import subprocess
import tempfile
import os

UPLOAD_DIR = '/var/www/xlsxapp/uploads'

# ---------------------------------------------------------------------------
# Main WSGI entry point
# ---------------------------------------------------------------------------

def application(environ, start_response):
    method      = environ.get('REQUEST_METHOD', '')
    path        = environ.get('PATH_INFO', '')
    script_name = environ.get('SCRIPT_NAME', '')

    # Apache with WSGIScriptAlias can move the matched segment into
    # SCRIPT_NAME and leave PATH_INFO as '' or '/'. Fall back to
    # SCRIPT_NAME so routing works regardless of how Apache rewrites it.
    route = path if path not in ('', '/') else script_name

    if method != 'POST':
        start_response('405 Method Not Allowed', [('Content-Type', 'text/plain')])
        return [b'Method Not Allowed']

    if '/process' in route:
        return handle_process(environ, start_response)
    elif '/import' in route:
        return handle_import(environ, start_response)
    else:
        start_response('404 Not Found', [('Content-Type', 'text/plain')])
        return [(f'Not Found (PATH_INFO={path!r} SCRIPT_NAME={script_name!r})').encode()]


# ---------------------------------------------------------------------------
# /process  — run processor.py, return the output .xlsx file
# ---------------------------------------------------------------------------

def handle_process(environ, start_response):
    try:
        file_data = read_uploaded_file(environ)
    except Exception as e:
        start_response('400 Bad Request', [('Content-Type', 'text/plain')])
        return [('Bad request: ' + str(e)).encode()]

    with tempfile.NamedTemporaryFile(dir=UPLOAD_DIR, suffix='.xlsx', delete=False) as tmp_in:
        tmp_in.write(file_data)
        input_path = tmp_in.name

    output_path = input_path.replace('.xlsx', '_processed.xlsx')

    try:
        result = subprocess.run(
            ['/var/www/xlsxapp/venv/bin/python3', '/var/www/xlsxapp/processor.py', input_path, output_path],
            capture_output=True,
            text=True,
            timeout=60
        )

        if result.returncode != 0:
            raise RuntimeError(result.stderr or 'processor.py exited with code ' + str(result.returncode))

        with open(output_path, 'rb') as f:
            output_bytes = f.read()

        headers = [
            ('Content-Type',        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
            ('Content-Disposition', 'attachment; filename="processed.xlsx"'),
            ('Content-Length',      str(len(output_bytes))),
        ]
        start_response('200 OK', headers)
        return [output_bytes]

    except Exception as e:
        start_response('500 Internal Server Error', [('Content-Type', 'text/plain')])
        return [('Processing error: ' + str(e)).encode()]

    finally:
        _remove(input_path)
        _remove(output_path)


# ---------------------------------------------------------------------------
# /import  — run importer.py, return plain-text output (no file download)
# ---------------------------------------------------------------------------

def handle_import(environ, start_response):
    try:
        file_data = read_uploaded_file(environ)
    except Exception as e:
        start_response('400 Bad Request', [('Content-Type', 'text/plain')])
        return [('Bad request: ' + str(e)).encode()]

    with tempfile.NamedTemporaryFile(dir=UPLOAD_DIR, suffix='.xlsx', delete=False) as tmp_in:
        tmp_in.write(file_data)
        input_path = tmp_in.name

    try:
        result = subprocess.run(
            ['/var/www/xlsxapp/venv/bin/python3', '/var/www/xlsxapp/importer.py', input_path],
            capture_output=True,
            text=True,
            timeout=120
        )

        if result.returncode != 0:
            raise RuntimeError(result.stderr or 'importer.py exited with code ' + str(result.returncode))

        output = (result.stdout or 'Import successful.').encode()
        start_response('200 OK', [
            ('Content-Type',   'text/plain'),
            ('Content-Length', str(len(output))),
        ])
        return [output]

    except Exception as e:
        start_response('500 Internal Server Error', [('Content-Type', 'text/plain')])
        return [('Import error: ' + str(e)).encode()]

    finally:
        _remove(input_path)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def read_uploaded_file(environ):
    """Parse the first file out of a multipart/form-data request body."""
    content_type = environ.get('CONTENT_TYPE', '')
    if 'multipart/form-data' not in content_type:
        raise ValueError('Expected multipart/form-data, got: ' + content_type)

    try:
        boundary = content_type.split('boundary=')[1].strip().encode()
    except IndexError:
        raise ValueError('Missing boundary in Content-Type header')

    try:
        content_length = int(environ.get('CONTENT_LENGTH', 0))
    except ValueError:
        raise ValueError('Invalid or missing Content-Length')

    if content_length <= 0:
        raise ValueError('Empty request body')

    body = environ['wsgi.input'].read(content_length)
    file_data = parse_multipart(body, boundary)

    if file_data is None:
        raise ValueError('No file field found in multipart body')

    return file_data


def parse_multipart(body, boundary):
    """Extract the bytes of the first part that has a filename parameter."""
    delimiter = b'--' + boundary
    parts = body.split(delimiter)

    for part in parts[1:]:
        if b'\r\n\r\n' not in part:
            continue
        headers_raw, content = part.split(b'\r\n\r\n', 1)
        if b'filename' in headers_raw:
            return content.rstrip(b'\r\n--')

    return None


def _remove(path):
    """Delete a file, silently ignoring errors."""
    try:
        os.unlink(path)
    except OSError:
        pass
