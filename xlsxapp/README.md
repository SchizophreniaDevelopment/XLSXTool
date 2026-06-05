# README

How to use:
## 1. Install dependencies:
```bash
sudo apt install apache2 libapache2-mod-wsgi-py3 python3-pip
pip install openpyxl # Or however you need to install it on your system
sudo a2enmod wsgi
```
```
/var/www/xlsxapp/
├── app.wsgi          ← WSGI entry point (Apache loads this)
├── processor.py      ← your existing CLI tool
└── uploads/          ← temp files (must be writable by www-data)
```
```bash
sudo mkdir -p /var/www/xlsxapp/uploads
sudo chown www-data:www-data /var/www/xlsxapp/uploads
```

## 2. Move xlsxapp.conf to /etc/apache2/site-available/xlsxapp.conf
```bash
sudo cp xlsxapp.conf /etc/apache2/site-available/xlsxapp.conf
cd /etc/apache2/site-available/xlsxapp.conf
sudo a2ensite xlsxapp
sudo a2dissite 000-default # Optional: disables the default site
sudo systemctl reload apache2
```
## 3. Test
```bash
# Quick smoke test from the command line
curl -X POST http://localhost/process \
  -F "file=@test.xlsx" \
  --output result.xlsx
```
Use the log if something goes wrong:
```bash
sudo tail -f /var/log/apache2/error.log
```
