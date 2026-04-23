.\venv\Scripts\Activate.ps1
pyinstaller --noconfirm --onefile --windowed --name "MarkItDown-GUI" --icon "assets/logo.ico" --add-data "assets;assets" --add-data "venv\Lib\site-packages\magika;magika" main.py
