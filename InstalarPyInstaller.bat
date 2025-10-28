# Instalar PyInstaller
pip install pyinstaller

# Compilar com inclusão de todos os módulos
pyinstaller --onefile ^
            --windowed ^
            --name="PrintF" ^
            --icon="assets/icon.ico" ^
            --add-data="modules;modules" ^
            --add-data="config.py;." ^
            --hidden-import="PIL" ^
            --hidden-import="PIL._tkinter_finder" ^
            --hidden-import="docx" ^
            --hidden-import="pyautogui" ^
            --hidden-import="pynput" ^
            main.py