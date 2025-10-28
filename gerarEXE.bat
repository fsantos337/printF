@echo off
echo Compilando PrintF...
pyinstaller --onefile --windowed --name "PrintF" --add-data "modules;modules" --add-data "config.py;." --hidden-import=modules.capture --hidden-import=modules.template_gen --hidden-import=modules.evidence_gen --hidden-import=modules.cleanup --hidden-import=modules.styles --hidden-import=PIL._tkinter_finder main.py

if exist "dist\PrintF.exe" (
    echo ✅ Executável criado com sucesso: dist\PrintF.exe
) else (
    echo ❌ Falha na compilação
    pause
)