# Script de build dos executáveis usando PyInstaller
# Este script gera as versões compiladas autônomas (.exe) para o CLI e a GUI.

Write-Host "[*] Iniciando compilação dos executáveis com PyInstaller..."

# 1. Compilação do CLI (Linha de comando)
Write-Host "[*] Compilando executável do CLI (docx-footnote-injector-by-jair-lima)..."
pyinstaller --onefile --add-data "dicionario_padrao.json;." --name "docx-footnote-injector-by-jair-lima" docx_footnote_injector.py

# 2. Compilação da GUI (Interface Gráfica)
Write-Host "[*] Compilando executável da GUI (docx-footnote-gui-by-jair-lima)..."
# É necessário utilizar --collect-all customtkinter para incluir recursos de tema e layout do CustomTkinter no binário.
pyinstaller --onefile --windowed --add-data "dicionario_padrao.json;." --collect-all customtkinter --name "docx-footnote-gui-by-jair-lima" docx_footnote_gui.py

Write-Host "[+] Compilação concluída!"
Write-Host "[+] Os executáveis resultantes (.exe) estão disponíveis no diretório 'dist/'."
