@echo off
REM Build GUIBatchTranslator in onedir mode (faster startup than onefile)

REM Clean old build/dist
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Run PyInstaller
pyinstaller.exe ^
  GUIBatchTranslator.py ^
  --name "GUIBatchTranslator" ^
  --onedir ^
  --noconsole ^
  --clean ^
  --noupx ^
  --runtime-tmpdir "%LOCALAPPDATA%\GUIBatchTranslator\_tmp" ^
  --exclude-module PyQt5.QtWebEngineWidgets ^
  --exclude-module PyQt5.QtWebEngineCore ^
  --exclude-module PyQt5.QtWebEngine ^
  --collect-submodules argostranslate ^
  --collect-submodules argos_translate_files ^
  --collect-data sentencepiece ^
  --collect-binaries ctranslate2 ^
  --add-data "Models;Models"

echo.
echo Build finished! Check the "dist\GUIBatchTranslator" folder.
pause
