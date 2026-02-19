@echo off
setlocal

set "APP_NAME=PDF Nova Desktop"
set "APP_DIR=%LOCALAPPDATA%\Programs\PDFNovaDesktop"
set "PAYLOAD=%~dp0PDFNovaDesktop_payload.zip"
set "EXE_PATH=%APP_DIR%\PDFNovaDesktop.exe"
set "UNINSTALL_CMD=%APP_DIR%\uninstall.cmd"

echo Installation de %APP_NAME%...

if not exist "%PAYLOAD%" (
  echo Payload introuvable: %PAYLOAD%
  exit /b 1
)

if exist "%APP_DIR%" (
  rmdir /s /q "%APP_DIR%" >nul 2>&1
)
mkdir "%APP_DIR%" >nul 2>&1

powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -LiteralPath '%PAYLOAD%' -DestinationPath '%APP_DIR%' -Force"
if errorlevel 1 (
  echo Echec extraction payload.
  exit /b 1
)

if not exist "%EXE_PATH%" (
  echo Executable introuvable apres extraction: %EXE_PATH%
  exit /b 1
)

> "%UNINSTALL_CMD%" echo @echo off
>> "%UNINSTALL_CMD%" echo setlocal
>> "%UNINSTALL_CMD%" echo set "APP_DIR=%LOCALAPPDATA%\Programs\PDFNovaDesktop"
>> "%UNINSTALL_CMD%" echo powershell -NoProfile -ExecutionPolicy Bypass -Command "$desk=[Environment]::GetFolderPath('Desktop');$start=[Environment]::GetFolderPath('Programs');$lnk1=Join-Path $desk 'PDF Nova Desktop.lnk';$lnk2=Join-Path $start 'PDF Nova Desktop.lnk';if(Test-Path $lnk1){Remove-Item $lnk1 -Force};if(Test-Path $lnk2){Remove-Item $lnk2 -Force}"
>> "%UNINSTALL_CMD%" echo if exist "%%APP_DIR%%" rmdir /s /q "%%APP_DIR%%"
>> "%UNINSTALL_CMD%" echo echo PDF Nova Desktop desinstalle.
>> "%UNINSTALL_CMD%" echo endlocal

powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell;$desk=[Environment]::GetFolderPath('Desktop');$start=[Environment]::GetFolderPath('Programs');$lnk1=$w.CreateShortcut((Join-Path $desk 'PDF Nova Desktop.lnk'));$lnk1.TargetPath='%EXE_PATH%';$lnk1.WorkingDirectory='%APP_DIR%';$lnk1.Save();$lnk2=$w.CreateShortcut((Join-Path $start 'PDF Nova Desktop.lnk'));$lnk2.TargetPath='%EXE_PATH%';$lnk2.WorkingDirectory='%APP_DIR%';$lnk2.Save()"
if errorlevel 1 (
  echo Raccourcis non crees. Installation continue.
)

echo Installation terminee.
start "" "%EXE_PATH%"
exit /b 0

