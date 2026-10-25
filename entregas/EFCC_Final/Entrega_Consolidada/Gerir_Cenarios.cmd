@echo off
setlocal
choice /C BCRT /N /M "B=Guardar Base, C=Guardar Comparativo, R=Restaurar versao selecionada, T=Validar: "
if errorlevel 4 goto validate
if errorlevel 3 goto restore
if errorlevel 2 goto comparative
powershell -NoProfile -File "%~dp0Gerir_Cenarios_Excel365.ps1" -Action SaveBase
goto end
:comparative
powershell -NoProfile -File "%~dp0Gerir_Cenarios_Excel365.ps1" -Action SaveComparative
goto end
:restore
powershell -NoProfile -File "%~dp0Gerir_Cenarios_Excel365.ps1" -Action Restore
goto end
:validate
powershell -NoProfile -File "%~dp0Gerir_Cenarios_Excel365.ps1" -Action Validate
:end
pause
endlocal
