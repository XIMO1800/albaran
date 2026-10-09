@echo off
rem Crea la tarea que sube los resumenes de la PDA (AUTOVENT) cada 2 minutos. Doble clic una vez.
schtasks /create /tn "Puente AUTOVENT" /tr "wscript.exe C:\PUENTE_ALBARAN\puente-autovent.vbs" /sc minute /mo 2 /it /f
if errorlevel 1 (echo. & echo NO SE HA PODIDO CREAR LA TAREA. Haz una foto de esta ventana.) else (echo. & echo LISTO: los resumenes de la PDA se subiran solos cada 2 minutos.)
echo.
pause
