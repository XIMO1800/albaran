@echo off
rem Crea la tarea programada que lanza el puente cada 5 minutos. Doble clic una vez.
schtasks /create /tn "Puente ALBTIEND" /tr "wscript.exe C:\Puente\puente-albtiend.vbs" /sc minute /mo 5 /it /f
if errorlevel 1 (echo. & echo NO SE HA PODIDO CREAR LA TAREA. Haz una foto de esta ventana.) else (echo. & echo LISTO: el puente se lanzara solo cada 5 minutos.)
echo.
pause
