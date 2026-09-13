@echo off
rem ============================================================
rem  Actualiza el Radar de Licitaciones de Capacitacion
rem  (Mercado Publico) y abre el informe HTML al terminar.
rem ============================================================
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "buscar_licitaciones.ps1"
if exist "salida\informe_licitaciones_capacitacion.html" start "" "salida\informe_licitaciones_capacitacion.html"
pause
