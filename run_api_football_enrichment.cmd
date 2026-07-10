@echo off
setlocal
set PYTHON_EXE=C:\Users\rafael.weber\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe
set SCRIPT_PATH=C:\Users\rafael.weber\Documents\Codex\2026-04-24\estoy-trabajando-en-un-juego-web\enrich_futbrain_with_api_football.py
set INPUT_XLSX=C:\Users\rafael.weber\Documents\Codex\2026-04-24\estoy-trabajando-en-un-juego-web\futbrain_db_con_titulos.xlsx
set OUTPUT_XLSX=C:\Users\rafael.weber\Documents\Codex\2026-04-24\estoy-trabajando-en-un-juego-web\futbrain_db_con_titulos_api_football.xlsx
set API_KEY=24ad8ed6257c8768b381665d7ccc318f

"%PYTHON_EXE%" "%SCRIPT_PATH%" --api-key "%API_KEY%" --input "%INPUT_XLSX%" --output "%OUTPUT_XLSX%"

echo.
echo Listo. Si todo salio bien, revisa:
echo %OUTPUT_XLSX%
pause
