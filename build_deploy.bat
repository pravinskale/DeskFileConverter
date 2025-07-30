@echo off
setlocal

:: === CONFIGURATION ===
set SCRIPT_NAME=app.py
set EXE_NAME=PDFToExcel.exe
set REMOTE_PATH=\\USLTCSBMPSVC01\Share\Pravin\PDFToExl
set PYINSTALLER_PATH=.\venv\Scripts\pyinstaller.exe

:: === STEP 1: Build the EXE ===
echo Building EXE from %SCRIPT_NAME%...
"%PYINSTALLER_PATH%" --onefile --windowed --name "PDFToExcel" %SCRIPT_NAME%

:: === STEP 2: Copy to Remote Location ===
echo Copying %EXE_NAME% to %REMOTE_PATH%...
copy /Y dist\%EXE_NAME% %REMOTE_PATH%

echo Done.
pause
