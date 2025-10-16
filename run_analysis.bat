@echo off
setlocal enabledelayedexpansion

REM === Configuration ===
set "INPUT_WORKBOOK=ALL SEM RESULTS.xlsx"
set "SUBJECT_CATALOG=CGPA.xlsx"
set "BIODATA_WORKBOOK=biodata.xlsx"
set "OUTPUT_DIR=outputs"
set "CURRENT_SEM=5"

REM === Verify required files exist ===
if not exist "%INPUT_WORKBOOK%" (
  echo [ERROR] Could not find %INPUT_WORKBOOK% in %cd%
  goto :fail
)
if not exist "%SUBJECT_CATALOG%" (
  echo [ERROR] Could not find %SUBJECT_CATALOG% in %cd%
  goto :fail
)
if not exist "%BIODATA_WORKBOOK%" (
  echo [ERROR] Could not find %BIODATA_WORKBOOK% in %cd%
  goto :fail
)

REM === Locate Python launcher ===
set "PYTHON_CMD=py -3"
where py >nul 2>&1
if errorlevel 1 (
  where python >nul 2>&1
  if errorlevel 1 (
    echo [ERROR] Python 3 is required but was not found on this system.
    goto :fail
  )
  set "PYTHON_CMD=python"
)

REM === Create virtual environment if missing ===
if not exist .venv (
  echo [INFO] Creating virtual environment...
  %PYTHON_CMD% -m venv .venv
  if errorlevel 1 goto :fail
)

call .venv\Scripts\activate
if errorlevel 1 goto :fail

REM === Install/update dependencies ===
python -m pip install --upgrade pip
if errorlevel 1 goto :fail
pip install -r requirements.txt
if errorlevel 1 goto :fail

REM === Run the analysis ===
python combine_results.py --input "%INPUT_WORKBOOK%" --biodata "%BIODATA_WORKBOOK%" --subject-catalog "%SUBJECT_CATALOG%" --outdir "%OUTPUT_DIR%" --current-semester %CURRENT_SEM%
if errorlevel 1 goto :fail

echo.
echo [SUCCESS] Analysis complete. Outputs are available in the %OUTPUT_DIR% folder.
goto :end

:fail
echo.
echo [FAILED] The analysis did not complete successfully. Review the messages above for details.

:end
call .venv\Scripts\deactivate.bat >nul 2>&1
pause
exit /b
