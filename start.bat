@echo off
setlocal

echo Checking if Python is installed...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found. Downloading and installing Python 3.11...
    powershell -Command "Invoke-WebRequest -Uri https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe -OutFile python-installer.exe"

    echo Installing Python in silent mode...
    python-installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
    if %errorlevel% neq 0 (
        echo Error during Python installation.
        pause
        exit /b 1
    )
    echo Deleting installer file...
    del python-installer.exe
    set "PATH=%SystemRoot%\system32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SYSTEMROOT%\System32\WindowsPowerShell\v1.0\"
    for /f "usebackq tokens=2,* skip=2" %%A in (`reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH`) do set "PATH=%PATH%;%%B"
)


if not exist "venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv venv
) else (
    echo Virtual environment already exists.
)

echo Activating virtual environment...
call venv\Scripts\activate

echo Installing/updating dependencies from requirements.txt...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo Running main script...
python main.py

echo.
echo Script has finished. Press any key to exit.
pause
