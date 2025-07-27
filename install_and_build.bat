@echo off
echo ====================================
echo Installing Python and building app
echo ====================================

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found! Downloading and installing Python...
    
    REM Download Python installer
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe' -OutFile 'python_installer.exe'"
    
    REM Install Python silently
    echo Installing Python 3.11.9...
    python_installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
    
    REM Wait for installation to complete
    timeout /t 30 /nobreak >nul
    
    REM Clean up installer
    del python_installer.exe
    
    echo Python installation completed!
    echo Please restart this script after Python installation.
    pause
    exit /b 1
) else (
    echo Python is already installed!
    python --version
)

echo.
echo ====================================
echo Installing required libraries...
echo ====================================

REM Upgrade pip first
python -m pip install --upgrade pip

REM Install from requirements.txt
if exist requirements.txt (
    echo Installing packages from requirements.txt...
    python -m pip install -r requirements.txt
    
    echo Installing additional dependencies...
    python -m pip install python-Levenshtein
    python -m pip install openpyxl
    python -m pip install "numpy<2"
) else (
    echo requirements.txt not found!
    echo Please make sure requirements.txt is in the same folder as this script.
    pause
    exit /b 1
)

echo.
echo ====================================
echo All libraries installed successfully!
echo ====================================

echo.
echo ====================================
echo Building executable with PyInstaller...
echo ====================================

REM Check if spec file exists
if exist calc_zp.spec (
    echo Using existing spec file...
    pyinstaller calc_zp.spec
) else (
    echo Creating executable...
    pyinstaller --onefile --windowed --name="Расчет_ЗП" main.py
)

echo.
if exist dist (
    echo ====================================
    echo Build completed successfully!
    echo ====================================
    echo Executable file location: dist\
    echo You can find the executable in the 'dist' folder
    start dist
) else (
    echo ====================================
    echo Build failed!
    echo ====================================
    echo Please check the error messages above
)

echo.
echo Press any key to exit...
pause >nul