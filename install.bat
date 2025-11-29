@echo off
echo Setting up Certificate Generator...

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed. Please install Python 3.9+ from python.org
    pause
    exit /b 1
)

:: Create virtual environment
echo Creating virtual environment...
python -m venv cert_env

:: Activate virtual environment
echo Activating virtual environment...
call cert_env\Scripts\activate.bat

:: Install dependencies
echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt

:: Create necessary directories
mkdir templates 2>nul
mkdir certificates 2>nul

echo Installation complete!
echo.
echo To run the application:
echo cert_env\Scripts\activate.bat
echo python main.py
pause