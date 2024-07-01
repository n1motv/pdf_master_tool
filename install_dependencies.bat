@echo off
echo Checking for pip...

:: Check if pip is installed
python -m pip --version
if %ERRORLEVEL% NEQ 0 (
    echo pip not found, installing pip...
    python -m ensurepip --upgrade
    python -m pip install --upgrade pip
)

echo Installing Python dependencies...
pip install ttkthemes PyPDF2 pdf2image pdf2docx pdfkit reportlab

echo Downloading and installing Poppler for Windows...
set POPPLER_URL=https://github.com/oschwartz10612/poppler-windows/releases/download/v23.01.0-0/poppler-23.01.0-0.zip
set POPPLER_DIR=C:\poppler

:: Check if Poppler directory exists
if not exist %POPPLER_DIR% (
    powershell -Command "& {
        Invoke-WebRequest -Uri '%POPPLER_URL%' -OutFile 'poppler.zip';
        Expand-Archive -Path 'poppler.zip' -DestinationPath '%POPPLER_DIR%';
        Remove-Item 'poppler.zip';
    }"
)

:: Add Poppler to PATH if not already added
for %%i in ("%POPPLER_DIR%\Library\bin") do (
    if not "%%~$PATH:i"=="" (
        echo Poppler is already in PATH
    ) else (
        setx PATH "%PATH%;%POPPLER_DIR%\Library\bin"
        echo Added Poppler to PATH
    )
)

echo Downloading and installing wkhtmltopdf for Windows...
set WKHTMLTOPDF_URL=https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6-1/wkhtmltox-0.12.6-1.msvc2015-win64.exe
set WKHTMLTOPDF_DIR=C:\Program Files\wkhtmltopdf

:: Download and install wkhtmltopdf if not already installed
if not exist "%WKHTMLTOPDF_DIR%\bin\wkhtmltopdf.exe" (
    powershell -Command "& {
        Invoke-WebRequest -Uri '%WKHTMLTOPDF_URL%' -OutFile 'wkhtmltox-installer.exe';
        Start-Process 'wkhtmltox-installer.exe' -ArgumentList '/SILENT' -Wait;
        Remove-Item 'wkhtmltox-installer.exe';
    }"
)

:: Add wkhtmltopdf to PATH if not already added
for %%i in ("%WKHTMLTOPDF_DIR%\bin") do (
    if not "%%~$PATH:i"=="" (
        echo wkhtmltopdf is already in PATH
    ) else (
        setx PATH "%PATH%;%WKHTMLTOPDF_DIR%\bin"
        echo Added wkhtmltopdf to PATH
    )
)

echo Dependencies installed successfully.
pause
