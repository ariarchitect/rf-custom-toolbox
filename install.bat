@echo off
for /f "tokens=2,*" %%a in ('reg query "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v "Personal" 2^>nul') do set DOCS=%%b
echo ----------------------
if not exist "%DOCS%\RFTools Files\RFCustomToolbox" (
    echo "ERROR: path %DOCS%\RFTools Files\RFCustomToolbox does not exist"
	goto EOF
)
if exist "%DOCS%\RFTools Files\RFCustomToolbox\0. Exyte.tab" (
	echo "removing old Exyte.tab folder..."
    rmdir /S /Q "%DOCS%\RFTools Files\RFCustomToolbox\0. Ariarchitect.tab"
)
echo "Copying new Exyte.tab folder..."
xcopy "0. Ariarchitect.tab" "%DOCS%\RFTools Files\RFCustomToolbox\0. Ariarchitect.tab" /E /I

Echo Ariarchitect panel for RF Tools Version 20250526 is installed
:EOF
pause