@echo off
echo Getting folder sizes for you...storing to folderSizes.csv
setlocal disabledelayedexpansion

:: Delete existing CSV file if it exists
if EXIST folderSizes.csv del folderSizes.csv

:: Create the CSV file with headers
echo Folder,Bytes Size,Short Size > folderSizes.csv

:: Set the folder path. Replace <FOLDER_PATH> with the actual path
set "folder=%<FOLDER_PATH>"

:: Use the current directory if no folder is specified
if not defined folder set "folder=%cd%"

:: Loop through each subdirectory
for /d %%a in ("%folder%\*") do (
    set "size=0"
    :: Get the folder size
    for /f "tokens=3,5" %%b in ('dir /-c /a /w /s "%%~fa\*" 2^>nul ^| findstr /b /c:"  "') do if "%%~c"=="" set "size=%%~b"
    setlocal enabledelayedexpansion
    call :GetUnit !size! unit
    call :ConvertBytes !size! !unit! newsize
    echo(%%~nxa,!size!,!newsize!!unit! >> folderSizes.csv
    endlocal 
)

endlocal
exit /b

:ConvertBytes bytes unit ret
setlocal
:: Determine the conversion factor based on the unit
if "%~2" EQU "KB" set val=/1024
if "%~2" EQU "MB" set val=/1024/1024
if "%~2" EQU "GB" set val=/1024/1024/1024
if "%~2" EQU "TB" set val=/1024/1024/1024/1024

:: Use VBScript to format the number
> %temp%\tmp.vbs echo wsh.echo FormatNumber(eval(%~1%val%),1)
for /f "delims=" %%a in ('cscript //nologo %temp%\tmp.vbs') do (
    endlocal
    set %~3=%%a
)
del %temp%\tmp.vbs
exit /b

:GetUnit bytes return
setlocal
set byt=00000000000%1X
set TB=000000000001099511627776X

:: Determine the unit based on the size
if %1 LEQ 1024 set "unit=Bytes"
if %1 GTR 1024 set "unit=KB"
if %1 GTR 1048576 set "unit=MB"
if %1 GTR 1073741824 set "unit=GB"
if %byt:~-14% GTR %TB:~-14% set "unit=TB"

endlocal & set %~2=%unit%
exit /b
