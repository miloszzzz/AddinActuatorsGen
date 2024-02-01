echo off
setlocal enabledelayedexpansion

set "project_folder=%1"

:: Katalog docelowy dla skopiowanych plików DLL
set "destination_folder=%2"

:: Ścieżka do programu .exe
set "program_exe=%3"

:: Ścieżka do folderu, w którym znajdują się pliki DLL
set "dll_folder=%project_folder%\bin\Debug"

:: Przechodzenie do katalogu źródłowego z plikami DLL
cd "%dll_folder%"

:: Pętla przechodząca przez pliki DLL i kopiująca je do katalogu docelowego
for %%f in (*.dll) do (
    set "filename=%%~nf"
    echo !filename! | find /i "System" >nul && (
        echo Ignorowanie pliku: %%f
    ) || (
        copy /Y "%%f" "%destination_folder%\"
    )
)

:: Uruchomienie programu .exe
cd "C:\Program Files\Siemens\Automation\Portal V16\PublicAPI\V16.AddIn\"
start "Publisher" /B "C:\Program Files\Siemens\Automation\Portal V16\PublicAPI\V16.AddIn\Siemens.Engineering.AddIn.Publisher.exe" -f "%project_folder%\Config.xml" -v -c

:: Zakończenie skryptu
exit /b