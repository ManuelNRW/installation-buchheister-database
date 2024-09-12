@echo off
:: Erstellt den Ordner "buchheister_database" im Home-Verzeichnis des aktuellen Benutzers

set "targetDir=%USERPROFILE%\buchheister_database"

:: Prüfen, ob der Ordner bereits existiert
if not exist "%targetDir%" (
    mkdir "%targetDir%"
    echo Der Ordner "buchheister_database" wurde erstellt.
) else (
    echo Der Ordner "buchheister_database" existiert bereits.
)

:: Dateien herunterladen
echo Starte Download der Dateien...

bitsadmin /transfer "DownloadFE" "https://www.internal.buchheister.de/buchheister_database/download/buchheister_database_FE.accde" "%targetDir%\buchheister_database_FE.accde"
bitsadmin /transfer "DownloadTXT1" "https://www.internal.buchheister.de/buchheister_database/download/1685163548135843521485.txt" "%targetDir%\1685163548135843521485.txt"
bitsadmin /transfer "DownloadVBS" "https://www.internal.buchheister.de/buchheister_database/download/update_buchheister_database.vbs" "%targetDir%\update_buchheister_database.vbs"
bitsadmin /transfer "DownloadTXT2" "https://www.internal.buchheister.de/buchheister_database/download/Updateanleitung.txt" "%targetDir%\Updateanleitung.txt"

echo Download der Dateien abgeschlossen.

:: Erstelle den Ordner "C:\ProgramData\", falls er noch nicht existiert
set "programDataDir=C:\ProgramData"

if not exist "%programDataDir%" (
    mkdir "%programDataDir%"
    echo Der Ordner "C:\ProgramData\" wurde erstellt.
) else (
    echo Der Ordner "C:\ProgramData\" existiert bereits.
)

:: Erstelle den Ordner "C:\ProgramData\buchheister_database", falls er noch nicht existiert
set "buchheisterDir=%programDataDir%\buchheister_database"

if not exist "%buchheisterDir%" (
    mkdir "%buchheisterDir%"
    echo Der Ordner "C:\ProgramData\buchheister_database" wurde erstellt.
) else (
    echo Der Ordner "C:\ProgramData\buchheister_database" existiert bereits.
)

:: Datei "b.ico" in den Ordner "C:\ProgramData\buchheister_database" herunterladen
echo Starte Download der Datei "b.ico"...

bitsadmin /transfer "DownloadICO" "https://www.internal.buchheister.de/buchheister_database/download/b.ico" "%buchheisterDir%\b.ico"

echo Download der Datei "b.ico" abgeschlossen.

:: Erstelle eine Desktop-Verknüpfung zur Datei "buchheister_database_FE.accde"
echo Erstelle Desktop-Verknüpfung...

set "shortcutPath=%USERPROFILE%\Desktop\buchheister_database.lnk"
set "targetFile=%targetDir%\buchheister_database_FE.accde"
set "iconFile=%buchheisterDir%\b.ico"

:: VBS-Script erstellen, um die Verknüpfung zu erzeugen
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%temp%\CreateShortcut.vbs"
echo sLinkFile = "%shortcutPath%" >> "%temp%\CreateShortcut.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%temp%\CreateShortcut.vbs"
echo oLink.TargetPath = "%targetFile%" >> "%temp%\CreateShortcut.vbs"
echo oLink.IconLocation = "%iconFile%" >> "%temp%\CreateShortcut.vbs"
echo oLink.Save >> "%temp%\CreateShortcut.vbs"

:: VBS-Script ausführen
cscript //nologo "%temp%\CreateShortcut.vbs"

:: VBS-Script löschen
del "%temp%\CreateShortcut.vbs"

echo Desktop-Verknüpfung wurde erstellt.

:: Registrierungseinstellung vornehmen
echo Füge Registrierungseinstellung hinzu...

reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Security" /v DisableHyperlinkWarning /t REG_DWORD /d 1 /f

echo Registrierungseinstellung wurde erfolgreich hinzugefügt.

:: Füge den Ordner "buchheister_database" im Homeverzeichnis als vertrauenswürdigen Speicherort in Access hinzu
echo Füge vertrauenswürdigen Speicherort für "buchheister_database" hinzu...

reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Access\Security\Trusted Locations\Location1" /v Path /t REG_SZ /d "%targetDir%" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Access\Security\Trusted Locations\Location1" /v AllowSubfolders /t REG_DWORD /d 1 /f

echo Vertrauenswürdiger Speicherort wurde hinzugefügt.