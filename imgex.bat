@echo off

:: Fordere den Benutzer auf, den Dateinamen einzugeben
set /P name=Datei: 
echo "%name%"

:: Überprüfe, ob die Datei eine Word oder PowerPoint Datei ist
set nameb="%~n1"
set ending="%~x1"
echo %ending%

if /I %ending%==.docx (
    echo Word Document detected
    set mediapath=word
    echo BP1
) else if /I %ending%==.pptx (
    set mediapath=ppt    
) else (
    echo Ungültige Datei
    pause
    exit
)

:: Konvertiere die word oder pptx Datei in eine zip Datei damit wir die Bilder extrahieren können
copy "%name%" "%nameb%.zip"

:: Kopiere die Bilder in den Ordner "unz" und lösche die gerade erstellte Datei
powershell -command "Expand-Archive '%nameb%.zip' -DestinationPath './unz'"
del "%nameb%.zip"

:: Erstelle den Ordner "images"
mkdir images

:: Kopiert alle Bilddateien in den Ordner "images"
xcopy /Q ".\unz\%mediapath%\media\*.*" "images" /s

:: Lösche den Ordner "unz"
rd /Q /S "unz"

:: Eine Nachricht anzeigen, dass die Bilder kopiert wurden
echo "Bilder kopiert"
