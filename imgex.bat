@echo off
set /P name=Datei: 
echo %name%
set nameb=%name:~0,-5%
set ending=%name:~-4%
echo %ending%
if %ending%==docx (
    echo Word Document detected
    %mediapath%=word
    echo BP1
) else if %ending%==pptx (
    %mediapath%="ppt"    
) else (
    echo Ungültige Datei
    pause
    exit
)
copy %name% %nameb%.zip
powershell -command "Expand-Archive %nameb%.zip ./unz"
del %nameb%.zip
mkdir images
xcopy /Q .\unz\%mediapath%\media\*.* "images" /s
rd /Q /S "unz"
echo "Bilder kopiert"
