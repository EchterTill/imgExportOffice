@echo off
set /P name=Datei: 
set nameb=%name:~0,-5%
copy %name% %nameb%.zip
powershell -command "Expand-Archive %nameb%.zip ./unz"
del %nameb%.zip
mkdir images
xcopy /Q .\unz\word\media\*.* "images" /s
rd /Q /S "unz"
echo "Bilder kopiert"