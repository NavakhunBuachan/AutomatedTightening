echo on
echo =================================================
echo Ending UDBS Interface Trial
echo =================================================
echo ... Copying VersionUpdater.exe to the location of VersionUpdater.exe.config
xcopy /q/y C:\Lumentum\VersionUpdater C:\MSIs\Lumentum\UDBS\UdbsInterface\Production\v3
echo.
cd C:\MSIs\Lumentum\UDBS\UdbsInterface\Production\v3
start VersionUpdater.exe -p
