echo off
echo =================================================
echo Installing latest UDBS Interface library version
echo =================================================
echo.
echo ... Copying VersionUpdater.exe to the location of VersionUpdater.exe.config
xcopy /q/y C:\Lumentum\VersionUpdater C:\MSIs\Lumentum\UDBS\UdbsInterface\Production\v3
cd C:\MSIs\Lumentum\UDBS\UdbsInterface\Production\v3
echo.
echo ... Starting VersionUpdater.exe...
call VersionUpdater.exe
exit