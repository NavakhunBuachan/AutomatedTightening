echo on
echo =================================================
echo Installing latest UDBS Interface TRIAL version
echo =================================================
echo ... Copying VersionUpdater.exe to the location of VersionUpdater.exe.config
xcopy /q/y C:\Lumentum\VersionUpdater C:\MSIs\Lumentum\UDBS\UdbsInterface\Trial\v3
echo.
cd C:\MSIs\Lumentum\UDBS\UdbsInterface\Trial\v3
call VersionUpdater.exe -t
exit
