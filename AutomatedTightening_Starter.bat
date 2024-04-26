echo off
echo =============================================================
echo Starting AutomatedTightening
echo =============================================================
echo.
echo Unpacking AutomatedTightening files ...
cd C:\TestSW.NET\AutomatedTightening
AutomatedTightening_installer.exe
echo ... unpacked files
echo.
echo Check and install UdbsInterface Shared Library...
::START /WAIT /b cmd /c C:\MSIs\Lumentum\UDBS\UdbsInterface\Production\v3\UdbsInterface_Install.bat
START /WAIT /b cmd /c C:\MSIs\Lumentum\UDBS\UdbsInterface\Trial\V3\UdbsInterface_StartTrial.bat
echo ...UdbsInterface Shared Library Version check completed.
echo.
echo Launching AutomatedTightening
echo (this window will close automatically when AutomatedTightening is closed) ...
cd C:\Program Files\Lumentum\AutomatedTightening
AutomatedTightening.exe
echo ... AutomatedTightening closed
exit