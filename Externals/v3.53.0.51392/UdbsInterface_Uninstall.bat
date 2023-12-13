@echo off
echo ========================================================
echo Uninstalling SUDBS
echo ========================================================
echo.

:: check for elevated permission
call :check_admin || goto :eof

REG DELETE HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Lumentum\Deployment\UdbsInterface\v3 /va /f 2> nul

REG DELETE HKEY_LOCAL_MACHINE\SOFTWARE\Lumentum\Deployment\UdbsInterface\v3 /va /f 2> nul

start C:\Lumentum\BindingRedirectUpdater\BindingRedirectUpdater.exe -r -t development

msiexec.exe /x C:\MSIs\Lumentum\UDBS\UdbsInterface\Production\v3\UdbsInterface_installer_x64.msi /QN
msiexec.exe /x C:\MSIs\Lumentum\UDBS\UdbsInterface\Production\v3\UdbsInterface_installer_x86.msi /QN
msiexec.exe /x C:\MSIs\Lumentum\UDBS\UdbsInterface\Trial\v3\UdbsInterface_installer_x64.msi /QN
msiexec.exe /x C:\MSIs\Lumentum\UDBS\UdbsInterface\Trial\v3\UdbsInterface_installer_x86.msi /QN
msiexec.exe /x C:\MSIs\Lumentum\UDBS\UdbsInterface\Development\v3\UdbsInterface_installer_x64.msi /QN
msiexec.exe /x C:\MSIs\Lumentum\UDBS\UdbsInterface\Development\v3\UdbsInterface_installer_x86.msi /QN
goto :eof

:check_admin
:: function to check whether we are running with elevated permission (run as administrator)
:: returns 0 if administrator
fsutil dirty query %systemdrive% >nul
if %errorlevel% neq 0 (
  echo Error! Elevated permissions are required. Access denied. 
  echo Press any key to exit and re-run this script as administrator.
  pause >nul
)
exit /b