:: ============================================================================================================
:: This batch file builds the given solution with the given configuration and platform, 
:: appending to build log files in the bin folder
:: Arg1 = path to solution
:: Arg2 = build configuration
:: Arg3 = build platform
:: ============================================================================================================
:: $HeadURL: svn+ssh://ottnrp01/AutomationTestSW/BatchScripts/FractalBuild/.BuildProject.cmd $
:: $Revision: 41715 $, $Author: pan62372 $, $Date: 2020-08-11 08:35:34 +0700 (Tue, 11 Aug 2020) $
:: ============================================================================================================

@ECHO OFF

:: Get the arguments for the script:
set PATH_SOURCE_SLN=%1
set CONFIG=%2
set PLATFORM=%3

:: Set the path to the build log files.
set PATH_SUMMARY_LOG="%~dp1\bin\Build-Summary.log"
set PATH_DETAILS_LOG="%~dp1\bin\Build-Details.log"

ECHO ===================
ECHO Building %CONFIG%-%PLATFORM%

:: Clean and rebuild solution, with logging (fl = filelogger, flp = fileloggerparamters)
:: You can specify the following verbosity levels: q[uiet], m[inimal], n[ormal], d[etailed], and diag[nostic].
:: Verbosity (v): minimal will include summary, with warnings and errors only
CALL MSBuild %PATH_SOURCE_SLN% /t:clean;restore /p:Configuration=%CONFIG% /p:Platform=%PLATFORM% /v:quiet /fl1 /fl2 /flp1:logfile=%PATH_SUMMARY_LOG%;Verbosity=minimal;Append /flp2:logfile=%PATH_DETAILS_LOG%;Verbosity=normal;Append
CALL MSBuild %PATH_SOURCE_SLN% /t:rebuild /p:Configuration=%CONFIG% /p:Platform=%PLATFORM% /v:quiet /fl1 /fl2 /flp1:logfile=%PATH_SUMMARY_LOG%;Verbosity=minimal;Append /flp2:logfile=%PATH_DETAILS_LOG%;Verbosity=normal;Append