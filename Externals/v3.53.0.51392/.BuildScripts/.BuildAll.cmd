:: ============================================================================================================
:: Builds all combinations of configurations and platforms, according to the solution file.
:: - Supports the following configurations: net40|x86, net40|x64, net462|x86, net462|x64
::   - To add more, update the NetFrameworks and TargetCpus variables at the top of the script.
:: - Looks for a single solution file in either the path of the script, or the parent folder.  
:: - Exits with error if either none found or more than one found.
:: - Cleans and rebuilds solution for all configurations and platforms that are contained in the solution file
:: ============================================================================================================
:: $HeadURL: svn+ssh://ottnrp01/AutomationTestSW/BatchScripts/FractalBuild/.BuildAll.cmd $
:: $Revision: 30936 $, $Author: sfrench $, $Date: 2019-05-05 07:35:58 +0700 (Sun, 05 May 2019) $
:: ============================================================================================================
@ECHO OFF

:: Global variables defining .NET Frameworks and TargetCpus over which to iterate
SET NetFrameworks=net40 net462
SET TargetCpus=x86 x64

:: Global variables that will be updated by functions
SET FileName=""
SET SolutionFile=""
SET ProjectFile=""
SET FoundConfig=0

:: Get the full path to the script
SET ScriptPath=%~dp0

:: Find the path to the solution file, in either the script path, or the parent folder
ECHO ===================
CALL :FIND_SOLUTION
IF %ERRORLEVEL% NEQ 0 (
	:: Pause to show results
	PAUSE
	EXIT /B %ERRORLEVEL%
)
ECHO Solution File is "%SolutionFile%"

:: Find the path to the project file, in the solution path
ECHO ===================
CALL :FIND_PROJECT
IF %ERRORLEVEL% NEQ 0 (
	:: Pause to show results
	PAUSE
	EXIT /B %ERRORLEVEL%
)
ECHO Project File is %ProjectFile%

:: Delete the bin and obj folders before building (if exist)
ECHO ===================
CALL :DELETE_BIN

:: Clean and rebuild solution for all configurations and platforms that are contained in the solution file
SET /A ErrCount = 0
SET /A BuildCount = 0

:: Required to use variable values inside the loops below
SetLocal EnableDelayedExpansion

:: Iterate over NET Frameworks and Target Cpus
FOR %%i IN (%NetFrameworks%) DO (
	FOR %%j IN (%TargetCpus%) DO (
		ECHO ===================
		CALL :FIND_CONFIG %%i %%j
		IF !FoundConfig! == 0 (
			CALL "%ScriptPath%.BuildProject" %ProjectFile% "%%i" "%%j" || SET /A ErrCount = %ErrCount% + 1
			SET /A BuildCount = %BuildCount% + 1
		)
	)
)

ECHO ===================
IF %ErrCount% NEQ 0 (
	ECHO Finished %BuildCount% builds with %ErrCount% errors
	PAUSE
) ELSE (
	ECHO Finished %BuildCount% builds successfully
)

EXIT /B %ErrCount%

:: Function to find the path to the solution file in the script folder or its parent
:FIND_SOLUTION

:: Find the solution path in current directory (script path)
CD %ScriptPath%
CALL :FIND_FILE "*.sln"
IF %ERRORLEVEL%==0 (
	SET SolutionFile=%FileName%
	EXIT /B 0
)

:: Try 1 level up
ECHO Looking in parent folder (CD ..\)
CD ..\
CALL :FIND_FILE "*.sln"
IF %ERRORLEVEL%==0 (
	SET SolutionFile=%FileName%
	EXIT /B 0
)

ECHO ERROR! Solution file not found in "%ScriptPath%" or parent folder.
EXIT /B 1

:FIND_PROJECT

:: Look for VB project
CALL :FIND_FILE "*.vbproj"
IF %ERRORLEVEL%==0 (
	SET ProjectFile=%FileName%
	EXIT /B 0
)

:: Look for C# project
CALL :FIND_FILE "*.csproj"
IF %ERRORLEVEL%==0 (
	SET ProjectFile=%FileName%
	EXIT /B 0
)

ECHO ERROR! Single Project file not found in path of Solution file.
EXIT /B 1

:: Function to find a single file in the current directory with the given search string
:FIND_FILE
ECHO Looking for single file matching %1% ...
ECHO ... CD is "%cd%"
IF NOT EXIST %1% (
	ECHO ... File not found matching %1%.
	EXIT /B 1
)

SET /A Count=0
FOR %%i IN (%1%) DO (
	SET /A Count+=1
	SET FileName=%%i
)

IF %Count% GTR 1 (
	ECHO ... ERROR! Found %Count% files matching %1%! Only 1 is allowed.
	EXIT /B %Count%
)

EXIT /B 0

:: Function to delete existing bin and obj folders
:DELETE_BIN
ECHO Deleting existing bin and obj folders ...
rmdir bin /s /q
rmdir obj /s /q
:: Suppress errors
EXIT /B 0

:: Function to find a given config string in the global solution file
:FIND_CONFIG
SET NetFramework=%1%
SET TargetCpu=%2%
ECHO Find "%NetFramework%|%TargetCpu%" in Solution ...
FIND /c "%NetFramework%|%TargetCpu% = %NetFramework%|%TargetCpu%" "%SolutionFile%"	
IF %ERRORLEVEL%==0 (
	ECHO ... "%NetFramework%|%TargetCpu%" Found
	SET FoundConfig=0
) ELSE (
	ECHO ... "%NetFramework%|%TargetCpu%" Not Found
	SET FoundConfig=1
)
EXIT /B FoundConfig