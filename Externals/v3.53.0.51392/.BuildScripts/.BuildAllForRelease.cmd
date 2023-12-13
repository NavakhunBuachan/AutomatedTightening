:: ============================================================================================================
:: .BuildAllForRelease.cmd batch script:
:: 1. Sets the current directory to the path containing the solution file (script path or parent), exiting with error if solution file is not found.
:: 2. Ensures that `SubWCRev.exe` can be found. 
:: 3. Ensures no local modifications or mixed revisions in the working copy 
:: 4. Ensures that `ReleaseNotes.md` file contains `$WCREV$` section
:: 5. If 1-4 are successful, updates the release notes to replace `$WCREV$` with the SVN revision (see [Edit Release Notes](ReleaseProcess.md#step-2-edit-release-notes) in [Release Process](ReleaseProcess.md)).
:: 6. Builds all combinations of configurations and platforms (calling `.BuildAll.cmd` script from same path)
:: 7. If 1-6 are successful, copies the binary output from `bin\` to `bin_version\`, for commit to Subversion
::
:: It should be included as an external to projects that include 'AssemblyVersion_Template.md'.
::
:: !!IMPORTANT!!
:: Before running this script:
:: - Make sure to update the "AssemblyVersion_Template.md" and "Documents\ReleaseNotes.md".
:: - Commit your code changes to SVN, to get the correct SVN revision in the above files.
::
:: After running this script:
:: - Commit the updated files in the "bin_version" folder to SVN for official release and tag the project.
:: See http://fractal.li.lumentuminc.net/fractal/articles/Fractal/Tutorials/ReleaseProcess.html
:: ============================================================================================================
:: $HeadURL: svn+ssh://ottnrp01/AutomationTestSW/BatchScripts/FractalBuild/.BuildAllForRelease.cmd $
:: $Revision: 37769 $, $Author: deo66257 $, $Date: 2020-03-06 03:33:38 +0700 (Fri, 06 Mar 2020) $
:: ============================================================================================================

@ECHO OFF

:: Find the path to the solution file, in either the script path, or the parent folder
ECHO ===================

:: Get the full path to the script
SET ScriptPath=%~dp0
SET SolutionFolder=%ScriptPath%

:: Find the solution path
IF NOT EXIST "%SolutionFolder%*.sln" (
	:: Try 1 level up
	SET SolutionFolder=%ScriptPath%..\
)

IF NOT EXIST "%SolutionFolder%*.sln" (
	ECHO ERROR! Solution file not found in "%ScriptPath%" or parent folder.
	EXIT /B 1
)

:: Change current directory to the solution folder
:: NOTE: This is preferred over using the SolutionFolder variable in the commands below, 
:: in case of a space in the path, so that commands do not need to be enclosed in quotations
CD "%SolutionFolder%"

:: By default, copy the output to the 'bin_version' folder after build
:: This will be set to zero if any errors occur 
SET COPY_OUTPUT=1

:: Make sure that SubWCRev is installed
ECHO Where SubWCRev
CALL Where SubWCRev
IF ERRORLEVEL 1 (
	SET COPY_OUTPUT=0
	ECHO ===================
	ECHO !!!Error!!! 
	ECHO Path to SubWCRev.exe not found.  Make sure TortoiseSVN is installed and path is included in System Environment Path variable.	
	GOTO :BUILD_ALL
)

:: 3. Check for local modification or mixed revisions in the working copy 
:: -n indicates that SubWCRev will exit with ERRORLEVEL 7 if the working copy contains local modifications
:: -m indicates that SubWCRev will exit with ERRORLEVEL 8 if the working copy contains mixed revisions to prevent building with a partially updated working copy. 
CALL SubWCRev.exe "%cd%." -nm

IF %ERRORLEVEL% NEQ 0 (
	SET COPY_OUTPUT=0
	ECHO ===================
	ECHO !!!Error!!! 
	if %ERRORLEVEL%==8 ( ECHO - Make sure that working copy has been updated from SVN before running this script )
	if %ERRORLEVEL%==7 ( ECHO - Make sure that all changes are committed to SVN before running this script )
	ECHO ===================
)

:: 4. Check that ReleaseNotes file contains $WCREV$ section
ECHO Checking Release Notes ...
Find /c "$WCREV$" "%cd%\Documents\ReleaseNotes.md"
IF ERRORLEVEL 1	(
	SET COPY_OUTPUT=0
	ECHO !!!Error!!! 'ReleaseNotes.md' does not contain '$WCREV$'
	ECHO ===================
)

:: 5. Only update the Release Notes with the SVN keywords if there are no local modifications and release notes are updated 
IF %COPY_OUTPUT% == 1 (	
	CALL SubWCRev.exe "%cd%." "Documents\ReleaseNotes.md" "Documents\ReleaseNotes.md"
	ECHO Successfully updated "Documents\ReleaseNotes.md" 
)

:BUILD_ALL

:: 6. Build all combinations of configurations and platforms, according to the solution in this folder
CALL "%ScriptPath%.BuildAll.cmd" || SET COPY_OUTPUT=0

IF %COPY_OUTPUT% == 1 (
	:: 7. Copy relevant files in the "bin" folder to the "bin_version" folder, suppressing prompt to overwrite
	ECHO Copying output to 'bin_version'
	CALL xcopy bin bin_version /i /s /y
	:: If the Debug folder exists, copy net40-x86 for legacy applications
	IF EXIST bin_version\Debug ( xcopy bin_version\net40-x86 bin_version\Debug /i /s /y	) 
	ECHO ===================
	ECHO Success! Build output copied to 'bin_version'
	EXIT
) ELSE (
	ECHO ===================
	ECHO !!!Error!!!
	ECHO See warnings or errors above.
	ECHO - Did not replace SVN keywords in 'ReleaseNotes.md'
	ECHO - Build output was not copied to 'bin_version' folder for release
	:: Pause to show results
	PAUSE
	:: Exit with code 1 to indicate error
	EXIT /B 1
)
