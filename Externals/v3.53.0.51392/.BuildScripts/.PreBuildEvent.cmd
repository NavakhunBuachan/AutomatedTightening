:: ============================================================================================================
:: This batch file runs SubWCRev to replace the keywords defined in the AssemblyVersion_Template.md file 
:: with the SVN working copy values, to link the Assembly version with the SVN revision.
::
:: See http://fractal.li.lumentuminc.net/tutorials/articles/Release_LinkingAssemblyVersionToSvn.html
::
:: Usage: 
:: - Add as external to the project, in the ".BuildScripts" subfolder
:: - Set the Pre-build event in the project file to the following:
::   cmd /c $(ProjectDir).BuildScripts\.PreBuildEvent.cmd $(ProjectDir)
::
:: Troubleshooting:
:: "'SubWCRev.exe' is not recognized as an internal or external command, operable program or batch file."
:: -> Add "C:\Program Files\TortoiseSVN\bin" to "Path" Environment Variable
:: -> If it still fails, reboot your computer
:: 
:: TODO - Add handling for case where SVN is not available, by 
:: - copying the AssemblyVersion_Template.vb file to AssemblyVersion.vb, and 
:: - replacing the "$WCREV$$WCMODS?*:$" with "SvnRev?" (for the revision field of the version, to indicate that SVN revision is unknown).
:: ============================================================================================================
:: $HeadURL: svn+ssh://ottnrp01/AutomationTestSW/BatchScripts/FractalBuild/.PreBuildEvent.cmd $
:: $Revision: 37769 $, $Author: deo66257 $, $Date: 2020-03-06 03:33:38 +0700 (Fri, 06 Mar 2020) $
:: ============================================================================================================

@ECHO OFF

SET ErrorCode=0

:: Project path is passed as argument
SET ProjectPath=%1%
ECHO ProjectPath %ProjectPath%

CD "%ProjectPath%"

:: Determine whether the project is VB or C#
SET Extension=vb
SET VersionFolder=My Project
IF NOT EXIST "%VersionFolder%" (
	SET Extension=cs
	SET VersionFolder=Properties
)

IF NOT EXIST "%VersionFolder%" (
	SET ErrorCode=1
	ECHO Could not determine if project is VB or C#. 
	GOTO SHOW_ERROR
)

:: Make sure the template exists
IF NOT EXIST "%VersionFolder%\AssemblyVersion_Template.md" (
	SET ErrorCode=2
	ECHO Could not find "AssemblyVersion_Template.md" in "%ProjectPath%" 
	GOTO SHOW_ERROR
)

:: Make sure that SubWCRev is installed
ECHO Where SubWCRev
CALL Where SubWCRev
IF ERRORLEVEL 1 (
	ECHO Path to SubWCRev.exe not found.  Make sure TortoiseSVN is installed and path is included in System Environment Path variable.	
	GOTO SHOW_ERROR
)

:: Update the AssemblyVersion with the SVN keywords
SubWCRev.exe "." "%VersionFolder%\AssemblyVersion_Template.md" "%VersionFolder%\AssemblyVersion.%Extension%"
IF %ERRORLEVEL% NEQ 0 (
	SET ErrorCode=3
	GOTO SHOW_ERROR
)

ECHO Successfully replaced AssemblyVersion with SVN revision
EXIT /B 0

:SHOW_ERROR

ECHO.
SET ERRORLEVEL=%ErrorCode%
ECHO Error!! Aborted early with Code %ERRORLEVEL%.
	
:: Exit with code to indicate error
EXIT /B %ErrorCode%
	
:FINISHED