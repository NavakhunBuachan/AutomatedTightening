===================================================================================================
7-Zip Self-Extracting Installer Creator 
===================================================================================================
The contents of this folder are used by Visual Studio during a Post Build event
to create a self-extracting executable (called the "installer" throughout this 
documet) that contains the output of the application build.

The purpose of this file is to ease deployment through only requiring a 
single executable file.


---------------------------------------------------------------------------------------------------
FILES
---------------------------------------------------------------------------------------------------
Files in this folder:

	7za.exe
		The 7-Zip program that compresses the build folder into a single file
	
	7zsd.sfx
		SFX extensions to 7-Zip that enables the creation of the self-extracting executable
	
	BuildSelfExtractingEXE.vbs
		VB Script called from Visual Studio that in turn configures and calls 7-Zip to create 
		the self-extracting executable
	
	README.txt
		This file
	
	
---------------------------------------------------------------------------------------------------
USE
---------------------------------------------------------------------------------------------------
To create a installer as part of the test software build in
Visual Studio 2010:

Via SVN ...
1. Add the 7ZipSelfExtractingInstallerCreator folder as an external to your
   solution and ensure that is a subfolder of the solution. 
   
   The 7ZipSelfExtractingInstallerCreator folder is located in SVN at: 
		svn+ssh://ottnrp01/AutomationTestSW/VB.NET/Utilities/7ZipSelfExtractingInstallerCreator

	Paste the following in the list of externals for your application:
       ^/VB.NET/Utilities/7ZipSelfExtractingInstallerCreator 7ZipSelfExtractingInstallerCreator
   
	Note that the destination folder must be "7ZipSelfExtractingInstallerCreator"
	
In Visual Studio 2010 ...
2. Go to Project properties
3. Select the Compile tab
4. Click the Build Events ... button
5. Click the Edit Post-build ... button
6. Paste the following line into Post-build Event Command Line text box:

	"$(SolutionDir)7ZipSelfExtractingInstallerCreator\BuildSelfExtractingEXE.vbs" "$(SolutionDir)" "$(TargetDir)" $(TargetFileName) UseVersion=FALSE ReleaseFolder=<path>
	
7. Click OK
8. Click OK to dismiss the Build Events dialog box
9. Build your solution

NOTES:

	UseVersion
		The UseVersion value can be set to either TRUE or FALSE. Can be omitted. If so, default value is FALSE
		
		TRUE =  inserts the major and minor build version into the install path and into the name of 
			    the installer. This is useful when deploying without overwriting 
				previous versions.
				The deployment path will be C:\Program Files\JDSU\<output-executable-name>_<major>.<minor>
				(eg. C:\Program Files\jdsu\JdsuROADMIntegration_1.11)
				The name of installer will be <output-executable-name>_<major>.<minor>_installer.exe
				(eg. JdsuROADMIntegration_1.11_installer.exe)
		
		FALSE = do not use the major and minor build version when naming paths or files. 
			    The deployment path will be C:\Program Files\JDSU\<output-executable-name>
				(eg. C:\Program Files\jdsu\JdsuROADMIntegration)
				The name of installer will be <output-executable-name>_installer.exe
				(eg. JdsuROADMIntegration_installer.exe)
		
		In both cases, any existing files in the deployment path are overwritten without prompting
		upon extraction.
		
	ReleaseFolder
		The full path where the application will be deployed on the target machine. Can be omitted. If
		so, the default path is C:\Program Files\JDSU
	
	Auto-Launching of Deployed Application
		By default, after the self-extraction is completed, the deployed application is _not_ launched.
		To change this behaviour, see comments in the BuildSelfExtractingEXE.vbs file.
		
	File List
		A file named FileList.csv is included as part of the deployment. This file contains a list
		of all files that will be deployed along with their sizes and dates.
		
	Log File
		A file named BuildSelfExtractingEXE.log is created. This file contains a log of the build
		process and is useful for debugging.
		
		
---------------------------------------------------------------------------------------------------
FUNCTIONAL DETAILS
---------------------------------------------------------------------------------------------------
The installer is created by a combination of VBScript and the 7-Zip archive utility with 
SFX extensions.

The VBScript file is a text file and can be edited in a standard text editor. The syntax is very 
similar to Visual Basic 6.

The VBScript, BuildSelfExtractingEXE.vbs, is called as a post-build event in Visual Studio and takes
the following parameters:
(Note that the VBScript could be used standalone if provided appropriate parameters when called)

	SolutionFolder
		The full path to the Visual Studio solution
		$(SolutionFolder) in Visual Studio

	BuildFolder
		The full path to the build output
		$(TargetDir) in Visual Studio
		
	ExeName
		The name of the executable file being deployed
		$(TargetFileName) in Visual Studio
		
	UseVersion
		Either TRUE or FALSE (see Use section above)

The VBScript handles the process of building the installer, configuring and calling 7-Zip when 
needed.

The build process is as follows:
1. Initialize
2. Parse the command-line parameters
3. Get the version of the application executable being deployed
4. Create the release folder name (the location of the deployed application on the target computer).
5. Delete the release folder if it already exists on the local computer
6. Copy the contents of the build folder to the release folder
7. Create the FileList.csv file
8. Create a batch file, named launcher.bat, that will be used to auto-launch the deployed executable
9. Use 7-zip to create a 7-Zip archive of the release folder named temp.7z
10. Create the configuration file, 7z_config.txt, for the self-extractor
11. Build the installer by concatenating the temp.7z archive, the 7z_config.txt, and the 
    7zsd.sfx files
12. Cleanup by deleting temporary files and the release folder


---------------------------------------------------------------------------------------------------
GETTING HELP
---------------------------------------------------------------------------------------------------
Contact Mike Ferris with any questions.


---------------------------------------------------------------------------------------------------
REVISION HISTORY
---------------------------------------------------------------------------------------------------
v1.3	13-Nov-2014		If the release folder is specified on the parameter list, use it explicitly
						and do not use either the solution name or the version in the folder path

v1.2	06-Nov-2014		Bug Fix: Release path in the 7Z_config file needs double '\' in path

v1.1	05-Nov-2014		Add command-line parameter RELEASEFOLDER
						Use parameter RELEASEFOLDER to override default install path 
							C:\Program Files\JDSU
						Add constant to track version and record in log
						Bug Fix when creating local release folder
						Do not delete the local release folder on completion
						
v1.0					Initial Release