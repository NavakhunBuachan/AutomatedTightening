---
uid: UdbsInterface.ProjectInfo
---
# UdbsInterface Project Info 
---

**Project Name:**          UdbsInterface ([README](ReadMe.md), [Release Notes](ReleaseNotes.md))   
**Root Namespace:**        UdbsInterface ([API Reference](xref:UdbsInterface))  
**Project GUID:**          `{C18E3D46-B578-4C10-8294-8C46581C8878}`  
**SVN Source URL:**        `/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk/`  
**SVN Commit Revision:**   51392  
**SVN Last Commit:**       9/5/2023 1:44:22 PM  

# Assembly Info

**AssemblyTitle:**         UdbsInterface  
**AssemblyCompany:**       Lumentum  
**AssemblyProduct:**       UdbsInterface  
**AssemblyCopyright:**     Copyright Lumentum 2023  
**AssemblyVersion:**       3.53.0.0  
**AssemblyFileVersion:**   3.53.0.$$  


# Release Info 

**Release Version:**       3.53.0.51392  
**Build Required:**        No  
**Build Configurations:**  net40-x64, net40-x86, net462-x64, net462-x86  
**SVN Release URL:**       `^/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk/bin_version/`  
**SVN Release Revision:**  51396  
**SVN Release Date:**      9/5/2023 1:56:59 PM  

# SVN Externals

```text
/AutomationTestSW/BatchScripts/FractalBuild@75968 .BuildScripts
/AutomationTestSW/VB.NET/Libraries/LoggingTools/tags/v1.1.0.0/bin_version Externals/LoggingTools
```

# Dependencies

This section lists the dependencies for the configurations defined in the UdbsInterface solution, at the SVN URL below.  The dependencies may be different when included in other solutions.

## UdbsInterface Solution

See <a href="UdbsInterface_DependencyGraph.svg" target="_blank">UdbsInterface Dependency Graph</a>  

**Solution SVN URL:** `^/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk/ Revision 51392`  
**Configurations:**   net40-x64, net40-x86, net462-x64, net462-x86  

### Project Dependencies (Config net40-x64, net40-x86)

| Dependency Name                          | Version     | Type                                     | Details                                  |
| ---------------------------------------- | ----------- | ---------------------------------------- | ---------------------------------------- |
| ClosedXML.Signed                         | 0.95.4      | NuGet                                    | ` NuGet ClosedXML.Signed 0.95.4` |
| LoggingTools                             | 1.1.0.0     | Assembly                                 | `/AutomationTestSW/VB.NET/Libraries/LoggingTools/trunk/bin_version/[config]/LoggingTools.dll` |
| Microsoft.CodeAnalysis.PublicApiAnalyzers | 3.3.3       | NuGet                                    | ` NuGet Microsoft.CodeAnalysis.PublicApiAnalyzers 3.3.3` |
| NLog                                     | 2.0.0.0     | Assembly                                 | `/AutomationTestSW/VB.NET/Libraries/LoggingTools/trunk/bin_version/[config]/NLog.dll` |
| StyleCop.Analyzers                       | 1.1.118     | NuGet                                    | ` NuGet StyleCop.Analyzers 1.1.118` |
| System.Data.SQLite.Core                  | 1.0.112     | NuGet                                    | ` NuGet System.Data.SQLite.Core 1.0.112` |


### Project Dependencies (Config net462-x64, net462-x86)

| Dependency Name                          | Version     | Type                                     | Details                                  |
| ---------------------------------------- | ----------- | ---------------------------------------- | ---------------------------------------- |
| ClosedXML.Signed                         | 0.95.4      | NuGet                                    | ` NuGet ClosedXML.Signed 0.95.4` |
| Microsoft.CodeAnalysis.PublicApiAnalyzers | 3.3.3       | NuGet                                    | ` NuGet Microsoft.CodeAnalysis.PublicApiAnalyzers 3.3.3` |
| NLog                                     | 4.7.15.867  | NuGet                                    | ` NuGet NLog 4.*` |
| NLog.Windows.Forms                       | 4.6.0.451   | NuGet                                    | ` NuGet NLog.Windows.Forms 4.*` |
| StyleCop.Analyzers                       | 1.1.118     | NuGet                                    | ` NuGet StyleCop.Analyzers 1.1.118` |
| System.Data.SQLite.Core                  | 1.0.112     | NuGet                                    | ` NuGet System.Data.SQLite.Core 1.0.112` |


# Consumers

This project is referenced by the projects in the solutions listed in the sections below, valid as of the SVN Revision.

## Consumer Solution:  MEMSXMLSiriusParser (trunk)

**Solution SVN URL:** `^/C%23/Libraries/MEMSTestSiriusParser/trunk/ Revision 67841`  
**Solution Info:** [MEMSXMLSiriusParser](xref:MEMSXMLSiriusParser.Documents.SolutionInfo)

### Consumer Projects (Config net462-source-x64, net462-source-x86)

| Consumer Project Name                    | Version     | SVN URL                                  |
| ---------------------------------------- | ----------- | ---------------------------------------- | 
| [UdbsTestUploader](xref:UdbsTestUploader.ProjectInfo) | 1.14.0.65046 | `^/C%23/Applications/UdbsTestUploader/trunk/UdbsTestUploader.csproj Revision 67845` |


## Consumer Solution:  MesTestDataLibrary (trunk)

**Solution SVN URL:** `^/C%23/Solutions/MesTestDataLibrary/trunk/ Revision 47422`  
**Solution Info:** [MesTestDataLibrary](xref:MesTestDataLibrary.Documents.SolutionInfo)

### Consumer Projects (Config net40-source-x64, net40-source-x86, net462-source-x64, net462-source-x86)

| Consumer Project Name                    | Version     | SVN URL                                  |
| ---------------------------------------- | ----------- | ---------------------------------------- | 
| [JdsuUdbsLibrary](xref:JDSU.UdbsLibrary.ProjectInfo) | 5.44.0.74219 | `^/VB.NET/Libraries/JdsuUdbsLibrary/branches/MESTestData/JdsuUdbsLibrary.vbproj Revision 75411` |
| [MesTestData.UDBS](xref:MesTestData.UDBS.ProjectInfo) | 1.13.0.75390 | `^/VB.NET/Libraries/MesTestData.UDBS/trunk/MesTestData.UDBS.vbproj Revision 75422` |
| [CommonUnitTestingUtilities](xref:CommonUnitTestingUtilities.ProjectInfo) | 1.34.0.74217 | `^/C%23/Libraries/CommonUnitTestingUtilities/trunk/CommonUnitTestingUtilities.csproj Revision 75406` |
| [MesTestData.Library](xref:MesTestDataLibrary.ProjectInfo) | 1.35.0.74208 | `^/VB.NET/FractalLibraries/MesTestData.Library/trunk/MesTestData.Library.vbproj Revision 75425` |


## Consumer Solution:  UdbsTestUploader (trunk)

**Solution SVN URL:** `^/C%23/Solutions/UdbsTestUploader/trunk/ Revision 64072`  
**Solution Info:** [UdbsTestUploader](xref:UdbsTestUploader.Documents.SolutionInfo)

### Consumer Projects (Config net462-source-x64, net462-source-x86)

| Consumer Project Name                    | Version     | SVN URL                                  |
| ---------------------------------------- | ----------- | ---------------------------------------- | 
| [UdbsTestUploader](xref:UdbsTestUploader.ProjectInfo) | 1.15.0.$$   | `^/C%23/Applications/UdbsTestUploader/trunk/UdbsTestUploader.csproj Revision 67845` |


## Consumer Solution:  Fractal (trunk)

**Solution SVN URL:** `^/VB.NET/Solutions/Fractal/trunk/ Revision 29527`  
**Solution Info:** [Fractal](xref:Fractal.Documents.SolutionInfo)

### Consumer Projects (Config net462-source-x64, net462-source-x86)

| Consumer Project Name                    | Version     | SVN URL                                  |
| ---------------------------------------- | ----------- | ---------------------------------------- | 
| [Fractal](xref:Fractal.ProjectInfo)      | 7.15.0.$$   | `^/VB.NET/FractalLibraries/Fractal/trunk/Fractal.vbproj Revision 76098` |
| [JdsuUdbsLibrary](xref:JDSU.UdbsLibrary.ProjectInfo) | 5.44.1.75411 | `^/VB.NET/Libraries/JdsuUdbsLibrary/branches/MESTestData/JdsuUdbsLibrary.vbproj Revision 75794` |
| [CommonUnitTestingUtilities](xref:CommonUnitTestingUtilities.ProjectInfo) | 1.34.1.75406 | `^/C%23/Libraries/CommonUnitTestingUtilities/trunk/CommonUnitTestingUtilities.csproj Revision 75786` |
| [Fractal.Steps](xref:Fractal.Steps.ProjectInfo) | 2.27.0.$$   | `^/VB.NET/FractalLibraries/FractalSteps/trunk/FractalSteps.vbproj Revision 76060` |
| [MesTestData.Library](xref:MesTestDataLibrary.ProjectInfo) | 1.35.1.75425 | `^/VB.NET/FractalLibraries/MesTestData.Library/trunk/MesTestData.Library.vbproj Revision 75809` |
| [MesTestData.UDBS](xref:MesTestData.UDBS.ProjectInfo) | 1.13.1.$$   | `^/VB.NET/Libraries/MesTestData.UDBS/trunk/MesTestData.UDBS.vbproj Revision 75976` |


## Consumer Solution:  FractalDemo (trunk)

**Solution SVN URL:** `^/VB.NET/Solutions/Fractal/trunk/ Revision 29527`  
**Solution Info:** [FractalDemo](xref:FractalDemo.Documents.SolutionInfo)

### Consumer Projects (Config net462-source-x86)

| Consumer Project Name                    | Version     | SVN URL                                  |
| ---------------------------------------- | ----------- | ---------------------------------------- | 
| [Fractal](xref:Fractal.ProjectInfo)      | 0.10.0.$$   | `^/VB.NET/FractalLibraries/Fractal/trunk/Fractal.vbproj Revision 32414` |
| [JdsuUdbsLibrary](xref:JDSU.UdbsLibrary.ProjectInfo) | 4.5.7.$$    | `^/VB.NET/Libraries/JdsuUdbsLibrary/branches/NET_DLL_Development/JdsuUdbsLibrary.vbproj Revision 32411` |


## Consumer Solution:  FractalDemo (trunk)

**Solution SVN URL:** `^/VB.NET/Solutions/FractalDemo/trunk/ Revision 45717`  
**Solution Info:** [FractalDemo](xref:FractalDemo.Documents.SolutionInfo)

### Consumer Projects (Config net462-source-x64, net462-source-x86)

| Consumer Project Name                    | Version     | SVN URL                                  |
| ---------------------------------------- | ----------- | ---------------------------------------- | 
| [Fractal](xref:Fractal.ProjectInfo)      | 7.1.0.$$    | `^/VB.NET/FractalLibraries/Fractal/trunk/Fractal.vbproj Revision 67229` |
| [JdsuUdbsLibrary](xref:JDSU.UdbsLibrary.ProjectInfo) | 5.27.0.67050 | `^/VB.NET/Libraries/JdsuUdbsLibrary/branches/MESTestData/JdsuUdbsLibrary.vbproj Revision 67146` |
| [CommonUnitTestingUtilities](xref:CommonUnitTestingUtilities.ProjectInfo) | 1.17.0.66702 | `^/C%23/Libraries/CommonUnitTestingUtilities/trunk/CommonUnitTestingUtilities.csproj Revision 67144` |
| [FractalSteps](xref:Fractal.Steps.ProjectInfo) | 2.12.0.$$   | `^/VB.NET/FractalLibraries/FractalSteps/trunk/FractalSteps.vbproj Revision 67229` |
| [FractalDemo](xref:Fractal.Demo.ProjectInfo) | 2.12.0.$$   | `^/VB.NET/FractalTemplates/Applications/FractalDemoApplication/trunk/FractalDemo.vbproj Revision 67231` |
| [FileTransferLibrary](xref:FileTransferLibrary.ProjectInfo) | 1.3.0.66733 | `^/C%23/Libraries/FileTransferLibrary/trunk/FileTransferLibrary.csproj Revision 67137` |
| [MesTestData.Library](xref:MesTestDataLibrary.ProjectInfo) | 1.16.0.$$   | `^/VB.NET/FractalLibraries/MesTestData.Library/trunk/MesTestData.Library.vbproj Revision 67228` |



***
*Auto-generated by Lumentum FractalBuildTool 1.40.0.73176 (9/5/2023 2:10:52 PM)*
