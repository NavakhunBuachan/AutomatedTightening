Imports System.Reflection

' *** IMPORTANT ***  
'  
' `AssemblyVersion_Template.md` contains SVN keywords that will be used to set the revision from the SVN commit revision.  
'  
' - Do edit `AssemblyVersion_Template.md` to change assembly version information, and do commit to SVN  
' - Do NOT edit `AssemblyVersion.vb`, as it will be replaced by the pre-build event, and do NOT commit to SVN  
'  
' See [How to Link the Assembly Version to the SVN Revision](http://fractal.li.lumentuminc.net/tutorials/articles/Release_LinkingAssemblyVersionToSvn.html)  
' See [How to Release New Versions of Shared Assemblies](http://fractal.li.lumentuminc.net/tutorials/articles/ReleaseProcess.html) for the complete release process.    
'  
' *** VERSION INFO GUIDELINES ***  
'  
' Version information for an assembly consists of the following four values:  
'   `Major Version`:  Increment for breaking API changes or adding significant new features  
'   `Minor Version`:  Increment for adding functionality in a backwards-compatible manner  
'   `Patch Number`:   Increment for bug fixes or small improvements to existing features  
'   `Revision`:       Linked to the Subversion revision by using a pre-build compile task  
'  
' See [How to Edit Assembly Version Fields](http://fractal.li.lumentuminc.net/tutorials/articles/Release_EditingAssemblyVersionFields.html)  
'   
' *** EDIT VERSION FIELDS BELOW ***  
' *** (AssemblyVersion, AssemblyInformationalVersion, AssemblyFileVersion) ***  
  
' `AssemblyVersion` is used by .NET for referencing  
' - Manually edit the `Major` and `Minor` versions, according to guidelines above  
' - Keep `Build` and `Revision` = 0   

<Assembly: AssemblyVersion("3.53.0.0")>  
  
' `AssemblyInformationalVersion` will be displayed as "Product Version" in the file details   
' - Manually set the `Major.Minor` to match the `AssemblyVersion` above (useful for information in file system)  

<Assembly: AssemblyInformationalVersion("3.53.0.0")>  
  
' `AssemblyFileVersion` is the version number given to a file as in file system, as displayed by Windows Explorer.   
' It is never used by .NET framework or runtime for referencing.  
' - Manually set the `Major.Minor` to match the `AssemblyVersion` above  
' - Manually set the `Patch`, according to guidelines above  
' - Revision = `$WCREV$$WCMODS?*:$`, where $WCREV$ = highest SVN commit revision in the working copy.   
'   If appended by `*`, it indicates that there are local modifications to the working copy.  
'   If appended by `?`, it indicates that the pre-build script was not able to find the SVN revision.

<Assembly: AssemblyFileVersion("3.53.0.$WCREV$$WCMODS?*:$")>  
  
' *** IMPORTANT *** Do not forget to update RELEASE NOTES !! ***  
'  
' After updating the version, make sure to insert a new section in the `ReleaseNotes.md` file.   
' Copy and paste the template below at the top of the Change Log section, replacing `X.Y.Z` and `[Your Name]`, and removing leading comments (') :  
'  
' ### vX.Y.Z.$WCREV$  
' *($WCDATE=%Y-%m-%d$, [Your Name], `$WCURL$ Rev $WCREV$`)*  
'  
' - High-level description of changes, or list of bugs/improvements/features addressed.  
' - [Jira Number](https://jira.lumentum.com/browse/Jira-Number) Example of Jira Issue  
'  
' See [How to Edit Release Notes](http://fractal.li.lumentuminc.net/tutorials/articles/ReleaseNotesGuidelines.html) 
