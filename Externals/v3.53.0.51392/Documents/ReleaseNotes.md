---
uid: UdbsInterface.ReleaseNotes
---
# UdbsInterface Release Notes
---

See [ProjectInfo](ProjectInfo.md) for more information.

See [Release Notes Guidelines](http://fractal.li.lumentuminc.net/tutorials/articles/ReleaseNotesGuidelines.html) for instructions on how to edit this file.

## Issue Tracking

*[TODO: Add link to [Jira](https://jira.lumentum.com/), with exact project information.]*

---

# Change Log

## Version3.53

### v3.53.0.51395
*(2023-09-05, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 51395`)*

[TMTD-644](https://jira.lumentum.com/browse/TMTD-644) Shared UDBS Backward/Forward Compatibility

Made Utility_SetTemporaryStationName accessibility "Friend" instead of public.

## Version3.52

### v3.52.1.51300
*(2023-08-21, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 51300`)*

[FRA-2059](https://jira.lumentum.com/browse/FRA-2059) Fix Fractal NuGet Packages Interdependency Conflicts

Repackaging every MES Test Data Library and NuGet package.

### v3.52.0.51249
*(2023-07-31, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 51249`)*

[TMTD-613](https://jira.lumentum.com/browse/TMTD-613) CProcessInstance::RestartUnit doesn't always save the right StationName

Added CUtility::Utility_SetTemporaryStationName to set the stationID static variable.
CUtility::Utility_GetStationName now checks for the StationID then checks the machine name.

[TMTD-639](https://jira.lumentum.com/browse/TMTD-639) Test result process flow shows incorrect failure

CTestData_Instance::Finish no longer calls StoreBlobSummaryInfo()
Removed CTestData_Instance::Finish(..)
CTestData_Instance::StoreBlobSummaryInfo is now public.

## Version3.51

### v3.51.0.50780
*(2023-05-24, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 50780`)*

[TMTD-591](https://jira.lumentum.com/browse/TMTD-591) CProcessInstance::Start() write operations to database should be within Transaction


[TMTD-604](https://jira.lumentum.com/browse/TMTD-604) TestDataProcessInfo class: Replace optional parameters with overloads

Fixed breaking change cause by adding optional parameter to TestDataProcessInfo::GetProcessInfo

[TMTD-609](https://jira.lumentum.com/browse/TMTD-609) Bug in TestDataProcessInfo::IsProcessFromThisStation

Added TestDataProcessInfo::IsProcessOwnedByThisStation which compares both the stationID and the machine name to the station name obtained from UDBS.
Marked TestDataProcessInfo::IsProcessFromThisStation as obsolete.

## Version3.50

### v3.50.0.50677
*(2023-05-09, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 50677`)*

[TMTD-531](https://jira.lumentum.com/browse/TMTD-531) Prevent overwrite of Test data of a COMPLETED test

Added a check to UpdateNetworkDB if the status is already COMPLETED in the network database

[TMTD-572](https://jira.lumentum.com/browse/TMTD-572) Add Test Instance Status check in Restart & Pause

A check for PAUSE status is performed before restarting. A check for IN_PROCESS status is performed before pausing.

[TMTD-575](https://jira.lumentum.com/browse/TMTD-575) Create Method to Purge Attached Files at the End of a Process

Added CProcessInstance.DeleteFilesAttachedToUDBS()

[TMTD-588](https://jira.lumentum.com/browse/TMTD-588) Add exception handling in CProcessInstance::DeleteFilesAttachedToUDBS

A message is now logged when file deletion fails.

[FRA-1970](https://jira.lumentum.com/browse/FRA-1970) Store Data Size Aggregate into a Test Process

Aggregate information about archived files and BLOBs is now stored in test data.

[FRA-1972](https://jira.lumentum.com/browse/FRA-1972) Clean up logging in Fractal

Lowering the severity of trace and debug messages.

## Version3.49

### v3.49.0.50510
*(2023-04-13, , `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 50510`)*

[TMTD-571](https://jira.lumentum.com/browse/TMTD-571) Make UdbsInterface LocalDB interactions Thread-Safe

Now using the DatabaseSupport.ExecuteLocalQuery() function that takes a Transaction as a parameter to make LocalDB queries Thread-Safe.

[TMTD-576](https://jira.lumentum.com/browse/TMTD-576) Prevent Creation of Units with Trailing Whitespaces in Serial Number

## Version3.48

### v3.48.0.50360
*(2023-03-22, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 50360`)*

[TMTD-538](https://jira.lumentum.com/browse/TMTD-538) Fix MES Test Data Library's DocFX

[TMTD-554](https://jira.lumentum.com/browse/TMTD-554) Make CProcessInstance.DeleteLocalProcess(...) Public

[TMTD-561](https://jira.lumentum.com/browse/TMTD-561) SQL Error Performing WIP Operations at Fabrinet

Removed the OUTPUT clause from UDBSNetworkDatabase::GenerateInsertSqlRequest function that was used to get the key of the InsertRecord.
Replaced it with SELECT @@IDENTITY clause.

[TMTD-562](https://jira.lumentum.com/browse/TMTD-562) UDBS MES API incorrectly reports inactive Oracle part number for UDBS Part number

[TMTD-568](https://jira.lumentum.com/browse/TMTD-568) Make the CBlob class public again

## Version3.47

### v3.47.0.50193
*(2023-02-24, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 50193`)*

[TMTD-544](https://jira.lumentum.com/browse/TMTD-544) Exception while logging exception in UdbsInterface

DatabaseSupport::LogErrorInDatabase has been fixed to better handle null object reference.
The fix provided before wasn't fully working. Now checking if the the exception's targetSite and its DeclaringType are nullable.

[TMTD-553](https://jira.lumentum.com/browse/TMTD-553) CProcessInstance::StoreProcessInstanceFields doesn't check for UpdateLocalRecord success

CProcessInstance::StoreProcessInstanceFields now checks for the return value of UpdateLocalRecord

## Version3.46

### v3.46.0.50170
*(2023-02-23, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 50170`)*

[TMTD-325](https://jira.lumentum.com/browse/TMTD-325) Replace CTestdata_Instance::LoadExisting calls with TestDataProcessInfo::GetProcessInfo when determining If Test Exists
Replaced LoadExisting call to GetProcessInfo in PrepareTestForRestart

[TMTD-525](https://jira.lumentum.com/browse/TMTD-525) Create function to check if a test process is active on another station

[TMTD-527](https://jira.lumentum.com/browse/TMTD-527) RestartUnit() should update the station DB field

RestartUnit() now updates the process station to the currently executing machine.

[TMTD-529](https://jira.lumentum.com/browse/TMTD-529) Improve Test Data Item List Load Time in UDBS Interface

Reducing the number of queries in order to enhance performances.

[TMTD-530](https://jira.lumentum.com/browse/TMTD-530) Don't save processes' item list to the local DB.

The item list revision is stored in the local DB, but is never read back.

[TMTD-541](https://jira.lumentum.com/browse/TMTD-541) Ability to clear or delete UDBS testdata_result records

Adding the ability to clear results of a group of an ongoing test data process.

[TMTD-544](https://jira.lumentum.com/browse/TMTD-544) Exception while logging exception in UdbsInterface

DatabaseSupport::LogErrorInDatabase has been fixed to better handle null object reference.

[TMTD-548](https://jira.lumentum.com/browse/TMTD-548) Make CWipProcess.Results Part of the Public Interface Again

Access modifier changed from Friend to Public

[FRA-1941](https://jira.lumentum.com/browse/FRA-1941) Update the NuGet Specs for other Fractal and UDBS Libraries

Fixing mistakes in the NuGet packaging in the previous release.

[TMDT-249](https://jira.lumentum.com/browse/TMTD-249) Added missing logs to exceptions

## Version3.45

### v3.45.0.49969
*(2023-01-31, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 49969`)*

[TMTD-527](https://jira.lumentum.com/browse/TMTD-527) RestartUnit() should update the station DB field

RestartUnit() now updates the process station to the currently executing machine.

[FRA-1941](https://jira.lumentum.com/browse/FRA-1941) Update the NuGet Specs for other Fractal and UDBS Libraries

Updating the NuGet package specifications.

## Version3.44

### v3.44.0.49869
*(2023-01-19, , `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 49869`)*

[TMTD-522](https://jira.lumentum.com/browse/TMTD-522) CProcessInstance::TerminateWithoutSynchronizing don't unregister process from local DB

CProcessInstance::TerminateWithoutSynchronizing now unregisters the process from local DB.

## Version3.43

### v3.43.0.49747
*(2022-12-16, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 49747`)*

[TMTD-345](https://jira.lumentum.com/browse/TMTD-345) Transactional integrity when starting WIP CWIP_Process::Begin_Process UDBS Interface

Methods of the CWIP_Process class create transactions, so that the changes are rolled-back if there's a failure along the way.

[TMTD-459](https://jira.lumentum.com/browse/TMTD-459) Implement Amazon S3 BLOB Retrieval in UDBS Interface

Files archived to Amazon S3 will automatically be restored to UDBS when accessed using the CBLOB::GetBLOB(...) method.

[TMTD-480](https://jira.lumentum.com/browse/TMTD-480) Create a class representing process info

Added ProcessInfo, TestDataProcessInfo, WIPprocessInfo and KittingProcessInfo classes.

[TMTD-483](https://jira.lumentum.com/browse/TMTD-483) UDBSIMes' GetUnitDetails(...) should use CProduct.UnitExists(...) now that performance issues have been addressed.

Static version of CProduct.UnitExists(...) created in order to check the presence of a unit without having to load all the product details.

[TMTD-487](https://jira.lumentum.com/browse/TMTD-487) CWIP_Process::LoadResultsCollection logs an error when starting the WIP process

No errors logged when there are no WIP results present for a given WIP process.

[TMTD-508](https://jira.lumentum.com/browse/TMTD-508) Add indication for "Shared" and "Trial" to UdbsInterface version in error_table

Shared assemblies comming from the GAC will now be displayed with a "S" suffix.
Assemblies in trial will now be displayed with a "T" suffix.

[TMTD-511](https://jira.lumentum.com/browse/TMTD-511) Process_info details should be uploaded all at once when finishing test

Added StoreProcessInstanceField() that takes a dictionary as parameter to make the operation atomic.

[TMTD-513](https://jira.lumentum.com/browse/TMTD-513) WIP Process Properties Not Accessible

Revert visibility changes to a the WIP Process' properties.

[FRA-1880](https://jira.lumentum.com/browse/FRA-1880) Fractal GUI shouldn't log stack traces

Stack traces are no longer logged into the Fractal message box and the log files. They are still logged in the UDBS error log.

[FRA-1899](https://jira.lumentum.com/browse/FRA-1899) Fractal apps hang event

Introducing a maximum 2 minutes timeout for acquiring the lock to the local SQLite database when starting a transaction.

Also, when UDBS debugging is turned on, the stack trace of the last successful lock acquisition will be included if this timeout
is exceeded. This will help in identifying what operation is holding onto the lock and causing this problematic situation.

## Version3.42

### v3.42.0.49507
*(2022-11-23, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 49507`)*

[TMTD-462](https://jira.lumentum.com/browse/TMTD-462) Create Synchronize() Method in ITest Interface
Added CProcessInstance::Synchronize to allow upload of local DB data to network DB without changing process status.
Reverted changes from TMTD-427 and TMTD-435:
Removed CProcessInstance::RestartUnitWithPopulatedLocalDB.
Removed "RemoveLocalData" argument of CProcessInstance::PauseProcessInstance.

[TTE-1961](https://jira.lumentum.com/browse/TTE-1961) Failed to start a test process
Local DB compacting on start-up is not a critical operation. Log a warning but do not abort the test start operation if it happens.

[TMTD-510](https://jira.lumentum.com/browse/TMTD-510) UDBS|Error unlocking WIP process
Fix: Wip process is actually unlocked successfully but the operation was wrongfully reported as a failure.

[FRA-1897](https://jira.lumentum.com/browse/FRA-1897) FRA-1897: Fix compilation warnings

## Version3.41

### v3.41.0.49238
*(2022-10-21, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 49238`)*

Built by accident. No new changes compared to v3.40.0.49231.

## Version3.40

### v3.40.0.49231
*(2022-10-21, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 49231`)*

[TMTD-182](https://jira.lumentum.com/browse/TMTD-182) CheckActiveProcesses() requires a restart of the app running to flush local DB
Process data is now uploaded to network DB without needind a restart of the software. CheckActiveProcesses() now takes an "Out" parameter to return the active process ID found in the local DB.

[TMTD-450](https://jira.lumentum.com/browse/TMTD-450) Unregister a process from local Db should be done in UpdateNetworkDB()
Unregister a process is now done within UpdateNetworkDB when remove loca data is allowed.

[FRA-1873](https://jira.lumentum.com/browse/FRA-1873) Refactor IsTestAvailable for restart and move it to the UdbsInterface
ProcessInfo::IsTestAvailableForRestart has been renamed to PrepareTestForRestart.
It now makes a call to CTestData_Instance::PrepareTestForRestart.
Furthermore, the function first checks for an active process in the local Db before looking in the network Db.

[TMTD-458](https://jira.lumentum.com/browse/TMTD-458) Change MasterInterface/clsPrdGrp class functions to be Public

[TMTD-463](https://jira.lumentum.com/browse/TMTD-463) Port VB6 ReadMACAddress(...) Method to UDBS Interface (.NET)
ClsPrdGrp::ReadMACAddresses was added. The function returns MAC IDs and MAC addresses for the specified unit and product group.


## Version3.39

### v3.39.0.49111
*(2022-10-06, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 49111`)*

[TMTD-447](https://jira.lumentum.com/browse/TMTD-447) Implement Microsoft .NET Public API Change Detection Tool

Making a breaking change to the public interface of UDBS will cause a compilation error from now on.

[TMTD-460](https://jira.lumentum.com/browse/TMTD-460) Dissociate UDBS Implementation of MES Test Data Library (Shim) from UDBS Interface DLL

New MesTestData.UDBS assembly marked as a "special friend" to allow access to "internal" functions and classes.


### v3.39.0.48873
*(2022-08-23, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48873`)*

[TMTD-435](https://jira.lumentum.com/browse/TMTD-435) Optimize CTestData.Pause() and RestartUnit() functions for performance

CProcessInstance::Pause(...) now takes an optional argument to specify removal of local DB data. CProcessInstance::RestartUnitWithPopulatedLocalDB() has been added
and takes into consideration that the process instance data is already in the local DB.

[TMTD-436](https://jira.lumentum.com/browse/TMTD-436) Provide ReadOnly access of the ItemListRevID through the process object
This is needed by FUSION to optimize access to its Splice ID.

## Version3.38

### v3.38.0.48723
*(2022-08-02, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48723`)*

[TMTD-431](https://jira.lumentum.com/browse/TMTD-431) Make InitializeNetworkDB call thread safe
-Serialize the calls to InitializeNetworkDB
-Compare SQL Client connection string to detect if the connection string has changed

[TMTD-427](https://jira.lumentum.com/browse/TMTD-427) Optimize CTestData_Instance.RestartUnit(...) for Performance
RestartUnit() doesn't require a new CTestData_instance object instantiation anymore.

## Version3.37

### v3.37.0.48673
*(2022-07-25, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48673`)*

[FRA-1802](https://jira.lumentum.com/browse/FRA-1802) UDBS Test Data Uploader should retrieve Product ID from MES when Creating Unit in Test Data

Adding the notion of "Unit Authority"; is the Test Data system the authority on units?

Rebuild with updated reference:
  - MesTestData.Interfaces v1.17.0.67519

## Version3.36

### v3.36.0.48583
*(2022-07-11, Liam Prieditis, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48583`)*

[TMTD-404](https://jira.lumentum.com/browse/TMTD-404) Unit is not created when starting a test if serial number is not unique

Product ID is being taken into account when determining whether a unit exists or not in UDBS' implementation of ITest.StartTest(...).

Addressing performance problems related to CProduct.UnitExists(...)

## Version3.35

### v3.35.0.48539
*(2022-07-07, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48539`)*

[TMTD-413](https://jira.lumentum.com/browse/TMTD-413) Deprecate (Remove) ITest.CreateUnit()

[TMTD-417](https://jira.lumentum.com/browse/TMTD-417) Failed to Start Test through ITest if Oracle Part Number is not Specified

Reverting refactoring activity.

[TMTD-418](https://jira.lumentum.com/browse/TMTD-418) Add exposure to the test instance result in MesTestData Library

Added the 'TestResultIsPass' property indicating if a test result has passed.

[TMTD-419](https://jira.lumentum.com/browse/TMTD-419) result_stringdata fields gets deleted in UpdateNetworkDB() function

Applied a fix to CProcessInstance.UpdateLocalDB to download all the testdata_result records to the local db correctly.

[TMTD-420](https://jira.lumentum.com/browse/TMTD-420) Make UDBS Interface Methods used by Fusion Public Again

CTestdata_Instance.EvaluateGroup() is made public again.
CWIP_Process.GetNextRecommendedStep() is made public again.
Created InterfaceSupport.IsSuccess() and made the module public.
DatabaseSupport.SetNetworkConnectionString() is made public again.

Rebuild with updated reference:
  - MesTestData.Interfaces v1.16.0.66992

## Version3.34

### v3.34.0.48448
*(2022-06-22, , `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48448`)*

[TMTD-411](https://jira.lumentum.com/browse/TMTD-411) UDBS Implementation of ITest.CreateUnit(...) doesn't work

The method can now be used as designed.

## Version3.33

### v3.33.0.48343
*(2022-06-10, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48343`)*

Rebuild with updated references:
  - MesTestData.Interfaces v1.15.0.65840
  - MesTestData.Models v1.7.0.65840

## Version3.32

### v3.32.0.48343
*(2022-06-10, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48343`)*

[TMTD-340](https://jira.lumentum.com/browse/TMTD-340) Remove stylecop rules override in code. (Part 2) - MesTestdata Interfaces

[TMTD-393](https://jira.lumentum.com/browse/TMTD-393) Method DatabaseSupport.CreateNetworkSqlConnection() Does Not Initialize DB Connection

[TMTD-398](https://jira.lumentum.com/browse/TMTD-398) Thread safety when creating units with the UDBSITest::CreateUnit call

CProduct::AddSNwVar(...) is now thread-safe.

Note: The Jira story mentions a problem with UDBSITest::CreateUnit(...). It turns out this method's implementation calls CProduct::AddSNwVar(...).
Making CProduct::AddSNwVar(...) thread-safe addresses the problem.

[TMTD-401](https://jira.lumentum.com/browse/TMTD-401) CProduct::GetUnit reports misleading error when multiple records exists for the unit with same serial number and product number

Example of the error message when this condition is encountered: "Duplicate entries for serial number: NEOBA00Golden-A1 of product 1-01745."

[TMTD-403](https://jira.lumentum.com/browse/TMTD-403) ITest::GetItemsInGroup should return the top level test items when the group name is null or empty

[TMTD-406](https://jira.lumentum.com/browse/TMTD-406) Store the application name in "unit_report" when creating a unit through UDBS Interface

Application name is now stored in the "unit_report" (Unit Creation Report) column.

Rebuild with updated references:
  - MesTestData.Interfaces v1.14.0.64918
  - MesTestData.Models v1.6.0.65837

## Version3.31

### v3.31.0.48343
*(2022-06-10, SandeepP, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48343`)*

Rebuild with updated reference:
  - MesTestData.Interfaces v1.14.0.64918

## Version3.30

### v3.30.0.48205
*(2022-05-17, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48205`)*

[TMTD-340](https://jira.lumentum.com/browse/TMTD-340) Remove stylecop rules override in code.

[TMTD-384](https://jira.lumentum.com/browse/TMTD-384) Fix instances where methods returning a ReturnCode are not properly checking for success

Error handling improved.

[TMTD-389](https://jira.lumentum.com/browse/TMTD-389) Recent Clean-Up of UDBS Public Interface Did Not Take new TED Tool's Usage of the Assembly into Account

Reverting recent visibility changes for a few methods of the UDBS Interface.

## Version3.29

### v3.29.0.48147
*(2022-05-06, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48147`)*

[TMTD-5](https://jira.lumentum.com/browse/TMTD-5) Prevent insertions of Thai dates in UDBS database from .NET UdbsInterface Lib

All dates that appear to be using the Thai calendar (i.e. > year 2400) are converted to the Common Era year by substracting 543. Day and month remain the same.

[TMTD-254](https://jira.lumentum.com/browse/TMTD-254) ITest.GetProcessInformation<int>(...) throws an exception when no data is available.

The default value of the type (0 in case of an integer) is now returned.

If one suspects no value to be present in the field, and needs to differentiate the value "0" from "no value", then invoke it using a nullable type, i.e.: GetProcessInformation<int?>(...)

[TMTD-284](https://jira.lumentum.com/browse/TMTD-284) Allow the creation of a "SqlConnection" to the Network DB without having to manipulate the Connection String

Introducing new DatabaseSupport.CreateNetworkSqlConnection() method for specialized tools to perform non-standard operations not available through the UDBS and MES Test Data interfaces.

Also making CWIP_Process.LoadActiveProcess(...) public again.

[TMTD-363](https://jira.lumentum.com/browse/TMTD-363) Cannot get BLOB or Files From Ongoing Test Instance

[TMTD-364](https://jira.lumentum.com/browse/TMTD-364) Refactor - CTestdata_Result.GetArray and GetFile

Maintenance improvements. No behavior or interface changes.

[TMTD-366](https://jira.lumentum.com/browse/TMTD-366) Dates not properly formatted in Text and Excel test data reports

Rebuild with updated reference:
  - MesTestData.Interfaces v1.12.0.63976

## Version3.28

### v3.28.0.48037
*(2022-04-14, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 48037`)*

[TMTD-251](https://jira.lumentum.com/browse/TMTD-251) Save UDBS system errors locally

System errors are now cached locally, so that they are not lost if the database connection is not available at the moment the error occurs.

[TMTD-288](https://jira.lumentum.com/browse/TMTD-288) Track version of shared UDBS DLL used by application

UDBS Interface assembly's version (and application) is logged to UDBS error table when DB connection is established.

[TMTD-313](https://jira.lumentum.com/browse/TMTD-313) Some log messages still show DB Connection password

Cleaned up the identified log messages.

[TMTD-314](https://jira.lumentum.com/browse/TMTD-314) MasterInterface: Update/Add units test for public interface

Cleaned up a few problems with the Master Interface.

[TMTD-316](https://jira.lumentum.com/browse/TMTD-316) WIPInterface: Update/Add units test for public interface

Fixed and issue with the CheckUserPrivileges() function that wouldn't work when the user only belongs to one group.

[TMTD-317](https://jira.lumentum.com/browse/TMTD-317) TestDataInterface: Update/Add units test for public interface

Cleaned up a few design problems the Test Data interface.

[TMTD-323](https://jira.lumentum.com/browse/TMTD-323) CUtility.Utility_GetStationName(...) should be Public

Method was wrongfully made "Friend" in a recent clean-up of the UDBS Interface's public interface. This breaks the Common Library.

[TMTD-324](https://jira.lumentum.com/browse/TMTD-324) Test data not uploaded to UDBS

Enhancing error log readability.

[TMTD-330](https://jira.lumentum.com/browse/TMTD-330) Validate Max BLOB File Size

The current size limit is that of a signed 32-bits integer (because of the UDBS BLOB DB column).

[TMTD-342](https://jira.lumentum.com/browse/TMTD-342) UDBS Interface should fail to start a test if that unit is "IN PROCESS" in the local DB and the associated Windows Process is running

The special case of a test "in process" associated with a running Windows process is now handled. Starting a new test or restarting the existing test will fail under that circumstance.

[TMTD-343](https://jira.lumentum.com/browse/TMTD-343) SecurityInterface falsely recognizes nonexistent employee number

Method GetGroupMembership(...) now correctly returns 'False' if it fails to find the user's groups.

[TMTD-357](https://jira.lumentum.com/browse/TMTD-357) Lower "CRAP" Score of UDBS Interface Methods

Refactoring UDBSNetworkDatabase.UpdateRecord(...) and UDBSNetworkDatabase.InsertRecord(...) methods to adopt better design.

Rebuild with updated reference:
  - MesTestData.Interfaces v1.11.0.63602

## Version3.27

### v3.27.0.47823
*(2022-03-04, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 47823`)*

[TMTD-212](https://jira.lumentum.com/browse/TMTD-212) Remove DatabaseSupport.OpenNetworkDBWithDeletePrivileges()

[TMTD-216](https://jira.lumentum.com/browse/TMTD-216) Always Attach BLOBs into the Local DB and synchronize at the end

[TMTD-218](https://jira.lumentum.com/browse/TMTD-218) Remove CProcessInstance::UpdateNetworkDB_1Item(...)

[TMTD-229](https://jira.lumentum.com/browse/TMTD-229) Remove 'Hack mode' (NetworkDb rows > LocalDbRows) from UpdateNetworkDb and replace it with an error message to the udbs_error table

[TMTD-235](https://jira.lumentum.com/browse/TMTD-235) Review and Clean-up Public Interface of UDBS Interface

Marking multiple methods used internally as "Friend" instead of "Public". Removing unused methods.

[TMTD-236](https://jira.lumentum.com/browse/TMTD-236) Review and Clean-up dependencies

SRMInterface (and JDSUFiles consequently) dependencies were removed from UDBSInterface.
Removed UdbsTools::CheckVersion

[TMTD-261](https://jira.lumentum.com/browse/TMTD-261) Renamed DatabaseSupport::DebugMessage to LogErrorInDatabase

[TMTD-265](https://jira.lumentum.com/browse/TMTD-265) Refactoring: Move UpdateNetworkDB(...) methods to a class of its own

[TMTD-268](https://jira.lumentum.com/browse/TMTD-268) Refactoring: CProcessInstance.UpdateNetworkDB(...)

Removed unnecessary Process Name and Process ID parameters; these two values are contained in member variables of the CProcessInstance class.

[TMTD-271](https://jira.lumentum.com/browse/TMTD-271) Incorrect Logic in DatabaseSupport.Releaser

The SQL transaction rollback occurs without the throwing of a new exception.

[TMTD-272](https://jira.lumentum.com/browse/TMTD-272) Application goes into "Hack Mode" during Process Recovery

[TMTD-290](https://jira.lumentum.com/browse/TMTD-290) CWIP_Process::FinishStep returns the wrong returnCode on last step in WIP

The Function now return UDBS_OP_SUCCESS on last step in WIP.

[TMTD-298](https://jira.lumentum.com/browse/TMTD-298) Improve code coverage of UdbsInterface/KittingInterface/KittingUtility.vb

Minor changes to the interface to fix errors.
- Kitting sequence number is a 32-bits integer, not a 64-bits one.
- Multipl parameters of the UpdateKittingData(...) method were wrongfully marked as output parameters.

Refactoring of the UpdateKittingData(...) to reduce excessive complexity. No changes in behavior.

[TMTD-306](https://jira.lumentum.com/browse/TMTD-306) Optimizing block size to speed-up transfer of DB BLOB

Fine-tuning the BLOBs' upload block size to improve file-transfer performances.

[FRA-1701](https://jira.lumentum.com/browse/FRA-1701) Bracketed Serial Number Cannot Initialize if using CIL

Accessing ISpec.Specs following a failure to load the test specifications no longer causes an exception to be thrown.

Also, improved the error log when a given product is not found.

## Version3.26

### v3.26.0.47678
*(2022-01-20, Boris Nzaramba, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 47678`)*

[TMTD-66](https://jira.lumentum.com/browse/TMTD-66) UDBS ITest AddNote(...) Fails Silently

An exception is now thrown when the function fails.

[TMTD-210](https://jira.lumentum.com/browse/TMTD-210) Original Stack Trace Lost in DatabaseSupport::DebugMessage

Overloaded DatabaseSupport::DebugMessage() with the following signature: Sub DebugMessage(functionName, Exception)

[TMTD-219](https://jira.lumentum.com/browse/TMTD-219) Store more information in udbs_error table

Added a Process Context section in the description field of the udbs_error table.
The description field new formatting is as such:

Process: Type= Name= ID= Product= SN= | Application: | {Class}::{Method} | {Message}

[TMTD-221](https://jira.lumentum.com/browse/TMTD-221) Don't log "normal" application errors in UDBS_Error

Removed the insertion of common application errors into the udbs_error table. Such errors are now logged into application log using
the new method LogError().

[TMTD-228](https://jira.lumentum.com/browse/TMTD-228) Replace the "ProcessID/ItemListID" bulk insert by full new rows bulk insert

Uploading every column during the "Build Insert" step of the DB synchronization.

[TMTD-238](https://jira.lumentum.com/browse/TMTD-238) Add Process instance details to classes using DebugMessage()

Add Process details to the description field of the udbs_error table:
Process: Type= Name= ID= Product= SN= | Application: | {Class}::{Method} | {Message}

[TMTD-258](https://jira.lumentum.com/browse/TMTD-258) Convert ProcedureErr to Try/Catch blocks in UdbsInterface

Converted ProcedureErr to Try/Catch blocks

[TMTD-264](https://jira.lumentum.com/browse/TMTD-264) Revert Change: Flush local DB in UdbsITest::LoadExisting

Reverted change made in TMTD-138 that had the local DB flushed in UdbsITest::LoadExisting. That change had introduced a bug.

## Version3.25

### v3.25.1.47695
*(2022-01-21, Boris Nzaramba, `svn+ssh://ottnrp01/AutomationTestSW/VB.NET/FractalLibraries/FractalSteps/trunk Rev 47695`)*

[TMTD-264](https://jira.lumentum.com/browse/TMTD-264) Revert Change: Flush local DB in UdbsITest::LoadExisting

Reverted change made in TMTD-138 that had the local DB flushed in UdbsITest::LoadExisting. That change had introduced a bug.

### v3.25.0.47556
*(2021-12-14, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 47556`)*

[TMTD-167](https://jira.lumentum.com/browse/TMTD-167) Validate that empty serial number does not allow to create unit

Added a check in CProduct.AddSNwVar(..) to make sure the new unit's SerialNumber is not empty.

[TMTD-214](https://jira.lumentum.com/browse/TMTD-214) Clamping the values before writing to NetworkDB

Validate double data for Nan, +Inf, -Inf before insert to network DB.

Software is already filtering data before inserting to local DB, however older unfiltered data in existing local DB fails to be uploaded to networkDB, during data recovery.

[FRA-1622](https://jira.lumentum.com/browse/FRA-1622) Console error logs generated when unit is not active in WIP

The full stack trace from of LoadActiveUnit() and LoadProcessID() is being logged in the application logs and pushed to the UDBS_Errors table when a unit is not active in WIP.

UDBSIMes.vb has been updated to prevent the exceptions from being thrown.

[TMTD-179](https://jira.lumentum.com/browse/TMTD-179) Improve error message on local DB initialization error

An error message is now logged when there is a failure to initlaize the local database.

[TMTD-208](https://jira.lumentum.com/browse/TMTD-208) Increased SQL Deadlocks on uploading testdata_result data

Validate double data for Nan, +Inf, -Inf before insert to network DB.

Software is already filtering data before inserting to local DB, however older unfiltered data in existing local DB fails to be uploaded to networkDB, during data recovery.

[FRA-1575](https://jira.lumentum.com/browse/FRA-1575) Get data system information.

Extended IDataSystem interface with property 'DsInfo', to get datasystem information, i.e., Name (UDBS, Camstar, SQLite) and the connection string.

[TMTD-196](https://jira.lumentum.com/browse/TMTD-196) UDBS Connection information string format should be fixed.

Cleaning UDBS database connection information string format used for logging.

[TMTD-213](https://jira.lumentum.com/browse/TMTD-213) Replace "MERGE" with "INSERT/UPDATE" logic in UpdateNetworkDb call

Replacing the MERGE command that seems to be the root cause of the SQL Server Deadlocks we have been investigating with an UPDATE command. See [TMTD-208](https://jira.lumentum.com/browse/TMTD-208)

[TMTD-225](https://jira.lumentum.com/browse/TMTD-225) Fractal applications not providing application name in connection string for UDBS

Name and version of the top-level application used to fill the Application Name.

[TMTD-247](https://jira.lumentum.com/browse/TMTD-247) Failure to retrieve VB6-encoded String Array BLOBs from UDBS

I/O stream needs to be reset before starting to load the array again using the VB6 format.

Rebuild with updated reference:
  - MesTestData.Interfaces v1.10.0.58931

## Version3.24

### v3.24.0.47172
*(2021-11-03, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 47172`)*

Rebuild with updated reference:
  - MesTestData.Interfaces v1.9.0.58517

## Version3.23

### v3.23.0.47164
*(2021-11-03, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 47164`)*

Rebuild with updated reference:
  - SRMInterface v1.7.0.58522

## Version3.22

### v3.22.0.47050
*(2021-11-01, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 47050`)*

[TMTD-86](https://jira.lumentum.com/browse/TMTD-86) Request for classification of ResultCodes

Extension methods added to the ResultCodes.

[TMTD-169](https://jira.lumentum.com/browse/TMTD-169) Exposing ReadOnly property to ITest

Implementing the ITest.ReadOnly property for UDBS.

[TMTD-172](https://jira.lumentum.com/browse/TMTD-172) Parameter 'deleteLocalFileOnSuccess' ignored in ITest:StoreDBFile(...)

Behavior implemented, as it should have.

[TMTD-173](https://jira.lumentum.com/browse/TMTD-173) Item list (specifications) revision not available to the caller when loading the "latest" revision.

Specifications' revision now available to the caller.

[TMTD-174](https://jira.lumentum.com/browse/TMTD-174) Deadlock on UDBS transaction rollback exception

Releasing of the lock object is now happening in a "Finally" block to ensure it is executed no matter the code path out of the method.

[TMTD-175](https://jira.lumentum.com/browse/TMTD-175) ITest - Exposing Item List (Specifications) Revision through ITest's UnitDetails

Item list revision is now exposed in the unit details.

[TMTD-180](https://jira.lumentum.com/browse/TMTD-180)

- Removed password from all database connection error messages in the application logs.
- Unit tests have been added to test removing the password from the database connection string.

Rebuild with updated references:
  - MesTestData.Interfaces v1.8.0.56950
  - SRMInterface v1.6.0.55251

### v3.21.0.45714
*(2021-09-16, , `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 45714`)*

__BUG FIXES__

[TMTD-150](https://jira.lumentum.com/browse/TMTD-150) FractalDemo : Blob data not saved correctly in UDBS

BigInt fix applied for TestData_result ID unveiled a bug with storage of Blobs in UDBS.
Added int64 versions of IDatabase.InsertRecord() to hande bigInt and store blobs to network db successfully.
Also fixed the bug in CProcessInstance.UpdateNetworkDB() that was causing blobs to be stored incorrectly.

__IMPROVEMENTS__

[TMTD-137](https://jira.lumentum.com/browse/TMTD-137) Manage UDBS Test instance STARTING limbo

Added handling for test instance status = "STARTING" in UDBSItest.StartTest() and cleaned up the function.

[TMTD-138](https://jira.lumentum.com/browse/TMTD-138) Flush locad DB should be performed in UDBSITest.LoadExisting()

CheckActiveProcesses() is now called in UDBSITest.LoadExisting() to flush test data stuck in the local db.

## Version3.20

### v3.20.0.45572
*(2021-08-30, Tariq Omar, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 45572`)*

__BUG FIXES__

[FRA-1533](https://jira.lumentum.com/browse/FRA-1533) Sequence 1 Log File Duplication in Test Data Folder

This change fixes an issue when there is no test instance in UDBS (Sequence = 0) so the testProcessID wouldn't be assigned.

__IMPROVEMENTS__

[TMTD-106](https://jira.lumentum.com/browse/TMTD-106) Implement DSI Interface - WIP Service Interface

 - Added IMes.ReRouteUnitAtWip() implementation to re-route unit for a given wipStep.

[TMTD-27](https://jira.lumentum.com/browse/TMTD-27) Use constant names for ITest.GetProcessInfo keys.

ITest.GetProcessInfo(...) used to take a string literal as parameter.
 - Added replacement method ITest.GetProcessInfo(key As Process_info) with an enumerated process attributes.
 - Marked ITest.GetProcessInfo(key As string) as obsolete.

[TMTD-106](https://jira.lumentum.com/browse/TMTD-106) Implement DSI Interface - WIP Service Interface

 - Added CWip_Utility.TrySoftwareReRoute(...) to return a ReturnCodes. This allows better debugging in case of failure.
 - Marked CWip_Utility.SoftwareReRoute(...) as obsolete.

[TMTD-129](https://jira.lumentum.com/browse/TMTD-129) Implement GetUnitsInWip API request in Camstar IMES / updated for Udbs IMES

 - Added data table structure to Base IMES class for consistency.

Rebuild with updated reference:
  - MesTestData.Interfaces v1.7.0.54005

## Version3.19

### v3.19.0.45297
*(2021-08-04, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/Libraries/UdbsInterface/trunk Rev 45297`)*

__BUG FIXES__

[FRA-1470](https://jira.lumentum.com/browse/FRA-1470) Missing column entries in UDBS recovery commit

Local data lost when test was "Completed" while UDBS was off-line and then recovered.

This fix will merge test data items values from local DB and network DB to ensure data is not lost on recovery.

[TMTD-117](https://jira.lumentum.com/browse/TMTD-117) Missing conversions from KillNullInteger(...) to KillNullLong(...)

Missed three instances of KillNullLong(...) in the initial fix.

## Version3.18

### v3.18.0.45015
*(2021-07-02, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 45015`)*

Rebuild with updated references:
  - MesTestData.Interfaces v1.6.0.52104
  - SRMInterface v1.5.0.52059
  - MesTestData.Models v1.3.0.52104

## Version3.17

### v3.17.2.45722
*(2021-09-16, Alexandre Lemieux, `svn+ssh://ottnrp01/AutomationTestSW/VB.NET/FractalLibraries/FractalSteps/trunk Rev 45722`)*

[TMTD-150](https://jira.lumentum.com/browse/TMTD-150) FractalDemo : Blob data not saved correctly in UDBS

- Merging the fix for TMTD-150 to the 3.17 hot-fix branch.

### v3.17.1.45203
*(2021-07-28, Alexandre Lemieux, `svn+ssh://ottnrp01/AutomationTestSW/VB.NET/FractalLibraries/FractalSteps/trunk Rev 45203`)*

[FRA-1515](https://jira.lumentum.com/browse/FRA-1515) BuildTool - Fractal v6.1.0 incorrect externals mapping to non-existent tags

- Release process to be updated.
- Will release MesTestDataLibrary in Net4.0/Net4.6.2 config and then release Fractal.

[TMTD-117](https://jira.lumentum.com/browse/TMTD-117) Missing conversions from KillNullInteger(...) to KillNullLong(...)

Missed three instances of KillNullLong(...) in the initial fix.

[TMTD-64](https://jira.lumentum.com/browse/TMTD-64) Fix UdbsInterface folder structure

[TMTD-102](https://jira.lumentum.com/browse/TMTD-102) Propagate BigInt Fix to UDBS Interface v2.0.8.39603

### v3.17.0.44828
*(2021-06-15, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 44828`)*

__BUG FIXES__

[TMTD-65](https://jira.lumentum.com/browse/TMTD-65) UDBS Implementation of ITest.ReadDBResultStored(...) and ITest.ReadDBResult(...) are Inconsistent

- ReadDBResultStored is now returning the correct information.
- Confusing property renamed to a more precise term (ResultStored is now ValueStored).
- Old property is still available, for backward compatibility, but is now marked as obsolete.

__IMPROVEMENTS__

[FRA-1465](https://jira.lumentum.com/browse/FRA-1465) Missing test data results in UDBS - Try to Reproduce through Unit Testing

Making the UDBS library more "testable" by allowing unit tests to setup simulate database interaction errors.

[TMTD-48](https://jira.lumentum.com/browse/TMTD-48) Added stylecop to the projects

[TMTD-66](https://jira.lumentum.com/browse/TMTD-66) We now log an error message when invoking AddNote(...) fails.
This does not fully address the Jira issue (an exception should be thrown) but at least the error condition is reported and will appear in the log file.

Rebuild with updated references:
  - MesTestData.Interfaces v1.5.0.51689
  - SRMInterface v1.4.0.51651
  - MesTestData.Models v1.2.0.51689

## Version3.16

### v3.16.0.44744
*(2021-05-28, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 44744`)*

Rebuild with updated reference:
  - SRMInterface v1.3.0.51143

## Version3.15

### v3.15.0.44724
*(2021-05-26, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 44724`)*

__IMPROVEMENTS__

[FRA-1319](https://jira.lumentum.com/browse/FRA-1319) Add Unit tests to validate different recovery sequences on process start

The addition of unit test required some improvements to the UdbsInterface library.

## Version3.14

### v3.14.0.44674
*(2021-05-19, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 44674`)*

__ BUG FIXES__

[FRA-1427](https://jira.lumentum.com/browse/FRA-1427) Port Fractal.ProcessInfoRecoveryTests.vb to run against Local DB

Fixing a Null Reference exception exposed by one of the Fractal unit tests.

## Version3.13

### v3.13.1.44502
*(2021-04-29, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 44502`)*

__IMPROVEMENTS__

[FRA-1410](https://jira.lumentum.com/browse/FRA-1410) Assembly Signing

Using signed versions of the dependencies.

### v3.13.0.44488
*(2021-04-29, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 44488`)*

- Rebuild with updated references:
  - SRMInterface v1.2.0.50219
  - JdsuFiles v1.5.0.50219

__IMPROVEMENTS__

[FRA-1410](https://jira.lumentum.com/browse/FRA-1410) Assembly Signing

Assembly is now be signed.

[FRA-1311](https://jira.lumentum.com/browse/FRA-1311) Make the log level of UDBS error configurable

Added "{App}UdbsErrorLogLevel" configuration setting.

Example xml:

```xml
<setting name="{App}UdbsErrorLogLevel" desc="Log level for UDBS Error logging. Valid levels (in ascending priority) are: Trace, Debug, Info, Error, Fatal" type="string">
	<value>Debug</value>
</setting>
```

[TMTD-10](https://jira.lumentum.com/browse/TMTD-10) Support big Int (Long) data type for UDBS TestData_result ID value

Adding support for 64-bits integer Result IDs table column. This remains backward compatible.

__BUG FIXES__

[TMTD-12](https://jira.lumentum.com/browse/TMTD-12) Unable to Store and Retrieve Arrays Attachments (BLOB)

Code ported from VB6 (CBLOB.vb) was not able to store .NET arrays. It expected COM/OLE Variant arrays. The CBLOB class can now store and retrieve .NET arrays.

## Version3.12

### v3.12.0.44379
*(2021-04-15, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 44379`)*

- Source is newer than Release:
  - Source SVN Revision 44375 > Release 44369 (v3.11.0.44225)
- Rebuild with updated references:
  - SRMInterface v1.1.0.49832
  - MesTestData.Interfaces v1.4.0.48847
  - MesTestData.Models v1.1.2.47681

__IMPROVEMENTS__

[FRA-1401](https://jira.lumentum.com/browse/FRA-1401) Integrate new SRM Interface's Deploy capability in Fractal libraries

The UdbsTools class' CheckVersion(...) method once again has the capability to 'deploy' SRM elements, without relying on the legacy 32-bits DLL.

__BUG FIXES__

[FRA-1325](https://jira.lumentum.com/browse/FRA-1325) Check Integrity of the Local SQL Lite database during the recovery process

Added method create a copy of the local database and delete the original copy

## Version3.11

### v3.11.0.44232
*(2021-03-25, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 44232`)*

__IMPROVEMENTS__

[TEDT-161](https://jira.lumentum.com/browse/TEDT-161) Remove the DeleleGroup function from SecurityInterface

This removes a function from the recently added SecurityInterface.

[TEDT-164](https://jira.lumentum.com/browse/TEDT-164) fix comment in DatabaseSupport.InitializeNetworkDB

A comment in InitializeNetworkDB said that the connection to the network database is only created if it does
not already exist.  This doesn't match what the code actually does -- the implementation always replaced
the existing network connection.

This changes the comment to match what the code really does.  A better solution might be to fix the code
to match the comment, but this would be a change in behaviour.  As an interim step, we should add a
warning to detect cases where the connection string is replaced.

[FRA-1272](https://jira.lumentum.com/browse/FRA-1272) Internal Product Revision Implementation

The IMes:GetUnitDetails() method now provides the optional "InternalProductRevision" field.

The ITest interface was also extended to include the TestStageExists(...) method. This change provides the
implementation of that method for UDBS.

[FRA-1258](https://jira.lumentum.com/browse/FRA-1258) Implement versionCheck x64 feature
The UDBS Interface library is now using the new 64-bits-compatible SRM Interface library to perform SRM Version Check.
Important: SRM Deployment through the new library is not supported yet.

__BUG FIXES__

[FRA-1325](https://jira.lumentum.com/browse/FRA-1325) Check Integrity of the Local SQL Lite database during the recovery process

Checks the integrity of the data in the local database when process is recovered with the following scenarios.
1) Process Id Does not exist in the local database in the process_registration table.
2) Test Data Process does not exist in testdata_process table.
3) Test Data value is missing in the testdata_result table i.e. both string and float values are blank.
4) Test Results table contains no entries.

- Rebuild with updated references:
  - MesTestData.Interfaces v1.4.0.48847
  - MesTestData.Models v1.1.2.47681

## Version3.10

### v3.10.0.44023
*(2021-02-18, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 44023`)*

[TEDT-156](https://jira.lumentum.com/browse/TEDT-156) TEDT-156 Log Database errors to the log file
- Log database errors to the log file so that there is traceability of errors in the log file.

[TEDT-157](https://jira.lumentum.com/browse/TEDT-157) TEDT-157 On Fractal startup log messages if any previous process data exists in Local DB
- Log messages indicating presence of residual data after recovery.

[TEDT-158](https://jira.lumentum.com/browse/TEDT-158) TEDT-158 Increase BulkCopyTimeout for bulk copy transfers from Local DB to Network DB
- Increase bulkcopy timeout to 120 seconds when copying data from local to network DB to account for slower response from database

[TEDT-160](https://jira.lumentum.com/browse/TEDT-160) TEDT-160 Use return code UDBS_OP_SUCCESS for success when updating the network database
- Use correct return code to indicate success of UpdateNetworkDB call

[FRA-1266](https://jira.lumentum.com/browse/FRA-1266) Incorporate UDBS_Security.dll functionality into UDBSInterface
- Replaces functionality of UDBS_Security.dll with directly database access in UdbsInterface

## Version3.9

### v3.9.1.43981
*(2021-02-11, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 43981`)*

__IMPROVEMENTS__

[TEDT-153](https://jira.lumentum.com/browse/TEDT-153) Marking some methods as "Overridable" in order to mock the CTestdata_Instance class.

### v3.9.0.43959
*(2021-02-10, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 43959`)*

__IMPROVEMENTS__

[TEDT-149](https://jira.lumentum.com/browse/TEDT-149)
- Case insensitive lookups for CTestData_instance ins

- Rebuild with updated references:
  - MesTestData.Interfaces v1.3.0.47681
  - MesTestData.Models v1.1.2.47681

## Version3.8

### v3.8.0.43875
*(2021-01-31, Tariq Omar, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 43875`)*

__IMPROVEMENTS__

[TEDT-94](https://jira.lumentum.com/browse/TEDT-94) Camstar TestData API Design & Implementation
   - Implementation of the new IMes::SetBinningCandidates(...) method for UDBS.
   - Adding CProduct::LookupAllPartNumbers(...) method to retrieve the Oracle Part Numbers from the UDBS product ID.

[FRA-1055](https://jira.lumentum.com/browse/FRA-1055) Improve logging for Wip WIP LoadProcessByID failures
   - On load errors, StepData contains two new keys with additional information.

- Rebuild with updated references:
  - MesTestData.Interfaces v1.2.1.47322
  - MesTestData.Models v1.1.1.47322

## Version3.7

### v3.7.0.43803
*(2021-01-20, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 43803`)*

- [FRA-1259](https://jira.lumentum.com/browse/FRA-1259) LookupKittingInfo should be overloaded to also retrieve the revision of the item

- [FRA-1051](https://jira.lumentum.com/browse/FRA-1051) Modify UdbsInterface to store instrument attributes
  - Requires a schema update, disabled by default.

## Version3.6

### v3.6.6.43429

*(2020-11-06, Sandeep Pradhananga, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 43429`)*

- Released with MESTestDataInterfaces version v1.1.5.45648

### v3.6.5.43329
*(2020-10-26, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 43329`)*

[TEDT-47](https://jira.lumentum.com/browse/TEDT-47) CUtility.vb My.Computer.Name does not work in .netCore
- Convert VB .Net specific API call to .Net

### v3.6.4.43057
*(2020-09-24, Tariq Omar, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 43057`)*

[TED-209](https://jira.lumentum.com/browse/TED-209)
- Replaces IMes.UnitDisposition property with a GetUnitDisposition/SetUnitDisposition
- Accepts an optional comment. The employee number is included in the record of the change,

### v3.6.3.43017
*(2020-09-21, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 43017`)*

[TEDT-24](https://jira.lumentum.com/browse/TEDT-24) UdbsInterface to get rid of VisualBasic dll use, to read registry.
- This is to ease calling UdbsInterface from a .net core application. And in future, ease of migration to .net core.

[TEDT-27](https://jira.lumentum.com/browse/TEDT-27) Missing station name on call to StartTest ITest
- Bug fix on UDBS ITest implementation for StartTest

[TEDT-28](https://jira.lumentum.com/browse/TEDT-28) Transaction scope not called in dbNetwork.Execute
- Bug fix on usage of Transaction

### v3.6.2.42938
*(2020-09-14, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 42938`)*

[TED-209](https://jira.lumentum.com/browse/TED-209)
   The Camstar implementation has an endpoint for MoveToMRB. This review adds similar functionality
to the UDBS, but in a more generic way.
As described in TED-209, there are six possible values; WIP, Inventory, Consumed, FailureAnalysis,
Scrap, and Review.
The new UnitDisposition property is able to get or set any of the six value. The MoveToMRB is then
implemented using this more generic function.
The disposition is stored in two tables. Each unit has a value, so the udbs_unit_details table is
used.  This table points to values in the udbs_product_group table, so that one is required as well.
The udbs_product_group creates the 'UNIT_DISP' group. The pg_string_value column is used to specify
the name of the enumerator.  There are currently six potential values (so a maximum of six rows).
When more enumerators are added each will get an additional row in this table the first time that
the value is used.
Values will be entered into the udbs_product_group table as needed. This allows the new enumerators
to be added in the code, without an explicit database update. For history tracking, values are never
removed from the udbs_product_group table.
When the disposition is set for a unit, a new row is added to udbs_unit_details. This keeps the full
history of the unit in the database. The ud_integer_value column is used to record the order of this
history.
The kitting details function still needs to be modified to use and update this value.
I've tested the SQL from a query window (using UDBS_CienaOA) and have run the attached unit tests.



### v3.6.1.42886
*(2020-09-04, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 42886`)*

- Rebuild using MesTestData.Models v1.0.2 nuget package
- Rebuild using MesTestData.Interfaces v1.1.2 nuget package

 ### v3.6.0.42884
*(2020-09-04, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 42884`)*

[TEDT-9](https://jira.lumentum.com/browse/TEDT-9) Data Models Targeting The New Traveller Module
- Rebuild using MesTestData.Models v1.0.1 nuget package
- Rebuild using MesTestData.Interfaces v1.1.1 nuget package

## v3.5

### v3.5.0.42751
*(2020-08-25, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 42751`)*

[TED-86](https://jira.lumentum.com/browse/TED-86) - Update UDBS DLL to include Operation Mode
- Partial Commit (excluding new SQL tables for Op Mode)
- MS-ACCESS local DB support dropped
  - Causes duplicate work for supporting new things, MS-Access can�t support x64
- Speedup of item list results using BulkInsert and BulkUpdate instead of for-loop insert/update
  - Uses SqlBulkCopy class
  - Requires SQL 2008 or higher (Ottawa is 2012)
  - Uses MERGE T-SQL under the hood for bulk UPSERT
  - Uses temp table for Bulk Update/Insert
- Converted a few On Error... syntax, to try-catch
- Compile Warnings reduction


### v3.0.16.42658
*(2020-08-07, Alexandre Lemieux, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 42658`)*

[FRA-979](https://jira.lumentum.com/browse/FRA-979) FRA-979 Make ITest calls to StoreValue case-insensitive
- Bug fix! When SpecItems is called, as it is now Key-Value pair
- Rebuilt with MesTestData.Interfaces v1.0.7.41480

[TEDT-7](https://jira.lumentum.com/browse/TEDT-7) Change ISpec interface, to use dictionary interface to expose Specs
- Reverting parts of the changes to TEDT-7 related to the use of a IDictionary.

### v3.0.15

Build 3.0.15 failed because some configurations are missing from the Solution.

### v3.0.14.42612
*(2020-07-30, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 42612`)*

[TEDT-7](https://jira.lumentum.com/browse/TEDT-7) Change ISpec interface, to use dictionary interface to expose Specs
- Bug fix! When SpecItems is called, as it is now Key-Value pair

Rebuilt with MesTestData.Interfaces v1.0.7.41480

### v3.0.13.42611
*(2020-07-30, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 42611`)*

[FRA-979](https://jira.lumentum.com/browse/FRA-979)
- Make ITest calls to StoreValue case-insensitive
- Regression found on exceeding 255 characters when directly calling CTestdata_Instance::StoreStringData

[TEDT-7](https://jira.lumentum.com/browse/TEDT-7) Change ISpec interface, to use dictionary interface to expose Specs

### v3.0.12.42455
*(2020-07-10, Spencer Belleau, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 42455`)*

[FRA-708](https://jira.lumentum.com/browse/FRA-708) UdbsInterface: Implement CTestdata_Instance::CreateExcelFile/CreateDatFile methods
- Adds in CreateDATFile and CreateExcelFile from the VB6 dlls, as verbatim as possible. Also includes required supporting functions.

### v3.0.11.42390
*(2020-07-07, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 42390`)*

[FRA-972](https://jira.lumentum.com/browse/FRA-972) IMes new API method to send sub-assy to MRB
- Placeholder implementation added

### v3.0.10.42230
*(2020-06-18, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 42230`)*

__BUG FIXES__

[FRA-997](https://jira.lumentum.com/browse/FRA-997) StateMachine ProcessID not passed properly into TestData object
- After test instance is started successfully, should update TestProcessID property
- Override the base class default value to empty string, from yyyyMMddHHmmss

### v3.0.9.41653
*(2020-05-14, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 41653`)*

[FRA-938](https://jira.lumentum.com/browse/FRA-938) Final Test and PD Calibration UDBS errors
- Clamp out of range values going into testdata_result, result_value column
- It is of SQL type float, and does not accept "infinity" values
- Values are clamped Min=Single.Min, Max=Single.Max

### v3.0.8.41382
*(2020-05-04, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 41382`)*

[FRA-925](https://jira.lumentum.com/browse/FRA-925) CWIP_Process::End_Process fails when there is a NaN ActiveDuration or InactiveDuration result value
- Cast NaN to zero in the for-loop accumulation

### v3.0.7.41208
*(2020-04-27, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 41208`)*

[FRA-915](https://jira.lumentum.com/browse/FRA-915) UDBS Library - Kitting Remarks are not returned by reference

### v3.0.6.40601
*(2020-03-23, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 40601`)*

- Rebuild with latest MesTestData.Interfaces v1.0.3.38771

### v3.0.5.40599
*(2020-03-23, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 40599`)*

### v3.0.4.40054
*(2020-02-05, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 40054`)*

[FRA-759](https://jira.lumentum.com/browse/FRA-759) Milliseconds date part showing up in Process Reporter
- Consistent reporting of Start and End dates (with milliseconds)

[FRA-789](https://jira.lumentum.com/browse/FRA-789) Error occurs when uploading the same blob twice.
- Should update existing row for a duplicate blob (CBLOB.vb)

### v3.0.3.40032
 *(2020-02-04, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 40032`)*
- Removed Debugger.Break()

### v3.0.2.39967
*(2020-01-30, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 39967`)*

[FRA-756](https://jira.lumentum.com/browse/FRA-756) New UdbsInterface Multi-Thread failures
 - Added locking on Transactions
 - Bug fix on SQL Transaction commit. Not atomic previously

### v3.0.1.39808
*(2020-01-21, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 39808`)*

[FRA-556](https://jira.lumentum.com/browse/FRA-556)
 - Added 64-bit


### v3.0.0.39801
*(2020-01-21, Eric Panorel, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/MesTestData/UdbsInterface Rev 39801`)*

[FRA-338](https://jira.lumentum.com/browse/FRA-338) Project Janus issues

[FRA-452](https://jira.lumentum.com/browse/FRA-452) Convert to native .Net from ADODB COM

[FRA-458](https://jira.lumentum.com/browse/FRA-458) UDBS Implemention of ISpec for MES TestData

[FRA-459](https://jira.lumentum.com/browse/FRA-459) Implement IMes in UDBSInterface

[FRA-460](https://jira.lumentum.com/browse/FRA-460) Implement ITest in UDBS

[FRA-710](https://jira.lumentum.com/browse/FRA-710)
 - Streaming upload and download to SQL Server and SQLite database
 - Pipelined streams to mitigate I/O bottlenecks
 - Bug fix for notes that are null

[FRA-733](https://jira.lumentum.com/browse/FRA-733)
 - Catch up to trunk requirements

## Version 2.0

### v2.0.4.37261
*(2019-06-07, Eric Panorel, `svn+ssh://pan62372@ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/Fractal/UdbsInterface Rev 37261`)*

[FRA-119](https://jira.lumentum.com/browse/FRA-119) - Include kitting sequence when retrieving Kitting Info

### v2.0.4.37028
*(2019-05-23, Susan French, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/Fractal/UdbsInterface Rev 37028`)*

[FRA-159](https://jira.lumentum.com/browse/FRA-159) - Simplify Imported Conditional References
- Rebuild

### v2.0.4.36760
*(2019-05-06, Susan French, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/Fractal/UdbsInterface Rev 36760`)*

[FRA-159](https://jira.lumentum.com/browse/FRA-159) - Simplify Imported Conditional References

### v2.0.3.36651
*(2019-04-30, Susan French, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/Fractal/UdbsInterface Rev 36651`)*

[FRA-52](https://jira.lumentum.com/browse/FRA-52) - Update NLog Import file to conditionally include LoggingTools for net40

### v2.0.2.36469
*(2019-04-17, Susan French, `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/Fractal/UdbsInterface Rev 36469`)*

[FRA-121](https://jira.lumentum.com/browse/FRA-121) - Reference latest NLog NuGet Package (Major 4)

### v2.0.1.36434
*(2019-04-12, [Your Name], `svn+ssh://ottnrp01/Manufacturing_Test_SW/.NET_Projects/UDBS.NET/UDBS DLLs/branches/Fractal/UdbsInterface Rev 36434`)*

[FRA-7](https://jira.lumentum.com/browse/FRA-7) - Link SVN Revision to Assembly Revisions

[FRA-61](https://jira.lumentum.com/browse/FRA-61) - Update Fractal Projects to Use General Build Scripts

Added Release Notes

## Previous Versions

See SVN Commit Logs












































