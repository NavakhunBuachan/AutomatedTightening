Option Explicit On
Option Strict On
Option Compare Text
Option Infer On

Imports System.Runtime.CompilerServices
Imports UdbsInterface.MasterInterface


Namespace TestDataInterface
    ' Result Enumeration
    Public Enum ResultCodes
        UDBS_ERROR = -1000000
        UDBS_SPECS_NONE = 11
        UDBS_SPECS_PASS = 10
        UDBS_SPECS_PASS_INC = 1010
        UDBS_SPECS_WARNING = -20
        UDBS_SPECS_WARNING_HI = -21
        UDBS_SPECS_WARNING_LO = -22
        UDBS_SPECS_WARNING_INC = -1020
        UDBS_SPECS_FAIL = -30
        UDBS_SPECS_FAIL_HI = -31
        UDBS_SPECS_FAIL_LO = -32
        UDBS_SPECS_FAIL_INC = -1030
        UDBS_SPECS_SANITY = -40
        UDBS_SPECS_SANITY_HI = -41
        UDBS_SPECS_SANITY_LO = -42
        UDBS_SPECS_SANITY_INC = -1040
    End Enum

    ''' <summary>
    ''' Module adding extension methods to the ResultCodes enumerated type.
    ''' </summary>
    ''' <remarks>
    ''' IMPORTANT!
    ''' These extension methods used to be a middle-layer calling the
    ''' implementation from the MesTestData.Interfaces library.
    ''' As it turns out, the dependency between UdbsInterface And
    ''' MesTestData.Interfaces as been severed.
    ''' The logic of those methods Is now duplicated in both UdbsInterface
    ''' and MesTestData.Interfaces.
    ''' Any changes made to these methods should be replicated in the
    ''' MesTestData.Interfaces library.
    ''' </remarks>
    Friend Module Extensions
        ''' <summary>
        ''' Whether or not this code represents a success.
        ''' </summary>
        ''' <remarks>
        ''' Most test checks for failure. If a test specification has
        ''' not min/max value associated with it, the test result will
        ''' be 'NONE'. This is not a failure. In fact, it's treated as
        ''' success.
        ''' </remarks>
        <Extension()>
        Public Function IsSuccess(Result As ResultCodes) As Boolean
            Select Case Result
                Case ResultCodes.UDBS_SPECS_PASS
                    Return True
                Case ResultCodes.UDBS_SPECS_PASS_INC
                    Return True
                Case ResultCodes.UDBS_SPECS_NONE
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' Whether or not this code represents a complete test.
        ''' </summary>
        <Extension()>
        Public Function IsComplete(Result As ResultCodes) As Boolean
            Select Case Result
                Case ResultCodes.UDBS_SPECS_PASS_INC
                    Return False
                Case ResultCodes.UDBS_SPECS_WARNING_INC
                    Return False
                Case ResultCodes.UDBS_SPECS_FAIL_INC
                    Return False
                Case ResultCodes.UDBS_SPECS_SANITY_INC
                    Return False
                Case ResultCodes.UDBS_ERROR
                    Return False
                Case Else
                    Return True
            End Select
        End Function

        ''' <summary>
        ''' Whether or not this code represents a warning.
        ''' </summary>
        <Extension()>
        Public Function IsWarning(Result As ResultCodes) As Boolean
            Select Case Result
                Case ResultCodes.UDBS_SPECS_WARNING
                    Return True
                Case ResultCodes.UDBS_SPECS_WARNING_HI
                    Return True
                Case ResultCodes.UDBS_SPECS_WARNING_INC
                    Return True
                Case ResultCodes.UDBS_SPECS_WARNING_LO
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' Whether or not this code represents a sanity failure.
        ''' </summary>
        <Extension()>
        Public Function IsSanity(Result As ResultCodes) As Boolean
            Select Case Result
                Case ResultCodes.UDBS_SPECS_SANITY
                    Return True
                Case ResultCodes.UDBS_SPECS_SANITY_HI
                    Return True
                Case ResultCodes.UDBS_SPECS_SANITY_INC
                    Return True
                Case ResultCodes.UDBS_SPECS_SANITY_LO
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' Whether or not this code represents a failure.
        ''' </summary>
        <Extension()>
        Public Function IsFailure(Result As ResultCodes) As Boolean
            Select Case Result
                Case ResultCodes.UDBS_SPECS_FAIL
                    Return True
                Case ResultCodes.UDBS_SPECS_FAIL_HI
                    Return True
                Case ResultCodes.UDBS_SPECS_FAIL_INC
                    Return True
                Case ResultCodes.UDBS_SPECS_FAIL_LO
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' Whether or not this code represents a value below the valid range.
        ''' </summary>
        <Extension()>
        Public Function IsLow(Result As ResultCodes) As Boolean
            Select Case Result
                Case ResultCodes.UDBS_SPECS_SANITY_LO
                    Return True
                Case ResultCodes.UDBS_SPECS_WARNING_LO
                    Return True
                Case ResultCodes.UDBS_SPECS_FAIL_LO
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' Whether or not this code represents a value above the valid range.
        ''' </summary>
        <Extension()>
        Public Function IsHigh(Result As ResultCodes) As Boolean
            Select Case Result
                Case ResultCodes.UDBS_SPECS_SANITY_HI
                    Return True
                Case ResultCodes.UDBS_SPECS_WARNING_HI
                    Return True
                Case ResultCodes.UDBS_SPECS_FAIL_HI
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' Whether or not this code represents an error.
        ''' </summary>
        <Extension()>
        Public Function IsError(Result As ResultCodes) As Boolean
            Select Case Result
                Case ResultCodes.UDBS_ERROR
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' Whether or not this code represents no specifications.
        ''' </summary>
        <Extension()>
        Public Function IsNone(Result As ResultCodes) As Boolean
            Select Case Result
                Case ResultCodes.UDBS_SPECS_NONE
                    Return True
                Case Else
                    Return False
            End Select
        End Function
    End Module

    Public Class CTestdata_Instance
        Inherits CProcessInstance

        ' Results Collection
        Private mResults As New Dictionary(Of String, CTestData_Result)(StringComparer.OrdinalIgnoreCase)

        ' Object State

        '**********************************************************************
        '* Properties
        '**********************************************************************

        ' Results Collection
        Public ReadOnly Property Results As Dictionary(Of String, CTestData_Result)
            Get
                Return mResults
            End Get
        End Property

        Public ReadOnly Property Results(item As String) As CTestData_Result
            Get
                Return mResults(item)
            End Get
        End Property

        '**********************************************************************
        '* Methods
        '**********************************************************************

        ''' <summary>
        ''' Start a test.
        ''' </summary>
        ''' <param name="ProductNumber">The product number.</param>
        ''' <param name="SerialNumber">The serial number of the unit to test.</param>
        ''' <param name="ProcessStage">The test stage to run.</param>
        ''' <param name="StageRevision">
        ''' The stage revision to use.
        ''' Set to 0 to use the latest revision.
        ''' This is meant to be a parameter "by reference". See TMTD-338.
        ''' In the meantime, use property "Revision" following this method call to retrieve
        ''' the sequence if you need it.
        ''' </param>
        ''' <returns></returns>
        Public Function Start(ProductNumber As String,
                              SerialNumber As String,
                              ProcessStage As String,
                              StageRevision As Integer) _
            As ReturnCodes
            ' Starts a testdata process instance
            Try
                Const fncName = "CTestdata_Instance::Start"
                If mREADONLY = True Or ProcessInstanceRunning = True Then
                    ' Process object already loaded/running
                    LogError(New Exception("Process is read only or process is already loaded."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                If StartProcessInstance(ProductNumber, SerialNumber, ProcessStage, StageRevision, mProcessID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Load up items collection
                Dim loadResult = LoadItemsCollection()
                If loadResult <> ReturnCodes.UDBS_OP_SUCCESS Then
                    logger.Warn($"{fncName}: {loadResult}")
                End If

                mREADONLY = False

                ' store incomplete fail to the overall result to start with
                ' this way we will get an overall result even if someone opens and immediately
                ' closes a test instance without saving any data
                Return StoreProcessInstanceField("result", CStr(ResultCodes.UDBS_SPECS_FAIL_INC))
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Completes the ongoing test, storing BLOB summary info.
        ''' </summary>
        ''' <returns>The outcome of this operation.</returns>
        ''' <remarks>
        ''' This method overrides <see cref="Finish(Boolean)"/> and is not using
        ''' an optional parameter instead because the signature would change and
        ''' this would cause a linker error when the UDBS interface library is
        ''' supplied through the GAC.
        ''' For the same reason, this method remains marked as Overridable but
        ''' class deriving from it really should be overriding the other method.
        ''' </remarks>
        Public Overridable Function Finish() As ReturnCodes
            Return Finish(False)
        End Function

        ''' <summary>
        ''' Completes the ongoing test.
        ''' </summary>
        ''' <param name="storeBlobSummaryInfo">Whether or not to store the BLOB summary info to the item list prior to completing the test.</param>
        ''' <returns>The outcome of this operation.</returns>
        Public Overridable Function Finish(storeBlobSummaryInfo As Boolean) As ReturnCodes
            Try
                If Not ProcessInstanceRunning Then
                    ' Process object not running
                    LogError(New Exception("Process is not currently running."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                If storeBlobSummaryInfo Then
                    logger.Warn("The 'storeBlobSummaryInfo' parameter was set to True but should be set to False. The storeBlobSummaryinfo operation is handled in another function and shouldn't be performed within the Finish operation.")
                    Me.StoreBlobSummaryInfo()
                End If

                If StopProcessInstance() <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                mREADONLY = True

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Store BLOB summary information into specific items under the 'General Info' section of the
        ''' item list revision.
        ''' Warnings are logged if the items are not present.
        ''' </summary>
        Friend Sub StoreBlobSummaryInfo()
            Dim count As Integer = 0
            Dim totalCompressedSize As Long = 0
            Dim totalUncompressedSize As Long = 0

            CBLOB.GetLocalBlobSize(Me.Process, Me.ID, count, totalCompressedSize, totalUncompressedSize)

            ' Those test data items are optional, but from a test data traceability point
            ' of view, they give a good visibility on data size and usage, so a warning
            ' will be posted if they do not exist in the item list.
            ' Failure to store will not be reported to the calling client since this is
            ' not a critical part of completing a test.

            Const BLOB_COUNT As String = "blob_count"
            Const BLOB_COMPRESSED_SIZE As String = "blob_compressed_size"
            Const BLOB_RAW_SIZE As String = "blob_raw_size"


            If Results.ContainsKey(BLOB_COUNT) Then
                StoreValue(BLOB_COUNT, count)
            Else
                logger.Warn($"Test data item '{BLOB_COUNT}' does not exist. Cannot store value '{count}' to testdata item.")
            End If

            If Results.ContainsKey(BLOB_COMPRESSED_SIZE) Then
                StoreValue(BLOB_COMPRESSED_SIZE, totalCompressedSize)
            Else
                logger.Warn($"Test data item '{BLOB_COMPRESSED_SIZE}' does not exist. Cannot store value '{totalCompressedSize}' to testdata item.")
            End If

            If Results.ContainsKey(BLOB_RAW_SIZE) Then
                StoreValue(BLOB_RAW_SIZE, totalUncompressedSize)
            Else
                logger.Warn($"Test data item '{BLOB_RAW_SIZE}' does not exist. Cannot store value '{totalUncompressedSize}' to testdata item.")
            End If
        End Sub

        ''' <summary>
        ''' Pauses a process instance temporary and upload all process data to the network DB. The test instance can be restarted
        ''' </summary>
        ''' <returns>Udbs return codes.</returns>
        Public Function Pause() As ReturnCodes
            Try
                If ProcessInstanceRunning = False Then
                    ' Process object not running
                    LogError(New Exception("Process is not currently running."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                If PauseProcessInstance() <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                mREADONLY = True

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Restart a paused process by process description
        ''' </summary>
        ''' <param name="ProductNumber">ProductNumber.</param>
        ''' <param name="SerialNumber">Unit's SerialNumber.</param>
        ''' <param name="ProcessStage">ProcessStage.</param>
        Public Overloads Function RestartUnit(ProductNumber As String,
                                            SerialNumber As String,
                                            ProcessStage As String) _
            As ReturnCodes

            If mREADONLY = True Or ProcessInstanceRunning = True Then
                ' Process object already loaded/running
                LogError(New Exception("Process is read only or process is already loaded."))
                Return ReturnCodes.UDBS_ERROR
            End If

            If MyBase.RestartUnit(ProcessStage, ProductNumber, SerialNumber) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_ERROR
            End If

            ' Load up items collection
            If LoadItemsCollection() <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_ERROR
            End If

            ' Get exiting results
            If LoadExistingResults() <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_ERROR
            End If

            mREADONLY = False

            Return ReturnCodes.UDBS_OP_SUCCESS
        End Function

        Public Overrides Function RestartProcessID(ProcessID As Integer) _
            As ReturnCodes
            ' restart a paused process by process id
            If mREADONLY = True Or ProcessInstanceRunning = True Then
                ' Process object already loaded/running
                LogError(New Exception("Process is read only or process is already loaded."))
                Return ReturnCodes.UDBS_ERROR
            End If

            If MyBase.RestartProcessID(ProcessID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_ERROR
            End If

            ' Load up items collection
            If LoadItemsCollection() <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_ERROR
            End If

            ' Get exiting results
            If LoadExistingResults() <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_ERROR
            End If

            mREADONLY = False

            Return ReturnCodes.UDBS_OP_SUCCESS
        End Function

        ''' <summary>
        ''' Retrieves existing test process and result data and loads it into this object.
        ''' </summary>
        ''' <param name="Stage">Stage name.</param>
        ''' <param name="ProductNumber">Product number.</param>
        ''' <param name="SerialNumber">Serial number.</param>
        ''' <param name="Sequence">
        ''' (In/Out) Test sequence. Is set to zero, the 'latest' sequence is loaded
        ''' And this parameter Is assigned the value representing the latest test sequence.
        ''' </param>
        ''' <param name="ConnectionString">The connection string to use to connect to the DB.</param>
        ''' <returns>Operation result <see cref="ReturnCodes"/>.</returns>
        Public Function LoadExisting(Stage As String,
                                     ProductNumber As String,
                                     SerialNumber As String,
                                     ByRef Sequence As Integer,
                                     Optional ByVal ConnectionString As String = "") As ReturnCodes
            If ConnectionString <> "" Then
                SetNetworkConnectionString(ConnectionString, False)
            End If

            Dim result = LoadProcessInstanceByUnit(Stage, ProductNumber, SerialNumber, Sequence)
            If (result = ReturnCodes.UDBS_OP_SUCCESS) Then
                Sequence = Me.Sequence
            End If

            Return result
        End Function

        Public Overrides Function LoadProcessInstanceByUnit(Stage As String,
                                                  ProductNumber As String,
                                                  SerialNumber As String,
                                                  Sequence As Integer) As ReturnCodes
            Try
                ' Archive related upgrade
                Dim oriConnString As String, archiveConnString As String
                Dim sSQL As String, rsTmp As New DataTable
                Dim rc As ReturnCodes

                If mREADONLY OrElse ProcessInstanceRunning Then
                    ' Process object already loaded/running
                    LogError(New Exception("Process is already loaded."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Archive related upgrade
                oriConnString = GetNetworkConnectionString()
                archiveConnString = ""
                sSQL = "SELECT process_archive_state " &
                       "FROM product with(nolock) , unit with(nolock) , testdata_process with(nolock) , testdata_itemlistrevision  with(nolock) " &
                       "WHERE product_id=unit_product_id and unit_id=process_unit_id AND process_itemlistrev_id=itemlistrev_id " &
                       "AND product_number='" & ProductNumber & "' AND unit_serial_number='" & SerialNumber & "' " &
                       "AND itemlistrev_stage='" & Stage & "' AND process_sequence=" & Sequence &
                       " AND process_archive_state<>0"
                rc = CUtility.Utility_ExecuteSQLStatement(sSQL, rsTmp)
                If rc = ReturnCodes.UDBS_OP_SUCCESS Then
                    If (If(rsTmp?.Rows?.Count, 0)) > 0 Then
                        sSQL = "SELECT site_connection_string with(nolock) FROM udbs_site WHERE site_name='" &
                               oriConnString & "' AND site_built_code='ARCHIVE'"
                        rsTmp = Nothing
                        rc = CUtility.Utility_ExecuteSQLStatement(sSQL, rsTmp)
                        If rc = ReturnCodes.UDBS_OP_SUCCESS Then
                            If (If(rsTmp?.Rows?.Count, 0)) > 0 Then
                                'If rsTmp(0) Is Not Null Then
                                archiveConnString = KillNull(rsTmp(0)(0))
                                SetNetworkConnectionString(archiveConnString, False)
                                'End If
                            End If
                        Else
                            logger.Error("Failed to query archive system connection string.")
                            Return ReturnCodes.UDBS_ERROR
                        End If
                    Else
                        ' Not archived.
                        ' Leave the connection string as-is.
                    End If
                Else
                    logger.Error("Failed to query process archive state.")
                    Return ReturnCodes.UDBS_ERROR
                End If

                If MyBase.LoadProcessInstanceByUnit(Stage, ProductNumber, SerialNumber, Sequence) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Process Instance does not exist
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Load up items collection
                Dim loadRes = LoadItemsCollection()
                If loadRes <> ReturnCodes.UDBS_OP_SUCCESS Then
                    logger.Warn($"LoadItemsCollection : {loadRes}")
                End If

                ' Get exiting results
                loadRes = LoadExistingResults()
                If loadRes <> ReturnCodes.UDBS_OP_SUCCESS Then
                    logger.Warn($"LoadExistingResults : {loadRes}")
                End If

                mREADONLY = True

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        Public Function AddNote(Note As String) As ReturnCodes
            Return AddProcessInstanceNote(Note)
        End Function

        Public Overrides Function AddProcessInstanceNote(ProcessNote As String) As ReturnCodes
            If mREADONLY Then
                LogError(New Exception("Process is read only."))
                Return ReturnCodes.UDBS_ERROR
            End If
            Return MyBase.AddProcessInstanceNote(ProcessNote)
        End Function


        ''' <summary>
        ''' Store a datum into the process record, altering the process result is prohibited.
        ''' </summary>
        ''' <param name="ProcessItem"></param>
        ''' <param name="ItemData"></param>
        ''' <returns></returns>
        Public Function StoreProcessData(ProcessItem As String,
                                         ItemData As String) _
            As ReturnCodes
            Dim DesiredField As String

            If mREADONLY Then
                LogError(New Exception("Process is read only."))
                Return ReturnCodes.UDBS_ERROR
            End If

            If InStr(ProcessItem, "process_") > 0 And Len(ProcessItem) > 9 Then
                DesiredField = LCase(Trim(ProcessItem))
            Else
                DesiredField = "process_" & LCase(Trim(ProcessItem))
            End If
            ' lock out process_result field so that people can't manually alter
            ' the overall test result of a unit
            If DesiredField = "process_result" Then
                LogError(New Exception("Please use EvaluateDevice to determine process_result."))
                Return ReturnCodes.UDBS_ERROR
            End If

            ' Stores the supplied data to the process instance
            Return StoreProcessInstanceField(ProcessItem, ItemData)
        End Function

        ''' <summary>
        ''' https://docs.microsoft.com/en-us/dotnet/api/system.data.sqldbtype?view=netframework-4.6.2
        ''' </summary>
        ''' <param name="ResultName"></param>
        ''' <param name="ResultValue">double in .NET but SQL float in SQL Server</param>
        ''' <returns></returns>
        Public Overridable Function StoreValue(ResultName As String,
                                   ResultValue As Double) _
            As ResultCodes
            ' Stores the supplied value to the specifed result item
            Dim result = ResultCodes.UDBS_ERROR
            If mREADONLY Then
                LogError(New Exception("Process is read only."))
                Return result
            End If

            Try
                ' NB: Semantics WARNING!!!
                ' Clamp to valid SQL Server float type value
                ' NB: +/- Infinity values are converted to Single.Min and Single.Max respectively
                If Not Double.IsNaN(ResultValue) AndAlso (Nothing <> ResultValue) Then
                    ResultValue = Clamp(ResultValue, Single.MinValue, Single.MaxValue)
                End If
                ' NB: Nothing and NaN ResultValue values become [NULL] in UDBS Server
                result = Results(ResultName.ToLower).StoreValue(ResultValue)
            Catch ex As Exception
                LogErrorInDatabase(New UDBSException($"Error storing {ResultName} with value {ResultValue}", ex))
            End Try
            Return result
        End Function


        ''' <summary>
        ''' Stores the supplied passflag to the specified result item
        ''' </summary>
        Public Function StoreResultFlag(ResultName As String,
                                        ResultFlag As Integer) As ReturnCodes
            If mREADONLY Then
                LogError(New Exception("Process is read only."))
                Return ReturnCodes.UDBS_ERROR
            End If

            Try
                Return Results(ResultName.ToLower).StoreField("passflag", CStr(ResultFlag))
            Catch ex As Exception
                LogErrorInDatabase(New UDBSException($"Error storing {ResultName} result with value {ResultFlag}", ex))
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function


        Public Overridable Function StoreStringData(ResultName As String, StringData As String) As ReturnCodes
            ' Stores the supplied string data to the specified result item

            Try
                If mREADONLY Then
                    Throw New Exception("Process is read only.")
                End If

                Const MAX_STRING_LENGTH = 255
                If StringData?.Length > MAX_STRING_LENGTH Then
                    logger.Warn(
                        $"Cannot store full string data for UDBS Item '{ResultName}', since the length of {StringData.Length _
                                   } exceeds the maximum of {MAX_STRING_LENGTH}.  String will be truncated at {MAX_STRING_LENGTH _
                                   } characters.")
                    logger.Warn($"Given string for '{ResultName}' = '{StringData}'")
                    StringData = StringData.Substring(0, MAX_STRING_LENGTH)
                End If

                Return Results(ResultName.ToLower).StoreField("StringData", StringData)

            Catch ex As Exception
                LogErrorInDatabase(New UDBSException($"Error storing string data {StringData} for {ResultName}", ex))
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Clear the test results under a item group.
        ''' </summary>
        ''' <param name="groupName">The group to clear.</param>
        ''' <returns>The outcome of the operation.</returns>
        Public Function ClearResults(groupName As String) As ReturnCodes
            If mREADONLY Then
                LogError(New Exception("Process is read only."))
                Return ReturnCodes.UDBS_ERROR
            End If

            If String.IsNullOrEmpty(groupName) Then
                LogError(New Exception("Group name must be specified."))
                Return ReturnCodes.UDBS_ERROR
            End If

            Dim items As String() = Nothing
            If GetItemsInGroup(groupName, items) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_ERROR
            End If

            For Each item In items
                Dim aResult = Results(item)
                If aResult.Clear() <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If
            Next

            ' Also clear the group item itself.
            If Results(groupName).Clear() <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_ERROR
            End If

            Return ReturnCodes.UDBS_OP_SUCCESS
        End Function

        ''' <summary>
        ''' Get item names belonging to a given group name (report level).
        ''' 
        ''' </summary>
        ''' <param name="groupName">Group name.</param>
        ''' <param name="items">(Out) Reference to the array of items.</param>
        ''' <returns>The outcome of the operation.</returns>
        Public Function GetItemsInGroup(groupName As String, ByRef items() As String) As ReturnCodes
            If String.IsNullOrEmpty(groupName) Then
                ' Special case: searching for top-level items.
                Return GetTopLevelItems(items)
            End If

            groupName = groupName.ToLowerInvariant()

            Try
                If Not Results.ContainsKey(groupName) Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                If Results(groupName).IsGroup Then
                    Dim groupItemNumber = Results(groupName).ItemNumber
                    Dim groupReportLevel = Results(groupName).ReportLevel
                    Dim numItems = 0

                    Dim candidates = Results.OrderBy(Function(kv) kv.Value.ItemNumber).
                        Where(Function(r) r.Value.ItemNumber > groupItemNumber).
                        Select(Function(r) r.Value)

                    For Each candidate As CTestData_Result In candidates
                        If candidate.ReportLevel > groupReportLevel Then
                            If candidate.IsGroup = False Then
                                numItems = numItems + 1
                                ReDim Preserve items(numItems - 1)
                                items(numItems - 1) = candidate.ItemName
                            End If
                        Else
                            'we've hit the next group
                            Exit For
                        End If
                    Next

                    Return ReturnCodes.UDBS_OP_SUCCESS
                Else
                    'Not a group.
                    logger.Debug($"Item {groupName} is not a group.")
                    Return ReturnCodes.UDBS_ERROR
                End If
            Catch ex As Exception
                logger.Error($"Unexpected error trying to get items in group ""{groupName}"": {ex.Message}")
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Get the top-level item names.
        ''' </summary>
        ''' <param name="items">(Out) The result of the function.</param>
        ''' <returns>Whether or not the operation succeeded (i.e. Did it find any top-level item?).</returns>
        Private Function GetTopLevelItems(ByRef items() As String) As ReturnCodes
            items = Results.OrderBy(Function(kv) kv.Value.ItemNumber).
                Where(Function(r) r.Value.ReportLevel = 1).
                Select(Function(r) r.Value.ItemName).ToArray
            If items.Count > 0 Then
                Return ReturnCodes.UDBS_OP_SUCCESS
            Else
                Return ReturnCodes.UDBS_ERROR
            End If
        End Function

        ' Candidate for removal.
        Private Function CheckValue(ResultName As String,
                                   ResultValue As Double) _
            As ResultCodes
            ' Compares the value provided with the specs for the item
            CheckValue = Results(ResultName.ToLower).CheckValue(ResultValue)
        End Function

        Private Function CDblNull(ByVal val As Object) As Double
            Try
                Return CDbl(val)
            Catch ex As Exception
                Return 0.0D
            End Try
        End Function

        Private Function CStrNull(ByVal val As Object) As String
            Try
                Return CStr(val)
            Catch ex As Exception
                Return ""
            End Try
        End Function

        '**********************************************************************
        '* Support Functions
        '**********************************************************************

        Private Function LoadItemsCollection() As ReturnCodes
            ' Function creates the items collection for this process instance object
            Dim ItemNumber As Integer
            Dim ItemName As String
            Dim Descriptor As String
            Dim Description As String
            Dim ReportLevel As Integer
            Dim Units As String
            Dim CriticalSpec As Integer
            Dim WarnMin As Double
            Dim WarnMax As Double
            Dim FailMin As Double
            Dim FailMax As Double
            Dim SanityMin As Double
            Dim SanityMax As Double
            Dim BlobExists As Integer

            Try
                ' Load an itemlist object
                Dim IL As New CItemlist()
                Dim result = IL.LoadItemList(_Process, ProductNumber, ProductRelease, Stage, Revision)
                If result <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return result
                End If

                mResults.Clear()
                For Each dr As DataRow In IL.Items_RS.Rows

                    ItemNumber = KillNullInteger(dr("itemlistdef_itemnumber"))
                    ItemName = KillNull(dr("itemlistdef_itemname"))
                    Descriptor = KillNull(dr("itemlistdef_descriptor"))
                    Description = KillNull(dr("itemlistdef_description"))
                    ReportLevel = KillNullInteger(dr("itemlistdef_report_level"))
                    Units = KillNull(dr("itemlistdef_units"))
                    CriticalSpec = KillNullInteger(dr("itemlistdef_critical_spec"))
                    WarnMin = KillNullDouble(dr("itemlistdef_warning_min"))
                    WarnMax = KillNullDouble(dr("itemlistdef_warning_max"))
                    FailMin = KillNullDouble(dr("itemlistdef_fail_min"))
                    FailMax = KillNullDouble(dr("itemlistdef_fail_max"))
                    SanityMin = KillNullDouble(dr("itemlistdef_sanity_min"))
                    SanityMax = KillNullDouble(dr("itemlistdef_sanity_max"))

                    If Not IsDBNull(dr("itemlistdef_blobdata_exists")) Then
                        BlobExists = KillNullInteger(dr("itemlistdef_blobdata_exists"))
                    End If

                    mResults.Add(ItemName,
                             New CTestData_Result(Me, ItemNumber, ItemName, Descriptor, Description, ReportLevel, Units,
                                                  CriticalSpec,
                                                  WarnMin, WarnMax, FailMin, FailMax, SanityMin, SanityMax, BlobExists))

                Next

                ' Process the itemlist groups
                SetGroupFlags()
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        Private Sub SetGroupFlags()
            ' Function walks through the results collection and sets the group flag for items that are groups
            For Each i In mResults
                i.Value.IsGroup = ItemIsGroup(i.Key)
            Next i
        End Sub

        'this function is duplciated from the Itemlist Object
        Private Function ItemIsGroup(ItemName As String) _
            As Boolean
            Try

                'This function assumes that the item name actually exists!
                Dim NextItemNumber As Integer
                Dim HasImmediateChildren As Boolean
                Dim NoSpecs As Boolean
                Dim retval As Boolean

                'Move the recordset pointer to the item name we're concerned about...
                Dim tmpItem As CTestData_Item = mResults(ItemName.ToLower)
                Dim NextItem As CTestData_Item = Nothing

                'Check to see if there is any specs on this item!
                If tmpItem.HasSpecs = False Then
                    NoSpecs = True
                Else
                    NoSpecs = False
                End If

                ' Get information on next item
                NextItemNumber = tmpItem.ItemNumber + 1

                If NextItemNumber > mResults.Count Then
                    ' This cannot be a group
                    retval = False
                Else
                    NextItem = mResults(GetItemName(NextItemNumber))

                    If NextItem.ReportLevel = tmpItem.ReportLevel + 1 Then
                        HasImmediateChildren = True
                    End If

                    If HasImmediateChildren And NoSpecs Then
                        retval = True
                    Else
                        retval = False
                    End If

                End If

                Return retval
            Catch ex As Exception
                Throw New UDBSException($"Failed to determine if item {ItemName} is a group.", ex)
            End Try
        End Function

        ' This function is duplciated from the Itemlist Object
        ' Candidate for removal.
        Private Function GetItemName(ItemNumber As Integer) _
            As String
            ' Function returns the item name of the specified item number
            Try
                Return mResults.First(Function(f) f.Value.ItemNumber = ItemNumber).Key
            Catch ex As Exception
                Throw New UDBSException("Item number " & ItemNumber & " does not exist in current itemlist.")
            End Try
        End Function

        ''' <summary>
        ''' Check all children of a group and return overall pass/fail status.
        ''' Dive into child groups if not trusted.
        ''' </summary>
        ''' <param name="GroupName"></param>
        ''' <returns></returns>
        Public Function EvaluateGroup(GroupName As String) As ResultCodes

            Try
                ' Set if group results are to be trusted
                'TrustGroupResults = True
                Dim TrustGroupResults As Boolean = False       ' always re-eval group, BC
                Dim GroupResult As Integer = Integer.MaxValue

                ' Work LOCALLY for all evaluations, leave network updates to EndInstance or user..

                ' If device is being checked then do not trust child group evaluations
                If GroupName = "~DEVICE" Then
                    LogError(New Exception("Please use EvaluateDevice function for device evals."))
                    Return ResultCodes.UDBS_ERROR
                End If

                ' Lets make sure that the indicated item is actually a group!
                If ItemIsGroup(GroupName) = False Then
                    LogError(New Exception($"The group selected '{GroupName}' is not a valid group."))
                    Return ResultCodes.UDBS_ERROR
                End If

                ' Begin scanning group, Scan will also store group result
                GroupResult = ScanGroupResults(GroupName, TrustGroupResults)

                Return CType(GroupResult, ResultCodes)

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ResultCodes.UDBS_ERROR
            End Try
        End Function

        Private Function ScanGroupResults(GroupName As String,
                                          TrustGroupResults As Boolean) _
            As Integer
            ' Function scans the resultlist and returns the worst result for the given group

            Dim TempResult As Integer
            Dim GroupResult As Integer
            Dim ItemCount As Integer
            Dim NumItems As Integer
            Dim GroupReportLevel As Integer
            Dim GroupItemNumber As Integer
            Dim ListIncompleteFlag As Boolean
            Dim CurrentItemNumber As Integer
            Dim CurrentItemLevel As Integer     'This will refer to the item in the list definition
            Dim CurrentItemName As String     'This will refer to the item in the list definition
            Dim CurrentItemResultFlag As Integer _
            'This will refer to the result flag actually stored for the unit in the result table

            GroupResult = Integer.MaxValue

            'Move to the item (which is supposed to be a group) that we want to evaluate...
            If GroupName = "~DEVICE" Then
                ' Get the report level and move to the next item
                GroupReportLevel = 0
                GroupItemNumber = 0
            Else
                ' Get the report level and move to the next item
                GroupReportLevel = mResults(GroupName).ReportLevel
                GroupItemNumber = mResults(GroupName).ItemNumber
            End If

            CurrentItemNumber = GroupItemNumber + 1
            NumItems = mResults.Count

            ItemCount = 1
            'Now lets scan through the following items until we get to the next group
            Do While CurrentItemNumber <= NumItems ' was < only
                ' Retrieve level and name for next item
                CurrentItemName = GetItemName(CurrentItemNumber)
                CurrentItemLevel = mResults(CurrentItemName).ReportLevel

                ' Are we at the end of the group?
                If CurrentItemLevel <= GroupReportLevel Then
                    If ItemCount = 1 Then
                        'The first item in the group is at the same level as the group!??!?!  ERROR!
                        ScanGroupResults = ResultCodes.UDBS_ERROR
                        Exit Function
                    Else
                        'we've finished all the items in the group... exit time
                        Exit Do
                    End If
                End If

                'Do a quick (yet majorly incomplete) list integrity check...
                If (CurrentItemLevel - GroupReportLevel) > ItemCount Then
                    'Somehow the itemlist is screwed!  We got to this point in the code assuming that this item
                    'was a group.  Either the the first item does not have level +1 as the level, or somehow the
                    'list has managed to have a report level increment of greater than one.
                    LogError(New Exception("Item List is bad. The group levels are incorrect."))
                    ScanGroupResults = ResultCodes.UDBS_ERROR
                    Exit Function
                End If

                'Check to see if the item we're looking at is a group, first...
                If mResults(CurrentItemName).IsGroup Then
                    '***** TrustGroupResults has been set for always False *****
                    '***** the following loop must be re-run *****
                    '***** sub-items may be updated since last evalGroup *****
                    'If (TrustGroupResults = False) Or ((Result(CurrentItemName).PassFlagStored) = False) Then
                    ' Evaluate the sub group
                    TempResult = ScanGroupResults(CurrentItemName, TrustGroupResults)
                    ' Catch List Incomplete Flag, but ignore if this is the first evaluation
                    StripIncompleteFlag(TempResult, ListIncompleteFlag)
                    ' Apply group result
                    GroupResult = Math.Min(GroupResult, TempResult)
                    'Result(CurrentItemName).PassFlag = GroupResult
                    'Result(CurrentItemName).PassFlagStored = True
                    'Else
                    'Do nothing special with this group... we'll just take the result flag as if it was a normal item
                    'End If
                End If

                If Not (mResults(CurrentItemName).PassFlagStored) Then
                    ' Check if this is a 'placeholder' entry
                    If mResults(CurrentItemName).CriticalSpec = 1 Then
                        ' There should have been a value here...
                        GroupResult = Math.Min(GroupResult, CInt(ResultCodes.UDBS_SPECS_FAIL))
                        ListIncompleteFlag = True
                    Else
                        ' Do not allow absense of this item to effect the outcome of an eval. This should be hit
                        ' by 'placeholder' items that are storing links like blob_id or string data
                        GroupResult = Math.Min(GroupResult, CInt(ResultCodes.UDBS_SPECS_PASS))
                    End If
                Else
                    'The item was at least stored to the database...
                    ' Get result value, may include an incomplete flag
                    CurrentItemResultFlag = mResults(CurrentItemName).PassFlag

                    ' Catch List Incomplete Flag
                    StripIncompleteFlag(CurrentItemResultFlag, ListIncompleteFlag)

                    ' Collapse individual results to result families
                    If _
                        CurrentItemResultFlag <= ResultCodes.UDBS_SPECS_SANITY And
                        CurrentItemResultFlag > ResultCodes.UDBS_SPECS_SANITY - 10 Then
                        ' Sanity series...
                        CurrentItemResultFlag = ResultCodes.UDBS_SPECS_SANITY
                    ElseIf _
                        CurrentItemResultFlag <= ResultCodes.UDBS_SPECS_FAIL And
                        CurrentItemResultFlag > ResultCodes.UDBS_SPECS_FAIL - 10 Then
                        ' Failure series...
                        CurrentItemResultFlag = ResultCodes.UDBS_SPECS_FAIL
                    ElseIf _
                        CurrentItemResultFlag <= ResultCodes.UDBS_SPECS_WARNING And
                        CurrentItemResultFlag > ResultCodes.UDBS_SPECS_WARNING - 10 Then
                        ' Warning series...
                        CurrentItemResultFlag = ResultCodes.UDBS_SPECS_WARNING
                    ElseIf _
                        CurrentItemResultFlag >= ResultCodes.UDBS_SPECS_PASS And
                        CurrentItemResultFlag < ResultCodes.UDBS_SPECS_PASS + 10 Then
                        ' Pass series...
                        CurrentItemResultFlag = ResultCodes.UDBS_SPECS_PASS
                    End If
                    GroupResult = Math.Min(GroupResult, CurrentItemResultFlag)
                End If
                'Debug.Print CurrentItemNumber & ": " & CurrentItemName & ": " & CurrentItemLevel & ": " & CurrentItemResultFlag
                ItemCount = ItemCount + 1
                CurrentItemNumber = CurrentItemNumber + 1
            Loop
            'Debug.Print GroupName & ": " & GroupResult
            ' Check for no spec results
            If GroupResult = Integer.MaxValue Then
                ' No values were checked against spec
                GroupResult = ResultCodes.UDBS_ERROR
            End If

            ' Apply incomplete flag
            If ListIncompleteFlag = True Then
                GroupResult = CInt(GroupResult + INCOMPLETE * (GroupResult / Math.Abs(GroupResult)))
            End If

            ' An incomplete pass is actually a FAIL
            If GroupResult >= INCOMPLETE Then
                GroupResult = ResultCodes.UDBS_SPECS_FAIL - INCOMPLETE
            End If

            ' Store group result only if it is not read only
            If (GroupName <> "~DEVICE") And (Not mREADONLY) Then
                If mResults(GroupName).StoreField("passflag", CStr(GroupResult)) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    GroupResult = ResultCodes.UDBS_ERROR
                Else
                    mResults(GroupName).PassFlagStored = True
                End If
            End If
            'Debug.Print GroupName & ": " & GroupResult
            ScanGroupResults = GroupResult
        End Function

        Private Sub StripIncompleteFlag(ByRef ResultValue As Integer,
                                        ByRef ListIncompleteFlag As Boolean)
            ' Function checks for incomplete flag in result value and removes it while setting flag
            If ResultValue = Integer.MaxValue Or ResultValue = ResultCodes.UDBS_ERROR Then
                Exit Sub
            End If

            If ResultValue > INCOMPLETE Then
                ResultValue = ResultValue - INCOMPLETE
                ListIncompleteFlag = True
            ElseIf ResultValue < -INCOMPLETE Then
                ResultValue = ResultValue + INCOMPLETE
                ListIncompleteFlag = True
            End If
        End Sub

        Public Overridable Function EvaluateDevice() As ResultCodes
            ' Function description

            ' Evaluate device for overall pass/fail
            ' This is the 'brute force' catch all that goes through EVERY item in the itemlist
            ' looking for the minimum value on the passflag. The function will return the
            ' Minimum value, and append the INCOMPLETE flag if necessary

            Dim TrustGroupResults As Boolean
            Dim GroupResult As Integer

            Try
                ' Set if group results are to be trusted
                TrustGroupResults = False
                GroupResult = Integer.MaxValue

                ' Work LOCALLY for all evaluations, leave network updates to EndInstance or user..

                ' Begin scanning group
                GroupResult = ScanGroupResults("~DEVICE", TrustGroupResults)

                ' Store device result
                If StoreProcessInstanceField("result", CStr(GroupResult)) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ResultCodes.UDBS_ERROR
                End If

                Return CType(GroupResult, ResultCodes)

                Exit Function
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ResultCodes.UDBS_ERROR
            End Try

        End Function

        Private Function LoadExistingResults() As ReturnCodes
            ' Function retrieves existing process and result data and loads it into object
            ' added PassFlagStored by BC Dec 4 2001
            Dim rsTemp As New DataTable
            ' The local PI object has been created and is in memory

            rsTemp = Results_RS

            For Each dr As DataRow In rsTemp.Rows
                Dim namee = KillNull(dr("itemlistdef_itemname"))
                If Not (IsDBNull(dr("result_passflag"))) Then
                    Results(KillNull(dr("itemlistdef_itemname"))).PassFlag = KillNullInteger(dr("result_passflag"))
                    Results(KillNull(dr("itemlistdef_itemname"))).PassFlagStored = True
                End If
                If Not (IsDBNull(dr("result_stringdata"))) Then
                    'Results(dr("itemlistdef_itemname")).StoreField "result_stringdata", KillNull(dr("result_stringdata"))
                    ' problem with "Friend" property during initiating (.StringData)???
                    ' .StringData - method or data not found
                    Results(KillNull(dr("itemlistdef_itemname"))).StringData = KillNull(dr("result_stringdata"))
                End If
                If Not (IsDBNull(dr("result_value"))) Then
                    Results(KillNull(dr("itemlistdef_itemname"))).Value = KillNullDouble(dr("result_value"))
                    Results(KillNull(dr("itemlistdef_itemname"))).ValueStored = True
                End If
                If Not (IsDBNull(dr("result_blobdata_exists"))) Then
                    Results(KillNull(dr("itemlistdef_itemname"))).ResultBlobDataExists =
                        CBool(dr("result_blobdata_exists"))
                End If

            Next
            Return ReturnCodes.UDBS_OP_SUCCESS
        End Function

        ''' <summary>
        ''' Validates if a test instance for the ProductID.SerialNumber.TestStage.StationID is alreday available for restart.
        ''' Starts by checking for active processes in the local DB. If a process Is found the network DB Is checked for a matching test instance.
        ''' This function will also check for local DB data integrity and discard it when data has been corrupted.
        ''' Furthermore, when the local DB is corrupted and the test instance is found to be IN PROCESS on the network DB,
        ''' then the test is TERMINATED.
        ''' </summary>
        ''' <param name="processName">process name.</param>
        ''' <param name="stage">process stage</param>
        ''' <param name="productID">udbs product ID</param>
        ''' <param name="serialNumber">unit serial number</param>
        ''' <param name="testSequence">test instance sequence number</param>
        ''' <param name="testIDLabel">Unique string label identifying the test instance.</param>
        ''' <param name="archiveFolder">Archive folder path. Used when the local DB data is corrupted.
        ''' It will be compared to the station name found on the network when recovering a test instabce.</param>
        ''' <param name="stationID">stationID that was provided at initialization of a test.</param>
        ''' <returns>True when the test instance can be restarted. False otherwise.</returns>
        Friend Shared Function PrepareTestForRestart(processName As String, stage As String, productID As String,
                                                         serialNumber As String, testSequence As Integer, testIDLabel As String,
                                                         archiveFolder As String, stationID As String) As Boolean

            Try
                If DatabaseSupport.UDBSDebugMode Then
                    logger.Debug($"Looking for UDBS Test instance for SN: {serialNumber}, ID: {productID}, Stage:{stage}, Sequence:{testSequence}.")
                End If

                'Check to see if there is already an active test process in the local DB
                Dim localDBCorrupted As Boolean
                Dim activeProcessID As Integer = -1

                If CProcessInstance.CheckActiveProcesses(processName, stage, productID, serialNumber, localDBCorrupted, activeProcessID) <> MasterInterface.ReturnCodes.UDBS_OP_SUCCESS Then
                    logger.Warn($"Failed to check the active processes in the local DB.")
                    Return False
                End If

                If activeProcessID <> -1 Then
                    'An active process was found and is available for restart.
                    logger.Info($"Found an active test process with process ID: {activeProcessID} with matching serial number. The test instance is available for restart.")
                    Return True
                Else
                    If localDBCorrupted Then
                        ' backup the local database to the recovery folder and delete the local database
                        LocalDBIntegrityChecker.DiscardLocalDB(archiveFolder, testIDLabel)
                    End If
                    logger.Info("No active process found in the local DB. Checking network DB...")

                    Dim testInstanceInfo = TestDataProcessInfo.GetProcessInfo(productID, serialNumber, stage, testSequence)

                    logger.Info($"Found a UDBS testdata instance for SN: {serialNumber}, Sequence {testInstanceInfo.Sequence}, started on: {testInstanceInfo.StartDate}, with status: {testInstanceInfo.Status}.")

                    Select Case testInstanceInfo.Status
                        Case UdbsProcessStatus.IN_PROCESS

                            'Make sure that it is not in process on a different station
                            If Not testInstanceInfo.IsProcessOwnedByThisStation(stationID) Then
                                logger.Warn($"The test instance with: SN " & serialNumber &
                                                ", ID " & productID & " is not avalaible for restart. It is already executing on station " & testInstanceInfo.Station &
                                                " with status = " & testInstanceInfo.Status & ".")
                                Return False
                            End If

                            If localDBCorrupted Then
                                'TODO: TMTD-537-Create function to Terminate test instance without using LoadExisting()
                                Using testInstance = New CTestdata_Instance
                                    logger.Info("Terminating the test instance as the local DB data is corrupted!")

                                    If (testInstance.LoadExisting(stage, productID, serialNumber, testSequence) <> ReturnCodes.UDBS_OP_SUCCESS) Then
                                        logger.Error("Couldn't load the test instance in order to terminate it.")
                                        Return False
                                    End If

                                    testInstance.TerminateWithoutSynchronizing()
                                    Return False
                                End Using

                            End If

                            logger.Info("The test sequence is available for restart.")
                            Return True

                        Case UdbsProcessStatus.PAUSED
                            logger.Info("The test sequence is available for restart.")
                            Return True

                        Case UdbsProcessStatus.COMPLETED, UdbsProcessStatus.STARTING, UdbsProcessStatus.TERMINATED
                            'If status is STARTING it means the test failed to start. Nothing to recover.
                            logger.Info("The test sequence is not available for restart.")
                            Return False
                    End Select

                    logger.Info("Couldn't load the test instance from the network DB. No test is available for restart.")
                    Return False
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return False
            End Try

        End Function

        Protected Overrides Sub Dispose(disposing As Boolean)
            MyBase.Dispose(disposing)
            mResults = Nothing
        End Sub

        Public Sub New()
            MyBase.New(InterfaceSupport.PROCESS)
        End Sub
    End Class
End Namespace
