Option Explicit On
Option Compare Text
Option Infer On
Option Strict On

Imports System.ComponentModel
Imports System.IO

Namespace MasterInterface
    Public Enum UdbsProcessStatus
        STARTING
        IN_PROCESS
        PAUSED
        TERMINATED
        COMPLETED
        UNKNOWN
    End Enum

    Public Class CProcessInstance
        Implements IDisposable

        Protected _Process As String

        ' Table Identification
        Private Shared ReadOnly mProductTable As String = "product"
        Private Shared ReadOnly mUnitTable As String = "unit"
        Private Shared ReadOnly mProcessRegistrationTable As String = "process_registration"

        Private mProcessTable As String
        Private mResultTable As String
        Private mItemListRevisionTable As String
        Private mItemListDefinitionTable As String
        Private mBlobTable As String
        Private mProcessAttributesTable As String
        Private mProcessAttributesHistoryTable As String

        ' Timing Variables
        Private mStartTime As Date
        Private mStopTime As Date
        Private mDurationS As Double
        Private mStartTicks As Long

        ' Object State Information
        Protected mREADONLY As Boolean
        Private mProcessRunning As Boolean = False

        ' ItemList Information
        Private ReadOnly mITEMLIST As New CItemlist

        ' Product Information
        Private ReadOnly mPRODUCT As New CProduct

        ' Unit Information
        Private mUnitID As Long
        Private mSerialNumber As String
        Private mUnitCreatedDate As Date
        Private mUnitEmployeeNumber As String
        Private mUnitNumLabels As Integer
        Private mUnitReport As String
        Private mUnitOraclePN As String
        Private mUnitCataloguePN As String
        Private mUnitVariance As String

        ' Process and Result Information
        Protected mProcessID As Integer = 0
        Private mProcessInfo As New DataTable
        Private mResultInfo As New DataTable

        Protected Friend _filesAttachedToUDBS As New List(Of String)

        Public Class DBEnumType
            Public Property ID As Integer
            Public Property Name As String
            Public Property Immutable As Boolean
            Public Overridable Property DBEnums As IEnumerable(Of DBEnum)
        End Class

        Public Class DBEnum
            Public Property ID As Integer
            Public Property Name As String

        End Class

        Private mEnumLookup As New List(Of DBEnumType)

        '**********************************************************************
        '* Properties
        '**********************************************************************

        ''' <summary>
        ''' The process' context. A human-readable string to be used in log messages
        ''' for debugging.
        ''' </summary>
        Friend ReadOnly Property ProcessContext As String
            Get
                Return $"Type={Process} Name={Stage} ID={ID} Product={ProductNumber} SN={UnitSerialNumber}"
            End Get
        End Property

        ' Process Information
        Public ReadOnly Property Process As String
            Get
                Return _Process
            End Get
        End Property

        Public ReadOnly Property ID As Integer
            Get
                Return mProcessID
            End Get
        End Property

        Public ReadOnly Property UnitID As Integer
            Get
                Return KillNullInteger(mProcessInfo(0)("process_unit_id"))
            End Get
        End Property

        Public ReadOnly Property Sequence As Integer
            Get
                Return KillNullInteger(mProcessInfo(0)?("process_sequence"))
            End Get
        End Property

        ' TODO: Start with an upper case letter.
        ' TODO: Convert to ResultCode enumerated type.
        Public ReadOnly Property result As Integer
            Get
                Return KillNullInteger(mProcessInfo(0)?("process_result"))
            End Get
        End Property

        ' TODO: Rename to OverwrittenResult
        ' TODO: Convert to ResultCode enumerated type.
        ' TODO: If not present, return NONE, not '0' (NONE != '0')
        Public ReadOnly Property overwritten_result As Integer
            Get
                If Not mProcessInfo.Columns.Contains("process_overwritten_result") _
                        OrElse IsDBNull(mProcessInfo(0)?("process_overwritten_result")) Then
                    Return 0
                Else
                    Return KillNullInteger(mProcessInfo(0)?("process_overwritten_result"))
                End If
            End Get
        End Property

        Public ReadOnly Property overwritten_date As Date
            Get
                If Not mProcessInfo.Columns.Contains("process_overwritten_date") Then
                    Return Date.MinValue
                Else
                    Return KillNullDate(mProcessInfo(0)?("process_overwritten_date"))
                End If
            End Get
        End Property

        Public ReadOnly Property overwritten_by As String
            Get
                If (Not mProcessInfo.Columns.Contains("process_overwritten_empnum")) Then
                    Return String.Empty
                Else
                    Return KillNull(mProcessInfo(0)?("process_overwritten_empnum"))
                End If
            End Get
        End Property

        Public ReadOnly Property StartDate As Date
            Get
                Return KillNullDate(mProcessInfo(0)?("process_start_date"))
            End Get
        End Property

        Public ReadOnly Property StopDate As Date
            Get
                Return KillNullDate(mProcessInfo(0)?("process_end_date"))
            End Get
        End Property

        Public ReadOnly Property ActiveDuration As Double
            Get
                If Not IsDBNull(mProcessInfo(0)?("process_active_duration")) Then
                    Return CDbl(mProcessInfo(0)?("process_active_duration"))
                Else
                    Return 0
                End If
            End Get
        End Property

        Public ReadOnly Property TotalDuration As Double
            Get
                If Not IsDBNull(mProcessInfo(0)?("process_total_duration")) Then
                    Return CDbl(mProcessInfo(0)?("process_total_duration"))
                Else
                    Return 0
                End If
            End Get
        End Property

        Public ReadOnly Property Status As String
            Get
                Return KillNull(mProcessInfo(0)?("process_status"))
            End Get
        End Property

        Public ReadOnly Property Notes As String
            Get
                Return KillNull(mProcessInfo(0)?("process_notes"))
            End Get
        End Property

        Public ReadOnly Property EmployeeNumber As String
            Get
                Return KillNull(mProcessInfo(0)?("process_employee_number"))
            End Get
        End Property

        Public ReadOnly Property StationName As String
            Get
                Return KillNull(mProcessInfo(0)?("process_station"))
            End Get
        End Property

        Public ReadOnly Property SoftwareVersion As String
            Get
                Return KillNull(mProcessInfo(0)("process_sw_version"))
            End Get
        End Property

        Friend ReadOnly Property Instance_RS As DataTable
            Get
                Return mProcessInfo.Copy()
            End Get
        End Property

        Friend ReadOnly Property Results_RS As DataTable
            Get
                Return mResultInfo.Copy()
            End Get
        End Property

        ' Not used.
        ' The property itself is not used.
        ' The member variable it encapsulate is still accessed, but does not seem to be
        ' assigned values.
        ' It might be an artifact from an older design.
        Public ReadOnly Property EnumLookup As IEnumerable(Of DBEnumType)
            Get
                Return mEnumLookup
            End Get
        End Property


        Public ReadOnly Property ProcessInstanceRunning As Boolean
            Get
                Return mProcessRunning
            End Get
        End Property

        ' Unit Information
        Public ReadOnly Property ProductNumber As String
            Get
                Return mPRODUCT.Number
            End Get
        End Property

        Public ReadOnly Property ProductRelease As Integer
            Get
                Return mPRODUCT.Release
            End Get
        End Property

        Public ReadOnly Property ProductDescriptor As String
            Get
                Return mPRODUCT.Descriptor
            End Get
        End Property

        Public ReadOnly Property ProductDescription As String
            Get
                Return mPRODUCT.Description
            End Get
        End Property

        Public ReadOnly Property ProductCreatedBy As String
            Get
                Return mPRODUCT.CreatedBy
            End Get
        End Property

        Public ReadOnly Property ProductCreatedDate As Date
            Get
                Return mPRODUCT.CreatedDate
            End Get
        End Property

        Public ReadOnly Property ProductReleaseReason As String
            Get
                Return mPRODUCT.ReleaseReason
            End Get
        End Property

        Public ReadOnly Property ProductSNProdCode As String
            Get
                Return mPRODUCT.SNProdCode
            End Get
        End Property

        Public ReadOnly Property ProductSNTemplate As String
            Get
                Return mPRODUCT.SNTemplate
            End Get
        End Property

        Public ReadOnly Property ProductSNLastUnit As Integer
            Get
                Return mPRODUCT.SNLastUnit
            End Get
        End Property

        Public ReadOnly Property ProductFamily As String
            Get
                Return mPRODUCT.Family
            End Get
        End Property

        Public ReadOnly Property UnitSerialNumber As String
            Get
                Return mSerialNumber
            End Get
        End Property

        Public ReadOnly Property UnitCreatedDate As Date
            Get
                Return mUnitCreatedDate
            End Get
        End Property

        Public ReadOnly Property UnitCreatedBy As String
            Get
                Return mUnitEmployeeNumber
            End Get
        End Property

        Public ReadOnly Property UnitNumLabels As Integer
            Get
                Return mUnitNumLabels
            End Get
        End Property

        Public ReadOnly Property UnitReport As String
            Get
                Return mUnitReport
            End Get
        End Property


        Public ReadOnly Property UnitOraclePN As String
            Get
                Return mUnitOraclePN
            End Get
        End Property

        Public ReadOnly Property UnitCataloguePN As String
            Get
                Return mUnitCataloguePN
            End Get
        End Property

        Public ReadOnly Property UnitVariance As String
            Get
                Return mUnitVariance
            End Get
        End Property

        ' Itemlist Information
        Public ReadOnly Property ItemListRevID As Integer
            Get
                Return mITEMLIST.ItemListRevID
            End Get
        End Property

        Public ReadOnly Property Stage As String
            Get
                Return mITEMLIST.Stage
            End Get
        End Property

        Public ReadOnly Property Revision As Integer
            Get
                Return mITEMLIST.Revision
            End Get
        End Property

        Public ReadOnly Property RevisionDescription As String
            Get
                Return mITEMLIST.RevisionDescription
            End Get
        End Property

        Public ReadOnly Property ItemListCreatedDate As Date
            Get
                Return mITEMLIST.CreatedDate
            End Get
        End Property

        Public ReadOnly Property ItemListCreatedBy As String
            Get
                Return mITEMLIST.EmployeeNumber
            End Get
        End Property

        ''' <summary>
        ''' Whether or not this process is read-only.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IsReadOnly As Boolean
            Get
                Return mREADONLY
            End Get
        End Property

        ''' <summary>
        ''' Gets the list of all the files attached to the database.
        ''' </summary>
        Public ReadOnly Property FilesAttachedtoUDBS As List(Of String)
            Get
                Return _filesAttachedToUDBS
            End Get
        End Property

        '**********************************************************************
        '* Methods
        '**********************************************************************

        ''' <summary>
        ''' Retrieve from the local store, the most recent process attribute of this instance
        ''' given a valid key
        ''' </summary>
        ''' <typeparam name="TValue">normally a numeric result</typeparam>
        ''' <typeparam name="TStringData">normally a string result</typeparam>
        ''' <param name="attributeName">Must by a valid enum name. <see cref="DBEnumType"/> </param>
        ''' <param name="value"></param>
        ''' <param name="stringData"></param>
        ''' <remarks>
        ''' Process instance attributes seems to be a prototype that was never completed.
        ''' See TMTD-201.
        ''' </remarks>
        Public Sub GetProcessInstanceAttribute(Of TValue, TStringData)(attributeName As String, ByRef value As TValue, ByRef stringData As TStringData)
            If Not mEnumLookup?.Any() Then
                Throw New UDBSException("Process instance not yet started")
            End If
            ' Is it a valid key?
            Dim enumType As DBEnumType = mEnumLookup.SingleOrDefault(Function(et) et.Name.Equals(attributeName, StringComparison.OrdinalIgnoreCase))
            If enumType Is Nothing Then
                Throw New UDBSException($"Undefined {_Process} process attribute key: {attributeName}. Check database definitions of valid enum/key types")
            End If
            ' Go Grab it

            Dim dataRS As New DataTable()
            Dim sqlQuery = $"select * from {mProcessAttributesTable}
                            where attribute_process_id={mProcessID} and attribute_attributetype_id={enumType.ID}"
            If OpenLocalRecordSet(dataRS, sqlQuery) = ReturnCodes.UDBS_OP_SUCCESS AndAlso dataRS?.AsEnumerable().Any() Then
                Dim resultObj = dataRS.AsEnumerable().First()("attribute_value")
                Dim typ As Type = GetType(TValue)
                Dim conv As TypeConverter = TypeDescriptor.GetConverter(typ)
                Dim typedResult = conv.ConvertTo(resultObj, typ)
                value = CType(typedResult, TValue)

                resultObj = dataRS.AsEnumerable().First()("attribute_stringdata")
                typ = GetType(TStringData)
                conv = TypeDescriptor.GetConverter(typ)
                typedResult = conv.ConvertTo(resultObj, typ)
                stringData = CType(typedResult, TStringData)
            Else
                'Not there yet?
                logger.Warn($"No process instance attribute found for {attributeName}. Returning nulls")
                value = Nothing
                stringData = Nothing
            End If

        End Sub

        ''' <summary>
        ''' Store a process attribute for this instance
        ''' It detects if the attribute name is enum-based and will enforce data integrity
        ''' otherwise treated as free-form entry against a valid attribute definition
        ''' </summary>
        ''' <param name="attributeName">Must by a valid enum type/attribute definition <see cref="DBEnumType"/></param>
        ''' <param name="attributeNumericValue">free-flow number</param>
        ''' <param name="attributeStringValue">free-flow text, or enum name</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' Process instance attributes seems to be a prototype that was never completed.
        ''' See TMTD-201.
        ''' </remarks>
        Public Function StoreProcessInstanceAttribute(attributeName As String,
                                          attributeNumericValue As Double,
                                                      attributeStringValue As String) As ReturnCodes

            ' Is it a valid key?
            Dim enumType As DBEnumType = mEnumLookup.SingleOrDefault(Function(et) et.Name.Equals(attributeName, StringComparison.OrdinalIgnoreCase))
            If enumType Is Nothing Then
                logger.Error($"Undefined {_Process} process attribute key: {attributeName}. Check database definitions of valid enum/key types")
                Return ReturnCodes.UDBS_ERROR
            End If
            Dim stringToSave As String = attributeStringValue
            Dim valueToSave As Double = attributeNumericValue
            ' Is it an enum entry
            If enumType?.DBEnums().Any() Then
                ' Check if valid
                Dim enumItem As DBEnum = enumType.DBEnums.SingleOrDefault(Function(e) e.Name.Equals(attributeStringValue, StringComparison.OrdinalIgnoreCase))
                If enumItem Is Nothing Then
                    logger.Error($"Undefined {_Process} process attribute enum [{attributeStringValue}] for enum type: {attributeName}. Check database definitions of valid enum/key types")
                    Return ReturnCodes.UDBS_ERROR
                End If
                stringToSave = enumItem.Name
                ' NB: Save Enum ID as well if equal
                If attributeNumericValue = enumItem.ID Then
                    valueToSave = enumItem.ID
                End If
            End If

            ' Save directly to local tables, no need to keep in memory as it is anticipated that these attributes are not read very frequently
            Dim elapsedSecondsAccurateSinceStart As Double = (Stopwatch.GetTimestamp() - mStartTicks) / Stopwatch.Frequency
            Dim serverDateNowApprox = mStartTime.AddSeconds(elapsedSecondsAccurateSinceStart)
            Dim utcDate = serverDateNowApprox.ToUniversalTime()

            Dim rsTemp As New DataTable()
            Dim sqlQuery = $"SELECT [attribute_id]
                                  ,[attribute_process_id]
                                  ,[attribute_attributetype_id]
                                  ,[attribute_value]
                                  ,[attribute_stringdata]
                                  ,[attribute_updated_date_utc]
                                  ,[attribute_updated_by]
                             FROM [{mProcessAttributesTable}]
                             where attribute_process_id={mProcessID} AND attribute_attributetype_id={enumType.ID}"
            Dim pk = "attribute_id"
            Dim attibuteID As Integer
            If OpenLocalRecordSet(rsTemp, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS Then
                logger.Error($"Error querying for process attributes for process {mProcessID}")
                Return ReturnCodes.UDBS_ERROR
            End If

            Dim columnNames As New List(Of String)() From {pk, "attribute_process_id", "attribute_attributetype_id", "attribute_value", "attribute_stringdata", "attribute_updated_date_utc", "attribute_updated_by"}
            Dim columnValues As New List(Of Object)() From {0, mProcessID, enumType.ID, valueToSave, stringToSave, utcDate, EmployeeNumber}
            Using localTx = BeginLocalTransaction()

                If Not rsTemp?.AsEnumerable().Any() Then
                    ' Add new result

                    attibuteID = InsertLocalRecord(columnNames.Skip(1).ToArray(), columnValues.Skip(1).ToArray(), mProcessAttributesTable, localTx, pk)
                    columnValues(0) = attibuteID
                Else
                    attibuteID = Convert.ToInt32(rsTemp(0)(0))
                    Dim keys As String() = {pk, "attribute_process_id", "attribute_attributetype_id"}
                    columnValues(0) = attibuteID
                    UpdateLocalRecord(keys, columnNames.ToArray(), columnValues.ToArray(), mProcessAttributesTable, localTx)

                End If
                ' Insert to history table too
                InsertLocalRecord(columnNames.ToArray(), columnValues.ToArray(), mProcessAttributesHistoryTable, localTx)
            End Using

            Return ReturnCodes.UDBS_OP_SUCCESS
        End Function

        ''' <summary>
        ''' Creates a process instance
        ''' </summary>
        ''' <param name="productNumber">Product number as string.</param>
        ''' <param name="SerialNumber">Unit's Serial number.</param>
        ''' <param name="ProcessStage">Process stage.</param>
        ''' <param name="ItemListRevision">Item list revision number.</param>
        ''' <param name="ProcessID">Process ID.</param>
        ''' <returns></returns>
        Public Function StartProcessInstance(productNumber As String,
                                             SerialNumber As String,
                                             ProcessStage As String,
                                             ByRef ItemListRevision As Integer,
                                             ByRef ProcessID As Integer) _
            As ReturnCodes

            If CheckLocalTables(Me.Process) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_OP_SUCCESS
            End If

            Dim rsTemp As New DataTable
            Try
                Dim sSQL As String
                Dim tmpStr As String, arrStr As String(), lUnitID As Integer

                If mProcessRunning Then
                    ' Already in process
                    LogError(New Exception("Already in process."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Check local DB integrity.
                ' We're starting a new process, so synchronize every process
                ' in the local DB. No need to ignore any process.
                ' (i.e. set the process to ignore to 0.)
                If CheckActiveProcesses(Process, ProcessStage, productNumber, SerialNumber) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    logger.Error("Cannot start process because another application is currently performing in on the same unit.")
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Creating object as a new Process instance
                mREADONLY = False

                ' Initialize private instance variables
                mSerialNumber = UCase(Trim(SerialNumber))

                ' get unit varianc/oracle PN/catalogue PN from product group
                sSQL = "SELECT " &
                            "unit.unit_id, " &
                            "product.product_id, " &
                            "product.product_number, " &
                            "product.product_catalogue_number, " &
                            "udbs_product_group.pg_string_value " &
                       "FROM unit INNER JOIN product ON " &
                            "product.product_id = unit.unit_product_id " &
                            "LEFT JOIN udbs_unit_details ON " &
                                "udbs_unit_details.ud_unit_id = unit.unit_id " &
                                "AND udbs_unit_details.ud_identifier = 'PRD_VAR' " &
                            "LEFT JOIN udbs_product_group  ON " &
                                "udbs_unit_details.ud_pg_product_group = udbs_product_group.pg_product_group " &
                                "AND udbs_unit_details.ud_pg_sequence = udbs_product_group.pg_sequence " &
                       "WHERE " &
                            "product.product_number='" & productNumber & "' " &
                            "AND unit.unit_serial_number='" & SerialNumber & "'"

                If QueryNetworkDB(sSQL, rsTemp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                If (rsTemp.Rows.Count = 0) Then
                    ' No such unit.
                    LogError(New Exception($"No such unit: '{SerialNumber}' / '{productNumber}'"))
                    Return ReturnCodes.UDBS_ERROR
                End If

                lUnitID = KillNullInteger(rsTemp(0)("unit_id"))
                Dim productId As Integer = KillNullInteger(rsTemp(0)("product_id"))
                mUnitOraclePN = KillNull(rsTemp(0)("product_number"))
                mUnitCataloguePN = KillNull(rsTemp(0)("product_catalogue_number"))
                mUnitVariance = ""

                ' Figure out the variance.
                tmpStr = KillNull(rsTemp(0)("pg_string_value"))
                If Not String.IsNullOrEmpty(tmpStr) Then
                    arrStr = Split(tmpStr, ",")
                    If UBound(arrStr) < 2 Then
                        LogError(New Exception("Invalid variance information found."))
                        Return ReturnCodes.UDBS_ERROR
                    End If
                    mUnitOraclePN = arrStr(0)
                    mUnitCataloguePN = arrStr(1)
                    If arrStr(2) <> "-1" Then
                        mUnitVariance = arrStr(2)
                    End If
                End If

                ' Load Product
                If mPRODUCT.GetProductByID(productId) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Could not load product object
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Load Itemlist
                If mITEMLIST.LoadItemList(Process, mPRODUCT, ProcessStage, ItemListRevision) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Could not load itemlist object
                    Return ReturnCodes.UDBS_ERROR
                End If

                Using transaction = BeginNetworkTransaction()
                    ' Create a new process instance in the network DB, retrieve process_index and sequence
                    ' if revision=0, i.e., debug list
                    If mITEMLIST.Revision >= 0 Then
                        Dim localSequence As Integer = 0 ' Unused.
                        If GetNewProcessId(mProcessID, mPRODUCT.Number, mSerialNumber, lUnitID, mITEMLIST.Stage, localSequence, transaction) <> ReturnCodes.UDBS_OP_SUCCESS Then
                            ' An error occurred
                            Return ReturnCodes.UDBS_ERROR
                        End If

                        '*****
                        ' The following has been altered from the original spec
                        ' Results of debug itemlist will be stored to UDBS,
                        ' However, the data (itemlist, process & results) will be saved to .dat file on local c:\udbs_v2 up on uprevved
                        '*****
                        '    Else
                        ' This is a debug itemlist. No data may be stored against it in the network DB.. Select a random - PID
                        '        mProcessID = -CInt(Rnd(Timer) * 10000)
                    Else
                        ' wrong revision, must be >=0
                        LogError(New Exception($"Wrong itemlistrev_revision: {mITEMLIST.Revision}"))
                        Return ReturnCodes.UDBS_ERROR
                    End If

                    ' Update the local process table with this information
                    If UpdateLocalDB() <> ReturnCodes.UDBS_OP_SUCCESS Then
                        ' An error occurred
                        transaction.HasError = True
                        Return ReturnCodes.UDBS_ERROR
                    End If

                    ' Register Process Instance
                    If RegisterProcessInstance() <> ReturnCodes.UDBS_OP_SUCCESS Then
                        ' An error occurred
                        transaction.HasError = True
                        Return ReturnCodes.UDBS_ERROR
                    End If

                    ' Create the in memory recordsets for process and result
                    If CreateInMemoryRecordSets() <> ReturnCodes.UDBS_OP_SUCCESS Then
                        ' An error occurred
                        transaction.HasError = True
                        Return ReturnCodes.UDBS_ERROR
                    End If


                    Try
                        ' Create all the result entries locally.
                        CreateMissingLocalTestResults()
                    Catch ex As Exception
                        transaction.HasError = True
                        Throw
                    End Try

                    ' Whewhhh!... <pant> <pant> <pant>
                    mProcessRunning = True

                    ' Change process status to 'in process'
                    Try
                        UpdateProcessStatus(UdbsProcessStatus.IN_PROCESS)

                    Catch ex As Exception
                        ' Error attempting to update the process info
                        ' not a critical function, resume
                    End Try

                    If UpdateNetworkDB(False, transaction) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        ' Error attempting to update the process info
                        ' not a critical function, don't generate an ReturnCodes.UDBS_ERROR???
                        logger.Warn("Failed to update the network database.")
                    End If
                End Using

                ' Update process timers
                Dim startTime = ""
                If GetProcessInstanceField("start_date", startTime) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                mStartTime = KillNullDate(startTime)
                mStartTicks = Stopwatch.GetTimestamp()

                ' return the process ID to caller
                ProcessID = mProcessID

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                ' log context from this method (not this class)
                DatabaseSupport.LogErrorInDatabase(ex, Process, ProcessStage, mProcessID, productNumber, SerialNumber)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' Gets the executing station's name.
        ''' </summary>
        ''' <returns></returns>
        Friend Shared Function GetStationName() As String

            Dim localStationName = ""

            If CUtility.Utility_GetStationName(localStationName) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New ApplicationException("Unable to retrieve the station name.")
            End If

            Return localStationName
        End Function

        ''' <summary>
        ''' Marks the process status as "COMPLETED" then closes the process instance and attempts to move all process instance data 
        ''' to the network DB, then closes all DB connections.
        ''' </summary>
        ''' <returns>The outcome of the operation.</returns>
        Public Function StopProcessInstance() As ReturnCodes

            Try
                ' Update process timers
                CUtility.Utility_GetServerTime(mStopTime)

                mDurationS = (mStopTime - mStartTime).TotalSeconds

                Dim endProcessInfo = New Dictionary(Of String, String) From {
                {"process_end_date", DBDateFormat(mStopTime)},
                {"process_total_duration", Format(mDurationS, "#0.0")},
                {"process_status", UdbsProcessStatus.COMPLETED.ToString}
            }

                If StoreProcessInstanceFields(endProcessInfo) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    logger.Error("Failed to update process instance fields.")
                    Return ReturnCodes.UDBS_ERROR
                End If

                If UpdateNetworkDB(True) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Transfer failed
                    Return ReturnCodes.UDBS_ERROR
                End If

                mProcessRunning = False
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Marks as "TERMINATED" a process in the network DB but without updating the result table.
        ''' This is useful when trying to recover a process and the local DB data is found to be corrupted. The process instance is then Terminated.
        ''' </summary>
        Protected Sub TerminateWithoutSynchronizing()

            If Not ProcessInstanceRunning Then
                Throw New UDBSException("Could not terminate the process as it's not currently running.")
            End If

            ' Update process timers
            CUtility.Utility_GetServerTime(mStopTime)

            Dim endProcessInfo = New Dictionary(Of String, String) From {
                {"process_end_date", DBDateFormat(mStopTime)},
                {"process_total_duration", Format(mDurationS, "#0.0")},
                {"process_status", UdbsProcessStatus.TERMINATED.ToString}
            }

            If StoreProcessInstanceFieldsInNetworkDB(endProcessInfo) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New UDBSException("Failed to update the process instance fields. Unable to terminate the process.")
            End If

            'Delete and unregister the process from the local DB
            DeleteLocalProcess()

            mREADONLY = True
        End Sub

        ''' <summary>
        ''' Pauses a Process instance, transferring current data to the network database.
        ''' </summary>
        ''' <returns>The outcome of the operation.</returns>
        Public Function PauseProcessInstance() As ReturnCodes

            Try
                CUtility.Utility_GetServerTime(mStopTime)

                'check if the process is IN_PROCESS before pausing (Locally)
                If Not mProcessRunning Then
                    LogError(New Exception("Cannot Pause a process not currently running."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                'check if the process is IN_PROCESS before pausing (Network)
                Dim netDBprocessInfo = TestDataInterface.TestDataProcessInfo.GetProcessInfo(ProductNumber, UnitSerialNumber, Stage, Sequence, ID)
                If netDBprocessInfo.Status <> UdbsProcessStatus.IN_PROCESS Then
                    'Process needs to be in process for pause
                    LogError(New Exception($"Error when pausing, the process status is {netDBprocessInfo.Status} but expected 'IN_PROCESS' for process: {ProcessContext}"))
                    Return ReturnCodes.UDBS_ERROR
                End If

                Dim pausedProcessInfo = New Dictionary(Of String, String) From {
                {"process_end_date", DBDateFormat(mStopTime)},
                {"process_status", UdbsProcessStatus.PAUSED.ToString}
                }

                If StoreProcessInstanceFields(pausedProcessInfo) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Error attempting to update the process info.
                    Return ReturnCodes.UDBS_ERROR
                End If

                If UpdateNetworkDB(RemoveLocalCopy:=True) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Transfer failed
                    logger.Warn("Failed to update network database. We will try again next time.")
                End If

                mProcessRunning = False
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Synchronizes local DB process data with the network Db. This operation doesn't remove local data nor does it change the process status.
        ''' </summary>
        Public Sub Synchronize()

            If ProcessInstanceRunning = False Then
                Throw New ApplicationException("Error while synchronizing Process data. Process is not currently running.")
            End If

            logger.Info("Syncing local process data with network DB...")
            If UpdateNetworkDB(RemoveLocalCopy:=False) <> ReturnCodes.UDBS_OP_SUCCESS Then
                ' Transfer failed
                Throw New UDBSException("Error while synchronizing Process data. Transfer to network DB failed.")
            End If

        End Sub

        ''' <summary>
        ''' Function restarts a paused process instance specified by the product/serial/stage, bringing the process to the local db
        ''' </summary>
        ''' <param name="ProcessStage">Process stage.</param>
        ''' <param name="ProductNumber">Udbs product ID.</param>
        ''' <param name="SerialNumber">Unit serial number.</param>
        ''' <returns></returns>
        Public Function RestartUnit(ProcessStage As String,
                                    ProductNumber As String,
                                    SerialNumber As String) As ReturnCodes

            Dim ProcessID As Integer = 0
            Try
                If mProcessRunning Then
                    ' Already in process
                    LogError(New Exception("Already in process."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                'check if the process is PAUSED before Restarting
                Dim netDBprocessInfo = TestDataInterface.TestDataProcessInfo.GetProcessInfo(ProductNumber, SerialNumber, ProcessStage)
                If netDBprocessInfo.Status <> UdbsProcessStatus.PAUSED Then
                    logger.Error($"Error when Restarting, the process status is {netDBprocessInfo.Status} but expected 'PAUSED' for process: {ProcessContext}")
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Create a local copy of the process tables
                If CheckLocalTables(Process) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Check local db integrity
                If CheckActiveProcesses(Process, ProcessStage, ProductNumber, SerialNumber, activeProcessID:=ProcessID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                Dim foundLocalProcess As Boolean = ProcessID <> 0

                If Not foundLocalProcess Then
                    ' Find the process ID of the latest sequence
                    Dim LastSequence As Integer = 0
                    If GetProcessID(ProductNumber, SerialNumber, ProcessStage, LastSequence, ProcessID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Return ReturnCodes.UDBS_ERROR
                    End If
                End If

                ' Check that the last process instance was paused
                ' TODO - Note that this method does not actually check the status, despite the comment above

                ' Always load the unit and product information into local fields, even if the process was found in the local DB
                If LoadProcessInstanceByID(ProcessID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Problem loading the old process instance
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Creating object as a new Process instance
                mREADONLY = False

                If Not foundLocalProcess Then
                    ' Update the local process table with this information
                    If UpdateLocalDB() <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Return ReturnCodes.UDBS_ERROR
                    End If

                    ' Register Process Instance
                    If RegisterProcessInstance() <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Return ReturnCodes.UDBS_ERROR
                    End If
                End If


                ' Create the Current Results recordset
                If CreateInMemoryRecordSets() <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Whewhhh!... <pant> <pant> <pant>
                mProcessRunning = True

                ' Change process status to 'in process'
                Try
                    UpdateProcessStatus(UdbsProcessStatus.IN_PROCESS)

                Catch ex As Exception
                    ' Error attempting to update the process info
                    ' not a critical function, resume
                End Try

                If UpdateNetworkDB(False) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' not critical, resume and try next time.
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                ' log context from this method (not this class)
                DatabaseSupport.LogErrorInDatabase(ex, Process, ProcessStage, ProcessID, ProductNumber, SerialNumber)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Retrieve the stage, product and serial number based on the process ID.
        ''' </summary>
        ''' <param name="processName">The name of the process (testdata, wip, etc.)</param>
        ''' <param name="processId">The ID of the process to fetch.</param>
        ''' <param name="stage">(Out) The stage of the process being executed.</param>
        ''' <param name="productId">(Out) The product ID of the unit being worked on.</param>
        ''' <param name="serialNumber">(Out) The serial number of the unit being worked on.</param>
        ''' <param name="sequence">(Out) Test instance sequence number.</param>
        ''' <exception cref="UDBSException">If we fail to query the network database, or there is no such process.</exception>
        Private Shared Sub GetStageProductAndSerialNumberByProcessId(processName As String, processId As Integer, ByRef stage As String, ByRef productId As String, ByRef serialNumber As String, ByRef sequence As Integer)
            Dim queryStr = $"SELECT unit_serial_number, product_number, process_sequence, itemlistrev_stage FROM {GetProcessTableName(processName)}, {GetItemListRevisionTableName(processName)}, unit, product WHERE {GetProcessTableName(processName)}.process_unit_id = unit.unit_id AND unit.unit_product_id = product.product_id AND process_itemlistrev_id = itemlistrev_id AND process_id = {processId}"
            Dim data As DataTable = Nothing
            OpenNetworkRecordSet(data, queryStr)
            If (data.Rows.Count <> 1) Then
                Throw New UDBSException($"No such process ID ({processId})")
            End If

            productId = KillNull(data.Rows(0)("product_number"))
            serialNumber = KillNull(data.Rows(0)("unit_serial_number"))
            stage = KillNull(data.Rows(0)("itemlistrev_stage"))
            sequence = KillNullInteger(data.Rows(0)("process_sequence"))
        End Sub

        ''' <summary>
        ''' Restarts a paused process instance specified by the process id, bringing the process to the local db
        ''' </summary>
        ''' <param name="ProcessID">Process ID.</param>
        ''' <returns></returns>
        ''' <remarks>checked and modified by BC, Dec 4, 2001, seemed working, may have problem if the local database has not been flushed</remarks>
        Public Overridable Function RestartProcessID(ProcessID As Integer) As ReturnCodes

            Dim serialNumber As String = Nothing
            Dim productId As String = Nothing
            Dim stage As String = Nothing
            Dim sequence As Integer = 0

            Try
                If mProcessRunning Then
                    ' Already in process
                    LogError(New Exception("Already in process."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Create a local copy of the process tables
                If CheckLocalTables(Process) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Check local db integrity
                GetStageProductAndSerialNumberByProcessId(Process, ProcessID, stage, productId, serialNumber, sequence)

                'check if the process is PAUSED before Restarting
                Dim netDBprocessInfo = TestDataInterface.TestDataProcessInfo.GetProcessInfo(productId, serialNumber, stage, sequence, ProcessID)
                If netDBprocessInfo.Status <> UdbsProcessStatus.PAUSED Then
                    logger.Error($"Error when Restarting, the process status is {netDBprocessInfo.Status} but expected 'PAUSED' for process: {ProcessContext}")
                    Return ReturnCodes.UDBS_ERROR
                End If

                If CheckActiveProcesses(Process, stage, productId, serialNumber) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Check that the last process instance was paused
                If LoadProcessInstanceByID(ProcessID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Problem loading the old process instance
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Creating object as a new Process instance
                mREADONLY = False

                ' Update the local process table with this information
                If UpdateLocalDB() <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Register Process Instance
                If RegisterProcessInstance() <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Create the Current Results recordset
                If CreateInMemoryRecordSets() <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Create all the result entries locally.
                CreateMissingLocalTestResults()

                ' Whewhhh!... <pant> <pant> <pant>
                mProcessRunning = True

                ' Change process status to 'in process'
                If StoreProcessInstanceField("process_status", "IN PROCESS") <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Error attempting to update the process info
                    ' Not critical, resume
                End If

                If UpdateNetworkDB(False) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Error attempting to update the process info
                    ' not critical, resume
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                ' log context from this method (not this class)
                DatabaseSupport.LogErrorInDatabase(ex, Process, stage, ProcessID, productId, serialNumber)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Function loads Process Information based on Product/Unit/Stage/Sequence for Read Only access
        ''' </summary>
        ''' <param name="ProcessStage">Process stage.</param>
        ''' <param name="ProductNumber">UDBS Product ID.</param>
        ''' <param name="SerialNumber">Unit serial number.</param>
        ''' <param name="Sequence">
        ''' Process sequence number. If set to zero, loads the latest one.
        ''' </param>
        ''' <returns>Whether or not this method successfully loaded the process.</returns>
        Public Overridable Function LoadProcessInstanceByUnit(ProcessStage As String,
                                                              ProductNumber As String,
                                                              SerialNumber As String,
                                                              Sequence As Integer) As ReturnCodes

            Dim ProcessID As Integer
            Try
                If mProcessRunning Then
                    ' Already in process
                    LogError(New Exception("Already in process."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Get the process id of the specified process instance
                If GetProcessID(ProductNumber, SerialNumber, ProcessStage, Sequence, ProcessID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Load the process instance object with this process id
                Return LoadProcessInstanceByID(ProcessID)
            Catch ex As Exception
                ' log context from this method (not this class)
                DatabaseSupport.LogErrorInDatabase(ex, Process, ProcessStage, ProcessID, ProductNumber, SerialNumber)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Loads the ItemList by ID then creates in memory records of item list definitions and results tables.
        ''' </summary>
        ''' <param name="itemlist_rev_id">Item List revision ID.</param>
        ''' <returns></returns>
        Private Function LoadItemList(itemlist_rev_id As Integer) As ReturnCodes

            If mITEMLIST.LoadItemListByID(Process, itemlist_rev_id) <> ReturnCodes.UDBS_OP_SUCCESS Then
                ' Could not load itemlist object
                Return ReturnCodes.UDBS_ERROR
            End If

            ' Load the process and result information into memory
            If CreateInMemoryRecordSets() <> ReturnCodes.UDBS_OP_SUCCESS Then
                ' Fail to create the local record sets.
                Return ReturnCodes.UDBS_ERROR
            End If

            Return ReturnCodes.UDBS_OP_SUCCESS

        End Function

        ''' <summary>
        ''' Populates local fields with the unit and product info for the given <paramref name="ProcessID"/>.
        ''' Loads the item list if not already loaded
        ''' </summary>
        ''' <param name="ProcessID">Process ID.</param>
        ''' <returns>The outcome of this operation.</returns>
        Private Function LoadProcessInstanceByID(ProcessID As Integer) As ReturnCodes
            Try
                Dim sqlQuery As String
                Dim rsTemp As New DataTable
                Dim itemlist_rev_id As Integer
                Dim tmpStr As String, arrStr As String()

                If mProcessRunning Then
                    ' Already in process
                    LogError(New Exception("Already in process."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Set object information
                mProcessID = ProcessID

                ' Need to find out what itemlist we want to use
                sqlQuery = "SELECT * FROM " & mProcessTable & " with(nolock) WHERE process_id = " & CStr(mProcessID)
                If QueryNetworkDB(sqlQuery, rsTemp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                itemlist_rev_id = KillNullInteger(rsTemp(0)("process_itemlistrev_id"))
                rsTemp = Nothing

                If Not mITEMLIST.ItemListLoaded Then
                    ' Now load the itemlist object so that it can be used by CreateLocalRecordsets
                    If LoadItemList(itemlist_rev_id) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Return ReturnCodes.UDBS_ERROR
                    End If
                End If

                ' Get unit specific information
                If GetUnitInfoByID(KillNullInteger(mProcessInfo(0)("process_unit_id")), rsTemp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                mSerialNumber = KillNull(rsTemp(0)("unit_serial_number"))
                mStartTime = KillNullDate(mProcessInfo(0)("process_start_date"))

                If Not IsDBNull(rsTemp(0)("unit_created_date")) Then
                    mUnitCreatedDate = KillNullDate(rsTemp(0)("unit_created_date"))
                End If

                If Not IsDBNull(rsTemp(0)("unit_created_by")) Then
                    mUnitEmployeeNumber = KillNull(rsTemp(0)("unit_created_by"))
                End If

                If Not IsDBNull(rsTemp(0)("unit_labels_no")) Then
                    mUnitNumLabels = KillNullInteger(rsTemp(0)("unit_labels_no"))
                End If

                If Not IsDBNull(rsTemp(0)("unit_report")) Then
                    mUnitReport = KillNull(rsTemp(0)("unit_report"))
                End If

                rsTemp = Nothing

                ' get unit variance/Oracle PN/catalogue PN from product group
                sqlQuery = "SELECT product_number, product_catalogue_number " &
                           "FROM product with(nolock), unit with(nolock) " &
                           "WHERE product_id=unit_product_id " &
                           "AND unit_id=" & KillNull(mProcessInfo(0)("process_unit_id"))
                If QueryNetworkDB(sqlQuery, rsTemp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If
                mUnitOraclePN = KillNull(rsTemp(0)("product_number"))
                mUnitCataloguePN = KillNull(rsTemp(0)("product_catalogue_number"))
                mUnitVariance = ""
                rsTemp = Nothing

                sqlQuery = "SELECT * FROM udbs_unit_details with(nolock), udbs_product_group  with(nolock) " &
                           "WHERE ud_pg_product_group=pg_product_group " &
                           "AND ud_pg_sequence=pg_sequence " &
                           "AND ud_identifier='PRD_VAR' " &
                           "AND ud_unit_id=" & KillNull(mProcessInfo(0)("process_unit_id"))
                If QueryNetworkDB(sqlQuery, rsTemp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If
                If (If(rsTemp?.Rows?.Count, 0)) = 0 Then
                    ' that is fine, non-RoHS stuffs
                Else
                    If (If(rsTemp?.Rows?.Count, 0)) > 1 Then
                        LogError(New Exception("Multiple variance records found."))
                        Return ReturnCodes.UDBS_ERROR
                    Else
                        tmpStr = KillNull(rsTemp(0)("pg_string_value"))
                        arrStr = Split(tmpStr, ",")
                        If UBound(arrStr) < 2 Then
                            LogError(New Exception("Invalid variance information found."))
                            Return ReturnCodes.UDBS_ERROR
                        End If
                        mUnitOraclePN = arrStr(0)
                        mUnitCataloguePN = arrStr(1)
                        If arrStr(2) <> "-1" Then mUnitVariance = arrStr(2)
                    End If
                End If
                rsTemp = Nothing

                ' Load Product object
                If mPRODUCT.GetProductByID(mITEMLIST.ProductID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Could not load product object
                    Return ReturnCodes.UDBS_ERROR

                End If

                If Me.Status = "IN PROCESS" Then
                    mProcessRunning = True
                End If

                ' Creating object as a Process instance copy
                mREADONLY = True

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        Public Overridable Function AddProcessInstanceNote(ProcessNote As String) _
            As ReturnCodes
            ' Adding notes to the existing process_note (never delete)
            ' modified Dec 6 2001, BC

            Try
                Dim NewMessage As String

                ProcessNote = CUtility.Utility_ConvertStringToASCIICondenseInvalidCharacters(ProcessNote)
                NewMessage = KillNull(mProcessInfo(0)("process_notes")) & ProcessNote & vbCrLf

                ' Store message to table
                Return StoreProcessInstanceField("notes", NewMessage)
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Stores field names and field value pairs of process instance information.
        ''' The operation is atomic.
        ''' </summary>
        ''' <param name="info">A dictionary of field name and field value pairs.</param>
        ''' <returns>Outcome of the operation as Udbs Return codes.</returns>
        Friend Function StoreProcessInstanceFields(info As Dictionary(Of String, String)) As ReturnCodes

            Try

                If info Is Nothing OrElse Not info.Any() Then
                    LogError(New Exception("No field name and field value pair was specified for storage."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                Dim columnNames = New List(Of String)
                Dim columnValues = New List(Of String)

                For Each KeyValuePair As KeyValuePair(Of String, String) In info
                    Dim fieldName = KeyValuePair.Key
                    Dim fieldValue = KeyValuePair.Value

                    Dim desiredField As String

                    If InStr(fieldName, "process_") > 0 AndAlso Len(fieldName) > 9 Then
                        desiredField = LCase(Trim(fieldName))
                    Else
                        desiredField = "process_" & LCase(Trim(fieldName))
                    End If

                    ' Make sure the calling function is not trying to update an ID field
                    If InStr(1, desiredField, "_id") > 0 Then
                        ' Cannot update an id field
                        LogError(New Exception("Cannot update an ID field."))
                        Return ReturnCodes.UDBS_ERROR
                    End If

                    mProcessInfo(0)(desiredField) = fieldValue
                    columnNames.Add(desiredField)
                    columnValues.Add(fieldValue)

                Next
                Dim keys As String() = {"process_id"}
                columnNames.Add("process_id")
                columnValues.Add(CStr(mProcessID))


                If UpdateLocalRecord(keys, columnNames.ToArray, columnValues.ToArray, mProcessTable) Then
                    Return ReturnCodes.UDBS_OP_SUCCESS
                Else
                    Return ReturnCodes.UDBS_ERROR
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR

            End Try

        End Function

        ''' <summary>
        ''' Stores a single field value of process instance information
        ''' </summary>
        ''' <param name="FieldName">Field Name.</param>
        ''' <param name="FieldValue">Field Value.</param>
        ''' <returns>Outcome of the operation as Udbs Return codes.</returns>
        Friend Function StoreProcessInstanceField(fieldName As String,
                                                  fieldValue As String) As ReturnCodes

            Return StoreProcessInstanceFields(New Dictionary(Of String, String) From {{fieldName, fieldValue}})

        End Function

        ''' <summary>
        ''' Stores field names and field value pairs of process instance information Process instance information directly into the Network DB.
        ''' The operation is atomic.
        ''' </summary>
        ''' <param name="info">A dictionary of field name and field value pairs.</param>
        ''' <returns>Outcome of the operation as Udbs Return codes.</returns>
        Friend Function StoreProcessInstanceFieldsInNetworkDB(info As Dictionary(Of String, String)) As ReturnCodes
            Try
                If info Is Nothing OrElse Not info.Any() Then
                    LogError(New Exception("No field name and field value pair was specified for storage."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                Dim columnNames = New List(Of String)
                Dim columnValues = New List(Of String)

                For Each KeyValuePair As KeyValuePair(Of String, String) In info
                    Dim fieldName = KeyValuePair.Key
                    Dim fieldValue = KeyValuePair.Value

                    Dim desiredField As String

                    If InStr(fieldName, "process_") > 0 AndAlso Len(fieldName) > 9 Then
                        desiredField = LCase(Trim(fieldName))
                    Else
                        desiredField = "process_" & LCase(Trim(fieldName))
                    End If

                    ' Make sure the calling function is not trying to update an ID field
                    If InStr(1, desiredField, "_id") > 0 Then
                        ' Cannot update an id field
                        LogError(New Exception("Cannot update an ID field."))
                        Return ReturnCodes.UDBS_ERROR
                    End If

                    mProcessInfo(0)(desiredField) = fieldValue
                    columnNames.Add(desiredField)
                    columnValues.Add(fieldValue)

                Next
                Dim keys As String() = {"process_id"}
                columnNames.Add("process_id")
                columnValues.Add(CStr(mProcessID))

                If UpdateNetworkRecord(keys, columnNames.ToArray, columnValues.ToArray, mProcessTable) Then
                    Return ReturnCodes.UDBS_OP_SUCCESS
                Else
                    Return ReturnCodes.UDBS_ERROR
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Stores a single field value of Process instance information directly into the Network DB.
        ''' </summary>
        ''' <param name="FieldName">Field</param>
        ''' <param name="FieldValue">Value</param>

        Friend Function StoreProcessInstanceFieldInNetworkDB(FieldName As String,
                                                  FieldValue As String) _
            As ReturnCodes

            Return StoreProcessInstanceFieldsInNetworkDB(New Dictionary(Of String, String) From {{FieldName, FieldValue}})

        End Function

        Friend Function GetProcessInstanceField(FieldName As String,
                                                ByRef FieldValue As String) _
            As ReturnCodes
            ' Returns a specific process field value
            ' verified Dec 6 2001, BC
            Try

                Dim DesiredField As String
                If InStr(FieldName, "process_") > 0 AndAlso Len(FieldName) > 9 Then
                    DesiredField = LCase(Trim(FieldName))
                Else
                    DesiredField = "process_" & LCase(Trim(FieldName))
                End If

                FieldValue = KillNull(mProcessInfo(0)(DesiredField))

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        <Obsolete("This method is no longer used and will be removed.")>
        Friend Function StoreResultRecord(ResultName As String,
                                          ByRef ResultInfo As DataTable) _
            As ReturnCodes
            ' Function adds a record in the the local result recordset with the information contained in the ResultInfo
            ' not called anyway???, Dec 7 2001, BC
            Dim rsTemp As New DataTable
            Try
                ' Find Item
                Dim dr = mResultInfo.AsEnumerable().
                        FirstOrDefault(Function(x) x.Field(Of String)("itemlistdef_itemname") = ResultName)

                If IsNothing(dr) Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Update the local DB copy
                Dim sqlQuery As String = "SELECT * FROM " & mResultTable & " " &
                           "WHERE result_process_id = " & CStr(mProcessID) & " " &
                           "AND result_itemlistdef_id = " & KillNull(dr("itemlistdef_id"))
                If OpenLocalRecordSet(rsTemp, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    logger.Error("Error querying for results.")
                    Return ReturnCodes.UDBS_ERROR
                End If

                Dim keyNames As New List(Of String)() From {"result_itemlistdef_id", "result_process_id"}
                Dim columnNames As New List(Of String)() From {"result_itemlistdef_id", "result_process_id"}
                Dim columnValues As New List(Of Object)() From {KillNull(dr("itemlistdef_id")), CStr(mProcessID)}
                Dim FieldItem As DataColumn
                For Each FieldItem In rsTemp.Columns
                    If InStr(1, FieldItem.ColumnName, "_id") <> 0 Then
                        ' This is a link column, skip it
                    Else
                        columnNames.Add(FieldItem.ColumnName)
                        columnValues.Add(ResultInfo(0)(FieldItem.ColumnName))
                        dr(FieldItem.ColumnName) = ResultInfo(0)(FieldItem.ColumnName)
                    End If
                Next FieldItem

                If (If(rsTemp?.Rows?.Count, 0)) < 1 Then
                    InsertLocalRecord(columnNames.ToArray(), columnValues.ToArray(), mResultTable)
                ElseIf (If(rsTemp?.Rows?.Count, 0)) = 1 Then
                    UpdateLocalRecord(keyNames.ToArray(), columnNames.ToArray(), columnValues.ToArray(), mResultTable)
                Else
                    ' Problem with local result set
                    LogError(New Exception("Duplicate local results."))
                    Return ReturnCodes.UDBS_RECORD_EXISTS
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try
        End Function

        Friend Function StoreResultField(ResultName As String,
                                         ResultField As String,
                                         ResultValue As String) _
            As ReturnCodes
            ' Storess a single field value for the specified result item
            Dim rsTemp As New DataTable
            Try
                Dim sqlQuery As String
                Dim DesiredField As String

                ' Find Item
                Dim dr = mResultInfo.AsEnumerable().
                        FirstOrDefault(Function(x) If(x.Field(Of String)("itemlistdef_itemname"), "") = ResultName)
                If IsNothing(dr) Then
                    ' Result item not found
                    LogError(New Exception($"Result item not found: {ResultName}"))
                    Return ReturnCodes.UDBS_ERROR
                End If

                If InStr(1, ResultField, "result_") > 0 AndAlso Len(ResultField) > 8 Then
                    DesiredField = LCase(Trim(ResultField))
                Else
                    DesiredField = "result_" & LCase(Trim(ResultField))
                End If

                ' Store to local database
                sqlQuery = "SELECT * FROM " & mResultTable & " " &
                           "WHERE result_process_id = " & mProcessID & " " &
                           "AND result_itemlistdef_id = " & KillNull(dr("itemlistdef_id"))
                If OpenLocalRecordSet(rsTemp, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Throw New Exception("Error querying for result.")
                End If

                If ResultValue Is Nothing Then
                    dr(DesiredField) = DBNull.Value
                Else
                    dr(DesiredField) = ResultValue
                End If

                Dim columnNames As New List(Of String)() From {"result_itemlistdef_id", "result_process_id"}
                Dim columnValues As New List(Of Object)() From {KillNull(dr("itemlistdef_id")), CStr(mProcessID)}
                columnNames.Add(DesiredField)
                columnValues.Add(dr(DesiredField))

                If (If(rsTemp?.Rows?.Count, 0)) < 1 Then
                    ' Add new result
                    InsertLocalRecord(columnNames.ToArray(), columnValues.ToArray(), mResultTable)
                Else
                    Dim keys As String() = {"result_itemlistdef_id", "result_process_id"}
                    UpdateLocalRecord(keys, columnNames.ToArray(), columnValues.ToArray(), mResultTable)
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try
        End Function

        '**********************************************************************
        '* Process Support Functions
        '**********************************************************************

        ''' <summary>
        ''' Function finds the latest sequence for the given unit/itemlist and enters a new process instance,
        ''' or retrieves process instance information and stores a local copy.
        ''' </summary>
        ''' <param name="ProcessID">Returns as ref the newly created process ID.</param>
        ''' <param name="ProductNumber">Udbs product number.</param>
        ''' <param name="SerialNumber">Unit's serial number.</param>
        ''' <param name="UnitID">Unit ID.</param>
        ''' <param name="ItemListStage">Item list stage.</param>
        ''' <param name="NextSequence">Next sequence number found for the unit.</param>
        ''' <param name="transaction">Database transaction scope.</param>
        ''' <returns></returns>
        Private Function GetNewProcessId(ByRef ProcessID As Integer,
                                         ProductNumber As String,
                                         SerialNumber As String,
                                         UnitID As Integer,
                                         ItemListStage As String,
                                         ByRef NextSequence As Integer,
                                         Optional ByRef transaction As ITransactionScope = Nothing) _
            As ReturnCodes

            Dim rsProcess As New DataTable
            Dim transactionCreated As Boolean = False

            Try
                If transaction Is Nothing Then
                    transaction = BeginNetworkTransaction()
                    transactionCreated = True
                End If

                Dim ItemListRevID As Integer
                Dim LastSequence As Integer
                Dim LastProcessId As Integer

                ' Starting a new process instance
                If GetLastProcessId(LastProcessId, LastSequence, ProductNumber, SerialNumber, UnitID, ItemListStage) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Calculate next sequence value
                NextSequence = LastSequence + 1

                ' Get server date/time
                Dim ServerTime As Date
                CUtility.Utility_GetServerTime(ServerTime)

                ItemListRevID = mITEMLIST.ItemListRevID

                ' Add a record to recordset with new sequence value
                Dim columnNames = New String() _
                        {"process_unit_id", "process_itemlistrev_id", "process_sequence", "process_start_date",
                         "process_status"}
                Dim columnValues = New Object() {UnitID, ItemListRevID, NextSequence, DBDateFormat(ServerTime), "STARTING"}

                ProcessID = InsertNetworkRecord(columnNames, columnValues, mProcessTable, transaction, "process_id")

                logger.Info($"New Process Id: {ProcessID} for Product: {ProductNumber} Serial Number: {SerialNumber} Stage: {ItemListStage}")

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                transaction.HasError = True
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsProcess?.Dispose()

                If transactionCreated Then
                    transaction?.Dispose()
                    transaction = Nothing
                End If
            End Try
        End Function

        Private Function GetLastProcessId(ByRef ProcessID As Integer,
                                          ByRef LastSequence As Integer,
                                          ProductNumber As String,
                                          SerialNumber As String,
                                          UnitID As Integer,
                                          Stage As String) _
            As ReturnCodes
            Try
                ' Function finds the latest process instance for the given unit/stage
                'Dim ProductRelease As Double -comment Simon: variable not used was commented out
                Dim sqlQuery As String
                Dim rsProcess As New DataTable

                ' Get the process instances for this unit
                sqlQuery = "SELECT p.* " &
                           "FROM " & mProcessTable & " p with(nolock), " & mItemListRevisionTable & " ir with(nolock) " &
                           "WHERE p.process_itemlistrev_id = ir.itemlistrev_id " &
                           "AND p.process_unit_id = " & CStr(UnitID) & " " &
                           "AND ir.itemlistrev_stage = '" & Stage & "' " &
                           "ORDER BY p.process_sequence "
                If QueryNetworkDB(sqlQuery, rsProcess) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    LastSequence = 0
                    Return ReturnCodes.UDBS_ERROR
                ElseIf (If(rsProcess?.Rows?.Count, 0)) > 0 Then
                    ' Get last sequence
                    LastSequence = KillNullInteger(rsProcess.AsEnumerable().Last()("process_sequence"))
                    ProcessID = KillNullInteger(rsProcess.AsEnumerable().Last()("process_id"))
                Else
                    ' There is no process instance in the db corresponding to this process id
                    LastSequence = 0
                End If
                rsProcess = Nothing
                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                ' log context from this method (not this class)
                DatabaseSupport.LogErrorInDatabase(ex, Process, Stage, ProcessID, ProductNumber, SerialNumber)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Fetch the row representing the current process from the local DB.
        ''' </summary>
        ''' <returns>
        ''' The row representing this process.
        ''' Will be 'null' (Nothing) if there is no such row in the local DB.       
        ''' </returns>
        Private Function RetrieveLocalProcessInfoRow() As DataRow
            Dim dtLocalProcess As New DataTable
            ' Load existing process info
            Dim sqlQuery = $"SELECT * FROM {mProcessTable} WHERE process_id = {mProcessID}"
            Dim result = OpenLocalRecordSet(dtLocalProcess, sqlQuery)
            If result = ReturnCodes.UDBS_TABLE_MISSING Then
                ' Make sure the tables exists.
                If CheckLocalTables(Process) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Throw New UDBSException($"Failed to initialize the local tables for process {Process}")
                End If

                ' Now the DB tables exist... but for sure, the process won't be there!
                Return Nothing
            End If

            If (dtLocalProcess.Rows.Count > 0) Then
                If (dtLocalProcess.Rows.Count > 1) Then
                    logger.Warn($"Duplicate local DB entries for process ID {mProcessID}")
                End If

                Return dtLocalProcess.Rows(0)
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' Retrieve the network process information table and the row
        ''' representing the current process.
        ''' </summary>
        ''' <param name="table">(Out) The process info table.</param>
        ''' <param name="row">
        ''' (Out) The row representing the current process.
        ''' May be null if not found.
        ''' </param>
        Private Sub RetrieveNetworkProcessDataTableAndRow(ByRef table As DataTable, ByRef row As DataRow)
            row = Nothing
            Dim sqlQuery = "SELECT * FROM " & mProcessTable & " with(nolock) WHERE process_id = " & CStr(mProcessID)
            If QueryNetworkDB(sqlQuery, table) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New Exception($"Error retrieving process information (id = {mProcessID}) from network database.")
            End If

            If (table.Rows.Count > 0) Then
                If (table.Rows.Count > 1) Then
                    logger.Warn($"Duplicate network DB entries for process ID {mProcessID}")
                End If

                row = table.Rows(0)
            End If
        End Sub

        ''' <summary>
        ''' Create the in-memory process information table (i.e. 'mProcessInfo' member variable)
        ''' by replicating the structure of the network database.
        ''' It also loads and merges the data from the network and local database.
        ''' </summary>
        Private Sub CreateInMemoryProcessInformationTable()
            Dim localProcessInfoRow As DataRow = RetrieveLocalProcessInfoRow()
            Dim networkProcessInfoTable = New DataTable()
            Dim networkProcessInfoRow As DataRow = Nothing
            RetrieveNetworkProcessDataTableAndRow(networkProcessInfoTable, networkProcessInfoRow)

            mProcessInfo = New DataTable()
            Dim aColumn As DataColumn
            For Each aColumn In networkProcessInfoTable.Columns
                mProcessInfo.Columns.Add(aColumn.ColumnName, aColumn.DataType)
            Next aColumn

            Dim mergedProcessInfo As Object() = CType(networkProcessInfoRow.ItemArray.Clone(), Object())
            Dim index As Integer = 0
            For Each networkValue In networkProcessInfoRow.ItemArray
                If (localProcessInfoRow IsNot Nothing) Then
                    Dim localValue = localProcessInfoRow(index)

                    If (localValue Is Nothing And networkValue Is Nothing) Then
                        ' Both values are null.
                        ' No ambiguity.
                    ElseIf (localValue Is Nothing) Then
                        ' Network value exists, but local value is null.
                        ' Was the local DB flushed?
                        logger.Warn("Network is more up-to-date than local DB.")
                    ElseIf (networkValue Is Nothing) Then
                        ' Local DB has value, but network DB has no value.
                        mergedProcessInfo(index) = localValue
                    ElseIf (Not networkValue.Equals(localValue)) Then
                        ' Values differ.
                        ' Local DB wins.
                        mergedProcessInfo(index) = localValue
                    End If

                    index += 1
                End If
            Next

            mProcessInfo.Rows.Add(mergedProcessInfo)
        End Sub

        ''' <summary>
        ''' Creates the in-memory test result table (i.e. the 'mResultInfo' member variable).
        ''' It replicates and combines the structure of both the network 'item specifications' 
        ''' and 'test results' tables into an hybrid 'joined' table.
        ''' </summary>
        Private Sub CreateInMemoryTestResultTable()
            Dim itemDefinitions As DataTable = mITEMLIST.Items_RS
            Dim aColumn As DataColumn
            mResultInfo = New DataTable()
            For Each aColumn In itemDefinitions.Columns
                mResultInfo.Columns.Add(aColumn.ColumnName, aColumn.DataType)
            Next aColumn

            ' Just query the top row. We don't really care about the data, we just
            ' need to get the table's structure.
            Dim networkResults As New DataTable()
            Dim sqlQuery = "SELECT TOP 1 * FROM " & mResultTable & " with(nolock) WHERE result_process_id = " &
                           CStr(mProcessID)
            If QueryNetworkDB(sqlQuery, networkResults) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New Exception("Error querying process information table structure.")
            End If

            For Each aColumn In networkResults.Columns
                mResultInfo.Columns.Add(aColumn.ColumnName, aColumn.DataType)
            Next aColumn
        End Sub

        ''' <summary>
        ''' Loads the data from the local and network database, and merges it
        ''' into the in-memory data table.
        ''' </summary>
        Private Sub LoadAndMergeInMemoryTestResults()
            Dim sqlQuery As String
            Dim dtResults As New DataTable()
            sqlQuery = "SELECT * FROM " & mResultTable & " with(nolock) WHERE result_process_id = " &
                           CStr(mProcessID)
            If QueryNetworkDB(sqlQuery, dtResults) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New Exception("Error retrieving results from network database.")
            End If

            ' Build the local result record set.
            Dim localResults = New DataTable()
            sqlQuery = "SELECT * FROM " & mResultTable & " WHERE result_process_id = " &
                           CStr(mProcessID)
            If OpenLocalRecordSet(localResults, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New Exception("Error retrieving results from local database.")
            End If

            Dim networkIndex = New IndexedResultSet(Of Long)(dtResults, "result_itemlistdef_id")
            Dim localIndex = New IndexedResultSet(Of Long)(localResults, "result_itemlistdef_id")

            For Each itemDefinitionRow As DataRow In mITEMLIST.Items_RS.Rows
                Dim drResultInfo As DataRow = mResultInfo.NewRow()
                ' It's a bit puzzling, but don't use 'copy row' here.
                ' Copy the ItemArray instead.
                drResultInfo.ItemArray = itemDefinitionRow.ItemArray

                Dim itemDefinitionId = KillNullInteger(itemDefinitionRow("itemlistdef_id"))

                Dim networkRecord = networkIndex.FindRow(itemDefinitionId)
                CopyRow(dtResults.Columns, networkRecord, drResultInfo)

                Dim localRecord = localIndex.FindRow(itemDefinitionId)
                CopyRow(localResults.Columns, localRecord, drResultInfo)

                mResultInfo.Rows.Add(drResultInfo)
            Next
        End Sub

        ''' <summary>
        ''' Create any missing local test results entry for the current process.
        ''' </summary>
        Private Sub CreateMissingLocalTestResults()
            Dim sqlQuery As String
            Dim networkResults As New DataTable()

            sqlQuery = "SELECT * FROM " & mResultTable & " with(nolock) WHERE result_process_id = " &
                               CStr(mProcessID)
            If QueryNetworkDB(sqlQuery, networkResults) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New Exception("Error retrieving results from network database.")
            End If

            ' Build the local result record set.
            Dim localResults = New DataTable()
            sqlQuery = "SELECT * FROM " & mResultTable & " WHERE result_process_id = " &
                               CStr(mProcessID)
            If OpenLocalRecordSet(localResults, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New Exception("Error retrieving results from local database.")
            End If

            Dim networkIndex = New IndexedResultSet(Of Long)(networkResults, "result_itemlistdef_id")
            Dim localIndex = New IndexedResultSet(Of Long)(localResults, "result_itemlistdef_id")

            Dim localItemListDefinitionMapping As New DataTable()
            localItemListDefinitionMapping.Columns.Add("result_itemlistdef_id", GetType(Integer))
            localItemListDefinitionMapping.Columns.Add("result_process_id", GetType(String))

            For Each itemDefinitionRow As DataRow In mITEMLIST.Items_RS.Rows
                Dim itemName = KillNull(itemDefinitionRow("itemlistdef_itemname"))
                Dim itemDefinitionId = KillNullInteger(itemDefinitionRow("itemlistdef_id"))

                Dim localRecord = localIndex.FindRow(itemDefinitionId)

                If localRecord Is Nothing Then
                    localItemListDefinitionMapping.Rows.Add({itemDefinitionId, CStr(mProcessID)})
                End If
            Next

            If localItemListDefinitionMapping.Rows.Count > 0 Then
                Dim columnNames As New List(Of String)() From {"result_itemlistdef_id", "result_process_id"}
                InsertLocalRecords(columnNames.ToArray(), localItemListDefinitionMapping, mResultTable)
            End If

        End Sub

        ''' <summary>
        ''' This builds an in-memory merge of itemlist definitions and results tables.
        ''' It compares the local DB values with network DB values and merges from network
        ''' DB if missing in local DB.
        ''' </summary>
        ''' <returns>A result code expressing the outcome of this operation.</returns>
        ''' <remarks>
        ''' Performance timing was measured. Invoking this methods takes roughly 15 msec.
        ''' for a set of roughly 1000 "test data items".
        ''' Performance should grow linearly, i.e. O(n).
        ''' </remarks>
        Private Function CreateInMemoryRecordSets() As ReturnCodes
            Try
                ' Process Info
                CreateInMemoryProcessInformationTable()

                ' Result Info
                CreateInMemoryTestResultTable()
                LoadAndMergeInMemoryTestResults()

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Copy a row's column into another row.
        ''' Don't copy IDs.
        ''' Don't overwrite 'data' with 'no-data' (partial data).
        ''' </summary>
        ''' <param name="columns">The columns of the source table.</param>
        ''' <param name="source">The row from the source table to copy from.</param>
        ''' <param name="destination">The row to copy the data to.</param>
        Private Sub CopyRow(columns As DataColumnCollection, source As DataRow, ByRef destination As DataRow)
            If (source Is Nothing) Then
                ' Nothing to copy. This is not an error.
                Return
            End If
            If (destination Is Nothing) Then
                Throw New ArgumentException("Destination row should not be null.")
            End If

            For Each aColumn As DataColumn In columns
                If (source(aColumn.ColumnName).GetType() = GetType(System.DBNull)) Then
                    ' Nothing to copy.
                    Continue For
                End If
                If (Not destination(aColumn.ColumnName).Equals(source(aColumn.ColumnName))) Then
                    If (aColumn.ColumnName.EndsWith("_id")) Then
                        ' Don't overwrite IDs.
                        ' The order into which they are inserted matters.
                        Continue For
                    ElseIf source(aColumn.ColumnName).ToString = String.Empty Then
                        ' Don't push an empty value onto a value.
                        Continue For
                    End If
                End If

                If (source(aColumn.ColumnName).GetType() <> aColumn.DataType) Then
                    logger.Warn($"Type mismatch for column {aColumn.ColumnName} of table {mResultTable}: {source(aColumn.ColumnName).GetType()} != {aColumn.DataType}")
                End If

                destination(aColumn.ColumnName) = source(aColumn.ColumnName)
            Next
        End Sub

        ''' <summary>
        ''' Get the ID of a given process (Product/Unit/Stage/Sequence).
        ''' </summary>
        ''' <param name="ProductNumber">The product number.</param>
        ''' <param name="SerialNumber">The unit's serial number.</param>
        ''' <param name="ProcessStage">The process stage.</param>
        ''' <param name="Sequence">
        ''' Process sequence number to load.
        ''' If set to zero, loads the latest one, and returns by reference.
        ''' </param>
        ''' <param name="ProcessID">(Out)</param>
        ''' <returns></returns>
        Private Function GetProcessID(ProductNumber As String,
                                      SerialNumber As String,
                                      ProcessStage As String,
                                      ByRef Sequence As Integer,
                                      ByRef ProcessID As Integer) As ReturnCodes

            Dim rsTemp As New DataTable
            Dim DesiredSequence As Integer = Sequence
            Try
                If Sequence = 0 Then
                    ' Select latest sequence
                    If GetSequenceCount(ProductNumber, SerialNumber, ProcessStage, DesiredSequence) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Return ReturnCodes.UDBS_ERROR
                    End If

                    Sequence = DesiredSequence
                End If

                Dim sqlQuery As String = "SELECT process_id " &
                           "FROM " & mProcessTable & " with(nolock), " & mItemListRevisionTable & " with(nolock), " &
                           mUnitTable & " with(nolock), " & mProductTable & "  with(nolock) " &
                           "WHERE process_itemlistrev_id=itemlistrev_id " &
                           "AND product_id=unit_product_id " &
                           "AND unit_id=process_unit_id " &
                           "AND product_number = '" & ProductNumber & "' " &
                           "AND unit_serial_number = '" & SerialNumber & "' " &
                           "AND process_sequence = " & CStr(DesiredSequence) &
                           " AND itemlistrev_stage = '" & ProcessStage & "'"

                OpenNetworkRecordSet(rsTemp, sqlQuery)

                If (If(rsTemp?.Rows?.Count, 0)) = 0 Then
                    Throw New Exception("Process Id not found in Network DB")
                End If

                ' Assign referenced parameter on success.
                ProcessID = KillNullInteger(rsTemp(0)("process_id"))

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                ' log context from this method (not this class)
                DatabaseSupport.LogErrorInDatabase(ex, Process, ProcessStage, DesiredSequence, ProductNumber, SerialNumber)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' Gets the number of sequences (maximum process_sequence) for the given product, serial number, and stage.
        ''' </summary>
        ''' <param name="ProductNumber"></param>
        ''' <param name="SerialNumber"></param>
        ''' <param name="ProcessStage"></param>
        ''' <param name="Sequences">Sequence Count returned by reference</param>
        ''' <returns><see cref="ReturnCodes.UDBS_ERROR"/> if no sequences found, else <see cref="ReturnCodes.UDBS_OP_SUCCESS"/></returns>
        Private Function GetSequenceCount(ProductNumber As String,
                                          SerialNumber As String,
                                          ProcessStage As String,
                                          ByRef Sequences As Integer) As ReturnCodes

            Dim logMessage As String = $"Product='{ProductNumber}; Serial Number='{SerialNumber}'; Process Stage='{ProcessStage}'"
            Dim rsTemp As New DataTable
            Try
                Dim sqlQuery As String = "SELECT MAX(process_sequence) AS maxSeq " &
                           "FROM " & mProductTable & " with(nolock), " & mUnitTable & " with(nolock), " & mProcessTable &
                           " with(nolock), " & mItemListRevisionTable & " with(nolock) " &
                           "WHERE product_id=unit_product_id " &
                           "AND unit_id=process_unit_id " &
                           "AND itemlistrev_id=process_itemlistrev_id " &
                           "AND product_number = '" & ProductNumber & "' " &
                           "AND unit_serial_number = '" & SerialNumber & "' " &
                           "AND itemlistrev_stage = '" & ProcessStage & "'"

                OpenNetworkRecordSet(rsTemp, sqlQuery)
                If (If(rsTemp?.Rows?.Count, 0)) = 0 OrElse IsDBNull(rsTemp(0)("maxSeq")) Then
                    logger.Trace($"No sequence found for {logMessage}.")
                    Sequences = 0
                    Return ReturnCodes.UDBS_OP_FAIL
                Else
                    Sequences = KillNullInteger(rsTemp(0)("maxSeq"))
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                ' log context from this method (not this class)
                DatabaseSupport.LogErrorInDatabase(ex, Process, ProcessStage, -1, ProductNumber, SerialNumber)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try
        End Function

        '**********************************************************************
        '* Unit Support Functions
        '**********************************************************************

        ' Function fills recordset argument with unit and product information
        Private Function GetUnitInfo(ProductNumber As String,
                                     SerialNumber As String,
                                     ByRef UnitInfo As DataTable) As ReturnCodes
            'TODO - Move GetUnitInfo to new class.  This could be a shared method that returns a UnitInfo object
            Try
                Dim sqlQuery As String =
                           "SELECT * FROM " & mProductTable & " with(nolock), " & mUnitTable & " with(nolock) " &
                           "WHERE product_id = unit_product_id " &
                           "AND product_number = '" & ProductNumber & "' " &
                           "AND unit_serial_number = '" & SerialNumber & "'"
                Return QueryNetworkDB(sqlQuery, UnitInfo)
            Catch ex As Exception
                ' log context from this method (not this class)
                DatabaseSupport.LogErrorInDatabase(ex, Process, "Unit", 0, ProductNumber, SerialNumber)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ' Function fills recordset argument with unit and product information
        Private Function GetUnitInfoByID(UnitID As Integer,
                                         ByRef UnitInfo As DataTable) As ReturnCodes
            'TODO - Move GetUnitInfoByID to new class.  This could be a shared method that returns a UnitInfo object
            Try
                Dim sqlQuery As String =
                           "SELECT * FROM " & mProductTable & " with(nolock), " & mUnitTable & " with(nolock) " &
                           "WHERE product_id = unit_product_id " &
                           "AND unit_id = " & CStr(UnitID)
                Return QueryNetworkDB(sqlQuery, UnitInfo)
            Catch ex As Exception
                ' log context from this method (not this class)
                DatabaseSupport.LogErrorInDatabase(ex, Process, "Unit", UnitID, String.Empty, String.Empty)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        Private Function GetUnitId(ProductNumber As String,
                                   SerialNumber As String,
                                   ByRef UnitID As Integer) _
            As ReturnCodes
            ' Function returns the unit id of the specified product/unit
            'TODO - Move GetUnitId to new class.  This could be a shared method.
            Try
                Dim rsTemp As DataTable = Nothing

                If GetUnitInfo(ProductNumber, SerialNumber, rsTemp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                UnitID = KillNullInteger(rsTemp(0)("unit_id"))

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                ' log context from this method (not this class)
                DatabaseSupport.LogErrorInDatabase(ex, Process, "Unit", 0, ProductNumber, SerialNumber)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        Friend Shared ReadOnly Property CurrentWindowsProcessID As String
            Get
                Return CustomFormatWindowsProcessID(Diagnostics.Process.GetCurrentProcess)
            End Get
        End Property

        ''' <summary>
        ''' Updates the current process' status along with the process station.
        ''' </summary>
        ''' <param name="newStatus">New UdbsProcessStatus</param>
        ''' <exception cref="UDBSException">Throws an exception when failed to store process info in local DB.</exception>
        Private Sub UpdateProcessStatus(newStatus As UdbsProcessStatus)

            Dim processInfo = New Dictionary(Of String, String)

            Try

                If (newStatus.ToString = "IN_PROCESS") Then
                    processInfo.Add("process_status", "IN PROCESS")
                Else
                    processInfo.Add("process_status", newStatus.ToString)
                End If

                processInfo.Add("process_station", GetStationName())

            Catch ex As ApplicationException
                ' GetStationName throws an exception when failed. Not crititcal, ignore.
                ' Just try to store the process status.
            End Try

            If StoreProcessInstanceFields(processInfo) <> ReturnCodes.UDBS_OP_SUCCESS Then
                ' Error attempting to update the process info
                ' not a critical function, resume
                logger.Error("Failed to update process instance fields in the local DB.")
                Throw New UDBSException("Failed to update process instance fields in the local DB.")
            End If

        End Sub

        Private Function RegisterProcessInstance() As ReturnCodes
            ' Function registers this process instance as in process, creating a MUTEX object
            ' old code used Mutexes and was messy. This code now uses .NET system.Diagnostics.Process 
            Dim CustomWindowsProcessID As String

            Try
                Try
                    ' Unique name for the mutex
                    CustomWindowsProcessID = CurrentWindowsProcessID
                Catch ex As Exception
                    Throw _
                        New Exception(
                            "Could not access current Windows process information to record the instance of this application.")
                End Try

                ' Store the information in the local registration table
                Dim sqlQuery = "INSERT INTO " & mProcessRegistrationTable &
                               " (pr_process, pr_mutex_name, pr_process_id) VALUES ('" & Process & "', '" &
                               CustomWindowsProcessID & "', " & CStr(mProcessID) & ") "
                ExecuteLocalQuery(sqlQuery)

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                Throw New Exception("Failed to register " & Process & " instance in local database. " & ex.Message, ex)

            End Try
        End Function

        ''' <summary>
        ''' Unregisters this process instance from the local registration table.
        ''' </summary>
        Private Sub UnRegisterProcessInstance(Optional transaction As ITransactionScope = Nothing)
            logger.Trace("Unregistering the process from the local DB.")

            Dim sqlQuery As String = "DELETE FROM " & mProcessRegistrationTable & " WHERE pr_process_id = " & mProcessID

            If transaction Is Nothing Then
                transaction = BeginLocalTransaction()
                ExecuteLocalQuery(sqlQuery, transaction)
                transaction.Dispose()

            Else
                ExecuteLocalQuery(sqlQuery, transaction)
            End If

        End Sub

        ''' <summary>
        ''' Unregister all processes. Applications shouldn't call this. This is only meant to be
        ''' used by the unit tests.
        ''' </summary>
        Friend Shared Sub UnRegisterAllProcessInstances()

            Using transaction = BeginLocalTransaction()
                Dim sqlQuery As String = "DELETE FROM " & mProcessRegistrationTable
                ExecuteLocalQuery(sqlQuery, transaction)
            End Using

        End Sub

        ''' <summary>
        ''' Checks for the presence of active processes in the local database. Returns by reference the active process ID found in the DB.
        ''' This does more than just 'checking'.
        ''' It actually uploads TestData stragglers stuck in the local data store.
        ''' It also validates that the process represented by the four (4) first parameters
        ''' is not currently being run by another application instance on the same station.
        ''' </summary>
        ''' <param name="processName">The name of the process.</param>
        ''' <param name="stage">The stage of the process.</param>
        ''' <param name="productId">The product of the unit being worked on.</param>
        ''' <param name="serialNumber">The serial number of the unit being worked on.</param>
        ''' <param name="localDBCorrupted">Returns true when the local Db integrity check failed for the active process found. Returns false otherwise.</param>
        ''' <param name="activeProcessID">Out Parameter. Returns the active process ID found in the local DB.</param>
        ''' <returns>
        ''' Whether or not the operation succeeded.
        ''' Failure to synchronize data to the network DB is logged, but not reported as
        ''' an error because it will be tried again next time.
        ''' The only error condition is if the process described by the four (4) first parameters
        ''' is currently being executed by another application on the same station.
        ''' </returns>
        Friend Shared Function CheckActiveProcesses(
                processName As String,
                stage As String,
                productId As String,
                serialNumber As String,
                Optional ByRef localDBCorrupted As Boolean = False,
                Optional ByRef activeProcessID As Integer = -1) As ReturnCodes

            Dim sqlQuery As String
            Dim rsTemp As New DataTable
            Dim sameUnitOnDifferentAppErrorCount = 0
            localDBCorrupted = False

            sqlQuery = "SELECT * FROM " & mProcessRegistrationTable
            If (OpenLocalRecordSet(rsTemp, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS) Then
                logger.Error("Unable to query the local DB in order to find active processes in the local DB.")
                Return ReturnCodes.UDBS_ERROR
            End If

            If (If(rsTemp?.Rows?.Count, 0)) <= 0 Then
                ' No active process.
                ' Nothing to synchronize.
                Return ReturnCodes.UDBS_OP_SUCCESS
            End If

            logger.Info($"{rsTemp.Rows.Count} active Processe(s) found in the local database.")
            For Each dr As DataRow In rsTemp.Rows
                Dim thatWindowsProcessCustomID As String = KillNull(dr("pr_mutex_name"))
                Dim thatProcessName As String = KillNull(dr("pr_process"))
                Dim thatProcessID As Integer = KillNullInteger(dr("pr_process_id"))
                Dim thatProcessSerialNumber As String = Nothing
                Dim thatProcessProductId As String = Nothing
                Dim thatProcessStage As String = Nothing
                Dim thatSequence As Integer = 0

                Try
                    GetStageProductAndSerialNumberByProcessId(thatProcessName, thatProcessID, thatProcessStage, thatProcessProductId, thatProcessSerialNumber, thatSequence)
                Catch ex As Exception
                    logger.Warn(ex, $"{ex.Message} The operation will be retried later.")
                    Continue For
                End Try

                ' Determine if the process is the one we are currently starting (or restarting)
                Dim isSameProcessAndUnit = (thatProcessName = processName And thatProcessStage = stage And thatProcessProductId = productId And thatProcessSerialNumber = serialNumber)

                Dim isSameWindowsProcess = (thatWindowsProcessCustomID = CurrentWindowsProcessID)

                Dim isProcessRunning = OsAbstractionLayer.Instance.IsProcessStillRunning(thatWindowsProcessCustomID)

                If isSameProcessAndUnit AndAlso Not isSameWindowsProcess AndAlso isProcessRunning Then
                    sameUnitOnDifferentAppErrorCount = +1
                    Continue For
                End If

                Dim thatProcessStatus = GetLocalProcessStatus(thatProcessName, thatProcessID, thatProcessStage, thatProcessProductId, thatProcessSerialNumber)

                logger.Info($"Local process found with ProcessID: {thatProcessID}, SerialNumber: {thatProcessSerialNumber}, ProductID: {thatProcessProductId}, Stage: {thatProcessStage}, Status: {thatProcessStatus}. Is same process and unit: {isSameProcessAndUnit}. Is same windows process: {isSameWindowsProcess}. Is process running: {isProcessRunning}.")

                Select Case thatProcessStatus

                    Case UdbsProcessStatus.IN_PROCESS, UdbsProcessStatus.PAUSED

                        If Not isSameProcessAndUnit And Not isSameWindowsProcess And Not isProcessRunning Then
                            'Different unit and process from a crashed application: Terminate, Upload and remove data
                            UploadLocalData(thatProcessName, thatWindowsProcessCustomID, thatProcessID, removeLocalData:=True)
                        Else
                            If (isSameProcessAndUnit) Then
                                ' Check the local database integrity
                                Dim integrityStatus As LocalDBIntegrityStatus = LocalDBIntegrityChecker.Check(thatProcessID.ToString)
                                If integrityStatus <> LocalDBIntegrityStatus.Good Then
                                    'local db integrity check failed
                                    localDBCorrupted = True
                                Else
                                    ' Just upload data but don't remove, use the local process (return active process ID)
                                    UploadLocalData(processName, thatWindowsProcessCustomID, thatProcessID, removeLocalData:=False)
                                    activeProcessID = thatProcessID
                                End If
                            End If
                        End If

                    Case UdbsProcessStatus.TERMINATED

                        If (isSameWindowsProcess AndAlso isProcessRunning) Or (Not isSameWindowsProcess AndAlso Not isProcessRunning) Then
                            'Upload and remove data
                            UploadLocalData(thatProcessName, thatWindowsProcessCustomID, thatProcessID, removeLocalData:=True)
                        Else
                            'Do nothing : Different process running in a different application : Let that application upload the data
                        End If

                    Case UdbsProcessStatus.COMPLETED

                        'Always upload and remove data
                        UploadLocalData(thatProcessName, thatWindowsProcessCustomID, thatProcessID, removeLocalData:=True)

                    Case Else
                        'It should never get here but this case would be to handle the "STARTING" or "UNKNOWN" status.
                        'Do nothing. The process will eventually be restarted.
                End Select
            Next

            If sameUnitOnDifferentAppErrorCount > 0 Then
                logger.Error($"Another Windows application is currently running on this computer for process '{processName}', process ID:'{productId}', and unit Serial Number:'{serialNumber}'.")

                Return ReturnCodes.UDBS_ERROR
            End If

            Return ReturnCodes.UDBS_OP_SUCCESS
        End Function

        ''' <summary>
        ''' Converts the enum UdbsProcessStatus to a string.
        ''' </summary>
        ''' <param name="processStatus">The process status as an enum.</param>
        ''' <returns>String: "IN PROCESS".</returns>
        Public Shared Function GetProcessStatusLabel(processStatus As UdbsProcessStatus) As String
            If processStatus = UdbsProcessStatus.IN_PROCESS Then
                Return "IN PROCESS"
            End If
            Return processStatus.ToString()
        End Function

        ''' <summary>
        ''' Returns the process status from the local DB.
        ''' </summary>
        ''' <param name="processName">Process name.</param>
        ''' <param name="localProcessID">Local Db process ID.</param>
        ''' <returns>The process status as an enum.</returns>
        Private Shared Function GetLocalProcessStatus(processName As String, localProcessID As Integer, stage As String, productID As String, serialNumber As String) As UdbsProcessStatus

            Dim sqlQuery = $"SELECT process_status FROM {processName}_process WHERE process_id = {localProcessID}"

            Dim rsTempProcessStatus As New DataTable

            If (OpenLocalRecordSet(rsTempProcessStatus, sqlQuery) = ReturnCodes.UDBS_OP_SUCCESS) Then
                If rsTempProcessStatus.Rows.Count = 1 Then
                    Return GetProcessStatusEnum(KillNull(rsTempProcessStatus.Rows(0)("process_status")))
                End If

                ' log context from this method (not this class)
                DatabaseSupport.LogErrorInDatabase(New UDBSException($"{rsTempProcessStatus.Rows.Count} row(s) were found in the {processName}_process table of the local DB. There should only be exactly one row per process. This a data integrity error!"), processName, stage, localProcessID, productID, serialNumber)
            End If

            Return UdbsProcessStatus.UNKNOWN
        End Function

        ''' <summary>
        ''' Converts a process status given as a string to the corresponding UdbsProcessStatus enum
        ''' </summary>
        ''' <param name="status">process status as string.</param>
        ''' <returns>Corresponding Udbs process status enum.</returns>
        Public Shared Function GetProcessStatusEnum(status As String) As UdbsProcessStatus
            Dim statusEnum As UdbsProcessStatus = UdbsProcessStatus.UNKNOWN
            Dim statusString As String = status.Replace(" ", "_")
            If Not [Enum].TryParse(Of UdbsProcessStatus)(statusString, statusEnum) Then
                logger.Warn($"UDBS Process Status string '{status}' is invalid. Returning '{statusEnum}'.")
            End If
            Return statusEnum
        End Function

        ''' <summary>
        ''' Upload orphaned process data found in the loca DB to the network DB
        ''' </summary>
        ''' <param name="processName">Process name</param>
        ''' <param name="processRegistrationName">Custom Windows process ID. Consists of the date the process was created followed by the Windows process ID.</param>
        ''' <param name="processID">Process ID as integer</param>
        ''' <param name="removeLocalData">Remove the local data after uploading to the network DB.</param>
        ''' <returns>True when the operation was successfull, False otherwise.</returns>
        Private Shared Function UploadLocalData(processName As String, processRegistrationName As String, processID As Integer, removeLocalData As Boolean) As Boolean

            Dim orphanedProcess As New CProcessInstance(processName)
            If orphanedProcess.LoadProcessInstanceByID(processID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                logger.Warn($"Couldn't load the process instance with custom process ID: {processRegistrationName} in order to upload data to network DB. Data remains in the local database. Upload operation will be retried later.")

                Return False
            End If

            'will set the process status to TERMINATED and upload and remove data from local DB.
            If orphanedProcess.UpdateNetworkDB(removeLocalData) <> ReturnCodes.UDBS_OP_SUCCESS Then

                logger.Warn($"Could not sync up with network database for Local custom Process Id: {processRegistrationName} PID: {processID} Process Name: {processName}. Data remains in the local database. Upload operation will be retried later.")
                Return False
            End If

            logger.Info($"Uploaded the following process: custom Process Id: {processRegistrationName} PID: {processID} Process Name: {processName} to the network DB.")

            Return True
        End Function


        ''' <summary>
        ''' The main engine.
        ''' This function uploads all process header and result information to the network DB
        ''' and optionally, upon succesful completion remove the old data from the local DB.
        ''' </summary>
        ''' <param name="RemoveLocalCopy">Whether or not to remove local copy on success.</param>
        ''' <returns>Whether or not the operation succeeds.</returns>
        Friend Overridable Function UpdateNetworkDB(RemoveLocalCopy As Boolean,
                                                    Optional transaction As ITransactionScope = Nothing) As ReturnCodes

            Dim updater As New CNetworkDatabaseProcessUpdater(Me, RemoveLocalCopy)
            Dim transactionCreated As Boolean = False

            Try
                If transaction Is Nothing Then
                    transaction = BeginNetworkTransaction()
                    transactionCreated = True
                End If

                Return updater.Execute(transaction)

            Finally
                If transactionCreated Then
                    transaction?.Dispose()
                    transaction = Nothing
                End If
            End Try

        End Function

        ''' <summary>
        ''' Delete the process from the local DB.
        ''' </summary>
        Public Sub DeleteLocalProcess()

            Using transaction = BeginLocalTransaction()
                Try
                    ' Clean up the local db
                    Dim sqlQuery = $"DELETE FROM {ResultTable} WHERE result_process_id = {ID}"
                    ExecuteLocalQuery(sqlQuery, transaction)
                    sqlQuery = $"DELETE FROM {ProcessTable} WHERE process_id = {ID}"
                    ExecuteLocalQuery(sqlQuery, transaction)
                    ' clean up blob table on local as well
                    sqlQuery = $"DELETE FROM {BlobTable} WHERE blob_ref_item_id = {ID}"
                    ExecuteLocalQuery(sqlQuery, transaction)
                    'Remove the process from the registration table
                    UnRegisterProcessInstance(transaction)

                Catch ex As Exception
                    logger.Error($"Unable to delete the local process {ID} from the LocalDB.")
                End Try

            End Using

        End Sub

        ''' <summary>
        ''' Deletes all the files that have been attached to the database.
        ''' </summary>
        ''' <remarks>This function should only be called at the very end of a process,
        ''' as these files may be needed from step to step.</remarks>
        Public Sub DeleteFilesAttachedToUDBS()

            For Each afile In _filesAttachedToUDBS

                Try
                    File.Delete(afile)
                Catch ex As Exception
                    'Not critical. Log error and skip the file.
                    logger.Warn(ex, $"Unable to delete the file '{afile}'.")
                    Continue For
                End Try

                Try
                    Dim directoryPath = Path.GetDirectoryName(afile)

                    If Not Directory.GetFiles(directoryPath).Any() Then
                        Directory.Delete(directoryPath)
                    End If
                Catch ex As Exception
                    'Not critical.
                    logger.Debug(ex, ex.Message)
                End Try

            Next

            _filesAttachedToUDBS.Clear()
        End Sub

        ''' <summary>
        ''' Logs the UDBS error in the application log and inserts it in the udbs_error table, with the process instance details:
        ''' Process Type, name, Process ID, UDBS product ID, Unit serial number.
        ''' </summary>
        ''' <param name="ex">Exception raised.</param>
        Private Sub LogErrorInDatabase(ex As Exception)

            'The Process Type is just set to the process name for this class as the type is unidentified as this point.
            DatabaseSupport.LogErrorInDatabase(ex, Process, Stage, mProcessID, ProductNumber, UnitSerialNumber)

        End Sub

        ''' <summary>
        ''' Function copies the process information as well as test data results from network db to local db.
        ''' </summary>
        ''' <returns>Udbs ReturnCode</returns>
        Private Function UpdateLocalDB() As ReturnCodes

            Dim sqlQuery As String
            Dim rsNetworkProcess As New DataTable
            Dim rsLocalProcess As New DataTable
            Dim FieldItem As DataColumn
            Const HINT = "WITH (NOLOCK) "
            Try

                OpenNetworkDB(120)
                sqlQuery = "SELECT * " &
                           "FROM " & Process & "_process " & HINT &
                           "WHERE process_id = " & Format(ID)
                OpenNetworkRecordSet(rsNetworkProcess, sqlQuery)

                If (If(rsNetworkProcess?.Rows?.Count, 0)) > 0 Then
                    ' Open Local Process table
                    sqlQuery = "SELECT * " &
                           "FROM " & Process & "_process " &
                           "WHERE process_id = " & Format(ID)
                    If OpenLocalRecordSet(rsLocalProcess, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Throw New Exception($"Error querying for local process {ID}.")
                    End If

                    Dim newRecord = False
                    ' Check for existing process instance locally
                    If (If(rsLocalProcess?.Rows?.Count, 0)) = 0 Then
                        ' Add a process record to local db
                        Dim dummy = rsLocalProcess.NewRow()
                        dummy("process_id") = ID
                        ' Copy data for each column
                        For Each FieldItem In rsNetworkProcess.Columns
                            dummy(FieldItem.ColumnName) = rsNetworkProcess(0)(FieldItem.ColumnName)
                        Next FieldItem

                        rsLocalProcess.Rows.Add(dummy)
                        newRecord = True
                    Else
                        ' Copy data for each column
                        For Each FieldItem In rsNetworkProcess.Columns
                            rsLocalProcess(0)(FieldItem.ColumnName) = rsNetworkProcess(0)(FieldItem.ColumnName)
                        Next FieldItem
                    End If

                    Dim columnNames =
                                rsLocalProcess.Columns.Cast(Of DataColumn)().[Select](Function(x) x.ColumnName).ToArray()


                    If Not newRecord Then
                        UpdateLocalRecord({"process_id"}, columnNames, rsLocalProcess(0).ItemArray, $"{Process}_process")
                    Else
                        InsertLocalRecord(columnNames, rsLocalProcess(0).ItemArray, $"{Process}_process")
                    End If

                    rsLocalProcess = Nothing

                    ' Copy the network results
                    rsNetworkProcess = Nothing

                    sqlQuery = "SELECT * FROM " & Process & "_result " & HINT &
                                       "WHERE result_process_id = " & CStr(ID)
                    OpenNetworkRecordSet(rsNetworkProcess, sqlQuery)

                    sqlQuery = "SELECT * FROM " & Process & "_result " &
                                       "WHERE result_process_id = " & CStr(ID)
                    If OpenLocalRecordSet(rsLocalProcess, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Throw New Exception($"Error querying for results of local process {ID}")
                    End If

                    If (If(rsNetworkProcess?.Rows?.Count, 0)) > 0 Then

                        columnNames =
                                    rsLocalProcess.Columns.Cast(Of DataColumn)().[Select](Function(x) x.ColumnName).ToArray()

                        For Each drNet As DataRow In rsNetworkProcess.Rows
                            ' Find the result locally..
                            newRecord = False
                            Dim drLocal = rsLocalProcess.AsEnumerable().
                                            FirstOrDefault(
                                                Function(x) _
                                                              KillNullInteger(x("result_itemlistdef_id")) =
                                                              KillNullInteger(drNet("result_itemlistdef_id")))
                            If IsNothing(drLocal) Then
                                newRecord = True
                                drLocal = rsLocalProcess.NewRow()
                                ' Copy the record
                                For Each FieldItem In rsNetworkProcess.Columns
                                    If FieldItem.ColumnName <> "result_id" Then
                                        drLocal(FieldItem.ColumnName) = drNet(FieldItem.ColumnName)
                                    End If
                                Next FieldItem

                            Else
                                ' Copy the record
                                For Each FieldItem In rsNetworkProcess.Columns
                                    If FieldItem.ColumnName <> "result_id" Then
                                        drLocal(FieldItem.ColumnName) = drNet(FieldItem.ColumnName)
                                    End If
                                Next FieldItem
                            End If

                            If Not newRecord Then
                                UpdateLocalRecord({"result_id"}, columnNames, drLocal.ItemArray,
                                                          $"{Process}_result")
                            Else
                                InsertLocalRecord(columnNames.Skip(1).ToArray(),
                                                          drLocal.ItemArray.Skip(1).ToArray(), $"{Process}_result")
                            End If
                        Next
                    End If

                    Return ReturnCodes.UDBS_OP_SUCCESS

                Else
                    ' There is no process with this id
                    LogError(New Exception($"Process with Id: {ID} registered on local database does not have matching record on server."))
                    Return ReturnCodes.UDBS_ERROR
                End If
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsNetworkProcess?.Dispose()
                rsLocalProcess?.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' Establishes the names of the tables. For example, if this object
        ''' is dealing with test data, then "testdata_process"
        ''' </summary>
        ''' <param name="ProcessName"></param>
        Private Sub InitializeTableNames(ProcessName As String)
            _Process = Trim(LCase(ProcessName))
            mProcessTable = GetProcessTableName(Process)
            mResultTable = GetResultTableName(Process)
            mBlobTable = GetBlobTableName(Process)
            mItemListRevisionTable = GetItemListRevisionTableName(Process)
            mItemListDefinitionTable = GetItemListDefinitionTableName(Process)
            mProcessAttributesTable = GetProcessAttributesTableName(Process)
            mProcessAttributesHistoryTable = GetProcessAttributesHistoryTableName(Process)
        End Sub

        Private Shared Function GetProcessTableName(Process As String) As String
            Return Process & "_process"
        End Function

        Private Shared Function GetItemListRevisionTableName(Process As String) As String
            Return Process & "_itemlistrevision"
        End Function

        Private Shared Function GetItemListDefinitionTableName(Process As String) As String
            Return Process & "_itemlistdefinition"
        End Function

        Private Shared Function GetResultTableName(Process As String) As String
            Return Process & "_result"
        End Function

        Private Shared Function GetBlobTableName(Process As String) As String
            Return Process & "_blob"
        End Function

        Private Shared Function GetProcessAttributesTableName(Process As String) As String
            Return Process & "_process_attributes"
        End Function

        Private Shared Function GetProcessAttributesHistoryTableName(Process As String) As String
            Return Process & "_process_attributes_history"
        End Function

        Friend ReadOnly Property ProcessTable As String
            Get
                Return mProcessTable
            End Get
        End Property

        Friend ReadOnly Property ResultTable As String
            Get
                Return mResultTable
            End Get
        End Property

        Friend ReadOnly Property BlobTable As String
            Get
                Return mBlobTable
            End Get
        End Property

        Friend ReadOnly Property ItemListRevisionTable As String
            Get
                Return mItemListRevisionTable
            End Get
        End Property

        Friend ReadOnly Property ItemListDefinitionTable As String
            Get
                Return mItemListDefinitionTable
            End Get
        End Property

        Public Sub New(ProcessName As String)
            InitializeTableNames(ProcessName)
        End Sub

#Region "IDisposable Support"

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    'Try
                    '    UnRegisterProcessInstance(mProcessID)
                    'Catch ex As Exception
                    'End Try
                    mProcessRunning = False
                    mProcessID = 0
                End If
                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region
    End Class
End Namespace
