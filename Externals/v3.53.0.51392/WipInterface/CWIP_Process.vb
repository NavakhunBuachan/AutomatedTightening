Option Explicit On
Option Compare Binary
Option Infer On
Option Strict On

Imports UdbsInterface.MasterInterface


Namespace WipInterface
    ' Result Enumeration
    Public Enum WIPResultCodes
        WIP_ERROR = -1000000
        WIP_PASS = 10
        WIP_SKIP = 5
        WIP_FAIL = -30
        WIP_INCOMPLETE = -10
    End Enum

    Public Enum LockStatus_Enum
        READ_ONLY = 0
        READ_WRITE = 1
    End Enum

    Public Class CWIP_Process
        Implements IDisposable

        Private Const ClsName = "CWIP_Process"

        Private ReadOnly mUt As New CWIP_Utility

        'class variables
        Private mLockStatus As LockStatus_Enum
        Private mLoaded As Boolean

        ' Results Collection
        Private mResults As Dictionary(Of Integer, CWIP_Result) = Nothing
        ' Itemlist Object
        Private mItemlist As CWIP_ItemList

        'object data Properties
        Private mfamily_name As String
        Private mproduct_id As Integer
        Private mproduct_number As String
        Private mproduct_descriptor As String
        Private munit_id As Integer
        Private munit_serial_number As String
        Private mitemlistrev_id As Integer
        Private mitemlistrev_stage As String
        Private mitemlistrev_revision As Integer
        Private mprocess_id As Integer
        Private mprocess_sequence As Integer
        Private mprocess_start_date As Date
        Private mprocess_end_date As Date
        Private mprocess_status As String
        Private mprocess_result As Integer
        Private mprocess_notes As String
        Private mprocess_work_order As String
        Private mprocess_tracking_number As String
        Private mprocess_employee_number As String
        Private mprocess_active_duration As Double
        Private mprocess_total_duration As Double
        Private mprocess_excursion_count As Integer
        Private mprocess_unit_info As String
        Private mprocess_active_step As String
        Private mprocess_blobdata_exists As Integer
        Private mprocess_locked_by As String

        '************************************************************************************************************
        ' PROPERTIES
        '************************************************************************************************************
        'Results Collection
        Public ReadOnly Property Results As Dictionary(Of Integer, CWIP_Result)
            Get
                If Not mLoaded Then Return Nothing
                If mResults Is Nothing Then
                    ' Properties getters should not throw exceptions; ignoring return code.
                    ' An error message is already logged inside LoadResultsCollection()
                    LoadResultsCollection()
                End If

                Return mResults
            End Get
        End Property

        ''' <remarks>This is used by the TED Tools.</remarks>
        Public ReadOnly Property Itemlist As CWIP_ItemList
            Get
                If Not mLoaded Then Return Nothing
                If mItemlist Is Nothing Then
                    mItemlist = New CWIP_ItemList
                End If
                If Not mItemlist.Loaded Then
                    mItemlist.LoadItemListByID(mitemlistrev_id)
                End If
                Itemlist = mItemlist
            End Get
        End Property

        '**********************************************************************
        '* READ-ONLY PROPERTIES
        '**********************************************************************
        Public ReadOnly Property Locked As LockStatus_Enum
            Get
                Return mLockStatus
            End Get
        End Property

        'Data Properties
        Public ReadOnly Property Family As String
            Get
                Return mfamily_name
            End Get
        End Property

        Public ReadOnly Property ProductID As Integer
            Get
                Return mproduct_id
            End Get
        End Property

        Public ReadOnly Property ProductNumber As String
            Get
                Return mproduct_number
            End Get
        End Property

        Public ReadOnly Property ProductDescriptor As String
            Get
                Return mproduct_descriptor
            End Get
        End Property

        Public ReadOnly Property UnitID As Integer
            Get
                Return munit_id
            End Get
        End Property

        Public ReadOnly Property SerialNumber As String
            Get
                Return munit_serial_number
            End Get
        End Property

        Public ReadOnly Property ItemlistRevID As Integer
            Get
                Return mitemlistrev_id
            End Get
        End Property

        Public ReadOnly Property Stage As String
            Get
                Return mitemlistrev_stage
            End Get
        End Property

        Public ReadOnly Property ItemlistRevision As Integer
            Get
                Return mitemlistrev_revision
            End Get
        End Property

        Public ReadOnly Property ID As Integer
            Get
                Return mprocess_id
            End Get
        End Property

        '**********************************************************************
        '* READ/WRITE PROPERTIES
        '**********************************************************************

        Public Property StartDate As Date
            Get
                Return mprocess_start_date
            End Get
            Set
                If StoreProcessField("process_start_date", Value, False) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_start_date = Value
                End If
            End Set
        End Property

        Public Property EndDate As Date
            Get
                Return mprocess_end_date
            End Get
            Set
                If StoreProcessField("process_end_date", Value, False) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_end_date = Value
                End If
            End Set
        End Property

        Public Property Status As String
            Get
                Return mprocess_status
            End Get
            Set
                If StoreProcessField("process_status", Value, False) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_status = Value
                End If
            End Set
        End Property

        Public Property Result As Integer
            Get
                Return mprocess_result
            End Get
            Set
                If StoreProcessField("process_result", Value, True) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_result = Value
                End If
            End Set
        End Property

        Public Property Notes As String
            Get
                Return mprocess_notes
            End Get
            Set
                If StoreProcessField("process_notes", Value, False) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_notes = Value
                End If
            End Set
        End Property

        ''' <remarks>This is used by the TED Tools.</remarks>
        Public Property WorkOrder As String
            Get
                Return mprocess_work_order
            End Get
            Set
                If StoreProcessField("process_work_order", Value, False) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_work_order = Value
                End If
            End Set
        End Property

        Public Property TrackingNumber As String
            Get
                Return mprocess_tracking_number
            End Get
            Set
                If StoreProcessField("process_tracking_number", Value, False) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_tracking_number = Value
                End If
            End Set
        End Property

        Public Property EmployeeNumber As String
            Get
                Return mprocess_employee_number
            End Get
            Set
                If StoreProcessField("process_employee_number", Value, False) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_employee_number = Value
                End If
            End Set
        End Property

        Public Property ActiveDuration As Double
            Get
                Return mprocess_active_duration
            End Get
            Set
                If StoreProcessField("process_active_duration", Value, True) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_active_duration = Value
                End If
            End Set
        End Property

        Public Property TotalDuration As Double
            Get
                Return mprocess_total_duration
            End Get
            Set
                If StoreProcessField("process_total_duration", Value, True) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_total_duration = Value
                End If
            End Set
        End Property

        Public Property ExcursionCount As Integer
            Get
                Return mprocess_excursion_count
            End Get
            Set
                If StoreProcessField("process_excursion_count", Value, True) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_excursion_count = Value
                End If
            End Set
        End Property

        Public Property UnitInfo As String
            Get
                Return mprocess_unit_info
            End Get
            Set
                If StoreProcessField("process_unit_info", Value, False) = ReturnCodes.UDBS_OP_SUCCESS Then
                    mprocess_unit_info = Value
                End If
            End Set
        End Property

        Public ReadOnly Property ActiveStep As String
            Get
                Return mprocess_active_step
            End Get
        End Property

        ''' <summary>
        ''' Set the active WIP step.
        ''' Process must first be locked for read-write access.
        ''' </summary>
        ''' <param name="value">The name of the new active step.</param>
        ''' <param name="transaction">The ongoing database transaction.</param>
        Private Sub SetActiveStep(value As String, transaction As ITransactionScope)
            If StoreProcessField("process_active_step", value, False, transaction) = ReturnCodes.UDBS_OP_SUCCESS Then
                mprocess_active_step = value
            End If
        End Sub

        Public ReadOnly Property LockedBy As String
            Get
                Return mprocess_locked_by
            End Get
        End Property

        ''' <summary>
        ''' Specify from where a WIP process is locked.
        ''' </summary>
        ''' <param name="value">
        ''' The station name where this WIP process is being locked from.
        ''' Use 'Nothing' when unlocking the process.
        ''' </param>
        ''' <param name="transaction">The ongoing database transaction.</param>
        Private Sub SetLockedBy(value As String, transaction As ITransactionScope)
            If StoreProcessField("process_locked_by", value, False, transaction) = ReturnCodes.UDBS_OP_SUCCESS Then
                mprocess_locked_by = value
            End If
        End Sub

        '**********************************************************************
        '* METHODS
        '**********************************************************************

        ''' <summary>
        ''' Unload the current process, then reload a stored process.
        ''' </summary>
        ''' <param name="SerialNumber">
        ''' The serial number of the unit to load.
        ''' Note: This method assumes that serial numbers are unique.
        ''' </param>
        ''' <param name="LockStatus">
        ''' The desired lock status. Whether the caller want to load the process in 'read-only' mode, or in 'read-write' mode.
        ''' </param>
        ''' <returns>
        ''' The outcome of the operation.
        ''' </returns>
        Public Function LoadActiveProcess(
                SerialNumber As String,
                Optional ByVal LockStatus As LockStatus_Enum = LockStatus_Enum.READ_ONLY) As ReturnCodes
            Using transaction As ITransactionScope = BeginNetworkTransaction()
                Try
                    ' Get the process id from the unit.
                    Using tmpUnit As New CWIP_Unit
                        Dim returnCode = tmpUnit.LoadActiveUnit(SerialNumber, mprocess_id)
                        If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then
                            Return returnCode
                        End If
                    End Using

                    ' Load the process that was found.
                    Return LoadProcessByID(mprocess_id, LockStatus, transaction)
                Catch ex As Exception
                    LogErrorInDatabase(ex)
                    Return ReturnCodes.UDBS_ERROR
                End Try
            End Using
        End Function

        ' Not used...
        ' Candidate for removal.
        ' Although, this one gets access to earlier sequences,
        ' and provide support for non-unique serial numbers.
        Friend Function LoadProcess(ProductNumber As String,
                                    SerialNumber As String,
                                    Stage As String,
                                    Sequence As Integer,
                                    Optional ByVal LockStatus As LockStatus_Enum = LockStatus_Enum.READ_ONLY) _
            As ReturnCodes
            Dim rsTemp As New DataTable
            Dim strSQL As String
            Dim returnCode As ReturnCodes

            Try
                UnloadProcess(Nothing)

                'check the new process is not already closed
                strSQL = "SELECT process_id " &
                         "FROM product with(nolock) , unit with(nolock) , WIP_process, WIP_itemlistrevision  with(nolock) " &
                         "WHERE unit_product_id = product_id " &
                         "AND process_unit_id = unit_id " &
                         "AND process_itemlistrev_id = itemlistrev_id " &
                         "AND product_number = '" & ProductNumber & "' " &
                         "AND unit_serial_number = '" & SerialNumber & "' " &
                         "AND itemlistrev_stage = '" & Stage & "' " &
                         "AND process_sequence = " & Sequence

                returnCode = QueryNetworkDB(strSQL, rsTemp)
                If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode

                If (If(rsTemp?.Rows?.Count, 0)) <> 1 Then
                    'none or too many processes found
                    LogError(New Exception("No unique process instance was found."))
                    Return ReturnCodes.UDBS_OP_FAIL

                End If

                'this is good - one process found
                Return LoadProcessByID(KillNullInteger(rsTemp(0)("process_id")), LockStatus)

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Load a process in memory from its ID (DB key).
        ''' </summary>
        ''' <param name="ProcessID">The ID of the WIP process to load.</param>
        ''' <param name="LockStatus">Whether to load the process 'read-only' or 'read-write'.</param>
        ''' <returns>The outcome of the operation.</returns>
        Public Function LoadProcessByID(ProcessID As Integer,
                                        Optional ByVal LockStatus As LockStatus_Enum = LockStatus_Enum.READ_ONLY) _
                As ReturnCodes
            Using transaction As ITransactionScope = BeginNetworkTransaction()
                Return LoadProcessByID(ProcessID, LockStatus, transaction)
            End Using
        End Function

        ''' <summary>
        ''' Load a process in memory from its ID (DB key).
        ''' </summary>
        ''' <param name="ProcessID">The ID of the WIP process to load.</param>
        ''' <param name="LockStatus">Whether to load the process 'read-only' or 'read-write'.</param>
        ''' <param name="transaction">The ongoing database transaction.</param>
        ''' <returns>The outcome of the operation.</returns>
        Friend Function LoadProcessByID(ProcessID As Integer,
                                        LockStatus As LockStatus_Enum,
                                        transaction As ITransactionScope) As ReturnCodes
            Try
                UnloadProcess(transaction)

                Dim strSQL =
                    "SELECT family_name, product_id, product_number, product_descriptor, unit_id, unit_serial_number, " &
                    "itemlistrev_revision, itemlistrev_stage, WIP_process.* " &
                    "FROM family with(nolock) , product with(nolock) , unit with(nolock) , WIP_process, WIP_itemlistrevision  with(nolock) " &
                    "WHERE product_family_id = family_id AND unit_product_id = product_id " &
                    "AND process_unit_id = unit_id AND process_itemlistrev_id = itemlistrev_id " &
                    "AND process_id = " & ProcessID

                ' If function is called by UnitSummary, then set the last argument as the
                ' tempory connection string to retrieve data in corresponding database
                Dim rsTemp As New DataTable
                OpenNetworkRecordSet(rsTemp, strSQL, transaction)

                If (If(rsTemp?.Rows?.Count, 0)) = 1 Then
                    mfamily_name = KillNull(rsTemp(0)("family_name"))
                    mproduct_id = KillNullInteger(rsTemp(0)("product_id"))
                    mproduct_number = KillNull(rsTemp(0)("product_number"))
                    mproduct_descriptor = KillNull(rsTemp(0)("product_descriptor"))
                    munit_id = KillNullInteger(rsTemp(0)("unit_id"))
                    munit_serial_number = KillNull(rsTemp(0)("unit_serial_number"))
                    mitemlistrev_id = KillNullInteger(rsTemp(0)("process_itemlistrev_id"))
                    mitemlistrev_stage = KillNull(rsTemp(0)("itemlistrev_stage"))
                    mitemlistrev_revision = KillNullInteger(rsTemp(0)("itemlistrev_revision"))
                    mprocess_id = KillNullInteger(rsTemp(0)("process_id"))
                    mprocess_sequence = KillNullInteger(rsTemp(0)("process_sequence"))
                    mprocess_start_date = KillNullDate(rsTemp(0)("process_start_date"))
                    mprocess_end_date = KillNullDate(rsTemp(0)("process_end_date"))
                    mprocess_status = KillNull(rsTemp(0)("process_status"))
                    mprocess_result = KillNullInteger(rsTemp(0)("process_result"))
                    mprocess_notes = KillNull(rsTemp(0)("process_notes"))
                    mprocess_work_order = KillNull(rsTemp(0)("process_work_order"))
                    mprocess_tracking_number = KillNull(rsTemp(0)("process_tracking_number"))
                    mprocess_employee_number = KillNull(rsTemp(0)("process_employee_number"))
                    mprocess_active_duration = KillNullDouble(rsTemp(0)("process_active_duration"))
                    mprocess_total_duration = KillNullDouble(rsTemp(0)("process_total_duration"))
                    mprocess_excursion_count = KillNullInteger(rsTemp(0)("process_excursion_count"))
                    mprocess_unit_info = KillNull(rsTemp(0)("process_unit_info"))
                    mprocess_active_step = KillNull(rsTemp(0)("process_active_step"))
                    mprocess_blobdata_exists = KillNullInteger(rsTemp(0)("process_blobdata_exists"))
                    mprocess_locked_by = KillNull(rsTemp(0)("process_locked_by"))

                    'no matter what happens, the process is now loaded
                    mLoaded = True

                    'the caller wants to load the process in read/write mode.
                    'lock the result for read/write (if required)
                    If LockStatus = LockStatus_Enum.READ_WRITE Then
                        If LockProcess(transaction) <> ReturnCodes.UDBS_OP_SUCCESS Then
                            logger.Error("Failed to lock WIP process.")
                            Return ReturnCodes.UDBS_ERROR
                        End If
                    End If

                    Return ReturnCodes.UDBS_OP_SUCCESS
                Else
                    LogError(New Exception("Couldn't load process."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        '**********************************************************************
        '* PRIVATE Methods
        '**********************************************************************

        ''' <summary>
        ''' Unlock the WIP process.
        ''' </summary>
        ''' <param name="transaction">The ongoing database transaction.</param>
        Private Sub UnloadProcess(transaction As ITransactionScope)
            ' Unlock the process if it's locked.
            If mLockStatus = LockStatus_Enum.READ_WRITE Then
                If UnlockProcess("", "", transaction) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    logger.Error("Error unlocking WIP process.")
                End If
            End If

            mLockStatus = LockStatus_Enum.READ_ONLY

            mResults = Nothing
            mItemlist = Nothing

            mfamily_name = Nothing
            mproduct_id = Nothing
            mproduct_number = Nothing
            mproduct_descriptor = Nothing
            munit_id = Nothing
            munit_serial_number = Nothing
            mitemlistrev_id = Nothing
            mitemlistrev_stage = Nothing
            mitemlistrev_revision = Nothing
            mprocess_id = Nothing
            mprocess_sequence = Nothing
            mprocess_start_date = Nothing
            mprocess_end_date = Nothing
            mprocess_status = Nothing
            mprocess_result = Nothing
            mprocess_notes = Nothing
            mprocess_work_order = Nothing
            mprocess_tracking_number = Nothing
            mprocess_employee_number = Nothing
            mprocess_active_duration = Nothing
            mprocess_total_duration = Nothing
            mprocess_excursion_count = Nothing
            mprocess_unit_info = Nothing
            mprocess_active_step = Nothing
            mprocess_blobdata_exists = Nothing
            mprocess_locked_by = Nothing

            mLoaded = False
        End Sub

        ''' <summary>
        ''' Loads the results for this WIP process.
        ''' </summary>
        ''' <param name="transaction">
        ''' The ongoing database transaction.
        ''' If not provided (or null), a new transaction is created.
        ''' </param>
        ''' <returns>The outcome of the operation.</returns>
        Friend Function LoadResultsCollection(Optional transaction As ITransactionScope = Nothing) As ReturnCodes
            If transaction Is Nothing Then
                transaction = BeginNetworkTransaction()
                Using transaction
                    Return LoadResultsCollection(transaction)
                End Using
            End If

            Try

                If Not mLoaded Then
                    LogError(New Exception("Cannot load results collection when process is not loaded."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                Dim strSQL = "SELECT * FROM WIP_result with(nolock), WIP_itemlistdefinition  with(nolock) " &
                         "WHERE result_itemlistdef_id = itemlistdef_id " &
                         "AND result_process_id = " & mprocess_id & " " &
                         "ORDER BY result_step_number ASC"
                Dim rsTemp As New DataTable
                OpenNetworkRecordSet(rsTemp, strSQL, transaction)

                mResults = New Dictionary(Of Integer, CWIP_Result)

                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                    For Each dr As DataRow In rsTemp.Rows
                        Dim tmpItem = New CWIP_Item(KillNullInteger(dr("itemlistdef_id")),
                                                    KillNullInteger(dr("itemlistdef_itemnumber")),
                                                    KillNull(dr("itemlistdef_itemname")),
                                                    KillNull(dr("itemlistdef_descriptor")),
                                                    KillNull(dr("itemlistdef_description")),
                                                    KillNullInteger(dr("itemlistdef_required_step")),
                                                    KillNull(dr("itemlistdef_processname")),
                                                    KillNull(dr("itemlistdef_stagename")),
                                                    KillNull(dr("itemlistdef_role")),
                                                    KillNull(dr("itemlistdef_pass_routing")),
                                                    KillNull(dr("itemlistdef_fail_routing")),
                                                    KillNullInteger(dr("itemlistdef_automated_process")),
                                                    KillNullInteger(dr("itemlistdef_oracle_routing")),
                                                    KillNullInteger(dr("itemlistdef_blobdata_exists")))

                        Dim tmpResult = New CWIP_Result(tmpItem, KillNullLong(dr("result_id")),
                                                        KillNullInteger(dr("result_process_id")),
                                                        KillNullInteger(dr("result_itemlistdef_id")),
                                                        KillNullInteger(dr("result_step_number")),
                                                        KillNull(dr("result_authorized_by")),
                                                        KillNullDate(dr("result_start_date")),
                                                        KillNull(dr("result_employee_number")),
                                                        KillNull(dr("result_station")),
                                                        KillNullInteger(dr("result_udbs_process_id")),
                                                        KillNullDate(dr("result_end_date")),
                                                        CType(KillNullInteger(dr("result_passflag")), WIPResultCodes),
                                                        KillNullDouble(dr("result_inactive_duration")),
                                                        KillNullDouble(dr("result_active_duration")),
                                                        KillNull(dr("result_wip_notes")),
                                                        KillNullInteger(dr("result_blobdata_exists")))

                        mResults.Add(KillNullInteger(dr("result_step_number")), tmpResult)
                    Next
                Else
                    ' This is not an error.
                    ' A brand-new unit with no WIP results is a valid state to be in.
                    logger.Debug("No results found for this process.")
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Stores some process data into the database.
        ''' </summary>
        ''' <param name="name">The name of the field to store.</param>
        ''' <param name="value">The value to store.</param>
        ''' <param name="isNumeric">Whether or not this field is a numeric value.</param>
        ''' <param name="transaction">
        ''' The ongoing database transaction.
        ''' If not provided (or null), a new transaction is created.
        ''' </param>
        ''' <returns>The outcome of the operation.</returns>
        Private Function StoreProcessField(
                name As String,
                value As Object,
                isNumeric As Boolean,
                Optional transaction As ITransactionScope = Nothing) As ReturnCodes

            If transaction Is Nothing Then
                transaction = BeginNetworkTransaction()
                Using transaction
                    Return StoreProcessField(name, value, isNumeric, transaction)
                End Using
            End If

            Try
                If mLockStatus = LockStatus_Enum.READ_ONLY Then
                    LogError(New Exception("Process is locked. Cannot update process data."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                ' Convert Thai date to common era if this is a date object.
                If value IsNot Nothing AndAlso value.GetType() = GetType(DateTime) Then
                    value = CUtility.ConvertThaiToCommonEra(CType(value, DateTime))
                End If

                If isNumeric Then
                    value = KillNull(value)
                Else
                    If Len(KillNull(value)) = 0 Then
                        value = Nothing
                    End If
                End If

                If Not UpdateNetworkRecord(
                        New String() {"process_id"},
                        New String() {"process_id", name},
                        New Object() {mprocess_id, value},
                        "WIP_process", transaction) Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ' Candidate for removal.
        Private Function GetUnitInfoElement(Token As String, TokenVal As String) As ReturnCodes
            'Retrieves an element of UnitInfo
            Dim TokenFound As Boolean
            Dim TokenNumber As Integer
            Dim ArrBuf() As String
            Dim i As Integer

            Try
                If Not mLoaded Then
                    LogError(New Exception("No Process Loaded."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                'find the token
                TokenFound = False
                ArrBuf = Split(Itemlist.UnitInfo, vbCrLf, Len(Itemlist.UnitInfo))
                For i = 0 To UBound(ArrBuf)
                    If ArrBuf(i) = Token Then
                        TokenNumber = i
                        TokenFound = True
                        Exit For
                    End If
                Next i

                If Not TokenFound Then Return ReturnCodes.UDBS_OP_FAIL

                'now find the value
                TokenVal = ""
                ArrBuf = Split(UnitInfo, vbCrLf, Len(Itemlist.UnitInfo))
                For i = 0 To UBound(ArrBuf)
                    If i = TokenNumber Then
                        TokenVal = ArrBuf(i)
                        Exit For
                    End If
                Next i

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ' Candidate for removal.
        Private Function SetUnitInfoElement(Token As String, TokenVal As String) As ReturnCodes
            'Retrieves an element of UnitInfo
            Dim TokenFound As Boolean
            Dim TokenNumber As Integer
            Dim ArrBuf() As String
            Dim buf As String
            Dim i As Integer

            Try

                If Not mLoaded Then
                    LogError(New Exception("No Process Loaded."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                'find the token
                TokenFound = False
                ArrBuf = Split(Itemlist.UnitInfo, vbCrLf, Len(Itemlist.UnitInfo))
                For i = 0 To UBound(ArrBuf)
                    If ArrBuf(i) = Token Then
                        TokenNumber = i
                        TokenFound = True
                        Exit For
                    End If
                Next i

                If Not TokenFound Then Return ReturnCodes.UDBS_OP_FAIL

                'now set the value
                ArrBuf = Split(UnitInfo, vbCrLf, Len(UnitInfo))
                buf = ""
                For i = 0 To UBound(ArrBuf)
                    If i = TokenNumber Then
                        buf = buf & vbCrLf & TokenVal
                    Else
                        buf = buf & vbCrLf & ArrBuf(i)
                    End If
                Next i

                UnitInfo = Mid(buf, 3)

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        '**********************************************************************
        '* PUBLIC PROCESS METHODS
        '**********************************************************************

        ''' <summary>Begin a new WIP process.</summary>
        ''' <param name="FamilyName">The product family.</param>
        ''' <param name="ProductNumber">The product number.</param>
        ''' <param name="SerialNumber">The unit's serial number.</param>
        ''' <param name="Stage">The WIP stage.</param>
        ''' <param name="EmployeeNumber">The ID of the employee starting the new WIP process.</param>
        ''' <remarks>This is used by the TED Tools.</remarks>
        ''' <returns>The outcome of the operation.</returns>
        Public Function Begin_Process(FamilyName As String,
                                      ProductNumber As String,
                                      SerialNumber As String,
                                      Stage As String,
                                      EmployeeNumber As String) As ReturnCodes

            Const fncName = ClsName & ":Begin_Process"
            Dim rsTemp As New DataTable
            Dim returnCode As ReturnCodes

            Using transaction As ITransactionScope = BeginNetworkTransaction()
                Using tmpUnit As New CWIP_Unit()
                    Using tmpILR As New CWIP_ItemList()
                        Try
                            UnloadProcess(transaction)

                            'first get some information about the unit.
                            returnCode = tmpUnit.LoadUnit(SerialNumber, FamilyName, ProductNumber)
                            If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then
                                Return returnCode
                            End If
                            Dim unitID = tmpUnit.UnitID

                            'look for any running WIP processes on this unit.
                            Dim strSQL = "SELECT product_number, unit_id FROM product with(nolock), unit with(nolock), WIP_process " &
                                 "WHERE product_id=unit_product_id and unit_id=process_unit_id " &
                                 "AND unit_serial_number='" & SerialNumber & "' " &
                                 "AND (process_status = 'PAUSED' OR process_status = 'IN PROCESS') " &
                                 "ORDER BY process_id DESC"
                            OpenNetworkRecordSet(rsTemp, strSQL, transaction)

                            If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                                ' The process is already loaded/running.
                                LogError(New Exception("Process is running under " & KillNull(rsTemp(0)("product_number")) & ". Cannot start a new process."))
                                Return ReturnCodes.UDBS_OP_FAIL
                            End If

                            'if everything is OK, initiate the new process record.
                            'get the itemlist id - we don't need to populate the itemlist yet.
                            returnCode = tmpILR.LoadItemList(tmpUnit.ProductNumber, tmpUnit.ProductRelease, Stage, 0)
                            AssertOperationSucceeded(returnCode)

                            'get the start time.
                            Dim ServerTime As Date
                            returnCode = CUtility.Utility_GetServerTime(ServerTime)
                            AssertOperationSucceeded(returnCode)

                            'determine how many sequences have already ocurred for this stage.
                            strSQL = "SELECT process_id, process_sequence FROM WIP_process, WIP_itemlistrevision " &
                                 "WHERE process_itemlistrev_id = itemlistrev_id " &
                                 "AND process_unit_id = " & unitID & " " &
                                 "AND itemlistrev_stage = '" & Stage & "' " &
                                 "ORDER BY process_sequence"
                            OpenNetworkRecordSet(rsTemp, strSQL, transaction)
                            Dim processSequence = (If(rsTemp?.Rows?.Count, 0)) + 1

                            'we're now ready to add the record.
                            Dim processID = InsertNetworkRecord(
                                New String() {"process_unit_id", "process_itemlistrev_id", "process_sequence", "process_start_date", "process_end_date", "process_status", "process_employee_number"},
                                New Object() {unitID, tmpILR.ID, processSequence, DBDateFormat(ServerTime), DBDateFormat(ServerTime), "PAUSED", EmployeeNumber}, "WIP_process", transaction, "process_id")

                            'reload the process object into memory, make sure it is locked.
                            returnCode = LoadProcessByID(processID, LockStatus_Enum.READ_WRITE, transaction)
                            If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode

                            'create the first WIP Step so the unit is waiting to begin.
                            Return CreateStep(Itemlist.GetItemByNumber(1).Name, "auto", transaction)

                        Catch ex As Exception
                            LogErrorInDatabase(ex)
                            Return ReturnCodes.UDBS_ERROR
                        End Try
                    End Using
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Terminate the WIP process.
        ''' </summary>
        ''' <param name="transaction">The ongoing database transaction.</param>
        ''' <returns>The outcome of the operation.</returns>
        Private Function End_Process(transaction As ITransactionScope) As ReturnCodes
            Dim CompletedStep = ""
            Dim PrcActiveTime As Double, PrcInactiveTime As Double
            Dim ServerTime As Date

            Try
                If mLoaded = False Or mLockStatus = LockStatus_Enum.READ_ONLY Then
                    LogError(New Exception("Cannot complete this action. Process may not be loaded or may be read-only."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                'look for any running WIP processes on this unit
                If mprocess_status = "IN PROCESS" Then
                    If FinishStep(WIPResultCodes.WIP_INCOMPLETE, "Aborted by dll", transaction) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Return ReturnCodes.UDBS_ERROR
                    End If
                Else
                    'There should be no active record but if there is, skip that step
                    'WIPStep_ByPass
                End If

                'COMPLETE And IN PROCESS steps will not be allowed because they can't be locked
                Dim ProcessStatus As String
                Dim ProcessResult = EvaluateProcess(CompletedStep)
                If ProcessResult > 0 Then
                    ProcessStatus = "COMPLETE"
                Else
                    ProcessStatus = "TERMINATED"
                End If

                'compute the active and inactive durations
                PrcActiveTime = 0
                PrcInactiveTime = 0
                For Each Res In Results.Values
                    PrcActiveTime = PrcActiveTime + If(Double.IsNaN(Res.ActiveDuration), 0, Res.ActiveDuration)
                    PrcInactiveTime = PrcInactiveTime + If(Double.IsNaN(Res.InactiveDuration), 0, Res.InactiveDuration)
                Next Res

                'get the End Time
                AssertOperationSucceeded(CUtility.Utility_GetServerTime(ServerTime))
                PrcInactiveTime = (ServerTime - mprocess_start_date).TotalMinutes - PrcActiveTime

                'we're now ready to add the record... build the string
                If Not UpdateNetworkRecord(
                        New String() {"process_id"},
                        New String() {"process_id", "process_end_date", "process_status", "process_result", "process_active_duration", "process_total_duration", "process_locked_by"},
                        New Object() {mprocess_id, ServerTime, ProcessStatus, ProcessResult, PrcActiveTime, PrcInactiveTime, Nothing},
                        "WIP_process", transaction) Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                'update the process data
                mLockStatus = LockStatus_Enum.READ_ONLY
                mprocess_end_date = ServerTime
                mprocess_status = ProcessStatus
                mprocess_result = ProcessResult
                mprocess_active_duration = PrcActiveTime
                mprocess_total_duration = PrcInactiveTime
                mprocess_locked_by = Nothing

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Checks the whole WIP process to make sure that everything has been completed (in order!)
        ''' If the process is not complete, the function returns WIP_INCOMPLETE.
        ''' </summary>
        ''' <param name="CompletedStep">(Out) The last completed step.</param>
        ''' <returns>The evaluation outcome.</returns>
        ''' <remarks>
        ''' If all steps are complete, the next step is "" and the function returns:
        '''   -30     if the process is incomplete and has an unclosed WIP excursion
        '''   -10     if the process is incomplete but so far, so good
        '''     5     if any manual overrides occurred
        '''    10     for a clean pass
        ''' </remarks>
        Private Function EvaluateProcess(ByRef CompletedStep As String) As WIPResultCodes
            Dim i As Integer, j As Integer, ConfirmedResult As Integer, ForcedPass As Boolean
            Dim StepFound As Boolean
            Dim activeStep As CWIP_Item
            Dim ResultCount As Integer

            Try
                CompletedStep = ""
                'make sure there is a valid process loaded
                If Not mLoaded Then
                    LogError(New Exception("Cannot complete this action. Process is not loaded."))
                    Return WIPResultCodes.WIP_ERROR
                End If

                'if no results have been stored, the CompleteStep is blank
                If Results.Count < 1 Then
                    Return WIPResultCodes.WIP_INCOMPLETE
                End If

                activeStep = Results(Results.Count).Item

                'now go through the itemlist (in ascending order) to verify each item in the results
                ResultCount = Results.Count
                'check if the last result has actually been stored
                If Results(ResultCount).Passflag = 0 Then ResultCount = ResultCount - 1
                ConfirmedResult = 0
                ForcedPass = False
                For i = 1 To Itemlist.Items.Count
                    'only check required steps
                    If Itemlist.GetItemByNumber(i).RequiredStep > 0 Then
                        StepFound = False
                        'now loop through the results until you find a step that matches this step
                        For j = ConfirmedResult + 1 To ResultCount
                            If Results(j).Item.Number = Itemlist.GetItemByNumber(i).Number Then
                                StepFound = True
                                ConfirmedResult = j
                            End If
                        Next j

                        'we now know the latest result for this required step
                        If StepFound Then
                            'check the passflag which identifies that the step has been performed and passed
                            Dim StepResult = Results(ConfirmedResult).Passflag
                            If StepResult > 0 Then
                                If StepResult = WIPResultCodes.WIP_SKIP Then ForcedPass = True
                                'if it passed then the process is still good
                                'this is, so far, the last completed step
                                CompletedStep = Itemlist.GetItemByNumber(i).Name
                            ElseIf Itemlist.GetItemByNumber(i).RequiredStep >= 2 Then _
                                ' RequiredStep was = 2, updated to >= 2 to take care of the new option "3: no skipping"
                                'step was found but failed - we abandon the evaluation here, process is incomplete
                                'if the activestep is beyond this point, the process is failing
                                If activeStep.Number > i Then
                                    Return WIPResultCodes.WIP_FAIL
                                Else
                                    Return WIPResultCodes.WIP_INCOMPLETE
                                End If
                                Exit Function
                            End If
                        Else
                            'step was not found - we abandon the evaluation here, process is incomplete
                            If activeStep.Number > i Then
                                Return WIPResultCodes.WIP_FAIL
                            Else
                                Return WIPResultCodes.WIP_INCOMPLETE
                            End If
                        End If
                    End If

                Next i

                'if we got here, this means that all required steps were found and passed
                If ForcedPass Then
                    Return WIPResultCodes.WIP_SKIP
                Else
                    Return WIPResultCodes.WIP_PASS
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return WIPResultCodes.WIP_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Unlocks the process.
        ''' </summary>
        ''' <param name="UserName">The username of the user unlocking the process.</param>
        ''' <param name="Pwd">The password of the user performing the operation.</param>
        ''' <returns>The outcome of the operation.</returns>
        ''' <remarks>This is used by the TED Tools.</remarks>
        Public Function Unlock_Process(Optional ByVal UserName As String = "", Optional ByVal Pwd As String = "") As ReturnCodes
            Using transaction As ITransactionScope = BeginNetworkTransaction()
                Return UnlockProcess(UserName, Pwd, transaction)
            End Using
        End Function

        ''' <summary>
        ''' Unlocks the process.
        ''' </summary>
        ''' <param name="UserName">The username of the user unlocking the process.</param>
        ''' <param name="Pwd">The password of the user performing the operation.</param>
        ''' <param name="transaction">The ongoing database transaction.</param>
        ''' <returns>The outcome of the operation.</returns>
        Friend Function UnlockProcess(UserName As String, Pwd As String, transaction As ITransactionScope) As ReturnCodes
            Dim strSQL As String
            Dim rsTemp As New DataTable

            Try
                If Not mLoaded Then
                    LogError(New Exception("Cannot complete this action. Process is not loaded."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                If mLockStatus = LockStatus_Enum.READ_WRITE Then
                    'this is easy, the process is locked by this object - so unlock it
                    SetLockedBy(Nothing, transaction)
                    mLockStatus = LockStatus_Enum.READ_ONLY
                    Return ReturnCodes.UDBS_OP_SUCCESS
                End If

                'this is trickier, the process is locked by another object - so we'll have to free it up

                'check if the process is actually locked in the database
                strSQL = "SELECT process_locked_by FROM WIP_process WHERE process_id = " & mprocess_id
                OpenNetworkRecordSet(rsTemp, strSQL, transaction)

                'a computer name will exist if the station is locked
                mprocess_locked_by = KillNull(rsTemp(0)("process_locked_by"))

                If mprocess_locked_by = "" Then
                    'object already unlocked
                    Return ReturnCodes.UDBS_OP_SUCCESS
                Else
                    'object is locked by another station
                    'this requires user validation
                    If Not mUt.CheckUserPrivileges(UserName, Pwd, False, "Administrators", "Engineers") Then
                        LogError(New Exception("Incorrect username or password. You cannot unlock the WIP process."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    Else
                        'we need to assign ourselves permission to unlock
                        mLockStatus = LockStatus_Enum.READ_WRITE
                        SetLockedBy(Nothing, transaction)
                        mLockStatus = LockStatus_Enum.READ_ONLY
                    End If
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Locks the process for read-write access.
        ''' </summary>
        ''' <param name="transaction">The ongoing database transaction.</param>
        ''' <returns>The outcome of the operation.</returns>
        Private Function LockProcess(transaction As ITransactionScope) As ReturnCodes

            Try
                If Not mLoaded Then Return ReturnCodes.UDBS_OP_FAIL

                If mLockStatus = LockStatus_Enum.READ_WRITE Then
                    'object already locked
                    Return ReturnCodes.UDBS_OP_SUCCESS
                Else
                    'cannot lock a closed process
                    If mprocess_status = "COMPLETE" Or mprocess_status = "TERMINATED" Then
                        LogError(New Exception("Cannot lock process. Process is " & mprocess_status))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If

                    'check if the process is already locked in the database
                    Dim rsTemp As New DataTable
                    Dim strSQL = "SELECT process_locked_by FROM WIP_process WHERE process_id = " & mprocess_id
                    OpenNetworkRecordSet(rsTemp, strSQL, transaction)

                    'a computer name will exist if the station is locked
                    mprocess_locked_by = KillNull(rsTemp(0)("process_locked_by"))

                    Dim ComputerName = ""
                    If CUtility.Utility_GetStationName(ComputerName) <> ReturnCodes.UDBS_OP_SUCCESS _
                                Or ComputerName = "" Then
                        LogError(New Exception("Station name not found. Cannot lock WIP process."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If

                    If String.IsNullOrEmpty(mprocess_locked_by) Then
                        'process is free, go ahead and lock it
                        'get the current computer name so we can store that

                        'now store the computer name - this will secure the process
                        mLockStatus = LockStatus_Enum.READ_WRITE
                        SetLockedBy(ComputerName, transaction)

                        'now just verify that it worked
                        OpenNetworkRecordSet(rsTemp, strSQL, transaction)
                        If KillNull(rsTemp(0)("process_locked_by")) <> ComputerName Then
                            mLockStatus = LockStatus_Enum.READ_ONLY
                        End If
                        Return ReturnCodes.UDBS_OP_SUCCESS
                    ElseIf ComputerName = mprocess_locked_by Then
                        'process is already locked by us.
                        mLockStatus = LockStatus_Enum.READ_WRITE
                        Return ReturnCodes.UDBS_OP_SUCCESS
                    Else
                        'process is already locked by another computer (or thread) - can't lock it
                        LogError(New Exception("Process is already locked by " & mprocess_locked_by))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Returns the next recommended WIP step.
        ''' NextStep is empty if there is no valid next step (i.e. last required step is done).
        ''' </summary>
        ''' <param name="CurrentStep">The current WIP step.</param>
        ''' <param name="CurrentResult">The outcome of the WIP process' evaluation.</param>
        ''' <param name="NextStep">(Out) The recommended next step.</param>
        ''' <returns>The outcome of the operation.</returns>
        Public Function GetNextRecommendedStep(CurrentStep As String, CurrentResult As Integer, ByRef NextStep As String) _
            As ReturnCodes

            Dim ThisStep As CWIP_Item
            Dim TempStep = ""
            Dim TempResult As WIPResultCodes

            Try
                NextStep = ""

                'if there are no results yet (probably won't happen) choose the first step
                If CurrentStep = "" Then
                    NextStep = Itemlist.GetItemByNumber(1).Name
                    Return ReturnCodes.UDBS_OP_SUCCESS
                End If

                'the item object can tell us where to go
                ThisStep = Itemlist.Items(CurrentStep)
                NextStep = ThisStep.Name

                If CurrentResult <= 0 And Not ThisStep.FailRouting = "" Then
                    'if there is a fail routing for the active step, choose that one if it failed
                    NextStep = ThisStep.FailRouting
                    Return ReturnCodes.UDBS_OP_SUCCESS
                ElseIf CurrentResult > 0 And Not ThisStep.PassRouting = "" Then
                    'if there is a pass routing for the active step, choose that one if it passed
                    NextStep = ThisStep.PassRouting
                    Return ReturnCodes.UDBS_OP_SUCCESS
                Else
                    'there is no routing info so we're on our own - look for the next step
                    TempResult = EvaluateProcess(TempStep)
                    Return GetNextRequiredStep(TempStep, TempResult, NextStep)
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Returns the next required WIP step.
        ''' NextStep is empty if there is no valid next step (i.e. last required step is done).
        ''' </summary>
        ''' <param name="CurrentStep">The current WIP step.</param>
        ''' <param name="CurrentResult">The outcome of the WIP process' evaluation.</param>
        ''' <param name="NextStep">(Out) The recommended next step.</param>
        ''' <returns>The outcome of the operation.</returns>
        Private Function GetNextRequiredStep(CurrentStep As String, CurrentResult As Integer, ByRef NextStep As String) _
            As ReturnCodes

            Dim i As Integer
            Dim ThisStep As CWIP_Item

            Try
                NextStep = ""

                'if there are no results yet (probably won't happen) : choose the first step
                'this will handle a Evaluate_Process that returns CompletedStep = ""
                If CurrentStep = "" Then
                    NextStep = Itemlist.GetItemByNumber(1).Name
                    Return ReturnCodes.UDBS_OP_SUCCESS
                End If

                'the item object tells us what to do
                ThisStep = Itemlist.Items(CurrentStep)
                NextStep = ThisStep.Name

                'look for the next required step
                If CurrentResult > 0 Or ThisStep.RequiredStep > 0 Then
                    For i = ThisStep.Number + 1 To Itemlist.Items.Count
                        If Itemlist.GetItemByNumber(i).RequiredStep > 0 Then
                            NextStep = Itemlist.GetItemByNumber(i).Name
                            Exit For
                        End If
                    Next i
                Else
                    'we'll just stay where we are
                    NextStep = CurrentStep
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        '**********************************************************************
        '* FRIEND WIP STEP METHODS
        '**********************************************************************

        ''' <summary>
        ''' Creates a new WIP step without actually starting the step
        ''' This step acts as a placeholder for the active WIP step
        ''' </summary>
        ''' <param name="StepName">The name of the WIP step to create.</param>
        ''' <param name="Authority">
        ''' The username of the user creating this step.
        ''' In the case of an automatic WIP step routing, use "auto".
        ''' </param>
        ''' <param name="transaction">The ongoing database transaction.</param>
        ''' <returns>The outcome of the operation.</returns>
        Private Function CreateStep(StepName As String, Authority As String, transaction As ITransactionScope) As ReturnCodes

            Dim StepNumber As Integer
            Dim ItemID As Integer

            Try
                'make sure there is a valid process loaded
                If mLoaded = False Or mLockStatus = LockStatus_Enum.READ_ONLY Then
                    LogError(New Exception("Cannot complete this action. Process may not be loaded or may be read-only."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                If mprocess_status <> "PAUSED" Then
                    LogError(New Exception("This unit is not ready for a WIP operation."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                'Now retrieve the WIP step number
                StepNumber = Results.Count + 1

                ItemID = Itemlist.Items(StepName).ID

                'it is OK to add the step, put the new info in the result table
                Dim resultId = InsertNetworkRecord(
                        New String() {"result_process_id", "result_itemlistdef_id", "result_authorized_by", "result_step_number"},
                        New Object() {mprocess_id, ItemID, Authority, StepNumber}, "WIP_result", transaction, "result_id")

                'update the active_step field, put the new info in the result table
                If Not UpdateNetworkRecord(
                        New String() {"process_id"},
                        New String() {"process_id", "process_active_step"},
                        New Object() {mprocess_id, Itemlist.Items(StepName).Name},
                        "WIP_process", transaction) Then
                    LogError(New Exception("Failed to set the WIP process' active step."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                SetActiveStep(StepName, transaction)

                Return LoadResultsCollection(transaction)

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Begins and ends a WIP Step storing only the result
        ''' There must be a current, unstarted process
        ''' </summary>
        ''' <param name="transaction">
        ''' The ongoing database transaction.
        ''' If none is provided (null) a new one will get created.
        ''' </param>
        ''' <param name="notes">The notes to append to the WIP step.</param>
        ''' <returns>The outcome of the operation.</returns>
        Friend Function SkipStep(Optional transaction As ITransactionScope = Nothing,
                                  Optional notes As String = "") As ReturnCodes
            If transaction Is Nothing Then
                transaction = BeginNetworkTransaction()
                Using transaction
                    Return SkipStep(transaction, notes)
                End Using
            End If

            Try
                notes = CUtility.Utility_ConvertStringToASCIICondenseInvalidCharacters(notes)
                'make sure there is a valid process loaded
                If mLoaded = False Or mLockStatus = LockStatus_Enum.READ_ONLY Then
                    LogError(New Exception("Cannot complete this action. Process may not be loaded or may be read-only."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                If mprocess_status <> "PAUSED" Then
                    LogError(New Exception("This unit is not ready to start a WIP step."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                'get the current result (i.e the last one)
                Dim ActiveResult = Results(Results.Count)

                'this is a sanity check - the active step should match the latest result record
                ' if not, the routing screwed up
                If Not (ActiveResult.Item.Name = mprocess_active_step) Then
                    LogError(New Exception("A database error occurred:" & vbCrLf & "The process_active_step field does not " & "match the last result record." & vbCrLf & vbCrLf &
                             "Please inform the software developer."))
                    'proceed using the result table as the true indicator
                End If

                If Not UpdateNetworkRecord(
                        New String() {"result_id"},
                        New String() {"result_id", "result_start_date", "result_end_date", "result_employee_number", "result_station", "result_udbs_process_id", "result_active_duration", "result_inactive_duration", "result_passflag", "result_wip_notes"},
                        New Object() {ActiveResult.ID, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, WIPResultCodes.WIP_SKIP, notes & " Skipped due to re-routing." & vbCrLf},
                        "WIP_result", transaction) Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                'results have changed so reload them
                Return LoadResultsCollection(transaction)

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Takes the unstarted WIP step and changes it into a new one without affecting the result
        ''' There must be a current, unstarted process.
        ''' </summary>
        ''' <param name="StepName">The new active step name.</param>
        ''' <param name="UserName">The username of the person making this change.</param>
        ''' <param name="Note">The note to append to the step.</param>
        ''' <param name="transaction">
        ''' The ongoing database transaction.
        ''' If none is provided (null) a new one will get created.
        ''' </param>
        ''' <returns>The outcome of the operation.</returns>
        Friend Function ChangeActiveStep(
                StepName As String,
                UserName As String,
                Note As String,
                Optional transaction As ITransactionScope = Nothing) As ReturnCodes

            If transaction Is Nothing Then
                transaction = BeginNetworkTransaction()
                Using transaction
                    Return ChangeActiveStep(StepName, UserName, Note, transaction)
                End Using
            End If

            Try
                'make sure there is a valid process loaded
                If mLoaded = False Or mLockStatus = LockStatus_Enum.READ_ONLY Then
                    LogError(New Exception("Cannot complete this action. Process may not be loaded or may be read-only."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                If mprocess_status <> "PAUSED" Then
                    LogError(New Exception("This unit does not have an unstarted WIP step."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                'get the current result (i.e the last one)
                Dim ActiveResult = Results(Results.Count)
                If ActiveResult.Passflag <> 0 Then
                    LogError(New Exception("This unit does not have an unstarted WIP step."))
                    Return CreateStep(StepName, UserName, transaction)
                End If

                'this is a sanity check - the active step should match the latest result record
                ' if not, the routing screwed up
                If Not (ActiveResult.Item.Name = mprocess_active_step) Then
                    LogError(New Exception("A database error occurred: The process_active_step field does not match the last result record." &
                                           "Please inform the software developer."))
                    'proceed using the result table as the true indicator
                End If

                If Not UpdateNetworkRecord(
                        New String() {"result_id"},
                        New String() {"result_id", "result_itemlistdef_id", "result_authorized_by", "result_wip_notes"},
                        New Object() {ActiveResult.ID, Itemlist.Items(StepName).ID, UserName, mprocess_notes & Note & vbCrLf},
                        "WIP_result", transaction) Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                'results have changed so reload them
                AssertOperationSucceeded(LoadResultsCollection(transaction))

                mprocess_active_step = StepName
                If Not UpdateNetworkRecord(
                        New String() {"process_id"},
                        New String() {"process_id", "process_active_step"},
                        New Object() {mprocess_id, mprocess_active_step},
                        "WIP_process", transaction) Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        '**********************************************************************
        '* PUBLIC WIP STEP METHODS
        '**********************************************************************

        ''' <remarks>This is used by the TED Tools.</remarks>
        Public Function AutoReRoute(StepName As String,
                                    RoutingType As String,
                                    UserName As String,
                                    Optional ByVal Pwd As String = "") _
            As ReturnCodes

            Dim SafeStep = "", ReqStep = ""
            Dim ReqStepNumber As Integer, NewStepNumber As Integer
            Dim localActiveStep As CWIP_Result
            Dim i As Integer, MaxIterations As Integer
            Dim Permit As Boolean
            Dim RetCode = ReturnCodes.UDBS_OP_INC

            Const AddNote = "Auto reroute after rework"

            Using transaction As ITransactionScope = BeginNetworkTransaction()
                Try
                    'make sure there is a valid process loaded
                    If mLoaded = False Or mLockStatus = LockStatus_Enum.READ_ONLY Then
                        RetCode = ReturnCodes.UDBS_OP_FAIL
                        LogError(New Exception("Cannot complete this action. Process may not be loaded or may be read-only."))
                        Return RetCode
                    End If

                    'make sure the WIP process is paused
                    If Not mprocess_status = "PAUSED" Then
                        LogError(New Exception("This unit is currently IN PROCESS."))
                        RetCode = ReturnCodes.UDBS_OP_FAIL
                        Return RetCode
                    End If

                    NewStepNumber = Itemlist.Items(StepName).Number

                    'get the last step that was completed successfully (i.e. "safe")
                    Dim WIPreturnCode = EvaluateProcess(SafeStep)
                    If WIPreturnCode <> WIPResultCodes.WIP_PASS Then Throw New UDBSException($"Evaluate_Process for '{SafeStep}' failed with WIP result code: {WIPreturnCode}")

                    'now get the required next step based on the 'safe' one
                    AssertOperationSucceeded(GetNextRequiredStep(SafeStep, WIPResultCodes.WIP_PASS, ReqStep))
                    ReqStepNumber = Itemlist.Items(ReqStep).Number

                    'the activestep is the one that is currently waiting to be done
                    localActiveStep = Results(Results.Count)

                    'if new step comes after the next required step, then the reroute is an 'exception'
                    'this means it may need to 'jump over' some required steps
                    'if the user has the right password etc., they can insert dummy steps to force a pass
                    If Itemlist.GetItemByNumber(NewStepNumber).RequiredStep > 0 And NewStepNumber > ReqStepNumber Then
                        ''RoutingType = InputBox("You are attempting to route the unit forward by skipping steps. " & vbCrLf & _
                        '    "The unit should go to '" & Itemlist.Items(ReqStep).Descriptor & "'" & vbCrLf & vbCrLf & _
                        '    "Please choose one of the following:" & vbCrLf & _
                        '    " (1) A 'safe' re-route. This allows you to temporarily skip steps but you must return and do them later." & vbCrLf & _
                        '    " (2) An 'exception' re-route. Permanently skip WIP steps (requires engineering permission)", , "1")
                        If Val(RoutingType) = 2 Then 'unsafe re-route
                            'this is an error trap of sorts - the open loop below could rapidly fill the DB with crap
                            MaxIterations = NewStepNumber - ReqStepNumber

                            'if the user has the right privileges, continue with the unsafe routing
                            Permit = mUt.CheckUserPrivileges(UserName, Pwd, True, "Administrators", "Engineers")
                            If Not Permit Then
                                If MsgBox("Insufficient privileges to perform an 'exception' routing." & vbCrLf & vbCrLf &
                                      "Proceed with 'safe routing?", vbYesNoCancel) = vbYes Then RoutingType = "1"
                            Else
                                'make sure to start with the last 'safe' point
                                If localActiveStep.Item.Number < ReqStepNumber Then
                                    RetCode = ChangeActiveStep(ReqStep, UserName, AddNote, transaction)
                                    If RetCode <> ReturnCodes.UDBS_OP_SUCCESS Then
                                        Return ReturnCodes.UDBS_ERROR
                                    End If

                                    localActiveStep = Results(Results.Count)
                                End If

                                i = 0
                                'create dummy steps to enable unit to evaluate to a pass
                                Do
                                    'make the result record a 'skipped' result
                                    RetCode = SkipStep(transaction)
                                    If RetCode <> ReturnCodes.UDBS_OP_SUCCESS Then
                                        Return RetCode
                                    End If

                                    'now evaluate to find the next required step
                                    AssertOperationSucceeded(GetNextRequiredStep(localActiveStep.Item.Name, WIPResultCodes.WIP_SKIP, ReqStep))
                                    ReqStepNumber = Itemlist.Items(ReqStep).Number

                                    If ReqStepNumber >= Itemlist.Items(StepName).Number Then
                                        'we have now filled in all the 'dummy steps'
                                        RetCode = CreateStep(StepName, UserName, transaction)
                                        Return RetCode
                                    Else
                                        'otherwise create that next required step
                                        RetCode = CreateStep(ReqStep, UserName, transaction)
                                    End If
                                    localActiveStep = Results(Results.Count)
                                    If RetCode <> ReturnCodes.UDBS_OP_SUCCESS Or i > MaxIterations Then
                                        LogError(New Exception("Problem occurred trying to create dummy WIP steps."))
                                        RetCode = ReturnCodes.UDBS_ERROR
                                        Return RetCode
                                    End If
                                    i = i + 1
                                Loop
                            End If
                        End If
                        If Val(RoutingType) = 1 Then
                            'check the privileges
                            Permit = mUt.CheckUserPrivileges(UserName, Pwd, True, "Administrators", "Engineers", "Technicians")
                            If Not Permit Then
                                Throw New ApplicationException("You do not have forward re-routing privileges.")
                            Else
                                RetCode = ChangeActiveStep(StepName, UserName, AddNote, transaction)
                            End If
                        End If
                    Else
                        RetCode = ChangeActiveStep(StepName, UserName, AddNote, transaction)
                    End If

                    Return RetCode
                Catch ex As Exception
                    LogErrorInDatabase(ex)
                    Return ReturnCodes.UDBS_ERROR
                End Try
            End Using
        End Function

        Friend Function ReRoute(StepName As String,
                                AddNote As String,
                                UserName As String,
                                Optional ByVal Pwd As String = "") _
            As ReturnCodes


            Dim SafeStep = "", ReqStep = ""
            Dim ReqStepNumber As Integer, NewStepNumber As Integer
            Dim localActiveStep As CWIP_Result
            Dim i As Integer, MaxIterations As Integer
            Dim RoutingType As String
            Dim Permit As Boolean

            Dim tmpStr As String
            Dim RetCode As ReturnCodes

            Using transaction As ITransactionScope = BeginNetworkTransaction()
                Try
                    'make sure there is a valid process loaded
                    If mLoaded = False Or mLockStatus = LockStatus_Enum.READ_ONLY Then
                        RetCode = ReturnCodes.UDBS_OP_FAIL
                        LogError(New Exception("Cannot complete this action. Process may not be loaded or may be read-only."))
                        Return RetCode
                    End If

                    'make sure the WIP process is paused
                    If Not mprocess_status = "PAUSED" Then
                        LogError(New Exception("This unit is currently IN PROCESS."))
                        RetCode = ReturnCodes.UDBS_OP_FAIL
                        Return RetCode
                    End If

                    NewStepNumber = Itemlist.Items(StepName).Number

                    'get the last step that was completed successfully (i.e. "safe")
                    Dim WIPreturnCode = EvaluateProcess(SafeStep)
                    If WIPreturnCode <> WIPResultCodes.WIP_PASS Then Throw New UDBSException($"Evaluate_Process for '{SafeStep}' failed with WIP result code: {WIPreturnCode}")

                    'now get the required next step based on the 'safe' one
                    AssertOperationSucceeded(GetNextRequiredStep(SafeStep, WIPResultCodes.WIP_PASS, ReqStep))
                    ReqStepNumber = Itemlist.Items(ReqStep).Number

                    'the activestep is the one that is currently waiting to be done
                    localActiveStep = Results(Results.Count)

                    'if new step comes after the next required step, then the reroute is an 'exception'
                    'this means it may need to 'jump over' some required steps
                    'if the user has the right password etc., they can insert dummy steps to force a pass
                    If Itemlist.GetItemByNumber(NewStepNumber).RequiredStep > 0 And NewStepNumber > ReqStepNumber Then
                        ' check if any NO SKIPPING skips between next reqd step and new step
                        ' if found, only safe reroute is allowed
                        tmpStr = ""
                        For i = ReqStepNumber To NewStepNumber - 1
                            If Itemlist.GetItemByNumber(i).RequiredStep = 3 Then
                                If tmpStr <> "" Then tmpStr = tmpStr & ","
                                tmpStr = tmpStr & Itemlist.GetItemByNumber(i).Descriptor
                            End If
                        Next i
                        If tmpStr = "" Then
                            RoutingType = InputBox("You are attempting to route the unit forward by skipping steps. " & vbCrLf &
                                               "The unit should go to '" & Itemlist.Items(ReqStep).Descriptor & "'" & vbCrLf &
                                               vbCrLf &
                                               "Please choose one of the following:" & vbCrLf &
                                               " (1) A 'safe' re-route. This allows you to temporarily skip steps but you must return and do them later." &
                                               vbCrLf &
                                               " (2) An 'exception' re-route. Permanently skip WIP steps (requires engineering permission).",
                                               , "1")
                        Else
                            RoutingType = InputBox("You are attempting to route the unit forward by skipping steps. " & vbCrLf &
                                               "The unit should go to '" & Itemlist.Items(ReqStep).Descriptor & "'" & vbCrLf &
                                               vbCrLf &
                                               tmpStr & " is(are) NO SKIPPING step(s)! Please choose one of the following:" &
                                               vbCrLf &
                                               " (0) Cancel reroute." & vbCrLf &
                                               " (1) A 'safe' re-route. This allows you to temporarily skip steps but you must return and do them later.",
                                               , "1")
                            If Val(RoutingType) <> 1 Then RoutingType = "0"
                        End If
                        If Val(RoutingType) = 2 Then 'unsafe re-route
                            'this is an error trap of sorts - the open loop below could rapidly fill the DB with crap
                            MaxIterations = NewStepNumber - ReqStepNumber

                            'if the user has the right privileges, continue with the unsafe routing
                            Permit = mUt.CheckUserPrivileges(UserName, Pwd, True, "Administrators", "Engineers")
                            If Not Permit Then
                                If MsgBox("Insufficient privileges to perform an 'exception' routing." & vbCrLf & vbCrLf &
                                      "Proceed with 'safe routing?", vbYesNoCancel) = vbYes Then RoutingType = "1"
                            Else
                                'make sure to start with the last 'safe' point
                                If localActiveStep.Item.Number < ReqStepNumber Then
                                    RetCode = ChangeActiveStep(ReqStep, UserName, AddNote, transaction)
                                    If RetCode <> ReturnCodes.UDBS_OP_SUCCESS Then
                                        Return RetCode
                                    End If

                                    localActiveStep = Results(Results.Count)
                                End If

                                i = 0
                                'create dummy steps to enable unit to evaluate to a pass
                                Do
                                    'make the result record a 'skipped' result
                                    RetCode = SkipStep(transaction, AddNote)
                                    If RetCode <> ReturnCodes.UDBS_OP_SUCCESS Then
                                        Return RetCode
                                    End If

                                    'now evaluate to find the next required step
                                    If GetNextRequiredStep(localActiveStep.Item.Name, WIPResultCodes.WIP_SKIP, ReqStep) <> ReturnCodes.UDBS_OP_SUCCESS Then
                                        Return ReturnCodes.UDBS_ERROR
                                    End If

                                    ReqStepNumber = Itemlist.Items(ReqStep).Number

                                    If ReqStepNumber >= Itemlist.Items(StepName).Number Then
                                        'we have now filled in all the 'dummy steps'
                                        RetCode = CreateStep(StepName, UserName, transaction)
                                        Return RetCode
                                    Else
                                        'otherwise create that next required step
                                        RetCode = CreateStep(ReqStep, UserName, transaction)
                                    End If

                                    localActiveStep = Results(Results.Count)
                                    If RetCode <> ReturnCodes.UDBS_OP_SUCCESS Or i > MaxIterations Then
                                        LogError(New Exception("Problem occurred trying to create dummy WIP steps."))
                                        Return RetCode
                                    End If

                                    i += 1
                                Loop
                            End If
                        End If

                        If Val(RoutingType) = 1 Then
                            'check the privileges
                            Permit = mUt.CheckUserPrivileges(UserName, Pwd, True, "Administrators", "Engineers", "Technicians")
                            If Not Permit Then
                                Throw New ApplicationException("You do not have forward re-routing privileges.")
                            Else
                                RetCode = ChangeActiveStep(StepName, UserName, AddNote, transaction)
                            End If
                        End If
                    Else
                        RetCode = ChangeActiveStep(StepName, UserName, AddNote, transaction)
                    End If

                    Return RetCode
                Catch ex As Exception
                    LogErrorInDatabase(ex)
                    Return ReturnCodes.UDBS_ERROR
                End Try
            End Using

        End Function

        ''' <summary>
        ''' Finish the current WIP step.
        ''' </summary>
        ''' <param name="UDBSResult">
        ''' Interface error: This parameter is not 'By Ref'. It is an input parameter.
        ''' TODO: Change this parameter to a normal input parameter. Note that this will
        ''' a breaking change!
        ''' TODO: Also, this should be of type WipResultCode.
        ''' </param>
        ''' <param name="WIPNote">The note to append to this step.</param>
        ''' <returns>The outcome of the operation.</returns>
        Public Function Finish_Step(
                ByRef UDBSResult As Integer,
                Optional ByVal WIPNote As String = "") As ReturnCodes
            Using transaction As ITransactionScope = BeginNetworkTransaction()
                Return FinishStep(UDBSResult, WIPNote, transaction)
            End Using
        End Function

        ''' <summary>
        ''' Finish the current WIP step.
        ''' </summary>
        ''' <param name="UDBSResult">The result of this WIP step.</param>
        ''' <param name="WIPNote">The note to append to this step.</param>
        ''' <param name="transaction">An ongoing database transaction.</param>
        ''' <returns>The outcome of the operation.</returns>
        Friend Function FinishStep(
                UDBSResult As Integer,
                WIPNote As String,
                transaction As ITransactionScope) As ReturnCodes

            ' Terminates a currently running WIP Step
            ' There must be a current, started process
            Dim ActiveResult As CWIP_Result
            Dim ServerTime As Date, StepActiveTime As Double
            Dim ProcessResult As WIPResultCodes
            Dim NewStep = "", ReqStep = ""
            Dim returnCode As ReturnCodes

            Try
                'make sure there is a valid process loaded
                If mLoaded = False Or mLockStatus = LockStatus_Enum.READ_ONLY Then
                    LogError(New Exception("Cannot complete this action. Process may not be loaded or may be read-only."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                If mprocess_status <> "IN PROCESS" Then
                    LogError(New Exception("This unit is not currently IN PROCESS."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                'get the current result (i.e the last one)
                ActiveResult = Results(Results.Count)

                'this is a sanity check - the active step should match the latest result record
                ' if not, the routing screwed up
                If Not (ActiveResult.Item.Name = mprocess_active_step) Then
                    LogError(New Exception("A database error occurred: The process_active_step field does not match the last result record." &
                                           "Please inform the software developer."))
                    'proceed using the result table as the true indicator
                End If

                'get the current time
                CUtility.Utility_GetServerTime(ServerTime)


                Dim columnNames = New List(Of String)({"result_id", "result_end_date", "result_passflag", "result_wip_notes"})
                Dim columnValues = New List(Of Object)({ActiveResult.ID, ServerTime, UDBSResult, ActiveResult.WIPNotes & WIPNote & vbCrLf})
                If Not ActiveResult.StartDate = Nothing Then
                    StepActiveTime = (ServerTime - ActiveResult.StartDate).TotalMinutes
                    columnNames.Add("result_active_duration")
                    columnValues.Add(StepActiveTime)
                End If

                If Not UpdateNetworkRecord(
                    New String() {"result_id"},
                    columnNames.ToArray(), columnValues.ToArray(),
                    "WIP_result", transaction) Then
                    Return returnCode
                End If

                'reload the results
                AssertOperationSucceeded(LoadResultsCollection(transaction))

                'update the active_step field, put the new info in the result table
                If Not UpdateNetworkRecord(
                        New String() {"process_id"},
                        New String() {"process_id", "process_end_date", "process_status"},
                        New Object() {mprocess_id, ServerTime, "PAUSED"},
                        "WIP_process", transaction) Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                mprocess_end_date = ServerTime
                mprocess_status = "PAUSED"

                ActiveResult = Results(Results.Count)

                'evaluate the WIP process to figure out where to route the unit
                ProcessResult = EvaluateProcess(NewStep)
                If ProcessResult > 0 Then
                    Return End_Process(transaction)
                End If

                'if the process result is failing (i.e. we've gone past the errant process) we should try to remedy the situation
                If ProcessResult = WIPResultCodes.WIP_FAIL Then
                    'check the next required step
                    AssertOperationSucceeded(GetNextRequiredStep(NewStep, UDBSResult, ReqStep))
                    'check the next recommended step
                    AssertOperationSucceeded(GetNextRecommendedStep(ActiveResult.Item.Name, UDBSResult, NewStep))
                    'we would like to trust the recommended step but
                    'if it's going to send us to some future 'required' step, we should stop and back up to the one that's failing
                    If Itemlist.Items(NewStep).RequiredStep > 0 _
                            And Itemlist.Items(NewStep).Number > Itemlist.Items(ReqStep).Number Then
                        NewStep = ReqStep
                    End If
                Else
                    If UDBSResult > 0 And Itemlist.Items(ActiveResult.Item.Name).PassRouting = "" Then
                        'there is no routing info for this passed step - so look for the next required step
                        AssertOperationSucceeded(GetNextRequiredStep(NewStep, UDBSResult, ReqStep))
                        NewStep = ReqStep
                    Else
                        AssertOperationSucceeded(GetNextRecommendedStep(ActiveResult.Item.Name, UDBSResult, NewStep))
                    End If
                End If

                Return CreateStep(NewStep, "auto", transaction)
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Begins a WIP Step.
        ''' There must be a current, unstarted process.
        ''' </summary>
        Public Function Start_Step(EmployeeNumber As String,
                                   StationName As String,
                                   Optional ByVal UDBSProcessID As Integer = -1) As ReturnCodes
            Dim ServerTime As Date
            Dim StepInactiveTime As Double


            Using transaction As ITransactionScope = BeginNetworkTransaction()
                Try
                    'make sure there is a valid process loaded
                    If mLoaded = False Or mLockStatus = LockStatus_Enum.READ_ONLY Then
                        LogError(New Exception($"Cannot complete this action. Process {UDBSProcessID} may not be loaded or may be read-only."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If

                    If mprocess_status <> "PAUSED" Then
                        LogError(New Exception($"This unit with Process Id {UDBSProcessID} is not ready to start a WIP step."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If

                    'get the current result (i.e the last one)
                    Dim ActiveResult = Results(Results.Count)

                    'this is a sanity check - the active step should match the latest result record
                    ' if not, the routing screwed up
                    If Not (ActiveResult.Item.Name = mprocess_active_step) Then
                        LogError(New Exception("A database error occurred:" & vbCrLf & "The process_active_step field does not " &
                                        "match the last result record." & vbCrLf & vbCrLf & "Please inform the software developer."))
                        'proceed using the result table as the true indicator
                    End If

                    'get the current time
                    CUtility.Utility_GetServerTime(ServerTime)

                    'update the new record, put the new info in the result table
                    Dim columnNames As New List(Of String)({"result_id", "result_start_date", "result_employee_number", "result_station"})
                    Dim columnValues As New List(Of Object)({ActiveResult.ID, ServerTime, EmployeeNumber, StationName})
                    If UDBSProcessID <> -1 Then
                        columnNames.Add("result_udbs_process_id")
                        columnValues.Add(UDBSProcessID)
                    End If
                    If Not mprocess_end_date = Nothing Then
                        StepInactiveTime = (ServerTime - mprocess_end_date).TotalMinutes
                        columnNames.Add("result_inactive_duration")
                        columnValues.Add(StepInactiveTime)
                    End If
                    If Not UpdateNetworkRecord(
                            New String() {"result_id"},
                            columnNames.ToArray(), columnValues.ToArray(), "WIP_result", transaction) Then
                        Return ReturnCodes.UDBS_ERROR
                    End If

                    'results have changed so reload them
                    AssertOperationSucceeded(LoadResultsCollection(transaction))

                    'update the active_step field, put the new info in the result table
                    mprocess_status = "IN PROCESS"
                    mprocess_active_step = ActiveResult.Item.Name

                    If Not UpdateNetworkRecord(
                            New String() {"process_id"},
                            New String() {"process_id", "process_active_step", "process_status"},
                            New Object() {mprocess_id, mprocess_active_step, "IN PROCESS"},
                            "WIP_process", transaction) Then
                        Return ReturnCodes.UDBS_ERROR
                    End If

                    Return ReturnCodes.UDBS_OP_SUCCESS

                Catch ex As Exception
                    LogErrorInDatabase(ex)
                    Return ReturnCodes.UDBS_ERROR
                End Try
            End Using

        End Function

        ''' <summary>
        ''' Logs the UDBS error in the application log and inserts it in the udbs_error table, with the process instance details:
        ''' Process Type, name, Process ID, UDBS product ID, Unit serial number.
        ''' </summary>
        ''' <param name="ex">Exception raised.</param>
        Private Sub LogErrorInDatabase(ex As Exception)

            DatabaseSupport.LogErrorInDatabase(ex, PROCESS, Stage, ID, ProductNumber, SerialNumber)

        End Sub

#Region "IDisposable Support"

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' Destroys collection when this class is terminated
                    If mLockStatus = LockStatus_Enum.READ_WRITE Then
                        If Unlock_Process() <> ReturnCodes.UDBS_OP_SUCCESS Then
                            logger.Error("Error unlocking WIP process.")
                        End If
                    End If

                    mUt.Dispose()
                    CloseNetworkDB()
                    mItemlist = Nothing
                    mResults = Nothing
                    mLoaded = False
                End If
            End If
            Me.disposedValue = True
        End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region
    End Class
End Namespace
