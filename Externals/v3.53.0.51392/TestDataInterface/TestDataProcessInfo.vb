Imports UdbsInterface.MasterInterface

Namespace TestDataInterface
    Public Class TestDataProcessInfo
        Inherits ProcessInfo

        Private _station As String
        Private _softwareVersion As String
        Private _fullTestNumber As Byte
        Private _blobDataExists As Byte
        Private _archiveState As Byte
        Private _archiveDate As Date

        ''' <summary>
        ''' Private to prevent direct instantiation.
        ''' </summary>
        Private Sub New(serialNumber As String, udbsProductID As String, stage As String)
            MyBase.New(serialNumber, udbsProductID, stage)
        End Sub

        Public ReadOnly Property Station As String
            Get
                Return _station
            End Get
        End Property

        Public ReadOnly Property SoftwareVersion As String
            Get
                Return _softwareVersion
            End Get
        End Property

        Public ReadOnly Property FullTestNumber As Byte
            Get
                Return _fullTestNumber
            End Get
        End Property

        Public ReadOnly Property BlobDataExists As Byte
            Get
                Return _blobDataExists
            End Get
        End Property

        Public ReadOnly Property ArchiveState As Byte
            Get
                Return _archiveState
            End Get
        End Property

        Public ReadOnly Property ArchiveDate As Date
            Get
                Return _archiveDate
            End Get
        End Property

        ''' <summary>
        ''' Checks if the loaded test process is from a different station than the current one.
        ''' </summary>
        ''' <returns>True if the loaded test instance is from another station. False otherwise.</returns>
        <Obsolete("Please use the 'IsProcessOwnedByThisStation' function instead.")>
        Public ReadOnly Property IsProcessFromThisStation As Boolean
            Get
                If (CProcessInstance.GetStationName() = _station) Then

                    Return True
                Else
                    Return False
                End If

            End Get
        End Property

        ''' <inheritdoc/>
        Protected Overrides Sub Load(ByRef resultTable As DataTable)

            If resultTable.Rows.Count = 0 Then

                Throw New ApplicationException($"No test data process information records were found for (id = {ProcessID}) from the network database.")

            ElseIf resultTable.Rows.Count > 1 Then
                Throw New ApplicationException($"Duplicate network DB entries for process ID {ProcessID}.")
            Else

                _unitID = KillNullInteger(resultTable(0)("process_unit_id"))
                _itemListRevisionID = KillNullInteger(resultTable(0)("process_itemlistrev_id"))
                _sequence = KillNullInteger(resultTable(0)("process_sequence"))
                _startDate = KillNullDate(resultTable(0)("process_start_date"))
                _endDate = KillNullDate(resultTable(0)("process_end_date"))
                _status = CProcessInstance.GetProcessStatusEnum(KillNull(resultTable(0)("process_status")))
                _result = UdbsTools.ConvertToResultCode(KillNullInteger(resultTable(0)("process_result")))
                _notes = KillNull(resultTable(0)("process_notes"))
                _employee = KillNull(resultTable(0)("process_employee_number"))
                _activeDuration = CDbl(IIf(Not IsDBNull(resultTable(0)("process_active_duration")), resultTable(0)("process_active_duration"), 0))
                _totalDuration = CDbl(IIf(Not IsDBNull(resultTable(0)("process_total_duration")), resultTable(0)("process_total_duration"), 0))
                _station = KillNull(resultTable(0)("process_station"))
                _softwareVersion = KillNull(resultTable(0)("process_sw_version"))
                _fullTestNumber = KillNullByte(resultTable(0)("process_fulltest_number"))
                _blobDataExists = KillNullByte(resultTable(0)("process_blobdata_exists"))
                _archiveState = KillNullByte(resultTable(0)("process_archive_state"))
                _archiveDate = KillNullDate(resultTable(0)("process_archive_date"))
            End If

        End Sub

        ''' <summary>
        ''' Gets the number of sequences for the specified Product/Unit/Stage
        ''' </summary>
        ''' <param name="productNumber">Udbs product ID</param>
        ''' <param name="serialNumber">unit's serial number.</param>
        ''' <param name="stage">test stage.</param>
        ''' <returns>Latest sequence number.</returns>
        Public Overloads Shared Function GetSequenceCount(productNumber As String,
                                          serialNumber As String,
                                          stage As String) As Integer

            Return ProcessInfo.GetSequenceCount(productNumber, serialNumber, stage, ProcessTypes.testdata)
        End Function

        ''' <summary>
        ''' Gets the process identifier for the specified Product/Unit/Stage/Sequence
        ''' </summary>
        ''' <param name="productNumber">Udbs product ID.</param>
        ''' <param name="serialNumber">Unit's serial number.</param>
        ''' <param name="stage">test stage.</param>
        ''' <param name="sequence">Test sequence number.</param>
        ''' <returns>The unique test process ID.</returns>
        Protected Overloads Shared Function GetProcessID(productNumber As String,
                                      serialNumber As String,
                                      stage As String, Optional sequence As Integer = 0) As Integer

            Return ProcessInfo.GetProcessID(productNumber, serialNumber, stage, ProcessTypes.testdata, sequence)
        End Function

        ''' <summary>
        ''' Gets the test data process information for the specified Product/Unit/Stage/Sequence
        ''' </summary>
        ''' <param name="productNumber">Udbs product ID.</param>
        ''' <param name="serialNumber">Unit's serial number.</param>
        ''' <param name="stage">test stage.</param>
        ''' <param name="sequence">test sequence.</param>
        ''' <returns>A process Info object.</returns>
        Public Overloads Shared Function GetProcessInfo(productNumber As String, serialNumber As String,
                                              stage As String, Optional sequence As Integer = 0) As TestDataProcessInfo


            Dim processID = GetProcessID(productNumber, serialNumber, stage, sequence)

            Return GetProcessInfo(productNumber, serialNumber, stage, sequence, processID)

        End Function

        ''' <summary>
        ''' Gets the test data process information for the specified Product/Unit/Stage/Sequence
        ''' </summary>
        ''' <param name="productNumber">Udbs product ID.</param>
        ''' <param name="serialNumber">Unit's serial number.</param>
        ''' <param name="stage">test stage.</param>
        ''' <param name="sequence">test sequence.</param>
        ''' <param name="processID">process ID </param>
        ''' <returns>A process Info object.</returns>
        Public Overloads Shared Function GetProcessInfo(productNumber As String, serialNumber As String,
                                              stage As String, sequence As Integer, processID As Integer) As TestDataProcessInfo

            Dim sqlQuery = "SELECT * FROM " & "testdata_process" & " with(nolock) WHERE process_id = " & CStr(processID)

            Dim resultTable = New DataTable
            If QueryNetworkDB(sqlQuery, resultTable) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New UdbsTestException($"Error retrieving process information for (id = {processID}) from the network database.")
            End If

            Dim processInfo = New TestDataProcessInfo(serialNumber, productNumber, stage) With {._processID = processID, ._sequence = sequence}

            processInfo.Load(resultTable)

            Return processInfo

        End Function

        ''' <summary>
        ''' Checks whether this process is from this station.
        ''' Compares both the machine name and the stationID to the station name obtained from UDBS.
        ''' </summary>
        ''' <param name="stationID">This is the 'StationID' system environment variable set by the application.
        ''' This optional parameter can be ignored for non-Fractal applications.</param>
        ''' <returns>True if the process belongs to this station. False otherwise.</returns>
        Public Function IsProcessOwnedByThisStation(Optional stationID As String = "") As Boolean

            Return stationID = _station Or CProcessInstance.GetStationName() = _station
        End Function

    End Class
End Namespace

