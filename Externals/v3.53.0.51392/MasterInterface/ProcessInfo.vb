
Namespace MasterInterface

    ''' <summary>
    ''' This is the base abstract class for a process info. Must be implemented for testdata, wip and kitting process info.
    ''' </summary>
    Public MustInherit Class ProcessInfo

        Public Enum ProcessTypes
            testdata
            wip
            kitting
        End Enum

        Protected _serialNumber As String
        Protected _ubdsProductID As String
        Protected _stage As String

        Protected _processID As Integer
        Protected _unitID As Integer
        Protected _itemListRevisionID As Integer
        Protected _sequence As Integer
        Protected _startDate As Date
        Protected _endDate As Date
        Protected _status As UdbsProcessStatus
        Protected _result As TestDataInterface.ResultCodes
        Protected _notes As String
        Protected _employee As String
        Protected _activeDuration As Double
        Protected _totalDuration As Double

        ''' <summary>
        ''' Constructor for a ProcessInfo object.
        ''' </summary>
        ''' <param name="serialNumber">Unit's serial number.</param>
        ''' <param name="udbsProductID">Udbs Product ID.</param>
        ''' <param name="stage">Process Stage.</param>
        Protected Sub New(serialNumber As String, udbsProductID As String, stage As String)

            If Not String.IsNullOrEmpty(serialNumber) AndAlso Not String.IsNullOrEmpty(udbsProductID) AndAlso Not String.IsNullOrEmpty(stage) Then
                _serialNumber = serialNumber
                _ubdsProductID = udbsProductID
                _stage = stage

            Else
                Throw New ApplicationException("The serialNumber, udbsProductID and stage must be specified!")
            End If
        End Sub

        Public ReadOnly Property ProcessID As Integer
            Get
                Return _processID
            End Get
        End Property

        Public ReadOnly Property UnitID As Integer
            Get
                Return _unitID
            End Get
        End Property

        Public ReadOnly Property ItemListRevisionID As Integer
            Get
                Return _itemListRevisionID
            End Get
        End Property

        Public ReadOnly Property Sequence As Integer
            Get
                Return _sequence
            End Get
        End Property

        Public ReadOnly Property StartDate As Date
            Get
                Return _startDate
            End Get
        End Property

        Public ReadOnly Property EndDate As Date
            Get
                Return _endDate
            End Get
        End Property

        Public ReadOnly Property Status As UdbsProcessStatus
            Get
                Return _status
            End Get
        End Property

        Public ReadOnly Property Result As TestDataInterface.ResultCodes
            Get
                Return _result
            End Get
        End Property

        Public ReadOnly Property Notes As String
            Get
                Return _notes
            End Get
        End Property

        Public ReadOnly Property Employee As String
            Get
                Return _employee
            End Get
        End Property

        Public ReadOnly Property ActiveDuration As Double
            Get
                Return _activeDuration
            End Get
        End Property

        Public ReadOnly Property TotalDuration As Double
            Get
                Return _totalDuration
            End Get
        End Property

        ''' <summary>
        ''' Loads all the process info properties of the process info object.
        ''' </summary>
        ''' <param name="resultTable">Result table containing the data obtained from the database.</param>
        Protected MustOverride Sub Load(ByRef resultTable As DataTable)

        ''' <summary>
        ''' Gets the number of sequences for the specified Product/Unit/Stage
        ''' </summary>
        ''' <param name="productNumber">Udbs product ID</param>
        ''' <param name="serialNumber">unit's serial number.</param>
        ''' <param name="stage">test stage.</param>
        ''' <param name="processType"><see cref="ProcessTypes"/></param>
        ''' <returns>Latest sequence number.</returns>
        Protected Shared Function GetSequenceCount(productNumber As String,
                                          serialNumber As String,
                                          stage As String, processType As ProcessTypes) As Integer

            Dim processTable = $"{processType}_process"
            Dim itemListRevisionTable = $"{processType}_itemlistrevision"

            Dim sqlQuery = "SELECT MAX(process_sequence) AS maxSeq " &
                           "FROM " & "product" & " with(nolock), " & "unit" & " with(nolock), " & processTable &
                           " with(nolock), " & itemListRevisionTable & " with(nolock) " &
                           "WHERE product_id=unit_product_id " &
                           "AND unit_id=process_unit_id " &
                           "AND itemlistrev_id=process_itemlistrev_id " &
                           "AND product_number = '" & productNumber & "' " &
                           "AND unit_serial_number = '" & serialNumber & "' " &
                           "AND itemlistrev_stage = '" & stage & "'"

            Dim rsTemp As New DataTable
            OpenNetworkRecordSet(rsTemp, sqlQuery)

            Dim sequence As Integer

            If (If(rsTemp?.Rows?.Count, 0)) = 0 Then

                sequence = 0

            ElseIf IsDBNull(rsTemp(0)("maxSeq")) Then

                sequence = 0
            Else

                sequence = KillNullInteger(rsTemp(0)("maxSeq"))
            End If

            Return sequence
        End Function

        ''' <summary>
        ''' Gets the process identifier for the specified Product/Unit/Stage/Sequence
        ''' </summary>
        ''' <param name="productNumber">Udbs product ID.</param>
        ''' <param name="serialNumber">Unit's serial number.</param>
        ''' <param name="stage">test stage.</param>
        ''' <param name="processType"><see cref="ProcessTypes"/></param>
        ''' <param name="sequence">Test sequence number.</param>
        ''' <returns>The unique test process ID.</returns>
        Protected Shared Function GetProcessID(productNumber As String,
                                      serialNumber As String,
                                      stage As String, processType As ProcessTypes,
                                      Optional sequence As Integer = 0) As Integer

            Dim desiredSequence As Integer
            Dim rsTemp As New DataTable

            Try
                If sequence = 0 Then
                    ' Select latest sequence
                    desiredSequence = GetSequenceCount(productNumber, serialNumber, stage, processType)
                Else
                    desiredSequence = sequence
                End If

                Dim processTable = $"{processType}_process"
                Dim itemListRevisionTable = $"{processType}_itemlistrevision"

                Dim sqlQuery = "SELECT process_id " &
                           "FROM " & processTable & " with(nolock), " & itemListRevisionTable & " with(nolock), " &
                           "unit" & " with(nolock), " & "product" & "  with(nolock) " &
                           "WHERE process_itemlistrev_id=itemlistrev_id " &
                           "AND product_id=unit_product_id " &
                           "AND unit_id=process_unit_id " &
                           "AND product_number = '" & productNumber & "' " &
                           "AND unit_serial_number = '" & serialNumber & "' " &
                           "AND process_sequence = " & CStr(desiredSequence) &
                           " AND itemlistrev_stage = '" & stage & "'"


                OpenNetworkRecordSet(rsTemp, sqlQuery)

                If (If(rsTemp?.Rows?.Count, 0)) = 0 Then
                    LogError(New Exception($"No test process ID found in the database for unit '{serialNumber}' at stage '{stage}'."))
                    Throw New ApplicationException($"No test process ID found in the database for unit '{serialNumber}' at stage '{stage}'.")
                End If

                Dim processID = KillNullInteger(rsTemp(0)("process_id"))

                Return processID
            Finally
                rsTemp?.Dispose()
            End Try


        End Function

    End Class

End Namespace

