Imports UdbsInterface.MasterInterface

Namespace KittingInterface
    Public Class KittingProcessInfo
        Inherits ProcessInfo

        Private _workOrder As String

        Public ReadOnly Property WorkOrder As String
            Get
                Return _workOrder
            End Get
        End Property

        ''' <summary>
        ''' Private to prevent direct instantiation.
        ''' </summary>
        Private Sub New(serialNumber As String, udbsProductID As String, stage As String)
            MyBase.New(serialNumber, udbsProductID, stage)
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub Load(ByRef resultTable As DataTable)
            If resultTable.Rows.Count = 0 Then

                Throw New ApplicationException($"No kitting process information records were found for (id = {ProcessID}) from the network database.")

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
                _workOrder = KillNull(resultTable(0)("process_workorder"))

            End If
        End Sub

        ''' <summary>
        ''' Gets the number of sequences for the specified Product/Unit/Stage
        ''' </summary>
        ''' <param name="productNumber">Udbs product ID</param>
        ''' <param name="serialNumber">unit's serial number.</param>
        ''' <param name="stage">kitting stage.</param>
        ''' <returns>Latest sequence number.</returns>
        Public Overloads Shared Function GetSequenceCount(productNumber As String,
                                          serialNumber As String,
                                          stage As String) As Integer

            Return ProcessInfo.GetSequenceCount(productNumber, serialNumber, stage, ProcessTypes.kitting)
        End Function

        ''' <summary>
        ''' Gets the process identifier for the specified Product/Unit/Stage/Sequence
        ''' </summary>
        ''' <param name="productNumber">Udbs product ID.</param>
        ''' <param name="serialNumber">Unit's serial number.</param>
        ''' <param name="stage">kitting stage.</param>
        ''' <param name="sequence">kitting sequence number.</param>
        ''' <returns>The unique process ID.</returns>
        Protected Overloads Shared Function GetProcessID(productNumber As String,
                                      serialNumber As String,
                                      stage As String, Optional sequence As Integer = 0) As Integer

            Return ProcessInfo.GetProcessID(productNumber, serialNumber, stage, ProcessTypes.kitting, sequence)
        End Function

        ''' <summary>
        ''' Gets the kitting data process information for the specified Product/Unit/Stage/Sequence
        ''' </summary>
        ''' <param name="productNumber">Udbs product ID.</param>
        ''' <param name="serialNumber">Unit's serial number.</param>
        ''' <param name="stage">kitting stage.</param>
        ''' <param name="sequence">kitting sequence.</param>
        ''' <returns>A process Info object.</returns>
        Public Shared Function GetProcessInfo(productNumber As String, serialNumber As String,
                                              stage As String, Optional sequence As Integer = 0) As KittingProcessInfo

            Dim processID = GetProcessID(productNumber, serialNumber, stage, sequence)

            Dim sqlQuery = "SELECT * FROM " & "kitting_process" & " with(nolock) WHERE process_id = " & CStr(processID)

            Dim resultTable = New DataTable
            If QueryNetworkDB(sqlQuery, resultTable) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New UdbsTestException($"Error retrieving process information for (id = {processID}) from the network database.")
            End If

            Dim processInfo As New KittingProcessInfo(serialNumber, productNumber, stage) With {._processID = processID}

            processInfo.Load(resultTable)

            Return processInfo

        End Function

    End Class

End Namespace


