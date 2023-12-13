Imports UdbsInterface.MasterInterface


Namespace KittingInterface
    ''' <summary>
    '''     Problem accessing the Kitting database
    ''' </summary>
    Friend Class UdbsKittingException
        Inherits Exception

        Public Sub New(reason As String)
            MyBase.New(reason & " Please use the UDBS Traveller software to update the kitting information.")
        End Sub
    End Class

    Friend Module KittingSupport
        Friend Logger As Logger = LogManager.GetLogger("UDBS")
    End Module

    Public Class KittingUtility


        ''' <summary>
        ''' Shim for back-compatibility of the method <see cref="LookupKittingInfo(String, String, ByRef String, ByRef String, ByRef String, ByRef Integer, ByRef String)"/>.
        ''' </summary>
        Public Shared Function LookupKittingInfo(serialNumber As String,
                                                 kittingItemname As String,
                                                 ByRef kittingItemSN As String,
                                                 ByRef kittingItemPartNumber As String,
                                                 ByRef kittingItemPartID As String,
                                                 ByRef kittingSequence As Integer) As Boolean
            Dim kittingRemarks As String = String.Empty ' Unused.
            Return LookupKittingInfo(serialNumber, kittingItemname, kittingItemSN, kittingItemPartNumber, kittingItemPartID, kittingSequence, kittingRemarks)
        End Function

        ''' <summary>
        ''' Based on the serial number of the optics module and the type of component given,
        ''' looks up the part number and serial number of the component part from the kitting database,
        ''' and returns the values by reference.
        ''' </summary>
        ''' <param name="serialNumber">Input serial number of optics module</param>
        ''' <param name="kittingItemname">Input name of item in kitting database</param>
        ''' <param name="kittingItemSN">Output Serial number of the found component part</param>
        ''' <param name="kittingItemPartNumber">Output Oracle part number of the found component part</param>
        ''' <param name="kittingItemPartID">Output Part ID of the found component part</param>
        ''' <param name="kittingSequence">Output Kitting sequence number</param>
        ''' <param name="kittingRemarks">Output Remarks provided when kitting information was updated.</param>
        ''' <returns>True if successful, False otherwise</returns>
        Public Shared Function LookupKittingInfo(serialNumber As String,
                                                 kittingItemname As String,
                                                 ByRef kittingItemSN As String,
                                                 ByRef kittingItemPartNumber As String,
                                                 ByRef kittingItemPartID As String,
                                                 ByRef kittingSequence As Integer,
                                                 ByRef kittingRemarks As String) As Boolean

            Dim rsTmp As New DataTable()

            Try
                ' Build query
                Dim query As String =
                        $"SELECT TOP 1 result_partnumber, result_serialnumber, process_sequence, result_stringdata 
                       FROM product with(nolock), unit with(nolock), kitting_process with(nolock), kitting_result with(nolock), kitting_itemlistdefinition with(nolock) 
                       WHERE process_id=result_process_id 
                       AND product_id=unit_product_id 
                       AND unit_id=process_unit_id 
                       AND itemlistdef_id=result_itemlistdef_id 
                       AND unit_serial_number='{serialNumber}' 
                       AND itemlistdef_itemname='{kittingItemname}' 
                       ORDER BY process_sequence DESC"

                ' Query database          
                OpenNetworkRecordSet(rsTmp, query)

                Dim fields As New Dictionary(Of String, Object)

                If rsTmp?.Rows.Count > 0 Then
                    Dim itemNum = 0
                    For Each tmpField As DataColumn In rsTmp.Columns
                        fields(tmpField.ColumnName) = rsTmp.Rows(0).Item(itemNum)
                        itemNum += 1
                    Next

                    If IsDBNull(fields("process_sequence")) Then
                        kittingSequence = Integer.MinValue
                    Else
                        kittingSequence = CInt(fields("process_sequence"))
                    End If

                    If IsDBNull(fields("result_partnumber")) Then
                        kittingItemPartNumber = String.Empty
                    Else
                        kittingItemPartNumber = UCase(Trim(fields("result_partnumber").ToString))
                        Try
                            kittingItemPartID = LookupPartID(kittingItemPartNumber)
                        Catch ex As Exception
                            kittingItemPartID = String.Empty
                        End Try
                    End If

                    If IsDBNull(fields("result_stringdata")) Then
                        kittingRemarks = String.Empty
                    Else
                        kittingRemarks = $"{fields("result_stringdata")}".Trim()
                    End If

                    'Return True, only if a non-empty serial number is found
                    If IsDBNull(fields("result_serialnumber")) Then
                        kittingItemSN = String.Empty
                        Return False
                    Else
                        kittingItemSN = UCase(Trim(fields("result_serialnumber").ToString))
                        Return Not String.IsNullOrEmpty(kittingItemSN)
                    End If

                Else
                    Return False
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)

                ' Return FALSE if error encountered
                Return False
            Finally
                rsTmp?.Dispose()
            End Try
        End Function

        ''' <summary>
        '''     Translates product/part number to product identifier
        ''' </summary>
        ''' <param name="partNumber">String. Oracle Part Number</param>
        ''' <returns>String. UDBS Product ID.</returns>
        ''' <remarks>Suppresses exceptions and returns null string if not found</remarks>
        Private Shared Function LookupPartID(partNumber As String) As String
            If String.IsNullOrWhiteSpace(partNumber) Then
                Return vbNullString
            End If

            Try
                Using utility As New CUtility
                    Dim productID As String = vbNullString
                    Dim rc As ReturnCodes = utility.Product_GetPartIdentifier(partNumber, productID)
                    If rc <> ReturnCodes.UDBS_OP_SUCCESS Then
                        If UDBSDebugMode Then
                            Logger.Debug("Could not find part ID for part number {0}", partNumber)
                        End If
                        Return Nothing
                    Else
                        Return productID
                    End If
                End Using
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        '''     Creates a new kitting process sequence for the given serial number with the given UDBS ID.
        '''     Copies all results from the most recent process sequence.
        '''     Updates the PN and SN for each item given in the sItemname array.
        ''' </summary>
        ''' <param name="udbsID"></param>
        ''' <param name="serialNo"></param>
        ''' <param name="empID"></param>
        ''' <param name="itemname"></param>
        ''' <param name="componentPN"></param>
        ''' <param name="componentSN"></param>
        ''' <param name="notes"></param>
        ''' <returns></returns>
        Public Shared Function UpdateKittingData(udbsID As String,
                                                 serialNo As String,
                                                 empID As String,
                                                 itemname() As String,
                                                 componentPN() As String,
                                                 componentSN() As String,
                                                 Optional notes As String = "") As Boolean
            Return UpdateKittingDataAndRemarks(udbsID, serialNo, empID, itemname, componentPN, componentSN, Nothing, notes)
        End Function

        ' Query the database for the most recent process sequence for this unit
        Private Shared Sub GetKittingProcessInfo(
                unitID As Integer,
                ByRef processID As Integer,
                ByRef workOrder As String,
                ByRef itemListID As Integer,
                ByRef sequence As Integer)
            Dim rsProcess As DataTable = Nothing
            Try
                Dim query =
                            "SELECT TOP 1 kitting_process.* FROM kitting_process with(nolock), kitting_itemlistrevision with(nolock) " &
                            "WHERE process_itemlistrev_id=itemlistrev_id AND process_unit_id=" &
                            unitID & " AND itemlistrev_stage='primary' " &
                            "ORDER BY process_sequence DESC"

                OpenNetworkRecordSet(rsProcess, query)

                If (If(rsProcess?.Rows?.Count, 0)) = 0 Then
                    Throw New UdbsKittingException("Cannot generate the first kitting sequence")
                End If

                ' Get the process ID and the workorder
                processID = KillNullInteger(rsProcess(0)("process_id"))
                workOrder = CStr(rsProcess(0)("process_workorder"))
                If String.IsNullOrEmpty(workOrder) Then
                    Throw New UdbsKittingException("Cannot create a new kitting sequence without a work order")
                End If

                ' Get the item list revision (relational link to the kitting_result table)
                itemListID = KillNullInteger(rsProcess(0)("process_itemlistrev_id"))
                sequence = KillNullInteger(rsProcess(0)("process_sequence"))
            Finally
                rsProcess?.Dispose()
            End Try
        End Sub

        ''' <summary>
        '''     Creates a new kitting process sequence for the given serial number with the given UDBS ID.
        '''     Copies all results from the most recent process sequence.
        '''     Updates the PN and SN for each item given in the sItemname array.
        ''' </summary>
        ''' <param name="udbsID">UDBS Product ID.</param>
        ''' <param name="serialNo">The unit's serial number.</param>
        ''' <param name="empID">The employee ID.</param>
        ''' <param name="itemname">The item name to update.</param>
        ''' <param name="componentPN">The kitted component's part number.</param>
        ''' <param name="componentSN">The kitted component's serial number.</param>
        ''' <param name="remarks">Remarks to tag to the kitting update operation.</param>
        ''' <param name="notes">Notes to tag to the kitting update operation.</param>
        ''' <returns>Whether or not the operation succeeds.</returns>
        Friend Shared Function UpdateKittingDataAndRemarks(udbsID As String,
                                                           serialNo As String,
                                                           empID As String,
                                                           itemname() As String,
                                                           componentPN() As String,
                                                           componentSN() As String,
                                                           remarks() As String,
                                                           Optional notes As String = "") As Boolean

            Try
                Dim productGroup As New clsPrdGrp()
                Dim unitID = productGroup.GetUnitID(udbsID, serialNo)

                Dim processID As Integer = -1
                Dim values As List(Of Object) = Nothing

                ' Query the database for the most recent process sequence for this unit
                Dim workOrder As String = Nothing
                Dim itemListID As Integer = -1
                Dim processSeq As Integer = -1
                GetKittingProcessInfo(unitID, processID, workOrder, itemListID, processSeq)

                ' Increment the process sequence
                processSeq += 1

                Dim curTime As Date
                CUtility.Utility_GetServerTime(curTime)

                ' Add a new process sequence
                values = New List(Of Object)() From {unitID,
                    itemListID,
                    processSeq,
                    curTime,
                    curTime,
                    empID,
                    "IN PROGRESS",
                    "INCOMPLETE",
                    workOrder,
                    My.Application.Info.Title & If(String.IsNullOrEmpty(notes), "", ": " & notes)}

                Using transactionScope As ITransactionScope = BeginNetworkTransaction()
                    Dim rsOldResult As DataTable = Nothing
                    Dim rsNewResult As DataTable = Nothing

                    Try
                        ' Get the list of components for this process ID (i.e. the last process sequence)
                        Dim query = "SELECT kitting_itemlistdefinition.*, kitting_result.* " &
                                "FROM kitting_itemlistdefinition with(nolock) INNER JOIN kitting_result with(nolock) ON itemlistdef_id = result_itemlistdef_id " &
                                "WHERE result_process_id = " & processID & " " &
                                "ORDER BY itemlistdef_itemnumber"
                        OpenNetworkRecordSet(rsOldResult, query)

                        Dim fields As New List(Of String)() From {"process_unit_id",
                                "process_itemlistrev_id",
                                "process_sequence",
                                "process_start_date",
                                "process_end_date",
                                "process_employee_number",
                                "process_status",
                                "process_result",
                                "process_workorder",
                                "process_notes"}
                        processID = InsertNetworkRecord(fields.ToArray(), values.ToArray(), "kitting_process",
                                                        transactionScope, "process_id")

                        query = "SELECT * FROM kitting_result with(nolock) WHERE result_process_id=" & processID
                        OpenNetworkRecordSet(rsNewResult, query)

                        If (If(rsNewResult?.Rows?.Count, 0)) > 0 Then
                            ' Something's fishy here, kitting result should be empty.
                            Throw New UdbsKittingException(
                                    "Component information already exists in UDBS for the new process ID " & processID &
                                    ". Contact technical staff to investigate.")
                        End If

                        ' Copy the old results into the new results, updating as necessary
                        Dim result = "COMPLETE"

                        If (If(rsOldResult?.Rows?.Count, 0)) > 0 Then
                            For Each drOld As DataRow In rsOldResult.Rows

                                ' See if this part is in the array of items
                                Dim idxItem As Integer = -1
                                Dim i As Integer = LBound(itemname)
                                Dim thisItemName = CStr(drOld("itemlistdef_itemname"))
                                Do While i <= UBound(itemname) And idxItem = -1
                                    If thisItemName = itemname(i) Then
                                        idxItem = i
                                    End If
                                    i = i + 1
                                Loop

                                ' Add a new result
                                fields = New List(Of String) From {"result_process_id", "result_itemlistdef_id"}
                                values = New List(Of Object) From {processID, drOld("result_itemlistdef_id")}

                                Dim oldSN As String =
                                        If _
                                        (IsDBNull(drOld("result_serialnumber")), "", CStr(drOld("result_serialnumber")))

                                If idxItem = -1 Then
                                    If (Not IsDBNull(drOld("result_partnumber"))) Then
                                        fields.Add("result_partnumber")
                                        values.Add(drOld("result_partnumber"))
                                    Else
                                        If KillNullInteger(drOld("itemlistdef_partnumber_flag")) = 1 Then _
                                            result = "INCOMPLETE"
                                    End If
                                    If (Not IsDBNull(drOld("result_serialnumber"))) Then
                                        fields.Add("result_serialnumber")
                                        values.Add(drOld("result_serialnumber"))
                                    Else
                                        If KillNullInteger(drOld("itemlistdef_serialnumber_flag")) = 1 Then _
                                            result = "INCOMPLETE"
                                    End If
                                    If (Not IsDBNull(drOld("result_stringdata"))) Then
                                        fields.Add("result_stringdata")
                                        values.Add(drOld("result_stringdata"))
                                    End If
                                    fields.Add("result_replacement")
                                    values.Add(0)
                                Else
                                    'Update the SN and PN
                                    If componentPN(idxItem) = "" Then
                                        fields.Add("result_partnumber")
                                        values.Add(drOld("result_partnumber"))
                                    Else
                                        fields.Add("result_partnumber")
                                        values.Add(componentPN(idxItem))
                                    End If
                                    fields.Add("result_serialnumber")
                                    values.Add(componentSN(idxItem))

                                    If (remarks IsNot Nothing) Then
                                        fields.Add("result_stringdata")
                                        values.Add(remarks(idxItem))
                                    End If

                                    If oldSN = "" Then
                                        fields.Add("result_replacement")
                                        values.Add(0)
                                        If (oldSN <> componentSN(idxItem)) Then
                                            Dim idx As Integer = fields.IndexOf("result_stringdata")
                                            If idx >= 0 Then
                                                values(idx) = $"{values(idx)}, Added by {My.Application.Info.Title}"
                                            Else
                                                fields.Add("result_stringdata")
                                                values.Add($"Added by {My.Application.Info.Title}")
                                            End If

                                        End If
                                    Else
                                        If (KillNullInteger(drOld("result_replacement")) = 1) Or
                                                (oldSN <> componentSN(idxItem)) Then
                                            fields.Add("result_replacement")
                                            values.Add(1)

                                            Dim idx As Integer = fields.IndexOf("result_stringdata")
                                            If idx >= 0 Then
                                                values(idx) = $"{values(idx)}, Changed by {My.Application.Info.Title}"
                                            Else
                                                fields.Add("result_stringdata")
                                                values.Add($"Changed by {My.Application.Info.Title}")
                                            End If
                                        Else
                                            fields.Add("result_replacement")
                                            values.Add(0)
                                        End If
                                    End If
                                End If

                                If (Not IsDBNull(drOld("result_linkto_process_id"))) Then
                                    fields.Add("result_linkto_process_id")
                                    values.Add(drOld("result_linkto_process_id"))
                                End If

                                InsertNetworkRecord(fields.ToArray(), values.ToArray(), "kitting_result",
                                                    transactionScope)
                            Next
                        End If

                        fields = New List(Of String) _
                            From {"process_status", "process_result", "process_id", "process_unit_id"}
                        values = New List(Of Object) From {"COMPLETED", result, processID, unitID}
                        Dim keys As String() = {"process_id", "process_unit_id"}

                        UpdateNetworkRecord(keys, fields.ToArray(), values.ToArray(), "kitting_process",
                                            transactionScope)
                    Catch ex As Exception
                        transactionScope.HasError = True ' when in doubt, rollback!
                        Throw
                    Finally
                        rsOldResult?.Dispose()
                        rsNewResult?.Dispose()
                    End Try
                End Using

                If UDBSDebugMode Then
                    Logger.Debug("Finished updating kitting database.")
                End If

                Return True

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Get the highest kitting item list revision used by a given unit.
        ''' </summary>
        ''' <param name="serialNumber">The unit's serial number.</param>
        ''' <returns>The highest revision. -1 if the operation fails.</returns>
        ''' <remarks>TODO: Rename and make the 'L' of 'List' upper-case.</remarks>
        Public Shared Function GetHighestKittingItemlistRevision(serialNumber As String) As Integer

            Dim rsTmp As DataTable = Nothing
            Dim sSQL As String = "SELECT max(itemlistrev_revision) " &
                                 "FROM unit with (nolock), kitting_process with (nolock), kitting_itemlistrevision with (nolock) " &
                                 "WHERE " &
                                 "process_unit_id = unit_id " &
                                 "AND " &
                                 "itemlistrev_id = process_itemlistrev_id " &
                                 "AND " &
                                 "itemlistrev_stage='primary' " &
                                 "AND " &
                                 "unit_serial_number='<sn>'"
            Try
                OpenNetworkRecordSet(rsTmp, sSQL.Replace("<sn>", serialNumber))
                If (If(rsTmp?.Rows?.Count, 0)) > 0 Then
                    Dim dr = rsTmp.AsEnumerable().Last()
                    Return CInt(dr(0))
                Else
                    Return -1
                End If
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return -1
            Finally
                rsTmp?.Dispose()
            End Try
        End Function
    End Class
End Namespace
