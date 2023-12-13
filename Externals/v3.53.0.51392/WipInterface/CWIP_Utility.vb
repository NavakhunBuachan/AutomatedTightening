Option Explicit On
Option Compare Binary
Option Infer On
Option Strict On

Imports UdbsInterface.MasterInterface

Namespace WipInterface
    Public Class CWIP_Utility
        Implements IDisposable

        Private Const ClsName As String = "CWIP_Utility"

        Friend Function CheckUserPrivileges(EmployeeNumber As String, Pword As String, AllowNewLogin As Boolean,
                                            ByVal ParamArray GroupNames As String()) As Boolean
            Dim tmpSecurity As New SecurityInterface
            Dim i As Integer, j As Integer
            Dim Permit As Boolean
            Dim UserGroups As String() = Nothing

            'check the privileges for the specified employee number
            Permit = False
            If Not tmpSecurity.GetGroupMembership(EmployeeNumber, UserGroups) Then
                logger.Debug($"Failed to retrieve groups of employee {EmployeeNumber}")
                Return False
            End If

            For i = 0 To UBound(UserGroups)
                For j = 0 To UBound(GroupNames)
                    If CStr(UserGroups.GetValue(i)).Equals(GroupNames(j), StringComparison.InvariantCultureIgnoreCase) _
                        Then
                        Permit = True
                        Exit For
                    End If
                Next j
                If Permit = True Then Exit For
            Next i
            Permit = Permit And tmpSecurity.VerifyPassword(EmployeeNumber, Pword)

            'give the user another chance to log on
            If Not Permit And AllowNewLogin Then
                If tmpSecurity.LogIn(True, EmployeeNumber, Pword) Then
                    tmpSecurity.GetGroupMembership(tmpSecurity.AuthenticatedEmployeeNumber, UserGroups)
                    For i = 0 To UBound(UserGroups)
                        For j = 0 To UBound(GroupNames)
                            If CStr(UserGroups.GetValue(i)) = GroupNames(j) Then
                                Permit = True
                                Exit For
                            End If
                        Next j
                        If Permit = True Then Exit For
                    Next i
                End If
            End If

            CheckUserPrivileges = Permit
        End Function

        ' Candidate for removal.
        Private Function LookupProductID(FamilyName As String, ProductNumber As String, ProductRelease As Integer) _
            As Integer
            'retrieves the product id of a given product for a given release; returns -1 if error ocurred
            Dim SQLstr As String
            Dim rsTemp As DataTable = Nothing
            Dim RC As ReturnCodes

            Try

                SQLstr = "SELECT product_id FROM family with(nolock), product with(nolock) " &
                         "WHERE product_family_id  = family_id " &
                         "AND family_name = '" & FamilyName & "' " &
                         "AND product_number = '" & ProductNumber & "' " &
                         "AND product_release = '" & ProductRelease & "'"
                RC = QueryNetworkDB(SQLstr, rsTemp)

                If RC <> ReturnCodes.UDBS_OP_SUCCESS Then Return ReturnCodes.UDBS_OP_FAIL

                If (If(rsTemp?.Rows?.Count, 0)) = 1 Then
                    Return KillNullInteger(rsTemp(0)("product_id"))
                Else
                    LogError(New Exception($"Cannot find product with Family: {FamilyName} ProductNumber: {ProductNumber} Release: {ProductRelease}"))
                    Return -1
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return -1
            End Try

        End Function

        ' Candidate for removal.
        Private Function GetStationName(ByRef StationName As String) As ReturnCodes
            ' Purpose: Return the name of the computer as specified in the network settings
            Return CUtility.Utility_GetStationName(StationName)
        End Function

        Friend Function GetMaxProductRelease(ProductNumber As String, ByRef ProductRelease As Integer) As ReturnCodes
            'retrieves the current release of a given product
            Dim SQLstr As String
            Dim rsTemp As DataTable = Nothing

            Try
                SQLstr = "SELECT product_release FROM product with(nolock) " &
                     "WHERE product_number = '" & ProductNumber & "' " &
                     "ORDER BY product_release DESC"

                Dim returnCode = QueryNetworkDB(SQLstr, rsTemp)
                If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode

                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                    ProductRelease = KillNullInteger(rsTemp(0)("product_release"))
                Else
                    LogError(New Exception($"Product: {ProductNumber} not found!"))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        Friend Function GetMaxItemlistRevision(ProductID As Integer, Stage As String, ByRef ILRevision As Integer) _
            As ReturnCodes
            'retrieves the current release of a given product
            Dim SQLstr As String
            Dim rsTemp As DataTable = Nothing

            Try
                SQLstr = "SELECT itemlistrev_revision FROM WIP_itemlistrevision with(nolock) " &
                     "WHERE itemlistrev_product_id  = " & ProductID & " " &
                     "AND itemlistrev_stage = '" & Stage & "' " &
                     "ORDER BY itemlistrev_revision DESC"

                Dim returnCode = QueryNetworkDB(SQLstr, rsTemp)
                If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode
                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then

                    ILRevision = KillNullInteger(rsTemp(0)("itemlistrev_revision"))
                Else
                    ILRevision = -1
                    LogError(New Exception($"No products match product: {ProductID}, stage: {Stage}"))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        '''     Gets a data table containing information for units in WIP for a given UDBS ID and WIP stage
        '''     This code ported and modified from the WIPTracker VB6 project
        ''' </summary>
        ''' <param name="udbsId"></param>
        ''' <param name="wipStep"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUnitsInWip(udbsId As String, wipStep As String) As DataTable

            Dim unitsInWip As New DataTable

            ' 1. Set up the datatable
            unitsInWip.Columns.Add("Family")
            unitsInWip.Columns.Add("ProductNumber")
            unitsInWip.Columns.Add("ProductDescriptor")
            unitsInWip.Columns.Add("SerialNumber")
            unitsInWip.Columns.Add("WorkOrder")
            unitsInWip.Columns.Add("TrackingNumber")
            unitsInWip.Columns.Add("StartDate")
            unitsInWip.Columns.Add("EndDate")
            unitsInWip.Columns.Add("ActiveStep")
            unitsInWip.Columns.Add("Status")
            unitsInWip.Columns.Add("Notes")
            unitsInWip.Columns.Add("LockedBy")
            unitsInWip.Columns.Add("Age")

            ' 2. Load the WIP
            Dim strSQL = "SELECT family_name, product_number, product_descriptor, " &
                         "unit_serial_number, unit_id, process_id, process_work_order, process_tracking_number, " &
                         "process_start_date, process_end_date, process_active_step, process_status, " &
                         "process_notes, itemlistrev_id, process_locked_by, itemlistrev_stage, " &
                         "itemlistdef_descriptor, itemlistdef_id, itemlistrev_revision " &
                         "FROM family with(nolock), product with(nolock), unit with(nolock), WIP_process with(nolock), WIP_itemlistrevision with(nolock), " &
                         "WIP_itemlistdefinition with(nolock) WHERE (product_family_id = family_id AND " &
                         "unit_product_id = product_id AND process_unit_id = unit_id " &
                         "AND process_itemlistrev_id = itemlistrev_id AND " &
                         "process_itemlistrev_id=itemlistdef_itemlistrev_id " &
                         "AND process_active_step=itemlistdef_itemname) " &
                         "AND (process_status = 'PAUSED' OR process_status = 'IN PROCESS') " &
                         "AND product_number LIKE '" & udbsId & "' " &
                         "AND process_active_step LIKE '" & wipStep & "' " &
                         "ORDER BY process_end_date ASC"
            Dim rs As DataTable = Nothing
            Dim rc As ReturnCodes = QueryNetworkDB(strSQL, rs)
            If Not rc = ReturnCodes.UDBS_OP_SUCCESS Then _
                Throw New Exception("Failed to get units in WIP due to database query error. Return code = " & rc)

            ' 3. Did we find anything?
            If (If(rs?.Rows?.Count, 0)) > 0 Then

                ' 3.1. Calculate aging on what we did find and populate the datatable
                Dim i = 1
                Dim currenttime As Date
                CUtility.Utility_GetServerTime(currenttime)
                unitsInWip.BeginLoadData()
                For Each dr As DataRow In rs.Rows

                    ' 3.1.1 Aging
                    Dim sSQL As String = "SELECT * FROM WIP_result with(nolock) WHERE result_process_id = " &
                                         KillNull(dr("process_id")) &
                                         " ORDER BY result_step_number DESC"
                    Dim rsWIP As DataTable = Nothing
                    Dim agingTime As New TimeSpan(0)
                    rc = QueryNetworkDB(sSQL, rsWIP)
                    If rc = ReturnCodes.UDBS_OP_SUCCESS Then
                        '//new aging time calculation
                        If KillNull(rsWIP(0)("result_start_date")) = "" Then
                            'process paused between steps, so aging time here is actually active step waiting time
                            agingTime = currenttime - CDate(KillNull(dr("process_end_date")))
                            'flxWIP.TextMatrix(i, 9) = "Waiting"
                        Else
                            'active step in process , so aging time here is really the aging time of active step
                            agingTime = currenttime - CDate(KillNull(rsWIP(0)("result_start_date")))
                            'flxWIP.TextMatrix(i, 9) = "In Process"
                        End If
                        rsWIP = Nothing
                    Else
                        Throw New Exception("Failed to query for WIP results.")
                    End If

                    If agingTime.TotalHours < 0 Then agingTime = New TimeSpan(0)

                    ' 3.1.2 DataTable
                    unitsInWip.Rows.Add(KillNull(dr("family_name")),
                                        KillNull(dr("product_number")),
                                        KillNull(dr("product_descriptor")),
                                        KillNull(dr("unit_serial_number")),
                                        KillNull(dr("process_work_order")),
                                        KillNull(dr("process_tracking_number")),
                                        KillNull(dr("process_start_date")),
                                        KillNull(dr("process_end_date")),
                                        KillNull(dr("process_active_step")),
                                        KillNull(dr("process_status")),
                                        KillNull(dr("process_notes")),
                                        KillNull(dr("process_locked_by")),
                                        agingTime.TotalDays)


                    i = i + 1

                Next
                unitsInWip.EndLoadData()
            End If

            ' 4. Return
            Return unitsInWip
        End Function

        ''' <summary>
        ''' Allows to reroute (repeat, skip) a process step. This Is used for bypassing optional steps in process.
        ''' </summary>
        ''' <param name="serialNumber">Serial number of unit being processed.</param>
        ''' <param name="stepName">Step to reroute to.</param>
        ''' <param name="note">Note for rerouting unit.</param>
        <ObsoleteAttribute("This method is obsolete. Please use TrySoftwareReRoute(...) instead.")>
        Public Sub SoftwareReRoute(SerialNumber As String, stepName As String, note As String)
            Using wipProcess As New CWIP_Process
                If wipProcess.LoadActiveProcess(SerialNumber, LockStatus_Enum.READ_WRITE) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Throw New UDBSException($"Fail to load active process for unit: {SerialNumber}")
                End If

                Try
                    If wipProcess.ReRoute(stepName, note, "eng", "eng") <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Throw New Exception($"Fail to re-route unit {SerialNumber}")
                    End If
                Finally
                    If wipProcess.Unlock_Process() <> ReturnCodes.UDBS_OP_SUCCESS Then
                        logger.Warn($"Fail to unlock process for unit {SerialNumber}")
                    End If
                End Try
            End Using
        End Sub

        ''' <summary>
        ''' Tries to reroute (repeat, skip) a process step. This Is used for bypassing optional steps in process.
        ''' </summary>
        ''' <param name="serialNumber">Serial number of unit being processed.</param>
        ''' <param name="stepName">Step to reroute to.</param>
        ''' <param name="note">Note for rerouting unit.</param>
        ''' <returns>Return Code indicating UDBS operation status. </returns>
        Friend Function TrySoftwareReRoute(SerialNumber As String, stepName As String, note As String) As ReturnCodes
            Using wipProcess As New CWIP_Process
                If wipProcess.LoadActiveProcess(SerialNumber, LockStatus_Enum.READ_WRITE) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                Try
                    Return wipProcess.ReRoute(stepName, note, "eng", "eng")
                Finally
                    If wipProcess.Unlock_Process() <> ReturnCodes.UDBS_OP_SUCCESS Then
                        logger.Warn($"Fail to unlock process for unit: {SerialNumber}")
                    End If
                End Try
            End Using
        End Function

        ''' <summary>
        ''' Logs the UDBS error in the application log and inserts it in the udbs_error table, with the process instance details:
        ''' Process Type, name, Process ID, UDBS product ID, Unit serial number.
        ''' </summary>
        ''' <param name="ex">Exception raised.</param>
        Private Sub LogErrorInDatabase(ex As Exception)

            DatabaseSupport.LogErrorInDatabase(ex, PROCESS, String.Empty, 0, String.Empty, String.Empty)

        End Sub

#Region "IDisposable Support"

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    CloseNetworkDB()
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
