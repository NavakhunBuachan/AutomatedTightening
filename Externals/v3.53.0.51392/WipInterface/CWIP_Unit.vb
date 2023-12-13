Option Explicit On
Option Compare Binary
Option Infer On
Option Strict On

Imports UdbsInterface.MasterInterface


Namespace WipInterface
    ''' <remarks>
    ''' This class' Dispose method doesn't do anything.
    ''' By contract, we have to call it, but is makes this class less user-friendly.
    ''' </remarks>
    Public Class CWIP_Unit
        Implements IDisposable

        Private Const ClsName = "CWIP_Unit"

        Private mfamily_name As String
        Private mproduct_id As Integer
        Private mproduct_number As String
        Private mproduct_release As Integer
        Private munit_id As Integer
        Private munit_serial_number As String
        Private munit_created_by As String
        Private munit_created_date As Date
        Private munit_create_report As String

        Private mLoaded As Boolean

        Public ReadOnly Property Loaded As Boolean
            Get
                Return mLoaded
            End Get
        End Property

        Public ReadOnly Property FamilyName As String
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

        Public ReadOnly Property ProductRelease As Integer
            Get
                Return mproduct_release
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

        Public ReadOnly Property CreatedBy As String
            Get
                Return munit_created_by
            End Get
        End Property

        Public ReadOnly Property CreatedDate As Date
            Get
                Return munit_created_date
            End Get
        End Property

        Public ReadOnly Property CreationReport As String
            Get
                Return munit_create_report
            End Get
        End Property

        Public Function LoadActiveUnit(SerialNumber As String, ByRef ProcessID As Integer) As ReturnCodes
            'Load a unit that has an active WIP process
            Dim strSQL As String
            Dim rsTemp As New DataTable

            Try

                UnloadUnit()

                strSQL = "SELECT unit_id, process_id " &
                         "FROM unit with(nolock) , WIP_process  with(nolock) " &
                         "WHERE process_unit_id = unit_id " &
                         "AND unit_serial_number = '" & SerialNumber & "' " &
                         "AND (process_status = 'IN PROCESS' OR process_status = 'PAUSED')"

                LoadActiveUnit = QueryNetworkDB(strSQL, rsTemp)
                If LoadActiveUnit <> ReturnCodes.UDBS_OP_SUCCESS Then Exit Function

                If (If(rsTemp?.Rows?.Count, 0)) <> 1 Then
                    logger.Debug($"Unit: {SerialNumber} is not active in WIP.")
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                ProcessID = KillNullInteger(rsTemp(0)("process_id"))
                Return LoadUnitByID(KillNullInteger(rsTemp(0)("unit_id")))

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Load the unit into this object.
        ''' If the product number is not provided, and the serial number is
        ''' not unique, the operation succeeds, the first one is returned and
        ''' an error message is logged.
        ''' </summary>
        ''' <param name="SerialNumber">The serial number of the unit to load.</param>
        ''' <param name="Family">The product family of unit to load. Helps for disambiguation.</param>
        ''' <param name="ProductNumber">The UDBS product number.</param>
        ''' <returns>Whether or not the unit was found and loaded.</returns>
        Public Function LoadUnit(SerialNumber As String,
                                 Optional ByVal Family As String = "",
                                 Optional ByVal ProductNumber As String = "") _
            As ReturnCodes
            ' Load itemlist by Product Number and Release
            Dim strSQL As String
            Dim rsTemp As New DataTable
            Dim localUnitID As Integer

            Try
                UnloadUnit()

                strSQL = "SELECT unit_id " &
                         "FROM family with(nolock) , product with(nolock) , unit  with(nolock) " &
                         "WHERE product_family_id = family_id " &
                         "AND unit_product_id = product_id " &
                         "AND unit_serial_number = '" & SerialNumber & "' "
                If Family <> "" Then strSQL = strSQL & "AND family_name = ('" & Family & "') "
                If ProductNumber <> "" Then strSQL = strSQL & "AND product_number = ('" & ProductNumber & "') "
                strSQL = strSQL & "ORDER BY product_release DESC"

                Dim returnCode = QueryNetworkDB(strSQL, rsTemp)
                If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode

                If (If(rsTemp?.Rows?.Count, 0)) = 1 Then
                    localUnitID = KillNullInteger(rsTemp(0)("unit_id"))
                ElseIf (If(rsTemp?.Rows?.Count, 0)) = 0 Then
                    LogError(New Exception($"Unit doesn't exist: {SerialNumber}"))
                    Return ReturnCodes.UDBS_OP_FAIL
                Else
                    'if there are duplicates, take the last one but issue a debug warning
                    localUnitID = KillNullInteger(rsTemp(0)("unit_id"))
                    LogError(New Exception($"Unit is non-unique: {SerialNumber}"))
                End If

                Return LoadUnitByID(localUnitID)

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        Public Sub UnloadUnit()

            mfamily_name = Nothing
            mproduct_id = Nothing
            mproduct_number = Nothing
            mproduct_release = Nothing
            munit_id = Nothing
            munit_serial_number = Nothing
            munit_created_by = Nothing
            munit_created_date = Nothing

            mLoaded = False
        End Sub

        ' Candidate for removal.
        Private Function LoadUnitByID(UnitID As Integer) As ReturnCodes
            ' Load info from database into object
            Dim strSQL As String
            Dim rsTemp As New DataTable

            Try
                UnloadUnit()

                strSQL = "SELECT family_name, product_number, product_release, unit.* " &
                         "FROM family with(nolock) , product with(nolock) , unit  with(nolock) " &
                         "WHERE product_family_id = family_id AND unit_product_id = product_id " &
                         "AND unit_id = " & UnitID

                Dim returnCode = QueryNetworkDB(strSQL, rsTemp)
                If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode

                If (If(rsTemp?.Rows?.Count, 0)) = 1 Then
                    mfamily_name = KillNull(rsTemp(0)("family_name"))
                    mproduct_number = KillNull(rsTemp(0)("product_number"))
                    mproduct_release = KillNullInteger(rsTemp(0)("product_release"))
                    mproduct_id = KillNullInteger(rsTemp(0)("unit_product_id"))
                    munit_id = KillNullInteger(rsTemp(0)("unit_id"))
                    munit_serial_number = KillNull(rsTemp(0)("unit_serial_number"))
                    munit_created_by = KillNull(rsTemp(0)("unit_created_by"))
                    munit_created_date = KillNullDate(rsTemp(0)("unit_created_date"))
                    munit_create_report = KillNull(rsTemp(0)("unit_report"))

                    mLoaded = True
                Else
                    LogError(New Exception($"Couldn't load unit with Id: {UnitID}."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ' Candidate for removal.
        Private Function GetUnitDuplicates(SerialNumber As String,
                                          ByRef NumUnits As Integer,
                                          ByRef UnitIDs() As Integer,
                                          Optional ByVal FamilyName As String = "",
                                          Optional ByVal product_number As String = "") _
            As ReturnCodes
            'checks to see if the unit is already in the database - and the number of duplicates if any
            'if a FamilyName is passed, the function checks only within that family
            'if a product number is passed, the function checks only within that product

            Dim rsTemp As DataTable = Nothing
            Dim SQLstr As String
            Dim i As Integer

            Try
                SQLstr = "SELECT family_name, product_number, unit_id, unit_serial_number " &
                         "FROM family with(nolock) , product with(nolock) , unit with(nolock)  " &
                         "WHERE product_family_id = family_id " &
                         "AND unit_product_id = product_id " &
                         "AND unit_serial_number = '" & SerialNumber & "'"

                'ignore duplicates outside the family
                If FamilyName <> "" Then
                    SQLstr = SQLstr & " AND family_name = '" & FamilyName & "'"
                End If

                'ignore duplicates outside the product
                If product_number <> "" Then
                    SQLstr = SQLstr & " AND product_number = '" & product_number & "'"
                End If

                GetUnitDuplicates = QueryNetworkDB(SQLstr, rsTemp)
                If GetUnitDuplicates <> ReturnCodes.UDBS_OP_SUCCESS Then Exit Function

                NumUnits = (If(rsTemp?.Rows?.Count, 0))
                ReDim UnitIDs(NumUnits - 1)
                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                    For i = 0 To NumUnits - 1
                        UnitIDs(i) = KillNullInteger(rsTemp(i)("unit_id"))

                    Next
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                GetUnitDuplicates = ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' Logs the UDBS error in the application log and inserts it in the udbs_error table, with the process instance details:
        ''' Process Type, name, Process ID, UDBS product ID, Unit serial number.
        ''' </summary>
        ''' <param name="ex">Exception raised.</param>
        Private Sub LogErrorInDatabase(ex As Exception)

            DatabaseSupport.LogErrorInDatabase(ex, PROCESS, "Unit", UnitID, ProductNumber, SerialNumber)

        End Sub

#Region "IDisposable Support"

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' Destroys collection when this class is terminated
                    CloseNetworkDB()
                    mLoaded = False
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
