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
    Public Class CWIP_ItemList
        Implements IDisposable

        Private Const ClsName = "CWIP_ItemList"

        'class variables
        Private ReadOnly mUt As New CWIP_Utility

        'Items Collection
        Private mItems As New Dictionary(Of String, CWIP_Item)

        'Object State
        Private mLoaded As Boolean

        'Object Data
        Private mfamily_name As String
        Private mproduct_id As Integer
        Private mproduct_number As String
        Private mproduct_release As Integer
        Private mitemlistrev_id As Integer
        Private mitemlistrev_stage As String
        Private mitemlistrev_revision As Integer
        Private mitemlistrev_description As String
        Private mitemlistrev_created_date As Date
        Private mitemlistrev_employee_number As String
        Private mitemlistrev_unit_info As String
        Private mitemlistrev_blobdata_exists As Integer

        '************************************************************************************************************
        ' PROPERTIES
        '************************************************************************************************************
        'Items Collection
        Public ReadOnly Property Items As Dictionary(Of String, CWIP_Item)
            Get
                If mItems Is Nothing Then
                    mItems = New Dictionary(Of String, CWIP_Item)
                End If
                Return mItems
            End Get
        End Property

        'object State
        Public ReadOnly Property Loaded As Boolean
            Get
                Return mLoaded
            End Get
        End Property

        'Object data
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

        Public ReadOnly Property ProductRelease As Integer
            Get
                Return mproduct_release
            End Get
        End Property

        ' ItemList Revision Information
        Public ReadOnly Property ID As Integer
            Get
                Return mitemlistrev_id
            End Get
        End Property

        Public ReadOnly Property Stage As String
            Get
                Return mitemlistrev_stage
            End Get
        End Property

        Public ReadOnly Property Revision As Integer
            Get
                Return mitemlistrev_revision
            End Get
        End Property

        Public ReadOnly Property Description As String
            Get
                Return mitemlistrev_description
            End Get
        End Property

        Public ReadOnly Property CreatedDate As Date
            Get
                Return mitemlistrev_created_date
            End Get
        End Property

        Public ReadOnly Property EmployeeNumber As String
            Get
                Return mitemlistrev_employee_number
            End Get
        End Property

        Public ReadOnly Property UnitInfo As String
            Get
                Return mitemlistrev_unit_info
            End Get
        End Property

        Public ReadOnly Property BlobDataExists As Boolean
            Get
                Return CBool(mitemlistrev_blobdata_exists)
            End Get
        End Property

        '**********************************************************************
        '* Methods
        '**********************************************************************
        Private Sub UnloadItemlist()

            mItems = Nothing

            mLoaded = False
            'Object Data
            mfamily_name = Nothing
            mproduct_id = Nothing
            mproduct_number = Nothing
            mproduct_release = Nothing
            mitemlistrev_id = Nothing
            mitemlistrev_stage = Nothing
            mitemlistrev_revision = Nothing
            mitemlistrev_description = Nothing
            mitemlistrev_created_date = Nothing
            mitemlistrev_employee_number = Nothing
            mitemlistrev_unit_info = Nothing
            mitemlistrev_blobdata_exists = Nothing
        End Sub

        ''' <param name="ProductNumber">API Error: This shouldn't be ByRef. This is an input parameter.</param>
        Public Function LoadItemList(ByRef ProductNumber As String,
                                     ByRef ProductRelease As Integer,
                                     ByRef Stage As String,
                                     ByRef Revision As Integer) _
            As ReturnCodes
            ' Load itemlist by Product Number and Release
            Dim strSQL As String
            Dim rsTemp As New DataTable
            Dim localProductID As Integer
            Dim returnCode As ReturnCodes

            Try
                UnloadItemlist()

                'get the product id
                'get latest release (if release was zero)
                If ProductRelease = 0 Then
                    returnCode = mUt.GetMaxProductRelease(ProductNumber, ProductRelease)
                    If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode
                End If

                strSQL = "SELECT product_id " &
                         "FROM product with(nolock) " &
                         "WHERE product_number = '" & ProductNumber & "' " &
                         "AND product_release = " & ProductRelease

                returnCode = QueryNetworkDB(strSQL, rsTemp)
                If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode

                If (If(rsTemp?.Rows?.Count, 0)) = 1 Then
                    localProductID = KillNullInteger(rsTemp(0)("product_id"))
                Else
                    'the product is not found or is not uniquely identified
                    LogError(New Exception($"Couldn't find product '{ProductNumber}', release '{ProductRelease}'"))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                'now get the revision info
                'get the latest revision (if none specified)
                If Revision = 0 Then
                    returnCode = mUt.GetMaxItemlistRevision(localProductID, Stage, Revision)
                    If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode
                End If

                strSQL = "SELECT itemlistrev_id " &
                         "FROM WIP_itemlistrevision with(nolock) " &
                         "WHERE itemlistrev_product_id = " & localProductID & " " &
                         "AND itemlistrev_stage = '" & Stage & "' " &
                         "AND itemlistrev_revision = " & Revision

                returnCode = QueryNetworkDB(strSQL, rsTemp)
                If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode

                'now load all the itemlist data
                If (If(rsTemp?.Rows?.Count, 0)) = 1 Then
                    returnCode = LoadItemListByID(KillNullInteger(rsTemp(0)("itemlistrev_id")))
                    If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode
                Else
                    LogError(New Exception("Couldn't find itemlist."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        Public Function LoadItemListByID(ItemlistRevID As Integer) As ReturnCodes
            ' Load info from database into object
            Dim strSQL As String
            Dim rsTemp As New DataTable
            Dim returnCode As ReturnCodes

            Try
                UnloadItemlist()

                strSQL = "SELECT family_name, product_number, product_release, WIP_itemlistrevision.* " &
                         "FROM family with(nolock), product with(nolock), WIP_itemlistrevision with(nolock) " &
                         "WHERE product_family_id = family_id AND itemlistrev_product_id = product_id " &
                         "AND itemlistrev_id = " & ItemlistRevID

                returnCode = QueryNetworkDB(strSQL, rsTemp)
                If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode

                If (If(rsTemp?.Rows?.Count, 0)) = 1 Then
                    mfamily_name = KillNull(rsTemp(0)("family_name"))
                    mproduct_number = KillNull(rsTemp(0)("product_number"))
                    mproduct_release = KillNullInteger(rsTemp(0)("product_release"))
                    mproduct_id = KillNullInteger(rsTemp(0)("itemlistrev_product_id"))
                    mitemlistrev_id = KillNullInteger(rsTemp(0)("itemlistrev_id"))
                    mitemlistrev_stage = KillNull(rsTemp(0)("itemlistrev_stage"))
                    mitemlistrev_revision = KillNullInteger(rsTemp(0)("itemlistrev_revision"))
                    mitemlistrev_description = KillNull(rsTemp(0)("itemlistrev_description"))
                    mitemlistrev_created_date = KillNullDate(rsTemp(0)("itemlistrev_created_date"))
                    mitemlistrev_employee_number = KillNull(rsTemp(0)("itemlistrev_created_by"))
                    mitemlistrev_unit_info = KillNull(rsTemp(0)("itemlistrev_unit_info"))
                    mitemlistrev_blobdata_exists = KillNullInteger(rsTemp(0)("itemlistrev_blobdata_exists"))
                    ' Populate Items collection
                    mLoaded = True

                    If LoadItemsCollection() <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If
                Else
                    LogError(New Exception($"Couldn't load itemlist for revision id: {ItemlistRevID}."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function


        '**********************************************************************
        '* Support Functions
        '**********************************************************************
        Private Function LoadItemsCollection() As ReturnCodes
            ' Function creates the items collection for this process instance object
            Dim strSQL As String
            Dim rsTemp As New DataTable
            Dim returnCode As ReturnCodes

            Try
                If Not mLoaded Then
                    LogError(New Exception("Cannot load Items collection when Itemlist is not loaded."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                mItems = New Dictionary(Of String, CWIP_Item)

                strSQL = "SELECT * FROM WIP_itemlistdefinition with(nolock) " &
                     "WHERE itemlistdef_itemlistrev_id = " & mitemlistrev_id & " " &
                     "ORDER BY itemlistdef_itemnumber ASC"

                returnCode = QueryNetworkDB(strSQL, rsTemp)
                If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Return returnCode

                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then

                    For Each dr As DataRow In rsTemp.Rows
                        Dim tmpItem = New CWIP_Item(KillNullInteger(dr("itemlistdef_id")),
                                                KillNullInteger(dr("itemlistdef_itemnumber")),
                                                KillNull(dr("itemlistdef_itemname")),
                                                KillNull(dr("itemlistdef_descriptor")),
                                                KillNull(dr("itemlistdef_description")),
                                                KillNullInteger(dr("itemlistdef_required_step")),
                                                KillNull(dr("itemlistdef_processname")),
                                                KillNull(dr("itemlistdef_stagename")), KillNull(dr("itemlistdef_role")),
                                                KillNull(dr("itemlistdef_pass_routing")),
                                                KillNull(dr("itemlistdef_fail_routing")),
                                                KillNullInteger(dr("itemlistdef_automated_process")),
                                                KillNullInteger(dr("itemlistdef_oracle_routing")),
                                                KillNullInteger(dr("itemlistdef_blobdata_exists")))

                        mItems.Add(KillNull(dr("itemlistdef_itemname")), tmpItem)

                    Next
                    Return ReturnCodes.UDBS_OP_SUCCESS
                Else
                    LogError(New Exception("No items found in itemlist."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        Public Function GetItemName(ItemNumber As Integer) As String
            ' Function returns the item name of the specified item number
            Try
                Return mItems.First(Function(f) f.Value.Number = ItemNumber).Key
            Catch ex As Exception
                Throw New Exception("Item number " & ItemNumber & " does not exist in current itemlist.")
            End Try
        End Function

        Public Function GetItemByNumber(ItemNumber As Integer) As CWIP_Item
            ' Function returns the item name of the specified item number
            Try
                Return mItems.First(Function(f) f.Value.Number = ItemNumber).Value
            Catch ex As Exception
                Throw New Exception("Item number " & ItemNumber & " does not exist in current itemlist.")
            End Try
        End Function

        ''' <summary>
        ''' Logs the UDBS error in the application log and inserts it in the udbs_error table, with the process instance details:
        ''' Process Type, name, Process ID, UDBS product ID, Unit serial number.
        ''' </summary>
        ''' <param name="ex">Exception raised.</param>
        Private Sub LogErrorInDatabase(ex As Exception)

            DatabaseSupport.LogErrorInDatabase(ex, "WIP", Stage, ID, ProductNumber, String.Empty)

        End Sub

#Region "IDisposable Support"

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    mItems = Nothing
                    mLoaded = False
                    ' Class Destructor
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

