Option Explicit On
Option Compare Text
Option Infer On
Option Strict On

Namespace MasterInterface
    Friend Class CItemlist
        Implements IDisposable

        ' Table Identification
        Private mProcess As String
        Private mUDBSProcessID As Integer = -1
        Private mBLOBTable As String
        Private mItemListRevisionTable As String
        Private mItemListDefinitionTable As String
        Private mProductTable As String
        Private mEquivalencyTable As String

        ' Object State
        Private mREADONLY As Boolean
        Private mObjectLoaded As Boolean

        ' ItemList Revision Information
        Private mItemListRevID As Integer
        Private mStage As String
        Private mRevision As Integer
        Private mRevisionDescription As String
        Private mCreatedDate As Date
        Private mEmployeeNumber As String
        Private mBlobDataExists As Integer
        Private mItemListRevision As New DataTable
        Private mItemListDefinition As New DataTable

        ' Product Information
        Protected mPRODUCT As CProduct

        ' Error Handling

        '**********************************************************************
        '* Properties
        '**********************************************************************

        ' Object Information

        ''' <summary>
        ''' Property specifies whether the ItemList was loaded.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ItemListLoaded As Boolean
            Get
                Return mObjectLoaded
            End Get
        End Property

        Public ReadOnly Property ProcessID As Integer
            Get
                If mUDBSProcessID < 0 Then
                    Using aUtility As New CUtility
                        ' Get the UDBSProcessID
                        If aUtility.UDBS_GetUDBSProcessID(Process, mUDBSProcessID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                            Return -1
                        End If
                    End Using
                End If

                Return mUDBSProcessID
            End Get
        End Property

        Public ReadOnly Property Process As String
            Get
                Return mProcess
            End Get
        End Property


        ' Product Information
        Public ReadOnly Property ProductID As Integer
            Get
                Return mPRODUCT.ProductID
            End Get
        End Property

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


        ' ItemList Revision Information
        Public ReadOnly Property ItemListRevID As Integer
            Get
                Return mItemListRevID
            End Get
        End Property

        Public ReadOnly Property Stage As String
            Get
                Return mStage
            End Get
        End Property

        Public ReadOnly Property Revision As Integer
            Get
                Return mRevision
            End Get
        End Property

        Public ReadOnly Property RevisionDescription As String
            Get
                Return mRevisionDescription
            End Get
        End Property

        Public ReadOnly Property CreatedDate As Date
            Get
                Return mCreatedDate
            End Get
        End Property

        Public ReadOnly Property EmployeeNumber As String
            Get
                Return mEmployeeNumber
            End Get
        End Property

        ''' <remarks>
        ''' This looks like a boolean property, but it is actually a numeric value.
        ''' TODO: Make this a boolean property.
        ''' </remarks>
        Public ReadOnly Property BlobDataExists As Integer
            Get
                Return mBlobDataExists
            End Get
        End Property

        Public ReadOnly Property ItemListRevision_RS As DataTable
            Get
                Return mItemListRevision.Copy()
            End Get
        End Property

        Public ReadOnly Property Items_RS As DataTable
            Get
                Return mItemListDefinition.Copy()
            End Get
        End Property

        '**********************************************************************
        '* Methods
        '**********************************************************************

        ' Function returns Stage information for specified Product/Release/Stage/Revision
        Public Function LoadItemList(ProcessName As String,
                                     ProductNumber As String,
                                     ProductRelease As Integer,
                                     Stage As String,
                                     ByRef Revision As Integer) _
            As ReturnCodes

            ' Load product object
            If mPRODUCT.GetProduct(ProductNumber, ProductRelease) <> ReturnCodes.UDBS_OP_SUCCESS Then
                ' Could not load product object
                Return ReturnCodes.UDBS_ERROR
            End If

            Return LoadItemList(ProcessName, mPRODUCT, Stage, Revision)
        End Function

        ' Function returns Stage information for specified Product/Release/Stage/Revision
        Public Function LoadItemList(ProcessName As String,
                                     Product As CProduct,
                                     Stage As String,
                                     ByRef Revision As Integer) _
            As ReturnCodes

            mPRODUCT = Product

            Dim UTILITY As New CUtility

            Try
                If mObjectLoaded = True Then
                    ' Cannot reload an itemlist.. Must create new IL object
                    LogError(New Exception("Could not reload itemlist."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                InitializeTableNames(ProcessName)

                ' Has calling function specified latest revision?
                If Revision = 0 Then
                    ' Return the latest revision
                    If UTILITY.ItemList_GetRevisionCount(ProcessName, mPRODUCT.Number, mPRODUCT.Release, Stage, Revision, mItemListRevID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        ' The specified itemlist does not exist
                        mObjectLoaded = False
                        Return ReturnCodes.UDBS_ERROR
                    End If
                Else
                    ' In the IF block, we retrieve the 'latest' revision and its matching ID.
                    ' If the revision is specified, we have to retrieve the ID.
                    ' TODO: Wouldn't there be a way to retrieve this at the same time as the
                    '       revision?
                    '
                    ' Get the revision ID for specified item list.
                    If GetItemListRevID(mPRODUCT.Number, mPRODUCT.Release, Stage, Revision, mItemListRevID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Return ReturnCodes.UDBS_ERROR
                    End If

                End If

                If Revision = -1 Then
                    ' No itemlist exists...
                    Return ReturnCodes.UDBS_ERROR
                End If

                Return LoadItemListByID(ProcessName, mItemListRevID)
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function


        ' Function returns Stage Information for the specified stage id
        Public Function LoadItemListByID(ProcessName As String,
                                         ItemListRevisionID As Integer) _
            As ReturnCodes

            If mObjectLoaded = True Then
                ' Cannot reload an itemlist.. Must create new IL object
                LogError(New Exception("Could not reload itemlist."))
                Return ReturnCodes.UDBS_ERROR
            End If

            ' Object has been created for read mode
            mREADONLY = True

            Dim UTILITY As New CUtility
            Dim sqlQuery As String
            Dim localProdId As Integer

            Dim iItemCount As Integer

            InitializeTableNames(ProcessName)
            InitializeObjectProperties()

            Try
                ' Get ItemList Revision Information
                sqlQuery = "SELECT * FROM " & mItemListRevisionTable &
                           " with(nolock) WHERE itemlistrev_id = " & CStr(ItemListRevisionID)
                If QueryNetworkDB(sqlQuery, mItemListRevision) <> ReturnCodes.UDBS_OP_SUCCESS _
                        OrElse mItemListRevision?.Rows?.Count <> 1 Then
                    LogError(New Exception("Error in retrieving itemlist."))
                    mObjectLoaded = False
                    Return ReturnCodes.UDBS_ERROR
                Else
                    ' Item list found.
                    ' Fill the object properties
                    mItemListRevID = KillNullInteger(mItemListRevision(0)("itemlistrev_id"))
                    mStage = KillNull(mItemListRevision(0)("itemlistrev_stage"))
                    mRevision = KillNullInteger(mItemListRevision(0)("itemlistrev_revision"))

                    If Not IsDBNull(mItemListRevision(0)("itemlistrev_description")) Then
                        mRevisionDescription = KillNull(mItemListRevision(0)("itemlistrev_description"))
                    End If

                    If Not IsDBNull(mItemListRevision(0)("itemlistrev_created_date")) Then
                        mCreatedDate = KillNullDate(mItemListRevision(0)("itemlistrev_created_date"))
                    End If

                    If Not IsDBNull(mItemListRevision(0)("itemlistrev_created_by")) Then
                        mEmployeeNumber = KillNull(mItemListRevision(0)("itemlistrev_created_by"))
                    End If

                    If Not IsDBNull(mItemListRevision(0)("itemlistrev_blobdata_exists")) Then
                        mBlobDataExists = KillNullInteger(mItemListRevision(0)("itemlistrev_blobdata_exists"))
                    End If

                    localProdId = KillNullInteger(mItemListRevision(0)("itemlistrev_product_id"))

                    If Not IsDBNull(mItemListRevision(0)("itemlistrev_reportdef")) Then
                        iItemCount = KillNullInteger(Val(mItemListRevision(0)("itemlistrev_reportdef")))
                    Else
                        iItemCount = -1
                    End If
                End If

                ' Retrieve the items for this stage
                sqlQuery = "SELECT * FROM " & mItemListDefinitionTable &
                           " with(nolock) WHERE itemlistdef_itemlistrev_id = " & CStr(mItemListRevID) &
                           " ORDER BY itemlistdef_itemnumber"
                If QueryNetworkDB(sqlQuery, mItemListDefinition) <> ReturnCodes.UDBS_OP_SUCCESS _
                        OrElse mItemListDefinition?.Rows?.Count = 0 Then
                    ' There is a problem finding these items
                    LogError(New Exception("Item not found."))
                    mObjectLoaded = False
                    Return ReturnCodes.UDBS_ERROR
                Else
                    ' Pass back the list of items
                    If iItemCount <> -1 Then
                        If (If(mItemListDefinition?.Rows?.Count, 0)) <> iItemCount Then
                            ' incorrect checksum
                            LogError(New Exception($"Incorrect checksum: rev=" & iItemCount & " def=" &
                                         (If(mItemListDefinition?.Rows?.Count, 0))))
                            mObjectLoaded = False
                            Return ReturnCodes.UDBS_ERROR
                        End If
                    End If
                End If

                ' Load product object
                If Not mPRODUCT.Loaded Then
                    If mPRODUCT.GetProductByID(localProdId) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        ' Could not load product object
                        Return ReturnCodes.UDBS_ERROR
                    End If
                End If

                ' Itemlist object is properly loaded
                mObjectLoaded = True

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Not used
        ''' </summary>
        Public Function ReleaseItemListFromDebugMode() _
            As ReturnCodes
            ' Function releases a Stage from debug mode to first revision.=
            Try
                If mObjectLoaded = False Or mRevision <> 0 Then
                    ' Must load a debug list before it can be released
                    LogError(New Exception("No itemlist loaded."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                Dim sqlQuery As String


                OpenNetworkDB(120)
                ' Change the revision from 0 to 1
                sqlQuery = "UPDATE " & mItemListRevisionTable &
                           " SET itemlistrev_revision = 1 WHERE itemlistrev_id = " & CStr(mItemListRevID) &
                           " AND itemlistrev_revision=0"
                ExecuteNetworkQuery(sqlQuery)

                ' Update object properties to reflect release
                mObjectLoaded = False
                If LoadItemListByID(mProcess, mItemListRevID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    logger.Error($"Failed to reload item list {Process} ID {mItemListRevID}")
                    Return ReturnCodes.UDBS_ERROR
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

        Private Function GetItemListRevID(ProductNumber As String,
                                          ProductRelease As Integer,
                                          Stage As String,
                                          Revision As Integer,
                                          ByRef ItemListRevID As Integer) _
            As ReturnCodes
            ' Function returns the id key of the specified Product/Release/Stage/Revision

            Dim rsTemp As New DataTable
            Dim sqlQuery As String
            Dim returnValue As ReturnCodes = ReturnCodes.UDBS_OP_SUCCESS

            Try
                OpenNetworkDB(120)

                sqlQuery = "SELECT itemlistrev_id FROM " & mItemListRevisionTable &
                           " with(nolock) WHERE itemlistrev_product_id IN (SELECT product_id FROM " & mProductTable &
                           " with(nolock) WHERE product_number = '" & ProductNumber & "' " &
                           "AND product_release=" & CStr(ProductRelease) & ") AND itemlistrev_stage='" & Stage & "' " &
                           "AND itemlistrev_revision=" & CStr(Revision)
                OpenNetworkRecordSet(rsTemp, sqlQuery)
                If (If(rsTemp?.Rows?.Count, 0)) = 1 Then
                    ' ItemList found
                    ItemListRevID = KillNullInteger(rsTemp(0)("itemlistrev_id"))
                    returnValue = ReturnCodes.UDBS_OP_SUCCESS
                Else
                    ' There is a problem finding this itemlist
                    LogError(New Exception("Itemlist not found."))
                    returnValue = ReturnCodes.UDBS_ERROR
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                returnValue = ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try

            Return returnValue
        End Function


        ''' <summary>
        ''' Function loads the appropriate table names for this process
        ''' </summary>
        ''' <param name="ProcessName">The name of the process.</param>
        Private Sub InitializeTableNames(ProcessName As String)
            mProcess = Trim(LCase(ProcessName))
            mProductTable = "product"
            mEquivalencyTable = "udbs_equivalency"
            mItemListRevisionTable = mProcess & "_itemlistrevision"
            mItemListDefinitionTable = mProcess & "_itemlistdefinition"
            mBLOBTable = mProcess & "_blob"
        End Sub

        Private Sub InitializeObjectProperties()
            mItemListRevID = 0
            mStage = ""
            mRevision = 0
            mRevisionDescription = ""
            mCreatedDate = Now
            mEmployeeNumber = ""
            mBlobDataExists = 0
        End Sub

        ''' <summary>
        ''' Logs the UDBS error in the application log and inserts it in the udbs_error table, with the process instance details:
        ''' Process Type, name, Process ID, UDBS product ID, Unit serial number.
        ''' </summary>
        ''' <param name="ex">Exception raised.</param>
        Private Sub LogErrorInDatabase(ex As Exception)

            'The Process Type is just set to the process name for this class as the type is unidentified as this point.
            DatabaseSupport.LogErrorInDatabase(ex, contextType:=Process, name:=Process, processID:=ProcessID, product:=ProductNumber, serialNumber:=String.Empty)

        End Sub


        '**********************************************************************
        '* Class Constructor/Destructor
        '**********************************************************************

        Public Sub New()

            ' create a product object for misc. use
            mPRODUCT = New CProduct

            ' An empty itemlist object can be used to create a new itemlist
            mREADONLY = False
            mObjectLoaded = False
        End Sub

#Region "IDisposable Support"

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    CloseNetworkDB()
                    ' destroy misc use product object
                    mPRODUCT = Nothing

                    ' An empty itemlist object can be used to create a new itemlist
                    mREADONLY = False
                    mObjectLoaded = False
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
