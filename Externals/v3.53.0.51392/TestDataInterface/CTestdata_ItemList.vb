Option Explicit On
Option Strict On
Option Compare Text
Option Infer On

Imports UdbsInterface.MasterInterface


Namespace TestDataInterface
    Public Class CTestData_ItemList
        Implements IDisposable

        ' ItemList Object
        Private ReadOnly mIL As New CItemlist

        ' Items Collection
        Private mItems As New Dictionary(Of String, CTestData_Item)

        ' Object State
        Private mReadOnly As Boolean
        Private mItemListLoaded As Boolean

        ' Temporary Product information used to build new lists
        Private mRelease As Integer
        Private mStage As String
        Private mDescription As String
        Private mEmployeeNumber As String
        Private mReportDef As String


        ' Item Collection
        Public ReadOnly Property Items As Dictionary(Of String, CTestData_Item)
            Get
                If mItems Is Nothing Then
                    mItems = New Dictionary(Of String, CTestData_Item)
                End If
                Return mItems
            End Get
        End Property

#Region "Properties"

        ' Product Information
        Public ReadOnly Property ProductID As Integer
            Get
                Return mIL.ProductID
            End Get
        End Property

        Public ReadOnly Property ProductNumber As String
            Get
                Return mIL.ProductNumber
            End Get
        End Property

        Public ReadOnly Property ProductRelease As Double
            Get
                Return mIL.ProductRelease
            End Get
        End Property

        Public ReadOnly Property ProductDescriptor As String
            Get
                Return mIL.ProductDescriptor
            End Get
        End Property

        Public ReadOnly Property ProductDescription As String
            Get
                Return mIL.ProductDescription
            End Get
        End Property

        Public ReadOnly Property ProductCreatedBy As String
            Get
                Return mIL.ProductCreatedBy
            End Get
        End Property

        Public ReadOnly Property ProductCreatedDate As Date
            Get
                Return mIL.ProductCreatedDate
            End Get
        End Property

        Public ReadOnly Property ProductReleaseReason As String
            Get
                Return mIL.ProductReleaseReason
            End Get
        End Property

        Public ReadOnly Property ProductSNProdCode As String
            Get
                Return mIL.ProductSNProdCode
            End Get
        End Property

        Public ReadOnly Property ProductSNTemplate As String
            Get
                Return mIL.ProductSNTemplate
            End Get
        End Property

        Public ReadOnly Property ProductSNLastUnit As Integer
            Get
                Return mIL.ProductSNLastUnit
            End Get
        End Property

        Public ReadOnly Property ProductFamily As String
            Get
                Return mIL.ProductFamily
            End Get
        End Property

        ' ItemList Revision Information
        Public ReadOnly Property ItemListRevID As Integer
            Get
                Return mIL.ItemListRevID
            End Get
        End Property

        Public ReadOnly Property Stage As String
            Get
                Return mIL.Stage
            End Get
        End Property

        Public ReadOnly Property Revision As Integer
            Get
                Return mIL.Revision
            End Get
        End Property

        Public ReadOnly Property RevisionDescription As String
            Get
                Return mIL.RevisionDescription
            End Get
        End Property

        Public ReadOnly Property CreatedDate As Date
            Get
                Return mIL.CreatedDate
            End Get
        End Property

        Public ReadOnly Property EmployeeNumber As String
            Get
                Return mIL.EmployeeNumber
            End Get
        End Property

        Public ReadOnly Property BlobDataExists As Integer
            Get
                Return mIL.BlobDataExists
            End Get
        End Property

#End Region

#Region "Methods"


        ' Load itemlist by Product Number and Release
        Public Function LoadItemList(ProductNumber As String,
                                     ProductRelease As Double,
                                     Stage As String,
                                     ByRef Revision As Integer) As ReturnCodes
            If mIL.LoadItemList(PROCESS, ProductNumber, CInt(ProductRelease), Stage, Revision) = ReturnCodes.UDBS_OP_SUCCESS _
                    AndAlso LoadItemsCollection() = ReturnCodes.UDBS_OP_SUCCESS Then
                ' Populate Items collection
                mItemListLoaded = True
                mReadOnly = True
                Return ReturnCodes.UDBS_OP_SUCCESS
            Else
                mItemListLoaded = False
                mReadOnly = False
                Return ReturnCodes.UDBS_ERROR
            End If
        End Function

#End Region

#Region "Support Functions"


        Private Function LoadItemsCollection() As ReturnCodes
            ' Function creates the items collection for this process instance object
            Dim ItemNumber As Integer
            Dim ItemName As String
            Dim Descriptor As String
            Dim Description As String
            Dim ReportLevel As Integer
            Dim Units As String
            Dim CriticalSpec As Integer
            Dim WarnMin As Double
            Dim WarnMax As Double
            Dim FailMin As Double
            Dim FailMax As Double
            Dim SanityMin As Double
            Dim SanityMax As Double
            Dim BlobExits As Integer
            Dim rsTemp As New DataTable
            Try
                Items.Clear()
                rsTemp = mIL.Items_RS
                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then

                    For Each dr As DataRow In rsTemp.Rows
                        ItemNumber = KillNullInteger(dr("itemlistdef_itemnumber"))
                        ItemName = KillNull(dr("itemlistdef_itemname"))
                        Descriptor = KillNull(dr("itemlistdef_descriptor"))
                        Description = KillNull(dr("itemlistdef_description"))
                        ReportLevel = KillNullInteger(dr("itemlistdef_report_level"))
                        Units = KillNull(dr("itemlistdef_units"))
                        CriticalSpec = KillNullInteger(dr("itemlistdef_critical_spec"))
                        WarnMin = KillNullDouble(dr("itemlistdef_warning_min"))
                        WarnMax = KillNullDouble(dr("itemlistdef_warning_max"))
                        FailMin = KillNullDouble(dr("itemlistdef_fail_min"))
                        FailMax = KillNullDouble(dr("itemlistdef_fail_max"))
                        SanityMin = KillNullDouble(dr("itemlistdef_sanity_min"))
                        SanityMax = KillNullDouble(dr("itemlistdef_sanity_max"))

                        If Not IsDBNull(dr("itemlistdef_blobdata_exists")) Then
                            BlobExits = KillNullInteger(dr("itemlistdef_blobdata_exists"))
                        End If

                        Items.Add(ItemName,
                                  New CTestData_Item(ItemNumber, ItemName, Descriptor, Description, ReportLevel, Units,
                                                     CriticalSpec,
                                                     WarnMin, WarnMax, FailMin, FailMax, SanityMin, SanityMax))

                    Next
                End If

                ' Process the itemlist groups
                SetGroupFlags()
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_OP_FAIL
            End Try
        End Function

        Private Sub SetGroupFlags()
            ' Function walks through the results collection and sets the group flag for items that are groups
            For Each i In mItems
                i.Value.IsGroup = ItemIsGroup(i.Key)
            Next i
        End Sub

        Private Function ItemIsGroup(ItemName As String) _
            As Boolean
            Try

                'This function assumes that the item name actually exists!
                Dim NextItemNumber As Integer
                Dim HasImmediateChildren As Boolean
                Dim NoSpecs As Boolean
                Dim retval As Boolean

                'Move the recordset pointer to the item name we're concerned about...
                Dim tmpItem As CTestData_Item = mItems(ItemName)
                Dim NextItem As CTestData_Item = Nothing

                'Check to see if there is any specs on this item!
                If tmpItem.HasSpecs = False Then
                    NoSpecs = True
                Else
                    NoSpecs = False
                End If

                ' Get information on next item
                NextItemNumber = tmpItem.ItemNumber + 1

                If NextItemNumber > mItems.Count Then
                    ' This cannot be a group
                    retval = False
                Else
                    NextItem = mItems(GetItemName(NextItemNumber))

                    If NextItem.ReportLevel = tmpItem.ReportLevel + 1 Then
                        HasImmediateChildren = True
                    End If

                    If HasImmediateChildren And NoSpecs Then
                        retval = True
                    Else
                        retval = False
                    End If

                End If

                Return retval
            Catch ex As Exception
                Throw New UDBSException($"Failed to determine if item {ItemName} is a group.", ex)
            End Try
        End Function

        Public Function GetItemName(ItemNumber As Integer) As String
            ' Function returns the item name of the specified item number
            Try
                Return mItems.First(Function(f) f.Value.ItemNumber = ItemNumber).Key
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

            If mIL Is Nothing Then
                DatabaseSupport.LogErrorInDatabase(ex)
            Else
                DatabaseSupport.LogErrorInDatabase(ex, mIL.Process, mIL.Stage, mIL.ProductID, mIL.ProductNumber, String.Empty)
            End If

        End Sub

#End Region

#Region "IDisposable Support"

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    mIL.Dispose()
                    mItems = Nothing
                    mItemListLoaded = False
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
