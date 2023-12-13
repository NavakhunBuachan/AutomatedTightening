Option Explicit On
Option Compare Binary
Option Infer On
Option Strict On

Namespace WipInterface
    Public Class CWIP_Item
        Private Const ClsName = "CWIP_Item"



        ' Properties (Read Only)
        Public ReadOnly Property ID As Integer

        Public ReadOnly Property Number As Integer

        Public ReadOnly Property Name As String

        Public ReadOnly Property Descriptor As String

        Public ReadOnly Property Description As String

        Public ReadOnly Property RequiredStep As Integer

        Public ReadOnly Property ProcessName As String

        Public ReadOnly Property StageName As String

        Public ReadOnly Property Role As String

        Public ReadOnly Property PassRouting As String

        Public ReadOnly Property FailRouting As String

        Public ReadOnly Property Automated As Integer

        Public ReadOnly Property Oracle_Routing As Integer

        Public ReadOnly Property BlobDataExists As Integer

        '**********************************************************************
        '* Methods
        '**********************************************************************
        Public Sub New(itemlistdef_id As Integer, itemlistdef_itemnumber As Integer, itemlistdef_itemname As String,
                       itemlistdef_descriptor As String, itemlistdef_description As String,
                       itemlistdef_required_step As Integer,
                       itemlistdef_processname As String, itemlistdef_stagename As String, itemlistdef_role As String,
                       itemlistdef_pass_routing As String,
                       itemlistdef_fail_routing As String, itemlistdef_automated_process As Integer,
                       itemlistdef_oracle_routing As Integer,
                       itemlistdef_blobdata_exists As Integer)

            ' Function populates the item object,
            ID = itemlistdef_id
            Number = itemlistdef_itemnumber
            Name = itemlistdef_itemname
            Descriptor = itemlistdef_descriptor
            Description = itemlistdef_description
            RequiredStep = itemlistdef_required_step
            ProcessName = itemlistdef_processname
            StageName = itemlistdef_stagename
            Role = itemlistdef_role
            PassRouting = itemlistdef_pass_routing
            FailRouting = itemlistdef_fail_routing
            Automated = itemlistdef_automated_process
            Oracle_Routing = itemlistdef_oracle_routing
            BlobDataExists = itemlistdef_blobdata_exists
        End Sub
    End Class
End Namespace
