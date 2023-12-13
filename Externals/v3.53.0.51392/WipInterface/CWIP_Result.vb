Option Explicit On
Option Compare Binary
Option Infer On
Option Strict On

Namespace WipInterface
    Public Class CWIP_Result
        Private Const ClsName = "CWIP_Result"

        'class variables
        'item sub-object

        'local variable(s) to hold property value(s)

        'item sub-object
        Public ReadOnly Property Item As CWIP_Item

        'data properties
        Public ReadOnly Property ID As Long

        Public ReadOnly Property ProcessID As Integer

        Public ReadOnly Property ItemlistDefID As Integer

        Public ReadOnly Property StepNumber As Integer

        Public ReadOnly Property AuthorizedBy As String

        Public ReadOnly Property StartDate As Date

        Public ReadOnly Property EmployeeNumber As String

        Public ReadOnly Property Station As String

        Public ReadOnly Property UDBSProcessID As Integer

        Public ReadOnly Property EndDate As Date

        Public ReadOnly Property Passflag As WIPResultCodes

        Public ReadOnly Property ActiveDuration As Double

        Public ReadOnly Property InactiveDuration As Double

        Public ReadOnly Property WIPNotes As String

        Public ReadOnly Property BlobDataExists As Integer


        '**********************************************************************
        '* Methods
        '**********************************************************************
        Friend Sub New(ItemListItem As CWIP_Item, result_id As Long, result_process_id As Integer,
                       result_itemlistdef_id As Integer, result_step_number As Integer, result_authorized_by As String,
                       result_start_date As Date,
                       result_employee_number As String, result_station As String, result_udbs_process_id As Integer,
                       result_end_date As Date,
                       result_passflag As WIPResultCodes, result_inactive_duration As Double,
                       result_active_duration As Double, result_wip_notes As String,
                       result_blobdata_exists As Integer)

            ID = result_id
            ProcessID = result_process_id
            ItemlistDefID = result_itemlistdef_id
            StepNumber = result_step_number
            AuthorizedBy = result_authorized_by
            StartDate = result_start_date
            EmployeeNumber = result_employee_number
            Station = result_station
            UDBSProcessID = result_udbs_process_id
            EndDate = result_end_date
            Passflag = result_passflag
            InactiveDuration = result_inactive_duration
            ActiveDuration = result_active_duration
            WIPNotes = result_wip_notes
            BlobDataExists = result_blobdata_exists

            Item = ItemListItem
        End Sub
    End Class
End Namespace
