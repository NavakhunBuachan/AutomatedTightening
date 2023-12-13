Option Explicit On
Option Strict On
Option Compare Text
Option Infer On

Imports UdbsInterface.MasterInterface

Namespace TestDataInterface
    Friend Module TestDataSupport
        Public mDebugMode As Boolean

        ' Constant Declarations
        Public Const C As Double = 2.99792458 * 10 ^ 8
        Public Const MIN_LONG As Long = -2147483647
        Public Const SHOW_NO_SPEC As String = "---"

        Public Function ExpandParents(ItemNumber As Integer,
                                      ByRef rsResultSet As DataTable) _
            As Long
            ' Set all previous parents to visible
            Dim ParentItem As Integer
            ParentItem = Convert.ToInt32(rsResultSet(ItemNumber)("itemlistdef_report_level"))
            Dim ctr As Integer = ItemNumber
            Do While ctr >= 0
                If Convert.ToInt32(rsResultSet(ctr)("itemlistdef_report_level")) < ParentItem Then
                    ParentItem = Convert.ToInt32(rsResultSet(ctr)("itemlistdef_report_level"))
                    ' Turn on this parent
                    rsResultSet(ctr)("Reported") = True
                End If
                ctr -= 1
            Loop
            ExpandParents = 1
        End Function

        Private Function FindColumn(Column As String,
                                    ByRef rsSearch As DataTable) _
            As Long

            ' Purpose: Look at a row of a recordset and determine if specified column is an element
            Dim x As Integer

            FindColumn = -1
            For x = 0 To rsSearch.Columns.Count - 1
                If rsSearch.Columns(x).ColumnName = Column Then
                    FindColumn = 1
                    Exit Function
                End If
            Next x
        End Function

        Public Function FindResultData(ItemName As String,
                                       Column As String,
                                       ByRef ItemValue As String,
                                       ByRef Flag As String,
                                       ByRef rsSearch As DataTable) _
            As Long


            Dim ctr = 0

            Dim dr = rsSearch.AsEnumerable().
                    FirstOrDefault(Function(z) If(z.Field(Of String)("itemlistdef_itemname"), "") = ItemName)
            If IsNothing(dr) Then
                ' no result data ?
                Return -1
            End If

            If Column = "" Then
                ItemValue = dr.Field(Of Double?)("result_value") & ""

                ' Mark result_value as warning, fail or ok
                Flag = "OK"
                If _
                    (If(dr.Field(Of Double?)("result_value"), Double.NaN) >
                     If(dr.Field(Of Double?)("itemlistdef_warning_max"), Double.NaN)) And
                    Convert.ToString(dr("itemlistdef_warning_max")) <> SHOW_NO_SPEC Then
                    Flag = "WARNING"
                End If

                If _
                    (If(dr.Field(Of Double?)("result_value"), Double.NaN) <
                     If(dr.Field(Of Double?)("itemlistdef_warning_min"), Double.NaN)) And
                    Convert.ToString(dr("itemlistdef_warning_min")) <> SHOW_NO_SPEC Then
                    Flag = "WARNING"
                End If

                If _
                    (If(dr.Field(Of Double?)("result_value"), Double.NaN) >
                     If(dr.Field(Of Double?)("itemlistdef_fail_max"), Double.NaN)) And
                    Convert.ToString(dr("itemlistdef_fail_max")) <> SHOW_NO_SPEC Then
                    Flag = "FAIL"
                End If

                If _
                    (If(dr.Field(Of Double?)("result_value"), Double.NaN) <
                     If(dr.Field(Of Double?)("itemlistdef_fail_min"), Double.NaN)) And
                    Convert.ToString(dr("itemlistdef_fail_min")) <> SHOW_NO_SPEC Then
                    Flag = "FAIL"
                End If

                FindResultData = 1
            Else
                If FindColumn(Column, rsSearch) = 1 Then
                    ItemValue = CStr(If(dr.Field(Of Double?)(Column), Double.NaN) & "")
                    FindResultData = 1
                Else
                    FindResultData = -1
                End If
            End If
        End Function


        Public Function FindProcessData(Column As String,
                                        ByRef ItemValue As String,
                                        ByRef rsSearch As DataTable) _
            As Long

            If FindColumn(Column, rsSearch) = 1 Then
                ItemValue = CStr(rsSearch(0)(Column))
                FindProcessData = 1
            Else
                FindProcessData = -1
            End If
        End Function


        Public Function Limits(PassValue As Object) _
            As String
            ' Function returns limit value as string, or dashed filler if limit is null.

            If String.IsNullOrEmpty(Convert.ToString(PassValue)) Then
                Return SHOW_NO_SPEC
            Else
                Return CStr(PassValue)
            End If
        End Function


        ' this function should be rewritten to use the enums instead of hard coding the
        ' pass flag values in here

        Public Function PassOrFail(PassValue As Long,
                                   FlagGroup As Boolean) _
            As String
            ' Function returns description of Pass/Fail enumerator


            Dim Temp As String

            If PassValue > 0 And PassValue < 1000 Then
                Temp = "PASS"
            ElseIf PassValue >= 1000 Then
                Temp = "INC/PASS"
            ElseIf PassValue = 0 Then
                Temp = ""
            ElseIf PassValue < 0 And PassValue > -20 Then
                Temp = "Error"
            ElseIf PassValue <= -20 And PassValue > -30 Then
                Temp = "WARNING"
                If FlagGroup = False Then
                    ' Give detailed result
                    Select Case PassValue
                        Case -21
                            Temp = Temp & " HIGH"
                        Case -22
                            Temp = Temp & " LOW"
                    End Select
                End If
            ElseIf PassValue <= -1020 And PassValue > -1030 Then
                Temp = "INC/WARNING"
                If FlagGroup = False Then
                    ' Give detailed result
                    Select Case PassValue
                        Case -1021
                            Temp = Temp & " HIGH"
                        Case -1022
                            Temp = Temp & " LOW"
                    End Select
                End If
            ElseIf PassValue <= -30 And PassValue > -40 Then
                Temp = "FAIL"
                If FlagGroup = False Then
                    ' Give detailed result
                    Select Case PassValue
                        Case -31
                            Temp = Temp & " HIGH"
                        Case -32
                            Temp = Temp & " LOW"
                    End Select
                End If
            ElseIf PassValue <= -1030 And PassValue > -1040 Then
                Temp = "INC/FAIL"
                If FlagGroup = False Then
                    ' Give detailed result
                    Select Case PassValue
                        Case -1031
                            Temp = Temp & " HIGH"
                        Case -1032
                            Temp = Temp & " LOW"
                    End Select
                End If
            Else
                Temp = "Unhandled"
            End If

            PassOrFail = Temp
        End Function


        Public Function ViewSummary(PreDef As String,
                                    ReportLevel As Integer,
                                    ShowFails As Boolean,
                                    ShowCriticals As Boolean,
                                    ShowMeasured As Boolean,
                                    ByRef rsResultSet As DataTable) _
            As ReturnCodes

            ' Function scans through the Summary array and sets the reported flag based on flags and expands
            Dim x As Integer
            Dim NumItems As Integer
            Dim bmLastRecord As Integer
            Dim bookMark As Integer
            NumItems = (If(rsResultSet?.Rows?.Count, 0))


            ' Hide all Records and convert passflags to text
            For x = 0 To NumItems - 1
                ' Clear item report flag
                rsResultSet(x)("Reported") = False
                rsResultSet(x)("ReportFlag") = PassOrFail(Convert.ToInt64(rsResultSet(x)("result_passflag")), False)

                ' QUICK FIX - Allowed string data to be shown in passflag column
                If _
                    Convert.ToString(rsResultSet(x)("result_stringdata")) <> "" And
                    Convert.ToString(rsResultSet(x)("ReportFlag")) = "" Then
                    ' There is stringdata to be displayed..
                    rsResultSet(x)("ReportFlag") = rsResultSet(x)("result_stringdata")
                End If

            Next x


            For x = 0 To NumItems - 1
                bmLastRecord = x

                '        ' First build report based on predefined reports
                '        If PreDef <> "" Then
                '            ' Pass through array looking for report letter in 'Reports' column
                '            If InStr(1, UCase(rsResultSet.Fields("itemlistdef_predef").Value), UCase(PreDef)) > 0 Then
                '                ' Show record
                '                rsResultSet.Fields("Reported").Value = True
                '            End If
                '        End If

                ' Second, expand based on report level
                If Convert.ToInt32(rsResultSet(x)("itemlistdef_report_level")) <= ReportLevel Then
                    ' Show record
                    rsResultSet(x)("Reported") = True
                End If

                ' Third, expand based on failures
                bookMark = bmLastRecord
                If ShowFails = True Then
                    If _
                        Convert.ToInt32(rsResultSet(x)("result_passflag")) <= 0 And
                        Convert.ToString(rsResultSet(x)("result_passflag")) <> "" Then
                        ' Show record
                        rsResultSet(x)("Reported") = True
                        ExpandParents(x, rsResultSet)
                    End If
                End If

                ' Fourth, expand based on critical specs
                bookMark = bmLastRecord
                If ShowCriticals = True Then
                    If Convert.ToInt32(rsResultSet(x)("itemlistdef_critical_spec")) > 0 Then
                        ' Show record
                        rsResultSet(x)("Reported") = True
                        ExpandParents(x, rsResultSet)
                    End If
                End If

                ' Fifth, expand based on values present
                bookMark = bmLastRecord
                If ShowMeasured = True Then
                    If _
                        Not DBNull.Value.Equals(rsResultSet(x)("result_value")) And
                        Convert.ToString(rsResultSet(x)("result_value")) <> "" Then
                        ' Show record
                        rsResultSet(x)("Reported") = True
                        ExpandParents(x, rsResultSet)
                    End If
                End If

                bookMark = bmLastRecord


            Next x

            Return ReturnCodes.UDBS_OP_SUCCESS
        End Function
    End Module
End Namespace
