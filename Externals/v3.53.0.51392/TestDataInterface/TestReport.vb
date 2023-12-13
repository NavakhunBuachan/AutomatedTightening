Imports UdbsInterface.MasterInterface

Namespace TestDataInterface
    ''' <summary>
    ''' Support class for producing test reports as Excel and DAT files.
    ''' </summary>
    ''' <remarks>
    ''' These methods used to be included in the CTestData_Instance class.
    ''' </remarks>
    Friend Class TestReport
        Private _test As CTestdata_Instance

        ''' <summary>
        ''' Default constructor.
        ''' </summary>
        ''' <param name="test">The test instance for which we want to create reports.</param>
        Public Sub New(test As CTestdata_Instance)
            _test = test
        End Sub

        ''' <summary>
        ''' Create a DAT test report file.
        ''' </summary>
        ''' <param name="TemplatePath">The path to the template file.</param>
        ''' <param name="OutputPath">The path of the file to create.</param>
        ''' <returns>The outcome of the operation.</returns>
        Public Function CreateDATFile(ByVal TemplatePath As String, ByVal OutputPath As String) As ReturnCodes
            Dim rsProcess As New DataTable
            Dim rsResult As New DataTable

            Dim strTemplateData As String() = New String() {}
            Dim strFileData As String()

            Dim i As Integer
            Dim Flag As String = ""
            Try
                ' Get Process Data
                rsProcess = _test.Instance_RS

                ' Get Result Data
                rsResult = _test.Results_RS

                ' Input Template File
                If InputFile(TemplatePath, strTemplateData) < 0 Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Set output array to same dimension as input array
                ReDim strFileData(0 To UBound(strTemplateData))

                ' Parse a line of the .dat file template
                For i = 1 To UBound(strTemplateData)
                    ParseLine(_test.ProductNumber & ":-:" & _test.UnitOraclePN & ":-:" & _test.UnitCataloguePN & ":-:" & _test.UnitVariance, _test.UnitSerialNumber,
                              CStr(_test.Sequence), _test.Stage, strTemplateData(i), strFileData(i), Flag, rsProcess, rsResult)
                Next i

                ' Create data file
                If OutputFile(OutputPath, strFileData) < 0 Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS
                'End If
            Catch e As Exception
                MsgBox(e.Message & vbNewLine & vbNewLine & e.StackTrace)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        Private Function InputFile(ByVal TemplatePath As String,
                          ByRef strTemplateData() As String) _
                          As Long
            'Purpose: Read template file into memory
            'Returns: 1 if successful, -1 if unsuccessful

            Dim fnum As Integer
            Dim i As Integer
            Dim buf As String

            Try

                fnum = FreeFile()
                'Load the template for the .dat file
                'Open TemplatePath For Input As #fnum
                FileOpen(fnum, TemplatePath, OpenMode.Input)
                Do Until EOF(fnum)
                    'Line Input #fnum, buf
                    buf = LineInput(fnum)
                    i = i + 1
                    ReDim Preserve strTemplateData(0 To i)
                    strTemplateData(i) = buf
                Loop
                'Close #fnum
                FileClose(fnum)

                Return 1

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return -1
            End Try
        End Function

        Private Function ParseLine(ByVal ProductNumber As String,
                          ByVal SerialNumber As String,
                          ByVal Sequence As String,
                          ByVal Stage As String,
                          ByVal InputLine As String,
                          ByRef OutputLine As String,
                          ByRef Flag As String,
                          ByRef rsProcess As DataTable,
                          ByRef rsResult As DataTable) _
                                       As Long
            ' [Token_Name,Column;Format]

            ' Purpose: Substitutes tokens for values from database
            ' Returns: 1 if successful, -1 if unsuccessful
            ' April 6, 2006: the new ProductNumber contains ProductIdentifier (old productnumber, oraclePN, catPN & variance

            Dim TokenStart As Integer
            Dim TokenEnd As Integer

            Dim StrToken As String = ""
            Dim StrColumn As String = ""
            Dim StrItemName As String = ""
            Dim StrFormat As String = ""
            Dim ItemValue As String = ""
            Dim orgStrToken As String, orgTokenStart As Integer, orgTokenEnd As Integer
            Dim strArr As String()

            ParseLine = 1

            strArr = Split(ProductNumber, ":-:")

            TokenStart = 1
            TokenStart = InStr(TokenStart, InputLine, "[")
            Do Until TokenStart = 0

                ' Get end of token position
                TokenEnd = InStr(TokenStart, InputLine, "]")
                orgTokenStart = TokenStart
                orgTokenEnd = TokenEnd
                ' Get token
                StrToken = Mid(InputLine, TokenStart + 1, (TokenEnd - TokenStart) - 1)
                orgStrToken = StrToken
                ' Check for specified column
                If InStr(1, StrToken, ",") <> 0 Then
                    TokenStart = InStr(1, StrToken, ",")
                    StrColumn = Trim(Mid(StrToken, TokenStart + 1))
                    StrToken = Mid(StrToken, 1, TokenStart - 1)
                    ' Check for specified Format
                    If InStr(1, StrColumn, ";") <> 0 Then
                        TokenStart = InStr(1, StrColumn, ";")
                        StrFormat = Trim(Mid(StrColumn, TokenStart + 1))
                        StrColumn = Mid(StrColumn, 1, TokenStart - 1)
                    End If
                Else
                    ' Check for specified Format
                    If InStr(1, StrToken, ";") <> 0 Then
                        TokenStart = InStr(1, StrToken, ";")
                        StrFormat = Trim(Mid(StrToken, TokenStart + 1))
                        StrToken = Mid(StrToken, 1, TokenStart - 1)
                    End If
                End If

                ' Get the ItemName in Token
                If InStr(1, StrFormat, ";") <> 0 Then
                    StrToken = "this is an array"
                End If
                StrItemName = Trim(StrToken)

                ' Check for Pre-Defined Itemnames
                Select Case UCase(StrItemName)
                    Case "SERIALNUMBER"
                        ItemValue = SerialNumber
                    Case "SEQUENCE"
                        ItemValue = Sequence
                    Case "PRODUCTNUMBER"
                        ItemValue = strArr(0)
                    Case "ORACLEPN"
                        If UBound(strArr) >= 1 Then
                            ItemValue = strArr(1)
                        Else
                            ItemValue = strArr(0)
                        End If
                    Case "CATALOGUEPN"
                        If UBound(strArr) >= 2 Then
                            ItemValue = strArr(2)
                        Else
                            ItemValue = ""
                        End If
                    Case "VARIANCE"
                        If UBound(strArr) >= 3 Then
                            ItemValue = strArr(3)
                        Else
                            ItemValue = ""
                        End If
                    Case Else
                        ' Search tables
                        ParseLine = 1
                        If FindResultData(StrItemName, StrColumn, ItemValue, Flag, rsResult) <= 0 AndAlso
                           FindProcessData(StrItemName, ItemValue, rsProcess(0)) <= 0 Then
                            'If UCase(ProductNumber) = "OA1510-LTWSM1" And UCase(StrItemName) Like "COND*G_VCB_V" Then
                            '    ItemValue = CStr(Val(ItemValue) * 1000)
                            'End If
                            'ElseIf FindProcessData(StrItemName, ItemValue, rsProcess) > 0 Then
                            '
                            'Else
                            ItemValue = "[" & orgStrToken & "]"
                            ParseLine = -1
                        End If
                End Select

                Dim FormattedItem As String = Microsoft.VisualBasic.Format(ItemValue, StrFormat)
                If Date.TryParse(ItemValue, Nothing) Then
                    Dim asDate As Date = Date.Parse(ItemValue)
                    FormattedItem = asDate.ToString(StrFormat)
                ElseIf Double.TryParse(ItemValue, Nothing) Then
                    FormattedItem = Microsoft.VisualBasic.Format(CDbl(ItemValue), StrFormat)
                ElseIf Integer.TryParse(ItemValue, Nothing) Then
                    FormattedItem = Microsoft.VisualBasic.Format(CInt(ItemValue), StrFormat)
                End If


                InputLine = Left(InputLine, orgTokenStart - 1) &
                    FormattedItem &
                    Right(InputLine, Len(InputLine) - orgTokenEnd)
                If InputLine <> "" Then
                    TokenStart = InStr(orgTokenStart + 1, InputLine, "[")
                    StrColumn = ""
                    StrFormat = ""
                    StrToken = ""
                Else
                    TokenStart = 0
                End If
            Loop

            OutputLine = InputLine.Replace("~", "") 'for some reason there's a tilde character that doesn't get snipped here
        End Function

        Private Function OutputFile(ByVal OutputPath As String,
                                    ByRef strFileData() As String) _
                                    As Long
            'Purpose: Generates lines of text for output file, from template lines of text
            'Returns: 1 if successful, -1 if unsuccessful

            Dim i As Integer
            Dim fnum As Integer

            Try

                If Dir(OutputPath) <> "" Then
                    If MsgBox("File already exsits!" & vbCrLf & "Do you wish to overwrite?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = vbNo Then
                        Return -1
                    End If
                End If

                fnum = FreeFile()
                'load the template for the .dat file
                'Open OutputPath For Output As #fnum
                FileOpen(fnum, OutputPath, OpenMode.Output)
                For i = 1 To UBound(strFileData)
                    PrintLine(fnum, strFileData(i))
                Next i
                FileClose(fnum)
                OutputFile = 1
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Create a Microsoft Excel test report file.
        ''' </summary>
        ''' <param name="TemplatePath">The path to the template file.</param>
        ''' <param name="OutputPath">The path of the file to create.</param>
        ''' <returns>The outcome of the operation.</returns>
        Public Function CreateExcelFile(ByVal TemplatePath As String,
                                ByVal OutputPath As String) _
                                            As ReturnCodes
            Dim Flag As String = ""
            Dim semi1st As Integer      ' 1st semi-colon location
            Dim moreThan1Semi As Boolean
            Dim ItemName As String = "", arrayName As String = "", valueFormat As String = "", fillStyle As String = ""
            Dim DataGroupName As String = "", NumElements As Integer, DataType As Long, IsHeader As Boolean
            Dim arrVar As Array = New Object() {}
            Dim i As Long
            Dim ProductRelease As Double
            Dim OutputLine As String = ""
            Dim InputLine As String = ""

            Dim rsResult As New DataTable
            Dim rsProcess As New DataTable

            Dim t As List(Of String)
            t = New List(Of String)
            Dim workbook As ClosedXML.Excel.XLWorkbook = Nothing

            ' Returns: 1 if successful, -1 if unsuccessful
            Try
                'Screen.MousePointer = vbHourglass
                ProductRelease = 0

                ' Get Process Data
                rsProcess = _test.Instance_RS
                workbook = New ClosedXML.Excel.XLWorkbook(TemplatePath)
                ' Get Result Data
                rsResult = _test.Results_RS
                For Each sheet In workbook.Worksheets
                    Dim cells = sheet.Cells().Where(Function(c)
                                                        Dim v As String
                                                        Try
                                                            v = c.GetValue(Of String)
                                                        Catch ex As Exception
                                                            Return False
                                                        End Try
                                                        Return v.Contains("~[")
                                                    End Function)
                    For Each cell In cells
                        InputLine = CStr(cell.Value)
                        If InputLine.Split({"~["}, StringSplitOptions.None).Count > 2 Then
#Disable Warning BC42322 ' Runtime errors might occur when converting to or from interface type
                            MsgBox("Invalid token definition (more than one ""~["")." & Chr(13) & "Address " & CStr(cell.Address), vbExclamation)
#Enable Warning BC42322 ' Runtime errors might occur when converting to or from interface type
                            Do While InStr(1, InputLine, "~[") > 0
                                i = InStr(1, InputLine, "~[")
                                InputLine = Mid(InputLine, 1, CInt(i - 1)) & Mid(InputLine, CInt(i + 1))
                            Loop
                            cell.Value = InputLine
                        Else
                            semi1st = InStr(1, InputLine, ";")
                            moreThan1Semi = False
                            If semi1st <> 0 Then
                                If InStr(semi1st + 1, InputLine, ";") <> 0 Then
                                    moreThan1Semi = True
                                End If
                            End If
                            If moreThan1Semi Then
                                semi1st = InStr(1, InputLine, "]")
                                If semi1st <> 0 Then
                                    If InStr(semi1st + 1, InputLine, "]") <> 0 Then
                                        ' most likely this is just a multiple items
                                        moreThan1Semi = False
                                    End If
                                End If
                            End If
                            If moreThan1Semi Then
                                ParseArrayName(ItemName, arrayName, valueFormat, fillStyle, InputLine)
                                NumElements = 0     ' initialize to avoid recalling previous arrVar
                                If Not _test.Results.Keys.Contains(ItemName) Then
                                    t.Add(ItemName)
                                Else
                                    If _test.Results(ItemName).GetArray(arrayName, DataGroupName, NumElements, CType(DataType, VariantType), IsHeader, arrVar) <> ReturnCodes.UDBS_OP_SUCCESS Then
                                        cell.Value = InputLine
                                    Else
                                        cell.Value = InputLine
                                        Dim tempCell As ClosedXML.Excel.IXLCell = cell
                                        For i = 0 To NumElements - 1
                                            If LCase(fillStyle) = "row" Then
                                                tempCell.Value = Format(arrVar.GetValue(i), valueFormat)
                                                tempCell = tempCell.CellRight()
                                            Else
                                                tempCell.Value = Format(arrVar.GetValue(i), valueFormat)
                                                tempCell = tempCell.CellBelow()
                                            End If
                                        Next i
                                    End If
                                End If
                            Else
                                ParseLine(_test.ProductNumber & ":-:" & _test.UnitOraclePN & ":-:" & _test.UnitCataloguePN & ":-:" & _test.UnitVariance,
                                        _test.UnitSerialNumber, CStr(_test.Sequence), _test.Stage, InputLine, OutputLine, Flag, rsProcess, rsResult)
                                cell.Value = OutputLine
                            End If

                        End If

                    Next
                Next
                workbook.ForceFullCalculation = True
                workbook.SaveAs(OutputPath)
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch e As Exception
                MsgBox(e.Message & vbNewLine & vbNewLine & e.StackTrace)
                Return ReturnCodes.UDBS_OP_FAIL
            Finally
                If workbook IsNot Nothing Then
                    workbook.Dispose()
                End If
            End Try
        End Function

        Private Function ParseArrayName(ByRef ItemName As String,
                               ByRef arrayName As String,
                               ByRef valueFormat As String,
                               ByRef fillStyle As String,
                               ByVal InputLine As String) _
                                       As Long
            ' [Token_Name;Format;arrayName;row/col]


            Dim TokenStart As Integer
            Dim TokenEnd As Integer

            Dim StrToken As String
            Dim StrColumn As String

            TokenStart = 1
            TokenStart = InStr(TokenStart, InputLine, "[")
            Do Until TokenStart = 0

                ' Get end of token position
                TokenEnd = InStr(TokenStart, InputLine, "]")

                ' Get token
                StrToken = Mid(InputLine, TokenStart + 1, (TokenEnd - TokenStart) - 1)

                ' Get the ItemName in Token
                If InStr(1, StrToken, ";") <> 0 Then
                    TokenStart = InStr(1, StrToken, ";")
                    ItemName = Trim(Mid(StrToken, 1, TokenStart - 1))
                    StrToken = Mid(StrToken, TokenStart + 1)
                End If

                ' Check for specified column
                '*** ignore this value
                If InStr(1, StrToken, ",") <> 0 Then
                    TokenStart = InStr(1, StrToken, ",")
                    StrColumn = Trim(Mid(StrToken, 1, TokenStart - 1))
                    StrToken = Mid(StrToken, TokenStart + 1)
                End If

                ' Check for specified Format
                If InStr(1, StrToken, ";") <> 0 Then
                    TokenStart = InStr(1, StrToken, ";")
                    valueFormat = Trim(Mid(StrToken, 1, TokenStart - 1))
                    StrToken = Mid(StrToken, TokenStart + 1)
                End If

                ' Check for specified arrayname
                If InStr(1, StrToken, ";") <> 0 Then
                    TokenStart = InStr(1, StrToken, ";")
                    arrayName = Trim(Mid(StrToken, 1, TokenStart - 1))
                    StrToken = Mid(StrToken, TokenStart + 1)
                End If

                ' Check for specified Fill style (row/col)
                fillStyle = Trim(StrToken)
                ' default fillStyle to row
                If fillStyle <> "col" Then fillStyle = "row"

                TokenStart = InStr(TokenStart + 1, InputLine, "[")
            Loop
            Return 0
        End Function

        ' Candidate for removal.
        Private Function FindProcessData(ByVal Column As String,
                                        ByRef ItemValue As String,
                                        ByRef rsSearch As DataRow) _
                                        As Long

            If FindColumn(Column, rsSearch) = 1 Then
                ItemValue = CStr(rsSearch(Column))
                FindProcessData = 1
            Else
                FindProcessData = -1
            End If
        End Function

        Private Function FindColumn(ByVal Column As String,
                                          ByRef rsSearch As DataRow) _
                                          As Long

            Return FindColumn(Column, rsSearch.Table)
        End Function

        Private Function FindColumn(ByVal Column As String, ByRef rsSearch As DataTable) As Long
            ' Purpose: Look at a row of a recordset and determine if specified column is an element
            Dim x As Integer
            FindColumn = -1
            For x = 0 To rsSearch.Columns.Count - 1
                If rsSearch.Columns(x).ColumnName = Column Then
                    Return 1
                End If
            Next x
        End Function

        ''' <summary>
        ''' Logs the UDBS error in the application log and inserts it in the udbs_error table, with the process instance details:
        ''' Process Type, name, Process ID, UDBS product ID, Unit serial number.
        ''' </summary>
        ''' <param name="ex">Exception raised.</param>
        Private Sub LogErrorInDatabase(ex As Exception)
            If _test Is Nothing Then
                DatabaseSupport.LogErrorInDatabase(ex)
            Else
                With _test
                    DatabaseSupport.LogErrorInDatabase(ex, .Process, .Stage, .ID, .ProductNumber, .UnitSerialNumber)
                End With
            End If
        End Sub
    End Class
End Namespace
