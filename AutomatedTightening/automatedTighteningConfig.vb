Public Class automatedTighteningConfig
    Public stationConfCsvPath As String = Application.StartupPath() & "\AutomatedTighteningConfig.csv"
    Friend logger As NLog.Logger = NLog.LogManager.GetLogger("Program")
    Public dtConfig As DataTable

    Public procName, prodNum, jobNum, screwNum As String

    Sub New()
        dtConfig = ReadCsvConfig()
    End Sub
    Function checkProc(procInput As String) As Boolean
        Dim dataRowResult() As DataRow
        Dim strTemp As String

        strTemp = "ProcName ='" & procInput & "'"

        dataRowResult = dtConfig.Select(strTemp)

        If dataRowResult.Length < 1 Then
            logger.Info("Unit is not in WIP")
            Console.WriteLine("Unit is not in WIP")
            dataRowResult = Nothing
            Return False
        End If
        Return True
    End Function
    Function hasConfig(prodNum As String, procName As String) As Boolean

        Dim dataRowResult() As DataRow
        Dim strTemp As String

        strTemp = "ProdNum ='" & prodNum & "' AND ProcName ='" & procName & "'"

        dataRowResult = dtConfig.Select(strTemp)

        If dataRowResult.Length <> 1 Then

            Console.WriteLine("Found DataRow not equal one")
            dataRowResult = Nothing
            Return False
        End If
        Return True

    End Function
    Function loadConfig(prodNumIn As String, procNameIn As String) As Boolean
        prodNum = prodNumIn
        procName = procNameIn

        Dim strSql As String = "ProdNum ='" & prodNum & "' AND ProcName ='" & procName & "'"
        Dim dataRowResult() As DataRow = dtConfig.Select(strSql)

        jobNum = dataRowResult(0)("JobNum")
        screwNum = dataRowResult(0)("ScrewNum")

        Return True

    End Function

    Public Function ReadCsvConfig() As DataTable

        Dim dt As New DataTable

        Using MyReader As New Microsoft.VisualBasic.
                      FileIO.TextFieldParser(stationConfCsvPath)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")

            Dim currentRow As String()
            Dim lineNum As Integer
            While Not MyReader.EndOfData
                'Console.WriteLine("Line Number : " & MyReader.LineNumber)
                Try

                    lineNum = MyReader.LineNumber
                    currentRow = MyReader.ReadFields()

                    Dim strArry(currentRow.Length) As String
                    Dim currentField As String
                    Dim i As Integer = 1

                    For Each currentField In currentRow

                        If lineNum = 1 Then
                            dt.Columns.Add(New DataColumn(currentField, GetType(String)))
                        Else
                            strArry(i) = currentField

                        End If
                        i = i + 1
                        'Console.WriteLine(currentField)
                    Next

                    If lineNum <> 1 Then
                        dt.Rows.Add(strArry(1), strArry(2), strArry(3), strArry(4), strArry(5))
                    End If

                Catch ex As Microsoft.VisualBasic.
                         FileIO.MalformedLineException
                    logger.Error("Cannot read csv config file")
                    MsgBox("Line " & ex.Message &
               "is not valid and will be skipped.")
                End Try
            End While
        End Using
        Return dt
    End Function
    Function printRow(targetCol As String) As String
        Dim targetColumn As String = targetCol
        Dim uniqueValues As New HashSet(Of String)()
        Dim uniqueRows As New List(Of DataRow)()

        ' Iterate through the DataTable and store unique rows based on the specified column
        For Each row As DataRow In dtConfig.Rows
            Dim value As String = row(targetColumn).ToString()
            If Not uniqueValues.Contains(value) Then
                uniqueValues.Add(value)
                uniqueRows.Add(row)
                Console.WriteLine(value)
            End If
        Next

        Return ""

    End Function
End Class
