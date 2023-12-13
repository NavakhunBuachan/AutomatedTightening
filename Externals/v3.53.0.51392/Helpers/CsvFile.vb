Imports System.Text
Imports Microsoft.VisualBasic.FileIO

Namespace MasterInterface
    ''' <summary>
    ''' Helper class to read a CSV file.
    ''' </summary>
    Public Module CsvFile
        ''' <summary>
        ''' Helper method for reading a CSV file.
        ''' </summary>
        ''' <param name="path"></param>
        ''' <returns></returns>
        Public Function OpenCsvFile(path As String) As TextFieldParser
            Dim myReader As New TextFieldParser(path) With {
                    .TextFieldType = FileIO.FieldType.Delimited,
                    .CommentTokens = {"#"},
                    .TrimWhiteSpace = False
                    }
            myReader.SetDelimiters(",")

            Return myReader
        End Function
    End Module

End Namespace