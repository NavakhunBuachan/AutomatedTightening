

Imports System.IO.Ports
    Public Class smartScrew
        Public portName As String = "COM3" ' Replace with the appropriate COM port name
        Public baudRate As Integer = 38400
        Public serialPortSmartScrew As SerialPort
        Public dtResult As DataTable
    Friend logger As NLog.Logger = NLog.LogManager.GetLogger("Program")

    Sub New()
        logger.Trace("SmartScrew Initial")
        serialPortSmartScrew = New SerialPort(portName, baudRate, Parity.None, 8, StopBits.One)
    End Sub
    Function connect() As Boolean
        Try
            If Not serialPortSmartScrew.IsOpen Then

                serialPortSmartScrew.Open()
                logger.Info("smartScrew is connected via" & portName)
            Else
                logger.Info("smartScrew was connected")
            End If

        Catch ex As Exception
            logger.Error("Cannot connect to smartScrew " & portName)
            Return False
        End Try

        Return True
    End Function
    Function close() As Boolean
        serialPortSmartScrew.Close()
        serialPortSmartScrew.Dispose()
        Return True
        End Function
        Function sendAndRecv(strIn As String) As String

            Dim sendData As Byte() = HexStringWithSpacesToByteArray(strIn)

            serialPortSmartScrew.Write(sendData, 0, sendData.Length)
            Threading.Thread.Sleep(200)
            Dim receivedData As String = serialPortSmartScrew.ReadExisting()
            'serialPortSmartScrew.
            'Dim receivedHexString As String = StringToHexRS232(receivedData)

            sendAndRecv = receivedData
        End Function

        Function intialTable() As DataTable

            Dim dt As New DataTable

            dt.Columns.Add("Num", GetType(Integer))
            dt.Columns.Add("Time", GetType(TimeSpan))
            dt.Columns.Add("F_Time", GetType(String))
            dt.Columns.Add("Preset", GetType(Integer))
        dt.Columns.Add("T_Tq", GetType(Integer))
        dt.Columns.Add("C_Tq", GetType(Integer))
        dt.Columns.Add("Speed", GetType(Integer))
            dt.Columns.Add("A1", GetType(Integer))
            dt.Columns.Add("A2", GetType(Integer))
            dt.Columns.Add("A3", GetType(Integer))
            dt.Columns.Add("Error", GetType(Integer))
            dt.Columns.Add("Count", GetType(Integer))
        dt.Columns.Add("F_L", GetType(Integer))
        dt.Columns.Add("Status", GetType(Integer))

            Return dt
        End Function
        Function HexStringWithSpacesToByteArray(hexString As String) As Byte()
            Dim hexBytes As String() = hexString.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
            Dim bytes(hexBytes.Length - 1) As Byte
            For i As Integer = 0 To hexBytes.Length - 1
                bytes(i) = Convert.ToByte(hexBytes(i), 16)
            Next
            Return bytes
        End Function
        'Function StringToHex(input As String) As String
        '    Dim hex As New Text.StringBuilder(input.Length * 3) ' Adjusted length to account for potential spaces

        '    For Each c As Char In input
        '        If c = " " Then
        '            hex.Append(" ") ' Insert a space into the hexadecimal representation
        '        Else
        '            hex.AppendFormat("{0:X2}", CInt(AscW(c)))
        '        End If
        '    Next

        '    Return hex.ToString()
        'End Function
        Function StringToHexRS232(input As String) As String
            Dim sb As New System.Text.StringBuilder(input.Length * 2)
            For Each c As Char In input
                sb.AppendFormat("{0:X2} ", AscW(c))
            Next
            Return sb.ToString().Trim()
        End Function
    Function recvSmartScrewData(inputStr As String) As String()
        logger.Trace("==> recvSmartSrewData")
        Dim delimiter As Char = ","
        'Dim trimDot() As String = inputStr.Split(".")
        'recvSmartSrewData = trimDot(1).Split(delimiter)
        Dim strTemp As String()
        recvSmartScrewData = inputStr.Split(delimiter)
        strTemp = recvSmartScrewData

        'strTemp(0) = strTemp(0).Substring(strTemp(0).Length - 5, 5)

        strTemp(0) = cutString(strTemp(0), "m")
        recvSmartScrewData = strTemp

        logger.Trace("<== recvSmartSrewData")
    End Function
    Function cutString(strIn As String, cutStr As String) As String
        Dim findIndex As Integer = strIn.IndexOf(cutStr)

        If findIndex > -1 Then
            cutString = strIn.Substring(findIndex + cutStr.Length, strIn.Length - (findIndex + cutStr.Length))
        Else
            Return "" ' cannot find string
        End If
    End Function
End Class
