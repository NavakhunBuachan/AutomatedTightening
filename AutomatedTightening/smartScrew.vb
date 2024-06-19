

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

    Function getPresetNumber() As String
        Dim response As String = sendAndRecvAscii("P009")
        'Dim response As String = "p00011"
        Dim index As Integer = response.IndexOf("p")
        If index = -1 Then
            logger.Error("Cannot found correct response")
            getPresetNumber = "0"
        Else
            getPresetNumber = response.Substring(index + 4, 1)
        End If

    End Function

    Function setPresetNumber(strIn As String) As Boolean

        If strIn.Length > 1 Then
            logger.Error("Preset should be 1 digit")
            Return False
        Else
            sendAndRecvAscii("S009000" + strIn)
            Return True
        End If

    End Function

    Function getTotalCount() As String
        Dim response As String = sendAndRecvAscii("P130")
        Dim index As Integer = response.IndexOf("p")
        If index = -1 Then
            logger.Error("Cannot found correct response")
            getTotalCount = "0"
        Else
            getTotalCount = response.Substring(index + 4, 1)
        End If

    End Function

    Function setTotalCount(strIn As String) As Boolean
        If strIn.Length > 1 Then
            logger.Error("This sw support 1 digit")
            Return False
        Else
            sendAndRecvAscii("S130000" + strIn)
            Return True
        End If
    End Function
    Function getMinAngle() As String
        Dim response As String = sendAndRecvAscii("P021")
        'Dim response As String = "p00011"
        Dim index As Integer = response.IndexOf("p")
        If index = -1 Then
            logger.Error("Cannot found correct response")
            getMinAngle = "0"
        Else
            getMinAngle = response.Substring(index + 2, 3)
        End If

    End Function

    Function setMinAngle(strIn As String) As Boolean

        If strIn.Length <> 3 Then
            logger.Error("MinAgles length should be 3")
            Return False
        Else
            sendAndRecvAscii("S0210" + strIn)
            Return True
        End If

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

    Function sendAndRecvAscii(asciiInput As String) As String
        Dim strSend As String
        'Dim inputData As String = strIn
        'Dim asciiInput As String = "M1"
        Dim checksum As String = CalculateChecksum(asciiInput)
        strSend = asciiInput + checksum
        strSend = AsciiToHex(strSend)
        strSend = addHeaderAndTail(strSend)
        'logger.Info("Send : " + strSend)

        Dim sendData As Byte() = HexStringWithSpacesToByteArray(strSend)
        serialPortSmartScrew.Write(sendData, 0, sendData.Length)
        Threading.Thread.Sleep(200)
        Dim receivedData As String = serialPortSmartScrew.ReadExisting()

        sendAndRecvAscii = receivedData
    End Function

    Function addHeaderAndTail(strIn As String) As String

        Dim strOut As String = "02 " + strIn + " 03"

        Return strOut

    End Function
    Function AsciiToHex(ByVal asciiInput As String) As String
        Dim hexOutput As String = ""
        For Each ch As Char In asciiInput
            ' Convert each character to a hexadecimal string
            hexOutput &= Asc(ch).ToString("X2") & " " ' X2 format to ensure two digits
        Next
        Return hexOutput.Trim()
    End Function

    Function CalculateChecksum(ByVal asciiInput As String) As String
        Dim sum As Integer = 0

        ' Calculate sum of ASCII values
        For Each ch As Char In asciiInput
            sum += Asc(ch)
        Next

        ' Convert sum to hexadecimal
        Dim sumHex As String = sum.ToString("X")

        ' Extract the last digit of the hexadecimal sum
        Dim lastHexDigit As String = sumHex(sumHex.Length - 1)

        ' Return the last digit as a string in uppercase
        Return lastHexDigit.ToUpper()
    End Function

    'Function StringToHex(ByVal text As String) As String
    '    Dim hex As String = ""
    '    For Each c As Char In text
    '        hex &= Asc(c).ToString("X2") & " "
    '    Next
    '    Return hex.Trim()
    'End Function

    'Function StringToHexRS232(input As String) As String
    '    Dim sb As New System.Text.StringBuilder(input.Length * 2)
    '    For Each c As Char In input
    '        sb.AppendFormat("{0:X2} ", AscW(c))
    '    Next
    '    Return sb.ToString().Trim()
    'End Function
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
