Imports System.Net.Sockets
Imports System.Text
Public Class janomeMachine
    Public ipAddr As String = "192.168.200.180"
    Public port As Integer = 10030
    Public tcpClient As TcpClient
    Friend logger As NLog.Logger = NLog.LogManager.GetLogger("Program")
    Public dicCmd As New Dictionary(Of String, String)
    Public sysInStatus() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
    Public sysOutStatus() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
    Public genInStatus() As Integer = {0, 0, 0, 0, 0, 0, 0, 0}
    Public genOutStatus() As Integer = {0, 0, 0, 0, 0, 0, 0, 0}
    Public strGenIn, strGenOut, strSysIn, strSysOut As String
    Public status As String


    Sub New()

        logger.Trace("Janome Initial")
        dicCmd.Add("B1", "00 00 00 08 03 03 42 31") '[Acquire Robot]
        dicCmd.Add("R0", "00 00 00 08 03 03 52 30") 'Initial
        dicCmd.Add("R1", "00 00 00 0C 03 03 52 31 ") ' Job change 
        dicCmd.Add("B2002", "00 00 00 0c 03 03 42 32 30 30 30 32") ' Enter to E_run mode
        dicCmd.Add("B2000", "00 00 00 0c 03 03 42 32 30 30 30 30") ' Enter to Teaching mode
        dicCmd.Add("K2", "00 00 00 14 03 03 4b 32 30 30 30 33 30 30 30 30 30 30 30 ") ' Set SysOut
        dicCmd.Add("K3", "00 00 00 14 03 03 4b 33 30 30 30 33 30 30 30 30 30 30 30 ") ' Reset SysOut
        '                                          K   2  0  0  0  3  0  0  0  0  0  0  0  1
        '                 ----Header-------
        dicCmd.Add("K0_0000", "00 00 00 0c 03 03 4b 30 30 30 30 30") ' read SysIn
        dicCmd.Add("K0_0001", "00 00 00 0c 03 03 4b 30 30 30 30 31") ' read GenIn
        dicCmd.Add("K0_0003", "00 00 00 0c 03 03 4b 30 30 30 30 33") ' read SysOut
        dicCmd.Add("K0_0004", "00 00 00 0c 03 03 4b 30 30 30 30 34") ' read GenOut

        strGenIn = ""
        strGenOut = ""
        strSysIn = ""
        strSysOut = ""
        tcpClient = New TcpClient

    End Sub

    Function connect() As Boolean
        Try
            If Not tcpClient.Connected Then

                tcpClient.Connect(ipAddr, port)
                logger.Info("Janome is connected via" & ipAddr)
            Else
                logger.Info("Janome was connected")
            End If

        Catch ex As Exception
            logger.Error("Cannot connect to Janome " & ipAddr)
            Return False
        End Try

        Return True
    End Function
    Function close() As Boolean
        Try
            tcpClient.Close()
            tcpClient.Dispose()
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Function readyToStart() As Boolean

        readSysIO()
        readGenIo()

        ' genInStatus 2, 3  = XY Spring sensors
        ' sysOutStatus 0  = ready signal
        ' sysOutStatus 12  = Clamped
        If genInStatus(2) = 1 And genInStatus(3) And sysOutStatus(12) = 1 And sysOutStatus(0) = 1 Then
            'ready signal from janome
            Return True
        End If
        Return False
    End Function

    'Function waitReady(waitingTime As Integer) As Boolean
    '    waitReady = False
    '    Dim stopwatch As New Stopwatch()
    '    stopwatch.Start()

    '    Do While stopwatch.ElapsedMilliseconds < waitingTime
    '        Threading.Thread.Sleep(1000)

    '        If getStatus() = "Waiting run" Then
    '            waitReady = True
    '            Exit Do
    '        End If

    '    Loop

    'End Function
    Function startJob() As Boolean

        Try
            setSysOut("8")
            Threading.Thread.Sleep(100)
            resetSysOut("8")
        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function

    Sub readSysIO()
        Dim strResp As String = sendHex(dicCmd("K0_0000"))
        Dim result As String = cutString(strResp, "k00000")
        ' logger.Debug(result)
        Dim binaryString As String
        Dim i2 As Integer = 0

        Dim intIOAddrLoop As Integer = 2 ' 2*8 = 16

        For iIO As Integer = 0 To intIOAddrLoop - 1

            Dim strTemp As String = result.Substring(iIO * 2, 2)
            binaryString = HexStringToBinaryString(strTemp)
            'logger.Debug(strTemp)
            'logger.Debug(binaryString)
            i2 = iIO * 8
            For i As Integer = binaryString.Length - 1 To 0 Step -1

                sysInStatus(i2) = CInt(binaryString(i).ToString)
                i2 = i2 + 1
            Next

        Next
        ' SysOut
        strResp = sendHex(dicCmd("K0_0003"))
        result = cutString(strResp, "k00003")
        'logger.Debug(result)
        i2 = 0

        For iIO As Integer = 0 To intIOAddrLoop - 1

            Dim strTemp As String = result.Substring(iIO * 2, 2)
            binaryString = HexStringToBinaryString(strTemp)
            'logger.Debug(strTemp)
            'logger.Debug(binaryString)
            i2 = iIO * 8
            For i As Integer = binaryString.Length - 1 To 0 Step -1

                sysOutStatus(i2) = CInt(binaryString(i).ToString)
                i2 = i2 + 1
            Next

        Next
        Dim strSysIo As String = ""

        For i = 0 To sysInStatus.Length - 1
            strSysIo = strSysIo + CStr(sysInStatus(i))
            'logger.Debug("SysIn " + CStr(i) + ": " + CStr(sysInStatus(i)))
        Next
        logger.Info("SysIn status :" + strSysIo)
        strSysIn = strSysIo
        strSysIo = ""
        For i = 0 To sysInStatus.Length - 1
            strSysIo = strSysIo + CStr(sysOutStatus(i))
            'logger.Debug("SysOut " + CStr(i) + ": " + CStr(sysOutStatus(i)))
        Next
        logger.Info("SysOut status :" + strSysIo)
        strSysOut = strSysIo

    End Sub
    Sub readGenIo()
        Dim strResp As String = sendHex(dicCmd("K0_0001"))
        Dim result As String = cutString(strResp, "k00001")
        logger.Debug(result)
        Dim binaryString As String
        Dim i2 As Integer = 0

        Dim intIOAddrLoop As Integer = 1 ' 1*8

        For iIO As Integer = 0 To intIOAddrLoop - 1

            Dim strTemp As String = result.Substring(iIO * 2, 2)
            binaryString = HexStringToBinaryString(strTemp)
            'logger.Debug(strTemp)
            'logger.Debug(binaryString)
            i2 = iIO * 8
            For i As Integer = binaryString.Length - 1 To 0 Step -1

                genInStatus(i2) = CInt(binaryString(i).ToString)
                i2 = i2 + 1
            Next

        Next
        ' genOut
        strResp = sendHex(dicCmd("K0_0004"))
        result = cutString(strResp, "k00004")
        logger.Debug(result)
        i2 = 0

        For iIO As Integer = 0 To intIOAddrLoop - 1

            Dim strTemp As String = result.Substring(iIO * 2, 2)
            binaryString = HexStringToBinaryString(strTemp)
            'logger.Debug(strTemp)
            'logger.Debug(binaryString)
            i2 = iIO * 8
            For i As Integer = binaryString.Length - 1 To 0 Step -1

                genOutStatus(i2) = CInt(binaryString(i).ToString)
                i2 = i2 + 1
            Next

        Next
        Dim strGenIo As String = ""

        For i = 0 To genInStatus.Length - 1
            strGenIo = strGenIo + CStr(genInStatus(i))
            'logger.Debug("SysIn " + CStr(i) + ": " + CStr(sysInStatus(i)))
        Next
        logger.Info("GenIn status :" + strGenIo)
        strGenIn = strGenIo
        strGenIo = ""
        For i = 0 To genInStatus.Length - 1
            strGenIo = strGenIo + CStr(genOutStatus(i))
            'logger.Debug("SysOut " + CStr(i) + ": " + CStr(sysOutStatus(i)))
        Next
        logger.Info("GenOut status :" + strGenIo)
        strGenOut = strGenIo
    End Sub
    Function HexStringToBinaryString(hexString As String) As String
        Dim binaryStringBuilder As New StringBuilder()

        For Each hexChar As Char In hexString
            ' Convert each hex character to its binary representation
            Dim binaryValue As String = Convert.ToString(Convert.ToInt32(hexChar.ToString(), 16), 2).PadLeft(4, "0"c)
            binaryStringBuilder.Append(binaryValue)
        Next

        Return binaryStringBuilder.ToString()
    End Function
    Function cutString(strIn As String, cutStr As String) As String
        Dim findIndex As Integer = strIn.IndexOf(cutStr)

        If findIndex > -1 Then
            cutString = strIn.Substring(findIndex + cutStr.Length, strIn.Length - (findIndex + cutStr.Length))
        Else
            Return "" ' cannot find string
        End If
    End Function
    Function setSysOut(strIn As String) As Boolean
        Dim strCmd As String = dicCmd("K2") + ASCIIToHex(strIn)
        logger.Info("Set sysOut : " + strIn)

        Dim strResp As String = sendHex(strCmd)
        'Need to cacth error return
        Return True
    End Function
    Function resetSysOut(strIn As String) As Boolean
        Dim strCmd As String = dicCmd("K3") + ASCIIToHex(strIn)
        logger.Trace("Set sysOut : " + strIn)

        Dim strResp As String = sendHex(strCmd)
        'Need to cacth error return
        Return True
    End Function
    Function getStatus() As String
        logger.Trace("GetStatus")
        Dim strResp As String = sendHex(dicCmd("B1"))
        'b10200000000780000000800029183000005710000C0A8C8B400000000FFFFFF00272F00000000
        'b10200000000780000002900029183000005710000C0A8C8B400000000FFFFFF00272F00000000
        If strResp.Substring(18, 4) = "0002" Then
            getStatus = "Waiting mechanical initialize"
        ElseIf strResp.Substring(18, 4) = "0008" Then
            getStatus = "Waiting run"
        ElseIf strResp.Substring(18, 4) = "000B" Then
            getStatus = "Intialing"
        ElseIf strResp.Substring(18, 4) = "0029" Then
            getStatus = "Running (moving)"
        ElseIf strResp.Substring(18, 4) = "0068" Then
            getStatus = "Running (waiting condition)"
        ElseIf strResp.Substring(18, 4) = "1100" Then
            getStatus = "Emergency stop"
        Else
            getStatus = "UndefineYet (" + strResp.Substring(18, 4) + ")"
        End If
        status = getStatus
        logger.Info("GetStatus = " + getStatus)
        'Return strResp

    End Function
    Function getMode() As String
        logger.Trace("GetMode")
        Dim strResp As String = sendHex(dicCmd("B1"))
        'b1020000000 0780000000 800029183000005710000C0A8C8B400000000FFFFFF00272F00000000
        'b1020000000 0780000002 900029183000005710000C0A8C8B400000000FFFFFF00272F00000000
        Dim mode As String = strResp.Substring(2, 2)
        If mode = "00" Then
            getMode = "Teaching Mode"
        ElseIf mode = "01" Then
            getMode = "S-Run Mode"
        ElseIf mode = "02" Then
            getMode = "E-Run Mode"
        ElseIf mode = "03" Then
            getMode = "Undefined Mode"
        ElseIf mode = "04" Then
            getMode = "Test-Run Mode"
        Else
            getMode = "UndefineYet (" + strResp.Substring(2, 2) + ")"
        End If
        logger.Info("GetMode = " + getMode)
    End Function
    Function getJobNum() As String
        logger.Trace("GetJob")
        Dim strResp As String = sendHex(dicCmd("B1"))
        'Dim temp As String = "b10200000000790000000800024392000005110000C0A8C8B400000000FFFFFF00272F00000000"
        Dim decimalValue As Integer = Convert.ToInt32(strResp.Substring(12, 2), 16)

        Return CStr(decimalValue)
    End Function
    Function setJobNum(strinput As String) As String
        Dim strResp As String = ""

        Dim decimalValue As Integer = CInt(strinput)    ' 121
        Dim strTemp As String = decimalValue.ToString("X4") '79

        Dim hexString As String = ""

        For Each c As Char In strTemp
            Dim asciiValue As Integer = AscW(c)
            hexString &= asciiValue.ToString("X2") & " " ' Append the hexadecimal representation to the result string
        Next
        hexString = hexString.Trim() ' 30 30 37 39

        Dim strCmd As String = dicCmd("R1") & hexString
        strResp = sendHex(strCmd)

        Return strResp
    End Function
    Function sendHex(strInput As String) As String
        Dim responseHexString As String = ""

        Dim hexBytes As Byte() = strInput.Split(" ").Select(Function(hex) Convert.ToByte(hex, 16)).ToArray()
        Dim stream As NetworkStream = tcpClient.GetStream()

        logger.Trace("SendHex send: " & strInput)
        logger.Trace("SendHex send: " & HexToASCII(strInput))

        stream.Write(hexBytes, 0, hexBytes.Length)

        Threading.Thread.Sleep(300)
        Dim responseBuffer(2048) As Byte
        Dim responseLength As Integer = stream.Read(responseBuffer, 0, responseBuffer.Length)
        responseHexString = BitConverter.ToString(responseBuffer, 0, responseLength).Replace("-", " ")
        'stream.Close()

        logger.Trace("SendHex recv: " & responseHexString)
        logger.Trace("SendHex recv: " & HexToASCII(responseHexString))

        Return HexToASCII(responseHexString)
    End Function
    Function HexToASCII(hex As String) As String

        Dim hexBytes As String() = hex.Split(New Char() {" "c, ","c}, StringSplitOptions.RemoveEmptyEntries)
        Dim ascii As New Text.StringBuilder(hexBytes.Length * 2) ' Adjusted length to account for added spaces
        'Dim ascii As New Text.StringBuilder(hexBytes.Length)

        For Each hexPair As String In hexBytes

            Dim asciiValue As Integer = Convert.ToInt32(hexPair, 16)
            If asciiValue = AscW(" "c) Then
                ascii.Append(" ") ' Insert a space in the ASCII representation
            Else
                ascii.Append(ChrW(asciiValue))
            End If
        Next
        HexToASCII = ascii.ToString()
        HexToASCII = HexToASCII.Substring(6)  ' Cut header for Char doesn't readable

    End Function
    Function ASCIIToHex(oneStr As String) As String
        If oneStr.Length > 1 Then
            logger.Error("This function use for convert only one charecter")
            Return False
        End If
        Return BitConverter.ToString(Encoding.ASCII.GetBytes(oneStr))
    End Function
    Function init() As Boolean

        sendHex(dicCmd("R0"))

        Return True
    End Function
    Function testRun() As Boolean

        Dim strCmd As String = dicCmd("K2 SysOut") + "" ' address
        Return False
    End Function
    Function modeERun() As Boolean
        sendHex(dicCmd("B2002"))
        Return True
    End Function
    Function modeTeaching() As Boolean
        sendHex(dicCmd("B2000"))
        Return True
    End Function
End Class

