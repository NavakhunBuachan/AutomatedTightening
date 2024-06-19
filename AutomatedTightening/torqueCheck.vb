Imports NLog

Public Class torqueCheck
    Friend logger As NLog.Logger = NLog.LogManager.GetLogger("Program")
    Friend torqueLogger As NLog.Logger = NLog.LogManager.GetLogger("TorqueCheck")
    'Friend screw As smartScrew
    'Friend janome As janomeMachine
    Dim hiosResponse As String = "None"
    Public lastHiosResponse As String = "None"
    Dim enterCapture As Boolean = False
    Private Sub torqueCheck_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        rtb01.Focus()
        Me.KeyPreview = True

    End Sub

    Private Sub rtb01_KeyDown(sender As Object, e As KeyEventArgs) Handles rtb01.KeyDown
        If e.KeyCode = Keys.Enter Then
            logger.Info("Found Enter on rtb01")

            'MessageBox.Show(getLastResponse(rtb01.Text))
            'logging(getLastResponse(rtb01.Text))
            lastHiosResponse = getLastResponse(rtb01.Text)
            logger.Info("Data : " + lastHiosResponse)
            enterCapture = True
        End If
    End Sub
    Function getLastResponse(strIn As String) As String
        Dim data As String() = strIn.Split(vbLf)
        Return data(data.Length - 1)
    End Function
    Private Function logging() As Boolean

        NLog.LogManager.Configuration.Variables("sn") = "torqueCheck"
        NLog.LogManager.Configuration.Variables("udbsPN") = "torqueCheck"
        NLog.LogManager.Configuration.Variables("janomeJobNum") = mainForm.strJanomeJob
        NLog.LogManager.Configuration.Variables("screwCount") = mainForm.strScrewCount
        NLog.LogManager.Configuration.Variables("sysIn") = mainForm.janome.strSysIn
        NLog.LogManager.Configuration.Variables("SysOut") = mainForm.janome.strSysOut
        NLog.LogManager.Configuration.Variables("fTime") = mainForm.returnPacketScrew(0)
        NLog.LogManager.Configuration.Variables("preset") = mainForm.returnPacketScrew(1)
        NLog.LogManager.Configuration.Variables("tTq") = mainForm.returnPacketScrew(2)
        NLog.LogManager.Configuration.Variables("cTq") = mainForm.returnPacketScrew(3)
        NLog.LogManager.Configuration.Variables("speed") = mainForm.returnPacketScrew(4)
        NLog.LogManager.Configuration.Variables("a1") = mainForm.returnPacketScrew(5)
        NLog.LogManager.Configuration.Variables("a2") = mainForm.returnPacketScrew(6)
        NLog.LogManager.Configuration.Variables("a3") = mainForm.returnPacketScrew(7)
        NLog.LogManager.Configuration.Variables("error") = mainForm.returnPacketScrew(8)
        NLog.LogManager.Configuration.Variables("count") = mainForm.returnPacketScrew(9)
        NLog.LogManager.Configuration.Variables("fL") = mainForm.returnPacketScrew(10)
        NLog.LogManager.Configuration.Variables("status") = mainForm.returnPacketScrew(11)
        NLog.LogManager.Configuration.Variables("torqueRead") = lastHiosResponse

        torqueLogger.Info("Reading")

        Return True

    End Function

    Function monitorResultTorqueCheck(timeLimit As Integer) As Boolean

        Dim response As String
        Dim returnPacket As String()
        Dim currentTime As TimeSpan
        Dim recvData As Boolean = False
        'Dim hiosResponse As List(Of String)
        logger.Trace("==> MonitorResult")
        mainForm.screw.dtResult = mainForm.screw.intialTable()

        Dim hiosColumn As New DataColumn()
        hiosColumn.DataType = System.Type.GetType("System.String")
        hiosColumn.ColumnName = "Hios"
        hiosColumn.DefaultValue = 0

        mainForm.screw.dtResult.Columns.Add(hiosColumn)

        Dim countLoop As Integer = My.Settings.torqueCheckTotalCount * 2
        Dim stopwatch As New Stopwatch()
        stopwatch.Start()

        DataGridView1.DataSource = mainForm.screw.dtResult
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        For i = 0 To countLoop - 1

            Do While stopwatch.ElapsedMilliseconds < timeLimit
                Threading.Thread.Sleep(1000)
                mainForm.janome.readSysIO()
                response = mainForm.screw.sendAndRecvAscii("M1")
                returnPacket = mainForm.screw.recvSmartScrewData(response)
                mainForm.returnPacketScrew = returnPacket

                If response.Length > 10 Then
                    currentTime = DateTime.Now.TimeOfDay
                    'rtb01.Focus()
                    Dim popup As New popupForm()
                    popup.ShowDialog()
                    lastHiosResponse = popup.KeyData
                    logger.Info("lastHiosRes : " & lastHiosResponse)

                    mainForm.screw.dtResult.Rows.Add(i + 1, currentTime, returnPacket(0), returnPacket(1), returnPacket(2), returnPacket(3), returnPacket(4),
                                        returnPacket(5), returnPacket(6), returnPacket(7), returnPacket(8), returnPacket(9), returnPacket(10),
                                        returnPacket(11), lastHiosResponse)
                    logging()
                    DataGridView1.Refresh()
                    Exit Do
                Else
                    logger.Trace("Cannot receive Data in small loop")
                    recvData = False

                End If

            Loop
        Next

        Dim str As String = ""
        Return True
    End Function

    Private Sub torqueCheck_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        rtb01.Focus()
        rtb01.Text = ""
        monitorResultTorqueCheck(90000)
    End Sub

    Private Sub torqueCheck_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        'Console.Write(e.KeyCode.ToString())
        ' MessageBox.Show("Key pressed: " & e.KeyCode.ToString())
        If e.KeyCode = Keys.Enter Then
            logger.Info("Found Enter on frm")
            'MessageBox.Show(getLastResponse(rtb01.Text))
            'logging(getLastResponse(rtb01.Text))
            lastHiosResponse = getLastResponse(hiosResponse)
            logger.Info("Data : " + lastHiosResponse)
            enterCapture = True
        End If

    End Sub

    Private Sub btnFin_Click(sender As Object, e As EventArgs) Handles btnFin.Click
        Me.Close()
    End Sub
End Class