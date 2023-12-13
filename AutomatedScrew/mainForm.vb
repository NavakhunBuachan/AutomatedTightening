Imports System.Net.Sockets
Imports System.Text
Imports System.IO
Imports System.Threading
Imports System.IO.Ports
Imports System.Threading.Tasks
Imports System.Windows.Forms
Public Class mainForm
    Public janome As janomeMachine
    Public screw As smartScrew
    Public config As automatedScrewConfig
    Friend logger As NLog.Logger = NLog.LogManager.GetLogger("Program")
    Friend timeStamplogger As NLog.Logger = NLog.LogManager.GetLogger("TimeStamp")
    Friend snLogger As NLog.Logger = NLog.LogManager.GetLogger("TimeStampSn")

    Public strSn, strUdbsPn, strEn, strJanomeJob, strScrewCount, strMesWipName As String
    Public returnPacketScrew, returnPacketDummy, picPassed, picFailed As String()
    Public strSnLogName As String

    Public testData As testDataInterface

    Private Sub mainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        logger.Trace("mainFormLoading")
        logger.Info("Initial Janome")
        janome = New janomeMachine
        janome.ipAddr = My.Settings.JanomeIP
        janome.port = My.Settings.JanomePort

        logger.Info("Initial smartScrew")
        screw = New smartScrew
        screw.portName = My.Settings.SmartScrewPort
        screw.baudRate = My.Settings.SmartScrewRate

        config = New automatedScrewConfig
        'config.loadConfig("1-01763", "Assy2")
        'config.printRow()
        Dim temp As String = "\00\00\00\12\03\03k00000003000"

        returnPacketDummy = {"None", "None", "None", "None", "None", "None", "None", "None", "None", "None", "None", "None"}

        guiReset()

        PictureBox1.Image = Image.FromFile(My.Application.Info.DirectoryPath + "\Picture\Capture.JPG")
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom


        picPassed = {My.Application.Info.DirectoryPath + "\Picture\Chassis_1_working.jpg",
            My.Application.Info.DirectoryPath + "\Picture\Chassis_2_working.jpg", My.Application.Info.DirectoryPath + "\Picture\Chassis_3_working.jpg",
            My.Application.Info.DirectoryPath + "\Picture\Chassis_4_working.jpg", My.Application.Info.DirectoryPath + "\Picture\Chassis_4_passed.jpg"}

        picFailed = {My.Application.Info.DirectoryPath + "\Picture\Chassis_1_working.jpg",
            My.Application.Info.DirectoryPath + "\Picture\Chassis_1_failed.jpg", My.Application.Info.DirectoryPath + "\Picture\Chassis_2_failed.jpg",
            My.Application.Info.DirectoryPath + "\Picture\Chassis_3_failed.jpg", My.Application.Info.DirectoryPath + "\Picture\Chassis_4_failed.jpg"}

        temp = ""

        'Dim testData = New testDataMES

        'If Not testData.initial("SWTEST02", "BUA66577", System.Net.Dns.GetHostName) Then
        '    testData.dispose()
        '    'Return False
        'End If

        'testData.iMesBuild.StartWip()
        'testData.iMesBuild.SetWipDisposition("FAIL")
        'testData.iMesBuild.EndWip()

    End Sub

    Private Sub btnMESCheck_Click(sender As Object, e As EventArgs) Handles btnMESCheck.Click

        logger.Info("==>MESCheck")
        printOutput("MES Check: Checking MES status and Application configuration", Color.Orange)

        If checkTbInput() Then
            strSn = tbSn.Text
            strEn = tbEn.Text

            If Not testData.initial(strSn, strEn, System.Net.Dns.GetHostName) Then
                testData.dispose()
            Else
                printOutput("Cannot initial MES, please contact Engineer", Color.Red)
                Return
            End If

        Else
            printOutput("Please input SN and EN in text box", Color.Red)
            Return

        End If


        If Not checkAndLoadConfig(testData.UDBSPartNumber, testData.wipstage) Then
            Return
        End If

        strJanomeJob = config.jobNum
        strScrewCount = config.screwNum

        If Not checkMachinesConnected() Then

            printOutput("MES Check: Cannot communicate to Janome/SmartScrew", Color.Red)
            Return
        End If

        logger.Debug("GetJob :" + janome.getJobNum() + "===" + config.jobNum)
        If janome.getJobNum() <> config.jobNum Then
            logger.Info("Set JobNum = " + config.jobNum)
            janome.setJobNum(config.jobNum)
        End If
        janome.readSysIO()
        janome.readGenIo()
        testData.dispose()

        timestampLog("MES/Conf Check is Done", False)

        printOutput("Load the unit then click 'Start'", Color.Green)
        btnMESCheck.Enabled = False
        btnMESCheck.BackColor = Color.Green
        btnStart.Enabled = True
        'btnLoadUnit.Enabled = True
        logger.Info("<==MESCheck")
    End Sub


    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        logger.Info("==> Start")
        logger.Info("CheckStatus before run")

        If Not testData.initial(strSn, strEn, System.Net.Dns.GetHostName) Then
            testData.dispose()
        Else
            printOutput("Cannot initial MES, please contact Engineer", Color.Red)
            Return
        End If

        testData.StartTest(testData.SerialNumber, testData.UDBSPartNumber)

        If Not janome.readyToStart() Then
            printOutput("Please check the unit was loaded properly", Color.Red)

            Return
        End If

        printOutput("Machines is running", Color.Orange)

        janome.readSysIO()
        janome.readGenIo()
        'janome.setSysOut("8")
        'Threading.Thread.Sleep(100)
        'janome.resetSysOut("8")
        'readSysIoLoop(30000)
        timestampLog("Job Start", False)

        janome.startJob()

        If monitorResult(5000) Then

            printOutput("Unit Passed. Please click 'Finish' then unload the unit", Color.Green)
            btnStart.BackColor = Color.Green
            btnStart.Enabled = False
            btnFin.Enabled = True
        Else
            printOutput("Unit Failed. Press 'Emergency' followed by 'Initial' button on machine", "Then unload unit", Color.Red)

        End If
        timestampLog("Job End", False)
        'logger.Info("Done")
        logger.Info("<== Start")
    End Sub
    Private Sub btnFin_Click(sender As Object, e As EventArgs) Handles btnFin.Click



        guiReset()
    End Sub

    Private Function storeTestData() As Boolean

        For Each row In screw.dtResult.Rows

        Next

        For i = 0 To screw.dtResult.Rows.Count - 1

            testData.TestInst.StoreValue("Screw_1", screw.dtResult.Rows(0).Item(2))


        Next

        'testData.TestInst.StoreValue()

    End Function


    Function janomeStart() As Boolean
        Try
            janome.setSysOut("8")
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Sub readSysIoLoop(timeLimit As Integer)
        Dim stopwatch As New Stopwatch()
        stopwatch.Start()

        Do While stopwatch.ElapsedMilliseconds < timeLimit

            Threading.Thread.Sleep(100)
            janome.readSysIO()
            janome.readGenIo()

        Loop

        stopwatch.Stop()

    End Sub
    Function monitorResult(timeLimit As Integer) As Boolean
        Dim i As Integer = 0
        Dim num As Integer = 1
        Dim countLoop As Integer = 0
        Dim response As String
        Dim returnPacket As String()
        Dim currentTime As TimeSpan
        Dim recvData As Boolean = False
        logger.Trace("==> MonitorResult")

        screw.dtResult = screw.intialTable()
        DataGridView1.DataSource = screw.dtResult
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        DataGridView1.ScrollBars = ScrollBars.Both

        countLoop = config.screwNum * 2
        'intScrewCount = countLoop
        PictureBox1.Image = Image.FromFile(picPassed(0))
        PictureBox1.Refresh()

        For i = 0 To countLoop - 1
            'Threading.Thread.Sleep(1000) ' for delay
            Dim picIndex As Integer = Math.Floor((i + 1) / 2)
            ' Dim timeLimit As Integer = 5000
            Dim stopwatch As New Stopwatch()
            stopwatch.Start()
            recvData = False

            Do While stopwatch.ElapsedMilliseconds < timeLimit

                Threading.Thread.Sleep(1000)
                janome.readSysIO()
                response = screw.sendAndRecv("02 4D 31 45 03")

                If response.Length > 10 Then

                    PictureBox1.Image = Image.FromFile(picPassed(picIndex))
                    PictureBox1.Refresh()
                    returnPacket = screw.recvSmartScrewData(response)
                    returnPacketScrew = returnPacket
                    currentTime = DateTime.Now.TimeOfDay
                    logger.Trace("==> Add data to dt")
                    'screw.dtResult.Rows.Add(i, currentTime, returnPacket(0).Substring(returnPacket(0).Length - 5, 5), returnPacket(1), returnPacket(2), returnPacket(3), returnPacket(4),
                    '            returnPacket(5), returnPacket(6), returnPacket(7), returnPacket(8), returnPacket(9), returnPacket(10),
                    '            returnPacket(11))
                    screw.dtResult.Rows.Add(i, currentTime, returnPacket(0), returnPacket(1), returnPacket(2), returnPacket(3), returnPacket(4),
                                returnPacket(5), returnPacket(6), returnPacket(7), returnPacket(8), returnPacket(9), returnPacket(10),
                                returnPacket(11))
                    timestampLog("SmartScrew " + Str(i), True)
                    logger.Trace("<== Add data to dt")
                    recvData = True
                    RefreshDataGridView()

                    Exit Do
                Else

                End If

            Loop

            stopwatch.Stop()

            If Not recvData Then
                PictureBox1.Image = Image.FromFile(picFailed(picIndex))
                PictureBox1.Refresh()
                logger.Trace("Cannot receive Data in time limit")
                Exit For
            End If

        Next
        logger.Trace("Compelete receive Data")
        Return recvData
    End Function


    Sub RefreshDataGridView()
        logger.Trace("==>RefreshDataGridView")
        'DataGridView1.ScrollBars = ScrollBars.Both
        DataGridView1.ScrollBars = ScrollBars.Vertical
        If DataGridView1.Rows.Count > 0 Then
            DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.Rows.Count - 1
        End If
        DataGridView1.Refresh()
        logger.Trace("<==RefreshDataGridView")
    End Sub


    Private Function checkAndLoadConfig(pn As String, proc As String) As Boolean

        If Not config.checkProc(proc) Then
            logger.Error("The unit is not in this process, please check WIP")
            printOutput("The unit is not in this process, please check WIP", Color.Red)
        End If

        If Not config.hasConfig(pn, proc) Then
            logger.Error("Not found partNumber in configuration, please contact engineer")
            printOutput("Not found partNumber in configuration, please contact engineer", Color.Red)
            Return False
        End If

        logger.Info("checkAndLoadConfig is done")
        Return True
        'If config.hasConfig(pn, proc) Then
        '    Return config.loadConfig(pn, proc)
        'Else
        '    logger.Error("Not found partNumber/ProcessName in configuration")
        '    Return False
        'End If

    End Function
    Public Function checkMachinesConnected() As Boolean

        checkMachinesConnected = janome.connect() And screw.connect()

        If Not checkMachinesConnected Then
            logger.Error("Cannot connect to Machines")
        End If
        logger.Info("Janome and SmartScrew is connected")

    End Function

    Function checkTbInput() As Boolean
        If tbEn.Text.Length <> 0 And tbEn.Text.Length <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function
#Region "Log"
    Function serialLogInitial() As String

        Dim config As NLog.Config.LoggingConfiguration = NLog.LogManager.Configuration

        strSnLogName = "${basedir}\" + strUdbsPn + "_" + strSn + "_" + strMesWipName + "_"

        For Each target As NLog.Targets.FileTarget In config.AllTargets

            If target.Name = "TimeStampSn" Then
                target.FileName = "${basedir}\TimeStampSn_Test_log.csv"
            End If

        Next

        Return strSnLogName

    End Function

    Private Function timestampLog(str As String, isAutoScrew As Boolean) As Boolean

        'lbStatus.Text = str
        NLog.LogManager.Configuration.Variables("sn") = strSn
        NLog.LogManager.Configuration.Variables("udbsPN") = strUdbsPn
        NLog.LogManager.Configuration.Variables("janomeJobNum") = strJanomeJob
        NLog.LogManager.Configuration.Variables("screwCount") = strScrewCount

        If isAutoScrew Then

            NLog.LogManager.Configuration.Variables("sysIn") = janome.strSysIn
            NLog.LogManager.Configuration.Variables("SysOut") = janome.strSysOut
            NLog.LogManager.Configuration.Variables("fTime") = returnPacketScrew(0)
            NLog.LogManager.Configuration.Variables("preset") = returnPacketScrew(1)
            NLog.LogManager.Configuration.Variables("tTq") = returnPacketScrew(2)
            NLog.LogManager.Configuration.Variables("cTq") = returnPacketScrew(3)
            NLog.LogManager.Configuration.Variables("speed") = returnPacketScrew(4)
            NLog.LogManager.Configuration.Variables("a1") = returnPacketScrew(5)
            NLog.LogManager.Configuration.Variables("a2") = returnPacketScrew(6)
            NLog.LogManager.Configuration.Variables("a3") = returnPacketScrew(7)
            NLog.LogManager.Configuration.Variables("error") = returnPacketScrew(8)
            NLog.LogManager.Configuration.Variables("count") = returnPacketScrew(9)
            NLog.LogManager.Configuration.Variables("fL") = returnPacketScrew(10)
            NLog.LogManager.Configuration.Variables("status") = returnPacketScrew(11)
        Else

            NLog.LogManager.Configuration.Variables("sysIn") = janome.strSysIn
            NLog.LogManager.Configuration.Variables("SysOut") = janome.strSysOut
            NLog.LogManager.Configuration.Variables("fTime") = returnPacketDummy(0)
            NLog.LogManager.Configuration.Variables("preset") = returnPacketDummy(1)
            NLog.LogManager.Configuration.Variables("tTq") = returnPacketDummy(2)
            NLog.LogManager.Configuration.Variables("cTq") = returnPacketDummy(3)
            NLog.LogManager.Configuration.Variables("speed") = returnPacketDummy(4)
            NLog.LogManager.Configuration.Variables("a1") = returnPacketDummy(5)
            NLog.LogManager.Configuration.Variables("a2") = returnPacketDummy(6)
            NLog.LogManager.Configuration.Variables("a3") = returnPacketDummy(7)
            NLog.LogManager.Configuration.Variables("error") = returnPacketDummy(8)
            NLog.LogManager.Configuration.Variables("count") = returnPacketDummy(9)
            NLog.LogManager.Configuration.Variables("fL") = returnPacketDummy(10)
            NLog.LogManager.Configuration.Variables("status") = returnPacketDummy(11)
        End If

        timeStamplogger.Info(str)
        snLogger.Info(str)

        Return True
    End Function
    Sub loggerTest()

        'testLogger.Factory.CreateNullLogger()

        Dim baseName As String = IO.Path.GetFileName(My.Application.Info.DirectoryPath + "\test.txt")

        Dim fileTarget As New NLog.Targets.FileTarget() With {
               .Name = baseName,
               .FileName = My.Application.Info.DirectoryPath + "\test.txt",
               .Layout = "${longdate}|${level}|${logger}|${gdc:item=OptionalLayout}${mdc:item=OptionalPerThreadLayout}${message} ${exception:format=tostring}",
               .ArchiveAboveSize = 50000000,
               .ArchiveNumbering = 5,
               .MaxArchiveFiles = 5,
               .EnableArchiveFileCompression = False,
               .ConcurrentWrites = False,
               .Encoding = Encoding.UTF8
           }

        Dim config As NLog.Config.LoggingConfiguration = NLog.LogManager.Configuration
        ' If config Is Nothing Then config = New Config.LoggingConfiguration
        'NLog.LogManager.ReconfigExistingLoggers()
        Dim fileTarget2 As New NLog.Targets.FileTarget()
        Dim FileRule2 = New NLog.Config.LoggingRule()
        snLogger.Info("Hello")
        For Each target As NLog.Targets.FileTarget In config.AllTargets

            If target.Name = "TimeStampSn" Then
                target.FileName = "${basedir}\TimeStampSn_Test_log.csv"
            End If

        Next
        snLogger.Info("Hello")
        ' NLog.LogManager.ReconfigExistingLoggers()
        snLogger.Info("Hello")
        timestampLog("test", False)
        ''Dim selectedIndex As Integer = 0

        ''For i = 0 To config.AllTargets.Count - 1
        ''    If config.AllTargets(i).Name = "TimeStampSn" Then
        ''        selectedIndex = i
        ''    End If
        ''Next

        ''fileTarget2 = config.AllTargets(selectedIndex)
        ''fileTarget2.Name = "Test3"
        ''fileTarget2.FileName = "${basedir}\TimeStamp_log2.csv"
        ''FileRule2 = config.LoggingRules(selectedIndex)
        ''FileRule2.LoggerNamePattern = "Test3"
        'config.AddTarget(fileTarget)

        'Dim logfile = New NLog.Targets.FileTarget("logfile")

        'Dim sToday As String = DateTime.Today.ToString("yyyy-MM-dd")
        'Dim FileTarget = New NLog.Targets.FileTarget() With {.FileName = sLogdir + sToday + ".log"}
        'Dim ConsoleTarget As New NLog.Targets.ConsoleTarget
        'ConsoleTarget.Layout = "${date:format=HH}:${date:format=mm}:${date:format=ss}: ${message}"
        'config.AddTarget("logfile", fileTarget)
        'config.AddTarget("console", ConsoleTarget)
        'Dim FileRule = New NLog.Config.LoggingRule("*", LogLevel.Debug, fileTarget)
        'Dim ConsoleRule = New NLog.Config.LoggingRule("*", LogLevel.Debug, ConsoleTarget)
        'config.LoggingRules.Add(FileRule)
        'config.LoggingRules.Add(ConsoleRule)
        'LogManager.Configuration = config
        '_logger = LogManager.GetCurrentClassLogger()
        '_logger.Info("Logger initialized.")

        Dim t As String = ""

    End Sub

    'Private Sub btnLoadUnit_Click(sender As Object, e As EventArgs)

    '    logger.Info("CheckStatus before run")

    '    If janome.readyToStart() Then
    '        lbOutput.BackColor = Color.Green
    '        lbOutput.Text = "Ready to Start"

    '        btnLoadUnit.Enabled = False
    '        btnLoadUnit.BackColor = Color.Green
    '        btnStart.Enabled = True

    '    Else
    '        lbOutput.BackColor = Color.Red
    '        lbOutput.Text = "Please check the unit was loaded properly"
    '    End If

    'End Sub
#End Region

#Region "GUI"
    Public Sub guiReset()
        lbOutput.BackColor = SystemColors.Control
        lbOutput2.BackColor = SystemColors.Control
        btnMESCheck.BackColor = SystemColors.Control
        'btnLoadUnit.BackColor = SystemColors.Control
        btnStart.BackColor = SystemColors.Control
        btnFin.BackColor = SystemColors.Control

        lbOutput.Text = ""
        lbOutput2.Text = ""
        tbSn.Text = ""

        btnMESCheck.Enabled = True
        'btnLoadUnit.Enabled = False
        btnStart.Enabled = False
        btnFin.Enabled = False

    End Sub
    Public Sub printOutput(strIn1 As String, color As Color)
        lbOutput.Text = strIn1
        lbOutput.BackColor = color
    End Sub
    Public Sub printOutput(strIn1 As String, strIn2 As String, color As Color)
        lbOutput.Text = strIn1
        lbOutput2.Text = strIn2
        lbOutput.BackColor = color
        lbOutput2.BackColor = color
    End Sub
#End Region
End Class
