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
    Public config As automatedTighteningConfig
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

        config = New automatedTighteningConfig

        testData = New testDataMES

        returnPacketDummy = {"None", "None", "None", "None", "None", "None", "None", "None", "None", "None", "None", "None"}

        Me.Text = "Automated Tightening v" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor

        guiReset()
        'forTest()

    End Sub

    Private Sub btnMESCheck_Click(sender As Object, e As EventArgs) Handles btnMESCheck.Click
        Dim checkresult As Boolean = True
        logger.Info("==> MESCheck")
        printOutput("MES Check: Checking MES status and Application configuration", Color.Orange)

        If checkTbInput() Then
            strSn = tbSn.Text
            strEn = tbEn.Text

            If Not testData.initial(strSn, strEn, System.Net.Dns.GetHostName) Then
                testData.dispose()
                printOutput("Cannot initial MES, please contact Engineer", Color.Red)
                checkresult = False
                Return
            Else

            End If
            logger.Info("MES initial Done")
        Else
            printOutput("Please input SN and EN in text box", Color.Red)
            checkresult = False
            Return

        End If

        If Not checkAndLoadConfig(testData.UDBSPartNumber, testData.wipstage) Then
            Return
        End If

        strJanomeJob = config.jobNum
        strScrewCount = config.screwNum
        loadPicture(config.jobNum)

        If Not checkMachinesConnected() Then
            printOutput("MES Check: Cannot communicate to Janome/SmartScrew", Color.Red)
            testData.dispose()
            Return
        End If

        logger.Debug("GetJob Janome :" + janome.getJobNum())
        logger.Debug("GetJob Config :" + config.jobNum)

        If janome.getJobNum() <> config.jobNum Then
            logger.Info("Set JobNum = " + config.jobNum)
            janome.setJobNum(config.jobNum)

        Else
            logger.Info("Don't need to set job No.")
        End If

        If Not checkJanomeWaitingRun(10000) Then
            printOutput("Press 'Emergency' followed by 'Initial' button on machine", Color.Red)
            'printOutput("Machine need to initialize, please contact engineer", Color.Red)
            testData.dispose()
            Return
        End If

        testData.dispose()

        printOutput("Load the unit then click 'Start'", Color.Green)
        btnMESCheck.Enabled = False
        btnMESCheck.BackColor = Color.Green
        btnStart.Enabled = True
        tbEn.Enabled = False
        tbSn.Enabled = False
        PictureBox1.Image = Image.FromFile(My.Application.Info.DirectoryPath + "\Picture\Check.png")
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        btnStart.Focus()
        logger.Info("<== MESCheck")
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        logger.Info("==> Start")
        logger.Info("CheckStatus before run")

        If Not testData.initial(strSn, strEn, System.Net.Dns.GetHostName) Then
            testData.dispose()
            printOutput("Cannot initial MES, please contact Engineer", Color.Red)
            Return
        Else

        End If
        'Dim str2 As String = UdbsInterface.MasterInterface.UdbsProcessStatus.
        strUdbsPn = testData.UDBSPartNumber
        strMesWipName = testData.wipstage

        unitLogInitial()

        If Not testData.StartTest(testData.SerialNumber, testData.UDBSPartNumber) Then
            printOutput("Cannot Start WIP", Color.Red)
            guiReset()
            Return
        End If

        If Not janome.readyToStart() Then
            printOutput("Please check the unit was loaded properly", Color.Red)
            Return
        End If

        printOutput("Machines is running", Color.Orange)

        timestampLog("Job Start", False)

        janome.startJob()

        If monitorResult(30000) Then

            'printOutput("Unit Passed. Please click 'Finish' then unload the unit", Color.Green)
            btnStart.BackColor = Color.Green

        Else
            ' printOutput("Unit Failed. Press 'Emergency' followed by 'Initial' button on machine", "Then Unload Unit", Color.Red)
            btnStart.BackColor = Color.Red
        End If

        'If Not checkJanomeWaitingRun(10000) Then
        '    'printOutput("Machine need to initialize, please contact engineer", Color.Red)
        '    logger.Info("Machine status is not normal after testing done, please check")
        'End If

        timestampLog("Job End", False)
        'logger.Info("Done")
        ' btnFin.Focus()
        logger.Info("<== Start")

        btnStart.Enabled = False
        'btnFin.Enabled = True

        finishTest()

    End Sub
    Function finishTest() As Boolean
        logger.Info("==> Finish")

        If Not storeTestData() Then
            Return False
        End If

        Dim rc As UdbsInterface.TestDataInterface.ResultCodes
        Dim strRouting As String = ""
        guiReset()
        rc = testData.TestInst.EvaluateDevice()
        If rc = UdbsInterface.TestDataInterface.ResultCodes.UDBS_SPECS_PASS Then
            strRouting = testData.passRouting
            testData.FinishTest("")
            'printOutput("Unit Passed. Please click 'Finish' then unload the unit", Color.Green)
            printOutput(strSn + " result is PASS", "Please unload the unit", Color.Green)
            finishTest = True
            logger.Info("Pass")
        Else
            strRouting = testData.failRouting
            testData.FinishTest("")
            'printOutput("Unit Failed. Press 'Emergency' followed by 'Initial' button on machine", "Then Unload Unit", Color.Red)
            printOutput(strSn + " result is FAILED", "Press 'Emergency' followed by 'Initial' button on machine Then Unload Unit", Color.Red)
            finishTest = False
            logger.Info("Fail")
        End If

        logger.Info("<== Finish")
    End Function
    'Private Sub btnFin_Click(sender As Object, e As EventArgs) Handles btnFin.Click
    '    logger.Info("==> Finish")

    '    If Not storeTestData() Then
    '        Return
    '    End If

    '    Dim rc As UdbsInterface.TestDataInterface.ResultCodes
    '    Dim strRouting As String = ""
    '    guiReset()
    '    rc = testData.TestInst.EvaluateDevice()
    '    If rc = UdbsInterface.TestDataInterface.ResultCodes.UDBS_SPECS_PASS Then
    '        strRouting = testData.passRouting
    '        testData.FinishTest("")
    '        printOutput(strSn + " result is PASS", Color.Green)
    '        logger.Info("Pass")
    '    Else
    '        strRouting = testData.failRouting
    '        testData.FinishTest("")
    '        printOutput(strSn + " result is FAILED", Color.Red)
    '        logger.Info("Fail")
    '    End If

    '    'janome.close()
    '    'screw.close()
    '    logger.Info("<== Finish")
    'End Sub

    Function loadPicture(jobNum As String) As Boolean

        picPassed = System.IO.Directory.GetFiles(My.Application.Info.DirectoryPath + "\Picture\" + jobNum + "\Passed\")
        System.Array.Sort(Of String)(picPassed)
        picFailed = System.IO.Directory.GetFiles(My.Application.Info.DirectoryPath + "\Picture\" + jobNum + "\Failed\")
        System.Array.Sort(Of String)(picFailed)


        Return True
    End Function
    Function checkJanomeWaitingRun(waitingTime As Integer) As Boolean
        checkJanomeWaitingRun = False
        Dim stopwatch As New Stopwatch()
        stopwatch.Start()

        Do While stopwatch.ElapsedMilliseconds < waitingTime
            Threading.Thread.Sleep(1000)
            printOutput("Waiting Machine running", Color.Orange)
            If janome.getStatus() = "Waiting run" Then
                checkJanomeWaitingRun = True
                Exit Do
            End If

        Loop
    End Function
    Private Function storeTestData() As Boolean
        logger.Info("==> StoreData")
        Dim strResultName As String = ""
        Dim screwCount As Integer = 1
        Dim rc As UdbsInterface.TestDataInterface.ResultCodes
        'logger.Info("screw.dtResult.Rows.Count-1 = " & CStr(screw.dtResult.Rows.Count - 1))
        logger.Info("dtResult.Columns.Count = " + CStr(screw.dtResult.Columns.Count))

        For i = 0 To screw.dtResult.Rows.Count - 1

            If (i Mod 2) = 0 Then 'Feeding result is in odd row

                strResultName = "Screw_" & CStr(screwCount) & "_Feed"
                strResultName = strResultName.ToLower
                If screw.dtResult.Rows(i).Item(10) = 0 Then ' Item(10) is error code from smartScrew, 0 is no error
                    rc = testData.TestInst.StoreValue(strResultName, CDbl(0))
                    logger.Info("Store =>" + strResultName + "=0" + " ==" + CStr(rc))
                Else
                    testData.TestInst.StoreValue(strResultName, CDbl(1))
                    logger.Info("Store =>" + strResultName + "=1" + " ==" + CStr(rc))
                End If

            Else ' Screw result in even row

                For colIndex = 2 To screw.dtResult.Columns.Count - 1

                    strResultName = "Screw_" & CStr(screwCount) & "_" & screw.dtResult.Columns(colIndex).ColumnName
                    rc = testData.TestInst.StoreValue(strResultName.ToLower, CDbl(screw.dtResult.Rows(i).Item(colIndex)))
                    logger.Info("Store =>" + strResultName.ToLower + "=" + CStr(screw.dtResult.Rows(i).Item(colIndex)) + " ==" + CStr(rc))

                    If screw.dtResult.Columns(colIndex).ColumnName = "Status" Then ' Count status result as screw result
                        testData.TestInst.StoreValue("Screw_" & CStr(screwCount), screw.dtResult.Rows(i).Item(colIndex))
                        logger.Info("Store =>" + strResultName + "=" + CStr(screw.dtResult.Rows(i).Item(colIndex)))
                    End If

                Next
                screwCount = screwCount + 1
            End If
            If rc < -10000 Then
                logger.Error("Store Data Error")
                Return False
            End If
        Next

        Dim testResult As UdbsInterface.TestDataInterface.CTestData_Result = testData.TestInst.Results("screw_1")

        testResult.StoreFile("LogFile", "GroupName", False, strSnLogName)

        'testData.TestInst.st
        logger.Info("<== StoreData")
        Return True
        ' testData.TestInst.StoreValue()

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

        Dim response As String
        Dim returnPacket As String()
        Dim currentTime As TimeSpan
        Dim recvData As Boolean = False
        logger.Trace("==> MonitorResult")

        screw.dtResult = screw.intialTable()

        Dim countLoop As Integer = config.screwNum * 2

        PictureBox1.Image = Image.FromFile(picPassed(0))
        PictureBox1.Refresh()

        Dim picindex As Integer = 0
        Dim resultScrew As Boolean = False

        For i = 0 To countLoop - 1

            'picindex = Math.Floor((i + 1) / 2)

            Dim stopwatch As New Stopwatch()
            stopwatch.Start()
            recvData = False
            PictureBox1.Image = Image.FromFile(picPassed(picindex))
            PictureBox1.Refresh()

            Do While stopwatch.ElapsedMilliseconds < timeLimit

                Threading.Thread.Sleep(1000)
                janome.readSysIO()
                'janome.getStatus()
                response = screw.sendAndRecv("02 4D 31 45 03")

                If response.Length > 10 Then

                    recvData = True
                    returnPacket = screw.recvSmartScrewData(response)
                    returnPacketScrew = returnPacket
                    currentTime = DateTime.Now.TimeOfDay

                    screw.dtResult.Rows.Add(i, currentTime, returnPacket(0), returnPacket(1), returnPacket(2), returnPacket(3), returnPacket(4),
                                returnPacket(5), returnPacket(6), returnPacket(7), returnPacket(8), returnPacket(9), returnPacket(10),
                                returnPacket(11))
                    timestampLog("SmartScrew " + Str(i), True)
                    logger.Trace("Count = " + CStr(screw.dtResult.Rows.Count))
                    If screw.dtResult.Rows.Count Mod 2 = 1 Then 'screw feeding
                        logger.Trace("Feeding")
                        logger.Trace(returnPacket(8))
                        If returnPacket(8) = "000" Then 'No error
                            resultScrew = True
                        Else
                            logger.Trace("Screw feeding Error")
                            resultScrew = False
                        End If
                    ElseIf screw.dtResult.Rows.Count Mod 2 = 0 Then 'screwing
                        logger.Trace("Screwing")
                        picindex = picindex + 1
                        If returnPacket(11) = 1 Then 'Status is good for screwing
                            resultScrew = True
                        Else
                            logger.Trace("Screwing result failed")
                            resultScrew = False
                        End If

                    End If

                    logger.Info("Picindex =" & picindex)
                    PictureBox1.Image = Image.FromFile(picPassed(picindex))
                    PictureBox1.Refresh()

                    Exit Do
                Else
                    logger.Trace("Cannot receive Data in small loop")
                    recvData = False
                End If
                logger.Info("Picindex =" & picindex)
            Loop

            stopwatch.Stop()

            ' Failure checking
            If (Not recvData) Or (janome.getStatus() <> "Running (moving)") Or (Not resultScrew) Then
                monitorResult = False
                PictureBox1.Image = Image.FromFile(picFailed(picindex))
                PictureBox1.Refresh()
                logger.Trace("Cannot receive Data in time limit/ Machine is not running normally/ Screwing failed")
                Exit For
            Else
                monitorResult = True
            End If

        Next

        logger.Trace("Compelete receive Data loop")
        Return monitorResult

    End Function


    Private Function checkAndLoadConfig(pn As String, proc As String) As Boolean

        If Not config.checkProc(proc) Then
            logger.Error("The unit is not in this process, please check WIP")
            printOutput("The unit is not in this process, please check WIP", Color.Red)
            Return False
        End If

        If Not config.hasConfig(pn, proc) Then
            logger.Error("Not found partNumber in configuration, please contact engineer")
            printOutput("Not found partNumber in configuration, please contact engineer", Color.Red)
            Return False
        End If

        Return config.loadConfig(pn, proc)

    End Function

    Public Function checkMachinesConnected() As Boolean

        checkMachinesConnected = janome.connect() And screw.connect()

        If Not checkMachinesConnected Then
            logger.Error("Cannot connect to Machines")
        End If
        logger.Info("Janome and SmartScrew are connected")

    End Function

    Function checkTbInput() As Boolean
        If tbEn.Text.Length <> 0 And tbEn.Text.Length <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function
#Region "Log"
    Function unitLogInitial() As String

        Dim config As NLog.Config.LoggingConfiguration = NLog.LogManager.Configuration

        strSnLogName = "C:\ProgramData\AutomatedScrew\" + strUdbsPn + "_" + strSn + "_" + strMesWipName + "_" + CStr(testData.calTestSeqByWipStage) + ".csv"
        logger.Info(strSnLogName)
        For Each target As NLog.Targets.FileTarget In config.AllTargets

            If target.Name = "TimeStampSn" Then
                target.FileName = strSnLogName
                'target.FileName = "${basedir}\TimeStampSn_Test_log.csv"
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

        PictureBox1.Image = Image.FromFile(My.Application.Info.DirectoryPath + "\Picture\Initial.png")
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom

        lbOutput.BackColor = SystemColors.Control
        lbOutput2.BackColor = SystemColors.Control
        btnMESCheck.BackColor = SystemColors.Control
        'btnLoadUnit.BackColor = SystemColors.Control
        btnStart.BackColor = SystemColors.Control
        btnFin.BackColor = SystemColors.Control

        lbOutput.Text = ""
        lbOutput2.Text = ""
        tbSn.Text = ""
        tbEn.Text = ""

        tbSn.Enabled = True
        tbEn.Enabled = True
        btnMESCheck.Enabled = True
        'btnLoadUnit.Enabled = False
        btnStart.Enabled = False
        btnFin.Enabled = False
        btnFin.Visible = False
        btnInit.Visible = False
        btnInit.Enabled = False

    End Sub
    Public Sub printOutput(strIn1 As String, color As Color)
        lbOutput.Text = strIn1
        lbOutput2.Text = ""
        lbOutput.BackColor = color
        lbOutput.Refresh()
        lbOutput2.BackColor = color
        lbOutput2.Refresh()
    End Sub
    Public Sub printOutput(strIn1 As String, strIn2 As String, color As Color)
        lbOutput.Text = strIn1
        lbOutput2.Text = strIn2
        lbOutput.BackColor = color
        lbOutput2.BackColor = color
        lbOutput.Refresh()
        lbOutput2.Refresh()
    End Sub
#End Region

#Region "Backup"
    Sub forTest()
        Dim temp As String = "\00\00\00\12\03\03k00000003000"
        Dim test As Integer = Math.Floor((2 + 1) / 2)
        Dim testDouble As Double = CDbl("0")
        temp = ""
        ' PictureBox1.Image = Image.FromFile(My.Application.Info.DirectoryPath + "\Picture\121\Chassis.JPG")
        'PictureBox1.Image = Image.FromFile(My.Application.Info.DirectoryPath + "\Picture\Start.png")
        'PictureBox1.SizeMode = PictureBoxSizeMode.Zoom

        'Dim strFiles As String() = System.IO.Directory.GetFiles(My.Application.Info.DirectoryPath + "\Picture\121\Passed\")
        'System.Array.Sort(Of String)(strFiles)
        'picPassed = strFiles

        'picPassed = {My.Application.Info.DirectoryPath + "\Picture\Chassis_1_working.jpg",
        '    My.Application.Info.DirectoryPath + "\Picture\Chassis_2_working.jpg", My.Application.Info.DirectoryPath + "\Picture\Chassis_3_working.jpg",
        '    My.Application.Info.DirectoryPath + "\Picture\Chassis_4_working.jpg", My.Application.Info.DirectoryPath + "\Picture\Chassis_4_passed.jpg"}

        'picFailed = {My.Application.Info.DirectoryPath + "\Picture\Chassis_1_working.jpg",
        '    My.Application.Info.DirectoryPath + "\Picture\Chassis_1_failed.jpg", My.Application.Info.DirectoryPath + "\Picture\Chassis_2_failed.jpg",
        '    My.Application.Info.DirectoryPath + "\Picture\Chassis_3_failed.jpg", My.Application.Info.DirectoryPath + "\Picture\Chassis_4_failed.jpg"}

        'loadPicture("121")

        ''Dim testData2 As UdbsInterface.TestDataInterface.CTestData_Result

        ''testData2.StoreFile()

        'If Not testData.initial("XCNE234ZTE001333-B", "BUA66577", System.Net.Dns.GetHostName) Then
        If Not testData.initial("SWTEST02", "BUA66577", System.Net.Dns.GetHostName) Then
            testData.dispose()
            'Return False
        End If

        If Not testData.StartTest(testData.SerialNumber, testData.UDBSPartNumber) Then
            printOutput("Cannot Start WIP", Color.Red)
            guiReset()
            Return
        End If
        'testData.FinishTest("")
        'testData.ime
        Dim testResult As UdbsInterface.TestDataInterface.CTestData_Result = testData.TestInst.Results("screw_1")
        'testResult.ItemName
        temp = My.Application.Info.DirectoryPath
        testResult.StoreFile("testdata_TRHUA20864-B_mt_insp_1.txt", "GroupName", False, "C:\testdata_TRHUA20864-B_mt_insp_1.txt")
        testData.FinishTest("")
        temp = ""
    End Sub
    Sub RefreshDataGridView()
        logger.Trace("==>RefreshDataGridView")
        'DataGridView1.ScrollBars = ScrollBars.Both
        'DataGridView1.ScrollBars = ScrollBars.Vertical
        'If DataGridView1.Rows.Count > 0 Then
        ' DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.Rows.Count - 1
        'End If
        'DataGridView1.Refresh()
        logger.Trace("<==RefreshDataGridView")
    End Sub
    Private Sub btnInit_Click(sender As Object, e As EventArgs) Handles btnInit.Click
        janome.getStatus()
    End Sub
    Function monitorResultOld(timeLimit As Integer) As Boolean
        Dim i As Integer = 0
        Dim num As Integer = 1
        Dim countLoop As Integer = 0
        Dim response As String
        Dim returnPacket As String()
        Dim currentTime As TimeSpan
        Dim recvData As Boolean = False
        logger.Trace("==> MonitorResult")

        screw.dtResult = screw.intialTable()
        ' DataGridView1.DataSource = screw.dtResult
        'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        ' DataGridView1.ScrollBars = ScrollBars.Both

        countLoop = config.screwNum * 2
        'intScrewCount = countLoop
        PictureBox1.Image = Image.FromFile(picPassed(0))
        PictureBox1.Refresh()

        Dim picindex As Integer = 0
        Dim resultScrew As Boolean = False

        For i = 0 To countLoop - 1

            'picindex = Math.Floor((i + 1) / 2)

            Dim stopwatch As New Stopwatch()
            stopwatch.Start()
            recvData = False
            PictureBox1.Image = Image.FromFile(picPassed(picindex))
            PictureBox1.Refresh()

            Do While stopwatch.ElapsedMilliseconds < timeLimit

                Threading.Thread.Sleep(1000)
                janome.readSysIO()
                response = screw.sendAndRecv("02 4D 31 45 03")

                If response.Length > 10 Then

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

                    If (screw.dtResult.Rows.Count Mod 2 = 0) Then 'And returnPacket(11) = "1" Then
                        picindex = picindex + 1
                        logger.Info("Picindex =" & picindex)
                        resultScrew = True
                        'ElseIf screw.dtResult.Rows.Count = countLoop Then
                        ' picindex = picindex + 1

                    End If

                    PictureBox1.Image = Image.FromFile(picPassed(picindex))
                    PictureBox1.Refresh()

                    Exit Do
                Else

                End If
                logger.Info("Picindex =" & picindex)
            Loop

            stopwatch.Stop()

            If Not recvData Then

                'If i Mod 2 = 1 Then
                '    picIndex = picIndex + 1
                'End If

                PictureBox1.Image = Image.FromFile(picFailed(picindex))
                PictureBox1.Refresh()
                logger.Trace("Cannot receive Data in time limit")
                Exit For
            End If

        Next
        logger.Trace("Compelete receive Data")
        Return recvData
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        'If Not testData.initial(strSn, strEn, System.Net.Dns.GetHostName) Then
        '    testData.dispose()
        '    printOutput("Cannot initial MES, please contact Engineer", Color.Red)
        '    Return
        'Else

        'End If
        'Dim str2 As String = UdbsInterface.MasterInterface.UdbsProcessStatus.
        'strUdbsPn = testData.UDBSPartNumber
        'strMesWipName = testData.wipstage

        ' testData.StartTest(testData.SerialNumber, testData.UDBSPartNumber)

        Dim rc As UdbsInterface.TestDataInterface.ResultCodes = testData.TestInst.StoreValue("screw_1_feed", CDbl(0))
        logger.Info("Store =>" + "screw_1_feed" + " ==" + CStr(rc))
        rc = testData.TestInst.StoreValue("screw_1_f_time", CDbl("2251"))
        logger.Info("Store =>" + "screw_1_f_time" + " ==" + CStr(rc))

        'testData.TestInst.StoreValue("screw_4_f_time", CDbl(0))


    End Sub
#End Region
End Class
