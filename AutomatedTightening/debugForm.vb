Imports NLog

Public Class debugForm

    Public screw As smartScrew
    Friend logger As NLog.Logger = NLog.LogManager.GetLogger("Program")
    Public returnPacketScrew, returnPacketDummy, picPassed, picFailed As String()

    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        Dim response As String
        'Dim returnPacket As String()
        screw.getPresetNumber()
        logger.Info("Send :" & tbSend.Text)
        response = screw.sendAndRecvAscii(tbSend.Text)
        'response = screw.sendAndRecv(tbSend.Text)
        logger.Info("Recv :" & response)

        rtbRecv.Text = response

    End Sub

    Private Sub debugForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        logger.Info("Initial smartScrew")

        screw = New smartScrew
        screw.portName = My.Settings.SmartScrewPort
        screw.baudRate = My.Settings.SmartScrewRate
        returnPacketDummy = {"None", "None", "None", "None", "None", "None", "None", "None", "None", "None", "None", "None"}
        Console.WriteLine("waiting.....")

        Do
            Dim keyInfo As ConsoleKeyInfo = Console.ReadKey(True) ' True to not display the pressed key
            If keyInfo.Key = ConsoleKey.Enter Then
                Exit Do ' Exit the loop if Enter is pressed
            Else
                Console.WriteLine("You pressed: " + keyInfo.KeyChar)
            End If
        Loop

        Console.WriteLine("Enter key pressed. Loop has been exited.")
    End Sub
    Private Sub btnInit_Click(sender As Object, e As EventArgs) Handles btnInit.Click
        screw.connect()
    End Sub


End Class