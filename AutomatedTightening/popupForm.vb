Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class popupForm
    Private WithEvents timer As New Timer()
    Friend logger As NLog.Logger = NLog.LogManager.GetLogger("Program")
    'Dim hiosResponse As String = "None"
    Public Property KeyData As String
    Private Sub PopupForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Enter Then
            logger.Info("Response : " & rtb01.Text)
            KeyData = rtb01.Text
            Me.DialogResult = DialogResult.OK
            Me.Close()  ' Close the popup
        Else

        End If
    End Sub

    Private Sub PopupForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True  ' Ensure the form receives all keypress events
        timer.Interval = 3000 'the old is 4000
        AddHandler timer.Tick, AddressOf Timer_Tick
        timer.Start()
        rtb01.Focus()
    End Sub
    Private Sub Timer_Tick(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        KeyData = "Timeout"
        Me.Close()  ' Close the form when time expires
    End Sub


End Class