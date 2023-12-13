Imports System.Windows.Forms
Imports UdbsInterface.MasterInterface.DatabaseSupport

''' <summary>
''' Reproduces behaviour from the LoginForm in the old UDBS_Security.dll.
''' </summary>
Friend Class LoginForm
    Private _employeeNumber As String
    Private _password As String

    Public Sub New(employeeNumber As String)
        ' This call is required by the designer.
        InitializeComponent()

        _employeeNumber = employeeNumber
    End Sub

    Public ReadOnly Property EmployeeNumber As String
        Get
            Return _employeeNumber
        End Get
    End Property

    Public ReadOnly Property Password As String
        Get
            Return _password
        End Get
    End Property

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        _employeeNumber = UsernameTextBox.Text
        _password = PasswordTextBox.Text
        DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Hint_Click(sender As Object, e As EventArgs) Handles Hint.Click
        Dim username As String = UsernameTextBox.Text

        ' Lookup the hint for the EmployeeNumber in username field.
        Dim sql As String =
            "SELECT employee_password_hint as Hint " &
            "  FROM security_employee " &
           $" WHERE employee_number = '{username}'"
        Dim result = New DataTable
        OpenNetworkRecordSet(result, sql)

        ' Display either the password hint or an error message.
        If result Is Nothing OrElse result.Rows Is Nothing OrElse result.Rows.Count <= 0 Then
            MessageBox.Show($"{username} is not a recognized user.",
                             "UDBS Security", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UsernameTextBox.BringToFront()
        Else
            Dim hint As String = KillNull(result(0)("Hint"))
            MessageBox.Show($"The password hint for employee {username} is:{vbCrLf}{vbCrLf}{hint}",
                             "UDBS Security", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' I don't know why this is needed, but the dialog does not show up without it.
        Visible = True
    End Sub
End Class
