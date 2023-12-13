
Imports System.Windows.Forms
Imports UdbsInterface.MasterInterface.DatabaseSupport

Public Class SecurityInterface

    ' The EmployeeNumber of the most recently verified password.  This is use for checking
    ' privileges in functions that need it.
    Private _authenticatedEmployeeNumber As String

    ' Modifying the security tables (e.g., to assign an employee to a group) requires the user
    ' to be in the "administrators" group.
    Private Const AdminGroupName = "administrators"

    ' Return the employee number of the most recently verified password.
    Public ReadOnly Property AuthenticatedEmployeeNumber As String
        Get
            Return _authenticatedEmployeeNumber
        End Get
    End Property

    ''' <summary>
    ''' Returns true if the given EmployeeNumber and Password are correct.  Returns false
    ''' on mismatch, or if UDBS cannot be reached.  Updates the receiver with the employee
    ''' number when verification is successful.
    ''' </summary>
    ''' <param name="EmployeeNumber">The employee number to verify.</param>
    ''' <param name="Password">The password for the given employee.</param>
    ''' <returns>A boolean describing whether the employee and password were verified.</returns>
    Public Function VerifyPassword(EmployeeNumber As String, Password As String) As Boolean
        Dim sql As String =
            "SELECT 1 " &
            "  FROM [dbo].[security_employee] " &
           $" WHERE employee_number = '{EmployeeNumber}' " &
           $"   AND employee_password = '{Password}'"

        Dim result = New DataTable
        OpenNetworkRecordSet(result, sql)
        If result Is Nothing OrElse result.Rows Is Nothing Then
            'log
            Return False
        End If

        If result.Rows.Count = 1 Then
            _authenticatedEmployeeNumber = EmployeeNumber
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Check if the given employee is a member of the given group.
    ''' </summary>
    ''' <param name="EmployeeNumber">The employee number to check.</param>
    ''' <param name="GroupName">The name of the group to check.</param>
    ''' <returns>True if the employee is a member of the group and false otherwise.</returns>
    Public Function IsGroupMember(EmployeeNumber As String, GroupName As String) As Boolean
        Dim sql As String =
            "SELECT 1 " &
            "  FROM security_employee " &
            "  JOIN security_membership ON membership_employee_id = employee_id " &
            "  JOIN security_group      ON group_id = membership_group_id " &
           $" WHERE employee_number = '{EmployeeNumber}' " &
           $"   AND group_name = '{GroupName}'"
        Dim result = New DataTable
        OpenNetworkRecordSet(result, sql)
        If result Is Nothing OrElse result.Rows Is Nothing Then
            'log
            Return False
        End If

        Return result.Rows.Count = 1
    End Function

    ''' <summary>
    ''' Find all groups to which the given employee belongs.
    ''' </summary>
    ''' <param name="EmployeeNumber">The employee number to check.</param>
    ''' <param name="Groups">
    ''' (Out) Return parameter holding the full list of groups for the given employee.
    ''' </param>
    ''' <returns>
    ''' True if the list of groups could be found and false otherwise.
    ''' The Groups return parameter
    ''' is only valid when the function returns true.
    ''' </returns>
    Public Function GetGroupMembership(EmployeeNumber As String, ByRef Groups() As String) As Boolean
        Dim sql As String =
            "SELECT group_name AS GroupName " &
            "  FROM security_employee " &
            "  JOIN security_membership ON membership_employee_id = employee_id " &
            "  JOIN security_group      ON group_id = membership_group_id " &
           $" WHERE employee_number = '{EmployeeNumber}' "
        Dim result = New DataTable
        OpenNetworkRecordSet(result, sql)
        If result Is Nothing OrElse result.Rows Is Nothing OrElse result.Rows.Count = 0 Then
            'log
            Return False
        End If

        Array.Resize(Groups, result.Rows.Count)
        For i As Integer = 0 To result.Rows.Count - 1
            Groups(i) = KillNull(result(i)("GroupName"))
        Next

        Return True
    End Function

    ''' <summary>
    ''' Adds a new group to the database.  Does nothing if the group already exists.
    ''' NOTE: This function requires an user in the "administrators" group to have been verified.  See
    '''       VerifyPassword.
    ''' </summary>
    ''' <param name="GroupName">The name of the group to add.</param>
    ''' <param name="ErrorMessage">
    '''     Return parameter with a message describing why the operation failed (if applicable).  The
    '''     most common reason for failure is that an admin user has not been authenticated.
    ''' </param>
    ''' <returns>True if successful and false otherwise.</returns>
    Public Function AddGroup(GroupName As String, ByRef ErrorMessage As String) As Boolean
        If Not IsGroupMember(AuthenticatedEmployeeNumber, AdminGroupName) Then
            ErrorMessage = "You do not have sufficient privileges to add groups."
            Return False
        End If

        ' Insert the name only if it does not already exist.
        Dim sql As String =
            "INSERT INTO security_group(group_name) " &
           $"     SELECT '{GroupName}' " &
           $"      WHERE '{GroupName}' NOT IN (SELECT group_name from security_group )"
        ExecuteNetworkQuery(sql)

        Return True
    End Function

    ''' <summary>
    ''' Modifies an existing group to add the given employees.  Ignores employee numbers that are already
    ''' a member of the group.
    ''' NOTE: This function requires an user in the "administrators" group to have been verified.  See
    '''       VerifyPassword.
    ''' </summary>
    ''' <param name="GroupName">The name of the group to be modified.</param>
    ''' <param name="EmployeeNumbers">The list of employees to add to the group.</param>
    ''' <param name="ErrorMessage">
    '''     Return parameter with a message describing why the operation failed (if applicable).  The
    '''     most common reason for failure is that an admin user has not been authenticated.
    ''' </param>
    ''' <returns>True if successful and false otherwise.</returns>
    Public Function EditGroupAddEmployees(GroupName As String, EmployeeNumbers As String(), ByRef ErrorMessage As String) As Boolean
        If Not IsGroupMember(AuthenticatedEmployeeNumber, AdminGroupName) Then
            ErrorMessage = "You do not have sufficient privileges to edit groups."
            Return False
        End If

        Using transaction = BeginNetworkTransaction()
            For i = 0 To UBound(EmployeeNumbers)
                Dim EmployeeNumber As String = EmployeeNumbers(i)

                ' A query that inserts a record only when there isn't already an membership
                ' record for the given employee and group.
                Dim sql As String =
                    "INSERT INTO security_membership(membership_employee_id, membership_group_id) " &
                    "     SELECT se.employee_id, sg.group_id " &
                    "       FROM security_employee se, " &
                    "            security_group sg " &
                   $"      WHERE se.employee_number = '{EmployeeNumber}' " &
                   $"        AND sg.group_name = '{GroupName}' " &
                    "        AND se.employee_number NOT IN " &
                    "      ( SELECT employee_number " &
                    "          FROM security_employee se " &
                    "          JOIN security_membership sm ON sm.membership_employee_id = se.employee_id " &
                    "          JOIN security_group sg on sg.group_id = sm.membership_group_id " &
                   $"         WHERE se.employee_number = '{EmployeeNumber}' " &
                   $"           AND sg.group_name = '{GroupName}' )"
                ExecuteNetworkQuery(sql, transaction)
            Next
        End Using ' NetworkTransaction

        Return True
    End Function

    ''' <summary>
    ''' Modifies an existing group to remove the given employees.  Ignores employee numbers that are not
    ''' a member of the group.
    ''' NOTE: This function requires an user in the "administrators" group to have been verified.  See
    '''       VerifyPassword.
    ''' </summary>
    ''' <param name="GroupName">The name of the group to modify.</param>
    ''' <param name="EmployeeNumbers">The list of employees to remove from the group.</param>
    ''' <param name="ErrorMessage">
    '''     Return parameter with a message describing why the operation failed (if applicable).  The
    '''     most common reason for failure is that an admin user has not been authenticated.
    ''' </param>
    ''' <returns>True if successful and false otherwise.</returns>
    Public Function EditGroupRemoveEmployees(GroupName As String, EmployeeNumbers As String(), ByRef ErrorMessage As String) As Boolean
        If Not IsGroupMember(AuthenticatedEmployeeNumber, AdminGroupName) Then
            ErrorMessage = "You do not have sufficient privileges to edit groups."
            Return False
        End If

        Using transaction = BeginNetworkTransaction()
            For i = 0 To UBound(EmployeeNumbers)
                Dim EmployeeNumber As String = EmployeeNumbers(i)

                Dim sql As String =
                    "DELETE sm FROM security_membership sm " &
                    "  JOIN security_employee se ON se.employee_id = sm.membership_employee_id " &
                    "  JOIN security_group sg on sg.group_id = sm.membership_group_id " &
                   $" WHERE se.employee_number = '{EmployeeNumber}' " &
                   $"   AND sg.group_name = '{GroupName}'"
                ExecuteNetworkQuery(sql, transaction)
            Next
        End Using ' NetworkTransaction

        Return True
    End Function

    ''' <summary>
    ''' Adds a new employee to the database.
    ''' NOTE: This function requires an user in the "administrators" group to have been verified.  See
    '''       VerifyPassword.
    ''' </summary>
    ''' <param name="EmployeeNumber">Employee number for the new entry.</param>
    ''' <param name="EmployeeName">Name for the new entry.</param>
    ''' <param name="Password">Password to use when verifying the new employee.</param>
    ''' <param name="PasswordHint">A hint that can be displayed if the new employee forgets their password.</param>
    ''' <param name="ErrorMessage">
    '''     Return parameter with a message describing why the operation failed (if applicable).  The
    '''     most common reason for failure is that an admin user has not been authenticated.
    ''' </param>
    ''' <returns>True if successful and false otherwise.</returns>
    Public Function AddEmployee(EmployeeNumber As String, EmployeeName As String, Password As String, PasswordHint As String, ByRef ErrorMessage As String) As Boolean
        If Not IsGroupMember(AuthenticatedEmployeeNumber, AdminGroupName) Then
            ErrorMessage = "You do not have sufficient privileges to delete users."
            Return False
        End If

        Using transaction = BeginNetworkTransaction()
            ' Make sure that this employee number is not already in use.
            Dim sql As String = $"SELECT employee_name as EmployeeName FROM security_employee WHERE employee_number = '{EmployeeNumber}'"
            Dim result = New DataTable
            OpenNetworkRecordSet(result, sql, transaction)
            If result Is Nothing OrElse result.Rows Is Nothing Then
                ErrorMessage = "Could not find employee table in database."
                Return False
            End If

            If result.Rows.Count > 0 Then
                Dim existingEmployeeName As String = KillNull(result(0)("EmployeeName"))

                ErrorMessage = $"Employee number {EmployeeNumber} already used for {existingEmployeeName}."
                Return False
            End If

            sql = "INSERT INTO security_employee( employee_number, employee_password, employee_password_hint,    employee_name ) " &
                 $"     VALUES (                 '{EmployeeNumber}',      '{Password}',       '{PasswordHint}', '{EmployeeName}' )"
            ExecuteNetworkQuery(sql, transaction)
        End Using ' NetworkTransaction

        Return True
    End Function

    ''' <summary>
    ''' Modifies an existing employee record to add membership in the given groups.  Ignores groups to which
    ''' the employee already belongs.
    ''' NOTE: This function requires an user in the "administrators" group to have been verified.  See
    '''       VerifyPassword.
    ''' </summary>
    ''' <param name="EmployeeNumber">The employee number for the record to be modified.</param>
    ''' <param name="GroupNames">The list of groups to which the employee should be added.</param>
    ''' <param name="ErrorMessage">
    '''     Return parameter with a message describing why the operation failed (if applicable).  The
    '''     most common reason for failure is that an admin user has not been authenticated.
    ''' </param>
    ''' <returns>True if successful and false otherwise.</returns>
    Public Function EditEmployeeAddGroups(EmployeeNumber As String, GroupNames As String(), ByRef ErrorMessage As String) As Boolean
        If Not IsGroupMember(AuthenticatedEmployeeNumber, AdminGroupName) Then
            ErrorMessage = "You do not have sufficient privileges to edit users."
            Return False
        End If

        Using transaction = BeginNetworkTransaction()
            For i = 0 To UBound(GroupNames)
                Dim GroupName As String = GroupNames(i)

                ' A query that inserts a record only when there isn't already an membership
                ' record for the given employee and group.
                Dim sql As String =
                    "INSERT INTO security_membership(membership_employee_id, membership_group_id) " &
                    "     SELECT se.employee_id, sg.group_id " &
                    "       FROM security_employee se, " &
                    "            security_group sg " &
                   $"      WHERE se.employee_number = '{EmployeeNumber}' " &
                   $"        AND sg.group_name = '{GroupName}' " &
                    "        AND se.employee_number NOT IN " &
                    "      ( SELECT employee_number " &
                    "          FROM security_employee se " &
                    "          JOIN security_membership sm ON sm.membership_employee_id = se.employee_id " &
                    "          JOIN security_group sg on sg.group_id = sm.membership_group_id " &
                   $"         WHERE se.employee_number = '{EmployeeNumber}' " &
                   $"           AND sg.group_name = '{GroupName}' )"
                ExecuteNetworkQuery(sql, transaction)
            Next
        End Using ' NetworkTransaction

        Return True
    End Function

    ''' <summary>
    ''' Modifies an existing group to remove the given employees.  Ignores employee numbers that are not
    ''' a member of the group.
    ''' NOTE: This function requires an user in the "administrators" group to have been verified.  See
    '''       VerifyPassword.
    ''' </summary>
    ''' <param name="EmployeeNumber">The employee number for the record to be modified.</param>
    ''' <param name="GroupNames">The list of groups from which the employee should be removed.</param>
    ''' <param name="ErrorMessage">
    '''     Return parameter with a message describing why the operation failed (if applicable).  The
    '''     most common reason for failure is that an admin user has not been authenticated.
    ''' </param>
    ''' <returns>True if successful and false otherwise.</returns>
    Public Function EditEmployeeRemoveGroups(EmployeeNumber As String, GroupNames As String(), ByRef ErrorMessage As String) As Boolean
        If Not IsGroupMember(AuthenticatedEmployeeNumber, AdminGroupName) Then
            ErrorMessage = "You do not have sufficient privileges to edit groups."
            Return False
        End If

        Using transaction = BeginNetworkTransaction()
            For i = 0 To UBound(GroupNames)
                Dim GroupName As String = GroupNames(i)

                Dim sql As String =
                    "DELETE sm FROM security_membership sm " &
                    "  JOIN security_employee se ON se.employee_id = sm.membership_employee_id " &
                    "  JOIN security_group sg on sg.group_id = sm.membership_group_id " &
                   $" WHERE se.employee_number = '{EmployeeNumber}' " &
                   $"   AND sg.group_name = '{GroupName}'"
                ExecuteNetworkQuery(sql, transaction)
            Next
        End Using ' NetworkTransaction

        Return True
    End Function

    ''' <summary>
    ''' Deletes an employee, as well as all membership records, from the database.  Does nothing if the
    ''' employee does exist.
    ''' NOTE: This function requires an user in the "administrators" group to have been verified.  See
    '''       VerifyPassword.
    ''' </summary>
    ''' <param name="EmployeeNumber">The employee number of the record to remove.</param>
    ''' <param name="ErrorMessage">
    '''     Return parameter with a message describing why the operation failed (if applicable).  The
    '''     most common reason for failure is that an admin user has not been authenticated.
    ''' </param>
    ''' <returns>True if successful and false otherwise.</returns>
    Public Function DeleteEmployee(EmployeeNumber As String, ByRef ErrorMessage As String) As Boolean
        If Not IsGroupMember(AuthenticatedEmployeeNumber, AdminGroupName) Then
            ErrorMessage = "You do not have sufficient privileges to delete users."
            Return False
        End If

        Using transaction = BeginNetworkTransaction()
            ' Remove all group membership records for the employee that is about to be deleted.  Then
            ' remove the employee.
            Dim sql As String =
                "DELETE FROM security_membership " &
                "      WHERE membership_id in " &
                "    ( SELECT membership_id " &
                "        FROM security_membership sm " &
                "        JOIN security_employee se ON se.employee_id = sm.membership_employee_id " &
               $"       WHERE se.employee_number = '{EmployeeNumber}' ); " &
               $"DELETE FROM security_employee WHERE employee_number = '{EmployeeNumber}';"
            ExecuteNetworkQuery(sql, transaction)
        End Using ' End Transaction

        Return True
    End Function

    Public Function LogIn(UseGUI As Boolean, Optional EmployeeNumber As String = Nothing, Optional Password As String = Nothing) As Boolean

        If Not UseGUI Then
            Return VerifyPassword(EmployeeNumber, Password)
        End If

        Dim form = New LoginForm(EmployeeNumber)
        While True
            Dim result = form.ShowDialog()
            If result <> DialogResult.OK Then
                Return False
            End If

            EmployeeNumber = form.EmployeeNumber
            Password = form.Password
            If VerifyPassword(EmployeeNumber, Password) Then
                Exit While
            End If

            MessageBox.Show("Password Incorrect", "Login", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End While

        _authenticatedEmployeeNumber = EmployeeNumber
        Return True

    End Function
End Class
