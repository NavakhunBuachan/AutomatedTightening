Imports UdbsInterface.MasterInterface

''' <summary>
''' Local DB Integrity Checker.
''' Checks to integrity of the local database to make sure the data in it is consistent.
''' </summary>
Public NotInheritable Class LocalDBIntegrityChecker
    ''' <summary>
    ''' This public property lets unit test mock the DB check code.
    ''' Production code should not need to access this.
    ''' This is achieved by implementing the 'strategy pattern': https://en.wikipedia.org/wiki/Strategy_pattern
    ''' </summary>
    ''' <returns>The integrity check function.</returns>
    Friend Shared Property CheckStrategy As Func(Of String, LocalDBIntegrityStatus) = AddressOf CheckImplementation

    ''' <summary>
    ''' Checks the integrity of the local database to make sure the data is in consistent state.
    ''' </summary>
    ''' <param name="processId">Process Id.</param>
    ''' <returns>Returns a value of <see cref="LocalDBIntegrityStatus"/></returns>
    Public Shared Function Check(processId As String) As LocalDBIntegrityStatus
        Return CheckStrategy.Invoke(processId)
    End Function

    ''' <summary>
    ''' Actual/default implementation of the integrity check. The public method uses the 
    ''' 'check strategy' property. The default value points to this method.
    ''' </summary>
    ''' <param name="processId">Process Id.</param>
    ''' <returns>Returns a value of <see cref="LocalDBIntegrityStatus"/></returns>
    Private Shared Function CheckImplementation(processId As String) As LocalDBIntegrityStatus
        logger.Info($"Checking integrity of the local database for the Process Id: {processId}.")
        Dim status As LocalDBIntegrityStatus = LocalDBIntegrityStatus.Good
        Try
            status = CheckProcessIdExistsInDatabase(processId)
            status = status Or CheckIfTestDataProcessEntryExists(processId)
            status = status Or CheckIfTestDataResultsExists(processId)
            status = status Or CheckIfTestDataResultValueMissing(processId)
            logger.Info($"Checking the local database integrity returned: {status}")
        Catch ex As Exception
            status = LocalDBIntegrityStatus.ErrorCheckingIntegrity
            logger.Error(ex, "Error checking local database integrity.")
        End Try

        Return status
    End Function

    ''' <summary>
    ''' Discards the file of the local database and copies it the destination folder so that a fresh new copy is generated when the connection is reinitialized.
    ''' </summary>
    ''' <param name="destinationFolderPath">Destination folder path.</param>
    ''' <param name="filePrefix">File Prefix.</param>
    Public Shared Sub DiscardLocalDB(destinationFolderPath As String, filePrefix As String)
        Try
            DatabaseSupport.BackupLocalDB(destinationFolderPath, filePrefix, True)
        Catch ex As Exception
            logger.Error(ex, "Error discarding local database.")
        End Try
    End Sub

    ''' <summary>
    ''' Checks whether process id exists in the process_registration table.
    ''' </summary>
    ''' <param name="processId">Process Id.</param>
    ''' <returns>Returns a value of <see cref="LocalDBIntegrityStatus"/></returns>
    Private Shared Function CheckProcessIdExistsInDatabase(processId As String) As LocalDBIntegrityStatus
        Dim status As LocalDBIntegrityStatus = LocalDBIntegrityStatus.ProcessIdDoesNotExist
        Dim rsLocalProcess As New DataTable
        Dim sqlQuery As String = $"SELECT COUNT(*) FROM process_registration WHERE pr_process_id ={processId}"
        If OpenLocalRecordSet(rsLocalProcess, sqlQuery) = ReturnCodes.UDBS_OP_SUCCESS Then
            If rsLocalProcess.Rows.Count > 0 AndAlso Convert.ToInt32(rsLocalProcess.Rows(0)(0)) > 0 Then
                status = LocalDBIntegrityStatus.Good
            End If
        Else
            status = LocalDBIntegrityStatus.ErrorCheckingIntegrity
        End If
        Return status
    End Function

    ''' <summary>
    ''' Checks whether entries exist in the testdata_result table.
    ''' </summary>
    ''' <param name="processId">Process Id.</param>
    ''' <returns>Returns a value of <see cref="LocalDBIntegrityStatus"/></returns>
    Private Shared Function CheckIfTestDataResultsExists(processId As String) As LocalDBIntegrityStatus
        Dim status As LocalDBIntegrityStatus = LocalDBIntegrityStatus.TestDataResultNotPresent
        Dim rsLocalProcess As New DataTable
        Dim sqlQuery As String = $"SELECT COUNT(*) FROM testdata_result WHERE result_process_id={processId}"
        If OpenLocalRecordSet(rsLocalProcess, sqlQuery) = ReturnCodes.UDBS_OP_SUCCESS Then
            If rsLocalProcess.Rows.Count > 0 AndAlso Convert.ToInt32(rsLocalProcess.Rows(0)(0)) > 0 Then
                status = LocalDBIntegrityStatus.Good
            End If
        Else
            status = LocalDBIntegrityStatus.ErrorCheckingIntegrity
        End If
        Return status
    End Function

    ''' <summary>
    ''' Checks whether there are rows containing values either in columns result_value, result_stringdata or result_passflag to indicate that there are is result data in the local database.
    ''' </summary>
    ''' <param name="processId">Process Id.</param>
    ''' <returns>Returns a value of <see cref="LocalDBIntegrityStatus"/></returns>
    Private Shared Function CheckIfTestDataResultValueMissing(processId As String) As LocalDBIntegrityStatus
        Dim status As LocalDBIntegrityStatus = LocalDBIntegrityStatus.TestDataResultValueMissing
        Dim rsLocalProcess As New DataTable
        Dim sqlQuery As String = $"SELECT COUNT(*) FROM testdata_result WHERE result_process_id = {processId} AND ( coalesce(result_value,'') != '' OR coalesce(result_stringdata,'') != '' OR coalesce(result_passflag,'') != '' )"

        If OpenLocalRecordSet(rsLocalProcess, sqlQuery) = ReturnCodes.UDBS_OP_SUCCESS Then
            If rsLocalProcess.Rows.Count > 0 AndAlso Convert.ToInt32(rsLocalProcess.Rows(0)(0)) > 0 Then
                status = LocalDBIntegrityStatus.Good
            End If
        Else
            status = LocalDBIntegrityStatus.ErrorCheckingIntegrity
        End If
        Return status
    End Function

    ''' <summary>
    ''' Checks whether there is an entry for the process id in the table testdata_process
    ''' </summary>
    ''' <param name="processId">Process Id.</param>
    ''' <returns>Returns a value of <see cref="LocalDBIntegrityStatus"/></returns>
    Private Shared Function CheckIfTestDataProcessEntryExists(processId As String) As LocalDBIntegrityStatus
        Dim status As LocalDBIntegrityStatus = LocalDBIntegrityStatus.TestDataProcessDoesNotExist
        Dim rsLocalProcess As New DataTable
        Dim sqlQuery As String = $"SELECT COUNT(*) FROM testdata_process WHERE process_id={processId}"
        If OpenLocalRecordSet(rsLocalProcess, sqlQuery) = ReturnCodes.UDBS_OP_SUCCESS Then
            If rsLocalProcess.Rows.Count > 0 AndAlso Convert.ToInt32(rsLocalProcess.Rows(0)(0)) > 0 Then
                status = LocalDBIntegrityStatus.Good
            End If
        Else
            status = LocalDBIntegrityStatus.ErrorCheckingIntegrity
        End If
        Return status
    End Function

End Class
