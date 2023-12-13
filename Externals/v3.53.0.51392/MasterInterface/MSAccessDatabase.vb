Option Explicit On
Option Compare Text
Option Infer On
Option Strict On

Imports System.Data.Common
Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading
Imports UdbsInterface.MasterInterface

''' <summary>
'''     MS-Access
''' </summary>
Public Class MSAccessDatabase
    Implements IDatabase
    Private _connection As OleDbConnection
    Private Const OptimumChunkSize As Integer = 8040 ' see:https://docs.microsoft.com/en-us/sql/t-sql/queries/update-transact-sql?redirectedfrom=MSDN&view=sql-server-ver15
    Private Const BlobColumnName As String = "blob_blob"
    Public Sub New(UDBSType As UDBS_DBType)
        DBType = UDBSType
    End Sub

    Public Property CommandTimeOut As Integer Implements IDatabase.CommandTimeOut

    Private _connectionString As String

    Private Sub EnsureOpen()
        If IsNothing(_connection) Then
            _connection = New OleDbConnection(_connectionString)
        End If

        If (_connection?.State <> ConnectionState.Open) Then
            _connection?.Close()
            _connection = New OleDbConnection(_connectionString)
            _connection.Open()
        End If
    End Sub

    ''' <summary>
    '''     OleDB Syntax as stored in the registry
    ''' </summary>
    ''' <returns></returns>
    Public Property ConnectionString As String Implements IDatabase.ConnectionString
        Get
            Return _connectionString
        End Get
        Set
            If Not Value.Equals(_connectionString, StringComparison.InvariantCultureIgnoreCase) Then
                _connection?.Close()
                _connection?.Dispose()
                _connectionString = Value
                _connection = New OleDbConnection(_connectionString)
                _connection.Open()
            End If
        End Set
    End Property


    Public Function ExecuteData(sqlCommand As String, ByRef blob As Stream) As DataRow Implements IDatabase.ExecuteData

        Dim drResult As DataRow = Nothing
        ' The bytes returned from GetBytes.  
        Dim bytesRead As Long
        ' The starting position in the BLOB output.  
        Dim startIndex As Long = 0
        ' The BLOB byte() buffer to be filled by GetBytes.  
        Dim outByte(OptimumChunkSize - 1) As Byte

        EnsureOpen()
        Using scope = New SingletonReleaser(Of OleDbConnection)(_connection)
            Using command As New OleDbCommand(sqlCommand, _connection, CType(scope.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}
                Dim dt = New DataTable()
                Using reader = command.ExecuteReader(CommandBehavior.SequentialAccess)


                    Dim columnData As List(Of Tuple(Of String, Type))
                    Using schemaTable = reader.GetSchemaTable()
                        columnData = schemaTable.Rows.Cast(Of DataRow).
                            Select(
                                Function(dr) _
                                      Tuple.Create(dr.Field(Of String)("ColumnName"),
                                                   dr.Field(Of Type)("DataType")
                                                   )).ToList()
                    End Using
                    For Each o As Tuple(Of String, Type) In columnData
                        If o.Item1.Equals(BlobColumnName, StringComparison.InvariantCultureIgnoreCase) Then
                            Continue For
                        End If
                        dt.Columns.Add(o.Item1, o.Item2)
                    Next

                    If reader.Read() Then
                        drResult = dt.NewRow()
                        Dim i As Integer = 0
                        For Each o As Tuple(Of String, Type) In columnData

                            If o.Item1.Equals(BlobColumnName, StringComparison.InvariantCultureIgnoreCase) Then
                                'stream
                                ' Read bytes into outByte() and retain the number of bytes returned.  
                                Do
                                    bytesRead = reader.GetBytes(i, startIndex, outByte, 0, OptimumChunkSize)
                                    If bytesRead <= 0 Then Exit Do
                                    blob.Write(outByte, 0, Convert.ToInt32(bytesRead))
                                    blob.Flush()
                                    ' Reposition start index to end of the last buffer and fill buffer.  
                                    startIndex += bytesRead
                                Loop

                                Continue For
                            End If
                            drResult(o.Item1) = reader(o.Item1)
                            i += 1
                        Next

                    End If
                End Using
            End Using
        End Using
        Return drResult
    End Function
    Public ReadOnly Property DBType As UDBS_DBType Implements IDatabase.DBType


    Public Function Execute(sqlCommand As String) As Integer Implements IDatabase.Execute
        EnsureOpen()
        Using scope = New SingletonReleaser(Of OleDbConnection)(_connection)
            Using command As New OleDbCommand(sqlCommand, _connection, CType(scope.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}
                Return command.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Function ExecuteScalar(sqlCommand As String) As Object Implements IDatabase.ExecuteScalar
        EnsureOpen()
        Using scope = New SingletonReleaser(Of OleDbConnection)(_connection)
            Using command As New OleDbCommand(sqlCommand, _connection, CType(scope.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}
                Return command.ExecuteScalar()
            End Using
        End Using
    End Function

    Public Function ExecuteData(sqlCommand As String) As DataTable Implements IDatabase.ExecuteData
        EnsureOpen()
        Using scope = New SingletonReleaser(Of OleDbConnection)(_connection)
            Using command As New OleDbCommand(sqlCommand, _connection, CType(scope.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}
                Using adapter = New OleDbDataAdapter(command)
                    Dim dt = New DataTable()
                    adapter.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    Public Function ConnectionStringWithoutPassword() As String Implements IDatabase.ConnectionStringWithoutPassword
        Try
            Dim tempS As New OleDbConnectionStringBuilder(ConnectionString)
            tempS.Remove("Password")
            Return tempS.ToString()
        Catch ex As Exception
            Return ConnectionString
        End Try
    End Function

    Public ReadOnly Property SystemAvailable As Boolean Implements IDataSystem.SystemAvailable
        Get
            Try
                logger.Trace($"Connecting To: {ConnectionString}")
                ExecuteScalar("select 1")
                Return True
            Catch ex As Exception
                logger.Error(ex, $"Failure connecting to: {ConnectionString}")
                Return False
            End Try
        End Get
    End Property

    Public Function GetColumnTypes(sqlCommand As String) As List(Of Tuple(Of String, String, Integer)) _
        Implements IDatabase.GetColumnTypes
        EnsureOpen()
        Using command As New OleDbCommand(sqlCommand, _connection) With {.CommandTimeout = CommandTimeOut}
            Using reader = command.ExecuteReader()
                Using schemaTable = reader.GetSchemaTable()
                    Return schemaTable.Rows.Cast(Of DataRow).
                        Select(
                            Function(dr) _
                                  Tuple.Create(dr.Field(Of String)("ColumnName"),
                                               dr.Field(Of Type)("DataType").ToString(),
                                               dr.Field(Of Integer)("ColumnSize"))).ToList()
                End Using
            End Using
        End Using
    End Function

    Public Function BeginTransaction() As ITransactionScope Implements IDatabase.BeginTransaction
        EnsureOpen()

        Return New SingletonReleaser(Of OleDbConnection)(_connection)
    End Function

    Public Function ExecuteData(sqlCommand As String, scope As ITransactionScope) As DataTable _
        Implements IDatabase.ExecuteData
        Dim trx = CType(scope, SingletonReleaser(Of OleDbConnection))
        Dim conn As OleDbConnection = CType(trx.Transaction.Connection, OleDbConnection)
        EnsureOpen()
        Using command As New OleDbCommand(sqlCommand, conn, CType(trx.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}
            Using adapter = New OleDbDataAdapter(command) With {.AcceptChangesDuringFill = False}
                Dim dt = New DataTable()
                adapter.Fill(dt)
                Return dt
            End Using
        End Using
    End Function

    Public Function CreateTableAdapter(selectSQL As String, scope As ITransactionScope, ByRef workTable As DataTable) _
        As IAdapterSession Implements IDatabase.CreateTableAdapter
        Dim trx = CType(scope, SingletonReleaser(Of OleDbConnection))
        Dim conn As OleDbConnection = CType(trx.Transaction.Connection, OleDbConnection)
        Dim command As New OleDbCommand(selectSQL, conn, CType(trx.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}
        Dim adapter As New OleDbDataAdapter(command)
        Dim builder As New OleDbCommandBuilder(adapter)
        builder.GetDeleteCommand()
        builder.GetInsertCommand()
        builder.GetUpdateCommand()
        adapter.Fill(workTable)

        Return New AdapterSession(command, adapter, builder)
    End Function

    Private Structure AdapterSession
        Implements IAdapterSession

        Private ReadOnly _command As OleDbCommand
        Private ReadOnly _adapter As OleDbDataAdapter
        Private ReadOnly _builder As OleDbCommandBuilder

        Friend Sub New(command As OleDbCommand, adapter As OleDbDataAdapter, builder As OleDbCommandBuilder)
            _command = command
            _adapter = adapter
            _builder = builder
        End Sub

        Public ReadOnly Property Adapter As DbDataAdapter Implements IAdapterSession.Adapter
            Get
                Return _adapter
            End Get
        End Property

        Public ReadOnly Property Builder As DbCommandBuilder Implements IAdapterSession.Builder
            Get
                Return _builder
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose
            _builder?.Dispose()
            _adapter?.Dispose()
            _command?.Dispose()
        End Sub
    End Structure

    Public Function Execute(sqlCommand As String, transactionScope As ITransactionScope) As Integer _
        Implements IDatabase.Execute
        Dim trx = CType(transactionScope, SingletonReleaser(Of OleDbConnection))
        Dim connection As OleDbConnection = CType(trx.Transaction.Connection, OleDbConnection)
        EnsureOpen()
        Using _
            command As _
                New OleDbCommand(sqlCommand, connection, CType(trx.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}
            Return command.ExecuteNonQuery()
        End Using
    End Function

    Public Function InsertRecord(columnNames() As String, columnValues() As Object, tableName As String,
                                 scope As ITransactionScope, Optional primaryKey As String = "") As Integer _
        Implements IDatabase.InsertRecord
        ' build insert query
        Dim returnKey As Boolean = Not String.IsNullOrEmpty(primaryKey)

        Dim input As Stream = Nothing
        Dim isBlobTable As Boolean = tableName.IndexOf("_blob", StringComparison.InvariantCultureIgnoreCase) > 0
        Dim blobIdx As Integer = -1
        If isBlobTable Then
            blobIdx = Array.IndexOf(columnNames, BlobColumnName)
            input = CType(columnValues(blobIdx), Stream)
        End If
        If columnNames.Length <> columnValues.Length Then
            Throw _
                New ArgumentException($"{NameOf(columnNames)} and {NameOf(columnValues)} do not have the same size!")
        End If
        Dim parms = columnNames.Select(Function(c, i) $"@p{i}").ToArray()
        Dim sb As New StringBuilder($"INSERT INTO {tableName}(")
        sb.Append($"{String.Join(",", columnNames)}) ")

        sb.Append($"VALUES({String.Join(",", parms)})")

        Dim trx = CType(scope, SingletonReleaser(Of OleDbConnection))
        Dim conn As OleDbConnection = CType(trx.Transaction.Connection, OleDbConnection)
        EnsureOpen()
        Using _
            command As _
                New OleDbCommand(sb.ToString(), conn, CType(trx.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}

            For i = 0 To columnValues.Length - 1
                ' detect function calls
                If columnValues(i)?.ToString().EndsWith("()") Then
                    Dim payload = ExecuteScalar($"select {columnValues(i)}")
                    command.Parameters.AddWithValue(parms(i), payload)
                Else
                    Dim dummy As Date
                    ' https://support.microsoft.com/en-us/help/320435/info-oledbtype-enumeration-vs-microsoft-access-data-types
                    If TypeOf (columnValues(i)) Is DateTime OrElse Date.TryParseExact(columnValues(i)?.ToString(), DBDateFormatting, enUS, DateTimeStyles.None, dummy) Then
                        command.Parameters.Add(parms(i), OleDbType.Date).Value = columnValues(i)
                    ElseIf isBlobTable AndAlso blobIdx = i Then
                        command.Parameters.Add(parms(i), OleDbType.Binary, -1).Value = StreamToBytes(input) ' May cause OOM
                    Else
                        command.Parameters.AddWithValue(parms(i), columnValues(i))
                    End If
                End If

            Next
            If returnKey Then
                command.ExecuteNonQuery()
                Dim rowId As Integer
                If returnKey Then
                    Using New OleDbCommand("SELECT @@IDENTITY", _connection, CType(trx.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}
                        rowId = Convert.ToInt32(command.ExecuteScalar())
                    End Using
                End If
                Return (rowId)
            Else
                Return command.ExecuteNonQuery()
            End If
        End Using
    End Function


    Public Function UpdateRecord(keys() As String, columnNames() As String, columnValues() As Object,
                                 tableName As String, scope As ITransactionScope) As Boolean _
        Implements IDatabase.UpdateRecord
        ' build update query
        If columnNames.Length <> columnValues.Length Then
            Throw _
                New ArgumentException($"{NameOf(columnNames)} and {NameOf(columnValues)} do not have the same size!")
        End If

        Dim input As Stream = Nothing
        Dim isBlobTable As Boolean = tableName.IndexOf("_blob", StringComparison.InvariantCultureIgnoreCase) > 0
        Dim blobIdx As Integer = -1

        Dim pairs = columnNames.Zip(columnValues, Function(name, valu) New With {name, .Value = valu}).ToList()
        Dim toUpdate = pairs.Where(Function(kv) Not keys.Contains(kv.name)).ToList()

        If isBlobTable Then
            blobIdx = Array.IndexOf(toUpdate.Select(Function(z) z.name).ToArray(), BlobColumnName)
            input = CType(toUpdate(blobIdx).Value, Stream)
        End If

        Dim wheres = pairs.Where(Function(kv) keys.Contains(kv.name)).ToList()

        Dim parms = toUpdate.Select(Function(c, i) $"{c.name}=@p{i}").ToArray()
        Dim constraints = wheres.Select(Function(c, i) $"{c.name}=@w{i}").ToArray()

        Dim sb As New StringBuilder($"UPDATE {tableName} Set ")
        sb.Append($"{String.Join(",", parms)} WHERE ")
        sb.Append($"{String.Join(" AND ", constraints)} ")

        Dim trx = CType(scope, SingletonReleaser(Of OleDbConnection))
        Dim conn As OleDbConnection = CType(trx.Transaction.Connection, OleDbConnection)
        EnsureOpen()
        Using _
            command As _
                New OleDbCommand(sb.ToString(), conn, CType(trx.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}

            For Each parUpdate In toUpdate.Select(Function(c, i) New With {.Data = c, .Idx = i})

                ' detect function calls
                If parUpdate.Data.Value?.ToString().EndsWith("()") Then
                    Dim payload = ExecuteScalar($"select {parUpdate.Data.Value}")
                    command.Parameters.AddWithValue($"@p{parUpdate.Idx}", payload)
                ElseIf isBlobTable AndAlso blobIdx = parUpdate.Idx Then
                    command.Parameters.Add($"@p{parUpdate.Idx}", OleDbType.Binary, -1).Value = StreamToBytes(input) ' May Cause OOM
                Else

                    ' https://support.microsoft.com/en-us/help/320435/info-oledbtype-enumeration-vs-microsoft-access-data-types
                    Dim dummy As Date
                    If TypeOf (parUpdate.Data.Value) Is DateTime OrElse Date.TryParseExact(parUpdate.Data.Value?.ToString(), DBDateFormatting, enUS, DateTimeStyles.None, dummy) Then
                        command.Parameters.Add($"@p{parUpdate.Idx}", OleDbType.Date).Value = parUpdate.Data.Value
                    Else
                        command.Parameters.AddWithValue($"@p{parUpdate.Idx}", parUpdate.Data.Value)
                    End If

                End If
            Next
            For Each constrain In wheres.Select(Function(c, i) New With {.Data = c, .Idx = i})

                ' detect function calls
                If constrain.Data.Value?.ToString().EndsWith("()") Then
                    Dim payload = ExecuteScalar($"select {constrain.Data.Value}")
                    command.Parameters.AddWithValue($"@w{constrain.Idx}", payload)
                Else
                    command.Parameters.AddWithValue($"@w{constrain.Idx}", constrain.Data.Value)
                End If
            Next

            Return (command.ExecuteNonQuery() >= 1) ' triggers may also count
        End Using
    End Function


    Private Shared Function StreamToBytes(ByVal input As Stream) As Byte()
        Dim buffer As Byte() = New Byte(16 * 1024) {}

        Using ms As MemoryStream = New MemoryStream()
            Dim read As Integer = input.Read(buffer, 0, buffer.Length)

            While read > 0
                ms.Write(buffer, 0, read)
                read = input.Read(buffer, 0, buffer.Length)
            End While

            Return ms.ToArray()
        End Using
    End Function

    Public Function InsertRecord(columnNames() As String, columnValues() As Object,
                                 tableName As String, Optional ByVal primaryKey As String = "") As Integer _
        Implements IDatabase.InsertRecord
        ' build insert query
        Dim returnKey As Boolean = Not String.IsNullOrEmpty(primaryKey)

        Dim input As Stream = Nothing
        Dim isBlobTable As Boolean = tableName.IndexOf("_blob", StringComparison.InvariantCultureIgnoreCase) > 0
        Dim blobIdx As Integer = -1
        If isBlobTable Then
            blobIdx = Array.IndexOf(columnNames, BlobColumnName)
            input = CType(columnValues(blobIdx), Stream)
            primaryKey = "blob_id"
        End If


        If columnNames.Length <> columnValues.Length Then
            Throw _
                New ArgumentException($"{NameOf(columnNames)} and {NameOf(columnValues)} do not have the same size!")
        End If
        Dim parms = columnNames.Select(Function(c, i) $"@p{i}").ToArray()
        Dim sb As New StringBuilder($"INSERT INTO {tableName}(")
        sb.Append($"{String.Join(",", columnNames)}) ")

        sb.Append($"VALUES({String.Join(",", parms)})")

        EnsureOpen()
        Using scope = New SingletonReleaser(Of OleDbConnection)(_connection)
            Try
                Using command As New OleDbCommand(sb.ToString(), _connection, CType(scope.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}

                    For i = 0 To columnValues.Length - 1
                        ' detect function calls
                        If columnValues(i)?.ToString().EndsWith("()") Then
                            Dim payload = ExecuteScalar($"select {columnValues(i)}")
                            command.Parameters.AddWithValue(parms(i), payload)
                        Else
                            ' https://support.microsoft.com/en-us/help/320435/info-oledbtype-enumeration-vs-microsoft-access-data-types
                            Dim dummy As Date
                            If TypeOf (columnValues(i)) Is DateTime OrElse Date.TryParseExact(columnValues(i)?.ToString(), DBDateFormatting, enUS, DateTimeStyles.None, dummy) Then
                                command.Parameters.Add(parms(i), OleDbType.Date).Value = columnValues(i)
                            ElseIf isBlobTable AndAlso blobIdx = i Then
                                command.Parameters.Add(parms(i), OleDbType.Binary, -1).Value = StreamToBytes(input) ' May Cause OOM
                            Else
                                command.Parameters.AddWithValue(parms(i), columnValues(i))
                            End If

                        End If

                    Next

                    If returnKey Then
                        command.ExecuteNonQuery()
                        Dim rowId As Integer

                        Using New OleDbCommand("SELECT @@IDENTITY", _connection, CType(scope.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}
                            rowId = Convert.ToInt32(command.ExecuteScalar())
                        End Using


                        Return (rowId)

                    Else
                        'NB: Keep this code sample for debugging INSERT issues
                        'If tableName = "testdata_result" Then
                        '    Dim cmdTxt = ActualCommandTextByNames(command)
                        '    If cmdTxt.IndexOf("14687874") >= 0 Then
                        '        Debug.WriteLine(cmdTxt)
                        '        If cmdTxt.IndexOf("result_passflag") >= 0 Then
                        '            Debugger.Break()
                        '        End If
                        '    End If
                        'End If

                        Return command.ExecuteNonQuery()
                    End If
                End Using

            Catch ex As Exception
                scope.HasError = True

                Throw New ApplicationException($"Error in Inserting to {tableName}", ex)

            End Try
        End Using

    End Function

    Public Function UpdateRecord(keys() As String, columnNames() As String, columnValues() As Object,
                                 tableName As String) As Boolean Implements IDatabase.UpdateRecord
        ' build update query
        If columnNames.Length <> columnValues.Length Then
            Throw _
                New ArgumentException($"{NameOf(columnNames)} and {NameOf(columnValues)} do not have the same size!")
        End If

        Dim input As Stream = Nothing
        Dim isBlobTable As Boolean = tableName.IndexOf("_blob", StringComparison.InvariantCultureIgnoreCase) > 0
        Dim blobIdx As Integer = -1

        Dim pairs = columnNames.Zip(columnValues, Function(name, valu) New With {name, .Value = valu}).ToList()
        Dim toUpdate = pairs.Where(Function(kv) Not keys.Contains(kv.name)).ToList()

        If isBlobTable Then
            blobIdx = Array.IndexOf(toUpdate.Select(Function(z) z.name).ToArray(), BlobColumnName)
            input = CType(toUpdate(blobIdx).Value, Stream)
        End If

        Dim wheres = pairs.Where(Function(kv) keys.Contains(kv.name)).ToList()

        Dim parms = toUpdate.Select(Function(c, i) $"{c.name}=@p{i}").ToArray()
        Dim constraints = wheres.Select(Function(c, i) $"{c.name}=@w{i}").ToArray()

        Dim sb As New StringBuilder($"UPDATE {tableName} Set ")
        sb.Append($"{String.Join(",", parms)} WHERE ")
        sb.Append($"{String.Join(" AND ", constraints)} ")

        EnsureOpen()
        Using scope = New SingletonReleaser(Of OleDbConnection)(_connection)
            Try
                Using command As New OleDbCommand(sb.ToString(), _connection, CType(scope.Transaction, OleDbTransaction)) With {.CommandTimeout = CommandTimeOut}

                    For Each parUpdate In toUpdate.Select(Function(c, i) New With {.Data = c, .Idx = i})

                        ' detect function calls
                        If parUpdate.Data.Value?.ToString().EndsWith("()") Then
                            Dim payload = ExecuteScalar($"select {parUpdate.Data.Value}")
                            command.Parameters.AddWithValue($"@p{parUpdate.Idx}", payload)
                        ElseIf isBlobTable AndAlso blobIdx = parUpdate.Idx Then
                            command.Parameters.Add($"@p{parUpdate.Idx}", OleDbType.Binary, -1).Value = StreamToBytes(input) ' May Cause OOM
                        Else
                            ' https://support.microsoft.com/en-us/help/320435/info-oledbtype-enumeration-vs-microsoft-access-data-types
                            Dim dummy As Date
                            If TypeOf (parUpdate.Data.Value) Is DateTime OrElse Date.TryParseExact(parUpdate.Data.Value?.ToString(), DBDateFormatting, enUS, DateTimeStyles.None, dummy) Then

                                command.Parameters.Add($"@p{parUpdate.Idx}", OleDbType.Date).Value = parUpdate.Data.Value

                            Else
                                command.Parameters.AddWithValue($"@p{parUpdate.Idx}", parUpdate.Data.Value)
                            End If
                        End If
                    Next
                    For Each constrain In wheres.Select(Function(c, i) New With {.Data = c, .Idx = i})

                        ' detect function calls
                        If constrain.Data.Value?.ToString().EndsWith("()") Then
                            Dim payload = ExecuteScalar($"select {constrain.Data.Value}")
                            command.Parameters.AddWithValue($"@w{constrain.Idx}", payload)
                        Else
                            ' https://support.microsoft.com/en-us/help/320435/info-oledbtype-enumeration-vs-microsoft-access-data-types
                            If TypeOf (constrain.Data.Value) Is DateTime Then
                                command.Parameters.AddWithValue($"@w{constrain.Idx}",
                                                                GetDateWithoutMilliseconds(
                                                                    Convert.ToDateTime(constrain.Data.Value)))
                            Else
                                command.Parameters.AddWithValue($"@w{constrain.Idx}", constrain.Data.Value)
                            End If
                        End If
                    Next


                    Return (command.ExecuteNonQuery() >= 1) ' triggers may also count


                End Using
            Catch ex As Exception
                scope.HasError = True
                Throw New ApplicationException($"Error in Updating to {tableName}", ex)
            End Try
        End Using
    End Function

    Public Function DeleteRecord(constraintKeys() As String, constraintValues() As Object, tableName As String) _
        As Boolean Implements IDatabase.DeleteRecord
        ' build delete query
        If constraintKeys.Length <> constraintValues.Length Then
            Throw _
                New ArgumentException(
                    $"{NameOf(constraintKeys)} and {NameOf(constraintValues)} do not have the same size!")
        End If
        Dim wheres =
                constraintKeys.Zip(constraintValues, Function(name, valu) New With {name, .Value = valu}).ToList()
        Dim constraints = wheres.Select(Function(c, i) $"{c.name}=@w{i}").ToArray()

        Dim sb As New StringBuilder($"DELETE FROM {tableName} WHERE ")
        sb.Append($"{String.Join(" AND ", constraints)} ")

        EnsureOpen()
        Using command As New OleDbCommand(sb.ToString(), _connection) With {.CommandTimeout = CommandTimeOut}

            For Each constrain In wheres.Select(Function(c, i) New With {.Data = c, .Idx = i})
                ' detect function calls
                If constrain.Data.Value?.ToString().EndsWith("()") Then
                    Dim payload = ExecuteScalar($"select {constrain.Data.Value}")
                    command.Parameters.AddWithValue($"@w{constrain.Idx}", payload)
                Else
                    ' https://support.microsoft.com/en-us/help/320435/info-oledbtype-enumeration-vs-microsoft-access-data-types
                    If TypeOf (constrain.Data.Value) Is DateTime Then
                        command.Parameters.AddWithValue($"@w{constrain.Idx}",
                                                        GetDateWithoutMilliseconds(
                                                            Convert.ToDateTime(constrain.Data.Value)))
                    Else
                        command.Parameters.AddWithValue($"@w{constrain.Idx}", constrain.Data.Value)
                    End If
                End If
            Next
            Return (command.ExecuteNonQuery() >= 1) ' triggers may also count
        End Using
    End Function



    Private Shared Function GetDateWithoutMilliseconds(d As DateTime) As DateTime
        Return New DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, d.Second)
    End Function

#Region "IDisposable"

    Private _disposed As Boolean = False

    Public Sub Dispose() _
        Implements IDisposable.Dispose
        ' Dispose of unmanaged resources.
        Dispose(True)
        ' Suppress finalization.
        GC.SuppressFinalize(Me)
    End Sub
    ' Protected implementation of Dispose pattern.
    Protected Overridable Sub Dispose(disposing As Boolean)
        If _disposed Then Return

        If disposing Then
            ' Free any other managed objects here.
            _connection?.Close()
            _connection.Dispose()
        End If

        ' Free any unmanaged objects here.
        '
        _disposed = True
    End Sub

#End Region
End Class