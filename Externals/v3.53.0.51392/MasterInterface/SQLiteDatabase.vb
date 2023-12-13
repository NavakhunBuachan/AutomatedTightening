Option Explicit On
Option Compare Text
Option Infer On
Option Strict On

Imports System.Data.Common
Imports System.Data.SQLite
Imports System.IO
Imports System.Text
Imports System.Threading
Imports UdbsInterface.MasterInterface

''' <summary>
'''     Local Database implementation using SQLite
''' </summary>
Friend Class SQLiteDatabase
    Implements IDatabase
    Private Const BlobColumnName As String = "blob_blob"
    Private Const OptimumChunkSize As Integer = 8040 ' see:https://docs.microsoft.com/en-us/sql/t-sql/queries/update-transact-sql?redirectedfrom=MSDN&view=sql-server-ver15
    Private _connection As SQLiteConnection

    Public Sub New()
        DBType = UDBS_DBType.LocalDB
    End Sub

    Public ReadOnly Property DBType As UDBS_DBType Implements IDatabase.DBType
    Public Property CommandTimeOut As Integer Implements IDatabase.CommandTimeOut
    Private _connectionString As String

    Public Property ConnectionString As String Implements IDatabase.ConnectionString
        Get
            Return _connectionString
        End Get
        Set
            If Not Value.Equals(_connectionString, StringComparison.InvariantCultureIgnoreCase) Then
                _connection?.Close()
                _connection?.Dispose()
                _connectionString = Value
                _connection = New SQLiteConnection(_connectionString)
                _connection.Open()
            End If
        End Set
    End Property

    Private Sub EnsureOpen()
        If IsNothing(_connection) Then
            _connection = New SQLiteConnection(_connectionString)
        End If

        If (_connection?.State <> ConnectionState.Open) Then
            _connection?.Close()
            _connection = New SQLiteConnection(_connectionString)
            _connection.Open()
        End If
    End Sub

    Public ReadOnly Property SystemAvailable As Boolean Implements IDatabase.SystemAvailable
        Get
            Try
                ExecuteScalar("select 1")
                Return True
            Catch ex As Exception
                logger.Error(ex, $"Failure connecting to: {ConnectionStringWithoutPassword()}")
                Return False
            End Try
        End Get
    End Property

    Public ReadOnly Property SQLClientConnectionString As String Implements IDatabase.SQLClientConnectionString
        Get
            Return ConnectionString
        End Get
    End Property

    Public Function ConnectionStringWithoutPassword() As String Implements IDatabase.ConnectionStringWithoutPassword
        ' The SQLite connection string does not contain the password.
        ' Return it as-is.
        Return ConnectionString
    End Function

    Public Function ExecuteData(sqlCommand As String, ByRef blob As Stream) As DataRow Implements IDatabase.ExecuteData
        Dim drResult As DataRow = Nothing
        ' The bytes returned from GetBytes.  
        Dim bytesRead As Long
        ' The starting position in the BLOB output.  
        Dim startIndex As Long = 0
        ' The BLOB byte() buffer to be filled by GetBytes.  
        Dim outByte(OptimumChunkSize - 1) As Byte

        EnsureOpen()

        Using scope = New SingletonReleaser(Of SQLiteConnection)(_connection)
            Using command As New SQLiteCommand(sqlCommand, _connection, CType(scope.Transaction, SQLiteTransaction)) With {.CommandTimeout = CommandTimeOut}
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
                                ' Continue while there are bytes beyond the size of the buffer.  
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


    Public Function Execute(sqlCommand As String) As Integer Implements IDatabase.Execute

        'The VACUUM command cannot be used within a transaction
        If sqlCommand.ToLower = "vacuum" Then
            EnsureOpen()
            Using command As New SQLiteCommand(sqlCommand, _connection) With {.CommandTimeout = CommandTimeOut}
                Return command.ExecuteNonQuery()
            End Using
        End If

        Using transaction = BeginLocalTransaction()
            Return Execute(sqlCommand, transaction)
        End Using

    End Function

    Public Function Execute(sqlCommand As String, transactionScope As ITransactionScope) As Integer _
        Implements IDatabase.Execute
        Dim trx = CType(transactionScope, SingletonReleaser(Of SQLiteConnection))
        Dim connection As SQLiteConnection = CType(trx.Transaction.Connection, SQLiteConnection)
        EnsureOpen()
        Using _
            command As _
                New SQLiteCommand(sqlCommand, connection, CType(trx.Transaction, SQLiteTransaction)) With {.CommandTimeout = CommandTimeOut}
            Return command.ExecuteNonQuery()
        End Using
    End Function

    Public Function ExecuteScalar(sqlCommand As String) As Object Implements IDatabase.ExecuteScalar
        EnsureOpen()
        Using command As New SQLiteCommand(sqlCommand, _connection) With {.CommandTimeout = CommandTimeOut}

            Return command.ExecuteScalar()
        End Using
    End Function

    Private Function ExecuteScalar(sqlCommand As String, sqlTransaction As SQLiteTransaction) As Object
        EnsureOpen()
        Using command As New SQLiteCommand(sqlCommand, _connection, sqlTransaction) With {.CommandTimeout = CommandTimeOut}

            Return command.ExecuteScalar()
        End Using
    End Function

    Public Function ExecuteData(sqlCommand As String) As DataTable Implements IDatabase.ExecuteData
        EnsureOpen()
        Using command As New SQLiteCommand(sqlCommand, _connection) With {.CommandTimeout = CommandTimeOut}
            Dim dt = New DataTable()
            dt.Load(command.ExecuteReader(CommandBehavior.SequentialAccess))
            Return dt
        End Using
    End Function

    Public Function ExecuteData(sqlCommand As String, scope As ITransactionScope) As DataTable _
        Implements IDatabase.ExecuteData
        Dim trx = CType(scope, SingletonReleaser(Of SQLiteConnection))
        Dim conn As SQLiteConnection = CType(trx.Transaction.Connection, SQLiteConnection)
        EnsureOpen()
        Using command As New SQLiteCommand(sqlCommand, conn, CType(trx.Transaction, SQLiteTransaction)) With {.CommandTimeout = CommandTimeOut}
            Dim dt = New DataTable()
            dt.Load(command.ExecuteReader(CommandBehavior.SequentialAccess))
            Return dt
        End Using
    End Function

    Public Function GetColumnTypes(sqlCommand As String) As List(Of Tuple(Of String, String, Integer)) _
        Implements IDatabase.GetColumnTypes
        EnsureOpen()
        Using command As New SQLiteCommand(sqlCommand, _connection) With {.CommandTimeout = CommandTimeOut}
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
        Return New SingletonReleaser(Of SQLiteConnection)(_connection)
    End Function

    Public Function InsertRecord64Bits(columnNames() As String, columnValues() As Object, tableName As String,
                                 Optional primaryKey As String = "") As Long Implements IDatabase.InsertRecord64Bits

        'SQLite does not support 64-bits
        Return InsertRecord(columnNames, columnValues, tableName, primaryKey)

    End Function

    Public Function InsertRecord(columnNames() As String, columnValues() As Object, tableName As String,
                                 Optional primaryKey As String = "") As Integer Implements IDatabase.InsertRecord
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
            Throw New ArgumentException($"{NameOf(columnNames)} and {NameOf(columnValues)} do not have the same size!")
        End If
        Dim parms = columnNames.Select(Function(c, i)
                                           If isBlobTable AndAlso i = blobIdx Then
                                               Return $"zeroblob(@p{i})" 'NB: can cause OOM
                                               'Return $"@p{i}"
                                           Else
                                               Return $"@p{i}"
                                           End If
                                       End Function).ToArray()
        Dim sb As New StringBuilder($"INSERT INTO {tableName}(")
        sb.Append($"{String.Join(",", columnNames)}) ")

        sb.Append($"VALUES({String.Join(",", parms)})")


        EnsureOpen()

        Using scope = New SingletonReleaser(Of SQLiteConnection)(_connection)
            Try
                Using command As New SQLiteCommand(sb.ToString(), _connection, CType(scope.Transaction, SQLiteTransaction)) With {.CommandTimeout = CommandTimeOut}
                    Dim rCount As Integer
                    For i = 0 To columnValues.Length - 1
                        ' detect function calls
                        If columnValues(i)?.ToString().EndsWith("()") Then
                            Dim payload = ExecuteScalar($"select {columnValues(i)}")
                            command.Parameters.AddWithValue(parms(i), payload)
                        ElseIf isBlobTable AndAlso blobIdx = i Then
                            command.Parameters.AddWithValue($"@p{i}", input.Length)
                            'command.Parameters.Add(parms(i), Data.DbType.Binary, -1).Value = input
                        Else
                            command.Parameters.AddWithValue(parms(i), columnValues(i))
                        End If

                    Next
                    If returnKey Or isBlobTable Then
                        rCount = command.ExecuteNonQuery()
                        Dim rowId = _connection.LastInsertRowId

                        If isBlobTable Then
                            Dim offset As Integer = 0
                            Dim buffer(OptimumChunkSize - 1) As Byte

                            input.Position = 0
                            Using blob = New SqliteBlobStream(_connection, tableName, BlobColumnName, OptimumChunkSize, Convert.ToInt32(rowId))
                                Using tOutWrite = New TransferStream(blob)
                                    input.CopyTo(tOutWrite, OptimumChunkSize)
                                End Using
                            End Using


                        End If


                        If returnKey Then
                            Return Convert.ToInt32(rowId)
                        Else
                            Return rCount
                        End If
                    Else

                        rCount = command.ExecuteNonQuery()

                        Return rCount
                    End If

                End Using

            Catch ex As Exception
                scope.HasError = True
                If ex.Message.IndexOf("out of memory") >= 0 Then
                    Throw New OutOfMemoryException(ex.Message)
                Else
                    Throw New ApplicationException($"Error in Inserting to {tableName}", ex)
                End If

            End Try
        End Using
    End Function

    Public Function InsertRecord64Bits(columnNames() As String, columnValues() As Object, tableName As String,
                                 scope As ITransactionScope, Optional primaryKey As String = "") As Long _
        Implements IDatabase.InsertRecord64Bits

        'SQLite does not support 64-bits
        Return InsertRecord(columnNames, columnValues, tableName, scope, primaryKey)
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
            Throw New ArgumentException($"{NameOf(columnNames)} and {NameOf(columnValues)} do not have the same size!")
        End If
        Dim parms = columnNames.Select(Function(c, i)
                                           If isBlobTable AndAlso i = blobIdx Then
                                               Return $"zeroblob(@p{i})"
                                           Else
                                               Return $"@p{i}"
                                           End If
                                       End Function).ToArray()
        Dim sb As New StringBuilder($"INSERT INTO {tableName}(")
        sb.Append($"{String.Join(",", columnNames)}) ")

        sb.Append($"VALUES({String.Join(",", parms)})")

        Dim trx = CType(scope, SingletonReleaser(Of SQLiteConnection))
        Dim conn As SQLiteConnection = CType(trx.Transaction.Connection, SQLiteConnection)
        EnsureOpen()
        Using command As New SQLiteCommand(sb.ToString(), conn, CType(trx.Transaction, SQLiteTransaction)) With {.CommandTimeout = CommandTimeOut}

            For i = 0 To columnValues.Length - 1
                ' detect function calls
                If columnValues(i)?.ToString().EndsWith("()") Then
                    Dim payload = ExecuteScalar($"select {columnValues(i)}")
                    command.Parameters.AddWithValue(parms(i), payload)
                ElseIf isBlobTable AndAlso blobIdx = i Then
                    command.Parameters.AddWithValue($"@p{i}", input.Length)
                Else
                    command.Parameters.AddWithValue(parms(i), columnValues(i))
                End If

            Next
            If returnKey Or isBlobTable Then
                Dim rCount = command.ExecuteNonQuery()
                Dim rowId = _connection.LastInsertRowId
                If isBlobTable Then
                    Dim offset As Integer = 0
                    Dim buffer(OptimumChunkSize - 1) As Byte

                    input.Position = 0
                    Using blob = New SqliteBlobStream(_connection, tableName, BlobColumnName, OptimumChunkSize, Convert.ToInt32(rowId))
                        Using tOutWrite = New TransferStream(blob)
                            input.CopyTo(tOutWrite, OptimumChunkSize)
                        End Using
                    End Using

                End If

                Return rCount
            Else
                Return command.ExecuteNonQuery()
            End If

        End Using
    End Function

    Public Function UpdateRecord(keys() As String, columnNames() As String, columnValues() As Object,
                                 tableName As String, scope As ITransactionScope) As Boolean _
        Implements IDatabase.UpdateRecord
        ' build update query

        Dim input As Stream = Nothing
        Dim isBlobTable As Boolean = tableName.IndexOf("_blob", StringComparison.InvariantCultureIgnoreCase) > 0
        Dim blobIdx As Integer = -1


        If columnNames.Length <> columnValues.Length Then
            Throw New ArgumentException($"{NameOf(columnNames)} and {NameOf(columnValues)} do not have the same size!")
        End If

        Dim pairs = columnNames.Zip(columnValues, Function(name, valu) New With {name, .Value = valu}).ToList()
        Dim toUpdate = pairs.Where(Function(kv) Not keys.Contains(kv.name)).ToList()
        If isBlobTable Then
            blobIdx = Array.IndexOf(toUpdate.Select(Function(z) z.name).ToArray(), BlobColumnName)
            input = CType(toUpdate(blobIdx).Value, Stream)
        End If

        Dim wheres = pairs.Where(Function(kv) keys.Contains(kv.name)).ToList()

        Dim parms = toUpdate.Select(Function(c, i)
                                        If isBlobTable AndAlso i = blobIdx Then
                                            Return $"{c.name}=zeroblob(@p{i})"
                                        Else
                                            Return $"{c.name}=@p{i}"
                                        End If
                                    End Function).ToArray()

        Dim constraints = wheres.Select(Function(c, i) $"{c.name}=@w{i}").ToArray()

        Dim sb As New StringBuilder($"UPDATE {tableName} SET ")
        sb.Append($"{String.Join(",", parms)} WHERE ")
        sb.Append($"{String.Join(" AND ", constraints)} ")

        Dim trx = CType(scope, SingletonReleaser(Of SQLiteConnection))
        Dim conn As SQLiteConnection = CType(trx.Transaction.Connection, SQLiteConnection)
        EnsureOpen()
        Using command As New SQLiteCommand(sb.ToString(), conn, CType(trx.Transaction, SQLiteTransaction)) With {.CommandTimeout = CommandTimeOut}

            For Each parUpdate In toUpdate.Select(Function(c, i) New With {.Data = c, .Idx = i})

                ' detect function calls
                If parUpdate.Data.Value?.ToString().EndsWith("()") Then
                    Dim payload = ExecuteScalar($"select {parUpdate.Data.Value}")
                    command.Parameters.AddWithValue($"@p{parUpdate.Idx}", payload)
                ElseIf isBlobTable AndAlso blobIdx = parUpdate.Idx Then
                    command.Parameters.AddWithValue($"@p{parUpdate.Idx}", input.Length)
                Else
                    command.Parameters.AddWithValue($"@p{parUpdate.Idx}", parUpdate.Data.Value)
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

            If isBlobTable Then
                Dim rCount = command.ExecuteNonQuery()
                sb = New StringBuilder($"select blob_id from {tableName} WHERE ")
                sb.Append($"{String.Join(" AND ", constraints)} ")
                Dim rowId As Integer
                command.Parameters.Clear()
                command.CommandText = sb.ToString()

                For Each constrain In wheres.Select(Function(c, i) New With {.Data = c, .Idx = i})
                    ' detect function calls
                    If constrain.Data.Value?.ToString().EndsWith("()") Then
                        Dim payload2 = ExecuteScalar($"select {constrain.Data.Value}")
                        command.Parameters.AddWithValue($"@w{constrain.Idx}", payload2)
                    Else
                        command.Parameters.AddWithValue($"@w{constrain.Idx}", constrain.Data.Value)
                    End If
                Next
                'NB: Assumes we only updated one row (where clause)
                rowId = Convert.ToInt32(command.ExecuteScalar())



                Dim offset As Integer = 0
                Dim buffer(OptimumChunkSize - 1) As Byte

                input.Position = 0

                Using blob = New SqliteBlobStream(_connection, tableName, BlobColumnName, OptimumChunkSize, Convert.ToInt32(rowId))
                    Using tOutWrite = New TransferStream(blob)
                        input.CopyTo(tOutWrite, OptimumChunkSize)
                    End Using
                End Using



                Return (rCount >= 1)
            Else
                Return (command.ExecuteNonQuery() >= 1) ' triggers may also count
            End If

        End Using
    End Function

    Public Function UpdateRecord(keys() As String, columnNames() As String, columnValues() As Object,
                                 tableName As String) As Boolean Implements IDatabase.UpdateRecord
        ' build update query
        If columnNames.Length <> columnValues.Length Then
            Throw New ArgumentException($"{NameOf(columnNames)} and {NameOf(columnValues)} do not have the same size!")
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
        Dim parms = toUpdate.Select(Function(c, i)
                                        If isBlobTable AndAlso i = blobIdx Then
                                            Return $"{c.name}=zeroblob(@p{i})"
                                        Else
                                            Return $"{c.name}=@p{i}"
                                        End If
                                    End Function).ToArray()

        Dim constraints = wheres.Select(Function(c, i) $"{c.name}=@w{i}").ToArray()

        Dim sb As New StringBuilder($"UPDATE {tableName} SET ")
        sb.Append($"{String.Join(",", parms)} WHERE ")
        sb.Append($"{String.Join(" AND ", constraints)} ")

        EnsureOpen()
        Using scope = New SingletonReleaser(Of SQLiteConnection)(_connection)
            Try
                Using command As New SQLiteCommand(sb.ToString(), _connection, CType(scope.Transaction, SQLiteTransaction)) With {.CommandTimeout = CommandTimeOut}

                    For Each parUpdate In toUpdate.Select(Function(c, i) New With {.Data = c, .Idx = i})

                        ' detect function calls
                        If parUpdate.Data.Value?.ToString().EndsWith("()") Then
                            Dim payload = ExecuteScalar($"select {parUpdate.Data.Value}")
                            command.Parameters.AddWithValue($"@p{parUpdate.Idx}", payload)
                        ElseIf isBlobTable AndAlso blobIdx = parUpdate.Idx Then
                            command.Parameters.AddWithValue($"@p{parUpdate.Idx}", input.Length)
                        Else
                            command.Parameters.AddWithValue($"@p{parUpdate.Idx}", parUpdate.Data.Value)
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
                    Dim rCount As Integer
                    If isBlobTable Then
                        rCount = command.ExecuteNonQuery()
                        sb = New StringBuilder($"select blob_id from {tableName} WHERE ")
                        sb.Append($"{String.Join(" AND ", constraints)} ")
                        Dim rowId As Integer
                        command.Parameters.Clear()
                        command.CommandText = sb.ToString()

                        For Each constrain In wheres.Select(Function(c, i) New With {.Data = c, .Idx = i})
                            ' detect function calls
                            If constrain.Data.Value?.ToString().EndsWith("()") Then
                                Dim payload2 = ExecuteScalar($"select {constrain.Data.Value}")
                                command.Parameters.AddWithValue($"@w{constrain.Idx}", payload2)
                            Else
                                command.Parameters.AddWithValue($"@w{constrain.Idx}", constrain.Data.Value)
                            End If
                        Next
                        'NB: Assumes we only updated one row (where clause), one BLOB can't belong to multiple records
                        rowId = Convert.ToInt32(command.ExecuteScalar())


                        Dim offset As Integer = 0
                        Dim buffer(OptimumChunkSize - 1) As Byte

                        input.Position = 0

                        Using blob = New SqliteBlobStream(_connection, tableName, BlobColumnName, OptimumChunkSize, Convert.ToInt32(rowId))
                            Using tOutWrite = New TransferStream(blob)
                                input.CopyTo(tOutWrite, OptimumChunkSize)
                            End Using
                        End Using


                        Return (rCount >= 1)
                    Else
                        rCount = command.ExecuteNonQuery()


                        Return (rCount >= 1) ' triggers may also count
                    End If
                End Using
            Catch ex As Exception
                scope.HasError = True
                If ex.Message.IndexOf("out of memory") >= 0 Then
                    Throw New OutOfMemoryException(ex.Message)
                Else
                    Throw New ApplicationException($"Error in Inserting to {tableName}", ex)
                End If
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
        Dim wheres = constraintKeys.Zip(constraintValues, Function(name, valu) New With {name, .Value = valu}).ToList()
        Dim constraints = wheres.Select(Function(c, i) $"{c.name}=@w{i}").ToArray()

        Dim sb As New StringBuilder($"DELETE FROM {tableName} WHERE ")
        sb.Append($"{String.Join(" AND ", constraints)} ")

        EnsureOpen()
        Using command As New SQLiteCommand(sb.ToString(), _connection) With {.CommandTimeout = CommandTimeOut}

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

    Public Function CreateTableAdapter(selectSQL As String, scope As ITransactionScope, ByRef workTable As DataTable) _
        As IAdapterSession Implements IDatabase.CreateTableAdapter
        Dim trx = CType(scope, SingletonReleaser(Of SQLiteConnection))
        Dim conn As SQLiteConnection = CType(trx.Transaction.Connection, SQLiteConnection)
        Dim command As New SQLiteCommand(selectSQL, conn, CType(trx.Transaction, SQLiteTransaction)) With {.CommandTimeout = CommandTimeOut}
        Dim adapter As New SQLiteDataAdapter(command)
        Dim builder As New SQLiteCommandBuilder(adapter)
        adapter.Fill(workTable)
        builder.GetDeleteCommand()
        builder.GetInsertCommand()
        builder.GetUpdateCommand()
        Return New AdapterSession(command, adapter, builder)
    End Function

#Region "IDisposable Support"

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If _connection IsNot Nothing AndAlso _connection.State <> ConnectionState.Closed AndAlso disposing Then
            _connection.Dispose()
        End If

        _connection = Nothing
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
    End Sub

    Public Sub BulkInsertOrUpdateRecords(targetColumnNames() As String, matchOnColumnNames() As String, sourceTable As DataTable, tableName As String, scope As ITransactionScope) Implements IDatabase.BulkInsertOrUpdateRecords
        Throw New NotImplementedException()
    End Sub

    ''' <summary>
    ''' Generate the SQL query to insert multiple record in one large operation.
    ''' </summary>
    ''' <param name="targetColumnNames">The name of the columns.</param>
    ''' <param name="sourceTable">The data to insert.</param>
    ''' <param name="tableName">The name of the table.</param>
    ''' <param name="beginIndex">The index where to tart.</param>
    ''' <param name="count">How many items to insert.</param>
    ''' <returns></returns>
    Private Function GenerateBulkInsertQuery(targetColumnNames() As String, sourceTable As DataTable, tableName As String, beginIndex As Integer, count As Integer) As String
        Dim sqlStr As New StringBuilder()

        sqlStr.Append($"INSERT INTO {tableName} (")
        For Each aColumn In targetColumnNames
            sqlStr.Append($"{aColumn}, ")
        Next

        ' Remove the trailing ', ' at end of the string.
        sqlStr.Remove(sqlStr.Length - 2, 2)

        sqlStr.Append(") VALUES ")

        For aRowIndex = beginIndex To (beginIndex + count) - 1
            Dim aRow As DataRow = sourceTable.Rows(aRowIndex)

            sqlStr.Append(" (")
            Dim i As Integer
            For i = 0 To targetColumnNames.Length - 1
                If aRow(i).GetType() = GetType(String) Then
                    sqlStr.Append($"'{KillNull(aRow(i))}'")
                Else
                    sqlStr.Append(KillNull(aRow(i)))
                End If

                sqlStr.Append(", ")
            Next

            ' Remove the trailing ', ' at end of the string.
            sqlStr.Remove(sqlStr.Length - 2, 2)

            sqlStr.Append("), ")
        Next

        ' Remove the trailing ', ' at end of the string.
        sqlStr.Remove(sqlStr.Length - 2, 2)

        sqlStr.Append(";")

        Return sqlStr.ToString()
    End Function

    ''' <summary>
    ''' Insert multiple record into the local database in a few large operation.
    ''' </summary>
    ''' <param name="targetColumnNames">The column names.</param>
    ''' <param name="sourceTable">The data to be inserted.</param>
    ''' <param name="tableName">The name of the table.</param>
    ''' <param name="scope">The transaction scope.</param>
    Public Sub BulkInsertRecords(targetColumnNames() As String, sourceTable As DataTable, tableName As String, scope As ITransactionScope) Implements IDatabase.BulkInsertRecords
        Dim beginIndex As Integer = 0
        Dim maxItems As Integer = 1000
        While True
            Dim remainingItems = sourceTable.Rows.Count - beginIndex
            Dim howManyItems = maxItems

            If remainingItems <= maxItems Then
                ' This is the last batch!
                howManyItems = remainingItems
            End If

            Dim sqlStr = GenerateBulkInsertQuery(targetColumnNames, sourceTable, tableName, beginIndex, howManyItems)
            Execute(sqlStr, scope)

            beginIndex += maxItems
            If beginIndex > sourceTable.Rows.Count Then
                Exit While
            End If
        End While
    End Sub

    Public Sub BulkUpdateRecords(targetColumnNames() As String, matchOnColumnNames() As String, sourceTable As DataTable, tableName As String, scope As ITransactionScope) Implements IDatabase.BulkUpdateRecords
        Throw New NotImplementedException()
    End Sub

#End Region

    Private Structure AdapterSession
        Implements IAdapterSession

        Private ReadOnly _command As SQLiteCommand
        Private ReadOnly _adapter As SQLiteDataAdapter
        Private ReadOnly _builder As SQLiteCommandBuilder

        Friend Sub New(command As SQLiteCommand, adapter As SQLiteDataAdapter, builder As SQLiteCommandBuilder)
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



End Class

