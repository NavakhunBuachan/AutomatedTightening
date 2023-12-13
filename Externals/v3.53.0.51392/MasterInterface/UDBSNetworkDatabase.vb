Option Explicit On
Option Compare Text
Option Infer On
Option Strict On

Imports System.Data.Common
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports UdbsInterface.MasterInterface
''' <summary>
'''     Database implementation using SQL .Net Provider for SQL Server
''' </summary>
Friend Class UDBSNetworkDatabase
    Implements IDatabase

#Disable Warning BC42304 ' XML documentation parse error
    ''' <summary>
    ''' The optimum chunk size should be a multiple of 8040 bytes (roughly 8 kb), according to the
    ''' following article: https://docs.microsoft.com/en-us/sql/t-sql/queries/update-transact-sql?redirectedfrom=MSDN&view=sql-server-ver15
    ''' We ran performance tests uploading a 35 Mb file with values such as...
    '''  - Blocks of 8 kb : 120 sec.
    '''  - Blocks of 80 kb : 65 sec.
    '''  - Blocks of 800 kb : 55 sec.
    '''  - Blocks of 1 Mb : 53 sec.
    '''  - Blocks of 8 Mb : 50 sec.
    ''' We decided to go with 1 Mb blocks (or the closest thing there is to it, in multiples of 8040...)
    ''' </summary>
#Enable Warning BC42304 ' XML documentation parse error

    Private Const OptimumChunkSize As Integer = 8040 * 128
    Private Const BlobColumnName As String = "blob_blob"
    Private Const TemporatyBulkCopyTable As String = "#tmpBulkUpdate"

    Public Sub New(UDBSType As UDBS_DBType)
        DBType = UDBSType
    End Sub

    Public Property CommandTimeOut As Integer Implements IDatabase.CommandTimeOut

    Private _connectionString As String
    Private _sqlClientConnectionString As String

    Public Property SQLClientConnectionString As String Implements IDatabase.SQLClientConnectionString
        Get
            Return _sqlClientConnectionString
        End Get
        Private Set(value As String)
            _sqlClientConnectionString = value
        End Set
    End Property

    ''' <summary>
    '''     OldeDB connections string syntax to conform what's stored in the registry
    ''' </summary>
    ''' <returns></returns>
    Public Property ConnectionString As String Implements IDatabase.ConnectionString
        Get
            Return SQLClientToODBC(_connectionString).ToString()
        End Get
        Set
            If String.IsNullOrEmpty(Value) Then
                _connectionString = Value
            ElseIf Value.Contains("(localdb)") Then
                _connectionString = Value
            Else
                'retain the converted SQL Client connection string to compare
                'the output from the SQLClientToODBC does not match the incoming ODBC connection string
                SQLClientConnectionString = ODBCToSQLClient(Value).ToString()
                _connectionString = SQLClientConnectionString
            End If
        End Set
    End Property

    Public ReadOnly Property DBType As UDBS_DBType Implements IDatabase.DBType

    ''' <summary>
    '''     Are we able to connect
    ''' </summary>
    ''' <returns></returns>
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

    Public Function ConnectionStringWithoutPassword() As String Implements IDatabase.ConnectionStringWithoutPassword
        Try
            Dim infoString As String = String.Empty
            Dim tempS As New SqlConnectionStringBuilder(ConnectionString)
            infoString = $"Server:'{tempS.DataSource}', DB:'{tempS.InitialCatalog}', User:'{tempS.UserID}'"
            Return infoString
        Catch ex As Exception
            Return ConnectionString
        End Try
    End Function

    ''' <summary>
    '''     Execute a query which returns number of rows affected
    '''     We Open and Close connections immediately to
    '''     take advantage of connection pooling
    '''     see https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/sql-server-connection-pooling
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    Public Function Execute(sqlCommand As String) As Integer Implements IDatabase.Execute
        Using conn = New SqlConnection(ConnectionString)
            Using command As New SqlCommand(sqlCommand, conn) With {.CommandTimeout = CommandTimeOut}
                conn.Open()
                Return command.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Function ExecuteScalar(sqlCommand As String) As Object Implements IDatabase.ExecuteScalar
        Using conn = New SqlConnection(ConnectionString)
            Using command As New SqlCommand(sqlCommand, conn) With {.CommandTimeout = CommandTimeOut}
                conn.Open()
                Return command.ExecuteScalar()
            End Using
        End Using
    End Function




    Public Function GetColumnTypes(sqlCommand As String) As List(Of Tuple(Of String, String, Integer)) _
        Implements IDatabase.GetColumnTypes
        Using conn = New SqlConnection(ConnectionString)
            Using command As New SqlCommand(sqlCommand, conn) With {.CommandTimeout = CommandTimeOut}
                conn.Open()
                Using reader = command.ExecuteReader()

                    Using schemaTable = reader.GetSchemaTable()
                        Return schemaTable.Rows.Cast(Of DataRow).
                            Select(
                                Function(dr) _
                                      Tuple.Create(dr.Field(Of String)("ColumnName"),
                                                   dr.Field(Of String)("DataTypeName"),
                                                   dr.Field(Of Integer)("ColumnSize"))).ToList()
                    End Using
                End Using
            End Using
        End Using
    End Function

    Public Shared Function ODBCToSQLClient(odbsFormatConnectionString As String) As SqlConnectionStringBuilder
        Dim odbcBuilder = New OdbcConnectionStringBuilder(odbsFormatConnectionString)
        Dim sqlBuilder = New SqlConnectionStringBuilder()

        sqlBuilder("Data Source") = odbcBuilder("Server")
        sqlBuilder("Initial Catalog") = odbcBuilder("Database")

        If odbcBuilder.ContainsKey("Uid") Then
            sqlBuilder("User Id") = odbcBuilder("Uid")
            sqlBuilder("Password") = odbcBuilder("Pwd")
        End If

        If odbcBuilder.ContainsKey("Trusted_Connection") Then
            sqlBuilder("Integrated Security") = odbcBuilder("Trusted_Connection").ToString().Equals("yes", StringComparison.OrdinalIgnoreCase)
        End If

        If odbcBuilder.ContainsKey("Application Name") Then
            ' The application name was provided in the connection string.
            sqlBuilder("Application Name") = odbcBuilder("Application Name")
        Else
            ' The application name was not provided. Use reflection to determine it.
            sqlBuilder("Application Name") = DetermineSoftwareName()
        End If

        Return sqlBuilder
    End Function

    Private Shared Function SQLClientToODBC(sqlConnectionString As String) As OdbcConnectionStringBuilder
        Dim sqlBuilder = New SqlConnectionStringBuilder(sqlConnectionString)
        Dim odbcBuilder = New OdbcConnectionStringBuilder()
        odbcBuilder("Server") = sqlBuilder("Data Source")
        odbcBuilder("Database") = sqlBuilder("Initial Catalog")
        If sqlBuilder.ContainsKey("User Id") Then
            odbcBuilder("Uid") = sqlBuilder("User Id")
            odbcBuilder("Pwd") = sqlBuilder("Password")
        End If
        If sqlBuilder.ContainsKey("Integrated Security") Then
            odbcBuilder("Trusted_Connection") = Convert.ToBoolean(sqlBuilder("Integrated Security"))
        End If
        If sqlBuilder.ContainsKey("Application Name") Then
            odbcBuilder("Application Name") = sqlBuilder("Application Name")
        Else
            ' The application name was not provided. Use reflection to determine it.
            odbcBuilder("Application Name") = DetermineSoftwareName()
        End If

        Return odbcBuilder
    End Function

    Public Function BeginTransaction() As ITransactionScope Implements IDatabase.BeginTransaction
        Dim conn = New SqlConnection(ConnectionString)
        conn.Open()
        Dim trx As SqlTransaction = conn.BeginTransaction()
        Return New Releaser(Of SqlTransaction)(trx)
    End Function

    ''' <summary>
    '''     Note: Call <see cref="BeginTransaction" />first to get a handle
    ''' </summary>
    ''' <param name="sqlCommand"></param>
    ''' <param name="transactionScope"></param>
    ''' <returns></returns>
    Public Function ExecuteData(sqlCommand As String, transactionScope As ITransactionScope) As DataTable _
        Implements IDatabase.ExecuteData
        Dim trx = CType(transactionScope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection
        Using command As New SqlCommand(sqlCommand, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
            Dim dt = New DataTable()
            dt.Load(command.ExecuteReader(CommandBehavior.SequentialAccess))
            Return dt

        End Using
    End Function

    Public Function ExecuteData(sqlCommand As String) As DataTable Implements IDatabase.ExecuteData
        Using conn = New SqlConnection(ConnectionString)
            Using command As New SqlCommand(sqlCommand, conn) With {.CommandTimeout = CommandTimeOut}
                conn.Open()
                Dim dt = New DataTable()
                dt.Load(command.ExecuteReader(CommandBehavior.SequentialAccess))
                Return dt

            End Using
        End Using
    End Function

    Public Function ExecuteData(sqlCommand As String, ByRef blob As Stream) As DataRow Implements IDatabase.ExecuteData

        Dim drResult As DataRow = Nothing
        ' The bytes returned from GetBytes.  
        Dim bytesRead As Long
        ' The starting position in the BLOB output.  
        Dim startIndex As Long = 0
        ' The BLOB byte() buffer to be filled by GetBytes.  
        Dim outByte(OptimumChunkSize - 1) As Byte
        Using conn = New SqlConnection(ConnectionString)
            Using command As New SqlCommand(sqlCommand, conn) With {.CommandTimeout = CommandTimeOut}
                conn.Open()
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

    Public Function CreateTableAdapter(selectSQL As String, scope As ITransactionScope, ByRef workTable As DataTable) _
        As IAdapterSession Implements IDatabase.CreateTableAdapter
        Dim trx = CType(scope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection
        Dim command As New SqlCommand(selectSQL, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
        Dim adapter As New SqlDataAdapter(command)
        Dim builder As New SqlCommandBuilder(adapter)
        adapter.Fill(workTable)
        builder.GetDeleteCommand()
        builder.GetInsertCommand()
        builder.GetUpdateCommand()
        Return New AdapterSession(command, adapter, builder)
    End Function

    Private Structure AdapterSession
        Implements IAdapterSession

        Private ReadOnly _command As SqlCommand
        Private ReadOnly _adapter As SqlDataAdapter
        Private ReadOnly _builder As SqlCommandBuilder

        Friend Sub New(command As SqlCommand, adapter As SqlDataAdapter, builder As SqlCommandBuilder)
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
        Dim trx = CType(transactionScope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection

        Using command As New SqlCommand(sqlCommand, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
            Return command.ExecuteNonQuery()
        End Using
    End Function

    Public Function InsertRecord(columnNames() As String, columnValues() As Object, tableName As String,
                                 scope As ITransactionScope, Optional primaryKey As String = "") As Integer _
        Implements IDatabase.InsertRecord

        Return CInt(InsertRecord64Bits(columnNames, columnValues, tableName, scope, primaryKey))
    End Function

    ''' <summary>
    ''' Converts two lists of the same size into a list of key-value pairs.
    ''' </summary>
    ''' <param name="keys">The keys.</param>
    ''' <param name="values">The values.</param>
    ''' <returns>The combined list.</returns>
    Private function ToKeyPairs(keys() as string, values() as Object) As List(Of KeyValuePair(Of String, Object))
        If keys.Length <> values.Length Then
            Throw New ArgumentException($"{NameOf(keys)} and {NameOf(values)} do not have the same size!")
        End If

        return keys.Zip(values, Function(name, value) New KeyValuePair(Of String, Object)(name, value)).ToList()
    End function

    ''' <summary>
    ''' Generate the SQL query string that will perform an UPDATE onto a single table.
    ''' </summary>
    ''' <param name="keys">The keys (for the WHERE clause).</param>
    ''' <param name="columnNames">
    ''' The name of the columns.
    ''' Includes both the columns that are the keys for the query and the values to be inserted.
    ''' </param>
    ''' <param name="columnValues">
    ''' The values to update and the values of the keys needed to select what row to update.
    ''' </param>
    ''' <param name="tableName">The name of the table.</param>
    ''' <returns>The SQL query to be performed.</returns>
    Private function GenerateUpdateQueryString(keys() As String, columnNames() As String, columnValues() As Object,
                                 tableName As String) As String
        ' Convert column names and values to a list of key-value pairs.
        Dim pairs = ToKeyPairs(columnNames, columnValues)

        ' The values to update are those that are not keys.
        Dim toUpdate = pairs.Where(Function(kv) Not keys.Contains(kv.Key))

        ' Why are we making this call, instead of simply using 'keys' ?
        ' My guess is that we are validating that the 'keys' are also part of the
        ' 'column names'.
        ' But if they aren't, then this will not work or, worst, would overwrite
        ' more than one row.
        Dim wheres = pairs.Where(Function(kv) keys.Contains(kv.Key))

        Dim parms = toUpdate.Select(Function(c, i) $"{c.Key}=@p{i}")
        Dim constraints = wheres.Select(Function(c, i) $"{c.Key}=@w{i}")

        Dim sb As New StringBuilder($"UPDATE {tableName} SET ")
        sb.Append($"{String.Join(",", parms)} WHERE ")
        sb.Append($"{String.Join(" AND ", constraints)} ")

        Return sb.ToString()
    End function

    ''' <summary>
    ''' Generate the SQL query to retrive the BLOB ID column of a given table.
    ''' </summary>
    ''' <param name="keys">The keys used for the WHERE clause to determine what row to select.</param>
    ''' <param name="tableName">The name of the table to query.</param>
    ''' <returns>The SQL command to execute in order to get the BLOB ID.</returns>
    Private function GenerateGetBlogIdQueryString(
            keys() As String,
            tableName As String) As String
        Dim constraints = keys.Select(Function(c, i) $"{c}=@w{i}")

        dim sb = New StringBuilder($"select blob_id from {tableName} WHERE ")
        sb.Append($"{String.Join(" AND ", constraints)} ")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Update the parameters of an SQL command object.
    ''' </summary>
    ''' <param name="command"></param>
    ''' <param name="keys">
    ''' The values of the keys to filter on in order to determine what row
    ''' to UPDATE.
    ''' When updating an INSERT command, there is no WHERE clause and the 
    ''' caller does not need to provide any keys.</param>
    ''' <param name="columnNames"></param>
    ''' <param name="columnValues"></param>
    Private sub UpdateParametersValues(
            command As SqlCommand,
            columnNames() As String,
            columnValues() As Object,
            Optional keys() As String = Nothing)
        Dim pairs = ToKeyPairs(columnNames, columnValues)
        Dim toUpdate As List(Of KeyValuePair(Of String, Object))
        If keys IsNot Nothing Then
            toUpdate = pairs.Where(Function(kv) Not keys.Contains(kv.Key)).ToList()
        Else
            toUpdate = pairs
        End If
        Dim wheres = pairs.Where(Function(kv) keys.Contains(kv.Key))

        Dim blobIndex As Integer = -1
        If IsBlobTable(columnNames) Then
            blobIndex = Array.IndexOf(toUpdate.Select(Function(z) z.Key).ToArray(), BlobColumnName)
        End If

        Dim i As Integer = 0
        For Each parUpdate In toUpdate
            If parUpdate.Value?.ToString().EndsWith("()") Then
                ' This is a function call.
                Dim payload = ExecuteScalar($"select {parUpdate.Value}")
                command.Parameters.AddWithValue($"@p{i}", payload)
            ElseIf blobIndex = i Then
                'NB: dummy payload to be replaced by a streaming scheme compatible to .Net 4.0
                ' Streaming, to avoid OOM issues
                command.Parameters.AddWithValue($"@p{i}", New Byte() {})
            Else
                If parUpdate.Value Is Nothing Then
                    command.Parameters.AddWithValue($"@p{i}", DBNull.Value)
                Else If parUpdate.Value.GetType() Is GetType(Double) Then
                    command.Parameters.AddWithValue($"@p{i}", Clamp(CType(parUpdate.Value, Double), Single.MinValue, Single.MaxValue))
                Else
                    command.Parameters.AddWithValue($"@p{i}", parUpdate.Value)
                End If
            End If

            i += 1
        Next

        If keys IsNot Nothing Then
            UpdateConstraintsParameters(command, keys, columnNames, columnValues)
        End If
    End sub

    ''' <summary>
    ''' Update the WHERE paremeters of an SQL command.
    ''' </summary>
    ''' <param name="command">The SQL command to update.</param>
    ''' <param name="keys">The keys for the WHERE clause.</param>
    ''' <param name="columnNames">The name of the columns. Every key column must be present.</param>
    ''' <param name="columnValues">The values of the columns.</param>
    Private sub UpdateConstraintsParameters(
            command As SqlCommand, 
            keys() As String, 
            columnNames() as String,
            columnValues() As Object)
        Dim i As Integer = 0
        For Each constraint In keys
            Dim columnIndex = columnNames.ToList().IndexOf(keys(i))
            Dim value = columnValues(columnIndex)
            If value?.ToString().EndsWith("()") Then
                ' This is a function call.
                Dim payload = ExecuteScalar($"select {value}")
                command.Parameters.AddWithValue($"@w{i}", payload)
            Else
                command.Parameters.AddWithValue($"@w{i}", value)
            End If

            i += 1
        Next
    End sub

    ''' <summary>
    ''' Determine if a given table is a BLOB table.
    ''' Check for the presence of the 'blob_blob' column.
    ''' </summary>
    ''' <param name="columnNames">The columns of the table to update.</param>
    ''' <returns>Whether or not this is a BLOB table.</returns>
    Private function IsBlobTable(columnNames As IEnumerable(Of String)) As Boolean
        return columnNames.Contains(BlobColumnName)
    End function

    ''' <summary>
    ''' Determine if this is a BLOB table based on the table name.
    ''' </summary>
    ''' <param name="tableName">The name of the table.</param>
    ''' <returns>Whether or not this is a BLOB table.</returns>
    Private function IsBlobTable(tableName As string) As Boolean
        return tableName.IndexOf("_blob", StringComparison.InvariantCultureIgnoreCase) > 0
    End function

    'NB: Assumes we only updated one row (where clause), one BLOB can't belong to multiple records
    Private function GetBlobId(keys() As String, columnNames() As String, columnValues() As Object,
                               tableName As String, scope As ITransactionScope) As Integer
        If Not IsBlobTable(tableName) Then
            Return -1
        End If

        Dim trx = CType(scope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection

        Dim getBlobIdSqlQueryStr = GenerateGetBlogIdQueryString(keys, tableName)
        Using command As New SqlCommand(getBlobIdSqlQueryStr, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
            UpdateConstraintsParameters(command, keys, columnNames, columnValues)

            'NB: Assumes we only updated one row (where clause), one BLOB can't belong to multiple records
            Return Convert.ToInt32(command.ExecuteScalar())
        End Using
    End function

    ' This method has one of the higher C.R.A.P. score.
    Public Function UpdateRecord(keys() As String, columnNames() As String, columnValues() As Object,
                                 tableName As String, scope As ITransactionScope) As Boolean _
        Implements IDatabase.UpdateRecord

        Dim trx = CType(scope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection

        Dim pairs = ToKeyPairs(columnNames, columnValues)
        Dim toUpdate = pairs.Where(Function(kv) Not keys.Contains(kv.Key)).ToList()
        Dim wheres = pairs.Where(Function(kv) keys.Contains(kv.Key)).ToList()

        Dim updatedRowCount As Integer = 0
        Dim sqlCommandStr = GenerateUpdateQueryString(keys, columnNames, columnValues, tableName)
        Using command As New SqlCommand(sqlCommandStr, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
            UpdateParametersValues(command, columnNames, columnValues, keys)
            updatedRowCount = command.ExecuteNonQuery()
        End using

        If IsBlobTable(tableName) Then
            Dim rowId = GetBlobId(keys, columnNames, columnValues, tableName, scope)
            Dim updateBlobSqlQueryStr = $"SELECT @Pointer = TEXTPTR([{BlobColumnName}]) FROM [{tableName}] WHERE blob_id = {rowId}"
            Using command As New SqlCommand(updateBlobSqlQueryStr, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
                Dim ptrParm As SqlParameter = command.Parameters.Add("@Pointer", SqlDbType.Binary, 16)
                ptrParm.Direction = ParameterDirection.Output
                command.ExecuteNonQuery()

                Using blob = New SqlImageStream(conn, trx.Transaction, tableName, BlobColumnName, OptimumChunkSize, CType(ptrParm.Value, Byte()))
                    Using tOutWrite = New TransferStream(blob)
                        Dim blobIdx As Integer = Array.IndexOf(toUpdate.Select(Function(z) z.Key).ToArray(), BlobColumnName)
                        Dim input As Stream = CType(toUpdate(blobIdx).Value, Stream)
                        input.CopyTo(tOutWrite, OptimumChunkSize)
                    End Using
                End Using
            End Using
        End If

        Return updatedRowCount >= 1
    End Function

    Public Function InsertRecord(columnNames() As String, columnValues() As Object,
                                 tableName As String, Optional ByVal primaryKey As String = "") As Integer _
        Implements IDatabase.InsertRecord

        Return CInt(InsertRecord64Bits(columnNames, columnValues, tableName, primaryKey))
    End Function

    Public Function InsertRecord64Bits(columnNames() As String, columnValues() As Object,
                                 tableName As String, Optional ByVal primaryKey As String = "") As Long _
        Implements IDatabase.InsertRecord64Bits

        Using transaction = BeginTransaction
            return InsertRecord64Bits(columnNames, columnValues, tableName, transaction, primaryKey)
        End Using
    End Function

    ''' <summary>
    ''' Generate the SQL command string to perform an INSERT operation.
    ''' </summary>
    ''' <param name="columnNames">The name of the columns.</param>
    ''' <param name="tableName">The name of the tables.</param>
    ''' <param name="primaryKey">
    ''' (Optional) The primary key.
    ''' If the primary key parameter is provided, the caller of the INSERT operation
    ''' will expect the key of the inserted row to be returned, instead of the number
    ''' of rows affected.
    ''' </param>
    ''' <returns>The SQL query to be performed.</returns>
    Private function GenerateInsertSqlRequest(
            columnNames As String(), tableName As String,
            Optional primaryKey As String = "") As string

        Dim returnInsertedKey As Boolean = Not String.IsNullOrEmpty(primaryKey)

        Dim parms = columnNames.Select(Function(c, i) $"@p{i}").ToArray()
        Dim sb As New StringBuilder($"INSERT INTO {tableName}(")
        sb.Append($"{String.Join(",", columnNames)}) ")
        sb.Append($"VALUES({String.Join(",", parms)})")
        If returnInsertedKey OrElse IsBlobTable(tableName) Then
            sb.Append($"SELECT @@IDENTITY")
        End If

        Return sb.ToString()
    End function

    ''' <inheritdoc/>
    Public Function InsertRecord64Bits(
            columnNames As String(), 
            columnValues As Object(), 
            tableName As String,
            scope As ITransactionScope, 
            Optional ByVal primaryKey As String = "") As Long Implements IDatabase.InsertRecord64Bits
        Dim trx = CType(scope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection
        Dim returnInsertedKey As Boolean = Not String.IsNullOrEmpty(primaryKey)

        If columnNames.Length <> columnValues.Length Then
            Throw New ArgumentException($"{NameOf(columnNames)} and {NameOf(columnValues)} do not have the same size!")
        End If

        If IsBlobTable(tableName) Then
            primaryKey = "blob_id"
        End If

        Dim numberOfRowsInserted As Long
        Dim keyOfInsertedRow As Long
        Dim insertSqlRequest = GenerateInsertSqlRequest(columnNames, tableName, primaryKey)
        Using command As New SqlCommand(insertSqlRequest, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
            UpdateParametersValues(command, columnNames, columnValues)

            If returnInsertedKey Or IsBlobTable(tableName) Then
                keyOfInsertedRow = Convert.ToInt64(command.ExecuteScalar())
                numberOfRowsInserted = 1
            Else
                numberOfRowsInserted = command.ExecuteNonQuery()
            End If
        End Using

        If IsBlobTable(tableName) Then
            Dim selectBlobSqlQuery = $"SELECT @Pointer = TEXTPTR([{BlobColumnName}]) FROM [{tableName}] WHERE {primaryKey} = {keyOfInsertedRow}"
            Using command As New SqlCommand(selectBlobSqlQuery, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
                Dim ptrParm As SqlParameter = command.Parameters.Add("@Pointer", SqlDbType.Binary, 16)
                ptrParm.Direction = ParameterDirection.Output
                command.ExecuteNonQuery()

                Using blob = New SqlImageStream(conn, trx.Transaction, tableName, BlobColumnName, OptimumChunkSize, CType(ptrParm.Value, Byte()))
                    Using tOutWrite = New TransferStream(blob)
                        Dim blobIdx = Array.IndexOf(columnNames, BlobColumnName)
                        Dim input As Stream = CType(columnValues(blobIdx), Stream)
                        input.CopyTo(tOutWrite, OptimumChunkSize)
                    End Using
                End Using
            End Using
        End If

        If returnInsertedKey Then
            Return keyOfInsertedRow
        Else
            Return numberOfRowsInserted
        End If
    End Function

    Public Function UpdateRecord(keys() As String, columnNames() As String, columnValues() As Object,
                                 tableName As String) As Boolean Implements IDatabase.UpdateRecord
        Using trx = Me.BeginTransaction()
            Try
                Return UpdateRecord(keys, columnNames, columnValues, tableName, trx)
            Catch ex As Exception
                trx.HasError = True
                Throw New ApplicationException($"Error updating {tableName}", ex)
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

        Using conn = New SqlConnection(ConnectionString)
            Using command As New SqlCommand(sb.ToString(), conn) With {.CommandTimeout = CommandTimeOut}
                conn.Open()
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
        End Using
    End Function

    ''' <summary>
    ''' Vanilla use of <see cref="SqlBulkCopy"></see> within a transaction scope/>
    ''' First copying the schema, fill datatable and use it as the payload
    ''' </summary>
    ''' <param name="targetColumnNames"></param>
    ''' <param name="sourceTable"></param>
    ''' <param name="tableName"></param>
    ''' <param name="scope"></param>
    Public Sub BulkInsertRecords(targetColumnNames() As String, sourceTable As DataTable, tableName As String, scope As ITransactionScope) Implements IDatabase.BulkInsertRecords
        Dim trx = CType(scope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection
        Dim dt As DataTable
        Dim dummySql = $"select top 0 * from {tableName} with(nolock)"
        Using command As New SqlCommand(dummySql, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
            dt = New DataTable()
            dt.Load(command.ExecuteReader())
        End Using
        For Each drSource In sourceTable.AsEnumerable()
            Dim dr = dt.NewRow()
            For Each col In targetColumnNames

                If (drSource(col).GetType() Is GetType(Double)) Then

                    dr(col) = Clamp(CType(drSource(col), Double), Single.MinValue, Single.MaxValue)
                Else
                    dr(col) = drSource(col)
                End If

            Next
            dt.Rows.Add(dr)
        Next
        Using bulkCopy =
            New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, trx.Transaction)
            bulkCopy.BulkCopyTimeout = DBCommandTimeout
            bulkCopy.DestinationTableName = tableName
            ' Write from the source to the destination.
            bulkCopy.WriteToServer(dt)
        End Using
    End Sub



    ''' <summary>
    ''' Make use of the temp table to bulk-insert, then update from there
    ''' </summary>
    ''' <param name="targetColumnNames"></param>
    ''' <param name="matchOnColumnNames"></param>
    ''' <param name="sourceTable"></param>
    ''' <param name="tableName"></param>
    ''' <param name="scope"></param>
    Public Sub BulkUpdateRecords(targetColumnNames() As String, matchOnColumnNames() As String, sourceTable As DataTable, tableName As String, scope As ITransactionScope) Implements IDatabase.BulkUpdateRecords
        Try
            BulkUpdate_CreateTemporaryTable(tableName, scope)

            Try
                BulkUpdate_InsertIntoTemporaryTable(scope, sourceTable, targetColumnNames)
                BulkUpdate_UpdateFromTemporaryTable(tableName, scope, targetColumnNames, matchOnColumnNames)
            Finally
                ' We drop the table, even if this is a temporary table, so that the BulkUpdateRecords(...) 
                ' method could be used again within the same transaction scope to update a different table.
                BulkUpdate_DropTemporaryTable(scope)
            End Try
        Catch e As Exception
            Throw New UDBSException("Error performing bulk update.", e)
        End Try
    End Sub

    ''' <summary>
    ''' Main Engine of Upsert
    ''' </summary>
    ''' <param name="targetColumnNames">Columns we intend to write</param>
    ''' <param name="matchOnColumnNames">Columns to match. Could be included in targetColumns</param>
    ''' <param name="sourceTable"></param>
    ''' <param name="tableName"></param>
    ''' <param name="scope"></param>
    ''' <param name="upsert"></param>
    <Obsolete("This method is unsafe. It is causing SQL Deadlocks. See TMTD-208. It might be taken out after June 2022.", False)>
    Private Sub BCP_Upsert(targetColumnNames() As String,
                           matchOnColumnNames() As String,
                           sourceTable As DataTable,
                           tableName As String,
                           scope As ITransactionScope,
                           upsert As Boolean)
        Dim trx = CType(scope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection
        Dim dummySql As String

#Region "Create Temp Table"
        'NB: Temp Table only exist within scope
        Dim tempTable As String = "#tmpBulkUpdate"
        dummySql = $"select top 0 * into {tempTable} from {tableName} with(nolock)"
        Using command As New SqlCommand(dummySql, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
            command.ExecuteNonQuery()
        End Using
#End Region

        Try

            ' Bulk Copy to the temp table first
#Region "BCP to the temp table"
            Dim dt As DataTable
            dummySql = $"select top 0 * from {tempTable} with(nolock)"
            Using command As New SqlCommand(dummySql, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
                dt = New DataTable()
                dt.Load(command.ExecuteReader())
            End Using
            For Each drSource In sourceTable.AsEnumerable()
                Dim dr = dt.NewRow()
                For Each col In targetColumnNames

                    If (drSource(col).GetType() Is GetType(Double)) Then
                        ' Validate numeric values.
                        ' SQLite stores 'Double', but 'UDBS' (SQL Server) stores 'Single' (Float).
                        ' Values such as 'NaN' or 'Double.Infinity' will get rejected
                        ' by the server.

                        dr(col) = Clamp(CType(drSource(col), Double), Single.MinValue, Single.MaxValue)
                    Else
                        dr(col) = drSource(col)
                    End If

                Next
                dt.Rows.Add(dr)
            Next
            Using bulkCopy =
                New SqlBulkCopy(conn, SqlBulkCopyOptions.KeepIdentity, trx.Transaction)
                bulkCopy.BulkCopyTimeout = DBCommandTimeout
                bulkCopy.DestinationTableName = tempTable
                ' Write from the source to the destination.
                bulkCopy.WriteToServer(dt)
            End Using
#End Region

            Dim matched = matchOnColumnNames.Select(Function(z) $"[Target].[{z}]=[Source].[{z}]")
            Dim toUpdate = targetColumnNames.Where(Function(k) Not matchOnColumnNames.Contains(k)).ToArray()

            Dim targets = toUpdate.Select(Function(z) $"[Target].[{z}]=[Source].[{z}]")

            ' Merge, see: https://docs.microsoft.com/en-us/sql/t-sql/statements/merge-transact-sql?view=sql-server-ver15&viewFallbackFrom=sql-server-previousversions
            dummySql = $"MERGE INTO {tableName} WITH (HOLDLOCK) AS Target 
                   USING  {tempTable} AS Source
                   ON ( {String.Join(" AND ", matched)} ) 
                   WHEN MATCHED THEN 
                   UPDATE SET 
                   {String.Join(",", targets)} {{0}};
                   DROP TABLE {tempTable};"

            Dim insertSql As String = String.Empty
            If upsert Then
                Dim insertValues = targetColumnNames.Select(Function(z) $"[Source].[{z}]")
                insertSql = $"WHEN NOT MATCHED BY TARGET THEN 
                            INSERT ({String.Join(",", targetColumnNames)}) VALUES ({String.Join(",", insertValues)})"
            End If

            dummySql = String.Format(dummySql, insertSql)
            ' Dropped Temp Table too
            Using command As New SqlCommand(dummySql, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
                command.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            logger.Error(ex, $"Error bulk-{If(upsert, "upsert", "update")} {tableName}")
            ' clean-up temp table if any
            dummySql = $"if object_id('tempdb..{tempTable}') is not null drop table {tempTable}"
            Using command As New SqlCommand(dummySql, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
                command.ExecuteNonQuery()
            End Using
            'rethrow
            Throw New UDBSException($"Error bulk-{If(upsert, "upsert", "update")} {tableName}", ex)
        End Try
    End Sub

    Private Sub BulkUpdate_CreateTemporaryTable(tableName As String, scope As ITransactionScope)
        Dim trx = CType(scope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection
        Dim sqlCommandStr As String = $"select top 0 * into {TemporatyBulkCopyTable} from {tableName} with(nolock)"
        Using command As New SqlCommand(sqlCommandStr, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
            command.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub BulkUpdate_DropTemporaryTable(scope As ITransactionScope)
        Dim trx = CType(scope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection
        Dim sqlCommandStr As String = $"DROP TABLE {TemporatyBulkCopyTable}"
        Using command As New SqlCommand(sqlCommandStr, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
            command.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub BulkUpdate_InsertIntoTemporaryTable(
            scope As ITransactionScope,
            sourceTable As DataTable,
            targetColumnNames() As String)
        Dim trx = CType(scope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection

        ' Bulk Copy to the temp table first
        Dim dt As DataTable
        Dim dummySql As String = $"select top 0 * from {TemporatyBulkCopyTable} with(nolock)"
        Using command As New SqlCommand(dummySql, conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
            dt = New DataTable()
            dt.Load(command.ExecuteReader())
        End Using
        For Each drSource In sourceTable.AsEnumerable()
            Dim dr = dt.NewRow()
            For Each col In targetColumnNames
                If (drSource(col).GetType() Is GetType(Double)) Then
                    dr(col) = Clamp(CType(drSource(col), Double), Single.MinValue, Single.MaxValue)
                Else
                    dr(col) = drSource(col)
                End If
            Next
            dt.Rows.Add(dr)
        Next
        Using bulkCopy =
                New SqlBulkCopy(conn, SqlBulkCopyOptions.KeepIdentity, trx.Transaction)
            bulkCopy.BulkCopyTimeout = DBCommandTimeout
            bulkCopy.DestinationTableName = TemporatyBulkCopyTable
            ' Write from the source to the destination.
            bulkCopy.WriteToServer(dt)
        End Using
    End Sub

    Private Sub BulkUpdate_UpdateFromTemporaryTable(tableName As String, scope As ITransactionScope, targetColumnNames() As String, matchOnColumnNames() As String)
        Dim trx = CType(scope, Releaser(Of SqlTransaction))
        Dim conn As SqlConnection = trx.Transaction.Connection

        Dim sourceTableAlias = "sourceTable"
        Dim destinationTableAlias = "destinationTable"

        Dim sqlCommandStr As New StringBuilder()
        sqlCommandStr.Append($"UPDATE {tableName} SET ")
        For Each aColumn In targetColumnNames
            If (matchOnColumnNames.Contains(aColumn)) Then
                ' Only include match names.
                Continue For
            End If

            sqlCommandStr.Append($"{tableName}.{aColumn} = {sourceTableAlias}.{aColumn}, ")
        Next

        ' Removing the trailing ", " at the end of the list of values.
        Dim numberOfCharactersToRemove = ", ".Length
        sqlCommandStr.Remove(sqlCommandStr.Length - numberOfCharactersToRemove, numberOfCharactersToRemove)

        sqlCommandStr.Append($" FROM {tableName} {destinationTableAlias} ")
        sqlCommandStr.Append($" INNER JOIN {TemporatyBulkCopyTable} {sourceTableAlias} ")

        sqlCommandStr.Append(" ON ")
        For Each aColumn In targetColumnNames
            If (Not matchOnColumnNames.Contains(aColumn)) Then
                ' Excluding columns on which we do the match.
                Continue For
            End If

            sqlCommandStr.Append($"{sourceTableAlias}.{aColumn} = {destinationTableAlias}.{aColumn} AND ")
        Next

        ' Remove training " AND "
        numberOfCharactersToRemove = " AND ".Length
        sqlCommandStr.Remove(sqlCommandStr.Length - numberOfCharactersToRemove, numberOfCharactersToRemove)

        ' Resulting SQL will look like:
        '
        ' UPDATE testdata_result
        ' SET testdata_result.result_process_id = sourceTable.result_process_id, 
        '     testdata_result.result_itemlistdef_id = sourceTable.result_itemlistdef_id,
        '     testdata_result.result_value = sourceTable.result_value,
        '     testdata_result.result_passflag = sourceTable.result_passflag, 
        '     testdata_result.result_stringdata = sourceTable.result_stringdata, 
        '     testdata_result.result_blobdata_exists = sourceTable.result_blobdata_exists 
        ' FROM testdata_result destinationTable 
        ' INNER JOIN #tmpBulkUpdate sourceTable
        ' ON sourceTable.result_id = destinationTable.result_id

        Using command As New SqlCommand(sqlCommandStr.ToString(), conn, trx.Transaction) With {.CommandTimeout = CommandTimeOut}
            command.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>
    ''' Make use of the temp table to bulk-insert, then update from there
    ''' If records match then update else insert
    ''' </summary>
    ''' <param name="targetColumnNames"></param>
    ''' <param name="matchOnColumnNames"></param>
    ''' <param name="sourceTable"></param>
    ''' <param name="tableName"></param>
    ''' <param name="scope"></param>
    Public Sub BulkInsertOrUpdateRecords(targetColumnNames() As String, matchOnColumnNames() As String, sourceTable As DataTable, tableName As String, scope As ITransactionScope) Implements IDatabase.BulkInsertOrUpdateRecords
        BCP_Upsert(targetColumnNames, matchOnColumnNames, sourceTable, tableName, scope, True)
    End Sub

    Dim disposed As Boolean = False
    Public Sub Dispose() _
        Implements IDisposable.Dispose
        ' Dispose of unmanaged resources.
        Dispose(True)
        ' Suppress finalization.
        GC.SuppressFinalize(Me)
    End Sub
    ' Protected implementation of Dispose pattern.
    Protected Overridable Sub Dispose(disposing As Boolean)
        If disposed Then Return

        If disposing Then
            ' Free any other managed objects here.
            '
        End If

        ' Free any unmanaged objects here.
        '
        disposed = True
    End Sub


End Class
