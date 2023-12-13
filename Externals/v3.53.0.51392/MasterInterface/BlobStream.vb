Imports System.Collections.Concurrent
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Data.SQLite
Imports System.IO
Imports System.Threading.Tasks

''' <summary>
''' Pseudo-Stream class to abstract the SQL UPDATETEXT  Syntax
''' This class is necessary because .Net4 does not support SqlStreams
''' see: https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/sqlclient-streaming-support
''' This is a work around for the image (old) data type !!!!
''' </summary>
Friend Class SqlImageStream
    Inherits AbstractDBStream

    Private ReadOnly cmdAppendChunk As SqlCommand
    Private ReadOnly cmdFirstChunk As SqlCommand
    Private ReadOnly connection As SqlConnection
    Private ReadOnly transaction As SqlTransaction
    Private ReadOnly blobParm As SqlParameter
    Private ReadOnly offsetParm As SqlParameter
    Private dbOffset As Long = 0
    Public Sub New(connection As SqlConnection, transaction As SqlTransaction, tableName As String, blobColumn As String,
                   bufferLen As Integer, pointer As Byte())
        Me.transaction = transaction
        Me.connection = connection

        cmdAppendChunk =
            New SqlCommand(
                $"UPDATETEXT [{tableName}].[{blobColumn}] @Pointer @Offset 0 @Bytes", connection, transaction)

        Dim ptrParm As SqlParameter = cmdAppendChunk.Parameters.Add(
            "@Pointer", SqlDbType.Binary, 16)
        ptrParm.Value = pointer

        blobParm = cmdAppendChunk.Parameters.Add("@Bytes", SqlDbType.Image, bufferLen)
        offsetParm = cmdAppendChunk.Parameters.Add("@Offset", SqlDbType.Int)
        offsetParm.Value = 0
    End Sub



    Public Overrides Sub Write(sourceBuffer As Byte(), offset As Integer, count As Integer)
        If IsNothing(sourceBuffer) Then Throw New ArgumentNullException("sourceBuffer")
        If (offset < 0 OrElse offset >= sourceBuffer.Length) Then Throw New ArgumentOutOfRangeException("offset")
        If (count < 0 OrElse offset + count > sourceBuffer.Length) Then Throw New ArgumentOutOfRangeException("count")
        If (count = 0) Then Return
        blobParm.Value = sourceBuffer
        cmdAppendChunk.ExecuteNonQuery()
        dbOffset += count
        offsetParm.Value = dbOffset

    End Sub

    Public Overrides ReadOnly Property CanWrite As Boolean = True

End Class

''' <summary>
''' Pseudo-Stream class to abstract the SQL Update .Write Syntax
''' This class is necessary because .Net4 does not support SqlStreams
''' see: https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/sqlclient-streaming-support
''' This ONLY WORKS if the Data Types are varchar(max), nvarchar(max), or varbinary(max)
''' image (old) data type doesn't work!!!!
''' </summary>
''' <remarks>Not used. Candidate for removal.</remarks>
Friend Class SqlBlobStream
    Inherits AbstractDBStream

    Private ReadOnly cmdAppendChunk As SqlCommand
    Private ReadOnly cmdFirstChunk As SqlCommand
    Private ReadOnly connection As SqlConnection
    Private ReadOnly transaction As SqlTransaction
    Private ReadOnly paramChunk As SqlParameter
    Private offset As Long

    Public Sub New(connection As SqlConnection, transaction As SqlTransaction, tableName As String, blobColumn As String,
                   keyColumn As String, keyValue As Long)
        Me.transaction = transaction
        Me.connection = connection
        cmdFirstChunk =
            New SqlCommand(
                $"
UPDATE [{tableName}]
    SET [{blobColumn}] = @firstChunk
    WHERE [{keyColumn}] = @key",
                connection, transaction)
        cmdFirstChunk.Parameters.AddWithValue("@key", keyValue)

        cmdAppendChunk =
            New SqlCommand(
                $"
UPDATE [{tableName}]
    SET [{blobColumn}].WRITE(@chunk, NULL, NULL)
    WHERE [{keyColumn _
                              }] = @key", connection, transaction)
        cmdAppendChunk.Parameters.AddWithValue("@key", keyValue)
        paramChunk = New SqlParameter("@chunk", SqlDbType.VarBinary, -1)
        cmdAppendChunk.Parameters.Add(paramChunk)
    End Sub



    Public Overrides Sub Write(buffer As Byte(), index As Integer, count As Integer)
        Dim bytesToWrite As Byte() = buffer

        If index <> 0 OrElse count <> buffer.Length Then
            bytesToWrite = New MemoryStream(buffer, index, count).ToArray()
        End If

        If offset = 0 Then
            cmdFirstChunk.Parameters.AddWithValue("@firstChunk", bytesToWrite)
            cmdFirstChunk.ExecuteNonQuery()
            offset = count
        Else
            paramChunk.Value = bytesToWrite
            cmdAppendChunk.ExecuteNonQuery()
            offset += count
        End If
    End Sub

    Public Overrides ReadOnly Property CanWrite As Boolean = True

End Class

''' <summary>
''' Same purpose as <see cref="SqlBlobStream"/>
''' </summary>
''' <remarks>Not used. Candidate for removal.</remarks>
Friend Class OleDbBlobStream
    Inherits AbstractDBStream

    Private ReadOnly cmdAppendChunk As OleDbCommand
    Private ReadOnly cmdFirstChunk As OleDbCommand
    Private ReadOnly connection As OleDbConnection
    Private ReadOnly transaction As OleDbTransaction
    Private ReadOnly paramChunk As OleDbParameter
    Private offset As Long

    Public Sub New(connection As OleDbConnection, transaction As OleDbTransaction, tableName As String,
                   blobColumn As String, keyColumn As String, keyValue As Long)
        Me.transaction = transaction
        Me.connection = connection
        cmdFirstChunk =
            New OleDbCommand(
                $"
UPDATE [{tableName}]
    SET [{blobColumn}] = @firstChunk
    WHERE [{keyColumn}] = @key",
                connection, transaction)
        cmdFirstChunk.Parameters.AddWithValue("@key", keyValue)
        cmdAppendChunk =
            New OleDbCommand(
                $"
UPDATE [{tableName}]
    SET [{blobColumn}].WRITE(@chunk, NULL, NULL)
    WHERE [{keyColumn _
                                }] = @key", connection, transaction)
        cmdAppendChunk.Parameters.AddWithValue("@key", keyValue)
        paramChunk = New OleDbParameter("@chunk", OleDbType.VarBinary, -1)
        cmdAppendChunk.Parameters.Add(paramChunk)
    End Sub



    Public Overrides Sub Write(buffer As Byte(), index As Integer, count As Integer)
        Dim bytesToWrite As Byte() = buffer

        If index <> 0 OrElse count <> buffer.Length Then
            bytesToWrite = New MemoryStream(buffer, index, count).ToArray()
        End If

        If offset = 0 Then
            cmdFirstChunk.Parameters.AddWithValue("@firstChunk", bytesToWrite)
            cmdFirstChunk.ExecuteNonQuery()
            offset = count
        Else
            paramChunk.Value = bytesToWrite
            cmdAppendChunk.ExecuteNonQuery()
            offset += count
        End If
    End Sub

    Public Overrides ReadOnly Property CanWrite As Boolean = True

End Class

Friend Class SqliteBlobStream
    Inherits AbstractDBStream

    Private dbOffset As Integer = 0
    Private ReadOnly blob As SQLiteBlob
    Public Sub New(connection As SQLiteConnection, tableName As String, blobColumn As String,
                   bufferLen As Integer, rowId As Integer)
        blob = SQLiteBlob.Create(connection, "main", tableName, blobColumn, rowId, False)

    End Sub


    Public Overrides Sub Write(sourceBuffer As Byte(), offset As Integer, count As Integer)
        If IsNothing(sourceBuffer) Then Throw New ArgumentNullException("sourceBuffer")
        If (offset < 0 OrElse offset >= sourceBuffer.Length) Then Throw New ArgumentOutOfRangeException("offset")
        If (count < 0 OrElse offset + count > sourceBuffer.Length) Then Throw New ArgumentOutOfRangeException("count")
        If (count = 0) Then Return
        blob.Write(sourceBuffer, count, dbOffset)
        dbOffset += count
    End Sub

    Public Overrides ReadOnly Property CanWrite As Boolean = True


    Public Overrides Sub Close()

        Try
            blob?.Dispose()
        Finally
            MyBase.Close()
        End Try
    End Sub

End Class


''' <summary>
''' Writeable stream for using a separate thread in a producer/consumer scenario.
''' This is useful if/when there is an I/O bottleneck, etc...
''' </summary>
Friend NotInheritable Class TransferStream
    Inherits AbstractDBStream
    ''' <summary>
    ''' The underlying stream to target.
    ''' </summary>
    Private _writeableStream As Stream
    ''' <summary>
    ''' The collection of chunks to be written
    ''' </summary>
    Private _chunks As BlockingCollection(Of Byte())
    ''' <summary>
    ''' The Task to use for background writing
    ''' </summary>
    Private _processingTask As Task
    Public Sub New(ByVal writeableStream As Stream)

        If Not writeableStream?.CanWrite Then Throw New ArgumentException("Target stream is NULL or not writeable.")
        _writeableStream = writeableStream
        _chunks = New BlockingCollection(Of Byte())()
        'Spin the worker thread
        _processingTask = Task.Factory.StartNew(Sub()
                                                    For Each chunk In _chunks.GetConsumingEnumerable()
                                                        _writeableStream.Write(chunk, 0, chunk.Length)
                                                    Next
                                                End Sub, TaskCreationOptions.LongRunning)
    End Sub
    Public Overrides ReadOnly Property CanWrite() As Boolean
        Get
            Return True
        End Get
    End Property


    Public Overrides Sub Write(
    ByVal sourceBuffer() As Byte, ByVal offset As Integer,
    ByVal count As Integer)
        If IsNothing(sourceBuffer) Then Throw New ArgumentNullException("sourceBuffer")
        If (offset < 0 OrElse offset >= sourceBuffer.Length) Then Throw New ArgumentOutOfRangeException("offset")
        If (count < 0 OrElse offset + count > sourceBuffer.Length) Then Throw New ArgumentOutOfRangeException("count")
        If (count = 0) Then Return

        Dim chunk = New Byte(count - 1) {}
        Buffer.BlockCopy(sourceBuffer, offset, chunk, 0, count)
        ' add Work
        _chunks.Add(chunk)
    End Sub
    Public Overrides Sub Close()
        _chunks.CompleteAdding()
        Try
            _processingTask.Wait()
        Finally
            MyBase.Close()
        End Try
    End Sub


End Class

Friend MustInherit Class AbstractDBStream
    Inherits Stream

    Public Overrides Sub Flush()
    End Sub

    Public Overrides ReadOnly Property CanRead As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property CanSeek As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property Length As Long
        Get
            Throw New NotSupportedException()
        End Get
    End Property
    Public Overrides Property Position As Long
        Get
            Throw New NotSupportedException()
        End Get
        Set(value As Long)
            Throw New NotSupportedException()
        End Set
    End Property

    Public Overrides Function Seek(offset As Long, origin As SeekOrigin) As Long
        Throw New NotSupportedException
    End Function

    Public Overrides Sub SetLength(value As Long)
        Throw New NotSupportedException
    End Sub

    Public Overrides Function Read(buffer As Byte(), offset As Integer, count As Integer) As Integer
        Throw New NotSupportedException
    End Function

End Class


