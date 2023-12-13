Option Explicit On
Option Compare Text
Option Infer On
Option Strict On


Imports System.IO
Imports System.IO.Compression
Imports System.Text
Namespace MasterInterface
    ''' <summary>
    '''     Copied from CBLOB of UDBSBLOBExtractor
    '''     Modified to use Streams all the way
    '''     to avoid OOM exceptions
    ''' </summary>
    Friend Module ZLibCompression
        ''' <remarks>Not used.</remarks>
        Public Sub CompressFile(FileName As String, CompressedFileName As String, Optional UseLegacyFormat As Boolean = False)

            If Not (File.Exists(FileName)) Then Throw New Exception("Missing file: " & FileName & ".")
            Using FS = New StreamReader(FileName)
                FS.BaseStream.Position = 0
                Using CFS As New FileStream(CompressedFileName, FileMode.Create)
                    CompressStream(FS.BaseStream, UseLegacyFormat, CFS)
                End Using
            End Using
        End Sub

        ''' <remarks>Not used.</remarks>
        Public Sub DecompressFile(CompressedFileName As String, DeCompressedFileName As String)
            Using FS = New StreamReader(CompressedFileName)
                FS.BaseStream.Position = 0
                Using CFS As New FileStream(DeCompressedFileName, FileMode.Create)
                    DeCompressStream(FS.BaseStream, CFS)
                End Using
            End Using

        End Sub

        ''' <remarks>Not used.</remarks>
        Public Function DecompressData(CompressedData() As Byte) As Byte()
            Using CompressedFileStream As New MemoryStream(CompressedData)
                Dim UnCompressedStream As New MemoryStream()
                DeCompressStream(CompressedFileStream, UnCompressedStream)
                Return UnCompressedStream.ToArray
            End Using
        End Function

        ''' <param name="deCompressedStream">
        ''' The output stream where to send the decompressed data.
        ''' </param>
        Public Sub DeCompressStream(Of TOut As {Stream, IDisposable})(compressedStreamSource As Stream, deCompressedStream As TOut)
            Dim B1 = compressedStreamSource.ReadByte() _
            'this is a weird thing - the decompression needs to ignore the first two bytes of an old compressed file
            Dim B2 = compressedStreamSource.ReadByte()

            If Not (B1 = 120 And B2 = 156) Then
                'Throw New Exception(String.Format("ZLib decompression detected wrong compression format. First two bytes should be (120, 156) but instead found ({0}, {1})", B1, B2))
                'try and go ahead - assuming new format
                compressedStreamSource.Position = 0
            End If

            Using tOutWrite = New TransferStream(deCompressedStream)
                Using deflateStream = New DeflateStream(compressedStreamSource, CompressionMode.Decompress)
                    deflateStream.CopyTo(tOutWrite)
                End Using
            End Using
        End Sub

        ''' <param name="compressedStream">
        ''' The stream where to send the compressed data.
        ''' TODO: This should not be 'by reference'.
        ''' </param>
        Public Sub CompressStream(Of TOut As {Stream, IDisposable})(unCompressedStreamSource As Stream, useLegacyFormat As Boolean, compressedStream As TOut)
            If useLegacyFormat Then
                'first two bytes should be (0x78, 0x9C) 
                compressedStream.WriteByte(120)
                compressedStream.WriteByte(156)
            End If
            ' Disk I/O is slower than compression
            Using threadTransfer = New TransferStream(compressedStream)
                Using deflateStream = New DeflateStream(threadTransfer, CompressionMode.Compress, True)
                    Using tComp = New TransferStream(deflateStream)
                        unCompressedStreamSource.CopyTo(tComp)
                    End Using
                End Using
            End Using

            If useLegacyFormat Then
                'Add Adler32 checksum at end
                Dim CSBytes As Byte() = ComputeAdler32(unCompressedStreamSource)
                compressedStream.WriteByte(CSBytes(3))
                compressedStream.WriteByte(CSBytes(2))
                compressedStream.WriteByte(CSBytes(1))
                compressedStream.WriteByte(CSBytes(0))
            End If
        End Sub




        Private Function ComputeAdler32(DataStream As Stream) As Byte()

            Const Modulus As Long = 65521
            Dim a As Long = 1
            Dim b As Long = 0

            DataStream.Position = 0

            For i = 0 To DataStream.Length - 1
                a = (a + DataStream.ReadByte) Mod Modulus
                b = (b + a) Mod Modulus
            Next i
            Dim CheckSum As Long = (b * 65536) + a
            Return BitConverter.GetBytes(CheckSum)
        End Function
    End Module

    ''' <summary>
    ''' A stream implementation with a Temporary file backing store
    ''' which is deleted upon dispose
    ''' </summary>
    Friend NotInheritable Class TempStream
        Inherits FileStream
        Sub New()
            MyBase.New(path:=Path.GetTempFileName(),
                       mode:=FileMode.Open,
                       access:=FileAccess.ReadWrite,
                       share:=FileShare.None,
                       bufferSize:=4096, ' default CTOR
                       options:=FileOptions.DeleteOnClose Or FileOptions.Asynchronous Or FileOptions.RandomAccess)
        End Sub
    End Class


    Public Class CBLOB
        Implements IDisposable

        ''' <summary>
        ''' How many attempts we try to use in-memory buffer stream before
        ''' falling back to a file-base temporary stream.
        ''' Memory buffer is faster, but we may run out of memory if we are
        ''' trying to download/upload a very large file.
        ''' If we fail with an out-of-memory exception, we retry using a
        ''' temporary file stream.
        ''' The default behavior is to retry after the first error.
        ''' We expose this property internally in order for the unit tests
        ''' to use the temporary stream all the time (by setting this property
        ''' to zero) and achieve higher code coverage.
        ''' </summary>
        ''' <returns></returns>
        Friend Shared Property TemporaryStreamFallbackRetryCount As Integer = 1

        ''' <summary>
        ''' Special type representing files.
        ''' </summary>
        Public Const FileType As Integer = 1000

        ''' <summary>
        ''' File size is stored in a 32-bits integer in the UDBS DB schema.
        ''' This is a property so that unit test could lower that number and validate that
        ''' the size validation is performed.
        ''' </summary>
        Public Property MaxFileSize As Integer = Integer.MaxValue

        ''' <summary>
        ''' Retrieve the list of attachments associated with a given result.
        ''' </summary>
        ''' <param name="ProcessName">The name of the process (failure/kitting/testdata/qc/rework/wip)</param>
        ''' <param name="LinkToTable">The name of the table being linked to. This is usually derived from the process name.</param>
        ''' <param name="LinkToID">The ID of the result being queried.</param>
        ''' <param name="ArrayNames">(Output) The name of the attachments.</param>
        ''' <param name="ArrayDataType">(Output) The types of the attachments (see Microsoft.VisualBasic.VariantType)</param>
        ''' <returns>The outcome of the operation.</returns>
        Public Function GetAttachmentList(ProcessName As String,
                                          LinkToTable As String,
                                          LinkToID As Long,
                                          ByRef ArrayNames() As String,
                                          ByRef ArrayDataType() As Integer) _
            As ReturnCodes
            ' Function returns a list of the arrays available for the specified item
            Dim sqlQuery As String
            Dim BLOBTable As String
            Dim rsTemp As New DataTable
            Dim Counter As Integer

            GetAttachmentList = ReturnCodes.UDBS_ERROR
            Try
                OpenNetworkDB(120)

                BLOBTable = LCase(Trim(ProcessName)) & "_blob"
                sqlQuery = "SELECT blob_array_name, blob_datatype FROM " & BLOBTable & " with(nolock) " &
                           "WHERE blob_ref_item_table = '" & LinkToTable & "' " &
                           "AND blob_ref_item_id = " & CStr(LinkToID) & " " &
                           "ORDER BY blob_isheader, blob_id"

                OpenNetworkRecordSet(rsTemp, sqlQuery)

                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                    ' There are blobs available for this item
                    ReDim ArrayNames(rsTemp.Rows.Count - 1)
                    ReDim ArrayDataType(rsTemp.Rows.Count - 1)
                    Counter = 0
                    For Each dr As DataRow In rsTemp.Rows
                        ArrayNames(Counter) = KillNull(dr("blob_array_name"))
                        If Not IsDBNull(dr("blob_datatype")) Then
                            ArrayDataType(Counter) = KillNullInteger(dr("blob_datatype"))
                        End If
                        Counter = Counter + 1

                    Next
                    GetAttachmentList = ReturnCodes.UDBS_OP_SUCCESS
                Else
                    LogError(New Exception($"No attachment found for Process: {ProcessName} LinkToTable: {LinkToTable} LinkToID: {LinkToID}."))
                    Array.Clear(ArrayNames, 0, ArrayNames.Length)
                    Array.Clear(ArrayDataType, 0, ArrayNames.Length)

                    ' Although some could argue that 'no attachment' is not an error,
                    ' this behavior has been exposed for a long while, and client code may
                    ' be expecting this return code, so we will not be changing this.
                    GetAttachmentList = ReturnCodes.UDBS_ERROR
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try
        End Function

        Private Function ValidateFileSize(fileName As String) As Boolean
            Dim fncName = "ValidateFileSize"
            Dim fileSize = New FileInfo(fileName).Length
            If (fileSize > MaxFileSize) Then
                LogError(New Exception($"File '{fileName}' is too large ({fileSize} b, maximum is {MaxFileSize} b)."))
                Return False
            ElseIf fileSize = 0 Then
                LogError(New Exception($"File is empty: {fileName}"))
                Return False
            Else
                Return True
            End If
        End Function

        ''' <summary>
        ''' Restore a file archived to Amazon S3.
        ''' </summary>
        ''' <param name="blobId">The blob ID.</param>
        ''' <returns>Whether or not the operation succeded.</returns>
        Friend Function RestoreFileIfArchived(blobId As Long) As Boolean
            Try
                OpenNetworkDB(120)
                Dim sqlStr = "exec stpr_BLOB_s3_restore_local " & blobId

                Dim rsTemp As New DataTable()
                OpenNetworkRecordSet(rsTemp, sqlStr)

                If rsTemp IsNot Nothing AndAlso rsTemp.Rows.Count > 0 Then
                    Dim tmpMsg = KillNull(rsTemp(0)(0))
                    If Left(tmpMsg, 1) = "1" Then 'should be something like "1:Success..."
                        Return True
                    Else
                        'should be something like "0:Some kind of error occurred..."
                        Throw New UDBSException("BLOB is no longer in the database. Tried to recall it from cloud storage but encountered the following error: " & tmpMsg)
                    End If
                Else
                    Throw New Exception("BLOB is no longer in the database. Failed to recall it from cloud storage")
                End If
            Catch ex As Exception
                logger.Error(ex)
                Return False
            End Try
        End Function

        ' TODO: The following method does too much. The parameters for handling of arrays & file
        '       are 'either or', which is a red-sign of a method with too many responsabilities.
        '       Split this method into two.

        ''' <summary>
        ''' Attaches a BLOB (from a data array or a file) to a result.
        ''' </summary>
        ''' <param name="ProcessName">The name of the process (failure/kitting/testdata/qc/rework/wip)</param>
        ''' <param name="LinkToTable">The name of the table being linked to. This is usually derived from the process name.</param>
        ''' <param name="LinkToID">The ID of the result to link this BLOB to.</param>
        ''' <param name="ArrayName">The name of the attachment.</param>
        ''' <param name="DataGroupName">The group name of the attachment.</param>
        ''' <param name="IsHeader">Whether or not this is flagged as a header (???)</param>
        ''' <param name="ArrayData">If storing from an array, this is the data being stored. 
        ''' This is ignored if the FileName parameter is not empty.</param>
        ''' <param name="FileName">If storing from a file, this is the path to the file being stored.</param>
        ''' <param name="lowBound">If storing from an array, this is the lower boundary of the data being saved.
        ''' Ignored when storing from a file.</param>
        ''' <param name="upBound">If storing from an array, this is the upper boundary of the data being saved.
        ''' Ignored when storing from a file.</param>
        ''' <returns>The outcome of the operation.</returns>
        Public Function StoreBLOB(ProcessName As String,
                                  LinkToTable As String,
                                  LinkToID As Long,
                                  ArrayName As String,
                                  DataGroupName As String,
                                  IsHeader As Boolean,
                                  ArrayData As Array,
                                  FileName As String,
                                  lowBound As Integer,
                                  upBound As Integer) _
                As ReturnCodes

            Dim sqlQuery As String
            Dim rsTemp As New DataTable
            Dim BLOBTable As String
            Dim TempFileName = ""
            Dim arrayDataType As Integer
            Dim payload As Stream = Nothing
            Dim UnCompressedFileSize As Integer
            Dim blobId As Long = Long.MinValue ' It is a stop gap solution to use negative primary key values
            Dim blobExist As Boolean = False

            Dim RetCode = ReturnCodes.UDBS_ERROR
            Try
                If String.IsNullOrEmpty(FileName) Then
                    ' this is an array, convert the array into binary file
                    If ArrayData Is Nothing Then
                        ' this is not an array
                        LogError(New Exception("Invalid array."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If
                    arrayDataType = DetermineArrayTypeCode(ArrayData)

                    ' Validate and adjust the lowwer and upper boundaries.
                    If upBound < 0 Then
                        ' No upper bound provided (-1)
                        upBound = UBound(ArrayData)
                    End If
                    If lowBound > upBound Then
                        ' Lower and upper boundary were inverted.
                        ' Swap them.
                        Dim tmpBound = lowBound
                        lowBound = upBound
                        upBound = tmpBound
                    End If

                    If lowBound < LBound(ArrayData) Then
                        ' Lower boundary too low. Use the actual lower boundary.
                        lowBound = LBound(ArrayData)
                    End If
                    If upBound > UBound(ArrayData) Then
                        ' Upper boundary too high. Use the actual upper boundary.
                        upBound = UBound(ArrayData)
                    End If

                    TempFileName = Path.GetTempFileName
                    SaveArrayToFile(ArrayData, TempFileName, lowBound, upBound)
                    FileName = TempFileName
                Else
                    arrayDataType = FileType
                End If

                BLOBTable = LCase(Trim(ProcessName)) & "_blob"
                'fileSpec = fileSpec 

                OpenNetworkDB(120)

                sqlQuery = "SELECT * FROM " & BLOBTable & " with(nolock) " &
                           "WHERE blob_ref_item_table = '" & LinkToTable & "' " &
                           "AND blob_ref_item_id = " & LinkToID & " " &
                           "AND blob_array_name = '" & ArrayName & "' "
                OpenNetworkRecordSet(rsTemp, sqlQuery)

                'If rsTemp.RecordCount > 1 Then
                '    ' There is a problem.. There should only be 1 instance if ANY...
                '    DebugMessage fncName, "More than one BLOB with arrayName=" & ArrayName & "."
                '    Exit Function
                'ElseIf rsTemp.RecordCount < 1 Then
                '    ' Add a new record
                '    rsTemp.AddNew
                'End If

                Dim columnNames = New List(Of String) From {"blob_ref_item_table", "blob_ref_item_id",
                                                "blob_array_name", "blob_datagroup_name", "blob_isheader",
                                                "blob_datatype", "blob_elements", "blob_blob", "blob_origsize"}

                ' Populate the fields of the new record.
                Dim columnValues = New List(Of Object) From {
                        LinkToTable, LinkToID,
                        ArrayName, DataGroupName, IsHeader, arrayDataType,
                        If(String.IsNullOrEmpty(TempFileName), 0, (upBound - lowBound + 1)),
                        DBNull.Value, 0}

                ' Update
                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                    columnNames.Insert(0, "blob_id")
                    blobId = Convert.ToInt64(rsTemp(0)("blob_id"))
                    columnValues.Insert(0, blobId)
                    blobExist = True
                End If

                If String.IsNullOrEmpty(FileName) OrElse Not (File.Exists(FileName)) Then
                    LogError(New Exception($"Missing file: {FileName} when attempting to store blob."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                If Not ValidateFileSize(FileName) Then
                    ' Error was already logged.
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                UnCompressedFileSize = CInt(New FileInfo(FileName).Length)

                Dim retryCount As Integer = 0

                Using FS = New StreamReader(FileName)
                    While retryCount <= 1
                        FS.BaseStream.Position = 0
                        Try
                            If retryCount < TemporaryStreamFallbackRetryCount Then
                                'Memory-based, faster
                                payload = New MemoryStream()
                            Else
                                ' File-based stream, slower
                                payload = New TempStream()
                            End If
                            'TODO: Perhaps if the file is REALLY REALLY BIG
                            'Don't save it as a blob in SQL Server but rather just keep a link of where
                            'the zipped file is, for upload later. This would violate the ACID principles of a database
                            'dependant system....
                            ' Or use the newer SQL data type called FileStream???

                            CompressStream(FS.BaseStream, True, payload)

                            payload.Position = 0
                            columnValues(columnValues.Count - 2) = payload
                            columnValues(columnValues.Count - 1) = UnCompressedFileSize

                            If Not blobExist Then
                                ' Insert 
                                InsertNetworkRecord(columnNames.ToArray(), columnValues.ToArray(), BLOBTable)
                            Else
                                UpdateNetworkRecord({columnNames(0)}, columnNames.ToArray(), columnValues.ToArray(), BLOBTable)
                            End If
                            Return ReturnCodes.UDBS_OP_SUCCESS
                        Catch oomex As OutOfMemoryException
                            logger.Warn(oomex, $"Problem {If(Not blobExist, "Inserting", "Updating")} {BLOBTable}. Retrying...")
                            retryCount += 1
                            GC.Collect()
                        Catch ex As Exception
                            logger.Error(ex, $"Problem {If(Not blobExist, "Inserting", "Updating")} {BLOBTable}.")
                            Throw ' some other error, don't bother retrying
                        End Try

                    End While

                End Using

                Return ReturnCodes.UDBS_OP_FAIL

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_OP_FAIL
            Finally
                payload?.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' Determine the code representing the type of the elements of an array.
        ''' The UDBS database was originally filled by an VB6 application, and the 'type'
        ''' column was populated with the value of the corresponding 'Variant Type'.
        ''' </summary>
        ''' <param name="anArray"></param>
        ''' <returns></returns>
        Public Shared Function DetermineArrayTypeCode(anArray As Array) As Integer
            Select Case anArray.GetType().GetElementType()
                Case GetType(Byte)
                    Return VariantType.Byte
                Case GetType(Short)
                    Return VariantType.Short
                Case GetType(Integer)
                    Return VariantType.Integer
                Case GetType(Single)
                    Return VariantType.Single
                Case GetType(Double)
                    Return VariantType.Double
                Case GetType(DateTime)
                    Return VariantType.Date
                Case GetType(String)
                    Return VariantType.String
                Case Else
                    ' Default to the old VB Variant code.
                    ' Not sure this works... This seems to return 'VariantType.Object' all the time.
                    Return VarType(anArray.GetType())
            End Select
        End Function

        Public Function StoreBLOB_Local(ProcessID As Integer,
                                        Itemname As String,
                                        ArrayName As String,
                                        DataGroupName As String,
                                        IsHeader As Boolean,
                                        ArrayData As Array,
                                        FileName As String,
                                        lowBound As Integer,
                                        upBound As Integer) As ReturnCodes
            ' Function store BLOB to local in case of network failure
            ' fileSpec is the file specification (path & name) of the image file to be saved to UDBS
            ' DataType: Integer = 2, Long = 3, Single = 4, Double = 5, Date = 7,
            '           String = 8, Byte = 17, Image(File) = 1000
            ' In case of Image(File) data, NumElements=0
            ' NumElements and DataType are determined from the supplied array

            Dim sqlQuery As String
            Dim rsTemp As New DataTable
            Dim TempFileName = ""


            Dim UnCompressedFileSize As Integer

            Dim RetCode = ReturnCodes.UDBS_OP_INC
            Dim payload As Stream = Nothing
            Dim blobId As Long = Long.MinValue
            Dim blobExist As Boolean = False
            Try
                If FileName = "" Then
                    ' this is an array, convert the array into binary file
                    If ArrayData Is Nothing Then
                        ' this is not an array
                        LogError(New Exception("Invalid array."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If
                    ' verify lowBound and upBound
                    If lowBound > upBound Then
                        Dim tmpBound = lowBound
                        lowBound = upBound
                        upBound = tmpBound
                    End If
                    If lowBound < LBound(ArrayData) Then
                        lowBound = LBound(ArrayData)
                    End If
                    If upBound > UBound(ArrayData) Then
                        upBound = UBound(ArrayData)
                    End If
                    TempFileName = Path.GetTempFileName
                    SaveArrayToFile(ArrayData, TempFileName, lowBound, upBound)
                    FileName = TempFileName
                End If

                Dim columnNames = New List(Of String) From {"blob_ref_item_table", "blob_ref_item_id",
                                                "blob_array_name", "blob_datagroup_name", "blob_isheader",
                                                "blob_datatype", "blob_elements", "blob_blob", "blob_origsize"}

                ' Populate the fields of the new record
                Dim columnValues = New List(Of Object) From {Itemname, ProcessID,
                                                 ArrayName, DataGroupName, IsHeader,
                                                 If(String.IsNullOrEmpty(TempFileName), 1000, DetermineArrayTypeCode(ArrayData)),
                                                 If(String.IsNullOrEmpty(TempFileName), 0, (upBound - lowBound + 1)),
                                                 DBNull.Value, 0}

                ' Use blob_ref_item_table & blob_ref_item_id to store itemname & processID on the local mdb temporary
                sqlQuery = "SELECT * FROM testdata_blob " &
                           "WHERE blob_ref_item_table = '" & Itemname & "' " &
                           "AND blob_ref_item_id = " & ProcessID & " " &
                           "AND blob_array_name = '" & ArrayName & "' "
                If OpenLocalRecordSet(rsTemp, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    LogError(New Exception($"Error querying for array ""{ArrayName}"" of item ""{Itemname}"" for process {ProcessID}."))
                    Return ReturnCodes.UDBS_OP_FAIL
                ElseIf (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                    columnNames.Insert(0, "blob_id")
                    blobId = Convert.ToInt64(rsTemp(0)("blob_id"))
                    columnValues.Insert(0, blobId)
                    blobExist = True
                End If

                If String.IsNullOrEmpty(FileName) OrElse Not (File.Exists(FileName)) Then
                    LogError(New Exception($"Missing file: {FileName}"))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                If Not ValidateFileSize(FileName) Then
                    ' Error was already logged.
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                UnCompressedFileSize = CInt(New FileInfo(FileName).Length)

                Dim retryCount As Integer = 0

                Using FS = New StreamReader(FileName)

                    While retryCount <= 1

                        FS.BaseStream.Position = 0
                        Try
                            If retryCount < TemporaryStreamFallbackRetryCount Then
                                'Memory-based, faster
                                payload = New MemoryStream()
                            Else
                                ' File-based stream, slower
                                payload = New TempStream()
                            End If
                            'TODO: Perhaps if the file is REALLY BIG
                            'Don't save it as a blob in SQLite/MS-ACCESS but rather just keep a link of where
                            'the zipped file is, for upload later. This would violate the ACID principles of a database
                            'dependant system....

                            CompressStream(FS.BaseStream, True, payload)

                            payload.Position = 0
                            columnValues(columnValues.Count - 2) = payload
                            columnValues(columnValues.Count - 1) = UnCompressedFileSize

                            ' Insert
                            If Not blobExist Then
                                InsertLocalRecord(columnNames.ToArray(), columnValues.ToArray(), "testdata_blob")
                            Else
                                UpdateLocalRecord({columnNames(0)}, columnNames.ToArray(), columnValues.ToArray(), "testdata_blob")
                            End If

                            Return ReturnCodes.UDBS_OP_SUCCESS
                        Catch oomex As OutOfMemoryException
                            logger.Warn(oomex, $"Problem {If(Not blobExist, "Inserting", "Updating")} testdata_blob (LocalDB). Retrying...")
                            retryCount += 1
                            GC.Collect()
                        Catch ex As Exception
                            logger.Error(ex, $"Problem {If(Not blobExist, "Inserting", "Updating")} testdata_blob (LocalDB).")
                            Throw ' some other error, don't bother retrying
                        End Try
                    End While

                End Using

                Return ReturnCodes.UDBS_OP_FAIL

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_OP_FAIL
            Finally
                payload?.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' Deletes the BLOB for the local DB
        ''' </summary>
        ''' <param name="ProcessID"></param>
        ''' <param name="Itemname"></param>
        ''' <param name="ArrayName"></param>
        ''' <returns><see cref="ReturnCodes"/></returns>
        Public Function RemoveLocalBLOB(ProcessID As Integer,
                                        Itemname As String,
                                        ArrayName As String) As ReturnCodes
            Try
                Using transaction = BeginLocalTransaction()
                    Dim sqlQuery As String = "DELETE FROM testdata_blob " &
           "WHERE blob_ref_item_table = '" & Itemname & "' " &
           "AND blob_ref_item_id = " & ProcessID & " " &
           "AND blob_array_name = '" & ArrayName & "' "
                    ExecuteLocalQuery(sqlQuery, transaction)
                End Using

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Determine whether or not a BLOB has been archived.
        ''' </summary>
        ''' <param name="processName">The process name (test_data, wip, etc.)</param>
        ''' <param name="itemId">The item this BLOB is linked to.</param>
        ''' <param name="arrayName">The name of the BLOB (file name, array name).</param>
        ''' <param name="blobId">(Out) The DB key to the BLOB.</param>
        ''' <returns></returns>
        Friend Function HasBeenArchived(
                processName As String,
                itemId As Long,
                arrayName As String,
                ByRef blobId As Long) As Boolean
            Dim BLOBTable As String = LCase(Trim(processName)) & "_blob"

            Dim sqlQuery = "SELECT blob_id AS id, datalength(blob_blob) AS length from " & BLOBTable & " with(nolock) " &
                           " WHERE blob_ref_item_id = " & CStr(itemId) &
                           " AND blob_array_name = '" & arrayName & "' " &
                           " ORDER BY blob_id DESC"

            Dim dataSetTmp As New DataTable()
            Try
                OpenNetworkRecordSet(dataSetTmp, sqlQuery)

                If dataSetTmp.Rows.Count > 0 Then
                    blobId = KillNullLong(dataSetTmp.Rows(0)("id"))
                    Dim actualSize As Integer = KillNullInteger(dataSetTmp.Rows(0)("length"))

                    If actualSize = 0 Then
                        Return True
                    End If
                End If

                Return False
            Finally
                dataSetTmp?.Dispose()
            End Try
        End Function

        ' TODO: There is a problem with this interface: When querying a file, you need to specify the
        '       file name. But the type is an output parameter, so you have to know its value before
        '       being able to make a valid query.
        '       Also, the fact that the ArrayData and FileName attributes are 'one or the other'
        '       tells that this method has too many responsabilities. It's begging to be split into
        '       two methods.
        '       Also, why couldn't we load a file in memory? Or save an array to a file?

        ''' <summary>
        ''' Retrieve an attachment.
        ''' </summary>
        ''' <param name="ProcessName">The name of the process (failure/kitting/testdata/qc/rework/wip)</param>
        ''' <param name="LinkToTable">The name of the table being linked to. This is usually derived from the process name.</param>
        ''' <param name="LinkToID">The ID of the result being queried.</param>
        ''' <param name="ArrayName">The name of the attachment.</param>
        ''' <param name="DataGroupName">(Output) The group name associated to this attachment.</param>
        ''' <param name="NumElements">(Output) The number of elements contained in this attachment.</param>
        ''' <param name="dataType">(Output) The data type of the attachment.</param>
        ''' <param name="IsHeader">(Output) Whether or not this is flagged as a header (???)</param>
        ''' <param name="ArrayData">(Output) The content of the attachment. Unused if retrieving a file attachment.</param>
        ''' <param name="FileName">If reading a file attachment, the file name where to store the data has to be specified.</param>
        ''' <returns></returns>
        Public Function GetBLOB(ProcessName As String,
                                LinkToTable As String,
                                LinkToID As Long,
                                ArrayName As String,
                                ByRef DataGroupName As String,
                                ByRef NumElements As Integer,
                                ByRef dataType As VariantType,
                                ByRef IsHeader As Boolean,
                                ByRef ArrayData As Array,
                                FileName As String) _
            As ReturnCodes

            Dim sqlQuery As String
            Dim rsTemp As DataRow = Nothing
            Dim BLOBTable As String
            Dim TempFileName = ""
            Dim payload As Stream = Nothing

            Try

                If LinkToTable = "" Or ArrayName = "" Then
                    ' Not enough information to retrieve BLOB
                    LogError(New Exception("Not enough information to retrieve BLOB."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If
                BLOBTable = LCase(Trim(ProcessName)) & "_blob"

                Dim blobId As Long = 0
                If (HasBeenArchived(ProcessName, LinkToID, ArrayName, blobId)) Then
                    RestoreFileIfArchived(blobId)
                End If

                OpenNetworkDB(120)

                sqlQuery = "SELECT * FROM " & BLOBTable & " with(nolock) " &
                           "WHERE blob_ref_item_table = '" & LinkToTable & "' " &
                           "AND blob_ref_item_id = " & CStr(LinkToID) & " " &
                           "AND blob_array_name = '" & ArrayName & "' " &
                           "ORDER BY blob_id DESC"
                Dim retryCount As Integer = 0
                While retryCount <= 1
                    Try
                        If retryCount < TemporaryStreamFallbackRetryCount Then
                            'Memory-based, faster
                            payload = New MemoryStream()
                        Else
                            ' File-based stream, slower
                            payload = New TempStream()
                        End If
                        OpenNetworkRecordSet(rsTemp, payload, sqlQuery)
                        Exit While
                    Catch oomex As OutOfMemoryException
                        retryCount += 1
                        GC.Collect()
                    Catch ex As Exception
                        Throw ' some other error, don't bother retrying
                    End Try
                End While

                If IsNothing(rsTemp) Then
                    ' no record found
                    LogError(New Exception($"No BLOB found: {ArrayName }"))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                ' Retrieve values
                ArrayName = KillNull(rsTemp("blob_array_name"))
                DataGroupName = KillNull(rsTemp("blob_datagroup_name"))
                NumElements = KillNullInteger(rsTemp("blob_elements"))
                dataType = CType(KillNull(rsTemp("blob_datatype")), VariantType)
                IsHeader = CBool(KillNull(rsTemp("blob_isheader")))

                If dataType = FileType Then
                    ' This is an image(file)
                    If NumElements <> 0 Then
                        ' NumElements must be 0 if this is a file
                        LogError(New Exception($"Wrong number of elements."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If
                    If String.IsNullOrEmpty(FileName) Then
                        ' No filespec supplied
                        LogError(New Exception("No file name supplied."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If
                Else
                    ' This is an array data
                    TempFileName = Path.GetTempFileName
                    ' If the data was passed as an input parameter, validate that
                    ' the type is valid.
                    If ArrayData IsNot Nothing AndAlso VarType(ArrayData) < vbArray Then
                        LogError(New Exception("This is not an array."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If
                    FileName = TempFileName
                End If

                'NB: Using Streams all the way to avoid OOM exceptions
                Using UnCompressedFileStream As New FileStream(FileName, FileMode.Create)
                    If IsDBNull(rsTemp("blob_origsize")) Then
                        payload.Position = 0
                        payload.CopyTo(UnCompressedFileStream)
                    Else
                        payload.Position = 0
                        DeCompressStream(payload, UnCompressedFileStream)
                    End If
                End Using

                'Debug.Print "Actual:" & actBLOBSize, "Original:" & rsTemp("blob_origsize").Value
                If dataType <> FileType Then
                    ' convert binary data into array as individual element
                    ArrayData = New Object() {}
                    Return getArrayFromFile_NewFormat(NumElements, FileName, dataType, ArrayData)
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_OP_FAIL
            Finally
                payload?.Dispose()
                If Not String.IsNullOrEmpty(TempFileName) AndAlso File.Exists(TempFileName) Then _
                    File.Delete(TempFileName)
            End Try
        End Function

        ''' <summary>
        ''' Retrieve an attachment from the local database.
        ''' </summary>
        ''' <param name="ProcessName">The name of the process (failure/kitting/testdata/qc/rework/wip)</param>
        ''' <param name="ItemName">The name of the process item.</param>
        ''' <param name="ProcessId">The ID of the process this is attached to.</param>
        ''' <param name="ArrayName">The name of the attachment.</param>
        ''' <param name="DataGroupName">(Output) The group name associated to this attachment.</param>
        ''' <param name="NumElements">(Output) The number of elements contained in this attachment.</param>
        ''' <param name="dataType">(Output) The data type of the attachment.</param>
        ''' <param name="IsHeader">(Output) Whether or not this is flagged as a header (???)</param>
        ''' <param name="ArrayData">(Output) The content of the attachment. Unused if retrieving a file attachment.</param>
        ''' <param name="FileName">If reading a file attachment, the file name where to store the data has to be specified.</param>
        ''' <returns></returns>
        Public Function GetBLOB_Local(ProcessName As String,
                                      ItemName As String,
                                      ProcessId As Long,
                                      ArrayName As String,
                                      ByRef DataGroupName As String,
                                      ByRef NumElements As Integer,
                                      ByRef dataType As VariantType,
                                      ByRef IsHeader As Boolean,
                                      ByRef ArrayData As Array,
                                      FileName As String) _
            As ReturnCodes

            Dim sqlQuery As String
            Dim rsTemp As DataRow = Nothing
            Dim BLOBTable As String
            Dim TempFileName = ""
            Dim payload As Stream = Nothing

            Try
                If ItemName = "" Or ArrayName = "" Then
                    ' Not enough information to retrieve BLOB
                    LogError(New Exception("Not enough information to retrieve BLOB."))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If
                BLOBTable = LCase(Trim(ProcessName)) & "_blob"

                sqlQuery = "SELECT * FROM " & BLOBTable & " " &
                           "WHERE blob_ref_item_table = '" & ItemName & "' " &
                           "AND blob_ref_item_id = " & CStr(ProcessId) & " " &
                           "AND blob_array_name = '" & ArrayName & "' " &
                           "ORDER BY blob_id DESC"
                Dim retryCount As Integer = 0
                While retryCount <= 1
                    Try
                        If retryCount < TemporaryStreamFallbackRetryCount Then
                            'Memory-based, faster
                            payload = New MemoryStream()
                        Else
                            ' File-based stream, slower
                            payload = New TempStream()
                        End If
                        OpenLocalRecordSet(rsTemp, payload, sqlQuery)
                        Exit While
                    Catch oomex As OutOfMemoryException
                        retryCount += 1
                        GC.Collect()
                    Catch ex As Exception
                        Throw ' some other error, don't bother retrying

                    End Try
                End While


                If IsNothing(rsTemp) Then
                    ' no record found
                    LogError(New Exception($"No BLOB found: {ArrayName }"))
                    Return ReturnCodes.UDBS_OP_FAIL
                End If

                ' Retrieve values
                ArrayName = KillNull(rsTemp("blob_array_name"))
                DataGroupName = KillNull(rsTemp("blob_datagroup_name"))
                NumElements = KillNullInteger(rsTemp("blob_elements"))
                dataType = CType(KillNull(rsTemp("blob_datatype")), VariantType)
                IsHeader = CBool(KillNull(rsTemp("blob_isheader")))

                If dataType = FileType Then
                    ' This is an image(file)
                    If NumElements <> 0 Then
                        ' NumElements must be 0 if this is a file
                        LogError(New Exception($"Wrong number of elements."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If
                    If String.IsNullOrEmpty(FileName) Then
                        ' No filespec supplied
                        LogError(New Exception("No file name supplied."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If
                Else
                    ' This is an array data
                    TempFileName = Path.GetTempFileName
                    ' If the data was passed as an input parameter, validate that
                    ' the type is valid.
                    If ArrayData IsNot Nothing AndAlso VarType(ArrayData) < vbArray Then
                        LogError(New Exception("This is not an array."))
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If
                    FileName = TempFileName
                End If

                'NB: Using Streams all the way to avoid OOM exceptions
                Using UnCompressedFileStream As New FileStream(FileName, FileMode.Create)
                    If IsDBNull(rsTemp("blob_origsize")) Then
                        payload.Position = 0
                        payload.CopyTo(UnCompressedFileStream)
                    Else
                        payload.Position = 0
                        DeCompressStream(payload, UnCompressedFileStream)
                    End If
                End Using

                payload?.Dispose()

                'Debug.Print "Actual:" & actBLOBSize, "Original:" & rsTemp("blob_origsize").Value
                If dataType <> FileType Then
                    ' convert binary data into array as individual element
                    ArrayData = New Object() {}
                    Return getArrayFromFile_NewFormat(NumElements, FileName, dataType, ArrayData)
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_OP_FAIL
            Finally
                CloseLocalDB()
                If Not String.IsNullOrEmpty(TempFileName) AndAlso File.Exists(TempFileName) Then _
                    File.Delete(TempFileName)
            End Try
        End Function

        ''' <summary>
        ''' Get the number of BLOBs attached to a given process in the local DB.
        ''' </summary>
        ''' <param name="processName">The name of the process.</param>
        ''' <param name="processId">The ID of the process.</param>
        ''' <param name="blobCount">(Output) How many BLOBs are attached to this process in the local DB.</param>
        ''' <returns>Whether or not the operation succeeded.</returns>
        Public Shared Function GetLocalBlobCount(processName As String, processId As Integer, ByRef blobCount As Integer) As ReturnCodes
            Dim results As DataTable = Nothing
            Try
                ' IMPORTANT: This may look like an error, but it is not.
                '            (i.e. We are filtering by Process ID, but we are applying the filter
                '            on a column named '...item_id'. But this is ok. The name of the local
                '            BLOB table's columns are a bit misleading. The 'blob_ref_item_id'
                '            column actually holds the 'Process ID'.)
                Dim sqlQuery = "SELECT COUNT(blob_id) AS BLOBCount " &
                                           "FROM " & processName & "_blob " &
                                           "WHERE blob_ref_item_id = " & processId
                If (OpenLocalRecordSet(results, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS) Then
                    logger.Error($"Error querying for blob count associated with process {processId}.")
                    Return ReturnCodes.UDBS_ERROR
                End If

                blobCount = KillNullInteger(results(0)(0))
                Return ReturnCodes.UDBS_OP_SUCCESS
            Finally
                results?.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' Get the number of BLOBs attached to a given process in the local DB.
        ''' </summary>
        ''' <param name="processName">The name of the process.</param>
        ''' <param name="processId">The ID of the process.</param>
        ''' <param name="blobCount">(Output) How many BLOBs are attached to this process in the local DB.</param>
        ''' <returns>Whether or not the operation succeeded.</returns>
        Public Shared Function GetLocalBlobSize(
                processName As String,
                processId As Integer,
                ByRef blobCount As Integer,
                ByRef totalCompressedSize As Long,
                ByRef totalUncompressedSize As Long) As ReturnCodes
            Dim results As DataTable = Nothing
            Try
                ' IMPORTANT: This may look like an error, but it is not.
                '            (i.e. We are filtering by Process ID, but we are applying the filter
                '            on a column named '...item_id'. But this is ok. The name of the local
                '            BLOB table's columns are a bit misleading. The 'blob_ref_item_id'
                '            column actually holds the 'Process ID'.)
                Dim sqlQuery = "SELECT blob_id, blob_array_name, length(blob_blob) as blob_compressed_size, blob_origsize " &
                                           "FROM " & processName & "_blob " &
                                           "WHERE blob_ref_item_id = " & processId
                If (OpenLocalRecordSet(results, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS) Then
                    logger.Error($"Error querying for blob count associated with process {processId}.")
                    Return ReturnCodes.UDBS_ERROR
                End If

                totalCompressedSize = 0
                totalUncompressedSize = 0
                blobCount = 0

                For Each aRow As DataRow In results.Rows
                    Dim compressedSize = KillNullLong(aRow("blob_compressed_size"))
                    Dim uncompressedSize = KillNullLong(aRow("blob_origsize"))

                    totalCompressedSize += compressedSize
                    totalUncompressedSize += uncompressedSize
                    blobCount += 1
                Next

                Return ReturnCodes.UDBS_OP_SUCCESS
            Finally
                results?.Dispose()
            End Try
        End Function

        Private Sub SaveArrayToFile(ArrayData As Array, FileName As String, lowBound As Integer, upBound As Integer)
            ' assuming upBound and lowBound have been verified by the previous fcn
            Dim i As Integer
            Dim j As Integer
            Dim numOfElements As Integer
            Dim dataType As VariantType

            If VarType(ArrayData) < vbArray Then
                ' not an array
                Throw New Exception("Tried to save array data but data type supplied is not array type.")
            End If

            If Not String.IsNullOrEmpty(FileName) AndAlso File.Exists(FileName) Then File.Delete(FileName)

            Using FS = New FileStream(FileName, FileMode.CreateNew)
                Dim BW = New BinaryWriter(FS, Encoding.Unicode)

                numOfElements = upBound - lowBound + 1
                dataType = VarType(ArrayData.GetValue(lowBound))
                j = 0
                For i = lowBound To upBound
                    Select Case dataType
                        Case VariantType.Short '2 16-bit Signed integer
                            BW.Write(CShort(ArrayData.GetValue(i)))
                        Case VariantType.Integer ' 3 32-bit signed integer (formerly long)
                            BW.Write(CInt(ArrayData.GetValue(i)))
                        Case VariantType.Single '4 single
                            BW.Write(CSng(ArrayData.GetValue(i)))
                        Case VariantType.Double '5 double
                            BW.Write(CDbl(ArrayData.GetValue(i)))
                        Case VariantType.Date '7 date
                            ' In VB6, one could cast a Date to a Double.
                            ' This no longer works in VB.NET
                            ' Use 'ToOADate()' instead.
                            Dim theDate As DateTime = CType(ArrayData.GetValue(i), DateTime)
                            Dim asDouble As Double = theDate.ToOADate()
                            BW.Write(asDouble)
                        Case VariantType.String '8 string
                            BW.Write(CStr(ArrayData.GetValue(i)))
                        Case VariantType.Byte '17 byte
                            BW.Write(CByte(ArrayData.GetValue(i)))
                        Case Else
                            ' data type not handled
                            Throw New Exception("Data type not handled.")
                    End Select
                Next i
            End Using
        End Sub


        Private Function getArrayFromFile_NewFormat(numOfElements As Integer, FileName As String,
                                                    dataType As VariantType, ByRef ArrayData As Array) As ReturnCodes
            Dim i As Integer

            If VarType(ArrayData) < vbArray Then
                ' not an array
                LogError(New Exception("This is not an array."))
                Return ReturnCodes.UDBS_ERROR
            End If

            Try
                Using FS As New FileStream(FileName, FileMode.Open)
                    Dim BR As New BinaryReader(FS, Encoding.Unicode)
                    'retrieve array an array at a time
                    '**** does not work with variant type
                    Select Case dataType
                        Case VariantType.Short '2 16-bit Signed integer
                            ArrayData = Array.CreateInstance(GetType(Short), numOfElements)
                            For i = 0 To numOfElements - 1
                                ArrayData.SetValue(BR.ReadInt16, i)
                            Next

                        Case VariantType.Integer ' 3 32-bit signed integer (formerly long)
                            ArrayData = Array.CreateInstance(GetType(Integer), numOfElements)
                            For i = 0 To numOfElements - 1
                                ArrayData.SetValue(BR.ReadInt32, i)
                            Next
                        Case VariantType.Single '4 single
                            ArrayData = Array.CreateInstance(GetType(Single), numOfElements)
                            For i = 0 To numOfElements - 1
                                ArrayData.SetValue(BR.ReadSingle, i)
                            Next
                        Case VariantType.Double '5 double
                            ArrayData = Array.CreateInstance(GetType(Double), numOfElements)
                            For i = 0 To numOfElements - 1
                                ArrayData.SetValue(BR.ReadDouble, i)
                            Next
                        Case VariantType.Date '7 date
                            ArrayData = Array.CreateInstance(GetType(Date), numOfElements)
                            For i = 0 To numOfElements - 1
                                ArrayData.SetValue(Date.FromOADate(BR.ReadDouble), i)
                            Next
                        Case VariantType.String '8 string
                            ArrayData = Array.CreateInstance(GetType(String), numOfElements)

                            Try
                                For i = 0 To numOfElements - 1
                                    ArrayData.SetValue(BR.ReadString, i)
                                Next
                            Catch eos As EndOfStreamException
                                ' Try old vb6

                                ' Recreate the reader to start reading from the beginning.
                                FS.Seek(0, 0)
                                BR = New BinaryReader(FS, Encoding.Unicode)

                                Try
                                    For i = 0 To numOfElements - 1
                                        Dim length = BitConverter.ToInt16(BR.ReadBytes(2), 0) 'first 2 bytes are length
                                        Dim item = Encoding.ASCII.GetString(BR.ReadBytes(length)) ' ASCII encoding
                                        ArrayData.SetValue(item, i)
                                    Next

                                Catch ex As Exception
                                    Throw
                                End Try

                            Catch ex As Exception
                                Throw
                            End Try

                        Case VariantType.Byte '17 byte
                            ArrayData = Array.CreateInstance(GetType(Byte), numOfElements)
                            For i = 0 To numOfElements - 1
                                ArrayData.SetValue(BR.ReadByte, i)
                            Next
                        Case Else
                            ' data type not handled
                            Throw New Exception("Data type not handled.")
                    End Select

                    getArrayFromFile_NewFormat = ReturnCodes.UDBS_OP_SUCCESS
                End Using

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

#Region "IDisposable Support"

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    CloseNetworkDB()
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region
    End Class
End Namespace
