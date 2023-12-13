Option Explicit On
Option Strict On
Option Compare Text
Option Infer On

Imports UdbsInterface.MasterInterface

Namespace TestDataInterface
    Public Class CTestData_Result
        Inherits CTestData_Item

        Implements IDisposable

        ' Index to PI for this item
        Private mPI As CProcessInstance
        Private mReadOnly As Boolean            ' set by results object

        ' Itemlist and Result data
        Private mPassFlagStored As Boolean
        Private mValueStored As Boolean
        Private mValue As Double
        Private mPassFlag As Integer
        Private mStringData As String

#Region "Properties"

        ''' <summary>
        ''' The measured value associated with that test item.
        ''' Check 'ValueStored' first, otherwise you will get NaN.
        ''' 
        ''' Old comment: Used to fill the results collection, not required to update DB. (Not sure what this mean...)
        ''' </summary>
        ''' <returns>The measured value for this test item.</returns>
        Public Property Value As Double
            Get
                Return mValue
            End Get
            Set(tmpValue As Double)
                mValue = tmpValue
            End Set
        End Property

        ''' <summary>
        ''' The outcome of the evaluation of the 'Value' against the test specifications or,
        ''' in the case of a 'group', the aggregation of the 'PassFlag' of all sub-items.
        ''' This integer should match one of the 'ResultCodes'.
        ''' </summary>
        ''' <returns>The outcome of this test item.</returns>
        ''' <see cref="ResultCodes"/>
        Public Property PassFlag As Integer
            Get
                Return mPassFlag
            End Get
            Set
                mPassFlag = Value
            End Set
        End Property

        Public Property StringData As String
            Get
                Return mStringData
            End Get
            Set
                mStringData = Value
            End Set
        End Property

        Public Property ResultBlobDataExists As Boolean

        ''' <summary>
        ''' Whether or not the 'PassFlag' has been evaluated for this particular
        ''' test item.
        ''' </summary>
        ''' <returns>Whether of not the 'PassFlag' has been evaluated.</returns>
        Public Property PassFlagStored As Boolean
            Get
                Return mPassFlagStored
            End Get
            Set
                mPassFlagStored = Value
            End Set
        End Property

        ''' <summary>
        ''' Whether or not a test value is stored.
        ''' This is causing confusion, because a 'result' could be interpreted as the outcome of a test; 
        ''' i.e. whether or not the test passes.
        ''' This is being renamed as 'ValueStored'.
        ''' The test outcome (whether or not the value meets the specifications) is the 'PassFlag' property.
        ''' </summary>
        ''' <returns>Whether or not the test value is store.</returns>
        ''' <see cref="ValueStored"/>
        <Obsolete("Use ValueStored instead. This will be removed some time after 2021-12.", False)>
        Public Property ResultStored As Boolean
            Get
                Return mValueStored
            End Get
            Set
                mValueStored = Value
            End Set
        End Property

        ''' <summary>
        ''' Whether or not a test value is stored.
        ''' Checks whether or not a value is stored before accessing this property.
        ''' You will get 'NaN' if no value is stored.
        ''' </summary>
        ''' <returns>Whether or not the test value is store.</returns>
        Public Property ValueStored As Boolean
            Get
                Return mValueStored
            End Get
            Set
                mValueStored = Value
            End Set
        End Property

        Friend Sub SetReadMode(ReadOnlyMode As Boolean)
            mReadOnly = ReadOnlyMode
        End Sub

#End Region
        '**********************************************************************
        '* Methods
        '**********************************************************************

        Private Function GetResultId() As Long
            ' acquire result_id from server
            Dim sSQL = "SELECT result_id FROM " & PROCESS & "_result with(nolock), " &
                PROCESS & "_itemlistdefinition with(nolock) " &
                "WHERE itemlistdef_id=result_itemlistdef_id " &
                "AND result_process_id = " & mPI.ID & " " &
                "AND itemlistdef_itemname = '" & mItemName & "'"

            Dim tmpRS As New DataTable
            If CUtility.Utility_ExecuteSQLStatement(sSQL, tmpRS) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New UDBSException("Failed to query server for result ID.")
            End If

            If (If(tmpRS?.Rows?.Count, 0)) = 1 Then
                Return KillNullLong(tmpRS(0)("result_id"))
            Else
                ' cnanot be more than ONE record
                Throw New UDBSException("Expecting exactly one result id to be returned.")
            End If
        End Function

        ''' <summary>
        ''' Get an array of data that was attached to that test item.
        ''' </summary>
        ''' <param name="arrayName">The name of the array (identifier).</param>
        ''' <param name="DataGroupName">(Out) The group to which it belongs.</param>
        ''' <param name="NumElements">(Out) The number or elements this array contains.</param>
        ''' <param name="DataType">(Out) The type of the data contained in the array.</param>
        ''' <param name="IsHeader">(Out) Whether or not this is a header. (TBD: What does that mean?)</param>
        ''' <param name="ArrayData">(Out) The data retrieved from the database.</param>
        ''' <returns></returns>
        Public Function GetArray(arrayName As String,
                                 ByRef DataGroupName As String,
                                 ByRef NumElements As Integer,
                                 ByRef DataType As VariantType,
                                 ByRef IsHeader As Boolean,
                                 ByRef ArrayData As Array) As ReturnCodes

            If Not mPI.IsReadOnly Then
                Return GetArray_Local(arrayName, DataGroupName, NumElements, DataType, IsHeader, ArrayData)
            End If

            Try
                Using blobObj As New CBLOB
                    Return blobObj.GetBLOB(PROCESS, PROCESS & "_result", GetResultId(), arrayName, DataGroupName, NumElements,
                                       DataType, IsHeader, ArrayData, "")
                End Using
            Catch ex As Exception
                LogError(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        Private Function GetArray_Local(arrayName As String,
                                 ByRef DataGroupName As String,
                                 ByRef NumElements As Integer,
                                 ByRef DataType As VariantType,
                                 ByRef IsHeader As Boolean,
                                 ByRef ArrayData As Array) _
            As ReturnCodes

            ' Get the array (ArrayData) to server
            Try
                Using blobObj As New CBLOB
                    ' get the array from server
                    Return blobObj.GetBLOB_Local(PROCESS, mItemName, mPI.ID, arrayName, DataGroupName, NumElements,
                                           DataType, IsHeader, ArrayData, "")
                End Using
            Catch ex As Exception
                LogError(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Retrieve a file associated with a test data item from the database.
        ''' </summary>
        ''' <param name="arrayName">The name of the array.</param>
        ''' <param name="DataGroupName">(Out) The group this array belongs to.</param>
        ''' <param name="IsHeader">(Out) Whether or not this is a header. (TBD: What does that mean?)</param>
        ''' <param name="FileSpec">Where to save the data retrieved from the database.</param>
        ''' <returns></returns>
        Public Function GetFile(arrayName As String,
                                ByRef DataGroupName As String,
                                ByRef IsHeader As Boolean,
                                FileSpec As String) As ReturnCodes
            Const fncName = "CTestdata_Result::GetFile"

            If String.IsNullOrEmpty(FileSpec) Then
                logger.Error($"Empty file name in {fncName}")
                Return ReturnCodes.UDBS_ERROR
            End If

            If Not mPI.IsReadOnly Then
                Return GetFile_Local(arrayName, DataGroupName, IsHeader, FileSpec)
            End If

            Try
                Using blobObj As New CBLOB()
                    ' Unused, mandatory arguments...
                    Dim NumElements As Integer
                    Dim DataType As VariantType
                    Dim ArrayData As Array = Nothing

                    GetFile = blobObj.GetBLOB(PROCESS, PROCESS & "_result", GetResultId(), arrayName, DataGroupName, NumElements,
                                          DataType, IsHeader, ArrayData, FileSpec)
                End Using
            Catch ex As Exception
                logger.Error(ex, $"Error in {fncName}")
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        Private Function GetFile_Local(arrayName As String,
                                ByRef DataGroupName As String,
                                ByRef IsHeader As Boolean,
                                FileSpec As String) _
            As ReturnCodes

            Const fncName = "CTestdata_Result::GetFile_Local"
            Try
                Using blobObj As New CBLOB()
                    ' Unused, mandatory arguments...
                    Dim NumElements As Integer
                    Dim DataType As VariantType
                    Dim ArrayData As Array = Nothing

                    Return blobObj.GetBLOB_Local(PROCESS, ItemName, mPI.ID, arrayName, DataGroupName, NumElements,
                                          DataType, IsHeader, ArrayData, FileSpec)
                End Using
            Catch ex As Exception
                logger.Error(ex, $"Error in {fncName}")
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Store an array as a BLOB in the local database.
        ''' The array will be pushed to the network database when the test is completed.
        ''' </summary>
        ''' <param name="arrayName">The name of the array.</param>
        ''' <param name="dataGroupName">The name of the group this array belongs to.</param>
        ''' <param name="isHeader">Whether or not this is a header. (More info needed...)</param>
        ''' <param name="arrayData">The data to be saved.</param>
        ''' <param name="lowBound">The lower boundary of the array to save.</param>
        ''' <param name="upBound">The upper boundary of the array to save.</param>
        ''' <returns>A return code indicating the outcome of this operation.</returns>
        Public Function StoreArray(
                arrayName As String,
                dataGroupName As String,
                isHeader As Boolean,
                arrayData As Array,
                lowBound As Integer,
                upBound As Integer) As ReturnCodes

            Try
                If mReadOnly Then
                    LogError(New Exception("Process is read only."))
                    Return ReturnCodes.UDBS_ERROR
                End If
                If (arrayData Is Nothing) Then
                    ' this is not an array
                    Throw New Exception("Invalid array")
                End If
                If (UBound(arrayData) - LBound(arrayData) + 1) > 1000000 Then
                    ' too many data, try using a zip file to store the data
                    Throw New Exception("Too many data (>1,000,000).")
                End If

                Dim blobObj As New CBLOB
                Dim currentPID As Integer = mPI.ID

                ' Force it to create a result record
                If mPI.StoreResultField(mItemName, "result_blobdata_exists", CStr(If(ResultBlobDataExists, -1, 0))) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Throw New Exception("failed to update result")
                End If

                Dim ret = blobObj.StoreBLOB_Local(currentPID, mItemName, arrayName,
                                                  dataGroupName, isHeader, arrayData,
                                                  "", lowBound, upBound)
                If ret = ReturnCodes.UDBS_OP_SUCCESS Then
                    ResultBlobDataExists = True
                    ret = mPI.StoreResultField(mItemName, "result_blobdata_exists", CStr(-1))
                End If

                Return ret
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_OP_FAIL
            End Try
        End Function


        ''' <summary>
        ''' Store a file as a BLOB in the local database.
        ''' The BLOB will be pushed to the network database when the test is completed.
        ''' </summary>
        ''' <param name="fileName">
        ''' The name of the file to store in the database.
        ''' Note: This does not have to match the name of the file on the local hard drive at
        ''' the time the file is stored.
        ''' </param>
        ''' <param name="dataGroupName">The name of the group this array belongs to.</param>
        ''' <param name="isHeader">Whether or not this is a header. (More info needed...)</param>
        ''' <param name="filePath">The path of the file to store.</param>
        ''' <returns>A return code indicating the outcome of this operation.</returns>
        Public Function StoreFile(
                fileName As String,
                dataGroupName As String,
                isHeader As Boolean,
                filePath As String) As ReturnCodes

            Try
                If mReadOnly Then
                    LogError(New Exception("Process is read only."))
                    Return ReturnCodes.UDBS_ERROR
                End If
                If String.IsNullOrEmpty(filePath) Then
                    ' this is not a file
                    Throw New ArgumentNullException("Empty [FileSpec] provided.")
                End If
                If Not IO.File.Exists(filePath) Then
                    ' No filespec supplied
                    Throw New IO.FileNotFoundException("Source file not found.", filePath)
                End If

                Dim currentPID As Integer = mPI.ID
                Dim blobObj As New CBLOB

                ' Force it to create a result record.
                If mPI.StoreResultField(mItemName, "result_blobdata_exists", CStr(If(ResultBlobDataExists, -1, 0))) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Throw New Exception("Failed to update result field")
                End If

                Dim ret = blobObj.StoreBLOB_Local(currentPID, mItemName, fileName,
                                                  dataGroupName, isHeader, Nothing, filePath, 0, 0)

                If ret = ReturnCodes.UDBS_OP_SUCCESS Then
                    ResultBlobDataExists = True
                    ret = mPI.StoreResultField(mItemName, "result_blobdata_exists", CStr(-1))
                    mPI._filesAttachedToUDBS.Add(filePath)
                End If

                Return ret
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_OP_FAIL
            End Try
        End Function

        Public Function StoreValue(MeasuredValue As Double) _
            As ResultCodes
            ' Store result to database and check against specs
            Dim Ret As ReturnCodes

            'Default the result to an error... we have to follow "guilty until proven innocent" very carefully...

            If mReadOnly Then
                LogError(New Exception("Process is read only."))
                Return ResultCodes.UDBS_ERROR
            End If

            Ret = mPI.StoreResultField(mItemName, "value", CStr(MeasuredValue))
            If Ret <> ReturnCodes.UDBS_OP_SUCCESS Then
                'Couldn't store the value for some reason...
                'This should really be a more defined error code other than the generic one
                'It would have been nice to store this error to the DB, but considering we couldn't store the value we can't store
                'the resultflag for it!
                'No sense compunding system errors!
                Return ResultCodes.UDBS_ERROR
            End If

            'Value stored successfully... check what the result of the value was...
            Dim Result = CheckValue(MeasuredValue)
            'If checking the value had any problems in evaluation, the error WILL be returned by CheckValue.
            If mPI.StoreResultField(mItemName, "passflag", CStr(Result)) <> ReturnCodes.UDBS_OP_SUCCESS Then
                'Some wierd problem storing the result (odd since other stuff previous to this must have worked)
                'Kick back an error to the calling function...
                Return ResultCodes.UDBS_ERROR
            End If

            mValue = MeasuredValue
            mPassFlag = Result
            mPassFlagStored = True
            mValueStored = True

            ' Reminder: Storing a value should clear the process_result column... it will no longer be valid.
            If mPI.StoreProcessInstanceField("result", CStr(ResultCodes.UDBS_SPECS_FAIL_INC)) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ResultCodes.UDBS_ERROR
            End If

            Return Result
        End Function

        ''' <summary>
        ''' Clear a given result, in case a measurement needs to be repeated.
        ''' </summary>
        ''' <returns>The outcome of the operation.</returns>
        Public Function Clear() As ReturnCodes
            If mReadOnly Then
                LogError(New Exception($"Attempting to clear result ""${ItemName}"" of read-only process."))
                Return ReturnCodes.UDBS_ERROR
            End If

            If mItemBlobDataExists <> 0 Then
                LogError(New Exception($"Unable to clear result ""${ItemName}"" because one or many files are attached to it."))
                Return ReturnCodes.UDBS_ERROR
            End If

            If mPI.StoreResultField(mItemName, "value", Nothing) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_ERROR
            End If
            If StoreField("stringdata", Nothing) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_ERROR
            End If
            If StoreField("passflag", Nothing) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Return ReturnCodes.UDBS_ERROR
            End If

            mValue = 0.0
            mPassFlagStored = False
            mValueStored = False

            Return ReturnCodes.UDBS_OP_SUCCESS
        End Function

        Public Function StoreField(FieldName As String,
                                   FieldValue As String) _
            As ReturnCodes
            ' Function allows any column to be edited"
            'Guilty before innocent
            Dim result = ReturnCodes.UDBS_ERROR

            If mReadOnly Then
                LogError(New Exception("Process is read only."))
                Return result
            End If
            If (LCase(FieldName) = "value") OrElse (LCase(FieldName) = "result_value") Then
                ' Should not use this function to store value!!!
                LogError(New Exception("Use 'StoreValue' to store result value."))
                Return result
            End If
            result = mPI.StoreResultField(mItemName, FieldName, FieldValue)
            If result = ReturnCodes.UDBS_OP_SUCCESS Then ' UDBS_OP_SUCCESS
                ' only stringdata or value field will be updated
                If (LCase(FieldName) = "stringdata") OrElse (LCase(FieldName) = "result_stringdata") Then
                    mStringData = FieldValue
                End If
                If (LCase(FieldName) = "passflag") OrElse (LCase(FieldName) = "result_passflag") Then
                    mPassFlag = CInt(Val(FieldValue))
                    mPassFlagStored = True
                End If
            End If
            Return result
        End Function


        Friend Sub New(ByRef ProcessInstance As CProcessInstance,
                       ItemNumber As Integer,
                       ItemName As String,
                       Descriptor As String,
                       Description As String,
                       ReportLevel As Integer,
                       Units As String,
                       CriticalSpec As Integer,
                       WarningMin As Double,
                       WarningMax As Double,
                       FailMin As Double,
                       FailMax As Double,
                       SanityMin As Double,
                       SanityMax As Double,
                       ItemBlobDataExists As Integer) _
            ' Function populates the item object,
            ' Specs are passed in as strings to allow for null (no specs)
            MyBase.New(ItemNumber, ItemName, Descriptor, Description, ReportLevel, Units, CriticalSpec, WarningMin,
                       WarningMax, FailMin, FailMax, SanityMin, SanityMax)
            mPI = ProcessInstance

            mItemBlobDataExists = ItemBlobDataExists

            mValueStored = False
            mPassFlagStored = False
        End Sub


        Public Function CheckValue(ResultValue As Double) _
            As ResultCodes

            'Default the result flag to a horrible error by default ("guilty until proven innocent").  This in case of any weird uncaught problem...
            CheckValue = ResultCodes.UDBS_ERROR

            Try
                'Check to see if there is any specs on this item!
                If mHasSpecs = True Then
                    'Need to determine how the value sits within the specs...
                    ' Are there specs on sanity?
                    If mSanityMinSpec = True Then
                        ' Does the value fail?
                        If ResultValue < mSanityMin Then
                            ' Failed
                            CheckValue = ResultCodes.UDBS_SPECS_SANITY_LO
                            Exit Function
                        End If
                    End If

                    If mSanityMaxSpec = True Then
                        ' Does the value fail?
                        If ResultValue > mSanityMax Then
                            ' Failed
                            CheckValue = ResultCodes.UDBS_SPECS_SANITY_HI
                            Exit Function
                        End If
                    End If

                    ' Are there specs on fail?
                    If mFailMinSpec = True Then
                        ' Does the value fail?
                        If ResultValue < mFailMin Then
                            ' Failed
                            CheckValue = ResultCodes.UDBS_SPECS_FAIL_LO
                            Exit Function
                        End If
                    End If

                    If mFailMaxSpec = True Then
                        ' Does the value fail?
                        If ResultValue > mFailMax Then
                            ' Failed
                            CheckValue = ResultCodes.UDBS_SPECS_FAIL_HI
                            Exit Function
                        End If
                    End If

                    ' Are there specs on warning?
                    If mWarningMinSpec = True Then
                        ' Does the value fail?
                        If ResultValue < mWarningMin Then
                            ' Failed
                            CheckValue = ResultCodes.UDBS_SPECS_WARNING_LO
                            Exit Function
                        End If
                    End If

                    If mWarningMaxSpec = True Then
                        ' Does the value fail?
                        If ResultValue > mWarningMax Then
                            ' Failed
                            CheckValue = ResultCodes.UDBS_SPECS_WARNING_HI
                            Exit Function
                        End If
                    End If

                    ' Measurement has passed all specs
                    CheckValue = ResultCodes.UDBS_SPECS_PASS

                Else
                    'We have nothing to check, so return the "PASS" indication that there is no specification
                    CheckValue = ResultCodes.UDBS_SPECS_NONE
                End If
            Catch ex As Exception
                logger.Error(ex, $"Error while checking the value: {ex.Message}")
                CheckValue = ResultCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Get information about the BLOBs attached to this result.
        ''' </summary>
        ''' <param name="ArrayNames">(Output) The names of the attachments.</param>
        ''' <param name="ArrayDataType">(Output) The types of the attachments.</param>
        ''' <returns>The outcome of the operation.</returns>
        Public Function GetAttachmentList(ByRef ArrayNames() As String,
                                          ByRef ArrayDataType() As Integer) _
            As ReturnCodes
            ' Returns a list of attachments available
            Const fncName = "CTestdata_Result::GetAttachmentList"

            Try
                Dim blobObj As New CBLOB
                Dim currentPID As Integer
                Dim resultID As Long
                Dim sSQL As String
                Dim tmpRS As New DataTable

                currentPID = mPI.ID
                ' acquire result_id from server
                sSQL = "SELECT result_id FROM " & PROCESS & "_result, " &
                   PROCESS & "_itemlistdefinition " &
                   "WHERE itemlistdef_id=result_itemlistdef_id " &
                   "AND result_process_id = " & currentPID & " " &
                   "AND itemlistdef_itemname = '" & mItemName & "'"
                GetAttachmentList = CUtility.Utility_ExecuteSQLStatement(sSQL, tmpRS)
                If GetAttachmentList <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Throw New Exception("Error searching for result ID.")
                End If
                If (If(tmpRS?.Rows?.Count, 0)) = 1 Then
                    resultID = KillNullLong(tmpRS(0)("result_id"))
                End If
                If (If(tmpRS?.Rows?.Count, 0)) = 0 Then
                    ' no blob
                    Throw New Exception("Itemname not found.")
                End If
                If (If(tmpRS?.Rows?.Count, 0)) > 1 Then
                    ' cnanot be more than ONE record
                    Throw New Exception("More than ONE result items returned.")
                End If
                GetAttachmentList = blobObj.GetAttachmentList(PROCESS, PROCESS & "_result", resultID, ArrayNames,
                                                          ArrayDataType)
            Catch ex As Exception
                logger.Error(ex, $"{fncName}: Error retrieving attachment lists: {ex.Message}")
                GetAttachmentList = ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Logs the UDBS error in the application log and inserts it in the udbs_error table, with the process instance details:
        ''' Process Type, name, Process ID, UDBS product ID, Unit serial number.
        ''' </summary>
        ''' <param name="ex">Exception raised.</param>
        Private Sub LogErrorInDatabase(ex As Exception)

            If mPI Is Nothing Then
                DatabaseSupport.LogErrorInDatabase(ex)
            Else
                DatabaseSupport.LogErrorInDatabase(ex, mPI.Process, mPI.Stage, mPI.ID, mPI.ProductNumber, mPI.UnitSerialNumber)
            End If

        End Sub

#Region "IDisposable Support"

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    mPI = Nothing
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
