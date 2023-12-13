Imports UdbsInterface.KittingInterface
Imports UdbsInterface.MasterInterface
Imports UdbsInterface.TestDataInterface

'Public Enum TestResultCode As Integer
'    UDBS_ERROR = ResultCodes.UDBS_ERROR
'    FAIL = ResultCodes.UDBS_SPECS_FAIL
'    FAIL_HI = ResultCodes.UDBS_SPECS_FAIL_HI
'    FAIL_INC = ResultCodes.UDBS_SPECS_FAIL_INC
'    FAIL_LO = ResultCodes.UDBS_SPECS_FAIL_LO
'    NO_SPECS = ResultCodes.UDBS_SPECS_NONE
'    PASS = ResultCodes.UDBS_SPECS_PASS
'    PASS_INC = ResultCodes.UDBS_SPECS_PASS_INC
'    SANITY = ResultCodes.UDBS_SPECS_SANITY
'    SANITY_HI = ResultCodes.UDBS_SPECS_SANITY_HI
'    SANITY_INC = ResultCodes.UDBS_SPECS_SANITY_INC
'    SANITY_LO = ResultCodes.UDBS_SPECS_SANITY_LO
'    WARNING = ResultCodes.UDBS_SPECS_WARNING
'    WARNING_HI = ResultCodes.UDBS_SPECS_WARNING_HI
'    WARNING_INC = ResultCodes.UDBS_SPECS_WARNING_INC
'    WARNING_LO = ResultCodes.UDBS_SPECS_WARNING_LO
'End Enum

''' <summary>
'''     Static Shared class containing useful methods for accessing the UDBS manufacturing database.
''' </summary>
''' <remarks></remarks>
''' <editHistory date="August 2011" developer="Susan French">Created</editHistory>
Public NotInheritable Class UdbsTools

#Region "Fields"

    Shared ReadOnly logger As Logger = LogManager.GetLogger("UDBS")

#End Region

#Region "Events"

#End Region

#Region "Properties"


    ''' <summary>
    '''     Gets a string description of the UDBS return code passed as argument
    ''' </summary>
    ''' <param name="udbsCode">ReturnCodes</param>
    ''' <returns>String</returns>
    Public Shared Function InterpretUDBSReturnCode(udbsCode As ReturnCodes) As String

        Select Case udbsCode
            Case ReturnCodes.UDBS_ERROR
                InterpretUDBSReturnCode = "UDBS Error"
            Case ReturnCodes.UDBS_LOCALDB_MISSING
                InterpretUDBSReturnCode = "Local DB is missing"
            Case ReturnCodes.UDBS_OP_FAIL
                InterpretUDBSReturnCode = "Fail"
            Case ReturnCodes.UDBS_OP_INC
                InterpretUDBSReturnCode = "Incomplete"
            Case ReturnCodes.UDBS_OP_SUCCESS
                InterpretUDBSReturnCode = "Success"
            Case ReturnCodes.UDBS_RECORD_EXISTS
                InterpretUDBSReturnCode = "Record exists"
            Case ReturnCodes.UDBS_TABLE_MISSING
                InterpretUDBSReturnCode = "Table is missing"
            Case Else
                InterpretUDBSReturnCode = "Unknown"
        End Select
    End Function

    ''' <summary>
    ''' Casts a result from integer to <see cref="ResultCodes"/>
    ''' </summary>
    ''' <param name="intResultCode">The result code as an integer.</param>
    ''' <returns><see cref="ResultCodes"/></returns>
    Friend Shared Function ConvertToResultCode(intResultCode As Integer) As ResultCodes

        Dim resultCode = CType(intResultCode, ResultCodes)
        Return resultCode
    End Function

    ''' <summary>
    '''     Gets a string description of the UDBS result code passed as argument
    ''' </summary>
    ''' <param name="udbsCode">ResultCodes</param>
    ''' <returns>String</returns>
    Public Shared Function InterpretResultCode(udbsCode As ResultCodes,
                                        Optional ByVal dataItem As CTestData_Item = Nothing) As String

        Select Case udbsCode
            Case ResultCodes.UDBS_SPECS_PASS
                InterpretResultCode = "PASS"
            Case ResultCodes.UDBS_SPECS_PASS_INC,
                ResultCodes.UDBS_SPECS_FAIL_INC,
                ResultCodes.UDBS_SPECS_SANITY_INC,
                ResultCodes.UDBS_SPECS_WARNING_INC
                InterpretResultCode = "INCOMPLETE"
            Case ResultCodes.UDBS_SPECS_NONE
                InterpretResultCode = "NO SPEC"
            Case ResultCodes.UDBS_SPECS_FAIL
                InterpretResultCode = "FAIL"
            Case ResultCodes.UDBS_SPECS_FAIL_HI
                InterpretResultCode = "FAIL HIGH"
                If dataItem IsNot Nothing Then
                    InterpretResultCode += " (Maximum = " & dataItem.FailMax & dataItem.Units & ")"
                End If
            Case ResultCodes.UDBS_SPECS_FAIL_LO
                InterpretResultCode = "FAIL LOW"
                If dataItem IsNot Nothing Then
                    InterpretResultCode += " (Minimum = " & dataItem.FailMin & dataItem.Units & ")"
                End If
            Case ResultCodes.UDBS_ERROR
                InterpretResultCode = "ERROR"
            Case ResultCodes.UDBS_SPECS_WARNING
                InterpretResultCode = "WARNING"
            Case ResultCodes.UDBS_SPECS_WARNING_HI
                InterpretResultCode = "WARNING HIGH"
                If dataItem IsNot Nothing Then
                    InterpretResultCode += " (Maximum = " & dataItem.WarningMax & dataItem.Units & ")"
                End If
            Case ResultCodes.UDBS_SPECS_WARNING_LO
                InterpretResultCode = "WARNING LOW"
                If dataItem IsNot Nothing Then
                    InterpretResultCode += " (Minimum = " & dataItem.WarningMin & dataItem.Units & ")"
                End If
            Case ResultCodes.UDBS_SPECS_SANITY
                InterpretResultCode = "SANITY FAIL"
            Case ResultCodes.UDBS_SPECS_SANITY_HI
                InterpretResultCode = "SANITY FAIL HIGH"
            Case ResultCodes.UDBS_SPECS_SANITY_LO
                InterpretResultCode = "SANITY FAIL LOW"
            Case Else
                InterpretResultCode = "UNKNOWN"
        End Select
    End Function

    Friend Shared ReadOnly Property StationName As String
        Get
            Try
                Dim station = "Unknown"
                If CTestData_Utility.GetStationName(station) = ReturnCodes.UDBS_OP_SUCCESS Then
                    Return station
                End If
            Catch ex As Exception
                logger.Debug(ex, ex.Message)
            End Try

            Return "Unknown"
        End Get
    End Property

    ''' <remarks>Not used from anywhere. Candidate for removal.</remarks>
    Friend Shared Property IsDebugMode As Boolean
        Get
            Return UDBSDebugMode
        End Get
        Set
            UDBSDebugMode = Value
        End Set
    End Property

    ''' <summary>
    ''' UDBS logging level. Read from configuration setting {App}UdbsErrorLogLevel
    ''' </summary>
    Public Shared Property ErrorLogLevel As NLog.LogLevel = LogLevel.Debug

    ''' <summary>
    ''' Gets boolean indicating whether the given serial number exists with the given part ID
    ''' </summary>
    ''' <param name="partID">String</param>
    ''' <param name="serialNumber">String</param>
    ''' <value>Boolean</value>
    ''' <remarks>
    ''' Only called from <see cref="IsTestAvailableForRestart(String, String, String, String, Integer, ByRef String)"/>,
    ''' which itself is not called from anywhere.
    ''' </remarks>
    Friend Shared ReadOnly Property UnitExists(partID As String, serialNumber As String) As Boolean
        Get
            Dim utility As New CTestData_Utility
            Try
                Return utility.isUnitExist(partID, 0, serialNumber)
            Catch ex As Exception
                logger.Warn(ex, ex.Message)
                Return False
            Finally
                utility.Dispose()
                utility = Nothing
            End Try
        End Get
    End Property

    ''' <remarks>Not called from anywhere. Candidate for removal.</remarks>
    Friend Shared Function IsTestAvailableForRestart(productID As String, serialNumber As String,
                                                     testStage As String, stationID As String,
                                                     testSequence As Integer,
                                                     Optional ByRef logMessage As String = vbNullString) As Boolean

        Dim testInstance = New CTestdata_Instance
        Dim available = False

        Try
            If UDBSDebugMode Then
                logger.Debug("Looking for UDBS Test {3} for SN {0}, ID {1}, Stage {2}", serialNumber, productID,
                             testStage, testSequence)
            End If
            UDBSDebugMode = False 'do not pop-up error messages from within UDBS calls

            ' make sure SN exists in db
            If Not UnitExists(productID, serialNumber) Then
                Throw New UdbsTestException(String.Format("SN {0}, ID {1} does not exist in database",
                                                          serialNumber, productID))
            End If

            ' check to see if there is already an open test instance
            Dim returnCode As ReturnCodes = testInstance.LoadExisting(testStage, productID, serialNumber, testSequence)
            If (returnCode = ReturnCodes.UDBS_OP_SUCCESS) Then

                logMessage =
                    String.Format(
                        "Found a UDBS testdata instance for SN {0}, Sequence {1}, started on {2}, with status = {3}",
                        serialNumber, testInstance.Sequence, testInstance.StartDate, testInstance.Status)

                ' Check the status of the test instance
                ' Possible values: STARTING, IN PROCESS, PAUSED, COMPLETED, TERMINATED
                Select Case testInstance.Status
                    Case "IN PROCESS", "PAUSED", "TERMINATED"

                        'Make sure that it is not in process on a different station
                        Dim udbsStationName As String = StationName
                        If testInstance.StationName <> udbsStationName And testInstance.StationName <> stationID Then
                            Throw _
                                New UdbsTestInProcessException(serialNumber, productID, testInstance.StationName,
                                                               testInstance.Status)
                        End If

                        available = True

                    Case "COMPLETED"

                        logMessage = String.Format("{0}, completed on {1}", logMessage, testInstance.StopDate)

                End Select
            Else
                Return False
            End If

            logger.Debug(logMessage)
            Return available

        Catch ex As Exception
            logMessage = logMessage & ". " & ex.Message
            logger.Warn(ex, ex.Message)
            Return False

        Finally
            testInstance.Dispose()
            testInstance = Nothing
        End Try
    End Function

    ''' <summary>
    '''     Tries to load the Item list for the given product ID and test stage.
    ''' </summary>
    ''' <param name="productID">UDBS Part ID</param>
    ''' <param name="testStage">UDBS Test Stage name</param>
    ''' <param name="errorMessage">String. Error message returned by reference.</param>
    ''' <returns>Boolean. True if Item list loaded successfully, False otherwise.</returns>
    ''' <remarks></remarks>
    Friend Shared Function TestStageExists(productID As String, testStage As String,
                                           ByRef errorMessage As String) As Boolean

        Dim itemList = New CTestData_ItemList
        Try
            ' load the item list
            Dim returnCode As ReturnCodes = itemList.LoadItemList(productID, 1, testStage, 0)
            If (returnCode = ReturnCodes.UDBS_OP_SUCCESS) Then
                If UDBSDebugMode Then
                    logger.Debug("Test stage {0} exists for ID {1}", testStage, productID)
                End If
                Return True
            Else
                Throw New UdbsTestException("Could not load " & testStage & " item list for ID " & productID,
                                            returnCode)
            End If

        Catch ex As Exception
            errorMessage = (ex.Message)
            Return False
        Finally
            'Release the item list from memory
            If itemList IsNot Nothing Then
                'release existing object
                itemList.Dispose()
                itemList = Nothing
            End If
        End Try
    End Function


#End Region

#Region "Methods"

    ''' <summary>
    '''     Translates product/part number to product identifier
    ''' </summary>
    ''' <param name="partNumber">String. Oracle Part Number</param>
    ''' <returns>String. UDBS Product ID.</returns>
    ''' <remarks>
    ''' Suppresses exceptions and returns null string if not found.
    ''' This is not used from anywhere. Candidate for removal.
    ''' </remarks>
    Friend Shared Function LookupPartID(partNumber As String) As String
        If String.IsNullOrWhiteSpace(partNumber) Then
            Return vbNullString
        End If

        Using utility = New CTestData_Utility
            Try
                Dim productID As String = vbNullString
                If utility.GetPartIdentifier(partNumber, productID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    logger.Debug("Could not find part ID for part number " & partNumber)
                End If
                Return productID
            Catch ex As Exception
                logger.Warn(ex, "Could not find part ID for part number {0}. {1}", partNumber, ex.Message)
            End Try
        End Using

        Return vbNullString
    End Function

    ''' <summary>
    '''     Looks up the Part ID for the given Part Number
    ''' </summary>
    ''' <param name="partNo">String. Oracle Part Number</param>
    ''' <param name="unitID">String. UDBS Product ID. Returned by reference</param>
    ''' <returns>Boolean. True if found, False otherwise</returns>
    ''' <remarks>This is not used from anywhere. Candidate for removal.</remarks>
    Friend Shared Function FindPartID(partNo As String, ByRef unitID As String) As Boolean

        Dim utility = New CUtility
        Try
            Dim retCode As ReturnCodes = utility.Product_GetPartIdentifier(partNo, unitID)

            Return (retCode = ReturnCodes.UDBS_OP_SUCCESS)
        Catch ex As Exception
            Return False
        Finally
            utility.Dispose()
            utility = Nothing
        End Try
    End Function

    ''' <summary>
    '''     Queries UDBS for the list of PN for the given ID. Returns the last PN in the list,
    '''     or empty string if not found.
    ''' </summary>
    ''' <param name="UnitID">String. Udbs Product ID</param>
    ''' <returns>String. Default part number</returns>
    ''' <remarks>Logs and suppresses exceptions. Returns empty string if not found</remarks>
    Friend Shared Function LookupDefaultPartNo(unitID As String) As String
        Return CProduct.LookupDefaultPartNo(unitID)
    End Function

    ''' <summary>
    '''     Based on the serial number of the optics module and the type of component given,
    '''     looks up the part number and serial number of the component part from the kitting database,
    '''     and returns the values by reference.
    ''' </summary>
    ''' <param name="serialNumber">Input serial number of optics module</param>
    ''' <param name="kittingItemname">Input name of item in kitting database</param>
    ''' <param name="kittingItemSN">Output Serial number of the found component part</param>
    ''' <param name="kittingItemPartNumber">Output Part number of the found component part</param>
    ''' <param name="kittingItemPartID">Output Part ID of the found component part</param>
    ''' <param name="kitting_sequence">process_sequence field of the kitting_process table</param>
    ''' <param name="kittingRemarks">Output Kitting remarks of the found component part</param>
    ''' <returns>True if successful, False otherwise</returns>
    ''' <remarks>Not called from anywhere. Candidate for removal.</remarks>
    <Obsolete("Use KittingUtility.LookupKittingInfo(...)")>
    Friend Shared Function LookupKittingInfo(serialNumber As String,
                                             kittingItemname As String,
                                             ByRef kittingItemSN As String,
                                             ByRef kittingItemPartNumber As String,
                                             ByRef kittingItemPartID As String,
                                             Optional ByRef kittingRemarks As String = vbNullString,
                                             Optional ByRef kitting_sequence As Integer = Integer.MinValue) As Boolean
        Return _
            KittingUtility.LookupKittingInfo(serialNumber, kittingItemname, kittingItemSN, kittingItemPartNumber,
                                             kittingItemPartID, kitting_sequence, kittingRemarks)
    End Function


    ''' <summary>
    '''     Creates a new kitting process sequence for the given serial number with the given UDBS ID.
    '''     Copies all results from the most recent process sequence.
    '''     Updates the PN and SN for each item given in the sItemname array.
    ''' </summary>
    ''' <param name="udbsID"></param>
    ''' <param name="serialNo"></param>
    ''' <param name="empID"></param>
    ''' <param name="itemname"></param>
    ''' <param name="componentPN"></param>
    ''' <param name="componentSN"></param>
    ''' <param name="notes"></param>
    ''' <returns></returns>
    <Obsolete("Use KittingUtility.UpdateKittingData(...)")>
    Friend Shared Function UpdateKittingData(udbsID As String,
                                      serialNo As String,
                                      empID As String,
                                      itemname() As String,
                                      componentPN() As String,
                                      componentSN() As String,
                                      Optional notes As String = "") As Boolean
        Return KittingUtility.UpdateKittingData(udbsID, serialNo, empID, itemname, componentPN, componentSN, notes)
    End Function

    <Obsolete("Use KittingUtility.GetHighestKittingItemlistRevision(...)")>
    Friend Shared Function GetHighestKittingItemlistRevision(serialNumber As String) As Integer
        Return KittingUtility.GetHighestKittingItemlistRevision(serialNumber)
    End Function

#End Region

#Region "Constructor"

    ''' <summary>
    '''     Constructor is Private, to prevent instantiation
    ''' </summary>
    Private Sub New()
        'will never get here
    End Sub

#End Region
End Class
