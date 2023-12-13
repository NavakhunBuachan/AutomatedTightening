Option Explicit On
Option Compare Text
Option Infer On
Option Strict On

Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions


Namespace MasterInterface
    ' Enumeration

    ''' <summary>
    ''' The return codes for the UDBS interface's many operations.
    ''' <see href="https://confluence.lumentum.com/display/TF/TMTD-269+Guidelines+for+UDBS+Return+Codes+Usage"/>
    ''' </summary>
    Public Enum ReturnCodes
        UDBS_ERROR = -1000000
        UDBS_OP_FAIL = -1
        UDBS_OP_INC = 0
        UDBS_OP_SUCCESS = 1
        UDBS_RECORD_EXISTS = -2147217900
        UDBS_LOCALDB_MISSING = -2147467259
        UDBS_TABLE_MISSING = -2147217865
    End Enum

    ''' <summary>
    ''' Generic helper class for the UDBS Interface library.
    ''' </summary>
    Public Class CUtility
        Implements IDisposable

        ' Table Identification
        Private ReadOnly mUDBSProcessesTable As String = "udbs_process"
        Private ReadOnly mProductTable As String = "product"
        Private ReadOnly mProductFamilyTable As String = "family"

        ''' <summary>
        ''' The time difference between the server time and the local time
        ''' the last time we successfully retrieved the server time.
        ''' If we ever lose connection to the server, we can reapply this
        ''' difference and produce a value closer to the actual server time.
        ''' </summary>
        Private Shared serverTimeDelta As Double = 0

        ''' <summary>
        ''' The stationID is an alternative, more human readable name for the computer than the default Machine name.
        ''' </summary>
        Private Shared stationName As String = String.Empty

        '**********************************************************************
        '* Standard Utility Functions
        '**********************************************************************

        ''' <summary>
        ''' Function allows calling application to change local files location strings for NEXT instance.
        ''' </summary>
        ''' <param name="LocationString"></param>
        Friend Sub Utility_SetLocalFilesLocationString(LocationString As String)
            SaveSetting("UDBS_V3", "LocalFiles", "Location", LocationString)
        End Sub

        ''' <summary>
        ''' Function allows calling application to change database connection strings for NEXT instance
        ''' </summary>
        ''' <param name="ConnectionString"></param>
        Friend Sub Utility_SetLocalDBConnectionString(ConnectionString As String)
            SaveSetting("UDBS_V3", "Database", "Local Connection", ConnectionString)
        End Sub

        Friend Function Utility_SetNetworkDBConnectionString(ConnectionString As String) As ReturnCodes
            Try
                SetNetworkConnectionString(ConnectionString, True)
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogError(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        Friend Sub Utility_SetNetworkHome(HomeDirectory As String)
            ' Function allows calling application to change network home folder
            Dim HomeDir As String
            HomeDir = Trim(HomeDirectory)
            If Right$(HomeDir, 1) <> "\" Then
                HomeDir = HomeDir & "\"
            End If
            SaveSetting("UDBS_V3", "Network", "Network Home", HomeDir)
        End Sub

        ''' <summary>
        ''' Return the stationID or the computer name when the stationID doesn't exist.
        ''' </summary>
        ''' <param name="StationName">Station name as string.</param>
        ''' <returns>Udbs return code specifiying whether the operation is successful.</returns>
        Public Shared Function Utility_GetStationName(ByRef StationName As String) As ReturnCodes
            Try

                If String.IsNullOrEmpty(CUtility.stationName) Then

                    Dim CompName As String = GetSetting("UDBS_V3", "Station", "StationName", "")

                    If CompName = "" Then
                        CompName = Environment.MachineName
                    End If

                    StationName = CompName
                Else
                    StationName = CUtility.stationName
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogError(ex)
                StationName = "Not Found"
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ' Not used... Candidate for removal. (This is used by WIP Tracker - could also be used by Fractal in future.  To discuss.)
        Friend Function Utility_SetStationName(ByRef StationName As String) As ReturnCodes
            ' Purpose: Return the name of the computer as specified in the network settings
            Try
                SaveSetting("UDBS_V3", "Station", "StationName", StationName)
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogError(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Check if we are getting in an infinite loop.
        ''' We're basically checking if a given method is present more than once in the
        ''' call stack.
        ''' </summary>
        ''' <param name="aMethod">
        ''' The method to check.
        ''' If null, evaluates the method invoking this method (i.e. 2nd element of the stack).
        ''' </param>
        ''' <returns>Whether or not a recursive loop is detected.</returns>
        Friend Shared Function IsRecursiveLoop(Optional aMethod As MethodBase = Nothing) As Boolean
            Dim currentStack = New StackTrace()

            If (aMethod Is Nothing) Then
                ' Method not specified. Using the calling method.
                aMethod = currentStack.GetFrames(1).GetMethod()
            End If

            Dim methodCount As Integer = 0

            For Each frame In currentStack.GetFrames()
                If (frame.GetMethod() = aMethod) Then
                    methodCount += 1
                End If
            Next

            Return methodCount > 1
        End Function

        ''' <summary>
        ''' Get the server time.
        ''' </summary>
        ''' <param name="ServerTime">(Out) The server time.</param>
        ''' <returns>
        ''' The outcome of the operation.
        ''' On failure, the local computer time is returned, adjusted for
        ''' the time difference of the last successful call.
        ''' </returns>
        Friend Shared Function Utility_GetServerTime(ByRef ServerTime As Date) As ReturnCodes
            ' When there's an error getting the server time, we log an error, which
            ' also trigger a query to the server, causing an infinite recursive loop.
            If IsRecursiveLoop() Then
                ServerTime = AdjustedLocalTime
                Return ReturnCodes.UDBS_OP_SUCCESS
            End If

            ' Function returns the db server date and time if available
            Try

                Dim rsTime As DataTable = Nothing

                ' Ask the server what time it is
                OpenNetworkRecordSet(rsTime, "SELECT ""Date"" = getdate()")
                ServerTime = ConvertThaiToCommonEra(KillNullDate(rsTime(0)("Date")))
                rsTime = Nothing

                ' Every time we successfully get the server time, we also update
                ' the difference between the local PC time and the server time,
                ' so that if we ever need to get the server time and we fail,
                ' our estimate is as precise as it can be.
                serverTimeDelta = (ConvertThaiToCommonEra(DateTime.Now) - ServerTime).TotalMilliseconds
                Const MaxAcceptableServerTimeDelta As Double = 60000 ' 1 minute
                If (Math.Abs(serverTimeDelta) > MaxAcceptableServerTimeDelta) Then
                    logger.Warn($"Local time is off by {serverTimeDelta} msec. (Local time: {DateTime.Now}, Server time: {ServerTime})")
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                ' Logging a system error will cause an infinite loop, because we
                ' will be trying to get the time at which the system error occurred.
                logger.Warn($"Failure to get server time: {ex.Message}")
                ServerTime = ConvertThaiToCommonEra(AdjustedLocalTime)
                Return ReturnCodes.UDBS_ERROR
            Finally
                CloseNetworkDB()
            End Try
        End Function

        ''' <summary>
        ''' Converts a date from the Thai calendar to the Common Era, if applicable.
        ''' </summary>
        ''' <param name="toConvert">The date to convert.</param>
        ''' <returns>
        ''' The date in the Common Era.
        ''' If the date is already in the common era, the date is returned unchanged.
        ''' </returns>
        Public Shared Function ConvertThaiToCommonEra(toConvert As DateTime) As DateTime
            ' IMPORTANT: If the common era year is 2399 and you are reading this comment:
            '            First of all, "Hi! from the 21st century!"
            '            Second, I want to say that I am really sorry you have to maintain
            '            VB.NET code that was written in 2022.
            '            I also have a few questions for you: Are you an A.I.? Have we
            '            colonized Mars?
            '            Now, back to business...
            '            Pick a date roughly 400 years in the future (that would be around
            '            2800 if you are updating this code in 2399...) and update the following
            '            line. This should future-proof you code for the next few hundred years.
            Dim thaiDatePivot = DateTime.Parse("2400/01/01 00:00:00.000")
            If toConvert > thaiDatePivot Then
                ' If the date is after year 2400, assume we are using the Thai calendar.
                Return toConvert.AddYears(-543)
            Else
                ' No need to convert the date.
                Return toConvert
            End If
        End Function

        ''' <summary>
        ''' Local time, adjusted for the difference between the server time and
        ''' the real local the last time we successfully queried for the server time.
        ''' </summary>
        Private Shared ReadOnly Property AdjustedLocalTime As DateTime
            Get
                Return ConvertThaiToCommonEra(Now).AddMilliseconds(-serverTimeDelta)
            End Get
        End Property

        Friend Shared Function Utility_ExecuteSQLStatement(sqlQuery As String,
                                                           ByRef Resultset As DataTable) _
            As ReturnCodes
            ' Function allows interface to make direct sql queries
            Try
                If UCase(Left$(Trim(sqlQuery), 6)) = "SELECT" Or UCase(Left$(Trim(sqlQuery), 6)) = "UPDATE" Then
                    Return QueryNetworkDB(sqlQuery, Resultset)
                Else
                    LogError(New Exception("Cannot use function for queries other than SELECT or UPDATE."))
                    Return ReturnCodes.UDBS_ERROR
                End If
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ' Not used. Candidate for removal.
        Friend Function Utility_AddProcessSpecificTableRecord(TableName As String,
                                                              ByRef OneRecord As DataTable) _
            As ReturnCodes
            ' Function adds a record to a process specific table
            ' BECAREFUL of autoNum field, it may cause an error

            Try
                If OneRecord?.Rows.Count <> 1 Then
                    ' Passed DataTable must contain exactly one record
                    LogError(New Exception($"More than one record passed."))
                    Return ReturnCodes.UDBS_ERROR

                End If

                ' Make sure that the calling function is not trying to modify a core table
                If IsCoreTable(TableName) Then
                    LogError(New Exception("Cannot use function to modify UDBS Core Table."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Since we do not know the format of the table, we must assume that the calling function is
                ' adding this record as a NEW record. If it is a modified record then the ModifyProcessSpecificTableRecord
                ' should be called
                Dim dbCommand As String
                Dim FieldItem As DataColumn
                Dim Columns = ""
                Dim Values = ""
                Dim IDColumn As String


                If InStr(1, TableName, "_") <> 0 Then
                    IDColumn = Right(TableName, Len(TableName) - InStr(1, TableName, "_")) & "_id"
                Else
                    LogError(New Exception($"This is not a process specific table: {TableName}"))
                    Return ReturnCodes.UDBS_ERROR

                End If
                For Each FieldItem In OneRecord.Columns
                    If LCase(FieldItem.ColumnName) = LCase(IDColumn) Then
                        ' this is an autoNum field, skip it
                    Else
                        If (FieldItem.DataType = GetType(String)) Or
                           (FieldItem.DataType = GetType(Date)) Then
                            Columns = Columns & FieldItem.ColumnName & ", "
                            Values = Values & "'" & KillNull(OneRecord(0)(FieldItem.ColumnName)) & "', "
                        Else
                            Columns = Columns & FieldItem.ColumnName & ", "
                            Values = Values & KillNull(OneRecord(0)(FieldItem.ColumnName)) & ", "
                        End If
                    End If
                Next FieldItem
                ' Remove false comma
                Columns = Left$(Columns, Len(Columns) - 2)
                Values = Left$(Values, Len(Values) - 2)


                dbCommand = "INSERT INTO " & TableName & " (" & Columns & ") VALUES (" & Values & ") "
                ExecuteNetworkQuery(dbCommand)

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ' Not used. Candidate for removal.
        Friend Function Utility_ModifyProcessSpecificTableRecord(TableName As String,
                                                                 RecordID As Integer,
                                                                 ByRef OneRecord As DataTable) _
            As ReturnCodes
            ' Function modifies a record in a process specific table
            Const fncName = "CUtility::Utility_ModifyProcessSpecificTableRecord"
            Dim IDColumn As String

            Try
                If InStr(1, TableName, "_") <> 0 Then
                    IDColumn = Right(TableName, Len(TableName) - InStr(1, TableName, "_")) & "_id"
                Else
                    LogError(New Exception($"This is not a process specific table: {TableName}"))
                    Return ReturnCodes.UDBS_ERROR
                End If
                ' Check to make sure that this is the proper record/record id
                If RecordID <> KillNullInteger(OneRecord(0)(IDColumn)) Then
                    LogError(New Exception($"Incorrect id: RecordID={RecordID}, IDColumn.Value={KillNull(OneRecord(0)(IDColumn))}"))
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Make sure that the calling function is not trying to modify a core table
                If IsCoreTable(TableName) Then
                    logger.Error($"{fncName} Cannot use function to modify UDBS Core Table")
                    Return ReturnCodes.UDBS_ERROR
                End If

                Dim dbCommand As String
                Dim FieldItem As DataColumn

                dbCommand = "UPDATE " & TableName & " SET "
                For Each FieldItem In OneRecord.Columns
                    If FieldItem.ColumnName <> IDColumn Then
                        If FieldItem.DataType = GetType(String) Or FieldItem.DataType = GetType(Date) Then
                            dbCommand = dbCommand & FieldItem.ColumnName & " = '" &
                                        KillNull(OneRecord(0)(FieldItem.ColumnName)) & "', "
                        Else
                            dbCommand = dbCommand & FieldItem.ColumnName & " = " &
                                        KillNull(OneRecord(0)(FieldItem.ColumnName)) & ", "
                        End If
                    End If
                Next FieldItem

                ' Remove false comma
                dbCommand = Left$(dbCommand, Len(dbCommand) - 2)
                dbCommand = dbCommand & " WHERE " & IDColumn & " = " & CStr(RecordID)
                ExecuteNetworkQuery(dbCommand)

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ' Not used. Candidate for removal.
        Friend Function Utility_DeleteProcessSpecificTableRecord(TableName As String,
                                                                 ByRef RecordID As Integer) _
            As ReturnCodes

            Dim IDColumn As String

            ' Function deletes the specified record in the process specific table

            Try
                ' Make sure that the calling function is not trying to modify a core table
                If IsCoreTable(TableName) Then
                    LogError(New Exception("Cannot use function to modify UDBS Core Table"))
                    Return ReturnCodes.UDBS_ERROR
                End If

                If InStr(1, TableName, "_") <> 0 Then
                    IDColumn = Right(TableName, Len(TableName) - InStr(1, TableName, "_")) & "_id"
                Else
                    LogError(New Exception($"This is not a process specific table: {TableName}"))
                    Return ReturnCodes.UDBS_ERROR
                End If

                Dim dbCommand As String = "DELETE FROM " & TableName & " " &
                            "WHERE " & IDColumn & " = " & CStr(RecordID)
                ExecuteNetworkQuery(dbCommand)
                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function


        '**********************************************************************
        '* ItemList Functions
        '**********************************************************************

        ''' <summary>
        '''     Processes a string as Unicode data, converting it to ASCII and replacing all sequences of invalid characters
        '''     (non-ASCII)
        '''     with a single ? each. If the string contains more than one ? in a row, they will also be condensed into a single ?
        ''' </summary>
        ''' <param name="str">The input string which may contain invalid characters</param>
        ''' <returns>A string of equal or lesser length with invalid sequences replaced with ?</returns>
        Friend Shared Function Utility_ConvertStringToASCIICondenseInvalidCharacters(str As String) As String
            Return If(String.IsNullOrEmpty(str), String.Empty, Regex.Replace(
                Encoding.ASCII.GetString(
                    Encoding.Convert('Convert text
                                     Encoding.Unicode, 'From Unicode
                                     Encoding.ASCII, 'To ASCII
                                     Encoding.Unicode.GetBytes(str)     'Using the unicode-encoded bytes of the string
                                     )
                    ), "\?\?+", "?"))
            'Replace all instances of 2 or more ? with a single ?
        End Function

        Friend Function ItemList_GetRevisionCount(ProcessName As String,
                                                  ProductNumber As String,
                                                  ProductRelease As Integer,
                                                  Stage As String,
                                                  ByRef Revisions As Integer,
                                                  ByRef RevisionId As Integer) _
            As ReturnCodes
            ' Function returns the number of revisions of a specified Product, Product Release & Stage
            Revisions = 0
            Dim ItemListRevTable As String
            Dim rsTemp As DataTable = Nothing
            Dim sqlQuery As String

            Try
                ItemListRevTable = LCase(Trim(ProcessName)) & "_itemlistrevision"

                sqlQuery = "SELECT TOP 1 itemlistrev_id, itemlistrev_revision FROM " & ItemListRevTable &
                           " with(nolock) WHERE itemlistrev_product_id IN (SELECT product_id FROM " &
                           mProductTable & " with(nolock) WHERE product_number = '" & ProductNumber & "' " &
                           "AND product_release = " & CStr(ProductRelease) & ") " &
                           "AND itemlistrev_stage = '" & Stage & "' " &
                           "ORDER BY itemlistrev_revision DESC "
                OpenNetworkRecordSet(rsTemp, sqlQuery)

                If (If(rsTemp?.Rows?.Count, 0)) > 0 AndAlso Not IsDBNull(rsTemp(0)(1)) Then
                    ' Stage found

                    RevisionId = KillNullInteger(rsTemp(0)(0))
                    Revisions = KillNullInteger(rsTemp(0)(1))
                    Return ReturnCodes.UDBS_OP_SUCCESS
                Else
                    ' This stage does not exist
                    LogError(New Exception($"Stage not found: {Stage}"))
                    Revisions = -1
                    Return ReturnCodes.UDBS_ERROR
                End If
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try

        End Function

        ' Not used. Candidate for removal.
        Friend Function ItemList_ReleaseItemListFromDebugMode(ProcessName As String,
                                                              ProductNumber As String,
                                                              ProductRelease As Integer,
                                                              Stage As String) _
            As ReturnCodes
            ' Function releases a Stage from debug mode to first revision.
            Dim Revision As Integer
            Revision = 0

            Try
                ' Load the specified itemlist, asking for the latest revision
                Dim ITEMLIST As New CItemlist
                If ITEMLIST.LoadItemList(ProcessName, ProductNumber, ProductRelease, Stage, Revision) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                If ITEMLIST.Revision <> 0 Then
                    ' This list is already releaseed
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Release the list
                If ITEMLIST.ReleaseItemListFromDebugMode() <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                ITEMLIST = Nothing

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ' Not used. Candidate for removal.
        Friend Function Product_GetStageCount(ProcessName As String,
                                              ProductNumber As String,
                                              ProductRelease As Integer,
                                              ByRef StageCount As Integer) _
            As ReturnCodes
            ' Function returns the count of Stages for a specified Product/Release
            Dim rsTemp As DataTable = Nothing
            Try
                If Product_GetStageList(ProcessName, ProductNumber, ProductRelease, rsTemp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    StageCount = -1
                    Return ReturnCodes.UDBS_ERROR
                End If

                StageCount = rsTemp.Rows.Count
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try

        End Function

        ' Not used, candidate for removal.
        Friend Function Product_GetStageList(ProcessName As String,
                                             ProductNumber As String,
                                             ProductRelease As Integer,
                                             ByRef StageList As DataTable) _
            As ReturnCodes
            ' Function returns a list of Stages available for the specified Product
            Try
                Dim sqlQuery As String
                Dim ItemListRevisionTable As String

                ItemListRevisionTable = LCase(Trim(ProcessName)) & "_itemlistrevision"

                sqlQuery = "SELECT DISTINCT itemlistrev_stage FROM " & ItemListRevisionTable &
                           " with(nolock) WHERE itemlistrev_product_id IN " &
                           "(SELECT product_id FROM " & mProductTable &
                           " WHERE product_number = '" & ProductNumber & "' " &
                           "AND product_release=" & CStr(ProductRelease) & ") " &
                           "ORDER BY itemlistrev_stage"
                If QueryNetworkDB(sqlQuery, StageList) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Throw New Exception($"Error querying for stages for product: {ProductNumber}")
                ElseIf StageList?.Rows.Count >= 0 Then
                    Return ReturnCodes.UDBS_OP_SUCCESS
                Else
                    ' There is a problem finding any stages
                    LogError(New Exception($"No Stage found for product: {ProductNumber}"))
                    Return ReturnCodes.UDBS_ERROR
                End If
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        '**********************************************************************
        '* Product Support Functions
        '**********************************************************************

        ' Function returns the products that are members of the specified family
        ' Not used. Candidate for removal.
        Friend Function Product_GetProductFamily(ProductFamily As String,
                                                 ByRef ProductInfo As DataTable) As ReturnCodes
            Try
                Dim sqlQuery As String =
                           "SELECT * FROM " & mProductTable &
                           " with(nolock) WHERE product_family_id IN " &
                           "(SELECT family_id FROM " & mProductFamilyTable &
                           " with(nolock) WHERE family_name = '" & Trim(ProductFamily) & "') " &
                           "ORDER BY product_number, product_release"
                Return QueryNetworkDB(sqlQuery, ProductInfo)
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function


        ' Function returns a list of the existing product families
        Friend Function Product_GetProductFamilyList(ByRef ProductFamilies As DataTable) As ReturnCodes
            Try
                Dim sqlQuery As String =
                    $"SELECT DISTINCT family_name FROM {mProductFamilyTable} with(nolock)"
                Return QueryNetworkDB(sqlQuery, ProductFamilies)
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        Public Function Product_GetUnitVariance(ProductNumber As String,
                                                SerialNumber As String,
                                                ByRef OraclePN As String,
                                                ByRef CataloguePN As String,
                                                ByRef Variance As String) As ReturnCodes
            ' Function returns a list of the existing product families
            Dim sSQL As String
            Dim rsTemp As New DataTable

            Dim lUnitID As Integer, tmpStr As String, arrStr() As String

            Try
                sSQL = "SELECT * FROM product with(nolock) , unit  with(nolock) " &
                   "WHERE product_id=unit_product_id " &
                   "AND product_number='" & ProductNumber & "' " &
                   "AND unit_serial_number='" & SerialNumber & "' "
                OpenNetworkRecordSet(rsTemp, sSQL)
                If (If(rsTemp?.Rows?.Count, 0)) = 0 Then
                    logger.Error($"{SerialNumber} Serial number not found.")
                    Return ReturnCodes.UDBS_ERROR
                Else
                    If (If(rsTemp?.Rows?.Count, 0)) > 1 Then
                        logger.Error("Duplicate serial number found." & Chr(13) & "Please contact engineering staff.")
                        Return ReturnCodes.UDBS_ERROR
                    Else

                        lUnitID = KillNullInteger(rsTemp(0)("unit_id"))
                        OraclePN = KillNull(rsTemp(0)("product_number"))        ' default value
                        CataloguePN = KillNull(rsTemp(0)("product_catalogue_number"))      ' default value
                        Variance = ""
                    End If
                End If
                rsTemp = Nothing

                sSQL = "SELECT * FROM udbs_unit_details with(nolock) , udbs_product_group  with(nolock) " &
                       "WHERE ud_pg_product_group=pg_product_group " &
                       "AND ud_pg_sequence=pg_sequence " &
                       "AND ud_identifier='PRD_VAR' " &
                       "AND ud_unit_id=" & lUnitID
                OpenNetworkRecordSet(rsTemp, sSQL)
                If (If(rsTemp?.Rows?.Count, 0)) = 0 Then
                    ' that is fine, non-RoHS stuffs, nothing need to be updated.
                Else
                    If (If(rsTemp?.Rows?.Count, 0)) > 1 Then
                        logger.Error("Multiple variance records found." & Chr(13) & "Please contact engineering staff.")
                        Return ReturnCodes.UDBS_ERROR

                    Else
                        tmpStr = KillNull(rsTemp(0)("pg_string_value"))
                        arrStr = Split(tmpStr, ",")
                        If UBound(arrStr) < 2 Then
                            logger.Error(
                                "Invalid variance information found." & Chr(13) & "Please contact engineering staff.")
                            Return ReturnCodes.UDBS_ERROR
                        End If
                        OraclePN = arrStr(0)
                        CataloguePN = arrStr(1)
                        If arrStr(2) <> "-1" Then Variance = arrStr(2)
                    End If
                End If
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try

        End Function

        Public Function Product_GetPartIdentifier(OraclePN As String,
                                                  ByRef PartIdentifier As String) As ReturnCodes
            ' Function returns a list of the existing product families
            Dim sSQL As String, rsTemp As New DataTable

            Try
                sSQL = "SELECT prdgrp_product_number " &
                   "FROM udbs_prdgrp with(nolock), udbs_product_group with(nolock) " &
                   "WHERE prdgrp_product_group=pg_product_group " &
                   "AND pg_product_group LIKE '%_variance' " &
                   "AND pg_string_value LIKE '" & OraclePN & ",%' "
                OpenNetworkRecordSet(rsTemp, sSQL)
                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                    PartIdentifier = KillNull(rsTemp(0)("prdgrp_product_number"))
                    Return ReturnCodes.UDBS_OP_SUCCESS
                Else
                    ' try product table
                    rsTemp = Nothing
                    sSQL = "SELECT product_number FROM product with(nolock) " &
                           "WHERE product_number='" & OraclePN & "' "
                    OpenNetworkRecordSet(rsTemp, sSQL)
                    If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                        PartIdentifier = KillNull(rsTemp(0)("product_number"))
                        Return ReturnCodes.UDBS_OP_SUCCESS
                    Else
                        LogError(New Exception($"Cannot find Part Identifier for the Oracle PN '{OraclePN}"))
                        Return ReturnCodes.UDBS_ERROR
                    End If
                End If
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try

        End Function

        Friend Function Product_UnitExists(ProductNumber As String,
                                           Release As Integer,
                                           SerialNumber As String) _
            As Boolean
            ' does the unit exist in the database
            Dim Temp As New CProduct

            Try
                If Temp.GetProduct(ProductNumber, Release) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' product not found
                    Return False
                Else
                    Return Temp.UnitExists(SerialNumber)
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return False
            End Try

        End Function


        '**********************************************************************
        '* UDBS System Support Functions
        '**********************************************************************

        Friend Function UDBS_GetUDBSProcessID(ProcessName As String,
                                              ByRef UDBSProcessID As Integer) _
            As ReturnCodes
            ' Function returns the UDBS id of the specified process
            Dim rsTemp As New DataTable
            Try
                If UDBS_GetUDBSProcessInfo(ProcessName, rsTemp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                UDBSProcessID = KillNullInteger(rsTemp(0)("udbs_process_id"))
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try
        End Function


        Friend Function UDBS_GetUDBSProcessInfo(ProcessName As String,
                                                ByRef UDBSProcessInfo As DataTable) _
            As ReturnCodes
            ' Function returns the UDBS process info of the specified process

            Dim sqlQuery As String

            Try
                sqlQuery = "SELECT * FROM " & mUDBSProcessesTable &
                           " with(nolock) WHERE udbs_process_name = '" & ProcessName & "' "
                If QueryNetworkDB(sqlQuery, UDBSProcessInfo) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Throw New Exception("Error querying for process information.")
                End If

                If (If(UDBSProcessInfo?.Rows?.Count, 0)) <> 1 Then
                    LogError(New Exception("Incorrect number of records returned."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function


        ' Function returns a list of the UDBS processes
        Friend Function UDBS_GetUDBSProcessList(ByRef Processes As DataTable) As Integer
            Try
                Dim sqlQuery As String =
                    $"SELECT * FROM {mUDBSProcessesTable}  with(nolock)"
                Return QueryNetworkDB(sqlQuery, Processes)
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        'Lookup the disposition of the unit with given product and serial number.  Populate the
        'the Dispoition return parameter if found.  The return parameter is set to the empty
        'String if the disposition has not been set  for this unit.  Returns an error if the
        'database is not in the expected state.
        Friend Function UnitDetails_GetDisposition(ProductNumber As String,
                                                SerialNumber As String,
                                                ByRef Disposition As String,
                                                ByRef EmployeeNumber As String,
                                                ByRef Comment As String) As ReturnCodes

            'A new row is inserted in udbs_unit_details each time the unit's disposition is modified.
            'The value of the ud_integer_field is incremented on each of these inserts.  The inner join
            'in this SQL is finding the largest value of ud_integer_field so that only the most recent
            'value will be returned from the outer select.
            'The product and unit tables are needed to find the unit corresponding to the given part
            'and serial numbers.
            'This query returns nothing when the unit's disposition has never been set for this unit.
            Dim sql As String =
                "SELECT pg.pg_string_value AS UnitDisposition, " &
                "       LEFT(ud.ud_string_value, CHARINDEX(',', ud.ud_string_value) - 1) as EmployeeNumber, " &
                "       RIGHT(ud.ud_string_value, LEN(ud.ud_string_value) - CHARINDEX(',', ud.ud_string_value)) as Comment " &
                "  FROM udbs_product_group pg, product p, unit u, udbs_unit_details ud " &
               $" WHERE p.product_number = '{ProductNumber}' " &
               $"   AND u.unit_serial_number = '{SerialNumber}' " &
                "   AND pg.pg_product_group = 'UNIT_DISP' " &
                "   AND ud.ud_identifier = 'CURRENT' " &
                "   AND ud.ud_unit_id = u.unit_id " &
                "   AND ud.ud_pg_product_group = pg.pg_product_group " &
                "   AND ud.ud_pg_sequence = pg.pg_sequence "
            Dim result = New DataTable
            OpenNetworkRecordSet(result, sql)
            If result?.Rows?.Count > 1 Then
                LogError(New Exception("Duplicate serial number found. Please contact engineering staff."))
                logger.Error($"UnitDetails_GetDisposition: Product '{ProductNumber}', serial number '{SerialNumber}'. {Err.Description}")
                Return ReturnCodes.UDBS_ERROR
            End If

            'The return parameter is the database value or the empty string when the database does
            'not have a value.
            If result?.Rows?.Count = 1 Then
                Disposition = KillNull(result(0)("UnitDisposition"))
                EmployeeNumber = KillNull(result(0)("EmployeeNumber"))
                Comment = KillNull(result(0)("Comment"))
            Else
                Disposition = ""
                EmployeeNumber = ""
                Comment = ""
            End If

            Return ReturnCodes.UDBS_OP_SUCCESS
        End Function

        Friend Function UnitDetails_SetDisposition(ProductNumber As String,
                                                SerialNumber As String,
                                                Disposition As String,
                                                EmployeeNumber As String,
                                                Comment As String
                                                ) As ReturnCodes

            If String.IsNullOrWhiteSpace(EmployeeNumber) Then
                Throw New InvalidOperationException("Modifying unit disposition requires a valid EmployeeNumber")
            End If

            'The unit's disposition is stored in two tables.  Each unit has a value, so the udbs_unit_details
            'table is used.  This table points to values in the udbs_product_group table, so that table is used
            'as well.
            'The udbs_product_group creates the 'UNIT_DISP' group.  The pg_string_value column is used to specify
            'the value.  A new row is added to udbs_unit_details each time that the value is changed.  This keeps
            'the full history of the unit in the database.  The ud_integer_value column is used to record the order
            'of this history.
            'The ud_string_value is a comment describing the employeeNumber that made the change, along with an
            'optional comment.  These two strings are comma-separated (employeeNumber first), and stored in the
            'ud_string_value column.

            'There are two parts to this transaction.
            '  1) Make sure the Disposition exists in the udbs_product_group table.
            '  2) Create an entry in the udbs_unit_details table that points to the corresponding group.

            Using transaction = BeginNetworkTransaction()

                'The types of unit disposition are defined in the code, but must exist in the udbs_product_group
                'table.  This SELECT happens on all sets, the INSERT and max() happen only if the value is missing.
                Dim sql As String =
                    "SELECT pg_sequence " &
                    "  FROM udbs_product_group " &
                    " WHERE pg_product_group = 'UNIT_DISP' " &
                   $"   AND pg_string_value = '{Disposition}'"
                Dim result = New DataTable
                OpenNetworkRecordSet(result, sql, transaction)
                If result?.Rows?.Count < 1 Then
                    'Create a new product group for this enumerator.
                    sql =
                        "INSERT INTO udbs_product_group " &
                        "    ( pg_product_group, pg_sequence, pg_description, pg_string_value ) " &
                       $" SELECT 'UNIT_DISP', coalesce(max(pg_sequence), 0) + 1, concat( 'Unit disposition for ', '{Disposition}', ' items' ), '{Disposition}'" &
                        "   FROM udbs_product_group " &
                        "  WHERE pg_product_group = 'UNIT_DISP' "
                    ExecuteNetworkQuery(sql, transaction)
                End If

                'There are two phases to updating the unit details table.
                '  1) Update the existing "most current" entry.
                '  2) Create a new entry that is now the "most current".
                'The new entry will have a ud_integer_value that is one higher than the previous "most current".  The
                'first query is to find that value.  If there isn't an existing "most current", then ud_integer_value
                'starts at 1.  The second query is to insert the new record.

                'Get highest existing ud_integer_value, or learn that there are currently no entries
                'for this unit.
                sql =
                    "SELECT ud.ud_unit_id, ud.ud_integer_value as max " &
                    "  FROM udbs_unit_details ud, udbs_product_group pg, product p, unit u " &
                    " WHERE ud.ud_identifier = 'CURRENT' " &
                   $"   AND p.product_number = '{ProductNumber}' " &
                   $"   AND u.unit_serial_number = '{SerialNumber}' " &
                    "   AND pg.pg_product_group = 'UNIT_DISP' " &
                    "   AND ud.ud_unit_id = u.unit_id " &
                    "   AND ud.ud_pg_product_group = pg.pg_product_group " &
                    "   AND ud.ud_pg_sequence = pg.pg_sequence "
                result.Clear()
                OpenNetworkRecordSet(result, sql, transaction)
                If result?.Rows?.Count > 1 Then
                    LogError(New Exception("Duplicate serial number found. Please contact engineering staff."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                'Find values for the existing entry (if any).  The next entry will use a ud_integer_value that is
                'one more.  The unit_id will be used to improve performance of the INSERT.
                Dim last_int_value = 0
                Dim unit_id = 0
                If result?.Rows?.Count = 1 Then
                    last_int_value = KillNullInteger(result(0)("max"))
                    unit_id = KillNullInteger(result(0)("ud_unit_id"))
                End If

                'The unit details table uses ud_identifier = 'CURRENT' to track the most recent entry.
                'We're about to insert a new "most recent" entry, so update existing one.
                If last_int_value > 0 Then
                    sql =
                        "UPDATE udbs_unit_details " &
                       $"   SET ud_identifier = 'HISTORY {last_int_value}' " &
                       $" WHERE ud_unit_id = {unit_id} " &
                       $"   AND ud_integer_value = {last_int_value} "
                    'ExecuteNetworkQuery updates the transaction and throws an exception on error,
                    'so there Is Nothing to check here.
                    ExecuteNetworkQuery(sql, transaction)
                End If

                'Increment the value that is used for sequencing history, and insert the new row.
                Dim next_int_value = last_int_value + 1

                'Insert a new record in udbs_unit_details that points to the appropriate record in
                'the pg_product_group table.
                sql =
                    "INSERT INTO udbs_unit_details " &
                    "     ( ud_unit_id,  ud_identifier, ud_integer_value, ud_pg_product_group, ud_pg_sequence, ud_string_value ) " &
                   $" SELECT u.unit_id, 'CURRENT',      {next_int_value}, pg.pg_product_group, pg.pg_sequence, '{EmployeeNumber},{Comment}'" &
                    "   FROM udbs_product_group pg, product p, unit u " &
                    "  WHERE pg.pg_product_group  = 'UNIT_DISP' " &
                   $"    AND pg.pg_string_value   = '{Disposition}' " &
                   $"    AND p.product_number     = '{ProductNumber}' " &
                   $"    AND u.unit_serial_number = '{SerialNumber}' "
                'ExecuteNetworkQuery updates the transaction and throws an exception on error,
                'so there Is Nothing to check here.
                ExecuteNetworkQuery(sql, transaction)
            End Using ' End Transaction

            Return ReturnCodes.UDBS_OP_SUCCESS
        End Function

        ''' <summary>
        '''     Compares the attribute values that are currently stored in the database with the attributes that are
        '''     currently reported by connected instruments.  Return a new dictionary with only the attributes that
        '''     need to be updated.
        ''' </summary>
        ''' <param name="dbAttrs">The collection of attribute currently stored in the database.</param>
        ''' <param name="instrAttrs">The collection of attributes reported by the currently attached instruments.</param>
        ''' <returns>
        '''     A new dictionary containing only the differences between the currently stored attribute values and
        '''     the values that are currently reported by the instruments.
        ''' </returns>
        Friend Function CalculateUpdatedAttrs(dbAttrs As Dictionary(Of String, String), instrAttrs As Dictionary(Of String, String)) As Dictionary(Of String, String)
            Dim updateAttrs = New Dictionary(Of String, String)

            'If an attribute is found in the database, but no longer has a value in the instrument,
            'then the database value is updated to NULL.
            For Each rm_key As String In dbAttrs.Keys.Except(instrAttrs.Keys)
                updateAttrs(rm_key) = Nothing
            Next rm_key

            'Attribute values that are currently reported by the instruments should be inserted into the
            'database if they are not already there.
            For Each kvp As KeyValuePair(Of String, String) In instrAttrs
                If Not dbAttrs.ContainsKey(kvp.Key) OrElse dbAttrs.Item(kvp.Key) <> kvp.Value Then
                    updateAttrs.Add(kvp.Key, kvp.Value)
                End If
            Next

            Return updateAttrs
        End Function

        ''' <summary>
        '''     Compares the attribute values that are currently stored in the database with the attributes that are
        '''     currently reported by connected instruments.  Return a new dictionary with only the attributes that
        '''     need to be updated.
        ''' </summary>
        ''' <param name="dbAttrs">The collection of attribute currently stored in the database.</param>
        ''' <param name="instrAttrs">The collection of attributes reported by the currently attached instruments.</param>
        ''' <returns>
        '''     A new dictionary containing only the differences between the currently stored attribute values and
        '''     the values that are currently reported by the instruments.
        ''' </returns>
        Friend Function CalculateUpdatedAttrs(dbAttrs As Dictionary(Of String, Dictionary(Of String, String)), instrAttrs As Dictionary(Of String, Dictionary(Of String, String))) As Dictionary(Of String, Dictionary(Of String, String))
            Dim updateAttrs = New Dictionary(Of String, Dictionary(Of String, String))

            'If an instrument is found in the database but no longer has a value in the hardware,
            'then all database attributes are updated to NULL.
            For Each rm_instr As String In dbAttrs.Keys.Except(instrAttrs.Keys)
                Dim rm_attrs = New Dictionary(Of String, String)
                For Each rm_key As String In dbAttrs.Item(rm_instr).Keys
                    rm_attrs.Add(rm_key, Nothing)
                Next rm_key

                updateAttrs.Add(rm_instr, rm_attrs)
            Next rm_instr

            'All other attributes from the current instruments should be merged with the currently stored
            'values.
            For Each kvp As KeyValuePair(Of String, Dictionary(Of String, String)) In instrAttrs
                Dim db_attrs = If(dbAttrs.ContainsKey(kvp.Key), dbAttrs.Item(kvp.Key), New Dictionary(Of String, String))
                Dim update_attrs = CalculateUpdatedAttrs(db_attrs, kvp.Value)

                'The instrument should only go into the update set when there is something to be updated.
                If update_attrs.Count <> 0 Then
                    updateAttrs.Add(kvp.Key, update_attrs)
                End If
            Next kvp

            Return updateAttrs
        End Function

        ''' <summary>
        '''     A utility function that converts from the DataTable returned from the database into the
        '''     multi-level dictionary used to represent instrument attributes in the rest of the application.
        ''' </summary>
        ''' <param name="data">The DataTable that is returned from UDBS.</param>
        ''' <param name="InstrColName">The DataTable's column name for the instrument name.</param>
        ''' <param name="KeyColName">The DataTable's column name for the attribute name.</param>
        ''' <param name="ValueColName">The DataTable's column name for the attribute value.</param>
        ''' <returns>
        '''     A two-level dictionary containing the corresponding instrument attribute values.  The first level
        '''     of key is the name of the instrument instance.  The second level of key is the name of the attribute
        '''     for the corresponding instrument.
        ''' </returns>
        Private Function ConvertToInstrumentAttrs(data As DataTable, Optional InstrColName As String = "InstrName", Optional KeyColName As String = "AttrKey", Optional ValueColName As String = "AttrValue") As Dictionary(Of String, Dictionary(Of String, String))
            Dim result = New Dictionary(Of String, Dictionary(Of String, String))
            For Each row As DataRow In data?.Rows
                Dim instr = row(InstrColName).ToString

                'Insert a new dictionary if this is the first time the instrument has been noticed.
                If Not result.ContainsKey(instr) Then
                    result.Add(instr, New Dictionary(Of String, String))
                End If

                'Put the attribute key and value into the result dictionary.
                result.Item(instr)(row(KeyColName).ToString) = row(ValueColName).ToString
            Next row

            Return result
        End Function

        ''' <summary>
        '''     Compare the given instrument attributes with values that are currently stored in UDBS.  Insert
        '''     new rows to store values for any attributes that have changed.
        ''' </summary>
        ''' <param name="StationName">The name of the station to which the instruments are attached.</param>
        ''' <param name="InstrAttrs">
        '''     A two-level dictionary containing the corresponding instrument attribute values.  The first level
        '''     of key is the name of the instrument instance.  The second level of key is the name of the attribute
        '''     for the corresponding instrument.
        ''' </param>
        Public Sub UpdateInstrumentAttrs(TestDataProcessID As String, StationName As String, InstrAttrs As Dictionary(Of String, Dictionary(Of String, String)))

            'CREATE TABLE station_attributes
            '  ( station_id         INT          IDENTITY( 1, 1 ),
            '    station_name       NVARCHAR(256),
            '    station_component  NVARCHAR(256),
            '    station_attr_key   NVARCHAR(256),
            '    station_attr_value NVARCHAR(256),
            '    station_updated_at DATETIME     DEFAULT CURRENT_TIMESTAMP,
            '    station_updated_by NVARCHAR(256) );

            'nvarchar instead of varchar
            'station-name of 256 instead of 50
            'compound index on station_name + station_udpated_at

            'These locals are used to make sure the same values are used in the SQL expression and the
            'code that looks at the result.
            Dim InstrName = "InstrName"
            Dim AttrKey = "AttrKey"
            Dim AttrValue = "AttrValue"

            Using transaction = BeginNetworkTransaction()

                'Read the end time of the test process entry.  That record has already been inserted, if we didn't read
                'that time, then the insertion time of this record would be after the test process where we observed the
                'new attribute values.  We don't want it to look like the instrument attributes changed after the test
                'process, so just collect the same time.  The time is read in a format that perserves milliseconds when
                'inserting into the station_attributes table.
                Dim dt = New DataTable
                OpenNetworkRecordSet(dt, $"SELECT FORMAT(process_end_date, 'yyyy-MM-dd HH:mm:ss.fff') FROM testdata_process WHERE process_id = {TestDataProcessID}", transaction)
                Dim processEndTime = dt.Rows(0).Item(0)

                'Lookup the current value of all instrument attributes associated with this station.
                Dim sql = "WITH most_recent_sas " &
                    "AS ( SELECT DISTINCT sa.station_component, station_attr_key, MAX(sa.station_updated_at) OVER (PARTITION BY sa.station_component, sa.station_attr_key) AS station_updated_at " &
                    "       FROM station_attributes sa " &
                   $"      WHERE sa.station_name = '{StationName}' " &
                    "   ) " &
                   $" SELECT sa.station_id, sa.station_component as {InstrName}, sa.station_attr_key as {AttrKey}, sa.station_attr_value as {AttrValue} " &
                    "   FROM station_attributes sa " &
                    "   JOIN most_recent_sas mrs ON mrs.station_component  = sa.station_component " &
                    "                           AND mrs.station_attr_key   = sa.station_attr_key " &
                    "                           AND mrs.station_updated_at = sa.station_updated_at " &
                    "  WHERE sa.station_attr_value IS NOT NULL"

                dt = New DataTable
                OpenNetworkRecordSet(dt, sql, transaction)
                Dim dbAttrs = ConvertToInstrumentAttrs(dt, InstrName, AttrKey, AttrValue)

                'Compare the currently available attributes with what is currently stored in the database.  Create
                'a dictionary with only the values that need to be updated.
                Dim updatedAttrs = CalculateUpdatedAttrs(dbAttrs, InstrAttrs)

                'If needed, update the database with the latest instrument attribute values.
                If updatedAttrs.Count > 0 Then
                    Dim first = True
                    sql = "INSERT INTO station_attributes ( station_name, station_component, station_attr_key, station_attr_value, station_updated_at, station_updated_by ) VALUES "

                    For Each instr_kvp As KeyValuePair(Of String, Dictionary(Of String, String)) In updatedAttrs
                        For Each kvp As KeyValuePair(Of String, String) In instr_kvp.Value
                            If first Then
                                first = False
                            Else
                                sql += ", "
                            End If

                            'When the attribute is deleted, a row in inserted in the database with the value set
                            'to NULL.
                            Dim db_value = If(kvp.Value Is Nothing, "NULL", $"'{kvp.Value}'")

                            sql += $"( '{StationName}', '{instr_kvp.Key}', '{kvp.Key}', {db_value}, '{processEndTime}', 'testdata_process_id:{TestDataProcessID}' )"
                        Next kvp
                    Next instr_kvp

                    ExecuteNetworkQuery(sql, transaction)
                End If 'End of updating instrument attribute values.
            End Using ' End Transaction
        End Sub

        ''' <summary>
        '''     Query UDBS to lookup and return the value of all attributes that were associated with the
        '''     given station at the given time.
        ''' </summary>
        ''' <param name="StationName">The name of the station to query.</param>
        ''' <param name="Timestamp">
        '''     Optional parameter to specify the maximum station revision that should be included in results.  Returns
        '''     the current value of attributes when revision is not provided.
        ''' </param>
        ''' <returns>
        '''     A two-level dictionary containing the corresponding instrument attribute values.  The first level
        '''     of key is the name of the instrument instance.  The second level of key is the name of the attribute
        '''     for the corresponding instrument.
        ''' </returns>
        ''' <remarks>Not used. Candidate for removal.</remarks>
        Friend Function GetInstrumentAttrs(StationName As String, Optional Timestamp As String = Nothing) As Dictionary(Of String, Dictionary(Of String, String))
            'These locals are used to make sure the same values are used in the SQL expression and the
            'code that looks at the result.
            Dim InstrName = "InstrName"
            Dim AttrKey = "AttrKey"
            Dim AttrValue = "AttrValue"

            Dim sql = "WITH most_recent_sas " &
                "AS ( SELECT DISTINCT sa.station_component, station_attr_key, MAX(sa.station_updated_at) OVER (PARTITION BY sa.station_component, sa.station_attr_key) AS station_updated_at " &
                "       FROM station_attributes sa " &
               $"      WHERE sa.station_name = '{StationName}' "

            'If the function was called with a value for the optional StationRev parameter, then limit the
            'search to that revision.  If the value is not provided then this clause is not needed, the
            'MAX(sa.station_id) part of the SQL request will search the entire table.
            If Not String.IsNullOrEmpty(Timestamp) Then
                sql += $"AND sa.station_updated_at <= '{Timestamp}' "
            End If

            sql += " ) " &
                  $"SELECT sa.station_component AS {InstrName}, sa.station_attr_key AS {AttrKey}, sa.station_attr_value AS {AttrValue} " &
                   "  FROM station_attributes sa " &
                   "  JOIN most_recent_sas mrs ON mrs.station_component  = sa.station_component " &
                   "                          AND mrs.station_attr_key   = sa.station_attr_key " &
                   "                          AND mrs.station_updated_at = sa.station_updated_at " &
                   " WHERE sa.station_attr_value IS NOT NULL "

            Dim dt = New DataTable
            OpenNetworkRecordSet(dt, sql)

            Return ConvertToInstrumentAttrs(dt, InstrName, AttrKey, AttrValue)
        End Function

        ''' <summary>
        ''' Sets the stationID.
        ''' </summary>
        ''' <param name="aStationName">StationID as string</param>
        Friend Shared Sub Utility_SetTemporaryStationName(aStationName As String)
            stationName = aStationName
        End Sub

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
