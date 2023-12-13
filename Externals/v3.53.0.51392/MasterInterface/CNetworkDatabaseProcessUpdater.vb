Imports System.IO
Imports System.Text

Namespace MasterInterface

	''' <summary>
	''' This class uploads all process header and result information to the network DB
	''' and optionally, upon succesful completion remove the old data from the local DB.
	''' 
	''' This code was moved from the <see cref="CProcessInstance"/> class to group all the code
	''' that is relevant to the DB update in a single location.
	''' </summary>
	Friend Class CNetworkDatabaseProcessUpdater

#Region "Properties containing information about what process is being updated."
		''' <summary>The process being updated.</summary>
		Private ReadOnly Property Process As CProcessInstance
#End Region

#Region "Parameters And options regarding the update operation."
		''' <summary>Whether or not to remove the entries from the local DB on success.</summary>
		Private ReadOnly Property RemoveLocalCopy As Boolean
#End Region

		''' <param name="RemoveLocalCopy">Whether or not to remove local copy on success.</param>
		Public Sub New(
					  Process As CProcessInstance,
					  RemoveLocalCopy As Boolean)
			If (Process Is Nothing) Then
				Throw New NullReferenceException("Process not provided.")
			End If

			Me.Process = Process
			Me.RemoveLocalCopy = RemoveLocalCopy
		End Sub

		''' <summary>
		''' Function to upload all process header and result information to network DB,
		''' and optionally, upon succesful completion, remove the old data from the local DB.
		''' </summary>
		''' <param name="transactionScope">Database transaction.</param>
		''' <returns>Operation result as Return Code.</returns>
		Public Function Execute(Optional ByRef transactionScope As ITransactionScope = Nothing) As ReturnCodes

			Const NOLOCK = " with(nolock) "
			Const BulkCopyFactor = 2 ' Magic number to start using Bulk Copy

			Dim sqlQuery As String
			Dim FieldItem As DataColumn
			Dim rsLocalProcess As New DataTable
			Dim rsLocalResult As New DataTable
			Dim rsNetworkProcess As New DataTable
			Dim rsNetworkResult As New DataTable
			Dim currentTime As Date = Date.MinValue
			Dim totalDuration As Integer
			Dim columnNames As List(Of String) = Nothing
			Dim transactionCreated As Boolean = False

			logger.Debug($"Updating Process {Process.Process}/{Process.ID} to Network UDBS DB. Removing local copy on success? {RemoveLocalCopy}.")

			Dim processQuery = $"SELECT * FROM {Process.ProcessTable} {{0}} WHERE process_id ={Process.ID}"

			Try
				rsLocalProcess = Process.Instance_RS

				' There is data to be transferred.
				' First get the Process ID, then transfer all related results.
				sqlQuery = String.Format(processQuery, NOLOCK)
				OpenNetworkRecordSet(rsNetworkProcess, sqlQuery)

				' check the network db for an already completed test with the same unit
				' When a process gets "stolen" by another station, the value of "Me.Status" becomes out-of-sync (in-memory vs. network DB)
				' so it is important to load it directly from the network DB. We cannot use "Me.Status" here.
				If CProcessInstance.GetProcessStatusEnum(KillNull(rsNetworkProcess(0)("process_status"))) = UdbsProcessStatus.COMPLETED Then
					BackupLocalDB(Path.GetDirectoryName(LocalDBPath), DateAndTime.Now.ToString("yyyy-MM-dd_") + "backup", False)
					Process.DeleteLocalProcess()
					LogErrorInDatabase(New Exception($"Error when trying to Update Network DB, a test instance is already completed for process: {Process.ID}. The process has been backed up and deleted from the Local DB"))
					If transactionScope IsNot Nothing Then
						transactionScope.HasError = True
					End If
					Return ReturnCodes.UDBS_ERROR
					End If

					If (If(rsNetworkProcess?.Rows?.Count, 0)) <> 1 Then
					' The network process instance no longer exists
					LogError(New Exception(
									"The process identification for this instance is not properly defined in the network database."))
					If transactionScope IsNot Nothing Then
						transactionScope.HasError = True
					End If
					Return ReturnCodes.UDBS_ERROR
				End If

				If transactionScope Is Nothing Then
					transactionScope = BeginNetworkTransaction()
					transactionCreated = True
				End If

				Try
					' Create result records on server
					rsNetworkResult = GetProcessResults(isNetworkDb:=True)

					' NB: Will exclude reading local results, if there already exists record for an item in the SQL Server
					'     already have results previously?
					rsLocalResult = GetProcessResults(
									isNetworkDb:=False,
									resultIdsToExclude:=GetColumnValuesAsCSV(rsNetworkResult, "result_itemlistdef_id"))

					'Adding new rows found in the local DB but not found in network DB yet.
					logger.Trace($"Local testdata_result table contains {rsLocalResult.Rows.Count} row(s).")
					Dim columns As String() = {"result_process_id",
									"result_itemlistdef_id",
									"result_value",
									"result_passflag",
									"result_stringdata",
									"result_blobdata_exists"}

					' In the following step, we will not need to update the records
					' we are about to insert into the network DB.
					Dim resultIdsCsvList = GetColumnValuesAsCSV(rsLocalResult, "result_itemlistdef_id")

					If rsLocalResult.AsEnumerable().Count < BulkCopyFactor Then
						For Each localResultRow In rsLocalResult.AsEnumerable()
							Dim columnValues = CreateNetworkResultRowData(columns, localResultRow)
							InsertNetworkRecord(columns, columnValues, Process.ResultTable, transactionScope)
						Next
					Else
						' Use Bulk-Copy for many rows
						BulkInsertNetwork(columns, rsLocalResult, Process.ResultTable, transactionScope)
					End If

					rsNetworkResult = Nothing
					rsLocalResult = Nothing

					' Move process result information
					rsLocalResult = GetProcessResults(isNetworkDb:=False)
					logger.Trace($"Local testdata_result table contains {rsLocalResult.Rows.Count} row(s).")
					If (If(rsLocalResult?.Rows?.Count, 0)) > 0 Then

						' There are associated results with this process instance.
						' The results here, would also include the ones inserted above, mostly blank
						rsNetworkResult = GetProcessResults(isNetworkDb:=True)

						' Hack mode condition used to occur on using older versions of the UDBS library and
						' Microsoft SQL server. The software is now accurate, more resilient, and the newer
						' version of SQL server does not cause this situation to arise.
						' This will soon be removed. See TMTD-229.
						If DetectHackModeCondition(rsLocalResult, rsNetworkResult) Then
							Throw New UDBSException($"Invalid state: The network database contains more results than the local database.")
						End If

						' We don't need to update records we have just pushed to the database in the
						' bulk insert.
						rsLocalResult = GetProcessResults(isNetworkDb:=False, resultIdsToExclude:=resultIdsCsvList)
						rsNetworkResult = GetProcessResults(isNetworkDb:=True, resultIdsToExclude:=resultIdsCsvList)

						If (If(rsNetworkResult?.Rows?.Count, 0)) < (If(rsLocalResult?.Rows?.Count, 0)) Then
							' There are more results in the network DB than in the local DB.
							' This should never happen, since all the network records are copied from the
							' local database.
							LogError(New Exception(
												"Number of records in the network database is less than the number in the local database."))
							' Rollback the transaction.
							transactionScope.HasError = True
							Return ReturnCodes.UDBS_ERROR
						End If

						columnNames =
									rsNetworkResult.Columns.Cast(Of DataColumn)().Select(Function(x) x.ColumnName).
										ToList()
						'NB: Using LINQ to perform inner join on 2 data tables
						Dim joinedLocalAndNet = rsLocalResult.AsEnumerable().
									Join(rsNetworkResult.AsEnumerable(),
											Function(rLoc) KillNullInteger(rLoc("result_itemlistdef_id")),
											Function(rNet) KillNullInteger(rNet("result_itemlistdef_id")),
											Function(drLoc, drNet) New With {.Local = drLoc, .NetDB = drNet})

						' Move the results to the network 
						For Each locNet In joinedLocalAndNet
							' we're in lock-step between the two databases so just update the network record
							locNet.NetDB("result_value") = locNet.Local("result_value")
							locNet.NetDB("result_passflag") = locNet.Local("result_passflag")
							locNet.NetDB("result_stringdata") = locNet.Local("result_stringdata")
							locNet.NetDB("result_blobdata_exists") = locNet.Local("result_blobdata_exists")
						Next

						If joinedLocalAndNet.Count < BulkCopyFactor Then

							For Each locNet In joinedLocalAndNet
								UpdateNetworkRecord(
												{"result_id"},
												columnNames.ToArray(),
												locNet.NetDB.ItemArray,
												Process.ResultTable,
												transactionScope)
							Next
						Else
							Dim matchedColumns As String() = {"result_id"}
							BulkUpdateNetwork(columnNames.ToArray(), matchedColumns, rsNetworkResult, Process.ResultTable, transactionScope)
						End If

						rsNetworkResult = Nothing
					End If
					' Close local results recordset
					rsLocalResult = Nothing

					' Update the network BLOBs.
					UpdateNetworkBlobs(transactionScope)

					' move process record
					For Each FieldItem In rsLocalProcess.Columns
						' Move each field value over
						If FieldItem.ColumnName <> "process_id" Then
							' Do not attempt to move the record_id, as it is an autonumber index in table
							rsNetworkProcess(0)(FieldItem.ColumnName) = rsLocalProcess(0)(FieldItem.ColumnName)
							If FieldItem.ColumnName = "process_status" AndAlso
											KillNull(rsLocalProcess(0)(FieldItem.ColumnName)) = "IN PROCESS" AndAlso
											RemoveLocalCopy Then
								' This process instance has not been completed...
								rsNetworkProcess(0)(FieldItem.ColumnName) = "TERMINATED"
								CUtility.Utility_GetServerTime(currentTime)
								If Not IsDBNull(rsLocalProcess(0)("process_start_date")) Then
									If IsDate(rsLocalProcess(0)("process_start_date")) Then
										totalDuration = CInt(DateDiff("s",
																			KillNullDate(
																				rsLocalProcess(0)("process_start_date")),
																			currentTime))
									End If
								End If
							End If
						End If
					Next FieldItem

					If currentTime <> Date.MinValue Then
						'process_end date has been rewritten, TERMINATED
						rsNetworkProcess(0)("process_end_date") = currentTime
						rsNetworkProcess(0)("process_total_duration") = totalDuration
					End If

					' Update the network db
					columnNames =
								rsNetworkProcess.Columns.Cast(Of DataColumn)().[Select](Function(x) x.ColumnName).
									ToList()

					UpdateNetworkRecord({"process_id"}, columnNames.ToArray(), rsNetworkProcess(0).ItemArray,
												Process.ProcessTable, transactionScope)

					' Close process recordsets
					rsNetworkProcess.Dispose()
					rsNetworkProcess = Nothing
					rsLocalProcess.Dispose()
					rsLocalProcess = Nothing

				Catch ex As Exception
					' Mark the transaction for a rollback, and throw the exception back.
					transactionScope.HasError = True
					Throw
				End Try

				rsLocalProcess = Nothing
				rsLocalResult = Nothing

				' Remove the local copy of the process instance?
				If RemoveLocalCopy Then
					logger.Debug($"Deleting local data for process id {Process.ID}")
					Try
						Process.DeleteLocalProcess()

						' If  the process and result tables are empty (ie, no other processes using them...)
						' Drop the local tables. (This ensures that we try to pick up table modifications at the network db)
						sqlQuery = "SELECT COUNT(process_id) FROM " & Process.ProcessTable
						If OpenLocalRecordSet(rsLocalProcess, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS Then
							Throw New Exception("Error querying for process count in local DB.")
						End If

						' Check on result table may not be necessary
						sqlQuery = "SELECT COUNT(result_id) FROM " & Process.ResultTable
						If OpenLocalRecordSet(rsLocalResult, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS Then
							Throw New Exception("Error querying for result count in local DB.")
						End If

						Dim localProcessCount = KillNullInteger(rsLocalProcess(0)(0))
						Dim localResultCount = KillNullInteger(rsLocalResult(0)(0))
						If localProcessCount = 0 And localResultCount = 0 Then
							'There are no active processes
							rsLocalProcess = Nothing
							rsLocalResult = Nothing
							If DropLocalTables(Process.Process) <> ReturnCodes.UDBS_OP_SUCCESS Then
								LogError(New Exception("An error occurred calling DropLocalTables."))
								transactionScope.HasError = True
								Return ReturnCodes.UDBS_ERROR
							End If
						Else
							rsLocalProcess = Nothing
							rsLocalResult = Nothing
						End If
					Catch ex As Exception
						LogError(New Exception($"Unexpected exception remove the local copy of the process instance: {ex.Message}", ex))
						transactionScope.HasError = True
						Return ReturnCodes.UDBS_ERROR
					End Try
				End If

				Return ReturnCodes.UDBS_OP_SUCCESS

			Catch ex As Exception
				transactionScope.HasError = True
				LogErrorInDatabase(ex)
				Return ReturnCodes.UDBS_ERROR
			Finally
				rsLocalProcess?.Dispose()
				rsLocalResult?.Dispose()
				rsNetworkProcess?.Dispose()
				rsNetworkResult?.Dispose()

				If transactionCreated Then
					transactionScope?.Dispose()
					transactionScope = Nothing
				End If
			End Try
		End Function

		''' <summary>
		''' Get the results associated with a given process (identified by name and ID).
		''' </summary>
		''' <param name="isNetworkDb">Whether we are looking for the results from the network or local DB.</param>
		''' <param name="resultIdsToExclude">
		''' Optionaly, we may want to exclude certain results. If this is the case, set this parameter to a comma-separated
		''' list of result IDs that will be injected in the SQL query.
		''' </param>
		''' <returns>The resulting results.</returns>
		''' <exception cref="Exception">If we fail to query the database.</exception>
		Private Function GetProcessResults(isNetworkDb As Boolean, Optional resultIdsToExclude As String = Nothing) As DataTable
			Dim queryStr As String = GenerateResultQuery(isNetworkDb, resultIdsToExclude)
			Dim resultSet As New DataTable()
			If (isNetworkDb) Then
				OpenNetworkRecordSet(resultSet, queryStr)
			Else
				If (OpenLocalRecordSet(resultSet, queryStr) <> ReturnCodes.UDBS_OP_SUCCESS) Then
					' The OpenNetworkRecordSet(...) and OpenLocalRecordSet(...) have a slightly different interface.
					' On error, OpenNetworkRecordSet(...) throws an exception, but OpenLocalRecordSet(...) returns
					' an error code. In order for this function to present an uniform signature, lets throw an
					' exception if the local query fails.
					Throw New UDBSException($"Failed to get local result set for Process {Process.Process}, {Process.ID}")
				End If
			End If
			Return resultSet
		End Function

		''' <summary>
		''' Generate the query to get all the results associated with a process.
		''' </summary>
		''' <param name="isNetworkDb">Whether or not the query is targetting the network DB.</param>
		''' <param name="resultIdsToExclude">
		''' The list of result IDs to exclude.
		''' The list should be comma-separated, usually prodiced by invoking <see cref="GetColumnValuesAsCSV(DataTable, String)"/>.
		''' Ignored if null or empty.
		''' </param>
		''' <returns>The SQL query to perform.</returns>
		Private Function GenerateResultQuery(isNetworkDb As Boolean, Optional resultIdsToExclude As String = Nothing) As String
			Dim hint As String = String.Empty
			If (isNetworkDb) Then
				hint = "with(nolock)"
			End If

			Dim sqlQueryAndClause = String.Empty
			If (Not String.IsNullOrEmpty(resultIdsToExclude)) Then
				sqlQueryAndClause = $"AND result_itemlistdef_id NOT IN ({resultIdsToExclude}) "
			End If

			Return $"SELECT * FROM {Process.ResultTable} {hint} WHERE result_process_id = {Process.ID} {sqlQueryAndClause} ORDER BY result_itemlistdef_id"
		End Function

		''' <summary>
		''' Create a comma-separated list of values from a column of a data table.
		''' </summary>
		''' <param name="results">The data table to iterate through.</param>
		''' <param name="columnName">The name of the column containing the data we care about.</param>
		''' <returns>The list of values, as a string.</returns>
		''' <remarks>
		''' The current implementation expects all rows to have a value, and do not surround
		''' the values with quotes. i.e. it's meant to be used with a column containing numbers.
		''' This would have to be improved if we need to exclude empty values or strings.
		''' </remarks>
		Private Function GetColumnValuesAsCSV(results As DataTable, columnName As String) As String
			Dim sb As New StringBuilder()
			For Each drNetResult As DataRow In results.Rows
				sb.Append($"{KillNull(drNetResult(columnName))},")
			Next
			sb.Length = If(sb.Length > 0, sb.Length - 1, 0)
			Return sb.ToString()
		End Function

		''' <summary>
		''' Extract the given column values from a local data row and create an Object array.
		''' </summary>
		''' <param name="columnNames">The name of the relevant columns.</param>
		''' <param name="localRow">The row from the local DB.</param>
		''' <returns>The resulting Object array.</returns>
		Private Function CreateNetworkResultRowData(columnNames As String(), localRow As DataRow) As Object()
			Dim columnValues = New List(Of Object)
			For Each aColumnName In columnNames
				columnValues.Add(localRow(aColumnName))
			Next
			Return columnValues.ToArray()
		End Function

		''' <summary>
		''' This method detects an invalid UDBS process instance data state.
		''' This was added at a time to work around a bug with an older version of Microsoft SQL Server
		''' that was causing duplicated rows to get generated.
		''' Since the UDBS servers are now running on up-to-date versions of SQL Server, this condition should
		''' no longer happen.
		''' </summary>
		''' <param name="localResultSet">The record set containing all results from the local DB.</param>
		''' <param name="networkResultSet">The record set containing all results from the network DB.</param>
		''' <returns>Whether or not the 'hack mode' condition was detected.</returns>
		Private Function DetectHackModeCondition(localResultSet As DataTable, networkResultSet As DataTable) As Boolean
			Dim localRowCount = If(localResultSet?.Rows?.Count, 0)
			Dim networkRowCount = If(networkResultSet?.Rows?.Count, 0)

			If (networkRowCount > localRowCount) Then
				logger.Info($"Hack mode condition detected. Local row count: {localRowCount}. Network row count: {networkRowCount}.")
			End If

			Return networkRowCount > localRowCount
		End Function

		''' <summary>
		''' Update the network BLOBs associated with a given process, from the local table.
		''' </summary>
		''' <param name="transactionScope">The ongoing transaction with the network SQL server.</param>
		''' <exception cref="Exception">If there is an error while updating the BLOBs.</exception>
		Private Sub UpdateNetworkBlobs(transactionScope As ITransactionScope)
			Dim sqlQuery As String = String.Empty
			Dim blobCount As Integer = 0
			If (CBLOB.GetLocalBlobCount(Process.Process, Process.ID, blobCount) <> ReturnCodes.UDBS_OP_SUCCESS) Then
				Throw New UDBSException($"Failed to retrieve BLOB count for process {Process.Process}, {Process.ID}.")
			End If
			logger.Debug($"Local {Process.BlobTable} table contains {blobCount} blob(s).")

			If blobCount > 0 Then
				' IMPORTANT: This may look like an error, but it is not.
				'            (i.e. We are filtering by Process ID, but we are applying the filter
				'            on a column named '...item_id'. But this is ok. The name of the local
				'            BLOB table's columns are a bit misleading. The 'blob_ref_item_id'
				'            column actually holds the 'Process ID'.)
				'            See the CBLOB class.
				sqlQuery = "SELECT blob_id,blob_ref_item_table,blob_ref_item_id,blob_array_name,blob_datagroup_name,blob_elements,blob_datatype,blob_isheader,blob_origsize " &
											   "FROM " & Process.BlobTable & " " &
											   "WHERE blob_ref_item_id = " & Process.ID & " " &
											   "ORDER BY blob_ref_item_table"
				Dim localBlobs As DataTable = Nothing
				If OpenLocalRecordSet(localBlobs, sqlQuery) <> ReturnCodes.UDBS_OP_SUCCESS Then
					Throw New UDBSException($"Failed to retrieve BLOB meta-data for process {Process.Process}, {Process.ID}.")
				End If

				' loop thru each blob record
				For Each drLocal As DataRow In localBlobs.Rows
					Dim blobPayload As Stream = Nothing
					Try
						' IMPORTANT: This may look like an error, but it is not.
						'            (i.e. We are looking for the item name, but we get the value of
						'            a column named '...table'. But this is ok. The name of the local
						'            BLOB table's columns are a bit misleading. The 'blob_ref_item_table'
						'            column actually holds the 'item name'.)
						'            See the CBLOB class.
						Dim itemName = KillNull(drLocal("blob_ref_item_table"))

						' Find the ID of the result item this BLOB is linked to.
						Dim resultId As Long = 0
						If (GetResultId(itemName, resultId) <> ReturnCodes.UDBS_OP_SUCCESS) Then
							Throw New UDBSException($"Failed to retrieve Result ID {itemName} of Process {Process.ID}.")
						End If

#Region "Load the Blob to a stream to avoid OOM"
						Dim rsBlob As DataRow = Nothing
						Dim retryCount As Integer = 0

						sqlQuery = $"select blob_blob from {Process.BlobTable} where blob_ref_item_id={Process.ID} And blob_id={KillNullInteger(drLocal("blob_id"))} order by blob_ref_item_table"
						While retryCount <= 1
							Try
								If retryCount < CBLOB.TemporaryStreamFallbackRetryCount Then
									'Memory-based, faster
									blobPayload = New MemoryStream()
								Else
									' File-based stream, slower
									blobPayload = New TempStream()
								End If
								OpenLocalRecordSet(rsBlob, blobPayload, sqlQuery)

								Exit While
							Catch oomex As OutOfMemoryException
								retryCount += 1
								GC.Collect()
							Catch ex As Exception
								Throw ' some other error, don't bother retrying
							End Try
						End While

						' Rewind the stream so we start from the beginning.
						blobPayload.Position = 0
#End Region

						sqlQuery = "SELECT blob_id,blob_ref_item_table,blob_ref_item_id,blob_array_name,blob_datagroup_name,blob_elements,blob_datatype,blob_isheader,blob_origsize  " &
													   "FROM " & Process.BlobTable & " with(nolock) " &
													   "WHERE blob_ref_item_id=" & resultId & " " &
													   "And blob_array_name='" & KillNull(drLocal("blob_array_name")) & "'"
						Dim networkBlobs As DataTable = Nothing
						OpenNetworkRecordSet(networkBlobs, sqlQuery, transactionScope)
						Dim columnNames As List(Of String) = networkBlobs.Columns.Cast(Of DataColumn)().[Select](
													Function(x) x.ColumnName).ToList()
						columnNames.Add("blob_blob")
						Dim newRecordCreated = False
						If (If(networkBlobs?.Rows?.Count, 0)) = 0 Then
							Dim newRecord = networkBlobs.NewRow()
							newRecord("blob_ref_item_table") = "testdata_result"
							newRecord("blob_ref_item_id") = resultId
							newRecord("blob_array_name") = drLocal("blob_array_name")
							newRecord("blob_datagroup_name") = drLocal("blob_datagroup_name")
							newRecord("blob_isheader") = drLocal("blob_isheader")
							newRecord("blob_datatype") = drLocal("blob_datatype")
							newRecord("blob_elements") = drLocal("blob_elements")
							newRecord("blob_origsize") = drLocal("blob_origsize")
							networkBlobs.Rows.Add(newRecord)
							newRecordCreated = True
						Else
							networkBlobs(0)("blob_ref_item_table") = "testdata_result"
							networkBlobs(0)("blob_ref_item_id") = resultId
							networkBlobs(0)("blob_array_name") = drLocal("blob_array_name")
							networkBlobs(0)("blob_datagroup_name") = drLocal("blob_datagroup_name")
							networkBlobs(0)("blob_isheader") = drLocal("blob_isheader")
							networkBlobs(0)("blob_datatype") = drLocal("blob_datatype")
							networkBlobs(0)("blob_elements") = drLocal("blob_elements")
							networkBlobs(0)("blob_origsize") = drLocal("blob_origsize")
						End If

						Dim payloadItems As List(Of Object) = New List(Of Object)(networkBlobs(0).ItemArray) From {
												blobPayload
											}
						If Not newRecordCreated Then
							UpdateNetworkRecord({"blob_id"}, columnNames.ToArray(), payloadItems.ToArray(),
																	Process.BlobTable, transactionScope)
						Else
							InsertNetworkRecord(columnNames.Skip(1).ToArray(),
																	payloadItems.Skip(1).ToArray(),
																	Process.BlobTable, transactionScope)
						End If
					Finally
						blobPayload?.Dispose()
					End Try
				Next
			End If
		End Sub

		''' <summary>
		''' Find the ID of a result item, given its name, in the context of the process.
		''' </summary>
		''' <param name="itemName">The item name.</param>
		''' <param name="resultId">(Output) Where the result ID will be store on success.</param>
		''' <returns>Whether or not the operation succeeded.</returns>
		Private Function GetResultId(itemName As String, ByRef resultId As Long) As ReturnCodes
			Dim sqlQuery = "SELECT result_id " &
					$"FROM {Process.ProcessTable} with(nolock), {Process.ResultTable} with(nolock), {Process.ItemListDefinitionTable} with(nolock) " &
					"WHERE process_id=result_process_id " &
					"And result_itemlistdef_id=itemlistdef_id " &
					"And itemlistdef_itemname='" & itemName & "' " &
					"AND process_id=" & Process.ID
			Dim resultIds As DataTable = Nothing
			OpenNetworkRecordSet(resultIds, sqlQuery)

			' Failed to determine the ID of the result this BLOB is linked to.
			If (If(resultIds?.Rows?.Count, 0)) <> 1 Then
				Return ReturnCodes.UDBS_OP_FAIL
			End If

			resultId = KillNullLong(resultIds(0)(0))
			Return ReturnCodes.UDBS_OP_SUCCESS
		End Function

		''' <summary>
		''' Hides the DatabaseSupport.DebugMesage(...) method and injects information such as the
		''' process name, ID, unit serial number and product.
		''' </summary>
		''' <param name="ex">The exception that needs to be logged.</param>
		Private Sub LogErrorInDatabase(ex As Exception)
			DatabaseSupport.LogErrorInDatabase(ex, Process.Process, Process.Stage, Process.ID, Process.ProductNumber, Process.UnitSerialNumber)
		End Sub
	End Class
End Namespace