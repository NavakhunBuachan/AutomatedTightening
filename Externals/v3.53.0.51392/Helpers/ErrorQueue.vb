Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Threading

Namespace MasterInterface
	''' <summary>
	''' Queue of system errors that are saved locally in a CSV file, then uploaded to UDBS.
	''' Local file is removed once everything in it has been saved to UDBS.
	''' </summary>
	Public Class ErrorQueue
		Inherits AsynchronousQueue(Of ErrorEntry)

		''' <summary>
		''' Description of an error and its context.
		''' </summary>
		Public Class ErrorEntry
			Property Guid As String
			Property PcDisplayName As String
			Property Description As String
			Property StackTrace As String
			Property FunctionName As String
			Property TimeOfEvent As Date
			' The version of the assembly, at the time the error occurs.
			' It is possible for an error to occur, get logged to CSV without being synchronized
			' with the network DB (i.e. network outage).
			' The application is upgraded, and the newer version of the application is the one
			' pushing the error entry to the database.
			' In that scenario, it needs to push the assembly version at the time the error occurs,
			' not the one pushing the error into the database.
			Property AssemblyVersion As String
		End Class

		''' <summary>
		''' Member variable for the <see cref="TemporaryErrorLogFile"/> property.
		''' </summary>
		Private _temporaryErrorLogFile As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
							 "JDSU", "UDBS", "errors.csv")

		''' <summary>
		''' Where system errors are logged.
		''' This is a temporary file; as errors are synchronized to the UDBS database this file is deleted
		''' (only when every errors stored in it have successfully been synchronized with the database).
		''' </summary>
		Public Property TemporaryErrorLogFile As String
			Get
				Return _temporaryErrorLogFile
			End Get
			Set(value As String)
				If (_temporaryErrorLogFile <> value) Then
					_temporaryErrorLogFile = value
					LoadFromFile(TemporaryErrorLogFile)
				End If
			End Set
		End Property

		''' <summary>
		''' Whether or not to delete error log file once all entries have been synchronized
		''' with the network database.
		''' This can turned off for the sake of unit testing.
		''' </summary>
		Public Property DeleteSynchronizedErrorLogFile As Boolean = True

		''' <inheritdoc/>
		Public Overrides Sub Start()
			Start(AddressOf HandleErrorEntry)
		End Sub

		''' <summary>
		''' Entry-processing method.
		''' </summary>
		''' <param name="entry">The entry to process.</param>
		Private Function HandleErrorEntry(entry As ErrorEntry) As Boolean
			If (LogErrorInDatabase(entry)) Then
				MarkAsSynchronized(entry)
				Return True
			End If

			Return False
		End Function

		''' <summary>
		''' Log a single system error into the database.
		''' </summary>
		''' <param name="entryToLog">The error entry to log.</param>
		Friend Function LogErrorInDatabase(entryToLog As ErrorQueue.ErrorEntry) As Boolean
			If CUtility.IsRecursiveLoop() Then
				Return True
			End If

			Try
				Dim description = entryToLog.Description
				If (Not String.IsNullOrWhiteSpace(entryToLog.StackTrace)) Then
					description = $"{description}{vbCrLf}Stack Trace: {entryToLog.StackTrace}"
				End If

				Dim sqlQuery = "INSERT INTO UDBS_ERROR (udbs_error_station, udbs_error_date, udbs_error_description, udbs_error_process) VALUES ('" &
				entryToLog.PcDisplayName & "', '" & DBDateFormat(CUtility.ConvertThaiToCommonEra(entryToLog.TimeOfEvent)) & "', '" & description &
				 $"', '{entryToLog.AssemblyVersion}')"
				ExecuteNetworkQuery(sqlQuery)
				Return True
			Catch ex As Exception
				logger.Error(ex, $"Function: {entryToLog.FunctionName}; Failed to log error to network.")
				Return False
			End Try
		End Function

		''' <summary>
		''' Load all entries from a local file.
		''' </summary>
		''' <param name="path">The file to be loaded.</param>
		Public Sub LoadFromFile(Optional path As String = Nothing)
			' Use the TemporaryErrorLogFile property if the parameter is not provided.
			If (String.IsNullOrEmpty(path)) Then
				path = TemporaryErrorLogFile
			End If

			Try
				If (Not File.Exists(path)) Then
					' No errors to log.
					' This is not an error.
					Return
				End If

				Using myReader = CsvFile.OpenCsvFile(path)
					ValidateErrorFileVersion(myReader.ReadFields())

					' Skip the second line (column header)
					myReader.ReadFields()

					While (Not myReader.EndOfData)
						HandleNextErrorFromFile(myReader.ReadFields())
					End While
				End Using
			Catch ex As Exception
				logger.Error(ex, $"Unexpected error loading error log from {path}.")
			End Try
		End Sub

		''' <summary>
		''' Validate that the file's version is compatible with the current software.
		''' For now, we only know about version 1.0.
		''' Anything other than that will throw an exception.
		''' </summary>
		''' <exception cref="System.Exception">
		''' If the version is not compatible with the software.
		''' </exception>
		''' <param name="columns">
		''' The CSV columns of the first line.
		''' Cell one (1) is expected to contain the word "Version"
		''' (anything else will cause an exception).
		''' Cell two (2) is expected to contain the version.
		''' At the moment, we only support "1.0".
		''' We can figure out versionning scheme and backward compatibility rules later...
		''' </param>
		Private Sub ValidateErrorFileVersion(columns() As String)
			If (columns(0) <> "Version") Then
				Throw New Exception("Invalid error log format.")
			End If
			Dim fileVersion = Double.Parse(columns(1))
			Const supportedVersion As Double = 1.0
			If (fileVersion <> supportedVersion) Then
				Throw New Exception($"Unsupported error log format version: {columns(1)}")
			End If
		End Sub

		''' <summary>
		''' Loads an entry (line) from the CSV file, and adds it to the queue.
		''' Entries are added to the queue asychronously, so this is not blocking.
		''' </summary>
		''' <param name="columns">The columns of a line from the CSV file.</param>
		Private Sub HandleNextErrorFromFile(columns() As String)
			If (columns Is Nothing) Then
				Throw New ArgumentException($"No columns provided.")
			ElseIf (columns.Count <> 8) Then
				Throw New ArgumentException($"Unexpected number of columns. Expecting 5, for {columns.Count}.")
			End If

			Dim synchronized As Boolean = Boolean.Parse(columns(0))
			If (synchronized) Then
				' This error log has already been pushed to the DB.
				Return
			End If

			Dim entry As New ErrorEntry With {
				.Guid = columns(1),
				.TimeOfEvent = DatabaseSupport.DBDateParse(columns(2)),
				.FunctionName = columns(3),
				.PcDisplayName = columns(4),
				.AssemblyVersion = columns(5),
				.Description = columns(6),
				.StackTrace = columns(7)}

			' Call add from the base method, since this is loaded from the file.
			' The overriden method also adds it to the CSV file.
			' This is not what we want to do while loading the CSV file.
			MyBase.Add(entry)
		End Sub

		Public Overrides Sub Add(toAdd As ErrorEntry)
			If IsCalledOnDequeueThread() Then
				logger.Debug("Entry added from the entry-processing thread. This might lead to an infinite recursive loop. Ignore it.")
				Return
			End If

			' Log the error locally in the CSV file and add it to the queue of errors 
			' to be pushed to the network DB.
			LogErrorToCsv(toAdd)
			MyBase.Add(toAdd)
		End Sub

		''' <summary>
		''' Write a single error to the local CSV file.
		''' </summary>
		''' <param name="entryToLog">The entry to log to the CSV file.</param>
		Public Sub LogErrorToCsv(
				entryToLog As ErrorEntry)
			Try
				Monitor.Enter(TemporaryErrorLogFile)

				Dim folder As String = Path.GetDirectoryName(TemporaryErrorLogFile)
				If (Not Directory.Exists(folder)) Then
					Directory.CreateDirectory(folder)
				End If

				If (Not File.Exists(TemporaryErrorLogFile)) Then
					InitializeCsvFile()
				End If

				Dim columns() As String = {
						"False", ' Not synchronized to Network DB yet.
						entryToLog.Guid,
						DBDateFormat(entryToLog.TimeOfEvent),
						entryToLog.FunctionName,
						entryToLog.PcDisplayName,
						entryToLog.AssemblyVersion,
						entryToLog.Description,
						entryToLog.StackTrace}

				File.AppendAllText(TemporaryErrorLogFile, ToCSV(columns))
			Catch ex As Exception
				logger.Error(ex, "Unexpected error while writing to the log.")
			Finally
				Monitor.Exit(TemporaryErrorLogFile)
			End Try
		End Sub

		''' <remarks>Call while holding the monitor onto ErrorEntryQueue.</remarks>
		Private Sub InitializeCsvFile(Optional filePath As String = Nothing)
			Dim str As New StringBuilder()

			Const CurrentVersion As String = "1.0"
			Dim columns() As String = {"Version", CurrentVersion}
			str.Append(ToCSV(columns))

			columns = {"Synchronized", "GUID", "Time Stamp", "Function", "Station", "Assembly Version", "Description", "Stack Trace"}
			str.Append(ToCSV(columns))

			If (filePath Is Nothing) Then
				filePath = TemporaryErrorLogFile
			End If

			File.WriteAllText(filePath, str.ToString())
		End Sub

		''' <summary>
		''' Expands an array of strings into a single CSV line.
		''' Escapes every value in quotes.
		''' The line ends with a new-line character.
		''' </summary>
		''' <param name="columns">The columns to be expanded.</param>
		''' <returns>The resulting line to write to the CSV file.</returns>
		Private Function ToCSV(columns() As String) As String
			Dim str As New StringBuilder()
			Dim aColumn As String

			For Each aColumn In columns
				str.Append($"""{ToCsvCell(aColumn)}"",")
			Next

			If columns.Length > 0 Then
				' Remove trailing ','
				str.Remove(str.Length - 1, 1)
			End If

			str.AppendLine()

			Return str.ToString()
		End Function

		''' <summary>
		''' Formats a single value to be inserted into a CSV file.
		''' Escapes special characters.
		''' </summary>
		''' <param name="value">The value to format.</param>
		''' <returns>The formated value.</returns>
		Private Function ToCsvCell(value As String) As String
			If (String.IsNullOrWhiteSpace(value)) Then
				Return value
			End If

			value = value.Replace("""", """""")

			Return value
		End Function

		''' <summary>
		''' Mark an entry as 'synchronized' in the temporary file.
		''' There is an option to delete the file when all entries have been synchronized.
		''' If this option is turned off, this lets this class know that already-synchronized entries 
		''' need not be pushed to the database again.
		''' In the case where there are a lot of errors, and not every one can be synchronized
		''' before the application is closed again, only entries that have not been synchronized
		''' will be retried when the application starts again.
		''' </summary>
		Private Sub MarkAsSynchronized(entry As ErrorEntry)
			Dim found As Boolean = False

			Monitor.Enter(TemporaryErrorLogFile)
			Try
				If (Not File.Exists(TemporaryErrorLogFile)) Then
					Return
				End If

				Dim temporaryFile = $"{TemporaryErrorLogFile}.tmp"

				InitializeCsvFile(temporaryFile)

				Dim unsynchronizedCount As Integer = 0

				Using errorLogFile = CsvFile.OpenCsvFile(TemporaryErrorLogFile)
					' Ignore the first two lines.
					Dim columns = errorLogFile.ReadFields()
					columns = errorLogFile.ReadFields()

					While Not errorLogFile.EndOfData
						columns = errorLogFile.ReadFields()
						Dim guid As String = columns(1)
						If (guid = entry.Guid) Then
							' This is the one we need to mark as synchronized.
							found = True
							columns(0) = "True"
						ElseIf Not Boolean.Parse(columns(0)) Then
							' This is not the one we are looking for, and it is not synchronized.
							unsynchronizedCount += 1
						End If

						File.AppendAllText(temporaryFile, ToCSV(columns))
					End While
				End Using

				If Not found Then
					logger.Warn($"Error entry '{entry.Guid}' not found.")
				End If

				' Instead of complex if/else statement, just remember whether or not the files were removed...
				Dim filesRemoved As Boolean = False

				If (unsynchronizedCount = 0) Then
					logger.Debug("All entries in the error logs have been pushed to the network DB.")
					If (DeleteSynchronizedErrorLogFile) Then
						File.Delete(temporaryFile)
						File.Delete(TemporaryErrorLogFile)
						filesRemoved = True
					End If
				End If

				If Not filesRemoved Then
					File.Replace(temporaryFile, TemporaryErrorLogFile, Nothing)
				End If
			Finally
				Monitor.Exit(TemporaryErrorLogFile)
			End Try
		End Sub

		''' <summary>
		''' Retry schedule is: 0, 100, 500, 1000, 2000, 5000.
		''' We retry immediatly, then we retry at slower and slower time interval, up to 5 seconds.
		''' Then we keep on retrying every 5 seconds.
		''' </summary>
		Protected Overrides Function GetRetryDelay(retryCount As Integer) As Integer
			Dim retryWaitDelay() As Integer = New Integer() {0, 100, 500, 1000, 2000, 5000}

			Dim howLongToWait = retryWaitDelay.Last
			If (retryCount < retryWaitDelay.Length) Then
				howLongToWait = retryWaitDelay(retryCount)
			End If

			logger.Warn($"Failed to write system error to network DB. Will retry in {howLongToWait} msec.")

			Return howLongToWait
		End Function
	End Class
End Namespace