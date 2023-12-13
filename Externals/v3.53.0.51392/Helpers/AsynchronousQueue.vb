Imports System.Threading

Namespace MasterInterface
	''' <summary>
	''' Generic implementation of a queue that deals with items asynchronously in a "first-in, first-out" fashion.
	''' </summary>
	''' <typeparam name="T">The type of items to be included.</typeparam>
	Public MustInherit Class AsynchronousQueue(Of T) : Implements IDisposable
		Private _dequeueThread As Thread
		''' <summary>
		''' First-in, first out.
		''' Add at the end. Dequeue at the beginning.
		''' </summary>
		Private _entries As New List(Of T)
		Private _dequeueMethod As Func(Of T, Boolean)
		Private _stopRequested As Boolean = False
		''' <summary>
		''' For throttling.
		''' Add at the end, dequeue at the beginning.
		''' </summary>
		Private _recentProcessedEntries As New List(Of HandledEntry(Of T))

		''' <summary>
		''' How big the throttling window size is (number of entries).
		''' </summary>
		''' <returns></returns>
		Public Property ThrottlingWindowSize As Integer = 5

		''' <summary>
		''' How long the throttling window period is (in milliseconds).
		''' </summary>
		''' <returns></returns>
		Public Property ThrottlingWindowPeriod As Integer = 5000

		''' <summary>
		''' Maximum queue size.
		''' </summary>
		Public Property MaxSize As Integer = 500

		''' <summary>
		''' Current queue size.
		''' </summary>
		Public ReadOnly Property Count As Integer
			Get
				Return _entries.Count
			End Get
		End Property

		''' <summary>
		''' Whether or not we can add an element to the queue
		''' (i.e. has the maximum number of entries been reached?)
		''' </summary>
		Public ReadOnly Property CanAdd As Boolean
			Get
				Return _entries.Count < MaxSize
			End Get
		End Property

		''' <summary>
		''' What was the oldest entry in the throttling window was added.
		''' </summary>
		Private ReadOnly Property TimeOfOldestEntryInThrottlingWindow As Date
			Get
				Monitor.Enter(_entries)
				Try
					If (_recentProcessedEntries.Count = 0) Then
						Return Date.MinValue
					Else
						Return _recentProcessedEntries(0).TimeWhenItWasProcessed
					End If
				Finally
					Monitor.Exit(_entries)
				End Try
			End Get
		End Property

		''' <summary>
		''' Delegate signature for a method processing an entry from the queue.
		''' </summary>
		''' <param name="toHandle"></param>
		Delegate Sub DequeueMethod(toHandle As T)

		''' <summary>
		''' Default constructor.
		''' </summary>
		Public Sub New()
		End Sub

		''' <summary>
		''' Implementors of this class should override this method and call
		''' <see cref="Start(Func(Of T, Boolean))"/> with the proper dequeue method.
		''' </summary>
		Public MustOverride Sub Start()

		''' <inheritdoc/>
		Public Sub Dispose() Implements IDisposable.Dispose
			Monitor.Enter(_entries)
			Try
				RequestStop()

				If _dequeueThread IsNot Nothing Then
					' Wait for the thread to complete.
					_dequeueThread.Join()
				End If
			Finally
				Monitor.Exit(_entries)
			End Try
		End Sub

		''' <summary>
		''' Raises a flag asking for the queue's processing thread to terminate.
		''' This is non-blocking.
		''' </summary>
		Public Sub RequestStop()
			Monitor.Enter(_entries)
			Try
				_stopRequested = True
				Monitor.PulseAll(_entries)
			Finally
				Monitor.Exit(_entries)
			End Try
		End Sub

		''' <summary>
		''' Start the thread with a given dequeue method.
		''' </summary>
		''' <param name="myDequeueMethod">The dequeue method to process entries added to the queue.</param>
		Protected Sub Start(myDequeueMethod As Func(Of T, Boolean))
			If _dequeueThread IsNot Nothing Then
				Throw New Exception("Thread already running.")
			End If

			_stopRequested = False
			_dequeueMethod = myDequeueMethod
			_dequeueThread = New Thread(AddressOf DequeueProcessingThread)
			_dequeueThread.IsBackground = True
			_dequeueThread.Start()
		End Sub

		''' <summary>
		''' Checks whether or not the processing thread is running.
		''' </summary>
		Public ReadOnly Property IsRunning As Boolean
			Get
				Return _dequeueThread IsNot Nothing AndAlso _dequeueThread.IsAlive
			End Get
		End Property

		''' <summary>
		''' Clears the queue.
		''' Unprocessed entries are removed.
		''' </summary>
		Public Sub Clear()
			Monitor.Enter(_entries)
			Try
				_entries.Clear()
				_recentProcessedEntries.Clear()
				Monitor.PulseAll(_entries)
			Finally
				Monitor.Exit(_entries)
			End Try
		End Sub

		''' <summary>
		''' Wait for the processing thread to terminate.
		''' </summary>
		Public Sub Join()
			_dequeueThread?.Join()
		End Sub

		''' <summary>
		''' Waits for the queue to be empty.
		''' </summary>
		Public Sub WaitForQueueToBeEmpty()
			While (True)
				Monitor.Enter(_entries)
				Try
					' When we stop the thread, it might never get the chance
					' to be emptied, so stop waiting.
					If (_stopRequested) Then
						Return
					End If

					If (_entries.Count = 0) Then
						Return
					Else
						Monitor.Wait(_entries)
					End If
				Finally
					Monitor.Exit(_entries)
				End Try
			End While
		End Sub

		''' <summary>
		''' Checks whether or not the caller is running on the dequeue thread.
		''' This is useful for detecting 'errors while processing an error' situation
		''' in the child class and prevent infinite recursive loop.
		''' </summary>
		''' <returns>Whether or not the caller is running on the dequeue thread.</returns>
		Protected Function IsCalledOnDequeueThread() As Boolean
			If (_dequeueThread?.Equals(Thread.CurrentThread)) Then
				Return True
			Else
				Return False
			End If
		End Function

		''' <summary>
		''' Adds an entry to the queue.
		''' </summary>
		''' <param name="anEntry">The entry to be added.</param>
		Public Overridable Sub Add(anEntry As T)
			Monitor.Enter(_entries)
			Try
				If Not CanAdd Then
					logger.Warn($"Maximum queue size reached ({MaxSize}). Discarding entry.")
					Return
				End If

				_entries.Add(anEntry)
				Monitor.PulseAll(_entries)
			Finally
				Monitor.Exit(_entries)
			End Try
		End Sub

		''' <summary>
		''' The processing thread's main method.
		''' </summary>
		Private Sub DequeueProcessingThread()
			Try
				While (True)
					Monitor.Enter(_entries)
					Try
						If (_stopRequested) Then
							Return
						End If

						If (_entries.Count = 0) Then
							Monitor.Wait(_entries)
							Continue While
						End If
					Finally
						Monitor.Exit(_entries)
					End Try

					Try
						Flush()
					Catch ex As Exception
						logger.Error(ex, "Unexpected error flushing the queue.")
						Continue While
					End Try

					Monitor.Enter(_entries)
					Try
						' Wake up everyone waiting for the queue to be empty.
						Monitor.PulseAll(_entries)
					Finally
						Monitor.Exit(_entries)
					End Try
				End While
			Finally
				_dequeueThread = Nothing
			End Try
		End Sub

		''' <summary>
		''' Wait if the throttling window parameters and content requires it.
		''' </summary>
		Private Sub ThrottleIfNeeded()
			RemoveExtraRecentlyProcessedEntries()

			If (_recentProcessedEntries.Count < ThrottlingWindowSize) Then
				Return
			End If

			Dim timeOfOldestEntry = Me.TimeOfOldestEntryInThrottlingWindow
			If (TimeOfOldestEntryInThrottlingWindow = Date.MinValue) Then
				' Throttling window empty.
				Return
			End If

			Dim delta = Date.Now - timeOfOldestEntry

			' Need to throttle.
			Dim howLongToWait = ThrottlingWindowPeriod - delta.TotalMilliseconds
			Dim waitUntil = Date.Now.AddMilliseconds(howLongToWait)
			While waitUntil > Date.Now
				' Wait will be interrupted everytime we add an entry.
				' So don't just wait once. Keep on waiting until we reached
				howLongToWait = (waitUntil - Date.Now).TotalMilliseconds
				Monitor.Wait(_entries, waitUntil - Date.Now)

				If (_stopRequested) Then
					' The thread is being interrupted.
					Return
				End If
			End While
			'End If
		End Sub

		''' <summary>
		''' How long to wait on a given retry.
		''' </summary>
		''' <param name="retryCount">
		''' How many time have we been retrying.
		''' The first time this method is invoked, the value will be zero (0).
		''' </param>
		''' <returns>
		''' How long to wait, in milliseconds.
		''' Default implementation returns one hundred (100) milliseconds all the time
		''' to prevent CPU starvation.
		''' </returns>
		Protected Overridable Function GetRetryDelay(retryCount As Integer) As Integer
			Return 100
		End Function

		''' <summary>
		''' Send all entries in the queue, throttling along the way as needed.
		''' </summary>
		Private Sub Flush()
			While (True)
				Dim entryToHandle As T = Nothing

				Monitor.Enter(_entries)
				Try
					ThrottleIfNeeded()

					If (_entries.Count = 0) Then
						Return
					End If
					If (_stopRequested) Then
						Return
					End If

					' Picks the entry to handle.
					' We will remove it later, only if the operation succeeds.
					entryToHandle = _entries(0)
				Finally
					Monitor.Exit(_entries)
				End Try

				Dim retryCount As Integer = 0
				Dim retryWaitDelay() As Integer = New Integer() {0, 100, 500, 1000, 2000, 5000}
				While (True)
					Dim processedSucccessfully As Boolean
					Try
						processedSucccessfully = _dequeueMethod.Invoke(entryToHandle)
					Catch ex As Exception
						logger.Error(ex, $"Error processing entry.")
						processedSucccessfully = False
					End Try

					If (processedSucccessfully) Then
						Exit While
					End If

					Dim endSleepTime = DateTime.Now.AddMilliseconds(Me.GetRetryDelay(retryCount))
					While endSleepTime > DateTime.Now
						Monitor.Enter(_entries)
						Try
							Monitor.Wait(_entries, endSleepTime - Date.Now)

							If (_stopRequested) Then
								' The thread is being interrupted.
								Return
							End If
						Finally
							Monitor.Exit(_entries)
						End Try
					End While

					retryCount += 1
				End While

				Dim processedTime = Date.Now

				' We have successfully handled the entry.
				' We can take it out of the queue.
				Monitor.Enter(_entries)
				Try
					_entries.RemoveAt(0)
					_recentProcessedEntries.Add(New HandledEntry(Of T) With
						{
							.Entry = entryToHandle,
							.TimeWhenItWasProcessed = processedTime
						})
				Finally
					Monitor.Exit(_entries)
				End Try
			End While
		End Sub

		''' <summary>
		''' Removes expired entries in the throttling window.
		''' </summary>
		Private Sub RemoveExtraRecentlyProcessedEntries()
			Dim currentTime = Date.Now
			Monitor.Enter(_entries)
			Try
				While _recentProcessedEntries.Count > 0
					Dim age = (currentTime - _recentProcessedEntries(0).TimeWhenItWasProcessed).TotalMilliseconds
					If (age > ThrottlingWindowPeriod) Then
						_recentProcessedEntries.RemoveAt(0)
					Else
						' All entries following this one will be younger.
						' No need to keep looking.
						Return
					End If
				End While
			Finally
				Monitor.Exit(_entries)
			End Try
		End Sub

		''' <summary>
		''' Entry in the throttling window.
		''' </summary>
		''' <typeparam name="T">The type of the entry.</typeparam>
		Protected Class HandledEntry(Of T)
			''' <summary>
			''' The item that has been processed.
			''' </summary>
			Public Property Entry As T

			''' <summary>
			''' When that entry was processed.
			''' </summary>
			Public Property TimeWhenItWasProcessed As Date
		End Class
	End Class

End Namespace