Imports System.IO
Imports Microsoft.Win32

Namespace MasterInterface
	''' <summary>
	''' Interface for an Operating System Abstraction Layer.
	''' This is useful for unit tests to mock interactions with
	''' the operating system that cannot be easily be mocked 
	''' otherwise.
	''' </summary>
	Public Interface IOsAbstractionLayer
		''' <summary>
		''' Checks whether or not a given process is running.
		''' </summary>
		''' <param name="CustomProcessID">
		''' The process ID (using our own custom pattern.)
		''' </param>
		''' <returns>Whether or not that process is currently running.</returns>
		Function IsProcessStillRunning(CustomProcessID As String) As Boolean

		''' <summary>
		''' Adapter to read and write values from/to the Windows registry.
		''' </summary>
		''' <returns></returns>
		ReadOnly Property RegistryAdapter As IRegistryAdapter
	End Interface

	''' <summary>
	''' Interface for saving and reading back values from the Windows registry.
	''' </summary>
	Public Interface IRegistryAdapter
		Sub SaveSetting(applicationName As String, sectionName As String, key As String, value As String)
		Function GetSetting(applicationName As String, sectionName As String, key As String, Optional defaultValue As String = "") As String
	End Interface

	''' <summary>
	''' Base implementation.
	''' Acts as a Singleton the unit tests can replace through the 'internal'
	''' setter of the <see cref="OsAbstractionLayer.Instance"/> property.
	''' </summary>
	Public MustInherit Class OsAbstractionLayer : Implements IOsAbstractionLayer
		Private Shared _instance As IOsAbstractionLayer = Nothing

		''' <summary>
		''' Singleton instance.
		''' </summary>
		Public Shared Property Instance As IOsAbstractionLayer
			Get
				If (_instance Is Nothing) Then
					_instance = New WindowsAbstractionLayer()
				End If

				Return _instance
			End Get
			Friend Set(value As IOsAbstractionLayer)
				_instance = value
			End Set
		End Property

		''' <inheritdoc/>
		Public MustOverride Property RegistryAdapter As IRegistryAdapter Implements IOsAbstractionLayer.RegistryAdapter

		''' <summary>
		''' Unit tests will want to reset this in their 'TearDown' method.
		''' </summary>
		Friend Shared Sub ResetInstance()
			_instance = Nothing
		End Sub

		''' <inheritdoc/>
		Public MustOverride Function IsProcessStillRunning(CustomProcessID As String) As Boolean Implements IOsAbstractionLayer.IsProcessStillRunning
	End Class

	''' <summary>
	''' Default implementation for the Microsoft Windows operating system.
	''' </summary>
	Friend Class WindowsAbstractionLayer : Inherits OsAbstractionLayer
		Public Overrides Property RegistryAdapter As IRegistryAdapter = New WindowsRegistryAdapter()

		''' <summary>
		''' Checks to see if a process entry from the local process registration table is still running.
		''' </summary>
		''' <remarks>
		''' This code was moved from the <see cref="ClassSupport"/> class.
		''' </remarks>
		Public Overrides Function IsProcessStillRunning(CustomProcessID As String) As Boolean
			Dim PID As Integer = -1
			Dim PDate = ""

			SplitCustomFormattedWindowsProcessID(CustomProcessID, PID, PDate)

			If PID <= 0 Then
				'something went wrong - this is not a valid process ID
				'let's say whatever process it was is no longer running
				Return False
			End If

			Try
				Dim AllProcesses = Process.GetProcesses

				For Each P In AllProcesses
					If P.Id = PID Then
						If String.Compare(CustomProcessID, CustomFormatWindowsProcessID(P), True) = 0 Then
							Return True
						End If
					End If
				Next
			Catch ex As Exception
				'this is just to prevent major crashes. Sometimes you get WIN32 exceptions. 
				'In this case - just assume the application in question is not running
			End Try

			Return False
		End Function
	End Class

	''' <summary>
	''' Implementation of the <see cref="IRegistryAdapter"/>. This is the real deal!
	''' </summary>
	Friend Class WindowsRegistryAdapter
		Implements IRegistryAdapter

		Private Const VB6RegistryHive As String = "Software\VB and VBA Program Settings"

		''' <inheritdoc/>
		Public Sub SaveSetting(applicationName As String, sectionName As String, key As String, payload As String) Implements IRegistryAdapter.SaveSetting
			Dim targetPath As String = String.Empty
			Try
				targetPath = IO.Path.Combine(VB6RegistryHive, applicationName, sectionName)
				Dim registryKey As RegistryKey = Registry.CurrentUser.OpenSubKey(targetPath.ToString(), True)

				If registryKey Is Nothing Then
					Throw New InvalidDataException($"Unable to open {targetPath} registry for writing {payload}")
				End If

				registryKey.SetValue(key, payload)
			Catch ex As Exception
				logger.Error(ex, $"Error reading {key}. Unable to write {payload} to {targetPath}")
				Throw
			End Try
		End Sub

		''' <inheritdoc/>
		Public Function GetSetting(applicationName As String, sectionName As String, key As String, Optional defaultValue As String = "") As String Implements IRegistryAdapter.GetSetting
			Try
				Dim targetPath As String = IO.Path.Combine(VB6RegistryHive, applicationName, sectionName)
				Dim registryKey As RegistryKey

				' open read only
				registryKey = Registry.CurrentUser.OpenSubKey(targetPath)

				If registryKey Is Nothing Then
					Return defaultValue
				End If

				Return CStr(registryKey.GetValue(key, defaultValue))
			Catch ex As Exception
				logger.Error(ex, $"Error reading {key}. Returning default value {defaultValue} instead")
				Return defaultValue
			End Try
		End Function
	End Class
End Namespace

