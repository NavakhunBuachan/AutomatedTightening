Option Explicit On
Option Compare Text
Option Infer On
Option Strict On

Namespace MasterInterface
    Friend Module ClassSupport
        ''' <summary>Creates a unique identifier for a process running in Windows</summary>
        ''' <returns>A string identifier that can be stored and later used to look up to see if a process still exists</returns>
        ''' <remarks>Because Windows can re-use process IDs, we concatenate the date to the process ID, making it unique forever</remarks>
        Friend Function CustomFormatWindowsProcessID(SystemProcess As Process) As String
            Dim tmpID As Integer
            Dim tmpStartTime As Date
            Try
                tmpID = SystemProcess.Id
                tmpStartTime = SystemProcess.StartTime
            Catch ex As Exception
                'sometimes, for example on processID 0 and 4 (Idle and Main processes), the start time is not available
                'one should avoid passing in process objects that refer to these special processes (though the run-time error is trapped here)
                tmpStartTime = Date.MinValue
            End Try
            Return String.Concat(tmpID.ToString, "-", tmpStartTime.ToString("yyyyMMddhhmmss"))
        End Function

        ''' <summary>Returns the parts of a formatted custom process ID so we can check against current windows processes</summary>
        ''' <remarks>see function CustomFormatProcessID to see reverse of this function</remarks>
        Friend Sub SplitCustomFormattedWindowsProcessID(CustomProcessID As String, ByRef ProcessID As Integer,
                                                         ByRef ProcessDate As String)
            ProcessID = CInt(Val(ProcessID))
            ProcessDate = ""
            Try
                Dim SplitString As String() = Split(CustomProcessID, "-")
                ProcessID = CInt(Val(SplitString(0)))
                If SplitString.Length > 1 Then ProcessDate = SplitString(1)
            Catch ex As Exception
                'just to handle unexpected stuff
            End Try
        End Sub

        ''' <summary>
        ''' Clamps a numeric value between a range.
        ''' </summary>
        ''' <param name="value">The numeric value to clamp.</param>
        ''' <param name="min">Minimum value.</param>
        ''' <param name="max">Maximum value.</param>
        ''' <returns></returns>
        Public Function Clamp(ByVal value As Double, ByVal min As Double, ByVal max As Double) As Double

            If value = Double.NaN Then
                Return Nothing
            End If

            If min > max Then
                Throw New UDBSException($"Unable to clamp [{value}] since Min=[{min}] is greater than Max=[{max}]")
            End If

            If value < min Then
                Return min
            ElseIf value > max Then
                Return max
            End If

            Return value
        End Function

    End Module
End Namespace
