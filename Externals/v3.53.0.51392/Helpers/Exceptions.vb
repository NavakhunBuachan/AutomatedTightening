Imports UdbsInterface.MasterInterface

''' <summary>
'''     General UDBS Exception
''' </summary>
Public Class UDBSException
    Inherits Exception

    Public Sub New()
        MyBase.New("UDBS general exception")
    End Sub

    Public Sub New(errorMessage As String)
        MyBase.New("UDBS exception: " & errorMessage)
    End Sub

    Public Sub New(inner_exception As Exception)
        MyBase.New("UDBS general exception", inner_exception)
    End Sub

    Public Sub New(errorMessage As String, inner_exception As Exception)
        MyBase.New("UDBS exception: " & errorMessage, inner_exception)
    End Sub
End Class

''' <summary>
'''     Exception thrown when accessing a UDBS test data instance
''' </summary>
Public Class UdbsTestException
    Inherits Exception

    Public Sub New(errorMessage As String)
        MyBase.New(errorMessage)
    End Sub

    Public Sub New(errorMessage As String, errorCode As ReturnCodes)
        MyBase.New(errorMessage & " (Return Code = " & UdbsTools.InterpretUDBSReturnCode(errorCode) & ")")
    End Sub
End Class

''' <summary>
'''     Exception thrown if test instance is not loaded
''' </summary>
''' <remarks></remarks>
Public Class UdbsTestNotLoadedException
    Inherits Exception

    Public Sub New()
        MyBase.New("Test instance not loaded.")
    End Sub
End Class

''' <summary>
'''     Exception thrown if test instance is not started / active
''' </summary>
''' <remarks></remarks>
Public Class UdbsTestNotStartedException
    Inherits Exception

    Public Sub New()
        MyBase.New("Test instance not started.")
    End Sub
End Class

''' <summary>
'''     Exception thrown if Item list is not loaded.
''' </summary>
Public Class UdbsItemListNotLoadedException
    Inherits Exception

    Public Sub New()
        MyBase.New("Item list is not loaded.")
    End Sub

    Public Sub New(errorMessage As String)
        MyBase.New(errorMessage & ": Item list is not loaded.")
    End Sub
End Class

''' <summary>
'''     Exception thrown if Item list does not contain the the item name.
''' </summary>
Public Class UdbsItemDoesNotExistException
    Inherits Exception

    Public Sub New(itemName As String)
        MyBase.New($"Item list does not contain item {itemName}.")
    End Sub

    Public Sub New(errorMessage As String, itemName As String)
        MyBase.New(String.Format("{1}: Item list does not contain item {0}.", itemName, errorMessage))
    End Sub
End Class

''' <summary>
'''     Error starting UDBS test instance - already in process
''' </summary>
Public Class UdbsTestInProcessException
    Inherits Exception

    ''' <summary>
    ''' </summary>
    ''' <param name="serialNo"></param>
    ''' <param name="partId"></param>
    ''' <param name="stationId"></param>
    ''' <param name="status"></param>
    ''' <remarks></remarks>
    Public Sub New(serialNo As String, partId As String, stationId As String, status As String)
        MyBase.New("Error starting test data instance: SN " & serialNo &
                   ", ID " & partId & " has already started on station " & stationId &
                   " with status = " & status & ".")
    End Sub
End Class

''' <summary>
'''     Problem accessing the Kitting database
''' </summary>
Public Class UdbsKittingException
    Inherits Exception

    Public Sub New(reason As String)
        MyBase.New(reason & " Please use the UDBS Traveller software to update the kitting information.")
    End Sub
End Class

''' <summary>
'''     Problem accessing the WIP database
''' </summary>
''' <remarks>ERR_ACCESS_WIP_PROCESS error code</remarks>
Public Class WipAccessException
    Inherits Exception

    Public Sub New()
        MyBase.New("There was a problem accessing the WIP database.")
    End Sub
End Class

''' <summary>
'''     Failed to load WIP process
''' </summary>
''' <remarks>ERR_LOAD_WIP_PROCESS error code</remarks>
Public Class WipLoadProcessException
    Inherits Exception

    Public Sub New(serialNumber As String, wipProcessID As Integer)
        MyBase.New("Serial number is active in WIP database, but there was a problem loading the WIP process. " &
                   "Please contact an engineer. [SN " & serialNumber & ", WIP PID " & wipProcessID & "]")
    End Sub
End Class

''' <summary>
'''     WIP Process is Locked (ReadOnly)
''' </summary>
''' <remarks>ERR_WIP_LOCKED</remarks>
Public Class WipLockedException
    Inherits Exception

    Public Sub New(serialNumber As String, lockedBy As String)
        MyBase.New(
            "WIP process is locked. The process may be active on another station or it was terminated unexpectedly. " &
            "Verify the Serial Number and contact an engineer to unlock the process. [SN " & serialNumber &
            ", Locked By " & lockedBy & "]")
    End Sub
End Class

''' <summary>
'''     WIP In Process
''' </summary>
Public Class WipInProcessException
    Inherits Exception

    Public Sub New(serialNumber As String, activeWipStep As String, startDate As Date)
        MyBase.New("WIP is already in process for the serial number. [SN " & serialNumber &
                   ", WIP step " & activeWipStep & " started at " &
                   Format(startDate, "hh:mm mm/dd/yyyy") & ".")
    End Sub
End Class

''' <summary>
'''     WIP Rework Exception
''' </summary>
Public Class WipReworkException
    Inherits Exception

    Public Sub New(serialNumber As String, requiredWipStep As String, activeWipStep As String)
        MyBase.New("The active step of the given Serial Number is Rework. " &
                   "Open the WIP Tracker software to re-route the unit to the required step. [SN " & serialNumber &
                   ", Required WIP Step = " & requiredWipStep & "]")
    End Sub
End Class

''' <summary>
'''     WIP Routing Exception.
''' </summary>
Public Class WipRoutingException
    Inherits Exception

    Public Sub New(serialNumber As String, requiredWipStep As String, activeWipStep As String)
        MyBase.New("Serial number has not been routed to the required WIP step. " &
                   "Verify the Serial Number and open the WIP Tracker software to re-route the unit to the required step. [SN " &
                   serialNumber &
                   ", Required WIP Step = " & requiredWipStep & ", Active WIP Step = " & activeWipStep & "]")
    End Sub
End Class

''' <summary>
'''     Serial Number inactive in WIP Exception
''' </summary>
''' <remarks>ERR_CHECK_WIP_FAIL error code</remarks>
Public Class WipSNNotActiveException
    Inherits Exception

    Public Sub New(serialNumber As String)
        MyBase.New("Serial number is not active in WIP database. " &
                   "Verify that you have entered the correct Serial Number, then open the WIP Tracker " &
                   "software and make sure that the unit is active. [SN " &
                   serialNumber & "]")
    End Sub
End Class

''' <summary>
'''     Serial number was not found in WIP database
''' </summary>
''' <remarks>ERR_CHECK_WIP_FAIL error code</remarks>
Public Class WipSNNotFoundException
    Inherits Exception

    Public Sub New(serialNumber As String)
        MyBase.New("Serial number was not found in WIP database. [SN " & serialNumber & "]")
    End Sub
End Class

''' <summary>
'''     Problem starting the WIP step.
''' </summary>
Public Class WipStartException
    Inherits Exception

    Public Sub New()
        MyBase.New("There was a error starting the WIP step. Please contact an engineer.")
    End Sub

    Public Sub New(returnCode As ReturnCodes)
        MyBase.New(
            "There was a error starting the WIP step. Please contact an engineer. Return Code = " &
            [Enum].GetName(GetType(ReturnCodes), returnCode))
    End Sub
End Class

''' <summary>
'''     Problem closing the WIP step.
''' </summary>
Public Class WipCloseException
    Inherits Exception

    Public Sub New(returnCode As ReturnCodes)
        MyBase.New(
            $"There was a error closing the WIP step. Return Code = {[Enum].GetName(GetType(ReturnCodes), returnCode) _
                      }. Please contact an engineer.")
    End Sub
End Class