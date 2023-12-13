Option Explicit On
Option Strict On
Option Compare Text
Option Infer On

Namespace TestDataInterface
    Public Module InterfaceSupport
        Public Const PROCESS As String = "testdata"
        Public Const INCOMPLETE As Integer = 1000

        Public Function IsSuccess(ByVal code As ResultCodes) As Boolean
            Select Case code
                Case ResultCodes.UDBS_SPECS_PASS, ResultCodes.UDBS_SPECS_PASS_INC, ResultCodes.UDBS_SPECS_NONE
                    Return True
                Case Else
                    Return False
            End Select
        End Function
    End Module
End Namespace
