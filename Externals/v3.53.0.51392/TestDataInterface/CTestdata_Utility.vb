Option Explicit On
Option Strict On
Option Compare Text
Option Infer On

Imports UdbsInterface.MasterInterface


Namespace TestDataInterface
    Public Class CTestData_Utility
        Inherits CUtility

        '**********************************************************************
        '* Standard Utility Functions
        '**********************************************************************

        ''' <summary>
        ''' Return the name of the computer as specified in the network settings.
        ''' </summary>
        ''' <param name="StationName">(Out) The station name.</param>
        ''' <returns>The outcome of this operation.</returns>
        Public Shared Function GetStationName(ByRef StationName As String) As ReturnCodes
            Return Utility_GetStationName(StationName)
        End Function

        '**********************************************************************
        '* Product Support Functions
        '**********************************************************************

        Public Function GetPartIdentifier(OraclePN As String,
                                          ByRef ProductNumber As String) As ReturnCodes
            ' get unit detail information from product group tables
            Return MyBase.Product_GetPartIdentifier(OraclePN, ProductNumber)
        End Function

        Public Function IsUnitExist(ProductNumber As String,
                                    Release As Integer,
                                    SerialNumber As String) _
            As Boolean
            ' does the unit exist in UDBS
            IsUnitExist = MyBase.Product_UnitExists(ProductNumber, Release, SerialNumber)
        End Function
    End Class
End Namespace
