Option Explicit On
Option Compare Binary
Option Infer On
Option Strict On

Imports UdbsInterface.MasterInterface


Namespace WipInterface
    Public Module Support
        Private Const ClsName = "Support"
        Friend Const PROCESS As String = "WIP"

        Private _
            Const GIVEN_ORACLE =
            "SELECT
                                                CASE
                                                    WHEN pg_string_value IS NULL THEN NULL
		                                            WHEN PATINDEX('%,%', pg_string_value)= 0 then pg_string_value
                                                    ELSE SUBSTRING(pg_string_value, 1, PATINDEX('%,%', pg_string_value) - 1)
                                                END AS OraclePartNumber
                                                ,prdg.pg_sequence
	                                            ,prdg.pg_product_group
												,grp.prdgrp_product_number
                                            FROM
                                                product prod with(NOLOCK)
                                                LEFT OUTER JOIN udbs_prdgrp grp   with(NOLOCK) ON prod.product_number = grp.prdgrp_product_number
                                                LEFT OUTER JOIN udbs_product_group prdg with(NOLOCK) ON (prdg.pg_product_group = grp.prdgrp_product_group)
										   
										   WHERE --prod.product_number='1-01395'
											    CASE
                                                    WHEN pg_string_value IS NULL THEN NULL
		                                            WHEN PATINDEX('%,%', pg_string_value)= 0 then pg_string_value
                                                    ELSE SUBSTRING(pg_string_value, 1, PATINDEX('%,%', pg_string_value) - 1)
                                                END='{0}'"

        ''' <summary>
        '''     Given an Oracle Part Number, Return uDBS
        ''' </summary>
        ''' <param name="oraclePN"></param>
        ''' <returns></returns>
        Public Function GetUDBSPartNumber(oraclePN As String) As String
            Dim rsTemp As DataTable = Nothing
            OpenNetworkRecordSet(rsTemp, String.Format(GIVEN_ORACLE, oraclePN))
            If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                ' record exists...
                Return (rsTemp(0)("prdgrp_product_number")).ToString()
            Else
                logger.Warn($"UDBS Part Number not found for Oracle Part Number ={oraclePN}. Returning {oraclePN}")
                Return oraclePN
            End If
        End Function

        ' Candidate for removal.
        Private Function OnErrorResumeNext(returnExceptions As Boolean, putNullWhenNoExceptionIsThrown As Boolean, ParamArray actions As Action()) As Exception()
            Dim exceptions = If(returnExceptions, New List(Of Exception)(), Nothing)

            For Each action In actions
                Dim exp As Exception = Nothing

                Try
                    action()
                Catch ex As Exception

                    If returnExceptions Then
                        exp = ex
                    End If
                End Try

                If exp IsNot Nothing OrElse putNullWhenNoExceptionIsThrown Then
                    exceptions.Add(exp)
                End If
            Next

            Return exceptions?.ToArray()
        End Function

    End Module
End Namespace
