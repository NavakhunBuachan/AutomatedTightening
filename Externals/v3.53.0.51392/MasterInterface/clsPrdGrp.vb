

Namespace MasterInterface
    Public Class clsPrdGrp
        Private Function HexToDbl(sHex As String) As Double

            Dim dTmp As Double
            Dim sTmp As String
            Dim iChrPos As Integer
            Dim iDec As Integer

            HexToDbl = -1
            sHex = UCase(sHex)
            dTmp = 0

            iChrPos = Len(sHex)
            Do While iChrPos <> 0
                sTmp = Mid(sHex, iChrPos, 1)
                Select Case sTmp
                    Case "0" To "9"
                        iDec = KillNullInteger(sTmp)
                    Case "A" To "F"
                        iDec = Asc(sTmp) - 55
                    Case Else
                        MsgBox("Invalid character: " & sTmp, vbExclamation)
                        Exit Function
                End Select
                dTmp = dTmp + iDec * (16 ^ (Len(sHex) - iChrPos))
                iChrPos = iChrPos - 1
            Loop

            Return dTmp
        End Function

        Private Function DblToHex(dValue As Double) As String

            Dim sChr As String
            Dim sTmp As String
            Dim dDiv As Double
            Dim dMod As Double
            Dim dTmp As Double

            DblToHex = ""
            sTmp = ""

            dDiv = dValue
            Do
                dTmp = dDiv
                ModDiv(dTmp, 16, dMod, dDiv)
                Select Case dMod
                    Case 0 To 9
                        sChr = CStr(dMod)
                    Case 10 To 16
                        sChr = Chr(CInt(dMod) + 55)
                    Case Else
                        Throw New Exception($"Invalid character: {sTmp}")
                End Select
                sTmp = sChr & sTmp
            Loop While dDiv <> 0

            Return sTmp
        End Function


        Private Sub ModDiv(dValue As Double,
                           dDivider As Double,
                           ByRef dMod As Double,
                           ByRef dDiv As Double)

            Dim lPos As Integer
            Dim dTmp As Double
            Dim sTmp As String

            dTmp = dValue / dDivider
            sTmp = CStr(Format(dTmp, "0.000000"))
            lPos = InStr(sTmp, ".")

            If lPos <> 0 Then
                dDiv = CDbl(Left(sTmp, lPos - 1))
                dMod = (dTmp - dDiv) * dDivider
            Else
                dDiv = dTmp
                dMod = 0
            End If
        End Sub

        ''' <summary>
        ''' Validates the unit serial number and gets the DB key, a.k.a unit ID.
        ''' </summary>
        ''' <param name="product_number">Udbs product number.</param>
        ''' <param name="serial_number">Unit serial number.</param>
        ''' <returns>An integer representing the unit ID.</returns>
        Friend Function GetUnitID(ByVal product_number As String, ByVal serial_number As String) As Integer

            Dim results As New DataTable
            Dim query = "SELECT DISTINCT unit_id " &
                   "FROM product with(nolock), unit with(nolock) " &
                   "WHERE product_id=unit_product_id " &
                   "AND unit_serial_number='" & serial_number & "' " &
                   "AND product_number='" & product_number & "'"

            OpenNetworkRecordSet(results, query)
            If (If(results?.Rows?.Count, 0)) = 0 Then
                Throw New ApplicationException("Unit serial number does not belong to " & product_number & ".")
            Else
                If (If(results?.Rows?.Count, 0)) <> 1 Then
                    Throw _
                        New ApplicationException(
                            $"{(If(results?.Rows?.Count, 0))} units have the same serial number in {product_number}.")

                Else
                    Return KillNullInteger(results(0)("unit_id"))
                End If
            End If
        End Function

        ''' <summary>
        ''' Gets or create a MAC address for the unit specified.
        ''' </summary>
        ''' <param name="product_number">UDBS Product number.</param>
        ''' <param name="product_group">Product Group.</param>
        ''' <param name="serial_number">Unit serial number.</param>
        ''' <param name="MAC_identifier">Mac identifier.</param>
        ''' <param name="new_address">Whether or not a new MAC address has been created.</param>
        ''' <param name="MACAddress">The MAC address.</param>
        ''' <returns>True when a MAC address was successfully obtained or created. False otherwise.</returns>
        Public Function CreateMACAddress(product_number As String,
                                      product_group As String,
                                      serial_number As String,
                                      MAC_identifier As String,
                                      ByRef new_address As Boolean,
                                      ByRef MACAddress As String) As Boolean

            Dim dbrs As New DataTable, sSQL As String
            Dim unitID As Integer
            Dim sStartMAC As String
            Dim lLastMACUsed As Integer
            Dim lMACBlockSize As Integer
            Dim iPrdGrpSeq As Integer
            Dim sMAC As String, dTmp As Double, iNumGrpAva As Integer
            Dim keys As String()
            Dim columns As String()
            Dim columnValues As Object()

            CreateMACAddress = False
            MACAddress = ""

            Try

                ' validating unit SN and get unit id at the same time
                unitID = GetUnitID(product_number, serial_number)

                ' check if SN has existing MAC address
                sSQL = "SELECT ud_string_value FROM udbs_unit_details with(nolock) " &
                       "WHERE ud_unit_id=" & unitID & " " &
                       "AND ud_identifier='" & MAC_identifier & "' " &
                       "AND ud_pg_product_group='" & product_group & "' " &
                       "ORDER BY ud_pg_sequence DESC"
                OpenNetworkRecordSet(dbrs, sSQL)
                If (If(dbrs?.Rows?.Count, 0)) > 0 Then
                    MACAddress = MakeMACAddress(KillNull(dbrs(0)("ud_string_value")) & "")
                    new_address = False
                    Return True
                Else
                    new_address = True
                End If
                dbrs = Nothing

                ' get product group info
                sSQL = "SELECT * " &
                       "FROM udbs_prdgrp with(nolock) , udbs_product_group with(nolock) " &
                       "WHERE prdgrp_product_group=pg_product_group " &
                       "AND prdgrp_product_number='" & product_number & "' " &
                       "AND pg_product_group='" & product_group & "' " &
                       "AND pg_integer_value<pg_float_value " &
                       "ORDER by pg_sequence"
                OpenNetworkRecordSet(dbrs, sSQL)
                If (If(dbrs?.Rows?.Count, 0)) = 0 Then
                    Throw New ApplicationException("Product - Product Group combination not found." & Chr(13) &
                                                   "                      OR" & Chr(13) &
                                                   "All addresses haved been used for this product group.")

                Else
                    iPrdGrpSeq = KillNullInteger(dbrs(0)("pg_sequence"))
                    sStartMAC = KillNull(dbrs(0)("pg_string_value"))
                    lLastMACUsed = KillNullInteger(dbrs(0)("pg_integer_value"))
                    lMACBlockSize = KillNullInteger(dbrs(0)("pg_float_value"))
                    iNumGrpAva = (If(dbrs?.Rows?.Count, 0))
                End If
                dbrs = Nothing

                dTmp = HexToDbl(sStartMAC) + lLastMACUsed '+ 1
                sMAC = DblToHex(dTmp)
                ' checking for uniqueness
                sSQL = "SELECT ud_unit_id FROM udbs_unit_details with(nolock) " &
                       "WHERE ud_string_value='" & sMAC & "' "
                OpenNetworkRecordSet(dbrs, sSQL)
                If (If(dbrs?.Rows?.Count, 0)) > 0 Then
                    Throw _
                        New ApplicationException(
                            "MAC address (" & sMAC & ") already belongs to another unit (id=" &
                            KillNull(dbrs(0)("ud_unit_id")) & ").")

                End If
                dbrs = Nothing

                lLastMACUsed = lLastMACUsed + 1
                If (Len(sMAC) > 12) Or (Len(sMAC) > Len(sStartMAC)) Then
                    Throw New ApplicationException("Wrong HEX returned: " & sMAC & ".")
                End If

                ' update product_group and unit_details
                If (lMACBlockSize - lLastMACUsed <= 5) And (Not iNumGrpAva > 1) Then
                    If (lMACBlockSize - lLastMACUsed) = 0 Then
                        Throw _
                            New ApplicationException(
                                "All MAC addresses have been used for this group after creating this address.")
                    Else
                        Throw _
                            New ApplicationException(
                                "There are only " & lMACBlockSize - lLastMACUsed &
                                " MAC addresses available for this group after creating this address.")
                    End If
                End If
                sSQL = "SELECT pg_integer_value " &
                       "FROM udbs_product_group with(nolock) " &
                       "WHERE pg_product_group='" & product_group & "' " &
                       "AND pg_sequence=" & iPrdGrpSeq
                OpenNetworkRecordSet(dbrs, sSQL)
                If (If(dbrs?.Rows?.Count, 0)) = 0 Then
                    Throw New ApplicationException("Product Group Sequence not found.")
                Else
                    keys = {"pg_product_group", "pg_sequence"}
                    columns = {"pg_product_group", "pg_sequence", "pg_integer_value"}
                    columnValues = {product_group, iPrdGrpSeq, lLastMACUsed}
                    UpdateNetworkRecord(keys, columns, columnValues, "udbs_product_group")
                End If
                dbrs = Nothing

                sSQL = "SELECT * FROM udbs_unit_details with(nolock) " &
                       "WHERE ud_unit_id=" & unitID & " " &
                       "AND ud_pg_product_group='" & product_group & "' " &
                       "AND ud_pg_sequence=" & iPrdGrpSeq & " " &
                       "AND ud_identifier='" & MAC_identifier & "'"
                OpenNetworkRecordSet(dbrs, sSQL)
                If (If(dbrs?.Rows?.Count, 0)) = 0 Then
                    Dim kv As New Dictionary(Of String, Object)

                    kv("ud_unit_id") = unitID
                    kv("ud_pg_product_group") = product_group
                    kv("ud_pg_sequence") = iPrdGrpSeq
                    kv("ud_identifier") = MAC_identifier
                    kv("ud_string_value") = sMAC

                    columns = kv.Keys.ToArray()
                    columnValues = columns.Select(Function(k) kv(k)).ToArray()
                    InsertNetworkRecord(columns, columnValues, "udbs_unit_details")

                Else
                    Throw New ApplicationException("Unit MAC address already existed.")

                End If

                MACAddress = MakeMACAddress(sMAC)

                Return True
            Catch ex As Exception
                MACAddress = ""
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Reads the MAC address and MAC identifier for the unit specified.
        ''' This function is different from <see cref="CreateMACAddress(String, String, String, String, ByRef Boolean, ByRef String)"/> in that it doesn't create a new address
        ''' when it doesn't find one.
        ''' </summary>
        ''' <param name="product_number">Udbs Product number.</param>
        ''' <param name="product_group">Product group.</param>
        ''' <param name="serial_number">unit serial number.</param>
        ''' <remarks>A certain unit could have multiple MAC identifiers and MAC addresses under the same product group.</remarks>
        ''' <returns>A dictionary of Mac_identifier, MACAddress.</returns>
        Public Function ReadMACAddresses(product_number As String,
                                      product_group As String,
                                      serial_number As String) As Dictionary(Of String, String)

            Dim unitID = GetUnitID(product_number, serial_number)

            Dim query = $"SELECT ud_string_value, ud_identifier FROM udbs_unit_details WHERE ud_unit_id ={unitID}  AND ud_pg_product_group ='{product_group}' ORDER BY ud_timestamp"

            Dim records As New DataTable

            OpenNetworkRecordSet(records, query)
            If (If(records?.Rows?.Count, 0)) = 0 Then
                Throw New ApplicationException($"No MAC address and MAC identifier found for unit '{serial_number}'.")

            Else
                Dim results As New Dictionary(Of String, String)

                For Each row As DataRow In records.Rows
                    Dim MAC_identifier = KillNull(row("ud_identifier"))
                    Dim MACAddress = MakeMACAddress(KillNull(row("ud_string_value")))

                    results.Add(MAC_identifier, MACAddress)
                Next

                Return results
            End If
        End Function

        ''' <summary>
        ''' Gets or creates a block of MAC addresses for the given MAC identifier.
        ''' </summary>
        ''' <param name="product_number">UDBS Product number.</param>
        ''' <param name="product_group">Product Group.</param>
        ''' <param name="serial_number">Unit serial number.</param>
        ''' <param name="MAC_identifier">MAC identifier.</param>
        ''' <param name="increment">Number of MAC addresses matching with the identifier.</param>
        ''' <param name="new_address">Whether or not new MAC addresses have been created.</param>
        ''' <param name="MACAddress">The MAC address.</param>
        ''' <param name="allMACAddresses">A string of all the MAC addresses found or created matching the MAC identifier.</param>
        ''' <returns>True when MAC addresses were successfully obtained or created. False otherwise.</returns>
        Public Function CreateBlockMACAddress(product_number As String,
                                           product_group As String,
                                           serial_number As String,
                                           MAC_identifier As String,
                                           ByRef increment As Integer,
                                           ByRef new_address As Boolean,
                                           ByRef MACAddress As String,
                                           ByRef allMACAddresses As String) As Boolean

            Dim dbrs As New DataTable, sSQL As String
            Dim unitID As Integer
            Dim sStartMAC As String
            Dim lLastMACUsed As Integer
            Dim lMACBlockSize As Integer
            Dim iPrdGrpSeq As Integer
            Dim sMAC As String, dTmp As Double, iNumGrpAva As Integer
            Dim i As Integer, sTmp As String

            If increment <= 0 Then
                Throw New ApplicationException("Invaild increment.")
            End If

            Try


                CreateBlockMACAddress = False
                MACAddress = ""
                allMACAddresses = ""


                ' validating unit SN and get unit id at the same time
                unitID = GetUnitID(product_number, serial_number)

                ' check if SN has existing MAC address
                sSQL = "SELECT ud_string_value FROM udbs_unit_details  with(nolock) " &
                       "WHERE ud_unit_id=" & unitID & " " &
                       "AND (ud_identifier='" & MAC_identifier & "' " &
                       "     OR ud_identifier LIKE '" & MAC_identifier & " ~*INC%*~' ) " &
                       "AND ud_pg_product_group='" & product_group & "' " &
                       "ORDER BY ud_pg_sequence DESC, ud_identifier"
                OpenNetworkRecordSet(dbrs, sSQL)
                Dim rCount = 0
                If (If(dbrs?.Rows?.Count, 0)) > 0 Then
                    MACAddress = MakeMACAddress(KillNull(dbrs(rCount)("ud_string_value")) & "")
                    rCount += 1
                    sTmp = MACAddress
                    Do While Not (rCount >= dbrs?.Rows?.Count)
                        sTmp = sTmp & "," & MakeMACAddress(KillNull(dbrs(rCount)("ud_string_value")) & "")
                        rCount += 1
                    Loop
                    If increment <> (If(dbrs?.Rows?.Count, 0)) Then
                        logger.Error(
                            "Number of addresses with matching identifer does not equal to increment!" & Chr(13) &
                            "Please check increment arguement for number found.")
                        increment = (If(dbrs?.Rows?.Count, 0))
                    End If
                    allMACAddresses = sTmp
                    new_address = False
                    Return True
                Else
                    new_address = True
                End If
                dbrs = Nothing

                ' get product group info
                sSQL = "SELECT * " &
                       "FROM udbs_prdgrp with(nolock) , udbs_product_group  with(nolock) " &
                       "WHERE prdgrp_product_group=pg_product_group " &
                       "AND prdgrp_product_number='" & product_number & "' " &
                       "AND pg_product_group='" & product_group & "' " &
                       "AND (pg_integer_value+" & increment - 1 & "<pg_float_value) " &
                       "ORDER by pg_sequence"
                OpenNetworkRecordSet(dbrs, sSQL)
                If (If(dbrs?.Rows?.Count, 0)) = 0 Then
                    Throw New ApplicationException("Product - Product Group combination not found." & Chr(13) &
                                                   "                      OR" & Chr(13) &
                                                   "Number of addresses available is not sufficient (Increment=" &
                                                   increment & ").")
                Else
                    iPrdGrpSeq = KillNullInteger(dbrs(0)("pg_sequence"))
                    sStartMAC = KillNull(dbrs(0)("pg_string_value"))
                    lLastMACUsed = KillNullInteger(dbrs(0)("pg_integer_value"))
                    lMACBlockSize = KillNullInteger(dbrs(0)("pg_float_value"))
                    iNumGrpAva = (If(dbrs?.Rows?.Count, 0))
                End If
                dbrs = Nothing

                dTmp = HexToDbl(sStartMAC) + lLastMACUsed '+ 1
                sMAC = DblToHex(dTmp)
                sSQL = ""
                For i = 2 To increment
                    sSQL = sSQL & ",'" & DblToHex(dTmp + i - 1) & "'"
                Next i
                ' checking for uniqueness
                sSQL =
                    "SELECT unit_serial_number, ud_string_value FROM unit with(nolock) , udbs_unit_details  with(nolock) " &
                    "WHERE unit_id=ud_unit_id AND ud_string_value IN ('" & sMAC & "'" & sSQL & ") "
                OpenNetworkRecordSet(dbrs, sSQL)
                If (If(dbrs?.Rows?.Count, 0)) > 0 Then
                    Throw _
                        New ApplicationException(
                            "MAC address (" & MakeMACAddress(KillNull(dbrs(0)("ud_string_value"))) &
                            ") already belongs to another unit (SN=" & KillNull(dbrs(0)("unit_serial_number")) & ").")

                End If
                dbrs = Nothing

                lLastMACUsed = lLastMACUsed + increment
                If (Len(sMAC) > 12) Or (Len(sMAC) > Len(sStartMAC)) Then
                    Throw New ApplicationException("Wrong HEX returned: " & sMAC & ".")
                End If

                ' update product_group and unit_details
                sSQL = "SELECT pg_integer_value " &
                       "FROM udbs_product_group  with(nolock) " &
                       "WHERE pg_product_group='" & product_group & "' " &
                       "AND pg_sequence=" & iPrdGrpSeq
                OpenNetworkRecordSet(dbrs, sSQL)
                If (If(dbrs?.Rows?.Count, 0)) = 0 Then
                    Throw New ApplicationException("Product Group Sequence not found.")
                Else
                    Dim keys As String() = {"pg_product_group", "pg_sequence"}
                    Dim columnNames As String() = keys.Concat({"pg_integer_value"}).ToArray()
                    Dim columnValues As Object() = {product_group, iPrdGrpSeq, lLastMACUsed}
                    UpdateNetworkRecord(keys, columnNames, columnValues, "udbs_product_group")
                End If
                dbrs = Nothing

                sSQL = "SELECT * FROM udbs_unit_details  with(nolock) " &
                       "WHERE ud_unit_id=" & unitID & " " &
                       "AND ud_pg_product_group='" & product_group & "' " &
                       "AND ud_pg_sequence=" & iPrdGrpSeq & " " &
                       "AND ud_identifier='" & MAC_identifier & "'"
                OpenNetworkRecordSet(dbrs, sSQL)
                If (If(dbrs?.Rows?.Count, 0)) = 0 Then
                    For i = 1 To increment
                        Dim dr As DataRow = dbrs.NewRow()

                        dr("ud_unit_id") = unitID
                        dr("ud_pg_product_group") = product_group
                        dr("ud_pg_sequence") = iPrdGrpSeq
                        If i = 1 Then
                            dr("ud_identifier") = MAC_identifier
                            dr("ud_string_value") = sMAC
                            sSQL = MakeMACAddress(sMAC)
                        Else
                            dr("ud_identifier") = MAC_identifier & " ~*INC" & i & "*~"
                            sTmp = DblToHex(dTmp + i - 1)
                            dr("ud_string_value") = sTmp
                            sSQL = sSQL & "," & MakeMACAddress(sTmp)
                        End If
                        InsertNetworkRecord(
                            dbrs.Columns.Cast(Of DataColumn)().Select(Function(dc) dc.ColumnName).Skip(1).ToArray(),
                            dr.ItemArray.Skip(1).ToArray(), "udbs_unit_details")
                    Next i
                Else
                    Throw New ApplicationException("Unit MAC address already existed.")

                End If
                dbrs = Nothing

                MACAddress = MakeMACAddress(sMAC)
                allMACAddresses = sSQL

                CreateBlockMACAddress = True

                ' Warning message
                If (lMACBlockSize - lLastMACUsed <= increment) And (Not iNumGrpAva > 1) Then
                    sSQL = "SELECT pg_sequence, pg_integer_value, pg_float_value " &
                           "FROM udbs_prdgrp with(nolock) , udbs_product_group  with(nolock) " &
                           "WHERE prdgrp_product_group=pg_product_group " &
                           "AND prdgrp_product_number='" & product_number & "' " &
                           "AND pg_product_group='" & product_group & "' " &
                           "AND pg_integer_value<pg_float_value " &
                           "ORDER by pg_sequence"
                    OpenNetworkRecordSet(dbrs, sSQL)
                    If (If(dbrs?.Rows?.Count, 0)) = 0 Then
                        Throw _
                            New ApplicationException(
                                "All MAC addresses have been used for this group after creating this address.")
                    Else
                        sTmp = "MAC addresses availability summary (increment=" & increment & "):  "
                        For Each dr As DataRow In dbrs.Rows
                            sTmp = sTmp & Chr(13) &
                                   (KillNullDouble(dr("pg_float_value")) - KillNullInteger(dr("pg_integer_value"))) &
                                   " address(es) left in block " & KillNull(dr("pg_sequence"))

                        Next
                        logger.Info(sTmp)
                    End If
                    dbrs = Nothing
                End If

                dbrs = Nothing
            Catch ex As Exception
                dbrs = Nothing

                Throw
            End Try
        End Function


        Private Function MakeMACAddress(sHex As String) As String

            Dim i As Integer
            Dim sTmp As String

            MakeMACAddress = ""
            If Len(sHex) > 12 Then
                MsgBox("Wrong Hex string length.", vbExclamation)
                Exit Function
            End If

            sTmp = sHex
            Do While Len(sTmp) < 12
                sTmp = "0" & sTmp
            Loop

            For i = 1 To 5
                sTmp = Left(sTmp, (6 - i) * 2) & " " & Mid(sTmp, (6 - i) * 2 + 1)
            Next i

            MakeMACAddress = sTmp
        End Function

        Public Function ProductGroupList_fromPrdNum(product_number As String,
                                                    ByRef rsPrdGrpList As DataTable) As Boolean
            Dim rsTmp As New DataTable, sSQL As String
            Try


                ProductGroupList_fromPrdNum = False


                sSQL = "SELECT prdgrp_product_number AS product_number, prdgrp_product_group AS product_group " &
                       "FROM udbs_prdgrp  with(nolock) " &
                       "WHERE prdgrp_product_number='" & product_number & "' " &
                       "ORDER BY prdgrp_product_group"

                OpenNetworkRecordSet(rsTmp, sSQL)
                rsPrdGrpList = rsTmp.Copy()

                ProductGroupList_fromPrdNum = True

                rsTmp = Nothing

            Catch ex As Exception
                rsTmp = Nothing
                logger.Error(ex)
                Throw
            End Try
        End Function

        Public Function ProductGroupList_fromPrdGrp(product_group As String,
                                                    ByRef rsPrdGrpList As DataTable) As Boolean


            Dim rsTmp As New DataTable, sSQL As String
            Try


                ProductGroupList_fromPrdGrp = False

                sSQL = "SELECT prdgrp_product_number AS product_number, prdgrp_product_group AS product_group " &
                       "FROM udbs_prdgrp with(nolock) " &
                       "WHERE prdgrp_product_group='" & product_group & "' " &
                       "ORDER BY prdgrp_product_number"

                OpenNetworkRecordSet(rsTmp, sSQL)
                rsPrdGrpList = rsTmp.Copy()

                ProductGroupList_fromPrdGrp = True

                rsTmp = Nothing


                Exit Function

            Catch ex As Exception
                rsTmp = Nothing
                logger.Error(ex)
                Throw
            End Try
        End Function

        Friend Function RemoveProductFromGroup(product_number As String,
                                               product_group As String) As Boolean


            Dim sSQL As String
            Try


                RemoveProductFromGroup = False


                sSQL = "DELETE FROM udbs_prdgrp " &
                       "WHERE prdgrp_product_number='" & product_number & "' " &
                       "AND prdgrp_product_group='" & product_group & "' "
                ExecuteNetworkQuery(sSQL)

                RemoveProductFromGroup = True


            Catch ex As Exception
                logger.Error(ex)
                Throw
            End Try
        End Function


        Friend Function AddProductToGroup(product_number As String,
                                          product_group As String) As Boolean
            On Error GoTo errHandler

            Dim rsTmp As New DataTable, sSQL As String

            AddProductToGroup = False

            If Trim(product_number) = "" Then
                MsgBox("Missing product number.", vbExclamation)
                GoTo exitHere
            End If
            If Trim(product_group) = "" Then
                MsgBox("Missing product group.", vbExclamation)
                GoTo exitHere
            End If


            sSQL = "SELECT product_id FROM product with(nolock) " &
                   "WHERE product_number='" & product_number & "'"

            OpenNetworkRecordSet(rsTmp, sSQL)
            If (If(rsTmp?.Rows?.Count, 0)) = 0 Then
                Throw New ApplicationException("Product Number not defined in UDBS product table.")
            End If
            rsTmp = Nothing

            sSQL = "SELECT * FROM udbs_prdgrp with(nolock) " &
                   "WHERE prdgrp_product_number='" & product_number & "' " &
                   "AND prdgrp_product_group='" & product_group & "' "
            OpenNetworkRecordSet(rsTmp, sSQL)
            If (If(rsTmp?.Rows?.Count, 0)) = 0 Then
                Dim keys As String() = {"prdgrp_product_number", "prdgrp_product_group"}
                Dim columnNames As String() = keys.Concat({"prdgrp_product_number", "prdgrp_product_group"}).ToArray()
                Dim columnVals As Object() = {product_number, product_group, product_number, product_group}
            Else
                Throw New ApplicationException("Product Group already existed.")
            End If
            rsTmp = Nothing

            AddProductToGroup = True

exitHere:
            rsTmp = Nothing


            Exit Function

errHandler:
            rsTmp = Nothing

            logger.Error(Err.Description)
            AddProductToGroup = False
        End Function


        Public Function ProductGroupDetailsList(product_group As String,
                                                ByRef rsPrdGrpdetails As DataTable) As Boolean

            On Error GoTo errHandler

            Dim rsTmp As New DataTable, sSQL As String

            ProductGroupDetailsList = False


            sSQL = "SELECT pg_product_group AS product_group, " &
                   "pg_sequence AS sequence, " &
                   "pg_description AS product_group_description, " &
                   "pg_string_value AS start_block_address, " &
                   "pg_float_value AS block_size, " &
                   "pg_integer_value AS last_address_requrested " &
                   "FROM udbs_product_group  with(nolock) " &
                   "WHERE pg_product_group='" & product_group & "' " &
                   "ORDER BY pg_sequence"

            OpenNetworkRecordSet(rsTmp, sSQL)
            rsPrdGrpdetails = rsTmp.Copy()

            ProductGroupDetailsList = True

            rsTmp = Nothing


            Exit Function

errHandler:
            rsTmp = Nothing

            logger.Error(Err.Description)
            ProductGroupDetailsList = False
        End Function

        Friend Function AddProductGroupDetails(product_group As String,
                                               group_description As String,
                                               start_block_address As String,
                                               block_size As Integer,
                                               ByRef sequence As Integer) As Boolean
            On Error GoTo errHandler

            Dim rsTmp As New DataTable, sSQL As String
            Dim iSeq As Integer

            AddProductGroupDetails = False

            If Trim(product_group) = "" Then
                Throw New ApplicationException("Missing product group.")
            End If
            If Trim(group_description) = "" Then
                Throw New ApplicationException("Missing group description.")
            End If
            If Trim(start_block_address) = "" Then
                Throw New ApplicationException("Missing starting block address.")
            End If
            If block_size < 0 Then
                Throw New ApplicationException("Block size must be greater than 0.")
            End If


            sSQL = "SELECT * FROM udbs_product_group with(nolock) " &
                   "WHERE pg_product_group='" & product_group & "' " &
                   "ORDER BY pg_sequence DESC"
            OpenNetworkRecordSet(rsTmp, sSQL)
            If (If(rsTmp?.Rows?.Count, 0)) = 0 Then
                iSeq = 1
            Else
                iSeq = KillNullInteger(rsTmp(0)("pg_sequence")) + 1
            End If

            Dim dr As DataRow = rsTmp.NewRow()

            dr("pg_product_group") = product_group
            dr("pg_sequence") = iSeq
            dr("pg_description") = group_description
            dr("pg_string_value") = start_block_address
            dr("pg_integer_value") = 0
            dr("pg_float_value") = block_size
            InsertNetworkRecord(
                rsTmp.Columns.Cast(Of DataColumn)().Skip(1).Select(Function(dc) dc.ColumnName).ToArray(),
                dr.ItemArray.Skip(1).ToArray(), "udbs_product_group")

            rsTmp = Nothing

            sequence = iSeq
            AddProductGroupDetails = True

exitHere:
            rsTmp = Nothing


            Exit Function

errHandler:
            rsTmp = Nothing
            logger.Error(Err.Description)
            AddProductGroupDetails = False
        End Function


        Public Function GetNextSNSequence(product_number As String,
                                          product_group As String,
                                          ByRef SN_Sequence As Integer) As Boolean

            Dim dbrs As New DataTable, sSQL As String
            Dim sStartBlock As String
            Dim lLastElementUsed As Integer
            Dim lBlockSize As Integer
            Dim iPrdGrpSeq As Integer
            Dim lSNSequence As Integer
            Dim iNumGrpAva As Integer

            On Error GoTo ProcedureErr

            GetNextSNSequence = False
            SN_Sequence = 0

            ' get product group info
            sSQL = "SELECT * " &
                   "FROM udbs_prdgrp with(nolock) , udbs_product_group with(nolock) " &
                   "WHERE prdgrp_product_group=pg_product_group " &
                   "AND prdgrp_product_number='" & product_number & "' " &
                   "AND pg_product_group='" & product_group & "' " &
                   "AND pg_integer_value<pg_float_value " &
                   "ORDER by pg_sequence"
            OpenNetworkRecordSet(dbrs, sSQL)
            If (If(dbrs?.Rows?.Count, 0)) = 0 Then
                Throw _
                    New ApplicationException(
                        "Product - Product Group combination not found." & Chr(13) & "                      OR" &
                        Chr(13) & "All addresses haved been used for this product group.")

            Else
                iPrdGrpSeq = KillNullInteger(dbrs(0)("pg_sequence"))
                sStartBlock = KillNull(dbrs(0)("pg_string_value"))
                lLastElementUsed = KillNullInteger(dbrs(0)("pg_integer_value"))
                lBlockSize = KillNullInteger(dbrs(0)("pg_float_value"))
                iNumGrpAva = dbrs.Rows.Count
            End If
            dbrs = Nothing

            lSNSequence = CInt(sStartBlock) + lLastElementUsed
            lLastElementUsed = lLastElementUsed + 1

            If lLastElementUsed > lBlockSize Then
                Throw New ApplicationException("Element exceeded block size.")
            End If

            ' update product_group and unit_details
            If (lBlockSize - lLastElementUsed <= 5) And (Not iNumGrpAva > 1) Then
                If (lBlockSize - lLastElementUsed) = 0 Then
                    logger.Info("All elements have been used for this group after creating this element.")
                Else
                    logger.Info(
                        "There are only " & lBlockSize - lLastElementUsed &
                        " elements available for this group after this.")
                End If
            End If
            sSQL = "SELECT pg_integer_value " &
                   "FROM udbs_product_group  with(nolock) " &
                   "WHERE pg_product_group='" & product_group & "' " &
                   "AND pg_sequence=" & iPrdGrpSeq
            OpenNetworkRecordSet(dbrs, sSQL)
            If (If(dbrs?.Rows?.Count, 0)) = 0 Then
                Throw New ApplicationException("Product Group Sequence not found.")

            Else

                Dim keys As String() = {"pg_product_group", "pg_sequence"}
                Dim columnNames As String() = keys.Concat({"pg_integer_value"}).ToArray()
                Dim columnValues As Object() = {product_group, iPrdGrpSeq, lLastElementUsed}
                UpdateNetworkRecord(keys, columnNames, columnValues, "udbs_product_group")

            End If
            dbrs = Nothing

            SN_Sequence = lSNSequence

            GetNextSNSequence = True

exitHere:
            dbrs = Nothing

            Exit Function

ProcedureErr:
            dbrs = Nothing

            SN_Sequence = 0
            logger.Error(Err.Description)
            GetNextSNSequence = False
        End Function

        Public Function GetUnitVariance(product_number As String,
                                        unit_SN As String,
                                        ByRef UnitVariance As Integer) As Boolean

            Dim dbrs As New DataTable, sSQL As String

            Try

                UnitVariance = 0
                ' get product group info
                sSQL = "SELECT ud_pg_sequence " &
                       "FROM product with(nolock) , unit with(nolock) , udbs_unit_details with(nolock)  " &
                       "WHERE product_id=unit_product_id " &
                       "AND unit_id=ud_unit_id " &
                       "AND ud_identifier='PRD_VAR' " &
                       "AND product_number='" & product_number & "' " &
                       "AND ud_pg_product_group='" & product_number & "_variance' " &
                       "AND unit_serial_number='" & unit_SN & "'"
                OpenNetworkRecordSet(dbrs, sSQL)
                If (If(dbrs?.Rows?.Count, 0)) = 0 Then
                    logger.Error("No unit variance information found.")
                    Return False
                Else
                    If (If(dbrs?.Rows?.Count, 0)) > 1 Then
                        logger.Error("More than 1 variance found." & vbCr & "Please contact UDBS administrator.")
                        Return False
                    End If
                    UnitVariance = KillNullInteger(dbrs(0)("ud_pg_sequence"))
                End If
                dbrs = Nothing

                Return True

            Catch ex As Exception
                UnitVariance = 0
                Throw
            End Try
        End Function


        Public Sub New()
        End Sub
    End Class
End Namespace