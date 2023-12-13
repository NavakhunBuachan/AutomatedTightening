Option Explicit On
Option Compare Text
Option Infer On
Option Strict On


Namespace MasterInterface
    Public Class CProduct
        Implements IDisposable

        ''' <summary>
        ''' Calling CProduct.UnitExists() and CProduct.AddSNwVar(...)
        ''' sequentially, from multiple threads, could lead to a race
        ''' condition and the creation of duplicated units.
        ''' AddSNwVar(...) performs the check and unit creation
        ''' in a synchronization lock, preventing the creation of
        ''' duplicated units.
        ''' This is 'shared' to prevent multiple CProduct instances
        ''' from creating duplicated units.
        ''' </summary>
        Private Shared mUnitCreationLock As Object = New Object()

        ' Table identification
        Private Const mProductTable As String = "product"
        Private Const mProductFamilyTable As String = "family"
        Private Const mUnitTable As String = "unit"

        ' Object State
        Private mREADONLY As Boolean = False
        Private mObjectLoaded As Boolean = False

        ' Product Properties
        Private mProductID As Integer
        Private mNumber As String
        Private mRelease As Integer
        Private mDescriptor As String
        Private mDescription As String
        Private mCreatedBy As String
        Private mCreatedDate As Date
        Private mReleaseReason As String
        Private mSNProdCode As String
        Private mSNTemplate As String
        Private mSNLastUnit As Integer
        Private mFamily As String
        Private mProductInfo As New DataTable
        Private mUnitInfo As Lazy(Of DataTable)

        ' Error Handling
        'Private mErrorDescription As String


        '**********************************************************************
        '* Properties
        '**********************************************************************

        ' Object Information


        ' Product Information
        Public ReadOnly Property ProductID As Integer
            Get
                Return mProductID
            End Get
        End Property

        Public ReadOnly Property Number As String
            Get
                Return mNumber
            End Get
        End Property

        Public ReadOnly Property Release As Integer
            Get
                Return mRelease
            End Get
        End Property

        Public ReadOnly Property Descriptor As String
            Get
                Return mDescriptor
            End Get
        End Property

        Public ReadOnly Property Description As String
            Get
                Return mDescription
            End Get
        End Property

        Public ReadOnly Property CreatedBy As String
            Get
                Return mCreatedBy
            End Get
        End Property

        Public ReadOnly Property CreatedDate As Date
            Get
                Return mCreatedDate
            End Get
        End Property

        Public ReadOnly Property ReleaseReason As String
            Get
                Return mReleaseReason
            End Get
        End Property

        Public ReadOnly Property SNProdCode As String
            Get
                Return mSNProdCode
            End Get
        End Property

        Public ReadOnly Property SNTemplate As String
            Get
                Return mSNTemplate
            End Get
        End Property

        Public ReadOnly Property SNLastUnit As Integer
            Get
                Return mSNLastUnit
            End Get
        End Property

        Public ReadOnly Property Family As String
            Get
                Return mFamily
            End Get
        End Property

        ' Candidate for removal.
        Private ReadOnly Property ProductInfo As DataTable
            Get
                Return mProductInfo.Copy()
            End Get
        End Property

        ' Candidate for removal.
        Private ReadOnly Property UnitInfo As DataTable
            Get
                Return mUnitInfo?.Value?.Copy()
            End Get
        End Property

        ''' <summary>
        ''' Whether or not this Product object has been loaded.
        ''' </summary>
        Friend ReadOnly Property Loaded As Boolean
            Get
                Return mObjectLoaded
            End Get
        End Property

        '**********************************************************************
        '* Methods
        '**********************************************************************

        Public Function GetProduct(ProductNumber As String,
                                   ProductRelease As Integer) _
            As ReturnCodes
            ' Function returns all information about a specified product/release
            ' modified by Billy Nov 1, 2001
            Try

                ' Object has been created for read mode
                mREADONLY = False

                Dim localProductID As Integer
                Dim DesiredRelease As Integer

                ' If latest release has been requested
                If ProductRelease = 0 Then
                    If GetLatestRelease(ProductNumber, DesiredRelease) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        ' Problem resolving release
                        Return ReturnCodes.UDBS_ERROR
                    End If
                Else
                    DesiredRelease = ProductRelease
                End If

                ' Get Product ID
                If GetProductID(ProductNumber, DesiredRelease, localProductID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Problem resolving release
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Load product object by ID
                Return GetProductByID(localProductID)

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function


        ' Function returns all information about a specified product/release
        Public Function GetProductByID(ProductID As Integer) As ReturnCodes
            ' Object has been created for read mode
            mREADONLY = True

            Try
                Dim sqlQuery As String
                sqlQuery = "SELECT * " &
                           "FROM " & mProductTable & " p with(nolock), " & mProductFamilyTable & " f with(nolock) " &
                           "WHERE product_family_id=family_id AND product_id=" & CStr(ProductID)
                If QueryNetworkDB(sqlQuery, mProductInfo) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Throw New Exception($"Error querying for product {ProductID}")
                End If

                If mProductInfo.Rows.Count = 0 Then
                    logger.Info($"No such product ID: {ProductID}")
                    Return ReturnCodes.UDBS_ERROR
                ElseIf mProductInfo.Rows.Count > 1 Then
                    ' Throw an exception; this will cause a system error to be logged.
                    Throw New Exception($"Multiple rows for product ID: {ProductID}")
                End If

                ' Fill Product Properties
                mProductID = KillNullInteger(mProductInfo(0)("product_id"))
                mNumber = KillNull(mProductInfo(0)("product_number"))
                mRelease = KillNullInteger(mProductInfo(0)("product_release"))
                mDescriptor = KillNull(mProductInfo(0)("product_descriptor"))
                mDescription = KillNull(mProductInfo(0)("product_description"))
                mCreatedBy = KillNull(mProductInfo(0)("product_created_by"))
                mCreatedDate = KillNullDate(mProductInfo(0)("product_created_date"))
                mReleaseReason = KillNull(mProductInfo(0)("product_release_reason"))
                mSNProdCode = KillNull(mProductInfo(0)("product_sn_prod_code"))
                mSNTemplate = KillNull(mProductInfo(0)("product_sn_template"))
                ' temporary fix for buildlocation, hardcode this for Ottawa build
                sqlQuery = KillNull(mProductInfo(0)("product_sn_last_unit"))
                If sqlQuery = "" Then sqlQuery = "0"
                mSNLastUnit = GetLastSN(sqlQuery, "J", False)
                mFamily = KillNull(mProductInfo(0)("family_name"))

                ' load units for a specific product_id!!!
                ' only units related to the product_id are pulled, not all units belong to that product number

                'TODO: What if there are 1 million rows???

                ' Defer
                mUnitInfo = New Lazy(Of DataTable)(Function()
                                                       Dim mrs As DataTable = Nothing
                                                       Dim uSql = "SELECT * " &
                                                                  "FROM " & mUnitTable &
                                                                  " with(nolock) WHERE unit_product_id=" & CStr(mProductID) &
                                                                  " ORDER BY unit_serial_number"
                                                       If QueryNetworkDB(uSql, mrs) <> ReturnCodes.UDBS_OP_SUCCESS Then
                                                           Throw New Exception($"Error querying for product information: {mProductID}")
                                                       End If
                                                       Return mrs
                                                   End Function)

                ' Object has been properly loaded
                mObjectLoaded = True

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function


        ''' <summary>
        ''' Attempts to load a product based on a Product Number/Serial Number combination.
        ''' Uses the Product Release of the unit specified by the product number.
        ''' </summary>
        ''' <param name="ProductNumber">The product number.</param>
        ''' <param name="SerialNumber">The serial number of the unit.</param>
        ''' <returns>The outcome of the operation.</returns>
        ''' <remarks>
        ''' This method cannot handle multiple units with the same serial number.
        ''' This is a valid condition, if they have different Oracle part numbers.
        ''' </remarks>
        Public Function GetUnit(
                ProductNumber As String,
                SerialNumber As String) As ReturnCodes

            Try
                ' Check for a SINGLE entry when product number and serial number are combined
                Dim rsTemp As New DataTable
                Dim sqlQuery = "SELECT * FROM " & mProductTable & " with(nolock),  " & mUnitTable & " with(nolock) " &
                           "WHERE product_id=unit_product_id " &
                           "AND product_number = '" & ProductNumber & "' " &
                           "and unit_serial_number = '" & SerialNumber & "' "
                OpenNetworkRecordSet(rsTemp, sqlQuery)

                If (If(rsTemp?.Rows?.Count, 0)) = 1 Then
                    ' There is only one matching PN/SN
                    Dim ProductRelease = KillNullInteger(rsTemp(0)("product_release"))

                    ' Load product object
                    If GetProduct(ProductNumber, ProductRelease) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        ' Could not load product
                        LogError(New Exception($"Product not found: " & ProductNumber & "."))
                        Return ReturnCodes.UDBS_ERROR
                    End If

                    Return ReturnCodes.UDBS_OP_SUCCESS
                ElseIf (If(rsTemp?.Rows?.Count, 0)) > 1 Then
                    ' Duplicate entries.
                    LogError(New Exception($"Duplicate entries for serial number {SerialNumber} and product number {ProductNumber}."))
                    Return ReturnCodes.UDBS_ERROR
                Else
                    LogError(New Exception($"Unit with serial number {SerialNumber} and product number {ProductNumber} not found."))
                    Return ReturnCodes.UDBS_ERROR
                End If
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ' Candidate for removal.
        Private Function AddProduct(ProductNumber As String,
                                   ProductRelease As Integer,
                                   Descriptor As String,
                                   Description As String,
                                   EmployeeNumber As String,
                                   ReleaseReason As String,
                                   SNPCode As String,
                                   SNTemplate As String,
                                   SNLastUnit As String,
                                   ProductFamily As String) _
            As ReturnCodes
            ' Function adds new product to UDBS

            Try
                ' Ensure that this is a fresh product object
                If mREADONLY = True Or mObjectLoaded = True Then
                    LogError(New Exception($"Product object already set to read only mode."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                Dim FamilyID As Integer

                ' Get server date/time
                Dim ServerTime As Date
                CUtility.Utility_GetServerTime(ServerTime)

                ' Get Product Family id
                If GetFamilyID(ProductFamily, FamilyID) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' Could not get family info
                    Return ReturnCodes.UDBS_ERROR

                End If

                ' All new Products must have release 1
                If ProductRelease < 1 Then
                    ProductRelease = 1
                End If

                ' Add the new product
                Dim sqlQuery As String
                Dim rsTemp As New DataTable


                sqlQuery = "SELECT * FROM " & mProductTable &
                           " with(nolock) WHERE product_number='" & ProductNumber & "' " &
                           "AND product_release=" & CStr(ProductRelease)
                OpenNetworkRecordSet(rsTemp, sqlQuery)
                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                    ' Product exists...
                    LogError(New Exception($"Product {ProductNumber} already exists."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' column name array
                Dim colummnNames() As String = {"product_number", "product_release", "product_descriptor",
                                                "product_description", "product_created_by", "product_created_date",
                                                "product_release_reason",
                                                "product_sn_prod_code", "product_sn_template", "product_sn_last_unit",
                                                "product_family_id"}
                ' column values array
                Dim colummnValues() As Object = {UCase(Trim(ProductNumber)), ProductRelease, Trim(Descriptor),
                                                 Trim(Description), Trim(EmployeeNumber), ServerTime, Trim(ReleaseReason),
                                                 UCase(Trim(SNPCode)), Trim(SNTemplate), SNLastUnit,
                                                 FamilyID}

                InsertNetworkRecord(colummnNames, colummnValues, mProductTable)

                Return GetProduct(ProductNumber, ProductRelease)
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        ''' <summary>
        ''' Checks if the unit exists.
        ''' </summary>
        ''' <param name="SerialNumber">The unit's serial number.</param>
        ''' <returns>True if the unit exists, false otherwise.</returns>
        Public Function UnitExists(SerialNumber As String) As Boolean
            ' Product object must be loaded
            If mObjectLoaded = False Then
                LogError(New Exception($"Product object has not been loaded."))
                Return False
            End If

            Return UnitExists(SerialNumber, mNumber)
        End Function

        ''' <summary>
        ''' Checks if the unit exists.
        ''' </summary>
        ''' <param name="SerialNumber">The unit's serial number.</param>
        ''' <param name="productNumber">The unit's product number.</param>
        ''' <returns>True if the unit exists, false otherwise.</returns>
        Public Shared Function UnitExists(SerialNumber As String, productNumber As String) As Boolean
            Dim rsTemp As New DataTable
            Try
                Dim sqlQuery As String
                sqlQuery = "SELECT unit.* " &
                           "FROM " & mUnitTable & " with(nolock), " & mProductTable & " with(nolock) " &
                           "WHERE product_id=unit_product_id " &
                           "AND unit_serial_number = '" & SerialNumber & "' " &
                           "AND product_number = '" & productNumber & "'"
                OpenNetworkRecordSet(rsTemp, sqlQuery)
                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                    ' Unit exists.
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                DatabaseSupport.LogErrorInDatabase(ex, "Unit", "", 0, productNumber, SerialNumber)
                Return False
            Finally
                rsTemp?.Dispose()
            End Try
        End Function

        '**********************************************************************
        '* Support Functions
        '**********************************************************************

        Private Function GetLatestRelease(ProductNumber As String,
                                          ByRef ProductRelease As Integer) _
            As ReturnCodes
            ' Function returns the latest product release of the specified product

            Dim rsTemp As New DataTable
            Try
                Dim sqlQuery As String = "SELECT product_number, product_release FROM " & mProductTable &
                       " with(nolock) WHERE product_number = '" & ProductNumber & "' ORDER BY product_release"
                OpenNetworkRecordSet(rsTemp, sqlQuery)
                If (If(rsTemp?.Rows?.Count, 0)) < 1 Then
                    ' Product not found
                    LogError(New Exception("Proudct not found: " & ProductNumber & "."))
                    Return ReturnCodes.UDBS_ERROR
                End If
                Dim dr = rsTemp.AsEnumerable().Last()
                ProductRelease = KillNullInteger(dr("product_release"))
                rsTemp = Nothing

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception

                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR

            Finally
                rsTemp?.Dispose()
            End Try

        End Function


        Private Function GetNumberScheme(ProductNumber As String,
                                         ProductRelease As Double,
                                         ByRef ProductIdentifier As String,
                                         ByRef Scheme As String,
                                         ByRef LastNumericCode As Integer,
                                         ByRef ReleaseCode As String) _
            As ReturnCodes
            ' Function retrieves the unit numbering scheme
            Dim sqlQuery As String
            Dim rsProducts As New DataTable
            Dim localSNLastUnit As String

            Try
                sqlQuery = "SELECT * " &
                           "FROM product with(nolock) " &
                           "WHERE product_number = '" & ProductNumber & "' " &
                           "AND product_release = " & CStr(ProductRelease)
                OpenNetworkRecordSet(rsProducts, sqlQuery)

                If (If(rsProducts?.Rows?.Count, 0)) <> 1 Then
                    ' Product not found
                    rsProducts = Nothing
                    LogError(New Exception("Unable to retrieve specified Product/Release."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Check to make sure a scheme exists
                If IsDBNull(rsProducts(0)("product_sn_template")) Then
                    LogError(New Exception("No serial number scheme exists for this product."))
                    Return ReturnCodes.UDBS_ERROR
                End If

                Scheme = KillNull(rsProducts(0)("product_sn_template"))
                ' hard coded the buildlocation to Ottawa=J
                localSNLastUnit = KillNull(rsProducts(0)("product_sn_last_unit"))
                If localSNLastUnit = "" Then localSNLastUnit = "0"
                LastNumericCode = GetLastSN(localSNLastUnit, "J", True)
                mSNLastUnit = LastNumericCode
                'If isdbnull(rsProducts("product_sn_last_unit").Value) Then
                '    LastNumericCode = 1
                'Else
                '    LastNumericCode = KillNullInteger(rsProducts("product_sn_last_unit").Value + 1)
                'End If
                ProductIdentifier = KillNull(rsProducts(0)("product_sn_prod_code"))
                ReleaseCode = Chr(KillNullInteger(rsProducts(0)("product_release")) + 64)

                If LastNumericCode > 9999 Then
                    LogError(New Exception("Maximum Serial Number reached!"))
                    Return ReturnCodes.UDBS_ERROR
                End If

                ' Update the Numeric Code
                rsProducts(0)("product_sn_last_unit") = localSNLastUnit
                Dim columnNames As String() = {"product_sn_last_unit", "product_id"}
                Dim keys As String() = {"product_id"}
                Dim columnValues As Object() = {localSNLastUnit, rsProducts(0)("product_id")}
                UpdateNetworkRecord(keys, columnNames, columnValues, "product")

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        Private Function GetProductID(ProductNumber As String,
                                      ProductRelease As Double,
                                      ByRef ProductID As Integer) _
            As ReturnCodes
            ' Function returns the product id of the specified product/release
            Dim sqlQuery As String
            Dim rsTemp As New DataTable

            Try
                sqlQuery = "SELECT product_id FROM " & mProductTable &
                       " with(nolock) WHERE product_number = '" & ProductNumber & "' " &
                       "AND product_release = " & CStr(ProductRelease)
                If QueryNetworkDB(sqlQuery, rsTemp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR

                End If

                If (rsTemp.Rows.Count = 0) Then
                    ' Don't log this with the 'error' severity; clients may call this to validated whether or not
                    ' a product exists, and should be handling the return code accordingly.
                    logger.Debug($"No such product. Product Number '{ProductNumber}', Release {ProductRelease}.")

                    Return ReturnCodes.UDBS_ERROR
                ElseIf (rsTemp.Rows.Count > 1) Then
                    ' Data integrity problem.
                    logger.Warn($"Ambiguous Product Number '{ProductNumber}', Release {ProductRelease}.")
                End If

                ProductID = KillNullInteger(rsTemp(0)("product_id"))
                rsTemp = Nothing
                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

        Private Function GetFamilyID(ProductFamily As String,
                                     ByRef FamilyID As Integer) _
            As ReturnCodes
            ' Function returns the product id of the specified product/release

            Dim rsTemp As DataTable = Nothing
            Try
                If GetFamilyInfo(ProductFamily, rsTemp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                End If

                FamilyID = KillNullInteger(rsTemp(0)("family_id"))
                rsTemp = Nothing
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try

        End Function


        ''' <summary>
        ''' Function returns information about the specified family
        ''' </summary>
        Private Function GetFamilyInfo(ProductFamily As String,
                                       ByRef FamilyInfo As DataTable) As ReturnCodes
            Try
                Dim sqlQuery As String =
                    $"SELECT * FROM {mProductFamilyTable} with(nolock) WHERE family_name = '{ProductFamily}'"
                Return QueryNetworkDB(sqlQuery, FamilyInfo)
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function


        Private Function AddNewProductFamily(FamilyName As String,
                                             FamilyDescription As String,
                                             EmployeeNumber As String) _
            As ReturnCodes
            ' Function add new product family to UDBS
            ' Get server date/time
            Dim ServerTime As Date

            Try
                CUtility.Utility_GetServerTime(ServerTime)

                ' Add the new product family
                Dim sqlQuery As String
                Dim rsTemp As New DataTable

                sqlQuery = "SELECT * FROM " & mProductFamilyTable & " with(nolock) WHERE family_name = '" & FamilyName &
                           "' "
                OpenNetworkRecordSet(rsTemp, sqlQuery)
                If (If(rsTemp?.Rows?.Count, 0)) > 0 Then
                    ' Family exists...
                    LogError(New Exception("Product Family already exists."))
                    Return ReturnCodes.UDBS_ERROR

                End If

                Dim columnNames = New String() _
                        {"family_name", "family_created_by", "family_created_date", "family_description"}
                Dim columnValues = New Object() _
                        {FamilyName.Trim(), EmployeeNumber?.Trim().ToUpperInvariant(), ServerTime, FamilyDescription?.Trim()}
                Dim result = InsertNetworkRecord(columnNames, columnValues, mProductFamilyTable)

                Return If(result > 0, ReturnCodes.UDBS_OP_SUCCESS, ReturnCodes.UDBS_ERROR)

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try

        End Function

#Region "New functions that did not exist in VB6 version"

        ''' <summary>
        '''     Queries UDBS for the list of PN for the given ID. Returns the last PN in the list,
        '''     or empty string if not found.
        ''' </summary>
        ''' <param name="UnitID">String. Udbs Product ID</param>
        ''' <returns>String. Default part number</returns>
        ''' <remarks>SD moved from JdsuUdbsLibrary 2014-10-09 Logs and suppresses exceptions. Returns empty string if not found</remarks>
        Public Shared Function LookupDefaultPartNo(unitID As String) As String

            Dim rsTmp As DataTable = Nothing

            Try

                ' Build query
                Dim sSQL As String = "SELECT pg_string_value FROM udbs_product_group with(nolock) " &
                                     "WHERE pg_product_group LIKE '" & unitID & "_variance' " &
                                     "AND pg_integer_value=0" &
                                     "ORDER BY pg_sequence"

                rsTmp = New DataTable
                OpenNetworkRecordSet(rsTmp, sSQL)
                If (If(rsTmp?.Rows?.Count, 0)) > 0 Then
                    Dim dr = rsTmp.AsEnumerable().Last()
                    Dim sTmp() As String = Split((rsTmp(0)("pg_string_value")).ToString, ",")
                    Return sTmp(LBound(sTmp))
                Else
                    Return ""
                End If

            Catch ex As Exception
                If UDBSDebugMode Then
                    logger.Debug(ex, "Could not find default part number for given ID {0}", unitID)
                End If

                Return ""

            Finally
                rsTmp = Nothing
            End Try
        End Function

        ' The CASE statement parses the "Product Group String Value" and extracts
        ' the first element (before the first comma), that is the Oracle Part Number.
        Private Const SEARCH_FOR_ORACLE_PN = "SELECT
            CASE
                WHEN pg_string_value IS NULL THEN NULL
	            WHEN PATINDEX('%,%', pg_string_value)= 0 then pg_string_value
                ELSE SUBSTRING(pg_string_value, 1, PATINDEX('%,%', pg_string_value) - 1)
            END AS OraclePartNumber
            FROM udbs_product_group with(nolock) 
            WHERE pg_product_group LIKE '{0}_variance'
            ORDER BY pg_sequence"

        ''' <summary>
        ''' Look-up all Oracle part numbers from the UDBS Product ID.
        ''' </summary>
        ''' <param name="udbsProductId">The UDBS product ID.</param>
        ''' <returns>The list of matching Oracle part numbers.</returns>
        Public Function LookupAllPartNumbers(udbsProductId As String) As List(Of String)
            Dim resultSet As DataTable = Nothing
            Try
                OpenNetworkRecordSet(resultSet, String.Format(SEARCH_FOR_ORACLE_PN, udbsProductId))

                Dim partNumbers = New List(Of String)

                For Each row As DataRow In resultSet.Rows
                    Dim partNumber As String = row("OraclePartNumber").ToString()
                    partNumbers.Add(partNumber)
                Next

                Return partNumbers
            Finally
                resultSet?.Dispose()
            End Try
        End Function

#End Region

        '**********************************************************************
        '* Unit Support Functions
        '**********************************************************************


        ' Function fills recordset argument with unit and product information
        Private Function GetUnitInfo(ProductNumber As String,
                                     SerialNumber As String,
                                     ByRef UnitInfo As DataTable) As ReturnCodes
            Try
                Dim sqlQuery As String =
                           "SELECT * FROM " & mProductTable & " with(nolock), " & mUnitTable & " with(nolock) " &
                           "WHERE product_id = unit_product_id " &
                           "AND product_number='" & ProductNumber & "' " &
                           "AND unit_serial_number = '" & SerialNumber & "'"
                Return QueryNetworkDB(sqlQuery, UnitInfo)
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function


        Private Function GetUnitId(ProductNumber As String,
                                   SerialNumber As String,
                                   ByRef UnitID As Integer) _
            As ReturnCodes
            ' Function returns the unit id of the specified product/unit
            Dim rsTemp As DataTable = Nothing
            Try
                If GetUnitInfo(ProductNumber, SerialNumber, rsTemp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Return ReturnCodes.UDBS_ERROR
                    Exit Function
                End If
                UnitID = KillNullInteger(rsTemp(0)("unit_id"))

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            Finally
                rsTemp?.Dispose()
            End Try

        End Function


        Private Function GetLastSN(ByRef LastSNColumn As String,
                                   BuildLocation As String,
                                   incrementSN As Boolean) _
            As Integer
            ' Function returns the last number of unit built at the specific buildlocation
            ' if incrementSN is TRUE, GetLastSN=last number+1
            Dim LocationPos As Integer
            Dim SemiColonPos As Integer
            Dim SN As Integer

            Try
                BuildLocation = UCase(BuildLocation)
                If BuildLocation = "" Then BuildLocation = "J" 'default it to Ottawa=J
                LocationPos = InStr(1, LastSNColumn, BuildLocation)

                If LocationPos <> 0 Then
                    ' this is coded for multi-site
                    SemiColonPos = InStr(LocationPos, LastSNColumn, ";")
                    SN = CInt(Mid(LastSNColumn, LocationPos + 2, (SemiColonPos - (LocationPos + 2))))
                    If incrementSN Then
                        SN = SN + 1
                    End If
                    LastSNColumn = Left(LastSNColumn, LocationPos) & ":" & SN & ";" &
                                   Right(LastSNColumn, Len(LastSNColumn) - SemiColonPos)
                Else
                    ' this is not coded for multi-site
                    SN = CInt(Val(LastSNColumn))
                    If incrementSN Then
                        SN = SN + 1
                    End If
                    LastSNColumn = SN.ToString
                End If

                Return SN

                Exit Function

            Catch ex As Exception
                LogErrorInDatabase(ex)
                If incrementSN Then
                    LastSNColumn = "1"
                    Return 1
                Else
                    LastSNColumn = "0"
                    Return 0
                End If
            End Try

        End Function

        ''' <summary>
        ''' Add to Unit Table and also doing the Variant.
        ''' Adds a new serial number to the unit table and relevant record in the udbs_unit_details table.
        ''' </summary>
        ''' <param name="SerialNumber"></param>
        ''' <param name="OPN"></param>
        ''' <param name="EmployeeNumber"></param>
        ''' <returns></returns>
        Public Function AddSNwVar(SerialNumber As String,
                                  OPN As String,
                                  EmployeeNumber As String) As ReturnCodes
            Dim result As ReturnCodes = ReturnCodes.UDBS_ERROR

            If (String.IsNullOrWhiteSpace(SerialNumber)) Then
                ' Calling function MUST pass in a Serial Number
                LogError(New Exception("Missing serial number."))
                Return ReturnCodes.UDBS_ERROR
            End If

            Try
                SerialNumber = SerialNumber.Trim()
                ' Get server date/time
                Dim ServerTime As Date
                CUtility.Utility_GetServerTime(ServerTime)
                'Stop
                ' verify from here
                Dim rsTmp As New DataTable, sSQL As String
                Dim localProductID As Integer, UnitID As Integer, isVar As Boolean, pgSeq As Integer
                Dim arrStr() As String, tmpStr As String = String.Empty

                ' Get Product ID (latest release)
                sSQL = "SELECT product_id FROM product " &
                       "with(nolock) WHERE product_number='" & mNumber & "' " &
                       "ORDER BY product_release DESC"
                If QueryNetworkDB(sSQL, rsTmp) <> ReturnCodes.UDBS_OP_SUCCESS _
                        OrElse rsTmp.Rows.Count = 0 Then
                    logger.Error($"No such product number: {mNumber}")
                    Return ReturnCodes.UDBS_ERROR
                End If

                localProductID = KillNullInteger(rsTmp(0)("product_id"))
                rsTmp = Nothing

                ' Check for variance.
                sSQL = "SELECT * FROM udbs_product_group with(nolock)" &
                       "WHERE pg_product_group='" & mNumber & "_variance' " &
                       "ORDER BY pg_sequence;"
                If QueryNetworkDB(sSQL, rsTmp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    Throw New Exception($"Error querying for product variance: {mNumber}")
                End If

                If OPN = "" Then
                    Select Case (If(rsTmp?.Rows?.Count, 0))
                        Case 0
                            ' non-RoHS product
                            isVar = False
                        Case Else
                            LogError(New Exception("Missing Oracle PN information."))
                            rsTmp = Nothing
                            Return ReturnCodes.UDBS_ERROR
                    End Select
                Else
                    If (If(rsTmp?.Rows?.Count, 0)) = 0 Then
                        ' non-RoHS product
                        isVar = False
                        If UCase(OPN) <> UCase(mNumber) Then
                            LogError(New Exception("Inconsistent part_number and OPN."))
                            ' no big deal, keep going.
                        End If
                    Else
                        Dim ctr = 0
                        For Each dr As DataRow In rsTmp.Rows
                            tmpStr = KillNull(dr("pg_string_value"))
                            arrStr = Split(tmpStr, ",")
                            If UBound(arrStr) >= 2 Then
                                If UCase(OPN) = UCase(arrStr(0)) Then
                                    tmpStr = arrStr(0)
                                    Exit For
                                End If
                            End If
                            ctr += 1
                        Next
                        If UCase(OPN) = UCase(tmpStr) Then
                            pgSeq = KillNullInteger(rsTmp(ctr)("pg_sequence"))
                            'sOPN = arrStr(0)
                            isVar = True
                        Else
                            LogError(New Exception("Invalid Oracle PN."))
                            rsTmp = Nothing
                            Return ReturnCodes.UDBS_ERROR
                        End If
                    End If
                End If
                rsTmp = Nothing

                SyncLock mUnitCreationLock
                    sSQL = "SELECT * " &
                       "FROM unit with(nolock) " &
                       "WHERE unit_serial_number = '" & SerialNumber & "' " &
                       "AND unit_product_id=" & localProductID
                    If QueryNetworkDB(sSQL, rsTmp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                        Throw New Exception($"Error querying to verify if a unit with serial number ""{SerialNumber}"" exists already.")
                    End If

                    If (If(rsTmp?.Rows?.Count, 0)) > 0 Then
                        ' Unit already exists...
                        rsTmp = Nothing
                        logger.Error($"Unit with serial # {SerialNumber} and product id {localProductID} already exists!")
                        Return ReturnCodes.UDBS_RECORD_EXISTS
                    End If

                    'atomic insert on 2 tables .Both must be successful
                    Using transactionScope = BeginNetworkTransaction()
                        Try
                            Dim columnNames = New String() _
                            {
                                "unit_product_id",
                                "unit_serial_number",
                                "unit_created_by",
                                "unit_created_date",
                                "unit_report"
                            }
                            Dim columnValues = New Object() _
                            {
                                mProductID,
                                SerialNumber.ToUpperInvariant(),
                                EmployeeNumber?.Trim().ToLowerInvariant(),
                                ServerTime,
                                DetermineSoftwareName()
                            }

                            ' Insert with transaction scope
                            UnitID = InsertNetworkRecord(columnNames, columnValues, "unit", transactionScope, "unit_id")

                            rsTmp = Nothing

                            If isVar Then
                                sSQL = "SELECT * FROM udbs_unit_details with(nolock) " &
                                           "WHERE ud_unit_id=" & UnitID & " " &
                                           "AND ud_identifier='PRD_VAR' " &
                                           "AND ud_pg_product_group='" & mNumber & "_variance'"
                                If QueryNetworkDB(sSQL, rsTmp) <> ReturnCodes.UDBS_OP_SUCCESS Then
                                    Throw New Exception("Error querying for variance.")
                                End If
                                If (If(rsTmp?.Rows?.Count, 0)) <> 0 Then
                                    ' variance already exists, couldn't be true since the SN was just added!!!
                                    ' rollback transaction
                                    transactionScope.HasError = True
                                    result = ReturnCodes.UDBS_RECORD_EXISTS
                                    Throw _
                                            New ApplicationException(
                                                "Variance already exists, couldn't be true since the SN was just added!!!")
                                Else
                                    columnNames = New String() _
                                            {"ud_unit_id", "ud_pg_product_group", "ud_pg_sequence", "ud_identifier",
                                             "ud_string_value"}
                                    columnValues = New Object() _
                                            {UnitID, $"{mNumber}_variance", pgSeq, "PRD_VAR", $"OPN: {OPN}"}
                                    ' Insert with transaction scope
                                    InsertNetworkRecord(columnNames, columnValues, "udbs_unit_details", transactionScope)
                                End If
                            End If
                        Catch ex As Exception
                            transactionScope.HasError = True
                            Throw
                        End Try
                    End Using ' End of transaction scope, commited or rolled back
                End SyncLock

                Return ReturnCodes.UDBS_OP_SUCCESS

            Catch ex As Exception
                ' Keep the bad result if one was set.
                result = If(result = ReturnCodes.UDBS_OP_SUCCESS, ReturnCodes.UDBS_ERROR, result)
                LogErrorInDatabase(ex)
                Return result
            End Try
        End Function

        ''' <summary>
        ''' Logs the UDBS error in the application log and inserts it in the udbs_error table, with the process instance details:
        ''' Process Type, name, Process ID, UDBS product ID, Unit serial number.
        ''' </summary>
        ''' <param name="ex">Exception raised.</param>
        Private Sub LogErrorInDatabase(ex As Exception)

            DatabaseSupport.LogErrorInDatabase(ex, "Product", String.Empty, Me.ProductID, Me.Number, String.Empty)

        End Sub

#Region "IDisposable Support"

        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    CloseNetworkDB()
                    ' An empty product object can be used to create a new product
                    mREADONLY = False
                    mObjectLoaded = False
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region
    End Class
End Namespace
