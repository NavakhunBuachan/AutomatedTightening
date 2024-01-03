Imports MesTestDataLibrary

Public Class testDataMES
    'Public Property specBuilder = MESTestDataFactory.Instance.GetMesTestDataBuilder(Of ISpecBuilder)()
    'Public Property mesBuilder = MESTestDataFactory.Instance.GetMesTestDataBuilder(Of IMesBuilder)()
    ' Public Property testBuilder = MESTestDataFactory.Instance.GetMesTestDataBuilder(Of ITestBuilder)()
    Implements testDataInterface
    Public Property testData As DataTable Implements testDataInterface.testData
    Public Property unitDetails As New Dictionary(Of String, String) From {{"", ""}}
    Public Property udbsPartID As String
    Public Property oraclePartID As String
    Public Property iMesBuild As IMes

    Public Property TestInst As UdbsInterface.TestDataInterface.CTestdata_Instance Implements testDataInterface.TestInst
    Public Property TestInstLast As UdbsInterface.TestDataInterface.CTestdata_Instance Implements testDataInterface.TestInstLast

    Friend logger As NLog.Logger = NLog.LogManager.GetLogger("Program")

    Public Property SerialNumber() As String Implements testDataInterface.SerialNumber
        Get
            Return iMesBuild.SerialNumber
        End Get
        Set(value As String)

        End Set
    End Property
    Public Property UDBSPartNumber() As String Implements testDataInterface.UDBSPartNumber
        Get
            Return udbsPartID
        End Get
        Set(value As String)

        End Set
    End Property
    Public Property wipstage() As String Implements testDataInterface.wipstage
        Get
            Return iMesBuild.WipStage
        End Get
        Set(value As String)

        End Set
    End Property
    Public Property passRouting() As String Implements testDataInterface.passRouting
        Get
            'MES not support 
            Return "Pass routing"
        End Get
        Set(value As String)

        End Set
    End Property
    Public Property failRouting() As String Implements testDataInterface.failRouting
        Get
            'MES not support 
            Return "Fail routing"
        End Get
        Set(value As String)

        End Set
    End Property

    Public Function initial(snStr As String, enStr As String, stationStr As String) As Boolean Implements testDataInterface.initial

        InitializeMesTestDataFactory()

        ' Dim temp As MesTestDataLibrary.MESTestDataFactory

        TestInst = New UdbsInterface.TestDataInterface.CTestdata_Instance
        TestInstLast = New UdbsInterface.TestDataInterface.CTestdata_Instance

        logger.Info("Initialize CWIP_Process and CTestdata_Instance")

        Dim errorString As String = ""
        Dim activeInWIP As Boolean

        Dim errMsg As String = String.Empty
        Dim configurationLoaded As Boolean = False

        Dim mesBuilder As IMesBuilder = MESTestDataFactory.Instance.GetMesTestDataBuilder(Of IMesBuilder)()

        Dim snFound As Boolean = WipTools.LookupSerialNo(
        snStr, oraclePartID, udbsPartID,
        errorString, activeInWIP, unitDetails)  ' it return true, even cannot find the unit

        If Not IsNothing(udbsPartID) Then
            mesBuilder.SetSerialNumber(snStr)
            mesBuilder.SetEmployeeId(enStr)
            mesBuilder.SetResourceName(stationStr)
            mesBuilder.SetUdbsProductNumber(udbsPartID)

            iMesBuild = mesBuilder.Build()
            unitDetails = iMesBuild.GetUnitDetails()

        Else
            logger.Error("Cannot find WIP process. Please check SN : " & snStr)
            '' dispose()
            'logger.Info("=>Unit Creating")
            'Dim DUT_UDBSID As String = "1-01763"
            'Dim DUT_SN As String = "SWTEST02"

            'Dim specBuilder = MESTestDataFactory.Instance.GetMesTestDataBuilder(Of ISpecBuilder)()
            'specBuilder.SetTestStage(iMesBuild.WipStage)
            'specBuilder.SetSpecRevision(0)
            'specBuilder.SetUdbsProductNumber(DUT_UDBSID)

            'Dim specInstance = specBuilder.Build()

            'If Not specInstance.LoadSpecs() Then
            '    Throw New Exception("Failed to load Specs.")
            'End If

            ''Dim mesBuilder2 = MESTestDataFactory.Instance.GetMesTestDataBuilder(Of IMesBuilder)()
            ''mesBuilder2.SetUdbsProductNumber(DUT_UDBSID).SetSerialNumber(SerialNumber)
            ''Dim mesInstance As IMes = mesBuilder2.Build()

            'Dim testBuilder = MESTestDataFactory.Instance.GetMesTestDataBuilder(Of ITestBuilder)()

            'testBuilder.SetUdbsProductNumber(DUT_UDBSID)
            'testBuilder.SetSerialNumber(DUT_SN)
            'testBuilder.SetTestStage(iMesBuild.WipStage)
            'testBuilder.SetResourceName(My.Computer.Name)
            'testBuilder.SetSpec(specInstance)
            ''testBuilder.SetMes(mesInstance)

            'testBuilder.DoCreateMissingUnit(True)
            'logger.Info("<=Unit Creating")

            Return False
        End If

        'Dim dt As DataTable = iMesBuild.GetUnitsAtWipStep("final_test")
        ' unitDetails = New Dictionary(Of String, String) From {{"", ""}}

        'Dim snFound As Boolean = WipTools.LookupSerialNo(
        'snStr, oraclePartnumber, partID,
        'errorString, activeInWIP, unitDetails)

        'TestInst.LoadExisting(iMesBuild.WipStage, partID, snStr, 0)

        logger.Info("Compeleted")
        Return True
    End Function


    Public Function initialMES(snStr As String, enStr As String, stationStr As String) As Boolean

        InitializeMesTestDataFactory()

        ' Dim temp As MesTestDataLibrary.MESTestDataFactory

        TestInst = New UdbsInterface.TestDataInterface.CTestdata_Instance
        TestInstLast = New UdbsInterface.TestDataInterface.CTestdata_Instance

        logger.Info("Initialize CWIP_Process and CTestdata_Instance")

        Dim errorString As String = ""
        Dim activeInWIP As Boolean

        Dim errMsg As String = String.Empty
        Dim configurationLoaded As Boolean = False

        Dim mesBuilder As IMesBuilder = MESTestDataFactory.Instance.GetMesTestDataBuilder(Of IMesBuilder)()

        Dim snFound As Boolean = WipTools.LookupSerialNo(
        snStr, oraclePartID, udbsPartID,
        errorString, activeInWIP, unitDetails)  ' it return true, even cannot find the unit

        If Not IsNothing(udbsPartID) Then
            mesBuilder.SetSerialNumber(snStr)
            mesBuilder.SetEmployeeId(enStr)
            mesBuilder.SetResourceName(stationStr)
            mesBuilder.SetUdbsProductNumber(udbsPartID)

            iMesBuild = mesBuilder.Build()
            unitDetails = iMesBuild.GetUnitDetails()

        Else
            logger.Error("Cannot find WIP process. Please check SN : " & snStr)
            ' dispose()
            Return False
        End If

        ' unitDetails = New Dictionary(Of String, String) From {{"", ""}}

        'Dim snFound As Boolean = WipTools.LookupSerialNo(
        'snStr, oraclePartnumber, partID,
        'errorString, activeInWIP, unitDetails)

        'TestInst.LoadExisting(iMesBuild.WipStage, partID, snStr, 0)

        logger.Info("Compeleted")
        Return True
    End Function

    Public Sub dispose() Implements testDataInterface.dispose
        logger.Info("MES and TestInst dispose")
        If Not IsNothing(TestInst) Then
            'WIPProcess.Dispose()
            TestInst.Dispose()
            If Not IsNothing(TestInstLast) Then
                TestInstLast.Dispose()
            End If
        End If

        'WIPProcess = Nothing
        TestInst = Nothing
        TestInstLast = Nothing
    End Sub
    Public Sub InitializeMesTestDataFactory()
        Dim errMsg As String = String.Empty
        Dim configurationLoaded As Boolean = False

        Try
            'configurationLoaded = MesTestDataConfigReader.Configure(errMsg)

            Dim mesConfig As IMESTestDataFactoryConfig = New MESTestDataFactoryConfig(
                "CamstarMes", "MesTestData.CamstarIMes+Builder", My.Settings.NavaMesServer)
            '"CamstarMes", "MesTestData.CamstarIMes+Builder", "http://thaappmciodev03.li.lumentuminc.net:128/CamInterfaceSvc.svc") ' for QA test
            '"CamstarMes", "MesTestData.CamstarIMes+Builder", "http://patappmesprd03h.li.lumentuminc.net:124/CamInterfaceSvc.svc")

            'Dim mesConfig As IMESTestDataFactoryConfig = New MESTestDataFactoryConfig(
            'MesTestDataConfigReader.Config.IMesAssemblyName,
            'MesTestDataConfigReader.Config.IMesBuilderClassName,
            'MesTestDataConfigReader.Config.IMesConnectionString)

            'Dim testConfig As IMESTestDataFactoryConfig = New TestDataFactoryConfig(
            '"UdbsInterface", "UdbsInterface.UDBSITest+Builder", "", "", True)
            'Dim testConfig As IMESTestDataFactoryConfig = New TestDataFactoryConfig(
            'MesTestDataConfigReader.Config.ITestAssemblyName,
            'MesTestDataConfigReader.Config.ITestBuilderClassName,
            'MesTestDataConfigReader.Config.ITestConnectionString,
            'MesTestDataConfigReader.Config.ITestLocalDBDriverName,
            'MesTestDataConfigReader.Config.ITestCreateMissingUnit)

            Dim specConfig As IMESTestDataFactoryConfig = New MESTestDataFactoryConfig(
                "UdbsInterface", "UdbsInterface.UDBSISpec+Builder", "")
            'Dim specConfig As IMESTestDataFactoryConfig = New MESTestDataFactoryConfig(
            'MesTestDataConfigReader.Config.ISpecAssemblyName,
            'MesTestDataConfigReader.Config.ISpecBuilderClassName,
            'MesTestDataConfigReader.Config.ISpecConnectionString)

            MESTestDataFactory.AddBuilderConfig(Of IMesBuilder)(mesConfig)
            'MESTestDataFactory.AddBuilderConfig(Of ITestBuilder)(testConfig)
            MESTestDataFactory.AddBuilderConfig(Of ISpecBuilder)(specConfig)

        Finally
            Dim configErr = Not String.IsNullOrEmpty(errMsg)
            'UpdateDatasystemsStatus(configurationLoaded, configErr)
        End Try
    End Sub
    Public Function calTestSeqByWipStage() As Integer Implements testDataInterface.calTestSeqByWipStage

        Dim count As Integer = 0

        TestInst = New UdbsInterface.TestDataInterface.CTestdata_Instance
        ' TestInst.LoadExisting("final_test", "1-01894", "CNE105CIE006094-B", count)
        TestInst.LoadExisting(iMesBuild.WipStage, udbsPartID, iMesBuild.SerialNumber, count)
        'TestInst.LoadExisting("final_test", udbsPartID, iMesBuild.SerialNumber, count)

        If IsNothing(TestInst.UnitSerialNumber) Then
            count = 1
        Else
            If TestInst.Status = "IN PROCESS" Then
                count = TestInst.Sequence
            Else
                count = TestInst.Sequence + 1
            End If

        End If

        Return count

    End Function

    Public Function loadTestItemNameStep(stage As String, testSeq As Integer, index As Integer) As DataTable Implements testDataInterface.loadTestItemNameStep

        Dim TestInstTemp = New UdbsInterface.TestDataInterface.CTestdata_Instance

        If index = 1 Then
            TestInstTemp = TestInst
        ElseIf index = 2 Then
            TestInstTemp = TestInstLast
        End If

        testData = New DataTable("TestData_itemName")

        testData.Columns.Add(New DataColumn("Serial", GetType(String)))
        testData.Columns.Add(New DataColumn("Test_Seq", GetType(Integer)))
        testData.Columns.Add(New DataColumn("MT_Seq", GetType(Integer)))
        testData.Columns.Add(New DataColumn("ItemName", GetType(String)))
        testData.Columns.Add(New DataColumn("Val", GetType(Single)))
        testData.Columns.Add(New DataColumn("PassFlag", GetType(Integer)))
        testData.Columns.Add(New DataColumn("Conn_Num", GetType(String)))
        testData.Columns.Add(New DataColumn("Conn_Type", GetType(String)))


        For Each itemkey In TestInstTemp.Results.Keys
            'Console.WriteLine(TestInst.Results.Item(itemkey).PassFlag)
            'Console.WriteLine(itemkey)
            Dim MtNum As String = "0"
            'If InStr(itemkey, "mt") > 0 Then
            '    MtNum = itemkey.Substring(3, 1)
            'End If
            If InStr(itemkey.ToUpper, stage.ToUpper) > 0 Then
                MtNum = itemkey.Substring(stage.Length + 1, 1)
            End If

            If stage = "LC" And itemkey.Length <= 5 Then
                MtNum = itemkey
            End If
            'testData.Rows.Add(TestInst.UnitSerialNumber, TestInst.Sequence, -1, itemkey, TestInst.Results.Item(itemkey).Value, TestInst.Results.Item(itemkey).PassFlag, MtNum, stage)
            testData.Rows.Add(TestInstTemp.UnitSerialNumber, testSeq, -1, itemkey, TestInstTemp.Results.Item(itemkey).Value, TestInstTemp.Results.Item(itemkey).PassFlag, MtNum, stage)
            'testData.Rows.Add(itemkey, "2")
        Next

        TestInstTemp.Dispose()
        TestInstTemp = Nothing

        Return testData

    End Function

    Public Function StartTest(EN As String, PN As String) As Boolean Implements testDataInterface.StartTest
        logger.Debug("==>MES Start Test")
        'Dim TestUtil As UdbsInterface.TestDataInterface.CTestData_Utility
        'Dim WIPUtil As UdbsInterface.WipInterface.CWIP_Utility
        'Dim WIPProcess As UdbsInterface.WipInterface.CWIP_Process
        Dim iProcessID As String

        Dim DUT_SN, DUT_PN, DUT_UDBSID As String

        DUT_SN = iMesBuild.SerialNumber
        DUT_UDBSID = iMesBuild.UDBSPartNumber
        DUT_PN = PN

        logger.Info("=>Check Unit Creating")

        Dim cprod As New UdbsInterface.MasterInterface.CProduct

        Dim temp As Boolean = False

        temp = cprod.GetProduct(DUT_UDBSID, 1) ' mObjectLoaded Must loaded 

        If cprod.UnitExists(DUT_SN) Then
            logger.Info("Don't need to create unit")
        Else
            logger.Info("=>Unit Creating")
            cprod.AddSNwVar(DUT_SN, iMesBuild.OraclePartNumber, EN)
            logger.Info("<=Unit Creating")
        End If
        logger.Info("<=Check Unit Creating")

        TestInst = New UdbsInterface.TestDataInterface.CTestdata_Instance

        If TestInst.LoadExisting(iMesBuild.WipStage, DUT_UDBSID, DUT_SN, 0) = 1 Then
            If TestInst.Status = "IN PROCESS" Then
                TestInst.Pause()
                logger.Info("test is in process")
                iProcessID = CStr(TestInst.ID)

                'txtLog.Text &= "Restarting Test Instance..." & vbNewLine
                TestInst = New UdbsInterface.TestDataInterface.CTestdata_Instance
                If TestInst.RestartUnit(DUT_UDBSID, DUT_SN, iMesBuild.WipStage) = 1 And (iProcessID = CStr(TestInst.ID)) Then
                    'TestSeq = TestInst.Sequence
                    logger.Info("Restart the unit")
                    ' if found exists testData, assume MES has start....
                    Return True
                    Exit Function

                End If
            End If
        Else
            'logger.Info("=>Unit Creating")

            'Dim specBuilder = MESTestDataFactory.Instance.GetMesTestDataBuilder(Of ISpecBuilder)()
            'specBuilder.SetTestStage(iMesBuild.WipStage)
            'specBuilder.SetSpecRevision(0)
            'specBuilder.SetUdbsProductNumber(DUT_UDBSID)

            'Dim specInstance = specBuilder.Build()

            'If Not specInstance.LoadSpecs() Then
            '    Throw New Exception("Failed to load Specs.")
            'End If

            'Dim mesBuilder = MESTestDataFactory.Instance.GetMesTestDataBuilder(Of IMesBuilder)()
            'mesBuilder.SetUdbsProductNumber(DUT_UDBSID).SetSerialNumber(SerialNumber)
            'Dim mesInstance As IMes = mesBuilder.Build()

            'Dim testBuilder = MESTestDataFactory.Instance.GetMesTestDataBuilder(Of ITestBuilder)()

            'testBuilder.SetUdbsProductNumber(DUT_UDBSID)
            'testBuilder.SetSerialNumber(DUT_SN)
            'testBuilder.SetTestStage(iMesBuild.WipStage)
            'testBuilder.SetResourceName(My.Computer.Name)
            'testBuilder.SetSpec(specInstance)
            ''testBuilder.SetMes(mesInstance)

            'testBuilder.DoCreateMissingUnit(True)
            'testBuilder.Build()
            'logger.Info("<=Unit Creating")
        End If


        Try
            logger.Info("=>MES Start WIP")
            If Not iMesBuild.StartWip() Then
                Return False
            End If
            logger.Info("<=MES Start WIP")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try

        'Start a new testdata instance
        Dim itemlist_rev As Integer = 0
        TestInst = New UdbsInterface.TestDataInterface.CTestdata_Instance

        If TestInst.Start(DUT_UDBSID, DUT_SN, iMesBuild.WipStage, itemlist_rev) < 0 Then
            MessageBox.Show("Error starting testdata instance", "UDBS Error", MessageBoxButtons.OK)
            logger.Info("Error starting testdata instance")
            'txtLog.Text &= "UDBS Error - Error starting testdata instance" & vbNewLine
            'TestInst = Nothing
            Return False
            Exit Function
        End If
        If TestInst.StoreProcessData("sw_version", My.Application.Info.Version.ToString) < 0 Then
            MessageBox.Show("Error sw_version Input", "UDBS Error", MessageBoxButtons.OK)
            logger.Info("Error sw_version Input")
            Return False
            Exit Function
        End If
        If TestInst.StoreProcessData("employee_number", EN) < 0 Then
            MessageBox.Show("Error employee_number Input", "UDBS Error", MessageBoxButtons.OK)
            logger.Info("Error employee_number Input")
            Return False
            Exit Function
        End If

        logger.Debug("<==MES Start Test")
        Return True

    End Function
    Public Sub FinishTest(wipNotes As String) Implements testDataInterface.FinishTest
        'Dim WIPProcess As UdbsInterface.WipInterface.CWIP_Process
        Dim rc As UdbsInterface.TestDataInterface.ResultCodes

        Dim DUT_SN As String = SerialNumber

        rc = TestInst.EvaluateDevice()
        If TestInst.Finish < 0 Then

            MessageBox.Show("UDBS Error", "UDBS Error", MessageBoxButtons.OK)
            'txtLog.Text &= "UBDS Error - Error completing testdata instance" & vbNewLine
        End If

        TestInst.Dispose()
        TestInst = Nothing

        Try
            If rc = UdbsInterface.TestDataInterface.ResultCodes.UDBS_SPECS_PASS Then
                iMesBuild.SetWipDisposition("PASS")
            Else
                iMesBuild.SetWipDisposition("FAIL")
            End If
            iMesBuild.EndWip()

            'WIPProcess.Finish_Step(rc, wipNotes)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally

            iMesBuild.Dispose()
            iMesBuild = Nothing
        End Try
        ' End If
    End Sub
End Class

