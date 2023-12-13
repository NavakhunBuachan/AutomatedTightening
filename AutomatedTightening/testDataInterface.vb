Public Interface testDataInterface
    Property testData As DataTable
    Property TestInst As UdbsInterface.TestDataInterface.CTestdata_Instance
    Property TestInstLast As UdbsInterface.TestDataInterface.CTestdata_Instance
    Property SerialNumber As String
    Property UDBSPartNumber As String
    Property wipstage As String
    Property passRouting As String
    Property failRouting As String
    Function initial(snStr As String, enStr As String, stationStr As String) As Boolean
    Sub dispose()
    Function calTestSeqByWipStage() As Integer
    Function loadTestItemNameStep(stage As String, testSeq As Integer, index As Integer) As DataTable
    Sub FinishTest(wipNotes As String)

    Function StartTest(EN As String, PN As String) As Boolean


End Interface
