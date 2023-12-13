''' <summary>
''' Enumerates the Data Integrity state of the local database.
''' </summary>
<Flags>
Public Enum LocalDBIntegrityStatus
    ''' <summary>
    ''' All Good. No Integrity Issues.
    ''' </summary>
    Good = 0

    ''' <summary>
    ''' Process Id Does not exist in the local database in the process_registration table.
    ''' </summary>
    ProcessIdDoesNotExist = 1

    ''' <summary>
    ''' Test Data Process does not exist in testdata_process table.
    ''' </summary>
    TestDataProcessDoesNotExist = 2

    ''' <summary>
    ''' Test Data value is missing in the testdata_result table.
    ''' </summary>
    TestDataResultValueMissing = 4

    ''' <summary>
    ''' Test Results table contains no entries.
    ''' </summary>
    TestDataResultNotPresent = 8

    ''' <summary>
    ''' Error occured while checking local database integrity.
    ''' </summary>
    ErrorCheckingIntegrity = 16
End Enum
