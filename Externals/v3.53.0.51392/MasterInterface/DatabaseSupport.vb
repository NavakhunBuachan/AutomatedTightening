Option Explicit On
Option Compare Text
Option Infer On
Option Strict On

Imports System.Data.Common
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports Microsoft.Win32

<Assembly: InternalsVisibleTo("FractalUnitTests, PublicKey=002400000480000094000000060200000024000052534131000400000100010059d5541166730691d4a7ef0f5f45c20bf8bc3a02a1a8ad83987063dc7942050134c01245c17347a9908a3124f1b4f20206dc7415f97e1b3a834ef21ee1247f4f736acca356e7f9e0d550dcdd2a404ea148a7577748ab5e05082032e1b614b25022e969179b4d9f6f7e23070e86f6838285ff77bc231f3cb02a2cb80db4461fa8")>
<Assembly: InternalsVisibleTo("UdbsInterfaceUnitTest, PublicKey=002400000480000094000000060200000024000052534131000400000100010059d5541166730691d4a7ef0f5f45c20bf8bc3a02a1a8ad83987063dc7942050134c01245c17347a9908a3124f1b4f20206dc7415f97e1b3a834ef21ee1247f4f736acca356e7f9e0d550dcdd2a404ea148a7577748ab5e05082032e1b614b25022e969179b4d9f6f7e23070e86f6838285ff77bc231f3cb02a2cb80db4461fa8")>
<Assembly: InternalsVisibleTo("CommonUnitTestingUtilities, PublicKey=002400000480000094000000060200000024000052534131000400000100010059d5541166730691d4a7ef0f5f45c20bf8bc3a02a1a8ad83987063dc7942050134c01245c17347a9908a3124f1b4f20206dc7415f97e1b3a834ef21ee1247f4f736acca356e7f9e0d550dcdd2a404ea148a7577748ab5e05082032e1b614b25022e969179b4d9f6f7e23070e86f6838285ff77bc231f3cb02a2cb80db4461fa8")>
<Assembly: InternalsVisibleTo("MesTestData.UDBS, PublicKey=002400000480000094000000060200000024000052534131000400000100010059d5541166730691d4a7ef0f5f45c20bf8bc3a02a1a8ad83987063dc7942050134c01245c17347a9908a3124f1b4f20206dc7415f97e1b3a834ef21ee1247f4f736acca356e7f9e0d550dcdd2a404ea148a7577748ab5e05082032e1b614b25022e969179b4d9f6f7e23070e86f6838285ff77bc231f3cb02a2cb80db4461fa8")>
Namespace MasterInterface
    ' Enumeration
    Friend Enum UDBS_DBType
        NetworkDB = 1
        LocalDB = 2
    End Enum

    ''' <summary>
    '''     Interface to Abstract Database operations
    ''' </summary>
    Friend Interface IDatabase
        Inherits IDisposable
        Property CommandTimeOut As Integer
        Property ConnectionString As String

        ''' <summary>
        ''' Gets the connection string SQL Client connection string format.
        ''' </summary>
        ''' <returns>Connection string</returns>
        ReadOnly Property SQLClientConnectionString As String

        ''' <summary>
        ''' Gets a value indicating whether returns True if the database system is available.
        ''' </summary>
        ReadOnly Property SystemAvailable As Boolean

        Function ConnectionStringWithoutPassword() As String

        ''' <summary>
        '''     Execute and return number of affected rows
        ''' </summary>
        ''' <param name="sqlCommand"></param>
        ''' <returns></returns>
        Function Execute(sqlCommand As String) As Integer

        ''' <summary>
        '''     Execute and return number of affected rows
        ''' </summary>
        ''' <param name="sqlCommand"></param>
        ''' <param name="transactionScope"></param>
        ''' <returns></returns>
        Function Execute(sqlCommand As String, transactionScope As ITransactionScope) As Integer

        ''' <summary>
        '''     Returns First column, first row
        ''' </summary>
        ''' <param name="sqlCommand"></param>
        ''' <returns></returns>
        Function ExecuteScalar(sqlCommand As String) As Object

        ''' <summary>
        '''     Returns a DataTable
        ''' </summary>
        ''' <param name="sqlCommand"></param>
        ''' <returns></returns>
        Function ExecuteData(sqlCommand As String) As DataTable

        ''' <summary>
        ''' Returns a DataRow
        ''' </summary>
        ''' <param name="sqlCommand"></param>
        ''' <param name="blob"> in lock-step with the rows of the DataTable</param>
        ''' <returns></returns>
        Function ExecuteData(sqlCommand As String, ByRef blob As Stream) As DataRow

        ReadOnly Property DBType As UDBS_DBType

        ''' <summary>
        '''     An internal throw-away data structure for holding schema information
        '''     Column Name, Column Type, Column Size
        ''' </summary>
        ''' <param name="sqlCommand"></param>
        ''' <returns></returns>
        Function GetColumnTypes(sqlCommand As String) As List(Of Tuple(Of String, String, Integer))

        ''' <summary>
        '''     Returns a handle for enclosing an atomic (scoped ) series of queries
        '''     to enable "atomic" transactions. For example, if you insert records on multiple
        '''     tables, AND you want to make sure that ALL are committed, or NONE are committed
        '''     if something fails
        ''' </summary>
        ''' <returns></returns>
        Function BeginTransaction() As ITransactionScope

        ''' <summary>
        '''     Returns a DataTable
        ''' </summary>
        ''' <param name="sqlCommand"></param>
        ''' <param name="scope">transaction handle</param>
        ''' <returns></returns>
        Function ExecuteData(sqlCommand As String, scope As ITransactionScope) As DataTable

        ''' <summary>
        '''     Simple utility using T-SQL for inserting records given name-value pairs, and a table name.
        ''' </summary>
        ''' <param name="columnNames"></param>
        ''' <param name="columnValues"></param>
        ''' <param name="tableName"></param>
        ''' <returns></returns>
        Function InsertRecord(columnNames As String(), columnValues As Object(), tableName As String,
                              Optional ByVal primaryKey As String = "") As Integer

        ''' <summary>
        '''     Simple utility using T-SQL for inserting records given name-value pairs, and a table name.
        '''     Similar to <see cref="InsertRecord"/> but returns a Long to support BigInt data types in Network DB.
        ''' </summary>
        ''' <param name="columnNames"></param>
        ''' <param name="columnValues"></param>
        ''' <param name="tableName"></param>
        ''' <returns>Row ID as Long</returns>
        Function InsertRecord64Bits(columnNames As String(), columnValues As Object(), tableName As String,
                               Optional ByVal primaryKey As String = "") As Long

        ''' <summary>
        '''     Simple utility using T-SQL for inserting records given name-value pairs, and a table name
        '''     which has a scope
        ''' </summary>
        ''' <param name="columnNames"></param>
        ''' <param name="columnValues"></param>
        ''' <param name="tableName"></param>
        ''' <param name="scope"></param>
        ''' <returns></returns>
        Function InsertRecord(columnNames As String(), columnValues As Object(), tableName As String,
                              scope As ITransactionScope, Optional ByVal primaryKey As String = "") As Integer

        ''' <summary>
        '''     Simple utility using T-SQL for inserting records given name-value pairs, and a table name.
        '''     Similar to <see cref="InsertRecord"/> but returns a Row ID as Long instead of Integer.
        ''' </summary>
        ''' <param name="columnNames"></param>
        ''' <param name="columnValues"></param>
        ''' <param name="tableName"></param>
        ''' <param name="scope"></param>
        ''' <returns>Row ID as Long</returns>
        Function InsertRecord64Bits(columnNames As String(), columnValues As Object(), tableName As String,
                              scope As ITransactionScope, Optional ByVal primaryKey As String = "") As Long

        ''' <summary>
        '''     Simple utility using T-SQL for updating records given keys, name-value pairs, and a table name
        ''' </summary>
        ''' <param name="constraintKeys">Keys for the update</param>
        ''' <param name="columnNames"></param>
        ''' <param name="columnValues"></param>
        ''' <param name="tableName"></param>
        ''' <returns></returns>
        Function UpdateRecord(constraintKeys As String(), columnNames As String(), columnValues As Object(),
                              tableName As String, scope As ITransactionScope) As Boolean

        Function UpdateRecord(constraintKeys As String(), columnNames As String(), columnValues As Object(),
                              tableName As String) As Boolean

        Function DeleteRecord(constraintKeys As String(), constraintValues As Object(), tableName As String) As Boolean

        ''' <summary>
        '''     Create a working table to perform edits
        '''     NOTE: Works only for single-table use!
        ''' </summary>
        ''' <param name="selectSQL"></param>
        ''' <param name="scope"></param>
        ''' <param name="workTable">
        '''     The datatable initially filled with data returned from executing <paramref name="selectSQL" />
        ''' </param>
        ''' <returns></returns>
        Function CreateTableAdapter(selectSQL As String, scope As ITransactionScope, ByRef workTable As DataTable) _
            As IAdapterSession


        ''' <summary>
        ''' Bulk insert or update records given the target columns, the columns that MUST match, a source DataTable and the destination table
        ''' Insert happens when no match occurs
        ''' </summary>
        ''' <param name="targetColumnNames"></param>
        ''' <param name="matchOnColumnNames"></param>
        ''' <param name="sourceTable"></param>
        ''' <param name="tableName"></param>
        ''' <param name="scope"></param>
        <Obsolete("This method is not used anymore and might be taken out after June 2022.", False)>
        Sub BulkInsertOrUpdateRecords(targetColumnNames As String(),
                      matchOnColumnNames As String(),
                      sourceTable As DataTable,
                      tableName As String,
                      scope As ITransactionScope)

        ''' <summary>
        ''' Bulk-insert records given the target columns, a source DataTable and the destination table
        ''' </summary>
        ''' <param name="targetColumnNames"></param>
        ''' <param name="sourceTable"></param>
        ''' <param name="tableName"></param>
        ''' <param name="scope"></param>
        Sub BulkInsertRecords(targetColumnNames As String(), sourceTable As DataTable, tableName As String,
                        scope As ITransactionScope)


        ''' <summary>
        ''' Bulk-update records given the target columns, the columns that MUST match, a source DataTable and the destination table
        ''' </summary>
        ''' <param name="targetColumnNames"></param>
        ''' <param name="matchOnColumnNames"></param>
        ''' <param name="sourceTable"></param>
        ''' <param name="tableName"></param>
        ''' <param name="scope"></param>
        Sub BulkUpdateRecords(targetColumnNames As String(),
                              matchOnColumnNames As String(),
                              sourceTable As DataTable,
                              tableName As String,
                        scope As ITransactionScope)

    End Interface

    Friend Interface IAdapterSession
        Inherits IDisposable
        ReadOnly Property Adapter As DbDataAdapter
        ReadOnly Property Builder As DbCommandBuilder
    End Interface

    Friend Interface ITransactionScope
        Inherits IDisposable
        Property HasError As Boolean
    End Interface


    ''' <summary>
    ''' Releaser of Lock and a Transaction for a singleton DB connection object.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    Friend NotInheritable Class SingletonReleaser(Of T As IDbConnection)
        Implements ITransactionScope

        ' Shared properties and member variables.

        ''' <summary>
        ''' A semaphore every instances of this class shares.
        ''' </summary>
        ''' <remarks>
        ''' Because this is a template class, each 'specialized' class has a different instance of the semaphore.
        ''' i.e.: SingletonRelease(int).lockObject is different than SingletonRelease(float).
        ''' </remarks>
        Private Shared ReadOnly lockObject As New SemaphoreSlim(1, 1)

        ''' <summary>
        ''' The stack trace of the thread at the moment the semaphore was last
        ''' acquired.
        ''' </summary>
        ''' <remarks>
        ''' According to the documentation, retrieving the stack trace is slow,
        ''' so this is only enabled when <see cref="UDBSDebugMode"/> is enabled.
        ''' </remarks>
        Private Shared _lastSemaphoreAcquisition As StackTrace = Nothing

        ' Member variables.

        ''' <summary>
        ''' The transaction of this releaser.
        ''' </summary>
        ''' <returns></returns>
        Friend ReadOnly Property Transaction As IDbTransaction

        ''' <summary>
        ''' Whether or not the transaction encountered an error.
        ''' When an error is encountered and this instance gets disposed,
        ''' the transaction is rolled-back.
        ''' </summary>
        Private _hasError As Boolean

        ''' <summary>
        ''' Whether or not the semaphore was acquired, so we don't clear the stack
        ''' trace by mistake when we dispose this object.
        ''' </summary>
        Private _semaphoreAcquired As Boolean = False

        ' Properties

        Friend Property HasError As Boolean Implements ITransactionScope.HasError
            Get
                Thread.MemoryBarrier()
                Return (_hasError)
                Thread.MemoryBarrier()
            End Get
            Set
                Thread.MemoryBarrier()
                _hasError = Value
                Thread.MemoryBarrier()
            End Set
        End Property

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="connection">The DB connection.</param>
        ''' <param name="semaphoreAccessTimeout">
        ''' How long to wait for accessing the semaphore.
        ''' Defaults to <see cref="DefaultSemaphoreAccessTimeout"/>.
        ''' </param>
        Friend Sub New(connection As T, Optional semaphoreAccessTimeout As Integer = -1)
            ' only one thread at a time
            ' object can grab it on thread A
            If semaphoreAccessTimeout < 0 Then
                semaphoreAccessTimeout = DefaultSemaphoreAccessTimeout
            End If

            If Not lockObject.Wait(semaphoreAccessTimeout) Then
                ' The lock was not acquired in time.
                ' If there's no timeout, this could cause the entire application to freeze.
                ' When we failed to acquire the lock in time, we throw an exception that
                ' will include the stack trace of the thread that last successfully acquired
                ' the lock and is the most likely suspect for the deadlock.
                Dim message As String
                If _lastSemaphoreAcquisition IsNot Nothing Then
                    message = $"Failed to acquire lock in time. Last semaphore acquisition:{Environment.NewLine}{_lastSemaphoreAcquisition}"
                Else
                    message = "Failed to acquire lock in time. No information related to last semaphore acquisition."
                End If

                Throw New UDBSException(message)
            End If

            _semaphoreAcquired = True
            HasError = False
            Transaction = connection.BeginTransaction()

            If DatabaseSupport.UDBSDebugMode Then
                _lastSemaphoreAcquisition = New StackTrace()
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                If Transaction IsNot Nothing Then
                    If Not HasError Then
                        Try
                            Transaction.Commit()
                        Catch ex As Exception
                            logger.Error(ex, "Exception during commit. Rolling back.")
                            Try
                                ' some error, so undo everything
                                Transaction.Rollback()
                            Catch ex2 As Exception
                                logger.Error(ex2, "Exception on rollback.")
                                Throw New Exception("Exception on rollback.", ex2)
                            End Try

                            Throw
                        End Try
                    Else
                        Try
                            ' some error, so undo everything
                            Transaction.Rollback()
                        Catch ex2 As Exception
                            logger.Error(ex2, "Exception on rollback.")
                            Throw New Exception("Exception on rollback.", ex2)
                        End Try
                    End If
                End If
            Finally
                If _semaphoreAcquired Then
                    _lastSemaphoreAcquisition = Nothing
                End If

                ' and possible to release it on thread B
                lockObject.Release()
            End Try
        End Sub
    End Class

    ''' <summary>
    '''     A throw-away structure that is exposed as IDisposable
    '''     to take advantage of using {} compiler feature
    ''' </summary>
    Friend Structure Releaser(Of T As IDbTransaction)
        Implements ITransactionScope

        Friend ReadOnly Property Transaction As T
        Private _hasError As Boolean

        Friend Property HasError As Boolean Implements ITransactionScope.HasError
            Get
                Thread.MemoryBarrier()
                Return (_hasError)
                Thread.MemoryBarrier()
            End Get
            Set
                Thread.MemoryBarrier()
                _hasError = Value
                Thread.MemoryBarrier()
            End Set
        End Property

        Friend Sub New(toRelease As T)
            Transaction = toRelease
            HasError = False
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If Transaction IsNot Nothing Then
                If Not HasError Then
                    Try
                        Transaction.Commit()
                    Catch ex As Exception
                        logger.Error(ex, "Exception during commit. Rolling back.")
                        Try
                            ' some error, so undo everything
                            Transaction.Rollback()
                        Catch ex2 As Exception
                            logger.Error(ex2, "Exception on rollback.")
                            Throw New Exception("Exception on rollback.", ex2)
                        End Try

                        Throw
                    End Try
                Else
                    Try
                        ' some error, so undo everything
                        Transaction.Rollback()
                    Catch ex2 As Exception
                        logger.Error(ex2, "Exception on rollback.")
                        Throw New Exception("Exception on rollback.", ex2)
                    End Try
                End If
            End If
        End Sub
    End Structure

    Public Module DatabaseSupport
        Friend Const DBDateFormatting As String = "dd MMM yyyy HH:mm:ss.fff"
        Private buildFolder As String = GetApplicationParentDirectory()
        Private stationName As String
        Private PCname As String = Environment.MachineName
        ' UDBS System Support Variables
        Private mSystemInitialized As Boolean = False

        ' synchronizing object
        Private ReadOnly syncObj As New Object

        ' Connection Objects
        Private dbLocal As IDatabase
        Private dbNetwork As IDatabase

        ' Count of objects requiring database support

        Public UDBSDebugMode As Boolean = False

        Friend Const DBCommandTimeout As Integer = 120
        Friend DefaultSemaphoreAccessTimeout As Integer = 120000

        Friend ReadOnly logger As Logger = LogManager.GetLogger("UDBS")
        Friend ReadOnly enUS As CultureInfo = CultureInfo.CreateSpecificCulture("en-US")

        Private _localDbPath As String = DefaultLocalDbPath

        ''' <summary>
        ''' During many operations, we check whether or not the set of tables
        ''' related to a given process exists in the local database.
        ''' This takes roughly 50 msec, but we do it over and over, so this
        ''' accounts for a measurable delay when starting a test instance.
        ''' Once the tables are created, we don't need to check for them
        ''' ever again.
        ''' This will contain the list of processes for which we have already
        ''' checked/created the set of tables in the local DB.
        ''' When the local DB is cleared or deleted, this list is reset.
        ''' </summary>
        Private _localProcessTablesCreated As New List(Of String)()

        ''' <summary>
        ''' Member variable to hold the pre-network query hook.
        ''' </summary>
        ''' <see cref="SetupPreNetworkQueryHook(Func(Of String, Tuple(Of String, DataTable)))"/>
        Private _preNetworkQueryHook As Func(Of String, Tuple(Of String, DataTable)) = Nothing

        ''' <summary>
        ''' The system error queue.
        ''' </summary>
        Private _systemErrorQueue As New ErrorQueue

        ''' <summary>
        ''' Whether or not the assembly version has already been logged.
        ''' We only want to log this once per application execution.
        ''' </summary>
        Private _assemblyVersionHasBeenLogged As Boolean = False

        Private Enum LocalDriverEnum
            SQLite
            MsAccess
        End Enum

        Private Property LocalDBDriver As LocalDriverEnum = LocalDriverEnum.SQLite

        ''' <summary>
        ''' Stores the network database connection string without the password.
        ''' </summary>
        ''' <returns>ConnectionStringWithoutPassword() value from UDBSNetworkDatabase.vb</returns>
        Friend ReadOnly Property ConnectionStrWithoutPwdNetworkDB As String
            Get
                Return dbNetwork.ConnectionStringWithoutPassword
            End Get
        End Property

        ''' <summary>
        ''' Error queue.
        ''' The logic for queuing, saving to a local CSV file, etc. is all handled by that class.
        ''' </summary>
        Friend ReadOnly Property SystemErrorQueue As ErrorQueue
            Get
                Return _systemErrorQueue
            End Get
        End Property

        ''' <summary>
        ''' Unit test friendly method, doesn't rely on Registry or UDBS installation.
        ''' This is also useful for the TED Tools, which allow the user to select
        ''' what UDBS instance s/he want to interact with.
        ''' </summary>
        ''' <param name="odbcConnectionString">The database connection string, in Microsoft ODBC format.</param>
        Public Sub InitializeNetworkDB(odbcConnectionString As String)
            Try
                If CUtility.Utility_GetStationName(stationName) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    ' This is not critical.
                End If

                LogManager.EnableLogging()
                If UDBSDebugMode Then
                    logger.Debug("Initializing UDBS Network Database")
                End If

                ' Keep a temporary reference to the old value (if any) so that a warning can be generated
                ' if the connection string is changed.
                Dim prevDbNetwork = dbNetwork

                Dim isChangingConnectionString As Boolean = False
                Dim isTheSameUdbsInstance As Boolean = False

                If odbcConnectionString.Contains("(localdb)") Then
                    isChangingConnectionString = prevDbNetwork IsNot Nothing
                    isTheSameUdbsInstance = False
                Else
                    ' compare sql connection string to see if we are trying to connect to the same database.
                    ' if we are connecting to the same database the network database object must not be re-initialized.
                    Dim newSQLClientConnectionString = UDBSNetworkDatabase.ODBCToSQLClient(odbcConnectionString).ToString()
                    isChangingConnectionString = prevDbNetwork IsNot Nothing AndAlso prevDbNetwork.SQLClientConnectionString <> newSQLClientConnectionString
                    isTheSameUdbsInstance = AreTheSameUdbsInstance(prevDbNetwork?.SQLClientConnectionString, newSQLClientConnectionString)
                End If

                ' Stop the error queue if we are connecting to a different UDBS instance.
                If (SystemErrorQueue.IsRunning And Not isTheSameUdbsInstance) Then
                    SystemErrorQueue.RequestStop()
                    SystemErrorQueue.Join()
                    SystemErrorQueue.Clear() ' It will be reloaded in a minute.
                End If

                If (dbNetwork Is Nothing Or isChangingConnectionString) Then
                    If (dbNetwork IsNot Nothing) Then
                        dbNetwork.Dispose()
                    End If

                    dbNetwork = New UDBSNetworkDatabase(UDBS_DBType.NetworkDB) With {
                        .ConnectionString = odbcConnectionString
                    }
                End If

                ' Generate a warning if the database connection string is being modified.  The log statement
                ' is after creating the new instance so that the ConnectionStringWithoutPassword can be used
                ' to avoid logging the password.
                If isChangingConnectionString Then
                    logger.Warn($"Changing database connection string from '{prevDbNetwork.ConnectionStringWithoutPassword()}' to '{ConnectionStrWithoutPwdNetworkDB}'")
                End If

                If Not isTheSameUdbsInstance Then
                    ' We are changing UDBS instance.
                    ' Log its usage.
                    _assemblyVersionHasBeenLogged = False
                End If

                ' check that we can connect
                If Not dbNetwork.SystemAvailable Then
                    Throw New ApplicationException($"Not able to connect to {ConnectionStrWithoutPwdNetworkDB}")
                End If

                If (Not SystemErrorQueue.IsRunning) Then
                    ' Any entries in the queue are also present in the CSV file. Flush the queue prior to loading the CSV file.
                    SystemErrorQueue.Clear()
                    SystemErrorQueue.LoadFromFile()
                    SystemErrorQueue.Start()
                End If

                ' Log even if the system is not available.
                ' This will be added to the error queue and pushed to the
                ' DB once the connection is re-established.
                ' This has to happen AFTER starting the error queue, otherwise
                ' the entry in the is logged twice as it will be present
                ' in the CSV file and the queue.
                LogAssemblyVersion()

            Catch ex As Exception
                dbNetwork = Nothing
                Throw New Exception("Couldn't initialize network database.", ex)
            End Try
        End Sub

        ''' <summary>
        ''' Loads all interface settings from the registry, called once at startup.  Replaces the existing
        ''' value if the system has already been initialized.
        ''' </summary>
        Public Sub InitializeNetworkDB()
            ' serialize calls so that internal objects can be initialized without contention.
            SyncLock syncObj
                Dim odbcFormatConnectionString As String = String.Empty
                Try
                    odbcFormatConnectionString = GetSetting("UDBS_V3", "Database", "Network Connection",
                                                            "Driver={SQL Server};Server=poamserv;database=UDBS_Modules;uid=udbs;pwd=masterinterface;")
                Catch ex As Exception
                    Throw New Exception("Failed to read UDBS Network Connection string from system registry. " & ex.Message)
                End Try
                InitializeNetworkDB(odbcFormatConnectionString)
            End SyncLock
        End Sub

        Friend Sub InitializeLocalDB()
            Dim connectionString =
                    "Data Source=%CommonAppDataFolder%;Version=3;UseUTF16Encoding=True;datetimeformat=CurrentCulture"

            Try
                If UDBSDebugMode Then
                    logger.Debug("Initializing UDBS Local Database")
                End If
                dbLocal = New SQLiteDatabase()
                dbLocal.ConnectionString = connectionString.Replace("%CommonAppDataFolder%", LocalDBPath)
            Catch ex As Exception
                logger.Error(ex, "Failure to initialize local database.")
                dbLocal = Nothing
                Throw New Exception("Couldn't initialize local database. ", ex)
            End Try

            Try
                ExecuteLocalQuery("vacuum")
            Catch ex As Exception
                ' Log a warning, but do not throw an exception.
                ' This is not a critical operation.
                logger.Warn(ex, "Failed to compact local DB.")
            End Try
        End Sub

        ''' <summary>
        ''' Creates a copy of the local database
        ''' </summary>
        ''' <param name="folderPath"></param>
        ''' <param name="filePrefix"></param>
        ''' <param name="delete"></param>
        Friend Sub BackupLocalDB(folderPath As String, filePrefix As String, delete As Boolean)
            If Not dbLocal Is Nothing Then
                dbLocal.Dispose()
                dbLocal = Nothing
            End If

            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            End If

            File.Copy(LocalDBPath, Path.Combine(folderPath, filePrefix + Path.GetFileName(LocalDBPath)), True)
            If delete Then
                DeleteLocalDB()
            End If
        End Sub

        ''' <summary>
        ''' The default local SQLite DB path: C:\ProgramData\JDSU\UDBS\local_process.db
        ''' </summary>
        ''' <returns>The path to the local DB file.</returns>
        Friend ReadOnly Property DefaultLocalDbPath As String
            Get
                Return Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                            "JDSU", "UDBS", "local_process.db")
            End Get
        End Property

        ''' <summary>
        ''' Gets and sets the file path of the local database.
        ''' </summary>
        ''' <returns>Returns the file path of the local database.</returns>
        ''' <remarks>
        ''' Setting the path causes the local DB connection to be closed and a new local DB to be created.
        ''' Applications usually do not need to set this.
        ''' This is mainly exposed for testability purposes.
        ''' </remarks>
        Friend Property LocalDBPath As String
            Get
                Return _localDbPath
            End Get
            Set(value As String)
                If (value <> _localDbPath) Then
                    _localProcessTablesCreated.Clear()
                    _localDbPath = value
                    CloseLocalDB()
                    InitializeLocalDB()
                End If
            End Set
        End Property

        ''' <summary>
        ''' Closes the local DB.
        ''' This 'really' closes the local DB.
        ''' There is usually no need to call this method.
        ''' This is mostly used by unit tests to manage multiple instances of the
        ''' local DB, to ensure each test starts with a fresh slate and that multiple
        ''' instance of the tests, running on the same computer, do not step on one
        ''' another's toes.
        ''' </summary>
        Friend Sub CloseLocalDB()
            If (dbLocal IsNot Nothing) Then
                dbLocal.Dispose()
                dbLocal = Nothing
            End If
        End Sub

        ''' <summary>
        ''' Delete the local DB.
        ''' Do not simply call File.Delete(...) using the local DB path; there is
        ''' house-keeping that needs to happen when the local DB is deleted.
        ''' </summary>
        Friend Sub DeleteLocalDB()
            File.Delete(LocalDBPath)
            _localProcessTablesCreated.Clear()
        End Sub

        ''' <summary>
        ''' Saves Registry Setting from the VB6 style registry hive
        ''' </summary>
        ''' <param name="applicationName"></param>
        ''' <param name="sectionName"></param>
        ''' <param name="key"></param>
        ''' <param name="payload"></param>
        Friend Sub SaveSetting(ByVal applicationName As String, ByVal sectionName As String, ByVal key As String, ByVal payload As String)
            OsAbstractionLayer.Instance.RegistryAdapter.SaveSetting(applicationName, sectionName, key, payload)
        End Sub

        ''' <summary>
        ''' Reads setting from the VB6 style Registry Hive
        ''' </summary>
        ''' <param name="applicationName"></param>
        ''' <param name="sectionName"></param>
        ''' <param name="key"></param>
        ''' <param name="defaultValue"></param>
        ''' <returns></returns>
        Friend Function GetSetting(ByVal applicationName As String, ByVal sectionName As String, ByVal key As String, ByVal defaultValue As String) As String
            Return OsAbstractionLayer.Instance.RegistryAdapter.GetSetting(applicationName, sectionName, key, defaultValue)
        End Function

        Public Sub SetNetworkConnectionString(ConnStr As String, Optional ByVal SaveToRegistry As Boolean = True)
            If dbNetwork Is Nothing Then InitializeNetworkDB()
            dbNetwork.ConnectionString = ConnStr
            If UDBSDebugMode Then
                logger.Debug(
                    "Setting UDBS Network Database Connection String to: " & ConnectionStrWithoutPwdNetworkDB)
            End If
            If SaveToRegistry Then SaveSetting("UDBS_V3", "Database", "Network Connection", ConnStr)
        End Sub

        Friend Function GetNetworkConnectionString() As String
            If dbNetwork Is Nothing Then InitializeNetworkDB()
            Return dbNetwork.ConnectionString
        End Function

        ''' <summary>
        ''' Create a SQL connection to the network DB.
        ''' We recommand not to make direct SQL queries into UDBS and use specialized
        ''' classes instead.
        ''' This method remains there for backward compatibility for tools that perform
        ''' operations not directly exposed through the UDBS or MES Test Data interfaces.
        ''' </summary>
        ''' <returns>
        ''' A newly created SQL connection. Do not forget to 'Dispose' of it.
        ''' </returns>
        Public Function CreateNetworkSqlConnection() As SqlConnection
            If dbNetwork Is Nothing Then InitializeNetworkDB()

            Return New SqlConnection(dbNetwork.ConnectionString)
        End Function

        ''' <summary>
        ''' This "Pre Network Query" hook lets one inspect or manipulate the SQL
        ''' query prior to execution.
        ''' This also lets one throw an exception to simulate a communication
        ''' problem with the database.
        ''' </summary>
        ''' <param name="hook">
        ''' The hook to invoke when the library is about to make a network DB query.
        ''' Setting this to 'Nothing' (null) will remove a previously installed
        ''' hook.
        ''' </param>
        ''' <remarks>
        ''' There can only be one hook setup at a time.
        ''' Setting-up a hook removes the previouly setup one.
        ''' </remarks>
        Friend Sub SetupPreNetworkQueryHook(hook As Func(Of String, Tuple(Of String, DataTable)))
            _preNetworkQueryHook = hook
        End Sub

        Friend Sub ExecuteNetworkQuery(SSQL As String, scope As ITransactionScope)
            If dbNetwork Is Nothing Then InitializeNetworkDB()

            Try
                If (_preNetworkQueryHook IsNot Nothing) Then
                    Dim result = _preNetworkQueryHook.Invoke(SSQL)
                    If (result IsNot Nothing AndAlso Not String.IsNullOrEmpty(result.Item1)) Then
                        SSQL = result.Item1
                    End If
                End If

                If UDBSDebugMode Then
                    logger.Debug("Executing network query: " & SSQL)
                End If
                dbNetwork.Execute(SSQL, scope)
            Catch ex As Exception
                scope.HasError = True
                Throw New UDBSException($"Failed to execute network query: '{SSQL}'. {ex.Message}", ex)
            End Try
        End Sub

        Friend Sub ExecuteNetworkQuery(SSQL As String)
            If dbNetwork Is Nothing Then InitializeNetworkDB()

            Try
                If (_preNetworkQueryHook IsNot Nothing) Then
                    Dim result = _preNetworkQueryHook.Invoke(SSQL)
                    If (result IsNot Nothing AndAlso Not String.IsNullOrEmpty(result.Item1)) Then
                        SSQL = result.Item1
                    End If
                End If

                dbNetwork.Execute(SSQL)
                If UDBSDebugMode Then
                    logger.Debug("Executing network query: " & SSQL)
                End If
            Catch ex As Exception
                Throw New UDBSException($"Failed to execute network query: '{SSQL}'. {ex.Message}", ex)
            End Try
        End Sub

        Friend Sub ExecuteLocalQuery(SSQL As String, scope As ITransactionScope)
            Try
                If dbLocal Is Nothing Then InitializeLocalDB()
                dbLocal.Execute(SSQL, scope)
                If UDBSDebugMode Then
                    logger.Debug("Executing local query: " & SSQL)
                End If
            Catch ex As Exception
                scope.HasError = True
                Throw New UDBSException($"Failed to execute local query: '{SSQL}'. {ex.Message}", ex)
            End Try
        End Sub


        Friend Sub ExecuteLocalQuery(SSQL As String)
            Try
                If dbLocal Is Nothing Then InitializeLocalDB()
                If UDBSDebugMode Then
                    logger.Debug("Executing local query: " & SSQL)
                End If
                dbLocal.Execute(SSQL)
            Catch ex As Exception
                Throw New UDBSException($"Failed to execute local query: '{SSQL}'. {ex.Message}", ex)
            End Try
        End Sub


        Friend Function InsertNetworkRecord(columnNames As String(), columnValues As Object(), tableName As String,
                                            scope As ITransactionScope, Optional ByVal primaryKey As String = "") _
            As Integer
            Return CInt(InsertNetworkRecord64Bits(columnNames, columnValues, tableName,
                                            scope, primaryKey))
        End Function

        ''' <summary>
        '''     Simple utility using T-SQL for inserting records given name-value pairs, and a table name.
        '''     Similar to <see cref="InsertNetworkRecord"/> but returns a Long to support BigInt data types in Network DB.
        ''' </summary>
        ''' <param name="columnNames"></param>
        ''' <param name="columnValues"></param>
        ''' <param name="tableName"></param>
        ''' <param name="scope"></param>
        ''' <returns>Row ID as Long</returns>
        Friend Function InsertNetworkRecord64Bits(columnNames As String(), columnValues As Object(), tableName As String,
                                            scope As ITransactionScope, Optional ByVal primaryKey As String = "") _
            As Long
            Try
                If dbNetwork Is Nothing Then InitializeNetworkDB()

                If (_preNetworkQueryHook IsNot Nothing) Then
                    _preNetworkQueryHook.Invoke($"INSERT INTO [{tableName}] ...")
                End If

                Return dbNetwork.InsertRecord64Bits(columnNames, columnValues, tableName, scope, primaryKey)
            Catch ex As Exception
                scope.HasError = True
                Throw New UDBSException($"Failed to insert a network record to: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Function

        Friend Function InsertLocalRecord(columnNames As String(), columnValues As Object(), tableName As String,
                                          scope As ITransactionScope, Optional ByVal primaryKey As String = "") _
            As Integer
            Try
                If dbLocal Is Nothing Then InitializeLocalDB()
                Return dbLocal.InsertRecord(columnNames, columnValues, tableName, scope, primaryKey)
            Catch ex As Exception
                scope.HasError = True
                Throw New UDBSException($"Failed to insert a local record to: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Function

        ''' <summary>
        ''' Insert multiple records in the local DB.
        ''' Similar to <see cref="InsertLocalRecord(String(), Object(), String, ITransactionScope, String)"/> but records
        ''' are bundled in a single operation for the sake of performance.
        ''' </summary>
        ''' <param name="columnNames">The names of the columns.</param>
        ''' <param name="table">The data to insert.</param>
        ''' <param name="tableName">The name of the table.</param>
        ''' <param name="transaction">(Optional) The transaction scope. A new, temporary transaction is created if none is provided.</param>
        Friend Sub InsertLocalRecords(columnNames As String(), table As DataTable, tableName As String, Optional transaction As ITransactionScope = Nothing)
            Dim newTransaction As Boolean = False
            If transaction Is Nothing Then
                transaction = dbLocal.BeginTransaction()
                newTransaction = True
            End If

            Try
                dbLocal.BulkInsertRecords(columnNames, table, tableName, transaction)
            Finally
                If newTransaction Then
                    transaction.Dispose()
                End If
            End Try
        End Sub

        Friend Function InsertLocalRecord(columnNames As String(), columnValues As Object(), tableName As String,
                                          Optional ByVal primaryKey As String = "") As Integer
            Try
                If dbLocal Is Nothing Then InitializeLocalDB()
                Return dbLocal.InsertRecord(columnNames, columnValues, tableName, primaryKey)
            Catch ex As Exception
                Throw New UDBSException($"Failed to insert a local record to: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Function

        Friend Function InsertNetworkRecord(columnNames As String(), columnValues As Object(), tableName As String,
                                            Optional ByVal primaryKey As String = "") As Integer

            Return CInt(InsertNetworkRecord64Bits(columnNames, columnValues, tableName, primaryKey))
        End Function

        ''' <summary>
        '''     Simple utility using T-SQL for inserting records given name-value pairs, and a table name.
        '''     Similar to <see cref="InsertNetworkRecord"/> but returns a Long to support BigInt data types in Network DB.
        ''' </summary>
        ''' <param name="columnNames"></param>
        ''' <param name="columnValues"></param>
        ''' <param name="tableName"></param>
        ''' <returns>Row ID as Long</returns>
        Private Function InsertNetworkRecord64Bits(columnNames As String(), columnValues As Object(), tableName As String,
                                            Optional ByVal primaryKey As String = "") As Long
            Try
                If dbNetwork Is Nothing Then InitializeNetworkDB()

                If (_preNetworkQueryHook IsNot Nothing) Then
                    _preNetworkQueryHook.Invoke($"INSERT INTO [{tableName}] ...")
                End If

                Return dbNetwork.InsertRecord64Bits(columnNames, columnValues, tableName, primaryKey)
            Catch ex As Exception
                Throw New UDBSException($"Failed to insert a network record to: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Function

        Friend Sub BulkInsertNetwork(columnNames As String(), sourceTable As DataTable, tableName As String, scope As ITransactionScope)
            Try
                If dbNetwork Is Nothing Then InitializeNetworkDB()

                If (_preNetworkQueryHook IsNot Nothing) Then
                    _preNetworkQueryHook.Invoke($"(BULK) INSERT INTO [{tableName}] ...")
                End If

                dbNetwork.BulkInsertRecords(columnNames, sourceTable, tableName, scope)
            Catch ex As Exception
                scope.HasError = True
                Throw New UDBSException($"Failed to bulk-insert network records to: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Sub

        Friend Sub BulkUpdateNetwork(columnNames As String(), matchedColumns As String(), sourceTable As DataTable, tableName As String, scope As ITransactionScope)
            Try
                If dbNetwork Is Nothing Then InitializeNetworkDB()

                If (_preNetworkQueryHook IsNot Nothing) Then
                    _preNetworkQueryHook.Invoke($"(BULK) INSERT INTO [{tableName}] ...")
                End If

                dbNetwork.BulkUpdateRecords(columnNames, matchedColumns, sourceTable, tableName, scope)
            Catch ex As Exception
                scope.HasError = True
                Throw New UDBSException($"Failed to bulk-insert network records to: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Sub

        Friend Function UpdateNetworkRecord(constraintKeys As String(), columnNames As String(),
                                            columnValues As Object(), tableName As String) As Boolean
            Try
                If dbNetwork Is Nothing Then InitializeNetworkDB()
                Return dbNetwork.UpdateRecord(constraintKeys, columnNames, columnValues, tableName)
            Catch ex As Exception
                Throw New UDBSException($"Failed to update a network record at: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Function

        Friend Function UpdateNetworkRecord(constraintKeys As String(), columnNames As String(),
                                            columnValues As Object(), tableName As String, scope As ITransactionScope) _
            As Boolean
            Try
                If dbNetwork Is Nothing Then InitializeNetworkDB()
                Return dbNetwork.UpdateRecord(constraintKeys, columnNames, columnValues, tableName, scope)
            Catch ex As Exception
                scope.HasError = True
                Throw New UDBSException($"Failed to update a network record at: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Function

        Friend Function UpdateLocalRecord(constraintKeys As String(), columnNames As String(), columnValues As Object(),
                                          tableName As String, scope As ITransactionScope) As Boolean
            Try
                If dbLocal Is Nothing Then InitializeLocalDB()
                Return dbLocal.UpdateRecord(constraintKeys, columnNames, columnValues, tableName, scope)
            Catch ex As Exception
                Throw New UDBSException($"Failed to update a local record at: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Function

        Friend Function UpdateLocalRecord(constraintKeys As String(), columnNames As String(), columnValues As Object(),
                                          tableName As String) As Boolean
            Try
                If dbLocal Is Nothing Then InitializeLocalDB()
                Return dbLocal.UpdateRecord(constraintKeys, columnNames, columnValues, tableName)
            Catch ex As Exception
                Throw New UDBSException($"Failed to update a local record at: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Function

        ' Candidate for removal.
        Private Function DeleteNetworkRecord(constraintKeys As String(), columnNames As String(), tableName As String) _
            As Boolean
            Try
                If dbNetwork Is Nothing Then InitializeNetworkDB()
                Return dbNetwork.DeleteRecord(constraintKeys, columnNames, tableName)
            Catch ex As Exception
                Throw New UDBSException($"Failed to delete a network record at: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Function

        ' Candidate for removal.
        Private Function DeleteLocalRecord(constraintKeys As String(), columnNames As String(), tableName As String) _
            As Boolean
            Try
                If dbLocal Is Nothing Then InitializeLocalDB()
                Return dbLocal.DeleteRecord(constraintKeys, columnNames, tableName)
            Catch ex As Exception
                Throw New UDBSException($"Failed to delete a local record at: {tableName} | Error: {ex.Message}", ex)
            End Try
        End Function

        Friend Sub OpenNetworkRecordSet(ByRef MyRS As DataRow, ByRef blob As Stream, SSQL As String)
            Try
                If dbNetwork Is Nothing Then InitializeNetworkDB()
                If UDBSDebugMode Then
                    logger.Debug("Opening Network DB Query: " & SSQL)
                End If
                MyRS = dbNetwork.ExecuteData(SSQL, blob)
            Catch ex As Exception
                Throw New UDBSException($"Failed to open network recordset: {SSQL} | Error: {ex.Message}", ex)
            End Try
        End Sub

        Friend Sub OpenNetworkRecordSet(ByRef MyRS As DataTable, SSQL As String)

            Try
                If (_preNetworkQueryHook IsNot Nothing) Then
                    Dim result = _preNetworkQueryHook.Invoke(SSQL)
                    If (result IsNot Nothing AndAlso Not String.IsNullOrEmpty(result.Item1)) Then
                        SSQL = result.Item1
                    End If
                    If (result IsNot Nothing AndAlso result.Item2 IsNot Nothing) Then
                        ' Data was provided, don't perform the query.
                        MyRS = result.Item2
                        Return
                    End If
                End If

                If dbNetwork Is Nothing Then InitializeNetworkDB()
                MyRS = dbNetwork.ExecuteData(SSQL)
                If UDBSDebugMode Then
                    logger.Debug("Opening Network DB Query: " & SSQL)
                End If
            Catch ex As Exception
                Throw New UDBSException($"Failed to open network recordset: {SSQL} | Error: {ex.Message}", ex)
            End Try
        End Sub

        Friend Function OpenLocalRecordSet(ByRef MyRS As DataTable, SSQL As String) As ReturnCodes
            Dim RET = ReturnCodes.UDBS_OP_SUCCESS
            Dim udbsEx As UDBSException = Nothing
            Try
                If dbLocal Is Nothing Then InitializeLocalDB()
                If UDBSDebugMode Then
                    logger.Debug("Opening Local DB Query: " & SSQL)
                End If
                MyRS = dbLocal.ExecuteData(SSQL)
            Catch ex As ExternalException
                If (ex.ErrorCode >= 0 AndAlso
                    ex.Message.IndexOf("no such table", StringComparison.InvariantCultureIgnoreCase) >= 0) Then
                    Return ReturnCodes.UDBS_TABLE_MISSING
                ElseIf ex.ErrorCode = ReturnCodes.UDBS_TABLE_MISSING Then
                    Return ReturnCodes.UDBS_TABLE_MISSING
                Else
                    udbsEx = New UDBSException($"Failed to open local recordset: {SSQL} | Error: {ex.Message}", ex)
                End If
            Catch ex As Exception
                udbsEx = New UDBSException($"Failed to open local recordset: {SSQL} | Error: {ex.Message}", ex)
            End Try

            If udbsEx IsNot Nothing Then
                LogErrorInDatabase(udbsEx)
                RET = ReturnCodes.UDBS_ERROR
            End If
            Return RET
        End Function

        Friend Sub OpenLocalRecordSet(ByRef MyRS As DataRow, ByRef blob As Stream, SSQL As String)
            Try
                If dbLocal Is Nothing Then InitializeLocalDB()
                If UDBSDebugMode Then
                    logger.Debug("Opening Network DB Query: " & SSQL)
                End If
                MyRS = dbLocal.ExecuteData(SSQL, blob)
            Catch ex As Exception
                Throw New UDBSException($"Failed to open local recordset: {SSQL} | Error: {ex.Message}", ex)
            End Try
        End Sub

        ' Candidate for removal.
        Private Function OpenLocalRecordSet(ByRef MyRS As DataTable, SSQL As String, scope As ITransactionScope) _
            As ReturnCodes
            Dim RET = ReturnCodes.UDBS_OP_SUCCESS
            Dim udbsEx As UDBSException = Nothing
            Try
                If dbLocal Is Nothing Then InitializeLocalDB()
                MyRS = dbLocal.ExecuteData(SSQL, scope)
                If UDBSDebugMode Then
                    logger.Debug("Opening Local DB Query: " & SSQL)
                End If
            Catch ex As ExternalException
                If _
                    (ex.ErrorCode >= 0 AndAlso
                     ex.Message.IndexOf("no such table", StringComparison.InvariantCultureIgnoreCase) >= 0) Then
                    Return ReturnCodes.UDBS_TABLE_MISSING
                ElseIf ex.ErrorCode = ReturnCodes.UDBS_TABLE_MISSING Then
                    Return ReturnCodes.UDBS_TABLE_MISSING
                Else
                    udbsEx = New UDBSException($"Failed to open local recordset: {SSQL} | Error: {ex.Message}", ex)
                End If
            Catch ex As Exception
                udbsEx = New UDBSException($"Failed to open local recordset: {SSQL} | Error: {ex.Message}", ex)
            End Try

            If udbsEx IsNot Nothing Then
                LogErrorInDatabase(udbsEx)
                RET = ReturnCodes.UDBS_ERROR
            End If
            Return RET
        End Function

        Private Sub OpenNetworkDB()
            If dbNetwork Is Nothing Then InitializeNetworkDB()
            If UDBSDebugMode Then
                logger.Debug("Opening Network Connection")
            End If
            dbNetwork.CommandTimeOut = DBCommandTimeout
        End Sub

        Friend Sub OpenNetworkDB(CommandTimeout As Integer)
            OpenNetworkDB()
            dbNetwork.CommandTimeOut = CommandTimeout
        End Sub

        Friend Function GetLocalAdapter(sqlSelectQuery As String, scope As ITransactionScope,
                                        ByRef workTable As DataTable) As IAdapterSession
            Try
                Return dbLocal.CreateTableAdapter(sqlSelectQuery, scope, workTable)
            Catch ex As Exception
                Throw New UDBSException($"Failed to create local table adapter: {sqlSelectQuery} | Error: {ex.Message}", ex)
            End Try
        End Function

        ' Candidate for removal.
        Private Function GetNetworkAdapter(sqlSelectQuery As String, scope As ITransactionScope,
                                          ByRef workTable As DataTable) As IAdapterSession
            Try
                Return dbNetwork.CreateTableAdapter(sqlSelectQuery, scope, workTable)
            Catch ex As Exception
                Throw New UDBSException($"Failed to create network table adapter: {sqlSelectQuery} | Error: {ex.Message}", ex)
            End Try
        End Function

        ''' <summary>
        '''     Perform atomic Queries
        '''     Call Dispose() to the returned object after you're done
        '''     This will commit everything within the scope's life
        ''' </summary>
        ''' <returns></returns>
        Friend Function BeginNetworkTransaction() As ITransactionScope
            If dbNetwork Is Nothing Then InitializeNetworkDB()
            Return dbNetwork.BeginTransaction()
        End Function

        Friend Function BeginLocalTransaction() As ITransactionScope
            If dbLocal Is Nothing Then InitializeLocalDB()
            Return dbLocal.BeginTransaction()
        End Function

        ''' <summary>
        '''     Perform a query with a given scope
        '''     Note: Call <see cref="BeginNetworkTransaction" /> first!
        ''' </summary>
        ''' <param name="MyRS"></param>
        ''' <param name="SSQL"></param>
        ''' <param name="scope"></param>
        Friend Sub OpenNetworkRecordSet(ByRef MyRS As DataTable, SSQL As String, scope As ITransactionScope)
            Try
                If UDBSDebugMode Then
                    logger.Debug("Opening Network DB Query: " & SSQL)
                End If
                MyRS = dbNetwork.ExecuteData(SSQL, scope)
            Catch ex As Exception
                scope.HasError = True
                Throw New UDBSException($"Failed to open network recordset: {SSQL} | Error: {ex.Message}", ex)
            End Try
        End Sub

        Friend Function IsNetworkConnectionOpen() As Boolean
            If dbNetwork Is Nothing Then InitializeNetworkDB()
            Return dbNetwork.SystemAvailable
        End Function

        Friend Sub CloseNetworkDB()
            ' We don't close network connections as we take advantage of built-in connection pooling
            ' TODO - remove this
        End Sub

        ''' <summary>
        ''' verifies if the process tables exist locally, if not creates them
        ''' </summary>
        ''' <param name="ProcessName"></param>
        ''' <returns></returns>
        Friend Function CheckLocalTables(ProcessName As String) As ReturnCodes

            If _localProcessTablesCreated.Contains(ProcessName) Then
                Console.WriteLine($">> We've already checked if the local tables exist for process '{ProcessName}'")
                Return ReturnCodes.UDBS_OP_SUCCESS
            Else
                Console.WriteLine($">> Checking if the local tables exist for process '{ProcessName}'...")
            End If

            Try
                Dim sqlQuery As String
                Dim rsTemp As New DataTable
                If UDBSDebugMode Then
                    logger.Debug("Checking UDBS Local Database " & ProcessName & " Tables")
                End If

                If LocalDBDriver = LocalDriverEnum.SQLite Then
                    sqlQuery =
                        "CREATE TABLE IF NOT EXISTS process_registration (
                            pr_id INTEGER PRIMARY KEY,
                            pr_mutex_name TEXT,
                            pr_process TEXT NOT NULL,
                            pr_process_id INTEGER
                        );"
                    If dbLocal Is Nothing Then InitializeLocalDB()
                    dbLocal.Execute(sqlQuery)
                End If

                CheckAndCreateLocalTable(
                    ProcessName,
                    ProcessName & "_process",
                    "SELECT process_id FROM " & ProcessName & "_process WHERE process_id = 1")
                CheckAndCreateLocalTable(
                    ProcessName,
                    ProcessName & "_result",
                    "SELECT result_id FROM " & ProcessName & "_result WHERE result_id = 1")
                CheckAndCreateLocalTable(
                    ProcessName,
                    ProcessName & "_blob",
                    "SELECT blob_id FROM " & ProcessName & "_blob WHERE blob_id = 1")

                _localProcessTablesCreated.Add(ProcessName)
                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Checks whether or not a table exists. If it doesn't exist, create it.
        ''' </summary>
        ''' <param name="processName">The name of the process. Only used for logging in case of an error.</param>
        ''' <param name="tableName">The name of the table to check and create.</param>
        ''' <param name="sqlQuery">The SQL query to perform in order to check if the table exists.</param>
        ''' <exception cref="Exception">If an error occurs while trying to determine whether or not the table exist, or while creating the table.</exception>
        ''' <remarks>
        ''' There is probably a better way to tell whether or not a table exists...
        ''' For example:
        '''   CREATE TABLE IF NOT EXISTS ...
        ''' </remarks>
        Private Sub CheckAndCreateLocalTable(processName As String, tableName As String, sqlQuery As String)
            If CreateLocalTable(tableName) <> ReturnCodes.UDBS_OP_SUCCESS Then
                Throw New Exception($"Failed to create local Blob table for process {processName}.")
            End If
        End Sub

        ' Candidate for removal.
        Private Function GetNetworkDBColumns(tableName As String) As List(Of Tuple(Of String, String, Integer))
            Dim sql As String = "SELECT top 1 * " &
                                "FROM " & tableName & " with(nolock)"
            OpenNetworkDB()
            Return dbNetwork.GetColumnTypes(sql)
        End Function
        Private Function CreateLocalTable(tableName As String) _
            As ReturnCodes
            If LocalDBDriver = LocalDriverEnum.SQLite Then
                Return CreateLocalTable_SqlLite(tableName)
            ElseIf LocalDBDriver = LocalDriverEnum.MsAccess Then
                Return CreateLocalTable_Access(tableName)
            End If
            Throw New ApplicationException("Invalid LocalDBDriver state")
        End Function

        ''' <summary>
        ''' Get the DB key of a given UDBS table.
        ''' </summary>
        ''' <remarks>
        ''' Not every table names are supported.
        ''' This is only meant to be used for the different UDBS process tables.
        ''' </remarks>
        ''' <param name="TableName">The name of the UDBS table.</param>
        ''' <returns>The name of the key for that table.</returns>
        Private Function GetDbKey(TableName As String) As String
            If InStr(1, TableName, "_itemlistrevision") > 0 Then
                Return "itemlistrev_id"
            ElseIf InStr(1, TableName, "_itemlistdefinition") > 0 Then
                Return "itemlistdef_id"
            ElseIf InStr(1, TableName, "_process") > 0 Then
                Return "process_id"
            ElseIf InStr(1, TableName, "_result") > 0 Then
                Return "result_id"
            ElseIf InStr(1, TableName, "_blob") > 0 Then
                Return "blob_id"
            ElseIf TableName = "product" Then ' Not used.
                Return "product_id"
            ElseIf TableName = "unit" Then ' Not used.
                Return "unit_id"
            Else
                ' Unknown table, just select all
                logger.Warn($"Unknown UDBS Process table: {TableName}")
                Return "1"
            End If
        End Function

        Private Function GetColumnTypes(TableName As String) As List(Of Tuple(Of String, String, Integer))
            Dim sqlQuery = "SELECT top 1 * " &
                           "FROM " & TableName & " with(nolock) " &
                           "WHERE " & GetDbKey(TableName) & " = 1 "
            OpenNetworkDB()
            Return dbNetwork.GetColumnTypes(sqlQuery)
        End Function

        ''' <summary>
        ''' Create the local SQLite database table.
        ''' </summary>
        ''' <param name="TableName">The name of the table to create.</param>
        ''' <returns>The outcome of the operation.</returns>
        Private Function CreateLocalTable_SqlLite(TableName As String) _
            As ReturnCodes
            ' Function copies an existing table from the network to the local database
            Try
                Dim dbKeyColumnName As String
                Dim PrimaryField As String
                Dim TypeField As String
                Dim sqlCommand As New StringBuilder()

                dbKeyColumnName = GetDbKey(TableName)

                Dim columnInformation = GetColumnTypes(TableName)

                ' Begin building table creation command
                sqlCommand.Append($"CREATE TABLE IF NOT EXISTS {TableName} (")

                Dim counter = 0
                For Each FieldItem In columnInformation
                    TypeField = FieldItem.Item2

                    If FieldItem.Item2.Equals("int", StringComparison.InvariantCultureIgnoreCase) Then
                        TypeField = "integer"
                    ElseIf FieldItem.Item2.Equals("bigint", StringComparison.InvariantCultureIgnoreCase) Then
                        ' SQLite doesn't like 'bigint' data type.
                        ' We're not expecting the local DB results to exceed 32-bits integer capacity; it took
                        ' years and many throusand of units to accumulate that much data.
                        ' It is safe to assume a single stage for a single unit will not contain that much data.
                        TypeField = "integer"
                    End If

                    ' Change to an AUTONUMBER field on result tables...
                    If counter = 0 Then
                        sqlCommand.Append($"{FieldItem.Item1} {TypeField} primary key") _
                        ' no need to create index as well
                        If dbKeyColumnName = "result_id" And FieldItem.Item1 = "result_id" Then
                            sqlCommand.Append(" AUTOINCREMENT")
                        End If

                        If dbKeyColumnName = "blob_id" And FieldItem.Item1 = "blob_id" Then
                            sqlCommand.Append(" AUTOINCREMENT")
                        End If
                    Else
                        sqlCommand.Append($"{FieldItem.Item1} {TypeField}")
                    End If

                    ' add precision for size varcharfields
                    If TypeField = "VARCHAR" Then
                        sqlCommand.Append($"(").Append(Format(FieldItem.Item3)).Append(")")
                    End If

                    sqlCommand.Append(", ")
                    counter += 1
                Next FieldItem

                PrimaryField = columnInformation.First().Item1

                ' Remove last comma
                sqlCommand.Length -= 2
                sqlCommand.Append(");")

                dbLocal.Execute(sqlCommand.ToString())

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(New UDBSException($"Failed to create local database | Error: {ex.Message}", ex))
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Create a table in the local UDBS cache, when that local DB is using Microsoft Access.
        ''' This is no longer supported.
        ''' </summary>
        ''' <remarks>MS Access is no longer supported. Candidate for removal.</remarks>
        Private Function CreateLocalTable_Access(TableName As String) As ReturnCodes
            ' Function copies an existing table from the network to the local database
            Try
                Dim IDCol As String
                Dim PrimaryField As String
                Dim TypeField As String
                Dim sqlQuery As String
                Dim sqlCommand As String

                If InStr(1, TableName, "_itemlistrevision") > 0 Then
                    IDCol = "itemlistrev_id"
                ElseIf InStr(1, TableName, "_itemlistdefinition") > 0 Then
                    IDCol = "itemlistdef_id"
                ElseIf InStr(1, TableName, "_process") > 0 Then
                    IDCol = "process_id"
                ElseIf InStr(1, TableName, "_result") > 0 Then
                    IDCol = "result_id"
                ElseIf InStr(1, TableName, "_blob") > 0 Then
                    IDCol = "blob_id"
                ElseIf TableName = "product" Then
                    IDCol = "product_id"
                ElseIf TableName = "unit" Then
                    IDCol = "unit_id"
                Else
                    ' Unknown table, just select all
                    IDCol = "1"
                End If

                sqlQuery = "SELECT top 1 * " &
                           "FROM " & TableName & " with(nolock) " &
                           "WHERE " & IDCol & " = 1 "
                OpenNetworkDB()
                Dim columnInformation As List(Of Tuple(Of String, String, Integer)) = dbNetwork.GetColumnTypes(sqlQuery)

                ' Begin building table creation command
                sqlCommand = "CREATE TABLE " &
                             TableName & " ("

                For Each FieldItem In columnInformation
                    TypeField = FieldItem.Item2

                    ' Change to an AUTONUMBER field on result tables...
                    If IDCol = "result_id" And FieldItem.Item1 = "result_id" Then
                        TypeField = "COUNTER"
                    End If

                    If IDCol = "blob_id" And FieldItem.Item1 = "blob_id" Then
                        TypeField = "COUNTER"
                    End If

                    sqlCommand = sqlCommand & FieldItem.Item1 & " " & TypeField

                    ' add precision for size varcharfields
                    If TypeField = "VARCHAR" Then
                        sqlCommand = sqlCommand & "(" & Format(FieldItem.Item3) & ")"
                    End If
                    sqlCommand = sqlCommand & ", "
                Next FieldItem

                PrimaryField = columnInformation.First().Item1

                ' Remove last comma
                sqlCommand = Left$(sqlCommand, Len(sqlCommand) - 2) & ");"

                dbLocal.Execute(sqlCommand)

                sqlCommand = "CREATE UNIQUE INDEX pk_" & TableName & " ON " & TableName & " (" & PrimaryField &
                             ") WITH PRIMARY"

                dbLocal.Execute(sqlCommand)

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ' This is process-specific and belong in the CProcessInstance class.
        ''' <summary>
        ''' Remove local tables when process finishes.
        ''' </summary>
        ''' <returns>The outcome of the operation.</returns>
        Friend Function DropLocalTables(ProcessName As String) As ReturnCodes

            Try
                If UDBSDebugMode Then
                    logger.Debug("Dropping UDBS Local Database " & ProcessName & " Tables.")
                End If

                logger.Info("Dropping local tables.")
                Dim TableName As String

                TableName = ProcessName & "_result"
                dbLocal.Execute("DROP TABLE IF EXISTS " & TableName)

                TableName = ProcessName & "_process"
                dbLocal.Execute("DROP TABLE IF EXISTS " & TableName)

                TableName = ProcessName & "_itemlistdefinition"
                dbLocal.Execute("DROP TABLE IF EXISTS " & TableName)

                TableName = ProcessName & "_itemlistrevision"
                dbLocal.Execute("DROP TABLE IF EXISTS " & TableName)

                TableName = ProcessName & "_blob"
                dbLocal.Execute("DROP TABLE IF EXISTS " & TableName)

                _localProcessTablesCreated.Remove(ProcessName)

                Return ReturnCodes.UDBS_OP_SUCCESS
            Catch e As Exception
                LogErrorInDatabase(e)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Checking if a given table is one of the core UDBS tables.
        ''' </summary>
        ''' <param name="TableName">The name of the table to validate.</param>
        ''' <returns>Whether or not this is a core UDBS table.</returns>
        ''' <remarks>
        ''' This is only used in the "Instrument Attributes" feature, which seems
        ''' to be a prototype or future feature that has not yet been enabled.
        ''' </remarks>
        Friend Function IsCoreTable(TableName As String) _
            As Boolean
            ' Function checks if named table is a core process table
            If TableName = "product" Or
               TableName = "product_family" Or
               TableName = "unit" Or
               TableName = "employee" Or
               InStr(1, TableName, "udbs_") > 0 Or
               InStr(1, TableName, "_process") > 0 Or
               InStr(1, TableName, "_result") > 0 Or
               InStr(1, TableName, "_itemlistdefinition") > 0 Or
               InStr(1, TableName, "_itemlistrevision") > 0 Then
                IsCoreTable = True
            Else
                IsCoreTable = False
            End If
        End Function

        ''' <summary>
        ''' Guards against null DB object reference.
        ''' Changes a DB object to a string, converting a null object to an empty string.
        ''' </summary>
        ''' <param name="DataValue">The DB object to convert.</param>
        ''' <returns>The DB object, as a string, or an empty string if the object is null.</returns>
        Friend Function KillNull(DataValue As Object) As String
            If Not IsDBNull(DataValue) Then
                Return CStr(DataValue)
            Else
                Return String.Empty
            End If
        End Function

        ''' <summary>
        ''' Guards against null DB object reference.
        ''' Changes a DB object to a double, converting a null object to the Double.NaN value.
        ''' </summary>
        ''' <param name="DataValue">The DB object to convert.</param>
        ''' <returns>The DB object, as a double, or NaN if the object is null.</returns>
        Friend Function KillNullDouble(DataValue As Object) As Double
            If IsDBNull(DataValue) Then
                Return Double.NaN
            Else
                Return Val(DataValue)
            End If
        End Function

        ''' <summary>
        ''' Guards against null DB object reference.
        ''' Changes a DB object to an integer, converting a null object to the value 0.
        ''' </summary>
        ''' <param name="DataValue">The DB object to convert.</param>
        ''' <returns>The DB object, as an integer, or the value 0 if the object is null.</returns>
        Friend Function KillNullInteger(DataValue As Object) As Integer
            If IsDBNull(DataValue) Then
                Return 0
            Else
                Return CInt(Val(DataValue))
            End If
        End Function

        ''' <summary>
        ''' Guards against null DB object reference.
        ''' Changes a DB object to an byte, converting a null object to the value 0.
        ''' </summary>
        ''' <param name="DataValue">The DB object to convert.</param>
        ''' <returns>The DB object, as a byte, or the value 0 if the object is null.</returns>
        Friend Function KillNullByte(DataValue As Object) As Byte
            If IsDBNull(DataValue) Then
                Return 0
            Else
                Return CByte(Val(DataValue))
            End If
        End Function

        ''' <summary>
        ''' Guards against null DB object reference.
        ''' Changes a DB object to a 64-bits integer, converting a null object to the value 0.
        ''' </summary>
        ''' <param name="DataValue">The DB object to convert.</param>
        ''' <returns>The DB object, as a 64-bits integer, or the value 0 if the object is null.</returns>
        Friend Function KillNullLong(DataValue As Object) As Long
            If IsDBNull(DataValue) Then
                Return 0
            Else
                Return CType(DataValue, Long)
            End If
        End Function

        ''' <summary>
        ''' Guards against null DB object reference and non-date objects.
        ''' Changes a DB object to a date, converting a null object to "epoch".
        ''' </summary>
        ''' <param name="DataValue">The DB object to convert.</param>
        ''' <returns>The DB object, as a date, or the "epoch" if the object is null or not a date.</returns>
        Friend Function KillNullDate(DataValue As Object) As Date
            If IsDBNull(DataValue) Then
                Return Date.MinValue
            Else
                If IsDate(DataValue) Then
                    Return CDate(DataValue)
                Else
                    Return Date.MinValue
                End If
            End If
        End Function

        ''' <summary>
        ''' Formats a date to the string value expected in a SQL query.
        ''' </summary>
        ''' <param name="Value">The date to format.</param>
        ''' <returns>The string value.</returns>
        ''' <remarks>
        ''' The name is somewhat confusing: is this the date format (noun) or the action of formatting?
        ''' If you simply look at the name of this method, you can't tell, then you have to look at the
        ''' method's signature.
        ''' Should be renamed to "FormatDateForDB(...)", for example.
        ''' </remarks>
        Friend Function DBDateFormat(Value As Date) As String
            Return Value.ToString(DBDateFormatting, enUS) '"yyyy/mm/dd HH:mm:ss")
        End Function

        ''' <summary>
        ''' Parses a date from the DB format to a System.DateTime object.
        ''' </summary>
        ''' <param name="toParse">The string to parse.</param>
        ''' <returns>The DateTime object matching the date expressin in that string.</returns>
        ''' <remarks>If/when <see cref="DBDateFormat(Date)"/> gets renamed, also rename this one to be reflective.</remarks>
        Friend Function DBDateParse(toParse As String) As DateTime
            Return Date.ParseExact(toParse, DBDateFormatting, CultureInfo.InvariantCulture)
        End Function

        Friend Function QueryNetworkDB(sqlQuery As String,
                                       ByRef rsResults As DataTable) _
            As ReturnCodes
            ' Function opens connection to DB and populates a recordset, then returns a close to the calling function

            rsResults = New DataTable
            Try
                OpenNetworkDB(120)

                OpenNetworkRecordSet(rsResults, sqlQuery)
                If Left(sqlQuery, 6) = "SELECT" Then

                    If rsResults?.Rows.Count >= 0 Then
                        If UDBSDebugMode Then
                            logger.Debug("Creating Network DB Recordset Clone")
                        End If

                        Return ReturnCodes.UDBS_OP_SUCCESS
                    Else
                        ' Query could not be processed
                        Return ReturnCodes.UDBS_OP_FAIL
                    End If
                Else
                    'Assume no query is returned
                    rsResults = Nothing
                    Return ReturnCodes.UDBS_OP_SUCCESS
                End If

            Catch ex As Exception
                LogErrorInDatabase(ex)
                Return ReturnCodes.UDBS_ERROR
            End Try
        End Function

        ''' <summary>
        ''' Uses reflection to determine the name of the software using the UDBS Interface assembly.
        ''' </summary>
        ''' <returns>The name and version of the software using the UDBS Interface.</returns>
        Friend Function DetermineSoftwareName() As String
            Return $"{My.Application.Info.Title} v{My.Application.Info.Version}"
        End Function

        ''' <summary>
        ''' Log the UDBS Interface's assembly version usage for traceability.
        ''' At the moment, this logs the assembly information to the "udbs_error" table.
        ''' In the future, we could either log to a new set of UDBS tables, or post
        ''' to Elastic Stack.
        ''' </summary>
        Private Sub LogAssemblyVersion()
            If (_assemblyVersionHasBeenLogged) Then
                Return
            End If

            ' Raise this flag right away.
            ' If the database has not been initialized yet, a call will be made to
            ' InitializeNetworkDb(...) as part of processing the error entry added
            ' to the queue near the end.
            ' If the user of this library doesn't start its interaction with a call
            ' to 'LogAssemblyVersion', it will be logged in there.
            _assemblyVersionHasBeenLogged = True

            Try
                Dim myAssembly = GetType(DatabaseSupport).Assembly

                Dim softwareName = DetermineSoftwareName()
                Dim stationId As String = Nothing
                If CUtility.Utility_GetStationName(stationId) <> ReturnCodes.UDBS_OP_SUCCESS Then
                    stationId = Environment.MachineName
                End If

                Dim entryToLog = New ErrorQueue.ErrorEntry With {
                .Guid = Guid.NewGuid().ToString(),
                .Description = $"Application ""{softwareName}"" launched.",
                .FunctionName = NameOf(DatabaseSupport),
                .PcDisplayName = stationId,
                .StackTrace = String.Empty,
                .AssemblyVersion = GenerateAssemblyVersionString(myAssembly)}

                ' Ignore return code; local time is returned if the server is not available.
                CUtility.Utility_GetServerTime(entryToLog.TimeOfEvent)

                SystemErrorQueue.Add(entryToLog)
            Catch
                ' We failed to log the assembly version.
                ' Clear the flag; we will try again next time.
                _assemblyVersionHasBeenLogged = False
                Throw
            End Try
        End Sub


        Friend Sub LogError(ex As Exception)
            If ex Is Nothing Then
                Return
            End If

            If (UdbsTools.ErrorLogLevel = Nothing) Then
                logger.Log(LogLevel.Debug, ex)
            Else
                logger.Log(UdbsTools.ErrorLogLevel, ex)
            End If
        End Sub

        ''' <summary>
        ''' Generates the version string of an assembly, for logging into UDBS as the application starts up.
        ''' Adds the 'S' suffix (for "shared") if the assembly is loaded from the GAC.
        ''' Adds the 'T' suffix (for "trial") if the assembly is registered for trial in the Windows Registry.
        ''' </summary>
        ''' <param name="anAssembly">The assembly to evaluate.</param>
        ''' <returns>
        ''' The string representing this assembly version.
        ''' Example:
        '''   UdbsInterface v3.42.0.49538*:T
        ''' </returns>
        Private Function GenerateAssemblyVersionString(anAssembly As Assembly) As String
            Dim name = anAssembly.GetName().Name
            Dim version = FileVersionInfo.GetVersionInfo(anAssembly.Location).FileVersion
            Dim isFromGAC As Boolean

            If String.IsNullOrEmpty(anAssembly.Location) Then
                'Running in Visual Studio, from source code.
                'Just use the assembly version (file location not available)
                isFromGAC = False
            Else
                isFromGAC = anAssembly.GlobalAssemblyCache
            End If

            Dim versionStr = $"{name} v{version}"

            If isFromGAC Then
                versionStr += ":S"
            End If

            If IsAssemblyTrial(anAssembly.GetName()) Then
                versionStr += ":T"
            End If

            Return versionStr
        End Function

        ''' <summary>
        ''' Checks whether or not the trial of an assembly is registered in the Windows Registry.
        ''' </summary>
        ''' <param name="anAssemblyName">The name of the assembly to evaluate.</param>
        ''' <returns>Whether or not a trial is in progress.</returns>
        ''' <remarks>
        ''' This method has been ported from the SRM Interface assembly.
        ''' The UDBS Interface and the SRM Interface assemblies do not share a common dependency
        ''' where code could be shared. Sadly, this mean this code had to be copied and pasted.
        ''' (Well, actually ported, since SRM Interface is implemented in C#.)
        ''' </remarks>
        Private Function IsAssemblyTrial(anAssemblyName As System.Reflection.AssemblyName) As Boolean
            Try
                Dim trialKey = $"SOFTWARE\\Lumentum\\Deployment\\{anAssemblyName.Name}\\v{anAssemblyName.Version.Major}"
                If (Environment.Is64BitOperatingSystem) Then
                    trialKey = $"SOFTWARE\\WOW6432Node\\Lumentum\\Deployment\\{anAssemblyName.Name}\\v{anAssemblyName.Version.Major}"
                End If

                Dim registryKey = Registry.LocalMachine.OpenSubKey(trialKey)
                If registryKey Is Nothing Then
                    Return False
                End If

                Dim versionTypeStr As String = CType(registryKey.GetValue("VersionType", String.Empty), String)
                Dim trialEnabled As Boolean
                If String.IsNullOrEmpty(versionTypeStr) Or versionTypeStr = "production" Then
                    trialEnabled = False
                ElseIf versionTypeStr = "trial" Then
                    trialEnabled = True
                Else
                    logger.Error($"Unexpected trial registration code: {versionTypeStr}")

                    ' There seem to be a problem with the trial registration itself.
                    ' Assume the trial Is ongoing.
                    ' This might capture the attention of the operator, factory supervisor, etc.
                    ' Otherwise, it Is possible this condition remains unnotice even if we log
                    ' an error log.
                    trialEnabled = True
                End If

                If trialEnabled Then
                    Dim trialEndDateStr As String = CType(registryKey.GetValue("TrialExpiryDate", String.Empty), String)
                    If String.IsNullOrEmpty(trialEndDateStr) Then
                        ' No expiry date for this trial.
                        Return True
                    Else
                        Dim trialEndDate As Date
                        If DateTime.TryParse(trialEndDateStr, trialEndDate) Then
                            Return trialEndDate > DateTime.Now
                        Else
                            logger.Error($"Failure to parse trial end date: {trialEndDateStr}")

                            ' There seem to be a problem with the trial registration itself.
                            ' Assume the trial is ongoing.
                            ' This might capture the attention of the operator, factory supervisor, etc.
                            ' Otherwise, it is possible this condition remains unnotice even if we log
                            ' an error log.
                            Return True
                        End If
                    End If
                Else
                    Return False
                End If
            Catch ex As Exception
                logger.Error(ex, $"Unexpected error trying to determine whether or not assembly {anAssemblyName?.Name} v{anAssemblyName?.Version?.Major} is a trial version.")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Gets udbs log level from configuration, and logs message with stack trace at the configured level.
        ''' </summary>
        Friend Sub LogError(message As String)
            If String.IsNullOrEmpty(message) Then
                Return
            End If

            If (UdbsTools.ErrorLogLevel = Nothing) Then
                logger.Log(LogLevel.Debug, message)
            Else
                logger.Log(UdbsTools.ErrorLogLevel, message)
            End If
        End Sub

        ''' <summary>
        ''' Logs the UDBS error in the application log and inserts it in the udbs_error table with process instance information.
        ''' </summary>
        ''' <param name="ex">Exception raised</param>
        ''' <param name="processContext">Context information to be displayed in the description field of the the udbs_error table.</param>
        Friend Sub LogErrorInDatabase(ex As Exception, Optional processContext As String = "")

            If ex Is Nothing Then
                Return
            End If

            ' Log error in the console.
            LogError(ex)

            Dim functionName As String = CStr(IIf(ex.TargetSite IsNot Nothing, $"{ex.TargetSite?.DeclaringType?.Name}::{ex.TargetSite?.Name}", "Unknown Function/Class name"))

            ' Make sure that we don't recurse into this method.
            ' If there is an error during the logging of an error, we
            ' don't want to also post that error, etc. This would
            ' create an infinite recursive loop.
            If CUtility.IsRecursiveLoop() Then
                Return
            End If

            ' Clean-up and format data to be logged.
            Dim pcDisplayName = $"{Replace(stationName, "'", ":")} | {PCname}"
            Dim applicationNameVersion = $"{My.Application.Info.Title} v{My.Application.Info.Version}"
            Dim description = $"Process: {processContext} | Application: {applicationNameVersion} | {functionName} | {ex.Message}"
            description = Replace(description, "'", ":")
            ' Remove the build folder occurrences from the stack trace for readability.
            Dim stackTrace = Replace(ex.StackTrace, buildFolder, "")

            Dim entryToLog = New ErrorQueue.ErrorEntry With {
                .Guid = Guid.NewGuid().ToString(),
                .Description = description,
                .FunctionName = functionName,
                .PcDisplayName = pcDisplayName,
                .StackTrace = stackTrace,
                .AssemblyVersion = AssemblyVersionString}

            ' Ignore return code; local time is returned if the server is not available.
            CUtility.Utility_GetServerTime(entryToLog.TimeOfEvent)

            SystemErrorQueue.Add(entryToLog)
        End Sub

        ''' <summary>
        ''' Logs the UDBS error in the application log and inserts it in the udbs_error table with process instance information.
        ''' </summary>
        ''' <param name="ex">Exception raised</param>
        ''' <param name="contextType">Process instance Type (Wip/Test/Kitting).</param>
        ''' <param name="name">Process instance name.</param>
        ''' <param name="processID">Process instance ID.</param>
        ''' <param name="product">Udbs product ID.</param>
        ''' <param name="serialNumber">Unit serial number.</param>
        Friend Sub LogErrorInDatabase(ex As Exception, ByVal contextType As String, ByVal name As String, ByVal processID As Integer, ByVal product As String, ByVal serialNumber As String)
            LogErrorInDatabase(ex, $"Type={contextType} Name={name} ID={processID} Product={product} SN={serialNumber}")
        End Sub

        ''' <summary>
        ''' Asserts that the UDBS return code of an operation is UDBS_OP_SUCCESS.
        ''' Throws an Exception otherwise.
        ''' </summary>
        ''' <param name="returnCode">The UDBS return code of an operation.</param>
        Friend Sub AssertOperationSucceeded(ByVal returnCode As ReturnCodes)

            If returnCode <> ReturnCodes.UDBS_OP_SUCCESS Then Throw New UDBSException($"The UDBS operation didn't succeed. Return code is: {returnCode}")
        End Sub

        ''' <summary>
        ''' Get the application's parent directory, without the application folder.
        ''' e.g If the application is Aura, will return: C:\Program Files\Lumentum
        ''' </summary>
        Private Function GetApplicationParentDirectory() As String

            Dim currentDirectoryArray = Split(My.Application.Info.DirectoryPath, "\")
            Dim builder As New StringBuilder

            If currentDirectoryArray.Count > 0 Then
                ' Get the folders in the path up to the application folder.
                For i As Integer = 0 To currentDirectoryArray.Length - 3
                    builder.Append($"{currentDirectoryArray(i)}\")
                Next
            End If

            Return builder.ToString
        End Function

        ''' <summary>
        ''' The name and version of the assembly (UDBS Interface)
        ''' </summary>
        Friend ReadOnly Property AssemblyVersionString As String
            Get
                Return _callingCode.Value
            End Get
        End Property

        ''' <summary>
        ''' Cached for less perf hit
        ''' </summary>
        Private ReadOnly _callingCode As New Lazy(Of String)(Function()

                                                                 Dim asm As Assembly = GetType(DatabaseSupport).Assembly
                                                                 Dim res As String = $"{asm.GetName().Name} v{asm.GetName().Version}"
                                                                 Try
                                                                     Dim fv = FileVersionInfo.GetVersionInfo(asm.Location)
                                                                     ' Use File Version to be accurate - we don't update assembly version that often
                                                                     res = $"{asm.GetName().Name} v{fv.FileVersion}"
                                                                 Catch
                                                                     ' ignore
                                                                 End Try
                                                                 Return res?.Substring(0, Math.Min(res.Length, 50))
                                                                 ' nvarchar(50)
                                                             End Function)

        ' Candidate for removal.
        <DebuggerStepThrough>
        Private Function ActualCommandTextByNames(sender As IDbCommand) As String
            Dim sb As New StringBuilder(sender.CommandText)
            Dim EmptyParameterNames =
                    (From T In sender.Parameters.Cast(Of IDataParameter)()
                     Where String.IsNullOrWhiteSpace(T.ParameterName)).FirstOrDefault

            If EmptyParameterNames IsNot Nothing Then
                Return sender.CommandText
            End If

            For Each p As IDataParameter In sender.Parameters

                Select Case p.DbType
                    Case Data.DbType.AnsiString, Data.DbType.AnsiStringFixedLength, Data.DbType.Date,
                        Data.DbType.DateTime,
                        Data.DbType.DateTime2, Data.DbType.Guid, Data.DbType.String, Data.DbType.StringFixedLength,
                        Data.DbType.Time, Data.DbType.Xml
                        If p.ParameterName(0) = "@" Then
                            If p.Value Is Nothing Then
                                Throw New Exception("no value given for parameter '" & p.ParameterName & "'")
                            End If
                            sb = sb.Replace(p.ParameterName, $"'{p.Value.ToString.Replace("'", "''")}'")
                        Else
                            sb = sb.Replace(String.Concat("@", p.ParameterName),
                                            $"'{p.Value.ToString.Replace("'", "''")}'")
                        End If
                    Case Else
                        sb = sb.Replace(p.ParameterName, p.Value.ToString)
                End Select
            Next
            Return sb.ToString
        End Function

        ''' <summary>
        ''' Detect whether two connection strings represent point to the same UDBS instance.
        ''' User name, password, and other arguments like the operation timeout, etc. are ignored in the
        ''' check. Only the server and database name matter.
        ''' </summary>
        ''' <param name="connectionStringA">One of the connection strings to compage.</param>
        ''' <param name="connectionStringB">The other one.</param>
        ''' <returns>Whether or not the two strings point to the same database.</returns>
        Private Function AreTheSameUdbsInstance(connectionStringA As String, connectionStringB As String) As Boolean
            If String.IsNullOrEmpty(connectionStringA) And String.IsNullOrEmpty(connectionStringB) Then
                ' Both connection string are null.
                ' No changes.
                Return True
            ElseIf String.IsNullOrEmpty(connectionStringA) Or String.IsNullOrEmpty(connectionStringB) Then
                ' One of them is null, the other one isn't.
                ' The connection string is changing.
                Return False
            ElseIf connectionStringA = connectionStringB Then
                ' Connections string are the same.
                ' No need to parse it and do complex logic.
                ' These two represent the same UDBS instance.
                Return True
            Else
                ' Connection string are 
                ' We only care about the server and database.

                Dim connectionA = New OdbcConnectionStringBuilder(connectionStringA)
                Dim connectionB = New OdbcConnectionStringBuilder(connectionStringB)

                If Not connectionA.ContainsKey("Server") Or Not connectionB.ContainsKey("Server") Then
                    ' We're running unit tests against the local DB.
                    ' Don't bother parsing the string.
                    ' The strings don't contain a 'server' argument, so the
                    ' operation would throw an exception.
                    ' Just compare the two strings. That's good-enough.
                    Return connectionStringA = connectionStringB
                End If

                Dim serverA As String = connectionA("Server").ToString()
                Dim serverB As String = connectionB("Server").ToString()
                Dim dbA As String = connectionA("Database").ToString()
                Dim dbB As String = connectionB("Database").ToString()

                Return serverA = serverB And dbA = dbB
            End If
        End Function
    End Module
End Namespace
