Option Explicit On
Option Compare Text
Option Infer On
Option Strict On

Namespace MasterInterface

    ''' <summary>
    ''' This helper class avoid having to iterate through an entire result
    ''' set in the search of a single line over and over again.
    ''' Instead, it indexes the row and lets the <see cref="Dictionary(Of TKey, TValue)"/>
    ''' class do the heavy lifting.
    ''' </summary>
    ''' <typeparam name="T">The type of the index key.</typeparam>
    Friend Class IndexedResultSet(Of T)
        Private mDictionary As Dictionary(Of T, DataRow) = New Dictionary(Of T, DataRow)

        ''' <summary>
        ''' Constructor. Creates a new index.
        ''' </summary>
        ''' <param name="table">The table which rows we need to index.</param>
        ''' <param name="indexColumnName">The name of the colum containing the value to index by.</param>
        Public Sub New(table As DataTable, indexColumnName As String)
            Index(table, indexColumnName)
        End Sub

        ''' <summary>
        ''' Indexes the table using the column with the specified name.
        ''' </summary>
        ''' <param name="table">The table to index.</param>
        ''' <param name="indexColumnName">The name of the column containing the key to index by.</param>
        ''' <remarks>
        ''' Performance timing was measured, and this takes less than 1 msec.
        ''' </remarks>
        Private Sub Index(table As DataTable, indexColumnName As String)
            For Each aRow As DataRow In table.Rows
                Dim indexValue As T = CType(aRow(indexColumnName), T)
                If (mDictionary.ContainsKey(indexValue)) Then
                    logger.Warn($"Problem indexing: there is already a row at index {indexValue}")
                    Continue For
                End If
                mDictionary.Add(indexValue, aRow)
            Next
        End Sub

        ''' <summary>
        ''' Find the row at a given index.
        ''' </summary>
        ''' <param name="index">The index of the row we are looking for.</param>
        ''' <returns>The row at that index. Null if there is no such row.</returns>
        Public Function FindRow(index As T) As DataRow
            Dim result As DataRow = Nothing
            mDictionary.TryGetValue(index, result)
            Return result
        End Function
    End Class
End Namespace
