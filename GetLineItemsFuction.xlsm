Function GetLineItems(table As Object) As Collection
    
    Dim TableCollection As Collection
    Dim RowDict As Dictionary
    Dim KeyNames As Variant
    
    Set TableCollection = New Collection
    KeyNames = Array("productCode", "description", "qty", "model", "serial", "startDate", "endDate")
    
    For i = 1 To table.Rows.Length - 1

        Set tablerow = table.Rows(i)
        Set RowDict = New Dictionary
        For ColumnNo = 0 To tablerow.cells.Length - 2
            'Debug.Print i, ColumnNo
            RowDict.Add KeyNames(ColumnNo), tablerow.cells(ColumnNo).innerText
        Next
            TableCollection.Add RowDict
            Set RowDict = Nothing
    Next
    Set GetLineItems = TableCollection
End Function
