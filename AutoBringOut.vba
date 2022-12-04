Sub AutoBringOut()

    Dim CompareSheet, StorageSheet As Worksheet
    Set CompareSheet = ThisWorkbook.Sheets("對照表")
    Set StorageSheet = ThisWorkbook.Sheets("入庫(U)")
    
    With CompareSheet
    
        CompareSheetLastRow = .Range("A1048576").End(xlUp).Row
        StorageSheetLastRow = StorageSheet.Range("A1048576").End(xlUp).Row
        
        For i = 2 To CompareSheetLastRow
            
            If .Range("E" & i) <> "" Then
                On Error Resume Next
                .Range("D" & i) = WorksheetFunction.VLookup(.Range("E" & i), StorageSheet.Range("A2:D" & StorageSheetLastRow), 3, False)
                .Range("F" & i) = WorksheetFunction.VLookup(.Range("E" & i), StorageSheet.Range("A2:D" & StorageSheetLastRow), 4, False)
                On Error GoTo 0
            Else
                '若沒有 vlookup 到對應的入庫品名就用 orders 品名，並標上紅色
                .Range("E" & i) = .Range("A" & i)
                .Range("E" & i).Font.ColorIndex = 3
                .Range("D" & i) = "TBD"
                .Range("D" & i).Font.ColorIndex = 3
                .Range("F" & i) = "A"
                .Range("F" & i).Font.ColorIndex = 3
            End If
        
        Next i

    End With
    
End Sub
