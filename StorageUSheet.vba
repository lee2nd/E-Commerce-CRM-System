Sub UpdateStorageUSheet()
    
    ' 更新入庫
    Dim StorageUSheet As Worksheet
    Set StorageUSheet = ThisWorkbook.Sheets("入庫(U)")
    
    With StorageUSheet
        StorageUSheetLastRow0 = .Range("A1048576").End(xlUp).Row
        StorageLastRow = ThisWorkbook.Sheets("入庫").Range("A1048576").End(xlUp).Row
        
        '名稱+規格
        For i = 2 To StorageLastRow
                ThisWorkbook.Sheets("入庫").Cells(i, 9) = ThisWorkbook.Sheets("入庫").Cells(i, 2) & "[" & ThisWorkbook.Sheets("入庫").Cells(i, 3) & "]"
        Next i
        
        '從入庫抓資料
        ThisWorkbook.Sheets("入庫").Range("A2:A" & StorageLastRow).Copy
        .Range("C" & StorageUSheetLastRow0 + 1).PasteSpecial Paste:=xlPasteValues
        ThisWorkbook.Sheets("入庫").Range("I2:I" & StorageLastRow).Copy
        .Range("A" & StorageUSheetLastRow0 + 1).PasteSpecial Paste:=xlPasteValues
        ThisWorkbook.Sheets("入庫").Range("H2:H" & StorageLastRow).Copy
        .Range("D" & StorageUSheetLastRow0 + 1).PasteSpecial Paste:=xlPasteValues
        
        '刪除重複資料
        .UsedRange.RemoveDuplicates Columns:=Array(1, 3, 4), Header:=xlYes
        
        ThisWorkbook.Sheets("入庫").Range("I:I").ClearContents
        
        '調整字體
        .Cells.Font.Size = 12
        .Cells.Font.Name = "微軟正黑體"
        
        '自動調整欄寬
        .Columns("A:D").AutoFit
        
        StorageUSheetLastRow1 = .Range("A1048576").End(xlUp).Row
        
        '依貨號排序
        With StorageUSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("C1:C" & StorageUSheetLastRow1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A1:D" & StorageUSheetLastRow1)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
            
        With ThisWorkbook.Sheets("Control Panel")
            .Range("G8") = StorageUSheetLastRow1 - StorageUSheetLastRow0
            .Range("G8").Font.Size = 12
            .Range("G8").Font.Name = "微軟正黑體"
            .Range("G8").VerticalAlignment = xlVAlignCenter
            .Range("G8").HorizontalAlignment = xlCenter
        End With
        
        MsgBox "Complete!"

    End With

End Sub
