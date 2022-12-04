Sub UpdateCompareSheet()

    Dim CompareSheet, Shopee_sheet, Yahoo_sheet, Ruten_sheet As Worksheet
    Set CompareSheet = ThisWorkbook.Sheets("對照表")
    Set Shopee_sheet = ThisWorkbook.Sheets("蝦皮orders")
    Set Yahoo_sheet = ThisWorkbook.Sheets("雅虎orders")
    Set Ruten_sheet = ThisWorkbook.Sheets("露天orders")
    
    CompareSheetLastRow = CompareSheet.Range("A1048576").End(xlUp).Row
    
    With Shopee_sheet
    
        ShopeeLastRow = .Range("A1048576").End(xlUp).Row
        
        If ShopeeLastRow <> 1 Then
        
            For i = 2 To ShopeeLastRow
                .Cells(i, 49) = .Cells(i, 22) & "[" & .Cells(i, 23) & "]"
            Next i
            
        .Range("AW2:AW" & ShopeeLastRow).Copy
        
        With CompareSheet
            .Activate
            .Range("A" & CompareSheetLastRow + 1).PasteSpecial Paste:=xlPasteValues
            .Range(Cells(CompareSheetLastRow + 1, 2), Cells(CompareSheetLastRow + ShopeeLastRow - 1, 2)).Value = "蝦皮"
            .Range(Cells(CompareSheetLastRow + 1, 2), Cells(CompareSheetLastRow + ShopeeLastRow - 1, 2)).Font.ColorIndex = 46
        End With
            
            .Range("AW:AW").ClearContents
            
        End If
        
    End With
    
    CompareSheetLastRow = CompareSheet.Range("A1048576").End(xlUp).Row
    
    With Yahoo_sheet
    
        YahooLastRow = .Range("A1048576").End(xlUp).Row
        
        If YahooLastRow <> 1 Then
        
                For i = 2 To YahooLastRow
                    .Cells(i, 42) = .Cells(i, 6) & "[" & .Cells(i, 10) & "," & .Cells(i, 11) & "]"
                Next i
                
            .Range("AP2:AP" & YahooLastRow).Copy
            
            With CompareSheet
                .Activate
                .Range("A" & CompareSheetLastRow + 1).PasteSpecial Paste:=xlPasteValues
                .Range(Cells(CompareSheetLastRow + 1, 2), Cells(CompareSheetLastRow + YahooLastRow - 1, 2)).Value = "雅虎"
                .Range(Cells(CompareSheetLastRow + 1, 2), Cells(CompareSheetLastRow + YahooLastRow - 1, 2)).Font.ColorIndex = 29
            End With
            
                .Range("AP:AP").ClearContents
                
        End If
        
    End With
    
    CompareSheetLastRow = CompareSheet.Range("A1048576").End(xlUp).Row
    
    With Ruten_sheet
    
        RutenLastRow = .Range("A1048576").End(xlUp).Row
        
        If RutenLastRow <> 1 Then
        
            For i = 2 To RutenLastRow
                .Cells(i, 23) = .Cells(i, 6) & "[" & .Cells(i, 7) & "," & .Cells(i, 8) & "]"
            Next i
            
        .Range("W2:W" & RutenLastRow).Copy
        
        With CompareSheet
            .Activate
            .Range("A" & CompareSheetLastRow + 1).PasteSpecial Paste:=xlPasteValues
            .Range(Cells(CompareSheetLastRow + 1, 2), Cells(CompareSheetLastRow + RutenLastRow - 1, 2)).Value = "露天"
            .Range(Cells(CompareSheetLastRow + 1, 2), Cells(CompareSheetLastRow + RutenLastRow - 1, 2)).Font.ColorIndex = 50
        End With
        
            .Range("W:W").ClearContents
            
        End If
        
    End With
    
    With CompareSheet
    
        '刪除重複資料
        .UsedRange.RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
        
        '調整字體
        .Cells.Font.Size = 12
        .Cells.Font.Name = "微軟正黑體"
        
        '自動調整欄寬
        .Columns("A:F").AutoFit
        .Columns("D:D").ColumnWidth = 8
        .Columns("E:E").ColumnWidth = 30
        
        '跳到左上角
        .Range("A1").Activate
        
        CompareSheetLastRow = .Range("A1048576").End(xlUp).Row
        
        '製作入庫(U)的下拉式選單
        .Range("E2:E" & CompareSheetLastRow).Select
        StorageLastRow = ThisWorkbook.Sheets("入庫(U)").Range("A1048576").End(xlUp).Row
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="='入庫(U)'!$A$2:$A$" & StorageLastRow
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        
        '計算未配對數量
        ThisWorkbook.Sheets("Control Panel").Range("G13") = Application.WorksheetFunction.CountBlank(.Range("E2:E" & CompareSheetLastRow))
        ThisWorkbook.Sheets("Control Panel").Range("G13").Font.Size = 12
        ThisWorkbook.Sheets("Control Panel").Range("G13").Font.Name = "微軟正黑體"
        ThisWorkbook.Sheets("Control Panel").Range("G13").VerticalAlignment = xlVAlignCenter
        ThisWorkbook.Sheets("Control Panel").Range("G13").HorizontalAlignment = xlCenter
    
    End With
    
    ThisWorkbook.Sheets("Control Panel").Activate
    MsgBox "Complete!"
    
End Sub
