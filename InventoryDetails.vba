Sub InventoryDetails()

    Dim InventoryDetailsSheet, StorageSheet, DeliverySheet As Worksheet
    Set InventoryDetailsSheet = ThisWorkbook.Sheets("庫存明細")
    Set StorageSheet = ThisWorkbook.Sheets("入庫")
    Set DeliverySheet = ThisWorkbook.Sheets("出庫")
    
    StorageSheetLastRow = StorageSheet.Range("A1048576").End(xlUp).Row
    ThisWorkbook.Sheets("入庫").Range("A2:C" & StorageSheetLastRow).Copy
    InventoryDetailsSheetLastRow = InventoryDetailsSheet.Range("A1048576").End(xlUp).Row
    InventoryDetailsSheetLastRow0 = InventoryDetailsSheet.Range("A1048576").End(xlUp).Row
    
    With InventoryDetailsSheet
    
        .Range("A" & InventoryDetailsSheetLastRow + 1).PasteSpecial Paste:=xlPasteValues
        '刪除重複資料
        .UsedRange.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
        InventoryDetailsSheetLastRow = .Range("A1048576").End(xlUp).Row
        
        For i = 3 To InventoryDetailsSheetLastRow
        
            AA = .Cells(i, 1)
            BB = .Cells(i, 2)
            CC = .Cells(i, 3)

            With StorageSheet
                .UsedRange.AutoFilter Field:=1, Criteria1:=AA
                .UsedRange.AutoFilter Field:=2, Criteria1:=BB
                .UsedRange.AutoFilter Field:=3, Criteria1:=CC
                TotalStorageNum = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
                TotalStorageAmount = Application.WorksheetFunction.Sum(.Range("F:F").SpecialCells(xlCellTypeVisible))
                AverageCost = Application.WorksheetFunction.Average(.Range("E:E").SpecialCells(xlCellTypeVisible))
            End With

            With DeliverySheet
                .UsedRange.AutoFilter Field:=1, Criteria1:=AA
                .UsedRange.AutoFilter Field:=2, Criteria1:=BB
                .UsedRange.AutoFilter Field:=3, Criteria1:=CC
                TotalDeliveryNum = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
                TotalDeliveryAmount = Application.WorksheetFunction.Sum(.Range("F:F").SpecialCells(xlCellTypeVisible))
            End With
            
            '進貨統計數量
            .Cells(i, 4) = TotalStorageNum
            '進貨統計合計
            .Cells(i, 5) = TotalStorageAmount
            '銷售統計數量
            .Cells(i, 6) = TotalDeliveryNum
            '銷售統計合計
            .Cells(i, 7) = TotalDeliveryAmount
            '現有庫存
            .Cells(i, 8) = .Cells(i, 4) - .Cells(i, 6)
            '平均成本(庫存明細)
            .Cells(i, 9) = .Cells(i, 5) / .Cells(i, 4)
            '平均成本(日報表)
            .Cells(i, 10) = AverageCost
            
            '若平均成本(庫存明細) 和 平均成本(日報表)數字不同則標黃底色
            If .Cells(i, 9) <> .Cells(i, 10) Then
            
                .Cells(i, 9).Interior.ColorIndex = 36
                .Cells(i, 10).Interior.ColorIndex = 36
                
            End If
            
        Next i

        InventoryDetailsSheetLastRow1 = InventoryDetailsSheet.Range("A1048576").End(xlUp).Row
        
        '調整字體
        .Cells.Font.Size = 11
        .Cells.Font.Name = "微軟正黑體"
        .Range("A3:J" & InventoryDetailsSheetLastRow1).VerticalAlignment = xlVAlignCenter
        .Range("A3:J" & InventoryDetailsSheetLastRow1).HorizontalAlignment = xlHAlignLeft
        
        With ThisWorkbook.Sheets("Control Panel")
            .Range("G23") = InventoryDetailsSheetLastRow1 - InventoryDetailsSheetLastRow0
            .Range("G23").Font.Size = 12
            .Range("G23").Font.Name = "微軟正黑體"
            .Range("G23").VerticalAlignment = xlVAlignCenter
            .Range("G23").HorizontalAlignment = xlCenter
        End With
        
        StorageSheet.AutoFilterMode = False
        DeliverySheet.AutoFilterMode = False
        
        ThisWorkbook.Sheets("Control Panel").Activate
        MsgBox "Complete!"
    
    End With

End Sub
