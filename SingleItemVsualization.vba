Sub SingleItemVsualization()

    Dim DeliverySheet, VisualizationSheet As Worksheet
    Set DeliverySheet = ThisWorkbook.Sheets("出庫")
    Set VisualizationSheet = ThisWorkbook.Sheets("圖表")
    Set InventoryDetailsSheet = ThisWorkbook.Sheets("庫存明細")
    
    With VisualizationSheet
    
        MyPos = InStr(1, .Range("D32"), ")", 1)
        ItemName = Right(.Range("D32"), Len(.Range("D32")) - MyPos)
        .Range("B28:C39").ClearContents
        ItemAvgCost = WorksheetFunction.VLookup(ItemName, InventoryDetailsSheet.Range("B:J"), 8, False)

        DeliverySheetLastRow = DeliverySheet.Range("A1048576").End(xlUp).Row
        
        '計算總銷售額
        For i = 1 To 12
            .Cells(i + 27, 2).Formula = "=SUMPRODUCT((MONTH(出庫!G2:G" & DeliverySheetLastRow & ")=" & i & ")*(出庫!B2:B" & DeliverySheetLastRow & "=""" & ItemName & """),出庫!F2:F" & DeliverySheetLastRow & ")"
        Next i
        
        '計算總淨利額
        For i = 1 To 12
            .Cells(i + 27, 100).Formula = "=SUMPRODUCT((MONTH(出庫!G2:G" & DeliverySheetLastRow & ")=" & i & ")*(出庫!B2:B" & DeliverySheetLastRow & "=""" & ItemName & """),出庫!D2:D" & DeliverySheetLastRow & ")"
            .Cells(i + 27, 3).Formula = "=SUMPRODUCT((MONTH(出庫!G2:G" & DeliverySheetLastRow & ")=" & i & ")*(出庫!B2:B" & DeliverySheetLastRow & "=""" & ItemName & """),出庫!F2:F" & DeliverySheetLastRow & ")-CV" & i + 27 & "*" & ItemAvgCost
        Next i
        
    End With
End Sub
