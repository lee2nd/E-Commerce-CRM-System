Sub SalesRanking()

    Dim DeliverySheet, VisualizationSheet As Worksheet
    Set DeliverySheet = ThisWorkbook.Sheets("出庫")
    Set VisualizationSheet = ThisWorkbook.Sheets("圖表")
    
        '清空原始資料
        VisualizationSheetLastRow = VisualizationSheet.Range("A1048576").End(xlUp).Row
        If VisualizationSheetLastRow > 44 Then
            VisualizationSheet.Range("A45:B" & VisualizationSheetLastRow).ClearContents
        End If
        
        With DeliverySheet
            .Activate
            DeliverySheetLastRow = .Range("A1048576").End(xlUp).Row
            
            For i = 2 To DeliverySheetLastRow
                VisualizationSheet.Range("A" & i + 43) = "(" & .Range("A" & i) & ")" & .Range("B" & i)
            Next i
            
            VisualizationSheet.Range("A45:A" & DeliverySheetLastRow + 43).RemoveDuplicates Columns:=1, Header:=no
            
        End With
        
        With VisualizationSheet
            .Activate
            VisualizationSheetLastRow = .Range("A1048576").End(xlUp).Row
            
            For i = 45 To VisualizationSheetLastRow
                MyPos = InStr(1, .Range("A" & i), ")", 1)
                ItemName = Right(.Range("A" & i), Len(.Range("A" & i)) - MyPos)
                .Range("B" & i).Formula = "=SUMPRODUCT((出庫!B2:B" & DeliverySheetLastRow & "=""" & ItemName & """)*1,出庫!F2:F" & DeliverySheetLastRow & ")"
                .Range("B" & i).NumberFormatLocal = "$#,##0_);[紅色]($#,##0)"
            Next i
            
            For i = VisualizationSheetLastRow To 45 Step -1
                If .Range("A" & i) Like "*TBD*" Then
                    .Range("A" & i & ":B" & i).Select
                    Selection.Delete Shift:=xlUp
                End If
            Next i
            
            VisualizationSheetLastRow = .Range("A1048576").End(xlUp).Row
            
            .Range("A45:B" & VisualizationSheetLastRow).Sort Key1:=.Range("B45"), Order1:=xlDescending, Header:=no
            
            With .Range("A45:A" & VisualizationSheetLastRow)
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            With .Range("B45:B" & VisualizationSheetLastRow)
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
        
        End With

End Sub
