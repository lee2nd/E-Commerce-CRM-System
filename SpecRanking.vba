Sub SpecRanking()

    Dim DeliverySheet, VisualizationSheet As Worksheet
    Set DeliverySheet = ThisWorkbook.Sheets("出庫")
    Set VisualizationSheet = ThisWorkbook.Sheets("圖表")
    
        With VisualizationSheet
            
            .Activate
            MyPos = InStr(1, .Range("D32"), ")", 1)
            ItemName = Right(.Range("D32"), Len(.Range("D32")) - MyPos)
            
            '清空原始資料
            VisualizationSheetLastRow = .Range("U1048576").End(xlUp).Row
            If VisualizationSheetLastRow > 26 Then
                .Range("U27:V" & VisualizationSheetLastRow).ClearContents
            End If
            
            DeliverySheetLastRow = DeliverySheet.Range("A1048576").End(xlUp).Row
            DeliverySheet.Range("A1:D" & DeliverySheetLastRow).AutoFilter Field:=2, Criteria1:=ItemName
            
            DeliverySheetLastRow = DeliverySheet.Range("A1048576").End(xlUp).Row
            DeliverySheet.Range("C2:C" & DeliverySheetLastRow).Copy _
            Destination:=.Range("U27")
            
            '移除重複規格
            VisualizationSheetLastRow = .Range("U1048576").End(xlUp).Row
            On Error Resume Next
            .Range("U27:U" & VisualizationSheetLastRow).RemoveDuplicates Columns:=1, Header:=no
            On Error GoTo 0
            
            '找出該商品該規格銷售量+強制將銷售量轉成數字
            VisualizationSheetLastRow = .Range("U1048576").End(xlUp).Row
            
            For i = 27 To VisualizationSheetLastRow
                .Cells(i, 22).Formula = "=SUMPRODUCT((出庫!B2:B" & DeliverySheetLastRow & "=""" & ItemName & """)*(出庫!C2:C" & DeliverySheetLastRow & "=""" & .Range("U" & i) & """),出庫!D2:D" & DeliverySheetLastRow & ")"
                .Range("V" & i) = CInt(.Range("V" & i))
            Next i

            '依銷售量排序
            .Range("U27:V" & VisualizationSheetLastRow).Sort Key1:=.Range("V27"), Order1:=xlDescending, Header:=no
            
            With .Range("U27:U" & VisualizationSheetLastRow)
                
                With .Font
                    .Color = -1003520
                    .TintAndShade = 0
                End With
                
            End With
            
            With .Range("V27:V" & VisualizationSheetLastRow)
            
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                
                With .Font
                    .Color = -1003520
                    .TintAndShade = 0
                End With
                
            End With
            
            '解除出庫的篩選
            DeliverySheet.ShowAllData
                    
        End With
        
End Sub
