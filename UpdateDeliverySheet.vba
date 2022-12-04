Sub UpdateDeliverySheet()

    Dim StorageUSheet, CompareSheet, DeliverySheet, Shopee_sheet, Yahoo_sheet, Ruten_sheet As Worksheet
    Set CompareSheet = ThisWorkbook.Sheets("對照表")
    Set StorageUSheet = ThisWorkbook.Sheets("入庫(U)")
    Set DeliverySheet = ThisWorkbook.Sheets("出庫")
    Set Shopee_sheet = ThisWorkbook.Sheets("蝦皮orders")
    Set Yahoo_sheet = ThisWorkbook.Sheets("雅虎orders")
    Set Ruten_sheet = ThisWorkbook.Sheets("露天orders")

    CompareSheetLastRow = CompareSheet.Range("A1048576").End(xlUp).Row
    StorageUSheetLastRow = StorageUSheet.Range("A1048576").End(xlUp).Row
    DeliverySheetLastRow0 = DeliverySheet.Range("A1048576").End(xlUp).Row
    
    With Shopee_sheet
    
        ShopeeLastRow = .Range("A1048576").End(xlUp).Row
        
            For i = 2 To ShopeeLastRow
            
                If .Cells(i, 3) = Empty And .Cells(i, 4) = Empty Then
                
                    DeliverySheetLastRow = DeliverySheet.Range("A1048576").End(xlUp).Row
                    DeliverySheet.Range("J" & DeliverySheetLastRow + 1) = .Cells(i, 22) & "[" & .Cells(i, 23) & "]"
                    If DeliverySheet.Range("J" & DeliverySheetLastRow + 1) Like "*~*" Then
                        DeliverySheet.Range("K" & DeliverySheetLastRow + 1) = WorksheetFunction.VLookup(WorksheetFunction.Substitute(DeliverySheet.Range("J" & DeliverySheetLastRow + 1), "~", "~~"), CompareSheet.Range("A2:E" & CompareSheetLastRow), 5, False)
                    Else
                        DeliverySheet.Range("K" & DeliverySheetLastRow + 1) = WorksheetFunction.VLookup(DeliverySheet.Range("J" & DeliverySheetLastRow + 1), CompareSheet.Range("A2:E" & CompareSheetLastRow), 5, False)
                    End If
                    ProductName = Split(DeliverySheet.Range("K" & DeliverySheetLastRow + 1), "[")
                    DeliverySheet.Range("B" & DeliverySheetLastRow + 1) = ProductName(0)
                    
                    Specification = Replace(ProductName(1), "]", "")
                    DeliverySheet.Range("C" & DeliverySheetLastRow + 1) = Specification
                    
                    On Error Resume Next
                    DeliverySheet.Range("I" & DeliverySheetLastRow + 1) = WorksheetFunction.VLookup(DeliverySheet.Range("K" & DeliverySheetLastRow + 1), StorageUSheet.Range("A2:D" & StorageUSheetLastRow), 4, False)
                    DeliverySheet.Range("A" & DeliverySheetLastRow + 1) = WorksheetFunction.VLookup(DeliverySheet.Range("K" & DeliverySheetLastRow + 1), StorageUSheet.Range("A2:D" & StorageUSheetLastRow), 3, False)
                    On Error GoTo 0

                    If DeliverySheet.Range("I" & DeliverySheetLastRow + 1) = "" Then
                        DeliverySheet.Range("I" & DeliverySheetLastRow + 1) = "A"
                        DeliverySheet.Range("I" & DeliverySheetLastRow + 1).Font.ColorIndex = 3
                        DeliverySheet.Range("A" & DeliverySheetLastRow + 1) = "TBD"
                        DeliverySheet.Range("A" & DeliverySheetLastRow + 1).Font.ColorIndex = 3
                    End If
                    
                    DeliverySheet.Range("D" & DeliverySheetLastRow + 1) = .Cells(i, 28)
                    DeliverySheet.Range("H" & DeliverySheetLastRow + 1) = "蝦皮"
                    DeliverySheet.Range("H" & DeliverySheetLastRow + 1).Font.ColorIndex = 46
                    
                    If Not IsEmpty(.Cells(i, 25)) Then
                        DeliverySheet.Range("E" & DeliverySheetLastRow + 1) = .Cells(i, 25)
                    ElseIf (IsEmpty(.Cells(i, 25))) And (Not IsEmpty(.Cells(i, 24))) Then
                        DeliverySheet.Range("E" & DeliverySheetLastRow + 1) = .Cells(i, 24)
                    Else
                        DeliverySheet.Range("E" & DeliverySheetLastRow + 1) = "查無金額"
                    End If
                    
                    DeliverySheet.Range("F" & DeliverySheetLastRow + 1) = DeliverySheet.Range("D" & DeliverySheetLastRow + 1) * DeliverySheet.Range("E" & DeliverySheetLastRow + 1)
                    DeliverySheet.Range("G" & DeliverySheetLastRow + 1) = Left(.Cells(i, 6), 10)
                    DeliverySheet.Range("G" & DeliverySheetLastRow + 1).NumberFormat = "yyyy年m月d日"

                End If
            Next i
        
    End With
    
    With Yahoo_sheet
    
        YahooLastRow = .Range("A1048576").End(xlUp).Row
        
            For i = 2 To YahooLastRow
            
                If Not .Cells(i, 38) Like "*已取消*" Then
                
                    DeliverySheetLastRow = DeliverySheet.Range("A1048576").End(xlUp).Row
                    DeliverySheet.Range("J" & DeliverySheetLastRow + 1) = .Cells(i, 6) & "[" & .Cells(i, 10) & "," & .Cells(i, 11) & "]"
                    DeliverySheet.Range("K" & DeliverySheetLastRow + 1) = WorksheetFunction.VLookup(DeliverySheet.Range("J" & DeliverySheetLastRow + 1), CompareSheet.Range("A2:E" & CompareSheetLastRow), 5, False)
                    
                    ProductName = Split(DeliverySheet.Range("K" & DeliverySheetLastRow + 1), "[")
                    DeliverySheet.Range("B" & DeliverySheetLastRow + 1) = ProductName(0)
                    
                    Specification = Replace(ProductName(1), "]", "")
                    DeliverySheet.Range("C" & DeliverySheetLastRow + 1) = Specification
                    
                    On Error Resume Next
                    DeliverySheet.Range("I" & DeliverySheetLastRow + 1) = WorksheetFunction.VLookup(DeliverySheet.Range("K" & DeliverySheetLastRow + 1), StorageUSheet.Range("A2:D" & StorageUSheetLastRow), 4, False)
                    DeliverySheet.Range("A" & DeliverySheetLastRow + 1) = WorksheetFunction.VLookup(DeliverySheet.Range("K" & DeliverySheetLastRow + 1), StorageUSheet.Range("A2:D" & StorageUSheetLastRow), 3, False)
                    On Error GoTo 0

                    If DeliverySheet.Range("I" & DeliverySheetLastRow + 1) = "" Then
                        DeliverySheet.Range("I" & DeliverySheetLastRow + 1) = "A"
                        DeliverySheet.Range("I" & DeliverySheetLastRow + 1).Font.ColorIndex = 3
                        DeliverySheet.Range("A" & DeliverySheetLastRow + 1) = "TBD"
                        DeliverySheet.Range("A" & DeliverySheetLastRow + 1).Font.ColorIndex = 3
                    End If
                    
                    DeliverySheet.Range("D" & DeliverySheetLastRow + 1) = .Cells(i, 15)
                    DeliverySheet.Range("H" & DeliverySheetLastRow + 1) = "Y拍"
                    DeliverySheet.Range("H" & DeliverySheetLastRow + 1).Font.ColorIndex = 29                  
                    DeliverySheet.Range("E" & DeliverySheetLastRow + 1) = .Cells(i, 14)                    
                    DeliverySheet.Range("F" & DeliverySheetLastRow + 1) = DeliverySheet.Range("D" & DeliverySheetLastRow + 1) * DeliverySheet.Range("E" & DeliverySheetLastRow + 1)
                    MyPos = InStr(1, .Cells(i, 1), " ", 1)
                    DeliverySheet.Range("G" & DeliverySheetLastRow + 1) = Left(.Cells(i, 1), MyPos - 1)
                    DeliverySheet.Range("G" & DeliverySheetLastRow + 1).NumberFormat = "yyyy年m月d日"

                End If
            Next i
        
    End With
    
    With Ruten_sheet

        RutenLastRow = .Range("A1048576").End(xlUp).Row

            For i = 2 To RutenLastRow

                If Not .Cells(i, 17) Like "*已領退貨*" Then

                    DeliverySheetLastRow = DeliverySheet.Range("A1048576").End(xlUp).Row
                    DeliverySheet.Range("J" & DeliverySheetLastRow + 1) = .Cells(i, 6) & "[" & .Cells(i, 7) & "," & .Cells(i, 8) & "]"
                    DeliverySheet.Range("K" & DeliverySheetLastRow + 1) = WorksheetFunction.VLookup(DeliverySheet.Range("J" & DeliverySheetLastRow + 1), CompareSheet.Range("A2:E" & CompareSheetLastRow), 5, False)

                    ProductName = Split(DeliverySheet.Range("K" & DeliverySheetLastRow + 1), "[")
                    DeliverySheet.Range("B" & DeliverySheetLastRow + 1) = ProductName(0)

                    Specification = Replace(ProductName(1), "]", "")
                    DeliverySheet.Range("C" & DeliverySheetLastRow + 1) = Specification
                    
                    On Error Resume Next
                    DeliverySheet.Range("I" & DeliverySheetLastRow + 1) = WorksheetFunction.VLookup(DeliverySheet.Range("K" & DeliverySheetLastRow + 1), StorageUSheet.Range("A2:D" & StorageUSheetLastRow), 4, False)
                    DeliverySheet.Range("A" & DeliverySheetLastRow + 1) = WorksheetFunction.VLookup(DeliverySheet.Range("K" & DeliverySheetLastRow + 1), StorageUSheet.Range("A2:D" & StorageUSheetLastRow), 3, False)
                    On Error GoTo 0

                    If DeliverySheet.Range("I" & DeliverySheetLastRow + 1) = "" Then
                        DeliverySheet.Range("I" & DeliverySheetLastRow + 1) = "A"
                        DeliverySheet.Range("I" & DeliverySheetLastRow + 1).Font.ColorIndex = 3
                        DeliverySheet.Range("A" & DeliverySheetLastRow + 1) = "TBD"
                        DeliverySheet.Range("A" & DeliverySheetLastRow + 1).Font.ColorIndex = 3
                    End If

                    DeliverySheet.Range("D" & DeliverySheetLastRow + 1) = .Cells(i, 10)
                    DeliverySheet.Range("H" & DeliverySheetLastRow + 1) = "露天"
                    DeliverySheet.Range("H" & DeliverySheetLastRow + 1).Font.ColorIndex = 50
                    DeliverySheet.Range("E" & DeliverySheetLastRow + 1) = .Cells(i, 11)
                    DeliverySheet.Range("F" & DeliverySheetLastRow + 1) = DeliverySheet.Range("D" & DeliverySheetLastRow + 1) * DeliverySheet.Range("E" & DeliverySheetLastRow + 1)
                    DeliverySheet.Range("G" & DeliverySheetLastRow + 1) = .Cells(i, 1)
                    DeliverySheet.Range("G" & DeliverySheetLastRow + 1).NumberFormat = "yyyy年m月d日"

                End If
            Next i

    End With

    With DeliverySheet
    
        '清除參考品名
        .Range("J:K").ClearContents
        
        '刪除重複資料
        .UsedRange.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9), Header:=xlYes
        
        '調整字體
        .Cells.Font.Size = 11
        .Cells.Font.Name = "微軟正黑體"
        .Cells.VerticalAlignment = xlVAlignCenter
        .Cells.HorizontalAlignment = xlHAlignLeft

        '自動調整欄寬
        .Columns("A:I").AutoFit
        
        DeliverySheetLastRow1 = DeliverySheet.Range("A1048576").End(xlUp).Row

        '依日期排序
        With DeliverySheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("G1:G" & DeliverySheetLastRow1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A1:I" & DeliverySheetLastRow1)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        With ThisWorkbook.Sheets("Control Panel")
            .Range("G18") = DeliverySheetLastRow1 - DeliverySheetLastRow0
            .Range("G18").Font.Size = 12
            .Range("G18").Font.Name = "微軟正黑體"
            .Range("G18").VerticalAlignment = xlVAlignCenter
            .Range("G18").HorizontalAlignment = xlCenter
        End With
    
    End With
    
    ThisWorkbook.Sheets("Control Panel").Activate
    MsgBox "Complete!"

End Sub
