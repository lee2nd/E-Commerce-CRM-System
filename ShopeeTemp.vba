Sub ShopeeTemp()

    Application.DisplayAlerts = False
    
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "Shopee_temp"
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "Shopee_Ratio"
    Dim StorageSheet, CompareSheet, ShopeetempSheet, Shopee_sheet, ShopeeRatioSheet, DaySheetA, DaySheetB As Worksheet
    Set ShopeetempSheet = ThisWorkbook.Sheets("Shopee_temp")
    Set ShopeeRatioSheet = ThisWorkbook.Sheets("Shopee_Ratio")
    Set CompareSheet = ThisWorkbook.Sheets("對照表")
    Set StorageSheet = ThisWorkbook.Sheets("入庫")
    Set Shopee_sheet = ThisWorkbook.Sheets("蝦皮orders")
    Set DaySheetA = ThisWorkbook.Sheets("日報表A")
    Set DaySheetB = ThisWorkbook.Sheets("日報表B")

    ShopeeSheetLastRow = Shopee_sheet.Range("A1048576").End(xlUp).Row
    CompareSheetLastRow = CompareSheet.Range("A1048576").End(xlUp).Row

    With ShopeeRatioSheet
        .Range("A1") = "訂單編號"
        .Range("B1") = "RatioA"
        .Range("C1") = "RatioB"
    End With
    
    With ShopeetempSheet
        .Range("A1") = "訂單編號"
        .Range("B1") = "商品名稱"
        .Range("C1") = "貨號"
        .Range("D1") = "營業額"
        .Range("E1") = "成本"
        .Range("F1") = "出貨人"
        .Range("G1") = "出貨狀態"
        .Range("H1") = "日期"
        .Range("I1") = "賣家折扣卷"
        .Range("J1") = "數量"
        .Range("K1") = "入庫名稱"
        
        For i = 2 To ShopeeSheetLastRow
            
            .Cells(i, 1) = Shopee_sheet.Cells(i, 1)
            .Cells(i, 1).NumberFormat = "0"
            .Cells(i, 2) = Shopee_sheet.Cells(i, 22) & "[" & Shopee_sheet.Cells(i, 23) & "]"
            .Cells(i, 3) = WorksheetFunction.VLookup(.Range("B" & i), CompareSheet.Range("A2:D" & CompareSheetLastRow), 4, False)
            If Shopee_sheet.Cells(i, 25) = "" Then
                .Cells(i, 4) = Shopee_sheet.Cells(i, 24) * Shopee_sheet.Cells(i, 28)
            Else
                .Cells(i, 4) = Shopee_sheet.Cells(i, 25) * Shopee_sheet.Cells(i, 28)
            End If
            
            .Cells(i, 6) = WorksheetFunction.VLookup(.Range("B" & i), CompareSheet.Range("A2:F" & CompareSheetLastRow), 6, False)
            .Cells(i, 8) = Left(Shopee_sheet.Cells(i, 6), 10)
            .Cells(i, 9) = Shopee_sheet.Cells(i, 14)
            .Cells(i, 10) = Shopee_sheet.Cells(i, 28)
            .Cells(i, 11) = WorksheetFunction.VLookup(.Range("B" & i), CompareSheet.Range("A2:E" & CompareSheetLastRow), 5, False)
            
            With StorageSheet
                    
                    Dim dctData1 As Object
                    Set dctData1 = CreateObject("scripting.dictionary")
                    Dim myArray() As Variant, X As Long
                    X = 0
                    
                    StorageSheetLastRow = StorageSheet.Range("A1048576").End(xlUp).Row
                    
                    .Range("I1") = "I1"
                    .Range("J1") = "J1"
                        
                    For j = 2 To StorageSheetLastRow
                        .Range("I" & j) = .Range("B" & j) & "[" & .Range("C" & j) & "]"
                        .Range("J" & j) = .Range("E" & j)
                    Next j
                    
                    '找出大於一個的品名
                    For q = 2 To StorageSheetLastRow
                        If Not dctData1.Exists(.Cells(q, 9)) Then
                            dctData1.Add (.Cells(q, 9)), ""
                        End If
                    Next q
                    
                    For Each a In dctData1
                        If WorksheetFunction.CountIf(.Range("I:I"), a) > 1 Then
                            ReDim Preserve myArray(0 To X)
                            myArray(X) = a
                            X = X + 1
                        End If
                    Next a

                    dctData1.RemoveAll

            End With
            
            If Not IsInArray(.Cells(i, 11), myArray) Then
            
                On Error Resume Next
                .Cells(i, 5) = .Cells(i, 10) * WorksheetFunction.VLookup(.Cells(i, 11), StorageSheet.Range("I2:J" & StorageSheetLastRow), 2, False)
                On Error GoTo 0
            
            Else
                                
                '篩選重複因子
                StorageSheet.UsedRange.AutoFilter Field:=9, Criteria1:=.Cells(i, 11)
                .Cells(i, 5) = .Cells(i, 10) * Application.WorksheetFunction.Average(StorageSheet.Range("J:J").SpecialCells(xlCellTypeVisible))

            End If
            
            On Error Resume Next
                StorageSheet.ShowAllData
            On Error GoTo 0
            
            Erase myArray
            
            If .Cells(i, 5) = "" Then .Cells(i, 5) = 0
            
            If Shopee_sheet.Cells(i, 4) <> "" Then
                .Cells(i, 7) = "!退貨!"
            End If
            
            If Shopee_sheet.Cells(i, 2) Like "*取消*" Then
                .Cells(i, 8) = "日期"
            End If
                            
        Next i
        
        StorageSheet.Range("I:J").ClearContents
        
        '依訂單排序
        With ShopeetempSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A1:A" & ShopeeSheetLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range("F1:F" & ShopeeSheetLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A1:K" & ShopeeSheetLastRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        'Unique 出所有訂單
        Dim d As Object, c As Range, k, tmp As String
        Set d = CreateObject("scripting.dictionary")
        
        Dim dctData As Object
        Set dctData = CreateObject("scripting.dictionary")
        
        For Each c In .Range("A2:A" & ShopeeSheetLastRow)
            tmp = Trim(c.Value)
            If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
        Next c
        
        For Each k In d.keys

                .UsedRange.AutoFilter Field:=1, Criteria1:=k
                
                'Type A
                .UsedRange.AutoFilter Field:=6, Criteria1:="A"
                TotalRevenueA = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
                TotalCostA = Application.WorksheetFunction.Sum(.Range("E:E").SpecialCells(xlCellTypeVisible))
                ShopeetempSheetLastRow = .Range("A1048576").End(xlUp).Row
                
                FilterdCount = Application.WorksheetFunction.Subtotal(3, .Range("B:B")) - 1
                
                OrderNumSet = ""
                NameSet = ""
                If FilterdCount = 0 Then
                        OrderNumSet = ""
                        NameSet = ""
                ElseIf FilterdCount = 1 Then
                        OrderNumSet = .Range("C" & ShopeetempSheetLastRow) & "(" & .Range("J" & ShopeetempSheetLastRow) & ")"
                        NameSet = .Range("K" & ShopeetempSheetLastRow)
                        '2021/2/4
                        If .Range("C" & ShopeetempSheetLastRow) = "TBD" Then
                                .Range("G" & ShopeetempSheetLastRow) = "!未匹配!"
                        End If
                Else
                        For m = ShopeetempSheetLastRow - FilterdCount + 1 To ShopeetempSheetLastRow
                                OrderNumSet = OrderNumSet & ";" & .Range("C" & m) & "(" & .Range("J" & m) & ")"
                                NameSet = NameSet & "," & .Range("K" & m)
                        Next m
                        OrderNumSet = Right(OrderNumSet, Len(OrderNumSet) - 1)
                        NameSet = Right(NameSet, Len(NameSet) - 1)
                        
                        '移除重複品名
                        
                        a = Split(NameSet, ",")
                        For q = 0 To UBound(a, 1)
                            If Not dctData.Exists(Trim(a(q))) Then
                                dctData.Add Trim(a(q)), ""
                            End If
                        Next q
                        
                        NameSet = ""
                        
                        For Each a In dctData
                            NameSet = NameSet & "," & a
                        Next a
                        
                        dctData.RemoveAll
                        
                        NameSet = Right(NameSet, Len(NameSet) - 1)
                        '2021/2/4
                        Set TBDRng = .UsedRange.Find(what:="TBD")
                        
                        If Not TBDRng Is Nothing Then
                            For f = ShopeetempSheetLastRow - FilterdCount + 1 To ShopeetempSheetLastRow
                                    .Range("G" & f) = "!未匹配!"
                            Next f
                        End If
                
                End If
                            
                With DaySheetA
                    DaySheetALastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(DaySheetALastRow + 1, 1) = ShopeetempSheet.Range("H" & ShopeetempSheetLastRow)
                    .Cells(DaySheetALastRow + 1, 1).NumberFormat = "yyyy年m月d日"
                    .Cells(DaySheetALastRow + 1, 1).NumberFormat = "m月d日"
                    .Cells(DaySheetALastRow + 1, 2) = ShopeetempSheet.Range("A" & ShopeetempSheetLastRow)
                    .Cells(DaySheetALastRow + 1, 3) = NameSet
                    .Cells(DaySheetALastRow + 1, 15) = OrderNumSet
                    .Cells(DaySheetALastRow + 1, 4) = TotalRevenueA
                    .Cells(DaySheetALastRow + 1, 17) = ShopeetempSheet.Range("I" & ShopeetempSheetLastRow)
                    .Cells(DaySheetALastRow + 1, 11) = TotalCostA
                    .Cells(DaySheetALastRow + 1, 13) = ShopeetempSheet.Range("G" & ShopeetempSheetLastRow)
                    .Cells(DaySheetALastRow + 1, 13).Font.ColorIndex = 3
                    .Cells(DaySheetALastRow + 1, 14) = "蝦皮"
                    .Cells(DaySheetALastRow + 1, 14).Font.ColorIndex = 46
                End With
                
                'Type B
                .UsedRange.AutoFilter Field:=6, Criteria1:="B"
                TotalRevenueB = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
                TotalCostB = Application.WorksheetFunction.Sum(.Range("E:E").SpecialCells(xlCellTypeVisible))
                ShopeetempSheetLastRow = .Range("A1048576").End(xlUp).Row

                FilterdCount = Application.WorksheetFunction.Subtotal(3, .Range("B:B")) - 1
                
                OrderNumSet = ""
                NameSet = ""
                If FilterdCount = 0 Then
                        OrderNumSet = ""
                        NameSet = ""
                ElseIf FilterdCount = 1 Then
                        OrderNumSet = .Range("C" & ShopeetempSheetLastRow) & "(" & .Range("J" & ShopeetempSheetLastRow) & ")"
                        NameSet = .Range("K" & ShopeetempSheetLastRow)
                        '2021/2/4
                        If .Range("C" & ShopeetempSheetLastRow) = "TBD" Then
                                .Range("G" & ShopeetempSheetLastRow) = "!未匹配!"
                        End If
                Else
                        For m = ShopeetempSheetLastRow - FilterdCount + 1 To ShopeetempSheetLastRow
                                OrderNumSet = OrderNumSet & ";" & .Range("C" & m) & "(" & .Range("J" & m) & ")"
                                NameSet = NameSet & "," & .Range("K" & m)
                        Next m
                        OrderNumSet = Right(OrderNumSet, Len(OrderNumSet) - 1)
                        NameSet = Right(NameSet, Len(NameSet) - 1)
                        
                        '移除重複品名
                        a = Split(NameSet, ",")
                        For q = 0 To UBound(a, 1)
                            If Not dctData.Exists(Trim(a(q))) Then
                                dctData.Add Trim(a(q)), ""
                            End If
                        Next q
                        
                        NameSet = ""
                        
                        For Each a In dctData
                            NameSet = NameSet & "," & a
                        Next a
                        
                        dctData.RemoveAll
                        
                        NameSet = Right(NameSet, Len(NameSet) - 1)
                        '2021/2/4
                        Set TBDRng = .UsedRange.Find(what:="TBD")
                        
                        If Not TBDRng Is Nothing Then
                            For f = ShopeetempSheetLastRow - FilterdCount + 1 To ShopeetempSheetLastRow
                                    .Range("G" & f) = "!未匹配!"
                            Next f
                        End If
                        
                End If
                                
                With DaySheetB
                    DaySheetBLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(DaySheetBLastRow + 1, 1) = ShopeetempSheet.Range("H" & ShopeetempSheetLastRow)
                    .Cells(DaySheetBLastRow + 1, 1).NumberFormat = "yyyy年m月d日"
                    .Cells(DaySheetBLastRow + 1, 1).NumberFormat = "m月d日"
                    .Cells(DaySheetBLastRow + 1, 2) = ShopeetempSheet.Range("A" & ShopeetempSheetLastRow)
                    .Cells(DaySheetBLastRow + 1, 3) = NameSet
                    .Cells(DaySheetBLastRow + 1, 15) = OrderNumSet
                    .Cells(DaySheetBLastRow + 1, 4) = TotalRevenueB
                    .Cells(DaySheetBLastRow + 1, 17) = ShopeetempSheet.Range("I" & ShopeetempSheetLastRow)
                    .Cells(DaySheetBLastRow + 1, 11) = TotalCostB
                    .Cells(DaySheetBLastRow + 1, 13) = ShopeetempSheet.Range("G" & ShopeetempSheetLastRow)
                    .Cells(DaySheetBLastRow + 1, 13).Font.ColorIndex = 3
                    .Cells(DaySheetBLastRow + 1, 14) = "蝦皮"
                    .Cells(DaySheetBLastRow + 1, 14).Font.ColorIndex = 46
                End With
                
                On Error Resume Next
                RatioA = TotalRevenueA / (TotalRevenueA + TotalRevenueB)
                RatioB = TotalRevenueB / (TotalRevenueA + TotalRevenueB)
                On Error GoTo 0

                With DaySheetA
                    .Cells(DaySheetALastRow + 1, 8) = RatioA * WorksheetFunction.VLookup(k, Shopee_sheet.Range("A2:S" & ShopeeSheetLastRow), 17, False)
                    .Cells(DaySheetALastRow + 1, 9) = RatioA * WorksheetFunction.VLookup(k, Shopee_sheet.Range("A2:S" & ShopeeSheetLastRow), 18, False)
                    .Cells(DaySheetALastRow + 1, 10) = RatioA * WorksheetFunction.VLookup(k, Shopee_sheet.Range("A2:S" & ShopeeSheetLastRow), 19, False)
                End With
                
                With DaySheetB
                    .Cells(DaySheetBLastRow + 1, 8) = RatioB * WorksheetFunction.VLookup(k, Shopee_sheet.Range("A2:S" & ShopeeSheetLastRow), 17, False)
                    .Cells(DaySheetBLastRow + 1, 9) = RatioB * WorksheetFunction.VLookup(k, Shopee_sheet.Range("A2:S" & ShopeeSheetLastRow), 18, False)
                    .Cells(DaySheetBLastRow + 1, 10) = RatioB * WorksheetFunction.VLookup(k, Shopee_sheet.Range("A2:S" & ShopeeSheetLastRow), 19, False)
                End With
                
                With ShopeeRatioSheet
                    ShopeeRatioSheetLastRow = ShopeeRatioSheet.Range("A1048576").End(xlUp).Row
                    .Range("A" & ShopeeRatioSheetLastRow + 1) = k
                    .Range("A" & ShopeeRatioSheetLastRow + 1).NumberFormat = "0"
                    .Range("B" & ShopeeRatioSheetLastRow + 1) = RatioA
                    .Range("C" & ShopeeRatioSheetLastRow + 1) = RatioB
                End With

        Next k
        
    End With
    
    With DaySheetA
            DaySheetALastRow = .Range("A1048576").End(xlUp).Row
            ShopeeRatioSheetLastRow = ShopeeRatioSheet.Range("A1048576").End(xlUp).Row
            For p = 2 To DaySheetALastRow
                    AA = ""
                    On Error Resume Next
                    AA = WorksheetFunction.VLookup(.Cells(p, 2), ShopeeRatioSheet.Range("A2:C" & ShopeeRatioSheetLastRow), 2, False)
                    On Error GoTo 0
                    If AA <> "" Then
                            .Cells(p, 16) = AA
                    End If
            Next p
    End With
    
    With DaySheetB
            DaySheetBLastRow = .Range("A1048576").End(xlUp).Row
            ShopeeRatioSheetLastRow = ShopeeRatioSheet.Range("A1048576").End(xlUp).Row
            For p = 2 To DaySheetBLastRow
                    BB = ""
                    On Error Resume Next
                    BB = WorksheetFunction.VLookup(.Cells(p, 2), ShopeeRatioSheet.Range("A2:C" & ShopeeRatioSheetLastRow), 3, False)
                    On Error GoTo 0
                    If BB <> "" Then
                            .Cells(p, 16) = BB
                    End If
            Next p
    End With
    
    ShopeetempSheet.Delete
    ShopeeRatioSheet.Delete
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Control Panel").Activate
    MsgBox "Complete!"

End Sub
