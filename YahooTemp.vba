Sub YahooTemp()

    Application.DisplayAlerts = False
    
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "Yahoo_temp"
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "Yahoo_Ratio"
    Dim StorageSheet, CompareSheet, YahootempSheet, YahooRatioSheet, Yahoo_sheet, DaySheetA, DaySheetB As Worksheet
    Set YahootempSheet = ThisWorkbook.Sheets("Yahoo_temp")
    Set YahooRatioSheet = ThisWorkbook.Sheets("Yahoo_Ratio")
    Set CompareSheet = ThisWorkbook.Sheets("對照表")
    Set StorageSheet = ThisWorkbook.Sheets("入庫")
    Set Yahoo_sheet = ThisWorkbook.Sheets("雅虎orders")
    Set DaySheetA = ThisWorkbook.Sheets("日報表A")
    Set DaySheetB = ThisWorkbook.Sheets("日報表B")

    YahoosheetLastRow = Yahoo_sheet.Range("A1048576").End(xlUp).Row
    CompareSheetLastRow = CompareSheet.Range("A1048576").End(xlUp).Row
    
    With YahooRatioSheet
        .Range("A1") = "訂單編號"
        .Range("B1") = "RatioA"
        .Range("C1") = "RatioB"
    End With
    
    With YahootempSheet
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
        
        For i = 2 To YahoosheetLastRow
            
            .Cells(i, 1) = Yahoo_sheet.Cells(i, 3)
            .Cells(i, 1).NumberFormat = "0"
            .Cells(i, 2) = Yahoo_sheet.Cells(i, 6) & "[" & Yahoo_sheet.Cells(i, 10) & "," & Yahoo_sheet.Cells(i, 11) & "]"
            .Cells(i, 3) = WorksheetFunction.VLookup(.Range("B" & i), CompareSheet.Range("A2:D" & CompareSheetLastRow), 4, False)
            .Cells(i, 4) = Yahoo_sheet.Cells(i, 16)
            .Cells(i, 6) = WorksheetFunction.VLookup(.Range("B" & i), CompareSheet.Range("A2:F" & CompareSheetLastRow), 6, False)
            .Cells(i, 8) = Left(Yahoo_sheet.Cells(i, 1), InStr(1, Yahoo_sheet.Cells(i, 1), " ") - 1)
            .Cells(i, 9) = Yahoo_sheet.Cells(i, 17)
            .Cells(i, 10) = Yahoo_sheet.Cells(i, 15)
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

            If Yahoo_sheet.Cells(i, 38) Like "*已取消*" And Yahoo_sheet.Cells(i, 41) Like "*賣家已取退貨*" Then
                .Cells(i, 7) = "!棄領!"
            End If
            
            If Yahoo_sheet.Cells(i, 38) Like "*已取消*" And Yahoo_sheet.Cells(i, 41) Like "*尚未出貨*" Then
                .Cells(i, 8) = "日期"
            End If
            
        Next i
        
        StorageSheet.Range("I:J").ClearContents
        
        '依訂單排序
        With YahootempSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A1:A" & YahoosheetLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range("F1:F" & YahoosheetLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A1:K" & YahoosheetLastRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        'Unique 出所有訂單
        Dim d As Object, c As Range, k, tmp As String
        
        Dim dctData As Object
        Set dctData = CreateObject("scripting.dictionary")
    
        Set d = CreateObject("scripting.dictionary")
        For Each c In .Range("A2:A" & YahoosheetLastRow)
            tmp = Trim(c.Value)
            If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
        Next c

        For Each k In d.keys
                
                .UsedRange.AutoFilter Field:=1, Criteria1:=k
                
                'Type A
                .UsedRange.AutoFilter Field:=6, Criteria1:="A"
                TotalRevenueA = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
                TotalCostA = Application.WorksheetFunction.Sum(.Range("E:E").SpecialCells(xlCellTypeVisible))
                YahootempSheetLastRow = .Range("A1048576").End(xlUp).Row
                
                FilterdCount = Application.WorksheetFunction.Subtotal(3, .Range("B:B")) - 1
                
                OrderNumSet = ""
                NameSet = ""
                If FilterdCount = 0 Then
                        OrderNumSet = ""
                        NameSet = ""
                ElseIf FilterdCount = 1 Then
                        OrderNumSet = .Range("C" & YahootempSheetLastRow) & "(" & .Range("J" & YahootempSheetLastRow) & ")"
                        NameSet = .Range("K" & YahootempSheetLastRow)
                        '2021/2/4
                        If .Range("C" & YahootempSheetLastRow) = "TBD" Then
                                .Range("G" & YahootempSheetLastRow) = "!未匹配!"
                        End If
                Else
                        For m = YahootempSheetLastRow - FilterdCount + 1 To YahootempSheetLastRow
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
                            For f = YahootempSheetLastRow - FilterdCount + 1 To YahootempSheetLastRow
                                    .Range("G" & f) = "!未匹配!"
                            Next f
                        End If
                End If

                            
                With DaySheetA
                    DaySheetALastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(DaySheetALastRow + 1, 1) = YahootempSheet.Range("H" & YahootempSheetLastRow)
                    .Cells(DaySheetALastRow + 1, 1).NumberFormat = "yyyy年m月d日"
                    .Cells(DaySheetALastRow + 1, 1).NumberFormat = "m月d日"
                    .Cells(DaySheetALastRow + 1, 2) = Val(YahootempSheet.Range("A" & YahootempSheetLastRow))
                    .Cells(DaySheetALastRow + 1, 2).NumberFormat = "0"
                    .Cells(DaySheetALastRow + 1, 3) = NameSet
                    .Cells(DaySheetALastRow + 1, 15) = OrderNumSet
                    .Cells(DaySheetALastRow + 1, 4) = TotalRevenueA
                    .Cells(DaySheetALastRow + 1, 17) = YahootempSheet.Range("I" & YahootempSheetLastRow)
                    .Cells(DaySheetALastRow + 1, 11) = TotalCostA
                    .Cells(DaySheetALastRow + 1, 13) = YahootempSheet.Range("G" & YahootempSheetLastRow)
                    .Cells(DaySheetALastRow + 1, 13).Font.ColorIndex = 3
                    .Cells(DaySheetALastRow + 1, 14) = "Y拍"
                    .Cells(DaySheetALastRow + 1, 14).Font.ColorIndex = 29
                End With
                
                'Type B
                .UsedRange.AutoFilter Field:=6, Criteria1:="B"
                TotalRevenueB = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
                TotalCostB = Application.WorksheetFunction.Sum(.Range("E:E").SpecialCells(xlCellTypeVisible))
                YahootempSheetLastRow = .Range("A1048576").End(xlUp).Row
                
                FilterdCount = Application.WorksheetFunction.Subtotal(3, .Range("B:B")) - 1
                
                OrderNumSet = ""
                NameSet = ""
                If FilterdCount = 0 Then
                        OrderNumSet = ""
                        NameSet = ""
                ElseIf FilterdCount = 1 Then
                        OrderNumSet = .Range("C" & YahootempSheetLastRow) & "(" & .Range("J" & YahootempSheetLastRow) & ")"
                        NameSet = .Range("K" & YahootempSheetLastRow)
                        '2021/2/4
                        If .Range("C" & YahootempSheetLastRow) = "TBD" Then
                                .Range("G" & YahootempSheetLastRow) = "!未匹配!"
                        End If
                Else
                        For m = YahootempSheetLastRow - FilterdCount + 1 To YahootempSheetLastRow
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
                            For f = YahootempSheetLastRow - FilterdCount + 1 To YahootempSheetLastRow
                                    .Range("G" & f) = "!未匹配!"
                            Next f
                        End If
                End If

                With DaySheetB
                    DaySheetBLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(DaySheetBLastRow + 1, 1) = YahootempSheet.Range("H" & YahootempSheetLastRow)
                    .Cells(DaySheetBLastRow + 1, 1).NumberFormat = "yyyy年m月d日"
                    .Cells(DaySheetBLastRow + 1, 1).NumberFormat = "m月d日"
                    .Cells(DaySheetBLastRow + 1, 2) = Val(YahootempSheet.Range("A" & YahootempSheetLastRow))
                    .Cells(DaySheetBLastRow + 1, 3) = NameSet
                    .Cells(DaySheetBLastRow + 1, 2).NumberFormat = "0"
                    .Cells(DaySheetBLastRow + 1, 15) = OrderNumSet
                    .Cells(DaySheetBLastRow + 1, 4) = TotalRevenueB
                    .Cells(DaySheetBLastRow + 1, 17) = YahootempSheet.Range("I" & YahootempSheetLastRow)
                    .Cells(DaySheetBLastRow + 1, 11) = TotalCostB
                    .Cells(DaySheetBLastRow + 1, 13) = YahootempSheet.Range("G" & YahootempSheetLastRow)
                    .Cells(DaySheetBLastRow + 1, 13).Font.ColorIndex = 3
                    .Cells(DaySheetBLastRow + 1, 14) = "Y拍"
                    .Cells(DaySheetBLastRow + 1, 14).Font.ColorIndex = 29
                End With
                
                On Error Resume Next
                RatioA = TotalRevenueA / (TotalRevenueA + TotalRevenueB)
                RatioB = TotalRevenueB / (TotalRevenueA + TotalRevenueB)
                On Error GoTo 0
                
                With YahooRatioSheet
                    YahooRatioSheetLastRow = YahooRatioSheet.Range("A1048576").End(xlUp).Row
                    .Range("A" & YahooRatioSheetLastRow + 1) = k
                    .Range("A" & YahooRatioSheetLastRow + 1).NumberFormat = "0"
                    .Range("B" & YahooRatioSheetLastRow + 1) = RatioA
                    .Range("C" & YahooRatioSheetLastRow + 1) = RatioB
                End With

        Next k
        
    End With
    
    With DaySheetA
            DaySheetALastRow = .Range("A1048576").End(xlUp).Row
            YahooRatioSheetLastRow = YahooRatioSheet.Range("A1048576").End(xlUp).Row
            For p = 2 To DaySheetALastRow
                    AA = ""
                    On Error Resume Next
                    AA = WorksheetFunction.VLookup(.Cells(p, 2), YahooRatioSheet.Range("A2:C" & YahooRatioSheetLastRow), 2, False)
                    On Error GoTo 0
                    If AA <> "" Then
                            .Cells(p, 16) = AA
                    End If
            Next p
    End With
    
    With DaySheetB
            DaySheetBLastRow = .Range("A1048576").End(xlUp).Row
            YahooRatioSheetLastRow = YahooRatioSheet.Range("A1048576").End(xlUp).Row
            For p = 2 To DaySheetBLastRow
                    BB = ""
                    On Error Resume Next
                    BB = WorksheetFunction.VLookup(.Cells(p, 2), YahooRatioSheet.Range("A2:C" & YahooRatioSheetLastRow), 3, False)
                    On Error GoTo 0
                    If BB <> "" Then
                            .Cells(p, 16) = BB
                    End If
            Next p
    End With
    
    YahootempSheet.Delete
    YahooRatioSheet.Delete
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Control Panel").Activate
    MsgBox "Complete!"
    
End Sub
