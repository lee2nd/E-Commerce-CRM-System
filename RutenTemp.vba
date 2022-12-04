Sub RutenTemp()

    Application.DisplayAlerts = False
    
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "Ruten_temp"
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "Ruten_Ratio"
    Dim StorageSheet, CompareSheet, RutentempSheet, RutenRatioSheet, Ruten_sheet, DaySheetA, DaySheetB As Worksheet
    Set RutentempSheet = ThisWorkbook.Sheets("Ruten_temp")
    Set RutenRatioSheet = ThisWorkbook.Sheets("Ruten_Ratio")
    Set CompareSheet = ThisWorkbook.Sheets("對照表")
    Set StorageSheet = ThisWorkbook.Sheets("入庫")
    Set Ruten_sheet = ThisWorkbook.Sheets("露天orders")
    Set DaySheetA = ThisWorkbook.Sheets("日報表A")
    Set DaySheetB = ThisWorkbook.Sheets("日報表B")

    RutensheetLastRow = Ruten_sheet.Range("A1048576").End(xlUp).Row
    CompareSheetLastRow = CompareSheet.Range("A1048576").End(xlUp).Row

    With RutenRatioSheet
        .Range("A1") = "訂單編號"
        .Range("B1") = "RatioA"
        .Range("C1") = "RatioB"
    End With
    
    With RutentempSheet
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
        
        For i = 2 To RutensheetLastRow
            
            .Cells(i, 1) = Ruten_sheet.Cells(i, 2)
            .Cells(i, 1).NumberFormat = "0"
            .Cells(i, 2) = Ruten_sheet.Cells(i, 6) & "[" & Ruten_sheet.Cells(i, 7) & "," & Ruten_sheet.Cells(i, 8) & "]"
            .Cells(i, 3) = WorksheetFunction.VLookup(.Range("B" & i), CompareSheet.Range("A2:D" & CompareSheetLastRow), 4, False)
            .Cells(i, 4) = Ruten_sheet.Cells(i, 10) * Ruten_sheet.Cells(i, 11)
            .Cells(i, 6) = WorksheetFunction.VLookup(.Range("B" & i), CompareSheet.Range("A2:F" & CompareSheetLastRow), 6, False)
            .Cells(i, 8) = Ruten_sheet.Cells(i, 1)
            .Cells(i, 9) = Ruten_sheet.Cells(i, 13)
            .Cells(i, 10) = Ruten_sheet.Cells(i, 10)
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
            
            If Ruten_sheet.Cells(i, 17) Like "*已領退貨*" Then
                .Cells(i, 7) = "!棄領!"
            End If
        
        Next i
        
        StorageSheet.Range("I:J").ClearContents
        
        '依訂單排序
        With RutentempSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A1:A" & RutensheetLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=Range("F1:F" & RutensheetLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A1:K" & RutensheetLastRow)
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
        For Each c In .Range("A2:A" & RutensheetLastRow)
            tmp = Trim(c.Value)
            If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
        Next c
        
        For Each k In d.keys
                
                .UsedRange.AutoFilter Field:=1, Criteria1:=k
                            
                'Type A
                .UsedRange.AutoFilter Field:=6, Criteria1:="A"
                TotalRevenueA = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
                TotalCostA = Application.WorksheetFunction.Sum(.Range("E:E").SpecialCells(xlCellTypeVisible))
                RutentempSheetLastRow = .Range("A1048576").End(xlUp).Row
                
                FilterdCount = Application.WorksheetFunction.Subtotal(3, .Range("B:B")) - 1
                
                OrderNumSet = ""
                NameSet = ""
                If FilterdCount = 0 Then
                        OrderNumSet = ""
                        NameSet = ""
                ElseIf FilterdCount = 1 Then
                        OrderNumSet = .Range("C" & RutentempSheetLastRow) & "(" & .Range("J" & RutentempSheetLastRow) & ")"
                        NameSet = .Range("K" & RutentempSheetLastRow)
                        '2021/2/4
                        If .Range("C" & RutentempSheetLastRow) = "TBD" Then
                                .Range("G" & RutentempSheetLastRow) = "!未匹配!"
                        End If
                Else
                        For m = RutentempSheetLastRow - FilterdCount + 1 To RutentempSheetLastRow
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
                            For f = RutentempSheetLastRow - FilterdCount + 1 To RutentempSheetLastRow
                                    .Range("G" & f) = "!未匹配!"
                            Next f
                        End If
                End If
                            
                With DaySheetA
                    DaySheetALastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(DaySheetALastRow + 1, 1) = RutentempSheet.Range("H" & RutentempSheetLastRow)
                    .Cells(DaySheetALastRow + 1, 1).NumberFormat = "yyyy年m月d日"
                    .Cells(DaySheetALastRow + 1, 1).NumberFormat = "m月d日"
                    .Cells(DaySheetALastRow + 1, 2) = Val(RutentempSheet.Range("A" & RutentempSheetLastRow))
                    .Cells(DaySheetALastRow + 1, 2).NumberFormat = "0"
                    .Cells(DaySheetALastRow + 1, 3) = NameSet
                    .Cells(DaySheetALastRow + 1, 15) = OrderNumSet
                    .Cells(DaySheetALastRow + 1, 4) = TotalRevenueA
                    .Cells(DaySheetALastRow + 1, 17) = RutentempSheet.Range("I" & RutentempSheetLastRow)
                    .Cells(DaySheetALastRow + 1, 11) = TotalCostA
                    .Cells(DaySheetALastRow + 1, 13) = RutentempSheet.Range("G" & RutentempSheetLastRow)
                    .Cells(DaySheetALastRow + 1, 13).Font.ColorIndex = 3
                    .Cells(DaySheetALastRow + 1, 14) = "露天"
                    .Cells(DaySheetALastRow + 1, 14).Font.ColorIndex = 50
                End With
                
                'Type B
                .UsedRange.AutoFilter Field:=6, Criteria1:="B"
                TotalRevenueB = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
                TotalCostB = Application.WorksheetFunction.Sum(.Range("E:E").SpecialCells(xlCellTypeVisible))
                RutentempSheetLastRow = .Range("A1048576").End(xlUp).Row

                FilterdCount = Application.WorksheetFunction.Subtotal(3, .Range("B:B")) - 1
                
                OrderNumSet = ""
                NameSet = ""
                If FilterdCount = 0 Then
                        OrderNumSet = ""
                        NameSet = ""
                ElseIf FilterdCount = 1 Then
                        OrderNumSet = .Range("C" & RutentempSheetLastRow) & "(" & .Range("J" & RutentempSheetLastRow) & ")"
                        NameSet = .Range("K" & RutentempSheetLastRow)
                        '2021/2/4
                        If .Range("C" & RutentempSheetLastRow) = "TBD" Then
                                .Range("G" & RutentempSheetLastRow) = "!未匹配!"
                        End If
                Else
                        For m = RutentempSheetLastRow - FilterdCount + 1 To RutentempSheetLastRow
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
                            For f = RutentempSheetLastRow - FilterdCount + 1 To RutentempSheetLastRow
                                    .Range("G" & f) = "!未匹配!"
                            Next f
                        End If
                                
                End If
                
                With DaySheetB
                    DaySheetBLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(DaySheetBLastRow + 1, 1) = RutentempSheet.Range("H" & RutentempSheetLastRow)
                    .Cells(DaySheetBLastRow + 1, 1).NumberFormat = "yyyy年m月d日"
                    .Cells(DaySheetBLastRow + 1, 1).NumberFormat = "m月d日"
                    .Cells(DaySheetBLastRow + 1, 2) = Val(RutentempSheet.Range("A" & RutentempSheetLastRow))
                    .Cells(DaySheetBLastRow + 1, 2).NumberFormat = "0"
                    .Cells(DaySheetBLastRow + 1, 3) = NameSet
                    .Cells(DaySheetBLastRow + 1, 15) = OrderNumSet
                    .Cells(DaySheetBLastRow + 1, 4) = TotalRevenueB
                    .Cells(DaySheetBLastRow + 1, 17) = RutentempSheet.Range("I" & RutentempSheetLastRow)
                    .Cells(DaySheetBLastRow + 1, 11) = TotalCostB
                    .Cells(DaySheetBLastRow + 1, 13) = RutentempSheet.Range("G" & RutentempSheetLastRow)
                    .Cells(DaySheetBLastRow + 1, 13).Font.ColorIndex = 3
                    .Cells(DaySheetBLastRow + 1, 14) = "露天"
                    .Cells(DaySheetBLastRow + 1, 14).Font.ColorIndex = 50
                End With
                
                On Error Resume Next
                RatioA = TotalRevenueA / (TotalRevenueA + TotalRevenueB)
                RatioB = TotalRevenueB / (TotalRevenueA + TotalRevenueB)
                On Error GoTo 0
                
                With RutenRatioSheet
                    RutenRatioSheetLastRow = RutenRatioSheet.Range("A1048576").End(xlUp).Row
                    .Range("A" & RutenRatioSheetLastRow + 1) = k
                    .Range("A" & RutenRatioSheetLastRow + 1).NumberFormat = "0"
                    .Range("B" & RutenRatioSheetLastRow + 1) = RatioA
                    .Range("C" & RutenRatioSheetLastRow + 1) = RatioB
                End With

        Next k
        
    End With
    
    With DaySheetA
            DaySheetALastRow = .Range("A1048576").End(xlUp).Row
            RutenRatioSheetLastRow = RutenRatioSheet.Range("A1048576").End(xlUp).Row
            For p = 2 To DaySheetALastRow
                    AA = ""
                    On Error Resume Next
                    AA = WorksheetFunction.VLookup(.Cells(p, 2), RutenRatioSheet.Range("A2:C" & RutenRatioSheetLastRow), 2, False)
                    On Error GoTo 0
                    If AA <> "" Then
                            .Cells(p, 16) = AA
                    End If
            Next p
    End With
    
    With DaySheetB
            DaySheetBLastRow = .Range("A1048576").End(xlUp).Row
            RutenRatioSheetLastRow = RutenRatioSheet.Range("A1048576").End(xlUp).Row
            For p = 2 To DaySheetBLastRow
                    BB = ""
                    On Error Resume Next
                    BB = WorksheetFunction.VLookup(.Cells(p, 2), RutenRatioSheet.Range("A2:C" & RutenRatioSheetLastRow), 3, False)
                    On Error GoTo 0
                    If BB <> "" Then
                            .Cells(p, 16) = BB
                    End If
            Next p
    End With
    
    RutentempSheet.Delete
    RutenRatioSheet.Delete
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Control Panel").Activate
    MsgBox "Complete!"

End Sub

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
