Sub DailyB()

    Dim DaySheetB, ControlPanelSheet, TagSheet, Shopee_sheet As Worksheet
    Set DaySheetB = ThisWorkbook.Sheets("日報表B")
    Set ControlPanelSheet = ThisWorkbook.Sheets("Control Panel")
    Set TagSheet = ThisWorkbook.Sheets("促銷組合標籤")
    Set Shopee_sheet = ThisWorkbook.Sheets("蝦皮orders")
    
    Application.ScreenUpdating = False

    DaySheetBLastRow = DaySheetB.Range("A1048576").End(xlUp).Row
    
    Dim SS As Variant
    
    With DaySheetB
    
        For i = DaySheetBLastRow To 2 Step -1
            If .Cells(i, 1) = "日期" Then
                .Rows(i).Delete
            End If
        Next i
        
        DaySheetBLastRow = .Range("A1048576").End(xlUp).Row
        
        For j = 2 To DaySheetBLastRow
        
            If .Cells(j, 13) = "!退貨!" Then
                .Cells(j, 4) = 0
                .Cells(j, 12) = 0
                .Cells(j, 11) = 0
            ElseIf .Cells(j, 13) = "!棄領!" Then
                .Cells(j, 4) = -60
                .Cells(j, 12) = -60
                .Cells(j, 11) = 0
            Else
                '賣家組合折扣
                .Cells(j, 5) = .Cells(j, 17) * .Cells(j, 16)
                
                '促銷組合折扣(只看蝦皮)
                If .Cells(j, 14) = "蝦皮" And (.Cells(j, 6) = "") Then
                
                    ShopeeSheetLastRow = Shopee_sheet.Range("A1048576").End(xlUp).Row
                    On Error Resume Next
                    If WorksheetFunction.VLookup(.Range("B" & j), Shopee_sheet.Range("A2:AD" & ShopeeSheetLastRow), 30, False) <> "" Then
                    On Error GoTo 0
                
                    SS = Split(.Cells(j, 15), ";")
                    For i = LBound(SS) To UBound(SS)
                        CulCount = Split(SS(i), "(")
                        Order = CulCount(0)
                        Count = Replace(CulCount(1), ")", "")
                        With TagSheet
                            TagSheetLastRow = .Range("A1048576").End(xlUp).Row
                            For k = 2 To TagSheetLastRow
                                If .Cells(k, 1) Like "*" & Order & "*" Then
                                    
                                    Counter = 0
                                    While .Cells(k, 8 + Counter) <> ""
                                        Counter = Counter + 1
                                    Wend
                                    
                                    .Cells(k, 8 + Counter) = Count
                                    
                                End If
                            Next k
                        End With
                    Next i
                    
                    '算總折扣金額
                    TotalDiscount = 0
                    
                    For o = 2 To TagSheetLastRow
                    
                        TotalCount = 0
                        If TagSheet.Cells(o, 8) <> "" Then
                            TotalCount = WorksheetFunction.Sum(TagSheet.Range(TagSheet.Cells(o, 8), TagSheet.Cells(o, 8).End(xlToRight)))
                        End If
                        
                        TotalDiscount = TotalDiscount + ((TagSheet.Cells(o, 7)) * (TotalCount \ TagSheet.Cells(o, 6)))
                            
                    Next o
                    
                    '累積資料全清掉
                    TagSheet.Range("H:XFD").ClearContents
                    
                    .Cells(j, 6) = TotalDiscount
                    
                    End If

                            .Cells(j, 12).Formula = "=D" & j & "-E" & j & "-F" & j & "-G" & j & "-H" & j & "-I" & j & "-J" & j & "-K" & j
                    
                    ElseIf .Cells(j, 14) = "Y拍" Then
                            .Cells(j, 8) = WorksheetFunction.Round((.Cells(j, 4) - .Cells(j, 5) - .Cells(j, 6) - .Cells(j, 7)) * 0.0199, 0)
                            .Cells(j, 9) = 0
                            .Cells(j, 10) = 0
                            .Cells(j, 12).Formula = "=D" & j & "-E" & j & "-F" & j & "-G" & j & "-H" & j & "-I" & j & "-J" & j & "-K" & j
                            
                    ElseIf .Cells(j, 14) = "露天" Then
                            .Cells(j, 8) = WorksheetFunction.Round((.Cells(j, 4) - .Cells(j, 5) - .Cells(j, 6) - .Cells(j, 7)) * 0.02, 0)
                            .Cells(j, 9) = 0
                            .Cells(j, 10) = 0
                            If .Cells(j, 1) > DateValue("2021/04/25") Then
                                .Cells(j, 10) = WorksheetFunction.Round((.Cells(j, 4) - .Cells(j, 5) - .Cells(j, 6) - .Cells(j, 7)) * 0.01, 0)
                            End If
                            .Cells(j, 12).Formula = "=D" & j & "-E" & j & "-F" & j & "-G" & j & "-H" & j & "-I" & j & "-J" & j & "-K" & j
                    
                End If
        
                '滿額運費折抵
                If (.Cells(j, 16) <> 0) Then
                    If (.Cells(j, 7) = "") And (.Cells(j, 14) = "蝦皮") And ((.Cells(j, 4) / .Cells(j, 16)) >= ControlPanelSheet.Range("Q3")) Then
                        .Cells(j, 7) = ControlPanelSheet.Range("Q4") * .Cells(j, 16)
                    ElseIf (.Cells(j, 7) = "") And (.Cells(j, 14) = "Y拍") And ((.Cells(j, 4) / .Cells(j, 16)) >= ControlPanelSheet.Range("R3")) Then
                        .Cells(j, 7) = ControlPanelSheet.Range("R4") * .Cells(j, 16)
                    ElseIf (.Cells(j, 7) = "") And (.Cells(j, 14) = "露天") And ((.Cells(j, 4) / .Cells(j, 16)) >= ControlPanelSheet.Range("S3")) Then
                        .Cells(j, 7) = ControlPanelSheet.Range("S4") * .Cells(j, 16)
                    End If
                Else
                        .Cells(j, 7) = 0
                End If
                
            End If
            
        Next j
        
        For j = 2 To DaySheetBLastRow
                If .Cells(j, 11) = 0 And .Cells(j, 13) = "" Then
                    .Cells(j, 13) = "!未匹配!"
                    .Cells(j, 13).Font.ColorIndex = 3
                End If
        Next j
        
    End With
    
    With DaySheetB
    
        '刪除重複資料
        .UsedRange.RemoveDuplicates Columns:=Array(1, 2, 6, 14), Header:=xlYes
        
        For i = .Range("B1").End(xlDown).Row To 2 Step -1
            If Application.WorksheetFunction.CountIf(.Range(.Cells(2, 2), .Cells(i, 2)), .Cells(i, 2)) > 1 Then
                .Rows(i).Delete
            End If
        Next i
        
        '調整字體
        .Cells.Font.Size = 11
        .Cells.Font.Name = "微軟正黑體"
        .Cells.VerticalAlignment = xlVAlignCenter
        .Cells.HorizontalAlignment = xlHAlignLeft

        '自動調整欄寬
        .Columns("A:O").AutoFit
        .Columns("C:C").ColumnWidth = 18
        
        DaySheetBLastRow = .Range("A1048576").End(xlUp).Row

        '依日期排序
        With DaySheetB.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A1:A" & DaySheetBLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A1:O" & DaySheetBLastRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
    End With
    
    ThisWorkbook.Sheets("Control Panel").Activate
    MsgBox "Complete!"
    
    Application.ScreenUpdating = True

End Sub



