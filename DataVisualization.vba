Sub data_visualization()

    Dim DaySheetA, DaySheetB, VisualizationSheet As Worksheet
    Set DaySheetA = ThisWorkbook.Sheets("日報表A")
    Set DaySheetB = ThisWorkbook.Sheets("日報表B")
    Set VisualizationSheet = ThisWorkbook.Sheets("圖表")
    
    VisualizationSheet.Range("B2:G13") = ""

    With VisualizationSheet
    
        Select Case VisualizationSheet.Range("D19")
    
            Case "A":
                
                '計算各平台訂單量
                DaySheetALastRow = DaySheetA.Range("A1048576").End(xlUp).Row
                For i = 1 To 12
                    .Cells(i + 1, 2).Formula = "=SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*(日報表A!M2:M" & DaySheetALastRow & "="""")*(日報表A!N2:N" & DaySheetALastRow & "=""蝦皮""))"
                Next i
                
                For i = 1 To 12
                    .Cells(i + 1, 3).Formula = "=SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*(日報表A!M2:M" & DaySheetALastRow & "="""")*(日報表A!N2:N" & DaySheetALastRow & "=""露天""))"
                Next i
            
                For i = 1 To 12
                    .Cells(i + 1, 4).Formula = "=SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*(日報表A!M2:M" & DaySheetALastRow & "="""")*(日報表A!N2:N" & DaySheetALastRow & "=""Y拍""))"
                Next i
                
                '計算總營業額
                For i = 1 To 12
                    .Cells(i + 1, 5).Formula = "=SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*1,日報表A!D2:D" & DaySheetALastRow & ")-SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*1,日報表A!E2:E" & DaySheetALastRow & ")-SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*1,日報表A!F2:F" & DaySheetALastRow & ")"
                Next i
                
                '計算總淨利額
                For i = 1 To 12
                    .Cells(i + 1, 6).Formula = "=ROUND(SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*1,日報表A!L2:L" & DaySheetALastRow & ")-月報表A!H" & i + 1 & "-月報表A!J" & i + 1 & ",0)"
                Next i
                
                '計算年度平均訂單量
                TotalMonth = 0
                
                For X = 2 To 13
                    ordersum = 0
                    ordersum = .Range("B" & X) + .Range("C" & X) + .Range("D" & X)
                    If ordersum > 0 Then
                        TotalMonth = TotalMonth + 1
                    End If
                Next X

                .Range("G2:G13") = Application.WorksheetFunction.Sum(.Range("B2:D13")) / TotalMonth
                    
            Case "B":
            
                '計算各平台訂單量
                DaySheetBLastRow = DaySheetB.Range("A1048576").End(xlUp).Row
                For i = 1 To 12
                    .Cells(i + 1, 2).Formula = "=SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*(日報表B!M2:M" & DaySheetBLastRow & "="""")*(日報表B!N2:N" & DaySheetBLastRow & "=""蝦皮""))"
                Next i
                
                For i = 1 To 12
                    .Cells(i + 1, 3).Formula = "=SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*(日報表B!M2:M" & DaySheetBLastRow & "="""")*(日報表B!N2:N" & DaySheetBLastRow & "=""露天""))"
                Next i
            
                For i = 1 To 12
                    .Cells(i + 1, 4).Formula = "=SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*(日報表B!M2:M" & DaySheetBLastRow & "="""")*(日報表B!N2:N" & DaySheetBLastRow & "=""Y拍""))"
                Next i
                
                '計算總營業額
                For i = 1 To 12
                    .Cells(i + 1, 5).Formula = "=SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*1,日報表B!D2:D" & DaySheetBLastRow & ")-SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*1,日報表B!E2:E" & DaySheetBLastRow & ")-SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*1,日報表B!F2:F" & DaySheetBLastRow & ")"
                Next i
                
                '計算總淨利額
                For i = 1 To 12
                    .Cells(i + 1, 6).Formula = "=ROUND(SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*1,日報表B!L2:L" & DaySheetBLastRow & ")-月報表B!I" & i + 1 & "-月報表B!K" & i + 1 & ",0)"
                Next i
                
                '計算年度平均訂單量
                TotalMonth = 0
                
                For X = 2 To 13
                    ordersum = 0
                    ordersum = .Range("B" & X) + .Range("C" & X) + .Range("D" & X)
                    If ordersum > 0 Then
                        TotalMonth = TotalMonth + 1
                    End If
                Next X

                .Range("G2:G13") = Application.WorksheetFunction.Sum(.Range("B2:D13")) / TotalMonth
                
            Case "A+B":
            
                '計算各平台訂單量
                DaySheetALastRow = DaySheetA.Range("A1048576").End(xlUp).Row
                DaySheetBLastRow = DaySheetB.Range("A1048576").End(xlUp).Row
                For i = 1 To 12
                    .Cells(i + 1, 2).Formula = "=SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*(日報表A!M2:M" & DaySheetALastRow & "="""")*(日報表A!N2:N" & DaySheetALastRow & "=""蝦皮""))+SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*(日報表B!M2:M" & DaySheetBLastRow & "="""")*(日報表B!N2:N" & DaySheetBLastRow & "=""蝦皮""))"
                Next i
                
                For i = 1 To 12
                    .Cells(i + 1, 3).Formula = "=SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*(日報表A!M2:M" & DaySheetALastRow & "="""")*(日報表A!N2:N" & DaySheetALastRow & "=""露天""))+SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*(日報表B!M2:M" & DaySheetBLastRow & "="""")*(日報表B!N2:N" & DaySheetBLastRow & "=""露天""))"
                Next i
            
                For i = 1 To 12
                    .Cells(i + 1, 4).Formula = "=SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*(日報表A!M2:M" & DaySheetALastRow & "="""")*(日報表A!N2:N" & DaySheetALastRow & "=""Y拍""))+SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*(日報表B!M2:M" & DaySheetBLastRow & "="""")*(日報表B!N2:N" & DaySheetBLastRow & "=""Y拍""))"
                Next i
                
                '計算總營業額
                For i = 1 To 12
                    .Cells(i + 1, 5).Formula = "=SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*1,日報表A!D2:D" & DaySheetALastRow & ")-SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*1,日報表A!E2:E" & DaySheetALastRow & ")-SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*1,日報表A!F2:F" & DaySheetALastRow & ")+SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*1,日報表B!D2:D" & DaySheetBLastRow & ")-SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*1,日報表B!E2:E" & DaySheetBLastRow & ")-SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*1,日報表B!F2:F" & DaySheetBLastRow & ")"
                Next i
                
                '計算總淨利額
                For i = 1 To 12
                    .Cells(i + 1, 6).Formula = "=ROUND((SUMPRODUCT((MONTH(日報表A!A2:A" & DaySheetALastRow & ")=" & i & ")*1,日報表A!L2:L" & DaySheetALastRow & ")-月報表A!H" & i + 1 & "-月報表A!J" & i + 1 & ")+(SUMPRODUCT((MONTH(日報表B!A2:A" & DaySheetBLastRow & ")=" & i & ")*1,日報表B!L2:L" & DaySheetBLastRow & ")-月報表B!I" & i + 1 & "-月報表B!K" & i + 1 & "),0)"
                Next i
            
                '計算年度平均訂單量
                TotalMonth = 0
                
                For X = 2 To 13
                    ordersum = 0
                    ordersum = .Range("B" & X) + .Range("C" & X) + .Range("D" & X)
                    If ordersum > 0 Then
                        TotalMonth = TotalMonth + 1
                    End If
                Next X

                .Range("G2:G13") = Application.WorksheetFunction.Sum(.Range("B2:D13")) / TotalMonth
            
        End Select
        
    End With
    
End Sub
