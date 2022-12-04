Sub MonthA()

    Dim DaySheetA, MonthAsheet As Worksheet
    Set DaySheetA = ThisWorkbook.Sheets("日報表A")
    Set MonthAsheet = ThisWorkbook.Sheets("月報表A")

        With DaySheetA
            
            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodJanuary
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)
            
            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 1
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With
            
            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodFebruray
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)

            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 2
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With

            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodMarch
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)
            
            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 3
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With

            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodApril
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)
            
            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 4
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With

            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodMay
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)
            
            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 5
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With

            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodJune
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)
            
            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 6
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With

            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodJuly
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)
            
            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 7
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With

            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodAugust
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)
            
            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 8
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With

            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodSeptember
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)

            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 9
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With
            
            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodOctober
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)
            
            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 10
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With
            
            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodNovember
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)
            
            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 11
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With
            
            TotalRevenue = 0
            TotalCost = 0
            .UsedRange.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterAllDatesInPeriodDecember
            TotalRevenue = Application.WorksheetFunction.Sum(.Range("D:D").SpecialCells(xlCellTypeVisible))
            TotalCost = Application.WorksheetFunction.Sum(.Range("K:K").SpecialCells(xlCellTypeVisible))
            TotalDiscount = Application.WorksheetFunction.Sum(.Range("E:G").SpecialCells(xlCellTypeVisible))
            TotalFee = Application.WorksheetFunction.Sum(.Range("H:J").SpecialCells(xlCellTypeVisible))
            
            TotalRevenue = WorksheetFunction.Round(TotalRevenue, 0)
            TotalCost = WorksheetFunction.Round(TotalCost, 0)
            TotalDiscount = WorksheetFunction.Round(TotalDiscount, 0)
            TotalFee = WorksheetFunction.Round(TotalFee, 0)
            
            With MonthAsheet
                    MonthAsheetLastRow = .Range("A1048576").End(xlUp).Row
                    .Cells(MonthAsheetLastRow + 1, 1) = 12
                    .Cells(MonthAsheetLastRow + 1, 2) = TotalRevenue
                    .Cells(MonthAsheetLastRow + 1, 3) = TotalCost
                    .Cells(MonthAsheetLastRow + 1, 4) = TotalDiscount
                    .Cells(MonthAsheetLastRow + 1, 5) = TotalFee
            End With
            
        End With
        
        With MonthAsheet
        
            '刪除重複資料
            .UsedRange.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5), Header:=xlYes
            
            '調整字體
            .Cells.Font.Size = 11
            .Cells.Font.Name = "微軟正黑體"
            .Cells.VerticalAlignment = xlVAlignCenter
            .Cells.HorizontalAlignment = xlHAlignLeft
    
            '自動調整欄寬
            .Columns("A:C").AutoFit
        
        End With
        
        DaySheetA.AutoFilterMode = False
        
        ThisWorkbook.Sheets("Control Panel").Activate
        MsgBox "Complete!"

End Sub
