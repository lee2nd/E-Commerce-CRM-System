Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim Shopee_book, Yahoo_book, Ruten_book As Workbook
    Dim Shopee_sheet, Yahoo_sheet, Ruten_sheet  As Worksheet
    
            If Me.opt1.Value Then
                'Open 蝦皮 Excel File
                With Application.FileDialog(msoFileDialogFilePicker)
                    .AllowMultiSelect = False
                    .Title = "選擇蝦皮資料"
                    .Filters.Clear
                    .Filters.Add "蝦皮", "*.xls;*.xlsx;*.csv"
                    If .Show = 0 Then
                        MsgBox "請選擇蝦皮資料", vbCritical
                        Exit Sub
                    Else
                        Shopee_Path = .SelectedItems(1)
                    End If
                End With
                
                Set Shopee_book = Workbooks.Open(Shopee_Path)
                Set Shopee_sheet = Shopee_book.Sheets(1)
                
                ShopeeColNum = Shopee_sheet.Range("A1").End(xlToRight).Column
                
                If ShopeeColNum = 48 Then
                    
                    With Shopee_sheet
                    
                        ShopeeLastRow = .Range("A65536").End(xlUp).Row
                        ShopeeOrdersLastRow = ThisWorkbook.Sheets("蝦皮orders").Range("A1048576").End(xlUp).Row
                        .Range("A2:AV" & ShopeeLastRow).Copy _
                        Destination:=ThisWorkbook.Sheets("蝦皮orders").Range("A" & ShopeeOrdersLastRow + 1)
                                       
                    End With
                                      
                    With ThisWorkbook.Sheets("Control Panel")

                        .Range("G3") = "蝦皮"
                        .Range("G3").VerticalAlignment = xlVAlignCenter
                        .Range("G3").HorizontalAlignment = xlCenter

                    End With
                    
                '2021/08/12 改版
                ElseIf ShopeeColNum = 50 Then
                                       
                    With Shopee_sheet
                    
                        ShopeeLastRow = .Range("A65536").End(xlUp).Row
                        ShopeeOrdersLastRow = ThisWorkbook.Sheets("蝦皮orders").Range("A1048576").End(xlUp).Row
                        .Columns("I:J").EntireColumn.Delete
                        .Range("A2:AU" & ShopeeLastRow).Copy _
                        Destination:=ThisWorkbook.Sheets("蝦皮orders").Range("A" & ShopeeOrdersLastRow + 1)
                                       
                    End With
        
                    With ThisWorkbook.Sheets("Control Panel")

                        .Range("G3") = "蝦皮"
                        .Range("G3").VerticalAlignment = xlVAlignCenter
                        .Range("G3").HorizontalAlignment = xlCenter

                    End With
                    
                Else
                
                    MsgBox "不符合蝦皮資料格式，請重新選擇"
                    
                End If

                Shopee_book.Close
                ThisWorkbook.Save
                
            ElseIf Me.opt2.Value Then
                'Open 雅虎Excel File
                With Application.FileDialog(msoFileDialogFilePicker)
                    .AllowMultiSelect = False
                    .Title = "選擇雅虎資料"
                    .Filters.Clear
                    .Filters.Add "雅虎", "*.xls;*.xlsx;*.csv"
                    If .Show = 0 Then
                        MsgBox "請選擇雅虎資料", vbCritical
                        Exit Sub
                    Else
                        Yahoo_Path = .SelectedItems(1)
                    End If
                End With
                
                Set Yahoo_book = Workbooks.Open(Yahoo_Path)
                Set Yahoo_sheet = Yahoo_book.Sheets(1)
                
                YahooColNum = Yahoo_sheet.Range("A1").End(xlToRight).Column
                
                If YahooColNum = 40 Then
                    
                    ThisWorkbook.Sheets("雅虎orders").Columns(23).EntireColumn.Delete
                    
                    With Yahoo_sheet
                    
                        YahooLastRow = .Range("A65536").End(xlUp).Row
                        YahooOrdersLastRow = ThisWorkbook.Sheets("雅虎orders").Range("A1048576").End(xlUp).Row
                        .Range("A2:AN" & YahooLastRow).Copy _
                        Destination:=ThisWorkbook.Sheets("雅虎orders").Range("A" & YahooOrdersLastRow + 1)
                                       
                    End With
                    
                    ThisWorkbook.Sheets("雅虎orders").Range("W:W").EntireColumn.Insert
                    ThisWorkbook.Sheets("雅虎orders").Range("W1") = "餘額部份支付金額"
                    
                    With ThisWorkbook.Sheets("Control Panel")
                    
                        .Range("G3") = "雅虎"
                        .Range("G3").VerticalAlignment = xlVAlignCenter
                        .Range("G3").HorizontalAlignment = xlCenter
                                          
                    End With
                    
                Else
                
                    MsgBox "不符合雅虎資料格式，請重新選擇"
                    
                End If

                Yahoo_book.Close
               ThisWorkbook.Save
                
            ElseIf Me.opt3.Value Then
                'Open 露天 Excel File
                With Application.FileDialog(msoFileDialogFilePicker)
                    .AllowMultiSelect = False
                    .Title = "選擇露天資料"
                    .Filters.Clear
                    .Filters.Add "露天", "*.xls;*.xlsx;*.csv"
                    If .Show = 0 Then
                        MsgBox "請選擇露天資料", vbCritical
                        Exit Sub
                    Else
                        Ruten_Path = .SelectedItems(1)
                    End If
                End With
                
                Set Ruten_book = Workbooks.Open(Ruten_Path)
                Set Ruten_sheet = Ruten_book.Sheets(1)
                
                RutenColNum = Ruten_sheet.Range("A1").End(xlToRight).Column
                
                If RutenColNum = 22 Then
                    
                    With Ruten_sheet
                    
                        RutenLastRow = .Range("A65536").End(xlUp).Row
                        RutenOrdersLastRow = ThisWorkbook.Sheets("露天orders").Range("A1048576").End(xlUp).Row
                        .Range("A2:V" & RutenLastRow).Copy _
                        Destination:=ThisWorkbook.Sheets("露天orders").Range("A" & RutenOrdersLastRow + 1)
                                       
                    End With
                    
                    With ThisWorkbook.Sheets("Control Panel")
                    
                        .Range("G3") = "露天"
                        .Range("G3").VerticalAlignment = xlVAlignCenter
                        .Range("G3").HorizontalAlignment = xlCenter
                                          
                    End With
                    
                '2021/09/15 改版
                ElseIf RutenColNum = 24 Then
                    
                    With Ruten_sheet
                    
                        RutenLastRow = .Range("A65536").End(xlUp).Row
                        RutenOrdersLastRow = ThisWorkbook.Sheets("露天orders").Range("A1048576").End(xlUp).Row
                        .Columns("P:P").EntireColumn.Delete
                        .Columns("C:C").EntireColumn.Delete
                        .Range("A2:V" & RutenLastRow).Copy _
                        Destination:=ThisWorkbook.Sheets("露天orders").Range("A" & RutenOrdersLastRow + 1)
                                       
                    End With
                    
                    With ThisWorkbook.Sheets("Control Panel")
                    
                        .Range("G3") = "露天"
                        .Range("G3").VerticalAlignment = xlVAlignCenter
                        .Range("G3").HorizontalAlignment = xlCenter
                                          
                    End With
                    
                 '2022/10/1 改版
                 ElseIf RutenColNum = 25 Then
                    
                    With Ruten_sheet
                    
                        RutenLastRow = .Range("A65536").End(xlUp).Row
                        RutenOrdersLastRow = ThisWorkbook.Sheets("露天orders").Range("A1048576").End(xlUp).Row
                        .Columns("Q:Q").EntireColumn.Delete
                        .Columns("N:N").EntireColumn.Delete
                        .Columns("C:C").EntireColumn.Delete
                        .Range("A2:V" & RutenLastRow).Copy _
                        Destination:=ThisWorkbook.Sheets("露天orders").Range("A" & RutenOrdersLastRow + 1)
                                       
                    End With
                    
                    With ThisWorkbook.Sheets("Control Panel")
                    
                        .Range("G3") = "露天"
                        .Range("G3").VerticalAlignment = xlVAlignCenter
                        .Range("G3").HorizontalAlignment = xlCenter
                                          
                    End With
                    
                Else
                
                    MsgBox "不符合露天資料格式，請重新選擇"
                    
                End If

                Ruten_book.Close
               ThisWorkbook.Save
                
            End If
            
            Unload Me
End Sub

Private Sub frmPlatform_Click()

End Sub

Private Sub opt1_Click()

End Sub

Private Sub opt2_Click()

End Sub

Private Sub opt3_Click()

End Sub

Private Sub UserForm_Click()

End Sub
