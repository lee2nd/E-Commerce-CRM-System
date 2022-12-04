Sub SingleItemSearch()

    Dim DeliverySheet, VisualizationSheet As Worksheet
    Set DeliverySheet = ThisWorkbook.Sheets("出庫")
    Set VisualizationSheet = ThisWorkbook.Sheets("圖表")
    
        '先找出 Unique 的商品
        With DeliverySheet
            .Activate
            .Columns("AZ:AZ").Clear
            DeliverySheetLastRow = .Range("A1048576").End(xlUp).Row
            
            For i = 2 To DeliverySheetLastRow
                .Range("AZ" & i) = "(" & .Range("A" & i) & ")" & .Range("B" & i)
                If .Range("AZ" & i) Like "*TBD*" Then
                    .Range("AZ" & i).Select
                    Selection.Delete Shift:=xlUp
                End If
            Next i
            
            .Range("AZ2:AZ" & DeliverySheetLastRow).RemoveDuplicates Columns:=1, Header:=no
            .Range("AZ2:AZ" & DeliverySheetLastRow).Sort Key1:=Range("AZ2"), Order1:=xlAscending, Header:=no
            
        End With

        '製作出庫的下拉式選單
        VisualizationSheet.Activate
        VisualizationSheet.Range("D32").Select
        DeliverySheetLastRow = DeliverySheet.Range("AZ1048576").End(xlUp).Row

        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="='出庫'!$AZ$2:$AZ$" & DeliverySheetLastRow
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
            
End Sub
