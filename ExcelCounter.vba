Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lookupCell As Range
    Dim productName As String
    Dim lastRowC As Long
    Dim foundCell As Range
    
    ' Only run if a cell in Column B is changed
    If Not Intersect(Target, Me.Range("B:B")) Is Nothing Then
        productName = Target.Value
        
        ' Check if the product already exists in Column C
        Set foundCell = Me.Range("C:C").Find(productName, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            ' Product already exists, update the count
            foundCell.Offset(0, 1).Value = foundCell.Offset(0, 1).Value + 1
        Else
            ' Product not found, add it to the first empty cell in Column C
            lastRowC = Me.Cells(Me.Rows.Count, 3).End(xlUp).Row + 1
            Me.Cells(lastRowC, 3).Value = productName
            Me.Cells(lastRowC, 4).Value = 1
        End If
    End If
End Sub

#Save your workbook as a Macro-Enabled Excel File (.xlsm).

