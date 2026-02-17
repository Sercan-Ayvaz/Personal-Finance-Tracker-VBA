Attribute VB_Name = "modLogic"
Public Sub UpdateRecentActions(lb As MSForms.ListBox, wsName As String, colCount As Integer)
    Dim ws As Worksheet
    Dim i As Long, last As Long, rowCount As Integer
    
    Set ws = ThisWorkbook.Sheets(wsName)
    lb.Clear
    
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Exit Sub
    
    rowCount = 0
    For i = last To 2 Step -1
        If rowCount >= 10 Then Exit For
        
        lb.AddItem ws.Cells(i, 1).Text
        Dim c As Integer
        For c = 1 To colCount - 1
            lb.List(lb.ListCount - 1, c) = ws.Cells(i, c + 1).Text
        Next c
        
        rowCount = rowCount + 1
    Next i
End Sub

Public Sub ResetFormFields(formType As String, frm As Object)
    On Error Resume Next
        Select Case formType
            Case "Expenses"
                frm.txtEDate.Value = ""
                frm.txtEAmount.Value = ""
                frm.txtEComment.Value = ""
                frm.cbECategory.ListIndex = -1
                frm.cbEPaymentMethod.ListIndex = -1
            Case "Revenues"
                frm.txtRDate.Value = ""
                frm.txtRSource.Value = ""
                frm.txtRAmount.Value = ""
                
            Case "Investments"
                frm.txtEntityName.Value = ""
                frm.txtPurchasePrice.Value = ""
                frm.txtCurrentPrice.Value = ""
                frm.txtQuantity.Value = ""
                frm.txtIDate.Value = ""
        End Select
    On Error GoTo 0
End Sub

Public Sub FillComboBox(cb As MSForms.ComboBox, tblName As String, colName As String)
    Dim tbl As ListObject
    Dim cell As Range
    
    cb.Clear
    On Error Resume Next
    Set tbl = ThisWorkbook.Worksheets("calculation").ListObjects(tblName)
    On Error GoTo 0
    
    If Not tbl Is Nothing Then
        For Each cell In tbl.listColumns(colName).DataBodyRange
            If cell.Value <> "" Then
                cb.AddItem cell.Value
            End If
        Next cell
    End If
End Sub



