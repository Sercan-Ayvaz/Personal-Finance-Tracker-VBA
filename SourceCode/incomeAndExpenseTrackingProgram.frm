VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} incomeAndExpenseTrackingProgram 
   Caption         =   "Income And Expense Tracking Program"
   ClientHeight    =   9990.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7860
   OleObjectBlob   =   "incomeAndExpenseTrackingProgram.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "incomeAndExpenseTrackingProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As Worksheet


Private Sub UserForm_Activate()
    obtnExpenses.Value = False
    obtnInvestments.Value = False
    obtnRevenues.Value = False
    
    fExpenses.Visible = False
    fRevenues.Visible = False
    fInvestments.Visible = False
End Sub



Private Sub obtnExpenses_Click()
    Dim i As Long, last As Long, rowCount As Integer
        
    If obtnExpenses.Value = True Then
        Set ws = ThisWorkbook.Sheets("tbl_Expenses")
        
        fExpenses.Visible = True
        fRevenues.Visible = False
        fInvestments.Visible = False
    End If
    
    Call FillComboBox(cbECategory, "tblVC", "Category")
    Call FillComboBox(cbEPaymentMethod, "tblPM", "Payment Methods")
    
    Call UpdateRecentActions(lstExpensesRecentAction, "tbl_Expenses", 6)
End Sub
Private Sub obtnRevenues_Click()
    Dim i As Long, last As Long, rowCount As Integer
    
    If obtnRevenues.Value = True Then
        Set ws = ThisWorkbook.Sheets("tbl_Revenues")
        fExpenses.Visible = False
        fRevenues.Visible = True
        fInvestments.Visible = False
    End If
    
    Call UpdateRecentActions(lstRevenuesRecentAction, "tbl_Revenues", 3)
End Sub
Private Sub obtnInvestments_Click()
    Dim i As Long, last As Long, rowCount As Integer
    If obtnInvestments.Value = True Then
        Set ws = ThisWorkbook.Sheets("tbl_Investments")
        
        fExpenses.Visible = False
        fRevenues.Visible = False
        fInvestments.Visible = True
        
        Call UpdateRecentActions(lstInvestmentsRecentActions, "tbl_Investments", 5)
    End If
End Sub

Private Sub cbtnExpensesSave_Click()
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If txtEDate.Value = "" And cbECategory.Value = "" And txtEAmount.Value = "" And cbEPaymentMethod.Value = "" Then
        MsgBox "There are blank spaces.", vbInformation
        Exit Sub
    End If
    If Not IsDate(txtEDate.Value) Then
        MsgBox "The date can only take a date value.(Example: 12.12.2000)", vbInformation
        Exit Sub
    End If
    If Not IsNumeric(txtEAmount.Value) Then
        MsgBox "The amount can only take a numerical value.", vbInformation
        Exit Sub
    End If
    
    ws.Cells(lastRow, 1).Value = Format(CDate(txtEDate.Value), "dd.mm.yyyy")
    ws.Cells(lastRow, 2).Value = cbECategory.Value
    ws.Cells(lastRow, 3).Value = txtEComment.Value
    ws.Cells(lastRow, 4).Value = CCur(txtEAmount.Value)
    ws.Cells(lastRow, 5).Value = cbEPaymentMethod.Value
    
    MsgBox "Data successfully saved.", vbInformation
    
    ActiveWorkbook.RefreshAll
    Call UpdateRecentActions(lstExpensesRecentAction, "tbl_Expenses", 6)
    Call ResetFormFields("Expenses", Me)
End Sub


Private Sub cbtnRevenuesSave_Click()
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    If txtRDate.Value = "" And txtRSource.Value = "" And txtRAmount.Value = "" Then
        MsgBox "There are blank spaces.", vbInformation
        Exit Sub
    End If
    If Not IsDate(txtRDate.Value) Then
        MsgBox "The date can only take a date value.(Example: 12.12.2000)", vbInformation
        Exit Sub
    End If
    If Not IsNumeric(txtRAmount.Value) Then
        MsgBox "The amount can only take a numerical value.", vbInformation
        Exit Sub
    End If
    
    ws.Cells(lastRow, 1).Value = Format(CDate(txtRDate.Value), "dd.mm.yyyy")
    ws.Cells(lastRow, 2).Value = txtRSource.Value
    ws.Cells(lastRow, 3).Value = CCur(txtRAmount.Value)
        
    MsgBox "Data successfully saved.", vbInformation
    
    ActiveWorkbook.RefreshAll
    Call UpdateRecentActions(lstRevenuesRecentAction, "tbl_Revenues", 3)
    Call ResetFormFields("Revenues", Me)
    
End Sub

Private Sub cbtnISave_Click()
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If txtEntityName.Value = "" And txtPurchasePrice.Value = "" And txtCurrentPrice.Value = "" And txtQuantity.Value = "" And txtIDate.Value = "" Then
        MsgBox "There are blank spaces.", vbInformation
        Exit Sub
    End If
    If Not IsNumeric(txtPurchasePrice.Value) Then
        MsgBox "The purchase price can only take a numerical value.", vbInformation
        Exit Sub
    End If
    If Not IsNumeric(txtCurrentPrice.Value) Then
        MsgBox "The current price can only take a numerical value.", vbInformation
        Exit Sub
    End If
    If Not IsNumeric(txtQuantity.Value) Then
        MsgBox "The quantity can only take a numerical value.", vbInformation
        Exit Sub
    End If
    If Not IsDate(txtIDate.Value) Then
        MsgBox "The date can only take a date value.(Example: 12.12.2000)", vbInformation
        Exit Sub
    End If
    
    ws.Cells(lastRow, 1).Value = Format(CDate(txtIDate.Value), "dd.mm.yyyy")
    ws.Cells(lastRow, 2).Value = txtEntityName.Value
    ws.Cells(lastRow, 3).Value = CCur(txtPurchasePrice.Value)
    ws.Cells(lastRow, 4).Value = CCur(txtCurrentPrice.Value)
    ws.Cells(lastRow, 5).Value = FormatNumber(CInt(txtQuantity.Value), 0)
    
    MsgBox "Data successfully saved.", vbInformation

    ActiveWorkbook.RefreshAll
    Call UpdateRecentActions(lstInvestmentsRecentActions, "tbl_Investments", 5)
    Call ResetFormFields("Investments", Me)
    
End Sub




