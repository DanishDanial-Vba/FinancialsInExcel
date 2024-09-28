VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_home 
   Caption         =   "Home"
   ClientHeight    =   2920
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8350.001
   OleObjectBlob   =   "frm_home.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
#Const DebugMode = True                          ' Set to True to enable debugging

Private Sub btn_Process_Click()
   
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim isValid As Boolean
    Set ws = Application.Workbooks(list_wb.value).Worksheets(list_ws.value)
       
    'X------------------------------------------------------------X
   
   
    
    Set wb = Application.Workbooks(list_wb.value)
    
    isValid = ValidateTrialBalance(ws)
    If isValid Then
        
        'GroupAndPrintStatements wb
        StoreTrialBalance ws, wb
            
        'GenerateErrorReport wb
        
        Set ws = Nothing
        Set wb = Nothing
        Exit Sub
    Else
        MsgBox "Not Okay"
        Exit Sub
    End If

    
End Sub

Private Sub UserForm_Initialize()
    list_wb.Clear
    list_ws.Clear
    
    List_WorkBooks_Combobox list_wb
    
    
    
    
    
End Sub

Private Sub list_wb_Change()
    Dim wb As Workbook
    Set wb = Application.Workbooks(list_wb.value)
    
    List_WorkSheets_Combobox wb, list_ws
    
    
    
End Sub

'2. Create a Separate Memory Release Method
'This method will clean up memory by explicitly setting all object references to Nothing and clearing the clsTrialBalance and clsAccountGrouping collections.


Private Sub UserForm_Terminate()
    
End Sub

