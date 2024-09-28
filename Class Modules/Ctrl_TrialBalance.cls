VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ctrl_TrialBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Controller"
Option Explicit

Private pAccounts As Collection

Private Sub Class_Initialize()
    Set pAccounts = New Collection
End Sub

Public Sub AddAccount(ByVal account As Mdl_Accounts)
    pAccounts.Add account
End Sub

Public Property Get accounts() As Collection
    Set accounts = pAccounts

End Property

Public Sub RemoveAccount(ByVal account As Mdl_Accounts)
    Dim i As Long
    Dim found As Boolean
    found = False

    ' Loop through the collection to find the account
    For i = 1 To pAccounts.Count
        ' Assuming Mdl_Accounts has a property called ID to identify accounts
        If pAccounts(i).accountName = account.accountName Then
            pAccounts.Remove i
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        MsgBox "Account not found."
    End If
End Sub


Private Sub Class_Terminate()
    Dim i As Long
        
    If Not pAccounts Is Nothing Then
        For i = pAccounts.Count To 1 Step -1
            pAccounts.Remove i
        Next i
                    
        Set pAccounts = Nothing
    End If
End Sub

