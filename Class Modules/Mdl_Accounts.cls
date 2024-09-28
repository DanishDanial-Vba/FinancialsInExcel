VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mdl_Accounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Models"
Option Explicit

Private pAccountCategory As String
Private pAccountSubCategory As String
Private pAccountSubCategoryL1 As String
Private pAccountSubCategoryL2 As String
Private pAccountName As String
Private pDebit As Double
Private pCredit As Double

Public Property Let accountCategory(ByVal value As String)
    pAccountCategory = value
End Property

Public Property Get accountCategory() As String
    accountCategory = pAccountCategory
End Property

Public Property Let accountSubCategory(ByVal value As String)
    pAccountSubCategory = value
End Property

Public Property Get accountSubCategory() As String
    accountSubCategory = pAccountSubCategory
End Property

Public Property Let accountSubCategoryL1(ByVal value As String)
    pAccountSubCategoryL1 = value
End Property

Public Property Get accountSubCategoryL1() As String
    accountSubCategoryL1 = pAccountSubCategoryL1
End Property

Public Property Let accountSubCategoryL2(ByVal value As String)
    pAccountSubCategoryL2 = value
End Property

Public Property Get accountSubCategoryL2() As String
    accountSubCategoryL2 = pAccountSubCategoryL2
End Property

Public Property Let accountName(ByVal value As String)
    pAccountName = value
End Property

Public Property Get accountName() As String
    accountName = pAccountName
End Property

Public Property Let Debit(ByVal value As Double)
    pDebit = value
End Property

Public Property Get Debit() As Double
    Debit = pDebit
End Property

Public Property Let Credit(ByVal value As Double)
    pCredit = value
End Property

Public Property Get Credit() As Double
    Credit = pCredit
End Property



Private Sub Class_Initialize()
    pAccountName = vbNullString
    pAccountCategory = vbNullString
    pAccountSubCategory = vbNullString
    pAccountSubCategoryL1 = vbNullString
    pAccountSubCategoryL2 = vbNullString
    pDebit = 0
    pDebit = 0
End Sub

