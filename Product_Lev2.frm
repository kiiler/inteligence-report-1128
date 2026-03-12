VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Product_Lev2 
   Caption         =   "请选择产品中类 Please Select secondary Product Type"
   ClientHeight    =   8505.001
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9821.001
   OleObjectBlob   =   "Product_Lev2.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Product_Lev2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_lev1_Click()
 Unload Me
 Exit Sub
End Sub

Private Sub Next_lev2_Click()
Dim i, N, Cn As Integer
Dim FR, FRid As String
Dim Url As String
Dim UF As UserForm
Dim postJson As Object
Set UF = Me
    N = UF.Controls.count
    Cn = 0
    FR = ""
 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
        If UF.Controls(i).Value = True Then
           Cn = Cn + 1
          Cat_Middle(Cn) = UF.Controls(i).Name
          Cat_Middle_ID(Cn) = UF.Controls(i).Tag
          Cat_MiddleE(Cn) = Mid(UF.Controls(i).Caption, InStrRev(UF.Controls(i).Caption, "|") + 1)
        End If
      End If
 Next i
  N_Cat_Middle = Cn
 If Cn = 0 Then
     MsgBox "You have not selected any product categories, the program will exit, please re-run the program if you need" & CHR(10) & _
    "你未选择任何产品类型，程序将退出，如果需要请你重新运行程序。" & CHR(10), vbCritical
     Unload Me
     Load_Product_Lev2
     Exit Sub
 ElseIf Cn > 10 Then
    MsgBox "You have selected " & Cn & " product categories, no more than 10 product categories allowed to be selected each time,please re-select." & CHR(10) & _
    "你选择了" & Cn & "产品类型，注意每次只允许选择不多于10种产品类型，请重新选择！" & CHR(10), vbCritical
     Unload Me
     Load_Product_Lev2
    Exit Sub
 Else
   
   Unload Me
   Call Load_Product_Lev3
 End If
End Sub


Private Sub PREVIOUS_Click()
 Unload Me
 Load_Product_Lev1
End Sub
