VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Product_ForDefects 
   Caption         =   "请选择1个产品用于搜索其缺陷列表 Please select one product for searching its defect list"
   ClientHeight    =   7504
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9821.001
   OleObjectBlob   =   "Product_ForDefects.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Product_ForDefects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_lev1_Click()
 Unload Me
 Exit Sub
End Sub

Private Sub Completed_Click()
Unload Me
'Exit Sub
End
End Sub

Private Sub Next_lev1_Click()

Dim N, Cn, i, j As Integer
Dim UF As UserForm
Dim Str As String


'====
Set UF = Me
    N = UF.Controls.count
    Cn = 0
     
 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
        If UF.Controls(i).Value = True Then
             Cn = Cn + 1
          End If
      End If
 Next i
 If Cn = 0 Then
     MsgBox "You have not selected any product, please select one product to search for its list of common defects." & CHR(10) & _
    "你未选择任何产品名称，请选择1个产品来搜索其常见的缺陷列表。" & CHR(10), vbCritical
    Load_Product_ForDefects
 ElseIf Cn > 1 Then
     MsgBox "You have not selected more than on product, each time, only one product allowed to select for searching its common defects，please re-select ." & CHR(10) & _
    "你选择了多个产品，每次只允许选择1个产品来查询其常见缺陷列表,请重新选择。" & CHR(10), vbCritical
    Load_Product_ForDefects
 End If
 
 
 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
        If UF.Controls(i).Value = True Then
             Defect_Product = UF.Controls(i).Name
             Str = UF.Controls(i).Caption
             Defect_Product = Mid(Str, InStr(1, Str, "|") + 1, InStrRev(Str, "|") - InStr(1, Str, "|") - 1)
             Defect_ProductE = Mid(Str, InStrRev(Str, "|") + 1, Len(Str) - InStrRev(Str, "|"))
             Defect_Product_ID = UF.Controls(i).Tag
          End If
      End If
 Next i
  
 ' Debug.Print Defect_Product & "====Defect_Product"
 ' Debug.Print Defect_ProductE & "====Defect_ProductE"
  'Debug.Print Defect_Product_ID & "==== Defect_Product_ID"
 
  Unload Me
  Load_Product_DefectList

End Sub
