VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Product_PrimaryTypeCheckPoints 
   Caption         =   "Please select checkpoints"
   ClientHeight    =   7504
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9821.001
   OleObjectBlob   =   "Product_PrimaryTypeCheckPoints.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Product_PrimaryTypeCheckPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_lev1_Click()
Unload Me
Exit Sub
End Sub

Private Sub Next_lev1_Click()
Dim i, j As Integer
Dim N As Integer
Dim Cn As Integer
Dim Un As Integer
Dim UF As UserForm
Dim Str As String
Dim AddcChkCode() As String
'====
Set UF = Me
    Un = UF.Controls.count
    Cn = 0

 For i = 0 To Un - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
             Cn = Cn + 1
      End If
 Next i

ReDim AddchkCode(1 To Cn)
 N = 0
 For i = 0 To Un - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
        If UF.Controls(i).Value = True Then
             N = N + 1
             AddchkCode(N) = UF.Controls(i).Tag
           End If
      End If
 Next i


'===

 If Cn = 0 Then
     MsgBox "You have not selected any Checkpoint, please select at least one Checkpoint or click cancel to exit." & CHR(10) & _
    "你未选择任何测试项目，请至少选择一个测试项目或者点Cancel退出。" & CHR(10), vbCritical
        Unload Me
        Call Load_Product_PrimaryTypeCheckPoints
 Else
   
        Unload Me
        Call Input_PrimaryTypeCheckPoint(Cn, N, AddchkCode())
 End If
End Sub
