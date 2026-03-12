VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Product_Lev1 
   Caption         =   "请选择产品大类 Please Select Primary Product Type"
   ClientHeight    =   8505.001
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9821.001
   OleObjectBlob   =   "Product_Lev1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Product_Lev1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddCheckPoints_Click()
Dim i, N, Cn As Integer
Dim UF As UserForm
    Set UF = Me
    N = UF.Controls.count
    Cn = 0
   
 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
        If UF.Controls(i).Value = True Then
           Cn = Cn + 1
           Cat_Large(Cn) = UF.Controls(i).Name
           Cat_Large_ID(Cn) = UF.Controls(i).Tag
           Cat_LargeE(Cn) = Mid(UF.Controls(i).Caption, InStrRev(UF.Controls(i).Caption, "|") + 1)
        End If
      End If
 Next i


 If Cn = 0 Then
     MsgBox "You have not selected any product categories, the program will exit, please re-run the program if you need" & CHR(10) & _
    "你未选择任何产品类型，程序将退出，如果需要请你重新运行程序。" & CHR(10), vbCritical
     Unload Me
     Load_Product_Lev1
     Exit Sub
 ElseIf Cn > 1 Then
    MsgBox "You have selected " & Cn & " product categories, for getting all check points of primary type,only one product categories allowed to be selected,please re-select." & CHR(10) & _
    "你选择了" & Cn & "产品类型，对于查询大类产品的所有测试项目，注意每次只允许1种产品大类，请重新选择！" & CHR(10), vbCritical
     Unload Me
    Exit Sub
 Else
    Unload Me
    Call Load_Product_PrimaryTypeCheckPoints
 End If
End Sub

Private Sub Cancel_lev1_Click()
 Unload Me
 Exit Sub
End Sub

Private Sub Next_lev1_Click()
Dim i, N, Cn As Integer
Dim FR, FRid As String
Dim Url As String
Dim UF As UserForm
'Dim PostJson As Object
Dim postJson As Object
  

    Set UF = Me
    N = UF.Controls.count
    Cn = 0
    FR = ""
 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
        If UF.Controls(i).Value = True Then
           Cn = Cn + 1
           Cat_Large(Cn) = UF.Controls(i).Name
           Cat_Large_ID(Cn) = UF.Controls(i).Tag
           Cat_LargeE(Cn) = Mid(UF.Controls(i).Caption, InStrRev(UF.Controls(i).Caption, "|") + 1)
           
        End If
      End If
 Next i
 N_Cat_Large = Cn

 If Cn = 0 Then
     MsgBox "You have not selected any product categories, the program will exit, please re-run the program if you need" & CHR(10) & _
    "你未选择任何产品类型，程序将退出，如果需要请你重新运行程序。" & CHR(10), vbCritical
     Unload Me
     Load_Product_Lev1
     Exit Sub
 ElseIf Cn > 3 Then
    MsgBox "You have selected " & Cn & " product categories, No more than 3 product categories allowed to be selected each time,please re-select." & CHR(10) & _
    "你选择了" & Cn & "产品类型，注意每次只允许选择不多于3种产品类型，请重新选择！" & CHR(10), vbCritical
     Unload Me
    Exit Sub
 Else
    Unload Me
    Call Load_Product_Lev2
 End If
End Sub

Private Sub Next_lev1_Click_Back()
Dim i, N, Cn As Integer
Dim FR, FRid As String
Dim Url As String
Dim UF As UserForm
'Dim PostJson As Object
Dim postJson As Object
  
    Set UF = Me
    N = UF.Controls.count
    Cn = 0
    FR = ""
 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
        If UF.Controls(i).Value = True Then
           Cn = Cn + 1
           Cat_Large = UF.Controls(i).Name
           Cat_Large_ID = UF.Controls(i).Tag
        End If
      End If
 Next i
 If Cn = 0 Then
     MsgBox "You have not selected any product categories, the program will exit, please re-run the program if you need" & CHR(10) & _
    "你未选择任何产品类型，程序将退出，如果需要请你重新运行程序。" & CHR(10), vbCritical
     Unload Me
     Exit Sub
 ElseIf Cn > 1 Then
    MsgBox "You have selected " & Cn & " product categories, only one product categories allowed to be selected each time,please re-select." & CHR(10) & _
    "你选择了" & Cn & "产品类型，注意每次只允许选择一种产品类型，请重新选择！" & CHR(10), vbCritical
     Unload Me
    Exit Sub
 Else
    Unload Me
    Load_Prodoct_Lev2
    Set postJson = New Scripting.Dictionary
    postJson("szdl") = Cat_Large_ID
    Url = WebBaseUrl & "/api/sqlhub/exec?appid=md_19&configName=getProductSecondaryType"
    Call getHUBdata(Url, postJson, jsonResponse)
    
 Debug.Print jsonResponse("data")(1)("产品中类中文")
    Dim Data As Object
    Dim Cat_Middle_IDs() As String
    Dim Cat_Middle_Codes() As String
    Dim Cat_Middle_Names() As String
    Dim Cat_Middle_NamesE() As String
    Dim Cat_Middle_Number As Integer
    'Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
    Set Data = jsonResponse("data")
    Cat_Middle_Number = Data.count
    ReDim Cat_Middle_IDs(1 To Cat_Middle_Number)
    ReDim Cat_Middle_Codes(1 To Cat_Middle_Number)
    ReDim Cat_Middle_Names(1 To Cat_Middle_Number)
    ReDim Cat_Middle_NamesE(1 To Cat_Middle_Number)
    For i = 1 To Cat_Middle_Number
          Cat_Middle_IDs(i) = Data(i)("产品中类id")
          Cat_Middle_Codes(i) = Data(i)("产品中类编码")
          Cat_Middle_Names(i) = Data(i)("产品中类中文")
          Cat_Middle_NamesE(i) = Data(i)("产品中类英文")
    Next i
   
    Set UF = New Product_Lev2
    Dim LeftPos As Integer
    Dim TopPos As Integer
    Dim Chk As Object
    
    LeftPos = 20 ' 设置复选框的初始Left位置
    TopPos = 10
    For i = 1 To Cat_Middle_Number
        Set Chk = UF.Controls.add("Forms.CheckBox.1", Cat_Middle_Names(i), True)
        With Chk
            .Caption = i & "|" & Cat_Middle_Names(i) & "|" & Cat_Middle_NamesE(i)
            .Tag = Cat_Middle_IDs(i)
            .Left = LeftPos
            .Top = 18 * i + TopPos
            .Width = 300 ' 设置复选框的宽度
            .Height = 18
            .Visible = True
           
            If Cat_Middle_IDs(i) = Cat_Middle_ID Then
              .Value = True
            End If
        End With
    Next i
    Unload Me
    NewMacros.LoadProdoct_Lev2
    'Product_Lev2.Show 'vbModal
 End If
End Sub

