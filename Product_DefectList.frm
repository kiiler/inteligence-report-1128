VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Product_DefectList 
   Caption         =   "请选择缺陷（可多选） please select defect(multiple choise are alloewed)"
   ClientHeight    =   7504
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9821.001
   OleObjectBlob   =   "Product_DefectList.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Product_DefectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents cmb As MSForms.ComboBox
Attribute cmb.VB_VarHelpID = -1
Dim cmbN As Integer
Dim CmbItems() As String
Dim ItemsTag() As String

'Dim LastText As String
'Dim CurrentText As String


Private Sub Cancel_lev1_Click()
 Unload Me
 Exit Sub
End Sub

Private Sub Next_lev1_Click()
Dim N, i, j As Integer
Dim Cn As Integer
Dim ChID As String
Dim UF As UserForm
Dim Str As String
Dim N_Defect As Integer
Dim Defect() As String
Dim DefectE() As String
Dim DefectCode() As String
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

 
 ReDim Defect(1 To Cn)
 ReDim DefectE(1 To Cn)
 ReDim DefectCode(1 To Cn)
 Cn = 0
 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
        If UF.Controls(i).Value = True Then
             Cn = Cn + 1
             Defect(Cn) = UF.Controls(i).Name
             Str = UF.Controls(i).Caption
             Defect(Cn) = Mid(Str, InStr(1, Str, "|") + 1, InStrRev(Str, "|") - InStr(1, Str, "|") - 1)
             DefectE(Cn) = Mid(Str, InStrRev(Str, "|") + 1, Len(Str) - InStrRev(Str, "|"))
             DefectCode(Cn) = UF.Controls(i).Tag
           End If
      End If
 Next i


'===

 If Cn = 0 Then
     MsgBox "You have not selected any defect, please select at least one defect." & CHR(10) & _
    "你未选择任何缺陷，请至少选择一个缺陷。" & CHR(10), vbCritical
        Unload Me
        Call Load_Product_DefectList
 Else
   
        Unload Me
        Call Input_Defects(Cn, Defect, DefectE, DefectCode)
 End If
 
End Sub



Private Sub Next_Add_Click()
Dim N, i, j As Integer
Dim Cn As Integer
Dim ChID As String
Dim UF As UserForm
Dim Str As String
Dim N_Defect As Integer
Dim Defect() As String
Dim DefectE() As String
Dim DefectCode() As String
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

 
 ReDim Defect(1 To Cn)
 ReDim DefectE(1 To Cn)
 ReDim DefectCode(1 To Cn)
 Dim BK As bookmark
 Cn = 0
 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
        If UF.Controls(i).Value = True Then
             Cn = Cn + 1
             Defect(Cn) = UF.Controls(i).Name
             Str = UF.Controls(i).Caption
             Defect(Cn) = Mid(Str, InStr(1, Str, "|") + 1, InStrRev(Str, "|") - InStr(1, Str, "|") - 1)
             DefectE(Cn) = Mid(Str, InStrRev(Str, "|") + 1, Len(Str) - InStrRev(Str, "|"))
             DefectCode(Cn) = UF.Controls(i).Tag & "_" & "1"
             j = 1
             Do While ThisDocument.Bookmarks.Exists(UF.Controls(i).Tag & "_" & CStr(j))
              j = j + 1
              DefectCode(Cn) = UF.Controls(i).Tag & "_" & CStr(j)
            Loop
           
           End If
      End If
 Next i


'===

 If Cn = 0 Then
     MsgBox "You have not selected any defect, please select at least one defect." & CHR(10) & _
    "你未选择任何缺陷，请至少选择一个缺陷。" & CHR(10), vbCritical
        Unload Me
        Call Load_Product_DefectList
 Else
   
        Unload Me
        Call Input_Defects(Cn, Defect, DefectE, DefectCode)
        Unload Me
 End If
End Sub

Private Sub Next_Replace_Click()
Dim N, i, j As Integer
Dim Cn As Integer
Dim ChID As String
Dim UF As UserForm
Dim Str As String
Dim N_Defect As Integer
Dim Defect() As String
Dim DefectE() As String
Dim DefectCode() As String
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

 
 ReDim Defect(1 To Cn)
 ReDim DefectE(1 To Cn)
 ReDim DefectCode(1 To Cn)
 Cn = 0
 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
        If UF.Controls(i).Value = True Then
             Cn = Cn + 1
             Defect(Cn) = UF.Controls(i).Name
             Str = UF.Controls(i).Caption
             Defect(Cn) = Mid(Str, InStr(1, Str, "|") + 1, InStrRev(Str, "|") - InStr(1, Str, "|") - 1)
             DefectE(Cn) = Mid(Str, InStrRev(Str, "|") + 1, Len(Str) - InStrRev(Str, "|"))
             DefectCode(Cn) = UF.Controls(i).Tag & "_" & "1"
             j = 1
             Do While ThisDocument.Bookmarks.Exists(UF.Controls(i).Tag & "_" & CStr(j))
              j = j + 1
              DefectCode(Cn) = UF.Controls(i).Tag & "_" & CStr(j)
            Loop
           End If
      End If
 Next i


'===

 If Cn = 0 Then
     MsgBox "You have not selected any defect, please select at least one defect." & CHR(10) & _
    "你未选择任何缺陷，请至少选择一个缺陷。" & CHR(10), vbCritical
        Unload Me
        Call Load_Product_DefectList
 ElseIf Cn > 1 Then
     MsgBox "You have selected more than one defects, please select only one defect." & CHR(10) & _
    "你选择了不止一个缺陷，请只选择一个缺陷。" & CHR(10), vbCritical
     Load Me
 Else
   
 
        Unload Me
        Call Replace_Defects(Cn, Defect, DefectE, DefectCode)
 End If
End Sub

Private Sub PREVIOUS_Click()
Unload Me
Load_Product_Lev4
End Sub

Private Sub UserForm_Activate()
  
   Dim i, N As Integer
   N = Me.Controls.count - 1
   Cn = 0
   For i = 1 To N
     If TypeName(Me.Controls(i)) = "CheckBox" Then
        Cn = Cn + 1
        ReDim Preserve CmbItems(1 To Cn)
        ReDim Preserve ItemsTag(1 To Cn)
        CmbItems(Cn) = Me.Controls(i).Caption
        ItemsTag(Cn) = Me.Controls(i).Tag
     End If
   Next i
   cmbN = Cn
    Dim ScreenWidth, ScreenHeight As Long
    Dim UFwidth, UFheight As Long
    Dim Defwidth, Defheight As Long
    
    ScreenWidth = Application.Width
    ScreenHeight = Application.Height
    Defwidth = 0.136 * ScreenWidth            'ScreenHeight=1091
    Defheight = 0.036 * ScreenHeight
    TopPos = 0.05 * ScreenHeight               'ScreenHeight=650
    LeftPos = 0.018 * ScreenWidth
     
    ' Set cmb = Me.Controls("ComBoBox1")
     Set cmb = Me.Controls.add("Forms.ComBoBox.1", "ComBoBox1", True)
     With cmb
       ' .AddItem ""
            .List = CmbItems
            .Top = TopPos + Defheight + 2
            .Left = LeftPos
            .Width = 0.2 * ScreenWidth    '300
            .Height = Defheight '50
            .Visible = True
            .Enabled = True
            .Style = fmStyleDropDownCombo
            .ListRows = 20
            .DropDown
    End With
    LastText = cmb.Text
    IsUserSelect = False
    'Set cmb = Me.Controls("cmb")
  
    
End Sub

Private Sub cmb_Change()

   Static InChange As Boolean
   If InChange Then Exit Sub
   InChange = True
   
   Dim KeyWords As String
   Dim Cn As Integer
   Dim SearchResult() As String
   Dim Ctrl As Control
   
   
       CurrentText = cmb.Text
       KeyWords = LCase(CurrentText)
   If LastText <> KeyWords Then
       Cn = 0
       cmb.Clear
      '---
       SearchResult = GetSearchResult(KeyWords)
        Cn = UBound(SearchResult)
        
        If Cn > 0 Then
            cc = cc + 1
            cmb.List = SearchResult
            cmb.Text = KeyWords ' 保持当前文本
            On Error Resume Next
              cmb.DropDown
            On Error GoTo 0
        End If
      '---
  
10    End If
    
    InChange = False ' 重置标志
    LastText = KeyWords
   
End Sub
Private Sub cmb_click()
       Dim DefName As String
       Dim DefNameE As String
       Dim DefCode As String
       Dim Str As String
       Dim i, j As Integer
       If cmb.ListIndex >= 0 Then
        ' 将选中的值设置为ComboBox的文本
          cmb.Text = cmb.List(cmb.ListIndex)
          For i = 1 To cmbN
            If cmb.Text = CmbItems(i) Then
                DefCode = ItemsTag(i)
                Str = cmb.Text
                GoTo 10
            End If
          Next i
10         j = 1
             Do While ThisDocument.Bookmarks.Exists(DefCode & "_" & CStr(j))
              j = j + 1
              DefCode = DefCode & "_" & CStr(j)
            Loop
             
             DefName = Mid(Str, InStr(1, Str, "|") + 1, InStrRev(Str, "|") - InStr(1, Str, "|") - 1)
             DefNameE = Mid(Str, InStrRev(Str, "|") + 1, Len(Str) - InStrRev(Str, "|"))
             
             Call InputDropDown_Defects(DefName, DefNameE, DefCode)

         
        ' 退出程序
          Unload Me
         Exit Sub
     
      End If
End Sub


Private Function GetSearchResult(KeyWords As String) As String()
    ' 这里应该是筛选逻辑，以下为示例代码
    Dim SearchResult() As String
    For Each Ctrl In Me.Controls
         If TypeName(Ctrl) = "CheckBox" Then
            If LCase(Ctrl.Caption) Like "*" & KeyWords & "*" Then
               Cn = Cn + 1
               ReDim Preserve SearchResult(1 To Cn)
               SearchResult(Cn) = Ctrl.Caption
            End If
         End If
       Next Ctrl
    
    GetSearchResult = SearchResult
End Function


