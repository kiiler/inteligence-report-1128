VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} selectproduct 
   Caption         =   "Select Product Category"
   ClientHeight    =   7504
   ClientLeft      =   5040
   ClientTop       =   3374
   ClientWidth     =   11816
   OleObjectBlob   =   "selectproduct.frx":0000
End
Attribute VB_Name = "selectproduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CheckBox1_Click()

End Sub

Private Sub Cancel_lev1_Click()
Unload Me
End Sub


  Private Sub ee_Click()
If EE = True Then ic.Value = 0: it.Value = 0
 End Sub
  Private Sub ic_Click()
If ic = True Then EE.Value = 0: it.Value = 0
 End Sub
  Private Sub it_Click()
If it = True Then ic.Value = 0: EE.Value = 0
 End Sub

Private Sub Next_lev1_Click()
Dim N, Cn As Integer
Dim FR(10) As String
  'Unload Me
  'MsgBox Energy_lev2.
   'Load Energy_lev2
   'Energy_lev2.Show
N = Level1.Controls.count
 For i = 0 To N - 1
   If Level1.Controls(i).Value = True Then
     Cn = Cn + 1
     FR(1) = Level1.Controls(i).Name
     
   End If
 Next i
 
 If Cn = 0 Then
     MsgBox "You have not selected any product categories, the program will exit, please re-run the program if you need" & CHR(10) & _
    "你未选择任何产品类型，程序将退出，如果需要请你重新运行程序。" & CHR(10), vbCritical
     Exit Sub
     
 ElseIf Cn > 1 Then
    MsgBox "You have selected " & Cn & " product categories, only one product categories allowed to be selected each time,please re-select." & CHR(10) & _
    "你选择了" & Cn & "产品类型，注意每次只允许选择一种产品类型，请重新选择！" & CHR(10), vbCritical
    Exit Sub
 Else
   Unload Me
   Cat_Large = FR(1)
   Select Case FR(1)
      Case "Electrical"
          Electrical_lev2.Show
      Case "Softgoods"
          Softgoods_lev2.Show
      Case "Hardgoods"
           Hardgoods_lev2.Show
      Case "Mechanical"
           Mechanical_lev2.Show
      Case "Oil"
             Oil_lev2.Show
      Case "Transportation"
             Transportation_lev2.Show
      Case "Building"
            Building_lev2.Show
      Case "Energy"
            Energy_lev2.Show
      Case "Mineral"
            Mineral_lev2.Show
      Case "Chemical"
           Chemical_lev2.Show
      Case "Medical"
            Medical_lev2.Show
      Case "Agriculture"
            Agriculture_lev2.Show
      Case Else
           Exit Sub
    End Select
   
 End If
 
 


End Sub

