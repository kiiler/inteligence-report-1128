VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Product_Features 
   Caption         =   "Please Select Product Features"
   ClientHeight    =   7504
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9821.001
   OleObjectBlob   =   "Product_Features.frx":0000
   StartUpPosition =   1  'Owner Center
End
Attribute VB_Name = "Product_Features"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_lev1_Click()
 Unload Me
 Exit Sub
End Sub

Private Sub Next_Lev5_Click()
Dim N, Cn As Integer
Dim UF As UserForm
Dim ChID As String
 
 Set UF = Me
 N = Me.Controls.count
 
 Erase ChT_ID
 
  For j = 1 To N_Cat_Large
    For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
        If Me.Controls(i).Value = True Then
             If Mid(Me.Controls(i).Tag, 1, InStr(1, Me.Controls(i).Tag, "|") - 1) = Cat_Large_ID(j) Then
               'cn = cn + 1
               'CHAR(cn) = Me.Controls(i).Caption
              ' CHAR(cn) = Mid(CHAR(cn), InStr(1, CHAR(cn), "|") + 1, Len(CHAR(cn)) - InStr(1, CHAR(cn), "|"))
              
               ChID = Right(Me.Controls(i).Tag, Len(Me.Controls(i).Tag) - InStrRev(Me.Controls(i).Tag, "|", Len(Me.Controls(i).Tag)))
               'Product_Feature_ID(cn) = Uf.Controls(i).Tag
               ChT_ID(j) = ChID & "," & ChT_ID(j)
             End If
        End If
    End If
  Next i
  
   If Len(ChT_ID(j)) > 1 Then
     ChT_ID(j) = Mid(ChT_ID(j), 1, Len(ChT_ID(j)) - 1)
   End If
Next j

     'ChN = cn
 
     Me.Hide
   'Unload Me
  Call Get_CheckPoints
 ' Call Get_ProductFeatures
End Sub
Private Sub PREVIOUS_Click()
 Unload Me
 Load_Product_Lev4
End Sub

