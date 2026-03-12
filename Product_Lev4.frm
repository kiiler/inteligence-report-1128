VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Product_Lev4 
   Caption         =   "Please Select Product Name"
   ClientHeight    =   8505.001
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9821.001
   OleObjectBlob   =   "Product_Lev4.frx":0000
   StartUpPosition =   1  'Owner Center
End
Attribute VB_Name = "Product_Lev4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_lev1_Click()
 Unload Me
 'Exit Sub
 End
End Sub

Private Sub Next_lev4_Click()

Dim N, Cn, i, j As Integer

Dim ChID As String
Dim UF As UserForm
Dim Str As String
  
'====
Set UF = Me
    N = UF.Controls.count
    Cn = 0

 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
       
        If UF.Controls(i).Value = True Then
          If UF.Controls(i).Tag <> "Choose Product Features" Then
             Cn = Cn + 1
             Product(Cn) = UF.Controls(i).Name
             Str = UF.Controls(i).Caption
             Product(Cn) = Mid(Str, InStr(1, Str, "|") + 1, InStrRev(Str, "|") - InStr(1, Str, "|") - 1)
             ProductE(Cn) = Mid(Str, InStrRev(Str, "|") + 1, Len(Str) - InStrRev(Str, "|"))
             Product_ID(Cn) = UF.Controls(i).Tag
           End If
         End If
      
      End If
 Next i
 N_Product = Cn

'===
  If SearchDefect = True Then
       If Cn = 0 Then
          MsgBox "You have not selected any product, please select one product which you like to search its defect descriptions." & CHR(10) & _
         UText(20320, 26410, 36873, 25321, 20219, 20309, 20135, 21697) & "，" & UText(35831, 36873, 25321, 19968, 20010, 20135, 21697, 26469, 26597, 35810, 20854, 32570, 38519, 25551, 36848) & "。" & CHR(10), vbCritical
         Unload Me
         Load_Product_Lev4
       ElseIf Cn > 1 Then
          MsgBox "You have not more than one product, please select only one product which you like to search its defect descriptions." & CHR(10) & _
         UText(20320, 36873, 25321, 20102, 22810, 20010, 20135, 21697) & "，" & UText(35831, 20165, 36873, 25321, 19968, 20010, 20135, 21697, 26469, 26597, 35810, 20854, 32570, 38519, 25551, 36848) & "。" & CHR(10), vbCritical
         Unload Me
         Load_Product_Lev4
       Else
         Defect_Product = Product(1)
         Defect_ProductE = ProductE(1)
         Defect_Product_ID = Product_ID(1)
         Unload Me
         Load_Product_DefectList
       End If
 Else
      If Cn = 0 Then
           MsgBox "You have not selected any product, will search check items according to the tertiary product categories you selected." & CHR(10) & _
           UText(20320, 26410, 36873, 25321, 20219, 20309, 20135, 21697, 21517, 31216) & "，" & UText(23558, 25353, 29031, 25152, 36873, 30340, 20135, 21697, 23567, 31867, 26597, 35810, 26816, 26597, 39033, 30446) & "。" & CHR(10), vbCritical
           Unload Me
           Load_Product_Lev4
       ElseIf Cn > 10 Then
            MsgBox "You have selected " & Cn & " products, no more than 10 products allowed to be selected，Please select again." & CHR(10) & _
           UText(20320, 36873, 25321, 20102) & Cn & UText(20010, 20135, 21697) & "，" & UText(27880, 24847, 27599, 27425, 21482, 20801, 35768, 36873, 25321, 26368, 22810) & "10" & UText(31181, 20135, 21697) & "，" & UText(35831, 37325, 26032, 36873, 25321) & "！" & CHR(10), vbCritical
            Unload Me
            Load_Product_Lev4
       Else
         'Me.Hide
          SLTN = Cn
          For i = 1 To N_Cat_Small
            ChID = Cat_Small_ID(i)
            For j = 1 To N_Product
              If Left(Product_ID(j), Len(ChID)) = ChID Then
                  GoTo 20
              End If
            Next j
            SLTN = SLTN + 1
            Product_ID(SLTN) = Cat_Small_ID(i) & "|"
20        Next i

          StoretoProperty "Selected Product Number", CStr(SLTN)
          For i = 1 To SLTN
             StoretoProperty "Product_ID" & " " & i, CStr(Product_ID(i))
             StoretoProperty "Product" & " " & i, CStr(Product(i))
             StoretoProperty "ProductE" & " " & i, CStr(ProductE(i))
         Next i

25      If Me.Choose_Character.Value = True Then
           'Product_Features.Show
            Unload Me
           Call Load_Product_Features
        Else
           ChT = ""
           Unload Me
          'Call Load_Product_ForDefects
           Call Get_CheckPoints
         End If
      End If
    End If
 
 
End Sub

Private Sub PREVIOUS_Click()
 Unload Me
 Load_Product_Lev3
End Sub
