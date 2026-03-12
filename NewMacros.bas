Attribute VB_Name = "NewMacros"
Public CustomerID As String
Public Cat_Large(1 To 3), Cat_Middle(1 To 10), Cat_Small(1 To 10), Product(1 To 10), Product_Feature(1 To 20) As String
Public Cat_LargeE(1 To 3), Cat_MiddleE(1 To 10), Cat_SmallE(1 To 10), ProductE(1 To 10), Product_FeatureE(1 To 20) As String
Public Cat_Large_Code(1 To 3), Cat_Middle_Code(1 To 10), Cat_Small_Code(1 To 10), Product_Code(1 To 10), Product_Feature_Code(1 To 10) As String
Public Cat_Large_ID(1 To 3), Cat_Middle_ID(1 To 10), Cat_Small_ID(1 To 10), Product_ID(1 To 10), Product_Feature_ID(1 To 20) As String
Public Defect_Product As String
Public Defect_ProductE As String
Public Defect_Product_ID As String
Public PriChk() As Variant
Public N_Cat_Large, N_Cat_Middle, N_Cat_Small, N_Product, N_Product_Feature As Integer
Public Srv_Large, Srv_Middle, Srv_Small As String
Public Factory_Name As String
Public tb, tbn, tbSno, SLTN As Integer
Public CHAR(100), ChT, ChT_ID(1 To 3) As String
Public ChN, Itms As Integer
Public Itx(), Ptx() As String
Public Preform As Object
Public tb_tt(12), BK(12) As String
Public WD As Document
Public Cel As Cell
Public tbl As table
Public Rc, cc As Integer
Public rpn As String
Public TokenJson As Object
Public SearchDefect As Boolean
Public EXT As String
'Public Const WebBaseUrl As String = "http://192.168.101.67"
Public Const WebBaseUrl As String = "https://www.hqts.cn"
Public Const AuthBaseUrl As String = "https://hub.hqts.cn"
Public Const AppId As String = "fuzhihao@hqts.cn"
Public Const EncryptedUserId = "V7NR8ZWF534oOVQFKgDQ+HIg1R4JYHM5KKV4E2ERhBaaQuCZfPsZDV7v30TKwMdwQgDr+efdHUP2P9y596AoKiRSt/iXnY8N8X10qEMDRYfwIg8Hd75C+RDj638MNY93vM6bNT+QMZrUDBDqj6ZvsO3Z61ieCMxeGIHz1tGADvLO5BtO9wUtCrkwbiB1M4L01HoCHk62tMU7WvPfvq+eoAsZUcb8qf4GIuVTvXHsVb0qP+wT+S1LetZ3fg1RHXsWFijtLevPYDqIi69yU3Mec9hlPjEMeIMvqMoimESVO4s7aqV328kZGVAh4apc0vCPmzOfG/5MTaogmpWgT5DOTw=="


Public Function UText(ParamArray codePoints() As Variant) As String
    Dim i As Long
    Dim result As String

    For i = LBound(codePoints) To UBound(codePoints)
        result = result & ChrW(CLng(codePoints(i)))
    Next i

    UText = result
End Function

Public Function HubKeyClientId() As String
    HubKeyClientId = UText(23458, 25143) & "id"
End Function

Public Function HubKeyClientFullName() As String
    HubKeyClientFullName = UText(23458, 25143, 20840, 31216)
End Function

Public Function HubKeySupplierId() As String
    HubKeySupplierId = UText(20379, 24212, 21830) & "id"
End Function

Public Function HubKeySupplierFullName() As String
    HubKeySupplierFullName = UText(20379, 24212, 21830, 20840, 31216)
End Function

Public Function HubKeyFactoryId() As String
    HubKeyFactoryId = UText(24037, 21378) & "id"
End Function

Public Function HubKeyFactoryFullName() As String
    HubKeyFactoryFullName = UText(24037, 21378, 20840, 31216)
End Function

Public Function HubKeyUnit() As String
    HubKeyUnit = UText(21333, 20301)
End Function


Public Function HubKeyProductTotalQuantity() As String
    HubKeyProductTotalQuantity = UText(20135, 21697, 24635, 25968, 37327)
End Function

Public Function HubKeyOrderNo() As String
    HubKeyOrderNo = UText(35746, 21333, 21495)
End Function

Public Function HubKeyItemNo() As String
    HubKeyItemNo = UText(27454, 21495)
End Function

Public Function HubKeyOrderProductName() As String
    HubKeyOrderProductName = UText(35746, 21333, 21697, 21517)
End Function

Public Function HubKeyServiceSmallCategory() As String
    HubKeyServiceSmallCategory = UText(26381, 21153, 23567, 31867)
End Function

Public Function HubKeyInspectionTimes() As String
    HubKeyInspectionTimes = UText(26816, 39564, 27425, 25968)
End Function

Public Function HubKeyActualStartDate() As String
    HubKeyActualStartDate = UText(23454, 38469, 23433, 25490, 24320, 22987, 26085, 26399)
End Function

Public Function HubKeyActualEndDate() As String
    HubKeyActualEndDate = UText(23454, 38469, 23433, 25490, 32467, 26463, 26085, 26399)
End Function

Public Function HubKeyServiceCity() As String
    HubKeyServiceCity = UText(26381, 21153, 22478, 24066)
End Function

Public Function HubKeyServiceCountryRegion() As String
    HubKeyServiceCountryRegion = UText(26381, 21153, 22269, 23478, 22320, 21306)
End Function

Public Function HubKeyGeneralWorkInstructionMIS() As String
    HubKeyGeneralWorkInstructionMIS = UText(36890, 29992, 20316, 19994, 25351, 23548, 20070) & "MIS"
End Function

Public Function HubKeyGeneralWorkInstruction() As String
    HubKeyGeneralWorkInstruction = UText(36890, 29992, 20316, 19994, 25351, 23548, 20070)
End Function

Public Function HubKeySpecialWorkInstruction() As String
    HubKeySpecialWorkInstruction = UText(19987, 29992, 20316, 19994, 25351, 23548, 20070)
End Function

Public Function HubKeyClientGeneralInspectionData() As String
    HubKeyClientGeneralInspectionData = UText(23458, 25143, 36890, 29992, 26816, 39564, 36164, 26009)
End Function

Public Function HubKeyTotalSampleSize() As String
    HubKeyTotalSampleSize = UText(24635, 25277, 26679, 25968)
End Function

Public Function HubKeyPrimaryProductLargeCategory() As String
    HubKeyPrimaryProductLargeCategory = UText(20027, 20135, 21697, 22823, 31867)
End Function

Public Function HubKeyPrimaryProductLargeCategoryId() As String
    HubKeyPrimaryProductLargeCategoryId = HubKeyPrimaryProductLargeCategory() & "id"
End Function

Public Function HubKeyPrimaryProductLargeCategoryCode() As String
    HubKeyPrimaryProductLargeCategoryCode = UText(20027, 20135, 21697, 22823, 31867, 32534, 30721)
End Function

Public Function HubKeyPrimaryProductMiddleCategory() As String
    HubKeyPrimaryProductMiddleCategory = UText(20027, 20135, 21697, 20013, 31867)
End Function

Public Function HubKeyPrimaryProductMiddleCategoryId() As String
    HubKeyPrimaryProductMiddleCategoryId = HubKeyPrimaryProductMiddleCategory() & "id"
End Function

Public Function HubKeyPrimaryProductMiddleCategoryCode() As String
    HubKeyPrimaryProductMiddleCategoryCode = UText(20027, 20135, 21697, 20013, 31867, 32534, 30721)
End Function

Public Function HubKeyPrimaryProductSmallCategory() As String
    HubKeyPrimaryProductSmallCategory = UText(20027, 20135, 21697, 23567, 31867)
End Function

Public Function HubKeyPrimaryProductSmallCategoryId() As String
    HubKeyPrimaryProductSmallCategoryId = HubKeyPrimaryProductSmallCategory() & "id"
End Function

Public Function HubKeyPrimaryProductSmallCategoryCode() As String
    HubKeyPrimaryProductSmallCategoryCode = UText(20027, 20135, 21697, 23567, 31867, 32534, 30721)
End Function

Public Function HubKeyProductName() As String
    HubKeyProductName = UText(20135, 21697, 21517, 31216)
End Function

Public Function HubKeyProductNameId() As String
    HubKeyProductNameId = HubKeyProductName() & "id"
End Function

Public Function HubKeyProductNameCode() As String
    HubKeyProductNameCode = UText(20135, 21697, 21517, 31216, 32534, 30721)
End Function

Public Function HubKeyServiceLargeCategory() As String
    HubKeyServiceLargeCategory = UText(26381, 21153, 22823, 31867)
End Function

Public Function HubKeyServiceMiddleCategory() As String
    HubKeyServiceMiddleCategory = UText(26381, 21153, 20013, 31867)
End Function

Public Function HubKeySamplingStandard() As String
    HubKeySamplingStandard = UText(25277, 26679, 26631, 20934)
End Function

Public Function HubKeySamplingPlan() As String
    HubKeySamplingPlan = UText(25277, 26679, 26041, 26696)
End Function

Public Function HubKeySamplingMethod() As String
    HubKeySamplingMethod = UText(25277, 26679, 26041, 24335)
End Function

Public Function HubValueSplitSampling() As String
    HubValueSplitSampling = UText(20998, 24320, 25277, 26679)
End Function

Public Function HubKeyInspectionLevel() As String
    HubKeyInspectionLevel = UText(26816, 39564, 27700, 24179)
End Function

Public Function HubKeyOrderQuantity() As String
    HubKeyOrderQuantity = UText(35746, 21333, 25968, 37327)
End Function

Public Function HubKeyProductMiddleCategoryNameZh() As String
    HubKeyProductMiddleCategoryNameZh = UText(20135, 21697, 20013, 31867, 20013, 25991)
End Function

Public Function HubKeyProductMiddleCategoryNameEn() As String
    HubKeyProductMiddleCategoryNameEn = UText(20135, 21697, 20013, 31867, 33521, 25991)
End Function

Public Function HubKeyProductMiddleCategoryId() As String
    HubKeyProductMiddleCategoryId = UText(20135, 21697, 20013, 31867) & "id"
End Function

Public Function HubKeyProductMiddleCategoryCode() As String
    HubKeyProductMiddleCategoryCode = UText(20135, 21697, 20013, 31867, 32534, 30721)
End Function


Public Function HubKeyProductLargeCategoryId() As String
    HubKeyProductLargeCategoryId = UText(20135, 21697, 22823, 31867) & "id"
End Function

Public Function HubKeyProductLargeCategoryCode() As String
    HubKeyProductLargeCategoryCode = UText(20135, 21697, 22823, 31867, 32534, 30721)
End Function

Public Function HubKeyProductLargeCategoryNameZh() As String
    HubKeyProductLargeCategoryNameZh = UText(20135, 21697, 22823, 31867, 20013, 25991)
End Function

Public Function HubKeyProductLargeCategoryNameEn() As String
    HubKeyProductLargeCategoryNameEn = UText(20135, 21697, 22823, 31867, 33521, 25991)
End Function

Public Function HubKeyProductSmallCategoryId() As String
    HubKeyProductSmallCategoryId = UText(20135, 21697, 23567, 31867) & "id"
End Function

Public Function HubKeyProductSmallCategoryCode() As String
    HubKeyProductSmallCategoryCode = UText(20135, 21697, 23567, 31867, 32534, 30721)
End Function

Public Function HubKeyProductSmallCategoryNameZh() As String
    HubKeyProductSmallCategoryNameZh = UText(20135, 21697, 23567, 31867, 20013, 25991)
End Function

Public Function HubKeyProductSmallCategoryNameEn() As String
    HubKeyProductSmallCategoryNameEn = UText(20135, 21697, 23567, 31867, 33521, 25991)
End Function

Public Function HubKeyProductNameZh() As String
    HubKeyProductNameZh = UText(20135, 21697, 21517, 31216, 20013, 25991)
End Function

Public Function HubKeyProductNameEn() As String
    HubKeyProductNameEn = UText(20135, 21697, 21517, 31216, 33521, 25991)
End Function

Public Function HubKeyProductFeatureId() As String
    HubKeyProductFeatureId = UText(20135, 21697, 29305, 24449) & "id"
End Function

Public Function HubKeyProductFeatureCode() As String
    HubKeyProductFeatureCode = UText(20135, 21697, 29305, 24449, 32534, 30721)
End Function

Public Function HubKeyProductFeatureNameZh() As String
    HubKeyProductFeatureNameZh = UText(20135, 21697, 29305, 24449, 20013, 25991)
End Function

Public Function HubKeyProductFeatureNameEn() As String
    HubKeyProductFeatureNameEn = UText(20135, 21697, 29305, 24449, 33521, 25991)
End Function

Public Function HubKeyCheckPointCode() As String
    HubKeyCheckPointCode = UText(27979, 35797, 39033, 30446, 32534, 30721)
End Function

Public Function HubKeyCheckPointName() As String
    HubKeyCheckPointName = UText(27979, 35797, 39033, 30446, 21517, 31216)
End Function

Public Function HubKeyCheckPointNameEn() As String
    HubKeyCheckPointNameEn = UText(27979, 35797, 39033, 30446, 21517, 31216) & "(" & UText(33521) & ")"
End Function

Public Function HubKeyAccordingStandardType() As String
    HubKeyAccordingStandardType = UText(25353, 24418, 24335, 24402, 31867)
End Function

Public Function HubKeyRequirement() As String
    HubKeyRequirement = UText(35201, 27714)
End Function

Public Function HubKeyRequirementEn() As String
    HubKeyRequirementEn = UText(35201, 27714) & "(" & UText(33521) & ")"
End Function

Public Function HubKeySampleQty() As String
    HubKeySampleQty = UText(25277, 26679, 37327)
End Function

Public Function HubKeySampleQtyEn() As String
    HubKeySampleQtyEn = UText(25277, 26679, 37327) & "(" & UText(33521) & ")"
End Function

Public Function HubKeyAcceptanceQty() As String
    HubKeyAcceptanceQty = UText(21442, 25968)
End Function

Public Function HubKeyAcceptanceQtyEn() As String
    HubKeyAcceptanceQtyEn = UText(21442, 25968) & "(" & UText(33521) & ")"
End Function

Public Function HubKeyRejectQty() As String
    HubKeyRejectQty = UText(20801, 24046, 20540)
End Function

Public Function HubKeyRejectQtyEn() As String
    HubKeyRejectQtyEn = UText(20801, 24046, 20540) & "(" & UText(33521) & ")"
End Function

Public Function HubKeyQuantityType() As String
    HubKeyQuantityType = UText(23450, 37327, 31867, 22411)
End Function

Public Function HubKeyProductFeature() As String
    HubKeyProductFeature = UText(20135, 21697, 29305, 24449)
End Function

Public Function HubKeyDefectName() As String
    HubKeyDefectName = UText(32570, 38519, 21517, 31216)
End Function

Public Function HubKeyDefectNameEn() As String
    HubKeyDefectNameEn = UText(32570, 38519, 21517, 31216) & "(" & UText(33521) & ")"
End Function

Public Function HubKeyDefectCode() As String
    HubKeyDefectCode = UText(32570, 38519, 32534, 30721)
End Function

Public Function HubKeyDefectLevel() As String
    HubKeyDefectLevel = UText(32570, 38519, 31867, 21035)
End Function

Sub Remark()
'Update to V1.1 item 20240125
'HUBquatityBreakdown，20240524,
End Sub
Sub InitializeParameter()
' translated comment
  tb_tt(0) = "INSPECTION INFORMATION":                 BK(0) = "INSPECTION INFORMATION"
  tb_tt(1) = "1. Quantity":                            BK(1) = "Quantity"
  tb_tt(2) = "2. Workmanship":                         BK(2) = "Workmanship"
  tb_tt(3) = "3. On Site Test":                        BK(3) = "OnSiteTest"
  tb_tt(4) = "4. Product Dimension & Weight":          BK(4) = "ProductDimensionWeight"
  tb_tt(5) = "5. Style / Material / Construction":         BK(5) = "StyleMaterialConstruction"
  tb_tt(6) = "6. Product Color":                       BK(6) = "ProductColor"
  tb_tt(7) = "7. Product Label / Marking":              BK(7) = "ProductLabelMarking"
  tb_tt(8) = "8. Shipping Mark":                       BK(8) = "ShippingMark"
  tb_tt(9) = "9. Packing / Packaging":                  BK(9) = "PackingPackaging"
  tb_tt(10) = "10. Other Note":                        BK(10) = "OtherNote"
  tb_tt(11) = "11. Other Photos":                      BK(11) = "OtherPhotos"
  Set WD = ThisDocument
  tb = WD.Tables.count - 1
End Sub

Function GetToken(username As String, password As String) As Object
    
    Dim req As Object 'WinHttp.WinHttpRequest
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim ReqStatus As Integer
    Dim ReqResponse As String
    Dim LoginApiURL As String
    Dim TokenApiURL As String
    Dim msg As String
    
    Dim RtlJson As New Scripting.Dictionary
    Dim postJson As Scripting.Dictionary
    Set postJson = New Scripting.Dictionary
    
    postJson("username") = username
    postJson("password") = password
    postJson("language") = "zh"
    
    RtlJson("success") = False
    
    
    'initiate login req
    LoginApiURL = AuthBaseUrl + "/api/ecology/login"
    With req
        .Open "POST", LoginApiURL, False
        .send JsonConverter.ConvertToJson(postJson)
        ReqStatus = .Status
        ReqResponse = .responseText
    End With
    If ReqStatus = "200" Then
        
        Dim LoginRespJson As Scripting.Dictionary
        Set LoginRespJson = JsonConverter.ParseJson(ReqResponse)
        
        If LoginRespJson("success") Then
            
            'initiate token req
            TokenApiURL = AuthBaseUrl + "/api/ecology/token"
            With req
                .Open "POST", TokenApiURL, False
                .setRequestHeader "userid", LoginRespJson("data")("userid")
                .setRequestHeader "identity", LoginRespJson("data")("identity")
                .send
                ReqStatus = .Status
                ReqResponse = .responseText
            End With
            
            If ReqStatus = "200" Then
                Dim TokenRespJson As New Scripting.Dictionary
                Set TokenRespJson = JsonConverter.ParseJson(ReqResponse)
                If TokenRespJson("success") Then
                    RtlJson("success") = True
                    Dim Data As New Scripting.Dictionary
                    Data("token") = TokenRespJson("data")("token")
                    Data("userid") = LoginRespJson("data")("userid")
                    
                    Set RtlJson("data") = Data
                    
                End If
                
            End If
            
            
        End If
    
    End If
    
    Set GetToken = RtlJson
End Function
                        

Function getTableNumber(tbSno)
Dim i As Integer
Dim Kw As String
Dim ttbno As Integer

Call InitializeParameter

Set WD = ActiveDocument
For i = 1 To tb
  With WD.Tables(i)
  
       If .Cell(1, 1).Range.Fields.count > 0 Then
            Kw = .Cell(1, 1).Range.Fields(1).Code
            Kw = Trim(Replace(Kw, "MACROBUTTON", ""))
            Kw = Trim(Mid(Kw, InStr(1, Kw, " ") + 1, Len(Kw) - InStr(1, Kw, " ")))
        
        Else
             Kw = Trim(Mid(.Cell(1, 1).Range.Text, 1, Len(.Cell(1, 1).Range.Text) - 2))
             'kw = .Cell(1, 1).Range.Text
             If Kw = "" Then GoTo 10
        End If

     
       'If tb_tt(tbSno) Like "*" & kw & "*" = True Then
       If Kw Like "*" & tb_tt(tbSno) & "*" = True And Val(Kw) = tbSno Then
          getTableNumber = i
          Exit Function
       End If
    
  End With
10 Next i
End Function
Function SendJSONRequest(ByVal Url As String, ByVal jsonInput As Object) As Object
    Dim httpRequest As Object
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim jsonOutput As Object
    Dim jsonSerialized As String
    Dim responseText As String
    
    ' Convert the input JSON object to a JSON string
    jsonSerialized = JsonConverter.ConvertToJson(jsonInput)
    
    
    ' Specify the method as POST, the URL, and set asynchronous to False
    httpRequest.Open "POST", Url, False
    
    ' Set the Content-Type header to 'application/json'
    httpRequest.setRequestHeader "Content-Type", "application/json"
    
    ' Send the JSON data
    httpRequest.send jsonSerialized
    
    ' Check the status code of the response
    If httpRequest.Status = 200 Then
        ' Get the response text
        responseText = httpRequest.responseText
        
        ' Parse the response text back into a JSON object
        Set jsonOutput = JsonConverter.ParseJson(responseText)
    Else
        ' Handle the error accordingly
        ' For example, setting the jsonOutput to Nothing or to an error description
        Set jsonOutput = Nothing
        ' Optionally, raise an error or log it
        ' Err.Raise vbObjectError + 1, "SendJSONRequest", "Error retrieving data: " & httpRequest.Status & " - " & httpRequest.statusText
    End If
    
    ' Return the JSON object
    Set SendJSONRequest = jsonOutput
    
    ' Clean up
    Set httpRequest = Nothing
End Function

Sub CNAS()
'by wyn
  On Error Resume Next
  Application.ScreenUpdating = False
  Dim myrange As Range
  Dim SearchStr, SearchChar, Position
  Dim bkRange As Range
  Dim hk As String
  Dim yourarray, renming As String
  
   yourarray = Array("cjh", "pys", "ft", "hm", "hzh", "ljl", "lsw", "lzc", "lzw", "wjf", "wyn", "xh", "zjs", "zyr", "zh", "cln", "csj")
                    '，，，，，，，，，，，，，，，,
                    
    renming = InputBox(UText(35831, 36755, 20837, 25216, 26415, 32463, 29702, 30340, 25340, 38899, 32553, 20889) & "，" & UText(22914) & "： " & UText(40644, 26126) & " " & UText(21017, 36755, 20837) & " MH " & UText(25110) & " hm," & UText(19981, 21306, 20998, 22823, 23567, 20889) & " ", , "hm")
    renming = LCase(renming)
    wenjian = "d:\hqtscnas\" & renming & ".png"
    If renming = "" Then Exit Sub
    For Each Item In yourarray
    If renming = Item Then KKK = "cunzai"
    Next
    If KKK = "" Then Exit Sub
  
  
'HQTS

   fpage = ThisDocument.PageSetup.DifferentFirstPageHeaderFooter

   If fpage = 0 Then
       With Selection.PageSetup
       .DifferentFirstPageHeaderFooter = True
        End With
    End If


    Set myrange = ThisDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range
    Set bkRange = myrange.Paragraphs(3).Range
    judge = Left(bkRange, 8)

    'If judge <> "HELMSMAN" Then Exit Sub
    Set myrange = ThisDocument.Sections(1).Footers(wdHeaderFooterFirstPage).Range
    Set bkRange = myrange.Paragraphs(5).Range
    If InStr(bkRange, UText(26412, 33609, 31295, 25253, 21578, 22522, 20110, 26102, 38388, 25152, 38480)) < 0 Then MsgBox UText(25253, 21578, 26684, 24335, 26377, 38382, 39064, 19981, 33021, 36827, 34892, 31614, 21517): Exit Sub
   
    
' translated comment
 
    If ThisDocument.PageSetup.DifferentFirstPageHeaderFooter = 0 Then
       With Selection.PageSetup
       .DifferentFirstPageHeaderFooter = True
        End With
    End If


 'CNAS

    If ThisDocument.Bookmarks.Exists("CNAS") = False Then
    Set myrange = ThisDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range
    Set bkRange = myrange.Paragraphs(7).Range
    bkRange.Bookmarks.add Name:="CNAS", Range:=bkRange
    End If
  
   'hqtssuper
    If ThisDocument.Bookmarks.Exists("hqtssuper") = False Then
    Set myrange = ThisDocument.Sections(1).Footers(wdHeaderFooterFirstPage).Range
    Set bkRange = myrange.Paragraphs(1).Range
    bkRange.Bookmarks.add Name:="hqtssuper", Range:=bkRange
    End If
  
   'Signnamer
    If ThisDocument.Bookmarks.Exists("signname") = False Then
    Set myrange = ThisDocument.Sections(1).Footers(wdHeaderFooterFirstPage).Range
    Set bkRange = myrange.Paragraphs(2).Range
    bkRange.Bookmarks.add Name:="signname", Range:=bkRange
    End If

    'xuanchuan
    If ThisDocument.Bookmarks.Exists("xuanchuan") = False Then
    Set myrange = ThisDocument.Sections(1).Footers(wdHeaderFooterFirstPage).Range
    Set bkRange = myrange.Paragraphs(5).Range
    bkRange.Bookmarks.add Name:="xuanchuan", Range:=bkRange
    End If

'======================================================================
  If ThisDocument.Bookmarks.Exists("signname") = True Then
       If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
       End If
       If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
       End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    
    ' translated comment
    ThisDocument.Bookmarks("signname").Select
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 12
    Selection.Font.Bold = 0
    Selection.Delete
    With Selection.ParagraphFormat
       .LineSpacingRule = wdLineSpaceMultiple
       .LineSpacing = LinesToPoints(1)
       .Alignment = wdAlignParagraphJustify
    End With
    
    
    
    Selection.InlineShapes.AddPicture FileName:=wenjian _
        , LinkToFile:=False, SaveWithDocument:=True
    Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
    Selection.Font.UnderlineColor = wdColorAutomatic
    'Selection.Font.Underline = wdUnderlineSingle

    With Selection.ParagraphFormat
       .LineSpacingRule = wdLineSpaceMultiple
       .LineSpacing = LinesToPoints(1)
       .Alignment = wdAlignParagraphJustify
        
    End With
    'HQTS supervisor
    If ThisDocument.Bookmarks.Exists("hqtssuper") = True Then
        ThisDocument.Bookmarks("hqtssuper").Select
        Selection.Font.Name = "Arial"
        Selection.Font.Size = 12
        Selection.Font.Bold = 1

        Selection.TypeText Text:="Approved by HQTS Supervisor:"
    End If
  
    '17020
    If ThisDocument.Bookmarks.Exists("xuanchuan") = True Then
        ThisDocument.Bookmarks("xuanchuan").Select
        Selection.SelectCell

      With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1)
        .Alignment = wdAlignParagraphJustify
      End With
    Selection.Font.Name = "Arial"
    Selection.Font.Size = 10
    Selection.TypeText Text:="This report reflects the facts as recorded by HQTS at the time and place of inspection. It does not relieve the manufacturers from their contractual obligations nor prejudice client's right for compensation for any apparent and/or hidden defects not detected during our random inspection or occuring thereafter."
    End If

   'CNAS
    If ThisDocument.Bookmarks.Exists("CNAS") = True Then
        ThisDocument.Bookmarks("CNAS").Select
        Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
       If Selection.InlineShapes(1).PictureFormat.Brightness <> 1# Then Selection.InlineShapes(1).PictureFormat.Brightness = 1#
       'If Selection.InlineShapes(1).PictureFormat.Contrast <> 1# Then Selection.InlineShapes(1).PictureFormat.Contrast = 1#: GoTo 10
       If Selection.InlineShapes(1).PictureFormat.Brightness <> 0.5 Then Selection.InlineShapes(1).PictureFormat.Brightness = 1#
       If Selection.InlineShapes(1).PictureFormat.Contrast <> 0.5 Then Selection.InlineShapes(1).PictureFormat.Contrast = 1#
    If ThisDocument.Bookmarks.Exists("xuanchuan") = True Then
        ThisDocument.Bookmarks("xuanchuan").Select
        Selection.Borders(wdBorderTop).Color = wdColorAutomatic
        Selection.Borders(wdBorderLeft).Color = wdColorAutomatic
        Selection.Borders(wdBorderBottom).Color = wdColorAutomatic
        Selection.Borders(wdBorderRight).Color = wdColorAutomatic
        Selection.Font.Color = wdColorAutomatic
       GoTo 20
     End If
       End If
End If


20:     ActiveWindow.ActivePane.Close
    ActiveWindow.ActivePane.View.Type = wdPrintView
ThisDocument.Range(Start:=0, End:=0).Select
    
    Application.ScreenUpdating = True

End Sub

Sub FillInformation()

'revised 241216
 Dim jsonResponse As Object
    Dim jsonPara As Object
    Dim Url As String
  
  
    Dim ReportNo As String
    Dim Pcn As Integer
    Dim StartDate, EndDate As String
    Dim myrange As Range
    Dim i As Integer
    Dim Kw As String
    Dim Descel As Cell
    Dim Cn As Integer
    Dim Client As String, Supplier As String, Factory As String
    Dim PO, Item, Quantity, Product_Description, Inspection_Type, Sequence, _
    Inspection_Date, Location, Inspection_Basis, Sample_Size As String
    Dim ClientID As String
    Dim SupplierID As String
    Dim FactoryID As String
    Dim ClientName As String
    Dim SupplierName As String
    Dim FactoryName As String
    Dim ReportID As String
    Dim Q As String
    
    tb = ThisDocument.Tables.count
    
    
3    Client = "Client"
     Supplier = "Supplier"
     Factory = "Factory"
     PO = "P.O. No."
     Item = "Item No."
     Quantity = "Quantity"
     Product_Description = "Product Description"
     Inspection_Type = "Inspection Type"
     Sequence = "Sequence"
     Inspection_Date = "Inspection Date"
     Location = "Location"
     Inspection_Basis = "Inspection Basis"
     Sample_Size = "Sample Size"
     
     '----------
     Set myrange = ThisDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Range
     Pcn = myrange.Cells.count
     rpn = myrange.Cells(Pcn).Range.Text
     rpn = Mid(rpn, 1, Len(rpn) - 2)
     rpn = Trim(Replace(rpn, "Report No: ", ""))
     If Len(rpn) < 13 Then
        rpn = InputBox("Please input the report number on HUB such as H2401045257HF-A1 & chr(10) & " & UText(35831, 36755, 20837) & "HUB" & UText(19978, 30340, 25253, 21578, 21495, 35937) & "H2401045257HF-A1", _
        "Input report Number")
         If Len(rpn) < 13 Then
           MsgBox "You have not input correct report number,the program exit" & CHR(10) & _
           UText(20320, 26410, 36755, 20837, 27491, 30830, 30340, 25253, 21578, 21495) & "，" & UText(31243, 24207, 36864, 20986) & "！", vbCritical
            Exit Sub
         End If
           myrange.Cells(Pcn).Range.Text = "Report No: " & rpn
     End If
       Call VerifyVersion(rpn, Q)
       If Q = "Quit" Then Exit Sub
      
      If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
       End If
       If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
       End If
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        
      Dim HeaderRange As Range
      Set HeaderRange = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range
      HeaderRange.Paragraphs(1).Range.Text = "Report No: " & rpn & CHR(13)
      ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
     
     '===========================

    Dim postJson As Object
     
         Set postJson = New Scripting.Dictionary
         postJson("bgh") = rpn
         
         Url = WebBaseUrl & "/api/sqlhub/exec?appid=md_330&configName=GetReportBasicInfo"
         Call getHUBdata(Url, postJson, jsonResponse)
         
   'Dim json_str As String
   'json_str = JsonConverter.ConvertToJson(jsonResponse)
   ' Debug.Print json_str
         
             '，dataIsEmpty
          Dim dataIsEmpty As Boolean
         If Not IsEmpty(jsonResponse("data")) Then
             dataIsEmpty = (TypeName(jsonResponse("data")) = "Collection") And (jsonResponse("data").count = 0)
          Else
             dataIsEmpty = True
          End If
         If dataIsEmpty Then
              MsgBox "Ｎo data is obtained from the HUB, the program will exit, please check 1. whether the HUB can be accessed via the Internet," & _
              "2. Whether the order is in the correct process stage on the HUB, such as the inspection notice has been issued and the report number has been generated. " & _
              "Whether the report number you input is correct!" & CHR(10) & CHR(10) & _
              UText(20174) & "HUB" & UText(26410, 33719, 21462, 21040, 25968, 25454) & "，" & UText(31243, 24207, 23558, 36864, 20986) & "，" & UText(35831, 26816, 26597) & "1." & UText(36890, 36807, 20114, 32852, 32593, 26159, 21542, 21487, 20197, 35775, 38382) & "HUB,2." & UText(36825, 20010, 35746, 21333, 22312) & "HUB" & UText(19978, 26159, 21542, 22788, 20110, 27491, 30830, 30340, 27969, 31243, 38454, 27573) & "，" & UText(27604, 22914, 39564, 36135, 36890, 30693, 21333, 24050, 21457, 20986) & "，" & UText(25253, 21578, 21495, 24050, 29983, 25104) & "," & _
              "3." & UText(20320, 36755, 20837, 30340, 25253, 21578, 21495, 26159, 21542, 27491, 30830), vbCritical
              Exit Sub
          End If
        
        '-----
        Set WD = ThisDocument
        WD.Activate
           ReportID = jsonResponse("data")(1)("bghid")
           StoretoProperty "ReportID", "ReportID" & "_" & ReportID
           ClientID = jsonResponse("data")(1)(HubKeyClientId())
           StoretoProperty "ClientID", ClientID
           ClientName = jsonResponse("data")(1)(HubKeyClientFullName())
           StoretoProperty Client, ClientName
           SupplierID = jsonResponse("data")(1)(HubKeySupplierId())
           StoretoProperty "SupplierID", SupplierID
           SupplierName = jsonResponse("data")(1)(HubKeySupplierFullName())
           StoretoProperty Supplier, SupplierName
           FactoryID = jsonResponse("data")(1)(HubKeyFactoryId())
           StoretoProperty "FactoryID", FactoryID
            FactoryID = Replace(FactoryID, ",", "_")
           FactoryName = jsonResponse("data")(1)(HubKeyFactoryFullName())
           StoretoProperty Factory, FactoryName
          
2        Set tbl = WD.Tables(1)
         With tbl
            Select Case ClientID
            Case "27153", "52368"   ''Starlight Textil-Handels GmbH,SL-Hunter,，sample size
            If .Cell(4, 1).Range.Text Like "*Factory*" = False Then
             .Rows(3).Select
              Selection.InsertRowsBelow 1
             .Cell(4, 1).Range.Text = "Factory"
            End If
            .Rows(5).Select
            Selection.InsertRowsBelow 1
            .Cell(6, 1).Range.Text = "Item No."
            .Cell(6, 3).Range.Text = "Sample Size"
            .Rows(7).Delete
          Case "28245", "28249", "28361"     'Tommy Teleshopping,JML Benelux,Valia Trading Corp
            If .Cell(4, 1).Range.Text Like "*Factory*" = False Then
             .Rows(3).Select
              Selection.InsertRowsBelow 1
             .Cell(4, 1).Range.Text = "Factory"
            End If
          Case "49702"     'NEW THINKING FASHION (CHINA) LTD
             WD.Range(.Cell(9, 2).Range.Start, .Cell(9, 4).Range.End).Cells.Merge
              .Cell(9, 2).Range.Text = ""
              .Rows(4).Delete
          Case "33439"     'IBA Hartmann GmbH & Co. KG， '
              Rn = .Rows.count
              For i = Rn To 1 Step -1
                If .Cell(i, 1).Range.Text Like "*Factory*" = True Then
                   .Rows(i).Delete
                ElseIf .Cell(i, 1).Range.Text Like "*Supplier*" = True Then
                   .Rows(i).Delete
                End If
              Next i
              
          Case Else
             
          End Select
        End With
  
      '------

      With tbl
         Cn = .Range.Cells.count
         For i = 2 To Cn - 1 Step 2
           Set Cel = .Range.Cells(i)
           Set Descel = .Range.Cells(i + 1)
              Kw = Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2)
           If Kw = Client Then
              Descel.Range.Text = jsonResponse("data")(1)(HubKeyClientFullName())
              Call DeleteBookMark(Descel)
              Descel.Range.Bookmarks.add Name:="ClientID" & "_" & ClientID, Range:=Descel.Range
              GoTo 20
           ElseIf Kw = Supplier Then
               Descel.Range.Text = jsonResponse("data")(1)(HubKeySupplierFullName())
                Call DeleteBookMark(Descel)
               Descel.Range.Bookmarks.add Name:="SupplierID" & "_" & SupplierID, Range:=Descel.Range
              GoTo 20
           ElseIf Kw = Factory Then
              Descel.Range.Text = jsonResponse("data")(1)(HubKeyFactoryFullName())
               Call DeleteBookMark(Descel)
              Descel.Range.Bookmarks.add Name:="FactoryID" & "_" & FactoryID, Range:=Descel.Range
              GoTo 20
           ElseIf Kw = Quantity Then
             Descel.Range.Text = jsonResponse("data")(1)(HubKeyProductTotalQuantity()) & " " & jsonResponse("data")(1)(HubKeyUnit())
              If jsonResponse("data")(1)(HubKeyUnit()) = "" Then Descel.Range.Text = jsonResponse("data")(1)(HubKeyProductTotalQuantity()) & " " & "Pieces"
              GoTo 20
           ElseIf Kw = PO Then
              Descel.Range.Text = Replace(jsonResponse("data")(1)(HubKeyOrderNo()), ",", ", ")
             
              GoTo 20
           ElseIf Kw = Item Then
              Descel.Range.Text = Replace(jsonResponse("data")(1)(HubKeyItemNo()), ",", ", ")
              GoTo 20
           ElseIf Kw = Product_Description Then
              Descel.Range.Text = jsonResponse("data")(1)(HubKeyOrderProductName())
              GoTo 20
           ElseIf Kw = Inspection_Type Then
              Descel.Range.Text = jsonResponse("data")(1)(HubKeyServiceSmallCategory())
              GoTo 20
          ElseIf Kw = Sequence Then
              Descel.Range.Text = jsonResponse("data")(1)(HubKeyInspectionTimes())
              If jsonResponse("data")(1)(HubKeyInspectionTimes()) = "" Then Descel.Range.Text = "Original Inspection"
              GoTo 20
           ElseIf Kw = Inspection_Date Then
               StartDate = jsonResponse("data")(1)(HubKeyActualStartDate())
               EndDate = jsonResponse("data")(1)(HubKeyActualEndDate())
               If StartDate = EndDate Then
                 Descel.Range.Text = StartDate
               Else
                   Descel.Range.Text = StartDate & "-" & EndDate
               End If
               GoTo 20
             ElseIf Kw = Location Then
              Descel.Range.Text = jsonResponse("data")(1)(HubKeyServiceCity()) & ", " & jsonResponse("data")(1)(HubKeyServiceCountryRegion())
              GoTo 20
            ElseIf Kw = Inspection_Basis Then
                Kw = jsonResponse("data")(1)(HubKeyGeneralWorkInstructionMIS())
                If Kw = "" Then
                     Kw = jsonResponse("data")(1)(HubKeyGeneralWorkInstruction())
                End If
                If jsonResponse("data")(1)(HubKeySpecialWorkInstruction()) <> "" Then
                  Kw = Kw & ", " & jsonResponse("data")(1)(HubKeySpecialWorkInstruction())
               End If
                If jsonResponse("data")(1)(HubKeyClientGeneralInspectionData()) <> "" Then
                  Kw = Kw & ", " & jsonResponse("data")(1)(HubKeyClientGeneralInspectionData())
               End If
              Descel.Range.Text = Kw
              GoTo 20
            ElseIf Kw = Sample_Size Then
             Descel.Range.Text = jsonResponse("data")(1)(HubKeyTotalSampleSize()) & " " & jsonResponse("data")(1)(HubKeyUnit())
             If jsonResponse("data")(1)(HubKeyUnit()) = "" Then Descel.Range.Text = jsonResponse("data")(1)(HubKeyTotalSampleSize()) & " " & "Pieces"
              GoTo 20
           End If
20     Next i
        
        
         ReportNo = rpn
         
      End With
    '------------------
     Cat_Large(1) = jsonResponse("data")(1)(HubKeyPrimaryProductLargeCategory())
     Cat_Large_ID(1) = jsonResponse("data")(1)(HubKeyPrimaryProductLargeCategoryId())
     Cat_Large_Code(1) = jsonResponse("data")(1)(HubKeyPrimaryProductLargeCategoryCode())
     Cat_Middle(1) = jsonResponse("data")(1)(HubKeyPrimaryProductMiddleCategory())
     Cat_Middle_ID(1) = jsonResponse("data")(1)(HubKeyPrimaryProductMiddleCategoryId())
     Cat_Middle_Code(1) = jsonResponse("data")(1)(HubKeyPrimaryProductMiddleCategoryCode())
     Cat_Small(1) = jsonResponse("data")(1)(HubKeyPrimaryProductSmallCategory())
     Cat_Small_ID(1) = jsonResponse("data")(1)(HubKeyPrimaryProductSmallCategoryId())
     Cat_Small_Code(1) = jsonResponse("data")(1)(HubKeyPrimaryProductSmallCategoryCode())
     Product(1) = jsonResponse("data")(1)(HubKeyProductName())
     Product_ID(1) = jsonResponse("data")(1)(HubKeyProductNameId())
     Product_Code(1) = jsonResponse("data")(1)(HubKeyProductNameCode())
     
     Srv_Large = jsonResponse("data")(1)(HubKeyServiceLargeCategory())
     Srv_Middle = jsonResponse("data")(1)(HubKeyServiceMiddleCategory())
     Srv_Small = jsonResponse("data")(1)(HubKeyServiceSmallCategory())
     Factory_Name = jsonResponse("data")(1)(HubKeyFactoryFullName())
    '----------------------
   
    Dim Occ As contentControl
    Dim dropdownEntry As ContentControlListEntry
    Dim Option_Selected As String
    Dim CCN  As Integer
    
    tbn = getTableNumber(2)

    With ThisDocument.Tables(tbn)
                             ' translated comment
    Dim SamplingStd_Name(), SamplingStd_Value() As String
    CCN = 9
    ReDim SamplingStd_Name(1 To CCN), SamplingStd_Value(1 To CCN)
   
  
  '--------
    SamplingStd_Name(1) = "ISO2859.1":        SamplingStd_Value(1) = "ISO2859.1"
    SamplingStd_Name(2) = "N/A":              SamplingStd_Value(2) = "N/A"
    SamplingStd_Name(3) = "Other":            SamplingStd_Value(3) = "Other"
    SamplingStd_Name(4) = "ANSI/ASQCZ1.4":    SamplingStd_Value(4) = "ANSI/ASQCZ1.4"
    SamplingStd_Name(5) = "BS6001.1":         SamplingStd_Value(5) = "BS6001.1"
    SamplingStd_Name(6) = "DIN40080":         SamplingStd_Value(6) = "DIN40080"
    SamplingStd_Name(7) = "NFX06-022":        SamplingStd_Value(7) = "NFX06-022"
    SamplingStd_Name(8) = "STD-019-HQT":      SamplingStd_Value(8) = "STD-019-HQT"
    SamplingStd_Name(9) = "STD-S-BLL-020":    SamplingStd_Value(9) = "STD-S-BLL-020"
    
    
    Option_Selected = jsonResponse("data")(1)(HubKeySamplingStandard())
    '-----------------------
        Set Cel = .Cell(3, 2)
        Cel.Range.Delete
        Set Occ = Cel.Range.ContentControls.add(wdContentControlComboBox)
        
        With Occ
            .Title = "Sampling Standart"
            .Tag = "Sampling Standart"
          For i = LBound(SamplingStd_Name) To UBound(SamplingStd_Name)
             .DropdownListEntries.add SamplingStd_Name(i), SamplingStd_Value(i)
          Next i
          
         
          For Each dropdownEntry In .DropdownListEntries
              If dropdownEntry.Text = Option_Selected Then
                dropdownEntry.Select
                Exit For
             End If
          Next dropdownEntry
         
                                              'Debug.Print "Sampling Standart=" & oCC.Range.Text
        End With
  
    '-----------------------
   
    Dim SamplingPlan_Name(), SamplingPlan_Value() As String
    CCN = 4
    ReDim SamplingPlan_Name(1 To CCN), SamplingPlan_Value(1 To CCN)
    
   
   
    '--------
    SamplingPlan_Name(1) = "Single, Normal":                                 SamplingPlan_Value(1) = "Single sampling plans for normal inspection"
    SamplingPlan_Name(2) = "Fixed Sample Size":                              SamplingPlan_Value(2) = "Fixed Sample"
    SamplingPlan_Name(3) = "N/A":                                            SamplingPlan_Value(3) = "N/A"
    SamplingPlan_Name(4) = "Percentage Sampling":                            SamplingPlan_Value(4) = "Percentage Sampling"
   ' SamplingPlan_Name(5) = "Single, Tightened":                              SamplingPlan_Value(5) = "Single sampling plans for tightened inspection"
    'SamplingPlan_Name(6) = "Single, Reduced":                                SamplingPlan_Value(6) = "Single sampling plans for reduced inspection"
        
        Option_Selected = jsonResponse("data")(1)(HubKeySamplingPlan())
        Set Cel = .Cell(4, 2)
        Cel.Range.Delete
        Set Occ = Cel.Range.ContentControls.add(wdContentControlComboBox)
        With Occ
            .Title = "Sampling Plan"
            .Tag = "Sampling Plan"
          For i = LBound(SamplingPlan_Name) To 3 'UBound(SamplingPlan_Name)
             .DropdownListEntries.add SamplingPlan_Name(i), SamplingPlan_Value(i)
          Next i
          For Each dropdownEntry In .DropdownListEntries
              If dropdownEntry.Value = Option_Selected Then
                dropdownEntry.Select
                Exit For
             End If
          Next dropdownEntry
                                                    ' Debug.Print "Sampling Plan=" & oCC.Range.Text
        End With
    
    '-------------
  
   Dim InspectionLevel_Name(), InspectionLevel_Value() As String
    CCN = 12
    ReDim InspectionLevel_Name(1 To CCN), InspectionLevel_Value(1 To CCN)
    If jsonResponse("data")(1)(HubKeySamplingPlan()) = "Fixed Sample" Then
      Option_Selected = "Fixed Sample Size"
     End If
    If jsonResponse("data")(1)(HubKeySamplingMethod()) = HubValueSplitSampling() Then
     Option_Selected = Option_Selected & " " & "per Item"
    End If
    
    '--------
    InspectionLevel_Name(1) = "G-II":               InspectionLevel_Value(1) = "G-II"
    InspectionLevel_Name(2) = "G-II per Item":      InspectionLevel_Value(2) = "G-II per Item"
    InspectionLevel_Name(3) = "G-I":                InspectionLevel_Value(3) = "G-I"
    InspectionLevel_Name(4) = "G-I per Item":       InspectionLevel_Value(4) = "G-I per Item"
    InspectionLevel_Name(5) = "Fixed Sample Size":  InspectionLevel_Value(5) = "Fixed Sample Size"
    InspectionLevel_Name(6) = "G-III":              InspectionLevel_Value(6) = "G-III"
    InspectionLevel_Name(7) = "G-III per Item":     InspectionLevel_Value(7) = "G-III per Item"
    InspectionLevel_Name(8) = "S4":                 InspectionLevel_Value(8) = "S4"
    InspectionLevel_Name(9) = "S3":                 InspectionLevel_Value(9) = "S3"
    InspectionLevel_Name(10) = "S2":                InspectionLevel_Value(10) = "S2"
    InspectionLevel_Name(11) = "S1":                InspectionLevel_Value(11) = "S1"
    InspectionLevel_Name(12) = "Fixed Sample Size per item":  InspectionLevel_Value(12) = "Fixed Sample Size per item"
        
        Option_Selected = jsonResponse("data")(1)(HubKeyInspectionLevel())
        Set Cel = .Cell(5, 2)
        Cel.Range.Delete
        Set Occ = Cel.Range.ContentControls.add(wdContentControlComboBox)
        With Occ
            .Title = "Inspection Level"
            .Tag = "Inspection Level"
          For i = LBound(InspectionLevel_Name) To UBound(InspectionLevel_Name)
             .DropdownListEntries.add InspectionLevel_Name(i), InspectionLevel_Value(i)
          Next i
          For Each dropdownEntry In .DropdownListEntries
              If dropdownEntry.Text = Option_Selected Then
                dropdownEntry.Select
                Exit For
             End If
          Next dropdownEntry
                                                   ' Debug.Print "Inspection Level=" & oCC.Range.Text
        End With
 
   '============================================================================
   
    Dim CriticalAql_Name(), CriticalAql_Value() As String
    CCN = 12
    ReDim CriticalAql_Name(1 To CCN), CriticalAql_Value(1 To CCN)
    
       '--------
    CriticalAql_Name(1) = "0":                  CriticalAql_Value(1) = "0"
    CriticalAql_Name(2) = "0.01":               CriticalAql_Value(2) = "0.01"
    CriticalAql_Name(3) = "0.015":              CriticalAql_Value(3) = "0.015"
    CriticalAql_Name(4) = "0.025":              CriticalAql_Value(4) = "0.025"
    CriticalAql_Name(5) = "0.04":               CriticalAql_Value(5) = "0.04"
    CriticalAql_Name(6) = "0.065":              CriticalAql_Value(6) = "0.065"
    CriticalAql_Name(7) = "0.1":                CriticalAql_Value(7) = "0.1"
    CriticalAql_Name(8) = "0.15":               CriticalAql_Value(8) = "0.15"
    CriticalAql_Name(9) = "0.25":               CriticalAql_Value(9) = "0.25"
    CriticalAql_Name(10) = "0.65":              CriticalAql_Value(10) = "0.65"
    CriticalAql_Name(11) = "1":                 CriticalAql_Value(11) = "1"
    CriticalAql_Name(12) = "1.5":               CriticalAql_Value(12) = "1.5"
        
        Option_Selected = jsonResponse("data")(1)("criticalaql")


        Set Cel = .Cell(4, 4)
        Cel.Range.Delete
        Set Occ = Cel.Range.ContentControls.add(wdContentControlComboBox)
        With Occ
            .Title = "CriticalAql"
            .Tag = "CriticalAql"
          For i = LBound(CriticalAql_Name) To UBound(CriticalAql_Name)
             .DropdownListEntries.add CriticalAql_Name(i), CriticalAql_Value(i)
          Next i
           For Each dropdownEntry In .DropdownListEntries
              If dropdownEntry.Text = Option_Selected Then
                dropdownEntry.Select
                Exit For
             End If
          Next dropdownEntry
                                                   ' Debug.Print "CriticalAql" & oCC.Range.Text
        End With
 
   '==================================================
   
    Dim MajorAql_Name(), MajorAql_Value() As String
    CCN = 13
    ReDim MajorAql_Name(1 To CCN), MajorAql_Value(1 To CCN)
    
       '--------
    MajorAql_Name(1) = "2.5":                MajorAql_Value(1) = "2.5"
    MajorAql_Name(2) = "1.5":                MajorAql_Value(2) = "1.5"
    MajorAql_Name(3) = "1":                  MajorAql_Value(3) = "1"
    MajorAql_Name(4) = "4.0":                MajorAql_Value(4) = "4.0"
    MajorAql_Name(5) = "0.65":               MajorAql_Value(5) = "0.65"
    MajorAql_Name(6) = "0.4":                MajorAql_Value(6) = "0.4"
    MajorAql_Name(7) = "0.25":               MajorAql_Value(7) = "0.25"
    MajorAql_Name(8) = "0.15":               MajorAql_Value(8) = "0.15"
    MajorAql_Name(9) = "0.1":                MajorAql_Value(9) = "0.1"
    MajorAql_Name(10) = "0.065":              MajorAql_Value(10) = "0.065"
    MajorAql_Name(11) = "0.04":             MajorAql_Value(11) = "0.04"
    MajorAql_Name(12) = "6.5":               MajorAql_Value(12) = "6.5"
    MajorAql_Name(13) = "NA":                MajorAql_Value(13) = "N/A"
    
        Option_Selected = jsonResponse("data")(1)("majoraql")
        Set Cel = .Cell(4, 5)
        Cel.Range.Delete
        Set Occ = Cel.Range.ContentControls.add(wdContentControlComboBox)
        With Occ
            .Title = "MajorAql"
            .Tag = "MajorAql"
          For i = LBound(MajorAql_Name) To UBound(MajorAql_Name)
             .DropdownListEntries.add MajorAql_Name(i), MajorAql_Value(i)
          Next i
           For Each dropdownEntry In .DropdownListEntries
              If dropdownEntry.Text = Option_Selected Then
                dropdownEntry.Select
                Exit For
             End If
          Next dropdownEntry
                                                   ' Debug.Print "MajorAql" & oCC.Range.Text
        End With
  '================
    
    Dim MinorAql_Name(), MinorAql_Value() As String
    CCN = 13
    ReDim MinorAql_Name(1 To CCN), MinorAql_Value(1 To CCN)
    
       '--------
    MinorAql_Name(1) = "4.0":                MinorAql_Value(1) = "4.0"
    MinorAql_Name(2) = "2.5":                MinorAql_Value(2) = "2.5"
    MinorAql_Name(3) = "1.5":                MinorAql_Value(3) = "1.5"
    MinorAql_Name(4) = "1":                  MinorAql_Value(4) = "1"
    MinorAql_Name(5) = "0.65":               MinorAql_Value(5) = "0.65"
    MinorAql_Name(6) = "0.4":                MinorAql_Value(6) = "0.4"
    MinorAql_Name(7) = "0.25":               MinorAql_Value(7) = "0.25"
    MinorAql_Name(8) = "0.15":               MinorAql_Value(8) = "0.15"
    MinorAql_Name(9) = "0.1":                MinorAql_Value(9) = "0.1"
    MinorAql_Name(10) = "0.065":              MinorAql_Value(10) = "0.065"
    MinorAql_Name(11) = "6.5":               MinorAql_Value(11) = "6.5"
    MinorAql_Name(12) = "10":                MinorAql_Value(12) = "10"
    MinorAql_Name(13) = "NA":                MinorAql_Value(13) = "N/A"
    
        Option_Selected = jsonResponse("data")(1)("minoraql")
        Set Cel = .Cell(4, 6)
        Cel.Range.Delete
        Set Occ = Cel.Range.ContentControls.add(wdContentControlComboBox)
        With Occ
            .Title = "MinorAql"
            .Tag = "MinorAql"
          For i = LBound(MinorAql_Name) To UBound(MinorAql_Name)
             .DropdownListEntries.add MinorAql_Name(i), MinorAql_Value(i)
          Next i
           For Each dropdownEntry In .DropdownListEntries
              If dropdownEntry.Text = Option_Selected Then
                 dropdownEntry.Select
                Exit For
             End If
          Next dropdownEntry
                                                   ' Debug.Print "MinorAql" & oCC.Range.Text
        End With
   End With
   '========================================================
    
      Url = WebBaseUrl & "/api/sqlhub/exec?appid=md_330&configName=GetReportQuantity"
      Call getHUBdata(Url, postJson, jsonResponse)
 
  
    Dim jsondata As String
 
    jsondata = JsonConverter.ConvertToJson(jsonResponse)
  
      Itms = jsonResponse("data").count
      If Itms = 0 Then
        msg = "No Quantity Breakdown data is available on Hub, Please input yourself!" & CHR(10) & "HUB" & UText(27809, 26377) & "Quanlity" & UText(26126, 32454, 25968, 25454) & "，" & UText(35831, 33258, 34892, 24405, 20837) & CHR(10)
        GoTo 50
      End If
      Dim PO_No(), Item_No(), Qty_No() As String
      ReDim PO_No(1 To Itms), Item_No(1 To Itms), Qty_No(1 To Itms)
    For i = 1 To Itms
        PO_No(i) = jsonResponse("data")(i)(HubKeyOrderNo())
        Item_No(i) = jsonResponse("data")(i)(HubKeyItemNo())
        Qty_No(i) = jsonResponse("data")(i)(HubKeyOrderQuantity())
    Next i
   
    tbn = getTableNumber(1)
    With ThisDocument.Tables(tbn)
       Rn = .Rows.count

       If Rn < Itms + 5 Then
         .Cell(.Rows.count - 1, .Columns.count).Select
         Selection.InsertRows Itms + 5 - Rn
       End If
         For i = 1 To Itms
        
          .Cell(i + 4, 1).Range.Text = PO_No(i)
          .Cell(i + 4, 2).Range.Text = Item_No(i)
          .Cell(i + 4, 3).Range.Text = Qty_No(i)
         Next i
       
    
    End With
      '========
   
    
         Dim AccountManagerID, AccountManagerNameC, AccountManagerNameE, AccountManagerEmail, AccountManagerMobile, AccountManagerPic, AccountManagerPicID, DeptID As String ', AccountManagerSkype
         Dim CSDeptManagerID, CSDeptManagerNameC, CSDeptManagerNameE, CSDeptManagerEmail, CSDeptManagerMobile, CSDeptManagerPic As String ', CSDeptManagerSkype
         Dim ddcode As String
         
         Dim OriginalTrackRevisions As Boolean
         Dim OriginalShowRevisions As Boolean
         OriginalTrackRevisions = WD.TrackRevisions
         OriginalShowRevisions = WD.ShowRevisions
   
         WD.TrackRevisions = False
         WD.ShowRevisions = False
         Set WD = ThisDocument
         tb = WD.Tables.count
         If WD.Tables(tb).Rows.count = 5 And WD.Tables(tb).Columns.count = 4 Then
            With WD.Tables(tb)
                postJson("ReportNo") = rpn
        
                 ' Url = WebBaseUrl & "/api/SelectPhotoInformation/SelectDdcode"
                  Url = WebBaseUrl & "/api/SelectPhotoInformation/SelectAccountManager"
                  Call getHUBdata(Url, postJson, jsonResponse)
                 'Dim json_str As String
                 'json_str = JsonConverter.ConvertToJson(jsonResponse)
                  ' Debug.Print json_str
   
                  AccountManagerNameC = jsonResponse("LASTNAME")
                  AccountManagerNameE = jsonResponse("LASTNAMEY")
                  AccountManagerEmail = jsonResponse("EMAIL")
                  'AccountManagerSkype = jsonResponse("Skype")
                  AccountManagerMobile = jsonResponse("MOBILE")
                  AccountManagerID = jsonResponse("id")
                  AccountManagerPicID = jsonResponse("resourceimageid")
                  AccountManagerPic = jsonResponse("MESSAGERURL")
                  DeptID = jsonResponse("SUBCOMID")
                  ddcode = jsonResponse("ddcode")
                
    Dim SHP As shape
    
                  .Cell(1, 1).Range.Delete
                     If AccountManagerPicID <> "" Then
                        .Cell(1, 1).Range.InlineShapes.AddPicture FileName:= _
                         WebBaseUrl & "/weaver/weaver.file.FileDownload?fileid=" & AccountManagerPicID & "&" & "ddcode=" & ddcode, LinkToFile:=False, _
                         SaveWithDocument:=True
                    
                        Set SHP = .Cell(1, 1).Range.InlineShapes(1).ConvertToShape
                        With SHP
                            .WrapFormat.Type = wdWrapInline
                            .AutoShapeType = msoShapeOval
                            .Line.Visible = msoFalse
                            .LockAspectRatio = msoTrue ' translated comment
                            .Height = 92.15  ' translated comment
                        End With
                    
                     Else
                       
                          .Cell(1, 1).Range.InlineShapes.AddPicture FileName:= _
                           WebBaseUrl & AccountManagerPic, LinkToFile:=False, _
                           SaveWithDocument:=True

                        Set SHP = .Cell(1, 1).Range.InlineShapes(1).ConvertToShape
                        With SHP
                            .WrapFormat.Type = wdWrapInline
                            .AutoShapeType = msoShapeOval
                            .Line.Visible = msoFalse
                            .LockAspectRatio = msoTrue ' translated comment
                            .Height = 92.15  ' translated comment
                        End With
                      
                    End If
                     ' If AccountManagerPicID  <> "" Then
                        .Cell(1, 2).Range.Text = AccountManagerNameE
                        .Cell(3, 2).Range.Text = "E:" & " " & AccountManagerEmail
                        '.Cell(4, 2).Range.Text = "Skype:" & " " & AccountManagerSkype
                        .Cell(4, 2).Range.Text = "M:" & " " & AccountManagerMobile
                      'End If
            
                      postJson("SUBCOMID") = DeptID
                      Url = WebBaseUrl & "/api/SelectPhotoInformation/SelectGeneralManager"
                      Call getHUBdata(Url, postJson, jsonResponse)
         
                      CSDeptManagerNameC = jsonResponse("LASTNAME")
                      CSDeptManagerNameE = jsonResponse("LASTNAMEY")
                      CSDeptManagerEmail = jsonResponse("EMAIL")
                      'CSDeptManagerSkype = jsonResponse("Skype")
                      CSDeptManagerMobile = jsonResponse("MOBILE")
                      CSDeptManagerID = jsonResponse("id")
                      CSDeptManagerPicID = jsonResponse("resourceimageid")
                      CSDeptManagerPic = jsonResponse("MESSAGERURL")
                     ddcode = jsonResponse("ddcode")
 
                     If CSDeptManagerPicID <> "" Then
                     .Cell(1, 3).Range.InlineShapes.AddPicture FileName:= _
                         WebBaseUrl & "/weaver/weaver.file.FileDownload?fileid=" & CSDeptManagerPicID & "&" & "ddcode=" & ddcode, LinkToFile:=False, _
                         SaveWithDocument:=True
                         
                        Set SHP = .Cell(1, 3).Range.InlineShapes(1).ConvertToShape
                        With SHP
                            .WrapFormat.Type = wdWrapInline
                            .AutoShapeType = msoShapeOval
                            .Line.Visible = msoFalse
                            .LockAspectRatio = msoTrue ' translated comment
                            .Height = 92.15  ' translated comment
                        End With
                        
                     Else
                   
                         ' If CSDeptManagerPic <> "" Then
                             .Cell(1, 3).Range.Delete
                             .Cell(1, 3).Range.InlineShapes.AddPicture FileName:= _
                              WebBaseUrl & CSDeptManagerPic, LinkToFile:=False, _
                               SaveWithDocument:=True
                        
                        Set SHP = .Cell(1, 3).Range.InlineShapes(1).ConvertToShape
                        With SHP
                            .WrapFormat.Type = wdWrapInline
                            .AutoShapeType = msoShapeOval
                            .Line.Visible = msoFalse
                            .LockAspectRatio = msoTrue ' translated comment
                            .Height = 92.15  ' translated comment
                        End With
                         ' End If
                  
                     End If
               
               
                      .Cell(1, 4).Range.Text = CSDeptManagerNameE
                      .Cell(3, 4).Range.Text = "E:" & " " & CSDeptManagerEmail
                      '.Cell(4, 4).Range.Text = "Skype:" & " " & CSDeptManagerSkype
                      .Cell(4, 4).Range.Text = "M:" & " " & CSDeptManagerMobile
            
                 End With
             Else
                msg = msg + UText(22312, 25253, 21578, 25991, 20214, 23614, 37096, 26410, 25214, 21040, 29992, 20110, 22635, 20805, 23458, 26381, 32463, 29702, 21644, 23458, 26381, 37096, 24635, 32463, 29702, 30340) & "4" & UText(21015) & "5" & UText(34892, 30340, 34920, 26684) & CHR(10) & _
                 "The table with 4 columns and 5 rows intended for filling in the customer service manager and  customer service department manager was not found at the end of the report file"
         
             End If
    
         
       
             WD.TrackRevisions = OriginalTrackRevisions
             WD.ShowRevisions = OriginalShowRevisions
50     Opentimes = GetPropertyValue("Opentimes")
    N = Val(Opentimes)
    N = N + 1
    StoretoProperty "Opentimes", CStr(N)

      ' ActiveWindow.ActivePane.View.Type = wdPrintView
      ThisDocument.Range(Start:=0, End:=0).Select
       MsgBox msg & "Import of Inspection Info Completed" & CHR(10) & UText(26816, 39564, 20449, 24687, 23548, 20837, 23436, 25104) & "！"
     
     
End Sub




Sub HandleContentControlSelection()
    Dim Occ As contentControl
    Dim selectedValue As String
    
   
    Set Occ = GetSelectedContentControl()
    
   
    If Occ.Tag = "Choose Sampling Standart" Then
       
        selectedValue = Occ.Range.Text
       
        MsgBox UText(36873, 20013, 30340, 20540, 26159) & "：" & selectedValue
    End If
End Sub
Function GetSelectedContentControl() As contentControl   ' translated comment
    Dim Occ As contentControl
    On Error Resume Next
    Set Occ = Selection.ContentControls(1)
    On Error GoTo 0
     Set GetSelectedContentControl = Occ
End Function

Function StoreJSON(propName As String, json_String As String)
Dim oProp As Variant
Dim bExists As Variant
   For Each oProp In ActiveDocument.CustomDocumentProperties
        If oProp.Name = propName Then
            oProp.Value = json_String
            bExists = True
            Exit For
        End If
    Next oProp
    If Not bExists Then
        ActiveDocument.CustomDocumentProperties.add Name:=propName, LinkToContent:=False, Type:=msoPropertyTypeString, Value:=json_String
    End If

        
End Function
Sub StoretoProperty(propName As String, propValue As String)
Dim oProp As Variant
Dim bExists As Variant
bExists = False
   For Each oProp In ActiveDocument.CustomDocumentProperties
        If oProp.Name = propName Then
            oProp.Value = propValue
            bExists = True
            Exit For
        End If
    Next oProp
    If Not bExists Then
        ActiveDocument.CustomDocumentProperties.add Name:=propName, LinkToContent:=False, Type:=msoPropertyTypeString, Value:=propValue
    End If
End Sub
Function GetPropertyValue(propName As String)
Dim oProp As DocumentProperty
Dim propValue As String
   For Each oProp In ActiveDocument.CustomDocumentProperties
        If oProp.Name = propName Then
            GetPropertyValue = oProp.Value
            Exit Function
        End If
    Next oProp
     GetPropertyValue = ""
End Function


Sub conclusion()
'
'by wyn conclusion
'
'
      
      ' translated comment
      
       Dim alr, msg, prompt As String
       Dim i, j, msgn As Integer
       Dim CF, DCF, NC, DNC, PD, DPD, NA As String
       Dim CFdis, NCdis, PDdis As String
       Dim rg As Range
       Dim Res, TBName(10), dres(10), ch, che As String
       Dim Rc, cc As Integer
       Dim qtyRow, qtyCol, k, p, Q, R, cnt, CeC, CRow As Integer
       Dim col(100, 9), scol(9), Nfw, Naw As Long
       Dim tem, txt(100, 10) As String
       Dim tb, tbn, SAP(10) As Integer
       Dim Ufw, Uaw As String
    
   TBName(1) = "*Quantity*"
   TBName(2) = "*Workmanship*"
   TBName(3) = "*Style, Color & Documentation*"
   TBName(4) = "*Field Test & Measurement*"
   TBName(5) = "*Shipping Mark & Packaging*"
   TBName(6) = "*Reference Note*"
   TBName(7) = "*Photos*"
   DCF = "Conform"
   DNC = "Not Conform"
   DPD = "Pending"
   NA = "NA"       'NANA，
   ch = "*Choose*"
   che = "*Choose*"
   msg = ""
   alt = ""
   tb = ThisDocument.Tables.count - 1
   ' translated comment
    
    
      
       prompt = "Ensure all inspection details has been input and  relusts of sub-items been  chosen correctly. If all are OK, click YES (left key)," _
      & "click NO (right key) will exit the program so you can revise the inspection details." & CHR(13) & _
      UText(35831, 39318, 20808, 30830, 35748, 21508, 20998, 39033, 30340, 25253, 21578, 20869, 23481, 26159, 24050, 22635, 22909, 24182, 36873, 25321, 20102, 32467, 35770) & "，" & UText(22914, 26524, 24050, 22635, 22909, 35831, 25353, 30830, 23450) & "，" & UText(21542, 21017, 25353, 21462, 28040, 36864, 20986, 32467, 35770, 35745, 31639, 21435, 22635, 20889, 20998, 39033, 20869, 23481) & "。"
    
      
      
      Key = MsgBox(prompt, vbOKCancel, UText(25552, 31034))
      If Key = vbCancel Then Exit Sub
    
  
   
    ' translated comment
    
     Call Dealingsap
     'SAP
     
     Selection.GoTo What:=wdGoToBookmark, Name:="sapcontent"

      With ThisDocument.Bookmarks("sapcontent").Range
      tNo = .Tables.count ' translated comment
       i = 0
       For T = 1 To tNo
         With ThisDocument.Bookmarks("sapcontent").Range.Tables(T)
         cc = .Columns.count
    
       If cc > 2 Then ' translated comment
        i = i + 1
         SAP(i) = .Rows.count
        End If
       End With
       Next T
      End With
       
       



  
   'Quantity
   tb = ThisDocument.Tables.count - 1
   For X = 1 To tb
     With ThisDocument.Range.Tables(X)
        If Trim(.Cell(1, 1).Range.Text) Like TBName(1) = True Then tbn = X
     End With
   Next X
  
   With ThisDocument.Range.Tables(tbn)
    
     qtyRow = .Rows.count

     qtyCol = .Columns.count
   
        ' translated comment
     CRow = qtyRow - 4
     CeC = ThisDocument.Tables(tbn).Range.Cells.count
     If (CeC - 23) / qtyCol <> CRow Then
        MsgBox "There is merge cells in Quantity table, the program can not deal with it, please split it to default frame and run program again." & CHR(13) & "Quantity" & UText(34920, 26684, 26377, 21512, 24182, 21333, 20803, 26684) & "，" & UText(31243, 24207, 19981, 33021, 36827, 34892, 32047, 21152, 35745, 31639) & "！！！," & UText(35831, 25286, 20998, 21333, 20803, 26684, 24674, 22797, 21040, 21021, 22987, 29366, 24577) & "。"
        For Q = 3 To 9
        scol(Q) = Val(.Cell(qtyRow, (Q - 1)).Range.Text)
        Next Q
        GoTo 12
     End If
        
     For p = qtyRow - 1 To 4 Step -1
         
        cnt = 0
        For Q = 1 To 9
        txt(p, Q) = Trim(.Cell(p, Q).Range.Text)
        col(p, Q) = Val(txt(p, Q))
        If Left(txt(p, Q), 1) = CHR(13) Then cnt = cnt + 1
        Next Q
        If cnt = 9 Then .Cell(p, 1).Range.Rows.Delete: GoTo 10
        
        If Left(txt(p, 1), 1) <> CHR(13) Or Left(txt(p, 2), 1) <> CHR(13) Then
        
          For Q = 3 To 9
          If Left(txt(p, Q), 1) = CHR(13) Then msg = msg + "Table of" & TBName(1) & "， row " & p & " without data, please check!" & CHR(13) & _
          TBName(1) & UText(34920, 26684, 31532) & p & UText(34892, 26377, 34920, 26684, 27809, 26377, 36755, 20837, 25968, 25454) & "，" & UText(35831, 26816, 26597, 36755, 20837) & "!" & CHR(13)
          Next Q
        End If
          
        If Left(txt(p, 1), 1) = CHR(13) And Left(txt(p, 2), 1) = CHR(13) Then
          For Q = 3 To 9
          If Left(txt(p, Q), 1) <> CHR(13) Then msg = msg + "Table of" & TBName(1) & "row" & p & " without P.O. No. or Item No.please check!" & CHR(13) & _
          TBName(1) & UText(34920, 26684, 31532) & p & UText(34892, 26377, 25968, 25454) & "，" & UText(20294, 21364, 27809, 26377, 36755, 20837) & "P.O. No. " & UText(25110) & "Item No." & UText(35831, 26816, 26597, 36755, 20837) & "!" & CHR(13)
          Next Q
        End If
        If msg <> "" Then MsgBox msg & "Please revice it and run program again！！！" & CHR(13) & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892, 31243, 24207) & "！！！": Exit Sub
      
        For Q = 3 To 9
        
       '.Cell(qtyRow, (q - 1)).Range.Text = 0
        scol(Q) = scol(Q) + col(p, Q)
        
        Next Q
    
10    Next p

   
    qtyRow = .Rows.count
    
     For R = 3 To 9
      If Val(.Cell(qtyRow, R - 1).Range.Text) <> scol(R) Then msg = msg + "Table:" & TBName(1) & " , Sum of column " & R & " was " & scol(R) & _
      " , however, the recent input data was " & Val(.Cell(qtyRow, R - 1).Range.Text) & CHR(13) & CHR(13) & _
      TBName(1) & UText(34920, 26684) & "total" & UText(25968, 25454, 31532) & R & UText(21015) & "，" & UText(32047, 35745, 25968, 25454, 26159) & "：" & scol(R) & UText(20294, 30446, 21069, 22635, 20889, 30340, 25968, 25454, 26159) & "：" & Val(.Cell(qtyRow, R - 1).Range.Text) & UText(20108, 32773, 19981, 30456, 31561) & CHR(13)
     Next R
     If msg <> "" Then MsgBox msg & "Please check and revise it" & CHR(13) & UText(35831, 26816, 26597, 20462, 25913) & "!", vbCritical: Exit Sub
     
     
    
12   ' translated comment

   .Cell(3, 3).Range.Text = .Cell(3, 4).Range.Text
   .Cell(3, 6).Range.Text = .Cell(3, 4).Range.Text
   .Cell(3, 8).Range.Text = .Cell(3, 4).Range.Text
   .Cell(3, 7).Range.Text = .Cell(3, 5).Range.Text
   .Cell(3, 9).Range.Text = .Cell(3, 5).Range.Text
     
   
     ' translated comment
     Ufw = ThisDocument.Tables(1).Cell(6, 4).Range.words(2)
     Ufw = Mid(Ufw, 1, Len(Ufw) - 1)
     Uaw = .Cell(3, 3).Range.words(2)
     If Uaw <> Ufw Then MsgBox "Unit of P.O. Quantity in table of first page was " & Ufw & ", however Unit in 1. Quantity table was" & Uaw _
     & ", they are not the same,please revise them and than run program again." & CHR(13) & _
     UText(25253, 21578, 39318, 39029, 34920, 26684) & "P.O. Quantity" & UText(30340, 20135, 21697, 21333, 20301, 26159) & Ufw & ", " & UText(32780) & "1. Quantity" & UText(34920, 26684, 30340, 20135, 21697, 21333, 20301, 26159) & Uaw _
     & UText(20004, 32773, 19981, 19968, 26679) & "，" & UText(35831, 20462, 25913, 21518, 20877, 27425, 36816, 34892, 26412, 31243, 24207) & "！！！" & CHR(13) & _
     UText(27880, 24847) & "1. Quantity" & UText(34920, 26684, 30340, 20135, 21697, 21333, 20301, 30340, 20462, 25913, 26159, 28857, 34920, 26684, 31532) & "4" & UText(21015, 30340) & "Shipment Qty" & UText(19979, 26041, 36827, 34892, 36873, 25321) & "，" & UText(36816, 34892, 31243, 24207, 21518, 34920, 26684, 20854, 20182, 22788, 30340, 21333, 20301, 23558, 33258, 34892, 21516, 27493, 20462, 25913) & "，" & UText(32780, 25253, 21578, 39318, 39029, 34920, 26684) & "P.O. Quantity" & UText(22788, 30340, 21333, 20301, 19981, 20250, 21516, 27493) & "，" & UText(38656, 35201, 25163, 24037, 21435, 20462, 25913) & "!": Exit Sub
      
     If ThisDocument.Tables(1).Cell(6, 4).Range.words(2) <> ThisDocument.Tables(1).Cell(9, 4).Range.words(2) Then MsgBox "In information table of first page, Unit of P.O. Quantity was " & _
     ThisDocument.Tables(1).Cell(6, 4).Range.words(2) & ", however Unit of sample size of product was" & ThisDocument.Tables(1).Cell(9, 4).Range.words(2) _
     & ", they are not the same,please revise them and than run program again." & CHR(13) & _
     UText(25253, 21578, 39318, 39029, 34920, 26684) & "P.O. Quantity" & UText(30340, 20135, 21697, 21333, 20301, 26159) & ThisDocument.Tables(1).Cell(6, 4).Range.words(2) & ", sample size" & UText(20013, 30340, 20135, 21697, 21333, 20301, 26159) & ThisDocument.Tables(1).Cell(9, 4).Range.words(2) _
     & UText(20004, 32773, 19981, 19968, 26679) & "，" & UText(35831, 20462, 25913, 21518, 20877, 27425, 36816, 34892, 26412, 31243, 24207) & "！！！" & CHR(13): Exit Sub
       
      
      'PO
    
     Nfw = Val(ThisDocument.Tables(1).Cell(6, 4).Range.words(1))

     Naw = Val(.Cell(qtyRow, 2).Range.Text)
     If Nfw <> Naw Then MsgBox "P.O. Quantity in table of first page was " & Nfw & " However P.O. Quantity in 1. Quantity table was " & Naw & _
     "they are not the same,please revise them and than run program again." & CHR(13) & UText(25253, 21578, 39318, 39029, 34920, 26684) & "P.O. Quantity" & UText(30340, 20135, 21697, 25968, 37327, 26159) & Nfw & ", " & UText(32780) & "1. Quantity" & UText(34920, 26684, 30340) & "PO Qty" & UText(25968, 37327, 26159) & Naw & _
     UText(20004, 32773, 19981, 19968, 26679) & "，" & UText(35831, 20462, 25913, 21518, 20877, 27425, 36816, 34892, 26412, 31243, 24207) & "！！！": Exit Sub
     
     
     If Val(.Cell(qtyRow, 5).Range.Text) > Val(.Cell(qtyRow, 3).Range.Text) Then MsgBox " Quantity table: the Packed Qty was larger than Shipment Qty，Please revise them and run program again." & CHR(13) & _
     " Quantity" & UText(34920, 26684) & ",Packed Qty" & UText(20135, 21697, 25968, 22823, 20110) & "Shipment Qty" & UText(20135, 21697, 25968) & "，" & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892, 31243, 24207) & "！！！": Exit Sub
     
     If Val(.Cell(qtyRow, 6).Range.Text) > Val(.Cell(qtyRow, 4).Range.Text) Then MsgBox " Quantity table: the Packed Qty was larger than Shipment Qty，Please revise them and run program again." & CHR(13) & _
     " Quantity" & UText(34920, 26684) & ",Packed Qty" & UText(21253, 35013, 25968, 22823, 20110) & "Shipment Qty" & UText(21253, 35013, 25968) & "，" & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892, 31243, 24207) & "！！！": Exit Sub
     
     ' translated comment
     If LCase(ThisDocument.Tables(1).Cell(9, 2).Range.Text) Like "*original inspection*" = True Or LCase(ThisDocument.Tables(1).Cell(9, 2).Range.Text) Like "*first inspection*" = True Then
         If Val(.Cell(qtyRow, 8).Range.Text) < Int(Sqr(Val(.Cell(qtyRow, 4).Range.Text))) Then MsgBox "table:" & TBName(1) & ",the selected carton number is less than sqrt of shipment carton number" & _
         Int(Sqr(scol(5))) & "please check and if you revise the data, run this program again！！！." & CHR(13) & TBName(1) & UText(34920, 26684, 20135, 21697, 25968, 37327, 20026) & scol(5) & UText(24320, 26041, 25968, 21462, 25972, 21518, 20026) & Int(Sqr(scol(5))) & UText(29616, 22312, 25277, 31665, 25968, 20026) & scol(9) & UText(26159, 21542, 22826, 23569) & "，" & UText(35831, 30830, 35748, 26159, 21542, 20462, 25913) & "，" & UText(33509, 26377, 20462, 25913, 35831, 25913, 21518, 20877, 36816, 34892, 31243, 24207), vbCritical, UText(35686, 21578)
     End If
     If ThisDocument.Tables(1).Cell(9, 2).Range.Text Like "*Re-inspection*" = True Or LCase(ThisDocument.Tables(1).Cell(9, 2).Range.Text) Like "*re-inspection*" = True Then
      If Val(.Cell(qtyRow, 8).Range.Text) < Int(Sqr(Val(.Cell(qtyRow, 4).Range.Text)) * 1.5) Then MsgBox "table:" & TBName(1) & "the selected carton number:" & scol(9) & " is less than sqrt of shipment carton number times 1.5= " & _
         Int(Sqr(scol(5)) * 1.5) & ",please check, and if you revise the data, run this program again！！！." & CHR(13) & TBName(1) & "table:" & UText(20135, 21697, 25968, 37327, 20026) & scol(5) & UText(24320, 26041, 25968, 20056, 20197) & "1.5" & UText(20493, 21462, 25972, 21518, 20026) & Int(Sqr(scol(5)) * 1.5) & UText(29616, 22312, 25277, 31665, 25968, 20026) & scol(9) & UText(26159, 21542, 22826, 23569) & "，" & UText(35831, 30830, 35748, 26159, 21542, 20462, 25913) & "，" & UText(33509, 26377, 20462, 25913, 35831, 25913, 21518, 20877, 36816, 34892, 31243, 24207), vbCritical, UText(35686, 21578)
     End If
     
     '，SAP
     If scol(4) <> scol(3) And SAP(1) = 1 Then msg = msg + TBName(1) & ": Qrder Quantity not eqaual to Shipment Quantity and without SAP，please check." & CHR(13) & TBName(1) & UText(34920, 26684) & "，Qrder Quantity" & UText(21644) & "Shipment Quantity" & UText(19981, 30456, 31561) & "，" & UText(21364, 27809, 26377, 20889) & "SAP，" & UText(35831, 26816, 26597) & "！" & CHR(13)
     If scol(6) < scol(4) * 0.8 And SAP(1) = 1 Then msg = msg + TBName(1) & ": Packed Quantity was smaller than 80% of Shipment Quantity and without SAP，Please check" & CHR(13) & TBName(1) & UText(34920, 26684) & "，Packed Quantity" & UText(23567, 20110) & "Shipment Quantity" & UText(30340) & "80%，" & UText(21364, 27809, 26377, 20889) & "SAP，" & UText(35831, 26816, 26597) & "！" & CHR(13)
     If Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1) = "Choose" Or Mid(.Cell(1, 2).Range.Text, 1, 4) = "Choose" Then msg = msg + TBName(1) & "You have not choose the result" & CHR(13) & _
     TBName(1) & UText(34920, 26684, 27809, 26377, 36873, 25321, 32467, 35770) & "!" & CHR(13): GoTo 13
     
     If SAP(1) > 1 And Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1) = "Conform" Then msg = msg + TBName(1) & ": There is SAP while the reslut was Conform,." & CHR(13) & _
     TBName(1) & UText(34920, 26684, 32467, 35770, 20026) & " Conform," & UText(20294, 21364, 26377) & "SAP" & CHR(13)
     If SAP(1) = 1 And Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1) <> "Conform" Then msg = msg + TBName(1) & ": There is not SAP while the resulit was not Conform." & CHR(13) & _
     TBName(1) & UText(34920, 26684, 32467, 35770, 19981, 26159) & "Conform," & UText(20294, 27809, 26377) & "SAP" & CHR(13)
13   If msg <> "" Then MsgBox msg & "Please revise them" & CHR(13) & UText(35831, 26816, 26597, 20462, 25913) & "！！！", vbCritical: Exit Sub
 End With
 
     'sample sizeAQL
 
       Dim Cn, Typ, Styp, lvl As Integer
       Dim itype, Aqlt, Smst, tmp As String
       Dim Wn, Pn, WL, Coln, Rown, ss, grp, Agrp, Hgrp, Cgrp, SmpS(10), AAc(10), ADef(10), Scode, Ass, Y As Integer
       Dim AQ(20), AQL(10), Acr, AMa, Ami, AFn, AFnMa As Single
       Dim rgs As Range
       Dim Lots(20) As Long
       Dim CD(20), LG(20, 10) As Integer
       Dim AC(20, 10), MV(20, 10) As Integer

       'Lots()，，-1
       Lots(1) = 2
       Lots(2) = 9
       Lots(3) = 16
       Lots(4) = 26
       Lots(5) = 51
       Lots(6) = 91
       Lots(7) = 151
       Lots(8) = 281
       Lots(9) = 501
       Lots(10) = 1201
       Lots(11) = 3201
       Lots(12) = 10001
       Lots(13) = 35001
       Lots(14) = 150001
       Lots(15) = 500001
    
       
       'CD()，A,B,C,....R"
        CD(0) = 0
        CD(1) = 2
        CD(2) = 3
        CD(3) = 5:
        CD(4) = 8:
        CD(5) = 13:
        CD(6) = 20
        CD(7) = 32
        CD(8) = 50
        CD(9) = 80
        CD(10) = 125
        CD(11) = 200
        CD(12) = 315
        CD(13) = 500
        CD(14) = 800
        CD(15) = 1250
        CD(16) = 2000
        
       'AQL() AQL " & UText(20540) & "," & UText(21482, 21462, 24120, 29992, 30340) & " 1.0" & UText(21040) & " 6.5
        AQ(0) = 0
        AQ(1) = 1#
        AQ(2) = 1.5
        AQ(3) = 2.5
        AQ(4) = 4#
        AQ(5) = 6.5
        
        
        
        ' LG() " & UText(25209, 37327, 23545, 24212, 30340, 26679, 26412, 20195, 30721, 23545, 24212, 30340, 26679, 26412, 20195, 30721) & "，" & UText(31532, 19968, 32500, 25968, 23383, 20195, 34920, 25209, 37327, 25968, 32534, 21495) & "，" & UText(31532, 20108, 32500, 25968, 23383, 34920, 31034) & " S-1,S-2，S-3，S-4，G-I，G-II,G-III
      
       LG(0, 1) = CD(0):      LG(0, 2) = CD(0):       LG(0, 3) = CD(0)
   
       LG(1, 1) = CD(1):      LG(1, 2) = CD(1):       LG(1, 3) = CD(1):    LG(1, 4) = CD(1):     LG(1, 5) = CD(1):   LG(1, 6) = CD(1):       LG(1, 7) = CD(2):
       LG(2, 1) = CD(1):      LG(2, 2) = CD(1):       LG(2, 3) = CD(1):    LG(2, 4) = CD(1):     LG(2, 5) = CD(1):   LG(2, 6) = CD(2):       LG(2, 7) = CD(3):
       LG(3, 1) = CD(1):      LG(3, 2) = CD(1):       LG(3, 3) = CD(2):    LG(3, 4) = CD(2):     LG(3, 5) = CD(2):   LG(3, 6) = CD(3):       LG(3, 7) = CD(4):
       LG(4, 1) = CD(1):      LG(4, 2) = CD(2):       LG(4, 3) = CD(2):    LG(4, 4) = CD(3):     LG(4, 5) = CD(3):   LG(4, 6) = CD(4):       LG(4, 7) = CD(5):
       LG(5, 1) = CD(2):      LG(5, 2) = CD(2):       LG(5, 3) = CD(3):    LG(5, 4) = CD(3):     LG(5, 5) = CD(3):   LG(5, 6) = CD(5):       LG(5, 7) = CD(6):
       LG(6, 1) = CD(2):      LG(6, 2) = CD(2):       LG(6, 3) = CD(3):    LG(6, 4) = CD(4):     LG(6, 5) = CD(4):   LG(6, 6) = CD(6):       LG(6, 7) = CD(7):
       LG(7, 1) = CD(2):      LG(7, 2) = CD(3):       LG(7, 3) = CD(4):    LG(7, 4) = CD(5):     LG(7, 5) = CD(5):   LG(7, 6) = CD(7):       LG(7, 7) = CD(8):
       LG(8, 1) = CD(2):      LG(8, 2) = CD(3):       LG(8, 3) = CD(4):    LG(8, 4) = CD(5):     LG(8, 5) = CD(6):   LG(8, 6) = CD(8):       LG(8, 7) = CD(9):
       LG(9, 1) = CD(3):      LG(9, 2) = CD(3):       LG(9, 3) = CD(5):    LG(9, 4) = CD(6):     LG(9, 5) = CD(7):   LG(9, 6) = CD(9):       LG(9, 7) = CD(10):
       LG(10, 1) = CD(3):     LG(10, 2) = CD(4):      LG(10, 3) = CD(5):   LG(10, 4) = CD(7):    LG(10, 5) = CD(8):  LG(10, 6) = CD(10):     LG(10, 7) = CD(11):
       LG(11, 1) = CD(3):     LG(11, 2) = CD(4):      LG(11, 3) = CD(6):   LG(11, 4) = CD(7):    LG(11, 5) = CD(9):  LG(11, 6) = CD(11):     LG(11, 7) = CD(12):
       LG(12, 1) = CD(3):     LG(12, 2) = CD(4):      LG(12, 3) = CD(6):   LG(12, 4) = CD(8):    LG(12, 5) = CD(10): LG(12, 6) = CD(12):     LG(12, 7) = CD(13):
       LG(13, 1) = CD(4):     LG(13, 2) = CD(5):      LG(13, 3) = CD(7):   LG(13, 4) = CD(9):    LG(13, 5) = CD(11): LG(13, 6) = CD(13):     LG(13, 7) = CD(14):
       LG(14, 1) = CD(4):     LG(14, 2) = CD(5):      LG(14, 3) = CD(7):   LG(14, 4) = CD(9):    LG(14, 5) = CD(12): LG(14, 6) = CD(14):     LG(14, 7) = CD(15):
       LG(15, 1) = CD(4):     LG(15, 2) = CD(5):      LG(15, 3) = CD(8):   LG(15, 4) = CD(10):   LG(15, 5) = CD(13): LG(15, 6) = CD(15):     LG(15, 7) = CD(16)
    

  With ThisDocument.Tables(1)
      
       SamP = .Cell(7, 4).Range.Text

       '" & UText(26816, 26597, 27491, 24120, 26816, 26597) & "，" & UText(37325, 26816, 25277, 26679, 26041, 26696, 26159, 21542, 20998, 21035, 26159) & "NORMAL" & UText(21644) & "TIGHTENED，" & UText(24182, 25552, 37266) & "。
       
            If SamP Like "*Double*" = True Then MsgBox "the sample plan in Information table of first page was double sample plan instead of single plan the progran can not work out the Sample Size and AQL, please deal with it by yourself." & CHR(13) & _
            "Information " & UText(34920, 25152, 22635, 20889, 30340, 25277, 26679, 26041, 26696, 26159, 20108, 27425, 25277, 26679, 26041, 26696) & "，" & UText(19981, 26159, 19968, 27425, 25277, 26679, 26041, 26696) & "，Sample Size" & UText(21644) & "AQL" & UText(35831, 25163, 24037, 26597, 35810) & "!!!": GoTo 100
            If SamP Like "*Normal*" = True Then Styp = 1                            '" & UText(27491, 24120, 26816, 39564) & "
            If SamP Like "*Tightened*" = True Then Styp = 2                           '" & UText(21152, 20005, 26816, 39564) & "
            If SamP Like "*Fixed*" = True Then
                cnt = 0
                For i = 1 To 16
                    If CD(i) = Val(.Cell(9, 4).Range.words(1)) Then Styp = 1: GoTo 15                   '" & UText(22266, 23450, 25277, 26679, 19988) & "Sample Size" & UText(21018, 22909, 26159, 25277, 26679, 34920, 20013, 26377, 23545, 24212) & "
                    If CD(i) <> Val(.Cell(9, 4).Range.words(1)) Then cnt = cnt + 1
                Next i
                 If cnt = 16 Then Styp = 3   '==============
15          End If
            
            If SamP Like "*Normal*" = False And SamP Like "*Tightened*" = False And SamP Like "*Fixed*" = False Then MsgBox _
            "The inspection sample plan if Information table of first page was neither Normal, nor Tightened, nor Fixed sample plan, the program can not work out the Sample Size and Ac value, please deal with it yourself." & CHR(13) & _
            "Information " & UText(34920, 25152, 22635, 20889, 30340, 26816, 39564, 26041, 26696) & "，" & UText(19981, 26159, 27491, 24120) & "（Normal）" & UText(26816, 39564) & "，" & UText(21152, 20005) & "（Tightened）" & UText(26816, 39564) & " " & UText(25110, 22266, 23450) & "（Fixed)" & UText(25277, 26679, 26816, 39564) & "," & UText(31243, 24207, 26080, 27861, 33258, 21160, 31639, 20986) & "Sample Size" & UText(21644) & "Ac" & UText(20540) & "," & UText(35831, 25163, 24037, 22788, 29702) & "，" & UText(25110, 26816, 26597, 26159, 21542, 26377, 36755, 20837, 38169, 35823), vbCritical: GoTo 100
       
           
          
    Select Case Styp
          
    Case 1   '" & UText(27491, 24120, 26816, 39564) & "
       
       
       'AQL AC" & UText(20195, 30721, 26597, 35810) & "
       'AC(),AQL " & UText(20805, 25910, 20540) & "，" & UText(31532, 19968, 32500, 34920, 31034) & " " & UText(26679, 26412, 20195, 30721, 39034, 24207) & "，" & UText(31532, 20108, 32500, 34920, 31034) & "AQL " & UText(20540) & "1.0,1.5,...
       ' MV()," & UText(36716, 31227, 30340, 19978, 19979, 31661, 22836) & "，" & UText(20195, 34920, 25913, 21464, 21518, 30340) & " Sample Size
       AC(0, 0) = CD(0)
       AC(0, 1) = CD(0):       AC(0, 2) = CD(0):   AC(0, 3) = CD(0):   AC(0, 4) = CD(0):   AC(0, 5) = CD(0):
       
       AC(1, 1) = 0:           AC(1, 2) = 0:       AC(1, 3) = 0:       AC(1, 4) = 0:       AC(1, 5) = 0:
       MV(1, 1) = CD(5):       MV(1, 2) = CD(4):   MV(1, 3) = CD(3):   MV(1, 4) = CD(2):
       AC(2, 1) = 0:           AC(2, 2) = 0:       AC(2, 3) = 0:       AC(2, 4) = 0:       AC(2, 5) = 0:
       MV(2, 1) = CD(5):       MV(2, 2) = CD(4):   MV(2, 3) = CD(3):                       MV(2, 5) = CD(1):
       AC(3, 1) = 0:           AC(3, 2) = 0:       AC(3, 3) = 0:       AC(3, 4) = 0:       AC(3, 5) = 1:
       MV(3, 1) = CD(5):       MV(3, 2) = CD(4):                       MV(3, 4) = CD(2):   MV(3, 5) = CD(4):
       AC(4, 1) = 0:           AC(4, 2) = 0:       AC(4, 3) = 0:       AC(4, 4) = 1:       AC(4, 5) = 1:
       MV(4, 1) = CD(5):                           MV(4, 3) = CD(3):   MV(4, 4) = CD(5):
       AC(5, 1) = 0:           AC(5, 2) = 0:       AC(5, 3) = 1:       AC(5, 4) = 1:       AC(5, 5) = 2:
                               MV(5, 2) = CD(4):   MV(5, 3) = CD(6):
       AC(6, 1) = 0:           AC(6, 2) = 1:       AC(6, 3) = 1:       AC(6, 4) = 2:       AC(6, 5) = 3:
       MV(6, 1) = CD(5):       MV(6, 2) = CD(7):
       AC(7, 1) = 1:           AC(7, 2) = 1:       AC(7, 3) = 2:       AC(7, 4) = 3:       AC(7, 5) = 5:
       MV(7, 1) = CD(8)
       AC(8, 1) = 1:           AC(8, 2) = 2:       AC(8, 3) = 3:       AC(8, 4) = 5:       AC(8, 5) = 7:
        
       AC(9, 1) = 2:           AC(9, 2) = 3:       AC(9, 3) = 5:       AC(9, 4) = 7:       AC(9, 5) = 10:
       
       AC(10, 1) = 3:          AC(10, 2) = 5:      AC(10, 3) = 7:      AC(10, 4) = 10:     AC(10, 5) = 14:
       
       AC(11, 1) = 5:          AC(11, 2) = 7:      AC(11, 3) = 10:     AC(11, 4) = 14:     AC(11, 5) = 21:
        
       AC(12, 1) = 7:          AC(12, 2) = 10:     AC(12, 3) = 14:     AC(12, 4) = 21:     AC(12, 5) = 21:
                                                                                           MV(12, 5) = CD(11):
       AC(13, 1) = 10:         AC(13, 2) = 14:     AC(13, 3) = 21:     AC(13, 4) = 21:     AC(13, 5) = 21:
                                                                       MV(13, 4) = CD(12): MV(13, 5) = CD(11):
       AC(14, 1) = 14:         AC(14, 2) = 21:     AC(14, 3) = 21:     AC(14, 4) = 21:     AC(14, 5) = 21:
                                                   MV(14, 3) = CD(13): MV(14, 4) = CD(12): MV(14, 5) = CD(11):
       AC(15, 1) = 21:         AC(15, 2) = 21:     AC(15, 3) = 21:     AC(15, 4) = 21:     AC(15, 5) = 21:
                               MV(15, 2) = CD(14): MV(15, 3) = CD(13): MV(15, 4) = CD(12): MV(15, 5) = CD(11):
       AC(16, 1) = 21:         AC(16, 2) = 21:     AC(16, 3) = 21:     AC(16, 4) = 21:     AC(16, 5) = 21:
       MV(16, 1) = CD(15):     MV(16, 2) = CD(14): MV(16, 3) = CD(13): MV(16, 4) = CD(12): MV(16, 5) = CD(11):
                            
   Case 2  '" & UText(21152, 20005, 26816, 39564) & "
   
     
       'AQL AC" & UText(20195, 30721, 26597, 35810) & "
       'AC(),AQL " & UText(20805, 25910, 20540) & "，" & UText(31532, 19968, 32500, 34920, 31034) & " " & UText(26679, 26412, 20195, 30721, 39034, 24207) & "，" & UText(31532, 20108, 32500, 34920, 31034) & "AQL " & UText(20540) & "1.0,1.5,...
       ' MV()," & UText(36716, 31227, 30340, 19978, 19979, 31661, 22836) & "，" & UText(20195, 34920, 25913, 21464, 21518, 30340) & " Sample Size
       AC(0, 0) = CD(0)
       AC(0, 1) = CD(0):       AC(0, 2) = CD(0):   AC(0, 3) = CD(0):   AC(0, 4) = CD(0):   AC(0, 5) = CD(0):
       
       AC(1, 1) = 0:           AC(1, 2) = 0:       AC(1, 3) = 0:       AC(1, 4) = 0:       AC(1, 5) = 0:
       MV(1, 1) = CD(6):       MV(1, 2) = CD(5):   MV(1, 3) = CD(4):   MV(1, 4) = CD(3):   MV(1, 5) = CD(2):
       AC(2, 1) = 0:           AC(2, 2) = 0:       AC(2, 3) = 0:       AC(2, 4) = 0:       AC(2, 5) = 0:
       MV(2, 1) = CD(6):       MV(2, 2) = CD(5):   MV(2, 3) = CD(4):   MV(2, 4) = CD(3):
       AC(3, 1) = 0:           AC(3, 2) = 0:       AC(3, 3) = 0:       AC(3, 4) = 0:       AC(3, 5) = 1:
       MV(3, 1) = CD(6):       MV(3, 2) = CD(5):   MV(3, 3) = CD(4):                       MV(3, 5) = CD(5):
       AC(4, 1) = 0:           AC(4, 2) = 0:       AC(4, 3) = 0:       AC(4, 4) = 1:       AC(4, 5) = 1:
       MV(4, 1) = CD(6):       MV(4, 2) = CD(5):                       MV(4, 4) = CD(6):   MV(4, 5) = CD(5):
       AC(5, 1) = 0:           AC(5, 2) = 0:       AC(5, 3) = 1:       AC(5, 4) = 1:       AC(5, 5) = 1:
       MV(5, 1) = CD(6):                           MV(5, 3) = CD(7):   MV(5, 4) = CD(6):
       AC(6, 1) = 0:           AC(6, 2) = 1:       AC(6, 3) = 1:       AC(6, 4) = 1:       AC(6, 5) = 2:
                               MV(6, 2) = CD(8):   MV(6, 3) = CD(7):
       AC(7, 1) = 1:           AC(7, 2) = 1:       AC(7, 3) = 1:       AC(7, 4) = 2:       AC(7, 5) = 3:
       MV(7, 1) = CD(9):       MV(7, 2) = CD(8):
       AC(8, 1) = 1:           AC(8, 2) = 1:       AC(8, 3) = 2:       AC(8, 4) = 3:       AC(8, 5) = 5:
       MV(8, 1) = CD(9):
       AC(9, 1) = 1:           AC(9, 2) = 2:       AC(9, 3) = 3:       AC(9, 4) = 5:       AC(9, 5) = 8:
        
       AC(10, 1) = 2:          AC(10, 2) = 3:      AC(10, 3) = 5:      AC(10, 4) = 8:      AC(10, 5) = 12:
       
       AC(11, 1) = 3:          AC(11, 2) = 5:      AC(11, 3) = 8:     AC(11, 4) = 12:      AC(11, 5) = 18:
        
       AC(12, 1) = 5:          AC(12, 2) = 8:      AC(12, 3) = 12:     AC(12, 4) = 18:     AC(12, 5) = 18:
                                                                                           MV(12, 5) = CD(11):
       AC(13, 1) = 8:          AC(13, 2) = 12:     AC(13, 3) = 18:     AC(13, 4) = 18:     AC(13, 5) = 18:
                                                                       MV(13, 4) = CD(12): MV(13, 5) = CD(11):
       AC(14, 1) = 12:         AC(14, 2) = 18:     AC(14, 3) = 18:     AC(14, 4) = 18:     AC(14, 5) = 18:
                                                   MV(14, 3) = CD(13): MV(14, 4) = CD(12): MV(14, 5) = CD(11):
       AC(15, 1) = 18:         AC(15, 2) = 18:     AC(15, 3) = 18:     AC(15, 4) = 18:     AC(15, 5) = 18:
                               MV(15, 2) = CD(14): MV(15, 3) = CD(13): MV(15, 4) = CD(12): MV(15, 5) = CD(11):
       AC(16, 1) = 18:         AC(16, 2) = 18:     AC(16, 3) = 18:     AC(16, 4) = 18:     AC(16, 5) = 18:
       MV(16, 1) = CD(15):     MV(16, 2) = CD(14): MV(16, 3) = CD(13): MV(16, 4) = CD(12): MV(16, 5) = CD(11):
   
    '-----------FIXED SAMPLE " & UText(19988, 27809, 26377, 21487, 23545, 24212, 30340) & "AQL
    Case 3
 
    SmpS(0) = Val(.Cell(9, 4).Range.words(1))
    SmpS(1) = SmpS(0)
    SmpS(2) = SmpS(0)
 
    GoTo 110
    
   '-------------------------
   
   
   End Select
                            
       
       
         '" & UText(26680, 23545, 26816, 26597, 27700, 24179, 30340, 36755, 20837) & "
       If .Cell(8, 4).Range.Text Like "*G-III*" = True Or .Cell(8, 4).Range.Text Like "*G-111*" = True Then
       lvl = 7
       ElseIf .Cell(8, 4).Range.Text Like "*G-II*" = True Or .Cell(8, 4).Range.Text Like "*G-11*" = True Then
       lvl = 6
       ElseIf .Cell(8, 4).Range.Text Like "*G-I*" = True Or .Cell(8, 4).Range.Text Like "*G-1*" = True Then
       lvl = 5
       ElseIf .Cell(8, 4).Range.Text Like "*S-4*" = True Or .Cell(8, 4).Range.Text Like "*s-4*" = True Then
       lvl = 4
       ElseIf .Cell(8, 4).Range.Text Like "*S-3*" = True Or .Cell(8, 4).Range.Text Like "*s-3*" = True Then
       lvl = 3
       ElseIf .Cell(8, 4).Range.Text Like "*S-2*" = True Or .Cell(8, 4).Range.Text Like "*s-2*" = True Then
       lvl = 2
       ElseIf .Cell(8, 4).Range.Text Like "*S-1*" = True Or .Cell(8, 4).Range.Text Like "*s-1*" = True Then
       lvl = 1
        ElseIf SamP Like "*Fixed*" = False Then '===
        MsgBox " The inspection level was neither  S-1 to S-4，nor G-I to G-III，the program can not deal with it, please deal with it by yourself." & CHR(13) & _
       UText(26816, 26597, 27700, 24179) & "，" & UText(19981, 26159) & "S-1" & UText(21040) & "S-4，" & UText(20063, 19981, 26159) & "G-I，G-II" & UText(25110) & "G-III，" & UText(26080, 27861, 35745, 31639) & "，" & UText(35831, 26816, 26597, 36755, 20837, 26159, 21542, 31526, 21512, 35201, 27714, 25110, 25163, 24037, 35745, 31639): Exit Sub
       Else
       Styp = 3
       GoTo 60
       End If
   
       
       '" & UText(21021, 22987, 21270) & "
       Acr = 0
   
       AMa = 0
  
       Ami = 0
       
     '" & UText(28040, 38500) & "AQL" & UText(26684, 20013, 30340, 31354, 26684) & "
     Cn = .Cell(10, 4).Range.Characters.count
     For i = Cn To 1 Step -1
     If .Cell(10, 4).Range.Characters(i) = " " Then .Cell(10, 4).Range.Characters(i) = ""
     Next i
     
      Aqlt = .Cell(10, 4).Range.Text

       Set rg = .Cell(10, 4).Range
       Wn = rg.words.count
       WL = Len(Aqlt)

     
       '" & UText(21028, 23450) & "AQL" & UText(36755, 20837, 26041, 24335, 26159, 21542, 31526, 21512, 35201, 27714) & "
       
        If Wn = 18 Then
           ' If Aqlt Like "*Critical*" = True And Aqlt Like "*Major*" = True And Aqlt Like "*Minor*" = True Then
           
            If rg.words(1) = "Critical" And rg.words(7) = "Major" = True And rg.words(13) = "Minor" Then

            Acr = Val(Mid(Aqlt, 10, 3))
            AMa = Val(Mid(Aqlt, 20, 3))      '" & UText(21462) & "Ma
            Ami = Val(Mid(Aqlt, 30, 3))
        
            ElseIf rg.words(1) = "Cr" And rg.words(7) = "Ma" = True And rg.words(13) = "Mi" Then
            
            'ElseIf Aqlt Like "*Cr*" = True And Aqlt Like "*Ma*" = True And Aqlt Like "*Mi*" = True Then
            

                    'aql " & UText(31867, 22411) & "，1， Cr,ma,mi
            Acr = Val(Mid(Aqlt, 4, 3))
            AMa = Val(Mid(Aqlt, 11, 3))      '" & UText(21462) & "Ma
            Ami = Val(Mid(Aqlt, 18, 3))
    
            Else
               
            MsgBox "You should input AQL into Information table on first page in the format of such as Critical:0,Major:1.0,Minor:2.5 or Cr:0,Ma:1.0,Mi:4.0" & CHR(13) & _
            "Information " & UText(34920, 26684) & "，AQL" & UText(26410, 23436, 20840, 36755, 20837, 25110, 36755, 20837, 26684, 24335, 19981, 31526, 35201, 27714) & "，" & UText(19981, 33021, 35745, 31639, 20986) & "Sample Size" & UText(21644) & "Ac！" & _
            UText(27491, 30830, 26684, 24335, 31867, 20284, 20110) & "Critical:0,Major:1.0,Minor:2.5" & UText(25110) & "Cr:0,Ma:1.0,Mi:4.0，" & UText(27880, 24847, 27599, 20010, 32570, 38519, 23383, 27597, 21518, 29992, 19968, 20010, 20882, 21495) & "，" & UText(32570, 38519, 25968, 23383, 21518) & "，" & UText(24517, 39035, 19988, 21482, 26377, 19968, 20010, 36887, 21495) & "，" & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892) & "，" & UText(25110, 25163, 24037, 35745, 31639), vbCritical: GoTo 100
           End If
         End If
      
         If Wn = 16 Then
           

            'If Aqlt Like "*Critical*" = True And Aqlt Like "*Major*" = True And Aqlt Like "*Minor*" = True Then
            
            If rg.words(1) = "Critical" And rg.words(5) = "Major" = True And rg.words(11) = "Minor" Then
            Acr = Val(Mid(Aqlt, 10, 1))
            AMa = Val(Mid(Aqlt, 18, 3))      '" & UText(21462) & "Ma
            Ami = Val(Mid(Aqlt, 28, 3))
            
             ElseIf rg.words(1) = "Cr" And rg.words(5) = "Ma" = True And rg.words(11) = "Mi" Then
           ' ElseIf Aqlt Like "*Cr*" = True And Aqlt Like "*Ma*" = True And Aqlt Like "*Mi*" = True Then
                    'aql " & UText(31867, 22411) & "，1， Cr,ma,mi
            Acr = Val(Mid(Aqlt, 4, 1))
            AMa = Val(Mid(Aqlt, 9, 3))      '" & UText(21462) & "Ma
            Ami = Val(Mid(Aqlt, 16, 3))

          Else
               
            MsgBox "You should input AQL into Information table on first page in the format of such as Critical:0,Major:1.0,Minor:2.5 or Cr:0,Ma:1.0,Mi:4.0" & CHR(13) & _
            "Information " & UText(34920, 26684) & "，AQL" & UText(26410, 23436, 20840, 36755, 20837, 25110, 36755, 20837, 26684, 24335, 19981, 31526, 35201, 27714) & "，" & UText(19981, 33021, 35745, 31639, 20986) & "Sample Size" & UText(21644) & "Ac！" & _
            UText(27491, 30830, 26684, 24335, 31867, 20284, 20110) & "Critical:0,Major:1.0,Minor:2.5" & UText(25110) & "Cr:0,Ma:1.0,Mi:4.0，" & UText(27880, 24847, 27599, 20010, 32570, 38519, 23383, 27597, 21518, 29992, 19968, 20010, 20882, 21495) & "，" & UText(32570, 38519, 25968, 23383, 21518) & "，" & UText(24517, 39035, 19988, 21482, 26377, 19968, 20010, 36887, 21495) & "，" & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892) & "，" & UText(25110, 25163, 24037, 35745, 31639), vbCritical: GoTo 100
          End If
        End If
     
       
       
       
       If Wn <> 16 And Wn <> 18 Then MsgBox "You should input AQL into Information table on first page in the format of such as Critical:0,Major:1.0,Minor:2.5 or Cr:0,Ma:1.0,Mi:4.0" & CHR(13) & _
            "Information " & UText(34920, 26684) & "，AQL" & UText(26410, 23436, 20840, 36755, 20837, 25110, 36755, 20837, 26684, 24335, 19981, 31526, 35201, 27714) & "，" & UText(19981, 33021, 35745, 31639, 20986) & "Sample Size" & UText(21644) & "Ac！" & _
            UText(27491, 30830, 26684, 24335, 31867, 20284, 20110) & "Critical:0,Major:1.0,Minor:2.5" & UText(25110) & "Cr:0,Ma:1.0,Mi:4.0，" & UText(27880, 24847, 27599, 20010, 32570, 38519, 23383, 27597, 21518, 29992, 19968, 20010, 20882, 21495) & "，" & UText(32570, 38519, 25968, 23383, 21518) & "，" & UText(24517, 39035, 19988, 21482, 26377, 19968, 20010, 36887, 21495) & "，" & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892) & "，" & UText(25110, 25163, 24037, 35745, 31639), vbCritical: GoTo 100
      
          '" & UText(25226) & "AQL" & UText(20998, 32452) & "
       If Acr = 0 Then Y = 1 'Cr=0 " & UText(26102) & "，" & UText(21482, 31639) & "MA，MI，" & UText(19981, 31561, 20110) & "0" & UText(26102, 20840, 31639) & "
       If Acr <> 0 Then Y = 0
  
       
       AQL(0) = Acr
  
       AQL(1) = AMa
   
       AQL(2) = Ami
       
       '" & UText(26816, 39564, 25209, 37327, 25968, 26159, 21542, 27491, 30830) & "

      'Sqty = InputBox("Shipment Quanty")
      
       SQty = scol(4)                'qty" & UText(34920, 26684, 20013, 30340) & "shipment Qty" & UText(32047, 21152, 20540) & "

                           

       If SQty < Lots(1) Then MsgBox "The batch size was too small." & CHR(13) & UText(25209, 37327, 25968, 22826, 23567) & "，" & UText(26412, 31243, 24207, 19981, 33021, 35745, 31639) & "，" & UText(35831, 25163, 24037, 35745, 31639): GoTo 100
       If SamP Like "*Fixed*" = False And SQty > Lots(15) Then MsgBox "The batch size was too large and the program can not deal with it please deal with it yourself." & CHR(13) & _
       UText(25209, 37327, 25968, 22826, 22823) & "，" & UText(26412, 31243, 24207, 19981, 33021, 35745, 31639) & "，" & UText(35831, 25163, 24037, 35745, 31639): GoTo 100
    
       '＊＊＊" & UText(22788, 29702) & "fixed sample " & UText(21448, 21018, 22909, 26159) & "AQL" & UText(34920, 20013, 30340, 24773, 20917) & "
       If Styp = 1 And SamP Like "*Fixed*" = True Then
           For i = 1 To 15
             If Val(.Cell(9, 4).Range.words(1)) = CD(i) Then grp = i: GoTo 20
           Next i
        End If

  
        
        '" & UText(31639, 25209, 37327, 25968, 22312, 27599, 19968, 24352, 34920) & "LotSiZE" & UText(30340, 31532, 20960, 32452) & "
       For i = 1 To 15
       
       If SQty >= Lots(i) And SQty < Lots(i + 1) Then grp = i: GoTo 20
       
       Next i
       
    '" & UText(31639) & "AQLSampleSiz" & UText(22312, 31532, 20108, 24352, 34920, 30340, 32437, 21521, 31532, 20960, 32452) & "
      
20     Scode = LG(grp, lvl)
       For i = 1 To 16
       If Scode = CD(i) Then Hgrp = i: GoTo 25
       Next i

25    For k = Y To 2  'k " & UText(23545, 24212, 23454, 38469) & "AQL" & UText(30340) & "Ma,Mi
       
        
        '" & UText(31639, 23454, 38469, 39564, 36135) & "AQL" & UText(22312, 31532, 20108, 24352, 34920, 27178, 21521, 30340, 32452, 25968) & "
         Cn = 0
         For i = 1 To 5
           If AQL(k) = AQ(i) Then Agrp = i: GoTo 30
           If AQL(k) <> AQ(i) And AQL(k) <> 0 Then Cn = Cn + 1 '" & UText(21028, 23450) & "AQL " & UText(26159, 21542, 21015, 20837, 34920, 26684, 33539, 22260) & "
         Next i
   
        If Cn = 5 Then MsgBox "The program can only deal with the AQL  of 1.0，1.5，2.5，4.0，6.5, your AQL beyond this scope please deal with it yourself." & CHR(13) & _
        "Information " & UText(34920, 26684) & ",AQL" & UText(36755, 20837, 25968, 23383, 20013, 26377, 25968, 23383, 36229, 36807, 26412, 31243, 24207, 22788, 29702, 33539, 22260) & "，" & UText(26412, 31243, 24207, 21482, 22788, 29702) & "AQL：1.0，1.5，2.5，4.0，6.5" & UText(30340, 25968, 20540) & "！！！，" & UText(35831, 20462, 25913, 25110, 25163, 24037, 22788, 29702) & " ", vbCritical: GoTo 100
30     If MV(Hgrp, Agrp) > 0 Then
         If MV(Hgrp, Agrp) >= SQty Then SmpS(k) = SQty              '" & UText(31639) & "sample size" & UText(20540) & "
         If MV(Hgrp, Agrp) < SQty Then SmpS(k) = MV(Hgrp, Agrp)
       End If
       If MV(Hgrp, Agrp) = 0 Then SmpS(k) = CD(Hgrp)
    
       AAc(k) = AC(Hgrp, Agrp)                      '" & UText(31639) & "AQL" & UText(20540) & "
       
40    If SmpS(k) > Ass Then Ass = SmpS(k)            '" & UText(21462) & "Sample Sise " & UText(21462, 22823, 20540) & "
      
      If AQL(0) = 0 Then AAc(0) = 0           'AQL=0 " & UText(26368) & "AC" & UText(35774, 20026) & "0
        
     Next k
     
     If Acr = 0 Then SmpS(0) = Ass

       '" & UText(21462, 22635, 20837, 30340) & "Sample size" & UText(36827, 34892, 26816, 26597) & "
        If Ass <> scol(8) And SamP Like "*Fixed*" = False Then MsgBox "The total sample size input in 1. Quantity table was " & scol(8) & " However in according to ISO2859 it should be " & Ass & _
       " Please check and revise it." & CHR(13) & _
       UText(34920, 26684) & "1. Quantity" & UText(32047, 32479, 35745, 20986, 30340) & "Sample Size" & UText(26159) & "：" & scol(8) & " " & UText(20294, 26159, 25353) & "ISO2859" & UText(26631, 20934, 26597, 20986, 30340) & "Sample Size" & UText(24212, 35813, 26159) & ":" & Ass & _
       " " & UText(35831, 36827, 34892, 20462, 25913) & "，" & UText(22914, 26524, 25353) & "ISO2859" & UText(26631, 20934) & "，" & UText(35831, 20462, 25913, 34920, 26684) & "1. Quantity" & UText(20013, 30340, 25968, 20540) & "，" & UText(21518, 20877, 27425, 36816, 34892, 31243, 24207) & "，" & UText(22914, 26524, 19981, 26159, 25353) & "ISO2859，" & UText(35831, 25163, 24037, 20462, 25913, 34920, 21333) & "！！！"
       
       
       Set rg = .Cell(9, 4).Range
       ss = Val(rg.words(1))
       If Ass <> ss And SamP Like "*Fixed*" = False Then
           cnt = MsgBox("Sample Size in information table on first page was" & ss & " , however in according to ISO2859, it should be " & Ass & " , click YES to replace sample size data in Information table, click NO to deal with it yourself." & CHR(13) & _
           "Information " & UText(34920, 26684, 22635, 20837, 30340) & "Sample Size" & UText(26159) & ss & " " & UText(20294, 25353) & "ISO2859" & UText(26597, 21040, 30340) & "Sample Size" & UText(26159) & Ass & UText(26159, 21542, 26367, 25442) & "Information " & UText(34920, 26684, 25968, 25454) & "！", vbYesNo)
           If cnt = vbYes Then rg.words(1) = Replace(ss, ss, Ass & " ")
           If cnt = vbNo Then MsgBox "Your choise was NO, please revise the sample size data yourself" & CHR(13) & _
           UText(20320, 36873, 25321, 19981, 26367, 25442, 39318, 39029, 34920, 26684) & "，" & UText(35828, 26126, 19981, 31526, 21512) & "ISO2859" & UText(25277, 26679, 35201, 27714) & "，" & UText(35831, 20462, 25913, 34920, 26684) & "1. Quantity" & UText(30340) & "Sample Size" & UText(20540) & "！！！"
       End If
     '--------- " & UText(22266, 23450, 25277, 26679, 26102) & "，" & UText(25226) & "sample SIZE " & UText(20174, 26597, 34920, 24378, 21046, 25913, 20026, 39318, 39029, 36755, 20837, 20540) & "
60   If Styp = 3 Then
       SmpS(0) = Val(.Cell(9, 4).Range.words(1))
       SmpS(1) = SmpS(0)
       SmpS(2) = SmpS(0)
     End If
     If Styp = 1 And SamP Like "*Fixed*" = True Then
       SmpS(0) = Val(.Cell(9, 4).Range.words(1))
       SmpS(1) = SmpS(0)
       SmpS(2) = SmpS(0)
     End If

     '---------
  End With
    

 
100    '" & UText(21462, 24471, 23454, 38469) & " defect allowed
    ADef(1) = AAc(0)      'cr
    ADef(2) = AAc(1)      'Ma
    ADef(3) = AAc(2)     'Mi
    
110 '" & UText(26816, 26597) & "workmanship " & UText(34920) & "
Dim myrow, M, sCr, sMa, sMi As Integer
Dim dCr(100), dMa(100), dMi(100), N As Integer
Dim tem1, temw, mbx, Fw As String
Dim rg3 As Range

   tb = ThisDocument.Tables.count - 1
   For X = 1 To tb
   With ThisDocument.Range.Tables(X)
        If Trim(.Cell(1, 1).Range.Text) Like TBName(2) = True Then tbn = X
     End With
  Next X
  
   With ThisDocument.Range.Tables(tbn)
      

    myrow = .Rows.count
  
    myrow = myrow - 2  '" & UText(33719, 24471) & "total foud " & UText(34920, 26684, 25152, 22312, 34892, 25968) & "
     

     For k = myrow - 1 To 3 Step -1 '" & UText(21024, 38500, 20840, 31354, 34892) & "
        tem = .Cell(k, 1).Range.Text
        If LCase(Left(Trim(tem), 4)) <> "item" Then
          If Left(Trim(tem), 1) = CHR(13) And _
          Val(Trim(.Cell(k, 2).Range.Text)) = 0 And _
          Val(Trim(.Cell(k, 3).Range.Text)) = 0 And _
          Val(Trim(.Cell(k, 4).Range.Text)) = 0 Then .Rows(k).Delete
        End If
     Next k
     
     '" & UText(21024, 38500, 19979, 34920, 27809, 20869, 23481, 30340, 22810, 20313) & "ITEM" & UText(34892) & "
     myrow = .Rows.count
  
      myrow = myrow - 2

     Do While LCase(Left(Trim(.Cell(myrow - 1, 1).Range.Text), 4)) = "item"

    .Rows(myrow - 1).Delete
      myrow = myrow - 1
     Loop
 
     
     msg = ""
     mbx = ""
     
     For k = 3 To (myrow - 1) '" & UText(20174, 31532, 19977, 34892, 24320, 22987) & "，" & UText(21028, 26029, 26159, 21542) & "item" & UText(34892) & "," & UText(33719, 24471, 27599, 34892, 21333, 20803, 26684, 30340) & "defect " & UText(25968) & "
       
     
      tem = .Cell(k, 1).Range.Text
      temw = Mid(tem, 1, Len(tem) - 1)
    
          If LCase(Left(Trim(tem), 4)) <> "item" Then
          
                    '.Cell(k, 1).Range.Text = StrConv(Mid(tem, 1, Len(tem) - 2), vbProperCase)
            
             Fw = UCase(Left(temw, 1)) & LCase(Mid(temw, 2, Len(temw) - 1))
   
           .Cell(k, 1).Range.Text = Mid(Fw, 1, Len(Fw) - 1)
            dCr(k) = Val(.Cell(k, 2).Range.Text)
            dMa(k) = Val(.Cell(k, 3).Range.Text)
            dMi(k) = Val(.Cell(k, 4).Range.Text)
           
            If dCr(k) < 0 Or dMa(k) < 0 Or dMi(k) < 0 Then msg = msg + " row " & k & ", the defect of " & temw & ", the defect number was minus!" & CHR(13) & _
            UText(31532) & k & UText(34892, 32570, 38519) & temw & UText(32570, 38519, 25968, 20026, 36127, 25968) & "!" & CHR(10)
            
            If Left(Trim(tem), 1) <> CHR(13) And UCase((Mid(Trim(tem), 1, 9))) <> "NO DEFECT" Then
              
                  If dCr(k) = 0 And dMa(k) = 0 And dMi(k) = 0 Then msg = msg + "the row " & k & ", the defect of " & temw & ", the defect number was not input" & CHR(10) & _
                  UText(31532) & k & UText(34892, 32570, 38519) & temw & UText(32570, 38519, 25968, 37327, 27809, 26377, 21028, 23450) & CHR(10)
                      
                  If (dCr(k) > 0 And dMa(k) > 0) Or (dCr(k) > 0 And dMi(k) > 0) Then msg = msg + " row " & k & ", the defect of" & temw & " input with  Critical and other defect." & CHR(10) & _
                  UText(31532) & k & UText(34892) & temw & UText(26377) & "Critical" & UText(32570, 38519) & "，" & UText(21448, 26377, 20854, 20182, 32570, 38519) & "，" & UText(37325, 22797, 21028, 23450) & "!" & CHR(10)
             End If
             
                 
              If UCase((Mid(Trim(tem), 1, 9))) = "NO DEFECT" Then
                     If dCr(k) > 0 Or dMa(k) > 0 Or dMi(k) > 0 Then msg = msg + "row " & k & " , with statement of No Defect and defect number." & CHR(10) & _
                     UText(31532) & k & UText(34892) & "No Defect," & UText(20294, 21448, 26377, 32570, 38519, 25968) & "!" & CHR(10)
              End If
                  
              If Left(Trim(tem), 1) = CHR(13) Then
              
                     If dCr(k) > 0 Or dMa(k) > 0 Or dMi(k) > 0 Then msg = msg + "row " & k & ", without defect statement!" & CHR(10) & _
                     UText(31532) & k & UText(34892, 32570, 38519, 25551, 36848, 26410, 20889) & "," & UText(20294, 21448, 26377, 32570, 38519, 25968) & "!" & CHR(10)
              End If
              
                    
         Else
            
           dCr(k) = 0
           dMa(k) = 0
           dMi(k) = 0
          
         End If
      
     Next k
          
           If msg <> "" Then MsgBox "2. Workmanship Table:" & CHR(10) & msg & "Please revise them and run program again" & CHR(10) & _
           "2. Workmanship" & UText(34920) & CHR(10) & msg & UText(35831, 23545, 20197, 19978, 36827, 34892, 20462, 25913, 21518, 20877, 27425, 36816, 34892) & "！！！", vbCritical, UText(35686, 21578): Exit Sub
           
       sCr = 0 '" & UText(28165, 38646) & "
       sMa = 0
       sMi = 0
     For M = 3 To (myrow - 1)
       sCr = sCr + dCr(M)
       sMa = sMa + dMa(M)
       sMi = sMi + dMi(M)
     Next M
     
    If Val(.Cell(myrow, 2).Range.Text) <> sCr Then msg = msg + TBName(2) & " table: sum of Critical defect was" & sCr & ", However the input date was " & Val(.Cell(myrow, 2).Range.Text) & ". They are not equal." & CHR(13) & _
    TBName(2) & UText(34920, 26684) & "Critical " & UText(32570, 38519, 30340, 32047, 21152, 20540, 35745, 31639, 32467, 26524, 20026) & "：" & sCr & UText(30446, 21069, 22635, 20837, 30340, 20540, 21364, 26159) & "：" & Val(.Cell(myrow, 2).Range.Text) & UText(20108, 32773, 19981, 30456, 31561) & CHR(13)
    
    If Val(.Cell(myrow, 3).Range.Text) <> sMa Then msg = msg + TBName(2) & " table: sum of Major defect was" & sMa & ", the input date was " & Val(.Cell(myrow, 3).Range.Text) & "They are not equal." & CHR(13) & _
    TBName(2) & UText(34920, 26684) & "Major " & UText(32570, 38519, 30340, 32047, 21152, 20540, 35745, 31639, 32467, 26524, 20026) & "：" & sMa & UText(30446, 21069, 22635, 20837, 30340, 20540, 21364, 26159) & "：" & Val(.Cell(myrow, 3).Range.Text) & UText(20108, 32773, 19981, 30456, 31561) & CHR(13)
    If Val(.Cell(myrow, 4).Range.Text) <> sMi Then msg = msg + TBName(2) & "table: sum of Minor defect was " & sMi & ", the input data was " & Val(.Cell(myrow, 4).Range.Text) & "They are not equal." & CHR(13) & _
    TBName(2) & UText(34920, 26684) & "Minor " & UText(32570, 38519, 30340, 32047, 21152, 20540, 35745, 31639, 32467, 26524, 20026) & "：" & sMi & UText(30446, 21069, 22635, 20837, 30340, 20540, 21364, 26159) & "：" & Val(.Cell(myrow, 4).Range.Text) & UText(20108, 32773, 19981, 30456, 31561) & CHR(13)
        
        Coln = .Columns.count
        Rown = .Rows.count
     If Styp <> 3 Then
        For i = 2 To Coln
         If Val(.Cell(Rown - 1, i).Range.Text) <> ADef(i - 1) Then msg = msg + TBName(2) & "table: row " & i & " the work out AQL was " & ADef(i - 1) & ", however the input data was " & Val(.Cell(Rown - 1, i).Range.Text) & ". Please check" & CHR(13) & _
         TBName(2) & UText(34920, 26684) & "，" & UText(31532) & i & UText(21015) & "," & UText(35745, 31639, 20986, 30340) & "AQL" & UText(30340) & "Ac" & UText(26159) & "：" & ADef(i - 1) & UText(30446, 21069, 22635, 20837, 30340, 25968, 25454, 21364, 26159) & "：" & Val(.Cell(Rown - 1, i).Range.Text) & UText(35831, 26816, 26597, 20462, 25913) & "！！！" & CHR(13)
         If Val(.Cell(Rown, i).Range.Text) <> SmpS(i - 2) Then msg = msg + TBName(2) & " Table: column " & i & ", the work out Ac was " & SmpS(i - 2) & ", However the input data was " & Val(.Cell(Rown, i).Range.Text) & ",Please check" & CHR(13) & _
         TBName(2) & UText(34920, 26684) & "，" & UText(31532) & i & UText(21015) & "," & UText(35745, 31639, 20986, 30340) & "AQL" & UText(30340) & "Ac" & UText(26159) & "：" & SmpS(i - 2) & UText(30446, 21069, 22635, 20837, 30340, 25968, 25454, 21364, 26159) & "：" & Val(.Cell(Rown, i).Range.Text) & UText(35831, 26816, 26597, 20462, 25913) & "！！！" & CHR(13)
        Next i
     End If
      myrow = Rown - 2
      If Val(.Cell(myrow, 2).Range.Text) > Val(.Cell(myrow + 1, 2).Range.Text) Then
            If .Cell(myrow, 2).Range.Font.Color <> wdColorRed Then msg = msg + TBName(2) & " table, row " & myrow & "Critical defect exceed allowance, however the total number was not marked with redcolor." & CHR(13) & _
            TBName(2) & " " & UText(34920, 31532) & myrow & UText(34892) & "，Critical defect" & UText(36229, 26631) & "，" & UText(20294, 20854, 32047, 21152, 25968, 27809, 26377, 26631, 32418) & CHR(13)
      Else
            If .Cell(myrow, 2).Range.Font.Color = wdColorRed Then msg = msg + TBName(2) & " table, row " & myrow & "Critical defect not exceed allowance, however the total number was marked with redcolor." & CHR(13) & _
            TBName(2) & " " & UText(34920, 31532) & myrow & UText(34892) & "，Critical defect" & UText(26410, 36229, 26631) & "，" & UText(20294, 20854, 32047, 21152, 25968, 21364, 26631, 32418) & CHR(13)
      End If
         
   
     If Val(.Cell(myrow, 3).Range.Text) > Val(.Cell(myrow + 1, 3).Range.Text) Then
         If .Cell(myrow, 3).Range.Font.Color <> wdColorRed Then msg = msg + TBName(2) & " table, row " & myrow & "Major defect exceed allowance, however the total number was not marked with redcolor." & CHR(13) & _
            TBName(2) & " " & UText(34920, 31532) & myrow & UText(34892) & "，Major defect" & UText(36229, 26631) & "，" & UText(20294, 20854, 32047, 21152, 25968, 27809, 26377, 26631, 32418) & CHR(13)
     Else
         If .Cell(myrow, 3).Range.Font.Color = wdColorRed Then msg = msg + TBName(2) & " table, row " & myrow & "Major defect not exceed allowance, however the total number was  marked with redcolor." & CHR(13) & _
             TBName(2) & " " & UText(34920, 31532) & myrow & UText(34892) & "，Major defect" & UText(26410, 36229, 26631) & "，" & UText(20294, 20854, 32047, 21152, 25968, 21364, 26631, 32418) & CHR(13)
     End If
         
  
      If Val(.Cell(myrow, 4).Range.Text) > Val(.Cell(myrow + 1, 4).Range.Text) Then
           If .Cell(myrow, 4).Range.Font.Color <> wdColorRed Then msg = msg + TBName(2) & " table, row " & myrow & "Minor defect exceed allowance, however the total number was not marked with redcolor." & CHR(13) & _
             TBName(2) & " " & UText(34920, 31532) & myrow & UText(34892) & "，Minor defect" & UText(36229, 26631) & "，" & UText(20294, 20854, 32047, 21152, 25968, 27809, 26377, 26631, 32418) & CHR(13)
      Else
          If .Cell(myrow, 4).Range.Font.Color = wdColorRed Then msg = msg + TBName(2) & " table, row " & myrow & "Minor defect not exceed allowance, however the total number was marked with redcolor." & CHR(13) & _
             TBName(2) & " " & UText(34920, 31532) & myrow & UText(34892) & "，Minor defect" & UText(26410, 36229, 26631) & "，" & UText(20294, 20854, 32047, 21152, 25968, 21364, 26631, 32418) & CHR(13)
      End If
      
     
     
    If sCr > Val(.Cell(Rown - 1, 2).Range.Text) Or sMa > Val(.Cell(Rown - 1, 3).Range.Text) Or sMi > _
    Val(.Cell(Rown - 1, 4).Range.Text) Then
       If Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1) <> "Not Conform" Then msg = msg + "Defect number in Workmanship table exceed the allowance, however the result was not 'Not Conform'，Please choose Not Conform!!!" & CHR(13) & _
       "Workmanship" & UText(34920, 20013, 30340, 32570, 38519, 25968, 22823, 20110, 20801, 35768, 20540) & "，" & UText(29616, 22312, 30340, 32467, 35770, 19981, 26159) & "Not Conform，" & UText(35831, 36873, 25321) & "Not Conform!!!" & CHR(13)
    End If
    
    If sCr <= Val(.Cell(Rown - 1, 2).Range.Text) And sMa <= Val(.Cell(Rown - 1, 3).Range.Text) And sMi <= _
    Val(.Cell(Rown - 1, 4).Range.Text) Then
      If SAP(2) = 1 And Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1) <> "Conform" Then msg = msg + "Defect number in Workmanship table within allowance and without SAP，However the result was not 'Conform'！！！" + CHR(13) & _
      "Workmanship" & UText(34920, 20013, 30340, 32570, 38519, 25968, 22312, 20801, 35768, 33539, 22260, 20869, 19988, 27809, 26377) & "SAP，" & UText(29616, 22312, 30340, 32467, 35770, 19981, 26159) & "Conform！！！" + CHR(13)
      If SAP(2) > 1 And Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1) = "Conform" Then msg = msg + "Defect number within allowance but there are SAP for workmanship, however the result was not Not Conform！！！" & CHR(13) & _
       "Workmanship" & UText(34920, 20013, 30340, 32570, 38519, 25968, 22312, 20801, 35768, 33539, 22260, 20869, 20294, 26377) & "SAP" & UText(39033) & "，" & UText(29616, 22312, 30340, 32467, 35770, 21364, 26159) & "Not Conform！！！" + CHR(13)
    
    End If
    If msg <> "" Then MsgBox msg & "Please check and revise！！！" & CHR(13) & UText(35831, 36827, 34892, 26816, 26597, 20462, 25913) & "！！！", vbCritical: Exit Sub
    
   
    
 End With
    
   
   Call Fullcheckitem
    
   Call checkscdform

  ' tb = ThisDocument.Tables.count-2
  ' For x = 1 To tb
  '   With Thisdocument.Range.Tables(x)
   '     If Trim(.Cell(1, 1).Range.Text) Like tbname(3) = True Then tbn = x
  '   End With
'  Next x
  
  ' With Thisdocument.Range.Tables(tbn)

    ' If Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1) = "Conform" And sap(3) > 1 Then MsgBox tbname(3) & "Table: the result was Conform but with SAP，Please note" & Chr(13) & _
     tbname(3) & UText(34920, 26684, 32467, 26524, 20026) & "Conform " & UText(20294, 26377) & "SAP，" & UText(35831, 27880, 24847) & Chr(13) & Chr(13)
   '  If Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1) <> "Conform" And sap(3) = 1 Then MsgBox tbname(3) & "Table: the result was not Conform，but without SAP，Please revise it." & Chr(13) & _
     tbname(3) & UText(34920, 26684, 32467, 35770, 19981, 26159) & "conform，" & UText(20294, 21364, 27809, 26377, 20889) & "SAP，" & UText(35831, 36827, 34892, 20462, 25913), vbCritical: Exit Sub
  ' End With
    
 
    Call checktestform
  ' tb = ThisDocument.Tables.count-2
   'For x = 1 To tb
  ' With Thisdocument.Range.Tables(x)
  '      If Trim(.Cell(1, 1).Range.Text) Like tbname(4) = True Then tbn = x
  '  End With
 '  Next x
  
  ' With Thisdocument.Range.Tables(tbn)
   '  If Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1) = "Conform" And sap(4) > 1 Then MsgBox tbname(4) & " Table: the result was Conform, nut with SAP，Please note." & Chr(13) & _
     tbname(4) & UText(34920, 26684, 32467, 26524, 20026) & "Conform " & UText(20294, 26377) & "SAP，" & UText(35831, 27880, 24847) & Chr(13)
  '   If Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1) <> "Conform" And sap(4) = 1 Then MsgBox tbname(4) & "Table: the result was not Conform，but without SAP，Please revise it." & Chr(13) & _
     tbname(4) & UText(34920, 26684, 32467, 35770, 19981, 26159) & "conform，" & UText(20294, 21364, 27809, 26377, 20889) & "SAP，" & UText(35831, 36827, 34892, 20462, 25913), vbCritical: Exit Sub
 '  End With
  
  
     Call checkNote
  
   
    


    
   For X = 1 To tb
   With ThisDocument.Range.Tables(X)
        If Trim(.Cell(1, 1).Range.Text) Like TBName(1) = True Then
        Res = .Cell(1, 2).Range.Text
        dres(1) = Mid(Res, 1, Len(Res) - 1)
            If dres(1) Like ch = True Or dres(1) Like che = True Or dres(1) = "" Then msg = msg + "1.Quantity Table without result" & CHR(13) & "1.Quantity" & UText(34920, 26684, 32467, 35770, 26410, 36873) & CHR(13)
        End If
        If Trim(.Cell(1, 1).Range.Text) Like TBName(2) = True Then
        Res = .Cell(1, 2).Range.Text
        dres(2) = Mid(Res, 1, Len(Res) - 1)
           If dres(2) Like ch = True Or dres(2) Like che = True Or dres(2) = "" Then msg = msg + "2.Workmanship Table without result" & CHR(13) & "2.Workmanship" & UText(34920, 26684, 32467, 35770, 26410, 36873) & CHR(13)
        End If
        If Trim(.Cell(1, 1).Range.Text) Like TBName(3) = True Then
        Res = .Cell(1, 2).Range.Text
        dres(3) = Mid(Res, 1, Len(Res) - 1)
            If dres(3) Like ch = True Or dres(3) Like che = True Or dres(3) = "" Then msg = msg + "3.Style, Color & Documentation table without result" & CHR(13) & "3.Style, Color & Documentation" & UText(34920, 26684, 32467, 35770, 26410, 36873) & CHR(13)
        End If
        If Trim(.Cell(1, 1).Range.Text) Like TBName(4) = True Then
        Res = .Cell(1, 2).Range.Text
        dres(4) = Mid(Res, 1, Len(Res) - 1)
   
            If dres(4) Like ch = True Or dres(4) Like che = True Or dres(4) = "" Then msg = msg + "4.Field Test & Measurement table without result" & CHR(13) & "4.Field Test & Measurement" & UText(34920, 26684, 32467, 35770, 26410, 36873) & CHR(13)
        End If
        If Trim(.Cell(1, 1).Range.Text) Like TBName(5) = True Then
        Res = .Cell(1, 2).Range.Text
        dres(5) = Mid(Res, 1, Len(Res) - 1)
             If dres(5) Like ch = True Or dres(5) Like che = True Or dres(5) = "" Then msg = msg + "5.Shipping Mark & Packaging table without result" & CHR(13) & "5.Shipping Mark & Packaging" & UText(34920, 26684, 32467, 35770, 26410, 36873) & CHR(13)
        End If
     End With
    Next X
    
    If msg <> "" Then
      With ThisDocument.Tables(3)
         For i = 3 To 7   '" & UText(21333, 20803, 26684, 28165, 38646) & "
           For j = 2 To 4
         .Cell(i, j).Range.Text = ""
           Next j
         Next i
     .Cell(8, 2).Range.Text = ""
      End With
      MsgBox msg & "Please choose the result and run program again." & CHR(13) & UText(35831, 36873, 25321, 21518, 20877, 36816, 34892, 31243, 24207) & "！！！": Exit Sub
       
   End If
 
       
   
  
     
             '" & UText(33719, 21462) & "SAP" & UText(34920, 30340, 32467, 35770) & "
             
        Dim sapQty, swms, sscd, stst, ssmp, kwd, rwd As String
        Dim lwd, lrd As Integer
             
        
            With ThisDocument.Bookmarks("sapcontent").Range
              tNo = .Tables.count '" & UText(35745, 31639, 34920, 26684, 24635, 25968) & "
              For T = 1 To tNo
                With ThisDocument.Bookmarks("sapcontent").Range.Tables(T)
      
                 Rc = .Rows.count
                 cc = .Columns.count
                 
                 If cc > 2 Then
                   kwd = .Cell(1, 2).Range.Text
                   rwd = .Cell(1, 3).Range.Text
                   lwd = Len(.Cell(1, 2).Range.Text)
                   lrd = Len(.Cell(1, 3).Range.Text)
                   kwd = Mid(kwd, 1, (lwd - 2))
                   rwd = Mid(rwd, 1, (lrd - 2))
                  
                   
                   If kwd Like "*Quantity*" = True Then sapQty = rwd
                     
                   If kwd Like "*Workmanship*" = True Then swms = rwd
                   If kwd Like "*Style, Color & Documentation*" = True Then sscd = rwd
                   If kwd Like "*Field Test & Measurement*" = True Then stst = rwd
                   If kwd Like "*Shipping Mark & Packaging*" = True Then ssmp = rwd
                   
                   '" & UText(26816, 26597) & "SAP" & UText(26159, 21542, 36816, 34892, 36807, 26816, 26597) & "
                   If Rc > 1 Then
                    
                      If Mid(.Cell(2, 2).Range.Text, 1, Len(.Cell(2, 2).Range.Text) - 2) Like ch = True Then MsgBox "Please run the micro program for SAP first." & CHR(13) & _
                      UText(35831, 22312) & "SAP" & UText(34920, 26684) & "，" & UText(20808, 36816, 34892) & "SAP" & UText(26816, 26597): Exit Sub
                   End If
                   
                 End If
                 
                End With
                
             Next T
             
            End With
            
        '" & UText(21028, 23450, 26377, 21542, 26377) & "SAP" & UText(30340, 34987, 21028, 20026, 21512, 26684) & "
            If dres(1) = DCF And sapQty = "" Then alt = alt + "There are SAP for Quantity while the result of Quantity was Conform" & CHR(13) & _
            "SAP 1 Quantity " & UText(20013, 26377, 22791, 27880, 20107, 39033) & "，Quantity" & UText(30340, 32467, 35770, 21364, 20026) & "Conform" & CHR(13)
            
            If dres(2) = DCF And swms = "" Then alt = alt + "There are SAP for Workmanship while the result of Workmanship was Conform" & CHR(13) & _
            "SAP 2 Workmanship" & UText(20013, 26377, 22791, 27880, 20107, 39033) & "，Workmanship" & UText(30340, 32467, 35770, 21364, 20026) & "Conform" & CHR(13)
            
            If dres(3) = DCF And sscd = "" Then alt = alt + "There are SAP for  Style, Color & Documentation while the result ofStyle, Color & Documentation was Conform" & CHR(13) & _
            "SAP 3 Style, Color & Documentation" & UText(20013, 26377, 22791, 27880, 20107, 39033) & "，Style, Color & Documentation" & UText(30340, 32467, 35770, 21364, 20026) & "Conform" & CHR(13)
            
            If dres(4) = DCF And stst = "" Then alt = alt + "There are SAP for  Field Test & Measurement while the result of Field Test & Measurement was Conform" & CHR(13) & _
            "SAP 4 Field Test & Measurement" & UText(20013, 26377, 22791, 27880, 20107, 39033) & "，Field Test & Measurement" & UText(30340, 32467, 35770, 21364, 20026) & "Conform" & CHR(13)
            
            If dres(5) = DCF And ssmp = "" Then alt = alt + "There are SAP for Shipping Mark & Packaging while the result of shipping Mark & Packaging was Conform" & CHR(13) & _
            "SAP 5Shipping Mark & Packaging" & UText(20013, 26377, 22791, 27880, 20107, 39033) & "Shipping Mark & Packaging" & UText(30340, 32467, 35770, 21364, 20026) & "Conform" & CHR(13)
            
            If alt <> "" Then MsgBox alt & "Please check, if you revise any please run program again." & CHR(13) & UText(35831, 26816, 26597, 30830, 35748) & "，" & UText(33509, 26377, 20462, 25913, 35831, 25913, 21518, 20877, 27425, 36816, 34892)
         
        With ThisDocument.Tables(3)
         For i = 3 To 7   '" & UText(21333, 20803, 26684, 28165, 38646) & "
           For j = 2 To 4
         .Cell(i, j).Range.Text = ""
           Next j
         Next i
         
    '" & UText(32467, 35770, 21333, 20803, 26684, 36171, 20540) & "
          
          For i = 3 To 7
              If dres(i - 2) = NA Then .Cell(i, 2).Range.Text = "-": .Cell(i, 3).Range.Text = "-": .Cell(i, 4).Range.Text = "-"
              If dres(i - 2) = DCF Then .Cell(i, 2).Range.Text = "X"
              If dres(i - 2) = DNC Then .Cell(i, 3).Range.Text = "X"
              If dres(i - 2) = DPD Then .Cell(i, 4).Range.Text = "X"
          Next i
         
           
          '" & UText(32467, 35770, 21028, 23450) & "
   
        CF = "CONFORM"
        NC = "NOT CONFORM"
        PD = "PENDING"
        CFdis = "to client's requirement"
        NCdis = "to client's requirement"
        PDdis = "for client's evaluation"
    
         If dres(1) = DNC Or dres(2) = DNC Or dres(3) = DNC Or dres(4) = DNC Or dres(5) = DNC Then
       .Cell(8, 2).Range.Text = NC & " " & NCdis
          With .Cell(8, 2).Range.words(1).Font
         .Color = 192
         .Size = 12
         .Bold = True
          End With
         .Cell(8, 2).Range.Text = NC & " " & NCdis
          With .Cell(8, 2).Range.words(2).Font
         .Color = 192
         .Size = 12
         .Bold = True
          End With
          With .Cell(8, 2).Range.words(3).Font
          .Color = wdColorAutomatic
          .Size = 10
          .Bold = False
          End With
         With .Cell(8, 2).Range.words(4).Font
        .Color = wdColorAutomatic
        .Size = 10
        .Bold = False
         End With  '
         With .Cell(8, 2).Range.words(5).Font
          .Color = wdColorAutomatic
          .Size = 10
          .Bold = False
         End With: GoTo 200
      End If
        
      If dres(1) = DPD Or dres(2) = DPD Or dres(3) = DPD Or dres(4) = DPD Or dres(5) = DPD Then
      .Cell(8, 2).Range.Text = PD & " " & PDdis
        With .Cell(8, 2).Range.words(1).Font
         .Color = 192
         .Size = 12
         .Bold = True
          End With
          With .Cell(8, 2).Range.words(2).Font
          .Color = wdColorAutomatic
          .Size = 10
          .Bold = False
          End With
          With .Cell(8, 2).Range.words(3).Font
         .Color = wdColorAutomatic
         .Size = 10
         .Bold = False
          End With  '
          With .Cell(8, 2).Range.words(4).Font
          .Color = wdColorAutomatic
          .Size = 10
          .Bold = False
          End With
        End If
        
        If (dres(1) = DCF Or dres(1) = NA) And _
           (dres(2) = DCF Or dres(2) = NA) And _
           (dres(3) = DCF Or dres(3) = NA) And _
           (dres(4) = DCF Or dres(4) = NA) And _
           (dres(5) = DCF Or dres(5) = NA) Then
        
       .Cell(8, 2).Range.Text = CF & " " & CFdis
          With .Cell(8, 2).Range.words(1).Font
         .Color = 192
         .Size = 12
         .Bold = True
          End With
         With .Cell(8, 2).Range.words(2).Font
          .Color = wdColorAutomatic
          .Size = 10
          .Bold = False
          End With
          With .Cell(8, 2).Range.words(3).Font
          .Color = wdColorAutomatic
          .Size = 10
          .Bold = False
          End With  '
          With .Cell(8, 2).Range.words(4).Font
          .Color = wdColorAutomatic
          .Size = 10
          .Bold = False
          End With

        End If
        
    End With
    
200   MsgBox "Complete calculation." & CHR(13) & "Note: if you revise any data, please run the program again." & CHR(13) & _
UText(32467, 35770, 35745, 31639, 23436, 25104) & "！" & CHR(13) & UText(27880, 24847) & "，" & UText(33509, 26377, 25552, 31034, 35201, 27714, 26816, 26597, 20462, 25913, 30340) & "，" & UText(35831, 20462, 25913, 21518, 20877, 27425, 36816, 34892) & "！"
    
End Sub
Sub sumdefect()

'
' sumdefect " & UText(23439) & ",Revby wyn/20241216
'
'" & UText(32047, 21152, 32570, 38519) & "
   
Dim myrow, i, j, k, M, sCr, sMa, sMi As Integer
Dim dCr(200), dMa(200), dMi(200), N, Dk, Sk As Integer
Dim tem, tem1, temw, msg, mbx, Fw, Res As String
Dim rg3 As Range
Dim InL(200) As Integer
Dim Coln As Integer
Dim TBName As String
Dim dictItem As Object
Dim ComItms  As Integer

    
    tbn = getTableNumber(1)
    With WD.Tables(tbn)
       qtyRow = .Rows.count
       Itms = qtyRow - 5
    
       Dim ComItx() As String
       ReDim ComItx(1 To Itms)
       Dim ComSm() As Integer
       ReDim ComSm(1 To Itms)
       Dim Sm() As Integer
       ReDim Sm(1 To Itms)
       N = 0
       Set dictItem = CreateObject("Scripting.Dictionary")
       For i = 5 To qtyRow - 1
         Kw = Mid(.Cell(i, 2).Range.Text, 1, Len(.Cell(i, 2).Range.Text) - 2)
         Sm(i - 4) = Val(.Cell(i, 8).Range.Text)
        
          If Not dictItem.Exists(Kw) Then
              N = N + 1
              dictItem.add Kw, Sm(i - 4)
              ComSm(N) = Sm(i - 4)
        Else
            dictItem(Kw) = dictItem(Kw) + Sm(i - 5)
             ComSm(N) = dictItem(Kw)
        End If
      Next i
        ComItms = dictItem.count
     End With
 
   TBName = "defects"
   tb = ThisDocument.Tables.count - 1
   For X = 1 To tb
   With ThisDocument.Range.Tables(X)
        If Left(LCase(Trim(.Cell(1, 1).Range.Text)), 7) = TBName Then
            tbn = X
            GoTo 10
        End If
     End With
  Next X
  If Itms = 0 Then MsgBox "The program has not got the items quantity,please run the Quantity total program at the Quantity table before run this program" & CHR(10) & _
  UText(26412, 31243, 24207, 36824, 26410, 33719, 24471, 27454, 25968) & "，" & UText(35831, 20808, 36816, 34892) & "Quantity" & UText(34920, 26684, 20013, 30340) & "Total" & UText(31243, 24207, 21518, 20877, 36816, 34892, 26412, 31243, 24207), vbCritical: Exit Sub
 
10   With ThisDocument.Range.Tables(tbn)
      

    myrow = .Rows.count
    Coln = .Columns.count
    myrow = myrow - 2  '" & UText(33719, 24471) & "total foud " & UText(19978, 19968, 34892) & "
    N = 0
    For k = 2 To myrow
        If .Rows(k).Cells.count = 1 And LCase(.Cell(k, 1).Range.Text) Like "*item*" = True Then
          N = N + 1
        End If
    Next k

    If N < ComItms Then
      If ComItms = 1 Then
        .Rows(2).Select
         Selection.InsertRowsAbove
         .Rows(2).Range.Select
         Selection.Range.Cells.Merge
         .Cell(2, 1).Range.Text = "Item No."
      Else
         MsgBox "The item quantity for this inspection is " & ComItms & " While ther are only " & N & "rows input with Item Info" & _
         "You may have delete some Item rows, Please add manually then run program again" & CHR(10) & _
         UText(26412, 27425, 39564, 36135, 27454, 25968, 20026) & ComItms & UText(27454) & "，" & UText(20294, 20165, 26377) & N & UText(34892, 36755, 20837) & "Item" & UText(21495) & "，" & UText(20320, 21487, 33021, 21024, 38500, 20102, 26576, 20123) & "ITem" & UText(34892) & "，" & UText(35831, 25163, 24037, 34917, 20805, 21518, 20877, 36816, 34892, 31243, 24207) & "！", vbCritical
          Exit Sub
      End If
    
    End If
    myrow = .Rows.count
    myrow = myrow - 2
        
     For k = myrow To 2 Step -1  '" & UText(21024, 38500, 20840, 31354, 34892) & "
        tem = .Cell(k, 1).Range.Text
        If LCase(Left(Trim(tem), 4)) <> "item" Then
          If Left(Trim(.Cell(k, Coln - 3).Range.Text), 1) = CHR(13) And _
             Val(Trim(.Cell(k, Coln - 2).Range.Text)) = 0 And _
             Val(Trim(.Cell(k, Coln - 1).Range.Text)) = 0 And _
             Val(Trim(.Cell(k, Coln).Range.Text)) = 0 Then
              If LCase(.Cell(k - 1, 1).Range.Text) Like "*item*" = True Then

                  If LCase(.Cell(k + 1, 1).Range.Text) Like "*item*" = True Then
                    .Cell(k, Coln - 3).Range.Text = "No defect found"
                  ElseIf .Cell(k + 1, Coln - 3).Range.Fields.count > 0 Then
                     .Cell(k, Coln - 3).Range.Text = "No defect found"
                  Else
                    .Rows(k).Delete
                  End If
               Else
                   .Rows(k).Delete
               End If
                  
              
          End If
        End If
     Next k
    '=======
    myrow = .Rows.count
    myrow = myrow - 2
    N = 0
    For k = 2 To myrow
      If LCase(Left(Trim(.Cell(k, 1).Range.Text), 4)) = "item" Then
        N = N + 1
      End If
    Next k
    If N > ComItms Then MsgBox "In defect table, there are " & N & " item rows, which is larger then inspection items of  " & ComItms & _
     "Please delete extra item's row and run program again" & CHR(10) & _
    "Defect" & UText(34920, 26377) & N & UText(20010) & "item" & UText(34892) & "，" & UText(27604, 26816, 39564, 30340, 27454, 25968) & ComItms & UText(22810) & "，" & UText(35831, 21024, 38500, 22810, 20313, 30340) & "Item" & UText(34892, 20877, 36816, 34892, 31243, 24207) & "！", vbCritical: Exit Sub
 
     '" & UText(21024, 38500, 19979, 34920, 27809, 20869, 23481, 30340, 22810, 20313) & "ITEM" & UText(34892) & "
    ' myrow = .Rows.count
    ' myrow = myrow - 2

    ' Do While LCase(Left(Trim(.Cell(myrow, 1).Range.Text), 4)) = "item"
        '.Rows(myrow).Delete
        'myrow = myrow - 1
     'Loop
    '========
    Call SortDefectTable
    myrow = .Rows.count
    myrow = myrow - 2
     
     msg = ""
     mbx = ""
     
     For k = 2 To myrow  '" & UText(20174, 31532, 20108, 34892, 24320, 22987) & "，" & UText(21028, 26029, 26159, 21542) & "item" & UText(34892) & "," & UText(33719, 24471, 27599, 34892, 21333, 20803, 26684, 30340) & "defect " & UText(25968) & "
        If .Rows(k).Cells.count = Coln Then
           If .Cell(k, Coln - 3).Range.Fields.count = 0 And .Cell(k, Coln - 3).Range.Text Like "*Allowed*" = False Then
     
             tem = Mid(.Cell(k, Coln - 3).Range.Text, 1, Len(.Cell(k, Coln - 3).Range.Text) - 2)
             If Len(tem) > 1 Then
               Fw = UCase(Left(tem, 1)) & LCase(Mid(tem, 2, Len(tem) - 1))
             Else
                Fw = UCase(tem)
             End If
   
           .Cell(k, Coln - 3).Range.Text = Fw
            dCr(k) = Val(.Cell(k, Coln - 2).Range.Text)
            dMa(k) = Val(.Cell(k, Coln - 1).Range.Text)
            dMi(k) = Val(.Cell(k, Coln).Range.Text)
           
           
            If dCr(k) < 0 Or dMa(k) < 0 Or dMi(k) < 0 Then msg = msg + " row " & k & ", the defect of " & temw & ", the defect number was minus!" & CHR(13) & _
            UText(31532) & k & UText(34892, 32570, 38519) & temw & UText(32570, 38519, 25968, 20026, 36127, 25968) & "!" & CHR(10)
            
            If Left(Trim(tem), 1) <> CHR(13) And UCase((Mid(Trim(tem), 1, 9))) <> "NO DEFECT" Then
              
                  If dCr(k) = 0 And dMa(k) = 0 And dMi(k) = 0 Then msg = msg + "the row " & k & ", the defect of " & temw & ", the defect number was not input" & CHR(10) & _
                  UText(31532) & k & UText(34892, 32570, 38519) & temw & UText(32570, 38519, 25968, 37327, 27809, 26377, 21028, 23450) & CHR(10)
                      
                  If (dCr(k) > 0 And dMa(k) > 0) Or (dCr(k) > 0 And dMi(k) > 0) Then msg = msg + " row " & k & ", the defect of" & temw & " input with  Critical and other defect." & CHR(10) & _
                  UText(31532) & k & UText(34892) & temw & UText(26377) & "Critical" & UText(32570, 38519) & "，" & UText(21448, 26377, 20854, 20182, 32570, 38519) & "，" & UText(37325, 22797, 21028, 23450) & "!" & CHR(10)
             End If
             
                 
             If UCase((Mid(Trim(tem), 1, 9))) = "NO DEFECT" Then
                     If dCr(k) > 0 Or dMa(k) > 0 Or dMi(k) > 0 Then msg = msg + "row " & k & " , with statement of No Defect and defect number." & CHR(10) & _
                     UText(31532) & k & UText(34892) & "No Defect," & UText(20294, 21448, 26377, 32570, 38519, 25968) & "!" & CHR(10)
              End If
                  
              If Len(tem) = 0 Then
              
                     If dCr(k) > 0 Or dMa(k) > 0 Or dMi(k) > 0 Then msg = msg + "row " & k & ", without defect statement!" & CHR(10) & _
                     UText(31532) & k & UText(34892, 32570, 38519, 25551, 36848, 26410, 20889) & "," & UText(20294, 21448, 26377, 32570, 38519, 25968) & "!" & CHR(10)
              End If
              
                    
         Else
            
           dCr(k) = 0
           dMa(k) = 0
           dMi(k) = 0
          
         End If
      
      End If
     Next k
          
           If msg <> "" Then MsgBox "Defects Table:" & CHR(10) & msg & "Please revise them and run program again" & CHR(10) & _
           "Defects" & UText(34920) & CHR(10) & msg & UText(35831, 23545, 20197, 19978, 36827, 34892, 20462, 25913, 21518, 20877, 27425, 36816, 34892) & "！！！", vbCritical, UText(35686, 21578): Exit Sub
         
         
         
 '====================
       sCr = 0 '" & UText(28165, 38646) & "
       sMa = 0
       sMi = 0
       Dk = 0
     For M = 3 To (myrow + 1)
       If .Rows(M).Cells.count <> 1 Then
        If .Cell(M, Coln - 3).Range.Fields.count = 0 Then
         sCr = sCr + dCr(M)
         sMa = sMa + dMa(M)
         sMi = sMi + dMi(M)
        Else
          .Cell(M, Coln - 2).Range.Text = sCr

          If sCr > Val(.Cell(M + 1, Coln - 2).Range.Text) Then
              With .Cell(M, Coln - 2).Range.Font
               .Color = wdColorRed
              End With
              Dk = Dk + 1
          Else
              With .Cell(M, Coln - 2).Range.Font
                .Color = wdColorBlack
              End With
          End If
         .Cell(M, Coln - 1).Range.Text = sMa
               If sMa > Val(.Cell(M + 1, Coln - 1).Range.Text) Then
                   With .Cell(M, Coln - 1).Range.Font
                    .Color = wdColorRed
                   End With
                   Dk = Dk + 1
               Else
                    With .Cell(M, Coln - 1).Range.Font
                      .Color = wdColorBlack
                    End With
               End If
          .Cell(M, Coln).Range.Text = sMi
               If sMi > Val(.Cell(M + 1, Coln).Range.Text) Then
                   With .Cell(M, Coln).Range.Font
                      .Color = wdColorRed
                  End With
                  Dk = Dk + 1
                Else
                    With .Cell(M, Coln).Range.Font
                       .Color = wdColorBlack
                    End With
                End If
           sCr = 0
           sMa = 0
           sMi = 0
        End If
      End If
     Next M
   Res = Mid(ThisDocument.Range.Tables(tbn - 1).Cell(1, 2).Range.Text, 1, Len(ThisDocument.Range.Tables(tbn - 1).Cell(1, 2).Range.Text) - 1)
 
     
    'SaP " & UText(39033) & "

   TBName = "*special attention point for workmanship*"
   tb = ThisDocument.Tables.count - 1
   For X = tbn To tb
   With ThisDocument.Range.Tables(X)
        If LCase(.Cell(1, 1).Range.Text) Like TBName = True Then
            tbn1 = X
            GoTo 20
        End If
     End With
  Next X
  Sk = 0
  GoTo 30
20    With ThisDocument.Range.Tables(tbn1)
     Rown = .Rows.count
     Sk = 0
     For i = 2 To Rown
        If Len(Trim(.Cell(i, 2).Range.Text)) > 2 Then
           Sk = Sk + 1
        End If
     Next i
     End With
 
30   If Dk > 0 Then
      If Res <> "Not Conformed" Then msg = msg + "In defects table, there are " & Dk & _
      " items which defect exceed the allowance, however the result of workmanship table was not 'Not Conformed'，Please choose Not Conformed!!!" & CHR(13) & _
       "Defect" & UText(34920, 20013, 26377) & Dk & UText(39033, 32570, 38519, 25968, 22823, 20110, 20801, 35768, 20540) & "，workmanship" & UText(34920, 20013, 30340, 32467, 35770, 19981, 26159) & "Not Conformed，" & UText(35831, 36873, 25321) & "Not Conformed!!!" & CHR(13)
    Else
      If Sk = 0 Then
        If Res <> "Conformed" Then msg = msg + "Defect number in defect table within allowance and without SAP，However the result of wokmanship was not 'Conformed'！！！" + CHR(13) & _
      "Defect" & UText(34920, 20013, 30340, 32570, 38519, 25968, 22312, 20801, 35768, 33539, 22260, 20869, 19988, 27809, 26377) & "SAP，" & UText(29616, 22312, 30340) & "Workmanship" & UText(32467, 35770, 21364, 19981, 26159) & "Conformed！！！" + CHR(13)
     Else
         If Res <> "Not Conformeded" Then msg = msg + "Defect number within allowance but there are SAPs for workmanship, however the result was not ‘Not Conformed’！！！" & CHR(13) & _
       "Degect" & UText(34920, 20013, 30340, 32570, 38519, 25968, 22312, 20801, 35768, 33539, 22260, 20869, 20294, 26377) & "SAP" & UText(39033) & "，" & UText(29616, 22312, 30340, 32467, 35770, 21364, 19981, 26159) & "Not Conformed！！！" + CHR(13)
    End If
  End If
 End With
 
    If msg <> "" Then MsgBox msg & CHR(10) & CHR(10) & "Please revise it and run program again!" & CHR(10) & _
    UText(35831, 26816, 26597, 26356, 27491, 24182, 37325, 26032, 36816, 34892, 31243, 24207) & "！", vbCritical
   
  
   
   
End Sub
Sub SortDefectTable()
Dim Rn, i, j, k As Integer
Dim N As Integer
Dim temp As Integer
Dim M As Integer

Dim Coln As Integer
Dim TBName As String
Dim StRow() As Integer
Dim EDRow() As Integer

Dim Dn As Integer

Dim BK() As String
Dim SortArray() As Variant
Dim tempArray() As Variant
Dim Cn As Integer
 
   
   TBName = "defects"
   tb = ThisDocument.Tables.count - 1
   For X = 1 To tb
   With ThisDocument.Range.Tables(X)
        If Left(LCase(Trim(.Cell(1, 1).Range.Text)), 7) = TBName Then
            tbn = X
            GoTo 10
        End If
     End With
  Next X
 
    
10  Set tbl = ThisDocument.Range.Tables(tbn)
    With tbl
        Rn = .Rows.count
     
        N = 0
        For i = 2 To Rn - 2
           If .Rows(i).Cells.count = 1 Then N = N + 1
        Next i
        
        ReDim StRow(1 To N)
        ReDim EDRow(1 To N)
       
        
        N = 0
        For i = 2 To Rn
           If .Rows(i).Cells.count = 1 Then
              N = N + 1
              StRow(N) = i + 1
              If N > 1 And Mid(.Cell(i - 1, 2).Range.Text, 1, Len(.Cell(i - 1, 2).Range.Text) - 2) <> "Allowed:" Then EDRow(N - 1) = i - 1
            Else
              If .Cell(i, 2).Range.Fields.count > 0 Or .Cell(i, 2).Range.Text Like "*Total Found*" = True Then
                EDRow(N) = i - 1
              End If
           End If
        Next i
        
   
    
    For M = 1 To N
      Dn = EDRow(M) - StRow(M) + 1
      ReDim BK(1 To Dn)
      ReDim SortArray(1 To Dn, 1 To 5)
       Cn = 0
      For i = StRow(M) To EDRow(M)
         Set Cel = .Cell(i, 2)
         If Cel.Range.Bookmarks.count > 0 Then
          BK(i - StRow(M) + 1) = Cel.Range.Bookmarks(1).Name
        Else
          Cn = Cn + 1
           BK(i - StRow(M) + 1) = "NA" & "_" & Cn
        End If
        For j = 2 To 5
            Set Cel = .Cell(i, j)
            SortArray(i - StRow(M) + 1, j) = Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2)
         Next j
      
    Next i
   
       Call BubbleSort(SortArray, BK, Dn)
   
    For i = StRow(M) To EDRow(M)
        For j = 2 To 5
            If j = 2 Then
                 Call DeleteBookMark(tbl.Cell(i, j))
                 tbl.Cell(i, j).Range.Bookmarks.add Name:=BK(i - StRow(M) + 1)
                 tbl.Cell(i, j).Range.Text = SortArray(i - StRow(M) + 1, j)
            ElseIf j = 1 Then
              
            Else
                tbl.Cell(i, j).Range.Text = SortArray(i - StRow(M) + 1, j)
            End If
        Next j
    Next i
    Next M

    
    End With
End Sub
Sub BubbleSort(ByRef SortArray As Variant, BK, Dn)
' " & UText(20882, 27873, 25490, 24207, 31639, 27861) & "，" & UText(21516, 26102, 25490, 24207, 20070, 31614) & "

    Dim i As Integer, j As Integer
    Dim tempBookmark  As String
    Dim tempArray As Variant
    ReDim tempArray(1 To 5)

    For i = 1 To Dn
        For j = i + 1 To Dn
            If SortArray(i, 3) < SortArray(j, 3) Or _
               (SortArray(i, 3) = SortArray(j, 3) And SortArray(i, 4) < SortArray(j, 4)) Or _
               (SortArray(i, 3) = SortArray(j, 3) And SortArray(i, 4) = SortArray(j, 4) And SortArray(i, 5) < SortArray(j, 5)) Then
                ' " & UText(20132, 25442, 34892) & "
                For k = 2 To 5
                    tempArray(k) = SortArray(i, k)
                    SortArray(i, k) = SortArray(j, k)
                    SortArray(j, k) = tempArray(k)
                Next k
                ' " & UText(20132, 25442, 20070, 31614) & "
                tempBookmark = BK(i)
                BK(i) = BK(j)
                BK(j) = tempBookmark
            End If
        Next j
    Next i
  
End Sub
Sub qtytotal()
'
'
' qtytotal " & UText(23439) & "  9" & UText(39033, 27169, 26495) & "，by wyn/20241216
'
'
Dim qtyRow, qtyCol, k, p, Q, R, cnt, CeC, CRow, Im(200), Sm(200) As Integer ' Itms
Dim col(200, 9), scol(9), Nfw, Naw As Long
Dim tem(10), txt(200, 10), Kw, msg, ALW As String
Dim Celn, PSQ, PSE, Dif As Integer
Dim TBName, Ufw, Uaw, KwQ As String

   TBName = "1. Quantity"
   tb = ThisDocument.Tables.count - 1
   For X = 1 To tb
   With ThisDocument.Range.Tables(X)
        Kw = Trim(.Cell(1, 1).Range.Text)
        Kw = Mid(Kw, 1, Len(Kw) - 2)
        If Kw = TBName Then
          tbn = X
          GoTo 5
        End If
      
     End With
  Next X
    'tbn = getTableNumber(1)
5   With ThisDocument.Range.Tables(tbn)
     qtyRow = .Rows.count
     qtyCol = .Columns.count
      For i = 5 To qtyRow
         .Cell(i, 1).Range.Rows.Select
          Selection.Range.Font.ColorIndex = wdBlack
      Next i
     
        '" & UText(21028, 23450, 26377, 27809, 26377, 21512, 24182, 21333, 20803, 26684) & "
     CRow = qtyRow - 5
     CeC = ThisDocument.Tables(tbn).Range.Cells.count


     If (CeC - 25) / qtyCol <> CRow Then MsgBox "Quantity Table has mergecell,program can't deal with it,please split it to default frame." & _
     CHR(10) & "(Quantity" & UText(34920, 26684, 26377, 21512, 24182, 21333, 20803, 26684) & "，" & UText(19981, 31526, 21512, 35201, 27714) & "，" & UText(31243, 24207, 19981, 33021, 35745, 31639) & "，" & UText(35831, 37325, 26032, 25286, 20998, 21333, 20803, 26684) & "，" & UText(24674, 22797, 21407, 26377, 26684, 24335, 21518, 20877, 36816, 34892, 26412, 31243, 24207) & ")！！！": Exit Sub
  
     For p = qtyRow - 1 To 5 Step -1
        cnt = 0
        For Q = 1 To 9
          txt(p, Q) = Trim(.Cell(p, Q).Range.Text)
          col(p, Q) = Val(txt(p, Q))
        If Left(txt(p, Q), 1) = CHR(13) Then cnt = cnt + 1
        Next Q
        
        If cnt = 9 Then .Cell(p, 1).Range.Rows.Delete: GoTo 10
        
        If Left(txt(p, 1), 1) <> CHR(13) Or Left(txt(p, 2), 1) <> CHR(13) Then
        
          For Q = 3 To 9
             If Left(txt(p, Q), 1) = CHR(13) Then msg = msg + TBName & "table: row" & p & ", some cells without inputing data,please note." & CHR(13) _
             & TBName & UText(34920, 26684, 31532) & p & UText(34892, 26377, 34920, 26684, 27809, 26377, 36755, 20837, 25968, 25454) & "，" & UText(35831, 27880, 24847) & "!" & CHR(10)
          Next Q
          End If
          
        If Left(txt(p, 1), 1) = CHR(13) And Left(txt(p, 2), 1) = CHR(13) Then
          For Q = 3 To 9
            If Left(txt(p, Q), 1) <> CHR(13) Then msg = msg + TBName & " Table: row" & p & " without inputing P.O. No. or Item No.,please note." & _
             CHR(10) & TBName & UText(34920, 26684, 31532) & p & UText(34892, 26377, 25968, 25454) & "，" & UText(20294, 21364, 27809, 26377, 36755, 20837) & "P.O. No. " & UText(25110) & "Item No." & UText(35831, 27880, 24847) & "!" & CHR(10)
          Next Q
        End If
           
        For Q = 3 To 9
          .Cell(qtyRow, (Q - 1)).Range.Text = 0
           scol(Q) = scol(Q) + col(p, Q)
        Next Q
    
10    Next p

               If msg <> "" Then
           Click = MsgBox(msg & CHR(10) & "Click (Yes) to Continue running the program or Click (No) to exit program" & CHR(10) & _
           UText(28857, 20987) & "（" & UText(26159) & "）" & UText(32487, 32493, 36816, 34892, 31243, 24207, 25110, 28857, 20987) & "（" & UText(21542) & "）" & UText(36864, 20986, 31243, 24207), vbYesNo)
            msg = ""
            If Click <> vbYes Then Exit Sub
        End If

 
       
      

    qtyRow = .Rows.count
    
     For R = 3 To 9
      .Cell(qtyRow, R - 1).Range.Text = scol(R) 'r-1 " & UText(25187, 38500, 21512, 24182, 21333, 26080, 26684) & "
     Next R
     
     Itms = qtyRow - 5 '" & UText(27454, 25968, 20197, 21450, 27454, 21495) & "，" & UText(21518, 38754, 20889, 20837) & "workmanship " & UText(34920, 29992) & "
     ReDim Ptx(1 To Itms)
     ReDim Itx(1 To Itms)
     Dim Psm() As Integer
     ReDim Psm(1 To Itms)
     StoretoProperty "ItemsQty", CStr(Itms)
     For i = 5 To qtyRow - 1
       Ptx(i - 4) = Mid(.Cell(i, 1).Range.Text, 1, Len(.Cell(i, 1).Range.Text) - 2)
       Itx(i - 4) = Mid(.Cell(i, 2).Range.Text, 1, Len(.Cell(i, 2).Range.Text) - 2)
       Sm(i - 4) = Val(.Cell(i, 8).Range.Text)     '" & UText(27599, 27454, 25277, 26679, 37327) & "
       Psm(i - 4) = Val(.Cell(i, 9).Range.Text)
     Next i
     sms = Val(.Cell(qtyRow, 7).Range.Text)
     
     Dim dictPO As Object
     Dim dictItem As Object
     Dim PONO, ItemNO As String
     
     Set dictPO = CreateObject("Scripting.Dictionary")
     For i = 1 To Itms
        If Not dictPO.Exists(Ptx(i)) Then
            dictPO.add Ptx(i), Nothing
        End If
    Next i
 
    StoretoProperty "ComPOsQty", CStr(dictPO.count)
   
    For Each Key In dictPO.Keys
        PONO = PONO & "," & Key
    Next Key
       PONO = Mid(PONO, 2, Len(PONO) - 1)
       
       Set tbl = ThisDocument.Tables(1)
       For Each Cel In tbl.Range.Cells
         If Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2) = "P.O. No." Then
           Cel.Next.Range.Text = PONO
           GoTo 12
         End If
       Next Cel
12   Set dictPO = Nothing
 
 Dim ComItx() As String
 ReDim ComItx(1 To Itms)
 Dim ComSm() As Integer
 ReDim ComSm(1 To Itms)
 Dim ComPSm() As Integer
 ReDim ComPSm(1 To Itms)
 Dim ComItms As Integer
      
   
     N = 0
      Set dictItem = CreateObject("Scripting.Dictionary")
      For i = 1 To Itms
       
        If Not dictItem.Exists(Itx(i)) Then
            N = N + 1
            dictItem.add Itx(i), Psm(i)
            ComPSm(N) = Psm(i)
        Else
            dictItem(Itx(i)) = dictItem(Itx(i)) + Psm(i)
            ComPSm(N) = dictItem(Itx(i))
        End If
    Next i
    ComItms = dictItem.count


    StoretoProperty "ComItemsQty", CStr(dictItem.count)
    
    For Each Key In dictItem.Keys
        ItemNO = ItemNO & "," & Key
    Next Key
       ItemNO = Mid(ItemNO, 2, Len(ItemNO) - 1)
        For Each Cel In tbl.Range.Cells
         If Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2) = "Item No." Then
           Cel.Next.Range.Text = ItemNO
           GoTo 15
         End If
       Next Cel
      
15    Set dictItem = Nothing

    '" & UText(21516, 27493, 36873, 25321, 30340, 21333, 20301) & "
    
   .Cell(4, 3).Range.Text = .Cell(4, 4).Range.Text
   .Cell(4, 6).Range.Text = .Cell(4, 4).Range.Text
   .Cell(4, 8).Range.Text = .Cell(4, 4).Range.Text
   .Cell(4, 7).Range.Text = .Cell(4, 5).Range.Text
   .Cell(4, 9).Range.Text = .Cell(4, 5).Range.Text
   
     '" & UText(26816, 26597, 20135, 21697, 21333, 20301) & " " & UText(26159, 21542, 19968, 33268) & "
     Celn = ThisDocument.Tables(1).Range.Cells.count
   
     For i = 1 To Celn
        KwQ = ThisDocument.Tables(1).Range.Cells(i).Range.Text
        KwQ = Mid(KwQ, 1, Len(KwQ) - 2)
        If KwQ = "Quantity" Then
           PSQ = i
        ElseIf KwQ = "Sequence" Then
           PSE = i
           GoTo 20
        End If
     Next i
     
20   If ThisDocument.Tables(1).Range.Cells(PSQ + 1).Range.words.count = 3 Then
       Ufw = ThisDocument.Tables(1).Range.Cells(PSQ + 1).Range.words(2)
     Else
      
       MsgBox UText(25253, 21578, 39318, 39029, 34920, 26684) & "Quantity" & UText(25968, 25454, 26410, 36755, 20837, 25110, 26684, 24335, 19981, 31526, 21512, 35201, 27714) & "，" & UText(27491, 30830, 30340, 36755, 20837, 26041, 24335, 26159, 25968, 23383) & "+" & UText(31354, 26684) & "+" & HubKeyUnit() & "，" & UText(35831, 36755, 20837, 21518, 36827, 34892, 36816, 34892, 31243, 24207) & "！" & CHR(13) & _
       "The quantity in firrst table of first page without data or the inputing is in incorrect format, the format shall be number+space+unit,please revise it and run program again!", vbApplicationModal: Exit Sub
     End If
     If Right(Ufw, 1) = "s" Then Ufw = Mid(Ufw, 1, Len(Ufw) - 1)
     Uaw = .Cell(4, 3).Range.words(2)
     Upw = .Cell(4, 7).Range.words(2)
     
    'If Uaw <> Ufw Then MsgBox "P.O. Quantity" & Ufw & ", 1. Quantity" & Uaw _
     & UText(20004, 32773, 19981, 19968, 26679) & "，" & UText(35831, 20462, 25913, 21518, 20877, 27425, 36816, 34892, 26412, 31243, 24207) & "！！！" & Chr(13) & UText(27880, 24847) & "1. Quantity" & UText(34920, 26684, 30340, 20135, 21697, 21333, 20301, 30340, 20462, 25913, 26159, 28857, 34920, 26684, 31532) & "4" & UText(21015, 30340) & "Shipment Qty" & UText(19979, 26041, 36827, 34892, 36873, 25321) & "，" & UText(36816, 34892, 31243, 24207, 21518, 34920, 26684, 20854, 20182, 22788, 30340, 21333, 20301, 23558, 33258, 34892, 21516, 27493, 20462, 25913) & "，" & UText(32780, 25253, 21578, 39318, 39029, 34920, 26684) & "P.O. Quantity" & UText(22788, 30340, 21333, 20301, 19981, 20250, 21516, 27493) & "，" & UText(38656, 35201, 25163, 24037, 21435, 20462, 25913) & "!": Exit Sub
      
        
     If LCase(Uaw) <> LCase(Ufw) Then msg = msg & "The unit in fist page is " & Ufw & " while it in Quantity table is " & Uaw _
     & " please check and revise it and than run this program again" & CHR(13) & _
     UText(25253, 21578, 39318, 39029, 34920, 26684) & "P.O. Quantity" & UText(30340, 20135, 21697, 21333, 20301, 26159) & Ufw & ", " & UText(32780) & "Quantity" & UText(34920, 26684, 30340, 20135, 21697, 21333, 20301, 26159) & Uaw _
     & UText(20004, 32773, 19981, 19968, 26679) & "，" & UText(35831, 20462, 25913, 21518, 20877, 27425, 36816, 34892, 26412, 31243, 24207) & "！！！" & CHR(13)
      
      
      '" & UText(26816, 26597) & "PO" & UText(25968, 37327, 26159, 21542, 19968, 33268) & "
    
     Nfw = Val(ThisDocument.Tables(1).Range.Cells(PSQ + 1).Range.words(1))
     Naw = Val(.Cell(qtyRow, 2).Range.Text)
     
    'If Nfw <> Naw Then MsgBox "P.O. Quantity" & Nfw & ", 1. QuantityPO Qty" & Naw & _
     UText(20004, 32773, 19981, 19968, 26679) & "，" & UText(35831, 20462, 25913, 21518, 20877, 27425, 36816, 34892, 26412, 31243, 24207) & "！！！": Exit Sub
     
      If Nfw <> Naw Then MsgBox "P.O. Quantity in first page is" & Nfw & "while it in Quantity table is" & Naw & _
     "Please check and revise it." & CHR(13) & UText(25253, 21578, 39318, 39029, 34920, 26684) & "P.O. Quantity" & UText(30340, 20135, 21697, 25968, 37327, 26159) & Nfw & ", " & UText(32780) & "1. Quantity" & UText(34920, 26684, 30340) & "PO Qty" & UText(25968, 37327, 26159) & Naw & _
     UText(20004, 32773, 19981, 19968, 26679) & "，" & UText(35831, 20462, 25913, 21518, 20877, 27425, 36816, 34892, 26412, 31243, 24207) & "！！！", vbCritical: Exit Sub
     
     If scol(6) > scol(4) Then MsgBox "Packed Qty is found larger than Shipment Qty，Please check and revise it！！！" & CHR(13) & _
     "Quantity" & UText(34920, 26684) & ",Packed Qty" & UText(20135, 21697, 25968, 22823, 20110) & "Shipment Qty" & UText(20135, 21697, 25968) & "，" & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892, 31243, 24207) & "！！！": Exit Sub
     If scol(7) > scol(5) Then MsgBox "Packed carton number is found larger than Shipment carton number，Please check and revise it！！！" & CHR(13) & _
     "Quantity" & UText(34920, 26684) & ",Packed carton" & UText(25968, 22823, 20110) & "Shipment carton" & UText(25968) & "，" & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892, 31243, 24207) & "！！！": Exit Sub
     
     '" & UText(21028, 23450, 25277, 31665, 25968, 26159, 21542, 22826, 23569) & "
     If LCase(ThisDocument.Tables(1).Range.Cells(PSE + 1).Range.Text) Like "*original inspection*" = True Or _
     LCase(ThisDocument.Tables(1).Range.Cells(PSE + 1).Range.Text) Like "*first inspection*" = True Then
         If scol(9) < Int(Sqr(scol(5))) Then msg = msg & TBName & "table" & "the selected carton number is less than sqrt of shipment carton number" & _
         Int(Sqr(scol(5))) & "please check and if you revise the data, run this program again！！！." & CHR(13) & _
         TBName & UText(34920, 26684, 20135, 21697, 25968, 37327, 20026) & scol(5) & UText(24320, 26041, 25968, 21462, 25972, 21518, 20026) & Int(Sqr(scol(5))) & UText(29616, 22312, 25277, 31665, 25968, 20026) & scol(9) & UText(26159, 21542, 22826, 23569) & "，" & UText(35831, 30830, 35748, 26159, 21542, 20462, 25913) & "，" & UText(33509, 26377, 20462, 25913, 35831, 25913, 21518, 20877, 36816, 34892, 31243, 24207)
     End If
     If ThisDocument.Tables(1).Range.Cells(PSE + 1).Range.Text Like "*Re-inspection*" = True Or LCase(ThisDocument.Tables(1).Range.Cells(PSE + 1).Range.Text) Like "*re-inspection*" = True Then
      If scol(9) < Int(Sqr(scol(5)) * 1.5) Then msg = msg & TBName & "table" & "the selected carton number:" & scol(9) & " is less than sqrt of shipment carton number times 1.5= " & _
         Int(Sqr(scol(5)) * 1.5) & ",please check, and if you revise the data, run this program again！！！." & CHR(13) & _
         TBName & UText(34920, 26684, 20135, 21697, 25968, 37327, 20026) & scol(5) & UText(24320, 26041, 25968, 20056, 20197) & "1.5" & UText(20493, 21462, 25972, 21518, 20026) & Int(Sqr(scol(5)) * 1.5) & UText(29616, 22312, 25277, 31665, 25968, 20026) & scol(9) & UText(26159, 21542, 22826, 23569) & "，" & UText(35831, 30830, 35748, 26159, 21542, 20462, 25913) & "，" & UText(33509, 26377, 20462, 25913, 35831, 25913, 21518, 20877, 36816, 34892, 31243, 24207)
     End If
     
 End With
    
     

     If scol(3) <> scol(4) Or scol(6) / scol(4) < 0.8 Then
     
         tbn1 = 0
         TBName = "Special Attention Point for Quantity"
         tb = ThisDocument.Tables.count - 1
         For X = 1 To tb
            With ThisDocument.Range.Tables(X)
               Kw = Trim(.Cell(1, 1).Range.Text)
               Kw = Mid(Kw, 1, Len(Kw) - 2)
               If Kw = TBName Then
                   tbn1 = X
                   GoTo 50
               End If
             End With
        Next X
        If ThisDocument.Tables(tbn + 1).Rows(1).Range.Cells.count = 2 Then
           ThisDocument.Tables(tbn + 1).Select
           Dif = 2
        Else
           ThisDocument.Tables(tbn).Select
           Dif = 1
        End If
        Selection.MoveDown
        Selection.TypeParagraph
        ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=2, NumColumns:= _
        2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
              With Selection.Tables(1)
                  If .Style <> UText(32593, 26684, 22411) Then
                   .Style = UText(32593, 26684, 22411)
                  End If
                  Call Setborders025
                 .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                 .Columns(1).Width = CentimetersToPoints(0.8)
                 .Rows(1).Range.Cells.Merge
                 .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                 .Rows(1).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 '" & UText(26080, 33394) & " -603917569  '" & UText(27973, 28784) & "
                 .Rows(2).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603914241 '" & UText(26080, 33394) & " -603917569  '" & UText(27973, 28784) & "
                 .Cell(1, 1).Range.Text = TBName
                
                  tbn1 = tbn + Dif
             End With
   

50 With ThisDocument.Tables(tbn1)
      Rn = .Rows.count
      If Rn = 1 Then
         .Cell(1, 1).Select
         Selection.InsertRowsBelow 1
         Selection.Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603914241 '" & UText(26080, 33394) & " -603917569  '" & UText(27973, 28784) & "
         Selection.Cells.Split NumRows:=1, NumColumns:=2, mergebeforesplit:=False
         .Cell(2, 1).Width = CentimetersToPoints(0.8)
         .Cell(2, 2).Width = CentimetersToPoints(17)
         Selection.InsertRowsBelow 1
      ElseIf Rn > 3 Then
         For i = Rn - 2 To 2 Step -1
           .Rows(i).Delete
         Next i
      End If
    
      R = 0
       For i = 5 To qtyRow - 1
          If col(i, 3) <> col(i, 4) Then '" & UText(21333, 27454, 20986, 36135, 37327, 21516, 35746, 21333, 19981, 21516) & "
             ThisDocument.Tables(tbn).Cell(i, 4).Range.Font.ColorIndex = wdRed
             ALW = ALW & "for item " & "((" & Itx(i - 4) & "))" & "," & " shipment quantity was " & "((" & col(i, 4) & "))" & _
             " " & LCase(Uaw) & "s" & " instead of " & "((" & col(i, 3) & "))" & " " & LCase(Uaw) & "s; "
           End If
       Next i
       'If R = 1 Then ALW = UCase(Mid(ALW, 1, 1)) & LCase(Mid(ALW, 2, Len(ALW) - 1)) & Chr(10)
 
     If scol(3) <> scol(4) Then   '" & UText(24635, 20986, 36135, 37327, 36319, 35746, 21333, 19981, 31526) & "
        ThisDocument.Tables(tbn).Cell(qtyRow, 3).Range.Font.ColorIndex = wdRed
        If R = 0 Then R = R + 1
        ALW = ALW & "total shipping quantity was " & "((" & scol(4) & "))" & " " & LCase(Uaw) & "s" & " instead of " & _
        "((" & scol(3) & "))" & " " & LCase(Uaw) & "s" & " as stated on the order sheet."
        'Thisdocument.Tables(tbn).Cell(R + 1, 1).Range.Text = "1." & R
        'Thisdocument.Tables(tbn).Cell(R + 1, 2).Range.Text = "Total shipping quantity was " & scol(3) & " " & Uaw & " instead of " & _
        scol(3) & " " & Uaw & " as stated on the order sheet."
      End If
   
      If R = 1 Then
        .Rows(R + 1).Select
        Selection.InsertRowsAbove 1
        .Cell(R + 1, 2).Range.Text = UCase(Mid(ALW, 1, 1)) & Mid(ALW, 2, Len(ALW) - 1)
      End If
   
      If scol(6) / scol(4) < 0.8 Then '" & UText(21253, 35013, 25968, 23569, 20110) & "80%
       R = R + 1
      .Rows(R + 1).Select
       Selection.InsertRowsAbove 1
       .Cell(R + 1, 2).Range.Text = "Upon arrival, only " & "((" & scol(6) & "))" & " " & LCase(Uaw) & _
       "s" & "((" & " (" & Format(scol(6) / scol(4) * 100, "00.00") & "% ) " & "))" & " of products were packed into " & "((" & LCase(Upw) & "s " & "))" & _
       "while " & "(( " & LCase(Uaw) & "s" & "))" & " of product has finished production but not packed and else (( " & LCase(Uaw) & _
       "s" & "))" & " of product are under production. " & "The inspection was approved and to carry out as usual. " & _
       "We try our best to draw samples for inspection from both packed and finished unpacked products in proportion. " & _
       "Please note the potential risk of sampling of this inspection."
     End If
     Rn = .Rows.count
     For i = 2 To Rn
      .Cell(i, 1).Range.Text = "1." & i - 1
     
     Next i
     
   End With
 End If
 If scol(6) / scol(4) >= 0.8 And scol(6) < scol(4) Then
       tbn1 = 0
       TBName = "Reference Note for Quantity"
       tb = ThisDocument.Tables.count - 1
       For X = 1 To tb
          With ThisDocument.Range.Tables(X)
             Kw = Trim(.Cell(1, 1).Range.Text)
             Kw = Mid(Kw, 1, Len(Kw) - 2)
             If Kw = TBName Then
              tbn1 = X
              GoTo 60
             End If
          End With
      Next X
      tbn = 0
      TBName = "2. Workmanship"
    
      tb = ThisDocument.Tables.count - 1
      For X = 1 To tb
        With ThisDocument.Range.Tables(X)
           Kw = Trim(.Cell(1, 1).Range.Text)
           Kw = Mid(Kw, 1, Len(Kw) - 2)
           If Kw = TBName Then
              tbn = X
              GoTo 55
            End If
        End With
      Next X
    
55   If ThisDocument.Tables(tbn - 1).Rows(1).Range.Cells.count = 2 Then
          ThisDocument.Tables(tbn - 1).Select
          Selection.MoveUp
          Selection.TypeParagraph
          ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=3, NumColumns:= _
             2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
              With Selection.Tables(1)
                If .Style <> UText(32593, 26684, 22411) Then
                  .Style = UText(32593, 26684, 22411)
                End If
                Call Setborders025
                .Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603914241 '" & UText(26080, 33394) & " -603917569  '" & UText(27973, 28784) & "
                .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                
                .Columns(1).Width = CentimetersToPoints(0.8)
                .Rows(1).Range.Cells.Merge
                .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Rows(1).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 '" & UText(26080, 33394) & " -603917569  '" & UText(27973, 28784) & "
                '.Rows(2).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603914241 '" & UText(26080, 33394) & " -603917569  '" & UText(27973, 28784) & "
                .Cell(1, 1).Range.Text = "Reference Note for Quantity"
               
             End With
    Else
            ThisDocument.Tables(tbn).Select
            Selection.MoveUp
            Selection.TypeParagraph
            ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=3, NumColumns:= _
             2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
              With Selection.Tables(1)
                 If .Style <> UText(32593, 26684, 22411) Then
                    .Style = UText(32593, 26684, 22411)
                  End If
                  Call Setborders025
                  .Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603914241 '" & UText(26080, 33394) & " -603917569  '" & UText(27973, 28784) & "
                  .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                  .Columns(1).Width = CentimetersToPoints(0.8)
                  .Rows(1).Range.Cells.Merge
                  .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                  .Rows(1).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 '" & UText(26080, 33394) & " -603917569  '" & UText(27973, 28784) & "
                 ' .Rows(2).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603914241 '" & UText(26080, 33394) & " -603917569  '" & UText(27973, 28784) & "
                  .Cell(1, 1).Range.Text = "Reference Note for Quantity"
                 
             End With
             Selection.MoveDown
               Selection.TypeParagraph
               ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=2, NumColumns:= _
               2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
                 With Selection.Tables(1)
                   If .Style <> UText(32593, 26684, 22411) Then
                      .Style = UText(32593, 26684, 22411)
                   End If
                   Call Setborders025
                  .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                  .Rows(1).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603914241 '" & UText(26080, 33394) & " -603917569  '" & UText(27973, 28784) & "
                  .Rows(2).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603914241 '" & UText(26080, 33394) & " -603917569  '" & UText(27973, 28784) & "
               
                End With
     End If
     tbn1 = tbn - 2
60  With ThisDocument.Tables(tbn1)
      Rn = .Rows.count
      If Rn > 3 Then
         For i = Rn - 2 To 2 Step -1
          .Rows(i).Delete
        Next i
      End If
     .Rows(2).Select
     Selection.InsertRowsAbove 1
    
      .Cell(2, 2).Range.Text = "Upon arrival, only " & "((" & scol(6) & " " & LCase(Uaw) & "s" & "))" & _
      "((" & " (" & Format(scol(6) / scol(4) * 100, "00.00") & "% ) " & "))" & " of products were packed into " & _
      "((" & scol(7) & " " & LCase(Upw) & "s" & "))" & " while " & "((" & scol(4) - scol(6) & LCase(Uaw) & "s" & "))" & _
      " of product has finished production" & " but not packed, the inspected samples were drawn from packed and unpacked products in proportion."
     Rn = .Rows.count
     For i = 2 To Rn
       .Cell(i, 1).Range.Text = "1." & i - 1
     Next i

    End With
End If

 
200  If msg <> "" Then MsgBox msg & "Please check and make the necessary corrections" & CHR(10) & _
UText(35831, 26816, 26597, 24182, 36827, 34892, 24517, 35201, 30340, 26356, 27491) & "！", vbCritical

 
1000 End Sub
Sub AQL(SQty As Long)
'
' by wyn " & UText(29983, 25104) & "sample size" & UText(31561, 23439) & "
'
       Dim Typ, Styp, lvl As Integer
       Dim itype, msg, Aqlt, Smst, tmp As String
       Dim i, j, k, X, Cn, msgn, Wn, Pn, WL, Coln, Rown, ss, grp, Agrp, Hgrp, Cgrp, SmpS(10), AAc(10), ADef(10), Scode, Ass As Integer
       Dim AQ(20), AQL(10), Acr, AMa, Ami, AFn, AFnMa As Single
       Dim rg, rgs As Range
       Dim Lots(20) As Long
       Dim CD(20), LG(20, 10) As Integer
       Dim AC(20, 10), MV(20, 10) As Integer

       'Lots()" & UText(25209, 37327, 25968) & "，" & UText(36215, 22987, 20540) & "，" & UText(21040, 19979, 19968, 25209, 37327, 25968) & "-1 " & UText(26159, 20854, 33539, 22260) & "
       Lots(1) = 2
       Lots(2) = 9
       Lots(3) = 16
       Lots(4) = 26
       Lots(5) = 51
       Lots(6) = 91
       Lots(7) = 151
       Lots(8) = 281
       Lots(9) = 501
       Lots(10) = 1201
       Lots(11) = 3201
       Lots(12) = 10001
       Lots(13) = 35001
       Lots(14) = 150001
       Lots(15) = 500001
    
       
       'CD()" & UText(26679, 26412, 20195, 30721) & "，" & UText(23545, 24212) & "A,B,C,....R"
        CD(0) = 0
        CD(1) = 2
        CD(2) = 3
        CD(3) = 5:
        CD(4) = 8:
        CD(5) = 13:
        CD(6) = 20
        CD(7) = 32
        CD(8) = 50
        CD(9) = 80
        CD(10) = 125
        CD(11) = 200
        CD(12) = 315
        CD(13) = 500
        CD(14) = 800
        CD(15) = 1250
        CD(16) = 2000
        
       'AQL() AQL , 1.0 6.5
        AQ(0) = 0
        AQ(1) = 1#
        AQ(2) = 1.5
        AQ(3) = 2.5
        AQ(4) = 4#
        AQ(5) = 6.5
        
        
        
        'LG() ，， S-1,S-2，S-3，S-4，G-I，G-II,G-III
      
       LG(0, 1) = CD(0):      LG(0, 2) = CD(0):       LG(0, 3) = CD(0)
   
       LG(1, 1) = CD(1):      LG(1, 2) = CD(1):       LG(1, 3) = CD(1):    LG(1, 4) = CD(1):     LG(1, 5) = CD(1):   LG(1, 6) = CD(1):       LG(1, 7) = CD(2):
       LG(2, 1) = CD(1):      LG(2, 2) = CD(1):       LG(2, 3) = CD(1):    LG(2, 4) = CD(1):     LG(2, 5) = CD(1):   LG(2, 6) = CD(2):       LG(2, 7) = CD(3):
       LG(3, 1) = CD(1):      LG(3, 2) = CD(1):       LG(3, 3) = CD(2):    LG(3, 4) = CD(2):     LG(3, 5) = CD(2):   LG(3, 6) = CD(3):       LG(3, 7) = CD(4):
       LG(4, 1) = CD(1):      LG(4, 2) = CD(2):       LG(4, 3) = CD(2):    LG(4, 4) = CD(3):     LG(4, 5) = CD(3):   LG(4, 6) = CD(4):       LG(4, 7) = CD(5):
       LG(5, 1) = CD(2):      LG(5, 2) = CD(2):       LG(5, 3) = CD(3):    LG(5, 4) = CD(3):     LG(5, 5) = CD(3):   LG(5, 6) = CD(5):       LG(5, 7) = CD(6):
       LG(6, 1) = CD(2):      LG(6, 2) = CD(2):       LG(6, 3) = CD(3):    LG(6, 4) = CD(4):     LG(6, 5) = CD(4):   LG(6, 6) = CD(6):       LG(6, 7) = CD(7):
       LG(7, 1) = CD(2):      LG(7, 2) = CD(3):       LG(7, 3) = CD(4):    LG(7, 4) = CD(5):     LG(7, 5) = CD(5):   LG(7, 6) = CD(7):       LG(7, 7) = CD(8):
       LG(8, 1) = CD(2):      LG(8, 2) = CD(3):       LG(8, 3) = CD(4):    LG(8, 4) = CD(5):     LG(8, 5) = CD(6):   LG(8, 6) = CD(8):       LG(8, 7) = CD(9):
       LG(9, 1) = CD(3):      LG(9, 2) = CD(3):       LG(9, 3) = CD(5):    LG(9, 4) = CD(6):     LG(9, 5) = CD(7):   LG(9, 6) = CD(9):       LG(9, 7) = CD(10):
       LG(10, 1) = CD(3):     LG(10, 2) = CD(4):      LG(10, 3) = CD(5):   LG(10, 4) = CD(7):    LG(10, 5) = CD(8):  LG(10, 6) = CD(10):     LG(10, 7) = CD(11):
       LG(11, 1) = CD(3):     LG(11, 2) = CD(4):      LG(11, 3) = CD(6):   LG(11, 4) = CD(7):    LG(11, 5) = CD(9):  LG(11, 6) = CD(11):     LG(11, 7) = CD(12):
       LG(12, 1) = CD(3):     LG(12, 2) = CD(4):      LG(12, 3) = CD(6):   LG(12, 4) = CD(8):    LG(12, 5) = CD(10): LG(12, 6) = CD(12):     LG(12, 7) = CD(13):
       LG(13, 1) = CD(4):     LG(13, 2) = CD(5):      LG(13, 3) = CD(7):   LG(13, 4) = CD(9):    LG(13, 5) = CD(11): LG(13, 6) = CD(13):     LG(13, 7) = CD(14):
       LG(14, 1) = CD(4):     LG(14, 2) = CD(5):      LG(14, 3) = CD(7):   LG(14, 4) = CD(9):    LG(14, 5) = CD(12): LG(14, 6) = CD(14):     LG(14, 7) = CD(15):
       LG(15, 1) = CD(4):     LG(15, 2) = CD(5):      LG(15, 3) = CD(8):   LG(15, 4) = CD(10):   LG(15, 5) = CD(13): LG(15, 6) = CD(15):     LG(15, 7) = CD(16)
    

       With ThisDocument.Tables(1)
       itype = .Cell(6, 2).Range.Text
       SamP = .Cell(7, 4).Range.Text

       '，NORMALTIGHTENED，。
       
        If itype Like "*Pre-shipment Inspection*" = False And itype Like "*PSI*" = False And _
          itype Like "*During Production Inspection*" = False And itype Like "*DPI*" = False And _
          itype Like "*Re-Inspection*" = False And itype Like "*Re-inspection*" = False Then MsgBox "Inspection Type" & UText(34920, 26684, 36755, 20837, 19981, 31526, 35201, 27714) & "，" & UText(35831, 26816, 26597, 26159, 21542, 27491, 30830): Exit Sub
          
       If itype Like "*Pre-shipment Inspection*" = True Or itype Like "*PSI*" = True Or _
          itype Like "*During Production Inspection*" = True Or itype Like "*DPI*" = True Then
          
            If SamP Like "*Normal*" = False And SamP Like "*Tightened*" = False Then MsgBox "Information " & UText(34920, 25152, 22635, 20889, 30340, 25277, 26679, 26041, 26696) & "，" & UText(19981, 26159) & "Normal" & UText(20063, 19981, 26159) & "Tightened," & UText(26412, 31243, 24207, 26080, 27861, 22788, 29702) & "," & UText(35831, 26816, 26597), vbCritical: Exit Sub
       
            If SamP Like "*Tightened*" = True Then
               AnS = MsgBox(UText(26412, 27425, 39564, 36135, 20026, 27491, 24120, 26816, 39564) & "，" & UText(20294) & "Sampling Plan" & UText(19981, 26159, 36755, 20837) & "Normal，" & UText(35831, 30830, 35748, 26159, 21542, 26377, 38169) & "。" & UText(22914, 26524, 38656, 35201, 36864, 20986, 20462, 25913, 35831, 25353, 26159) & "，" & UText(22914, 26524, 19981, 36864, 20986, 20462, 25913, 32780, 32487, 32493, 36816, 34892) & "，" & UText(35831, 25353, 21542), vbYesNo)
               If AnS = vbYes Then Exit Sub
               Styp = 2                           ' translated comment
            End If
          
            If SamP Like "*Normal*" = True Then Styp = 1 ' translated comment
      End If
          
            
       If itype Like "*Re-Inspection*" = True Or itype Like "*Re-inspection*" = True Then
          If SamP Like "*Normal*" = False And SamP Like "*Tightened*" = False Then MsgBox "Information " & UText(34920, 25152, 22635, 20889, 30340, 25277, 26679, 26041, 26696) & "，" & UText(19981, 26159) & "Normal" & UText(20063, 19981, 26159) & "Tightened," & UText(26412, 31243, 24207, 26080, 27861, 22788, 29702) & "," & UText(35831, 26816, 26597), vbCritical: Exit Sub
               If SamP Like "*Normal*" = True Then
               AnS = MsgBox(UText(26412, 27425, 39564, 36135, 20026, 37325, 39564) & "，" & UText(20294) & "Sampling Plan" & UText(19981, 26159, 36755, 20837) & "Tightened，" & UText(35831, 30830, 35748, 26159, 21542, 26377, 38169) & "，" & UText(22914, 26524, 38656, 35201, 36864, 20986, 20462, 25913, 35831, 25353, 26159) & "，" & UText(22914, 26524, 19981, 36864, 20986, 20462, 25913, 32780, 32487, 32493, 36816, 34892, 35831, 25353, 21542), vbYesNo)
               If AnS = vbYes Then Exit Sub
               Styp = 1
            End If
          
            If SamP Like "*Tightened*" = True Then Styp = 2
      End If
    End With
       
    Select Case Styp
          
    Case 1   ' translated comment
       
       
       'AQL AC
       'AC(),AQL ， ，AQL 1.0,1.5,...
       'MV(),， Sample Size
       AC(0, 0) = CD(0)
       AC(0, 1) = CD(0):       AC(0, 2) = CD(0):   AC(0, 3) = CD(0):   AC(0, 4) = CD(0):   AC(0, 5) = CD(0):
       
       AC(1, 1) = 0:           AC(1, 2) = 0:       AC(1, 3) = 0:       AC(1, 4) = 0:       AC(1, 5) = 0:
       MV(1, 1) = CD(5):       MV(1, 2) = CD(4):   MV(1, 3) = CD(3):   MV(1, 4) = CD(2):
       AC(2, 1) = 0:           AC(2, 2) = 0:       AC(2, 3) = 0:       AC(2, 4) = 0:       AC(2, 5) = 0:
       MV(2, 1) = CD(5):       MV(2, 2) = CD(4):   MV(2, 3) = CD(3):                       MV(2, 5) = CD(1):
       AC(3, 1) = 0:           AC(3, 2) = 0:       AC(3, 3) = 0:       AC(3, 4) = 0:       AC(3, 5) = 1:
       MV(3, 1) = CD(5):       MV(3, 2) = CD(4):                       MV(3, 4) = CD(2):   MV(3, 5) = CD(4):
       AC(4, 1) = 0:           AC(4, 2) = 0:       AC(4, 3) = 0:       AC(4, 4) = 1:       AC(4, 5) = 1:
       MV(4, 1) = CD(5):                           MV(4, 3) = CD(3):   MV(4, 4) = CD(5):
       AC(5, 1) = 0:           AC(5, 2) = 0:       AC(5, 3) = 1:       AC(5, 4) = 1:       AC(5, 5) = 2:
                               MV(5, 2) = CD(4):   MV(5, 3) = CD(6):
       AC(6, 1) = 0:           AC(6, 2) = 1:       AC(6, 3) = 1:       AC(6, 4) = 2:       AC(6, 5) = 3:
       MV(6, 1) = CD(5):       MV(6, 2) = CD(7):
       AC(7, 1) = 1:           AC(7, 2) = 1:       AC(7, 3) = 2:       AC(7, 4) = 3:       AC(7, 5) = 5:
       MV(7, 1) = CD(8)
       AC(8, 1) = 1:           AC(8, 2) = 2:       AC(8, 3) = 3:       AC(8, 4) = 5:       AC(8, 5) = 7:
        
       AC(9, 1) = 2:           AC(9, 2) = 3:       AC(9, 3) = 5:       AC(9, 4) = 7:       AC(9, 5) = 10:
       
       AC(10, 1) = 3:          AC(10, 2) = 5:      AC(10, 3) = 7:      AC(10, 4) = 10:     AC(10, 5) = 14:
       
       AC(11, 1) = 5:          AC(11, 2) = 7:      AC(11, 3) = 10:     AC(11, 4) = 14:     AC(11, 5) = 21:
        
       AC(12, 1) = 7:          AC(12, 2) = 10:     AC(12, 3) = 14:     AC(12, 4) = 21:     AC(12, 5) = 21:
                                                                                           MV(12, 5) = CD(11):
       AC(13, 1) = 10:         AC(13, 2) = 14:     AC(13, 3) = 21:     AC(13, 4) = 21:     AC(13, 5) = 21:
                                                                       MV(13, 4) = CD(12): MV(13, 5) = CD(11):
       AC(14, 1) = 14:         AC(14, 2) = 21:     AC(14, 3) = 21:     AC(14, 4) = 21:     AC(14, 5) = 21:
                                                   MV(14, 3) = CD(13): MV(14, 4) = CD(12): MV(14, 5) = CD(11):
       AC(15, 1) = 21:         AC(15, 2) = 21:     AC(15, 3) = 21:     AC(15, 4) = 21:     AC(15, 5) = 21:
                               MV(15, 2) = CD(14): MV(15, 3) = CD(13): MV(15, 4) = CD(12): MV(15, 5) = CD(11):
       AC(16, 1) = 21:         AC(16, 2) = 21:     AC(16, 3) = 21:     AC(16, 4) = 21:     AC(16, 5) = 21:
       MV(16, 1) = CD(15):     MV(16, 2) = CD(14): MV(16, 3) = CD(13): MV(16, 4) = CD(12): MV(16, 5) = CD(11):
                            
   Case 2  ' translated comment
   
     
       'AQL AC
       'AC(),AQL ， ，AQL 1.0,1.5,...
       'MV(),， Sample Size
       AC(0, 0) = CD(0)
       AC(0, 1) = CD(0):       AC(0, 2) = CD(0):   AC(0, 3) = CD(0):   AC(0, 4) = CD(0):   AC(0, 5) = CD(0):
       
       AC(1, 1) = 0:           AC(1, 2) = 0:       AC(1, 3) = 0:       AC(1, 4) = 0:       AC(1, 5) = 0:
       MV(1, 1) = CD(6):       MV(1, 2) = CD(5):   MV(1, 3) = CD(4):   MV(1, 4) = CD(3):   MV(1, 5) = CD(2):
       AC(2, 1) = 0:           AC(2, 2) = 0:       AC(2, 3) = 0:       AC(2, 4) = 0:       AC(2, 5) = 0:
       MV(2, 1) = CD(6):       MV(2, 2) = CD(5):   MV(2, 3) = CD(4):   MV(2, 4) = CD(3):
       AC(3, 1) = 0:           AC(3, 2) = 0:       AC(3, 3) = 0:       AC(3, 4) = 0:       AC(3, 5) = 1:
       MV(3, 1) = CD(6):       MV(3, 2) = CD(5):   MV(3, 3) = CD(4):                       MV(3, 5) = CD(5):
       AC(4, 1) = 0:           AC(4, 2) = 0:       AC(4, 3) = 0:       AC(4, 4) = 1:       AC(4, 5) = 1:
       MV(4, 1) = CD(6):       MV(4, 2) = CD(5):                       MV(4, 4) = CD(6):   MV(4, 5) = CD(5):
       AC(5, 1) = 0:           AC(5, 2) = 0:       AC(5, 3) = 1:       AC(5, 4) = 1:       AC(5, 5) = 1:
       MV(5, 1) = CD(6):                           MV(5, 3) = CD(7):   MV(5, 4) = CD(6):
       AC(6, 1) = 0:           AC(6, 2) = 1:       AC(6, 3) = 1:       AC(6, 4) = 1:       AC(6, 5) = 2:
                               MV(6, 2) = CD(8):   MV(6, 3) = CD(7):
       AC(7, 1) = 1:           AC(7, 2) = 1:       AC(7, 3) = 1:       AC(7, 4) = 2:       AC(7, 5) = 3:
       MV(7, 1) = CD(9):       MV(7, 2) = CD(8):
       AC(8, 1) = 1:           AC(8, 2) = 1:       AC(8, 3) = 2:       AC(8, 4) = 3:       AC(8, 5) = 5:
       MV(8, 1) = CD(9):
       AC(9, 1) = 1:           AC(9, 2) = 2:       AC(9, 3) = 3:       AC(9, 4) = 5:       AC(9, 5) = 8:
        
       AC(10, 1) = 2:          AC(10, 2) = 3:      AC(10, 3) = 5:      AC(10, 4) = 8:      AC(10, 5) = 12:
       
       AC(11, 1) = 3:          AC(11, 2) = 5:      AC(11, 3) = 8:     AC(11, 4) = 12:      AC(11, 5) = 18:
        
       AC(12, 1) = 5:          AC(12, 2) = 8:      AC(12, 3) = 12:     AC(12, 4) = 18:     AC(12, 5) = 18:
                                                                                           MV(12, 5) = CD(11):
       AC(13, 1) = 8:          AC(13, 2) = 12:     AC(13, 3) = 18:     AC(13, 4) = 18:     AC(13, 5) = 18:
                                                                       MV(13, 4) = CD(12): MV(13, 5) = CD(11):
       AC(14, 1) = 12:         AC(14, 2) = 18:     AC(14, 3) = 18:     AC(14, 4) = 18:     AC(14, 5) = 18:
                                                   MV(14, 3) = CD(13): MV(14, 4) = CD(12): MV(14, 5) = CD(11):
       AC(15, 1) = 18:         AC(15, 2) = 18:     AC(15, 3) = 18:     AC(15, 4) = 18:     AC(15, 5) = 18:
                               MV(15, 2) = CD(14): MV(15, 3) = CD(13): MV(15, 4) = CD(12): MV(15, 5) = CD(11):
       AC(16, 1) = 18:         AC(16, 2) = 18:     AC(16, 3) = 18:     AC(16, 4) = 18:     AC(16, 5) = 18:
       MV(16, 1) = CD(15):     MV(16, 2) = CD(14): MV(16, 3) = CD(13): MV(16, 4) = CD(12): MV(16, 5) = CD(11):
   
   
   
   End Select
                            
       
       With ThisDocument.Tables(1)
      
         ' translated comment
       If .Cell(8, 4).Range.Text Like "*G-III*" = True Or .Cell(8, 4).Range.Text Like "*G-111*" = True Then
       lvl = 7
       ElseIf .Cell(8, 4).Range.Text Like "*G-II*" = True Or .Cell(8, 4).Range.Text Like "*G-11*" = True Then
       lvl = 6
       ElseIf .Cell(8, 4).Range.Text Like "*G-I*" = True Or .Cell(8, 4).Range.Text Like "*G-1*" = True Then
       lvl = 5
       ElseIf .Cell(8, 4).Range.Text Like "*S-4*" = True Or .Cell(8, 4).Range.Text Like "*s-4*" = True Then
       lvl = 4
       ElseIf .Cell(8, 4).Range.Text Like "*S-3*" = True Or .Cell(8, 4).Range.Text Like "*s-3*" = True Then
       lvl = 3
       ElseIf .Cell(8, 4).Range.Text Like "*S-2*" = True Or .Cell(8, 4).Range.Text Like "*s-2*" = True Then
       lvl = 2
       ElseIf .Cell(8, 4).Range.Text Like "*S-1*" = True Or .Cell(8, 4).Range.Text Like "*s-1*" = True Then
       lvl = 1
       Else
       MsgBox UText(26816, 26597, 27700, 24179) & "，" & UText(19981, 26159) & "S-1" & UText(21040) & "S-4，" & UText(20063, 19981, 26159) & "G-I，G-II" & UText(25110) & "G-III，" & UText(26080, 27861, 35745, 31639) & "，" & UText(35831, 26816, 26597, 36755, 20837, 26159, 21542, 31526, 21512, 35201, 27714, 25110, 25163, 24037, 35745, 31639): Exit Sub
       End If

       
       ' translated comment
       Acr = 0
       AFn = 0
       AMa = 0
       AFnMa = 0
       Ami = 0
       
       
       Aqlt = .Cell(10, 4).Range.Text
    
       Set rg = .Cell(10, 4).Range
       Wn = rg.words.count
       WL = Len(Aqlt)
      ' Pn = rg.Characters.Count
      
      
     
       'AQL,
       
       If Wn < 14 Or Aqlt Like "*Cr*" = False Or Aqlt Like "*Ma*" = False Or _
       Aqlt Like "*Mi*" = False Then MsgBox "Information " & UText(34920, 26684) & "，AQL" & UText(26410, 23436, 20840, 36755, 20837) & "，" & UText(25110, 36755, 20837, 26684, 24335, 19981, 31526, 35201, 27714) & "，" & UText(27491, 30830, 26684, 24335, 31867, 20284, 20110) & "Cr:0 Ma:1.0 Mi:4.0，" & UText(27880, 24847, 27599, 20010, 32570, 38519, 23383, 27597, 21518, 29992, 19968, 20010, 20882, 21495) & "，" & UText(32570, 38519, 25968, 23383, 21518) & "，" & UText(24517, 39035, 19988, 21482, 26377, 26377, 19968, 20010, 31354, 26684) & "，" & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892) & "！！", vbCritical: Exit Sub
           
       If WL <> 20 And WL <> 30 And WL <> 37 Then MsgBox "Information " & UText(34920, 26684) & "，AQL" & UText(36755, 20837, 26684, 24335, 19981, 31526, 35201, 27714) & "，" & UText(27491, 30830, 26684, 24335, 31867, 20284, 20110) & "Cr:0 Ma:1.0 Mi:4.0，" & UText(27880, 24847, 27599, 20010, 32570, 38519, 23383, 27597, 21518, 29992, 19968, 20010, 20882, 21495) & "，" & UText(32570, 38519, 25968, 23383, 21518) & "，" & UText(24517, 39035, 19988, 21482, 26377, 26377, 19968, 20010, 31354, 26684) & "，" & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892) & "！！", vbCritical: Exit Sub
       
       If Aqlt Like "*Fn*" = False Then
       
         If Wn = 14 Then
         Typ = 1                      'aql ，1， Cr,ma,mi
         Acr = Val(Mid(Aqlt, 4, 1))
         AMa = Val(Mid(Aqlt, 9, 3))    'Ma
         Ami = Val(Mid(Aqlt, 16, 3))
         ElseIf Wn <> 14 Then MsgBox "Information " & UText(34920, 26684) & "，AQL" & UText(36755, 20837, 26684, 24335, 19981, 31526, 35201, 27714) & "，" & UText(27491, 30830, 26684, 24335, 31867, 20284, 20110) & "Cr:0 Ma:1.0 Mi:4.0，" & UText(27880, 24847, 27599, 20010, 32570, 38519, 23383, 27597, 21518, 29992, 19968, 20010, 20882, 21495) & "，" & UText(32570, 38519, 25968, 23383, 21518) & "，" & UText(24517, 39035, 19988, 21482, 26377, 26377, 19968, 20010, 31354, 26684) & "，" & UText(19981, 35201, 26377, 22810, 20313, 22238, 36710, 21644, 31354, 26684) & "，" & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892) & "！！", vbCritical: Exit Sub
         End If
       End If
       If Aqlt Like "*Fn*" = True Then
          If Wn = 21 Then
          Typ = 2
          Acr = Val(Mid(Aqlt, 4, 1))
          AFn = Val(Mid(Aqlt, 9, 3))
          AFnMa = Val(Mid(Aqlt, 19, 3))
          Ami = Val(Mid(Aqlt, 26, 3))
          End If
          If Wn = 26 Then
          Typ = 3
          Acr = Val(Mid(Aqlt, 4, 1))
          AFn = Val(Mid(Aqlt, 9, 3))
          AMa = Val(Mid(Aqlt, 16, 3))
          AFnMa = Val(Mid(Aqlt, 26, 3))
          Ami = Val(Mid(Aqlt, 33, 3))
          End If
          If Wn <> 21 And Wn <> 26 Then MsgBox "Information " & UText(34920, 26684) & "，AQL" & UText(36755, 20837, 25968, 25454, 30340, 26684, 24335, 19981, 31526, 21512, 35201, 27714) & "，" & UText(27491, 30830, 26684, 24335, 31867, 20284, 20110) & "Cr:0 Fn:1.0 Fn+Ma:1.0 Mi:4.0," & UText(27880, 24847, 27599, 20010, 32570, 38519, 23383, 27597, 21518, 29992, 19968, 20010, 20882, 21495) & "，" & UText(32570, 38519, 25968, 23383, 21518) & "，" & UText(24517, 39035, 19988, 21482, 26377, 26377, 19968, 20010, 31354, 26684) & "," & UText(19981, 35201, 26377, 22810, 20313, 22238, 36710, 21644, 31354, 26684) & "，" & UText(35831, 20462, 25913, 21518, 20877, 36816, 34892) & "！！", vbCritical: Exit Sub
       End If

       'AQL
       If Acr <> 0 Then MsgBox UText(27880, 24847) & "：Information " & UText(34920, 26684) & "，Critical " & UText(30340) & "AQL" & UText(19981, 26159) & "0，" & UText(38656, 35201, 25163, 24037, 26597, 35810) & "Cr" & UText(30340) & "AQL" & UText(34920) & "。"
       
       AQL(0) = Acr
       AQL(1) = AFn
       AQL(2) = AMa
       AQL(3) = AFnMa
       AQL(4) = Ami

      'Sqty = InputBox(UText(35831, 36755, 20837) & "Shipment Quanty")
  
       SQty = Val(SQty)
       If SQty < Lots(1) Then MsgBox UText(25209, 37327, 25968, 22826, 23567) & "，" & UText(26412, 31243, 24207, 19981, 33021, 35745, 31639) & "，" & UText(35831, 25163, 24037, 35745, 31639): Exit Sub
       If SQty > Lots(15) Then MsgBox UText(25209, 37327, 25968, 22826, 22823) & "，" & UText(26412, 31243, 24207, 19981, 33021, 35745, 31639) & "，" & UText(35831, 25163, 24037, 35745, 31639): Exit Sub
    
     
  
        
        'LotSiZE
       For i = 1 To 15
       
       If SQty >= Lots(i) And SQty < Lots(i + 1) Then grp = i: GoTo 10
       
       Next i
       
    'AQLSampleSiz
      
10     Scode = LG(grp, lvl)
       For i = 1 To 16
       If Scode = CD(i) Then Hgrp = i: GoTo 20
       Next i
20    For k = 1 To 4  'k AQLMa,Fn,Fnma,Mi
       
        
        'AQL
        Cn = 0
     For i = 1 To 5
        If AQL(k) = AQ(i) Then Agrp = i: GoTo 30
        If AQL(k) <> AQ(i) And AQL(k) <> 0 Then Cn = Cn + 1 'AQL
        Next i
        If Cn = 5 Then MsgBox "Information " & UText(34920, 26684) & ",AQL" & UText(36755, 20837, 25968, 23383, 20013, 26377, 25968, 23383, 36229, 36807, 26412, 31243, 24207, 22788, 29702, 33539, 22260) & "，" & UText(26412, 31243, 24207, 21482, 22788, 29702) & "AQL：1.0，1.5，2.5，4.0，6.5" & UText(30340, 25968, 20540) & "！！！，" & UText(35831, 20462, 25913, 25110, 25163, 24037, 22788, 29702) & " ", vbCritical: Exit Sub
30     If MV(Hgrp, Agrp) > 0 Then
         If MV(Hgrp, Agrp) >= SQty Then SmpS(k) = SQty              'sample size
         If MV(Hgrp, Agrp) < SQty Then SmpS(k) = MV(Hgrp, Agrp)
       End If
       If MV(Hgrp, Agrp) = 0 Then SmpS(k) = CD(Hgrp)
    
       AAc(k) = AC(Hgrp, Agrp)                      'AQL
       
40    If SmpS(k) > Ass Then Ass = SmpS(k)            'Sample Sise
      
      If AQL(0) = 0 Then AAc(0) = 0           'AQL=0 AC0
        
     Next k
     

       'Sample size
       
       Set rg = .Cell(9, 4).Range
       ss = Val(rg.words(1))
       If Ass <> ss Then MsgBox "Information " & UText(34920, 26684) & "," & UText(35745, 31639, 24471, 21040, 30340) & "Sample Size" & UText(26159) & Ass & UText(20294, 34920, 26684, 20013, 22635, 20837, 30340, 21364, 26159) & ss & UText(23558, 34987, 26356, 25913, 26367, 25442) & "！", vbCritical
       rg.words(1) = Replace(ss, ss, Ass & " ")
      
    End With

  
   'workmanship
    
    Select Case Typ
    
    Case 1
    'defect allowed
    ADef(1) = AAc(0)      'cr
    ADef(2) = AAc(2)      'Ma
    ADef(3) = AAc(4)     'Mi
    
    TBName = "*Workmanship*"
    tb = ThisDocument.Tables.count - 1
    
    For X = tb To 1 Step -1
    With ThisDocument.Range.Tables(X)
        Coln = .Columns.count
      If Trim(.Cell(1, 1).Range.Text) Like TBName = True And Coln = 4 Then
        Rown = .Rows.count
        For i = 2 To Coln
      .Cell(Rown, i).Range.Text = ADef(i - 1)
       Next i
      End If
      If Trim(.Cell(1, 1).Range.Text) Like TBName = True And Coln > 4 Then
         ThisDocument.Range.Tables(X).Select
         ThisDocument.Range.Tables(X).Delete
         Selection.Delete
     
      End If
     End With
    
    Next X
   
     Case 2
     
    ADef(1) = AAc(0)      'cr
    ADef(2) = AAc(1)      'Fn
    ADef(3) = AAc(3)      'FnMa
    ADef(4) = AAc(4)     'Mi
    
    TBName = "*Workmanship*"
    tb = ThisDocument.Tables.count - 1
    
    For X = tb To 1 Step -1
    With ThisDocument.Range.Tables(X)
        Coln = .Columns.count
       If Trim(.Cell(1, 1).Range.Text) Like TBName = True And Coln = 5 Then
        Rown = .Rows.count
        For i = 2 To Coln
      .Cell(Rown, i).Range.Text = ADef(i - 1)
       Next i
      End If
      If Trim(.Cell(1, 1).Range.Text) Like TBName = True And Coln <> 5 Then
         ThisDocument.Range.Tables(X).Select
         ThisDocument.Range.Tables(X).Delete
         Selection.Delete
     
      End If
     End With
    
    Next X
   
     Case 3
     
    ADef(1) = AAc(0)      'cr
    ADef(2) = AAc(1)      'Fn
    ADef(3) = AAc(2)     'Ma
    ADef(4) = AAc(3)     'FnMA
    ADef(5) = AAc(4)     'Mi
    
    TBName = "*Workmanship*"
    tb = ThisDocument.Tables.count - 1
    
    For X = tb To 1 Step -1
    With ThisDocument.Range.Tables(X)
        Coln = .Columns.count
      If Trim(.Cell(1, 1).Range.Text) Like TBName = True And Coln = 6 Then
        Rown = .Rows.count
        For i = 2 To Coln
      .Cell(Rown, i).Range.Text = ADef(i - 1)
       Next i
      End If
      If Trim(.Cell(1, 1).Range.Text) Like TBName = True And Coln <> 6 Then
         ThisDocument.Range.Tables(X).Select
         ThisDocument.Range.Tables(X).Delete
         Selection.Delete
     
      End If
     End With
    
    Next X
  
  End Select
  
 
  
    
End Sub



Sub Makefolders()
' Quick function keys: ALT -1
' by wyn 20241216


Dim count, Url, words, rwindex, url1, fso
Dim sx1, sx2, p(200), S(10), urlt(500), mypath, Rn(200), checkf As String
Dim i, j, ch, qtyRow As Integer
Dim tb_sap_qty, tb_sap_wms, tb_sap_ost, tb_sap_pdw, tb_sap_smc, tb_sap_pc, tb_sap_plm, tb_sap_sm, tb_sap_pp  As String
Dim tb_no_qty, tb_no_wms, tb_no_ost, tb_no_pdw, tb_no_smc, tb_no_pc, tb_no_plm, tb_no_sm, tb_no_pp  As String
Dim iW As String
Dim Urlx, Urlx1, Urlx2, Urlx3 As String
Dim Rown, CLn As Integer
Dim Cn, Dn As Integer
Dim DefName() As String
Dim DefNo() As Integer
Dim Dw As String
 exitf = 0

 Set fso = CreateObject("Scripting.FileSystemObject")
 
 sx1 = InputBox("Input the path you want to save the photos under it, su as: " & CHR(10) & CHR(10) & UText(35831, 36755, 20837, 35201, 29983, 25104, 25991, 20214, 22841, 30340, 29238, 25991, 20214, 22841, 36335, 24452) & "," & UText(27604, 22914) & ":", "Notice", "D:\InspectionPhotos")

 If sx1 = "" Then Exit Sub
 
 If fso.FolderExists(sx1) = False Then MkDir (sx1)
 
 Urlx = sx1 & "\" & Format(Date, "yyyy mmm d") & " report Photo\"

 If fso.FolderExists(Urlx) = False Then
     MkDir (Urlx)
 Else
    If MsgBox(Urlx & " folder existed, to create new folder will delete all existing files under the folder. Do you want to continue? " & CHR(10) & CHR(10) & _
    Urlx & " " & UText(30340, 25991, 20214, 22841, 24050, 32463, 23384, 22312) & "，" & UText(21019, 24314, 26032, 25991, 20214, 22841, 23558, 21024, 38500, 21407, 26377, 35813, 25991, 20214, 22841, 19979, 25152, 26377, 25991, 20214) & "，" & UText(26159, 21542, 32487, 32493, 21019, 24314, 26032, 25991, 20214, 22841) & "？", vbYesNoCancel + vbInformation) = vbYes Then
      Set fs = CreateObject("Scripting.FileSystemObject")
      fs.DeleteFolder Urlx & "*"
  
    Else
      Exit Sub
    End If
   ' exitf = 1 Else Exit Sub
 End If

' urlx = url
' If fso.FolderExists(urlx) = False Then MkDir (urlx)
 
 p(0) = "00. First page photos(No naming required)\"
 p(1) = "01. Quantity & Warehouse & Storage(No naming required)\"
 p(2) = "02. Workmanship（No naming required)\"
 p(3) = "03. On Site Test（naming required)\"
 p(4) = "04. Product Dimension & Weight（no naming required)\"
 p(5) = "05. Style & Material & Construction(No naming required)\"
 p(6) = "06. Product Color（No naming required)\"
 p(7) = "07.  Product Label & Marking（No naming required)\"
 p(8) = "08. Shipping Mark（No naming required)\"
 p(9) = "09. Packing & Packaging（No naming required)\"
 p(10) = "10. Other Note（No naming required)\"
 p(11) = "11. Other photos（Naming or no naming required)\"

 For i = 0 To 11
  Urlx1 = Urlx & p(i)
  If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
Next i

 
 '
   TBName = "*1. Quantity*"
   tb = ThisDocument.Tables.count - 1
   For X = 1 To tb
   With ThisDocument.Range.Tables(X)
        If Trim(.Cell(1, 1).Range.Text) Like TBName = True Then tbn = X
     End With
  Next X
  
   With ThisDocument.Range.Tables(tbn)
      qtyRow = .Rows.count
       Itms = qtyRow - 5 '，workmanship
     ReDim Ptx(1 To Itms)
     ReDim Itx(1 To Itms)
     
     
     
     For i = 5 To qtyRow - 1
       Itx(i - 4) = Mid(.Cell(i, 2).Range.Text, 1, Len(.Cell(i, 2).Range.Text) - 2)
       If Itx(i - 4) = "" Then
          j = j + 1
          Itx(i - 4) = "No name" & CStr(j)
       End If
     Next i
   End With
      
   If j > 0 Then MsgBox "The sub-folder will generate and name by item No., However, there are " & j & "lines in 1.Quantity table without input item number, " & _
  "the sub-folders may be with wrong item number, " & "please input tiem number into 1.Qauntity table and run program again " & _
  " or rename all sub-folders " & CHR(10) & _
    UText(25991, 20214, 22841, 19979, 30340, 23376, 25991, 20214, 22841, 23558, 20197, 27454, 21495, 21629, 21517) & "，" & UText(20294, 30446, 21069, 21457, 29616, 22312) & "1.Quantity" & UText(34920, 20013, 26377) & j & UText(34892, 27809, 26377, 36755, 20837) & "item" & UText(21495) & "，" & UText(29983, 25104, 30340, 23376, 25991, 20214, 22841, 27454, 21495, 30340, 21517, 31216, 21487, 33021, 19981, 23545) & _
  UText(35831, 22312, 25253, 21578) & "1.Quantity" & UText(34920, 26684, 20013, 37325, 26032, 36755, 20837, 27454, 21495, 20877, 36816, 34892, 31243, 24207, 25110, 32773, 23545, 24050, 29983, 25104, 30340, 25152, 26377, 23376, 25991, 20214, 22841, 21152, 20837, 27454, 21495), vbCritical
   
   
   
          '=============
 Dim ComItx() As String
 ReDim ComItx(1 To Itms)

 Dim ComItms As Integer
      
   
      N = 0
      Set dictItem = CreateObject("Scripting.Dictionary")
      For i = 1 To Itms
        If Not dictItem.Exists(Itx(i)) Then
            N = N + 1
            dictItem.add Itx(i), Nothing
            ComItx(N) = Itx(i)
        End If
    Next i
    ComItms = dictItem.count
    For i = 1 To dictItem.count
      Itx(i) = ComItx(i)
    Next i
   Itms = ComItms
   '==============

   
  '----
  
  
  tbn = getTableNumber(2) + 1
  With ThisDocument.Tables(tbn)
    Rown = .Rows.count
    CLn = .Columns.count
    ReDim DefName(1 To Itms, 1 To Rown)
    ReDim DefNo(1 To Itms)
    Cn = 0
    For i = 2 To Rown - 2
      If .Rows(i).Range.Cells.count = 1 Then
         Cn = Cn + 1
         Dn = 0
      Else
        
        If .Cell(i, CLn - 3).Range.Fields.count = 0 Then
           Dw = Mid(.Cell(i, CLn - 3).Range.Text, 1, Len(.Cell(i, CLn - 3).Range.Text) - 2)
           If Val(.Cell(i, CLn - 2).Range.Text) > 0 Then
              Dn = Dn + 1
              DefName(Cn, Dn) = Dw & "-" & "Critical"
              DefNo(Cn) = Dn
           ElseIf Val(.Cell(i, CLn - 1).Range.Text) > 0 Then
              Dn = Dn + 1
              DefName(Cn, Dn) = Dw & "-" & "Major"
              DefNo(Cn) = Dn
           ElseIf Val(.Cell(i, CLn).Range.Text) > 0 Then
             Dn = Dn + 1
             DefName(Cn, Dn) = Dw & "-" & "Minor"
             DefNo(Cn) = Dn
           End If
        Else
          Dn = 0
          i = i + 1
        End If
      End If
    Next i
  End With
  

  
  k = 2
  
  
   For i = 1 To Itms
       If i < 10 Then iW = "0" & CStr(i) Else iW = CStr(i)
         Urlx1 = Urlx & p(k) & iW & " Item " & Itx(i)
        If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
       'workmanshipitemdefect
       For j = 1 To DefNo(i)
          If j < 10 Then iW = "0" & CStr(j) Else iW = CStr(j)
             If InStr(1, DefName(i, j), "/") > 0 Then DefName(i, j) = Replace(DefName(i, j), "/", "-")
             Urlx2 = Urlx1 & "\" & iW & " " & DefName(i, j)
            If fso.FolderExists(Urlx2) = False Then MkDir (Urlx2)
       
       Next j
   Next i
  
   
  
  '-----------
   For k = 3 To 9
   For i = 1 To Itms
       If i < 10 Then iW = "0" & CStr(i) Else iW = CStr(i)
     Urlx1 = Urlx & p(k) & iW & " Item " & Itx(i)
       If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
   Next i
   Next k
    Urlx1 = Urlx & p(4) & "Big Photo"
     If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
    For i = 1 To Itms
        If i < 10 Then iW = "0" & CStr(i) Else iW = CStr(i)
        Urlx1 = Urlx & p(11) & iW & " Item " & Itx(i)
       If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
   Next i
   
   


  
  
  tb_sap_qty = "Special Attention Point for Quantity"
  tb_sap_qty = LCase(tb_sap_qty)
  tb_sap_wms = "Special Attention Point for Workmanship"
  tb_sap_wms = LCase(tb_sap_wms)
  tb_sap_ost = "Special Attention Point for On Site Test"
  tb_sap_ost = LCase(tb_sap_ost)
  tb_sap_pdw = "Special Attention Point for Product Dimension & Weight"
  tb_sap_pdw = LCase(tb_sap_pdw)
  tb_sap_smc = "Special Attention Point for Style / Material / Construction"
  tb_sap_smc = LCase(tb_sap_smc)
  tb_sap_pc = "Special Attention Point for Product Color"
  tb_sap_pc = LCase(tb_sap_pc)
  tb_sap_plm = "Special Attention Point for Product Label / Marking"
  tb_sap_plm = LCase(tb_sap_plm)
  tb_sap_sm = "Special Attention Point for Shipping Mark"
  tb_sap_sm = LCase(tb_sap_sm)
  tb_sap_pp = "Special Attention Point for Packing / Packaging"
  tb_sap_pp = LCase(tb_sap_pp)
  '----
  tb_no_qty = "Reference Note for Quantity"
  tb_no_qty = LCase(tb_no_qty)
  tb_no_wms = "Reference Note for Workmanship"
  tb_no_wms = LCase(tb_no_wms)
  tb_no_ost = "Reference Note for On Site Test"
  tb_no_ost = LCase(tb_no_ost)
  tb_no_pdw = "Reference Note for Product Dimension & Weight"
  tb_no_pdw = LCase(tb_no_pdw)
  tb_no_smc = "Reference Note for Style / Material / Construction"
  tb_no_smc = LCase(tb_no_smc)
  tb_no_pc = "Reference Note for Product Color"
  tb_no_pc = LCase(tb_no_pc)
  tb_no_plm = "Reference Note for Product Label / Marking"
  tb_no_plm = LCase(tb_no_plm)
  tb_no_sm = "Reference Note for Shipping Mark"
  tb_no_sm = LCase(tb_no_sm)
  tb_no_pp = "Reference Note for Packing / Packaging"
  tb_no_pp = LCase(tb_no_pp)
  
  '===========================
  tb = ThisDocument.Tables.count - 1
  For X = 1 To tb
     With ThisDocument.Tables(X)
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_sap_qty & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(1) & "Special Attention Point"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_no_qty & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(1) & "Reference Note"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       '------------------------------
        If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_sap_wms & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(2) & "Special Attention Point"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_no_wms & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(2) & "Reference Note"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       '------------------------------
        If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_sap_ost & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(3) & "Special Attention Point"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_no_ost & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(3) & "Reference Note"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       '------------------------------
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_sap_pdw & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(4) & "Special Attention Point"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_no_pdw & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(4) & "Reference Note"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       '------------------------------
        If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_sap_smc & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(5) & "Special Attention Point"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_no_smc & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(5) & "Reference Note"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       '------------------------------
        If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_sap_pc & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(6) & "Special Attention Point"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_no_pc & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(6) & "Reference Note"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       '------------------------------
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_sap_plm & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(7) & "Special Attention Point"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_no_plm & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(7) & "Reference Note"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       '------------------------------
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_sap_sm & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(8) & "Special Attention Point"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_no_sm & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(8) & "Reference Note"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       '------------------------------
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_sap_pp & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(9) & "Special Attention Point"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       If LCase(.Cell(1, 1).Range.Text) Like "*" & tb_no_pp & "*" = True Then
        If .Rows.count > 1 And Len(.Cell(2, 2).Range.Text) > 2 Then
             Urlx1 = Urlx & p(9) & "Reference Note"
             If fso.FolderExists(Urlx1) = False Then MkDir (Urlx1)
        End If
        GoTo 100
       End If
       '------------------------------
     End With
100  Next X

Shell "Explorer.exe" & " " & Urlx, vbNormalFocus

End Sub

Sub Pastephotos()
'Quick function keys: CTRL-1
'CTRL-1,by wyn 20230605

Dim i, N, i2, i2P, i3, I4(1000), i7, i8, tts, ttp, Rown, PicNP(200), TpicNP, NP As Integer
Dim S, ss2, sp, PhotoName1, PhotoPath(200), PhotoPathP(200), CurrentPhotoPath As String
Dim Show_Photo_Name As String
Dim PhotoName(1000) As String
Dim PhotoNameN(200, 500) As String
Dim skipnum, index, ii, j, X, rnp, cnp, Jia, JiaP, FN, SF(10) As Integer
Dim picN(1000) As Integer
Dim pic_Number As Integer
Dim tb_sap_qty, tb_sap_wms, tb_sap_ost, tb_sap_pdw, tb_sap_smc, tb_sap_pc, tb_sap_plm, tb_sap_sm, tb_sap_pp  As String
Dim tb_no_qty, tb_no_wms, tb_no_ost, tb_no_pdw, tb_no_smc, tb_no_pc, tb_no_plm, tb_no_sm, tb_no_pp  As String
Dim tb_sap(11), tb_no(11), Ltb_sap(11), Ltb_no(11), tb_tt(12) As String
Dim tempnum As String
Dim msg As String
Dim tt As Integer
Dim TbNumber As Integer
Dim Tn, BG As Integer
Dim Path, mypath As String
Dim Sap_Note_type As String
Dim myrange, rg As Range
Dim fd As FileDialog
Dim R(200), cs As Integer
Dim BigTbN, BigPhotoN As Integer
Dim folder As Object
Dim pic As InlineShape

      Set WD = ThisDocument
      For Each tbl In WD.Tables
         If tbl.Columns.count = 2 Then
            If Len(Mid(Trim(tbl.Cell(1, 1).Range.Text), 1, Len(Trim(tbl.Cell(1, 1).Range.Text)) - 2)) < 2 Or _
            Left(Trim(tbl.Cell(1, 1).Range.Text), 4) = "Item" Then
               For Each Cel In tbl.Range.Cells
                 For Each pic In tbl.Range.InlineShapes
                    pic.Delete
                Next pic
           
             Next Cel
           End If
         End If
      Next
      
       ' For Each pic In wd.InlineShapes
           'pic.Delete
       ' Next
'---------------------
  tb_tt(1) = "1. Quantity"
  tb_tt(2) = "2. Workmanship"
  tb_tt(3) = "3. On Site Test"
  tb_tt(4) = "4. Product Dimension & Weight"
  tb_tt(5) = "5. Style / Material / Construction"
  tb_tt(6) = "6. Product Color"
  tb_tt(7) = "7. Product Label / Marking"
  tb_tt(8) = "8. Shipping Mark"
  tb_tt(9) = "9. Packing / Packaging"
  tb_tt(10) = "10. Other Note"
  tb_tt(11) = "11. Other Photos"
 
          ss2 = InputBox("Input 0 or 1 and click “Confirm” key for pasting photos." & CHR(10) & _
            "   0： Exact photo name was not required for other photos, will be named with serial number." & CHR(10) & _
            "   1： Exact photo name was required, will remain the name you input." & CHR(10) & _
            CHR(10) & _
             UText(35831, 36755, 20837, 25554, 20837, 22270, 29255, 30340, 26041, 24335) & "：(" & UText(21482, 36755, 20837) & "0 " & UText(25110) & " 1 " & UText(30340, 25968, 25454) & ")，" & UText(20877, 28857, 30830, 23450, 38190) & CHR(10) & _
            "   0：Other Photo" & UText(22270, 29255, 19981, 38656, 35201, 20934, 30830, 21629, 21517) & "，" & UText(23558, 29992, 39034, 24207, 21495, 21629, 21517, 22270, 29255) & CHR(10) & _
            "   1：Other Photo" & UText(22270, 29255, 38656, 35201, 20934, 30830, 21629, 21517) & "，" & UText(23558, 20445, 30041, 20320, 36755, 20837, 30340, 22270, 29255, 21517) & "。" & CHR(10), "NOTICE", 0)


             If ss2 <> "0" And ss2 <> "1" Then MsgBox "Your inputting was not 0 or 1, the program exit. please run the program again and input 0 or 1." & _
             CHR(13) & UText(20320, 36755, 20837, 30340, 19981, 26159) & "0" & UText(25110) & "1," & UText(31243, 24207, 23558, 36864, 20986) & "，" & UText(35831, 37325, 26032, 36816, 34892, 24182, 36755, 20837) & "0" & UText(25110) & "1": Exit Sub

2    mypath = ThisDocument.Path
     ChDrive Split(mypath, "\")(0)
     ChDir mypath

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
     Dim AA As Variant
     With fd
         If .Show = -1 Then
             For Each AA In .SelectedItems
               S = AA & "\"
             Next AA
         Else
            Exit Sub
         End If
     End With
     Set fd = Nothing
     PhotoPath(0) = S    ' translated comment
     
     If Dir(PhotoPath(0), vbDirectory) = "" Or S = "" Then
        MsgBox "The folder is not existed!", vbCritical, "Warning"
        Exit Sub
     End If
   
   s2 = 1
   i8 = 1
   i2 = 1
   '----
  CurrentPhotoPath = Dir(S, vbDirectory)
  Do While CurrentPhotoPath <> ""
     If CurrentPhotoPath <> "." And CurrentPhotoPath <> ".." And (GetAttr(S & CurrentPhotoPath) And vbDirectory) = vbDirectory Then
       Kw = Trim(CurrentPhotoPath)
       Kw = Mid(Kw, 1, 1)
   
       If Not IsNumeric(Kw) Or Val(CurrentPhotoPath) > 11 Then MsgBox _
        "photo folders were not numbered or it's number larger than 11, " & _
       "please check. You can press “alt-1” to re-generate the 11 folders." & CHR(10) & _
         UText(23384, 22312, 26410, 32534, 21495, 25110, 32534, 21495, 22823, 20110) & "11" & UText(30340, 25991, 20214, 22841) & "，" & UText(35831, 26816, 26597) & "，" & UText(25110, 25353) & "“alt—1”" & UText(37325, 26032, 36816, 34892, 31243, 24207, 29983, 25104) & "11" & UText(20010, 25991, 20214, 22841) & "！！！", vbCritical: Exit Sub
       i2 = Val(CurrentPhotoPath)
       I4(i2) = i2
       PhotoPath(i2) = CurrentPhotoPath
         FN = FN + 1
     End If
      CurrentPhotoPath = Dir
    Loop
  Jia = FN ' translated comment
    If FN < 11 Then MsgBox "The photo folders was smaller than 12, the correct number of folder shall be 00 to 11,total 12 folders,please check, You can press “alt-1” to re-generate the 12 folders." & CHR(10) & _
         UText(25991, 20214, 22841, 23569, 20110) & "12" & UText(20010) & "，" & UText(27491, 30830, 25991, 20214, 22841, 32534, 21495, 24212, 26159) & "00" & UText(33267) & "11" & UText(20849) & "12" & UText(20010) & "，" & UText(35831, 26816, 26597, 25110, 25353) & "“alt—1”" & UText(36816, 34892, 31243, 24207, 37325, 26032, 29983, 25104) & "12" & UText(20010, 25991, 20214, 22841) & "！！！", vbCritical: Exit Sub
  '-------
    tt = 0  ' translated comment
    
     PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
      If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
      For i3 = 1 To picN(tt)
         PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
      If WD.Tables(2).Columns.count <> 2 Then
           WD.Tables(1).Cell(WD.Tables(1).Rows.count, 1).Select
           Selection.MoveDown
            Selection.TypeParagraph
            WD.Tables.add Range:=Selection.Range, NumRows:=1, NumColumns:=2, _
        DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
            WD.Tables(2).Select
             With Selection.Tables(1)
                If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
                 End If
              End With
              Call Setborders025
      Else
          If WD.Tables(2).Rows.count > 1 Then
            For i = WD.Tables(2).Rows.count To 2 Step -1
              WD.Tables(2).Rows(i).Delete
            Next i
          End If
          WD.Tables(2).Select
      End If
  
              If picN(tt) = 0 Then
               msg = msg + PhotoPath(tt) & "has not any photo, it should have two photos normally." & CHR(13) & PhotoPath(tt) & UText(27809, 26377, 22270, 29255) & "," & UText(27491, 24120, 35201, 27714, 35201, 26377, 20004, 24352, 22270, 29255) & CHR(13)
       
            ElseIf picN(tt) = 1 Then
                 msg = msg + PhotoPath(tt) & " has only one photo, it should have two photos normally." & CHR(13) & PhotoPath(tt) & UText(22270, 29255, 21482, 26377, 19968, 24352) & "，" & UText(27491, 24120, 35201, 27714, 35201, 26377, 20004, 24352, 22270, 29255) & CHR(13)
                 With Selection.Tables(1)
                    .Cell(1, 1).Range.InlineShapes.AddPicture FileName:= _
                      PhotoNameN(tt, 1), LinkToFile:=False, _
                      SaveWithDocument:=True
                     .Cell(1, 1).Range.InlineShapes(1).Height = 170
                     .Cell(1, 1).Range.InlineShapes(1).Width = 236
                     .Cell(1, 2).Range.Text = "NA"
                 End With
             Else
                
                   With Selection.Tables(1)
                    If picN(tt) > 2 Then
                      .Rows(1).Select
                      Selection.InsertRowsBelow Int((picN(tt) - 2) / 2 + 0.5)
                    End If
                     For i = 1 To picN(tt)
                           If i Mod 2 = 0 Then
                               cnp = 2
                               rnp = Int(i / 2)
                           Else
                              cnp = 1
                              rnp = Int(i / 2 + 0.5)
                           End If
                          .Cell(rnp, cnp).Range.InlineShapes.AddPicture FileName:= _
                          PhotoNameN(tt, i), LinkToFile:=False, _
                          SaveWithDocument:=True
                       If .Cell(rnp, cnp).Range.InlineShapes.count > 0 Then
                         .Cell(rnp, cnp).Range.InlineShapes(1).Height = 170
                         .Cell(rnp, cnp).Range.InlineShapes(1).Width = 236
                        End If
                     Next i
                   End With
            End If
'-------------------------------------------------------
5    tt = 1  'Quantity/warehouse
     picN(tt) = 0
     PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
     If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
      If picN(tt) = 0 Then
         msg = msg + PhotoPath(tt) & "has not any photo." & CHR(13) & PhotoPath(tt) & UText(27809, 26377, 22270, 29255) & CHR(13)
        GoTo 12
      End If
      If picN(tt) >= 10 Then '10
          i = picN(tt)
          i5 = 1
         For X = 1 To i
           PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
          If InStr(1, PhotoName1, " ") > 0 Then
             I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
           Else
              I4(X) = Val(PhotoName1)
           End If
         Next X
         Call PhotoSort(i, I4, i5, PhotoName)
      End If
      For i3 = 1 To picN(tt)
        PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
      tb = ThisDocument.Tables.count - 1
          For X = 1 To tb
            If ThisDocument.Tables(X).Cell(1, 1).Range.Text Like "*1. Quantity*" = True Then
              Tn = X
              GoTo 10
            End If
          Next X
      GoTo 15
10  Call GeneratePhotoTable(Tn)
            TbNumber = Tn + 1
            pic_Number = picN(tt)
           ' rnp = 1
            Show_Photo_Name = "0"
            Call PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)
            
12   '---
15   Path = S & PhotoPath(tt) & "\"
     Sap_Note_type = "Special Attention Point"
     Call Sap_photo(tt, Path, Sap_Note_type, msg)
     Sap_Note_type = "Reference Note"
     Call Note_photo(tt, Path, Sap_Note_type, msg)
   
        ' -------------workmanship
         tt = 2
         tb = WD.Tables.count - 1
         For X = 1 To tb
            With WD.Tables(X)
              If .Cell(1, 1).Range.Text Like "*Defect Photos*" = True And .Range.Rows.count = 1 Then
                Tn = X
                GoTo 20
              End If
            End With
         Next X
         msg = msg + "The table name as Defect Photos for locating defects photo position was not found at report, the defect photo was not pasted" & _
         CHR(10) & UText(25253, 21578, 20013, 26410, 21457, 29616, 29992, 20110, 23450, 20301) & "defect" & UText(22270, 29255, 20301, 32622, 21517, 23383, 20026) & "Defect Photos" & UText(30340, 34920, 26684) & "，" & UText(22240, 32780) & "defect" & UText(22270, 29255, 26410, 36148, 19978) & CHR(10)
         GoTo 25
     
20      Call GeneratePhotoTable(Tn)
    
   
   
     PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
     picN(tt) = 0
     If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
   
    If picN(tt) = 0 Then GoTo 22
     If picN(tt) >= 10 Then '10
        i = picN(tt)
        i5 = 1
         For X = 1 To i
           PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
           If InStr(1, PhotoName1, " ") > 0 Then
             I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
           Else
              I4(X) = Val(PhotoName1)
           End If
         Next X
        Call PhotoSort(i, I4, i5, PhotoName)
     End If
      For i3 = 1 To picN(tt)
        PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
         
            Show_Photo_Name = "1"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            'rnp = 1
            Call PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)

22          If picN(tt) > 0 Then
               WD.Tables(TbNumber).Rows(rnp + 1).Select
               Selection.InsertRowsBelow 2
            End If
     
            Show_Photo_Name = "1"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
  
             Call PasteAutoNameItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)
           
           ' Call PasteItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)
 
25  Path = S & PhotoPath(tt) & "\"
   Sap_Note_type = "Special Attention Point"
   Call Sap_photo(tt, Path, Sap_Note_type, msg)
   
   Sap_Note_type = "Reference Note"
   Call Note_photo(tt, Path, Sap_Note_type, msg)
      
    '-------onsite test
    tt = 3
    
     tb = ThisDocument.Tables.count - 1
          For X = 1 To tb
            If ThisDocument.Tables(X).Cell(1, 1).Range.Text Like "*3. On Site Test*" = True Then
              Tn = X
              GoTo 30
            End If
         Next X
         GoTo 35
30   '----------------
     Call GeneratePhotoTable(Tn)
     PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
     If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
31   If picN(tt) = 0 Then GoTo 32
     If picN(tt) >= 10 Then '10
        i = picN(tt)
        i5 = 1
         For X = 1 To i
             PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
             If InStr(1, PhotoName1, " ") > 0 Then
             I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
           Else
              I4(X) = Val(PhotoName1)
           End If
         Next X
          Call PhotoSort(i, I4, i5, PhotoName)
      End If
      For i3 = 1 To picN(tt)
          PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
            
            Show_Photo_Name = "1"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            'rnp = 1
            Call PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)
            
32         If picN(tt) > 0 Then
               WD.Tables(TbNumber).Rows(rnp + 1).Select
               Selection.InsertRowsBelow 2
            End If
      
            Show_Photo_Name = "1"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            Call PasteItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)

               
35 Path = S & PhotoPath(tt) & "\"
   Sap_Note_type = "Special Attention Point"
   Call Sap_photo(tt, Path, Sap_Note_type, msg)
   
   Sap_Note_type = "Reference Note"
   Call Note_photo(tt, Path, Sap_Note_type, msg)
   
   '-----------------------Product Dimension & Weight
   tt = 4
   
    tb = ThisDocument.Tables.count - 1
          For X = 1 To tb
            If ThisDocument.Tables(X).Cell(1, 1).Range.Text Like "*4. Product Dimension & Weight*" = True Then
              Tn = X
              GoTo 40
            End If
          Next X
          GoTo 45
       
40   '----------------
     Call GeneratePhotoTable(Tn)
    
     PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
     If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
    If picN(tt) = 0 Then GoTo 42
    If picN(tt) >= 10 Then '10
        i = picN(tt)
        i5 = 1
        For X = 1 To i
            PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
           If InStr(1, PhotoName1, " ") > 0 Then
             I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
           Else
              I4(X) = Val(PhotoName1)
           End If
        Next X
        Call PhotoSort(i, I4, i5, PhotoName)
    
    End If
      For i3 = 1 To picN(tt)
        PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
      
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            'rnp = 1
            Call PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)
42         If picN(tt) > 0 Then
               WD.Tables(TbNumber).Rows(rnp + 1).Select
               Selection.InsertRowsBelow 2
            End If
      
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            Call PasteItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)
               BigPhotoN = 0
               sp = S & PhotoPath(tt) & "\" & "Big Photo\"
               PhotoName(1) = Dir(sp & "*.*")
               If PhotoName(1) <> "" Then
                   i = 1
                 Do While PhotoName(i) <> ""
                     i = i + 1
                     PhotoName(i) = Dir
                 Loop
                    BigPhotoN = i - 1
               End If
               If BigPhotoN = 0 Then
                   tb = WD.Tables.count - 1
                   For X = tb To TbNumber Step -1
                     With WD.Tables(X)
                        If .Columns.count = 1 Then
                          If Trim(Mid(.Cell(1, 1).Range.Text, 1, Len(.Cell(1, 1).Range.Text) - 2)) = "" Then
                             WD.Tables(X).Select
                             Selection.Tables(1).Delete
                             Selection.Delete
                          End If
                        End If
                     End With
                   Next X
                  GoTo 45
               End If
               tb = WD.Tables.count - 1
               BigTbN = 0
               For i = TbNumber + 1 To tb
                 If WD.Tables(i).Columns.count = 1 Then
                    BigTbN = BigTbN + 1
                 Else
                    GoTo 43
                 End If
               Next i
43             If BigTbN > BigPhotoN Then
                  For i = TbNumber + BigTbN To TbNumber + BigPhotoN + 1 Step -1
                    WD.Tables(i).Select
                    Selection.Tables(1).Delete
                    Selection.Delete
                  Next i
               ElseIf BigTbN < BigPhotoN Then
                   For i = 1 To BigPhotoN - BigTbN
                      WD.Tables(TbNumber).Select
                      Selection.MoveDown
                      Selection.TypeParagraph
                      WD.Tables.add Range:=Selection.Range, NumRows:=2, NumColumns:=1, _
                      DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
                      With Selection.Tables(1)
                         If .Style <> UText(32593, 26684, 22411) Then
                           .Style = UText(32593, 26684, 22411)
                         End If
                      End With
                      Call Setborders025
                   Next i
               End If
               For i3 = TbNumber + 1 To TbNumber + BigPhotoN
                  With WD.Tables(i3)
                   Rown = .Rows.count
                   If Rown > 2 Then
                      For i = Rown To 3 Step -1
                        .Rows(i).Delete
                      Next i
                   ElseIf Rown = 1 Then
                      .Rows(1).Select
                      Selection.InsertRowsBelow 1
                   End If
                 End With
              Next i3
             
                   For i3 = 1 To BigPhotoN
                    With WD.Tables(TbNumber + i3)
                       .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                       .Cell(1, 1).VerticalAlignment = wdCellAlignVerticalCenter
                     .Cell(1, 1).Range.InlineShapes.AddPicture FileName:= _
                      sp & PhotoName(i3), LinkToFile:=False, _
                      SaveWithDocument:=True
                         .Cell(1, 1).Range.InlineShapes(1).Height = 340
                         .Cell(1, 1).Range.InlineShapes(1).Width = 470
                    
                      If Show_Photo_Name = "1" Then
                          Do While IsNumeric(Mid(PhotoName(i3), 1, 1)) = True
                            PhotoName(i3) = Mid(PhotoName(i3), 2, Len(PhotoName(i3)))
                          Loop
                          PhotoName(i3) = Trim(PhotoName(i3))
                          PhotoName(i3) = Mid(PhotoName(i3), 1, InStrRev(PhotoName(i3), ".") - 1)
                          PhotoName(i3) = UCase(Left(PhotoName(i3), 1)) & Mid(PhotoName(i3), 2, Len(PhotoName(i3)) - 1)
                      Else
                          PhotoName(i3) = i3
                      End If
                   
                     .Rows(2).Height = 0.6
                     .Cell(2, 1).Range.Text = PhotoName(i3)
                     .Cell(2, 1).Range.Font.Bold = False
                     .Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                     .Cell(2, 1).VerticalAlignment = wdCellAlignVerticalCenter
                   End With
                   Next i3
 
 
45  Path = S & PhotoPath(tt) & "\"
   Sap_Note_type = "Special Attention Point"
   Call Sap_photo(tt, Path, Sap_Note_type, msg)
   
   Sap_Note_type = "Reference Note"
   Call Note_photo(tt, Path, Sap_Note_type, msg)
   
   '------------------------------Style & Material & Construction
     tt = 5
      tb = WD.Tables.count - 1
       
         For X = 1 To tb
            With WD.Tables(X)
              If .Cell(1, 1).Range.Text Like "*5. Style / Material / Construction*" = True Then
                Tn = X
                GoTo 50
              End If
            End With
         Next X
         GoTo 55
     '-------
      NP = 0
     
50   Call GeneratePhotoTable(Tn)
     
     PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
     If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
    If picN(tt) = 0 Then GoTo 52
    If picN(tt) >= 10 Then '10
        i = picN(tt)
        i5 = 1
        For X = 1 To i
            PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
          If InStr(1, PhotoName1, " ") > 0 Then
             I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
           Else
              I4(X) = Val(PhotoName1)
           End If
        Next X
        Call PhotoSort(i, I4, i5, PhotoName)
    
    End If
      For i3 = 1 To picN(tt)
        PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
      
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            'rnp = 1
            Call PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)
            
52         If picN(tt) > 0 Then
             WD.Tables(TbNumber).Rows(rnp + 1).Select
               Selection.InsertRowsBelow 2
            End If
      
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            Call PasteItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)
     

55 Path = S & PhotoPath(tt) & "\"
   Sap_Note_type = "Special Attention Point"
   Call Sap_photo(tt, Path, Sap_Note_type, msg)
   
   Sap_Note_type = "Reference Note"
   Call Note_photo(tt, Path, Sap_Note_type, msg)
   
  '---------------------Product Color
     tt = 6
       tb = WD.Tables.count - 1
       
         For X = 1 To tb
            With WD.Tables(X)
              If .Cell(1, 1).Range.Text Like "*6. Product Color*" = True Then
                Tn = X
                GoTo 60
              End If
            End With
         Next X
         GoTo 65
     '-------
      NP = 0
     
60    Call GeneratePhotoTable(Tn)
       PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
     If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
    If picN(tt) = 0 Then GoTo 62
    If picN(tt) >= 10 Then '10
        i = picN(tt)
        i5 = 1
        For X = 1 To i
            PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
           If InStr(1, PhotoName1, " ") > 0 Then
             I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
           Else
              I4(X) = Val(PhotoName1)
           End If
        Next X
        Call PhotoSort(i, I4, i5, PhotoName)
    
    End If
      For i3 = 1 To picN(tt)
        PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
      
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            'rnp = 1
            Call PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)
            
62         If picN(tt) > 0 Then
               WD.Tables(TbNumber).Rows(rnp + 1).Select
               Selection.InsertRowsBelow 2
            End If
      
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            Call PasteItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)

65   Path = S & PhotoPath(tt) & "\"
   Sap_Note_type = "Special Attention Point"
   Call Sap_photo(tt, Path, Sap_Note_type, msg)
   
   Sap_Note_type = "Reference Note"
   Call Note_photo(tt, Path, Sap_Note_type, msg)
   
   '---------------------------------7. Product Label / Marking
68   tt = 7
     tb = WD.Tables.count - 1
       
         For X = 1 To tb
            With WD.Tables(X)
              If .Cell(1, 1).Range.Text Like "*7. Product Label / Marking*" = True Then
                Tn = X
                GoTo 70
              End If
            End With
         Next X
         GoTo 75
     '-------
      NP = 0
     
70   Call GeneratePhotoTable(Tn)
        PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
     If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
    If picN(tt) = 0 Then GoTo 72
    If picN(tt) >= 10 Then '10
        i = picN(tt)
        i5 = 1
        For X = 1 To i
            PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
           If InStr(1, PhotoName1, " ") > 0 Then
             I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
           Else
              I4(X) = Val(PhotoName1)
           End If
        Next X
        Call PhotoSort(i, I4, i5, PhotoName)
    
    End If
      For i3 = 1 To picN(tt)
        PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
      
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            'rnp = 1
            Call PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)
            
72         If picN(tt) > 0 Then
               WD.Tables(TbNumber).Rows(rnp + 1).Select
               Selection.InsertRowsBelow 2
            End If
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            Call PasteItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)
     
   
75 Path = S & PhotoPath(tt) & "\"
   Sap_Note_type = "Special Attention Point"
   Call Sap_photo(tt, Path, Sap_Note_type, msg)
   
   Sap_Note_type = "Reference Note"
   Call Note_photo(tt, Path, Sap_Note_type, msg)
   
   
78   '--------------------------8. Shipping Mark
    tt = 8
    tb = WD.Tables.count - 1
       
         For X = 1 To tb
            With WD.Tables(X)
              If .Cell(1, 1).Range.Text Like "*8. Shipping Mark*" = True Then
                Tn = X
                GoTo 80
              End If
            End With
         Next X
         GoTo 85
     '-------
      NP = 0
     
80   Call GeneratePhotoTable(Tn)
        PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
     If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
    If picN(tt) = 0 Then GoTo 82
    If picN(tt) >= 10 Then '10
        i = picN(tt)
        i5 = 1
        For X = 1 To i
            PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
            If InStr(1, PhotoName1, " ") > 0 Then
             I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
           Else
              I4(X) = Val(PhotoName1)
           End If
        Next X
        Call PhotoSort(i, I4, i5, PhotoName)
    
    End If
      For i3 = 1 To picN(tt)
        PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
      
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            'rnp = 1
            Call PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)
82         If picN(tt) > 0 Then
               WD.Tables(TbNumber).Rows(rnp + 1).Select
               Selection.InsertRowsBelow 2
            End If
      
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            Call PasteItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)
   
85 Path = S & PhotoPath(tt) & "\"
   Sap_Note_type = "Special Attention Point"
   Call Sap_photo(tt, Path, Sap_Note_type, msg)
   
   Sap_Note_type = "Reference Note"
   Call Note_photo(tt, Path, Sap_Note_type, msg)
   
   '--------------------------9. Packing / Packaging
    tt = 9
      tb = WD.Tables.count - 1
       
         For X = 1 To tb
            With WD.Tables(X)
              If .Cell(1, 1).Range.Text Like "*9. Packing / Packaging*" = True Then
                Tn = X
                GoTo 90
              End If
            End With
          Next X
          GoTo 95
     '-------
      NP = 0
     
90   Call GeneratePhotoTable(Tn)
     PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
     If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
    If picN(tt) = 0 Then GoTo 92
    If picN(tt) >= 10 Then '10
        i = picN(tt)
        i5 = 1
        For X = 1 To i
            PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
           If InStr(1, PhotoName1, " ") > 0 Then
             I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
           Else
              I4(X) = Val(PhotoName1)
           End If
        Next X
        Call PhotoSort(i, I4, i5, PhotoName)
    
    End If
      For i3 = 1 To picN(tt)
        PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
      
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            'rnp = 1
            Call PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)
            
92         If picN(tt) > 0 Then
               WD.Tables(TbNumber).Rows(rnp + 1).Select
               Selection.InsertRowsBelow 2
            End If
            
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            Call PasteItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)
   
95 Path = S & PhotoPath(tt) & "\"
   Sap_Note_type = "Special Attention Point"
   Call Sap_photo(tt, Path, Sap_Note_type, msg)
   
   Sap_Note_type = "Reference Note"
   Call Note_photo(tt, Path, Sap_Note_type, msg)
   
   '-------------10. Other Note
   tt = 10
         tb = ThisDocument.Tables.count - 1
          For X = 1 To tb
            If ThisDocument.Tables(X).Cell(1, 1).Range.Text Like "*10. Other Note*" = True Then
              Tn = X
              GoTo 100
            End If
        Next X
        GoTo 105
       
100  Call GeneratePhotoTable(Tn)
      PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
     If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
    If picN(tt) = 0 Then GoTo 102
    If picN(tt) >= 10 Then '10
        i = picN(tt)
        i5 = 1
        For X = 1 To i
            PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
           If InStr(1, PhotoName1, " ") > 0 Then
             I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
           Else
              I4(X) = Val(PhotoName1)
           End If
        Next X
        Call PhotoSort(i, I4, i5, PhotoName)
    
    End If
      For i3 = 1 To picN(tt)
        PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
      
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            'rnp = 1
            Call PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)
            
102        If picN(tt) > 0 Then
               WD.Tables(TbNumber).Rows(rnp + 1).Select
               Selection.InsertRowsBelow 2
            End If
            Show_Photo_Name = "0"
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            Call PasteItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)
            
                
    
     
    '------------------------------11. Other photo
105    tt = 11
     tb = WD.Tables.count - 1
         For X = 1 To tb
            With WD.Tables(X)
              If .Cell(1, 1).Range.Text Like "*Other Photos*" = True And .Range.Rows.count = 1 Then
                Tn = X
                GoTo 110
              End If
            End With
         Next X
         msg = msg + "The table name as Other Photos for locating other photo position was not found at report, the other photos were not pasted" & _
         CHR(10) & UText(25253, 21578, 20013, 26410, 21457, 29616, 29992, 20110, 23450, 20301) & "Other Photo" & UText(22270, 29255, 20301, 32622, 21517, 23383, 20026) & "Other Photo" & UText(30340, 34920, 26684) & "，" & UText(22240, 32780) & "Other Photo" & UText(22270, 29255, 26410, 36148, 19978) & CHR(10)
         GoTo 125
      NP = 0
     
110        If WD.Tables(Tn + 1).Columns.count = 4 And LCase(WD.Tables(Tn + 1).Cell(1, 1).Range.Text) Like "*inspector*" = True Then
                Rown = WD.Tables(Tn).Rows.count
                WD.Tables(Tn).Cell(Rown, 1).Select
                Selection.MoveDown
                Selection.TypeParagraph
                WD.Tables.add Range:=Selection.Range, NumRows:=2, NumColumns:=2, _
                DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
                With Selection.Tables(1)
                 If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
                 End If
              End With
              Call Setborders025
             End If
                
              WD.Tables(Tn + 1).Select
              With Selection.Tables(1).Range
                   Selection.Delete
                   Rown = .Rows.count
                   If Rown < 2 Then
                    .Rows(1).Select
                     Selection.InsertRows (2 - Rown)
                   ElseIf Rown > 2 Then
                     For i = Rown To 3 Step -1
                      .Rows(i).Delete
                    Next i
                   End If
              End With
              
        PhotoName(1) = Dir(S & PhotoPath(tt) & "\" & "*.*")
     If PhotoName(1) <> "" Then
         i = 1
         Do While PhotoName(i) <> ""
              i = i + 1
          PhotoName(i) = Dir
         Loop
         picN(tt) = i - 1
     End If
    If picN(tt) = 0 Then GoTo 112
    If picN(tt) >= 10 Then '10
        i = picN(tt)
        i5 = 1
        For X = 1 To i
            PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
           If InStr(1, PhotoName1, " ") > 0 Then
             I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
           Else
              I4(X) = Val(PhotoName1)
           End If
        Next X
        Call PhotoSort(i, I4, i5, PhotoName)
    
    End If
      For i3 = 1 To picN(tt)
        PhotoNameN(tt, i3) = S & PhotoPath(tt) & "\" & PhotoName(i3)
      Next i3
      
            Show_Photo_Name = ss2
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            'rnp = 1
            Call PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)
            
         If picN(tt) > 0 Then
               WD.Tables(TbNumber).Rows(rnp + 1).Select
               Selection.InsertRowsBelow 2
            End If
      
112         Show_Photo_Name = ss2
            TbNumber = Tn + 1
            pic_Number = picN(tt)
            Call PasteItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)
                    
124   '-----------
   
125 '---

 tb = WD.Tables.count - 1
  
  For i = tb - 2 To 4 Step -1
 

     With WD.Tables(i)
   
        If Len(Trim(.Cell(1, 1).Range.Text)) = 2 Then
           If .Columns.count = 2 And .Range.InlineShapes.count = 0 Then
              WD.Tables(i).Select
              Selection.Tables(1).Delete
              Selection.Delete
           End If
        End If
     End With
   Next i    '===========================================

 If msg <> "" Then
    MsgBox msg & " Please check and revise, if you are sure all are right, neglect this alarm." & CHR(10) & UText(35831, 26816, 26597) & "，" & UText(22914, 26524, 30830, 23454, 27809, 26377, 38382, 39064) & "，" & UText(35831, 24573, 30053, 26412, 25552, 31034), vbCritical
 Else
    MsgBox "Complete pasting photos." & CHR(10) & UText(36148, 22270, 23436, 25104) & "！", vbCritical
 End If
 
 
End Sub
Sub PhotoSort(i, I4, i5, PhotoName)
Dim skipnum, ii, j, index, i7 As Integer
Dim tempnum As String

  skipnum = Round(i / 2)
     Do While skipnum > 0
        ii = 1 + skipnum
        For j = ii To i
            index = j
            Do
                If index >= (1 + skipnum) Then
                    If i5 = 1 Then
                        If I4(index) < I4(index - skipnum) Then
                            tempnum = PhotoName(index)
                            PhotoName(index) = PhotoName(index - skipnum)
                            PhotoName(index - skipnum) = tempnum
                            i7 = I4(index)
                            I4(index) = I4(index - skipnum)
                            I4(index - skipnum) = i7
                            index = index - skipnum
                        Else
                            Exit Do
                        End If
                    Else
                        If PhotoName(index) < PhotoName(index - skipnum) Then
                            tempnum = PhotoName(index)
                            PhotoName(index) = PhotoName(index - skipnum)
                            PhotoName(index - skipnum) = tempnum
                            index = index - skipnum
                        Else
                            Exit Do
                        End If
                    End If
                Else
                    Exit Do
                End If
          Loop
       Next
'
        skipnum = Round(skipnum / 2)
     Loop
     
End Sub
Sub SortDefects(N, p)

   
    Dim i As Long, j As Long
    Dim temp As String
    Dim SortP() As String
    Dim Kp, Ks As String
    ReDim SortP(1 To N)
    SortP = p
    
    For i = 1 To N
        For j = i + 1 To N
           Kp = UCase(Right(SortP(i), Len(SortP(i)) - InStrRev(SortP(i), "|")))
           Ks = UCase(Right(SortP(j), Len(SortP(j)) - InStrRev(SortP(j), "|")))

            If Kp > Ks Then
                temp = SortP(i)
                 SortP(i) = SortP(j)
                 SortP(j) = temp
            End If
        Next j
    Next i
    p = SortP
End Sub
Sub PasteTwoColumnPhoto(tt, TbNumber, pic_Number, PhotoNameN, PhotoName, Show_Photo_Name, rnp, msg)

Dim i, i3, cnp As Integer


       Set WD = ThisDocument
       WD.Tables(TbNumber).Select
                With Selection.Tables(1)
                   For i3 = 1 To pic_Number
                      If i3 / 2 = Round(i3 / 2) Then
                        cnp = 2
                        rnp = i3 - 1
                      Else
                        cnp = 1
                        rnp = i3
                      End If
                      If .Rows(rnp).Range.Cells.count = 1 Then .Rows(rnp).Range.Cells.Split NumRows:=1, NumColumns:=2, mergebeforesplit:=True
                     .Cell(rnp, cnp).Range.InlineShapes.AddPicture FileName:= _
                      PhotoNameN(tt, i3), LinkToFile:=False, _
                      SaveWithDocument:=True
                      If .Cell(rnp, cnp).Range.InlineShapes.count > 0 Then
                        .Cell(rnp, cnp).Range.InlineShapes(1).Height = 170
                        .Cell(rnp, cnp).Range.InlineShapes(1).Width = 235
                      End If
                      If Show_Photo_Name = "1" Then
                          Do While IsNumeric(Mid(PhotoName(i3), 1, 1)) = True
                            PhotoName(i3) = Mid(PhotoName(i3), 2, Len(PhotoName(i3)))
                          Loop
                          PhotoName(i3) = Trim(PhotoName(i3))
                          PhotoName(i3) = Mid(PhotoName(i3), 1, InStrRev(PhotoName(i3), ".") - 1)
                          PhotoName(i3) = UCase(Left(PhotoName(i3), 1)) & Mid(PhotoName(i3), 2, Len(PhotoName(i3)) - 1)
                      Else
                          PhotoName(i3) = i3
                      End If
                     If .Rows(rnp + 1).Range.Cells.count = 1 Then .Rows(rnp + 1).Range.Cells.Split NumRows:=1, NumColumns:=2, mergebeforesplit:=True
                     .Rows(rnp + 1).Height = 0.6
                     .Cell(rnp + 1, cnp).Range.Text = PhotoName(i3) 'i3
                     .Cell(rnp + 1, cnp).Range.Font.Bold = False
                     .Cell(rnp + 1, cnp).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                     .Cell(rnp + 1, cnp).VerticalAlignment = wdCellAlignVerticalCenter
                     .Rows(rnp + 1).Select
                       If i3 < pic_Number Then
                            If i3 / 2 = Round(i3 / 2) Then Selection.InsertRowsBelow 2
                            
                           
                       Else
                           If i3 / 2 <> Round(i3 / 2) Then .Cell(rnp, 2).Range.Text = "NA"
                          .Cell(rnp, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                          .Cell(rnp, 2).VerticalAlignment = wdCellAlignVerticalCenter
                       End If
                    Next i3
                  End With
End Sub

Sub PasteItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)

Dim PhotoPathP(4000) As String   '2004000 Vega 20251126
Dim i2P, i, i2, I4(500), i5, JiaP, X, NP, TpicNP, ttp, PicNP(1000), cnp As Integer
Dim PhotoName(1000), TPhotoNameNP(1000) As String
ReDim Itx(1 To 500) '200500 Vega 20251126


       Set WD = ThisDocument
       i2P = 1
        sp = S & PhotoPath(tt) & "\"
        CurrentPhotoPath = Dir(sp, vbDirectory)
        Do While CurrentPhotoPath <> ""
           If CurrentPhotoPath <> "." And CurrentPhotoPath <> ".." And (GetAttr(sp & CurrentPhotoPath) And vbDirectory) = vbDirectory Then
             If Val(CurrentPhotoPath) > 0 Then
                PhotoPathP(i2P) = CurrentPhotoPath
                i2 = Val(CurrentPhotoPath)
                I4(i2) = i2
                Itx(i2P) = PhotoPathP(i2P)
                Do While IsNumeric(Mid(Itx(i2P), 1, 1)) = True
                  Itx(i2P) = Mid(Itx(i2P), 2, Len(Itx(i2P)) - 1)
                Loop
                i2P = i2P + 1
             End If
          End If
             CurrentPhotoPath = Dir
        Loop
         i2P = i2P - 1
         JiaP = i2P
         If JiaP = 0 Then
           msg = msg + PhotoPath(tt) & "has not any subfolder nameed with item name." & CHR(13) & _
           PhotoPath(tt) & UText(27809, 26377, 20197, 27454, 21495, 21629, 21517, 30340, 22270, 29255, 23376, 25991, 20214, 22841) & CHR(13)
           GoTo 25
         End If
         If JiaP >= 10 Then
                     i = i2P
                     i5 = 1
                     For X = 1 To i
                         'PhotoPathP1 = Left(PhotoPathP(x), 4)
                        I4(X) = Val(PhotoPathP(X))
                     Next X
          Call PhotoInSubfolderSort(i, I4, i5, PhotoPathP)
            
            '=== === '100 Vega 20251126
            Dim tmp As String, X1 As Long
            For X1 = 1 To JiaP
                tmp = PhotoPathP(X1)
                Do While IsNumeric(Left(tmp, 1))
                    tmp = Mid(tmp, 2)
                Loop
                Itx(X1) = tmp
            Next X1
            '========
        End If
         '-------
              NP = 0
              TpicNP = 0
              If pic_Number = 0 Then rnp = 1 Else rnp = rnp + 2
              For ttp = 1 To JiaP
                  PicNP(ttp) = 0
                  PhotoName(1) = Dir(sp & PhotoPathP(ttp) & "\" & "*.*")
                  If PhotoName(1) <> "" Then
                     i = 1
                     Do While PhotoName(i) <> ""
                       i = i + 1
                       PhotoName(i) = Dir
                     Loop
                      PicNP(ttp) = i - 1
                  End If
                  If PicNP(ttp) = 0 Then
                     msg = msg + sp & PhotoPathP(ttp) & UText(23376, 25991, 20214, 20013, 27809, 26377, 22270, 29255, 25991, 20214) & CHR(10) & sp & PhotoPathP(ttp) & " subfolder without photo file" & CHR(10)
                     GoTo 24
                  Else
                     TpicNP = TpicNP + PicNP(ttp)
                  End If
                  If PicNP(ttp) >= 10 Then
                     i = PicNP(ttp)
                     i5 = 1
                     For X = 1 To i
                         PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
                        If InStr(1, PhotoName1, " ") > 0 Then
                           I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
                        Else
                            I4(X) = Val(PhotoName1)
                        End If
                     Next X
                    Call PhotoSort(i, I4, i5, PhotoName)
                  End If
    
                      WD.Tables(TbNumber).Select
                      With Selection.Tables(1)
                        '.PreferredWidthType = wdPreferredWidthPoints
                        '.PreferredWidth = CentimetersToPoints(17.8)
                      For i3 = 1 To PicNP(ttp)
                          If i3 = 1 Then
                             .Rows(rnp).Range.Cells.Merge
                             .Rows(rnp).Height = 0.6
                            .Cell(rnp, 1).Range.Text = Itx(ttp)
                            .Cell(rnp, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                            .Cell(rnp, 1).VerticalAlignment = wdCellAlignVerticalCenter
                            .Cell(rnp, 1).Range.Font.Bold = True
                             If .Rows(rnp + 1).Range.Cells.count = 1 Then .Rows(rnp + 1).Range.Cells.Split NumRows:=1, NumColumns:=2, mergebeforesplit:=True
                                .Rows(rnp + 1).Select
                                 Selection.InsertRowsBelow 1
                                 rnp = rnp + 1
                          End If
                          If i3 Mod 2 = 0 Then
                               cnp = 2
                           Else
                              cnp = 1
                           End If
                           TPhotoNameNP(i3) = sp & PhotoPathP(ttp) & "\" & PhotoName(i3)
                          .Cell(rnp, cnp).Range.InlineShapes.AddPicture FileName:= _
                           TPhotoNameNP(i3), LinkToFile:=False, _
                           SaveWithDocument:=True
                           If .Cell(rnp, cnp).Range.InlineShapes.count > 0 Then
                            .Cell(rnp, cnp).Range.InlineShapes(1).Height = 170
                            .Cell(rnp, cnp).Range.InlineShapes(1).Width = 235
                           End If
                           If Show_Photo_Name = "1" Then
                              PhotoName(i3) = Trim(Left(PhotoName(i3), InStrRev(PhotoName(i3), ".") - 1))
                               Do While IsNumeric(Mid(PhotoName(i3), 1, 1)) = True
                                 PhotoName(i3) = Mid(PhotoName(i3), 2, Len(PhotoName(i3)) - 1)
                               Loop
                           Else
                               PhotoName(i3) = Itx(ttp) & "-" & i3
                           End If
                          .Rows(rnp + 1).Height = 0.6
                          .Cell(rnp + 1, cnp).Range.Text = PhotoName(i3)
                          .Cell(rnp + 1, cnp).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                          .Cell(rnp + 1, cnp).VerticalAlignment = wdCellAlignVerticalCenter
                          .Cell(rnp + 1, cnp).Range.Font.Bold = False
                          
                            If i3 < PicNP(ttp) Then
                               If i3 Mod 2 = 0 Then
                                 Selection.Tables(1).Rows(rnp + 1).Select
                                 Selection.InsertRowsBelow 2
                                 rnp = rnp + 2
                               End If
                            Else
                               If i3 Mod 2 <> 0 Then .Cell(rnp, 2).Range.Text = "NA"
                               .Cell(rnp, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                               .Cell(rnp, 2).VerticalAlignment = wdCellAlignVerticalCenter
                            End If
                        Next i3
                     End With
                  If ttp < JiaP Then
                        Selection.Tables(1).Rows(rnp + 1).Select
                        Selection.InsertRowsBelow 2
                        rnp = rnp + 2
                    End If
24              Next ttp

25              With WD.Tables(TbNumber)
                  Rown = .Rows.count
                  If .Cell(Rown - 1, 1).Range.InlineShapes.count = 0 Then
                    For X = Rown To Rown - 1 Step -1
                     .Rows(X).Delete
                    Next X
                  End If
                End With
  End Sub
Sub PasteAutoNameItemPhoto(tt, TbNumber, S, PhotoPath, pic_Number, Show_Photo_Name, rnp, msg)

Dim PhotoPathP(4000) As String                                                    '1004000 Vega 20251126
Dim i2P, i, i2, I4(500), i5, JiaP, X, NP, TpicNP, ttp, PicNP(1000), cnp As Integer
Dim PhotoName(1000), TPhotoNameNP(1000) As String
Dim PhotoNameP(500, 100) As String
 Dim TPhotoNameP(500, 100) As String
    
     Dim fso As Object
     Set fso = CreateObject("Scripting.FileSystemObject")
     Dim fld As Object
     Dim subfolder As Object
     Dim Itemfld As Object
     Dim Deffld As Object
     Dim Picfld As Object
     Dim Pht As Object
     Dim JiaPS As Integer
     Dim ItemFC() As Integer
     Dim DefFC() As Integer
     Dim Pname, PNameS As String
     Dim C As Integer
     i3 = 0
     JiaP = 0
     Set fld = fso.GetFolder(S & PhotoPath(tt))      '...\02. Workmanship（naming required)
     ReDim ItemFC(1 To fld.SubFolders.count)
     ReDim Itx(1 To fld.SubFolders.count)
     For Each subfolder In fld.SubFolders
        If LCase(subfolder.Name) Like "*item*" = True Then
                JiaP = JiaP + 1
                PicNP(JiaP) = 0
                Itx(JiaP) = subfolder.Name
                Do While IsNumeric(Mid(Itx(JiaP), 1, 1)) = True
                  Itx(JiaP) = Mid(Itx(JiaP), 2, Len(Itx(JiaP)) - 1)
                Loop
      
                Set Itemfld = fso.GetFolder(subfolder.Path & "\") '...\02. Workmanship（naming required)\01 Item abc
               
                  For Each Picfld In Itemfld.SubFolders
                    JiaPS = Itemfld.SubFolders.count
                     'Debug.Print Picfld.Path & "=Picfld.Path"
                     'Debug.Print Picfld.Name & "=Picfld.name"               '1 Critical
                     
                     Set Deffld = fso.GetFolder(Picfld.Path & "\")  '....\02. Workmanship（naming required)\01 Item abc\1 Critical
                   ' Debug.Print Deffld.Path
                     C = 0
                     If Deffld.Files.count > 0 Then
                        For Each Pht In Deffld.Files
                           Pname = Pht.Name
                           If InStrRev(Pname, ".") > 0 And Len(Pname) - InStrRev(Pname, ".") < 4 Then
                              PNameS = Mid(Pname, InStrRev(Pname, ".") + 1, Len(Pname) - InStrRev(Pname, "."))
                           Else
                              PNameS = "JPG"
                           End If
                           C = C + 1
                           Name Deffld.Path & "\" & Pname As Deffld.Path & "\" & Picfld.Name & " " & C & "." & PNameS
                           'Name Deffld.Path & "\" & Pht.Name As Deffld.Path & "\" & Picfld.Name & " " & C & "." & "JpG"
                           i3 = i3 + 1
                           PicNP(JiaP) = PicNP(JiaP) + 1
                           PhotoPathP(i3) = Deffld.Path
                           PhotoNameP(JiaP, PicNP(JiaP)) = Picfld.Name & " " & C & "." & PNameS
                           TPhotoNameP(JiaP, PicNP(JiaP)) = PhotoPathP(i3) & "\" & PhotoNameP(JiaP, PicNP(JiaP))
                         Next Pht
                  
                   
                    End If
               Next Picfld
        End If
     Next subfolder
     
   
      
      pic_Number = i3

       Set WD = ThisDocument
      
         If JiaP = 0 Then
           msg = msg + PhotoPath(tt) & "has not any subfolder nameed with item name." & CHR(13) & _
           PhotoPath(tt) & UText(27809, 26377, 20197, 27454, 21495, 21629, 21517, 30340, 22270, 29255, 23376, 25991, 20214, 22841) & CHR(13)
           GoTo 25
         End If
         
        
              NP = 0
              TpicNP = 0
              'If pic_Number = 0 Then rnp = 1 Else rnp = rnp + 2
              rnp = 1
              For ttp = 1 To JiaP
                 
                  If PicNP(ttp) = 0 Then
                     msg = msg + sp & PhotoPathP(ttp) & UText(23376, 25991, 20214, 20013, 27809, 26377, 22270, 29255, 25991, 20214) & CHR(10) & sp & PhotoPathP(ttp) & " subfolder without photo file" & CHR(10)
                     GoTo 24
                  End If
                
    
                      WD.Tables(TbNumber).Select
                      With Selection.Tables(1)
                        '.PreferredWidthType = wdPreferredWidthPoints
                        '.PreferredWidth = CentimetersToPoints(17.8)
                      For i3 = 1 To PicNP(ttp)
                          If i3 = 1 Then
                             .Rows(rnp).Range.Cells.Merge
                             .Rows(rnp).Height = 0.6
                            .Cell(rnp, 1).Range.Text = Itx(ttp)
                            .Cell(rnp, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                            .Cell(rnp, 1).VerticalAlignment = wdCellAlignVerticalCenter
                            .Cell(rnp, 1).Range.Font.Bold = True
                             If .Rows(rnp + 1).Range.Cells.count = 1 Then .Rows(rnp + 1).Range.Cells.Split NumRows:=1, NumColumns:=2, mergebeforesplit:=True
                                .Rows(rnp + 1).Select
                                 Selection.InsertRowsBelow 1
                                 rnp = rnp + 1
                          End If
                          If i3 Mod 2 = 0 Then
                               cnp = 2
                           Else
                              cnp = 1
                           End If
                           
                          .Cell(rnp, cnp).Range.InlineShapes.AddPicture FileName:= _
                           TPhotoNameP(ttp, i3), LinkToFile:=False, _
                           SaveWithDocument:=True
                           If .Cell(rnp, cnp).Range.InlineShapes.count > 0 Then
                            .Cell(rnp, cnp).Range.InlineShapes(1).Height = 170
                            .Cell(rnp, cnp).Range.InlineShapes(1).Width = 235
                           End If
                           If Show_Photo_Name = "1" Then
                              PhotoNameP(ttp, i3) = Trim(Left(PhotoNameP(ttp, i3), InStrRev(PhotoNameP(ttp, i3), ".") - 1))
                               Do While IsNumeric(Mid(PhotoNameP(ttp, i3), 1, 1)) = True
                                  PhotoNameP(ttp, i3) = Mid(PhotoNameP(ttp, i3), 2, Len(PhotoNameP(ttp, i3)) - 1)
                               Loop
                           Else
                               PhotoNameP(ttp, i3) = Itx(ttp) & "-" & i3
                           End If
                          .Rows(rnp + 1).Height = 0.6
                          .Cell(rnp + 1, cnp).Range.Text = PhotoNameP(ttp, i3)
                          .Cell(rnp + 1, cnp).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                          .Cell(rnp + 1, cnp).VerticalAlignment = wdCellAlignVerticalCenter
                          .Cell(rnp + 1, cnp).Range.Font.Bold = False
                          
                            If i3 < PicNP(ttp) Then
                               If i3 Mod 2 = 0 Then
                                 Selection.Tables(1).Rows(rnp + 1).Select
                                 Selection.InsertRowsBelow 2
                                 rnp = rnp + 2
                               End If
                            Else
                               If i3 Mod 2 <> 0 Then .Cell(rnp, 2).Range.Text = "NA"
                               .Cell(rnp, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                               .Cell(rnp, 2).VerticalAlignment = wdCellAlignVerticalCenter
                            End If
                        Next i3
                     End With
                  If ttp < JiaP Then
                        Selection.Tables(1).Rows(rnp + 1).Select
                        Selection.InsertRowsBelow 2
                        rnp = rnp + 2
                    End If
24              Next ttp

25              With WD.Tables(TbNumber)
                  Rown = .Rows.count
                  If .Cell(Rown - 1, 1).Range.InlineShapes.count = 0 Then
                    For X = Rown To Rown - 1 Step -1
                     .Rows(X).Delete
                    Next X
                  End If
                End With
  End Sub

Sub PhotoInSubfolderSort(i, I4, i5, PhotoPathP)
Dim skipnum, ii, j, index, i7 As Integer
Dim tempnum As String
                
                skipnum = Round(i / 2)
                Do While skipnum > 0
                        ii = 1 + skipnum
                        For j = ii To i
                          index = j
                          Do
                            If index >= (1 + skipnum) Then
                               If i5 = 1 Then
                                  If I4(index) < I4(index - skipnum) Then
                                     tempnum = PhotoPathP(index)
                                     PhotoPathP(index) = PhotoPathP(index - skipnum)
                                     PhotoPathP(index - skipnum) = tempnum
                                     i7 = I4(index)
                                     I4(index) = I4(index - skipnum)
                                     I4(index - skipnum) = i7
                                     index = index - skipnum
                                   Else
                                     Exit Do
                                   End If
                                Else
                                   If PhotoPathP(index) < PhotoPathP(index - skipnum) Then
                                       tempnum = PhotoPathP(index)
                                       PhotoPathP(index) = PhotoPathP(index - skipnum)
                                       PhotoPathP(index - skipnum) = tempnum
                                       index = index - skipnum
                                    Else
                                       Exit Do
                                    End If
                                End If
                            Else
                                 Exit Do
                            End If
                        Loop
                     Next j
                         skipnum = Round(skipnum / 2)
                   Loop
End Sub

   Sub InitialWms()
'
'
',by wyn/2025010824
'
'
Dim i, j, k, p, Q, R, Y, Cn, tb, ss, Qtb, cnt, CeC, CRow, Typ, lvl, Ass, grp, tt, ch, FX As Integer
Dim Styp, Agrp, Hgrp, Cgrp, Scode, Coln, Rown, msgn, CSmS, SamT As Integer
Dim qtyRow, qtyCol As Integer
Dim CD(20), SmpS(10), DSmpS(200, 3), ADef(10), AAc(10), DAAc(200, 3), AAc2(10), RRe(10), RRe2(10), Sm(200), Im(200) As Integer
Dim LG(20, 30), AC(20, 30), AC2(20, 30), Re(20, 30), MV(20, 30) As Integer
Dim Nfw, Naw, SQty, DSQty(200), scol(9), Lots(20), col(200, 9) As Long
Dim Acr, AMa, Ami, AQ(30), AQL(30), CHS, CHS1 As Single
Dim Ufw, Uaw, tmp, sms, msg, Aqlt, ACRE, itype, Smst, TBName, Unt As String
Dim kwSamp, kwInsL, kwAql, kwInsS, kwQty, kwSamS, Inslev, InsS, InsAql, POQty As String
Dim tem(10), txt(200, 10) As String
Dim rg, rgs As Range
Dim pic As InlineShape
Dim StartTime As Single
 'For Each pic In Thisdocument.InlineShapes
    'pic.Delete
 
 'Next
 '------------
         Dim SamplingStd_Name(), SamplingStd_Value() As String
         CCN = 9
         ReDim SamplingStd_Name(1 To CCN), SamplingStd_Value(1 To CCN)
  
          '--------
          SamplingStd_Name(1) = "ISO2859.1":        SamplingStd_Value(1) = "ISO2859.1"
          SamplingStd_Name(2) = "N/A":              SamplingStd_Value(2) = "N/A"
          SamplingStd_Name(3) = "Other":            SamplingStd_Value(3) = "Other"
          SamplingStd_Name(4) = "ANSI/ASQCZ1.4":    SamplingStd_Value(4) = "ANSI/ASQCZ1.4"
          SamplingStd_Name(5) = "BS6001.1":         SamplingStd_Value(5) = "BS6001.1"
          SamplingStd_Name(6) = "DIN40080":         SamplingStd_Value(6) = "DIN40080"
          SamplingStd_Name(7) = "NFX06-022":        SamplingStd_Value(7) = "NFX06-022"
          SamplingStd_Name(8) = "STD-019-HQT":      SamplingStd_Value(8) = "STD-019-HQT"
          SamplingStd_Name(9) = "STD-S-BLL-020":    SamplingStd_Value(9) = "STD-S-BLL-020"
        Dim InspectionLevel_Name(), InspectionLevel_Value() As String
        CCN = 12
        ReDim InspectionLevel_Name(1 To CCN), InspectionLevel_Value(1 To CCN)
           '--------
                 InspectionLevel_Name(1) = "G-II":               InspectionLevel_Value(1) = "G-II"
                 InspectionLevel_Name(2) = "G-II per Item":      InspectionLevel_Value(2) = "G-II per Item"
                 InspectionLevel_Name(3) = "G-I":                InspectionLevel_Value(3) = "G-I"
                 InspectionLevel_Name(4) = "G-I per Item":       InspectionLevel_Value(4) = "G-I per Item"
                 InspectionLevel_Name(5) = "Fixed Sample Size":  InspectionLevel_Value(5) = "Fixed Sample Size"
                 InspectionLevel_Name(6) = "G-III":              InspectionLevel_Value(6) = "G-III"
                 InspectionLevel_Name(7) = "G-III per Item":     InspectionLevel_Value(7) = "G-III per Item"
                 InspectionLevel_Name(8) = "S4":                 InspectionLevel_Value(8) = "S4"
                 InspectionLevel_Name(9) = "S3":                 InspectionLevel_Value(9) = "S3"
                 InspectionLevel_Name(10) = "S2":                InspectionLevel_Value(10) = "S2"
                 InspectionLevel_Name(11) = "S1":                InspectionLevel_Value(11) = "S1"
                 InspectionLevel_Name(12) = "Fixed Sample Size per Item":  InspectionLevel_Value(12) = "Fixed Sample Size per Item"
 '----
 
   
   TBName = "*1. Quantity*"
   tb = ThisDocument.Tables.count - 1
   For X = 1 To tb
   With ThisDocument.Range.Tables(X)
        If Trim(.Cell(1, 1).Range.Text) Like TBName = True Then
          tbn = X
          GoTo 5
        End If
     End With
  Next X
  
5   With ThisDocument.Range.Tables(tbn)    'shipment qty
       qtyRow = .Rows.count
       SQty = Val(.Cell(qtyRow, 3).Range.Text)
       If SQty = 0 Then MsgBox "1.Quantity" & UText(34920, 26684, 30340) & "total shipment quantity=0," & UText(35831, 20808, 36755, 22909, 35813, 34920, 26684, 25968, 20540) & "！" & CHR(10) & _
       "1.Quantity" & UText(34920, 26684) & ", shipment quantity=0, please complete inputing data first!", vbCritical: Exit Sub
        Itms = qtyRow - 5 '，workmanship
        ReDim Itx(1 To Itms)
          For i = 5 To qtyRow - 1
              Itx(i - 4) = Mid(.Cell(i, 2).Range.Text, 1, Len(.Cell(i, 2).Range.Text) - 2)
              Sm(i - 4) = Val(.Cell(i, 8).Range.Text)     ' translated comment
              DSQty(i - 4) = Val(.Cell(i, 4).Range.Text) ' translated comment
          Next i
         'SmS = Val(.Cell(qtyRow, 7).Range.Text)
          sco = Val(.Cell(qtyRow, 7).Range.Text)
          Unit = .Cell(4, 3).Range.words(2)
   End With

                '-----
   Dim ComItx_dict As Object
   Dim ComItx() As String
   ReDim ComItx(1 To Itms)
   Dim Key As String
   Dim TemSm As Integer
   Dim ComSm() As String
   ReDim ComSm(1 To Itms)
   Dim ComItms As Integer
   Dim ComDSQty() As Long
   ReDim ComDSQty(1 To Itms)
   Set ComItx_dict = CreateObject("Scripting.Dictionary")
        N = 0
        For i = 1 To Itms
          Key = Itx(i)
          TemSm = Sm(i)
          
          If Not ComItx_dict.Exists(Key) Then
           N = N + 1
           ComItx_dict.add Key, TemSm
           ComItx(N) = Key
           ComSm(N) = TemSm
           ComDSQty(N) = DSQty(i)
          Else
            ComItx_dict(Key) = ComItx_dict(Key) + TemSm
            ComSm(N) = ComItx_dict(Key)
            ComDSQty(N) = ComDSQty(N) + DSQty(i)
         End If
       Next i
      ComItms = ComItx_dict.count
   '------

   
   TBName = "*2. Workmanship*"
   tb = ThisDocument.Tables.count - 1
   For X = 1 To tb
   With ThisDocument.Range.Tables(X)
        If Trim(.Cell(1, 1).Range.Text) Like TBName = True Then
          tbn = X
          GoTo 10
        End If
     End With
  Next X

10 With ThisDocument.Range.Tables(tbn)
    
       If .Cell(3, 2).Range.Text Like "*Choose*" = True Or .Cell(4, 2).Range.Text Like "*Choose*" = True _
       Or .Cell(5, 2).Range.Text Like "*Choose*" = True Then MsgBox HubKeySamplingStandard() & "，" & HubKeySamplingPlan() & "，" & UText(26816, 39564, 27700, 24179, 26410, 23436, 25972, 36755, 20837) & "，" & UText(35831, 36755, 23436, 25972) & "!" & CHR(10) & _
       "Sampling Standard,Sampling Plan,or Inspection Level not input completely,please input!", vbCritical: Exit Sub
       
       If .Cell(4, 4).Range.ContentControls.count = 0 And Len(.Cell(4, 4).Range.Text) <= 2 _
       Or .Cell(4, 5).Range.ContentControls.count = 0 And Len(.Cell(4, 5).Range.Text) <= 2 _
       Or .Cell(4, 6).Range.ContentControls.count = 0 And Len(.Cell(4, 6).Range.Text) <= 2 Then _
       MsgBox "AQL" & UText(26410, 23436, 25972, 36755, 20837) & "，" & UText(22914, 26524, 27809, 26377) & "AQL，" & UText(35831, 36755, 20837) & "‘-’!" & CHR(10) & _
       "AQL not input completely,if there was not AQL please input‘-’!", vbCritical: Exit Sub

       Inslev = .Cell(5, 2).Range.Text
       If Inslev Like "*G-III*" = True Or Inslev Like "*III*" = True Then
       lvl = 7
       If LCase(Inslev) Like "*per item*" = True Then SamT = 1 Else SamT = 2   '1,，2，
       
       ElseIf Inslev Like "*G-II*" = True Or Inslev Like "*II*" = True Then
       lvl = 6
       If LCase(Inslev) Like "*per item*" = True Then SamT = 1 Else SamT = 2
       
       ElseIf Inslev Like "*G-I*" = True Or Inslev Like "*I*" = True Then
       lvl = 5
       If LCase(Inslev) Like "*per item*" = True Then SamT = 1 Else SamT = 2
       
       ElseIf Inslev Like "*S-4*" = True Then
       lvl = 4
       If LCase(Inslev) Like "*per item*" = True Then SamT = 1 Else SamT = 2
       
       ElseIf Inslev Like "*S-3*" = True Then
       lvl = 3
       If LCase(Inslev) Like "*per item*" = True Then SamT = 1 Else SamT = 2
       
       ElseIf Inslev Like "*S-2*" = True Then
       lvl = 2
       If LCase(Inslev) Like "*per item*" = True Then SamT = 1 Else SamT = 2
       
       ElseIf Inslev Like "*S-1*" = True Then
       lvl = 1
       If LCase(Inslev) Like "*per item*" = True Then SamT = 1 Else SamT = 2
       
       Else
        lvl = 0
       End If
   
    If LCase(.Cell(4, 2).Range.Text) Like "*fixed sample*" = True Or .Cell(4, 2).Range.Text Like "*" & UText(22266, 23450, 25277, 26679) & "*" = True Then
        'Styp = 5
        FX = 1
        If LCase(Inslev) Like "*per item*" = True Then SamT = 1 Else SamT = 2
       ' .Cell(3, 2).Range.Text = "N/A"
     
       
    ElseIf LCase(.Cell(4, 2).Range.Text) Like "*single, normal*" = True Or .Cell(4, 2).Range.Text Like "*" & UText(27491, 24120, 26816, 26597) & "," & UText(19968, 27425, 25277, 26679) & "*" = True Then
        Styp = 1
       
    End If
    
  
     ' ss = Val(.Cell(5, 3).Range.Text)

       

   'sample sizeAQL
 

       Lots(0) = 0
       Lots(1) = 2
       Lots(2) = 9
       Lots(3) = 16
       Lots(4) = 26
       Lots(5) = 51
       Lots(6) = 91
       Lots(7) = 151
       Lots(8) = 281
       Lots(9) = 501
       Lots(10) = 1201
       Lots(11) = 3201
       Lots(12) = 10001
       Lots(13) = 35001
       Lots(14) = 150001
       Lots(15) = 500001
       Lots(16) = 2147483647
'-------------------------------------------
        AQ(0) = 0
        AQ(1) = 0.01
        AQ(2) = 0.015
        AQ(3) = 0.025
        AQ(4) = 0.04
        AQ(5) = 0.065
        AQ(6) = 0.1
        AQ(7) = 0.15
        AQ(8) = 0.25
        AQ(9) = 0.4
        AQ(10) = 0.65
        AQ(11) = 1#
        AQ(12) = 1.5
        AQ(13) = 2.5
        AQ(14) = 4#
        AQ(15) = 6.5
        AQ(16) = 10
        AQ(17) = 15
        AQ(18) = 25
        AQ(19) = 40
        AQ(20) = 65
        AQ(21) = 100
        AQ(22) = 150
        AQ(23) = 250
        AQ(24) = 400
        AQ(25) = 650
        AQ(26) = 1000
        
        
        

   Select Case SamT   ' translated comment
   
   Case 2  ' translated comment
    
      
     'Select Case Styp
        'Case 1, 5 '2859,
        Call SingelNormal(CD, AC, AC2, Re, MV)
        ' Case 2
        ' Call SingelTightened(CD, AC, AC2, RE, MV)
         'Case 3
        'Call DoubleNormal(CD, AC, AC2, RE, MV)

        ' sampling code LG(), d1:for lot size, d2:inspecyion leval S-1,S-2，S-3，S-4，G-I，G-II,G-III
       LG(0, 1) = CD(0):      LG(0, 2) = CD(0):       LG(0, 3) = CD(0):    LG(0, 4) = CD(0):     LG(0, 5) = CD(0):   LG(0, 6) = CD(0):       LG(0, 7) = CD(0):
       LG(1, 1) = CD(1):      LG(1, 2) = CD(1):       LG(1, 3) = CD(1):    LG(1, 4) = CD(1):     LG(1, 5) = CD(1):   LG(1, 6) = CD(1):       LG(1, 7) = CD(2):
       LG(2, 1) = CD(1):      LG(2, 2) = CD(1):       LG(2, 3) = CD(1):    LG(2, 4) = CD(1):     LG(2, 5) = CD(1):   LG(2, 6) = CD(2):       LG(2, 7) = CD(3):
       LG(3, 1) = CD(1):      LG(3, 2) = CD(1):       LG(3, 3) = CD(2):    LG(3, 4) = CD(2):     LG(3, 5) = CD(2):   LG(3, 6) = CD(3):       LG(3, 7) = CD(4):
       LG(4, 1) = CD(1):      LG(4, 2) = CD(2):       LG(4, 3) = CD(2):    LG(4, 4) = CD(3):     LG(4, 5) = CD(3):   LG(4, 6) = CD(4):       LG(4, 7) = CD(5):
       LG(5, 1) = CD(2):      LG(5, 2) = CD(2):       LG(5, 3) = CD(3):    LG(5, 4) = CD(3):     LG(5, 5) = CD(3):   LG(5, 6) = CD(5):       LG(5, 7) = CD(6):
       LG(6, 1) = CD(2):      LG(6, 2) = CD(2):       LG(6, 3) = CD(3):    LG(6, 4) = CD(4):     LG(6, 5) = CD(4):   LG(6, 6) = CD(6):       LG(6, 7) = CD(7):
       LG(7, 1) = CD(2):      LG(7, 2) = CD(3):       LG(7, 3) = CD(4):    LG(7, 4) = CD(5):     LG(7, 5) = CD(5):   LG(7, 6) = CD(7):       LG(7, 7) = CD(8):
       LG(8, 1) = CD(2):      LG(8, 2) = CD(3):       LG(8, 3) = CD(4):    LG(8, 4) = CD(5):     LG(8, 5) = CD(6):   LG(8, 6) = CD(8):       LG(8, 7) = CD(9):
       LG(9, 1) = CD(3):      LG(9, 2) = CD(3):       LG(9, 3) = CD(5):    LG(9, 4) = CD(6):     LG(9, 5) = CD(7):   LG(9, 6) = CD(9):       LG(9, 7) = CD(10):
       LG(10, 1) = CD(3):     LG(10, 2) = CD(4):      LG(10, 3) = CD(5):   LG(10, 4) = CD(7):    LG(10, 5) = CD(8):  LG(10, 6) = CD(10):     LG(10, 7) = CD(11):
       LG(11, 1) = CD(3):     LG(11, 2) = CD(4):      LG(11, 3) = CD(6):   LG(11, 4) = CD(7):    LG(11, 5) = CD(9):  LG(11, 6) = CD(11):     LG(11, 7) = CD(12):
       LG(12, 1) = CD(3):     LG(12, 2) = CD(4):      LG(12, 3) = CD(6):   LG(12, 4) = CD(8):    LG(12, 5) = CD(10): LG(12, 6) = CD(12):     LG(12, 7) = CD(13):
       LG(13, 1) = CD(4):     LG(13, 2) = CD(5):      LG(13, 3) = CD(7):   LG(13, 4) = CD(9):    LG(13, 5) = CD(11): LG(13, 6) = CD(13):     LG(13, 7) = CD(14):
       LG(14, 1) = CD(4):     LG(14, 2) = CD(5):      LG(14, 3) = CD(7):   LG(14, 4) = CD(9):    LG(14, 5) = CD(12): LG(14, 6) = CD(14):     LG(14, 7) = CD(15):
       LG(15, 1) = CD(4):     LG(15, 2) = CD(5):      LG(15, 3) = CD(8):   LG(15, 4) = CD(10):   LG(15, 5) = CD(13): LG(15, 6) = CD(15):     LG(15, 7) = CD(16)
    
  '-----------------------------------------------------------
 
       
      If .Cell(4, 4).Range.Text Like "*" & UText(19981, 20801, 35768) & "*" = True Or LCase(.Cell(4, 4).Range.Text) Like "*not allowed*" = True Or Val(.Cell(4, 4).Range.Text) = 0 Then
        Acr = 0
        Y = 1
      Else
       Acr = Val(.Cell(4, 4).Range.Text)
       Y = 0
    
      End If
   
       AMa = Val(.Cell(4, 5).Range.Text)
       Ami = Val(.Cell(4, 6).Range.Text)
       AQL(0) = Acr
       AQL(1) = AMa
       AQL(2) = Ami
       
      
      
       If SQty < Lots(1) Then MsgBox "The batch size was too small." & CHR(13) & UText(25209, 37327, 25968, 22826, 23567) & "，" & UText(26412, 31243, 24207, 19981, 33021, 35745, 31639) & "，" & UText(35831, 25163, 24037, 35745, 31639): Exit Sub
        ss = Val(.Cell(5, 4).Range.Text)
        '＊＊＊fixed sample AQL
      
        Dim Cell As Cell
        
        If FX = 1 Then
            ss = Val(.Cell(5, 4).Range.Text)   'workmanship
           If ss = 0 Then MsgBox UText(22266, 23450, 25277, 26679) & "，" & UText(35831, 20808, 36755, 20837, 22312) & "workmanship" & UText(34920, 36755, 20837) & "sample size" & CHR(10) & _
           "For fixed sample plan, please input sample size at workmanship table !", vbCritical: Exit Sub
           If ss <> sco Then MsgBox UText(22266, 23450, 25277, 26679, 22312) & "2.Workmanship" & UText(34920, 36755, 20837, 30340, 25277, 26679, 25968, 19981, 31561, 20110) & "1.Quantity" & UText(34920, 30340, 25277, 26679, 25968) & "，" & UText(35831, 26356, 27491, 21518, 20877, 36816, 34892, 31243, 24207) & CHR(10) & _
           "The sample size in 2.Workmanship table is not the same with which in 1.Quantity table, please revise it and run program again!!", vbCritical

           For i = 1 To 15
             If ss = CD(i) Then
                 grp = i
                 
                 Set Cell = .Cell(5, 2)
                 If LCase(Cell.Range.Text) Like "*per item*" = True Then
                    Option_Selected = "Fixed Sample Size per Item"
                 Else
                     Option_Selected = "Fixed Sample Size"
                 End If
                
                 Cell.Range.Delete
                 Set Occ = Cell.Range.ContentControls.add(wdContentControlComboBox)
                 With Occ
                   .Title = "Inspection Level"
                   .Tag = "Inspection Level"
                     For j = LBound(InspectionLevel_Name) To UBound(InspectionLevel_Name)
                      .DropdownListEntries.add InspectionLevel_Name(j), InspectionLevel_Value(j)
                     Next j
                     For Each dropdownEntry In .DropdownListEntries
                       If dropdownEntry.Text = Option_Selected Then
                             dropdownEntry.Select
                             Exit For
                        End If
                      Next dropdownEntry
                                                   ' Debug.Print "Inspection Level=" & oCC.Range.Text
                 End With
                  ' .cell(5, 2).Range.Text = "Fixed Sample Size"
                   GoTo 20
               End If
             Next i
            FX = 2    'sample size
           .Cell(4, 4).Range.Text = "-"
           .Cell(4, 5).Range.Text = "-"
           .Cell(4, 6).Range.Text = "-"
           .Cell(5, 4).Range.Text = ss
           .Cell(5, 5).Range.Text = ss
           .Cell(5, 6).Range.Text = ss
         ' .cell(5, 2).Range.Text = "-"
           GoTo 70
        End If
  

        'LotSiZE
       For i = 1 To 15
         If SQty >= Lots(i) And SQty < Lots(i + 1) Then grp = i: GoTo 20
       Next i
       
    'AQLSampleSiz
      
20     Scode = LG(grp, lvl)
       For i = 1 To 16
          If Scode = CD(i) Then Hgrp = i: GoTo 25
       Next i
       
25    For k = Y To 2  'k AQLMa,Mi
        'AQL
         Cn = 0
         For i = 1 To 16
           If AQL(k) = AQ(i) Then Agrp = i: GoTo 30
           If AQL(k) <> AQ(i) And AQL(k) <> 0 Then Cn = Cn + 1 'AQL
         Next i

        If Cn = 16 Then
           If FX = 1 Then
                FX = 3       'AQL
                .Cell(4, 4).Range.Text = "-"
                .Cell(4, 5).Range.Text = "-"
                .Cell(4, 6).Range.Text = "-"
               .Cell(5, 4).Range.Text = ss
               .Cell(5, 5).Range.Text = ss
               .Cell(5, 6).Range.Text = ss
                GoTo 70
           Else
             MsgBox "2.Workmanship" & UText(34920, 26684, 20013, 36755, 20837, 30340) & "AQL" & UText(20540, 19981, 22312, 31243, 24207, 35268, 23450, 33539, 22260, 20869) & "，" & UText(35831, 20462, 25913, 25110, 25163, 24037, 22788, 29702) & "！" & CHR(10) & _
             "2.Workmanship table, input AQL value out of the program data scope, please revised it or dealing with it by hand！", vbCritical: Exit Sub
          End If
       End If
30       If MV(Hgrp, Agrp) > 0 Then
             If MV(Hgrp, Agrp) >= SQty Then SmpS(k) = SQty
             If MV(Hgrp, Agrp) < SQty Then SmpS(k) = MV(Hgrp, Agrp)
         End If
         
         If MV(Hgrp, Agrp) > 0 And FX = 1 Then
           If MV(Hgrp, Agrp) > ss Then    'sQty
              FX = 4                     'fixed sample
                .Cell(4, 4).Range.Text = "-"
                .Cell(4, 5).Range.Text = "-"
                .Cell(4, 6).Range.Text = "-"
            
                .Cell(5, 4).Range.Text = ss
                .Cell(5, 5).Range.Text = ss
                .Cell(5, 6).Range.Text = ss
               GoTo 70
           End If
         End If
        If MV(Hgrp, Agrp) = 0 Then SmpS(k) = CD(Hgrp)
        AAc(k) = AC(Hgrp, Agrp)
 
 '-------，，
        If Styp = 1 Or Styp = 2 Or Styp = 4 Then
             RRe(k) = AAc(k) + 1
        Else
            AAc2(k) = AC2(Hgrp, Agrp)
            RRe(k) = Re(Hgrp, Agrp)
            RRe2(k) = AC2(Hgrp, Agrp) + 1
        End If
        If AAc(k) = 0 Then RRe(k) = 1
        If AAc2(k) = "" Then AAc2(k) = "-"
        If RRe2(k) = "" Then RRe2(k) = "-"

40    If SmpS(k) > Ass Then Ass = SmpS(k)
        If AAc(k) = o And RRe(k) = 1 Then
           AAc2(k) = "-"
           RRe2(k) = "-"
        End If
    '=======================================
50   Next k
      If Acr = 0 Then
         If Ass >= Scode Then
             SmpS(0) = Ass
          Else
            SmpS(0) = Scode
            Ass = Scode
          End If
          'SmpS(0) = Ass
          AAc(0) = 0
          RRe(0) = 1
          AAc2(0) = "-"
          RRe2(0) = "-"
      End If
  '------------------------
      
 
 
        If Ass <> sco And Styp = 1 Then MsgBox UText(34920, 26684) & "1.Quantity" & UText(32479, 35745, 20986, 30340, 25277, 26679, 25968, 26159) & "：" & sco & " " & UText(20294, 26159, 25353) & "ISO2859" & UText(26631, 20934, 26597, 20986, 30340) & "Sample Size" & UText(24212, 35813, 26159) & ":" & Ass & _
       " " & UText(20004, 32773, 19981, 19968, 33268) & "，" & UText(35831, 36827, 34892, 20462, 25913, 21518, 20877, 27425, 36816, 34892, 31243, 24207) & "！！！" & CHR(10) & _
       "The sample size in 1.Quantity table is " & sco & ", however according to ISO28591 the sample size shall be " & Ass & _
       ", they are not the same, please revise it and run program again!!!": GoTo 1000
      .Cell(5, 4).Range.Text = SmpS(0)
      .Cell(5, 5).Range.Text = SmpS(1)
      .Cell(5, 6).Range.Text = SmpS(2)
      
       If FX = 1 Then
        
           Set Cell = .Cell(3, 2)
           Cell.Range.Delete
           Set Occ = Cell.Range.ContentControls.add(wdContentControlComboBox)
        
           With Occ
            .Title = "Sampling Standart"
            .Tag = "Sampling Standart"
             For i = LBound(SamplingStd_Name) To UBound(SamplingStd_Name)
              .DropdownListEntries.add SamplingStd_Name(i), SamplingStd_Value(i)
             Next i
             .DropdownListEntries(8).Select   '"STD-019-HQT"
           End With
           '.Cell(3, 2).Range.Text = "STD-019-HQT"
       End If
70      With ThisDocument.Range.Tables(tbn + 1) 'defects---------
          If LCase(.Cell(1, 1).Range.Text) Like "*defects*" = False Then MsgBox UText(22312) & "2.wordmanship" & UText(34920, 19979, 31532, 19968, 20010, 34920, 26684, 19981, 26159) & "Defect" & UText(34920) & "，" & UText(35831, 26356, 27491) & "！" & CHR(10) & _
          "Not Defects Table is found below the 2.workmanship table, please revise it!!!", vbCritical: Exit Sub
            Coln = .Columns.count
            Rown = .Rows.count
  
            '----- -------------
               ' translated comment
               N = 0
                For i = Rown - 2 To 3 Step -1
                 If .Rows(i).Cells.count > 1 Then
                   If .Cell(i, 2).Range.Text Like "*Allowed:*" = True Or .Cell(i, 2).Range.Text Like "*Total Found:*" = True Then
                      .Rows(i).Delete
                    Else
                       If .Cell(i, 2).Range.Fields.count > 0 Then
                          If .Cell(i, 2).Range.Fields(1).Code Like "*Total Found:*" = True Then
                            .Rows(i).Delete
                          End If
                       End If
                    End If
                End If
                Next i
                
                For i = 2 To Rown - 2
                   If .Rows(i).Cells.count = 1 And LCase(.Cell(i, 1).Range.Text) Like "*item*" = False Then
                     .Cell(i, 1).Range.Text = "Item No."
                   End If
                   If .Rows(i).Cells.count = 1 And LCase(.Cell(i, 1).Range.Text) Like "*item*" = True Then
                       Im(N + 1) = i   'Item
                       N = N + 1       'item
                    End If
                  Next i
                  If N = 0 Then MsgBox "2. Workmanship" & UText(34920, 19979, 30340) & "Defec" & UText(34920, 27809, 26377, 19968, 25972, 34892, 29992, 20110, 22635, 20889) & "item" & UText(21495, 30340, 34892) & "，" & UText(35831, 22686, 21152, 21518, 20877, 36816, 34892, 31243, 24207) & "。" & CHR(13) & _
                  "The Defect tablle has is not any row for inputing Item No.,please revise it and run program again！": Exit Sub

                If N > ComItms Then
                  Set rg = ThisDocument.Range(Start:=.Rows(Im(ComItms + 1)).Range.Start, End:=.Rows(Rown - 2).Range.End)
                  rg.Select
                  Selection.Rows.Delete
                 ElseIf N < ComItms And N > 1 Then
                   Set rg = ThisDocument.Range(Start:=.Rows(Im(1)).Range.Start, End:=.Rows(Im(2) - 1).Range.End)
                   rg.Copy
                   For j = 1 To ComItms - N
                    .Rows(Im(1)).Select
                      
                        StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
                     Selection.Paste
                  Next j
                 Else
                    Set rg = ThisDocument.Range(Start:=.Rows(Im(1)).Range.Start, End:=.Rows(Rown - 2).Range.End)
                    rg.Copy
                    For j = 1 To ComItms - N
                     .Rows(Im(1)).Select
                        StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
                       Selection.Paste
                     Next j
                End If

             Rown = .Rows.count
             N = 0
             For i = 2 To Rown - 2
                If .Rows(i).Cells.count = 1 And LCase(.Cell(i, 1).Range.Text) Like "*item*" = True Then
                N = N + 1
               .Cell(i, 1).Range.Text = "Item No. " & ComItx(N) & "," & " Sample Size= " & ComSm(N) & " " & Unit & "s"
                
                End If
             Next i
   
           If FX < 2 Then
            .Cell(Rown, Coln - 2).Range.Text = AAc(0)
            .Cell(Rown, Coln - 1).Range.Text = AAc(1)
            .Cell(Rown, Coln).Range.Text = AAc(2)
           ElseIf FX > 1 Then
             .Cell(Rown, Coln - 2).Range.Text = "-"
             .Cell(Rown, Coln - 1).Range.Text = "-"
             .Cell(Rown, Coln).Range.Text = "-"
          
           End If
       End With
           If FX > 1 Then
               .Cell(4, 4).Range.Text = "-"
               .Cell(4, 5).Range.Text = "-"
               .Cell(4, 6).Range.Text = "-"
                Set Cell = .Cell(3, 2)
                Cell.Range.Delete
                Set Occ = Cell.Range.ContentControls.add(wdContentControlComboBox)
        
                With Occ
                 .Title = "Sampling Standart"
                 .Tag = "Sampling Standart"
                  For i = LBound(SamplingStd_Name) To UBound(SamplingStd_Name)
                    .DropdownListEntries.add SamplingStd_Name(i), SamplingStd_Value(i)
                   Next i
                   .DropdownListEntries(2).Select   '"N/A"
                 End With
               
             End If
       
       
       GoTo 1000
      'End select
    
    
      Case 1    ' translated comment
          '------------------- -----------------------
       Set dict = CreateObject("Scripting.Dictionary")
    
      For i = 1 To ComItms
        If Not dict.Exists(ComSm(i)) Then
            dict.add ComSm(i), Nothing
        End If
     Next i
     N = dict.count
     Set dict = Nothing

     If N > 1 Then   'per size
     
       FX = 5
       With WD.Tables(getTableNumber(2))
         .Cell(5, 4).Range.Text = sco
         .Cell(5, 5).Range.Text = sco
         .Cell(5, 6).Range.Text = sco
       End With
       
       GoTo 170
     End If
  
      
      For tt = 1 To ComItms
      Ass = 0
  
        Call SingelNormal(CD, AC, AC2, Re, MV)
        ' sampling code LG(), d1:for lot size, d2:inspecyion leval S-1,S-2，S-3，S-4，G-I，G-II,G-III
       LG(0, 1) = CD(0):      LG(0, 2) = CD(0):       LG(0, 3) = CD(0):    LG(0, 4) = CD(0):     LG(0, 5) = CD(0):   LG(0, 6) = CD(0):       LG(0, 7) = CD(0):
       LG(1, 1) = CD(1):      LG(1, 2) = CD(1):       LG(1, 3) = CD(1):    LG(1, 4) = CD(1):     LG(1, 5) = CD(1):   LG(1, 6) = CD(1):       LG(1, 7) = CD(2):
       LG(2, 1) = CD(1):      LG(2, 2) = CD(1):       LG(2, 3) = CD(1):    LG(2, 4) = CD(1):     LG(2, 5) = CD(1):   LG(2, 6) = CD(2):       LG(2, 7) = CD(3):
       LG(3, 1) = CD(1):      LG(3, 2) = CD(1):       LG(3, 3) = CD(2):    LG(3, 4) = CD(2):     LG(3, 5) = CD(2):   LG(3, 6) = CD(3):       LG(3, 7) = CD(4):
       LG(4, 1) = CD(1):      LG(4, 2) = CD(2):       LG(4, 3) = CD(2):    LG(4, 4) = CD(3):     LG(4, 5) = CD(3):   LG(4, 6) = CD(4):       LG(4, 7) = CD(5):
       LG(5, 1) = CD(2):      LG(5, 2) = CD(2):       LG(5, 3) = CD(3):    LG(5, 4) = CD(3):     LG(5, 5) = CD(3):   LG(5, 6) = CD(5):       LG(5, 7) = CD(6):
       LG(6, 1) = CD(2):      LG(6, 2) = CD(2):       LG(6, 3) = CD(3):    LG(6, 4) = CD(4):     LG(6, 5) = CD(4):   LG(6, 6) = CD(6):       LG(6, 7) = CD(7):
       LG(7, 1) = CD(2):      LG(7, 2) = CD(3):       LG(7, 3) = CD(4):    LG(7, 4) = CD(5):     LG(7, 5) = CD(5):   LG(7, 6) = CD(7):       LG(7, 7) = CD(8):
       LG(8, 1) = CD(2):      LG(8, 2) = CD(3):       LG(8, 3) = CD(4):    LG(8, 4) = CD(5):     LG(8, 5) = CD(6):   LG(8, 6) = CD(8):       LG(8, 7) = CD(9):
       LG(9, 1) = CD(3):      LG(9, 2) = CD(3):       LG(9, 3) = CD(5):    LG(9, 4) = CD(6):     LG(9, 5) = CD(7):   LG(9, 6) = CD(9):       LG(9, 7) = CD(10):
       LG(10, 1) = CD(3):     LG(10, 2) = CD(4):      LG(10, 3) = CD(5):   LG(10, 4) = CD(7):    LG(10, 5) = CD(8):  LG(10, 6) = CD(10):     LG(10, 7) = CD(11):
       LG(11, 1) = CD(3):     LG(11, 2) = CD(4):      LG(11, 3) = CD(6):   LG(11, 4) = CD(7):    LG(11, 5) = CD(9):  LG(11, 6) = CD(11):     LG(11, 7) = CD(12):
       LG(12, 1) = CD(3):     LG(12, 2) = CD(4):      LG(12, 3) = CD(6):   LG(12, 4) = CD(8):    LG(12, 5) = CD(10): LG(12, 6) = CD(12):     LG(12, 7) = CD(13):
       LG(13, 1) = CD(4):     LG(13, 2) = CD(5):      LG(13, 3) = CD(7):   LG(13, 4) = CD(9):    LG(13, 5) = CD(11): LG(13, 6) = CD(13):     LG(13, 7) = CD(14):
       LG(14, 1) = CD(4):     LG(14, 2) = CD(5):      LG(14, 3) = CD(7):   LG(14, 4) = CD(9):    LG(14, 5) = CD(12): LG(14, 6) = CD(14):     LG(14, 7) = CD(15):
       LG(15, 1) = CD(4):     LG(15, 2) = CD(5):      LG(15, 3) = CD(8):   LG(15, 4) = CD(10):   LG(15, 5) = CD(13): LG(15, 6) = CD(15):     LG(15, 7) = CD(16)
    
  '-----------------------------------------------------------
 
       
      If .Cell(4, 4).Range.Text Like "*" & UText(19981, 20801, 35768) & "*" = True Or Val(.Cell(4, 4).Range.Text) = 0 Then
        Acr = 0
        Y = 1
        AAc(0) = 0
      Else
       Acr = Val(.Cell(4, 4).Range.Text)
       Y = 0
    
      End If
       AMa = Val(.Cell(4, 5).Range.Text)
       Ami = Val(.Cell(4, 6).Range.Text)
       AQL(0) = Acr
       AQL(1) = AMa
       AQL(2) = Ami
       

     
       SQty = ComDSQty(tt)


       If SQty < Lots(1) Then MsgBox HubKeyItemNo() & ComItx(tt) & UText(25209, 37327, 25968, 22826, 23567) & "，" & UText(26412, 31243, 24207, 19981, 33021, 35745, 31639) & "，" & UText(35831, 25163, 24037, 35745, 31639) & CHR(10) & _
       "Lot size of item" & ComItx(tt) & "is two small for this program, please calculate by hand！": Exit Sub

        '＊＊＊fixed sample AQL
         If FX = 1 Then
           ss = Val(.Cell(5, 4).Range.Text)   'workmanship
           If ss = 0 Then MsgBox UText(22266, 23450, 25277, 26679) & "，" & UText(35831, 20808, 36755, 20837, 22312) & "workmanship" & UText(34920, 36755, 20837) & "sample size" & CHR(10) & _
           "For fixed sample plan, please input sample size at workmanship table !", vbCritical: Exit Sub
             If ss <> sco Then MsgBox UText(22266, 23450, 25277, 26679, 22312) & "2.Workmanship" & UText(34920, 36755, 20837, 30340, 25277, 26679, 25968, 19981, 31561, 20110) & "1.Quantity" & UText(34920, 30340, 25277, 26679, 25968) & "，" & UText(35831, 26356, 27491, 21518, 20877, 36816, 34892, 31243, 24207) & CHR(10) & _
            "The sample size in 2.Workmanship table is not the same with which in 1.Quantity table, please revise it and run program again!!", vbCritical
           ss = ss / ComItms
           For i = 1 To ComItms
                If ss <> ComSm(i) Then MsgBox UText(22266, 23450, 25277, 26679, 27454, 21495) & "：" & ComItx(i) & UText(22312) & "2.Workmanship" & UText(34920, 36755, 20837, 30340, 25277, 26679, 25968, 19981, 31561, 20110) & "1.Quantity" & UText(34920, 30340, 25277, 26679, 25968) & "，" & UText(35831, 26356, 27491, 21518, 20877, 36816, 34892, 31243, 24207) & CHR(10) & _
               "The sample size in 2.Workmanship table is not the same with which in 1.Quantity table, please revise it and run program again!!", vbCritical
           Next i
          
          
           For i = 1 To 15
              
              If ss = CD(i) Then
                 grp = i
        
                 Option_Selected = "Fixed Sample Size per Item"
                 Set Cell = .Cell(5, 2)
                 Cell.Range.Delete
                 Set Occ = Cell.Range.ContentControls.add(wdContentControlComboBox)
                 With Occ
                   .Title = "Inspection Level"
                   .Tag = "Inspection Level"
                     For j = LBound(InspectionLevel_Name) To UBound(InspectionLevel_Name)
                      .DropdownListEntries.add InspectionLevel_Name(j), InspectionLevel_Value(j)
                     Next j
                     For Each dropdownEntry In .DropdownListEntries
                       If dropdownEntry.Text = Option_Selected Then
                          dropdownEntry.Select
                          Exit For
                        End If
                      Next dropdownEntry
                                                   ' Debug.Print "Inspection Level=" & oCC.Range.Text
                 End With
                 GoTo 120
               End If
               
           Next i
            FX = 2    'sample size
           .Cell(4, 4).Range.Text = "-"
           .Cell(4, 5).Range.Text = "-"
           .Cell(4, 6).Range.Text = "-"
           .Cell(5, 4).Range.Text = ss * ComItms
           .Cell(5, 5).Range.Text = ss * ComItms
           .Cell(5, 6).Range.Text = ss * ComItms
         ' .cell(5, 2).Range.Text = "-"
           GoTo 170
        End If
      
        '--------------
 

        'LotSiZE
       For i = 1 To 15
         If SQty >= Lots(i) And SQty < Lots(i + 1) Then grp = i: GoTo 120
       Next i
       
    'AQLSampleSiz
      
120    Scode = LG(grp, lvl)
       For i = 1 To 16
          If Scode = CD(i) Then Hgrp = i: GoTo 125
       Next i
       
125    For k = Y To 2  'k AQLMa,Mi
        'AQL
         Cn = 0
         For i = 1 To 16
           If AQL(k) = AQ(i) Then Agrp = i: GoTo 130
           If AQL(k) <> AQ(i) And AQL(k) <> 0 Then Cn = Cn + 1 'AQL
         Next i
 

         If Cn = 16 Then
           If FX = 1 Then
                FX = 3       'AQL
                .Cell(4, 4).Range.Text = "-"
               .Cell(4, 5).Range.Text = "-"
               .Cell(4, 6).Range.Text = "-"
               .Cell(5, 4).Range.Text = ss * ComItms
               .Cell(5, 5).Range.Text = ss * ComItms
               .Cell(5, 6).Range.Text = ss * ComItms
                GoTo 170
             Else
               MsgBox "2.Workmanship" & UText(34920, 26684, 20013, 36755, 20837, 30340) & "AQL" & UText(20540, 19981, 22312, 31243, 24207, 35268, 23450, 33539, 22260, 20869) & "，" & UText(35831, 20462, 25913, 25110, 25163, 24037, 22788, 29702) & "！" & CHR(10) & _
                "2.Workmanship table, input AQL value out of the program data scope, please revised it or dealing with it by hand！", vbCritical: Exit Sub
            End If
       End If
       
130     If MV(Hgrp, Agrp) > 0 Then
             If MV(Hgrp, Agrp) >= SQty Then SmpS(k) = SQty
             If MV(Hgrp, Agrp) < SQty Then SmpS(k) = MV(Hgrp, Agrp)
         End If
        
         If MV(Hgrp, Agrp) > 0 And FX = 1 Then
            If MV(Hgrp, Agrp) > ss Then     'sQty
              FX = 4                     'fixed sample
               .Cell(4, 4).Range.Text = "-"
               .Cell(4, 5).Range.Text = "-"
               .Cell(4, 6).Range.Text = "-"
               .Cell(5, 4).Range.Text = ss * ComItms
               .Cell(5, 5).Range.Text = ss * ComItms
               .Cell(5, 6).Range.Text = ss * ComItms
             GoTo 170
            End If
         End If
        If MV(Hgrp, Agrp) = 0 Then SmpS(k) = CD(Hgrp)
        AAc(k) = AC(Hgrp, Agrp)
 '-------，，
        If Styp = 1 Or Styp = 2 Or Styp = 4 Then
             RRe(k) = AAc(k) + 1
        Else
            AAc2(k) = AC2(Hgrp, Agrp)
            RRe(k) = Re(Hgrp, Agrp)
            RRe2(k) = AC2(Hgrp, Agrp) + 1
        End If
        If AAc(k) = 0 Then RRe(k) = 1
        If AAc2(k) = "" Then AAc2(k) = "-"
        If RRe2(k) = "" Then RRe2(k) = "-"
   
140    If SmpS(k) > Ass Then Ass = SmpS(k)
        If AAc(k) = 0 And RRe(k) = 1 Then
           AAc2(k) = "-"
           RRe2(k) = "-"
        End If
150   Next k
      
      If Acr = 0 Then
          'SmpS(0) = Ass
          If Ass >= Scode Then
             SmpS(0) = Ass
          Else
            SmpS(0) = Scode
         Ass = Scode
          End If
          AAc(0) = 0
          RRe(0) = 1
          AAc2(0) = "-"
          RRe2(0) = "-"
      End If
  '------------------------

  
       DAAc(tt, 0) = AAc(0)
       DAAc(tt, 1) = AAc(1)
       DAAc(tt, 2) = AAc(2)
       DSmpS(tt, 0) = SmpS(0)
       DSmpS(tt, 1) = SmpS(1)
       DSmpS(tt, 2) = SmpS(2)
       

 
       'Sample size
       
       
        If Ass <> ComSm(tt) And Styp = 1 Then msg = msg + "Item " & ComItx(tt) & ",1.Quantity" & UText(34920, 26684, 22635, 20837, 30340, 25277, 26679, 25968, 26159) & ComSm(tt) & " " & UText(20294, 26159, 25353) & "ISO2859" & UText(26631, 20934, 26597, 20986, 30340) & "Sample Size" & UText(24212, 35813, 26159) & ":" & Ass & CHR(10) & _
         "Item " & ComItx(tt) & ", the sample size in 1.quantity table is " & ComSm(tt) & "however according to ISO2859.1 it shall be " & Ass & CHR(10)
  Next tt
     
     If msg <> "" Then MsgBox msg & " " & UText(35831, 36827, 34892, 20462, 25913, 21518, 20877, 27425, 36816, 34892, 31243, 24207) & "!！！" & "Pease revise and run progaram again!!!": GoTo 1000

      .Cell(5, 4).Range.Text = SmpS(0)
      .Cell(5, 5).Range.Text = SmpS(1)
      .Cell(5, 6).Range.Text = SmpS(2)
       
       If FX = 1 Then
           Set Cell = .Cell(3, 2)
           Cell.Range.Delete
           Set Occ = Cell.Range.ContentControls.add(wdContentControlComboBox)
        
           With Occ
            .Title = "Sampling Standart"
            .Tag = "Sampling Standart"
             For i = LBound(SamplingStd_Name) To UBound(SamplingStd_Name)
              .DropdownListEntries.add SamplingStd_Name(i), SamplingStd_Value(i)
             Next i
             .DropdownListEntries(8).Select   '"STD-019-HQT"
           End With
         
       End If
     
     
     
    

170   '-------- '
 

        With ThisDocument.Range.Tables(tbn + 1)
        Rown = .Rows.count
        Coln = .Columns.count
        N = 0
        For i = Rown - 2 To 2 Step -1
           If .Rows(i).Cells.count > 1 And .Cell(i, 1).Range.Fields.count = 0 Then .Rows(i).Delete
        Next i
        Rown = .Rows.count
        For i = Rown - 2 To 2 Step -1
           If .Rows(i).Cells.count = 1 Then
              .Rows(i).Select
              Selection.Copy
              .Rows(2).Select
              Selection.Paste
              Exit For
           End If
        Next i
         Rown = .Rows.count
          For i = Rown - 2 To 3 Step -1
              If .Rows(i).Cells.count = 1 Then .Rows(i).Delete
          Next i
          Rown = .Rows.count
          If Rown < 4 Then
             MsgBox "The row for input defect details was deleted,please recover it from correct template" & CHR(10) & _
            UText(29992, 20110, 36755, 20837, 32570, 38519, 30340, 34892, 34987, 21024, 38500, 20102) & "，" & UText(35831, 20808, 20174, 27491, 30830, 30340, 27169, 26495, 20013, 24674, 22797) & "！"
             Exit Sub
          ElseIf Rown = 5 Then
             .Rows(5).Select
              Selection.Copy
              .Rows(5).Select
              Selection.Paste
             
          ElseIf Rown > 6 Then
            For i = Rown - 2 To 6 Step -1
              .Rows(i).Delete
            Next i
          End If
        
           'If .Rows(i).Range.Cells.count = 1 Then
            '  .cell(i, 1).Range.Text = "Item No."
          ' ElseIf .Rows(i).Range.Cells.count > 4 Then
             ' Set rg = ThisDocument.Range(Start:=.cell(i, 2).Range.Start, End:=.cell(i, Coln).Range.End)
              'rg.Delete
           'Else
             ' Set rg = ThisDocument.Range(Start:=.cell(i, 1).Range.Start, End:=.cell(i, Coln).Range.End)
              'rg.Delete
          ' End If
        ' Next i
         'Set rg = ThisDocument.Range(Start:=.Rows(2).Range.Start, End:=.Rows(Rown - 2).Range.End)
         'rg.Delete
         'Set rg = ThisDocument.Range(Start:=.cell(2, 1).Range.Start, End:=.cell(Rown - 2, 1).Range.End)
         'rg.Shading.BackgroundPatternColor = wdColorWhite
         'rg.ParagraphFormat.Alignment = wdAlignParagraphLeft
          Rown = .Rows.count
           For i = 2 To Rown - 2
                If .Rows(i).Cells.count = 1 And LCase(.Cell(i, 1).Range.Text) Like "*item*" = False Then
               .Cell(i, 1).Range.Text = "Item No."
                End If
                If .Rows(i).Cells.count = 1 And LCase(.Cell(i, 1).Range.Text) Like "*item*" = True Then
                   Im(N + 1) = i   'Item
                     N = N + 1       'item
                End If
            Next i
            
      '-----------------
            If N = 0 Then MsgBox "2. Workmanship" & UText(34920, 19979, 30340) & "Defec" & UText(34920, 27809, 26377, 19968, 25972, 34892, 29992, 20110, 22635, 20889) & "item" & UText(21495, 30340, 34892) & "，" & UText(35831, 22686, 21152, 21518, 20877, 36816, 34892, 31243, 24207) & "。" & CHR(13) & _
                  "The Defect tablle has is not any row for inputing Item No.,please revise it and run program again！": Exit Sub

            If N > ComItms Then
                  Set rg = ThisDocument.Range(Start:=.Rows(Im(ComItms + 1)).Range.Start, End:=.Rows(Rown - 2).Range.End)
                  rg.Select
                  Selection.Rows.Delete
            ElseIf N < ComItms And N > 1 Then
                   Set rg = ThisDocument.Range(Start:=.Rows(Im(1)).Range.Start, End:=.Rows(Im(2) - 1).Range.End)
                   rg.Copy
                   For j = 1 To ComItms - N
                    
                      
                     .Rows(Im(1)).Select
                       StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
                     Selection.Paste
                  Next j
            Else
                    Set rg = ThisDocument.Range(Start:=.Rows(Im(1)).Range.Start, End:=.Rows(Rown - 2).Range.End)
                    rg.Copy
                    For j = 1 To ComItms - N
                     .Rows(Im(1)).Select
                        StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
                       Selection.Paste
                    Next j
            End If
                  Rown = .Rows.count
                  N = 0
                  For i = 2 To Rown
                     If .Rows(i).Cells.count = 1 And LCase(.Cell(i, 1).Range.Text) Like "*item*" = True Then
                      N = N + 1
                     .Cell(i, 1).Range.Text = "Item No. " & ComItx(N) & "," & "Sample Size= " & ComSm(N) & " " & Unit & "s"
                     End If
                 Next i
                
                   Rown = .Rows.count
                   
                ' .Cell(Rown, Coln - 2).Range.Text = DAAc(N, 0)
                ' .Cell(Rown, Coln - 1).Range.Text = DAAc(N, 1)
                ' .Cell(Rown, Coln).Range.Text = DAAc(N, 2)
                   Set rg = ThisDocument.Range(Start:=.Rows(Rown - 1).Range.Start, End:=.Rows(Rown).Range.End)
                   rg.Copy
               
                   For i = Rown - 2 To 3 Step -1
                      If .Rows(i).Cells.count = 1 And LCase(.Cell(i, 1).Range.Text) Like "*item*" = True Then
                   
                        .Rows(i).Select
                        time1 = timeGetTime
                         Do
                           DoEvents
                        Loop While timeGetTime - time1 < T
                        Selection.Paste
                      End If
                    Next i
                   
                   Rown = .Rows.count
                   N = 0
                   For i = 3 To Rown
                     If .Rows(i).Cells.count > 1 Then
                       If .Cell(i, 2).Range.Text Like "*Allowed:*" = True Then
                          N = N + 1
                          If FX < 2 Then
                           .Cell(i, Coln - 2).Range.Text = DAAc(N, 0)
                           .Cell(i, Coln - 1).Range.Text = DAAc(N, 1)
                           .Cell(i, Coln).Range.Text = DAAc(N, 2)
                          ElseIf FX > 1 Then
                            .Cell(i, Coln - 2).Range.Text = "-"
                            .Cell(i, Coln - 1).Range.Text = "-"
                           .Cell(i, Coln).Range.Text = "-"
                          End If
                       End If
                     End If
                   Next i
   
    '===================
      Dim colW() As Single
      ReDim colW(1 To Coln)
      Dim TbW As Single
     
      i = 0
      For Each Cel In .Rows(3).Range.Cells
         i = i + 1
         colW(i) = Cel.Width
       Next Cel
      TbW = .Rows(2).Range.Cells(1).Width
      Rown = .Rows.count
              For i = 1 To Rown
                    If .Rows(i).Range.Cells.count = 1 Then
                       .Rows(i).Cells(1).Width = TbW
                    ElseIf .Rows(i).Range.Cells.count = Coln Then
                        For j = 1 To Coln
                          .Rows(i).Range.Cells(j).Width = colW(j)
                        Next j
                    
                    End If
             Next i
     '============================================
     
 End With
      SmpS(0) = 0
      SmpS(1) = 0
      SmpS(2) = 0
     For tt = 1 To ComItms
       SmpS(0) = SmpS(0) + DSmpS(tt, 0)
       SmpS(1) = SmpS(1) + DSmpS(tt, 1)
       SmpS(2) = SmpS(2) + DSmpS(tt, 2)
       
     Next tt
     
             If FX > 1 Then
               .Cell(4, 4).Range.Text = "-"
               .Cell(4, 5).Range.Text = "-"
               .Cell(4, 6).Range.Text = "-"
                Set Cell = .Cell(3, 2)
                Cell.Range.Delete
                Set Occ = Cell.Range.ContentControls.add(wdContentControlComboBox)
        
                With Occ
                 .Title = "Sampling Standart"
                 .Tag = "Sampling Standart"
                  For i = LBound(SamplingStd_Name) To UBound(SamplingStd_Name)
                    .DropdownListEntries.add SamplingStd_Name(i), SamplingStd_Value(i)
                   Next i
                   .DropdownListEntries(2).Select   '"N/A"
                 End With
                 
                 Set Cell = .Cell(5, 2)
                   Cell.Range.Delete
                   Set Occ = Cell.Range.ContentControls.add(wdContentControlComboBox)
                   With Occ
                    .Title = "Inspection Level"
                    .Tag = "Inspection Level"
                      For j = LBound(InspectionLevel_Name) To UBound(InspectionLevel_Name)
                      .DropdownListEntries.add InspectionLevel_Name(j), InspectionLevel_Value(j)
                      Next j
                     .DropdownListEntries(12).Select  '"Fixed Sample Size per Item"
                    
                                                
                 End With
             End If
             If FX < 2 Then
               .Cell(5, 4).Range.Text = SmpS(0)
               .Cell(5, 5).Range.Text = SmpS(1)
               .Cell(5, 6).Range.Text = SmpS(2)
                Set Cell = .Cell(5, 2)
                   Cell.Range.Delete
                   Set Occ = Cell.Range.ContentControls.add(wdContentControlComboBox)
                   With Occ
                    .Title = "Inspection Level"
                    .Tag = "Inspection Level"
                      For j = LBound(InspectionLevel_Name) To UBound(InspectionLevel_Name)
                      .DropdownListEntries.add InspectionLevel_Name(j), InspectionLevel_Value(j)
                      Next j
                     .DropdownListEntries(12).Select  '"Fixed Sample Size per Item"
                 End With
               
             End If
     
     
     
      '-------------------
    
    End Select
       
       
    End With
    
 
     
   If msg <> "" Then
      MsgBox msg & CHR(10) & "Completed!" & CHR(10) & UText(35745, 31639, 23436, 25104) & "！"
    Else
      MsgBox "Completed!" & CHR(10) & UText(35745, 31639, 23436, 25104) & "！"
    End If

    
1000 End Sub


Sub SingelNormal(CD(), AC(), AC2(), Re(), MV() As Integer)
'by wyn single, normal data

        CD(0) = 0
        CD(1) = 2
        CD(2) = 3
        CD(3) = 5:
        CD(4) = 8:
        CD(5) = 13:
        CD(6) = 20
        CD(7) = 32
        CD(8) = 50
        CD(9) = 80
        CD(10) = 125
        CD(11) = 200
        CD(12) = 315
        CD(13) = 500
        CD(14) = 800
        CD(15) = 1250
        CD(16) = 2000
        CD(17) = 3150
     

       AC(0, 0) = CD(0)
       AC(0, 1) = CD(0):       AC(0, 2) = CD(0):   AC(0, 3) = CD(0):   AC(0, 4) = CD(0):   AC(0, 5) = CD(0):
       '---------------------------------------------
       AC(1, 1) = 0:            AC(1, 2) = 0:        AC(1, 3) = 0:        AC(1, 4) = 0:        AC(1, 5) = 0:
       MV(1, 1) = CD(15):       MV(1, 2) = CD(14):   MV(1, 3) = CD(13):   MV(1, 4) = CD(12):   MV(1, 5) = CD(11):
       AC(2, 1) = 0:            AC(2, 2) = 0:        AC(2, 3) = 0:        AC(2, 4) = 0:        AC(2, 5) = 0:
       MV(2, 1) = CD(15):       MV(2, 2) = CD(14):   MV(2, 3) = CD(13):   MV(2, 4) = CD(12):   MV(2, 5) = CD(11):
       AC(3, 1) = 0:            AC(3, 2) = 0:        AC(3, 3) = 0:        AC(3, 4) = 0:        AC(3, 5) = 0:
       MV(3, 1) = CD(15):       MV(3, 2) = CD(14):   MV(3, 3) = CD(13):   MV(3, 4) = CD(12):   MV(3, 5) = CD(11):
       AC(4, 1) = 0:            AC(4, 2) = 0:        AC(4, 3) = 0:        AC(4, 4) = 0:        AC(4, 5) = 0:
       MV(4, 1) = CD(15):       MV(4, 2) = CD(14):   MV(4, 3) = CD(13):   MV(4, 4) = CD(12):   MV(4, 5) = CD(11):
       AC(5, 1) = 0:            AC(5, 2) = 0:        AC(5, 3) = 0:        AC(5, 4) = 0:        AC(5, 5) = 0:
       MV(5, 1) = CD(15):       MV(5, 2) = CD(14):   MV(5, 3) = CD(13):   MV(5, 4) = CD(12):   MV(5, 5) = CD(11):
       AC(6, 1) = 0:            AC(6, 2) = 0:        AC(6, 3) = 0:        AC(6, 4) = 0:        AC(6, 5) = 0:
       MV(6, 1) = CD(15):       MV(6, 2) = CD(14):   MV(6, 3) = CD(13):   MV(6, 4) = CD(12):   MV(6, 5) = CD(11):
       AC(7, 1) = 0:            AC(7, 2) = 0:        AC(7, 3) = 0:        AC(7, 4) = 0:        AC(7, 5) = 0:
       MV(7, 1) = CD(15):       MV(7, 2) = CD(14):   MV(7, 3) = CD(13):   MV(7, 4) = CD(12):   MV(7, 5) = CD(11):
       AC(8, 1) = 0:            AC(8, 2) = 0:        AC(8, 3) = 0:        AC(8, 4) = 0:        AC(8, 5) = 0:
       MV(8, 1) = CD(15):       MV(8, 2) = CD(14):   MV(8, 3) = CD(13):   MV(8, 4) = CD(12):   MV(8, 5) = CD(11):
       AC(9, 1) = 0:            AC(9, 2) = 0:        AC(9, 3) = 0:        AC(9, 4) = 0:        AC(9, 5) = 0:
       MV(9, 1) = CD(15):       MV(9, 2) = CD(14):   MV(9, 3) = CD(13):   MV(9, 4) = CD(12):   MV(9, 5) = CD(11):
       AC(10, 1) = 0:           AC(10, 2) = 0:       AC(10, 3) = 0:       AC(10, 4) = 0:       AC(10, 5) = 0:
       MV(10, 1) = CD(15):      MV(10, 2) = CD(14):  MV(10, 3) = CD(13):  MV(10, 4) = CD(12):  MV(10, 5) = CD(11):
       AC(11, 1) = 0:           AC(11, 2) = 0:       AC(11, 3) = 0:       AC(11, 4) = 0:       AC(11, 5) = 0:
       MV(11, 1) = CD(15):      MV(11, 2) = CD(14):  MV(11, 3) = CD(13):  MV(11, 4) = CD(12):
       AC(12, 1) = 0:           AC(12, 2) = 0:       AC(12, 3) = 0:       AC(12, 4) = 0:       AC(12, 5) = 0:
       MV(12, 1) = CD(15):      MV(12, 2) = CD(14):  MV(12, 3) = CD(13):                       MV(12, 5) = CD(11):
       AC(13, 1) = 10:          AC(13, 2) = 0:       AC(13, 3) = 0:       AC(13, 4) = 0:       AC(13, 5) = 1:
       MV(13, 1) = CD(15):      MV(13, 2) = CD(14):                       MV(13, 4) = CD(12):  MV(13, 5) = CD(14):
       AC(14, 1) = 0:           AC(14, 2) = 0:       AC(14, 3) = 0:       AC(14, 4) = 1:       AC(14, 5) = 1:
       MV(14, 1) = CD(15):                           MV(14, 3) = CD(13):  MV(14, 4) = CD(15):
       AC(15, 1) = 0:           AC(15, 2) = 0:       AC(15, 3) = 1:       AC(15, 4) = 1:       AC(15, 5) = 2:
                                MV(15, 2) = CD(14):  MV(15, 3) = CD(16):
       AC(16, 1) = 0:           AC(16, 2) = 0:       AC(16, 3) = 1:       AC(16, 4) = 2:       AC(16, 5) = 3:
       MV(16, 1) = CD(15):      MV(16, 2) = CD(14)
       '----------------------------------------------
       AC(1, 6) = 0:            AC(1, 7) = 0:        AC(1, 8) = 0:        AC(1, 9) = 0:        AC(1, 10) = 0:
       MV(1, 6) = CD(10):       MV(1, 7) = CD(9):    MV(1, 8) = CD(8):    MV(1, 9) = CD(7):    MV(1, 10) = CD(6):
       AC(2, 6) = 0:            AC(2, 7) = 0:        AC(2, 8) = 0:        AC(2, 9) = 0:        AC(2, 10) = 0:
       MV(2, 6) = CD(10):       MV(2, 7) = CD(9):    MV(2, 8) = CD(8):    MV(2, 9) = CD(7):    MV(2, 10) = CD(6):
       AC(3, 6) = 0:            AC(3, 7) = 0:        AC(3, 8) = 0:        AC(3, 9) = 0:        AC(3, 10) = 0:
       MV(3, 6) = CD(10):       MV(3, 7) = CD(9):    MV(3, 8) = CD(8):    MV(3, 9) = CD(7):    MV(3, 10) = CD(6):
       AC(4, 6) = 0:            AC(4, 7) = 0:        AC(4, 8) = 0:        AC(4, 9) = 0:        AC(4, 10) = 0:
       MV(4, 6) = CD(10):       MV(4, 7) = CD(9):    MV(4, 8) = CD(8):    MV(4, 9) = CD(7):    MV(4, 10) = CD(6):
       AC(5, 6) = 0:            AC(5, 7) = 0:        AC(5, 8) = 0:        AC(5, 9) = 0:        AC(5, 10) = 0:
       MV(5, 6) = CD(10):       MV(5, 7) = CD(9):    MV(5, 8) = CD(8):    MV(5, 9) = CD(7):    MV(5, 10) = CD(6):
       AC(6, 6) = 0:            AC(6, 7) = 0:        AC(6, 8) = 0:        AC(6, 9) = 0:        AC(6, 10) = 0:
       MV(6, 6) = CD(10):       MV(6, 7) = CD(9):    MV(6, 8) = CD(8):    MV(6, 9) = CD(7):
       AC(7, 6) = 0:            AC(7, 7) = 0:        AC(7, 8) = 0:        AC(7, 9) = 0:        AC(7, 10) = 0:
       MV(7, 6) = CD(10):       MV(7, 7) = CD(9):    MV(7, 8) = CD(8):                         MV(7, 10) = CD(6):
       AC(8, 6) = 0:            AC(8, 7) = 0:        AC(8, 8) = 0:        AC(8, 9) = 0:        AC(8, 10) = 1:
       MV(8, 6) = CD(10):       MV(8, 7) = CD(9):                         MV(8, 9) = CD(7):    MV(8, 10) = CD(9):
       AC(9, 6) = 0:            AC(9, 7) = 0:        AC(9, 8) = 0:        AC(9, 9) = 1:        AC(9, 10) = 1:
       MV(9, 6) = CD(10):                            MV(9, 8) = CD(8):    MV(9, 9) = CD(10):
       AC(10, 6) = 0:           AC(10, 7) = 0:       AC(10, 8) = 1:       AC(10, 9) = 1:       AC(10, 10) = 2:
                                MV(10, 7) = CD(9):   MV(10, 8) = CD(11):
       AC(11, 6) = 0:           AC(11, 7) = 1:       AC(11, 8) = 1:       AC(11, 9) = 2:       AC(11, 10) = 3:
       MV(11, 6) = CD(10):      MV(11, 7) = CD(12):
       AC(12, 6) = 1:           AC(12, 7) = 1:       AC(12, 8) = 2:       AC(12, 9) = 3:       AC(12, 10) = 5:
       MV(12, 6) = CD(13):
       AC(13, 6) = 1:           AC(13, 7) = 2:       AC(13, 8) = 3:       AC(13, 9) = 5:       AC(13, 10) = 7:
                                                                                              
       AC(14, 6) = 2:           AC(14, 7) = 3:       AC(14, 8) = 5:       AC(14, 9) = 7:       AC(14, 10) = 10:
                                                                          
       AC(15, 6) = 3:           AC(15, 7) = 5:       AC(15, 8) = 7:       AC(15, 9) = 10:       AC(15, 10) = 14:
                                                      
       AC(16, 6) = 5:           AC(16, 7) = 7:       AC(16, 8) = 10:      AC(16, 9) = 14:       AC(16, 10) = 21::
       
       
       '----------------------------------------------
       AC(1, 11) = 0:           AC(1, 12) = 0:       AC(1, 13) = 0:       AC(1, 14) = 0:       AC(1, 15) = 0:
       MV(1, 11) = CD(5):       MV(1, 12) = CD(4):   MV(1, 13) = CD(3):   MV(1, 14) = CD(2):
       AC(2, 11) = 0:           AC(2, 12) = 0:       AC(2, 13) = 0:       AC(2, 14) = 0:       AC(2, 15) = 0:
       MV(2, 11) = CD(5):       MV(2, 12) = CD(4):   MV(2, 13) = CD(3):                        MV(2, 15) = CD(1):
       AC(3, 11) = 0:           AC(3, 12) = 0:       AC(3, 13) = 0:       AC(3, 14) = 0:       AC(3, 15) = 1:
       MV(3, 11) = CD(5):       MV(3, 12) = CD(4):                        MV(3, 14) = CD(2):   MV(3, 15) = CD(4):
       AC(4, 11) = 0:           AC(4, 12) = 0:       AC(4, 13) = 0:       AC(4, 14) = 1:       AC(4, 15) = 1:
       MV(4, 11) = CD(5):                            MV(4, 13) = CD(3):   MV(4, 14) = CD(5):
       AC(5, 11) = 0:           AC(5, 12) = 0:       AC(5, 13) = 1:       AC(5, 14) = 1:       AC(5, 15) = 2:
                                MV(5, 12) = CD(4):   MV(5, 13) = CD(6):
       AC(6, 11) = 0:           AC(6, 12) = 1:       AC(6, 13) = 1:       AC(6, 14) = 2:       AC(6, 15) = 3:
       MV(6, 11) = CD(5):       MV(6, 12) = CD(7):
       AC(7, 11) = 1:           AC(7, 12) = 1:       AC(7, 13) = 2:       AC(7, 14) = 3:       AC(7, 15) = 5:
       MV(7, 11) = CD(8)
       AC(8, 11) = 1:           AC(8, 12) = 2:       AC(8, 13) = 3:       AC(8, 14) = 5:       AC(8, 15) = 7:
        
       AC(9, 11) = 2:           AC(9, 12) = 3:       AC(9, 13) = 5:       AC(9, 14) = 7:       AC(9, 15) = 10:
       
       AC(10, 11) = 3:          AC(10, 12) = 5:      AC(10, 13) = 7:      AC(10, 14) = 10:     AC(10, 15) = 14:
       
       AC(11, 11) = 5:          AC(11, 12) = 7:      AC(11, 13) = 10:     AC(11, 14) = 14:     AC(11, 15) = 21:
        
       AC(12, 11) = 7:          AC(12, 12) = 10:     AC(12, 13) = 14:     AC(12, 14) = 21:     AC(12, 15) = 21:
                                                                                               MV(12, 15) = CD(11):
       AC(13, 11) = 10:         AC(13, 12) = 14:     AC(13, 13) = 21:     AC(13, 14) = 21:     AC(13, 15) = 21:
                                                                          MV(13, 14) = CD(12): MV(13, 15) = CD(11):
       AC(14, 11) = 14:         AC(14, 12) = 21:     AC(14, 13) = 21:     AC(14, 14) = 21:     AC(14, 15) = 21:
                                                     MV(14, 13) = CD(13): MV(14, 14) = CD(12): MV(14, 15) = CD(11):
       AC(15, 11) = 21:         AC(15, 12) = 21:     AC(15, 13) = 21:     AC(15, 14) = 21:     AC(15, 15) = 21:
                                MV(15, 12) = CD(14): MV(15, 13) = CD(13): MV(15, 14) = CD(12): MV(15, 15) = CD(11):
       AC(16, 11) = 21:         AC(16, 12) = 21:     AC(16, 13) = 21:     AC(16, 14) = 21:     AC(16, 15) = 21:
       MV(16, 11) = CD(15):     MV(16, 12) = CD(14): MV(16, 13) = CD(13): MV(16, 14) = CD(12): MV(16, 15) = CD(11):
           
       '----------------------------------------------
       AC(1, 16) = 1:           AC(1, 17) = 1:       AC(1, 18) = 1:       AC(1, 19) = 2:       AC(1, 20) = 3:
       MV(1, 16) = CD(3):       MV(1, 17) = CD(2):
       AC(2, 16) = 1:           AC(2, 17) = 1:       AC(2, 18) = 2:       AC(2, 19) = 3:       AC(2, 20) = 5:
       MV(2, 16) = CD(3):
       AC(3, 16) = 1:           AC(3, 17) = 2:       AC(3, 18) = 3:       AC(3, 19) = 5:       AC(3, 20) = 7:
                                                                                               
       AC(4, 16) = 2:           AC(4, 17) = 3:       AC(4, 18) = 5:       AC(4, 19) = 7:       AC(4, 20) = 10:
                                                                          
       AC(5, 16) = 3:           AC(5, 17) = 5:       AC(5, 18) = 7:       AC(5, 19) = 10:       AC(5, 20) = 14:
                                                       
       AC(6, 16) = 5:           AC(6, 17) = 7:       AC(6, 18) = 10:      AC(6, 19) = 14:       AC(6, 20) = 21:
                                 
       AC(7, 16) = 7:           AC(7, 17) = 10:      AC(7, 18) = 14:      AC(7, 19) = 21:       AC(7, 20) = 21:
                                                                                                MV(7, 20) = CD(6)
       AC(8, 16) = 10:          AC(8, 17) = 14:      AC(8, 18) = 21:      AC(8, 19) = 21:       AC(8, 20) = 21:
                                                                          MV(8, 19) = CD(7):    MV(8, 20) = CD(6)
       AC(9, 16) = 14:          AC(9, 17) = 21:      AC(9, 18) = 21:      AC(9, 19) = 21:       AC(9, 20) = 21:
                                                     MV(9, 18) = CD(8):   MV(9, 19) = CD(7):    MV(9, 20) = CD(6)
       AC(10, 16) = 21:         AC(10, 17) = 21:     AC(10, 18) = 21:     AC(10, 19) = 21:      AC(10, 20) = 21:
                                MV(10, 17) = CD(9):  MV(10, 18) = CD(8):  MV(10, 19) = CD(7):   MV(10, 20) = CD(6)
       AC(11, 16) = 21:         AC(11, 17) = 21:     AC(11, 18) = 21:     AC(11, 19) = 21:      AC(11, 20) = 21:
       MV(11, 16) = CD(10):     MV(11, 17) = CD(9):  MV(11, 18) = CD(8):  MV(11, 19) = CD(7):   MV(11, 20) = CD(6)
       AC(12, 16) = 21:         AC(12, 17) = 21:     AC(12, 18) = 21:     AC(12, 19) = 21:      AC(12, 20) = 21:
       MV(12, 16) = CD(10):     MV(12, 17) = CD(9):  MV(12, 18) = CD(8):  MV(12, 19) = CD(7):   MV(12, 20) = CD(6):
       AC(13, 16) = 21:         AC(13, 17) = 21:     AC(13, 18) = 21:     AC(13, 19) = 21:      AC(13, 20) = 21:
       MV(13, 16) = CD(10):     MV(13, 17) = CD(9):  MV(13, 18) = CD(8):  MV(13, 19) = CD(7):   MV(13, 20) = CD(6):
       AC(14, 16) = 21:         AC(14, 17) = 21:     AC(14, 18) = 21:     AC(14, 19) = 21:      AC(14, 20) = 21:
       MV(14, 16) = CD(10):     MV(14, 17) = CD(9):  MV(14, 18) = CD(8):  MV(14, 19) = CD(7):   MV(14, 20) = CD(6):
       AC(15, 16) = 21:         AC(15, 17) = 21:     AC(15, 18) = 21:     AC(15, 19) = 21:      AC(15, 20) = 21:
       MV(15, 16) = CD(10):     MV(15, 17) = CD(9):  MV(15, 18) = CD(8):  MV(15, 19) = CD(7):   MV(15, 20) = CD(6):
       AC(16, 16) = 21:         AC(16, 17) = 21:     AC(16, 18) = 21:     AC(16, 19) = 21:      AC(16, 20) = 21:
       MV(16, 16) = CD(10):     MV(16, 17) = CD(9):  MV(16, 18) = CD(8):  MV(16, 19) = CD(7):   MV(16, 20) = CD(6):
   
        '----------------------------------------------
       AC(1, 21) = 5:            AC(1, 22) = 7:        AC(1, 23) = 10:      AC(1, 24) = 14:       AC(1, 25) = 21:
                            
       AC(2, 21) = 7:            AC(2, 22) = 10:       AC(2, 23) = 14:      AC(2, 24) = 21:       AC(2, 25) = 30:
       
       AC(3, 21) = 10:           AC(3, 22) = 14:       AC(3, 23) = 21:      AC(3, 24) = 30:       AC(3, 25) = 44:
                                                                                               
       AC(4, 21) = 14:           AC(4, 22) = 21:       AC(4, 23) = 30:      AC(4, 24) = 44:       AC(4, 25) = 44:
                                                                                                  MV(4, 25) = CD(3)
       AC(5, 21) = 21:           AC(5, 22) = 30:       AC(5, 23) = 44:      AC(5, 24) = 44:       AC(5, 25) = 44:
                                                                            MV(5, 24) = CD(4):    MV(5, 25) = CD(3)
       AC(6, 21) = 21:           AC(6, 22) = 30:       AC(6, 23) = 44:      AC(6, 24) = 44:       AC(6, 25) = 44:
       MV(6, 21) = CD(5):        MV(6, 22) = CD(5):    MV(6, 23) = CD(5):   MV(6, 24) = CD(4):    MV(6, 25) = CD(3)
       AC(7, 21) = 21:           AC(7, 22) = 30:       AC(7, 23) = 44:      AC(7, 24) = 44:       AC(7, 25) = 44:
       MV(7, 21) = CD(5):        MV(7, 22) = CD(5):    MV(7, 23) = CD(5):   MV(7, 24) = CD(4):    MV(7, 25) = CD(3)
       AC(8, 21) = 21:           AC(8, 22) = 30:       AC(8, 23) = 44:      AC(8, 24) = 44:       AC(8, 25) = 44:
       MV(8, 21) = CD(5):        MV(8, 22) = CD(5):    MV(8, 23) = CD(5):   MV(8, 24) = CD(4):    MV(8, 25) = CD(3)
       AC(9, 21) = 21:           AC(9, 22) = 30:       AC(9, 23) = 44:      AC(9, 24) = 44:       AC(9, 25) = 44:
       MV(9, 21) = CD(5):        MV(9, 22) = CD(5):    MV(9, 23) = CD(5):   MV(9, 24) = CD(4):    MV(9, 25) = CD(3)
       AC(10, 21) = 21:          AC(10, 22) = 30:      AC(10, 23) = 44:     AC(10, 24) = 44:      AC(10, 25) = 44:
       MV(10, 21) = CD(5):       MV(10, 22) = CD(5):   MV(10, 23) = CD(5):  MV(10, 24) = CD(4):   MV(10, 25) = CD(3)
       AC(11, 21) = 21:          AC(11, 22) = 30:      AC(11, 23) = 44:     AC(11, 24) = 44:      AC(11, 25) = 44:
       MV(11, 21) = CD(5):       MV(11, 22) = CD(5):   MV(11, 23) = CD(5):  MV(11, 24) = CD(4):   MV(11, 25) = CD(3)
       AC(12, 21) = 21:          AC(12, 22) = 30:      AC(12, 23) = 44:     AC(12, 24) = 44:      AC(12, 25) = 44:
       MV(12, 21) = CD(5):       MV(12, 22) = CD(5):   MV(12, 23) = CD(5):  MV(12, 24) = CD(4):   MV(12, 25) = CD(3):
       AC(13, 21) = 21:          AC(13, 22) = 30:      AC(13, 23) = 44:     AC(13, 24) = 44:      AC(13, 25) = 44:
       MV(13, 21) = CD(5):       MV(13, 22) = CD(5):   MV(13, 23) = CD(5):  MV(13, 24) = CD(4):   MV(13, 25) = CD(3):
       AC(14, 21) = 21:          AC(14, 22) = 30:      AC(14, 23) = 44:     AC(14, 24) = 44:      AC(14, 25) = 44:
       MV(14, 21) = CD(5):       MV(14, 22) = CD(5):   MV(14, 23) = CD(5):  MV(14, 24) = CD(4):   MV(14, 25) = CD(3):
       AC(15, 21) = 21:          AC(15, 22) = 30:      AC(15, 23) = 44:     AC(15, 24) = 44:      AC(15, 25) = 44:
       MV(15, 21) = CD(5):       MV(15, 22) = CD(5):   MV(15, 23) = CD(5):  MV(15, 24) = CD(4):   MV(15, 25) = CD(3):
       AC(16, 21) = 21:          AC(16, 22) = 30:      AC(16, 23) = 44:     AC(16, 24) = 44:      AC(16, 25) = 44:
       MV(16, 21) = CD(5):       MV(16, 22) = CD(5):   MV(16, 23) = CD(5):  MV(16, 24) = CD(4):   MV(16, 25) = CD(3):
   
       '----------------------------------------------
       AC(1, 26) = 30:
                            
       AC(2, 26) = 44:
       
       AC(3, 26) = 44:
       MV(3, 26) = CD(2)
       AC(4, 26) = 44:
       MV(4, 26) = CD(2)
       AC(5, 26) = 44:
       MV(5, 26) = CD(2)
       AC(6, 26) = 44:
       MV(6, 26) = CD(2)
       AC(7, 26) = 44:
       MV(7, 26) = CD(2)
       AC(8, 26) = 44:
       MV(8, 26) = CD(2)
       AC(9, 26) = 44:
       MV(9, 26) = CD(2)
       AC(10, 26) = 44:
       MV(10, 26) = CD(2)
       AC(11, 26) = 44:
       MV(11, 26) = CD(2)
       AC(12, 26) = 44:
       MV(12, 26) = CD(2):
       AC(13, 26) = 44:
       MV(13, 26) = CD(2):
       AC(14, 26) = 44:
       MV(14, 26) = CD(2):
       AC(15, 26) = 44:
       MV(15, 26) = CD(2):
       AC(16, 26) = 44:
       MV(16, 26) = CD(2):

End Sub



Sub Sap_photo(tt As Integer, Path, Sap_Note_type, msg As String)
'by wyn


Dim PhotoNameN(11, 50), PhotoPath(11), PhotoName(100) As String
Dim i, i3, rnp, cnp, picN(200) As Integer
Dim tb_sap(11), Ltb_sap(11) As String
Dim Tn As Integer
Dim I4(200) As Integer

  tb_sap(1) = "Special Attention Point for Quantity":                       Ltb_sap(1) = LCase(tb_sap(1))
  tb_sap(2) = "Special Attention Point for Workmanship":                    Ltb_sap(2) = LCase(tb_sap(2))
  tb_sap(3) = "Special Attention Point for On Site Test":                   Ltb_sap(3) = LCase(tb_sap(3))
  tb_sap(4) = "Special Attention Point for Product Dimension & Weight":     Ltb_sap(4) = LCase(tb_sap(4))
  tb_sap(5) = "Special Attention Point for Style / Material / Construction":    Ltb_sap(5) = LCase(tb_sap(5))
  tb_sap(6) = "Special Attention Point for Product Color":                  Ltb_sap(6) = LCase(tb_sap(6))
  tb_sap(7) = "Special Attention Point for Product Label / Marking":         Ltb_sap(7) = LCase(tb_sap(7))
  tb_sap(8) = "Special Attention Point for Shipping Mark":                  Ltb_sap(8) = LCase(tb_sap(8))
  tb_sap(9) = "Special Attention Point for Packing / Packaging":             Ltb_sap(9) = LCase(tb_sap(9))
 


Set WD = ThisDocument


   'Quantity sap
        picN(tt) = 0
       
        If Dir(Path & Sap_Note_type, vbDirectory) = Sap_Note_type Then
                    PhotoName(1) = Dir(Path & Sap_Note_type & "\" & "*.*")
                If PhotoName(1) <> "" Then
                    i = 1
                    Do While PhotoName(i) <> ""
                       i = i + 1
                       PhotoName(i) = Dir
                     Loop
                       picN(tt) = i - 1
                    
                        
                     
               End If
       End If
        If picN(tt) >= 10 Then '10
     i = picN(tt)
     i5 = 1
     For X = 1 To i
           PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
           I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
     Next X

     skipnum = Round(i / 2)
     Do While skipnum > 0
        ii = 1 + skipnum
        For j = ii To i
            index = j
            Do
                If index >= (1 + skipnum) Then
                    If i5 = 1 Then
                        If I4(index) < I4(index - skipnum) Then
                            tempnum = PhotoName(index)
                            PhotoName(index) = PhotoName(index - skipnum)
                            PhotoName(index - skipnum) = tempnum
                            i7 = I4(index)
                            I4(index) = I4(index - skipnum)
                            I4(index - skipnum) = i7
                            index = index - skipnum
                        Else
                            Exit Do
                        End If
                    Else
                        If PhotoName(index) < PhotoName(index - skipnum) Then
                            tempnum = PhotoName(index)
                            PhotoName(index) = PhotoName(index - skipnum)
                            PhotoName(index - skipnum) = tempnum
                            index = index - skipnum
                        Else
                            Exit Do
                        End If
                    End If
                Else
                    Exit Do
                End If
          Loop
       Next j
'        skipnum = Int((skipnum - 1) / 2) - 1
        skipnum = Round(skipnum / 2)
      Loop
    End If
       
        If picN(tt) > 0 Then
         tb = WD.Tables.count - 1
          For X = 1 To tb
            If LCase(WD.Tables(X).Cell(1, 1).Range.Text) Like "*" & Ltb_sap(tt) & "*" = True Then
              Tn = X
              GoTo 15
            End If
          Next X

         msg = msg & "There are " & picN(tt) & " photos in " & tb_sap(tt) & " folder, However, not" & tb_sap(tt) & _
         " table was found at report" & CHR(10) & tb_sap(tt) & _
         UText(25991, 20214, 22841, 20013, 26377) & picN(tt) & UText(24352, 22270, 29255) & "，" & UText(20294, 22312, 25253, 21578, 20013, 27809, 26377) & tb_sap(tt) & UText(34920, 26684) & CHR(10)

         GoTo 100

15      If WD.Tables(Tn + 1).Columns.count <> 2 Or WD.Tables(Tn + 1).Rows(1).Range.Cells.count = 1 Then
        Rown = WD.Tables(Tn).Rows.count
        WD.Tables(Tn).Cell(Rown, 1).Select
        Selection.MoveDown
        Selection.TypeParagraph
        WD.Tables.add Range:=Selection.Range, NumRows:=2, NumColumns:=2, _
        DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
            With Selection.Tables(1)
                If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
               End If
            End With
            Call Setborders025
     End If
         
        
                 WD.Tables(Tn + 1).Select
                 Selection.Delete
                  With Selection.Tables(1)
                    Rown = .Rows.count
                       If Rown = 1 Then
                         .Rows(1).Select
                           Selection.InsertRows 1
                        ElseIf Rown > 2 Then
                           For i = Rown To 3 Step -1
                          .Rows(i).Delete
                           Next i
                        End If
                  End With
            
             If picN(tt) <> 0 Then
                WD.Tables(Tn + 1).Select
                With Selection.Tables(1)
                   For i3 = 1 To picN(tt)
                      If i3 / 2 = Round(i3 / 2) Then
                        rnp = i3 - 1
                        cnp = 2
                      Else
                        rnp = i3
                        cnp = 1
                      End If
           
                    PhotoNameN(tt, i3) = Path & Sap_Note_type & "\" & PhotoName(i3)
                    .Cell(rnp, cnp).Range.InlineShapes.AddPicture FileName:= _
                      PhotoNameN(tt, i3), LinkToFile:=False, _
                      SaveWithDocument:=True
                      If .Cell(rnp, cnp).Range.InlineShapes.count > 0 Then
                       .Cell(rnp, cnp).Range.InlineShapes(1).Height = 170
                       .Cell(rnp, cnp).Range.InlineShapes(1).Width = 235
                      End If
                       Do While IsNumeric(Mid(PhotoName(i3), 1, 1)) = True
                         PhotoName(i3) = Mid(PhotoName(i3), 2, Len(PhotoName(i3)))
                       Loop
                       PhotoName(i3) = Trim(PhotoName(i3))
                       PhotoName(i3) = Mid(PhotoName(i3), 1, InStrRev(PhotoName(i3), ".") - 1)
                       PhotoName(i3) = UCase(Left(PhotoName(i3), 1)) & Mid(PhotoName(i3), 2, Len(PhotoName(i3)) - 1)
                       PhotoName(i3) = "SAP " & PhotoName(i3)
                     .Rows(rnp + 1).Height = 0.6
                     .Cell(rnp + 1, cnp).Range.Text = PhotoName(i3)
                     .Cell(rnp + 1, cnp).Range.Font.Bold = False
                     .Cell(rnp + 1, cnp).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                     .Cell(rnp + 1, cnp).VerticalAlignment = wdCellAlignVerticalCenter
                      
                     .Rows(rnp + 1).Select
                       If i3 < picN(tt) Then
                            If i3 / 2 = Round(i3 / 2) Then Selection.InsertRowsBelow 2
                       Else
                            If i3 / 2 <> Round(i3 / 2) Then .Cell(rnp, 2).Range.Text = "NA"
                           .Cell(rnp, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                          .Cell(rnp, 2).VerticalAlignment = wdCellAlignVerticalCenter
                       End If
                
                    Next i3
                  End With
              End If


      End If
        
100 End Sub


Sub Note_photo(tt As Integer, Path, Sap_Note_type, msg As String)
'by wyn


Dim PhotoNameN(11, 50), PhotoPath(11), PhotoName(100) As String
Dim i, i3, rnp, cnp, picN(200) As Integer
Dim tb_no(11), Ltb_no(11) As String
Dim Tn As Integer
Dim I4(200) As Integer
   
  tb_no(1) = "Reference Note for Quantity":                                 Ltb_no(1) = LCase(tb_no(1))
  tb_no(2) = "Reference Note for Workmanship":                              Ltb_no(2) = LCase(tb_no(2))
  tb_no(3) = "Reference Note for On Site Test":                             Ltb_no(3) = LCase(tb_no(3))
  tb_no(4) = "Reference Note for Product Dimension & Weight":               Ltb_no(4) = LCase(tb_no(4))
  tb_no(5) = "Reference Note for Style / Material / Construction":              Ltb_no(5) = LCase(tb_no(5))
  tb_no(6) = "Reference Note for Product Color":                            Ltb_no(6) = LCase(tb_no(6))
  tb_no(7) = "Reference Note for Product Label / Marking":                   Ltb_no(7) = LCase(tb_no(7))
  tb_no(8) = "Reference Note for Shipping Mark":                            Ltb_no(8) = LCase(tb_no(8))
  tb_no(9) = "Reference Note for Packing / Packaging":                       Ltb_no(9) = LCase(tb_no(9))
 


Set WD = ThisDocument


   'Quantity sap
        picN(tt) = 0
       
        If Dir(Path & Sap_Note_type, vbDirectory) = Sap_Note_type Then
                    PhotoName(1) = Dir(Path & Sap_Note_type & "\" & "*.*")
                If PhotoName(1) <> "" Then
                    i = 1
                    Do While PhotoName(i) <> ""
                       i = i + 1
                       PhotoName(i) = Dir
                     Loop
                       picN(tt) = i - 1
                   
                       
                    
               End If
       End If
        If picN(tt) >= 10 Then '10
     i = picN(tt)
     i5 = 1
     For X = 1 To i
           PhotoName1 = Left(PhotoName(X), InStrRev(PhotoName(X), ".") - 1)
           I4(X) = Val(Mid(PhotoName1, 1, InStr(1, PhotoName1, " ") - 1))
     Next X

     skipnum = Round(i / 2)
     Do While skipnum > 0
        ii = 1 + skipnum
        For j = ii To i
            index = j
            Do
                If index >= (1 + skipnum) Then
                    If i5 = 1 Then
                        If I4(index) < I4(index - skipnum) Then
                            tempnum = PhotoName(index)
                            PhotoName(index) = PhotoName(index - skipnum)
                            PhotoName(index - skipnum) = tempnum
                            i7 = I4(index)
                            I4(index) = I4(index - skipnum)
                            I4(index - skipnum) = i7
                            index = index - skipnum
                        Else
                            Exit Do
                        End If
                    Else
                        If PhotoName(index) < PhotoName(index - skipnum) Then
                            tempnum = PhotoName(index)
                            PhotoName(index) = PhotoName(index - skipnum)
                            PhotoName(index - skipnum) = tempnum
                            index = index - skipnum
                        Else
                            Exit Do
                        End If
                    End If
                Else
                    Exit Do
                End If
          Loop
       Next j
'        skipnum = Int((skipnum - 1) / 2) - 1
        skipnum = Round(skipnum / 2)
      Loop
    End If
        If picN(tt) > 0 Then
         tb = WD.Tables.count - 1
          For X = 1 To tb
            If LCase(WD.Tables(X).Cell(1, 1).Range.Text) Like "*" & Ltb_no(tt) & "*" = True Then
              Tn = X
              GoTo 15
            End If
          Next X

         msg = msg & "There are " & picN(tt) & " photos in " & tb_no(tt) & " folder, However, not" & tb_no(tt) & _
         " table was found at report" & CHR(10) & tb_no(tt) & _
         UText(25991, 20214, 22841, 20013, 26377) & picN(tt) & UText(24352, 22270, 29255) & "，" & UText(20294, 22312, 25253, 21578, 20013, 27809, 26377) & tb_no(tt) & UText(34920, 26684) & CHR(10)

         GoTo 100

15      If WD.Tables(Tn + 1).Columns.count <> 2 Or WD.Tables(Tn + 1).Rows(1).Range.Cells.count = 1 Then
        Rown = WD.Tables(Tn).Rows.count
        WD.Tables(Tn).Cell(Rown, 1).Select
        Selection.MoveDown
        Selection.TypeParagraph
        WD.Tables.add Range:=Selection.Range, NumRows:=2, NumColumns:=2, _
        DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
            With Selection.Tables(1)
                If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
               End If
            End With
            Call Setborders025
     End If
         
        
                 WD.Tables(Tn + 1).Select
                 Selection.Delete
                  With Selection.Tables(1)
                    Rown = .Rows.count
                       If Rown = 1 Then
                         .Rows(1).Select
                           Selection.InsertRows 1
                        ElseIf Rown > 2 Then
                           For i = Rown To 3 Step -1
                          .Rows(i).Delete
                           Next i
                        End If
                  End With
            
             If picN(tt) <> 0 Then
                WD.Tables(Tn + 1).Select
                With Selection.Tables(1)
                   For i3 = 1 To picN(tt)
                      If i3 / 2 = Round(i3 / 2) Then
                        rnp = i3 - 1
                        cnp = 2
                      Else
                        rnp = i3
                        cnp = 1
                      End If
           
                     PhotoNameN(tt, i3) = Path & Sap_Note_type & "\" & PhotoName(i3)
                    .Cell(rnp, cnp).Range.InlineShapes.AddPicture FileName:= _
                      PhotoNameN(tt, i3), LinkToFile:=False, _
                      SaveWithDocument:=True
                      If .Cell(rnp, cnp).Range.InlineShapes.count > 0 Then
                         .Cell(rnp, cnp).Range.InlineShapes(1).Height = 170
                         .Cell(rnp, cnp).Range.InlineShapes(1).Width = 235
                      End If
                       Do While IsNumeric(Mid(PhotoName(i3), 1, 1)) = True
                         PhotoName(i3) = Mid(PhotoName(i3), 2, Len(PhotoName(i3)))
                       Loop
                       PhotoName(i3) = Trim(PhotoName(i3))
                       PhotoName(i3) = Mid(PhotoName(i3), 1, InStrRev(PhotoName(i3), ".") - 1)
                       PhotoName(i3) = UCase(Left(PhotoName(i3), 1)) & Mid(PhotoName(i3), 2, Len(PhotoName(i3)) - 1)
                       PhotoName(i3) = "NOTE " & PhotoName(i3)
                     .Rows(rnp + 1).Height = 0.6
                     .Cell(rnp + 1, cnp).Range.Text = PhotoName(i3)
                     .Cell(rnp + 1, cnp).Range.Font.Bold = False
                     .Cell(rnp + 1, cnp).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                     .Cell(rnp + 1, cnp).VerticalAlignment = wdCellAlignVerticalCenter
                     .Rows(rnp + 1).Select
                       If i3 < picN(tt) Then
                            If i3 / 2 = Round(i3 / 2) Then Selection.InsertRowsBelow 2
                       Else
                            If i3 / 2 <> Round(i3 / 2) Then .Cell(rnp, 2).Range.Text = "NA"
                          .Cell(rnp, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                          .Cell(rnp, 2).VerticalAlignment = wdCellAlignVerticalCenter
                       End If
                
                    Next i3
                  End With
              End If


      End If
        
100 End Sub

Sub InsResult()
'by wyn 20241216


Dim tb_tt(12), tb_sap(12), Ltb_sap(12), tb_no(12), Ltb_no(12), BK(12) As String
Dim i, j, k, tb, Restbn(12), Saptbn(12), Notbn(12), Cn, Rn, Ctrn(12), RCtrn(12), SapP(12), Res(12), BG As Integer
Dim Kw, Cw(100), Kcr, Kma, Kmi As String
Dim CF, NC, PD, ch, NA, FI As String
Dim CFN, NCN, PDN, FIN, Celn, SCF(12), SNC(12), SPD(12), NAP(12), INF(12), RI, Lw(100), ST(50), ED(50) As Integer
Dim rg, rg1 As Range
Dim HyperlinkText As String
Dim BKText(100) As String
Dim TLn As Integer

'Call SetTablePosition
Set WD = ThisDocument
Call FindAndHighlightText

  tb_tt(1) = "1. Quantity":                            BK(1) = "Quantity"
  tb_tt(2) = "2. Workmanship":                         BK(2) = "Workmanship"
  tb_tt(3) = "3. On Site Test":                        BK(3) = "OnSiteTest"
  tb_tt(4) = "4. Product Dimension & Weight":          BK(4) = "ProductDimensionWeight"
  tb_tt(5) = "5. Style / Material / Construction":         BK(5) = "StyleMaterialConstruction"
  tb_tt(6) = "6. Product Color":                       BK(6) = "ProductColor"
  tb_tt(7) = "7. Product Label / Marking":              BK(7) = "ProductLabelMarking"
  tb_tt(8) = "8. Shipping Mark":                       BK(8) = "ShippingMark"
  tb_tt(9) = "9. Packing / Packaging":                  BK(9) = "PackingPackaging"
  tb_tt(10) = "10. Other Note":                        BK(10) = "OtherNote"
  tb_tt(11) = "11. Other Photos":                      BK(11) = "OtherPhotos"
  '----------------------------------------------------------
  tb_sap(1) = "Special Attention Point for Quantity":                       Ltb_sap(1) = LCase(tb_sap(1))
  tb_sap(2) = "Special Attention Point for Workmanship":                    Ltb_sap(2) = LCase(tb_sap(2))
  tb_sap(3) = "Special Attention Point for On Site Test":                   Ltb_sap(3) = LCase(tb_sap(3))
  tb_sap(4) = "Special Attention Point for Product Dimension & Weight":     Ltb_sap(4) = LCase(tb_sap(4))
  tb_sap(5) = "Special Attention Point for Style / Material / Construction":    Ltb_sap(5) = LCase(tb_sap(5))
  tb_sap(6) = "Special Attention Point for Product Color":                  Ltb_sap(6) = LCase(tb_sap(6))
  tb_sap(7) = "Special Attention Point for Product Label / Marking":         Ltb_sap(7) = LCase(tb_sap(7))
  tb_sap(8) = "Special Attention Point for Shipping Mark":                  Ltb_sap(8) = LCase(tb_sap(8))
  tb_sap(9) = "Special Attention Point for Packing / Packaging":             Ltb_sap(9) = LCase(tb_sap(9))
  '-----------------------------------------------------------
  tb_no(1) = "Reference Note for Quantity":                                 Ltb_no(1) = LCase(tb_no(1))
  tb_no(2) = "Reference Note for Workmanship":                              Ltb_no(2) = LCase(tb_no(2))
  tb_no(3) = "Reference Note for On Site Test":                             Ltb_no(3) = LCase(tb_no(3))
  tb_no(4) = "Reference Note for Product Dimension & Weight":               Ltb_no(4) = LCase(tb_no(4))
  tb_no(5) = "Reference Note for Style / Material / Construction":              Ltb_no(5) = LCase(tb_no(5))
  tb_no(6) = "Reference Note for Product Color":                            Ltb_no(6) = LCase(tb_no(6))
  tb_no(7) = "Reference Note for Product Label / Marking":                   Ltb_no(7) = LCase(tb_no(7))
  tb_no(8) = "Reference Note for Shipping Mark":                            Ltb_no(8) = LCase(tb_no(8))
  tb_no(9) = "Reference Note for Packing / Packaging":                       Ltb_no(9) = LCase(tb_no(9))
  '--------------------------------------------------
  

  ch = "Choose"
  CF = "Conformed"
  NC = "Not Conformed"
  PD = "Pending"
  FI = "For Information"
  NA = "Not Applicable"

'Call Datacheck

 '，sap,note
  tb = WD.Tables.count - 1
  
  For i = tb - 2 To 4 Step -1

     With WD.Tables(i)
   
         'If Len(Trim(.Cell(1, 1).Range.Text)) = 2 Then
            'If.Columns.count = 2 And.Range.InlineShapes.count = 2 Then.Delete
          If LCase(.Cell(1, 1).Range.Text) Like "*" & "reference note for" & "*" = True Then

             Rn = .Rows.count
             For k = Rn To 2 Step -1
                If Len(Trim(.Cell(k, 2).Range.Text)) <= 2 Or Mid(Trim(.Cell(k, 2).Range.Text), 1, Len(Trim(.Cell(k, 2).Range.Text)) - 1) = "Choose" Then .Rows(k).Delete
             Next k
             Rn = .Rows.count
             If Rn = 1 Then
              WD.Tables(i).Select
              Selection.Tables(1).Delete
              Selection.Delete
            End If
         ElseIf LCase(.Cell(1, 1).Range.Text) Like "*" & "special attention point for" & "*" = True Then

             Rn = .Rows.count
             For k = Rn To 2 Step -1
                If Len(Trim(.Cell(k, 2).Range.Text)) <= 2 Or Mid(Trim(.Cell(k, 2).Range.Text), 1, Len(Trim(.Cell(k, 2).Range.Text)) - 1) = "Choose" Then .Rows(k).Delete
             Next k
             Rn = .Rows.count
            If Rn = 1 Then
              WD.Tables(i).Select
              Selection.Tables(1).Delete
              Selection.Delete
            End If
         ElseIf LCase(.Cell(1, 1).Range.Text) Like "*" & "special attention point summary" & "*" = True Then
              WD.Tables(i).Select
              Selection.Tables(1).Delete
              Selection.Delete
         ElseIf LCase(.Cell(1, 1).Range.Text) Like "*" & "reference note summary " & "*" = True Then
              WD.Tables(i).Select
              Selection.Tables(1).Delete
              Selection.Delete
         End If
    End With
10 Next i
  
 
   '----Note
 tb = WD.Tables.count - 1
  For i = tb - 2 To 4 Step -1
    With WD.Tables(i)
        For j = 1 To 9
          Set rg = WD.Tables(3).Cell(j + 2, 4).Range
          If LCase(.Cell(1, 1).Range.Text) Like "*" & Ltb_no(j) & "*" = True Then
             Notbn(j) = i
             Rn = .Rows.count
             Kw = ""
             N = 0
           
             TLn = 0
                For k = 2 To Rn
                  N = N + 1
                  Cw(N) = Mid(.Cell(k, 1).Range.Text, 1, Len(.Cell(k, 1).Range.Text) - 2)
                  Lw(N) = Len(.Cell(k, 1).Range.Text) - 2
                 
                  ST(N) = TLn + 1 ' translated comment
                  ED(N) = TLn + Lw(N)
                  TLn = TLn + Lw(N) + 2
                  BKText(N) = "Note" & Replace(Cw(N), ".", "")
                  Set Cel = .Cell(k, 1)
                  Call DeleteBookMark(Cel)
                 .Cell(k, 1).Range.Bookmarks.add Name:=BKText(N), Range:=.Cell(k, 1).Range
                  Kw = Kw & "," & " " & Cw(N)
                Next k
                
                 Kw = Mid(Kw, 3, Len(Kw) - 2)
                 WD.Tables(3).Cell(j + 2, 4).Range.Text = Kw
                 If Kw <> "" Then
                   For k = 1 To N
                      Set rg1 = WD.Range(Start:=rg.Characters(ST(k)).Start, End:=rg.Characters(ED(k)).End)
                      rg1.Select
                      With Selection
                        HyperlinkText = BKText(k)
                       .Hyperlinks.add Anchor:=Selection.Range, Address:="", SubAddress:=HyperlinkText
                      
                      End With
                   Next k
                End If
             GoTo 20
          End If
        Next j
    End With
20 Next i
  
  
  '----sap
 tb = WD.Tables.count - 1
  For i = tb - 2 To 4 Step -1
    With WD.Tables(i)
        For j = 1 To 9
           Set rg = WD.Tables(3).Cell(j + 2, 3).Range
          If LCase(.Cell(1, 1).Range.Text) Like "*" & Ltb_sap(j) & "*" = True Then
             Saptbn(j) = i
             Rn = .Rows.count
             SapP(j) = Rn - 1    'SAP
             Kw = ""
             N = 0
             TLn = 0
             For k = 2 To Rn
                  N = N + 1
                  Cw(N) = Mid(.Cell(k, 1).Range.Text, 1, Len(.Cell(k, 1).Range.Text) - 2)
                  Lw(N) = Len(.Cell(k, 1).Range.Text) - 2
             
                  ST(N) = TLn + 1 ' translated comment
                  ED(N) = TLn + Lw(N)
                  TLn = TLn + Lw(N) + 2
                  BKText(N) = "Sap" & Replace(Cw(N), ".", "")
                  Set Cel = .Cell(k, 1)
                  Call DeleteBookMark(Cel)
                  Cel.Range.Bookmarks.add Name:=BKText(N), Range:=Cel.Range
                  Kw = Kw & "," & " " & Cw(N)
             Next k
             
             Kw = Mid(Kw, 3, Len(Kw) - 2)
             WD.Tables(3).Cell(j + 2, 3).Range.Text = Kw
            If Kw <> "" Then
              For k = 1 To N
                   Set rg1 = WD.Range(Start:=rg.Characters(ST(k)).Start, End:=rg.Characters(ED(k)).End)
                      rg1.Select
                      With Selection
                        HyperlinkText = BKText(k)
                       .Hyperlinks.add Anchor:=Selection.Range, Address:="", SubAddress:=HyperlinkText
                      
                    End With
              Next k
            End If
             GoTo 30
         End If
        Next j
    End With
30 Next i
  
'---------
  tb = WD.Tables.count - 1

  For i = 4 To tb     ' translated comment
     With WD.Tables(i)
      
        For j = 1 To 10
          If .Cell(1, 1).Range.Text Like "*" & tb_tt(j) & "*" = True Then
             Restbn(j) = i
             If WD.Bookmarks.Exists(BK(j)) = False Then
               Set rg = WD.Tables(i).Cell(1, 1).Range
               rg.Bookmarks.add Name:=BK(j), Range:=rg
             End If
             If j < 10 Then
               Ctrn(j) = .Range.ContentControls.count ' translated comment
               If .Cell(1, 2).Range.ContentControls.count > 0 Then
                  Res(j) = .Cell(1, 2).Range.ContentControls(1).Range.Text
                  BG = 2
               Else
                 Res(j) = Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text))
                 BG = 1
               End If
                
               WD.Tables(3).Cell(j + 2, 2).Range.Text = Res(j)
               If Res(j) = CF Then
                 WD.Tables(3).Cell(j + 2, 2).Range.Font.ColorIndex = wdAuto
               ElseIf Res(j) = PD Then
                 WD.Tables(3).Cell(j + 2, 2).Range.Font.Color = 194
               Else
                 WD.Tables(3).Cell(j + 2, 2).Range.Font.Color = 194
               End If
               If Res(j) <> CF And Res(j) <> NC And Res(j) <> PD Then
                  msg = msg & "The result of " & tb_tt(j) & "has not choosen or input" & CHR(10) & _
                  tb_tt(j) & UText(30340, 26816, 39564, 32467, 26524, 36824, 26410, 36873, 25321, 25110, 36755, 20837) & CHR(10) & CHR(10)
               End If
              GoTo 40
            End If
          End If
        Next j
    End With
40  Next i
  If msg <> "" Then MsgBox msg & "Please check, revice and run program again!" & CHR(10) & _
  UText(35831, 26816, 26597, 26356, 27491, 24182, 37325, 26032, 36816, 34892, 31243, 24207) & "！" & CHR(10), vbCritical: Exit Sub
  
  
   For i = 3 To 9  ' translated comment
      
        N = 0
        SCF(i) = 0: SNC(i) = 0: SPD(i) = 0: NAP(i) = 0
      With WD.Tables(Restbn(i))
         Rn = .Rows.count
         Cn = .Columns.count
        
         For j = 4 To Rn
             Set rg = .Cell(j, Cn).Range
             If rg.ContentControls.count = 1 Then
               Kw = rg.ContentControls(1).Range.Text
             Else
               Kw = Mid(rg.Text, 1, Len(rg.Text) - 2)
             End If
            
               If Kw = CF Then             ' translated comment
                 SCF(i) = SCF(i) + 1
               ElseIf Kw = NC Then
                  SNC(i) = SNC(i) + 1
               ElseIf Kw = PD Then
                 SPD(i) = SPD(i) + 1
               ElseIf Kw = NA Then
                 NAP(i) = NAP(i) + 1
               ElseIf Kw = FI Then
                 INF(i) = INF(i) + 1
                  
               ElseIf Kw = ch Or Trim(Kw) = "" Then
                 N = N + 1
              End If
           
        Next j

         If N > 0 Then msg = msg & N & " options of sub result items in " & tb_tt(i) & " section has not been choosen or input" & CHR(10) & _
         tb_tt(i) & UText(21306, 22495, 26377) & N & UText(20010, 20998, 39033, 30340, 32467, 26524, 36873, 39033, 26410, 36873, 25321, 25110, 36755, 20837) & CHR(10)
         
         N = 0
      

         For j = Restbn(i) To Restbn(i + 1) - 1 ' translated comment
            With WD.Tables(j)
              Rn = .Rows.count
              Cn = .Columns.count
              Celn = .Range.Cells.count
  
              For k = 1 To Celn
                 If Len(Trim(.Range.Cells(k).Range.Text)) >= 8 Then
                 If Mid(Trim(.Range.Cells(k).Range.Text), 1, Len(Trim(.Range.Cells(k).Range.Text)) - 2) = "Result" Then
                  RI = .Range.Cells(k).rowIndex
                    For L = RI + 1 To Rn
                      Kw = Trim(.Cell(L, Cn).Range.Text)
                      Kw = Trim(Mid(Kw, 1, Len(Kw) - 2))
                       If Kw = "" Then             ' translated comment
                          N = N + 1
                       End If
                    Next L
                    GoTo 38
                End If
               End If
              Next k
            End With
38     Next j
         If N > 0 Then msg = msg & "For tables of " & tb_tt(i) & " section , there are total " & N & " rows without input result data " & CHR(10) & _
         tb_tt(i) & UText(21306, 22495, 30340, 34920, 26684, 20013, 26377) & N & UText(34892, 32467, 26524, 39033, 26410, 36755, 20837, 25968, 25454) & CHR(10)
     End With
   Next i
     
     If msg <> "" Then MsgBox CHR(10) & msg & "Please check,revise and run program again!" & CHR(10) & _
   UText(35831, 26816, 26597, 20462, 25913, 24182, 37325, 26032, 36816, 34892, 31243, 24207), vbCritical: Exit Sub
   
 '---
   
  
    
    With WD.Tables(3)
        CFN = 0: NCN = 0: PDN = 0
        Rn = .Rows.count
        For i = 1 To 9
           If Res(i) = CF Then CFN = CFN + 1
           If Res(i) = NC Then NCN = NCN + 1
           If Res(i) = PD Then PDN = PDN + 1
          
        Next i
          If NCN > 0 Then
            .Cell(Rn, 2).Range.Text = "NOT CONFORM to client's requirement"
              With .Cell(Rn, 2).Range.words(1).Font
                 .Color = 192
                 .Size = 12
                 .Bold = True
              End With
              With .Cell(Rn, 2).Range.words(2).Font
                 .Color = 192
                 .Size = 12
                 .Bold = True
              End With
               With .Cell(Rn, 2).Range.words(3).Font
                 .Color = wdColorAutomatic
                 .Size = 10
                 .Bold = True
              End With
              With .Cell(Rn, 2).Range.words(4).Font
                 .Color = wdColorAutomatic
                 .Size = 10
                 .Bold = True
              End With
              With .Cell(Rn, 2).Range.words(5).Font
                 .Color = wdColorAutomatic
                 .Size = 10
                 .Bold = True
              End With
             
          ElseIf PDN > 0 Then
             .Cell(Rn, 2).Range.Text = "PENDING for client's evaluation"
               With .Cell(Rn, 2).Range.words(1).Font
                 .Color = 192
                 .Size = 12
                 .Bold = True
              End With
              
               With .Cell(Rn, 2).Range.words(2).Font
                 .Color = wdColorAutomatic
                 .Size = 10
                 .Bold = True
              End With
              With .Cell(Rn, 2).Range.words(3).Font
                 .Color = wdColorAutomatic
                 .Size = 10
                 .Bold = True
              End With
              With .Cell(Rn, 2).Range.words(4).Font
                 .Color = wdColorAutomatic
                 .Size = 10
                 .Bold = True
              End With
          Else
             .Cell(Rn, 2).Range.Text = "CONFORM to client's requirement"
              With .Cell(Rn, 2).Range.words(1).Font
                 .Color = wdColorAutomatic
                 .Size = 12
                 .Bold = True
              End With
              
               With .Cell(Rn, 2).Range.words(2).Font
                 .Color = wdColorAutomatic
                 .Size = 10
                 .Bold = True
              End With
              With .Cell(Rn, 2).Range.words(3).Font
                 .Color = wdColorAutomatic
                 .Size = 10
                 .Bold = True
              End With
              With .Cell(Rn, 2).Range.words(4).Font
                 .Color = wdColorAutomatic
                 .Size = 10
                 .Bold = True
              End With
          End If
    End With
  

'--------------
50 For i = 1 To 9
   Select Case i
     
     Case 1
       If SapP(i) > 0 Then
          If Res(i) = CF Then msg = msg & "For " & tb_tt(i) & ", There were " & SapP(i) & _
           " Special Attention Points, however the inspection reuslt was " & _
            Res(i) & ". It shall be " & NC & " Or " & PD & CHR(10) & tb_tt(i) & "," & UText(26377) & SapP(i) & _
           UText(39033) & "Special Attention Points," & UText(20294, 32467, 35770, 21364, 26159) & Res(i) & "," & UText(24212, 20026) & NC & UText(25110) & PD & CHR(10)
          
       Else
         If Res(i) <> CF Then msg = msg & "For " & tb_tt(i) & ", not any Special Attention Points was input" & _
         ", however the inspection reuslt was " & Res(i) & ", It shall be " & CF & CHR(10) & _
          tb_tt(i) & "," & UText(27809, 26377) & "Special Attention Points," & UText(20294, 32467, 35770, 21364, 26159) & Res(i) & "," & UText(24212, 20026) & CF & CHR(10)
       End If
    Case 2
       With WD.Tables(Restbn(i) + 1)
         Rn = .Rows.count
         For j = 2 To Rn
           If .Rows(j).Range.Cells.count <> 1 Then
              If .Cell(j, 1).Range.Text Like "*Allowed*" = True Then
                Kcr = Trim(Mid(.Cell(j, 2).Range.Text, 1, Len(.Cell(j, 2).Range.Text) - 2))
                Kma = Trim(Mid(.Cell(j, 3).Range.Text, 1, Len(.Cell(j, 3).Range.Text) - 2))
                Kmi = Trim(Mid(.Cell(j, 4).Range.Text, 1, Len(.Cell(j, 4).Range.Text) - 2))
                 If Not IsNumeric(Kcr) Or Not IsNumeric(Kma) Or Not IsNumeric(Kmi) Then               'AQL
                   If Res(i) <> PD Then
                      msg = msg + "For " & tb_tt(i) & ", not Allowed data was input in defects table, " & _
                      " however, the inspection reuslt was " & Res(i) & ".It shall be " & PD & CHR(10) & _
                      tb_tt(i) & ",Defect" & UText(34920, 20013, 26410, 36755, 20837) & "Allowed" & UText(25968, 20540) & "," & UText(20294, 32467, 35770, 21364, 26159) & Res(i) & "," & UText(24212, 20026) & PD & CHR(10)
                      GoTo 55
                    End If
                 Else
                     If Val(.Cell(j - 1, 2).Range.Text) > Val(Kcr) Or Val(.Cell(j - 1, 3).Range.Text) > Val(Kma) Or _
                        Val(.Cell(j - 1, 4).Range.Text) > Val(Kmi) Then     'AQL
                         If Res(i) <> NC Then
                            msg = msg + "For " & tb_tt(i) & ", defects found exceed allowance, " & _
                           " however, the inspection reuslt was " & Res(i) & ".It shall be " & NC & CHR(10) & _
                            tb_tt(i) & "," & UText(21457, 29616, 32570, 38519, 25968, 36229, 36807, 20801, 35768, 20540) & "," & UText(20294, 32467, 35770, 21364, 26159) & Res(i) & "," & UText(24212, 20026) & NC & CHR(10)
                            GoTo 55
                          End If
                     Else
                          If SapP(i) = 0 And Res(i) <> CF Then
                            msg = msg + "For " & tb_tt(i) & ", defects not exceed allowance and not special attention point was record, " & _
                           " however, the inspection reuslt was " & Res(i) & ". It shall be " & CF & CHR(10) & _
                            tb_tt(i) & "," & UText(32570, 38519, 25968, 26410, 36229, 36807, 20801, 35768, 20540, 19988, 27809, 26377) & "special attention point," & UText(20294, 32467, 35770, 21364, 26159) & Res(i) & "," & UText(24212, 20026) & CF & CHR(10)
                            GoTo 55
                          End If
                     End If
                 End If
              End If
           End If
         Next j
         If SapP(i) > 0 And Res(i) <> NC Then msg = msg & "For " & tb_tt(i) & ", There were " & SapP(i) & _
           " Special Attention Points, however, the inspection reuslt was " & _
            Res(i) & ". It shall be " & NC & CHR(10) & tb_tt(i) & "," & UText(26377) & SapP(i) & _
           UText(39033) & "Special Attention Points," & UText(20294, 32467, 35770, 21364, 26159) & Res(i) & "," & UText(24212, 20026) & NC & CHR(10)
         
55     End With
      
      Case 3, 4, 5, 6, 7, 8, 9
      
       If SNC(i) > 0 Then
           If Res(i) <> NC Then msg = msg & "For " & tb_tt(i) & ", There were " & SNC(i) & " sub items were " & NC & _
           ", however, the inspection reuslt was " & Res(i) & ". It shall be " & NC & CHR(10) & _
           tb_tt(i) & "," & UText(26377) & SNC(i) & UText(23376, 39033, 32467, 26524, 26159) & NC & "," & UText(20294, 32467, 35770, 21364, 26159) & Res(i) & "," & UText(24212, 20026) & NC & CHR(10)
          
       Else
           If SPD(i) > 0 Then
             If Res(i) <> PD Then msg = msg & "For " & tb_tt(i) & ", There were " & SPD(i) & " sub items  were " & PD & _
             ", however, the inspection reuslt was " & Res(i) & ". It shall be " & PD & CHR(10) & _
                tb_tt(i) & "," & UText(26377) & SPD(i) & UText(23376, 39033, 32467, 26524, 26159) & PD & "," & UText(20294, 32467, 35770, 21364, 26159) & Res(i) & "," & UText(24212, 20026) & PD & CHR(10)
          Else
                If SapP(i) > 0 Then
                  If Res(i) <> PD Then msg = msg & "For " & tb_tt(i) & ", There were " & SapP(i) & _
                    " Special Attention Points , however, the inspection reuslt was " & Res(i) & ". It shall be " & PD & CHR(10) & _
                     tb_tt(i) & "," & UText(26377) & SapP(i) & "Special Attention Points" & UText(39033) & "," & UText(20294, 32467, 35770, 21364, 26159) & Res(i) & "," & UText(24212, 20026) & PD & CHR(10)
                 Else
                    If Res(i) <> CF Then msg = msg & "For " & tb_tt(i) & ", All sub items were " & CF & _
                    " and with out Special Attention Points , however, the inspection reuslt was " & _
                     Res(i) & ". It shall be " & CF & CHR(10) & _
                     tb_tt(i) & "," & UText(25152, 26377, 23376, 39033, 37117, 26159) & CF & UText(19988, 27809, 26377) & "Special Attention Points," & UText(20294, 32467, 35770, 21364, 26159) & Res(i) & "," & UText(24212, 20026) & CF & CHR(10)
                
                End If
        End If
          
       End If
  
   End Select

  Next i

    
   '==================SAP,

 
  ' CustomerID = GetPropertyValue("CustomerID")
    
  'If CustomerID <> "" Then
    'Select Case CustomerID
       'Case "32927", "49858" 'Jokey Sohland GmbH,Topitop SAP
       Dim sn As Integer
   
      '-----------------------------
   Dim Note_No(100), Note_Txt(100) As String

       tb = WD.Tables.count - 1
       sn = 0
       For j = 3 To tb
          Set tbl = WD.Tables(j)
            If LCase(tbl.Cell(1, 1).Range.Text) Like "*" & "reference note for" & "*" = True Then
               Rn = tbl.Rows.count
               For i = 2 To Rn
                    sn = sn + 1
                     Note_No(sn) = Trim(Mid(tbl.Cell(i, 1).Range.Text, 1, Len(tbl.Cell(i, 1).Range.Text) - 2))
                     
                    If tbl.Cell(i, 2).Range.ContentControls.count > 0 Then
                     
                         Note_Txt(sn) = Trim(Mid(tbl.Cell(i, 2).Range.Text, 1, Len(tbl.Cell(i, 2).Range.Text) - 1))
                      Else
                         Note_Txt(sn) = Trim(Mid(tbl.Cell(i, 2).Range.Text, 1, Len(tbl.Cell(i, 2).Range.Text) - 2))
                      End If
                  Next i
             End If
        Next j

        If sn > 0 Then
          Set tbl = WD.Tables(3)
          Rn = tbl.Rows.count

          WD.Tables(3).Cell(Rn, 1).Select
          Selection.MoveDown
          Selection.TypeParagraph
          WD.Tables.add Range:=Selection.Range, NumRows:=sn + 1, NumColumns:= _
             2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
             wdAutoFitFixed
           With Selection.Tables(1)
               If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
               End If
               Call Setborders025
              .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
               With .Range.Font
                 .Size = 10
                 .Name = "Arial"
               End With
              .Rows(1).Height = 16
              .Columns(1).Width = 28.2 'CentimetersToPoints(W1)
              .Columns(2).Width = 504.65 - 28.2 'CentimetersToPoints(W2)
              .Rows(1).Range.Cells.Merge
              .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Rows(1).Range.Cells.Shading.BackgroundPatternColor = 12040421 '12040421 ，-603914241 ' -603917569 '
              '.Rows(2).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603914241 ' -603917569 '
              .Cell(1, 1).Range.Text = "Reference Note Summary "
                With .Cell(1, 1).Range.Font
                 .Bold = True
                End With
            End With
            With WD.Tables(4)
              For i = 1 To sn
                 .Cell(i + 1, 1).Range.Text = Note_No(i)
                 .Cell(i + 1, 2).Range.Text = Note_Txt(i)
                    BKText(Int(Val(Note_No(i)))) = "Note" & Replace(Note_No(i), ".", "")
                    HyperlinkText = BKText(Int(Val(Note_No(i))))
            
                    Set rg = .Cell(i + 1, 1).Range
                   rg.Select
                    With Selection
                      .Hyperlinks.add Anchor:=Selection.Range, Address:="", SubAddress:=HyperlinkText
                    End With
                   ' rg.Font.Underline = wdUnderlineNone
                 Next i
             End With
         End If
    '-------------------------------------
    Dim SAP_No(100), SAP_Txt(100) As String

       tb = WD.Tables.count - 1
       sn = 0
       For j = 3 To tb
          Set tbl = WD.Tables(j)
            If LCase(tbl.Cell(1, 1).Range.Text) Like "*" & "special attention point for" & "*" = True Then
               Rn = tbl.Rows.count
               For i = 2 To Rn
                    sn = sn + 1
                     SAP_No(sn) = Trim(Mid(tbl.Cell(i, 1).Range.Text, 1, Len(tbl.Cell(i, 1).Range.Text) - 2))
                     
                    If tbl.Cell(i, 2).Range.ContentControls.count > 0 Then
                     
                         SAP_Txt(sn) = Trim(Mid(tbl.Cell(i, 2).Range.Text, 1, Len(tbl.Cell(i, 2).Range.Text) - 1))
                      Else
                         SAP_Txt(sn) = Trim(Mid(tbl.Cell(i, 2).Range.Text, 1, Len(tbl.Cell(i, 2).Range.Text) - 2))
                      End If
                  Next i
             End If
        Next j
        
        If sn > 0 Then

         Set tbl = WD.Tables(3)
         Rn = tbl.Rows.count

          WD.Tables(3).Cell(Rn, 1).Select
          Selection.MoveDown
          Selection.TypeParagraph
          WD.Tables.add Range:=Selection.Range, NumRows:=sn + 1, NumColumns:= _
             2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
             wdAutoFitFixed
           With Selection.Tables(1)
               If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
               End If
               Call Setborders025
              .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
               With .Range.Font
                 .Size = 10
                 .Name = "Arial"
               End With
              .Rows(1).Height = 16
              .Columns(1).Width = 28.2 'CentimetersToPoints(W1)
              .Columns(2).Width = 504.65 - 28.2 'CentimetersToPoints(W2)
              .Rows(1).Range.Cells.Merge
              .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Rows(1).Range.Cells.Shading.BackgroundPatternColor = 12040421 '12040421 ，-603914241 ' -603917569 '
              '.Rows(2).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603914241 ' -603917569 '
              .Cell(1, 1).Range.Text = "Special Attention Point Summary "
                With .Cell(1, 1).Range.Font
                 .Bold = True
                End With
            End With
            With WD.Tables(4)
              For i = 1 To sn
                 .Cell(i + 1, 1).Range.Text = SAP_No(i)
                 .Cell(i + 1, 2).Range.Text = SAP_Txt(i)
                    BKText(Int(Val(SAP_No(i)))) = "Sap" & Replace(SAP_No(i), ".", "")
                    HyperlinkText = BKText(Int(Val(SAP_No(i))))
            
                    Set rg = .Cell(i + 1, 1).Range
                   rg.Select
                    With Selection
                      .Hyperlinks.add Anchor:=Selection.Range, Address:="", SubAddress:=HyperlinkText
                    End With
                   ' rg.Font.Underline = wdUnderlineNone
                 Next i
             End With
         End If
         'End Select
    'End If
  
  
 
  
  
  
100  '========================
      If msg <> "" Then
       MsgBox msg & "Please check,if you are sure all are correct, negled this alart, otherwise, revise these and run program again!" & CHR(10) & _
      UText(35831, 26816, 26597) & "," & UText(22914, 26524, 20320, 30830, 23450, 19968, 20999, 27491, 30830, 35831, 24573, 30053, 26412, 25552, 31034) & "，" & UText(21542, 21017, 35831, 20462, 25913, 24182, 37325, 26032, 36816, 34892, 31243, 24207), vbCritical
      Else
       MsgBox "Result check completed!" & CHR(10) & "If you find that all tables are in inappropriate positions due to the version of Office," & _
       "please click CTRL-0 to make adjustment." & CHR(10) & UText(32467, 35770, 26816, 26597, 23436, 25104) & "！" & CHR(10) & _
       UText(22914, 26524, 20320, 21457, 29616, 22240, 20026) & "office" & UText(29256, 26412, 30340, 21407, 22240, 25152, 26377, 34920, 26684, 22788, 20110, 19981, 24688, 24403, 30340, 20301, 32622, 35831, 25353) & "CTRL-0" & UText(36827, 34892, 35843, 25972) & "！"
      End If

'------------------

End Sub


Sub Input_CheckPoint(ByRef N As Integer, ByRef p() As Variant)

'by wyn 20241129

Dim Ino(), Num(), Typ(), Chk(), Cri(), Spe(), sms(), Tol(), Nat() As String

ReDim Ino(1 To N)
ReDim Num(1 To N)
ReDim Typ(1 To N)
ReDim Chk(1 To N)
ReDim Cri(1 To N)
ReDim Spe(1 To N)
ReDim sms(1 To N)
ReDim Tol(1 To N)
ReDim Nat(1 To N)

Dim Cha(200) As String

Dim NumO(), NumP(), NumW(), NumD(), NumL(), NumS(), NumC(), NumM() As String
ReDim NumO(1 To N)
ReDim NumP(1 To N)
ReDim NumW(1 To N)
ReDim NumD(1 To N)
ReDim NumL(1 To N)
ReDim NumS(1 To N)
ReDim NumC(1 To N)
ReDim NumM(1 To N)

Dim ChkO(), ChkP(), ChkW(), ChkD(), ChkL(), ChkS(), ChkC(), ChkM() As String
ReDim ChkO(1 To N)
ReDim ChkP(1 To N)
ReDim ChkW(1 To N)
ReDim ChkD(1 To N)
ReDim ChkL(1 To N)
ReDim ChkS(1 To N)
ReDim ChkC(1 To N)
ReDim ChkM(1 To N)

 
Dim i, j, k, X, tb, tbn, TbnB, TbnE, cnt, tNum, Z As Integer
Dim W, o, PP, L, D, S, C, M As Integer
Dim cno, cnp, cnw, cnd, cnl, cns, cnc, cnm, Ps, Psn, NN, NN1 As Integer
Dim msg, DelTab, Kw1, Kw2 As String

Dim a, TCHAR
Dim uC(200) As String
Dim CHAR(100, 50) As String
Dim KwChk As String
Dim Sm() As Integer

Dim TableChoise As String
Dim TbCount1, TbCount2, TbCount3 As Integer

For i = 1 To N_Cat_Middle
  If Cat_Middle(i) = UText(38795, 31867) Then
    TableChoise = 1
    GoTo 3
  End If
Next i

TableChoise = InputBox("The program will generate tabels for datasheet, please input number for the table type you would like" & CHR(10) & _
"1:  normal datasheet, one check point with one table, one row for one item of product." & CHR(10) & _
"2:  complex datasheet, some cell units for one item of product,each row for one sub check point, and support pasting in drawings photo." & CHR(10) & _
"3:  Both 1 and 2 mentioned above" & CHR(10) & CHR(10) & _
UText(31243, 24207, 23558, 29983, 25104, 25968, 25454, 34920, 26684) & "，" & UText(35831, 36755, 20837, 25968, 23383, 36873, 25321, 29983, 25104, 30340, 25968, 25454, 34920, 26684, 31867, 22411) & CHR(10) & _
"1: " & UText(36890, 29992, 25968, 25454, 34920) & "，" & UText(27599, 20010, 26816, 26597, 39033, 30446, 19968, 20010, 34920, 26684) & "，" & UText(27599, 27454, 20135, 21697, 25968, 25454, 21344, 29992, 19968, 34892) & CHR(10) & _
"2: " & UText(22797, 21512, 25968, 25454, 34920, 26684) & "，" & UText(27599, 27454, 20135, 21697, 21344, 29992, 22810, 20010, 21333, 20803, 26684, 32452, 21512) & "，" & UText(27599, 34892, 21487, 29992, 20110, 36755, 20837, 23376, 26816, 26597, 39033, 30446) & "，" & UText(25903, 25345, 31896, 36148, 20837, 27979, 37327, 22270, 32440, 29031, 29255) & CHR(10) & _
"3: " & UText(21516, 26102, 29983, 25104, 19978, 36848) & "1，2" & UText(34920, 26684) & CHR(10), "Please Choose datasheet type" & "  " & UText(35831, 36873, 25321, 34920, 26684, 31867, 22411), "1")
   
3 StoretoProperty "TableChoise", TableChoise

    W1 = 3.8    ' translated comment
    W2 = 4
    W3 = 2.2
    W4 = 7.8

5 Set WD = ThisDocument

   Dim qtyRow As Integer
    With ThisDocument.Range.Tables(getTableNumber(1))
     qtyRow = .Rows.count
     Itms = qtyRow - 5
    ReDim Ptx(1 To Itms)
    ReDim Itx(1 To Itms)
    ReDim Sm(1 To Itms)
     For i = 5 To qtyRow - 1
       If Mid(.Cell(i, 1).Range.Text, 1, Len(.Cell(i, 1).Range.Text) - 2) <> "" Or Mid(.Cell(i, 2).Range.Text, 1, Len(.Cell(i, 2).Range.Text) - 2) <> "" Then
          Ptx(i - 4) = Mid(.Cell(i, 1).Range.Text, 1, Len(.Cell(i, 1).Range.Text) - 2)
          Itx(i - 4) = Mid(.Cell(i, 2).Range.Text, 1, Len(.Cell(i, 2).Range.Text) - 2)
          Sm(i - 4) = Val(.Cell(i, 8).Range.Text)     ' translated comment
      End If
     Next i
   End With
   
 Dim ComItx() As String
 ReDim ComItx(1 To Itms)
 
 Dim ComSm() As Integer
 ReDim ComSm(1 To Itms)

 Dim ComItms As Integer
 Dim NI As Integer
   
      NI = 0
      Set dictItem = CreateObject("Scripting.Dictionary")
      For i = 1 To Itms
        If Not dictItem.Exists(Itx(i)) Then
            NI = NI + 1
            dictItem.add Itx(i), Sm(i)
             ComItx(NI) = Itx(i)
             ComSm(NI) = Sm(i)
        Else
            dictItem(Itx(i)) = dictItem(Itx(i)) + Sm(i)
            ComSm(NI) = dictItem(Itx(i))
        End If
    Next i
    ComItms = dictItem.count
    

     'Itms = Val(GetPropertyValue("ItemsQty"))
    
     StoretoProperty "ItemsQty", CStr(Itms)
   '-----------
 
 '（0）-（1）-（2)-（3）-（4）-（5）-（6）-（7）-（8）

 For i = 1 To N
    p(i) = Replace(p(i), "，", ",")
    a = Split(p(i), "|")
    u = UBound(a)

    Num(i) = i
    Ino(i) = a(0)
    Chk(i) = a(1)
    Typ(i) = Trim(a(2))
    Cri(i) = Trim(a(3))
    sms(i) = Trim(a(4))

    NN1 = Val(sms(i))
    Psn = InStr(1, sms(i), "≥")
    If Psn > 0 Then
     NN = Val(Mid(sms(i), Psn + 1, Len(sms(i)) - Psn))
      If ComItms = 1 Then
        sms(i) = CStr(NN)
      Else
         sms(i) = Mid(sms(i), 1, Psn - 1)
         Do While NN1 * ComItms < NN
            sms(i) = CStr((NN1 + 1)) & Mid(sms(i), 2, Len(sms(i)) - 1)
            NN1 = Val(sms(i))
         Loop
      End If
    Else
       If ComItms = 1 Then sms(i) = Replace(sms(i), "/item", "")
    End If
    
    Spe(i) = Trim(a(5))
    Tol(i) = Trim(a(6))
    Nat(i) = Trim(a(7))
    
    Cha(i) = Trim(a(8))      'char（）
       TCHAR = Split(Cha(i), ",")
       uC(i) = UBound(TCHAR)
       For j = 0 To uC(i)
         CHAR(i, j) = TCHAR(j)
       Next j
   Next i
   
   
  
    W = 0:      o = 0:     D = 0:     S = 0:     C = 0:     L = 0:     M = 0:     PP = 0
 
 
   For i = 1 To N
      
      Select Case LCase(Typ(i))
       '-------------------
      Case UText(25163, 24037, 26816, 26597)
      'Case "workmanship"
                    
                            W = W + 1
                            NumW(W) = Num(i)
                            ChkW(W) = Chk(i)
                            GoTo 10
                
       '-------------------
      Case UText(29305, 27530, 27979, 35797)
     ' Case "on site test"
         
                            o = o + 1
                            NumO(o) = Num(i)
                            ChkO(o) = Chk(i)
                            GoTo 10
                 
    '------------------
   Case UText(20135, 21697, 27979, 37327), UText(37325, 37327, 26816, 26597)
   ' Case "product dimension and weight"
        
                            D = D + 1
                           NumD(D) = Num(i)
                           ChkD(D) = Chk(i)
                            GoTo 10
    '===================
    Case UText(27454, 24335, 26816, 26597), UText(25991, 20214, 26816, 26597)
                           S = S + 1
                           NumS(S) = Num(i)
                           ChkS(S) = Chk(i)
                            GoTo 10
    
    '----------
    Case UText(39068, 33394, 26816, 26597)
    
                           C = C + 1
                           NumC(C) = Num(i)
                           ChkC(C) = Chk(i)
                           GoTo 10
     
    '===
     Case UText(26631, 31614, 26631, 35782, 26816, 26597)
       ' Case "labeling check"
                             L = L + 1
                             NumL(L) = Num(i)
                            ChkL(L) = Chk(i)
                            GoTo 10
                            
       '==================
     Case UText(21787, 22836, 26816, 26597)
                            M = M + 1
                           NumM(M) = Num(i)
                           ChkM(M) = Chk(i)
                           GoTo 10
       
      '--------------------------------------
       Case UText(21253, 35013, 26816, 26597)
      'Case "packing and packaging"
                
                             PP = PP + 1
                             NumP(PP) = Num(i)
                              ChkP(PP) = Chk(i)
                              GoTo 10
  
    End Select
10    Next i

     cnw = W:        cno = o:       cnd = D:       cns = S:       cnc = C:       cnl = L:       cnm = M:       cnp = PP

  
20    '------------------------------------------workmanship
  tbn = 0
  tbn = getTableNumber(2)
     If tbn = 0 Then
        msg = msg & UText(26410, 21457, 29616) & "2. Workmanship" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
        "2. Workmanship table is not found, the check item for workmanship is not imported" & CHR(10)
        GoTo 40
    End If
30  With WD.Tables(tbn)
       Rn = .Rows.count
        If cnw > 0 Then
          If Rn = 5 Then
             .Cell(Rn, 1).Select
              Selection.InsertRowsBelow 1
             .Rows(Rn + 1).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 '
             .Rows(Rn + 1).Range.Cells.Merge
             .Rows(Rn + 1).Range.Cells.Split NumRows:=1, NumColumns:=3, mergebeforesplit:=True
             .Cell(Rn + 1, 1).SetWidth ColumnWidth:=36, RulerStyle:=wdAdjustFirstColumn
             .Cell(Rn + 1, 2).SetWidth ColumnWidth:=110, RulerStyle:=wdAdjustFirstColumn
              '.Cell(Rn + 1, 3).SetWidth ColumnWidth:=260, RulerStyle:=wdAdjustFirstColumn
             .Rows(Rn + 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
             .Cell(Rn + 1, 1).Range.Text = "No."
             .Cell(Rn + 1, 2).Range.Text = "Check Point"
             .Cell(Rn + 1, 3).Range.Text = "Criteria"
              '.Cell(Rn + 1, 4).Range.Text = "Sample Size"
              
             .Cell(Rn + 1, 1).Select
              Selection.InsertRowsBelow 1
            
             .Rows(Rn + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  ' translated comment
             .Rows(Rn + 2).Range.Cells(1).Shading.BackgroundPatternColor = -603914241  '-603917569 '
             .Cell(Rn + 2, 1).Select
              Selection.InsertRowsBelow cnw - 1
          ElseIf Rn = 6 Then
              .Rows(Rn).Range.Cells.Shading.BackgroundPatternColor = -603917569  '-603914241 '
              .Rows(Rn).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
              .Cell(Rn, 1).Select
               Selection.InsertRowsBelow 1
             .Rows(Rn + 1).Range.Cells.Shading.BackgroundPatternColor = -603914241  ' translated comment
             .Rows(Rn + 1).Range.Cells(1).Shading.BackgroundPatternColor = -603914241  '-603917569 '
             .Cell(Rn + 1, 1).Select
               Selection.InsertRowsBelow cnw - 1
          ElseIf Rn < cnw + 6 Then
           
             .Cell(Rn, 1).Select
              Selection.InsertRowsBelow cnw + 6 - Rn
          ElseIf Rn > cnw + 6 Then
                 For j = Rn To cnw + 6 + 1 Step -1
                 .Rows(j).Delete
                 Next j
           End If
        Else
           If Rn > 5 Then
              For j = Rn To 6 Step -1
                 .Rows(j).Delete
              Next j
           End If

        End If
       
     
       
       For i = 1 To cnw
         .Rows(i + 6).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Cell(i + 6, 1).Range.Text = "2." & i
          If Len(Chk(NumW(i))) > 1 Then
            KwChk = UCase(Mid(Chk(NumW(i)), 1, 1)) & Mid(Chk(NumW(i)), 2, Len(Chk(NumW(i))) - 1)
          Else
            KwChk = Chk(NumW(i))
          End If
          If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
          
         .Cell(i + 6, 2).Range.Text = KwChk
         Set Cel = .Cell(i + 6, 2)
         Call DeleteBookMark(Cel)
         Cel.Range.Bookmarks.add Name:=Ino(NumW(i)), Range:=Cel.Range
        
          If Len(Cri(NumW(i))) > 1 Then
                 .Cell(i + 6, 3).Range.Text = UCase(Mid(Cri(NumW(i)), 1, 1)) & Mid(Cri(NumW(i)), 2, Len(Cri(NumW(i))) - 1)
           Else
              .Cell(i + 6, 3).Range.Text = UCase(Mid(Cri(NumW(i)), 1, 1))
          End If
       Next i
     End With
  
     '---workmanship
    TbnB = tbn
    DelTab = ""
    
  
   For i = 1 To cnw
     If Nat(NumW(i)) = UText(23450, 37327) Then
          DelTab = DelTab & "," & "Data Sheet for " & Chk(NumW(i))  ' translated comment
          TbCount1 = 0
          tb = WD.Tables.count - 1
          For j = TbnB To tb
            With WD.Tables(j)
               If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumW(i)) & "*" = True Then
                   TbCount1 = TbCount1 + 1
                   If TbCount1 > 1 Then
                     .Delete
                     j = j - 1
                   Else
                     Set Cel = .Cell(1, 1)
                     Call DeleteBookMark(Cel)
                     Cel.Range.Bookmarks.add Name:=Ino(NumW(i)) & "_Data", Range:=Cel.Range
                      Rn = .Rows.count
                      If Rn < ComItms + 2 Then
                          .Cell(Rn, 1).Select
                          Selection.InsertRowsBelow ComItms + 2 - Rn
                      ElseIf Rn > ComItms + 2 Then
                         For k = Rn To ComItms + 2 + 1 Step -1
                           .Rows(k).Delete
                         Next k
                      End If
                      For k = 1 To ComItms
                        .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603917569 '
                        .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                        .Cell(k + 2, 1).Range.Text = ComItx(k)
                        .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                        .Cell(k + 2, 2).Range.Text = Spe(NumW(i))
                        .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                        .Cell(k + 2, 3).Range.Text = Tol(NumW(i))
                        .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                      Next k
                    End If
                 End If
            End With
        Next j
        
35    tbn = 0
      tbn = getTableNumber(3)
      If tbn = 0 Then
          msg = msg & UText(26410, 21457, 29616) & "3. On Site Test" & UText(30340, 34920, 26684) & "，Workmanship " & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
         "3. On Site Test table is not found, Data sheet for Workmanship form can not generate" & CHR(10)
          GoTo 40
      End If
37    If TbCount1 = 0 Then
              Call MakeNormalDataSheet(tbn, ComItms)
              
              With WD.Tables(tbn)
                Set Cel = .Cell(1, 1)
               .Cell(1, 1).Range.Text = "Data Sheet for " & Chk(NumW(i))
                Call DeleteBookMark(Cel)
                Cel.Range.Bookmarks.add Name:=Ino(NumW(i)) & "_Data", Range:=Cel.Range
                For j = 1 To ComItms
                 .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                 .Cell(j + 2, 1).Range.Text = ComItx(j)
                 .Cell(j + 2, 2).Range.Text = Spe(NumW(i))
                 .Cell(j + 2, 3).Range.Text = Tol(NumW(i))
               Next j
            End With
       End If
     End If
38 Next i
  
 
40  ''------------------------------------------3.onsitetest
     tbn = 0
     tbn = getTableNumber(3)
      If tbn = 0 Then
         msg = msg & UText(26410, 21457, 29616) & "3. On Site Test" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
        "3. On Site Test table is not found, the check item for On Site Test is not imported" & CHR(10)
         GoTo 60
     End If
50  With WD.Tables(tbn)
        Rn = .Rows.count
         If Rn = 3 Then
          .Cell(Rn, 1).Select
            Selection.InsertRowsBelow 1
          .Rows(Rn + 1).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603917569 '
           Rn = .Rows.count
        End If
       If Rn < cno + 3 Then
         .Rows(4).Select
          Selection.Copy
          For j = 1 To cno + 3 - Rn
           .Rows(Rn).Select
             Dim StartTime As Single
            StartTime = Timer '' 100
            Do While Timer - StartTime < 0.1
               DoEvents
            Loop
           Selection.Paste
           Next j
        ElseIf Rn > cno + 3 Then
          For j = Rn To cno + 3 + 1 Step -1
          .Rows(j).Delete
          Next j
       End If
   
       For i = 1 To cno
        .Rows(i + 3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Cell(i + 3, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
         
        .Cell(i + 3, 1).Range.Text = "3." & i
          If Len(Chk(NumO(i))) > 1 Then
           KwChk = UCase(Mid(Chk(NumO(i)), 1, 1)) & Mid(Chk(NumO(i)), 2, Len(Chk(NumO(i))) - 1)
          Else
             KwChk = Chk(NumO(i))
          End If
          If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
          
       
         .Cell(i + 3, 2).Range.Text = KwChk
          Set Cel = .Cell(i + 3, 2)
          Call DeleteBookMark(Cel)
         Cel.Range.Bookmarks.add Name:=Ino(NumO(i)), Range:=Cel.Range
         
         If Len(Cri(NumO(i))) > 1 Then
         .Cell(i + 3, 3).Range.Text = UCase(Mid(Cri(NumO(i)), 1, 1)) & Mid(Cri(NumO(i)), 2, Len(Cri(NumO(i))) - 1)
         Else
          .Cell(i + 3, 3).Range.Text = Cri(NumO(i))
         End If
         If sms(NumO(i)) > "" Then
             If ComItms = 1 Then
                sms(NumO(i)) = Replace(sms(NumO(i)), " /item", "")
                sms(NumO(i)) = Replace(sms(NumO(i)), " /Item", "")
                .Cell(i + 3, 4).Range.Text = sms(NumO(i))
             Else
                If Len(sms(NumO(i))) > 1 Then
                  .Cell(i + 3, 4).Range.Text = UCase(Mid(sms(NumO(i)), 1, 1)) & Mid(sms(NumO(i)), 2, Len(sms(NumO(i))) - 1)
                Else
                  .Cell(i + 3, 4).Range.Text = sms(NumO(i))
                End If
            End If
         Else
             Set Cel = .Cell(i + 3, 4)
             Call CellAddDropDownList(Cel)
           
          End If
       Next i
  End With
       
       '---On site test
    TbnB = tbn
    DelTab = ""
   
   For i = 1 To cno
     If Nat(NumO(i)) = UText(23450, 37327) Then
     

          DelTab = DelTab & "," & "Data Sheet for " & Chk(NumO(i))  ' translated comment
          TbCount1 = 0
          tb = WD.Tables.count - 1
          For j = 3 To tb - 3
            With WD.Tables(j)
               If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumO(i)) & "*" = True Then
                      TbCount1 = TbCount1 + 1
                      If TbCount1 > 1 Then
                          .Delete
                          j = j - 1
                      Else
                          Set Cel = .Cell(1, 1)
                          Call DeleteBookMark(Cel)
                          Cel.Range.Bookmarks.add Name:=Ino(NumO(i)) & "_Data", Range:=Cel.Range
                           Rn = .Rows.count
                              If Rn < ComItms + 2 Then
                                  .Cell(Rn, 1).Select
                                  Selection.InsertRowsBelow ComItms + 2 - Rn
                               ElseIf Rn > ComItms + 2 Then
                                  For k = Rn To ComItms + 2 + 1 Step -1
                                   .Rows(k).Delete
                                  Next k
                             End If
                             For k = 1 To ComItms
                               .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603917569 '
                               .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                               .Cell(k + 2, 1).Range.Text = ComItx(k)
                               .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                               .Cell(k + 2, 2).Range.Text = Spe(NumO(i))
                               .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                               .Cell(k + 2, 3).Range.Text = Tol(NumO(i))
                               .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                              Next k
                       End If
                End If
            End With
          Next j
55     tbn = 0
       tbn = getTableNumber(4)
         If tbn = 0 Then
            msg = msg & UText(26410, 21457, 29616) & "4. Product Dimension & Weight" & UText(30340, 34920, 26684) & "，On Site Test " & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
           "4. Product Dimension & Weight table is not found, Data sheet for On Site Test form can not generate" & CHR(10)
            GoTo 60
        End If
57      If TbCount1 = 0 Then
              Call MakeNormalDataSheet(tbn, ComItms)
              With WD.Tables(tbn)
                  Set Cel = .Cell(1, 1)
                  .Cell(1, 1).Range.Text = "Data Sheet for " & Chk(NumO(i))
                   Call DeleteBookMark(Cel)
                   Cel.Range.Bookmarks.add Name:=Ino(NumO(i)) & "_Data", Range:=Cel.Range
             
                   For j = 1 To ComItms
                     .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                     .Cell(j + 2, 1).Range.Text = ComItx(j)
                     .Cell(j + 2, 2).Range.Text = Spe(NumO(i))
                     .Cell(j + 2, 3).Range.Text = Tol(NumO(i))
                    Next j
               End With
         End If
       End If
58   Next i

    
       
60 '---------------------------------------------------4. Product Dimension & Weight
  
    tbn = 0
    tbn = getTableNumber(4)
      If tbn = 0 Then
          msg = msg & UText(26410, 21457, 29616) & "4. Product Dimension & Weight" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
          "4. Product Dimension & Weight table is not found, the check item for Product Dimension & Weight is not imported" & CHR(10)
          GoTo 80
      End If
70  With WD.Tables(tbn)
        Rn = .Rows.count
         If Rn = 3 Then
           .Cell(Rn, 1).Select
            Selection.InsertRowsBelow 1
           .Rows(Rn + 1).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603917569 '
             Rn = .Rows.count
          End If

        If Rn < cnd + 3 Then
         .Rows(4).Select
          Selection.Copy
          For j = 1 To cnd + 3 - Rn
          .Rows(Rn).Select
          StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
            Selection.Paste
           Next j
        ElseIf Rn > cnd + 3 Then
          For j = Rn To cnd + 3 + 1 Step -1
          .Rows(j).Delete
          Next j
       End If
   
       For i = 1 To cnd
        .Rows(i + 3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Cell(i + 3, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cell(i + 3, 1).Range.Text = "4." & i
         If Len(Chk(NumD(i))) > 1 Then
           KwChk = UCase(Mid(Chk(NumD(i)), 1, 1)) & Mid(Chk(NumD(i)), 2, Len(Chk(NumD(i))) - 1)
         Else
           KwChk = Chk(NumD(i))
         End If
         If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
         
        .Cell(i + 3, 2).Range.Text = KwChk
         Set Cel = .Cell(i + 3, 2)
         Call DeleteBookMark(Cel)
         Cel.Range.Bookmarks.add Name:=Ino(NumD(i)), Range:=Cel.Range
         '.Cell(i + 3, 3).Range.Text = UCase(Mid(Cri(NumD(i)), 1, 1)) & Mid(Cri(NumD(i)), 2, Len(Cri(NumD(i))) - 1)
        
        
        If sms(NumD(i)) <> "" Then
            If ComItms = 1 Then
              sms(NumD(i)) = Replace(sms(NumD(i)), " /item", "")
              sms(NumD(i)) = Replace(sms(NumD(i)), " /Item", "")
              .Cell(i + 3, 3).Range.Text = sms(NumD(i))
            Else
               If Len(sms(NumD(i))) > 1 Then
                  .Cell(i + 3, 3).Range.Text = UCase(Mid(sms(NumD(i)), 1, 1)) & Mid(sms(NumD(i)), 2, Len(sms(NumD(i))) - 1)
               Else
                  .Cell(i + 3, 3).Range.Text = sms(NumD(i))
               End If
            End If
         Else
             Set Cel = .Cell(i + 3, 3)
               Call CellAddDropDownList(Cel)
            
          End If
        
       Next i
  End With
   
        'Product Dimension & Weight---
     TbnB = tbn
     DelTab = ""
     
 Dim WithShoe As String
 WithShoe = ""
   For i = 1 To cnd
     If Nat(NumD(i)) = UText(23450, 37327) Then

          DelTab = DelTab & "," & "Data Sheet for " & Chk(NumD(i))  ' translated comment
          tb = WD.Tables.count - 1
        
          Select Case TableChoise
           Case 1
             TbCount1 = 0
             If Ino(NumD(i)) = "TSG056" Then   ' translated comment
                WithShoe = DelTab
               For j = 3 To tb - 3
                With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumD(i)) & "*" = True Then
                     If .Columns.count < 10 Then
                       .Delete
                       j = j - 1
                     Else
                       TbCount1 = TbCount1 + 1
                       If TbCount1 > ComItms Then
                        .Delete
                        j = j - 1
                       Else
                         Set Cel = .Cell(1, 1)
                         Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                         Call DeleteBookMark(Cel)
                         Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data" & TbCount1, Range:=Cel.Range
                         Set Cel = .Cell(2, 1)
                         Cel.Range.Text = "Item No.:" & ComItx(TbCount1)
                       End If
                     End If
                  End If
               End With
              Next j
              GoTo 75
            Else
             '==  type1
               TbCount1 = 0
              
               For j = 3 To tb - 3
                 With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumD(i)) & "*" = True Then
                    If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                     TbCount1 = TbCount1 + 1
                     If TbCount1 > 1 Then
                        .Delete
                        j = j - 1
                     Else
                       Set Cel = .Cell(1, 1)
                       Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                       Call DeleteBookMark(Cel)
                       Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data", Range:=Cel.Range
                       Rn = .Rows.count
                       If Rn < ComItms + 2 Then
                         .Cell(Rn, 1).Select
                          Selection.InsertRowsBelow ComItms + 2 - Rn
                          ElseIf Rn > ComItms + 2 Then
                          For k = Rn To ComItms + 2 + 1 Step -1
                            .Rows(k).Delete
                         Next k
                       End If
                       For k = 1 To ComItms
                        .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                        .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                        .Cell(k + 2, 1).Range.Text = ComItx(k)
                        .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                        .Cell(k + 2, 2).Range.Text = Spe(NumD(i))
                        .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                        .Cell(k + 2, 3).Range.Text = Tol(NumD(i))
                        .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                      Next k
                    End If
                    GoTo 78
                  End If
                 End If
                End With
            Next j
         End If
       Case 2
            TbCount2 = 0
            For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumD(i)) & "*" = True Then
                      If .Rows(2).Cells.count = 1 Then
                        TbCount2 = TbCount2 + 1
                         If TbCount2 > ComItms Then
                           .Delete
                           j = j - 1
                         Else
                           Set Cel = .Cell(1, 1)
                           Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                           Call DeleteBookMark(Cel)
                           Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data" & TbCount2, Range:=Cel.Range
                           Set Cel = .Cell(2, 1)
                           Cel.Range.Text = "Item No.:" & ComItx(TbCount2)
                        End If
                    End If
                  End If
                 End With
            Next j
       Case 3
            TbCount1 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
                 With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumD(i)) & "*" = True Then
                    If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                      TbCount1 = TbCount1 + 1
                      If TbCount1 > 1 Then
                         .Delete
                         j = j - 1
                      Else
                         Set Cel = .Cell(1, 1)
                         Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                         Call DeleteBookMark(Cel)
                         Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data", Range:=Cel.Range
                         Rn = .Rows.count
                           If Rn < ComItms + 2 Then
                             .Cell(Rn, 1).Select
                              Selection.InsertRowsBelow ComItms + 2 - Rn
                              ElseIf Rn > ComItms + 2 Then
                              For k = Rn To ComItms + 2 + 1 Step -1
                               .Rows(k).Delete
                              Next k
                           End If
                           For k = 1 To ComItms
                               .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                               .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                               .Cell(k + 2, 1).Range.Text = ComItx(k)
                               .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                               .Cell(k + 2, 2).Range.Text = Spe(NumD(i))
                               .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                               .Cell(k + 2, 3).Range.Text = Tol(NumD(i))
                              .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                            Next k
                        End If
                      End If
                  End If
                End With
            Next j
            
            TbCount2 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumD(i)) & "*" = True Then
                      If .Rows(2).Cells.count = 1 Then
                         TbCount2 = TbCount2 + 1
                         If TbCount2 > ComItms Then
                           .Delete
                           j = j - 1
                         Else
                           Set Cel = .Cell(1, 1)
                           Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                           Call DeleteBookMark(Cel)
                           Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data" & TbCount2, Range:=Cel.Range
                           Set Cel = .Cell(2, 1)
                           Cel.Range.Text = "Item No.:" & ComItx(TbCount2)
                        End If
                     End If
                   End If
                 End With
            Next j
       
     End Select
75      tbn = 0
        tbn = getTableNumber(5)
         If tbn = 0 Then
           msg = msg & UText(26410, 21457, 29616) & "5. Style / Material / Construction" & UText(34920, 26684) & "，Product Dimension & Weight " & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
           "5. Style / Material / Construction table is not found, Data sheet for Product Dimension & Weight form can not generate" & CHR(10)
           GoTo 80
        End If
77      Select Case TableChoise
         Case "1"
         If Ino(NumD(i)) = "TSG056" Then   ' translated comment
            For j = 1 To ComItms - TbCount1
              tbn = getTableNumber(5)
              Call MakeShoesDataSheet(tbn, ComItms)
              With WD.Tables(tbn)
                     Set Cel = .Cell(1, 1)
                     Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                     Call DeleteBookMark(Cel)
                     Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data" & "_" & TbCount1 + j, Range:=Cel.Range
                     Set Cel = .Cell(2, 1)
                     Cel.Range.Text = "Item No.:" & ComItx(TbCount1 + j)
              End With
            Next j
         Else
              If TbCount1 = 0 Then
                tbn = getTableNumber(5)
                Call MakeNormalDataSheet(tbn, ComItms)
                With WD.Tables(tbn)
                 Set Cel = .Cell(1, 1)
                  Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                  Call DeleteBookMark(Cel)
                  Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data", Range:=Cel.Range
                  For j = 1 To ComItms
                   .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                   .Cell(j + 2, 1).Range.Text = ComItx(j)
                   .Cell(j + 2, 2).Range.Text = Spe(NumD(i))
                   .Cell(j + 2, 3).Range.Text = Tol(NumD(i))
                  Next j
                End With
             End If
        End If
        Case "2"
            For j = 1 To ComItms - TbCount2
                   tbn = getTableNumber(5)
                   Call MakeWithPicDataSheet(tbn, ComItms)
                   With WD.Tables(tbn)
                     Set Cel = .Cell(1, 1)
                     Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                     Call DeleteBookMark(Cel)
                     Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                     Set Cel = .Cell(2, 1)
                     Cel.Range.Text = "Item No.:" & ComItx(j + TbCount2)
               
               'For j = 1 To ComItms 'MakeWithPicDataSheet_combined
                '.cell((j - 1) * 6 + 2, 1).Range.Text = "Item No.:" & Comitx(j)
                '.cell(j + 2, 2).Range.Text = Spe(NumD(i))
               ' .cell(j + 2, 3).Range.Text = Tol(NumD(i))
               'Next j
              End With
            Next j
         Case "3"
              If TbCount1 = 0 Then
                tbn = getTableNumber(5)
                Call MakeNormalDataSheet(tbn, ComItms)
                With WD.Tables(tbn)
                   Set Cel = .Cell(1, 1)
                   Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                  Call DeleteBookMark(Cel)
                  Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data", Range:=Cel.Range
                  For j = 1 To ComItms
                     .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                     .Cell(j + 2, 1).Range.Text = ComItx(j)
                     .Cell(j + 2, 2).Range.Text = Spe(NumD(i))
                    .Cell(j + 2, 3).Range.Text = Tol(NumD(i))
                  Next j
                End With
             End If
              
              For j = 1 To ComItms - TbCount2
                  tbn = getTableNumber(5)
                  Call MakeWithPicDataSheet(tbn, ComItms)
                     With WD.Tables(tbn)
                       Set Cel = .Cell(1, 1)
                        Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                        Call DeleteBookMark(Cel)
                        Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                        Set Cel = .Cell(2, 1)
                        Cel.Range.Text = "Item No.:" & ComItx(j + TbCount2)
               'For j = 1 To ComItms 'MakeWithPicDataSheet_combined
                '.cell((j - 1) * 6 + 2, 1).Range.Text = "Item No.:" & Comitx(j)
                '.cell(j + 2, 2).Range.Text = Spe(NumD(i))
               ' .cell(j + 2, 3).Range.Text = Tol(NumD(i))
               'Next j
                   End With
             Next j
             
        End Select
        
       End If
78  Next i
    
    
 
       ' translated comment
       If WithShoe = "" Then
         tb = WD.Tables.count - 1
         For i = tb To 3 Step -1
           With WD.Tables(i)
             If .Cell(1, 1).Range.Text Like "*Data Sheet for*" = True And LCase(.Cell(1, 1).Range.Text) Like "*length of footwear*" = True Then .Delete
           End With
         Next i
       End If
       
   
        
80     '---------------------------------------------------5. Style / Material / Construction
     tbn = 0
     tbn = getTableNumber(5)
      If tbn = 0 Then
          msg = msg & UText(26410, 21457, 29616) & "5. Style / Material / Construction" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
         "5. Style / Material / Construction table is not found, the check item for Style / Material / Construction is not imported" & CHR(10)
          GoTo 100
     End If
90  With WD.Tables(tbn)
        Rn = .Rows.count
         If Rn = 3 Then
           .Cell(Rn, 1).Select
            Selection.InsertRowsBelow 1
           .Rows(Rn + 1).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603917569 '
             Rn = .Rows.count
          End If

        If Rn < cns + 3 Then
         .Rows(4).Select
          Selection.Copy
          For j = 1 To cns + 3 - Rn
          .Rows(Rn).Select
          StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
            Selection.Paste
           Next j
        ElseIf Rn > cns + S Then
          For j = Rn To cns + 3 + 1 Step -1
          .Rows(j).Delete
          Next j
       End If
   
       For i = 1 To cns
        .Rows(i + 3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Cell(i + 3, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cell(i + 3, 1).Range.Text = "5." & i
         If Len(Chk(NumS(i))) > 1 Then
           KwChk = UCase(Mid(Chk(NumS(i)), 1, 1)) & Mid(Chk(NumS(i)), 2, Len(Chk(NumS(i))) - 1)
         Else
           KwChk = Chk(NumS(i))
         End If
         If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
         
        .Cell(i + 3, 2).Range.Text = KwChk
         Set Cel = .Cell(i + 3, 2)
         Call DeleteBookMark(Cel)
         Cel.Range.Bookmarks.add Name:=Ino(NumS(i)), Range:=Cel.Range
         '.Cell(i + 3, 3).Range.Text = UCase(Mid(Cri(NumD(i)), 1, 1)) & Mid(Cri(NumD(i)), 2, Len(Cri(NumD(i))) - 1)
        
         
          
        If sms(NumS(i)) <> "" Then
            If ComItms = 1 Then
                sms(NumS(i)) = Replace(sms(NumS(i)), " /item", "")
                sms(NumS(i)) = Replace(sms(NumS(i)), " /Item", "")
                .Cell(i + 3, 3).Range.Text = sms(NumS(i))
            Else
                 If Len(sms(NumS(i))) > 1 Then
                    .Cell(i + 3, 3).Range.Text = UCase(Mid(sms(NumS(i)), 1, 1)) & Mid(sms(NumS(i)), 2, Len(sms(NumS(i))) - 1)
                 Else
                    .Cell(i + 3, 3).Range.Text = sms(NumS(i))
                 End If
            End If
         Else
             Set Cel = .Cell(i + 3, 3)
              Call CellAddDropDownList(Cel)
            
         End If
        
91       Next i
  End With
   
        '5. Style / Material / Construction---
    TbnB = tbn
     DelTab = ""

   For i = 1 To cns
     If Nat(NumS(i)) = UText(23450, 37327) Then

       DelTab = DelTab & "," & "Data Sheet for " & Chk(NumS(i))  ' translated comment
       If Ino(NumS(i)) <> "TEE086" Then    'check item was not Constructional Data Form (CDF)
            'If Chk(NumS(i)) Like "*Constructional Data Form*" = False Then
           TbCount1 = 0
           tb = WD.Tables.count - 1
           For j = 3 To tb - 3
            With WD.Tables(j)
               If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumS(i)) & "*" = True Then
                  TbCount1 = TbCount1 + 1
                  If TbCount1 > 1 Then
                     .Delete
                     j = j - 1
                    
                  Else
                     Set Cel = .Cell(1, 1)
                     Call DeleteBookMark(Cel)
                     Cel.Range.Bookmarks.add Name:=Ino(NumS(i)) & "_Data", Range:=Cel.Range
                     Rn = .Rows.count
                     If Rn < ComItms + 2 Then
                        .Cell(Rn, 1).Select
                        Selection.InsertRowsBelow ComItms + 2 - Rn
                     ElseIf Rn > ComItms + 2 Then
                        For k = Rn To ComItms + 2 + 1 Step -1
                         .Rows(k).Delete
                       Next k
                     End If
                     For k = 1 To ComItms
                      .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                      .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                      .Cell(k + 2, 1).Range.Text = ComItx(k)
                      .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                      .Cell(k + 2, 2).Range.Text = Spe(NumS(i))
                      .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                      .Cell(k + 2, 3).Range.Text = Tol(NumS(i))
                      .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                     Next k
                  End If
              End If
            End With
           Next j
        'End If
95      tbn = 0
        tbn = getTableNumber(6)
         If tbn = 0 Then
             msg = msg & UText(26410, 21457, 29616) & "6. Product Color" & UText(34920, 26684) & "，Style / Material / Construction " & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
            "6. Product Color table is not found, Data sheet for Style / Material / Construction form can not generate" & CHR(10)
            GoTo 110
        End If
        
97       If TbCount1 = 0 Then
               Call MakeNormalDataSheet(tbn, ComItms)
               With WD.Tables(tbn)
                  Set Cel = .Cell(1, 1)
                 Cel.Range.Text = "Data Sheet for " & Chk(NumS(i))
                 Call DeleteBookMark(Cel)
                 Cel.Range.Bookmarks.add Name:=Ino(NumS(i)) & "_Data", Range:=Cel.Range
                 For j = 1 To ComItms
                    .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                    .Cell(j + 2, 1).Range.Text = ComItx(j)
                    .Cell(j + 2, 2).Range.Text = Spe(NumS(i))
                    .Cell(j + 2, 3).Range.Text = Tol(NumS(i))
                 Next j
              End With
          End If
        End If
      End If
     Next i
     
     '===== ，CDF
     For i = 1 To cns
        If Nat(NumS(i)) = UText(23450, 37327) Then
          If Ino(NumS(i)) = "TEE086" Then    'check item was Constructional Data Form (CDF)
            TbCount1 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
               With WD.Tables(j)
                 If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumS(i)) & "*" = True Then
                   TbCount1 = TbCount1 + 1
                   If TbCount1 > ComItms Then
                      .Delete
                      j = j - 1
                    Else
                       Set Cel = .Cell(1, 1)
                       Call DeleteBookMark(Cel)
                       Cel.Range.Bookmarks.add Name:=Ino(NumS(i)) & "_Data" & "_" & TbCount1, Range:=Cel.Range
                       .Cell(2, 1).Range.Text = "Item No.:" & ComItx(TbCount1)
                    End If
                End If
              End With
           Next j
100        tbn = 0
           tbn = getTableNumber(6)
           If tbn = 0 Then
               msg = msg & UText(26410, 21457, 29616) & "6. Product Color" & UText(34920, 26684) & "，Style / Material / Construction " & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
               "6. Product Color table is not found, Data sheet for Style / Material / Construction form can not generate" & CHR(10)
               GoTo 110
            End If
            For j = 1 To ComItms - TbCount1
               Call MakeCDFDataSheet(tbn, ComItms)
               With WD.Tables(tbn)
                  Set Cel = .Cell(1, 1)
                  Cel.Range.Text = "Data Sheet for " & Chk(NumS(i))
                  Call DeleteBookMark(Cel)
                  Cel.Range.Bookmarks.add Name:=Ino(NumS(i)) & "_Data" & "_" & TbCount1 + j, Range:=Cel.Range
                  .Cell(2, 1).Range.Text = "Item No.:" & ComItx(TbCount1 + j)
                End With
            Next j
        End If
       End If
108  Next i
       TbnE = getTableNumber(6)
     
110  '------------------------------------------------6. Product Color

   tbn = 0
   tbn = getTableNumber(6)
      If tbn = 0 Then
         msg = msg & UText(26410, 21457, 29616) & "6. Product Color" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
        "6. Product Color table is not found, the check item for Product Color is not imported" & CHR(10)
        GoTo 130
     End If
120  With WD.Tables(tbn)
        Rn = .Rows.count
         If Rn = 3 Then
           .Cell(Rn, 1).Select
            Selection.InsertRowsBelow 1
           .Rows(Rn + 1).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603917569 '
             Rn = .Rows.count
          End If

        If Rn < cnc + 3 Then
         .Rows(4).Select
          Selection.Copy
          For j = 1 To cnc + 3 - Rn
          .Rows(Rn).Select
          StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
            Selection.Paste
           Next j
        ElseIf Rn > cnc + 3 Then
          For j = Rn To cnc + 3 + 1 Step -1
          .Rows(j).Delete
          Next j
       End If
   
       For i = 1 To cnc
        .Rows(i + 3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Cell(i + 3, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cell(i + 3, 1).Range.Text = "6." & i
         If Len(Chk(NumC(i))) > 1 Then
           KwChk = UCase(Mid(Chk(NumC(i)), 1, 1)) & Mid(Chk(NumC(i)), 2, Len(Chk(NumC(i))) - 1)
         Else
           KwChk = Chk(NumC(i))
         End If
         If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
         
        .Cell(i + 3, 2).Range.Text = KwChk
         Set Cel = .Cell(i + 3, 2)
         Call DeleteBookMark(Cel)
         Cel.Range.Bookmarks.add Name:=Ino(NumC(i)), Range:=Cel.Range
         '.Cell(i + 3, 3).Range.Text = UCase(Mid(Cri(NumC(i)), 1, 1)) & Mid(Cri(Numc(i)), 2, Len(Cri(Numc(i))) - 1)
        
        
        If sms(NumC(i)) <> "" Then
            If ComItms = 1 Then
               sms(NumC(i)) = Replace(sms(NumC(i)), " /item", "")
               sms(NumC(i)) = Replace(sms(NumC(i)), " /Item", "")
               .Cell(i + 3, 3).Range.Text = sms(NumC(i))
            Else
               If Len(sms(NumC(i))) > 1 Then
                   .Cell(i + 3, 3).Range.Text = UCase(Mid(sms(NumC(i)), 1, 1)) & Mid(sms(NumC(i)), 2, Len(sms(NumC(i))) - 1)
                Else
                   .Cell(i + 3, 3).Range.Text = sms(NumC(i))
                End If
            End If
         Else
           
             Set Cel = .Cell(i + 3, 3)
             Call CellAddDropDownList(Cel)
           
       End If
121  Next i
  End With
   
        '6. Product Color--
    TbnB = tbn
     DelTab = ""
   
   For i = 1 To cnc
     If Nat(NumC(i)) = UText(23450, 37327) Then

          DelTab = DelTab & "," & "Data Sheet for " & Chk(NumC(i))  ' translated comment
          TbCount1 = 0
          tb = WD.Tables.count - 1
          For j = 3 To tb - 3
            With WD.Tables(j)
               If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumC(i)) & "*" = True Then
                     TbCount1 = TbCount1 + 1
                     If TbCount1 > 1 Then
                       .Delete
                       j = j - 1
                     Else
                        Set Cel = .Cell(1, 1)
                        Call DeleteBookMark(Cel)
                        Cel.Range.Bookmarks.add Name:=Ino(NumC(i)) & "_Data", Range:=Cel.Range
                        Rn = .Rows.count
                        If Rn < ComItms + 2 Then
                           .Cell(Rn, 1).Select
                            Selection.InsertRowsBelow ComItms + 2 - Rn
                        ElseIf Rn > ComItms + 2 Then
                           For k = Rn To ComItms + 2 + 1 Step -1
                            .Rows(k).Delete
                          Next k
                        End If
                        For k = 1 To ComItms
                         .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                         .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                         .Cell(k + 2, 1).Range.Text = ComItx(k)
                         .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                         .Cell(k + 2, 2).Range.Text = Spe(NumC(i))
                         .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                         .Cell(k + 2, 3).Range.Text = Tol(NumC(i))
                        .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                       Next k
                  End If
              End If
           End With
      Next j

125     tbn = 0
        tbn = getTableNumber(7)
         If tbn = 0 Then
             msg = msg & UText(26410, 21457, 29616) & "7. Product Label / Marking" & UText(34920, 26684) & "，Product Color" & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
            "7. Product Label / Marking table is not found, Data sheet for Product Color form can not generate" & CHR(10)
             GoTo 130
         End If
127      If TbCount1 = 0 Then
              Call MakeNormalDataSheet(tbn, ComItms)
              With WD.Tables(tbn)
                 Set Cel = .Cell(1, 1)
                 Cel.Range.Text = "Data Sheet for " & Chk(NumC(i))
                 Call DeleteBookMark(Cel)
                 Cel.Range.Bookmarks.add Name:=Ino(NumC(i)) & "_Data", Range:=Cel.Range
             
                 For j = 1 To ComItms
                    .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                    .Cell(j + 2, 1).Range.Text = ComItx(j)
                    .Cell(j + 2, 2).Range.Text = Spe(NumC(i))
                    .Cell(j + 2, 3).Range.Text = Tol(NumC(i))
                  Next j
              End With
         End If
       End If
128  Next i
    
   
130    ''------------------------------------------7. Product Label / Marking

     tbn = 0
     tbn = getTableNumber(7)
     If tbn = 0 Then
         msg = msg & UText(26410, 21457, 29616) & "7. Product Label / Marking" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
        "7. Product Label / Marking table is not found, the check item for Product Label / Marking is not imported" & CHR(10)
        GoTo 150
     End If
140       With WD.Tables(tbn)
        
        Rn = .Rows.count
        If Rn = 3 Then
           .Cell(Rn, 1).Select
            Selection.InsertRowsBelow 1
           .Rows(Rn + 1).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603917569 '
             Rn = .Rows.count
        End If
        If Rn < cnl + 3 Then
         .Rows(4).Select
          Selection.Copy
          For j = 1 To cnl + 3 - Rn
          .Rows(Rn).Select
          StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
            Selection.Paste
           Next j
        ElseIf Rn > cnl + 3 Then
          For j = Rn To cnl + 3 + 1 Step -1
          .Rows(j).Delete
          Next j
       End If
   
       For i = 1 To cnl
         .Rows(i + 3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
         .Cell(i + 3, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
         
        .Cell(i + 3, 1).Range.Text = "7." & i
         If Len(Chk(NumL(i))) > 1 Then
           KwChk = UCase(Mid(Chk(NumL(i)), 1, 1)) & Mid(Chk(NumL(i)), 2, Len(Chk(NumL(i))) - 1)
         Else
            KwChk = Chk(NumL(i))
         End If
         If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
         
        .Cell(i + 3, 2).Range.Text = KwChk
         Set Cel = .Cell(i + 3, 2)
         Call DeleteBookMark(Cel)
         Cel.Range.Bookmarks.add Name:=Ino(NumL(i)), Range:=Cel.Range
         '.Cell(i + 3, 3).Range.Text = UCase(Mid(Cri(NumL(i)), 1, 1)) & Mid(Cri(NumL(i)), 2, Len(Cri(NumL(i))) - 1)
         
          
        If sms(NumL(i)) <> "" Then
           If ComItms = 1 Then
               sms(NumL(i)) = Replace(sms(NumL(i)), " /item", "")
               sms(NumL(i)) = Replace(sms(NumL(i)), " /Item", "")
               .Cell(i + 3, 3).Range.Text = sms(NumL(i))
           Else
               If Len(sms(NumL(i))) > 1 Then
                 .Cell(i + 3, 3).Range.Text = UCase(Mid(sms(NumL(i)), 1, 1)) & Mid(sms(NumL(i)), 2, Len(sms(NumL(i))) - 1)
               Else
                 .Cell(i + 3, 3).Range.Text = sms(NumL(i))
               End If
           End If
        Else
          Set Cel = .Cell(i + 3, 3)
           Call CellAddDropDownList(Cel)
            
        End If
        
141   Next i
  End With
        
        '7. Product Label / Marking--
    TbnB = tbn
     DelTab = ""
   
   For i = 1 To cnl
      If Nat(NumL(i)) = UText(23450, 37327) Then
         If Chk(NumL(i)) = "Readability test for Barcode,QR code" Then GoTo 148   'BadCode，GenerateBarCodeDataSheet
         DelTab = DelTab & "," & "Data Sheet for " & Chk(NumL(i))  ' translated comment
         Select Case TableChoise
         Case 1
             TbCount1 = 0
               For j = 3 To tb - 3
                With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumL(i)) & "*" = True Then
                       If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                            TbCount1 = TbCount1 + 1
                            If TbCount1 > 1 Then
                                 .Delete
                                  j = j - 1
                            Else
                                 Set Cel = .Cell(1, 1)
                                 Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                                 Call DeleteBookMark(Cel)
                                 Cel.Range.Bookmarks.add Name:=Ino((NumL(i))) & "_Data", Range:=Cel.Range
                                 Rn = .Rows.count
                                   If Rn < ComItms + 2 Then
                                      .Cell(Rn, 1).Select
                                       Selection.InsertRowsBelow ComItms + 2 - Rn
                                   ElseIf Rn > ComItms + 2 Then
                                       For k = Rn To ComItms + 2 + 1 Step -1
                                         .Rows(k).Delete
                                       Next k
                                   End If
                                   For k = 1 To ComItms
                                     .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                                      .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                                      .Cell(k + 2, 1).Range.Text = ComItx(k)
                                      .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                                      .Cell(k + 2, 2).Range.Text = Spe((NumL(i)))
                                      .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                                      .Cell(k + 2, 3).Range.Text = Tol(NumL(i))
                                      .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                                   Next k
                           End If
                           GoTo 148
                    
                    End If
                  End If
              End With
            Next j
       Case 2
            TbCount2 = 0
            For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumL(i)) & "*" = True Then
                      If .Rows(2).Cells.count = 1 Then
                        TbCount2 = TbCount2 + 1
                         If TbCount2 > ComItms Then
                           .Delete
                           j = j - 1
                         Else
                           Set Cel = .Cell(1, 1)
                           Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                           Call DeleteBookMark(Cel)
                           Cel.Range.Bookmarks.add Name:=Ino(NumL(i)) & "_Data" & TbCount2, Range:=Cel.Range
                           Set Cel = .Cell(2, 1)
                           Cel.Range.Text = "Item No.:" & ComItx(TbCount2)
                        End If
                    End If
                  End If
                 End With
            Next j
       Case 3
            TbCount1 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
                 With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumL(i)) & "*" = True Then
                    If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                      TbCount1 = TbCount1 + 1
                      If TbCount1 > 1 Then
                         .Delete
                         j = j - 1
                      Else
                         Set Cel = .Cell(1, 1)
                         Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                         Call DeleteBookMark(Cel)
                         Cel.Range.Bookmarks.add Name:=Ino(NumL(i)) & "_Data", Range:=Cel.Range
                         Rn = .Rows.count
                           If Rn < ComItms + 2 Then
                             .Cell(Rn, 1).Select
                              Selection.InsertRowsBelow ComItms + 2 - Rn
                              ElseIf Rn > ComItms + 2 Then
                              For k = Rn To ComItms + 2 + 1 Step -1
                               .Rows(k).Delete
                              Next k
                           End If
                           For k = 1 To ComItms
                               .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                               .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                               .Cell(k + 2, 1).Range.Text = ComItx(k)
                               .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                               .Cell(k + 2, 2).Range.Text = Spe(NumL(i))
                               .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                               .Cell(k + 2, 3).Range.Text = Tol(NumL(i))
                              .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                            Next k
                        End If
                      End If
                  End If
                End With
            Next j
            
            TbCount2 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumL(i)) & "*" = True Then
                      If .Rows(2).Cells.count = 1 Then
                         TbCount2 = TbCount2 + 1
                         If TbCount2 > ComItms Then
                           .Delete
                           j = j - 1
                         Else
                           Set Cel = .Cell(1, 1)
                           Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                           Call DeleteBookMark(Cel)
                           Cel.Range.Bookmarks.add Name:=Ino(NumL(i)) & "_Data" & TbCount2, Range:=Cel.Range
                           Set Cel = .Cell(2, 1)
                           Cel.Range.Text = "Item No.:" & ComItx(TbCount2)
                        End If
                     End If
                   End If
                 End With
            Next j
     End Select

145  tbn = 0
     tbn = getTableNumber(8)
         If tbn = 0 Then
            msg = msg & UText(26410, 21457, 29616) & "8. Shipping Mark" & UText(34920, 26684) & "，Product Label / Marking" & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
           "8. Shipping Mark table is not found, Data sheet for Product Label / Marking form can not generate" & CHR(10)
            GoTo 150
         End If
147      Select Case TableChoise
         Case "1"
             If TbCount1 = 0 Then
                tbn = getTableNumber(8)
                Call MakeNormalDataSheet(tbn, ComItms)
                With WD.Tables(tbn)
                 Set Cel = .Cell(1, 1)
                  Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                  Call DeleteBookMark(Cel)
                  Cel.Range.Bookmarks.add Name:=Ino(NumL(i)) & "_Data", Range:=Cel.Range
                  For j = 1 To ComItms
                   .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                   .Cell(j + 2, 1).Range.Text = ComItx(j)
                   .Cell(j + 2, 2).Range.Text = Spe(NumL(i))
                   .Cell(j + 2, 3).Range.Text = Tol(NumL(i))
                  Next j
            End With
         End If
 
        Case "2"
               For j = 1 To ComItms - TbCount2
                    tbn = getTableNumber(8)
                    Call MakeWithPicDataSheet(tbn, ComItms)
                    With WD.Tables(tbn)
                      Set Cel = .Cell(1, 1)
                      Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                      Call DeleteBookMark(Cel)
                      Cel.Range.Bookmarks.add Name:=Ino(NumL(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                      Set Cel = .Cell(2, 1)
                      Cel.Range.Text = "Item No.:" & ComItx(j + TbCount2)
           
                    End With
              Next j
         Case "3"
              If TbCount1 = 0 Then
                   tbn = getTableNumber(8)
                   Call MakeNormalDataSheet(tbn, ComItms)
                   With WD.Tables(tbn)
                       Set Cel = .Cell(1, 1)
                       Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                       Call DeleteBookMark(Cel)
                       Cel.Range.Bookmarks.add Name:=Ino(NumL(i)) & "_Data", Range:=Cel.Range
                       For j = 1 To ComItms
                          .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                          .Cell(j + 2, 1).Range.Text = ComItx(j)
                          .Cell(j + 2, 2).Range.Text = Spe(NumL(i))
                         .Cell(j + 2, 3).Range.Text = Tol(NumL(i))
                       Next j
                  End With
             End If
              
              For j = 1 To ComItms - TbCount2
                   tbn = getTableNumber(8)
                   Call MakeWithPicDataSheet(tbn, ComItms)
                     With WD.Tables(tbn)
                         Set Cel = .Cell(1, 1)
                         Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                         Call DeleteBookMark(Cel)
                         Cel.Range.Bookmarks.add Name:=Ino(NumL(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                         Set Cel = .Cell(2, 1)
                   End With
             Next j
        End Select
       End If
148  Next i
      
    
  
150  '-----------------------------------------8. Shipping Mark
     tbn = 0
     tbn = getTableNumber(8)
     tb = WD.Tables.count - 1
    If tbn = 0 Then
         msg = msg & UText(26410, 21457, 29616) & "8. Shipping Mark" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
        "8. Shipping Mark table is not found, the check item for  Shipping Mark is not imported" & CHR(10)
         GoTo 170
     End If
160  With WD.Tables(tbn)
        Rn = .Rows.count
         If Rn = 3 Then
           .Cell(Rn, 1).Select
            Selection.InsertRowsBelow 1
           .Rows(Rn + 1).Range.Cells.Shading.BackgroundPatternColor = -603914241 '-603917569 '
             Rn = .Rows.count
          End If

        If Rn < cnm + 3 Then
         .Rows(4).Select
          Selection.Copy
          For j = 1 To cnm + 3 - Rn
          .Rows(Rn).Select
          StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
            Selection.Paste
           Next j
        ElseIf Rn > cnm + 3 Then
          For j = Rn To cnm + 3 + 1 Step -1
          .Rows(j).Delete
          Next j
       End If
   
       For i = 1 To cnm
        .Rows(i + 3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Cell(i + 3, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cell(i + 3, 1).Range.Text = "8." & i
         If Len(Chk(NumM(i))) > 1 Then
           KwChk = UCase(Mid(Chk(NumM(i)), 1, 1)) & Mid(Chk(NumM(i)), 2, Len(Chk(NumM(i))) - 1)
         Else
           KwChk = Chk(NumM(i))
         End If
         If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
         
        .Cell(i + 3, 2).Range.Text = KwChk
         Set Cel = .Cell(i + 3, 2)
         Call DeleteBookMark(Cel)
         Cel.Range.Bookmarks.add Name:=Ino(NumM(i)), Range:=Cel.Range
         '.Cell(i + 3, 3).Range.Text = UCase(Mid(Cri(Numm(i)), 1, 1)) & Mid(Cri(Numm(i)), 2, Len(Cri(Numm(i))) - 1)
        
        If sms(NumM(i)) <> "" Then
           If ComItms = 1 Then
              sms(NumM(i)) = Replace(sms(NumM(i)), " /item", "")
              sms(NumM(i)) = Replace(sms(NumM(i)), " /Item", "")
              .Cell(i + 3, 3).Range.Text = sms(NumM(i))
           Else
              If Len(sms(NumM(i))) > 1 Then
                 .Cell(i + 3, 3).Range.Text = UCase(Mid(sms(NumM(i)), 1, 1)) & Mid(sms(NumM(i)), 2, Len(sms(NumM(i))) - 1)
              Else
                 .Cell(i + 3, 3).Range.Text = sms(NumM(i))
              End If
           End If
        Else
             Set Cel = .Cell(i + 3, 3)
             Call CellAddDropDownList(Cel)
          
        End If
        
161  Next i
  End With
   
        '8. Shipping Mark-
    TbnB = tbn
     DelTab = ""
   
   For i = 1 To cnm
     If Nat(NumM(i)) = UText(23450, 37327) Then

          DelTab = DelTab & "," & "Data Sheet for " & Chk(NumM(i))  ' translated comment
         Select Case TableChoise
         Case 1
              TbCount1 = 0
               For j = 3 To tb - 3
                With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumM(i)) & "*" = True Then
                       If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                            TbCount1 = TbCount1 + 1
                            If TbCount1 > 1 Then
                                 .Delete
                                  j = j - 1
                            Else
                                 Set Cel = .Cell(1, 1)
                                 Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                                 Call DeleteBookMark(Cel)
                                 Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data", Range:=Cel.Range
                                 Rn = .Rows.count
                                   If Rn < ComItms + 2 Then
                                      .Cell(Rn, 1).Select
                                       Selection.InsertRowsBelow ComItms + 2 - Rn
                                   ElseIf Rn > ComItms + 2 Then
                                       For k = Rn To ComItms + 2 + 1 Step -1
                                         .Rows(k).Delete
                                       Next k
                                   End If
                                   For k = 1 To ComItms
                                     .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                                      .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                                      .Cell(k + 2, 1).Range.Text = ComItx(k)
                                      .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                                      .Cell(k + 2, 2).Range.Text = Spe(NumM(i))
                                      .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                                      .Cell(k + 2, 3).Range.Text = Tol(NumM(i))
                                      .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                                   Next k
                           End If
                           GoTo 168
                    
                    End If
                  End If
              End With
            Next j
       Case 2
            TbCount2 = 0
            For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumM(i)) & "*" = True Then
                      If .Rows(2).Cells.count = 1 Then
                        TbCount2 = TbCount2 + 1
                         If TbCount2 > ComItms Then
                           .Delete
                           j = j - 1
                         Else
                           Set Cel = .Cell(1, 1)
                           Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                           Call DeleteBookMark(Cel)
                           Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data" & TbCount2, Range:=Cel.Range
                           Set Cel = .Cell(2, 1)
                           Cel.Range.Text = "Item No.:" & ComItx(TbCount2)
                        End If
                    End If
                  End If
                 End With
            Next j
       Case 3
            TbCount1 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
                 With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumM(i)) & "*" = True Then
                    If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                      TbCount1 = TbCount1 + 1
                      If TbCount1 > 1 Then
                         .Delete
                         j = j - 1
                      Else
                         Set Cel = .Cell(1, 1)
                         Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                         Call DeleteBookMark(Cel)
                         Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data", Range:=Cel.Range
                         Rn = .Rows.count
                           If Rn < ComItms + 2 Then
                             .Cell(Rn, 1).Select
                              Selection.InsertRowsBelow ComItms + 2 - Rn
                              ElseIf Rn > ComItms + 2 Then
                              For k = Rn To ComItms + 2 + 1 Step -1
                               .Rows(k).Delete
                              Next k
                           End If
                           For k = 1 To ComItms
                               .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                               .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                               .Cell(k + 2, 1).Range.Text = ComItx(k)
                               .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                               .Cell(k + 2, 2).Range.Text = Spe(NumM(i))
                               .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                               .Cell(k + 2, 3).Range.Text = Tol(NumM(i))
                              .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                            Next k
                        End If
                      End If
                  End If
                End With
            Next j
            
            TbCount2 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumM(i)) & "*" = True Then
                      If .Rows(2).Cells.count = 1 Then
                         TbCount2 = TbCount2 + 1
                         If TbCount2 > ComItms Then
                           .Delete
                           j = j - 1
                         Else
                           Set Cel = .Cell(1, 1)
                           Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                           Call DeleteBookMark(Cel)
                           Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data" & TbCount2, Range:=Cel.Range
                           Set Cel = .Cell(2, 1)
                           Cel.Range.Text = "Item No.:" & ComItx(TbCount2)
                        End If
                     End If
                   End If
                 End With
            Next j
     End Select
       
165   tbn = 0
      tbn = getTableNumber(9)
      If tbn = 0 Then
          msg = msg & UText(26410, 21457, 29616) & "9. Packing / Packaging" & UText(34920, 26684) & "，Shipping Mark-" & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
         "9. Packing / Packaging table is not found, Data sheet for Shipping Mark form can not generate" & CHR(10)
          GoTo 170
      End If
167     Select Case TableChoise
         Case "1"
             If TbCount1 = 0 Then
                 tbn = getTableNumber(9)
                 Call MakeNormalDataSheet(tbn, ComItms)
                With WD.Tables(tbn)
                 Set Cel = .Cell(1, 1)
                  Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                  Call DeleteBookMark(Cel)
                  Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data", Range:=Cel.Range
                  For j = 1 To ComItms
                   .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                   .Cell(j + 2, 1).Range.Text = ComItx(j)
                   .Cell(j + 2, 2).Range.Text = Spe(NumM(i))
                   .Cell(j + 2, 3).Range.Text = Tol(NumM(i))
                  Next j
            End With
         End If
 
        Case "2"
               For j = 1 To ComItms - TbCount2
                    tbn = getTableNumber(9)
                    Call MakeWithPicDataSheet(tbn, ComItms)
                    With WD.Tables(tbn)
                      Set Cel = .Cell(1, 1)
                      Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                      Call DeleteBookMark(Cel)
                      Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                      Set Cel = .Cell(2, 1)
                      Cel.Range.Text = "Item No.:" & ComItx(j + TbCount2)
           
                    End With
              Next j
         Case "3"
              If TbCount1 = 0 Then
                   tbn = getTableNumber(9)
                   Call MakeNormalDataSheet(tbn, ComItms)
                   With WD.Tables(tbn)
                       Set Cel = .Cell(1, 1)
                       Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                       Call DeleteBookMark(Cel)
                       Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data", Range:=Cel.Range
                       For j = 1 To ComItms
                          .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                          .Cell(j + 2, 1).Range.Text = ComItx(j)
                          .Cell(j + 2, 2).Range.Text = Spe(NumM(i))
                         .Cell(j + 2, 3).Range.Text = Tol(NumM(i))
                       Next j
                  End With
             End If
              
              For j = 1 To ComItms - TbCount2
                   tbn = getTableNumber(9)
                   Call MakeWithPicDataSheet(tbn, ComItms)
                     With WD.Tables(tbn)
                         Set Cel = .Cell(1, 1)
                         Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                         Call DeleteBookMark(Cel)
                         Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                         Set Cel = .Cell(2, 1)
                   End With
             Next j
        End Select
       End If
168  Next i
    
 
  
170 '------------------------------------------package packaging
     tbn = 0
     tbn = getTableNumber(9)
     If tbn = 0 Then
        msg = msg & UText(26410, 21457, 29616) & "9. Packing / Packaging" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
       "9. Packing / Packaging is not found, the check item for Packing / Packaging is not imported" & CHR(10)
       GoTo 190
     End If
180  With WD.Tables(tbn)
        Rn = .Rows.count

        If Rn < cnp + 3 Then
         .Rows(4).Select
          Selection.Copy
          For j = 1 To cnp + 3 - Rn
          .Rows(Rn).Select
            StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
            Selection.Paste
           Next j
        ElseIf Rn > cnp + 3 Then
          For j = Rn To cnp + 3 + 1 Step -1
          .Rows(j).Delete
          Next j
       End If
   
       For i = 1 To cnp
        .Rows(i + 3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Cell(i + 3, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
         
        .Cell(i + 3, 1).Range.Text = "9." & i
         .Cell(i + 3, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
         If Len(Chk(NumP(i))) > 1 Then
            KwChk = UCase(Mid(Chk(NumP(i)), 1, 1)) & Mid(Chk(NumP(i)), 2, Len(Chk(NumP(i))) - 1)
         Else
            KwChk = Chk(NumP(i))
         End If
         If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
        
      
         .Cell(i + 3, 2).Range.Text = KwChk
          Set Cel = .Cell(i + 3, 2)
          Call DeleteBookMark(Cel)
          Cel.Range.Bookmarks.add Name:=Ino(NumP(i)), Range:=Cel.Range
         '.Cell(i + 3, 3).Range.Text = UCase(Mid(Cri(NumP(i)), 1, 1)) & Mid(Cri(NumP(i)), 2, Len(Cri(NumP(i))) - 1)
          If sms(NumP(i)) <> "" Then
              If ComItms = 1 Then
                sms(NumP(i)) = Replace(sms(NumP(i)), " /item", "")
                sms(NumP(i)) = Replace(sms(NumP(i)), " /Item", "")
                .Cell(i + 3, 3).Range.Text = sms(NumP(i))
               Else
                  If Len(sms(NumP(i))) > 1 Then
                    .Cell(i + 3, 3).Range.Text = UCase(Mid(sms(NumP(i)), 1, 1)) & Mid(sms(NumP(i)), 2, Len(sms(NumP(i))) - 1)
                  Else
                     .Cell(i + 3, 3).Range.Text = sms(NumP(i))
                  End If
               End If
          Else
             Set Cel = .Cell(i + 3, 3)
               Call CellAddDropDownList(Cel)
           
          End If
       Next i
  End With
      
      'Packing / Packaging---
     TbnB = tbn
     DelTab = ""
   
   For i = 1 To cnp
     If Nat(NumP(i)) = UText(23450, 37327) Then

          DelTab = DelTab & "," & "Data Sheet for " & Chk(NumP(i))  ' translated comment
          TbCount1 = 0
           tb = WD.Tables.count - 1
           For j = 3 To tb - 3
            With WD.Tables(j)
               If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumP(i)) & "*" = True Then
                  TbCount1 = TbCount1 + 1
                  If TbCount1 > 1 Then
                     .Delete
                     j = j - 1
                  Else
                     Set Cel = .Cell(1, 1)
                     Call DeleteBookMark(Cel)
                     Cel.Range.Bookmarks.add Name:=Ino(NumP(i)) & "_Data", Range:=Cel.Range
                     Rn = .Rows.count
                     If Rn < ComItms + 2 Then
                        .Cell(Rn, 1).Select
                        Selection.InsertRowsBelow ComItms + 2 - Rn
                     ElseIf Rn > ComItms + 2 Then
                        For k = Rn To ComItms + 2 + 1 Step -1
                         .Rows(k).Delete
                       Next k
                    End If
                    For k = 1 To ComItms
                     .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                     .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                     .Cell(k + 2, 1).Range.Text = ComItx(k)
                     .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                     .Cell(k + 2, 2).Range.Text = Spe(NumP(i))
                     .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                     .Cell(k + 2, 3).Range.Text = Tol(NumP(i))
                     .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                    Next k
                   End If
                 End If
            End With
         Next j
         

185       tbn = 0
         tbn = getTableNumber(10)
          If tbn = 0 Then
            msg = msg & UText(26410, 21457, 29616) & "10. Other Note" & UText(34920, 26684) & "Packing / Packaging" & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
            "10. Other Note table is not found, Data sheet for Packing / Packagingform can not generate" & CHR(10)
            GoTo 190
          End If
187     If TbCount1 = 0 Then
               Call MakeNormalDataSheet(tbn, ComItms)
               With WD.Tables(tbn)
                  Set Cel = .Cell(1, 1)
                 Cel.Range.Text = "Data Sheet for " & Chk(NumP(i))
                 Call DeleteBookMark(Cel)
                 Cel.Range.Bookmarks.add Name:=Ino(NumP(i)) & "_Data", Range:=Cel.Range
                 For j = 1 To ComItms
                    .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                    .Cell(j + 2, 1).Range.Text = ComItx(j)
                    .Cell(j + 2, 2).Range.Text = Spe(NumP(i))
                    .Cell(j + 2, 3).Range.Text = Tol(NumP(i))
                 Next j
              End With
        End If
      End If
188   Next i
     
     
   
      
190      'CDF
    
      For k = 1 To 3
         If Cat_Large(k) = UText(30005, 23376, 30005, 22120, 31867) Then
           GoTo 300
          End If
      Next k
      
           tb = WD.Tables.count - 1           '5. Style / Material / Construction,CDF
           For j = tb To 3 Step -1
             With WD.Tables(j)
               If .Cell(1, 1).Range.Text Like "*Data Sheet for Constructional Data Form (CDF)*" = True Then
                     WD.Tables(j).Select
                     Selection.Tables(1).Delete
                     Selection.Delete
               ElseIf .Cell(1, 1).Range.Text Like "*5. Style / Material / Construction*" = True Then
                  GoTo 300
               End If
             End With
203        Next j

 
300  If msg <> "" Then
      MsgBox msg & CHR(10) & CHR(10) & UText(27979, 35797, 39033, 30446, 23548, 20837, 23436, 25104) & "，" & UText(35831, 26681, 25454, 23454, 38469, 24773, 20917, 36827, 34892, 20462, 25913, 25110, 34917, 20805) & "," & UText(35831, 19981, 35201, 20877, 36816, 34892, 26412, 31243, 24207, 21542, 21017, 20462, 25913, 25110, 34917, 20805, 30340, 25968, 25454, 23558, 34987, 35206, 30422) & "!" & CHR(10) & _
     "Complete importing test items, please modify or add according to the actual situation,  do not run this procedure again or the data will be overwritten"
  Else
     MsgBox UText(27979, 35797, 39033, 30446, 23548, 20837, 23436, 25104) & "，" & UText(35831, 26681, 25454, 23454, 38469, 24773, 20917, 36827, 34892, 20462, 25913, 25110, 34917, 20805) & "," & UText(35831, 19981, 35201, 20877, 36816, 34892, 26412, 31243, 24207, 21542, 21017, 20462, 25913, 25110, 34917, 20805, 30340, 25968, 25454, 23558, 34987, 35206, 30422) & "!" & CHR(10) & _
     "Complete importing test items, please modify or add according to the actual situation,  do not run this procedure again or the data will be overwritten"
  End If

1000 End Sub
Sub GenerateBarCodeDataSheet()

Dim InputText As String
Set WD = ThisDocument

   Dim qtyRow As Integer
    With WD.Range.Tables(getTableNumber(1))
     qtyRow = .Rows.count
     Itms = qtyRow - 5
  
    ReDim Itx(1 To Itms)
    ReDim Sm(1 To Itms)
     For i = 5 To qtyRow - 1
       If Mid(.Cell(i, 1).Range.Text, 1, Len(.Cell(i, 1).Range.Text) - 2) <> "" Or Mid(.Cell(i, 2).Range.Text, 1, Len(.Cell(i, 2).Range.Text) - 2) <> "" Then
         
          Itx(i - 4) = Mid(.Cell(i, 2).Range.Text, 1, Len(.Cell(i, 2).Range.Text) - 2)
          Sm(i - 4) = Val(.Cell(i, 8).Range.Text)     ' translated comment
      End If
     Next i
   End With
   
 Dim ComItx() As String
 ReDim ComItx(1 To Itms)
 
 Dim ComSm() As Integer
 ReDim ComSm(1 To Itms)

 Dim ComItms As Integer
 Dim NI As Integer
   
      NI = 0
      Set dictItem = CreateObject("Scripting.Dictionary")
      For i = 1 To Itms
        If Not dictItem.Exists(Itx(i)) Then
            NI = NI + 1
            dictItem.add Itx(i), Sm(i)
             ComItx(NI) = Itx(i)
             ComSm(NI) = Sm(i)
        Else
            dictItem(Itx(i)) = dictItem(Itx(i)) + Sm(i)
            ComSm(NI) = dictItem(Itx(i))
        End If
    Next i
    ComItms = dictItem.count

 Dim ColorN() As Integer
 Dim SizeN() As Integer
 ReDim ColorN(1 To ComItms)
 ReDim SizeN(1 To ComItms)
 Dim Response As VbMsgBoxResult
 
 tbn = getTableNumber(8)
 Response = MsgBox("Dose each color and size of the sampe item of product has the same barcode?" & CHR(10) & _
 UText(26159, 21542, 21516, 27454, 20135, 21697, 20013, 19981, 21516, 39068, 33394, 21644, 23610, 30721, 30340, 26465, 27454, 37117, 30456, 21516) & "?", vbYesNo)

 
   If Response = vbNo Then
     For i = 1 To ComItms
        InputText = InputBox("Please enter the number of colors and sizes of " & ComItx(i) & "separated by commas." & CHR(10) & _
        UText(35831, 36755, 20837, 27454, 21495, 20026) & ComItx(i) & UText(20135, 21697, 30340, 39068, 33394, 25968, 21644, 23610, 23544, 25968) & "," & UText(24182, 29992, 36887, 21495, 38548, 24320) & CHR(10), "Please input the number of colors and sizes for the item" & " " & UText(35831, 36755, 20837, 26412, 27454, 20135, 21697, 30340, 23610, 23544, 25968, 21644, 23610, 30721, 25968), _
       "1,1")
       If InputText = "" Then Exit Sub
       ColorN(i) = Val(Mid(InputText, 1, InStr(1, InputText, ",") - 1))
       SizeN(i) = Val(Mid(InputText, InStr(1, InputText, ",") + 1, Len(InputText) - InStr(1, InputText, ",")))
     Next i
   ElseIf Response = vbYes Then
       For i = 1 To ComItms
           ColorN(i) = 1
           SizeN(i) = 1
       Next i
   Else
      Exit Sub
   End If
   For i = 1 To ComItms
     For j = 1 To ColorN(i)
       For k = 1 To SizeN(i)
        Call MakeBarCodeDataSheet(tbn, ComItms, ColorN, SizeN)
           With WD.Tables(tbn)
             .Cell(1, 1).Range.Text = "Data Sheet for" & " " & "Readability test for Barcode,QR code"
             .Cell(2, 1).Range.Text = "Item No.:" & ComItx(i)
           End With
       Next k
     Next j
   Next i
   
End Sub




Sub MakeNormalDataSheet(tbn, ComItms)
   Dim W1, W2, W3, W4 As Single
    
    W1 = 3.8    ' translated comment
    W2 = 4
    W3 = 2.2
    W4 = 7.8
      Set WD = ThisDocument
      WD.Tables(tbn).Cell(1, 1).Select
          Selection.MoveUp
          Selection.TypeParagraph
          ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=ComItms + 2, NumColumns:= _
             4, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
           With Selection.Tables(1)
               If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
               End If
               Call Setborders025
              .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
              .Columns(1).Width = CentimetersToPoints(W1)
              .Columns(2).Width = CentimetersToPoints(W2)
              .Columns(3).Width = CentimetersToPoints(W3)
              .Columns(4).Width = CentimetersToPoints(W4)
              .Rows(1).Range.Cells.Merge
              .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Rows(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Cell(2, 1).Range.Text = "Item No."
              .Cell(2, 2).Range.Text = "Specification"
              .Cell(2, 3).Range.Text = "Tolerance"
              .Cell(2, 4).Range.Text = "Result"
              .Rows(1).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 ' -603917569 '
              .Rows(2).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 ' -603917569 '
         End With
End Sub
Sub MakeBarCodeDataSheet(tbn, ComItms, ColorN, SizeN)
'BAOCDOE
   Dim W1, W2, W3, W4 As Single

   Dim RST As Integer
    W1 = 3.8    ' translated comment
    W2 = 5
    W3 = 9
    
      Set WD = ThisDocument
      WD.Tables(tbn).Cell(1, 1).Select
          Selection.MoveUp
          Selection.TypeParagraph
          ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=9, NumColumns:= _
             3, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
           With Selection.Tables(1)
               If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
               End If
               Call Setborders025
              .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Columns(1).Width = CentimetersToPoints(W1)
              .Columns(2).Width = CentimetersToPoints(W2)
              .Columns(3).Width = CentimetersToPoints(W3)
             
              .Rows(1).Range.Cells.Merge
              .Rows(2).Range.Cells.Merge
              .Rows(3).Range.Cells.Merge
              .Rows(4).Range.Cells.Merge
              '.Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
             ' .Rows(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Cell(2, 1).Range.Text = "Item No.:"
              .Cell(3, 1).Range.Text = "Color:"
              .Cell(4, 1).Range.Text = "Size:"
              .Cell(5, 1).Range.Text = "Position"
              .Cell(5, 2).Range.Text = "Specification"
              .Cell(5, 3).Range.Text = "Findings"
              .Rows(1).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 ' -603917569 '
              .Rows(5).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 ' -603917569 '
         End With
End Sub
Sub MakeCDFDataSheet(tbn, ComItms)
   Dim TblW, W1, W2, W3, W4, W5 As Single
    TblW = 17.8
    W1 = 1.31  ' translated comment
    W2 = 3.57
    W3 = 4.61
    W4 = 4.83
    W5 = 3.48
      
            Set WD = ThisDocument
      
            WD.Tables(tbn).Cell(1, 1).Select
                    Selection.MoveUp
                    Selection.TypeParagraph
                    ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=5, NumColumns:= _
                     5, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
                   
                    With Selection.Tables(1)
                      If .Style <> UText(32593, 26684, 22411) Then
                         .Style = UText(32593, 26684, 22411)
                      End If
                      Call Setborders025
                     .PreferredWidth = CentimetersToPoints(TblW)
                     .Columns(1).Width = CentimetersToPoints(W1)
                     .Columns(2).Width = CentimetersToPoints(W2)
                     .Columns(3).Width = CentimetersToPoints(W3)
                     .Columns(4).Width = CentimetersToPoints(W4)
                     .Columns(5).Width = CentimetersToPoints(W5)
                      .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                      .Rows(1).Range.Cells.Merge
                      .Cell(1, 1).Range.Text = "Data Sheet for Constructional Data Form (CDF)"
                      .Rows(1).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 ' -603917569 '
                      .Rows(2).Range.Cells.Merge
                      .Rows(2).Range.Cells.Split NumRows:=1, NumColumns:=3, mergebeforesplit:=True
                      '.Rows(2).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                      .Cell(2, 1).Width = CentimetersToPoints(4.87)
                      .Cell(2, 1).Range.Text = "Item No.:"
                      .Cell(2, 2).Width = CentimetersToPoints(7.13)
                      .Cell(2, 2).Range.Text = "Manufacture Model No.:"
                       .Cell(2, 3).Width = CentimetersToPoints(5.8)
                       .Cell(2, 3).Range.Text = "Report No.:"
                       .Rows(3).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 ' -603917569 '
                       .Cell(3, 1).Range.Text = "No."
                       .Cell(3, 2).Range.Text = "Component Name"
                       .Cell(3, 3).Range.Text = "On CDF"
                       .Cell(3, 4).Range.Text = "Findings"
                       .Cell(3, 5).Range.Text = "Result"
                       .Cell(4, 1).Range.Text = "1"
                       .Cell(5, 1).Range.Text = "2"
                       .Cell(4, 5).Select
                        Selection.Range.Font.Bold = True
                        Selection.Range.Font.Color = 194
                        Selection.Range.ContentControls.add (wdContentControlComboBox)
                        Selection.ParentContentControl.DropdownListEntries.Clear
                        Selection.ParentContentControl.DropdownListEntries.add Text:="Choose", Value:="Choose"
                        Selection.Font.Color = 194
                        Selection.ParentContentControl.DropdownListEntries.add Text:="Conformed", Value:="Conformed"
                        Selection.Font.Color = 194
                        Selection.ParentContentControl.DropdownListEntries.add Text:="Not Conformed", Value:="Not Conformed"
                        Selection.Font.Color = 194
                        Selection.ParentContentControl.DropdownListEntries.add Text:="Pending", Value:="Pending"
                        Selection.Font.Color = 194
                        Selection.ParentContentControl.DropdownListEntries(1).Select
                        'ActiveWindow.ActivePane.HorizontalPercentScrolled = 11
                        .Cell(5, 5).Select
                        Selection.Range.Font.Bold = True
                        Selection.Range.Font.Color = 194
                        Selection.Range.ContentControls.add (wdContentControlComboBox)
                        Selection.ParentContentControl.Tag = "rowresult"
                        Selection.ParentContentControl.DropdownListEntries.Clear
                        Selection.ParentContentControl.DropdownListEntries.add Text:="Choose", Value:="Choose"
                        Selection.Font.Color = 194
                        Selection.ParentContentControl.DropdownListEntries.add Text:="Conformed", Value:="Conformed"
                        Selection.Font.Color = 194
                        Selection.ParentContentControl.DropdownListEntries.add Text:="Not Conformed", Value:="Not Conformed"
                        Selection.Font.Color = 194
                        Selection.ParentContentControl.DropdownListEntries.add Text:="Pending", Value:="Pending"
                        Selection.Font.Color = 194
                        Selection.ParentContentControl.DropdownListEntries(1).Select
                        ActiveWindow.ActivePane.HorizontalPercentScrolled = 11
                   
                   End With
End Sub
Sub MakeWithPicDataSheet(tbn, ComItms)
   Dim TblW, W1, W2, W3, W4 As Single
   Dim Rn As Integer
    
    

    TblW = 17.8 ' translated comment
    W1 = 3.8
    W2 = 4
    W3 = 2.2
    W4 = 7.8
    Rn = 7
      Set WD = ThisDocument
    
      WD.Tables(tbn).Cell(1, 1).Select
          Selection.MoveUp
          Selection.TypeParagraph
          ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=Rn, NumColumns:= _
             4, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
           With Selection.Tables(1)
               If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
               End If
               Call Setborders025
              .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Columns(1).Width = CentimetersToPoints(W1)
              .Columns(2).Width = CentimetersToPoints(W2)
              .Columns(3).Width = CentimetersToPoints(W3)
              .Columns(4).Width = CentimetersToPoints(W4)
              .Rows(1).Range.Cells.Merge
             ' .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Rows(1).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 ' -603917569 '
              i = 2
                .Rows(i).Range.Cells.Merge
              '  .Rows(i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                .Rows(i + 1).Range.Cells.Merge
              '  .Rows(i + 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Rows(i + 2).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 ' -603917569
                .Rows(i + 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Cell(i + 2, 1).Range.Text = "Check point"
                .Cell(i + 2, 2).Range.Text = "Specification"
                .Cell(i + 2, 3).Range.Text = "Tolerance"
                .Cell(i + 2, 4).Range.Text = "Result"
                .Cell(i + 1, 1).Range.Text = "Only for photo of measurement drawings, if not any photo please delete this row"
                .Cell(i + 1, 1).Range.Font.ColorIndex = wdGray25
         End With
     
End Sub
Sub MakeWithPicDataSheet_Combinded(tbn, ComItms)
   Dim TblW, W1, W2, W3, W4 As Single
   Dim Rn As Integer
    
    

    TblW = 17.8 ' translated comment
    W1 = 3.8
    W2 = 4
    W3 = 2.2
    W4 = 7.8
    Rn = 1 + ComItms + 5 * ComItms
      Set WD = ThisDocument
      WD.Tables(tbn).Cell(1, 1).Select
          Selection.MoveUp
          Selection.TypeParagraph
          ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=Rn, NumColumns:= _
             4, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
           With Selection.Tables(1)
               If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
               End If
               Call Setborders025
              .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Columns(1).Width = CentimetersToPoints(W1)
              .Columns(2).Width = CentimetersToPoints(W2)
              .Columns(3).Width = CentimetersToPoints(W3)
              .Columns(4).Width = CentimetersToPoints(W4)
              .Rows(1).Range.Cells.Merge
            '  .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Rows(1).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 ' -603917569 '
              For i = 2 To Rn Step 6
                .Rows(i).Range.Cells.Merge
              ' .Rows(i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                .Rows(i + 1).Range.Cells.Merge
               ' .Rows(i + 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Rows(i + 2).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 ' -603917569
                .Rows(i + 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Cell(i + 2, 1).Range.Text = "Check point"
                .Cell(i + 2, 2).Range.Text = "Specification"
                .Cell(i + 2, 3).Range.Text = "Tolerance"
                .Cell(i + 2, 4).Range.Text = "Result"
                .Cell(i + 1, 1).Range.Text = "Only for photo of measurement drawings, if not any photo please delete this row"
                .Cell(i + 1, 1).Range.Font.ColorIndex = wdGray25
              Next i
           
         End With
End Sub
Sub MakeShoesDataSheet(tbn, ComItms)
   Dim TblW, TtlW, ClW, DatW, TolW As Single
   Dim Rn, CLn As Integer
   Dim Rng As Range


    
    TblW = 17.8 ' translated comment
    TtlW = 3.6
    ClW = 0.8
    DatW = 1.2
    TolW = 1
    
    Rn = 2 + 4 * 3
    CLn = 1 + 4 * 3 + 1
      Set WD = ThisDocument
      WD.Tables(tbn).Cell(1, 1).Select
          Selection.MoveUp
          Selection.TypeParagraph
          ActiveDocument.Tables.add Range:=Selection.Range, NumRows:=Rn, NumColumns:= _
              CLn, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
           With Selection.Tables(1)
               If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
               End If
               Call Setborders025
              .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Columns(1).Width = CentimetersToPoints(TtlW)
              For i = 2 To 10 Step 4
                 .Columns(i).Width = CentimetersToPoints(ClW)
                 .Columns(i + 1).Width = CentimetersToPoints(DatW)
                 .Columns(i + 2).Width = CentimetersToPoints(DatW)
                 .Columns(i + 3).Width = CentimetersToPoints(DatW)
              Next i
                 .Columns(CLn).Width = CentimetersToPoints(TolW)
                ' .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                 .Rows(1).Range.Cells.Shading.BackgroundPatternColor = -603917569 '-603914241 ' -603917569 '
              '   .Rows(2).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
              For i = 11 To 3 Step -4
                For j = 10 To 2 Step -4
                   .Cell(i, j).Range.Text = "SZ"
                   .Cell(i + 1, j).Range.Text = "CL"
                Next j
              Next i
               For i = 3 To 11 Step 4
                .Rows(i).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Cell(i, 1).Range.Text = "Measurement Point" 'Meas.
                .Cell(i + 2, 1).Range.Text = "Insole length (mm)"
                .Cell(i + 3, 1).Range.Text = "Lace Length (cm)"
                .Cell(i, CLn).Range.Text = "TOL" 'Meas.
              
              Next i
              For i = 11 To 3 Step -4
                   Set Rng = .Cell(i, CLn).Range
                   Rng.End = .Cell(i + 1, CLn).Range.End
                   Rng.Cells.Merge
                   'Rng.ParagraphFormat.Alignment = wdAlignVerticalBottom
              Next i
              For i = 11 To 3 Step -4
                For j = 11 To 3 Step -4
                   Set Rng = .Cell(i, j).Range
                   Rng.End = .Cell(i, j + 2).Range.End
                   Rng.Cells.Merge
                Next j
              Next i
              
              For i = 11 To 3 Step -4
                   Set Rng = WD.Range(Start:=.Cell(i, 1).Range.Start, End:=.Cell(i + 1, 1).Range.End)
                   Rng.Select
                   Selection.Cells.Merge
               
              Next i
                   Set Rng = .Cell(1, 1).Range
                   Rng.End = .Cell(i, CLn).Range.End
                   Rng.Cells.Merge
                   Set Rng = .Cell(2, 1).Range
                   Rng.End = .Cell(2, CLn).Range.End
                   Rng.Cells.Merge

              
           
         End With
End Sub

Sub Input_PrimaryTypeCheckPoint(Cn As Integer, N As Integer, AddchkCode() As Variant)
'by wyn 20241216

Dim Ino(), Num(), Typ(), Chk(), Cri(), Spe(), sms(), Tol(), Nat() As String
ReDim Ino(1 To N)
ReDim Num(1 To N)
ReDim Typ(1 To N)
ReDim Chk(1 To N)
ReDim Cri(1 To N)
ReDim Spe(1 To N)
ReDim sms(1 To N)
ReDim Tol(1 To N)
ReDim Nat(1 To N)

Dim Cha(200) As String

Dim NumO(), NumP(), NumW(), NumD(), NumL(), NumS(), NumC(), NumM() As String
ReDim NumO(1 To N)
ReDim NumP(1 To N)
ReDim NumW(1 To N)
ReDim NumD(1 To N)
ReDim NumL(1 To N)
ReDim NumS(1 To N)
ReDim NumC(1 To N)
ReDim NumM(1 To N)

Dim ChkO(), ChkP(), ChkW(), ChkD(), ChkL(), ChkS(), ChkC(), ChkM() As String
ReDim ChkO(1 To N)
ReDim ChkP(1 To N)
ReDim ChkW(1 To N)
ReDim ChkD(1 To N)
ReDim ChkL(1 To N)
ReDim ChkS(1 To N)
ReDim ChkC(1 To N)
ReDim ChkM(1 To N)

Dim TableChoise As String

TableChoise = GetPropertyValue("TableChoise")
If TableChoise = "" Then

  TableChoise = InputBox("The program will generate tabels for datasheet, please input number for the table type you would like" & CHR(10) & _
  "1:  normal datasheet, one check point with one table, one row for one item of product." & CHR(10) & _
  "2:  complex datasheet, some cell units for one item of product,each row for one sub check point, and support pasting in drawings photo." & CHR(10) & _
  "3:  Both 1 and 2 mentioned above" & CHR(10) & CHR(10) & _
  UText(31243, 24207, 23558, 29983, 25104, 25968, 25454, 34920, 26684) & "，" & UText(35831, 36755, 20837, 25968, 23383, 36873, 25321, 29983, 25104, 30340, 25968, 25454, 34920, 26684, 31867, 22411) & CHR(10) & _
  "1: " & UText(36890, 29992, 25968, 25454, 34920) & "，" & UText(27599, 20010, 26816, 26597, 39033, 30446, 19968, 20010, 34920, 26684) & "，" & UText(27599, 27454, 20135, 21697, 25968, 25454, 21344, 29992, 19968, 34892) & CHR(10) & _
  "2: " & UText(22797, 21512, 25968, 25454, 34920, 26684) & "，" & UText(27599, 27454, 20135, 21697, 21344, 29992, 22810, 20010, 21333, 20803, 26684, 32452, 21512) & "，" & UText(27599, 34892, 21487, 29992, 20110, 36755, 20837, 23376, 26816, 26597, 39033, 30446) & "，" & UText(25903, 25345, 31896, 36148, 20837, 27979, 37327, 22270, 32440, 29031, 29255) & CHR(10) & _
  "3: " & UText(21516, 26102, 29983, 25104, 19978, 36848) & "1，2" & UText(34920, 26684) & CHR(10), "Please Choose datasheet type" & "  " & UText(35831, 36873, 25321, 34920, 26684, 31867, 22411), "1")
End If

Dim i, j, k, X, tb, tbn, TbnB, TbnE, cnt, tNum, Z As Integer
Dim W, o, PP, L, D, S, C, M As Integer
Dim cno, cnp, cnw, cnd, cnl, cns, cnc, cnm, Ps, Psn, NN, NN1 As Integer
Dim msg, DelTab, Kw1, Kw2 As String

Dim a, TCHAR
Dim uC(200) As String
Dim CHAR(100, 50) As String
Dim KwChk As String
Dim Sm() As Integer

Set WD = ThisDocument
Dim qtyRow As Integer
    
5   With ThisDocument.Range.Tables(getTableNumber(1))
     qtyRow = .Rows.count
     Itms = qtyRow - 5
    ReDim Ptx(1 To Itms)
    ReDim Itx(1 To Itms)
    ReDim Sm(1 To Itms)
     For i = 5 To qtyRow - 1
       If Mid(.Cell(i, 1).Range.Text, 1, Len(.Cell(i, 1).Range.Text) - 2) <> "" Or Mid(.Cell(i, 2).Range.Text, 1, Len(.Cell(i, 2).Range.Text) - 2) <> "" Then
          
          Ptx(i - 4) = Mid(.Cell(i, 1).Range.Text, 1, Len(.Cell(i, 1).Range.Text) - 2)
          Itx(i - 4) = Mid(.Cell(i, 2).Range.Text, 1, Len(.Cell(i, 2).Range.Text) - 2)
          Sm(i - 4) = Val(.Cell(i, 8).Range.Text)     ' translated comment
      End If
     Next i
   End With
   
 Dim ComItx() As String
 ReDim ComItx(1 To Itms)
 
 Dim ComSm() As Integer
 ReDim ComSm(1 To Itms)

 Dim ComItms As Integer
 Dim NI As Integer
   
      NI = 0
      Set dictItem = CreateObject("Scripting.Dictionary")
      For i = 1 To Itms
        If Not dictItem.Exists(Itx(i)) Then
            NI = NI + 1
            dictItem.add Itx(i), Sm(i)
             ComSm(NI) = Sm(i)
             ComItx(NI) = Itx(i)
        Else
            dictItem(Itx(i)) = dictItem(Itx(i)) + Sm(i)
            ComSm(NI) = dictItem(Itx(i))
        End If
    Next i
    ComItms = dictItem.count
    
     'Itms = Val(GetPropertyValue("ItemsQty"))
    
     StoretoProperty "ItemsQty", CStr(Itms)
   
   '-----------
 
 '（0）-（1）-（2)-（3）-（4）-（5）-（6）-（7）-（8）


Dim SelPriChkNo() As Integer
ReDim SelPriChkNo(1 To N)
 
 For i = 1 To N
   For j = 1 To Cn
     If AddchkCode(i) = Mid(PriChk(j), 1, InStr(PriChk(j), "|") - 1) Then
        SelPriChkNo(i) = j
     
        GoTo 8
     End If
   Next j
8 Next i

For i = 1 To N
    PriChk(SelPriChkNo(i)) = Replace(PriChk(SelPriChkNo(i)), "，", ",")

    a = Split(PriChk(SelPriChkNo(i)), "|")
    u = UBound(a)

    Num(i) = i
    Ino(i) = a(0)
    Chk(i) = a(1)
    Typ(i) = Trim(a(2))
    Cri(i) = Trim(a(3))
    sms(i) = Trim(a(4))
    
    NN1 = Val(sms(i))
    Psn = InStr(1, sms(i), "≥")
    If Psn > 0 Then
     NN = Val(Mid(sms(i), Psn + 1, Len(sms(i)) - Psn))
      If ComItms = 1 Then
        sms(i) = CStr(NN)
      Else
         sms(i) = Mid(sms(i), 1, Psn - 1)
         Do While NN1 * ComItms < NN
            sms(i) = CStr((NN1 + 1)) & Mid(sms(i), 2, Len(sms(i)) - 1)
            NN1 = Val(sms(i))
         Loop
      End If
    Else
       If ComItms = 1 Then sms(i) = Replace(sms(i), "/item", "")
    End If
    
    Spe(i) = Trim(a(5))
    Tol(i) = Trim(a(6))
    Nat(i) = Trim(a(7))
    
      Cha(i) = Trim(a(8))      'char（）
        TCHAR = Split(Cha(i), ",")
        uC(i) = UBound(TCHAR)
       For j = 0 To uC(i)
         CHAR(i, j) = TCHAR(j)
       Next j
   Next i
  
          W = 0
     o = 0
     D = 0
     S = 0
     C = 0
     L = 0
     M = 0
     PP = 0

     
        j = 0
        msg = ""
 For X = 1 To Cn
   If ThisDocument.Bookmarks.Exists(AddchkCode(X)) = False Then  ' translated comment
    For i = 1 To N
     If AddchkCode(X) = Ino(i) Then
            Select Case LCase(Typ(i))
                '-------------------
            Case UText(25163, 24037, 26816, 26597)
             'Case "workmanship"
                    
                            W = W + 1
                            NumW(W) = Num(i)
                            ChkW(W) = Chk(i)
                            GoTo 10
                
               '-------------------
           Case UText(29305, 27530, 27979, 35797)
         ' Case "on site test"
         
                            o = o + 1
                            NumO(o) = Num(i)
                            ChkO(o) = Chk(i)
                            GoTo 10
                 
            '------------------
          Case UText(20135, 21697, 27979, 37327), UText(37325, 37327, 26816, 26597)
        ' Case "product dimension and weight"
        
                            D = D + 1
                           NumD(D) = Num(i)
                           ChkD(D) = Chk(i)
                            GoTo 10
             '===================
            Case UText(27454, 24335, 26816, 26597), UText(25991, 20214, 26816, 26597)
                           S = S + 1
                           NumS(S) = Num(i)
                           ChkS(S) = Chk(i)
                            GoTo 10
    
              '----------
              Case UText(39068, 33394, 26816, 26597)
    
                           C = C + 1
                           NumC(C) = Num(i)
                           ChkC(C) = Chk(i)
                           GoTo 10
     
               '===
             Case UText(26631, 31614, 26631, 35782, 26816, 26597)
             ' Case "labeling check"
                             L = L + 1
                             NumL(L) = Num(i)
                            ChkL(L) = Chk(i)
                            GoTo 10
                            
                '==================
              Case UText(21787, 22836, 26816, 26597)
                            M = M + 1
                           NumM(M) = Num(i)
                           ChkM(M) = Chk(i)
                           GoTo 10
       
                 '--------------------------------------
               Case UText(21253, 35013, 26816, 26597)
              'Case "packing and packaging"
                
                             PP = PP + 1
                             NumP(PP) = Num(i)
                              ChkP(PP) = Chk(i)
                              GoTo 10
          End Select
      End If
10   Next i
   Else
        For i = 1 To N
          If AddchkCode(X) = Ino(i) Then
            j = j + 1
             msg = msg & j & "." & Chk(i) & "," & CHR(10)
          End If
        Next i
   End If
 Next X
    If msg <> "" Then
       MsgBox "The report exist the checkitems of" & CHR(10) & msg & "to avoid deplicating, these checkitems will not be imported" & CHR(10) _
       & CHR(10) & UText(25253, 21578, 20013, 24050, 21253, 21547, 26377, 19979, 21015, 26816, 26597, 39033, 30446) & CHR(10) & msg & UText(20026, 36991, 20813, 37325, 22797) & "，" & UText(23558, 19981, 20877, 23548, 20837, 36825, 20123, 39033, 30446)
    End If
    msg = ""
    
       cnw = W
       cno = o
       cnd = D
       cns = S
       cnc = C
       cnl = L
       cnm = M
       cnp = PP


 '------------------------------------------workmanship
20  If cnw > 0 Then
        tbn = getTableNumber(2)
        If tbn = 0 Then
           msg = msg & UText(26410, 21457, 29616) & "2. Workmanship" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
          "2. Workmanship table is not found, the check item for workmanship is not imported" & CHR(10)
          GoTo 40
        End If

     
30     With WD.Tables(tbn)
          Rn = .Rows.count
          .Cell(Rn, 1).Select
          Selection.InsertRowsBelow cnw
          For i = 1 To cnw
            .Rows(Rn + i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cell(Rn + i, 1).Range.Text = "2." & Rn + i - 6
             If Len(Chk(NumW(i))) > 1 Then
               KwChk = UCase(Mid(Chk(NumW(i)), 1, 1)) & Mid(Chk(NumW(i)), 2, Len(Chk(NumW(i))) - 1)
             Else
                 KwChk = Chk(NumW(i))
             End If
             
             If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
          
             .Cell(Rn + i, 2).Range.Text = KwChk
             Set Cel = .Cell(Rn + i, 2)
             Call DeleteBookMark(Cel)
             Cel.Range.Bookmarks.add Name:=Ino(NumW(i)), Range:=Cel.Range
             If Len(Cri(NumW(i))) >= 2 Then
                 .Cell(Rn + i, 3).Range.Text = UCase(Mid(Cri(NumW(i)), 1, 1)) & Mid(Cri(NumW(i)), 2, Len(Cri(NumW(i))) - 1)
             Else
                 .Cell(Rn + i, 3).Range.Text = UCase(Mid(Cri(NumW(i)), 1, 1))
             End If
           Next i
       End With
  
     '---workmanship
      TbnB = tbn
       DelTab = ""
       
 
35     For i = 1 To cnw
        If Nat(NumW(i)) = UText(23450, 37327) Then
          tbn = 0
          tbn = getTableNumber(3)
          If tbn = 0 Then
             msg = msg & UText(26410, 21457, 29616) & "3. On Site Test" & UText(30340, 34920, 26684) & "，Workmanship " & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
            "3. On Site Test table is not found, Data sheet for Workmanship form can not generate" & CHR(10)
             GoTo 40
          End If

37        Call MakeNormalDataSheet(tbn, ComItms)
              
              With WD.Tables(tbn)
              .Cell(1, 1).Range.Text = "Data Sheet for " & Chk(NumW(i))
               Set Cel = .Cell(1, 1)
               Call DeleteBookMark(Cel)
               Cel.Range.Bookmarks.add Name:=Ino(NumW(i)) & "_Data", Range:=Cel.Range
             
               For j = 1 To ComItms
                .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                .Cell(j + 2, 1).Range.Text = ComItx(j)
                .Cell(j + 2, 2).Range.Text = Spe(NumW(i))
                .Cell(j + 2, 3).Range.Text = Tol(NumW(i))
               Next j
             End With
          End If
38     Next i
    End If
  
40  ''------------------------------------------onsitetest
   If cno > 0 Then
       tbn = 0
       tbn = getTableNumber(3)
       If tbn = 0 Then
          msg = msg & UText(26410, 21457, 29616) & "3. On Site Test" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
         "3. On Site Test table is not found, the check item for On Site Test is not imported" & CHR(10)
          GoTo 60
       End If

50     With WD.Tables(tbn)
          Rn = .Rows.count
          Coln = .Columns.count
          .Cell(Rn, 1).Select
          Selection.InsertRowsBelow cno
          For i = 1 To cno
              .Rows(Rn + i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
              .Cell(Rn + i, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
               .Cell(Rn + i, 1).Range.Text = "3." & Rn + i - 3
               If Len(Chk(NumO(i))) > 1 Then
                 KwChk = UCase(Mid(Chk(NumO(i)), 1, 1)) & Mid(Chk(NumO(i)), 2, Len(Chk(NumO(i))) - 1)
               Else
                  KwChk = Chk(NumO(i))
               End If
               If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
               .Cell(Rn + i, 2).Range.Text = KwChk
                Set Cel = .Cell(Rn + i, 2)
                Call DeleteBookMark(Cel)
                Cel.Range.Bookmarks.add Name:=Ino(NumO(i)), Range:=Cel.Range
                If Len(Cri(NumO(i))) > 1 Then
                 .Cell(Rn + i, 3).Range.Text = UCase(Mid(Cri(NumO(i)), 1, 1)) & Mid(Cri(NumO(i)), 2, Len(Cri(NumO(i))) - 1)
                Else
                  .Cell(Rn + i, 3).Range.Text = Cri(NumO(i))
                End If
              '  .cell(Rn + i, 4).Range.Text = UCase(Mid(sms(NumO(i)), 1, 1)) & Mid(sms(NumO(i)), 2, Len(sms(NumO(i))) - 1)
                 
                 If sms(NumO(i)) <> "" Then
                 
                     If ComItms = 1 Then
                         sms(NumO(i)) = Replace(sms(NumO(i)), " /item", "")
                         sms(NumO(i)) = Replace(sms(NumO(i)), " /Item", "")
                         .Cell(Rn + i, 4).Range.Text = sms(NumO(i))
                     Else
                         If Len(sms(NumO(i))) > 1 Then
                           .Cell(Rn + i, 4).Range.Text = UCase(Mid(sms(NumO(i)), 1, 1)) & Mid(sms(NumO(i)), 2, Len(sms(NumO(i))) - 1)
                         Else
                           .Cell(Rn + i, 4).Range.Text = sms(NumO(i))
                         End If
                    End If
             Else
                 Set Cel = .Cell(Rn + i, 4)
                 Cel.Select
                 Call CellAddDropDownList(Cel)

             End If
   
                    
                 If .Cell(Rn, Coln).Range.ContentControls.count > 0 Then
                  .Cell(Rn, Coln).Range.Select
                   Selection.Copy
                  .Cell(Rn + i, Coln).Range.Select
                   StartTime = Timer '
                   Do While Timer - StartTime < 0.1
                     DoEvents
                   Loop
                   Selection.Paste
                  Selection.ContentControls(1).DropdownListEntries(1).Select
                  End If
           Next i
     
       End With
       '---On site test
       TbnB = tbn
       DelTab = ""
       
       For i = 1 To cno
        If Nat(NumO(i)) = UText(23450, 37327) Then
55        tbn = 0
          tbn = getTableNumber(4)
          If tbn = 0 Then
             msg = msg & UText(26410, 21457, 29616) & "4. Product Dimension & Weight" & UText(30340, 34920, 26684) & "，On Site Test " & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
             "4. Product Dimension & Weight table is not found, Data sheet for On Site Test form can not generate" & CHR(10)
             GoTo 60
          End If

57         Call MakeNormalDataSheet(tbn, ComItms)
              
              With WD.Tables(tbn)
              
              .Cell(1, 1).Range.Text = "Data Sheet for " & Chk(NumO(i))
               Set Cel = .Cell(1, 1)
               Call DeleteBookMark(Cel)
               Cel.Range.Bookmarks.add Name:=Ino(NumO(i)) & "_Data", Range:=Cel.Range
             
               For j = 1 To ComItms
                .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                .Cell(j + 2, 1).Range.Text = ComItx(j)
                .Cell(j + 2, 2).Range.Text = Spe(NumO(i))
                .Cell(j + 2, 3).Range.Text = Tol(NumO(i))
               Next j
            End With
         End If
58    Next i
    End If
60 '---------------------------------------------------4. Product Dimension & Weight
    If cnd > 0 Then
       tbn = 0
       tbn = getTableNumber(4)
       If tbn = 0 Then
          msg = msg & UText(26410, 21457, 29616) & "4. Product Dimension & Weight" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
          "4. Product Dimension & Weight table is not found, the check item for Product Dimension & Weight is not imported" & CHR(10)
          GoTo 80
       End If

70      With WD.Tables(tbn)
         Rn = .Rows.count
         Coln = .Columns.count
         .Cell(Rn, 1).Select
         Selection.InsertRowsBelow cnd
         For i = 1 To cnd
              .Rows(Rn + i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
              .Cell(Rn + i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Cell(Rn + i, 1).Range.Text = "4." & Rn + i - 3
               If Len(Chk(NumD(i))) > 1 Then
                  KwChk = UCase(Mid(Chk(NumD(i)), 1, 1)) & Mid(Chk(NumD(i)), 2, Len(Chk(NumD(i))) - 1)
               Else
                  KwChk = Chk(NumD(i))
               End If
               If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
               .Cell(Rn + i, 2).Range.Text = KwChk
               Set Cel = .Cell(Rn + i, 2)
               Call DeleteBookMark(Cel)
               Cel.Range.Bookmarks.add Name:=Ino(NumD(i)), Range:=Cel.Range
              ' .cell(Rn + i, 3).Range.Text = UCase(Mid(Cri(NumD(i)), 1, 1)) & Mid(Cri(NumD(i)), 2, Len(Cri(NumD(i))) - 1)
            If sms(NumD(i)) <> "" Then
               If ComItms = 1 Then
                 sms(NumD(i)) = Replace(sms(NumD(i)), " /item", "")
                 sms(NumD(i)) = Replace(sms(NumD(i)), " /Item", "")
                 .Cell(Rn + i, 3).Range.Text = sms(NumD(i))
              Else
                 If Len(sms(NumD(i))) > 1 Then
                     .Cell(Rn + i, 3).Range.Text = UCase(Mid(sms(NumD(i)), 1, 1)) & Mid(sms(NumD(i)), 2, Len(sms(NumD(i))) - 1)
                Else
                     .Cell(Rn + i, 3).Range.Text = sms(NumD(i))
                 End If
              End If
           Else
              Set Cel = .Cell(Rn + i, 3)
              Cel.Select
              Call CellAddDropDownList(Cel)

           End If

           If .Cell(Rn, Coln).Range.ContentControls.count > 0 Then
                 .Cell(Rn, Coln).Range.Select
                  Selection.Copy
                 .Cell(Rn + i, Coln).Range.Select
                  StartTime = Timer '
                  Do While Timer - StartTime < 0.1
                    DoEvents
                  Loop
                  Selection.Paste
                  Selection.ContentControls(1).DropdownListEntries(1).Select
           End If
        Next i
      End With
     
        'Product Dimension & Weight---
      TbnB = tbn
      DelTab = ""
     
      Dim WithShoe As String
      WithShoe = ""
      For i = 1 To cnd
        If Nat(NumD(i)) = UText(23450, 37327) Then
          DelTab = DelTab & "," & "Data Sheet for " & Chk(NumD(i))  ' translated comment
          tb = WD.Tables.count - 1
          Select Case TableChoise
          Case 1
             TbCount1 = 0
             If Ino(NumD(i)) = "TSG056" Then   ' translated comment
                WithShoe = DelTab
               For j = 3 To tb - 3
                With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumD(i)) & "*" = True Then
                       TbCount1 = TbCount1 + 1
                  End If
                End With
               Next j
               GoTo 75
            Else
             '==  type1
               TbCount1 = 0
               For j = 3 To tb - 3
                 With WD.Tables(j)
                    If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumD(i)) & "*" = True Then
                       If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                          TbCount1 = TbCount1 + 1
                          GoTo 78
                       End If
                    End If
                 End With
             Next j
           End If
         Case 2
             TbCount2 = 0
             For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumD(i)) & "*" = True Then
                       If .Rows(2).Cells.count = 1 Then
                          TbCount2 = TbCount2 + 1
                       End If
                    End If
                 End With
            Next j
       Case 3
            TbCount1 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
                 With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumD(i)) & "*" = True Then
                     If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                       TbCount1 = TbCount1 + 1
                     End If
                   End If
                End With
            Next j
            
            TbCount2 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumD(i)) & "*" = True Then
                      If .Rows(2).Cells.count = 1 Then
                         TbCount2 = TbCount2 + 1
                     End If
                   End If
                 End With
            Next j
      End Select
75    tbn = 0
      tbn = getTableNumber(5)
      If tbn = 0 Then
           msg = msg & UText(26410, 21457, 29616) & "5. Style / Material / Construction" & UText(34920, 26684) & "，Product Dimension & Weight " & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
           "5. Style / Material / Construction table is not found, Data sheet for Product Dimension & Weight form can not generate" & CHR(10)
           GoTo 80
      End If
77    Select Case TableChoise
         Case "1"
           If Ino(NumD(i)) = "TSG056" Then   ' translated comment
             For j = 1 To ComItms
                tbn = getTableNumber(5)
                Call MakeShoesDataSheet(tbn, ComItms)
                With WD.Tables(tbn)
                     Set Cel = .Cell(1, 1)
                     Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                     Call DeleteBookMark(Cel)
                     Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data" & "_" & TbCount1 + j, Range:=Cel.Range
                     Set Cel = .Cell(2, 1)
                     Cel.Range.Text = "Item No.:" & ComItx(TbCount1 + j)
                End With
             Next j
           Else
              tbn = getTableNumber(5)
              Call MakeNormalDataSheet(tbn, ComItms)
              With WD.Tables(tbn)
                  Set Cel = .Cell(1, 1)
                  Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                  Call DeleteBookMark(Cel)
                  Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data", Range:=Cel.Range
                  For j = 1 To ComItms
                    .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                    .Cell(j + 2, 1).Range.Text = ComItx(j)
                    .Cell(j + 2, 2).Range.Text = Spe(NumD(i))
                    .Cell(j + 2, 3).Range.Text = Tol(NumD(i))
                  Next j
              End With
           End If
        Case "2"
            For j = 1 To ComItms
               tbn = getTableNumber(5)
               Call MakeWithPicDataSheet(tbn, ComItms)
               With WD.Tables(tbn)
                     Set Cel = .Cell(1, 1)
                     Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                     Call DeleteBookMark(Cel)
                     Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                     Set Cel = .Cell(2, 1)
                     Cel.Range.Text = "Item No.:" & ComItx(j + TbCount2)
               End With
            Next j
         Case "3"
            tbn = getTableNumber(5)
            Call MakeNormalDataSheet(tbn, ComItms)
              With WD.Tables(tbn)
                   Set Cel = .Cell(1, 1)
                   Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                   Call DeleteBookMark(Cel)
                  Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data", Range:=Cel.Range
                  For j = 1 To ComItms
                     .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                     .Cell(j + 2, 1).Range.Text = ComItx(j)
                     .Cell(j + 2, 2).Range.Text = Spe(NumD(i))
                    .Cell(j + 2, 3).Range.Text = Tol(NumD(i))
                  Next j
             End With
             For j = 1 To ComItms
                  tbn = getTableNumber(5)
                  Call MakeWithPicDataSheet(tbn, ComItms)
                     With WD.Tables(tbn)
                        Set Cel = .Cell(1, 1)
                        Cel.Range.Text = "Data Sheet for " & Chk(NumD(i))
                        Call DeleteBookMark(Cel)
                        Cel.Range.Bookmarks.add Name:=Ino(NumD(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                        Set Cel = .Cell(2, 1)
                        Cel.Range.Text = "Item No.:" & ComItx(j + TbCount2)
                    End With
             Next j
        End Select
    End If
78  Next i
    
  End If
80 '---------------------------------------------------5. Style / Material / Construction
    If cns > 0 Then
            tbn = 0
     tbn = getTableNumber(5)
      If tbn = 0 Then
          msg = msg & UText(26410, 21457, 29616) & "5. Style / Material / Construction" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
         "5. Style / Material / Construction table is not found, the check item for Style / Material / Construction is not imported" & CHR(10)
          GoTo 110
     End If

90      With WD.Tables(tbn)
          Rn = .Rows.count
          Coln = .Columns.count
          .Cell(Rn, 1).Select
          Selection.InsertRowsBelow cns
          For i = 1 To cns
              .Rows(Rn + i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
              .Cell(Rn + i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Cell(Rn + i, 1).Range.Text = "5." & Rn + i - 3
               If Len(Chk(NumS(i))) > 1 Then
                  KwChk = UCase(Mid(Chk(NumS(i)), 1, 1)) & Mid(Chk(NumS(i)), 2, Len(Chk(NumS(i))) - 1)
               Else
                  KwChk = Chk(NumS(i))
               End If
               If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
               .Cell(Rn + i, 2).Range.Text = KwChk
               Set Cel = .Cell(Rn + i, 2)
               Call DeleteBookMark(Cel)
               Cel.Range.Bookmarks.add Name:=Ino(NumS(i)), Range:=Cel.Range
              ' .cell(Rn + i, 3).Range.Text = UCase(Mid(Cri(NumD(i)), 1, 1)) & Mid(Cri(NumD(i)), 2, Len(Cri(NumD(i))) - 1)
            If sms(NumS(i)) <> "" Then
               If ComItms = 1 Then
                 sms(NumS(i)) = Replace(sms(NumS(i)), " /item", "")
                 sms(NumS(i)) = Replace(sms(NumS(i)), " /Item", "")
                 .Cell(Rn + i, 3).Range.Text = sms(NumS(i))
              Else
                 If Len(sms(NumS(i))) > 1 Then
                     .Cell(Rn + i, 3).Range.Text = UCase(Mid(sms(NumS(i)), 1, 1)) & Mid(sms(NumS(i)), 2, Len(sms(NumS(i))) - 1)
                Else
                     .Cell(Rn + i, 3).Range.Text = sms(NumS(i))
                 End If
              End If
           Else
              Set Cel = .Cell(Rn + i, 3)
              Cel.Select
               Call CellAddDropDownList(Cel)
           End If

            If .Cell(Rn, Coln).Range.ContentControls.count > 0 Then
                 .Cell(Rn, Coln).Range.Select
                  Selection.Copy
                 .Cell(Rn + i, Coln).Range.Select
                  StartTime = Timer '
                  Do While Timer - StartTime < 0.1
                    DoEvents
                  Loop
                  Selection.Paste
                  Selection.ContentControls(1).DropdownListEntries(1).Select
            End If
        Next i
      End With
     
        '5. Style / Material / Construction---
     TbnB = tbn
     DelTab = ""
    
   
    For i = 1 To cns
      If Nat(NumS(i)) = UText(23450, 37327) Then
        'DelTab = DelTab & "," & "Data Sheet for " & Chk(NumS(i)) '
         If Chk(NumS(i)) Like "*Constructional Data Form*" = False Then
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
              With WD.Tables(j)
                 If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumS(i)) & "*" = True Then
                      Set Cel = .Cell(1, 1)
                      Call DeleteBookMark(Cel)
                      Cel.Range.Bookmarks.add Name:=Ino(NumS(i)) & "_Data", Range:=Cel.Range
                     Rn = .Rows.count
                     If Rn < ComItms + 2 Then
                        .Cell(Rn, 1).Select
                        Selection.InsertRowsBelow ComItms + 2 - Rn
                     ElseIf Rn > ComItms + 2 Then
                        For k = Rn To ComItms + 2 + 1 Step -1
                         .Rows(k).Delete
                       Next k
                    End If
                    For k = 1 To ComItms
                     .Rows(k + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                     .Cell(k + 2, 1).Width = CentimetersToPoints(W1)
                     .Cell(k + 2, 1).Range.Text = ComItx(k)
                     .Cell(k + 2, 2).Width = CentimetersToPoints(W2)
                     .Cell(k + 2, 2).Range.Text = Spe(NumS(i))
                     .Cell(k + 2, 3).Width = CentimetersToPoints(W3)
                     .Cell(k + 2, 3).Range.Text = Tol(NumS(i))
                     .Cell(k + 2, 4).Width = CentimetersToPoints(W4)
                    Next k
                    GoTo 98
                 End If
               End With
             Next j
95    tbn = 0
      tbn = getTableNumber(6)
         If tbn = 0 Then
             msg = msg & UText(26410, 21457, 29616) & "6. Product Color" & UText(34920, 26684) & "，Style / Material / Construction " & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
            "6. Product Color table is not found, Data sheet for Style / Material / Construction form can not generate" & CHR(10)
            GoTo 110
        End If
         
         
97        Call MakeNormalDataSheet(tbn, ComItms)
              
              With WD.Tables(tbn)
              .Cell(1, 1).Range.Text = "Data Sheet for " & Chk(NumS(i))
               Set Cel = .Cell(1, 1)
               Call DeleteBookMark(Cel)
               Cel.Range.Bookmarks.add Name:=Ino(NumS(i)) & "_Data", Range:=Cel.Range
            
               For j = 1 To ComItms
                 .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                .Cell(j + 2, 1).Range.Text = ComItx(j)
                .Cell(j + 2, 2).Range.Text = Spe(NumS(i))
                .Cell(j + 2, 3).Range.Text = Tol(NumS(i))
               Next j
           End With
       End If
    Else      '========================CDF
        tb = WD.Tables.count - 1
        tbn = 0
        Cn = 0
        For j = tb To 3 Step -1
          With WD.Tables(j)
            If .Cell(1, 1).Range.Text Like "*Data Sheet for Constructional Data Form (CDF)*" = True Then
              Cn = Cn + 1
              tbn = j
            End If
          End With
        Next j
        If tbn = 0 Then
           tb = WD.Tables.count - 1
           For j = 4 To tb
             With WD.Tables(j)
               If .Cell(1, 1).Range.Text Like "*6. Product Color*" = True Then
                    tbn = j
                    Call MakeCDFDataSheet(tbn, ComItms)
                    
                   
                       Cn = 1
                       'tbn = j
                       GoTo 98
                 End If
               End With
              Next j
              msg = msg + "The CDF data tables were not generated due to not 6. Product Color for locating the table position." & CHR(10) & _
             UText(30001, 20110, 27809, 26377, 25214, 21040) & "6. Product Color" & UText(34920, 26684, 29992, 20110, 23450, 20301) & "，Data Sheet for Constructional Data Form (CDF)" & UText(34920, 26684, 26410, 29983, 25104) & "！"
              GoTo 110
           End If
          
98       If Cn > ComItms Then
             cn1 = 0
             For j = tb To 3 Step -1
                With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*Data Sheet for Constructional Data Form (CDF)*" = True Then
                      WD.Tables(j).Select
                      Selection.Tables(1).Delete
                      Selection.Delete
                       cn1 = cn1 + 1
                       If cn1 = Cn - ComItms Then GoTo 100
                    End If
                End With
              Next j
          ElseIf Cn < ComItms Then
             For k = 1 To ComItms - Cn
                WD.Tables(tbn).Select
                Selection.Copy
                Selection.MoveDown
                Selection.TypeParagraph
                StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
                Selection.Paste
              Next k
          End If
100      tb = WD.Tables.count - 1
         Cn = 0
         For j = 3 To tb - 3
          With WD.Tables(j)
            If .Cell(1, 1).Range.Text Like "*Data Sheet for Constructional Data Form (CDF)*" = True Then
              Cn = Cn + 1
               Set Cel = .Cell(1, 1)
               Call DeleteBookMark(Cel)
               Cel.Range.Bookmarks.add Name:=Ino(NumW(i)) & "_Data" & "_" & Cn, Range:=Cel.Range
              .Cell(2, 1).Range.Text = "Item No.:" & ComItx(Cn)
            End If
          End With
          Next j
     End If
       '=====
108  Next i
   End If
 
 
110  '---------------------------------------------------6. Product Color
    If cnc > 0 Then
      tbn = 0
      tbn = getTableNumber(6)
      If tbn = 0 Then
         msg = msg & UText(26410, 21457, 29616) & "6. Product Color" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
        "6. Product Color table is not found, the check item for Product Color is not imported" & CHR(10)
        GoTo 130
     End If

120      With WD.Tables(tbn)
         Rn = .Rows.count
         Coln = .Columns.count
         .Cell(Rn, 1).Select
         Selection.InsertRowsBelow cnc
         For i = 1 To cnc
              .Rows(Rn + i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
              .Cell(Rn + i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
              .Cell(Rn + i, 1).Range.Text = "6." & Rn + i - 3
               If Len(Chk(NumC(i))) > 1 Then
                  KwChk = UCase(Mid(Chk(NumC(i)), 1, 1)) & Mid(Chk(NumC(i)), 2, Len(Chk(NumC(i))) - 1)
               Else
                  KwChk = Chk(NumC(i))
               End If
               If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
               .Cell(Rn + i, 2).Range.Text = KwChk
               Set Cel = .Cell(Rn + i, 2)
               Call DeleteBookMark(Cel)
               Cel.Range.Bookmarks.add Name:=Ino(NumC(i)), Range:=Cel.Range
              ' .cell(Rn + i, 3).Range.Text = UCase(Mid(Cri(NumD(i)), 1, 1)) & Mid(Cri(NumD(i)), 2, Len(Cri(NumD(i))) - 1)
            If sms(NumC(i)) <> "" Then
               If ComItms = 1 Then
                 sms(NumC(i)) = Replace(sms(NumC(i)), " /item", "")
                 sms(NumC(i)) = Replace(sms(NumC(i)), " /Item", "")
                 .Cell(Rn + i, 3).Range.Text = sms(NumC(i))
              Else
                 If Len(sms(NumC(i))) > 1 Then
                     .Cell(Rn + i, 3).Range.Text = UCase(Mid(sms(NumC(i)), 1, 1)) & Mid(sms(NumC(i)), 2, Len(sms(NumC(i))) - 1)
                Else
                     .Cell(Rn + i, 3).Range.Text = sms(NumC(i))
                 End If
              End If
           Else
              Set Cel = .Cell(Rn + i, 3)
              Cel.Select
               Call CellAddDropDownList(Cel)

           End If

            If .Cell(Rn, Coln).Range.ContentControls.count > 0 Then
                 .Cell(Rn, Coln).Range.Select
                  Selection.Copy
                 .Cell(Rn + i, Coln).Range.Select
                  StartTime = Timer '
                  Do While Timer - StartTime < 0.1
                    DoEvents
                  Loop
                  Selection.Paste
                  Selection.ContentControls(1).DropdownListEntries(1).Select
            End If
        Next i
      End With
     
        '6. Product Color---
     TbnB = tbn
     DelTab = ""
     
   
125  For i = 1 To cnc
      If Nat(NumC(i)) = UText(23450, 37327) Then
        tbn = 0
        tbn = getTableNumber(7)
         If tbn = 0 Then
             msg = msg & UText(26410, 21457, 29616) & "7. Product Label / Marking" & UText(34920, 26684) & "，Product Color" & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
            "7. Product Label / Marking table is not found, Data sheet for Product Color form can not generate" & CHR(10)
             GoTo 130
         End If

127    Call MakeNormalDataSheet(tbn, ComItms)
              
              With WD.Tables(tbn)
              .Cell(1, 1).Range.Text = "Data Sheet for " & Chk(NumC(i))
               Set Cel = .Cell(1, 1)
               Call DeleteBookMark(Cel)
               Cel.Range.Bookmarks.add Name:=Ino(NumC(i)) & "_Data", Range:=Cel.Range
              
               For j = 1 To ComItms
                 .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                .Cell(j + 2, 1).Range.Text = ComItx(j)
                .Cell(j + 2, 2).Range.Text = Spe(NumC(i))
                .Cell(j + 2, 3).Range.Text = Tol(NumC(i))
               Next j
           End With
       End If
128  Next i
   End If
   
130    ''------------------------------------------7. Product Label / Marking
   If cnl > 0 Then
         tbn = 0
     tbn = getTableNumber(7)
     If tbn = 0 Then
     
         msg = msg & UText(26410, 21457, 29616) & "7. Product Label / Marking" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
        "7. Product Label / Marking table is not found, the check item for Product Label / Marking is not imported" & CHR(10)
        GoTo 150
     End If

140   With WD.Tables(tbn)
            Rn = .Rows.count
            Coln = .Columns.count
           .Cell(Rn, 1).Select
            Selection.InsertRowsBelow cnl
           For i = 1 To cnl
              .Rows(Rn + i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
              .Cell(Rn + i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
         
              .Cell(Rn + i, 1).Range.Text = "7." & Rn + i - 3
              If Len(Chk(NumL(i))) > 1 Then
                KwChk = UCase(Mid(Chk(NumL(i)), 1, 1)) & Mid(Chk(NumL(i)), 2, Len(Chk(NumL(i))) - 1)
              Else
                KwChk = Chk(NumL(i))
              End If
              If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
         
             .Cell(Rn + i, 2).Range.Text = KwChk
              Set Cel = .Cell(Rn + i, 2)
              Call DeleteBookMark(Cel)
              Cel.Range.Bookmarks.add Name:=Ino(NumL(i)), Range:=Cel.Range
              ' .cell(Rn + i, 3).Range.Text = UCase(Mid(Cri(NumL(i)), 1, 1)) & Mid(Cri(NumL(i)), 2, Len(Cri(NumL(i))) - 1)
             
                If sms(NumL(i)) <> "" Then
                  If ComItms = 1 Then
                    sms(NumL(i)) = Replace(sms(NumL(i)), " /item", "")
                    sms(NumL(i)) = Replace(sms(NumL(i)), " /Item", "")
                    .Cell(Rn + i, 3).Range.Text = sms(NumL(i))
                  Else
                     If Len(sms(NumL(i))) > 1 Then
                       .Cell(Rn + i, 3).Range.Text = UCase(Mid(sms(NumL(i)), 1, 1)) & Mid(sms(NumL(i)), 2, Len(sms(NumL(i))) - 1)
                     Else
                        .Cell(Rn + i, 3).Range.Text = sms(NumL(i))
                     End If
                  End If
               Else
                  Set Cel = .Cell(Rn + i, 3)
                  Cel.Select
                  Call CellAddDropDownList(Cel)

             End If
        
                If .Cell(Rn, Coln).Range.ContentControls.count > 0 Then
                 .Cell(Rn, Coln).Range.Select
                  Selection.Copy
                 .Cell(Rn + i, Coln).Range.Select
                  StartTime = Timer '
                  Do While Timer - StartTime < 0.1
                    DoEvents
                  Loop
                  Selection.Paste
                  Selection.ContentControls(1).DropdownListEntries(1).Select
                End If
             Next i
       End With
           
           '7. Product Label / Marking--
    TbnB = tbn
     DelTab = ""
   
   For i = 1 To cnl
      If Nat(NumL(i)) = UText(23450, 37327) Then
         If Chk(NumL(i)) = "Readability test for Barcode,QR code" Then GoTo 148   'BadCode，GenerateBarCodeDataSheet
         DelTab = DelTab & "," & "Data Sheet for " & Chk(NumL(i))  ' translated comment
         
         Select Case TableChoise
         Case 1
             TbCount1 = 0
               For j = 3 To tb - 3
                With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumL(i)) & "*" = True Then
                       If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                            TbCount1 = TbCount1 + 1
                           GoTo 148
                      End If
                  End If
              End With
            Next j
       Case 2
            TbCount2 = 0
            For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumL(i)) & "*" = True Then
                      If .Rows(2).Cells.count = 1 Then
                        TbCount2 = TbCount2 + 1
                    End If
                  End If
                 End With
            Next j
       Case 3
            TbCount1 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
                 With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumL(i)) & "*" = True Then
                   If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                      TbCount1 = TbCount1 + 1
                    End If
                  End If
                End With
            Next j
            
            TbCount2 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumL(i)) & "*" = True Then
                      If .Rows(2).Cells.count = 1 Then
                         TbCount2 = TbCount2 + 1
                     End If
                   End If
                 End With
            Next j
     End Select

145  tbn = 0
     tbn = getTableNumber(8)
         If tbn = 0 Then
            msg = msg & UText(26410, 21457, 29616) & "8. Shipping Mark" & UText(34920, 26684) & "，Product Label / Marking" & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
           "8. Shipping Mark table is not found, Data sheet for Product Label / Marking form can not generate" & CHR(10)
            GoTo 150
         End If
147      Select Case TableChoise
         Case "1"
             If TbCount1 = 0 Then
                tbn = getTableNumber(8)
                Call MakeNormalDataSheet(tbn, ComItms)
                With WD.Tables(tbn)
                 Set Cel = .Cell(1, 1)
                  Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                  Call DeleteBookMark(Cel)
                  Cel.Range.Bookmarks.add Name:=Ino(NumL(i)) & "_Data", Range:=Cel.Range
                  For j = 1 To ComItms
                   .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                   .Cell(j + 2, 1).Range.Text = ComItx(j)
                   .Cell(j + 2, 2).Range.Text = Spe(NumL(i))
                   .Cell(j + 2, 3).Range.Text = Tol(NumL(i))
                  Next j
              End With
            End If
 
        Case "2"
               For j = 1 To ComItms
                    tbn = getTableNumber(8)
                    Call MakeWithPicDataSheet(tbn, ComItms)
                    With WD.Tables(tbn)
                      Set Cel = .Cell(1, 1)
                      Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                      Call DeleteBookMark(Cel)
                      Cel.Range.Bookmarks.add Name:=Ino(NumL(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                      Set Cel = .Cell(2, 1)
                      Cel.Range.Text = "Item No.:" & ComItx(j + TbCount2)
           
                    End With
              Next j
         Case "3"
              If TbCount1 = 0 Then
                   tbn = getTableNumber(8)
                   Call MakeNormalDataSheet(tbn, ComItms)
                   With WD.Tables(tbn)
                       Set Cel = .Cell(1, 1)
                       Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                       Call DeleteBookMark(Cel)
                       Cel.Range.Bookmarks.add Name:=Ino(NumL(i)) & "_Data", Range:=Cel.Range
                       For j = 1 To ComItms
                          .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                          .Cell(j + 2, 1).Range.Text = ComItx(j)
                          .Cell(j + 2, 2).Range.Text = Spe(NumL(i))
                         .Cell(j + 2, 3).Range.Text = Tol(NumL(i))
                       Next j
                  End With
             End If
              
              For j = 1 To ComItms
                   tbn = getTableNumber(8)
                   Call MakeWithPicDataSheet(tbn, ComItms)
                     With WD.Tables(tbn)
                         Set Cel = .Cell(1, 1)
                         Cel.Range.Text = "Data Sheet for " & Chk(NumL(i))
                         Call DeleteBookMark(Cel)
                         Cel.Range.Bookmarks.add Name:=Ino(NumL(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                         Set Cel = .Cell(2, 1)
                   End With
             Next j
        End Select
       End If
148  Next i
      
  End If
150 ''------------------------------------------8. Shipping Mark
   If cnm > 0 Then
     tbn = 0
     tbn = getTableNumber(8)
     tb = WD.Tables.count - 1
    If tbn = 0 Then
         msg = msg & UText(26410, 21457, 29616) & "8. Shipping Mark" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
        "8. Shipping Mark table is not found, the check item for  Shipping Mark is not imported" & CHR(10)
         GoTo 170
     End If

160   With WD.Tables(tbn)
            Rn = .Rows.count
            Coln = .Columns.count
           .Cell(Rn, 1).Select
            Selection.InsertRowsBelow cnm
           For i = 1 To cnm
              .Rows(Rn + i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
              .Cell(Rn + i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
         
              .Cell(Rn + i, 1).Range.Text = "8." & Rn + i - 3
              If Len(Chk(NumM(i))) > 1 Then
                KwChk = UCase(Mid(Chk(NumM(i)), 1, 1)) & Mid(Chk(NumM(i)), 2, Len(Chk(NumM(i))) - 1)
              Else
                KwChk = Chk(NumM(i))
              End If
              If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
         
             .Cell(Rn + i, 2).Range.Text = KwChk
              Set Cel = .Cell(Rn + i, 2)
              Call DeleteBookMark(Cel)
              Cel.Range.Bookmarks.add Name:=Ino(NumM(i)), Range:=Cel.Range
              '.Cell(rn + i, 3).Range.Text = UCase(Mid(Cri(NumL(i)), 1, 1)) & Mid(Cri(NumL(i)), 2, Len(Cri(NumL(i))) - 1)
             
                If sms(NumM(i)) <> "" Then
                  If ComItms = 1 Then
                    sms(NumM(i)) = Replace(sms(NumM(i)), " /item", "")
                    sms(NumM(i)) = Replace(sms(NumM(i)), " /Item", "")
                    .Cell(Rn + i, 3).Range.Text = sms(NumM(i))
                  Else
                     If Len(sms(NumM(i))) > 1 Then
                       .Cell(Rn + i, 3).Range.Text = UCase(Mid(sms(NumM(i)), 1, 1)) & Mid(sms(NumM(i)), 2, Len(sms(NumM(i))) - 1)
                     Else
                        .Cell(Rn + i, 3).Range.Text = sms(NumM(i))
                     End If
                  End If
               Else
                  Set Cel = .Cell(Rn + i, 3)
                  Cel.Select
                 Call CellAddDropDownList(Cel)

             End If
        
                If .Cell(Rn, Coln).Range.ContentControls.count > 0 Then
                 .Cell(Rn, Coln).Range.Select
                  Selection.Copy
                 .Cell(Rn + i, Coln).Range.Select
                  StartTime = Timer '
                  Do While Timer - StartTime < 0.1
                    DoEvents
                  Loop
                  Selection.Paste
                  Selection.ContentControls(1).DropdownListEntries(1).Select
                End If
             Next i
       End With
         
         
         '8. Shipping Mark-
    TbnB = tbn
     DelTab = ""
   
   For i = 1 To cnm
     If Nat(NumM(i)) = UText(23450, 37327) Then

          DelTab = DelTab & "," & "Data Sheet for " & Chk(NumM(i))  ' translated comment
         Select Case TableChoise
         Case 1
              TbCount1 = 0
               For j = 3 To tb - 3
                With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumM(i)) & "*" = True Then
                      If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                            TbCount1 = TbCount1 + 1
                           GoTo 168
                    End If
                  End If
              End With
            Next j
       Case 2
            TbCount2 = 0
            For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumM(i)) & "*" = True Then
                      If .Rows(2).Cells.count = 1 Then
                        TbCount2 = TbCount2 + 1
                    End If
                  End If
                 End With
            Next j
       Case 3
            TbCount1 = 0
            tb = WD.Tables.count - 1
            For j = 3 To tb - 3
                 With WD.Tables(j)
                   If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumM(i)) & "*" = True Then
                  If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                      TbCount1 = TbCount1 + 1
                    End If
                  End If
                End With
            Next j
            
            TbCount2 = 0
            tb = WD.Tables.count - 1
             For j = 3 To tb - 3
                 With WD.Tables(j)
                     If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumM(i)) & "*" = True Then
                      If .Rows(2).Cells.count = 1 Then
                         TbCount2 = TbCount2 + 1
                      End If
                   End If
                 End With
            Next j
     End Select
       
165   tbn = 0
      tbn = getTableNumber(9)
      If tbn = 0 Then
          msg = msg & UText(26410, 21457, 29616) & "9. Packing / Packaging" & UText(34920, 26684) & "，Shipping Mark-" & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
         "9. Packing / Packaging table is not found, Data sheet for Shipping Mark form can not generate" & CHR(10)
          GoTo 170
      End If
167     Select Case TableChoise
         Case "1"
             If TbCount1 = 0 Then
                 tbn = getTableNumber(9)
                 Call MakeNormalDataSheet(tbn, ComItms)
                With WD.Tables(tbn)
                  Set Cel = .Cell(1, 1)
                  Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                  Call DeleteBookMark(Cel)
                  Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data", Range:=Cel.Range
                  For j = 1 To ComItms
                   .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                   .Cell(j + 2, 1).Range.Text = ComItx(j)
                   .Cell(j + 2, 2).Range.Text = Spe(NumM(i))
                   .Cell(j + 2, 3).Range.Text = Tol(NumM(i))
                  Next j
            End With
         End If
 
        Case "2"
               For j = 1 To ComItms
                    tbn = getTableNumber(9)
                    Call MakeWithPicDataSheet(tbn, ComItms)
                    With WD.Tables(tbn)
                      Set Cel = .Cell(1, 1)
                      Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                      Call DeleteBookMark(Cel)
                      Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                      Set Cel = .Cell(2, 1)
                      Cel.Range.Text = "Item No.:" & ComItx(j + TbCount2)
           
                    End With
              Next j
         Case "3"
              If TbCount1 = 0 Then
                   tbn = getTableNumber(9)
                   Call MakeNormalDataSheet(tbn, ComItms)
                   With WD.Tables(tbn)
                       Set Cel = .Cell(1, 1)
                       Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                       Call DeleteBookMark(Cel)
                       Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data", Range:=Cel.Range
                       For j = 1 To ComItms
                          .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                          .Cell(j + 2, 1).Range.Text = ComItx(j)
                          .Cell(j + 2, 2).Range.Text = Spe(NumM(i))
                         .Cell(j + 2, 3).Range.Text = Tol(NumM(i))
                       Next j
                  End With
             End If
              
              For j = 1 To ComItms
                   tbn = getTableNumber(9)
                   Call MakeWithPicDataSheet(tbn, ComItms)
                     With WD.Tables(tbn)
                         Set Cel = .Cell(1, 1)
                         Cel.Range.Text = "Data Sheet for " & Chk(NumM(i))
                         Call DeleteBookMark(Cel)
                         Cel.Range.Bookmarks.add Name:=Ino(NumM(i)) & "_Data" & "_" & j + TbCount2, Range:=Cel.Range
                         Set Cel = .Cell(2, 1)
                   End With
             Next j
        End Select
       End If
168  Next i
   End If
    
170 '------------------------------------------9. Packing / Packaging
   
   If cnp > 0 Then
          tbn = 0
     tbn = getTableNumber(9)
     If tbn = 0 Then
        msg = msg & UText(26410, 21457, 29616) & "9. Packing / Packaging" & UText(30340, 34920, 26684) & "，" & UText(35813, 37096, 20998, 27979, 35797, 39033, 30446, 26410, 23548, 20837) & CHR(10) & _
       "9. Packing / Packaging is not found, the check item for Packing / Packaging is not imported" & CHR(10)
       GoTo 190
     End If

180    With WD.Tables(tbn)
            Rn = .Rows.count
            Coln = .Columns.count
           .Cell(Rn, 1).Select
            Selection.InsertRowsBelow cnp
            For i = 1 To cnp
              .Rows(Rn + i).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
              .Cell(Rn + i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
               .Cell(Rn + i, 1).Range.Text = "9." & Rn + i - 3
               If Len(Chk(NumP(i))) > 1 Then
                  KwChk = UCase(Mid(Chk(NumP(i)), 1, 1)) & Mid(Chk(NumP(i)), 2, Len(Chk(NumP(i))) - 1)
               Else
                  KwChk = Chk(NumP(i))
               End If
               If InStr(1, KwChk, "(") > 0 Then KwChk = Trim(Mid(KwChk, 1, InStr(1, KwChk, "(") - 1))
             .Cell(Rn + i, 2).Range.Text = KwChk
              Set Cel = .Cell(Rn + i, 2)
              Call DeleteBookMark(Cel)
              Cel.Range.Bookmarks.add Name:=Ino(NumP(i)), Range:=Cel.Range
             '.Cell(i + 6, 3).Range.Text = UCase(Mid(Cri(NumP(i)), 1, 1)) & Mid(Cri(NumP(i)), 2, Len(Cri(NumP(i))) - 1)
               
            If sms(NumP(i)) <> "" Then
              If ComItms = 1 Then
                sms(NumP(i)) = Replace(sms(NumP(i)), " /item", "")
                sms(NumP(i)) = Replace(sms(NumP(i)), " /Item", "")
                 .Cell(Rn + i, 3).Range.Text = sms(NumP(i))
               Else
                  If Len(sms(NumP(i))) > 1 Then
                     .Cell(Rn + i, 3).Range.Text = UCase(Mid(sms(NumP(i)), 1, 1)) & Mid(sms(NumP(i)), 2, Len(sms(NumP(i))) - 1)
                  Else
                      .Cell(Rn + i, 3).Range.Text = sms(NumP(i))
                  End If
               End If
            Else
               Set Cel = .Cell(Rn + i, 3)
               Cel.Select
               Call CellAddDropDownList(Cel)

            End If

                 
                 If .Cell(Rn, Coln).Range.ContentControls.count > 0 Then
                 .Cell(Rn, Coln).Range.Select
                  Selection.Copy
                 .Cell(Rn + i, Coln).Range.Select
                  StartTime = Timer '
                  Do While Timer - StartTime < 0.1
                    DoEvents
                  Loop
                  Selection.Paste
                  Selection.ContentControls(1).DropdownListEntries(1).Select
                End If
            Next i
           
      End With
      
      'Packing / Packaging---
     TbnB = tbn
     DelTab = ""
      

185   For i = 1 To cnp
        If Nat(NumP(i)) = UText(23450, 37327) Then
           tb = WD.Tables.count - 1
           For j = 3 To tb - 3
              With WD.Tables(j)
               If .Cell(1, 1).Range.Text Like "*" & "Data Sheet for " & Chk(NumP(i)) & "*" = True Then
                  tbn = j
                  GoTo 187
                End If
              End With
           Next j
          tbn = 0
         tbn = getTableNumber(10)
          If tbn = 0 Then
             msg = msg & UText(26410, 21457, 29616) & "10. Other Note" & UText(34920, 26684) & ",Packing / Packaging" & UText(25968, 25454, 34920, 26684, 26080, 27861, 29983, 25104) & CHR(10) & _
           " 10. Other Note table is not found, Data sheet for Packing / Packaging form can not generate" & CHR(10)
            GoTo 190
          End If
187       Call MakeNormalDataSheet(tbn, ComItms)
              
              With WD.Tables(tbn)
              .Cell(1, 1).Range.Text = "Data Sheet for " & Chk(NumP(i))
               Set Cel = .Cell(1, 1)
               Call DeleteBookMark(Cel)
               Cel.Range.Bookmarks.add Name:=Ino(NumP(i)) & "_Data", Range:=Cel.Range
            
               For j = 1 To ComItms
                .Rows(j + 2).Range.Cells.Shading.BackgroundPatternColor = -603914241  '-603914241 ' -603917569 '
                .Cell(j + 2, 1).Range.Text = ComItx(j)
                .Cell(j + 2, 2).Range.Text = Spe(NumP(i))
                .Cell(j + 2, 3).Range.Text = Tol(NumP(i))
               Next j
           End With
       End If
188   Next i
    End If
    
190 '-

 
300  If msg <> "" Then
      MsgBox msg & CHR(10) & CHR(10) & UText(27979, 35797, 39033, 30446, 23548, 20837, 23436, 25104) & "，" & UText(35831, 26681, 25454, 23454, 38469, 24773, 20917, 36827, 34892, 20462, 25913, 25110, 34917, 20805) & "," & UText(35831, 19981, 35201, 37325, 22797, 36816, 34892, 26412, 31243, 24207, 21542, 21017, 20462, 25913, 25110, 34917, 20805, 30340, 25968, 25454, 23558, 34987, 35206, 30422) & "!" & CHR(10) & _
     "Complete importing test items, please modify or add according to the actual situation, do no Run this program repeatedly or the data will be overwritten"
    Else
     MsgBox UText(27979, 35797, 39033, 30446, 23548, 20837, 23436, 25104) & "，" & UText(35831, 26681, 25454, 23454, 38469, 24773, 20917, 36827, 34892, 20462, 25913, 25110, 34917, 20805) & "," & UText(35831, 19981, 35201, 20877, 36816, 34892, 26412, 31243, 24207, 21542, 21017, 20462, 25913, 25110, 34917, 20805, 30340, 25968, 25454, 23558, 34987, 35206, 30422) & "!" & CHR(10) & _
     "Complete importing test items, please modify or add according to the actual situation,  do not Run this program repeatedly or the data will be overwritten"
    End If

1000 End Sub
Sub Input_Defects(Cn As Integer, Defect As Variant, DefectE As Variant, DefectCode As Variant)

'by wyn 202407
Dim i, j As Integer
Dim Rn As Integer
Dim Rd As Integer
Dim Coln As Integer
Dim Rng As Range
Dim BK As bookmark
  
10   If Selection.Information(wdWithInTable) Then
       Set tbl = Selection.Tables(1)
       Set Cel = Selection.Cells(1)
  
  With tbl
     Rn = .Rows.count
     Coln = .Columns.count
     Rc = Cel.rowIndex
     cc = Cel.ColumnIndex
     For i = Rc To Rn
       If .Rows(i).Cells.count = 1 Then
          Rd = i - Rc
          GoTo 20
       Else
         If .Cell(i, Coln - 3).Range.Fields.count > 0 Then
            Rd = i - Rc
            GoTo 20
         End If
       End If
     Next i
     
20  If Rd <= Cn Then
       .Rows(Rc).Select
        Selection.Copy
       For i = 1 To Cn - Rd + 1
            .Rows(Rc).Select
                        StartTime = Timer '' 100
                        Do While Timer - StartTime < 0.1
                          DoEvents
                        Loop
         Selection.Paste
       Next i
     ElseIf Rd > Cn Then
       For i = 1 To Rd - Cn - 1
         .Rows(Rc).Delete
       Next i
     End If
     For i = 1 To Cn
       If Language = "Chinese" Then
         .Cell(Rc + i - 1, Coln - 3).Range.Text = Defect(i)
       Else
          DefectE(i) = UCase(Mid(DefectE(i), 1, 1)) & Mid(DefectE(i), 2, Len(DefectE(i)) - 1)
          .Cell(Rc + i - 1, Coln - 3).Range.Text = DefectE(i)
       End If
       If .Cell(Rc + i - 1, Coln - 3).Range.Bookmarks.count > 0 Then
           For Each BK In .Cell(Rc + i - 1, Coln - 3).Range.Bookmarks
              BK.Delete
           Next BK
       End If
       Set Cel = .Cell(Rc + i - 1, Coln - 3)
       Call DeleteBookMark(Cel)
       Cel.Range.Bookmarks.add Name:=DefectCode(i), Range:=Cel.Range
     Next i
   End With
 End If
End Sub
Sub Replace_Defects(Cn As Integer, Defect As Variant, DefectE As Variant, DefectCode As Variant)

'by wyn 202407
Dim i, j As Integer
Dim Rn As Integer
Dim Rd As Integer
Dim Coln As Integer
Dim Rng As Range
Dim bookmark As bookmark
  
10   If Selection.Information(wdWithInTable) Then
       Set tbl = Selection.Tables(1)
       Set Cel = Selection.Cells(1)
  
  With tbl
     Coln = .Columns.count
     Rc = Cel.rowIndex
     cc = Cel.ColumnIndex
     
     Set Cel = .Cell(Rc, cc + 1)
     
       If Language = "Chinese" Then
         Cel.Range.Text = Defect(Cn)
       Else
         DefectE(Cn) = UCase(Mid(DefectE(Cn), 1, 1)) & Mid(DefectE(Cn), 2, Len(DefectE(Cn)) - 1)
         Cel.Range.Text = DefectE(Cn)
       End If
       Call DeleteBookMark(Cel)
       Cel.Range.Bookmarks.add Name:=DefectCode(Cn), Range:=Cel.Range
  
   End With
 End If
End Sub
Sub InputDropDown_Defects(DefName As String, DefNameE As String, DefCode As String)

'by wyn 20250107
Dim i, j As Integer
Dim Rn As Integer
Dim Rd As Integer
Dim Coln As Integer
Dim Rng As Range
Dim bookmark As bookmark
  
10   If Selection.Information(wdWithInTable) Then
       Set tbl = Selection.Tables(1)
       Set Cel = Selection.Cells(1)
  
  With tbl
     Coln = .Columns.count
     Rc = Cel.rowIndex
     cc = Cel.ColumnIndex
     
     Set Cel = .Cell(Rc, cc + 1)
     
       If Language = "Chinese" Then
         Cel.Range.Text = DefName
       Else
         DefNameE = UCase(Mid(DefNameE, 1, 1)) & Mid(DefNameE, 2, Len(DefNameE) - 1)
         Cel.Range.Text = DefNameE
       End If
       Call DeleteBookMark(Cel)
       Cel.Range.Bookmarks.add Name:=DefCode, Range:=Cel.Range
  
   End With
 End If
End Sub
Sub Setborders025()
With Selection.Tables(1)
                  With .Borders(wdBorderLeft)
                      .LineStyle = wdLineStyleSingle
                      .LineWidth = wdLineWidth025pt
                    .Color = wdColorAutomatic
                     End With
                     With .Borders(wdBorderRight)
                       .LineStyle = wdLineStyleSingle
                       .LineWidth = wdLineWidth025pt
                       .Color = wdColorAutomatic
                     End With
                     With .Borders(wdBorderTop)
                      .LineStyle = wdLineStyleSingle
                      .LineWidth = wdLineWidth025pt
                      .Color = wdColorAutomatic
                     End With
                     With .Borders(wdBorderBottom)
                      .LineStyle = wdLineStyleSingle
                      .LineWidth = wdLineWidth025pt
                      .Color = wdColorAutomatic
                    End With
                    With .Borders(wdBorderHorizontal)
                       .LineStyle = wdLineStyleSingle
                       .LineWidth = wdLineWidth025pt
                       .Color = wdColorAutomatic
                    End With
                    With .Borders(wdBorderVertical)
                        .LineStyle = wdLineStyleSingle
                         .LineWidth = wdLineWidth025pt
                        .Color = wdColorAutomatic
                   End With
End With

End Sub




Sub TestNo()
Dim Rn As Integer

With Selection.Tables(1)
  Rn = .Rows.count
  For i = 4 To Rn
   .Cell(i, 1).Range.Text = "3." & i - 3
  
  Next i

End With
End Sub


Sub WorkmanshipNo()
Dim Rn As Integer

With Selection.Tables(1)
  Rn = .Rows.count
  For i = 7 To Rn
   .Cell(i, 1).Range.Text = "2." & i - 6
  
  Next i

End With
End Sub

Sub DimensionNo()
Dim Rn As Integer

With Selection.Tables(1)
  Rn = .Rows.count
  For i = 4 To Rn
   .Cell(i, 1).Range.Text = "4." & i - 3
  
  Next i

End With
End Sub


Sub StyleNo()
Dim Rn As Integer

With Selection.Tables(1)
  Rn = .Rows.count
  For i = 4 To Rn
   .Cell(i, 1).Range.Text = "5." & i - 3
  
  Next i

End With
End Sub

Sub ColorNo()
Dim Rn As Integer

With Selection.Tables(1)
  Rn = .Rows.count
  For i = 4 To Rn
   .Cell(i, 1).Range.Text = "6." & i - 3
  
  Next i

End With
End Sub

Sub LabelNo()
Dim Rn As Integer

With Selection.Tables(1)
  Rn = .Rows.count
  For i = 4 To Rn
   .Cell(i, 1).Range.Text = "7." & i - 3
  
  Next i

End With
End Sub


Sub ShippingNo()
Dim Rn As Integer

With Selection.Tables(1)
  Rn = .Rows.count
  For i = 4 To Rn
   .Cell(i, 1).Range.Text = "8." & i - 3
  
  Next i

End With
End Sub
Sub PackingNo()
Dim Rn As Integer

With Selection.Tables(1)
  Rn = .Rows.count
  For i = 4 To Rn
   .Cell(i, 1).Range.Text = "9." & i - 3
  
  Next i

End With
End Sub

Sub NoteNo()
Dim Rn As Integer

With Selection.Tables(1)
  Rn = .Rows.count
  For i = Rn To 4 Step -1
     If Len(Trim(.Cell(i, 2).Range.Text)) <= 2 Then
    .Rows(i).Delete
     End If
  Next i
  Rn = .Rows.count
  If Rn > 3 Then
 
   For i = 4 To Rn
    .Cell(i, 1).Range.Text = "10." & i - 3
  
   Next i
  End If
End With
End Sub
Sub CDFNo()
Dim Rn As Integer

With Selection.Tables(1)
  Rn = .Rows.count
  For i = Rn To 5 Step -1
     If Len(Trim(.Cell(i, 2).Range.Text)) <= 2 And Len(Trim(.Cell(i, 3).Range.Text)) <= 2 Then
    .Rows(i).Delete
     End If
  Next i
  Rn = .Rows.count
 
 
   For i = 4 To Rn
    .Cell(i, 1).Range.Text = i - 3
  
   Next i
 
End With
End Sub

Sub SetRemarkTitle(ByRef Title() As String)
        ReDim Title(0 To 10)
             Title(1) = "RemarkType"
             Title(2) = "RemarkCode"
             Title(3) = "No"
             Title(4) = "Description"
             Title(5) = "Category"
End Sub
Sub SetCheckPointTitle(ByRef Title() As String)
        ReDim Title(0 To 10)
             Title(0) = "CheckPointCode"
             Title(1) = "No"
             Title(2) = "CheckPoint"
             Title(3) = "Criteria"
             Title(4) = "SampleSize"
             Title(5) = "Results"
End Sub
Sub SetTestDataTitle(ByRef Title() As String)
           ReDim Title(0 To 10)
           Title(1) = "ItemNo"
           Title(2) = "Specification"
           Title(3) = "Tolerance"
           Title(4) = "Result"
           Title(5) = "Position"  '"Position"
           Title(6) = "Size"
           Title(7) = "Color"
           'Title(8) = "Position"
End Sub

Sub SAPForQuantity(Remark_col, ByRef Title() As String)
          
 Dim tbn, Rn, CLn, Cn, Itms As Integer
 Dim i, j, k As Integer
 Dim Kw  As String


 Dim cell_dict As Object
 Set WD = ThisDocument

      tbn = 0
      tbn = getTableNumber(1)
      If tbn = 0 Then
         MsgBox tb_tt(1) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(1) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
      Dim Qty_Dict As Object
      Set Qty_Dict = CreateObject("Scripting.Dictionary")
       With WD.Tables(tbn)
          Rn = .Rows.count
          CLn = .Columns.count
          Itms = Rn - 5
          Cn = 0
     
 
    End With
       Dim SAPForQty_Dict As Object
       Set SAPForQty_Dict = CreateObject("Scripting.Dictionary")
       Dim SAPForQty_Rows As Collection
       Set SAPForQty_Rows = New Collection
       tb = WD.Tables.count - 1
       
       For i = 1 To tb
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Special Attention Point for Quantity*" = True Then
            Rn = .Rows.count
            SetRemarkTitle Title
            For j = 2 To Rn
                Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Special Attention Point"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                cell_dict.add Title(5), "Quantity"
                Remark_col.add cell_dict
                'SpecialAttentionPointforQty_Rows.add cell_dict
            Next j
            GoTo 28
         End If
        End With
       Next i
            Set cell_dict = New Scripting.Dictionary
            SetRemarkTitle Title
            cell_dict.add Title(1), "Special Attention Point"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
            cell_dict.add Title(5), "Quantity"
            Remark_col.add cell_dict

28 End Sub
Sub NoteForQuantity(Remark_col, ByRef Title() As String)

 Dim tbn, Rn, CLn, Cn, Itms As Integer
 Dim i, j, k As Integer
 Dim Kw  As String

 Dim cell_dict As Object
 
Set WD = ThisDocument
      
      tbn = 0
      tbn = getTableNumber(1)
      If tbn = 0 Then
         MsgBox tb_tt(1) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(1) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
      Dim Qty_Dict As Object
      Set Qty_Dict = CreateObject("Scripting.Dictionary")
       With WD.Tables(tbn)
          Rn = .Rows.count
          CLn = .Columns.count
          Itms = Rn - 5
          Cn = 0
     
 
    End With
       
       
       
       Dim NoteForQty_Dict As Object
       Set NoteForQty_Dict = CreateObject("Scripting.Dictionary")
       Dim NoteForQty_Rows As Collection
       Set NoteForQty_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Reference Note for Quantity*" = True Then
            Rn = .Rows.count
            SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Reference Note"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
               cell_dict.add Title(5), "Quantity"
               Remark_col.add cell_dict
               'NoteForQty_Rows.add cell_dict
            Next j
            GoTo 29
         End If
        End With
       Next i
            Set cell_dict = New Scripting.Dictionary
            SetRemarkTitle Title
            cell_dict.add Title(1), "Reference Note"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
            cell_dict.add Title(5), "Quantity"
            Remark_col.add cell_dict
29 End Sub
Sub SAPForWorkmanship(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
      
       Dim SAPForWMS_dict As Object
       Set SAPForWMS_dict = CreateObject("Scripting.Dictionary")
       Dim SAPForWMS_Rows As Collection
       Set SAPForWMS_Rows = New Collection
    
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Special Attention Point for Workmanship*" = True Then
            Rn = .Rows.count
            
            SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Special Attention Point"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                cell_dict.add Title(5), "Workmanship"
                Remark_col.add cell_dict
               'SAPForWMS_Rows.add cell_dict
            Next j
            GoTo 38
         End If
        End With
       Next i
           Set cell_dict = New Scripting.Dictionary
            SetRemarkTitle Title
            cell_dict.add Title(1), "Special Attention Point"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
            cell_dict.add Title(5), "Workmanship"
            Remark_col.add cell_dict

       
38       'SAPForWMS_dict.add "SpecialAttentionPointforWorkmanship", SAPForWMS_Rows
End Sub
Sub NoteForWorkmanship(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
      
       Dim NoteForWMS_dict As Object
       Set NoteForWMS_dict = CreateObject("Scripting.Dictionary")
       Dim NoteForWorkmanship_Rows As Collection
       Set NoteForWorkmanship_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Reference Note for Workmanship*" = True Then
            Rn = .Rows.count
            SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Reference Note"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
               cell_dict.add Title(5), "Workmanship"
               Remark_col.add cell_dict
               'NoteForWorkmanship_Rows.add cell_dict
            Next j
            GoTo 39
         End If
        End With
       Next i
            Set cell_dict = New Scripting.Dictionary
           SetRemarkTitle Title
            cell_dict.add Title(1), "Reference Note"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
            cell_dict.add Title(5), "Workmanship"
            Remark_col.add cell_dict
            'NoteForWorkmanship_Rows.add cell_dict
39        ' NoteForWMS_dict.add "ReferenceNoteforWorkmanship", NoteForWorkmanship_Rows

End Sub
Sub SAPForOnSiteTest(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
             
       Dim SAPForOST_Dict As Object
       Set SAPForOST_Dict = CreateObject("Scripting.Dictionary")
       Dim SAPForOST_Rows As Collection
       Set SAPForOST_Rows = New Collection

       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Special Attention Point for On Site Test*" = True Then
             Rn = .Rows.count
             SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Special Attention Point"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                cell_dict.add Title(5), "On Site Test"
                Remark_col.add cell_dict
               'SAPForOST_Rows.add cell_dict
            Next j
            GoTo 48
         End If
        End With
       Next i
            Set cell_dict = New Scripting.Dictionary
            SetRemarkTitle Title
            cell_dict.add Title(1), "Special Attention Point"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
            cell_dict.add Title(5), "On Site Test"
            Remark_col.add cell_dict

       
       
48        'SAPForOST_Dict.add "SpecialAttentionPointforOnSiteTest", SAPForOST_Rows
       
   
       
End Sub
Sub NoteForOnSiteTest(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object

      Set WD = ThisDocument
           
       
       Dim NoteforOST_Dict As Object
       Set NoteforOST_Dict = CreateObject("Scripting.Dictionary")
       Dim NoteforOST_Rows As Collection
       Set NoteforOST_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Reference Note for On Site Test*" = True Then
            Rn = .Rows.count
             SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Reference Note"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
               cell_dict.add Title(5), "On Site Test"
               Remark_col.add cell_dict
               'NoteforOST_Rows.add cell_dict
            Next j
            GoTo 49
         End If
        End With
       Next i
          Set cell_dict = New Scripting.Dictionary
           SetRemarkTitle Title
            cell_dict.add Title(1), "Reference Note"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
             cell_dict.add Title(5), "On Site Test"
            Remark_col.add cell_dict
            'NoteforOST_Rows.add cell_dict
       
49         'NoteforOST_Dict.add "ReferenceNoteforOnSiteTest", NoteforOST_Rows
       
End Sub
Sub SAPForDimensionweight(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
             
       Dim SAPForDW_Dict As Object
       Set SAPForDW_Dict = CreateObject("Scripting.Dictionary")
       Dim SAPForDW_Rows As Collection
       Set SAPForDW_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Special Attention Point for Product Dimension & Weight*" = True Then
            Rn = .Rows.count
            SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Special Attention Point"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                cell_dict.add Title(5), "Product Dimension & Weight"
                Remark_col.add cell_dict
               'SAPForDW_Rows.add cell_dict
            Next j
            GoTo 58
         End If
        End With
       Next i
             Set cell_dict = New Scripting.Dictionary
            SetRemarkTitle Title
            cell_dict.add Title(1), "Special Attention Point"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
             cell_dict.add Title(5), "Product Dimension & Weight"
            Remark_col.add cell_dict

           ' SAPForDW_Rows.add cell_dict
58        'SAPForDW_Dict.add "SpecialAttentionPointforProductDimension&Weight", SAPForDW_Rows
       
   
   
   
       
End Sub
Sub NoteForDimensionweight(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
                   
       
       Dim NoteForDW_Dict As Object
       Set NoteForDW_Dict = CreateObject("Scripting.Dictionary")
       Dim NoteforDW_Rows As Collection
       Set NoteforDW_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Reference Note for Product Dimension & Weight*" = True Then
            Rn = .Rows.count
             SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Reference Note"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
               cell_dict.add Title(5), "Product Dimension & Weight"
               Remark_col.add cell_dict
               'NoteforDW_Rows.add cell_dict
            Next j
            GoTo 59
         End If
        End With
       Next i
            Set cell_dict = New Scripting.Dictionary
           SetRemarkTitle Title
            cell_dict.add Title(1), "Reference Note"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
            cell_dict.add Title(5), "Product Dimension & Weight"
            Remark_col.add cell_dict
            'NoteforDW_Rows.add cell_dict
         
59       ' NoteForDW_Dict.add "ReferenceNoteforProductDimension&Weight", NoteforDW_Rows
   
       
End Sub
Sub SAPForStyleMaterialConstruction(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
             
        Dim SAPforSMC_Dict As Object
       Set SAPforSMC_Dict = CreateObject("Scripting.Dictionary")
       Dim SAPForSMC_Rows As Collection
       Set SAPForSMC_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Special Attention Point for Style / Material / Construction*" = True Then
            Rn = .Rows.count
             SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Special Attention Point"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                cell_dict.add Title(5), "Style / Material / Construction"
                Remark_col.add cell_dict
               'SAPForSMC_Rows.add cell_dict
            Next j
            GoTo 68
         End If
        End With
       Next i
          Set cell_dict = New Scripting.Dictionary
         SetRemarkTitle Title
            cell_dict.add Title(1), "Special Attention Point"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
            cell_dict.add Title(5), "Style / Material / Construction"
            Remark_col.add cell_dict

            'SAPForSMC_Rows.add cell_dict
       
68       ' SAPforSMC_Dict.add "SpecialAttentionPointforStyle/Material/Construction", SAPForSMC_Rows
       
  
   
       
End Sub
Sub NoteForStyleMaterialConstruction(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
                   
       
            
       Dim NoteForSMC_Dict As Object
       Set NoteForSMC_Dict = CreateObject("Scripting.Dictionary")
       Dim NoteForSMC_Rows As Collection
       Set NoteForSMC_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
         With WD.Tables(i)
           If .Cell(1, 1).Range.Text Like "*Reference Note for Style / Material / Construction*" = True Then
              Rn = .Rows.count
           SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Reference Note"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
               cell_dict.add Title(5), "Style / Material / Construction"
               Remark_col.add cell_dict
                 'NoteForSMC_Rows.add cell_dict
               Next j
               GoTo 69
             End If
          End With
       Next i
         
         Set cell_dict = New Scripting.Dictionary
           SetRemarkTitle Title
            cell_dict.add Title(1), "Reference Note"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
            cell_dict.add Title(5), "Style / Material / Construction"
            Remark_col.add cell_dict
            'NoteForSMC_Rows.add cell_dict
69          'NoteForSMC_Dict.add "ReferenceNoteforStyle/Material/Construction", NoteForSMC_Rows
      
   
       
End Sub
Sub SAPForProductColor(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
             
       Dim SAPForCL_Dict As Object
       Set SAPForCL_Dict = CreateObject("Scripting.Dictionary")
       Dim SAPForCL_Rows As Collection
       Set SAPForCL_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Special Attention Point for Product Color*" = True Then
            Rn = .Rows.count
            SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Special Attention Point"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                cell_dict.add Title(5), "Product Color"
                Remark_col.add cell_dict
               'SAPForCL_Rows.add cell_dict
            Next j
            GoTo 78
         End If
        End With
       Next i
        Set cell_dict = New Scripting.Dictionary
            SetRemarkTitle Title
            cell_dict.add Title(1), "Special Attention Point"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
             cell_dict.add Title(5), "Product Color"
            Remark_col.add cell_dict

            'SAPForCL_Rows.add cell_dict
       
78        ' SAPForCL_Dict.add "SpecialAttentionPointforProductColor", SAPForCL_Rows
       
  
     
       
End Sub
Sub NoteForProductColor(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
                   
       
       Dim NoteforCL_Dict As Object
       Set NoteforCL_Dict = CreateObject("Scripting.Dictionary")
       Dim NoteforCL_Rows As Collection
       Set NoteforCL_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
         With WD.Tables(i)
           If .Cell(1, 1).Range.Text Like "*Reference Note for Product Color*" = True Then
              Rn = .Rows.count
           SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Reference Note"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
              cell_dict.add Title(5), "Product Color"
               Remark_col.add cell_dict
                  'NoteforCL_Rows.add cell_dict
               Next j
               GoTo 79
            End If
         End With
       Next i
        Set cell_dict = New Scripting.Dictionary
           SetRemarkTitle Title
            cell_dict.add Title(1), "Reference Note"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
             cell_dict.add Title(5), "Product Color"
            Remark_col.add cell_dict
            'NoteforCL_Rows.add cell_dict
         
79         'NoteforCL_Dict.add "ReferenceNoteforProductColor", NoteforCL_Rows
     
   
   
       
End Sub
Sub SAPForProductLableMarking(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
             
       
      Dim SAPForLM_Dict As Object
       Set SAPForLM_Dict = CreateObject("Scripting.Dictionary")
       Dim SAPForLM_Rows As Collection
       Set SAPForLM_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
         With WD.Tables(i)
           If .Cell(1, 1).Range.Text Like "*Special Attention Point for Product Label / Marking*" = True Then
              Rn = .Rows.count
           SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Special Attention Point"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                 cell_dict.add Title(5), "Product Label / Marking"
                Remark_col.add cell_dict
                  'SAPForLM_Rows.add cell_dict
             Next j
             GoTo 88
           End If
         End With
       Next i
        Set cell_dict = New Scripting.Dictionary
           SetRemarkTitle Title
            cell_dict.add Title(1), "Special Attention Point"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
            cell_dict.add Title(5), "Product Label / Marking"
            Remark_col.add cell_dict

            'SAPForLM_Rows.add cell_dict
         
88         'SAPForLM_Dict.add "SpecialAttentionPointforProductLabel/Marking", SAPForLM_Rows
       
     
           
    
   
End Sub
Sub NoteForProductLableMarking(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
                   
        Dim NoteForLM_Dict As Object
       Set NoteForLM_Dict = CreateObject("Scripting.Dictionary")
       Dim NoteforLM_Rows As Collection
       Set NoteforLM_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
          With WD.Tables(i)
            If .Cell(1, 1).Range.Text Like "*Reference Note for Product Label / Marking*" = True Then
              Rn = .Rows.count
            SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Reference Note"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
               cell_dict.add Title(5), "Product Label / Marking"
               Remark_col.add cell_dict
                 'NoteforLM_Rows.add cell_dict
               Next j
               GoTo 89
            End If
          End With
       Next i
       Set cell_dict = New Scripting.Dictionary
           SetRemarkTitle Title
            cell_dict.add Title(1), "Reference Note"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
            cell_dict.add Title(5), "Product Label / Marking"
            Remark_col.add cell_dict
            'NoteforLM_Rows.add cell_dict
89        ' NoteForLM_Dict.add "ReferenceNoteforProductLabel/Marking", NoteforLM_Rows
  
       
End Sub
Sub SAPForShippingMark(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
             
       
       Dim SAPForSM_Dict As Object
       Set SAPForSM_Dict = CreateObject("Scripting.Dictionary")
       Dim SAPForSM_Rows As Collection
       Set SAPForSM_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Special Attention Point for Shipping Mark*" = True Then
            Rn = .Rows.count
           SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Special Attention Point"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                cell_dict.add Title(5), "Shipping Mark"
                Remark_col.add cell_dict
                  ' SAPForSM_Rows.add cell_dict
             Next j
             GoTo 98
         End If
        End With
       Next i
        Set cell_dict = New Scripting.Dictionary
          SetRemarkTitle Title
            cell_dict.add Title(1), "Special Attention Point"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
              cell_dict.add Title(5), "Shipping Mark"
            Remark_col.add cell_dict

           ' SAPForSM_Rows.add cell_dict
98       ' SAPForSM_Dict.add "SpecialAttentionPointforShippingMark", SAPForSM_Rows
   
End Sub
Sub NoteForShippingMark(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
                   
        
       Dim NoteForSM_Dict As Object
       Set NoteForSM_Dict = CreateObject("Scripting.Dictionary")
       Dim NoteForSM_Rows As Collection
       Set NoteForSM_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Reference Note for Shipping Mark*" = True Then
            Rn = .Rows.count
          SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Reference Note"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                 cell_dict.add Title(5), "Shipping Mark"
               Remark_col.add cell_dict
                'NoteForSM_Rows.add cell_dict
             Next j
             GoTo 99
         End If
        End With
       Next i
       Set cell_dict = New Scripting.Dictionary
           SetRemarkTitle Title
            cell_dict.add Title(1), "Reference Note"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
              cell_dict.add Title(5), "Shipping Mark"
            Remark_col.add cell_dict
            'NoteForSM_Rows.add cell_dict
99       'NoteForSM_Dict.add "ReferenceNoteforShippingMark", NoteForSM_Rows
       
       
End Sub
Sub SAPForPackingPackaging(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
             
       
        Dim SAPForPP_Dict As Object
       Set SAPForPP_Dict = CreateObject("Scripting.Dictionary")
       Dim SAPForPP_Rows As Collection
       Set SAPForPP_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Special Attention Point for Packing / Packaging" = True Then
            Rn = .Rows.count
          SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Special Attention Point"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                cell_dict.add Title(5), "Packing / Packaging"
                Remark_col.add cell_dict
               'SAPForPP_Rows.add cell_dict
            Next j
            GoTo 108
         End If
        End With
       Next i
        Set cell_dict = New Scripting.Dictionary
         SetRemarkTitle Title
            cell_dict.add Title(1), "Special Attention Point"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
             cell_dict.add Title(5), "Packing / Packaging"
            Remark_col.add cell_dict

            'SAPForPP_Rows.add cell_dict
108        'SAPForPP_Dict.add "SpecialAttentionPointforPacking/Packaging", SAPForPP_Rows
       
         
    
  
   
End Sub
Sub NoteForPackingPackaging(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
                   
         Dim NoteforPP_Dict As Object
       Set NoteforPP_Dict = CreateObject("Scripting.Dictionary")
      Dim NoteforPP_Rows As Collection
      Set NoteforPP_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Reference Note for Packing / Packaging*" = True Then
            Rn = .Rows.count
           SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Reference Note"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
              
                 cell_dict.add Title(5), "Packing / Packaging"
               Remark_col.add cell_dict
               'NoteforPP_Rows.add cell_dict
            Next j
            GoTo 109
         End If
        End With
       Next i
       Set cell_dict = New Scripting.Dictionary
           SetRemarkTitle Title
            cell_dict.add Title(1), "Reference Note"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
             cell_dict.add Title(5), "Packing / Packaging"
            Remark_col.add cell_dict
            'NoteforPP_Rows.add cell_dict
109         ' NoteforPP_Dict.add "ReferenceNoteforPacking/Packaging", NoteforPP_Rows
       
       
End Sub
Sub SAPForSemiProduct(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
             
       
      Dim SAPForSemi_Dict As Object
       Set SAPForSemi_Dict = CreateObject("Scripting.Dictionary")
       Dim SAPForSemi_Rows As Collection
       Set SAPForSemi_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Special Attention Point for Semi-finished Product Inspection and Process Viewing*" = True Then
            Rn = .Rows.count
           SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Special Attention Point"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                cell_dict.add Title(5), "Semi-finished Product Inspection and Process Viewing"
                Remark_col.add cell_dict
               'SAPForSemi_Rows.add cell_dict
            Next j
            GoTo 118
         End If
        End With
       Next i
        Set cell_dict = New Scripting.Dictionary
            SetRemarkTitle Title
            cell_dict.add Title(1), "Special Attention Point"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
                
                cell_dict.add Title(5), "Semi-finished Product Inspection and Process Viewing"
                Remark_col.add cell_dict
           'SAPForSemi_Rows.add cell_dict
118       'SAPForSemi_Dict.add "SpecialAttentionPointforSemifinishedProductInspectionandProcessViewing", SAPForSemi_Rows
       
 
   
End Sub
Sub NoteForSemiProduct(Remark_col, ByRef Title() As String)
       
      Dim tbn, Rn, CLn, Cn, Itms As Integer
      Dim i, j, k As Integer
      Dim Kw  As String
      
      Dim cell_dict As Object
      Set WD = ThisDocument
                   
            
       Dim NoteForSemi_Dict As Object
       Set NoteForSemi_Dict = CreateObject("Scripting.Dictionary")
      Dim NoteForSemi_Rows As Collection
      Set NoteForSemi_Rows = New Collection
       tb = WD.Tables.count - 1
       For i = 1 To tb
    
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Reference Note for Semi-finished Product Inspection and Process Viewing*" = True Then
            Rn = .Rows.count
           SetRemarkTitle Title
            For j = 2 To Rn
              Set cell_dict = New Scripting.Dictionary
                cell_dict.add Title(1), "Reference Note"
                cell_dict.add Title(2), ""
              For k = 1 To 2
               If .Cell(j, k).Range.ContentControls.count > 0 Then
                 Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 1)
               Else
                Kw = Mid(Trim(.Cell(j, k).Range.Text), 1, Len(Trim(.Cell(j, k).Range.Text)) - 2)
               End If
                cell_dict.add Title(k + 2), Kw
              Next k
                cell_dict.add Title(5), "Semi-finished Product Inspection and Process Viewing"
               Remark_col.add cell_dict
               'NoteForSemi_Rows.add cell_dict
            Next j
            GoTo 119
         End If
        End With
       Next i
       Set cell_dict = New Scripting.Dictionary
         SetRemarkTitle Title
            cell_dict.add Title(1), "Reference Note"
            cell_dict.add Title(2), ""
            cell_dict.add Title(3), ""
            cell_dict.add Title(4), ""
             cell_dict.add Title(5), "Semi-finished Product Inspection and Process Viewing"
            Remark_col.add cell_dict
            'NoteForSemi_Rows.add cell_dict
119        ' NoteForSemi_Dict.add "ReferenceNoteforSemifinishedProductInspectionandProcessViewing", NoteForSemi_Rows
      
  
       
End Sub
Sub Export1545ResultToHUB()
'
' by wuyannan ,20241122,collection,dictionary，revised at 20241217
     
   Dim PackageUnit, SamplingStandard, SamplingPlan, InspectionLevel, CriticalAQL, MajorAQL, MinorAQL As String

     Dim CriticalSampleSize, MajorSampleSize, MinorSamleSize As String
     Dim CLn As Integer
     Dim TbnB, TbnE, i, j, k, Rn, Cn, Celn, ST, ED As Integer
     Dim tb_tt(0 To 14), BK(0 To 14), Kw, VW, VWL, VWR As String
     Dim rg As Range
     Dim a As String
     Dim tb, tbn As Integer
     Dim Ln As Integer
     
     Dim ReportID, ClientID, SupplierID, FactoryID As String
     Dim Client, Supplier, Factory As String
     Dim Unit As String
     
     Dim Title() As String
     
    Dim CItemIdx, CSpeIdx, CToLIdx, CActIdx, CChkIdx, CMeaIdx, CCpnIdx, CCDFIdx As Integer
    Dim RItemIdx, RSpeIdx, RToLIdx, RActIdx, RChkIdx, RMeaIdx, RCpnIdx, RCDFIdx As Integer
    Dim StRow, StColumn As Integer
    Dim TableType As String
 
    Dim KItmNo As String
    Dim ChkPoints As String
    Dim ChkG(200) As Integer
    Dim TemChkPoint As String
    Dim KwC, KwS As String
    Dim DataTableG(50) As String
    Dim DTN, Tn, Dn, DR As Integer
     VerifyReportID
     
     If EXT = "Quit" Then Exit Sub
     
     Set WD = ThisDocument
    ' Set wd = ActiveDocument
     tb_tt(0) = "INSPECTION INFORMATION"
     tb_tt(1) = "1. Quantity"
     tb_tt(2) = "2. Workmanship"
     tb_tt(3) = "3. On Site Test"
     tb_tt(4) = "4. Product Dimension & Weight"
     tb_tt(5) = "5. Style / Material / Construction"
     tb_tt(6) = "6. Product Color"
     tb_tt(7) = "7. Product Label / Marking"
     tb_tt(8) = "8. Shipping Mark"
     tb_tt(9) = "9. Packing / Packaging"
      '--------------------PSI，DPI
     tb_tt(10) = "10. Semi-finished product Inspection and Process Viewing"
     '--------------------
     
     tb_tt(11) = "10. Other Note"
     tb_tt(12) = "Other Photos"
     tb_tt(13) = "INSPECTION RESULT SUMMARY"
     tb_tt(14) = "OVERALL CONCLUSION"
    '-----------------------------
     Application.StatusBar = "Packing report data and uploading data to HUB, please wait......" & "" & _
     UText(25171, 21253, 25968, 25454, 24182, 19978, 20256) & "HUB" & UText(20013) & "，" & UText(35831, 31245, 20505) & "......"
     
     
     ReportID = GetPropertyValue("ReportID")
     ReportID = Replace(ReportID, "ReportID_", "")
     Set rg = WD.Sections(1).Headers(wdHeaderFooterFirstPage).Range
     i = rg.Tables(1).Range.Cells.count
     ReportNo = Trim(rg.Tables(1).Range.Cells(i).Range.Text)
     ReportNo = Mid(ReportNo, 1, Len(ReportNo) - 2)
     ReportNo = Replace(ReportNo, "Report No.:", "")
     ReportNo = Replace(ReportNo, "Report No:", "")
    
  
     Set rg = WD.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Range
     Pcn = rg.Cells.count
     rpn = rg.Cells(Pcn).Range.Text
     rpn = Mid(rpn, 1, Len(rpn) - 2)
     rpn = Trim(Replace(rpn, "Report No: ", ""))
     
     Dim Tc As Integer
     Dim Tr, Trn As String
     Set rg = WD.Sections(1).Footers(wdHeaderFooterFirstPage).Range
     Tc = rg.Tables.count
     Tr = Mid(rg.Tables(Tc).Cell(1, 1).Range.Text, 1, Len(rg.Tables(Tc).Cell(1, 1).Range.Text) - 2)
     If InStr(Tr, "-") = 0 Then
        Tr = "N/A"
     Else
       Do While Val(Tr) = 0
         Tr = Mid(Tr, InStr(Tr, "-") + 1)
        Loop
         Trn = Tr
     End If
  
  
     Dim Issued_Date As String
     Dim Inspector, Auditor As String
     
     Issued_Date = rg.Tables(1).Cell(1, 4).Range.Text
     Issued_Date = Mid(Issued_Date, 1, Len(Issued_Date) - 2)
    
     tb = WD.Tables.count - 1
     Inspector = WD.Tables(tb - 1).Cell(1, 2).Range.Text
     Inspector = Mid(Inspector, 1, Len(Inspector) - 2)
     Auditor = WD.Tables(tb - 1).Cell(1, 4).Range.Text
     Auditor = Mid(Auditor, 1, Len(Auditor) - 2)
     With WD.Tables(1).Range
     Cn = .Cells.count
     For i = 2 To Cn - 1 Step 2
        If Mid(.Cells(i).Range.Text, 1, Len(.Cells(i).Range.Text) - 2) = "Client" Then
        If .Cells(i + 1).Range.Bookmarks.count > 0 Then
          ClientID = .Cells(i + 1).Range.Bookmarks(1).Name
          ClientID = Replace(ClientID, "ClientID_", "")
        Else
          ClientID = GetPropertyValue("ClientID ")
        End If
          Client = Mid(.Cells(i + 1).Range.Text, 1, Len(.Cells(i + 1).Range.Text) - 2)
          GoTo 5
        End If
        If Mid(.Cells(i).Range.Text, 1, Len(.Cells(i).Range.Text) - 2) = "Supplier" Then
         If .Cells(i + 1).Range.Bookmarks.count > 0 Then
           SupplierID = .Cells(i + 1).Range.Bookmarks(1).Name
           SupplierID = Replace(SupplierID, "SupplierID_", "")
        Else
           SupplierID = "N/A"
        End If
           Supplier = Mid(.Cells(i + 1).Range.Text, 1, Len(.Cells(i + 1).Range.Text) - 2)
           GoTo 5
        End If
        If Mid(.Cells(i).Range.Text, 1, Len(.Cells(i).Range.Text) - 2) = "Factory" Then
           Factory = Mid(.Cells(i + 1).Range.Text, 1, Len(.Cells(i + 1).Range.Text) - 2)
           If .Cells(i + 1).Range.Bookmarks.count > 0 Then
              FactoryID = .Cells(i + 1).Range.Bookmarks(1).Name
              FactoryID = Replace(FactoryID, "FactoryID_", "")
              FactoryID = Replace(FactoryID, "_", ",")
           Else
              FactoryID = "N/A"
           End If
          GoTo 5
        End If
       
5   Next i
    End With
    
     'BK(0), "INSPECTION INFORMATION"
    Dim Info_Dict As Object
    Set Info_Dict = CreateObject("Scripting.Dictionary")
    Info_Dict.add "ReportTemplateNo", Trn
    Info_Dict.add "ReportID", ReportID
    Info_Dict.add "ReportNo", ReportNo
    Info_Dict.add "ReportIssuedDate", Issued_Date
    Info_Dict.add "Inspector", Inspector
    Info_Dict.add "Auditor", Auditor
    
    With WD.Tables(1)
      Celn = .Range.Cells.count
      For i = 2 To Celn - 1 Step 2
         Kw = Trim(.Range.Cells(i).Range.Text)
         Kw = Mid(Kw, 1, Len(Kw) - 2)
         Kw = Replace(Kw, " ", "")
         VW = Trim(.Range.Cells(i + 1).Range.Text)
         If .Range.Cells(i + 1).Range.ContentControls.count <> 0 Then
            VW = Mid(VW, 1, Len(VW) - 1)
         Else
           VW = Mid(VW, 1, Len(VW) - 2)
         End If
         If Kw = "Client" Then
           Info_Dict.add "ClientID", ClientID
         ElseIf Kw = "Supplier" Then
           Info_Dict.add "SupplierID", SupplierID
         ElseIf Kw = "Factory" Then
          Info_Dict.add "FactoryID", FactoryID
         ElseIf Kw = "InspectionDate" Then

           Ln = Len(VW)
           If Ln > 15 Then
             VWL = Mid(VW, 1, Int(Ln / 2))
             VWR = Mid(VW, Int(Ln / 2) + 1, Ln - Int(Ln / 2))
           Else
             VWL = VW
             VWR = VW
           End If
           Info_Dict.add "InspectionStartDate", VWL
           Info_Dict.add "InspectionEndDate", VWR

        ElseIf Kw = "Location" Then
            If InStr(1, VW, ",") > 0 Then
               VWL = Trim(Mid(VW, 1, InStr(1, VW, ",") - 1))
               VWR = Trim(Mid(VW, InStr(1, VW, ",") + 1, Len(VW) - InStr(1, VW, ",")))
            Else
               VWL = VW
               VWR = ""
            End If
            Info_Dict.add "Location-City", VWL
            Info_Dict.add "Location-Country", VWR
         End If
         If InStr(1, Kw, ".") > 0 Then Kw = Replace(Kw, ".", "")
         
         If Kw = "Quantity" Then
            Info_Dict.add "Quantity", Mid(VW, 1, InStr(1, VW, " ") - 1)
            Info_Dict.add "Unit", Mid(VW, InStr(1, VW, " ") + 1, Len(VW) - InStr(1, VW, " "))
         Else
            Info_Dict.add Kw, VW
         End If
         
       
10     Next i
    End With
         
         
    'Json_string = jsonConverter.ConvertToJson(Report_Dict)
    'Debug.Print Json_string
     Dim InspectionType As String
     Dim InsRes_Sum_Rows As Collection
     Set InsRes_Sum_Rows = New Collection
     With WD.Tables(3)
           Rn = .Rows.count
           CLn = .Columns.count
           Cn = 0
          Dim CTitle() As String
          ReDim CTitle(1 To CLn + 1)
          For j = 1 To CLn
            CTitle(j) = Mid(.Cell(2, j).Range.Text, 1, Len(.Cell(2, j).Range.Text) - 2)
            CTitle(j) = Replace(CTitle(j), " ", "")
          Next j
           CTitle(CLn + 1) = "DetailTableResult"
           
           Dim InsRes_Sum() As String
           ReDim InsRes_Sum(0 To Rn - 1, CLn + 1)
           Dim cell_dict As Object
           
            With WD.Tables(1)
              Celn = .Range.Cells.count
              For i = 2 To Celn - 1 Step 2
                 Set Cel = .Range.Cells(i)
                 Set cel1 = .Range.Cells(i + 1)
                 If Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2) = "Inspection Type" Then
                    InspectionType = Mid(cel1.Range.Text, 1, Len(cel1.Range.Text) - 2)
                    GoTo 15
                  End If
              Next i
            End With
15          If InspectionType = "Pre-shipment Inspection" Then
                Rn = 12
              ElseIf InspectionType = "During Production Inspection" Then
                Rn = 13
             End If
         
            For i = 2 To Rn - 1
              Set cell_dict = New Scripting.Dictionary
              For j = 1 To CLn
                Kw = Trim(.Cell(i, j).Range.Text)
                If .Cell(i, j).Range.ContentControls.count > 0 Then
                  Kw = Mid(Kw, 1, Len(Kw) - 1)
                Else
                  Kw = Mid(Kw, 1, Len(Kw) - 2)
                End If
                  If j = 1 Then Kw = Mid(Kw, InStr(1, Kw, " ") + 1, Len(Kw) - InStr(1, Kw, " "))
                  InsRes_Sum(Cn, j) = Kw
                  cell_dict.add CTitle(j), InsRes_Sum(Cn, j)
              Next j
           
              
             If Cn = 0 Then
                 cell_dict.add "DetailTableResult", "Result of detail table"
             Else
       
                 If WD.Tables(getTableNumber(Cn)).Cell(1, 2).Range.ContentControls.count > 0 Then
                      Kw = Mid(WD.Tables(getTableNumber(Cn)).Cell(1, 2).Range.Text, 1, Len(WD.Tables(getTableNumber(Cn)).Cell(1, 2).Range.Text) - 1)
                 Else
                      Kw = Mid(WD.Tables(getTableNumber(Cn)).Cell(1, 2).Range.Text, 1, Len(WD.Tables(getTableNumber(Cn)).Cell(1, 2).Range.Text) - 2)
                 End If
                
                 cell_dict.add "DetailTableResult", Kw
              End If
                Cn = Cn + 1
                InsRes_Sum_Rows.add cell_dict
           Next i
             
             Set cell_dict = New Scripting.Dictionary
             If InspectionType = "Pre-shipment Inspection" Then
                 cell_dict.add CTitle(1), "SemifinisheproductInspectionandProcessViewing"
                For j = 2 To CLn
                   Kw = ""
                   cell_dict.add CTitle(j), Kw
                Next j
                   cell_dict.add "DetailTableResult", ""
                   InsRes_Sum_Rows.add cell_dict
             End If
                
        
           
           Info_Dict.add "OVERALLCONCLUSION", Trim(Mid(.Cell(Rn, 2).Range.Text, 1, Len(.Cell(Rn, 2).Range.Text) - 2))
           
           Dim InsRes_Sum_Dict As Object
           Set InsRes_Sum_Dict = CreateObject("Scripting.Dictionary")
           
           'InsRes_Sum_Dict.add "InsRes_Sum", InsRes_Sum_Rows
           
           'InsRes_Sum_Dict.add "OVERALL CONCLUSION", Trim(Mid(.cell(Rn, 2).Range.Text, 1, Len(.cell(Rn, 2).Range.Text) - 2))
      
           'json_String = jsonConverter.ConvertToJson(InsRes_Sum_Dict)
            'Debug.Print json_String
      
      
    End With
       

       
    ' json_String = JsonConverter.ConvertToJson(Info_dict)
     ' Debug.Print json_String
      Dim Remark_col As Collection
      Set Remark_col = New Collection
       
       ' ---BK(1) , "Quantity"
      tbn = 0
      tbn = getTableNumber(1)
      If tbn = 0 Then
         MsgBox tb_tt(1) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(1) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
20    Dim Qty_Dict As Object
      Set Qty_Dict = CreateObject("Scripting.Dictionary")
       With WD.Tables(tbn)
          Rn = .Rows.count
          CLn = .Columns.count
          Itms = Rn - 5
          Cn = 0
          Kw = Trim(.Cell(4, 4).Range.Text)
          Unit = Mid(Kw, InStr(1, Kw, "(") + 1, InStr(1, Kw, ")") - 2)
          Kw = .Cell(4, 5).Range.Text
          PackageUnit = Mid(Kw, InStr(1, Kw, "(") + 1, InStr(1, Kw, ")") - 2)
          '-------------------------------------ST----

    ' Dim Title() As String
     ReDim Title(0 To CLn)
     
     Title(1) = Mid(.Cell(3, 1).Range.Text, 1, Len(.Cell(3, 1).Range.Text) - 2)
     Title(2) = Mid(.Cell(3, 2).Range.Text, 1, Len(.Cell(3, 2).Range.Text) - 2)
     Title(3) = Mid(.Cell(3, 3).Range.Text, 1, Len(.Cell(3, 3).Range.Text) - 2)
     Title(4) = Mid(.Cell(3, 4).Range.Text, 1, Len(.Cell(3, 4).Range.Text) - 2) & "(Pieces)"
     Title(5) = Mid(.Cell(3, 4).Range.Text, 1, Len(.Cell(3, 4).Range.Text) - 2) & "(Cartons)"
     Title(6) = Mid(.Cell(3, 5).Range.Text, 1, Len(.Cell(3, 5).Range.Text) - 2) & "(Pieces)"
     Title(7) = Mid(.Cell(3, 5).Range.Text, 1, Len(.Cell(3, 5).Range.Text) - 2) & "(Cartons)"
     Title(8) = Mid(.Cell(3, 6).Range.Text, 1, Len(.Cell(3, 6).Range.Text) - 2) & "(Pieces)"
     Title(9) = Mid(.Cell(3, 6).Range.Text, 1, Len(.Cell(3, 6).Range.Text) - 2) & "(Cartons)"
     
     
     Dim Qty_Rows As New Collection
     
     Dim PONo_G(), ItemNo_G() As String
     ReDim PONo_G(1 To Itms), ItemNo_G(1 To Itms)
          '--------------
          For i = 5 To Rn - 1
            Cn = Cn + 1
            Set cell_dict = New Scripting.Dictionary
              For j = 1 To CLn
                  Title(j) = Replace(Title(j), " ", "")
                  Title(j) = Replace(Title(j), ".", "")
                 If j <= 2 Then
                     cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
                 Else
                     If .Cell(4, j).Range.ContentControls.count > 0 Then
                        Kw = Mid(.Cell(4, j).Range.Text, 1, Len(.Cell(4, j).Range.Text) - 1)
                      Else
                        Kw = Mid(.Cell(4, j).Range.Text, 1, Len(.Cell(4, j).Range.Text) - 2)
                      End If
                      Unit = Mid(Kw, InStr(1, Kw, "(") + 1, InStr(1, Kw, ")") - 2)
                     cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2) & " " & Unit & "s"
                 End If
                 
              Next j
              Qty_Rows.add cell_dict
        
         Next i
           
         
          'Qty_Dict.add "Quantity Details", Qty_Rows
          Dim TotalQty_Dict As Object
          Set TotalQty_Dict = New Scripting.Dictionary
         
           'Set cell_dict = New Scripting.Dictionary
           For j = 2 To CLn - 1
                If .Cell(4, j + 1).Range.ContentControls.count > 0 Then
                        Kw = Mid(.Cell(4, j + 1).Range.Text, 1, Len(.Cell(4, j + 1).Range.Text) - 1)
                      Else
                        Kw = Mid(.Cell(4, j + 1).Range.Text, 1, Len(.Cell(4, j + 1).Range.Text) - 2)
                      End If
                      Unit = Mid(Kw, InStr(1, Kw, "(") + 1, InStr(1, Kw, ")") - 2)
                 TotalQty_Dict.add Title(j + 1), Mid(.Cell(Rn, j).Range.Text, 1, Len(.Cell(Rn, j).Range.Text) - 2) & " " & Unit & "s"
           Next j
         
     
       End With
       '--------------------------
      ' json_String = jsonConverter.ConvertToJson(Qty_Dict)
       'Debug.Print json_String
       
        Call SAPForQuantity(Remark_col, Title)
        Call NoteForQuantity(Remark_col, Title)
        
        'json_String = JsonConverter.ConvertToJson(Remark_col, Title)
       ' Debug.Print json_String
      
      '  BK(2) , "Workmanship"
       
      tbn = 0
      tbn = getTableNumber(2)
      If tbn = 0 Then
         MsgBox tb_tt(2) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(2) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
        
30   Dim Workmanship_Dict As Object
     Set Workmanship_Dict = CreateObject("Scripting.Dictionary")
     
     With WD.Tables(tbn)
          If .Cell(1, 2).Range.ContentControls.count > 0 Then
            Kw = Mid(Trim(.Cell(1, 2).Range.Text), 1, Len(Trim(.Cell(1, 2).Range.Text)) - 1)
          Else
             Kw = Mid(Trim(.Cell(1, 2).Range.Text), 1, Len(Trim(.Cell(1, 2).Range.Text)) - 2)
          End If
          'Workmanship_Dict.add "2. Workmanship Result", Kw
          Kw = Trim(.Cell(3, 2).Range.Text)
          If .Cell(3, 2).Range.ContentControls.count <> 0 Then
             Kw = Mid(Kw, 1, Len(Kw) - 1)
          Else
             Kw = Mid(Kw, 1, Len(Kw) - 2)
          End If
          SamplingStandard = Kw
          Kw = Trim(.Cell(4, 2).Range.Text)
          If .Cell(4, 2).Range.ContentControls.count <> 0 Then
             Kw = Mid(Kw, 1, Len(Kw) - 1)
          Else
             Kw = Mid(Kw, 1, Len(Kw) - 2)
          End If
          SamplingPlan = Kw
          Kw = Trim(.Cell(5, 2).Range.Text)
          If .Cell(4, 2).Range.ContentControls.count <> 0 Then
              Kw = Mid(Kw, 1, Len(Kw) - 1)
          Else
              Kw = Mid(Kw, 1, Len(Kw) - 2)
          End If
          InspectionLevel = Kw
          If .Cell(4, 4).Range.ContentControls.count > 0 Then
             CriticalAQL = Mid(Trim(.Cell(4, 4).Range.Text), 1, Len(Trim(.Cell(4, 4).Range.Text)) - 1)
          Else
             CriticalAQL = Mid(Trim(.Cell(4, 4).Range.Text), 1, Len(Trim(.Cell(4, 4).Range.Text)) - 2)
          End If
          If CriticalAQL = "" Then CriticalAQL = "0"
          If .Cell(4, 5).Range.ContentControls.count > 0 Then
               MajorAQL = Mid(Trim(.Cell(4, 5).Range.Text), 1, Len(Trim(.Cell(4, 5).Range.Text)) - 1)
          Else
               MajorAQL = Mid(Trim(.Cell(4, 5).Range.Text), 1, Len(Trim(.Cell(4, 5).Range.Text)) - 2)
          End If
          If MajorAQL = "" Then MajorAQL = "0"
     
          If .Cell(4, 6).Range.ContentControls.count > 0 Then
                MinorAQL = Mid(Trim(.Cell(4, 6).Range.Text), 1, Len(Trim(.Cell(4, 6).Range.Text)) - 1)
          Else
                MinorAQL = Mid(Trim(.Cell(4, 6).Range.Text), 1, Len(Trim(.Cell(4, 6).Range.Text)) - 2)
          End If
          If MinorAQL = "" Then MajorAQL = "0"
       
          CriticalSampleSize = CStr(Val(.Cell(5, 4).Range.Text))
          MajorSampleSize = CStr(Val(.Cell(5, 5).Range.Text))
          MinorSamleSize = CStr(Val(.Cell(5, 6).Range.Text))
     
          Info_Dict.add "SamplingStandard", SamplingStandard
          Info_Dict.add "SamplingPlan", SamplingPlan
          Info_Dict.add "InspectionLevel", InspectionLevel
          Info_Dict.add "AQL-Critical", CriticalAQL
          Info_Dict.add "AQL-Major", MajorAQL
          Info_Dict.add "AQL-Minor", MinorAQL
          Info_Dict.add "SampleSize-Critical", CriticalSampleSize
          Info_Dict.add "SampleSize-Major", MajorSampleSize
          Info_Dict.add "SampleSize-Minor", MinorSamleSize
     
          SetCheckPointTitle Title
     
          Dim Workmanship_Rows As New Collection
     
          Rn = .Rows.count
          For i = 7 To Rn
             Set cell_dict = New Scripting.Dictionary
             cell_dict.add "Category", "Workmanship"
             If .Cell(i, 2).Range.Bookmarks.count > 0 Then
                cell_dict.add Title(0), .Cell(i, 2).Range.Bookmarks(1).Name
             Else
                cell_dict.add Title(0), "N/A"
             End If
       
             For j = 1 To 3
                cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
             Next j
             cell_dict.add Title(4), ""
             cell_dict.add Title(5), ""
             Workmanship_Rows.add cell_dict
          Next i
          'Workmanship_Dict.add "WorkmanshipCheckPoint", Workmanship_Rows
    End With
   
     '------------Defect-------------
     
     Dim Defect_Dict As Object
     Dim ItemTotal_Dict As Object
     Dim AllTotal_Dict As Object
     Dim Item_Dict As Object
     Dim Defects_rows As New Collection
     Dim TotalFoundCritical, TotalFoundMajor, TotalFoundMinor As Integer
     Dim TotalAllowedCritical, TotalAllowedMajor, TotalAllowedMinor As Integer
     
     With WD.Tables(tbn + 1)
        Rn = .Rows.count
        CLn = .Columns.count
     
        If CLn = 5 Then Title(1) = "DefectCode"
        Title(CLn - 3) = "Defects"
        Title(CLn - 2) = "Critical"
        Title(CLn - 1) = "Major"
        Title(CLn) = "Minor"
   
    
        Cn = 0
        Dim SamplingType As Integer
        If LCase(InspectionLevel) Like "*per item*" = False Then
           SamplingType = 1
        Else
           SamplingType = 2
        End If
     
        Dim Item_Col As Collection
        Set Item_Col = New Collection
     
        Select Case SamplingType
        Case 1
       
           For i = 2 To Rn - 2
              Set cell_dict = New Scripting.Dictionary
             If .Rows(i).Range.Cells.count = 1 Then
                 Set Item_Dict = New Scripting.Dictionary
                 Set Defects_rows = New Collection
                 Set ItemTotal_Dict = New Scripting.Dictionary
          
                 Kw = Mid(.Cell(i, 1).Range.Text, 1, Len(.Cell(i, 1).Range.Text) - 2)
                 Kw = Trim(Replace(Kw, "Item No.", ""))
                 If InStr(1, Kw, ",") > 0 Then
                   Kw = Trim(Left(Kw, InStr(1, Kw, ",") - 1))
                 End If
                 Item_Dict.add "ItemNo", Kw
                 Kw = Mid(.Cell(i, 1).Range.Text, 1, Len(.Cell(i, 1).Range.Text) - 2)
                 Kw = Mid(Kw, InStr(1, Kw, "=") + 1, Len(Kw) - InStr(1, Kw, "="))
                 Item_Dict.add "SampleSize", Kw
                'Defects_rows.add Item_Dict
                 TotalFoundCritical = 0
                 TotalFoundMajor = 0
                 TotalFoundMinor = 0
            Else
            
                 If .Cell(i, 2).Range.Bookmarks.count > 0 Then
                     Kw = .Cell(i, 2).Range.Bookmarks(1).Name
                     If InStr(1, Kw, "_") > 0 Then
                        Kw = Mid(Kw, 1, InStr(1, Kw, "_") - 1)
                     End If
                     cell_dict.add Title(1), Kw
                 Else
                     cell_dict.add Title(1), "N/A"
                 End If
                 For j = CLn - 3 To CLn
                    cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
                 Next j
                 Defects_rows.add cell_dict
            
                 TotalFoundCritical = TotalFoundCritical + Val(.Cell(i, CLn - 2).Range.Text)
                 TotalFoundMajor = TotalFoundMajor + Val(.Cell(i, CLn - 1).Range.Text)
                 TotalFoundMinor = TotalFoundMinor + Val(.Cell(i, CLn).Range.Text)
            
                 If .Rows(i + 1).Range.Cells.count = 1 Then
                    Item_Dict.add "DefectDetail", Defects_rows
                    ItemTotal_Dict.add "TotalFoundCritical", CStr(TotalFoundCritical)
                    ItemTotal_Dict.add "TotalFoundMajor", CStr(TotalFoundMajor)
                    ItemTotal_Dict.add "TotalFoundMinor", CStr(TotalFoundMinor)
              
                    ItemTotal_Dict.add "TotalAllowedCritical", "N/A"
                    ItemTotal_Dict.add "TotalAllowedMajor", "N/A"
                    ItemTotal_Dict.add "TotalAllowedMinor", "N/A"
                    'Defects_rows.add TotalAllowed_Dict
                    Item_Dict.add "ItemTotal", ItemTotal_Dict
                    Item_Col.add Item_Dict
               Else
                    If .Cell(i + 1, CLn - 3).Range.Fields.count > 0 Then
                      Item_Dict.add "DefectDetail", Defects_rows
                      ItemTotal_Dict.add "TotalFoundCritical", CStr(TotalFoundCritical)
                      ItemTotal_Dict.add "TotalFoundMajor", CStr(TotalFoundMajor)
                      ItemTotal_Dict.add "TotalFoundMinor", CStr(TotalFoundMinor)
                      'Defects_rows.add TotalFound_Dict
               
                      ItemTotal_Dict.add "TotalAllowedCritical", "N/A"
                      ItemTotal_Dict.add "TotalAllowedMajor", "N/A"
                      ItemTotal_Dict.add "TotalAllowedMinor", "N/A"
                     'Defects_rows.add TotalAllowed_Dict
                      Item_Dict.add "ItemTotal", ItemTotal_Dict
                      Item_Col.add Item_Dict
                   End If
               End If
            End If
        
33       Next i
         AllTotalFoundCritical = Val(.Cell(Rn - 1, CLn - 2).Range.Text)
         AllTotalFoundMajor = Val(.Cell(Rn - 1, CLn - 1).Range.Text)
         AllTotalFoundMinor = Val(.Cell(Rn - 1, CLn).Range.Text)
         AllTotalAllowedCritical = Val(.Cell(Rn, CLn - 2).Range.Text)
         AllTotalAllowedMajor = Val(.Cell(Rn, CLn - 1).Range.Text)
         AllTotalAllowedMinor = Val(.Cell(Rn, CLn).Range.Text)
        
      Case 2
         For i = 2 To Rn
            Set cell_dict = New Scripting.Dictionary
            If .Rows(i).Range.Cells.count = 1 Then
               Set Item_Dict = New Scripting.Dictionary
               Set Defects_rows = New Collection
               Set ItemTotal_Dict = New Scripting.Dictionary
               Kw = Mid(.Cell(i, 1).Range.Text, 1, Len(.Cell(i, 1).Range.Text) - 2)
               Kw = Trim(Replace(Kw, "Item No.", ""))
               If InStr(1, Kw, ",") > 0 Then
                  Kw = Trim(Left(Kw, InStr(1, Kw, ",") - 1))
               End If
               Item_Dict.add "ItemNo", Kw
               Kw = Mid(.Cell(i, 1).Range.Text, 1, Len(.Cell(i, 1).Range.Text) - 2)
               Kw = Mid(Kw, InStr(1, Kw, "=") + 1, Len(Kw) - InStr(1, Kw, "="))
               Item_Dict.add "SampleSize", Kw
              ' Defects_rows.add Item_Dict
            Else
               If .Cell(i, CLn - 3).Range.Fields.count = 0 Then
                  If .Cell(i, 2).Range.Bookmarks.count > 0 Then
                     Kw = .Cell(i, 2).Range.Bookmarks(1).Name
                     If InStr(1, Kw, "_") > 0 Then
                        Kw = Mid(Kw, 1, InStr(1, Kw, "_") - 1)
                     End If
                     cell_dict.add Title(1), Kw
                  Else
                     cell_dict.add Title(1), "N/A"
                  End If
                  For j = CLn - 3 To CLn
                    cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
                  Next j
                 Defects_rows.add cell_dict
              Else
                 Item_Dict.add "DefectDetail", Defects_rows
                 TotalFoundCritical = Val(.Cell(i, CLn - 2).Range.Text)
                 TotalFoundMajor = Val(.Cell(i, CLn - 1).Range.Text)
                 TotalFoundMinor = Val(.Cell(i, CLn).Range.Text)
                 ItemTotal_Dict.add "TotalFoundCritical", CStr(TotalFoundCritical)
                 ItemTotal_Dict.add "TotalFoundMajor", CStr(TotalFoundMajor)
                 ItemTotal_Dict.add "TotalFoundMinor", CStr(TotalFoundMinor)
                'Defects_rows.add TotalFound_Dict
              
                 AllTotalFoundCritical = AllTotalFoundCritical + TotalFoundCritical
                 AllTotalFoundMajor = AllTotalFoundMajor + TotalFoundMajor
                 AllTotalFoundMinor = AllTotalFoundMinor + TotalFoundMinor
                   
                 TotalAllowedCritical = Val(.Cell(i + 1, CLn - 2).Range.Text)
                 TotalAllowedMajor = Val(.Cell(i + 1, CLn - 1).Range.Text)
                 TotalAllowedMinor = Val(.Cell(i + 1, CLn).Range.Text)

                 ItemTotal_Dict.add "TotalAllowedCritical", CStr(TotalAllowedCritical)
                 ItemTotal_Dict.add "TotalAllowedMajor", CStr(TotalAllowedMajor)
                 ItemTotal_Dict.add "TotalAllowedMinor", CStr(TotalAllowedMinor)
                 'Defects_rows.add TotalAllowed_Dict
                 Item_Dict.add "ItemTotal", ItemTotal_Dict
               
                 AllTotalAllowedCritical = AllTotalAllowedCritical + TotalAllowedCritical
                 AllTotalAllowedMajor = AllTotalAllowedMajor + TotalAllowedMajor
                 AllTotalAllowedMinor = AllTotalAllowedMinor + TotalAllowedMinor
                 Item_Col.add Item_Dict
                   i = i + 1
               End If
            End If
35       Next i
      End Select
    
     'Set AllTotal_Dict = New Scripting.Dictionary
      Info_Dict.add "AllTotalFoundCritical", CStr(AllTotalFoundCritical)
      Info_Dict.add "AllTotalFoundMajor", CStr(AllTotalFoundMajor)
      Info_Dict.add "AllTotalFoundMinor", CStr(AllTotalFoundMinor)
      'Defects_rows.add AllTotalFound_Dict
       
        
      Info_Dict.add "AllTotalAllowedCritical", CStr(AllTotalAllowedCritical)
      Info_Dict.add "AllTotalAllowedMajor", CStr(AllTotalAllowedMajor)
      Info_Dict.add "AllTotalAllowedMinor", CStr(AllTotalAllowedMinor)
         
       'Info_Dict.add "AllTotal", AllTotal_Dict
        
      ' Defects_rows.add AllTotalAllowed_Dict
         
      'Set Defect_Dict = New Scripting.Dictionary
      'Item_Dict.add "DefectDetail", Defects_rows
      'Workmanship_Dict.add "Defect", Item_Col
        
         '---------------------------------------------------------Ed--
     End With
     
      ' json_String = JsonConverter.ConvertToJson(Defect_Dict)
        'Debug.Print json_String
        'wd.Tables(2).cell(1, 1).Range.Text = json_String
  
      Call SAPForWorkmanship(Remark_col, Title)
      Call NoteForWorkmanship(Remark_col, Title)
       '---------------------------
        
        'Json_string = jsonConverter.ConvertToJson(Defect_Dict)
        'Debug.Print Json_string
     
     ' BK (3), "OnSiteTest"
      tbn = 0
      tbn = getTableNumber(3)
      If tbn = 0 Then
         MsgBox tb_tt(3) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(3) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
      TbnE = 0
      TbnE = getTableNumber(4)
       
      If TbnE = 0 Then
         MsgBox tb_tt(4) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(4) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
     
     
40    With WD.Tables(tbn)
         Rn = .Rows.count
         CLn = .Columns.count
         Cn = 0
         Dim OST_Dict As Object
         Set OST_Dict = CreateObject("Scripting.Dictionary")
         If .Cell(1, 2).Range.ContentControls.count > 0 Then
              Kw = Mid(Trim(.Cell(1, 2).Range.Text), 1, Len(Trim(.Cell(1, 2).Range.Text)) - 1)
         Else
              Kw = Mid(Trim(.Cell(1, 2).Range.Text), 1, Len(Trim(.Cell(1, 2).Range.Text)) - 2)
         End If
         ' OST_Dict.add "3. On Site Test Result", Kw
     
         SetCheckPointTitle Title
         Dim OST_Rows As New Collection
   
         For i = 4 To Rn
            Set cell_dict = New Scripting.Dictionary
            cell_dict.add "Category", "On Site Test"
            If .Cell(i, 2).Range.Bookmarks.count > 0 Then
                cell_dict.add Title(0), .Cell(i, 2).Range.Bookmarks(1).Name
             Else
                cell_dict.add Title(0), "N/A"
             End If

             For j = 1 To CLn
                If .Cell(i, j).Range.ContentControls.count > 0 Then
                   cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
                Else
                   cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
                End If
              Next j
              OST_Rows.add cell_dict
         Next i
         'OST_Dict.add "OnSiteTest CheckPoint", OST_Rows
      
       End With
        'json_String = JsonConverter.ConvertToJson(OST_Dict)
         'Debug.Print json_String
        'wd.Tables(2).cell(1, 1).Range.Text = json_String
       
       '--------------------------
       Call SAPForOnSiteTest(Remark_col, Title)
       Call NoteForOnSiteTest(Remark_col, Title)
       '-------------------------------
  
49       '---------------------------
      
      
      ' BK(4) , "Product Dimension & Weight"
      tbn = 0
      tbn = getTableNumber(4)
      If tbn = 0 Then
         MsgBox tb_tt(4) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(4) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
     
     
50  With WD.Tables(tbn)
       Rn = .Rows.count
       CLn = .Columns.count
       Cn = 0
      
       Dim DW_Dict As Object
       Set DW_Dict = CreateObject("Scripting.Dictionary")
       If .Cell(1, 2).Range.ContentControls.count > 0 Then
            Kw = Trim(Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1))
       Else
             Kw = Trim(Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 2))
       End If
          'DW_Dict.add "4. Product Dimension & Weight Result", Kw
       SetCheckPointTitle Title

       Dim DW_Rows As New Collection
       For i = 4 To Rn
          Set cell_dict = New Scripting.Dictionary
          cell_dict.add "Category", "Product Dimension & Weight"
          If .Cell(i, 2).Range.Bookmarks.count > 0 Then
               cell_dict.add Title(0), .Cell(i, 2).Range.Bookmarks(1).Name
          Else
              cell_dict.add Title(0), "N/A"
          End If
       
          For j = 1 To 2
             If .Cell(i, j).Range.ContentControls.count > 0 Then
               cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
             Else
               cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
            End If
          Next j
          cell_dict.add Title(3), ""
       
          For j = 3 To CLn
             If .Cell(i, j).Range.ContentControls.count > 0 Then
               cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
             Else
               cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
             End If
          Next j
       
          DW_Rows.add cell_dict
       Next i
       DW_Dict.add "CheckPointDetail", DW_Rows
      
    End With
      
         '--------------------------
      Call SAPForDimensionweight(Remark_col, Title)
      Call NoteForDimensionweight(Remark_col, Title)
        '-------------------------------

       
  
59       '---------------------------
      
     
          '  BK(5) , "StyleMaterialConstruction""
      tbn = 0
      tbn = getTableNumber(5)
      If tbn = 0 Then
         MsgBox tb_tt(5) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(5) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
     
     
60   With WD.Tables(tbn)
      Rn = .Rows.count
      CLn = .Columns.count
      Cn = 0
      Dim SMC_Dict As Object
      Set SMC_Dict = CreateObject("Scripting.Dictionary")
          If .Cell(1, 2).Range.ContentControls.count > 0 Then
            Kw = Trim(Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1))
          Else
             Kw = Trim(Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 2))
          End If
          'SMC_Dict.add "5. Style / Material / Construction Result", Kw
        SetCheckPointTitle Title
       
       Dim SMC_Rows As New Collection
   
       For i = 4 To Rn
          Set cell_dict = New Scripting.Dictionary
          cell_dict.add "Category", "Style / Material / Construction"
          If .Cell(i, 2).Range.Bookmarks.count > 0 Then
             cell_dict.add Title(0), .Cell(i, 2).Range.Bookmarks(1).Name
          Else
            cell_dict.add Title(0), "N/A"
          End If

         For j = 1 To 2
           If .Cell(i, j).Range.ContentControls.count > 0 Then
           cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
         Else
           cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
         End If
       Next j
       cell_dict.add Title(3), ""
       
       For j = 3 To CLn
           If .Cell(i, j).Range.ContentControls.count > 0 Then
           cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
         Else
           cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
         End If
       Next j
          SMC_Rows.add cell_dict
        Next i
     
        SMC_Dict.add "CheckPointDetail", SMC_Rows
      
     End With
      
       '--------------------------
      Call SAPForStyleMaterialConstruction(Remark_col, Title)
      Call NoteForStyleMaterialConstruction(Remark_col, Title)
       '---------------------------
  
      
69        '-------------------------------
  
          '   BK(6) , "ProductColor"
      tbn = 0
      tbn = getTableNumber(6)
      If tbn = 0 Then
         MsgBox tb_tt(6) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(6) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
     
70   With WD.Tables(tbn)
        Rn = .Rows.count
        CLn = .Columns.count
        Cn = 0
        Dim CL_Dict As Object
        Set CL_Dict = CreateObject("Scripting.Dictionary")
         If .Cell(1, 2).Range.ContentControls.count > 0 Then
              Kw = Trim(Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1))
          Else
               Kw = Trim(Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 2))
          End If
          'CL_Dict.add "6. Product Color Result", Kw
          SetCheckPointTitle Title
          Dim CL_Rows As New Collection
   
          For i = 4 To Rn
             Set cell_dict = New Scripting.Dictionary
              cell_dict.add "Category", "Product Color"
           If .Cell(i, 2).Range.Bookmarks.count > 0 Then
             cell_dict.add Title(0), .Cell(i, 2).Range.Bookmarks(1).Name
          Else
            cell_dict.add Title(0), "N/A"
          End If

            For j = 1 To 2
           If .Cell(i, j).Range.ContentControls.count > 0 Then
           cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
         Else
           cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
         End If
       Next j
       cell_dict.add Title(3), ""
       
       For j = 3 To CLn
           If .Cell(i, j).Range.ContentControls.count > 0 Then
           cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
         Else
           cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
         End If
       Next j
               CL_Rows.add cell_dict
           Next i
     
           CL_Dict.add "CheckPointDetail", CL_Rows
    
     End With
     
      
       '-- ------------------------
       Call SAPForProductColor(Remark_col, Title)
       Call NoteForProductColor(Remark_col, Title)
       '---------------------------
   
79      '-----------------------
        '     BK(7) , "ProductLabelMarking""
      tbn = 0
      tbn = getTableNumber(7)
      If tbn = 0 Then
         MsgBox tb_tt(7) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(7) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
     
     
80   With WD.Tables(tbn)
        Rn = .Rows.count
        CLn = .Columns.count
        Cn = 0
        Dim LM_Dict As Object
        Set LM_Dict = CreateObject("Scripting.Dictionary")
           If .Cell(1, 2).Range.ContentControls.count > 0 Then
              Kw = Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1)
           Else
              Kw = Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 2)
          End If
         ' LM_Dict.add "7. Product Label / Marking Result", Kw
          
          SetCheckPointTitle Title
          Dim LM_Rows As New Collection
          For i = 4 To Rn
             Set cell_dict = New Scripting.Dictionary
             cell_dict.add "Category", "Product Label / Marking"
             If .Cell(i, 2).Range.Bookmarks.count > 0 Then
                 cell_dict.add Title(0), .Cell(i, 2).Range.Bookmarks(1).Name
             Else
                 cell_dict.add Title(0), "N/A"
             End If

             For j = 1 To 2
                If .Cell(i, j).Range.ContentControls.count > 0 Then
                   cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
                Else
                   cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
                End If
             Next j
             cell_dict.add Title(3), ""
       
             For j = 3 To CLn
               If .Cell(i, j).Range.ContentControls.count > 0 Then
                  cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
               Else
                  cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
               End If
             Next j
             LM_Rows.add cell_dict
          Next i
     
           LM_Dict.add "CheckPointDetail", LM_Rows
      End With
          
     
       
       '--------------------------
       Call SAPForProductLableMarking(Remark_col, Title)
       Call NoteForProductLableMarking(Remark_col, Title)
  
     
89        ' --    BK(8) , "ShippingMark"""
      tbn = 0
      tbn = getTableNumber(8)
      If tbn = 0 Then
         MsgBox tb_tt(8) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(8) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
     
     
90   With WD.Tables(tbn)
        Rn = .Rows.count
        CLn = .Columns.count
        Cn = 0
        Dim SM_Dict As Object
        Set SM_Dict = CreateObject("Scripting.Dictionary")
        If .Cell(1, 2).Range.ContentControls.count > 0 Then
            Kw = Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1)
        Else
             Kw = Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 2)
        End If
          'SM_Dict.add "8. Shipping Mark Result", Kw
        SetCheckPointTitle Title
           
        Dim SM_Rows As New Collection
   
        For i = 4 To Rn
              Set cell_dict = New Scripting.Dictionary
               cell_dict.add "Category", "Shipping Mark"
              If .Cell(i, 2).Range.Bookmarks.count > 0 Then
                cell_dict.add Title(0), .Cell(i, 2).Range.Bookmarks(1).Name
              Else
               cell_dict.add Title(0), "N/A"
              End If

              For j = 1 To 2
                If .Cell(i, j).Range.ContentControls.count > 0 Then
                   cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
                Else
                   cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
                End If
             Next j
            cell_dict.add Title(3), ""
       
            For j = 3 To CLn
              If .Cell(i, j).Range.ContentControls.count > 0 Then
                 cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
               Else
                 cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
              End If
            Next j
            SM_Rows.add cell_dict
       Next i
     
       SM_Dict.add "CheckPointDetail", SM_Rows
          
    End With
     
      '--------------------------
      Call SAPForShippingMark(Remark_col, Title)
      Call NoteForShippingMark(Remark_col, Title)
       
   
99         '---    BK(9), "PackingPackaging""
      tbn = 0
      tbn = getTableNumber(9)
      If tbn = 0 Then
         MsgBox tb_tt(9) & "was not found,program exit" & CHR(10) & UText(26410, 25214, 21040, 21517, 31216, 20026) & tb_tt(9) & UText(30340, 34920, 26684) & "," & UText(31243, 24207, 36864, 20986)
         Exit Sub
      End If
     
     
100  With WD.Tables(tbn)
      Rn = .Rows.count
      CLn = .Columns.count
      Cn = 0
      Dim PP_Dict As Object
      Set PP_Dict = CreateObject("Scripting.Dictionary")
          If .Cell(1, 2).Range.ContentControls.count > 0 Then
            Kw = Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 1)
          Else
             Kw = Mid(.Cell(1, 2).Range.Text, 1, Len(.Cell(1, 2).Range.Text) - 2)
          End If
          'PP_Dict.add "9. Packing / Packaging", Kw
         SetCheckPointTitle Title
       
         Dim PP_Rows As New Collection
   
        For i = 4 To Rn
          Set cell_dict = New Scripting.Dictionary
           cell_dict.add "Category", "Packing / Packaging"
          If .Cell(i, 2).Range.Bookmarks.count > 0 Then
             cell_dict.add Title(0), .Cell(i, 2).Range.Bookmarks(1).Name
          Else
            cell_dict.add Title(0), "N/A"
          End If

        For j = 1 To 2
           If .Cell(i, j).Range.ContentControls.count > 0 Then
           cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
         Else
           cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
         End If
       Next j
       cell_dict.add Title(3), ""
       
       For j = 3 To CLn
           If .Cell(i, j).Range.ContentControls.count > 0 Then
           cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
         Else
           cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
         End If
       Next j
            PP_Rows.add cell_dict
          Next i
     
           PP_Dict.add "CheckPointDetail", PP_Rows
          
     End With
     
      '--------------------------
        Call SAPForPackingPackaging(Remark_col, Title)
        Call NoteForPackingPackaging(Remark_col, Title)
       
       '---------------------------
    
       
       
       
109       '-----Semi-finished Product Inspection and Process Viewing
        For i = 1 To tb
         Kw = WD.Tables(i).Cell(1, 1).Range.Text
           If Kw Like "*" & "Semi-finished Product Inspection and Process Viewing" & "*" = True Then
              tbn = i
              GoTo 110
           End If
       Next i
        GoTo 120
        Exit Sub
     
      
110   With WD.Tables(tbn)
        Rn = .Rows.count
        CLn = .Columns.count
        Cn = 0
      
       SetCheckPointTitle Title
       
         Dim Semi_Rows As New Collection
   
        For i = 4 To Rn
          Set cell_dict = New Scripting.Dictionary
             cell_dict.add "Category", "Semi-finished Product Inspection and Process Viewing"
          If .Cell(i, 2).Range.Bookmarks.count > 0 Then
             cell_dict.add Title(0), .Cell(i, 2).Range.Bookmarks(1).Name
          Else
            cell_dict.add Title(0), "N/A"
          End If

        For j = 1 To 2
           If .Cell(i, j).Range.ContentControls.count > 0 Then
           cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
         Else
           cell_dict.add Title(j), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
         End If
       Next j
       cell_dict.add Title(3), ""
       
       For j = 3 To CLn
           If .Cell(i, j).Range.ContentControls.count > 0 Then
           cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 1)
         Else
           cell_dict.add Title(j + 1), Mid(.Cell(i, j).Range.Text, 1, Len(.Cell(i, j).Range.Text) - 2)
         End If
       Next j
            Semi_Rows.add cell_dict
          Next i
     
           'SemifinishedProduct_Dict.add "CheckPointDetail", Semi_Rows
          
     End With
     
      '--------------------------
         Call SAPForSemiProduct(Remark_col, Title)
         Call NoteForSemiProduct(Remark_col, Title)
       
      
       
       
      '------------------------Other note--
120    Dim OtherNote_Dict As Object
       Set OtherNote_Dict = CreateObject("Scripting.Dictionary")
       Dim OtherNote_Rows As Collection
       Dim Key As String
       Title(1) = "No"
       Title(2) = "NoteDescription"
       tb = WD.Tables.count - 1
       For i = 1 To tb
         With WD.Tables(i)
         If .Cell(1, 1).Range.Text Like "*Other Note*" = True Then
           ' Set OtherNote_Rows = New Collection
            Rn = .Rows.count
            SetRemarkTitle Title
            For j = 3 To Rn
               Set cell_dict = New Scripting.Dictionary
                 cell_dict.add Title(1), "Other Note"
                 cell_dict.add Title(2), ""

                For k = 1 To 2
                   If .Cell(j, k).Range.ContentControls.count > 0 Then
                     cell_dict.add Title(k + 2), Mid(.Cell(j, k).Range.Text, 1, Len(.Cell(j, k).Range.Text) - 1)
                   Else
                     cell_dict.add Title(k + 2), Mid(.Cell(j, k).Range.Text, 1, Len(.Cell(j, k).Range.Text) - 2)
                   End If
                Next k
                  cell_dict.add Title(5), ""
                  Remark_col.add cell_dict
                
                'OtherNote_Rows.add cell_dict
               
            Next j
         End If
        End With
       Next i
          ' OtherNote_Dict.add "Other Note", OtherNote_Rows

   '=========
      'GoTo 130
      Dim OnSiteTestData_dict As Object
      Dim CheckpointCode, CheckPoint As String
      Dim TestData_dict As Object
      Dim TestData_row As Collection
      Dim TestData_Table As Collection
      Set TestData_Table = New Collection
      
      Dim CDF_dict, BarCode_Dict, Shoe_dict As Object
      Dim CDFCheckpointCode, CDFCheckPoint As String
      Dim Kw1 As String
      Dim CDFTableCount, BarCodeTableCount, ShoeTableCount, WithPicTableCount As Integer
      
      For M = 2 To 9
         TbnB = getTableNumber(M)
         TbnE = getTableNumber(M + 1)
         ChkPoints = ""         '" & UText(29992, 20110, 36991, 20813, 21516, 19968, 27979, 35797, 28857, 26377, 22810, 34920, 26684, 30340, 37325, 22797, 35760, 24405) & "
         CDFTableCount = 0
         BarCodeTableCount = 0
         ShoeTableCount = 0
         WithPicTableCount = 0
         DTN = 0
         
         '" & UText(34920, 26684, 31181, 31867) & "
         For i = TbnB + 1 To TbnE - 1
            With WD.Tables(i)
                Kw = Mid(.Cell(1, 1).Range.Text, 1, Len(.Cell(1, 1).Range.Text) - 2)
                If Kw Like "*" & "Data Sheet for" & "*" = True Then
                  Kw = Trim(Replace(Kw, "Data Sheet for", ""))
                  For j = 1 To DTN
                     If DataTableG(j) = Kw Then
                       GoTo 122
                     End If
                  Next j
                  DTN = DTN + 1
                  DataTableG(DTN) = Kw
              End If
            End With
122      Next i
         If DTN = 0 Then GoTo 129
         For N = 1 To DTN
         Set TestData_dict = New Scripting.Dictionary
          Tn = 0       '" & UText(21516, 31867, 34920, 26684, 35745, 25968) & "
           For i = TbnB + 1 To TbnE - 1
             With WD.Tables(i)
               Kw = Mid(.Cell(1, 1).Range.Text, 1, Len(.Cell(1, 1).Range.Text) - 2)
               If Kw Like "*" & "Data Sheet for" & "*" = True Then
                  Kw = Trim(Replace(Kw, "Data Sheet for", ""))
                 
                  If Kw <> DataTableG(N) Then GoTo 128
                  CheckPoint = DataTableG(N)
                  
                  Tn = Tn + 1
                  For Each Cel In .Range.Cells
                     If Cel.Range.ContentControls.count > 0 Then
                        Kw = Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 1)
                      Else
                        Kw = Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2)
                     End If
                     If Left(Trim(LCase(Kw)), 4) = "item" Then       'normal table
                        CItemIdx = Cel.ColumnIndex
                        RItemIdx = Cel.rowIndex
                     ElseIf LCase(Kw) = "specification" Then
                       RSpeIdx = Cel.rowIndex
                       CSpeIdx = Cel.ColumnIndex
                     ElseIf LCase(Kw) = "tolerance" Then
                       RToLIdx = Cel.rowIndex
                       CToLIdx = Cel.ColumnIndex
                     ElseIf LCase(Kw) = "check point" Or LCase(Kw) = "check points" Then     'withpictable
                       RChkIdx = Cel.rowIndex
                       CChkIdx = Cel.ColumnIndex
                     ElseIf LCase(Kw) = "measurement point" Or LCase(Kw) = "measurement points" Then       '" & UText(38795, 31867, 23610, 23544, 34920) & "=Barcode" & UText(34920, 30340) & "position
                       RPsIdx = Cel.rowIndex
                       CPsIdx = Cel.ColumnIndex
                       StRow = RPsIdx + 2
                       GoTo 123
                     ElseIf LCase(Kw) = "measurement point" Or LCase(Kw) = "position" Then   'barcode tebla
                       RPsIdx = Cel.rowIndex
                       CPsIdx = Cel.ColumnIndex
                       StRow = RPsIdx + 1
                     ElseIf LCase(Kw) = "component name" Or LCase(Kw) = "component" Then  'CDF" & UText(34920) & "
                       RCpnIdx = Cel.rowIndex
                       CCpnIdx = Cel.ColumnIndex
                     ElseIf LCase(Kw) = "on cdf" Then
                       RCDFIdx = Cel.rowIndex
                       CCDFIdx = Cel.ColumnIndex
                     ElseIf LCase(Kw) = "result" Or LCase(Kw) = "actual" Or LCase(Kw) = "actual finding" Or LCase(Kw) = "findings" Then
                       RActIdx = Cel.rowIndex
                       CActIdx = Cel.ColumnIndex
                       StRow = RActIdx + 1
                       Exit For
                     End If
                  Next Cel
                  
123               Rn = .Rows.count
                  CLn = .Columns.count
                  If CLn > 10 Then
                     TableType = "ShoeTable"
                      StRow = RPsIdx + 2
                  ElseIf CLn = 5 Then
                     TableType = "CDFTable"
                     StRow = RCpnIdx + 1
                  ElseIf CLn = 4 Then
                     'If .Rows(2).Range.Cells.count = CLn Then
                     If Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2) = "Item No." Then
                        TableType = "NormalTable"
                         StRow = RItemIdx + 1
                     ElseIf LCase(Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2)) Like "*item no*" = True Then
                        TableType = "WithPicTable"
                         StRow = RChkIdx + 1
                     End If
                  ElseIf CLn = 3 Then
                     TableType = "BarCodeTable"
                     StRow = RPsIdx + 1
                  End If
              
                  If .Cell(1, 1).Range.Bookmarks.count > 0 Then
                     CheckpointCode = .Cell(1, 1).Range.Bookmarks(1).Name
                     If CheckpointCode = "N/A" Then
                        Set tbl = WD.Tables(TbnB)
                        For Each Cel In tbl.Range.Cells
                           If Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2) = CheckPoint Then
                               If Cel.Range.Bookmarks.count > 0 Then
                                   CheckpointCode = Cel.Range.Bookmarks(1).Name
                               Else
                                   CheckpointCode = "N/A"
                               End If
                               GoTo 125
                           End If
                        Next Cel
                    Else
                        Set tbl = WD.Tables(TbnB)
                        For Each Cel In tbl.Range.Cells
                           If Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2) = CheckPoint Then
                               If Cel.Range.Bookmarks.count > 0 Then
                                   CheckpointCode = Cel.Range.Bookmarks(1).Name
                               Else
                                   CheckpointCode = "N/A"
                               End If
                               GoTo 125
                           End If
                        Next Cel
                    End If
                 End If
              
125     If Tn = 1 Then
              SetTestDataTitle Title
              Set TestData_row = New Collection
              
        End If
      
        Select Case TableType
           
           Case "NormalTable"
                  
                   For k = StRow To Rn
                      Set cell_dict = New Scripting.Dictionary
                      For j = 1 To CLn
                          cell_dict.add Title(j), Mid(.Cell(k, j).Range.Text, 1, Len(.Cell(k, j).Range.Text) - 2)
                      Next j
                      For j = 5 To 7
                         cell_dict.add Title(j), ""
                      Next j
                      TestData_row.add cell_dict
                    Next k
                    
                Case "WithPicTable"
                   
                    KItmNo = Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2)
                    KItmNo = Replace(KItmNo, "Item No.:", "")
                    For k = StRow To Rn
                        Set cell_dict = New Scripting.Dictionary
                        cell_dict.add Title(1), KItmNo
                        cell_dict.add Title(2), Mid(.Cell(RSpeIdx + k, CSpeIdx).Range.Text, 1, Len(.Cell(RSpeIdx + k, CSpeIdx).Range.Text) - 2)
                        cell_dict.add Title(3), Mid(.Cell(RToLIdx + k, CToLIdx).Range.Text, 1, Len(.Cell(RToLIdx + k, CToLIdx).Range.Text) - 2)
                        cell_dict.add Title(4), Mid(.Cell(RActIdx + k, CActIdx).Range.Text, 1, Len(.Cell(RActIdx + k, CActIdx).Range.Text) - 2)
                        cell_dict.add Title(5), Mid(.Cell(PsIdx + k, CPsIdx).Range.Text, 1, Len(.Cell(RPsIdx + k, CPsIdx).Range.Text) - 2)
                        cell_dict.add Title(6), ""
                        cell_dict.add Title(7), ""
                        TestData_row.add cell_dict
                     Next k
               
                   Case "ShoeTable"
                 
                      KItmNo = Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2)
                      KItmNo = Replace(KItmNo, "Item No.:", "")
                      '====
                      Dn = 0
                      For j = StRow To Rn
                         Set Cel = .Cell(j - 2, 2)
                         If UCase(Left(Cel.Range.Text, 2)) = "SZ" Then
                            Dn = Dn + 1
                            ChkG(Dn) = j
              
                         End If
                      Next j
                      DR = ChkG(2) - ChkG(1)
                      For j = 2 To Dn
                          If ChkG(j) - ChkG(j - 1) <> DR Then
                            MsgBox "On Data Sheet for Insole Length of footwear table, different size has different number of check point rows, please correct before upload data to HUB!" & CHR(10) & _
                            UText(38795, 31867, 20013, 24213, 23610, 23544, 34920, 20013) & "，" & UText(19981, 21516, 30721, 26377, 19981, 21516, 30340, 27979, 37327, 28857, 34892, 25968) & "，" & UText(35831, 22312, 19978, 20256, 25968, 25454, 21040) & "HUB" & UText(21069) & "，" & UText(20808, 26356, 27491) & "！"
                            Exit Sub
                          End If
                      Next j
                     
                      For j = StRow To Rn Step DR
                          
                         Set Cel = .Cell(j - 2, 2)
                            Cn = 0
                            For k = 2 To CLn - 4 Step 4
                                Cn = Cn + 1
                               If Mid(.Cell(j, k).Range.Text, 1, Len(.Cell(j, k).Range.Text) - 2) = "" Then
                                  GoTo 128
                               End If
                               Set cell_dict = New Scripting.Dictionary
                               cell_dict.add Title(1), KItmNo
                               cell_dict.add Title(5), Mid(.Cell(j, 1).Range.Text, 1, Len(.Cell(j, 1).Range.Text) - 2)                 'checkpoint,position
                               cell_dict.add Title(7), Mid(.Cell(j - 2, Cn * 2 + 1).Range.Text, 1, Len(.Cell(j - 2, Cn * 2 + 1).Range.Text) - 2)    'size
                               cell_dict.add Title(6), Mid(.Cell(j - 1, k + 1).Range.Text, 1, Len(.Cell(j - 1, k + 1).Range.Text) - 2) 'color
                               cell_dict.add Title(2), Mid(.Cell(j, k).Range.Text, 1, Len(.Cell(j, k).Range.Text) - 2)           'specification
                               cell_dict.add Title(3), Mid(.Cell(j, CLn).Range.Text, 1, Len(.Cell(j, CLn).Range.Text) - 2)             'tolerance
                               cell_dict.add Title(4), Mid(.Cell(j, k + 1).Range.Text, 1, Len(.Cell(j, k + 1).Range.Text) - 2) & "," & _
                                                     Mid(.Cell(j, k + 2).Range.Text, 1, Len(.Cell(j, k + 2).Range.Text) - 2) & "," & _
                                                     Mid(.Cell(j, k + 3).Range.Text, 1, Len(.Cell(j, k + 3).Range.Text) - 2)           'result
                         
                               TestData_row.add cell_dict
                            Next k
                       
                      Next j
                 
                 Case "BarCodeTable"
                    
                     KItmNo = Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2)
                     KItmNo = Replace(KItmNo, "Item No.:", "")
                    
                      KwC = ""
                      KwS = ""
                      Set Cel = .Cell(3, 1)
                      Kw = Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2)
                      Kw = Replace(Kw, ":", ":")
                      
                      If Kw Like "Color:*" = True Then
                         KwC = Replace(Kw, "Color:", "")
                      ElseIf Kw Like "Size:*" = True Then
                         KwS = Replace(Kw, "Size:", "")
                      End If
                      Set Cel = .Cell(4, 1)
                      Kw = Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2)
                      Kw = Replace(Kw, ":", ":")
                      If Kw Like "Color:*" = True Then
                         KwC = Replace(Kw, "Color:", "")
                      ElseIf Kw Like "Size:*" = True Then
                         KwS = Replace(Kw, "Size:", "")
                      End If
                      
                      For j = StRow To Rn
                          Set cell_dict = New Scripting.Dictionary
                          cell_dict.add Title(1), KItmNo
                          cell_dict.add Title(6), KwC
                          cell_dict.add Title(7), KwS
                          Set Cel = .Cell(j, 1)
                          Kw = Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2)
                          cell_dict.add Title(5), Kw
                         
                          Set Cel = .Cell(j, 2)
                          Kw = Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2)
                          cell_dict.add Title(2), Kw
                         
                          Set Cel = .Cell(j, 3)
                          Kw = Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2)
                          cell_dict.add Title(4), Kw
                         
                          cell_dict.add Title(3), ""
                          
                          TestData_row.add cell_dict
                      Next j
             
            
                 Case "CDFTable"
                 
                    Title(1) = "ItemNo"
                    Title(2) = "ManufactureModelNo"
                    Title(3) = "ReportNo"
                    Title(4) = "No"
                    Title(5) = "ComponentName"
                    Title(6) = "OnCDF"
                    Title(7) = "Findings"
                    Title(8) = "Result"
                  
                 
                     KItmNo = Mid(.Cell(2, 1).Range.Text, 1, Len(.Cell(2, 1).Range.Text) - 2)
                     KItmNo = Replace(KItmNo, "Item No.:", "")
                   
                     Kw2 = Mid(.Cell(2, 2).Range.Text, 1, Len(.Cell(2, 2).Range.Text) - 2)
                     Kw2 = Replace(Kw2, "Manufacture Model No.:", "")
                    
                     Kw3 = Mid(.Cell(2, 3).Range.Text, 1, Len(.Cell(2, 3).Range.Text) - 2)
                     Kw3 = Replace(Kw3, "Report No.:", "")
                   
                     For j = StRow To Rn
                            Set cell_dict = New Scripting.Dictionary
                             
                             cell_dict.add Title(1), KItmNo
                             cell_dict.add Title(2), Kw2
                             cell_dict.add Title(3), Kw3
                             cell_dict.add Title(4), Trim(Mid(.Cell(j, 1).Range.Text, 1, Len(.Cell(j, 1).Range.Text) - 2))
                             cell_dict.add Title(5), Trim(Mid(.Cell(j, 2).Range.Text, 1, Len(.Cell(j, 2).Range.Text) - 2))
                             cell_dict.add Title(6), Trim(Mid(.Cell(j, 3).Range.Text, 1, Len(.Cell(j, 3).Range.Text) - 2))
                             cell_dict.add Title(7), Trim(Mid(.Cell(j, 4).Range.Text, 1, Len(.Cell(j, 4).Range.Text) - 2))
                             If .Cell(j, 5).Range.ContentControls.count > 0 Then
                               Kw = Trim(Mid(.Cell(j, 5).Range.Text, 1, Len(.Cell(j, 5).Range.Text) - 1))
                             Else
                              Kw = Trim(Mid(.Cell(j, 5).Range.Text, 1, Len(.Cell(j, 5).Range.Text) - 2))
                             End If
                              cell_dict.add Title(8), Kw
                            
                            TestData_row.add cell_dict
                        
                      Next j
                     
                      
                  End Select
              End If
           End With
128     Next i
            
            
           
       
          
          'TestData_dict.add "CheckPointCode", CheckpointCode
          'TestData_dict.add "CheckPoint", CheckPoint
         ' TestData_dict.add "CheckPointDataDetail", TestData_row
          'TestData_Table.add TestData_dict
            
        'json_String = JsonConverter.ConvertToJson(TestData_dict)
        'WD.Tables(2).cell(1, 1).Range.Text = json_String
         
      Next N
            
         'If DTN > 0 Then    '" & UText(26576, 20123, 20998, 39033, 27809, 26377, 25968, 25454, 20250, 23548, 33268) & "TestData_dict" & UText(20250, 37325, 22797) & "
            
           TestData_dict.add "CheckPointCode", CheckpointCode
           TestData_dict.add "CheckPoint", CheckPoint
           TestData_dict.add "CheckPointDataDetail", TestData_row
        
           TestData_Table.add TestData_dict
        'End If
      ' json_String = JsonConverter.ConvertToJson(TestData_Table)
        'WD.Tables(2).cell(1, 1).Range.Text = json_String
        
129   Next M
      
       ' json_String = JsonConverter.ConvertToJson(TestData_Table)
       ' WD.Tables(2).cell(1, 1).Range.Text = json_String
     
     
        'If CDFtableCount > 0 Then
         'Set TestData_dict = New Scripting.Dictionary
        ' TestData_dict.add "CheckPointCode", CheckpointCode
        ' TestData_dict.add "CheckPoint", CheckPoint
        ' TestData_dict.add "CDFCheckData", CDF_dict
         'TestData_Table.add TestData_dict
       ' End If
    Set OnSiteTestData_dict = New Scripting.Dictionary
        'TestData_dict.add "CheckpointCode of " & CDFCheckPoint, CDFCheckpointCode
        'TestData_dict.add "CheckPoint of " & CDFCheckPoint, CDF_dict
     OnSiteTestData_dict.add "OnSiteTestData", TestData_Table
       
      'json_String = JsonConverter.ConvertToJson(OnSiteTestData_dict)
      'WD.Tables(2).cell(1, 1).Range.Text = json_String
        
        
   '---------------------------------
130    Dim Report_Dict As Object
       Set Report_Dict = CreateObject("Scripting.Dictionary")
    'Report_Dict.add "Report No.", ReportNo
     Report_Dict.add "INSPECTIONINFORMATION", Info_Dict
    
    'Report_Dict.add "INSPECTIONRESULTSUMMARY", InsRes_Sum_Dict
     Report_Dict.add "INSPECTIONRESULTSUMMARY", InsRes_Sum_Rows
    
    'Report_Dict.add "1. Quantity", Qty_Dict
     Report_Dict.add "Quantity", Qty_Rows
     
     Report_Dict.add "QuantityTotal", TotalQty_Dict
   ' Report_Dict.add "SpecialAttentionPointforQuantity", SAPForQty_Dict
   
    'Report_Dict.add "ReferenceNoteforQuantity", NoteForQty_Dict
   
    
    'Report_Dict.add "Workmanship", Workmanship_Dict
    Report_Dict.add "Workmanship", Workmanship_Rows
    'Report_Dict.add "Defects", Defect_Dict
    Report_Dict.add "Defects", Item_Col
    'Report_Dict.add "SpecialAttentionPointforWorkmanship", SAPForWMS_dict
   
    'Report_Dict.add "ReferenceNoteforWorkmanship", NoteForWMS_dict
  
    
    'Report_Dict.add "OnSiteTest", OST_Dict
    Report_Dict.add "OnSiteTest", OST_Rows
    'Report_Dict.add "SpecialAttentionPointforOnSiteTest", SAPForOST_Dict
   
    'Report_Dict.add "ReferenceNoteforOnSiteTest", NoteforOST_Dict
    
    
    'Report_Dict.add "ProductDimension&Weight", DW_Dict
    Report_Dict.add "ProductDimension&Weight", DW_Rows
    'Report_Dict.add "SpecialAttentionPointforProductDimension&Weight", SAPForDW_Dict
  
    'Report_Dict.add "ReferenceNoteforProductDimension&Weight", NoteForDW_Dict
  
    
    'Report_Dict.add "Style/Material/Construction", SMC_Dict
    Report_Dict.add "Style/Material/Construction", SMC_Rows
    'Report_Dict.add "SpecialAttentionPointforStyle/Material/Construction", SAPforSMC_Dict
  
    'Report_Dict.add "ReferenceNoteforStyle/Material/Construction", NoteForSMC_Dict
   
    
    'Report_Dict.add "ProductColor", CL_Dict
    Report_Dict.add "ProductColor", CL_Rows
   ' Report_Dict.add "SpecialAttentionPointforProductColor", SAPForCL_Dict
   
    'Report_Dict.add "ReferenceNoteforProductColor", NoteforCL_Dict
   
    
    'Report_Dict.add "ProductLabel/Marking", LM_Dict
    Report_Dict.add "ProductLabel/Marking", LM_Rows
    'Report_Dict.add "SpecialAttentionPointforProductLabel/Marking", SAPForLM_Dict
  
    'Report_Dict.add "ReferenceNoteforProductLabel/Marking", NoteForLM_Dict
    
    
    'Report_Dict.add "ShippingMark", SM_Dict
    Report_Dict.add "ShippingMark", SM_Rows
    'Report_Dict.add "SpecialAttentionPointforShippingMark", SAPForSM_Dict
   
    'Report_Dict.add "ReferenceNoteforShippingMark", NoteForSM_Dict
    
    
    'Report_Dict.add "Packing/Packaging", PP_Dict
     Report_Dict.add "Packing/Packaging", PP_Rows
    'Report_Dict.add "SpecialAttentionPointforPacking/Packaging", SAPForPP_Dict
    
   ' Report_Dict.add "ReferenceNoteforPacking/Packaging", NoteforPP_Dict
     
    
    Report_Dict.add "OnSiteTestData", OnSiteTestData_dict
    
    Report_Dict.add "Remarks", Remark_col
    
   
    
200   json_String = JsonConverter.ConvertToJson(Report_Dict)
       
      'WD.Tables(2).cell(1, 1).Range.Text = json_String
       
       Dim Url As String
       Dim jsonResponse As Object
        
        'Set postjson = Report_Dict
        Url = WebBaseUrl & "/api/Interpretation/InspectionReport"
       Call sendDataToHUB(Url, json_String, jsonResponse)
      'Debug.Print jsonResponse
   
   'json_String = Replace(json_String, "/", "\/")
    'Debug.Print json_String
     'wd.Tables(2).Range.Text = json_String
   ' Call StoreJSON("Report All Info", json_String)
    
     'Dim Report_info As Object
    
   ' Set Report_info = jsonConverter.ParseJson(json_String)

End Sub

Sub DatacheckWMS()

Dim X As Integer

Set WD = ThisDocument
tb = WD.Tables.count - 1
For X = 1 To tb
   With WD.Tables(X)
     If .Cell(1, 1).Range.Text Like "*2. Workmanship*" = True Then
        MBG = X
     ElseIf .Cell(1, 1).Range.Text Like "*3. On Site Test*" = True Then
        MED = X
        GoTo 50
     End If
   End With
Next X
50 Call DatacheckModule(MBG, MED)
MsgBox "Data check completed！" & CHR(10) & UText(25968, 25454, 26816, 26597, 23436, 25104) & "！"

End Sub
Sub DatacheckOST()

Dim X As Integer

Set WD = ThisDocument
tb = WD.Tables.count - 1
For X = 1 To tb
   With WD.Tables(X)
     If .Cell(1, 1).Range.Text Like "*3. On Site Test*" = True Then
        MBG = X
     ElseIf .Cell(1, 1).Range.Text Like "*4. Product Dimension & Weight*" = True Then
        MED = X
        GoTo 50
     End If
   End With
Next X
50 Call DatacheckModule(MBG, MED)
MsgBox "Data check completed！" & CHR(10) & UText(25968, 25454, 26816, 26597, 23436, 25104) & "！"

End Sub

Sub DatacheckPDW()

Dim X As Integer

Set WD = ThisDocument
tb = WD.Tables.count - 1
For X = 1 To tb
   With WD.Tables(X)
     If .Cell(1, 1).Range.Text Like "*4. Product Dimension & Weight*" = True Then
        MBG = X
     ElseIf .Cell(1, 1).Range.Text Like "*5. Style / Material / Construction*" = True Then
        MED = X
        GoTo 50
     End If
   End With
Next X
50 Call DatacheckModule(MBG, MED)
MsgBox "Data check completed！" & CHR(10) & UText(25968, 25454, 26816, 26597, 23436, 25104) & "！"

End Sub
Sub DatacheckSMC()

Dim X As Integer

Set WD = ThisDocument
tb = WD.Tables.count - 1
For X = 1 To tb
   With WD.Tables(X)
     If .Cell(1, 1).Range.Text Like "*5. Style / Material / Construction*" = True Then
        MBG = X
     ElseIf .Cell(1, 1).Range.Text Like "*6. Product Color*" = True Then
        MED = X
        GoTo 50
     End If
   End With
Next X
50 Call DatacheckModule(MBG, MED)
MsgBox "Data check completed！" & CHR(10) & UText(25968, 25454, 26816, 26597, 23436, 25104) & "！"

End Sub
Sub DatacheckLM()

Dim X As Integer

Set WD = ThisDocument
tb = WD.Tables.count - 1
For X = 1 To tb
   With WD.Tables(X)
     If .Cell(1, 1).Range.Text Like "*7. Product Label / Marking*" = True Then
        MBG = X
     ElseIf .Cell(1, 1).Range.Text Like "*8. Shipping Mark*" = True Then
        MED = X
        GoTo 50
     End If
   End With
Next X
50 Call DatacheckModule(MBG, MED)
MsgBox "Data check completed！" & CHR(10) & UText(25968, 25454, 26816, 26597, 23436, 25104) & "！"
End Sub
Sub DatacheckPP()

Dim X As Integer

Set WD = ThisDocument
tb = WD.Tables.count - 1
For X = 1 To tb
   With WD.Tables(X)
     If .Cell(1, 1).Range.Text Like "*9. Packing / Packaging*" = True Then
        MBG = X
     ElseIf .Cell(1, 1).Range.Text Like "*Other Note*" = True Then
        MED = X
        GoTo 50
     End If
   End With
Next X
50 Call DatacheckModule(MBG, MED)
MsgBox "Data check completed！" & CHR(10) & UText(25968, 25454, 26816, 26597, 23436, 25104) & "！"

End Sub
Sub DatacheckModule(MBG, MED)
  '20241216
Dim i, j, k, X, Y, Rown, sn, Tn, An, AGn, ASn, TAn As Integer
Dim SType, TType, AType As String
Dim Spe, TAct, ActG
Dim VSpe(10), VAct(10, 10), VTolP, VTolM As Single
Dim Act(10, 10), BGG(10), EDG(10), Ln, BG(10, 10), ED(10, 10), Km1, Km2, Max, Min As Single
Dim KSpe, KAct, KTol, KTT, WVTolP, WVTolM As String
Dim msg As String
Dim MsgG
Dim uM As Integer
Dim dic As Object
Dim CItemIdx, CSpeIdx, CToLIdx, CActIdx, CChkIdx, CMeaIdx, CCpnIdx, CCDFIdx, CPsIdx As Integer
Dim RItemIdx, RSpeIdx, RToLIdx, RActIdx, RChkIdx, RMeaIdx, RCpnIdx, RCDFIdx, RPsIdx As Integer
Dim StRow, StColumn As Integer
Dim TableType As String
Dim Rn, CLn As Integer
Dim Cel As Cell
Dim NI As Integer
Set WD = ThisDocument

For X = MBG To MED
  With WD.Tables(X)
     If LCase(.Cell(1, 1).Range.Text) Like "*data sheet for*" = True Then
        Rn = .Rows.count
        CLn = .Columns.count
        If CLn > 10 Then
           TableType = "ShoeTable"
        ElseIf CLn = 5 Then
            TableType = "CDFTable"
        ElseIf CLn = 4 And .Cell(2, 1).Width = .PreferredWidth Then
            TableType = "WithPicTable"
        ElseIf CLn = 4 Then
          If .Rows(2).Cells.count = CLn Then
            TableType = "NormalTable"
          End If
        ElseIf CLn = 3 Then
          TableType = "BarCodeTable"
        
        End If
        
        For Each Cel In .Range.Cells
            If Cel.Range.ContentControls.count > 0 Then
              Kw = Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 1)
            Else
              Kw = Mid(Cel.Range.Text, 1, Len(Cel.Range.Text) - 2)
            End If
            If Left(Trim(LCase(Kw)), 4) = "item" Then
               CItemIdx = Cel.ColumnIndex
               RItemIdx = Cel.rowIndex
             
            ElseIf LCase(Kw) = "specification" Then
               RSpeIdx = Cel.rowIndex
               CSpeIdx = Cel.ColumnIndex
            ElseIf LCase(Kw) = "tolerance" Then
               RToLIdx = Cel.rowIndex
               CToLIdx = Cel.ColumnIndex
           
            ElseIf LCase(Kw) = "check point" Or LCase(Kw) = "check points" Then
               RChkIdx = Cel.rowIndex
               CChkIdx = Cel.ColumnIndex
            ElseIf LCase(Kw) = "measurement point" Or LCase(Kw) = "measurement points" Then
               RMeaIdx = Cel.rowIndex
               CMeaIdx = Cel.ColumnIndex
               StRow = RMeaIdx + 1
               
             ElseIf LCase(Kw) = "component name" Or LCase(Kw) = "component" Then
               RCpnIdx = Cel.rowIndex
               CCpnIdx = Cel.ColumnIndex
            ElseIf LCase(Kw) = "on cdf" Then
               RCDFIdx = Cel.rowIndex
               CCDFIdx = Cel.ColumnIndex
            ElseIf LCase(Kw) = "Position" Then
              RPsIdx = Cel.rowIndex
              CPsIdx = Cel.ColumnIndex
            
             ElseIf LCase(Kw) = "result" Or LCase(Kw) = "actual" Or LCase(Kw) = "actual finding" Or LCase(Kw) = "findings" Then
               RActIdx = Cel.rowIndex
               CActIdx = Cel.ColumnIndex
               StRow = RActIdx + 1
               Exit For
            End If
        Next Cel
     
        Select Case TableType
        
        Case "NormalTable", "WithPicTable"
       
       ' If Mid(.cell(2, 2).Range.Text, 1, Len(.cell(2, 2).Range.Text) - 2) = "Specification" Then
          ' If Mid(.cell(2, 3).Range.Text, 1, Len(.cell(2, 3).Range.Text) - 2) = "Tolerance" Then
               KTT = Mid(.Cell(1, 1).Range.Text, 1, Len(.Cell(1, 1).Range.Text) - 2)
               Rown = .Rows.count
               For i = StRow To Rown
               .Cell(i, CActIdx).Range.Font.ColorIndex = wdAuto
                  KSpe = Mid(Trim(.Cell(i, CSpeIdx).Range.Text), 1, Len(Trim(.Cell(i, CSpeIdx).Range.Text)) - 2)
                  KSpe = Replace(KSpe, "：", ":")
                  If InStr(1, KSpe, " ") > 0 Then
                    Do While InStr(1, KSpe, " ") > 0
                       KSpe = Replace(KSpe, " ", "")
                    Loop
                  End If
                  If InStr(1, KSpe, ":") > 0 Then
                     KSpe = Mid(KSpe, InStr(1, KSpe, ":") + 1, Len(KSpe) - InStr(1, KSpe, ":"))
                  End If
                    '.cell(i, CSpeIdx).Range.Text = KSpe
                     
                  KSpe = Replace(KSpe, "x", "*", , , vbTextCompare)
                  KSpe = Replace(KSpe, "X", "*", , , vbTextCompare)
                  KSpe = Replace(KSpe, "(", "")
                  KSpe = Replace(KSpe, ")", "")
                  KSpe = Replace(KSpe, "（", "")
                  KSpe = Replace(KSpe, "）", "")
                 
                  If KSpe = "" Or KSpe = "-" Then
                      SType = ""
                      msg = msg + KTT & " table " & i & " row, without specification data,  please corrct it!" & CHR(10) & _
                         KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(27809, 26377, 25351, 26631, 35201, 27714, 25968, 25454) & "！，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                         GoTo 50
                  ElseIf InStr(1, KSpe, "-") > 1 Then
                     SType = "-"
                     Km1 = Val(KSpe)
                     Km2 = Val(Mid(KSpe, InStr(1, KSpe, "-") + 1, Len(KSpe) - InStr(1, KSpe, "-")))
                     If Km1 < Km2 Then
                        Min = Km1
                        Max = Km2
                     Else
                        Min = Km2
                        Max = Km1
                     End If
                  ElseIf IsNumeric(Left(KSpe, 1)) Then
                     SType = "0"
                  ElseIf Left(KSpe, 1) = "<" Then
                     SType = "<"
                      KSpe = Val(Mid(KSpe, 2, Len(KSpe) - 1))
                  ElseIf Left(KSpe, 1) = "≤" Then
                     SType = "≤"
                      KSpe = Val(Mid(KSpe, 2, Len(KSpe) - 1))
                  ElseIf Left(KSpe, 1) = ">" Then
                     SType = ">"
                      KSpe = Val(Mid(KSpe, 2, Len(KSpe) - 1))
                  ElseIf Left(KSpe, 1) = "≥" Then
                      SType = "≥"
                      KSpe = Val(Mid(KSpe, 2, Len(KSpe) - 1))
                 ElseIf IsNumeric(Left(KSpe, 1)) = False Then
                     SType = "Character"
                 End If
                  Spe = Split(KSpe, "*")
                  sn = UBound(Spe)
                  For j = 0 To sn
                    VSpe(j) = Val(Spe(j))
                  Next j
                  '--------------------------
                  KTol = Mid(Trim(.Cell(i, CToLIdx).Range.Text), 1, Len(Trim(.Cell(i, CToLIdx).Range.Text)) - 2)
                  If InStr(1, KTol, " ") > 0 Then
                    Do While InStr(1, KTol, " ") > 0
                       KTol = Replace(KTol, " ", "")
                    Loop
                  End If
                  .Cell(i, CToLIdx).Range.Text = KTol
                  If KTol = "" Then
                      TType = ""
                      VTolP = 0
                      VTolM = 0
                  ElseIf InStr(1, KTol, "+/-") > 0 Then
                      WVTolP = Mid(KTol, InStr(1, KTol, "+/-") + 1, Len(KTol) - InStr(1, KTol, "+/-"))
                      If InStr(1, WVTolP, "%") > 0 Then
                          TType = "+/-%"
                          VTolP = Val(WVTolP) / 100
                          VTolM = VTolP
                      Else
                         TType = "+/-"
                         VTolP = Val(WVTolP)
                         VTolM = VTolP
                      End If
                  
                  ElseIf InStr(1, KTol, "+") > 0 And InStr(1, KTol, "/-") > 0 Then
                       WVTolP = Mid(KTol, InStr(1, KTol, "+") + 1, InStr(1, KTol, "/") - 1 - InStr(1, KTol, "+"))
                       WVTolM = Mid(KTol, InStr(1, KTol, "-") + 1, Len(KTol) - InStr(1, KTol, "-"))
                       If InStr(1, WVTolP, "%") > 0 Then
                           TType = "+X/-X%"
                           VTolP = Val(WVTolP) / 100
                           VTolM = Val(WVTolM) / 100
                       Else
                          TType = "+X/-X"
                          VTolP = Val(WVTolP)
                          VTolM = Val(WVTolM)
                      End If
                  
                  ElseIf IsNumeric(Left(KTol, 1)) = True And InStr(1, KTol, "/-") > 0 Then
                       WVTolP = Mid(KTol, 1, InStr(1, KTol, "/") - 1)
                       WVTolM = Mid(KTol, InStr(1, KTol, "-") + 1, Len(KTol) - InStr(1, KTol, "-"))
                       If InStr(1, WVTolP, "%") > 0 Then
                         TType = "X/-X%"
                         VTolP = Val(WVTolP) / 100
                         VTolM = Val(WVTolM) / 100
                      Else
                         TType = "X/-X"
                         VTolP = Val(WVTolP)
                         VTolM = Val(WVTolM) / 100
                      End If
                       
                  ElseIf IsNumeric(Left(KTol, 1)) Then
                    
                     If InStr(1, KTol, "%") > 0 Then
                       TType = "0%"
                       VTolP = Val(KTol) / 100
                       VTolM = 0
                     Else
                         TType = "0"
                         VTolP = Val(KTol)
                         VTolM = 0
                    End If
                     
                 ElseIf Left(KTol, 1) = "-" Then
                      WVTolM = Mid(KTol, 2, Len(KTol) - 1)
                     If InStr(1, WVTolM, "%") > 0 Then
                       TType = "-%"
                       VTolP = 0
                       VTolM = Val(WVTolM) / 100
                     Else
                       TType = "-"
                        VTolP = 0
                       VTolM = Val(WVTolM)
                     End If
                     
                ElseIf Left(KTol, 1) = "+" Then
                    WVTolP = Mid(KTol, 2, Len(KTol) - 1)
                    If InStr(1, WVTolP, "%") > 0 Then
                      TType = "+%"
                      VTolP = Val(WVTolP / 100)
                      VTolM = 0
                    Else
                       TType = "+"
                       VTolP = Val(WVTolP)
                       VTolM = 0
                    End If
                    
                ElseIf InStr(1, KTol, "±") > 0 Then
                      WVTolP = Mid(KTol, InStr(1, KTol, "±") + 1, Len(KTol) - InStr(1, KTol, "±"))
                      If InStr(1, WVTolP, "%") > 0 Then
                         TType = "±%"
                         VTolP = Val(WVTolP) / 100
                         VTolM = VTolP
                      Else
                         TType = "±"
                         VTolP = Val(WVTolP)
                         VTolM = VTolP
                      End If
                End If
                  
                  KAct = Mid(Trim(.Cell(i, CActIdx).Range.Text), 1, Len(Trim(.Cell(i, CActIdx).Range.Text)) - 2)
                  If InStr(1, KAct, " ") > 0 Then
                    Do While InStr(1, KAct, " ") > 0
                       KAct = Replace(KAct, " ", "")
                    Loop
                     
                  End If
                    
                    '.cell(i,  CActIdx).Range.Text = KAc
                 
                    .Cell(i, CActIdx).Range.Text = Replace(KAct, ",", ", ")
                  If Len(Trim(KAct)) = 0 Then
                      msg = msg + KTT & " table " & i & " row, without result data or the inputing result begin with text instead of data,  please corrct it!" & CHR(10) & _
                         KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(26410, 36755, 20837, 32467, 26524, 25968, 25454, 25110, 36755, 20837, 30340, 32467, 26524, 25968, 25454, 20197, 25991, 26412, 32780, 19981, 26159, 25968, 25454, 24320, 22836) & "！，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                         GoTo 50
                  End If
                  KAct = Replace(KAct, " ", "")
                  KAct = Replace(KAct, ";", ",")
                  KAct = Replace(KAct, "；", ",")
                  KAct = Replace(KAct, "，", ",")
                  KAct = Replace(KAct, CHR(13), ",")
                  KAct = Replace(KAct, "x", "*")
                  KAct = Replace(KAct, "X", "*")
                  KAct = Replace(KAct, "(", "")
                  KAct = Replace(KAct, ")", "")
                  KAct = Replace(KAct, "（", "")
                  KAct = Replace(KAct, "）", "")
                  
                  ActG = Split(KAct, ",")
                  AGn = UBound(ActG)
                  
                  BGG(0) = 1
                  EDG(0) = Len(ActG(0))
                  
                  For j = 1 To AGn
                    Ln = Len(ActG(j))
                    BGG(j) = EDG(j - 1) + 2
                    EDG(j) = BGG(j) + Ln - 1
                  Next j
                  
                  For j = 0 To AGn
                     TAct = Split(ActG(j), "*")
                     TAn = UBound(TAct)
                     BG(j, 0) = BGG(j)
                     ED(j, 0) = BGG(j) + Len(TAct(0)) - 1
                     
                     For k = 1 To TAn
                       Ln = Len(TAct(k))
                       BG(j, k) = ED(j, k - 1) + 2
                       ED(j, k) = BG(j, k) + Ln - 1
                     Next k
                      
                    If TAn <> sn Then    '" & UText(26816, 26597, 25968, 25454, 32500, 25968) & "
                       If KSpe <> "" And KSpe <> "-" Then
                         msg = msg + KTT & " table " & i & " row, the data format of result not match to specification's,  please corrct it!" & CHR(10) & _
                         KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(32467, 26524, 25968, 25454, 30340, 26684, 24335, 21644, 21442, 25968, 30340, 19981, 21305, 37197) & "！，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                         GoTo 50
                      End If
                    End If
                   
                    For k = 0 To TAn
                        Act(j, k) = TAct(k)
                        VAct(j, k) = Val(Act(j, k))
                        
                        Select Case TType       'Tol type
                        Case ""
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications and tolerance data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                               msg = msg + KTT & " table " & i & " row, no tolerance data was input,  please check!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(20801, 24046, 20540, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "<"
                                
                                If VAct(j, k) >= VSpe(k) Then
                                    For Y = BG(j, k) + j To ED(j, k) + j
                                       .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                    Next Y
                                End If
                                
                             Case "≤"
                                If VAct(j, k) > VSpe(k) Then
                                    For Y = BG(j, k) + j To ED(j, k) + j
                                       .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                    Next Y
                                End If
                             Case ">"
                                If VAct(j, k) <= VSpe(k) Then
                                    For Y = BG(j, k) + j To ED(j, k) + j
                                       .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                    Next Y
                                End If
                             Case "≥"
                                 If VAct(j, k) < VSpe(k) Then
                                    For Y = BG(j, k) + j To ED(j, k) + j
                                       .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                    Next Y
                                 End If
                             Case "-"
                                  If VAct(j, k) < Min Or VAct(j, k) > Max Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                  End If
                            End Select
                       Case "0%"
                            Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications and tolerance data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) > VSpe(k) + VTolP * VSpe(k) Or VAct(j, k) < VSpe(k) Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                 End If
                             Case "<"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                                
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                            
                       Case "0"
                            Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications and tolerance data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) > VSpe(k) + VTolP Or VAct(j, k) < VSpe(k) Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                 End If
                             Case "<"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                                
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                        Case "-%"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) > VSpe(k) Or VAct(j, k) < VSpe(k) - VTolM * VSpe(k) Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                              Case "Character"
                              If Act(j, k) <> Spe(k) Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             End Select
                            
                        Case "-"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) > VSpe(k) Or VAct(j, k) < VSpe(k) - VTolM Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                              Case "Character"
                              If Act(j, k) <> Spe(k) Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             End Select
                         Case "+%"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) < VSpe(k) Or VAct(j, k) > VSpe(k) + VTolP * VSpe(k) Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                         Case "+"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) < VSpe(k) Or VAct(j, k) > VSpe(k) + VTolP Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                         Case "±%"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                            
                             Case "0"
                                   If VAct(j, k) < VSpe(k) - VTolM * VSpe(k) Or VAct(j, k) > VSpe(k) + VTolP * VSpe(k) Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                   End If
                                
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                            
                            
                         Case "±"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                            
                             Case "0"
                                   If VAct(j, k) < VSpe(k) - VTolM Or VAct(j, k) > VSpe(k) + VTolP Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                   End If
                                
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                        
                        Case "+/-%"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) < VSpe(k) - VTolM * VSpe(k) Or VAct(j, k) > VSpe(k) + VTolP * VSpe(k) Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & "table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            
                            End Select
                        Case "+/-"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) < VSpe(k) - VTolM Or VAct(j, k) > VSpe(k) + VTolP Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & "table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            
                            End Select
                         
                         Case "+X/-X%"
                            Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) < VSpe(k) - VTolM * VSpe(k) Or VAct(j, k) > VSpe(k) + VTolP * VSpe(k) Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & "table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            
                            End Select
                        
                        Case "+X/-X"
                            Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) < VSpe(k) - VTolM Or VAct(j, k) > VSpe(k) + VTolP Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & "table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            
                            End Select
                        
                        Case "X/-X%"
                            Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) < VSpe(k) - VTolM * VSpe(k) Or VAct(j, k) > VSpe(k) + VTolP * VSpe(k) Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & "table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                       
                        Case "X/-X"
                            Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) < VSpe(k) - VTolM Or VAct(j, k) > VSpe(k) + VTolP Then
                                        For Y = BG(j, k) + j To ED(j, k) + j
                                          .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                                        Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & "table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(21442, 25968, 21644, 20801, 24046, 20540, 30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                       
                                
                       End Select
                    Next k
                  Next j
                  
50             Next i
          ' End If
        'End If
        '--
         Case "CDFTable"
          
          KTT = Mid(.Cell(1, 1).Range.Text, 1, Len(.Cell(1, 1).Range.Text) - 2)
               Rown = .Rows.count
               For i = StRow To Rown
                  .Cell(i, CActIdx).Range.Font.ColorIndex = wdAuto
                 If .Cell(i, CCDFIdx).Range.InlineShapes.count = 0 Then
                    KSpe = Mid(Trim(.Cell(i, CCDFIdx).Range.Text), 1, Len(Trim(.Cell(i, CCDFIdx).Range.Text)) - 2)
                    KSpe = Replace(KSpe, "：", ":")
                      'If InStr(1, KSpe, " ") > 0 Then
                         'Do While InStr(1, KSpe, " ") > 0
                            'KSpe = Replace(KSpe, " ", "")
                         'Loop
                       'End If
                    If InStr(1, KSpe, ":") > 0 Then
                       KSpe = Mid(KSpe, InStr(1, KSpe, ":") + 1, Len(KSpe) - InStr(1, KSpe, ":"))
                    End If
                    If KSpe = "" Or KSpe = "-" Then
                       msg = msg + KTT & " table " & i & " row, without specification data,  please corrct it!" & CHR(10) & _
                         KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(27809, 26377, 25351, 26631, 35201, 27714, 25968, 25454) & "！，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                         GoTo 100
                    End If
                    KAct = Mid(Trim(.Cell(i, CActIdx).Range.Text), 1, Len(Trim(.Cell(i, CActIdx).Range.Text)) - 2)
                    If InStr(1, KAct, " ") > 0 Then
                       Do While InStr(1, KAct, " ") > 0
                          KAct = Replace(KAct, " ", "")
                       Loop
                    End If
                     .Cell(i, CActIdx).Range.Text = Replace(KAct, ",", ", ")
                    If Len(Trim(KAct)) = 0 Then
                         msg = msg + KTT & " table " & i & " row, without result data or the inputing result,  please corrct it!" & CHR(10) & _
                         KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(26410, 36755, 20837, 32467, 26524, 25968, 25454) & "！，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                         GoTo 100
                    End If
                    KAct = Replace(KAct, " ", "")
                    KAct = Replace(KAct, ";", ",")
                    KAct = Replace(KAct, "；", ",")
                    KAct = Replace(KAct, "，", ",")
                    KAct = Replace(KAct, CHR(13), ",")
                    KAct = Replace(KAct, "(", "")
                    KAct = Replace(KAct, ")", "")
                    KAct = Replace(KAct, "（", "")
                    KAct = Replace(KAct, "）", "")
                    
                    ActG = Split(KAct, ",")
                    AGn = UBound(ActG)
                  
                    BGG(0) = 1
                    EDG(0) = Len(ActG(0))
                  
                    For j = 1 To AGn
                      Ln = Len(ActG(j))
                      BGG(j) = EDG(j - 1) + 2
                      EDG(j) = BGG(j) + Ln - 1
                    Next j
                  
                  
                    .Cell(i, CActIdx).Range.Text = Replace(KAct, ",", ", ")
                
                    For j = 0 To AGn
                       If Len(KSpe) <> Len(ActG(j)) Then
                          For Y = BGG(j) + j To EDG(j) + j
                            .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                          Next Y
                       Else
                          NI = 0
                          For Y = BGG(j) + j To EDG(j) + j
                             NI = NI + 1
                            If Mid(ActG(j), NI, 1) <> Mid(KSpe, NI, 1) Then
                                .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                             End If
                          Next Y
                       End If
                     Next j
                 End If
100          Next i

        Case "BarCodeTable"
          
          KTT = Mid(.Cell(1, 1).Range.Text, 1, Len(.Cell(1, 1).Range.Text) - 2)
               Rown = .Rows.count
               For i = StRow To Rown
                  .Cell(i, CActIdx).Range.Font.ColorIndex = wdAuto
                  If .Cell(i, CSpeIdx).Range.InlineShapes.count = 0 Then
                    KSpe = Mid(Trim(.Cell(i, CSpeIdx).Range.Text), 1, Len(Trim(.Cell(i, CSpeIdx).Range.Text)) - 2)
                    KSpe = Replace(KSpe, "：", ":")
                      'If InStr(1, KSpe, " ") > 0 Then
                         'Do While InStr(1, KSpe, " ") > 0
                            'KSpe = Replace(KSpe, " ", "")
                         'Loop
                       'End If
                    If InStr(1, KSpe, ":") > 0 Then
                       KSpe = Mid(KSpe, InStr(1, KSpe, ":") + 1, Len(KSpe) - InStr(1, KSpe, ":"))
                    End If
                    If KSpe = "" Or KSpe = "-" Then
                       msg = msg + KTT & " table " & i & " row, without specification data,  please corrct it!" & CHR(10) & _
                         KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(27809, 26377, 25351, 26631, 35201, 27714, 25968, 25454) & "！，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                         GoTo 100
                    End If
                    KAct = Mid(Trim(.Cell(i, CActIdx).Range.Text), 1, Len(Trim(.Cell(i, CActIdx).Range.Text)) - 2)
                    If InStr(1, KAct, " ") > 0 Then
                       Do While InStr(1, KAct, " ") > 0
                          KAct = Replace(KAct, " ", "")
                       Loop
                    End If
                     .Cell(i, CActIdx).Range.Text = Replace(KAct, ",", ", ")
                    If Len(Trim(KAct)) = 0 Then
                         msg = msg + KTT & " table " & i & " row, without result data or the inputing result,  please corrct it!" & CHR(10) & _
                         KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(26410, 36755, 20837, 32467, 26524, 25968, 25454) & "！，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                         GoTo 100
                    End If
                    KAct = Replace(KAct, " ", "")
                    KAct = Replace(KAct, ";", ",")
                    KAct = Replace(KAct, "；", ",")
                    KAct = Replace(KAct, "，", ",")
                    KAct = Replace(KAct, CHR(13), ",")
                    KAct = Replace(KAct, "(", "")
                    KAct = Replace(KAct, ")", "")
                    KAct = Replace(KAct, "（", "")
                    KAct = Replace(KAct, "）", "")
                    
                    ActG = Split(KAct, ",")
                    AGn = UBound(ActG)
                  
                    BGG(0) = 1
                    EDG(0) = Len(ActG(0))
                  
                    For j = 1 To AGn
                      Ln = Len(ActG(j))
                      BGG(j) = EDG(j - 1) + 2
                      EDG(j) = BGG(j) + Ln - 1
                    Next j
                  
                  
                    .Cell(i, CActIdx).Range.Text = Replace(KAct, ",", ", ")
                
                    For j = 0 To AGn
                       If Len(KSpe) <> Len(ActG(j)) Then
                          For Y = BGG(j) + j To EDG(j) + j
                            .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                          Next Y
                       Else
                          NI = 0
                          For Y = BGG(j) + j To EDG(j) + j
                             NI = NI + 1
                            If Mid(ActG(j), NI, 1) <> Mid(KSpe, NI, 1) Then
                                .Cell(i, CActIdx).Range.Characters(Y).Font.ColorIndex = wdRed
                             End If
                          Next Y
                       End If
                     Next j
                End If
        Next i


      Case "ShoeTable"
        KTT = Mid(.Cell(1, 1).Range.Text, 1, Len(.Cell(1, 1).Range.Text) - 2)
        Rown = .Rows.count
        For i = StRow + 1 To Rown
           NI = 0
           Set Cel = .Cell(i, 2)
           Kw = Trim(Cel.Range.Text)
           Kw = Mid(Kw, 1, Len(Kw) - 2)
        
          If LCase(Kw) <> "sz" And LCase(Kw) <> "cl" Then
             For k = 2 To 13
                Set Cel = .Cell(i, k)
                Kw = Cel.Range.Text
                Kw = Mid(Kw, 1, Len(Kw) - 2)
                If Trim(Kw) = "" Then
                    NI = NI + 1
                End If
            Next k
            If NI > 0 Then
                 msg = msg + KTT & " table " & i & " row" & NI & "　cells, without inputing specification or measurement data,  please corrct it!" & CHR(10) & _
                KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(26377) & NI & UText(20010, 21333, 20803, 26684, 27809, 26377, 23610, 23544, 35201, 27714, 25110, 23454, 38469, 27979, 37327, 25968, 25454) & "！，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                GoTo 150
            End If
         
            Set Cel = .Cell(i, 14) 'cln
            KTol = Cel.Range.Text
            KTol = Mid(KTol, 1, Len(KTol) - 2)
               
            If KTol = "" Then
                   msg = msg + KTT & " table " & i & " row, without Tol data,  please corrct it!" & CHR(10) & _
                   KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，" & UText(27809, 26377, 20801, 24046, 25968, 25454) & "！，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                   GoTo 150
            ElseIf InStr(1, KTol, "+/-") > 0 Then
            
                  TType = "+/-"
                  WVTolP = Mid(KTol, InStr(1, KTol, "+/-") + 1, Len(KTol) - InStr(1, KTol, "+/-"))
                  If InStr(1, WVTolP, "%") > 0 Then
                     VTolP = Val(WVTolP) / 100
                     VTolM = VTolP
                  Else
                    VTolP = Val(WVTolP)
                    VTolM = VTolP
                 End If
                  
             ElseIf InStr(1, KTol, "+") > 0 And InStr(1, KTol, "/-") > 0 Then
                 TType = "+X/-X"
                 WVTolP = Mid(KTol, InStr(1, KTol, "+") + 1, InStr(1, KTol, "/") - 1 - InStr(1, KTol, "+"))
                 If InStr(1, WVTolP, "%") > 0 Then
                    VTolP = Val(WVTolP) / 100
                 Else
                    VTolP = Val(WVTolP)
                 End If
                      
                 WVTolM = Mid(KTol, InStr(1, KTol, "-") + 1, Len(KTol) - InStr(1, KTol, "-"))
                 If InStr(1, WVTolM, "%") > 0 Then
                    VTolM = Val(WVTolM) / 100
                 Else
                    VTolM = Val(WVTolM)
                 End If
                  
            ElseIf IsNumeric(Left(KTol, 1)) = True And InStr(1, KTol, "/-") > 0 Then
                 TType = "X/-X"
                 WVTolP = Mid(KTol, 1, InStr(1, KTol, "/") - 1)
                 If InStr(1, WVTolP, "%") > 0 Then
                    VTolP = Val(WVTolP) / 100
                 Else
                    VTolP = Val(WVTolP)
                 End If
                      
                 WVTolM = Mid(KTol, InStr(1, KTol, "-") + 1, Len(KTol) - InStr(1, KTol, "-"))
                 If InStr(1, WVTolM, "%") > 0 Then
                    VTolM = Val(WVTolM) / 100
                 Else
                    VTolM = Val(WVTolM)
                 End If
                       
            ElseIf IsNumeric(Left(KTol, 1)) Then
                TType = "0"
                If InStr(1, KTol, "%") > 0 Then
                   VTolP = Val(KTol) / 100
                Else
                   VTolP = Val(KTol)
               End If
               VTolM = 0
            ElseIf Left(KTol, 1) = "-" Then
               TType = "-"
               VTolP = 0
               WVTolM = Mid(KTol, 2, Len(KTol) - 1)
               If InStr(1, WVTolM, "%") > 0 Then
                   VTolM = Val(WVTolM) / 100
               Else
                   VTolM = Val(WVTolM)
              End If
                     
            ElseIf Left(KTol, 1) = "+" Then
               TType = "+"
               WVTolP = Mid(KTol, 2, Len(KTol) - 1)
               If InStr(1, WVTolP, "%") > 0 Then
                  VTolP = Val(WVTolP / 100)
               Else
                  VTolP = Val(WVTolP)
               End If
               VTolM = 0
            ElseIf InStr(1, KTol, "±") > 0 Then
               TType = "±"
                   
               WVTolP = Mid(KTol, InStr(1, KTol, "±") + 1, Len(KTol) - InStr(1, KTol, "±"))
               If InStr(1, WVTolP, "%") > 0 Then
                  VTolP = Val(WVTolP) / 100
                  VTolM = VTolP
               Else
                  VTolP = Val(WVTolP)
                  VTolM = VTolP
              End If
            End If
                  
            For j = 2 To 10 Step 4
                Set Cel = .Cell(i, j)
                Cel.Range.Font.ColorIndex = wdAuto
                KSpe = Val(Cel.Range.Text)
                For k = 1 To 3
                  Set Cel = .Cell(i, j + k)
                  KAct = Cel.Range.Text
                  Cel.Range.Font.ColorIndex = wdAuto
                  ValAct = Val(KAct)
                  Select Case TType
                  Case "-"
                      Cel.Range.Font.ColorIndex = wdAuto
                  Case Else
                    
                       If ValAct > VTolP Or ValAct < -VTolM Then Cel.Range.Font.ColorIndex = wdRed
                     
                  End Select
               Next k
            Next j
         End If
150     Next i
      End Select
     End If
  End With
Next X
    If msg <> "" Then
    MsgG = Split(msg, CHR(10))
    uM = UBound(MsgG)
    Set dic = CreateObject("Scripting.Dictionary")
      
       For i = 0 To uM
          dic(MsgG(i)) = MsgG(i)
         'Dic(Chkw(i)) = ChkW(i)
       Next i
       ArrChk = dic.Keys
       Dn = dic.count
       For i = 0 To Dn - 1
         msg1 = msg1 & ArrChk(i) & CHR(10)
       Next i
   End If
    
    If msg1 <> "" Then MsgBox msg1
   
End Sub
Sub Datacheck()

Dim i, j, k, X, Y, Rown, sn, Tn, An, AGn, ASn, TAn As Integer
Dim SType, TType, AType As String
Dim Spe, TAct, ActG
Dim VSpe(10), VAct(10, 10), VTolP, VTolM As Single
Dim Act(10, 10), BGG(10), EDG(10), Ln, BG(10, 10), ED(10, 10), Km1, Km2, Max, Min As Single
Dim KSpe, KAct, KTol, KTT, WVTolP, WVTolM As String
Dim msg As String
Dim MsgG
Dim uM As Integer
Dim dic As Object

Set WD = ThisDocument
tb = ThisDocument.Tables.count - 1
For X = 1 To tb
  With WD.Tables(X)
     If LCase(.Cell(1, 1).Range.Text) Like "*data sheet for*" = True Then
        If Mid(.Cell(2, 2).Range.Text, 1, Len(.Cell(2, 2).Range.Text) - 2) = "Specification" Then
           If Mid(.Cell(2, 3).Range.Text, 1, Len(.Cell(2, 3).Range.Text) - 2) = "Tolerance" Then
               KTT = Mid(.Cell(1, 1).Range.Text, 1, Len(.Cell(1, 1).Range.Text) - 2)
               Rown = .Rows.count
               For i = 3 To Rown
               .Cell(i, 4).Range.Font.ColorIndex = wdAuto
                  KSpe = Mid(Trim(.Cell(i, 2).Range.Text), 1, Len(Trim(.Cell(i, 2).Range.Text)) - 2)
                  KSpe = Replace(KSpe, "x", "*", , , vbTextCompare)
                  KSpe = Replace(KSpe, "X", "*", , , vbTextCompare)
                  KSpe = Replace(KSpe, "(", "")
                  KSpe = Replace(KSpe, ")", "")
                  KSpe = Replace(KSpe, "（", "")
                  KSpe = Replace(KSpe, "）", "")
                  
                 
                  If KSpe = "" Or KSpe = "-" Then
                      SType = ""
                  ElseIf InStr(1, KSpe, "-") > 1 Then
                     SType = "-"
                     Km1 = Val(KSpe)
                     Km2 = Val(Mid(KSpe, InStr(1, KSpe, "-") + 1, Len(KSpe) - InStr(1, KSpe, "-")))
                     If Km1 < Km2 Then
                        Min = Km1
                        Max = Km2
                     Else
                        Min = Km2
                        Max = Km1
                     End If
                  ElseIf IsNumeric(Left(KSpe, 1)) Then
                     SType = "0"
                  ElseIf Left(KSpe, 1) = "<" Then
                     SType = "<"
                      KSpe = Val(Mid(KSpe, 2, Len(KSpe) - 1))
                  ElseIf Left(KSpe, 1) = "≤" Then
                     SType = "≤"
                      KSpe = Val(Mid(KSpe, 2, Len(KSpe) - 1))
                  ElseIf Left(KSpe, 1) = ">" Then
                     SType = ">"
                      KSpe = Val(Mid(KSpe, 2, Len(KSpe) - 1))
                  ElseIf Left(KSpe, 1) = "≥" Then
                     SType = "≥"
                      KSpe = Val(Mid(KSpe, 2, Len(KSpe) - 1))
                 End If
                  Spe = Split(KSpe, "*")
                  sn = UBound(Spe)
                  For j = 0 To sn
                    VSpe(j) = Val(Spe(j))
                  Next j
                  '--------------------------
                  KTol = Mid(Trim(.Cell(i, 3).Range.Text), 1, Len(Trim(.Cell(i, 3).Range.Text)) - 2)
                  
                  If KTol = "" Then
                      TType = ""
                      VTolP = 0
                      VTolM = 0
                  ElseIf InStr(1, KTol, "/") > 0 Then
                      TType = "+/-"
                      WVTolP = Mid(KTol, InStr(1, KTol, "+") + 1, InStr(1, KTol, "/") - 1)
                      If InStr(1, WVTolP, "%") > 0 Then
                         VTolP = Val(WVTolP) / 100
                      Else
                         VTolP = Val(WVTolP)
                      End If
                      WVTolM = Mid(KTol, InStr(1, KTol, "-") + 1, Len(KTol) - InStr(1, KTol, "-"))
                      If InStr(1, WVTolM, "%") > 0 Then
                         VTolM = Val(WVTolM) / 100
                      Else
                         VTolM = Val(WVTolM)
                      End If
                       
                  ElseIf IsNumeric(Left(KTol, 1)) Then
                     TType = "0"
                     If InStr(1, KTol, "%") > 0 Then
                       VTolP = Val(KTol) / 100
                     Else
                         VTolP = Val(KTol)
                    End If
                     VTolM = 0
                  ElseIf Left(KTol, 1) = "-" Then
                     TType = "-"
                     VTolP = 0
                     WVTolM = Mid(KTol, 2, Len(KTol) - 1)
                     If InStr(1, WVTolM, "%") > 0 Then
                       VTolM = Val(WVTolM) / 100
                     Else
                       VTolM = Val(WVTolM)
                     End If
                     
                  ElseIf Left(KTol, 1) = "+" Then
                    TType = "+"
                    WVTolP = Mid(KTol, 2, Len(KTol) - 1)
                    If InStr(1, WVTolP, "%") > 0 Then
                      VTolP = Val(WVTolP / 100)
                    Else
                       VTolP = Val(WVTolP)
                    End If
                    VTolM = 0
                  ElseIf InStr(1, KTol, "±") > 0 Then
                     TType = "±"
                     If InStr(1, KTol, "%") > 0 Then
                       VTolP = Val(Mid(KTol, 2, Len(KTol) - 1)) / 100
                       VTolM = Val(Mid(KTol, 2, Len(KTol) - 1)) / 100
                     Else
                       VTolP = Val(Mid(KTol, 2, Len(KTol) - 1))
                       VTolM = Val(Mid(KTol, 2, Len(KTol) - 1))
                     
                     End If
                  End If
                  
                  KAct = Mid(Trim(.Cell(i, 4).Range.Text), 1, Len(Trim(.Cell(i, 4).Range.Text)) - 2)
                  KAct = Replace(KAct, " ", "")
                  KAct = Replace(KAct, ";", ",")
                  KAct = Replace(KAct, "；", ",")
                  KAct = Replace(KAct, "，", ",")
                  KAct = Replace(KAct, CHR(13), ",")
                  KAct = Replace(KAct, "x", "*")
                  KAct = Replace(KAct, "X", "*")
                  KAct = Replace(KAct, "(", "")
                  KAct = Replace(KAct, ")", "")
                  KAct = Replace(KAct, "（", "")
                  KAct = Replace(KAct, "）", "")
                  
                  ActG = Split(KAct, ",")
                  AGn = UBound(ActG)
                  
                  BGG(0) = 1
                  EDG(0) = Len(ActG(0))
                  
                  For j = 1 To AGn
                    Ln = Len(ActG(j))
                    BGG(j) = EDG(j - 1) + 2
                    EDG(j) = BGG(j) + Ln - 1
                  Next j
                  
                  For j = 0 To AGn
                     TAct = Split(ActG(j), "*")
                     TAn = UBound(TAct)
                     BG(j, 0) = BGG(j)
                     ED(j, 0) = BGG(j) + Len(TAct(0)) - 1
                     
                     For k = 1 To TAn
                       Ln = Len(TAct(k))
                       BG(j, k) = ED(j, k - 1) + 2
                       ED(j, k) = BG(j, k) + Ln - 1
                     Next k
                      
                    If TAn <> sn Then    '" & UText(26816, 26597, 25968, 25454, 32500, 25968) & "
                       If KSpe <> "" And KSpe <> "-" Then
                         msg = msg + KTT & " table " & i & " row, the data format of result not match to specification's,  please corrct it!" & CHR(10) & _
                         KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，Result" & UText(25968, 25454, 30340, 26684, 24335, 21644) & "specification" & UText(30340, 19981, 21305, 37197) & "！，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                         GoTo 50
                      End If
                    End If
                   
                    For k = 0 To TAn
                      Act(j, k) = TAct(k)
                      VAct(j, k) = Val(Act(j, k))
                      
                        Select Case TType       'Tol type
                        
                        Case ""
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications and tolerance data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification" & UText(21644) & "tolerance" & UText(25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                               msg = msg + KTT & " table " & i & " row, no tolerance data was input,  please check!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，tolerance" & UText(25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "<"
                                If VAct(j, k) >= VSpe(k) Then
                                   For Y = BG(j, k) To ED(j, k)
                                      .Cell(i, 4).Range.Characters(Y).Font.ColorIndex = wdRed
                                   Next Y
                                End If
                             Case "≤"
                                If VAct(j, k) > VSpe(k) Then
                                   For Y = BG(j, k) To ED(j, k)
                                      .Cell(i, 4).Range.Characters(Y).Font.ColorIndex = wdRed
                                   Next Y
                                End If
                             Case ">"
                                If VAct(j, k) <= VSpe(k) Then
                                   For Y = BG(j, k) To ED(j, k)
                                      .Cell(i, 4).Range.Characters(Y).Font.ColorIndex = wdRed
                                   Next Y
                                End If
                             Case "≥"
                                 If VAct(j, k) < VSpe(k) Then
                                   For Y = BG(j, k) To ED(j, k)
                                      .Cell(i, 4).Range.Characters(Y).Font.ColorIndex = wdRed
                                   Next Y
                                 End If
                             Case "-"
                                  If VAct(j, k) < Min Or VAct(j, k) > Max Then
                                   For Y = BG(j, k) To ED(j, k)
                                      .Cell(i, 4).Range.Characters(Y).Font.ColorIndex = wdRed
                                   Next Y
                                  End If
                            End Select
                            
                       Case "0"
                            Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications and tolerance data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification" & UText(21644) & "tolerance" & UText(25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) > VSpe(k) + VSpe(k) * VTolP Or VAct(j, k) < VSpe(k) Then
                                   For Y = BG(j, k) To ED(j, k)
                                      .Cell(i, 4).Range.Characters(Y).Font.ColorIndex = wdRed
                                   Next Y
                                 End If
                             Case "<"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                                
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                            
                        Case "-"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification" & UText(25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) > VSpe(k) Or VAct(j, k) < VSpe(k) - VSpe(k) * VTolM Then
                                   For Y = BG(j, k) To ED(j, k)
                                      .Cell(i, 4).Range.Characters(Y).Font.ColorIndex = wdRed
                                   Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             End Select
                             
                         Case "+"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification" & UText(25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) < VSpe(k) Or VAct(j, k) > VSpe(k) + VSpe(k) * VTolP Then
                                   For Y = BG(j, k) To ED(j, k)
                                      .Cell(i, 4).Range.Characters(Y).Font.ColorIndex = wdRed
                                   Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                            
                          Case "±"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification" & UText(25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) < VSpe(k) - VSpe(k) * VTolM Or VAct(j, k) > VSpe(k) + VSpe(k) * VTolP Then
                                   For Y = BG(j, k) To ED(j, k)
                                      .Cell(i, 4).Range.Characters(Y).Font.ColorIndex = wdRed
                                   Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                            
                          Case "+/-"
                             Select Case SType
                             Case ""
                               msg = msg + KTT & " table " & i & " row, no specifications data was input,  please corrct it!" & CHR(10) & _
                               KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification" & UText(25968, 25454, 26410, 24405, 20837) & "，" & UText(35831, 26356, 27491) & "!" & CHR(10)
                             Case "0"
                                If VAct(j, k) < VSpe(k) - VSpe(k) * VTolM Or VAct(j, k) > VSpe(k) + VSpe(k) * VTolP Then
                                   For Y = BG(j, k) To ED(j, k)
                                      .Cell(i, 4).Range.Characters(Y).Font.ColorIndex = wdRed
                                   Next Y
                                End If
                             Case "<"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≤"
                                msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case ">"
                                 msg = msg + KTT & "table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "≥"
                                  msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                             Case "-"
                                 msg = msg + KTT & " table " & i & " row, the requirement of specification and tolerance was conflicted. please check!" & CHR(10) & _
                                 KTT & UText(34920, 26684, 31532) & i & UText(34892) & "，specification and tolerance" & UText(30340, 35201, 27714, 20914, 31361) & "，" & UText(35831, 26816, 26597) & "!" & CHR(10)
                            End Select
                       
                       End Select
                    Next k
                  Next j
                  
50             Next i
           End If
        End If
     End If
  End With
Next X
    If msg <> "" Then
    MsgG = Split(msg, CHR(10))
    uM = UBound(MsgG)
    Set dic = CreateObject("Scripting.Dictionary")
      
       For i = 0 To uM
          dic(MsgG(i)) = MsgG(i)
         'Dic(Chkw(i)) = ChkW(i)
       Next i
       ArrChk = dic.Keys
       Dn = dic.count
       For i = 0 To Dn - 1
         msg1 = msg1 & ArrChk(i) & CHR(10)
       Next i
   End If
    
    If msg1 <> "" Then MsgBox msg1
   
End Sub

Sub GeneratePhotoTable(Tn)


Set WD = ThisDocument

       If WD.Tables(Tn + 1).Columns.count <> 2 Or LCase(WD.Tables(Tn + 1).Cell(1, 1).Range.Text) Like "*special attention point*" = True Or _
              LCase(WD.Tables(Tn + 1).Cell(1, 1).Range.Text) Like "*reference note*" = True Then
        Rown = WD.Tables(Tn).Rows.count
        WD.Tables(Tn).Cell(Rown, 1).Select
        Selection.MoveDown
        Selection.TypeParagraph
        WD.Tables.add Range:=Selection.Range, NumRows:=2, NumColumns:=2, _
        DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
            With Selection.Tables(1)
                If .Style <> UText(32593, 26684, 22411) Then
                 .Style = UText(32593, 26684, 22411)
               End If
            End With
            Call Setborders025
     End If
                WD.Tables(Tn + 1).Select
                  With Selection.Tables(1)
                    Rown = .Rows.count
                       If Rown = 1 Then
                         .Rows(1).Select
                           Selection.InsertRows 1
                        ElseIf Rown > 2 Then
                           For i = Rown To 3 Step -1
                          .Rows(i).Delete
                           Next i
                        End If
                  End With
End Sub






Sub getHUBdata(Url, postJson, jsonResponse)
   
    Dim ReqResponse  As String
    Dim TokenJson As Object
    
    Set TokenJson = GetToken("AiReport", "nNt9=*FdpW8U")
    
    If TokenJson("success") Then
    
        Dim req As Object
        Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
        Dim ReqStatus As Integer
  
         
        With req
            .Open "POST", Url, False
            .setRequestHeader "token", TokenJson("data")("token")
            .setRequestHeader "userid", EncryptedUserId
            .setRequestHeader "appid", AppId
            .send JsonConverter.ConvertToJson(postJson)
             ReqStatus = .Status
             ReqResponse = .responseText
        End With
     End If
 
     If Not ReqStatus = 200 Then
        MsgBox Url & "No response form HUB or an error occurred" & CHR(10) & "HUB" & UText(26410, 21709, 24212, 25110, 26377, 38169, 35823, 21457, 29983) & "!", vbCritical
        Set req = Nothing
        Exit Sub
    End If
     Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
     
End Sub
Sub sendDataToHUB(Url, json_String, jsonResponse)
   
    Dim ReqResponse As String
    Dim TokenJson As Object
    Set TokenJson = GetToken("AiReport", "nNt9=*FdpW8U")
    If TokenJson("success") Then
        Dim req As Object
        Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
        Dim ReqStatus As Integer
        With req
            .Open "POST", Url, False
            .setRequestHeader "token", TokenJson("data")("token")
            .setRequestHeader "userid", EncryptedUserId
            .setRequestHeader "appid", AppId
            .send json_String
             ReqStatus = .Status
             ReqResponse = .responseText
        End With
     End If
   
    Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
   
  
    If ReqStatus <> 200 Then
        MsgBox "No response form HUB or an error occurred, please check the internet or try again later" & CHR(10) & _
         "HUB" & UText(26410, 21709, 24212, 25110, 26377, 38169, 35823, 21457, 29983) & "，" & UText(35831, 26816, 26597, 32593, 32476, 25110, 31245, 21518, 20877, 35797, 19968, 27425) & "!"
          Exit Sub
     Else
         Dim Code As String
         Code = jsonResponse("code")
        If Code = "200" Then
           MsgBox "The data of the inspection report has been successfully uploaded to the HUB" & CHR(10) & UText(26816, 39564, 25253, 21578, 30340, 25968, 25454, 24050, 25104, 21151, 19978, 20256, 36882, 21040) & "HUB!"
        ElseIf Code = "201" Then
          MsgBox "An error of " & "'" & jsonResponse("jsonMsg") & "'" & " Occured,Please contact the IT team for support" & CHR(10) & _
           UText(21457, 29983) & "'" & jsonResponse("jsonMsg") & "'" & UText(38169, 35823, 23548, 33268, 25253, 21578, 25968, 25454, 26410, 25104, 21151, 19978, 20256, 21040) & "HUB!，" & UText(35831, 32852, 31995) & "IT" & UText(32452, 20197, 33719, 24471, 25903, 25345) & "!"
           
        ElseIf Code = "202" Then
           MsgBox "Due to an system error, the data of the inspection report was not successfully uploaded to the HUB. Please contact the IT team for support" & CHR(10) & _
           UText(30001, 20110, 21457, 29983, 38169, 35823, 23548, 33268, 25253, 21578, 25968, 25454, 26410, 25104, 21151, 19978, 20256, 21040) & "HUB!，" & UText(35831, 32852, 31995) & "IT" & UText(32452, 20197, 33719, 24471, 25903, 25345) & "!"
          
        End If
    End If

    Set req = Nothing
   ' Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
     
End Sub



Sub Load_Product_Lev1()
   
    Dim i, N, Cn, k As Integer
    Dim postJson As Object
    Dim jsonResponse As Object
    Dim Url As String
    
    Dim Data As Object
    Dim Cat_Large_IDs() As String
    Dim Cat_Large_Codes() As String
    Dim Cat_Large_Names() As String
    Dim Cat_Large_NamesE() As String
    Dim Cat_Large_Number As Integer
    
    Dim UF As Product_Lev1
    Set UF = New Product_Lev1
    Dim LeftPos As Integer
    Dim TopPos As Integer
    Dim Chk As Object
     
     N = UF.Controls.count
     Cn = 0

 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CommandButton" Then
       If SearchDefect = True Then
         If UF.Controls(i).Caption = "AddCheckPoints" Then
            UF.Controls(i).Visible = False
         End If
      End If
      End If
 Next i
     
    Url = WebBaseUrl & "/api/sqlhub/exec?appid=md_69&configName=getProductPrimaryType"
    Call getHUBdata(Url, postJson, jsonResponse)
      
        Set Data = jsonResponse("data")
        Cat_Large_Number = Data.count
        
        ReDim Cat_Large_IDs(1 To Cat_Large_Number)
        ReDim Cat_Large_Codes(1 To Cat_Large_Number)
        ReDim Cat_Large_Names(1 To Cat_Large_Number)
        ReDim Cat_Large_NamesE(1 To Cat_Large_Number)
        For i = 1 To Cat_Large_Number
          Cat_Large_IDs(i) = Data(i)(HubKeyProductLargeCategoryId())
          Cat_Large_Codes(i) = Data(i)(HubKeyProductLargeCategoryCode())
          Cat_Large_Names(i) = Data(i)(HubKeyProductLargeCategoryNameZh())
          Cat_Large_NamesE(i) = Data(i)(HubKeyProductLargeCategoryNameEn())
        Next i

 
   If N < Cat_Large_Number Then N = Cat_Large_Number
   
    Dim ScreenWidth, ScreenHeight As Long
    Dim UFwidth, UFheight As Long
    Dim Defwidth, Defheight As Long
    
    ScreenWidth = Application.Width
    ScreenHeight = Application.Height
    Defwidth = 0.17 * ScreenWidth * 0.8 * 2
    Defheight = 0.04 * ScreenHeight * 0.9
    TopPos = 0.09 * ScreenHeight * 0.9
    
    Cn = 0
    LeftPos = 20 ' " & UText(35774, 32622, 22797, 36873, 26694, 30340, 21021, 22987) & "Left" & UText(20301, 32622) & "

    For i = 1 To Cat_Large_Number
        Set Chk = UF.Controls.add("Forms.CheckBox.1", Cat_Large_Names(i), True)
        With Chk
            .Caption = i & "|" & Cat_Large_Names(i) & "|" & Cat_Large_NamesE(i)
            .Tag = Cat_Large_IDs(i)
            If i Mod 20 = 0 Then
              Cn = Cn + 1
             .Top = TopPos
            Else
             .Top = TopPos + (i Mod 20) * Defheight
            
            End If
             .Left = LeftPos + (Cn) * Defwidth
         
            .Width = Defwidth  ' " & UText(35774, 32622, 22797, 36873, 26694, 30340, 23485, 24230) & "
            .Height = Defheight
            .Visible = True
            For j = 1 To N_Cat_Large
              If Cat_Large_IDs(i) = Cat_Large_ID(j) Then
               .Value = True
              End If
            Next j
        .Font.Size = 10
        End With
    Next i
    UF.Width = LeftPos + (Cn + 1) * Defwidth + 100
    If i < 20 Then
      UF.Height = TopPos + (i + 2 + 2) * Defheight
    Else
      UF.Height = TopPos + (20 + 2 + 2) * Defheight
    End If
   
    UF.Show vbModal
    
100 End Sub

Sub Load_Product_Lev2()

Dim i, N, Cn, k As Integer
Dim Url As String
Dim postJson As Object
Dim postID As String
Dim Data As Object
Dim Cat_Middle_IDs() As String
Dim Cat_Middle_Codes() As String
Dim Cat_Middle_Names() As String
Dim Cat_Middle_NamesE() As String
Dim Cat_Middle_Number As Integer

Dim UF As Product_Lev2
Set UF = New Product_Lev2
Dim LeftPos As Integer
Dim TopPos As Integer
Dim Chk As Object
Dim Str As String
Dim p As Integer


  Set postJson = New Scripting.Dictionary
  Cn = 0
  For k = 1 To N_Cat_Large
   
    postJson("szdl") = Cat_Large_ID(k)
    Url = WebBaseUrl & "/api/sqlhub/exec?appid=md_19&configName=getProductSecondaryType"
    Call getHUBdata(Url, postJson, jsonResponse)
    Dim jasonstr As String
    jasonstr = JsonConverter.ConvertToJson(jsonResponse)
    'Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
    Set Data = jsonResponse("data")
    Cat_Middle_Number = Data.count
    ReDim Cat_Middle_IDs(1 To Cat_Middle_Number)
    ReDim Cat_Middle_Codes(1 To Cat_Middle_Number)
    ReDim Cat_Middle_Names(1 To Cat_Middle_Number)
    ReDim Cat_Middle_NamesE(1 To Cat_Middle_Number)
    For i = 1 To Cat_Middle_Number
          Cat_Middle_IDs(i) = Data(i)(HubKeyProductMiddleCategoryId())
          Cat_Middle_Codes(i) = Data(i)(HubKeyProductMiddleCategoryCode())
          Cat_Middle_Names(i) = Data(i)(HubKeyProductMiddleCategoryNameZh())
          Cat_Middle_NamesE(i) = Data(i)(HubKeyProductMiddleCategoryNameEn())
    Next i
    
    Dim ScreenWidth, ScreenHeight As Long
    Dim UFwidth, UFheight As Long
    Dim Defwidth, Defheight As Long
    
    ScreenWidth = Application.Width
    ScreenHeight = Application.Height
    Defwidth = 0.17 * ScreenWidth * 0.8 * 2
    Defheight = 0.04 * ScreenHeight * 0.9
    TopPos = 0.09 * ScreenHeight * 0.9
     LeftPos = 20 ' " & UText(35774, 32622, 22797, 36873, 26694, 30340, 21021, 22987) & "Left" & UText(20301, 32622) & "
  
      For i = 1 To Cat_Middle_Number
        Set Chk = UF.Controls.add("Forms.CheckBox.1", Cat_Middle_Names(i), True)
        With Chk
            .Caption = i & "|" & Cat_Middle_Names(i) & "|" & Cat_Middle_NamesE(i)
            .Tag = Cat_Large_ID(k) & "|" & Cat_Middle_IDs(i)
            If i Mod 20 = 0 Then
              Cn = Cn + 1
             .Top = TopPos
            Else
             .Top = TopPos + (i Mod 20) * Defheight
              '.Top = 18 * i + TopPos
            End If
            .Left = LeftPos + (k - 1 + Cn) * Defwidth
            .Width = Defwidth
            .Height = Defheight
            .Visible = True
           
             For j = 1 To N_Cat_Middle
                p = InStrRev(Cat_Middle_ID(j), "|", Len(Cat_Middle_ID(j)))
                Str = Right(Cat_Middle_ID(j), Len(Cat_Middle_ID(j)) - p)
              
                If Cat_Middle_IDs(i) = Str Then
                  .Value = True
                End If
              Next j
            .Font.Size = 10
        End With
    Next i
    
    If N < Cat_Middle_Number Then N = Cat_Middle_Number
     
    Set Chk = UF.Controls.add("Forms.TextBox.1", "TextBox1", True)
     With Chk
        .Top = 40
         .Left = LeftPos + (k - 1) * Defwidth
        .Value = Cat_Large(k) & "|" & Cat_LargeE(k)
     
       .Width = (Cn + 1) * Defwidth
       .Height = Defheight
       .Visible = True
    End With
Next k
    UF.Width = LeftPos + (k - 1) * Defwidth + 40
     
    If i < 20 Then
      UF.Height = TopPos + (i + 2 + 2) * Defheight
    Else
      UF.Height = TopPos + (20 + 2 + 2) * Defheight
    End If
   
     UF.Show vbModal
End Sub

Sub Load_Product_Lev3()
   
    Dim i, N, Cn, k As Integer
    Dim postJson As Object
    Dim jsonResponse As Object
    Dim Url As String
    Dim postID As String
    Dim pos As Integer
    Dim Data As Object
    Dim Cat_Small_IDs() As String
    Dim Cat_Small_Codes() As String
    Dim Cat_Small_Names() As String
    Dim Cat_Small_NamesE() As String
    Dim Cat_Small_Number As Integer
    
    Dim UF As Product_Lev3
    Set UF = New Product_Lev3
    Dim LeftPos As Integer
    Dim TopPos As Integer
    Dim Chk As Object
    Dim Str As String
    Dim p As Integer

    Set postJson = New Scripting.Dictionary
    Cn = 0
    For k = 1 To N_Cat_Middle
     
      pos = InStrRev(Cat_Middle_ID(k), "|")
      postID = Mid(Cat_Middle_ID(k), pos + 1)
     postJson("szcpzl") = postID
     Url = WebBaseUrl & "/api/sqlhub/exec?appid=md_39&configName=getProductTertiaryType"
     Call getHUBdata(Url, postJson, jsonResponse)
   
    'Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
    Set Data = jsonResponse("data")
    Cat_Small_Number = Data.count
    ReDim Cat_Small_IDs(1 To Cat_Small_Number)
    ReDim Cat_Small_Codes(1 To Cat_Small_Number)
    ReDim Cat_Small_Names(1 To Cat_Small_Number)
    ReDim Cat_Small_NamesE(1 To Cat_Small_Number)
    For i = 1 To Cat_Small_Number
          Cat_Small_IDs(i) = Data(i)(HubKeyProductSmallCategoryId())
          Cat_Small_Codes(i) = Data(i)(HubKeyProductSmallCategoryCode())
          Cat_Small_Names(i) = Data(i)(HubKeyProductSmallCategoryNameZh())
          Cat_Small_NamesE(i) = Data(i)(HubKeyProductSmallCategoryNameEn())
    Next i
  
    Dim ScreenWidth, ScreenHeight As Long
    Dim UFwidth, UFheight As Long
    Dim Defwidth, Defheight As Long
    
    ScreenWidth = Application.Width
    ScreenHeight = Application.Height
    Defwidth = 0.17 * ScreenWidth * 0.8 * 2
    Defheight = 0.04 * ScreenHeight * 0.9
    TopPos = 0.09 * ScreenHeight * 0.9

    LeftPos = 20
    For i = 1 To Cat_Small_Number
        Set Chk = UF.Controls.add("Forms.CheckBox.1", Cat_Small_Names(i), True)
        With Chk
            .Caption = i & "|" & Cat_Small_Names(i) & "|" & Cat_Small_NamesE(i)
            .Tag = Cat_Middle_ID(k) & "|" & Cat_Small_IDs(i)
            If i Mod 20 = 0 Then
              Cn = Cn + 1
             .Top = TopPos
            Else
             .Top = TopPos + (i Mod 20) * Defheight
              '.Top = 18 * i + TopPos
            End If
             .Left = LeftPos + (k - 1 + Cn) * Defwidth
           
            .Width = Defwidth ' " & UText(35774, 32622, 22797, 36873, 26694, 30340, 23485, 24230) & "
            .Height = Defheight
            .Visible = True
           
         
             For j = 1 To N_Cat_Small
               p = InStrRev(Cat_Small_ID(j), "|", Len(Cat_Small_ID(j)))
               Str = Right(Cat_Small_ID(j), Len(Cat_Small_ID(j)) - p)
              
               If Cat_Small_IDs(i) = Str Then
                .Value = True
               End If
             Next j
        End With
    Next i
     If N < Cat_Small_Number Then N = Cat_Small_Number
    
    Set Chk = UF.Controls.add("Forms.TextBox.1", "TextBox1", True)
     With Chk
        .Top = 40
        .Left = LeftPos + (k - 1) * Defheight
       .Value = Cat_Middle(k) & "|" & Cat_MiddleE(k)
       .Width = (Cn + 1) * Defwidth ' " & UText(35774, 32622, 22797, 36873, 26694, 30340, 23485, 24230) & "
       .Height = Defheight
       .Visible = True
       
    End With
   Next k
    UF.Width = LeftPos + (k - 1) * Defwidth + 40
    
    If i < 20 Then
      UF.Height = TopPos + (i + 2 + 2) * Defheight
    Else
      UF.Height = TopPos + (20 + 2 + 2) * Defheight
    End If
    UF.Show vbModal
100 End Sub


Sub Load_Product_Lev4()
    
      Dim i, N, Cn, k As Integer
    Dim postJson As Object
    Dim jsonResponse As Object
    Dim Url As String
    Dim postID As String
    Dim pos As Integer
    Dim Data As Object
    Dim Product_IDs() As String
    Dim Product_Codes() As String
    Dim Product_Names() As String
    Dim Product_NamesE() As String
    Dim Product_Number As Integer
    
    Dim UF As Product_Lev4
    Set UF = New Product_Lev4
    Dim LeftPos As Integer
    Dim TopPos As Integer
    Dim Chk As Object
    Dim Str As String
    Dim p As Integer
    
     N = UF.Controls.count
     Cn = 0

 For i = 0 To N - 1
     If TypeName(UF.Controls(i)) = "CheckBox" Then
       If SearchDefect = True Then
         If UF.Controls(i).Caption = UText(36873, 25321, 20135, 21697, 29305, 24449) & "|Choose Product Features" Then
            UF.Controls(i).Visible = False
         End If
      End If
      End If
 Next i

    Set postJson = New Scripting.Dictionary
    Cn = 0
    For k = 1 To N_Cat_Small
      pos = InStrRev(Cat_Small_ID(k), "|")
      postID = Mid(Cat_Small_ID(k), pos + 1)
     postJson("szcpxl") = postID
     Url = WebBaseUrl & "/api/sqlhub/exec?appid=md_40&configName=getProductName"
     Call getHUBdata(Url, postJson, jsonResponse)
   
    'Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
    Set Data = jsonResponse("data")
    Product_Number = Data.count
    ReDim Product_IDs(1 To Product_Number)
    ReDim Product_Codes(1 To Product_Number)
    ReDim Product_Names(1 To Product_Number)
    ReDim Product_NamesE(1 To Product_Number)
    For i = 1 To Product_Number
          Product_IDs(i) = Data(i)(HubKeyProductNameId())
          Product_Codes(i) = Data(i)(HubKeyProductNameCode())
          Product_Names(i) = Data(i)(HubKeyProductNameZh())
          Product_NamesE(i) = Data(i)(HubKeyProductNameEn())
    Next i
    
    Dim ScreenWidth, ScreenHeight As Long
    Dim UFwidth, UFheight As Long
    Dim Defwidth, Defheight As Long
    
    ScreenWidth = Application.Width
    ScreenHeight = Application.Height
    Defwidth = 0.17 * ScreenWidth * 0.8 * 2
    Defheight = 0.04 * ScreenHeight * 0.9
    TopPos = 0.09 * ScreenHeight * 0.9
    LeftPos = 20
    For i = 1 To Product_Number
        Set Chk = UF.Controls.add("Forms.CheckBox.1", Product_Names(i), True)
        With Chk
            .Caption = i & "|" & Product_Names(i) & "|" & Product_NamesE(i)
            .Tag = Cat_Small_ID(k) & "|" & Product_IDs(i)
            If i Mod 20 = 0 Then
              Cn = Cn + 1
             .Top = TopPos + 10
            Else
             .Top = TopPos + (i Mod 20) * 25 + 10
              '.Top = 18 * i + TopPos
            End If
            .Left = LeftPos + (k - 1 + Cn) * Defwidth
            .Width = Defwidth
            .Height = Defheight
            .Visible = True
            .Enabled = True
             
             For j = 1 To N_Product
               p = InStrRev(Product_ID(j), "|", Len(Product_ID(j)))
               Str = Right(Product_ID(j), Len(Product_ID(j)) - p)
               If Product_IDs(i) = Str Then
                .Value = True
               End If
            Next j
          .Font.Size = 10
        End With
    Next i
     If N < Product_Number Then N = Product_Number
     Set Chk = UF.Controls.add("Forms.TextBox.1", "TextBox1", True)
     With Chk
       .Top = 40
       .Left = LeftPos + (k - 1) * Defwidth
       .Value = Cat_Small(k) & "|" & Cat_SmallE(k)
       .Width = (Cn + 1) * Defwidth
       .Height = Defheight
       .Visible = True
    End With
  Next k
   

      'Uf.ScrollBars = 3
      
        UF.Width = LeftPos + (k - 1 + Cn) * Defwidth + 120
        If i < 20 Then
           UF.Height = TopPos + (i + 2 + 2) * Defheight
        Else
           UF.Height = TopPos + (20 + 2 + 2) * Defheight
        End If
    
     UF.Show vbModal
100 End Sub

Sub Load_Product_Features()
    
    Dim i, N, Cn, k As Integer
    Dim postJson As Object
    Dim jsonResponse As Object
    Dim Url As String
    
    Dim Data As Object
    Dim Product_Feature_IDs() As String
    Dim Product_Feature_Codes() As String
    Dim Product_Feature_Names() As String
    Dim Product_Feature_NamesE() As String
    Dim Product_Feature_Number As Integer
    
    Dim UF As Product_Features
    Set UF = New Product_Features
    Dim LeftPos As Integer
    Dim TopPos As Integer
    Dim Chk As Object
    Dim W(4) As Integer
    
       
     Set postJson = New Scripting.Dictionary
     Cn = 0
     For k = 1 To N_Cat_Large
       postJson("cpdlid") = Cat_Large_ID(k)
       Url = WebBaseUrl & "/api/sqlhub/exec?appid=md_755&configName=GetProductFeatures"
       Call getHUBdata(Url, postJson, jsonResponse)
    
   
    'Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
     Set Data = jsonResponse("data")
     Product_Feature_Number = Data.count

     ReDim Product_Feature_IDs(1 To Product_Feature_Number)
     ReDim Product_Feature_Codes(1 To Product_Feature_Number)
     ReDim Product_Feature_Names(1 To Product_Feature_Number)
     ReDim Product_Feature_NamesE(1 To Product_Feature_Number)
     For i = 1 To Product_Feature_Number
          Product_Feature_IDs(i) = Data(i)(HubKeyProductFeatureId())
          Product_Feature_Codes(i) = Data(i)(HubKeyProductFeatureCode())
          Product_Feature_Names(i) = Data(i)(HubKeyProductFeatureNameZh())
          Product_Feature_NamesE(i) = Data(i)(HubKeyProductFeatureNameEn())
     Next i
  '======
    Dim ScreenWidth, ScreenHeight As Long
    Dim UFwidth, UFheight As Long
    Dim Defwidth, Defheight As Long
    
    ScreenWidth = Application.Width
    ScreenHeight = Application.Height
    Defwidth = 0.17 * ScreenWidth * 0.8 * 2
    Defheight = 0.043 * ScreenHeight * 0.9
    TopPos = 0.09 * ScreenHeight * 0.9

    
    'TopPos = 60
  
    For i = 1 To Product_Feature_Number
        Set Chk = UF.Controls.add("Forms.CheckBox.1", Product_Feature_Names(i), True)
        With Chk
            .Caption = i & "|" & Product_Feature_Names(i) & "|" & Product_Feature_NamesE(i)
            .Tag = Cat_Large_ID(k) & "|" & Product_Feature_IDs(i)
            If i Mod 20 = 0 Then
              Cn = Cn + 1
             .Top = TopPos
            
            Else
             .Top = TopPos + (i Mod 20) * Defheight  ' 25
              '.Top = 18 * i + TopPos
            End If
             .Left = 20 + Cn * Defwidth ' 300
           
            .Width = Defwidth '300 ' " & UText(35774, 32622, 22797, 36873, 26694, 30340, 23485, 24230) & "
            .Height = Defheight ' 25
            .Visible = True
            .Font.Size = 10
        End With
    Next i
    
    W(k) = W(k) + (Cn + 1) * Defwidth '300
    
    If N < Product_Feature_Number Then N = Product_Feature_Number
    Set Chk = UF.Controls.add("Forms.TextBox.1", "TextBox1", True)
     With Chk
        .Top = Int(0.05 * ScreenHeight * 0.9) '40
        .Left = 20 + W(k - 1)
        .Value = Cat_Large(k) & "|" & Cat_LargeE(k)
        .Width = W(k) ' " & UText(35774, 32622, 22797, 36873, 26694, 30340, 23485, 24230) & "
        .Height = Defheight
       .Visible = True
     End With
     
   Cn = Cn + 1
   Next k
   
     UF.Width = 20 + (Cn) * Defwidth + 40  '(k - 1) * 300 + 40
     UF.Height = (N + 1) * Defheight
     UF.StartUpPosition = 2
    UF.Show vbModal
    
100 End Sub

Sub Get_CheckPoints()
    
    Dim i, Cn, k As Integer
    Dim N As Integer
    Dim postJson As Object
    Dim jsonResponse As Object
    Dim Url As String
    Dim Prarr As Variant

    Dim Data As Object
    Dim Product_IDs() As String
    Dim Product_Codes() As String
    Dim Product_Names() As String
    Dim Product_NamesE() As String
    Dim Product_Number As Integer
    
    Dim UF As Product_Lev4
    Set UF = New Product_Lev4
    Dim LeftPos As Integer
    Dim TopPos As Integer
    Dim Chk As Object
    Dim gotChks As Integer
    Dim gotList() As String
    Dim p() As Variant
    Dim ky As String
    Dim TPN As Integer
    Set postJson = New Scripting.Dictionary
    Dim dicList As Object
    Set dicList = New Scripting.Dictionary
    N = 0
    For k = 1 To SLTN
     Prarr = Split(Product_ID(k), "|")
     postJson("productLargeType") = Prarr(0)
     postJson("productMiddleType") = Prarr(1)
     postJson("productSmallType") = Prarr(2)
     postJson("productName") = Prarr(3)
     For i = 1 To N_Cat_Large
        If Prarr(0) = Cat_Large_ID(i) Then
            postJson("productFeatures_G") = ChT_ID(i)
        End If
     Next i
     
     Url = WebBaseUrl & "/api/OnSiteTestProject/getCheckPoints"
     Call getHUBdata(Url, postJson, jsonResponse)
    
      Dim jasonstr As String
      jasonstr = JsonConverter.ConvertToJson(jsonResponse)
     'ThisDocument.Tables(2).cell(1, 1).Range.Text = jasonstr
      'Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
      
       Set Data = jsonResponse("data")
       gotChks = Data.count
     
       ReDim gotList(1 To gotChks)
       ReDim p(1 To gotChks)
   
       For i = 1 To gotChks
         ky = Data(i)(HubKeyCheckPointCode())
         If Language = "Chinese" Then
           dicList.Item(ky) = Data(i)(HubKeyCheckPointCode()) & "|" & Data(i)(HubKeyCheckPointName()) & "|" & Data(i)(HubKeyAccordingStandardType()) & "|" & _
           Data(i)(HubKeyRequirement()) & "|" & Data(i)(HubKeySampleQty()) & "|" & Data(i)(HubKeyAcceptanceQty()) & "|" & Data(i)(HubKeyRejectQty()) & "|" & _
           Data(i)(HubKeyQuantityType()) & "|" & Data(i)(HubKeyProductFeature())
        Else
           dicList.Item(ky) = Data(i)(HubKeyCheckPointCode()) & "|" & Data(i)(HubKeyCheckPointNameEn()) & "|" & Data(i)(HubKeyAccordingStandardType()) & "|" & _
           Data(i)(HubKeyRequirementEn()) & "|" & Data(i)(HubKeySampleQtyEn()) & "|" & Data(i)(HubKeyAcceptanceQtyEn()) & "|" & Data(i)(HubKeyRejectQtyEn()) & "|" & _
           Data(i)(HubKeyQuantityType()) & "|" & Data(i)(HubKeyProductFeature())
        
        End If
         
       Next i
       TPN = TPN + dicList.count
    
    '（0）" & HubKeyCheckPointCode() & "-（1）" & HubKeyCheckPointName() & "-（2)" & HubKeyAccordingStandardType() & "-（3）" & HubKeyRequirement() & "-（4）" & HubKeySampleQty() & "-（5）" & HubKeyAcceptanceQty() & "-（6）" & HubKeyRejectQty() & "-（7）" & HubKeyQuantityType() & "-（8）" & HubKeyProductFeature() & "
    

10   Next k
    ReDim p(1 To TPN)
    
    For Each i In dicList.Keys
        N = N + 1
        p(N) = dicList.Item(i)
       ' Debug.Print n & ": " & dicList.Item(i)
    Next i
      'Debug.Print "n=" & n
      
      Call Input_CheckPoint(N, p())
  
100 End Sub
Sub Load_Product_PrimaryTypeCheckPoints()
    
   Dim i, Cn, k As Integer
    Dim N As Integer
    Dim postJson As Object
    Dim jsonResponse As Object
    Dim Url As String
    Dim UF As Product_PrimaryTypeCheckPoints
    Set UF = New Product_PrimaryTypeCheckPoints
    Dim Data As Object
    Dim ChkCode(), ChkName(), ChkNameE() As String
    Dim Chk As Object
    Dim gotChks As Integer
 
    Dim pShow() As Variant
    Dim ch
    Dim ky As String
    Dim TPN As Integer
    Set postJson = New Scripting.Dictionary
    Dim dicList As Object
    Set dicList = New Scripting.Dictionary
    Dim dicListShow As Object
    Set dicListShow = New Scripting.Dictionary
     
     postJson("productLargeType") = Cat_Large_ID(1)
     Url = WebBaseUrl & "/api/sqlhub/exec?appid=md_551&configName=getPrimaryTypeCheckPoints"
   ' URL = WebBaseUrl & "/api/OnSiteTestProject/getCheckPoints"
     Call getHUBdata(Url, postJson, jsonResponse)
    
      Dim jasonstr As String
      jasonstr = JsonConverter.ConvertToJson(jsonResponse)
       'ThisDocument.Tables(2).cell(1, 1).Range.Text = jasonstr
      'Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
      
       Set Data = jsonResponse("data")
       gotChks = Data.count
 
       For i = 1 To gotChks
         ky = Data(i)(HubKeyCheckPointCode())
         If Language = "Chinese" Then
           dicList.Item(ky) = Data(i)(HubKeyCheckPointCode()) & "|" & Data(i)(HubKeyCheckPointName()) & "|" & Data(i)(HubKeyAccordingStandardType()) & "|" & _
           Data(i)(HubKeyRequirement()) & "|" & Data(i)(HubKeySampleQty()) & "|" & Data(i)(HubKeyAcceptanceQty()) & "|" & Data(i)(HubKeyRejectQty()) & "|" & _
           Data(i)(HubKeyQuantityType()) & "|" & Data(i)(HubKeyProductFeature())
        Else
           dicList.Item(ky) = Data(i)(HubKeyCheckPointCode()) & "|" & Data(i)(HubKeyCheckPointNameEn()) & "|" & Data(i)(HubKeyAccordingStandardType()) & "|" & _
           Data(i)(HubKeyRequirementEn()) & "|" & Data(i)(HubKeySampleQtyEn()) & "|" & Data(i)(HubKeyAcceptanceQtyEn()) & "|" & Data(i)(HubKeyRejectQtyEn()) & "|" & _
           Data(i)(HubKeyQuantityType()) & "|" & Data(i)(HubKeyProductFeature()) & "|"
        End If
         dicListShow(ky) = Data(i)(HubKeyCheckPointCode()) & "|" & Data(i)(HubKeyCheckPointName()) & "|" & Data(i)(HubKeyCheckPointNameEn())
       Next i
       TPN = TPN + dicList.count
    
    '（0）" & HubKeyCheckPointCode() & "-（1）" & HubKeyCheckPointName() & "-（2)" & HubKeyAccordingStandardType() & "-（3）" & HubKeyRequirement() & "-（4）" & HubKeySampleQty() & "-（5）" & HubKeyAcceptanceQty() & "-（6）" & HubKeyRejectQty() & "-（7）" & HubKeyQuantityType() & "-（8）" & HubKeyProductFeature() & "
    
       ReDim pShow(1 To gotChks)
       ReDim ChkCode(1 To gotChks)
       ReDim ChkName(1 To gotChks)
       ReDim ChkNameE(1 To gotChks)
       
       ReDim PriChk(1 To TPN)
    N = 0
    For Each i In dicList.Keys
        N = N + 1
        PriChk(N) = dicList.Item(i)
        pShow(N) = dicListShow.Item(i)
        ch = Split(pShow(N), "|")
        ChkCode(N) = ch(0)
        ChkName(N) = ch(1)
        ChkNameE(N) = ch(2)
    Next i
    gotChks = N
    '====
   
    Dim ScreenWidth, ScreenHeight As Long
    Dim UFwidth, UFheight As Long
    Dim Defwidth, Defheight As Long
    
    ScreenWidth = Application.Width
    ScreenHeight = Application.Height
    Defwidth = 0.17 * ScreenWidth * 0.8
    Defheight = 0.04 * ScreenHeight * 0.9
    
    TopPos = 0.092 * ScreenHeight * 0.9
    LeftPos = 20
    Cn = 0
    For i = 1 To gotChks
        Set Chk = UF.Controls.add("Forms.CheckBox.1", ChkName(i), True)
        With Chk
            .Caption = i & "|" & ChkName(i) & "|" & ChkNameE(i)
            .Tag = ChkCode(i)
            If i Mod 20 = 0 Then
              Cn = Cn + 1
             .Top = TopPos + 5
            Else
             .Top = TopPos + (i Mod 20) * Defheight + 5
         
            End If
             .Left = LeftPos + Cn * Defwidth
           
            .Width = Defwidth
            .Height = Defheight
            .Visible = True
         .Font.Size = 10
        End With
    Next i
       
    Set Def = UF.Controls.add("Forms.TextBox.1", "TextBox1", True)
     With Def
        .Top = 30
        .Left = LeftPos
        .Value = Cat_Large(1) & "|" & Cat_LargeE(1)
        .Width = (Cn + 1) * Defwidth
        .Height = Defheight
        .Visible = True
    End With
     
'=====
    
    UF.Width = LeftPos + (Cn + 1) * Defwidth + 40 '(k - 1) * 300 + 40
    
     UF.Height = TopPos + (20 + 2 + 1) * Defheight
  
  
     UF.Show vbModal
    
  
100 End Sub
Sub Load_Product_ForDefects()
    
   Dim i, Cn, k As Integer
    Dim N As Integer
    Dim UF As Product_ForDefects
    Set UF = New Product_ForDefects
    Dim LeftPos As Integer
    Dim TopPos As Integer
    Dim Chk As Object
 
    SearchDefect = True
   
    
    Dim ScreenWidth, ScreenHeight As Long
    Dim UFwidth, UFheight As Long
    Dim Defwidth, Defheight As Long
    
    ScreenWidth = Application.Width
    ScreenHeight = Application.Height
    Defwidth = 0.17 * ScreenWidth * 0.8 * 2
    Defheight = 0.04 * ScreenHeight * 0.9
    TopPos = 0.09 * ScreenHeight * 0.9
    LeftPos = 20
    Cn = 0
    
  

    If SLTN = 0 Then
       SLTN = Val(GetPropertyValue("Selected Product Number"))
     End If

  If SLTN = 0 Then
         Call Load_Product_Lev1
    Else
  
       For i = 1 To SLTN
       If ProductE(i) = "" Then ProductE(i) = GetPropertyValue("ProductE" & " " & i)
       If Product(i) = "" Then Product(i) = GetPropertyValue("Product" & " " & i)
       If Product_ID(i) = "" Then Product_ID(i) = GetPropertyValue("Product_ID" & " " & i)
      
       Set Chk = UF.Controls.add("Forms.CheckBox.1", ProductE(i), True)
        With Chk
            .Caption = i & "|" & Product(i) & "|" & ProductE(i)
            .Tag = Product_ID(i)
            .Top = TopPos + i * Defheight
            .Left = LeftPos
            .Width = Defwidth   '300
            .Height = Defheight '25
            .Visible = True
            .Enabled = True
        End With
       Next i
   End If
     
     UF.Width = LeftPos + i * Defwidth + 40
     
     UF.Height = TopPos + (i + 2 + 2) * Defheight

   
   If SLTN = 1 Then
        Defect_Product_ID = Product_ID(1)
        Load_Product_DefectList
   ElseIf SLTN > 1 Then
        UF.Show vbModal
      
  End If


  
100 End Sub
Sub Load_Product_DefectList()
    
    Dim i, Cn, k As Integer
    Dim N As Integer
    Dim postJson As Object
    Dim jsonResponse As Object
    Dim Url As String
    Dim Prarr As Variant
     Dim UF As Product_DefectList
     Set UF = New Product_DefectList
  
    Dim LeftPos As Integer
    Dim TopPos As Integer
    Dim Def As Object
    Dim gotDefects As Integer
    Dim gotList As String
    Dim p() As String
    Dim TPN As Integer
    Set postJson = New Scripting.Dictionary
    Dim dicList As Object
    Set dicList = New Scripting.Dictionary
    Dim Ps As Integer
    Dim Str As String
    Dim Defects(), DefectsE(), DefectCodes() As String
    Dim DF
    Dim Chk As Object
   
   '----

     Prarr = Split(Defect_Product_ID, "|")
     postJson("productLargeType") = Prarr(0)
     postJson("productMiddleType") = Prarr(1)
     postJson("productSmallType") = Prarr(2)
     postJson("productName") = Prarr(3)
     
     Url = WebBaseUrl & "/api/OnSiteTestProject/getDefect"
     Call getHUBdata(Url, postJson, jsonResponse)
    
      Dim jasonstr As String
      jasonstr = JsonConverter.ConvertToJson(jsonResponse)
      'ThisDocument.Tables(2).cell(1, 1).Range.Text = jasonstr
      'Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
      
       Set Data = jsonResponse("data")
       gotDefects = Data.count
       
        If gotDefects = 0 Then
         MsgBox "HUB" & UText(19978, 26410, 25214, 21040, 35813, 20135, 21697, 30340, 32570, 38519, 25551, 36848, 21442, 25968) & "，" & UText(32570, 38519, 25968, 25454, 26080, 27861, 23548, 20837) & "，" & UText(35831, 32852, 31995, 36136, 25511, 37096, 25110) & "IT" & UText(37096, 23547, 27714, 35299, 20915, 26041, 26696) & CHR(10) & _
         "The defect description for this product is not found at HUB, the defect is not imported, please contact the Quality Control Department or IT Deparment for the solutions", vbCritical
          Exit Sub
      
       End If
 
       ReDim Defects(1 To gotDefects)
       ReDim DefectsE(1 To gotDefects)
       ReDim DefectCodes(1 To gotDefects)
   
       For i = 1 To gotDefects
           Defects(i) = Data(i)(HubKeyDefectName())
           DefectsE(i) = Data(i)(HubKeyDefectNameEn())
           DefectCodes(i) = Data(i)(HubKeyDefectCode())
           dicList.Item(DefectCodes(i)) = DefectCodes(i) & "|" & Defects(i) & "|" & DefectsE(i)
       Next i
     '====
       TPN = TPN + dicList.count
       
    ReDim p(1 To TPN)
    N = 0
    For Each i In dicList.Keys
        N = N + 1
        p(N) = dicList.Item(i)
    Next i
    gotDefects = N
    
    Call SortDefects(gotDefects, p)
    
    For i = 1 To N
        DF = Split(p(i), "|")
        DefectCodes(i) = DF(0)
        Defects(i) = DF(1)
        DefectsE(i) = DF(2)
    Next i
 
    Dim ScreenWidth, ScreenHeight As Long
    Dim UFwidth, UFheight As Long
    Dim Defwidth, Defheight As Long
    
    ScreenWidth = Application.Width
    ScreenHeight = Application.Height
    Defwidth = 0.136 * ScreenWidth            'ScreenHeight=1091
    Defheight = 0.036 * ScreenHeight
    TopPos = 0.05 * ScreenHeight               'ScreenHeight=650
    LeftPos = 0.018 * ScreenWidth
    
     
    
     Set Def = UF.Controls.add("Forms.TextBox.1", "TextBox1", True)
     With Def
        .Top = TopPos
        .Left = LeftPos
        .Value = "Selected Product=" & Defect_Product & "|" & Defect_ProductE
        .Width = (Int(N / 20) + 1) * Defwidth ' " & UText(35774, 32622, 22797, 36873, 26694, 30340, 23485, 24230) & "
        .Height = Defheight
        .Visible = True
    End With
     
     'Dim cmb As MSForms.ComboBox
     'Dim LastText As String
     
     'Set cmb = UF.Controls.add("Forms.ComBox.1", "ComBox1", True)
    ' With cmb
       ' .AddItem ""
       ' For i = 1 To gotDefects
           ' .AddItem p(i)
        'Next i
          '  .Top = TopPos + Defheight + 2
           ' .Left = LeftPos
           ' .Width = 0.2 * ScreenWidth    '300
          '  .Height = Defheight '50
           ' .Visible = True
           ' .Enabled = True
           ' .Style = fmStyleDropDownCombo
           ' .ListRows = 20
          '  .DropDown
    'End With
   ' LastText = cmb.Text
   
     'GoTo 10
    ' Set Chk = UF.Controls.add("Forms.TextBox.1", "KeyWords", True)
     'With Chk
            '.Caption = "KeyWords"
         '   .Tag = "KeyWords"
           ' .Top = TopPos + Defheight + 2
          '  .Left = LeftPos
            '.Width = 0.2 * ScreenWidth    '300
            '.Height = Defheight '50
            '.Visible = True
            '.Enabled = True
        'End With
        
   
10       TopPos = TopPos + Defheight + 2
        
     
'====
    Cn = 0
 
    For i = 1 To gotDefects
        Set Def = UF.Controls.add("Forms.CheckBox.1", Defects(i), True)
        With Def
            .Caption = i & "|" & Defects(i) & "|" & DefectsE(i)
            .Tag = DefectCodes(i)
            If i Mod 20 = 0 Then
              Cn = Cn + 1
             .Top = TopPos + Defheight
            Else
              .Top = TopPos + (i Mod 20 + 1) * Defheight 'Int((UFheight - TopPos) / 21) '25
         
            End If
            .Left = LeftPos + Cn * Defwidth   'Int(UFwidth / 6) '150
            .Width = Defwidth  '150 ' " & UText(35774, 32622, 22797, 36873, 26694, 30340, 23485, 24230) & "
            .Height = Defheight '25
            .Visible = True
            .Font.Size = 10
        End With
    Next i
   
     
'=====
    UF.Width = LeftPos + (Cn + 1) * Defwidth + 40
    UF.Height = ScreenHeight * 0.9
    UF.StartUpPosition = 2  ' " & UText(23621, 20013, 26174, 31034) & "
   ' UF.Width = 20 + (cn + 1) * 150 + 40 '(k - 1) * 300 + 40
    'UF.Height = 620

    UF.Show vbModal
    
  
  
100 End Sub
Sub Get_Defects()
    
    Dim i, Cn, k As Integer
    Dim N As Integer
    Dim postJson As Object
    Dim jsonResponse As Object
    Dim Url As String
    Dim Prarr As Variant
    Dim UF As Product_ForDefects
    Set UF = New Product_ForDefects
    Dim LeftPos As Integer
    Dim TopPos As Integer
    Dim Def As Object
    Dim gotDefects As Integer
    Dim gotList As String
    Dim ky As String
    Dim TPN As Integer
    Set postJson = New Scripting.Dictionary
    Dim dicList As Object
    Set dicList = New Scripting.Dictionary
    Dim Ps As Integer
    Dim Str As String
    
 
    
    
             Defect_Product = UF.Controls(i).Name
             Str = UF.Controls(i).Caption
             Defect_Product = Mid(Str, InStr(1, Str, "|") + 1, InStrRev(Str, "|") - InStr(1, Str, "|") - 1)
             Defect_ProductE = Mid(Str, InStrRev(Str, "|") + 1, Len(Str) - InStrRev(Str, "|"))
             Defect_Product_ID = UF.Controls(i).Tag

   '----
     Prarr = Split(Defect_Product_ID, "|")
     postJson("productLargeType") = Prarr(0)
     postJson("productMiddleType") = Prarr(1)
     postJson("productSmallType") = Prarr(2)
     postJson("productName") = Prarr(3)
     
     Url = WebBaseUrl & "/api/OnSiteTestProject/getDefect"
     Call getHUBdata(Url, postJson, jsonResponse)
    
      Dim jasonstr As String
      jasonstr = JsonConverter.ConvertToJson(jsonResponse)
       ThisDocument.Tables(2).Cell(1, 1).Range.Text = jasonstr
      'Set jsonResponse = JsonConverter.ParseJson(ReqResponse)
      
       Set Data = jsonResponse("data")
       gotDefects = Data.count
     
       ReDim gotList(1 To gotDefects)
       ReDim p(1 To gotChks)
   
       For i = 1 To gotDefects
         ky = Data(i)(HubKeyDefectCode())
         If Language = "Chinese" Then
           dicList.Item(ky) = Data(i)(HubKeyDefectCode()) & "|" & Data(i)(HubKeyDefectName()) & "|" & Data(i)(HubKeyDefectLevel())
        Else
           dicList.Item(ky) = Data(i)(HubKeyDefectCode()) & "|" & Data(i)(HubKeyDefectNameEn()) & "|" & Data(i)(HubKeyDefectLevel())
        End If
         
       Next i
       TPN = TPN + dicList.count
'====

    Cn = 0
    LeftPos = 20
    TopPos = 50
   
    
    If SLTN = 0 Then
       SLTN = Val(GetPropertyValue("Selected Product Number"))
     End If
    
    If SLTN = 0 Then
         Call Load_Product_Lev1
    Else
       For i = 1 To SLTN
       If ProductE(i) = "" Then ProductE(i) = GetPropertyValue("ProductE" & " " & i)
       If Product(i) = "" Then Product(i) = GetPropertyValue("Product" & " " & i)
       If Product_ID(i) = "" Then Product_ID(i) = GetPropertyValue("Product_ID" & " " & i)
      
       Set Chk = UF.Controls.add("Forms.CheckBox.1", ProductE(i), True)
        With Chk
            .Caption = i & "|" & Product(i) & "|" & ProductE(i)
            .Tag = Product_ID(i)
            .Top = TopPos + i * 25
            .Left = 20
            .Width = 300
            .Height = 25
            .Visible = True
            .Enabled = True
        End With
       Next i
   End If
  UF.Show vbModal
  
  
100 End Sub
Sub VerifyVersion(rpn, Q)
Dim rg As Range
Dim Tc As Integer
Dim Tr, Vs As String
Dim Trn, HTrn As Integer
Dim HTrnText, HVsnText As String
Dim Vsn, HVsn As Single
Dim postJson As Object
Dim Pcn As Integer
Dim ch As VbMsgBoxResult

     Set WD = ThisDocument
     Set rg = WD.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Range
     Pcn = rg.Cells.count
     rpn = rg.Cells(Pcn).Range.Text
     rpn = Mid(rpn, 1, Len(rpn) - 2)
     rpn = Trim(Replace(rpn, "Report No: ", ""))
     
     Set rg = WD.Sections(1).Footers(wdHeaderFooterFirstPage).Range
     Tc = rg.Tables.count
     Tr = rg.Tables(Tc).Cell(1, 1).Range.Text
     Vs = rg.Tables(Tc).Cell(1, 2).Range.Text
    
     If InStr(Tr, "-") = 0 Then MsgBox "The report template without the document No. or the format not comply with requirement, the progaram can not deal with it": Exit Sub
     
     Do While Val(Tr) = 0
      Tr = Mid(Tr, InStr(Tr, "-") + 1)
     Loop
     Trn = Val(Tr)
     
     Do While Val(Vs) = 0
      Vs = Mid(Vs, 2)
     Loop
     Vsn = Val(Vs)
     

     'rpn = "H24096737HF-A1"
         Set postJson = New Scripting.Dictionary
          postJson("ReportNumber") = rpn
                             
         Url = WebBaseUrl & "/api/SelectFileEncoding/EncodingValidation"

         Call getHUBdata(Url, postJson, jsonResponse)
         
           'Dim json_str As String
           'json_str = JsonConverter.ConvertToJson(jsonResponse)
          
          
          HTrnText = jsonResponse("FileEncoding")
          If HTrnText = "" Then
            Exit Sub
          Else
            Do While IsNumeric(Mid(HTrnText, 1, 1)) = False
              HTrnText = Mid(HTrnText, 2, Len(HTrnText) - 1)
            Loop
            HTrn = Val(HTrnText)
          
          HVsnText = jsonResponse("VersionNumber")
          Do While IsNumeric(Mid(HVsnText, 1, 1)) = False
              HVsnText = Mid(HVsnText, 2, Len(HVsnText) - 1)
          Loop
          HVsn = Val(HVsnText)
          
         
            
         If HTrn <> Trn Then ch = MsgBox("DocNo.of this report template is" & Trn & _
         "and is not same with the statement of the inspection notice form on HUB, which shall be " & HTrn & _
         ". If you are sure this report template be right, click YES to continue otherwise cilck No " & _
         "to exit the program." & CHR(10) & UText(26412, 25253, 21578, 27169, 26495, 26159, 25991, 20214, 21495) & "DocNo." & UText(26159) & Trn & _
         UText(21516) & "HUB" & UText(26816, 39564, 36890, 30693, 21333, 19978, 26631, 27880, 30340) & HTrn & UText(19981, 21516) & "，" & UText(22914, 26524, 20320, 30830, 23450, 26412, 25253, 21578, 27169, 26495, 26159, 27491, 30830, 30340) & "，" & UText(25353, 26159, 32487, 32493) & "，" & UText(21542, 21017, 25353, 21542, 36864, 20986, 31243, 24207), vbYesNo)
          If ch = vbNo Then Q = "Quit"
         
          
         If HVsn <> Vsn Then ch = MsgBox("Version numner of this report template is" & Vsn & _
         "and is not same with the newest Version numner on HUB, which shall be " & HVsn & _
         ". If you are sure Version numner of this report template be right, click YES to continue otherwise cilck No " & _
         "to exit the program." & CHR(10) & UText(26412, 25253, 21578, 27169, 26495, 30340, 29256, 26412, 21495, 26159) & Vsn & _
         UText(21516) & "HUB" & UText(19978, 26368, 26032, 30340, 29256, 26412, 21495) & HVsn & UText(19981, 21516) & "，" & UText(22914, 26524, 20320, 30830, 23450, 35813, 29256, 26412, 21495, 26159, 27491, 30830, 30340) & "，" & UText(25353, 26159, 32487, 32493) & "，" & UText(21542, 21017, 25353, 21542, 36864, 20986, 31243, 24207), vbYesNo)
          If ch = vbNo Then Q = "Quit"
      End If
End Sub

Sub FindAndHighlightText()
  
    Set WD = ThisDocument
 
    Dim Cel As Cell
    Dim Rng As Range
    Dim found As Boolean
    Dim Cn As Integer
    Dim searchText As String
    searchText = "(())"
    Cn = 0
   

    For Each tbl In WD.Tables
        For Each Cel In tbl.Range.Cells
            With Cel.Range.Find
                .Text = searchText
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
              
                found = .Execute()
                
                Do While found
      
                    Set Rng = .Parent
                    Rng.Font.Color = wdColorRed
                    found = .Execute()
                    Cn = Cn + 1
                    
                Loop
            End With
        Next Cel
    Next tbl
     If Cn > 0 Then
        MsgBox "Some " & searchText & " was found without filling content, we have highlighted them in red." & _
       "Please fill in content and change the color to balck. please don't delete the double parentheses (())" & _
       " especially when the report should translate to other language, other program will delete them automatically." & _
         vbCrLf & UText(19968, 20123) & searchText & UText(34987, 21457, 29616, 26410, 22635, 20805, 20869, 23481) & "，" & UText(24050, 29992, 32418, 33394, 26631, 20986) & "，" & UText(35831, 22635, 20889, 20869, 23481, 23558, 39068, 33394, 25913, 22238, 40657, 33394) & "," & UText(35831, 19981, 35201, 21024, 38500, 21452, 25324, 21495) & "(())，" & _
         UText(29305, 21035, 24403, 25253, 21578, 38656, 35201, 32763, 35793, 25104, 20854, 20182, 35821, 31181, 26102) & "，" & UText(20854, 20182, 31243, 24207, 20250, 33258, 21160, 21024, 38500, 20182, 20204) & "。", vbInformation
     End If
End Sub

Sub SetTablePosition()
  
    Dim Cell As Cell
    Dim Rng As Range
   
    Set WD = ThisDocument

    ' " & UText(35774, 32622, 39029, 38754, 24038, 36793, 36317) & "
    With WD.PageSetup
        .LeftMargin = 48
       ' .LeftMargin = CentimetersToPoints(1.6) ' 3" & UText(21400, 31859, 30340, 24038, 36793, 36317) & "
       ' .RightMargin = CentimetersToPoints(1) ' 3" & UText(21400, 31859, 30340, 21491, 36793, 36317) & "
    End With
  End Sub

Sub DeleteBookMark(Cel)

Dim BK As bookmark
If Cel.Range.Bookmarks.count > 0 Then
For Each BK In Cel.Range.Bookmarks
   BK.Delete
Next BK
End If
End Sub

Sub CellAddDropDownList(ByRef Cel As Cell)
Dim Occ As contentControl
             Cel.Range.Text = " "
             Cel.Range.Collapse Direction:=wdCollapseEnd
             On Error Resume Next
             Set Occ = Cel.Range.ContentControls.add(wdContentControlComboBox, Range:=Cel.Range)
       
                With Occ
                 '.Title = "Sample Size"
                 .Tag = "Sample Size"
                 .DropdownListEntries.add "Choose ", " "
                 .DropdownListEntries.add "sample", "sample"
                 .DropdownListEntries.add "samples", "samples"
                 .DropdownListEntries.add "sample/item", "sample/item"
                 .DropdownListEntries.add "samples/item", "samples/item"
                 .DropdownListEntries.add "all samples", "all samples"
             End With
End Sub


Sub Translate()
   'bywynwyn, 20241122，" & UText(29992) & "HUB" & UText(23383, 20856) & "
Dim Dict_Original_EnText() As String
Dim SDict_Original_EnText() As String
Dim Dict_Original_RuText() As String

Dim Dict_Inner_Text() As String
Dim Wd_Inner_Text() As String
Dim Trs_Inner_Text() As String

Dim Dict_Temp_Text As String
Dim Wd_Temp_Text As String
Dim Trs_Temp_Text As String

Dim Wd_Cell_Text As String

Dim Dict_Temp_Text_PosL() As Integer
Dim Dict_Temp_Text_PosR() As Integer
Dim Wd_Temp_Text_PosL() As Integer
Dim Wd_Temp_Text_PosR() As Integer
Dim Trs_Temp_Text_PosL() As Integer
Dim Trs_Temp_Text_PosR() As Integer

Set WD = ThisDocument

Dim tbl As table
Dim Cell As Cell
Dim exFilePath As String
Dim Rn, i, j, N As Long

Dim T, C, CI, RI As Integer
Dim TbTitle As String
Dim CountL, CountR As Integer
Dim Trs_Inner_Index() As Integer
Dim Path As String
Dim DocName As String
Dim BDocName As String
Dim DocExt As String
Dim Field As Field
Dim Displaytext As String
Dim fd As FileDialog
Dim S As String
Dim ch As VbMsgBoxResult
Dim SLanguage, DLanguage As String
Dim LangSuffix As String
Dim SelectedLang As String
       
      SLanguage = UText(33521, 25991, 25991, 26412)
      
      Path = WD.Path
      DocName = WD.Name
   
  ' " & UText(35753, 29992, 25143, 36873, 25321, 30446, 26631, 35821, 35328) & "
  SelectedLang = InputBox("Please select target language / " & UText(35831, 36873, 25321, 30446, 26631, 35821, 35328) & ":" & CHR(10) & _
    "1 - Russian / " & UText(20420, 25991) & CHR(10) & _
    "2 - Japanese / " & UText(26085, 25991) & CHR(10) & _
    "3 - Vietnamese / " & UText(36234, 21335, 25991) & CHR(10) & CHR(10) & _
    "Enter 1, 2, or 3 / " & UText(35831, 36755, 20837) & " 1、2 " & UText(25110) & " 3", _
    "Language Selection / " & UText(35821, 35328, 36873, 25321), "1")
  
  ' " & UText(22788, 29702, 29992, 25143, 21462, 28040, 25110, 26080, 25928, 36755, 20837) & "
  If SelectedLang = "" Then
      Exit Sub
  End If
  
  SelectedLang = Trim(SelectedLang)
  
  If SelectedLang = "1" Then
      DLanguage = UText(20420, 25991, 25991, 26412)
      LangSuffix = "-RU"
  ElseIf SelectedLang = "2" Then
      DLanguage = UText(26085, 25991, 25991, 26412)
      LangSuffix = "-JA"
  ElseIf SelectedLang = "3" Then
      DLanguage = UText(36234, 21335, 35821, 25991, 26412)
      LangSuffix = "-VI"
  Else
      MsgBox "Invalid selection. Translation cancelled." & CHR(10) & UText(26080, 25928, 36873, 25321) & "。" & UText(32763, 35793, 24050, 21462, 28040) & "。", vbExclamation
      Exit Sub
  End If

  ch = MsgBox("Is it necessary to translate the report into " & _
    IIf(SelectedLang = "1", "Russian", IIf(SelectedLang = "2", "Japanese", "Vietnamese")) & "?" & CHR(10) & _
    UText(38656, 35201, 23558, 27492, 25253, 21578, 32763, 35793, 25104) & _
    IIf(SelectedLang = "1", UText(20420, 25991), IIf(SelectedLang = "2", UText(26085, 25991), UText(36234, 21335, 25991))) & UText(21527) & "?", vbYesNo)

  If ch = 6 Then

    DocName = WD.Name
    BDocName = Left(DocName, InStrRev(DocName, ".") - 1)
    DocExt = Mid(DocName, InStrRev(DocName, ".") + 1)
    WD.Save
       
     Application.StatusBar = "Translating, pease wait...., " & "  " & UText(32763, 35793, 20013) & "，" & UText(35831, 31245, 20505) & "...."
15   '----------------------------" & UText(22312, 19979, 20351, 29992) & "HUB" & UText(23383, 20856, 20570, 20026, 27491, 24335, 36816, 34892, 30340, 20195, 30721) & "
    
      Url = WebBaseUrl & "/api/sqlhub/exec?appid=md_763&configName=SelectReportTextTranslationDictionaryLedger"
      Call getHUBdata(Url, postJson, jsonResponse)
    'Dim jsondata As String
    'jsondata = JsonConverter.ConvertToJson(jsonResponse)
    'Debug.Print jsondata
    Set Data = jsonResponse("data")
    Rn = Data.count
    ReDim Dict_Original_EnText(1 To Rn)
    ReDim SDict_Original_EnText(1 To Rn)
    ReDim Dict_Original_RuText(1 To Rn)
    For i = 1 To Rn
          Dict_Original_EnText(i) = Data(i)(SLanguage)
          ' " & UText(23581, 35797, 33719, 21462, 32763, 35793, 25991, 26412) & "，" & UText(22914, 26524, 23383, 27573, 19981, 23384, 22312, 25110, 20026, 31354, 21017, 20351, 29992, 33521, 25991, 25991, 26412) & "，" & UText(36991, 20813, 32763, 35793, 21518, 23383, 27573, 21464, 25104, 31354, 30333) & "
          Dim TransText As Variant
          Dim FieldExists As Boolean
          FieldExists = True
          TransText = ""
          On Error Resume Next
          Err.Clear
          TransText = Data(i)(DLanguage)
          If Err.Number <> 0 Then
              FieldExists = False
              Err.Clear
          End If
          On Error GoTo 0
          ' " & UText(21028, 26029, 26159, 21542, 20351, 29992, 32763, 35793, 25991, 26412) & "：" & UText(23383, 27573, 24517, 39035, 23384, 22312, 19988, 20540, 19981, 20026, 31354) & "
          If Not FieldExists Then
              ' " & UText(23383, 27573, 19981, 23384, 22312) & "，" & UText(20351, 29992, 33521, 25991, 25991, 26412) & "
              Dict_Original_RuText(i) = Data(i)(SLanguage)
          ElseIf IsNull(TransText) Then
              ' " & UText(23383, 27573, 20540, 20026) & "Null，" & UText(20351, 29992, 33521, 25991, 25991, 26412) & "
              Dict_Original_RuText(i) = Data(i)(SLanguage)
          ElseIf Trim(CStr(TransText)) = "" Then
              ' " & UText(23383, 27573, 20540, 20026, 31354, 23383, 31526, 20018) & "，" & UText(20351, 29992, 33521, 25991, 25991, 26412) & "
              Dict_Original_RuText(i) = Data(i)(SLanguage)
          Else
              ' " & UText(23383, 27573, 23384, 22312, 19988, 26377, 20540) & "，" & UText(20351, 29992, 32763, 35793, 25991, 26412) & "
              Dict_Original_RuText(i) = CStr(TransText)
          End If
          Dict_Temp_Text = Dict_Original_EnText(i)
          If InStr(1, Dict_Temp_Text, "((") > 0 Then
             Dict_Temp_Text_PosL = PositionsL(Dict_Temp_Text)
             Dict_Temp_Text_PosR = PositionsR(Dict_Temp_Text)
             CountL = UBound(Dict_Temp_Text_PosL)
             CountR = UBound(Dict_Temp_Text_PosR)
              If CountL <> CountR Then
                msg = msg + "Within Parameter of " & Dict_Temp_Text & ", the number of '((' does not match the number of '))'" & CHR(10) & _
               UText(21442, 25968, 34920, 20013, 30340) & Dict_Temp_Text & "'((’" & UText(30340, 25968, 37327, 21644) & "'))'" & UText(25968, 37327, 19981, 21305, 37197) & "。" & CHR(10)
                GoTo 10
              End If
              ReDim Dict_Inner_Text(1 To CountR)
              For j = 1 To CountR
                Dict_Inner_Text(j) = Mid(Dict_Temp_Text, Dict_Temp_Text_PosL(j), Dict_Temp_Text_PosR(j) + 2 - Dict_Temp_Text_PosL(j))
                
              Next j
              For j = 1 To CountR
                Dict_Temp_Text = Replace(Dict_Temp_Text, Dict_Inner_Text(j), "")
              Next j
              SDict_Original_EnText(i) = Dict_Temp_Text
            Else
              SDict_Original_EnText(i) = Dict_Original_EnText(i)
            End If
10    Next i
   
  
    With WD
      T = 0
    
      For Each tbl In WD.Tables
        T = T + 1
         C = 0
        For Each Cell In tbl.Range.Cells
          C = C + 1
          If Cell.Range.Fields.count > 0 Then
             If Cell.Range.Fields(1).Type = wdFieldMacroButton Then
               Wd_Cell_Text = Cell.Range.Fields(1).Code
               Wd_Cell_Text = Trim(Replace(Wd_Cell_Text, "MACROBUTTON", ""))
               Wd_Cell_Text = Mid(Wd_Cell_Text, InStr(1, Wd_Cell_Text, " ") + 1, Len(Wd_Cell_Text) - InStr(1, Wd_Cell_Text, " "))
     
             Else
                Displaytext = ""
                For Each Field In Cell.Range.Fields
                   If Field.Type = wdFieldHyperlink Then
                      Displaytext = Displaytext & Field.result.Text
                   Else
                      Displaytext = Displaytext & Field.Code
                   End If
                Next Field
                   Wd_Cell_Text = Displaytext
             End If
          ElseIf Cell.Range.ContentControls.count > 0 Then
             Wd_Cell_Text = Mid(Cell.Range.Text, 1, Len(Cell.Range.Text) - 1)
          Else
            Wd_Cell_Text = Mid(Cell.Range.Text, 1, Len(Cell.Range.Text) - 2)
            If Wd_Cell_Text = "" Then GoTo 20
          End If

12       If C = 1 Then TbTitle = Wd_Cell_Text
         If InStr(1, Wd_Cell_Text, "((") > 0 Or InStr(1, Wd_Cell_Text, "))") > 0 Then
             Wd_Temp_Text = Wd_Cell_Text

             Wd_Temp_Text_PosL = PositionsL(Wd_Temp_Text)
             Wd_Temp_Text_PosR = PositionsR(Wd_Temp_Text)
             CountL = UBound(Wd_Temp_Text_PosL)
             CountR = UBound(Wd_Temp_Text_PosR)
    
             If CountL <> CountR Then
                CI = Cell.ColumnIndex
                RI = Cel.rowIndex
                msg = msg + "For Tables " & i & ", Title : " & TbTitle & ", cell of Rows Number " & RI & ", Columns number " & CI & ",the '((' number does not match '))'" & CHR(10) & _
               UText(34920, 26684) & i & "," & UText(26631, 39064) & ":" & TbTitle & UText(31532) & RI & UText(34892) & "，" & CI & UText(21015, 21333, 20803, 26684, 20013, 30340) & "'(('" & UText(25968, 37327, 21644) & "‘))'" & UText(25968, 37327, 19981, 21305, 37197) & "。"
               GoTo 20
             End If
             ReDim Wd_Inner_Text(1 To CountR)
             For j = 1 To CountR
                 Wd_Inner_Text(j) = Mid(Wd_Temp_Text, Wd_Temp_Text_PosL(j), Wd_Temp_Text_PosR(j) + 2 - Wd_Temp_Text_PosL(j))

             Next j
             For j = 1 To CountR
                Wd_Temp_Text = Replace(Wd_Temp_Text, Wd_Inner_Text(j), "")
             Next j
             SWd_Cell_Text = Wd_Temp_Text

             For i = 2 To Rn
               If SWd_Cell_Text = SDict_Original_EnText(i) Then
                   Trs_Temp_Text = Dict_Original_RuText(i)
                   ' " & UText(22914, 26524, 32763, 35793, 25991, 26412, 20026, 31354) & "，" & UText(21017, 20445, 30041, 21407, 25991) & "，" & UText(36991, 20813, 21333, 20803, 26684, 34987, 28165, 31354) & "
                   If Trim(Trs_Temp_Text) = "" Then
                       GoTo 20
                   End If
                   Trs_Temp_Text_PosL = PositionsL(Trs_Temp_Text)
                   Trs_Temp_Text_PosR = PositionsR(Trs_Temp_Text)
                   CountL = UBound(Trs_Temp_Text_PosL)
                   CountR = UBound(Trs_Temp_Text_PosR)
                   If CountL <> CountR Then
                      msg = msg + "Within Parameter of " & Trs_Temp_Text & ", the number of '((' does not match the number of '))'" & CHR(10) & _
                      UText(21442, 25968, 34920, 20013, 30340) & Trs_Temp_Text & "'(('" & UText(30340, 25968, 37327, 21644) & "'))'" & UText(25968, 37327, 19981, 21305, 37197) & "。" & CHR(10)
                      GoTo 20
                   End If

                   ReDim Trs_Inner_Text(1 To CountR)
                   ReDim Trs_Inner_Index(1 To CountR)
                   
                   For j = 1 To CountR
                     Trs_Inner_Text(j) = Mid(Trs_Temp_Text, Trs_Temp_Text_PosL(j), Trs_Temp_Text_PosR(j) + 2 - Trs_Temp_Text_PosL(j))
                     If Val(Mid(Trs_Inner_Text(j), 3, Len(Trs_Inner_Text(j)) - 2)) > 0 Then
                       Trs_Inner_Index(j) = Val(Mid(Trs_Inner_Text(j), 3, Len(Trs_Inner_Text(j)) - 2))
                     Else
                      Trs_Inner_Index(j) = j
                     End If
                   Next j

                   For j = 1 To CountR
                     Trs_Temp_Text = Replace(Trs_Temp_Text, Trs_Inner_Text(j), Wd_Inner_Text(Trs_Inner_Index(j)))
                   Next j
               Cell.Range.Text = Trs_Temp_Text
               Cell.Range.Font.Name = "Times New Roman"
               GoTo 20
              End If
            Next i
         Else

            For i = 2 To Rn
              If Wd_Cell_Text = SDict_Original_EnText(i) Then
                  ' " & UText(22914, 26524, 32763, 35793, 25991, 26412, 20026, 31354) & "，" & UText(21017, 20445, 30041, 21407, 25991) & "，" & UText(36991, 20813, 21333, 20803, 26684, 34987, 28165, 31354) & "
                  If Trim(Dict_Original_RuText(i)) <> "" Then
                      Cell.Range.Text = Dict_Original_RuText(i)
                      Cell.Range.Font.Name = "Times New Roman"
                  End If
                 GoTo 20
              End If
            Next i
         End If
20    Next Cell
     tbl.Range.Font.Name = "Times New Roman"
 
     Next tbl
   End With
   
    '============
   Dim RWd_Cell_Text As String
     Set myrange = WD.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Range
     Pcn = myrange.Cells.count
     Set Cell = myrange.Cells(2)
     Wd_Cell_Text = Mid(Cell.Range.Text, 1, Len(Cell.Range.Text) - 2)
     For i = 2 To Rn
        If Wd_Cell_Text = SDict_Original_EnText(i) Then
            ' " & UText(22914, 26524, 32763, 35793, 25991, 26412, 20026, 31354) & "，" & UText(21017, 20445, 30041, 21407, 25991) & "，" & UText(36991, 20813, 21333, 20803, 26684, 34987, 28165, 31354) & "
            If Trim(Dict_Original_RuText(i)) <> "" Then
                Cell.Range.Text = Dict_Original_RuText(i)
                Cell.Range.Font.Name = "Times New Roman"
            End If
            Exit For
        End If
     Next i
     Set Cell = myrange.Cells(Pcn)
     SWd_Cell_Text = Mid(Cell.Range.Text, 1, Len(Cell.Range.Text) - 2)
     SWd_Cell_Text = Replace(SWd_Cell_Text, "：", ":")
     Wd_Cell_Text = Left(SWd_Cell_Text, InStr(1, SWd_Cell_Text, ":") - 1)
     RWd_Cell_Text = Mid(SWd_Cell_Text, Len(Wd_Cell_Text), Len(SWd_Cell_Text) - Len(Wd_Cell_Text))
     For i = 2 To Rn
        If Wd_Cell_Text = SDict_Original_EnText(i) Then
            ' " & UText(22914, 26524, 32763, 35793, 25991, 26412, 20026, 31354) & "，" & UText(21017, 20445, 30041, 21407, 25991) & "，" & UText(36991, 20813, 21333, 20803, 26684, 34987, 28165, 31354) & "
            If Trim(Dict_Original_RuText(i)) <> "" Then
                Cell.Range.Text = Dict_Original_RuText(i) & RWd_Cell_Text
                Cell.Range.Font.Name = "Times New Roman"
            End If
            Exit For
        End If
     Next i
   Set myrange = WD.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Range
   Set Cell = myrange.Cells(1)
   Wd_Cell_Text = Mid(Cell.Range.Text, 1, Len(Cell.Range) - 2)
   Wd_Cell_Text = Left(Wd_Cell_Text, Len(Wd_Cell_Text) - 1)
    For i = 2 To Rn
        If Wd_Cell_Text = SDict_Original_EnText(i) Then
            ' " & UText(22914, 26524, 32763, 35793, 25991, 26412, 20026, 31354) & "，" & UText(21017, 20445, 30041, 21407, 25991) & "，" & UText(36991, 20813, 21333, 20803, 26684, 34987, 28165, 31354) & "
            If Trim(Dict_Original_RuText(i)) <> "" Then
                Cell.Range.Text = Dict_Original_RuText(i) & ":"
                Cell.Range.Font.Name = "Times New Roman"
            End If
            Exit For
        End If
     Next i
    Set Cell = myrange.Cells(3)
    Wd_Cell_Text = Mid(Cell.Range.Text, 1, Len(Cell.Range) - 2)
    Wd_Cell_Text = Left(Wd_Cell_Text, Len(Wd_Cell_Text) - 1)
     For i = 2 To Rn
        If Wd_Cell_Text = SDict_Original_EnText(i) Then
            ' " & UText(22914, 26524, 32763, 35793, 25991, 26412, 20026, 31354) & "，" & UText(21017, 20445, 30041, 21407, 25991) & "，" & UText(36991, 20813, 21333, 20803, 26684, 34987, 28165, 31354) & "
            If Trim(Dict_Original_RuText(i)) <> "" Then
                Cell.Range.Text = Dict_Original_RuText(i) & ":"
                Cell.Range.Font.Name = "Times New Roman"
            End If
            Exit For
        End If
     Next i

    Set myrange = WD.Sections(1).Footers(wdHeaderFooterFirstPage).Range
    For Each Paragraph In myrange.Paragraphs
      If Len(Paragraph) > 50 Then       '" & UText(29992, 38271, 24230, 21028, 26029) & "，" & UText(22914, 26524, 19981, 25913, 26684, 24335) & "，" & UText(21487, 29992, 25152, 22788, 27573, 33853, 21028, 26029) & "
       'Wd_Cell_Text = myrange.Paragraphs(7).Range.Text
       Wd_Cell_Text = Paragraph.Range.Text
       Wd_Cell_Text = Left(Wd_Cell_Text, Len(Wd_Cell_Text) - 1)
        For i = 2 To Rn
        If Wd_Cell_Text = SDict_Original_EnText(i) Then
            ' " & UText(22914, 26524, 32763, 35793, 25991, 26412, 20026, 31354) & "，" & UText(21017, 20445, 30041, 21407, 25991) & "，" & UText(36991, 20813, 27573, 33853, 34987, 28165, 31354) & "
            If Trim(Dict_Original_RuText(i)) <> "" Then
                Paragraph.Range.Text = Dict_Original_RuText(i)
                Paragraph.Range.Font.Name = "Times New Roman"
            End If
            Exit For
        End If
     Next i
    End If
    Next Paragraph
    
    GoTo 30
    
     Set myrange = WD.Sections(1).Headers(wdHeaderFooterPrimary).Range
    
     SWd_Cell_Text = Mid(myrange.Paragraphs(1).Range.Text, 1, Len(myrange.Paragraphs(1).Range.Text) - 1)
     SWd_Cell_Text = Replace(SWd_Cell_Text, "：", ":")
     Wd_Cell_Text = Left(SWd_Cell_Text, InStr(1, SWd_Cell_Text, ":") - 1)
     RWd_Cell_Text = Mid(SWd_Cell_Text, Len(Wd_Cell_Text), Len(SWd_Cell_Text) - Len(Wd_Cell_Text))
     For i = 2 To Rn
        If Wd_Cell_Text = SDict_Original_EnText(i) Then
            ' " & UText(22914, 26524, 32763, 35793, 25991, 26412, 20026, 31354) & "，" & UText(21017, 20445, 30041, 21407, 25991) & "，" & UText(36991, 20813, 21333, 20803, 26684, 34987, 28165, 31354) & "
            If Trim(Dict_Original_RuText(i)) <> "" Then
                Cell.Range.Text = Dict_Original_RuText(i) & RWd_Cell_Text & CHR(13)
                Cell.Range.Font.Name = "Times New Roman"
            End If
            Exit For
        End If
      Next i
     Wd_Cell_Text = Mid(myrange.Paragraphs(2).Range.Text, 1, Len(myrange.Paragraphs(2).Range.Text) - 1)
    
     For i = 2 To Rn
        If Wd_Cell_Text = SDict_Original_EnText(i) Then
            ' " & UText(22914, 26524, 32763, 35793, 25991, 26412, 20026, 31354) & "，" & UText(21017, 20445, 30041, 21407, 25991) & "，" & UText(36991, 20813, 21333, 20803, 26684, 34987, 28165, 31354) & "
            If Trim(Dict_Original_RuText(i)) <> "" Then
                Cell.Range.Text = Dict_Original_RuText(i) & RWd_Cell_Text
                Cell.Range.Font.Name = "Times New Roman"
            End If
            Exit For
        End If
      Next i
      
30     Dim hl As Hyperlink
     If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
     End If
     If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
     End If
     ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        'ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
     Dim HeaderRange As Range
     Set HeaderRange = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range
     With HeaderRange.Paragraphs(1).Range
         SWd_Cell_Text = Mid(.Text, 1, Len(.Text) - 1)
         SWd_Cell_Text = Replace(SWd_Cell_Text, "：", ":")
         Wd_Cell_Text = Left(SWd_Cell_Text, InStr(1, SWd_Cell_Text, ":") - 1)
         RWd_Cell_Text = Mid(SWd_Cell_Text, Len(Wd_Cell_Text) + 1, Len(SWd_Cell_Text) - Len(Wd_Cell_Text))
         For i = 2 To Rn
            If Wd_Cell_Text = SDict_Original_EnText(i) Then
                ' " & UText(22914, 26524, 32763, 35793, 25991, 26412, 20026, 31354) & "，" & UText(21017, 20445, 30041, 21407, 25991) & "，" & UText(36991, 20813, 27573, 33853, 34987, 28165, 31354) & "
                If Trim(Dict_Original_RuText(i)) <> "" Then
                    .Text = Dict_Original_RuText(i) & RWd_Cell_Text & CHR(13)
                    .Font.Name = "Times New Roman"
                End If
                Exit For
              End If
         Next i
     End With
    
    With HeaderRange.Paragraphs(2).Range
        For Each hl In .Hyperlinks
           hl.Delete
        Next hl
        Wd_Cell_Text = Mid(.Text, 1, Len(.Text) - 1)
        For i = 2 To Rn
           If Wd_Cell_Text = SDict_Original_EnText(i) Then
                ' " & UText(22914, 26524, 32763, 35793, 25991, 26412, 20026, 31354) & "，" & UText(21017, 20445, 30041, 21407, 25991) & "，" & UText(36991, 20813, 27573, 33853, 34987, 28165, 31354) & "
                If Trim(Dict_Original_RuText(i)) <> "" Then
                    .Text = Replace(.Text, .Text, Dict_Original_RuText(i))
                    .Font.Name = "Times New Roman"
                End If
                HyperlinkText = "Category"
                If Not WD.Tables(3).Cell(2, 1).Range.Bookmarks.Exists(HyperlinkText) Then
                    WD.Tables(3).Cell(2, 1).Range.Bookmarks.add Name:=HyperlinkText, Range:=WD.Tables(3).Cell(2, 1).Range
                End If
                .Select
                Selection.Hyperlinks.add Anchor:=Selection.Range, Address:="", SubAddress:=HyperlinkText
                With .Font
                 .Name = "Times New Roman"
                 .Underline = wdUnderlineSingle ' " & UText(21333, 19979, 21010, 32447) & "
                End With
                Exit For
            End If
        Next i
    End With
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
   '======================
   
      WD.SaveAs2 Path & "\" & BDocName & LangSuffix & "." & DocExt
      For Each tbl In WD.Tables
        For Each Cell In tbl.Range.Cells
        
          If Cell.Range.Fields.count = 0 And Cell.Range.ContentControls.count = 0 And _
          Cell.Range.InlineShapes.count = 0 Then
            Kw = Cell.Range.Text
            Kw = Mid(Kw, 1, Len(Kw) - 2)

            Kw = Replace(Kw, "（（", "((")
            Kw = Replace(Kw, "（(", "((")
            Kw = Replace(Kw, "(（", "((")
            Kw = Replace(Kw, "））", "))")
            Kw = Replace(Kw, "）)", "))")
            Kw = Replace(Kw, ")）", "))")
           If InStr(1, Kw, "((") > 0 Then
              Kw = Replace(Kw, "((", "")
           End If
           If InStr(1, Kw, "))") > 0 Then
              Kw = Replace(Kw, "))", "")
           End If
            Cell.Range.Text = Kw
          End If
        Next Cell
      Next tbl
     
      Dim doc As Document
      Set doc = Documents.Open(Path & "\" & DocName)
      doc.Activate
      For Each tbl In doc.Tables
        For Each Cell In tbl.Range.Cells
            If Cell.Range.Fields.count = 0 And Cell.Range.ContentControls.count = 0 And _
            Cell.Range.InlineShapes.count = 0 Then
            Kw = Cell.Range.Text
            Kw = Mid(Kw, 1, Len(Kw) - 2)

            Kw = Replace(Kw, "（（", "((")
            Kw = Replace(Kw, "（(", "((")
            Kw = Replace(Kw, "(（", "((")
            Kw = Replace(Kw, "））", "))")
            Kw = Replace(Kw, "）)", "))")
            Kw = Replace(Kw, ")）", "))")
           If InStr(1, Kw, "((") > 0 Then
              Kw = Replace(Kw, "((", "")
           End If
           If InStr(1, Kw, "))") > 0 Then
              Kw = Replace(Kw, "))", "")
           End If
           Cell.Range.Text = Kw
          End If
        Next Cell
      Next tbl
     
      Call Export1545ResultToHUB
      doc.Save
  Else
       For Each tbl In WD.Tables
        For Each Cell In tbl.Range.Cells
          If Cell.Range.Fields.count = 0 And Cell.Range.ContentControls.count = 0 And _
          Cell.Range.InlineShapes.count = 0 Then
            Kw = Cell.Range.Text
            Kw = Mid(Kw, 1, Len(Kw) - 2)
            Kw = Replace(Kw, "（（", "((")
            Kw = Replace(Kw, "（(", "((")
            Kw = Replace(Kw, "(（", "((")
            Kw = Replace(Kw, "））", "))")
            Kw = Replace(Kw, "）)", "))")
            Kw = Replace(Kw, ")）", "))")
           If InStr(1, Kw, "((") > 0 Then
              Kw = Replace(Kw, "((", "")
           End If
           If InStr(1, Kw, "))") > 0 Then
              Kw = Replace(Kw, "))", "")
           End If
            Cell.Range.Text = Kw
          End If
        Next Cell
     
      Next tbl
      Call Export1545ResultToHUB
  End If
End Sub


Function PositionsL(Str As String) As Integer()
    Dim pos As Integer
    Dim CountL As Integer
    Dim Positions() As Integer
    CountL = 0
    pos = 1
    Do While pos > 0
        pos = InStr(pos, Str, "(", vbBinaryCompare)
        If pos > 0 Then
            If Mid(Str, pos + 1, 1) = "(" Then
                CountL = CountL + 1
                ReDim Preserve Positions(1 To CountL)
                Positions(CountL) = pos
                pos = pos + 1
            Else
                pos = pos + 1
            End If
        End If
    Loop
    PositionsL = Positions
End Function
Function PositionsR(Str As String) As Integer()
    Dim pos As Integer
    Dim CountR As Integer
    Dim Positions() As Integer
    CountR = 0
    pos = 1
    Do While pos > 0
        pos = InStr(pos, Str, ")", vbBinaryCompare)
        If pos > 0 Then
            If Mid(Str, pos + 1, 1) = ")" Then
                CountR = CountR + 1
                ReDim Preserve Positions(1 To CountR)
                Positions(CountR) = pos
                pos = pos + 1
            Else
                pos = pos + 1
            End If
        End If
    Loop
    PositionsR = Positions
End Function

Sub AdjustMargin()
    
    b = InputBox(UText(35831, 36755, 20837, 25968, 23383, 29992, 20110, 35843, 25972, 24038, 36793, 36317) & "。" & UText(25968, 23383, 36234, 23567, 36234, 38752, 24038, 21453, 20043, 25968, 23383, 36234, 22823, 36234, 38752, 21491) & "，" & UText(30446, 21069, 27169, 26495, 35774, 23450, 25968, 23383) & "48" & CHR(10) & _
    "Please enter a number to adjust the left margin. The smaller the number, the closer it is to the left; conversely, the larger the number, the closer it is to the right. The current template setting is number 48.", UText(35843, 25972, 34920, 26684, 20301, 32622) & CHR(10) & "Enter a number to adjust the left margin", 48)
    Set WD = ThisDocument

  Dim sec As Section
    With WD
        For Each sec In .Sections
            With sec.PageSetup
                .LeftMargin = b
            End With
        Next sec
    End With
  End Sub


Sub VerifyReportID()
   
    '20241108
    
     Dim CLn As Integer
     Dim TbnE, i, j, k, Rn, Cn, Celn, ST, ED As Integer
     Dim Kw As String
     Dim rg As Range
  
     Dim ReportID, ClientID, SupplierID, FactoryID As String
     Dim Client, Supplier, Factory As String
     Dim Unit As String
     
    '-----------------" & UText(39564, 35777, 25253, 21578) & "ID，" & UText(25253, 21578, 21495) & "
    Dim jsonResponse As Object
    Dim jsonPara As Object
    Dim Url As String
  
    Dim ReportNo As String
    Dim Pcn As Integer
   
   

   
    Dim ClientName As String
    Dim SupplierName As String
    Dim FactoryName As String

    
3    Client = "Client"
     Supplier = "Supplier"
     Factory = "Factory"
     PO = "P.O. No."
     Item = "Item No."
     Quantity = "Quantity"
     Product_Description = "Product Description"
     Inspection_Type = "Inspection Type"
     Sequence = "Sequence"
     Inspection_Date = "Inspection Date"
     Location = "Location"
     Inspection_Basis = "Inspection Basis"
     Sample_Size = "Sample Size"
     
     '----------
     Set myrange = ThisDocument.Sections(1).Headers(wdHeaderFooterFirstPage).Range.Tables(1).Range
     Pcn = myrange.Cells.count
     rpn = myrange.Cells(Pcn).Range.Text
     rpn = Mid(rpn, 1, Len(rpn) - 2)
     rpn = Trim(Replace(rpn, "Report No: ", ""))
     

     '===========================
  
    Dim postJson As Object
     
         Set postJson = New Scripting.Dictionary
         postJson("bgh") = rpn
         
         Url = WebBaseUrl & "/api/sqlhub/exec?appid=md_330&configName=GetReportBasicInfo"
         Call getHUBdata(Url, postJson, jsonResponse)
  
         
           ' " & UText(26816, 26597, 36820, 22238, 25968, 25454, 26159, 21542, 20026, 31354) & "，" & UText(24182, 26681, 25454) & "dataIsEmpty" & UText(30340, 20540, 26469, 22788, 29702, 32467, 26524) & "
          Dim dataIsEmpty As Boolean
         If Not IsEmpty(jsonResponse("data")) Then
             dataIsEmpty = (TypeName(jsonResponse("data")) = "Collection") And (jsonResponse("data").count = 0)
          Else
             dataIsEmpty = True
          End If
         If dataIsEmpty Then
              MsgBox "Ｎo data is obtained from the HUB, the program will exit, please check 1. whether the HUB can be accessed via the Internet," & _
              "2. Whether the order is in the correct process stage on the HUB, such as the inspection notice has been issued and the report number has been generated. " & _
              "Whether the report number you input is correct!" & CHR(10) & CHR(10) & _
              UText(20174) & "HUB" & UText(26410, 33719, 21462, 21040, 25968, 25454) & "，" & UText(31243, 24207, 23558, 36864, 20986) & "，" & UText(35831, 26816, 26597) & "1." & UText(36890, 36807, 20114, 32852, 32593, 26159, 21542, 21487, 20197, 35775, 38382) & "HUB,2." & UText(36825, 20010, 35746, 21333, 22312) & "HUB" & UText(19978, 26159, 21542, 22788, 20110, 27491, 30830, 30340, 27969, 31243, 38454, 27573) & "，" & UText(27604, 22914, 39564, 36135, 36890, 30693, 21333, 24050, 21457, 20986) & "，" & UText(25253, 21578, 21495, 24050, 29983, 25104) & "," & _
              "3." & UText(20320, 36755, 20837, 30340, 25253, 21578, 21495, 26159, 21542, 27491, 30830) & CHR(10) & CHR(10), vbCritical
              Exit Sub
          End If
    
       
        Set WD = ThisDocument
           
           ReportID = jsonResponse("data")(1)("bghid")
           If ReportID <> Replace(GetPropertyValue("ReportID"), "ReportID" & "_", "") Then
              MsgBox UText(20320, 19978, 20256, 30340, 25253, 21578) & "ID" & UText(21644) & "HUB" & UText(19978, 30340, 25253, 21578) & "ID" & UText(19981, 21305, 37197) & "，" & UText(23558, 32456, 27490, 19978, 20256, 25253, 21578, 25968, 25454) & "。" & UText(35831, 30830, 35748, 25253, 21578, 30340, 22522, 26412, 20449, 24687, 26159, 21542, 26159, 36890, 36807, 36816, 34892) & "FILLINFORMATION" & UText(31243, 24207, 20174) & "HUB" & UText(23548, 20837, 32780, 19981, 26159, 20174, 20854, 20182, 22788, 25335, 36125, 20837, 30340) & "," & _
              UText(35831, 36890, 36807, 23548, 20837, 22522, 26412, 20449, 24687, 26356, 27491, 25253, 21578) & "ID" & UText(21518, 20877, 36816, 34892, 26412, 31243, 24207) & "！" & CHR(10) & CHR(10) & _
              "The report ID you uploaded does not match the report ID on the HUB. The upload of the report data will be terminated. " & _
              "Please verify that the basic information of the report was imported from the HUB by running the INFORMATION program, rather than copied from elsewhere. " & _
              "Correct the report ID by importing the basic information and then run this program again." & CHR(10) & CHR(10), vbCritical
              EXT = "Quit"
              Exit Sub
              
           End If
           
           ClientID = jsonResponse("data")(1)(HubKeyClientId())
           If ClientID <> GetPropertyValue("ClientID") Then
              MsgBox UText(20320, 19978, 20256, 30340, 23458, 25143) & "id" & UText(21644) & "HUB" & UText(19978, 30340, 23458, 25143) & "id" & UText(19981, 21305, 37197) & "，" & UText(23558, 32456, 27490, 19978, 20256, 25253, 21578, 25968, 25454) & "。" & UText(35831, 30830, 35748, 25253, 21578, 30340, 22522, 26412, 20449, 24687, 26159, 21542, 26159, 36890, 36807, 36816, 34892) & "FILLINFORMATION" & UText(31243, 24207, 20174) & "HUB" & UText(23548, 20837, 32780, 19981, 26159, 20174, 20854, 20182, 22788, 25335, 36125, 20837, 30340) & "，" & _
              UText(35831, 36890, 36807, 23548, 20837, 22522, 26412, 20449, 24687, 26356, 27491, 23458, 25143) & "id" & UText(20877, 36816, 34892, 26412, 31243, 24207) & "！" & UText(22914, 26524) & "HUB" & UText(23548, 20837, 30340, 23458, 25143, 20449, 24687, 26159, 38169, 35823, 30340) & "，" & UText(35831, 32852, 31995) & "CS" & UText(26356, 27491) & "HUB" & UText(19978, 30340, 25490, 21333, 20449, 24687) & "，" & UText(37325, 26032, 23548, 20837, 22522, 26412, 20449, 24687, 21518, 20877, 36816, 34892, 26412, 31243, 24207) & "。" & CHR(10) & CHR(10) & _
              "The ClientID you uploaded does not match the ClientID on the HUB. The upload of the report data will be terminated. " & _
              "Please verify that the basic information of the report was imported from the HUB by running the INFORMATION program, rather than copied from elsewhere. " & _
              "Correct the ClientID by importing the basic information and then run this program again. if the Client information imported from HUB is not correct," & _
              "Please contact CS to revise it, re-import it again and run this program again!" & CHR(10) & CHR(10), vbCritical
          EXT = "Quit"
           Exit Sub
           
           End If
           
           ClientName = jsonResponse("data")(1)(HubKeyClientFullName())
           If ClientName <> GetPropertyValue("Client") Then
              MsgBox UText(20320, 19978, 20256, 30340, 23458, 25143, 20840, 31216, 21644) & "HUB" & UText(19978, 30340, 23458, 25143, 20840, 31216, 19981, 21305, 37197) & "，" & UText(23558, 32456, 27490, 19978, 20256, 25253, 21578, 25968, 25454) & "。" & UText(35831, 30830, 35748, 25253, 21578, 30340, 22522, 26412, 20449, 24687, 26159, 21542, 26159, 36890, 36807, 36816, 34892) & "FILLINFORMATION" & UText(31243, 24207, 20174) & "HUB" & UText(23548, 20837, 32780, 19981, 26159, 20174, 20854, 20182, 22788, 25335, 36125, 20837, 30340) & "，" & _
              UText(35831, 36890, 36807, 23548, 20837, 22522, 26412, 20449, 24687, 26356, 27491, 23458, 25143, 20840, 31216, 20877, 36816, 34892, 26412, 31243, 24207) & "！" & UText(22914, 26524) & "HUB" & UText(23548, 20837, 30340, 23458, 25143, 20449, 24687, 26159, 38169, 35823, 30340) & "，" & UText(35831, 32852, 31995) & "CS" & UText(26356, 27491) & "HUB" & UText(19978, 30340, 25490, 21333, 20449, 24687) & "，" & UText(37325, 26032, 23548, 20837, 22522, 26412, 20449, 24687, 21518, 20877, 36816, 34892, 26412, 31243, 24207) & "。" & CHR(10) & CHR(10) & _
              "The Client name you uploaded does not match the Client name the HUB. The upload of the report data will be terminated. " & _
              "Please verify that the basic information of the report was imported from the HUB by running the INFORMATION program, rather than copied from elsewhere. " & _
              "Correct the Client name by importing the basic information and then run this program again. if the Client information imported from HUB is not correct," & _
              "Please contact CS to revise it, re-import it again and run this program again!" & CHR(10) & CHR(10), vbCritical
          EXT = "Quit"
           Exit Sub
        
           End If
            
            SupplierID = jsonResponse("data")(1)(HubKeySupplierId())
           If SupplierID <> GetPropertyValue("SupplierID") Then
              MsgBox UText(20320, 19978, 20256, 30340, 20379, 24212, 21830) & "ID" & UText(21644) & "HUB" & UText(19978, 30340, 20379, 24212, 21830) & "ID" & UText(19981, 21305, 37197) & "，" & UText(23558, 32456, 27490, 19978, 20256, 25253, 21578, 25968, 25454) & "。" & UText(35831, 30830, 35748, 25253, 21578, 30340, 22522, 26412, 20449, 24687, 26159, 21542, 26159, 36890, 36807, 36816, 34892) & "FILLINFORMATION" & UText(31243, 24207, 20174) & "HUB" & UText(23548, 20837, 32780, 19981, 26159, 20174, 20854, 20182, 22788, 25335, 36125, 20837, 30340) & "，" & _
              UText(35831, 36890, 36807, 23548, 20837, 22522, 26412, 20449, 24687, 26356, 27491, 20379, 24212, 21830) & "ID" & UText(21518, 20877, 36816, 34892, 26412, 31243, 24207) & "！" & UText(22914, 26524) & "HUB" & UText(23548, 20837, 30340, 20379, 24212, 21830, 20449, 24687, 26159, 38169, 35823, 30340) & "，" & UText(35831, 32852, 31995) & "CS" & UText(26356, 27491) & "HUB" & UText(19978, 30340, 25490, 21333, 20449, 24687) & "，" & UText(37325, 26032, 23548, 20837, 22522, 26412, 20449, 24687, 21518, 20877, 36816, 34892, 26412, 31243, 24207) & "。" & CHR(10) & CHR(10) & _
              "The SupplierID you uploaded does not match the SupplierID on the HUB. The upload of the report data will be terminated. " & _
              "Please verify that the basic information of the report was imported from the HUB by running the INFORMATION program, rather than copied from elsewhere. " & _
              "Correct the SupplierID by importing the basic information and then run this program again. if the supplier information imported from HUB is not correct," & _
              "Please contact CS to revise it, re-import it again and run this program again!" & CHR(10) & CHR(10), vbCritical
          EXT = "Quit"
           Exit Sub
        
           End If
          
            SupplierName = jsonResponse("data")(1)(HubKeySupplierFullName())
            If SupplierName <> GetPropertyValue("Supplier") Then
              MsgBox UText(20320, 19978, 20256, 30340, 20379, 24212, 21830, 20840, 31216, 21644) & "HUB" & UText(19978, 30340, 20379, 24212, 21830, 20840, 31216, 19981, 21305, 37197) & "，" & UText(23558, 32456, 27490, 19978, 20256, 25253, 21578, 25968, 25454) & "。" & UText(35831, 30830, 35748, 25253, 21578, 30340, 22522, 26412, 20449, 24687, 26159, 21542, 26159, 36890, 36807, 36816, 34892) & "FILLINFORMATION" & UText(31243, 24207, 20174) & "HUB" & UText(23548, 20837, 32780, 19981, 26159, 20174, 20854, 20182, 22788, 25335, 36125, 20837, 30340) & "，" & _
              UText(35831, 36890, 36807, 23548, 20837, 22522, 26412, 20449, 24687, 26356, 27491, 20379, 24212, 21830, 20840, 31216, 20877, 36816, 34892, 26412, 31243, 24207) & "！" & UText(22914, 26524) & "HUB" & UText(23548, 20837, 30340, 20379, 24212, 21830, 20449, 24687, 26159, 38169, 35823, 30340) & "，" & UText(35831, 32852, 31995) & "CS" & UText(26356, 27491) & "HUB" & UText(19978, 30340, 25490, 21333, 20449, 24687) & "，" & UText(37325, 26032, 23548, 20837, 22522, 26412, 20449, 24687, 21518, 20877, 36816, 34892, 26412, 31243, 24207) & "。" & CHR(10) & CHR(10) & _
              "The Supplier name you uploaded does not match the Supplier name on the HUB. The upload of the report data will be terminated. " & _
              "Please verify that the basic information of the report was imported from the HUB by running the INFORMATION program, rather than copied from elsewhere. " & _
              "Correct the Supplier name by importing the basic information and then run this program again. if the supplier information imported from HUB is not correct," & _
              "Please contact CS to revise it, re-import it again and run this program again!" & CHR(10) & CHR(10), vbCritical
          EXT = "Quit"
           Exit Sub
          
           End If
        
           FactoryID = jsonResponse("data")(1)(HubKeyFactoryId())
           If FactoryID <> GetPropertyValue("FactoryID") Then
              MsgBox UText(20320, 19978, 20256, 30340, 24037, 21378) & "id" & UText(21644) & "HUB" & UText(19978, 30340, 24037, 21378) & "id" & UText(19981, 21305, 37197) & "，" & UText(23558, 32456, 27490, 19978, 20256, 25253, 21578, 25968, 25454) & "。" & UText(35831, 30830, 35748, 25253, 21578, 30340, 22522, 26412, 20449, 24687, 26159, 21542, 26159, 36890, 36807, 36816, 34892) & "FILLINFORMATION" & UText(31243, 24207, 20174) & "HUB" & UText(23548, 20837, 32780, 19981, 26159, 20174, 20854, 20182, 22788, 25335, 36125, 20837, 30340) & "，" & _
              UText(35831, 36890, 36807, 23548, 20837, 22522, 26412, 20449, 24687, 26356, 27491, 20379, 24212, 21830, 20840, 31216, 20877, 36816, 34892, 26412, 31243, 24207) & "！" & UText(22914, 26524) & "HUB" & UText(23548, 20837, 30340, 20379, 24212, 21830, 20449, 24687, 26159, 38169, 35823, 30340) & "，" & UText(35831, 32852, 31995) & "CS" & UText(26356, 27491) & "HUB" & UText(19978, 30340, 25490, 21333, 20449, 24687) & "，" & UText(37325, 26032, 23548, 20837, 22522, 26412, 20449, 24687, 21518, 20877, 36816, 34892, 26412, 31243, 24207) & "。" & CHR(10) & CHR(10) & _
              "The FactoryID you uploaded does not match the FactoryID on the HUB. The upload of the report data will be terminated. " & _
              "Please verify that the basic information of the report was imported from the HUB by running the INFORMATION program, rather than copied from elsewhere. " & _
              "Correct the FactoryID by importing the basic information and then run this program again. if the Factory information imported from HUB is not correct," & _
              "Please contact CS to revise it, re-import it again and run this program again!" & CHR(10) & CHR(10), vbCritical
          EXT = "Quit"
           Exit Sub
         
           End If
           'StoretoProperty "FactoryID", FactoryID
           
           
           FactoryName = jsonResponse("data")(1)(HubKeyFactoryFullName())
            If FactoryName <> GetPropertyValue("Factory") Then
              MsgBox UText(20320, 19978, 20256, 30340, 24037, 21378, 20840, 31216, 21644) & "HUB" & UText(19978, 30340, 24037, 21378, 20840, 31216, 19981, 21305, 37197) & "，" & UText(23558, 32456, 27490, 19978, 20256, 25253, 21578, 25968, 25454) & "。" & UText(35831, 30830, 35748, 25253, 21578, 30340, 22522, 26412, 20449, 24687, 26159, 21542, 26159, 36890, 36807, 36816, 34892) & "FILLINFORMATION" & UText(31243, 24207, 20174) & "HUB" & UText(23548, 20837, 32780, 19981, 26159, 20174, 20854, 20182, 22788, 25335, 36125, 20837, 30340) & "，" & _
              UText(35831, 36890, 36807, 23548, 20837, 22522, 26412, 20449, 24687, 26356, 27491, 24037, 21378, 20840, 31216, 20877, 36816, 34892, 26412, 31243, 24207) & "！" & UText(22914, 26524) & "HUB" & UText(23548, 20837, 30340, 24037, 21378, 20840, 31216, 20449, 24687, 26159, 38169, 35823, 30340) & "，" & UText(35831, 32852, 31995) & "CS" & UText(26356, 27491) & "HUB" & UText(19978, 30340, 25490, 21333, 20449, 24687) & "，" & UText(37325, 26032, 23548, 20837, 22522, 26412, 20449, 24687, 21518, 20877, 36816, 34892, 26412, 31243, 24207) & "。" & CHR(10) & CHR(10) & _
              "The Factory name you uploaded does not match the Factory name on the HUB. The upload of the report data will be terminated. " & _
              "Please verify that the basic information of the report was imported from the HUB by running the INFORMATION program, rather than copied from elsewhere. " & _
              "Correct the Factory name by importing the basic information and then run this program again. if the supplier information imported from HUB is not correct," & _
              "Please contact CS to revise it, re-import it again and run this program again!" & CHR(10) & CHR(10), vbCritical
          EXT = "Quit"
           Exit Sub
          End If
           'StoretoProperty Factory, FactoryName
 
    '-------------------------------------------------
End Sub


Sub tt()
   Debug.Print ThisDocument.Tables(1).Cell(1, 1).Shading.BackgroundPatternColorIndex
   Debug.Print ThisDocument.Tables(1).Cell(2, 1).Shading.BackgroundPatternColorIndex
   Debug.Print ThisDocument.Tables(1).Cell(2, 2).Range.ContentControls(1).Range.Font.Color
   Debug.Print ThisDocument.Tables(1).Cell(3, 2).Shading.BackgroundPatternColorIndex
 
End Sub




Option Explicit

'==================== " & UText(21024, 38500, 31354, 25991, 20214, 22841, 20027, 20837, 21475) & " ====================
Sub DeleteEmptyFolders()
    Dim targetPath As String
    Dim deletedCount As Long
    
    '" & UText(35753, 29992, 25143, 36873, 25321, 25991, 20214, 22841) & "
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = UText(36873, 25321, 35201, 28165, 29702, 30340, 25991, 20214, 22841)
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub          '" & UText(29992, 25143, 21462, 28040) & "
        targetPath = .SelectedItems(1)
    End With
    
    If Right(targetPath, 1) <> "\" Then targetPath = targetPath & "\"
    
    If MsgBox(UText(30830, 23450, 21024, 38500) & " """ & targetPath & """ " & UText(19979, 25152, 26377, 23618, 32423, 30340, 31354, 25991, 20214, 22841) & "？" & vbCrLf & _
              UText(27492, 25805, 20316, 19981, 21487, 25764, 38144) & "！" & vbCrLf & "Delete all empty fold？This action is irreversible.", vbYesNo + vbExclamation, "") = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    deletedCount = DeleteEmptyFoldersRecursive_NoLog(targetPath)
    Application.ScreenUpdating = True
    
    MsgBox UText(23436, 25104) & "！" & UText(20849, 21024, 38500) & " " & deletedCount & " " & UText(20010, 31354, 25991, 20214, 22841) & "。", vbInformation
End Sub

'==================== " & UText(36882, 24402, 21024, 38500) & " ====================
Private Function DeleteEmptyFoldersRecursive_NoLog(ByVal folderPath As String) As Long
    Dim fso  As Object
    Dim fld  As Object
    Dim subF As Object
    Dim cnt  As Long
    
    On Error GoTo ErrH
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Function
    
    Set fld = fso.GetFolder(folderPath)
    
    ' translated comment
    For Each subF In fld.SubFolders
        cnt = cnt + DeleteEmptyFoldersRecursive_NoLog(subF.Path)
    Next subF
    
    ' translated comment
    If fld.Files.count = 0 And fld.SubFolders.count = 0 Then
        On Error Resume Next
        fso.DeleteFolder folderPath, True           'True=
        If Err.Number = 0 Then cnt = cnt + 1
        On Error GoTo ErrH
    End If
    
Done:
    DeleteEmptyFoldersRecursive_NoLog = cnt
    Exit Function
ErrH:
    ' translated comment
    Resume Next
End Function


