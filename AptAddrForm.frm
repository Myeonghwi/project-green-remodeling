VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AptAddrForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "AptAddrForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "AptAddrForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kaptCodeList(1000)  As String
Dim kaptNameList(1000)  As String

Private Sub CmdCancelButton_Click()
    Unload Me
End Sub

Private Sub CmdOKButton_Click()

    With Me.ListBox1
    
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                GetAPTInfo (kaptCodeList(i))
            End If
        Next i
    
    End With
    
    Range("Result_AptSearch").Offset(i, 0).Interior.Color = RGB(255, 255, 255)
    Range("Result_AptSearch").Offset(i, 0).Locked = Ture
    
    For i = 0 To 16
    
        Range("Result_AptSearch").Offset(i, 0).Interior.Color = RGB(255, 255, 255)
        Range("Result_AptSearch").Offset(i, 0).Locked = Ture
    
    Next
            
    
    Unload Me
    
End Sub

Private Sub CmdSearchButton_Click()
    
    GetRoadInfo (Me.TextBox1.Value)


    Dim xmlDoc As New MSXML2.DOMDocument60
        xmlDoc.async = False
        xmlDoc.Load (ThisWorkbook.Path & API_XML_PATH & "\kaptcode.xml")
    
    Set KaptList = xmlDoc.getElementsByTagName("item")
        
    Dim i As Integer
    
        'list배열에 아파트 정보 입력
        i = 0
        For Each NodeLevel1 In KaptList
        
            For Each NodeLevel2 In NodeLevel1.getElementsByTagName("kaptCode")
    
               kaptCodeList(i) = NodeLevel2.Text
               
            Next
            
            For Each NodeLevel2 In NodeLevel1.getElementsByTagName("kaptName")
            
               kaptNameList(i) = NodeLevel2.Text
               
            Next
            
            i = i + 1
            
        Next
        
        
        'list box에 입력
        i = 0
        For Each NodeLevel1 In KaptList
        
            With Me.ListBox1
                .AddItem
                .Column(0, Me.ListBox1.ListCount - 1) = i + 1 & ". " & kaptNameList(i) & " (코드 : " & kaptCodeList(i) & ")"
            End With
            
            i = i + 1
            
        Next
End Sub

Sub GetAPTInfo(kaptcode As String)

    Dim strURL As String
        strURL = "http://apis.data.go.kr/1611000/AptBasisInfoService/getAphusBassInfo?serviceKey=qmtrltW6G7zoOxVeLWJJ%2FE%2BYEnmZeicm4b8mQQmJPnS1ZKDpg1dg1xLMIiKfzleMVxHD%2F9%2FvECvVBINhw2QcEw%3D%3D&kaptCode=" & kaptcode
        
    Dim hReq As New WinHttpRequest
        hReq.Open "GET", strURL, False
        hReq.send

    Dim FileStream, strData
    Set FileStream = CreateObject("ADODB.Stream")
    
        FileStream.Open
        FileStream.Position = 0
        FileStream.Charset = "iso-8859-1"
        FileStream.WriteText hReq.responseText
        FileStream.SaveToFile ThisWorkbook.Path & API_XML_PATH & "\kaptinfo.xml", 2
        strData = FileStream.ReadText
        FileStream.Close
    
    
    Dim xmlDoc As New MSXML2.DOMDocument60
        xmlDoc.async = False
        xmlDoc.Load (ThisWorkbook.Path & API_XML_PATH & "\kaptinfo.xml")
    
    Set KaptInfoList = xmlDoc.getElementsByTagName("item")
    
        For Each NodeLevel1 In KaptInfoList
               Range("Result_AptSearch").Offset(0, 0).Value = NodeLevel1.getElementsByTagName("kaptName")(0).Text
               Range("Result_AptSearch").Offset(1, 0).Value = NodeLevel1.getElementsByTagName("kaptAddr")(0).Text
               Range("Result_AptSearch").Offset(2, 0).Value = NodeLevel1.getElementsByTagName("codeSaleNm")(0).Text
               Range("Result_AptSearch").Offset(3, 0).Value = NodeLevel1.getElementsByTagName("codeHeatNm")(0).Text
               Range("Result_AptSearch").Offset(4, 0).Value = NodeLevel1.getElementsByTagName("kaptdaCnt")(0).Text
               Range("Result_AptSearch").Offset(5, 0).Value = NodeLevel1.getElementsByTagName("kaptDongCnt")(0).Text
               Range("Result_AptSearch").Offset(6, 0).Value = NodeLevel1.getElementsByTagName("codeHallNm")(0).Text
               Range("Result_AptSearch").Offset(7, 0).Value = NodeLevel1.getElementsByTagName("kaptMarea")(0).Text
               Range("Result_AptSearch").Offset(8, 0).Value = NodeLevel1.getElementsByTagName("kaptTarea")(0).Text
               Range("Result_AptSearch").Offset(9, 0).Value = NodeLevel1.getElementsByTagName("kaptBcompany")(0).Text
               Range("Result_AptSearch").Offset(10, 0).Value = NodeLevel1.getElementsByTagName("kaptAcompany")(0).Text
               Range("Result_AptSearch").Offset(11, 0).Value = NodeLevel1.getElementsByTagName("kaptUsedate")(0).Text
               Range("Result_AptSearch").Offset(12, 0).Value = NodeLevel1.getElementsByTagName("kaptMparea_60")(0).Text
               Range("Result_AptSearch").Offset(13, 0).Value = NodeLevel1.getElementsByTagName("kaptMparea_85")(0).Text
               Range("Result_AptSearch").Offset(14, 0).Value = NodeLevel1.getElementsByTagName("kaptMparea_135")(0).Text
               Range("Result_AptSearch").Offset(15, 0).Value = NodeLevel1.getElementsByTagName("kaptdaSize_136")(0).Text
               Range("Result_AptSearch").Offset(16, 0).Value = NodeLevel1.getElementsByTagName("kaptTel")(0).Text
        Next
        
End Sub

Sub GetRoadInfo(roadcode As String)

    Dim strURL As String
        strURL = "http://apis.data.go.kr/1611000/AptListService/getRoadnameAptList?serviceKey=qmtrltW6G7zoOxVeLWJJ%2FE%2BYEnmZeicm4b8mQQmJPnS1ZKDpg1dg1xLMIiKfzleMVxHD%2F9%2FvECvVBINhw2QcEw%3D%3D&loadCode=" & roadcode & "&pageNo=1&startPage=1&numOfRows=10&pageSize=10"
        
    Dim hReq As New WinHttpRequest
        hReq.Open "GET", strURL, False
        hReq.send

    Dim FileStream, strData
    Set FileStream = CreateObject("ADODB.Stream")
    
        FileStream.Open
        FileStream.Position = 0
        FileStream.Charset = "iso-8859-1"
        FileStream.WriteText hReq.responseText
        FileStream.SaveToFile ThisWorkbook.Path & API_XML_PATH & "\kaptcode.xml", 2
        strData = FileStream.ReadText
        FileStream.Close
    
    
    Dim xmlDoc As New MSXML2.DOMDocument60
        xmlDoc.async = False
        xmlDoc.Load (ThisWorkbook.Path & API_XML_PATH & "\kaptcode.xml")
        

End Sub
