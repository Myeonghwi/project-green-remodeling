Attribute VB_Name = "Module1"
Sub GetAPTInfo()

    Dim strURL As String
        strURL = Range("APIurl").Value
        
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
    
    Set KaptList = xmlDoc.getElementsByTagName("item")
        
        For Each NodeLevel1 In KaptList
            For Each NodeLevel2 In NodeLevel1.getElementsByTagName("kaptCode")
               Debug.Print NodeLevel2.Text
            Next
            
            For Each NodeLevel2 In NodeLevel1.getElementsByTagName("kaptName")
               Debug.Print NodeLevel2.Text
            Next
        Next
        

End Sub



