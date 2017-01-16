Attribute VB_Name = "ExeMdl"
Sub StartExcutionFile()

        Dim wsh As Object
        Dim waitOnReturn As Boolean: waitOnReturn = True
        Dim windowStyle As Integer: windowStyle = 0
        Set wsh = CreateObject("WScript.Shell")
        
            wsh.Run ThisWorkbook.Path & BATCH_FILE_PATH, windowsStyle, waitOnReturn
            Application.ScreenUpdating = True
            'wsh.Run ThisWorkbook.Path & BATCH_FILE_PATH
            'Application.Wait (Now + TimeValue("00:00:09"))
            'TODO: Progress Bar 추가해주기
End Sub

Sub StartSimulation(arrSpec() As Double, rngMin As Double, rngMax As Double, term As Double)

    Dim num As Integer
    
        num = 1
        
        '단열재 범위값만을 위해서 테스트했을 때
        'arrSpec -> Index(0) = 열전도율, Index(1) = 밀도, Index(2) = 열용량
'        Do Until rngMin > rngMax
'
'            DeleteStringInFile
'
'            CopyStringInFile
'
'            Call ReplaceInsulationFile(rngMin, arrSpec())
'
'            StartExcutionFile
'
'            rngMin = rngMin + term
'
'            ParseXML
'
'            OutputResult (num)
'
'            num = num + 1
'
'        Loop

        
        
End Sub

Sub ExecuteHtml()

    Dim wsh As Object
    Dim waitOnReturn As Boolean: waitOnReturn = False
    Dim windowStyle As Integer: windowStyle = 0
    Set wsh = CreateObject("WScript.Shell")
    
    wsh.Run ThisWorkbook.Path & API_XML_PATH & "\roadcode.html", windowsStyle, waitOnReturn
    
End Sub
