Attribute VB_Name = "ReplaceMdl"
'Very Slow!!!!!! No answer
'Sub ReplaceStringInFile(solutionSet() As Double, setNumber As Integer)
'
'Dim sBuf As String
'Dim sTemp As String
'Dim iFileNum As Integer
'Dim sFileName As String
'
'    ' Edit as needed
'    sFileName = ThisWorkbook.Path & PROGRESS_PATH
'
'    iFileNum = FreeFile
'    Open sFileName For Input As iFileNum
'
'    Do Until EOF(iFileNum)
'        Line Input #iFileNum, sBuf
'        sTemp = sTemp & sBuf & vbCrLf
'    Loop
'    Close iFileNum
'
'    'source, origin, target
'    sTemp = Replace(sTemp, "var1", solutionSet(1, setNumber))
'    sTemp = Replace(sTemp, "var2", solutionSet(2, setNumber))
'
'    iFileNum = FreeFile
'    Open sFileName For Output As iFileNum
'    Print #iFileNum, sTemp
'    Close iFileNum
'
'End Sub

Sub CopyStringInFile()

    Dim srcFileObject As Object
    Dim strSourcePath As String
    Dim strDestinationPath As String
    Set srcFileObject = VBA.CreateObject("Scripting.FileSystemObject")

        strSourcePath = ThisWorkbook.Path & TEMPLATE_PATH
        strDestinationPath = ThisWorkbook.Path & PROGRESS_PATH
        
        Call srcFileObject.CopyFile(strSourcePath, strDestinationPath)

End Sub

Sub DeleteStringInFile()

    Dim srcFileObject As Object
    Dim strSourcePath As String
    Set srcFileObject = VBA.CreateObject("Scripting.FileSystemObject")

        strSourcePath = ThisWorkbook.Path & PROGRESS_PATH
        
        On Error GoTo ErrorHandler
            Call srcFileObject.DeleteFile(strSourcePath)

ErrorHandler:
Exit Sub

End Sub

'Very Very Faster
Sub ReplaceStringInFile(solutionSet() As Double, setNumber As Integer)

    Dim res As String
    Dim strData() As String
    Dim iFileNum As Integer
    Dim sTemp As String
    
    Open ThisWorkbook.Path & PROGRESS_PATH For Binary As #1
    res = Space$(LOF(1))
    Get #1, , res
    Close #1
    strData = Split(res, vbCrLf)
    
    iFileNum = 1
    
    Open ThisWorkbook.Path & PROGRESS_PATH For Output As iFileNum
    
    For i = 0 To UBound(strData)
    
        strData(i) = Replace(strData(i), "var1", solutionSet(1, setNumber))
        strData(i) = Replace(strData(i), "var2", solutionSet(2, setNumber))
        strData(i) = Replace(strData(i), "var3", solutionSet(3, setNumber))
        Print #iFileNum, strData(i)
        
    Next
   
    Close iFileNum
    
End Sub

'TODO: Boolean-If문으로 프로시저 없애주기
Sub ReplaceBestInFile(solutionSet() As Double)

    Dim res As String
    Dim strData() As String
    Dim iFileNum As Integer
    Dim sTemp As String
    
    Open ThisWorkbook.Path & PROGRESS_PATH For Binary As #1
    res = Space$(LOF(1))
    Get #1, , res
    Close #1
    strData = Split(res, vbCrLf)
    
    iFileNum = 1
    
    Open ThisWorkbook.Path & PROGRESS_PATH For Output As iFileNum
    
    For i = 0 To UBound(strData)
    
        strData(i) = Replace(strData(i), "var1", solutionSet(1, 1))
        strData(i) = Replace(strData(i), "var2", solutionSet(2, 1))
        strData(i) = Replace(strData(i), "var3", solutionSet(3, 1))
        Print #iFileNum, strData(i)
    Next
   
    Close iFileNum
    
End Sub

Sub ReplaceInsulationFile(Thickness As Double, arrSpec() As Double)
    
    Dim res As String
    Dim strData() As String
    Dim iFileNum As Integer
    Dim sTemp As String
    
    Open ThisWorkbook.Path & PROGRESS_PATH For Binary As #1
    res = Space$(LOF(1))
    Get #1, , res
    Close #1
    strData = Split(res, vbCrLf)
    
    iFileNum = 1
    
    Open ThisWorkbook.Path & PROGRESS_PATH For Output As iFileNum
    
    Dim arrTest
    arrTest = Array(10, 20, 30)
    
    For i = 0 To UBound(strData)
    
'        strData(i) = Replace(strData(i), "var1", arrSpec(0)) 'mm to m
'        strData(i) = Replace(strData(i), "var2", arrSpec(1))
'        strData(i) = Replace(strData(i), "var3", arrSpec(2))
'        strData(i) = Replace(strData(i), "var4", arrSpec(3))

        Print #iFileNum, strData(i)
    Next
   
    Close iFileNum
    
End Sub

Sub ReplaceTestFile()
    
    Dim res As String
    Dim strData() As String
    Dim iFileNum As Integer
    Dim sTemp As String
    
    Open ThisWorkbook.Path & PROGRESS_PATH For Binary As #1
    res = Space$(LOF(1))
    Get #1, , res
    Close #1
    strData = Split(res, vbCrLf)
    
    iFileNum = 1
    
    Open ThisWorkbook.Path & PROGRESS_PATH For Output As iFileNum
    
    ScrapeList          'Replacement_Module Sheet에 있는 요소들을 모두 스크래핑
    
    For i = 0 To UBound(strData)
    
        For j = 0 To rowCount           '요소들의 열수를 기반으로 Replacement를 진행함
        
            strData(i) = Replace(strData(i), lst(j, VAR_NAME), lst(j, REPLA_VALUE))
        
        Next

        Print #iFileNum, strData(i)
    Next
   
    Close iFileNum
    
End Sub
