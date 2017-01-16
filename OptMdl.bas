Attribute VB_Name = "OptMdl"
Option Explicit

Dim acceptableValueTypes As Integer '-1==>reals, 0==>not set yet, 1==> integers
Dim lowValue As Double
Dim highValue As Double
Dim bestGeneration As Integer
Dim startFromScratch As Integer '-1==>no (start from last run), 0==>not set yet, 1==>yes

Dim numberOfVariable As Long
Dim varSetRange As String
Dim conditionSetRange As String

Dim objectiveFunction As String
Dim bestSetNumber As Integer
Dim bestScore As Double
Dim bestSetValue() As Double

Dim setPerGeneration As Integer
Dim elite As Integer
Dim parentPool As Integer
Dim mutation As Integer
Dim generationRequested As Integer
Dim maxScoreWanted As Integer

Dim score() As Double
Dim ranking() As Integer
Dim tempArray() As Double
Dim solutionSet() As Double

Dim genNumber As Integer
Dim calcSheet As String
Dim calcMode As Long
Dim bolExtendedForm As Boolean

Dim populationArray() As Double

Sub ShowParameterForm()

    With ParameterForm
    
        .Show
        
    End With
    
End Sub

Sub StartCalibrationAuto() '자동보정 시작
    
    
    Dim temp() As String
    
    Dim addrRngVarSet
    Dim addrRngSetCondition

    Set addrRngVarSet = Range("RngSetVariable")
    Set addrRngSetCondition = Range("RngSetCondition")
    

        temp = Split(addrRngVarSet.Name, "=")
        varSetRange = temp(1)
        numberOfVariable = ComputeSetLength(varSetRange)
        
        temp = Split(addrRngSetCondition.Name, "=")
        conditionSetRange = temp(1)

        objectiveFunction = "Calibration_Engine!$H18"

        setPerGeneration = 10
        
        elite = 20
        elite = isElite
        
        parentPool = 50
        
        mutation = 8
        
        generationRequested = 1
        
        StartFromRandomGeneration

        maxScoreWanted = -1
        
        bestScore = 1.79E+308
        
        StartCalibration

End Sub

Sub StartCalibration() '직접보정 설정 후 시작

    Dim runComplete As Boolean
    
        With ParameterForm
            On Error GoTo Cleanup
            
            calcMode = Application.Calculation
            'Application.Calculation = xlCalculationManual
            'Application.Cursor = xlWait
            
            ReDim solutionSet(numberOfVariable, setPerGeneration)
            ReDim tempArray(numberOfVariable, setPerGeneration)
            ReDim bestSetValue(numberOfVariable)
            ReDim score(setPerGeneration)
            ReDim ranking(setPerGeneration)
    
            Application.ScreenUpdating = False
            
            If .firstGenerationPreviousOption.Value = True Then
            
                StartFromPreviousGeneration
                
            Else
            
                PopulateInitialGeneration
                
            End If
            
            runComplete = False
            genNumber = 1
            
            While Not runComplete
            
                CalculateBestSet
                
                RankTheScore
                
                AddChildren
                
                AddMutation
                
                runComplete = isCompleted
                
                Application.ScreenUpdating = True
                
            Wend
            
        End With

Cleanup:
    Application.Cursor = xlDefault
    Application.Calculation = calcMode
    
End Sub

'TODO : Should add exception process
Function GetParameter() As Boolean

    Dim readyToRun As Boolean
    
        With ParameterForm
        
            If .inputSetRangeTextBox.Value <> "" Then
                varSetRange = .inputSetRangeTextBox.Value
                numberOfVariable = ComputeSetLength(varSetRange)
            End If
            
            If .conditionSetTextBox.Value <> "" Then
                conditionSetRange = .conditionSetTextBox.Value
            End If
            
            If .objectiveFunctionTextBox.Value <> "" Then
                objectiveFunction = .objectiveFunctionTextBox.Value
            End If
            
            If .generationTextBox <> "" Then
                setPerGeneration = .generationTextBox
            End If
            
            If .eliteTextBox <> "" Then
                elite = CInt(.eliteTextBox)
                elite = isElite
            End If
            
            If .parentPerChildTextBox <> "" Then
                parentPool = CInt(.parentPerChildTextBox)
            End If
            
            If .mutationTextBox <> "" Then
                mutation = CInt(.mutationTextBox)
            End If
            
            If .genRequestedTextBox <> "" Then
                generationRequested = CInt(.genRequestedTextBox)
            End If
            
            If .firstGenerationRandomOption.Value = True Then
                StartFromRandomGeneration
            End If
    
            If .goalTypeMaxOption = True Then
                maxScoreWanted = 1
            Else
                maxScoreWanted = -1
                .goalTypeMinOption.Value = True
            End If
            
            If (maxScoreWanted = 1) Or (maxScoreWanted = 0) Then
                bestScore = 4.94065645841247E-324
            Else
                bestScore = 1.79E+308
            End If
            
    
    '
    '        If .elitesTextBox = "" Then
    '            elites = -1
    '        Else
    '            elites = CInt(.elitesTextBox)
    '        End If
    '        If elites > setsPerGeneration * 0.5 Then
    '            elites = Int(setsPerGeneration * 0.5)
    '        End If
    '        If elites < 1 Then
    '            errorString = errorString & vbCrLf & "Your number of elites must be an integer greater than 0" _
    '            & vbCrLf & "   and at most half the number of sets per generation."
    '        End If
    '
    '        If .parentsPerChildTextBox = "" Then
    '            parentPool = -1
    '        Else
    '            parentPool = CInt(.parentsPerChildTextBox)
    '        End If
    '        If (parentPool < 1) Or (parentPool > setsPerGeneration) Then
    '            errorString = errorString & vbCrLf & "Your number of parents per child must be an integer greater than 0" _
    '                & vbCrLf & "   and at most the number of sets per generation."
    '        End If
    '
    '        If .generationsRequestedTextBox = "" Then
    '            generationsRequested = -1
    '        Else
    '            generationsRequested = CInt(.generationsRequestedTextBox)
    '        End If
    '        If generationsRequested < 1 Then
    '            errorString = errorString & vbCrLf & "Your number of generations must be an integer greater than 0."
    '        End If
    '
    '        If .firstGenerationRandomOption.Value = True Then
    '            startFromScratch = 1
    '        Else
    '            startFromScratch = -1
    '        End If
    '
    '        If .inputSetRangeTextBox.Value <> "" Then
    '            testSetRange = .inputSetRangeTextBox.Value
    '        Else
    '            errorString = errorString & vbCrLf & "invalid range for test set (must be valid range in this workbook)"
    '        End If
    '        If .conditionsSetTextBox.Value <> "" Then
    '            conditionsSetRange = .conditionsSetTextBox.Value
    '        Else
    '            errorString = errorString & vbCrLf & "invalid range for conditions set (must be valid range in this workbook)"
    '        End If
    '        setLength = fn_ComputeSetLength(testSetRange)
    '        If setLength <= 0 Then
    '            errorString = errorString & vbCrLf & "invalid range for test set (must be valid range in this workbook)"
    '        End If
    '
    '        If .scoreRangeTextBox.Value <> "" Then
    '            scoreRange = .scoreRangeTextBox.Value
    '        Else
    '            errorString = errorString & vbCrLf & "invalid range for score cell (must be valid cell address in this workbook)"
    '        End If
    '
    '        If .mutationsTextBox = "" Then
    '            mutations = -1
    '        Else
    '            mutations = CInt(.mutationsTextBox)
    '        End If
    '        If (mutations < 0) Or (mutations > setLength) Then
    '            errorString = errorString & vbCrLf & "Your number of mutations must be an integer greater than or equal to 0" _
    '                & vbCrLf & " and at most the length of a set (in this case, " & setLength & ")."
    '        End If
    '
    '        If errorString = "" Then
    '            .Hide
    '            If startFromScratch = 1 Then
    '                'delete old version of DNA sheet if it exists, and recreate it
    '                Dim xlSheet As Excel.Worksheet, currentSheet As String
    '                For Each xlSheet In ActiveWorkbook.Sheets
    '                    If xlSheet.Name = "DNA" Then
    '                        Application.DisplayAlerts = False
    '                        xlSheet.Delete
    '                        Application.DisplayAlerts = True
    '                        Exit For
    '                    End If
    '                Next xlSheet
    '                currentSheet = ActiveSheet.Name
    '                Sheets.Add
    '                ActiveSheet.Name = "DNA"
    '                Sheets(currentSheet).Activate
    '            End If
    '            readyToRun = True
    '        Else
    '            MsgBox errorString & vbCrLf & vbCrLf & "Please revise and continue"
    '            readyToRun = False
    '        End If
    '
    '        If .goalTypeMaxOption = True Then
    '            maxScoreWanted = 1
    '        Else
    '            maxScoreWanted = -1
    '            .goalTypeMinOption.Value = True
    '        End If
    '
    '        If (maxScoreWanted = 1) Or (maxScoreWanted = 0) Then
    '            bestScore = 4.94065645841247E-324
    '        Else
    '            bestScore = 1.79E+308
    '        End If
    '    End With
            readyToRun = True
            GetParameter = readyToRun
        End With
End Function

Function isElite() As Double

    If elite > Int(setPerGeneration * 0.5) Then
        elite = Int(setPerGeneration * 0.5)
    End If

isElite = elite

End Function

Function ComputeSetLength(setRange As String) As Integer

    Dim numberOfVariable As Integer
    
    Dim sheetName As String
    Dim sht As Excel.Worksheet
    Dim sheetExists As Boolean
    
    Dim pos As Integer
    Dim rng As Excel.Range
    Dim localRange As String
    
        'check to see if sheet exists
        pos = InStr(setRange, "!")
        
        If pos > 1 Then
            sheetName = Left(setRange, pos - 1)
            localRange = Mid(setRange, pos + 1)
            For Each sht In ActiveWorkbook.Sheets
                If (sht.Name = sheetName) Or (sht.Name = Replace(sheetName, "'", "")) Then
                    sheetExists = True
                    'Debug.Print sht.Name, sheetName
                    Exit For
                End If
            Next sht
        End If
        
        If sheetExists Then
            Set rng = Sheets(Replace(sheetName, "'", "")).Range(localRange)
            numberOfVariable = rng.Rows.count
        End If
        
        ComputeSetLength = numberOfVariable
End Function

Sub PopulateInitialGeneration()
    
    Dim setNumber As Integer
    Dim snpNum As Long
    Dim geneVal As Double
        
        For setNumber = 1 To setPerGeneration
            'Populate each item of the strand (trial solution set) with a random acceptable value
            For snpNum = 1 To numberOfVariable
                geneVal = ValidVariable(snpNum)
                solutionSet(snpNum, setNumber) = geneVal
            Next snpNum
            solutionSet(0, setNumber) = setNumber
        Next setNumber
        CopySetToDNAsheet
    
End Sub

Function ValidVariable(snpNum As Long) As Double '실질적으로 Random 값을 생성함

    Const sheetName = "Calibration_Engine"
    Const maxLoopCounter = 1000
    
    Dim ValidValue As Double
    Dim bolValid As Boolean
    Dim loopCounter As Integer
    
    Dim rowWithValue As Long
    Dim rowWithTest As Long
    Dim colWithValue As Integer
    Dim colWithTest As Integer
    
    Dim minValue As Double
    Dim maxValue As Double
    Dim mustBeInt As Boolean
        
        'plug successive values into cell for the snpNum, and test the ValueOK cell for that snpNum
        Sheets(sheetName).Activate
        With Sheets(sheetName)
            colWithValue = .Range("InputSet").Column
            colWithTest = .Range("ConditionSet").Column
            rowWithValue = .Range("InputSet").Row + snpNum - 1
            rowWithTest = .Range("ConditionSet").Row + snpNum - 1
            minValue = .Cells(rowWithTest, colWithTest + 1)
            maxValue = .Cells(rowWithTest, colWithTest + 2)
            
            If UCase(.Cells(rowWithTest, colWithTest + 3)) = "Y" Then
                mustBeInt = True
            Else
                mustBeInt = False
            End If
            
            bolValid = False
            loopCounter = 0
            
            Do Until (bolValid = True) Or (loopCounter > maxLoopCounter)
                ValidValue = (maxValue - minValue) * Rnd() + minValue
                
                If mustBeInt Then
                    ValidValue = Int(ValidValue)
                End If
                
                Cells(rowWithValue, colWithValue) = ValidValue
                .Range("ConditionSet").Calculate
                DoEvents
                
                If Cells(rowWithTest, colWithTest) = True Then
                    bolValid = True
                End If
                
                loopCounter = loopCounter + 1
            Loop
        End With
        
        If loopCounter > maxLoopCounter Then
            Debug.Print Now & " unable to find acceptable value for snpNum " & snpNum & " in "; maxLoopCounter & " tries"
        End If
        
        ValidVariable = ValidValue
        
End Function

Sub CopySetToDNAsheet(Optional x As Integer = 0)

    Dim i As Integer
    Dim r As Long
    Dim solutionSetRange As String
        
        solutionSetRange = Cells(1, 1).Address & ":" & Cells(numberOfVariable + 1, setPerGeneration + 1).Address
        Sheets("INFO_DNA").Range(solutionSetRange) = solutionSet
        
        For i = 1 To setPerGeneration
            Sheets("INFO_DNA").Cells(numberOfVariable + 2, i + 1) = i
            Sheets("INFO_DNA").Cells(numberOfVariable + 3, i + 1) = score(i)
        Next i
        
        Sheets("INFO_DNA").Cells(1, 1) = "snp\Set"
        
        For i = 1 To numberOfVariable
            Sheets("INFO_DNA").Cells(i + 1, 1) = i
        Next i
        
        Sheets("INFO_DNA").Cells(numberOfVariable + 2, 1) = "Set"
        Sheets("INFO_DNA").Cells(numberOfVariable + 3, 1) = "Score"
        
End Sub

Sub StartFromRandomGeneration()

    'delete old version of DNA sheet if it exists, and recreate it
    Dim xlSheet As Excel.Worksheet
    Dim currentSheet As String
    
        For Each xlSheet In ActiveWorkbook.Sheets
            If xlSheet.Name = "INFO_DNA" Then
                Application.DisplayAlerts = False
                xlSheet.Delete
                Application.DisplayAlerts = True
                Exit For
            End If
        Next xlSheet
        
        currentSheet = ActiveSheet.Name
        Sheets.Add
        ActiveSheet.Name = "INFO_DNA"
        Sheets(currentSheet).Activate
        
End Sub

Sub StartFromPreviousGeneration(Optional x As Integer = 0)

    Dim setNumber As Integer, snpNum As Long
    
        For snpNum = 1 To numberOfVariable
            For setNumber = 1 To setPerGeneration
                solutionSet(snpNum, setNumber) = Sheets("INFO_DNA").Cells(snpNum + 1, setNumber + 1)
            Next setNumber
        Next snpNum
    
End Sub

Sub CalculateBestSet(Optional x As Integer = 0)
    
    Dim setNumber As Integer
    Dim snpNum As Long

        'Debug.Print "Entering TestsSet for Generation " & genNumber
        For setNumber = 1 To setPerGeneration
            For snpNum = 1 To numberOfVariable
                Range(varSetRange).Cells(snpNum, 1) = solutionSet(snpNum, setNumber)
            Next snpNum
            
            DeleteStringInFile
            CopyStringInFile
            Call ReplaceStringInFile(solutionSet, setNumber)
            StartExcutionFile
            ParseXML
            
            Calculate
            DoEvents
            score(setNumber) = Range(objectiveFunction).Cells(1, 1)
    
        If maxScoreWanted = 1 Then
            If score(setNumber) >= bestScore Then
                bestSetNumber = setNumber
                bestScore = score(setNumber)
                For snpNum = 1 To numberOfVariable
                    bestSetValue(snpNum) = solutionSet(snpNum, setNumber)
                Next snpNum
            End If
        Else
            If score(setNumber) <= bestScore Then
                bestSetNumber = setNumber
                bestScore = score(setNumber)
                For snpNum = 1 To numberOfVariable
                    bestSetValue(snpNum) = solutionSet(snpNum, setNumber)
                Next snpNum
            End If
        End If
        
        Next setNumber
    
        CopySetToDNAsheet
        
End Sub

Sub RankTheScore(Optional x As Integer = 0)
    
    'simple bubble sort for now
    Dim i As Integer
    Dim j As Integer
    Dim temp As Double
    Dim iTemp As Integer
    
        With ParameterForm
            For i = 1 To setPerGeneration
                ranking(i) = i
            Next i
            'Application.StatusBar = "generation: " & genNumber & " ... Best score = " _
            '& Format(bestScore, "###,###,###,##0.######") & " from set " & bestSetNumber
            For i = 1 To setPerGeneration - 1
                For j = i + 1 To setPerGeneration
                    If maxScoreWanted = 1 Then
                        If score(i) < score(j) Then
                            temp = score(j)
                            score(j) = score(i)
                            score(i) = temp
                            iTemp = ranking(j)
                            ranking(j) = ranking(i)
                            ranking(i) = iTemp
                        End If
                    Else
                        If score(i) > score(j) Then
                            temp = score(j)
                            score(j) = score(i)
                            score(i) = temp
                            iTemp = ranking(j)
                            ranking(j) = ranking(i)
                            ranking(i) = iTemp
                        End If
                    End If
                Next j
            Next i
            
            For i = 1 To setPerGeneration
                For j = 1 To numberOfVariable
                    tempArray(j, i) = solutionSet(j, ranking(i))
                Next j
                tempArray(0, i) = i
            Next i
            
            solutionSet = tempArray
            
            CopySetToDNAsheet
            
            DeleteStringInFile
            CopyStringInFile
            Call ReplaceBestInFile(solutionSet)
            StartExcutionFile
            ParseXML
            
            Calculate
            
        End With
End Sub

Sub AddChildren(Optional x As Integer = 0)

    Dim parent As Integer
    Dim snpNum As Long
    Dim child As Integer
    Dim children As Integer
        
        children = setPerGeneration - elite
            For child = 1 To children
                For snpNum = 1 To numberOfVariable
                    parent = Int(parentPool * Rnd()) + 1
                    solutionSet(snpNum, elite + child) = solutionSet(snpNum, parent)
                Next snpNum
            Next child
        
End Sub

Sub AddMutation()

    Dim snpNum As Long
    Dim mutationCount As Integer
    Dim setNum As Integer
    Dim geneVal As Double
    
        For setNum = elite + 1 To setPerGeneration
            For mutationCount = 1 To mutation
                snpNum = Int((numberOfVariable - 1) * Rnd()) + 1
                geneVal = ValidVariable(snpNum)
                solutionSet(snpNum, setNum) = geneVal
            Next mutationCount
        Next setNum

End Sub

Function isCompleted()
    
    Dim doneOrNotDone As Boolean
    Dim snpNum As Long
    
        CopySetToDNAsheet
        
        If genNumber >= generationRequested Then
            doneOrNotDone = True
            For snpNum = 1 To numberOfVariable
                Range(varSetRange).Cells(snpNum, 1) = solutionSet(snpNum, 1)
            Next snpNum
            Calculate
        Else
            genNumber = genNumber + 1
        End If
                
        isCompleted = doneOrNotDone
        
End Function
