VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsulationForm 
   Caption         =   "단열재 선택"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7155
   OleObjectBlob   =   "InsulationForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "InsulationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelButton_Click()

    Unload Me
    
End Sub

Private Sub CmdDirInputButton_Click()

    Unload Me
    
    InsulationDirForm.Show

End Sub

Private Sub CmdOKButton_Click()

    Dim arrSpec() As Double
    Dim splitRng() As String
    
    Set rngType_ = Range(Range("InsulationType"), Range("InsulationType").End(xlDown))

'        With Sheet4
'            .Range("G10").Value = Me.cmb01.Value & ", " & Me.cmb02.Value
'        End With
    With Range("Repla_Insulation")
    
        ReDim arrSpec(1 To 3)
    
        If ChkRngBox1.Value = True Then     '범위입력
                
            For Each rngType In rngType_        '외벽
                
                If rngType.Value = cmb01.Value Then
                
                    .Offset(2, IS_RANGE) = "TRUE"
                    .Offset(2, REPLA_VALUE).Value = ""
                   
                    For i = 1 To 3
                    
                        arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                        .Offset(2, REPLA_VALUE + i).Value = InsulationForm.Controls("TextBox" & i).Value      '두께 설정
                        .Offset(i + 2, REPLA_VALUE).Value = arrSpec(i)        '물성치설정
                    Next

                End If
            
            Next rngType
                
        Else                               '범위입력X
                 
            For Each rngType In rngType_        '외벽
                
                If rngType.Value = cmb01.Value Then
                        .Offset(2, IS_RANGE) = "FALSE"
                        splitRng = Split(cmb02.Value, " ")
                        .Offset(2, REPLA_VALUE).Value = CDbl(splitRng(0)) / 1000
                        
                    For i = 1 To 3
                        arrSpec(i) = rngType.Cells(1, i + 1).Value
                        .Offset(2, REPLA_VALUE + i).Value = ""
                        .Offset(i + 2, REPLA_VALUE).Value = arrSpec(i)
                    Next

                End If
            
            Next rngType
        
        End If
        
        
        If ChkRngBox2.Value = True Then     '범위입력

            For Each rngType In rngType_        '측벽
                
                If rngType.Value = cmb03.Value Then
                
                    .Offset(6, IS_RANGE) = "TRUE"
                    .Offset(6, REPLA_VALUE).Value = ""
                   
                    For i = 1 To 3
                    
                        arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                        .Offset(6, REPLA_VALUE + i).Value = InsulationForm.Controls("TextBox" & i + 3).Value      '두께 설정
                        .Offset(i + 6, REPLA_VALUE).Value = arrSpec(i)        '물성치설정
                    Next

                End If
            
            Next rngType
                
        Else                               '범위입력X
                 
            For Each rngType In rngType_        '측벽
                
                If rngType.Value = cmb03.Value Then
                        
                    .Offset(6, IS_RANGE) = "FALSE"
                    splitRng = Split(cmb04.Value, " ")
                    .Offset(6, REPLA_VALUE).Value = CDbl(splitRng(0)) / 1000
                    
                    For i = 1 To 3
                        arrSpec(i) = rngType.Cells(1, i + 1).Value
                        .Offset(6, REPLA_VALUE + i).Value = ""
                        .Offset(i + 6, REPLA_VALUE).Value = arrSpec(i)
                        
                    Next

                End If
            
            Next rngType
        
        End If
        
    End With
    
    InputToCell
        
    'Call StartSimulation(arrSpec(), CDbl(TextBox1.Value), CDbl(TextBox2.Value), CDbl(TextBox3.Value))
        
End Sub

Private Sub ChkRngBox1_Click() '외벽

    With InsulationForm
    
        If ChkRngBox1.Value = True Then
        
            cmb02.Enabled = False
            cmb02.Text = ""
            cmb02.BackColor = &H80000016
            
            For i = 1 To 3
            
                .Controls("TextBox" & i).Enabled = True
                .Controls("TextBox" & i).BackColor = &H80000005
                
            Next i
    '
        Else
            cmb02.Enabled = True
            cmb02.BackColor = &H80000005

            For i = 1 To 3
            
                .Controls("TextBox" & i).Enabled = Flase
                .Controls("TextBox" & i).BackColor = &H80000016
                
            Next i
            
        End If
        
    End With
End Sub

Private Sub ChkRngBox2_Click() '측벽

    With InsulationForm
    
        If ChkRngBox2.Value = True Then
        
            cmb04.Enabled = False
            cmb04.Text = ""
            cmb04.BackColor = &H80000016
            
            For i = 4 To 6
            
                .Controls("TextBox" & i).Enabled = True
                .Controls("TextBox" & i).BackColor = &H80000005
                
            Next i
    '
        Else
            cmb04.Enabled = True
            cmb04.BackColor = &H80000005

            For i = 4 To 6
            
                .Controls("TextBox" & i).Enabled = Flase
                .Controls("TextBox" & i).BackColor = &H80000016
                
            Next i
            
        End If
        
    End With
End Sub

Private Sub UserForm_Initialize()

    Dim rngType As Range
    Dim strType As String
    
    Set rngType_ = Range(Range("InsulationType"), Range("InsulationType").End(xlDown))
        
        cmb01.Clear
        cmb03.Clear
        
        For Each rngType In rngType_
            
            If rngType.Value <> "종류" Then
           
                cmb01.AddItem rngType.Value
            
            End If
        
        Next rngType
        
        For Each rngType In rngType_
            
            If rngType.Value <> "종류" Then
           
                cmb03.AddItem rngType.Value
            
            End If
        
        Next rngType
        
        cmb01.Value = cmb01.list(0)
        cmb02.Value = cmb02.list(20)
        
        cmb03.Value = cmb03.list(4)
        cmb04.Value = cmb04.list(30)
End Sub

Private Sub cmb01_Change()

    Set rngType_ = Range(Range("InsulationType"), Range("InsulationType").End(xlDown))
    Set rngTn_ = Range(Range("InsulationTn"), Range("InsulationTn").End(xlDown))
        
        With cmb02
            .Enabled = True
            .BackColor = &H80000005

            For Each rngTn In rngTn_
                
                If rngTn.Value <> "두께" Then

                    cmb02.AddItem rngTn.Value & " mm"

                End If
            
            Next rngTn
            
            For Each rngType In rngType_
            
                If rngType = cmb01.Value Then
                    
                    Dim temp
                    
                    temp = Split(rngType, " ")
                    
                    Image1.Picture = LoadPicture(ThisWorkbook.Path & "\files\image\insulation\" & temp(0) & ".jpg")
                
                End If
            
            Next rngType
            
            .SetFocus
            
        End With
    
End Sub

Private Sub cmb03_Change()

    Set rngType_ = Range(Range("InsulationType"), Range("InsulationType").End(xlDown))
    Set rngTn_ = Range(Range("InsulationTn"), Range("InsulationTn").End(xlDown))
        
        With cmb04
            .Enabled = True
            .BackColor = &H80000005

            For Each rngTn In rngTn_
                
                If rngTn.Value <> "두께" Then

                    cmb04.AddItem rngTn.Value & " mm"

                End If
            
            Next rngTn
            
            For Each rngType In rngType_
            
                If rngType = cmb03.Value Then
                    
                    Dim temp
                    
                    temp = Split(rngType, " ")
                    
                    Image2.Picture = LoadPicture(ThisWorkbook.Path & "\files\image\insulation\" & temp(0) & ".jpg")
                
                End If
            
            Next rngType
            
            .SetFocus
            
        End With
    
End Sub

Private Sub InputToCell()

    With InsulationForm
    
        If ChkRngBox1.Value = False Then      '범위입력이 아닐때
            
            For i = 1 To 2
            
                Range("Cell_Main_Insulation").Cells(i, 1) = .Controls("cmb0" & i).Value
                Range("Cell_Main_Insulation").Cells(i + 2, 1) = "범위 선택 안됨"
            Next
            
            For i = 1 To 2
            
                Range("Cell_Main_Insulation").Cells(i + 6, 1) = .Controls("cmb0" & i + 2).Value
                Range("Cell_Main_Insulation").Cells(i + 8, 1) = "범위 선택 안됨"
            Next
            
        Else        '범위 입력일 때
        
            '추후작성
            
        End If

    End With
End Sub
