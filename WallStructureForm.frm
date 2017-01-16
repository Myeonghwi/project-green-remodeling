VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WallStructureForm 
   Caption         =   "벽체 설정"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760.001
   OleObjectBlob   =   "WallStructureForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "WallStructureForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelButton_Click()

    Unload Me
    
End Sub

Private Sub CmdOKButton_Click()

    Dim temp() As String
    
    Set rngType_ = Range(Range("ConcreteType"), Range("ConcreteType").End(xlDown))
    
    With Range("Repla_Wall")
    
        ReDim arrSpec(1 To 3)
        
        For Each rngType In rngType_        '외벽 콘크리트
        
            If rngType = cmb01.Value Then
            
                For i = 1 To 3
                
                    arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                    .Offset(i + 2, REPLA_VALUE).Value = arrSpec(i)
                    
                Next
                
                temp() = Split(cmb02.Value, " ")
                .Offset(2, REPLA_VALUE).Value = CDbl(temp(0)) / 1000 'mm to m
                
            End If
        Next
        
        For Each rngType In rngType_        '측벽 콘크리트
        
            If rngType = cmb07.Value Then
            
                For i = 1 To 3
                
                    arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                    .Offset(i + 10, REPLA_VALUE).Value = arrSpec(i)
                    
                Next
                
                temp() = Split(cmb08.Value, " ")
                .Offset(10, REPLA_VALUE).Value = CDbl(temp(0)) / 1000 'mm to m
                
            End If
        Next
        
        
        
        
    Set rngType_ = Range(Range("GypsumType"), Range("GypsumType").End(xlDown))
    
            For Each rngType In rngType_        '외벽 석고보드
        
            If rngType = cmb05.Value Then
            
                For i = 1 To 3
                
                    arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                    .Offset(i + 6, REPLA_VALUE).Value = arrSpec(i)
                    
                Next
                
                temp() = Split(cmb06.Value, " ")
                .Offset(6, REPLA_VALUE).Value = CDbl(temp(0)) / 1000 'mm to m
                
            End If
        Next
        
        For Each rngType In rngType_        '측벽 석고보드
        
            If rngType = cmb11.Value Then
            
                For i = 1 To 3
                
                    arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                    .Offset(i + 14, REPLA_VALUE).Value = arrSpec(i)
                    
                Next
                
                temp() = Split(cmb12.Value, " ")
                .Offset(14, REPLA_VALUE).Value = CDbl(temp(0)) / 1000 'mm to m
                
            End If
        Next
        
    End With
    
    
    
    
    With Range("Repla_Insulation")
    
    Set rngType_ = Range(Range("InsulationType"), Range("InsulationType").End(xlDown))
    
        For Each rngType In rngType_        '외벽 단열재
        
            If rngType = cmb03.Value Then
            
                For i = 1 To 3
                
                    arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                    .Offset(i + 2, REPLA_VALUE).Value = arrSpec(i)
                    
                Next
                
                temp() = Split(cmb04.Value, " ")
                .Offset(2, REPLA_VALUE).Value = CDbl(temp(0)) / 1000 'mm to m
                
            End If
        Next
        
        For Each rngType In rngType_        '측벽 단열재
        
            If rngType = cmb09.Value Then
            
                For i = 1 To 3
                
                    arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                    .Offset(i + 6, REPLA_VALUE).Value = arrSpec(i)
                    
                Next
                
                temp() = Split(cmb10.Value, " ")
                .Offset(6, REPLA_VALUE).Value = CDbl(temp(0)) / 1000 'mm to m
                
            End If
        Next
        
    End With
    
    InputToCell
    
End Sub

Private Sub UserForm_Initialize()

    Dim rngType As Range
    Dim strType As String
    
    Set rngType_ = Range(Range("ConcreteType"), Range("ConcreteType").End(xlDown))
        
        For Each rngType In rngType_
            
            If rngType.Value <> "종류" Then
           
                cmb01.AddItem rngType.Value
                cmb07.AddItem rngType.Value
            
            End If
        
        Next rngType
    
    
    Set rngType_ = Range(Range("InsulationType"), Range("InsulationType").End(xlDown))
            
        For Each rngType In rngType_
            
            If rngType.Value <> "종류" Then
           
                cmb03.AddItem rngType.Value
                cmb09.AddItem rngType.Value
            
            End If
        
        Next rngType


    Set rngType_ = Range(Range("GypsumType"), Range("GypsumType").End(xlDown))
            
        For Each rngType In rngType_
            
            If rngType.Value <> "종류" Then
           
                cmb05.AddItem rngType.Value
                cmb11.AddItem rngType.Value
            
            End If
        
        Next rngType
        
        cmb01.Value = cmb01.list(0)
        cmb02.Value = cmb02.list(38)
        
        cmb03.Value = cmb03.list(0)
        cmb04.Value = cmb04.list(20)
        
        cmb05.Value = cmb05.list(0)
        cmb06.Value = cmb06.list(10)
        
        cmb07.Value = cmb07.list(0)
        cmb08.Value = cmb08.list(50)
        
        cmb09.Value = cmb09.list(0)
        cmb10.Value = cmb10.list(20)
        
        cmb11.Value = cmb11.list(0)
        cmb12.Value = cmb12.list(10)

        Image1.Picture = LoadPicture(ThisWorkbook.Path & "\files\image\wall\wallstructure_line.jpg")
        Image2.Picture = LoadPicture(ThisWorkbook.Path & "\files\image\wall\wallstructure_line.jpg")
End Sub


Private Sub cmb01_Change()

    Set rngTn_ = Range(Range("InsulationTn"), Range("InsulationTn").End(xlDown))
        
        With cmb02
            .Enabled = True
            .BackColor = &H80000005

            For Each rngTn In rngTn_
                
                If rngTn.Value <> "두께" Then

                    .AddItem rngTn.Value & " mm"

                End If
            
            Next rngTn
            
            .SetFocus
            
            Label11.Caption = cmb01.Value & " / Thk-" & cmb02.Value
            
        End With
    
End Sub

Private Sub cmb02_Change()

    Label11.Caption = cmb01.Value & " / Thk-" & cmb02.Value

End Sub

Private Sub cmb03_Change()

    Set rngTn_ = Range(Range("InsulationTn"), Range("InsulationTn").End(xlDown))
        
        With cmb04
            .Enabled = True
            .BackColor = &H80000005

            For Each rngTn In rngTn_
                
                If rngTn.Value <> "두께" Then

                    .AddItem rngTn.Value & " mm"

                End If
            
            Next rngTn
            
            .SetFocus
            
            Label12.Caption = cmb03.Value & " / Thk-" & cmb04.Value
            
        End With
    
End Sub

Private Sub cmb04_Change()

    Label12.Caption = cmb03.Value & " / Thk-" & cmb04.Value

End Sub

Private Sub cmb05_Change()

    Set rngTn_ = Range(Range("InsulationTn"), Range("InsulationTn").End(xlDown))
        
        With cmb06
            .Enabled = True
            .BackColor = &H80000005

            For Each rngTn In rngTn_
                
                If rngTn.Value <> "두께" Then

                    .AddItem rngTn.Value & " mm"

                End If
            
            Next rngTn
            
            .SetFocus
            
            Label13.Caption = cmb05.Value & " / Thk-" & cmb06.Value
            
        End With
    
End Sub

Private Sub cmb06_Change()

    Label13.Caption = cmb05.Value & " / Thk-" & cmb06.Value

End Sub

Private Sub cmb07_Change()

    Set rngTn_ = Range(Range("InsulationTn"), Range("InsulationTn").End(xlDown))
        
        With cmb08
            .Enabled = True
            .BackColor = &H80000005

            For Each rngTn In rngTn_
                
                If rngTn.Value <> "두께" Then

                    .AddItem rngTn.Value & " mm"

                End If
            
            Next rngTn
            
            .SetFocus
            
            Label14.Caption = cmb07.Value & " / Thk-" & cmb08.Value
            
        End With
    
End Sub

Private Sub cmb08_Change()

    Label14.Caption = cmb07.Value & " / Thk-" & cmb08.Value

End Sub

Private Sub cmb09_Change()

    Set rngTn_ = Range(Range("InsulationTn"), Range("InsulationTn").End(xlDown))
        
        With cmb10
            .Enabled = True
            .BackColor = &H80000005

            For Each rngTn In rngTn_
                
                If rngTn.Value <> "두께" Then

                    .AddItem rngTn.Value & " mm"

                End If
            
            Next rngTn
            
            .SetFocus
            
            Label15.Caption = cmb09.Value & " / Thk-" & cmb10.Value
            
        End With
    
End Sub

Private Sub cmb10_Change()

    Label15.Caption = cmb09.Value & " / Thk-" & cmb10.Value

End Sub

Private Sub cmb11_Change()

    Set rngTn_ = Range(Range("InsulationTn"), Range("InsulationTn").End(xlDown))
        
        With cmb12
            .Enabled = True
            .BackColor = &H80000005

            For Each rngTn In rngTn_
                
                If rngTn.Value <> "두께" Then

                    .AddItem rngTn.Value & " mm"

                End If
            
            Next rngTn
            
            .SetFocus
            
            Label16.Caption = cmb11.Value & " / Thk-" & cmb12.Value
            
        End With
    
End Sub

Private Sub cmb12_Change()

    Label16.Caption = cmb11.Value & " / Thk-" & cmb12.Value

End Sub


Private Sub InputToCell()

    With WallStructureForm
        
        For i = 1 To 6      '외벽
        
            Range("Cell_Cali_Wall").Cells(i, 1) = .Controls("cmb0" & i).Value
    
        Next
        
        For i = 8 To 13      '측벽
        
            If i < 11 Then
            
                Range("Cell_Cali_Wall").Cells(i, 1) = .Controls("cmb0" & i - 1).Value
            
            Else
            
                Range("Cell_Cali_Wall").Cells(i, 1) = .Controls("cmb" & i - 1).Value
            
            End If
    
        Next

    End With
End Sub
