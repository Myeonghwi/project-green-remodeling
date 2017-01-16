VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LightingForm 
   Caption         =   "조명"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "LightingForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "LightingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelButton_Click()

    Unload Me
    
End Sub

Private Sub UserForm_Initialize()

    'Dim rngType As Range
    Dim arrLightingType
    
    arrLightingType = Array("형광등", "LED조명")
    
    Set rngType_ = Range(Range("LightingSetupType"), Range("LightingSetupType").End(xlDown))
    
        cmb01.list = arrLightingType
        
        For Each rngType In rngType_
        
            If rngType.Value <> "조명 설치 형태" Then

                cmb02.AddItem rngType.Value
                
            End If
            
        Next rngType
        
        cmb01.Value = arrLightingType(0)
        cmb02.Value = cmb02.list(0)
        
End Sub

Private Sub CmdOKButton_Click()

    Set rngType_ = Range(Range("LightingSetupType"), Range("LightingSetupType").End(xlDown))
    
    With Range("Repla_Lighting")
    
        ReDim arrSpec(1 To 4)
        
        For Each rngType In rngType_
        
            If rngType = cmb02.Value Then
            
                For i = 1 To 4
                
                    arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                    .Offset(i + 1, REPLA_VALUE).Value = arrSpec(i)
                    
                Next
            End If
        Next
        
    End With
    
    InputToCell
    
End Sub

Private Sub CmdDirInputButton_Click()

    LightingDirForm.Show
    
End Sub

Private Sub cmb01_Change()

    If cmb01.Value = "형광등" Then
        Label11.Caption = "847"
    Else
        Label11.Caption = "479"
    End If

End Sub

Private Sub cmb02_Change()

    Dim rngType As Range
    
    Set rngType_ = Range(Range("LightingSetupType"), Range("LightingSetupType").End(xlDown))
    
    With LightingForm
    
        ReDim arrSpec(1 To 3)
        
        For Each rngType In rngType_
        
            If rngType = cmb02.Value Then
            
                For i = 1 To 3
                    arrSpec(i) = CDbl(rngType.Offset(0, i + 1).Value)
                Next
                
                For i = 1 To 3
                    .Controls("Label1" & i + 1).Caption = arrSpec(i)
                Next
            
                Image1.Picture = LoadPicture(ThisWorkbook.Path & "\files\image\lighting\" & rngType & ".jpg")
                
            End If
        
        Next rngType
    
    End With
End Sub

Private Sub InputToCell()

    With LightingForm
        
        For i = 1 To 2
        
            Range("Cell_Cali_Lighting").Cells(i, 1) = .Controls("cmb0" & i).Value
            Range("Cell_Main_Lighting").Cells(i, 1) = .Controls("cmb0" & i).Value
    
        Next
        
        For i = 1 To 4
        
            Range("Cell_Cali_Lighting").Cells(i + 2, 1) = CDbl(.Controls("Label1" & i).Caption)
            
        Next

    End With
End Sub
