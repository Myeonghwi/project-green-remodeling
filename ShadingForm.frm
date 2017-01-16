VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ShadingForm 
   Caption         =   "차양"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "ShadingForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "ShadingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelButton_Click()

    Unload Me

End Sub


Private Sub UserForm_Initialize()

    Set rngType_ = Range(Range("ShadingType"), Range("ShadingType").End(xlDown))

        For Each rngType In rngType_

            If rngType.Value <> "종류" Then

                cmb01.AddItem rngType.Value

            End If

        Next rngType

End Sub

Private Sub CmdOKButton_Click()

    Set rngType_ = Range(Range("ShadingType"), Range("ShadingType").End(xlDown))
    
    With Range("Repla_Shading")
    
        ReDim arrSpec(1 To 8)
        
        For Each rngType In rngType_
        
            If rngType = cmb01.Value Then
            
                For i = 1 To 8
                
                    arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                    .Offset(i + 1, REPLA_VALUE).Value = arrSpec(i)
                    
                Next
            End If
        Next
        
    End With
    
    InputToCell
    
End Sub

Private Sub cmb01_Change()

    Set rngType_ = Range(Range("ShadingType"), Range("ShadingType").End(xlDown))
    
    With ShadingForm
    
        ReDim arrSpec(1 To 6)
        
        For Each rngType In rngType_
        
            If rngType = cmb01.Value Then
            
                For i = 1 To 3
                    arrSpec(i) = CDbl(rngType.Offset(0, i).Value)
                Next
                
                For i = 4 To 6
                    arrSpec(i) = CDbl(rngType.Offset(0, i + 5).Value)
                Next
                
                For i = 1 To 6
                    .Controls("Label1" & i).Caption = arrSpec(i)
                Next
            
                'Image1.Picture = LoadPicture(ThisWorkbook.Path & "\files\image\lighting\" & rngType & ".jpg")
                
            End If
        
        Next rngType
    
    End With
    
End Sub

Private Sub InputToCell()

    With ShadingForm
     
        Range("Cell_Main_Shading").Cells(1, 1) = cmb01.Value
        
    End With
End Sub
