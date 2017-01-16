VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ParameterForm 
   Caption         =   "자동보정 상세입력"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   OleObjectBlob   =   "ParameterForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "ParameterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancelButton_Click()

    Unload Me
    
End Sub

Private Sub CmdSetRangeButton_Click()
    
    SelectRangeWithMouse
    
End Sub

Private Sub startButton_Click()

Dim readyToRun As Boolean

    readyToRun = GetParameter
    
    If Not readyToRun Then
        
        MsgBox "모든 사항을 입력해주세요"

    Else
    
        'Unload Me
        StartCalibration
        'Progress Bar 만들기

    End If
    
End Sub

Sub SelectRangeWithMouse()

    Set thisRng = Application.InputBox("범위를 선택하세요", "범위 선택", Type:=8)
    

End Sub

