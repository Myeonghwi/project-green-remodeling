VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ParameterForm 
   Caption         =   "�ڵ����� ���Է�"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   OleObjectBlob   =   "ParameterForm.frx":0000
   StartUpPosition =   1  '������ ���
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
        
        MsgBox "��� ������ �Է����ּ���"

    Else
    
        'Unload Me
        StartCalibration
        'Progress Bar �����

    End If
    
End Sub

Sub SelectRangeWithMouse()

    Set thisRng = Application.InputBox("������ �����ϼ���", "���� ����", Type:=8)
    

End Sub

