VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CaliCheckForm 
   Caption         =   "�ڵ�����"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3975
   OleObjectBlob   =   "CaliCheckForm.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "CaliCheckForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelButton_Click()

    Unload Me
    
End Sub

Private Sub CmdStartButton_Click()
    
    StartCalibrationAuto
    
End Sub

Private Sub UserForm_Initialize()

    '����Ʈ�����
    For i = 1 To 4
        
        Controls("Label" & i).Caption = Sheet1.Range("I" & (i + 15)).Value
        
    Next
End Sub
