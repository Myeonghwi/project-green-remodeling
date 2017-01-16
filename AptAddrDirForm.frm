VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AptAddrDirForm 
   Caption         =   "공동주택 정보 직접입력"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "AptAddrDirForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "AptAddrDirForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelButton_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    Range("Result_AptSearch").Offset(-1, 0).Value = Blank
    Range("Result_AptSearch").Offset(i, 0).Interior.Color = RGB(210, 210, 255)
    Range("Result_AptSearch").Offset(i, 0).Locked = Ture
    
    For i = 0 To 16
    
        Range("Result_AptSearch").Offset(i, 0).Value = Blank
        Range("Result_AptSearch").Offset(i, 0).Interior.Color = RGB(210, 210, 255)
        Range("Result_AptSearch").Offset(i, 0).Locked = Ture
            
    Next
    
    cmb01.list = Array("중부지방", "남부지방", "제주도")
    cmb02.list = Array("36 m2", "46 m2", "59 m2", "84 m2", "125 m2")
    
'    Dim rngType As Range
'    Dim strType As String
'
'    Set rngType_ = Range(Range("InsulationType"), Range("InsulationType").End(xlDown))
'
'        cmb01.Clear
'        cmb03.Clear
'
'        For Each rngType In rngType_
'
'            If rngType.Value <> "종류" Then
'
'                cmb01.AddItem rngType.Value
'
'            End If
'
'        Next rngType
'
'        For Each rngType In rngType_
'
'            If rngType.Value <> "종류" Then
'
'                cmb03.AddItem rngType.Value
'
'            End If
'
'        Next rngType
'
'        cmb01.Value = cmb01.list(0)
'        cmb02.Value = cmb02.list(20)
'
'        cmb03.Value = cmb03.list(4)
'        cmb04.Value = cmb04.list(30)
End Sub
