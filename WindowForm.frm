VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WindowForm 
   Caption         =   "창호 선택"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   4065
   ClientWidth     =   5760
   OleObjectBlob   =   "WindowForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "WindowForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelButton_Click()

    Unload Me
    
End Sub

Private Sub UserForm_Initialize()

    Dim rngType As Range
    Dim strType As String
    
    Set rngType_ = Range(Range("WindowType"), Range("WindowType").End(xlDown))
        
        For Each rngType In rngType_
            
            If rngType.Value <> "종류" Then
           
                cmb01.AddItem rngType.Value
            
            End If
        
        Next rngType
        
        cmb01.Value = cmb01.list(0)
End Sub

Private Sub CmdOKButton_Click()

    Set rngType_ = Range(Range("WindowType"), Range("WindowType").End(xlDown))
    
    With Range("Repla_Window")
    
        ReDim arrSpec(1 To 3)
        
        For Each rngType In rngType_
        
            If rngType = cmb01.Value Then
            
                For i = 1 To 3
                
                    arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                    .Offset(i + 1, REPLA_VALUE).Value = arrSpec(i)
                    
                Next
            End If
        Next
        
    End With
    
    InputToCell
    
End Sub

Private Sub cmb01_Change()

    Dim rngType As Range
    Dim strType As String
    
    Set rngType_ = Range(Range("WindowType"), Range("WindowType").End(xlDown))
    
    With WindowForm
    
        ReDim arrSpec(1 To 3)
        
        For Each rngType In rngType_
        
            If rngType = cmb01.Value Then
            
                For i = 1 To 3
                
                    arrSpec(i) = CDbl(rngType.Cells(1, i + 1).Value)
                    
                Next
                
                For i = 1 To 3
                    .Controls("Label1" & i).Caption = arrSpec(i)
                Next
                
            End If
            
        Next rngType
        
    End With
End Sub

Private Sub InputToCell()

    With WindowForm
     
        Range("Cell_Main_Window").Cells(1, 1) = cmb01.Value
        
    End With
End Sub
