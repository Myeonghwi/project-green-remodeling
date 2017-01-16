VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10455
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()

    For i = 0 To ListBox1.ListCount - 1
        
        If ListBox1.Selected(i) = True Then
        
            ListBox2.AddItem ListBox1.list(i) & "|" & cmb01.Value & "|" & cmb02.Value
        
        End If
        
    Next i
            
End Sub

Private Sub btnDelete_Click()

    Dim counter As Integer
    
    'counter = 0
    
    For i = 0 To ListBox2.ListCount - 1
    
        If ListBox2.Selected(i) = True Then
            ListBox2.RemoveItem (i)
        '    counter = counter + 1
        End If
        
    Next i

End Sub


Private Sub ListBox1_Click()

    Set rngType_ = Range(Range("InsulationType"), Range("InsulationType").End(xlDown))
    
    For Each rngType In rngType_
        For i = 0 To ListBox1.ListCount - 1
            
            If ListBox1.Selected(i) = True Then
        
                Dim temp

                temp = Split(ListBox1.Value, " ")
                
                Image1.Picture = LoadPicture(ThisWorkbook.Path & "\files\image\insulation\" & temp(0) & ".jpg")
            
            End If
        
        Next i
    Next rngType
    
End Sub

Private Sub UserForm_Initialize()

    

    Set rngType_ = Range(Range("InsulationType"), Range("InsulationType").End(xlDown))
    
    For Each rngType In rngType_
    
        If rngType.Value <> "종류" Then
     
            ListBox1.AddItem rngType.Value
        
        End If
        
    Next rngType
    
    
    
    Set rngTn_ = Range(Range("InsulationTn"), Range("InsulationTn").End(xlDown))
    
    
    For Each rngTn In rngTn_
        
        If rngTn.Value <> "두께" Then
            
            cmb01.AddItem rngTn.Value & "mm"
        
        End If
        
    Next rngTn
    
    cmb02.list = Array("외벽", "천장", "바닥")
    
    
End Sub


