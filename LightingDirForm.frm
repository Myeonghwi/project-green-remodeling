VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LightingDirForm 
   Caption         =   "등기구 상세입력"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760.001
   OleObjectBlob   =   "LightingDirForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "LightingDirForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelButton_Click()

    Unload Me

End Sub

Private Sub UserForm_Initialize()

    With LightingDirForm
    
        For i = 1 To 15
            
            .Controls("cmb" & i).Enabled = False
            .Controls("cmb" & i).Text = ""
            .Controls("cmb" & i).BackColor = &H80000016
            
            .Controls("TextBox" & i).Enabled = False
            .Controls("TextBox" & i).Text = ""
            .Controls("TextBox" & i).BackColor = &H80000016
            
            .Controls("SpinButton" & i).Enabled = False
            
        Next

    End With
End Sub


Private Sub ChkRngBox1_Click()

    With LightingDirForm
    
        If ChkRngBox1.Value = True Then
        
            cmb1.Enabled = True
            cmb1.Text = ""
            cmb1.BackColor = &H80000005
            
            TextBox1.Enabled = True
            TextBox1.Text = ""
            TextBox1.BackColor = &H80000005
            
            SpinButton1.Enabled = True
            
        Else
        
            cmb1.Enabled = False
            cmb1.Text = ""
            cmb1.BackColor = &H80000016
            
            TextBox1.Enabled = False
            TextBox1.Text = ""
            TextBox1.BackColor = &H80000016
            
            SpinButton1.Enabled = False
        
        End If
        
    End With
End Sub

Private Sub SpinButton1_Change()

    With LightingDirForm
    
        If SpinButton1.Enabled = True Then
        
            TextBox1.Value = SpinButton1.Value
    
        End If
        
    End With

End Sub
