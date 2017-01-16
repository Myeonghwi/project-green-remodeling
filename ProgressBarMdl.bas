Attribute VB_Name = "ProgressBarMdl"
Option Explicit

Sub SetProgress()
    
Application.EnableEvents = False
Application.ScreenUpdating = False

Dim lngCounter As Long
Dim lngNumberOfTasks As Long

lngNumberOfTasks = 10000

Call ShowProgress(0, lngNumberOfTasks, _
                    "Excel is working on Task Number 1", False, _
                    "Progress Bar Test")


For lngCounter = 1 To lngNumberOfTasks

    Call ShowProgress(lngCounter, lngNumberOfTasks, _
                    "Excel is working on Task Number " & lngCounter + 1, False)
Next lngCounter

'Enable ScreenUpdating and Events
Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub

Sub ShowProgress(ByVal ActionNumber As Long, _
                ByVal TotalActions As Long, _
                Optional ByVal StatusMessage As String = vbNullString, _
                Optional ByVal CloseWhenDone As Boolean = True, _
                Optional ByVal Title As String = vbNullString)

DoEvents

If isFormOpen("ProgressBar") Then

    Call ProgressBar.UpdateForm(ActionNumber, TotalActions, StatusMessage)
    
Else

    ProgressBar.Show
    
    If Not Title = vbNullString Then
    
        ProgressBar.Caption = Title
        
    End If
    
    Call ProgressBar.UpdateForm(ActionNumber, TotalActions, StatusMessage)
    
End If

If CloseWhenDone And CBool(ActionNumber >= TotalActions) Then
    Unload ProgressBar
End If

End Sub

Function isFormOpen(ByVal FormName As String) As Boolean

Dim ufForm As Object

isFormOpen = False

For Each ufForm In VBA.UserForms

    If ufForm.Name = FormName Then

        isFormOpen = True

        Exit For
    End If
Next ufForm
End Function

