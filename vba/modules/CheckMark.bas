Attribute VB_Name = "CheckMark"
Public Sub ToggleMaru(Target As Range)
    If Target.Value = "" Then
        Target.ClearContents
    Else
        Target.Value = ""
    End If
End Sub
