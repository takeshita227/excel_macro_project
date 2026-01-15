Attribute VB_Name = "ToggleMaru"
Public Sub ToggleMaru(Target As Range)
    If Target.Value = "" Then
        Target.ClearContents
    Else
        Target.Value = ""
    End If
End Sub
