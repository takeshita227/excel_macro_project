Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("AC8:AC73")) Is Nothing Then
        TenKey.Show
    End If
End Sub
