Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    ' E4 を選択したとき → openingcheck を開く
    If Target.Address = "$E$4" Then
        openingcheck.Tag = Target.Address
        openingcheck.Show
        Exit Sub
    End If

    ' AC8:AC73 を選択したとき → TenKey を開く
    If Not Intersect(Target, Me.Range("AC8:AC73")) Is Nothing Then
        TenKey.Tag = Target.Address
        TenKey.Show
        Exit Sub
    End If

End Sub
