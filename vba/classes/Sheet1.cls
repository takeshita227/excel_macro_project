VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Application.ScreenUpdating = False   ' ★ 画面更新停止

    ' ▼ C4 を選択したとき → LineIDCheck を起動
    If Target.Address = "$C$4" Then
        LineIDCheck.Show
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' ▼ E4 を選択したとき → openingcheck を表で表示
    If Target.Address = "$E$4" Then
        openingcheck.Tag = Target.Address
        openingcheck.Show
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' ▼ E8:E73 を選択したとき → YOBIBAN を起動
    If Not Intersect(Target, Me.Range("E8:E73")) Is Nothing Then
        YOBIBAN.Tag = Target.Address
        YOBIBAN.Show
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' ▼ AC8:AC73 を選択したとき → TenKey を起動
    If Not Intersect(Target, Me.Range("AC8:AC73")) Is Nothing Then
        TenKey.Tag = Target.Address
        TenKey.Show
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' ▼ AD8:AD73 を選択したとき → LineStop を起動
    If Not Intersect(Target, Me.Range("AD8:AD73")) Is Nothing Then
        LineStop.Tag = Target.Address
        LineStop.Show
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Application.ScreenUpdating = True   ' 念のため

End Sub
