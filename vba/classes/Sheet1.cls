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

' E4 を選択したとき → openingcheck を開く
If Target.Address = "$E$4" Then
    openingcheck.Tag = Target.Address
    openingcheck.Show
    Application.ScreenUpdating = True
    Exit Sub
End If

' AC8:AC73 を選択したとき → TenKey を開く
If Not Intersect(Target, Me.Range("AC8:AC73")) Is Nothing Then
    TenKey.Tag = Target.Address
    TenKey.Show
    Application.ScreenUpdating = True
    Exit Sub
End If

Application.ScreenUpdating = True   ' 念のため

End Sub












