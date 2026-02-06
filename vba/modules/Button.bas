Attribute VB_Name = "Button"
Sub ボタン379_Click()
Defective_in.Show
End Sub
Sub ボタン380_Click() '完成品入力ボタン
Defective_assy.Show
End Sub
Sub ボタン381_Click()
Defective_parts.Show
End Sub
'不良数のシートへ移動
Sub ボタン338_Click() '不良品閲覧ボタン
Sheets("不良数").Select
End Sub
'生産状況のシートへ移動
Sub ボタン1_Click()  '生産状況へもどるボタン
Sheets("生産状況").Select
End Sub

