Attribute VB_Name = "nikoMasterSearch"
Option Explicit

Public Const masterPath As String = "C:\Users\h_i_d\Desktop\testmaster\ko\"

' ▼ C列以降の列名（1行目の文字）を ComboBox にセットする
Public Sub LoadnikoMasterToComboBox(targetCombo As MSForms.ComboBox, masterFileName As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fullPath As String
    Dim lastCol As Long
    Dim c As Long
    Dim val As String

    fullPath = masterPath & masterFileName

    ' ★ 更新する/しないダイアログを絶対に出さない
    Set wb = Workbooks.Open(Filename:=fullPath, ReadOnly:=True, UpdateLinks:=False)

    Set ws = wb.Sheets(1)

    ' ★ 1行目の最終列を取得
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    targetCombo.Clear

    ' ★ C列(3列目) ? 最終列までの列名を選択肢に追加
    For c = 3 To lastCol
        val = Trim(ws.Cells(1, c).Value)
        If val <> "" Then
            targetCombo.AddItem val
        End If
    Next c

    ' ★ 読み込んだら閉じる（リンク扱いにならない）
    wb.Close SaveChanges:=False

End Sub
