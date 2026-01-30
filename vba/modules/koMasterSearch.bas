Attribute VB_Name = "koMasterSearch"
Option Explicit

Public Const masterPath As String = "C:\Users\h_i_d\Desktop\testmaster\ko\"

' ▼ B列の値だけを ComboBox にセットするバージョン
Public Sub LoadkoMasterToComboBox(targetCombo As MSForms.ComboBox, masterFileName As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fullPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim val As String

    ' フルパス作成
    fullPath = masterPath & masterFileName

    ' 外部リンク警告を出さないためにフルパスで開く
    Set wb = Workbooks.Open(fullPath, ReadOnly:=True)

    ' シート名が Sheet1 の前提
    Set ws = wb.Sheets("Sheet1")

    ' B列の最終行
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' ComboBox 初期化
    targetCombo.Clear

    ' B列の値をすべて追加（空白はスキップ）
    For i = 2 To lastRow
        val = Trim(ws.Cells(i, 2).Value)
        If val <> "" Then
            targetCombo.AddItem val
        End If
    Next i

    ' 読み込んだら閉じる
    wb.Close SaveChanges:=False

End Sub
