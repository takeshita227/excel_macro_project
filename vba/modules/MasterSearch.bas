Attribute VB_Name = "MasterSearch"
Option Explicit

Public Const masterPath As String = "C:\Users\h_i_d\Desktop\testmaster\ko\"

' ▼ TRUE 行の B列を取得して ListBox にセットする共通関数
Public Sub LoadMasterToListBox(targetListBox As MSForms.ListBox, masterFileName As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fullPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim items As Collection
    Dim itemValue As String

    fullPath = masterPath & masterFileName

    ' マスターを開く（既に開いていれば再利用）
    On Error Resume Next
    Set wb = Workbooks(masterFileName)
    On Error GoTo 0

    If wb Is Nothing Then
        Set wb = Workbooks.Open(fullPath)
    End If

    Set ws = wb.Sheets(1) ' 必要ならシート名に変更

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set items = New Collection

    ' ▼ C列が TRUE の行だけ抽出し、B列を追加
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = True Then
            itemValue = ws.Cells(i, 2).Value   ' ★ B列
            items.Add itemValue
        End If
    Next i

    ' ▼ ListBox に反映
    targetListBox.Clear
    For i = 1 To items.Count
        targetListBox.AddItem items(i)
    Next i

    ' ▼ ListBox の高さを項目数に応じて自動調整
    Dim baseHeight As Single
    Dim itemHeight As Single

    itemHeight = 12   ' 1行あたりの高さ
    baseHeight = 6    ' 余白

    targetListBox.Height = baseHeight + (items.Count * itemHeight)

End Sub
