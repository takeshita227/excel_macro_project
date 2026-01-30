Attribute VB_Name = "nikoMasterSearch"
Option Explicit

Public Const masterPath As String = "C:\Users\h_i_d\Desktop\testmaster\ko\"

' ▼ ComboBox1：detail_line-stop_testmaster の C列以降の列名を読み込む
Public Sub LoadnikoMasterToComboBox(targetCombo As MSForms.ComboBox, masterFileName As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fullPath As String
    Dim lastCol As Long
    Dim c As Long
    Dim val As String

    fullPath = masterPath & masterFileName

    ' 更新ダイアログを出さない
    Set wb = Workbooks.Open(Filename:=fullPath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Sheets(1)

    ' 1行目の最終列
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    targetCombo.Clear

    ' C列(3列目) ? 最終列の列名を追加
    For c = 3 To lastCol
        val = Trim(ws.Cells(1, c).Value)
        If val <> "" Then
            targetCombo.AddItem val
        End If
    Next c

    wb.Close SaveChanges:=False

End Sub


' ▼ ComboBox2：選ばれた列の TRUE 行の B列の値を読み込む
Public Sub Loadniko2MasterToComboBox(targetCombo As MSForms.ComboBox, masterFileName As String, selectedColumnName As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fullPath As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim r As Long
    Dim val As String

    fullPath = masterPath & masterFileName

    ' 更新ダイアログを出さない
    Set wb = Workbooks.Open(Filename:=fullPath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Sheets(1)

    ' 1行目から列名一致を探す
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    col = 0

    For r = 1 To lastCol
        If Trim(ws.Cells(1, r).Value) = selectedColumnName Then
            col = r
            Exit For
        End If
    Next r

    If col = 0 Then
        wb.Close False
        Exit Sub
    End If

    ' 該当列の最終行
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    targetCombo.Clear

    ' TRUE の行だけ B列を追加
    For r = 2 To lastRow
        If ws.Cells(r, col).Value = True Then
            val = Trim(ws.Cells(r, 2).Value)   ' ← B列
            If val <> "" Then
                targetCombo.AddItem val
            End If
        End If
    Next r

    wb.Close SaveChanges:=False

End Sub
