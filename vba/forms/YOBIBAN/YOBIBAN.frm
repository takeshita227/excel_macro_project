VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} YOBIBAN 
   Caption         =   "呼番登録"
   ClientHeight    =   5085
   ClientLeft      =   135
   ClientTop       =   570
   ClientWidth     =   6570
   OleObjectBlob   =   "YOBIBAN.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "YOBIBAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton3_Click()

    Dim ws As Worksheet
    Dim targetCell As Range
    Dim writeValue As String
    Dim r As Long
    Dim nextRow As Long
    
    Set ws = ThisWorkbook.Sheets("生産状況")
    
    ' ▼ YOBIBAN.Tag に入っているセルを取得
    Set targetCell = ws.Range(Me.Tag)
    
    ' ▼ 書き込む値（TextBox2 の内容）
    writeValue = Me.TextBox2.Value
    
    ' ▼ 選択されていたセルに書き込む
    If writeValue <> "" Then
        targetCell.Value = writeValue
    End If
    
    ' ▼ AI9?AI17 の中で「一番上の空きセル」を探す
    nextRow = 0
    For r = 9 To 17
        If ws.Cells(r, "AI").Value = "" Then
            nextRow = r
            Exit For
        End If
    Next r
    
    ' ▼ 空きが見つかったときだけ書き込む
    If nextRow <> 0 Then
        ws.Cells(nextRow, "AI").Value = writeValue
    End If
    
    ' ▼ フォームを閉じる
    Unload Me

End Sub
