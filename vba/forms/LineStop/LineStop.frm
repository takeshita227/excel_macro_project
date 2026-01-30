VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LineStop 
   Caption         =   "ライン停止内容"
   ClientHeight    =   12390
   ClientLeft      =   300
   ClientTop       =   750
   ClientWidth     =   7665
   OleObjectBlob   =   "LineStop.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "LineStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' フォーム読み込み時、TextBoxにデフォルト値を設定
Private Sub UserForm_Initialize()
    Dim selectedRow As Long
    Dim startTime As String
    
    selectedRow = ActiveCell.Row
    
    On Error Resume Next
    startTime = Cells(selectedRow, 3).Value
    On Error GoTo 0
    
    If startTime <> "" Then
        TextBox1.Value = Format(startTime, "hh:mm")
        TextBox2.Value = Format(startTime, "hh:mm")
    Else
        TextBox1.Value = ""
        TextBox2.Value = ""
    End If
    
    TextBox3.Value = "時間エラー"
    
    ' ▼ 現在時刻をセット
    TextBox0.Value = Format(Now, "hh:mm")
    
    TextBox0.Enabled = False
    TextBox3.Enabled = False
    TextBox0.BackColor = vbWhite
    TextBox3.BackColor = vbWhite
    
    TextBox5.Enabled = False
    TextBox5.BackColor = vbWhite
End Sub


' ▼ TextBox0 が消えたら復元
Private Sub TextBox0_Change()
    If TextBox0.Value = "" Then
        TextBox0.Value = Format(Now, "hh:mm")
    End If
End Sub

Private Sub TextBox0_Enter()
    TextBox0.Value = Format(Now, "hh:mm")
End Sub


' TextBox2（復旧時刻）の変更時に停止時間を計算
Private Sub TextBox2_Change()
    Call CalculateStopTime
End Sub

' TextBox1（発生時刻）の変更時に停止時間を計算
Private Sub TextBox1_Change()
    Call CalculateStopTime
End Sub


' 停止時間を計算する関数
Private Sub CalculateStopTime()
    Dim startSeconds As Long, recoverySeconds As Long
    Dim stopSeconds As Long, hours As Long, minutes As Long
    
    On Error GoTo ErrorHandler
    
    If TextBox1.Value = "" Or TextBox2.Value = "" Then
        TextBox3.Value = "時間エラー"
        Exit Sub
    End If
    
    startSeconds = TimeToSeconds(TextBox1.Value)
    recoverySeconds = TimeToSeconds(TextBox2.Value)
    
    If recoverySeconds < startSeconds Then
        TextBox3.Value = "時間エラー"
        Exit Sub
    End If
    
    stopSeconds = recoverySeconds - startSeconds
    hours = Int(stopSeconds / 3600)
    minutes = Int((stopSeconds Mod 3600) / 60)
    
    TextBox3.Value = Format(hours, "00") & ":" & Format(minutes, "00")
    Exit Sub
    
ErrorHandler:
    TextBox3.Value = "時間エラー"
End Sub


' 時刻文字列を秒単位に変換
Private Function TimeToSeconds(timeStr As String) As Long
    Dim parts() As String, hours As Long, minutes As Long
    
    On Error GoTo ErrorHandler
    parts = Split(timeStr, ":")
    
    If UBound(parts) >= 1 Then
        hours = CLng(parts(0))
        minutes = CLng(parts(1))
    End If
    
    TimeToSeconds = hours * 3600 + minutes * 60
    Exit Function
    
ErrorHandler:
    TimeToSeconds = 0
End Function


' ▼ TextBox4（担当者ID）が変更されたときに TextBox5（担当者名）を自動入力
Private Sub TextBox4_Change()
    Dim userId As String, userName As String
    
    userId = Me.TextBox4.Text
    
    If Len(userId) > 8 Then
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Exit Sub
    End If
    
    If Len(userId) < 8 Then
        Me.TextBox5.Text = ""
        Exit Sub
    End If
    
    userName = GetUserName(userId)
    Me.TextBox5.Text = userName
    
    If userName <> "" And userName <> "ファイルなし" And userName <> "ファイル開失敗" Then
        ThisWorkbook.Sheets("生産状況").Range("E4").Value = userName
    End If
End Sub


' ▼ TextBox4 にフォーカスが当たったときに中身を削除
Private Sub TextBox4_Enter()
    Me.TextBox4.Text = ""
    Me.TextBox5.Text = ""
End Sub


' ▼ CommandButton3（送信ボタン）
Private Sub CommandButton3_Click()

    Dim inputTime As String
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim matchRow As Long
    Dim stopTime As String
    Dim stopMinutes As Long
    Dim extraRows As Long
    Dim i As Long
    
    inputTime = Trim(Me.TextBox1.Value)
    
    If inputTime = "" Then
        MsgBox "発生時刻が入力されていません。", vbExclamation
        Exit Sub
    End If
    
    ' ▼ 10分単位に丸める
    On Error Resume Next
    inputTime = Format(RoundToNearest10Minutes(inputTime), "h:mm")
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets("生産状況")
    Set rng = ws.Range("C8:C73")
    
    matchRow = 0
    
    ' ▼ C8:C73 の中から一致する時刻を探す
    For Each cell In rng
        If Format(cell.Value, "h:mm") = inputTime Then
            matchRow = cell.Row
            Exit For
        End If
    Next cell
    
    If matchRow = 0 Then
        MsgBox "時間表に一致する時刻がありません。", vbExclamation
        Exit Sub
    End If
    
    ' ▼ 停止時間を取得
    stopTime = Me.TextBox3.Value
    
    ' ▼ 停止時間を分に変換
    On Error Resume Next
    stopMinutes = Hour(CDate(stopTime)) * 60 + Minute(CDate(stopTime))
    On Error GoTo 0
    
    ' ▼ 停止時間に応じて追加行を決定
    If stopMinutes <= 15 Then
        extraRows = 0
    Else
        extraRows = (stopMinutes - 16) \ 10 + 1
    End If
    
    ' ▼ 行を塗る（基本1行 + extraRows）
    For i = 0 To extraRows
        If matchRow + i <= 73 Then
            ws.Cells(matchRow + i, "D").Interior.Color = RGB(255, 200, 200)
        End If
    Next i
    
    
    ' ▼ フォームを閉じる
    Unload Me

End Sub

Private Sub ComboBox1_Enter()
    LoadnikoMasterToComboBox Me.ComboBox1, "detail_line-stop_testmaster.xlsx"
End Sub
