VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LineStop 
   Caption         =   "ライン停止内容"
   ClientHeight    =   11175
   ClientLeft      =   195
   ClientTop       =   795
   ClientWidth     =   8295.001
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
    TextBox0.Value = Format(Now, "hh:mm")
    
    TextBox0.Enabled = False
    TextBox3.Enabled = False
    TextBox0.BackColor = vbWhite
    TextBox3.BackColor = vbWhite
    
    TextBox5.Enabled = False
    TextBox5.BackColor = vbWhite
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

' ▼ 開始時刻を10分単位に丸め（常に切り捨て）
Private Function RoundBlockStart(timeStr As String) As String
    Dim parts() As String, h As Long, m As Long
    parts = Split(timeStr, ":")
    h = CLng(parts(0))
    m = CLng(parts(1))
    m = (m \ 10) * 10
    RoundBlockStart = Format(h, "00") & ":" & Format(m, "00")
End Function

' ▼ 終了時刻を10分単位に丸め（56分までは切り捨て、57分以降は切り上げ）
Private Function RoundBlockEnd(timeStr As String) As String
    Dim parts() As String, h As Long, m As Long
    parts = Split(timeStr, ":")
    h = CLng(parts(0))
    m = CLng(parts(1))
    
    If m <= 56 Then
        m = (m \ 10) * 10
    Else
        m = ((m \ 10) + 1) * 10
        If m = 60 Then
            h = h + 1
            m = 0
        End If
    End If
    
    RoundBlockEnd = Format(h, "00") & ":" & Format(m, "00")
End Function

' ▼ CommandButton3 を押したら開始～復旧の間と停止時間ブロックを塗ってフォームを閉じる
Private Sub CommandButton3_Click()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim startRow As Long, endRow As Long
    Dim startTimeRounded As String, endTimeRounded As String
    
    Set ws = ThisWorkbook.Sheets("生産状況")
    Set rng = ws.Range("C8:C73")   ' C列の対象範囲
    
    startRow = 0
    endRow = 0
    
    ' 入力時刻を10分単位に丸める
    startTimeRounded = RoundBlockStart(TextBox1.Value)
    endTimeRounded = RoundBlockEnd(TextBox2.Value)
    
    ' 開始時刻の行を探す
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            If Format(cell.Value, "hh:mm") = startTimeRounded Then
                startRow = cell.Row
                Exit For
            End If
        End If
    Next cell
    
    ' 終了時刻の行を探す
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            If Format(cell.Value, "hh:mm") = endTimeRounded Then
                endRow = cell.Row
                Exit For
            End If
        End If
    Next cell
    
    ' 開始～復旧の間を塗る（D列のみ）
    If startRow > 0 And endRow > 0 Then
        Dim r As Long
        For r = startRow To endRow
            ws.Cells(r, 4).Interior.Color = RGB(255, 200, 200) ' D列のみ塗る
        Next r
    End If
    
    ' フォームを閉じる
    Unload Me
End Sub
