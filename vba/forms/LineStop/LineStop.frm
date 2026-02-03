VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LineStop 
   Caption         =   "ライン停止内容"
   ClientHeight    =   12396
   ClientLeft      =   336
   ClientTop       =   852
   ClientWidth     =   9564.001
   OleObjectBlob   =   "LineStop.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "LineStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' フォーム読み込み時、初期設定
'==========================================================
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

    ' ▼ 時計 TextBox（共通関数）
    SetupClockTextBox Me.TextBox0

    ' ▼ 停止時間表示欄は編集不可
    TextBox3.Enabled = False
    TextBox3.BackColor = vbWhite

    ' ▼ 担当者名は編集不可
    TextBox5.Enabled = False
    TextBox5.BackColor = vbWhite

End Sub


'==========================================================
' TextBox0（時計）保護
'==========================================================
Private Sub TextBox0_Change()
    RestoreClockIfEmpty Me.TextBox0
End Sub

Private Sub TextBox0_Enter()
    RestoreClockIfEmpty Me.TextBox0
End Sub


'==========================================================
' 停止時間の自動計算
'==========================================================
Private Sub TextBox2_Change()
    CalculateStopTime
End Sub

Private Sub TextBox1_Change()
    CalculateStopTime
End Sub


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
    hours = stopSeconds \ 3600
    minutes = (stopSeconds Mod 3600) \ 60

    TextBox3.Value = Format(hours, "00") & ":" & Format(minutes, "00")
    Exit Sub

ErrorHandler:
    TextBox3.Value = "時間エラー"

End Sub


'==========================================================
' 時刻文字列 → 秒
'==========================================================
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


'==========================================================
' 担当者ID → 担当者名
'==========================================================
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


Private Sub TextBox4_Enter()
    Me.TextBox4.Text = ""
    Me.TextBox5.Text = ""
End Sub


'==========================================================
' 書き込み処理
'==========================================================
Private Sub CommandButton3_Click()

    Dim inputTime As String
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim matchRow As Long
    Dim stopTime As String
    Dim stopMinutes As Long
    Dim stopSeconds As Long
    Dim extraRows As Long
    Dim i As Long

    inputTime = Trim(Me.TextBox1.Value)

    If inputTime = "" Then
        MsgBox "発生時刻が入力されていません。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    inputTime = Format(RoundToNearest10Minutes(inputTime), "h:mm")
    On Error GoTo 0

    Set ws = ThisWorkbook.Sheets("生産状況")
    Set rng = ws.Range("C8:C73")

    matchRow = 0

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

    stopTime = Me.TextBox3.Value
    stopSeconds = TimeToSeconds(stopTime)
    stopMinutes = stopSeconds \ 60

    If stopMinutes <= 15 Then
        extraRows = 0
    Else
        extraRows = ((stopMinutes - 16) \ 10) + 1
    End If

    For i = 0 To extraRows
        If matchRow + i <= 73 Then
            ws.Cells(matchRow + i, "D").Interior.Color = RGB(255, 200, 200)
        End If
    Next i

    ws.Cells(matchRow, 30).Value = Me.ComboBox1.Value
    ws.Cells(matchRow, 31).Value = Me.ComboBox2.Value
    ws.Cells(matchRow, 32).Value = Me.ComboBox3.Value & "｜" & Me.ComboBox4.Value

    Unload Me

End Sub


Private Sub ComboBox1_Enter()
    LoadnikoMasterToComboBox Me.ComboBox1, "detail_line-stop_testmaster.xlsx"
End Sub

Private Sub ComboBox2_Enter()
    If Me.ComboBox1.Value = "" Then Exit Sub

    Loadniko2MasterToComboBox Me.ComboBox2, _
                              "detail_line-stop_testmaster.xlsx", _
                              Me.ComboBox1.Value
End Sub

