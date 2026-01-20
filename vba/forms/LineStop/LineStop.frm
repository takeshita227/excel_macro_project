VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LineStop 
   Caption         =   "ライン停止内容"
   ClientHeight    =   8940.001
   ClientLeft      =   132
   ClientTop       =   576
   ClientWidth     =   5328
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

' ▼ TextBox4（担当者ID）が変更されたときに TextBox5（担当者名）を自動入力
Private Sub TextBox4_Change()
    Dim userId As String, userName As String
    
    userId = Me.TextBox4.Text
    
    ' 8桁を超えた場合はリセット
    If Len(userId) > 8 Then
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Exit Sub
    End If
    
    ' 8桁揃うまで何もしない
    If Len(userId) < 8 Then
        Me.TextBox5.Text = ""
        Exit Sub
    End If
    
    ' 8桁ちょうど揃ったら検索
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

' マスターファイルから user_name を検索する関数
Private Function GetUserName(userId As String) As String
    Dim masterFile As String, masterBook As Workbook
    Dim masterSheet As Worksheet, searchRange As Range
    Dim foundCell As Range, rowIndex As Long
    
    masterFile = "C:\Users\h_i_d\Desktop\ProductionSystem\master\excel\user_master.xlsx"
    
    If Dir(masterFile) = "" Then
        GetUserName = "ファイルなし"
        Exit Function
    End If
    
    On Error Resume Next
    Set masterBook = Workbooks.Open(masterFile, , True)
    On Error GoTo 0
    
    If masterBook Is Nothing Then
        GetUserName = "ファイル開失敗"
        Exit Function
    End If
    
    Set masterSheet = masterBook.Sheets("Sheet1")
    Set searchRange = masterSheet.Range("A:A")
    
    Set foundCell = searchRange.Find(What:=userId, LookAt:=xlWhole, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        rowIndex = foundCell.Row
        GetUserName = masterSheet.Range("B" & rowIndex).Value
    Else
        GetUserName = ""
    End If
    
    masterBook.Close SaveChanges:=False
End Function
