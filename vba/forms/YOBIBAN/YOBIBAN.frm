VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} YOBIBAN 
   Caption         =   "呼番登録"
   ClientHeight    =   6336
   ClientLeft      =   168
   ClientTop       =   696
   ClientWidth     =   8208.001
   OleObjectBlob   =   "YOBIBAN.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "YOBIBAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' フォーム初期化：TextBox0 を時計としてセット
'==========================================================
Private Sub UserForm_Initialize()

    ' ▼ 時計 TextBox（編集不可）
    SetupClockTextBox Me.TextBox0

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
' Enter を完全無効化（フォーム送信防止）
'==========================================================
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0
End Sub


'==========================================================
' TextBox1（QR入力）Enter が押された瞬間だけ処理
'==========================================================
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Dim rawQR As String
    Dim displayID As String

    If KeyCode = vbKeyReturn Then

        rawQR = Trim(TextBox1.Text)

        If Len(rawQR) <> 75 Then
            MsgBox "スキャンするQRコード違います！完成品かんばんをスキャンしてください", vbExclamation
            TextBox1.Value = ""
            TextBox2.Value = ""
            KeyCode = 0
            Exit Sub
        End If

        displayID = Mid(rawQR, 26, 18)
        TextBox2.Value = displayID

        TextBox1.Value = ""

        KeyCode = 0
    End If

End Sub


'==========================================================
' CommandButton3 / CommandButton4 → セルに書き込み & フォーム閉じる
'==========================================================
Private Sub CommandButton3_Click()
    WriteYOBIBANValue
End Sub

Private Sub CommandButton4_Click()
    WriteYOBIBANValue
End Sub


'==========================================================
' 共通処理：選択セル（Tag）に TextBox2 の値を書き込む
'==========================================================
Private Sub WriteYOBIBANValue()

    Dim ws As Worksheet
    Dim targetCell As Range

    Set ws = ThisWorkbook.Sheets("生産状況")
    Set targetCell = ws.Range(Me.Tag)

    targetCell.Value = TextBox2.Value

    Unload Me

End Sub
