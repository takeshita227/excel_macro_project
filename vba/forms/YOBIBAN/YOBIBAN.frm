VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} YOBIBAN 
   Caption         =   "呼番登録"
   ClientHeight    =   7920
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

    ' Enterキーが押されたときだけ処理
    If KeyCode = vbKeyReturn Then

        rawQR = Trim(TextBox1.Text)

        ' 文字数チェック
        If Len(rawQR) <> 75 Then
            MsgBox "スキャンするQRコード違います！完成品かんばんをスキャンしてください", vbExclamation
            TextBox1.Value = ""     ' 初期化
            TextBox2.Value = ""     ' 初期化
            KeyCode = 0
            Exit Sub
        End If

        ' 26文字目から18文字を抜き出す（あなたの仕様）
        displayID = Mid(rawQR, 26, 18)
        TextBox2.Value = displayID

        ' QR入力欄はクリアして次回スキャンに備える
        TextBox1.Value = ""

        ' Enterをフォームに渡さない
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
