VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} openingcheck 
   Caption         =   "名前入力"
   ClientHeight    =   4905
   ClientLeft      =   180
   ClientTop       =   675
   ClientWidth     =   7830
   OleObjectBlob   =   "openingcheck.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "openingcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub user_id_Change()
    Dim userId As String
    Dim userName As String
    
    userId = Me.user_id.Text
    
    ' 8桁揃うまで何もしない
    If Len(userId) < 8 Then
        Me.user_name.Text = ""
        Exit Sub
    End If
    
    ' 標準モジュールの関数を呼ぶ
    userName = GetUserName(userId)
    Me.user_name.Text = userName

    ' E4 に書き込む
    If userName <> "" And userName <> "ファイルなし" And userName <> "ファイル開失敗" Then
        ThisWorkbook.Sheets("生産状況").Range("E4").Value = userName
    End If

    ' フォームを閉じる
    Unload Me
End Sub

