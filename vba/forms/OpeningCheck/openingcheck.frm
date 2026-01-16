VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} openingcheck 
   Caption         =   "名前入力"
   ClientHeight    =   3144
   ClientLeft      =   156
   ClientTop       =   576
   ClientWidth     =   6264
   OleObjectBlob   =   "openingcheck.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "openingcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' user_idの値が変更されたときにuser_nameを自動入力
Private Sub user_id_Change()
    Dim userId As String
    Dim userName As String
    
    userId = Me.user_id.Text
    
    ' ★ 8桁揃うまで何もしない（チカチカ防止）
    If Len(userId) < 8 Then
        Me.user_name.Text = ""
        Exit Sub
    End If
    
    ' 8桁揃ったらマスターファイルから検索
    userName = GetUserName(userId)
    Me.user_name.Text = userName

    ' ★ user_name が取得できたら E4 に書き込む
    If userName <> "" And userName <> "ファイルなし" And userName <> "ファイル開失敗" Then
        ThisWorkbook.Sheets("生産状況").Range("E4").Value = userName
    End If

    ' ★ openingcheck を閉じる
    Unload Me
End Sub


' マスターファイルからuser_nameを検索する関数
Private Function GetUserName(userId As String) As String
    Dim masterFile As String
    Dim masterBook As Workbook
    Dim masterSheet As Worksheet
    Dim searchRange As Range
    Dim foundCell As Range
    Dim rowIndex As Long
    
    masterFile = "C:\Users\h_i_d\Desktop\ProductionSystem\master\excel\user_master.xlsx"
    
    ' マスターファイルが存在するか確認
    If Dir(masterFile) = "" Then
        GetUserName = "ファイルなし"
        Exit Function
    End If
    
    ' マスターファイルを開く
    On Error Resume Next
    Set masterBook = Workbooks.Open(masterFile, , True)
    On Error GoTo 0
    
    If masterBook Is Nothing Then
        GetUserName = "ファイル開失敗"
        Exit Function
    End If
    
    Set masterSheet = masterBook.Sheets("Sheet1")
    Set searchRange = masterSheet.Range("A:A")
    
    ' user_idを検索
    Set foundCell = searchRange.Find(What:=userId, LookAt:=xlWhole, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        ' 同じ行のB列を取得
        rowIndex = foundCell.Row
        GetUserName = masterSheet.Range("B" & rowIndex).Value
    Else
        GetUserName = ""
    End If
    
    ' マスターファイルを閉じる
    masterBook.Close SaveChanges:=False
    
End Function
