VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TenKey 
   Caption         =   "数字入力"
   ClientHeight    =   5856
   ClientLeft      =   168
   ClientTop       =   696
   ClientWidth     =   5268
   OleObjectBlob   =   "TenKey.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "TenKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 数字 1?9
Private Sub NumberButton1_Click(): TextBox1.Text = TextBox1.Text & "1": End Sub
Private Sub NumberButton2_Click(): TextBox1.Text = TextBox1.Text & "2": End Sub
Private Sub NumberButton3_Click(): TextBox1.Text = TextBox1.Text & "3": End Sub
Private Sub NumberButton4_Click(): TextBox1.Text = TextBox1.Text & "4": End Sub
Private Sub NumberButton5_Click(): TextBox1.Text = TextBox1.Text & "5": End Sub
Private Sub NumberButton6_Click(): TextBox1.Text = TextBox1.Text & "6": End Sub
Private Sub NumberButton7_Click(): TextBox1.Text = TextBox1.Text & "7": End Sub
Private Sub NumberButton8_Click(): TextBox1.Text = TextBox1.Text & "8": End Sub
Private Sub NumberButton9_Click(): TextBox1.Text = TextBox1.Text & "9": End Sub

' 0
Private Sub NumberButton10_Click()
    TextBox1.Text = TextBox1.Text & "0"
End Sub

' 全削除（AC）
Private Sub NumberButton12_Click()
    TextBox1.Text = ""
End Sub

' 00
Private Sub NumberButton11_Click()
    TextBox1.Text = TextBox1.Text & "00"
End Sub

' 小数点
Private Sub NumberButton13_Click()
    If InStr(TextBox1.Text, ".") = 0 Then
        TextBox1.Text = TextBox1.Text & "."
    End If
End Sub

' バックスペース（1文字削除）
Private Sub NumberButton14_Click()
    If Len(TextBox1.Text) > 0 Then
        TextBox1.Text = Left(TextBox1.Text, Len(TextBox1.Text) - 1)
    End If
End Sub

' 決定（セルに反映して閉じる）
Private Sub NumberButton15_Click()

    Dim addr As String
    addr = Me.Tag

    If addr <> "" Then
        Range(addr).Value = TextBox1.Text
    End If

    Unload Me

End Sub
