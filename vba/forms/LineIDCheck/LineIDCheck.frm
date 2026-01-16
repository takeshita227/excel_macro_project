VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LineIDCheck 
   Caption         =   "UserForm1"
   ClientHeight    =   16140
   ClientLeft      =   84
   ClientTop       =   372
   ClientWidth     =   2808
   OleObjectBlob   =   "LineIDCheck.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "LineIDCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ListBox1_Click()
    ThisWorkbook.Sheets("生産状況").Range("C4").Value = Me.ListBox1.Value
    Unload Me
End Sub


Private Sub UserForm_Initialize()

    ' B-1 ～ B-22
    Dim i As Long
    For i = 1 To 22
        Me.ListBox1.AddItem "B-" & i
    Next i

    ' WP 系
    Me.ListBox1.AddItem "WP-1"
    Me.ListBox1.AddItem "WP-3"
    Me.ListBox1.AddItem "WP-4"
    Me.ListBox1.AddItem "WP-5"

    ' D 系
    Me.ListBox1.AddItem "D-20"
    Me.ListBox1.AddItem "D-24"
    Me.ListBox1.AddItem "D-25"

End Sub
