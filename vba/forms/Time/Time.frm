VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Time 
   Caption         =   "時刻変更"
   ClientHeight    =   3795
   ClientLeft      =   132
   ClientTop       =   552
   ClientWidth     =   5736
   OleObjectBlob   =   "Time.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "TIME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    ListBox1.List = Array(0, 1, 2)
    ListBox2.List = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    ListBox3.List = Array(0, 1, 2, 3, 4, 5)
    ListBox4.List = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
End Sub
