VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

    If Sh.Name = "ê∂éYèÛãµ" Then
        If Not Intersect(Target, Sh.Range("AD4, AE4")) Is Nothing Then
            Cancel = True
            ToggleMaru Target
        End If
    End If

End Sub

Private Sub Workbook_Open()
    LineIDCheck.Show
End Sub
