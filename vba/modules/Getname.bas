Attribute VB_Name = "Getname"
Public Function GetUserName(userId As String) As String
    Dim masterFile As String
    Dim masterBook As Workbook
    Dim masterSheet As Worksheet
    Dim searchRange As Range
    Dim foundCell As Range
    Dim rowIndex As Long
    
    masterFile = "C:\Users\h_i_d\Desktop\ProductionSystem\master\excel\user_master.xlsx"
    
    If Dir(masterFile) = "" Then
        GetUserName = "ファイルなし"
        Exit Function
    End If
    
    On Error Resume Next
    Set masterBook = Workbooks.Open(masterFile, ReadOnly:=True)
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

