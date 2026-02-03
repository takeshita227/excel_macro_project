Attribute VB_Name = "TheClock"
' ▼ TextBox を時計専用に設定（編集不可＋現在時刻セット）
Public Sub SetupClockTextBox(tb As MSForms.TextBox)

    tb.Value = Format(Now, "hh:mm")

    ' 編集不可（カーソルも入らない）
    tb.Enabled = False

    ' 見た目は白のまま
    tb.BackColor = vbWhite

End Sub


' ▼ TextBox が空になったときに復元
Public Sub RestoreClockIfEmpty(tb As MSForms.TextBox)

    If tb.Value = "" Then
        tb.Value = Format(Now, "hh:mm")
    End If

End Sub
