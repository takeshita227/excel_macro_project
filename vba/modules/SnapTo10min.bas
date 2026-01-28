Attribute VB_Name = "Snapto10min"
'@ŠÔŠÛ‚ß‚é‚½‚ß‚ÌŠÖ”(10•ª’PˆÊ)
Public Function RoundToNearest10Minutes(timeStr As String) As Date
    Dim t As Date
    Dim totalMinutes As Long
    Dim roundedMinutes As Long
    
    t = CDate(timeStr)
    totalMinutes = Hour(t) * 60 + Minute(t)
    
    ' Å‚à‹ß‚¢10•ª‚ÉŠÛ‚ß‚é
    roundedMinutes = Round(totalMinutes / 10) * 10
    
    RoundToNearest10Minutes = TimeSerial(roundedMinutes \ 60, roundedMinutes Mod 60, 0)
End Function
