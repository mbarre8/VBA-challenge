# VBA-challenge
Challenge2


Sub Stock_Challenge()

Dim tickersymbol As String
Dim lastrow As String
Dim summarytable As Long


lastrow = Cells(Rows.Count, 1).End(xlUp).Row
summarytable = 2
Range("I1").Value = "Ticker"


For i = 2 To lastrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
tickersymbol = Cells(i, 1).Value
Range("I" & summarytable).Value = tickersymbol


summarytable = summarytable + 1
End If
Next i
End Sub••••ˇˇˇˇ
