# VBA-challenge
Sub stockmarket()

'declare variables
Dim ticker As String
Dim lastrow As Double
Dim summarytable As Integer
Dim openprice As Double
Dim closedprice As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim totalstockvolume As Double
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatesttotalvolume As Double



'Inserting column and adding title

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"


summarytable = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'loop through row 2 to the last row of data
For i = 2 To lastrow

'set opening price to zero before selecting a ticker symbol
If openprice = 0 Then
'then set open price to first value in column 3 for the first row of ticker symbol selected
openprice = Cells(i, 3).Value
End If

'As you go row by row once ticker symbol changes then
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
'set ticker symbol stock name
ticker = Cells(i, 1).Value
'and Printed ticker symbol name in summary chart under Column I (titled ticker)
Range("I" & summarytable).Value = ticker

'add total stock volume for a specfic ticker
totalstockvolume = totalstockvolume + Cells(i, 7).Value
'Print ticker name in summary chart under Column L (titled Total Stock Volume)
Range("L" & summarytable).Value = totalstockvolume
'setting total stock volume back to zero to calculate total stock volume for next ticker
totalstockvolume = 0

greatesttotalvolume = Application.WorksheetFunction.Max(Range("L:L"))
Range("Q4").Value = greatesttotalvolume
If greatesttotalvolume = Range("L" & summarytable).Value Then
Range("P4").Value = Range("I" & summarytable).Value
End If

'setting row location of closing price to last row of specific ticker symbol based on if statement above
closedprice = Cells(i, 6).Value
'Calculating yearly change
yearlychange = closedprice - openprice
'Printing Yearly Change values in summary table
Range("J" & summarytable).Value = yearlychange


'if yearly change is a positive change then fill interior of that cell green
If Range("J" & summarytable).Value >= 0 Then
Range("J" & summarytable).Interior.ColorIndex = 4
'otherwise (if value is negative) fill interior of that cell red
Else
Range("J" & summarytable).Interior.ColorIndex = 3
End If

'Since you can not divide by zero if open price of current ticker is zero
If openprice = 0 Then
'then that means percent change is zero
percentchange = 0
Else
'otherwise percent change is yearly change divided by open price
percentchange = yearlychange / openprice
'printer percent change calue in summary table and format it to display a percentage
Range("K" & summarytable).Value = Format(percentchange, "Percent")
'setting opening price back to zero to calculate yearly change for next ticker
openprice = 0
End If

greatestincrease = Application.WorksheetFunction.Max(Range("K:K"))
Range("Q2").Value = Format(greatestincrease, "Percent")
If greatestincrease = Range("K" & summarytable).Value Then
Range("P2").Value = Range("I" & summarytable).Value
End If
greatestdecrease = Application.WorksheetFunction.Min(Range("K:K"))
Range("Q3").Value = Format(greatestdecrease, "Percent")
If greatestdecrease = Range("K" & summarytable).Value Then
Range("P3").Value = Range("I" & summarytable).Value
End If
'Go to next row in summary chart to enter next ticker information
summarytable = summarytable + 1

Else

'adding to the Total Stock volume
totalstockvolume = totalstockvolume + Cells(i, 7).Value


'closing out all open if statements and loops
End If

Next i

End Sub


Sub ResetButton()
 Range("I:Q").ClearContents
 Range("I:Q").ClearFormats
End Sub

