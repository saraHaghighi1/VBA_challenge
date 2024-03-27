## VBA_challenge
#VBA coding
**I tried to write  my first code with VBA and i had a challenge but i got through it**

Attribute VB_Name = "Module1"

Sub FillRowForTicker()
**1.complete arow for a ticker**
  Dim Ticker_counts As Integer
  GetUniqueTickers

Ticker_counts = Cells(Rows.Count, "I").End(xlUp).Row


For i = 2 To Ticker_counts
Cells(i, 10).Value = CalculateYearlyChange(Cells(i, 9).Value)
Cells(i, 11).Value = CalculatePercentageChange(Cells(i, 9).Value)
Cells(i, 11).NumberFormat = "0.00%"
Cells(i, 12).Value = CalculateTotalStock(Cells(i, 9).Value)
Next i

**2.ApplyConditionalFormatting**

 Range("J1").Value = "Yearly Change"
 Range("k1").Value = "Percentage change"
 Range("L1").Value = "Total Stock Volume"
 Range("Q3").Value = "Greatest % Increase"
 Range("Q4").Value = "Greatest % Decrease"
 Range("Q5").Value = "Total Stock Volume"
 Range("R2").Value = "Tiker"
 Range("S2").Value = "Value"
-**Greatest % Increase**
  Range("I:L").Sort key1:=Range("K1"), Order1:=xlDescending, Header:=xlYes
  Range("S3").Value = Range("K2").Value
  Range("R3").Value = Range("I2").Value
  Range("S3").NumberFormat = "0.00%"

-**Greatest % Decrease**

Range("I:L").Sort key1:=Range("K1"), Order1:=xlAscending, Header:=xlYes
Range("S4").Value = Range("K2").Value
Range("R4").Value = Range("I2").Value
Range("S4").NumberFormat = "0.00%"

*Greatest Total Volume*

Range("I:L").Sort key1:=Range("L1"), Order1:=xlDescending, Header:=xlYes
Range("S5").Value = Range("L2").Value
Range("R5").Value = Range("I2").Value
End Sub
--------------------------------------------------------------------------------
Sub debuging()
'for debuging other functions
 MsgBox (Cells(Rows.Count, "A").End(xlUp).Row)


End Sub
------------------------------------------------------------------------
Sub GetUniqueTickers()

Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True

End Sub
-----------------------------------------------------------------------------
**Do you have any other sugestion for this code?**
Function CalculateTotalStock(ticker As String) As LongLong
'calculate Total Stock Volume for ticker
Dim TotalStock As LongLong
TotalStock = 0
For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row
If Cells(i, 1).Value = ticker Then

TotalStock = TotalStock + Cells(i, 7).Value

End If
Next i
CalculateTotalStock = TotalStock

End Function
--------------------------------------------------------------------------------

Function CalculateYearlyChange(ticker As String) As Double

Dim lastRow As Long
Dim openPrice As Double
Dim closePrice As Double

'let s find open openPrice
lastRow = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To lastRow
If Cells(i, 1).Value = ticker Then
' we found the first row of ticker
openPrice = Cells(i, 3).Value

Exit For
End If
Next i
---------------------------------------------------------------------------------------
'let s find closePrice
For i = 2 To lastRow
If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
' we found the last row of ticker
closePrice = Cells(i, 6).Value

Exit For
End If
Next i
CalculateYearlyChange = openPrice - closePrice
End Function
------------------------------------------------------------------------------------------


Function CalculatePercentageChange(ticker As String) As Double

Dim lastRow As Long
Dim openPrice As Double
Dim closePrice As Double

'let s find open openPrice
lastRow = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To lastRow
If Cells(i, 1).Value = ticker Then
' we found the first row of ticker
openPrice = Cells(i, 3).Value

Exit For
End If
Next i
---------------------------------------------------------------------------------------
'let s find closePrice
For i = 2 To lastRow
If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
' we found the last row of ticker
closePrice = Cells(i, 6).Value

Exit For
End If
