Sub MarketChallenge()

'Declare variables to name column headers

Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Long


'Name additional columns
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"


'Declare variables

Dim Ticker_Names_Found As Long
Dim LastRow As Long
Dim TickerOpen As Double
Dim LastTickerRow As Long
Dim TickerClose As Double
Dim Current_Row As Long
Dim TickerChange As Double
Dim TickerPercent As Double
Dim First_Ticker_Row As Long

Application.ScreenUpdating = False

 'Identify last row in worksheet
 
Last_Row = Range("A1").End(xlDown).Row
Ticker_Names_Found = 0

For Current_Row = 2 To Last_Row
Cells(Current_Row, 1).Select

'move ticker row / find ticker symbols

If ActiveCell.Value <> ActiveCell.Offset(-1, 0).Value Then


Ticker_Names_Found = Ticker_Names_Found + 1


First_Ticker_Row = ActiveCell.Row
Range("I1").Offset(Ticker_Names_Found, 0).Value = ActiveCell.Value
TickerOpen = ActiveCell.Offset(0, 2).Value

End If

If ActiveCell.Value <> ActiveCell.Offset(1, 0).Value Then

'Create loops/formulas to solve for columns J-L

LastTickerRow = ActiveCell.Row
TickerClose = ActiveCell.Offset(0, 5).Value
TickerChange = TickerClose - TickerOpen

If TickerOpen = 0 Then

TickerPercent = 0

Else

TickerPercent = TickerChange / TickerOpen
 
End If

Range("J1").Offset(Ticker_Names_Found, 0).Value = TickerChange
Range("K1").Offset(Ticker_Names_Found, 0).Value = TickerPercent

Set MyRange = Range(Cells(First_Ticker_Row, 7), Cells(LastTickerRow, 7))

Range("L1").Offset(Ticker_Names_Found, 0).Value = Application.WorksheetFunction.Sum(MyRange)

End If


Next Current_Row
Range("A1").Select

Application.ScreenUpdating = True

End Sub




