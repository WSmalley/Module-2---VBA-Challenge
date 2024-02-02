VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub Stock_Sort()

 Const TICKER_COL As Integer = 1
 Const OPEN_COL As Integer = 3
 Const CLOSE_COL As Integer = 6
 Const VOL_COL As Integer = 7
 Const FIRST_DATA_ROW As Integer = 2
 'Define variables
 Dim Ticker As String
 Dim Open_Value As Double
 Dim Close_Value As Double
 Dim Total_Volume As LongLong
 Dim Yearly_Change As Double
 Dim Output_Row As Long
 Dim Input_Row As Long
 Dim Change_Frac As Double
 Dim Max_Increase As Double
 Dim Max_Decrease As Double
 Dim Max_Volume As LongLong
 Dim Max_Decrease_Ticker As String
 Dim Max_Increase_Ticker As String
 Dim Max_Volume_Ticker As String
 Dim ws As Worksheet

 'Loop through each worksheet
 For Each ws In Worksheets


'print the label headings in each worksheet
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1") = "Yearly Change"
 ws.Range("K1") = "Percent Change"
 ws.Range("L1") = "Total Stock Volume"
 ws.Range("O1") = "Ticker"
 ws.Range("P1") = "Value"
 ws.Range("N2") = "Greatest % Increase"
 ws.Range("N3") = "Greatest % Decrease"
 ws.Range("N4") = "Greatest Total Volume"

 'Prepare for first Stock
 Open_Value = ws.Cells(FIRST_DATA_ROW, OPEN_COL).Value
 Yearly_Change = 0
 Output_Row = FIRST_DATA_ROW
 Max_Increase = 0
 Max_Decrease = 0
 Max_Volume = 0

 For Input_Row = FIRST_DATA_ROW To 800000
 Ticker = ws.Cells(Input_Row, TICKER_COL).Value
 Total_Volume = Total_Volume + ws.Cells(Input_Row, VOL_COL).Value

 'Last row of current stock
 If ws.Cells(Input_Row + 1, TICKER_COL).Value <> Ticker Then

 'Input
 Close_Value = ws.Cells(Input_Row, CLOSE_COL)

 'Calculations
 Yearly_Change = Close_Value - Open_Value
 Change_Frac = Yearly_Change / Open_Value
 Total_Volume = Total_Volume + ws.Cells(Input_Row, VOL_COL).Value

 'IS this a new "BEST"?
 If Change_Frac < Max_Decrease Then
 Max_Decrease = Change_Frac
 Max_Decrease_Ticker = Ticker

 ElseIf Change_Frac > Max_Increase Then
 Max_Increase = Change_Frac
 Max_Increase_Ticker = Ticker

 End If

 If Total_Volume > Max_Volume Then
 Max_Volume = Total_Volume
 Max_Volume_Ticker = Ticker
 End If

 'Output
 ws.Range("I" & Output_Row).Value = Ticker
 ws.Range("J" & Output_Row).Value = Yearly_Change
 ws.Range("K" & Output_Row).Value = FormatPercent(Change_Frac)
 ws.Range("L" & Output_Row).Value = Total_Volume

 'Color negative changes as red and positive and zero changes as green
 If Yearly_Change < 0 Then
 ws.Range("J" & Output_Row).Interior.ColorIndex = 3
 'Cells(Output_Row, 10).Interior.ColorIndex = 3
 ElseIf Yearly_Change > 0 Then
 ws.Range("J" & Output_Row).Interior.ColorIndex = 4
 'Cells(Output_Row, 10).InteriorColorIndex = 4
 End If

 'Prepare for Next Stock
 Open_Value = ws.Cells(Input_Row + 1, OPEN_COL).Value

 Output_Row = Output_Row + 1
 Yearly_Change = 0
 Total_Volume = 0

 End If

 Next Input_Row

 ws.Range("O2").Value = Max_Increase_Ticker
 ws.Range("P2").Value = FormatPercent(Max_Increase)
 ws.Range("O3").Value = Max_Decrease_Ticker
 ws.Range("P3").Value = FormatPercent(Max_Decrease)
 ws.Range("O4").Value = Max_Volume_Ticker
 ws.Range("P4").Value = Max_Volume

 'LastRow = ws.Cells(Input_Row, TICKER_COL).End(x1Up).Row
 'TUTOR: last_row = Cells(1, TICKER_COL).End(x1Down)Row
 Next ws
 MsgBox "done"
End Sub
