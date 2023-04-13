VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks()
'Ethan Clark 4/13/2023 VBA challenge
Dim ws As worksheet

For Each ws In ActiveWorkbook.Worksheets
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Tot Vol"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Tot Vol"

'Preallocations

Dim Tick As String: Tick = " "
Dim Ticker_vol As Double: Ticker_vol = 0
Dim tick_row As Integer: tick_row = 2
Dim last As Long
Dim i As Long
Dim j As Integer

'Allocate search vars
    Dim BiggestNegChange As Double
    Dim BiggestPosChange As Double
    Dim BiggestVol As Double
    Dim VolTick As String
    Dim PosTick As String
    Dim NegTick As String
    
    BiggestVol = 0
    BiggestNegChange = 0
    BiggestPosChange = 0
'last row calc
last = ws.Cells(Rows.Count, 1).End(xlUp).Row

'More Preallocations
Dim OpenPrice As Double: OpenPrice = 0
Dim ClosePrice As Double: ClosePrice = 0
Dim PriceChange As Double: PriceChange = 0
Dim PriceChangeP As Double: PriceChangeP = 0

'Put the initial price down
OpenPrice = ws.Cells(2, 3).Value

'start the loop across tick values
For i = 2 To last

'Volume
Ticker_vol = Ticker_vol + ws.Cells(i, 7).Value

'make da ticker
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Tick = ws.Cells(i, 1).Value
    
    
    'print da ticker
    ws.Cells(tick_row, 9).Value = Tick
    
    'print da volume
    ws.Cells(tick_row, 12) = Ticker_vol


    'Calc PriceChange

    'Pull the closing price
    ClosePrice = ws.Cells(i, 6).Value
    
    'Subtract the open price from the closing
    PriceChange = ClosePrice - OpenPrice
    
            'Catch zero open prices
            If OpenPrice = 0 Then
            PriceChangeP = 0
            Else
            PriceChangeP = PriceChange / OpenPrice
            End If
    
    'print da change
    ws.Cells(tick_row, 10).Value = PriceChange
    
    'print da % change
    ws.Cells(tick_row, 11).Value = PriceChangeP
    ws.Cells(tick_row, 11).NumberFormat = "0.00%"

    'progress where the tick info is printed
    tick_row = tick_row + 1
    
    'Compare this ticks volume to the running max volume

            'Allocate the running greatest volume
            If Ticker_vol > BiggestVol Then
                BiggestVol = Ticker_vol
                VolTick = Tick
            End If
            
            'Allocate the running greatest pos/neg % change
            If PriceChangeP > 0 Then
                    If PriceChangeP > BiggestPosChange Then
                    BiggestPosChange = PriceChangeP
                    PosTick = Tick
                    End If
                Else
                    If PriceChangeP < BiggestNegChange Then
                    BiggestNegChange = PriceChangeP
                    NegTick = Tick
                    End If
            
            
        End If
            
            
    'Reset volume for the ticker
    Ticker_vol = 0
    'Pull the next Open Price
    OpenPrice = Cells(i + 1, 3).Value
    'Reset the Price Change Metrics
    PriceChange = 0
    PriceChangeP = 0
    
End If
 
Next i

'Establish what the last row in the calculated table is
lastSum = ws.Cells(Rows.Count, 9).End(xlUp).Row
'Loop to format the Changes
For j = 2 To lastSum
    If ws.Cells(j, 10).Value > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(j, 10).Interior.ColorIndex = 3
    End If
Next j
    'Printing da rightmost table
        'Biggest Positive change
        ws.Cells(2, 16).Value = PosTick
        ws.Cells(2, 17).Value = BiggestPosChange
        ws.Cells(2, 17).NumberFormat = "0.00%"
        'Biggest Neg Change
        ws.Cells(3, 16).Value = NegTick
        ws.Cells(3, 17).Value = BiggestNegChange
        ws.Cells(3, 17).NumberFormat = "0.00%"
        'Biggest Vol
        ws.Cells(4, 16).Value = VolTick
        ws.Cells(4, 17).Value = BiggestVol
        
 
    
Next ws
'Resetting Tick Row and Global Metrics
tick_row = 2
BiggestVol = 0
BiggestNegChange = 0
BiggestPosChange = 0

'It is so hard to not do comments with %, MATLAB has ruined me
End Sub
