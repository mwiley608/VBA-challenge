{\rtf1\ansi\ansicpg1252\cocoartf2709
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub AnnualStockSummary()\
\
'loop through all sheets\
Dim ws As Worksheet\
For Each ws In Worksheets\
    ws.Activate\
\
'variables for ticker, open, close, yearly change, percent change, total volume\
    Dim Ticker_Symbol As String\
    Dim open_amt As Single\
        open_amt = Cells(2, 3).Value\
    Dim close_amt As Single\
    Dim yearly_change As Single\
    Dim percent_change As Single\
    Dim total_volume As Double\
        total_volume = 0\
\
'track location of each ticker symbol\
    Dim summary_table_row As Integer\
    summary_table_row = 2\
\
'Determine last row\
    Dim LastRow As Long\
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row\
'MsgBox "There are " & LastRow & " Rows"\
\
'Add headers\
Cells(1, 9).Value = "Ticker"\
Cells(1, 10).Value = "Year Open"\
Cells(1, 11).Value = "Year Close"\
Cells(1, 12).Value = "Yearly Change"\
Cells(1, 13).Value = "Percent Change"\
Cells(1, 14).Value = "Total Stock Volume"\
Cells(1, 17).Value = "Ticker"\
Cells(1, 18).Value = "Value"\
Cells(2, 16).Value = "Greatest % Increase"\
Cells(3, 16).Value = "Greatest % Decrease"\
Cells(4, 16).Value = "Greatest Total Volume"\
\
\
'Loop through ticker symbols\
For I = 2 To LastRow\
\
'check if ticker symbol is the same, if not\
If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then\
\
'set ticker Symbol\
Ticker_Symbol = Cells(I, 1).Value\
        \
'set open and close value\
close_amt = Cells(I, 6).Value\
\
Range("J" & summary_table_row).Value = open_amt\
\
Range("K" & summary_table_row).Value = close_amt\
\
'calculate yearly change\
yearly_change = close_amt - open_amt\
\
'calculate percent change\
percent_change = yearly_change / open_amt\
\
'add to volume total\
total_volume = total_volume + Cells(I, 7).Value\
\
'print ticker symbol in summary table\
Range("I" & summary_table_row).Value = Ticker_Symbol\
\
'print yearly change in summary table\
Range("L" & summary_table_row).Value = yearly_change\
\
'determine last row of summary table\
Dim ST_LastRow As Long\
ST_LastRow = Cells(Rows.Count, 9).End(xlUp).Row\
\
'set cell colors for yearly_change\
For l = 2 To ST_LastRow\
    If Cells(l, 12) > 0 Then\
    Cells(l, 12).Interior.ColorIndex = 4\
    Else\
    Cells(l, 12).Interior.ColorIndex = 3\
    End If\
    Next l\
\
'print percent change in summary table\
Range("M" & summary_table_row).Value = FormatPercent(percent_change)\
\
'print total stock volume in summary table\
Range("N" & summary_table_row).Value = total_volume\
\
'add one to summary table row\
summary_table_row = summary_table_row + 1\
\
'reset volume total\
total_volume = 0\
\
'reset open value\
open_amt = Cells(I + 1, 3).Value\
\
'if next cell is same symbol\
Else\
\
'save new close value\
close_amt = Cells(I + 1, 6).Value\
Range("K" & summary_table_row).Value = close_amt\
\
'add to volume total\
total_volume = total_volume + Cells(I, 7).Value\
\
End If\
Next I\
\
'create variables for max values\
Dim max_incr As Single\
Dim max_incr_tick As String\
Dim max_decr As Single\
Dim max_decr_tick As String\
Dim max_vol As Single\
Dim max_vol_tick As String\
max_incr = 0\
max_decr = 0\
max_vol = 0\
\
\
'find max values\
For m = 2 To ST_LastRow\
    If Cells(m, 13).Value > max_incr Then\
    max_incr = Cells(m, 13).Value\
    max_incr_tick = Cells(m, 9).Value\
    \
    End If\
\
    If Cells(m, 13) < max_decr Then\
    max_decr = Cells(m, 13).Value\
    max_decr_tick = Cells(m, 9).Value\
    \
    End If\
    \
    \
    If Cells(m, 14) > max_vol Then\
    max_vol = Cells(m, 14).Value\
    max_vol_tick = Cells(m, 9).Value\
    \
    End If\
    Next m\
\
\
'print max values\
Cells(2, 17).Value = max_incr_tick\
Cells(2, 18).Value = FormatPercent(max_incr)\
Cells(3, 17).Value = max_decr_tick\
Cells(3, 18).Value = FormatPercent(max_decr)\
Cells(4, 17).Value = max_vol_tick\
Cells(4, 18).Value = max_vol\
\
Next ws\
\
End Sub\
\
}