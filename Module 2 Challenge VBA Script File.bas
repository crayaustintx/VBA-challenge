Attribute VB_Name = "Module1"
'Run SubRoutine titled 'RunMacro' to complete all tasks in the workbook
Sub RunMacro()
Call Stocks
Call Summary_Table
End Sub


'For each worksheet, SubRoutine titled Stocks creates a summary table in columns I:L
Sub Stocks()


Dim WS As Worksheet
For Each WS In Worksheets


Dim Ticker As String
Dim Q_Change_Start As Double
Dim Q_Change_Qty As Double
Dim Percent_Change As Double
Dim TSV As Double
Dim i As Long
Dim Summary_Table As Integer

WS.Cells(1, 9) = "Ticker"
WS.Cells(1, 10) = "Quarterly Change"
WS.Cells(1, 11) = "Percent Change"
WS.Cells(1, 12) = "Total Stock Volume"
WS.Range("J:J").NumberFormat = "#,##0.00"
WS.Range("K:K").NumberFormat = "0.0%"
WS.Range("L:L").NumberFormat = "#,##0"



'Set initial values
'Ticker = WS.Cells(2, 1).Value
Q_Change_Start = WS.Cells(2, 3).Value
TSV = 0
Summary_Table = 1

'Determine the Last Row
lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    'If not the last record of ticker
    If WS.Cells(i, 1).Value = WS.Cells(i + 1, 1).Value Then
        'Cumulate stock volume
        TSV = TSV + WS.Cells(i, 7).Value
    'If last record of ticker
    Else
        'Add ticker to list
        Summary_Table = Summary_Table + 1
        WS.Cells(Summary_Table, 9).Value = WS.Cells(i, 1).Value
        'get quarterly change and write to table
        Q_Change_Qty = WS.Cells(i, 6).Value - Q_Change_Start
        WS.Cells(Summary_Table, 10).Value = Q_Change_Qty
        'get percent change and write to table
        Percent_Change = Q_Change_Qty / Q_Change_Start
        WS.Cells(Summary_Table, 11).Value = Percent_Change
        'get write total volume to table
        TSV = TSV + WS.Cells(i, 7).Value
        WS.Cells(Summary_Table, 12).Value = TSV
    
        'Reset Total_Stock_Volume
        TSV = 0
        Ticker = WS.Cells(i + 1, 1).Value
        Q_Change_Start = WS.Cells(i + 1, 3).Value
        
    End If

Next i

WS.Columns("I:L").AutoFit

'Next WS
Next


'End Sub Stocks
End Sub

'For each worksheet, SubRoutine titled Summary_Table adds color formatting to the Summary Table in Col I:L and creates an additional summary table in Col N:P
Sub Summary_Table()

Dim WS As Worksheet
For Each WS In Worksheets

Dim i As Integer
Dim Max_Pct_Tracker As Double
Dim Min_Pct_Tracker As Double
Dim TSV As Double

'Initiate variables to first record
Max_Pct_Tracker = WS.Cells(2, 11)
Min_Pct_Tracker = WS.Cells(2, 11)
Max_TSV = WS.Cells(2, 12)

'Create Summary of Summary table
WS.Cells(2, 14) = "Greatest % Increase"
WS.Cells(3, 14) = "Greatest % Decrease"
WS.Cells(4, 14) = "Greatest Total Volume"
WS.Cells(1, 15) = "Ticker"
WS.Cells(1, 16) = "Value"
WS.Cells(2, 16).NumberFormat = "0.0%"
WS.Cells(3, 16).NumberFormat = "0.0%"
WS.Cells(4, 16).NumberFormat = "#,##0"


'Determine the Last Row
lastrow = WS.Cells(Rows.Count, 9).End(xlUp).Row


'Color Summary Table and evauate summary table for max/min records
For i = 2 To lastrow
    
    'Color Quarterly Change Red/Green
    If WS.Cells(i, 10) < 0 Then
    WS.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
    ElseIf WS.Cells(i, 10) > 0 Then
    WS.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
    End If
    
    'Find max value of Percent Change
    If WS.Cells(i, 11).Value > Max_Pct_Tracker Then
    Max_Pct_Tracker = WS.Cells(i, 11).Value
    End If
    
    'Find min value of Percent Change
    If WS.Cells(i, 11).Value < Min_Pct_Tracker Then
    Min_Pct_Tracker = WS.Cells(i, 11).Value
    End If
    
    'Find max value of Total Stock Volume
    If WS.Cells(i, 12).Value > Max_TSV Then
    Max_TSV = WS.Cells(i, 12).Value
    End If

Next i

'Write Max/Min values to Summary of Summary table
WS.Cells(2, 16).Value = Max_Pct_Tracker
WS.Cells(3, 16).Value = Min_Pct_Tracker
WS.Cells(4, 16).Value = Max_TSV


'Get ticker associated with max values
For i = 2 To lastrow
    'write max and min tickers to summary of summary table
    If WS.Cells(i, 11).Value = Max_Pct_Tracker Then
    WS.Cells(2, 15) = WS.Cells(i, 9).Value
    End If
    
    'write min tickers to summary of summary table
    If WS.Cells(i, 11).Value = Min_Pct_Tracker Then
    WS.Cells(3, 15) = WS.Cells(i, 9).Value
    End If
    
   'write total stock volume ticker to summary of summary table
    If WS.Cells(i, 12).Value = Max_TSV Then
    WS.Cells(4, 15) = WS.Cells(i, 9).Value
    End If
   
Next i


WS.Columns("N:P").AutoFit

'Next WS
Next


'End Sub Summary_Table()
End Sub





