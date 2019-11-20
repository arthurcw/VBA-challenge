Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------------------------------------
'Column constants used in this procedure/module
'------------------------------------------------------------------
Private Const colTicker = 1         'column A for <ticker> field
Private Const colOpen = 3           'column C for <open>
Private Const colClose = 6          'column F for <close>
Private Const colVol = 7            'column G for <vol>

Private Const colTickerSum = 9     'column I for <ticker> in summary
Private Const colChange = 10        'column J for <Yearly Change> in summary
Private Const colPercentChange = 11 'column K for <Percent Change> in summary
Private Const colTotalVol = 12      'column L for <Total Stock Volume> in summary

Private Const colParameter = 15     'column O for paramter in greatest increase/decrease/volume
Private Const colTickerSum2 = 16    'column P for <ticker> in greatest increase/decrease/volume
Private Const colValue = 17         'column Q for <value> in greatest increase/decrease/volume

Dim i As Long                       'used in for loop


'--------------------------------------------------------------------------------------------
'Main procedure: summarize stock performance by year
'The script only works if data are already sorted in ascending order by ticker and then date
'--------------------------------------------------------------------------------------------
Sub mainStockSummary():
    Application.ScreenUpdating = False
        
    'Declare local variables
    Dim ws As Worksheet         'worksheet
    Dim dataLastRow As Long     'last row of data in worksheet
    Dim summaryLastRow As Long  'last row in summary column
    Dim stockTicker As String   'stock name
    Dim stockOpen As Double     'open price
    Dim stockClose As Double    'close price
    Dim stockVol As Double      'stock volume
    
    '------------------------
    ' Loop through each sheet
    '------------------------
    For Each ws In ActiveWorkbook.Worksheets
        'Activate sheet and clear summary
        ws.Activate
        Call clearSummary

        'Last row in data and in summary
        dataLastRow = Cells(Rows.Count, "A").End(xlUp).row
        summaryLastRow = 2
        
        'Initialize values
        stockTicker = Cells(2, colTicker).Value2
        stockVol = 0
        stockOpen = Cells(2, colOpen).Value2
        
        '----------------------------
        ' Go through each row of data
        '----------------------------
        For i = 2 To dataLastRow
            'Add trade volume
            stockVol = stockVol + Cells(i, colVol).Value2
            
            '----------------------------------------------------------
            'Summarize data if this is the last row of data for a stock
            '----------------------------------------------------------
            If stockTicker <> Cells(i + 1, colTicker).Value2 Then
                'Store close price, then add data to summary columns
                stockClose = Cells(i, colClose).Value2
                Call populateSummary(stockTicker, stockOpen, stockClose, _
                    stockVol, summaryLastRow)
                    
                'Reset variables
                stockTicker = Cells(i + 1, colTicker).Value2
                stockVol = 0
                stockOpen = Cells(i + 1, colOpen).Value2
                summaryLastRow = summaryLastRow + 1
            End If
        Next i
        
        '----------------------------------------------------------
        'Find the stock with greatest increase, decrease and volume
        '----------------------------------------------------------
        Call greatestValue(summaryLastRow - 1)
        
        'Autofit summary columns
        Range(Columns(colTickerSum), Columns(colValue)).EntireColumn.AutoFit
        
        'Done with one sheet, move onto the next
    Next
    
    'End of script message
    Application.ScreenUpdating = True
    MsgBox ("Stock data are summarized")

End Sub

'Sub procedure 1: Clear the summary columns and repopulate header
Private Sub clearSummary():
    
    'Clear summary
    Range(Columns(colTickerSum), Columns(colValue)).Clear
    
    '(Re)populate header
    Cells(1, colTickerSum).Value = "Ticker"
    Cells(1, colChange).Value = "Yearly Change"
    Cells(1, colPercentChange).Value = "Percent Change"
    Cells(1, colTotalVol).Value = "Total Stock Volume"
    Cells(1, colTickerSum2).Value = "Ticker"
    Cells(1, colValue).Value = "Value"
    Cells(2, colParameter).Value = "Greatest % Increase"
    Cells(3, colParameter).Value = "Greatest % Decrease"
    Cells(4, colParameter).Value = "Greatest Total Volume"
    
    'Autofit columns
    Range(Columns(colTickerSum), Columns(colValue)).EntireColumn.AutoFit

End Sub

'Sub procedure 2: Summarize data and populate to summary
Private Sub populateSummary(ticker As String, openPrice As Double, _
        closePrice As Double, volume As Double, row As Long):

    '(1) Ticker
    Cells(row, colTickerSum).Value2 = ticker
    
    '(2) Yearly Change & conditional formatting
    Cells(row, colChange).Value2 = closePrice - openPrice
    If Cells(row, colChange).Value2 > 0 Then
        Cells(row, colChange).Interior.ColorIndex = 4
    ElseIf Cells(row, colChange).Value2 < 0 Then
        Cells(row, colChange).Interior.ColorIndex = 3
    End If
    
    '(3) Percent change & number format
    If openPrice = 0 Then
        Cells(row, colPercentChange).Value2 = 0
    Else
        Cells(row, colPercentChange).Value2 = (closePrice - openPrice) / openPrice
    End If
    Cells(row, colPercentChange).NumberFormat = "0.00%"
    
    '(4) Total stock volume & number format
    Cells(row, colTotalVol).Value2 = volume
    Cells(row, colTotalVol).NumberFormat = "#,##0"

End Sub

'Sub procedure 3: Find greatest increase, greatest decrease and greatest total volume
Private Sub greatestValue(lastRow As Long):
    
    'Declare local variables
    Dim maxIncTicker As String          'stock with greatest percent increase
    Dim maxDecTicker As String          'stock with greatest percent decrease
    Dim maxVolTicker As String          'stock with largest total volume
    Dim maxInc As Double                'greatest percent increase
    Dim maxDec As Double                'greatest percent decrease
    Dim maxVol As Double                'largest total volume
    
    'Initialize values, set first stock in the summary as initial threshold
    maxIncTicker = Cells(2, colTickerSum).Value2
    maxDecTicker = Cells(2, colTickerSum).Value2
    maxVolTicker = Cells(2, colTickerSum).Value2
    maxInc = Cells(2, colPercentChange).Value2
    maxDec = Cells(2, colPercentChange).Value2
    maxVol = Cells(2, colTotalVol).Value2
    
    '------------------------------------------------
    ' Loop through each row/stock and compare values
    '------------------------------------------------
    For i = 3 To lastRow
        'compare percent
        If Cells(i, colPercentChange).Value2 > maxInc Then
            maxInc = Cells(i, colPercentChange).Value2
            maxIncTicker = Cells(i, colTickerSum).Value2
        ElseIf Cells(i, colPercentChange).Value2 < maxDec Then
            maxDec = Cells(i, colPercentChange).Value2
            maxDecTicker = Cells(i, colTickerSum).Value2
        End If
        
        'compare total volume
        If Cells(i, colTotalVol).Value2 > maxVol Then
            maxVol = Cells(i, colTotalVol).Value2
            maxVolTicker = Cells(i, colTickerSum).Value2
        End If
    Next i

    '-----------------
    ' Populate result
    '-----------------
    Cells(2, colTickerSum2).Value2 = maxIncTicker
    Cells(2, colValue).Value2 = maxInc
    Cells(2, colValue).NumberFormat = "0.00%"
    
    Cells(3, colTickerSum2).Value2 = maxDecTicker
    Cells(3, colValue).Value2 = maxDec
    Cells(3, colValue).NumberFormat = "0.00%"
    
    Cells(4, colTickerSum2).Value2 = maxVolTicker
    Cells(4, colValue).Value2 = maxVol
    Cells(4, colValue).NumberFormat = "#,##0"

End Sub

