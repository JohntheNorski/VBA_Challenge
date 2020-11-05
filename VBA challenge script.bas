Attribute VB_Name = "Module1"
Sub Stock_Loop():

'Setting variables that I will need
Dim Volume, OpenPrice, EndPrice, Change, highestVolume, highestChange, lowestChange As Double
Dim ticker, highestVolumeTicker, highestChangeTicker, lowestChangeTicker, worksheetvalue As String
Dim wb As Workbook
Dim ws As Worksheet
Dim FinalRow As Long
Dim x As Integer

    'Gathering user input to know which year to parse
    worksheetvalue = InputBox("Enter which year you would like to parse (2014, 2015, 2016): ", "Year selector", 2014)
    While worksheetvalue <> 2014 And worksheetvalue <> 2015 And worksheetvalue <> 2016
        worksheetvalue = InputBox("Enter in a correct year  (2014, 2015, 2016): ", "Year selector", 2014)
    Wend
    
    'Initialization including no screenupdating, variable setting, formatting and headers, finding Finalrow
    Application.ScreenUpdating = False
    highestVolume = 0
    highestChange = 0
    lowestChange = 0
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(worksheetvalue)
    ws.Activate
    ws.Range("H:Z").Clear
    FinalRow = ws.Range("A1").End(xlDown).Row
    x = 2
    OpenPrice = ws.Cells(2, 3).Value
    ws.Range("J:J").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 9).Font.Bold = True
    ws.Cells(1, 10).Value = "% price change"
    ws.Cells(1, 10).Font.Bold = True
    ws.Cells(1, 11).Value = "Price change %"
    ws.Cells(1, 11).Font.Bold = True
    ws.Cells(1, 12).Value = "Total Volume"
    ws.Cells(1, 12).Font.Bold = True
    
    'Beginning loop through all rows
    For i = 2 To FinalRow
    
        'setting ticker and summing volume
        ticker = ws.Cells(i, 1).Value
        Volume = Volume + ws.Cells(i, 7).Value
        
        'if the next ticker isn't the current ticker, need to paste current results and reset
        If ws.Cells(i + 1, 1).Value <> ticker Then
        
            'this was just in case the endvalue was corrupted, it would simply take the last open value instead so we have something
            If ws.Cells(i, 6).Value = 0 Then
                EndPrice = ws.Cells(i, 3).Value
            Else
                EndPrice = ws.Cells(i, 6).Value
            End If
            
            'Checking open price, as it's our comparison and if 0 we have an error, otherwise get the change percentage
            If OpenPrice = 0 Then
                Change = 0
            Else
                Change = (EndPrice - OpenPrice) / OpenPrice
            End If
            
            'Insert all of our values now to the summary table
            ws.Cells(x, 9).Value = ticker
            ws.Cells(x, 10).Value = EndPrice - OpenPrice
            
            'quick check if percent is negative or positive, then respond with green or red fill
            If ws.Cells(x, 10).Value >= 0 Then
                ws.Cells(x, 10).Interior.Color = vbGreen
            Else
                ws.Cells(x, 10).Interior.Color = vbRed
            End If
            ws.Cells(x, 11).Value = Change
            ws.Cells(x, 12).Value = Volume
            
            'adding 1 to our x variable, as we use this to keep going down a row for every new entry
            x = x + 1
            
            'Bonus part, checking current values against global standing values for a new winner
            If Volume > highestVolume Then
                highestVolume = Volume
                highestVolumeTicker = ticker
            End If
            If Change > highestChange Then
                highestChange = Change
                highestChangeTicker = ticker
            End If
            If Change < lowestChange Then
                lowestChange = Change
                lowestChangeTicker = ticker
            End If
            
            'resetting OpenPrice to next row's value, resetting volume to 0
            OpenPrice = ws.Cells(i + 1, 3).Value
            Volume = 0
        End If
    Next i
    
    'Bonus summary table entry, formatting, and headers
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 16).Font.Bold = True
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(1, 17).Font.Bold = True
    ws.Cells(2, 15).Value = "Highest Volume: "
    ws.Cells(2, 15).Font.Bold = True
    ws.Cells(2, 16).Value = highestVolumeTicker
    ws.Cells(2, 17).Value = highestVolume
    ws.Cells(3, 15).Value = "Greatest % Increase: "
    ws.Cells(3, 15).Font.Bold = True
    ws.Cells(3, 16).Value = highestChangeTicker
    ws.Cells(3, 17).Value = highestChange
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 15).Value = "Lowest % Decrease: "
    ws.Cells(4, 15).Font.Bold = True
    ws.Cells(4, 16).Value = lowestChangeTicker
    ws.Cells(4, 17).Value = lowestChange
    ws.Cells(4, 17).NumberFormat = "0.00%"



End Sub
