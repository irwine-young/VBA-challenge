' trialled VBA coding on alphabetical_testing excel then export VBA file to this file
Sub VBA_stock()

'Loop through all sheets
For Each ws In Worksheets

'create heading for new outputs
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

' set variable and create summary table to start at row 2
Dim ticker_summary As Integer
ticker_summary = 2

' set variable and create first yearly open value that will go through loop below.
Dim yearly_open As Double
yearly_open = ws.Cells(2, 3).Value


' set variable and find the last row from range A
Dim LastRow As Long
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
    ' Loop through rows 2 to last row for ticker symbol
    For i = 2 To LastRow
           
        ' Check if next cell has the same "ticker" cell value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' Calculation for total volume per ticker symbol
            ' Set variable for ticker; Assign values to "ticker" symbol
            Dim ticker As String
            ticker = ws.Cells(i, 1).Value
            
            ' Set variable to volume; create formula for total volume
            Dim volume As Double
            volume = volume + ws.Cells(i, 7).Value
            
            ' Print "ticker" symbol in the summary table
            ws.Range("I" & ticker_summary).Value = ticker
            
            ' Print total volume in the summary table
            ws.Range("L" & ticker_summary).Value = volume
                                      
            ' Check if ticker is the first row
            If i = 2 Then
            
            ' Calculation for Yearly Changea and % yearly change
            ' set variable for yearly_open; Assign values to "yearly_open"
            
            ' set value to open price
            yearly_open = ws.Cells(i, 3).Value
            
            End If
                                    
            ' set variable for yearly_close; Assign values to "yearly_close"
            Dim yearly_close As Double
            yearly_close = ws.Cells(i, 6).Value
                                           
            ' Check if yearly open is not ) - to prevent division by zero error
            If (yearly_open <> 0) Then
                        
            ' set variable for yearly; Assign values to "yearly" and set formula
            Dim yearly As Double
            yearly = yearly_close - yearly_open
            
            ' set variable for percent_yearly; Assign values to "percent_yearly" and set formula
            Dim percent_yearly As Double
            percent_yearly = (yearly_close - yearly_open) / yearly_open
                 
            ' Print Yearly change as "yearly" in the summary table
            ws.Range("J" & ticker_summary).Value = yearly
            
            ' Print Percent change as "percent_yearly" in the summary table
            ws.Range("K" & ticker_summary).Value = FormatPercent(percent_yearly)
                                                             
            End If
                                                                         
             ' add one to the summary table for next ticker symbol
            ticker_summary = ticker_summary + 1
            
            ' reset the volume
            volume = 0
            
            ' get first open price for next ticker
            yearly_open = ws.Cells(i + 1, 3)
                          
        ' If the cell following a row is the same ticker
        Else
            ' add to the volume
            volume = volume + ws.Cells(i, 7).Value
            
        End If
    Next i

' set variable and find the last row from range I
Dim LastRow_2 As Long
LastRow_2 = ws.Range("I" & Rows.Count).End(xlUp).Row

' set array for Yearly change column, range J
Set myArray_change = ws.Range("J" & 2 & ":" & "J" & LastRow_2)

    ' Loop through rows 2 to last row_2 for Range J
    For J = 2 To LastRow_2
    
        ' Add conditional formatting for Yearly change of >=0 and <0
        If ws.Cells(J, 10) >= 0 Then
        ws.Cells(J, 10).Interior.ColorIndex = 4
    
        Else
        ws.Cells(J, 10).Interior.ColorIndex = 3
        
        End If
    
    Next J
    
' set variable for min and max value of Percent change, max value for total stock volume
greatest_increase = WorksheetFunction.Max(ws.Range(("K" & 2 & ":" & "K" & LastRow_2)))
greatest_decrease = WorksheetFunction.Min(ws.Range(("K" & 2 & ":" & "K" & LastRow_2)))
greatest_total = WorksheetFunction.Max(ws.Range(("L" & 2 & ":" & "L" & LastRow_2)))
    
' Print min and max value of Percent change, max value for total stock volume
ws.Cells(2, 17).Value = FormatPercent(greatest_increase)
ws.Cells(3, 17).Value = FormatPercent(greatest_decrease)
ws.Cells(4, 17).Value = greatest_total
  
' set array for Ticker (range I), Percent change (range K) and total stock volume (range L)
Set myArray_ticker = ws.Range("I" & 2 & ":" & "I" & LastRow_2)
Set myArray_percent = ws.Range("K" & 2 & ":" & "K" & LastRow_2)
Set myArray_total = ws.Range("L" & 2 & ":" & "L" & LastRow_2)

' Use index to return ticker symbol with min and max value of Percent change, max value for total stock volume
ticker_increase = Application.WorksheetFunction.Index(myArray_ticker, WorksheetFunction.Match(greatest_increase, myArray_percent, 0))
ticker_decrease = Application.WorksheetFunction.Index(myArray_ticker, WorksheetFunction.Match(greatest_decrease, myArray_percent, 0))
ticker_total = Application.WorksheetFunction.Index(myArray_ticker, WorksheetFunction.Match(greatest_total, myArray_total, 0))
   
' Print ticker symbol that has min and max value of Percent change, max value for total stock volume
ws.Cells(2, 16).Value = ticker_increase
ws.Cells(3, 16).Value = ticker_decrease
ws.Cells(4, 16).Value = ticker_total

' autofit entire column and row for formatting
ws.Cells.EntireColumn.AutoFit
ws.Cells.EntireRow.AutoFit

Next ws
    
End Sub