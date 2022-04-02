Attribute VB_Name = "Module6"
Sub wsloop()

'Create loop for worksheets
For Each ws In Worksheets

'Create column names
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Find the last non-blank cell in a single column
Dim lastrow As Long
'Find the last non-blank cell in column A(1)
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'-------------------------------------------
'Loop for ticker and Volume
'-------------------------------------------

'create variables and assign values
Dim tickervar As String
Dim volume As LongLong
Dim summary_table_row As Integer

volume = 0
summary_table_row = 2

    'Create loop for ticker and volume total
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'set ticker variable to current ticker string
        tickervar = ws.Cells(i, 1).Value
        
        'Add to the volume total when being the first new ticker
        volume = volume + ws.Cells(i, 7).Value
        
        'Print ticker into summary table
        ws.Range("I" & summary_table_row).Value = tickervar
        
        'Print volume into summary table
        ws.Range("L" & summary_table_row).Value = volume
        
        'Add one to the summary_table_row
        summary_table_row = summary_table_row + 1
        
        'Reset volume for next ticker
        volume = 0
        
        Else
        
        'Add to volume when sticker is the same
        volume = volume + ws.Cells(i, 7).Value

        End If
        
    Next i
      
'---------------------------------------------
'Yearly Change AND Percent change loop
'---------------------------------------------
    'Create variables for yearly change loop
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim opening As Double
    Dim closing As Double
    Dim summary_table_row2
    
    'set variables equal to 0
    yearly_change = 0
    percent_change = 0
    opening = 0
    closing = 0
    summary_table_row2 = 2
    
    'For this loop I'm using lastRow and summary_table_row variable from above
    
    'Create loop for yearly change and percent change
    For i = 2 To lastrow
        
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
          'set closing value to cell value
          closing = closing + ws.Cells(i, 6).Value
          
          'calculate yearly_change
          yearly_change = closing - opening
          
          'calculate percent_change
          percent_change = (yearly_change / opening) * 100
          
          'print yearly_change in summary table
          ws.Range("J" & summary_table_row2).Value = yearly_change
          
          'print percent change in summary table
          ws.Range("K" & summary_table_row2).Value = percent_change
           
          'add one to the summary table row
          summary_table_row2 = summary_table_row2 + 1
          
          'Reset variables
          closing = 0
          opening = 0
          yearly_change = 0
          percent_change = 0
          
          ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
          opening = ws.Cells(i, 3).Value
        
          End If
        
    Next i

'----------------------------------------
'Conditional Formatting Loop
'----------------------------------------
'Create last row for summary table
'Find the last non-blank cell in a single column
Dim lastRow2 As Long
'Find the last non-blank cell in column A(1)
lastRow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row

'Create conditional formatting loop
    For i = 2 To lastRow2
        If ws.Cells(i, 10).Value < 0 Then
        'Negative numbers are red
        ws.Cells(i, 10).Interior.ColorIndex = 3
     
        ws.Cells(i, 11).Interior.ColorIndex = 3
     
        ElseIf ws.Cells(i, 10).Value > 0 Then
        
        'Positive numbers are green
        ws.Cells(i, 10).Interior.ColorIndex = 4
     
        ws.Cells(i, 11).Interior.ColorIndex = 4
        
        Else
        ws.Cells(i, 10).Interior.ColorIndex = 2
     
        ws.Cells(i, 11).Interior.ColorIndex = 2
        
        End If
     
    Next i
    
Next ws
    
End Sub
