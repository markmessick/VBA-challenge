Attribute VB_Name = "Module1"
Sub HomeWork()

' Create the headings for my summary table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' Determine the Last Row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Create the variables I'll be using
Dim Ticker As String
Dim Volume As Double
Dim Opening As Double
Dim Closing As Double
Dim Yearly As Double

' Assign my variables
Yearly = 0
Opening = Cells(2, 3).Value
Volume = 0

' Create a place for my summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Run a For Loop that looks for changes in the ticker symbol
' and uses the assigned variables to calculate the yearly change,
' percentage change, and total volume for each ticker
For i = 2 To LastRow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Closing = Cells(i, 6).Value
        Ticker = Cells(i, 1).Value
        Volume = Volume + Cells(i, 7).Value
        Yearly = Closing - Opening
        
        ' Add the values into the summary table

        Range("I" & Summary_Table_Row).Value = Ticker
        Range("J" & Summary_Table_Row).Value = Closing - Opening
        Range("K" & Summary_Table_Row).Value = (Yearly / Opening) * 100
        Range("L" & Summary_Table_Row).Value = Volume
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Reset the values and begin looking for the new ticker change
        
        Opening = Cells(i + 1, 3).Value
        Yearly = Cells(2 + 1, 10).Value
        Volume = 0
      
        Else

        Volume = Volume + Cells(i, 7).Value

        End If
    
Next i

LastSummary = Cells(Rows.Count, 10).End(xlUp).Row

For j = 2 To LastSummary
        If Cells(j, 10).Value > 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
        Else
        Cells(j, 10).Interior.ColorIndex = 3
        End If

        Cells(j, 11).NumberFormat = "0.00"

Next j
    
End Sub
