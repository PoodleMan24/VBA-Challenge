Attribute VB_Name = "Module1"
Sub Stock()
'defining variables
Dim lastrow As Long
Dim Ticker As String
Dim AnswerRow As Integer
Dim TotalVol As Double
Dim ClosePrice As Double
Dim OpenPrice As Double
Dim YearChange As Double
Dim Percentage As Double

'Naming all cells that require it (Summary Table)
Range("I1, P1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

'Setting Values for predefined variables outside of loop
TotalVol = 0
AnswerRow = 2
OpenPrice = Cells(2, 3).Value
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Creating a for-loop to run through each line
For i = 2 To lastrow
    Ticker = Cells(i, 1).Value
    ClosePrice = Cells(i, 6).Value
    If Ticker <> Cells(i + 1, 1).Value Then
        Cells(AnswerRow, 9).Value = Ticker
        
        'This will print the total value from the else statements
        TotalVol = TotalVol + Cells(i, 7).Value
        Cells(AnswerRow, 12).Value = TotalVol
        
        'Now we calculate the percentage comparison by dividing the
        'Closing Price with the Opening Price and formatting the cell
        Percentage = ((ClosePrice / OpenPrice)) - 1
        Cells(AnswerRow, 11).NumberFormat = "0.00%"
        Cells(AnswerRow, 11).Value = Percentage
        
        'We can calculate the change value of the year with following formula
        YearChange = ClosePrice - OpenPrice
        
        'and then we can format the cells to change colour depending
        'on whether the value is negative, positive or 0
            If YearChange > 0 Then
                    Cells(AnswerRow, 10).Interior.ColorIndex = 4
                ElseIf YearChange < 0 Then
                    Cells(AnswerRow, 10).Interior.ColorIndex = 3
                    Cells(AnswerRow, 10).Font.ColorIndex = 2
                Else: Cells(AnswerRow, 10).Interior.ColorIndex = 6
            
            End If
        Cells(AnswerRow, 10).Value = YearChange

        
        
        'Setting next Tickers starting price value
        OpenPrice = Cells((i + 1), 3).Value
        TotalVol = 0
        AnswerRow = AnswerRow + 1
    
    
    
    Else: TotalVol = TotalVol + Cells(i, 7).Value
    
    
    End If

Next i
'End of ForLoop, start of Bonus Section table

'Defining new Variables, could be done at top of sheet but I thought
'It'd be a good idea to keep it here with the values
Dim LastSummary As Double
Dim MaxValue As Double
Dim MinValue As Double
Dim MaxBrand As String
Dim MinBrand As String
Dim Vol As Double
Dim VolBra As String

'Assigning Values to Variables that need them outside the loop
LastSummary = Cells(Rows.Count, 9).End(xlUp).Row
MinValue = Cells(2, 11).Value
MaxValue = 0
Vol = 0

'For loop to scan our summary table from previous loop
For i = 2 To LastSummary
'If the value is greater than previous value, starting at 0, then assign
'current cell as new maximum value and continue to next row
    If Cells(i, 11) > MaxValue Then
        MaxValue = Cells(i, 11)
        MaxBrand = Cells(i, 9)
    End If
'New If statement to calculate Min Value, same as before except
'starting with the first cell as initial value, outside of loop
    If Cells(i, 11) < MinValue Then
        MinValue = Cells(i, 11)
        MinBrand = Cells(i, 9)
    End If
'Last new if statement to check the final Volume in each row of Summ Table
    If Cells(i, 12) > Vol Then
        Vol = Cells(i, 12)
        VolBra = Cells(i, 9)
    End If
'End of If statement and Closure of ForLoop
Next i
'Inserting final Values into a table and changing to % value for max and min
Range("Q2:Q3").NumberFormat = "0.00%"
Cells(2, 17).Value = MaxValue
Cells(2, 16).Value = MaxBrand
Cells(3, 17).Value = MinValue
Cells(3, 16).Value = MinBrand
Cells(4, 17).Value = Vol
Cells(4, 16).Value = VolBra

'Autofit to make columns neater, runs at end of script so itll adjust to
'all datasets calculated
Columns("A:Q").AutoFit
End Sub

