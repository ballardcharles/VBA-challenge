Attribute VB_Name = "Module1"
Sub Stock_Analysis():

For Each ws In Worksheets

    ' Set Stock Analysis Summary Table Headers
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Yearly Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O1").Value = "Ticker Symbol"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest Percent Increase"
    ws.Range("N3").Value = "Greatest Percent Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"

    ' Declare Variables
    Dim Ticker_Symbol As String
    Dim Total_Volume As Double
    Dim Yearly_Change As Double
    Dim Yearly_Percent_Change As Double
    Dim Yearly_Open As Double
    Dim Yearly_Close As Double
    Dim Price As Double
    Dim Stock_Analysis_Summary As Integer
    Dim LastRow As Long
    Dim Stock_Change_Summary As Integer
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Volume As Long
    Dim Greatest_Increase_Symbol As String
    Dim Greatest_Decrease_Symbol As String
    Dim Greatest_Volume_Symbol As String

    ' Set initial values
    Total_Volume = 0
    Price = 2
    Stock_Analysis_Summary = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Stock_Change_Summary = 2

    ' Loop through the data to find ticker symbol, yearly change, percent change, and total volume
    For i = 2 To LastRow
    
     ' Check that we are still on the same ticker symbol, if not move to next symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ' Set the Ticker Symbol
            Ticker_Symbol = ws.Cells(i, 1).Value
    
            ' Add to the Total Stock Volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
            ' Print Ticker Symbol into Data Analysis Summary Table
            ws.Range("I" & Stock_Analysis_Summary).Value = Ticker_Symbol
    
            ' Print Total volume into Stock Analysis Summary Table
            ws.Range("L" & Stock_Analysis_Summary).Value = Total_Volume
        
            ' Determine stock price change from year open to year close
            Yearly_Open = ws.Range("C" & Price).Value
            Yearly_Close = ws.Range("F" & i).Value
            Yearly_Change = Yearly_Close - Yearly_Open
        
            ' Pring the yearly change for each symbol
            ws.Range("J" & Stock_Analysis_Summary).Value = Yearly_Change
        
            ' Determine the percent change from year open to year close
            If Yearly_Open = 0 Then
                Yearly_Percent_Change = 0
            Else
                Yearly_Open = ws.Range("C" & Price).Value
                Yearly_Percent_Change = Yearly_Change / Yearly_Open
            End If
            
            ' Print the yearly percent change in the Stock Analysis summary table
            ws.Range("K" & Stock_Analysis_Summary).NumberFormat = "0.00%"
            ws.Range("K" & Stock_Analysis_Summary) = Yearly_Percent_Change
        
            ' Set conditional formating for yearly percent change
            If ws.Range("J" & Stock_Analysis_Summary) >= 0 Then
                ws.Range("J" & Stock_Analysis_Summary).Interior.ColorIndex = 4
            Else
                ws.Range("J" & Stock_Analysis_Summary).Interior.ColorIndex = 3
            End If
        
            ' Add to the Stock Analysis Summary Table row
            Stock_Analysis_Summary = Stock_Analysis_Summary + 1
            Price = i + 1
        
            ' Reset Total Stock Volume
            Total_Volume = 0
        
        Else
            ' Add to Total Volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
     
        End If
    
    Next i

        LastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        ' Determine the stock with the greatest percent increase over the year and print to table
        Greatest_Percent_Increase = WorksheetFunction.Max(ws.Range("K:K"))
        Greatest_Increase_Symbol = WorksheetFunction.Match(Greatest_Percent_Increase, ws.Range("K:K"), 0)
        ws.Range("O2").Value = ws.Cells(Greatest_Increase_Symbol + 0, 9)
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P2").Value = Greatest_Percent_Increase
    
        ' Determine the stock with the greatest percent decrease over the year and print to table
        Greatest_Percent_Decrease = WorksheetFunction.Min(ws.Range("K:K"))
        Greatest_Decrease_Symbol = WorksheetFunction.Match(Greatest_Percent_Decrease, ws.Range("K:K"), 0)
        ws.Range("O3").Value = ws.Cells(Greatest_Decrease_Symbol + 0, 9)
        ws.Range("P3").NumberFormat = "0.00%"
        ws.Range("P3").Value = Greatest_Percent_Decrease

        'Determine the stock with the greatest total volume over the year and print to table
        Greatest_Vol = WorksheetFunction.Max(ws.Range("L:L"))
        Greatest_Volume_Symbol = WorksheetFunction.Match(Greatest_Vol, ws.Range("L:L"), 0)
        ws.Range("O4").Value = ws.Cells(Greatest_Volume_Symbol + 0, 9)
        ws.Range("P4").NumberFormat = "0"
        ws.Range("P4").Value = Greatest_Vol
    
        ws.Range("I:P").Columns.AutoFit
    Next ws
    
End Sub

