Attribute VB_Name = "Module1"
Sub Stock_Info()

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "% Change"
Range("L1") = "Stock Volume"

'set initial variable for holding ticker symbol
Dim Stock_Ticker As String

'set intital variable for opening price
Dim Open_Price As Double
Open_Price = 0

'set intital variable for opening price
Dim Close_Price As Double
Close_Price = 0

'set initial variable for yearly change
Dim Yearly_Change As Double
Yearly_Change = 0

'set initial variable for percentage change
Dim Percentage_Change As Double
Percentage_Change = 0

'set initial variable for holding trading volume per ticker symbol
Dim Volume_Total As Double
Volume_Total = 0

'keep track of location for each ticker symbol in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'determine the last row
LastRow = Cells(Rows.Count, 2).End(xlUp).Row

'loop through all ticker symbols
For i = 2 To LastRow

    'check to see if still within same ticker symbol and if not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'set the ticker name
    Stock_Ticker = Cells(i, 1).Value
               
    'set the opening price
    Open_Price = Cells(i, 3).Value
              
    'set the opening price
    Close_Price = Cells(i, 6).Value
        
    'add the Yearly Change by Ticker
    Yearly_Change = (Close_Price - Open_Price)
            
    'add the Percentage Change by Ticker
    Percentage_Change = (Yearly_Change / Close_Price)
                    
    'add volume to the ticker total
    Volume_Total = Volume_Total + Cells(i, 7).Value
    
    'display ticker symbol in summary table
    Range("I" & Summary_Table_Row).Value = Stock_Ticker
    
    'display yearly change in summary table with color coding
    Range("J" & Summary_Table_Row).Value = Yearly_Change
    Range("J" & Summary_Table_Row).NumberFormat = "0.00"
    
        If (Yearly_Change >= 0) Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf (Yearly_Change < 0) Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
          
    'display percentage change in summary table
    Range("K" & Summary_Table_Row).Value = Percentage_Change
    Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    
    'display ticker volume in summary table
    Range("L" & Summary_Table_Row).Value = Volume_Total
    
    'add to next row of summary table
    Summary_Table_Row = Summary_Table_Row + 1
    
    'reset total per ticker symbol
    Volume_Total = 0
    
    'if cell above is same
Else

    'add ticker volume to total
    Volume_Total = Volume_Total + Cells(i, 7).Value
    
End If

Next i

End Sub
