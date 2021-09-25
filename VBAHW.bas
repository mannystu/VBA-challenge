Attribute VB_Name = "Module1"
Sub stock_info()

  
 
  
  ' Set an initial variable for holding the ticker symbol
  Dim ticker As String

  ' Set an initial variable for holding the total volume per stock
  Dim volume As Double
  volume = 0
  
  ' Set an initial variable for holding the opening stock price
  Dim Open_Price As Double
  
  ' Set an initial variable for holding the closing stock price
  Dim Close_Price As Double
  
  ' Set an initial variable for holding the closing stock price
  Dim Change_Price As Double

  ' Mah Row Tracker
  Dim Stock_info_row As Integer
  
  ' Mah greatest hits declaration y'all
  Dim great_gain_ticker As String
  
  Dim great_loss_ticker As String
  
  Dim great_vol_ticker As String
  
  Dim great_gain As Double
  
  Dim great_loss As Double
  
  Dim great_vol As Double
  
  Dim per_gain As Double
  
'cycle throu mah sheets
For n = 1 To Sheets.Count
    
 Worksheets(Sheets(n).Name).Activate
  
  'Print the Headings
  Range("I1").Value = "Ticker Symbol"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Volume"
  
  Range("O2").Value = "Greatest % Increase"
  Range("O3").Value = "Greatest % Decrease"
  Range("O4").Value = "Greatest Total Volume"
  Range("P1").Value = "Ticker"
  Range("Q1").Value = "Value"
  
  'initialize the greatest values
  
  great_gain = 0
  great_loss = 0
  great_vol = 0

'assuming every sheet has headers. start with row 2
    Stock_info_row = 2
        
        
  ' Loop through all rows of stock info
  For i = 2 To ActiveSheet.UsedRange.Rows.Count

    ' Check if we are still at the same stock ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker name
      ticker = Cells(i, 1).Value

      ' Set the closing price value
      Close_Price = Cells(i, 6).Value
      
          
      ' Add to the volume Total
      volume = volume + Cells(i, 7).Value

           
      ' Print the ticker symbol to the summary table
      Range("I" & Stock_info_row).Value = ticker

      ' Print the Change in Price
      Change_Price = Close_Price - Open_Price
      Range("J" & Stock_info_row).Value = Change_Price
      
      'Conditional Formating for Change in Price
      If Change_Price > 0 Then
        Range("J" & Stock_info_row).Interior.Color = RGB(204, 255, 204)
        'I want mah % to be colored too
        Range("K" & Stock_info_row).Interior.Color = RGB(204, 255, 204)
        
        Else
            Range("J" & Stock_info_row).Interior.Color = RGB(254, 206, 206)
            Range("K" & Stock_info_row).Interior.Color = RGB(254, 206, 206)
      End If
      
      
        
      
      ' Print the % Change in Price
      If Open_Price = 0 Then
            Range("K" & Stock_info_row).Value = 0
      
        Else
            per_gain = (Close_Price - Open_Price) / Open_Price
            Range("K" & Stock_info_row).Value = per_gain
      End If
      
      'conditinals for summary of greatest stuff
      
      If per_gain > great_gain Then
        great_gain = per_gain
        great_gain_ticker = ticker
        
      End If
      
      If per_gain < great_loss Then
        great_loss = per_gain
        great_loss_ticker = ticker
        
      End If
      
      If volume > great_vol Then
        great_vol = volume
        great_vol_ticker = ticker
        
      End If
      
      
      ' Print the Volume Amount the summary table
      Range("L" & Stock_info_row).Value = volume

      
      ' Reset the volume
      volume = 0

     ' Moves to next cell row
      Stock_info_row = Stock_info_row + 1
    
    ' If the cell immediately following a row is the same brand...
    ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then

      'Set the Opening Price
      Open_Price = Cells(i, 3).Value
      
    Else
      ' Add to the stock volume
      volume = volume + Cells(i, 7).Value
      
       
       

    End If



  Next i

'print mah greatest table vales
Range("P2").Value = great_gain_ticker
Range("P3").Value = great_loss_ticker
Range("P4").Value = great_vol_ticker
Range("Q2").Value = great_gain
Range("Q3").Value = great_loss
Range("Q4").Value = great_vol

'ugh! them numbers be ugly. make em pretty
Range("Q2").NumberFormat = "0.00%"
Range("Q3").NumberFormat = "0.00%"
Worksheets(Sheets(n).Name).Columns("K").EntireColumn.NumberFormat = "0.00%"
Worksheets(Sheets(n).Name).Columns("I:Q").EntireColumn.AutoFit

Next n



End Sub
