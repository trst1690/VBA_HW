Attribute VB_Name = "Module1"
Sub stocks()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

Dim Ticker_name As String
'determine last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim yearly_change As Double
Dim open_price As Double
Dim percent_change As Double

Cells(1, "m").Value = "Ticker"
Cells(1, "n").Value = "Yearly Change"
Cells(1, "p").Value = "Total Volume"
Cells(1, "o").Value = " Percentage change"

yearly_change = 0

Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

open_price = Cells(2, 3)
For i = 2 To lastrow

 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ' set tickers
    Ticker_name = Cells(i, 1).Value
    
    'set close price
    close_price = Cells(i, 6).Value
    
   
    
    'calculate yearly change
    yearly_change = (close_price - open_price)
    
    'find percent change
    percent_change = yearly_change / open_price
    
    
   
    
    
    
    
    ' print ticker names
    
    Range("m" & Summary_Table_Row).Value = Ticker_name
    
    'print yearly change
     Range("n" & Summary_Table_Row).Value = yearly_change
     
     'print percent_change
     Range("o" & Summary_Table_Row).Value = percent_change
     Range("o" & Summary_Table_Row).NumberFormat = "0.00%"
     
     ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
       
       
       ' Reset the yearly change
             yearly_change = 0
      'Reset open_price
            open_price = Cells(i + 1, 3).Value
      
       End If
       
Next i





Dim ticker2 As String
Dim volume As Double
volume = 0

Dim Summary_Table_Roww As Integer
Summary_Table_Roww = 2
ticker2 = Cells(2, 1).Value

For i = 2 To lastrow + 1


If Cells(i + 1, 1).Value <> Cells(i, 1) Then

 volume = volume + Cells(i, 7).Value

 ' Add one to the summary table row
      Summary_Table_Roww = Summary_Table_Roww + 1



 ' Print the volume to the Summary Table
      Range("P" & Summary_Table_Roww - 1).Value = volume
       
        
        ' reset volume
  
    volume = 0
   
    
    Else
    volume = volume + Cells(i, 7).Value
     
    
    
End If
Next i

lastrow2 = ws.Cells(Rows.Count, 14).End(xlUp).Row


For i = 2 To lastrow2

        If Cells(i, 14).Value > 0.00000001 Then
        Cells(i, 14).Interior.ColorIndex = 10
Else
        Cells(i, 14).Interior.ColorIndex = 3
 
 End If
 Next i

Lastrow3 = ws.Cells(Rows.Count, 15).End(xlUp).Row

For i = 2 To Lastrow3

    If Cells(i, 15).Value > 0.00000000000001 Then
        Cells(i, 15).Interior.ColorIndex = 10

Else
    Cells(i, 15).Interior.ColorIndex = 3

End If
Next i





Cells(1, "r").Value = "Greatest % Increase"
Cells(2, "r").Value = "Greatest % Decrease"
Cells(3, "r").Value = "Greatest Total Volume"


Dim perRangeO As Range
Dim perRangeP As Range
Dim Greatest_percent_change As Double
Dim lastrow4 As Long
Dim Greatest_percent_decrease As Double
Dim Greatest_total_volume As Double
Dim lastrow5 As Long




' Finding greatest percent change

lastrow4 = Range("O" & Rows.Count).End(xlUp).Row
Set perRangeO = Range("o2:o" & lastrow4)


Greatest_percent_change = Application.WorksheetFunction.Max(perRangeO)
Cells(1, "s").Value = Greatest_percent_change
Cells(1, "s").NumberFormat = "0.00%"



' finding greatest percent decrease
Greatest_percent_decrease = Application.WorksheetFunction.Min(perRangeO)
Cells(2, "s").Value = Greatest_percent_decrease
Cells(2, "s").NumberFormat = "0.00%"

' find greatest total volume
lastrow5 = Range("P" & Rows.Count).End(xlUp).Row
Set perRangeP = Range("p2:p" & lastrow5)

    Greatest_total_volume = Application.WorksheetFunction.Max(perRangeP)
    Cells(3, "s").Value = Greatest_total_volume

Next ws













End Sub

