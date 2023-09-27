Attribute VB_Name = "Module1"
Sub repeat_sheets():

    Dim xSh As Worksheet
    
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call titles
        Call Stocks
        Call Colors
        
    Next
    
       Application.ScreenUpdating = True
End Sub

Sub titles():

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 17).Value = "Ticker"
Cells(1, 18).Value = "Value"
Cells(2, 16).Value = "Greatest % Increase"
Cells(3, 16).Value = "Greatest % Decrease"
Cells(4, 16).Value = "Greatest Total Volume"

End Sub

Sub Stocks():
 
  Dim ticker, year_date As String
  Dim Volume_Total, init, per_change, change As Double
  Volume_Total = 0
  Dim Summary_Table_Row, n  As Integer
  Summary_Table_Row = 2

n = Worksheets("A").UsedRange.Rows.Count
init = 2
  For i = 2 To n
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
      ticker = Cells(i, 1).Value
      Volume_Total = Volume_Total + Cells(i, 7).Value
      Range("I" & Summary_Table_Row).Value = ticker
      Range("L" & Summary_Table_Row).Value = Volume_Total
      
      Volume_Total = 0
      
      close_price = Cells(i, 6).Value
      open_price = Cells(init, 3).Value
      change = close_price - open_price
      Cells(Summary_Table_Row, 10).Value = change
      perc_change = ((close_price / open_price) - 1) * 100
      Cells(Summary_Table_Row, 11).Value = perc_change
      Summary_Table_Row = Summary_Table_Row + 1
      init = i + 1
    Else
      Volume_Total = Volume_Total + Cells(i, 7).Value
     
      
    End If
  Next i

End Sub

Sub Colors():

    
    Dim n, i As Integer
    
    
    
    n = Worksheets("A").UsedRange.Rows.Count
    
    
    
        For i = 2 To n
        
    If Cells(i, 11).Value > 0 Then
    Cells(i, 11).Interior.Color = RGB(0, 176, 80)
    
    
    ElseIf Cells(i, 11).Value < 0 Then
    Cells(i, 11).Interior.Color = RGB(204, 51, 0)
    
    
        End If
    Next i
   Call Max_Min_perctchange
   
End Sub


    Sub Max_Min_perctchange():
    

    Dim ws As Worksheet
    Dim lastRow, i As Long
    Dim currentPercentageChange, maxPercentageChange, minpercentagechange, maxVolume  As Double
    Dim CurrentMax_ticker, ticker_max, ticker_min As String

       
   For Each ws In ThisWorkbook.Worksheets
    
    
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    maxPercentageChange = ws.Cells(2, "K").Value
    minpercentagechange = ws.Cells(2, "K").Value
    ticker_max = ws.Cells(2, "I").Value
    ticker_min = ws.Cells(2, "I").Value
          
    For i = 2 To lastRow
    
        currentPercentageChange = ws.Cells(i, "K").Value
        CurrentMax_ticker = ws.Cells(i, "K").Value
        
        If currentPercentageChange > maxPercentageChange Then
           maxPercentageChange = currentPercentageChange
           CurrentMax_ticker = ticker_max
           
        End If
        
        currentPercentageChange = ws.Cells(i, "K").Value
        CurrentMin_Ticker = ws.Cells(i, "I").Value
        
        If currentPercentageChange < minpercentagechange Then
           minpercentagechange = currentPercentageChange
           ticker_min = CurrentMin_Ticker
           
        End If
        
    Next i
    
      ws.Cells(2, 18).Value = maxPercentageChange
      ws.Cells(3, 18).Value = minpercentagechange
      ws.Cells(2, 17).Value = ticker_max
      ws.Cells(3, 17).Value = ticker_min
      
   Next ws
  Call Max_volume
  
              
 End Sub
 
Sub Max_volume():

    Dim ws As Worksheet
    Dim lastRow_volume As Double, maxVolume As Double, currentvolume As Double
    Dim j As Long
    Dim ticker_max_vol, currentticker As String

   
    For Each ws In ThisWorkbook.Worksheets
    
    
    lastRow_volume = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    maxVolume = ws.Cells(2, "L").Value
    ticker_max_vol = ws.Cells(2, "I").Value

    For j = 2 To lastRow_volume
    
        currentvolume = ws.Cells(j, "L").Value
        currentticker = ws.Cells(j, "I").Value
        
        If currentvolume > maxVolume Then
            maxVolume = currentvolume
            ticker_max_vol = currentticker
            
            
        End If
    Next j
       ws.Cells(4, 18).Value = maxVolume
       ws.Cells(4, 17).Value = ticker_max_vol
    
    Next ws
        
    
End Sub




