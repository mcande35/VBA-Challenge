Attribute VB_Name = "Module1"
Sub VBAHomework()
Dim Ticker As String 'for word

Dim table_row As Double 'for numbers

Dim opening_value As Double

Dim closing_value As Double

Dim Yearly_change As Double

Dim Percentage_change As Double

Dim ticker_volume As Double

Dim greatest_percent As Double

Dim WS As Worksheet

For Each WS In Sheets

ticker_volume = 0

table_row = 2

lastrow_main = Cells(Rows.Count, 1).End(xlUp).Row 'given

opening_value = Cells(2, 3).Value 'outside of loop since it is the first


For i = 2 To lastrow_main

    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    Ticker = WS.Cells(i, 1).Value
    
    Range("I" & table_row).Value = Ticker
    
    closing_value = WS.Cells(i, 6).Value 'just i because it is the current
    
    Yearly_change = closing_value - opening_value
    
    WS.Range("J" & table_row).Value = Yearly_change
    
    Percentage_change = Yearly_change / opening_value
    
    WS.Range("K" & table_row).Value = Percentage_change
    
    opening_value = Cells(i + 1, 3).Value
    
    WS.Range("L" & table_row).Value = ticker_volume + WS.Cells(i, 7).Value
    
    table_row = table_row + 1
      
     Else
     ticker_volume = WS.Cells(i, 7).Value + ticker_volume 'see it as a set
     
    End If
    Next i
    
  Next WS
  
  End Sub
