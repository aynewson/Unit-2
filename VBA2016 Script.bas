Attribute VB_Name = "Module1"
Sub StockMarket()

Dim Ticker As String

Dim Ticker_Volume As Double
Ticker_Volume = 0

Dim Ticker_Summary As Double
Ticker_Summary = 2


Dim i As Double

For i = 2 To 797711

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        Ticker = Cells(i, 1).Value
        
        Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
        
        Range("I" & Ticker_Summary).Value = Ticker
        
        Range("J" & Ticker_Summary).Value = Ticker_Volume
        
        Ticker_Summary = Ticker_Summary + 1
        
        Ticker_Volume = 0
        
    Else
    
        Ticker_Volume = Ticker_Volume + Cells(i, 7).Value


    End If
    
    
Next i
        
End Sub
