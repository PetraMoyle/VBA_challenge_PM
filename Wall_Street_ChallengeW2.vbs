Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

    Range("I1").Value = "Ticker Symbol"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Dim Ticker_Symbol As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As LongLong
    Dim Summary_Table_Row As Double
    Dim Opening_Value As Double
    Dim Closing_Value As Double
    Dim i As Long

    Summary_Table_Row = 2
    Total_Stock_Volume = 0
    Opening_Value = Cells(2, 3)
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        
        For i = 2 To LastRow
        
         
        
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                     Ticker_Symbol = Cells(i, 1).Value
                     Cells(Summary_Table_Row, 9).Value = Ticker_Symbol
                     
                     Closing_Value = Cells(i, 6).Value
                       
                     Yearly_Change = Closing_Value - Opening_Value
                     Cells(Summary_Table_Row, 10).Value = Yearly_Change
                     
                     Percent_Change = Round((Yearly_Change / Opening_Value) * 100, 2)
                     Cells(Summary_Table_Row, 11).Value = Percent_Change
        
                     Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                     Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
                     
                     Total_Stock_Volume = 0
                     
                     Opening_Value = Cells(i + 1, 3).Value
                
                     Summary_Table_Row = Summary_Table_Row + 1
                     
                 Else
                
                    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
                 End If
       
        Next i

End Sub
        
Sub Formatting()

Dim i As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To LastRow
        
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.Color = RGB(0, 255, 0)
            
            Else
                Cells(i, 10).Interior.Color = RGB(255, 0, 0)
            
            End If
            
        Next i
        
    
End Sub


