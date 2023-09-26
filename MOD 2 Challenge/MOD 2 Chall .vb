Sub VBACHAL()
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_vol As Double
    Dim row_num As Integer
    Dim ws As Worksheet
    Dim maxpercentageIncrease As Double
    Dim maxpercentageticker As String
    Dim minpercentageIncrease As Double
    Dim minpercentageticker As String
    Dim total_volhold As Double
    Dim total_volticker As String
    Dim colorval As Double
    
    
    For Each ws In Worksheets
        Worksheets(ws.Name).Activate
        total_volume = 0
        row_num = 2
        maxpercentageIncrease = 0
        colorval = 0

        date_first_row = Cells(2, 2).Value
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        date_last_row = Cells(last_row, 2).Value
        previousValue = ws.Cells(2, 1).Value
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "maxPercentageIncrease"
        Cells(2, 16).Value = "minPercentageIncrease"
        Cells(1, 17).Value = "Ticker Symbol"
        Cells(1, 18).Value = "Value"
        
        
        
        
        
            For i = 2 To last_row
                total_vol = total_vol + Cells(i, 7).Value
                
                If Cells(i, 2).Value = date_first_row Then
                    ticker = Cells(i, 1).Value
                    open_price = Cells(i, 3).Value
                
                ElseIf Cells(i, 2).Value = date_last_row Then
                    close_price = Cells(i, 6).Value
                    
                    yearly_change = close_price - open_price
                    percent_change = Round((yearly_change / open_price) * 100, 2)
                    
                    If percent_change > maxpercentageIncrease Then
                        maxpercentageIncrease = percent_change
                        maxpercentageticker = ticker
              
                    End If
                    
                    If percent_change > mixpercentageIncrease Then
                        minpercentageIncrease = percent_change
                        minpercentageticker = ticker
                    
                    End If
                    
                    
                    If total_vol > total_volhold Then
                        total_volhold = total_vol
                        total_volticker = ticker
                    
                    End If
                    








                    If yearly_change > colorval Then
                        Cells(row_num, 10).Interior.Color = RGB(255, 0, 0)
                
                    
                    End If


                    If yearly_change < colorval Then
                        Cells(row_num, 10).Interior.Color = RGB(0, 255, 0)
                
                    
                    End If



                    Cells(row_num, 9).Value = ticker
                    Cells(row_num, 10).Value = yearly_change
                    Cells(row_num, 11).Value = percent_change & "%"
                    Cells(row_num, 12).Value = total_vol
                    row_num = row_num + 1
                    total_vol = 0
                End If
                
            Next i
            Cells(2, 18).Value = maxpercentageIncrease
            Cells(2, 17).Value = maxpercentageticker
            Cells(3, 18).Value = minpercentageIncrease
            Cells(3, 17).Value = minpercentageticker
            Cells(4, 18).Value = total_volhold
            Cells(4, 17).Value = total_volticker
    Next ws
       

End Sub
