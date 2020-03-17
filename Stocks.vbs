Sub Stocks()
  Dim ws As Worksheet
  
  For Each ws In Worksheets
    
    Dim k As Integer
    k = 2
    Dim i As Long
    i = 2
    Dim LastRow As Long
    Dim opening_value As Double
    Dim closing_value As Double
    Dim volume As Double
    volume = 0
    Dim firstticker As Integer
    firstticker = 0
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim firsttime As Integer
    firsttime = 0
    
    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        
'Label the headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "% Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
     'Loop through rows in the column Print the Ticker Symbol
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            'Add up the volume
            volume = volume + ws.Cells(i, 7)
            firsttime = firsttime + 1
            'get opening stock price
                If firsttime = 1 Then
                    opening_value = ws.Cells(i, 3)
                Else
                End If
            Else
            'Add up the volume
            volume = volume + ws.Cells(i, 7)
            ws.Cells(k, 12) = volume
            
            'Get the Ticker symbol
            ws.Cells(k, 9) = ws.Cells(i, 1)
        
            
            'get the closing price
            closing_value = ws.Cells(i, 6)
            
            If opening_value <> 0 Then
                'Calculate yearly change
                yearly_change = closing_value - opening_value
                
            
                'calculate % change
                percent_change = ((closing_value - opening_value) / opening_value) * 100
            Else
                yearly_change = 0
                percent_change = 0
            End If
            'Print percent change and yearly change
            ws.Cells(k, 11) = percent_change
            ws.Cells(k, 10) = yearly_change
            'If yearly change is positive then Green or then red
            If ws.Cells(k, 10).Value > 0 Then
                ws.Cells(k, 10).Interior.Color = vbGreen
            Else
                ws.Cells(k, 10).Interior.Color = vbRed
            End If
            
            'increment counter
            k = k + 1
            'reset volume
            volume = 0
            'reset firsttime
            firsttime = 0
        End If
    Next i
    
    Dim j As Integer
    j = 2
    Dim m As Integer
    
    Dim newlastrow As Long
    Dim MinValue As Double
    Dim MaxValue As Double
    Dim GreatestVol As Double
    Dim HighTckr As String
    Dim LowTckr As String
    Dim MaxValue2 As Double
    Dim HighTckr2 As String
  
    
    
    ' Determine the Last Row
    newlastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
         
    'Label the headers of new Table
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greaest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    
        For j = 2 To newlastrow
            If ws.Cells(j, 11).Value > MaxValue Then
                MaxValue = ws.Cells(j, 11).Value
                HighTckr = ws.Cells(j, 9)
            Else
            End If
            If ws.Cells(j, 11).Value < MinValue Then
                MinValue = ws.Cells(j, 11).Value
                LowTckr = ws.Cells(j, 9).Value
            Else
            End If
            
            If ws.Cells(j, 12).Value > MaxValue2 Then
                MaxValue2 = ws.Cells(j, 12).Value
                HighTckr2 = ws.Cells(j, 9)
            Else
            End If
            
        Next j
        ws.Cells(2, 15).Value = HighTckr
        ws.Cells(3, 15).Value = LowTckr
        ws.Cells(2, 16).Value = MaxValue
        ws.Cells(3, 16).Value = MinValue
        ws.Cells(4, 15).Value = HighTckr2
        ws.Cells(4, 16).Value = MaxValue2
        
        'Reset Value
        HighTckr = 0
        HighTckr2 = 0
        LowTckr = 0
        MinValue = 0
        MaxValue = 0
        MaxValue2 = 0
   
 Next ws
End Sub
    
    
