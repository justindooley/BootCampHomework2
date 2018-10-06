Attribute VB_Name = "Module1"
Sub tickertotaler_moderate()


'define everything
Dim ws As Worksheet
Dim Ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Long
Dim Summary_Table As Integer

vol = 0
year_open = 0
year_close = 0
yearly_change = 0
percent_change = 0



'run through each worksheet
For Each ws In ThisWorkbook.Worksheets

    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'setup integers for loop
    Summary_Table = 2

    'loop
      
      
        For i = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                'find all the values
                Ticker = ws.Cells(i, 1).Value
                vol = vol + ws.Cells(i, 7).Value
                year_open = ws.Cells(i, 3).Value
                year_close = ws.Cells(i + 261, 6).Value

                year_open = year_open
                year_close = year_close
            
                yearly_change = year_close - year_open
                percent_change = (yearly_change / year_open)
                
                
            'reset everything
                yearly_change = 0
                percent_change = 0

                'insert values into summary
                ws.Cells(Summary_Table, 9).Value = Ticker
                ws.Cells(Summary_Table, 10).Value = yearly_change
                ws.Cells(Summary_Table, 11).Value = percent_change
                ws.Cells(Summary_Table, 12).Value = vol
                Summary_Table = Summary_Table + 1
                
            'reset everything
                vol = 0
                year_open = 0
                year_close = 0
                yearly_change = 0
                percent_change = 0
                
           Else
            vol = vol + ws.Cells(i, 7).Value
            year_open = ws.Cells(i, 3).Value
            year_close = ws.Cells(i + 261, 6).Value
            yearly_change = year_close - year_open
            percent_change = (yearly_change / year_open)
    

            End If
              
    'finish i loop
        Next i
            
     
    ws.Columns("K").NumberFormat = "0.00%"
    


'format columns colors
   '     Dim rg As Range
   '     Dim g As Double
   '     Dim color_cell As Range
   
   '     Set rg = ws.Range("J2", Range("J2").End(xlDown))
        
        
   '     For g = 1 To ws.Range("J2").rows.count


   '     Set color_cell = rg(g)
   '     Select Case color_cell
   '         Case Is >= 0
   '             With color_cell
   '                 .Interior.Color = vbGreen
   '             End With
   '         Case Is < 0
   '             With color_cell
   '                 .Interior.Color = vbRed
   '             End With
   '        End Select
   '     Next g


         '  If ws.Cells(i, 10).Value >= 0 Then
          ' ws.Cells(i, 10).Interior.Color.Index = 4
           'Else
          ' ws.Cells(i, 10).Interior.Color.Index = 3
           'End If


'move to next worksheet
Next ws

End Sub
