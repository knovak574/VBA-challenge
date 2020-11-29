Attribute VB_Name = "Module1"
Sub ticker_symbol()

    Dim i As Integer
    Dim ws_num As Integer
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet
    ws_num = ThisWorkbook.Worksheets.Count


For i = 1 To ws_num
    ThisWorkbook.Worksheets(i).Activate
                
Call yearly_volume

Next
starting_ws.Activate

End Sub

Sub yearly_volume()

    'set X as row number and Y as ticker symbol'
    'set number of rows in spreadsheet which is huge and will vary from sheet to sheet'
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
        
    Dim LastRow As Long
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    Range("A1").Select
    


    'set columns that contains ticker, volume, and total volume'
    ticker = 1
    opening = 3
    closing = 6
    volume = 7
    Symbol = 9
    yearly_change = 10
    percent_change = 11
    total_volume = 12
    Y = 2
    'new'
    Z = 1
    

        For X = 2 To NumRows
          
    
            If Cells(X, ticker).Value = Cells(X + 1, ticker).Value Then
            Cells(Y, total_volume).Value = (Cells(X, volume) + Cells(Y, total_volume))
            

    
            Else: Cells(Y, total_volume).Value = (Cells(X, volume) + Cells(Y, total_volume))
                  Cells(Y + 1, total_volume).Value = (Cells(X + 1, volume) + Cells(Y + 1, total_volume))
                  Cells(Y, Symbol).Value = Cells(X, ticker)
            
                  
                  Cells(Y, percent_change).Value = Round((Cells(Y, yearly_change) / Cells(Z + 1, opening)) * 100, 2)
                  Cells(Y, yearly_change).Value = (Cells(X, closing) - Cells(Z + 1, opening))
                  
                  If Cells(Y, yearly_change).Value > 0 Then Cells(Y, yearly_change).Interior.ColorIndex = 10
                  If Cells(Y, yearly_change).Value < 0 Then Cells(Y, yearly_change).Interior.ColorIndex = 3
                      
    
                  Y = Y + 1
                  Z = X
        
            End If
   
         
        Next X
End Sub

