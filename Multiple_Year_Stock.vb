Sub Multiple_year_stock_data():

Dim Ticker As String
Dim Volumen_Tot As Double, Delta_precio As Double, Delta_porcentaje As Double, max_delta_porcentaje As Double, min_delta_porcentaje As Double, volumen_max As Double
Dim registros As Integer
Dim open_, close_ As Double


For Each Worksheet In Worksheets
    

    Volumen_Tot = 0
    registros = 2
    
    LastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row

    open_ = Worksheet.Cells(2, 3)
    For i = registros To LastRow
    
        
        If Worksheet.Cells(i + 1, 1) <> Worksheet.Cells(i, 1) Then
            Ticker = Worksheet.Cells(i, 1)
            Volumen_Tot = Volumen_Tot + Worksheet.Cells(i, 7)
        
            close_ = Worksheet.Cells(i, 6)
           
            Delta_precio = close_ - open_
            
            
            If open_ = 0 Then
                Delta_porcentaje = 0
            Else
                Delta_porcentaje = Delta_precio / open_ * 100
           End If
            
            Worksheet.Range("I" & registros).Value = Ticker
            Worksheet.Range("J" & registros).Value = Delta_precio
            Worksheet.Range("K" & registros).Value = (Delta_porcentaje & "%")
            Worksheet.Range("L" & registros).Value = Volumen_Tot
            
             
            Volumen_Tot = 0
       
            open_ = Worksheet.Cells(i + 1, 3)
            
            
             If Worksheet.Range("J" & registros).Value >= 0 Then
                Worksheet.Range("J" & registros).Interior.ColorIndex = 4
            
            Else
            
                Worksheet.Range("J" & registros).Interior.ColorIndex = 3
            
                End If
            
                registros = registros + 1
        Else
            Volumen_Tot = Volumen_Tot + Worksheet.Cells(i, 7)
           
        
        
       
        
        End If
        

        
    Next i
    
Next Worksheet


End Sub

