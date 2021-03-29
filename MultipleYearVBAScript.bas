Attribute VB_Name = "Module1"
Sub Stockmarket():

Dim Ticker As String
Dim Yearly As Double
Dim Percent As Double
Dim Summary As Integer
Dim OpenP As Double
Dim CloseP As Double
Dim Volume As Double
Dim i As Long
Dim j As Long
Dim ws As Worksheet
Dim Lastrow As Long


WS_count = ActiveWorkbook.Worksheets.Count

Summary = 2
j = 1


Set ws = ActiveWorkbook.ActiveSheet

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

Cells(1, "I").Value = "Ticker Symbol"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Stock Volume"


    'Open Price
      OpenP = Cells(2, j + 2).Value
      
      
     For i = 1 To Lastrow
                If Cells(i + 1, j).Value <> Cells(i, j).Value Then
                
                    'Ticker
                    Ticker = Ticker + Cells(i + 1, j).Value
                

                    Range("I" & Summary).Value = Ticker

                   'Close Price
                    CloseP = Cells(i + 1, j + 6).Value

                   
                    
                    'Yearly
                    Yearly = CloseP - OpenP
                    Range("J" & Summary).Value = Yearly


                            'Percent Change
                            If OpenP = CloseP Then
                                Percent = 0
                            Else
                                Percent = Yearly / OpenP

                            Range("K" & Summary).Value = Percent

                            End If
                            
                        
                    'Total Volume
                     Volume = Volume + Cells(i + 1, 7 + j).Value
                    
                    Range("L" & Summary).Value = Volume
                    
                    
                    
            'Reset Volume and Open Price
            Volume = 0
            OpenP = Cells(i + 1, j + 3)
            
            End If


Next i

End Sub



