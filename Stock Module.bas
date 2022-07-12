Attribute VB_Name = "Module7"
Sub Stonks()


'Assign

Dim DayTicker As String
Dim DayDate As Long
Dim DayOpen As Double
Dim DayClose As Double
Dim DayVolume As Long

Dim YearVolume As Long
Dim YearClosePrice As Double
Dim yearOpenPrice As Double

Dim YearlyChange As Double
Dim PercentChange As Double


Dim CurrentRow As Integer
CurrentRow = 2

Dim checkpoint As Integer



    'I Loop
    For i = 2 To 753002
    
    
    
    
    
    'If Ticker is in different group, calculate and  print
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
        DayTicker = Cells(i, 1).Value
        Range("I" & CurrentRow).Value = DayTicker
        
        
        TotalVolume = TotalVolume + Cells(i, 7).Value
        Range("P" & CurrentRow).Value = TotalVolume
        
        YearClosePrice = Cells(i, 6).Value
        Range("M" & CurrentRow).Value = YearClosePrice
        
        Range("L" & CurrentRow).Value = yearOpenPrice
        
        YearlyChange = YearClosePrice - yearOpenPrice
        Range("N" & CurrentRow).Value = YearlyChange
        
'            If Range("N" & CurrentRow).Value > 0 Then
'            Cells(i, 14).Interior.ColorIndex = 3
'
'            ElseIf Range("N" & CurrentRow).Value < 0 Then
'
'           Cells(i, 14).Interior.ColorIndex = 4
            
'            End If
        
        PercentChange = (Range("N" & CurrentRow).Value / yearOpenPrice)
        Range("O" & CurrentRow).Value = PercentChange
        
        PecentChange = 0
        YearlyChange = 0
        DayTicker = 0
        TotalVolume = 0
        YearClosePrice = 0
        yearOpenPrice = 0
        CurrentRow = CurrentRow + 1
        
    Else
         
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
    
            If i = 2 Or Cells(i, 2).Value < Cells(i - 1, 2).Value Then
                yearOpenPrice = Cells(i, 3).Value
            End If
                
       
                

                
                
    
            
    End If
    
    Next i
    
    
  
    
End Sub

