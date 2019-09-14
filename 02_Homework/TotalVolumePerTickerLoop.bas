Attribute VB_Name = "TotalVolumePerTickerLoop"
Sub HomeworkVBAVolume()

    Dim Ticker As String
    Dim Totalvolume As String
    Dim NextTick As String
    Dim i As Double
    Dim UniqueTick As Double
    Dim LastRow As Double
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Totalvolume = 0
    UniqueTick = 1
    
        'Do unique total volume first, then add around it unique ticker
        For i = 1 To LastRow
            Ticker = Cells(i, 1).Value
            NextTick = Cells(i + 1, 1).Value
            
        If Ticker = NextTick Then
                'this line references the information we want by value
                Totalvolume = Totalvolume + Cells(i + 1, 7).Value
            
            ElseIf Ticker <> NextTick Then
                Cells(UniqueTick, 10).Value = Totalvolume
                
            ' hold variables by row label of each ticker into the label and total columns located at 9 and 10
            Cells(UniqueTick, 10).Value = Totalvolume
            Cells(UniqueTick, 9).Value = Ticker
            
        UniqueTick = UniqueTick + 1
        
        'this line totals the volume located in the 7th column within the loop defined at the top
        Totalvolume = Cells(i + 1, 7).Value
            
            End If
        Next i
        
        
    'label the column headers specifically - just easy writing into cells
    Cells(1, 10) = "total stock volume"
    Cells(1, 9) = "ticker symbol"
        
End Sub

Sub RepeatonEverySheet()
    
    Dim WS As Worksheet
    For Each WS In Worksheets
    WS.Activate
    
    Call HomeworkVBAVolume
    
    Next WS
    
End Sub
