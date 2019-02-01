Attribute VB_Name = "Module1"
Sub VbaStock()

    Dim lastrow As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim opening As Double
    Dim closing As Double

    Dim ws As Worksheet

    ' Sheet Loop
    
    For Each ws In Worksheets
        ws.Activate
        
        lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row  ' number of rows
        
        opening = Cells(2, 3) ' the first stock opening value
        j = 1
        For i = 2 To lastrow
            
            If Cells(i, 1) <> Cells(j, 9) Then
                j = j + 1
                Cells(j, 9) = Cells(i, 1)
                Cells(j, 12) = Cells(i, 7)
                If i > 2 Then
                    closing = Cells(i - 1, 6) ' stock closing volume
                If opening <> 0 Then ' to avoid division with 0
                    Cells(j - 1, 11) = -1 * (opening - closing) / opening
                End If
                Cells(j - 1, 10) = closing - opening 'red/green value and filling
                If Cells(j - 1, 10).Value >= 0 Then
                    Cells(j - 1, 10).Interior.ColorIndex = 4
                Else
                    Cells(j - 1, 10).Interior.ColorIndex = 3
                End If
             opening = Cells(i, 3)
             End If
                
            Else
                Cells(j, 12) = Cells(j, 12) + Cells(i, 7)
            End If
    
        Next i
        
            ' the last stock value and filling
            closing = Cells(lastrow, 6)
            Cells(j, 10) = opening - closing
            Cells(j, 11) = -1 * (opening - closing) / opening
            If Cells(j, 10).Value >= 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            Else
                Cells(j, 10).Interior.ColorIndex = 3
                End If
                
        ' ==========================
        ' Getting the min/max values
        ' ==========================
          
         k = 2
         Min = 0
         Max = 0
         maxvolume = 0
         
         Do While Cells(k, 11) <> ""
            If Cells(k, 11) > Max Then
                Max = Cells(k, 11)
                maxtick = Cells(k, 9)
                End If
            If Cells(k, 11) < Min Then
                Min = Cells(k, 11)
                mintick = Cells(k, 9)
                End If
            If Cells(k, 12) > maxvolume Then
                maxvolume = Cells(k, 12)
                maxvolumeTick = Cells(k, 9)
                End If
            k = k + 1
            Loop
            
            ' =================================================
            ' printing, formating and fitting of the max values
            ' =================================================
            
            Range("K2:K" & k).NumberFormat = "0.00%"
            Range("J2:J" & k).NumberFormat = "0.00"
            Cells(3, 17) = Min
            Cells(2, 17) = Max
            Range("Q2:Q3").NumberFormat = "0.00%"
            Cells(4, 17) = maxvolume
            Cells(3, 16) = mintick
            Cells(2, 16) = maxtick
            Cells(4, 16) = maxvolumeTick
            Cells(3, 15) = "Greatest % decrease"
            Cells(2, 15) = "Greatest % increase"
            Cells(4, 15) = "Greatest Total Volume"
            Cells(1, 9) = "Ticker"
            Cells(1, 10) = "Yearly change"
            Cells(1, 11) = "Percent Chnage"
            Cells(1, 12) = "Total Stock Volume"
    
            Columns(15).AutoFit
            Columns(16).AutoFit
            Columns(17).AutoFit
        
    Next ws
End Sub

