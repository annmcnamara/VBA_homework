Attribute VB_Name = "Module1"
  Sub countTicker()
    Dim nRows As Long
    Dim ticker As String
    Dim TotalSTockValue As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim r As Long
    Dim presentationRow As Long
    
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    
    Set starting_ws = ActiveSheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate

        '   nRows = Range("A1").End(xlDown).Row
        nRows = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        presentationRow = 2
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
              
        Cells(2, 15) = "Greatest % Increase"
        Cells(3, 15) = "Greatest % Decrease"
        Cells(4, 15) = "Greatest Total Volume"
            
        Cells(1, 16) = "Ticker"
        Cells(1, 17) = "Value"

        ticker = Cells(2, 1)
        openingPrice = Cells(2, 3).Value
        'totalStockVolume = Cells(2, 7)
    
        For r = 2 To nRows
          If (Cells(r + 1, 1) = ticker) Then
            totalStockVolume = totalStockVolume + Cells(r + 1, 7)
    
         Else:
            closingPrice = Cells(r, 6)
            yearlyChange = closingPrice - openingPrice
            
            PercentageChange = 0
            If (openingPrice <> 0) Then
                PercentageChange = yearlyChange / openingPrice
            End If
            Cells(presentationRow, 9) = ticker
            Cells(presentationRow, 10) = yearlyChange
            Cells(presentationRow, 11) = PercentageChange
            Cells(presentationRow, 12) = totalStockVolume

            If (yearlyChange <= 0) Then
                Cells(presentationRow, 11).Interior.ColorIndex = 3
            Else
                Cells(presentationRow, 11).Interior.ColorIndex = 4
            End If
              
            ticker = Cells((r + 1), 1)
            presentationRow = presentationRow + 1
            
            openingPrice = Cells(r + 1, 3).Value 'update opening price
            totalStockVolume = Cells(r + 1, 7)
            yearlyChange = 0
            
        End If
     
    Next r
    
    Call Summary
    

    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
 Next ws
    
End Sub
Sub Summary()
    Dim maxInc As Double
    Dim maxDec As Double
    Dim GTV As Double
  
    Dim maxIncTKR As String
    Dim maxDecTKR As String
    Dim GTVTKR As String

    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    
    Dim PercentageChange As Double
    Dim totalStockVolume As Double
    
    Set starting_ws = ActiveSheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
    
        maxInc = 0
        maxDec = 0
        GTV = 0
        
        Dim nRowsSummary As Long
        nRowsSummary = Cells(Rows.Count, 9).End(xlUp).Row

        'loop over summary values to get further summary values
        For r = 2 To nRowsSummary
            PercentageChange = Cells(r, 11)
            totalStockVolume = Cells(r, 12)
            
            ticker = Cells(r, 9)
    
            If (PercentageChange > maxInc) Then
                maxInc = PercentageChange
                maxIncTKR = ticker
            End If

            If (PercentageChange < maxDec) Then
                maxDec = PercentageChange
                maxDecTKR = ticker
            End If

            If (totalStockVolume > GTV) Then
                GTV = totalStockVolume
                GTVTKR = ticker
            End If

        Next r

    Cells(2, 16).Value = maxIncTKR
    Cells(2, 17).Value = maxInc

    Cells(3, 16).Value = maxDecTKR
    Cells(3, 17).Value = maxDec

    Cells(4, 16).Value = GTVTKR
    Cells(4, 17).Value = GTV

    
    Next ws
    
    
End Sub

