# VBA-Challenge
Challenge 2 (VBA)
You will be able to find  in the following lines the code for the Challenge 2 The CODE STARTS FROM LINE 11








Sub stockname6_AllWorksheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        stockname6 ws
    Next ws
End Sub

Sub stockname6(ws As Worksheet)
    Dim Ticker As String
    Dim Maxyear As Double
    Dim Minyear As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim MinyearRow As Long
    Dim MaxyearRow As Long
    Dim StockVol As Double
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim countrow As Integer
    Dim lastrow As Long
    Dim i As Long

    ' Initial variables
    Minyear = 99999999
    Maxyear = 0
    OpenPrice = 0
    ClosePrice = 0
    StockVol = 0
    MaxIncrease = 0
    MaxDecrease = 0
    MaxVolume = 0
    countrow = 2

    ' Find the last used row
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Set headers
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(1, 17) = "Value"

    ws.Columns(11).ColumnWidth = 15
    ws.Columns(12).ColumnWidth = 25
    ws.Columns(15).ColumnWidth = 25
    ws.Columns(17).ColumnWidth = 25

    For i = 2 To lastrow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            Minyear = Application.WorksheetFunction.Min(Minyear, ws.Cells(i, 2).Value)
            If Minyear = ws.Cells(i, 2).Value Then
                ' Store the row number
                MinyearRow = i
                ' Retrieve OpenPrice from the same row
                OpenPrice = ws.Cells(MinyearRow, 3).Value
            End If
        End If

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            Maxyear = Application.WorksheetFunction.Max(Maxyear, ws.Cells(i, 2).Value)
            If Maxyear = ws.Cells(i, 2).Value Then
                ' Store the row
                MaxyearRow = i
                ' Retrieve ClosePrice from the same row
                ClosePrice = ws.Cells(MaxyearRow, 6).Value
            End If
            StockVol = StockVol + ws.Cells(i, 7).Value

            ' Results to corresponding Columns
            ws.Cells(countrow, 9).Value = Ticker
            ws.Cells(countrow, 10).Value = ClosePrice - OpenPrice

            ' Check for zero OpenPrice to avoid division by zero
            If OpenPrice <> 0 Then
                ws.Cells(countrow, 11).Value = Format((ClosePrice - OpenPrice) / OpenPrice, "0.00%")
            Else
                ws.Cells(countrow, 11).Value = 0
            End If

            ws.Cells(countrow, 12).Value = StockVol
            
            'Assign Colors to column
            If ws.Cells(countrow, 10).Value < 0 Then
                ws.Cells(countrow, 10).Interior.Color = RGB(255, 0, 0) ' Red
            ElseIf ws.Cells(countrow, 10).Value >= 0 Then
                ws.Cells(countrow, 10).Interior.Color = RGB(144, 238, 144) ' Light Green
            End If

            '  Maximum increase, decrease, and volume
            If ws.Cells(countrow, 11).Value > MaxIncrease Then
                MaxIncrease = ws.Cells(countrow, 11).Value
                ws.Cells(2, 16).Value = Ticker
                ws.Cells(2, 17).Value = MaxIncrease
            End If

            If ws.Cells(countrow, 11).Value < MaxDecrease Then
                MaxDecrease = ws.Cells(countrow, 11).Value
                ws.Cells(3, 16).Value = Ticker
                ws.Cells(3, 17).Value = MaxDecrease
            End If

            If ws.Cells(countrow, 12).Value > MaxVolume Then
                MaxVolume = ws.Cells(countrow, 12).Value
                ws.Cells(4, 16).Value = Ticker
                ws.Cells(4, 17).Value = MaxVolume
            End If

            ' Move to the next row for output
            countrow = countrow + 1

            ' Reset variables
            Minyear = 99999999
            Maxyear = 0
            OpenPrice = 0
            ClosePrice = 0
            StockVol = 0
        Else
            StockVol = StockVol + ws.Cells(i, 7).Value
        End If
    Next i
End Sub


![Result 2019](https://github.com/jrobertofg/VBA-Challenge/assets/27827806/19dd0542-0f49-42d9-a4c5-414069b35e65)
![Result 2018](https://github.com/jrobertofg/VBA-Challenge/assets/27827806/9bb253ec-256b-4239-aaf4-577cc16fc31c)
![Result 2020](https://github.com/jrobertofg/VBA-Challenge/assets/27827806/de549dd5-3666-4f11-a327-d1d05821f341)

