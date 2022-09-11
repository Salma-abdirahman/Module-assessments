VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
    
'Sheet - Multiple Year Stock Data - The VBA of Wall Street
    Sub WallStreetStock()

        'To create a loop through worksheets 2018-2020 declare ws as a worksheet object variable.
        Dim ws As Worksheet

        'Loop through all the worksheets in he workbook.
        For Each ws In Worksheets

            'Column Titles
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percentage Change"
            ws.Range("L1").Value = "Total Stock Volume"

            'Column Titles for Bonus.
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"

            'Declaring
            Dim Ticker_Name As String
            Dim Yearly_Open As Double
            Dim Yearly_Close As Double
            Dim Volume As Double
            Dim Row_Length As Double
            Dim Ticker_Row As Integer

            'Declaring 2nd table
            Dim Yearly_Change As Double
            Dim Percent_Change As Double





            Ticker_Row = 2

            ' Finding last row
            Row_Length = ws.Cells(Rows.Count, 1).End(xlUp).Row


            For i = 2 To Row_Length
                'Checking if <open> on row is equal to 0
                If ws.Cells(i, 3).Value = 0 Then

                    'Checking if row below the one currently on is not equal to ticker name on current row
                    If ws.Cells(i + 1).Value <> ws.Cells(i, 1).Value Then
                        'Making the ticker name equal to ticker of current row - So setting the Ticker name.
                        Ticker_Name = ws.Cells(i, 1).Value

                    End If

                    'If <open> is not equal to 0 then check:
                    'If row below ticker is the same as current row Ticker
                ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then

                    'Total volumes of the rows with the same tickers.
                    Volume = Volume + ws.Cells(i, 7).Value

                    'If the row aboves ticker is not the same as current row ticker then:
                    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                        'Year open is equal to current row open
                        Year_Open = ws.Cells(i, 3).Value
                    End If

                    'If not this happens if <open> is not equal to 0 AND row belows ticker is NOT same as current row ticker
                Else
                    'Uses all the current rows data for variables
                    Ticker_Name = ws.Cells(i, 1).Value
                    Volume = Volume + ws.Cells(i, 7).Value
                    Year_Close = ws.Cells(i, 6).Value


                    ws.Range("I" & Ticker_Row).Value = Ticker_Name
                    ws.Range("L" & Ticker_Row).Value = Volume

                    If Volume > 0 Then
                        'Yearly change = close-open
                        ws.Range("J" & Ticker_Row).Value = Year_Close - Year_Open

                        ' Changing colour
                        If ws.Range("J" & Ticker_Row).Value > 0 Then
                            'Make cell colour green if positive number
                            ws.Range("J" & Ticker_Row).Interior.ColorIndex = 4

                        Else
                            'Make cell colour red if negative number
                            ws.Range("J" & Ticker_Row).Interior.ColorIndex = 3

                        End If
                        'Percentage change = yearly change/year open
                        ws.Range("K" & Ticker_Row).Value = ws.Range("J" & Ticker_Row) / Year_Open

                        'If volume is less than 0
                    Else

                        ws.Range("J" & Ticker_Row) = 0
                        ws.Range("K" & Ticker_Row) = 0


                    End If

                    ws.Range("K" & Ticker_Row).NumberFormat = "0.00%"


                    Volume = 0
                    Ticker_Row = Ticker_Row + 1

                End If

            Next i

            'Starting the bonus.
            'Last row to calculate the Greatest %Increase/%Decrease/TotalVolume.
            Row_Length = ws.Cells(Rows.Count, 11).End(xlUp).Row

            'Creating another loop but for the bonus question.
            For i = 2 To Row_Length

                'If Percent change value of all is more than greatest % increase value then note down the ticker and value in bonus box.
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    'Note greatest % increase ticker
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                    'Note greatest % increase value
                    ws.Range("Q2").Value = ws.Range("K" & i).Value

                End If

                'If Percentage change of all is less than greatest % decrease then note down ticker and value in bonus box.
                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    'Note greatest % decrease ticker
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                    'Note greatest % decrease value.
                    ws.Range("Q3").Value = ws.Range("K" & i).Value

                End If


                'If Total stock volume value of all is more than total volume value note down ticker and value in bonus box.
                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    'Note greatest total volume ticker.
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                    'Note greatest total volume value.
                    ws.Range("Q4").Value = ws.Range("L" & i).Value

                End If

            Next i
            ' Formatting greatest percentage increase to 2 decimal places.
            ws.Range("Q2").NumberFormat = "0.00%"
            'Formatting greatest percentage decrease to 2 decimal places.
            ws.Range("Q3").NumberFormat = "0.00%"







        Next ws

    End Sub
