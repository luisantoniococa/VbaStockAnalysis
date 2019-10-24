Sub stock2():
    For Each ws In Worksheets
        Dim ticker As String
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim stockvolume As Long
        Dim j As Long
        Dim openflag As Boolean
        Dim openday, closeday As Double
        openflag = True
        j = 2
        stockvolume = 0
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total stock Volume"
        For i = 2 To Application.CountA(ws.Range("A:A"))
            ticker = ws.Cells(i, 1).Value
            If openflag = True Then
                openday = ws.Cells(i, 3).Value
                openflag = False
            End If

            If ticker = ws.Cells(i, 1).Value And ws.Cells(i + 1, 1).Value = ticker Then
                'stockvolume = stockvolume + CLng(ws.Cells(i, 7).Value)
                ws.Cells(j, 13).Value = ws.Cells(j, 13).Value + ws.Cells(i, 7).Value
                
            Else
                'stockvolume = stockvolume + CLng(ws.Cells(i, 7))
                ws.Cells(j, 13).Value = ws.Cells(j, 13).Value + ws.Cells(i, 7).Value
                'MsgBox (Cells(j, 13))
                closeday = ws.Cells(i, 6).Value
                yearlychange = closeday - openday
                If openday <> 0 Then
                    percentchange = yearlychange / openday
                End If
                ws.Cells(j, 10).Value = ticker
                ws.Cells(j, 11).Value = yearlychange
                ws.Cells(j, 12).Value = percentchange
                If yearlychange > 0 Then
                        ws.Cells(j, 12).Style = "Percent"
                        ws.Cells(j, 11).NumberFormat = "0.00000"
                        ws.Cells(j, 11).Interior.ColorIndex = 4
                        ws.Cells(j, 11).Font.Color = 1
                        ws.Cells(j, 12).Interior.ColorIndex = 4
                        ws.Cells(j, 12).Font.Color = 1
                ElseIf yearlychange < 0 Then
                        ws.Cells(j, 12).Style = "Percent"
                        ws.Cells(j, 11).NumberFormat = "0.00000"
                        ws.Cells(j, 11).Interior.ColorIndex = 3
                        ws.Cells(j, 11).Font.Color = 1
                        ws.Cells(j, 12).Interior.ColorIndex = 3
                        ws.Cells(j, 12).Font.Color = 1
                        
                End If

                'ws.Cells(j, 13).Value = stockvolume
                stockvolume = 0
                j = j + 1
                openflag = True
            End If
        Next i
        
        ws.Range("n2").Value = "Greatest Percent Increase"
        ws.Range("n3").Value = "Greatest Percent Decrease"
        ws.Range("n4").Value = "Greatest stock Volume"
        ws.Range("o1").Value = "Value"
        ws.Range("p1").Value = "Ticker"

        For Row = 2 To Application.CountA(ws.Range("J:J"))
            If ws.Cells(Row, 12).Value > ws.Cells(2, 15).Value Then
                ws.Cells(2, 15).Value = ws.Cells(Row, 12).Value
                ws.Cells(2, 16).Value = ws.Cells(Row, 10).Value
            End If
            If ws.Cells(Row, 12).Value < ws.Cells(3, 15).Value Then
                ws.Cells(3, 15).Value = ws.Cells(Row, 12).Value
                ws.Cells(3, 16).Value = ws.Cells(Row, 10).Value
            End If
            If ws.Cells(Row, 13).Value > ws.Cells(4, 15).Value Then
                ws.Cells(4, 15).Value = ws.Cells(Row, 13).Value
                ws.Cells(4, 16).Value = ws.Cells(Row, 10).Value
            End If
        Next Row
        
       ws.Columns("A:P").AutoFit
       ws.Range("o2:o3").Style = "Percent"
    Next ws
End Sub
