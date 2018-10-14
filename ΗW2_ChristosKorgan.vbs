Sub process2014()
   Call Process(2014)
End Sub
Sub process2015()
   Call Process(2015)
End Sub
Sub process2016()
   Call Process(2016)
End Sub
Sub Process(ByVal Year As String)
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.StatusBar = True
   Dim WsData, WsRpt As Worksheet
   Dim Row, nRows, RowRpt As Long
   Dim ActTicker, LastTicker As String
   Dim OpenPrice, ClosePrice As Double
   Set WsData = ThisWorkbook.Sheets(Year)
   Set WsRpt = ThisWorkbook.Sheets("Report-" & Year)
   Row = GWsLastRowInCol(WsRpt, "A")
   If Row > 1 Then
      WsRpt.Range(WsRpt.Cells(2, "A"), WsRpt.Cells(Row, "F")).ClearContents
   End If
   nRows = GWsLastRowInCol(WsData, "A")
   LastTicker = ""
   RowRpt = 1
   For Row = 2 To nRows
      If LastTicker <> WsData.Cells(Row, "A") Then
         Application.StatusBar = "Processing... " & Round(Row / nRows * 100, 0) & "%..."
         If LastTicker <> "" Then
            WsRpt.Cells(RowRpt, "C") = WsData.Cells(Row - 1, "C")
         End If
         RowRpt = RowRpt + 1
         WsRpt.Cells(RowRpt, "A") = WsData.Cells(Row, "A")
         WsRpt.Cells(RowRpt, "B") = WsData.Cells(Row, "C")
         WsRpt.Cells(RowRpt, "D").Formula = "=C" & RowRpt & "-B" & RowRpt
         WsRpt.Cells(RowRpt, "E").Formula = "=D" & RowRpt & "/B" & RowRpt
         WsRpt.Cells(RowRpt, "F") = 0
      End If
      LastTicker = WsData.Cells(Row, "A")
      WsRpt.Cells(RowRpt, "F") = WsRpt.Cells(RowRpt, "F") + WsData.Cells(Row, "G")
   Next
   WsRpt.Cells(RowRpt, "C") = WsData.Cells(Row - 1, "C")
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.StatusBar = True
End Sub
Function GWsLastRowInCol(ByVal Ws As Worksheet, ByVal NCol As Variant) As Long
   If IsNumeric(NCol) Then
      NCol = GWsColumnLetter(NCol)
   End If
   GWsLastRowInCol = Ws.Cells(Rows.Count, NCol).End(xlUp).Row
End Function
Function GWsColumnLetter(ByVal nColumn As Long) As String
   GWsColumnLetter = Split(Cells(1, nColumn).Address, "$")(1)
End Function
