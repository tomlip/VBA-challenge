' VBA code by Thomas Lippoli
Function tickerArray(ws As Worksheet)

  On Error Resume Next

  Dim rCnt As Double
  rCnt = ws.Cells(Rows.Count,1).End(xlUp).Row
  
  'array to hold results
  Dim arr() As Variant
  ReDim arr(1 To 4, 1 To 1) As Variant

  'initialize variables for result array
  Dim tickStr, newTick As String
  tickStr = ws.Cells(2,1).Value 'set 1st ticker value
  Dim valueChange, percentChange, volume As Double
  volume = 0 'start at 0

  'additional variables to support the above
  Dim openVal, closeVal, tickVol As Double
  openVal = ws.Cells(2,3).Value
  
  'array index

  'loop through ticker rows and create table
  'rCnt - 1 as we are skipping the 1st row
  For i = 1 To rCnt -1
    newTick = ws.Cells(i+1,1).Value
    tickVol = ws.Cells(i+1,7).Value
    If tickStr = newTick Then

      'increase stock count
      volume = volume + tickVol

    'last row catch
    ElseIf i = rCnt - 1 Then
      volume = volume + tickVol
      closeVal = ws.Cells(i,6).Value
      valueChange =  closeVal - openVal
      percentChange = 100 * (valueChange/openVal)
      'add to results array befor preparing for the next ticker
      n = UBound(arr,2)
      arr(1,n) = tickStr
      arr(2,n) = valueChange
      arr(3,n) = percentChange
      arr(4,n) = volume

    Else
      closeVal = ws.Cells(i,6).Value
      valueChange =  closeVal - openVal
      percentChange = 100 * (valueChange/openVal)
      'add to results array befor preparing for the next ticker
      n = UBound(arr,2)
      arr(1,n) = tickStr
      arr(2,n) = valueChange
      arr(3,n) = percentChange
      arr(4,n) = volume

      'add new array column for new ticker
      ReDim Preserve arr(1 To 4, 1 To n + 1)

      'reset values for new ticker
      openVal = ws.Cells(i+1,3).Value
      volume = tickVol
      tickStr = newTick

    End If
  Next

'add new array column for new ticker
      ReDim Preserve arr(1 To 4, 1 To n)

  'variables for largest numbers
  Dim gInc, gDec, gVol As Double
  Dim incStr, decStr, volStr As String
  gInc = 0.0
  gDec = 0.0
  gVol = 0.0

  For j = 1 To rCnt - 1
    'write result array onto ws
    For i = 1 To 4
      ws.Cells(j+1,8+i).Value = arr(i,j)
    Next
    'largest numbers check
    'greatest increase in %
    If arr(3,j) > gInc Then
      gInc = arr(3,j)
      incStr = arr(1,j)
    End If
    'greatest decrease in %
    If arr(3,j) < gDec Then
      gDec = arr(3,j)
      decStr = arr(1,j)
    End If
    'greatest volume
    If arr(4,j) > gVol Then
      gVol = arr(4,j)
      volStr = arr(1,j)
    End If
    

    'conditional formatting
    If arr(2, j) < 0 Then
      ws.Cells(j+1,10).Interior.ColorIndex = 3
    Else
      ws.Cells(j+1,10).Interior.ColorIndex = 4
    End If

  Next

  'write largest numbers to worksheet
  ws.Range("P2").Value = incStr
  ws.Range("P3").Value = decStr
  ws.Range("P4").Value = volStr
  ws.Range("Q2").Value = gInc
  ws.Range("Q3").Value = gDec
  ws.Range("Q4").Value = gVol

End Function

Sub VBA_Stocks()

  'stop screen flickering
  Application.ScreenUpdating = False
  
  'initialize variables for worksheet looping
  Dim wb As Workbook
  Dim ws As Worksheet
  Set wb = ThisWorkbook
  
  'loop for all sheets in worksheet collection
  For Each ws In wb.Worksheets
    ws.Range("I1") ="Ticker"
    ws.Range("J1") ="Yearly Change"
    ws.Range("K1") ="Percent Change"
    ws.Range("L1") ="Total Stock Volume"
    ws.Range("O2") ="Greatest % Increase"
    ws.Range("O3") ="Greatest % Decrease"
    ws.Range("O4") ="Greatest Total Volume"
    ws.Range("P1") ="Ticker"
    ws.Range("Q1") ="Value"
    Call tickerArray(ws)
  Next ws
    
  'turn it back on for changes
  Application.ScreenUpdating = True

End Sub