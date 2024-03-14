Attribute VB_Name = "Module1"
Sub macro1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Assignment")
    Dim stationA As Integer
    Dim stationB As Integer
    Dim totalA As Integer
    Dim totalB As Integer
    Dim lastRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim temp As Integer

    
    totalA = 0
    totalB = 0
    stationA = 0
    stationB = 0
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    Dim timeArray() As Integer
    ReDim timeArray(1 To lastRow - 1, 1 To 2)
    


    For i = 2 To lastRow
        timeArray(i - 1, 1) = ws.Cells(i, "C").Value
        timeArray(i - 1, 2) = i
    Next i
    

  
    For i = LBound(timeArray) To UBound(timeArray) - 1
        For j = i + 1 To UBound(timeArray)
            If timeArray(i, 1) < timeArray(j, 1) Then
                temp = timeArray(i, 1)
                timeArray(i, 1) = timeArray(j, 1)
                timeArray(j, 1) = temp
                temp = timeArray(i, 2)
                timeArray(i, 2) = timeArray(j, 2)
                timeArray(j, 2) = temp
            End If
        Next j
    Next i

    For i = LBound(timeArray) To UBound(timeArray)
        If totalA <= totalB Then
            ws.Cells(timeArray(i, 2), "E").Value = "A"
            ws.Cells(timeArray(i, 2), "F").Value = stationA + 1
            totalA = totalA + timeArray(i, 1)
            stationA = stationA + 1
        Else
            ws.Cells(timeArray(i, 2), "E").Value = "B"
            ws.Cells(timeArray(i, 2), "F").Value = stationB + 1
            totalB = totalB + timeArray(i, 1)
            stationB = stationB + 1
        End If
    Next i
End Sub

