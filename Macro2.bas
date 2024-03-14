Attribute VB_Name = "Module2"
Sub Macro2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Assignment")
    Dim lastRow As Integer
    Dim currentTime As Integer
    Dim i As Integer
    Dim order As Integer
    Dim station As String
    Dim timeA As Integer
    Dim timeB As Integer
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    Dim orderA() As Integer
    Dim orderB() As Integer
    ReDim orderA(1 To lastRow - 1)
    ReDim orderB(1 To lastRow - 1)
       
    ws.Range("H2:H" & lastRow).ClearContents
    currentTime = ws.Range("K5").Value
    
    For i = 2 To lastRow
        station = ws.Cells(i, "E").Value
        order = ws.Cells(i, "F").Value
        
        If station = "A" Then
            orderA(order) = i
        Else
            orderB(order) = i
        End If
    Next i
    
    TaskStatus ws, orderA, timeA, currentTime
    TaskStatus ws, orderB, timeB, currentTime
    Call CurrentAndNext(ws)


End Sub

Sub TaskStatus(ws As Worksheet, ByRef orderArray() As Integer, ByRef timeAccumulated As Integer, currentTime As Integer)
    Dim idx As Integer
    Dim row As Integer
    Dim TaskTime As Integer, dueTime As Integer
    Dim status As String
    Dim finishingTime As Integer
    
    For idx = LBound(orderArray) To UBound(orderArray)
        If orderArray(idx) = 0 Then Exit For
        
        row = orderArray(idx)
        TaskTime = ws.Cells(row, "C").Value
        dueTime = ws.Cells(row, "D").Value
        
        finishingTime = timeAccumulated + TaskTime
        
       
        If finishingTime <= currentTime Then
            If finishingTime <= dueTime Then
                status = "Finished in Time"
            Else
                status = "Finished Late"
            End If
           
            ws.Cells(row, "H").Value = finishingTime
        ElseIf timeAccumulated < currentTime And finishingTime > currentTime Then
            status = "In Process"
        Else
            status = "Waiting"
        End If
        
        
        ws.Cells(row, "G").Value = status
        
        timeAccumulated = finishingTime
    Next idx
End Sub
Sub CurrentAndNext(ws As Worksheet)
    Dim lastRow As Integer
    Dim nextOrderA As Integer
    Dim nextOrderB As Integer
    Dim minNextOrderA As Integer
    Dim minNextOrderB As Integer
    Dim taskNameA As String
    Dim taskNameB As String
    Dim nextTaskNameA As String
    Dim nextTaskNameB As String
    Dim currentOrderA As Integer
    Dim currentOrderB As Integer
    Dim station As String
    Dim status As String
    Dim order As Integer
    Dim i As Integer

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    minNextOrderA = lastRow + 2  ' a control mechanism to check weather tasks in stationA are completed or not
    minNextOrderB = lastRow + 2  ' a control mechanism to check weather tasks in stationB are completed or not

    For i = 2 To lastRow

        station = ws.Cells(i, "E").Value
        status = ws.Cells(i, "G").Value
        order = ws.Cells(i, "F").Value

        If status = "In Process" Then
            If station = "A" And order > currentOrderA Then
                currentOrderA = order
                taskNameA = ws.Cells(i, "A").Value
            ElseIf station = "B" And order > currentOrderB Then
                currentOrderB = order
                taskNameB = ws.Cells(i, "A").Value
            End If
        End If
    Next i

    For i = 2 To lastRow

        station = ws.Cells(i, "E").Value
        status = ws.Cells(i, "G").Value
        order = ws.Cells(i, "F").Value

        If status = "Waiting" Then
            If station = "A" And order > currentOrderA And order < minNextOrderA Then
                minNextOrderA = order
                nextTaskNameA = ws.Cells(i, "A").Value
            ElseIf station = "B" And order > currentOrderB And order < minNextOrderB Then
                minNextOrderB = order
                nextTaskNameB = ws.Cells(i, "A").Value
            End If
        End If
    Next i
    
    If taskNameA = "" Then
        ws.Range("K13").Value = " All Tasks are finished"
    Else
        ws.Range("K13").Value = taskNameA
    End If
    If nextTaskNameA = "" Then
        ws.Range("L13").Value = "All tasks are finished"
    Else
        ws.Range("L13").Value = nextTaskNameA
    End If
    
    If taskNameB = "" Then
        ws.Range("K14").Value = " All Tasks are finished"
    Else
        ws.Range("K14").Value = taskNameB
    End If
    If nextTaskNameB = "" Then
        ws.Range("L14").Value = "All tasks are finished"
    Else
        ws.Range("L14").Value = nextTaskNameB
    End If
        
    
   
End Sub
