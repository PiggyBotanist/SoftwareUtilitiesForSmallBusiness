'Project Title: Media_Inventory_tracking
'Written By: Piggy Botanist
'Date: 2023-09-25

Sub main()
    ' Set variables
    Dim received_column As Variant
    Dim used_column As Variant
    Dim uniqueValues As Variant
    Dim barcodeList As Variant
    
    ' Define worksheet we will be working with
    Set wsReceiving = ThisWorkbook.Sheets("receiving")
    Set wsMasterList = ThisWorkbook.Sheets("master_list")
    
    ' Get received & used barcode column
    received_barcode = readColumnToVariant(wsReceiving, "A", 2)
    received_counts = readCounts(wsReceiving, received_barcode, "B", 2)
    used_barcode = readColumnToVariant(wsReceiving, "C", 2)
    used_counts = readCounts(wsReceiving, used_barcode, "D", 2)

    ' Get symbol
    item_symbol = readColumnToVariant(wsMasterList, "B", 2)
    defaultUnitsPerBatch = readColumnToVariant(wsMasterList, "E", 2)
    
    itemCounts = tallyCounts(received_barcode, received_counts, used_barcode, used_counts, item_symbol, defaultUnitsPerBatch)
    
    Call displayValuesToColumn(wsMasterList, itemCounts, "k", 2)

    unique_barcode = getNotUsedBarcode(received_barcode, used_barcode)
    split_unique_barcode = splitValues(unique_barcode)
    Call displayBarcodes(wsMasterList, split_unique_barcode, item_symbol)
       
    
End Sub

' Function that reads a column and store that column into a variant
Function readColumnToVariant(ByVal ws As Worksheet, ByVal colLetter As String, ByVal startRow As Integer) As Variant
    Dim i As Integer
    Dim LastRow As Integer
    Dim columnData() As Variant
    
    ' Find the last row in the specified column
    LastRow = ws.Cells(ws.Rows.count, colLetter).End(xlUp).Row
    
    ' Check if there is any data in the column
    If LastRow >= startRow Then
        ' Redefine columnData to the length of column that contains data
        ReDim columnData(1 To (LastRow - startRow + 1))
        
        ' Loop through each cell to obtain the value
        For Each cell In ws.Range(colLetter & startRow & ":" & colLetter & LastRow)
            i = i + 1
            columnData(i) = cell.Value
        Next cell
    Else
        ' If no data, return an empty variant
        ReDim columnData(1 To 1)
        columnData(1) = Empty
    End If
    
    ' Return the variant array
    readColumnToVariant = columnData
End Function

' Function that retrieves counts based on barcode entry
Function readCounts(ByVal ws As Worksheet, ByVal reference As Variant, ByVal colLetter As String, ByVal startRow As Integer) As Variant
    Dim result As Variant
    
    ReDim result(1 To UBound(reference))
    
    For i = LBound(reference) To UBound(reference)
        If IsEmpty(ws.Cells(i + startRow - 1, colLetter)) Then
            result(i) = 0
        Else
            result(i) = ws.Cells(i + startRow - 1, colLetter)
        End If
    Next i
    
    readCounts = result
End Function

' Function that returns barcode that are received but not used in variant
Function getNotUsedBarcode(ByVal colA As Variant, ByVal colB As Variant) As Variant
    Dim result As Variant
    Dim count As Integer
    Dim temp As Integer
    
    ' If colA and colB are not empty
    If Not IsEmpty(colA(1)) And Not IsEmpty(colB(1)) Then
        ' Loop through column B
        For i = LBound(colB) To UBound(colB)
            ' Loop through column A for each colB
            For j = LBound(colA) To UBound(colA)
                ' If colA and colB values are the sample
                If colA(j) = colB(i) Then
                    ' Set the value to empty and go to next value in B
                    colA(j) = Empty
                    Exit For
                End If
            Next j
        Next i
        
        
        ' Count how many values are not empty
        count = 0
        For i = LBound(colA) To UBound(colA)
            If Not colA(i) = Empty Then
                count = count + 1
            End If
        Next i
        
        temp = 1
        ' Redefine the dimension of variant, and store all nonempty values
        ReDim result(1 To count)
        For i = LBound(colA) To UBound(colA)
            If Not colA(i) = Empty Then
                result(temp) = colA(i)
                temp = temp + 1
            End If
        Next i
    
        getNotUsedBarcode = result
        
    ' Else if colB is empty, return colA without any processing
    ElseIf Not IsEmpty(colA) Then
        getNotUsedBarcode = colA
    ' If both empty, return a variant with first value equal to empty
    Else
        ReDim result(1 To 1)
        result(1) = Empty
        getNotUsedBarcode = result
    End If

End Function

' Function that display variant values onto a column
Sub displayValuesToColumn(ByVal ws As Worksheet, data As Variant, colLetter As String, startRow As Integer)
    Dim i As Long

    If Not IsEmpty(data) Then
        ' Loop through the array and write values to the worksheet
        For i = LBound(data) To UBound(data)
            ws.Cells(startRow + i - 1, colLetter).Value = data(i)
        Next i
    End If
End Sub

' Function that returns barcodes in 3 values, symbol, LOT, expiry date
Function splitValues(arr As Variant) As Variant
    Dim result() As Variant
    
    If Not IsEmpty(arr(1)) Then
    
        ReDim result(1 To UBound(arr) + 1, 1 To 3)
        
        For i = LBound(arr) To UBound(arr)
            Dim parts() As String
            parts = Split(arr(i), " ")
            
            If UBound(parts) >= 2 Then
                result(i + 1, 1) = parts(0)
                result(i + 1, 2) = parts(1)
                result(i + 1, 3) = parts(2)
            Else
                result(i + 1, 1) = CVErr(xlErrValue)
                result(i + 1, 2) = CVErr(xlErrValue)
                result(i + 1, 3) = CVErr(xlErrValue)
            End If
        Next i
        
        splitValues = result
    Else
        ReDim result(1 To 1, 1 To 3)
        result(1, 1) = Empty
        splitValues = result
    End If
End Function

' Function that writes the barcode onto the masterlist
Sub displayBarcodes(ByVal ws As Worksheet, barcodes As Variant, media As Variant)
    Dim count As Integer
    
    Call clearColumn(ws, "H")
    Call clearColumn(ws, "I")
    Call clearColumn(ws, "J")
  
    For i = LBound(media) To UBound(media)
        For j = LBound(barcodes) To UBound(barcodes)
            If media(i) = LCase(barcodes(j, 1)) Then
                count = count + 1
                ws.Cells(1 + i, "H").Value = barcodes(j, 2)
                    
                ' Convert to date using DateSerial
                If Not barcodes(j, 3) = "" Then
                    Dim dayPart As Integer
                    Dim monthPart As Integer
                    Dim yearPart As Integer
                    Dim resultDate As Date
                    
                    ' Extract day, month, and year parts
                    yearPart = Val(Mid(barcodes(j, 3), 1, 2))
                    monthPart = Val(Mid(barcodes(j, 3), 3, 2))
                    dayPart = Val(Mid(barcodes(j, 3), 5, 2))
                
                    resultDate = DateSerial(yearPart, monthPart, dayPart)
                    ws.Cells(1 + i, "I").Value = Format(resultDate, "yyyy/mm/dd")
                End If
            End If
        Next j
        
        If Not media(i) = "" Then
            ws.Cells(1 + i, "J").Value = count
        End If
        count = 0
    Next i
End Sub


Sub clearColumn(ByVal ws As Worksheet, ByVal col As String)
    Dim LastRow As Long
    
    ' Find the last row with data in column A
    LastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' Clear the contents of column A from A2 to the last row
    ws.Range(col & "2:" & col & LastRow).ClearContents

End Sub

Function tallyCounts(received_barcode As Variant, received_counts As Variant, used_barcode As Variant, used_counts As Variant, media As Variant, defaultUnit As Variant) As Variant
    Dim result As Variant
    Dim barcode() As String
    ReDim result(1 To UBound(media))
    
    'Initialize result to 0 for all media
    For i = LBound(result) To UBound(result)
        result(i) = 0
    Next i
    
    'test
    'For i = LBound(received_counts) To UBound(received_counts)
    '    MsgBox (received_barcode(i) & " " & received_counts(i))
    'Next i
    
    received_barcode_split = splitValues(received_barcode)
    used_barcode_split = splitValues(used_barcode)
    

    For i = LBound(media) To UBound(media)
        ' Tally all entries from recieved media (addition)
        If Not IsEmpty(received_barcode(1)) Then
            For j = 1 To UBound(received_barcode)
                ' Split the barcode
                barcode = Split(received_barcode(j), " ")
                
                ' If we have the same media then...
                If LCase(barcode(0)) = LCase(media(i)) Then
                    ' Add counts if we have values, else just add default unit
                    If CInt(received_counts(j)) <> 0 Then
                        'MsgBox ("Not 0!")
                        result(i) = result(i) + received_counts(j)
                    Else
                        'MsgBox ("Is 0!")
                        result(i) = result(i) + defaultUnit(i)
                    End If
                End If
            Next j
        End If
        
        ' Tally all entries from used media (subtraction)
        If Not IsEmpty(used_barcode(1)) Then
            For j = 1 To UBound(used_barcode)
                barcode = Split(used_barcode(j), " ")
                ' If we have the same media then...
                If LCase(barcode(0)) = LCase(media(i)) Then
                    ' Subtract counts if we have values, else just add default unit
                    If CInt(used_counts(j)) <> 0 Then
                        'MsgBox ("Not 0!")
                        result(i) = result(i) - used_counts(j)
                    Else
                        'MsgBox ("Is 0!")
                        result(i) = result(i) - defaultUnit(i)
                    End If
                End If
            Next j
        End If
             
    Next i
      
    tallyCounts = result
    
End Function



