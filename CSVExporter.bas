Option Explicit


Sub csvExporter(ByVal sht As Worksheet, ByVal path As String, ByVal header As Boolean, ParamArray forceText() As Variant)
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim startRow As Long
    Dim i As Long, j As Long, k As Long
    Dim fso As New FileSystemObject
    Dim csvFile As TextStream
    Dim colTypeArray() As String
    Dim dataArray As Variant
    Dim line As String
    
    lastRow = sht.UsedRange.Rows.count
    lastCol = sht.UsedRange.columns.count
    
    If UBound(forceText) > 0 Then
        For k = LBound(forceText) To UBound(forceText)
            If Not IsNumeric(forceText(k)) Then Err.Raise 555, "force text param array not numeric"
            If forceText(k) > lastCol Then Err.Raise 555, "force text param array item greater than sheet columns"
        Next
    End If
    
    ReDim colTypeArray(1 To lastCol)
    If header Then startRow = 2 Else startRow = 1
    For j = 1 To lastCol
        If UBound(forceText) > 0 Then
            For k = LBound(forceText) To UBound(forceText)
                If forceText(k) = j Then
                colTypeArray(j) = "String"
                GoTo Continue
                End If
            Next
        End If
        
        For i = startRow To lastRow
            If i = 2 And sht.Cells(i, j).NumberFormat = "d-mmm-yy" Then colTypeArray(j) = "Date"
            
            If IsError(sht.Cells(i, j).Value) Then
                sht.Cells(i, j).ClearContents
            ElseIf IsNumeric(sht.Cells(i, j).Value) Then
                If colTypeArray(j) <> "Dbl" And colTypeArray(j) <> "Date" Then colTypeArray(j) = "Int"
                If InStr(1, sht.Cells(i, j).Text, ".", vbTextCompare) Then colTypeArray(j) = "Dbl"
            Else
                If colTypeArray(j) = "Date" And (IsDate(sht.Cells(i, j).Value) Or sht.Cells(i, j).Value = "") Then
                Else
                    colTypeArray(j) = "String"
                    Exit For
                End If
            End If
        Next
        
Continue:
    Next
    
    dataArray = sht.UsedRange

    Set csvFile = fso.CreateTextFile(path, True)
    
    
    For i = 1 To lastRow
    line = ""
        For j = 1 To lastCol
            If header And i = 1 Then
                If j = 1 Then line = """" & dataArray(i, j) & """"
                If j > 1 Then line = line & ",""" & dataArray(i, j) & """"
            Else
                If colTypeArray(j) = "String" Then
    
                    If j = 1 Then line = """" & CStr(dataArray(i, j)) & """"
                    If j > 1 Then line = line & ",""" & CStr(dataArray(i, j)) & """"
                ElseIf colTypeArray(j) = "Int" Then
                    If j = 1 Then line = dataArray(i, j)
                    If j > 1 Then line = line & "," & dataArray(i, j)
                ElseIf colTypeArray(j) = "Dbl" Then
                    If dataArray(i, j) = 0 Then dataArray(i, j) = "0.00"
                    If j = 1 Then line = dataArray(i, j)
                    If j > 1 Then line = line & "," & dataArray(i, j)
                ElseIf colTypeArray(j) = "Date" Then
                    If dataArray(i, j) <> "" Then dataArray(i, j) = "#" & dataArray(i, j) & "#"
                    If j = 1 Then line = dataArray(i, j)
                    If j > 1 Then line = line & "," & dataArray(i, j)
                End If
            End If

        Next
        
    csvFile.WriteLine line
    Next
    
    csvFile.Close
    Set csvFile = Nothing

End Sub
