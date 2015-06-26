Option Explicit
' ***** References required *****
' 1.  Microsoft Scripting Runtime
' 2.  Microsoft ActiveX Data Objects 2.8 Library
' 3.  Microsoft ActiveX Data Recordset Objects 2.8 Library

Public Sub clearTillRange(ByVal rng As Range, ByVal rngEnd As Range, ByVal columnCount As Integer, ByVal rowCount As Integer)
' Deletes the row inbetween 2 ranges on the same column by a determined number of columns
' Reinserts a determined number of rows of the determined number of columns

' Useful for pasting in dynamically generated data which has different number of rows on each run while
' maining the position of the table in the worksheet

    Dim i As Integer, k As Integer
    k = 1
    i = 0
    If rng.Row + 1 < rngEnd.Row Then Range(rng.Offset(1, 0), rngEnd.Offset(-1, columnCount - 1)).Delete xlUp
    If rowCount > 1 Then
        Range(rng.Offset(1, 0), rng.Offset(rowCount - k, columnCount - 1)).Insert xlDown
    ElseIf rowCount = 0 Then
        Range(rng.Offset(1, 0), rngEnd.Offset(-1, columnCount - 1)).ClearContents
    End If
    
    If rowCount = 1 Then Exit Sub
    i = 0
    While (i < columnCount)
        If rng.Offset(0, i).HasFormula Then
            With Range(rng.Offset(1, i), rng.Offset(rowCount - 1, i))
                .FormulaR1C1 = rng.Offset(0, i).FormulaR1C1
            End With
        End If
        i = i + 1
    Wend
End Sub


Public Sub PasteRecordsetHeader(record As adodb.Recordset, rng As Range)
    Dim fld As adodb.field
    Dim i As Long
    Dim pasteArray As Variant
    ReDim pasteArray(0 To 0, 0 To record.fields.count - 1)
    i = 0
    For Each fld In record.fields
        pasteArray(0, i) = fld.Name
        'rng.Offset(0, i).Value = fld.Name
        i = i + 1
    Next
    Range(rng, rng.Offset(0, record.fields.count - 1)).Value = pasteArray
End Sub


Public Sub FormulaCopyDown(ws As Worksheet, Optional refRow As Long = 2)
    Dim i As Long
    For i = 1 To ws.UsedRange.columns.count
        If ws.Cells(refRow, i).HasFormula Then ws.Cells(refRow, i).Copy ws.Range(ws.Cells(refRow, i), ws.Cells(ws.UsedRange.Rows.count, i))
    Next
End Sub


Public Sub SmartPasteRecordset(record As adodb.Recordset, ByVal rng As Range, ByVal rngEnd As Range)
    Dim i As Integer, k As Integer
    Dim formulaArray As scripting.Dictionary
    Dim rowCount As Long, columnCount As Long

    rowCount = record.RecordCount
    columnCount = record.fields.count

    'Determine formula cells for repopulation later
    i = 0
    While (i < rowCount)
        If rng.Offset(0, i).HasFormula Then
            formulaArray.Add i, rng.Offset(0, i).FormulaR1C1
        End If
        i = i + 1
    Wend

    k = 1
    i = 0
    If rng.Row + 1 < rngEnd.Row Then Range(rng.Offset(1, 0), rngEnd.Offset(-1, columnCount - 1)).Delete xlUp
    If rowCount > 1 Then
        Range(rng.Offset(1, 0), rng.Offset(rowCount - k, columnCount - 1)).Insert xlDown
    Else
        Range(rng.Offset(1, 0), rngEnd.Offset(-1, columnCount - 1)).ClearContents
    End If
    
    For Each i In formulaArray.Keys
        rng.Offset(0, i).FormulaR1C1 = formulaArray.Item(i)
        With Range(rng.Offset(1, i), rng.Offset(rowCount - 1, i))
            .FormulaR1C1 = rng.Offset(0, i).FormulaR1C1
        End With
    Next
    
End Sub


Public Sub RecordSetToSheetHeaderCheck(record As adodb.Recordset, ws As Worksheet)
    Dim fld As field
    Dim i As Long
    
    i = 1
    For Each fld In record.fields
        If ws.Cells(1, i).Value <> fld.Name Then Err.Raise 800, "File header does not match " & ws.Name & " sheet header", "File header does not match " & ws.Name & " sheet header"
        i = i + 1
    Next

End Sub


