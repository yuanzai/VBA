Option Explicit

'****************************************** Array Join/Split utility ********************************************

Public Function JoinText(lineArray As Variant, ByVal delimiter As String)
' Joins elements of an array with a delimiter as well as applying escaped double quotes chars to all elements
' such that it prints as "element" in a text file

    Dim i As Long
    For i = LBound(lineArray) To UBound(lineArray)
        If Left(lineArray(i), 1) <> """" And Right(lineArray(i), 1) <> """" Then lineArray(i) = """" & lineArray(i) & """"
        
    Next
    If delimiter = "" Then delimiter = ","
    JoinText = Join(lineArray, delimiter)
End Function


Public Function JoinWithNumeric(lineArray As Variant, ByVal delimiter As String, ParamArray aVar() As Variant)
' Joins elements of an array with a delimiter as well as applying escaped double quotes chars to all elements
' such that it prints as "element" in a text file. The paramarray takes in the elements which should not be
' within escaped double quotes

    Dim k As Variant
    Dim i As Long

    For i = LBound(lineArray) To UBound(lineArray)
        If Left(lineArray(i), 1) <> """" And Right(lineArray(i), 1) <> """" Then lineArray(i) = """" & lineArray(i) & """"
    Next
    For Each k In aVar
        lineArray(k) = Mid(lineArray(k), 2, Len(lineArray(k)) - 2)
    Next
    If delimiter = "" Then delimiter = ","
    JoinWithNumeric = Join(lineArray, delimiter)
End Function


Public Function SplitWithTextQualifier(ByVal inputString As String, _
                                       ByVal delimiter As String, _
                                       ByVal textQualifier As String, _
                                       Optional ByVal removeQualifiers As Boolean = False) As Variant
' Splits a string which may contain elements that has a text qualifier such as a double quote. Option to
' retain the qualifier in the element if needed.

    Dim resultArray As Variant
    Dim i As Long, elementCount As Long, startPosition As Long
    Dim insideTQ As Boolean
    ReDim resultArray(0 To 0)
    insideTQ = False
    startPosition = 1
    elementCount = 0
    
    For i = 1 To Len(inputString) + 1
        If ((Mid(inputString, i, 1)) = delimiter And Not insideTQ) Or i = Len(inputString) + 1 Then
            ReDim Preserve resultArray(0 To elementCount)
            
            If removeQualifiers Then
                If Left(Mid(inputString, startPosition, i - startPosition), 1) = textQualifier And Right(Mid(inputString, startPosition, i - startPosition), 1) = textQualifier Then
                    resultArray(elementCount) = Mid(inputString, startPosition + 1, i - startPosition - 2)
                Else
                    resultArray(elementCount) = Mid(inputString, startPosition, i - startPosition)
                End If
            Else
                resultArray(elementCount) = Mid(inputString, startPosition, i - startPosition)
            End If
            startPosition = i + 1
            elementCount = elementCount + 1
        ElseIf (Mid(inputString, i, 1)) = textQualifier Then
            insideTQ = Not insideTQ
        End If
    Next
    
    SplitWithTextQualifier = resultArray
End Function

