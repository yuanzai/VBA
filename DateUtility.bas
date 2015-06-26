Option Explicit

' ***** References required *****
' 1.  Microsoft Scripting Runtime

Const DateReferenceWorksheet = "Config"   ' Amend to the desire worksheet that as the data range
Const DateReferenceRange = "ReportDate"   ' Amend to the desire range name

'****************************************** DATE PARSING ********************************************

Public Function Parse(ByVal Target As String, ByVal DateValue As Date) As String

' *** Replaces a string containing a datetime format string and applies the desired excel date time to it.
' *** The syntax/tag used to identify the datetime format string - eg. [0:dd mmm yyy]
' *** Use "[0:" to open and "]" to close

    If Not IsDate(DateValue) Then Err.Raise 555, , "Date value does not contain a proper date"
    
    Dim i As Long
    Dim targetArray, targetSubArray As Variant
    
    targetArray = Split(Target, "[0:")
    For i = 1 To UBound(targetArray)
        targetSubArray = Split(targetArray(i), "]", 2)
        targetSubArray(0) = VBA.Format(DateValue, targetSubArray(0))
        If UBound(targetSubArray) = 0 Then Err.Raise 700, "Date Fault", "Cannot find ']'": Exit Function
        targetArray(i) = Join(targetSubArray, "")
    Next
    Parse = Join(targetArray, "")

End Function

Public Function Check(ByVal Target As String) As String
' *** Check if file exists
    
    Dim fso As New scripting.FileSystemObject
    If Not fso.FileExists(Target) Then Err.Raise 700, "File not found", Target & " cannot be found": Exit Function
    Check = Target
    
End Function

' *** Variant functions of the above 2 functions which uses excel named ranges instead

Public Function ParseWithDateRange(ByVal Target As String, _
                                   Optional ByVal DateRange As String = DateReferenceRange, _
                                   Optional ByVal WSName As String = DateReferenceWorksheet) As String
    HasRange DateRange
    ParseWithDateRange = Parse(Target, ThisWorkbook.Sheets(WSName).Range(DateRange))
End Function

Public Function ParseAndCheckWithDateRange(ByVal Target As String, _
                                   Optional ByVal DateRange As String = DateReferenceRange, _
                                   Optional ByVal WSName As String = DateReferenceWorksheet) As String
    ParseAndCheckWithDateRange = Check(ParseWithDateRange(Target, ThisWorkbook.Sheets(WSName).Range(DateRange)))
End Function

Public Function ParseRangeWithDateRange(ByVal TargetRange As String, _
                                   Optional ByVal DateRange As String = DateReferenceRange, _
                                   Optional ByVal WSName As String = DateReferenceWorksheet) As String
    HasRange TargetRange
    ParseRangeWithDateRange = ParseWithDateRange(ThisWorkbook.Sheets(WSName).Range(TargetRange), ThisWorkbook.Sheets(WSName).Range(DateRange))
End Function

Public Function ParseRangeAndCheckWithDateRange(ByVal TargetRange As String, _
                                   Optional ByVal DateRange As String = DateReferenceRange, _
                                   Optional ByVal WSName As String = DateReferenceWorksheet) As String
    ParseRangeAndCheckWithDateRange = Check(ParseRangeWithDateRange(TargetRange, ThisWorkbook.Sheets(WSName).Range(DateRange)))
End Function

Public Function ParseRangeInFolderRangeAndCheckWithDateRange(ByVal TargetRange As String, _
                                   Optional ByVal FolderRange As String = "", _
                                   Optional ByVal DateRange As String = DateReferenceRange, _
                                   Optional ByVal WSName As String = DateReferenceWorksheet) As String
Dim slash As String
slash = ""
    If FolderRange = "" Then
        ParseRangeInFolderRangeAndCheckWithDateRange = ParseRangeAndCheckWithDateRange(TargetRange, DateRange, WSName)
    Else
        HasRange FolderRange
        If Right(ThisWorkbook.Sheets(WSName).Range(FolderRange).Value, 1) <> "\" Then slash = "\"
        ParseAndCheckWithDateRange ThisWorkbook.Sheets(WSName).Range(FolderRange).Value & slash & ThisWorkbook.Sheets(WSName).Range(TargetRange).Value, DateRange
        ParseRangeInFolderRangeAndCheckWithDateRange = ParseWithDateRange(ThisWorkbook.Sheets(WSName).Range(TargetRange).Value, DateRange, WSName)
    End If
End Function
