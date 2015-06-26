Option Explicit
' ***** References required *****
' 1.  Microsoft Scripting Runtime
' 2.  Microsoft Outlook 15.0 Object Library (older versions should be fine)

'************ Utility to create emails and create images of excel ranges to attach to email. **************

Sub NewEmail(ByRef wb As Workbook, ByVal subject As String, ByVal fromAddress As String, ByVal toAddress As String, ByVal ccAddress As String, ByVal body As String, ByVal atmtArray As Variant)
    Dim imgDict As scripting.Dictionary
    Dim imgPath As Variant
    Set imgDict = ParseImageRange(wb, body)
    Call SendEmail(subject, "", toAddress, ccAddress, imgDict.Item("HTMLBody"), atmtArray, imgDict.Item("Attachments"))
    If UBound(imgDict.Item("Attachments")) > 0 Then
        For Each imgPath In imgDict.Item("Attachments")
            Kill imgPath
        Next
    End If
End Sub


Private Function ParseImageRange(ByRef wb As Workbook, htmlBody As String) As scripting.Dictionary
    Set ParseImageRange = New scripting.Dictionary
    Dim fso As New scripting.FileSystemObject
    
    Dim i As Long
    Dim targetArray, targetSubArray As Variant
    Dim imgRngArray As Variant
    Dim fullrng As Variant
    Dim shtname As String
    Dim rngname As String
    Dim path As String
    
    
    targetArray = Split(htmlBody, "[imgrng:")
    For i = 1 To UBound(targetArray)
        targetSubArray = Split(targetArray(i), "]")
        fullrng = Split(Replace(targetSubArray(0), "'", ""), "!")
        If UBound(fullrng) <> 1 Then
            Err.Raise 700, "Img Range Fault"
        Else
            shtname = fullrng(0)
            rngname = fullrng(1)
        End If
        path = createJpg(wb.Sheets(shtname).Range(rngname))
        If i = 1 Then ReDim imgRngArray(0 To 0)
        ReDim Preserve imgRngArray(0 To i - 1)
        imgRngArray(i - 1) = path
        targetSubArray(0) = "<img src='cid:" & fso.GetFileName(path) & "'>"
        If UBound(targetSubArray) = 0 Then Err.Raise 700, "Img Range Fault", "Cannot find ']'": Exit Function
        targetArray(i) = Join(targetSubArray, "")
    Next
    ParseImageRange.Add "HTMLBody", Join(targetArray, "")
    ParseImageRange.Add "Attachments", imgRngArray
    
End Function


Private Sub SendEmail(ByVal subject As String, ByVal fromAddress As String, ByVal toAddress As String, ByVal ccAddress As String, ByVal body As String, ByVal atmtArray As Variant, ByVal hiddenAtmt As Variant)
    Dim myOlApp As Outlook.Application
    Dim MItem As Outlook.MailItem
    Dim atmt As Variant

    'Get Outlook Application
    On Error Resume Next
    Set myOlApp = GetObject(, "Outlook.Application")
    If myOlApp Is Nothing Then
        Set myOlApp = CreateObject("Outlook.Application")
        Debug.Print "Outlook Started"
        
        If myOlApp Is Nothing Then
            Debug.Print "Unable to start Outlook."
            Exit Sub
        End If
    End If
               
'Create Mail Item and send it
    Set MItem = myOlApp.CreateItem(0)
    With MItem
    
        If Not IsEmpty(atmtArray) Then
            For Each atmt In atmtArray
                .Attachments.Add atmt
            Next
        End If
        
        If Not IsEmpty(hiddenAtmt) Then
            For Each atmt In hiddenAtmt
                .Attachments.Add atmt, Outlook.olByValue, 0
            Next
        End If
        
        .SentOnBehalfOfName = fromAddress
        .To = toAddress
        .CC = ccAddress
        .subject = subject
        .htmlBody = body
        .Display
    End With
End Sub


Private Function createJpg(rng As Range) As String
    Dim path As String
    Dim imgType As String
    imgType = "PNG"
    rng.Worksheet.Activate
    ActiveWindow.DisplayGridlines = False

    path = VBA.Environ$("temp") & "\temp" & CStr(CLng(Rnd() * 100000)) & CStr(CLng(Rnd() * 100000)) & "." & imgType
    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    With rng.Worksheet.ChartObjects.Add(rng.Left, rng.Top, rng.Width, rng.Height)
        .Activate
        .Chart.Paste
        .Chart.ChartArea.Border.LineStyle = xlNone
        .Chart.Export path, imgType
    End With
    
    rng.Worksheet.ChartObjects(rng.Worksheet.ChartObjects.count).Delete
    rng.Worksheet.Activate
    ActiveWindow.DisplayGridlines = True
    createJpg = path
End Function
