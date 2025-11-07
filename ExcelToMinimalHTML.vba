Sub ExcelToMinimalHTML()
    Dim rng As Range
    Dim htmlStr As String
    Dim fso As Object
    Dim ts As Object
    Dim tempFile As String
    Dim re As Object

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range to convert!", vbExclamation
        Exit Sub
    End If

    Set rng = Selection
    tempFile = Environ("TEMP") & "\temp_excel_" & Format(Now, "yyyymmddhhmmss") & ".html"

    With ActiveWorkbook.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
        Filename:=tempFile, _
        Sheet:=rng.Worksheet.Name, _
        Source:=rng.Address, _
        HtmlType:=xlHtmlStatic)
        .Publish True
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(tempFile, 1, False, -2)
    htmlStr = ts.ReadAll
    ts.Close
    fso.DeleteFile tempFile

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True

    re.Pattern = "<!\[if\s+supportMisalignedColumns\]>[\s\S]*?<!\[endif\]>"
    htmlStr = re.Replace(htmlStr, "")

    re.Pattern = "<head\b[^>]*>[\s\S]*?</head>"
    htmlStr = re.Replace(htmlStr, "")

    re.Pattern = "<!--[\s\S]*?-->"
    htmlStr = re.Replace(htmlStr, "")

    re.Pattern = "<span\b[^>]*>\s*&nbsp;\s*</span>"
    htmlStr = re.Replace(htmlStr, "")

    re.Pattern = ">\s+<"
    htmlStr = re.Replace(htmlStr, "><")

    re.Pattern = "\s+(?!(rowspan|colspan)\b)[A-Za-z0-9\-:]+=(?:""[^""]*""|'[^']*'|[^>\s]+)"
    htmlStr = re.Replace(htmlStr, "")

    re.Pattern = "　"
    htmlStr = re.Replace(htmlStr, "")

    re.Pattern = "\s{2,}"
    htmlStr = re.Replace(htmlStr, " ")

    htmlStr = Trim(htmlStr)

    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText htmlStr
        .PutInClipboard
    End With

End Sub
