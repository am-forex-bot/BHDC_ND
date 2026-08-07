Attribute VB_Name = "modInsertDocId"
Option Explicit

' =============================================================================
' modInsertDocId - NetDocuments Document ID footer functions
' =============================================================================
' 45 public subs for ribbon: format(5) x pages(3) x alignment(3)
' Formats: Num, Ver, Name, NameNum, NameNumVer
' Naming: NDFooter_{format}_{All|First|NotFirst}_{L|C|R}
' Plus: InsertNDDocNum - insert at cursor position
'
' Multi-section support: if a document has more than one section, the user
' is shown a picker listing page ranges, break types and link status so
' they can choose which sections to update.
' =============================================================================

' --- Shared helpers ---

Private Function GetNDDocIdFromTitleBar(Optional includeVersion As Boolean = True) As String
    Dim titleText As String
    titleText = Application.ActiveWindow.Caption

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "([A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4})\s*(\.\d+|v\.?\d+)?"

    If regex.Test(titleText) Then
        Dim matches As Object
        Set matches = regex.Execute(titleText)
        Dim baseId As String
        baseId = matches(0).SubMatches(0)

        If includeVersion And Len(matches(0).SubMatches(1)) > 0 Then
            GetNDDocIdFromTitleBar = baseId & matches(0).SubMatches(1)
        Else
            GetNDDocIdFromTitleBar = baseId
        End If
    End If
End Function

Private Function GetDocNameFromTitleBar() As String
    Dim titleText As String
    titleText = Application.ActiveWindow.Caption

    Dim appSuffix As String
    appSuffix = " - " & Application.Name
    If Len(titleText) > Len(appSuffix) And _
       StrComp(Right(titleText, Len(appSuffix)), appSuffix, vbTextCompare) = 0 Then
        titleText = Left(titleText, Len(titleText) - Len(appSuffix))
    End If

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}\s*(\.\d+|v\.?\d+)?"
    titleText = regex.Replace(titleText, "")

    ' Strip file extension
    Dim extRegex As Object
    Set extRegex = CreateObject("VBScript.RegExp")
    extRegex.Global = False
    extRegex.IgnoreCase = True
    extRegex.Pattern = "\.(docx?|dotx?|docm|dotm|rtf|pdf|xlsx?|xlsm|pptx?|pptm)$"
    titleText = extRegex.Replace(Trim(titleText), "")

    titleText = Trim(titleText)
    If Len(titleText) > 0 And (Left(titleText, 1) = "-" Or Left(titleText, 1) = ChrW(&H2013)) Then
        titleText = Trim(Mid(titleText, 2))
    End If
    If Len(titleText) > 0 And (Right(titleText, 1) = "-" Or Right(titleText, 1) = ChrW(&H2013)) Then
        titleText = Trim(Left(titleText, Len(titleText) - 1))
    End If

    GetDocNameFromTitleBar = titleText
End Function

Private Function CreateNDRegex() As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}\s*(\.\d+|v\.?\d+)?"
    Set CreateNDRegex = regex
End Function

Private Function CreateIManageRegex() As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "\d{5,}(v\d+)?"
    Set CreateIManageRegex = regex
End Function

Private Function AlignmentConst(align As String) As Long
    Select Case LCase(align)
        Case "left": AlignmentConst = wdAlignParagraphLeft
        Case "center", "centre": AlignmentConst = wdAlignParagraphCenter
        Case Else: AlignmentConst = wdAlignParagraphRight
    End Select
End Function

Private Function BuildFooterText(mode As String) As String
    Dim docName As String
    Dim docId As String

    Select Case LCase(mode)
        Case "num"
            BuildFooterText = GetNDDocIdFromTitleBar(includeVersion:=False)
        Case "ver"
            BuildFooterText = GetNDDocIdFromTitleBar(includeVersion:=True)
        Case "name"
            If GetNDDocIdFromTitleBar(False) = "" Then
                BuildFooterText = ""
            Else
                BuildFooterText = GetDocNameFromTitleBar()
            End If
        Case "namenum"
            docName = GetDocNameFromTitleBar()
            docId = GetNDDocIdFromTitleBar(includeVersion:=False)
            If docName <> "" And docId <> "" Then
                BuildFooterText = docName & " - " & docId
            ElseIf docId <> "" Then
                BuildFooterText = docId
            Else
                BuildFooterText = docName
            End If
        Case "namenumver"
            docName = GetDocNameFromTitleBar()
            docId = GetNDDocIdFromTitleBar(includeVersion:=True)
            If docName <> "" And docId <> "" Then
                BuildFooterText = docName & " - " & docId
            ElseIf docId <> "" Then
                BuildFooterText = docId
            Else
                BuildFooterText = docName
            End If
    End Select
End Function

Private Function BuildNDLinkCard(docName As String, ndUrl As String) As String
    ' Show doc name with extension but without ND doc number
    Dim displayName As String
    displayName = GetDocNameFromTitleBar()
    ' Add back the file extension
    Dim dotPos As Long
    dotPos = InStrRev(ActiveDocument.Name, ".")
    If dotPos > 0 Then
        displayName = displayName & Mid(ActiveDocument.Name, dotPos)
    End If
    If displayName = "" Then displayName = ActiveDocument.Name

    Dim h As String
    h = "<table cellpadding='0' cellspacing='0' style='border:1px solid #d0d0d0;" & _
        "border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;" & _
        "max-width:520px;background-color:#f5f5f5;'>"
    h = h & "<tr>"
    h = h & "<td style='padding:10px 14px;font-size:13px;color:#333;font-weight:500;'>" & displayName & "</td>"
    h = h & "<td style='padding:10px 14px;text-align:right;white-space:nowrap;'>"
    h = h & "<a href='" & ndUrl & "' style='color:#0563C1;text-decoration:none;" & _
        "font-size:11px;font-weight:bold;letter-spacing:0.5px;'>OPEN</a>"
    h = h & "</td></tr>"
    h = h & "</table>"
    ' Trailing paragraph with an explicit reset background, so the cursor
    ' lands below the table on a clean line and pressing Enter can't pull
    ' the grey card background down with it
    h = h & "<p style='margin:0;padding:0;background:#ffffff;" & _
        "font-family:Segoe UI,Arial,sans-serif;font-size:11pt;color:#000000;'>&nbsp;</p>"
    BuildNDLinkCard = h
End Function

Private Function SanitizeFileName(name As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    Dim result As String
    result = name
    For i = LBound(bad) To UBound(bad)
        result = Replace(result, CStr(bad(i)), "")
    Next i
    SanitizeFileName = Trim(result)
End Function

Private Function BuildRefreshText(existingText As String, newDocId As String) As String
    Dim ndRegex As Object
    Set ndRegex = CreateNDRegex()

    Dim cleanText As String
    cleanText = existingText
    If Len(cleanText) > 0 Then
        If Asc(Right(cleanText, 1)) = 13 Then
            cleanText = Left(cleanText, Len(cleanText) - 1)
        End If
    End If

    If ndRegex.Test(cleanText) Then
        Dim matches As Object
        Set matches = ndRegex.Execute(cleanText)
        If matches(0).FirstIndex > 0 Then
            Dim docName As String
            docName = GetDocNameFromTitleBar()
            If docName <> "" Then
                BuildRefreshText = docName & " - " & newDocId
                Exit Function
            End If
        End If
    End If

    BuildRefreshText = newDocId
End Function

' --- Section helpers ---

Private Function GetSectionInfo() As String
    Dim info As String
    Dim sec As Section
    Dim i As Long
    Dim rng As Range
    Dim startPage As Long
    Dim endPage As Long
    Dim pageStr As String
    Dim details As String

    i = 0
    For Each sec In ActiveDocument.Sections
        i = i + 1

        Set rng = sec.Range
        rng.Collapse wdCollapseStart
        startPage = rng.Information(wdActiveEndPageNumber)

        Set rng = sec.Range
        rng.Collapse wdCollapseEnd
        endPage = rng.Information(wdActiveEndPageNumber)

        If startPage = endPage Then
            pageStr = "p." & startPage
        Else
            pageStr = "pp." & startPage & "-" & endPage
        End If

        details = ""
        If i > 1 Then
            Select Case sec.PageSetup.SectionStart
                Case wdSectionNewPage: details = "New Page"
                Case wdSectionContinuous: details = "Continuous"
                Case wdSectionEvenPage: details = "Even Page"
                Case wdSectionOddPage: details = "Odd Page"
            End Select

            If sec.Footers(wdHeaderFooterPrimary).LinkToPrevious Then
                If details <> "" Then details = details & ", "
                details = details & "linked"
            Else
                If details <> "" Then details = details & ", "
                details = details & "independent"
            End If
        End If

        If sec.Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection Then
            If details <> "" Then details = details & ", "
            details = details & "numbering restarts at " & _
                      sec.Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber
        End If

        If sec.Footers(wdHeaderFooterPrimary).PageNumbers.NumberStyle <> wdPageNumberStyleArabic Then
            If details <> "" Then details = details & ", "
            details = details & "format " & PageNumStyleName(sec)
        End If

        info = info & "  " & i & ". " & pageStr
        If details <> "" Then info = info & " (" & details & ")"
        info = info & vbCrLf
    Next sec

    GetSectionInfo = info
End Function

Private Function ParseSectionSelection(userInput As String, sectionCount As Long) As Collection
    Dim result As New Collection
    Dim inp As String
    Dim parts() As String
    Dim part As Variant
    Dim trimmed As String
    Dim rangeParts() As String
    Dim startNum As Long
    Dim endNum As Long
    Dim k As Long
    Dim num As Long

    inp = Trim(LCase(userInput))

    If inp = "all" Then
        For k = 1 To sectionCount
            result.Add k
        Next k
        Set ParseSectionSelection = result
        Exit Function
    End If

    parts = Split(inp, ",")
    For Each part In parts
        trimmed = Trim(CStr(part))
        If InStr(trimmed, "-") > 0 Then
            rangeParts = Split(trimmed, "-")
            If UBound(rangeParts) = 1 Then
                If IsNumeric(Trim(rangeParts(0))) And IsNumeric(Trim(rangeParts(1))) Then
                    startNum = CLng(Trim(rangeParts(0)))
                    endNum = CLng(Trim(rangeParts(1)))
                    For k = startNum To endNum
                        If k >= 1 And k <= sectionCount Then result.Add k
                    Next k
                End If
            End If
        ElseIf IsNumeric(trimmed) Then
            num = CLng(trimmed)
            If num >= 1 And num <= sectionCount Then result.Add num
        End If
    Next part

    Set ParseSectionSelection = result
End Function

Private Function PickSectionsForUpdate(dlgTitle As String) As Collection
    ' Shows the section picker for multi-section documents.
    ' Returns a Collection of section indices, or Nothing if cancelled/invalid.
    Dim sectionCount As Long
    sectionCount = ActiveDocument.Sections.Count

    Dim result As Collection

    If sectionCount = 1 Then
        Set result = New Collection
        result.Add 1
        Set PickSectionsForUpdate = result
        Exit Function
    End If

    Dim info As String
    info = "This document has " & sectionCount & " sections:" & vbCrLf & vbCrLf
    info = info & GetSectionInfo()
    info = info & vbCrLf & "Enter sections to update (e.g. ""1,3"" or ""2-4"" or ""all""):"

    Dim userInput As String
    userInput = InputBox(info, dlgTitle, "all")
    If Len(userInput) = 0 Then Exit Function

    Set result = ParseSectionSelection(userInput, sectionCount)
    If result.Count = 0 Then
        MsgBox "No valid sections selected.", vbExclamation, dlgTitle
        Exit Function
    End If

    Set PickSectionsForUpdate = result
End Function

Private Sub UnlinkSectionFooters(sec As Section)
    If sec.Index = 1 Then Exit Sub
    If sec.Footers(wdHeaderFooterPrimary).LinkToPrevious Then
        sec.Footers(wdHeaderFooterPrimary).LinkToPrevious = False
    End If
    If sec.Footers(wdHeaderFooterFirstPage).LinkToPrevious Then
        sec.Footers(wdHeaderFooterFirstPage).LinkToPrevious = False
    End If
End Sub

Private Sub UnlinkUnselectedNeighbours(selectedSections As Collection)
    ' Linked footers share storage, so a change to one section shows in its
    ' linked neighbours. Break the link in BOTH directions at every boundary
    ' between a selected and an unselected section:
    '  - a selected section linked to an unselected previous one would
    '    otherwise rewrite the previous section's footer
    '  - an unselected next section linked to a selected one would otherwise
    '    pick up the change
    Dim sectionCount As Long
    sectionCount = ActiveDocument.Sections.Count
    If sectionCount = 1 Then Exit Sub

    Dim isSelected() As Boolean
    ReDim isSelected(1 To sectionCount)
    Dim s As Variant
    For Each s In selectedSections
        isSelected(CLng(s)) = True
    Next s

    Dim i As Long
    For i = 1 To sectionCount
        If isSelected(i) Then
            If i > 1 Then
                If Not isSelected(i - 1) Then
                    UnlinkSectionFooters ActiveDocument.Sections(i)
                End If
            End If
            If i < sectionCount Then
                If Not isSelected(i + 1) Then
                    UnlinkSectionFooters ActiveDocument.Sections(i + 1)
                End If
            End If
        End If
    Next i
End Sub

' --- Footer manipulation ---

Private Sub ReplaceParaText(para As Paragraph, docId As String, align As Long)
    Dim rng As Range
    Set rng = para.Range
    rng.MoveEnd wdCharacter, -1
    rng.Text = docId
    rng.Font.Size = 8
    rng.Font.Color = RGB(128, 128, 128)
    rng.ParagraphFormat.Alignment = align
End Sub

Private Sub InsertOrReplaceInFooter(ftr As HeaderFooter, footerText As String, align As Long)
    Dim ndRegex As Object
    Set ndRegex = CreateNDRegex()
    Dim imRegex As Object
    Set imRegex = CreateIManageRegex()

    Dim para As Paragraph
    For Each para In ftr.Range.Paragraphs
        If ndRegex.Test(para.Range.Text) Or imRegex.Test(para.Range.Text) Then
            ReplaceParaText para, footerText, align
            Exit Sub
        End If
    Next para

    ' Second pass: previously inserted ref (size 8, grey text)
    For Each para In ftr.Range.Paragraphs
        Dim rng As Range
        Set rng = para.Range
        rng.MoveEnd wdCharacter, -1
        If Len(rng.Text) > 0 And rng.Font.Size = 8 And rng.Font.Color = RGB(128, 128, 128) Then
            ReplaceParaText para, footerText, align
            Exit Sub
        End If
    Next para

    ' No existing ref found - append new paragraph
    Dim insertRange As Range
    Set insertRange = ftr.Range
    insertRange.Collapse wdCollapseEnd
    insertRange.Text = vbCr & footerText
    insertRange.Font.Size = 8
    insertRange.Font.Color = RGB(128, 128, 128)
    insertRange.ParagraphFormat.Alignment = align
End Sub

Private Sub RemoveDocRefFromFooter(ftr As HeaderFooter)
    Dim ndRegex As Object
    Set ndRegex = CreateNDRegex()
    Dim imRegex As Object
    Set imRegex = CreateIManageRegex()

    Dim para As Paragraph
    For Each para In ftr.Range.Paragraphs
        If ndRegex.Test(para.Range.Text) Or imRegex.Test(para.Range.Text) Then
            para.Range.Delete
            Exit Sub
        End If
    Next para

    For Each para In ftr.Range.Paragraphs
        Dim rng As Range
        Set rng = para.Range
        rng.MoveEnd wdCharacter, -1
        If Len(rng.Text) > 0 And rng.Font.Size = 8 And rng.Font.Color = RGB(128, 128, 128) Then
            para.Range.Delete
            Exit Sub
        End If
    Next para
End Sub

Private Sub SweepStaleWatermarks(hdr As HeaderFooter)
    ' When DifferentFirstPageHeaderFooter is flipped from False to True,
    ' the first-page header story becomes visible. It may contain stale
    ' watermark shapes from the template. Remove them - but only if the
    ' header is not linked to a previous section (to avoid deleting
    ' intentional watermarks from an earlier section).
    On Error Resume Next
    Dim shp As Shape
    Dim i As Long
    For i = hdr.Shapes.Count To 1 Step -1
        Set shp = hdr.Shapes(i)
        If InStr(1, shp.Name, "PowerPlusWaterMarkObject", vbTextCompare) = 1 Then
            shp.Delete
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub ApplyFooterToSection(sec As Section, footerText As String, pageMode As String, al As Long)
    Select Case LCase(pageMode)
        Case "all"
            InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), footerText, al
            If sec.PageSetup.DifferentFirstPageHeaderFooter Then
                InsertOrReplaceInFooter sec.Footers(wdHeaderFooterFirstPage), footerText, al
            End If

        Case "first"
            If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
                sec.PageSetup.DifferentFirstPageHeaderFooter = True
                If sec.Index = 1 Or Not sec.Headers(wdHeaderFooterFirstPage).LinkToPrevious Then
                    SweepStaleWatermarks sec.Headers(wdHeaderFooterFirstPage)
                End If
            End If
            InsertOrReplaceInFooter sec.Footers(wdHeaderFooterFirstPage), footerText, al
            RemoveDocRefFromFooter sec.Footers(wdHeaderFooterPrimary)

        Case "notfirst"
            If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
                sec.PageSetup.DifferentFirstPageHeaderFooter = True
                If sec.Index = 1 Or Not sec.Headers(wdHeaderFooterFirstPage).LinkToPrevious Then
                    SweepStaleWatermarks sec.Headers(wdHeaderFooterFirstPage)
                End If
            End If
            InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), footerText, al
            RemoveDocRefFromFooter sec.Footers(wdHeaderFooterFirstPage)
    End Select
End Sub

' --- Page numbers ---

Private Function ParaHasPageField(para As Paragraph) As Boolean
    Dim f As Field
    For Each f In para.Range.Fields
        If f.Type = wdFieldPage Or f.Type = wdFieldNumPages Or _
           f.Type = wdFieldSectionPages Then
            ParaHasPageField = True
            Exit Function
        End If
    Next f
End Function

Private Function StoryHasPageField(hf As HeaderFooter) As Boolean
    Dim f As Field
    For Each f In hf.Range.Fields
        If f.Type = wdFieldPage Then
            StoryHasPageField = True
            Exit Function
        End If
    Next f
End Function

Private Function PageNumStyleName(sec As Section) As String
    Select Case sec.Footers(wdHeaderFooterPrimary).PageNumbers.NumberStyle
        Case wdPageNumberStyleLowercaseRoman: PageNumStyleName = "i, ii, iii"
        Case wdPageNumberStyleUppercaseRoman: PageNumStyleName = "I, II, III"
        Case wdPageNumberStyleLowercaseLetter: PageNumStyleName = "a, b, c"
        Case wdPageNumberStyleUppercaseLetter: PageNumStyleName = "A, B, C"
        Case Else: PageNumStyleName = "1, 2, 3"
    End Select
End Function

Private Function DescribePageNumLocations(sec As Section) As String
    ' Lists the active header/footer stories that contain page numbers
    Dim parts As String

    If StoryHasPageField(sec.Footers(wdHeaderFooterPrimary)) Then
        parts = "footer"
    End If
    If sec.PageSetup.DifferentFirstPageHeaderFooter Then
        If StoryHasPageField(sec.Footers(wdHeaderFooterFirstPage)) Then
            If parts <> "" Then parts = parts & ", "
            parts = parts & "first-page footer"
        End If
    End If
    If sec.PageSetup.OddAndEvenPagesHeaderFooter Then
        If StoryHasPageField(sec.Footers(wdHeaderFooterEvenPages)) Then
            If parts <> "" Then parts = parts & ", "
            parts = parts & "even-page footer"
        End If
    End If

    If StoryHasPageField(sec.Headers(wdHeaderFooterPrimary)) Then
        If parts <> "" Then parts = parts & ", "
        parts = parts & "header"
    End If
    If sec.PageSetup.DifferentFirstPageHeaderFooter Then
        If StoryHasPageField(sec.Headers(wdHeaderFooterFirstPage)) Then
            If parts <> "" Then parts = parts & ", "
            parts = parts & "first-page header"
        End If
    End If
    If sec.PageSetup.OddAndEvenPagesHeaderFooter Then
        If StoryHasPageField(sec.Headers(wdHeaderFooterEvenPages)) Then
            If parts <> "" Then parts = parts & ", "
            parts = parts & "even-page header"
        End If
    End If

    DescribePageNumLocations = parts
End Function

Private Function ParaIsOnlyPageNumber(para As Paragraph) As Boolean
    ' True if the paragraph contains nothing but page-number content:
    ' PAGE/NUMPAGES field results plus words like "Page"/"of" and digits
    Dim t As String
    t = LCase(para.Range.Text)
    t = Replace(t, "page", "")
    t = Replace(t, "of", "")

    Dim k As Long
    For k = 1 To Len(t)
        If Mid(t, k, 1) Like "[a-z]" Then
            ParaIsOnlyPageNumber = False
            Exit Function
        End If
    Next k
    ParaIsOnlyPageNumber = True
End Function

Private Sub RemovePageNumFromFooter(ftr As HeaderFooter)
    Dim i As Long
    Dim para As Paragraph
    For i = ftr.Range.Paragraphs.Count To 1 Step -1
        Set para = ftr.Range.Paragraphs(i)
        If ParaHasPageField(para) Then
            If ParaIsOnlyPageNumber(para) Then
                para.Range.Delete
            Else
                ' Paragraph has other content (e.g. doc ref on the same
                ' line) - surgically delete just the page fields
                Dim j As Long
                For j = para.Range.Fields.Count To 1 Step -1
                    If para.Range.Fields(j).Type = wdFieldPage Or _
                       para.Range.Fields(j).Type = wdFieldNumPages Or _
                       para.Range.Fields(j).Type = wdFieldSectionPages Then
                        para.Range.Fields(j).Delete
                    End If
                Next j
            End If
        End If
    Next i
End Sub

Private Function RemoveAllPageNumbers() As Long
    ' Sweeps every header and footer story in every section, including
    ' first-page and even-page variants, and resets numbering back to
    ' continuous so a future insert behaves predictably.
    ' Returns the number of sections processed.
    Dim sec As Section
    Dim hfTypes As Variant
    hfTypes = Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)
    Dim v As Variant
    Dim n As Long

    For Each sec In ActiveDocument.Sections
        For Each v In hfTypes
            RemovePageNumFromFooter sec.Footers(CLng(v))
            RemovePageNumFromFooter sec.Headers(CLng(v))
        Next v
        sec.Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = False
        n = n + 1
    Next sec

    RemoveAllPageNumbers = n
End Function

Private Sub InsertPageNumInFooter(ftr As HeaderFooter, fmt As String, al As Long)
    ' Replace any existing page number, then append a new paragraph with
    ' PAGE (and NUMPAGES for "Page X of Y") fields
    RemovePageNumFromFooter ftr

    Dim insertRange As Range
    Set insertRange = ftr.Range
    insertRange.Collapse wdCollapseEnd

    ' If the footer is empty, use its existing paragraph rather than
    ' appending a new one (avoids a stray blank line above the number)
    Dim prefix As String
    If Len(ftr.Range.Text) <= 1 Then
        prefix = ""
    Else
        prefix = vbCr
    End If

    Dim basePos As Long
    Dim fldRng As Range

    If LCase(fmt) = "oftotal" Or LCase(fmt) = "ofsection" Then
        ' "oftotal" counts the whole document (NUMPAGES);
        ' "ofsection" counts just this section (SECTIONPAGES)
        Dim totalFieldType As Long
        If LCase(fmt) = "ofsection" Then
            totalFieldType = wdFieldSectionPages
        Else
            totalFieldType = wdFieldNumPages
        End If

        insertRange.Text = prefix & "Page  of "
        basePos = insertRange.Start

        ' Add the total field at the end first so the earlier position stays valid
        Set fldRng = ftr.Range.Duplicate
        fldRng.Start = insertRange.End
        fldRng.End = insertRange.End
        ftr.Range.Fields.Add fldRng, totalFieldType

        ' PAGE field goes after "Page "
        Set fldRng = ftr.Range.Duplicate
        fldRng.Start = basePos + Len(prefix) + 5
        fldRng.End = basePos + Len(prefix) + 5
        ftr.Range.Fields.Add fldRng, wdFieldPage
    Else
        If prefix <> "" Then insertRange.Text = prefix
        Set fldRng = ftr.Range.Duplicate
        fldRng.Start = insertRange.End
        fldRng.End = insertRange.End
        ftr.Range.Fields.Add fldRng, wdFieldPage
    End If

    Dim lastPara As Paragraph
    Set lastPara = ftr.Range.Paragraphs(ftr.Range.Paragraphs.Count)
    lastPara.Range.ParagraphFormat.Alignment = al
End Sub

Private Sub ApplyPageNumToSection(sec As Section, fmt As String, pageMode As String, al As Long)
    ' Documents with "Different Odd & Even Pages" keep a separate even-page
    ' footer story - cover it too, or even pages silently get no number
    Dim hasEvenPages As Boolean
    hasEvenPages = sec.PageSetup.OddAndEvenPagesHeaderFooter

    Select Case LCase(pageMode)
        Case "all"
            InsertPageNumInFooter sec.Footers(wdHeaderFooterPrimary), fmt, al
            If sec.PageSetup.DifferentFirstPageHeaderFooter Then
                InsertPageNumInFooter sec.Footers(wdHeaderFooterFirstPage), fmt, al
            End If
            If hasEvenPages Then
                InsertPageNumInFooter sec.Footers(wdHeaderFooterEvenPages), fmt, al
            End If

        Case "first"
            If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
                sec.PageSetup.DifferentFirstPageHeaderFooter = True
                If sec.Index = 1 Or Not sec.Headers(wdHeaderFooterFirstPage).LinkToPrevious Then
                    SweepStaleWatermarks sec.Headers(wdHeaderFooterFirstPage)
                End If
            End If
            InsertPageNumInFooter sec.Footers(wdHeaderFooterFirstPage), fmt, al
            RemovePageNumFromFooter sec.Footers(wdHeaderFooterPrimary)
            If hasEvenPages Then
                RemovePageNumFromFooter sec.Footers(wdHeaderFooterEvenPages)
            End If

        Case "notfirst"
            If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
                sec.PageSetup.DifferentFirstPageHeaderFooter = True
                If sec.Index = 1 Or Not sec.Headers(wdHeaderFooterFirstPage).LinkToPrevious Then
                    SweepStaleWatermarks sec.Headers(wdHeaderFooterFirstPage)
                End If
            End If
            InsertPageNumInFooter sec.Footers(wdHeaderFooterPrimary), fmt, al
            If hasEvenPages Then
                InsertPageNumInFooter sec.Footers(wdHeaderFooterEvenPages), fmt, al
            End If
            RemovePageNumFromFooter sec.Footers(wdHeaderFooterFirstPage)
    End Select
End Sub

Private Sub DoInsertPageNumbers(fmt As String, pageMode As String, align As String)
    On Error GoTo ErrorHandler

    Dim cursorPos As Range
    Set cursorPos = Selection.Range

    Dim al As Long
    al = AlignmentConst(align)

    Dim selectedSections As Collection
    Set selectedSections = PickSectionsForUpdate("Page Numbers")
    If selectedSections Is Nothing Then Exit Sub

    UnlinkUnselectedNeighbours selectedSections

    Dim idx As Variant
    For Each idx In selectedSections
        ApplyPageNumToSection ActiveDocument.Sections(CLng(idx)), fmt, pageMode, al
    Next idx

    ' "Page X of Y (this section)" reads wrong unless the section restarts
    ' numbering at 1 (e.g. "Page 5 of 3") - offer to fix it
    If LCase(fmt) = "ofsection" Then
        Dim needsRestart As Boolean
        For Each idx In selectedSections
            If CLng(idx) > 1 Then
                If Not ActiveDocument.Sections(CLng(idx)).Footers(wdHeaderFooterPrimary) _
                       .PageNumbers.RestartNumberingAtSection Then
                    needsRestart = True
                End If
            End If
        Next idx

        If needsRestart Then
            If MsgBox("""Page X of Y"" for a section usually reads best when the " & _
                      "section restarts numbering at 1 (otherwise you can get " & _
                      """Page 5 of 3"")." & vbCrLf & vbCrLf & _
                      "Restart the selected section(s) at 1?", _
                      vbYesNo + vbQuestion, "Page Numbers") = vbYes Then
                For Each idx In selectedSections
                    If CLng(idx) > 1 Then
                        With ActiveDocument.Sections(CLng(idx)) _
                                 .Footers(wdHeaderFooterPrimary).PageNumbers
                            .StartingNumber = 1
                            .RestartNumberingAtSection = True
                        End With
                    End If
                Next idx
            End If
        End If
    End If

    cursorPos.Select

    Dim pageTxt As String
    Select Case LCase(pageMode)
        Case "all": pageTxt = "all pages"
        Case "first": pageTxt = "first page footer"
        Case "notfirst": pageTxt = "all pages except first"
    End Select

    MsgBox "Page numbers inserted into " & pageTxt & ".", _
           vbInformation, "Page Numbers"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred inserting page numbers." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Page Numbers"
End Sub

' --- Shared implementation ---

Private Sub DoInsertFooter(mode As String, pageMode As String, align As String)
    On Error GoTo ErrorHandler

    Dim footerText As String
    footerText = BuildFooterText(mode)

    If footerText = "" Then
        MsgBox "Could not find a NetDocuments reference in the title bar." & vbCrLf & vbCrLf & _
               "Make sure the document is saved to NetDocuments first.", _
               vbExclamation, "Insert ND Ref"
        Exit Sub
    End If

    Dim cursorPos As Range
    Set cursorPos = Selection.Range

    Dim al As Long
    al = AlignmentConst(align)

    Dim selectedSections As Collection
    Set selectedSections = PickSectionsForUpdate("Insert ND Ref")
    If selectedSections Is Nothing Then Exit Sub

    UnlinkUnselectedNeighbours selectedSections

    Dim idx As Variant
    For Each idx In selectedSections
        ApplyFooterToSection ActiveDocument.Sections(CLng(idx)), footerText, pageMode, al
    Next idx

    cursorPos.Select

    Dim pageTxt As String
    Select Case LCase(pageMode)
        Case "all": pageTxt = "all pages"
        Case "first": pageTxt = "first page footer"
        Case "notfirst": pageTxt = "all pages except first"
    End Select

    MsgBox footerText & vbCrLf & vbCrLf & _
           "Inserted into " & pageTxt & ".", _
           vbInformation, "Insert ND Ref"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred inserting the ND reference." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Insert ND Ref"
End Sub

' =============================================================================
' Number Only - All Pages
' =============================================================================
Public Sub NDFooter_Num_All_L()
    DoInsertFooter "num", "all", "left"
End Sub

Public Sub NDFooter_Num_All_C()
    DoInsertFooter "num", "all", "center"
End Sub

Public Sub NDFooter_Num_All_R()
    DoInsertFooter "num", "all", "right"
End Sub

' =============================================================================
' Number Only - First Page Only
' =============================================================================
Public Sub NDFooter_Num_First_L()
    DoInsertFooter "num", "first", "left"
End Sub

Public Sub NDFooter_Num_First_C()
    DoInsertFooter "num", "first", "center"
End Sub

Public Sub NDFooter_Num_First_R()
    DoInsertFooter "num", "first", "right"
End Sub

' =============================================================================
' Number Only - All Except First
' =============================================================================
Public Sub NDFooter_Num_NotFirst_L()
    DoInsertFooter "num", "notfirst", "left"
End Sub

Public Sub NDFooter_Num_NotFirst_C()
    DoInsertFooter "num", "notfirst", "center"
End Sub

Public Sub NDFooter_Num_NotFirst_R()
    DoInsertFooter "num", "notfirst", "right"
End Sub

' =============================================================================
' With Version - All Pages
' =============================================================================
Public Sub NDFooter_Ver_All_L()
    DoInsertFooter "ver", "all", "left"
End Sub

Public Sub NDFooter_Ver_All_C()
    DoInsertFooter "ver", "all", "center"
End Sub

Public Sub NDFooter_Ver_All_R()
    DoInsertFooter "ver", "all", "right"
End Sub

' =============================================================================
' With Version - First Page Only
' =============================================================================
Public Sub NDFooter_Ver_First_L()
    DoInsertFooter "ver", "first", "left"
End Sub

Public Sub NDFooter_Ver_First_C()
    DoInsertFooter "ver", "first", "center"
End Sub

Public Sub NDFooter_Ver_First_R()
    DoInsertFooter "ver", "first", "right"
End Sub

' =============================================================================
' With Version - All Except First
' =============================================================================
Public Sub NDFooter_Ver_NotFirst_L()
    DoInsertFooter "ver", "notfirst", "left"
End Sub

Public Sub NDFooter_Ver_NotFirst_C()
    DoInsertFooter "ver", "notfirst", "center"
End Sub

Public Sub NDFooter_Ver_NotFirst_R()
    DoInsertFooter "ver", "notfirst", "right"
End Sub

' =============================================================================
' Name Only - All Pages
' =============================================================================
Public Sub NDFooter_Name_All_L()
    DoInsertFooter "name", "all", "left"
End Sub

Public Sub NDFooter_Name_All_C()
    DoInsertFooter "name", "all", "center"
End Sub

Public Sub NDFooter_Name_All_R()
    DoInsertFooter "name", "all", "right"
End Sub

' =============================================================================
' Name Only - First Page Only
' =============================================================================
Public Sub NDFooter_Name_First_L()
    DoInsertFooter "name", "first", "left"
End Sub

Public Sub NDFooter_Name_First_C()
    DoInsertFooter "name", "first", "center"
End Sub

Public Sub NDFooter_Name_First_R()
    DoInsertFooter "name", "first", "right"
End Sub

' =============================================================================
' Name Only - All Except First
' =============================================================================
Public Sub NDFooter_Name_NotFirst_L()
    DoInsertFooter "name", "notfirst", "left"
End Sub

Public Sub NDFooter_Name_NotFirst_C()
    DoInsertFooter "name", "notfirst", "center"
End Sub

Public Sub NDFooter_Name_NotFirst_R()
    DoInsertFooter "name", "notfirst", "right"
End Sub

' =============================================================================
' Name + Number - All Pages
' =============================================================================
Public Sub NDFooter_NameNum_All_L()
    DoInsertFooter "namenum", "all", "left"
End Sub

Public Sub NDFooter_NameNum_All_C()
    DoInsertFooter "namenum", "all", "center"
End Sub

Public Sub NDFooter_NameNum_All_R()
    DoInsertFooter "namenum", "all", "right"
End Sub

' =============================================================================
' Name + Number - First Page Only
' =============================================================================
Public Sub NDFooter_NameNum_First_L()
    DoInsertFooter "namenum", "first", "left"
End Sub

Public Sub NDFooter_NameNum_First_C()
    DoInsertFooter "namenum", "first", "center"
End Sub

Public Sub NDFooter_NameNum_First_R()
    DoInsertFooter "namenum", "first", "right"
End Sub

' =============================================================================
' Name + Number - All Except First
' =============================================================================
Public Sub NDFooter_NameNum_NotFirst_L()
    DoInsertFooter "namenum", "notfirst", "left"
End Sub

Public Sub NDFooter_NameNum_NotFirst_C()
    DoInsertFooter "namenum", "notfirst", "center"
End Sub

Public Sub NDFooter_NameNum_NotFirst_R()
    DoInsertFooter "namenum", "notfirst", "right"
End Sub

' =============================================================================
' Name + Number + Version - All Pages
' =============================================================================
Public Sub NDFooter_NameNumVer_All_L()
    DoInsertFooter "namenumver", "all", "left"
End Sub

Public Sub NDFooter_NameNumVer_All_C()
    DoInsertFooter "namenumver", "all", "center"
End Sub

Public Sub NDFooter_NameNumVer_All_R()
    DoInsertFooter "namenumver", "all", "right"
End Sub

' =============================================================================
' Name + Number + Version - First Page Only
' =============================================================================
Public Sub NDFooter_NameNumVer_First_L()
    DoInsertFooter "namenumver", "first", "left"
End Sub

Public Sub NDFooter_NameNumVer_First_C()
    DoInsertFooter "namenumver", "first", "center"
End Sub

Public Sub NDFooter_NameNumVer_First_R()
    DoInsertFooter "namenumver", "first", "right"
End Sub

' =============================================================================
' Name + Number + Version - All Except First
' =============================================================================
Public Sub NDFooter_NameNumVer_NotFirst_L()
    DoInsertFooter "namenumver", "notfirst", "left"
End Sub

Public Sub NDFooter_NameNumVer_NotFirst_C()
    DoInsertFooter "namenumver", "notfirst", "center"
End Sub

Public Sub NDFooter_NameNumVer_NotFirst_R()
    DoInsertFooter "namenumver", "notfirst", "right"
End Sub

' =============================================================================
' Legacy names (keep for backward compatibility with existing ribbon tags)
' =============================================================================
Public Sub InsertNDDocIdAllPages()
    DoInsertFooter "ver", "all", "right"
End Sub

Public Sub InsertNDDocIdFirstOnly()
    DoInsertFooter "ver", "first", "right"
End Sub

Public Sub InsertNDDocIdAllButFirst()
    DoInsertFooter "ver", "notfirst", "right"
End Sub

Public Sub InsertNDDocIdAllPagesNoVer()
    DoInsertFooter "num", "all", "right"
End Sub

Public Sub InsertNDDocIdFirstOnlyNoVer()
    DoInsertFooter "num", "first", "right"
End Sub

Public Sub InsertNDDocIdAllButFirstNoVer()
    DoInsertFooter "num", "notfirst", "right"
End Sub

' =============================================================================
' Page Numbers - Number Only
' =============================================================================
Public Sub NDPageNum_Num_All_L()
    DoInsertPageNumbers "num", "all", "left"
End Sub

Public Sub NDPageNum_Num_All_C()
    DoInsertPageNumbers "num", "all", "center"
End Sub

Public Sub NDPageNum_Num_All_R()
    DoInsertPageNumbers "num", "all", "right"
End Sub

Public Sub NDPageNum_Num_First_L()
    DoInsertPageNumbers "num", "first", "left"
End Sub

Public Sub NDPageNum_Num_First_C()
    DoInsertPageNumbers "num", "first", "center"
End Sub

Public Sub NDPageNum_Num_First_R()
    DoInsertPageNumbers "num", "first", "right"
End Sub

Public Sub NDPageNum_Num_NotFirst_L()
    DoInsertPageNumbers "num", "notfirst", "left"
End Sub

Public Sub NDPageNum_Num_NotFirst_C()
    DoInsertPageNumbers "num", "notfirst", "center"
End Sub

Public Sub NDPageNum_Num_NotFirst_R()
    DoInsertPageNumbers "num", "notfirst", "right"
End Sub

' =============================================================================
' Page Numbers - Page X of Y
' =============================================================================
Public Sub NDPageNum_OfTotal_All_L()
    DoInsertPageNumbers "oftotal", "all", "left"
End Sub

Public Sub NDPageNum_OfTotal_All_C()
    DoInsertPageNumbers "oftotal", "all", "center"
End Sub

Public Sub NDPageNum_OfTotal_All_R()
    DoInsertPageNumbers "oftotal", "all", "right"
End Sub

Public Sub NDPageNum_OfTotal_First_L()
    DoInsertPageNumbers "oftotal", "first", "left"
End Sub

Public Sub NDPageNum_OfTotal_First_C()
    DoInsertPageNumbers "oftotal", "first", "center"
End Sub

Public Sub NDPageNum_OfTotal_First_R()
    DoInsertPageNumbers "oftotal", "first", "right"
End Sub

Public Sub NDPageNum_OfTotal_NotFirst_L()
    DoInsertPageNumbers "oftotal", "notfirst", "left"
End Sub

Public Sub NDPageNum_OfTotal_NotFirst_C()
    DoInsertPageNumbers "oftotal", "notfirst", "center"
End Sub

Public Sub NDPageNum_OfTotal_NotFirst_R()
    DoInsertPageNumbers "oftotal", "notfirst", "right"
End Sub

' =============================================================================
' Page Numbers - Page X of Y (This Section)
' =============================================================================
Public Sub NDPageNum_OfSection_All_L()
    DoInsertPageNumbers "ofsection", "all", "left"
End Sub

Public Sub NDPageNum_OfSection_All_C()
    DoInsertPageNumbers "ofsection", "all", "center"
End Sub

Public Sub NDPageNum_OfSection_All_R()
    DoInsertPageNumbers "ofsection", "all", "right"
End Sub

Public Sub NDPageNum_OfSection_First_L()
    DoInsertPageNumbers "ofsection", "first", "left"
End Sub

Public Sub NDPageNum_OfSection_First_C()
    DoInsertPageNumbers "ofsection", "first", "center"
End Sub

Public Sub NDPageNum_OfSection_First_R()
    DoInsertPageNumbers "ofsection", "first", "right"
End Sub

Public Sub NDPageNum_OfSection_NotFirst_L()
    DoInsertPageNumbers "ofsection", "notfirst", "left"
End Sub

Public Sub NDPageNum_OfSection_NotFirst_C()
    DoInsertPageNumbers "ofsection", "notfirst", "center"
End Sub

Public Sub NDPageNum_OfSection_NotFirst_R()
    DoInsertPageNumbers "ofsection", "notfirst", "right"
End Sub

' =============================================================================
' Page Numbers - Number format (per section)
' =============================================================================
Private Sub DoSetNumberFormat(styleVal As Long, styleLabel As String)
    On Error GoTo ErrorHandler

    Dim selectedSections As Collection
    Set selectedSections = PickSectionsForUpdate("Number Format")
    If selectedSections Is Nothing Then Exit Sub

    Dim idx As Variant
    For Each idx In selectedSections
        ActiveDocument.Sections(CLng(idx)).Footers(wdHeaderFooterPrimary) _
            .PageNumbers.NumberStyle = styleVal
    Next idx

    MsgBox "Page number format set to """ & styleLabel & """ for the " & _
           "selected section(s).", vbInformation, "Number Format"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred setting the number format." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Number Format"
End Sub

Public Sub NDPageNum_Fmt_Arabic()
    DoSetNumberFormat wdPageNumberStyleArabic, "1, 2, 3"
End Sub

Public Sub NDPageNum_Fmt_RomanLower()
    DoSetNumberFormat wdPageNumberStyleLowercaseRoman, "i, ii, iii"
End Sub

Public Sub NDPageNum_Fmt_RomanUpper()
    DoSetNumberFormat wdPageNumberStyleUppercaseRoman, "I, II, III"
End Sub

Public Sub NDPageNum_Fmt_LetterLower()
    DoSetNumberFormat wdPageNumberStyleLowercaseLetter, "a, b, c"
End Sub

Public Sub NDPageNum_Fmt_LetterUpper()
    DoSetNumberFormat wdPageNumberStyleUppercaseLetter, "A, B, C"
End Sub

' =============================================================================
' Page Numbers - Check / diagnose
' =============================================================================
Public Sub NDPageNum_Check()
    On Error GoTo ErrorHandler

    Dim report As String
    Dim warnings As String
    Dim sec As Section
    Dim i As Long
    Dim sectionCount As Long
    sectionCount = ActiveDocument.Sections.Count

    Dim hasNum() As Boolean
    ReDim hasNum(1 To sectionCount)
    Dim anyHave As Boolean

    For Each sec In ActiveDocument.Sections
        i = sec.Index

        ' Page range
        Dim rng As Range
        Set rng = sec.Range
        rng.Collapse wdCollapseStart
        Dim startPage As Long
        startPage = rng.Information(wdActiveEndPageNumber)
        Set rng = sec.Range
        rng.Collapse wdCollapseEnd
        Dim endPage As Long
        endPage = rng.Information(wdActiveEndPageNumber)

        Dim pageStr As String
        If startPage = endPage Then
            pageStr = "p." & startPage
        Else
            pageStr = "pp." & startPage & "-" & endPage
        End If

        Dim locations As String
        locations = DescribePageNumLocations(sec)
        hasNum(i) = (locations <> "")
        If hasNum(i) Then anyHave = True

        Dim line As String
        line = "Section " & i & " (" & pageStr & "): "
        If locations = "" Then
            line = line & "no page numbers"
        Else
            line = line & "numbers in " & locations
        End If

        line = line & "; format " & PageNumStyleName(sec)

        If sec.Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection Then
            line = line & "; RESTARTS at " & _
                   sec.Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber
        ElseIf i > 1 Then
            line = line & "; continues from previous"
        End If

        If i > 1 Then
            If sec.Footers(wdHeaderFooterPrimary).LinkToPrevious Then
                line = line & "; footer linked to previous"
            End If
        End If

        ' First page of the section deliberately unnumbered?
        If sec.PageSetup.DifferentFirstPageHeaderFooter Then
            If StoryHasPageField(sec.Footers(wdHeaderFooterPrimary)) And _
               Not StoryHasPageField(sec.Footers(wdHeaderFooterFirstPage)) Then
                line = line & "; no number on its first page"
            End If
        End If

        report = report & line & vbCrLf

        ' Duplicate numbering - header AND footer both carry numbers
        Dim footerHas As Boolean, headerHas As Boolean
        footerHas = StoryHasPageField(sec.Footers(wdHeaderFooterPrimary))
        If Not footerHas And sec.PageSetup.DifferentFirstPageHeaderFooter Then
            footerHas = StoryHasPageField(sec.Footers(wdHeaderFooterFirstPage))
        End If
        headerHas = StoryHasPageField(sec.Headers(wdHeaderFooterPrimary))
        If Not headerHas And sec.PageSetup.DifferentFirstPageHeaderFooter Then
            headerHas = StoryHasPageField(sec.Headers(wdHeaderFooterFirstPage))
        End If
        If footerHas And headerHas Then
            warnings = warnings & "- Section " & i & " has page numbers in BOTH " & _
                       "the header and the footer." & vbCrLf
        End If
    Next sec

    ' Sections missing numbers while others have them
    If anyHave Then
        Dim missing As String
        For i = 1 To sectionCount
            If Not hasNum(i) Then
                If missing <> "" Then missing = missing & ", "
                missing = missing & i
            End If
        Next i
        If missing <> "" Then
            warnings = warnings & "- Section(s) " & missing & " have no page " & _
                       "numbers while other sections do." & vbCrLf
        End If
    End If

    Dim msg As String
    msg = "PAGE NUMBERING REPORT" & vbCrLf & _
          String(40, "-") & vbCrLf & report

    If warnings <> "" Then
        msg = msg & vbCrLf & "POSSIBLE PROBLEMS" & vbCrLf & _
              String(40, "-") & vbCrLf & warnings & vbCrLf & _
              "Fix restarts with ""Continue from Previous"", or use " & _
              """Remove Page Numbers"" and re-insert for a clean start."
    ElseIf Not anyHave Then
        msg = msg & vbCrLf & "This document has no page numbers."
    Else
        msg = msg & vbCrLf & "No numbering problems found."
    End If

    MsgBox msg, vbInformation, "Check Numbering"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred checking page numbering." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Check Numbering"
End Sub

' =============================================================================
' Page Numbers - Numbering control (restart / continue / remove)
' =============================================================================
Private Sub DoSetNumbering(dlgTitle As String, restart As Boolean, Optional startNum As Long = 1)
    On Error GoTo ErrorHandler

    Dim selectedSections As Collection
    Set selectedSections = PickSectionsForUpdate(dlgTitle)
    If selectedSections Is Nothing Then Exit Sub

    Dim idx As Variant
    Dim sec As Section
    For Each idx In selectedSections
        Set sec = ActiveDocument.Sections(CLng(idx))
        If restart Then
            sec.Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = startNum
            sec.Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
        Else
            sec.Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = False
        End If
    Next idx

    If restart Then
        MsgBox "Page numbering will restart at " & startNum & _
               " in the selected section(s).", vbInformation, dlgTitle
    Else
        MsgBox "Page numbering will continue from the previous section " & _
               "in the selected section(s).", vbInformation, dlgTitle
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred updating page numbering." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, dlgTitle
End Sub

Public Sub NDPageNum_RestartAt1()
    DoSetNumbering "Restart Page Numbering", True, 1
End Sub

Public Sub NDPageNum_StartAt()
    Dim userInput As String
    userInput = InputBox("Start page numbering at:", "Start at Number", "1")
    If Len(userInput) = 0 Then Exit Sub

    If Not IsNumeric(userInput) Or CLng(Val(userInput)) < 0 Then
        MsgBox "Please enter a whole number (0 or higher).", _
               vbExclamation, "Start at Number"
        Exit Sub
    End If

    DoSetNumbering "Start at Number", True, CLng(Val(userInput))
End Sub

Public Sub NDPageNum_Continue()
    DoSetNumbering "Continue Page Numbering", False
End Sub

Public Sub NDPageNum_RemoveAll()
    On Error GoTo ErrorHandler

    Dim response As VbMsgBoxResult
    response = MsgBox("Remove ALL page numbers from every section of this document?" & _
                      vbCrLf & vbCrLf & _
                      "This clears page numbers from all headers and footers " & _
                      "(including first-page and even-page variants) and resets " & _
                      "numbering back to continuous.", _
                      vbYesNo + vbQuestion, "Remove Page Numbers")
    If response <> vbYes Then Exit Sub

    Dim cursorPos As Range
    Set cursorPos = Selection.Range

    Dim n As Long
    n = RemoveAllPageNumbers()

    cursorPos.Select

    MsgBox "Page numbers removed from " & n & " section(s).", _
           vbInformation, "Remove Page Numbers"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred removing page numbers." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Remove Page Numbers"
End Sub

' =============================================================================
' Refresh footer version (all sections)
' =============================================================================
Public Sub RefreshFooterVersion()
    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar(includeVersion:=True)

    If docId = "" Then
        MsgBox "Could not find a NetDocuments reference in the title bar." & vbCrLf & vbCrLf & _
               "Make sure the document is saved to NetDocuments first.", _
               vbExclamation, "Refresh Footer"
        Exit Sub
    End If

    Dim ndRegex As Object
    Set ndRegex = CreateNDRegex()
    Dim imRegex As Object
    Set imRegex = CreateIManageRegex()

    Dim sec As Section
    Dim updated As Boolean
    updated = False
    Dim para As Paragraph
    Dim existingAlign As Long
    Dim refreshText As String

    For Each sec In ActiveDocument.Sections
        For Each para In sec.Footers(wdHeaderFooterPrimary).Range.Paragraphs
            If ndRegex.Test(para.Range.Text) Or imRegex.Test(para.Range.Text) Then
                existingAlign = para.Range.ParagraphFormat.Alignment
                refreshText = BuildRefreshText(para.Range.Text, docId)
                ReplaceParaText para, refreshText, existingAlign
                updated = True
                Exit For
            End If
        Next para

        If sec.PageSetup.DifferentFirstPageHeaderFooter Then
            For Each para In sec.Footers(wdHeaderFooterFirstPage).Range.Paragraphs
                If ndRegex.Test(para.Range.Text) Or imRegex.Test(para.Range.Text) Then
                    existingAlign = para.Range.ParagraphFormat.Alignment
                    refreshText = BuildRefreshText(para.Range.Text, docId)
                    ReplaceParaText para, refreshText, existingAlign
                    updated = True
                    Exit For
                End If
            Next para
        End If
    Next sec

    If updated Then
        MsgBox "Footer updated to " & docId & ".", vbInformation, "Refresh Footer"
    Else
        MsgBox "No existing document reference found in the footer.", _
               vbExclamation, "Refresh Footer"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred refreshing the footer." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Refresh Footer"
End Sub

' =============================================================================
' Insert at cursor position
' =============================================================================
Public Sub InsertNDDocNum()
    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar()

    If docId = "" Then
        MsgBox "Could not find a NetDocuments reference in the title bar." & vbCrLf & vbCrLf & _
               "Make sure the document is saved to NetDocuments first.", _
               vbExclamation, "Insert ND Doc Number"
        Exit Sub
    End If

    Selection.TypeText Text:=docId

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred inserting the ND document number." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Insert ND Doc Number"
End Sub

' =============================================================================
' Email a copy of the current document
' =============================================================================
Public Sub EmailDocCopy()
    On Error GoTo ErrorHandler

    ' Document must be saved to a file before we can attach it
    If ActiveDocument.Path = "" Then
        MsgBox "This document hasn't been saved yet." & vbCrLf & vbCrLf & _
               "Please save it to NetDocuments first, then try again.", _
               vbExclamation, "Email a Copy"
        Exit Sub
    End If

    Dim response As VbMsgBoxResult
    response = MsgBox("Save latest changes before emailing?" & vbCrLf & vbCrLf & _
                      "Yes = Save first, then email" & vbCrLf & _
                      "No = Email the last saved version", _
                      vbYesNoCancel + vbQuestion, "Email a Copy")

    If response = vbCancel Then Exit Sub

    If response = vbYes Then
        On Error Resume Next
        ActiveDocument.Save
        If Err.Number <> 0 Then
            MsgBox "Could not save the document." & vbCrLf & _
                   "Error " & Err.Number & ": " & Err.Description, _
                   vbExclamation, "Email a Copy"
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    End If

    ' Build a clean name without the ND doc number
    Dim docName As String
    docName = GetDocNameFromTitleBar()
    If docName = "" Then
        docName = ActiveDocument.Name
        Dim dp As Long
        dp = InStrRev(docName, ".")
        If dp > 1 Then docName = Left(docName, dp - 1)
    End If

    ' Preserve the original file extension
    Dim ext As String
    Dim dotPos As Long
    dotPos = InStrRev(ActiveDocument.Name, ".")
    If dotPos > 0 Then ext = Mid(ActiveDocument.Name, dotPos)

    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")
    Dim olMail As Object
    Set olMail = olApp.CreateItem(0)

    ' Attach the file with a clean display name (no ND doc number).
    ' The 4th parameter of Attachments.Add sets the name the recipient sees.
    Dim cleanName As String
    cleanName = SanitizeFileName(docName) & ext
    If cleanName = ext Or cleanName = "" Then cleanName = ActiveDocument.Name
    olMail.Attachments.Add ActiveDocument.FullName, , , cleanName

    olMail.Subject = docName
    olMail.Display

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred creating the email." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Email a Copy"
End Sub

' =============================================================================
' Email a NetDocuments link to the current document
' =============================================================================
Public Sub EmailDocLink()
    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar(includeVersion:=False)

    If docId = "" Then
        MsgBox "Could not find a NetDocuments reference in the title bar." & vbCrLf & vbCrLf & _
               "Make sure the document is saved to NetDocuments first.", _
               vbExclamation, "Email Link"
        Exit Sub
    End If

    Dim ndUrl As String
    ndUrl = "https://eu.netdocuments.com/neWeb2/searchRes.aspx?newSearch=E&filter=%3D999%28" & docId & "%29&open=1"

    Dim docName As String
    docName = GetDocNameFromTitleBar()
    If docName = "" Then docName = docId

    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")
    Dim olMail As Object
    Set olMail = olApp.CreateItem(0)

    olMail.Subject = docName
    olMail.HTMLBody = BuildNDLinkCard(docName, ndUrl)
    olMail.Display

    ' Use Outlook's WordEditor to place the cursor below the table,
    ' so pressing Enter doesn't extend the card
    On Error Resume Next
    Dim wdDoc As Object
    Set wdDoc = olMail.GetInspector.WordEditor
    If Not wdDoc Is Nothing Then
        Dim rng As Object
        Set rng = wdDoc.Content
        rng.Collapse 0
        rng.Select
    End If
    On Error GoTo ErrorHandler

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred creating the email." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Email Link"
End Sub
