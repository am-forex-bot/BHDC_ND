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

    Dim sectionCount As Long
    sectionCount = ActiveDocument.Sections.Count

    If sectionCount = 1 Then
        ApplyFooterToSection ActiveDocument.Sections(1), footerText, pageMode, al
    Else
        Dim info As String
        info = "This document has " & sectionCount & " sections:" & vbCrLf & vbCrLf
        info = info & GetSectionInfo()
        info = info & vbCrLf & "Enter sections to update (e.g. ""1,3"" or ""2-4"" or ""all""):"

        Dim userInput As String
        userInput = InputBox(info, "Select Sections", "all")

        If Len(userInput) = 0 Then Exit Sub

        Dim selectedSections As Collection
        Set selectedSections = ParseSectionSelection(userInput, sectionCount)

        If selectedSections.Count = 0 Then
            MsgBox "No valid sections selected.", vbExclamation, "Insert ND Ref"
            Exit Sub
        End If

        Dim idx As Variant
        For Each idx In selectedSections
            ApplyFooterToSection ActiveDocument.Sections(CLng(idx)), footerText, pageMode, al
        Next idx
    End If

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
