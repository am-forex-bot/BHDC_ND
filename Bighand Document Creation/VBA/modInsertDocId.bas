Attribute VB_Name = "modInsertDocId"
Option Explicit

' =============================================================================
' modInsertDocId - NetDocuments Document ID footer functions
' =============================================================================
' 45 public subs for ribbon: format(5) x pages(3) x alignment(3)
' Formats: Num, Ver, Name, NameNum, NameNumVer
' Naming: NDFooter_{format}_{All|First|NotFirst}_{L|C|R}
' Plus: InsertNDDocNum - insert at cursor position
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

    ' Strip trailing " - Microsoft Word" (or localized app name)
    Dim appSuffix As String
    appSuffix = " - " & Application.Name
    If Len(titleText) > Len(appSuffix) And _
       StrComp(Right(titleText, Len(appSuffix)), appSuffix, vbTextCompare) = 0 Then
        titleText = Left(titleText, Len(titleText) - Len(appSuffix))
    End If

    ' Remove ND doc ID and optional version
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}\s*(\.\d+|v\.?\d+)?"
    titleText = regex.Replace(titleText, "")

    ' Clean up residual separators and whitespace
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
            ' Verify doc is in NetDocuments before extracting name
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
    ' Detect if the existing footer has a name prefix before the doc ID
    ' and rebuild accordingly, preserving the name
    Dim ndRegex As Object
    Set ndRegex = CreateNDRegex()

    Dim cleanText As String
    cleanText = existingText
    ' Remove trailing paragraph mark
    If Len(cleanText) > 0 Then
        If Asc(Right(cleanText, 1)) = 13 Then
            cleanText = Left(cleanText, Len(cleanText) - 1)
        End If
    End If

    If ndRegex.Test(cleanText) Then
        Dim matches As Object
        Set matches = ndRegex.Execute(cleanText)
        If matches(0).FirstIndex > 0 Then
            ' Text exists before the doc ID - rebuild with fresh name
            Dim docName As String
            docName = GetDocNameFromTitleBar()
            If docName <> "" Then
                BuildRefreshText = docName & " - " & newDocId
                Exit Function
            End If
        End If
    End If

    ' No name prefix - just use the doc ID
    BuildRefreshText = newDocId
End Function

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

    ' First pass: check for ND/iManage doc ID patterns
    Dim para As Paragraph
    For Each para In ftr.Range.Paragraphs
        If ndRegex.Test(para.Range.Text) Or imRegex.Test(para.Range.Text) Then
            ReplaceParaText para, footerText, align
            Exit Sub
        End If
    Next para

    ' Second pass: check for previously inserted ref (size 8, grey text)
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

    ' Also check for previously inserted ref (name-only mode: size 8, grey)
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

    Dim sec As Section
    Set sec = ActiveDocument.Sections(1)

    Dim al As Long
    al = AlignmentConst(align)

    Select Case LCase(pageMode)
        Case "all"
            InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), footerText, al
            If sec.PageSetup.DifferentFirstPageHeaderFooter Then
                InsertOrReplaceInFooter sec.Footers(wdHeaderFooterFirstPage), footerText, al
            End If

        Case "first"
            If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
                sec.PageSetup.DifferentFirstPageHeaderFooter = True
            End If
            InsertOrReplaceInFooter sec.Footers(wdHeaderFooterFirstPage), footerText, al
            RemoveDocRefFromFooter sec.Footers(wdHeaderFooterPrimary)

        Case "notfirst"
            If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
                sec.PageSetup.DifferentFirstPageHeaderFooter = True
            End If
            InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), footerText, al
            RemoveDocRefFromFooter sec.Footers(wdHeaderFooterFirstPage)
    End Select

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
' Refresh footer version
' =============================================================================
Public Sub RefreshFooterVersion()
    ' Updates existing ND ref in footer to match current title bar version.
    ' Preserves doc name prefix if present.
    ' Use after Save As / New Version in NetDocuments.

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
    Set sec = ActiveDocument.Sections(1)
    Dim updated As Boolean
    updated = False

    ' Check primary footer
    Dim para As Paragraph
    For Each para In sec.Footers(wdHeaderFooterPrimary).Range.Paragraphs
        If ndRegex.Test(para.Range.Text) Or imRegex.Test(para.Range.Text) Then
            Dim existingAlign As Long
            existingAlign = para.Range.ParagraphFormat.Alignment
            Dim refreshText As String
            refreshText = BuildRefreshText(para.Range.Text, docId)
            ReplaceParaText para, refreshText, existingAlign
            updated = True
            Exit For
        End If
    Next para

    ' Check first page footer
    If sec.PageSetup.DifferentFirstPageHeaderFooter Then
        For Each para In sec.Footers(wdHeaderFooterFirstPage).Range.Paragraphs
            If ndRegex.Test(para.Range.Text) Or imRegex.Test(para.Range.Text) Then
                Dim existingAlign2 As Long
                existingAlign2 = para.Range.ParagraphFormat.Alignment
                Dim refreshText2 As String
                refreshText2 = BuildRefreshText(para.Range.Text, docId)
                ReplaceParaText para, refreshText2, existingAlign2
                updated = True
                Exit For
            End If
        Next para
    End If

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
