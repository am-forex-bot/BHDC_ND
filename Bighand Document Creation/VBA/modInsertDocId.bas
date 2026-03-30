Attribute VB_Name = "modInsertDocId"
Option Explicit

' =============================================================================
' modInsertDocId - NetDocuments Document ID footer functions
' =============================================================================
' 18 public subs for ribbon: version(2) x pages(3) x alignment(3)
' Naming: NDFooter_{Ver|Num}_{All|First|NotFirst}_{L|C|R}
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

Private Sub ReplaceParaText(para As Paragraph, docId As String, align As Long)
    Dim rng As Range
    Set rng = para.Range
    rng.MoveEnd wdCharacter, -1
    rng.Text = docId
    rng.Font.Size = 8
    rng.Font.Color = RGB(128, 128, 128)
    rng.ParagraphFormat.Alignment = align
End Sub

Private Sub InsertOrReplaceInFooter(ftr As HeaderFooter, docId As String, align As Long)
    Dim ndRegex As Object
    Set ndRegex = CreateNDRegex()
    Dim imRegex As Object
    Set imRegex = CreateIManageRegex()

    Dim para As Paragraph
    For Each para In ftr.Range.Paragraphs
        If ndRegex.Test(para.Range.Text) Or imRegex.Test(para.Range.Text) Then
            ReplaceParaText para, docId, align
            Exit Sub
        End If
    Next para

    ' No existing ref found - append new paragraph
    Dim insertRange As Range
    Set insertRange = ftr.Range
    insertRange.Collapse wdCollapseEnd
    insertRange.Text = vbCr & docId
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
End Sub

' --- Shared implementation ---

Private Sub DoInsertFooter(includeVersion As Boolean, pageMode As String, align As String)
    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar(includeVersion)

    If docId = "" Then
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
            InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), docId, al
            If sec.PageSetup.DifferentFirstPageHeaderFooter Then
                InsertOrReplaceInFooter sec.Footers(wdHeaderFooterFirstPage), docId, al
            End If

        Case "first"
            If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
                sec.PageSetup.DifferentFirstPageHeaderFooter = True
            End If
            InsertOrReplaceInFooter sec.Footers(wdHeaderFooterFirstPage), docId, al
            RemoveDocRefFromFooter sec.Footers(wdHeaderFooterPrimary)

        Case "notfirst"
            If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
                sec.PageSetup.DifferentFirstPageHeaderFooter = True
            End If
            InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), docId, al
            RemoveDocRefFromFooter sec.Footers(wdHeaderFooterFirstPage)
    End Select

    cursorPos.Select

    Dim pageTxt As String
    Select Case LCase(pageMode)
        Case "all": pageTxt = "all pages"
        Case "first": pageTxt = "first page footer"
        Case "notfirst": pageTxt = "all pages except first"
    End Select

    MsgBox "NetDocuments reference " & docId & " inserted into " & pageTxt & ".", _
           vbInformation, "Insert ND Ref"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred inserting the ND reference." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Insert ND Ref"
End Sub

' =============================================================================
' With Version - All Pages
' =============================================================================
Public Sub NDFooter_Ver_All_L()
    DoInsertFooter True, "all", "left"
End Sub

Public Sub NDFooter_Ver_All_C()
    DoInsertFooter True, "all", "center"
End Sub

Public Sub NDFooter_Ver_All_R()
    DoInsertFooter True, "all", "right"
End Sub

' =============================================================================
' With Version - First Page Only
' =============================================================================
Public Sub NDFooter_Ver_First_L()
    DoInsertFooter True, "first", "left"
End Sub

Public Sub NDFooter_Ver_First_C()
    DoInsertFooter True, "first", "center"
End Sub

Public Sub NDFooter_Ver_First_R()
    DoInsertFooter True, "first", "right"
End Sub

' =============================================================================
' With Version - All Except First
' =============================================================================
Public Sub NDFooter_Ver_NotFirst_L()
    DoInsertFooter True, "notfirst", "left"
End Sub

Public Sub NDFooter_Ver_NotFirst_C()
    DoInsertFooter True, "notfirst", "center"
End Sub

Public Sub NDFooter_Ver_NotFirst_R()
    DoInsertFooter True, "notfirst", "right"
End Sub

' =============================================================================
' Number Only - All Pages
' =============================================================================
Public Sub NDFooter_Num_All_L()
    DoInsertFooter False, "all", "left"
End Sub

Public Sub NDFooter_Num_All_C()
    DoInsertFooter False, "all", "center"
End Sub

Public Sub NDFooter_Num_All_R()
    DoInsertFooter False, "all", "right"
End Sub

' =============================================================================
' Number Only - First Page Only
' =============================================================================
Public Sub NDFooter_Num_First_L()
    DoInsertFooter False, "first", "left"
End Sub

Public Sub NDFooter_Num_First_C()
    DoInsertFooter False, "first", "center"
End Sub

Public Sub NDFooter_Num_First_R()
    DoInsertFooter False, "first", "right"
End Sub

' =============================================================================
' Number Only - All Except First
' =============================================================================
Public Sub NDFooter_Num_NotFirst_L()
    DoInsertFooter False, "notfirst", "left"
End Sub

Public Sub NDFooter_Num_NotFirst_C()
    DoInsertFooter False, "notfirst", "center"
End Sub

Public Sub NDFooter_Num_NotFirst_R()
    DoInsertFooter False, "notfirst", "right"
End Sub

' =============================================================================
' Legacy names (keep for backward compatibility with existing ribbon tags)
' =============================================================================
Public Sub InsertNDDocIdAllPages()
    DoInsertFooter True, "all", "right"
End Sub

Public Sub InsertNDDocIdFirstOnly()
    DoInsertFooter True, "first", "right"
End Sub

Public Sub InsertNDDocIdAllButFirst()
    DoInsertFooter True, "notfirst", "right"
End Sub

Public Sub InsertNDDocIdAllPagesNoVer()
    DoInsertFooter False, "all", "right"
End Sub

Public Sub InsertNDDocIdFirstOnlyNoVer()
    DoInsertFooter False, "first", "right"
End Sub

Public Sub InsertNDDocIdAllButFirstNoVer()
    DoInsertFooter False, "notfirst", "right"
End Sub

' =============================================================================
' Refresh footer version
' =============================================================================
Public Sub RefreshFooterVersion()
    ' Updates existing ND ref in footer to match current title bar version.
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
            ReplaceParaText para, docId, existingAlign
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
                ReplaceParaText para, docId, existingAlign2
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
