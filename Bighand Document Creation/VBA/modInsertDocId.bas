Attribute VB_Name = "modInsertDocId"
Option Explicit

' =============================================================================
' modInsertDocId - NetDocuments Document ID functions
' =============================================================================
' AutoExec / CheckActiveDocForMigration - Auto-replace iManage refs with ND ref on open
' InsertNDDocIdAllPages          - Insert ND ref (with version) into footer on all pages
' InsertNDDocIdFirstOnly         - Insert ND ref (with version) into first page footer only
' InsertNDDocIdAllButFirst       - Insert ND ref (with version) into footer except first
' InsertNDDocIdAllPagesNoVer     - Insert ND ref (number only) into footer on all pages
' InsertNDDocIdFirstOnlyNoVer    - Insert ND ref (number only) into first page footer only
' InsertNDDocIdAllButFirstNoVer  - Insert ND ref (number only) into footer except first
' InsertNDDocNum                 - Insert ND ref at cursor position
' =============================================================================

Private Function GetNDDocIdFromTitleBar(Optional includeVersion As Boolean = True) As String
    ' When a document is open via ndOffice, the title bar contains the ND reference
    ' Formats seen:
    '   "DocumentName - 1234-5678-9012 v.1 - NetDocuments"
    '   "DocumentName - 1234-5678-9012.1 - NetDocuments"
    '   "DocumentName [1234-5678-9012] - Word"

    Dim titleText As String
    titleText = Application.ActiveWindow.Caption

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True

    ' Pattern: xxxx-xxxx-xxxx optionally followed by version in various formats:
    '   .1   v1   v.1   (space)v.1   (space).1
    regex.Pattern = "([A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4})\s*(\.\d+|v\.?\d+)?"

    If regex.Test(titleText) Then
        Dim matches As Object
        Set matches = regex.Execute(titleText)
        Dim baseId As String
        baseId = matches(0).SubMatches(0)

        If includeVersion And Len(matches(0).SubMatches(1)) > 0 Then
            ' Join base ID and version with no space
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
    ' Match ND ref in footer: xxxx-xxxx-xxxx optionally followed by version
    regex.Pattern = "[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}\s*(\.\d+|v\.?\d+)?"
    Set CreateNDRegex = regex
End Function

Private Function CreateIManageRegex() As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    ' Match iManage ref: 5+ digits optionally followed by v and version number (e.g. 5973487v1)
    regex.Pattern = "\d{5,}(v\d+)?"
    Set CreateIManageRegex = regex
End Function

Private Sub ReplaceParaText(para As Paragraph, docId As String)
    ' Replaces the text content of a paragraph without touching the paragraph mark
    ' This preserves the paragraph structure and position in the footer
    Dim rng As Range
    Set rng = para.Range
    ' Shrink range to exclude the trailing paragraph mark
    rng.MoveEnd wdCharacter, -1
    rng.Text = docId
    ' Format the replaced text
    rng.Font.Size = 8
    rng.Font.Color = RGB(128, 128, 128)
    rng.ParagraphFormat.Alignment = wdAlignParagraphRight
End Sub

Private Sub InsertOrReplaceInFooter(ftr As HeaderFooter, docId As String)
    ' Inserts or replaces an ND or iManage ref in the given footer
    ' Preserves all other content (logo, disclaimer, etc.)

    Dim ndRegex As Object
    Set ndRegex = CreateNDRegex()
    Dim imRegex As Object
    Set imRegex = CreateIManageRegex()

    ' Loop through paragraphs to find one containing an ND ref or iManage ref
    Dim para As Paragraph
    For Each para In ftr.Range.Paragraphs
        If ndRegex.Test(para.Range.Text) Or imRegex.Test(para.Range.Text) Then
            ReplaceParaText para, docId
            Exit Sub
        End If
    Next para

    ' No existing ref found - append on a new paragraph at the end
    Dim insertRange As Range
    Set insertRange = ftr.Range
    insertRange.Collapse wdCollapseEnd
    insertRange.Text = vbCr & docId
    insertRange.Font.Size = 8
    insertRange.Font.Color = RGB(128, 128, 128)
    insertRange.ParagraphFormat.Alignment = wdAlignParagraphRight
End Sub

Private Sub RemoveDocRefFromFooter(ftr As HeaderFooter)
    ' Removes an ND or iManage ref paragraph from a footer
    ' Preserves all other content (logo, disclaimer, etc.)

    Dim ndRegex As Object
    Set ndRegex = CreateNDRegex()
    Dim imRegex As Object
    Set imRegex = CreateIManageRegex()

    ' Loop through paragraphs to find and delete the one with a doc ref
    Dim para As Paragraph
    For Each para In ftr.Range.Paragraphs
        If ndRegex.Test(para.Range.Text) Or imRegex.Test(para.Range.Text) Then
            para.Range.Delete
            Exit Sub
        End If
    Next para
End Sub

' =============================================================================
' Auto-migration: monitors for new documents and migrates iManage refs
' BigHand loads templates via COM add-in, not Word's Startup mechanism,
' so AutoOpen doesn't fire. Instead we use AutoExec + polling.
' =============================================================================

' Tracks the last document we checked, so we only migrate once per document
Private mLastCheckedDoc As String

Public Sub AutoExec()
    ' Runs when the template is loaded into Word.
    ' Starts a document change monitor that checks every 3 seconds.
    mLastCheckedDoc = ""
    ScheduleDocCheck
End Sub

Private Sub ScheduleDocCheck()
    On Error Resume Next
    Application.OnTime When:=Now + TimeValue("00:00:03"), Name:="CheckActiveDocForMigration"
End Sub

Public Sub CheckActiveDocForMigration()
    ' Called every 3 seconds. If the active document has changed,
    ' checks for iManage refs and migrates them.

    On Error GoTo Reschedule

    Dim currentDoc As String
    currentDoc = ActiveDocument.FullName

    If currentDoc <> mLastCheckedDoc Then
        mLastCheckedDoc = currentDoc
        MigrateIManageFooter
    End If

Reschedule:
    ScheduleDocCheck
End Sub

Private Sub MigrateIManageFooter()
    ' If the current document is open via ndOffice (ND ref in title bar),
    ' scans all footers for iManage doc numbers and replaces them.

    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar(includeVersion:=True)

    ' Only proceed if this document is open via NetDocuments
    If docId = "" Then Exit Sub

    Dim imRegex As Object
    Set imRegex = CreateIManageRegex()

    Dim sec As Section
    Set sec = ActiveDocument.Sections(1)

    ' Check primary footer
    Dim para As Paragraph
    For Each para In sec.Footers(wdHeaderFooterPrimary).Range.Paragraphs
        If imRegex.Test(para.Range.Text) Then
            ReplaceParaText para, docId
            Exit For
        End If
    Next para

    ' Check first page footer if Different First Page is enabled
    If sec.PageSetup.DifferentFirstPageHeaderFooter Then
        For Each para In sec.Footers(wdHeaderFooterFirstPage).Range.Paragraphs
            If imRegex.Test(para.Range.Text) Then
                ReplaceParaText para, docId
                Exit For
            End If
        Next para
    End If

    Exit Sub

ErrorHandler:
    ' Silently fail - don't interrupt the user opening their document
End Sub

' =============================================================================
' Ribbon button handlers - With Version
' =============================================================================

Public Sub InsertNDDocIdAllPages()
    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar(includeVersion:=True)

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

    InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), docId

    If sec.PageSetup.DifferentFirstPageHeaderFooter Then
        InsertOrReplaceInFooter sec.Footers(wdHeaderFooterFirstPage), docId
    End If

    cursorPos.Select

    MsgBox "NetDocuments reference " & docId & " inserted into footer on all pages.", _
           vbInformation, "Insert ND Ref"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred inserting the ND reference." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Insert ND Ref"
End Sub

Public Sub InsertNDDocIdFirstOnly()
    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar(includeVersion:=True)

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

    If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
        sec.PageSetup.DifferentFirstPageHeaderFooter = True
    End If

    InsertOrReplaceInFooter sec.Footers(wdHeaderFooterFirstPage), docId
    RemoveDocRefFromFooter sec.Footers(wdHeaderFooterPrimary)

    cursorPos.Select

    MsgBox "NetDocuments reference " & docId & " inserted into first page footer.", _
           vbInformation, "Insert ND Ref"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred inserting the ND reference." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Insert ND Ref"
End Sub

Public Sub InsertNDDocIdAllButFirst()
    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar(includeVersion:=True)

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

    If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
        sec.PageSetup.DifferentFirstPageHeaderFooter = True
    End If

    InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), docId
    RemoveDocRefFromFooter sec.Footers(wdHeaderFooterFirstPage)

    cursorPos.Select

    MsgBox "NetDocuments reference " & docId & " inserted into footer on all pages except first.", _
           vbInformation, "Insert ND Ref"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred inserting the ND reference." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Insert ND Ref"
End Sub

' =============================================================================
' Ribbon button handlers - Number Only (no version suffix)
' =============================================================================

Public Sub InsertNDDocIdAllPagesNoVer()
    On Error GoTo ErrorHandler
    Dim docId As String
    docId = GetNDDocIdFromTitleBar(includeVersion:=False)
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
    InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), docId
    If sec.PageSetup.DifferentFirstPageHeaderFooter Then
        InsertOrReplaceInFooter sec.Footers(wdHeaderFooterFirstPage), docId
    End If
    cursorPos.Select
    MsgBox "NetDocuments reference " & docId & " inserted into footer on all pages.", _
           vbInformation, "Insert ND Ref"
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred inserting the ND reference." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Insert ND Ref"
End Sub

Public Sub InsertNDDocIdFirstOnlyNoVer()
    On Error GoTo ErrorHandler
    Dim docId As String
    docId = GetNDDocIdFromTitleBar(includeVersion:=False)
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
    If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
        sec.PageSetup.DifferentFirstPageHeaderFooter = True
    End If
    InsertOrReplaceInFooter sec.Footers(wdHeaderFooterFirstPage), docId
    RemoveDocRefFromFooter sec.Footers(wdHeaderFooterPrimary)
    cursorPos.Select
    MsgBox "NetDocuments reference " & docId & " inserted into first page footer.", _
           vbInformation, "Insert ND Ref"
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred inserting the ND reference." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Insert ND Ref"
End Sub

Public Sub InsertNDDocIdAllButFirstNoVer()
    On Error GoTo ErrorHandler
    Dim docId As String
    docId = GetNDDocIdFromTitleBar(includeVersion:=False)
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
    If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
        sec.PageSetup.DifferentFirstPageHeaderFooter = True
    End If
    InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), docId
    RemoveDocRefFromFooter sec.Footers(wdHeaderFooterFirstPage)
    cursorPos.Select
    MsgBox "NetDocuments reference " & docId & " inserted into footer on all pages except first.", _
           vbInformation, "Insert ND Ref"
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred inserting the ND reference." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Insert ND Ref"
End Sub

Public Sub InsertNDDocNum()
    ' Reads ND document number from title bar and inserts at current cursor position

    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar()

    If docId = "" Then
        MsgBox "Could not find a NetDocuments reference in the title bar." & vbCrLf & vbCrLf & _
               "Make sure the document is saved to NetDocuments first.", _
               vbExclamation, "Insert ND Doc Number"
        Exit Sub
    End If

    ' Insert at current cursor position
    Selection.TypeText Text:=docId

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred inserting the ND document number." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Insert ND Doc Number"
End Sub
