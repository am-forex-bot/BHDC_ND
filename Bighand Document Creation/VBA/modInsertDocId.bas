Attribute VB_Name = "modInsertDocId"
Option Explicit

' =============================================================================
' modInsertDocId - NetDocuments Document ID functions
' =============================================================================
' InsertNDDocIdAllPages      - Insert ND ref into footer on all pages
' InsertNDDocIdFirstOnly     - Insert ND ref into first page footer only
' InsertNDDocIdAllButFirst   - Insert ND ref into footer on all pages except first
' InsertNDDocNum             - Insert ND ref at cursor position
' =============================================================================

Private Function GetNDDocIdFromTitleBar() As String
    ' When a document is open via ndOffice, the title bar contains the ND reference
    ' Typical format: "DocumentName - 1234-5678-9012.1 - NetDocuments"
    ' Or it may appear as: "DocumentName [1234-5678-9012] - Word"

    Dim titleText As String
    Dim docId As String

    titleText = Application.ActiveWindow.Caption

    ' Try to extract ND document ID pattern (xxxx-xxxx-xxxx format)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True

    ' Pattern: 4 alphanumeric chars, dash, 4, dash, 4 (optionally followed by .version)
    regex.Pattern = "([A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4})(\.\d+)?"

    If regex.Test(titleText) Then
        Dim matches As Object
        Set matches = regex.Execute(titleText)
        docId = matches(0).Value
    End If

    GetNDDocIdFromTitleBar = docId
End Function

Private Sub InsertOrReplaceInFooter(ftr As HeaderFooter, docId As String)
    ' Inserts or replaces an ND ref in the given footer, right-aligned, 8pt

    Dim ftrRange As Range
    Set ftrRange = ftr.Range

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}(\.\d+)?"

    If regex.Test(ftrRange.Text) Then
        ' Replace existing ND ref in footer using VBScript.RegExp
        Dim newText As String
        newText = regex.Replace(ftrRange.Text, docId)
        ftrRange.Text = newText
    Else
        ' Append to footer
        Dim insertRange As Range
        Set insertRange = ftr.Range
        insertRange.Collapse wdCollapseEnd
        insertRange.InsertAfter vbCrLf & docId
    End If

    ' Right-align and set consistent font size (8pt) for the footer
    ftr.Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    ftr.Range.Font.Size = 8
End Sub

Public Sub InsertNDDocIdAllPages()
    ' Insert ND ref into footer on all pages (primary + first page if enabled)

    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar()

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

    ' Insert into primary footer (applies to all pages, or non-first pages if Different First Page is on)
    InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), docId

    ' If Different First Page is enabled, also insert into first page footer
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
    ' Insert ND ref into first page footer only

    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar()

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

    ' Enable Different First Page if not already
    If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
        sec.PageSetup.DifferentFirstPageHeaderFooter = True
    End If

    ' Insert into first page footer only
    InsertOrReplaceInFooter sec.Footers(wdHeaderFooterFirstPage), docId

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
    ' Insert ND ref into footer on all pages except the first

    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar()

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

    ' Enable Different First Page if not already
    If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then
        sec.PageSetup.DifferentFirstPageHeaderFooter = True
    End If

    ' Insert into primary footer only (skips first page)
    InsertOrReplaceInFooter sec.Footers(wdHeaderFooterPrimary), docId

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
