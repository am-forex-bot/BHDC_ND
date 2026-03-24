Attribute VB_Name = "modInsertDocId"
Option Explicit

' =============================================================================
' modInsertDocId - NetDocuments Document ID functions
' =============================================================================
' InsertNDDocId    - Reads ND doc ID from title bar, inserts into footer
' InsertNDDocNum   - Reads ND doc ID from title bar, inserts at cursor position
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

Public Sub InsertNDDocId()
    ' Reads ND document reference from title bar and inserts into the footer

    On Error GoTo ErrorHandler

    Dim docId As String
    docId = GetNDDocIdFromTitleBar()

    If docId = "" Then
        MsgBox "Could not find a NetDocuments reference in the title bar." & vbCrLf & vbCrLf & _
               "Make sure the document is saved to NetDocuments first.", _
               vbExclamation, "Insert ND Ref"
        Exit Sub
    End If

    ' Store current cursor position
    Dim cursorPos As Range
    Set cursorPos = Selection.Range

    ' Insert into primary footer of section 1
    Dim sec As Section
    Set sec = ActiveDocument.Sections(1)

    Dim ftr As HeaderFooter
    Set ftr = sec.Footers(wdHeaderFooterPrimary)

    ' Check if footer already contains an ND ref and replace it
    Dim ftrRange As Range
    Set ftrRange = ftr.Range

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}(\.\d+)?"

    If regex.Test(ftrRange.Text) Then
        ' Replace existing ND ref in footer
        Dim findObj As Find
        Set findObj = ftrRange.Find
        findObj.ClearFormatting
        findObj.Replacement.ClearFormatting
        findObj.Text = ""
        findObj.MatchWildcards = True
        findObj.Text = "[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}"
        findObj.Replacement.Text = docId
        findObj.Execute Replace:=wdReplaceAll
    Else
        ' Append to footer
        Dim insertRange As Range
        Set insertRange = ftr.Range
        insertRange.Collapse wdCollapseEnd
        insertRange.InsertAfter vbCrLf & docId
    End If

    ' Restore cursor position
    cursorPos.Select

    MsgBox "NetDocuments reference " & docId & " inserted into footer.", _
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
