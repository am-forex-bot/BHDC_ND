Attribute VB_Name = "modInsertDocId"
Option Explicit

' =============================================================================
' Module:  modInsertDocId
' Purpose: Extracts the NetDocuments envelope ID and version from the Word
'          window caption and inserts it into the document's primary footer.
'
' The ND title bar format is:
'   "filename.docx 4162-3353-2006 v.2 [other stuff]"
'
' Pattern: XXXX-XXXX-XXXX (4 digits, dash, 4 digits, dash, 4 digits)
' Version: v.N (where N is one or more digits)
'
' To install: Import this .bas file into Wallace Shared Code.dotm
' =============================================================================

Public Sub InsertNDDocId()

    Dim sCaption    As String
    Dim sDocNumber  As String
    Dim sVersion    As String
    Dim sDocId      As String
    Dim oFooter     As HeaderFooter
    Dim rng         As Range

    ' Get the window caption
    sCaption = ActiveWindow.Caption

    ' Extract the ND envelope ID (XXXX-XXXX-XXXX) and version (v.N)
    sDocNumber = ExtractNDNumber(sCaption)
    sVersion = ExtractNDVersion(sCaption)

    ' Validate we found something
    If sDocNumber = "" Then
        MsgBox "No NetDocuments reference found in the document title." & vbCrLf & vbCrLf & _
               "Please save the document to NetDocuments first, then try again." & vbCrLf & vbCrLf & _
               "Expected format in title bar: XXXX-XXXX-XXXX v.N", _
               vbExclamation, "Insert Doc Reference"
        Exit Sub
    End If

    ' Build the formatted doc ID string
    If sVersion <> "" Then
        sDocId = sDocNumber & "." & sVersion
    Else
        sDocId = sDocNumber & ".1"
    End If

    ' Insert into the primary footer of the first section
    On Error GoTo ErrHandler

    Set oFooter = ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary)

    ' Check if footer already has an ND reference and update it,
    ' otherwise append to existing footer content
    If ReplaceExistingDocId(oFooter, sDocId) Then
        ' Successfully replaced existing reference
    Else
        ' No existing reference found - append to footer
        AppendDocIdToFooter oFooter, sDocId
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error inserting document reference: " & Err.Description, _
           vbCritical, "Insert Doc Reference"

End Sub

' -----------------------------------------------------------------------------
' Extracts the XXXX-XXXX-XXXX pattern from a string
' Returns empty string if not found
' -----------------------------------------------------------------------------
Private Function ExtractNDNumber(ByVal sText As String) As String

    Dim i       As Long
    Dim sChunk  As String

    ' Walk through the string looking for the pattern ####-####-####
    ' Minimum length needed: 14 characters (4-4-4 + 2 dashes)
    If Len(sText) < 14 Then
        ExtractNDNumber = ""
        Exit Function
    End If

    For i = 1 To Len(sText) - 13
        sChunk = Mid$(sText, i, 14)
        If IsNDNumber(sChunk) Then
            ' Make sure the character before (if any) is not a digit
            ' to avoid partial matches
            If i > 1 Then
                If IsNumeric(Mid$(sText, i - 1, 1)) Then GoTo NextChar
            End If
            ' Make sure the character after (if any) is not a digit
            If i + 14 <= Len(sText) Then
                If IsNumeric(Mid$(sText, i + 14, 1)) Then GoTo NextChar
            End If
            ExtractNDNumber = sChunk
            Exit Function
        End If
NextChar:
    Next i

    ExtractNDNumber = ""

End Function

' -----------------------------------------------------------------------------
' Checks if a 14-character string matches ####-####-####
' -----------------------------------------------------------------------------
Private Function IsNDNumber(ByVal s As String) As Boolean

    Dim j As Long

    IsNDNumber = False
    If Len(s) <> 14 Then Exit Function

    For j = 1 To 14
        Select Case j
            Case 5, 10  ' Dash positions
                If Mid$(s, j, 1) <> "-" Then Exit Function
            Case Else   ' Digit positions
                If Not IsNumeric(Mid$(s, j, 1)) Then Exit Function
        End Select
    Next j

    IsNDNumber = True

End Function

' -----------------------------------------------------------------------------
' Extracts the version number from "v.N" or "v.NN" pattern
' Returns just the number portion (e.g. "2" from "v.2")
' Returns empty string if not found
' -----------------------------------------------------------------------------
Private Function ExtractNDVersion(ByVal sText As String) As String

    Dim i       As Long
    Dim sAfter  As String
    Dim sNum    As String
    Dim ch      As String

    ' Look for " v." pattern (space before v.)
    i = InStr(1, sText, " v.", vbTextCompare)

    If i = 0 Then
        ExtractNDVersion = ""
        Exit Function
    End If

    ' Get everything after " v."
    sAfter = Mid$(sText, i + 3)

    ' Collect consecutive digits
    sNum = ""
    Dim k As Long
    For k = 1 To Len(sAfter)
        ch = Mid$(sAfter, k, 1)
        If IsNumeric(ch) Then
            sNum = sNum & ch
        Else
            Exit For
        End If
    Next k

    ExtractNDVersion = sNum

End Function

' -----------------------------------------------------------------------------
' Searches footer for existing ND reference pattern and replaces it
' Returns True if a replacement was made
' -----------------------------------------------------------------------------
Private Function ReplaceExistingDocId(oFooter As HeaderFooter, sNewDocId As String) As Boolean

    Dim rng     As Range
    Dim sText   As String
    Dim i       As Long
    Dim iStart  As Long
    Dim iEnd    As Long

    ReplaceExistingDocId = False

    Set rng = oFooter.Range
    sText = rng.Text

    ' Look for an existing ND number pattern in the footer
    For i = 1 To Len(sText) - 13
        If IsNDNumber(Mid$(sText, i, 14)) Then
            iStart = i
            ' Find the end (include .version if present)
            iEnd = i + 13  ' End of XXXX-XXXX-XXXX
            ' Check for .N version suffix
            If iEnd < Len(sText) Then
                If Mid$(sText, iEnd + 1, 1) = "." Then
                    iEnd = iEnd + 1
                    ' Collect version digits
                    Do While iEnd < Len(sText)
                        If IsNumeric(Mid$(sText, iEnd + 1, 1)) Then
                            iEnd = iEnd + 1
                        Else
                            Exit Do
                        End If
                    Loop
                End If
            End If

            ' Replace the text
            Dim rngReplace As Range
            Set rngReplace = oFooter.Range
            rngReplace.Start = rng.Start + iStart - 1
            rngReplace.End = rng.Start + iEnd
            rngReplace.Text = sNewDocId

            ReplaceExistingDocId = True
            Exit Function
        End If
    Next i

End Function

' -----------------------------------------------------------------------------
' Appends the doc ID to the footer on a new line
' Uses the same font/style as existing footer text
' -----------------------------------------------------------------------------
Private Sub AppendDocIdToFooter(oFooter As HeaderFooter, sDocId As String)

    Dim rng As Range
    Set rng = oFooter.Range

    ' If footer already has content, go to the end
    rng.Collapse wdCollapseEnd

    ' Check if footer is empty
    If Len(Trim$(oFooter.Range.Text)) <= 1 Then
        ' Empty footer - just insert the doc ID
        rng.Text = sDocId
    Else
        ' Has content - add on new line
        rng.Text = vbCr & sDocId
    End If

    ' Format: use a smaller font size typical for doc references
    rng.Font.Size = 8
    rng.Font.Color = wdColorGray50

End Sub
