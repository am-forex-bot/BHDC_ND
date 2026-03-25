Attribute VB_Name = "modAutoMigrate"
Option Explicit

' =============================================================================
' modAutoMigrate - Auto-migrate iManage doc refs to NetDocuments on document open
' =============================================================================
' Import this module into Normal.dotm (NOT the BigHand template).
' Normal.dotm is always loaded by Word, so AutoOpen always fires.
'
' How it works:
'   1. AutoOpen fires when any document is opened
'   2. After a 2-second delay (gives ndOffice time to update the title bar),
'      MigrateIManageToND checks the title bar for a NetDocuments reference
'   3. If found, scans all footers for iManage doc numbers (e.g. 5973487v1)
'      and replaces them with the NetDocuments reference
'   4. Preserves all other footer content (logos, disclaimers, etc.)
'
' To install:
'   1. Open Word
'   2. Press Alt+F11 to open the VBA Editor
'   3. In the Project Explorer, find "Normal" (Normal.dotm)
'   4. Right-click Normal > Import File > select this .bas file
'   5. Close the VBA Editor
'   6. Word will auto-save Normal.dotm on exit
' =============================================================================

Public Sub AutoOpen()
    On Error Resume Next
    Application.OnTime When:=Now + TimeValue("00:00:02"), Name:="MigrateIManageToND"
End Sub

Public Sub MigrateIManageToND()
    ' Checks if the active document has a NetDocuments ref in the title bar.
    ' If so, scans footers for iManage doc numbers and replaces them.

    On Error GoTo ErrorHandler

    ' --- Get ND reference from title bar ---
    Dim titleText As String
    titleText = Application.ActiveWindow.Caption

    Dim ndRegex As Object
    Set ndRegex = CreateObject("VBScript.RegExp")
    ndRegex.Global = False
    ndRegex.IgnoreCase = True
    ndRegex.Pattern = "([A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4})\s*(\.\d+|v\.?\d+)?"

    If Not ndRegex.Test(titleText) Then Exit Sub

    Dim ndMatches As Object
    Set ndMatches = ndRegex.Execute(titleText)
    Dim docId As String
    docId = ndMatches(0).SubMatches(0)
    If Len(ndMatches(0).SubMatches(1)) > 0 Then
        docId = docId & ndMatches(0).SubMatches(1)
    End If

    ' --- Build iManage regex ---
    Dim imRegex As Object
    Set imRegex = CreateObject("VBScript.RegExp")
    imRegex.Global = True
    imRegex.IgnoreCase = True
    imRegex.Pattern = "\d{5,}(v\d+)?"

    ' --- Scan and replace in footers ---
    Dim sec As Section
    Set sec = ActiveDocument.Sections(1)

    Dim para As Paragraph

    ' Check primary footer
    For Each para In sec.Footers(wdHeaderFooterPrimary).Range.Paragraphs
        If imRegex.Test(para.Range.Text) Then
            ReplaceMigratePara para, docId
            Exit For
        End If
    Next para

    ' Check first page footer if Different First Page is enabled
    If sec.PageSetup.DifferentFirstPageHeaderFooter Then
        For Each para In sec.Footers(wdHeaderFooterFirstPage).Range.Paragraphs
            If imRegex.Test(para.Range.Text) Then
                ReplaceMigratePara para, docId
                Exit For
            End If
        Next para
    End If

    Exit Sub

ErrorHandler:
    ' Silently fail - don't interrupt the user opening their document
End Sub

Private Sub ReplaceMigratePara(para As Paragraph, docId As String)
    Dim rng As Range
    Set rng = para.Range
    rng.MoveEnd wdCharacter, -1
    rng.Text = docId
    rng.Font.Size = 8
    rng.Font.Color = RGB(128, 128, 128)
    rng.ParagraphFormat.Alignment = wdAlignParagraphRight
End Sub
