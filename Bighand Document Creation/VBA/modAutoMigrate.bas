Attribute VB_Name = "modAutoMigrate"
Option Explicit

' =============================================================================
' modAutoMigrate - Auto-update footer doc refs on open and save
' =============================================================================
' Import this module into Normal.dotm (NOT the BigHand template).
' Normal.dotm is always loaded by Word, so auto macros always fire.
'
' Hooks:
'   AutoOpen   - fires when a document is opened
'   FileSave   - fires on Ctrl+S / Save
'   FileSaveAs - fires on Save As (e.g. creating a new version in ND)
'
' After each event, schedules MigrateIManageToND with a short delay
' to give ndOffice time to update the title bar, then:
'   - Replaces iManage doc numbers (e.g. 5973487v1) with ND ref
'   - Updates outdated ND version numbers (e.g. v.1 -> v.2)
'   - Preserves all other footer content and alignment
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

Public Sub FileSave()
    ' Let the normal save happen first, then check for version updates
    On Error Resume Next
    ActiveDocument.Save
    Application.OnTime Now + TimeValue("00:00:03"), "MigrateIManageToND"
End Sub

Public Sub FileSaveAs()
    ' Let the normal Save As happen first (ndOffice intercepts this for versioning)
    On Error Resume Next
    Dialogs(wdDialogFileSaveAs).Show
    Application.OnTime Now + TimeValue("00:00:03"), "MigrateIManageToND"
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

    ' --- Build regexes ---
    Dim imRegex As Object
    Set imRegex = CreateObject("VBScript.RegExp")
    imRegex.Global = True
    imRegex.IgnoreCase = True
    imRegex.Pattern = "\d{5,}(v\d+)?"

    ' ND regex to detect existing ND refs in footer (may have old version)
    Dim footerNdRegex As Object
    Set footerNdRegex = CreateObject("VBScript.RegExp")
    footerNdRegex.Global = True
    footerNdRegex.IgnoreCase = True
    footerNdRegex.Pattern = "[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}\s*(\.\d+|v\.?\d+)?"

    ' --- Scan and replace in footers (all sections) ---
    Dim sec As Section
    For Each sec In ActiveDocument.Sections
        CheckAndUpdateFooter sec.Footers(wdHeaderFooterPrimary), docId, imRegex, footerNdRegex

        If sec.PageSetup.DifferentFirstPageHeaderFooter Then
            CheckAndUpdateFooter sec.Footers(wdHeaderFooterFirstPage), docId, imRegex, footerNdRegex
        End If
    Next sec

    Exit Sub

ErrorHandler:
    ' Silently fail - don't interrupt the user opening their document
End Sub

Private Sub CheckAndUpdateFooter(ftr As HeaderFooter, docId As String, imRegex As Object, footerNdRegex As Object)
    Dim para As Paragraph
    For Each para In ftr.Range.Paragraphs
        Dim paraText As String
        paraText = para.Range.Text

        ' Check for iManage ref first
        If imRegex.Test(paraText) Then
            ReplaceMigratePara para, docId
            Exit Sub
        End If

        ' Check for existing ND ref that doesn't match current version
        If footerNdRegex.Test(paraText) Then
            Dim footerMatches As Object
            Set footerMatches = footerNdRegex.Execute(paraText)
            ' Only replace if the ref is different (e.g. old version)
            If footerMatches(0).Value <> docId Then
                ReplaceMigratePara para, docId
            End If
            Exit Sub
        End If
    Next para
End Sub

Private Sub ReplaceMigratePara(para As Paragraph, docId As String)
    Dim rng As Range
    Set rng = para.Range
    ' Preserve existing alignment
    Dim existingAlign As Long
    existingAlign = rng.ParagraphFormat.Alignment
    rng.MoveEnd wdCharacter, -1
    rng.Text = docId
    rng.Font.Size = 8
    rng.Font.Color = RGB(128, 128, 128)
    rng.ParagraphFormat.Alignment = existingAlign
End Sub
