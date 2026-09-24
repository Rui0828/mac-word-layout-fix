Attribute VB_Name = "LayoutFix"
' mac-word-layout-fix
' Word for Mac reads unitless HTML table heights (<td height="36">) as points,
' while Word for Windows reads them as 96-dpi pixels (36px = 27pt).
' Rows end up 4/3 taller on Mac. This rescales them after an HTML file opens.

Option Explicit

Private Const PX_TO_PT As Single = 0.75

Private appEvents As LayoutFixEvents

' Runs when Word loads this template from the Startup folder.
Public Sub AutoExec()
    Set appEvents = New LayoutFixEvents
    Set appEvents.App = Word.Application
End Sub

' Manual entry point: Tools > Macro > Macros > FixActiveDocument.
Public Sub FixActiveDocument()
    If Documents.Count = 0 Then Exit Sub
    FixDocument ActiveDocument, True
End Sub

Public Sub FixDocument(ByVal doc As Document, Optional ByVal force As Boolean = False)
    If Not force And Not IsHtmlDocument(doc) Then Exit Sub

    Dim wasSaved As Boolean
    wasSaved = doc.Saved

    Dim t As Table
    For Each t In doc.Tables
        FixTable t
    Next t

    ' Only the on-screen layout changed; don't prompt to save on close.
    doc.Saved = wasSaved
End Sub

Private Function IsHtmlDocument(ByVal doc As Document) As Boolean
    Select Case doc.SaveFormat
        Case wdFormatHTML, wdFormatFilteredHTML, wdFormatWebArchive
            IsHtmlDocument = True
    End Select
End Function

Private Sub FixTable(ByVal t As Table)
    Dim c As Cell
    Dim nested As Table
    Dim doneRows As String
    Dim key As String

    On Error Resume Next
    For Each c In t.Range.Cells
        ' Range.Cells also returns cells of nested tables; skip those here.
        If c.NestingLevel = t.NestingLevel Then
            key = "|" & c.RowIndex & "|"
            If InStr(doneRows, key) = 0 Then
                doneRows = doneRows & key
                If c.HeightRule <> wdRowHeightAuto And c.Height > 0 Then
                    c.Height = c.Height * PX_TO_PT
                End If
            End If
            For Each nested In c.Tables
                FixTable nested
            Next nested
        End If
    Next c
    On Error GoTo 0
End Sub
