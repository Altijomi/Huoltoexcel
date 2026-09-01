Attribute VB_Name = "Module2"
' ============================================================
'   ALARM TRACKER — FULL BODY DIAGNOSTIC INTAKE
'   Dumps the ENTIRE unedited email body into column C
' ============================================================
Option Explicit

Const FOLDER_NAME As String = "Saapuneet"
Const DIAG_SHEET  As String = "Raw Emails Test"

Sub ImportRawEmailsForAnalysis()
    Dim olApp        As Object
    Dim olNS         As Object
    Dim olFolder     As Object
    Dim olStore      As Object
    Dim targetStore  As Object
    Dim olItem       As Object
    Dim ws           As Worksheet
    Dim checkWs      As Worksheet
    Dim lastRow      As Long
    Dim added        As Long
    Dim sheetExists  As Boolean
    
    sheetExists = False
    
    ' 1. Locate or create the diagnostic sheet safely
    For Each checkWs In ThisWorkbook.Worksheets
        If checkWs.Name = DIAG_SHEET Then
            Set ws = checkWs
            sheetExists = True
            Exit For
        End If
    Next checkWs
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = DIAG_SHEET
    End If
    
    ' Reset the worksheet view
    ws.Cells.ClearContents
    ws.Cells.Interior.ColorIndex = xlNone
    
    ' Establish clean tracking headers
    ws.Cells(1, 1).Value = "Received Date & Time"
    ws.Cells(1, 2).Value = "RAW Unedited Subject Line"
    ws.Cells(1, 3).Value = "FULL Unedited Email Body"
    
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' 2. Establish connection to Outlook mailbox
    Set olApp = CreateObject("Outlook.Application")
    Set olNS = olApp.GetNamespace("MAPI")
    olNS.Logon "", "", False, False

    For Each olStore In olNS.Stores
        If LCase(olStore.DisplayName) Like "*pohjoinen*" Then
            Set targetStore = olStore
            Exit For
        End If
    Next olStore

    If targetStore Is Nothing Then
        MsgBox "Could not find mailbox 'Pohjoinen Hälytykset'.", vbExclamation, "Mailbox Missing"
        Exit Sub
    End If

    Dim fld As Object
    For Each fld In targetStore.GetRootFolder.Folders
        If fld.Name = FOLDER_NAME Then
            Set olFolder = fld
            Exit For
        End If
    Next fld

    If olFolder Is Nothing Then
        MsgBox "Could not find folder '" & FOLDER_NAME & "'.", vbExclamation, "Folder Missing"
        Exit Sub
    End If

    ' 3. Extract unedited text content structures
    lastRow = 1
    added = 0
    
    On Error GoTo ErrHandler
    For Each olItem In olFolder.Items
        If olItem.Class = 43 Then ' MailItem
            lastRow = lastRow + 1
            
            ws.Cells(lastRow, 1).Value = olItem.ReceivedTime
            ws.Cells(lastRow, 2).Value = olItem.Subject
            
            ' Replace line breaks with a simple space so the whole body fits beautifully on one Excel row
            ws.Cells(lastRow, 3).Value = Trim(Replace(Replace(olItem.Body, vbCrLf, " "), vbLf, " "))
            
            added = added + 1
        End If
    Next olItem

    ' Set clean column sizes for evaluation
    ws.Columns("A:B").AutoFit
    ws.Columns("C").ColumnWidth = 100 ' Keeps column C wide and scrollable
    
    MsgBox added & " raw email entries fully loaded.", vbInformation, "Diagnostic View Ready"
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Diagnostic Error"
End Sub
