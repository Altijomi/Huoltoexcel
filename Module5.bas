Attribute VB_Name = "Module5"
' ============================================================
'    DIAGNOSTIC MODULE — DATE FILTER ISOLATION TEST SUITE
' ============================================================
Option Explicit

Const FOLDER_NAME As String = "Saapuneet"

Sub CreateTestDropdowns()
    Dim olApp As Object, olNS As Object, olFolder As Object, olStore As Object, targetStore As Object, olItem As Object
    Dim wsActive As Worksheet, fld As Object
    Dim mailDateStr As String
    Dim dictDates As Object: Set dictDates = CreateObject("Scripting.Dictionary")
    Dim dateKey As Variant
    Dim dropdownList As String
    
    On Error Resume Next
    Set wsActive = ThisWorkbook.Sheets("Alarms Log")
    On Error GoTo 0
    
    If wsActive Is Nothing Then Set wsActive = ActiveSheet
    
    wsActive.Activate
    
    wsActive.Range("I1").Value = "START DATE:"
    wsActive.Range("K1").Value = "END DATE:"
    wsActive.Range("I1,K1").Font.Bold = True
    wsActive.Range("I1,K1").Font.Color = RGB(0, 102, 204)
    
    With wsActive.Range("J1,L1")
        .Interior.Color = RGB(255, 255, 153)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = RGB(200, 0, 0)
    End With
    
    Set olApp = CreateObject("Outlook.Application")
    Set olNS = olApp.GetNamespace("MAPI")
    olNS.Logon "", "", False, False

    For Each olStore In olNS.Stores
        If LCase(olStore.DisplayName) Like "*pohjoinen*" Then
            Set targetStore = olStore
            Exit For
        End If
    Next olStore

    If targetStore Is Nothing Then Set targetStore = olNS.DefaultStore

    For Each fld In targetStore.GetRootFolder.Folders
        If fld.Name = FOLDER_NAME Then
            Set olFolder = fld
            Exit For
        End If
    Next fld

    If olFolder Is Nothing Then
        MsgBox "Target folder '" & FOLDER_NAME & "' missing.", vbCritical
        Exit Sub
    End If

    For Each olItem In olFolder.Items
        If olItem.Class = 43 Then
            mailDateStr = Format(olItem.ReceivedTime, "d.m.yyyy")
            dictDates(mailDateStr) = True
        End If
    Next olItem

    If dictDates.count = 0 Then
        MsgBox "No emails found in the folder.", vbExclamation
        Exit Sub
    End If

    dropdownList = ""
    For Each dateKey In dictDates.Keys
        dropdownList = dropdownList & dateKey & ","
    Next dateKey
    
    If Len(dropdownList) > 0 Then dropdownList = Left(dropdownList, Len(dropdownList) - 1)

    On Error Resume Next
    With wsActive.Range("J1,L1").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=dropdownList
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0

    MsgBox "Dropdown menus loaded into J1 and L1!", vbInformation
End Sub

Sub RunIsolatedDateTest()
    Dim olApp As Object, olNS As Object, olFolder As Object, olStore As Object, targetStore As Object, olItem As Object
    Dim wsTest As Worksheet, wsActive As Worksheet, wsCheck As Worksheet
    Dim lastRow As Long, added As Long, fld As Object
    Dim filterStartDate As Date, filterEndDate As Date
    Dim hasStartFilter As Boolean, hasEndFilter As Boolean
    Dim mailDateOnly As Date
    Dim sheetExists As Boolean
    
    ' Explicitly link the inputs to your main dashboard sheet BEFORE creating the test sheet
    On Error Resume Next
    Set wsActive = ThisWorkbook.Sheets("Alarms Log")
    On Error GoTo 0
    If wsActive Is Nothing Then Set wsActive = ActiveSheet
    
    ' Read filter boundaries safely right now while wsActive is guaranteed secure
    hasStartFilter = False
    hasEndFilter = False
    
    If Trim(CStr(wsActive.Range("J1").Value)) <> "" Then
        filterStartDate = LocalConvertDotDate(wsActive.Range("J1").Value)
        If filterStartDate > 0 Then hasStartFilter = True
    End If
    
    If Trim(CStr(wsActive.Range("L1").Value)) <> "" Then
        filterEndDate = LocalConvertDotDate(wsActive.Range("L1").Value)
        If filterEndDate > 0 Then hasEndFilter = True
    End If
    
    ' Locate or build the test sheet tab canvas
    sheetExists = False
    For Each wsCheck In ThisWorkbook.Worksheets
        If wsCheck.Name = "Date Filter Isolation Test" Then
            Set wsTest = wsCheck
            sheetExists = True
            Exit For
        End If
    Next wsCheck
    
    If Not sheetExists Then
        Set wsTest = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsTest.Name = "Date Filter Isolation Test"
    End If
    
    ' FIX: Strictly wipe the test sheet layout only, leaving your inputs safe
    wsTest.Cells.Clear
    
    wsTest.Cells(1, 1).Value = "Outlook Received Time"
    wsTest.Cells(1, 2).Value = "Parsed Date Only"
    wsTest.Cells(1, 3).Value = "Evaluation Status Message"
    wsTest.Cells(1, 4).Value = "Subject Line"
    
    With wsTest.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With

    wsTest.Cells(2, 6).Value = "Active Filter Boundaries:"
    wsTest.Cells(3, 6).Value = "Start Boundary (J1): " & IIf(hasStartFilter, Format(filterStartDate, "yyyy.mm.dd"), "NOT SET (BLANK)")
    wsTest.Cells(4, 6).Value = "End Boundary (L1): " & IIf(hasEndFilter, Format(filterEndDate, "yyyy.mm.dd"), "NOT SET (BLANK)")
    wsTest.Range("F2:F4").Font.Bold = True
    wsTest.Range("F2:F4").Font.Color = RGB(180, 50, 50)

    Set olApp = CreateObject("Outlook.Application")
    Set olNS = olApp.GetNamespace("MAPI")
    olNS.Logon "", "", False, False

    For Each olStore In olNS.Stores
        If LCase(olStore.DisplayName) Like "*pohjoinen*" Then
            Set targetStore = olStore
            Exit For
        End If
    Next olStore

    If targetStore Is Nothing Then Set targetStore = olNS.DefaultStore

    For Each fld In targetStore.GetRootFolder.Folders
        If fld.Name = FOLDER_NAME Then
            Set olFolder = fld
            Exit For
        End If
    Next fld

    If olFolder Is Nothing Then
        MsgBox "Target folder '" & FOLDER_NAME & "' could not be mapped.", vbCritical
        Exit Sub
    End If

    lastRow = 1
    added = 0

    For Each olItem In olFolder.Items
        If olItem.Class = 43 Then
            lastRow = lastRow + 1
            
            mailDateOnly = DateSerial(Year(olItem.ReceivedTime), Month(olItem.ReceivedTime), Day(olItem.ReceivedTime))
            
            wsTest.Cells(lastRow, 1).Value = olItem.ReceivedTime
            wsTest.Cells(lastRow, 2).Value = Format(mailDateOnly, "yyyy.mm.dd")
            wsTest.Cells(lastRow, 4).Value = olItem.Subject
            
            Dim isMatch As Boolean: isMatch = True
            Dim feedbackMsg As String: feedbackMsg = "MATCHED: Within specified dates."
            
            If hasStartFilter Then
                If mailDateOnly < filterStartDate Then
                    isMatch = False
                    feedbackMsg = "SKIPPED: Out of range (Before Start Date: " & Format(filterStartDate, "yyyy.mm.dd") & ")"
                End If
            End If
            
            If hasEndFilter And isMatch Then
                If mailDateOnly > filterEndDate Then
                    isMatch = False
                    feedbackMsg = "SKIPPED: Out of range (After End Date: " & Format(filterEndDate, "yyyy.mm.dd") & ")"
                End If
            End If
            
            wsTest.Cells(lastRow, 3).Value = feedbackMsg
            
            If isMatch Then
                wsTest.Range(wsTest.Cells(lastRow, 1), wsTest.Cells(lastRow, 3)).Interior.Color = RGB(215, 245, 215)
                added = added + 1
            Else
                wsTest.Range(wsTest.Cells(lastRow, 1), wsTest.Cells(lastRow, 3)).Interior.Color = RGB(255, 220, 220)
            End If
        End If
    Next olItem

    wsTest.Columns("A:D").AutoFit
    wsTest.Columns("F").AutoFit
    
    MsgBox "Isolated verification check completed successfully!", vbInformation
End Sub

Private Function LocalConvertDotDate(ByVal rawInput As Variant) As Date
    Dim txt As String: txt = Trim(CStr(rawInput))
    If txt = "" Then
        LocalConvertDotDate = 0
        Exit Function
    End If
    
    If IsDate(txt) Then
        LocalConvertDotDate = DateValue(txt)
        Exit Function
    End If
    
    Dim elements() As String
    elements = Split(txt, ".")
    If UBound(elements) = 2 Then
        On Error Resume Next
        LocalConvertDotDate = DateSerial(CInt(elements(2)), CInt(elements(1)), CInt(elements(0)))
        On Error GoTo 0
    Else
        LocalConvertDotDate = 0
    End If
End Function

