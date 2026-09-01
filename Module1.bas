Attribute VB_Name = "Module1"
Option Explicit

Const SUM_SHEET   As String = "Summary by Property"
Const FOLDER_NAME As String = "Saapuneet"

Sub ImportAlarmsFromOutlook()
    Dim olApp        As Object
    Dim olNS         As Object
    Dim olFolder     As Object
    Dim olStore      As Object
    Dim targetStore  As Object
    Dim olItem       As Object
    Dim wsLog        As Worksheet
    Dim wsSum        As Worksheet
    Dim lastRow      As Long
    Dim subj         As String
    Dim bodyText     As String
    Dim cleanBody    As String
    Dim parts()      As String
    Dim prop         As String
    Dim alarmName    As String
    Dim alarmCode    As String
    Dim dt           As String
    Dim tm           As String
    Dim recv         As String
    Dim entryID      As String
    Dim added        As Long
    Dim i            As Long
    Dim alreadyExists As Boolean
    Dim fld          As Object
    Dim processedCount As Long
    Dim isMailCleared As Boolean
    
    Dim filterStartDate As Date, filterEndDate As Date
    Dim hasStartFilter As Boolean, hasEndFilter As Boolean
    Dim mailDateOnly   As Date
    
    Dim regExDate    As Object: Set regExDate = CreateObject("VBScript.RegExp")
    Dim regExTime    As Object: Set regExTime = CreateObject("VBScript.RegExp")
    Dim match        As Object
    
    regExDate.Pattern = "\b\d{4}\.\d{2}\.\d{2}\b"
    regExTime.Pattern = "\b\d{2}[:\.]\d{2}[:\.]\d{2}\b"

    Set wsLog = ActiveSheet
    
    On Error Resume Next
    Set wsSum = ThisWorkbook.Sheets(SUM_SHEET)
    On Error GoTo 0
    
    If wsSum Is Nothing Then
        MsgBox "Critical Error: Could not find summary sheet named '" & SUM_SHEET & "'.", vbCritical, "Sheet Missing"
        Exit Sub
    End If

    ' --- READ USER FILTER BOUNDARIES ---
    hasStartFilter = False: hasEndFilter = False
    If Trim(CStr(wsLog.Range("J5").Value)) <> "" Then
        filterStartDate = ConvertDotDate(wsLog.Range("J5").Value)
        If filterStartDate > 0 Then hasStartFilter = True
    End If
    If Trim(CStr(wsLog.Range("L5").Value)) <> "" Then
        filterEndDate = ConvertDotDate(wsLog.Range("L5").Value)
        If filterEndDate > 0 Then hasEndFilter = True
    End If

    ' --- RESET ACTIVE WORK AREA (COLUMNS A:H ONLY) ---
    If wsLog.AutoFilterMode Then wsLog.AutoFilterMode = False
    lastRow = wsLog.Cells(wsLog.Rows.count, 1).End(xlUp).Row
    
    If lastRow >= 2 Then
        With wsLog.Range("A2:H" & lastRow)
            .ClearContents
            .Interior.ColorIndex = xlNone
            .Borders.LineStyle = xlNone
        End With
    End If

    ' --- RE-WRITE HEADERS ---
    wsLog.Cells(1, 1).Value = "Property"
    wsLog.Cells(1, 2).Value = "Alarm Name"
    wsLog.Cells(1, 3).Value = "Alarm Code"
    wsLog.Cells(1, 4).Value = "Date"
    wsLog.Cells(1, 5).Value = "Time"
    wsLog.Cells(1, 6).Value = "Subject"
    wsLog.Cells(1, 7).Value = "Entry ID"
    wsLog.Cells(1, 8).Value = "Received Date"

    ' --- INITIALIZE UI COUNTER ---
    added = 0
    processedCount = 0
    wsLog.Range("I1").Value = "Connecting..."
    wsLog.Range("I1").Font.Bold = True

    Set olApp = CreateObject("Outlook.Application")
    Set olNS = olApp.GetNamespace("MAPI")
    olNS.Logon "", "", False, False

    For Each olStore In olNS.Stores
        If LCase(olStore.DisplayName) Like "*pohjoinen*" Then
            Set targetStore = olStore: Exit For
        End If
    Next olStore
    If targetStore Is Nothing Then Set targetStore = olNS.DefaultStore

    ' --- FOLDER MAPPING ---
    For Each fld In targetStore.GetRootFolder.Folders
        If LCase(fld.Name) Like "*pohjoinenhalytys*" Then Set olFolder = fld: Exit For
    Next fld
    
    If olFolder Is Nothing Then
        For Each fld In targetStore.GetRootFolder.Folders
            If fld.Name = FOLDER_NAME Then Set olFolder = fld: Exit For
        Next fld
    End If
    
    If Not olFolder Is Nothing And olFolder.Name = FOLDER_NAME Then
        Dim subFld As Object
        For Each subFld In olFolder.Folders
            If LCase(subFld.Name) Like "*pohjoinenhalytys*" Then Set olFolder = subFld: Exit For
        Next subFld
    End If

    If olFolder Is Nothing Then MsgBox "Could not find alarm folder.", vbExclamation: Exit Sub

    lastRow = wsLog.Cells(wsLog.Rows.count, 1).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    ' --- MAIN PROCESSING LOOP ---
    For Each olItem In olFolder.Items
        If olItem.Class = 43 Then
            processedCount = processedCount + 1
            If processedCount Mod 25 = 0 Then
                wsLog.Range("I1").Value = "Scanning item #" & processedCount & "..."
                DoEvents
            End If
            
            mailDateOnly = Int(olItem.ReceivedTime)
            If hasStartFilter And mailDateOnly < filterStartDate Then GoTo NextItem
            If hasEndFilter And mailDateOnly > filterEndDate Then GoTo NextItem
            
            subj = Trim(olItem.Subject)
            entryID = olItem.entryID
            
            If InStr(LCase(subj), "uusi jäsenyys") > 0 Or InStr(LCase(subj), "tilille") > 0 Then GoTo NextItem
            
            bodyText = Replace(Replace(olItem.Body, vbCrLf, " "), vbLf, " ")
            bodyText = Trim(bodyText)
            If bodyText = "" Or Left(LCase(bodyText), 4) = "http" Then GoTo NextItem

            ' Dupe check
            alreadyExists = False
            For i = 2 To lastRow
                If wsLog.Cells(i, 7).Value = entryID Then alreadyExists = True: Exit For
            Next i
            If alreadyExists Then GoTo NextItem

            ' --- DYNAMIC TIMESTAMP TARGETING ---
            isMailCleared = (InStr(LCase(subj), "poistunut") > 0) Or (InStr(LCase(wsLog.Cells(i, 3).Value), "poistunut") > 0)
            
            dt = "": tm = ""
            If isMailCleared Then
                dt = Format(olItem.ReceivedTime, "yyyy.mm.dd")
                tm = Format(olItem.ReceivedTime, "HH.mm.ss")
            Else
                If regExDate.Test(bodyText) Then dt = regExDate.Execute(bodyText)(0).Value
                If regExTime.Test(bodyText) Then tm = Replace(regExTime.Execute(bodyText)(0).Value, ":", ".")
                If dt = "" Then dt = Format(olItem.ReceivedTime, "yyyy.mm.dd")
                If tm = "" Then tm = Format(olItem.ReceivedTime, "HH.mm.ss")
            End If
            
            cleanBody = Replace(Replace(Replace(Replace(bodyText, " : ", "||"), " :", "||"), ": ", "||"), ":", "||")
            parts = Split(cleanBody, "||")
            
            prop = "": alarmName = "": alarmCode = ""
            If UBound(parts) >= 3 Then
                prop = Trim(parts(1)): alarmName = Trim(parts(2))
                Dim tailWords() As String: tailWords = Split(Trim(parts(3)), " ")
                Dim baseCode As String: baseCode = ""
                For i = 0 To UBound(tailWords)
                    If tailWords(i) <> "" Then baseCode = baseCode & IIf(baseCode = "", "", " ") & tailWords(i)
                Next i
                alarmCode = Trim(parts(0)) & IIf(baseCode <> "", " - " & baseCode, "")
            Else
                If UBound(parts) >= 1 Then prop = Trim(parts(1)): alarmCode = Trim(parts(0)) Else prop = subj: alarmCode = "Hälytys"
                alarmName = "System Alert Notification"
            End If

            lastRow = lastRow + 1
            wsLog.Cells(lastRow, 1).Value = prop: wsLog.Cells(lastRow, 2).Value = alarmName
            wsLog.Cells(lastRow, 3).Value = alarmCode: wsLog.Cells(lastRow, 4).Value = CStr(dt)
            wsLog.Cells(lastRow, 5).Value = "'" & CStr(tm): wsLog.Cells(lastRow, 6).Value = subj
            wsLog.Cells(lastRow, 7).Value = entryID: wsLog.Cells(lastRow, 8).Value = Format(olItem.ReceivedTime, "dd/mm/yyyy")
            
            added = added + 1
        End If
NextItem:
    Next olItem

    wsLog.Range("I1").Value = "Done (" & processedCount & " items checked)"
    RefreshSummary
    MsgBox added & " new items added successfully.", vbInformation, "Sync Complete"
    
    ThisWorkbook.Sheets("Alarms Log").AutoFilterMode = False
    ThisWorkbook.Sheets("Alarms Log").Range("A1:H1").AutoFilter
    
End Sub

Sub RefreshSummary()
    Dim wsLog     As Worksheet
    Dim wsSum     As Worksheet
    Dim lastRow   As Long
    Dim i         As Long
    Dim prop      As String
    Dim alarmName As String
    Dim subj      As String
    Dim dtStr     As String
    Dim isCleared As Boolean
    Dim monthKey  As String
    
    Dim monthVariant As Variant
    Dim sortedMonths() As String
    Dim monthIdx  As Long
    
    Dim filterStartDate As Date, filterEndDate As Date
    Dim hasStartFilter  As Boolean, hasEndFilter  As Boolean
    Dim logRowDate As Date
    Dim sParts()  As String
    
    Dim tsArr()   As String
    Dim idx       As Long
    Dim x         As Long
    Dim y         As Long
    Dim tempStr   As String
    Dim dtX       As Date
    Dim dtY       As Date
    Dim count     As Long
    
    Set wsLog = ThisWorkbook.Sheets("Alarms Log")
    On Error Resume Next
    Set wsSum = ThisWorkbook.Sheets(SUM_SHEET)
    On Error GoTo 0
    
    If wsSum Is Nothing Then Exit Sub

    With wsSum
        .Range("B1:C1").ClearContents
        .Range("B1:C1").Interior.ColorIndex = xlNone
        .Range("B1:C1").Borders.LineStyle = xlNone
        
        .Range("D1").Value = "Start Date:"
        .Range("D1").Font.Bold = True
        .Range("D1").HorizontalAlignment = xlRight
        
        With .Range("E1")
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(255, 255, 153)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .NumberFormat = "m/d/yyyy"
        End With
        
        .Range("F1").Value = "End Date:"
        .Range("F1").Font.Bold = True
        .Range("F1").HorizontalAlignment = xlRight
        
        With .Range("G1")
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(255, 255, 153)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .NumberFormat = "m/d/yyyy"
        End With
    End With

    hasStartFilter = False
    hasEndFilter = False
    
    If Trim(CStr(wsSum.Range("E1").Value)) <> "" Then
        filterStartDate = ConvertDotDate(wsSum.Range("E1").Value)
        If filterStartDate > 0 Then hasStartFilter = True
    ElseIf Trim(CStr(wsLog.Range("J5").Value)) <> "" Then
        filterStartDate = ConvertDotDate(wsLog.Range("J5").Value)
        If filterStartDate > 0 Then hasStartFilter = True
    End If
    
    If Trim(CStr(wsSum.Range("G1").Value)) <> "" Then
        filterEndDate = ConvertDotDate(wsSum.Range("G1").Value)
        If filterEndDate > 0 Then hasEndFilter = True
    ElseIf Trim(CStr(wsLog.Range("L5").Value)) <> "" Then
        filterEndDate = ConvertDotDate(wsLog.Range("L5").Value)
        If filterEndDate > 0 Then hasEndFilter = True
    End If

    On Error Resume Next
    wsSum.AutoFilterMode = False
    wsSum.Cells.ClearOutline
    
    With wsSum.Rows("2:" & wsSum.Rows.count)
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Font.Bold = False
        .Font.Color = vbBlack
        .NumberFormat = "General"
    End With
    On Error GoTo 0

    lastRow = wsLog.Cells(wsLog.Rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim dictCounts   As Object: Set dictCounts = CreateObject("Scripting.Dictionary")
    Dim dictMonths   As Object: Set dictMonths = CreateObject("Scripting.Dictionary")
    Dim dictTS       As Object: Set dictTS = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        dtStr = wsLog.Cells(i, 4).Value
        If Left(dtStr, 1) = "'" Then dtStr = Mid(dtStr, 2)
        
        If dtStr <> "" Then
            logRowDate = ConvertDotDate(dtStr)
            If logRowDate > 0 Then
                If hasStartFilter Then
                    If logRowDate < filterStartDate Then GoTo NextLogRow
                End If
                If hasEndFilter Then
                    If logRowDate > filterEndDate Then GoTo NextLogRow
                End If
            End If
        End If

        prop = Trim(wsLog.Cells(i, 1).Value)
        alarmName = Trim(wsLog.Cells(i, 2).Value)
        subj = wsLog.Cells(i, 6).Value
        
        If prop = "" Then GoTo NextLogRow

        Dim subCode As String: subCode = ""
        Dim commaPos As Long: commaPos = InStr(prop, ",")
        If commaPos > 0 Then
            subCode = Trim(Mid(prop, commaPos + 1))
            prop = Trim(Left(prop, commaPos - 1))
        End If
        
        If subCode <> "" Then
            alarmName = subCode & " - " & alarmName
        End If

        isCleared = (InStr(LCase(subj), "poistunut") > 0) Or (InStr(LCase(wsLog.Cells(i, 3).Value), "poistunut") > 0)

        monthKey = "Unknown Month"
        If logRowDate > 0 Then
            monthKey = Format(DateSerial(Year(logRowDate), Month(logRowDate), 1), "mmm-yy")
        End If

        If monthKey <> "Unknown Month" Then dictMonths(monthKey) = True

        Dim compoundKey As String
        compoundKey = prop & "||" & alarmName & "||" & monthKey

        Dim tsKey As String
        tsKey = prop & "||" & alarmName
        
        Dim fullTS As String
        fullTS = wsLog.Cells(i, 4).Value & " " & wsLog.Cells(i, 5).Value
        
        Dim stateLabel As String
        stateLabel = IIf(isCleared, "Cleared", "Active")
        
        Dim tsData As String
        tsData = fullTS & "||" & monthKey & "||" & stateLabel

        dictCounts(compoundKey) = dictCounts(compoundKey) + 1
        If Not dictTS.Exists(tsKey) Then Set dictTS(tsKey) = New Collection
        dictTS(tsKey).Add tsData
        
NextLogRow:
    Next i

    If dictMonths.count > 0 Then
        ReDim sortedMonths(0 To dictMonths.count - 1)
        monthIdx = 0
        For Each monthVariant In dictMonths.Keys
            sortedMonths(monthIdx) = monthVariant
            monthIdx = monthIdx + 1
        Next monthVariant
        
        Dim m1 As Long, m2 As Long, tempMonth As String
        For m1 = 0 To UBound(sortedMonths) - 1
            For m2 = m1 + 1 To UBound(sortedMonths)
                If CDate("1 " & sortedMonths(m1)) > CDate("1 " & sortedMonths(m2)) Then
                    tempMonth = sortedMonths(m1)
                    sortedMonths(m1) = sortedMonths(m2)
                    sortedMonths(m2) = tempMonth
                End If
            Next m2
        Next m1
    Else
        ReDim sortedMonths(0 To 0)
        sortedMonths(0) = "No Data"
    End If

    Dim HDR_FILL_A  As Long: HDR_FILL_A = RGB(31, 56, 100)
    Dim PARENT_BG   As Long: PARENT_BG = RGB(242, 242, 242)
    
    Dim r As Long: r = 3
    Dim monthCount As Long
    monthCount = IIf(dictMonths.count > 0, UBound(sortedMonths) + 1, 0)

    wsSum.Outline.SummaryRow = xlSummaryAbove

    wsSum.Cells(r, 1).Value = "PROPERTY ALARM TIMELINE LOG (Active & Cleared)"
    wsSum.Cells(r, 1).Font.Bold = True
    wsSum.Cells(r, 1).Font.Size = 13
    wsSum.Cells(r, 1).Font.Color = RGB(255, 255, 255)
    wsSum.Range(wsSum.Cells(r, 1), wsSum.Cells(r, 2 + monthCount)).Interior.Color = HDR_FILL_A
    r = r + 1

    wsSum.Cells(r, 1).Value = "Property Structure / Alarm Metrics"
    Dim mi As Long
    For mi = 0 To monthCount - 1
        wsSum.Cells(r, 2 + mi).Value = sortedMonths(mi)
    Next mi
    wsSum.Cells(r, 2 + monthCount).Value = "Total"

    Dim c As Long
    For c = 1 To 2 + monthCount
        With wsSum.Cells(r, c)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = HDR_FILL_A
            .HorizontalAlignment = xlCenter
        End With
    Next c
    r = r + 1

    Dim dictPropGroup As Object: Set dictPropGroup = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In dictCounts.Keys
        sParts = Split(k, "||")
        If Not dictPropGroup.Exists(sParts(0)) Then
            Set dictPropGroup(sParts(0)) = CreateObject("Scripting.Dictionary")
        End If
        dictPropGroup(sParts(0))(sParts(1)) = True
    Next k

    Dim propKey As Variant, alarmKey As Variant
    Dim propRow As Long, alarmRow As Long
    Dim startTSLong As Long, endTSLong As Long
    Dim startPropChildren As Long, endPropChildren As Long
    Dim alarmRowsColl As Collection
    Dim colLetter As String
    
    For Each propKey In dictPropGroup.Keys
        propRow = r
        wsSum.Cells(propRow, 1).Value = propKey & " Total"
        wsSum.Cells(propRow, 1).Font.Bold = True
        wsSum.Range(wsSum.Cells(propRow, 1), wsSum.Cells(propRow, 2 + monthCount)).Interior.Color = PARENT_BG
        r = r + 1
        
        Set alarmRowsColl = New Collection
        startPropChildren = r
        
        For Each alarmKey In dictPropGroup(propKey).Keys
            alarmRow = r
            alarmRowsColl.Add alarmRow
            
            wsSum.Cells(alarmRow, 1).Value = "  " & alarmKey
            wsSum.Cells(alarmRow, 1).Font.Bold = False
            wsSum.Cells(alarmRow, 1).Font.Color = RGB(40, 40, 40)
            r = r + 1
            
            tsKey = propKey & "||" & alarmKey
            startTSLong = r
            
            If dictTS.Exists(tsKey) Then
                count = dictTS(tsKey).count
                If count > 0 Then
                    ReDim tsArr(1 To count)
                    For idx = 1 To count
                        tsArr(idx) = dictTS(tsKey)(idx)
                    Next idx
                    
                    ' Sort descending (Newest to Oldest)
                    For x = 1 To count - 1
                        For y = x + 1 To count
                            dtX = ParseFullTS(Split(tsArr(x), "||")(0))
                            dtY = ParseFullTS(Split(tsArr(y), "||")(0))
                            If dtX < dtY Then
                                tempStr = tsArr(x)
                                tsArr(x) = tsArr(y)
                                tsArr(y) = tempStr
                            End If
                        Next y
                    Next x
                    
                    For idx = 1 To count
                        Dim tsParts() As String
                        tsParts = Split(tsArr(idx), "||")
                        
                        Dim textTimelineRow As String
                        If tsParts(2) = "Active" Then
                            textTimelineRow = "    Active:   " & tsParts(0)
                        Else
                            textTimelineRow = "    Cleared: " & tsParts(0)
                        End If
                        
                        wsSum.Cells(r, 1).Value = textTimelineRow
                        wsSum.Cells(r, 1).Font.Bold = False
                        wsSum.Cells(r, 1).Font.Color = vbBlack
                        
                        For mi = 0 To monthCount - 1
                            If sortedMonths(mi) = tsParts(1) Then
                                wsSum.Cells(r, 2 + mi).Value = 1
                            Else
                                wsSum.Cells(r, 2 + mi).Value = 0
                            End If
                            wsSum.Cells(r, 2 + mi).Font.Bold = False
                            wsSum.Cells(r, 2 + mi).NumberFormat = "#,##0"
                        Next mi
                        wsSum.Cells(r, 2 + monthCount).Value = 1
                        wsSum.Cells(r, 2 + monthCount).Font.Bold = False
                        wsSum.Cells(r, 2 + monthCount).NumberFormat = "#,##0"
                        r = r + 1
                    Next idx
                End If
            End If
            endTSLong = r - 1
            
            For mi = 0 To monthCount
                colLetter = Split(wsSum.Cells(1, 2 + mi).Address, "$")(1)
                If endTSLong >= startTSLong Then
                    wsSum.Cells(alarmRow, 2 + mi).Formula = "=SUM(" & colLetter & startTSLong & ":" & colLetter & endTSLong & ")"
                Else
                    wsSum.Cells(alarmRow, 2 + mi).Value = 0
                End If
                wsSum.Cells(alarmRow, 2 + mi).Font.Bold = False
                wsSum.Cells(alarmRow, 2 + mi).NumberFormat = "#,##0"
            Next mi
            
            If endTSLong >= startTSLong Then
                wsSum.Rows(startTSLong & ":" & endTSLong).Group
            End If
        Next alarmKey
        endPropChildren = r - 1
        
        For mi = 0 To monthCount
            colLetter = Split(wsSum.Cells(1, 2 + mi).Address, "$")(1)
            If alarmRowsColl.count > 0 Then
                Dim formulaStr As String: formulaStr = "="
                Dim alarmRowIdx As Variant
                For Each alarmRowIdx In alarmRowsColl
                    formulaStr = formulaStr & colLetter & alarmRowIdx & "+"
                Next alarmRowIdx
                formulaStr = Left(formulaStr, Len(formulaStr) - 1)
                wsSum.Cells(propRow, 2 + mi).Formula = formulaStr
            Else
                wsSum.Cells(propRow, 2 + mi).Value = 0
            End If
            wsSum.Cells(propRow, 2 + mi).Font.Bold = True
            wsSum.Cells(propRow, 2 + mi).NumberFormat = "#,##0"
        Next mi
        
        If endPropChildren >= startPropChildren Then
            wsSum.Rows(startPropChildren & ":" & endPropChildren).Group
        End If
    Next propKey

    Dim finalRow As Long: finalRow = r - 1
    Dim finalCol As Long: finalCol = 2 + monthCount
    
    With wsSum.Range("A3:A" & finalRow)
        .Font.Name = "Calibri"
        .Font.Size = 13
        .Font.Italic = False
        .HorizontalAlignment = xlLeft
    End With
    
    If finalCol >= 2 Then
        With wsSum.Range(wsSum.Cells(5, 2), wsSum.Cells(finalRow, finalCol))
            .Font.Name = "Calibri"
            .Font.Size = 11
            .HorizontalAlignment = xlRight
        End With
    End If

    wsSum.Columns(1).ColumnWidth = 55
    
    wsSum.Range(wsSum.Cells(4, 1), wsSum.Cells(4, 2 + monthCount)).AutoFilter
    wsSum.Range("A4").AutoFilter Field:=1, VisibleDropdown:=False
    
    For mi = 0 To monthCount
        wsSum.Columns(2 + mi).ColumnWidth = 16
    Next mi
    
    ' --- RUN FINNISH WORK LEDGER AUTOMATION ---
    UpdateDoneWorkSheet dictPropGroup
    
    On Error Resume Next
    wsSum.Activate
    
    ActiveWindow.FreezePanes = False
    wsSum.Cells(5, 1).Select
    ActiveWindow.FreezePanes = True
    
    wsSum.Outline.ShowLevels RowLevels:=3
    wsSum.Outline.ShowLevels RowLevels:=2
    wsSum.Outline.ShowLevels RowLevels:=1
    On Error GoTo 0
End Sub
Sub UpdateDoneWorkSheet(ByRef dictProps As Object)
    Dim wsDone As Worksheet
    On Error Resume Next
    Set wsDone = ThisWorkbook.Sheets("Done work")
    On Error GoTo 0
    If wsDone Is Nothing Then Exit Sub
    
    ' 1. Clear existing Filters/Outline
    On Error Resume Next
    wsDone.Cells.ClearOutline
    wsDone.AutoFilterMode = False
    On Error GoTo 0
    
    ' 2. Initialize Headers
    If Trim(CStr(wsDone.Cells(1, 1).Value)) <> "HUOLTOTYÖMAAT" Then
        wsDone.Cells.Clear
        wsDone.Cells(1, 1).Value = "HUOLTOTYÖMAAT"
        wsDone.Cells(1, 1).Font.Bold = True: wsDone.Cells(1, 1).Font.Size = 14: wsDone.Cells(1, 1).Font.Color = RGB(255, 255, 255)
        wsDone.Range("A1:B1").Interior.Color = RGB(31, 56, 100)
        wsDone.Cells(2, 1).Value = "Kohteet": wsDone.Cells(2, 2).Value = "Tiedot"
        wsDone.Range("A2:B2").Font.Bold = True: wsDone.Range("A2:B2").Interior.Color = RGB(242, 242, 242)
        wsDone.Columns(1).ColumnWidth = 45: wsDone.Columns(2).ColumnWidth = 75
    End If
    
    ' 3. GRID GUARD: Find the last valid block and wipe out any broken rows below it
    Dim lastHeaderRow As Long: lastHeaderRow = 0
    Dim checkRow As Long
    For checkRow = wsDone.Cells(wsDone.Rows.count, 1).End(xlUp).Row To 3 Step -1
        If wsDone.Cells(checkRow, 1).Font.Bold = True And Left(wsDone.Cells(checkRow, 1).Value, 4) <> "    " And Trim(wsDone.Cells(checkRow, 1).Value) <> "" Then
            If (checkRow - 3) Mod 6 = 0 Then
                lastHeaderRow = checkRow
                Exit For
            End If
        End If
    Next checkRow
    
    Dim lastRow As Long
    If lastHeaderRow = 0 Then
        lastRow = 2
    Else
        lastRow = lastHeaderRow + 5
    End If
    
    ' Automatically slice away any corrupted rows at the bottom (like the ones in image_5a9642.png)
    Dim absoluteLastRow As Long: absoluteLastRow = wsDone.Cells(wsDone.Rows.count, 1).End(xlUp).Row
    If absoluteLastRow > lastRow Then
        wsDone.Rows((lastRow + 1) & ":" & absoluteLastRow).Clear
    End If
    
    ' 4. Map Valid Existing Properties
    Dim existingProps As Object: Set existingProps = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 3 To lastRow Step 6
        If wsDone.Cells(i, 1).Value <> "" Then existingProps(Trim(wsDone.Cells(i, 1).Value)) = True
    Next i
    
    ' 5. Add New Properties strictly aligned to the grid
    Dim propKey As Variant, r As Long
    For Each propKey In dictProps.Keys
        Dim cleanProp As String: cleanProp = Trim(Split(CStr(propKey), ",")(0))
        If Not existingProps.Exists(cleanProp) Then
            r = wsDone.Cells(wsDone.Rows.count, 1).End(xlUp).Row + 1
            
            ' Absolute Guard: Enforce strict step-6 structural tracking
            If (r - 3) Mod 6 <> 0 Then
                r = 3 + 6 * Application.WorksheetFunction.Ceiling((r - 2) / 6, 1)
            End If
            
            ' Header
            wsDone.Cells(r, 1).Value = cleanProp
            wsDone.Cells(r, 1).Font.Name = "Calibri": wsDone.Cells(r, 1).Font.Size = 13: wsDone.Cells(r, 1).Font.Bold = True
            wsDone.Range(wsDone.Cells(r, 1), wsDone.Cells(r, 2)).Interior.Color = RGB(235, 241, 251)
            wsDone.Rows(r).RowHeight = 26
            
            ' Labels
            Dim labels As Variant: labels = Array("    Mitä myyty:", "    Tehdyt Työt:", "    Myyjä:", "    Hinta €:", "    Lisätietoja:")
            Dim L As Long
            For L = 0 To 4
                wsDone.Cells(r + L + 1, 1).Value = labels(L)
                wsDone.Cells(r + L + 1, 1).Font.Name = "Calibri": wsDone.Cells(r + L + 1, 1).Font.Size = 11
                
                With wsDone.Cells(r + L + 1, 2)
                    .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(225, 225, 225)
                    .Font.Name = "Calibri": .VerticalAlignment = xlTop: .WrapText = True
                End With
                
                wsDone.Rows(r + L + 1).RowHeight = IIf(L = 1 Or L = 4, 60, 21)
            Next L
            existingProps(cleanProp) = True
        End If
    Next propKey
    
    ' 6. Clean Batch Grouping Pass
    lastRow = wsDone.Cells(wsDone.Rows.count, 1).End(xlUp).Row
    For i = 3 To lastRow Step 6
        If wsDone.Cells(i, 1).Font.Bold = True And Left(wsDone.Cells(i, 1).Value, 4) <> "    " Then
            On Error Resume Next
            wsDone.Rows((i + 1) & ":" & (i + 5)).Group
            On Error GoTo 0
        End If
    Next i
    
    wsDone.Range("A2:B2").AutoFilter
    wsDone.Outline.SummaryRow = xlSummaryAbove
    wsDone.Outline.ShowLevels RowLevels:=1
End Sub

Private Function ParseFullTS(ByVal tsStr As String) As Date
    On Error Resume Next
    Dim spacePos As Long: spacePos = InStr(tsStr, " ")
    Dim dStr As String, tStr As String
    
    If spacePos > 0 Then
        dStr = Trim(Left(tsStr, spacePos - 1))
        tStr = Trim(Mid(tsStr, spacePos + 1))
    Else
        dStr = Trim(tsStr)
        tStr = "00:00:00"
    End If
    
    tStr = Replace(tStr, ".", ":")
    
    Dim parts() As String
    parts = Split(dStr, ".")
    If UBound(parts) = 2 Then
        ParseFullTS = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2))) + CDate(tStr)
    Else
        ParseFullTS = CDate(tsStr)
    End If
    On Error GoTo 0
End Function

Private Function ConvertDotDate(ByVal rawVal As Variant) As Date
    Dim s As String
    s = Trim(CStr(rawVal))
    If s = "" Then
        ConvertDotDate = 0
        Exit Function
    End If
    
    If IsDate(s) Then
        ConvertDotDate = DateValue(s)
        Exit Function
    End If
    
    Dim parts() As String
    
    If InStr(s, "-") > 0 Then
        parts = Split(s, "-")
        If UBound(parts) = 2 Then
            On Error Resume Next
            If Len(parts(0)) = 4 Then
                ConvertDotDate = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
            End If
            On Error GoTo 0
            Exit Function
        End If
    End If
    
    If InStr(s, ".") > 0 Then
        parts = Split(s, ".")
        If UBound(parts) = 2 Then
            On Error Resume Next
            If Len(parts(0)) = 4 Then
                ConvertDotDate = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
            Else
                ConvertDotDate = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
            End If
            On Error GoTo 0
            Exit Function
        End If
    End If
    
    ConvertDotDate = 0
End Function

