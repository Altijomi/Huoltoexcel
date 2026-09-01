Attribute VB_Name = "Module4"
Sub UpdateDatePickerDropdowns()
    Dim wsLog As Worksheet
    Dim dateList As String
    Dim i As Long
    Dim targetDate As Date
    
    ' Dynamically target whatever sheet you are looking at right now
    Set wsLog = ActiveSheet
    
    ' Generate a rolling list of the past 60 days
    For i = 0 To 60
        targetDate = Date - i
        dateList = dateList & Format(targetDate, "yyyy.mm.dd") & ","
    Next i
    
    ' Clean off trailing comma
    If Len(dateList) > 0 Then dateList = Left(dateList, Len(dateList) - 1)
    
    ' Safely force setup of the UI cells if they were cleared
    If wsLog.Range("I1").Value <> "Start Date:" Then
        wsLog.Range("I1").Value = "Start Date:"
        wsLog.Range("I1").Font.Bold = True
    End If
    If wsLog.Range("I2").Value <> "End Date:" Then
        wsLog.Range("I2").Value = "End Date:"
        wsLog.Range("I2").Font.Bold = True
    End If
    
    ' Style input cells J1 and J2 and drop in the drop-down selectors
    With wsLog.Range("J1:J2")
        .Interior.Color = RGB(255, 255, 224) ' Soft light yellow
        .Borders.LineStyle = xlContinuous
    End With
    
    On Error Resume Next
    With wsLog.Range("J1:J2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=dateList
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub

