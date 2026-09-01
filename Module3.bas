Attribute VB_Name = "Module3"
Public SelectedDate As String

Sub LaunchCalendar()
    Dim TempForm As Object
    Dim NewButton As Object
    Dim i As Long, Row As Long, Col As Long
    Dim DayNum As Date, StartDate As Date
    Dim CurrentMonth As Date
    
    CurrentMonth = DateSerial(Year(Date), Month(Date), 1)
    StartDate = CurrentMonth - Weekday(CurrentMonth, vbMonday) + 1
    
    ' 1. Create a temporary UserForm on the fly
    Set TempForm = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    With TempForm
        .Properties("Width") = 220
        .Properties("Height") = 210
        .Properties("Caption") = "Valitse P‰iv‰m‰‰r‰ (" & Format(Date, "mmmm yyyy") & ")"
    End With
    
    ' 2. Programmatically spawn a 7x6 grid of date buttons
    DayNum = StartDate
    For Row = 0 To 5
        For Col = 0 To 6
            Set NewButton = TempForm.Designer.Controls.Add("Forms.CommandButton.1")
            With NewButton
                .Name = "Btn_" & Format(DayNum, "yyyymmdd")
                .Caption = Day(DayNum)
                .Left = 10 + (Col * 28)
                .Top = 10 + (Row * 24)
                .Width = 25
                .Height = 22
                ' Dim old/future month dates slightly
                If Month(DayNum) <> Month(Date) Then .ForeColor = RGB(160, 160, 160)
            End With
            
            ' Add the execution macro code for each button click
            TempForm.CodeModule.InsertLines TempForm.CodeModule.CountOfLines + 1, _
                "Private Sub Btn_" & Format(DayNum, "yyyymmdd") & "_Click()" & vbCrLf & _
                "    SelectedDate = """ & Format(DayNum, "yyyy.mm.dd") & """" & vbCrLf & _
                "    Unload Me" & vbCrLf & _
                "End Sub"
                
            DayNum = DayNum + 1
        Next Col
    Next Row
    
    ' 3. Display the freshly manufactured calendar form
    VBA.UserForms.Add(TempForm.Name).Show
    
    ' 4. Output the choice to the active cell and delete the temporary form form memory
    If SelectedDate <> "" Then ActiveCell.Value = SelectedDate
    ThisWorkbook.VBProject.VBComponents.Remove TempForm
End Sub

