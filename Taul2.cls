VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Taul2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Dim calShape As Shape

    

    ' 1. Safely locate our summary calendar shape

    On Error Resume Next

    Set calShape = Me.Shapes("SummaryCalendar")

    On Error GoTo 0

    

    If calShape Is Nothing Then Exit Sub

    

    ' 2. Trigger ONLY when selecting cells E1 (Start) or G1 (End)

    If Not Intersect(Target, Me.Range("E1,G1")) Is Nothing Then

        If Target.Cells.count = 1 Then

            ' Drop down perfectly aligned below the selected cell

            calShape.Left = Target.Left

            calShape.Top = Target.Top + Target.Height + 2

            calShape.Visible = True

            Exit Sub

        End If

    End If

    

    ' Hide calendar if clicking anywhere else

    calShape.Visible = False

End Sub
