VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Taul1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' --- AUTOMAATTINEN MUOTOILUN KORJAUS ---
    ' Jos soluun J5 tai L5 syˆtet‰‰n kalenterista viivallinen p‰iv‰m‰‰r‰, korjataan se heti
    If Not Intersect(Target, Me.Range("J5,L5")) Is Nothing Then
        If Target.Cells.count = 1 Then
            Dim valAsStr As String
            valAsStr = Trim(CStr(Target.Value))
            
            ' Jos p‰iv‰m‰‰r‰ss‰ on viivoja (esim. 2026-05-11)
            If InStr(valAsStr, "-") > 0 And IsDate(valAsStr) Then
                Application.EnableEvents = False ' Estet‰‰n ikuinen silmukka
                
                ' Muutetaan muoto lennosta pisteelliseksi suomalaiseksi tekstiksi
                Target.NumberFormat = "@"
                Target.Value = Format(CDate(valAsStr), "d.m.yyyy")
                
                Application.EnableEvents = True
            End If
        End If
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim calShape As Shape
    
    On Error Resume Next
    Set calShape = Me.Shapes("Calendar")
    On Error GoTo 0
    
    If calShape Is Nothing Then Exit Sub
    
    ' Kalenterin n‰ytt‰minen rivill‰ 5
    If Not Intersect(Target, Me.Range("J5,L5")) Is Nothing Then
        If Target.Cells.count = 1 Then
' --- ALIGN DIRECTLY BELOW THE CELL ---
            ' Left aligns perfectly with the left edge of the cell
            calShape.Left = Target.Left
            ' Top drops down to the exact bottom edge of the cell + 2 pixels of padding
            calShape.Top = Target.Top + Target.Height + 2
            calShape.Visible = True
            Exit Sub
        End If
    End If
    
    calShape.Visible = False
End Sub
