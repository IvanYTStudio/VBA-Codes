'Prank #1
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim i As Integer
Dim j As Integer

i = Int(((Target.Row() - 1) * -1) + (10 - ((Target.Row() - 1) * -1) + 1) * Rnd())
j = Int(((Target.Column() - 1) * -1) + (10 - ((Target.Column() - 1) * -1) + 1) * Rnd())

Application.EnableEvents = False
Target.Offset(i, j).Select
Application.EnableEvents = True

End Sub