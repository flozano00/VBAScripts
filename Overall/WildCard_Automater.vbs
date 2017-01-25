Option Explicit


Public Sub Automatic_Filler()





If ActiveCell.Value = "" Then
    ActiveCell.Value = "*"
    
  
End If





End Sub

Public Sub Automtic_StarLoop()

Do
    Automatic_Filler
    ActiveCell.Offset(1, 0).Select
    
Loop Until ActiveCell.Value <> ""

End Sub

