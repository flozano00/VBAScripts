Option Explicit


Public Sub County_Auto()
Dim county As String
Dim abb As String

county = ActiveCell.Text

Select Case county
Case "Morris County":
    abb = "Morris"
Case "Somerset County":
    abb = "Somerset"
Case "Passaic County":
    abb = "Passaic"
Case "Atlantic County":
    abb = "Atlantic"
Case "Bergen County":
    abb = "Bergen"
Case "Broward"
    abb = "Broward"
Case "Essex County"
    abb = "Essex"
Case "Camden County"
    abb = "Camden"
Case "Glouster County"
    abb = "Glouster"
Case "Hudson County"
    abb = "Hudson"
Case "Middlesex County"
    abb = "Middlesex"
Case "Passaic County"
    abb = "Passaic"
Case "Sussex County"
    abb = "Sussex"
Case "Union County"
    abb = "Union"
Case "Warren County"
    abb = "Warren"
Case "Burlington County"
    abb = "Burlington"
Case "Ocean County"
    abb = "Ocean"
Case "Passaic"
    abb = "Passaic"
Case "Ocean"
    abb = "Ocean"
Case "Middlesex"
    abb = "Middlesex"
Case "Mlddlesex"
    abb = "Mlddlesex"
Case "Morris"
    abb = "Morris"
    
    
End Select

ActiveCell.Offset(0, 1).Value = abb



End Sub

Public Sub Auto_loop()

Do Until ActiveCell.Value = ""
    County_Auto
    ActiveCell.Offset(1, 0).Select

Loop


End Sub

