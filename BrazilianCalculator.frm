VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Brazilian Test Calculator"
   ClientHeight    =   11460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16470
   OleObjectBlob   =   "BrazilianCalculator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Calculate_Click()

Dim i As Integer

For i = 1 To 10
    If Controls("T" & i).Value = "" Then

Controls("O" & i).Value = 0

    Else

Controls("O" & i).Value = Application.Evaluate(2 * Controls("P" & i).Value / (3.141592 * Controls("T" & i).Value * Controls("L" & i).Value))
Controls("O" & i).Value = Round(Controls("O" & i).Value)


    End If

Next






End Sub

Private Sub CommandButton1_Click()

For i = 1 To 10
Controls("T" & i).Value = ""
Controls("P" & i).Value = ""
Controls("L" & i).Value = ""
Controls("O" & i).Value = ""
Next
End Sub
