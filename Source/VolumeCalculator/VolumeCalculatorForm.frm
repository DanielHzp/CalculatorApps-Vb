VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TankForm 
   Caption         =   "Tank Calculator v. 3.7.2"
   ClientHeight    =   5628
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   8868.001
   OleObjectBlob   =   "VolumeCalculatorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TankForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Calculate volume when this button is clicked
Private Sub CalculateButton_Click()


Dim pi As Double, V As Double
pi = WorksheetFunction.pi()

'Data entry validations
'Input boxes cant be negative
If Ht < 0 Or Radius < 0 Or rho < 0 Or depth < 0 Then
MsgBox ("At least one of the inputs is negative, try again.")
    Exit Sub
End If


'Input boxes shoudnt be strings
If Not IsNumeric(Ht) Or Not IsNumeric(Radius) Or Not IsNumeric(rho) Or Not IsNumeric(depth) Then
MsgBox ("at least one of the inputx is not a number, please try again")
Exit Sub
End If



'No blank input boxes
If Ht = "" Or Radius = "" Or rho = "" Or depth = "" Then
MsgBox ("At least one of the inputs is missing, please try again")
    Exit Sub
End If


'If depth is greater than height throw alert
If 1 * depth > 1 * Ht Then
MsgBox ("fluid depth exceed the height of the tank,try again")
Exit Sub
End If




'If parameters are correct estimate the volumeof the shape
If depth <= Radius Then
    V = pi = depth ^ 2 / 3 * (3 * Radius - depth)
    
ElseIf depth <= Ht - Radius Then

    V = 2 / 3 * pi * Radius ^ 3 + pi * Radius ^ 2 * (depth - Radius)
    
Else

    V = 4 / 3 * pi * Radius ^ 3 + pi * Radius ^ 2 * (Ht - 2 * Radius) - pi * (Ht - depth) ^ 2 / 3 * (3 * Radius - Ht + depth)
End If


mass = rho * V

End Sub


'Close form
Private Sub QuitButton_Click()
Unload TankForm

End Sub
