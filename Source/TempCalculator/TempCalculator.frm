VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConversionForm 
   Caption         =   "Conversion Calculator v. 8.7.3"
   ClientHeight    =   4488
   ClientLeft      =   120
   ClientTop       =   444
   ClientWidth     =   7428
   OleObjectBlob   =   "TempCalculator.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "ConversionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Estimates temperature conversion based on measure units selected
Private Sub CalculateButton_Click()

'Declare variables
Dim low As Double, mid As Double, high As Double, Ke As Double, i As Integer
Dim flow As Double, fmid As Double, fhigh As Double, idx As Integer




'Validates the input fields make sense and checks numeric logic
If Pressure = "" Then
MsgBox "Pressure cannot be empty!"
Exit Sub
End If

If Not IsNumeric(Pressure) Then
MsgBox "Pressure must be numeric"

Exit Sub

End If

'Conversions go from bar to selected measure units
If Pressure < 0.001 Or Pressure > 1000 Then

MsgBox "Pressure must be between 0.001 and 1000 bar."

Exit Sub
End If


low = 0
high = 0.9999999


'Gets the index of the temperature selected in the dropdown list
idx = ConversionForm.TempSelect.ListIndex

'Sets the Ke conversion parameter based on the dropdown list selected value
'Ideally it should be parametrized
If idx = 0 Then

Ke = 0.00434

ElseIf idx = 1 Then

Ke = 0.000164

ElseIf idx = 2 Then

Ke = 0.0000451

ElseIf idx = 3 Then

Ke = 0.0000145

ElseIf idx = 4 Then

Ke = 0.00000538

ElseIf idx = 5 Then

Ke = 0.00000225

End If



'CONVERSION CALCULATED USING THE BISECTION METHOD
For i = 1 To 20


    mid = (low + high) / 2
    
    'call aux function to estimate conversion using input args
    flow = func(Pressure, Ke, low)
    fmid = func(Pressure, Ke, high)
    fhigh = func(Pressure, Ke, high)



    If flow * fmid < 0 Then
        high = mid
        Else
        low = mid
        End If
Next i



  'Output that is rendered in the userform
Conversion = FormatNumber(100 * (low + high) / 2, 1)



End Sub

'FunCtion that estimates the temperature conversion result.It is used recursively by the userform 'Calculate' button
Function func(P, Ke, x)
func = 16 * x ^ 2 * (2 - x) ^ 2 / 27 / (1 - x) ^ 4 / P ^ 2 - Ke

End Function



Private Sub QuitButton_Click()
Unload ConversionForm

End Sub

Private Sub ResetButton_Click()
Unload ConversionForm
'repopulate the temperature select
OpenUserForm
End Sub
