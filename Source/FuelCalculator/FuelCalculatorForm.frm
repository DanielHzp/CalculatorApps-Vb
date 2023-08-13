VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FuelCalculator 
   Caption         =   "Fuel Efficiency Calculator version 3.4.11"
   ClientHeight    =   5460
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   9120.001
   OleObjectBlob   =   "FuelCalculatorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FuelCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
'call GoButton_Click
End Sub


'Estimate fuel consumption conversions depending on the measure units selected

Private Sub GoButton_Click()
'RADIO BUTTONS ARE BOLEAN EATHIER TRUE OR FALSE
If miles Then

    If gal Then
    
        mpg = dist / vol
        
        kpl = (dist / 0.62) / (vol * 3.78)
    Else
        mpg = dist / vol * 3.78
        
        kpl = dist / 0.62 / vol
    End If
    
 Else
    If gal Then
    
        mpg = dist * 0.62 / vol
        
        kpl = dist / (vol * 3.78)
    Else
        mpg = dist * 0.62 / vol * 3.78
        
        kpl = dist / vol
    End If
    
 End If

End Sub

Private Sub QuitButton_Click()

Unload FuelCalculator

End Sub
