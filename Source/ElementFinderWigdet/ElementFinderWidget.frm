VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Element Finder"
   ClientHeight    =   4788
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   6420
   OleObjectBlob   =   "ElementFinderWidget.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'For easier loop iterations
Option Base 1

'Execute search of the element selected in the dropdown list
Private Sub ComboBox1_Change()
'WHEN ITEM IN COMBO BOX ES SELECTED  THIS REEVALUATES WHATS INSIDE
'NO NEED TO PRESS GO BUTTON
Dim index As Integer


'Obtain index from combo box
index = UserForm1.ComboBox1.ListIndex


symbol = Range("A2:D" & n).Cells(index + 1, 2)
atomicnumber = Range("A2:D" & n).Cells(index + 1, 3)
atomicmass = Range("A2:D" & n).Cells(index + 1, 4)
End Sub


'Obtain the element object selected in the dropdown list
Private Sub GoButton_Click()
Dim index As Integer


'Obtain index from the combo box to search the element parameters in the dataset (periodic table)
index = UserForm1.ComboBox1.ListIndex
symbol = Range("A2:D" & n).Cells(index + 1, 2)
atomicnumber = Range("A2:D" & n).Cells(index + 1, 3)
atomicmass = Range("A2:D" & n).Cells(index + 1, 4)


End Sub



Private Sub QuitButton_Click()
Unload UserForm1

End Sub
