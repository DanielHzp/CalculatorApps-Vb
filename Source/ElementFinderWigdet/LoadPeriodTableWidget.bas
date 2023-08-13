Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

'it can be shared in other subroutines used in the user form

Public n As Integer

'FILL COMBOBOX DROPDOWN MENU IN PERIODIC TABLE WIDGET
Sub PopulateComboBox()
Dim i As Integer
n = WorksheetFunction.CountA(Columns("A:A"))
For i = 2 To n
    UserForm1.ComboBox1.AddItem Range("A" & i)
Next i
UserForm1.ComboBox1.Text = Range("A2")
    

End Sub


Sub RunForm()

PopulateComboBox
UserForm1.Show
End Sub
