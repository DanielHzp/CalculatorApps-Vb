Attribute VB_Name = "Module1"
Option Explicit

'Populate dropdownlist

Sub OpenUserForm()

ConversionForm.TempSelect.AddItem "300"
ConversionForm.TempSelect.AddItem "400"
ConversionForm.TempSelect.AddItem "500"
ConversionForm.TempSelect.AddItem "550"
ConversionForm.TempSelect.AddItem "600"
ConversionForm.TempSelect.Text = "300"
ConversionForm.Show

End Sub
