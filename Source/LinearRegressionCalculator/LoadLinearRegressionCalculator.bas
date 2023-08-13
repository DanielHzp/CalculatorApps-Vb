Attribute VB_Name = "Module1"
Option Explicit
Option Base 1


'Execute linear regression calculator runtime
Sub RunForm()
UserForm1.Show
End Sub


Sub validation()

If IsNumeric(UserForm1.fxn1.Value) = True Then
    MsgBox " input has to be a function of x"
    ElseIf IsNumeric(UserForm1.fxn2.Value) = True Then
    MsgBox " input has to be a function of x"
    ElseIf IsNumeric(UserForm1.fxn3.Value) = True Then
    MsgBox " input has to be a function of x"
    ElseIf IsNumeric(UserForm1.fxn4.Value) = True Then
    MsgBox " input has to be a function of x"
    End If
    Exit Sub
    

End Sub




Sub test()
Dim test As String, X As Double, UserXRange As Range
test = "X^2+X+2*X^5"
MsgBox UserForm1.fxn3.Value
MsgBox UserForm1.fxn3.Text
Set UserXRange = Application.InputBox("X Input Range", "X Input", "Sheet1!$A$1:$A$10", Type:=8)
MsgBox UserXRange.Cells(5, 1)
 X = Evaluate(Replace("Ln(X)", "X", 2))
X = Evaluate(Replace("Sqrt(X)", "X", 2))
 X = Evaluate(Replace("1/X", "X", 2))
  X = Evaluate(Replace("1/X", "X", 2))
   X = Evaluate(Replace("1/X+X^2+X^4", "X", 2))
    X = Evaluate(Replace(test, "X", 2))
    MsgBox UserForm1.fxn1.Value

End Sub
