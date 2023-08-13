VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Regression Toolbox"
   ClientHeight    =   6000
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   8016
   OleObjectBlob   =   "LinearRegressionCalculator.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Option Base 1




'This app automates the calculation of regression models that best fit the input equations (max.4 due to computational constraints)
'This solution lets the user predict the behavior of any model (using one variable max.) based on the best regression fit automatically calculated






'This method creates the regression model based on input expressions f(x)
Private Sub GoButton_Click()

'Declare variables and and aux
Dim tWB As Workbook, ans As Variant, regfit As String, temp
Dim UserXRange As Range, UserYRange As Range, X() As Variant, Y() As Variant, nrowsX As Integer, nrowsY As Integer, i As Integer, j As Integer, k As Integer
Dim Xt As Variant, XtX As Variant, XtXinv As Variant, B As Variant, XtY As Variant



'PLACE YOUR ADDITIONAL DIM STATEMENTS IN THIS REGION

'Try-catch behavior
On Error GoTo here



'Input expressions f(x) validations for model fit
If IsNumeric(UserForm1.fxn1.Value) = True Then
    MsgBox " input has to be a function of x"
    Exit Sub
    ElseIf IsNumeric(UserForm1.fxn2.Value) = True Then
    MsgBox " input has to be a function of x"
    Exit Sub
    ElseIf IsNumeric(UserForm1.fxn3.Value) = True Then
    MsgBox " input has to be a function of x"
    Exit Sub
    ElseIf IsNumeric(UserForm1.fxn4.Value) = True Then
    MsgBox " input has to be a function of x"
    Exit Sub
    End If
    
    
'Save reference to current workbook
Set tWB = ThisWorkbook
tWB.Activate





'THE FOLLOWING TWO LINES JUST SETS A DEFAULT RANGE IN THE INPUT BOXES, THAT'S ALL

Set UserXRange = Application.InputBox("X Input Range", "X Input", "Sheet1!$A$1:$A$10", Type:=8)
Set UserYRange = Application.InputBox("Y Input Range", "Y Input", "Sheet1!$B$1:$B$10", Type:=8)



'for j=2 to number of functions (the first column is 1)
'Evaluate(Replace(UserForm1.Controls("fxn" & j - 1).Value, "x", smallx(i))
'determine number of input functions:
For j = 1 To 4
    If UserForm1.Controls("fxn" & j).Value <> "" Then
    k = k + 1
    End If
Next j

 MsgBox "number of functions input:" & k
 
 
 
 
 

 
If k = 0 Then MsgBox " Enter at least one function "







'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'PLACE THE MAIN BULK OF YOUR CODE IN THIS REGION!
'try to populate matrix X CONSIDERING 4 FUNCTIONS MAX


'count number of data points
 nrowsX = UserXRange.Rows.Count
 nrowsY = UserYRange.Rows.Count
 
 If nrowsY < k + 2 Then
 MsgBox "There arent at least num functions+2 data points"
 Exit Sub
 End If
 
 'Update the size (iterations loop) of the array that stores the regression model calculation
 ReDim X(nrowsX, k + 1) As Variant
 ReDim Y(nrowsY, 1) As Variant
 
 
 For i = 1 To nrowsX
    X(i, 1) = 1
 Next i
 
 'No blank input formula boxes , let all nonempty input boxes CONCATENATE in a sequence
For j = 2 To 4
    For i = 2 To 4
    
        If Not UserForm1.Controls("fxn" & i - 1).Value <> "" Then
        temp = UserForm1.Controls("fxn" & i).Value
        
        UserForm1.Controls("fxn" & i).Value = UserForm1.Controls("fxn" & i - 1).Value
        UserForm1.Controls("fxn" & i - 1).Value = temp
        
        End If
 Next i
 Next j
 
 
 
'Populate vector X
    
For i = 1 To nrowsX
    For j = 2 To k + 1
    If UserForm1.Controls("fxn" & j - 1).Value <> "" Then
    X(i, j) = Evaluate(Replace(UserForm1.Controls("fxn" & j - 1).Value, "x", UserXRange.Cells(i, 1)))
    End If
    Next j
Next i

'TRANSFORM THE ARRAYS TO FIT THE LINEAR REGRESSION MODEL STRUCTURE:
Xt = WorksheetFunction.Transpose(X)
XtX = WorksheetFunction.MMult(Xt, X)
XtXinv = WorksheetFunction.MInverse(XtX)




'Populate vector Y
For j = 1 To nrowsY
    Y(j, 1) = UserYRange.Cells(j, 1)
Next j

XtY = WorksheetFunction.MMult(Xt, Y)
B = WorksheetFunction.MMult(XtXinv, XtY)

'Regressions equation start point
regfit = "Y" & "=" & B(1, 1)
For i = 1 To 4

    If UserForm1.Controls("fxn" & i).Value = "" Then Exit For
   
   'Recursively create the regression equation with the input linear expressions
  regfit = regfit & "+" & B(i + 1, 1) & "*" & UserForm1.Controls("fxn" & i).Text
  
Next i



'DISPLAY THE REGRESSION MODEL EQUATION RESULT:
MsgBox regfit


'The code below creates and populates the linear regression output PLOT
'-----------------------------------------------------------------------------------------------------------------
Dim SSE As Double, SST As Double, par As Integer, n As Integer, resq() As Variant, aux As Variant, Yexp As String
Dim YminusYave(), RAdj As Double, averg As Double
par = k + 1
n = nrowsX


ReDim resq(n, 1)
ReDim YminusYave(n, 1)


'Obtain the equation output in terms of x
aux = Split(regfit, "=")

Yexp = aux(1)

'Calculate SST, SSE ANOVA parameters of the regression model estimated
For i = 1 To n
    averg = averg + UserYRange.Cells(i, 1)
    resq(i, 1) = (Evaluate(Replace(Yexp, "x", UserXRange.Cells(i, 1))) - UserYRange.Cells(i, 1)) ^ 2
    SSE = resq(i, 1) + SSE
Next i
For i = 1 To n
    YminusYave(i, 1) = (UserYRange.Cells(i, 1) - (averg / n)) ^ 2
    SST = SST + YminusYave(i, 1)
Next i


'Estimate SSR based on SSE and SST
RAdj = 1 - (SSE / (n - par)) / (SST / (n - 1))


'DISPLAY RESULT:
MsgBox "R^2 Adjusted :" & RAdj





'Display the regression model PLOT:
'---------------------------------------------------------------------------------------------------------
Dim ypVector() As Variant
ReDim ypVector(nrowsY, 1)

    For j = 1 To nrowsY
        ypVector(j, 1) = Evaluate(Replace(Yexp, "x", UserXRange.Cells(j, 1)))
    Next j
    
ans = MsgBox("Would you like to plot the data?", vbYesNo)


'IF USERS SELECTS YES
If ans = 6 Then
      
      
      
    'SET PLOT PARAMETERS
    
    
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.SetSourceData Source:=UserYRange
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    Selection.Caption = "y"
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Caption = "x"
    ActiveChart.ChartArea.Select
    ActiveChart.PlotArea.Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).XValues = UserXRange
    ActiveChart.FullSeriesCollection(2).Values = ypVector
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
    End With
    Selection.MarkerStyle = -4142
    ActiveChart.FullSeriesCollection(2).Smooth = True
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.PlotArea.Select
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).Name = "=""Experimental Data"""
    ActiveChart.FullSeriesCollection(2).Name = "=""Model prediction"""
End If







MsgBox "Quit to see the plot"
Exit Sub





here:
MsgBox "Try again, error"

End Sub










Private Sub QuitButton_Click()
'PLACE YOUR CODE HERE
 'MsgBox UserForm1.fxn1.Value
Unload UserForm1
End Sub


