VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   11964
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   18180
   OleObjectBlob   =   "CashFlowNPVCalculatorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'For easier iteration loops over arrays
Option Base 1


'Simulate and calculate the cash flow according to the amount of simulations indicated by the user
'Each cash flow parameter has a statistical distribution associated
Private Sub CommandButton1_Click()


'Declare variables, arrays and aux to create the model
Dim Ran As Double, i As Integer, alpha As Double, beta As Double, tWB As Workbook, alphaSR As Double, betaSR As Double, betadistSR As Double, a As Double, b As Double, c As Integer, Ppc As Double, Rtax As Double
Dim ProbProf As Double, R() As Double, datamin As Double, datamax As Double, datarange As Double, lowbins As Integer, highbins As Integer, nbins As Double, binrangeinit As Double, binrangefinal As Double
Dim bins() As Double, bincenters() As Double, j As Integer, bincounts() As Integer, ChartRange As String, nr As Integer


    'Save current workbook reference and set It as active
    Set tWB = ThisWorkbook
    tWB.Activate
    
    
    
    
    
    'Try catch behavior
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Main").Select
    
    
    
    
    'Dynamically expand the array size depending on the input number of simulations
    ReDim R(nsimulations) As Double
  
  
    'Loop over the number of simulations given by the user
    'A MONTECARLO ALGORITHM IS USED TO ESTIMATE THE CASH FLOW PROJECTION RESULT
    'Monte Carlo simulations aim to simulate the possible pathway of a future endeavor, experiment, or process, given multiple inputs that each have uncertainty.
    'These models are oftentimes used by financial planners to try to predict what might happen in the future.
    'Inputs to financial planning processes include probabilities and projections for various costs and amount of sales, and many other financial variables.
    For i = 1 To nsimulations.Value
    
    
    'Estimate the cost of land parameter and save the result in a spreadsheet cell for the model
    
        Ran = Rnd
        If Ran <= 1 * clandper1 / 100 Then
        
        Range("B3") = 1 * cland1
        
        ElseIf Ran <= 1 * clandper1 / 100 + 1 * clandper2 / 100 Then
        
        Range("B3") = 1 * cland2
        Else
        
        Range("B3") = 1 * cland3
        End If
        
        
        
    'Estimate the cost of royalties as a beta pert probability distribution
    
    alpha = (4 * croyalmod + 1 * croyalhigh - 5 * croyalow) / (1 * croyalhigh - 1 * croyalow)
    beta = (5 * croyalhigh - 1 * croyalow - 4 * croyalmod) / (1 * croyalhigh - 1 * croyalow)
    
    'Save the result in a spreadsheet cell
    Range("B4") = -1 * WorksheetFunction.Beta_Inv(Rnd, alpha, beta, -1 * croyalow, -1 * croyalhigh)
    
    
    'Calculate TDC as a normal distribution
    
    Range("B5") = WorksheetFunction.Norm_Inv(Rnd, 1 * tdpave, 1 * tdpstd)
    
    'Working capital calculation as a uniform distribution
    
    Range("B6") = 1 * wcmin + (1 * wcmax - 1 * wcmin) * Rnd
    
    'Calculate startup costs as a normal dist
     Range("B7") = WorksheetFunction.Norm_Inv(Rnd, scave, scstd)

    'Calculate sales revenue as a beta pert distribution
    
      alphaSR = (4 * srmode + 1 * srhigh - 5 * srlow) / (1 * srhigh - 1 * srlow)
    betaSR = (5 * srhigh - 1 * srlow - 4 * srmode) / (1 * srhigh - 1 * srlow)
    
    'Use a worksheet function to calculate the revenue and save It in a spreadsheet cell
    tWB.Worksheets("Main").Range("E3") = WorksheetFunction.Beta_Inv(Rnd, alphaSR, betaSR, srlow, srhigh)
   
    
    'Calculate production costs as a triangular distribution
    
    Ppc = Rnd
    
    'Estimate thr distribution parameters depending on the statistics input
If Ppc < (1 * pcmode - 1 * pclow) / (1 * pchigh - 1 * pclow) Then
    a = -1 * 1
    b = -1 * (-2 * pclow)
    c = -1 * (1 * pclow ^ 2 - Ppc * (1 * pcmode - 1 * pclow) * (1 * pchigh - 1 * pclow))
    
    'triangular inverse
    Range("H3") = (-b + Sqr(b ^ 2 - 4 * a * c)) / 2 / a
    
    
ElseIf Ppc <= 1 Then

    a = -1 * 1
    b = -1 * (-2 * pchigh)
    c = -1 * (1 * pchigh ^ 2 - (1 - 1 * Ppc) * (1 * pchigh - 1 * pclow) * (1 * pchigh - 1 * pcmode))
    
    'Left tail
    Range("H3") = (-b - Sqr(b ^ 2 - 4 * a * c)) / 2 / a
End If





    'Calculate the taxes based on a discrete distribution
      Rtax = Rnd
      
        If Rtax <= 1 * taxper1 / 100 Then
        
        Range("E4") = 1 * taxrate1.Value
        
        ElseIf Rtax <= 1 * taxper1 / 100 + 1 * taxper2 / 100 Then
        
        Range("E4") = 1 * taxrate2.Value
    End If
    
    'Calculate the interest rate as a using a Uniform distribution
    
    Range("H4") = 1 * irmin / 100 + (1 * irmax / 100 - 1 * irmin / 100) * Rnd
    
    '-----------------------------------------------------------
    
    
    
    'Save the clash flow result of each iteration in an array (ith simulation of n total simulations)
   R(i) = tWB.Worksheets("Main").Range("N24").Value
    
    
    'Update profitability counter for the model plot
    If tWB.Worksheets("Main").Range("N24").Value > 0 Then
    
    ProbProf = ProbProf + 1
    End If
    
    Next i
    
    
'When the model is calculate using the previous algorithm, the cash flow projections are displayed in a histogram plot on a worksheet:
'--------------------------------------------------------------------------------------------------------------------

'CREATE HISTOGRAM AND ASSIGN PLOT PARAMETERS:
    
datamin = WorksheetFunction.Min(R)
datamax = WorksheetFunction.Max(R)
datarange = datamax - datamin
lowbins = Int(WorksheetFunction.Log(nsimulations, 2)) + 1
highbins = Int(Sqr(nsimulations))
nbins = (lowbins + highbins) / 2
binrangeinit = datarange / nbins

ReDim bins(1) As Double


'Create histogram bar sizes according to frequencies calculated
If binrangeinit < 1 Then
    c = 1
    Do
        If 10 * binrangeinit > 1 Then
            binrangefinal = 10 * binrangeinit Mod 10
            Exit Do
        Else
            binrangeinit = 10 * binrangeinit
            c = c + 1
        End If
    Loop
    
    binrangefinal = binrangefinal / 10 ^ c
    
ElseIf binrangeinit < 10 Then

    binrangefinal = binrangeinit Mod 10
Else
    c = 1
    Do
        If binrangeinit / 10 < 10 Then
            binrangefinal = binrangeinit / 10 Mod 10
            Exit Do
        Else
            binrangeinit = binrangeinit / 10
            c = c + 1
        End If
    Loop
    binrangefinal = binrangefinal * 10 ^ c
End If
i = 1
bins(1) = (datamin - ((datamin) - (binrangefinal * Fix(datamin / binrangefinal))))
Do
    i = i + 1
    ReDim Preserve bins(i) As Double
    bins(i) = bins(i - 1) + binrangefinal
Loop Until bins(i) > datamax
nbins = i
ReDim Preserve bincounts(nbins - 1) As Integer
ReDim Preserve bincenters(nbins - 1) As Double
For j = 1 To nbins - 1
    c = 0
    For i = 1 To nsimulations
        If R(i) > bins(j) And R(i) <= bins(j + 1) Then
            c = c + 1
        End If
    Next i
    bincounts(j) = c
    bincenters(j) = (bins(j) + bins(j + 1)) / 2
Next j





'SAVE AND DISPLAY RESULT IN THE MODEL
Sheets("Histogram Data").Select
Cells.Clear
Range("A1").Select
Range("A1:A" & nbins - 1) = WorksheetFunction.Transpose(bincenters)
Range("B1:B" & nbins - 1) = WorksheetFunction.Transpose(bincounts)
UserForm1.Hide
Application.ScreenUpdating = False
Charts("Histogram").Delete
ActiveCell.Range("A1:B1").Select

'Set display properties of the diagram
    Range(Selection, Selection.End(xlDown)).Select
    nr = Selection.Rows.Count
    ChartRange = Selection.Address
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Histogram Data'!" & ChartRange)
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.PlotArea.Select
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).Delete
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).XValues = "='Histogram Data'!" & "$A$1:$A$" & nr
    ActiveChart.Legend.Select
    Selection.Delete
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    Selection.Caption = "Count"
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Caption = "Bin Center"
    ActiveChart.ChartArea.Select
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="Histogram"
    
'------------------
    
    
    
    'DISPLAY AND CALCULATE THE NET PRESENT VALUE USING THE RESULTS OF ALL THE ALGORITHM ITERATIONS:
    MsgBox " The percentage of simulations that resulted in positive net present value (NPV) is :" & " " & (ProbProf / (1 * nsimulations)) * 100 & "%"
    
    
    
End Sub






Private Sub CommandButton2_Click()
Unload UserForm1
End Sub



