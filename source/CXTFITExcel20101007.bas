Attribute VB_Name = "Module1"
'-----------------------------------------------------------------------------------------------
'   CXTFIT/Excel
'   October 07, 2010

'   Guoping Tang
'   Melanie A. Mayes
'   Phil M. Jardine
'   Environmental Sciences Division
'   Oak Ridge National Laboratory
'   PO Box 2008
'   Oak Ridge, TN 37831-6038
'   Tel: 865-574-7314
'   Fax: 865-576-8646
'   Email: tangg@ornl.gov
'          mayesma@ornl.gov
'          jardinepm@ornl.gov

'   Jack C. Parker
'   Department of Civil and Environmental Engineering
'   University of Tennessee
'   62 Perkins Hall
'   Knoxville, TN 37996-2010
'   865-974-7718
'   jparker@utk.edu
'-----------------------------------------------------------------------------------------------

Option Explicit
Public Const PI = 3.1415926


'-----------------------------------------------------------------------------------------------
' Interfaces (Menu and Dialogs)
'-----------------------------------------------------------------------------------------------
Public Sub AddMenu()
    Dim MenuObject As CommandBarPopup
    Dim MenuItem As Object
    
    Call DeleteMenu
    
    Set MenuObject = Application.CommandBars(1).Controls.Add( _
                     Type:=msoControlPopup, _
                     Before:=10, _
                     temporary:=True)
                     
    MenuObject.Caption = "CXTFIT"
    
    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
    MenuItem.OnAction = "SolveDialog"
    MenuItem.Caption = "Solve..."
    
    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
    MenuItem.OnAction = "JacobianDialog"
    MenuItem.Caption = "Calculate sensitivity..."
    
    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
    MenuItem.OnAction = "PropagateDialog"
    MenuItem.Caption = "Propagate parameter uncertainty..."

    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
    MenuItem.OnAction = "MonteCarloDialog"
    MenuItem.Caption = "Monte Carlo analysis..."

    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
    MenuItem.OnAction = "ResponseDialog"
    MenuItem.Caption = "Calculate response surface..."

'    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
'    MenuItem.OnAction = "DeleteMenu"
'    MenuItem.Caption = "Hide CXTFIT menu..."

    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
    MenuItem.OnAction = "AboutDialog"
    MenuItem.Caption = "About..."
End Sub

Public Sub DeleteMenu()
    On Error Resume Next
    Application.CommandBars(1).Controls("CXTFIT").Delete
    On Error GoTo 0
End Sub

Sub SolveDialog()
    DlgSolve.Show
End Sub

Sub JacobianDialog()
    DlgJacobian.Show
End Sub

Sub PropagateDialog()
    DlgPropagate.Show
End Sub

Sub MonteCarloDialog()
    DlgMonteCarloAnalyze.Show
End Sub

Sub ResponseDialog()
    DlgResponse.Show
End Sub

Sub AboutDialog()
    DlgAbout.Show
End Sub

'-----------------------------------------------------------------------------------------------
' Interfaces (Macros)
'-----------------------------------------------------------------------------------------------


Public Sub Solve()
    Call Optimize
    Call Analyze
End Sub


Public Sub Optimize()
' CXTFIT/Excel macro to call Solver Add-In to perform optimization
' with the defined name ObjFuncCell and ParameterRange
'
    If (Not (NameExists("ObjFuncCell", "Local"))) Then
        MsgBox ("Local name ObjFuncCell is not defined!!!")
        Exit Sub
    End If
    
    If (Not (NameExists("ParameterRange", "Local"))) Then
        MsgBox ("Local Name ParameterRange is not defined!!!")
        Exit Sub
    End If
    
    Dim I As Integer
    
    SolverReset
    
    If (NameExists("ObjFuncCell", "Local")) Then
        SolverOk SetCell:=Range("ObjFuncCell"), _
                 ByChange:=Range("ParameterRange"), _
                 MaxMinVal:=2
    End If
        
    If (Not (NameExists("NoParaConstraint", "Local"))) Then
        For I = 1 To Range("ParameterRange").Count
            SolverAdd CellRef:=Range("ParameterRange").Cells(I, 1), _
                      Relation:=3, _
                      FormulaText:=Range("ParameterRange").Cells(I, 2)
            SolverAdd CellRef:=Range("ParameterRange").Cells(I, 1), _
                      Relation:=1, _
                      FormulaText:=Range("ParameterRange").Cells(I, 3)
        Next
    End If
    
    If (Not (NameExists("SolverUserFinish", "Local"))) Then
        SolverSolve UserFinish:=False
    Else
        SolverSolve UserFinish:=True
    End If

End Sub

Public Sub Analyze()
                  
    Dim objv As Double
    Dim Penalty As Double
    Dim ParameterRange As Range
    Dim PredictionRange As Range
        
    If (Not (NameExists("ObjFuncCell", "Local"))) Then
        MsgBox ("Local name ObjFuncCell is not defined!!!")
        Exit Sub
    Else
        objv = Range("ObjFuncCell").Value
    End If
    
    If (Not (NameExists("ParameterRange", "Local"))) Then
        MsgBox ("Local name ParameterRange is not defined!!!")
        Exit Sub
    Else
        Set ParameterRange = Range("ParameterRange")
    End If
    
    If (Not (NameExists("PredictionRange", "Local"))) Then
        MsgBox ("Local name PredictionRange is not defined!!!")
        Exit Sub
    Else
        Set PredictionRange = Range("PredictionRange")
    End If

    Dim OffsetPredJcb, OffsetParaStdOut As Integer
    
    If (Not (NameExists("OffsetParaStdOut", "Local"))) Then
        OffsetParaStdOut = 3
    Else
        OffsetParaStdOut = Evaluate("OffsetParaStdOut")
    End If

    If (Not (NameExists("OffsetPredJcb", "Local"))) Then
        OffsetPredJcb = 0
    Else
        OffsetPredJcb = Evaluate("OffsetPredJcb")
    End If
    
    Dim OffsetParaPtb, OffsetParaPriorStd, OffsetPredObsStd As Integer
    
    If (Not (NameExists("OffsetParaPtb", "Local"))) Then
        OffsetParaPtb = 0
    Else
        OffsetParaPtb = Evaluate("OffsetParaPtb")
    End If
    
    If (Not (NameExists("OffsetParaStd", "Local"))) Then
        OffsetParaPriorStd = 0
    Else
        OffsetParaPriorStd = Evaluate("OffsetParaStd")
    End If
    
    If (Not (NameExists("OffsetPredStd", "Local"))) Then
        OffsetPredObsStd = 0
    Else
        OffsetPredObsStd = Evaluate("OffsetPredStd")
    End If
    
    If (Not (NameExists("PenaltyCell", "Local"))) Then
        Penalty = 1#
    Else
        Penalty = Range("PenaltyCell").Value
    End If
    
    If (NameExists("ObjFuncCell", "Local")) Then
        objv = Range("ObjFuncCell").Value
    End If
    
    If (Not (NameExists("ParameterRange", "Local"))) Then
        MsgBox ("Local name ParameterRange is not defined!!!")
        Exit Sub
    End If
    
    'Check PredictionRange
    Dim nPara, nData As Integer
    nPara = Range("ParameterRange").Count    'number of parameters to be estimated
    nData = Range("PredictionRange").Count   'number of observations
    
    ReDim CovarianceMatrix(1 To nPara, 1 To nPara) As Double
    Dim I, J As Integer
    Dim CorrelationRange As Range
    Dim dtmp As Double
    Dim MSE As Double
    
    If (nPara > nData) Then
        MsgBox ("The number of observations is less than the number of estimated parameters!!!")
        Return
    End If
    
    ReDim JacobianMatrix(1 To nData + nPara, 1 To nPara) As Double 'derivative
 '   ReDim CovarianceMatrix(1 To nPara, 1 To nPara) As Double
    ReDim CorrelationMatrix(1 To nPara, 1 To nPara) As Double
    ReDim StandardDeviation(1 To nPara, 1 To 1) As Double
    ReDim St(1 To nPara, 1 To nData) As Double
    ReDim Sw(1 To nPara, 1 To nData) As Double
    ReDim Jt(1 To nPara, 1 To nData + nPara) As Double
    ReDim aa(1 To nPara, 1 To nPara) As Double
    ReDim Bb(1 To nPara, 1 To nData + nPara) As Double
    ReDim iA(1 To nPara, 1 To nPara) As Double
    ReDim WeightMatrixC(1 To nData, 1 To nData) As Double
    
    ' compute sensitivity matrix related to observations
    Call CalculateJacobian(JacobianMatrix)
    
    'compute covariance matrix
    'weight for observation
    For I = 1 To nData
        If (IsMissing(OffsetPredObsStd)) Then
            WeightMatrixC(I, I) = 1#
        ElseIf (OffsetPredObsStd = 0) Then
            WeightMatrixC(I, I) = 1#
        Else
            dtmp = PredictionRange.Offset(0, OffsetPredObsStd).Cells(I, 1)
            WeightMatrixC(I, I) = 1# / dtmp / dtmp
        End If
    Next I
    
    Call MatrixTranspose(JacobianMatrix, nData, nPara, St)
    Call MatrixMultiplication(St, nPara, nData, WeightMatrixC, nData, Sw)
    Call MatrixMultiplication(Sw, nPara, nData, JacobianMatrix, nPara, aa)
        
    If (Not (IsMissing(OffsetParaPriorStd))) Then
         If (Not (OffsetParaPriorStd = 0)) Then
        'weight for parameters in penalty function
            If (IsMissing(Penalty) Or (Abs(Penalty) < 0.00000001)) Then
                Penalty = 1#
            End If
        
            For I = 1 To nPara
                dtmp = ParameterRange.Offset(0, OffsetParaPriorStd).Cells(I, 1).Value
                aa(I, I) = aa(I, I) + Penalty / dtmp / dtmp
            Next I
            
            MSE = 1#
        Else
            MSE = objv / (nData - nPara)
        End If
    Else
        MSE = objv / (nData - nPara)
    End If
    
    Call MatrixInverse(aa, nPara, iA)
    
    'calculate the covariance matrix
    For I = 1 To nPara
        For J = 1 To nPara
            CovarianceMatrix(I, J) = MSE * iA(I, J)
        Next J
    Next I
    
        'calculate the standard deviation
    For I = 1 To nPara
        If (CovarianceMatrix(I, I) < 0) Then
             MsgBox ("Standard deviation of estimated parameter is too close to zero!")
             StandardDeviation(I, 1) = -99
        Else
             StandardDeviation(I, 1) = Sqr(CovarianceMatrix(I, I))
        End If
    Next I
        
    'write the standard deviation for estimated parameters
    ParameterRange.Offset(0, OffsetParaStdOut).Value = StandardDeviation
    
    'calculate the correlation matrix
    For I = 1 To nPara
        For J = 1 To nPara
            If (Abs(StandardDeviation(I, 1)) < 0.000000001 Or Abs(StandardDeviation(J, 1)) < 0.000000001) Then
                MsgBox ("Standard deviation for estimated parameters is close to zero, correlation is set to-999")
                CorrelationMatrix(I, J) = -999
            Else
                CorrelationMatrix(I, J) = CovarianceMatrix(I, J) / StandardDeviation(I, 1) / StandardDeviation(J, 1)
            End If
        Next J
    Next I
    
    Set CorrelationRange = ParameterRange.Offset(0, OffsetParaStdOut + 1)
   
    For I = 2 To nPara
        Set CorrelationRange = Union(CorrelationRange, ParameterRange.Offset(0, I + OffsetParaStdOut))
    Next I

    CorrelationRange.Value = CorrelationMatrix
    CorrelationRange.NumberFormat = "0.000"
    
End Sub

'Calculate and output Jacobian matrix to the range defined by OffPredJcb offset from PredictionRange
Public Sub GetJacobianMatrix()
    'Check OffsetPredJcb
    Dim OffsetPredJcb As Integer
    
    If (Not (NameExists("OffsetPredJcb", "Local"))) Then
        OffsetPredJcb = 0
        MsgBox ("Local name OffsetPredJcb is not defined!!!")
        Exit Sub
    Else
        OffsetPredJcb = Evaluate("OffsetPredJcb")
        If (OffsetPredJcb = 0) Then
            MsgBox ("OffsetPredJcb = 0 !!!")
            Exit Sub
        End If
    End If
    
    'Check ParameterRange
    If (Not (NameExists("ParameterRange", "Local"))) Then
        MsgBox ("Local name ParameterRange is not defined!!!")
        Exit Sub
    End If
    
    'Check PredictionRange
    If (Not (NameExists("PredictionRange", "Local"))) Then
        MsgBox ("Local name PredictionRange is not defined!!!")
        Exit Sub
    End If
    
    'Define the size of Jacobian matrix
    
    Dim nPara, nData As Integer
        
    nPara = Range("ParameterRange").Count    'number of parameters to be estimated
    nData = Range("PredictionRange").Count   'number of observations
    
    ReDim JacobianMatrix(1 To nData, 1 To nPara) As Double  'derivative
    
    'Calculate Jacobian matrix
    Call CalculateJacobian(JacobianMatrix)
    
End Sub

'Calculate Jacobian matrix
Public Sub CalculateJacobian(ByRef JacobianMatrix() As Double)

    If (Not (NameExists("ParameterRange", "Local"))) Then
        MsgBox ("Local name ParameterRange is not defined!!!")
        Exit Sub
    End If
    
    If (Not (NameExists("PredictionRange", "Local"))) Then
        MsgBox ("Local name PredictionRange is not defined!!!")
        Exit Sub
    End If

    Dim OffsetParaPtb, OffsetPredJcb As Integer
    
    If (Not (NameExists("OffsetParaPtb", "Local"))) Then
        OffsetParaPtb = 0
    Else
        OffsetParaPtb = Evaluate("OffsetParaPtb")
    End If
    
    Dim I, J, nPara, nData As Integer
    Dim ParaOldValue, dtmp, Delta As Double
    
    nPara = Range("ParameterRange").Count    'number of parameters to be estimated
    nData = Range("PredictionRange").Count   'number of observations
    
    ReDim JacobianMatrix(1 To nData, 1 To nPara) As Double  'derivative
    ReDim PredMeanValue(1 To nData) As Variant                      'model predition for average P
    ReDim PredPerturbationValue(1 To nData) As Variant              'model prediction for average P + 0.01P
    
    PredMeanValue = Range("PredictionRange").Value
    
    For I = 1 To nPara
'       parameter value
        ParaOldValue = Range("ParameterRange").Cells(I, 1).Value
        
        If (Abs(ParaOldValue) < 0.0000000001) Then
            Delta = 0.01    'to avoid zero perturbation for parameters of zero value
        Else
            'perturbation is not specified
            If (IsMissing(OffsetParaPtb)) Then
                Delta = ParaOldValue * 0.01
            ElseIf (OffsetParaPtb = 0) Then
                Delta = ParaOldValue * 0.01
            Else
                Delta = ParaOldValue * Range("ParameterRange").Offset(0, OffsetParaPtb).Cells(I, 1).Value
            End If
        End If
        
        Range("ParameterRange").Cells(I, 1).Value = ParaOldValue + Delta
 
'    derivative
        PredPerturbationValue = Range("PredictionRange").Value
        
        For J = 1 To nData
            JacobianMatrix(J, I) = (PredPerturbationValue(J, 1) - PredMeanValue(J, 1)) / Delta
        Next J
        
        Range("ParameterRange").Cells(I, 1).Value = ParaOldValue
    Next I
    
    'output Jacobian matrix
    If (Not (NameExists("OffsetPredJcb", "Local"))) Then
        Exit Sub
    Else
        OffsetPredJcb = Evaluate("OffsetPredJcb")
    End If
    
    Dim JacobianRange As Range
    Set JacobianRange = Range("PredictionRange").Offset(0, OffsetPredJcb)
    For I = 2 To nPara
        Set JacobianRange = Union(JacobianRange, Range("PredictionRange").Offset(0, I + OffsetPredJcb - 1))
    Next I
    
    JacobianRange.Value = JacobianMatrix

End Sub

Public Sub Propagate()
                  
    Dim ParameterRange As Range
    Dim PredictionRange As Range
    
    If (Not (NameExists("ParameterRange", "Local"))) Then
        MsgBox ("Local name ParameterRange is not defined!!!")
        Exit Sub
    Else
        Set ParameterRange = Range("ParameterRange")
    End If
    
    If (Not (NameExists("PredictionRange", "Local"))) Then
        MsgBox ("Local name PredictionRange is not defined!!!")
        Exit Sub
    Else
        Set PredictionRange = Range("PredictionRange")
    End If
    
    Dim OffsetPredEP As Integer
    
    If (Not (NameExists("OffsetPredEP", "Local"))) Then
        MsgBox ("Local name OffsetPredEP is not defined!!!")
        Exit Sub
    Else
        OffsetPredEP = Evaluate("OffsetPredEP")
    End If
    
    Dim OffsetParaStdOut As Integer
    
    If (Not (NameExists("OffsetParaStdOut", "Local"))) Then
        OffsetParaStdOut = 3
    Else
        OffsetParaStdOut = Evaluate("OffsetParaStdOut")
    End If
    
    Dim I, J, nPara, nData As Integer
    Dim CorrelationRange As Range
        
    nPara = ParameterRange.Count    'number of parameters to be estimated
    nData = PredictionRange.Count   'number of observations
    
    ReDim JacobianMatrix(1 To nData + nPara, 1 To nPara) As Double 'derivative
    ReDim CovarianceMatrix(1 To nPara, 1 To nPara) As Double
    ReDim CorrelationMatrix(1 To nPara, 1 To nPara) As Double
    ReDim StandardDeviation(1 To nPara, 1 To 1) As Double
    
    ReDim s2Para(1 To nData, 1 To 1) As Double
    ReDim TA(1 To 1, 1 To nPara) As Double
    ReDim tAt(1 To nPara, 1 To 1) As Double
    ReDim tM1(1 To 1, 1 To nPara) As Double
    ReDim tM2(1 To 1, 1 To 1) As Double
    
    ' calculate Jacobian matrix
    Call CalculateJacobian(JacobianMatrix)

   'Read the standard deviation
    For I = 1 To nPara
        StandardDeviation(I, 1) = ParameterRange.Offset(0, OffsetParaStdOut).Cells(I, 1).Value
    Next I
      
   'Read the correlation matrix
   For J = 1 To nPara
       For I = 1 To nPara
          CorrelationMatrix(I, J) = ParameterRange.Offset(0, OffsetParaStdOut + J).Cells(I, 1).Value
       Next I
   Next J
      
    'Calculate covariance matrix
    For I = 1 To nPara
        For J = 1 To nPara
            CovarianceMatrix(I, J) = StandardDeviation(I, 1) * StandardDeviation(J, 1) * CorrelationMatrix(I, J)
        Next J
    Next I
            
    'Calculate error propagation
    
    'calculate uncertainty due to estimated parameter uncertainty
    For I = 1 To nData
          For J = 1 To nPara
               TA(1, J) = JacobianMatrix(I, J)
               tAt(J, 1) = TA(1, J)
          Next J
            
          Call MatrixMultiplication(TA, 1, nPara, CovarianceMatrix, nPara, tM1)
          Call MatrixMultiplication(tM1, 1, nPara, tAt, 1, tM2)
             
          s2Para(I, 1) = tM2(1, 1)
    Next I
    
    PredictionRange.Offset(0, OffsetPredEP).Value = s2Para
End Sub

Sub MonteCarloAnalyze()
    
    Dim I, J, nPara, nInput, nCol As Integer
    
    Dim ParaRange As Range
    Dim InputRange As Range
    Dim ObjFuncCell As Range
    Dim CurCell As Range
    Dim DisplayCur As Boolean
    
    DisplayCur = False
    
    If (Not (NameExists("ParameterRange", "Local"))) Then
        MsgBox ("Local name ParameterRange is not defined!!!")
        Exit Sub
    Else
        Set ParaRange = Range("ParameterRange")
        nPara = ParaRange.Count
    End If
    
    If (Not (NameExists("MCParaInputRange", "Local"))) Then
        MsgBox ("Local name MCParaInputRange is not defined!!!")
        Exit Sub
    Else
        Set InputRange = Range("MCParaInputRange")
        nCol = InputRange.Columns.Count
        nInput = InputRange.Rows.Count
        If (nPara > nCol) Then
            MsgBox ("The specified parameter input range does not have enough columns for the parameter inputs!!!")
            Exit Sub
        ElseIf (nPara < nCol) Then
            MsgBox ("The specified parameter input range has not have more columns for the parameter inputs!!!")
        End If
    End If
    
    If (Not (NameExists("MCMeritCell", "Local"))) Then
        MsgBox ("Local name MCMeritCell is not defined!!!")
        Exit Sub
    Else
        Set ObjFuncCell = Range("MCMeritCell")
    End If
    
    If (NameExists("MCCurrentCell", "Local")) Then
        Set CurCell = Range("MCCurrentCell")
        DisplayCur = True
    End If
    
    If (nInput > 1000) Then
        If (MsgBox("It may take a long time to finish Monte Carlo simulation. You may stop at any time using Esc", vbOKCancel) = vbCancel) Then
        Exit Sub
        End If
    End If
    
    With Application
        .Calculation = xlCalculationManual
    End With
    
    For I = 1 To nInput
        If (DisplayCur) Then
            CurCell.Value = I
        End If
            
        For J = 1 To nPara
            ParaRange.Cells(J).Value = InputRange(I, J)
        Next
                
        Application.Calculate
        
        InputRange.Cells(I, nPara).Offset(0, 1).Value = ObjFuncCell.Value
    Next
    
    With Application
        .Calculation = xlCalculationAutomatic
    End With
    
End Sub

Sub CalculateResponse()
    Dim I, J, nRowInput, nColInput, nColOffset As Integer
    Dim ParaRange As Range
    Dim InputRowRange As Range
    Dim InputColRange As Range
    Dim ObjFuncCell As Range
    
    If (Not (NameExists("ParameterRange", "Local"))) Then
        MsgBox ("Local name ParameterRange is not defined!!!")
        Exit Sub
    Else
        Set ParaRange = Range("ParameterRange")
    End If
    
    If (Not (NameExists("RSVRange", "Local"))) Then
        MsgBox ("Local name RSVRange is not defined!!!")
        Exit Sub
    Else
        Set InputRowRange = Range("RSVRange")
        nRowInput = InputRowRange.Rows.Count
    End If
    
    If (Not (NameExists("RSHRange", "Local"))) Then
        MsgBox ("Local name RSHRange is not defined!!!")
        Exit Sub
    Else
        Set InputColRange = Range("RSHRange")
        nColInput = InputColRange.Columns.Count
    End If
    
    If (Not (NameExists("RSMeritCell", "Local"))) Then
        MsgBox ("Local name RSMeritCell is not defined!!!")
        Exit Sub
    Else
        Set ObjFuncCell = Range("RSMeritCell")
    End If
    
    nColOffset = InputColRange.Cells(1, 1).Column - InputRowRange.Cells(1, 1).Column
    
    If (nColOffset < 0) Then
        MsgBox ("Column range starts at the left of the row range!")
        Exit Sub
    End If
    
    With Application
        .Calculation = xlCalculationManual
    End With
    
    For I = 1 To nRowInput
        For J = 1 To nColInput
            ParaRange.Cells(1).Value = InputRowRange.Cells(I, 1).Value
            ParaRange.Cells(2).Value = InputColRange.Cells(1, J).Value
                
            Application.Calculate
        
            InputRowRange.Cells(I, 1).Offset(0, J + nColOffset - 1).Value = ObjFuncCell.Value
        
            ObjFuncCell.Offset(1, 0).Value = I
            ObjFuncCell.Offset(2, 0).Value = J
        Next
    Next
    
    With Application
        .Calculation = xlCalculationAutomatic
    End With
    
End Sub

'-----------------------------------------------------------------------------------------------
' CDE solution
'-----------------------------------------------------------------------------------------------


'Solving equilibrium convection-diffusion equation with first order decay and production
' R dc/dt = D d2c/dx2 - v dc/dx - mu c + gamma
' initial condition c(x, 0) = ci
' inlet boundary condition vc(0, t)-Ddc(0,t)/dx = vc0(t)
' outlet boundary condition Ddc(infinity, t)/dx = 0
' Eq. 11 in Parker and van Genuchten 1984)

Public Function CDE(ByVal x As Double, _
                    ByVal T As Double, _
                    ByVal t0 As Double, _
                    ByVal v As Double, _
                    ByVal D As Double, _
                    Optional ByVal R As Variant, _
                    Optional ByVal mu As Variant, _
                    Optional ByVal gamma As Variant, _
                    Optional ByVal strFluxConc As Variant, _
                    Optional ByVal ci As Variant, _
                    Optional ByVal c0 As Variant) As Double


If IsMissing(R) Then
    R = 1#
End If

If IsMissing(mu) Then
    mu = 0#
End If

If IsMissing(gamma) Then
    gamma = 0#
End If

If IsMissing(strFluxConc) Then
    strFluxConc = "flux"
End If
             
If IsMissing(ci) Then
    ci = 0#
End If
             
If IsMissing(c0) Then
    c0 = 1#
End If
             
Select Case LCase(strFluxConc)
    Case "flux"
        'flux average concentration, using Eqs in 3.3 and 3.4 in
        'Parker and van Genuchten 1984
        If (Abs(mu) < 0.0000000001) Then
            CDE = ci + (c0 - ci) * AXTFU0(v, D, R, x, T) _
                     + BXTFU0(v, D, R, gamma, x, T)
            
            If (T > t0) Then
                CDE = CDE - c0 * AXTFU0(v, D, R, x, T - t0)
            End If
        Else
            CDE = gamma / mu - (ci - gamma / mu) * AXTFU1(v, D, R, mu, x, T) _
                             + (c0 - gamma / mu) * BXTFU1(v, D, R, mu, x, T)
            
            If (T > t0) Then
                CDE = CDE - c0 * BXTFU1(v, D, R, mu, x, T - t0)
            End If
        End If
    
    Case Else
        'volume average concentration, using Eqs in 3.1 and 3.2 in
        'Parker and van Genuchten 1984
        
        If (Abs(mu) < 0.0000000001) Then
            CDE = ci + (c0 - ci) * AXTRU0(v, D, R, x, T) _
                     + BXTRU0(v, D, R, gamma, x, T)
            
            If (T > t0) Then
                CDE = CDE - c0 * AXTRU0(v, D, R, x, T - t0)
            End If
        Else
            CDE = gamma / mu + (ci - gamma / mu) * AXTRU1(v, D, R, mu, x, T) _
                             + (c0 - gamma / mu) * BXTRU1(v, D, R, mu, x, T)
            
            If (T > t0) Then
                CDE = CDE - c0 * BXTRU1(v, D, R, mu, x, T - t0)
            End If
        End If
    End Select

End Function


'flux average concentration, using Eqs in 3.4 in Parker and van Genuchten 1984 mu = 0
Private Function AXTFU0(ByVal v As Double, _
                        ByVal D As Double, _
                        ByVal R As Double, _
                        ByVal x As Double, _
                        ByVal T As Double) As Double
If (T < 0.0000000001) Then
    AXTFU0 = 0#
Else
    AXTFU0 = 0.5 * VBAERFC((R * x - v * T) / 2# / Sqr(D * R * T)) _
           + 0.5 * EXF(v * x / D, (R * x + v * T) / 2# / Sqr(D * R * T))
End If
End Function

'flux average concentration, using Eqs in 3.4 in Parker and van Genuchten 1984 mu = 0
Private Function BXTFU0(ByVal v As Double, _
                        ByVal D As Double, _
                        ByVal R As Double, _
                        ByVal gamma As Double, _
                        x As Double, _
                        T As Double) As Double
If (T < 0.0000000001) Then
    BXTFU0 = 0#
Else
    BXTFU0 = gamma / R * (T _
                            + (R * x - v * T) / 2# / v * VBAERFC((R * x - v * T) / 2# / Sqr(D * R * T)) _
                            - (R * x + v * T) / 2# / v * EXF(v * x / D, (R * x + v * T) / 2# / Sqr(D * R * T)))
End If
End Function

'flux average concentration, using Eqs in 3.3 in Parker and van Genuchten 1984 mu .NE. 0
Private Function AXTFU1(ByVal v As Double, _
                        ByVal D As Double, _
                        ByVal R As Double, _
                        ByVal mu As Double, _
                        ByVal x As Double, _
                        ByVal T As Double) _
                        As Double
If (T < 0.0000000001) Then
    AXTFU1 = 0#
Else
    AXTFU1 = Exp(-mu * T / R) * (1# - 0.5 * VBAERFC((R * x - v * T) / 2# / Sqr(D * R * T)) _
           - 0.5 * EXF(v * x / D, (R * x + v * T) / 2# / Sqr(D * R * T)))
End If
End Function

'flux average concentration, using Eqs in 3.3 in Parker and van Genuchten 1984 mu .NE. 0
Private Function BXTFU1(ByVal v As Double, _
                        ByVal D As Double, _
                        ByVal R As Double, _
                        ByVal mu As Double, _
                        ByVal x As Double, _
                        ByVal T As Double) As Double
If (T < 0.0000000001) Then
    BXTFU1 = 0#
Else
    BXTFU1 = 0.5 * EXF((v - Sqr(v * v + 4# * mu * D)) * x / 2# / D, (R * x - Sqr(v * v + 4# * mu * D) * T) / 2# / Sqr(D * R * T)) _
           + 0.5 * EXF((v + Sqr(v * v + 4# * mu * D)) * x / 2# / D, (R * x + Sqr(v * v + 4# * mu * D) * T) / 2# / Sqr(D * R * T))
End If
End Function

'volume average concentration, using Eqs in 3.2 in Parker and van Genuchten 1984 mu = 0
Private Function AXTRU0(ByVal v As Double, _
                        ByVal D As Double, _
                        ByVal R As Double, _
                        ByVal x As Double, _
                        ByVal T As Double) As Double
If (T < 0.0000000001) Then
    AXTRU0 = 0#
Else
    AXTRU0 = 0.5 * VBAERFC((R * x - v * T) / 2# / Sqr(D * R * T)) _
           + Sqr(v * v * T / PI / D / R) * Exp(-(R * x - v * T) * (R * x - v * T) / 4# / D / R / T) _
           - 0.5 * (1# + v * x / D + v * v * T / D / R) * EXF(v * x / D, (R * x + v * T) / 2# / Sqr(D * R * T))
End If
End Function

'volume average concentration, using Eqs in 3.2 in Parker and van Genuchten 1984 mu = 0
Private Function BXTRU0(ByVal v As Double, _
                        ByVal D As Double, _
                        ByVal R As Double, _
                        ByVal gamma As Double, _
                        ByVal x As Double, _
                        ByVal T As Double) As Double
If (T < 0.0000000001) Then
    BXTRU0 = 0#
Else
    BXTRU0 = gamma / R * (T + (R * x - v * T + D * R / v) / 2# / v * VBAERFC((R * x - v * T) / 2# / Sqr(D * R * T)) _
           - Sqr(T / 4# / PI / D / R) * (R * x + v * T + 2# * D * R / v) * Exp(-(R * x - v * T) * (R * x - v * T) / 4# / D / R / T) _
           + (0.5 * T - 0.5 * D * R / v / v + (R * x + v * T) * (R * x + v * T) / 4# / D / R) * EXF(v * x / D, (R * x + v * T) / 2# / Sqr(D * R * T)))
End If
End Function

'volume average concentration, using Eqs in 3.1 in Parker and van Genuchten 1984 mu .NE. 0
Private Function AXTRU1(ByVal v As Double, _
                        ByVal D As Double, _
                        ByVal R As Double, _
                        ByVal mu As Double, _
                        ByVal x As Double, _
                        ByVal T As Double) As Double
If (T < 0.0000000001) Then
    AXTRU1 = 0#
Else
    AXTRU1 = Exp(-mu * T / R) * (1# - 0.5 * VBAERFC((R * x - v * T) / 2# / Sqr(D * R * T)) _
                                    - Sqr(v * v * T / PI / D / R) * Exp(-(R * x - v * T) * (R * x - v * T) / 4# / D / R / T) _
                                    + 0.5 * (1# + v * x / D + v * v * T / D / R) * EXF(v * x / D, (R * x + v * T) / 2# / Sqr(D * R * T)))
End If
End Function

'volume average concentration, using Eqs in 3.1 in Parker and van Genuchten 1984 mu .NE. 0
Private Function BXTRU1(ByVal v As Double, _
                        ByVal D As Double, _
                        ByVal R As Double, _
                        ByVal mu As Double, _
                        ByVal x As Double, _
                        ByVal T As Double) As Double
Dim U As Double
If (T < 0.0000000001) Then
    BXTRU1 = 0#
Else
    U = Sqr(v * v + 4# * mu * D)
    BXTRU1 = v / (v + U) * EXF((v - U) * x / 2# / D, (R * x - U * T) / 2# / Sqr(D * R * T)) _
           + v / (v - U) * EXF((v + U) * x / 2# / D, (R * x + U * T) / 2# / Sqr(D * R * T)) _
           + v * v / 2# / mu / D * EXF(v * x / D - mu * T / R, (R * x + v * T) / 2# / Sqr(D * R * T))
End If
End Function


'-----------------------------------------------------------------------------------------------
' function MIM
'-----------------------------------------------------------------------------------------------

'Solving non-equilibrium convection-diffusion equation with first order decay
' beta R dC1/dT = 1/P d2C1/dZ2 - dC1/dZ - omega(C1 - C2) - mu1 C1
' (1 - beta) R dC2/dT = omega(C1 - C2) - mu2 C2
' initial condition C1(Z, 0) = C2(Z, 0) = 0
' inlet boundary condition C1(0, T) - 1/P dC1(0,T)/dZ = C0(t)
' outlet boundary condition dC1(infinity, T)/dZ = 0
' Eq. 3.5 and 3.6 in Toride et al. 1995 using solution Eq. 3.20-22

Public Function MIM(ByVal Z As Double, _
                    ByVal T As Double, _
                    ByVal t0 As Double, _
                    ByVal p As Double, _
                    ByVal beta As Double, _
                    ByVal omega As Double, _
                    Optional ByVal R As Variant, _
                    Optional ByVal mu1 As Variant, _
                    Optional ByVal mu2 As Variant, _
                    Optional ByVal strFluxConc As Variant, _
                    Optional ByVal strMobile As Variant, _
                    Optional ByVal ci As Variant, _
                    Optional ByVal c0 As Variant, _
                    Optional ByVal bln As Variant, _
                    Optional ByVal strNI As Variant, _
                    Optional ByVal tol As Variant, _
                    Optional ByVal nPoint As Variant) As Double


If (IsMissing(R)) Then
    R = 1#
End If

If (IsMissing(mu1)) Then
    mu1 = 0#
End If

If (IsMissing(mu2)) Then
    mu2 = 0#
End If

If (IsMissing(tol)) Then
    tol = 0.0000001      'integration tolerance
End If
    
If (IsMissing(strNI)) Then
    strNI = "chebyshev"      'integration tolerance
End If
    
If (IsMissing(bln)) Then
    bln = False      'integration tolerance
End If
    
If (IsMissing(nPoint)) Then
    nPoint = 75      'integration tolerance
End If
    
If (IsMissing(strFluxConc)) Then
    strFluxConc = "flux"   ' any input other than Flux is treated as residence concentration
End If

If (IsMissing(strMobile)) Then
    strMobile = "mobile"   ' any input other than mobile is treated as  concentration in immobile phase
End If


MIM = A1zt(Z, T, p, R, beta, omega, mu1, mu2, tol, strFluxConc, strMobile, strNI, bln, nPoint)

If (T > t0) Then
    MIM = MIM - A1zt(Z, T - t0, p, R, beta, omega, mu1, mu2, tol, strFluxConc, strMobile, strNI, bln, nPoint)
End If

If IsMissing(c0) Then
    c0 = 1#
End If
MIM = MIM * c0
        
If IsMissing(ci) Then
    ci = 0#
End If
        
If (Abs(ci) > 0.0000000001) Then
    MIM = MIM + IVP(ci, T, Z, 0#, p, beta, _
                    omega, R, mu1, mu2, strFluxConc, tol, strNI, bln, nPoint)
End If

End Function

Private Function A1zt(ByVal Z As Double, _
                      ByVal T As Double, _
                      ByVal p As Double, _
                      ByVal R As Double, _
                      ByVal beta As Double, _
                      ByVal omega As Double, _
                      ByVal mu1 As Double, _
                      ByVal mu2 As Double, _
                      ByVal tol As Double, _
                      Optional ByVal strFluxConc As Variant, _
                      Optional ByVal strMobile As Variant, _
                      Optional ByVal strNI As Variant, _
                      Optional ByVal bln As Variant, _
                      Optional ByVal nPoint As Variant) As Double

'Eq 3.21 in Toride et al. 1995

Dim a, b, T1, T2 As Double
'cxt 1.0 Parker and van Genuchten 1984
'    a = beta * R * Z + 40# * beta * R / P * (1# - Sqr((1# + P * Z / 20#)))
'    b = beta * R * Z + 40# * beta * R / P * (1# + Sqr((1# + P * Z / 20#)))

'cxt 2.0 Toride et al. 1995
    a = beta * R * Z + 60# * beta * R / p * (1# - Sqr((1# + p * Z / 30#)))
    b = beta * R * Z + 60# * beta * R / p * (1# + Sqr((1# + p * Z / 30#)))
    
    T1 = Max(0#, a)
    T2 = Min(T, b)
    
'    If (IsMissing(strNI)) Then
'        strNI = "Chebyshev"
'    End If
    
'    If (IsMissing(bln)) Then
'        bln = False
'    End If
    
    Select Case LCase(strNI)
        Case "romberg"
            A1zt = ROMB("gztJab", T1, T2, tol, bln, Z, T, p, R, beta, omega, mu1, mu2, strFluxConc, strMobile)
        
        Case "simpson"
            A1zt = AdaptSim("gztJab", T1, T2, tol, bln, Z, T, p, R, beta, omega, mu1, mu2, strFluxConc, strMobile)
        
        Case "lobatto"
            A1zt = AdaptLob("gztJab", T1, T2, tol, bln, Z, T, p, R, beta, omega, mu1, mu2, strFluxConc, strMobile)
        
        Case "chebyshev"
'            If (IsMissing(nPoint)) Then
'                nPoint = 50
'            End If
            
'            If (nPoint < 1) Then
'                nPoint = 1
'            End If
            A1zt = ChebyQ("gztJab", T1, T2, nPoint, bln, Z, T, p, R, beta, omega, mu1, mu2, strFluxConc, strMobile)
    End Select
End Function


Private Function gztJab(ByVal T As Double, _
                ByVal Z As Double, _
                ByVal t0 As Double, _
                ByVal p As Double, _
                ByVal R As Double, _
                ByVal beta As Double, _
                ByVal omega As Double, _
                ByVal mu1 As Double, _
                ByVal mu2 As Double, _
                Optional ByVal strFluxConc As Variant, _
                Optional ByVal strMobile As Variant) As Double

'integrand in Eq. 3.21
    Dim GOLD, decay, gzt As Double
    Dim a, b As Double

    Select Case LCase(strFluxConc)
        Case "flux"
           gzt = gztF(Z, T, p, R, beta)
        Case Else
           gzt = gztR(Z, T, p, R, beta)
    End Select
    
'    GOLD = Goldstein(omega * omega * T / beta / R / (omega + mu2), (omega + mu2) * (t0 - T) / (1# - beta) / R)
    decay = Exp(-mu1 * T / beta / R) * Exp(-omega * mu2 * T / (omega + mu2) / beta / R)

    a = omega * omega * T / beta / R / (omega + mu2)
    b = (omega + mu2) * (t0 - T) / (1# - beta) / R
    
'        J(a,b) in Table 3.4 Toride et al. 1995
    Select Case LCase(strMobile)
        Case "mobile"
            GOLD = Goldstein(a, b)
            gztJab = decay * gzt * GOLD
        Case Else
            GOLD = Goldstein(b, a)
            gztJab = omega / (omega + mu2) * decay * gzt * (1 - GOLD)
    End Select
    
'    gztJab = decay * gzt * GOLD

End Function

Private Function gztR(ByVal Z As Double, _
              ByVal T As Double, _
              ByVal p As Double, _
              ByVal R As Double, _
              ByVal beta As Double) As Double
'Table 3.2 Third-type, Cr, Toride et al. 1995 for residence concentration
              
    gztR = Sqr(p / PI / beta / R / T) * Exp(-p * (beta * R * Z - T) ^ 2 / 4# / beta / R / T) _
         - p / 2# / beta / R * EXF(p * Z, Sqr(p / 4# / beta / R / T) * (beta * R * Z + T))

End Function
    
Private Function gztF(ByVal Z As Double, _
              ByVal T As Double, _
              ByVal p As Double, _
              ByVal R As Double, _
              ByVal beta As Double) As Double
'Table 3.2 Cf, Toride et al. 1995 for flux average concentration
    
    gztF = Z / T * Sqr(p * beta * R / 4# / PI / T) * Exp(-p * (beta * R * Z - T) * (beta * R * Z - T) / 4# / beta / R / T)

End Function

Private Function Goldstein(ByVal x As Double, ByVal y As Double) As Double
'     PURPOSE: TO CALCULATE GOLDSTEIN'S J-FUNCTION J(a,b)
'    in Eq. 3.21 in Toride et al. 1995,
'    copied from CXTFIT 1.0 Parker and van Genuchten 1984
    
    Dim a, b, B0, c0, C1, C2, C3, C4, E, p, Z As Double
    Dim BF, GOLD, NT, DA, Db, ERF, Sum, GXY As Double
    Dim GXYO, GX, GY, GZ As Double
    Dim I, k As Integer
    
    
    GOLD = 0#
    BF = 0#
    E = 2# * Sqr(Max(0#, x * y))
    Z = x + y - E
    
    If (Z > 17#) Then
        If (x < y) Then
            GOLD = 1# + BF - GOLD
        End If
        Goldstein = GOLD
        Exit Function
    End If
        
    If (Abs(E) < 0.0000000001) Then
        Goldstein = Exp(-x)
        Exit Function
    End If
  
    a = Max(x, y)
    b = Min(x, y)
    NT = 11# + 2# * b + 0.3 * a
    If (NT > 25#) Then
        DA = Sqr(a)
        Db = Sqr(b)
        p = 3.75 / E
        B0 = (0.3989423 + p * (0.01328592 + p * (0.00225319 - p * (0.00157565 - p * (0.00916281 - p * (0.02057706 - p * (0.02635537 - p * (0.01647633 - 0.00392377 * p)))))))) / Sqr(E)
        BF = B0 * EXF(-Z, 0#)
        p = 1# / (1# + 0.3275911 * (DA - Db))
        ERF = p * (0.2548296 - p * (0.2844967 - p * (1.421414 - p * (1.453152 - p * 1.061405))))
        p = 0.25 / E
        c0 = 1# - 1.772454 * (DA - Db) * ERF
        C1 = 0.5 - Z * c0
        C2 = 0.75 - Z * C1
        C3 = 1.875 - Z * C2
        C4 = 6.5625 - Z * C3
        Sum = 0.1994711 * (a - b) * p * (c0 + 1.5 * p * (C1 + 1.666667 * p * (C2 + 1.75 * p * (C3 + p * (C4 * (1.8 - 3.3 * p * Z) + 97.45313 * p)))))
        GOLD = 0.5 * BF + (0.3535534 * (DA + Db) * ERF + Sum) * BF / (B0 * Sqr(E))
        If (x < y) Then
            GOLD = 1# + BF - GOLD
        End If
        Goldstein = GOLD
        Exit Function
    End If

    I = 0
    If (x < y) Then
        I = 1
    End If
    
    GXY = 1# + I * (b - 1#)
    GXYO = GXY
    GX = 1#
    GY = GXY
    GZ = 1#
    
    For k = 1 To NT
      GX = GX * a / k
      GY = GY * b / (k + I)
      GZ = GZ + GX
      GXY = GXY + GY * GZ
      If ((GXY - GXYO) / GXY < 0.00000001) Then
            GOLD = GXY * EXF(-x - y, 0#)
            If (x < y) Then
                GOLD = 1# + BF - GOLD
            End If
            Goldstein = GOLD
            Exit Function
      End If
      
      GXYO = GXY
    Next

    GOLD = GXY * EXF(-x - y, 0#)
    If (x < y) Then
        GOLD = 1# + BF - GOLD
    End If
    
    Goldstein = GOLD
End Function


'stepwise initial concentration
'Eq. (2.27) in Toride et al. 1995
Private Function PsiE1(ByVal Z As Double, _
                       ByVal Zi As Double, _
                       ByVal T As Double, _
                       ByVal p As Double, _
                       Optional ByVal R As Variant, _
                       Optional ByVal mu As Variant, _
                       Optional ByVal strFluxConc As Variant) As Double

If IsMissing(R) Then
    R = 1#
End If

If IsMissing(mu) Then
    mu = 0#
End If

If IsMissing(strFluxConc) Then
    strFluxConc = "flux"
End If

Select Case LCase(strFluxConc)
    Case "flux"
        PsiE1 = 0.5 * Exp(-mu * T / R) * (2# - _
                VBAERFC((R * (Z - Zi) - T) / Sqr(4# * R * T / p)) - _
                EXF(p * Z, (R * (Z + Zi) + T) / Sqr(4# * R * T / p)) - _
                Sqr(R / PI / p / T) * Exp(-p * (R * (Z - Zi) - T) * (R * (Z - Zi) - T) / 4# / R / T) + _
                Sqr(R / PI / p / T) * EXF(p * Z, -p * (R * (Z + Zi) + T) * (R * (Z + Zi) + T) / 4# / R / T))
    Case Else
'        PsiE1 = Exp(-mu * T / R) * (1# - _
'                0.5 * VBAERFC((R * (Z - Zi) - T) / Sqr(4# * R * T / P)) - _
'                0.5 * EXF(P * Z, (R * (Z + Zi) + T) / Sqr(4# * R * T / P)))
        PsiE1 = Exp(-mu * T / R) * (1# - _
                0.5 * VBAERFC((R * (Z - Zi) - T) / Sqr(4# * R * T / p)) - _
                Sqr(p * T / PI / R) * Exp(p * Z - p * (R * (Z + Zi) + T) * (R * (Z + Zi) + T) / 4# / R / T) + _
                0.5 * (1# + p * (Z + Zi) + p * T / R) * EXF(p * Z, (R * (Z + Zi) + T) / Sqr(4# * R * T / p)))
    End Select
'End If    Andrew Moore 01/07/2010

End Function

'boundary value problem for MIM

Private Function Gamma1N(ByVal Z As Double, _
                         ByVal tau As Double, _
                         ByVal p As Double, _
                         ByVal beta As Double, _
                         ByVal omega As Double, _
                         Optional ByVal R As Variant, _
                         Optional ByVal mu As Variant, _
                         Optional ByVal strFluxConc As Variant) As Double

If IsMissing(R) Then
    R = 1#
End If

If IsMissing(mu) Then
    mu = 0#
End If

If IsMissing(strFluxConc) Then
    strFluxConc = "flux"
End If

Select Case LCase(strFluxConc)
    Case "flux"
'If (bFluxConc) Then
        Gamma1N = Exp(-mu * tau / beta / R) * Z / tau * Sqr(beta * R * p / 4# / PI / tau) * _
                  Exp(-p * (beta * R * Z - tau) * (beta * R * Z - tau) / 4# / beta / R / tau)
    Case Else
'Else
        Gamma1N = Exp(-mu * tau / beta / R) * ( _
                Sqr(p / PI / beta / R / tau) * Exp(-p * (beta * R * Z - tau) * (beta * R * Z - tau) / 4# / beta / R / tau) - _
                p / 2# / beta / R * EXF(p * Z, (beta * R * Z + tau) / Sqr(4# * beta * R * tau / p)))
End Select
End If

End Function

Private Function EXPBI0(ByVal x As Double, ByVal Z As Double) As Double
Dim y As Double
y = x / 3.75
If (x >= -3.71 And x <= 3.75) Then
    EXPBI0 = Exp(Z) * (1# + 3.5156229 * y * y _
                         + 3.0899424 * y * y * y * y _
                         + 1.2067492 * y * y * y * y * y * y _
                         + 0.2659732 * y * y * y * y * y * y * y * y _
                         + 0.0360768 * y * y * y * y * y * y * y * y * y * y _
                         + 0.0045813 * y * y * y * y * y * y * y * y * y * y * y * y)
ElseIf (x >= 3.75) Then
    EXPBI0 = Exp(Z + x) / Sqr(x) * ( _
                         0.39894228 _
                         + 0.01328592 / y _
                         + 0.00225319 / y / y _
                         - 0.00157565 / y / y / y _
                         + 0.00916281 / y / y / y / y _
                         - 0.02057706 / y / y / y / y / y _
                         + 0.02635537 / y / y / y / y / y / y _
                         - 0.01647633 / y / y / y / y / y / y / y _
                         + 0.00392377 / y / y / y / y / y / y / y / y)
End If
End Function


Private Function EXPBI1(ByVal x As Double, ByVal Z As Double) As Double
Dim y As Double

y = x / 3.75
If (x >= -3.71 And x <= 3.75) Then
    EXPBI1 = x * Exp(Z) * (0.5 _
                         + 0.87890594 * y * y _
                         + 0.51498869 * y * y * y * y _
                         + 0.15084934 * y * y * y * y * y * y _
                         + 0.02658733 * y * y * y * y * y * y * y * y _
                         - 0.00301532 * y * y * y * y * y * y * y * y * y * y _
                         + 0.00032411 * y * y * y * y * y * y * y * y * y * y * y * y)
ElseIf (x >= 3.75) Then
    EXPBI1 = Exp(Z + x) / Sqr(x) * ( _
                           0.39894228 _
                         - 0.03988024 / y _
                         - 0.00362018 / y / y _
                         + 0.00163801 / y / y / y _
                         - 0.01031555 / y / y / y / y _
                         + 0.02282967 / y / y / y / y / y _
                         - 0.02895312 / y / y / y / y / y / y _
                         + 0.01787654 / y / y / y / y / y / y / y _
                         - 0.00420059 / y / y / y / y / y / y / y / y)
End If
End Function



'integral in 3.15 in Toride et al. 1995
Private Function Gamma1NH1(ByVal tau As Double, _
                           ByVal Z As Double, _
                           ByVal T As Double, _
                           ByVal p As Double, _
                           ByVal beta As Double, _
                           ByVal omega As Double, _
                           Optional ByVal R As Variant, _
                           Optional ByVal mu1 As Variant, _
                           Optional ByVal mu2 As Variant, _
                           Optional ByVal strFluxConc As Variant) As Double
Gamma1NH1 = Sqr(tau / beta / (1# - beta) / (T - tau))
Gamma1NH1 = Gamma1NH1 * Gamma1N(Z, tau, p, beta, omega, R, mu1, strFluxConc)
Gamma1NH1 = Gamma1NH1 * EXPBI1(-omega * tau / beta / R - (omega + mu2) * (T - tau) / (1# - beta) / R, _
                                2# * omega / R * Sqr((T - tau) * tau / beta / (1# - beta)))
End Function

'Eq. 3.15 in Toride et al 1995
Private Function FZT(ByVal Z As Double, _
                     ByVal T As Double, _
                     ByVal p As Double, _
                     ByVal beta As Double, _
                     ByVal omega As Double, _
                     Optional ByVal R As Variant, _
                     Optional ByVal mu1 As Variant, _
                     Optional ByVal mu2 As Variant, _
                     Optional ByVal strFluxConc As Variant) As Double

    If IsMissing(R) Then
        R = 1#
    End If

    If IsMissing(mu1) Then
        mu1 = 0#
    End If

    If IsMissing(mu2) Then
        mu2 = 0#
    End If

    If IsMissing(strFluxConc) Then
        strFluxConc = "flux"
    End If
    
    Dim T1, T2 As Double
    T1 = beta * R * Z + 60# * beta * R / p * (1 - Sqr(1 + p * Z / 30))
    T2 = beta * R * Z + 60# * beta * R / p * (1 + Sqr(1 + p * Z / 30))
    
    Dim tmp As Double
    If (T < T1) Then
        tmp = 0#
    Else
        If (T2 > T) Then
            T2 = 0.999999999 * T
        End If
            
        tmp = ROMBLN("Gamma1NH1", T1, T2, 0.00000001, Z, T, p, beta, omega, p, mu1, mu2, strFluxConc)
    End If
    
    FZT = Gamma1N(Z, T, p, beta, omega, R, mu1, strFluxConc) * Exp(-omega * T / beta / R)
    FZT = FZT + tmp * omega / R

End Function


Public Sub ErrorAnalysis1(ByVal SSR As Double, _
                         ByVal ParameterRange As Range, _
                         ByVal PredictionRange As Range, _
                         ByVal OffsetParaStdOut As Integer, _
                         Optional ByVal OffsetPredS2P As Variant, _
                         Optional ByVal OffsetPredJcb As Variant, _
                         Optional ByVal OffsetParaDlt As Variant, _
                         Optional ByVal PenaltyWeight As Variant, _
                         Optional ByVal OffsetParaPriorStd As Variant, _
                         Optional ByVal OffsetPredObsStd As Variant)
                  
    Dim I, J As Integer
    Dim nPara, nData As Integer
    Dim ParaOldValue As Double
    Dim CorrelationRange As Range
    Dim SensitivityRange As Range
    Dim dtmp As Double
    Dim Delta As Double
    Dim MSE As Double
    
    nPara = ParameterRange.Count    'number of parameters to be estimated
    nData = PredictionRange.Count   'number of observations
    
    If (nPara > nData) Then
        MsgBox ("The number of observations is less than the number of estimated parameters!!!")
        Return
    End If
    
    ReDim SensitivityMatrix(1 To nData + nPara, 1 To nPara) As Double 'derivative
    ReDim PredMeanValue(1 To nData) As Variant      'model predition for average P
    ReDim PredPerturbationValue(1 To nData) As Variant      'model prediction for average P + 0.01P
    ReDim CovarianceMatrix(1 To nPara, 1 To nPara) As Double
    ReDim CorrelationMatrix(1 To nPara, 1 To nPara) As Double
    ReDim StandardDeviation(1 To nPara, 1 To 1) As Double
    ReDim St(1 To nPara, 1 To nData) As Double
    ReDim Sw(1 To nPara, 1 To nData) As Double
    ReDim Jt(1 To nPara, 1 To nData + nPara) As Double
    ReDim aa(1 To nPara, 1 To nPara) As Double
    ReDim Bb(1 To nPara, 1 To nData + nPara) As Double
    ReDim iA(1 To nPara, 1 To nPara) As Double
    ReDim sPred(1 To nData, 1 To 1) As Double
    ReDim s2Para(1 To nData, 1 To 1) As Double
    ReDim TA(1 To 1, 1 To nPara) As Double
    ReDim tAt(1 To nPara, 1 To 1) As Double
    ReDim tM1(1 To 1, 1 To nPara) As Double
    ReDim tM2(1 To 1, 1 To 1) As Double
    ReDim WeightMatrixC(1 To nData, 1 To nData) As Double
    '----------------------------------------------------------------
    ' compute sensitivity matrix related to observations
    '----------------------------------------------------------------
    PredMeanValue = PredictionRange.Value
    
    For I = 1 To nPara
'       parameter value
        ParaOldValue = ParameterRange.Cells(I, 1).Value
        
        If (Abs(ParaOldValue) < 0.0000000001) Then
            Delta = 0.01    'to avoid zero perturbation for parameters of zero value
        Else
            'perturbation is not specified
            If (IsMissing(OffsetParaDlt)) Then
                Delta = ParaOldValue * 0.01
            ElseIf (OffsetParaDlt = 0) Then
                Delta = ParaOldValue * 0.01
            Else
                Delta = ParaOldValue * ParameterRange.Offset(0, OffsetParaDlt).Cells(I, 1).Value
            End If
        End If
        
        ParameterRange.Cells(I, 1).Value = ParaOldValue + Delta
 
'    derivative
        PredPerturbationValue = PredictionRange.Value
        
        For J = 1 To nData
            SensitivityMatrix(J, I) = (PredPerturbationValue(J, 1) - PredMeanValue(J, 1)) / Delta
        Next J
        
        ParameterRange.Cells(I, 1).Value = ParaOldValue
    Next I
    
    If (Not (IsMissing(OffsetPredJcb))) Then
        If (Not (OffsetPredJcb = 0)) Then
    '----------------------------------------------------------------
    ' output sensitivity matrix related to measurements
    '----------------------------------------------------------------
            Set SensitivityRange = PredictionRange.Offset(0, OffsetPredJcb)
            For I = 2 To nPara
                Set SensitivityRange = Union(SensitivityRange, PredictionRange.Offset(0, I + OffsetPredJcb - 1))
            Next I
            SensitivityRange.Value = SensitivityMatrix
        End If
    End If
    
    '----------------------------------------------------------------
    ' compute covariance matrix
    '----------------------------------------------------------------
    'weight for observation
    For I = 1 To nData
        If (IsMissing(OffsetPredObsStd)) Then
            WeightMatrixC(I, I) = 1#
        ElseIf (OffsetPredObsStd = 0) Then
            WeightMatrixC(I, I) = 1#
        Else
            dtmp = PredictionRange.Offset(0, OffsetPredObsStd).Cells(I, 1)
            WeightMatrixC(I, I) = 1# / dtmp / dtmp
        End If
    Next I
    
    Call MatrixTranspose(SensitivityMatrix, nData, nPara, St)
    Call MatrixMultiplication(St, nPara, nData, WeightMatrixC, nData, Sw)
    Call MatrixMultiplication(Sw, nPara, nData, SensitivityMatrix, nPara, aa)
        
    If (Not (IsMissing(OffsetParaPriorStd))) Then
         If (Not (OffsetParaPriorStd = 0)) Then
        'weight for parameters in penalty function
            If (IsMissing(PenaltyWeight) Or (Abs(PenaltyWeight) < 0.00000001)) Then
                PenaltyWeight = 1#
            End If
        
            For I = 1 To nPara
                dtmp = ParameterRange.Offset(0, OffsetParaPriorStd).Cells(I, 1).Value
                aa(I, I) = aa(I, I) + PenaltyWeight / dtmp / dtmp
            Next I
            
            MSE = 1#
        Else
            MSE = SSR / (nData - nPara)
        End If
    Else
        MSE = SSR / (nData - nPara)
    End If
    
    Call MatrixInverse(aa, nPara, iA)
    
    'calculate the covariance matrix
    For I = 1 To nPara
        For J = 1 To nPara
            CovarianceMatrix(I, J) = MSE * iA(I, J)
        Next J
    Next I
    
        'calculate the standard deviation
    For I = 1 To nPara
        If (CovarianceMatrix(I, I) < 0) Then
             MsgBox ("Standard deviation of estimated parameter is too close to zero!")
             StandardDeviation(I, 1) = -99
        Else
             StandardDeviation(I, 1) = Sqr(CovarianceMatrix(I, I))
        End If
    Next I
        
    'write the standard deviation for estimated parameters
    ParameterRange.Offset(0, OffsetParaStdOut).Value = StandardDeviation
    
    'calculate the correlation matrix
    For I = 1 To nPara
        For J = 1 To nPara
            If (Abs(StandardDeviation(I, 1)) < 0.000000001 Or Abs(StandardDeviation(J, 1)) < 0.000000001) Then
                MsgBox ("Standard deviation for estimated parameters is close to zero, correlation is set to-999")
                CorrelationMatrix(I, J) = -999
            Else
                CorrelationMatrix(I, J) = CovarianceMatrix(I, J) / StandardDeviation(I, 1) / StandardDeviation(J, 1)
            End If
        Next J
    Next I
    
    Set CorrelationRange = ParameterRange.Offset(0, OffsetParaStdOut + 1)
   
    For I = 2 To nPara
        Set CorrelationRange = Union(CorrelationRange, ParameterRange.Offset(0, I + OffsetParaStdOut))
    Next I

    CorrelationRange.Value = CorrelationMatrix
    CorrelationRange.NumberFormat = "0.000"
    
   If (Not (IsMissing(OffsetPredS2P))) Then
       If (Not (OffsetPredS2P = 0)) Then
    'calculate uncertainty due to estimated parameter uncertainty
            For I = 1 To nData
                For J = 1 To nPara
                    TA(1, J) = SensitivityMatrix(I, J)
                    tAt(J, 1) = TA(1, J)
                Next J
            
                Call MatrixMultiplication(TA, 1, nPara, CovarianceMatrix, nPara, tM1)
                Call MatrixMultiplication(tM1, 1, nPara, tAt, 1, tM2)
             
                s2Para(I, 1) = tM2(1, 1)
            Next I
    
            PredictionRange.Offset(0, OffsetPredS2P).Value = s2Para
       End If
   End If
End Sub

Public Sub ErrorAnalysis(ByVal SSR As Double, _
                         ByVal ParameterRange As Range, _
                         ByVal PredictionRange As Range, _
                         ByVal OffsetParaStdOut As Integer, _
                         Optional ByVal OffsetPredS2P As Variant, _
                         Optional ByVal OffsetPredJcb As Variant, _
                         Optional ByVal OffsetParaDlt As Variant, _
                         Optional ByVal PenaltyWeight As Variant, _
                         Optional ByVal OffsetParaPriorStd As Variant, _
                         Optional ByVal OffsetPredObsStd As Variant)
                  
    Dim I, J As Integer
    Dim nPara, nData As Integer
    Dim ParaOldValue As Double
    Dim CorrelationRange As Range
    Dim SensitivityRange As Range
    Dim dtmp As Double
    Dim Delta As Double
    Dim MSE As Double
    Dim UnitStdDev As Boolean       'check if the standard deviation for observations are set to be 1.0
    
    nPara = ParameterRange.Count    'number of parameters to be estimated
    nData = PredictionRange.Count   'number of observations
    
    If (nPara > nData) Then
        MsgBox ("The number of observations is less than the number of estimated parameters!!!")
        Return
    End If
    
    ReDim SensitivityMatrix(1 To nData + nPara, 1 To nPara) As Double 'derivative
    ReDim PredMeanValue(1 To nData) As Variant      'model predition for average P
    ReDim PredPerturbationValue(1 To nData) As Variant      'model prediction for average P + 0.01P
    ReDim CovarianceMatrix(1 To nPara, 1 To nPara) As Double
    ReDim CorrelationMatrix(1 To nPara, 1 To nPara) As Double
    ReDim StandardDeviation(1 To nPara, 1 To 1) As Double
    ReDim St(1 To nPara, 1 To nData) As Double
    ReDim Sw(1 To nPara, 1 To nData) As Double
    ReDim Jt(1 To nPara, 1 To nData + nPara) As Double
    ReDim aa(1 To nPara, 1 To nPara) As Double
    ReDim Bb(1 To nPara, 1 To nData + nPara) As Double
    ReDim iA(1 To nPara, 1 To nPara) As Double
    ReDim sPred(1 To nData, 1 To 1) As Double
    ReDim s2Para(1 To nData, 1 To 1) As Double
    ReDim TA(1 To 1, 1 To nPara) As Double
    ReDim tAt(1 To nPara, 1 To 1) As Double
    ReDim tM1(1 To 1, 1 To nPara) As Double
    ReDim tM2(1 To 1, 1 To 1) As Double
    ReDim WeightMatrixC(1 To nData, 1 To nData) As Double
    '----------------------------------------------------------------
    ' compute sensitivity matrix related to observations
    '----------------------------------------------------------------
    PredMeanValue = PredictionRange.Value
    
    For I = 1 To nPara
'       parameter value
        ParaOldValue = ParameterRange.Cells(I, 1).Value
        
        If (Abs(ParaOldValue) < 0.0000000001) Then
            Delta = 0.01    'to avoid zero perturbation for parameters of zero value
        Else
            'perturbation is not specified
            If (IsMissing(OffsetParaDlt)) Then
                Delta = ParaOldValue * 0.01
            ElseIf (OffsetParaDlt = 0) Then
                Delta = ParaOldValue * 0.01
            Else
                Delta = ParaOldValue * ParameterRange.Offset(0, OffsetParaDlt).Cells(I, 1).Value
            End If
        End If
        
        ParameterRange.Cells(I, 1).Value = ParaOldValue + Delta
 
'    derivative
        PredPerturbationValue = PredictionRange.Value
        
        For J = 1 To nData
            SensitivityMatrix(J, I) = (PredPerturbationValue(J, 1) - PredMeanValue(J, 1)) / Delta
        Next J
        
        ParameterRange.Cells(I, 1).Value = ParaOldValue
    Next I
    
    If (Not (IsMissing(OffsetPredJcb))) Then
        If (Not (OffsetPredJcb = 0)) Then
    '----------------------------------------------------------------
    ' output sensitivity matrix related to measurements
    '----------------------------------------------------------------
            Set SensitivityRange = PredictionRange.Offset(0, OffsetPredJcb)
            For I = 2 To nPara
                Set SensitivityRange = Union(SensitivityRange, PredictionRange.Offset(0, I + OffsetPredJcb - 1))
            Next I
            SensitivityRange.Value = SensitivityMatrix
        End If
    End If
    
    '----------------------------------------------------------------
    ' compute covariance matrix
    '----------------------------------------------------------------
    'weight for observation
    UnitStdDev = True
    If (IsMissing(OffsetPredObsStd) Or OffsetPredObsStd = 0) Then
        For I = 1 To nData
            WeightMatrixC(I, I) = 1#
        Next I
    Else
        For I = 1 To nData
            dtmp = PredictionRange.Offset(0, OffsetPredObsStd).Cells(I, 1)
            WeightMatrixC(I, I) = 1# / dtmp / dtmp
           
            If (Abs(dtmp - 1#) > 0.0000000001) Then
                UnitStdDev = False   'if any of the standard deviation is not equal to 1, it is not unit.
            End If
        Next I
    End If
    
    If (UnitStdDev) Then
        MSE = SSR / (nData - nPara)
    Else
        MSE = 1#
    End If
    
    Call MatrixTranspose(SensitivityMatrix, nData, nPara, St)
    Call MatrixMultiplication(St, nPara, nData, WeightMatrixC, nData, Sw)
    Call MatrixMultiplication(Sw, nPara, nData, SensitivityMatrix, nPara, aa)
        
    If (Not (IsMissing(OffsetParaPriorStd))) Then
         If (Not (OffsetParaPriorStd = 0) And Not (IsMissing(PenaltyWeight) And (Abs(PenaltyWeight) > 0.00000001))) Then
        'weight for parameters in penalty function
            For I = 1 To nPara
                dtmp = ParameterRange.Offset(0, OffsetParaPriorStd).Cells(I, 1).Value
                aa(I, I) = aa(I, I) + PenaltyWeight / dtmp / dtmp
            Next I
        End If
    End If
    
    Call MatrixInverse(aa, nPara, iA)
    
    'calculate the covariance matrix
    For I = 1 To nPara
        For J = 1 To nPara
            CovarianceMatrix(I, J) = MSE * iA(I, J)
        Next J
    Next I
    
        'calculate the standard deviation
    For I = 1 To nPara
        If (CovarianceMatrix(I, I) < 0) Then
             MsgBox ("Standard deviation of estimated parameter is too close to zero!")
             StandardDeviation(I, 1) = -99
        Else
             StandardDeviation(I, 1) = Sqr(CovarianceMatrix(I, I))
        End If
    Next I
        
    'write the standard deviation for estimated parameters
    ParameterRange.Offset(0, OffsetParaStdOut).Value = StandardDeviation
    
    'calculate the correlation matrix
    For I = 1 To nPara
        For J = 1 To nPara
            If (Abs(StandardDeviation(I, 1)) < 0.000000001 Or Abs(StandardDeviation(J, 1)) < 0.000000001) Then
                MsgBox ("Standard deviation for estimated parameters is close to zero, correlation is set to-999")
                CorrelationMatrix(I, J) = -999
            Else
                CorrelationMatrix(I, J) = CovarianceMatrix(I, J) / StandardDeviation(I, 1) / StandardDeviation(J, 1)
            End If
        Next J
    Next I
    
    Set CorrelationRange = ParameterRange.Offset(0, OffsetParaStdOut + 1)
   
    For I = 2 To nPara
        Set CorrelationRange = Union(CorrelationRange, ParameterRange.Offset(0, I + OffsetParaStdOut))
    Next I

    CorrelationRange.Value = CorrelationMatrix
    CorrelationRange.NumberFormat = "0.000"
    
   If (Not (IsMissing(OffsetPredS2P))) Then
       If (Not (OffsetPredS2P = 0)) Then
    'calculate uncertainty due to estimated parameter uncertainty
            For I = 1 To nData
                For J = 1 To nPara
                    TA(1, J) = SensitivityMatrix(I, J)
                    tAt(J, 1) = TA(1, J)
                Next J
            
                Call MatrixMultiplication(TA, 1, nPara, CovarianceMatrix, nPara, tM1)
                Call MatrixMultiplication(tM1, 1, nPara, tAt, 1, tM2)
             
                s2Para(I, 1) = tM2(1, 1)
            Next I
    
            PredictionRange.Offset(0, OffsetPredS2P).Value = s2Para
       End If
   End If
End Sub

Public Function NameExists(ByVal FindName As String, Optional ByVal SheetName As Variant) As Boolean
Dim Rng As Range
Dim myName As String
On Error Resume Next

If IsMissing(SheetName) Then
    myName = ActiveWorkbook.Names(FindName).Name
Else
    myName = ActiveSheet.Names(FindName).Name
End If

If Err.Number = 0 Then NameExists = True
End Function

'matrix inverse
Public Sub MatrixInverse(ByRef Min() As Double, ByVal nrc As Integer, ByRef Mout() As Double)

' The square input and output matrices are Min and Mout
' respectively; nrc is the number of rows and columns in
' Min and Mout
    
Dim I As Integer, icol As Integer, irow As Integer
Dim J As Integer, k As Integer, L As Integer, LL As Integer
Dim big As Double, dummy As Double
Dim n As Integer, pivinv As Double
Dim U As Double

n = nrc + 1
ReDim Bb(1 To n, 1 To n) As Double
ReDim ipivot(1 To n) As Double
ReDim Index(1 To n) As Double
ReDim indexr(1 To n) As Double
ReDim indexc(1 To n) As Double
U = 1

' Copy the input matrix in order to retain it

For I = 1 To nrc
  For J = 1 To nrc
    Mout(I, J) = Min(I, J)      'Min rather than M1
  Next J
Next I

' The following is the Gauss-Jordan elimination routine
' GAUSSJ from J. C. Sprott, "Numerical Recipes: Routines
' and Examples in BASIC", Cambridge University Press,
' Copyright (C)1991 by Numerical Recipes Software. Used by
' permission. Use of this routine other than as an integral
' part of the present book requires an additional license
' from Numerical Recipes Software. Further distribution is
' prohibited. The routine has been modified to yield
' double-precision results.

For J = 1 To nrc
  ipivot(J) = 0
Next J
For I = 1 To nrc
  big = 0
  For J = 1 To nrc
    If ipivot(J) <> U Then
      For k = 1 To nrc
        If ipivot(k) = 0 Then
          If Abs(Mout(J, k)) >= big Then
            big = Abs(Mout(J, k))
            irow = J
            icol = k
          End If
          ElseIf ipivot(k) > 1 Then Exit Sub
        End If
      Next k
    End If
  Next J
  ipivot(icol) = ipivot(icol) + 1
  If irow <> icol Then
    For L = 1 To nrc
      dummy = Mout(irow, L)
      Mout(irow, L) = Mout(icol, L)
      Mout(icol, L) = dummy
    Next L
    For L = 1 To nrc
      dummy = Bb(irow, L)
      Bb(irow, L) = Bb(icol, L)
      Bb(icol, L) = dummy
    Next L
  End If
  indexr(I) = irow
  indexc(I) = icol
  If Mout(icol, icol) = 0 Then Exit Sub
  pivinv = U / Mout(icol, icol)
  Mout(icol, icol) = U
  For L = 1 To nrc
    Mout(icol, L) = Mout(icol, L) * pivinv
    Bb(icol, L) = Bb(icol, L) * pivinv
  Next L
  For LL = 1 To nrc
    If LL <> icol Then
      dummy = Mout(LL, icol)
      Mout(LL, icol) = 0
      For L = 1 To nrc
        Mout(LL, L) = Mout(LL, L) - Mout(icol, L) * dummy
        Bb(LL, L) = Bb(LL, L) - Bb(icol, L) * dummy
      Next L
    End If
  Next LL
Next I
For L = nrc To 1 Step -1
  If indexr(L) <> indexc(L) Then
    For k = 1 To nrc
      dummy = Mout(k, indexr(L))
      Mout(k, indexr(L)) = Mout(k, indexc(L))
      Mout(k, indexc(L)) = dummy
    Next k
  End If
Next L
'Erase indexc, indexr, ipivot

End Sub

'matrix multiplication
Public Sub MatrixMultiplication(ByRef a() As Double, _
                                 ByVal nRow As Integer, _
                                 ByVal nCol As Integer, _
                                 ByRef b() As Double, _
                                 ByVal nCol2 As Integer, _
                                 ByRef AB() As Double)
Dim I, J, k As Integer

For I = 1 To nRow
  For J = 1 To nCol2
    AB(I, J) = 0
    For k = 1 To nCol
      AB(I, J) = AB(I, J) + a(I, k) * b(k, J)
    Next k
  Next J
Next I

End Sub

'matrix transpose
Public Sub MatrixTranspose(ByRef a() As Double, _
                            ByVal nRow As Integer, _
                            ByVal nCol As Integer, _
                            ByRef At() As Double)

Dim I, J As Integer
ReDim At(1 To nCol, 1 To nRow) As Double

For I = 1 To nCol
  For J = 1 To nRow
    At(I, J) = a(J, I)
  Next J
Next I

End Sub


Public Function Max(ByVal x1 As Double, ByVal x2 As Double) As Double
    If (x1 > x2) Then
        Max = x1
    Else
        Max = x2
    End If
End Function

Public Function Min(ByVal x1 As Double, ByVal x2 As Double) As Double
    If (x1 > x2) Then
        Min = x2
    Else
        Min = x1
    End If
End Function

Private Function VBARoundDown(x As Double, Num As Integer) As Double
    Dim y As Double
    y = Round(x, Num)
    If (x > 0) Then
        If (y > x) Then
            y = y - 10 ^ (-Num)
        End If
    Else
        If (y < x) Then
            y = y + 10 ^ (-Num)
        End If
    End If
        
    VBARoundDown = y

End Function


Public Function EXF(ByVal a As Double, ByVal b As Double) As Double
'   EXF(A, B) = exp(A)erfc(B)
'   erfc(B) = \frac{2}{\sqrt{\pi}}\int_B^\infty e^{-t^2} dt
    Dim C, x, y, T, cc, res As Double

      EXF = 0#
      If ((Abs(a) > 100#) And (b <= 0#)) Then
        Exit Function
      End If
        
      C = a - b * b
      
      If ((Abs(C) > 100#) And (b >= 0#)) Then
        Exit Function
      End If
        
      If (C < -100#) Then
         If (b < 0#) Then
            EXF = 2# * Exp(a) - EXF
         End If
        Exit Function
      End If
      
      x = Abs(b)
      If (x > 3#) Then
         y = 0.5641896 / (x + 0.5 / (x + 1# / (x + 1.5 / (x + 2# / (x + 2.5 / x + 1#)))))
         EXF = y * Exp(C)
         If (b < 0#) Then
            EXF = 2# * Exp(a) - EXF
         End If
        Exit Function
      End If
      
      T = 1# / (1# + 0.3275911 * x)
      cc = T * (0.2844967 - T * (1.421414 - T * (1.453152 - 1.061405 * T)))
      y = T * (0.2548296 - cc)
'      Y = T * (0.2548296 - T * (0.2844967 - T * (1.421414 - T * (1.453152 - 1.061405 * T))))
      EXF = y * Exp(C)
      If (b < 0#) Then
         EXF = 2# * Exp(a) - EXF
      End If
        Exit Function
End Function



Public Function VBAERF(x As Double) As Double
'ERF Error function.
'   Y = ERF(X) is the error function for each element of X.  X must be
'   real. The error function is defined as:
'
'     erf(x) = 2/sqrt(pi) * integral from 0 to x of exp(-t^2) dt.
'
'   See also ERFC, ERFCX, ERFINV.

'   Ref: Abramowitz & Stegun, Handbook of Mathematical Functions, sec. 7.1.

'   Copyright 1984-2002 The MathWorks, Inc.
'   $Revision: 5.13 $  $Date: 2002/04/09 00:29:47 $

' Derived from a FORTRAN program by W. J. Cody.
' See ERFCORE.

'if ~isreal(x), error('X must be real.'); end
'siz = size(x);
'x = x(:);
    VBAERF = VBAERFCORE(x, 0)
'y = reshape(y,siz);

End Function

Public Function VBAERFC(x As Double) As Double
    VBAERFC = VBAERFCORE(x, 1)
End Function

Private Function VBAERFCORE(x As Double, jint As Integer) As Double
'function result = erfcore(x,jint)
'ERFCORE Core algorithm for error functions.
'   erf(x) = erfcore(x,0)
'   erfc(x) = erfcore(x,1)
'   erfcx(x) = exp(x^2)*erfc(x) = erfcore(x,2)

'   C. Moler, 2-1-91.
'   Copyright 1984-2002 The MathWorks, Inc.
'   $Revision: 5.15 $  $Date: 2002/04/09 00:29:47 $

'   This is a translation of a FORTRAN program by W. J. Cody,
'   Argonne National Laboratory, NETLIB/SPECFUN, March 19, 1990.
'   The main computation evaluates near-minimax approximations
'   from "Rational Chebyshev approximations for the error function"
'   by W. J. Cody, Math. Comp., 1969, PP. 631-638.

'   Note: This M-file is intended to document the algorithm.
'   If a .DLL file or .MEX file for a particular architecture exists,
'   it will be executed instead, but its functionality is the same.
'#mex

'    if ~isreal(x),
'       error('Input argument must be real.')
'    End
'    result = repmat(NaN,size(x));
'
'   evaluate  erf  for  |x| <= 0.46875
'
    Dim y, Z, xden, xnum, res, del As Double
        If (Abs(x) <= 0.46875) Then
            y = Abs(x)
            Z = y * y
            xnum = 0.185777706184603 * Z
            xden = Z
' 1
            xnum = (xnum + 3.16112374387057) * Z
            xden = (xden + 23.6012909523441) * Z
' 2
            xnum = (xnum + 113.86415415105) * Z
            xden = (xden + 244.024637934444) * Z
' 3
            xnum = (xnum + 377.485237685302) * Z
            xden = (xden + 1282.61652607737) * Z
            
            res = x * (xnum + 3209.37758913847) / (xden + 2844.23683343917)
            If (jint <> 0) Then
                res = 1 - res
            End If
                
            If (jint = 2) Then
                res = Exp(Z) * res
            End If
 '           Return
        End If

'   evaluate  erfc  for 0.46875 <= |x| <= 4.0
'
    If ((Abs(x) > 0.46875) And (Abs(x) <= 4#)) Then
        
            y = Abs(x)
            xnum = 2.15311535474404E-08 * y
            xden = y
'1
               xnum = (xnum + 0.56418849698867) * y
               xden = (xden + 15.7449261107098) * y
'2
               xnum = (xnum + 8.88314979438838) * y
               xden = (xden + 117.693950891312) * y
'3
               xnum = (xnum + 66.1191906371416) * y
               xden = (xden + 537.18110186201) * y
'4
               xnum = (xnum + 298.6351381974) * y
               xden = (xden + 1621.38957456669) * y
               
'5
               xnum = (xnum + 881.952221241769) * y
               xden = (xden + 3290.79923573346) * y
'6
               xnum = (xnum + 1712.04761263407) * y
               xden = (xden + 4362.61909014325) * y
'7
               xnum = (xnum + 2051.07837782607) * y
               xden = (xden + 3439.36767414372) * y
               
               res = (xnum + 1230.339354798) / (xden + 1230.33935480375)
               If (jint <> 2) Then
                    Z = VBARoundDown(y * 16, 6) / 16
                    del = (y - Z) * (y + Z)
                    res = Exp(-Z * Z) * Exp(-del) * res
                End If
'                Return
            End If
    
'   evaluate  erfc  for |x| > 4.0
    If (Abs(x) > 4#) Then
            y = Abs(x)
            Z = 1# / (y * y)
            xnum = 1.63153871373021E-02 * Z
            xden = Z
'1
               xnum = (xnum + 0.305326634961232) * Z
               xden = (xden + 2.56852019228982) * Z
'2
               xnum = (xnum + 0.360344899949804) * Z
               xden = (xden + 1.87295284992346) * Z
'3
               xnum = (xnum + 0.125781726111229) * Z
               xden = (xden + 0.527905102951428) * Z
'4
               xnum = (xnum + 1.60837851487423E-02) * Z
               xden = (xden + 6.05183413124413E-02) * Z

            res = Z * (xnum + 6.58749161529838E-04) / (xden + 2.33520497626869E-03)
            res = (1 / Sqr(3.1415926) - res) / y
            If (jint <> 2) Then
               Z = VBARoundDown(y * 16, 10) / 16
               del = (y - Z) * (y + Z)
               res = Exp(-Z * Z) * Exp(-del) * res
'               k = find(~isfinite(result));
'               result(k) = 0*k;
            End If
'            Return
    End If
    
'   fix up for negative argument, erf, etc.
'
    If (jint = 0) Then
            If (x > 0.46875) Then
                res = (0.5 - res) + 0.5
            End If
                
            If (x < -0.46875) Then
                res = (-0.5 + res) - 0.5
            End If
    ElseIf (jint = 1) Then
        If (x < -0.46875) Then
            res = 2# - res
        End If
    Else 'jint == 2
            If (x < -0.46875) Then
                Z = VBARoundDown(x * 16, 0.000001) / 16
            del = (x - Z) * (x + Z)
            y = Exp(Z * Z) * Exp(del)
            res = (y + y) - res
            End If
    End If
    
    VBAERFCORE = res
End Function

'function for initial value problem
'Eq. 3.37 Toride et al. 1995
Public Function IVP(ByVal ci As Double, _
                    ByVal T As Double, _
                    ByVal Z As Double, _
                    ByVal Zi As Double, _
                    ByVal p As Double, _
                    ByVal beta As Double, _
                    ByVal omega As Double, _
                    Optional ByVal R As Variant, _
                    Optional ByVal mu1 As Variant, _
                    Optional ByVal mu2 As Variant, _
                    Optional ByVal strFluxConc As Variant, _
                    Optional ByVal tol As Variant, _
                    Optional ByVal strNI As Variant, _
                    Optional ByVal bln As Variant, _
                    Optional ByVal nPoint As Variant) As Double
    
    If (IsMissing(R)) Then
        R = 1#
    End If
        
    If (IsMissing(tol)) Then
        tol = 0.0001
    End If
        
    If (IsMissing(strNI)) Then
        strNI = "chebyshev"
    End If
        
    If (IsMissing(bln)) Then
        bln = False   'Andrew Moore 01/07/2010
    End If
        
    Select Case LCase(strNI)
        Case "romberg"
            IVP = ROMB("EqInt337", 0.0000000001, T - 0.0000000001, tol, bln, T, Z, Zi, p, beta, omega, R, mu1, mu2, strFluxConc)
        
        Case "simpson"
            IVP = AdaptSim("EqInt337", 0.0000000001, T - 0.0000000001, tol, bln, T, Z, Zi, p, beta, omega, R, mu1, mu2, strFluxConc)
        
        Case "lobatto"
            IVP = AdaptLob("EqInt337", 0.0000000001, T - 0.0000000001, tol, bln, T, Z, Zi, p, beta, omega, R, mu1, mu2, strFluxConc)
        
        Case "chebyshev"
            If (IsMissing(nPoint)) Then
                nPoint = 50
            End If
            
            If (nPoint < 1) Then
                nPoint = 1
            End If
            IVP = ChebyQ("EqInt337", 0.0000000001, T - 0.0000000001, nPoint, bln, T, Z, Zi, p, beta, omega, R, mu1, mu2, strFluxConc)
    End Select
    
'    IVP = ROMBLN("EqInt337", 0.0000000001, T - 0.0000000001, tol, T, Z, Zi, p, beta, omega, r, mu1, mu2, bFluxConc)
    IVP = IVP * omega / beta / R * ci + Exp(-omega * T / beta / R) * Psi1N(Z, 0#, T, p, beta, R, mu1, strFluxConc)
    
End Function

Private Function Psi1N(ByVal Z As Double, _
                       ByVal Zi As Double, _
                       ByVal tau As Double, _
                       ByVal p As Double, _
                       ByVal beta As Double, _
                       Optional ByVal R As Variant, _
                       Optional ByVal mu As Variant, _
                       Optional ByVal strFluxConc As Variant) As Double

If IsMissing(R) Then
    R = 1#
End If

If IsMissing(mu) Then
    mu = 0#
End If

If IsMissing(strFluxConc) Then
    strFluxConc = "flux"
End If

Select Case LCase(strFluxConc)
    Case "flux"
        Psi1N = 1# - 0.5 * VBAERFC((beta * R * (Z - Zi) - tau) / Sqr(4# * beta * R * tau / p)) _
                   - 0.5 * EXF(p * Z, (beta * R * (Z + Zi) + tau) / Sqr(4# * beta * R * tau / p)) _
                   + Sqr(beta * R / (4# * PI * p * tau)) * ( _
                     Exp(p * Z - p * (beta * R * (Z + Zi) + tau) * (beta * R * (Z + Zi) + tau) / (4# * beta * R * tau) _
                   - Exp(-p * (beta * R * (Z - Zi) - tau) * (beta * R * (Z - Zi) - tau) / (4# * beta * R * tau))))
    Case Else
'Else
        Psi1N = 1# - 0.5 * VBAERFC((beta * R * (Z - Zi) - tau) / Sqr(4# * beta * R * tau / p)) _
                   - Sqr(p * tau / PI / beta / R) * Exp(p * Z - p * (beta * R * (Z + Zi) + tau) * (beta * R * (Z + Zi) + tau) / (4# * beta * R * tau)) _
                   + 0.5 * (1# + p * (Z + Zi) + p * tau / beta / R) * EXF(p * Z, (beta * R * (Z + Zi) + tau) / Sqr(4# * beta * R * tau / p))
    End Select
End If
    
    Psi1N = Psi1N * Exp(-mu * tau / beta / R)
    
End Function

'function to calculate the integrand in Eq. 3.37 in Toride et al. 1995
Private Function EqInt337(ByVal tau As Double, _
                        ByVal T As Double, _
                        ByVal Z As Double, _
                        ByVal Zi As Double, _
                        ByVal p As Double, _
                        ByVal beta As Double, _
                        ByVal omega As Double, _
                        Optional ByVal R As Variant, _
                        Optional ByVal mu1 As Variant, _
                        Optional ByVal mu2 As Variant, _
                        Optional ByVal strFluxConc As Variant) As Double
    
    Dim ax, bx As Double
    
    If IsMissing(R) Then
        R = 1#
    End If

    If IsMissing(mu2) Then
        mu2 = 0#
    End If

    If IsMissing(strFluxConc) Then
        strFluxConc = "flux"
    End If
    
    ax = -omega * tau / beta / R - (omega + mu2) * (T - tau) / (1# - beta) / R
    bx = 2# * omega / R * Sqr((T - tau) * tau / beta / (1# - beta))
    
    EqInt337 = EXPBI0(ax, bx) + Sqr(beta * tau / (1# - beta) / (T - tau)) * EXPBI1(ax, bx)
    EqInt337 = EqInt337 * Psi1N(Z, Zi, tau, p, beta, R, mu1, strFluxConc)
End Function

'production value problem
Public Function PVP(ByVal gamma As Double, _
                      ByVal T As Double, _
                      ByVal Z As Double, _
                      ByVal Zi As Double, _
                      ByVal p As Double, _
                      ByVal beta As Double, _
                      ByVal omega As Double, _
                      Optional ByVal R As Variant, _
                      Optional ByVal mu1 As Variant, _
                      Optional ByVal mu2 As Variant, _
                      Optional ByVal strFluxConc As Variant, _
                      Optional ByVal tol As Variant, _
                      Optional ByVal strNI As Variant, _
                      Optional ByVal bln As Variant, _
                      Optional ByVal nPoint As Variant) As Double
    
    If (IsMissing(R)) Then
        R = 1#
    End If
    
    If (IsMissing(tol)) Then
        tol = 0.0001
    End If
    
    If (IsMissing(strNI)) Then
        strNI = "chebyshev"
    End If
    
    If (IsMissing(bln)) Then
        bln = False
    End If
    
    Select Case LCase(strNI)
        Case "romberg"
            PVP = ROMB("EqInt347", 0.0000000001, T - 0.0000000001, tol, bln, T, Z, Zi, p, beta, omega, R, mu1, mu2, strFluxConc)
        
        Case "simpson"
            PVP = AdaptSim("EqInt347", 0.0000000001, T - 0.0000000001, tol, bln, T, Z, Zi, p, beta, omega, R, mu1, mu2, strFluxConc)
        
        Case "lobatto"
            PVP = AdaptLob("EqInt347", 0.0000000001, T - 0.0000000001, tol, bln, T, Z, Zi, p, beta, omega, R, mu1, mu2, strFluxConc)
        
        Case "chebyshev"
            If (IsMissing(nPoint)) Then
                nPoint = 50
            End If
            If (nPoint < 1) Then
                nPoint = 1
            End If
            PVP = ChebyQ("EqInt347", 0.0000000001, T - 0.0000000001, nPoint, bln, T, Z, Zi, p, beta, omega, R, mu1, mu2, strFluxConc)
    End Select
    
'    PVP = ROMBLN("EqInt347", 0.0000000001, T, tol, T, Z, Zi, p, beta, omega, r, mu1, mu2, bFluxConc)
    PVP = PVP / beta / R * gamma
End Function


'function to calculate the integrand in Eq. 3.47 in Toride et al. 1995
Private Function EqInt347(ByVal tau As Double, _
                        ByVal T As Double, _
                        ByVal Z As Double, _
                        ByVal Zi As Double, _
                        ByVal p As Double, _
                        ByVal beta As Double, _
                        ByVal omega As Double, _
                        Optional ByVal R As Variant, _
                        Optional ByVal mu1 As Variant, _
                        Optional ByVal mu2 As Variant, _
                        Optional ByVal strFluxConc As Variant) As Double
    
    Dim GOLD As Double
    If IsMissing(R) Then
        R = 1#
    End If

    If IsMissing(mu2) Then
        mu2 = 0#
    End If

    If IsMissing(strFluxConc) Then
        strFluxConc = "flux"
    End If
    
    GOLD = Goldstein(omega * omega * tau / beta / R / (omega + mu2), (omega + mu2) * (T - tau) / (1# - beta) / R)
    EqInt347 = Exp(-omega * mu2 * tau / (omega + mu2) / beta / R) * GOLD
    EqInt347 = EqInt347 * Psi1N(Z, Zi, tau, p, beta, R, mu1, strFluxConc)
    
End Function

Public Function ChebyQ(ByVal func As String, _
                         ByVal a As Double, _
                         ByVal b As Double, _
                         ByVal n As Integer, _
                         Optional ByVal ln As Variant, _
                         Optional ByVal P1 As Variant, _
                         Optional ByVal P2 As Variant, _
                         Optional ByVal P3 As Variant, _
                         Optional ByVal P4 As Variant, _
                         Optional ByVal P5 As Variant, _
                         Optional ByVal P6 As Variant, _
                         Optional ByVal P7 As Variant, _
                         Optional ByVal P8 As Variant, _
                         Optional ByVal P9 As Variant, _
                         Optional ByVal P10 As Variant) As Double

    Dim k As Integer
    Dim qq, tk, xk, fk As Double
    
    If (Abs(b - a) < 0.0000000001) Then
        ChebyQ = 0#
        Exit Function
    End If
    
    If (b < a) Then
        ChebyQ = 0#
        Exit Function
    End If
      
    If (IsMissing(ln)) Then
        ln = False
    End If
        
    qq = 0#
    
    For k = 1 To n
        tk = (2# * k - 1#) * PI / 2# / n
        tk = Cos(tk)
        If (ln) Then
            xk = 0.5 * ((Log(b) - Log(a)) * tk + (Log(a) + Log(b)))
            fk = Run(func, Exp(xk), P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * Exp(xk)
        Else
            xk = 0.5 * ((b - a) * tk + (a + b))
            fk = Run(func, xk, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
        End If
        fk = fk * Sqr(1# - tk * tk)
        qq = qq + fk
    Next
    
    If (ln) Then
        ChebyQ = qq * 0.5 * (Log(b) - Log(a)) * PI / n
    Else
        ChebyQ = qq * 0.5 * (b - a) * PI / n
    End If

End Function

Public Function ROMB(ByVal func As String, _
                      ByVal a As Double, _
                      ByVal b As Double, _
                      Optional ByVal tol As Variant, _
                      Optional ByVal ln As Variant, _
                      Optional ByVal P1 As Variant, _
                      Optional ByVal P2 As Variant, _
                      Optional ByVal P3 As Variant, _
                      Optional ByVal P4 As Variant, _
                      Optional ByVal P5 As Variant, _
                      Optional ByVal P6 As Variant, _
                      Optional ByVal P7 As Variant, _
                      Optional ByVal P8 As Variant, _
                      Optional ByVal P9 As Variant, _
                      Optional ByVal P10 As Variant) _
                      As Double

Dim R(1 To 15, 1 To 15) As Double
Dim I, J, k, Level As Integer
Dim h As Double
Dim Er As Double

If IsMissing(tol) Then
    tol = 0.0001
End If

If (a >= b) Then
    ROMB = 0#
    Exit Function
End If
      
If (IsMissing(ln)) Then
    ln = False
End If

      
Level = 15
      
If (ln) Then
    R(1, 1) = 0.5 * (Log(b) - Log(a)) * (Run(func, a, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * a + Run(func, b, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * b)
    h = Log(b) - Log(a)
Else
    R(1, 1) = 0.5 * (b - a) * (Run(func, a, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) + Run(func, b, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10))
    h = b - a
End If

For J = 2 To Level
    
'   From R(k, 1) to R(k + 1, 1)
    h = h / 2#
    R(J, 1) = 0.5 * R(J - 1, 1)
    For I = 1 To 2 ^ (J - 2)
        If (ln) Then
            R(J, 1) = R(J, 1) + h * Run(func, Exp(Log(a) + (2 * I - 1) * h), P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * Exp(Log(a) + (2 * I - 1) * h)
        Else
            R(J, 1) = R(J, 1) + h * Run(func, a + (2 * I - 1) * h, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
        End If
    Next
    
    
    For k = 2 To J
        R(J, k) = (4 ^ (k - 1) * R(J, k - 1) - R(J - 1, k - 1)) / (4 ^ (k - 1) - 1)
        
        If (Abs(R(J, k) - R(J, k - 1)) < tol * Abs(R(J, k)) Or R(J, k) < tol) Then
            ROMB = R(J, k)
            Exit Function
        End If
        
    Next
    
    If (Abs(R(J, J) - R(J - 1, J - 1)) < tol * Abs(R(J, J)) Or R(J, J) < tol) Then
        ROMB = R(J, J)
        Exit Function
    End If
    
Next
ROMB = R(Level, Level)
End Function

Public Function AdaptSim(ByVal func As String, _
                         ByVal a As Double, _
                         ByVal b As Double, _
                         Optional ByVal tol As Variant, _
                         Optional ByVal ln As Variant, _
                         Optional ByVal P1 As Variant, _
                         Optional ByVal P2 As Variant, _
                         Optional ByVal P3 As Variant, _
                         Optional ByVal P4 As Variant, _
                         Optional ByVal P5 As Variant, _
                         Optional ByVal P6 As Variant, _
                         Optional ByVal P7 As Variant, _
                         Optional ByVal P8 As Variant, _
                         Optional ByVal P9 As Variant, _
                         Optional ByVal P10 As Variant) As Double

Dim m, fm, fa, fb As Double

If (IsMissing(tol)) Then
    tol = 0.0001
End If
    
If (b <= a) Then
    AdaptSim = 0#
    Exit Function
End If
    
If (IsMissing(ln)) Then
    ln = False
End If

If (ln) Then
    m = (Log(b) + Log(a)) / 2#
    fm = Run(func, Exp(m), P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * Exp(m)
    fa = Run(func, a, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * a
    fb = Run(func, b, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * b
Else
    m = (b + a) / 2#
    fm = Run(func, m, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
    fa = Run(func, a, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
    fb = Run(func, b, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
End If

AdaptSim = AdaptSimStp(func, a, b, fa, fm, fb, tol, ln, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)

End Function

Public Function AdaptSimStp(ByVal func As String, _
                         ByVal a As Double, _
                         ByVal b As Double, _
                         ByVal fa As Double, _
                         ByVal fm As Double, _
                         ByVal fb As Double, _
                         Optional ByVal tol As Variant, _
                         Optional ByVal ln As Variant, _
                         Optional ByVal P1 As Variant, _
                         Optional ByVal P2 As Variant, _
                         Optional ByVal P3 As Variant, _
                         Optional ByVal P4 As Variant, _
                         Optional ByVal P5 As Variant, _
                         Optional ByVal P6 As Variant, _
                         Optional ByVal P7 As Variant, _
                         Optional ByVal P8 As Variant, _
                         Optional ByVal P9 As Variant, _
                         Optional ByVal P10 As Variant) As Double

Dim m, h, fml, fmr As Double
Dim i1, i2 As Double

If (a >= b) Then
    AdaptSimStp = 0#
    Exit Function
End If

If (IsMissing(tol)) Then
    tol = 0.0001
End If
    
If (IsMissing(ln)) Then
    ln = False
End If
    
If (ln) Then
    m = (Log(b) + Log(a)) / 2#
    h = (Log(b) - Log(a)) / 4#
    fml = Run(func, Exp(Log(a) + h), P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * Exp(Log(a) + h)
    fmr = Run(func, Exp(Log(b) - h), P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * Exp(Log(b) - h)
    m = Exp(m)
Else
    m = (b + a) / 2#
    h = (b - a) / 4#
    fml = Run(func, a + h, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
    fmr = Run(func, b - h, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
End If


i1 = h / 1.5 * (fa + 4# * fm + fb)
i2 = h / 3# * (fa + 4# * fml + 2# * fm + 4# * fmr + fb)
i1 = (16 * i2 - i1) / 15

If ((Abs(i1 - i2) < tol * i1) Or (i1 < tol)) Then
    AdaptSimStp = i1
Else
    AdaptSimStp = AdaptSimStp(func, a, m, fa, fml, fm, tol, ln, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) + _
                  AdaptSimStp(func, m, b, fm, fmr, fb, tol, ln, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
End If

End Function

Public Function AdaptLob(ByVal func As String, _
                         ByVal a As Double, _
                         ByVal b As Double, _
                         Optional ByVal tol As Variant, _
                         Optional ByVal ln As Variant, _
                         Optional ByVal P1 As Variant, _
                         Optional ByVal P2 As Variant, _
                         Optional ByVal P3 As Variant, _
                         Optional ByVal P4 As Variant, _
                         Optional ByVal P5 As Variant, _
                         Optional ByVal P6 As Variant, _
                         Optional ByVal P7 As Variant, _
                         Optional ByVal P8 As Variant, _
                         Optional ByVal P9 As Variant, _
                         Optional ByVal P10 As Variant) As Double
'   rewritten from
'   ADAPTLOB  Numerically evaluate integral using adaptive Lobatto rule.
'   See also ADAPTLOBSTP.
'   Walter Gautschi, 08/03/98
'   Reference: Gander, Computermathematik, Birkhaeuser, 1992.

Dim m, h, alpha, beta As Double
Dim fa, fb As Double

If (a >= b) Then
    AdaptLob = 0#
    Exit Function
End If
    
If IsMissing(tol) Then
    tol = 0.0001
End If

If IsMissing(ln) Then
    ln = False
End If

fa = Run(func, a, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
fb = Run(func, b, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)

AdaptLob = AdaptLobStp(func, a, b, fa, fb, tol, ln, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)

End Function

Public Function AdaptLobStp(ByVal func As String, _
                         ByVal a As Double, _
                         ByVal b As Double, _
                         ByVal fa As Double, _
                         ByVal fb As Double, _
                         Optional ByVal tol As Variant, _
                         Optional ByVal ln As Variant, _
                         Optional ByVal P1 As Variant, _
                         Optional ByVal P2 As Variant, _
                         Optional ByVal P3 As Variant, _
                         Optional ByVal P4 As Variant, _
                         Optional ByVal P5 As Variant, _
                         Optional ByVal P6 As Variant, _
                         Optional ByVal P7 As Variant, _
                         Optional ByVal P8 As Variant, _
                         Optional ByVal P9 As Variant, _
                         Optional ByVal P10 As Variant) As Double
'ADAPTLOBSTP  Recursive function used by ADAPTLOB.
'   See also ADAPTLOB.
'   Walter Gautschi, 08/03/98
Dim h, m, alpha, beta, mll, ml, mr, mrr As Double
Dim fmll, fm, fml, fmr, fmrr, i1, i2, eps As Double

If (a >= b) Then
    AdaptLobStp = 0#
    Exit Function
End If

If IsMissing(tol) Then
    tol = 0.0001
End If

If IsMissing(ln) Then
    ln = False
End If

alpha = Sqr(2# / 3#)
beta = 1# / Sqr(5#)

If (ln) Then
    h = (Log(b) - Log(a)) / 2#
    m = (Log(a) + Log(b)) / 2#
    
    mll = Exp(m - alpha * h)
    ml = Exp(m - beta * h)
    mr = Exp(m + beta * h)
    mrr = Exp(m + alpha * h)
    m = Exp(m)
    
    fmll = Run(func, mll, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * mll
    fml = Run(func, ml, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * ml
    fm = Run(func, m, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * m
    fmr = Run(func, mr, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * mr
    fmrr = Run(func, mrr, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) * mrr
Else
    h = (b - a) / 2#
    m = (a + b) / 2#
    
    mll = m - alpha * h
    ml = m - beta * h
    mr = m + beta * h
    mrr = m + alpha * h
    
    fmll = Run(func, mll, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
    fml = Run(func, ml, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
    fm = Run(func, m, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
    fmr = Run(func, mr, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
    fmrr = Run(func, mrr, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
End If

i2 = (h / 6#) * (fa + fb + 5# * (fml + fmr))
i1 = (h / 1470#) * (77# * (fa + fb) + 432# * (fmll + fmrr) + 625# * (fml + fmr) + 672# * fm)

If (Abs(i1 - i2) < tol * i1) Or (mll <= a) Or (b <= mrr) Or (i1 < tol) Then
    AdaptLobStp = i1
Else
    AdaptLobStp = AdaptLobStp(func, a, mll, fa, fmll, tol, ln, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) + _
                  AdaptLobStp(func, mll, ml, fmll, fml, tol, ln, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) + _
                  AdaptLobStp(func, ml, m, fml, fm, tol, ln, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) + _
                  AdaptLobStp(func, m, mr, fm, fmr, tol, ln, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) + _
                  AdaptLobStp(func, mr, mrr, fmr, fmrr, tol, ln, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10) + _
                  AdaptLobStp(func, mrr, b, fmrr, fb, tol, ln, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10)
  End If

End Function



