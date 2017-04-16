VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DlgPropagate 
   Caption         =   "CXTFIT/Excel Error Propagation"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   OleObjectBlob   =   "DlgPropagate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DlgPropagate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Initialize()
    If (NameExists("ParameterRange", "Local")) Then
        reParameter.Text = "ParameterRange"
    End If
    
    If (NameExists("PredictionRange", "Local")) Then
        rePrediction.Text = "PredictionRange"
    End If

    If (Not (NameExists("OffsetParaPtb", "Local"))) Then
        cbParaPtb.Value = False
        tbParaPtb.Enabled = False
        sbParaPtb.Enabled = False
        tbParaPtb.Value = -2
        sbParaPtb.Value = -2
    Else
        ParaPtb = Evaluate("OffsetParaPtb")
        cbParaPtb.Value = True
        tbParaPtb.Enabled = True
        sbParaPtb.Enabled = True
        tbParaPtb.Value = ParaPtb
        sbParaPtb.Value = ParaPtb
    End If

    If (Not (NameExists("OffsetPredJcb", "Local"))) Then
        tbPredJcb.Enabled = False
        sbPredJcb.Enabled = False
        tbPredJcb.Value = 1
        sbPredJcb.Value = 1
        cbPredJcb.Value = False
    Else
        PredJcb = Evaluate("OffsetPredJcb")
        tbPredJcb.Value = PredJcb
        sbPredJcb.Value = PredJcb
        cbPredJcb.Value = True
    End If
    
    If (Not (NameExists("OffsetPredEP", "Local"))) Then
        tbPredEP.Value = 2
        sbPredEP.Value = 2
    Else
        PredEP = Evaluate("OffsetPredEP")
        tbPredEP.Value = PredEP
        sbPredEP.Value = PredEP
    End If
End Sub

Private Sub cbParaPtb_Click()
    If (cbParaPtb.Value = True) Then
        tbParaPtb.Enabled = True
        sbParaPtb.Enabled = True
        tbParaPtb.Value = sbParaPtb.Value
    Else
        tbParaPtb.Enabled = False
        sbParaPtb.Enabled = False
    End If
        
End Sub


Private Sub cbPredJcb_Click()

    If (cbPredJcb.Value = True) Then
        tbPredJcb.Enabled = True
        sbPredJcb.Enabled = True
        tbPredJcb.Value = sbPredJcb.Value
    Else
        tbPredJcb.Value = 1
        tbPredJcb.Enabled = False
        sbPredJcb.Enabled = False
    End If
        
End Sub

Private Sub sbParaPtb_Change()
    tbParaPtb.Value = sbParaPtb.Value
End Sub

Private Sub sbPredJcb_Change()
    tbPredJcb.Value = sbPredJcb.Value
End Sub

Private Sub tbPredJcb_Change()
    If (tbPredJcb.Value > 0) Then
        sbPredJcb.Value = tbPredJcb.Value
    Else
        MsgBox ("Offset from prediction range for output of Jacobian needs to be greater to avoid overwritting!")
    End If
End Sub

Private Sub sbPredEP_Change()
    tbPredEP.Value = sbPredEP.Value
End Sub

Private Sub tbPredEP_Change()
    If (tbPredEP.Value > 0) Then
        sbPredEP.Value = tbPredEP.Value
    Else
'        MsgBox ("Offset from prediction range for output of Jacobian needs to be greater to avoid overwritting!")
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCalculate_Click()
    Dim ParaPtb, PredJcb, PredEP As Integer

    PredEP = tbPredEP.Value
    If (PredEP > 0) Then
'        If (Not NameExists("OffsetPredEP", "Local")) Then
            Names.Add Name:="'" + ActiveSheet.Name + "'" + "!OffsetPredEP", _
                      RefersTo:=PredEP, Visible:=True
        
'        End If
    Else
        MsgBox "Offset for error propagation output is expected to be greater than 0"
        Return
    End If
   
    If (cbParaPtb.Value = True) Then
        ParaPtb = tbParaPtb.Value
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!OffsetParaPtb", _
                  RefersTo:=ParaPtb, Visible:=True
    Else
        If (NameExists("OffsetParaPtb", "Local")) Then
            Names("'" + ActiveSheet.Name + "'" + "!OffsetParaPtb").Delete
        End If
    End If
        
    If (cbPredJcb.Value = True) Then
        PredJcb = tbPredJcb.Value
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!OffsetPredJcb", _
                  RefersTo:=PredJcb, Visible:=True
    Else
        If (NameExists("OffsetPredJcb", "Local")) Then
            Names("'" + ActiveSheet.Name + "'" + "!OffsetPredJcb").Delete
        End If
    End If
        
    If (reParameter = "") Then
        MsgBox ("Parameter range not selected!")
        Exit Sub
    End If
    
    If (rePrediction = "") Then
        MsgBox ("Prediction range not selected!")
        Exit Sub
    End If
    
    If (reParameter.Text <> "ParameterRange") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!ParameterRange", _
                  RefersTo:="=" + reParameter, Visible:=True
    End If
    
    If (rePrediction.Text <> "PredictionRange") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!PredictionRange", _
                  RefersTo:="=" + rePrediction, Visible:=True
    End If
   
    Call Propagate
        
    Unload Me
End Sub



