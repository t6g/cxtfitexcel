VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DlgSolveOption 
   Caption         =   "CXTFIT/Excel Solve Option"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8205
   OleObjectBlob   =   "DlgSolveOption.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DlgSolveOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Dim ParaStd, ParaPtb, PredStd, PredJcb As Integer
        
    If (NameExists("NoParaConstraint", "Local")) Then
        cbConstraint.Value = False
    Else
        cbConstraint.Value = True
    End If
    
    If (Not (NameExists("PenaltyCell", "Local"))) Then
        cbPenalty.Value = False
        rePenalty.Enabled = False
    Else
        cbPenalty.Value = True
        rePenalty.Enabled = True
        rePenalty.Text = "PenaltyCell"
    End If

    If (Not (NameExists("OffsetParaStd", "Local"))) Then
        cbParaStd.Value = False
        tbParaStd.Enabled = False
        sbParaStd.Enabled = False
        tbParaStd.Value = -1
        sbParaStd.Value = -1
    Else
        ParaStd = Evaluate("OffsetParaStd")
        cbParaStd.Value = True
        tbParaStd.Enabled = True
        sbParaStd.Enabled = True
        tbParaStd.Value = ParaStd
        sbParaStd.Value = ParaStd
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

    If (Not (NameExists("OffsetPredStd", "Local"))) Then
        cbPredStd.Value = False
        tbPredStd.Enabled = False
        sbPredStd.Enabled = False
        tbPredStd.Value = -2
        sbPredStd.Value = -2
    Else
        PredStd = Evaluate("OffsetPredStd")
        cbPredStd.Value = True
        tbPredStd.Enabled = True
        sbPredStd.Enabled = True
        tbPredStd.Value = PredStd
        sbPredStd.Value = PredStd
    End If
   
    If (Not (NameExists("OffsetPredJcb", "Local"))) Then
        cbPredJcb.Value = False
        tbPredJcb.Enabled = False
        sbPredJcb.Enabled = False
        tbPredJcb.Value = 2
        sbPredJcb.Value = 2
    Else
        PredJcb = Evaluate("OffsetPredJcb")
        cbPredJcb.Value = True
        tbPredJcb.Enabled = True
        sbPredJcb.Enabled = True
        tbPredJcb.Value = PredJcb
        sbPredJcb.Value = PredJcb
    End If
    
End Sub

Private Sub cbPenalty_Click()
    If (cbPenalty.Value) Then
        rePenalty.Enabled = True
    Else
        rePenalty.Enabled = False
    End If
End Sub

Private Sub cbParaStd_Click()
    If (cbParaStd.Value = True) Then
        tbParaStd.Enabled = True
        sbParaStd.Enabled = True
        tbParaStd.Value = sbParaStd.Value
    Else
        tbParaStd.Enabled = False
        sbParaStd.Enabled = False
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

Private Sub cbPredStd_Click()
    If (cbPredStd.Value = True) Then
        tbPredStd.Enabled = True
        sbPredStd.Enabled = True
        tbPredStd.Value = sbPredStd.Value
    Else
        tbPredStd.Enabled = False
        sbPredStd.Enabled = False
    End If
End Sub

Private Sub cbPredJcb_Click()
    If (cbPredJcb.Value = True) Then
        tbPredJcb.Enabled = True
        sbPredJcb.Enabled = True
        tbPredJcb.Value = sbPredJcb.Value
    Else
        tbPredJcb.Enabled = False
        sbPredJcb.Enabled = False
    End If
End Sub

Private Sub sbParaStd_Change()
    tbParaStd.Value = sbParaStd.Value
End Sub

Private Sub sbParaPtb_Change()
    tbParaPtb.Value = sbParaPtb.Value
End Sub

Private Sub sbPredJcb_Change()
    tbPredJcb.Value = sbPredJcb.Value
End Sub

Private Sub sbPredStd_Change()
    tbPredStd.Value = sbPredStd.Value
End Sub

Private Sub tbParaStd_Change()
    sbParaStd.Value = tbParaStd.Value
End Sub

Private Sub tbParaPtb_Change()
    sbParaPtb.Value = tbParaPtb.Value
End Sub

Private Sub tbPredStd_Change()
    sbPredStd.Value = tbPredStd.Value
End Sub

Private Sub tbPredJcb_Change()
    If (tbPredJcb.Value > 1) Then
        sbPredJcb.Value = tbPredJcb.Value
    Else
        MsgBox ("Offset from prediction range for output of Jacobian needs to be greater to avoid overwritting!")
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim ParaPtb, ParaStd, PredStd, PredJcb As Integer

    If (cbConstraint.Value) Then
        If (NameExists("NoParaConstraint", "Local")) Then
            Names("'" + ActiveSheet.Name + "'" + "!NoParaConstraint").Delete
        End If
    Else
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!NoParaConstraint", _
                  RefersTo:=1, Visible:=True
    End If
            
    If (cbPenalty.Value = True) Then
        If (rePenalty.Text <> "PenaltyCell") Then
            Names.Add Name:="'" + ActiveSheet.Name + "'" + "!PenaltyCell", _
                      RefersTo:="=" + rePenalty, Visible:=True
        End If
    Else
        If (NameExists("PenaltyCell", "Local")) Then
            Names("'" + ActiveSheet.Name + "'" + "!PenaltyCell").Delete
        End If
    End If
        
    If (cbParaStd.Value = True) Then
        ParaStd = tbParaStd.Value
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!OffsetParaStd", _
                  RefersTo:=ParaStd, Visible:=True
    Else
        If (NameExists("OffsetParaStd", "Local")) Then
            Names("'" + ActiveSheet.Name + "'" + "!OffsetParaStd").Delete
        End If
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
        
    If (cbPredStd.Value = True) Then
        PredStd = tbPredStd.Value
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!OffsetPredStd", _
                  RefersTo:=PredStd, Visible:=True
    Else
        If (NameExists("OffsetPredStd", "Local")) Then
            ActiveWorkbook.Names("'" + ActiveSheet.Name + "'" + "!OffsetPredStd").Delete
        End If
    End If

    Unload Me
End Sub


