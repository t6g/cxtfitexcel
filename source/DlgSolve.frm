VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DlgSolve 
   Caption         =   "CXTFIT/Excel Solve"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   OleObjectBlob   =   "DlgSolve.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DlgSolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdOptions_Click()
    DlgSolveOption.Show
End Sub

Private Sub UserForm_Initialize()
    If (NameExists("ObjFuncCell", "Local")) Then
        reObj.Text = "ObjFuncCell"
    End If
    
    If (NameExists("ParameterRange", "Local")) Then
        reParameter.Text = "ParameterRange"
    End If
    
    If (NameExists("PredictionRange", "Local")) Then
        rePrediction.Text = "PredictionRange"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSolve_Click()
    If (reObj = "") Then
        MsgBox ("Objective function cell not selected!")
        Exit Sub
    End If
        
    If (reParameter = "") Then
        MsgBox ("Parameter range not selected!")
        Exit Sub
    End If
    
    If (rePrediction = "") Then
        MsgBox ("Prediction range not selected!")
        Exit Sub
    End If
    
    If (reObj.Text <> "ObjFuncCell") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!ObjFuncCell", _
                  RefersTo:="=" + reObj, Visible:=True
    End If
    
    If (reParameter.Text <> "ParameterRange") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!ParameterRange", _
                  RefersTo:="=" + reParameter, Visible:=True
    End If
    
    If (rePrediction.Text <> "PredictionRange") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!PredictionRange", _
                  RefersTo:="=" + rePrediction, Visible:=True
    End If
    
    Call Solve
    
    Unload Me
End Sub
