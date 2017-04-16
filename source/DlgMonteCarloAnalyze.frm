VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DlgMonteCarloAnalyze 
   Caption         =   "CXTFIT/Excel Monte Carlo Analysis"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   OleObjectBlob   =   "DlgMonteCarloAnalyze.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DlgMonteCarloAnalyze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub UserForm_Initialize()

    If (NameExists("MCMeritCell", "Local")) Then
        reObj.Text = "MCMeritCell"
    End If
    
    If (NameExists("ParameterRange", "Local")) Then
        reParameter.Text = "ParameterRange"
    End If
    
    If (NameExists("MCParaInputRange", "Local")) Then
        reInput.Text = "MCParaInputRange"
    End If

    If (NameExists("MCCurrentCell", "Local")) Then
        reProgress.Text = "MCCurrentCell"
    End If

End Sub

Private Sub cmdCalculate_Click()
    If (reObj = "") Then
        MsgBox ("Merit function cell not selected!")
        Exit Sub
    End If
        
    If (reParameter = "") Then
        MsgBox ("Parameter range not selected!")
        Exit Sub
    End If
    
    If (reInput = "") Then
        MsgBox ("Parameter input range not selected!")
        Exit Sub
    End If
    
    If (reObj.Text <> "MCMeritCell") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!MCMeritCell", _
                  RefersTo:="=" + reObj, Visible:=True
    End If
    
    If (reParameter.Text <> "ParameterRange") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!ParameterRange", _
                  RefersTo:="=" + reParameter, Visible:=True
    End If
    
    If (reInput.Text <> "MCParaInputRange") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!MCParaInputRange", _
                  RefersTo:="=" + reInput, Visible:=True
    End If
    
    If (reProgress <> "MCCurrentCell") Then
        If (reProgress <> "") Then
            Names.Add Name:="'" + ActiveSheet.Name + "'" + "!MCCurrentCell", _
                      RefersTo:="=" + reProgress, Visible:=True
        Else
            If (NameExists("MCCurrentCell", "Local")) Then
                Names("'" + ActiveSheet.Name + "'" + "!MCCurrentCell").Delete
            End If
        End If
    End If
        
    Call MonteCarloAnalyze
    
    Unload Me
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub



