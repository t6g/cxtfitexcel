VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DlgResponse 
   Caption         =   "CXTFIT/Excel Response surface"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   OleObjectBlob   =   "DlgResponse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DlgResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

    If (NameExists("RSMeritCell", "Local")) Then
        reObj.Text = "RSMeritCell"
    End If
    
    If (NameExists("ParameterRange", "Local")) Then
        reParameter.Text = "ParameterRange"
    End If
    
    If (NameExists("RSVRange", "Local")) Then
        reRowInput.Text = "RSVRange"
    End If

    If (NameExists("RSHRange", "Local")) Then
        reColInput.Text = "RSHRange"
    End If

End Sub

Private Sub cmdCalculate_Click()
   
    If (reObj.Text <> "RSMeritCell") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!RSMeritCell", _
                  RefersTo:="=" + reObj, Visible:=True
    End If
    
    If (reParameter.Text <> "ParameterRange") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!ParameterRange", _
                  RefersTo:="=" + reParameter, Visible:=True
    End If
    
    If (reRowInput.Text <> "RSVRange") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!RSVRange", _
                  RefersTo:="=" + reRowInput, Visible:=True
    End If
    
    If (reColInput.Text <> "RSHRange") Then
        Names.Add Name:="'" + ActiveSheet.Name + "'" + "!RSHRange", _
                  RefersTo:="=" + reColInput, Visible:=True
    End If
    
    
    Call CalculateResponse
    
    Unload Me
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub





