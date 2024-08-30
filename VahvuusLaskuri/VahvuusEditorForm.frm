VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VahvuusEditorForm 
   Caption         =   "Muokkaa Vahvuuksia"
   ClientHeight    =   10800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10005
   OleObjectBlob   =   "VahvuusEditorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VahvuusEditorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents henkil�tKirjavahvuus As Henkil�vahvuus
Attribute henkil�tKirjavahvuus.VB_VarHelpID = -1
Private WithEvents henkil�tRivivahvuus As Henkil�vahvuus
Attribute henkil�tRivivahvuus.VB_VarHelpID = -1
Private WithEvents henkil�tPaikalla As Henkil�vahvuus
Attribute henkil�tPaikalla.VB_VarHelpID = -1

Private WithEvents aseetPaikalla As Asiavahvuus
Attribute aseetPaikalla.VB_VarHelpID = -1
Private WithEvents aseetKirjavahvuus As Asiavahvuus
Attribute aseetKirjavahvuus.VB_VarHelpID = -1



Private Sub UserForm_Activate()
    UpdateKirjavahvuusNytLabel
    UpdateRivivahvuusNytLabel
    UpdatePaikallaNytLabel
    
    UpdateAseetPaikallaNytLabel
    UpdateAseetKirjavahvuusNytLabel
End Sub

Sub SetHenkil�tKirjavahvuusReference(newVahvuus As Henkil�vahvuus)
    Set henkil�tKirjavahvuus = newVahvuus
End Sub

Sub SetHenkil�tRivivahvuusReference(newVahvuus As Henkil�vahvuus)
    Set henkil�tRivivahvuus = newVahvuus
End Sub

Sub SetHenkil�tPaikallaReference(newVahvuus As Henkil�vahvuus)
    Set henkil�tPaikalla = newVahvuus
End Sub


Sub SetAseetPaikallaReference(newVahvuus As Asiavahvuus)
    Set aseetPaikalla = newVahvuus
End Sub

Sub SetAseetKirjavahvuusReference(newVahvuus As Asiavahvuus)
    Set aseetKirjavahvuus = newVahvuus
End Sub


' ### HELPER FUNCTIONS ###

Private Function BuildVahvuusNytStringFrom(vahvuus As Henkil�vahvuus) As String
    BuildVahvuusNytStringFrom = "Vahvuus nyt: " + CStr(vahvuus.GetUpseerit) + " + " + CStr(vahvuus.GetAliupseerit) + " + " + CStr(vahvuus.GetMiehist�) + " = " + CStr(vahvuus.GetHenkil�tYhteens�)
End Function

Private Function BuildAsevahvuusNytStringFrom(vahvuus As Asiavahvuus) As String
    BuildAsevahvuusNytStringFrom = "Vahvuus nyt: " + CStr(vahvuus.GetM��r�)
End Function


'### KIRJAVAHVUUS ###

Private Function GetUpseeritKirjavahvuusTextBoxValueAsInt() As Integer
    GetUpseeritKirjavahvuusTextBoxValueAsInt = SafeConversion.ToInt(UpseeritKirjavahvuusTextBox.Text)
End Function

Private Function SetUpseeritKirjavahvuusTextBoxValue(value As Integer)
    UpseeritKirjavahvuusTextBox.Text = SafeConversion.ToString(value)
End Function

Private Function GetAliupseeritKirjavahvuusTextBoxValueAsInt() As Integer
    GetAliupseeritKirjavahvuusTextBoxValueAsInt = SafeConversion.ToInt(AliupseeritKirjavahvuusTextBox.Text)
End Function

Private Function SetAliupseeritKirjavahvuusTextBoxValue(value As Integer)
    AliupseeritKirjavahvuusTextBox.Text = SafeConversion.ToString(value)
End Function

Private Function GetMiehist�KirjavahvuusTextBoxValueAsInt() As Integer
    GetMiehist�KirjavahvuusTextBoxValueAsInt = SafeConversion.ToInt(Miehist�KirjavahvuusTextBox.Text)
End Function

Private Function SetMiehist�KirjavahvuusTextBoxValue(value As Integer)
    Miehist�KirjavahvuusTextBox.Text = SafeConversion.ToString(value)
End Function

Private Sub UpdateKirjavahvuusNytLabel()
    KirjavahvuusNytLabel.Caption = BuildVahvuusNytStringFrom(henkil�tKirjavahvuus)
End Sub

Private Sub ResetKirjavahvuusTextBoxes()
    SetUpseeritKirjavahvuusTextBoxValue (0)
    SetAliupseeritKirjavahvuusTextBoxValue (0)
    SetMiehist�KirjavahvuusTextBoxValue (0)
End Sub



Private Sub Lis��VahvuuteenKirjavahvuusButton_Click()
    henkil�tKirjavahvuus.SetUpseerit (henkil�tKirjavahvuus.GetUpseerit + GetUpseeritKirjavahvuusTextBoxValueAsInt)
    henkil�tKirjavahvuus.SetAliupseerit (henkil�tKirjavahvuus.GetAliupseerit + GetAliupseeritKirjavahvuusTextBoxValueAsInt)
    henkil�tKirjavahvuus.SetMiehist� (henkil�tKirjavahvuus.GetMiehist� + GetMiehist�KirjavahvuusTextBoxValueAsInt)
    
    UpdateKirjavahvuusNytLabel
    ResetKirjavahvuusTextBoxes
End Sub



Private Sub V�henn�VahvuudestaKirjavahvuusButton_Click()
    henkil�tKirjavahvuus.SetUpseerit (henkil�tKirjavahvuus.GetUpseerit - GetUpseeritKirjavahvuusTextBoxValueAsInt)
    henkil�tKirjavahvuus.SetAliupseerit (henkil�tKirjavahvuus.GetAliupseerit - GetAliupseeritKirjavahvuusTextBoxValueAsInt)
    henkil�tKirjavahvuus.SetMiehist� (henkil�tKirjavahvuus.GetMiehist� - GetMiehist�KirjavahvuusTextBoxValueAsInt)
    
    UpdateKirjavahvuusNytLabel
    ResetKirjavahvuusTextBoxes
End Sub

Private Sub NollaaKirjavahvuusButton_Click()
    henkil�tKirjavahvuus.SetUpseerit (0)
    henkil�tKirjavahvuus.SetAliupseerit (0)
    henkil�tKirjavahvuus.SetMiehist� (0)
    
    UpdateKirjavahvuusNytLabel
    ResetKirjavahvuusTextBoxes
End Sub

Private Sub AsetaKirjavahvuusButton_Click()
    henkil�tKirjavahvuus.SetUpseerit (GetUpseeritKirjavahvuusTextBoxValueAsInt)
    henkil�tKirjavahvuus.SetAliupseerit (GetAliupseeritKirjavahvuusTextBoxValueAsInt)
    henkil�tKirjavahvuus.SetMiehist� (GetMiehist�KirjavahvuusTextBoxValueAsInt)
    
    UpdateKirjavahvuusNytLabel
    ResetKirjavahvuusTextBoxes
End Sub

Private Sub SpinButtonUpseeritKirjavahvuus_SpinUp()
    SetUpseeritKirjavahvuusTextBoxValue (GetUpseeritKirjavahvuusTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonUpseeritKirjavahvuus_SpinDown()
        If (GetUpseeritKirjavahvuusTextBoxValueAsInt <= 0) Then
        SetUpseeritKirjavahvuusTextBoxValue (0)
    Else
        SetUpseeritKirjavahvuusTextBoxValue (GetUpseeritKirjavahvuusTextBoxValueAsInt - 1)
    End If
End Sub

Private Sub SpinButtonAliupseeritKirjavahvuus_SpinUp()
    SetAliupseeritKirjavahvuusTextBoxValue (GetAliupseeritKirjavahvuusTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonAliupseeritKirjavahvuus_SpinDown()
        If (GetAliupseeritKirjavahvuusTextBoxValueAsInt <= 0) Then
        SetAliupseeritKirjavahvuusTextBoxValue (0)
    Else
        SetAliupseeritKirjavahvuusTextBoxValue (GetAliupseeritKirjavahvuusTextBoxValueAsInt - 1)
    End If
End Sub

Private Sub SpinButtonMiehist�Kirjavahvuus_SpinUp()
    SetMiehist�KirjavahvuusTextBoxValue (GetMiehist�KirjavahvuusTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonMiehist�Kirjavahvuus_SpinDown()
        If (GetMiehist�KirjavahvuusTextBoxValueAsInt <= 0) Then
        SetMiehist�KirjavahvuusTextBoxValue (0)
    Else
        SetMiehist�KirjavahvuusTextBoxValue (GetMiehist�KirjavahvuusTextBoxValueAsInt - 1)
    End If
End Sub





'### RIVIVAHVUUS ###

Private Function GetUpseeritRivivahvuusTextBoxValueAsInt() As Integer
    GetUpseeritRivivahvuusTextBoxValueAsInt = SafeConversion.ToInt(UpseeritRivivahvuusTextBox.Text)
End Function

Private Function SetUpseeritRivivahvuusTextBoxValue(value As Integer)
    UpseeritRivivahvuusTextBox.Text = SafeConversion.ToString(value)
End Function

Private Function GetAliupseeritRivivahvuusTextBoxValueAsInt() As Integer
    GetAliupseeritRivivahvuusTextBoxValueAsInt = SafeConversion.ToInt(AliupseeritRivivahvuusTextBox.Text)
End Function

Private Function SetAliupseeritRivivahvuusTextBoxValue(value As Integer)
    AliupseeritRivivahvuusTextBox.Text = SafeConversion.ToString(value)
End Function

Private Function GetMiehist�RivivahvuusTextBoxValueAsInt() As Integer
    GetMiehist�RivivahvuusTextBoxValueAsInt = SafeConversion.ToInt(Miehist�RivivahvuusTextBox.Text)
End Function

Private Function SetMiehist�RivivahvuusTextBoxValue(value As Integer)
    Miehist�RivivahvuusTextBox.Text = SafeConversion.ToString(value)
End Function


Private Sub UpdateRivivahvuusNytLabel()
    RivivahvuusNytLabel.Caption = BuildVahvuusNytStringFrom(henkil�tRivivahvuus)
End Sub

Private Sub ResetRivivahvuusTextBoxes()
    SetUpseeritRivivahvuusTextBoxValue (0)
    SetAliupseeritRivivahvuusTextBoxValue (0)
    SetMiehist�RivivahvuusTextBoxValue (0)
End Sub


Private Sub Lis��VahvuuteenRivivahvuusButton_Click()
    henkil�tRivivahvuus.SetUpseerit (henkil�tRivivahvuus.GetUpseerit + GetUpseeritRivivahvuusTextBoxValueAsInt)
    henkil�tRivivahvuus.SetAliupseerit (henkil�tRivivahvuus.GetAliupseerit + GetAliupseeritRivivahvuusTextBoxValueAsInt)
    henkil�tRivivahvuus.SetMiehist� (henkil�tRivivahvuus.GetMiehist� + GetMiehist�RivivahvuusTextBoxValueAsInt)
    
    UpdateRivivahvuusNytLabel
    ResetRivivahvuusTextBoxes
End Sub

Private Sub V�henn�VahvuudestaRivivahvuusButton_Click()
    henkil�tRivivahvuus.SetUpseerit (henkil�tRivivahvuus.GetUpseerit - GetUpseeritRivivahvuusTextBoxValueAsInt)
    henkil�tRivivahvuus.SetAliupseerit (henkil�tRivivahvuus.GetAliupseerit - GetAliupseeritRivivahvuusTextBoxValueAsInt)
    henkil�tRivivahvuus.SetMiehist� (henkil�tRivivahvuus.GetMiehist� - GetMiehist�RivivahvuusTextBoxValueAsInt)
    
    UpdateRivivahvuusNytLabel
    ResetRivivahvuusTextBoxes
End Sub

Private Sub NollaaRivivahvuusButton_Click()
    henkil�tRivivahvuus.SetUpseerit (0)
    henkil�tRivivahvuus.SetAliupseerit (0)
    henkil�tRivivahvuus.SetMiehist� (0)
    
    UpdateRivivahvuusNytLabel
    ResetRivivahvuusTextBoxes
End Sub

Private Sub AsetaRivivahvuusButton_Click()
    henkil�tRivivahvuus.SetUpseerit (GetUpseeritRivivahvuusTextBoxValueAsInt)
    henkil�tRivivahvuus.SetAliupseerit (GetAliupseeritRivivahvuusTextBoxValueAsInt)
    henkil�tRivivahvuus.SetMiehist� (GetMiehist�RivivahvuusTextBoxValueAsInt)
    
    UpdateRivivahvuusNytLabel
    ResetRivivahvuusTextBoxes
End Sub

Private Sub SpinButtonUpseeritRivivahvuus_SpinUp()
    SetUpseeritRivivahvuusTextBoxValue (GetUpseeritRivivahvuusTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonUpseeritRivivahvuus_SpinDown()
        If (GetUpseeritRivivahvuusTextBoxValueAsInt <= 0) Then
        SetUpseeritRivivahvuusTextBoxValue (0)
    Else
        SetUpseeritRivivahvuusTextBoxValue (GetUpseeritRivivahvuusTextBoxValueAsInt - 1)
    End If
End Sub

Private Sub SpinButtonAliupseeritRivivahvuus_SpinUp()
    SetAliupseeritRivivahvuusTextBoxValue (GetAliupseeritRivivahvuusTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonAliupseeritRivivahvuus_SpinDown()
        If (GetAliupseeritRivivahvuusTextBoxValueAsInt <= 0) Then
        SetAliupseeritRivivahvuusTextBoxValue (0)
    Else
        SetAliupseeritRivivahvuusTextBoxValue (GetAliupseeritRivivahvuusTextBoxValueAsInt - 1)
    End If
End Sub

Private Sub SpinButtonMiehist�Rivivahvuus_SpinUp()
    SetMiehist�RivivahvuusTextBoxValue (GetMiehist�RivivahvuusTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonMiehist�Rivivahvuus_SpinDown()
        If (GetMiehist�RivivahvuusTextBoxValueAsInt <= 0) Then
        SetMiehist�RivivahvuusTextBoxValue (0)
    Else
        SetMiehist�RivivahvuusTextBoxValue (GetMiehist�RivivahvuusTextBoxValueAsInt - 1)
    End If
End Sub



'### PAIKALLA ###

Private Function GetUpseeritPaikallaTextBoxValueAsInt() As Integer
    GetUpseeritPaikallaTextBoxValueAsInt = SafeConversion.ToInt(UpseeritPaikallaTextBox.Text)
End Function

Private Function SetUpseeritPaikallaTextBoxValue(value As Integer)
    UpseeritPaikallaTextBox.Text = SafeConversion.ToString(value)
End Function

Private Function GetAliupseeritPaikallaTextBoxValueAsInt() As Integer
    GetAliupseeritPaikallaTextBoxValueAsInt = SafeConversion.ToInt(AliupseeritPaikallaTextBox.Text)
End Function

Private Function SetAliupseeritPaikallaTextBoxValue(value As Integer)
    AliupseeritPaikallaTextBox.Text = SafeConversion.ToString(value)
End Function

Private Function GetMiehist�PaikallaTextBoxValueAsInt() As Integer
    GetMiehist�PaikallaTextBoxValueAsInt = SafeConversion.ToInt(Miehist�PaikallaTextBox.Text)
End Function

Private Function SetMiehist�PaikallaTextBoxValue(value As Integer)
    Miehist�PaikallaTextBox.Text = SafeConversion.ToString(value)
End Function


Private Sub UpdatePaikallaNytLabel()
    PaikallaNytLabel.Caption = BuildVahvuusNytStringFrom(henkil�tPaikalla)
End Sub

Private Sub ResetPaikallaTextBoxes()
    SetUpseeritPaikallaTextBoxValue (0)
    SetAliupseeritPaikallaTextBoxValue (0)
    SetMiehist�PaikallaTextBoxValue (0)
End Sub


Private Sub Lis��VahvuuteenPaikallaButton_Click()
    henkil�tPaikalla.SetUpseerit (henkil�tPaikalla.GetUpseerit + GetUpseeritPaikallaTextBoxValueAsInt)
    henkil�tPaikalla.SetAliupseerit (henkil�tPaikalla.GetAliupseerit + GetAliupseeritPaikallaTextBoxValueAsInt)
    henkil�tPaikalla.SetMiehist� (henkil�tPaikalla.GetMiehist� + GetMiehist�PaikallaTextBoxValueAsInt)
    
    UpdatePaikallaNytLabel
    ResetPaikallaTextBoxes
End Sub

Private Sub V�henn�VahvuudestaPaikallaButton_Click()
    henkil�tPaikalla.SetUpseerit (henkil�tPaikalla.GetUpseerit - GetUpseeritPaikallaTextBoxValueAsInt)
    henkil�tPaikalla.SetAliupseerit (henkil�tPaikalla.GetAliupseerit - GetAliupseeritPaikallaTextBoxValueAsInt)
    henkil�tPaikalla.SetMiehist� (henkil�tPaikalla.GetMiehist� - GetMiehist�PaikallaTextBoxValueAsInt)
    
    UpdatePaikallaNytLabel
    ResetPaikallaTextBoxes
End Sub

Private Sub NollaaPaikallaButton_Click()
    henkil�tPaikalla.SetUpseerit (0)
    henkil�tPaikalla.SetAliupseerit (0)
    henkil�tPaikalla.SetMiehist� (0)
    
    UpdatePaikallaNytLabel
    ResetPaikallaTextBoxes
End Sub

Private Sub AsetaPaikallaButton_Click()
    henkil�tPaikalla.SetUpseerit (GetUpseeritPaikallaTextBoxValueAsInt)
    henkil�tPaikalla.SetAliupseerit (GetAliupseeritPaikallaTextBoxValueAsInt)
    henkil�tPaikalla.SetMiehist� (GetMiehist�PaikallaTextBoxValueAsInt)
    
    UpdatePaikallaNytLabel
    ResetPaikallaTextBoxes
End Sub

Private Sub SpinButtonUpseeritPaikalla_SpinUp()
    SetUpseeritPaikallaTextBoxValue (GetUpseeritPaikallaTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonUpseeritPaikalla_SpinDown()
        If (GetUpseeritPaikallaTextBoxValueAsInt <= 0) Then
        SetUpseeritPaikallaTextBoxValue (0)
    Else
        SetUpseeritPaikallaTextBoxValue (GetUpseeritPaikallaTextBoxValueAsInt - 1)
    End If
End Sub

Private Sub SpinButtonAliupseeritPaikalla_SpinUp()
    SetAliupseeritPaikallaTextBoxValue (GetAliupseeritPaikallaTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonAliupseeritPaikalla_SpinDown()
        If (GetAliupseeritPaikallaTextBoxValueAsInt <= 0) Then
        SetAliupseeritPaikallaTextBoxValue (0)
    Else
        SetAliupseeritPaikallaTextBoxValue (GetAliupseeritPaikallaTextBoxValueAsInt - 1)
    End If
End Sub

Private Sub SpinButtonMiehist�Paikalla_SpinUp()
    SetMiehist�PaikallaTextBoxValue (GetMiehist�PaikallaTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonMiehist�Paikalla_SpinDown()
        If (GetMiehist�PaikallaTextBoxValueAsInt <= 0) Then
        SetMiehist�PaikallaTextBoxValue (0)
    Else
        SetMiehist�PaikallaTextBoxValue (GetMiehist�PaikallaTextBoxValueAsInt - 1)
    End If
End Sub



'### ASEET PAIKALLA ###

Private Function GetAseetPaikallaTextBoxValueAsInt() As Integer
    GetAseetPaikallaTextBoxValueAsInt = SafeConversion.ToInt(AseetPaikallaTextBox.Text)
End Function

Private Function SetAseetPaikallaTextBoxValue(value As Integer)
    AseetPaikallaTextBox.Text = SafeConversion.ToString(value)
End Function

Private Sub UpdateAseetPaikallaNytLabel()
    AseetPaikallaNytLabel.Caption = BuildAsevahvuusNytStringFrom(aseetPaikalla)
End Sub

Private Sub ResetAseetPaikallaTextBoxes()
    SetAseetPaikallaTextBoxValue (0)
End Sub


Private Sub Lis��VahvuuteenAseetPaikallaButton_Click()
    aseetPaikalla.SetM��r� (aseetPaikalla.GetM��r� + GetAseetPaikallaTextBoxValueAsInt)
    
    UpdateAseetPaikallaNytLabel
    ResetAseetPaikallaTextBoxes
End Sub

Private Sub V�henn�AseetVahvuudestaPaikallaButton_Click()
    aseetPaikalla.SetM��r� (aseetPaikalla.GetM��r� - GetAseetPaikallaTextBoxValueAsInt)
    
    UpdateAseetPaikallaNytLabel
    ResetAseetPaikallaTextBoxes
End Sub

Private Sub NollaaAseetPaikallaButton_Click()
    aseetPaikalla.SetM��r� (0)
    
    UpdateAseetPaikallaNytLabel
    ResetAseetPaikallaTextBoxes
End Sub

Private Sub AsetaAseetPaikallaButton_Click()
    aseetPaikalla.SetM��r� (GetAseetPaikallaTextBoxValueAsInt)
    
    UpdateAseetPaikallaNytLabel
    ResetAseetPaikallaTextBoxes
End Sub

Private Sub SpinButtonAseetPaikalla_SpinUp()
    SetAseetPaikallaTextBoxValue (GetAseetPaikallaTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonAseetPaikalla_SpinDown()
        If (GetAseetPaikallaTextBoxValueAsInt <= 0) Then
        SetAseetPaikallaTextBoxValue (0)
    Else
        SetAseetPaikallaTextBoxValue (GetAseetPaikallaTextBoxValueAsInt - 1)
    End If
End Sub



'### ASEET KIRJAVAHVUUS ###

Private Function GetAseetKirjavahvuusTextBoxValueAsInt() As Integer
    GetAseetKirjavahvuusTextBoxValueAsInt = SafeConversion.ToInt(AseetKirjavahvuusTextBox.Text)
End Function

Private Function SetAseetKirjavahvuusTextBoxValue(value As Integer)
    AseetKirjavahvuusTextBox.Text = SafeConversion.ToString(value)
End Function


Private Sub UpdateAseetKirjavahvuusNytLabel()
    AseetKirjavahvuusNytLabel.Caption = BuildAsevahvuusNytStringFrom(aseetKirjavahvuus)
End Sub

Private Sub ResetAseetKirjavahvuusTextBoxes()
    SetAseetKirjavahvuusTextBoxValue (0)
End Sub


Private Sub Lis��VahvuuteenAseetKirjavahvuusButton_Click()
    aseetKirjavahvuus.SetM��r� (aseetKirjavahvuus.GetM��r� + GetAseetKirjavahvuusTextBoxValueAsInt)
    
    UpdateAseetKirjavahvuusNytLabel
    ResetAseetKirjavahvuusTextBoxes
End Sub

Private Sub V�henn�AseetVahvuudestaKirjavahvuusButton_Click()
    aseetKirjavahvuus.SetM��r� (aseetKirjavahvuus.GetM��r� - GetAseetKirjavahvuusTextBoxValueAsInt)
    
    UpdateAseetKirjavahvuusNytLabel
    ResetAseetKirjavahvuusTextBoxes
End Sub

Private Sub NollaaAseetKirjavahvuusButton_Click()
    aseetKirjavahvuus.SetM��r� (0)
    
    UpdateAseetKirjavahvuusNytLabel
    ResetAseetKirjavahvuusTextBoxes
End Sub

Private Sub AsetaAseetKirjavahvuusButton_Click()
    aseetKirjavahvuus.SetM��r� (GetAseetKirjavahvuusTextBoxValueAsInt)
    
    UpdateAseetKirjavahvuusNytLabel
    ResetAseetKirjavahvuusTextBoxes
End Sub

Private Sub SpinButtonAseetKirjavahvuus_SpinUp()
    SetAseetKirjavahvuusTextBoxValue (GetAseetKirjavahvuusTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonAseetKirjavahvuus_SpinDown()
        If (GetAseetKirjavahvuusTextBoxValueAsInt <= 0) Then
        SetAseetKirjavahvuusTextBoxValue (0)
    Else
        SetAseetKirjavahvuusTextBoxValue (GetAseetKirjavahvuusTextBoxValueAsInt - 1)
    End If
End Sub
