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
Private WithEvents henkilötKirjavahvuus As Henkilövahvuus
Attribute henkilötKirjavahvuus.VB_VarHelpID = -1
Private WithEvents henkilötRivivahvuus As Henkilövahvuus
Attribute henkilötRivivahvuus.VB_VarHelpID = -1
Private WithEvents henkilötPaikalla As Henkilövahvuus
Attribute henkilötPaikalla.VB_VarHelpID = -1

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

Sub SetHenkilötKirjavahvuusReference(newVahvuus As Henkilövahvuus)
    Set henkilötKirjavahvuus = newVahvuus
End Sub

Sub SetHenkilötRivivahvuusReference(newVahvuus As Henkilövahvuus)
    Set henkilötRivivahvuus = newVahvuus
End Sub

Sub SetHenkilötPaikallaReference(newVahvuus As Henkilövahvuus)
    Set henkilötPaikalla = newVahvuus
End Sub


Sub SetAseetPaikallaReference(newVahvuus As Asiavahvuus)
    Set aseetPaikalla = newVahvuus
End Sub

Sub SetAseetKirjavahvuusReference(newVahvuus As Asiavahvuus)
    Set aseetKirjavahvuus = newVahvuus
End Sub


' ### HELPER FUNCTIONS ###

Private Function BuildVahvuusNytStringFrom(vahvuus As Henkilövahvuus) As String
    BuildVahvuusNytStringFrom = "Vahvuus nyt: " + CStr(vahvuus.GetUpseerit) + " + " + CStr(vahvuus.GetAliupseerit) + " + " + CStr(vahvuus.GetMiehistö) + " = " + CStr(vahvuus.GetHenkilötYhteensä)
End Function

Private Function BuildAsevahvuusNytStringFrom(vahvuus As Asiavahvuus) As String
    BuildAsevahvuusNytStringFrom = "Vahvuus nyt: " + CStr(vahvuus.GetMäärä)
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

Private Function GetMiehistöKirjavahvuusTextBoxValueAsInt() As Integer
    GetMiehistöKirjavahvuusTextBoxValueAsInt = SafeConversion.ToInt(MiehistöKirjavahvuusTextBox.Text)
End Function

Private Function SetMiehistöKirjavahvuusTextBoxValue(value As Integer)
    MiehistöKirjavahvuusTextBox.Text = SafeConversion.ToString(value)
End Function

Private Sub UpdateKirjavahvuusNytLabel()
    KirjavahvuusNytLabel.Caption = BuildVahvuusNytStringFrom(henkilötKirjavahvuus)
End Sub

Private Sub ResetKirjavahvuusTextBoxes()
    SetUpseeritKirjavahvuusTextBoxValue (0)
    SetAliupseeritKirjavahvuusTextBoxValue (0)
    SetMiehistöKirjavahvuusTextBoxValue (0)
End Sub



Private Sub LisääVahvuuteenKirjavahvuusButton_Click()
    henkilötKirjavahvuus.SetUpseerit (henkilötKirjavahvuus.GetUpseerit + GetUpseeritKirjavahvuusTextBoxValueAsInt)
    henkilötKirjavahvuus.SetAliupseerit (henkilötKirjavahvuus.GetAliupseerit + GetAliupseeritKirjavahvuusTextBoxValueAsInt)
    henkilötKirjavahvuus.SetMiehistö (henkilötKirjavahvuus.GetMiehistö + GetMiehistöKirjavahvuusTextBoxValueAsInt)
    
    UpdateKirjavahvuusNytLabel
    ResetKirjavahvuusTextBoxes
End Sub



Private Sub VähennäVahvuudestaKirjavahvuusButton_Click()
    henkilötKirjavahvuus.SetUpseerit (henkilötKirjavahvuus.GetUpseerit - GetUpseeritKirjavahvuusTextBoxValueAsInt)
    henkilötKirjavahvuus.SetAliupseerit (henkilötKirjavahvuus.GetAliupseerit - GetAliupseeritKirjavahvuusTextBoxValueAsInt)
    henkilötKirjavahvuus.SetMiehistö (henkilötKirjavahvuus.GetMiehistö - GetMiehistöKirjavahvuusTextBoxValueAsInt)
    
    UpdateKirjavahvuusNytLabel
    ResetKirjavahvuusTextBoxes
End Sub

Private Sub NollaaKirjavahvuusButton_Click()
    henkilötKirjavahvuus.SetUpseerit (0)
    henkilötKirjavahvuus.SetAliupseerit (0)
    henkilötKirjavahvuus.SetMiehistö (0)
    
    UpdateKirjavahvuusNytLabel
    ResetKirjavahvuusTextBoxes
End Sub

Private Sub AsetaKirjavahvuusButton_Click()
    henkilötKirjavahvuus.SetUpseerit (GetUpseeritKirjavahvuusTextBoxValueAsInt)
    henkilötKirjavahvuus.SetAliupseerit (GetAliupseeritKirjavahvuusTextBoxValueAsInt)
    henkilötKirjavahvuus.SetMiehistö (GetMiehistöKirjavahvuusTextBoxValueAsInt)
    
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

Private Sub SpinButtonMiehistöKirjavahvuus_SpinUp()
    SetMiehistöKirjavahvuusTextBoxValue (GetMiehistöKirjavahvuusTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonMiehistöKirjavahvuus_SpinDown()
        If (GetMiehistöKirjavahvuusTextBoxValueAsInt <= 0) Then
        SetMiehistöKirjavahvuusTextBoxValue (0)
    Else
        SetMiehistöKirjavahvuusTextBoxValue (GetMiehistöKirjavahvuusTextBoxValueAsInt - 1)
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

Private Function GetMiehistöRivivahvuusTextBoxValueAsInt() As Integer
    GetMiehistöRivivahvuusTextBoxValueAsInt = SafeConversion.ToInt(MiehistöRivivahvuusTextBox.Text)
End Function

Private Function SetMiehistöRivivahvuusTextBoxValue(value As Integer)
    MiehistöRivivahvuusTextBox.Text = SafeConversion.ToString(value)
End Function


Private Sub UpdateRivivahvuusNytLabel()
    RivivahvuusNytLabel.Caption = BuildVahvuusNytStringFrom(henkilötRivivahvuus)
End Sub

Private Sub ResetRivivahvuusTextBoxes()
    SetUpseeritRivivahvuusTextBoxValue (0)
    SetAliupseeritRivivahvuusTextBoxValue (0)
    SetMiehistöRivivahvuusTextBoxValue (0)
End Sub


Private Sub LisääVahvuuteenRivivahvuusButton_Click()
    henkilötRivivahvuus.SetUpseerit (henkilötRivivahvuus.GetUpseerit + GetUpseeritRivivahvuusTextBoxValueAsInt)
    henkilötRivivahvuus.SetAliupseerit (henkilötRivivahvuus.GetAliupseerit + GetAliupseeritRivivahvuusTextBoxValueAsInt)
    henkilötRivivahvuus.SetMiehistö (henkilötRivivahvuus.GetMiehistö + GetMiehistöRivivahvuusTextBoxValueAsInt)
    
    UpdateRivivahvuusNytLabel
    ResetRivivahvuusTextBoxes
End Sub

Private Sub VähennäVahvuudestaRivivahvuusButton_Click()
    henkilötRivivahvuus.SetUpseerit (henkilötRivivahvuus.GetUpseerit - GetUpseeritRivivahvuusTextBoxValueAsInt)
    henkilötRivivahvuus.SetAliupseerit (henkilötRivivahvuus.GetAliupseerit - GetAliupseeritRivivahvuusTextBoxValueAsInt)
    henkilötRivivahvuus.SetMiehistö (henkilötRivivahvuus.GetMiehistö - GetMiehistöRivivahvuusTextBoxValueAsInt)
    
    UpdateRivivahvuusNytLabel
    ResetRivivahvuusTextBoxes
End Sub

Private Sub NollaaRivivahvuusButton_Click()
    henkilötRivivahvuus.SetUpseerit (0)
    henkilötRivivahvuus.SetAliupseerit (0)
    henkilötRivivahvuus.SetMiehistö (0)
    
    UpdateRivivahvuusNytLabel
    ResetRivivahvuusTextBoxes
End Sub

Private Sub AsetaRivivahvuusButton_Click()
    henkilötRivivahvuus.SetUpseerit (GetUpseeritRivivahvuusTextBoxValueAsInt)
    henkilötRivivahvuus.SetAliupseerit (GetAliupseeritRivivahvuusTextBoxValueAsInt)
    henkilötRivivahvuus.SetMiehistö (GetMiehistöRivivahvuusTextBoxValueAsInt)
    
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

Private Sub SpinButtonMiehistöRivivahvuus_SpinUp()
    SetMiehistöRivivahvuusTextBoxValue (GetMiehistöRivivahvuusTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonMiehistöRivivahvuus_SpinDown()
        If (GetMiehistöRivivahvuusTextBoxValueAsInt <= 0) Then
        SetMiehistöRivivahvuusTextBoxValue (0)
    Else
        SetMiehistöRivivahvuusTextBoxValue (GetMiehistöRivivahvuusTextBoxValueAsInt - 1)
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

Private Function GetMiehistöPaikallaTextBoxValueAsInt() As Integer
    GetMiehistöPaikallaTextBoxValueAsInt = SafeConversion.ToInt(MiehistöPaikallaTextBox.Text)
End Function

Private Function SetMiehistöPaikallaTextBoxValue(value As Integer)
    MiehistöPaikallaTextBox.Text = SafeConversion.ToString(value)
End Function


Private Sub UpdatePaikallaNytLabel()
    PaikallaNytLabel.Caption = BuildVahvuusNytStringFrom(henkilötPaikalla)
End Sub

Private Sub ResetPaikallaTextBoxes()
    SetUpseeritPaikallaTextBoxValue (0)
    SetAliupseeritPaikallaTextBoxValue (0)
    SetMiehistöPaikallaTextBoxValue (0)
End Sub


Private Sub LisääVahvuuteenPaikallaButton_Click()
    henkilötPaikalla.SetUpseerit (henkilötPaikalla.GetUpseerit + GetUpseeritPaikallaTextBoxValueAsInt)
    henkilötPaikalla.SetAliupseerit (henkilötPaikalla.GetAliupseerit + GetAliupseeritPaikallaTextBoxValueAsInt)
    henkilötPaikalla.SetMiehistö (henkilötPaikalla.GetMiehistö + GetMiehistöPaikallaTextBoxValueAsInt)
    
    UpdatePaikallaNytLabel
    ResetPaikallaTextBoxes
End Sub

Private Sub VähennäVahvuudestaPaikallaButton_Click()
    henkilötPaikalla.SetUpseerit (henkilötPaikalla.GetUpseerit - GetUpseeritPaikallaTextBoxValueAsInt)
    henkilötPaikalla.SetAliupseerit (henkilötPaikalla.GetAliupseerit - GetAliupseeritPaikallaTextBoxValueAsInt)
    henkilötPaikalla.SetMiehistö (henkilötPaikalla.GetMiehistö - GetMiehistöPaikallaTextBoxValueAsInt)
    
    UpdatePaikallaNytLabel
    ResetPaikallaTextBoxes
End Sub

Private Sub NollaaPaikallaButton_Click()
    henkilötPaikalla.SetUpseerit (0)
    henkilötPaikalla.SetAliupseerit (0)
    henkilötPaikalla.SetMiehistö (0)
    
    UpdatePaikallaNytLabel
    ResetPaikallaTextBoxes
End Sub

Private Sub AsetaPaikallaButton_Click()
    henkilötPaikalla.SetUpseerit (GetUpseeritPaikallaTextBoxValueAsInt)
    henkilötPaikalla.SetAliupseerit (GetAliupseeritPaikallaTextBoxValueAsInt)
    henkilötPaikalla.SetMiehistö (GetMiehistöPaikallaTextBoxValueAsInt)
    
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

Private Sub SpinButtonMiehistöPaikalla_SpinUp()
    SetMiehistöPaikallaTextBoxValue (GetMiehistöPaikallaTextBoxValueAsInt + 1)
End Sub

Private Sub SpinButtonMiehistöPaikalla_SpinDown()
        If (GetMiehistöPaikallaTextBoxValueAsInt <= 0) Then
        SetMiehistöPaikallaTextBoxValue (0)
    Else
        SetMiehistöPaikallaTextBoxValue (GetMiehistöPaikallaTextBoxValueAsInt - 1)
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


Private Sub LisääVahvuuteenAseetPaikallaButton_Click()
    aseetPaikalla.SetMäärä (aseetPaikalla.GetMäärä + GetAseetPaikallaTextBoxValueAsInt)
    
    UpdateAseetPaikallaNytLabel
    ResetAseetPaikallaTextBoxes
End Sub

Private Sub VähennäAseetVahvuudestaPaikallaButton_Click()
    aseetPaikalla.SetMäärä (aseetPaikalla.GetMäärä - GetAseetPaikallaTextBoxValueAsInt)
    
    UpdateAseetPaikallaNytLabel
    ResetAseetPaikallaTextBoxes
End Sub

Private Sub NollaaAseetPaikallaButton_Click()
    aseetPaikalla.SetMäärä (0)
    
    UpdateAseetPaikallaNytLabel
    ResetAseetPaikallaTextBoxes
End Sub

Private Sub AsetaAseetPaikallaButton_Click()
    aseetPaikalla.SetMäärä (GetAseetPaikallaTextBoxValueAsInt)
    
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


Private Sub LisääVahvuuteenAseetKirjavahvuusButton_Click()
    aseetKirjavahvuus.SetMäärä (aseetKirjavahvuus.GetMäärä + GetAseetKirjavahvuusTextBoxValueAsInt)
    
    UpdateAseetKirjavahvuusNytLabel
    ResetAseetKirjavahvuusTextBoxes
End Sub

Private Sub VähennäAseetVahvuudestaKirjavahvuusButton_Click()
    aseetKirjavahvuus.SetMäärä (aseetKirjavahvuus.GetMäärä - GetAseetKirjavahvuusTextBoxValueAsInt)
    
    UpdateAseetKirjavahvuusNytLabel
    ResetAseetKirjavahvuusTextBoxes
End Sub

Private Sub NollaaAseetKirjavahvuusButton_Click()
    aseetKirjavahvuus.SetMäärä (0)
    
    UpdateAseetKirjavahvuusNytLabel
    ResetAseetKirjavahvuusTextBoxes
End Sub

Private Sub AsetaAseetKirjavahvuusButton_Click()
    aseetKirjavahvuus.SetMäärä (GetAseetKirjavahvuusTextBoxValueAsInt)
    
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
