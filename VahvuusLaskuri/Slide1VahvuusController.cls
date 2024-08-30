VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Slide1VahvuusController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private targetSlide As Slide1
Private editorForm As VahvuusEditorForm
Private WithEvents henkilötKirjavahvuus As Henkilövahvuus
Attribute henkilötKirjavahvuus.VB_VarHelpID = -1
Private WithEvents henkilötRivivahvuus As Henkilövahvuus
Attribute henkilötRivivahvuus.VB_VarHelpID = -1
Private WithEvents henkilötPaikalla As Henkilövahvuus
Private WithEvents aseetPaikalla As Asiavahvuus
Attribute aseetPaikalla.VB_VarHelpID = -1
Private WithEvents aseetKirjavahvuus As Asiavahvuus
Attribute aseetKirjavahvuus.VB_VarHelpID = -1


Sub SetTargetSlideReference(newTargetSlide As Slide1)
    Set targetSlide = newTargetSlide
    
    InitializeAllHenkilövahvuudet
    InitializeAllAsevahvuudet
    InitializeEditorForm
End Sub

Private Sub InitializeAllHenkilövahvuudet()
    InitializeHenkilötKirjavahvuusFromTargetSlide
    InitializeHenkilötRivivahvuusFromTargetSlide
    InitializeHenkilötPaikallaFromTargetSlide
End Sub

Private Sub InitializeHenkilötKirjavahvuusFromTargetSlide()
    Dim newVahvuus As Henkilövahvuus
    Set newVahvuus = New Henkilövahvuus
    
    newVahvuus.SetUpseerit (SafeConversion.ToInt(targetSlide.GetUpseeritKirjavahvuus))
    newVahvuus.SetAliupseerit (SafeConversion.ToInt(targetSlide.GetAliupseeritKirjavahvuus))
    newVahvuus.SetMiehistö (SafeConversion.ToInt(targetSlide.GetMiehistöKirjavahvuus))
    
    Set henkilötKirjavahvuus = newVahvuus
End Sub

Private Sub InitializeHenkilötRivivahvuusFromTargetSlide()
    Dim newVahvuus As Henkilövahvuus
    Set newVahvuus = New Henkilövahvuus
    
    newVahvuus.SetUpseerit (SafeConversion.ToInt(targetSlide.GetUpseeritRivivahvuus))
    newVahvuus.SetAliupseerit (SafeConversion.ToInt(targetSlide.GetAliupseeritRivivahvuus))
    newVahvuus.SetMiehistö (SafeConversion.ToInt(targetSlide.GetMiehistöRivivahvuus))
    
    Set henkilötRivivahvuus = newVahvuus
End Sub

Private Sub InitializeHenkilötPaikallaFromTargetSlide()
    Dim newVahvuus As Henkilövahvuus
    Set newVahvuus = New Henkilövahvuus
    
    newVahvuus.SetUpseerit (SafeConversion.ToInt(targetSlide.GetUpseeritPaikalla))
    newVahvuus.SetAliupseerit (SafeConversion.ToInt(targetSlide.GetAliupseeritPaikalla))
    newVahvuus.SetMiehistö (SafeConversion.ToInt(targetSlide.GetMiehistöPaikalla))
    
    Set henkilötPaikalla = newVahvuus
End Sub

Private Sub InitializeAllAsevahvuudet()
    InitializeAseetPaikallaFromTargetSlide
    InitializeAseetKirjavahvuusFromTargetSlide
End Sub

Private Sub InitializeAseetPaikallaFromTargetSlide()
    Dim newVahvuus As Asiavahvuus
    Set newVahvuus = New Asiavahvuus
    
    newVahvuus.SetMäärä (SafeConversion.ToInt(targetSlide.GetAseetPaikalla))
    
    Set aseetPaikalla = newVahvuus
End Sub

Private Sub InitializeAseetKirjavahvuusFromTargetSlide()
    Dim newVahvuus As Asiavahvuus
    Set newVahvuus = New Asiavahvuus
    
    newVahvuus.SetMäärä (SafeConversion.ToInt(targetSlide.GetAseetKirjavahvuus))
    
    Set aseetKirjavahvuus = newVahvuus
End Sub

Private Sub InitializeEditorForm()
    Set editorForm = New VahvuusEditorForm
    
    editorForm.SetHenkilötKirjavahvuusReference henkilötKirjavahvuus
    editorForm.SetHenkilötRivivahvuusReference henkilötRivivahvuus
    editorForm.SetHenkilötPaikallaReference henkilötPaikalla
    
    editorForm.SetAseetPaikallaReference aseetPaikalla
    editorForm.SetAseetKirjavahvuusReference aseetKirjavahvuus
End Sub



'### EVENT HANDLERS ###

Private Sub henkilötKirjavahvuus_OnVahvuusChanged()
    UpdateTargetSlideHenkilötKirjavahvuus
End Sub

Private Sub UpdateTargetSlideHenkilötKirjavahvuus()
    targetSlide.SetUpseeritKirjavahvuus (SafeConversion.ToString(henkilötKirjavahvuus.GetUpseerit))
    targetSlide.SetAliupseeritKirjavahvuus (SafeConversion.ToString(henkilötKirjavahvuus.GetAliupseerit))
    targetSlide.SetMiehistöKirjavahvuus (SafeConversion.ToString(henkilötKirjavahvuus.GetMiehistö))
    targetSlide.SetKirjavahvuusYhteensä (SafeConversion.ToString(henkilötKirjavahvuus.GetHenkilötYhteensä))
End Sub


Private Sub henkilötRivivahvuus_OnVahvuusChanged()
    UpdateTargetSlideHenkilötRivivahvuus
End Sub


Private Sub UpdateTargetSlideHenkilötRivivahvuus()
    targetSlide.SetUpseeritRivivahvuus (SafeConversion.ToString(henkilötRivivahvuus.GetUpseerit))
    targetSlide.SetAliupseeritRivivahvuus (SafeConversion.ToString(henkilötRivivahvuus.GetAliupseerit))
    targetSlide.SetMiehistöRivivahvuus (SafeConversion.ToString(henkilötRivivahvuus.GetMiehistö))
    targetSlide.SetRivivahvuusYhteensä (SafeConversion.ToString(henkilötRivivahvuus.GetHenkilötYhteensä))
End Sub


Private Sub henkilötPaikalla_OnVahvuusChanged()
    UpdateTargetSlideHenkilötPaikalla
End Sub

Private Sub UpdateTargetSlideHenkilötPaikalla()
    targetSlide.SetUpseeritPaikalla (SafeConversion.ToString(henkilötPaikalla.GetUpseerit))
    targetSlide.SetAliupseeritPaikalla (SafeConversion.ToString(henkilötPaikalla.GetAliupseerit))
    targetSlide.SetMiehistöPaikalla (SafeConversion.ToString(henkilötPaikalla.GetMiehistö))
    targetSlide.SetPaikallaYhteensä (SafeConversion.ToString(henkilötPaikalla.GetHenkilötYhteensä))
End Sub


Private Sub aseetPaikalla_OnMääräChanged()
    UpdateTargetSlideAseetPaikalla
End Sub

Private Sub UpdateTargetSlideAseetPaikalla()
    targetSlide.SetAseetPaikalla (SafeConversion.ToString(aseetPaikalla.GetMäärä))
End Sub


Private Sub aseetKirjavahvuus_OnMääräChanged()
    UpdateTargetSlideAseetKirjavahvuus
End Sub

Private Sub UpdateTargetSlideAseetKirjavahvuus()
    targetSlide.SetAseetKirjavahvuus (SafeConversion.ToString(aseetKirjavahvuus.GetMäärä))
End Sub

'### INTERFACE FUNCTIONS ###

Sub ShowEditorForm()
    editorForm.Show
End Sub
