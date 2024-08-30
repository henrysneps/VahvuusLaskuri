VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Slide1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private vahvuusController As Slide1VahvuusController



Private Sub MuokkaaVahvuuksiaButton_Click()
    InitializeVahvuusController
    vahvuusController.ShowEditorForm
End Sub

Private Sub InitializeVahvuusController()
    Set vahvuusController = New Slide1VahvuusController
    vahvuusController.SetTargetSlideReference Me
End Sub



'### UI GETTERS/SETTERS FOR "KIRJAVAHVUUS" ###

Function GetUpseeritKirjavahvuus() As String
    GetUpseeritKirjavahvuus = Slide1.UpseeritKirjavahvuusLabel.Caption
End Function

Sub SetUpseeritKirjavahvuus(value As String)
    Slide1.UpseeritKirjavahvuusLabel.Caption = value
End Sub

Function GetAliupseeritKirjavahvuus() As String
    GetAliupseeritKirjavahvuus = Slide1.AliupseeritKirjavahvuusLabel.Caption
End Function

Sub SetAliupseeritKirjavahvuus(value As String)
    Slide1.AliupseeritKirjavahvuusLabel.Caption = value
End Sub

Function GetMiehistöKirjavahvuus() As String
    GetMiehistöKirjavahvuus = Slide1.MiehistöKirjavahvuusLabel.Caption
End Function

Sub SetMiehistöKirjavahvuus(value As String)
    Slide1.MiehistöKirjavahvuusLabel.Caption = value
End Sub

Sub SetKirjavahvuusYhteensä(value As String)
    Slide1.YhteensäKirjavahvuusLabel.Caption = value
End Sub



'### UI GETTERS/SETTERS FOR "RIVIVAHVUUS" ###

Function GetUpseeritRivivahvuus() As String
    GetUpseeritRivivahvuus = Slide1.UpseeritRivivahvuusLabel.Caption
End Function

Sub SetUpseeritRivivahvuus(value As String)
    Slide1.UpseeritRivivahvuusLabel.Caption = value
End Sub

Function GetAliupseeritRivivahvuus() As String
    GetAliupseeritRivivahvuus = Slide1.AliupseeritRivivahvuusLabel.Caption
End Function

Sub SetAliupseeritRivivahvuus(value As String)
    Slide1.AliupseeritRivivahvuusLabel.Caption = value
End Sub

Function GetMiehistöRivivahvuus() As String
    GetMiehistöRivivahvuus = Slide1.MiehistöRivivahvuusLabel.Caption
End Function

Sub SetMiehistöRivivahvuus(value As String)
    Slide1.MiehistöRivivahvuusLabel.Caption = value
End Sub

Sub SetRivivahvuusYhteensä(value As String)
    Slide1.YhteensäRivivahvuusLabel.Caption = value
End Sub



'### UI GETTERS/SETTERS FOR "PAIKALLA" ###

Function GetUpseeritPaikalla() As String
    GetUpseeritPaikalla = Slide1.UpseeritPaikallaLabel.Caption
End Function

Sub SetUpseeritPaikalla(value As String)
    Slide1.UpseeritPaikallaLabel.Caption = value
End Sub

Function GetAliupseeritPaikalla() As String
    GetAliupseeritPaikalla = Slide1.AliupseeritPaikallaLabel.Caption
End Function

Sub SetAliupseeritPaikalla(value As String)
    Slide1.AliupseeritPaikallaLabel.Caption = value
End Sub

Function GetMiehistöPaikalla() As String
    GetMiehistöPaikalla = Slide1.MiehistöPaikallaLabel.Caption
End Function

Sub SetMiehistöPaikalla(value As String)
    Slide1.MiehistöPaikallaLabel.Caption = value
End Sub

Sub SetPaikallaYhteensä(value As String)
    Slide1.YhteensäPaikallaLabel.Caption = value
End Sub



'### UI GETTERS/SETTERS FOR "ASEET" ###

Function GetAseetPaikalla() As String
    GetAseetPaikalla = Slide1.AseetPaikallaLabel.Caption
End Function

Sub SetAseetPaikalla(value As String)
    Slide1.AseetPaikallaLabel.Caption = value
End Sub

Function GetAseetKirjavahvuus() As String
    GetAseetKirjavahvuus = Slide1.AseetKirjavahvuusLabel.Caption
End Function

Sub SetAseetKirjavahvuus(value As String)
    Slide1.AseetKirjavahvuusLabel.Caption = value
End Sub
