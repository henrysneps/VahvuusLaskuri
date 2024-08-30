VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Henkilövahvuus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event OnVahvuusChanged()

Private upseerit As Integer
Private aliupseerit As Integer
Private miehistö As Integer
Private henkilötYhteensä As Integer



Private Sub Class_Initialize()
    InitializeAllVahvuudet
End Sub

Private Sub InitializeAllVahvuudet()
    upseerit = 0
    aliupseerit = 0
    miehistö = 0
    henkilötYhteensä = 0
End Sub



'### GETTERS/SETTERS ###

Private Sub UpdateHenkilötYhteensä()
    henkilötYhteensä = upseerit + aliupseerit + miehistö
    RaiseEvent OnVahvuusChanged
End Sub

Public Function GetHenkilötYhteensä()
    GetHenkilötYhteensä = henkilötYhteensä
End Function

Public Function GetUpseerit() As Integer
    GetUpseerit = upseerit
End Function

Public Function SetUpseerit(value As Integer)
    upseerit = value
    UpdateHenkilötYhteensä
End Function

Public Function GetAliupseerit() As Integer
    GetAliupseerit = aliupseerit
End Function

Public Function SetAliupseerit(value As Integer)
    aliupseerit = value
    UpdateHenkilötYhteensä
End Function

Public Function GetMiehistö() As Integer
    GetMiehistö = miehistö
End Function

Public Function SetMiehistö(value As Integer)
    miehistö = value
    UpdateHenkilötYhteensä
End Function

