VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Asiavahvuus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event OnM‰‰r‰Changed()

Private m‰‰r‰ As Integer

Private Sub Class_Initialize()
    m‰‰r‰ = 0
End Sub

Function GetM‰‰r‰()
    GetM‰‰r‰ = m‰‰r‰
End Function

Function SetM‰‰r‰(value As Integer)
    m‰‰r‰ = value
    RaiseEvent OnM‰‰r‰Changed
End Function
