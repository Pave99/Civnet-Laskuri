VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Terrain"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Option Explicit
Private name As String
Private def As Double
Private def_fe As String
Private food As Integer
Private food_irri As Integer



Property Get terrname() As String
terrname = name
End Property

Property Let terrname(newname As String)
name = newname
End Property

Property Get defbonus() As Double
defbonus = def
End Property

Property Let defbonus(newdef As Double)
def = newdef
End Property

Property Get deffront() As String
deffront = def_fe
End Property

Property Let deffront(newdef_fe As String)
def_fe = newdef_fe
End Property
