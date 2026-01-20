VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Option Explicit
Private name As String
Private att As Integer
Private def As Integer

Property Get Unitname() As String
Unitname = name
End Property
Property Let Unitname(newname As String)
name = newname
End Property
Property Get attack() As Integer
attack = att
End Property
Property Let attack(newattack As Integer)
att = newattack
End Property
Property Get defend() As Integer
defend = def
End Property
Property Let defend(newdefend As Integer)
def = newdefend
End Property
