VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Unit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
