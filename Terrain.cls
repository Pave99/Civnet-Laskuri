VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Terrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private name As String
Private def As Double
Private def_fe As String
Private food As Integer
Private terr_pic_id As Integer
Private prod As Integer
Private trade As Integer

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

Property Get foodvalue() As Integer
foodvalue = food
End Property

Property Let foodvalue(newfood As Integer)
food = newfood
End Property

Property Let terrpicid(newid As Integer)
terr_pic_id = newid
End Property

Property Get terrpicid() As Integer
terrpicid = terr_pic_id
End Property

Property Get prodvalue() As Integer
prodvalue = prod
End Property

Property Get tradevalue() As Integer
tradevalue = trade
End Property
