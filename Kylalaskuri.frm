VERSION 5.00
Begin VB.Form Kylalaskuri 
   Caption         =   "Kylä-laskuri"
   ClientHeight    =   9840
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Kylalaskuri"
   ScaleHeight     =   9840
   ScaleMode       =   0  'User
   ScaleWidth      =   6467.336
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6453
      Begin VB.CommandButton laske 
         Caption         =   "Laske"
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CheckBox Check21 
         Caption         =   "Check21"
         Height          =   255
         Left            =   4095
         TabIndex        =   21
         Top             =   5140
         Width           =   255
      End
      Begin VB.CheckBox Check20 
         Caption         =   "Check20"
         Height          =   255
         Left            =   3015
         TabIndex        =   20
         Top             =   5140
         Width           =   255
      End
      Begin VB.CheckBox Check19 
         Caption         =   "Check19"
         Height          =   255
         Left            =   1945
         TabIndex        =   19
         Top             =   5140
         Width           =   255
      End
      Begin VB.CheckBox Check18 
         Caption         =   "Check18"
         Height          =   255
         Left            =   5175
         TabIndex        =   18
         Top             =   3940
         Width           =   255
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Check17"
         Height          =   255
         Left            =   4095
         TabIndex        =   17
         Top             =   3940
         Width           =   255
      End
      Begin VB.CheckBox Check16 
         Caption         =   "Check16"
         Height          =   255
         Left            =   3015
         TabIndex        =   16
         Top             =   3940
         Width           =   255
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Check15"
         Height          =   255
         Left            =   1935
         TabIndex        =   15
         Top             =   3940
         Width           =   255
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Check14"
         Height          =   255
         Left            =   975
         TabIndex        =   14
         Top             =   3940
         Width           =   255
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Check13"
         Height          =   255
         Left            =   5175
         TabIndex        =   13
         Top             =   2740
         Width           =   255
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Check12"
         Height          =   255
         Left            =   4095
         TabIndex        =   12
         Top             =   2740
         Width           =   255
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Check11"
         Height          =   255
         Left            =   3015
         TabIndex        =   11
         Top             =   2740
         Width           =   255
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Check10"
         Height          =   255
         Left            =   1935
         TabIndex        =   10
         Top             =   2740
         Width           =   255
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Check9"
         Height          =   255
         Left            =   975
         TabIndex        =   9
         Top             =   2740
         Width           =   255
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check8"
         Height          =   255
         Left            =   5175
         TabIndex        =   8
         Top             =   1660
         Width           =   255
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check7"
         Height          =   255
         Left            =   4095
         TabIndex        =   7
         Top             =   1660
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check6"
         Height          =   255
         Left            =   3015
         TabIndex        =   6
         Top             =   1660
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check5"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   1660
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   1660
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   4095
         TabIndex        =   3
         Top             =   580
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   580
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1935
         TabIndex        =   1
         Top             =   580
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   7320
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   7080
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   6720
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Image Image21 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   735
      End
      Begin VB.Image Image20 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   735
      End
      Begin VB.Image Image19 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   735
      End
      Begin VB.Image Image18 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   4920
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   735
      End
      Begin VB.Image Image17 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   735
      End
      Begin VB.Image Image16 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   735
      End
      Begin VB.Image Image15 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   735
      End
      Begin VB.Image Image14 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   720
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   735
      End
      Begin VB.Image Image13 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   4920
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   735
      End
      Begin VB.Image Image12 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   735
      End
      Begin VB.Image Image11 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   735
      End
      Begin VB.Image Image10 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   735
      End
      Begin VB.Image Image9 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   720
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   735
      End
      Begin VB.Image Image8 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   4920
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   735
      End
      Begin VB.Image Image7 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   735
      End
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   735
      End
      Begin VB.Image Image5 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   735
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   720
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   735
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   840
         Width           =   735
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   840
         Width           =   735
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   840
         Width           =   735
      End
   End
End
Attribute VB_Name = "Kylalaskuri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim terrchoice As Terrain
Dim terrcollection_kyla As New Collection
Dim terrcollection_2 As New Collection
Dim idcollection As New Collection


Dim food As Integer
Dim trade As Integer
Dim prod As Integer

Dim id1 As Integer
Dim id2 As Integer
Dim id3 As Integer
Dim id4 As Integer
Dim id5 As Integer
Dim id6 As Integer
Dim id7 As Integer
Dim id8 As Integer
Dim id9 As Integer
Dim id10 As Integer
Dim id11 As Integer
Dim id12 As Integer
Dim id13 As Integer
Dim id14 As Integer
Dim id15 As Integer
Dim id16 As Integer
Dim id17 As Integer
Dim id18 As Integer
Dim id19 As Integer
Dim id20 As Integer
Dim id21 As Integer





Dim Plains As Terrain
Dim Grassland As Terrain
Dim Desert As Terrain
Dim Arctic As Terrain
Dim Tundra As Terrain
Dim River As Terrain
Dim Forest As Terrain
Dim Swamp As Terrain
Dim Jungle As Terrain
Dim Hills As Terrain
Dim Mountains As Terrain
Dim Ocean As Terrain

Private Sub Check1_Click()
If Check1.Value = 1 Then
Image1 = LoadResPicture(terrcollection_2.Item(id1).terrpicid, 0)
idcollection.Remove 1
idcollection.Add terrcollection_2.Item(id1).foodvalue, "id1", before:=1
Else
Image1 = LoadResPicture(terrcollection_kyla.Item(id1).terrpicid, 0)
idcollection.Remove 1
idcollection.Add terrcollection_kyla.Item(id1).foodvalue, "id1", before:=1
End If
End Sub
Private Sub Check2_Click()
If Check2.Value = 1 Then
Image2 = LoadResPicture(terrcollection_2.Item(id2).terrpicid, 0)
idcollection.Remove 2
idcollection.Add terrcollection_2.Item(id2).foodvalue, "id2", after:=1
Else
Image2 = LoadResPicture(terrcollection_kyla.Item(id2).terrpicid, 0)
idcollection.Remove 2
idcollection.Add terrcollection_kyla.Item(id2).foodvalue, "id2", after:=1
End If
End Sub
Private Sub Check3_Click()
If Check3.Value = 1 Then
Image3 = LoadResPicture(terrcollection_2.Item(id3).terrpicid, 0)
Else
Image3 = LoadResPicture(terrcollection_kyla.Item(id3).terrpicid, 0)
End If
End Sub
Private Sub Check4_Click()
If Check4.Value = 1 Then
Image4 = LoadResPicture(terrcollection_2.Item(id4).terrpicid, 0)
Else
Image4 = LoadResPicture(terrcollection_kyla.Item(id4).terrpicid, 0)
End If
End Sub
Private Sub Check5_Click()
If Check5.Value = 1 Then
Image5 = LoadResPicture(terrcollection_2.Item(id5).terrpicid, 0)
Else
Image5 = LoadResPicture(terrcollection_kyla.Item(id5).terrpicid, 0)
End If
End Sub
Private Sub Check6_Click()
If Check6.Value = 1 Then
Image6 = LoadResPicture(terrcollection_2.Item(id6).terrpicid, 0)
Else
Image6 = LoadResPicture(terrcollection_kyla.Item(id6).terrpicid, 0)
End If
End Sub
Private Sub Check7_Click()
If Check7.Value = 1 Then
Image7 = LoadResPicture(terrcollection_2.Item(id7).terrpicid, 0)
Else
Image7 = LoadResPicture(terrcollection_kyla.Item(id7).terrpicid, 0)
End If
End Sub
Private Sub Check8_Click()
If Check8.Value = 1 Then
Image8 = LoadResPicture(terrcollection_2.Item(id8).terrpicid, 0)
Else
Image8 = LoadResPicture(terrcollection_kyla.Item(id8).terrpicid, 0)
End If
End Sub
Private Sub Check9_Click()
If Check9.Value = 1 Then
Image9 = LoadResPicture(terrcollection_2.Item(id9).terrpicid, 0)
Else
Image9 = LoadResPicture(terrcollection_kyla.Item(id9).terrpicid, 0)
End If
End Sub
Private Sub Check10_Click()
If Check10.Value = 1 Then
Image10 = LoadResPicture(terrcollection_2.Item(id10).terrpicid, 0)
Else
Image10 = LoadResPicture(terrcollection_kyla.Item(id10).terrpicid, 0)
End If
End Sub
Private Sub Check11_Click()
If Check11.Value = 1 Then
Image11 = LoadResPicture(terrcollection_2.Item(id11).terrpicid, 0)
Else
Image11 = LoadResPicture(terrcollection_kyla.Item(id11).terrpicid, 0)
End If
End Sub
Private Sub Check12_Click()
If Check12.Value = 1 Then
Image12 = LoadResPicture(terrcollection_2.Item(id12).terrpicid, 0)
Else
Image12 = LoadResPicture(terrcollection_kyla.Item(id12).terrpicid, 0)
End If
End Sub
Private Sub Check13_Click()
If Check13.Value = 1 Then
Image13 = LoadResPicture(terrcollection_2.Item(id13).terrpicid, 0)
Else
Image13 = LoadResPicture(terrcollection_kyla.Item(id13).terrpicid, 0)
End If
End Sub
Private Sub Check14_Click()
If Check14.Value = 1 Then
Image14 = LoadResPicture(terrcollection_2.Item(id14).terrpicid, 0)
Else
Image14 = LoadResPicture(terrcollection_kyla.Item(id14).terrpicid, 0)
End If
End Sub
Private Sub Check15_Click()
If Check15.Value = 1 Then
Image15 = LoadResPicture(terrcollection_2.Item(id15).terrpicid, 0)
Else
Image15 = LoadResPicture(terrcollection_kyla.Item(id15).terrpicid, 0)
End If
End Sub
Private Sub Check16_Click()
If Check16.Value = 1 Then
Image16 = LoadResPicture(terrcollection_2.Item(id16).terrpicid, 0)
Else
Image16 = LoadResPicture(terrcollection_kyla.Item(id16).terrpicid, 0)
End If
End Sub
Private Sub Check17_Click()
If Check17.Value = 1 Then
Image17 = LoadResPicture(terrcollection_2.Item(id17).terrpicid, 0)
Else
Image17 = LoadResPicture(terrcollection_kyla.Item(id17).terrpicid, 0)
End If
End Sub
Private Sub Check18_Click()
If Check18.Value = 1 Then
Image18 = LoadResPicture(terrcollection_2.Item(id18).terrpicid, 0)
Else
Image18 = LoadResPicture(terrcollection_kyla.Item(id18).terrpicid, 0)
End If
End Sub
Private Sub Check19_Click()
If Check19.Value = 1 Then
Image19 = LoadResPicture(terrcollection_2.Item(id19).terrpicid, 0)
Else
Image19 = LoadResPicture(terrcollection_kyla.Item(id19).terrpicid, 0)
End If
End Sub
Private Sub Check20_Click()
If Check20.Value = 1 Then
Image20 = LoadResPicture(terrcollection_2.Item(id20).terrpicid, 0)
Else
Image20 = LoadResPicture(terrcollection_kyla.Item(id20).terrpicid, 0)
End If
End Sub
Private Sub Check21_Click()
If Check21.Value = 1 Then
Image21 = LoadResPicture(terrcollection_2.Item(id21).terrpicid, 0)
Else
Image21 = LoadResPicture(terrcollection_kyla.Item(id21).terrpicid, 0)
End If
End Sub

Private Sub Form_Load()

id1 = 0
id2 = 0
id3 = 0
id4 = 0
id5 = 0
id6 = 0
id7 = 0
id8 = 0
id9 = 0
id10 = 0
id11 = 0
id12 = 0
id13 = 0
id14 = 0
id15 = 0
id16 = 0
id17 = 0
id18 = 0
id19 = 0
id20 = 0
id21 = 0


idcollection.Add id1, "0"
idcollection.Add id2, "1"





Set Arctic = New Terrain
Arctic.terrname = "Arctic"
Arctic.foodvalue = 0
Arctic.terrpicid = 101
terrcollection_kyla.Add Arctic

Set Grassland = New Terrain
Grassland.terrname = "Grassland"
Grassland.foodvalue = 3
Grassland.terrpicid = 104
terrcollection_kyla.Add Grassland

Set Hills = New Terrain
Hills.terrname = "Hills"
Hills.foodvalue = 2
Hills.terrpicid = 105
terrcollection_kyla.Add Hills

Set Forest = New Terrain
Forest.terrname = "Forest"
Forest.foodvalue = 2
Forest.terrpicid = 103
terrcollection_kyla.Add Forest

Set Mountains = New Terrain
Mountains.terrname = "Mountains"
Mountains.foodvalue = 0
Mountains.terrpicid = 107
terrcollection_kyla.Add Mountains

Set Tundra = New Terrain
Tundra.terrname = "Tundra"
Tundra.foodvalue = 1
Tundra.terrpicid = 112
terrcollection_kyla.Add Tundra

Set Jungle = New Terrain
Jungle.terrname = "Jungle"
Jungle.foodvalue = 3
Jungle.terrpicid = 106
terrcollection_kyla.Add Jungle

Set Plains = New Terrain
Plains.terrname = "Plains"
Plains.foodvalue = 2
Plains.terrpicid = 109
terrcollection_kyla.Add Plains

Set Desert = New Terrain
Desert.terrname = "Desert"
Desert.foodvalue = 1
Desert.terrpicid = 102
terrcollection_kyla.Add Desert

Set Swamp = New Terrain
Swamp.terrname = "Swamp"
Swamp.foodvalue = 3
Swamp.terrpicid = 111
terrcollection_kyla.Add Swamp

Set Ocean = New Terrain
Ocean.terrname = "Ocean"
Ocean.foodvalue = 1
Ocean.terrpicid = 108
terrcollection_kyla.Add Ocean

Set River = New Terrain
River.terrname = "River"
River.foodvalue = 3
River.terrpicid = 110
terrcollection_kyla.Add River

Set Seal = New Terrain
Seal.terrname = "Seal"
Seal.foodvalue = Arctic.foodvalue + 2
Seal.terrpicid = 113
terrcollection_2.Add Seal

Set Resources = New Terrain
Resources.terrname = "Resources"
Resources.foodvalue = Grassland.foodvalue
Resources.terrpicid = 114
terrcollection_2.Add Resources

Set Coal = New Terrain
Coal.terrname = "Coal"
Coal.foodvalue = Hills.foodvalue
Coal.terrpicid = 115
terrcollection_2.Add Coal

Set Deer = New Terrain
Deer.terrname = "Deer"
Deer.foodvalue = Forest.foodvalue + 2
Deer.terrpicid = 116
terrcollection_2.Add Deer

Set Gold = New Terrain
Gold.terrname = "Gold"
Gold.foodvalue = Mountains.foodvalue
Gold.terrpicid = 117
terrcollection_2.Add Gold

Set Moose = New Terrain
Moose.terrname = "Moose"
Moose.foodvalue = Tundra.foodvalue + 2
Moose.terrpicid = 118
terrcollection_2.Add Moose

Set Diamond = New Terrain
Diamond.terrname = "Diamond"
Diamond.foodvalue = Jungle.foodvalue
Diamond.terrpicid = 119
terrcollection_2.Add Diamond

Set Horse = New Terrain
Horse.terrname = "Horse"
Horse.foodvalue = Plains.foodvalue
Horse.terrpicid = 120
terrcollection_2.Add Horse

Set Oasis = New Terrain
Oasis.terrname = "Oasis"
Oasis.foodvalue = Desert.foodvalue + 3
Oasis.terrpicid = 121
terrcollection_2.Add Oasis

Set Oil = New Terrain
Oil.terrname = "Oil"
Oil.foodvalue = Swamp.foodvalue
Oil.terrpicid = 122
terrcollection_2.Add Oil

Set Fish = New Terrain
Fish.terrname = "Fish"
Fish.foodvalue = Ocean.foodvalue + 2
Fish.terrpicid = 123
terrcollection_2.Add Fish

Set Testoriver = New Terrain
Testoriver.terrname = "Testoriver"
Testoriver.foodvalue = River.foodvalue
Testoriver.terrpicid = 124
terrcollection_2.Add Testoriver


End Sub

Function Getactivecombolistchoice(CList As ComboBox)
Getactivecombolistchoice = CList.Text
End Function

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id1 = id1 + 1
If Button = 2 Then id1 = id1 - 1
If id1 > 12 Then id1 = 1
If id1 < 1 Then id1 = 12
If Check1.Value = 1 Then
Image1 = LoadResPicture(terrcollection_2.Item(id1).terrpicid, 0)
idcollection.Remove 1
idcollection.Add terrcollection_2.Item(id1).foodvalue, "id1", before:=1
Else
Image1 = LoadResPicture(terrcollection_kyla.Item(id1).terrpicid, 0)
idcollection.Remove 1
idcollection.Add terrcollection_kyla.Item(id1).foodvalue, "id1", before:=1
End If

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id2 = id2 + 1
If Button = 2 Then id2 = id2 - 1
If id2 > 12 Then id2 = 1
If id2 < 1 Then id2 = 12
If Check2.Value = 1 Then
Image2 = LoadResPicture(terrcollection_2.Item(id2).terrpicid, 0)
idcollection.Remove 2
idcollection.Add terrcollection_2.Item(id2).foodvalue, "id2", after:=1
Else
Image2 = LoadResPicture(terrcollection_kyla.Item(id2).terrpicid, 0)
idcollection.Remove 2
idcollection.Add terrcollection_kyla.Item(id2).foodvalue, "id2", after:=1
End If

End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id3 = id3 + 1
If Button = 2 Then id3 = id3 - 1
If id3 > 12 Then id3 = 1
If id3 < 1 Then id3 = 12
If Check3.Value = 1 Then
Image3 = LoadResPicture(terrcollection_2.Item(id3).terrpicid, 0)
Else
Image3 = LoadResPicture(terrcollection.Item(id3).terrpicid, 0)
End If
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id4 = id4 + 1
If Button = 2 Then id4 = id4 - 1
If id4 > 12 Then id4 = 1
If id4 < 1 Then id4 = 12
If Check4.Value = 1 Then
Image4 = LoadResPicture(terrcollection_2.Item(id4).terrpicid, 0)
Else
Image4 = LoadResPicture(terrcollection.Item(id4).terrpicid, 0)
End If
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id5 = id5 + 1
If Button = 2 Then id5 = id5 - 1
If id5 > 12 Then id5 = 1
If id5 < 1 Then id5 = 12
If Check5.Value = 1 Then
Image5 = LoadResPicture(terrcollection_2.Item(id5).terrpicid, 0)
Else
Image5 = LoadResPicture(terrcollection.Item(id5).terrpicid, 0)
End If
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id6 = id6 + 1
If Button = 2 Then id6 = id6 - 1
If id6 > 12 Then id6 = 1
If id6 < 1 Then id6 = 12
If Check6.Value = 1 Then
Image6 = LoadResPicture(terrcollection_2.Item(id6).terrpicid, 0)
Else
Image6 = LoadResPicture(terrcollection.Item(id6).terrpicid, 0)
End If
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id7 = id7 + 1
If Button = 2 Then id7 = id7 - 1
If id7 > 12 Then id7 = 1
If id7 < 1 Then id7 = 12
If Check7.Value = 1 Then
Image7 = LoadResPicture(terrcollection_2.Item(id7).terrpicid, 0)
Else
Image7 = LoadResPicture(terrcollection.Item(id7).terrpicid, 0)
End If
End Sub

Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id8 = id8 + 1
If Button = 2 Then id8 = id8 - 1
If id8 > 12 Then id8 = 1
If id8 < 1 Then id8 = 12
If Check8.Value = 1 Then
Image8 = LoadResPicture(terrcollection_2.Item(id8).terrpicid, 0)
Else
Image8 = LoadResPicture(terrcollection.Item(id8).terrpicid, 0)
End If
End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id9 = id9 + 1
If Button = 2 Then id9 = id9 - 1
If id9 > 12 Then id9 = 1
If id9 < 1 Then id9 = 12
If Check9.Value = 1 Then
Image9 = LoadResPicture(terrcollection_2.Item(id9).terrpicid, 0)
Else
Image9 = LoadResPicture(terrcollection.Item(id9).terrpicid, 0)
End If
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id10 = id10 + 1
If Button = 2 Then id10 = id10 - 1
If id10 > 12 Then id10 = 1
If id10 < 1 Then id10 = 12
If Check10.Value = 1 Then
Image10 = LoadResPicture(terrcollection_2.Item(id10).terrpicid, 0)
Else
Image10 = LoadResPicture(terrcollection.Item(id10).terrpicid, 0)
End If
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id11 = id11 + 1
If Button = 2 Then id11 = id11 - 1
If id11 > 12 Then id11 = 1
If id11 < 1 Then id11 = 12
If Check11.Value = 1 Then
Image11 = LoadResPicture(terrcollection_2.Item(id11).terrpicid, 0)
Else
Image11 = LoadResPicture(terrcollection.Item(id11).terrpicid, 0)
End If
End Sub

Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id12 = id12 + 1
If Button = 2 Then id12 = id12 - 1
If id12 > 12 Then id12 = 1
If id12 < 1 Then id12 = 12
If Check12.Value = 1 Then
Image12 = LoadResPicture(terrcollection_2.Item(id12).terrpicid, 0)
Else
Image12 = LoadResPicture(terrcollection.Item(id12).terrpicid, 0)
End If
End Sub

Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id13 = id13 + 1
If Button = 2 Then id13 = id13 - 1
If id13 > 12 Then id13 = 1
If id13 < 1 Then id13 = 12
If Check13.Value = 1 Then
Image13 = LoadResPicture(terrcollection_2.Item(id13).terrpicid, 0)
Else
Image13 = LoadResPicture(terrcollection.Item(id13).terrpicid, 0)
End If
End Sub

Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id14 = id14 + 1
If Button = 2 Then id14 = id14 - 1
If id14 > 12 Then id14 = 1
If id14 < 1 Then id14 = 12
If Check14.Value = 1 Then
Image14 = LoadResPicture(terrcollection_2.Item(id14).terrpicid, 0)
Else
Image14 = LoadResPicture(terrcollection.Item(id14).terrpicid, 0)
End If
End Sub

Private Sub Image15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id15 = id15 + 1
If Button = 2 Then id15 = id15 - 1
If id15 > 12 Then id15 = 1
If id15 < 1 Then id15 = 12
If Check15.Value = 1 Then
Image15 = LoadResPicture(terrcollection_2.Item(id15).terrpicid, 0)
Else
Image15 = LoadResPicture(terrcollection.Item(id15).terrpicid, 0)
End If
End Sub

Private Sub Image16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id16 = id16 + 1
If Button = 2 Then id16 = id16 - 1
If id16 > 12 Then id16 = 1
If id16 < 1 Then id16 = 12
If Check16.Value = 1 Then
Image16 = LoadResPicture(terrcollection_2.Item(id16).terrpicid, 0)
Else
Image16 = LoadResPicture(terrcollection.Item(id16).terrpicid, 0)
End If
End Sub

Private Sub Image17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id17 = id17 + 1
If Button = 2 Then id17 = id17 - 1
If id17 > 12 Then id17 = 1
If id17 < 1 Then id17 = 12
If Check17.Value = 1 Then
Image17 = LoadResPicture(terrcollection_2.Item(id17).terrpicid, 0)
Else
Image17 = LoadResPicture(terrcollection.Item(id17).terrpicid, 0)
End If
End Sub

Private Sub Image18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id18 = id18 + 1
If Button = 2 Then id18 = id18 - 1
If id18 > 12 Then id18 = 1
If id18 < 1 Then id18 = 12
If Check18.Value = 1 Then
Image18 = LoadResPicture(terrcollection_2.Item(id18).terrpicid, 0)
Else
Image18 = LoadResPicture(terrcollection.Item(id18).terrpicid, 0)
End If
End Sub

Private Sub Image19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id19 = id19 + 1
If Button = 2 Then id19 = id19 - 1
If id19 > 12 Then id19 = 1
If id19 < 1 Then id19 = 12
If Check19.Value = 1 Then
Image19 = LoadResPicture(terrcollection_2.Item(id19).terrpicid, 0)
Else
Image19 = LoadResPicture(terrcollection.Item(id19).terrpicid, 0)
End If
End Sub

Private Sub Image20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id20 = id20 + 1
If Button = 2 Then id20 = id20 - 1
If id20 > 12 Then id20 = 1
If id20 < 1 Then id20 = 12
If Check20.Value = 1 Then
Image20 = LoadResPicture(terrcollection_2.Item(id20).terrpicid, 0)
Else
Image20 = LoadResPicture(terrcollection.Item(id20).terrpicid, 0)
End If
End Sub

Private Sub Image21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id21 = id21 + 1
If Button = 2 Then id21 = id21 - 1
If id21 > 12 Then id21 = 1
If id21 < 1 Then id21 = 12
If Check21.Value = 1 Then
Image21 = LoadResPicture(terrcollection_2.Item(id21).terrpicid, 0)
Else
Image21 = LoadResPicture(terrcollection.Item(id21).terrpicid, 0)
End If
End Sub

Private Sub laske_Click()
Dim totalfood As Integer
totalfood = 0
For Each id In idcollection
totalfood = totalfood + id
Label2.Caption = CStr(id)
Next
Label1.Caption = CStr(totalfood)

End Sub
