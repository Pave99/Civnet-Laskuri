VERSION 5.00
Begin VB.Form Kylalaskuri 
   Caption         =   "Kylä-laskuri"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Kylalaskuri"
   ScaleHeight     =   7590
   ScaleMode       =   0  'User
   ScaleWidth      =   6467.336
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6453
      Begin VB.CommandButton laske 
         Caption         =   "Laske"
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   5800
         Width           =   1215
      End
      Begin VB.CheckBox Check21 
         Caption         =   "Check21"
         Height          =   255
         Left            =   4095
         TabIndex        =   21
         Top             =   4700
         Width           =   255
      End
      Begin VB.CheckBox Check20 
         Caption         =   "Check20"
         Height          =   255
         Left            =   3015
         TabIndex        =   20
         Top             =   4700
         Width           =   255
      End
      Begin VB.CheckBox Check19 
         Caption         =   "Check19"
         Height          =   255
         Left            =   1945
         TabIndex        =   19
         Top             =   4700
         Width           =   255
      End
      Begin VB.CheckBox Check18 
         Caption         =   "Check18"
         Height          =   255
         Left            =   5175
         TabIndex        =   18
         Top             =   3620
         Width           =   255
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Check17"
         Height          =   255
         Left            =   4095
         TabIndex        =   17
         Top             =   3620
         Width           =   255
      End
      Begin VB.CheckBox Check16 
         Caption         =   "Check16"
         Height          =   255
         Left            =   3015
         TabIndex        =   16
         Top             =   3620
         Width           =   255
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Check15"
         Height          =   255
         Left            =   1935
         TabIndex        =   15
         Top             =   3620
         Width           =   255
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Check14"
         Height          =   255
         Left            =   975
         TabIndex        =   14
         Top             =   3620
         Width           =   255
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Check13"
         Height          =   255
         Left            =   5175
         TabIndex        =   13
         Top             =   2540
         Width           =   255
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Check12"
         Height          =   255
         Left            =   4095
         TabIndex        =   12
         Top             =   2540
         Width           =   255
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Check11"
         Height          =   255
         Left            =   3015
         TabIndex        =   11
         Top             =   2540
         Width           =   255
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Check10"
         Height          =   255
         Left            =   1935
         TabIndex        =   10
         Top             =   2540
         Width           =   255
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Check9"
         Height          =   255
         Left            =   975
         TabIndex        =   9
         Top             =   2540
         Width           =   255
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check8"
         Height          =   255
         Left            =   5175
         TabIndex        =   8
         Top             =   1460
         Width           =   255
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check7"
         Height          =   255
         Left            =   4095
         TabIndex        =   7
         Top             =   1460
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check6"
         Height          =   255
         Left            =   3015
         TabIndex        =   6
         Top             =   1460
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check5"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   1460
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   1460
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   4095
         TabIndex        =   3
         Top             =   380
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   380
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1935
         TabIndex        =   1
         Top             =   380
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   24
         Top             =   7005
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   22
         Top             =   6660
         Width           =   60
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   6720
         Y1              =   6320
         Y2              =   6320
      End
      Begin VB.Image Image21 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   4960
         Width           =   735
      End
      Begin VB.Image Image20 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   4960
         Width           =   735
      End
      Begin VB.Image Image19 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   4960
         Width           =   735
      End
      Begin VB.Image Image18 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   4920
         Stretch         =   -1  'True
         Top             =   3880
         Width           =   735
      End
      Begin VB.Image Image17 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   3880
         Width           =   735
      End
      Begin VB.Image Image16 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   3880
         Width           =   735
      End
      Begin VB.Image Image15 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   3880
         Width           =   735
      End
      Begin VB.Image Image14 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   720
         Stretch         =   -1  'True
         Top             =   3880
         Width           =   735
      End
      Begin VB.Image Image13 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   4920
         Stretch         =   -1  'True
         Top             =   2800
         Width           =   735
      End
      Begin VB.Image Image12 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   2800
         Width           =   735
      End
      Begin VB.Image Image11 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   2800
         Width           =   735
      End
      Begin VB.Image Image10 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   2800
         Width           =   735
      End
      Begin VB.Image Image9 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   720
         Stretch         =   -1  'True
         Top             =   2800
         Width           =   735
      End
      Begin VB.Image Image8 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   4920
         Stretch         =   -1  'True
         Top             =   1720
         Width           =   735
      End
      Begin VB.Image Image7 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   1720
         Width           =   735
      End
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   1720
         Width           =   735
      End
      Begin VB.Image Image5 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   1720
         Width           =   735
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   720
         Stretch         =   -1  'True
         Top             =   1720
         Width           =   735
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   640
         Width           =   735
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   640
         Width           =   735
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   640
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
If id1 > 0 Then
If Check1.Value = 1 Then
Image1 = LoadResPicture(terrcollection_2.Item(id1).terrpicid, 0)
idcollection.Remove 1
idcollection.Add terrcollection_2.Item(id1).foodvalue, "id1", before:=1
Else
Image1 = LoadResPicture(terrcollection_kyla.Item(id1).terrpicid, 0)
idcollection.Remove 1
idcollection.Add terrcollection_kyla.Item(id1).foodvalue, "id1", before:=1
End If
End If
End Sub
Private Sub Check2_Click()
If id2 > 0 Then
If Check2.Value = 1 Then
Image2 = LoadResPicture(terrcollection_2.Item(id2).terrpicid, 0)
idcollection.Remove 2
idcollection.Add terrcollection_2.Item(id2).foodvalue, "id2", after:=1
Else
Image2 = LoadResPicture(terrcollection_kyla.Item(id2).terrpicid, 0)
idcollection.Remove 2
idcollection.Add terrcollection_kyla.Item(id2).foodvalue, "id2", after:=1
End If
End If
End Sub
Private Sub Check3_Click()
If id3 > 0 Then
If Check3.Value = 1 Then
Image3 = LoadResPicture(terrcollection_2.Item(id3).terrpicid, 0)
Else
Image3 = LoadResPicture(terrcollection_kyla.Item(id3).terrpicid, 0)
End If
End If
End Sub
Private Sub Check4_Click()
If id4 > 0 Then
If Check4.Value = 1 Then
Image4 = LoadResPicture(terrcollection_2.Item(id4).terrpicid, 0)
Else
Image4 = LoadResPicture(terrcollection_kyla.Item(id4).terrpicid, 0)
End If
End If
End Sub
Private Sub Check5_Click()
If id5 > 0 Then
If Check5.Value = 1 Then
Image5 = LoadResPicture(terrcollection_2.Item(id5).terrpicid, 0)
Else
Image5 = LoadResPicture(terrcollection_kyla.Item(id5).terrpicid, 0)
End If
End If
End Sub
Private Sub Check6_Click()
If id6 > 0 Then
If Check6.Value = 1 Then
Image6 = LoadResPicture(terrcollection_2.Item(id6).terrpicid, 0)
Else
Image6 = LoadResPicture(terrcollection_kyla.Item(id6).terrpicid, 0)
End If
End If
End Sub
Private Sub Check7_Click()
If id7 > 0 Then
If Check7.Value = 1 Then
Image7 = LoadResPicture(terrcollection_2.Item(id7).terrpicid, 0)
Else
Image7 = LoadResPicture(terrcollection_kyla.Item(id7).terrpicid, 0)
End If
End If
End Sub
Private Sub Check8_Click()
If id8 > 0 Then
If Check8.Value = 1 Then
Image8 = LoadResPicture(terrcollection_2.Item(id8).terrpicid, 0)
Else
Image8 = LoadResPicture(terrcollection_kyla.Item(id8).terrpicid, 0)
End If
End If
End Sub
Private Sub Check9_Click()
If id9 > 0 Then
If Check9.Value = 1 Then
Image9 = LoadResPicture(terrcollection_2.Item(id9).terrpicid, 0)
Else
Image9 = LoadResPicture(terrcollection_kyla.Item(id9).terrpicid, 0)
End If
End If
End Sub
Private Sub Check10_Click()
If id10 > 0 Then
If Check10.Value = 1 Then
Image10 = LoadResPicture(terrcollection_2.Item(id10).terrpicid, 0)
Else
Image10 = LoadResPicture(terrcollection_kyla.Item(id10).terrpicid, 0)
End If
End If
End Sub
Private Sub Check11_Click()
If id11 > 0 Then
If Check11.Value = 1 Then
Image11 = LoadResPicture(terrcollection_2.Item(id11).terrpicid, 0)
Else
Image11 = LoadResPicture(terrcollection_kyla.Item(id11).terrpicid, 0)
End If
End If
End Sub
Private Sub Check12_Click()
If id12 > 0 Then
If Check12.Value = 1 Then
Image12 = LoadResPicture(terrcollection_2.Item(id12).terrpicid, 0)
Else
Image12 = LoadResPicture(terrcollection_kyla.Item(id12).terrpicid, 0)
End If
End If
End Sub
Private Sub Check13_Click()
If id13 > 0 Then
If Check13.Value = 1 Then
Image13 = LoadResPicture(terrcollection_2.Item(id13).terrpicid, 0)
Else
Image13 = LoadResPicture(terrcollection_kyla.Item(id13).terrpicid, 0)
End If
End If
End Sub
Private Sub Check14_Click()
If id14 > 0 Then
If Check14.Value = 1 Then
Image14 = LoadResPicture(terrcollection_2.Item(id14).terrpicid, 0)
Else
Image14 = LoadResPicture(terrcollection_kyla.Item(id14).terrpicid, 0)
End If
End If
End Sub
Private Sub Check15_Click()
If id15 > 0 Then
If Check15.Value = 1 Then
Image15 = LoadResPicture(terrcollection_2.Item(id15).terrpicid, 0)
Else
Image15 = LoadResPicture(terrcollection_kyla.Item(id15).terrpicid, 0)
End If
End If
End Sub
Private Sub Check16_Click()
If id16 > 0 Then
If Check16.Value = 1 Then
Image16 = LoadResPicture(terrcollection_2.Item(id16).terrpicid, 0)
Else
Image16 = LoadResPicture(terrcollection_kyla.Item(id16).terrpicid, 0)
End If
End If
End Sub
Private Sub Check17_Click()
If id17 > 0 Then
If Check17.Value = 1 Then
Image17 = LoadResPicture(terrcollection_2.Item(id17).terrpicid, 0)
Else
Image17 = LoadResPicture(terrcollection_kyla.Item(id17).terrpicid, 0)
End If
End If
End Sub
Private Sub Check18_Click()
If id18 > 0 Then
If Check18.Value = 1 Then
Image18 = LoadResPicture(terrcollection_2.Item(id18).terrpicid, 0)
Else
Image18 = LoadResPicture(terrcollection_kyla.Item(id18).terrpicid, 0)
End If
End If
End Sub
Private Sub Check19_Click()
If id19 > 0 Then
If Check19.Value = 1 Then
Image19 = LoadResPicture(terrcollection_2.Item(id19).terrpicid, 0)
Else
Image19 = LoadResPicture(terrcollection_kyla.Item(id19).terrpicid, 0)
End If
End If
End Sub
Private Sub Check20_Click()
If id20 > 0 Then
If Check20.Value = 1 Then
Image20 = LoadResPicture(terrcollection_2.Item(id20).terrpicid, 0)
Else
Image20 = LoadResPicture(terrcollection_kyla.Item(id20).terrpicid, 0)
End If
End If
End Sub
Private Sub Check21_Click()
If id21 > 0 Then
If Check21.Value = 1 Then
Image21 = LoadResPicture(terrcollection_2.Item(id21).terrpicid, 0)
Else
Image21 = LoadResPicture(terrcollection_kyla.Item(id21).terrpicid, 0)
End If
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
idcollection.Add id3, "2"
idcollection.Add id4, "3"
idcollection.Add id5, "4"
idcollection.Add id6, "5"
idcollection.Add id7, "6"
idcollection.Add id8, "7"
idcollection.Add id9, "8"
idcollection.Add id10, "9"
idcollection.Add id11, "10"
idcollection.Add id12, "11"
idcollection.Add id13, "12"
idcollection.Add id14, "13"
idcollection.Add id15, "14"
idcollection.Add id16, "15"
idcollection.Add id17, "16"
idcollection.Add id18, "17"
idcollection.Add id19, "18"
idcollection.Add id20, "19"
idcollection.Add id21, "20"

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
    idcollection.Remove 3
    idcollection.Add terrcollection_2.Item(id3).foodvalue, "id3", after:=2
Else
    Image3 = LoadResPicture(terrcollection_kyla.Item(id3).terrpicid, 0)
    idcollection.Remove 3
    idcollection.Add terrcollection_kyla.Item(id3).foodvalue, "id3", after:=2
End If
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id4 = id4 + 1
If Button = 2 Then id4 = id4 - 1
If id4 > 12 Then id4 = 1
If id4 < 1 Then id4 = 12
If Check4.Value = 1 Then
    Image4 = LoadResPicture(terrcollection_2.Item(id4).terrpicid, 0)
    idcollection.Remove 4
    idcollection.Add terrcollection_2.Item(id4).foodvalue, "id4", after:=3
Else
    Image4 = LoadResPicture(terrcollection_kyla.Item(id4).terrpicid, 0)
    idcollection.Remove 4
    idcollection.Add terrcollection_kyla.Item(id4).foodvalue, "id4", after:=3
End If
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id5 = id5 + 1
If Button = 2 Then id5 = id5 - 1
If id5 > 12 Then id5 = 1
If id5 < 1 Then id5 = 12
If Check5.Value = 1 Then
    Image5 = LoadResPicture(terrcollection_2.Item(id5).terrpicid, 0)
    idcollection.Remove 5
    idcollection.Add terrcollection_2.Item(id5).foodvalue, "id5", after:=4
Else
    Image5 = LoadResPicture(terrcollection_kyla.Item(id5).terrpicid, 0)
    idcollection.Remove 5
    idcollection.Add terrcollection_kyla.Item(id5).foodvalue, "id5", after:=4
End If
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id6 = id6 + 1
If Button = 2 Then id6 = id6 - 1
If id6 > 12 Then id6 = 1
If id6 < 1 Then id6 = 12
If Check6.Value = 1 Then
    Image6 = LoadResPicture(terrcollection_2.Item(id6).terrpicid, 0)
    idcollection.Remove 6
    idcollection.Add terrcollection_2.Item(id6).foodvalue, "id6", after:=5
Else
    Image6 = LoadResPicture(terrcollection_kyla.Item(id6).terrpicid, 0)
    idcollection.Remove 6
    idcollection.Add terrcollection_kyla.Item(id6).foodvalue, "id6", after:=5
End If
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id7 = id7 + 1
If Button = 2 Then id7 = id7 - 1
If id7 > 12 Then id7 = 1
If id7 < 1 Then id7 = 12
If Check7.Value = 1 Then
    Image7 = LoadResPicture(terrcollection_2.Item(id7).terrpicid, 0)
    idcollection.Remove 7
    idcollection.Add terrcollection_2.Item(id7).foodvalue, "id7", after:=6
Else
    Image7 = LoadResPicture(terrcollection_kyla.Item(id7).terrpicid, 0)
    idcollection.Remove 7
    idcollection.Add terrcollection_kyla.Item(id7).foodvalue, "id7", after:=6
End If
End Sub

Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id8 = id8 + 1
If Button = 2 Then id8 = id8 - 1
If id8 > 12 Then id8 = 1
If id8 < 1 Then id8 = 12
If Check8.Value = 1 Then
    Image8 = LoadResPicture(terrcollection_2.Item(id8).terrpicid, 0)
    idcollection.Remove 8
    idcollection.Add terrcollection_2.Item(id8).foodvalue, "id8", after:=7
Else
    Image8 = LoadResPicture(terrcollection_kyla.Item(id8).terrpicid, 0)
    idcollection.Remove 8
    idcollection.Add terrcollection_kyla.Item(id8).foodvalue, "id8", after:=7
End If
End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id9 = id9 + 1
If Button = 2 Then id9 = id9 - 1
If id9 > 12 Then id9 = 1
If id9 < 1 Then id9 = 12
If Check9.Value = 1 Then
    Image9 = LoadResPicture(terrcollection_2.Item(id9).terrpicid, 0)
    idcollection.Remove 9
    idcollection.Add terrcollection_2.Item(id9).foodvalue, "id9", after:=8
Else
    Image9 = LoadResPicture(terrcollection_kyla.Item(id9).terrpicid, 0)
    idcollection.Remove 9
    idcollection.Add terrcollection_kyla.Item(id9).foodvalue, "id9", after:=8
End If
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id10 = id10 + 1
If Button = 2 Then id10 = id10 - 1
If id10 > 12 Then id10 = 1
If id10 < 1 Then id10 = 12
If Check10.Value = 1 Then
    Image10 = LoadResPicture(terrcollection_2.Item(id10).terrpicid, 0)
    idcollection.Remove 10
    idcollection.Add terrcollection_2.Item(id10).foodvalue, "id10", after:=9
Else
    Image10 = LoadResPicture(terrcollection_kyla.Item(id10).terrpicid, 0)
    idcollection.Remove 10
    idcollection.Add terrcollection_kyla.Item(id10).foodvalue, "id10", after:=9
End If
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id11 = id11 + 1
If Button = 2 Then id11 = id11 - 1
If id11 > 12 Then id11 = 1
If id11 < 1 Then id11 = 12
If Check11.Value = 1 Then
    Image11 = LoadResPicture(terrcollection_2.Item(id11).terrpicid, 0)
    idcollection.Remove 11
    idcollection.Add terrcollection_2.Item(id11).foodvalue, "id11", after:=10
Else
    Image11 = LoadResPicture(terrcollection_kyla.Item(id11).terrpicid, 0)
    idcollection.Remove 11
    idcollection.Add terrcollection_kyla.Item(id11).foodvalue, "id11", after:=10
End If
End Sub

Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id12 = id12 + 1
If Button = 2 Then id12 = id12 - 1
If id12 > 12 Then id12 = 1
If id12 < 1 Then id12 = 12
If Check12.Value = 1 Then
    Image12 = LoadResPicture(terrcollection_2.Item(id12).terrpicid, 0)
    idcollection.Remove 12
    idcollection.Add terrcollection_2.Item(id12).foodvalue, "id12", after:=11
Else
    Image12 = LoadResPicture(terrcollection_kyla.Item(id12).terrpicid, 0)
    idcollection.Remove 12
    idcollection.Add terrcollection_kyla.Item(id12).foodvalue, "id12", after:=11
End If
End Sub

Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id13 = id13 + 1
If Button = 2 Then id13 = id13 - 1
If id13 > 12 Then id13 = 1
If id13 < 1 Then id13 = 12
If Check13.Value = 1 Then
    Image13 = LoadResPicture(terrcollection_2.Item(id13).terrpicid, 0)
    idcollection.Remove 13
    idcollection.Add terrcollection_2.Item(id13).foodvalue, "id13", after:=12
Else
    Image13 = LoadResPicture(terrcollection_kyla.Item(id13).terrpicid, 0)
    idcollection.Remove 13
    idcollection.Add terrcollection_kyla.Item(id13).foodvalue, "id13", after:=12
End If
End Sub

Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id14 = id14 + 1
If Button = 2 Then id14 = id14 - 1
If id14 > 12 Then id14 = 1
If id14 < 1 Then id14 = 12
If Check14.Value = 1 Then
    Image14 = LoadResPicture(terrcollection_2.Item(id14).terrpicid, 0)
    idcollection.Remove 14
    idcollection.Add terrcollection_2.Item(id14).foodvalue, "id14", after:=13
Else
    Image14 = LoadResPicture(terrcollection_kyla.Item(id14).terrpicid, 0)
    idcollection.Remove 14
    idcollection.Add terrcollection_kyla.Item(id14).foodvalue, "id14", after:=13
End If
End Sub

Private Sub Image15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id15 = id15 + 1
If Button = 2 Then id15 = id15 - 1
If id15 > 12 Then id15 = 1
If id15 < 1 Then id15 = 12
If Check15.Value = 1 Then
    Image15 = LoadResPicture(terrcollection_2.Item(id15).terrpicid, 0)
    idcollection.Remove 15
    idcollection.Add terrcollection_2.Item(id15).foodvalue, "id15", after:=14
Else
    Image15 = LoadResPicture(terrcollection_kyla.Item(id15).terrpicid, 0)
    idcollection.Remove 15
    idcollection.Add terrcollection_kyla.Item(id15).foodvalue, "id15", after:=14
End If
End Sub

Private Sub Image16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id16 = id16 + 1
If Button = 2 Then id16 = id16 - 1
If id16 > 12 Then id16 = 1
If id16 < 1 Then id16 = 12
If Check16.Value = 1 Then
    Image16 = LoadResPicture(terrcollection_2.Item(id16).terrpicid, 0)
    idcollection.Remove 16
    idcollection.Add terrcollection_2.Item(id16).foodvalue, "id16", after:=15
Else
    Image16 = LoadResPicture(terrcollection_kyla.Item(id16).terrpicid, 0)
    idcollection.Remove 16
    idcollection.Add terrcollection_kyla.Item(id16).foodvalue, "id16", after:=15
End If
End Sub

Private Sub Image17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id17 = id17 + 1
If Button = 2 Then id17 = id17 - 1
If id17 > 12 Then id17 = 1
If id17 < 1 Then id17 = 12
If Check17.Value = 1 Then
    Image17 = LoadResPicture(terrcollection_2.Item(id17).terrpicid, 0)
    idcollection.Remove 17
    idcollection.Add terrcollection_2.Item(id17).foodvalue, "id17", after:=16
Else
    Image17 = LoadResPicture(terrcollection_kyla.Item(id17).terrpicid, 0)
    idcollection.Remove 17
    idcollection.Add terrcollection_kyla.Item(id17).foodvalue, "id17", after:=16
End If
End Sub

Private Sub Image18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id18 = id18 + 1
If Button = 2 Then id18 = id18 - 1
If id18 > 12 Then id18 = 1
If id18 < 1 Then id18 = 12
If Check18.Value = 1 Then
    Image18 = LoadResPicture(terrcollection_2.Item(id18).terrpicid, 0)
    idcollection.Remove 18
    idcollection.Add terrcollection_2.Item(id18).foodvalue, "id18", after:=17
Else
    Image18 = LoadResPicture(terrcollection_kyla.Item(id18).terrpicid, 0)
    idcollection.Remove 18
    idcollection.Add terrcollection_kyla.Item(id18).foodvalue, "id18", after:=17
End If
End Sub

Private Sub Image19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id19 = id19 + 1
If Button = 2 Then id19 = id19 - 1
If id19 > 12 Then id19 = 1
If id19 < 1 Then id19 = 12
If Check19.Value = 1 Then
    Image19 = LoadResPicture(terrcollection_2.Item(id19).terrpicid, 0)
    idcollection.Remove 19
    idcollection.Add terrcollection_2.Item(id19).foodvalue, "id19", after:=18
Else
    Image19 = LoadResPicture(terrcollection_kyla.Item(id19).terrpicid, 0)
    idcollection.Remove 19
    idcollection.Add terrcollection_kyla.Item(id19).foodvalue, "id19", after:=18
End If
End Sub

Private Sub Image20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id20 = id20 + 1
If Button = 2 Then id20 = id20 - 1
If id20 > 12 Then id20 = 1
If id20 < 1 Then id20 = 12
If Check20.Value = 1 Then
    Image20 = LoadResPicture(terrcollection_2.Item(id20).terrpicid, 0)
    idcollection.Remove 20
    idcollection.Add terrcollection_2.Item(id20).foodvalue, "id20", after:=19
Else
    Image20 = LoadResPicture(terrcollection_kyla.Item(id20).terrpicid, 0)
    idcollection.Remove 20
    idcollection.Add terrcollection_kyla.Item(id20).foodvalue, "id20", after:=19
End If
End Sub

Private Sub Image21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then id21 = id21 + 1
If Button = 2 Then id21 = id21 - 1
If id21 > 12 Then id21 = 1
If id21 < 1 Then id21 = 12
If Check21.Value = 1 Then
    Image21 = LoadResPicture(terrcollection_2.Item(id21).terrpicid, 0)
    idcollection.Remove 21
    idcollection.Add terrcollection_2.Item(id21).foodvalue, "id21", after:=20
Else
    Image21 = LoadResPicture(terrcollection_kyla.Item(id21).terrpicid, 0)
    idcollection.Remove 21
    idcollection.Add terrcollection_kyla.Item(id21).foodvalue, "id21", after:=20
End If
End Sub

Private Sub laske_Click()
Dim totalfood As Integer
totalfood = 0
For Each id In idcollection
totalfood = totalfood + id
Next
Label1.Caption = "Ruokaa yhteensä: " + CStr(totalfood)
Label2.Caption = "Suurin mahdollinen kylän koko: " + CStr(Fix(totalfood / 2))

End Sub
