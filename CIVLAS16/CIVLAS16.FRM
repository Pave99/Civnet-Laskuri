VERSION 4.00
Begin VB.Form Laskuri 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Civnet-laskuri"
   ClientHeight    =   5475
   ClientLeft      =   3375
   ClientTop       =   3060
   ClientWidth     =   6720
   Height          =   5880
   Icon            =   "CIVLAS16.frx":0000
   Left            =   3315
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6720
   Top             =   2715
   Width           =   6840
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CheckBox mute 
         Caption         =   "Mute"
         Height          =   495
         Left            =   240
         TabIndex        =   31
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox winchance_def 
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   4920
         Width           =   2415
      End
      Begin VB.TextBox winchance_att 
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4320
         Width           =   2415
      End
      Begin VB.CommandButton calc 
         Caption         =   "Laske"
         Height          =   615
         Left            =   5400
         TabIndex        =   25
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CheckBox def_f_check 
         Height          =   495
         Left            =   4200
         TabIndex        =   21
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton Optiontwothirds 
         Caption         =   "Option2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Optiononethird 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5160
         TabIndex        =   15
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox thirds_check 
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Citywalls_check 
         Caption         =   "Check2"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox Veteran_check_def 
         Height          =   495
         Left            =   2520
         TabIndex        =   9
         Top             =   1680
         Width           =   495
      End
      Begin VB.CheckBox Veteran_check_att 
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.ComboBox terr_list 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox def_list 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox att_list 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.OLE we_die 
         AutoActivate    =   0  'Manual
         Class           =   "SoundRec"
         Height          =   495
         Left            =   4560
         OleObjectBlob   =   "CIVLAS16.frx":08CA
         SourceDoc       =   "c:\civ\we_die.wav"
         TabIndex        =   30
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OLE they_die 
         AutoActivate    =   0  'Manual
         Class           =   "SoundRec"
         Height          =   495
         Left            =   3840
         OleObjectBlob   =   "CIVLAS16.frx":7CE2
         SourceDoc       =   "c:\civ\they_die.wav"
         TabIndex        =   29
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label winchance_def_text 
         Caption         =   "Voittomahdollisuus puolustajalla:"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   4920
         Width           =   3615
      End
      Begin VB.Label winprob_text 
         Caption         =   "Voittomahdollisuus hyökkääjällä:"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   4320
         Width           =   3855
      End
      Begin VB.Label terr_result 
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   18
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3600
         TabIndex        =   23
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Line Line38 
         X1              =   4800
         X2              =   4800
         Y1              =   1560
         Y2              =   1200
      End
      Begin VB.Line Line37 
         X1              =   4800
         X2              =   3840
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line36 
         X1              =   4800
         X2              =   3840
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line35 
         X1              =   3840
         X2              =   3840
         Y1              =   1200
         Y2              =   1560
      End
      Begin VB.Label def_f_check_text 
         Caption         =   "Fortify"
         Height          =   375
         Left            =   3960
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label def_result 
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   18
            underline       =   -1  'True
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5880
         TabIndex        =   20
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label att_result 
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   400
            size            =   18
            underline       =   -1  'True
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5880
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.Line Line34 
         X1              =   5520
         X2              =   5040
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line33 
         X1              =   5040
         X2              =   5040
         Y1              =   240
         Y2              =   1560
      End
      Begin VB.Line Line32 
         X1              =   5520
         X2              =   5520
         Y1              =   1560
         Y2              =   240
      End
      Begin VB.Line Line31 
         X1              =   5040
         X2              =   5520
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label twothirds_text 
         Caption         =   "2"
         Height          =   255
         Left            =   5205
         TabIndex        =   18
         Top             =   960
         Width           =   255
      End
      Begin VB.Label onethird 
         Caption         =   "1"
         Height          =   255
         Left            =   5205
         TabIndex        =   17
         Top             =   360
         Width           =   255
      End
      Begin VB.Line Line30 
         X1              =   3840
         X2              =   4800
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line29 
         X1              =   4800
         X2              =   4800
         Y1              =   600
         Y2              =   240
      End
      Begin VB.Line Line28 
         X1              =   3840
         X2              =   3840
         Y1              =   240
         Y2              =   600
      End
      Begin VB.Line Line27 
         X1              =   4800
         X2              =   3840
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Moves_text 
         Caption         =   "Kolmasosat"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Line Line26 
         X1              =   2160
         X2              =   2160
         Y1              =   2400
         Y2              =   2760
      End
      Begin VB.Line Line25 
         X1              =   3120
         X2              =   2160
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line24 
         X1              =   2160
         X2              =   3120
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line23 
         X1              =   3120
         X2              =   3120
         Y1              =   2400
         Y2              =   2760
      End
      Begin VB.Label Citywalls_text 
         Caption         =   "City Walls"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Line Line22 
         X1              =   240
         X2              =   240
         Y1              =   2760
         Y2              =   2400
      End
      Begin VB.Line Line21 
         X1              =   240
         X2              =   1320
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line20 
         X1              =   1320
         X2              =   1320
         Y1              =   2400
         Y2              =   2760
      End
      Begin VB.Line Line19 
         X1              =   240
         X2              =   1320
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Terrain_text 
         Caption         =   "Maasto"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Line Line18 
         X1              =   2280
         X2              =   2160
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line17 
         X1              =   2160
         X2              =   2160
         Y1              =   1200
         Y2              =   1560
      End
      Begin VB.Line Line16 
         X1              =   3120
         X2              =   3120
         Y1              =   1200
         Y2              =   1560
      End
      Begin VB.Line Line15 
         X1              =   2160
         X2              =   3120
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line14 
         X1              =   2280
         X2              =   3120
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Veteran_text_def 
         Caption         =   "Veteraani"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.Line Line13 
         X1              =   240
         X2              =   240
         Y1              =   1200
         Y2              =   1560
      End
      Begin VB.Line Line12 
         X1              =   240
         X2              =   1320
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line11 
         X1              =   1320
         X2              =   1320
         Y1              =   1560
         Y2              =   1200
      End
      Begin VB.Line Line10 
         X1              =   240
         X2              =   1320
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Def_text 
         Caption         =   "Puolustava"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Line Line9 
         X1              =   3120
         X2              =   2160
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line8 
         X1              =   2160
         X2              =   2160
         Y1              =   600
         Y2              =   240
      End
      Begin VB.Line Line7 
         X1              =   3120
         X2              =   3120
         Y1              =   600
         Y2              =   240
      End
      Begin VB.Line Line6 
         X1              =   3120
         X2              =   2160
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Veteran_text_att 
         Caption         =   "Veteraani"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   1320
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   240
         Y1              =   600
         Y2              =   240
      End
      Begin VB.Line Line2 
         X1              =   1320
         X2              =   1320
         Y1              =   600
         Y2              =   240
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   1320
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Att_text 
         Caption         =   "Hyökkäävä"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Laskuri"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Dim Militia As Unit
Dim Phalanx As Unit
Dim Cavalry As Unit
Dim Legion As Unit
Dim Chariot As Unit
Dim Trireme As Unit
Dim Catapult As Unit
Dim Knight As Unit
Dim Sail As Unit
Dim Ironclad As Unit
Dim Musketeer As Unit
Dim Frigate As Unit
Dim Cannon As Unit
Dim Cruiser As Unit
Dim Transport As Unit
Dim Fighter As Unit
Dim Bomber As Unit
Dim Carrier As Unit
Dim Rifleman As Unit
Dim Submarine As Unit
Dim Battleship As Unit
Dim Armor As Unit
Dim MechInf As Unit
Dim Artillery As Unit
Dim Nuclear As Unit

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
Dim Citywalls As Terrain


Dim attchoice As Unit
Dim defchoice As Unit
Dim terrchoice As Terrain
Dim attunit As Unit
Dim bens As Unit



Dim unitcollection As New Collection
Dim terrcollection As New Collection


Dim defbonus As Double
Const fortifybonus As Double = 0.5
Const veteranbonus As Double = 0.5


Dim terr As Terrain
Dim Unit As Unit

Function Getactivecombolistchoice(CList As ComboBox)
Getactivecombolistchoice = CList.Text
End Function

Private Sub att_list_Click()
Set attchoice = New Unit
attchoice.Unitname = Getactivecombolistchoice(att_list)
    
For Each Unit In unitcollection
    If Unit.Unitname = attchoice.Unitname Then
        If attchoice.Unitname = "Trireme" Or attchoice.Unitname = "Sail" Or attchoice.Unitname = "Ironclad" Or attchoice.Unitname = "Frigate" Or attchoice.Unitname = "Cruiser" Or attchoice.Unitname = "Submarine" Or attchoice.Unitname = "Battleship" Or attchoice.Unitname = "Transport" Or attchoice.Unitname = "Carrier" Or attchoice.Unitname = "Fighter" Or attchoice.Unitname = "Bomber" Or attchoice.Unitname = "Nuclear" Then
            thirds_check.Enabled = False
            'thirds_check.Value = 2
            Optiononethird.Enabled = False
            Optiononethird.Value = False
            Optiontwothirds.Enabled = False
            Optiontwothirds.Value = False

            If Veteran_check_att.Value = 1 Then
                att_result.Caption = Fix(Unit.attack * (1 + veteranbonus))
            ElseIf Unit.Unitname = attchoice.Unitname Then
                att_result.Caption = Unit.attack
            End If
        Else
            thirds_check.Enabled = True
            If Unit.Unitname = attchoice.Unitname And Veteran_check_att.Value = 1 And Optiononethird.Value = True Then
                att_result.Caption = Fix(Unit.attack * (1 + veteranbonus) * 0.33334)
            ElseIf Unit.Unitname = attchoice.Unitname And Veteran_check_att.Value = 1 And Optiontwothirds.Value = True Then
                att_result.Caption = Fix(Unit.attack * (1 + veteranbonus) * 0.66667)
            ElseIf Unit.Unitname = attchoice.Unitname And Veteran_check_att.Value = 0 And Optiononethird.Value = True Then
                att_result.Caption = Fix(Unit.attack * (1) * 0.33334)
            ElseIf Unit.Unitname = attchoice.Unitname And Veteran_check_att.Value = 0 And Optiontwothirds.Value = True Then
                att_result.Caption = Fix(Unit.attack * (1) * 0.66667)
            ElseIf Unit.Unitname = attchoice.Unitname And Veteran_check_att.Value = 1 Then
                att_result.Caption = Fix(Unit.attack * (1 + veteranbonus))
            ElseIf Unit.Unitname = attchoice.Unitname Then
                att_result.Caption = Unit.attack
            End If
        End If
            If thirds_check.Enabled = False And thirds_check.Value = 1 Then
                thirds_check.Value = 0
            End If
    End If
    
Next
End Sub

Private Sub Citywalls_check_Click()

Set terrchoice = New Terrain
terrchoice.terrname = Getactivecombolistchoice(terr_list)

Set attchoice = New Unit
attchoice.Unitname = Getactivecombolistchoice(att_list)

Set defchoice = New Unit
defchoice.Unitname = Getactivecombolistchoice(def_list)

If Citywalls_check.Value = 0 Or attchoice.Unitname = "Bomber" Or attchoice.Unitname = "Artillery" Then

For Each Terrain In terrcollection
    If Terrain.terrname = terrchoice.terrname Then
        terr_result.Caption = Terrain.deffront
        defbonus = Terrain.defbonus
    End If
Next

For Each Unit In unitcollection
    If Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And def_f_check.Value = 1 Then
        def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + fortifybonus + defbonus))
    ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And def_f_check.Value = 0 Then
        def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + defbonus))
    ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And def_f_check.Value = 1 Then
        def_result.Caption = Fix(Unit.defend * (1 + fortifybonus + defbonus))
    ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And def_f_check.Value = 0 Then
        def_result.Caption = Fix(Unit.defend * (1 + defbonus))
    End If
Next

Else

For Each Unit In unitcollection
    If Unit.Unitname = defchoice.Unitname Then
        For Each terr In terrcollection
            If terr.terrname = terrchoice.terrname Then
                terr_result.Caption = terr.deffront
                defbonus = terr.defbonus
                If def_f_check.Value = 1 And Veteran_check_def.Value = 1 Then
                    def_result.Caption = Fix(Unit.defend * (((1 + veteranbonus + defbonus) * 3) + (fortifybonus)))
                ElseIf def_f_check.Value = 1 And Veteran_check_def.Value = 0 Then
                    def_result.Caption = Fix(Unit.defend * (((1 + defbonus) * 3) + (fortifybonus)))
                ElseIf def_f_check.Value = 0 And Veteran_check_def.Value = 1 Then
                    def_result.Caption = Fix(Unit.defend * (((1 + veteranbonus + defbonus) * 3)))
                ElseIf def_f_check.Value = 0 And Veteran_check_def.Value = 0 Then
                    def_result.Caption = Fix(Unit.defend * (((1 + defbonus) * 3)))
                End If
            End If
        Next
    End If
Next

End If
End Sub

Private Sub calc_Click()

On Error GoTo Errorhandler

intatt = CInt(att_result.Caption)
intdef = CInt(def_result.Caption)

inttotal = intatt + intdef
intattprob = intatt / inttotal
'intattprob = Round(intattprob, 4)

winchance_att.Text = CStr(intatt) + " / " + CStr(inttotal) + " = " + CStr(intattprob)
winchance = CDbl(winchance_att.Text)

winchance_def.Text = 1 - intattprob

If intattprob >= 0.5 And mute.Value = 0 Then
they_die.DoVerb 0
ElseIf intattprob < 0.5 And mute.Value = 0 Then
we_die.DoVerb 0
End If

Errorhandler:
If Err.Number = 13 Then
inttotal = 0
intatt = 0
intdef = 0

End If
Resume Next

End Sub

Private Sub def_f_check_Click()

If def_f_check.Value = 1 Then

Set terrchoice = New Terrain
terrchoice.terrname = Getactivecombolistchoice(terr_list)

For Each Terrain In terrcollection
    If Terrain.terrname = terrchoice.terrname Then
        defbonus = Terrain.defbonus
    End If
Next

Set defchoice = New Unit
defchoice.Unitname = Getactivecombolistchoice(def_list)

Set attchoice = New Unit
attchoice.Unitname = Getactivecombolistchoice(att_list)


For Each Unit In unitcollection
    If Unit.Unitname = defchoice.Unitname Then
        If (attchoice.Unitname = "Artillery" Or attchoice.Unitname = "Bomber") And Citywalls_check.Value = 1 Then
            If Veteran_check_def.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (1 + fortifybonus + veteranbonus + defbonus))
            Else
                def_result.Caption = Fix(Unit.defend * (1 + fortifybonus + defbonus))
            End If
        Else
            If Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + fortifybonus + veteranbonus + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + fortifybonus + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + defbonus) * 3) + (fortifybonus)))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + veteranbonus + defbonus) * 3) + (fortifybonus)))
            End If
        End If
    End If
Next

Else

Set terrchoice = New Terrain
terrchoice.terrname = Getactivecombolistchoice(terr_list)

For Each Terrain In terrcollection
    If Terrain.terrname = terrchoice.terrname Then
    defbonus = Terrain.defbonus
    End If
Next

Set defchoice = New Unit
defchoice.Unitname = Getactivecombolistchoice(def_list)

Set attchoice = New Unit
attchoice.Unitname = Getactivecombolistchoice(att_list)

For Each Unit In unitcollection
    If Unit.Unitname = defchoice.Unitname Then
        If (attchoice.Unitname = "Artillery" Or attchoice.Unitname = "Bomber") And Citywalls_check.Value = 1 Then
            If Veteran_check_def.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + defbonus))
            Else
                def_result.Caption = Fix(Unit.defend * (1 + defbonus))
            End If
        Else
            If Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + defbonus) * 3)))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * ((1 + veteranbonus + defbonus) * 3))
            End If
        End If
    End If
Next

End If
End Sub

Private Sub def_list_Click()

Set terrchoice = New Terrain
terrchoice.terrname = Getactivecombolistchoice(terr_list)

For Each Terrain In terrcollection
    If Terrain.terrname = terrchoice.terrname Then
        terr_result.Caption = Terrain.deffront
        defbonus = Terrain.defbonus
    End If
Next

Set defchoice = New Unit
defchoice.Unitname = Getactivecombolistchoice(def_list)

Set attchoice = New Unit
attchoice.Unitname = Getactivecombolistchoice(att_list)

For Each Unit In unitcollection
    If Unit.Unitname = defchoice.Unitname Then
        If (attchoice.Unitname = "Artillery" Or attchoice.Unitname = "Bomber") And Citywalls_check.Value = 1 Then
            If def_f_check.Value = 1 And Veteran_check_def.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (1 + fortifybonus + veteranbonus + defbonus))
            ElseIf (def_f_check.Value = 1 And Veteran_check_def.Value = 0) Or (def_f_check.Value = 0 And Veteran_check_def.Value = 1) Then
                def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + defbonus))
            Else
                def_result.Caption = Fix(Unit.defend * (1 + defbonus))
            End If
        Else

            If (defchoice.Unitname = "Trireme" Or defchoice.Unitname = "Sail" Or defchoice.Unitname = "Frigate" Or defchoice.Unitname = "Ironclad" Or defchoice.Unitname = "Cruiser" Or defchoice.Unitname = "Battleship" Or defchoice.Unitname = "Carrier" Or defchoice.Unitname = "Transport" Or defchoice.Unitname = "Submarine" Or defchoice.Unitname = "Fighter" Or defchoice.Unitname = "Bomber" Or defchoice.Unitname = "Nuclear") Then
                def_f_check.Enabled = False
            Else
                def_f_check.Enabled = True
            End If
            
            If Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And def_f_check.Value = 1 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + fortifybonus + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And ((Veteran_check_def.Value = 1 And def_f_check.Value = 0) Or (Veteran_check_def.Value = 0 And def_f_check.Value = 1)) And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And def_f_check.Value = 0 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And def_f_check.Value = 1 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + veteranbonus + defbonus) * 3) + (fortifybonus)))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And def_f_check.Value = 0 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + veteranbonus + defbonus) * 3)))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And def_f_check.Value = 1 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + defbonus) * 3) + (fortifybonus)))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And def_f_check.Value = 0 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + defbonus) * 3)))
            End If
        End If
        If def_f_check.Enabled = False And def_f_check.Value = 1 Then
            def_f_check.Value = 0
        End If
    End If
Next
End Sub
Private Sub Form_Load()

Set Militia = New Unit
Militia.Unitname = "Militia"
Militia.attack = 1
Militia.defend = 1
unitcollection.Add Militia

Set Phalanx = New Unit
Phalanx.Unitname = "Phalanx"
Phalanx.attack = 1
Phalanx.defend = 2
unitcollection.Add Phalanx

Set Cavalry = New Unit
Cavalry.Unitname = "Cavalry"
Cavalry.attack = 2
Cavalry.defend = 1
unitcollection.Add Cavalry

Set Legion = New Unit
Legion.Unitname = "Legion"
Legion.attack = 3
Legion.defend = 1
unitcollection.Add Legion

Set Chariot = New Unit
Chariot.Unitname = "Chariot"
Chariot.attack = 4
Chariot.defend = 1
unitcollection.Add Chariot

Set Trireme = New Unit
Trireme.Unitname = "Trireme"
Trireme.attack = 1
Trireme.defend = 1
unitcollection.Add Trireme

Set Catapult = New Unit
Catapult.Unitname = "Catapult"
Catapult.attack = 6
Catapult.defend = 1
unitcollection.Add Catapult

Set Knight = New Unit
Knight.Unitname = "Knight"
Knight.attack = 4
Knight.defend = 2
unitcollection.Add Knight

Set Sail = New Unit
Sail.Unitname = "Sail"
Sail.attack = 1
Sail.defend = 1
unitcollection.Add Sail

Set Ironclad = New Unit
Ironclad.Unitname = "Ironclad"
Ironclad.attack = 4
Ironclad.defend = 4
unitcollection.Add Ironclad

Set Musketeer = New Unit
Musketeer.Unitname = "Musketeer"
Musketeer.attack = 2
Musketeer.defend = 3
unitcollection.Add Musketeer

Set Frigate = New Unit
Frigate.Unitname = "Frigate"
Frigate.attack = 2
Frigate.defend = 2
unitcollection.Add Frigate

Set Cannon = New Unit
Cannon.Unitname = "Cannon"
Cannon.attack = 8
Cannon.defend = 1
unitcollection.Add Cannon

Set Cruiser = New Unit
Cruiser.Unitname = "Cruiser"
Cruiser.attack = 6
Cruiser.defend = 6
unitcollection.Add Cruiser

Set Transport = New Unit
Transport.Unitname = "Transport"
Transport.attack = 0
Transport.defend = 3
unitcollection.Add Transport

Set Fighter = New Unit
Fighter.Unitname = "Fighter"
Fighter.attack = 4
Fighter.defend = 2
unitcollection.Add Fighter

Set Bomber = New Unit
Bomber.Unitname = "Bomber"
Bomber.attack = 12
Bomber.defend = 1
unitcollection.Add Bomber

Set Carrier = New Unit
Carrier.Unitname = "Carrier"
Carrier.attack = 1
Carrier.defend = 12
unitcollection.Add Carrier

Set Rifleman = New Unit
Rifleman.Unitname = "Rifleman"
Rifleman.attack = 3
Rifleman.defend = 5
unitcollection.Add Rifleman

Set Submarine = New Unit
Submarine.Unitname = "Submarine"
Submarine.attack = 8
Submarine.defend = 2
unitcollection.Add Submarine

Set Battleship = New Unit
Battleship.Unitname = "Battleship"
Battleship.attack = 18
Battleship.defend = 12
unitcollection.Add Battleship

Set Armor = New Unit
Armor.Unitname = "Armor"
Armor.attack = 10
Armor.defend = 5
unitcollection.Add Armor

Set MechInf = New Unit
MechInf.Unitname = "Mech Inf"
MechInf.attack = 6
MechInf.defend = 6
unitcollection.Add MechInf

Set Artillery = New Unit
Artillery.Unitname = "Artillery"
Artillery.attack = 12
Artillery.defend = 2
unitcollection.Add Artillery

Set Nuclear = New Unit
Nuclear.Unitname = "Nuclear"
Nuclear.attack = 99
Nuclear.defend = 0
unitcollection.Add Nuclear

Set Plains = New Terrain
Plains.terrname = "Plains"
Plains.defbonus = 0
Plains.deffront = "0%"
terrcollection.Add Plains

Set Grassland = New Terrain
Grassland.terrname = "Grassland"
Grassland.defbonus = 0
Grassland.deffront = "0%"
terrcollection.Add Grassland

Set Desert = New Terrain
Desert.terrname = "Desert"
Desert.defbonus = 0
Desert.deffront = "0%"
terrcollection.Add Desert

Set Arctic = New Terrain
Arctic.terrname = "Arctic"
Arctic.defbonus = 0
Arctic.deffront = "0%"
terrcollection.Add Arctic

Set Tundra = New Terrain
Tundra.terrname = "Tundra"
Tundra.defbonus = 0
Tundra.deffront = "0%"
terrcollection.Add Tundra

Set River = New Terrain
River.terrname = "River"
River.defbonus = 0.5
River.deffront = "50%"
terrcollection.Add River

Set Forest = New Terrain
Forest.terrname = "Forest"
Forest.defbonus = 0.5
Forest.deffront = "50%"
terrcollection.Add Forest

Set Swamp = New Terrain
Swamp.terrname = "Swamp"
Swamp.defbonus = 0.5
Swamp.deffront = "50%"
terrcollection.Add Swamp

Set Jungle = New Terrain
Jungle.terrname = "Jungle"
Jungle.defbonus = 0.5
Jungle.deffront = "50%"
terrcollection.Add Jungle

Set Hills = New Terrain
Hills.terrname = "Hills"
Hills.defbonus = 1
Hills.deffront = "100%"
terrcollection.Add Hills

Set Mountains = New Terrain
Mountains.terrname = "Mountains"
Mountains.defbonus = 2
Mountains.deffront = "200%"
terrcollection.Add Mountains


att_list.AddItem "Militia", 0
att_list.AddItem "Phalanx", 1
att_list.AddItem "Cavalry", 2
att_list.AddItem "Legion", 3
att_list.AddItem "Chariot", 4
att_list.AddItem "Trireme", 5
att_list.AddItem "Catapult", 6
att_list.AddItem "Knight", 7
att_list.AddItem "Sail", 8
att_list.AddItem "Ironclad", 9
att_list.AddItem "Musketeer", 10
att_list.AddItem "Frigate", 11
att_list.AddItem "Cannon", 12
att_list.AddItem "Cruiser", 13
att_list.AddItem "Transport", 14
att_list.AddItem "Fighter", 15
att_list.AddItem "Bomber", 16
att_list.AddItem "Carrier", 17
att_list.AddItem "Rifleman", 18
att_list.AddItem "Submarine", 19
att_list.AddItem "Battleship", 20
att_list.AddItem "Armor", 21
att_list.AddItem "Mech Inf", 22
att_list.AddItem "Artillery", 23
att_list.AddItem "Nuclear", 24


def_list.AddItem "Militia", 0
def_list.AddItem "Phalanx", 1
def_list.AddItem "Cavalry", 2
def_list.AddItem "Legion", 3
def_list.AddItem "Chariot", 4
def_list.AddItem "Trireme", 5
def_list.AddItem "Catapult", 6
def_list.AddItem "Knight", 7
def_list.AddItem "Sail", 8
def_list.AddItem "Ironclad", 9
def_list.AddItem "Musketeer", 10
def_list.AddItem "Frigate", 11
def_list.AddItem "Cannon", 12
def_list.AddItem "Cruiser", 13
def_list.AddItem "Transport", 14
def_list.AddItem "Fighter", 15
def_list.AddItem "Bomber", 16
def_list.AddItem "Carrier", 17
def_list.AddItem "Rifleman", 18
def_list.AddItem "Submarine", 19
def_list.AddItem "Battleship", 20
def_list.AddItem "Armor", 21
def_list.AddItem "Mech Inf", 22
def_list.AddItem "Artillery", 23
def_list.AddItem "Nuclear", 24

terr_list.AddItem "Plains", 0
terr_list.AddItem "Grassland", 1
terr_list.AddItem "Desert", 2
terr_list.AddItem "Arctic", 3
terr_list.AddItem "Tundra", 4

terr_list.AddItem "River", 5
terr_list.AddItem "Forest", 6
terr_list.AddItem "Swamp", 7
terr_list.AddItem "Jungle", 8

terr_list.AddItem "Hills", 9
terr_list.AddItem "Mountains", 10
End Sub
Private Sub Optiononethird_Click()

Set attchoice = New Unit
attchoice.Unitname = Getactivecombolistchoice(att_list)

For Each Unit In unitcollection
    If Unit.Unitname = attchoice.Unitname And Veteran_check_att.Value = 1 Then
        att_result.Caption = Fix(Unit.attack * 1.5 * 0.33334)
    ElseIf Unit.Unitname = attchoice.Unitname And Veteran_check_att.Value = 0 Then
        att_result.Caption = Fix(Unit.attack * 0.33334)
    End If
Next

End Sub
Private Sub Optiontwothirds_Click()

Set attchoice = New Unit
attchoice.Unitname = Getactivecombolistchoice(att_list)

For Each Unit In unitcollection
    If Unit.Unitname = attchoice.Unitname And Veteran_check_att.Value = 1 Then
        att_result.Caption = Fix(Unit.attack * 1.5 * 0.66667)
    ElseIf Unit.Unitname = attchoice.Unitname And Veteran_check_att.Value = 0 Then
        att_result.Caption = Fix(Unit.attack * 0.66667)
    End If
Next

End Sub
Private Sub terr_list_Click()

Set terrchoice = New Terrain
terrchoice.terrname = Getactivecombolistchoice(terr_list)

For Each Terrain In terrcollection
    If Terrain.terrname = terrchoice.terrname Then
        terr_result.Caption = Terrain.deffront
        defbonus = Terrain.defbonus
    End If
Next

Set defchoice = New Unit
defchoice.Unitname = Getactivecombolistchoice(def_list)

Set attchoice = New Unit
attchoice.Unitname = Getactivecombolistchoice(att_list)

For Each Unit In unitcollection
    If Unit.Unitname = defchoice.Unitname Then
        If (attchoice.Unitname = "Artillery" Or attchoice.Unitname = "Bomber") And Citywalls_check.Value = 1 Then
            If def_f_check.Value = 1 And Veteran_check_def.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (1 + fortifybonus + veteranbonus + defbonus))
            ElseIf (def_f_check.Value = 1 And Veteran_check_def.Value = 0) Or (def_f_check.Value = 0 And Veteran_check_def.Value = 1) Then
                def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + defbonus))
            Else
                def_result.Caption = Fix(Unit.defend * (1 + defbonus))
            End If
        Else
            If Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And def_f_check.Value = 1 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + fortifybonus + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And ((Veteran_check_def.Value = 1 And def_f_check.Value = 0) Or (Veteran_check_def.Value = 0 And def_f_check.Value = 1)) And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And def_f_check.Value = 0 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And def_f_check.Value = 1 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + veteranbonus + defbonus) * 3) + (fortifybonus)))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 1 And def_f_check.Value = 0 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + veteranbonus + defbonus) * 3)))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And def_f_check.Value = 1 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + defbonus) * 3) + (fortifybonus)))
            ElseIf Unit.Unitname = defchoice.Unitname And Veteran_check_def.Value = 0 And def_f_check.Value = 0 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + defbonus) * 3)))
            End If
        End If
    End If
Next

End Sub
Private Sub thirds_check_Click()

If thirds_check.Value = 1 Then
    Optiononethird.Enabled = True
    Optiontwothirds.Enabled = True
Else
    Optiononethird.Value = False
    Optiontwothirds.Value = False
    Optiononethird.Enabled = False
    Optiontwothirds.Enabled = False
End If

If Veteran_check_att.Value = 1 Then
    Set attchoice = New Unit
    attchoice.Unitname = Getactivecombolistchoice(att_list)

    For Each Unit In unitcollection
        If Unit.Unitname = attchoice.Unitname Then
            att_result.Caption = Fix(Unit.attack * 1.5)
        End If
    Next
Else
    Set attchoice = New Unit
    attchoice.Unitname = Getactivecombolistchoice(att_list)

    For Each Unit In unitcollection
        If Unit.Unitname = attchoice.Unitname Then
            att_result.Caption = Fix(Unit.attack)
        End If
    Next
    
End If
End Sub
Private Sub Veteran_check_att_Click()

If Veteran_check_att.Value = 1 Then
    Set attchoice = New Unit
    attchoice.Unitname = Getactivecombolistchoice(att_list)
    
    For Each Unit In unitcollection
        If Unit.Unitname = attchoice.Unitname Then
            If attchoice.Unitname = "Trireme" Or attchoice.Unitname = "Sail" Or attchoice.Unitname = "Ironclad" Or attchoice.Unitname = "Frigate" Or attchoice.Unitname = "Cruiser" Or attchoice.Unitname = "Submarine" Or attchoice.Unitname = "Battleship" Or attchoice.Unitname = "Transport" Or attchoice.Unitname = "Carrier" Or attchoice.Unitname = "Fighter" Or attchoice.Unitname = "Bomber" Or attchoice.Unitname = "Nuclear" Then
                Optiononethird.Enabled = False
                Optiononethird.Value = False
                Optiontwothirds.Enabled = False
                Optiontwothirds.Value = False
                thirds_check.Enabled = False
                'thirds_check.Value = 0
                att_result.Caption = Fix(Unit.attack * (1 + veteranbonus))
            Else
                thirds_check.Enabled = True
                If Unit.Unitname = attchoice.Unitname And Optiononethird.Value = True Then
                    att_result.Caption = Fix(Unit.attack * (1 + veteranbonus) * 0.33334)
                ElseIf Unit.Unitname = attchoice.Unitname And Optiontwothirds.Value = True Then
                    att_result.Caption = Fix(Unit.attack * (1 + veteranbonus) * 0.66667)
                ElseIf Unit.Unitname = attchoice.Unitname Then
                    att_result.Caption = Fix(Unit.attack * (1 + veteranbonus))
                End If
            End If
        End If
    Next
    
Else

    Set attchoice = New Unit
    attchoice.Unitname = Getactivecombolistchoice(att_list)
    
    For Each Unit In unitcollection
        If Unit.Unitname = attchoice.Unitname Then
            If attchoice.Unitname = "Trireme" Or attchoice.Unitname = "Sail" Or attchoice.Unitname = "Ironclad" Or attchoice.Unitname = "Frigate" Or attchoice.Unitname = "Cruiser" Or attchoice.Unitname = "Submarine" Or attchoice.Unitname = "Battleship" Or attchoice.Unitname = "Transport" Or attchoice.Unitname = "Carrier" Or attchoice.Unitname = "Fighter" Or attchoice.Unitname = "Bomber" Or attchoice.Unitname = "Nuclear" Then
                Optiononethird.Enabled = False
                Optiononethird.Value = False
                Optiontwothirds.Enabled = False
                Optiontwothirds.Value = False
                thirds_check.Enabled = False
                att_result.Caption = Fix(Unit.attack)
            Else
                thirds_check.Enabled = True
                If Unit.Unitname = attchoice.Unitname And Optiononethird.Value = True Then
                    att_result.Caption = Fix(Unit.attack * 1 * 0.33334)
                ElseIf Unit.Unitname = attchoice.Unitname And Optiontwothirds.Value = True Then
                    att_result.Caption = Fix(Unit.attack * 1 * 0.66667)
                ElseIf Unit.Unitname = attchoice.Unitname Then
                    att_result.Caption = Fix(Unit.attack)
                End If
            End If
        End If
    Next
End If
End Sub

Private Sub Veteran_check_def_Click()

If Veteran_check_def.Value = 1 Then
    Set defchoice = New Unit
    Set terrchoice = New Terrain
    Set attchoice = New Unit
    defchoice.Unitname = Getactivecombolistchoice(def_list)
    terrchoice.terrname = Getactivecombolistchoice(terr_list)
    attchoice.Unitname = Getactivecombolistchoice(att_list)

    For Each terr In terrcollection
        If terr.terrname = terrchoice.terrname Then
            defbonus = terr.defbonus
        End If
    Next

For Each Unit In unitcollection
    If Unit.Unitname = defchoice.Unitname Then
        If (attchoice.Unitname = "Artillery" Or attchoice.Unitname = "Bomber") And Citywalls_check.Value = 1 Then
            If def_f_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (1 + fortifybonus + veteranbonus + defbonus))
            Else
                def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + defbonus))
            End If
        Else
            If Unit.Unitname = defchoice.Unitname And def_f_check.Value = 1 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + defbonus + fortifybonus))
            ElseIf Unit.Unitname = defchoice.Unitname And def_f_check.Value = 0 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + veteranbonus + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And def_f_check.Value = 1 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + veteranbonus + defbonus) * 3) + (fortifybonus)))
            ElseIf Unit.Unitname = defchoice.Unitname And def_f_check.Value = 0 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + veteranbonus + defbonus) * 3)))
            End If
        End If
    End If
Next

Else

Set terrchoice = New Terrain
terrchoice.terrname = Getactivecombolistchoice(terr_list)

For Each Terrain In terrcollection
    If Terrain.terrname = terrchoice.terrname Then
        defbonus = Terrain.defbonus
    End If
Next

Set defchoice = New Unit
defchoice.Unitname = Getactivecombolistchoice(def_list)

For Each Unit In unitcollection
    If Unit.Unitname = defchoice.Unitname Then
        If (attchoice.Unitname = "Artillery" Or attchoice.Unitname = "Bomber") And Citywalls_check.Value = 1 Then
            If def_f_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (1 + fortifybonus + defbonus))
            Else
                def_result.Caption = Fix(Unit.defend * (1 + defbonus))
            End If
        Else
            If Unit.Unitname = defchoice.Unitname And def_f_check.Value = 1 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + defbonus + fortifybonus))
            ElseIf Unit.Unitname = defchoice.Unitname And def_f_check.Value = 0 And Citywalls_check.Value = 0 Then
                def_result.Caption = Fix(Unit.defend * (1 + defbonus))
            ElseIf Unit.Unitname = defchoice.Unitname And def_f_check.Value = 1 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + defbonus) * 3) + (fortifybonus)))
            ElseIf Unit.Unitname = defchoice.Unitname And def_f_check.Value = 0 And Citywalls_check.Value = 1 Then
                def_result.Caption = Fix(Unit.defend * (((1 + defbonus) * 3)))
            End If
        End If
    End If
Next
End If
End Sub
