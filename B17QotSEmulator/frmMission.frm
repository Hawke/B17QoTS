VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMission 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mission:"
   ClientHeight    =   10125
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMissionSelection 
      Caption         =   "Mission"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   163
      Top             =   120
      Width           =   5055
      Begin VB.Label lblBombRun 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   179
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblBomberName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   178
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblBomberMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   177
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblTargetName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   176
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblTargetType 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   175
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblFormationPos 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   174
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblSquadronPos 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   173
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Bomber"
         Height          =   255
         Left            =   120
         TabIndex        =   172
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Mission"
         Height          =   255
         Left            =   2880
         TabIndex        =   171
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Target"
         Height          =   255
         Left            =   2880
         TabIndex        =   170
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Type"
         Height          =   255
         Left            =   2880
         TabIndex        =   169
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Formation Pos"
         Height          =   255
         Left            =   120
         TabIndex        =   168
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Squadron Pos"
         Height          =   255
         Left            =   120
         TabIndex        =   167
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblBombRunLabel 
         Caption         =   "Bomb %"
         Height          =   255
         Left            =   2880
         TabIndex        =   166
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   165
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   164
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdInterrupt 
      Caption         =   "Enter Next Zone"
      Height          =   375
      Left            =   8280
      TabIndex        =   162
      Top             =   9120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame fraWave 
      Caption         =   " Wave Selection "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   7800
      TabIndex        =   141
      Top             =   5640
      Width           =   2895
      Begin VB.CheckBox chkEvadeFighters 
         Caption         =   "Evade Fighters"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   142
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label lblToHit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   161
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label lblToHit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   160
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label lblToHit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   159
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label lblToHit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   158
         Top             =   960
         Width           =   405
      End
      Begin VB.Label lblToHit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   157
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblToHit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   156
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblType 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   155
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblType 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   154
         Top             =   600
         Width           =   645
      End
      Begin VB.Label lblType 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   153
         Top             =   960
         Width           =   645
      End
      Begin VB.Label lblType 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   152
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lblType 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   151
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lblType 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   150
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label lblPosition 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   149
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblPosition 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   148
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label lblPosition 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   147
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblPosition 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   146
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label lblPosition 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   145
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblPosition 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   144
         Top             =   2040
         Width           =   1200
      End
      Begin VB.Label lblMiscWave 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   143
         Top             =   2400
         Width           =   2655
      End
   End
   Begin VB.Frame fraGuns 
      Caption         =   " Guns "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   5400
      TabIndex        =   109
      Top             =   5640
      Width           =   2175
      Begin VB.CheckBox chkSpray 
         Enabled         =   0   'False
         Height          =   255
         Index           =   10
         Left            =   600
         TabIndex        =   120
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox chkSpray 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   119
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkSpray 
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   118
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkSpray 
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   117
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkSpray 
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   116
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkSpray 
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   115
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkSpray 
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   114
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkSpray 
         Enabled         =   0   'False
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   113
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox chkSpray 
         Enabled         =   0   'False
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   112
         Top             =   2760
         Width           =   255
      End
      Begin VB.CheckBox chkSpray 
         Enabled         =   0   'False
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   111
         Top             =   3120
         Width           =   255
      End
      Begin VB.CheckBox chkSwapAmmo 
         Caption         =   "Swap Ammo"
         Height          =   195
         Left            =   120
         TabIndex        =   110
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lblGunAmmo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   140
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblGunName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mid-Upper"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   139
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblGunName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tail Turret"
         Height          =   255
         Index           =   10
         Left            =   960
         TabIndex        =   138
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblGunName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ball Turret"
         Height          =   255
         Index           =   9
         Left            =   960
         TabIndex        =   137
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblGunName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Top Turret"
         Height          =   255
         Index           =   8
         Left            =   960
         TabIndex        =   136
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lblGunName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stbd Waist"
         Height          =   255
         Index           =   7
         Left            =   960
         TabIndex        =   135
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblGunName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Port Waist"
         Height          =   255
         Index           =   6
         Left            =   960
         TabIndex        =   134
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblGunName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Radio Room"
         Height          =   255
         Index           =   5
         Left            =   960
         TabIndex        =   133
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblGunName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stbd Cheek"
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   132
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblGunName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Port Cheek"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   131
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblGunName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nose"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   130
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblGunAmmo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   129
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label lblGunAmmo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   128
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label lblGunAmmo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   127
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label lblGunAmmo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   126
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblGunAmmo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   125
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label lblGunAmmo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   124
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblGunAmmo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   123
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label lblGunAmmo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   122
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label lblGunAmmo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   121
         Top             =   600
         Width           =   360
      End
   End
   Begin VB.Frame fraGauges 
      Caption         =   "Gauges"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   70
      Top             =   2040
      Width           =   5055
      Begin MSComctlLib.ProgressBar proFuelGauge 
         Height          =   255
         Left            =   3720
         TabIndex        =   71
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar proOilPressure 
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   72
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   4
      End
      Begin MSComctlLib.ProgressBar proOilPressure 
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   73
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   4
      End
      Begin MSComctlLib.ProgressBar proOilPressure 
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   74
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   4
      End
      Begin MSComctlLib.ProgressBar proOilPressure 
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   75
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   4
      End
      Begin VB.Label lblEnvironment 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Environment"
         Height          =   255
         Left            =   2760
         TabIndex        =   182
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblBombSight 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bomb Sight"
         Height          =   255
         Left            =   2760
         TabIndex        =   78
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblBombBayDoors 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BombBayDoor"
         Height          =   255
         Left            =   2760
         TabIndex        =   181
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Weather"
         Height          =   255
         Left            =   2760
         TabIndex        =   103
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblDirection 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3480
         TabIndex        =   108
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Bearing"
         Height          =   255
         Left            =   2760
         TabIndex        =   107
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblLocation 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "zone / terrain"
         Height          =   255
         Left            =   3480
         TabIndex        =   106
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Location"
         Height          =   255
         Left            =   2760
         TabIndex        =   105
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblWeather 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3480
         TabIndex        =   104
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblAltitude 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3480
         TabIndex        =   102
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Altitude"
         Height          =   255
         Left            =   2760
         TabIndex        =   101
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Fuel Gauge"
         Height          =   255
         Left            =   2760
         TabIndex        =   100
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Oil Pressure"
         Height          =   255
         Left            =   1440
         TabIndex        =   99
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblRudder 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stbd Rudder"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   98
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblRudder 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Port Rudder"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   97
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblEngine 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Engine #3"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   96
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblEngine 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Engine #2"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   95
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblEngine 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Engine #1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   94
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblElevator 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stbd Elevator"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   93
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblAileron 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stbd Aileron"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   92
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblWingFlap 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stbd Wing Flap"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   91
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblEngine 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Engine #4"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   90
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label lblElevator 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Port Elevator"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   89
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblAileron 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Port Aileron"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   88
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblWingFlap 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Port Wing Flap"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   87
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Damage"
         Height          =   255
         Left            =   2760
         TabIndex        =   86
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblPeckhamPoints 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3720
         TabIndex        =   85
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label lblBrake 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Brake"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblThirdWheel 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tail Wheel"
         Height          =   255
         Left            =   1440
         TabIndex        =   83
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblLandingGear 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Landing Gear"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblIntercom 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intercom"
         Height          =   255
         Left            =   3960
         TabIndex        =   81
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblRadio 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Radio"
         Height          =   255
         Left            =   3960
         TabIndex        =   80
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblRubberRafts 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rubber Raft"
         Height          =   255
         Left            =   3960
         TabIndex        =   79
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblNavigationEquipment 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Navig. Sys."
         Height          =   255
         Left            =   3960
         TabIndex        =   77
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblBombsOnBoard 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bombs Aboard"
         Height          =   255
         Left            =   2760
         TabIndex        =   76
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.Frame fraCrew 
      Caption         =   " Crew "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.Label Label2 
         Caption         =   "Kills"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   69
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   68
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   12
         Left            =   3120
         TabIndex        =   67
         Top             =   4560
         Width           =   600
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Tail Gunner"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   66
         Top             =   4560
         Width           =   1440
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   12
         Left            =   3840
         TabIndex        =   65
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   12
         Left            =   4680
         TabIndex        =   64
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   11
         Left            =   3120
         TabIndex        =   63
         Top             =   4200
         Width           =   600
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Tail Gunner"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   11
         Left            =   1560
         TabIndex        =   62
         Top             =   4200
         Width           =   1440
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   11
         Left            =   3840
         TabIndex        =   61
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   11
         Left            =   4680
         TabIndex        =   60
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   13
         Left            =   3120
         TabIndex        =   59
         Top             =   4920
         Width           =   600
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Tail Gunner"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   13
         Left            =   1560
         TabIndex        =   58
         Top             =   4920
         Width           =   1440
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   13
         Left            =   3840
         TabIndex        =   57
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   13
         Left            =   4680
         TabIndex        =   56
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Ammo Stocker"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   55
         Top             =   4920
         Width           =   1275
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Mid-Upper Gunner"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   54
         Top             =   3120
         Width           =   1305
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Nose Gunner"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   53
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "RKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   52
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "RKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   51
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "RKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   50
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "RKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   49
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Pilot"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   48
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   47
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   3120
         TabIndex        =   46
         Top             =   3840
         Width           =   600
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   45
         Top             =   960
         Width           =   600
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   44
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   43
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   42
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   41
         Top             =   2400
         Width           =   600
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   3120
         TabIndex        =   40
         Top             =   2760
         Width           =   600
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   39
         Top             =   3120
         Width           =   600
      End
      Begin VB.Label lblAirmanStatus 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   3120
         TabIndex        =   38
         Top             =   3480
         Width           =   600
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Tail Gunner"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   10
         Left            =   1560
         TabIndex        =   37
         Top             =   3840
         Width           =   1440
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Co-Pilot"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   36
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Bombardier"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   35
         Top             =   1320
         Width           =   1440
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Navigator"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   34
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Radio Operator"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   33
         Top             =   2040
         Width           =   1440
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "P Waist Gunner"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   32
         Top             =   2400
         Width           =   1440
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "S Waist Gunner"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   31
         Top             =   2760
         Width           =   1440
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Engineer"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   30
         Top             =   3120
         Width           =   1440
      End
      Begin VB.Label lblAirmanName 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Ball Gunner"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   9
         Left            =   1560
         TabIndex        =   29
         Top             =   3480
         Width           =   1440
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Ball Gunner"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   28
         Top             =   3480
         Width           =   1305
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Engineer"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Stbd Waist Gunner"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   26
         Top             =   4200
         Width           =   1350
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Port Waist Gunner"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   25
         Top             =   3840
         Width           =   1305
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Radio Operator"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   1320
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Navigator"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Bombardier"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Co-Pilot"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Tail Gunner"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   20
         Top             =   4560
         Width           =   1305
      End
      Begin VB.Label lblCrewPosition 
         AutoSize        =   -1  'True
         Caption         =   "Pilot"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "PMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "CPMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "BMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   16
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "NMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   15
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "RMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   14
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "PWMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   6
         Left            =   3840
         TabIndex        =   13
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "SWMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   12
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "EMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   11
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "BallMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   9
         Left            =   3840
         TabIndex        =   10
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblAirmanMission 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TMission"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   9
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Mission"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "RKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   5
         Left            =   4680
         TabIndex        =   7
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "PWKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   6
         Left            =   4680
         TabIndex        =   6
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "SWKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   5
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "EKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   8
         Left            =   4680
         TabIndex        =   4
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "BallKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   9
         Left            =   4680
         TabIndex        =   3
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblAirmanKills 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TKills"
         DataSource      =   "Data1"
         Height          =   255
         Index           =   10
         Left            =   4680
         TabIndex        =   2
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin RichTextLib.RichTextBox rtbMessages 
      Height          =   3735
      Left            =   120
      TabIndex        =   180
      Top             =   6120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6588
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmMission.frx":0000
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "&Index"
      End
      Begin VB.Menu MenuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' frmMission.frm
'
' @author Preston V. McMurry III, http://www.prestonm.com
' @copyright (C) Copyright 2002, 2010 by Preston V. McMurry III, http://www.prestonm.com
'
' *****************************************************************************
'
' This file is part of B17QotS, the "B-17: Queen of the Skies" Emulator.
'
' B17QotS is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' B17QotS is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with B17QotS. If not, see <http://www.gnu.org/licenses/>.
'******************************************************************************

'******************************************************************************
' START HERE
'******************************************************************************

Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private gintRemovalsRemaining As Integer
Private gintCurrGun As Integer
Private gintCurrTarget As Integer

' Dynamically created run-time controls. In order for the controls to respond
' to events, they must be globally declared using the WithEvents key word.

Private WithEvents cmdTakeOff As VB.CommandButton
Attribute cmdTakeOff.VB_VarHelpID = -1
Private WithEvents fraTemp As VB.Frame
Attribute fraTemp.VB_VarHelpID = -1
Private WithEvents optOneTemp As VB.OptionButton
Attribute optOneTemp.VB_VarHelpID = -1
Private WithEvents optTwoTemp As VB.OptionButton
Attribute optTwoTemp.VB_VarHelpID = -1

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    MsgBox "DOWN: KeyCodei = '" & KeyCode & "', Shift = '" & Shift & "'"
    
'    If KeyCode = vbKeyShift _
'    Or KeyCode = vbKeyControl _
'    Or KeyCode = vbKeyMenu Then
    If KeyCode = vbKeyControl Then
        ' Record whether the user pressed shift, control or alt.
        pintDoo(1) = KeyCode
    ElseIf KeyCode = vbKeyB Then
        ' Record any other keystroke.
        pintDoo(2) = KeyCode
    End If
        
    If pintDoo(1) = vbKeyControl _
    And pintDoo(2) = vbKeyB Then
        MsgBox "User pressed CTRL-B"
        pintDoo(1) = 0
        pintDoo(2) = 0
        pblnDoBailOrCrash = True
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    MsgBox "PRESS: KeyAscii = '" & KeyAscii & "'"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    ' If a key is released, then the current state is reset.
'    If KeyCode = vbKeyShift _
'    Or KeyCode = vbKeyControl _
'    Or KeyCode = vbKeyMenu Then
    If KeyCode = vbKeyControl Then
        pintDoo(1) = 0
    ElseIf KeyCode = vbKeyB Then
        pintDoo(2) = 0
    End If

End Sub

'******************************************************************************
' Form_Load()
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Launch the form, initialize the controls, then wait for the bomber
'         to takeoff.
'******************************************************************************
Private Sub Form_Load()

    'CenterForm Me
    
    ' Fiddle the form bottom, as adding a menu bar otherwise seems to
    ' randomly cut off the bottom of the form
    frmMission.Height = rtbMessages.Top + rtbMessages.Height + 880
    
    Me.Caption = "Mission: " & Mission.TargetName

    Call InitializeMissionInfo
    Call InitializeGauges
    Call InitializeCrew
    Call InitializeGuns
    
    cmdInterrupt.Visible = False
    
    ' This control is only temporarily needed prior to take off, so dynamically
    ' create it.
    
    Set cmdTakeOff = Me.Controls.Add("VB.CommandButton", "cmdTakeOff", Me)
    cmdTakeOff.Width = cmdInterrupt.Width
    cmdTakeOff.Left = cmdInterrupt.Left
    cmdTakeOff.Height = cmdInterrupt.Height
    cmdTakeOff.Top = cmdInterrupt.Top
    cmdTakeOff.Caption = "Take Off"
    cmdTakeOff.Visible = True
    cmdTakeOff.Enabled = True
    
End Sub

'******************************************************************************
' cmdTakeOff_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Launch the mission engine.
'******************************************************************************
Private Sub cmdTakeOff_Click()

    ' The cmdTakeOff control is dynamically created in the Form_Load()
    ' routine. The control must exist for this routine to work. If the user
    ' clicked the button, he decided to take off. Reveal the next step button,
    ' delete the take off button, then launch the mission engine.

    cmdInterrupt.Visible = True

    Me.Controls.Remove ("cmdTakeOff")
            
    ' Launch the mission engine.
            
    Call FlyMission
    
    ' Display final status.
    
    Call RefreshMissionForm
    
    Call Interrupt(cmdInterrupt, FINISH_MISSION)

    ' User clicked the next action button.
    
    Call EndMission
    
    Call Interrupt(cmdInterrupt, EXIT_MISSION)

    frmMainMenu.Show

    Unload Me

End Sub

'******************************************************************************
' RefreshMissionForm
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Update the displayed data. This routine should be called after each
'        combat phase.
'******************************************************************************
Private Sub RefreshMissionForm()

    Call RefreshMissionInfo
    Call RefreshGauges
    Call RefreshCrew
    Call RefreshGuns
    
End Sub

'******************************************************************************
' InitializeMissionInfo
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub InitializeMissionInfo()

    lblBomberName.Caption = Bomber.Name
    lblBomberMission.Caption = Bomber.Mission
    lblTargetName.Caption = Mission.TargetName
    lblTargetType.Caption = Mission.TargetType
    lblFormationPos.Caption = Bomber.FormationPos
    lblSquadronPos.Caption = Bomber.SquadronPos
    
    If Bomber.BomberModel = YB40 Then
        lblBombRunLabel.Visible = False
        lblBombRun.Visible = False
    Else
        lblBombRun.Caption = ""
    End If
    
    lblDate.Caption = frmMainMenu.cboMonth.Text & ", " & frmMainMenu.cboYear.Text

End Sub

'******************************************************************************
' RefreshMissionInfo
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub RefreshMissionInfo()

' TODO: if target switching is allowed while mission is in progress
'    lblTargetName.Caption = Mission.TargetName
'    lblTargetType.Caption = Mission.TargetType
    
    If Bomber.InFormation = False Then
        lblFormationPos.Caption = "Out"
        If lblSquadronPos.Caption <> "Abort" Then
            lblSquadronPos.Caption = ""
        End If
    Else
        lblFormationPos.Caption = Bomber.FormationPos
        lblSquadronPos.Caption = Bomber.SquadronPos
    End If
    
    ' lblBombRun.Caption is updated, once, in TargetZoneCombat().

End Sub

'******************************************************************************
' InitializeGauges
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub InitializeGauges()
    
    lblWingFlap(PORT_SIDE).BackColor = PaleGreen()
    lblWingFlap(STBD_SIDE).BackColor = PaleGreen()
   
    lblAileron(PORT_SIDE).BackColor = PaleGreen()
    lblAileron(STBD_SIDE).BackColor = PaleGreen()
   
    lblElevator(PORT_SIDE).BackColor = PaleGreen()
    lblElevator(STBD_SIDE).BackColor = PaleGreen()
   
    lblRudder(PORT_SIDE).BackColor = PaleGreen()
    
    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM _
    Or Bomber.BomberModel = AVRO_LANCASTER Then
        lblRudder(STBD_SIDE).BackColor = PaleGreen()
    Else
        lblRudder(PORT_SIDE).Caption = "Rudder"
        lblRudder(STBD_SIDE).Visible = False
    End If
   
    lblLandingGear.BackColor = PaleGreen()
    
    lblThirdWheel.BackColor = PaleGreen()
    
    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
        lblThirdWheel.Caption = "Nose Wheel"
    End If
    
    lblBrake.BackColor = PaleGreen()
    
    lblEngine(1).BackColor = PaleGreen()
    lblEngine(2).BackColor = PaleGreen()
    lblEngine(3).BackColor = PaleGreen()
    lblEngine(4).BackColor = PaleGreen()
    
    proOilPressure(1).Value = 4
    proOilPressure(2).Value = 4
    proOilPressure(3).Value = 4
    proOilPressure(4).Value = 4
    
    Call ColorOilPressureGauge(1)
    Call ColorOilPressureGauge(2)
    Call ColorOilPressureGauge(3)
    Call ColorOilPressureGauge(4)
    
    lblDirection.Caption = "Outbound"
    
    lblAltitude.Caption = CStr(Bomber.Altitude)
    
    lblLocation.Caption = "Zone " & Bomber.CurrentZone & " / " & _
                          Mission.Zone(Bomber.CurrentZone).Terrain

    lblWeather.Caption = WeatherText(Mission.Zone(Bomber.CurrentZone).Weather)

    lblEnvironment.BackColor = PaleGreen()
    
    If Bomber.BomberModel = YB40 Then
        lblBombSight.Visible = False
        lblBombBayDoors.Visible = False
        lblBombsOnBoard.Visible = False
    Else
        lblBombSight.BackColor = PaleGreen()
        lblBombBayDoors.BackColor = PaleGreen()
        lblBombsOnBoard.BackColor = PaleYellow()
    End If
    
    lblNavigationEquipment.BackColor = PaleGreen()
    lblRadio.BackColor = PaleGreen()
    lblIntercom.BackColor = PaleGreen()
    lblRubberRafts.BackColor = PaleGreen()

    proFuelGauge.Max = Bomber.FuelPoints
    proFuelGauge.Value = Bomber.FuelPoints
    Call ColorFuelGauge

    lblPeckhamPoints.Caption = Damage.PeckhamPoints
    Call ColorPeckhamPointsGauge

End Sub

'******************************************************************************
' RefreshGauges
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Update display due to any damage that may have occured.
'******************************************************************************
Private Sub RefreshGauges()

    If Damage.WingFlap(PORT_SIDE) = True _
    Or Damage.WingFlapControls = True Then
        lblWingFlap(PORT_SIDE).BackColor = PaleRed()
    End If
    
    If Damage.WingFlap(STBD_SIDE) = True _
    Or Damage.WingFlapControls = True Then
        lblWingFlap(STBD_SIDE).BackColor = PaleRed()
    End If
    
    If Damage.Aileron(PORT_SIDE) = True _
    Or Damage.AileronControls = True Then
        lblAileron(PORT_SIDE).BackColor = PaleRed()
    End If
    
    If Damage.Aileron(STBD_SIDE) = True _
    Or Damage.AileronControls = True Then
        lblAileron(STBD_SIDE).BackColor = PaleRed()
    End If
    
    If Damage.Elevator(PORT_SIDE) = True _
    Or Damage.ElevatorControls = True Then
        lblElevator(PORT_SIDE).BackColor = PaleRed()
    End If
    
    If Damage.Elevator(STBD_SIDE) = True _
    Or Damage.ElevatorControls = True Then
        lblElevator(STBD_SIDE).BackColor = PaleRed()
    End If
    
    If Damage.Rudder(PORT_SIDE) = 2 _
    Or Damage.RudderControls = True Then
        lblRudder(PORT_SIDE).BackColor = PaleRed()
    ElseIf Damage.Rudder(PORT_SIDE) = 1 Then
        lblRudder(PORT_SIDE).BackColor = PaleYellow()
    End If
    
    If Damage.Rudder(STBD_SIDE) = 2 _
    Or Damage.RudderControls = True Then
        lblRudder(STBD_SIDE).BackColor = PaleRed()
    ElseIf Damage.Rudder(STBD_SIDE) = 1 Then
        lblRudder(STBD_SIDE).BackColor = PaleYellow()
    End If
    
    If Damage.LandingGear = True Then
        lblLandingGear.BackColor = PaleRed()
    End If
    
    If Damage.NoseWheel = True _
    Or Damage.Tailwheel = True Then
        lblThirdWheel.BackColor = PaleRed()
    End If
    
    If Damage.Brake = True Then
        lblBrake.BackColor = PaleRed()
    End If
    
    Call ColorEngineGauge(1)
    Call ColorEngineGauge(2)
    Call ColorEngineGauge(3)
    Call ColorEngineGauge(4)

    proOilPressure(1).Value = (4 - Damage.OilTankLeak(1))
    proOilPressure(2).Value = (4 - Damage.OilTankLeak(2))
    proOilPressure(3).Value = (4 - Damage.OilTankLeak(3))
    proOilPressure(4).Value = (4 - Damage.OilTankLeak(4))
    
    Call ColorOilPressureGauge(1)
    Call ColorOilPressureGauge(2)
    Call ColorOilPressureGauge(3)
    Call ColorOilPressureGauge(4)
    
    If Bomber.Direction = OUTBOUND Then
        lblDirection.Caption = "Outbound"
    Else
        lblDirection.Caption = "Return Trip"
    End If
    
    lblAltitude.Caption = CStr(Bomber.Altitude)
    
    lblLocation.Caption = "Zone " & Bomber.CurrentZone & " / " & _
                          Mission.Zone(Bomber.CurrentZone).Terrain
    
    lblWeather.Caption = WeatherText(Mission.Zone(Bomber.CurrentZone).Weather)

    Call ColorEnvironmentGauge

    If Damage.BombSight = True Then
        lblBombSight.BackColor = PaleRed()
    End If
    
    If Damage.BombBayDoors = True Then
        lblBombBayDoors.BackColor = PaleRed()
    End If
    
    If Bomber.BombsOnBoard = False Then
        lblBombsOnBoard.Caption = "Bombs Away"
        lblBombsOnBoard.BackColor = vbButtonFace
    End If
    
    If Damage.NavigationEquipment = True Then
        lblNavigationEquipment.BackColor = PaleRed()
    End If
    
    If Damage.Radio = True Then
        lblRadio.BackColor = PaleRed()
    End If
    
    If Damage.IntercomSystem = True Then
        lblIntercom.BackColor = PaleRed()
    End If
    
    If Damage.RubberRafts = True Then
        lblRubberRafts.BackColor = PaleRed()
    End If
    
    proFuelGauge.Value = Bomber.FuelPoints
    Call ColorFuelGauge
    
    lblPeckhamPoints.Caption = Damage.PeckhamPoints
    Call ColorPeckhamPointsGauge

End Sub

'******************************************************************************
' ColorEnvironmentGauge
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Update display due to any damage that may have occured. This gauge
'        individually and collectively represents all oxygen and heating
'        systems.
'******************************************************************************
Private Sub ColorEnvironmentGauge()
    Dim intPos As Integer

    If Damage.OxygenSystem = True _
    Or Damage.HeatingSystem = True Then
        
        lblEnvironment.BackColor = PaleRed()
    
    Else

        For intPos = PILOT To AMMO_STOCKER
            
            If Damage.Oxygen(intPos) = 1 _
            Or Damage.Heater(intPos) = True Then
                
                lblEnvironment.BackColor = PaleYellow()
            
            ElseIf Damage.Oxygen(intPos) >= 2 Then
                
                lblEnvironment.BackColor = PaleRed()
            
            End If
            
        Next intPos
    
    End If
    
End Sub

'******************************************************************************
' ColorEngineGauge
'
' INPUT:  Key to a particular engine.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Update display due to any damage that may have occured.
'******************************************************************************
Private Sub ColorEngineGauge(ByVal intEngine As Integer)

    If Damage.EngineOut(intEngine) = True Then
        lblEngine(intEngine).BackColor = PaleRed()
    ElseIf Damage.Turbocharger(intEngine) = True Then
        lblEngine(intEngine).BackColor = PaleYellow()
    Else
        ' Engine miraculously restarted despite any previous damage.
        lblEngine(intEngine).BackColor = PaleGreen()
    End If
    
End Sub

'******************************************************************************
' ColorOilPressureGauge
'
' INPUT:  Key to a particular engine.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Update display due to any damage that may have occured.
'******************************************************************************
Private Sub ColorOilPressureGauge(ByVal intEngine As Integer)
    
    If proOilPressure(intEngine).Value = 4 Then
        Call SetBackColor(proOilPressure(intEngine).hwnd, PaleGreen())
    ElseIf proOilPressure(intEngine).Value = 3 _
    Or proOilPressure(intEngine).Value = 2 Then
        Call SetBackColor(proOilPressure(intEngine).hwnd, PaleYellow())
    ElseIf proOilPressure(intEngine).Value = 1 _
    Or proOilPressure(intEngine).Value = 0 Then
        Call SetBackColor(proOilPressure(intEngine).hwnd, PaleRed())
    End If

    Call SetBarColor(proOilPressure(intEngine).hwnd, vbBlack)

End Sub

'******************************************************************************
' ColorFuelGauge
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Update display due to any damage that may have occured, or fuel that
'        was normally consumed.
'******************************************************************************
Private Sub ColorFuelGauge()

    Dim intPct As Integer
    
    intPct = CInt((proFuelGauge.Value / proFuelGauge.Max) * 100)
    
    If intPct >= 51 Then
        Call SetBackColor(proFuelGauge.hwnd, PaleGreen())
    ElseIf intPct <= 50 _
    And intPct >= 26 Then
        Call SetBackColor(proFuelGauge.hwnd, PaleYellow())
    ElseIf intPct <= 25 Then
        Call SetBackColor(proFuelGauge.hwnd, PaleRed())
    End If

    Call SetBarColor(proFuelGauge.hwnd, vbBlack)

End Sub

'******************************************************************************
' ColorPeckhamPointsGauge
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Update display due to any damage that may have occured.
'******************************************************************************
Private Sub ColorPeckhamPointsGauge()

    If Damage.BurstInPlane = True Then
        
        lblPeckhamPoints.ForeColor = vbWhite
        lblPeckhamPoints.BackColor = vbBlack
    
    ElseIf lblPeckhamPoints.Caption >= 51 _
    And lblPeckhamPoints.Caption <= 200 Then
        
        lblPeckhamPoints.BackColor = PaleYellow()
    
    ElseIf lblPeckhamPoints.Caption >= 201 _
    And lblPeckhamPoints.Caption <= 400 Then
        
        lblPeckhamPoints.BackColor = PaleRed()
    
    ElseIf lblPeckhamPoints.Caption >= 401 Then
        
        lblPeckhamPoints.ForeColor = vbWhite
        lblPeckhamPoints.BackColor = vbBlack
    
    End If

End Sub

'******************************************************************************
' InitializeCrew
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Hide controls for non-existant positions, oterwise initialize them.
'******************************************************************************
Private Sub InitializeCrew()

    Dim intPos As Integer

    For intPos = PILOT To AMMO_STOCKER
        
        If PosExists(intPos) = False Then
            
            ' The position does not exist on this bomber. Hide the irrelevant
            ' controls.
            
            lblCrewPosition(intPos).Visible = False
            lblAirmanName(intPos).Visible = False
            lblAirmanStatus(intPos).Visible = False
            lblAirmanMission(intPos).Visible = False
            lblAirmanKills(intPos).Visible = False
            
        ElseIf PosOccupied(intPos) = True Then
            
            ' Unmanned positions will be visible, but grey.
            
            lblAirmanName(intPos).Caption = Bomber.Airman(intPos).Name
            lblAirmanStatus(intPos).Caption = "OK"
            lblAirmanStatus(intPos).BackColor = PaleGreen()
            lblAirmanMission(intPos).Caption = Bomber.Airman(intPos).Mission
            lblAirmanKills(intPos).Caption = Bomber.Airman(intPos).Kills
            
            ' The crew position captions are defaulted to their most common
            ' value, however the caption may vary depending on bomber model.
            
            Select Case Bomber.BomberModel
                
                Case B17_C:
                    
                    If intPos = BALL_GUNNER Then
                        lblCrewPosition(intPos).Caption = "Tunnel Gunner"
                    End If
                
                Case YB40:
                    
                    If intPos = MID_UPPER_GUNNER Then
                        lblCrewPosition(intPos).Caption = "Mid-Upper Gunner"
                    End If
                
                Case B24_D:
                    
                    If intPos = BALL_GUNNER Then
                        lblCrewPosition(intPos).Caption = "Tunnel Gunner"
'                    ElseIf intPos = STBD_WAIST_GUNNER Then
'                        lblCrewPosition(intPos).Caption = "Waist Gunner"
                    End If
                
                Case B24_E:
                    
                    If intPos = BALL_GUNNER Then
                        lblCrewPosition(intPos).Caption = "Tunnel Gunner"
'                    ElseIf intPos = STBD_WAIST_GUNNER Then
'                        lblCrewPosition(intPos).Caption = "Waist Gunner"
                    End If
                
                Case B24_GHJ:
                    
'                    If intPos = STBD_WAIST_GUNNER Then
'                        lblCrewPosition(intPos).Caption = "Waist Gunner"
'                    End If
                
                Case B24_LM:
                    
                    If intPos = BALL_GUNNER Then
                        lblCrewPosition(intPos).Caption = "Floor Ring Gunner"
'                    ElseIf intPos = STBD_WAIST_GUNNER Then
'                        lblCrewPosition(intPos).Caption = "Waist Gunner"
                    ElseIf intPos = TAIL_GUNNER Then
                        lblCrewPosition(intPos).Caption = "Stinger Gunner"
                    End If
                
                Case AVRO_LANCASTER:
                    
                    If intPos = BOMBARDIER Then
                        lblCrewPosition(intPos).Caption = "Bomb Aimer"
                    ElseIf intPos = ENGINEER Then
                        lblCrewPosition(intPos).Caption = "Flight Engineer"
                    ElseIf intPos = RADIO_OPERATOR Then
                        lblCrewPosition(intPos).Caption = "Wireless Operator"
                    ElseIf intPos = TAIL_GUNNER Then
                        lblCrewPosition(intPos).Caption = "Rear Gunner"
                    End If
                
            End Select
                    
        End If
        
    Next intPos
            
End Sub

'******************************************************************************
' RefreshCrew
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Update display due to any damage that may have occured, or kills they
'        may have scored.
'******************************************************************************
Public Sub RefreshCrew()

    Dim intPos As Integer
    Dim intIndex As Integer

    For intPos = PILOT To AMMO_STOCKER
        
        If PosOccupied(intPos) = False Then
            
            lblAirmanName(intPos).Caption = ""
            lblAirmanStatus(intPos).Caption = ""
            lblAirmanMission(intPos).Caption = ""
            lblAirmanKills(intPos).Caption = ""
            lblAirmanStatus(intPos).BackColor = vbButtonFace
        
        ElseIf PosExists(intPos) = True Then
            
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
            
            lblAirmanName(intPos).Caption = Bomber.Airman(intIndex).Name
            
            Select Case Bomber.Airman(intIndex).Status
            
                Case DUTY_STATUS:
            
                    If Bomber.Airman(intIndex).Frostbite = False Then
                    
                        lblAirmanStatus(intPos).Caption = "OK"
                        lblAirmanStatus(intPos).ForeColor = vbBlack
                        lblAirmanStatus(intPos).BackColor = PaleGreen()
                    
                    Else
                    
                        ' If the airman is wounded, displaying that takes
                        ' precedence, otherwise display frostbite.
                    
                        lblAirmanStatus(intPos).Caption = "Frost"
                        lblAirmanStatus(intPos).ForeColor = vbBlack
                        lblAirmanStatus(intPos).BackColor = PaleYellow()
                    
                    End If
                
                Case LW1_STATUS:
                    
                    lblAirmanStatus(intPos).Caption = "LW1"
                    lblAirmanStatus(intPos).ForeColor = vbBlack
                    lblAirmanStatus(intPos).BackColor = PaleYellow()
                
                Case LW2_STATUS:
                    
                    lblAirmanStatus(intPos).Caption = "LW2"
                    lblAirmanStatus(intPos).ForeColor = vbBlack
                    lblAirmanStatus(intPos).BackColor = PaleYellow()
                
                Case SW_STATUS:
                    
                    lblAirmanStatus(intPos).Caption = "SW"
                    lblAirmanStatus(intPos).ForeColor = vbBlack
                    lblAirmanStatus(intPos).BackColor = PaleRed()
                
                Case KIA_STATUS:
                
                    lblAirmanStatus(intPos).Caption = "KIA"
                    lblAirmanStatus(intPos).ForeColor = vbWhite
                    lblAirmanStatus(intPos).BackColor = vbBlack
                
            End Select
            
            lblAirmanMission(intPos).Caption = Bomber.Airman(intIndex).Mission
            lblAirmanKills(intPos).Caption = Bomber.Airman(intIndex).Kills
            
        End If
        
    Next intPos
            
End Sub

'******************************************************************************
' InitializeGuns
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Hide controls for non-existant positions, oterwise initialize them.
'******************************************************************************
Private Sub InitializeGuns()

    Dim intGun As Integer

    For intGun = MID_UPPER_MG To TAIL_MG
        
        If GunExists(intGun) = False Then
            
            lblGunAmmo(intGun).Visible = False
            chkSpray(intGun).Visible = False
            lblGunName(intGun).Visible = False
        
        Else
            
            lblGunAmmo(intGun).Caption = Bomber.Gun(intGun).Ammo
            lblGunName(intGun).BackColor = PaleGreen()
        
            ' The gun captions are defaulted to their most common
            ' value, however the caption may vary depending on bomber model.
            
            Select Case Bomber.BomberModel
                
                Case B17_C:
                    
                    If intGun = BALL_TURRET_MG Then
                        lblGunName(intGun).Caption = "Tunnel Gun"
                    End If
                
                Case YB40:
                    
                    If intGun = MID_UPPER_MG Then
                        lblGunName(intGun).Caption = "Mid-Upper Turret"
                    End If
                
                Case B24_D:
                    
                    If intGun = BALL_TURRET_MG Then
                        lblGunName(intGun).Caption = "Tunnel Gun"
                    End If
                
                Case B24_E:
                    
                    If intGun = BALL_TURRET_MG Then
                        lblGunName(intGun).Caption = "Tunnel Gun"
                    End If
                
                Case B24_LM:
                    
                    If intGun = BALL_TURRET_MG Then
                        lblGunName(intGun).Caption = "Floor Ring"
                    ElseIf intGun = TAIL_MG Then
                        lblGunName(intGun).Caption = "Stinger"
                    End If
                
                Case AVRO_LANCASTER:
                    
                    If intGun = TAIL_MG Then
                        lblGunName(intGun).Caption = "Rear Turret"
                    End If
                
            End Select
                    
        End If
    
    Next intGun

End Sub

'******************************************************************************
' UnjamGuns
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Assume that combat ready gunners are trying to unjam their guns,
'        rather than twiddling their fingers for 10 zones.
'******************************************************************************
Private Sub UnjamGuns()
    Dim intGun As Integer
    Dim intIndex As Integer
    Dim intRoll As Integer
    
    For intGun = MID_UPPER_MG To TAIL_MG

        If Bomber.Gun(intGun).Status = MG_JAMMED _
        And GunOccupied(intGun) = True Then
'        And Bomber.Gun(intGun).MannedBy <> HIDDEN_MG _
'        And Bomber.Gun(intGun).MannedBy <> UNMANNED_MG Then
        
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Gun(intGun).MannedBy)
            
            If Bomber.Airman(intIndex).Status <= LW2_STATUS Then
                
                intRoll = Random1D6()
                
                If Bomber.Airman(intIndex).Kills >= 5 Then
                    intRoll = intRoll + 1
                End If
        
                If intRoll >= 6 Then
                
                    Bomber.Gun(intGun).Status = MG_OKAY
'                    lblGunName(intGun).BackColor = PaleGreen()
                    UpdateMessage lblGunName(intGun).Caption & " unjammed"
                
                End If
            
            End If
        
        End If
        
    Next intGun
    
End Sub

'******************************************************************************
' RefreshGuns
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: Update display due to any damage that may have occured.
'******************************************************************************
Private Sub RefreshGuns()

    Dim intGun As Integer

    For intGun = MID_UPPER_MG To TAIL_MG
        
        If GunExists(intGun) = True Then
            
            lblGunAmmo(intGun).BackColor = vbButtonFace
            
            lblGunAmmo(intGun).Caption = Bomber.Gun(intGun).Ammo
            
            chkSpray(intGun).Value = vbUnchecked
            
            Select Case Bomber.Gun(intGun).Status
                
                Case MG_OKAY:
                
                    lblGunName(intGun).BackColor = PaleGreen()
                
                Case MG_JAMMED:
                
                    lblGunName(intGun).BackColor = PaleYellow()
                
                Case MG_INOPERABLE:
            
                    lblGunName(intGun).BackColor = PaleRed()
            
            End Select
            
        End If
    
    Next intGun
            
End Sub

'******************************************************************************
' FlyMission
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is the "mission engine". First, create a mission on the main
'         menu's Generate Mission tab, then click the Fly Mission button,
'         launching the Mission form. When the user clicks the Take Off button,
'         the mission engine will be invoked. The engine runs until the bomber
'         is shot down or returns to base. After the mission is complete, the
'         Mission form is closed.
'
' This is the general format of how a mission is conducted:
'
'         Take Off
'            If B-24
'                Roll for crash
'
'         Enter Next/Prev Zone
'            increment/decrement zone
'            If zone is target zone
'                roll for weather in zone ... O1Weather()
'            If zone is base zone
'                roll for weather in zone ... O1Weather()
'                land ... G9LandingOnLand()
'                If bomber had a BIP, it is unfit for further missions (scrapped)
'            If bomber is at low altitude (10,000 feet)
'            & over enemy territory
'                light flak
'                Repeat flak procedure from "roll for flak hits three times" point forward
'            Determine quality of fighter cover (Lanc = none) ... G5FighterCover()
'            Determine number of German waves ... B1NumberOfGermanFighterWaves()
'                If Lancaster
'                    0 or 1 waves
'            For each wave
'                Determine type of wave ... B3AttackingFighterWave()
'                    If Lancaster
'                        Tame Boar / Wild Boar
'                        If Tame Boar
'                            Spotting phase
'                            If Tame Boar not spotted
'                                Tame Boar gets bonus surprise attack before resolving normal combat
'                If out of formation
'                    Add Me-109 at 12 level
'                else if lead bomber
'                    Add Me-109 at 12 high
'                else if tail bomber
'                    Add Me-109 at 6 high
'                For each German fighter
'                    Determine pilot quality
'                loop
'                while German fighters remain
'                    Place German planes at clock positions
'                    Determine quantity of fighter cover (Lanc = none) ... M4FighterCoverDefense()
'                    Select German fighters to be removed
'                    Remove German fighters
'                    If out of formation
'                        If operating engines >= 3
'                        & control cables operative
'                        & negative landing modifiers <= 2
'                        & pilot and copilot in position
'                        & other damage does not prevent evasion
'                            Evasive action?
'                    Determine which MGs may fire at which German fighters
'                        If MG unmanned
'                        Or MG is out of ammo
'                        Or MG is inoperative/jammed
'                            May not fire at any target
'                    Designate targets for MGs
'                    For each MG with a designated target
'                        If remaining ammo >= 3
'                        & attack from rear, sides or underneath
'                            Spray fire?
'                        If spray fire
'                            Mark off three ammo boxes
'                            Determine if MG hit German fighter ... M5SprayFire()
'                        Else
'                            Mark off one ammo box
'                            Determine if MG hit German fighter ... M1DefensiveFire()
'                        If German fighter was hit
'                            roll for damage ... M2HitDamageAgainstGermanFighter()
'                            place damage marker
'                            If German fighter destroyed
'                                Give credit to gunner
'                                Remove it from map
'                    loop
'                    For each remaining German fighter
'                        Determine if it hits bomber ... M3GermanOffensiveFire()
'                        If bomber was hit
'                            determine number of hits ... B4ShellHitsByArea()
'                            For each hit
'                                determine location of hit ... B5AreaDamage()
'                                if type of hit is "walking"
'                                    delete all hits
'                                    set hits = quantity for the type of walk
'                                    set type of each hit to type of hit(s) associated with the type of walk
'                                    exit loop
'                            loop
'                            For each hit
'                                determine damage (appropriate P-chart or BL-chart)
'                            loop
'                        else if bomber is in formation
'                            Remove German fighter from map
'                        if German fighter has FBOA damage
'                            Remove German fighter from map
'                        if German fighter attacked from 10:30, 12 or 1
'                        & tail MG is manned
'                        & tail MG has ammo
'                        & tail MG is operative
'                            If tail gunner wishes to make a passing shot
'                                Perform defensive fire routine from "Designate targets for MGs" point forward
'                        increment number attacks German fighter has made
'                        if German fighter has FBOA damage, or has made three attacks
'                            Remove German fighter from map
'                        else
'                            determine successive attacking position ... B6SuccessiveAttacks()
'                    loop
'                    crash landing? (water or enemy territory)
'                    Swap crew positions? (within same compartment only)
'                loop
'                Swap crew positions? (any)
'                Swap ammo between positions?
'                If out of formation
'                    Repeat wave process from "Determine number of German waves" point forward
'            loop
'            Abort mission?
'            If zone is target zone
'                If Lancaster
'                & inbound
'                    Searchlight phase
'                Resolve flak
'                    determine flak over target ... O2FlakOverTarget()
'                    roll for flak hits three times ... O3FlakToHitBomber()
'                    For each flak hit
'                        determine number of shell hits ... O4EffectOfFlakHits()
'                        if number of shell hits = "burst in plane"
'                            determine location of BIP ... O5AreaAffectedByFlakHit()
'                            if area = bomb bay
'                                all crew KIA
'                                bomber destroyed
'                                end mission
'                            all crew in area KIA
'                            inflict every possible damage result for that area (appropriate P-chart or BL-chart)
'                            if area = wing, tail or flight deck
'                                emergency bail out
'                            else
'                                out of formation
'                                -4 to landing roll
'                                no evasive action allowed
'                        else
'                            For each shell hit
'                                determine location of hit ... O5AreaAffectedByFlakHit()
'                                determine damage (appropriate P-chart or BL-chart)
'                            loop
'                        loop
'                    loop
'               If Lancaster
'               & inbound
'                   Wild Boar ... Repeat wave process from "Determine number of German waves" point forward
'               Resolve Bomb run
'                   determine if bombs were on target ... O6BombRun()
'                   determine bombing accuracy ... O7BombingAccuracy()
'               Turn the plane around
'            Repeat wave procedure in target zone
'******************************************************************************
Private Function FlyMission()
    Dim blnContinueMission As Boolean
    
    blnContinueMission = True

    If TakeOff() = False Then
        Exit Function
    End If
    
    ' Endless loop that exits only when the bomber returns to base or is shot
    ' down.
    
    Do
       
        If EnterZone() = END_MISSION Then
            Exit Function
        End If
        
        Call RefreshMissionForm

        If Bomber.CurrentZone = BASE_ZONE Then
            If BailOverBase() = False Then
                Call G9LandingOnLand
            End If
            Exit Function
        End If

        If Bomber.Altitude = LOW_ALTITUDE _
        And EnemyTerritory() = True Then
            ' Bomber takes light flak when as it flies over a zone at low
            ' altitude. (If the bomber spendS two turns in the zone, it checks
            ' twice.) Separate from flak surrounding the target, this represents
            ' ad hoc local resistance.
            
            If FlakCombat(LIGHT_FLAK, False) = END_MISSION Then
                Exit Function
            End If
            
            Call RefreshMissionForm
        End If

        If NormalZoneCombat() = END_MISSION Then
            Exit Function
        End If
        
        Call AbortMission
        
        If Bomber.CurrentZone = Mission.TargetZone _
        And Bomber.Direction = OUTBOUND Then
        
            UpdateMessage vbCrLf & "Target Zone"
    
            If Mission.Zone(Bomber.CurrentZone).Weather = STORM_WEATHER Then
                ' If there is a storm in the target zone, there is no flak and
                ' the bomber may not bomb the target. The bomber must abort or
                ' attack an alternate target.
                UpdateMessage "Target obscured by stormy weather"
' TODO: alternate target
                Call JettisonPayload(True, False)
            ElseIf TargetZoneCombat() = END_MISSION Then
                Exit Function
            End If
        
            Call RefreshMissionForm
        
        End If
        
    Loop While blnContinueMission

End Function

'******************************************************************************
' BailOverBase
'
' INPUT:
'
' OUTPUT: n/a
'
' RETURN: True if the crew bailed out, otherwise false.
'
' NOTES:  If there is the possibility that the bomber will crash while attempting
'         to land at its base, the user will have the option of instead bailing
'         out over the base. As crash landing is *always* possible, even when
'         the plane is whole, the possibility must be very high for the user to
'         be presented with the option.
'
' the lowest possible roll is 2. On 2d6 that is a 3% chance. Roll -3 is necessary
' to automatically kill all aboard in a crash landing. If the only DM is -11 for
' pilot incapacity, there is a 72% of an all kia crash. If DM -6, there is a
' 17% chance of an all kia crash, 28% chance of a serious wound crash, and 42%
' of a wounding crash.
'

'******************************************************************************
Private Function BailOverBase() As Integer

    Dim intDM As Integer
    Dim intEnginesOut As Integer
    Dim blnBailOut
    
    BailOverBase = False
    
    ' O-1 Weather: Note a and Note b, plus "The General" (Volume 24, #6)
    ' variant.
    
    If Mission.Zone(BASE_ZONE).Weather = POOR_WEATHER Then
        intDM = intDM - 1
    ElseIf Mission.Zone(BASE_ZONE).Weather = BAD_WEATHER Then
        intDM = intDM - 2
    ElseIf Mission.Zone(BASE_ZONE).Weather = STORM_WEATHER Then
        intDM = intDM - 3
    End If
    
    If Bomber.Airman(PILOT).Status >= SW_STATUS _
    And Bomber.Airman(COPILOT).Status >= SW_STATUS Then
        ' Note f: Pilot and copilot both unable to operate the controls.
        ' Another airman must land the plane.
        intDM = intDM - 11
    ElseIf (Bomber.Airman(PILOT).Status <= LW1_STATUS And Bomber.Airman(PILOT).Mission >= 11) _
    Or (Bomber.Airman(COPILOT).Status <= LW1_STATUS And Bomber.Airman(COPILOT).Mission >= 11) Then
        ' Note b: The controls are manned by the pilot and/or copilot, at
        ' least one whom is a veteran who has not been lightly wounded
        ' more than once.
        intDM = intDM + 1
    End If
    
    intEnginesOut = CountEnginesOut()
    
    If intEnginesOut = 3 Then
        ' Note h.
        intDM = intDM - 3
    ElseIf intEnginesOut = 4 Then
        ' Note i.
        intDM = intDM - 7
    End If
    
    If Damage.Window >= 2 Then
        ' P-2 Pilot Compartment.
        intDM = intDM - 1
    End If
    
    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
        
        If Damage.Autopilot = True Then
            ' B24s were more difficult and exhausting to fly.
            intDM = intDM - 2
        End If

    End If
    
    If Damage.ControlCables >= 2 Then
        ' P-2 Pilot Compartment.
        intDM = intDM - 1
    End If
    
    If Damage.ElevatorControls = True _
    Or (Damage.Elevator(PORT_SIDE) = True _
    And Damage.Elevator(STBD_SIDE) = True) Then
        ' P-6 Tail Section: Note b.
        intDM = intDM - 1
    End If
    
    ' P-6 Tail Section.
    
    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.BomberModel = B17_F _
    Or Bomber.BomberModel = B17_G _
    Or Bomber.BomberModel = YB40 Then
            
        ' A B-17 only has one rudder, so by default it is the 'port side'.
        
        If Damage.RudderControls = True _
        Or Damage.Rudder(PORT_SIDE) >= 3 Then
            intDM = intDM - 1
        End If
            
    ElseIf Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
            
        If Damage.RudderControls = True _
        Or (Damage.Rudder(PORT_SIDE) >= 2 _
        And Damage.Rudder(STBD_SIDE) >= 2) Then
            intDM = intDM - 2
        End If

    Else ' Lancaster
    
        If Damage.RudderControls = True _
        Or (Damage.Rudder(PORT_SIDE) >= 2 _
        And Damage.Rudder(STBD_SIDE) >= 2) Then
            intDM = intDM - 1
        End If

    End If
    
    If Damage.WingFlapControls = True _
    Or (Damage.WingFlap(PORT_SIDE) = True _
    And Damage.WingFlap(STBD_SIDE) = True) Then
        ' BL-1 Wings: Note b.
        intDM = intDM - 1
    End If
    
    If Damage.AileronControls = True _
    Or (Damage.Aileron(PORT_SIDE) = True _
    And Damage.Aileron(STBD_SIDE) = True) Then
        ' BL-1 Wings: Note b.
        intDM = intDM - 1
    End If
    
    If Damage.Brake = True Then
        ' BL-1 Wings: Note h.
        intDM = intDM - 1
    End If

    If Damage.LandingGear = True Then
        
        If Bomber.BomberModel = B24_D _
        Or Bomber.BomberModel = B24_E _
        Or Bomber.BomberModel = B24_GHJ _
        Or Bomber.BomberModel = B24_LM Then
            ' Flimsy roll up doors tended to collapse when doing belly
            ' landings.
            intDM = intDM - 4
        Else
            ' BL-1 Wings: Note i.
            intDM = intDM - 3
        End If
    
    End If
    
    If Damage.Tailwheel = True Then
        ' P-6 Tail Section. B-17s and Lancasters only.
        intDM = intDM - 1
    ElseIf Damage.NoseWheel = True Then
        ' B-24s are only models which can have nosewheel damage, which makes
        ' it more likely the bomber will flip over.
        intDM = intDM - 2
    End If
    
    If Damage.BurstInPlane = True Then
        ' Rule 19.2.d.
        intDM = intDM - 4
    End If
    
    If intDM <= -6 _
    Or Bomber.BombsOnBoard = True Then
        ' High probability of a serious crash landing. Ask if the user
        ' would prefer to bailout over base instead. Pause the program.
    
        Call DisplayTempOptions(BAIL_OR_CRASH, BAILOUT_BASE, ATTEMPT_LANDING)
        
        Call Interrupt(cmdInterrupt, BAIL_OR_CRASH)
    
        ' User clicked the next action button.
        
        If optOneTemp.Value = True Then
            blnBailOut = True
        Else ' optTwoTemp was clicked
            blnBailOut = False
        End If
        
        Call RemoveTempOptions
    
        If blnBailOut = True Then
            
            Call G6ControlledBailout(False)
            
            ' No need to set function = END_MISSION, as we would not be
            ' here if the mission was not already over. All we determined
            ' above was how the mission should end.
            
            BailOverBase = True
        
        End If
            
    End If
    
End Function

'******************************************************************************
' EnterZone
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, or default value.
'
' NOTES:  Controls the number of turns a bomber spends in a zone and the change
'         of direction after the target is bombed.
'******************************************************************************
Private Function EnterZone() As Integer
    Dim intEng As Integer
    Dim intEnginesOut As Integer
    Dim blnSpendAnotherTurnInZone As Boolean

    blnSpendAnotherTurnInZone = False
    
    For intEng = 1 To 4
        
        If Damage.EngineOut(intEng) = True Then
            intEnginesOut = intEnginesOut + 1
        End If
            
    Next intEng
    
    If Bomber.TurnsInZone = 1 _
    And Bomber.CurrentZone < Mission.TargetZone Then
        
        If intEnginesOut = 1 _
        And Bomber.BombsOnBoard = True Then
            
            ' Two turns in zone
            blnSpendAnotherTurnInZone = True
        
        ElseIf intEnginesOut >= 2 Then
            
            ' Two turns in zone
            blnSpendAnotherTurnInZone = True
        
        ElseIf Bomber.InFormation = False _
        And Damage.NavigationEquipment = True _
        And IsOdd(Bomber.CurrentZone) = True Then
            
            ' Two turns in zone
            blnSpendAnotherTurnInZone = True
        
        End If
    
    End If
        
    If blnSpendAnotherTurnInZone = True Then
'    If Bomber.InFormation = False _
'    And Bomber.TurnsInZone = 1 _
'    And Bomber.CurrentZone <> Mission.TargetZone Then
        ' If the bomber is out of formation, it must spend a second turn in
        ' the zone. If the current zone is the target zone, then the bomber
        ' would spend one turn inbound and one turn on the return trip.
        Bomber.TurnsInZone = 2
        ' Don't burn fuel the second turn in the normal zone.
        
        If Bomber.Direction = OUTBOUND Then
            UpdateMessage vbCrLf & "Zone " & Bomber.CurrentZone & ", Turn 2 (Out)"
        Else
            UpdateMessage vbCrLf & "Zone " & Bomber.CurrentZone & ", Turn 2 (Ret)"
        End If
        
        UpdateMessage "Terrain: " & Mission.Zone(Bomber.CurrentZone).Terrain
        
        If Mission.Zone(Bomber.CurrentZone).Contrail = True Then
            UpdateMessage "Weather: " & WeatherText(Mission.Zone(Bomber.CurrentZone).Weather) & " (contrails form)"
        Else
            UpdateMessage "Weather: " & WeatherText(Mission.Zone(Bomber.CurrentZone).Weather)
        End If

        Exit Function
    End If
    
    ' Exit the previous/current zone before entering the next zone.
' erroneous comment: actually all ExitPrevZone() does is some end of zone cleanup.
' the actual exitting of the zone occurs in this function, below.
    If ExitPrevZone() = END_MISSION Then
        EnterZone = END_MISSION
        Exit Function
    End If
 
    '---------------------------------------------------------
    ' Exitting the previous zone may have knocked out more engines, so we need
    ' to check again.
    
    intEnginesOut = CountEnginesOut()

    If QuitBeforeWater(intEnginesOut) = END_MISSION Then
        EnterZone = END_MISSION
        Exit Function
    End If

    If intEnginesOut = 3 Then
        ' One engine operating: May fly one additional zone, then either ditch,
        ' crash or bail out. May fly one zone further than that by throwing all
        ' ammo and handheld extinguishers overboard, plus jettisoning bombs.
        ' (If bomb bay doors are jammed, this can't be done.) No engines
        ' operating is already properly handled.
        Call JettisonPayload(True, True)
    End If
    '---------------------------------------------------------

    ' Enter the next zone.
    
    If Bomber.CurrentZone = Mission.TargetZone _
    And Bomber.TurnsInZone = 1 Then
        ' Bomber must spend two turns in the target zone. The first turn
        ' inbound, the second turn on the return trip. Change direction,
        ' but not zone.
        Bomber.Direction = RETURN_TRIP
        Bomber.TurnsInZone = 2
        ' Don't burn fuel the second turn in the target zone.
        UpdateMessage vbCrLf & "Zone " & Bomber.CurrentZone & ", Turn 2 (Ret)"
        UpdateMessage "Terrain: " & Mission.Zone(Bomber.CurrentZone).Terrain
        
        If Mission.Zone(Bomber.CurrentZone).Contrail = True Then
            UpdateMessage "Weather: " & WeatherText(Mission.Zone(Bomber.CurrentZone).Weather) & " (contrails form)"
        Else
            UpdateMessage "Weather: " & WeatherText(Mission.Zone(Bomber.CurrentZone).Weather)
        End If

    ElseIf Bomber.Direction = OUTBOUND Then
        Bomber.CurrentZone = Bomber.CurrentZone + 1
        Bomber.TurnsInZone = 1
        Bomber.FuelPoints = Bomber.FuelPoints - 1
        UpdateMessage vbCrLf & "Zone " & Bomber.CurrentZone & ", Turn 1 (Out)"
'        UpdateMessage "Terrain: " & Mission.Zone(Bomber.CurrentZone).Terrain
    
        UpdateMessage "Terrain: " & Mission.Zone(Bomber.CurrentZone).Terrain

        If Mission.Zone(Bomber.CurrentZone).Contrail = True Then
            UpdateMessage "Weather: " & WeatherText(Mission.Zone(Bomber.CurrentZone).Weather) & " (contrails form)"
        Else
            UpdateMessage "Weather: " & WeatherText(Mission.Zone(Bomber.CurrentZone).Weather)
        End If

        Call ConsumeOil
        
        If Mission.Options.MechanicalFailures = True Then
            If MechanicalFailure() = END_MISSION Then
                EnterZone = END_MISSION
                Exit Function
            End If
        End If
    
        If AlpsDirection() = ALPS_BELOW _
        And Mission.Zone(Bomber.CurrentZone).Weather = POOR_WEATHER _
        And Random1D6() = 6 Then
        
            ' Player chose not to abort despite fog over Alps. Bomber entered
            ' the Alps, could not see in the fog, then flew into a mountain.

            UpdateMessage "Dense fog over Alps. Mountain ahead! Too late ..."
            Bomber.Status = CRASHED_STATUS
            Call CrewFinish(KIA_STATUS)
            EnterZone = END_MISSION
            Exit Function

        End If
        
    Else ' RETURN_TRIP
        Bomber.CurrentZone = Bomber.CurrentZone - 1
        Bomber.TurnsInZone = 1
        Bomber.FuelPoints = Bomber.FuelPoints - 1
        UpdateMessage vbCrLf & "Zone " & Bomber.CurrentZone & ", Turn 1 (Ret)"
'        UpdateMessage "Terrain: " & Mission.Zone(Bomber.CurrentZone).Terrain
    
        UpdateMessage "Terrain: " & Mission.Zone(Bomber.CurrentZone).Terrain
        
        If Mission.Zone(Bomber.CurrentZone).Contrail = True Then
            UpdateMessage "Weather: " & WeatherText(Mission.Zone(Bomber.CurrentZone).Weather) & " (contrails form)"
        Else
            UpdateMessage "Weather: " & WeatherText(Mission.Zone(Bomber.CurrentZone).Weather)
        End If

        Call ConsumeOil
        
        If Mission.Options.MechanicalFailures = True Then
            If MechanicalFailure() = END_MISSION Then
                EnterZone = END_MISSION
                Exit Function
            End If
        End If
    
    End If
Dim a
a = 1

'    UpdateMessage "Terrain: " & Mission.Zone(Bomber.CurrentZone).Terrain
'    UpdateMessage "Weather: " & WeatherText(Mission.Zone(Bomber.CurrentZone).Weather)

End Function

'******************************************************************************
' QuitBeforeWater
'
' INPUT:  The number of engines that are out.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  If there is the possibility that the bomber will run out of power or
'         fuel in a water zone, then the user will have the option of abandoning
'         ship while over land, or continuing (i.e., spend the rest of the war
'         as a POW, or risk death in the sea?)
'******************************************************************************
Private Function QuitBeforeWater(ByVal intEnginesOut As Integer) As Integer

    Dim intLastZone As Integer
    Dim blnWaterHazard As Boolean
    Dim intIndex As Integer
    Dim blnQuitNow As Boolean
    
    intLastZone = Bomber.CurrentZone - Bomber.FuelPoints
    
    If intLastZone >= 2 Then
        
        If OverWater2(intLastZone) = True Then
            
            ' If the last drop of fuel is used in a water zone, then there is a
            ' water hazard.
            
            blnWaterHazard = True
        
        End If
    
    End If
    
    If Bomber.CurrentZone >= 3 _
    And intEnginesOut >= 2 Then
        
        ' Every mission crosses water in zone 2, except for where a few targets
        ' are actually in zone 2. If two engines are out, the bomber may lose
        ' power over a water zone.
        
        blnWaterHazard = True
    
    End If
 
    If intLastZone <= 0 Then
        ' If last zone is 0, the bomber has just enough fuel to get back to
        ' base. If the value is negative, then the bomber has more than enough
        ' fuel to get back to base.
        intLastZone = 1
    End If
 
    If blnWaterHazard = True Then
    
        ' Bomber is in danger of having having to bailout over water or ditch
        ' later in the mission.
        
        If (Bomber.Position(NAVIGATOR).CurrentSerialNum = Bomber.Position(NAVIGATOR).AssignedSerialNum _
        And Bomber.Airman(NAVIGATOR).Status <= LW2_STATUS) Then
    
            ' The navigator is qualified and capable of making the distance
            ' calculation.
    
            If OverWater2(intLastZone) = True _
            And OverWater2(Bomber.CurrentZone) = False Then
            
                ' The last zone is over water
        
                Call DisplayTempOptions(QUIT_BEFORE_WATER, "Yes", "No")
                
                Call Interrupt(cmdInterrupt, QUIT_BEFORE_WATER)
        
                ' User clicked the next action button.
                If optOneTemp.Value = True Then
                    blnQuitNow = True
                Else ' optTwoTemp was clicked
                    blnQuitNow = False
                End If
                
                Call RemoveTempOptions
            
                If blnQuitNow = True Then
                
                    Call BailOrCrash
                    QuitBeforeWater = END_MISSION
                    Exit Function
                
                End If
            
            End If
            
        End If
    
    End If

End Function

'******************************************************************************
' JettisonPayload
'
' INPUT:  Indcate whether bombs and/or excess weight should be jettisoned.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Bombs need to be jettisoned to lower odds of explosion due to combat
'         or rough landing. Excess weight ought to be jettisoned when the
'         bomber is low on power.
'******************************************************************************
Private Sub JettisonPayload(ByVal blnBombs As Boolean, ByVal blnExcess As Boolean)
    
    Dim blnBombsJettisoned As Boolean
    Dim blnExcessJettisoned As Boolean
    Dim intAbleBodied As Integer
    Dim intPos As Integer
    Dim intIndex As Integer
    Dim blnJettisonExcess As Boolean
    Dim intGun As Integer
    Dim blnManual As Boolean
    
    blnBombsJettisoned = True
    blnExcessJettisoned = True
    intAbleBodied = 0
    blnManual = False
    
    ' Cycle through the crew's originally assigned positions.
    For intPos = PILOT To AMMO_STOCKER
        
        If PosOccupied(intPos) = True Then
        
            ' Airman currently in position
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
            
            If Bomber.Airman(intIndex).Status <= LW2_STATUS Then
                
                ' There must be at least three able-bodied airmen to jettison
                ' excess or to manually jettison bombs. (Two to fly the bomber;
                ' one to kick the payload out the door.)
    
                intAbleBodied = intAbleBodied + 1
            
            End If
        
        End If
    
    Next intPos
    
    If blnBombs = True _
    And Bomber.BombsOnBoard = True Then
        ' Assume this is done automatically, if the plane is mechanically
        ' capable. (The pilot would have to be pretty stupid to forget.)

        If Damage.BombBayDoors = False Then
            
            ' The bomb bay doors are functioning, so the payload may be
            ' jettisoned.
            
            If Damage.BombControls = True _
            Or Damage.BombRelease = True Then
            
                ' The payload must be manually jettisoned.
                
                If intAbleBodied <= 2 Then
                    blnBombsJettisoned = False
                Else
                    blnManual = True
                End If

            End If
        
        Else ' Can't open bomb bay doors
            
            blnBombsJettisoned = False
        
        End If
    
        If blnBombsJettisoned = True Then
        
            If blnManual = True Then
                UpdateMessage "Bombs manually jettisoned."
            Else
                UpdateMessage "Bombs jettisoned."
            End If
            
            Bomber.BombsOnBoard = False
        
        Else
            
            UpdateMessage "Unable to jettison bombs."
        
        End If
    
    End If
    
    If blnExcess = True _
    And Damage.LastZone = 0 Then
    
        ' This is not automatic. Ask the user.
    
        Call DisplayTempOptions(JETTISON_EXCESS, "Yes", "No")
        
        ' Pause the program while the user decides whether or not to fly one zone
        ' further by throwing all excess weight overboard.
    
        Call Interrupt(cmdInterrupt, JETTISON_EXCESS)
    
        ' User clicked the next action button.
        
        If optOneTemp.Value = True Then
            blnJettisonExcess = True
        Else ' optTwoTemp was clicked
            blnJettisonExcess = False
        End If
        
        Call RemoveTempOptions
    
        If blnJettisonExcess = False _
        Or intAbleBodied <= 2 Then
            blnExcessJettisoned = False
        End If

        If blnExcessJettisoned = True Then
        
            ' Jettison hand held extinguishers, all ammo, and extra fuel.
            
            Bomber.HandHeldExtinguishers = 0
            
            For intGun = MID_UPPER_MG To TAIL_MG
                Bomber.Gun(intGun).Ammo = 0
            Next intGun
        
            Bomber.ExtraAmmo = 0
            
            UpdateMessage "Excess weight thrown overboard."
        
            If Bomber.ExtraFuelInBombBay = True Then
                Bomber.ExtraFuelInBombBay = False
                UpdateMessage "Excess fuel dumped."
            End If
        
            ' Bomber should have already aborted, being on its return trip.
            ' So, we assume the zones are counting down.
        
            Damage.LastZone = Bomber.CurrentZone - 2
        
            UpdateMessage "Bomber may fly two more zones."
                
        Else
        
            If intAbleBodied <= 2 Then
            
                UpdateMessage "Unable to throw excess weight overboard."
                
            Else
            
                UpdateMessage "Pilot elected to maintain vital supplies."
                
            End If
                
            Damage.LastZone = Bomber.CurrentZone - 1
        
            UpdateMessage "Bomber may fly one more zone."
                
        End If
    
    End If
    
End Sub

'******************************************************************************
' ConsumeOil
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  If an engine is leaking oil, and some oil remains, drain one point.
'******************************************************************************
Private Sub ConsumeOil()
    Dim intEngine As Integer
    
    For intEngine = 1 To 4

        If Damage.OilTankLeak(intEngine) >= LT_LEAK _
        And Damage.OilTankLeak(intEngine) <= HVY_LEAK Then
            ' There is a leak, but some oil remains. Consume one zone's worth.
            Damage.OilTankLeak(intEngine) = Damage.OilTankLeak(intEngine) + 1
        End If
    
    Next intEngine

End Sub

'******************************************************************************
' MechanicalFailure
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, or default value.
'
' NOTES:  From the "Theater Modifications" article in "The General" (Volume 24,
'         #6). If a failure is rolled for, but does not occur -- e.g., due to
'         previous damage to the system -- then the failure is ignored. The
'         game is hard enough as it is, and the mechanical failure option makes
'         it more difficult, so cut the player that break.
'******************************************************************************
Private Function MechanicalFailure() As Integer
    Dim intRoll As Integer
    Dim intEngine As Integer
    Dim intGun As Integer
    Dim intPos As Integer
    Dim strEffect As String

    If Random2D6() <> 12 Then
        Exit Function
    End If

    intRoll = Random2D6()

    Select Case intRoll
    
        Case 2: ' Engine

            intEngine = RandomDX(4)
            
            If Damage.EngineOut(intEngine) = False Then
                Damage.EngineOut(intEngine) = True
                UpdateMessage "#" & intEngine & " engine shut down due to " & _
                              "malfunction."
                              
                If Damage.FeatheringCtrl = True Then
                    UpdateMessage "#" & intEngine & " engine not feathered."
                    Damage.EngineDrag(intEngine) = True
                    Call DropOutOfFormation
                End If
            
            End If
        
        Case 3: ' Turbocharger
    
            intEngine = RandomDX(4)
            
            If Damage.EngineOut(intEngine) = False _
            And Damage.Turbocharger(intEngine) = False Then
                
                Damage.Turbocharger(intEngine) = True
                
                UpdateMessage "#" & intEngine & " turbocharger malfunction."
                Call DropOutOfFormation
            
                If AlpsDirection() = ALPS_BELOW Then
                    ' Bomber must maintain high altitude over the Alps. Since
                    ' it cannot, the crew must bailout.
                    UpdateMessage "Bomber cannot maintain altitude in Alps."
                    G6ControlledBailout (False)
                    MechanicalFailure = END_MISSION
                Else
                    Call LoseAltitude
                End If
            
            End If
        
        Case 4: ' Heating System
        
            If Damage.HeatingSystem = False Then
            
                Damage.HeatingSystem = True
            
                For intPos = PILOT To AMMO_STOCKER
                
                    If PosExists(intPos) = True Then
                        Damage.Heater(intPos) = True
                    End If
                
                Next intPos
        
                UpdateMessage "Heating System: Complete malfunction."
                
            End If
        
        Case 5: ' Fuel Transfer

            If Damage.FuelTransferSystem = False Then
            
                Damage.FuelTransferSystem = True
            
                Select Case Random1D6()
                    Case 1, 2: Bomber.FuelPoints = 4
                    Case 3, 4: Bomber.FuelPoints = 3
                    Case 5, 6: Bomber.FuelPoints = 2
                End Select
                
                UpdateMessage "Fuel transfer malfunction: Bomber may fly " & _
                              Bomber.FuelPoints & " more zones."
        
            End If

        Case 6: ' Oil Tank
            
            intEngine = RandomDX(4)
            
            If OilTankHit(intEngine, strEffect) = END_MISSION Then
                MechanicalFailure = END_MISSION
            End If

            UpdateMessage strEffect
        
        Case 7: ' Intercom
    
            If Damage.IntercomSystem = False Then
            
                Damage.IntercomSystem = True
                UpdateMessage "Intercom: Complete malfunction."
                
            End If
        
        Case 8: ' Oxygen System
    
            If Damage.OxygenSystem = False Then
            
                Damage.OxygenSystem = True
                
                UpdateMessage "Oxygen System: Complete malfunction."
                
            End If
        
        Case 9: ' Electrical System
            
            If Bomber.RabbitsFoot >= 1 Then
                ' Expend luck to prevent loss of aircraft.
                UpdateMessage "Electrical System: Throws sparks & smoke, but luckily does not malfunction."
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
            Else
            
                Damage.Electrical = True
                UpdateMessage "Electrical System: Complete malfunction."
                G6ControlledBailout (OverWater())
                MechanicalFailure = END_MISSION
            
            End If
        
        Case 10, 11: ' Random Gun Jam
    
            ' This is a modification of the variant, which only accounted for
            ' B-17Fs. Rather than a turret power malfunction requiring 6 to
            ' hit due to manual traverse, a random gun will jam.

            intGun = RandomDX(TAIL_MG)

            If Bomber.Gun(intGun).Status = MG_OKAY _
            And GunOccupied(intGun) = True Then
            
                Bomber.Gun(intGun).Status = MG_JAMMED
                
                UpdateMessage lblGunName(intGun).Caption & " jammed."
            
            End If
        
        Case 12: ' Bomb Release Mechanism
    
            If Bomber.BomberModel <> YB40 Then
            
                Damage.BombRelease = True
                ' No message. The crew wouldn't know this until they attempt
                ' to drop their ordnance.
            
            End If
    
    End Select

End Function

'******************************************************************************
' ExitPrevZone
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  All movement and combat in the current zone is complete. Consume
'         fuel and do other housekeeping chores.
'******************************************************************************
Private Function ExitPrevZone() As Integer
    Dim intPos As Integer
    Dim intEng As Integer
    Dim blnOxygenOut As Boolean
    Dim blnHeaterOut As Boolean
    Dim intIndex As Integer
    Dim intAlpsDirection As Integer
    
    intAlpsDirection = AlpsDirection()
    
    For intEng = 1 To 4
        
        If Damage.OilTankLeak(intEng) = NO_OIL _
        And Damage.EngineOut(intEng) = False Then
            
            UpdateMessage "#" & intEng & " engine shut down due to oil tank hit."
            Damage.EngineOut(intEng) = True
        
            If Damage.FeatheringCtrl Then
                'See BL-1 result 9.
                
                UpdateMessage "#" & intEng & " engine not feathered."
                Damage.EngineDrag(intEng) = True
                
                If Bomber.InFormation = True Then
                    Call DropOutOfFormation
                End If
            
            End If
            
        End If
    
    Next intEng
    
    If Bomber.Altitude = HIGH_ALTITUDE Then
    
        If Damage.OxygenSystem = True Then
            blnOxygenOut = True
        Else
        
            ' Cycle through the existing positions on the bomber.
            For intPos = PILOT To AMMO_STOCKER
                
                If PosOccupied(intPos) = True Then
                    
                    ' Airman currently in position
                    intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
                    
                    If Bomber.Airman(intIndex).Status <= SW_STATUS Then
                        
                        If Damage.Oxygen(intPos) >= 2 Then
                            
                            ' If the airman is alive, he must have air.
                    
                            blnOxygenOut = True
                            Exit For
                    
                        ElseIf Bomber.Airman(intIndex).Frostbite = False _
                        And Damage.Heater(intPos) = True Then
                            
                            ' If the airman is alive, he should have heat.
                    
                            blnHeaterOut = True
                            Exit For
                        
                        End If
                    
                    End If
                    
                End If
            
            Next intPos
        
        End If

        If blnOxygenOut = True Then
'            ' Oxygen is out at a manned position(s): The bomber must descend.
'
'            Call DropOutOfFormation
'
            If intAlpsDirection = ALPS_BELOW Then

                ' Bomber must maintain high altitude over the Alps. Since
                ' it cannot, the crew must bailout.
                UpdateMessage "Bomber cannot maintain altitude in Alps."
                G6ControlledBailout (False)
                ExitPrevZone = END_MISSION
                Exit Function

'            ElseIf intAlpsDirection = ALPS_NEXT_ZONE Then
'
'                UpdateMessage "Bomber cannot maintain altitude in Alps."
'                Call BailOrCrash
'                ExitPrevZone = END_MISSION
'                Exit Function
'
            Else
                
                Call DropOutOfFormation
                Call LoseAltitude
                Call RefreshMissionForm
            
            End If
        
        ElseIf blnHeaterOut = True Then
        
'            If intAlpsDirection = ALPS_BELOW _
'            Or intAlpsDirection = ALPS_NEXT_ZONE Then
            If intAlpsDirection = ALPS_NEXT_ZONE Then
            
                UpdateMessage "Bomber must maintain altitude entering Alps."
            
            ElseIf intAlpsDirection = ALPS_BELOW Then
            
                UpdateMessage "Bomber must maintain altitude over Alps."
            
            Else
            
                Call DropToLowAltitude
                Call RefreshMissionForm
            
            End If
        
        End If
        
    ElseIf Bomber.Altitude = LOW_ALTITUDE Then
    
'        If intAlpsDirection = ALPS_AHEAD _
'        And Bomber.Direction = OUTBOUND Then
'            UpdateMessage "Bomber can't gain altitude to cross Alps."
'            Call BomberAbort
'            Exit Function
'        ElseIf intAlpsDirection = ALPS_NEXT_ZONE _
'        And Bomber.Direction = RETURN_TRIP Then
        If intAlpsDirection = ALPS_NEXT_ZONE _
        And Bomber.Direction = RETURN_TRIP Then
            UpdateMessage "Bomber can't gain altitude to re-cross Alps."
            Call BailOrCrash
            ExitPrevZone = END_MISSION
            Exit Function
        End If
        
    End If

    If Bomber.FuelPoints = 0 Then
        
        UpdateMessage "Fuel tanks empty."
        Call BailOrCrash
        ExitPrevZone = END_MISSION
        Exit Function
    
    ElseIf Bomber.CurrentZone = Damage.LastZone Then
        
        UpdateMessage "Bomber can't maintain lift."
        Call BailOrCrash
        ExitPrevZone = END_MISSION
        Exit Function
    
    ElseIf Bomber.FuelPoints = 1 _
    And intAlpsDirection = ALPS_NEXT_ZONE _
    And Bomber.Direction = RETURN_TRIP Then
    
        UpdateMessage "Bomber doesn't have enough fuel to re-cross Alps."
        Call BailOrCrash
        ExitPrevZone = END_MISSION
        Exit Function
    
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    frmMainMenu.Show
End Sub

Private Sub mnuHelpAbout_Click()
    ' Mission Screen
    frmAbout.Show
End Sub

Private Sub mnuHelpIndex_Click()
    ' Mission Screen
'    frmHelpBrowser.Hide
    
    frmHelpBrowser.txtPageName.Text = "doc/B17HelpIndex.html"

    frmHelpBrowser.Show
End Sub

'******************************************************************************
' optOneTemp_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Dynamically created control. The control must exist for this routine
'         to work.
'******************************************************************************
Private Sub optOneTemp_Click()

    If optOneTemp.Value = True Then
        
        If fraTemp.Caption = BAIL_OR_CRASH Then
            cmdInterrupt.Caption = optOneTemp.Caption
        ElseIf fraTemp.Caption = ABORT_MISSION Then
            cmdInterrupt.Caption = ABORT_MISSION
        ElseIf fraTemp.Caption = DESCEND_ALTITUDE Then
            cmdInterrupt.Caption = DESCEND_ALTITUDE
        ElseIf fraTemp.Caption = JETTISON_EXCESS Then
            cmdInterrupt.Caption = optOneTemp.Caption
        ElseIf fraTemp.Caption = QUIT_BEFORE_WATER Then
            cmdInterrupt.Caption = optOneTemp.Caption
        End If
        
    End If

End Sub

'******************************************************************************
' optTwoTemp_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Dynamically created control. The control must exist for this routine
'         to work.
'******************************************************************************
Private Sub optTwoTemp_Click()

    If optTwoTemp.Value = True Then
        
        If fraTemp.Caption = BAIL_OR_CRASH Then
            cmdInterrupt.Caption = optTwoTemp.Caption
        ElseIf fraTemp.Caption = ABORT_MISSION Then
            cmdInterrupt.Caption = "Continue Mission"
        ElseIf fraTemp.Caption = DESCEND_ALTITUDE Then
            cmdInterrupt.Caption = "Maintain Altitude"
        ElseIf fraTemp.Caption = JETTISON_EXCESS Then
            cmdInterrupt.Caption = optTwoTemp.Caption
        ElseIf fraTemp.Caption = QUIT_BEFORE_WATER Then
            cmdInterrupt.Caption = optTwoTemp.Caption
        End If
        
    End If

End Sub

'******************************************************************************
' RemoveTempOptions
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Delete dynamically created controls, then redisplay the wave frame.
'******************************************************************************
Private Sub RemoveTempOptions()

    ' Delete the controls now that we have the user's decision. (The user
    ' may change his choice prior to clicking the command button.)
    
    Me.Controls.Remove ("fraTemp")
    
    ' Put wave back in its proper place.
    
    fraWave.Visible = True
    
End Sub

'******************************************************************************
' DisplayTempOptions
'
' INPUT:  Captions for the frame and two option buttons.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Hide the wave frame, then dynamically create controls.
'******************************************************************************
Private Sub DisplayTempOptions(ByVal strFrameCaption As String, ByVal strOptOneCaption As String, ByVal strOptTwoCaption As String)

    ' The temporary frame will be displayed in wave's place.
    
    fraWave.Visible = False
    
    ' Create the necessary controls on the fly.
    
    Set fraTemp = Me.Controls.Add("VB.Frame", "fraTemp", Me)
    fraTemp.Caption = strFrameCaption
    fraTemp.Width = 1575
    fraTemp.Left = (((cmdInterrupt.Left * 2) + cmdInterrupt.Width) / 2) - (fraTemp.Width / 2)
    fraTemp.Height = 975
    fraTemp.Top = (fraWave.Top + fraWave.Height) - 975
    fraTemp.BackColor = PaleRed()
    
    Set optOneTemp = Me.Controls.Add("VB.OptionButton", "optOneTemp", fraTemp)
    optOneTemp.Caption = strOptOneCaption
    optOneTemp.Width = 1335
    optOneTemp.Left = 120
    optOneTemp.Height = 255
    optOneTemp.Top = 240
    optOneTemp.BackColor = PaleRed()
    optOneTemp.Value = True
    
    Set optTwoTemp = Me.Controls.Add("VB.OptionButton", "optTwoTemp", fraTemp)
    optTwoTemp.Caption = strOptTwoCaption
    optTwoTemp.Width = 1335
    optTwoTemp.Left = 120
    optTwoTemp.Height = 255
    optTwoTemp.Top = 600
    optTwoTemp.BackColor = PaleRed()
    optTwoTemp.Value = False
    
    fraTemp.Visible = True
    optOneTemp.Visible = True
    optTwoTemp.Visible = True

End Sub

'******************************************************************************
' DropToLowAltitude
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: If circumstances warrant, ask if the user wants to descend to low
'        altitude. The LoseAltitude() function actually does the descent.
'******************************************************************************
Public Sub DropToLowAltitude()
    Dim blnDescend As Boolean
    Dim intPos As Integer
    Dim intIndex As Integer
        
    Call DisplayTempOptions(DESCEND_ALTITUDE, "Yes", "No")
    
    ' Pause the program while the user decides whether or not to descend.

    Call Interrupt(cmdInterrupt, DESCEND_ALTITUDE)

    ' User clicked the next action button.
    
    If optOneTemp.Value = True Then
        blnDescend = True
    Else ' optTwoTemp was clicked
        blnDescend = False
    End If
    
    Call RemoveTempOptions

    If blnDescend = True Then
        Call DropOutOfFormation
        Call LoseAltitude
        Exit Sub
    End If

    Call BL5Frostbite

End Sub

'******************************************************************************
' AbortMission
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Forces abort, if necessary, otherwise gives the user the option of
'         aborting, if circumstances warrant. The user may not abort on a whim.
'******************************************************************************
Private Sub AbortMission()

    ' Rule 8.0

    Dim blnHeaterOut As Boolean
    Dim blnOxygenOut As Boolean
    Dim intEnginesOut As Integer
    Dim blnOilTankLeak As Boolean
    Dim intEng As Integer
    Dim intPos As Integer
    Dim intIndex As Integer
    Dim blnAbortMission As Boolean
    Dim blnWeatherAbort As Boolean
            
    blnHeaterOut = False
    blnOxygenOut = False
    intEnginesOut = 0
    blnOilTankLeak = False
    blnAbortMission = False
    blnWeatherAbort = False

    ' No sense in aborting if the plane is in the target zone or on the return
    ' trip, since returning (early) is the reason for aborting.
    
    If Bomber.Direction = OUTBOUND _
    And Bomber.CurrentZone < Mission.TargetZone Then
    
    
        If Bomber.BombsOnBoard = False _
        And Bomber.InFormation = False Then
            ' It may make sense to continue if the bomber is out of formation
            ' with bombs onboard, and it may make sense to continue if the
            ' bomber is in formation with no bombs, but it makes no sense
            ' to continue if the bomber is out of formation with no bombs.
            Call BomberAbort
            Exit Sub
        End If

'        If Bomber.BombsOnBoard = True Then

'    If Bomber.Direction = OUTBOUND _
'    And Bomber.CurrentZone < Mission.TargetZone Then
    
        ' Can't call CountEnginesOut() because we also need to check oil.
        
        For intEng = 1 To 4
            
            If Damage.EngineOut(intEng) = True Then
                intEnginesOut = intEnginesOut + 1
            End If
            
            If Damage.OilTankLeak(intEng) >= LT_LEAK Then
                blnOilTankLeak = True
            End If
            
        Next intEng
        
        ' Rule 8.0.d-e and h: Player must abort.
                
        If intEnginesOut >= 2 Then
            
            UpdateMessage intEnginesOut & " engines out: Bomber must abort."
            Call BomberAbort
            Exit Sub
                
        ElseIf Damage.FuelTankHits >= 1 _
        Or Damage.FuelTransferSystem = True Then

            ' Though it isn't mentioned in the rules, there isn't much sense
            ' in continuing a mission when the bomber is leaking like a sieve.

            UpdateMessage "Fuel shortage: Bomber must abort."
            Call BomberAbort
            Exit Sub
            
        ElseIf Bomber.InFormation = False Then
            
            If Bomber.Airman(NAVIGATOR).Status >= SW_STATUS _
            Or (Bomber.Airman(PILOT).Status >= SW_STATUS _
            And Bomber.Airman(COPILOT).Status >= SW_STATUS) Then
                
                UpdateMessage "Personnel losses: Bomber must abort."
                Call BomberAbort
                Exit Sub
            
            End If
        
        End If
        
        If Damage.OxygenSystem = True Then
            blnOxygenOut = True
        Else
        
            ' Cycle through the existing positions on the bomber.
            For intPos = PILOT To AMMO_STOCKER
                
                If PosOccupied(intPos) = True Then
                    
                    ' Airman currently in position
                    intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
                    
                    If Bomber.Airman(intIndex).Status <= SW_STATUS Then
                        
                        ' If the airman is alive, he must have air and should have heat.
                    
                        If Damage.Oxygen(intPos) >= 2 Then
                            blnOxygenOut = True
                            Exit For
                        ElseIf Damage.Heater(intPos) = True Then
                            blnHeaterOut = True
                            Exit For
                        End If
                    
                    End If
                    
                End If
            
            Next intPos
        
        End If
        
        ' Alps check block
        
        If Mission.AlpsZone <> 0 Then
        
            If blnOxygenOut = True Then
                    
                If AlpsDirection() = ALPS_AHEAD _
                Or AlpsDirection() = ALPS_NEXT_ZONE Then
                
                    UpdateMessage "Alps ahead and oxygen shortage: Bomber must abort."
                    Call BomberAbort
                    Exit Sub
                End If
            
            ElseIf Bomber.CurrentZone = (Mission.AlpsZone - 1) _
            And (Mission.Zone(Mission.AlpsZone).Weather = BAD_WEATHER _
            Or Mission.Zone(Mission.AlpsZone).Weather = STORM_WEATHER) Then
                    
                If Bomber.InFormation = True Then
                    UpdateMessage "Severe weather over Alps; entire group aborts."
                    Call GroupAbort
                Else
                    UpdateMessage "Severe weather over Alps; bomber must abort."
                    Call BomberAbort
                End If
                    
                Exit Sub
                
            ElseIf Bomber.CurrentZone = (Mission.AlpsZone - 1) _
            And Mission.Zone(Mission.AlpsZone).Weather = POOR_WEATHER Then
                
                If Bomber.InFormation = True Then
                    ' If the player chooses to abort, assume the entire group
                    ' also chose to abort.
                    UpdateMessage "Dense fog over Alps; group may abort."
                Else
                    UpdateMessage "Dense fog over Alps; bomber may abort."
                End If
                
                blnWeatherAbort = True
                
            ElseIf Bomber.Altitude = LOW_ALTITUDE Then
                
                If AlpsDirection() = ALPS_AHEAD _
                Or AlpsDirection() = ALPS_NEXT_ZONE Then
                
                    UpdateMessage "Alps ahead and low altitude: Bomber must abort."
                    Call BomberAbort
                    Exit Sub
                
                End If
            
            End If

        End If
        
        ' Rule 8.0.a-c and f-h: Player may choose to abort.

        If Damage.BombBayDoors = True _
        Or Damage.IntercomSystem = True _
        Or Damage.BombSight = True _
        Or blnHeaterOut = True _
        Or Bomber.Airman(BOMBARDIER).Status >= SW_STATUS _
        Or Bomber.InFormation = False _
        Or blnOxygenOut = True _
        Or intEnginesOut = 1 _
        Or blnOilTankLeak = True _
        Or blnWeatherAbort = True Then

            Call DisplayTempOptions(ABORT_MISSION, "Yes", "No")
            
            ' Pause the program while the user decides whether or not to abort.
    
            Call Interrupt(cmdInterrupt, ABORT_MISSION)
    
            ' User clicked the next action button.
            
            If optOneTemp.Value = True Then
                blnAbortMission = True
            Else ' optTwoTemp was clicked
                blnAbortMission = False
            End If
            
            Call RemoveTempOptions
        
            If blnAbortMission = True Then
                
                If blnWeatherAbort = True Then
                
                    If Bomber.InFormation = True Then
                        UpdateMessage "Entire group aborts due to weather."
                        Call GroupAbort
                    Else
                        UpdateMessage "Bomber aborts due to weather."
                        Call BomberAbort
                    End If
                    
                    Exit Sub
                
                Else
                    
                    UpdateMessage "Bomber aborts mission."
                    Call BomberAbort
                
                End If
            
            Else
            
                If intEnginesOut = 1 Then
                    ' By electing to continue the mission, with bombs, the player
                    ' has by default chosen to drop out of formation.
                    Call DropOutOfFormation
                End If
            
            End If
            
        End If
    
    End If

End Sub

'******************************************************************************
' BomberAbort
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  The bomber aborts, while the rest of the group continues.
'******************************************************************************
Private Sub BomberAbort()

    If Bomber.InFormation = True Then
        ' Bomber is in formation, so it must drop out before it can abort.
        lblSquadronPos.Caption = "Abort"
        Call DropOutOfFormation
        Bomber.Direction = RETURN_TRIP
        Call JettisonPayload(True, False)
        Call RefreshMissionInfo
    Else
        ' Bomber is already out of formation, so it can abort right away.
        lblSquadronPos.Caption = "Abort"
        Bomber.Direction = RETURN_TRIP
        Call JettisonPayload(True, False)
        Call RefreshMissionInfo
    End If

End Sub

'******************************************************************************
' GroupAbort
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  The entire group, including the bomber, is aborting.
'******************************************************************************
Private Sub GroupAbort()

    Call JettisonPayload(True, False)
    Bomber.Direction = RETURN_TRIP
    Call RefreshMissionInfo

End Sub

'******************************************************************************
' TakeOff
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if successful, otherwise false.
'
' NOTES:  For B-17s and Lancasters, this is routine, but B-24s must roll for
'         success due to their tendency to have a hard time getting airborne.
'******************************************************************************
Private Function TakeOff() As Boolean
    Dim intRoll As Integer
    Dim blnPayloadExploded As Boolean
    
    TakeOff = False
    
    ' All bombers other than the B-24 are assume to safely takeoff. The B-24,
    ' due to overloading, may crash on takeoff, so it must be checked.
    
    UpdateMessage Bomber.Name & " takes off."
    
    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
    
        ' Basically, this is the G9LandingOnLand procedure, but only with
        ' modifiers for weight, weather and pilot/copilot skill.
    
        intRoll = Random2D6()
        
        ' Adjust the roll for the extra weight carried by the B-24.
        
        intRoll = intRoll - 1
        
        ' O-1 Weather: Note a and Note b, plus "The General" (Volume 24, #6)
        ' variant.
        
        If Mission.Zone(BASE_ZONE).Weather = POOR_WEATHER Then
            intRoll = intRoll - 1
        ElseIf Mission.Zone(BASE_ZONE).Weather = BAD_WEATHER Then
            intRoll = intRoll - 2
        ElseIf Mission.Zone(BASE_ZONE).Weather = STORM_WEATHER Then
            intRoll = intRoll - 3
        End If
            
        ' Adjust for pilot and copilot experience.
            
        If Bomber.Airman(PILOT).Mission >= 11 _
        And Bomber.Airman(COPILOT).Mission >= 11 Then
            intRoll = intRoll + 1
        End If
        
        If intRoll <= 1 Then
            UpdateMessage Bomber.Name & " fails to clear end of runway."
        End If

        If intRoll <= 0 _
        And Random1D6() = 6 Then
            
            If Bomber.BombsOnBoard = True Then
                ' Note e.
                UpdateMessage "Bombs still aboard!"
                blnPayloadExploded = True
            ElseIf Bomber.ExtraFuelInBombBay = True Then
                blnPayloadExploded = True
            ElseIf Bomber.ExtraAmmo > 0 Then
                UpdateMessage "Extra ammo still aboard!"
                blnPayloadExploded = True
            End If
                
        End If
            
        If blnPayloadExploded = True Then
            
            Bomber.Status = CRASHED_STATUS
            UpdateMessage "Explosion. Bomber destroyed."
            Call CrewFinish(KIA_STATUS)
            TakeOff = False
            Exit Function
        
        Else

            Select Case intRoll
                    
                Case Is <= -3:
                        
                    Bomber.Status = CRASHED_STATUS
                    UpdateMessage "Bomber wrecked."
                    Call CrewFinish(KIA_STATUS)
                
                Case -2:
                        
                    Bomber.Status = CRASHED_STATUS
                    UpdateMessage "Bomber wrecked."
                    Call CrewFinish(BAD_CRASH_STATUS)
                    
                Case -1:
                        
                    Bomber.Status = CRASHED_STATUS
                    UpdateMessage "Bomber wrecked."
                    Call CrewFinish(CRASHED_STATUS)
                    
                Case 0:
                        
                    Bomber.Status = CRASHED_STATUS
                    UpdateMessage "Bomber crashed; irrepairably damaged."
                    
                Case 1:
                        
                    Bomber.Status = DUTY_STATUS
                    UpdateMessage "Bomber crashed; repairable by next mission."
                    
                Case Is >= 2:
                        
                    ' B-24 successfully takes off.
                    TakeOff = True
                    
            End Select
            
        End If
    
    Else
        
        ' B-17s and Lancasters always successfully takeoff.
        TakeOff = True
    
    End If
    
    If TakeOff = True Then
    
        Bomber.Altitude = HIGH_ALTITUDE
        
        ' B-24s that successfully took off, plus all other bomber models.
        
        If Bomber.BomberModel = AVRO_LANCASTER Then
            UpdateMessage Bomber.Name & " joins the bomber stream heading for Europe."
        Else
            UpdateMessage Bomber.Name & " forms up with squadron heading for Europe."
        End If

        If Mission.Options.RedTailAngels = True Then
            UpdateMessage "332nd Fighter Group, Red Tail Angels, are providing escort."
        End If
    
    End If
    
End Function

'******************************************************************************
' FlakCombat
'
' INPUT:  Flak intensity.
'
' OUTPUT: n/a
'
' RETURN: Number of flak hits, or end of mission.
'
' NOTES:  Called for both low level and target zone flak. Determines number of
'         flak hits, plus resulting damage.
'******************************************************************************
Private Function FlakCombat(ByVal intFlakLevel As Integer, ByVal blnTargetZoneCombat As Boolean) As Integer
    Dim intFlakBursts As Integer
    Dim intFlakHits As Integer
    Dim intCounter As Integer
    Dim intTemp As Integer

    FlakCombat = 0
    intFlakBursts = 0
    intFlakHits = 0
    intCounter = 0
    intTemp = 0

    ' Determine number of bursts in close proximity to the bomber.
    
    intFlakBursts = O3FlakToHitBomber(intFlakLevel, blnTargetZoneCombat)
    
    ' Determine the number of hits on the bomber due to those close bursts.
    
    For intCounter = 1 To intFlakBursts
        
        intTemp = O4EffectOfFlakHits()
        
        If intTemp = BURST_IN_PLANE Then
            ' Burst in plane negates all other flak hits due to a BIP's
            ' overwhelming damage.
            intFlakHits = BURST_IN_PLANE
            Exit For
        Else
            intFlakHits = intFlakHits + intTemp ' O4EffectOfFlakHits()
        End If
    
    Next intCounter

    If intFlakHits = BURST_IN_PLANE Then
        intFlakHits = 1
    End If
    
    UpdateMessage intFlakHits & " flak hits."

    ' Generate the damage due to those hits.
    
    If intFlakHits = BURST_IN_PLANE Then
    
        FlakCombat = BurstInPlane()
    
    Else
    
        For intCounter = 1 To intFlakHits
            FlakCombat = O5AreaAffectedByFlakHit()
            If FlakCombat = END_MISSION Then
                Exit Function
            End If
        Next intCounter

    End If

    FlakCombat = intFlakHits
    
End Function

'******************************************************************************
' NormalZoneCombat
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Determine number of waves. Only call AirToAirCombat() if at least one
'         wave should be generated.
'******************************************************************************
Private Function NormalZoneCombat() As Integer
    Dim intNumberOfWaves As Integer

    intNumberOfWaves = 0
    
    ' Fighter cover, for each zone, outbound and return trip, was determined
    ' when the mission was generated. Determine number of waves.
        
    If Bomber.BomberModel = AVRO_LANCASTER Then

        intNumberOfWaves = B1TameBoarWave()

    Else
        
        intNumberOfWaves = B1NumberOfGermanFighterWaves()
    
    End If
    
    If intNumberOfWaves >= 1 Then
        NormalZoneCombat = AirToAirCombat(intNumberOfWaves, False)
    Else
        Call RefreshMissionForm ' bernie
    End If

'DEBUG block
'If Bomber.CurrentZone = Mission.AlpsZone + 1 Then
'    Dim a
'    a = 1
'    Damage.OxygenSystem = True
'    Damage.EngineOut(1) = True
'    Damage.EngineOut(3) = True
'    Damage.FuelTransferSystem = True
'    Bomber.InFormation = True
'    Bomber.InFormation = False
'    Bomber.Airman(NAVIGATOR).Status = SW_STATUS
'    Damage.Oxygen(3) = 2
'    Mission.Zone(Bomber.CurrentZone + 1).Weather = BAD_WEATHER
'    Bomber.Altitude = LOW_ALTITUDE
'    Mission.Zone(Mission.AlpsZone).Weather = POOR_WEATHER
'    Mission.Options.MechanicalFailures = True
'    Damage.Turbocharger(1) = True
'    Damage.Heater(4) = True
'    Bomber.Altitude = LOW_ALTITUDE
'    Bomber.FuelPoints = 0
'End If

End Function

'******************************************************************************
' TargetZoneCombat
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, or default value.
'
' NOTES:  Conduct flak combat, then Wild Boar combat (for Lancasters), then
'         perform bomb run.
'******************************************************************************
Private Function TargetZoneCombat() As Integer
    Dim intFlakHits As Integer
    Dim intFlakLevel As Integer
    Dim blnOnTarget As Boolean
    Dim strResult As String
    Dim intPctHit As Integer
    
    If Bomber.BomberModel = AVRO_LANCASTER Then
        
        Bomber.SpottedBySearchLight = SpottedBySearchLight()

    End If

    intFlakLevel = O2FlakOverTarget(False)
    
    If intFlakLevel >= LIGHT_FLAK Then
        intFlakHits = FlakCombat(intFlakLevel, True)
        
        TargetZoneCombat = intFlakHits
        
        If TargetZoneCombat = END_MISSION Then
            Exit Function
        End If
    End If

    ' Regular air-to-air combat has already been conducted in the target
    ' zone, so the number of waves for non-Lancasters will always be 0.
        
    If Bomber.BomberModel = AVRO_LANCASTER Then

        ' Wild Boar attacks occurred after Lancasters entered the flak box.

        If B1WildBoarWave() = 1 Then
            ' Reflect flak damage prior to the Wild Boar attack.
            Call RefreshMissionForm

            TargetZoneCombat = AirToAirCombat(1, True)
            
            If TargetZoneCombat = END_MISSION Then
                Exit Function
            End If
        End If

    End If

    If Damage.BombBayDoors = True Then
    
        UpdateMessage "Bombs could not be released due to door damage."
    
    ElseIf Bomber.BomberModel = YB40 Then
    
        UpdateMessage "Escorted formation drops bombs."
    
    Else

        blnOnTarget = O6BombRun(intFlakHits)
        
        If blnOnTarget = False Then
            If BombNeutralTerritory() = True Then
                UpdateMessage "You just dropped your bombs on neutral Switzerland!"
                Exit Function
            End If
        End If
        
        intPctHit = O7BombingAccuracy(blnOnTarget)
        
        If blnOnTarget = True Then
            strResult = "On"
        Else
            strResult = "Off"
        End If
        
        lblBombRun.Caption = strResult & " / " & intPctHit
        
        UpdateMessage "Bomb run was " & LCase(strResult) & " target: " & _
                      intPctHit & "%."
        
        Bomber.BombsOnBoard = False
    
    End If
    
    If Damage.EngineOut(1) = True _
    And Damage.EngineOut(2) = True _
    And Damage.EngineOut(3) = True _
    And Damage.EngineOut(4) = True Then
        Call BailOrCrash
        TargetZoneCombat = END_MISSION
    End If
    
End Function

'******************************************************************************
' BombNeutralTerritory
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if neutral territory was bombed, otherwise false.
'
' NOTES:  This function is only called if the bombs fell off target. If the
'         bomber is in formation and the weather is clear, bombing neutral
'         territory is impossible (due to modifiers). The base chance is 1%.
'         If the bomber is out formation, and there are navigation problems,
'         and the weather is bad, then the chance goes up to 5%.
'******************************************************************************
Private Function BombNeutralTerritory() As Boolean
    Dim intRoll As Integer
    
    BombNeutralTerritory = False
    
    ' Only targets close to Switzerland (neutral territory) may be accidentally
    ' bombed.
    
    If Mission.TargetName = "Friedrichshaven" _
    Or Mission.TargetName = "Mulhouse" Then
    
        intRoll = RandomD100()
    
        If Bomber.InFormation = False Then
        
            intRoll = intRoll - 1
            
            If Damage.NavigationEquipment = True _
            Or Bomber.Airman(NAVIGATOR).Status >= SW_STATUS Then
                intRoll = intRoll - 1
            End If
                
        End If
            
        ' O-1 Weather: Note a and Note b, plus "The General" (Volume 24, #6)
        ' variant.
        If Mission.Zone(Bomber.CurrentZone).Weather = CLEAR_WEATHER Then
            intRoll = intRoll + 1
        ElseIf Mission.Zone(Bomber.CurrentZone).Weather = POOR_WEATHER Then
            intRoll = intRoll - 1
        ElseIf Mission.Zone(Bomber.CurrentZone).Weather = BAD_WEATHER Then
            intRoll = intRoll - 2
        End If
        
        If intRoll <= 1 Then
            BombNeutralTerritory = True
        End If
    
    End If

End Function

'******************************************************************************
' AirToAirCombat
'
' INPUT:  Number of waves; whether or not this is a Lancaster's Wild Boar phase.
'
' OUTPUT: n/a
'
' RETURN: End of mission, or default value.
'
' NOTES:  This routine drives all combat between the bomber, friendly fighters
'         and enemy fighters. It truly is the meat of the "B-17 Queen of the
'         Skies" gaming experience.
'******************************************************************************
Private Function AirToAirCombat(ByVal intNumberOfWaves As Integer, ByVal blnWildBoarPhase As Boolean) As Integer

    Dim intWaveCount As Integer
    Dim intFightersInWave As Integer
    Dim intTotalHits As Integer
    Dim intHit As Integer
    Dim strCover As String
    Dim intRemovals As Integer

    intWaveCount = 0
    intFightersInWave = 0
    intTotalHits = 0
    intHit = 0
    strCover = ""
    intRemovals = 0
    
    ' Determine composition of each wave, conducting any combat before
    ' proceeding to the next wave.
    
    For intWaveCount = 1 To intNumberOfWaves
    
'        gintCurrGun = 0
'        gintCurrTarget = 0
    
        chkEvadeFighters.Value = vbUnchecked
        chkEvadeFighters.Enabled = False
        
        ' Create the wave.
        
        If Bomber.BomberModel = AVRO_LANCASTER Then
    
            If blnWildBoarPhase = True Then
                intFightersInWave = B3WildBoarFighter()
            Else
                intFightersInWave = B3TameBoarFighter()
            End If

        Else
            
            intFightersInWave = B3AttackingFighterWave()
        
        End If
    
        If intFightersInWave = END_MISSION Then
            AirToAirCombat = END_MISSION
            Exit Function
        End If
        
        ' Sometimes a "wave" may be a random event, or may otherwise consist
        ' of no fighters.
        
        If intFightersInWave >= 1 Then
        
            UpdateMessage "" ' Blank line ' xyz
        
            If Wave.Fighter(1).Special = "TameBoar" _
            And Wave.Fighter(1).Position = F6_LOW Then
                
                ' Surprise attack. Fighter gets a free shot prior to defensive
                ' fire and normal combat.

                UpdateMessage "Surprise attack!"

                If OffensiveFire(True) = END_MISSION Then
                    AirToAirCombat = END_MISSION
                    Exit Function
                End If

                Call RefreshMissionForm
        
            End If
        
            ' Normal attacks.
            
            Wave.Attack = 1
            
            Do While GetWaveSize() >= 1 _
            And Wave.Attack <= 3
            
                gintCurrGun = 0
                gintCurrTarget = 0
                
                UpdateMessage "Wave " & intWaveCount & _
                              " (Attack " & Wave.Attack & "): " & _
                              GetWaveSize() & " fighters"

                ' Display the wave's initial attack, or redisplay it for
                ' successive attacks.
            
                Call DisplayWave

                chkEvadeFighters.Value = vbUnchecked
                chkEvadeFighters.Enabled = False
        
                ' Fighter cover phase.
                
                If Bomber.Direction = OUTBOUND Then
                    strCover = Mission.Zone(Bomber.CurrentZone).CoverOut
                Else
                    strCover = Mission.Zone(Bomber.CurrentZone).CoverBack
                End If
                
                UpdateMessage "Fighter cover: " & strCover
                
                intRemovals = M4FighterCoverDefense(strCover)
                
'UpdateMessage "intRemovals = [" & intRemovals & "]" ' DEBUG
    
                If intRemovals >= 1 Then
                    intRemovals = GetMaxRemovals(intRemovals)
                End If
                
'UpdateMessage "intRemovals = [" & intRemovals & "]" & vbCrLf ' DEBUG

                If intRemovals >= 1 Then
                
                    ' Do not allow the user to continue until the required
                    ' number of fighters are marked for removal.
                    
                    cmdInterrupt.Enabled = False
                    
                    ' Indicate how many fighters must be removed.
                    
                    gintRemovalsRemaining = intRemovals
                
                    lblMiscWave.Caption = "Remove " & gintRemovalsRemaining & _
                                          " Enemy Fighters"
    
                    ' Pause the program while the user selects the fighters to be
                    ' removed.
                    
                    Call Interrupt(cmdInterrupt, REMOVE_ENEMY_FIGHTERS)
                    
                    ' User clicked the next action button.
                    ' Remove fighters driven off by friendly cover; leave the
                    ' survivors in their current spot.

                    Call RemoveFightersFromWave(False, False)
                    
                    Call DisplayWave
                
                End If

                UpdateMessage intRemovals & " enemy fighters chased off by cover"
                
                ' Defensive fire phase.

                If GetWaveSize() >= 1 Then
                
                    ' Indicate which guns have a possible target.
                    
                    Call PreAssignGuns
                    
                    ' Pause the program while the user selects the weapons
                    ' to be fired, and whether or not the bomber will take
                    ' evasive action.

                    If EvasiveActionAllowed() = True Then
                        chkEvadeFighters.Enabled = True ' blahblahblah
                    Else
                        chkEvadeFighters.Value = vbUnchecked
                        chkEvadeFighters.Enabled = False
                    End If
        
                    Call Interrupt(cmdInterrupt, FIRE_GUNS)
    
                    ' User clicked the next action button.
                    
                    Call DefensiveFire
                    
                    ' Remove fighters shot down by the bomber's guns; leave the
                    ' survivors in their current spot.
                    Call RemoveFightersFromWave(False, False)
                    
                    Call DisplayWave
                
                    Call RefreshCrew
Call RefreshGuns
                
                End If
' asd
                gintCurrGun = 0
                gintCurrTarget = 0
                
                ' Offensive fire phase.
                
                If GetWaveSize() >= 1 Then
                
                    If OffensiveFire(False) = END_MISSION Then
                        AirToAirCombat = END_MISSION
                        Exit Function
                    End If

                    Call RefreshMissionForm
        
                End If
                
                Wave.Attack = Wave.Attack + 1
           
                ' Passing fire phase.
                
                If GetWaveSize() >= 1 _
                And PassingFireAllowed() = True Then
                
                    Call PassingFireSetup
                
                    ' Pause the program while the user decides whether or not
                    ' to use passing fire and, if so, which target.

                    Call Interrupt(cmdInterrupt, PASSING_FIRE)
    
                    ' User clicked the next action button.
                    
                    Call DefensiveFire
                    
                End If
                
                gintCurrGun = 0
                gintCurrTarget = 0
    
                ' Damaged fighters, and those which missed their shots,
                ' break off contact.
                
                ' Remove fighters shot down by passing fire, FBOA by normal or
                ' passing fire, or which missed their own attack. Reposition
                ' any survivors.
                Call RemoveFightersFromWave(True, True)
                
'                gintCurrGun = 0
'                gintCurrTarget = 0
    
                If chkSwapAmmo.Value = vbChecked Then
                    
                    ' Pause the program while the user decides whether or not to
                    ' swap ammo between weapons.
    
                    Call Interrupt(cmdInterrupt, SWAP_AMMO)
        
                    ' User clicked the next action button.
                
                    Call SwapAmmo
                    
                End If
                    
' TODO: Swap positions
' TODO: other functionality ???
           
            Loop
        
            ' The current wave is complete. Clear its display.
            
            Call ClearWaveDisplay(True) ' iop
        Else
            Call RefreshMissionForm ' bernie
        End If
        
        If Damage.EngineOut(1) = True _
        And Damage.EngineOut(2) = True _
        And Damage.EngineOut(3) = True _
        And Damage.EngineOut(4) = True Then
            Call BailOrCrash
            AirToAirCombat = END_MISSION
            Exit Function
        End If
    
    Next intWaveCount

    Call UnjamGuns
                
End Function

'******************************************************************************
' BailOrCrash
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  The mission has come to a premature conclusion. Give the user a
'         choice of bailing out or trying to land the bomber.
'******************************************************************************
Private Sub BailOrCrash()
    Dim blnOverWater As Boolean
    Dim blnBailOut As Boolean

    blnOverWater = OverWater()

    If blnOverWater = True Then
        Call DisplayTempOptions(BAIL_OR_CRASH, BAILOUT_WATER, DITCH_WATER)
    Else
        Call DisplayTempOptions(BAIL_OR_CRASH, BAILOUT_LAND, CRASH_LAND)
    End If
    
    ' Pause the program while the user decides whether or not
    ' to bailout or crashland.

    Call Interrupt(cmdInterrupt, BAIL_OR_CRASH)

    ' User clicked the next action button.
    
    If optOneTemp.Value = True Then
        blnBailOut = True
    Else ' optTwoTemp was clicked
        blnBailOut = False
    End If
    
    Call RemoveTempOptions

    If blnBailOut = True Then
        
        Call G6ControlledBailout(blnOverWater)
    
    Else

        Call JettisonPayload(True, False)
        
        If blnOverWater = True Then
            Call G10LandingInWater
        Else
            Call G9LandingOnLand
        End If
        
    End If
        
End Sub

'******************************************************************************
' SwapAmmo
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Launch the Swap Ammo form. Update the ammo display based on any
'         user-entered changes.
'******************************************************************************
Private Sub SwapAmmo()
    Dim intGun As Integer

    ' The user may not continue the game until they either swap ammo, or
    ' cancel the swap.
    chkSwapAmmo.Value = Unchecked
    
    frmSwapAmmo.Show vbModal
                
    Call RefreshGuns

End Sub

'******************************************************************************
' ClearWaveDisplay
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Clear out all the fighters prior to displaying a successive attacks
'         or a new wave.
'******************************************************************************
Private Sub ClearWaveDisplay(ByVal blnTypeAndPos As Boolean)
    Dim intFighter As Integer

    For intFighter = 1 To 6
        
        lblToHit(intFighter).Caption = ""
        lblToHit(intFighter).Tag = ""
        lblToHit(intFighter).BackColor = vbButtonFace

' iop
        If blnTypeAndPos = True Then
        
            lblType(intFighter).Caption = ""
            lblType(intFighter).Tag = ""
            lblType(intFighter).BackColor = vbButtonFace
        
            lblPosition(intFighter).Caption = ""
            lblPosition(intFighter).Tag = ""
            lblPosition(intFighter).BackColor = vbButtonFace

        End If

    Next intFighter

    lblMiscWave.Caption = ""
    
End Sub

'******************************************************************************
' PassingFireAllowed
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if passing fire is allowed, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Private Function PassingFireAllowed() As Boolean
    Dim intFighter As Integer

    PassingFireAllowed = False

    ' Rule 9.2
    
    If MayFireGun(TAIL_MG) = True _
    And Bomber.Gun(TAIL_MG).QualifiedGunner = True _
    And Damage.IntercomSystem = False Then

        ' Gunner is able to fire a passing shot.
        
        For intFighter = 1 To GetWaveSize()
        
            ' Gunner has a target to shoot at.
            
            If Wave.Fighter(intFighter).Position = F1030_HIGH _
            Or Wave.Fighter(intFighter).Position = F1030_LEVEL _
            Or Wave.Fighter(intFighter).Position = F1030_LOW _
            Or Wave.Fighter(intFighter).Position = F12_HIGH _
            Or Wave.Fighter(intFighter).Position = F12_LEVEL _
            Or Wave.Fighter(intFighter).Position = F12_LOW _
            Or Wave.Fighter(intFighter).Position = F130_HIGH _
            Or Wave.Fighter(intFighter).Position = F130_LEVEL _
            Or Wave.Fighter(intFighter).Position = F130_LOW Then
    
                PassingFireAllowed = True
                Exit For

            End If

        Next intFighter

    End If

End Function

'******************************************************************************
' PassingFireSetup
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub PassingFireSetup()
    Dim intGun As Integer
    Dim intFighter As Integer

    ' Clear the bomber's weapons.
    
    For intGun = MID_UPPER_MG To TAIL_MG
                    
        lblGunAmmo(intGun).BackColor = vbButtonFace
        lblGunAmmo(intGun).Tag = ""
    
        chkSpray(intGun).Enabled = False
        chkSpray(intGun).Value = vbUnchecked
        
    Next intGun

    ' Preselect the tail weapon.
    
'    lblGunAmmo(TAIL_MG).BackColor = MedDkCyan()
'    lblGunAmmo(TAIL_MG).BackColor = PaleCyan()
'    gintCurrGun = TAIL_MG

    For intFighter = 1 To GetWaveSize()
        
        ' Indicate potential passing fire targets.
        
        If Wave.Fighter(intFighter).Position = F1030_HIGH _
        Or Wave.Fighter(intFighter).Position = F1030_LEVEL _
        Or Wave.Fighter(intFighter).Position = F1030_LOW _
        Or Wave.Fighter(intFighter).Position = F12_HIGH _
        Or Wave.Fighter(intFighter).Position = F12_LEVEL _
        Or Wave.Fighter(intFighter).Position = F12_LOW _
        Or Wave.Fighter(intFighter).Position = F130_HIGH _
        Or Wave.Fighter(intFighter).Position = F130_LEVEL _
        Or Wave.Fighter(intFighter).Position = F130_LOW Then
            
            lblGunAmmo(TAIL_MG).BackColor = MedDkCyan()
            
            lblToHit(intFighter).Caption = "6"
            lblToHit(intFighter).BackColor = MedDkCyan()
        
        Else
            
            lblToHit(intFighter).Caption = ""
            lblToHit(intFighter).BackColor = vbButtonFace
        
        End If
    
    Next intFighter
    
End Sub

'******************************************************************************
' lblGunAmmo_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Select the gun, indicating the odds of hitting all targets in the
'         gun's field of fire.
'******************************************************************************
Private Sub lblGunAmmo_Click(Index As Integer)
    Dim intPrevGun As Integer
    Dim intPrevTarget As Integer
    Dim intNewGun As Integer
    Dim intFighter As Integer
    Dim intToHit As Integer
    Dim intTarget As Integer
    Dim intAirman As Integer
    Dim intAirmanStatus As Integer

Dim a
a = 1

    If cmdInterrupt.Caption = FIRE_GUNS _
    Or cmdInterrupt.Caption = PASSING_FIRE Then
        
        ' Operate on local variables.
        
        intPrevGun = gintCurrGun
        intPrevTarget = gintCurrTarget
        intNewGun = Index
    
        ' Set lblGuns' colors.
        
        If intNewGun = intPrevGun Then
        
            ' Re-selecting the same gun has no effect.
            Exit Sub
        
        Else
            
            If intPrevGun = 0 Then
                ' This current gun is the first one to be selected. Dummy the
                ' previous gun to avoid array subscript error.
                intPrevGun = 1
            End If
            
            If lblGunAmmo(intNewGun).BackColor = PaleCyan() Then
            
                ' This should only be possible if the user re-selected the
                ' same gun. This condition should have been handled above.
                
            ElseIf lblGunAmmo(intNewGun).BackColor = MedDkCyan() Then
                
                ' User selected a gun which may fire.
                
                If lblGunAmmo(intPrevGun).BackColor = PaleCyan() Then
                    
                    ' Unselect the old gun; select the new gun.
                    lblGunAmmo(intPrevGun).BackColor = MedDkCyan()
                    lblGunAmmo(intNewGun).BackColor = PaleCyan()
                
                ElseIf lblGunAmmo(intPrevGun).BackColor = MedDkCyan() Then
                    
                    ' Old gun wasn't selected. It was probably dummied from
                    ' 0 to 1, meaning the new gun is the first gun to be
                    ' selected. Select the new gun.
                    lblGunAmmo(intNewGun).BackColor = PaleCyan()
                    
                ElseIf lblGunAmmo(intPrevGun).BackColor = vbButtonFace Then
                    
                    ' Old gun can't be fired. Select the new gun.
                    lblGunAmmo(intNewGun).BackColor = PaleCyan()
                
                End If
                
            ElseIf lblGunAmmo(intNewGun).BackColor = vbButtonFace Then
                
                ' Newly selected gun can not fire.
            
                If lblGunAmmo(intPrevGun).BackColor = PaleCyan() Then
                    
                    ' Unselect the old gun; ignore the new gun.
                    lblGunAmmo(intPrevGun).BackColor = MedDkCyan()

'iop
Call ClearWaveDisplay(False) ' clear the to hit values only
'                    If intPrevTarget >= 1 Then
'                        ' The old gun had a target; unselect it as well.
'                        lblToHit(intPrevTarget).BackColor = MedDkCyan()
'                    End If
' iop

'                    gintCurrGun = 0
'                    gintCurrTarget = 0
                
                ElseIf lblGunAmmo(intPrevGun).BackColor = MedDkCyan() Then
                    
                    ' Old gun wasn't selected. It was probably dummied from
                    ' 0 to 1, meaning the new gun is the first gun to be
'                    ' selected. But the new gun can't fire, so simply exit.
'                    Exit Sub
                    
                ElseIf lblGunAmmo(intPrevGun).BackColor = vbButtonFace Then
                    
                    
                    ' Olg gun can't be fired; new gun can't be fired.

'                    ' Old gun couldn't have been selected because it isn't
'                    ' allowed to fire. The new gun can't fire either, so
'                    ' simply exit.
'                    Exit Sub
                
                End If
                
            End If
            
        End If

        If lblGunAmmo(intNewGun).BackColor <> vbButtonFace Then
        
            ' Determine who is actually manning the gun. If no one is manning
            ' the gun, switch an airman in, if possible.

            If GunManned(intNewGun) = True Then
                
                ' Temporary swap is not allowed. To gun must be unmanned to execute
                ' a temporary swap. If the gun is manned, then a slow swap may be
                ' performed after the wave attack.
                intAirman = Bomber.Gun(intNewGun).MannedBy
                
            Else
    
                If intNewGun = PORT_CHEEK_MG _
                Or intNewGun = STBD_CHEEK_MG Then
                
                    ' A position for which a quick swap can be made.
                    intAirman = QuickPosSwap(intNewGun)
                
                ElseIf (intNewGun = PORT_WAIST_MG _
                Or intNewGun = STBD_WAIST_MG) _
                And (Bomber.BomberModel = B24_D _
                Or Bomber.BomberModel = B24_E _
                Or Bomber.BomberModel = B24_GHJ _
                Or Bomber.BomberModel = B24_LM) Then
    
                    ' A position for which a quick swap can be made.
                    intAirman = QuickPosSwap(intNewGun)
                
                ElseIf (intNewGun = BALL_TURRET_MG _
                And (Bomber.BomberModel = B24_D _
                Or Bomber.BomberModel = B24_E)) Then
                
                    ' A position for which a quick swap can be made.
                    intAirman = QuickPosSwap(intNewGun)
                
                End If
            
            End If
    
'            If PosOccupied(intPos) = True Then
' this should have been done already, oterwise the gun would not be cyan colored
            
            ' Airman currently in position
            ' The airman's status may affect to the hit value.
            
            intAirmanStatus = Bomber.Airman(GetAirmanIndexBySerialNumber(intAirman)).Status
        
            ' The new gun has a possible target.
        
            For intFighter = 1 To GetWaveSize()
                
                ' Initialize each target to grey with no to hit value.
                
                lblToHit(intFighter).Caption = ""
                lblToHit(intFighter).BackColor = vbButtonFace
            
                ' Determine which fighters are in the new gun's field of fire. There
                ' should be at least one possible target, otherwise PreAssignGuns()
                ' would have made the new gun grey, meaning we would not have gotten
                ' to this point.
                
                If cmdInterrupt.Caption = FIRE_GUNS Then
                
                    ' Lookup to hit in the database.
                    
                    intToHit = M1DefensiveFire(Bomber.BomberModel, _
                                               intNewGun, _
                                               Wave.Fighter(intFighter).Position, _
                                               Wave.Fighter(intFighter).Type, _
                                               intAirmanStatus)
                    
                    If intToHit > 0 Then
                        
                        ' The new gun may hit this fighter.
                        
                        lblToHit(intFighter).Caption = CStr(intToHit)
                        lblToHit(intFighter).BackColor = MedDkCyan()
                        
                    End If
                
                ElseIf cmdInterrupt.Caption = PASSING_FIRE Then
                    
                    ' To hit is always 6 for passing fire.
                    
                    If lblGunAmmo(intNewGun).BackColor = PaleCyan() _
                    And (Wave.Fighter(intFighter).Position = F1030_HIGH _
                    Or Wave.Fighter(intFighter).Position = F1030_LEVEL _
                    Or Wave.Fighter(intFighter).Position = F1030_LOW _
                    Or Wave.Fighter(intFighter).Position = F12_HIGH _
                    Or Wave.Fighter(intFighter).Position = F12_LEVEL _
                    Or Wave.Fighter(intFighter).Position = F12_LOW _
                    Or Wave.Fighter(intFighter).Position = F130_HIGH _
                    Or Wave.Fighter(intFighter).Position = F130_LEVEL _
                    Or Wave.Fighter(intFighter).Position = F130_LOW) Then
                        
                        lblToHit(intFighter).BackColor = MedDkCyan()
                        lblToHit(intFighter).Caption = "6"
                    
                    End If
                    
                End If
            
            Next intFighter
        
        End If
        
        If lblGunAmmo(intNewGun).Tag <> "" Then
            
            ' The new gun was previously assigned a target. Left of the slash
            ' is the fighter; right of the slash is the to hit value.

            intTarget = CInt(Mid(lblGunAmmo(intNewGun).Tag, 1, 1))
            
            If lblToHit(intTarget).Caption <> "" Then
                lblToHit(intTarget).BackColor = PaleCyan()
            End If
            
        Else
        
            ' New gun doesn't or can't have a target.
            intTarget = 0

        End If
        
'        If lblGunAmmo(intNewGun).BackColor = PaleCyan() Then
'
'            ' The newly selected gun becomes the currently selected gun.
'            gintCurrGun = intNewGun
'            gintCurrTarget = intTarget
'
'        End If
        
        gintCurrGun = intNewGun
        gintCurrTarget = intTarget
    
    End If

a = 1

End Sub

'******************************************************************************
' QuickPosSwap
'
' INPUT:  The gun that the airman should be switched to.
'
' OUTPUT: n/a
'
' RETURN: The airman manning the gun being switched to. (The return value
'         should never be UNMANNED_MG.)
'
' NOTES:  This function alters the Mission form. The port cheek gun is initially
'         unmanned on the B-17F, YB-40 and B-24E. The port waist gun is initially
'         unmanned on the B-24D, B-24E, B-24G/H/J and B-24L/M. Both waist guns
'         may be manned at the same time; only one cheek gun may be manned at
'         one time. Otherwise, every position must be manned at all times. A
'         quick position swap may only be performed between these positions:
'             1) stbd -> port cheek
'             .) port -> stbd cheek
'             2) stbd -> port waist
'             .) port -> stbd waist
'             3) tunnel -> port waist
'             .) port waist -> tunnel
'         Quick swaps are undone by clicking on the original from position.
'******************************************************************************
Private Function QuickPosSwap(ByVal intToGun As Integer) As Integer
    Dim blnHasTunnelGun As Boolean
Dim a
a = 1
    
'    QuickPosSwap = Bomber.Gun(intToGun).MannedBy

'            x1) stbd -> port cheek
'            x.) port -> stbd cheek
'            x2) stbd -> port waist
'            x.) port -> stbd waist
'             3) tunnel -> port waist
'             .) port waist -> tunnel
    
'---( cheek )---------------------------------------------------------------

    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E Then
        ' The "ball turret" position on the B-24D and B-24E was actually a
        ' tunnel gun from which the tunnel gunner could quickly assume waist
        ' gunner duties.
        blnHasTunnelGun = True
    Else
        ' The ball turret really is a ball turret, rather than a tunnel gun,
        ' making a quick port waist swap is impossible.
        blnHasTunnelGun = False
    End If
    
    If GunOccupied(STBD_CHEEK_MG) = True _
    And intToGun = PORT_CHEEK_MG _
    And Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(NAVIGATOR).CurrentSerialNum)).Status <= LW2_STATUS Then ' navig avail
            
        ' Switch from stbd to port side.
        Bomber.Gun(PORT_CHEEK_MG).MannedBy = Bomber.Gun(STBD_CHEEK_MG).MannedBy
        
        ' Unman stbd side.
        Bomber.Gun(STBD_CHEEK_MG).MannedBy = UNMANNED_MG
        lblGunAmmo(STBD_CHEEK_MG).Tag = ""
        
        GoTo XitFunction
        
    ElseIf GunOccupied(PORT_CHEEK_MG) = True _
    And intToGun = STBD_CHEEK_MG _
    And Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(NAVIGATOR).CurrentSerialNum)).Status <= LW2_STATUS Then ' navig avail
            
        ' Switch from port to stbd side.
        Bomber.Gun(STBD_CHEEK_MG).MannedBy = Bomber.Gun(PORT_CHEEK_MG).MannedBy
        
        ' Unman port side.
        Bomber.Gun(PORT_CHEEK_MG).MannedBy = UNMANNED_MG
        lblGunAmmo(PORT_CHEEK_MG).Tag = ""
        
        GoTo XitFunction
        
    End If

'---( waist )---------------------------------------------------------------
    
    If intToGun = PORT_WAIST_MG Then ' to port
    
        If GunOccupied(PORT_WAIST_MG) = True Then ' port occ
        
            ' do nothing
            GoTo XitFunction
    
        End If
            
        If GunOccupied(BALL_TURRET_MG) = True Then ' ball occ
    
            If blnHasTunnelGun = True _
            And GunOccupied(BALL_TURRET_MG) = True _
            And Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(BALL_GUNNER).CurrentSerialNum)).Status <= LW2_STATUS Then
    
                ' tunn -> port (1/1)
                Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum = Bomber.Position(BALL_GUNNER).CurrentSerialNum
                Bomber.Gun(PORT_WAIST_MG).MannedBy = Bomber.Gun(BALL_TURRET_MG).MannedBy
                
                ' Unman tunnel.
                Bomber.Position(BALL_GUNNER).CurrentSerialNum = UNMANNED_POSITION
                Bomber.Gun(BALL_TURRET_MG).MannedBy = UNMANNED_MG
                lblGunAmmo(BALL_TURRET_MG).Tag = ""
    
                GoTo XitFunction
        
            End If
    
        End If
    
        If GunOccupied(STBD_WAIST_MG) = True Then ' stbd occ
        
            If Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum)).Status <= LW2_STATUS Then
    
                ' stbd -> port (1/2)
                Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum = Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum
                Bomber.Gun(PORT_WAIST_MG).MannedBy = Bomber.Gun(STBD_WAIST_MG).MannedBy
                
                ' Unman stbd side.
                Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum = UNMANNED_POSITION
                Bomber.Gun(STBD_WAIST_MG).MannedBy = UNMANNED_MG
                lblGunAmmo(STBD_WAIST_MG).Tag = ""
    
                GoTo XitFunction
        
            End If
    
        End If
    
    End If
    
    If intToGun = STBD_WAIST_MG Then ' to stbd
    
        If GunOccupied(STBD_WAIST_MG) = True Then ' stbd occ
        
            ' do nothing
            GoTo XitFunction
    
        End If
    
        If GunOccupied(PORT_WAIST_MG) = True Then ' port occ
    
            If Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum)).Status <= LW2_STATUS Then
    
                ' port -> stbd (2/3)
                Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum = Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum
                Bomber.Gun(STBD_WAIST_MG).MannedBy = Bomber.Gun(PORT_WAIST_MG).MannedBy
                
                ' Unman port side.
                Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum = UNMANNED_POSITION
                Bomber.Gun(PORT_WAIST_MG).MannedBy = UNMANNED_MG
                lblGunAmmo(PORT_WAIST_MG).Tag = ""
    
                GoTo XitFunction
        
            End If
    
        End If
    
        If GunOccupied(BALL_TURRET_MG) = True Then ' ball occ
    
            If blnHasTunnelGun = True _
            And GunOccupied(BALL_TURRET_MG) = True _
            And Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(BALL_GUNNER).CurrentSerialNum)).Status <= LW2_STATUS Then
    
                ' tunn -> stbd (2/1)
                Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum = Bomber.Position(BALL_GUNNER).CurrentSerialNum
                Bomber.Gun(STBD_WAIST_MG).MannedBy = Bomber.Gun(BALL_TURRET_MG).MannedBy
                
                ' Unman tunnel.
                Bomber.Position(BALL_GUNNER).CurrentSerialNum = UNMANNED_POSITION
                Bomber.Gun(BALL_TURRET_MG).MannedBy = UNMANNED_MG
                lblGunAmmo(BALL_TURRET_MG).Tag = ""
    
                GoTo XitFunction
        
            End If
    
        End If
    
    End If
    
    If intToGun = BALL_TURRET_MG Then ' to ball
    
        If GunOccupied(BALL_TURRET_MG) = True Then ' ball occ
        
            ' do nothing
            GoTo XitFunction
            
        End If
    
        If GunOccupied(STBD_WAIST_MG) = True Then ' stbd occ
        
            If Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum)).Status <= LW2_STATUS Then
    
                ' stbd -> tunn (3/2)
                Bomber.Position(BALL_GUNNER).CurrentSerialNum = Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum
                Bomber.Gun(BALL_TURRET_MG).MannedBy = Bomber.Gun(STBD_WAIST_MG).MannedBy
                
                ' Unman stbd side.
                Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum = UNMANNED_POSITION
                Bomber.Gun(STBD_WAIST_MG).MannedBy = UNMANNED_MG
                lblGunAmmo(STBD_WAIST_MG).Tag = ""
    
                GoTo XitFunction
        
            End If
    
        End If
    
        If GunOccupied(PORT_WAIST_MG) = True Then ' port occ
    
            If Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum)).Status <= LW2_STATUS Then
    
                ' port -> tunn (3/3)
                Bomber.Position(BALL_GUNNER).CurrentSerialNum = Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum
                Bomber.Gun(BALL_TURRET_MG).MannedBy = Bomber.Gun(PORT_WAIST_MG).MannedBy
                
                ' Unman port side.
                Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum = UNMANNED_POSITION
                Bomber.Gun(PORT_WAIST_MG).MannedBy = UNMANNED_MG
                lblGunAmmo(PORT_WAIST_MG).Tag = ""
    
                GoTo XitFunction
        
            End If
    
        End If
    
    End If

XitFunction:

'---------------------------------------------------------------------------

    Call RefreshCrew
            
    ' Return the serial number of the airman who is now firing the gun.

    QuickPosSwap = Bomber.Gun(intToGun).MannedBy

End Function
    
'******************************************************************************
' lblToHit_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Select the fighter to be fired at, or to be removed (due to friendly
'         cover).
'******************************************************************************
Private Sub lblToHit_Click(Index As Integer)
    Dim intPrevTarget As Integer
    Dim intNewTarget As Integer
    Dim intCurrGun As Integer
    
Dim a
a = 1
    
    If cmdInterrupt.Caption = REMOVE_ENEMY_FIGHTERS Then
      
        ' Operate on local variables.
        
        intNewTarget = Index
    
        If intNewTarget > GetWaveSize() Then
            ' User tried to remove a fighter which doesn't exist.
            Exit Sub
        End If
        
        If Wave.Fighter(intNewTarget).Special = "B" _
        Or Wave.Fighter(intNewTarget).Special = "D" Then
            ' Fighter can't be removed due its attack angle.
            Exit Sub
        End If
        
        If lblToHit(intNewTarget).Tag = REMOVE_FIGHTER Then
         
            ' Fighter was previously selected for removal. By selecting
            ' it again, the user is indicating that he changed his mind.
            ' Deselect the fighter.
            
            lblToHit(intNewTarget).BackColor = vbButtonFace
            lblToHit(intNewTarget).Tag = ""
            gintRemovalsRemaining = gintRemovalsRemaining + 1
      
        ElseIf gintRemovalsRemaining >= 1 Then
            
            ' Select the fighter for removal.
            
            lblToHit(intNewTarget).BackColor = PaleCyan()
            lblToHit(intNewTarget).Tag = REMOVE_FIGHTER
            gintRemovalsRemaining = gintRemovalsRemaining - 1
      
        End If
        
        ' Display the number of remaining removals.
        
        lblMiscWave.Caption = "Remove " & gintRemovalsRemaining & _
                              " Enemy Fighters"
    
        If gintRemovalsRemaining = 0 Then
            ' There are no more remaining removals. Therefore the user may
            ' commit the removals.
            cmdInterrupt.Enabled = True
        Else
            cmdInterrupt.Enabled = False
        End If
    
    ElseIf cmdInterrupt.Caption = FIRE_GUNS _
    Or cmdInterrupt.Caption = PASSING_FIRE Then
      
        ' Operate on local variables.
        
        intCurrGun = gintCurrGun
        intPrevTarget = gintCurrTarget
        intNewTarget = Index
    
        If intCurrGun = 0 Then
            ' No gun was selected, therefore no target may be selected.
            Exit Sub
        End If

        If lblToHit(intNewTarget).BackColor = PaleCyan() Then
        
            If lblGunAmmo(intCurrGun).BackColor = PaleCyan() Then
            
                ' The target was previously assigned to the gun. By selecting
                ' the target again, the user deselects the target.
                lblToHit(intNewTarget).BackColor = MedDkCyan()
                lblGunAmmo(intCurrGun).Tag = ""
                gintCurrTarget = 0
            
            ElseIf lblGunAmmo(intCurrGun).BackColor = MedDkCyan() Then
            
                ' Somehow the target was previously selected despite the
                ' the current gun being unable to fire at it. Correct the
'                ' mistake, then exit.
                lblToHit(intNewTarget).BackColor = vbButtonFace
'                Exit Sub
                
            ElseIf lblGunAmmo(intCurrGun).BackColor = vbButtonFace Then
            
                ' Somehow the target was previously selected despite the
                ' the current gun being unable to fire at it. Correct the
'                ' mistake, then exit.
                lblToHit(intNewTarget).BackColor = vbButtonFace
'                Exit Sub
            
            End If
        
        ElseIf lblToHit(intNewTarget).BackColor = MedDkCyan() Then
        
            If lblGunAmmo(intCurrGun).BackColor = PaleCyan() Then
            
                ' Assign the target to the gun. Save the fighter and to hit of
                ' the selected target for the current gun.
                lblToHit(intNewTarget).BackColor = PaleCyan()
                lblGunAmmo(intCurrGun).Tag = CStr(intNewTarget) & "/" & lblToHit(intNewTarget).Caption
            
                If intPrevTarget >= 1 Then
                    ' Deselect the gun's previous target.
                    lblToHit(intPrevTarget).BackColor = MedDkCyan()
                End If
            
                ' The new target is now current.
                gintCurrTarget = intNewTarget
            
            ElseIf lblGunAmmo(intCurrGun).BackColor = MedDkCyan() Then
            
                ' A gun must be selected before a target may be assigned to
                ' it. The target should not be dark unless a gun was selected.
'                ' Correct the mistake, then exit.
                lblToHit(intNewTarget).BackColor = vbButtonFace
'                Exit Sub
            
            ElseIf lblGunAmmo(intCurrGun).BackColor = vbButtonFace Then
            
                ' A gun must be selected before a target may be assigned to
                ' it. The target should not be dark unless a gun was selected.
'                ' Correct the mistake, then exit.
                lblToHit(intNewTarget).BackColor = vbButtonFace
'                Exit Sub
            
            End If
        
        ElseIf lblToHit(intNewTarget).BackColor = vbButtonFace Then
        
'            ' User selected fighter which can't be targetted by the current
'            ' gun. If there is a selected weapon, the target is deselected,
'            ' otherwise nothing happens.
'            If lblGunAmmo(intCurrGun).BackColor = PaleCyan() Then
''                lblGunAmmo(intCurrGun).BackColor = MedDkCyan()
'                lblGunAmmo(intCurrGun).Tag = ""
'                gintCurrGun = 0
'                lblToHit(intPrevTarget).BackColor = MedDkCyan()
'                gintCurrTarget = 0
'            End If
            
            If lblGunAmmo(intCurrGun).BackColor = PaleCyan() Then
            
                ' User selected fighter which can't be targetted by the current
                ' gun. If the gun was previously assigned a target, deselect
                ' it.
                
                If intPrevTarget >= 1 Then
                    lblToHit(intPrevTarget).BackColor = MedDkCyan()
                    lblGunAmmo(intCurrGun).Tag = ""
                    gintCurrTarget = 0
                End If
            
            ElseIf lblGunAmmo(intCurrGun).BackColor = MedDkCyan() Then
            
a = 1
            ElseIf lblGunAmmo(intCurrGun).BackColor = vbButtonFace Then
            
a = 1
            
            End If
        
        End If
    
'        gintCurrGun = intGun
'        gintCurrTarget = intNewTarget
    
    End If

a = 1

End Sub

'******************************************************************************
' SprayFireAllowed
'
' INPUT:  The gun that is firing and the fighter it is being fired at.
'
' OUTPUT: True if spray fire is allowed, otherwise false.
'
' RETURN: n/a
'
' NOTES: If spray is not allowed, uncheck chkSpray so normal fire will occur.
'******************************************************************************
Private Function SprayFireAllowed(ByVal intGun As Integer, ByVal intTarget As Integer) As Boolean

    SprayFireAllowed = True
    
    If chkSpray(intGun).Value = vbChecked Then
        ' Rule 9.5
        If InStr(1, lblPosition(intTarget).Caption, "10:30", vbTextCompare) >= 1 _
        Or InStr(1, lblPosition(intTarget).Caption, "12", vbTextCompare) >= 1 _
        Or InStr(1, lblPosition(intTarget).Caption, "1:30", vbTextCompare) >= 1 _
        Or InStr(1, lblPosition(intTarget).Caption, "Vertical Dive", vbTextCompare) >= 1 Then
            chkSpray(intGun).Value = vbUnchecked
            SprayFireAllowed = False
        End If
    End If

End Function

'******************************************************************************
' OffensiveFire
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, or default value.
'
' NOTES:  All surviving German fighters shoot at the bomber. Generate damage
'         due to any hits.
'******************************************************************************
Private Function OffensiveFire(ByVal blnSurpriseAttack As Boolean) As Integer
    Dim blnEvadeFighters As Boolean
    Dim intFighterPos As Integer
    Dim intPilotSkill As Integer
    Dim intDamage As Integer
    Dim strFighterType As String
    Dim intFighter As Integer
    Dim intTotalHits As Integer
    Dim intHit As Integer
    Dim intSide As Integer
    Dim intHitArea As Integer
    Dim arrHitArea() As Integer

    If chkEvadeFighters.Value = vbChecked Then
        blnEvadeFighters = True
    Else
        blnEvadeFighters = False
    End If
    
    For intFighter = 1 To GetWaveSize()
    
        intFighterPos = Wave.Fighter(intFighter).Position
        intPilotSkill = Wave.Fighter(intFighter).PilotSkill
        intDamage = Wave.Fighter(intFighter).Damage
        strFighterType = Wave.Fighter(intFighter).Type
    
        Select Case intFighterPos
            ' Ignore side if the fighter is not attacking from a side.
            Case F12_HIGH To F12_LOW: intSide = 0
            Case F130_HIGH To F130_LOW: intSide = STBD_SIDE
            Case F3_HIGH To F3_LOW: intSide = STBD_SIDE
            Case F6_HIGH To F6_LOW: intSide = 0
            Case F9_HIGH To F9_LOW: intSide = PORT_SIDE
            Case F1030_HIGH To F1030_LOW: intSide = PORT_SIDE
            Case VERT_DIVE: intSide = 0
            Case VERT_CLIMB: intSide = 0
        End Select

        If M3GermanOffensiveFire(intFighterPos, _
                                 intPilotSkill, _
                                 intDamage, _
                                 blnEvadeFighters) = True Then
            
            If blnSurpriseAttack = True Then
                    
                ' Surprise attacks can be very devastating ...

                intTotalHits = Random2D6()

            Else
            
                intTotalHits = B4ShellHitsByArea(intFighterPos, strFighterType)
            
            End If
        
            ' Walking hits negate all other hits from an attack. Queue the
            ' damage location rolls, one by one. If a roll is a walking hit,
            ' delete any previous hits, then skip succeeding hits.
    
            ReDim arrHitArea(1 To intTotalHits)
                
            For intHit = 1 To intTotalHits
            
                intHitArea = B5AreaDamage(intFighterPos, intSide)
                
                If intHitArea = WALKING_HITS_FUSELAGE _
                Or intHitArea = WALKING_HITS_WINGS _
                Or intHitArea = WALKING_HITS_BOTH Then
                
                    ReDim arrHitArea(1)
                    arrHitArea(1) = intHitArea
                    intTotalHits = 1
                    Exit For
                
                Else
                
                    arrHitArea(intHit) = intHitArea
                
                End If

            Next intHit

            UpdateMessage Wave.Fighter(intFighter).Type & " at " & _
                          PositionText(Wave.Fighter(intFighter).Position, "PosNoToString") & _
                          " registers " & intTotalHits & " hits."

            ' Now that the number and location of hits has been finalized,
            ' acually inflict the damage.

            For intHit = 1 To intTotalHits
            
                If B5AreaDamageRouter(arrHitArea(intHit), intSide) = END_MISSION Then
                    OffensiveFire = END_MISSION
                    Exit Function
                End If
                
            Next intHit

        Else
        
            lblToHit(intFighter).Tag = FIGHTER_MISSED
        
            UpdateMessage Wave.Fighter(intFighter).Type & " at " & _
                          PositionText(Wave.Fighter(intFighter).Position, "PosNoToString") & _
                          " missed."

        End If
    
        If AlpsDirection() = ALPS_BELOW Then
            ' German fighters are only allowed to make one pass in Alps.
            lblToHit(intFighter).Tag = REMOVE_FIGHTER
        End If

    Next intFighter
    
End Function

'******************************************************************************
' EvasiveActionAllowed
'
' INPUT:  True if evasive action is allowed, otherwise false.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Function EvasiveActionAllowed() As Boolean
    Dim intEnginesOut As Integer
    Dim intControlsOut As Integer
    
    ' Rule 15.2
    
    EvasiveActionAllowed = True

    intControlsOut = 0
    
    intEnginesOut = CountEnginesOut()
        
    ' Rule 15.2.d states that -3 or worse cumulative landing modifiers
    ' disallows evasive action, but things such the operability of the
    ' landing gear (-3 when out) are irrelevant. Therefore, only consider
    ' the condition of the aircraft's control surfaces.
    
    If Damage.RudderControls = True Then
        intControlsOut = intControlsOut + 1
    Else
        If Bomber.BomberModel = B17_C _
        Or Bomber.BomberModel = B17_E _
        Or Bomber.BomberModel = B17_F _
        Or Bomber.BomberModel = B17_G _
        Or Bomber.BomberModel = YB40 Then
                
            ' A B-17 only has one rudder, so by default it is the 'port side'.
            
            If Damage.Rudder(PORT_SIDE) >= 3 Then
                intControlsOut = intControlsOut + 1
            End If
                
        Else
                
            If Damage.Rudder(PORT_SIDE) >= 2 _
            And Damage.Rudder(STBD_SIDE) >= 2 Then
                intControlsOut = intControlsOut + 1
            End If
    
        End If
    End If
    
    If Damage.WingFlapControls = True _
    Or (Damage.WingFlap(PORT_SIDE) = True _
    And Damage.WingFlap(STBD_SIDE) = True) Then
        intControlsOut = intControlsOut + 1
    End If

    If Damage.AileronControls = True _
    Or (Damage.Aileron(PORT_SIDE) = True _
    And Damage.Aileron(STBD_SIDE) = True) Then
        intControlsOut = intControlsOut + 1
    End If

    If Damage.ElevatorControls = True _
    Or (Damage.Elevator(PORT_SIDE) = True _
    And Damage.Elevator(STBD_SIDE) = True) Then
        intControlsOut = intControlsOut + 1
    End If

    ' Bomber must be out of formation, must have 3 or 4 operational engines,
    ' must have operational control cables, must have two or more operational
    ' control surfaces, and either the pilot or copilot must be at the controls,
    ' otherwise evasion is impossible.

    If Bomber.InFormation = True _
    Or Damage.BurstInPlane = True _
    Or intEnginesOut >= 2 _
    Or Damage.ControlCables >= 2 _
    Or intControlsOut >= 3 Then
        
        EvasiveActionAllowed = False
    
    ElseIf Bomber.Airman(PILOT).Status >= SW_STATUS _
    And Bomber.Airman(COPILOT).Status >= SW_STATUS Then
        
        EvasiveActionAllowed = False
    
    End If
    
End Function

'******************************************************************************
' DefensiveFire
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Fire the bomber's guns at all selected targets.
'******************************************************************************
Private Sub DefensiveFire()
    Dim intGun As Integer
    Dim intTarget As Integer
    Dim intToHit As Integer
    Dim blnFW190 As Boolean
    Dim intToHitRoll As Integer
    Dim intNewDamage As Integer
    Dim intOldDamage As Integer
    Dim strDamage As String
    Dim intOrigPos As Integer
    Dim intSpray As Integer
    
    For intGun = MID_UPPER_MG To TAIL_MG
        
        If GunManned(intGun) = True _
        And Len(Mid(lblGunAmmo(intGun).Tag, 3, 1)) = 1 Then
            
            ' Gun has been assigned a target. Left of the slash is the
            ' fighter; right of the slash is the to hit value.

            intTarget = CInt(Mid(lblGunAmmo(intGun).Tag, 1, 1))
            intToHit = CInt(Mid(lblGunAmmo(intGun).Tag, 3, 1))
            
            intOldDamage = Wave.Fighter(intTarget).Damage
                
            ' Get original position of Airman that is currently manning the
            ' weapon.
            
            intOrigPos = GetAirmanIndexBySerialNumber(Bomber.Gun(intGun).MannedBy)
                
            If chkSpray(intGun).Value = vbChecked _
            And SprayFireAllowed(intGun, intTarget) = True Then
            
                ' Rule 9.5: Gun is being spray-fired. (High volume, low
                ' accuracy fire.)
                
                Bomber.Gun(intGun).Ammo = Bomber.Gun(intGun).Ammo - 3
                lblGunAmmo(intGun).Caption = CStr(Bomber.Gun(intGun).Ammo)
        
                intSpray = M5SprayAreaFire()
            
                Select Case intSpray
                    Case SPRAY_FIRE_JAM:
                        ' Gun jams. Fighter is unaffected.
                        
'                        UpdateMessage "Spray fire jammed " & lblGunName(intGun).Caption & "."
                        UpdateMessage lblGunName(intGun).Caption & " jammed by spray fire."
                        Bomber.Gun(intGun).Status = MG_JAMMED
                        GoTo CheckNextGun
                    
                    Case SPRAY_FIRE_NOEFFECT:
                        ' Fighter is unaffected.
                        
                        UpdateMessage lblGunName(intGun).Caption & " missed spray fire."
                        GoTo CheckNextGun
                    
                    Case SPRAY_FIRE_BREAKOFF:
                        ' Fighter breaks off before attacking. Other defensive
                        ' fires may still affect the fighter.
                        
                        UpdateMessage lblGunName(intGun).Caption & " spray fire drove off " & _
                                      lblType(intTarget).Caption & " at " & _
                                      lblPosition(intTarget).Caption
                        lblToHit(intTarget).Tag = REMOVE_FIGHTER
                        GoTo CheckNextGun
                    
                    Case SPRAY_FIRE_HIT:
                        ' Fall through to normal damage resolution.
                        
                        intToHitRoll = 6
                
                End Select
            
            Else
            
                ' Normal to hit determination.
            
                Bomber.Gun(intGun).Ammo = Bomber.Gun(intGun).Ammo - 1
                lblGunAmmo(intGun).Caption = CStr(Bomber.Gun(intGun).Ammo)
        
                intToHitRoll = Random1D6()
                
                If Bomber.Airman(intOrigPos).Status = LW2_STATUS _
                Or Bomber.Gun(intGun).QualifiedGunner = False _
                Or chkEvadeFighters.Value = vbChecked Then
                    ' BL-4 Wounds: Note a.
                    ' Rule 14.2.b
                    ' Rule 15.1.b
                    intToHit = 6
                Else
                    ' Duty or LW1 status.
                
                    If Bomber.Airman(intOrigPos).Kills >= 5 _
                    And Bomber.Airman(intOrigPos).Status = DUTY_STATUS _
                    And Bomber.Airman(intOrigPos).Frostbite = False Then
                        
                        ' Rule 9.3
                        intToHitRoll = intToHitRoll + 1
                    
                    ElseIf Bomber.Airman(intOrigPos).Mission <= 5 _
                    And Mission.Options.CrewExperience = True Then
                        
                        ' Crew experience variant from the "Theater Modifications"
                        ' article in "The General" (Volume 24, #6).
                        
                        If intToHitRoll = 1 Then
                            
                            ' Unmodified 1 jams weapon.
                            Bomber.Gun(intGun).Status = MG_JAMMED
                            UpdateMessage lblGunName(intGun).Caption & " jammed."
                            GoTo CheckNextGun
                        
                        Else
                            
                            ' Inexperienced gunner less likely to get a hit.
                            intToHitRoll = intToHitRoll - 1
                        
                        End If
                    
                    End If
                
                End If
            
            End If
            
            lblGunAmmo(intGun).BackColor = vbButtonFace
            
            If (intToHitRoll >= intToHit _
            Or intToHitRoll = 6) _
            And intOldDamage < SHOT_DOWN_DAMAGE Then
            
                If lblType(intTarget).Caption = "FW190" Then
                    ' M-2 Hit Damage: Note b.
                    blnFW190 = True
                Else
                    blnFW190 = False
                End If
                
                intNewDamage = M2HitDamageAgainstGermanFighter(Bomber.Gun(intGun).Bonus, blnFW190)
                
                If lblType(intTarget).Caption = "Ju88" _
                And Bomber.Gun(intGun).Bonus = 0 _
                And Bomber.Airman(intOrigPos).Kills <= 4 Then
                    
                    ' Single-mount .50, and dual-mount .303, hits on the Ju88
                    ' fired by a non-ace gunner are reduced one level.
                    
                    Select Case intNewDamage
                        Case FCA_DAMAGE: intNewDamage = NO_DAMAGE
                        Case FBOA_DAMAGE: intNewDamage = FCA_DAMAGE
                        Case SHOT_DOWN_DAMAGE: intNewDamage = FBOA_DAMAGE
                    End Select
                
                End If
                
'UpdateMessage vbCrLf & "Bef: " & lblType(intTarget).Caption & " at " & lblPosition(intTarget).Caption & "Dam = " & Wave.Fighter(intTarget).Damage ' DEBUG
'UpdateMessage "Damage = [" & intNewDamage & "]" ' DEBUG
                
                ' Increment the fighter's damage.

                If intOldDamage >= SHOT_DOWN_DAMAGE Then
                    
                    ' Do nothing. The fighter was shot down by another gun.
                
                ElseIf intOldDamage >= FBOA_DAMAGE Then
                    
                    If intNewDamage = SHOT_DOWN_DAMAGE Then
                        
                        Wave.Fighter(intTarget).Damage = SHOT_DOWN_DAMAGE
                        strDamage = "shot down"
                        Bomber.Airman(intOrigPos).Kills = Bomber.Airman(intOrigPos).Kills + 1
                    
                    ElseIf intNewDamage = FBOA_DAMAGE Then
                        
                        Wave.Fighter(intTarget).Damage = SHOT_DOWN_DAMAGE
                        strDamage = "shot down"
                        Bomber.Airman(intOrigPos).Kills = Bomber.Airman(intOrigPos).Kills + 1
                    
                    ElseIf intNewDamage = FCA_DAMAGE Then
                        
                        If intOldDamage + intNewDamage = SHOT_DOWN_DAMAGE Then
                            
                            Wave.Fighter(intTarget).Damage = SHOT_DOWN_DAMAGE
                            strDamage = "shot down"
                            Bomber.Airman(intOrigPos).Kills = Bomber.Airman(intOrigPos).Kills + 1
                        
                        Else
                            
                            Wave.Fighter(intTarget).Damage = FBOA_DAMAGE + FCA_DAMAGE
                            strDamage = "FBOA"
                        
                        End If
                    
                    End If
                
                ElseIf intOldDamage = FCA_DAMAGE Then
                    
                    If intNewDamage = SHOT_DOWN_DAMAGE Then
                        
                        Wave.Fighter(intTarget).Damage = SHOT_DOWN_DAMAGE
                        strDamage = "shot down"
                        Bomber.Airman(intOrigPos).Kills = Bomber.Airman(intOrigPos).Kills + 1
                    
                    ElseIf intNewDamage = FBOA_DAMAGE Then
                        
                        If intOldDamage + intNewDamage = SHOT_DOWN_DAMAGE Then
                            
                            Wave.Fighter(intTarget).Damage = SHOT_DOWN_DAMAGE
                            strDamage = "shot down"
                            Bomber.Airman(intOrigPos).Kills = Bomber.Airman(intOrigPos).Kills + 1
                        
                        Else
                            
                            Wave.Fighter(intTarget).Damage = FBOA_DAMAGE + FCA_DAMAGE
                            strDamage = "FBOA"
                        
                        End If
                    
                    ElseIf intNewDamage = FCA_DAMAGE Then
                        
                        Wave.Fighter(intTarget).Damage = FBOA_DAMAGE
                        strDamage = "FBOA"
                    
                    End If
                
                ElseIf intOldDamage = NO_DAMAGE Then
                    
                    Wave.Fighter(intTarget).Damage = intNewDamage
                
                    Select Case Wave.Fighter(intTarget).Damage
                        Case SHOT_DOWN_DAMAGE:
                            strDamage = "shot down"
                            Bomber.Airman(intOrigPos).Kills = Bomber.Airman(intOrigPos).Kills + 1
                        Case FBOA_DAMAGE: strDamage = "FBOA"
                        Case FCA_DAMAGE: strDamage = "FCA"
                        Case NO_DAMAGE: strDamage = "no damage"
                    End Select
                
                End If
                
'UpdateMessage "Fighter Damage (Aft) = [" & Wave.Fighter(intTarget).Damage & "]" ' DEBUG
'UpdateMessage "Aft: " & lblType(intTarget).Caption & " at " & lblPosition(intTarget).Caption & "Dam = " & Wave.Fighter(intTarget).Damage ' DEBUG
                
                ' Assemble & print a damage message. The message shows the
                ' cumulative damage, rather than the effect of this particular
                ' hit. Thus the airman may get credited for the 'kill shot'
                ' even if he only got a FBOA result.
                
                If intSpray = SPRAY_FIRE_HIT Then
                    UpdateMessage lblGunName(intGun).Caption & " spray fire hit " & _
                                  lblType(intTarget).Caption & " at " & _
                                  lblPosition(intTarget).Caption & " - " & _
                                  strDamage
                Else
                    UpdateMessage lblGunName(intGun).Caption & " hit " & _
                                  lblType(intTarget).Caption & " at " & _
                                  lblPosition(intTarget).Caption & " - " & _
                                  strDamage
                End If

            Else
            
                UpdateMessage lblGunName(intGun).Caption & " missed " & _
                              lblType(intTarget).Caption & " at " & _
                              lblPosition(intTarget).Caption

            End If
        
        End If
        
CheckNextGun:
    
    Next intGun

End Sub

'******************************************************************************
' MayFireGun
'
' INPUT:  Key to a particular gun.
'
' OUTPUT: n/a
'
' RETURN: True if gun may be fired, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Private Function MayFireGun(ByVal intGun As Integer) As Boolean
    Dim intIndex As Integer
    
    MayFireGun = True
        
    If Bomber.Gun(intGun).Status <> MG_OKAY _
    Or Bomber.Gun(intGun).Ammo = 0 _
    Or GunExists(intGun) = False Then

        ' User tried to fire a MG which is jammed, inoperable, out of ammo,
        ' or which does not even exist on the bomber.
        
        MayFireGun = False
        Exit Function

    End If

    If GunManned(intGun) = True Then
            
        ' Airman currently manning gun
        intIndex = GetAirmanIndexBySerialNumber(Bomber.Gun(intGun).MannedBy)

        If Bomber.Airman(intIndex).Status >= LW2_STATUS Then
            MayFireGun = False
        End If
    
    Else 'If GunManned(intGun) = False Then
    
        ' Nobody is currently manning the gun, but another airman may be able
        ' to quickly swap into the position.

        If intGun = PORT_CHEEK_MG _
        Or intGun = STBD_CHEEK_MG Then
            
            If PosManned(NAVIGATOR) = True Then
                
                intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(NAVIGATOR).CurrentSerialNum)
            
                If Bomber.Airman(intIndex).Status >= LW2_STATUS Then
                    ' The airman in the position is too wounded to make a
                    ' quick pos swap.
                    MayFireGun = False
                End If
            
            Else
                
                MayFireGun = False
            
            End If
        
        ElseIf intGun = PORT_WAIST_MG _
        And (Bomber.BomberModel = B24_D _
        Or Bomber.BomberModel = B24_E _
        Or Bomber.BomberModel = B24_GHJ _
        Or Bomber.BomberModel = B24_LM) Then

            ' Port waist is unmanned, so we know the stbd waist and tunnel
            ' guns must be manned.

            If Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum)).Status >= LW2_STATUS _
            And Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(BALL_GUNNER).CurrentSerialNum)).Status >= LW2_STATUS Then
                ' The airman in the position is too wounded to make a
                ' quick pos swap.
                MayFireGun = False
            End If
        
        ElseIf intGun = STBD_WAIST_MG _
        And (Bomber.BomberModel = B24_D _
        Or Bomber.BomberModel = B24_E _
        Or Bomber.BomberModel = B24_GHJ _
        Or Bomber.BomberModel = B24_LM) Then

            ' Stbd waist is unmanned, so we know the port waist and tunnel
            ' guns must be manned.

            If Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum)).Status >= LW2_STATUS _
            And Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(BALL_GUNNER).CurrentSerialNum)).Status >= LW2_STATUS Then
                ' The airman in the position is too wounded to make a
                ' quick pos swap.
                MayFireGun = False
            End If
        
        ElseIf (intGun = BALL_TURRET_MG _
        And (Bomber.BomberModel = B24_D _
        Or Bomber.BomberModel = B24_E)) Then
        
            ' Tunnel is unmanned, so we know the stbd waist and port waist
            ' guns must be manned.

            If Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum)).Status >= LW2_STATUS _
            And Bomber.Airman(GetAirmanIndexBySerialNumber(Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum)).Status >= LW2_STATUS Then
                ' The airman in the position is too wounded to make a
                ' quick pos swap.
                MayFireGun = False
            End If
        
        End If
    
    End If
    
End Function

'******************************************************************************
' PreAssignGuns
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Basically, color the gun if it has a target in its field of fire.
'******************************************************************************
Private Sub PreAssignGuns()

    Dim intGun As Integer
    Dim intFighter As Integer
    
    For intGun = MID_UPPER_MG To TAIL_MG
        
        If MayFireGun(intGun) = True Then
            
            ' The gun may fire. See if there are any fighters within its
            ' line-of-sight.
            
            For intFighter = 1 To GetWaveSize()
                
                If HasATarget(Bomber.BomberModel, _
                              intGun, _
                              Wave.Fighter(intFighter).Position) = True Then
                
                    lblGunAmmo(intGun).BackColor = MedDkCyan()
                    
                    If Bomber.Gun(intGun).Ammo >= 3 Then
                        chkSpray(intGun).Enabled = True
                    Else
                        chkSpray(intGun).Enabled = False
                    End If
                    
                    Exit For
                
                Else
                    lblGunAmmo(intGun).BackColor = vbButtonFace
                    chkSpray(intGun).Enabled = False
                End If
            
            Next intFighter
        
        Else
            lblGunAmmo(intGun).BackColor = vbButtonFace
            chkSpray(intGun).Enabled = False
        End If
    
        ' Regardless of whether the weapon has a target or not.
    
        lblGunAmmo(intGun).Tag = ""
        chkSpray(intGun).Value = vbUnchecked
    
    Next intGun

End Sub

'******************************************************************************
' HasATarget
'
' INPUT:  Key to a particular gun and fighter.
'
' OUTPUT: n/a
'
' RETURN: True if the fighter is in the gun's field of fire.
'
' NOTES:  n/a
'******************************************************************************
Private Function HasATarget(ByVal intBomberModel As Integer, ByVal intGun As Integer, ByVal intPosition As Integer) As Boolean
    Dim rsGunnery As New ADODB.Recordset
    Dim strErrMsg As String

    HasATarget = False
    
    pobjCmnd.CommandText = " SELECT * FROM Gunnery" & _
                           " WHERE BomberModel = " & intBomberModel & _
                           " AND GunPos = " & intGun & _
                           " AND FighterPos = " & intPosition

    rsGunnery.CursorLocation = adUseClient
    rsGunnery.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    
    If RecordsInSet(rsGunnery) = 0 Then
        HasATarget = False
    Else
        HasATarget = True
    End If
    
CleanUp:
   
    If Not rsGunnery Is Nothing Then
        If rsGunnery.State = adStateClosed Then rsGunnery.Close
        Set rsGunnery = Nothing
    End If
   
    Exit Function
   
ErrorTrap:
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "HasATarget() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    HasATarget = False
    
    Resume CleanUp

End Function

'******************************************************************************
' GetMaxRemovals
'
' INPUT:  Number of removals due to fighter cover.
'
' OUTPUT: n/a
'
' RETURN: Number of Germans that may actually be removed.
'
' NOTES:  n/a
'******************************************************************************
Private Function GetMaxRemovals(ByVal intRemovals As Integer) As Integer
    
    Dim intMaxRemovals As Integer
    Dim intCanBeRemoved As Integer
    Dim intFighterCount As Integer
    Dim intFighters As Integer
    
    intMaxRemovals = 0
    intCanBeRemoved = 0
    intFighterCount = 0
    
    intFighters = GetWaveSize()
    
'UpdateMessage "GetWaveSize() = [" & GetWaveSize() & "]" ' DEBUG
    
    ' Rule 18.0.d.
    
    If RandomEvent.BadLuftwaffeComm = True Then
        intRemovals = intRemovals + 1
    End If
                
    If intRemovals > intFighters Then
        ' Can't remove any more fighters than there are in the wave.
        intMaxRemovals = intFighters
    Else
        intMaxRemovals = intRemovals
    End If

'UpdateMessage "intMaxRemovals = [" & intMaxRemovals & "]" ' DEBUG

    For intFighterCount = 1 To intFighters
        
        If Wave.Fighter(intFighterCount).Special = "B" _
        Or Wave.Fighter(intFighterCount).Special = "D" Then
            
            ' Friendlies can't drive off this fighter.
            
            lblToHit(intFighterCount).BackColor = vbRed
            
        Else
            
            intCanBeRemoved = intCanBeRemoved + 1
        
        End If
        
    Next intFighterCount
        
'UpdateMessage "intCanBeRemoved = [" & intCanBeRemoved & "]" ' DEBUG

    If intMaxRemovals > intCanBeRemoved Then
        intMaxRemovals = intCanBeRemoved
    End If
    
    GetMaxRemovals = intMaxRemovals

'UpdateMessage "GetMaxRemovals = [" & GetMaxRemovals & "]" ' DEBUG

End Function

'******************************************************************************
' FightersInWave
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Number of functional fighters in the wave.
'
' NOTES:  A fighter may temporarily remain in the wave, after it has been shot
'         down or missed its own shot, prior to being cleaned up after the
'         wave's air-to-air combat is complete. Count fighters which are still
'         capable of attacking or being shot at.
'******************************************************************************
Private Function FightersInWave() As Integer
    Dim intFighterCount As Integer

    intFighterCount = 0
    FightersInWave = 0

    ' For each fighter that was originally part of the wave ...
        
    For intFighterCount = 1 To GetWaveSize()
        
        ' If the fighter was shot down or severely damaged during a previous
        ' attack, then it is no longer part of the current wave.
            
        If Wave.Fighter(intFighterCount).Damage <> FBOA_DAMAGE _
        And Wave.Fighter(intFighterCount).Damage <> SHOT_DOWN_DAMAGE Then
            
            ' Fighter has light or no damage, so it may make another attack.
            
            FightersInWave = FightersInWave + 1
        
        End If
    
    Next intFighterCount

End Function

'******************************************************************************
' RemoveFightersFromWave
'
' INPUT:  Flag to drop fighters with FBOA, or worse, damage. Flag to reposition
'         fighters.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Pop the fighter from the wave structure if it missed its shot, or
'         was heavily damaged or shot down.
'******************************************************************************
Private Sub RemoveFightersFromWave(ByVal blnDropFBOA As Boolean, ByVal blnReposition As Boolean)
    Dim intFighterCount As Integer
    Dim WaveTemp As WaveInfoNew
    Dim intKeep As Integer
    
    intKeep = 0
    
    WaveTemp.Attack = Wave.Attack
    WaveTemp.JG26 = Wave.JG26
    WaveTemp.Ju88 = Wave.Ju88

'UpdateMessage "cmdInterrupt.Caption = '" & cmdInterrupt.Caption & "'" ' DEBUG
    
    ' For each fighter that is currently part of the wave ...
        
    For intFighterCount = 1 To GetWaveSize()
    
        If lblToHit(intFighterCount).Tag = REMOVE_FIGHTER _
        Or Wave.Fighter(intFighterCount).Damage >= SHOT_DOWN_DAMAGE _
        Or (Wave.Fighter(intFighterCount).Damage >= FBOA_DAMAGE _
        And blnDropFBOA = True) _
        Or (lblToHit(intFighterCount).Tag = FIGHTER_MISSED _
        And Bomber.InFormation = True) Then
        
            ' If German fighter was marked for removal, was shot down or too
            ' damaged to continue, or missed its attack, it should be removed
            ' from the wave. The "removal" is accomplished by not copying the
            ' fighter to WaveTemp.

'            UpdateMessage vbCrLf & "----------------------------------------" ' DEBUG
'            UpdateMessage intFighterCount & " Drop: " & _
                          "Type = '" & Wave.Fighter(intFighterCount).Type & "', " & _
                          "Pos = " & Wave.Fighter(intFighterCount).Position & ", " & _
                          "Dam = " & Wave.Fighter(intFighterCount).Damage
'            UpdateMessage "----------------------------------------"
    
        Else
        
            ' This fighter may make another attack. Copy it to WaveTemp.
            
            intKeep = intKeep + 1
        
            ReDim Preserve WaveTemp.Fighter(1 To intKeep)

            WaveTemp.Fighter(intKeep) = Wave.Fighter(intFighterCount)
        
'            UpdateMessage vbCrLf & "----------------------------------------" ' DEBUG
'            UpdateMessage intFighterCount & " Keep: " & _
                          "Type = '" & WaveTemp.Fighter(intKeep).Type & "', " & _
                          "Pos = " & WaveTemp.Fighter(intKeep).Position & ", " & _
                          "Spec = " & WaveTemp.Fighter(intKeep).Special & ", " & _
                          "Dam = " & WaveTemp.Fighter(intKeep).Damage
'            UpdateMessage "----------------------------------------"
        
            If blnReposition = True _
            And Wave.Attack <= 3 Then
            
                WaveTemp.Fighter(intKeep).Position = B6SuccessiveAttacks()
            
                If WaveTemp.Fighter(intKeep).Special = "B" _
                Or WaveTemp.Fighter(intKeep).Special = "D" _
                Or WaveTemp.Fighter(intKeep).Special = "E" Then
                    WaveTemp.Fighter(intKeep).Special = ""
                End If
            
            End If
        
        End If
    
    Next intFighterCount

    For intFighterCount = 1 To intKeep
        
'        UpdateMessage vbCrLf & "----------------------------------------" ' DEBUG
'        UpdateMessage intFighterCount & " New: " & _
                      "Type = '" & WaveTemp.Fighter(intFighterCount).Type & "', " & _
                      "Pos = " & WaveTemp.Fighter(intFighterCount).Position & ", " & _
                      "Spec = " & WaveTemp.Fighter(intFighterCount).Special & ", " & _
                      "Dam = " & WaveTemp.Fighter(intFighterCount).Damage
'        UpdateMessage "----------------------------------------"
        
    Next intFighterCount
    
'    UpdateMessage vbCrLf ' DEBUG

    Wave = WaveTemp

End Sub

'******************************************************************************
' DisplayWave
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Display all the fighters in the wave that are capable of further
'         attacks.
'******************************************************************************
Private Sub DisplayWave()

    Dim intFighterCount As Integer
    Dim strPosition As String
    
    intFighterCount = 0
    strPosition = ""
    
    Call ClearWaveDisplay(True) ' iop
    
    ' Only display data for fighters in the current wave.
    
    For intFighterCount = 1 To GetWaveSize()
    
        If Wave.Fighter(intFighterCount).Damage < SHOT_DOWN_DAMAGE Then
            
            lblToHit(intFighterCount).BackColor = vbButtonFace 'yoyoyo
            
            lblType(intFighterCount).Caption = Wave.Fighter(intFighterCount).Type
            
            If Wave.Fighter(intFighterCount).Damage = FCA_DAMAGE Then
                lblType(intFighterCount).BackColor = PaleYellow()
            ElseIf Wave.Fighter(intFighterCount).Damage >= FBOA_DAMAGE Then
                lblType(intFighterCount).BackColor = PaleRed()
            End If
            
            strPosition = PositionText(Wave.Fighter(intFighterCount).Position, "PosNoToString")
            
            Select Case Wave.Fighter(intFighterCount).Special
                Case "B":
                    strPosition = strPosition & " (B)"
                Case "D":
                    strPosition = strPosition & " (D)"
                Case "E":
                    strPosition = strPosition & " (E)"
            End Select
            
            lblPosition(intFighterCount).Caption = strPosition
    
        End If
    
    Next intFighterCount

End Sub

'******************************************************************************
'
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:
'******************************************************************************
Private Function PositionText(intPosNo As Integer, strConvert As String) As String

   If strConvert = "PosNoToString" Then
      Select Case intPosNo
         Case F12_HIGH: PositionText = "12 High"
         Case F12_LEVEL: PositionText = "12 Level"
         Case F12_LOW: PositionText = "12 Low"
         Case F130_HIGH: PositionText = "1:30 High"
         Case F130_LEVEL: PositionText = "1:30 Level"
         Case F130_LOW: PositionText = "1:30 Low"
         Case F3_HIGH: PositionText = "3 High"
         Case F3_LEVEL: PositionText = "3 Level"
         Case F3_LOW: PositionText = "3 Low"
         Case F6_HIGH: PositionText = "6 High"
         Case F6_LEVEL: PositionText = "6 Level"
         Case F6_LOW: PositionText = "6 Low"
         Case F9_HIGH: PositionText = "9 High"
         Case F9_LEVEL: PositionText = "9 Level"
         Case F9_LOW: PositionText = "9 Low"
         Case F1030_HIGH: PositionText = "10:30 High"
         Case F1030_LEVEL: PositionText = "10:30 Level"
         Case F1030_LOW: PositionText = "10:30 Low"
         Case VERT_CLIMB: PositionText = "Vertical Climb"
         Case VERT_DIVE: PositionText = "Vertical Dive"
      End Select
   End If
   
   If strConvert = "GunNoToString" Then
      Select Case intPosNo
         Case MID_UPPER_MG: PositionText = "Mid-Upper Gun"
         Case NOSE_MG: PositionText = "Nose Gun"
         Case PORT_CHEEK_MG: PositionText = "Port Cheek Gun"
         Case STBD_CHEEK_MG: PositionText = "Starboard Cheek Gun"
         Case RADIO_ROOM_MG: PositionText = "Radio Gun"
         Case PORT_WAIST_MG: PositionText = "Port Waist Gun"
         Case STBD_WAIST_MG: PositionText = "Starboard Waist Gun"
         Case TOP_TURRET_MG: PositionText = "Top Turret"
         Case BALL_TURRET_MG: PositionText = "Ball Turret"
         Case TAIL_MG: PositionText = "Tail Guns"
      End Select
   End If

End Function

'******************************************************************************
' Interrupt
'
' INPUT:  The button for which a click is pending and the caption to be
'         displayed on that button.
'
' OUTPUT: None (The button is passed ByRef so we have access to its properties.)
'
' RETURN: n/a
'
' NOTES:  The mission engine keeps chugging along until some user input is
'         required. Calling this function pauses the engine while the input is
'         performed, then allows the engine to continue after the button is
'         clicked. Check every fraction of a second to see if the button has
'         been clicked, otherwise sleep.
'
' ALSO!!! need to check if the window has been closed. If it has, the entire
'         program needs to shut down. Otherwise the program looks like it is
'         "closed", when it actually is still a running process.
'
'******************************************************************************
Private Function Interrupt(ByRef tmpButton As CommandButton, ByVal strCaption As String) As Integer
    Static blnPause As Boolean

    ' Determine whether or not the program is currently interrupted, waiting
    ' for user input.
    
    If blnPause = True Then
      
        ' Cancel the interrupt. Continue processing.
        
        tmpButton.Caption = strCaption
        
        blnPause = False
   
    Else
      
        blnPause = True
        
        ' Display text indicating the input that is required before processing
        ' should continue.
        
        tmpButton.Caption = strCaption
        
        ' Loop forever. (Unless an event is generated which can exit the loop,
        ' such as clicking a button or shutting down the app.)

        Do While blnPause
         
            ' Wait for approximately 1/10th of a second. Note that this pause
            ' consumes very few CPU cycles.
                 
            Sleep 100
            
            ' Check for user input. (DoEvents interrupts the outter forever
            ' loop.)
    
            DoEvents
            
            ' If no interrupt event occurred, loop again.
        
        Loop

        blnPause = False

   End If

End Function

'******************************************************************************
' Interrupt
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Wrapper to the Interrupt() function. When the button is clicked, it
'         tells Interrupt() that the user is done with his input, so that the
'         program may resume normal processing.
'******************************************************************************
Private Sub cmdInterrupt_Click()

    Call Interrupt(cmdInterrupt, cmdInterrupt.Caption)
    
End Sub



