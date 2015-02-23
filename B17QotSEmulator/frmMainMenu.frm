VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   """B-17 Queen of the Skies"" Emulator"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9765
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   240
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4275
      TabIndex        =   180
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2835
      TabIndex        =   179
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1395
      TabIndex        =   178
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Frame fraExitHelp 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2835
      TabIndex        =   181
      Top             =   7680
      Width           =   5535
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   200
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdFlyMission 
         Caption         =   "Fly Mission"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   199
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   183
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   182
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame fraTab 
      Height          =   6840
      Index           =   4
      Left            =   120
      TabIndex        =   146
      Top             =   512
      Visible         =   0   'False
      Width           =   9515
      Begin VB.ComboBox cboBomberModel 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   150
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox cboName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   148
         Top             =   480
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   3840
         TabIndex        =   184
         Top             =   120
         Width           =   5535
         Begin VB.TextBox txtLogSpeed 
            Height          =   285
            Left            =   120
            TabIndex        =   207
            Text            =   "2"
            Top             =   5400
            Width           =   375
         End
         Begin VB.CheckBox chkRedTailAngels 
            Caption         =   "Red Tail Angels"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   201
            Top             =   5040
            Width           =   1935
         End
         Begin VB.CheckBox chkUnescorted 
            Caption         =   "Unescorted"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   197
            Top             =   4680
            Width           =   1575
         End
         Begin VB.CheckBox chkExtraAmmoInBombBay 
            Caption         =   "Extra Ammo in Bomb Bay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   120
            TabIndex        =   194
            ToolTipText     =   "from ""The General"", vol. 26 no. 5"
            Top             =   4320
            Width           =   2655
         End
         Begin VB.CheckBox chkAlternateWeather 
            Caption         =   "Alternate Weather"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   120
            TabIndex        =   174
            ToolTipText     =   "from ""The General"", vol. 24 no. 6"
            Top             =   2520
            Width           =   2175
         End
         Begin VB.CheckBox chkCrewExperience 
            Caption         =   "Crew Experience"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   120
            TabIndex        =   173
            ToolTipText     =   "from ""The General"", vol. 24 no. 6"
            Top             =   2880
            Width           =   3135
         End
         Begin VB.CheckBox chkFormationDefensiveGunnery 
            Caption         =   "Formation Defensive Gunnery"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   120
            TabIndex        =   171
            ToolTipText     =   "from ""The General"", vol. 24 no. 6"
            Top             =   1800
            Width           =   3255
         End
         Begin VB.CheckBox chkEvadeFlak 
            Caption         =   "Evade Flak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   120
            TabIndex        =   172
            ToolTipText     =   "from ""The General"", vol. 24 no. 6"
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CheckBox chkJG26StationedInAbbeville 
            Caption         =   "JG-26 Stationed in Abbeville"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   176
            Top             =   3600
            Width           =   3135
         End
         Begin VB.CheckBox chkJu88sUsedAsFighters 
            Caption         =   "Ju-88s Used as Fighters"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   177
            Top             =   3960
            Width           =   2655
         End
         Begin VB.CheckBox chkExpandedTargetList 
            Caption         =   "Expanded Target List"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   120
            TabIndex        =   167
            ToolTipText     =   "from ""The General"", vol. 25 no. 5"
            Top             =   360
            Width           =   2415
         End
         Begin VB.CheckBox chkRandomEvents 
            Caption         =   "Random Events"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   120
            TabIndex        =   168
            Top             =   720
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkGermanFighterPilotSkill 
            Caption         =   "German Fighter Pilot Skill"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   120
            TabIndex        =   175
            Top             =   3240
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.CheckBox chkTimePeriodSpecificFormations 
            Caption         =   "Time Period Specific Formations"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   120
            TabIndex        =   170
            ToolTipText     =   "from ""The General"", vol. 24 no. 6"
            Top             =   1440
            Width           =   3255
         End
         Begin VB.CheckBox chkMechanicalFailures 
            Caption         =   "Mechanical Failures"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   120
            TabIndex        =   169
            ToolTipText     =   "from The General, Vol 24. no. 6"
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "Log Speed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   208
            Top             =   5400
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboTarget 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmMainMenu.frx":0ECA
         Left            =   120
         List            =   "frmMainMenu.frx":0ECC
         Style           =   2  'Dropdown List
         TabIndex        =   155
         Top             =   4080
         Width           =   2415
      End
      Begin VB.ComboBox cboMonth 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   159
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdRandomDate 
         Caption         =   "Random"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   161
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdRandomTarget 
         Caption         =   "Random"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   156
         Top             =   4080
         Width           =   975
      End
      Begin VB.ComboBox cboYear 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   160
         Top             =   5040
         Width           =   975
      End
      Begin VB.ComboBox cboFormationPos 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   165
         Top             =   6000
         Width           =   1215
      End
      Begin VB.ComboBox cboSquadronPos 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   164
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton cmdRandomPosition 
         Caption         =   "Random"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   166
         Top             =   6000
         Width           =   975
      End
      Begin VB.Frame fraBase 
         Caption         =   "Base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   4
         Left            =   120
         TabIndex        =   151
         Top             =   2040
         Width           =   1455
         Begin VB.TextBox txtBase 
            Height          =   285
            Left            =   960
            TabIndex        =   198
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.OptionButton optEngland 
            Caption         =   "England"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   152
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optItaly 
            Caption         =   "Italy"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   153
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.Label lblPlane 
         Caption         =   "Bomber"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   147
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblTarget 
         Caption         =   "Target"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   154
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblMonth 
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   157
         Top             =   4680
         Width           =   735
      End
      Begin VB.Label lblYear 
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   158
         Top             =   4680
         Width           =   735
      End
      Begin VB.Label lblBomberPos 
         Caption         =   "Formation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   163
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label lblSquadronPos 
         Caption         =   "Squadron"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   162
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label lblBomberModel 
         Caption         =   "Bomber Model"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   149
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Frame fraTab 
      Height          =   6840
      Index           =   3
      Left            =   120
      TabIndex        =   108
      Top             =   512
      Visible         =   0   'False
      Width           =   9515
      Begin VB.TextBox txtKeyField 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   114
         TabStop         =   0   'False
         Text            =   "txtKeyField(3)"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CheckBox chkDefault 
         Caption         =   "Default Airman"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   121
         Top             =   5880
         Width           =   1455
      End
      Begin VB.ComboBox cboAssignment 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   3360
         Width           =   2415
      End
      Begin VB.ComboBox cboCrewPosition 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   118
         Top             =   4320
         Width           =   2415
      End
      Begin VB.ComboBox cboRank 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Frame fraHistory 
         Caption         =   "Personnel File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Index           =   3
         Left            =   2880
         TabIndex        =   122
         Top             =   120
         Width           =   4575
         Begin VB.TextBox txtMeritoriousUnitCitation 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   5520
            Width           =   735
         End
         Begin VB.TextBox txtDistinguishedUnitCitation 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   5040
            Width           =   735
         End
         Begin VB.TextBox txtAirMedal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   4560
            Width           =   735
         End
         Begin VB.TextBox txtSorties 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtKills 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   125
            TabStop         =   0   'False
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtMedalOfHonor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtDistinguishedServiceCross 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtSilverStar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtDistinguishedFlyingCross 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox txtBronzeStarV 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox txtPurpleHeart 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   4080
            Width           =   735
         End
         Begin VB.Label lblMedalOfHonor 
            BackStyle       =   0  'Transparent
            Caption         =   "Medal of Honor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   129
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label lblDistinguishedServiceCross 
            BackStyle       =   0  'Transparent
            Caption         =   "Distinguished Service Cross"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   131
            Top             =   2160
            Width           =   2655
         End
         Begin VB.Label lblSilverStar 
            BackStyle       =   0  'Transparent
            Caption         =   "Silver Star"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   133
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label lblDistinguishedFlyingCross 
            BackStyle       =   0  'Transparent
            Caption         =   "Distinguished Flying Cross"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   135
            Top             =   3120
            Width           =   2415
         End
         Begin VB.Label lblBronzeStarV 
            BackStyle       =   0  'Transparent
            Caption         =   "Bronze Star w/t V"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   137
            Top             =   3600
            Width           =   1695
         End
         Begin VB.Label lblPurpleHeart 
            BackStyle       =   0  'Transparent
            Caption         =   "Purple Heart"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   139
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label lblAirMedal 
            BackStyle       =   0  'Transparent
            Caption         =   "Air Medal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   141
            Top             =   4560
            Width           =   855
         End
         Begin VB.Label lblDistinguishedUnitCitation 
            BackStyle       =   0  'Transparent
            Caption         =   "Distinguished Unit Citation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   143
            Top             =   5040
            Width           =   2415
         End
         Begin VB.Label lblMeritoriousUnitCitation 
            BackStyle       =   0  'Transparent
            Caption         =   "Meritorious Unit Citation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   145
            Top             =   5520
            Width           =   2175
         End
         Begin VB.Label lblAwards 
            BackStyle       =   0  'Transparent
            Caption         =   "Awards"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   127
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblSorties 
            BackStyle       =   0  'Transparent
            Caption         =   "Missions"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   960
            TabIndex        =   124
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblKills 
            BackStyle       =   0  'Transparent
            Caption         =   "Kills"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   960
            TabIndex        =   126
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.ComboBox cboName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   110
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   109
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblRank 
         Caption         =   "Rank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   111
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblCrewPosition 
         Caption         =   "Crew Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   117
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label lblStatusLabel 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   119
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label lblSerialNumberLabel 
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   113
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblAssignment 
         Caption         =   "Assignment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   115
         Top             =   3000
         Width           =   1215
      End
   End
   Begin VB.Frame fraTab 
      Height          =   6840
      Index           =   1
      Left            =   120
      TabIndex        =   43
      Top             =   512
      Visible         =   0   'False
      Width           =   9515
      Begin VB.TextBox txtKeyField 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "txtKeyField(1)"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkDefault 
         Caption         =   "Default Squadron"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   5040
         Width           =   1815
      End
      Begin VB.ComboBox cboCommander 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox cboName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   2415
      End
      Begin VB.Frame fraBomberType 
         Caption         =   "Bomber Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   50
         Top             =   3000
         Width           =   2535
         Begin VB.OptionButton optB17FlyingFortress 
            Caption         =   "B-17 Flying Fortress"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optB24Liberator 
            Caption         =   "B-24 Liberator"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton optAvroLancaster 
            Caption         =   "Avro Lancaster"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1320
            Width           =   1815
         End
      End
      Begin VB.ComboBox cboGroup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmMainMenu.frx":0ECE
         Left            =   120
         List            =   "frmMainMenu.frx":0ED0
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Frame fraHistory 
         Caption         =   "Unit History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Index           =   1
         Left            =   2880
         TabIndex        =   56
         Top             =   120
         Width           =   6495
         Begin VB.TextBox txtMIA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   192
            TabStop         =   0   'False
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtPurpleHeart 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox txtBronzeStarV 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtDistinguishedFlyingCross 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtSilverStar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtDistinguishedServiceCross 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtMedalOfHonor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtPOW 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox txtWounded 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox txtKIA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtPlanesLost 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtKills 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtSorties 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtAirMedal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox txtDistinguishedUnitCitation 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   4080
            Width           =   735
         End
         Begin VB.TextBox txtMeritoriousUnitCitation 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   4560
            Width           =   735
         End
         Begin VB.Label lblMIALabel 
            BackStyle       =   0  'Transparent
            Caption         =   "MIA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   193
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label lblCasualties 
            BackStyle       =   0  'Transparent
            Caption         =   "Casualties"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   61
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblKills 
            BackStyle       =   0  'Transparent
            Caption         =   "Kills"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   60
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblSorties 
            BackStyle       =   0  'Transparent
            Caption         =   "Sorties"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   58
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblPOWLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "POW"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   69
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label lblWoundedLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Wounded"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   67
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label lblKIALabel 
            BackStyle       =   0  'Transparent
            Caption         =   "KIA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   65
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label lblPlanesLost 
            BackStyle       =   0  'Transparent
            Caption         =   "Planes Lost"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   63
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lblAwards 
            BackStyle       =   0  'Transparent
            Caption         =   "Awards"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2520
            TabIndex        =   70
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblMeritoriousUnitCitation 
            BackStyle       =   0  'Transparent
            Caption         =   "Meritorious Unit Citation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   88
            Top             =   4560
            Width           =   2175
         End
         Begin VB.Label lblDistinguishedUnitCitation 
            BackStyle       =   0  'Transparent
            Caption         =   "Distinguished Unit Citation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   86
            Top             =   4080
            Width           =   2415
         End
         Begin VB.Label lblAirMedal 
            BackStyle       =   0  'Transparent
            Caption         =   "Air Medal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   84
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label lblPurpleHeart 
            BackStyle       =   0  'Transparent
            Caption         =   "Purple Heart"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   82
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lblBronzeStarV 
            BackStyle       =   0  'Transparent
            Caption         =   "Bronze Star w/t V"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   80
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label lblDistinguishedFlyingCross 
            BackStyle       =   0  'Transparent
            Caption         =   "Distinguished Flying Cross"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   78
            Top             =   2160
            Width           =   2415
         End
         Begin VB.Label lblSilverStar 
            BackStyle       =   0  'Transparent
            Caption         =   "Silver Star"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   76
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lblDistinguishedServiceCross 
            BackStyle       =   0  'Transparent
            Caption         =   "Distinguished Service Cross"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   74
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label lblMedalOfHonor 
            BackStyle       =   0  'Transparent
            Caption         =   "Medal of Honor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   72
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Label lblCommander 
         Caption         =   "Commander"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblGroup 
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   2040
         Width           =   735
      End
   End
   Begin VB.Frame fraTab 
      Height          =   6840
      Index           =   2
      Left            =   120
      TabIndex        =   89
      Top             =   512
      Visible         =   0   'False
      Width           =   9515
      Begin VB.TextBox txtTailNumber 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   206
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtPlant 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   204
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1935
      End
      Begin VB.TextBox txtManufacturer 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   202
         TabStop         =   0   'False
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Frame fraHistory 
         Caption         =   "Flight Log"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Index           =   2
         Left            =   2880
         TabIndex        =   101
         Top             =   120
         Width           =   4575
         Begin VB.TextBox txtRabbitsFoot 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   195
            TabStop         =   0   'False
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtKills 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtSorties 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblRabbitsFoot 
            BackStyle       =   0  'Transparent
            Caption         =   "Rabbits Foot"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   196
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblKills 
            BackStyle       =   0  'Transparent
            Caption         =   "Kills"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   960
            TabIndex        =   105
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblSorties 
            BackStyle       =   0  'Transparent
            Caption         =   "Missions"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   960
            TabIndex        =   103
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdRetireBomber 
         Caption         =   "Retire Bomber"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   107
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1575
      End
      Begin VB.CommandButton cmdAssignCrew 
         Caption         =   "Assign Crew"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   106
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboBomberModel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   2400
         Width           =   2415
      End
      Begin VB.ComboBox cboSquadron 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CheckBox chkDefault 
         Caption         =   "Default Aircraft"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   100
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox txtKeyField 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         Text            =   "txtKeyField(2)"
         Top             =   6240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   91
         Text            =   "cboName(2)"
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblPlant 
         Caption         =   "Plant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   205
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label lblManufacturer 
         Caption         =   "Manufacturer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   203
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label lblStatusLabel 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   98
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label lblBomberModel 
         Caption         =   "Bomber Model"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   94
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblSquadron 
         Caption         =   "Squadron"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   96
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Tail Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   92
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   90
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame fraTab 
      Height          =   6840
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   512
      Visible         =   0   'False
      Width           =   9515
      Begin VB.TextBox txtKeyField 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "txtKeyField(0)"
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkDefault 
         Caption         =   "Default Group"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Frame fraHistory 
         Caption         =   "Unit History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Index           =   0
         Left            =   2880
         TabIndex        =   10
         Top             =   120
         Width           =   6495
         Begin VB.TextBox txtMIA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   190
            TabStop         =   0   'False
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtPurpleHeart 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox txtBronzeStarV 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtDistinguishedFlyingCross 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtSilverStar 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtDistinguishedServiceCross 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtMedalOfHonor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtPOW 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox txtWounded 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox txtKIA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtPlanesLost 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtKills 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtSorties 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtAirMedal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox txtDistinguishedUnitCitation 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   4080
            Width           =   735
         End
         Begin VB.TextBox txtMeritoriousUnitCitation 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   4560
            Width           =   735
         End
         Begin VB.Label lblMIALabel 
            BackStyle       =   0  'Transparent
            Caption         =   "MIA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   191
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label lblCasualties 
            BackStyle       =   0  'Transparent
            Caption         =   "Casualties"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblKills 
            BackStyle       =   0  'Transparent
            Caption         =   "Kills"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   14
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblSorties 
            BackStyle       =   0  'Transparent
            Caption         =   "Sorties"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   12
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblPOWLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "POW"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   23
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label lblWoundedLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Wounded"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   21
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label lblKIALabel 
            BackStyle       =   0  'Transparent
            Caption         =   "KIA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   19
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label lblPlanesLost 
            BackStyle       =   0  'Transparent
            Caption         =   "Planes Lost"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   17
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lblAwards 
            BackStyle       =   0  'Transparent
            Caption         =   "Awards"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2520
            TabIndex        =   24
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblMeritoriousUnitCitation 
            BackStyle       =   0  'Transparent
            Caption         =   "Meritorious Unit Citation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   42
            Top             =   4560
            Width           =   2175
         End
         Begin VB.Label lblDistinguishedUnitCitation 
            BackStyle       =   0  'Transparent
            Caption         =   "Distinguished Unit Citation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   40
            Top             =   4080
            Width           =   2415
         End
         Begin VB.Label lblAirMedal 
            BackStyle       =   0  'Transparent
            Caption         =   "Air Medal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   38
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label lblPurpleHeart 
            BackStyle       =   0  'Transparent
            Caption         =   "Purple Heart"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   36
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lblBronzeStarV 
            BackStyle       =   0  'Transparent
            Caption         =   "Bronze Star w/t V"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   34
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label lblDistinguishedFlyingCross 
            BackStyle       =   0  'Transparent
            Caption         =   "Distinguished Flying Cross"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   32
            Top             =   2160
            Width           =   2415
         End
         Begin VB.Label lblSilverStar 
            BackStyle       =   0  'Transparent
            Caption         =   "Silver Star"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   30
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lblDistinguishedServiceCross 
            BackStyle       =   0  'Transparent
            Caption         =   "Distinguished Service Cross"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   28
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label lblMedalOfHonor 
            BackStyle       =   0  'Transparent
            Caption         =   "Medal of Honor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   26
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame fraBase 
         Caption         =   "Base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1455
         Begin VB.OptionButton optEngland 
            Caption         =   "England"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optItaly 
            Caption         =   "Italy"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboCommander 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox cboName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblCommander 
         Caption         =   "Commander"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bomber Maintenance"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   3840
      TabIndex        =   189
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Generate Mission"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   7440
      TabIndex        =   188
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Airman Maintenance"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   5640
      TabIndex        =   187
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Squadron Maintenance"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1920
      TabIndex        =   186
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Group Maintenance"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   185
      Top             =   120
      Width           =   1740
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
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
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' frmMainMenu.frm
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
' The main menu makes use of Stefaan Casier's tab pseudo-control.
'******************************************************************************
'
' A Tab construction without the Tab-control
' Works as long as you don't need more than one row of Tabs
'
' H o w t o p r o c e e d
' 1. insert a copy of frmMainMenu into your project
' 2. remove the End keyword in cmdOK & cmdCancel
' 3. from within your project you call frmMainMenu as follows:
'
'         Load frmMainMenu
'         frmMainMenu. ... = ...         ' set current values
'         ...
'         frmMainMenu.Show 1, Me         ' show as modal form
'         If frmMainMenu.OK = True Then
'             ... = frmMainMenu. ...     ' read & process new values
'             ...
'             End If
'         Unload frmMainMenu
'
' 4. provide the proper amount of tab-controls (delete or add existing ones)
'         lblTab(...)
'         fraTab(...)
'         lblTabTitle(...) - if you want to use titles, that is
'    adjust the TabsCount constant, here beneath
' 5. fill in the lblTab() / lblTabTitle() .captions
' 6. position the cmdOK (cmdCancel accordingly) - the drawing of
'    raised borders + width of this form, will be aligned/adjusted
'    to this position
' 7. adjust the fraTab() size to what you need and place your own
'    option controls inside each frame plus their proper code
'
'******************************************************************************

Option Explicit

Const TabsCount = 5
Dim CurrentTab As Integer

'******************************************************************************
' Form_Activate
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Gets fired when the emulator is started and when the user returns
'         to this form from another one.
'******************************************************************************
Private Sub Form_Activate()

'MsgBox Bomber.Name & ", " & Bomber.Status

    If Bomber.Status <> DUTY_STATUS Then
        ' Bomber status is 0 when the program starts, or after a mission where
        ' the bomber is not shot down and loses no airmen. If the bomber is not
        ' on duty status (0), then it must be removed from the combo.
Call AdjustAvailableBombers ' Nov04
'        Call AdjustMissionAvailableBombers
    End If

End Sub

'******************************************************************************
' Form_Load
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Initialize the randomizer, create the tab pseudo-control, open a DB
'         connection, get recordsets, populate combos, bookmark the first
'         record on each tab.
'******************************************************************************
Private Sub Form_Load()
    On Error GoTo ErrorTrap
    
    Dim strErrMsg As String
    Dim dbConn
    
    Randomize
    
    ' Draw the pseudo-tab control, then center the form.
    
    Call StartTab
    'CenterForm Me
    
    ' Fiddle the form bottom, as adding a menu bar otherwise seems to
    ' randomly cut off the bottom of the form
    'Me.Height = cmdAdd.Top + cmdAdd.Height + 880 '(Me.Height - Me.ScaleHeight)
    
    ' Open the DB connection.
' ??? Leave it open as long as the emulator is running???

    dbConn = OpenDBConnection()

    If dbConn = False Then
        Call ExitEmulator
    End If
    
    ' Grab all our recordsets to avoid making multiple database hits each
    ' time a different record is displayed on a tab. Once the recordsets
    ' are all queried, the database will only be hit to add, update or
    ' delete records. (The recordsets will be modified first, then the
    ' corresponding table: If one fails, they both fail.) If a mission is
    ' being flown, then the plane and airmen will not be updated until
    ' the mission is complete -- giving the user a de facto abort/cheat.

    If GetGroupRecordset() = False Then
        ' The table is named "GroupT" because "Group" is a reserved SQL word.
        Call ExitEmulator
    ElseIf GetSquadronRecordset() = False Then
        Call ExitEmulator
    ElseIf GetBomberRecordset() = False Then
        Call ExitEmulator
    ElseIf GetBomberModelRecordset() = False Then
        Call ExitEmulator
    ElseIf GetBomberStatusRecordset() = False Then
        Call ExitEmulator
    ElseIf GetAirmanRecordset() = False Then
        Call ExitEmulator
    ElseIf GetRankRecordset() = False Then
        Call ExitEmulator
    ElseIf GetCrewPositionRecordset() = False Then
        Call ExitEmulator
    ElseIf GetAirmanStatusRecordset() = False Then
        Call ExitEmulator
    ElseIf GetTargetRecordset() = False Then
        Call ExitEmulator
    End If

    ' Successfully obtained all recordsets. Populate the combos, then repoint
    ' the recordsets and initialize the tabs.

    Call PopulateGroupCombos
    Call PopulateSquadronCombos
    Call PopulateBomberCombos
    Call PopulateBomberModelCombo
    Call PopulateAirmanCombos
    Call PopulateRankCombo
    Call PopulateCrewPositionCombo
    
    ' Bookmark the first record (that will be displayed on the tab), so
    ' so that it can be returned to if off-tab operations iterate the
    ' recordset.
    
    prsGroup.MoveFirst
    varGroupCurrentlyOnTab = prsGroup.Bookmark
    
    prsSquadron.MoveFirst
    varSquadronCurrentlyOnTab = prsSquadron.Bookmark
    
    prsBomber.MoveFirst
    varBomberCurrentlyOnTab = prsBomber.Bookmark
    
    prsAirman.MoveFirst
    varAirmanCurrentlyOnTab = prsAirman.Bookmark
    
    ' The MISSION_TAB does not have any saved records to be displayed on
    ' the tab. Missions are created on the tab, then either flown or lost.

    ' If the combos were not pointed at the first 0-base record, they would
    ' initially be pointing at the -1 wildspace record. Setting the values
    ' this way invokes the click method, which in turn calls the fill tab
    ' fields function for the object.
    
    cboName(GROUP_TAB).ListIndex = 0
    cboName(SQUADRON_TAB).ListIndex = 0
    cboName(BOMBER_TAB).ListIndex = 0
    cboName(AIRMAN_TAB).ListIndex = 0

    ' Populate hard-coded combos that don't have an associated recordset.
    
    Call PopulatePositionCombos
    Call PopulateTargetCombo
    Call PopulateDateCombos
    
    ' Since each tab is displaying the first record, the tab-associated
    ' recordsets should be re-pointed to the first record so the recordset
    ' is in synch with the tab.

    prsGroup.MoveFirst
    prsSquadron.MoveFirst
    prsBomber.MoveFirst
    prsAirman.MoveFirst

    Exit Sub

ErrorTrap:
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & _
                "Form_Load() " & vbCrLf

    strErrMsg = strErrMsg & Err.Description
    
    MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    Call ExitEmulator

End Sub

'******************************************************************************
' StartTab
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Draw the tab pseudo-control.
'******************************************************************************
Private Sub StartTab()
    Dim i As Long
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    
    Me.AutoRedraw = True

    ' Position the tab buttons and frames, based on the location of
    ' the base button and frame. Hide the design-time borders.

    For i = lblTab.LBound To lblTab.UBound
        ' Hide design-time borders, so we can draw them later.
        lblTab(i).BorderStyle = vbBSNone
        fraTab(i).BorderStyle = vbBSNone
        If i = 0 Then
            ' Position first tab button in upper left corner.
            lblTab(i).Left = 80
            lblTab(i).Top = 160
        Else
            lblTab(i).Left = lblTab(i - 1).Left + lblTab(i - 1).Width
            lblTab(i).Top = lblTab(0).Top
            fraTab(i).Left = fraTab(0).Left
            fraTab(i).Top = fraTab(0).Top
            fraTab(i).BorderStyle = 0
            fraTab(i).Visible = False
        End If
    Next i
    
    ' Draw tab buttons.
    
    For i = lblTab.LBound To lblTab.UBound
        X1 = lblTab(i).Left                                    ' horizontal start point
        Y1 = lblTab(i).Top - 64                                ' vertical start point
        X2 = lblTab(i).Left + lblTab(i).Width - 32             ' horizontal end point
        Y2 = lblTab(i).Top + lblTab(i).Height                  ' vertical end point
        Line (X1 + 16, Y1)-(X2 - 16, Y1), vb3DHighlight     ' top line of tab label
        Line (X1, Y1 + 16)-(X1, Y2), vb3DHighlight          ' left line of tab label
        Line (X2, Y1 + 16)-(X2, Y2), vb3DShadow             ' right line of tab label
        Line (X2 + 16, Y1 + 32)-(X2 + 16, Y2), vb3DDKShadow ' right line shadow
    Next i
    
    ' Draw tab frame.
    
    X1 = lblTab(0).Left                   ' horizontal start point
    Y1 = lblTab(0).Top + lblTab(0).Height ' vertical start point
    X2 = fraTab(0).Left + fraTab(0).Width ' horizontal end point
    Y2 = fraTab(0).Top + fraTab(0).Height ' vertical end point

    Line (X1, Y1)-(X2, Y1), vb3DHighlight               ' top line of tab body
    Line (X1, Y1 + 16)-(X1, Y2 - 16), vb3DHighlight     ' left line of tab body
    Line (X1 + 16, Y2)-(X2, Y2), vb3DShadow             ' bottom line of tab body
    Line (X1, Y2 + 16)-(X2 + 16, Y2 + 16), vb3DDKShadow ' bottom line shadow
    Line (X2, Y1 + 16)-(X2, Y2 + 16), vb3DShadow        ' right line of tab body
    Line (X2 + 16, Y1)-(X2 + 16, Y2 + 16), vb3DDKShadow ' right line shadow
    
    Picture = Image ' Make the 'frame' we just drew the permanent background
    AutoRedraw = False
    
    ' Group is the default tab, so make that the form caption.
    
    Me.Caption = App.Title & " (" & lblTab(GROUP_TAB).Caption & ")"

End Sub

'******************************************************************************
' SelectTab
'
' INPUT:  The selected tab.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Redraw the tab pseudo-control.
'******************************************************************************
Public Sub SelectTab(ByVal i As Integer)
    Dim X1 As Long
    Dim X2 As Long
    Dim Y As Long
    fraTab(CurrentTab).Visible = False
    
    CurrentTab = i
    fraTab(CurrentTab).Visible = True

    X1 = lblTab(i).Left + 1 ' draw new Tab selection
    X2 = lblTab(i).Left + lblTab(i).Width - 2 ' ???
    Y = lblTab(i).Top + lblTab(i).Height
    Me.Cls
    Me.Line (X1, Y)-(X2, Y), vbButtonFace
    Me.PSet (X1 - 1, Y), vb3DHighlight
    Me.PSet (X2, Y), vb3DShadow
    Me.PSet (X2 + 1, Y), vb3DDKShadow
    
End Sub

'******************************************************************************
' Form_MouseDown
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Call SelectTab.
'******************************************************************************
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim i As Long
   For i = 0 To TabsCount - 1
      If X > lblTab(i).Left And X < lblTab(i).Left + lblTab(i).Width Then SelectTab i: Exit For
   Next i
End Sub

'******************************************************************************
' Form_Paint()
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub Form_Paint()
    SelectTab CurrentTab
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ExitEmulator
End Sub

'******************************************************************************
' lblTab_Click
'
' INPUT:  The selected tab.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Reset the caption to reflect the current tab, then show the appropriate
'         buttons.
'******************************************************************************
Private Sub lblTab_Click(intIndex As Integer)
    
    SelectTab intIndex
    Me.Caption = App.Title & " (" & lblTab(intIndex).Caption & ")"

    If intIndex = MISSION_TAB Then
        ' Missions are not recorded in the database, so the add, update and
        ' delete buttons are hidden.
            
        cmdAdd.Visible = False
        cmdUpdate.Visible = False
        cmdDelete.Visible = False

        ' These buttons are only visible on the generate mission tab.
        cmdSave.Visible = True
        cmdFlyMission.Visible = True
        'CenterControl fraExitHelp, Me
    Else
        ' The user is doing maintenance, so the database buttons are revealed.
        
        fraExitHelp.Left = 2835 '5715
        cmdAdd.Visible = True
        cmdUpdate.Visible = True
        cmdDelete.Visible = True
        
        ' These buttons are hidden when on the maintenance tabs.
        cmdSave.Visible = False
        cmdFlyMission.Visible = False
    End If

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
' qwe
Public Sub ExitEmulator()
    ' Gracefully shut down the emulator: Free memory, close the DB connection,
    ' then exit.
    
    Call FreeRecordset(prsGroup)
    Call FreeRecordset(prsSquadron)
    Call FreeRecordset(prsBomber)
    Call FreeRecordset(prsBomberModel)
    Call FreeRecordset(prsBomberStatus)
    Call FreeRecordset(prsBomberSquadron)
    Call FreeRecordset(prsAirman)
    Call FreeRecordset(prsRank)
    Call FreeRecordset(prsCrewPosition)
    Call FreeRecordset(prsAirmanStatus)
    
    Call CloseDBConnection

    Unload Me
    End

End Sub

'******************************************************************************
' cmdAdd_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Branch to the routine appropriate to the current tab.
'******************************************************************************
Private Sub cmdAdd_Click()
    ' MAIN SCREEN
        
    Select Case CurrentTab
        Case GROUP_TAB:
            
            If AddGroup() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

        Case SQUADRON_TAB:
            
            If AddSquadron() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

        Case BOMBER_TAB:
        
            If AddBomber() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

        Case AIRMAN_TAB:
            
            If AddAirman() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

    End Select
    
End Sub

'******************************************************************************
' cmdUpdate_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Branch to the routine appropriate to the current tab.
'******************************************************************************
Private Sub cmdUpdate_Click()
    ' MAIN SCREEN
        
    Select Case CurrentTab
        Case GROUP_TAB:
            
            If ModifyGroup() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

        Case SQUADRON_TAB:

' TODO: Should not be able to change plane type of squadron if the squadron has
' planes assigned to it
            
            If ModifySquadron() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

        Case BOMBER_TAB:
        
            If ModifyBomber() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

        Case AIRMAN_TAB:
            
            If ModifyAirman() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

    End Select
    
End Sub

'******************************************************************************
' cmdDelete_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Branch to the routine appropriate to the current tab.
'******************************************************************************
Private Sub cmdDelete_Click()
    ' MAIN SCREEN

    Select Case CurrentTab
        Case GROUP_TAB:
            
            If DeleteGroup() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

        Case SQUADRON_TAB:
            
            If DeleteSquadron() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

        Case BOMBER_TAB:
        
            If DeleteBomber() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

        Case AIRMAN_TAB:
            
            If DeleteAirman() = False Then
' qwe                Exit Sub
                Call ExitEmulator
            End If

    End Select
    
End Sub

'******************************************************************************
' mnuFileExit_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Shut down the app.
'******************************************************************************
Private Sub mnuFileExit_Click()
    ' MENU
    If MsgBox("Do you wish to exit?", (vbYesNo + vbDefaultButton2 + vbQuestion)) = vbYes Then
        Call ExitEmulator
    End If
End Sub

'******************************************************************************
' cmdExit_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Shut down the app.
'******************************************************************************
Private Sub cmdExit_Click()
    ' MAIN SCREEN
    If MsgBox("Do you wish to exit?", (vbYesNo + vbDefaultButton2 + vbQuestion)) = vbYes Then
        Call ExitEmulator
'Unload Me 'frmMainMenu
    End If
End Sub

'******************************************************************************
' cmdHelp_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Pop the swap ammo help screen.
'******************************************************************************
Private Sub cmdHelp_Click()
    ' MAIN SCREEN
    
    frmHelpBrowser.Hide

    frmHelpBrowser.txtPageName.Text = "doc/B17" & Replace(lblTab(CurrentTab).Caption, " ", "") & "Help.html"

    frmHelpBrowser.Show

End Sub
    
'******************************************************************************
' mnuHelpAbout_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Pop the about screen.
'******************************************************************************
Private Sub mnuHelpAbout_Click()
    ' MENU
    frmAbout.Show
End Sub

'******************************************************************************
' mnuHelpIndex_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Pop the help index screen.
'******************************************************************************
Private Sub mnuHelpIndex_Click()
    ' MENU
    frmHelpBrowser.Hide
    
    frmHelpBrowser.txtPageName.Text = "doc/B17HelpIndex.html"

    frmHelpBrowser.Show
End Sub

'******************************************************************************
' cmdAssignCrew_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Pop the crew assignment form.
'******************************************************************************
Private Sub cmdAssignCrew_Click()
    ' BOMBER_TAB
' qwe
    gblnCrewAssigned = True
    
    frmCrewAssignment.Show vbModal
    
    If gblnCrewAssigned = False Then
        Call ExitEmulator
    End If
    
Call AdjustAvailableBombers ' Nov04
'    Call AdjustMissionAvailableBombers
End Sub

'******************************************************************************
' cmdRetireBomber_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Voluntarily retire a bomber from the game. Otherwise the bomber may
'         be used until it is shot down.
'******************************************************************************
Private Sub cmdRetireBomber_Click()
    ' BOMBER_TAB
    Dim strMsg As String
    
    strMsg = "Retiring a bomber means that it will no longer be " & _
             "available for combat duty, and its whole crew will " & _
             "placed on admin duty pending new assignments. " & _
             vbCrLf & vbCrLf & _
             "It cannot be undone. " & _
             vbCrLf & vbCrLf & _
             "Are you sure?"

    If MsgBox(strMsg, (vbYesNo + vbDefaultButton2 + vbQuestion)) = vbNo Then
        Exit Sub
    End If

    If RetireBomber() = False Then
        Exit Sub
    End If

End Sub

'******************************************************************************
' cmdRandomTarget_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  User wants to be assigned a target, rather than choosing one for
'         himself.
'******************************************************************************
Private Sub cmdRandomTarget_Click()
    ' MISSION_TAB
    cboTarget.ListIndex = G1MissionTarget(cboTarget.ListCount)
End Sub
    
'******************************************************************************
' cmdRandomDate_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  User wants to be assigned a date, rather than choosing one for
'         himself.
'******************************************************************************
Private Sub cmdRandomDate_Click()
    ' MISSION_TAB

    ' Strategic bombing missions did not begin until August, 1942,
    ' continuing until late April, 1945. (In Europe.) VE-Day was in May,
    ' 1945, so we allow for the possibility of missions in the war's
    ' final month as well.
    
    Dim intMonth As Integer
    
    Do While True
        
        ' Loop until a valid month is generated.
    
        intMonth = RandomMonth()
    
        ' If this is the 8th Air Force, or either air force on or after
        ' November, 1943, then a legitimate month was generated. Otherwise,
        ' get a new month. (The 15th Air Force did not fly missions until
        ' November, 1943.)
    
        If optEngland(MISSION_TAB).Value = True _
        Or intMonth >= NOV_1943 Then
            Exit Do
        End If
    
    Loop

    Select Case intMonth
        Case AUG_1942 To DEC_1942:
            
            cboYear.ListIndex = 0
            cboMonth.ListIndex = intMonth
        
        Case JAN_1943 To DEC_1943:
            
            If optEngland(MISSION_TAB).Value = True Then
                cboYear.ListIndex = 1
                cboMonth.ListIndex = (intMonth - 5)
            Else ' optItaly(MISSION_TAB).Value = True
                cboYear.ListIndex = 0
                cboMonth.ListIndex = (intMonth - 15)
            End If
        
        Case JAN_1944 To DEC_1944:
            
            If optEngland(MISSION_TAB).Value = True Then
                cboYear.ListIndex = 2
                cboMonth.ListIndex = (intMonth - 17)
            Else ' optItaly(MISSION_TAB).Value = True
                cboYear.ListIndex = 1
                cboMonth.ListIndex = (intMonth - 17)
            End If
        
        Case JAN_1945 To MAY_1945:
            
            If optEngland(MISSION_TAB).Value = True Then
                cboYear.ListIndex = 3
                cboMonth.ListIndex = (intMonth - 29)
            Else ' optItaly(MISSION_TAB).Value = True
                cboYear.ListIndex = 2
                cboMonth.ListIndex = (intMonth - 29)
            End If
        
    End Select
    
End Sub

'******************************************************************************
' cboYear_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  The available months and options depends on the date.
'******************************************************************************
Private Sub cboYear_Click()
    ' MISSION_TAB
    Call AdjustDateLists
    Call AdjustMissionOptions
End Sub

'******************************************************************************
' cboMonth_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES: The available options depends on year and month.
'******************************************************************************
Private Sub cboMonth_Click()
    ' MISSION_TAB
    Call AdjustMissionOptions
End Sub

'******************************************************************************
' chkTimePeriodSpecificFormations_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Availability of options depends on whether the user is using
'         historical tactics.
'******************************************************************************
Private Sub chkTimePeriodSpecificFormations_Click()
    Call AdjustMissionOptions
End Sub

'******************************************************************************
' AdjustMissionOptions
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  The available options depends on year, month and tactics.
'******************************************************************************
Private Sub AdjustMissionOptions()
    Dim intDate

    intDate = GetDateValue()

    If chkTimePeriodSpecificFormations.Value = vbChecked Then
        
        Select Case intDate
            Case AUG_1942:
            
                chkEvadeFlak.Enabled = True
                chkCrewExperience.Enabled = False
                chkCrewExperience.Value = vbUnchecked
                chkFormationDefensiveGunnery.Enabled = False
                chkFormationDefensiveGunnery.Value = vbUnchecked
            
            Case SEP_1942 To MAR_1943:
            
                chkEvadeFlak.Enabled = True
                chkCrewExperience.Enabled = False
                chkCrewExperience.Value = vbUnchecked
                chkFormationDefensiveGunnery.Enabled = True
            
            Case APR_1943 To MAY_1945:
            
                chkEvadeFlak.Enabled = False
                chkEvadeFlak.Value = vbUnchecked
                chkCrewExperience.Enabled = True
                chkFormationDefensiveGunnery.Enabled = True
            
        End Select
    
    Else
    
        chkEvadeFlak.Enabled = True
        chkCrewExperience.Enabled = True
        chkFormationDefensiveGunnery.Enabled = True
    
    End If
    
End Sub

'******************************************************************************
' cmdRandomPosition_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  User wants to be assigned a position, rather than choosing one for
'         himself.
'******************************************************************************
Private Sub cmdRandomPosition_Click()
    ' MISSION_TAB

    cboSquadronPos.Text = G4SquadronPosition()
    cboFormationPos.Text = G4FormationPosition()
    
End Sub

'******************************************************************************
' cmdFlyMission_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Validate the required mission parameters, initialize the mission
'         structures, launch mission form, then hide this form.
'******************************************************************************
Private Sub cmdFlyMission_Click()
    ' MISSION_TAB

    ' Validate required fields were entered/selected.
    
    If cboName(MISSION_TAB).Text = "" Then
        MsgBox "Bomber?", (vbQuestion + vbOKOnly)
        Exit Sub
    ElseIf cboTarget.Text = "" Then
        MsgBox "Target?", (vbQuestion + vbOKOnly)
        Exit Sub
    ElseIf cboMonth.Text = "" Then
        MsgBox "Month?", (vbQuestion + vbOKOnly)
        Exit Sub
    ElseIf cboYear.Text = "" Then
        MsgBox "Year?", (vbQuestion + vbOKOnly)
        Exit Sub
    ElseIf cboSquadronPos.Text = "" Then
        MsgBox "Squadron Position?", (vbQuestion + vbOKOnly)
        Exit Sub
    ElseIf cboFormationPos.Text = "" Then
        MsgBox "Formation Position?", (vbQuestion + vbOKOnly)
        Exit Sub
    End If
    
    If LookupBomber(intBomberMission(cboName(MISSION_TAB).ListIndex), LOOKUP_BY_KEYFIELD, vbNullString) = False Then
        Call ExitEmulator
    End If

    Call InitializeMission

' qwe   Call InitializeBomber
    If InitializeBomber() = False Then
        Call ExitEmulator
    End If
    
    Call InitializeRandomEvents

    prsBomber.Bookmark = varBomberCurrentlyOnTab
        
    Load frmMission
    frmMission.Show

    Me.Hide

End Sub

'******************************************************************************
' cmdSave_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Save the mission parameters to a neatly formatted HTML document.
'******************************************************************************
Private Sub cmdSave_Click()
    On Error GoTo ErrorTrap
    ' MISSION_TAB
    
    Dim blnOpenFile As Boolean
    Dim strErrMsg As String
    Dim intIndex As Integer
    Dim strFileType As String
    Dim strHeader As String
    Dim strBody As String
    Dim strFooter As String

    If LookupBomber(intBomberMission(cboName(MISSION_TAB).ListIndex), LOOKUP_BY_KEYFIELD) = False Then
        Call ExitEmulator
    End If

    Call InitializeMission
' qwe   Call InitializeBomber
    If InitializeBomber() = False Then
        Call ExitEmulator
    End If
    
    prsBomber.Bookmark = varBomberCurrentlyOnTab
        
    ' Remove any text from previous invocation.
    
    dlgFile.FileName = ""

    ' Treat Cancel button as an error, so we can exit the file type loop.
    
    dlgFile.CancelError = True
    
    ' Display overwrite prompt if file exists.
    
    dlgFile.Flags = cdlOFNOverwritePrompt
    
    ' Only display files ending in allowable types.
    
    dlgFile.Filter = "HTML (*.html;*.htm)|*.html;*.htm"

    Do While True
    
        ' Loop until a valid file type is entered.
        
        dlgFile.ShowSave
        
        ' Wait here for the user to click OK or cancel.
        
        intIndex = InStr(1, dlgFile.FileName, ".")
        strFileType = LCase(Mid(dlgFile.FileName, (intIndex + 1)))

        If strFileType <> "html" _
        And strFileType <> "htm" Then
            strErrMsg = "Missions should be saved as an HTML file." & vbCrLf & vbCrLf & _
                        "Save mission as " & dlgFile.FileTitle & "?"

            If MsgBox(strErrMsg, (vbExclamation + vbYesNo)) = vbYes Then
                Exit Do
            End If
        Else
            Exit Do
        End If

    Loop

    Open dlgFile.FileName For Output As #1

    blnOpenFile = True

    Call CreateMissionHTML(strHeader, strBody, strFooter)
    
    Print #1, strHeader
    Print #1, strBody
    Print #1, strFooter
    
CleanUp:
   
    If blnOpenFile = True Then
        Close #1
    End If

    Exit Sub

ErrorTrap:
    
    If Err.Number = FILE_DIALOG_CANCEL Then
        Resume CleanUp
    End If
    
    ' If FILE_NOT_FOUND, the file will be created, so FILE_NOT_FOUND is not
    ' an error.
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & _
                "cmdSave_Click() " & vbCrLf

    strErrMsg = strErrMsg & Err.Description
    
    MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    Resume CleanUp

End Sub

'******************************************************************************
' cboTarget_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  YB-40s have limitations depending on the range of the chosen target.
'******************************************************************************
Private Sub cboTarget_Click()

    ' All we need is the pointer to the matching record.
    If LookupBomberTarget((cboTarget.ListIndex + 1), LOOKUP_BY_LISTINDEX) = False Then
        Call ExitEmulator
    End If

End Sub

'******************************************************************************
' cboBomberModel_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  The list of available squadrons depends on the chosen model (e.g.,
'         the user can't assign a B-17 to a Lancaster squadron).
'******************************************************************************
Private Sub cboBomberModel_Click(Index As Integer)
'MsgBox "cboBomberModel_Click"
    If Index <> MISSION_TAB Then
        Call PopulateBomberSquadronCombo
    End If
End Sub

'******************************************************************************
' cboBomberModel_Change
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  The list of available squadrons depends on the chosen model (e.g.,
'         the user can't assign a B-17 to a Lancaster squadron).
'******************************************************************************
Private Sub cboBomberModel_Change(Index As Integer)
MsgBox "cboBomberModel_Change"
    Call PopulateBomberSquadronCombo
End Sub

'******************************************************************************
' cboName_Change
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Names can't be changed. If a user attempts to 'change' a name, then
'         that is the signal to create a new instance of that entity.
'******************************************************************************
Private Sub cboName_Change(Index As Integer)

    ' This event should only fire when the user types something in the
    ' text portion of the name combo. When the user types something in
    ' the text portion of the name combo, that indicates the user wants
    ' to add a new record for the current tab. Regardless of whether or
    ' not the current recordset is a default record, enable all
    ' modifiable fields on the tab.
    
    Select Case CurrentTab
    
        Case GROUP_TAB:
            
            If cboName(GROUP_TAB).Text <> prsGroup![Name] Then
                Call ZeroGroupTabFields
            End If
    
        Case SQUADRON_TAB:
            
            If cboName(SQUADRON_TAB).Text <> prsSquadron![Name] Then
                Call ZeroSquadronTabFields
            End If
    
        Case BOMBER_TAB:
            
            If cboName(BOMBER_TAB).Text <> prsBomber![Name] Then
' qwe           Call ZeroBomberTabFields
                If ZeroBomberTabFields() = False Then
                    Call ExitEmulator
                End If
            End If
    
        Case AIRMAN_TAB:
            
            If cboName(AIRMAN_TAB).Text <> prsAirman![Name] Then
' qwe                Call ZeroAirmanTabFields
                If ZeroAirmanTabFields() = False Then
                    Call ExitEmulator
                End If
            End If
    
    End Select
        
End Sub

'******************************************************************************
' cboName_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  When the user selects a name -- either by clicking on the drop down
'         or scrolling through the list -- then the data related to that item
'         should be displayed on the current tab.
'******************************************************************************
Private Sub cboName_Click(Index As Integer)
    ' GROUP_TAB, SQUADRON_TAB, BOMBER_TAB, AIRMAN_TAB, MISSION_TAB


    ' The ListIndex property is 0-base, but emulator KeyFields are 1-base.
    ' Therefore, 1 must be added to .ListIndex before a lookup may be
    ' correctly performed.

    Select Case Index
        Case GROUP_TAB:
            
            If LookupGroup((cboName(GROUP_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX) = True Then
                ' All we need is the pointer to the matching record (the
                ' group name was selected). Fill in the tab fields.
                If FillGroupTabFields() = False Then
                    Call ExitEmulator
                End If
            Else
                Call ExitEmulator
            End If

            ' A new record is being displayed on the tab. Bookmark it.
            varGroupCurrentlyOnTab = prsGroup.Bookmark

        Case SQUADRON_TAB:
            
            If LookupSquadron((cboName(SQUADRON_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX) = True Then
                ' All we need is the pointer to the matching record (the
                ' squadron name was selected). Fill in the tab fields.
                If FillSquadronTabFields() = False Then
                    Call ExitEmulator
                End If
            Else
                Call ExitEmulator
            End If

            ' A new record is being displayed on the tab. Bookmark it.
            varSquadronCurrentlyOnTab = prsSquadron.Bookmark

        Case BOMBER_TAB:
            
            If LookupBomber((cboName(BOMBER_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX) = True Then
                ' All we need is the pointer to the matching record (the
                ' bomber name was selected). Fill in the tab fields.
                If FillBomberTabFields() = False Then
                    Call ExitEmulator
                End If
            Else
                Call ExitEmulator
            End If

            ' A new record is being displayed on the tab. Bookmark it.
            varBomberCurrentlyOnTab = prsBomber.Bookmark
        
        Case AIRMAN_TAB:
            
            If LookupAirman((cboName(AIRMAN_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX) = True Then
                ' All we need is the pointer to the matching record (the
                ' airman name was selected). Fill in the tab fields.
                
                If FillAirmanTabFields() = False Then
                    Call ExitEmulator
                End If
            Else
                Call ExitEmulator
            End If

            ' A new record is being displayed on the tab. Bookmark it.
            varAirmanCurrentlyOnTab = prsAirman.Bookmark

        Case MISSION_TAB:
    
            If FillMissionTabFields() = False Then
                Call ExitEmulator
            End If
        
            ' The mission list and available dates may change when a
            ' different bomber is selected, therefore the combos must
            ' be re-populated. Also, blank out the text portions of the
            ' target, date and position controls.
            
            Call PopulateTargetCombo
            Call PopulateDateCombos
            cboSquadronPos.ListIndex = -1
            cboFormationPos.ListIndex = -1
    
    End Select

End Sub

'******************************************************************************
' chkExpandedTargetList_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  The user wants to choose from more targets than were part of the
'         original game/variant.
'******************************************************************************
Private Sub chkExpandedTargetList_Click()
    Call PopulateTargetCombo
End Sub

'******************************************************************************
' txtLogSpeed_LostFocus
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  The amount of delay, in milliseconds, after each line is printed
'         to the mission log.
'******************************************************************************
Private Sub txtLogSpeed_LostFocus()
    On Error GoTo ErrorTrap
    
    If CInt(txtLogSpeed.Text) > 10 Then
        txtLogSpeed.Text = "10"
    ElseIf CInt(txtLogSpeed.Text) < 0 Then
        txtLogSpeed.Text = "0"
    End If
    
    Exit Sub
   
ErrorTrap:
    
    txtLogSpeed.Text = "0"
    
    Resume Next

End Sub
