VERSION 5.00
Begin VB.Form frmCrewAssignment 
   Caption         =   "Crew Assignment"
   ClientHeight    =   6570
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5310
   Icon            =   "frmCrewAssignment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCancelHelp 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2040
      TabIndex        =   40
      Top             =   5880
      Width           =   2655
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
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
         TabIndex        =   42
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help"
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
         TabIndex        =   41
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   13
      Left            =   2160
      TabIndex        =   39
      Text            =   "0"
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   12
      Left            =   1200
      TabIndex        =   38
      Text            =   "0"
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   11
      Left            =   240
      TabIndex        =   37
      Text            =   "0"
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   10
      Left            =   4080
      TabIndex        =   36
      Text            =   "0"
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   9
      Left            =   3120
      TabIndex        =   35
      Text            =   "0"
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   8
      Left            =   2160
      TabIndex        =   34
      Text            =   "0"
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   7
      Left            =   1200
      TabIndex        =   33
      Text            =   "0"
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   32
      Text            =   "0"
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   5
      Left            =   4080
      TabIndex        =   31
      Text            =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   4
      Left            =   3120
      TabIndex        =   30
      Text            =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   29
      Text            =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   28
      Text            =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Text            =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
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
      Index           =   12
      Left            =   2760
      TabIndex        =   24
      Top             =   5280
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
      Index           =   11
      Left            =   2760
      TabIndex        =   22
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   608
      TabIndex        =   21
      Top             =   5880
      Width           =   1215
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
      Index           =   10
      Left            =   2760
      TabIndex        =   18
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
      Index           =   9
      Left            =   2760
      TabIndex        =   16
      Top             =   2400
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
      Index           =   8
      Left            =   2760
      TabIndex        =   14
      Top             =   1440
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
      Index           =   7
      Left            =   2760
      TabIndex        =   12
      Top             =   480
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
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   5280
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
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   4320
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
      Index           =   4
      Left            =   120
      TabIndex        =   6
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
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   2400
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
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1440
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
      Index           =   1
      ItemData        =   "frmCrewAssignment.frx":0ECA
      Left            =   120
      List            =   "frmCrewAssignment.frx":0ECC
      TabIndex        =   0
      Top             =   480
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
      Index           =   13
      Left            =   2760
      TabIndex        =   25
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      Index           =   12
      Left            =   2760
      TabIndex        =   23
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      Index           =   11
      Left            =   2760
      TabIndex        =   20
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      Index           =   10
      Left            =   2760
      TabIndex        =   19
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      Index           =   9
      Left            =   2760
      TabIndex        =   17
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      Index           =   8
      Left            =   2760
      TabIndex        =   15
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      Index           =   7
      Left            =   2760
      TabIndex        =   13
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      TabIndex        =   7
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      TabIndex        =   5
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblCrewPosition 
      Caption         =   "lblCrewPosition"
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
      Index           =   13
      Left            =   2760
      TabIndex        =   26
      Top             =   5880
      Width           =   2415
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
Attribute VB_Name = "frmCrewAssignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' frmCrewAssignment.frm
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

Option Explicit

Dim strErrMsg As String
Dim varBookmark As Variant
Dim intCurrentAirmanOldAssigment As Integer

' This form utilizes a jagged multidimensional array of the following
' format:
'
'   (PILOT)(0,1)(0,1)(0,1)(0,1)
'   (COPILOT)(0,1)
'   ...
'   (AMMO_STOCKER)(0,1)(0,1)
'
' The number of rows is fixed to the number of crew positions, and the
' nodes on each row are two-dimensional, but the number of nodes on each
' row varies. The number of nodes on a row should equal the number of rows
' in the position's combo, not including the blank row. (There is no blank
' node in the matrix.) It is possible a row may not have any nodes at all.
Dim lvntCrewMatrix(PILOT To AMMO_STOCKER) As Variant
Dim lvntNode() As Variant
Dim PositionNames As Variant
Dim B17CPositions As Variant
Dim B17CPositionNames As Variant
Dim B17EFGPositions As Variant
Dim B17EFGPositionNames As Variant
Dim YB40Positions As Variant
Dim YB40PositionNames As Variant
Dim B24DEPositions As Variant
Dim B24DEPositionNames As Variant
Dim B24GHJPositions As Variant
Dim B24GHJPositionNames As Variant
Dim B24LMPositions As Variant
Dim B24LMPositionNames As Variant
Dim LancasterPositions As Variant
Dim LancasterPositionNames As Variant
Dim CurrentBomberPositions As Variant
Dim CurrentBomberPositionNames As Variant
'******************************************************************************
' Form_Load
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Load the form, hide non-existant weapons, set weapon names, set ammo,
'         and display totals.
'******************************************************************************
Private Sub Form_Load()
    Dim cbo As ComboBox
    
    B17CPositions = Array(PILOT, COPILOT, BOMBARDIER, ENGINEER, RADIO_OPERATOR, BALL_GUNNER, PORT_WAIST_GUNNER, STBD_WAIST_GUNNER)
    B17CPositionNames = Array("Pilot", "Co-Pilot", "Bombardier", "Engineer", "Radio Operator", "Tunnel Gunner", "Port Waist Gunner", "Stbd. Waist Gunner")
    B17EFGPositions = Array(PILOT, COPILOT, BOMBARDIER, NAVIGATOR, ENGINEER, RADIO_OPERATOR, BALL_GUNNER, PORT_WAIST_GUNNER, STBD_WAIST_GUNNER, TAIL_GUNNER)
    B17EFGPositionNames = Array("Pilot", "Co-Pilot", "Bombardier", "Navigator", "Engineer", "Radio Operator", "Ball Gunner", "Port Waist Gunner", "Stbd. Waist Gunner", "Tail Gunner")
    YB40Positions = Array(PILOT, COPILOT, NAVIGATOR, ENGINEER, RADIO_OPERATOR, NOSE_GUNNER, MID_UPPER_GUNNER, BALL_GUNNER, PORT_WAIST_GUNNER, STBD_WAIST_GUNNER, TAIL_GUNNER, AMMO_STOCKER)
    YB40PositionNames = Array("Pilot", "Co-Pilot", "Navigator", "Engineer", "Radio Operator", "Nose Gunner", "Mid-Upper Gunner", "Ball Gunner", "Port Waist Gunner", "Stbd. Waist Gunner", "Tail Gunner", "Ammo Stocker")
    B24DEPositions = Array(PILOT, COPILOT, BOMBARDIER, NAVIGATOR, ENGINEER, RADIO_OPERATOR, BALL_GUNNER, STBD_WAIST_GUNNER, TAIL_GUNNER)
    B24DEPositionNames = Array("Pilot", "Co-Pilot", "Bombardier", "Navigator", "Engineer", "Radio Operator", "Tunnel Gunner", "Waist Gunner", "Tail Gunner")
    B24GHJPositions = Array(PILOT, COPILOT, BOMBARDIER, NAVIGATOR, ENGINEER, RADIO_OPERATOR, NOSE_GUNNER, BALL_GUNNER, STBD_WAIST_GUNNER, TAIL_GUNNER)
    B24GHJPositionNames = Array("Pilot", "Co-Pilot", "Bombardier", "Navigator", "Engineer", "Radio Operator", "Nose Gunner", "Ball Gunner", "Waist Gunner", "Tail Gunner")
    B24LMPositions = Array(PILOT, COPILOT, BOMBARDIER, NAVIGATOR, ENGINEER, RADIO_OPERATOR, NOSE_GUNNER, BALL_GUNNER, STBD_WAIST_GUNNER, TAIL_GUNNER)
    B24LMPositionNames = Array("Pilot", "Co-Pilot", "Bombardier", "Navigator", "Engineer", "Radio Operator", "Nose Gunner", "Floor Ring Gunner", "Waist Gunner", "Stinger Gunner")
    LancasterPositions = Array(PILOT, BOMBARDIER, NAVIGATOR, ENGINEER, RADIO_OPERATOR, MID_UPPER_GUNNER, TAIL_GUNNER)
    LancasterPositionNames = Array("Pilot", "Bomb Aimer", "Navigator", "Flight Engineer", "Wireless Operator", "Mid-Upper Gunner", "Tail Gunner")
'    CenterForm Me

    ' Fiddle the form bottom, as adding a menu bar otherwise seems to
    ' randomly cut off the bottom of the form
    frmCrewAssignment.Height = cmdOK.Top + cmdOK.Height + 880
    
    ' Point to the record currently on the airman tab.
    
'    varBookmark = prsAirman.Bookmark
    
    intCurrentAirmanOldAssigment = prsAirman![Assignment]

'MsgBox "1: intCurrentAirmanOldAssigment = " & intCurrentAirmanOldAssigment & vbCrLf & _
'       "prsAirman![Assignment] = " & prsAirman![Assignment] & vbCrLf & _
'       "varAirmanCurrentlyOnTab = " & varAirmanCurrentlyOnTab & vbCrLf & _
'       "prsAirman![Name] = " & prsAirman![Name]
    
    For Each cbo In cboCrewPosition
        cbo.Tag = cbo.Index
    Next
    
    With frmMainMenu
    
        Me.Caption = .cmdAssignCrew.Caption & " Dialog " & " (" & .cboName(BOMBER_TAB).Text & ")"
    
        ' Position the combos that will be visible for the bomber. Record
        ' which ones will be hidden.
        
'MsgBox "(.cboBomberModel(BOMBER_TAB).ListIndex + 1) = " & (.cboBomberModel(BOMBER_TAB).ListIndex + 1)

        Select Case (.cboBomberModel(BOMBER_TAB).ListIndex + 1)
            
            Case B17_C:
                CurrentBomberPositions = B17CPositions
                CurrentBomberPositionNames = B17CPositionNames
            
            Case B17_E, B17_F, B17_G:
                CurrentBomberPositions = B17EFGPositions
                CurrentBomberPositionNames = B17EFGPositionNames
            
            Case YB40:
                CurrentBomberPositions = YB40Positions
                CurrentBomberPositionNames = YB40PositionNames
                
            Case B24_D, B24_E:
                CurrentBomberPositions = B24DEPositions
                CurrentBomberPositionNames = B24DEPositionNames
                
            Case B24_GHJ:
                CurrentBomberPositions = B24GHJPositions
                CurrentBomberPositionNames = B24GHJPositionNames
                
            Case B24_LM:
                CurrentBomberPositions = B24LMPositions
                CurrentBomberPositionNames = B24LMPositionNames
                
            Case AVRO_LANCASTER:
                CurrentBomberPositions = LancasterPositions
                CurrentBomberPositionNames = LancasterPositionNames
        End Select
        
        Call PositionCrewCombos
        ' Fill in the text portion of the visible and enabled combos.

'MsgBox "Call FillCrewAssignmentDialogFields"
        If FillCrewAssignmentDialogFields() = False Then
' qwe            Call ExitEmulator
            gblnCrewAssigned = False
            Unload Me
        End If

        ' Populate the combos that will be visible and enabled, otherwise
        ' disable the combos.

        If .cmdAssignCrew.Caption = "Assign Crew" Then
        
            Call PopulateCrewPositionCombos
        
        Else ' Default Crew or Last Crew
        
            Call DisableCrewPositionCombos
            cmdOK.Visible = False
            'CenterControl fraCancelHelp, Me
        
        End If

        ' Hide the combos that don't exist on the bomber.

'MsgBox "Call HideUnusedCombos"
        Call HideUnusedCombos

    End With
End Sub
            
'******************************************************************************
' FillCrewAssignmentDialogFields
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Initialize the crew position drop downs with the names of the airmen
'         currently assigned to the bomber. Do not lookup airmen for positions
'         which do not exist on the bomber (indicated by "BLANK"), or for
'         positions which are unfilled (indicated by 0).
'******************************************************************************
Private Function FillCrewAssignmentDialogFields()
    Dim strIgnore As String
    
    FillCrewAssignmentDialogFields = True

'MsgBox "1 - prsBomber![Name]: " & prsBomber![Name]
    
'MsgBox "a: " & prsBomber![PILOT]
'    prsBomber.AbsolutePosition = 4 ' error here
    If prsBomber![PILOT] <> UNMANNED_POSITION _
    And prsBomber![PILOT] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![PILOT], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(PILOT).Text = prsAirman![Name]
            txtSerialNumber(PILOT).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "b: " & prsBomber![COPILOT]

    If prsBomber![COPILOT] <> UNMANNED_POSITION _
    And prsBomber![COPILOT] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![COPILOT], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(COPILOT).Text = prsAirman![Name]
            txtSerialNumber(COPILOT).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "c: " & prsBomber![BOMBARDIER]
    
    If prsBomber![BOMBARDIER] <> UNMANNED_POSITION _
    And prsBomber![BOMBARDIER] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![BOMBARDIER], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(BOMBARDIER).Text = prsAirman![Name]
            txtSerialNumber(BOMBARDIER).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "d: " & prsBomber![NAVIGATOR]
    
    If prsBomber![NAVIGATOR] <> UNMANNED_POSITION _
    And prsBomber![NAVIGATOR] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![NAVIGATOR], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(NAVIGATOR).Text = prsAirman![Name]
            txtSerialNumber(NAVIGATOR).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "e: " & prsBomber![ENGINEER]
    
    If prsBomber![ENGINEER] <> UNMANNED_POSITION _
    And prsBomber![ENGINEER] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![ENGINEER], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(ENGINEER).Text = prsAirman![Name]
            txtSerialNumber(ENGINEER).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "f: " & prsBomber![RadioOperator]
    
    If prsBomber![RadioOperator] <> UNMANNED_POSITION _
    And prsBomber![RadioOperator] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![RadioOperator], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(RADIO_OPERATOR).Text = prsAirman![Name]
            txtSerialNumber(RADIO_OPERATOR).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "g: " & prsBomber![NoseGunner]
    
    If prsBomber![NoseGunner] <> UNMANNED_POSITION _
    And prsBomber![NoseGunner] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![NoseGunner], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(NOSE_GUNNER).Text = prsAirman![Name]
            txtSerialNumber(NOSE_GUNNER).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "h: " & prsBomber![MidUpperGunner]
    
    If prsBomber![MidUpperGunner] <> UNMANNED_POSITION _
    And prsBomber![MidUpperGunner] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![MidUpperGunner], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(MID_UPPER_GUNNER).Text = prsAirman![Name]
            txtSerialNumber(MID_UPPER_GUNNER).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "i: " & prsBomber![BallGunner]
    
    If prsBomber![BallGunner] <> UNMANNED_POSITION _
    And prsBomber![BallGunner] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![BallGunner], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(BALL_GUNNER).Text = prsAirman![Name]
            txtSerialNumber(BALL_GUNNER).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "j: " & prsBomber![PortWaistGunner]
    
    If prsBomber![PortWaistGunner] <> UNMANNED_POSITION _
    And prsBomber![PortWaistGunner] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![PortWaistGunner], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(PORT_WAIST_GUNNER).Text = prsAirman![Name]
            txtSerialNumber(PORT_WAIST_GUNNER).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "k: " & prsBomber![StbdWaistGunner]
    
    If prsBomber![StbdWaistGunner] <> UNMANNED_POSITION _
    And prsBomber![StbdWaistGunner] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![StbdWaistGunner], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(STBD_WAIST_GUNNER).Text = prsAirman![Name]
            txtSerialNumber(STBD_WAIST_GUNNER).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "l: " & prsBomber![TailGunner]
    
    If prsBomber![TailGunner] <> UNMANNED_POSITION _
    And prsBomber![TailGunner] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![TailGunner], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(TAIL_GUNNER).Text = prsAirman![Name]
            txtSerialNumber(TAIL_GUNNER).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "m: " & prsBomber![AmmoStocker]
    
'    If txtSerialNumber(AMMO_STOCKER).Text <> HIDDEN_POSITION _
'    And prsBomber![AmmoStocker] <> UNMANNED_POSITION Then
    If prsBomber![AmmoStocker] <> UNMANNED_POSITION _
    And prsBomber![AmmoStocker] <> HIDDEN_POSITION Then
        If LookupAirman(prsBomber![AmmoStocker], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillCrewAssignmentDialogFields = False
            Exit Function
        Else
            cboCrewPosition(AMMO_STOCKER).Text = prsAirman![Name]
            txtSerialNumber(AMMO_STOCKER).Text = prsAirman![KeyField]
        End If
    End If

'MsgBox "n"

'MsgBox "2 - prsBomber![Name]: " & prsBomber![Name]
    
End Function

'******************************************************************************
' PopulateCrewPositionCombos
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is tricky ...
'******************************************************************************
Private Sub PopulateCrewPositionCombos()
    Dim frsTemp As New ADODB.Recordset
    Dim intPos As Integer
    Dim intIndex As Integer
    Dim strBaseFilter As String
    Dim strPositionFilter As String
    Dim strFilter As String
    
    Set frsTemp = prsAirman.Clone

    ' Only airmen on duty status, and not default airmen, and not in
    ' flight should be listed in the combos. A further AND clause is
    ' appended prior to filling each position's combo, so that only
    ' airman qualified for that position will be listed. Finally, since
    ' filters are very picky about using both AND and OR clauses, and
    ' because a filter cannot be filtered (the second filter replaces
    ' the first rather than supplementing it), as the filtered recordset
    ' is looped, only the airmen which are on admin duty, or which are
    ' already assigned to the bomber will be added to the combo.
    
    strBaseFilter = "Status = " & DUTY_STATUS & " AND " & _
                    "Default = " & vbUnchecked

    ' At the very least, every visible combo should have a blank first
    ' row. After adding the blank row, fill in the other rows.
    
    For intPos = 1 To txtSerialNumber.UBound
        If txtSerialNumber(intPos).Text <> HIDDEN_POSITION Then
            
            cboCrewPosition(intPos).AddItem ""
            
            strPositionFilter = " AND " & "CrewPosition = " & intPos
        
            strFilter = strBaseFilter & strPositionFilter
        
            frsTemp.Filter = strFilter
                    
            ' Two loops of the temporary recordset are necessary: The
            ' first to get the intIndex dimensioning value, the second
            ' to fill the combo and matrix row for the current position.
            ' Get the dimensioning value.
            
            intIndex = 0

'Msgbox frsTemp.RecordCount
'Msgbox prsAirman.RecordCount
'Msgbox frsTemp.Filter
'            If frsTemp.RecordCount = 0 Then
'
            If frsTemp.RecordCount <> 0 Then
                frsTemp.MoveFirst
                
'MsgBox "3 - prsBomber![Name]: " & prsBomber![Name]
    
                Do Until frsTemp.EOF
                    If frsTemp![Assignment] = ADMIN_DUTY _
                    Or frsTemp![Assignment] = prsBomber![KeyField] Then
                        intIndex = intIndex + 1
                    End If
                        
                    frsTemp.MoveNext
                Loop
            End If
'            End If
        
            ' Dimension the array, then append it to the appropriate row.
            
'Msgbox "intPos = " & intPos & vbCrLf & _
       "intIndex = " & intIndex
            
            If intIndex = 0 Then
                ' There are no nodes for this row. Get the next row.
                GoTo Continue
            End If
            
            ReDim lvntNode(1 To intIndex, 0 To 1)
            
            lvntCrewMatrix(intPos) = lvntNode
            
' MsgBox "Crew Position " & intPos & " = " & LBound(lvntCrewMatrix(intPos), 1) & " ... " & UBound(lvntCrewMatrix(intPos), 1)
' MsgBox "Crew Position " & intPos & " = " & LBound(lvntCrewMatrix(intPos), 2) & " ... " & UBound(lvntCrewMatrix(intPos), 2)
        
            ' Loop the temporary recordset for the second time, filling the
            ' combo and each node on the current row.
            
            intIndex = 0
            frsTemp.MoveFirst
' MsgBox "frsTemp.CursorLocation = " & frsTemp.CursorLocation

            Do Until frsTemp.EOF
' MsgBox "frsTemp.AbsolutePosition = " & frsTemp.AbsolutePosition
                If frsTemp![Assignment] = ADMIN_DUTY _
                Or frsTemp![Assignment] = prsBomber![KeyField] Then
                    cboCrewPosition(intPos).AddItem frsTemp![Name]
                
                    intIndex = intIndex + 1
                
                    lvntCrewMatrix(intPos)(intIndex, 0) = frsTemp![KeyField]
                    lvntCrewMatrix(intPos)(intIndex, 1) = frsTemp![Name]
                
                End If
        
                frsTemp.MoveNext
            Loop

        End If
    
Continue:
    
    Next intPos
    
    Call FreeRecordset(frsTemp)
    
'    If Not frsTemp Is Nothing Then
'        If frsTemp.State = adStateClosed Then frsTemp.Close
'        Set frsTemp = Nothing
'    End If

' debug
'For intPos = 1 To txtSerialNumber.UBound
'MsgBox "intPos = " & intPos
'    If txtSerialNumber(intPos).Text <> HIDDEN_POSITION Then
'        For intIndex = 1 To UBound(lvntCrewMatrix(intPos), 1)
'MsgBox "intIndex = " & intIndex
'            ' MsgBox "Crew Position = " & intPos & vbCrLf & _
'                   "Airman = " & intIndex & vbCrLf & _
'                   "Serial = " & lvntCrewMatrix(intPos)(intIndex, 0) & vbCrLf & _
'                   "Name = '" & lvntCrewMatrix(intPos)(intIndex, 1) & "'"
'        Next intIndex
'    End If
'Next intPos
' debug

'MsgBox "4 - prsBomber![Name]: " & prsBomber![Name]

End Sub

'******************************************************************************
' DisableCrewPositionCombos
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is done when the bomber is a default plane or it has been shot
'         down, retired, or is otherwise no longer eligible to perform missions.
'******************************************************************************
Private Sub DisableCrewPositionCombos()
    Dim intPos As Integer
    
    For intPos = 1 To cboCrewPosition.UBound
        If txtSerialNumber(intPos).Text <> HIDDEN_POSITION Then
            cboCrewPosition(intPos).Enabled = False
            cboCrewPosition(intPos).BackColor = vbButtonFace
        End If
    Next intPos
    
End Sub

'******************************************************************************
' HideUnusedCombos
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub HideUnusedCombos()
    Dim intPos As Integer
    
    For intPos = 1 To cboCrewPosition.UBound
        If txtSerialNumber(intPos).Text = HIDDEN_POSITION Then
            lblCrewPosition(intPos).Visible = False
            cboCrewPosition(intPos).Visible = False
        ElseIf intPos = PORT_WAIST_GUNNER _
        And (prsBomber![BomberModel] = B24_D _
        Or prsBomber![BomberModel] = B24_E _
        Or prsBomber![BomberModel] = B24_GHJ _
        Or prsBomber![BomberModel] = B24_LM) Then
            lblCrewPosition(intPos).Visible = False
            cboCrewPosition(intPos).Visible = False
        End If
    Next intPos
    
End Sub

'******************************************************************************
' cboCrewPosition_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  We need to track each airman's KeyField so that we can update his
'         assigned bomber, and the bomber's crew list. If there was just one
'         airman combo, we wouldn't need a placeholder field, but with a
'         variable multitude of airman combos, all scrolling through the
'         airman recordset, we need to be more certain of which airmen we
'         are dealing with.
'******************************************************************************
Private Sub cboCrewPosition_Click(intPos As Integer)

    If cboCrewPosition(intPos).ListIndex = 0 Then
        ' Blank row. If this is a required position, the bomber will not be
        ' able to fly missions until the position is filled.
        txtSerialNumber(intPos).Text = UNMANNED_POSITION
        Exit Sub
    End If

' MsgBox "Store " & lvntCrewMatrix(intPos)(cboCrewPosition(intPos).ListIndex, 0) & " in txtSerialNumber."

    ' Row is not blank. Get the airman's serial number. The combo row and
    ' matrix node should be equal. Therefore use ListIndex to index to
    ' the associated node.
    
    txtSerialNumber(intPos).Text = lvntCrewMatrix(intPos)(cboCrewPosition(intPos).ListIndex, 0)

End Sub

' qwe
Private Sub ExitCrewAssign()
Dim strAssignment As String
    ' Re-point the recordset to the record currently on the airman tab.

    prsAirman.Bookmark = varAirmanCurrentlyOnTab

    ' Ensure the current airman's assignment reflects any update that just
    ' occured.

    If prsAirman![Assignment] <> intCurrentAirmanOldAssigment Then

        If LookupBomber(prsAirman![Assignment], LOOKUP_BY_KEYFIELD, strAssignment) = False Then
MsgBox "' TODO: Bail completely???"
            gblnCrewAssigned = False
'            frmMainMenu.cboAssignment.Text = strAssignment
            Exit Sub
        Else
            frmMainMenu.cboAssignment.Text = strAssignment
        End If
        
        prsBomber.Bookmark = varBomberCurrentlyOnTab

    End If

Call AdjustAvailableBombers ' Nov04
'    Call AdjustMissionAvailableBombers
    prsBomber.Bookmark = varBomberCurrentlyOnTab
    
    Unload Me

End Sub

'******************************************************************************
' cmdOK_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Validate data, then commit assignments.
'******************************************************************************
Private Sub cmdOK_Click()
    Dim strAssignment As String

 'MsgBox "cmdOK_Click()"

    If ValidData() = False Then
        Exit Sub
    End If

    If CommitAssignments() = False Then
' qwe            Call ExitEmulator
        gblnCrewAssigned = False
'        Unload Me
    End If

'MsgBox "4: intCurrentAirmanOldAssigment = " & intCurrentAirmanOldAssigment & vbCrLf & _
       "prsAirman![Assignment] = " & prsAirman![Assignment] & vbCrLf & _
       "varAirmanCurrentlyOnTab = " & varAirmanCurrentlyOnTab & vbCrLf & _
       "prsAirman![Name] = " & prsAirman![Name]
    
    ' Re-point the recordset to the record currently on the airman tab.

    prsAirman.Bookmark = varAirmanCurrentlyOnTab

'MsgBox "5: intCurrentAirmanOldAssigment = " & intCurrentAirmanOldAssigment & vbCrLf & _
       "prsAirman![Assignment] = " & prsAirman![Assignment] & vbCrLf & _
       "varAirmanCurrentlyOnTab = " & varAirmanCurrentlyOnTab & vbCrLf & _
       "prsAirman![Name] = " & prsAirman![Name]
    
    ' Ensure the current airman's assignment reflects any update that just
    ' occured.

    If prsAirman![Assignment] <> intCurrentAirmanOldAssigment Then

'MsgBox "Reset the assignment combo"
        If LookupBomber(prsAirman![Assignment], LOOKUP_BY_KEYFIELD, strAssignment) = False Then
MsgBox "' TODO: Bail completely???"
            With frmMainMenu
                .cboAssignment.Text = strAssignment
            End With
'            FillAirmanTabFields = False
            gblnCrewAssigned = False ' qwe
            Exit Sub
        
        Else
            With frmMainMenu
'MsgBox "strAssignment = '" & strAssignment & "'"
                .cboAssignment.Text = strAssignment
            End With
        End If
'MsgBox "prsBomber.Bookmark = varBomberCurrentlyOnTab"
        
        prsBomber.Bookmark = varBomberCurrentlyOnTab

    End If

Call AdjustAvailableBombers ' Nov04
'    Call AdjustMissionAvailableBombers
    prsBomber.Bookmark = varBomberCurrentlyOnTab
    
    Unload Me

End Sub

'******************************************************************************
' CommitAssignments
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  No controls or checks for non-Flight Duty status bombers or crew are
'         necessary, as we would not be able to reach this function if were not
'         able to perform updates.
'******************************************************************************
Private Function CommitAssignments() As Boolean
'    On Error GoTo ErrorTrap

'MsgBox "CommitAssignments()"

    Dim strAirman As String
    Dim intPositionsOnBomber As Integer
    Dim intIndex As Integer
    Dim blnStandDown As Boolean
    Dim strBomberStatus As String

'MsgBox "5 - prsBomber![Name]: " & prsBomber![Name]

    CommitAssignments = True
    
    ' Even if the bomber was on stand down status when the dialog was brought
    ' up, we assume that the empty position(s) was filled.
    blnStandDown = False
    
'MsgBox "2: intCurrentAirmanOldAssigment = " & intCurrentAirmanOldAssigment & vbCrLf & _
       "varAirmanCurrentlyOnTab = " & varAirmanCurrentlyOnTab & vbCrLf & _
       "prsAirman![Name] = " & prsAirman![Name]

    pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
        
        ' Determine how many positions the bomber currently has filled.
        ' The result will be used as a control to determine when we no
        ' longer search prsAirman for assignments to wipe.
        
        intPositionsOnBomber = 0
        
        For intIndex = 1 To txtSerialNumber.UBound

            If intIndex = PORT_WAIST_GUNNER _
            And (prsBomber![BomberModel] = B24_D _
            Or prsBomber![BomberModel] = B24_E _
            Or prsBomber![BomberModel] = B24_GHJ _
            Or prsBomber![BomberModel] = B24_LM) Then
            
                ' The port waist gunner is supposed to be unmanned on B-24s.
            
            ElseIf txtSerialNumber(intIndex).Text = UNMANNED_POSITION Then

                blnStandDown = True
                intPositionsOnBomber = intPositionsOnBomber + 1
                
            ElseIf txtSerialNumber(intIndex).Text <> HIDDEN_POSITION Then
                intPositionsOnBomber = intPositionsOnBomber + 1
            End If

        Next intIndex

'MsgBox "intPositionsOnBomber = " & intPositionsOnBomber
        
        If blnStandDown = True Then
            prsBomber![Status] = STAND_DOWN_STATUS
        Else
            prsBomber![Status] = DUTY_STATUS
        End If

        If LookupBomberStatus(prsBomber![Status], strBomberStatus) = False Then
            CommitAssignments = False
' qwe            Exit Function ' TODO: error out instead?
            GoTo CleanUp
        Else
'MsgBox "frmMainMenu.txtStatus(BOMBER_TAB).Text = " & strBomberStatus
            frmMainMenu.txtStatus(BOMBER_TAB).Text = strBomberStatus
        End If
'MsgBox "intPositionsOnBomber = " & intPositionsOnBomber
        
        ' Update the bomber's modified crew.
        
        prsBomber![PILOT] = txtSerialNumber(PILOT).Text
        prsBomber![COPILOT] = txtSerialNumber(COPILOT).Text
        prsBomber![BOMBARDIER] = txtSerialNumber(BOMBARDIER).Text
        prsBomber![NAVIGATOR] = txtSerialNumber(NAVIGATOR).Text
        prsBomber![ENGINEER] = txtSerialNumber(ENGINEER).Text
        prsBomber![RadioOperator] = txtSerialNumber(RADIO_OPERATOR).Text
        prsBomber![NoseGunner] = txtSerialNumber(NOSE_GUNNER).Text
        prsBomber![MidUpperGunner] = txtSerialNumber(MID_UPPER_GUNNER).Text
        prsBomber![BallGunner] = txtSerialNumber(BALL_GUNNER).Text
        prsBomber![PortWaistGunner] = txtSerialNumber(PORT_WAIST_GUNNER).Text
        prsBomber![StbdWaistGunner] = txtSerialNumber(STBD_WAIST_GUNNER).Text
        prsBomber![TailGunner] = txtSerialNumber(TAIL_GUNNER).Text
        prsBomber![AmmoStocker] = txtSerialNumber(AMMO_STOCKER).Text
            
        ' Update the previous crew's assignments. Rather than record who
        ' the current crew is, then do a bunch of lookups and updates on
        ' only the airman who've been transferred off the bomber, we simply
        ' wipe the assignments of all assigned airmen, then rewrite them.
        
        intIndex = 1
    
'MsgBox "Loop prsAirman, wiping previous crew assignments"
        
        prsAirman.MoveFirst
        Do Until prsAirman.EOF

'MsgBox "prsAirman.CursorLocation = " & prsAirman.CursorLocation & vbCrLf & _
       "intIndex = " & intIndex & vbCrLf & _
       "intPositionsOnBomber = " & intPositionsOnBomber
            
            If intIndex > intPositionsOnBomber Then

'MsgBox "Even though there might be more airmen in the recordset, we've already wiped the assignments of all previous crew."
                
                ' Even though there might be more airmen in the recordset,
                ' we've already wiped the assignments of all previous crew.
                Exit Do
            End If
            
            If prsAirman![Assignment] = prsBomber![KeyField] Then

'MsgBox "Airman " & prsAirman![Name] & "has been assigned to 0."
                
                prsAirman![Assignment] = ADMIN_DUTY
                intIndex = intIndex + 1
            End If

            prsAirman.MoveNext
        Loop
        
'MsgBox "Update the modified crew's assignments and positions."
        
        ' Update the modified crew's assignments and positions. UBound is
        ' the number of hidden serial number textboxes on the form, where
        ' each textbox is associated with one crew position combo.
        
        For intIndex = 1 To txtSerialNumber.UBound
            If txtSerialNumber(intIndex).Text = UNMANNED_POSITION _
            Or txtSerialNumber(intIndex).Text = HIDDEN_POSITION Then
'MsgBox "empty or hidden, get next airman's serial number"
                GoTo Continue
            End If
        
'MsgBox "Key to the airman occupying the position in question."
            
            ' Key to the airman occupying the position in question.

            If LookupAirman(txtSerialNumber(intIndex).Text, LOOKUP_BY_KEYFIELD, strAirman) = False Then
'MsgBox "airman not found"
                CommitAssignments = False
' qwe                Exit Function ' TODO: error out instead?
                GoTo CleanUp
            Else
'MsgBox "The airman's assignment was previously wiped. Now, it will be set to the current bomber."
                ' The airman's assignment was previously wiped. Now, it will
                ' be set to the current bomber.
                prsAirman![Assignment] = prsBomber![KeyField]

            End If
        
Continue:
        
        Next intIndex
            
'MsgBox "BEF: prsBomber![Name] = " & prsBomber![Name]
        prsBomber.UpdateBatch
        prsAirman.UpdateBatch
'MsgBox "AFT: prsBomber![Name] = " & prsBomber![Name]

    pobjConn.CommitTrans
        
    pintOpenTrans = pintOpenTrans - 1
        
'MsgBox "3: intCurrentAirmanOldAssigment = " & intCurrentAirmanOldAssigment & vbCrLf & _
       "varAirmanCurrentlyOnTab = " & varAirmanCurrentlyOnTab & vbCrLf & _
       "prsAirman![Name] = " & prsAirman![Name]
    
    Exit Function

CleanUp:

    If pintOpenTrans Then
        pobjConn.RollbackTrans
        pintOpenTrans = pintOpenTrans - 1
    End If
    
'MsgBox "6 - prsBomber![Name]: " & prsBomber![Name]

    Exit Function

ErrorTrap:

    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "CommitAssignments() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)

    Err.Clear

    CommitAssignments = False
    
    Resume CleanUp

End Function

'******************************************************************************
' ValidData
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  An airman may not occupy more than one crew position on a given
'         bomber. An airman may not be on the crew of more than one active
'         bomber. An airman may be on the crew of an unlimited number of
'         retired bombers. Blank crew positions are permissible; they should
'         not give a duplicate airman error. A bomber with blank crew positions
'         is not allowed to fly missions.
'******************************************************************************
Private Function ValidData()
    Dim intIndex As Integer
    Dim blnMissingCrew As Boolean
    
    ValidData = True
    blnMissingCrew = False

    For intIndex = 1 To txtSerialNumber.UBound
        ' Duplicate crew positions are illegal. Blank crew positions are
        ' permissible. Multiple blank crew positions should not give the
        ' duplicate error.
        
        If intIndex = PORT_WAIST_GUNNER _
        And (prsBomber![BomberModel] = B24_D _
        Or prsBomber![BomberModel] = B24_E _
        Or prsBomber![BomberModel] = B24_GHJ _
        Or prsBomber![BomberModel] = B24_LM) Then
        
            ' The port waist gunner is supposed to be unmanned on B-24s.
        
        ElseIf txtSerialNumber(intIndex).Text = UNMANNED_POSITION Then

            ' The position really is blank / unoccupied.
            blnMissingCrew = True
        
        End If
    
    Next intIndex
    
    If blnMissingCrew = True Then
        strErrMsg = "One or more crew positions are unmanned. The bomber " & _
                    "will not be able to fly a mission until all positions " & _
                    "are manned."

        MsgBox strErrMsg, (vbInformation + vbOKOnly)
    End If

End Function

Private Sub PositionCrewCombos()
'Position and label bomber crew comboboxes
    Dim X  As Integer
    Dim Y As Integer
    Dim i As Integer
    Dim j As Integer
    Dim hasPosition As Boolean
    Dim firstColumn As Boolean
    Dim bomberPositions As Integer
    
    X = 120
    Y = 120
    firstColumn = True
    
    For i = cboCrewPosition.LBound To cboCrewPosition.UBound
        hasPosition = False
        For j = LBound(CurrentBomberPositions) To UBound(CurrentBomberPositions)
            If CurrentBomberPositions(j) = i Then
                hasPosition = True
                lblCrewPosition(i).Caption = CurrentBomberPositionNames(j)
                bomberPositions = bomberPositions + 1
            End If
        Next
        If hasPosition Then
            lblCrewPosition(i).Left = X
            lblCrewPosition(i).Top = Y
            Y = Y + 360
            cboCrewPosition(i).Left = X
            cboCrewPosition(i).Top = Y
            Y = Y + 600
            txtSerialNumber(i).Text = UNMANNED_POSITION
        Else
            txtSerialNumber(i).Text = HIDDEN_POSITION
        End If
        If firstColumn And bomberPositions > (UBound(CurrentBomberPositions) - LBound(CurrentBomberPositions)) / 2 Then
            X = 2760
            Y = 120
            firstColumn = False
        End If
    Next
End Sub

'******************************************************************************
' cmdCancel_Click
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Cancel the changes.
'******************************************************************************
Private Sub cmdCancel_Click()

    Unload Me

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
'    frmHelpBrowser.Hide
    
    frmHelpBrowser.txtPageName.Text = "doc/B17CrewAssignmentHelp.html"
    
    frmHelpBrowser.Show vbModal
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
'    frmHelpBrowser.Hide
    
    frmAbout.Show vbModal
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
'    frmHelpBrowser.Hide
    
    frmHelpBrowser.txtPageName.Text = "doc/B17HelpIndex.html"

    frmHelpBrowser.Show vbModal
End Sub

