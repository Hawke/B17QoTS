'******************************************************************************
' frmSwapAmmo.frm
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
VERSION 5.00
Begin VB.Form frmSwapAmmo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Swap Ammo"
   ClientHeight    =   5040
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   2520
      TabIndex        =   43
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtBombBayAmmo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   41
      Top             =   4080
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txtMaxBombBayAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   40
      Top             =   4080
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txtMaxTotal 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   37
      Top             =   4560
      Width           =   360
   End
   Begin VB.TextBox txtMaxTotal 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   36
      Top             =   4200
      Width           =   360
   End
   Begin VB.TextBox txtMaxAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   720
      TabIndex        =   35
      Top             =   3360
      Width           =   360
   End
   Begin VB.TextBox txtMaxAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   720
      TabIndex        =   34
      Top             =   3000
      Width           =   360
   End
   Begin VB.TextBox txtMaxAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   720
      TabIndex        =   33
      Top             =   2640
      Width           =   360
   End
   Begin VB.TextBox txtMaxAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   720
      TabIndex        =   32
      Top             =   2280
      Width           =   360
   End
   Begin VB.TextBox txtMaxAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   720
      TabIndex        =   31
      Top             =   1920
      Width           =   360
   End
   Begin VB.TextBox txtMaxAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   720
      TabIndex        =   30
      Top             =   1560
      Width           =   360
   End
   Begin VB.TextBox txtMaxAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   29
      Top             =   1200
      Width           =   360
   End
   Begin VB.TextBox txtMaxAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   28
      Top             =   840
      Width           =   360
   End
   Begin VB.TextBox txtMaxAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   27
      Top             =   480
      Width           =   360
   End
   Begin VB.TextBox txtMaxAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   720
      TabIndex        =   26
      Top             =   3720
      Width           =   360
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   4560
      Width           =   360
   End
   Begin VB.TextBox txtGunAmmo 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   360
   End
   Begin VB.TextBox txtGunAmmo 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   360
   End
   Begin VB.TextBox txtGunAmmo 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   360
   End
   Begin VB.TextBox txtGunAmmo 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   360
   End
   Begin VB.TextBox txtGunAmmo 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   360
   End
   Begin VB.TextBox txtGunAmmo 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   360
   End
   Begin VB.TextBox txtGunAmmo 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   360
   End
   Begin VB.TextBox txtGunAmmo 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   360
   End
   Begin VB.TextBox txtGunAmmo 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   360
   End
   Begin VB.TextBox txtGunAmmo 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   360
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   360
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblBombBayAmmo 
      Caption         =   "Bomb Bay Ammo"
      Height          =   255
      Left            =   1200
      TabIndex        =   42
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Max"
      Height          =   255
      Left            =   720
      TabIndex        =   39
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Ammo"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblTotal 
      Caption         =   "Twin Gun Total"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   25
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblGunName 
      Caption         =   "Mid-Upper"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   23
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblGunName 
      Caption         =   "Tail Turret"
      Height          =   255
      Index           =   10
      Left            =   1200
      TabIndex        =   22
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblGunName 
      Caption         =   "Ball Turret"
      Height          =   255
      Index           =   9
      Left            =   1200
      TabIndex        =   21
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblGunName 
      AutoSize        =   -1  'True
      Caption         =   "Top Turret"
      Height          =   255
      Index           =   8
      Left            =   1200
      TabIndex        =   20
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblGunName 
      Caption         =   "Stbd Waist"
      Height          =   255
      Index           =   7
      Left            =   1200
      TabIndex        =   19
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblGunName 
      Caption         =   "Port Waist"
      Height          =   255
      Index           =   6
      Left            =   1200
      TabIndex        =   18
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblGunName 
      Caption         =   "Radio Room"
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblGunName 
      Caption         =   "Stbd Cheek"
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   16
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblGunName 
      Caption         =   "Port Cheek"
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblGunName 
      Caption         =   "Nose"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   14
      Top             =   840
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2520
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblTotal 
      Caption         =   "Single Gun Total"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   13
      Top             =   4200
      Width           =   1335
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
Attribute VB_Name = "frmSwapAmmo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gintTotal() As Integer ' Values remain unchanged once it is set on load
Private gintMaxTotal() As Integer ' Values remain unchanged once it is set on load
Private gintMaxAmmoPts As Integer
Private gintOldAmmo As Integer

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
    Dim intGun As Integer

    ReDim gintTotal(SINGLE_GUN_AMMO To TWIN_GUN_AMMO)
    ReDim gintMaxTotal(SINGLE_GUN_AMMO To TWIN_GUN_AMMO)
            
    If Bomber.ExtraAmmo >= 1 Then
        gintMaxAmmoPts = Bomber.ExtraAmmo
    Else
        gintMaxAmmoPts = 0
    End If
    
    For intGun = MID_UPPER_MG To TAIL_MG
        
        If GunExists(intGun) = False Then
        
            txtGunAmmo(intGun).Visible = False
            txtMaxAmmo(intGun).Visible = False
            lblGunName(intGun).Visible = False
        
        Else
            
            txtGunAmmo(intGun).Text = Bomber.Gun(intGun).Ammo
            
            ' The gun captions are defaulted to their most common
            ' value, however the caption may vary depending on bomber model.
            
            Select Case Bomber.BomberModel
                
                Case B17_C:
                    
                    If intGun = BALL_TURRET_MG Then
                        lblGunName(intGun).Caption = "Tunnel"
                    End If
                
                Case YB40:
                    
                    If intGun = MID_UPPER_MG Then
                        lblGunName(intGun).Caption = "Mid-Upper"
                    End If
                
                Case B24_D:
                    
                    If intGun = BALL_TURRET_MG Then
                        lblGunName(intGun).Caption = "Tunnel"
                    End If
                
                Case B24_E:
                    
                    If intGun = BALL_TURRET_MG Then
                        lblGunName(intGun).Caption = "Tunnel"
                    End If
                
                Case B24_LM:
                    
                    If intGun = BALL_TURRET_MG Then
                        lblGunName(intGun).Caption = "Floor Ring"
                    ElseIf intGun = TAIL_MG Then
                        lblGunName(intGun).Caption = "Stinger"
                    End If
                
                Case AVRO_LANCASTER:
                    
                    If intGun = TAIL_MG Then
                        lblGunName(intGun).Caption = "Rear"
                    End If
                
            End Select
                    
            txtMaxAmmo(intGun).Text = Bomber.Gun(intGun).MaxAmmo
            
            ' Twin .50's in cyan, single guns in default white. The amount of
            ' each type of ammo is fixed: It cannot be raised or lowered.
                
            If Bomber.ExtraAmmo >= 1 Then
                
                If Bomber.Gun(intGun).Bonus = 0 Then
                    gintTotal(SINGLE_GUN_AMMO) = gintTotal(SINGLE_GUN_AMMO) + Bomber.Gun(intGun).Ammo
                    gintMaxTotal(SINGLE_GUN_AMMO) = gintMaxTotal(SINGLE_GUN_AMMO) + Bomber.Gun(intGun).MaxAmmo
                    gintMaxAmmoPts = gintMaxAmmoPts + Bomber.Gun(intGun).Ammo
                Else
                    txtGunAmmo(intGun).BackColor = PaleCyan()
                    gintTotal(TWIN_GUN_AMMO) = gintTotal(TWIN_GUN_AMMO) + Bomber.Gun(intGun).Ammo
                    gintMaxTotal(TWIN_GUN_AMMO) = gintMaxTotal(TWIN_GUN_AMMO) + Bomber.Gun(intGun).MaxAmmo
                    gintMaxAmmoPts = gintMaxAmmoPts + (Bomber.Gun(intGun).Ammo * 2)
                End If
            
            Else
                
                If Bomber.Gun(intGun).Bonus = 0 Then
                    gintTotal(SINGLE_GUN_AMMO) = gintTotal(SINGLE_GUN_AMMO) + Bomber.Gun(intGun).Ammo
                Else
                    txtGunAmmo(intGun).BackColor = PaleCyan()
                    gintTotal(TWIN_GUN_AMMO) = gintTotal(TWIN_GUN_AMMO) + Bomber.Gun(intGun).Ammo
                End If
            
            End If
            
        End If
    
    Next intGun

    ' Display the current totals
    
    txtTotal(SINGLE_GUN_AMMO).Text = gintTotal(SINGLE_GUN_AMMO)
    txtTotal(TWIN_GUN_AMMO).Text = gintTotal(TWIN_GUN_AMMO)
    
    ' Display the permanent/required totals
    
    If Bomber.ExtraAmmo >= 1 Then
        txtMaxTotal(SINGLE_GUN_AMMO).Text = gintMaxTotal(SINGLE_GUN_AMMO)
        txtMaxTotal(TWIN_GUN_AMMO).Text = gintMaxTotal(TWIN_GUN_AMMO)
        Call DisplayBombBayAmmo
    Else
        txtMaxTotal(SINGLE_GUN_AMMO).Text = gintTotal(SINGLE_GUN_AMMO)
        txtMaxTotal(TWIN_GUN_AMMO).Text = gintTotal(TWIN_GUN_AMMO)
    End If

    ' Fiddle the form bottom, as adding a menu bar otherwise seems to
    ' randomly cut off the bottom of the form
    frmCrewAssignment.Height = txtTotal(SINGLE_GUN_AMMO).Top + txtTotal(SINGLE_GUN_AMMO).Height + 440
    
End Sub

'******************************************************************************
' DisplayBombBayAmmo
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Insert extra ammo in bomb bay prior to totals, adjusting totals
'         controls downward.
'******************************************************************************
Private Sub DisplayBombBayAmmo()
    Me.Height = Me.Height + 360
    Line1.Y1 = Line1.Y1 + 360
    Line1.Y2 = Line1.Y2 + 360
    txtTotal(1).Top = txtTotal(1).Top + 360
    txtMaxTotal(1).Top = txtMaxTotal(1).Top + 360
    lblTotal(1).Top = lblTotal(1).Top + 360
    txtTotal(2).Top = txtTotal(2).Top + 360
    txtMaxTotal(2).Top = txtMaxTotal(2).Top + 360
    lblTotal(2).Top = lblTotal(2).Top + 360
    txtBombBayAmmo.Text = Bomber.ExtraAmmo
    txtMaxBombBayAmmo.Text = Bomber.ExtraAmmo
    txtBombBayAmmo.Visible = True
    txtMaxBombBayAmmo.Visible = True
    lblBombBayAmmo.Visible = True
End Sub

'******************************************************************************
' txtGunAmmo_GotFocus
'
' INPUT:  Index to specific weapon.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub txtGunAmmo_GotFocus(Index As Integer)
    ' Save this value before the user is able to change it.
    gintOldAmmo = CInt(txtGunAmmo(Index).Text)
End Sub

'******************************************************************************
' txtBombBayAmmo_LostFocus
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub txtBombBayAmmo_LostFocus()
    On Error GoTo ErrorTrap
    
    txtBombBayAmmo.Text = CInt(txtBombBayAmmo.Text)

    If CInt(txtBombBayAmmo.Text) > CInt(txtMaxBombBayAmmo.Text) Then
        txtBombBayAmmo.Text = txtMaxBombBayAmmo.Text
    End If
    
    Exit Sub
   
ErrorTrap:
    
    txtBombBayAmmo.Text = "0"
    
    Resume Next

End Sub

'******************************************************************************
' txtGunAmmo_LostFocus
'
' INPUT:  Index to specific weapon.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub txtGunAmmo_LostFocus(intGun As Integer)
    On Error GoTo ErrorTrap
    
    Dim intDiff As Integer
    Dim intIndex As Integer
    Dim intNewAmmo As Integer
    Dim intOldTotal As Integer

    If Bomber.Gun(intGun).Bonus = 0 Then
        intIndex = SINGLE_GUN_AMMO
    Else
        intIndex = TWIN_GUN_AMMO
    End If
    
    intOldTotal = CInt(txtTotal(intIndex).Text)

    intNewAmmo = CInt(txtGunAmmo(intGun).Text)
    
    If intNewAmmo > Bomber.Gun(intGun).MaxAmmo Then
        intNewAmmo = Bomber.Gun(intGun).MaxAmmo
        txtGunAmmo(intGun).Text = intNewAmmo
    End If
    
    intDiff = intNewAmmo - gintOldAmmo
    
    txtTotal(intIndex).Text = intOldTotal + intDiff
    
    Exit Sub
   
ErrorTrap:
    
    txtGunAmmo(intGun).Text = "0"
    intNewAmmo = 0
    
    Resume Next

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
' NOTES:  Accept the changes.
'******************************************************************************
Private Sub cmdOK_Click()
    Dim strMessage As String
    Dim intGun As Integer
    Dim intAmmoPts As Integer
    
    ' Validate ammo swap.
    
    If Bomber.ExtraAmmo >= 1 Then
        
        intAmmoPts = CInt(txtTotal(1).Text) + _
                    (CInt(txtTotal(2).Text) * 2) + _
                    CInt(txtBombBayAmmo.Text)
        
        If CInt(txtBombBayAmmo.Text) > CInt(txtMaxBombBayAmmo.Text) Then
    
            strMessage = "Bomb bay ammo (" & txtBombBayAmmo.Text & ") may not total " & _
                         "more than " & txtMaxBombBayAmmo.Text & "."
        
            MsgBox strMessage, (vbOKOnly + vbInformation)
        
            Exit Sub
        
        ElseIf intAmmoPts <> gintMaxAmmoPts Then
        
            strMessage = "Total ammo points (" & intAmmoPts & ") -- from " & _
                         "bomb bay, single guns and double guns -- must " & _
                         "equal " & gintMaxAmmoPts & "."
        
            MsgBox strMessage, (vbOKOnly + vbInformation)
        
            Exit Sub
        
        ElseIf gintMaxTotal(SINGLE_GUN_AMMO) > CInt(txtMaxTotal(SINGLE_GUN_AMMO).Text) Then
        
            strMessage = "Single gun ammo (" & txtTotal(SINGLE_GUN_AMMO).Text & ") may not exceed " & _
                         gintTotal(SINGLE_GUN_AMMO) & "."
        
            MsgBox strMessage, (vbOKOnly + vbInformation)
        
            Exit Sub
        
        ElseIf gintMaxTotal(TWIN_GUN_AMMO) <> CInt(txtMaxTotal(TWIN_GUN_AMMO).Text) Then
        
            strMessage = "Twin gun ammo (" & txtTotal(TWIN_GUN_AMMO).Text & ") may not exceed " & _
                         gintTotal(TWIN_GUN_AMMO) & "."
        
            MsgBox strMessage, (vbOKOnly + vbInformation)
            
            Exit Sub
        
        End If
    
        Bomber.ExtraAmmo = CInt(txtBombBayAmmo.Text)

    Else
    
        If gintTotal(SINGLE_GUN_AMMO) <> CInt(txtTotal(SINGLE_GUN_AMMO).Text) Then
        
            strMessage = "Single gun ammo (" & txtTotal(SINGLE_GUN_AMMO).Text & ") must total " & _
                         gintTotal(SINGLE_GUN_AMMO) & "."
        
            MsgBox strMessage, (vbOKOnly + vbInformation)
        
            Exit Sub
        
        ElseIf gintTotal(TWIN_GUN_AMMO) <> CInt(txtTotal(TWIN_GUN_AMMO).Text) Then
        
            strMessage = "Twin gun ammo (" & txtTotal(TWIN_GUN_AMMO).Text & ") must total " & _
                         gintTotal(TWIN_GUN_AMMO) & "."
        
            MsgBox strMessage, (vbOKOnly + vbInformation)
            
            Exit Sub
        
        End If
    
    End If

    For intGun = MID_UPPER_GUNNER To TAIL_MG

        If CInt(txtGunAmmo(intGun).Text) > Bomber.Gun(intGun).MaxAmmo Then
            strMessage = "The " & lblGunName(intGun).Caption & " gun may " & _
                         "not have more than " & Bomber.Gun(intGun).MaxAmmo & _
                         " ammo."
            
            MsgBox strMessage, (vbOKOnly + vbInformation)
            
            txtGunAmmo(intGun).SetFocus
            Call SelectField(txtGunAmmo(intGun))
            
            Exit Sub
        
        End If

    Next intGun
    
    ' Swap was valid: The total of both types of ammo is the same as before.
    ' Assign the modified totals back to the bomber.
    
    For intGun = MID_UPPER_MG To TAIL_MG
        
        If GunExists(intGun) = True Then
            Bomber.Gun(intGun).Ammo = CInt(txtGunAmmo(intGun).Text)
        End If

    Next intGun

    Unload Me

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
    frmHelpBrowser.txtPageName.Text = "doc/B17" & Replace(Me.Caption, " ", "") & "Help.html"
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
    frmHelpBrowser.txtPageName.Text = "doc/B17HelpIndex.html"

    frmHelpBrowser.Show vbModal
End Sub


