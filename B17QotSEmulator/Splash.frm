'******************************************************************************
' Splash.frm
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
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   """B-17 Queen of the Skies"" Emulator"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSplash 
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Sub Form_Load()
    
    Dim ret As Long

    ret = mciSendString("OPEN sound/great-escape.mp3 Alias Sonido", 0, 0, 0)
    ret = mciSendString("Play sonido", 0, 0, 0)
            
    ' Pause for dramatic effect. Display, approximately, on first drum beat.
    Sleep 6000
    
    frmSplash.Width = 12090
    frmSplash.Height = 10417
    
    picSplash.Width = 12090
    picSplash.Height = 10417
    picSplash.Picture = LoadPicture(App.Path + "\image\SplashLg.jpg")

End Sub

Private Sub picSplash_Click()
    Dim ret As Long
    
    Load frmMainMenu
    frmMainMenu.Show
    ret = mciSendString("CLOSE Sonido", 0, 0, 0)
    Me.Hide
End Sub
