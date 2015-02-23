'******************************************************************************
' modProgressBar.bas
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

Attribute VB_Name = "modProgressBar"
Option Explicit

' Progress Bar Color, Interface Definition
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_USER = &H400
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const SB_SETBKCOLOR = CCM_SETBKCOLOR

'******************************************************************************
' SetBackColor
'
' INPUT:  Pointer to control and the color it should be changed to.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  API call to change the background color of a progress bar control.
'******************************************************************************
Public Sub SetBackColor(ProgressBarHwnd As Long, RGBValue As Long)
    Call SendMessage(ProgressBarHwnd, SB_SETBKCOLOR, 0, _
      ByVal RGBValue)
End Sub
 
'******************************************************************************
' SetBarColor
'
' INPUT:  Pointer to control and the color it should be changed to.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  API call to change the foreground color of a progress bar control.
'******************************************************************************
Public Sub SetBarColor(ProgressBarHwnd As Long, RGBValue As Long)
    Call SendMessage(ProgressBarHwnd, PBM_SETBARCOLOR, 0, _
        ByVal RGBValue)
End Sub



