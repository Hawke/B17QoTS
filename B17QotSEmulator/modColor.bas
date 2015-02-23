'******************************************************************************
' modColor.bas
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

Attribute VB_Name = "modColor"
Option Explicit

'******************************************************************************
' Color Functions
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Whole number representing an RGB color value.
'
' NOTES:  These functions provide mnemonic wrappers to the RGB function.
'******************************************************************************
Public Function PaleYellow() As Long
    PaleYellow = RGB(255, 255, 153)
End Function

Public Function PaleRed() As Long
    PaleRed = RGB(255, 153, 153)
End Function

Public Function PaleGreen() As Long
    PaleGreen = RGB(153, 255, 153)
End Function

Public Function PaleCyan() As Long
    PaleCyan = RGB(208, 255, 255)
End Function

Public Function MedDkCyan() As Long
    MedDkCyan = RGB(0, 192, 192)
End Function

Public Function White() As Long
    ' Not in current use.
    White = RGB(255, 255, 255)
End Function

Public Function PaleOrange() As Long
    ' Not in current use.
    PaleOrange = RGB(255, 208, 153)
End Function




