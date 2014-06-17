'******************************************************************************
' modRandomizer.bas
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

Attribute VB_Name = "modRandomizer"
Option Explicit

'******************************************************************************
' RandomMonth
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Number that represents some month between August, 1942, and May, 1945.
'
' NOTES:  n/a
'******************************************************************************
Public Function RandomMonth() As Integer
    ' 0-base function.
    RandomMonth = Int(CAMPAIGN_DURATION * Rnd)
End Function

'******************************************************************************
' Random1D6
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Number between 1 and 6.
'
' NOTES:  n/a
'******************************************************************************
Public Function Random1D6() As Integer
    ' 1-base function.
    Random1D6 = Int((6 * Rnd) + 1)
End Function

'******************************************************************************
' Random2D6
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Number between 2 and 12.
'
' NOTES:  n/a
'******************************************************************************
Public Function Random2D6() As Integer
    ' 2-base function.
    Random2D6 = Int((6 * Rnd) + 1) + Int((6 * Rnd) + 1)
End Function

'******************************************************************************
' RandomD66
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Number between 11-16, 21-26, 31-36, 41-46, 51-56 and 61-66.
'
' NOTES:  n/a
'******************************************************************************
Public Function RandomD66() As Integer
    ' 11-base function.
    RandomD66 = (Int((6 * Rnd) + 1) * 10) + Int((6 * Rnd) + 1)
End Function

'******************************************************************************
' RandomD100
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Number between 1 and 100.
'
' NOTES:  n/a
'******************************************************************************
Public Function RandomD100() As Integer
    ' 1-base function.
    RandomD100 = Int((100 * Rnd) + 1)
End Function

'******************************************************************************
' RandomDX
'
' INPUT:  The maximum value to be rolled (X).
'
' OUTPUT: n/a
'
' RETURN: Number between 1 and X (the passed in value).
'
' NOTES:  n/a
'******************************************************************************
Public Function RandomDX(ByVal intSides As Integer) As Integer
    ' 1-base function.
    RandomDX = Int((intSides * Rnd) + 1)
End Function

