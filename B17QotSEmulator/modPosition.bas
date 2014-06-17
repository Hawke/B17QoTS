'******************************************************************************
' modPosition.bas
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

Attribute VB_Name = "modPosition"
Option Explicit

'******************************************************************************
' PosManned
'
' INPUT:  Position to be examined.
'
' OUTPUT: n/a
'
' RETURN: True if the position is occupied, otherwise false.
'
' NOTES:  The airman's status is not considered, only whether there is an airman
'         at the position.
'******************************************************************************
Public Function PosManned(ByVal intPos As Integer) As Boolean

    PosManned = True
    
    If intPos <= 0 Then
        PosManned = False
    ElseIf Bomber.Position(intPos).CurrentSerialNum = UNMANNED_POSITION Then
        PosManned = False
    End If

End Function

'******************************************************************************
' PosExists
'
' INPUT:  Position to be examined.
'
' OUTPUT: n/a
'
' RETURN: True if the position exists, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function PosExists(ByVal intPos As Integer) As Boolean

    PosExists = True
    
    If intPos <= 0 Then
        PosExists = False
    ElseIf Bomber.Position(intPos).AssignedSerialNum = HIDDEN_POSITION Then
        PosExists = False
    End If

End Function

'******************************************************************************
' PosOccupied
'
' INPUT:  Position to be examined.
'
' OUTPUT: n/a
'
' RETURN: True if the position is occupied, otherwise false.
'
' NOTES:  The airman's status is not considered, only whether there is an airman
'         at the position.
'******************************************************************************
Public Function PosOccupied(ByVal intPos As Integer) As Boolean

    PosOccupied = True

    If PosManned(intPos) = False _
    Or PosExists(intPos) = False Then
    
        ' The position is either currently unoccupied, or the position does
        ' not even exist on the bomber.
    
        PosOccupied = False
    
    End If

End Function

'******************************************************************************
' GunManned
'
' INPUT:  Gun to be examined.
'
' OUTPUT: n/a
'
' RETURN: True if the gun is manned, otherwise false.
'
' NOTES:  The airman's status is not considered, only whether there is an airman
'         at the gun.
'******************************************************************************
Public Function GunManned(ByVal intGun As Integer) As Boolean
    
    GunManned = True
    
    If intGun <= 0 Then
        GunManned = False
    ElseIf Bomber.Gun(intGun).MannedBy = UNMANNED_POSITION Then
        GunManned = False
    End If

End Function

'******************************************************************************
' GunExists
'
' INPUT:  Gun to be examined.
'
' OUTPUT: n/a
'
' RETURN: True if the gun exists, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GunExists(ByVal intGun As Integer) As Boolean

    GunExists = True
    
    If intGun <= 0 Then
        GunExists = False
    ElseIf Bomber.Gun(intGun).MannedBy = HIDDEN_POSITION Then
        GunExists = False
    End If

End Function

'******************************************************************************
' GunOccupied
'
' INPUT:  Gun to be examined.
'
' OUTPUT: n/a
'
' RETURN: True if the gun is occupied, otherwise false.
'
' NOTES:  The airman's status is not considered, only whether there is an airman
'         at the gun.
'******************************************************************************
Public Function GunOccupied(ByVal intGun As Integer) As Boolean

    GunOccupied = True

    If GunManned(intGun) = False _
    Or GunExists(intGun) = False Then
    
        ' The gun is either currently unmanned, or the gun does not even exist
        ' on the bomber.
    
        GunOccupied = False
    
    End If

End Function

'******************************************************************************
' InStartingPos
'
' INPUT:  Position to be examined.
'
' OUTPUT: n/a
'
' RETURN: True if the airman currently occupies his assigned position, otherwise
'         false.
'
' NOTES:  Not in current use.
'******************************************************************************
Public Function InStartingPos(ByVal intCurrPos As Integer) As Boolean

    InStartingPos = False

    If Bomber.Airman(intCurrPos).AssignedPosition = intCurrPos Then
        ' The airman's assigned (starting) position is the same as his current
        ' position.
        InStartingPos = True
    End If

End Function

'******************************************************************************
' IsAssignedAirman
'
' INPUT:  Position to be examined.
'
' OUTPUT: n/a
'
' RETURN: True if the position is currently occupied by the originally assigned
'         airman, otherwise false.
'
' NOTES:  Not in current use.
'******************************************************************************
Public Function IsAssignedAirman(ByVal intPos As Integer) As Boolean

    IsAssignedAirman = False

    If Bomber.Position(intPos).AssignedSerialNum = Bomber.Position(intPos).CurrentSerialNum Then
        IsAssignedAirman = True
    End If

End Function



