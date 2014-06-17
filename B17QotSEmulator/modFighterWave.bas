'******************************************************************************
' modFighterWave.bas
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

Attribute VB_Name = "modFighterWave"
Option Explicit

Dim strErrMsg As String

'******************************************************************************
' GetFighterWaveRecordset
'
' INPUT:  The B-3 roll, between 11 and 66. (The value may be greater for
'         special waves such as Ju-88s.)
'
' OUTPUT: Fighter wave recordset.
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetFighterWaveRecordset(ByVal intRoll As Integer, ByRef rsFighterWave As ADODB.Recordset) As Boolean
    On Error GoTo ErrorTrap

'    ' Clear old data before getting new data ???
'    Call FreeRecordset(rsFighterWave)
   
    GetFighterWaveRecordset = True
    
    pobjCmnd.CommandText = "SELECT * " & _
                           "FROM WaveSelection " & _
                           "WHERE Roll = " & intRoll ' 24 ' 63 ' intRoll

    rsFighterWave.CursorLocation = adUseClient
    rsFighterWave.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
'    rsFighterWave!KeyField.Properties("Optimize") = True
    rsFighterWave.Sort = "Position ASC"
    
    Exit Function
    
CleanUp:
   
    Call FreeRecordset(rsFighterWave)
   
    Exit Function
   
ErrorTrap:
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetFighterWaveRecordset() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    GetFighterWaveRecordset = False
    
    Resume CleanUp

End Function

'******************************************************************************
' GetWaveSize
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Number of fighter nodes attached to the wave.
'
' NOTES:  If all the fighter nodes have been popped, UBound() will throw an
'         error. If that happens return a friendly 0.
'******************************************************************************
Public Function GetWaveSize() As Integer
    On Error GoTo NoSize

    GetWaveSize = UBound(Wave.Fighter)

    Exit Function
    
NoSize:
    
    GetWaveSize = 0
    Resume Next

End Function


