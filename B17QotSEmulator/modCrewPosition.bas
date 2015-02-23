'******************************************************************************
' modCrewPosition.bas
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

Attribute VB_Name = "modCrewPosition"
Option Explicit

Public prsCrewPosition As New ADODB.Recordset

Dim strErrMsg As String

'******************************************************************************
' GetCrewPositionRecordset
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetCrewPositionRecordset() As Boolean
    On Error GoTo ErrorTrap
   
    GetCrewPositionRecordset = True

    pobjCmnd.CommandText = "SELECT * FROM CrewPosition ORDER BY KeyField"

    prsCrewPosition.CursorLocation = adUseClient
    prsCrewPosition.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    prsCrewPosition!KeyField.Properties("Optimize") = True
    prsCrewPosition.Sort = "KeyField ASC"
    
    Exit Function
   
CleanUp:
   
    Call FreeRecordset(prsCrewPosition)
   
    Exit Function
   
ErrorTrap:
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetCrewPositionRecordset() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    GetCrewPositionRecordset = False
    
    Resume CleanUp

End Function

'******************************************************************************
' LookupCrewPosition
'
' INPUT:  n/a
'
' OUTPUT: The position string if successful, otherwise blank.
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  Search for LookupKeyField in prsCrewPosition. If it is found, return
'         true and CrewPosition; if it is not found (which should never happen),
'         return false and blank.
'******************************************************************************
Public Function LookupCrewPosition(ByVal LookupKeyField As Integer, ByRef CrewPosition As String) As Boolean
    
    LookupCrewPosition = False
    CrewPosition = ""

    With frmMainMenu
        
        prsCrewPosition.MoveFirst
        Do Until prsCrewPosition.EOF
            
            If LookupKeyField = prsCrewPosition![KeyField] Then
                CrewPosition = prsCrewPosition![CrewPosition]
                LookupCrewPosition = True
                Exit Function
            End If
            
            prsCrewPosition.MoveNext
        Loop
    
    End With

    ' If the crew position had been found, we would have previously
    ' exitted. Therefore, an error condition exists.
    
    strErrMsg = "LookupCrewPosition() " & vbCrLf & vbCrLf & _
                "CrewPosition " & LookupKeyField & " not found."

    MsgBox strErrMsg, (vbCritical + vbOKOnly)

End Function

'******************************************************************************
' PopulateCrewPositionCombo
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Populate the position combo on the airman tab.
'******************************************************************************
Public Sub PopulateCrewPositionCombo()
    With frmMainMenu
        
        prsCrewPosition.MoveFirst
        Do Until prsCrewPosition.EOF
            
            .cboCrewPosition.AddItem prsCrewPosition![CrewPosition] ' AIRMAN_TAB
            
            prsCrewPosition.MoveNext
        Loop
        
    End With
End Sub


