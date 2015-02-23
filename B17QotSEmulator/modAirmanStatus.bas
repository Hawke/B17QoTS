'******************************************************************************
' modAirmanStatus.bas
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

Attribute VB_Name = "modAirmanStatus"
Option Explicit

Public prsAirmanStatus As New ADODB.Recordset

Dim strErrMsg As String

'******************************************************************************
' GetAirmanStatusRecordset
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetAirmanStatusRecordset() As Boolean
    On Error GoTo ErrorTrap
   
    GetAirmanStatusRecordset = True

    pobjCmnd.CommandText = "SELECT * FROM AirmanStatus ORDER BY KeyField"

    prsAirmanStatus.CursorLocation = adUseClient
    prsAirmanStatus.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    prsAirmanStatus!KeyField.Properties("Optimize") = True
    prsAirmanStatus.Sort = "KeyField ASC"
    
    Exit Function
   
CleanUp:

    Call FreeRecordset(prsAirmanStatus)

    Exit Function
   
ErrorTrap:
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetAirmanStatusRecordset() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    GetAirmanStatusRecordset = False
    
    Resume CleanUp

End Function

'******************************************************************************
' LookupAirmanStatus
'
' INPUT:  n/a
'
' OUTPUT: The status string if successful, otherwise blank.
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  Search for LookupKeyField in prsAirmanStatus. If it is found, return
'         true and AirmanStatus; if it is not found (which should never
'         happen), return false and blank.
'******************************************************************************
Public Function LookupAirmanStatus(ByVal LookupKeyField As Integer, ByRef AirmanStatus As String) As Boolean
    
    LookupAirmanStatus = False
    AirmanStatus = ""

    With frmMainMenu
        
        prsAirmanStatus.MoveFirst
        Do Until prsAirmanStatus.EOF
            
            If LookupKeyField = prsAirmanStatus![KeyField] Then
                AirmanStatus = prsAirmanStatus![Status]
                LookupAirmanStatus = True
                Exit Function
            End If
            
            prsAirmanStatus.MoveNext
        Loop
    
    End With

    ' If the airman's status had been found, we would have previously
    ' exitted. Therefore, an error condition exists.
    
    strErrMsg = "LookupAirmanStatus() " & vbCrLf & vbCrLf & _
                "AirmanStatus " & LookupKeyField & " not found."

    MsgBox strErrMsg, (vbCritical + vbOKOnly)

End Function


