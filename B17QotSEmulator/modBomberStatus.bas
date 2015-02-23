'******************************************************************************
' modBomberStatus.bas
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

Attribute VB_Name = "modBomberStatus"
Option Explicit

Public prsBomberStatus As New ADODB.Recordset

Dim strErrmsg As String

'******************************************************************************
' GetBomberStatusRecordset
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetBomberStatusRecordset() As Boolean
    On Error GoTo ErrorTrap
   
    GetBomberStatusRecordset = True

    pobjCmnd.CommandText = "SELECT * FROM BomberStatus ORDER BY KeyField"

    prsBomberStatus.CursorLocation = adUseClient
    prsBomberStatus.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    prsBomberStatus!KeyField.Properties("Optimize") = True
    prsBomberStatus.Sort = "KeyField ASC"
    
    Exit Function
   
CleanUp:

    Call FreeRecordset(prsBomberStatus)

    Exit Function
   
ErrorTrap:
    
    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetBomberStatusRecordset() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    GetBomberStatusRecordset = False
    
    Resume CleanUp

End Function

'******************************************************************************
' LookupBomberStatus
'
' INPUT:  n/a
'
' OUTPUT: The status string if successful, otherwise blank.
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  Search for LookupKeyField in prsBomberStatus. If it is found, return
'         true and BomberStatus; if it is not found (which should never
'         happen), return false and blank.
'******************************************************************************
Public Function LookupBomberStatus(ByVal LookupKeyField As Integer, ByRef BomberStatus As String) As Boolean
    
    LookupBomberStatus = False
    BomberStatus = ""

    With frmMainMenu
        
        prsBomberStatus.MoveFirst
        Do Until prsBomberStatus.EOF
            
            If LookupKeyField = prsBomberStatus![KeyField] Then
                BomberStatus = prsBomberStatus![Status]
                LookupBomberStatus = True
                Exit Function
            End If
            
            prsBomberStatus.MoveNext
        Loop
    
    End With

    ' If the bomber's status had been found, we would have previously
    ' exitted. Therefore, an error condition exists.
    
    strErrmsg = "LookupBomberStatus() " & vbCrLf & vbCrLf & _
                "BomberStatus " & LookupKeyField & " not found."

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

End Function




