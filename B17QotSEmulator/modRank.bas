'******************************************************************************
' modRank.bas
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

Attribute VB_Name = "modRank"
Option Explicit

Public prsRank As New ADODB.Recordset

Dim strErrMsg As String

'******************************************************************************
' GetRankRecordset
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetRankRecordset() As Boolean
    On Error GoTo ErrorTrap
   
    GetRankRecordset = True

    pobjCmnd.CommandText = "SELECT * FROM Rank ORDER BY KeyField"

    prsRank.CursorLocation = adUseClient
    prsRank.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    prsRank!KeyField.Properties("Optimize") = True
    prsRank.Sort = "KeyField ASC"
    
    Exit Function
   
CleanUp:
   
    Call FreeRecordset(prsRank)
   
    Exit Function
   
ErrorTrap:
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetRankRecordset() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    GetRankRecordset = False
    
    Resume CleanUp

End Function

'******************************************************************************
' LookupRank
'
' INPUT:  n/a
'
' OUTPUT: Rank string if successful, otherwise blank.
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function LookupRank(ByVal LookupKeyField As Integer, ByRef Rank As String) As Boolean
    
    LookupRank = False
    Rank = ""

    With frmMainMenu
        
        prsRank.MoveFirst
        Do Until prsRank.EOF
            
            If LookupKeyField = prsRank![KeyField] Then
                Rank = prsRank![Rank]
                LookupRank = True
                Exit Function
            End If
            
            prsRank.MoveNext
        Loop
    
    End With

    ' If the rank had been found, we would have previously exitted.
    ' Therefore, an error condition exists.
    
    strErrMsg = "LookupRank() " & vbCrLf & vbCrLf & _
                "Rank " & LookupKeyField & " not found."

    MsgBox strErrMsg, (vbCritical + vbOKOnly)

End Function

'******************************************************************************
' PopulateRankCombo
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Public Sub PopulateRankCombo()
    With frmMainMenu
        
        prsRank.MoveFirst
        Do Until prsRank.EOF
            
            .cboRank.AddItem prsRank![Rank] ' AIRMAN_TAB
            
            prsRank.MoveNext
        Loop
    
    End With
End Sub


