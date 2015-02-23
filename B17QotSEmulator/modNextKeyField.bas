'******************************************************************************
' modNextKeyField.bas
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

Attribute VB_Name = "modNextKeyField"
Option Explicit

Dim strErrMsg As String

'******************************************************************************
' NextKeyField
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The next available numeric key. Return 0 if there was an error, then
'         call ExitEmulator().
'
' NOTES:  Check if there is a gap in KeyFields (due to deletion); if so, that
'         is the next available KeyField. For instance, if there are airmen
'         numbered 1,2,4 and 6, the next available KeyField is 3. If there is
'         no gap, the next KeyField is one more than the highest KeyField.
'         For instance, if there are bombers numbered 1, 2, 3, and 4, then the
'         next available key is 5. Note that each keyed table has a distinct
'         set of keys. Returning 0 indicates an error occured. This function is
'         called like so:
'
'         intRetVal = NextKeyField(prsBomber, "Bomber")
'
'******************************************************************************
Public Function NextKeyField(frsOrig As Recordset, strTable As String) As Integer
    On Error GoTo ErrorTrap

    Dim intLastKey As Integer
    Dim frsTemp As New ADODB.Recordset

    NextKeyField = 0
    intLastKey = 0

    ' Create a temporary recordset, then sort it, to avoid altering the
    ' original recordset.
    
    Set frsTemp = frsOrig.Clone
    
    frsTemp.Sort = "KeyField ASC"
    
    frsTemp.MoveFirst
    Do Until frsTemp.EOF
        
        If frsTemp![KeyField] < 1 Then
            
            ' KeyFields drive everything the emulator does. Though KeyField
            ' should never be less than 1, it could be if there is a bug
            ' elsewhere in the system, or if some ignorant user was data
            ' fiddling. So, we need to check.
            
            strErrMsg = "KeyField less than 1!" & vbCrLf & vbCrLf & _
                        "There is a record in the " & strTable & _
                        " table where the KeyField is " & frsTemp![KeyField] & _
                        ". All KeyField primary keys must be 1 or more (though a " & _
                        "foreign key may be 0)." & vbCrLf & vbCrLf & _
                        "If you don't want to lose all your data, you can " & _
                        "try fiddling the database in Access, but you will " & _
                        "probably make the situation worse. Your best bet " & _
                        "is to re-initialize the database."
            
            MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
            GoTo CleanUp
        End If

        If frsTemp![KeyField] > (intLastKey + 1) Then
            
            ' Each KeyField should be one more than the previous KeyField;
            ' if not, that indicates a gap. Use the missing value as the
            ' new KeyField. If there is more than one value in the gap,
            ' then the lowest missing value will be returned.
            
            NextKeyField = intLastKey + 1
            
            GoTo CleanUp
        Else
            
            ' When the comparison was conducted, KeyField should never be
            ' less than or equal to intLastKey, therefore it must be one
            ' more than intLastKey. Set intLastKey equal to KeyField.
            
            intLastKey = frsTemp![KeyField]
        
        End If
            
        frsTemp.MoveNext
    Loop

    ' If a gap had been found, the function would have already returned the
    ' lowest missing value, therefore there was no gap. Instead return the
    ' next available number.

    frsTemp.MoveLast
    
    NextKeyField = frsTemp![KeyField] + 1

CleanUp:

    If Not frsTemp Is Nothing Then
        If frsTemp.State = adStateClosed Then frsTemp.Close
        Set frsTemp = Nothing
    End If

    Exit Function
   
ErrorTrap:
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "NextKeyField() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    Resume CleanUp

End Function

