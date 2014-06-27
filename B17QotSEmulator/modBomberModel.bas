Attribute VB_Name = "modBomberModel"
'******************************************************************************
' modBomberModel.bas
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

Option Explicit

Public prsBomberModel As New ADODB.Recordset

Dim strErrMsg As String

'******************************************************************************
' GetBomberModelRecordset
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetBomberModelRecordset() As Boolean
    On Error GoTo ErrorTrap
   
    GetBomberModelRecordset = True

    pobjCmnd.CommandText = "SELECT * FROM BomberModel ORDER BY KeyField"

    prsBomberModel.CursorLocation = adUseClient
    prsBomberModel.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    prsBomberModel!KeyField.Properties("Optimize") = True
    prsBomberModel.Sort = "KeyField ASC"
    
    Exit Function
   
CleanUp:
   
    Call FreeRecordset(prsBomberModel)
   
    Exit Function
   
ErrorTrap:
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetBomberModelRecordset() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    GetBomberModelRecordset = False
    
    Resume CleanUp

End Function

'******************************************************************************
' LookupBomberModel
'
' INPUT:  n/a
'
' OUTPUT: The model string if successful, otherwise blank.
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  Search for LookupKeyField in prsBomberModel. If it is found, return
'         true and BomberModel; if it is not found (which should never happen),
'         return false and blank.
'******************************************************************************
Public Function LookupBomberModel(ByVal LookupKeyField As Integer, Optional ByRef BomberModel As String = vbNullString) As Boolean
    
    LookupBomberModel = False
    BomberModel = ""

    With frmMainMenu
        
        prsBomberModel.MoveFirst
        Do Until prsBomberModel.EOF
            
            If LookupKeyField = prsBomberModel![KeyField] Then
                BomberModel = prsBomberModel![BomberModel]
                LookupBomberModel = True
                Exit Function
            End If
            
            prsBomberModel.MoveNext
        Loop
    
    End With

    ' If the BomberModel had been found, we would have previously exitted.
    ' Therefore, an error condition exists.
    
    strErrMsg = "LookupBomberModel() " & vbCrLf & vbCrLf & _
                "BomberModel " & LookupKeyField & " not found."

    MsgBox strErrMsg, (vbCritical + vbOKOnly)

End Function

'******************************************************************************
' PopulateBomberModelCombo
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  There are two bomber model combos: One on the bomber tab and one on
'         the mission tab.
'******************************************************************************
Public Sub PopulateBomberModelCombo()
    
    With frmMainMenu

        prsBomberModel.MoveFirst
        Do Until prsBomberModel.EOF

            .cboBomberModel(BOMBER_TAB).AddItem prsBomberModel![BomberModel]
            .cboBomberModel(MISSION_TAB).AddItem prsBomberModel![BomberModel]

            prsBomberModel.MoveNext
        Loop

    End With

End Sub




