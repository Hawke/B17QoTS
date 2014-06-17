'******************************************************************************
' modMisc.bas
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

Attribute VB_Name = "modMisc"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'******************************************************************************
'
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:
'******************************************************************************
Public Sub FreeRecordset(ByRef rsTemp As ADODB.Recordset)

    If Not rsTemp Is Nothing Then
        If rsTemp.State <> adStateClosed Then rsTemp.Close
        Set rsTemp = Nothing
    End If
   
End Sub

'*******************************************************************************
' RecordsInSet
'
' INPUT:  Any recordset.
'
' OUPUT:  n/a
'
' RETURN: Number of records in the set.
'
' NOTES:  n/a
'*******************************************************************************
Public Function RecordsInSet(ByRef rsTemp As ADODB.Recordset) As Integer

    If Not rsTemp Is Nothing Then
        ' 0 or some number of records
        RecordsInSet = rsTemp.RecordCount
    Else
        RecordsInSet = 0
    End If

End Function

'*******************************************************************************
' UpdateMessage
'
' INPUT:  Text to add to mission display.
'
' OUPUT:  n/a
'
' RETURN: n/a
'
' NOTES:  This routine was placed in a module, rather than the form's code, so
'         that it would be available to all areas of the program from which it
'         is called. frmMission needs to be loaded for it to work.
'*******************************************************************************
Public Sub UpdateMessage(ByVal strNewText As String)
'Static intLines As Integer
'        UpdateMessage (Mission.Options.Delay / 1000) & " second delay"
'        MsgBox (Mission.Options.Delay / 1000) & " second delay"
'Sleep 250 ' Mission.Options.Delay

    With frmMission
'        intLines = intLines + 1

'        If intLines <= 17 Then
'            .rtbMessages.Text = .rtbMessages.Text & strNewText & vbCrLf
'        ElseIf intLines = 18 Then
'            .rtbMessages.Text = .rtbMessages.Text & strNewText & vbCrLf
'            Sleep (Mission.Options.Delay * 10)
'        ElseIf intLines >= 19 Then
            .rtbMessages.SelStart = Len(.rtbMessages.Text)
            .rtbMessages.Text = .rtbMessages.Text & strNewText & vbCrLf
            Sleep Mission.Options.Delay
'        End If
 
        .Refresh
    
    End With

End Sub


