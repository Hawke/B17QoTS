'******************************************************************************
' modLiberation.bas
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

Attribute VB_Name = "modLiberation"
Option Explicit

Public prsLiberation As New ADODB.Recordset

Dim strErrMsg As String

'******************************************************************************
' GetLiberationPct
'
' INPUT:  The terrain the bomber is over and the mission date.
'
' OUTPUT: n/a
'
' RETURN: Number indicating what percentage of terrain was liberated on that
'         date.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetLiberationPct(ByVal strTer As String, ByVal intDate As Integer) As Integer
    On Error GoTo ErrorTrap
    
    Dim arrTer() As String

    GetLiberationPct = 0

    If InStr(1, strTer, "-", 1) >= 1 Then

        ' Split territory. Determine which side of the split the bomber is on.
    
        arrTer = Split(strTer, "-")
        
        If Random1D6() <= 3 Then
            strTer = arrTer(0)
        Else
            strTer = arrTer(1)
        End If
    
    End If
    
    ' If the country is England, it is always in Allied hands; if the
    ' country is Norway, it is always in Axis hands.
    
    If strTer = "Base" _
    Or strTer = ENGLAND_TER _
    Or strTer = WATER_TER Then ' TODO: nothing is initialized to england_ter
        GetLiberationPct = 100
        Exit Function
    ElseIf strTer = NORWAY_TER _
    Or strTer = ALPS_TER Then
        GetLiberationPct = 0
        Exit Function
    End If
    
    ' For a given country, get all liberation values, regardless of date.
    
    pobjCmnd.CommandText = "SELECT Date, LiberationPct " & _
                           "FROM Liberation " & _
                           "WHERE Ter = '" & strTer & "'"

    prsLiberation.CursorLocation = adUseClient
    prsLiberation.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
' ???    prsLiberation!KeyField.Properties("Optimize") = True
    prsLiberation.Sort = "Date ASC"
    
    ' If an error was encountered selecting the data, we would not be here.
    ' Return the percent liberated. Fall through to cleanup, as we only
    ' temporarily needed the recordset.
    
    prsLiberation.MoveFirst
    Do Until prsLiberation.EOF

        If intDate < prsLiberation![Date] Then
            If prsLiberation.AbsolutePosition = 1 Then
                ' The first record is when the first chunk of the country was
                ' liberated. Since the current date falls before the first
                ' liberation date, none of the country has been liberated.
                GetLiberationPct = 0
                Exit Do
            End If
        ElseIf intDate = prsLiberation![Date] Then
            ' The current date is when a chunk was liberated.
            GetLiberationPct = prsLiberation![LiberationPct]
            Exit Do
        Else
            ' Record the liberation percentage for the last month in which a
            ' chunk of the country was liberated.
            GetLiberationPct = prsLiberation![LiberationPct]
        End If
            
        prsLiberation.MoveNext
    Loop
    
CleanUp:
   
    Call FreeRecordset(prsLiberation)
   
    Exit Function
   
ErrorTrap:
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetLiberationPct() " & vbCrLf & vbCrLf & _
                Err.Description & vbCrLf & _
                " The emulator will assume that 0% of the country (" & _
                strTer & ") has been liberated."

    MsgBox strErrMsg, (vbExclamation + vbOKOnly)
    
    Err.Clear
    
    GetLiberationPct = 0
    
    Resume CleanUp

End Function

