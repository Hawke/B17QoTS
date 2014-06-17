'******************************************************************************
' modDBConnection.bas
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

Attribute VB_Name = "modDBConnection"
Option Explicit

Public pstrMyDB As String           ' Emulator database path\filename
Public pstrConnString As String     ' Database connection string to avoid DSN
Public pobjConn As ADODB.Connection ' Database connection object
Public pintOpenTrans As Integer     ' Count of open database transactions
Public pobjCmnd As ADODB.Command ' xxx test ...

Dim strErrmsg As String

'******************************************************************************
' OpenDBConnection
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Open a database connection. If successful, return true. If there
'         is an error, pop a msgbox, close the connection, free memory, then
'         return false. The database connection should remain open until the
'         program exits, either deliberately, or due to error.
'******************************************************************************
Public Function OpenDBConnection()
    On Error GoTo ErrorTrap

'MsgBox "OpenDBConnection()"
    
    OpenDBConnection = True
   
    pstrMyDB = App.Path + "\B17QotSDatabase.mdb"

    pstrConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pstrMyDB & ";Persist Security Info=False"
   
    Set pobjConn = New ADODB.Connection
    
    pobjConn.ConnectionTimeout = 120
    
    pobjConn.CommandTimeout = 120
    
    pobjConn.ConnectionString = pstrConnString
    
    pobjConn.Open
   
    Set pobjCmnd = New ADODB.Command ' xxx test ...
    
    Set pobjCmnd.ActiveConnection = pobjConn ' xxx test ...
    
    ' Even though the DB connection is open, there are no open transactions.
    
    pintOpenTrans = 0
   
    Exit Function
   
CleanUp:
   
    If Not pobjConn Is Nothing Then
        If pobjConn.State <> adStateClosed Then pobjConn.Close
        Set pobjConn = Nothing
    End If
   
    Exit Function
   
ErrorTrap:
    
    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "OpenDBConnection() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    OpenDBConnection = False
    
    Resume CleanUp

End Function

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
Public Sub ExitEmulatorX() ' qwe
    ' Gracefully shut down the emulator: Free memory, close the DB connection,
    ' then exit.
    
    Call FreeRecordset(prsGroup)
    Call FreeRecordset(prsSquadron)
    Call FreeRecordset(prsBomber)
    Call FreeRecordset(prsBomberModel)
    Call FreeRecordset(prsBomberStatus)
    Call FreeRecordset(prsBomberSquadron)
    Call FreeRecordset(prsAirman)
    Call FreeRecordset(prsRank)
    Call FreeRecordset(prsCrewPosition)
    Call FreeRecordset(prsAirmanStatus)
    
    Call CloseDBConnection
Dim a
a = 1
Unload frmMainMenu
    End

End Sub

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
Public Function CloseDBConnection()
    ' Close a database connection.  If successful, return true. If there
    ' is an error, pop a msgbox, try to close the connection and free
    ' memory anyway, then return false.
'    On Error GoTo ErrorTrap
   
'MsgBox "CloseDBConnection()"
    
    CloseDBConnection = True
   
CleanUp:
   
    If Not pobjConn Is Nothing Then
        If pobjConn.State <> adStateClosed Then pobjConn.Close
        Set pobjConn = Nothing
    End If
   
    Exit Function
   
ErrorTrap:
    
    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "CloseDBConnection() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    CloseDBConnection = False
    
    Resume CleanUp

End Function


