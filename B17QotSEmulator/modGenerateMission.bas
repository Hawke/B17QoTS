'******************************************************************************
' modGenerateMission.bas
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

Attribute VB_Name = "modGenerateMission"
Option Explicit

Public prsTarget As New ADODB.Recordset
Public prsBomberTarget As New ADODB.Recordset

Dim strErrmsg As String

'******************************************************************************
' FillMissionTabFields
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if the fields were filled, otherwise false.
'
' NOTES:  Some prsBomber record must be pointed at -- either using MoveFirst
'         or a LookupBomber() call -- before this function is called.
'******************************************************************************
Public Function FillMissionTabFields() As Boolean
    
    Dim strIgnore As String
    
    FillMissionTabFields = True
    
    With frmMainMenu

        ' Populate the recordset lookup fields. Work our way back up the
        ' object tree, getting relevant pieces of information along the way.
        ' Afterwards, repoint the recordsets so they are back in synch with
        ' the information on the corresponding tabs.
        
'MsgBox "FillMissionTabFields():" & vbCrLf & _
       ".cboName(MISSION_TAB).ListIndex = " & .cboName(MISSION_TAB).ListIndex & vbCrLf & _
       "intBomberMission(.cboName(MISSION_TAB).ListIndex) = " & intBomberMission(.cboName(MISSION_TAB).ListIndex) & vbCrLf & _
       "strIgnore = " & strIgnore
    
        If LookupBomber(intBomberMission(.cboName(MISSION_TAB).ListIndex), LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillMissionTabFields = False
            Exit Function
        End If

'MsgBox "prsBomber![BomberModel] = " & prsBomber![BomberModel]
        
        If LookupBomberModel(prsBomber![BomberModel], strIgnore) = False Then
            FillMissionTabFields = False
            Exit Function
        Else
            .cboBomberModel(MISSION_TAB).Text = prsBomberModel![BomberModel]
            
            If prsBomberModel![KeyField] = YB40 Then
                .chkExtraAmmoInBombBay.Enabled = True
            Else
                ' Other models must carry bombs.
                .chkExtraAmmoInBombBay.Enabled = False
                .chkExtraAmmoInBombBay.Value = vbUnchecked
            End If
        End If

'MsgBox "prsBomber![Squadron] = " & prsBomber![Squadron]
        
        If LookupSquadron(prsBomber![Squadron], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillMissionTabFields = False
            Exit Function
        ElseIf LookupGroup(prsSquadron![Group], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            FillMissionTabFields = False
            Exit Function
        End If
        
        ' Need to get group for the squadron before the base may be
        ' examined, otherwise we are pointing at whatever group we last
        ' pointed at.
            
        .txtBase.Text = prsGroup![Base] ' Save value in hidden control.
        
        ' Bomber model is displayed, but it is actually bomber type that
        ' determines the filtered target list.
            
        If .txtBase.Text = ENGLAND_TER _
        And (prsSquadron![BomberType] = B17_TYPE _
        Or prsSquadron![BomberType] = B24_TYPE) Then
            .chkExpandedTargetList.Enabled = True
        Else
            .chkExpandedTargetList.Enabled = False
            .chkExpandedTargetList.Value = vbUnchecked
        End If
        
'MsgBox "prsSquadron![Group] = " & prsSquadron![Group]
        
        Select Case .txtBase.Text
            Case ENGLAND_TER:
                .optEngland(MISSION_TAB).Value = True
                .chkRedTailAngels.Enabled = False
                .chkRedTailAngels.Value = vbUnchecked
            Case ITALY_TER:
                .optItaly(MISSION_TAB).Value = True
                .chkRedTailAngels.Enabled = True
        End Select
        
    End With

    prsBomber.Bookmark = varBomberCurrentlyOnTab
    prsSquadron.Bookmark = varSquadronCurrentlyOnTab
    prsGroup.Bookmark = varGroupCurrentlyOnTab

End Function

'******************************************************************************
' GetTargetRecordset
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetTargetRecordset() As Boolean
    On Error GoTo ErrorTrap
   
    GetTargetRecordset = True

    pobjCmnd.CommandText = "SELECT * FROM Target ORDER BY Name"

    prsTarget.CursorLocation = adUseClient
    prsTarget.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    prsTarget!KeyField.Properties("Optimize") = True
    prsTarget.Sort = "Name ASC"
    
    Exit Function
   
CleanUp:
   
    Call FreeRecordset(prsTarget)
   
    Exit Function
   
ErrorTrap:
    
    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetTargetRecordset() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    GetTargetRecordset = False
    
    Resume CleanUp

End Function

'******************************************************************************
' LookupTarget
'
' INPUT:  n/a
'
' OUTPUT: Target name if successful, otherwise blank.
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function LookupTarget(ByVal LookupKeyField As Integer, ByVal LookupType As Integer, ByRef TargetName As String)
    
    Dim intIndex As Integer
    
    LookupTarget = False
    TargetName = ""
    intIndex = 1

    With frmMainMenu
        
        prsTarget.MoveFirst
        Do Until prsTarget.EOF
            
            Select Case LookupType
                Case LOOKUP_BY_LISTINDEX:
                    If LookupKeyField = intIndex Then
                        TargetName = prsTarget![Name]
                        LookupTarget = True
                        Exit Function
                    End If
                Case LOOKUP_BY_KEYFIELD:
                    If LookupKeyField = prsTarget![KeyField] Then
                        TargetName = prsTarget![Name]
                        LookupTarget = True
                        Exit Function
                    End If
            End Select
            
            intIndex = intIndex + 1
            prsTarget.MoveNext
        Loop
    
    End With

    ' If the Target had been found, we would have previously exitted.
    ' Therefore, an error condition exists.
    
    strErrmsg = "LookupTarget() " & vbCrLf & vbCrLf & _
                "Target " & LookupKeyField & " not found."

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

End Function

'******************************************************************************
' PopulateDateCombos
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  There are two bomber combos: One for months, the other for years.
'         The list of months depends on the year, so the month combo is
'         dynamically populated each time a year is selected.
'******************************************************************************
Public Sub PopulateDateCombos()
    Dim intIndex As Integer
    
    With frmMainMenu
    
        ' Delete the previous combo items to prevent concatenating two sets
        ' of dates.

        For intIndex = 0 To (.cboMonth.ListCount - 1)
            .cboMonth.RemoveItem 0
        Next intIndex
        
        For intIndex = 0 To (.cboYear.ListCount - 1)
            .cboYear.RemoveItem 0
        Next intIndex
        
        If .optEngland(MISSION_TAB).Value = True Then

                .cboMonth.AddItem "January"
                .cboMonth.AddItem "February"
                .cboMonth.AddItem "March"
                .cboMonth.AddItem "April"
                .cboMonth.AddItem "May"
                .cboMonth.AddItem "June"
                .cboMonth.AddItem "July"
                .cboMonth.AddItem "August"
                .cboMonth.AddItem "September"
                .cboMonth.AddItem "October"
                .cboMonth.AddItem "November"
                .cboMonth.AddItem "December"
            
                .cboYear.AddItem "1942"
                .cboYear.AddItem "1943"
                .cboYear.AddItem "1944"
                .cboYear.AddItem "1945"
        
        Else ' 15th Air Force

            .cboMonth.AddItem "January"
            .cboMonth.AddItem "February"
            .cboMonth.AddItem "March"
            .cboMonth.AddItem "April"
            .cboMonth.AddItem "May"
            .cboMonth.AddItem "June"
            .cboMonth.AddItem "July"
            .cboMonth.AddItem "August"
            .cboMonth.AddItem "September"
            .cboMonth.AddItem "October"
            .cboMonth.AddItem "November"
            .cboMonth.AddItem "December"
                
            .cboYear.AddItem "1943"
            .cboYear.AddItem "1944"
            .cboYear.AddItem "1945"
        
        End If
    
    End With
End Sub

'******************************************************************************
' AdjustDateLists
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Dynamically determine the list of months based on the bomber's base
'         and the selected year.
'******************************************************************************
Public Sub AdjustDateLists()
    With frmMainMenu
        
        ' Adjust the available months based on the year.
    
        If .cboMonth.List(0) = "January" Then
            
            ' The previous year was blank, 1943 (England), 1944 or 1945.
            
            If .cboYear.Text = "1942" Then
                
                ' The current year is 1942.
            
                If .cboMonth.ListCount = 12 Then
    
                    ' The year was changed from either blank, 1943 (England)
                    ' or 1944 to 1942. In inverse order, delete January
                    ' through July.
            
                    .cboMonth.RemoveItem 6 ' July
                    .cboMonth.RemoveItem 5 ' June
                    .cboMonth.RemoveItem 4 ' May
                    .cboMonth.RemoveItem 3 ' April
                    .cboMonth.RemoveItem 2 ' March
                    .cboMonth.RemoveItem 1 ' February
                    .cboMonth.RemoveItem 0 ' January
                
                Else ' .cboMonth.ListCount = 5
    
                    ' The year was changed from 1945 to 1942. Insert August
                    ' through December, then in inverse order delete January
                    ' through May.
    
                    .cboMonth.AddItem "August"
                    .cboMonth.AddItem "September"
                    .cboMonth.AddItem "October"
                    .cboMonth.AddItem "November"
                    .cboMonth.AddItem "December"
    
                    .cboMonth.RemoveItem 4 ' May
                    .cboMonth.RemoveItem 3 ' April
                    .cboMonth.RemoveItem 2 ' March
                    .cboMonth.RemoveItem 1 ' February
                    .cboMonth.RemoveItem 0 ' January
    
                End If

            ElseIf .cboYear.Text = "1943" Then
            
                ' The current year is 1943.
                
                If .optEngland(MISSION_TAB).Value = True Then
                
                    ' The base is in England.
                
                    If .cboMonth.ListCount = 5 Then
                    
                        ' The year was changed from 1945 to 1943 (England).
    
                        .cboMonth.AddItem "June"
                        .cboMonth.AddItem "July"
                        .cboMonth.AddItem "August"
                        .cboMonth.AddItem "September"
                        .cboMonth.AddItem "October"
                        .cboMonth.AddItem "November"
                        .cboMonth.AddItem "December"
                    
                    End If
    
                Else ' .optItaly(MISSION_TAB).Value = True
                
                    ' The base is in Italy.
                
                    If .cboMonth.ListCount = 12 Then
            
                        ' The year was changed from either blank, 1943
                        ' (England) or 1944 to 1943 (Italy). In inverse
                        ' order, delete January through October.
                    
                        .cboMonth.RemoveItem 9 ' October
                        .cboMonth.RemoveItem 8 ' September
                        .cboMonth.RemoveItem 7 ' August
                        .cboMonth.RemoveItem 6 ' July
                        .cboMonth.RemoveItem 5 ' June
                        .cboMonth.RemoveItem 4 ' May
                        .cboMonth.RemoveItem 3 ' April
                        .cboMonth.RemoveItem 2 ' March
                        .cboMonth.RemoveItem 1 ' February
                        .cboMonth.RemoveItem 0 ' January
                        
                    Else ' .cboMonth.ListCount = 5
           
                        ' The year was changed from 1945 to 1943 (Italy).
                        ' Insert November through December, then in inverse
                        ' order delete January through May.
            
                        .cboMonth.AddItem "November"
                        .cboMonth.AddItem "December"
                    
                        .cboMonth.RemoveItem 4 ' May
                        .cboMonth.RemoveItem 3 ' April
                        .cboMonth.RemoveItem 2 ' March
                        .cboMonth.RemoveItem 1 ' February
                        .cboMonth.RemoveItem 0 ' January
        
                    End If
    
                End If
            
            ElseIf .cboYear.Text = "1944" Then
            
                ' The current year is 1944.
                
                If .cboMonth.ListCount = 5 Then
    
                    ' The year was changed from 1945 to 1944. Insert June
                    ' through December.
            
                    .cboMonth.AddItem "June"
                    .cboMonth.AddItem "July"
                    .cboMonth.AddItem "August"
                    .cboMonth.AddItem "September"
                    .cboMonth.AddItem "October"
                    .cboMonth.AddItem "November"
                    .cboMonth.AddItem "December"
    
                End If
            
            Else ' .cboYear.Text = "1945"
    
                ' The current year is 1945.
    
                If .cboMonth.ListCount = 12 Then
    
                    ' The year was changed from either blank, 1943 (England)
                    ' or 1944 to 1945. In inverse order, delete June through
                    ' December.
            
                    .cboMonth.RemoveItem 11 ' December
                    .cboMonth.RemoveItem 10 ' November
                    .cboMonth.RemoveItem 9  ' October
                    .cboMonth.RemoveItem 8  ' September
                    .cboMonth.RemoveItem 7  ' August
                    .cboMonth.RemoveItem 6  ' July
                    .cboMonth.RemoveItem 5  ' June
                
                End If
    
            End If
            
        ElseIf .cboMonth.List(0) = "August" Then
            
            ' The previous year was 1942.
            
            If .cboYear.Text = "1943" Then
            
                ' The current year is 1943.
                
                If .optEngland(MISSION_TAB).Value = True Then
                
                    ' The year was changed from 1942 to 1943 (England). In
                    ' inverse order, insert January through July.
                
                    .cboMonth.AddItem "July", 0
                    .cboMonth.AddItem "June", 0
                    .cboMonth.AddItem "May", 0
                    .cboMonth.AddItem "April", 0
                    .cboMonth.AddItem "March", 0
                    .cboMonth.AddItem "February", 0
                    .cboMonth.AddItem "January", 0
    
                Else ' .optItaly(MISSION_TAB).Value = True
                
                    ' The year was changed from 1942 to 1943 (Italy). In
                    ' inverse order, delete August through October.
                
                    .cboMonth.RemoveItem 9 ' October
                    .cboMonth.RemoveItem 8 ' September
                    .cboMonth.RemoveItem 7 ' August
                        
                End If
            
            ElseIf .cboYear.Text = "1944" Then
            
                ' The year was changed from 1942 to 1944. In inverse order,
                ' insert January through July.
            
                .cboMonth.AddItem "July", 0
                .cboMonth.AddItem "June", 0
                .cboMonth.AddItem "May", 0
                .cboMonth.AddItem "April", 0
                .cboMonth.AddItem "March", 0
                .cboMonth.AddItem "February", 0
                .cboMonth.AddItem "January", 0
            
            ElseIf .cboYear.Text = "1945" Then
            
                ' The year was changed from 1942 to 1945. In inverse order,
                ' insert January through May, then in inverse order delete
                ' August through December.
    
                .cboMonth.AddItem "May", 0
                .cboMonth.AddItem "April", 0
                .cboMonth.AddItem "March", 0
                .cboMonth.AddItem "February", 0
                .cboMonth.AddItem "January", 0
            
                .cboMonth.RemoveItem 9 ' December
                .cboMonth.RemoveItem 8 ' November
                .cboMonth.RemoveItem 7 ' October
                .cboMonth.RemoveItem 6 ' September
                .cboMonth.RemoveItem 5 ' August
            
            End If
            
        ElseIf .cboMonth.List(0) = "November" Then
            
            ' The previous year was 1943 (Italy).
            
            If .cboYear.Text = "1942" Then
            
                ' The year was changed from 1943 (Italy) to 1942. In inverse order,
                ' insert August through October.
                
                .cboMonth.AddItem "October", 0
                .cboMonth.AddItem "September", 0
                .cboMonth.AddItem "August", 0
            
            ElseIf .cboYear.Text = "1943" _
            And .optEngland(MISSION_TAB).Value = True Then
            
                ' The year was changed from 1943 (Italy) to 1943 (England).
                ' In inverse order, insert January through October.
            
                .cboMonth.AddItem "October", 0
                .cboMonth.AddItem "September", 0
                .cboMonth.AddItem "August", 0
                .cboMonth.AddItem "July", 0
                .cboMonth.AddItem "June", 0
                .cboMonth.AddItem "May", 0
                .cboMonth.AddItem "April", 0
                .cboMonth.AddItem "March", 0
                .cboMonth.AddItem "February", 0
                .cboMonth.AddItem "January", 0
            
            ElseIf .cboYear.Text = "1944" Then

                ' The year was changed from 1943 (Italy) to 1944. In inverse
                ' order, insert January through October.
            
                .cboMonth.AddItem "October", 0
                .cboMonth.AddItem "September", 0
                .cboMonth.AddItem "August", 0
                .cboMonth.AddItem "July", 0
                .cboMonth.AddItem "June", 0
                .cboMonth.AddItem "May", 0
                .cboMonth.AddItem "April", 0
                .cboMonth.AddItem "March", 0
                .cboMonth.AddItem "February", 0
                .cboMonth.AddItem "January", 0
            
            ElseIf .cboYear.Text = "1945" Then
            
                ' The year was changed from 1943 (Italy) to 1945. In inverse
                ' order, insert January through May, then in inverse order
                ' delete November through December.
    
                .cboMonth.AddItem "May", 0
                .cboMonth.AddItem "April", 0
                .cboMonth.AddItem "March", 0
                .cboMonth.AddItem "February", 0
                .cboMonth.AddItem "January", 0
            
                .cboMonth.RemoveItem 6 ' December
                .cboMonth.RemoveItem 5 ' November
            
            End If
            
        End If

    End With
End Sub

'******************************************************************************
' PopulatePositionCombos
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Public Sub PopulatePositionCombos()
    With frmMainMenu

        .cboSquadronPos.AddItem "High"
        .cboSquadronPos.AddItem "Middle"
        .cboSquadronPos.AddItem "Low"
    
        .cboFormationPos.AddItem "Lead"
        .cboFormationPos.AddItem "Middle"
        .cboFormationPos.AddItem "Tail"

    End With
End Sub

'******************************************************************************
' PopulateTargetCombo
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  The list of targets is dynamically adjusted based on bomber type,
'         base country, date and options.
'******************************************************************************
Public Sub PopulateTargetCombo()
    Dim strTypeFilter As String
    Dim intIndex As Integer
    Dim strIgnore As String

    With frmMainMenu
        
         ' Clone the target recordset so we do not affect the original
         ' recordset.
       
        Set prsBomberTarget = prsTarget.Clone
    
        ' Delete the previous combo items to prevent concatenating two sets
        ' of targets.
        
        For intIndex = 0 To (.cboTarget.ListCount - 1)
            .cboTarget.RemoveItem 0
        Next intIndex
        
        ' Filter the clone so that it only contains targets which have the
        ' same general bomber type as the specific bomber model.
        
        Select Case (.cboBomberModel(MISSION_TAB).ListIndex + 1)
            
            Case B17_C To YB40:
                
                If .optItaly(MISSION_TAB).Value = True Then
                    
                    ' Italy-based B-17s can only fly 15th Air Force missions.
                    ' ADO apears to object to "15thAirForceVariant" as a
                    ' field name, perhaps due to the numbers, so we call it
                    ' "ItalyVariant" instead.
                    
                    strTypeFilter = "ItalyVariant = " & True
            
                ElseIf .chkExpandedTargetList.Value = vbChecked Then
                    
                    ' England-based B17s may bomb additional targets if the
                    ' option is chosen.
                    
                    strTypeFilter = "ExpandedTargetVariant = " & True
            
                Else
                    
                    ' England-based B-17s bombing the original B-17QotS
                    ' target list is the default.
                    
                    strTypeFilter = "OriginalTarget = " & True
                
                End If

            Case B24_D To B24_LM:
                
                If .optItaly(MISSION_TAB).Value = True Then
                    
                    ' Italy-based B24s can only fly 15th Air Force missions.
                    
                    strTypeFilter = "ItalyVariant = " & True
                
                Else
                    
                    ' England-based B-24s bombing the Flying Boxcar target
                    ' list is the default.
                    
                    strTypeFilter = "FlyingBoxcarVariant = " & True
                
                End If

            Case AVRO_LANCASTER:
                
                ' Lancaster's may only bomb the Battle of Berlin target list.
                 
                 strTypeFilter = "LancasterVariant = " & True
        
        End Select
    
' OK                    strTypeFilter = "ExpandedTargetVariant = " & True
' OK                    strTypeFilter = "OriginalTarget = " & True
' BAD                    strTypeFilter = "15thAirForceVariant = " & True
'ADO apears to object to '15thAirForceVariant' as a field name, perhaps due to the numbers, so we call it 'ItalyVariant' instead.
'strTypeFilter = "ItalyVariant = " & True
' OK                 strTypeFilter = "LancasterVariant = " & True
' BAD                   strTypeFilter = "FlyingBoxcarVariant = " & True
'                   strTypeFilter = "FlyingBoxcarVariant = " & True
'MsgBox "strTypeFilter = '" & strTypeFilter & "'"

        prsBomberTarget.Filter = strTypeFilter
    
        prsBomberTarget.MoveFirst
        Do Until prsBomberTarget.EOF
                
            .cboTarget.AddItem prsBomberTarget![Name] ' MISSION_TAB
                
            prsBomberTarget.MoveNext
        Loop

    End With

End Sub

'******************************************************************************
' LookupBomberTarget
'
' INPUT:  n/a
'
' OUTPUT: Target name if successful, otherwise blank.
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function LookupBomberTarget(ByVal LookupKeyField As Integer, ByVal LookupType As Integer, ByRef TargetName As String) As Boolean
    
    Dim intIndex As Integer
    
    LookupBomberTarget = False
    TargetName = ""
    intIndex = 1

    With frmMainMenu
        
        prsBomberTarget.MoveFirst
        Do Until prsBomberTarget.EOF
            
            Select Case LookupType
                Case LOOKUP_BY_LISTINDEX:
                    If LookupKeyField = intIndex Then
                        TargetName = prsBomberTarget![Name]
                        LookupBomberTarget = True
                        Exit Function
                    End If
                Case LOOKUP_BY_KEYFIELD:
                    If LookupKeyField = prsBomberTarget![KeyField] Then
                        TargetName = prsBomberTarget![Name]
                        LookupBomberTarget = True
                        Exit Function
                    End If
            End Select
            
            intIndex = intIndex + 1
            prsBomberTarget.MoveNext
        Loop
    
    End With

    ' If the target had been found, we would have previously exitted.
    ' Therefore, an error condition exists.
    
    strErrmsg = "LookupBomberTarget() " & vbCrLf & vbCrLf & _
                "Target " & LookupKeyField & " not found."

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

End Function


