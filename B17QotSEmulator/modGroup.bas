'******************************************************************************
' modGroup.bas
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

Attribute VB_Name = "modGroup"
' The table is named "GroupT" because "Group" is a reserved SQL word.

Option Explicit

Public prsGroup As New ADODB.Recordset
Public varGroupCurrentlyOnTab As Variant

Dim strErrmsg As String

'******************************************************************************
' FillGroupTabFields
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if the fields were filled, otherwise false.
'
' NOTES:  Some prsGroup record must be pointed at -- either using MoveFirst
'         or a LookupGroup() call -- before this function is called.
'******************************************************************************
Public Function FillGroupTabFields() As Boolean
    
    Dim strCommander As String
    
    FillGroupTabFields = True
    
    With frmMainMenu
        ' Populate the non-lookup fields.
        
        .txtKeyField(GROUP_TAB).Text = prsGroup![KeyField]
        .cboName(GROUP_TAB).Text = prsGroup![Name]
        .txtSorties(GROUP_TAB).Text = prsGroup![Sorties]
        .txtKills(GROUP_TAB).Text = prsGroup![Kills]
        .txtPlanesLost(GROUP_TAB).Text = prsGroup![PlanesLost]
        .txtKIA(GROUP_TAB).Text = prsGroup![KIA]
        .txtMIA(GROUP_TAB).Text = prsGroup![MIA]
        .txtWounded(GROUP_TAB).Text = prsGroup![Wounded]
        .txtPOW(GROUP_TAB).Text = prsGroup![POW]
        .txtMedalOfHonor(GROUP_TAB).Text = prsGroup![MedalOfHonor]
        .txtDistinguishedServiceCross(GROUP_TAB).Text = prsGroup![DistinguishedServiceCross]
        .txtSilverStar(GROUP_TAB).Text = prsGroup![SilverStar]
        .txtDistinguishedFlyingCross(GROUP_TAB).Text = prsGroup![DistinguishedFlyingCross]
        .txtBronzeStarV(GROUP_TAB).Text = prsGroup![BronzeStarV]
        .txtPurpleHeart(GROUP_TAB).Text = prsGroup![PurpleHeart]
        .txtAirMedal(GROUP_TAB).Text = prsGroup![AirMedal]
        .txtDistinguishedUnitCitation(GROUP_TAB).Text = prsGroup![DistinguishedUnitCitation]
        .txtMeritoriousUnitCitation(GROUP_TAB).Text = prsGroup![MeritoriousUnitCitation]
    
        Select Case prsGroup![Base]
            Case ENGLAND_TER:
                .optEngland(GROUP_TAB).Value = True
            Case ITALY_TER:
                .optItaly(GROUP_TAB).Value = True
        End Select
        
        ' Populate the recordset lookup fields.
        
        If LookupAirman(prsGroup![Commander], LOOKUP_BY_KEYFIELD, strCommander) = False Then
            FillGroupTabFields = False
            Exit Function
        Else
            .cboCommander(GROUP_TAB).Text = strCommander
            
            ' Repoint the airman recordset to the record displayed on the airman tab.
            prsAirman.Bookmark = varAirmanCurrentlyOnTab
        End If

        ' If this is a default Group, disable the fields, otherwise ensure
        ' they are enabled.
        
        If prsGroup![Default] = True Then
            .chkDefault(GROUP_TAB).Value = vbChecked
            
            .cboCommander(GROUP_TAB).Enabled = False
            .cboCommander(GROUP_TAB).BackColor = vbButtonFace
            .optEngland(GROUP_TAB).Enabled = False
            .optItaly(GROUP_TAB).Enabled = False
        Else
            .chkDefault(GROUP_TAB).Value = vbUnchecked
        
            .cboCommander(GROUP_TAB).Enabled = True
            .cboCommander(GROUP_TAB).BackColor = vbWhite
            .optEngland(GROUP_TAB).Enabled = True
            .optItaly(GROUP_TAB).Enabled = True
        End If

    End With

End Function

'******************************************************************************
' ZeroGroupTabFields
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Name is not blanked, because this function should only be called
'         when the user types something in name, and we don't want to blank
'         out what the user is typing.
'******************************************************************************
Public Sub ZeroGroupTabFields()
    With frmMainMenu
        
        .txtKeyField(GROUP_TAB).Text = 0
        .txtSorties(GROUP_TAB).Text = 0
        .txtKills(GROUP_TAB).Text = 0
        .txtPlanesLost(GROUP_TAB).Text = 0
        .txtKIA(GROUP_TAB).Text = 0
        .txtWounded(GROUP_TAB).Text = 0
        .txtPOW(GROUP_TAB).Text = 0
        .txtMedalOfHonor(GROUP_TAB).Text = 0
        .txtDistinguishedServiceCross(GROUP_TAB).Text = 0
        .txtSilverStar(GROUP_TAB).Text = 0
        .txtDistinguishedFlyingCross(GROUP_TAB).Text = 0
        .txtBronzeStarV(GROUP_TAB).Text = 0
        .txtPurpleHeart(GROUP_TAB).Text = 0
        .txtAirMedal(GROUP_TAB).Text = 0
        .txtDistinguishedUnitCitation(GROUP_TAB).Text = 0
        .txtMeritoriousUnitCitation(GROUP_TAB).Text = 0
        
        .optEngland(GROUP_TAB).Value = True
        
        .cboCommander(GROUP_TAB).ListIndex = 0
            
        .chkDefault(GROUP_TAB).Value = vbUnchecked
        
        .cboCommander(GROUP_TAB).Enabled = True
        .cboCommander(GROUP_TAB).BackColor = vbWhite
        
        .optEngland(GROUP_TAB).Enabled = True
        .optItaly(GROUP_TAB).Enabled = True

    End With
End Sub

'******************************************************************************
' GetGroupRecordset
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetGroupRecordset() As Boolean
    On Error GoTo ErrorTrap
   
    GetGroupRecordset = True

    pobjCmnd.CommandText = "SELECT * FROM GroupT ORDER BY Name"

    prsGroup.CursorLocation = adUseClient
    prsGroup.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    prsGroup!KeyField.Properties("Optimize") = True
    prsGroup.Sort = "Name ASC"
    
    Exit Function
   
CleanUp:
   
    Call FreeRecordset(prsGroup)
   
    Exit Function
   
ErrorTrap:
    
    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetGroupRecordset() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    GetGroupRecordset = False
    
    Resume CleanUp

End Function

'******************************************************************************
' LookupGroup
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  Search for LookupKeyField in prsGroup. If it is found, point at the
'         prsGroup record that was found, then return true and GroupName;
'         if it is not found (which should never happen), then return false and
'         blank.
'******************************************************************************
Public Function LookupGroup(ByVal LookupKeyField As Integer, ByVal LookupType As Integer, ByRef GroupName As String) As Boolean
    
    Dim intIndex As Integer
    
    LookupGroup = False
    GroupName = ""
    intIndex = 1

    With frmMainMenu
        
        prsGroup.MoveFirst
        Do Until prsGroup.EOF
            
            Select Case LookupType
                Case LOOKUP_BY_LISTINDEX:
                    If LookupKeyField = intIndex Then
                        GroupName = prsGroup![Name]
                        LookupGroup = True
                        Exit Function
                    End If
                Case LOOKUP_BY_KEYFIELD:
                    If LookupKeyField = prsGroup![KeyField] Then
                        GroupName = prsGroup![Name]
                        LookupGroup = True
                        Exit Function
                    End If
            End Select
            
            intIndex = intIndex + 1
            prsGroup.MoveNext
        Loop
    
    End With

    ' If the group had been found, we would have previously exitted.
    ' Therefore, an error condition exists.
    
    strErrmsg = "LookupGroup() " & vbCrLf & vbCrLf & _
                "Group " & LookupKeyField & " not found."

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

End Function

'******************************************************************************
' HasAssignedSquadron
'
' INPUT:  n/a
'
' OUTPUT: Comma-delimited list of squadrons.
'
' RETURN: True if squadrons are assigned, otherwise false.
'
' NOTES:  Determine if the current group has any squadrons assigned to it.
'******************************************************************************
Private Function HasAssignedSquadron(ByRef strReturn As String) As Boolean
    Dim intSecondLastCommaLoc As Integer
    Dim intLastCommaLoc As Integer
    Dim strTemp As String
    Dim intIndex As Integer

    HasAssignedSquadron = False

    intSecondLastCommaLoc = 0
    intLastCommaLoc = 0
    strReturn = ""
    strTemp = ""
    
    ' Determine if any squadrons are assigned to the group.
    
    prsSquadron.MoveFirst
    Do Until prsSquadron.EOF
    
        If prsSquadron![Group] = prsGroup![KeyField] Then
            HasAssignedSquadron = True
            
            strReturn = strReturn & prsSquadron![Name] & ", "
            
            intSecondLastCommaLoc = intLastCommaLoc
            
            intLastCommaLoc = Len(strReturn) - 1
        End If
            
        prsSquadron.MoveNext
    Loop
    
    ' If there are dependent squadrons, correct the return string's grammar.
        
    If strReturn <> "" Then
        
        For intIndex = 1 To intLastCommaLoc
            
            If intIndex = intSecondLastCommaLoc Then
                ' Change the second last comma to " and".
                strTemp = strTemp & " and"
            ElseIf intIndex <> intLastCommaLoc Then
                ' Copy every letter but the trailing ", ".
                strTemp = strTemp & Mid(strReturn, intIndex, 1)
            End If
        
        Next intIndex
   
        strReturn = strTemp
    
    End If
    
    ' Re-point the squadron recordset to the record currently on the squadron tab.

    prsSquadron.Bookmark = varSquadronCurrentlyOnTab

End Function

'******************************************************************************
' PopulateGroupCombos
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  There are two group combos: One on the group tab and one on the
'         squadron tab.
'******************************************************************************
Public Sub PopulateGroupCombos()
    With frmMainMenu
        
        prsGroup.MoveFirst
        Do Until prsGroup.EOF
                
            .cboName(GROUP_TAB).AddItem prsGroup![Name]
            .cboGroup.AddItem prsGroup![Name] ' SQUADRON_TAB
                
            prsGroup.MoveNext
        
        Loop
    
    End With
End Sub

'******************************************************************************
' AddGroup
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Add a new group. Name is a required field: It cannot be blank, and it
'         cannot be a duplicate name. If nothing else is changed from the
'         current group, then a new group with the same commander and base will
'         be created. User created groups are never default groups.
'******************************************************************************
Public Function AddGroup() As Boolean
    On Error GoTo ErrorTrap
    
    Dim intKeyField As Integer
    Dim strIgnore As String
    
    AddGroup = True

    With frmMainMenu
        If ValidateRequiredInput(.cboName(GROUP_TAB)) = False Then
            ' This is not a severe system error, so return true.
            Exit Function
        End If
        
        If IsDupeGroupName() = True Then
            strErrmsg = "Failed to add group." & vbCrLf & vbCrLf & _
                        "The " & _
                        .cboName(GROUP_TAB).Text & _
                        " already exists."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
        
        intKeyField = NextKeyField(prsGroup, "Group")
        
        If intKeyField = 0 Then
            ' Fatal error. A value > 0 should always be returned.
' qwe            Call ExitEmulator
            AddGroup = False
            Exit Function
        End If
        
        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
        
            prsGroup.AddNew
            
            prsGroup![KeyField] = intKeyField
            prsGroup![Name] = .cboName(GROUP_TAB).Text
            prsGroup![Sorties] = 0
            prsGroup![Kills] = 0
            prsGroup![PlanesLost] = 0
            prsGroup![KIA] = 0
            prsGroup![MIA] = 0
            prsGroup![Wounded] = 0
            prsGroup![POW] = 0
            prsGroup![MedalOfHonor] = 0
            prsGroup![DistinguishedServiceCross] = 0
            prsGroup![SilverStar] = 0
            prsGroup![DistinguishedFlyingCross] = 0
            prsGroup![BronzeStarV] = 0
            prsGroup![PurpleHeart] = 0
            prsGroup![AirMedal] = 0
            prsGroup![DistinguishedUnitCitation] = 0
            prsGroup![MeritoriousUnitCitation] = 0
            
            If LookupAirman((.cboCommander(GROUP_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = False Then
' qwe            Call ExitEmulator
                AddGroup = False
                Exit Function
            Else
                prsGroup![Commander] = prsAirman![KeyField]
            
                ' Repoint the airman recordset to the record displayed on the airman tab.
                prsAirman.Bookmark = varAirmanCurrentlyOnTab
            End If

            If .optEngland(GROUP_TAB).Value = True Then
                prsGroup![Base] = ENGLAND_TER
            Else ' .optItaly(GROUP_TAB).Value = True
                prsGroup![Base] = ITALY_TER
            End If
        
            prsGroup![Default] = vbUnchecked
            
            prsGroup.UpdateBatch
    
        pobjConn.CommitTrans
        
        pintOpenTrans = pintOpenTrans - 1
            
        ' Insert into group tab name combo, then change the group tab to
        ' the new group. Also add the new group to the squadron tab group
        ' combo. (Though it does not need to be re-indexed because no
        ' squadrons, including the current one, are dependent on the new
        ' group.)

        .cboName(GROUP_TAB).AddItem prsGroup![Name], (prsGroup.AbsolutePosition - 1)
        .cboGroup.AddItem prsGroup![Name], (prsGroup.AbsolutePosition - 1)

        .cboName(GROUP_TAB).ListIndex = (prsGroup.AbsolutePosition - 1)

'        If LookupGroup((.cboName(GROUP_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = True Then
'            ' All we need is the pointer to the matching record (the
'            ' Group name was selected). Fill in the tab fields.
'            If FillGroupTabFields() = False Then
'                Call ExitEmulator
'            End If
'        Else
'            Call ExitEmulator
'        End If

    End With

    Exit Function

CleanUp:

    If pintOpenTrans Then
        pobjConn.RollbackTrans
        pintOpenTrans = pintOpenTrans - 1
    End If
    
    Exit Function

ErrorTrap:

    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "AddGroup() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

    Err.Clear

    AddGroup = False
    
    Resume CleanUp

End Function

'******************************************************************************
' ModifyGroup
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Update the current group. The group's name may not be changed, nor
'         may it be blank.
'******************************************************************************
Public Function ModifyGroup() As Boolean
     On Error GoTo ErrorTrap
 
    Dim strIgnore As String
 
    ModifyGroup = True
    
    With frmMainMenu
        
        ' Default groups cannot be deleted.
        
        If prsGroup![Default] = True Then
            strErrmsg = "Failed to update group." & vbCrLf & vbCrLf & _
                        prsGroup![Name] & _
                        " is a default group."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
    
        If ValidateRequiredInput(.cboName(GROUP_TAB)) = False Then
            Exit Function
        ElseIf .cboName(GROUP_TAB).Text <> prsGroup![Name] Then
            strErrmsg = "Failed to update group." & vbCrLf & vbCrLf & _
                        "You are not allowed to change the " & _
                        prsGroup![Name] & _
                        "'s name."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
        
        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
        
            If LookupAirman((.cboCommander(GROUP_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = False Then
' qwe            Call ExitEmulator
                ModifyGroup = False
                Exit Function
            Else
                prsGroup![Commander] = prsAirman![KeyField]
            
                ' Repoint the airman recordset to the record displayed on the airman tab.
                prsAirman.Bookmark = varAirmanCurrentlyOnTab
            End If

            If .optEngland(GROUP_TAB).Value = True Then
                prsGroup![Base] = ENGLAND_TER
            Else '.optItaly(GROUP_TAB).Value = True
                prsGroup![Base] = ITALY_TER
            End If
        
            prsGroup.UpdateBatch
    
        pobjConn.CommitTrans
        
        pintOpenTrans = pintOpenTrans - 1
            
    End With

    Exit Function

CleanUp:

    If pintOpenTrans Then
        pobjConn.RollbackTrans
        pintOpenTrans = pintOpenTrans - 1
    End If
    
    Exit Function

ErrorTrap:

    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "ModifyGroup() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

    Err.Clear

    ModifyGroup = False
    
    Resume CleanUp

End Function

'******************************************************************************
' DeleteGroup
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Delete the current group. If the group is a default group, it cannot
'         be deleted. If there is a dependency, pop a msgbox, then return true.
'         (It is not a system error.) Due to defaults being impervious, it is
'         not possible to delete all records, therefore there will always be
'         some records remaining in the recordset and combo.
'******************************************************************************
Public Function DeleteGroup() As Boolean
    On Error GoTo ErrorTrap
    
    Dim strSquadronList As String
    Dim strIgnore As String
    Dim intListIndex As Integer
    
    DeleteGroup = True

    With frmMainMenu
        
        ' Default groups cannot be deleted.
        
        If prsGroup![Default] = True Then
            strErrmsg = "Failed to delete group." & vbCrLf & vbCrLf & _
                        prsGroup![Name] & _
                        " is a default group."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
    
        ' Determine if the group has assigned squadrons.
        
        If HasAssignedSquadron(strSquadronList) = True Then
            strErrmsg = "Failed to delete group." & vbCrLf & vbCrLf & _
                        strSquadronList & _
                        " are assigned to the " & _
                        prsGroup![Name] & _
                        ". A group cannot be deleted if it has " & _
                        "dependent squadrons."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
    
        ' Confirm the deletion.
    
        If MsgBox("Are you sure?", (vbYesNo + vbDefaultButton2 + vbQuestion)) = vbNo Then
            Exit Function
        End If

        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
    
            prsGroup.Delete
    
            prsGroup.UpdateBatch
        
        pobjConn.CommitTrans
            
        pintOpenTrans = pintOpenTrans - 1
            
        ' Delete from group tab name combo, then change the group tab to
        ' another listed group. Also delete from squadron tab group combo.
        ' (Though it does not need to be re-indexed because no squadrons,
        ' including the current one, were dependent on the deleted group.)
    
        intListIndex = .cboName(GROUP_TAB).ListIndex
        
        .cboName(GROUP_TAB).RemoveItem intListIndex
        .cboGroup.RemoveItem intListIndex

        If intListIndex = 0 Then
            
            ' Deleted the first listed record. Point to the second record,
            ' which just became the first record.
            .cboName(GROUP_TAB).ListIndex = 0
        
        ElseIf intListIndex = .cboName(GROUP_TAB).ListCount Then
            
            ' Deleted the last listed record. Point to the second-to-last
            ' record, which just became the last record. ListIndex is 0-base,
            ' while ListCount is 1-base, so we need to subtract 1.
            .cboName(GROUP_TAB).ListIndex = intListIndex - 1 '(.cboName(GROUP_TAB).ListCount - 1)
        
        Else
            
            ' Deleted some record between the first and last records. Point
            ' to the record after the one that was deleted.
            .cboName(GROUP_TAB).ListIndex = intListIndex
        
        End If
    
        If LookupGroup((.cboName(GROUP_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = True Then
            If FillGroupTabFields() = False Then
' qwe            Call ExitEmulator
                DeleteGroup = False
                Exit Function
            End If
        Else
' qwe            Call ExitEmulator
            DeleteGroup = False
            Exit Function
        End If

    End With

    Exit Function

CleanUp:

    If pintOpenTrans Then
        pobjConn.RollbackTrans
        pintOpenTrans = pintOpenTrans - 1
    End If
    
    Exit Function

ErrorTrap:

    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "DeleteGroup() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

    Err.Clear

    DeleteGroup = False
    
    Resume CleanUp

End Function

'******************************************************************************
' IsDupeGroupName
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Multiple groups may not have the exact same name.
'******************************************************************************
Private Function IsDupeGroupName() As Boolean

    IsDupeGroupName = False

    With frmMainMenu
    
        ' Determine if any other group has the same name.
         
        prsGroup.MoveFirst
        Do Until prsGroup.EOF
         
            If prsGroup![Name] = .cboName(GROUP_TAB).Text Then
                IsDupeGroupName = True
                Exit Do
            End If
        
            prsGroup.MoveNext
        Loop
    
    End With
    
    ' Re-point the group recordset to the record currently on the group tab.

    prsGroup.Bookmark = varGroupCurrentlyOnTab

End Function



