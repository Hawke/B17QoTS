'******************************************************************************
' modSquadron.bas
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

Attribute VB_Name = "modSquadron"
Option Explicit

Public prsSquadron As New ADODB.Recordset
Public prsBomberSquadron As New ADODB.Recordset
Public varSquadronCurrentlyOnTab As Variant

Dim strErrmsg As String

'******************************************************************************
' FillSquadronTabFields
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if the fields were filled, otherwise false.
'
' NOTES:  Some prsSquadron record must be pointed at -- either using MoveFirst
'         or a LookupSquadron() call -- before this function is called.
'******************************************************************************
Public Function FillSquadronTabFields() As Boolean
    
    Dim strCommander As String
    Dim strGroup As String
    
    FillSquadronTabFields = True
    
    With frmMainMenu
        .txtKeyField(SQUADRON_TAB).Text = prsSquadron![KeyField]
        .cboName(SQUADRON_TAB).Text = prsSquadron![Name]
        .txtSorties(SQUADRON_TAB).Text = prsSquadron![Sorties]
        .txtKills(SQUADRON_TAB).Text = prsSquadron![Kills]
        .txtPlanesLost(SQUADRON_TAB).Text = prsSquadron![PlanesLost]
        .txtKIA(SQUADRON_TAB).Text = prsSquadron![KIA]
        .txtMIA(SQUADRON_TAB).Text = prsSquadron![MIA]
        .txtWounded(SQUADRON_TAB).Text = prsSquadron![Wounded]
        .txtPOW(SQUADRON_TAB).Text = prsSquadron![POW]
        .txtMedalOfHonor(SQUADRON_TAB).Text = prsSquadron![MedalOfHonor]
        .txtDistinguishedServiceCross(SQUADRON_TAB).Text = prsSquadron![DistinguishedServiceCross]
        .txtSilverStar(SQUADRON_TAB).Text = prsSquadron![SilverStar]
        .txtDistinguishedFlyingCross(SQUADRON_TAB).Text = prsSquadron![DistinguishedFlyingCross]
        .txtBronzeStarV(SQUADRON_TAB).Text = prsSquadron![BronzeStarV]
        .txtPurpleHeart(SQUADRON_TAB).Text = prsSquadron![PurpleHeart]
        .txtAirMedal(SQUADRON_TAB).Text = prsSquadron![AirMedal]
        .txtDistinguishedUnitCitation(SQUADRON_TAB).Text = prsSquadron![DistinguishedUnitCitation]
        .txtMeritoriousUnitCitation(SQUADRON_TAB).Text = prsSquadron![MeritoriousUnitCitation]
    
        Select Case prsSquadron![BomberType]
            Case B17_TYPE:
                .optB17FlyingFortress.Value = True
            Case B24_TYPE:
                .optB24Liberator.Value = True
            Case AVRO_TYPE:
                .optAvroLancaster.Value = True
        End Select
        
        ' Populate the recordset lookup fields.
        
        If LookupAirman(prsSquadron![Commander], LOOKUP_BY_KEYFIELD, strCommander) = False Then
            FillSquadronTabFields = False
            Exit Function
        Else
            .cboCommander(SQUADRON_TAB).Text = strCommander
            
            ' Repoint the airman recordset to the record displayed on the airman tab.
            prsAirman.Bookmark = varAirmanCurrentlyOnTab
        End If

        If LookupGroup(prsSquadron![Group], LOOKUP_BY_KEYFIELD, strGroup) = False Then
            FillSquadronTabFields = False
            Exit Function
        Else
            .cboGroup.Text = strGroup
            
            ' Repoint the group recordset to the record displayed on the group tab.
            prsGroup.Bookmark = varGroupCurrentlyOnTab
        End If
        
        ' If this is a default Squadron, disable the fields, otherwise ensure
        ' they are enabled.
        
        If prsSquadron![Default] = True Then
            .chkDefault(SQUADRON_TAB).Value = vbChecked
        
            .cboCommander(SQUADRON_TAB).Enabled = False
            .cboCommander(SQUADRON_TAB).BackColor = vbButtonFace
            .cboGroup.Enabled = False
            .cboGroup.BackColor = vbButtonFace
            .optB17FlyingFortress.Enabled = False
            .optB24Liberator.Enabled = False
            .optAvroLancaster.Enabled = False
        Else
            .chkDefault(SQUADRON_TAB).Value = vbUnchecked
        
            .cboCommander(SQUADRON_TAB).Enabled = True
            .cboCommander(SQUADRON_TAB).BackColor = vbWhite
            .cboGroup.Enabled = True
            .cboGroup.BackColor = vbWhite
            .optB17FlyingFortress.Enabled = True
            .optB24Liberator.Enabled = True
            .optAvroLancaster.Enabled = True
        End If

    End With

End Function

'******************************************************************************
' ZeroSquadronTabFields
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
Public Sub ZeroSquadronTabFields()
    
    With frmMainMenu
        
        .txtKeyField(SQUADRON_TAB).Text = 0
        .txtSorties(SQUADRON_TAB).Text = 0
        .txtKills(SQUADRON_TAB).Text = 0
        .txtPlanesLost(SQUADRON_TAB).Text = 0
        .txtKIA(SQUADRON_TAB).Text = 0
        .txtWounded(SQUADRON_TAB).Text = 0
        .txtPOW(SQUADRON_TAB).Text = 0
        .txtMedalOfHonor(SQUADRON_TAB).Text = 0
        .txtDistinguishedServiceCross(SQUADRON_TAB).Text = 0
        .txtSilverStar(SQUADRON_TAB).Text = 0
        .txtDistinguishedFlyingCross(SQUADRON_TAB).Text = 0
        .txtBronzeStarV(SQUADRON_TAB).Text = 0
        .txtPurpleHeart(SQUADRON_TAB).Text = 0
        .txtAirMedal(SQUADRON_TAB).Text = 0
        .txtDistinguishedUnitCitation(SQUADRON_TAB).Text = 0
        .txtMeritoriousUnitCitation(SQUADRON_TAB).Text = 0
        
        .optB17FlyingFortress.Value = True
        
        .cboCommander(SQUADRON_TAB).ListIndex = 0
            
        .cboGroup.ListIndex = 0
            
        .chkDefault(SQUADRON_TAB).Value = vbUnchecked
                    
        .cboCommander(SQUADRON_TAB).Enabled = True
        .cboCommander(SQUADRON_TAB).BackColor = vbWhite
        
        .cboGroup.Enabled = True
        .cboGroup.BackColor = vbWhite
        
        .optB17FlyingFortress.Enabled = True
        .optB24Liberator.Enabled = True
        .optAvroLancaster.Enabled = True
    End With
End Sub

'******************************************************************************
' GetSquadronRecordset
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetSquadronRecordset() As Boolean
    On Error GoTo ErrorTrap
   
    GetSquadronRecordset = True

    pobjCmnd.CommandText = "SELECT * FROM Squadron ORDER BY Name"

    prsSquadron.CursorLocation = adUseClient
    prsSquadron.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    prsSquadron!KeyField.Properties("Optimize") = True
    prsSquadron.Sort = "Name ASC"
    
    Exit Function
   
CleanUp:
   
    Call FreeRecordset(prsSquadron)
   
    Exit Function
   
ErrorTrap:
    
    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetSquadronRecordset() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    GetSquadronRecordset = False
    
    Resume CleanUp

End Function

'******************************************************************************
' LookupSquadron
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  Search for LookupKeyField in prsSquadron. If it is found, point at the
'         prsSquadron record that was found, then return true and SquadronName;
'         if it is not found (which should never happen), then return false and
'         blank.
'******************************************************************************
Public Function LookupSquadron(ByVal LookupKeyField As Integer, ByVal LookupType As Integer, ByRef SquadronName As String) As Boolean
    
    Dim intIndex As Integer
    
    LookupSquadron = False
    SquadronName = ""
    intIndex = 1

    With frmMainMenu
        
        prsSquadron.MoveFirst
        Do Until prsSquadron.EOF
            
            Select Case LookupType
                Case LOOKUP_BY_LISTINDEX:
                    If LookupKeyField = intIndex Then
                        SquadronName = prsSquadron![Name]
                        LookupSquadron = True
                        Exit Function
                    End If
                Case LOOKUP_BY_KEYFIELD:
                    If LookupKeyField = prsSquadron![KeyField] Then
                        SquadronName = prsSquadron![Name]
                        LookupSquadron = True
                        Exit Function
                    End If
            End Select
            
            intIndex = intIndex + 1
            prsSquadron.MoveNext
        Loop
    
    End With

    ' If the squadron had been found, we would have previously exitted.
    ' Therefore, an error condition exists.
    
    strErrmsg = "LookupSquadron() " & vbCrLf & vbCrLf & _
                "Squadron " & LookupKeyField & " not found."

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

End Function

'******************************************************************************
' HasAssignedBomber
'
' INPUT:  n/a
'
' OUTPUT: Comma-delimited list of bombers.
'
' RETURN: True if bombers are assigned, otherwise false.
'
' NOTES:  Determine if the current squadron has any bombers assigned to it.
'******************************************************************************
Private Function HasAssignedBomber(ByRef strReturn As String) As Boolean
    Dim intSecondLastCommaLoc As Integer
    Dim intLastCommaLoc As Integer
    Dim strTemp As String
    Dim intIndex As Integer

    HasAssignedBomber = False

    intSecondLastCommaLoc = 0
    intLastCommaLoc = 0
    strReturn = ""
    strTemp = ""

    ' Determine if any bombers are assigned to the squadron.
    
    prsBomber.MoveFirst
    Do Until prsBomber.EOF
    
        If prsBomber![Squadron] = prsSquadron![KeyField] Then
            HasAssignedBomber = True
            
            strReturn = strReturn & prsBomber![Name] & ", "
            
            intSecondLastCommaLoc = intLastCommaLoc
            
            intLastCommaLoc = Len(strReturn) - 1
        End If
            
        prsBomber.MoveNext
    Loop
    
    ' If there are dependent bombers, correct the return string's grammar.
        
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
    
    ' Re-point the recordset to the record currently on the bomber tab.

    prsBomber.Bookmark = varBomberCurrentlyOnTab

End Function

'******************************************************************************
' PopulateSquadronCombos
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  There are two squadron combos: One on the squadron tab and one on the
'         bomber tab.
'******************************************************************************
Public Sub PopulateSquadronCombos()
    With frmMainMenu

        prsSquadron.MoveFirst
        Do Until prsSquadron.EOF
                
            .cboName(SQUADRON_TAB).AddItem prsSquadron![Name]
            .cboSquadron.AddItem prsSquadron![Name] ' BOMBER_TAB
                
            prsSquadron.MoveNext
        
        Loop
    
    End With
End Sub

'******************************************************************************
' AddSquadron
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Add a new squadron. Name is a required field: It cannot be blank, and
'         it cannot be a duplicate name. If nothing else is changed from the
'         current squadron, then a new squadron with the same commander, group
'         and bomber type will be created. User created squadrons are never
'         default squadrons.
'******************************************************************************
Public Function AddSquadron() As Boolean
    On Error GoTo ErrorTrap
    
    Dim intKeyField As Integer
    Dim strIgnore As String
    
    AddSquadron = True

    With frmMainMenu
        If ValidateRequiredInput(.cboName(SQUADRON_TAB)) = False Then
            ' This is not a severe system error, so return true.
            Exit Function
        End If
        
        If IsDupeSquadronName() = True Then
            strErrmsg = "Failed to add squadron." & vbCrLf & vbCrLf & _
                        "The " & _
                        .cboName(SQUADRON_TAB).Text & _
                        " already exists."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
        
        intKeyField = NextKeyField(prsSquadron, "Squadron")
        
        If intKeyField = 0 Then
            ' Fatal error. A value > 0 should always be returned.
' qwe            Call ExitEmulator
            AddSquadron = False
            Exit Function
        End If
        
        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
        
            prsSquadron.AddNew
            
            prsSquadron![KeyField] = intKeyField
            prsSquadron![Name] = .cboName(SQUADRON_TAB).Text
            prsSquadron![Sorties] = 0
            prsSquadron![Kills] = 0
            prsSquadron![PlanesLost] = 0
            prsSquadron![KIA] = 0
            prsSquadron![MIA] = 0
            prsSquadron![Wounded] = 0
            prsSquadron![POW] = 0
            prsSquadron![MedalOfHonor] = 0
            prsSquadron![DistinguishedServiceCross] = 0
            prsSquadron![SilverStar] = 0
            prsSquadron![DistinguishedFlyingCross] = 0
            prsSquadron![BronzeStarV] = 0
            prsSquadron![PurpleHeart] = 0
            prsSquadron![AirMedal] = 0
            prsSquadron![DistinguishedUnitCitation] = 0
            prsSquadron![MeritoriousUnitCitation] = 0
            
            If LookupAirman((.cboCommander(SQUADRON_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = False Then
' qwe            Call ExitEmulator
                AddSquadron = False
                Exit Function
            Else
                prsSquadron![Commander] = prsAirman![KeyField]
            
                ' Repoint the airman recordset to the record displayed on the airman tab.
                prsAirman.Bookmark = varAirmanCurrentlyOnTab
            End If

            If LookupGroup((.cboGroup.ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = False Then
' qwe            Call ExitEmulator
                AddSquadron = False
                Exit Function
            Else
                prsSquadron![Group] = prsGroup![KeyField]
            
                ' Repoint the group recordset to the record displayed on the group tab.
                prsGroup.Bookmark = varGroupCurrentlyOnTab
            End If

            If .optB17FlyingFortress.Value = True Then
                prsSquadron![BomberType] = B17_TYPE
            ElseIf .optB24Liberator.Value = True Then
                prsSquadron![BomberType] = B24_TYPE
            Else ' .optAvroLancaster.Value = True
                prsSquadron![BomberType] = AVRO_TYPE
            End If
        
            prsSquadron![Default] = vbUnchecked
            
            prsSquadron.UpdateBatch
    
        pobjConn.CommitTrans
        
        pintOpenTrans = pintOpenTrans - 1
            
        ' Insert into squadron tab name combo, then change the squadron tab
        ' to the new squadron. Also add the new squadron to the bomber tab
        ' squadron combo. (Though it does not need to be re-indexed because
        ' no bombers, including the current one, are dependent on the new
        ' squadron.)

        .cboName(SQUADRON_TAB).AddItem prsSquadron![Name], (prsSquadron.AbsolutePosition - 1)

' xcv       .cboSquadron.AddItem prsSquadron![Name], (prsSquadron.AbsolutePosition - 1)
Call PopulateBomberSquadronCombo ' xcv
        
        .cboName(SQUADRON_TAB).ListIndex = (prsSquadron.AbsolutePosition - 1)

'        If LookupSquadron((.cboName(SQUADRON_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = True Then
'            ' All we need is the pointer to the matching record (the
'            ' Squadron name was selected). Fill in the tab fields.
'            If FillSquadronTabFields() = False Then
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
                "AddSquadron() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

    Err.Clear

    AddSquadron = False
    
    Resume CleanUp

End Function

'******************************************************************************
' ModifySquadron
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Update the current squadron. The squadron's name may not be changed,
'         nor may it be blank.
'******************************************************************************
Public Function ModifySquadron() As Boolean
     On Error GoTo ErrorTrap
 
    Dim strIgnore As String
 
    ModifySquadron = True
    
    With frmMainMenu
        
        ' Default squadrons cannot be updated.
        
        If prsSquadron![Default] = True Then
            strErrmsg = "Failed to update squadron." & vbCrLf & vbCrLf & _
                        prsSquadron![Name] & _
                        " is a default squadron."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
    
        If ValidateRequiredInput(.cboName(SQUADRON_TAB)) = False Then
            Exit Function
        ElseIf .cboName(SQUADRON_TAB).Text <> prsSquadron![Name] Then
            strErrmsg = "Failed to update squadron." & vbCrLf & vbCrLf & _
                        "You are not allowed to change the " & _
                        prsSquadron![Name] & _
                        "'s name."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
        
        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
        
            If LookupAirman((.cboCommander(SQUADRON_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = False Then
' qwe            Call ExitEmulator
                ModifySquadron = False
                Exit Function
            Else
                prsSquadron![Commander] = prsAirman![KeyField]
            
                ' Repoint the airman recordset to the record displayed on the airman tab.
                prsAirman.Bookmark = varAirmanCurrentlyOnTab
            End If

            If LookupGroup((.cboGroup.ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = False Then
' qwe            Call ExitEmulator
                ModifySquadron = False
                Exit Function
            Else
                prsSquadron![Group] = prsGroup![KeyField]
            
                ' Repoint the group recordset to the record displayed on the group tab.
                prsGroup.Bookmark = varGroupCurrentlyOnTab
            End If

            If .optB17FlyingFortress.Value = True Then
                prsSquadron![BomberType] = B17_TYPE
            ElseIf .optB24Liberator.Value = True Then
                prsSquadron![BomberType] = B24_TYPE
            Else ' .optAvroLancaster.Value = True
                prsSquadron![BomberType] = AVRO_TYPE
            End If
        
            prsSquadron.UpdateBatch
    
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
                "ModifySquadron() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

    Err.Clear

    ModifySquadron = False
    
    Resume CleanUp

End Function

'******************************************************************************
' DeleteSquadron
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Delete the current squadron. If the squadron is a default squadron, it
'         cannot be deleted. If there is a dependency, pop a msgbox, then return
'         true. (It is not a system error.) Due to defaults being impervious, it
'         is not possible to delete all records, therefore there will always be
'         some records remaining in the recordset and combo.
'******************************************************************************
Public Function DeleteSquadron() As Boolean
    On Error GoTo ErrorTrap
    
    Dim strBomberList As String
    Dim strIgnore As String
    Dim intListIndex As Integer
    
    DeleteSquadron = True

    With frmMainMenu
        
        ' Default squadrons cannot be deleted.
        
        If prsSquadron![Default] = True Then
            strErrmsg = "Failed to delete squadron." & vbCrLf & vbCrLf & _
                        prsSquadron![Name] & _
                        " is a default squadron."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
    
        ' Determine if the squadron has assigned bombers.
        
        If HasAssignedBomber(strBomberList) = True Then
            strErrmsg = "Failed to delete squadron." & vbCrLf & vbCrLf & _
                        strBomberList & _
                        " are assigned to the " & _
                        prsSquadron![Name] & _
                        ". A squadron cannot be deleted if it has " & _
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
    
            prsSquadron.Delete
    
            prsSquadron.UpdateBatch
        
        pobjConn.CommitTrans
            
        pintOpenTrans = pintOpenTrans - 1
            
        ' Delete from squadron tab name combo, then change the squadron tab to
        ' another listed squadron. Also delete from bomber tab squadron combo.
        ' (Though it does not need to be re-indexed because no bombers,
        ' including the current one, were dependent on the deleted squadron.)
    
        intListIndex = .cboName(SQUADRON_TAB).ListIndex
        
        .cboName(SQUADRON_TAB).RemoveItem intListIndex
        
' xcv        .cboSquadron.RemoveItem intListIndex
Call PopulateBomberSquadronCombo ' xcv
        
        If intListIndex = 0 Then
            
            ' Deleted the first listed record. Point to the second record,
            ' which just became the first record.
            .cboName(SQUADRON_TAB).ListIndex = 0
        
        ElseIf intListIndex = .cboName(SQUADRON_TAB).ListCount Then
            
            ' Deleted the last listed record. Point to the second-to-last
            ' record, which just became the last record. ListIndex is 0-base,
            ' while ListCount is 1-base, so we need to subtract 1.
            .cboName(SQUADRON_TAB).ListIndex = intListIndex - 1 '(.cboName(SQUADRON_TAB).ListCount - 1)
        
        Else
            
            ' Deleted some record between the first and last records. Point
            ' to the record after the one that was deleted.
            .cboName(SQUADRON_TAB).ListIndex = intListIndex
        
        End If
    
        If LookupSquadron((.cboName(SQUADRON_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = True Then
            If FillSquadronTabFields() = False Then
' qwe            Call ExitEmulator
                DeleteSquadron = False
                Exit Function
            End If
        Else
' qwe            Call ExitEmulator
            DeleteSquadron = False
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
                "DeleteSquadron() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

    Err.Clear

    DeleteSquadron = False
    
    Resume CleanUp

End Function

'******************************************************************************
' IsDupeSquadronName
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Multiple squadrons may not have the exact same name.
'******************************************************************************
Private Function IsDupeSquadronName() As Boolean

    IsDupeSquadronName = False

    With frmMainMenu
    
        ' Determine if any other squadron has the same name.
         
        prsSquadron.MoveFirst
        Do Until prsSquadron.EOF
         
            If prsSquadron![Name] = .cboName(SQUADRON_TAB).Text Then
                IsDupeSquadronName = True
                Exit Do
            End If
        
            prsSquadron.MoveNext
        Loop
    
    End With
    
    ' Re-point the squadron recordset to the record currently on the squadron tab.

    prsSquadron.Bookmark = varSquadronCurrentlyOnTab

End Function

'******************************************************************************
' PopulateBomberSquadronCombo
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  The list of available squadrons is restricted to squadrons of the
'         same bomber type as the bomber.
'******************************************************************************
Public Sub PopulateBomberSquadronCombo()
    Dim strTypeFilter As String
    Dim intIndex As Integer
    Dim strIgnore As String

    With frmMainMenu
        
         ' Clone the squadron recordset so we do not affect the squadron tab.
       
        Set prsBomberSquadron = prsSquadron.Clone
    
        ' Delete the previous combo items to prevent concatenating two sets
        ' of squadrons.
        
        For intIndex = 0 To (.cboSquadron.ListCount - 1)
            .cboSquadron.RemoveItem 0
        Next intIndex
        
        ' Filter the clone so that it only contains squadrons which have the
        ' same general bomber type as the specific bomber model.
        
        Select Case (.cboBomberModel(BOMBER_TAB).ListIndex + 1)
            
            Case B17_C To YB40:
                
                strTypeFilter = "BomberType = " & B17_TYPE
            
            Case B24_D To B24_LM:
                
                strTypeFilter = "BomberType = " & B24_TYPE
            
            Case AVRO_LANCASTER:
                
                strTypeFilter = "BomberType = " & AVRO_TYPE
        
        End Select
    
        prsBomberSquadron.Filter = strTypeFilter
        
        prsBomberSquadron.MoveFirst
        Do Until prsBomberSquadron.EOF
                
            .cboSquadron.AddItem prsBomberSquadron![Name] ' BOMBER_TAB
                
            prsBomberSquadron.MoveNext
        Loop
Dim a
a = 1
' Scroll to the squadron to which the bomber belongs.

' We know the keyfield of the bomber's squadron, but we do not know where in
' the combo list the squadron is located. (Due to the combo not listing all
' squadrons, only those of the same bomber type as the bomber.)

' Get the name of the bomber's squadron
Dim strSquadron As String

'        ' If the bomber name is only one character long, then the user just
'        ' started to add a new bomber, by typing in its name. Don't lookup
'        ' the new bomber's squadron, because it doesn't have one yet.
'
'MsgBox (prsBomberSquadron.RecordCount & " records in set")
'MsgBox (.cboSquadron.ListCount & " records in combo")
'
'        If Len(.cboName(BOMBER_TAB)) >= 2 Then
        If .cboSquadron.ListCount = 1 Then
            
            ' If there is only one record in the recordset/combo, there is
            ' no sense in doing a lookup. Just point at the first/only row.
            
            .cboSquadron.ListIndex = 0
        
        Else

            If LookupBomberSquadron(prsBomber![Squadron], LOOKUP_BY_KEYFIELD, strSquadron) = False Then
                ' Don't worry about it: The situation is temporary.
            Else
                .cboSquadron.Text = strSquadron
    
                ' Repoint the squadron recordset to the record displayed on the squadron tab.
                prsSquadron.Bookmark = varSquadronCurrentlyOnTab
            End If
        
        End If

' Find the squadron in the combo.

'        For intIndex = 0 To (.cboSquadron.ListCount - 1)
'            .cboSquadron.RemoveItem 0
'        Next intIndex
    


    End With

End Sub

'******************************************************************************
' LookupBomberSquadron
'
' INPUT:  n/a
'
' OUTPUT: Squadron name if found, otherwise blank.
'
' RETURN: True if the squadron is found, otherwise false.
'
' NOTES:  Search for LookupKeyField in prsBomberSquadron.
'******************************************************************************
Public Function LookupBomberSquadron(ByVal LookupKeyField As Integer, ByVal LookupType As Integer, ByRef SquadronName As String) As Boolean
    
    Dim intIndex As Integer
    
    LookupBomberSquadron = False
    SquadronName = ""
    intIndex = 1

    With frmMainMenu
        
        prsBomberSquadron.MoveFirst
        Do Until prsBomberSquadron.EOF
            
            Select Case LookupType
                Case LOOKUP_BY_LISTINDEX:
                    If LookupKeyField = intIndex Then
                        SquadronName = prsBomberSquadron![Name]
                        LookupBomberSquadron = True
                        Exit Function
                    End If
                Case LOOKUP_BY_KEYFIELD:
                    If LookupKeyField = prsBomberSquadron![KeyField] Then
                        SquadronName = prsBomberSquadron![Name]
                        LookupBomberSquadron = True
                        Exit Function
                    End If
            End Select
            
            intIndex = intIndex + 1
            prsBomberSquadron.MoveNext
        Loop
    
    End With

    ' If the squadron had been found, we would have previously exitted.
    ' Therefore, an error condition exists.
    
    strErrmsg = "LookupBomberSquadron() " & vbCrLf & vbCrLf & _
                "Squadron " & LookupKeyField & " not found."

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

End Function



