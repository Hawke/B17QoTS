Attribute VB_Name = "modAirman"
'******************************************************************************
' modAirman.bas
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

Public prsAirman As New ADODB.Recordset
Public varAirmanCurrentlyOnTab As Variant

Dim strErrMsg As String

'******************************************************************************
' FillAirmanTabFields
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if the fields were filled, otherwise false.
'
' NOTES:  Some prsAirman record must be pointed at -- either using MoveFirst
'         or a LookupAirman() call -- before this function is called.
'******************************************************************************
Public Function FillAirmanTabFields() As Boolean
    
    Dim strRank As String
    Dim strAssignment As String
    Dim strCrewPosition As String
    Dim strAirmanStatus As String
    
    FillAirmanTabFields = True

    Call AdjustAvailableBombers ' Nov04

    With frmMainMenu
        ' Populate the non-lookup fields.
        
        .txtKeyField(AIRMAN_TAB).Text = prsAirman![KeyField]
        .cboName(AIRMAN_TAB).Text = prsAirman![Name]
        .cboRank.ListIndex = prsAirman![Rank] - 1
        .cboCrewPosition.ListIndex = prsAirman![CrewPosition] - 1
'        .cboAssignment.ListIndex = prsAirman![Assignment] ' kgreer (12 Dec 04) bug
        .cboAssignment.ListIndex = MapBomberKeyToAssignmentIndex(prsAirman![Assignment]) ' kgreer (12 Dec 04) bug
        .txtStatus(AIRMAN_TAB).Text = prsAirman![Status]
        .txtSorties(AIRMAN_TAB).Text = prsAirman![Sorties]
        .txtKills(AIRMAN_TAB).Text = prsAirman![Kills]
        .txtMedalOfHonor(AIRMAN_TAB).Text = prsAirman![MedalOfHonor]
        .txtDistinguishedServiceCross(AIRMAN_TAB).Text = prsAirman![DistinguishedServiceCross]
        .txtSilverStar(AIRMAN_TAB).Text = prsAirman![SilverStar]
        .txtDistinguishedFlyingCross(AIRMAN_TAB).Text = prsAirman![DistinguishedFlyingCross]
        .txtBronzeStarV(AIRMAN_TAB).Text = prsAirman![BronzeStarV]
        .txtPurpleHeart(AIRMAN_TAB).Text = prsAirman![PurpleHeart]
        .txtAirMedal(AIRMAN_TAB).Text = prsAirman![AirMedal]
        .txtDistinguishedUnitCitation(AIRMAN_TAB).Text = prsAirman![DistinguishedUnitCitation]
        .txtMeritoriousUnitCitation(AIRMAN_TAB).Text = prsAirman![MeritoriousUnitCitation]
    
        If prsAirman![Default] = True Then
            .chkDefault(AIRMAN_TAB).Value = vbChecked
        Else
            .chkDefault(AIRMAN_TAB).Value = vbUnchecked
        End If

        ' Populate the recordset lookup fields.
        If Not IsNull(prsAirman("Assignment")) Then
            If LookupBomber(prsAirman![Assignment], LOOKUP_BY_KEYFIELD, strAssignment) = False Then
                FillAirmanTabFields = False
                Exit Function
            Else
                .cboAssignment.Text = strAssignment
                
                ' Repoint the bomber recordset to the record displayed on the bomber tab.
                prsBomber.Bookmark = varBomberCurrentlyOnTab
            End If
        End If
        If LookupCrewPosition(prsAirman![CrewPosition], strCrewPosition) = False Then
            FillAirmanTabFields = False
            Exit Function
        Else
            .cboCrewPosition.Text = strCrewPosition
        End If

        If LookupAirmanStatus(prsAirman![Status], strAirmanStatus) = False Then
            FillAirmanTabFields = False
            Exit Function
        Else
            .txtStatus(AIRMAN_TAB).Text = strAirmanStatus
        End If

        ' If the airman is on duty status, enable the fields, otherwise
        ' ensure they are disabled. Also, set the status colors.
        
        Select Case prsAirman![Status]
            Case DUTY_STATUS:
        
                .cboRank.Enabled = True
                .cboRank.BackColor = vbWhite
                
                .cboAssignment.Enabled = True
                .cboAssignment.BackColor = vbWhite
                
                .cboCrewPosition.Enabled = True
                .cboCrewPosition.BackColor = vbWhite
                
                .txtStatus(AIRMAN_TAB).ForeColor = vbBlack
                .txtStatus(AIRMAN_TAB).BackColor = PaleGreen()
                
            Case LW1_STATUS, LW2_STATUS, SW_STATUS:
                
                .cboRank.Enabled = False
                .cboRank.BackColor = vbButtonFace
                
                .cboAssignment.Enabled = False
                .cboAssignment.BackColor = vbButtonFace
                
                .cboCrewPosition.Enabled = False
                .cboCrewPosition.BackColor = vbButtonFace
                
                .txtStatus(AIRMAN_TAB).ForeColor = vbBlack
                .txtStatus(AIRMAN_TAB).BackColor = PaleYellow()
            
            Case INVALID_STATUS, POW_STATUS, MIA_STATUS:
                
                .cboRank.Enabled = False
                .cboRank.BackColor = vbButtonFace
                
                .cboAssignment.Enabled = False
                .cboAssignment.BackColor = vbButtonFace
                
                .cboCrewPosition.Enabled = False
                .cboCrewPosition.BackColor = vbButtonFace
                
                .txtStatus(AIRMAN_TAB).ForeColor = vbBlack
                .txtStatus(AIRMAN_TAB).BackColor = PaleRed()
            
            Case DOW_STATUS, KIA_STATUS:
            
                .cboRank.Enabled = False
                .cboRank.BackColor = vbButtonFace
                
                .cboAssignment.Enabled = False
                .cboAssignment.BackColor = vbButtonFace
                
                .cboCrewPosition.Enabled = False
                .cboCrewPosition.BackColor = vbButtonFace
                
                .txtStatus(AIRMAN_TAB).ForeColor = vbWhite
                .txtStatus(AIRMAN_TAB).BackColor = vbBlack
            
            Case TOUR_COMPLETE_STATUS:
            
                .cboRank.Enabled = False
                .cboRank.BackColor = vbButtonFace
                
                .cboAssignment.Enabled = False
                .cboAssignment.BackColor = vbButtonFace
                
                .cboCrewPosition.Enabled = False
                .cboCrewPosition.BackColor = vbButtonFace
                
                .txtStatus(AIRMAN_TAB).ForeColor = vbBlack
                .txtStatus(AIRMAN_TAB).BackColor = vbButtonFace
            
        End Select
            
        ' Regardless of the airman's status, if he is a default airman,
        ' disable the fields.
        
        If .chkDefault(AIRMAN_TAB).Value = vbChecked Then
            .cboRank.Enabled = False
            .cboRank.BackColor = vbButtonFace
            .cboAssignment.Enabled = False
            .cboAssignment.BackColor = vbButtonFace
            .cboCrewPosition.Enabled = False
            .cboCrewPosition.BackColor = vbButtonFace
        End If
    
    End With

End Function

'******************************************************************************
' ZeroAirmanTabFields
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
Public Function ZeroAirmanTabFields()
    Dim strAirmanStatus As String
    
    ZeroAirmanTabFields = True
    
    With frmMainMenu
        
        .txtKeyField(AIRMAN_TAB).Text = 0
        .txtSorties(AIRMAN_TAB).Text = 0
        .txtKills(AIRMAN_TAB).Text = 0
        .txtMedalOfHonor(AIRMAN_TAB).Text = 0
        .txtDistinguishedServiceCross(AIRMAN_TAB).Text = 0
        .txtSilverStar(AIRMAN_TAB).Text = 0
        .txtDistinguishedFlyingCross(AIRMAN_TAB).Text = 0
        .txtBronzeStarV(AIRMAN_TAB).Text = 0
        .txtPurpleHeart(AIRMAN_TAB).Text = 0
        .txtAirMedal(AIRMAN_TAB).Text = 0
        .txtDistinguishedUnitCitation(AIRMAN_TAB).Text = 0
        .txtMeritoriousUnitCitation(AIRMAN_TAB).Text = 0
        
        .cboRank.ListIndex = 0
        .cboCrewPosition.ListIndex = 0
        .cboAssignment.ListIndex = 0
        
        If LookupAirmanStatus(DUTY_STATUS, strAirmanStatus) = False Then
' qwe            Call ExitEmulator
            ZeroAirmanTabFields = False
            Exit Function
        Else
            .txtStatus(AIRMAN_TAB).Text = strAirmanStatus
        End If
        
        .chkDefault(AIRMAN_TAB).Value = vbUnchecked
                    
        .cboRank.Enabled = True
        .cboRank.BackColor = vbWhite
                            
        .cboAssignment.Enabled = True
        .cboAssignment.BackColor = vbWhite
                            
        .cboCrewPosition.Enabled = True
        .cboCrewPosition.BackColor = vbWhite
                            
        .txtStatus(AIRMAN_TAB).ForeColor = vbBlack
        .txtStatus(AIRMAN_TAB).BackColor = PaleGreen()

    End With
End Function

'******************************************************************************
' GetAirmanRecordset
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetAirmanRecordset() As Boolean
    On Error GoTo ErrorTrap
   
    GetAirmanRecordset = True

    pobjCmnd.CommandText = "SELECT * FROM Airman ORDER BY Name"
    
    prsAirman.CursorLocation = adUseClient
    prsAirman.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    prsAirman!KeyField.Properties("Optimize") = True
    prsAirman.Sort = "Name ASC"
    
    Exit Function
   
CleanUp:
   
    Call FreeRecordset(prsAirman)
    
    Exit Function
   
ErrorTrap:
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetAirmanRecordset() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    GetAirmanRecordset = False
    
    Resume CleanUp

End Function

'******************************************************************************
' LookupAirman
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  Search for LookupKeyField in prsAirman. If it is found, point at the
'         prsAirman record that was found, then return true and AirmanName;
'         if it is not found (which should never happen), then return false and
'         blank.
'******************************************************************************
Public Function LookupAirman(ByVal LookupKeyField As Integer, ByVal LookupType As Integer, ByRef AirmanName As String) As Boolean
    
    Dim intIndex As Integer
    
    LookupAirman = False
    AirmanName = ""
    intIndex = 1

    With frmMainMenu
        
        prsAirman.MoveFirst
        Do Until prsAirman.EOF
            
            Select Case LookupType
                Case LOOKUP_BY_LISTINDEX:
                    If LookupKeyField = intIndex Then
                        AirmanName = prsAirman![Name]
                        LookupAirman = True
                        Exit Function
                    End If
                Case LOOKUP_BY_KEYFIELD:
                    If LookupKeyField = prsAirman![KeyField] Then
                        AirmanName = prsAirman![Name]
                        LookupAirman = True
                        Exit Function
                    End If
            End Select

            intIndex = intIndex + 1
            prsAirman.MoveNext
        Loop
    
    End With

    ' If the airman had been found, we would have previously exitted.
    ' Therefore, an error condition exists.
    
    strErrMsg = "LookupAirman() " & vbCrLf & vbCrLf & _
                "Airman " & LookupKeyField & " not found."

    MsgBox strErrMsg, (vbCritical + vbOKOnly)

End Function

'******************************************************************************
' PopulateAirmanCombos
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  There are three airman combos: On the airman tab, squadron tab and
'         group tab.
'******************************************************************************
Public Sub PopulateAirmanCombos()
    With frmMainMenu
        
        prsAirman.MoveFirst
        Do Until prsAirman.EOF
            
' TODO: population of commander assignment combos should be restricted
' based on airman's status
            
            .cboCommander(GROUP_TAB).AddItem prsAirman![Name]
            .cboCommander(SQUADRON_TAB).AddItem prsAirman![Name]
            .cboName(AIRMAN_TAB).AddItem prsAirman![Name]
            
            prsAirman.MoveNext
        Loop
    
    End With
End Sub

'******************************************************************************
' AddAirman
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Add a new airman. Name is a required field: It may not be blank. If
'         nothing is changed from the current airman, then a new airman with
'         the same rank and position, but a new serial number and assigned to
'         admin duty, will be created. User created airmen are never default
'         airmen.
'******************************************************************************
Public Function AddAirman() As Boolean
    On Error GoTo ErrorTrap
    
    Dim intKeyField As Integer
    Dim strIgnore As String
    Dim intToPlane As Integer
    
    AddAirman = True
    
    With frmMainMenu
        
'MsgBox ".cboName(AIRMAN_TAB).Text = " & .cboName(AIRMAN_TAB).Text & vbCrLf & _
'       ".cboRank.ListIndex(" & (.cboRank.ListIndex + 1) & ") = " & .cboRank.Text & vbCrLf & _
'       ".cboAssignment.ListIndex(" & (.cboAssignment.ListIndex + 1) & ") = " & .cboAssignment.Text & vbCrLf & _
'       ".cboCrewPosition.ListIndex(" & (.cboCrewPosition.ListIndex + 1) & ") = " & .cboCrewPosition.Text
        
        If ValidateRequiredInput(.cboName(AIRMAN_TAB)) = False Then
            Exit Function
        End If
        
        intKeyField = NextKeyField(prsAirman, "Airman")
        
        If intKeyField = 0 Then
            ' Fatal error. A value > 0 should always be returned.
' qwe            Call ExitEmulator
            AddAirman = False
            Exit Function
        End If
        
        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
        
' TODO: where does the combo listindex to when there is more than one
' person with the same name?
    
            ' If the user did not change the airman's name, then an airman
            ' with the same name will be added.
                
            prsAirman.AddNew
            
            prsAirman![KeyField] = intKeyField
            prsAirman![Name] = .cboName(AIRMAN_TAB).Text
            prsAirman![Rank] = (.cboRank.ListIndex + 1)
            
            If UCase(.cboAssignment.Text) = "ADMIN DUTY" Then
            
                intToPlane = ADMIN_DUTY
                
            Else
            
                If LookupBomber(.cboAssignment.ListIndex, LOOKUP_BY_LISTINDEX, strIgnore) = False Then
                    'ModifyAirman = False
                    'GoTo CleanUp
                End If
                
                If prsBomber![Default] = True Then
                    intToPlane = ADMIN_DUTY
                Else
                    intToPlane = prsBomber![KeyField]
                End If
            
            End If
            
            Call AssignAirmanToBomber(prsAirman![KeyField], intToPlane, (.cboCrewPosition.ListIndex + 1))
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            
            prsAirman![Status] = DUTY_STATUS
            prsAirman![Sorties] = 0
            prsAirman![Kills] = 0
            prsAirman![MedalOfHonor] = 0
            prsAirman![DistinguishedServiceCross] = 0
            prsAirman![SilverStar] = 0
            prsAirman![DistinguishedFlyingCross] = 0
            prsAirman![BronzeStarV] = 0
            prsAirman![PurpleHeart] = 0
            prsAirman![AirMedal] = 0
            prsAirman![DistinguishedUnitCitation] = 0
            prsAirman![MeritoriousUnitCitation] = 0
            prsAirman![Default] = vbUnchecked
        
            prsAirman.UpdateBatch
            prsBomber.UpdateBatch
    
        pobjConn.CommitTrans
        
        pintOpenTrans = pintOpenTrans - 1
            
        ' Insert into airman tab name combo, then change the airman tab to
        ' the new airman. Also add the new airman to the squadron and group
        ' commander combos.

'MsgBox prsAirman.AbsolutePosition

'Msgbox "prsAirman![Name] = " & prsAirman![Name] & vbCrLf & _
       "prsAirman![Rank] = " & prsAirman![Rank] & vbCrLf & _
       "prsAirman![KeyField] = " & prsAirman![KeyField] & vbCrLf & _
       "prsAirman![Assignment] = " & prsAirman![Assignment] & vbCrLf & _
       ".cboCrewPosition.ListIndex = " & (.cboCrewPosition.ListIndex + 1)

        .cboCommander(GROUP_TAB).AddItem prsAirman![Name], (prsAirman.AbsolutePosition - 1)
        .cboCommander(SQUADRON_TAB).AddItem prsAirman![Name], (prsAirman.AbsolutePosition - 1)
        .cboName(AIRMAN_TAB).AddItem prsAirman![Name], (prsAirman.AbsolutePosition - 1)
        .cboName(AIRMAN_TAB).ListIndex = (prsAirman.AbsolutePosition - 1)
                
'MsgBox .cboName(AIRMAN_TAB).ListIndex
                
        If FillAirmanTabFields() = False Then
' qwe            Call ExitEmulator
            AddAirman = False
            Exit Function
        End If

'    prsAirman.MoveFirst
'    Do Until prsAirman.EOF
'        MsgBox prsAirman![Name]
'        prsAirman.MoveNext
'    Loop

    End With

CleanUp:

Call AdjustAvailableBombers ' Nov04
'    Call AdjustMissionAvailableBombers
    
    ' Repoint the recordset to the record displayed on the tab.
    prsBomber.Bookmark = varBomberCurrentlyOnTab
                
    If pintOpenTrans Then
        pobjConn.RollbackTrans
        pintOpenTrans = pintOpenTrans - 1
    End If
    
    Exit Function

ErrorTrap:

    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "AddAirman() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)

    Err.Clear

    AddAirman = False
    
    Resume CleanUp

End Function

'******************************************************************************
' ModifyAirman
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Update the current airman. The airman's name may not be changed, nor
'         may it be blank.
'******************************************************************************
Public Function ModifyAirman() As Boolean
 '    On Error GoTo ErrorTrap
 
    Dim strIgnore As String
    Dim intPrevCrewman As Integer
    Dim intToPlane As Integer
 
    ' Update the current airman. If the airman is a group or squadron
    ' commander, is a default airman, or is in flight, he cannot be
    ' deleted. If there is a dependency, pop a msgbox, then return true. (It
    ' is not a system error.) Status is ignored to allow inactive -- i.e.,
    ' POW, DOW, KIA, etc. -- airmen to be deleted. If the airman is on a
    ' bomber's crew, he may be deleted, but the bomber will be unable to
    ' fly missions until a replacement is assigned. Due to defaults being
    ' impervious, it is not possible to delete all records, therefore there
    ' will always be some records remaining in the recordset and combo.
 
    ModifyAirman = True
    
    With frmMainMenu
        
'MsgBox ".cboName(AIRMAN_TAB).Text = " & .cboName(AIRMAN_TAB).Text & vbCrLf & _
       ".cboRank.ListIndex(" & (.cboRank.ListIndex + 1) & ") = " & .cboRank.Text & vbCrLf & _
       ".cboAssignment.ListIndex(" & (.cboAssignment.ListIndex + 1) & ") = " & .cboAssignment.Text & vbCrLf & _
       ".cboCrewPosition.ListIndex(" & (.cboCrewPosition.ListIndex + 1) & ") = " & .cboCrewPosition.Text

        ' Default airmen cannot be updated.
        
        If prsAirman![Default] = True Then
            strErrMsg = "Failed to update airman." & vbCrLf & vbCrLf & _
                        frmMainMenu.cboRank.Text & _
                        " " & _
                        prsAirman![Name] & _
                        ", serial #" & _
                        prsAirman![KeyField] & _
                        ", is either in flight and/or a default airman."
    
            MsgBox strErrMsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If

        If ValidateRequiredInput(.cboName(AIRMAN_TAB)) = False Then
            Exit Function
        ElseIf .cboName(AIRMAN_TAB).Text <> prsAirman![Name] Then
            strErrMsg = "Failed to update airman." & vbCrLf & vbCrLf & _
                        "You are not allowed to change " & _
                        frmMainMenu.cboRank.Text & _
                        " " & _
                        prsAirman![Name] & _
                        "'s name."

            MsgBox strErrMsg, (vbExclamation + vbOKOnly)

            ' This is not a severe system error, so return true.
            Exit Function
        End If
        
        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
        
            prsAirman![Rank] = (.cboRank.ListIndex + 1)
            
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            If UCase(.cboAssignment.Text) = "ADMIN DUTY" Then
            
                intToPlane = ADMIN_DUTY
                
            Else
            
                If LookupBomber(.cboAssignment.ListIndex, LOOKUP_BY_LISTINDEX, strIgnore) = False Then
                    'ModifyAirman = False
                    'GoTo CleanUp
                End If
                
                If prsBomber![Default] = True Then
                    intToPlane = ADMIN_DUTY
                Else
                    intToPlane = prsBomber![KeyField]
                End If
            
            End If
            
            Call AssignAirmanToBomber(prsAirman![KeyField], intToPlane, (.cboCrewPosition.ListIndex + 1))
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

            ' Update both airman and the bomber.
                
            prsAirman.UpdateBatch
            prsBomber.UpdateBatch
    
        pobjConn.CommitTrans
        
        pintOpenTrans = pintOpenTrans - 1
            
    End With

CleanUp:

Call AdjustAvailableBombers ' Nov04
'    Call AdjustMissionAvailableBombers
    
    ' Repoint the recordset to the record displayed on the tab.
    prsBomber.Bookmark = varBomberCurrentlyOnTab
                
    If pintOpenTrans Then
        pobjConn.RollbackTrans
        pintOpenTrans = pintOpenTrans - 1
    End If
    
    Exit Function

ErrorTrap:

    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "ModifyAirman() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)

    Err.Clear

    ModifyAirman = False
    
    Resume CleanUp

End Function

'******************************************************************************
' AssignAirmanToPos
'
' INPUT:  The airman and the position to which he should be assigned.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub AssignAirmanToPos(ByVal intAirman As Integer, ByVal intPos As Integer)

    Select Case intPos
        Case PILOT:
            prsBomber![PILOT] = intAirman
        Case COPILOT:
            prsBomber![COPILOT] = intAirman
        Case BOMBARDIER:
            prsBomber![BOMBARDIER] = intAirman
        Case NAVIGATOR:
            prsBomber![NAVIGATOR] = intAirman
        Case ENGINEER:
            prsBomber![ENGINEER] = intAirman
        Case RADIO_OPERATOR:
            prsBomber![RadioOperator] = intAirman
        Case NOSE_GUNNER:
            prsBomber![NoseGunner] = intAirman
        Case MID_UPPER_GUNNER:
            prsBomber![MidUpperGunner] = intAirman
        Case BALL_GUNNER:
            prsBomber![BallGunner] = intAirman
        Case PORT_WAIST_GUNNER:
            prsBomber![PortWaistGunner] = intAirman
        Case STBD_WAIST_GUNNER:
            prsBomber![StbdWaistGunner] = intAirman
        Case TAIL_GUNNER:
            prsBomber![TailGunner] = intAirman
        Case AMMO_STOCKER:
            prsBomber![AmmoStocker] = intAirman
    End Select

End Sub

'******************************************************************************
' AssignAirmanToBomber
'
' INPUT:  The airman and the plane and position to which he should be assigned.
'         (intNewAirman does not mean the airman is totally new, rather that he
'         is new to the bomber.)
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Things That May be Changed By This Function
'         -------------------------------------------
'         New Airman:  plane and position (plane intNewAirman currently assigned to)
'         From Plane:  position and status (position intNewAirman currently assigned to)
'         From Pos:    assigned airman
'         Prev Airman: plane (intPrevAirman currently occupies intToPos on intToPlane)
'         To Plane:    position and status
'         To Pos:      assigned airman
'******************************************************************************
Private Sub AssignAirmanToBomber(ByVal intNewAirman As Integer, ByVal intToPlane As Integer, ByVal intToPos As Integer)

    Dim strIgnore As String
    Dim intFromPos As Integer
    Dim intFromPlane As Integer
    Dim intPrevAirman As Integer
    Dim varNewAirmanBookmark As Variant
    
    If IsNull(prsAirman![CrewPosition]) = True Then
        intFromPos = ADMIN_DUTY
    Else
        intFromPos = prsAirman![CrewPosition]
    End If
    
    If IsNull(prsAirman![Assignment]) = True Then
        intFromPlane = ADMIN_DUTY
    Else
        intFromPlane = prsAirman![Assignment]
    End If
    
    If intFromPlane = intToPlane _
    And intFromPos = intToPos Then
        ' Airman is not switching planes or positions.
        Exit Sub
    End If

    If intFromPlane > ADMIN_DUTY Then
        
        ' The current airman actually is assigned to a bomber. Point at it.
        
        If LookupBomber(intFromPlane, LOOKUP_BY_KEYFIELD, strIgnore) = False Then
            'AssignAirmanToBomber = False
            'GoTo CleanUp
        End If

        ' Mark the position as unmanned.
        
        Call AssignAirmanToPos(UNMANNED_POSITION, intFromPos)
        
        ' Change bomber to stand down status.
        
        prsBomber![Status] = STAND_DOWN_STATUS
    
    Else
        
        ' If the current airman was on admin duty, then there obviously is no
        ' intFromPlane or intFromPos that needs to be updated.

    End If
    
    If intToPlane > ADMIN_DUTY Then
    
        ' Point at the plane the current airman is being assigned to (if he
        ' is not being reassigned within the plane).
        
        If intFromPlane <> intToPlane Then
            
            If LookupBomber(intToPlane, LOOKUP_BY_KEYFIELD, strIgnore) = False Then
                'AssignAirmanToBomber = False
                'GoTo CleanUp
            End If
        
        End If
        
        ' Find the position the airman is being assigned to, then save
        ' the previously assigned airman's key.
    
        Select Case intToPos
            Case PILOT:
                intPrevAirman = prsBomber![PILOT]
            Case COPILOT:
                intPrevAirman = prsBomber![COPILOT]
            Case BOMBARDIER:
                intPrevAirman = prsBomber![BOMBARDIER]
            Case NAVIGATOR:
                intPrevAirman = prsBomber![NAVIGATOR]
            Case ENGINEER:
                intPrevAirman = prsBomber![ENGINEER]
            Case RADIO_OPERATOR:
                intPrevAirman = prsBomber![RadioOperator]
            Case NOSE_GUNNER:
                intPrevAirman = prsBomber![NoseGunner]
            Case MID_UPPER_GUNNER:
                intPrevAirman = prsBomber![MidUpperGunner]
            Case BALL_GUNNER:
                intPrevAirman = prsBomber![BallGunner]
            Case PORT_WAIST_GUNNER:
                intPrevAirman = prsBomber![PortWaistGunner]
            Case STBD_WAIST_GUNNER:
                intPrevAirman = prsBomber![StbdWaistGunner]
            Case TAIL_GUNNER:
                intPrevAirman = prsBomber![TailGunner]
            Case AMMO_STOCKER:
                intPrevAirman = prsBomber![AmmoStocker]
        End Select
        
        If intPrevAirman > UNMANNED_POSITION Then
        
            ' Point at the current airman, so we can point back to him after
            ' updating the previous airman.
        
            varNewAirmanBookmark = prsAirman.Bookmark
        
            ' Point at the previous airman.
            
            If LookupAirman(intPrevAirman, LOOKUP_BY_KEYFIELD, strIgnore) = False Then
                'AssignAirmanToBomber = False
                'GoTo CleanUp
            End If
                
            ' Assign the previous Airman to admin duty.
            
            prsAirman![Assignment] = ADMIN_DUTY
        
            ' Return to the airman we are assigning to the position.
            
            prsAirman.Bookmark = varNewAirmanBookmark
        
        End If
        
        ' Assign the new airman to the position.
        
        Call AssignAirmanToPos(intNewAirman, intToPos)
        
        ' If all the positions that exist on the bomber are occupied, then
        ' the bomber should be set to duty status.
        
        If prsBomber![BomberModel] = B24_D _
        Or prsBomber![BomberModel] = B24_E _
        Or prsBomber![BomberModel] = B24_GHJ _
        Or prsBomber![BomberModel] = B24_LM Then
        
            If prsBomber![PILOT] <> UNMANNED_POSITION _
            And prsBomber![COPILOT] <> UNMANNED_POSITION _
            And prsBomber![BOMBARDIER] <> UNMANNED_POSITION _
            And prsBomber![NAVIGATOR] <> UNMANNED_POSITION _
            And prsBomber![ENGINEER] <> UNMANNED_POSITION _
            And prsBomber![RadioOperator] <> UNMANNED_POSITION _
            And prsBomber![NoseGunner] <> UNMANNED_POSITION _
            And prsBomber![MidUpperGunner] <> UNMANNED_POSITION _
            And prsBomber![BallGunner] <> UNMANNED_POSITION _
            And prsBomber![StbdWaistGunner] <> UNMANNED_POSITION _
            And prsBomber![TailGunner] <> UNMANNED_POSITION _
            And prsBomber![AmmoStocker] <> UNMANNED_POSITION Then
                prsBomber![Status] = DUTY_STATUS
            End If

        Else
        
            If prsBomber![PILOT] <> UNMANNED_POSITION _
            And prsBomber![COPILOT] <> UNMANNED_POSITION _
            And prsBomber![BOMBARDIER] <> UNMANNED_POSITION _
            And prsBomber![NAVIGATOR] <> UNMANNED_POSITION _
            And prsBomber![ENGINEER] <> UNMANNED_POSITION _
            And prsBomber![RadioOperator] <> UNMANNED_POSITION _
            And prsBomber![NoseGunner] <> UNMANNED_POSITION _
            And prsBomber![MidUpperGunner] <> UNMANNED_POSITION _
            And prsBomber![BallGunner] <> UNMANNED_POSITION _
            And prsBomber![PortWaistGunner] <> UNMANNED_POSITION _
            And prsBomber![StbdWaistGunner] <> UNMANNED_POSITION _
            And prsBomber![TailGunner] <> UNMANNED_POSITION _
            And prsBomber![AmmoStocker] <> UNMANNED_POSITION Then
                prsBomber![Status] = DUTY_STATUS
            End If

        End If
        
        ' Set the current airman's position.
        
        prsAirman![CrewPosition] = intToPos
        
        ' Set the current airman's assignment.
    
        prsAirman![Assignment] = intToPlane
                
    Else

        ' Set the current airman's position.
        
        prsAirman![CrewPosition] = intToPos
           
    End If

    prsBomber.Bookmark = varBomberCurrentlyOnTab

End Sub

'******************************************************************************
' DeleteAirman
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Delete the current airman. If the airman is a group or squadron
'         commander, or is a default airman, he cannot be deleted. If there
'         is a dependency, pop a msgbox, then return true. (It is not a system
'         error.) Status is ignored to allow inactive -- i.e., POW, DOW, KIA,
'         etc. -- airmen to be deleted. If the airman is on a Bomber 's crew,
'         he may be deleted, but the bomber will be unable to fly missions
'         until a replacement is assigned. Due to defaults being impervious,
'         it is not possible to delete all records, therefore there will always
'         be some records remaining in the recordset and combo.
'******************************************************************************
Public Function DeleteAirman() As Boolean
    On Error GoTo ErrorTrap
    
' TODO: What happens if you delete an airman that is assigned to an
' inactive bomber?
    
    Dim strUnitCommanded As String
    Dim strBomberName As String
    Dim blnCommander As Boolean
    Dim blnCrewman As Boolean
    Dim strIgnore As String
    Dim strAirman As String
    Dim intListIndex As Integer

    DeleteAirman = True

    With frmMainMenu
        
        ' Default airmen cannot be deleted.
        
        If prsAirman![Default] = True Then
            strErrMsg = "Failed to delete airman." & vbCrLf & vbCrLf & _
                        .cboRank.Text & _
                        " " & _
                        prsAirman![Name] & _
                        ", serial #" & _
                        prsAirman![KeyField] & _
                        ", is either in flight and/or a default airman."
    
            MsgBox strErrMsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
    
        ' Commanders cannot be deleted.
        
        blnCommander = IsCommander(prsGroup, strUnitCommanded)
    
        If blnCommander = False Then
            blnCommander = IsCommander(prsSquadron, strUnitCommanded)
        End If
        
        If blnCommander = True Then
            strErrMsg = "Failed to delete airman." & vbCrLf & vbCrLf & _
                        .cboRank.Text & _
                        " " & _
                        prsAirman![Name] & _
                        ", serial #" & _
                        prsAirman![KeyField] & _
                        ", is the " & _
                        strUnitCommanded & _
                        " commander."
    
            MsgBox strErrMsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
    
        ' Crewmen may be deleted, but pop a warning first.
        
        blnCrewman = IsCrewman(strBomberName)
        
        If blnCrewman = True Then
            strErrMsg = .cboRank.Text & _
                        " " & _
                        prsAirman![Name] & _
                        ", serial #" & _
                        prsAirman![KeyField] & _
                        ", is on the " & _
                        strBomberName & _
                        "'s crew. Deleting the airman will leave a " & _
                        "position unfilled, preventing the bomber from " & _
                        "flying further missions until a replacement is " & _
                        "assigned." & vbCrLf & vbCrLf & _
                        "Do you wish to continue?"
                
            If MsgBox(strErrMsg, (vbYesNo + vbDefaultButton2 + vbQuestion)) = vbNo Then
                Exit Function
            End If
        Else
            ' Normal deletion confirmation is in the else block to avoid double
            ' warnings where the airman is a crewman.
            
            If MsgBox("Are you sure?", (vbYesNo + vbDefaultButton2 + vbQuestion)) = vbNo Then
                Exit Function
            End If
        End If

'Msgbox prsAirman![Name] & ", serial #" & prsAirman![KeyField] & ", is about to be deleted."
        
        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
            
            If LookupBomber(prsAirman![Assignment], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
                DeleteAirman = False
                GoTo CleanUp
            End If

            If blnCrewman = True Then
                    
                Select Case prsAirman![CrewPosition]
                    Case PILOT:
                        prsBomber![PILOT] = UNMANNED_POSITION
                    Case COPILOT:
                        prsBomber![COPILOT] = UNMANNED_POSITION
                    Case BOMBARDIER:
                        prsBomber![BOMBARDIER] = UNMANNED_POSITION
                    Case NAVIGATOR:
                        prsBomber![NAVIGATOR] = UNMANNED_POSITION
                    Case ENGINEER:
                        prsBomber![ENGINEER] = UNMANNED_POSITION
                    Case RADIO_OPERATOR:
                        prsBomber![RadioOperator] = UNMANNED_POSITION
                    Case NOSE_GUNNER:
                        prsBomber![NoseGunner] = UNMANNED_POSITION
                    Case MID_UPPER_GUNNER:
                        prsBomber![MidUpperGunner] = UNMANNED_POSITION
                    Case BALL_GUNNER:
                        prsBomber![BallGunner] = UNMANNED_POSITION
                    Case PORT_WAIST_GUNNER:
                        prsBomber![PortWaistGunner] = UNMANNED_POSITION
                    Case STBD_WAIST_GUNNER:
                        prsBomber![StbdWaistGunner] = UNMANNED_POSITION
                    Case TAIL_GUNNER:
                        prsBomber![TailGunner] = UNMANNED_POSITION
                    Case AMMO_STOCKER:
                        prsBomber![AmmoStocker] = UNMANNED_POSITION
                End Select
                    
                ' Repoint the recordset to the record displayed on the tab.
                prsBomber.Bookmark = varBomberCurrentlyOnTab
                
            End If
    
            prsAirman.Delete
    
            prsAirman.UpdateBatch
            prsBomber.UpdateBatch
        
        pobjConn.CommitTrans
            
        pintOpenTrans = pintOpenTrans - 1
            
        ' Delete from airman tab name combo, then change the airman tab to
        ' another listed airman. Also delete from squadron tab and group tab
        ' commander combos. (Though they do not need to be re-indexed because
        ' no squadrons or groups, including the current ones, were dependent
        ' on the deleted airman.)
    
        intListIndex = .cboName(AIRMAN_TAB).ListIndex
        
        .cboName(AIRMAN_TAB).RemoveItem intListIndex
        .cboCommander(GROUP_TAB).RemoveItem intListIndex
        .cboCommander(SQUADRON_TAB).RemoveItem intListIndex

        If intListIndex = 0 Then
            
            ' Deleted the first listed record. Point to the second record,
            ' which just became the first record.
            .cboName(AIRMAN_TAB).ListIndex = 0
        
        ElseIf intListIndex = .cboName(AIRMAN_TAB).ListCount Then
            
            ' Deleted the last listed record. Point to the second-to-last
            ' record, which just became the last record. ListIndex is 0-base,
            ' while ListCount is 1-base, so we need to subtract 1.
            .cboName(AIRMAN_TAB).ListIndex = intListIndex - 1 '(.cboName(AIRMAN_TAB).ListCount - 1)
        
        Else
            
            ' Deleted some record between the first and last records. Point
            ' to the record after the one that was deleted.
            .cboName(AIRMAN_TAB).ListIndex = intListIndex
        
        End If

        If LookupAirman((.cboName(AIRMAN_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = True Then
            If FillAirmanTabFields() = False Then
' qwe            Call ExitEmulator
                DeleteAirman = False
                Exit Function
            End If
        Else
' qwe            Call ExitEmulator
            DeleteAirman = False
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

    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "DeleteAirman() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)

    Err.Clear

    DeleteAirman = False
    
    Resume CleanUp

End Function

'******************************************************************************
' IsCommander
'
' INPUT:  Either a squadron or group recordset.
'
' OUTPUT: The airman's name if he is a commander, otherwise blank.
'
' RETURN: True if the airman is a commander, otherwise false.
'
' NOTES:
'******************************************************************************
Public Function IsCommander(ByRef frsUnit As ADODB.Recordset, ByRef strUnitCommanded As String) As Boolean
    Dim varBookmark As Variant

    IsCommander = False

    ' Point to the record currently on the unit tab.
    
    varBookmark = frsUnit.Bookmark
    
    ' Determine if the airman is a commander.
    
    frsUnit.MoveFirst
    Do Until frsUnit.EOF
    
        If frsUnit![Commander] = prsAirman![KeyField] Then
            strUnitCommanded = frsUnit![Name]
            IsCommander = True
            Exit Do
        End If
    
        frsUnit.MoveNext
    Loop

    ' Re-point the recordset to the record currently on the unit tab.

    frsUnit.Bookmark = varBookmark

End Function

'******************************************************************************
' IsCrewman
'
' INPUT:  n/a
'
' OUTPUT: The airman's name if he assigned to a bomber, otherwise blank.
'
' RETURN: True if the airman is assigned to a bomber, otherwise false.
'
' NOTES:
'******************************************************************************
Public Function IsCrewman(ByRef strBomberName As String) As Boolean

    IsCrewman = False

    ' Determine if the airman is on a bomber's crew.
    
    prsBomber.MoveFirst
    Do Until prsBomber.EOF
    
        If prsBomber![PILOT] = prsAirman![KeyField] _
        Or prsBomber![COPILOT] = prsAirman![KeyField] _
        Or prsBomber![BOMBARDIER] = prsAirman![KeyField] _
        Or prsBomber![NAVIGATOR] = prsAirman![KeyField] _
        Or prsBomber![ENGINEER] = prsAirman![KeyField] _
        Or prsBomber![RadioOperator] = prsAirman![KeyField] _
        Or prsBomber![NoseGunner] = prsAirman![KeyField] _
        Or prsBomber![MidUpperGunner] = prsAirman![KeyField] _
        Or prsBomber![BallGunner] = prsAirman![KeyField] _
        Or prsBomber![PortWaistGunner] = prsAirman![KeyField] _
        Or prsBomber![StbdWaistGunner] = prsAirman![KeyField] _
        Or prsBomber![TailGunner] = prsAirman![KeyField] _
        Or prsBomber![AmmoStocker] = prsAirman![KeyField] Then
            
            strBomberName = prsBomber![Name]
            IsCrewman = True
            Exit Do
        
        End If
    
        prsBomber.MoveNext
    Loop

    ' Re-point the recordset to the record currently on the bomber tab.

    prsBomber.Bookmark = varBomberCurrentlyOnTab

End Function

