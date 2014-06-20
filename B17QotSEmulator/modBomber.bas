Attribute VB_Name = "modBomber"
'******************************************************************************
' modBomber.bas
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

Public prsBomber As New ADODB.Recordset
Public varBomberCurrentlyOnTab As Variant

Public intBomberMission() As Integer

' kgreer (12 Dec 04)
Public intMapBomberKeyToAssignmentIndex() As Integer

Dim strErrmsg As String

'******************************************************************************
' FillBomberTabFields
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
Public Function FillBomberTabFields() As Boolean
    
    Dim strBomberModel As String
    Dim strSquadron As String
    Dim strBomberStatus As String
    
    FillBomberTabFields = True
    
    With frmMainMenu
        
        ' Populate the non-lookup fields.
        
        .txtKeyField(BOMBER_TAB).Text = prsBomber![KeyField]
        .cboName(BOMBER_TAB).Text = prsBomber![Name]
        .txtStatus(BOMBER_TAB).Text = prsBomber![Status]
        .txtSorties(BOMBER_TAB).Text = prsBomber![Sorties]
        .txtKills(BOMBER_TAB).Text = prsBomber![Kills]
        .txtRabbitsFoot.Text = prsBomber![RabbitsFoot]
        
        .txtManufacturer = prsBomber![Manufacturer]
        .txtPlant = prsBomber![Plant]
        .txtTailNumber = prsBomber![TailNumber]
        
        If prsBomber![Default] = True Then
            .chkDefault(BOMBER_TAB).Value = vbChecked
        Else
            .chkDefault(BOMBER_TAB).Value = vbUnchecked
        End If

        ' Populate the recordset lookup fields.
        
        If LookupBomberModel(prsBomber![BomberModel], strBomberModel) = False Then
            FillBomberTabFields = False
            Exit Function
        Else
            .cboBomberModel(BOMBER_TAB).Text = strBomberModel
        End If

        ' We can do this because the combo is sorted in KeyField order. We
        ' must do it so the proper positions are displayed on the crew
        ' assignment dialog.
        
        .cboBomberModel(BOMBER_TAB).ListIndex = (prsBomber![BomberModel] - 1)

        Call PopulateBomberSquadronCombo
        
        If LookupBomberSquadron(prsBomber![Squadron], LOOKUP_BY_KEYFIELD, strSquadron) = False Then
            FillBomberTabFields = False
            Exit Function
        Else
            .cboSquadron.Text = strSquadron

            ' Repoint the squadron recordset to the record displayed on the squadron tab.
            prsSquadron.Bookmark = varSquadronCurrentlyOnTab
        End If
        
        If LookupBomberStatus(prsBomber![Status], strBomberStatus) = False Then
            FillBomberTabFields = False
            Exit Function
        Else
            .txtStatus(BOMBER_TAB).Text = strBomberStatus
        End If

        ' If the bommber is on duty status, enable the fields, otherwise
        ' ensure they are disabled. Also, set the status colors.
        
        Select Case prsBomber![Status]
            Case DUTY_STATUS, STAND_DOWN_STATUS:
        
                .cboBomberModel(BOMBER_TAB).Enabled = False
                .cboBomberModel(BOMBER_TAB).BackColor = vbButtonFace
                
                .cboSquadron.Enabled = True
                .cboSquadron.BackColor = vbWhite
                
                .txtStatus(BOMBER_TAB).ForeColor = vbBlack
                .txtStatus(BOMBER_TAB).BackColor = PaleGreen()
                
            Case CRASHED_STATUS, SCRAPPED_STATUS:
                
                .cboBomberModel(BOMBER_TAB).Enabled = False
                .cboBomberModel(BOMBER_TAB).BackColor = vbButtonFace
                
                .cboSquadron.Enabled = False
                .cboSquadron.BackColor = vbButtonFace
                
                .txtStatus(BOMBER_TAB).ForeColor = vbBlack
                .txtStatus(BOMBER_TAB).BackColor = PaleRed()
            
            Case CAPTURED_STATUS, DITCHED_STATUS, SHOT_DOWN_STATUS:
            
                .cboBomberModel(BOMBER_TAB).Enabled = False
                .cboBomberModel(BOMBER_TAB).BackColor = vbButtonFace
                
                .cboSquadron.Enabled = False
                .cboSquadron.BackColor = vbButtonFace
                
                .txtStatus(BOMBER_TAB).ForeColor = vbWhite
                .txtStatus(BOMBER_TAB).BackColor = vbBlack
            
            Case RETIRED_STATUS:
            
                .cboBomberModel(BOMBER_TAB).Enabled = False
                .cboBomberModel(BOMBER_TAB).BackColor = vbButtonFace
                
                .cboSquadron.Enabled = False
                .cboSquadron.BackColor = vbButtonFace
                
                .txtStatus(BOMBER_TAB).ForeColor = vbBlack
                .txtStatus(BOMBER_TAB).BackColor = vbButtonFace
            
        End Select
            
        ' Regardless of the bomber's status, if it is a default bomber,
        ' disable the fields.
        
        If .chkDefault(BOMBER_TAB).Value = vbChecked Then
            .cboBomberModel(BOMBER_TAB).Enabled = False
            .cboBomberModel(BOMBER_TAB).BackColor = vbButtonFace
            .cboSquadron.Enabled = False
            .cboSquadron.BackColor = vbButtonFace
        End If
    
        If prsBomber![Status] = DUTY_STATUS _
        Or prsBomber![Status] = STAND_DOWN_STATUS Then
            If .chkDefault(BOMBER_TAB).Value = vbChecked Then
                .cmdAssignCrew.Caption = "Default Crew"
                .cmdRetireBomber.Visible = False
            Else
                .cmdAssignCrew.Caption = "Assign Crew"
                .cmdRetireBomber.Visible = True
            End If
        Else ' Bomber no longer in service
            .cmdAssignCrew.Caption = "Last Crew"
            .cmdRetireBomber.Visible = False
        End If
    
    End With

End Function

'******************************************************************************
' ZeroBomberTabFields
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
Public Function ZeroBomberTabFields() As Boolean
    Dim strBomberStatus As String
    
    ZeroBomberTabFields = True
    
    With frmMainMenu
        
        .txtKeyField(BOMBER_TAB).Text = 0
        .txtSorties(BOMBER_TAB).Text = 0
        .txtKills(BOMBER_TAB).Text = 0
        .txtRabbitsFoot.Text = 0
        .txtManufacturer = ""
        .txtPlant = ""
        .txtTailNumber = ""
        
        .cboBomberModel(BOMBER_TAB).ListIndex = 0
        
        If LookupBomberStatus(DUTY_STATUS, strBomberStatus) = False Then
' qwe            Call ExitEmulator
            ZeroBomberTabFields = False
            Exit Function
        Else
            .txtStatus(BOMBER_TAB).Text = strBomberStatus
        End If

        .cboSquadron.ListIndex = 0
        
        .chkDefault(BOMBER_TAB).Value = vbUnchecked
                    
        .cboBomberModel(BOMBER_TAB).Enabled = True
        .cboBomberModel(BOMBER_TAB).BackColor = vbWhite
                            
        .cboSquadron.Enabled = True
        .cboSquadron.BackColor = vbWhite
                            
        .txtStatus(BOMBER_TAB).ForeColor = vbBlack
        .txtStatus(BOMBER_TAB).BackColor = PaleGreen()
    End With
End Function

'******************************************************************************
' GetBomberRecordset
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function GetBomberRecordset() As Boolean
    On Error GoTo ErrorTrap

    GetBomberRecordset = True

    pobjCmnd.CommandText = "SELECT * FROM Bomber ORDER BY Name"
    
    prsBomber.CursorLocation = adUseClient
    prsBomber.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    prsBomber!KeyField.Properties("Optimize") = True
    prsBomber.Sort = "Name ASC"
    
    Exit Function
   
CleanUp:
   
    Call FreeRecordset(prsBomber)
   
    Exit Function
   
ErrorTrap:
    
    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "GetBomberRecordset() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    GetBomberRecordset = False
    
    Resume CleanUp

End Function

'******************************************************************************
' LookupBomber
'
' INPUT:  n/a
'
' OUTPUT: Bomber name if successful, otherwise blank.
'
' RETURN: True if succesful, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function LookupBomber(ByVal LookupKeyField As Integer, ByVal LookupType As Integer, ByRef BomberName As String) As Boolean
    
    Dim intIndex As Integer

    LookupBomber = False
    BomberName = ""
    intIndex = 1

    If LookupKeyField = 0 Then
        ' Assignment to Bomber 0 indicates admin/desk duty. This is the
        ' only case in the emulator where a lower level entity may exist
        ' without belonging to some higher level entity. Avoid the not
        ' found error by aborting this function.
    
        BomberName = "Admin Duty"
        LookupBomber = True
        Exit Function
    End If
        
    With frmMainMenu
        
        prsBomber.MoveFirst
        Do Until prsBomber.EOF
            
            Select Case LookupType
                Case LOOKUP_BY_LISTINDEX:
                    If LookupKeyField = intIndex Then
                        BomberName = prsBomber![Name]
                        LookupBomber = True
                        Exit Function
                    End If
                Case LOOKUP_BY_KEYFIELD:
                    If LookupKeyField = prsBomber![KeyField] Then
                        BomberName = prsBomber![Name]
                        LookupBomber = True
                        Exit Function
                    End If
            End Select
            
            intIndex = intIndex + 1
            prsBomber.MoveNext
        Loop
    
    End With

    ' If the bomber had been found, we would have previously exitted.
    ' Therefore, an error condition exists.
    
    strErrmsg = "LookupBomber() " & vbCrLf & vbCrLf & _
                "Bomber " & LookupKeyField & " not found."

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

End Function
 
' kgreer (12 Dec 04)
'******************************************************************************
' MapBomberKeyToAssignmentIndex
'
' INPUT:  An airman's assignment (which the keyfield for a bomber).
'
' OUTPUT: n/a
'
' RETURN: The 0-base location in .cboAssignment which contains the name of
'         the airman's assigned bomber. If the bomber isn't found, the return
'         value is 0, which indexes to "Admin Duty".
'
' NOTES:  n/a
'******************************************************************************
Function MapBomberKeyToAssignmentIndex(intBomberKey)
    Dim intIndex As Integer
    Dim intLast As Integer
    
    MapBomberKeyToAssignmentIndex = 0
    
    intLast = UBound(intMapBomberKeyToAssignmentIndex)

    For intIndex = 0 To intLast
        
        If intBomberKey = intMapBomberKeyToAssignmentIndex(intIndex) Then
            MapBomberKeyToAssignmentIndex = intIndex
            Exit For
        End If
        
    Next intIndex

End Function
 
'******************************************************************************
' HasAssignedAirman
'
' INPUT:  n/a
'
' OUTPUT: Comma-delimited list of airman.
'
' RETURN: True if airmen are assigned, otherwise false.
'
' NOTES:  Determine if the current bomber has any airmen assigned to it.
'******************************************************************************
Private Function HasAssignedAirman(ByRef strReturn As String) As Boolean
    Dim intSecondLastCommaLoc As Integer
    Dim intLastCommaLoc As Integer
    Dim strTemp As String
    Dim intIndex As Integer

    HasAssignedAirman = False

    intSecondLastCommaLoc = 0
    intLastCommaLoc = 0
    strReturn = ""
    strTemp = ""
    
    ' Determine if any airmen are assigned to the bomber.
    
    prsAirman.MoveFirst
    Do Until prsAirman.EOF
    
        If prsAirman![Assignment] = prsBomber![KeyField] Then
            HasAssignedAirman = True
            
            strReturn = strReturn & prsAirman![Name] & ", "
            
            intSecondLastCommaLoc = intLastCommaLoc
            
            intLastCommaLoc = Len(strReturn) - 1
        End If
            
        prsAirman.MoveNext
    Loop
    
    ' If there are dependent airmen, correct the return string's grammar.
        
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
    
    ' Re-point the airman recordset to the record currently on the airman tab.

    prsAirman.Bookmark = varAirmanCurrentlyOnTab

End Function

'******************************************************************************
' PopulateBomberCombos
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  There are two bomber combos: One on the bomber tab and one on the
'         airman tab.
'******************************************************************************
Public Sub PopulateBomberCombos()
    Dim intIndex As Integer
    
    intIndex = 0
    
    With frmMainMenu
    
        prsBomber.MoveFirst
        
        ' Admin duty indicates the airman is not assigned to a bomber.
        ' Since "not a bomber" does not belong in the Bomber table,
        ' add it to the top of the combo.
        .cboAssignment.AddItem "Admin Duty"
        
        Do Until prsBomber.EOF
            
            .cboName(BOMBER_TAB).AddItem prsBomber![Name]
            .cboAssignment.AddItem prsBomber![Name] ' AIRMAN_TAB
            
            prsBomber.MoveNext
        
        Loop
    
    End With

' ??? cause airmen = 0 error ???
    ' Normally we would set prsBomber.Bookmark = varBomberCurrentlyOnTab after
    ' adjusting the mission availabl bombers, but when PopulateBomberCombos()
    ' is first called, when the app is loaded, the bookmark has not yet been
    ' set.
Call AdjustAvailableBombers ' Nov04
'    Call AdjustMissionAvailableBombers
            
End Sub

'******************************************************************************
' AddBomber
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Add a new bomber. Name is a required field: It cannot be blank, and
'         it cannot be a duplicate name. If nothing else is changed from the
'         current bomber, then a new bomber with the same commander, group
'         and airman type will be created. User created bombers are never
'         default bombers.
'******************************************************************************
Public Function AddBomber() As Boolean
    On Error GoTo ErrorTrap
    
    Dim intKeyField As Integer
    Dim strIgnore As String
    
    AddBomber = True

    With frmMainMenu
        If ValidateRequiredInput(.cboName(BOMBER_TAB)) = False _
        Or ValidateRequiredInput(.cboSquadron) = False Then
' qwe        If ValidateRequiredInput(.cboName(BOMBER_TAB)) = False Then
            ' This is not a severe system error, so return true.
            Exit Function
        End If
        
        If IsDupeBomberName() = True Then
            strErrmsg = "Failed to add bomber." & vbCrLf & vbCrLf & _
                        "The " & _
                        .cboName(BOMBER_TAB).Text & _
                        " already exists."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
        
        intKeyField = NextKeyField(prsBomber, "Bomber")
        
        If intKeyField = 0 Then
            ' Fatal error. A value > 0 should always be returned.
' qwe            Call ExitEmulator
            AddBomber = False
            Exit Function
        End If
        
        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
        
            prsBomber.AddNew

            prsBomber![KeyField] = intKeyField
            prsBomber![Name] = .cboName(BOMBER_TAB).Text
            prsBomber![BomberModel] = (.cboBomberModel(BOMBER_TAB).ListIndex + 1)
            prsBomber![Status] = STAND_DOWN_STATUS
            prsBomber![Sorties] = 0
            prsBomber![Kills] = 0
            prsBomber![RabbitsFoot] = 0
            
            ' Get tail number, manufacturer and plant.
            Call BomberBuildData

            If LookupBomberSquadron((.cboSquadron.ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = False Then
' qwe            Call ExitEmulator
                AddBomber = False
                Exit Function
            Else
                prsBomber![Squadron] = prsBomberSquadron![KeyField]
            End If

            prsBomber![Default] = vbUnchecked
            
            ' If a position exists on the bomber, it is initially empty (and
            ' will remain so until airmen are assigned). If the position does
            ' not exist on the bomber, it will be hidden forever.
            
            prsBomber![PILOT] = UNMANNED_POSITION
            prsBomber![NAVIGATOR] = UNMANNED_POSITION
            prsBomber![ENGINEER] = UNMANNED_POSITION
            prsBomber![RadioOperator] = UNMANNED_POSITION
            
            If prsBomber![BomberModel] <> AVRO_LANCASTER Then
                prsBomber![COPILOT] = UNMANNED_POSITION
            Else
                prsBomber![COPILOT] = HIDDEN_POSITION
            End If
            
            If prsBomber![BomberModel] <> YB40 Then
                prsBomber![BOMBARDIER] = UNMANNED_POSITION
            Else
                prsBomber![BOMBARDIER] = HIDDEN_POSITION
            End If
            
            If prsBomber![BomberModel] = B24_GHJ _
            Or prsBomber![BomberModel] = B24_LM _
            Or prsBomber![BomberModel] = YB40 Then
                prsBomber![NoseGunner] = UNMANNED_POSITION
            Else
                prsBomber![NoseGunner] = HIDDEN_POSITION
            End If
            
            If prsBomber![BomberModel] = YB40 _
            Or prsBomber![BomberModel] = AVRO_LANCASTER Then
                prsBomber![MidUpperGunner] = UNMANNED_POSITION
            Else
                prsBomber![MidUpperGunner] = HIDDEN_POSITION
            End If
            
            If prsBomber![BomberModel] <> AVRO_LANCASTER Then
                prsBomber![BallGunner] = UNMANNED_POSITION
            Else
                prsBomber![BallGunner] = HIDDEN_POSITION
            End If
            
            If prsBomber![BomberModel] = B17_C _
            Or prsBomber![BomberModel] = B17_E _
            Or prsBomber![BomberModel] = B17_F _
            Or prsBomber![BomberModel] = B17_G _
            Or prsBomber![BomberModel] = YB40 _
            Or prsBomber![BomberModel] = B24_D _
            Or prsBomber![BomberModel] = B24_E _
            Or prsBomber![BomberModel] = B24_GHJ _
            Or prsBomber![BomberModel] = B24_LM Then
                prsBomber![PortWaistGunner] = UNMANNED_POSITION
            Else
                prsBomber![PortWaistGunner] = HIDDEN_POSITION
            End If
            
            If prsBomber![BomberModel] <> AVRO_LANCASTER Then
                prsBomber![StbdWaistGunner] = UNMANNED_POSITION
            Else
                prsBomber![StbdWaistGunner] = HIDDEN_POSITION
            End If
            
            If prsBomber![BomberModel] <> B17_C Then
                prsBomber![TailGunner] = UNMANNED_POSITION
            Else
                prsBomber![TailGunner] = HIDDEN_POSITION
            End If
            
            If prsBomber![BomberModel] = YB40 Then
                prsBomber![AmmoStocker] = UNMANNED_POSITION
            Else
                prsBomber![AmmoStocker] = HIDDEN_POSITION
            End If
            
            prsBomber.UpdateBatch
    
        pobjConn.CommitTrans
        
        pintOpenTrans = pintOpenTrans - 1
            
        ' Insert into bomber tab name combo, then change the bomber tab
        ' to the new bomber. Also add the new bomber to the airman tab
        ' bomber combo. (Though it does not need to be re-indexed because
        ' no airmen, including the current one, are dependent on the new
        ' bomber.)

        .cboName(BOMBER_TAB).AddItem prsBomber![Name], (prsBomber.AbsolutePosition - 1)
        .cboAssignment.AddItem prsBomber![Name], (prsBomber.AbsolutePosition - 1)

        .cboName(BOMBER_TAB).ListIndex = (prsBomber.AbsolutePosition - 1)

'        If LookupBomber((.cboName(BOMBER_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = True Then
'            ' All we need is the pointer to the matching record (the
'            ' Bomber name was selected). Fill in the tab fields.
'            If FillBomberTabFields() = False Then
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
                "AddBomber() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

    Err.Clear

    AddBomber = False
    
    Resume CleanUp

End Function

'******************************************************************************
' ModifyBomber
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Update the current bomber. The bomber's name may not be changed,
'         nor may it be blank.
'******************************************************************************
Public Function ModifyBomber() As Boolean
     On Error GoTo ErrorTrap
 
    Dim strIgnore As String
 
    ModifyBomber = True
    
    With frmMainMenu
        
        ' Default bombers cannot be updated.
        
        If prsBomber![Default] = True Then
            strErrmsg = "Failed to update bomber." & vbCrLf & vbCrLf & _
                        prsBomber![Name] & _
                        " is a default bomber."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
    
        If ValidateRequiredInput(.cboName(BOMBER_TAB)) = False Then
            Exit Function
        ElseIf .cboName(BOMBER_TAB).Text <> prsBomber![Name] Then
            strErrmsg = "Failed to update bomber." & vbCrLf & vbCrLf & _
                        "You are not allowed to change the " & _
                        prsBomber![Name] & _
                        "'s name."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
        
        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
        
            If LookupBomberSquadron((.cboSquadron.ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = False Then
' qwe            Call ExitEmulator
                ModifyBomber = False
                Exit Function
            Else
                prsBomber![Squadron] = prsBomberSquadron![KeyField]
            End If

            prsBomber.UpdateBatch
    
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
                "ModifyBomber() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

    Err.Clear

    ModifyBomber = False
    
    Resume CleanUp

End Function

'******************************************************************************
' DeleteBomber
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Delete the current bomber. If the bomber is a default bomber, it
'         cannot be deleted. If there is a dependency, pop a msgbox, then return
'         true. (It is not a system error.) Due to defaults being impervious, it
'         is not possible to delete all records, therefore there will always be
'         some records remaining in the recordset and combo.
'******************************************************************************
Public Function DeleteBomber() As Boolean
    On Error GoTo ErrorTrap
    
    Dim strAirmanList As String
    Dim strIgnore As String
    Dim intListIndex As Integer
    
    DeleteBomber = True

    With frmMainMenu
        
        ' Default bombers cannot be deleted.
        
        If prsBomber![Default] = True Then
            strErrmsg = "Failed to delete bomber." & vbCrLf & vbCrLf & _
                        prsBomber![Name] & _
                        " is a default bomber."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
    
        ' Determine if the bomber has assigned airmen.
        
        If HasAssignedAirman(strAirmanList) = True Then
            strErrmsg = "Failed to delete bomber." & vbCrLf & vbCrLf & _
                        strAirmanList & _
                        " are assigned to the " & _
                        prsBomber![Name] & _
                        ". A bomber cannot be deleted if it has " & _
                        "dependent airmen."
    
            MsgBox strErrmsg, (vbExclamation + vbOKOnly)
            
            ' This is not a severe system error, so return true.
            Exit Function
        End If
    
        ' Confirm the deletion.
    
        If MsgBox("Are you sure?", (vbYesNo + vbDefaultButton2 + vbQuestion)) = vbNo Then
            Exit Function
        End If

        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
    
            prsBomber.Delete
    
            prsBomber.UpdateBatch
        
        pobjConn.CommitTrans
        
        pintOpenTrans = pintOpenTrans - 1
            
        ' Delete from bomber tab name combo, then change the bomber tab
        ' to another listed bomber. Also delete from airman tab bomber
        ' combo. (Though it does not need to be re-indexed because no
        ' airmen, including the current one, were dependent on the
        ' deleted bomber.)
    
        intListIndex = .cboName(BOMBER_TAB).ListIndex
        
        .cboName(BOMBER_TAB).RemoveItem intListIndex
        .cboAssignment.RemoveItem intListIndex

        If intListIndex = 0 Then
            
            ' Deleted the first listed record. Point to the second record,
            ' which just became the first record.
            .cboName(BOMBER_TAB).ListIndex = 0
        
        ElseIf intListIndex = .cboName(BOMBER_TAB).ListCount Then
            
            ' Deleted the last listed record. Point to the second-to-last
            ' record, which just became the last record. ListIndex is 0-base,
            ' while ListCount is 1-base, so we need to subtract 1.
            .cboName(BOMBER_TAB).ListIndex = intListIndex - 1 '(.cboName(BOMBER_TAB).ListCount - 1)
        
        Else
            
            ' Deleted some record between the first and last records. Point
            ' to the record after the one that was deleted.
            .cboName(BOMBER_TAB).ListIndex = intListIndex
        
        End If
    
        If LookupBomber((.cboName(BOMBER_TAB).ListIndex + 1), LOOKUP_BY_LISTINDEX, strIgnore) = True Then
            If FillBomberTabFields() = False Then
' qwe            Call ExitEmulator
                DeleteBomber = False
                Exit Function
            End If
        Else
' qwe            Call ExitEmulator
            DeleteBomber = False
            Exit Function
        End If

    End With

    ' The deleted bomber will be removed from both the array of bomber
    ' keyfields and the mission tab bomber combo.
Call AdjustAvailableBombers ' Nov04
'    Call AdjustMissionAvailableBombers
    
    Exit Function

CleanUp:

    If pintOpenTrans Then
        pobjConn.RollbackTrans
        pintOpenTrans = pintOpenTrans - 1
    End If
    
    Exit Function

ErrorTrap:

    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "DeleteBomber() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrmsg, (vbCritical + vbOKOnly)

    Err.Clear

    DeleteBomber = False
    
    Resume CleanUp

End Function

'******************************************************************************
' IsDupeBomberName
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Multiple bombers may not have the exact same name.
'******************************************************************************
Private Function IsDupeBomberName() As Boolean

    IsDupeBomberName = False

    With frmMainMenu
    
        ' Determine if any other bomber has the same name.
         
        prsBomber.MoveFirst
        Do Until prsBomber.EOF
         
            If prsBomber![Name] = .cboName(BOMBER_TAB).Text Then
                IsDupeBomberName = True
                Exit Do
            End If
        
            prsBomber.MoveNext
        Loop
    
    End With
    
    ' Re-point the bomber recordset to the record currently on the bomber tab.

    prsBomber.Bookmark = varBomberCurrentlyOnTab

End Function

'******************************************************************************
' AdjustMissionAvailableBombers
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  If a bomber is put on duty status, it should be added to the bomber
'         list on the generate mission screen; if it goes off duty status, it
'         should be removed from the list.
'******************************************************************************
Public Sub AdjustMissionAvailableBombers()
    Dim intIndex As Integer
    Dim intTemp() As Integer
    
    With frmMainMenu
    
        intIndex = 0
    
        ' Clear the mission tab bomber combo so it can be rebuilt from
        ' scratch.

        If .cboName(MISSION_TAB).ListCount > 0 Then
            For intIndex = 0 To (.cboName(MISSION_TAB).ListCount - 1)
                .cboName(MISSION_TAB).RemoveItem 0
            Next intIndex
        End If
    
        intIndex = 0
    
        prsBomber.MoveFirst
        
        Do Until prsBomber.EOF
            
            ' All bombers that can be listed on the mission tab bomber combo
            ' -- i.e., those on duty or stand down status -- should be placed
            ' in the array. Bombers should only be removed from the array when
            ' they change to some other status.
            
            ' Bombers in the array should be placed in the mission tab bomber
            ' combo if they are on duty status and are not in flight. Bombers
            ' should be added to, or removed from, the combo as their states
            ' change.
            
            ' Only bombers available for missions should be listed on the
            ' mission tab. Keep an array of tail numbers matching the
            ' mission tab bomber list so that the bomber's information
            ' can be looked up in prsBomber if the bomber is chosen for a
            ' mission. The array is 0-base to match the 0-base of the
            ' combo's ListIndex.
            
            If prsBomber![Status] = DUTY_STATUS Then
            
                ' Add space to the array.
        
                ReDim Preserve intTemp(intIndex)
                intTemp(intIndex) = prsBomber![KeyField]

                .cboName(MISSION_TAB).AddItem prsBomber![Name]
                
                intIndex = intIndex + 1
            End If
            
            prsBomber.MoveNext
        Loop
    
        ' Any bombers whose state changed will be added to, or removed from,
        ' the array when intTemp is copied over intBomberMission.
        
        intBomberMission = intTemp

    End With

    If varBomberCurrentlyOnTab <> Empty Then
        prsBomber.Bookmark = varBomberCurrentlyOnTab
    End If

End Sub

'******************************************************************************
' AdjustAvailableBombers ' Nov04
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  If a bomber is put on duty status, it should be added to the bomber
'         list on the generate mission screen; if it goes off duty status, it
'         should be removed from the list.
'******************************************************************************
Public Sub AdjustAvailableBombers()
    Dim intIndex As Integer
    Dim intTemp() As Integer
    Dim intAssIndex As Integer
    Dim strAssignment As String
    
    With frmMainMenu
    
        intIndex = 0
    
        ' Clear the mission tab bomber combo so it can be rebuilt from
        ' scratch.

        If .cboName(MISSION_TAB).ListCount > 0 Then
            For intIndex = 0 To (.cboName(MISSION_TAB).ListCount - 1)
                .cboName(MISSION_TAB).RemoveItem 0
            Next intIndex
        End If
    
        If .cboAssignment.ListCount > 0 Then
            For intIndex = 0 To (.cboAssignment.ListCount - 1)
                .cboAssignment.RemoveItem 0
            Next intIndex
        End If
    
        intIndex = 0
    
        .cboAssignment.AddItem "Admin Duty"
        
        ' kgreer (12 Dec 04)
        intAssIndex = 0
        ReDim intMapBomberKeyToAssignmentIndex(intAssIndex)
        intMapBomberKeyToAssignmentIndex(intAssIndex) = 0
        
        prsBomber.MoveFirst
        
        Do Until prsBomber.EOF
            
            ' All bombers that can be listed on the mission tab bomber combo
            ' -- i.e., those on duty or stand down status -- should be placed
            ' in the array. Bombers should only be removed from the array when
            ' they change to some other status.
            
            ' Bombers in the array should be placed in the mission tab bomber
            ' combo if they are on duty status and are not in flight. Bombers
            ' should be added to, or removed from, the combo as their states
            ' change.
            
            ' Only bombers available for missions should be listed on the
            ' mission tab. Keep an array of tail numbers matching the
            ' mission tab bomber list so that the bomber's information
            ' can be looked up in prsBomber if the bomber is chosen for a
            ' mission. The array is 0-base to match the 0-base of the
            ' combo's ListIndex.
            
            If prsBomber![Status] = DUTY_STATUS Then
            
                ' Add space to the array.
        
                ReDim Preserve intTemp(intIndex)
                intTemp(intIndex) = prsBomber![KeyField]

                .cboName(MISSION_TAB).AddItem prsBomber![Name]
                .cboAssignment.AddItem prsBomber![Name]
                
                intIndex = intIndex + 1
            
                ' kgreer (12 Dec 04)
                intAssIndex = intAssIndex + 1
                ReDim Preserve intMapBomberKeyToAssignmentIndex(intAssIndex)
                intMapBomberKeyToAssignmentIndex(intAssIndex) = prsBomber![KeyField]
                
            ElseIf prsBomber![Status] = STAND_DOWN_STATUS _
            Or prsAirman![Assignment] = prsBomber![KeyField] Then
                
                .cboAssignment.AddItem prsBomber![Name]
            
                ' kgreer (12 Dec 04)
                intAssIndex = intAssIndex + 1
                ReDim Preserve intMapBomberKeyToAssignmentIndex(intAssIndex)
                intMapBomberKeyToAssignmentIndex(intAssIndex) = prsBomber![KeyField]
                
            End If
            
            prsBomber.MoveNext
        Loop
    
        ' Any bombers whose state changed will be added to, or removed from,
        ' the array when intTemp is copied over intBomberMission.
        
        intBomberMission = intTemp

        If IsNull(prsAirman("Assignment")) Then
            .cboAssignment.Text = "Admin Duty"
        Else
            If LookupBomber(prsAirman![Assignment], LOOKUP_BY_KEYFIELD, strAssignment) = True Then
                .cboAssignment.Text = strAssignment
'
'                ' Repoint the bomber recordset to the record displayed on the bomber tab.
'                prsBomber.Bookmark = varBomberCurrentlyOnTab
            End If
        End If
    
    End With

    If varBomberCurrentlyOnTab <> Empty Then
        prsBomber.Bookmark = varBomberCurrentlyOnTab
    End If

End Sub

'******************************************************************************
' RetireBomber
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Voluntarily remove the bomber from the game (i.e., send it home to
'         do war bond tours like the Memphis Belle).
'******************************************************************************
Public Function RetireBomber() As Boolean
    On Error GoTo ErrorTrap

    Dim strBomberStatus As String
    Dim intPos As Integer
    Dim intIndex As Integer
    Dim strErrmsg As String
    Dim strIgnore As String

    RetireBomber = True
    
    With frmMainMenu
        
'+++++++++++++++++++++++++++++++++++++
' TODO: update crew assignments to "ADMIN DUTY"

        Call InitializeBomber
        
        ' Sequentially point at each of the airmen assigned to the bomber by
        ' cycling through the crew's originally assigned positions.
        
        For intPos = PILOT To AMMO_STOCKER
            
            If PosOccupied(intPos) = True Then
                
                ' Airman currently in position
                intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
                
                ' Point at the airman.
                    
                If LookupAirman(Bomber.Airman(intIndex).SerialNumber, LOOKUP_BY_KEYFIELD, strIgnore) = True Then
                    prsAirman![Assignment] = ADMIN_DUTY ' Nov04
                End If
    
            End If ' not a hidden position
        
        Next intPos
        
'+++++++++++++++++++++++++++++++++++++
        
        prsBomber![Status] = RETIRED_STATUS
        
        pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
                
            prsAirman.UpdateBatch
            prsBomber.UpdateBatch
            
        pobjConn.CommitTrans
            
        pintOpenTrans = pintOpenTrans - 1
        
' TODO: what purpose does this serve?
'        If LookupBomberStatus(prsBomber![Status], strBomberStatus) = False Then
'            RetireBomber = False
'            Exit Function
'        Else
'            .txtStatus(BOMBER_TAB).Text = prsBomberStatus![Status] ' TODO: strBomberStatus
'        End If

'        .cboBomberModel(BOMBER_TAB).Enabled = False
'        .cboBomberModel(BOMBER_TAB).BackColor = vbButtonFace
'
'        .cboSquadron.Enabled = False
'        .cboSquadron.BackColor = vbButtonFace
'
'        .txtStatus(BOMBER_TAB).ForeColor = vbBlack
'        .txtStatus(BOMBER_TAB).BackColor = vbButtonFace
'
'        .cmdAssignCrew.Caption = "Last Crew"
'        .cmdRetireBomber.Visible = False

    End With

Call AdjustAvailableBombers ' Nov04
'    Call AdjustMissionAvailableBombers
    prsBomber.Bookmark = varBomberCurrentlyOnTab
    prsAirman.Bookmark = varAirmanCurrentlyOnTab

Call FillBomberTabFields ' Nov04
Call FillAirmanTabFields ' Nov04

    Exit Function

CleanUp:

    If pintOpenTrans Then
        pobjConn.RollbackTrans
        pintOpenTrans = pintOpenTrans - 1
    End If
    
    Call FreeRecordset(prsBomber)
    
    Exit Function

ErrorTrap:

    strErrmsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "RetireBomber() " & vbCrLf & vbCrLf & _
                Err.Description & vbCrLf & vbCrLf & _
                prsBomber!Name & vbCrLf & vbCrLf & _
                prsBomber!BomberModel & vbCrLf & vbCrLf & _
                prsBomber!Status


    MsgBox strErrmsg, (vbCritical + vbOKOnly)

    Err.Clear

    RetireBomber = False

    Resume CleanUp

End Function





