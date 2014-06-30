Attribute VB_Name = "modMission"
'******************************************************************************
' modMission.bas
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

Dim strErrMsg As String

'******************************************************************************
' InitializeMission
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Fill the mission structure and all its sub structures.
'******************************************************************************
Public Sub InitializeMission()

    Dim intIndex As Integer
    Dim intDateValue As Integer
    Dim intFinalCoverZone As Integer
    Dim intExtendedCoverZone As Integer

    With frmMainMenu
    
        Mission.TargetName = prsBomberTarget![Name]
        Mission.TargetType = prsBomberTarget![TargetType]
        Mission.HeavyFlak = prsBomberTarget![HeavyFlak]
        Mission.TargetZone = prsBomberTarget![TargetZone]
        Mission.Date = GetDateValue()
 
        ' Options
        Mission.Options.RandomEvents = .chkRandomEvents.Value
        Mission.Options.MechanicalFailures = .chkMechanicalFailures.Value
        Mission.Options.TimePeriodSpecificFormations = .chkTimePeriodSpecificFormations.Value
        Mission.Options.FormationDefensiveGunnery = .chkFormationDefensiveGunnery.Value
        Mission.Options.EvadeFlak = .chkEvadeFlak.Value
        Mission.Options.CrewExperience = .chkCrewExperience.Value
        Mission.Options.AlternateWeather = .chkAlternateWeather.Value
        Mission.Options.GermanFighterPilotSkill = .chkGermanFighterPilotSkill.Value
        Mission.Options.JG26StationedInAbbeville = .chkJG26StationedInAbbeville.Value
        Mission.Options.Ju88sUsedAsFighters = .chkJu88sUsedAsFighters.Value
        Mission.Options.Unescorted = .chkUnescorted.Value
        Mission.Options.RedTailAngels = .chkRedTailAngels.Value
        
        If RandomDX(7) = 1 Then
            ' There were seven fighter groups in the 15th Air Force, so there is
            ' a 1-in-7 chance that selecting this option will have an effect.
            Mission.Options.RedTailAngels = .chkRedTailAngels.Value
        End If
       
        ' The amount of delay, in milliseconds, after each line is printed
        ' to the mission log.
        Mission.Options.Delay = CInt(.txtLogSpeed.Text) * 100
        
        Mission.Zone(1).Modifier = 0
        
        Mission.Zone(1).Terrain = "Base"
        
        If Mission.TargetZone >= 2 Then
            Mission.Zone(2).Modifier = prsBomberTarget![Zone2Mod]
            Mission.Zone(2).Terrain = prsBomberTarget![Zone2Ter]
        End If
        
        If Mission.TargetZone >= 3 Then
            Mission.Zone(3).Modifier = prsBomberTarget![Zone3Mod]
            Mission.Zone(3).Terrain = prsBomberTarget![Zone3Ter]
        End If
        
        If Mission.TargetZone >= 4 Then
            Mission.Zone(4).Modifier = prsBomberTarget![Zone4Mod]
            Mission.Zone(4).Terrain = prsBomberTarget![Zone4Ter]
        End If
        
        If Mission.TargetZone >= 5 Then
            Mission.Zone(5).Modifier = prsBomberTarget![Zone5Mod]
            Mission.Zone(5).Terrain = prsBomberTarget![Zone5Ter]
        End If
        
        If Mission.TargetZone >= 6 Then
            Mission.Zone(6).Modifier = prsBomberTarget![Zone6Mod]
            Mission.Zone(6).Terrain = prsBomberTarget![Zone6Ter]
        End If
        
        If Mission.TargetZone >= 7 Then
            Mission.Zone(7).Modifier = prsBomberTarget![Zone7Mod]
            Mission.Zone(7).Terrain = prsBomberTarget![Zone7Ter]
        End If
        
        If Mission.TargetZone >= 8 Then
            Mission.Zone(8).Modifier = prsBomberTarget![Zone8Mod]
            Mission.Zone(8).Terrain = prsBomberTarget![Zone8Ter]
        End If
       
        If Mission.TargetZone >= 9 Then
            Mission.Zone(9).Modifier = prsBomberTarget![Zone9Mod]
            Mission.Zone(9).Terrain = prsBomberTarget![Zone9Ter]
        End If
       
        If Mission.TargetZone >= 10 Then
            Mission.Zone(10).Modifier = prsBomberTarget![Zone10Mod]
            Mission.Zone(10).Terrain = prsBomberTarget![Zone10Ter]
        End If
       
        If Mission.TargetZone >= 11 Then
            Mission.Zone(11).Modifier = prsBomberTarget![Zone11Mod]
            Mission.Zone(11).Terrain = prsBomberTarget![Zone11Ter]
        End If
       
        If Mission.TargetZone >= 12 Then
            Mission.Zone(12).Modifier = prsBomberTarget![Zone12Mod]
            Mission.Zone(12).Terrain = prsBomberTarget![Zone12Ter]
        End If
       
        Mission.AlpsZone = AlpsZone()
 
        intFinalCoverZone = GetFinalCoverZone()
       
        For intIndex = 2 To MAX_ZONE 'intFinalCoverZone

            If prsBomber![BomberModel] = AVRO_LANCASTER Then
' TODO: change after bomber structure is tested (change what ???)
                
                ' Lancasters flew unescorted night missions.
                Mission.Zone(intIndex).CoverOut = NO_COVER
                Mission.Zone(intIndex).CoverBack = NO_COVER
            ElseIf Mission.Options.Unescorted = True Then
                Mission.Zone(intIndex).CoverOut = NO_COVER
                Mission.Zone(intIndex).CoverBack = NO_COVER
            ElseIf intIndex <= intFinalCoverZone Then
                Mission.Zone(intIndex).CoverOut = G5FighterCover(Mission.Date, False)
                Mission.Zone(intIndex).CoverBack = G5FighterCover(Mission.Date, False)
            Else
                Mission.Zone(intIndex).CoverOut = NO_COVER
                Mission.Zone(intIndex).CoverBack = NO_COVER
            End If
        
'MsgBox "Mission.Zone(intIndex).CoverOut = " & Mission.Zone(intIndex).CoverOut & vbCrLf & _
       "Mission.Zone(intIndex).CoverBack = " & Mission.Zone(intIndex).CoverBack

        Next intIndex
        
' TODO:
'Oct-44: B-24s able to return as far as Zone 3 may make an emergency
'landing at an Allied tactical airfield in France. From this point forward,
'the B-2 table die roll modifier for zones in Germany (see above) only applies
'to Zone 10, and even then only when Berlin is the target.
                    
'Jan-45: In addition, German fighter waves will no longer be encountered over
'France. B-24s able to return as far as Zone 5 may make an emergency landing
'at an Allied tactical airfield in France.
                    
        If prsBomber![BomberModel] <> AVRO_LANCASTER Then  ' TODO: change after bomber structure is tested
            ' Sometimes a fighter pilot would use his last drop of gas to fly
            ' past his fighter's maximum range, but only on the outbound leg
            ' of a mission.
            
            If intFinalCoverZone < Mission.TargetZone _
            And Mission.TargetZone < 12 _
            And Mission.Options.Unescorted = False Then
                intExtendedCoverZone = intFinalCoverZone + 1
            
                Mission.Zone(intExtendedCoverZone).CoverOut = G5FighterCover(Mission.Date, True)
                Mission.Zone(intExtendedCoverZone).CoverBack = NO_COVER
            End If
        End If

        ' There is weather in every zone, including the base zone.
    
        For intIndex = BASE_ZONE To Mission.TargetZone
            Mission.Zone(intIndex).Weather = O1Weather()
            Mission.Zone(intIndex).Contrail = Contrail(intIndex)
        Next intIndex
    
    End With
End Sub

'******************************************************************************
' Contrail
'
' INPUT:  n/a
'
' OUTPUT: True if a contrail formed, otherwise false.
'
' RETURN: n/a
'
' NOTES:  In zones 2-12, if weather is clear, roll 1d6. On 5 or 6, contrails
'         form. If contrails form, apply +1 to tables B-1, B-2 and O-2. (The
'         variant states O-3, but that makes no sense as O-3 is a 2d6 table,
'         whereas O-2 is a 1d6 table with graduated results.) Contrails apply
'         to both outbound and return trip in the zone in which they are rolled.
'         Contrails do not apply to night missions (Lancasters). From the
'         "Theater Modifications" article in "The General" (Volume 24, #6).
'******************************************************************************
Private Function Contrail(ByVal intZone As Integer) As Boolean

    Contrail = False

'    If intZone >= 2 _
'    And Mission.Zone(intZone).Weather = CLEAR_WEATHER _
'    And Bomber.BomberModel <> AVRO_LANCASTER _
'    And Random1D6() >= 5 Then
    If intZone >= 2 _
    And Mission.Zone(intZone).Weather = CLEAR_WEATHER _
    And prsBomber![BomberModel] <> AVRO_LANCASTER _
    And Random1D6() >= 5 Then
        Contrail = True
    End If

End Function

'******************************************************************************
' InitializeBomber
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Fill the bomber structure and all its sub structures.
'******************************************************************************
Public Function InitializeBomber() As Boolean
    Dim strIgnore As String
    Dim intPos As Integer
    Dim intIndex As Integer
    Dim intZones As Integer
    
    InitializeBomber = True
    
    With frmMainMenu
        
        Bomber.Name = prsBomber![Name]
        Bomber.KeyField = prsBomber![KeyField]
        Bomber.BomberModel = prsBomber![BomberModel]
        Bomber.TailNumber = ""
        Bomber.Manufacturer = ""
        Bomber.Plant = ""
        Bomber.Mission = prsBomber![Sorties] + 1 ' next mission
        Bomber.Status = prsBomber![Status]
        Bomber.RabbitsFoot = prsBomber![RabbitsFoot]
        
        Bomber.TailNumber = prsBomber![TailNumber]
        Bomber.Manufacturer = prsBomber![Manufacturer]
        Bomber.Plant = prsBomber![Plant]
        
        If Bomber.BomberModel = AVRO_LANCASTER Then
            ' The bomber stream did not employ a defensive box, therefore all
            ' Lancasters are considered to be in the middle of the formation.
            Bomber.SquadronPos = MIDDLE_SQUADRON
            Bomber.FormationPos = MIDDLE_PLANE
        Else
            Bomber.SquadronPos = .cboSquadronPos.Text
            Bomber.FormationPos = .cboFormationPos.Text
        End If
        
        Bomber.CurrentZone = BASE_ZONE
        Bomber.Direction = OUTBOUND
        
        ' Lancasters are always in formation (i.e., in the bomber stream).
        ' Other bomber models may drop out of formation.
        
        Bomber.InFormation = True
        Bomber.Altitude = NO_ALTITUDE
        Bomber.TurnsInZone = 0
        
        ' Only Lancasters may be spotted by searchlights, since American
        ' bombers flew during daylight.
        
        Bomber.SpottedBySearchLight = False

        ' Copy all positions, whether that position actually exists on the
        ' bomber or not.
        
        Bomber.Position(PILOT).AssignedSerialNum = prsBomber![PILOT]
        Bomber.Position(COPILOT).AssignedSerialNum = prsBomber![COPILOT]
        Bomber.Position(BOMBARDIER).AssignedSerialNum = prsBomber![BOMBARDIER]
        Bomber.Position(NAVIGATOR).AssignedSerialNum = prsBomber![NAVIGATOR]
        Bomber.Position(ENGINEER).AssignedSerialNum = prsBomber![ENGINEER]
        Bomber.Position(RADIO_OPERATOR).AssignedSerialNum = prsBomber![RadioOperator]
        Bomber.Position(NOSE_GUNNER).AssignedSerialNum = prsBomber![NoseGunner]
        Bomber.Position(MID_UPPER_GUNNER).AssignedSerialNum = prsBomber![MidUpperGunner]
        Bomber.Position(BALL_GUNNER).AssignedSerialNum = prsBomber![BallGunner]
        Bomber.Position(PORT_WAIST_GUNNER).AssignedSerialNum = prsBomber![PortWaistGunner]
        Bomber.Position(STBD_WAIST_GUNNER).AssignedSerialNum = prsBomber![StbdWaistGunner]
        Bomber.Position(TAIL_GUNNER).AssignedSerialNum = prsBomber![TailGunner]
        Bomber.Position(AMMO_STOCKER).AssignedSerialNum = prsBomber![AmmoStocker]
        
        ' Fill in remaining airman data only for those positions that are not
        ' hidden or empty.
        
        For intPos = PILOT To AMMO_STOCKER
            
            ' Every airman -- even those that are hidden -- starts out in
            ' his assigned position. Non-hidden airman may be moved between
            ' non-hidden positions during the course of the mission. Airman
            ' currently at position:
            
            Bomber.Position(intPos).CurrentSerialNum = Bomber.Position(intPos).AssignedSerialNum
            
            If PosOccupied(intPos) = True Then
                
                ' Lookup the assigned airman, then initialize his data to the
                ' structure.
                If LookupAirman(Bomber.Position(intPos).AssignedSerialNum, LOOKUP_BY_KEYFIELD, strIgnore) = False Then
                    ' qwe Call ExitEmulator
                    InitializeBomber = False
                    Exit Function
                End If
            
                Bomber.Airman(intPos).Name = prsAirman![Name]
                Bomber.Airman(intPos).SerialNumber = prsAirman![KeyField]
                Bomber.Airman(intPos).Mission = prsAirman![Sorties] + 1 ' next mission
                Bomber.Airman(intPos).Kills = prsAirman![Kills]
                Bomber.Airman(intPos).Status = prsAirman![Status]
                Bomber.Airman(intPos).Wounded = False
                Bomber.Airman(intPos).Frostbite = False

'If intPos = NAVIGATOR Then Bomber.Airman(intPos).Status = SW_STATUS

                If IsNull(prsAirman![LeadCrewExp]) = False Then
                    Bomber.Airman(intPos).LeadCrewExp = prsAirman![LeadCrewExp]
                End If
                
'MsgBox "Bomber.Airman(intPos).LeadCrewExp = '" & Bomber.Airman(intPos).LeadCrewExp & "'"
'MsgBox "prsBomberTarget![KeyField] = '" & prsBomberTarget![KeyField] & "'"
                
                If prsAirman![Sorties] >= 11 _
                And (intPos = BOMBARDIER _
                Or intPos = NAVIGATOR) _
                And Bomber.SquadronPos = LOW_SQUADRON _
                And Bomber.FormationPos = LEAD_PLANE _
                And Mission.Date >= APR_1943 Then
                
                    ' The airman is an experienced bombardier or navigator,
                    ' flying in the lead bomber of the entire group, on or
                    ' after April, 1943, so the airman gains Lead Crew
                    ' Experience on this mission.
                    If Bomber.Airman(intPos).LeadCrewExp = "" Then
                        Bomber.Airman(intPos).LeadCrewExp = prsBomberTarget![KeyField]
                    Else
                        Bomber.Airman(intPos).LeadCrewExp = Bomber.Airman(intPos).LeadCrewExp & "|" & prsBomberTarget![KeyField]
                    End If
                End If

'MsgBox "Bomber.Airman(intPos).LeadCrewExp = '" & Bomber.Airman(intPos).LeadCrewExp & "'"
                
                ' The airman's initial position is the one he is assigned to.
                ' He may be moved to another position during the course of the
                ' mission. The index never changes, and so always indicates the
                ' airman's normal position.
                Bomber.Airman(intPos).AssignedPosition = prsAirman![CrewPosition]
            
            End If
        
            ' Heaters and oxygen can't moved during the game, so they are
            ' assigned to the position.
            
            Damage.Heater(intPos) = False
            Damage.Oxygen(intPos) = 0
            
        Next intPos

        prsAirman.Bookmark = varAirmanCurrentlyOnTab

        ' If the position is hidden -- i.e., does not exist for the particular
        ' bomber model -- then copying the airman's serial number to gun's
        ' manned by field will copy the hidden-ness. Thus, the gun is marked
        ' as non-existant. Manned by should not be empty unless an airman is
        ' wounded or moved to another position while in flight. A hidden
        ' position or gun can never be unhidden.
        
        ' --------------------------------------------------
        
        If Bomber.BomberModel = YB40 Then
            Bomber.Gun(MID_UPPER_MG).Bonus = 1
        Else
            Bomber.Gun(MID_UPPER_MG).Bonus = 0
        End If
        
        Bomber.Gun(MID_UPPER_MG).Ammo = 16
        Bomber.Gun(MID_UPPER_MG).MaxAmmo = 16
        
        ' The gun is either manned or hidden.
        Bomber.Gun(MID_UPPER_MG).MannedBy = Bomber.Position(MID_UPPER_GUNNER).AssignedSerialNum
        
        ' --------------------------------------------------
        
        If Bomber.BomberModel = B17_G _
        Or Bomber.BomberModel = YB40 _
        Or Bomber.BomberModel = B24_GHJ _
        Or Bomber.BomberModel = B24_LM Then
            Bomber.Gun(NOSE_MG).Bonus = 1
        Else
            Bomber.Gun(NOSE_MG).Bonus = 0
        End If

        If Bomber.BomberModel = YB40 Then
            Bomber.Gun(NOSE_MG).Ammo = 30
            Bomber.Gun(NOSE_MG).MaxAmmo = 30
        Else
            Bomber.Gun(NOSE_MG).Ammo = 15
            Bomber.Gun(NOSE_MG).MaxAmmo = 15
        End If

        If Bomber.BomberModel = YB40 _
        Or Bomber.BomberModel = B24_GHJ _
        Or Bomber.BomberModel = B24_LM Then
            ' The gun is manned.
            Bomber.Gun(NOSE_MG).MannedBy = Bomber.Position(NOSE_GUNNER).AssignedSerialNum
        Else
            ' The gun is either manned or hidden.
            Bomber.Gun(NOSE_MG).MannedBy = Bomber.Position(BOMBARDIER).AssignedSerialNum
        End If
        
        ' --------------------------------------------------
        
        Bomber.Gun(STBD_CHEEK_MG).Bonus = 0
        Bomber.Gun(STBD_CHEEK_MG).Ammo = 10
        Bomber.Gun(STBD_CHEEK_MG).MaxAmmo = 10
        
        Bomber.Gun(PORT_CHEEK_MG).Bonus = 0
        Bomber.Gun(PORT_CHEEK_MG).Ammo = 10
        Bomber.Gun(PORT_CHEEK_MG).MaxAmmo = 10
        
        If Bomber.BomberModel = B17_F _
        Or Bomber.BomberModel = B17_G _
        Or Bomber.BomberModel = YB40 _
        Or Bomber.BomberModel = B24_E Then
            ' The navigator starts out manning the stbd cheek gun.
            Bomber.Gun(STBD_CHEEK_MG).MannedBy = Bomber.Position(NAVIGATOR).AssignedSerialNum
            Bomber.Gun(PORT_CHEEK_MG).MannedBy = UNMANNED_MG
        Else
            ' The gun doesn't exist, it is hidden.
            Bomber.Gun(STBD_CHEEK_MG).MannedBy = HIDDEN_MG
            Bomber.Gun(PORT_CHEEK_MG).MannedBy = HIDDEN_MG
        End If
        
        ' --------------------------------------------------
        
        Bomber.Gun(TOP_TURRET_MG).Bonus = 1
        
        If Bomber.BomberModel = YB40 Then
            Bomber.Gun(TOP_TURRET_MG).Ammo = 32
            Bomber.Gun(TOP_TURRET_MG).MaxAmmo = 32
        Else
            Bomber.Gun(TOP_TURRET_MG).Ammo = 16
            Bomber.Gun(TOP_TURRET_MG).MaxAmmo = 16
        End If

        If Bomber.BomberModel = B17_C _
        Or Bomber.BomberModel = AVRO_LANCASTER Then
            ' The gun doesn't exist, it is hidden.
            Bomber.Gun(TOP_TURRET_MG).MannedBy = HIDDEN_MG
        Else
            ' The gun is manned by the engineer.
            Bomber.Gun(TOP_TURRET_MG).MannedBy = Bomber.Position(ENGINEER).AssignedSerialNum
        End If
        
        ' --------------------------------------------------
        
        If Bomber.BomberModel = B17_C Then
            Bomber.Gun(RADIO_ROOM_MG).Bonus = 0
            Bomber.Gun(RADIO_ROOM_MG).Ammo = 15
            Bomber.Gun(RADIO_ROOM_MG).MaxAmmo = 15
        Else
            Bomber.Gun(RADIO_ROOM_MG).Bonus = 0
            Bomber.Gun(RADIO_ROOM_MG).Ammo = 10
            Bomber.Gun(RADIO_ROOM_MG).MaxAmmo = 10
        End If
        
        ' All bomber models had a radio operator (thus the position is never
        ' hidden), however only the B-17 C, E and F actually had a radio room
        ' MG. Therefore, even though the position is not hidden, the weapon
        ' must be hidden.
        
        If Bomber.BomberModel = B17_C _
        Or Bomber.BomberModel = B17_E _
        Or Bomber.BomberModel = B17_F Then
            ' The gun is manned.
            Bomber.Gun(RADIO_ROOM_MG).MannedBy = Bomber.Position(RADIO_OPERATOR).AssignedSerialNum
        Else
            ' The gun is hidden.
            Bomber.Gun(RADIO_ROOM_MG).MannedBy = HIDDEN_MG
        End If
        
        ' --------------------------------------------------
        
        If Bomber.BomberModel = B24_D _
        Or Bomber.BomberModel = B24_E Then
            Bomber.Gun(BALL_TURRET_MG).Bonus = 0
        Else
            ' B-17C had twin guns in the tunnel position.
            Bomber.Gun(BALL_TURRET_MG).Bonus = 1
        End If

        If Bomber.BomberModel = YB40 Then
            Bomber.Gun(BALL_TURRET_MG).Ammo = 24
            Bomber.Gun(BALL_TURRET_MG).MaxAmmo = 24
        ElseIf Bomber.BomberModel = B17_C Then
            Bomber.Gun(BALL_TURRET_MG).Ammo = 18
            Bomber.Gun(BALL_TURRET_MG).MaxAmmo = 18
        Else
            Bomber.Gun(BALL_TURRET_MG).Ammo = 20
            Bomber.Gun(BALL_TURRET_MG).MaxAmmo = 20
        End If
        
        ' The gun is either manned or hidden.
        Bomber.Gun(BALL_TURRET_MG).MannedBy = Bomber.Position(BALL_GUNNER).AssignedSerialNum
        
        ' --------------------------------------------------
        
        If Bomber.BomberModel = YB40 Then
            Bomber.Gun(PORT_WAIST_MG).Bonus = 1
        Else
            Bomber.Gun(PORT_WAIST_MG).Bonus = 0
        End If
        
        If Bomber.BomberModel = B17_C Then
            Bomber.Gun(PORT_WAIST_MG).Ammo = 15
            Bomber.Gun(PORT_WAIST_MG).MaxAmmo = 15
        Else
            Bomber.Gun(PORT_WAIST_MG).Ammo = 20
            Bomber.Gun(PORT_WAIST_MG).MaxAmmo = 20
        End If
        
        ' B-24s had port and starboard waist guns, but only one waist gunner.
        ' At the start of a mission, the waist gunner is assumed to be manning
        ' the starboard gun. Because the position is not hidden on a B-24, we
        ' must explicity mark it as empty.
        
        If Bomber.BomberModel = B24_D _
        Or Bomber.BomberModel = B24_E _
        Or Bomber.BomberModel = B24_GHJ _
        Or Bomber.BomberModel = B24_LM Then
            ' The waist gunner starts out manning the stbd waist gun.
            Bomber.Gun(PORT_WAIST_MG).MannedBy = UNMANNED_MG
        Else
            ' The gun is either manned or hidden.
            Bomber.Gun(PORT_WAIST_MG).MannedBy = Bomber.Position(PORT_WAIST_GUNNER).AssignedSerialNum
        End If
        
        ' --------------------------------------------------
        
        If Bomber.BomberModel = YB40 Then
            Bomber.Gun(STBD_WAIST_MG).Bonus = 1
        Else
            Bomber.Gun(STBD_WAIST_MG).Bonus = 0
        End If
        
        If Bomber.BomberModel = B17_C Then
            Bomber.Gun(STBD_WAIST_MG).Ammo = 15
            Bomber.Gun(STBD_WAIST_MG).MaxAmmo = 15
        Else
            Bomber.Gun(STBD_WAIST_MG).Ammo = 20
            Bomber.Gun(STBD_WAIST_MG).MaxAmmo = 20
        End If
        
        ' The gun is either manned or hidden.
        Bomber.Gun(STBD_WAIST_MG).MannedBy = Bomber.Position(STBD_WAIST_GUNNER).AssignedSerialNum
        
        ' --------------------------------------------------
        
        Bomber.Gun(TAIL_MG).Bonus = 1
        
        If Bomber.BomberModel = YB40 Then
            Bomber.Gun(TAIL_MG).Ammo = 25
            Bomber.Gun(TAIL_MG).MaxAmmo = 25
        ElseIf Bomber.BomberModel = AVRO_LANCASTER Then
            Bomber.Gun(TAIL_MG).Ammo = 63
            Bomber.Gun(TAIL_MG).MaxAmmo = 63
        Else
            Bomber.Gun(TAIL_MG).Ammo = 23
            Bomber.Gun(TAIL_MG).MaxAmmo = 23
        End If
        
        ' The gun is either manned or hidden.
        Bomber.Gun(TAIL_MG).MannedBy = Bomber.Position(TAIL_GUNNER).AssignedSerialNum
        
        ' --------------------------------------------------
        
        ' Initialize Guns
        For intIndex = MID_UPPER_MG To TAIL_MG
            Bomber.Gun(intIndex).Status = MG_OKAY
            Bomber.Gun(intIndex).QualifiedGunner = True
        Next
        
        ' --------------------------------------------------
        
        Bomber.HandHeldExtinguishers = 5
        
        Bomber.EngineExtinguisher(1) = 2
        Bomber.EngineExtinguisher(2) = 2
        Bomber.EngineExtinguisher(3) = 2
        Bomber.EngineExtinguisher(4) = 2

        ' --------------------------------------------------
        
        ' Regardless of model, the bomber had to carry enough fuel.
        
        Bomber.FuelPoints = Mission.TargetZone * 2 '(Mission.TargetZone - 1) * 2
        
        ' YB-40s normally have a range of six zones. If they have a more
        ' distant mission, they must carry extra fuel in the bomb bay.
        
        If Bomber.BomberModel = YB40 _
        And Mission.TargetZone >= 7 Then
            Bomber.ExtraFuelInBombBay = True
        Else
            Bomber.ExtraFuelInBombBay = False
        End If
        
        ' If a YB-40 isn't carrying extra fuel, it may carry extra ammo in
        ' the bomb bay.
        
        If .chkExtraAmmoInBombBay.Value = vbChecked _
        And Bomber.ExtraFuelInBombBay = False Then
            Bomber.ExtraAmmo = 140
        Else
            Bomber.ExtraAmmo = 0
        End If
        
        ' The YB-40 didn't carry bombs -- it was a gunship -- but other
        ' models have to carry bombs.
        
        If Bomber.BomberModel = YB40 Then
            Bomber.BombsOnBoard = False
        Else
            Bomber.BombsOnBoard = True
        End If

'++++++++++++++++++++
        ' TODO: Update database, also put this in a separate function.
        
        If .txtKeyField(BOMBER_TAB).Text = prsBomber![KeyField] Then
            ' The bomber is currently being displayed on the bomber tab.
            ' Disable the tab fields, by re-filling them.
            If FillBomberTabFields() = False Then
                ' qwe Call ExitEmulator
                InitializeBomber = False
                Exit Function
            End If
        End If
'++++++++++++++++++++

    End With

    Call InitializeDamage

'MsgBox "Bomber Info" & vbCrLf & _
       "-----------" & vbCrLf & _
       "Name = " & Bomber.Name & vbCrLf & _
       "KeyField = " & Bomber.KeyField & vbCrLf & _
       "BomberModel = " & Bomber.BomberModel & vbCrLf & _
       "Mission = " & Bomber.Mission & vbCrLf & _
       "Status = " & Bomber.Status & vbCrLf & _
       "SquadronPos = " & Bomber.SquadronPos & vbCrLf & _
       "FormationPos = " & Bomber.FormationPos & vbCrLf & _
       "CurrentZone = " & Bomber.CurrentZone & vbCrLf & _
       "InFormation = " & Bomber.InFormation & vbCrLf & _
       "ExtraAmmo = " & Bomber.ExtraAmmo

'MsgBox "Bomber.Airman Info" & vbCrLf & _
       "------------------" & vbCrLf & _
       "Name = " & Bomber.Airman(NOSE_GUNNER).Name & vbCrLf & _
       "SerialNumber = " & Bomber.Airman(NOSE_GUNNER).SerialNumber & vbCrLf & _
       "Mission = " & Bomber.Airman(NOSE_GUNNER).Mission & vbCrLf & _
       "Kills = " & Bomber.Airman(NOSE_GUNNER).Kills & vbCrLf & _
       "Status = " & Bomber.Airman(NOSE_GUNNER).Status & vbCrLf & _
       "Frostbite = " & Bomber.Airman(NOSE_GUNNER).Frostbite & vbCrLf & vbCrLf & _
       "Bomber.Position Info" & vbCrLf & _
       "--------------------" & vbCrLf & _
       "Serial Number = " & Bomber.Position(NOSE_GUNNER).CurrentSerialNum & vbCrLf & _
       "Heater = " & Damage.Heater(NOSE_GUNNER) & vbCrLf & _
       "Oxygen = " & Damage.Oxygen(NOSE_GUNNER) & vbCrLf & vbCrLf & _
       "Bomber.Gun Info" & vbCrLf & _
       "---------------" & vbCrLf & _
       "Bonus = " & Bomber.Gun(NOSE_MG).Bonus & vbCrLf & _
       "Ammo = " & Bomber.Gun(NOSE_MG).Ammo & vbCrLf & _
       "Status = " & Bomber.Gun(NOSE_MG).Status & vbCrLf & _
       "MannedBy = " & Bomber.Gun(NOSE_MG).MannedBy & vbCrLf & _
       "QualifiedGunner = " & Bomber.Gun(NOSE_MG).QualifiedGunner

' TEST xyz
'If Bomber.Gun(1).Ammo >= 1 Then
'    Dim intDoo As Integer
'    For intDoo = MID_UPPER_MG To TAIL_MG
'        Bomber.Gun(intDoo).Ammo = 0
'    Next intDoo
'End If
    
End Function

'******************************************************************************
' InitializeDamage
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Fill the damage structure and all its sub structures.
'******************************************************************************
Public Sub InitializeDamage()
    
    ' Regardless of the bomber model, and what types of equipment it may
    ' posess, the equipment always starts out in working order. Therefore,
    ' all damage is initialized to nothing.
    
    Damage.PeckhamPoints = 0
    Damage.BombSight = False
    Damage.NoseWheel = False
    Damage.NavigationEquipment = False
    Damage.BombControls = False
    Damage.Window = 0
    Damage.ControlCables = 0
    Damage.RubberRafts = False
    Damage.BombRelease = False
    Damage.BombBayDoors = False
    Damage.IntercomSystem = False
    Damage.OxygenSystem = False
    Damage.WingFlapControls = False
    Damage.AileronControls = False
    Damage.ElevatorControls = False
    Damage.RudderControls = False
    Damage.Radio = False
    Damage.BallTurretMech = False
    Damage.Tailwheel = False
    Damage.Autopilot = False
    
    Damage.Rudder(PORT_SIDE) = 0
    Damage.Rudder(STBD_SIDE) = 0
    Damage.TailplaneRoot(PORT_SIDE) = 0
    Damage.TailplaneRoot(STBD_SIDE) = 0
    Damage.WingRoot(PORT_SIDE) = 0
    Damage.WingRoot(STBD_SIDE) = 0
    Damage.WingFlap(PORT_SIDE) = False
    Damage.WingFlap(STBD_SIDE) = False
    Damage.Aileron(PORT_SIDE) = False
    Damage.Aileron(STBD_SIDE) = False
    Damage.Elevator(PORT_SIDE) = False
    Damage.Elevator(STBD_SIDE) = False
    Damage.EngineOut(1) = False
    Damage.EngineOut(2) = False
    Damage.EngineOut(3) = False
    Damage.EngineOut(4) = False
    Damage.EngineDrag(1) = False
    Damage.EngineDrag(2) = False
    Damage.EngineDrag(3) = False
    Damage.EngineDrag(4) = False
    Damage.OilTankLeak(1) = NO_LEAK
    Damage.OilTankLeak(2) = NO_LEAK
    Damage.OilTankLeak(3) = NO_LEAK
    Damage.OilTankLeak(4) = NO_LEAK
    Damage.FuelTankHits = 0
    Damage.FuelTransferSystem = False
    Damage.HeatingSystem = False
    Damage.Turbocharger(1) = False
    Damage.Turbocharger(2) = False
    Damage.Turbocharger(3) = False
    Damage.Turbocharger(4) = False
    
    Damage.BurstInPlane = False
    
'+++++++++++++++++++++++++++++++++++++++
    
    Damage.Brake = False
    Damage.LandingGear = False
    Damage.FeatheringCtrl = False
    Damage.EngineExtCtrl = False
    Damage.Electrical = False
    Damage.PortAmmoTrack = False
    Damage.StbdAmmoTrack = False
    Damage.AmmoTrackFirings = 0
    Damage.PortAmmoBox = 0
    Damage.StbdAmmoBox = 0
    Damage.AmmoBoxFirings = 0

    Damage.LastZone = 0
End Sub

'******************************************************************************
' InitializeRandomEvents
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Public Sub InitializeRandomEvents()

    RandomEvent.AceForADay = False ' TODO: add effecting logic
    RandomEvent.AggressiveCover = False
    RandomEvent.BadLuftwaffeComm = False
    RandomEvent.EngineFailure = 0
    RandomEvent.ExtremeCold = False
    RandomEvent.FormationCasualties = False
    RandomEvent.LooseFormation = False
    RandomEvent.MidAirAccident = False
    RandomEvent.TightFormation = False

End Sub

'******************************************************************************
' GetFinalCoverZone
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The final cover zone if succesful, otherwise 0.
'
' NOTES:  Fighter cover extends to the target zone or the maximum fighter
'         range, whichever is less. From the B-24 variant included in this
'         distribution.
'******************************************************************************
Public Function GetFinalCoverZone() As Integer

    GetFinalCoverZone = 0
    
    Select Case Mission.Date
        Case AUG_1942 To NOV_1943:
            GetFinalCoverZone = 4
        Case DEC_1943 To MAY_1944:
            GetFinalCoverZone = 5
        Case JUN_1944 To SEP_1944:
            GetFinalCoverZone = 7
        Case OCT_1944 To DEC_1944:
            GetFinalCoverZone = 9
        Case JAN_1945 To MAY_1945:
            GetFinalCoverZone = 11
    End Select
        
    If Mission.TargetZone <= GetFinalCoverZone Then
         GetFinalCoverZone = Mission.TargetZone
    End If
        
'MsgBox "GetFinalCoverZone()" & vbCrLf & _
       "GetFinalCoverZone = " & GetFinalCoverZone & vbCrLf & _
       "Mission.TargetZone = " & Mission.TargetZone & vbCrLf & _
       "GetFinalCoverZone = " & GetFinalCoverZone

End Function

'******************************************************************************
' GetDateValue
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Number that represents some month between August, 1942, and May, 1945.
'
' NOTES:  Convert the text in the month and year combos into a numeric value.
'******************************************************************************
Public Function GetDateValue() As Integer
    GetDateValue = -1
    
    With frmMainMenu
        
        Select Case .cboMonth.Text
            Case "January":
                GetDateValue = -7
            Case "February":
                GetDateValue = -6
            Case "March":
                GetDateValue = -5
            Case "April":
                GetDateValue = -4
            Case "May":
                GetDateValue = -3
            Case "June":
                GetDateValue = -2
            Case "July":
                GetDateValue = -1
            Case "August":
                GetDateValue = 0
            Case "September":
                GetDateValue = 1
            Case "October":
                GetDateValue = 2
            Case "November":
                GetDateValue = 3
            Case "December":
                GetDateValue = 4
        End Select
    
        Select Case .cboYear.Text
            Case "1943":
                GetDateValue = GetDateValue + 12
            Case "1944":
                GetDateValue = GetDateValue + 24
            Case "1945":
                GetDateValue = GetDateValue + 36
        End Select
    
'MsgBox "GetDateValue()" & vbCrLf & _
       "Month = '" & .cboMonth.Text & "'" & vbCrLf & _
       "Year = '" & .cboYear.Text & "'" & vbCrLf & _
       "GetDateValue = " & GetDateValue

    End With

End Function

