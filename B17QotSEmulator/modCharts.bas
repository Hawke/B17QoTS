Attribute VB_Name = "modCharts"
'******************************************************************************
' modCharts.bas
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

'******************************************************************************
' G1MissionTarget
'
' INPUT:  The number of rows that appear in the target list box on the mission
'         tab of the Main Menu form.
'
' OUTPUT: n/a
'
' RETURN: The name of the randomly chosen target city.
'
' NOTES:  n/a
'******************************************************************************
Public Function G1MissionTarget(ByVal intListCount As Integer) As String

    ' G-1 Mission Targets (Missions 1-5)
    ' G-2 Mission Targets (Missions 6-10)
    ' G-3 Mission Targets (Missions 11-25)
    '
    ' Plus variants from "The General" (Volume 23, #5), "The General" (Volume
    ' 28, #4), "The General" (Volume 23, #1), the B-24 variant included
    ' in this distribution and new targets added specifcically for the
    ' emulator.
    '
    ' PopulateTargetCombo() has already determined the set of targets, so all
    ' we need to do is pick one from the list.
    
    G1MissionTarget = RandomDX(intListCount) - 1

End Function

'******************************************************************************
' G4SquadronPosition
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The randomly chosen squadron position.
'
' NOTES:  n/a
'******************************************************************************
Public Function G4SquadronPosition() As String
    
    Select Case Random1D6()
        Case 1 To 2:
            G4SquadronPosition = "High"
        Case 3 To 4:
            G4SquadronPosition = "Middle"
        Case 5 To 6:
            G4SquadronPosition = "Low"
    End Select

End Function

'******************************************************************************
' G4FormationPosition
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The randomly chosen formation position.
'
' NOTES:  n/a
'******************************************************************************
Public Function G4FormationPosition() As String
    
    Select Case Random2D6()
        Case 2:
            G4FormationPosition = "Lead"
        Case 3 To 11:
            G4FormationPosition = "Middle"
        Case 12:
            G4FormationPosition = "Tail"
    End Select

End Function

'******************************************************************************
' G5FighterCover
'
' INPUT:  The mission date, and whether or not fighters are pushing cover past
'         their range.
'
' OUTPUT: n/a
'
' RETURN: The randomly chosen level of fighter cover.
'
' NOTES:  G5FighterCover determines how good coverage is. M4FighterCoverDefense
'         determines how many German fighters are chased away.
'******************************************************************************
Public Function G5FighterCover(intDate As Integer, blnExtendedCover As Boolean) As String
    Dim intRoll As Integer
    
    intRoll = Random1D6()

    ' Fighter cover, regardless of range, improved as the war progressed.
    
    Select Case intDate
        Case AUG_1942 To NOV_1943:
            intRoll = intRoll + 0
        Case DEC_1943 To MAY_1944:
            intRoll = intRoll + 1
        Case JUN_1944 To SEP_1944:
            ' Fighters flying tactical support in Normandy.
            intRoll = intRoll + 0
        Case OCT_1944 To DEC_1944:
            intRoll = intRoll + 2 ' TODO: +1 ???
        Case JAN_1945 To MAY_1945:
            intRoll = intRoll + 2
    End Select

    If Mission.Options.RedTailAngels = True Then
    
        G5FighterCover = "Good"
    
    ElseIf blnExtendedCover = False Then
        
        Select Case intRoll
            Case Is <= 2:
                G5FighterCover = "Poor"
            Case 3 To 4:
                G5FighterCover = "Fair"
            Case Is >= 5:
                G5FighterCover = "Good"
        End Select
    
    Else
        
        ' A few friendly fighters may fly past normal range.
        
        If intRoll = 6 Then
            G5FighterCover = "Poor"
        Else
            G5FighterCover = "None"
        End If
    
    End If

End Function

'******************************************************************************
' G6ControlledBailout
'
' INPUT:  Whether or not the bomber is over water.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Public Sub G6ControlledBailout(ByVal blnOverWater As Boolean)
    Dim intRoll As Integer
    Dim intPos As Integer
    Dim intIndex As Integer
    
    If blnOverWater = True Then
        UpdateMessage "Controlled bailout over water:"
    Else
        UpdateMessage "Controlled bailout over land:"
    End If
    
    ' Cycle through the crew's originally assigned positions.
    For intPos = PILOT To AMMO_STOCKER
        
        If PosOccupied(intPos) = True Then
        
            ' Airman currently in position
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
        
            If Bomber.Airman(intIndex).Status = SW_STATUS _
            Or (intPos = BALL_GUNNER And Damage.BallTurretMech = True) Then
                
                ' Note b: Seriously wounded airman cannot bailout.
                ' P-5 Waist: Note c. Ball gunner can't bail out of inoperative
                ' turret.
                
                Bomber.Airman(intIndex).Status = MIA_STATUS
                UpdateMessage Bomber.Airman(intIndex).Name & " cannot bailout."
            
            ElseIf Bomber.Airman(intIndex).Status = KIA_STATUS Then
                
                UpdateMessage Bomber.Airman(intIndex).Name & " already dead."
            
            Else
                
                intRoll = Random1D6()

                If Bomber.Airman(intIndex).Status = LW2_STATUS Then
                    intRoll = intRoll - 1
                End If
                
                Select Case intRoll
                    
                    Case 1:
                        
                        If Random1D6() = 6 Then
                            Bomber.Airman(intIndex).Status = KIA_STATUS
                            Bomber.Airman(intIndex).Wounded = True
                            UpdateMessage Bomber.Airman(intIndex).Name & " killed in accident."
                        Else
                            UpdateMessage Bomber.Airman(intIndex).Name & " bailed out OK."
                        End If
                    
                    Case Else
                        
                        UpdateMessage Bomber.Airman(intIndex).Name & " bailed out OK."
                
                        If blnOverWater = True Then
                            Call G8BailoutOverWater(intIndex)
                        End If
    
                End Select

            End If
            
        End If
    
    Next intPos

    Bomber.Status = SHOT_DOWN_STATUS

    If blnOverWater = False Then
        If EnemyTerritory() = True Then
            Call CrewCaptured
        End If
    End If
    
End Sub

'******************************************************************************
' G7UncontrolledBailout
'
' INPUT:  Whether or not the bomber is over water.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub G7UncontrolledBailout(ByVal blnOverWater As Boolean)
    Dim intRoll As Integer
    Dim intPos As Integer
    Dim intIndex As Integer
    
    If blnOverWater = True Then
        UpdateMessage "Uncontrolled bailout over water:"
    Else
        UpdateMessage "Uncontrolled bailout over land:"
    End If
    
    ' Cycle through the crew's originally assigned positions.
    For intPos = PILOT To AMMO_STOCKER
        
        If PosOccupied(intPos) = True Then
        
            ' Airman currently in position
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
            
            If Bomber.Airman(intIndex).Status = SW_STATUS _
            Or (intPos = BALL_GUNNER And Damage.BallTurretMech = True) Then
                
                ' Note b: Seriously wounded airman cannot bailout.
                ' P-5 Waist: Note c. Ball gunner can't bail out of inoperative
                ' turret.
                
                Bomber.Airman(intIndex).Status = MIA_STATUS
                UpdateMessage Bomber.Airman(intIndex).Name & " cannot bailout."
            
            ElseIf Bomber.Airman(intIndex).Status = KIA_STATUS Then
                
                UpdateMessage Bomber.Airman(intIndex).Name & " already dead."
            
            Else
                
                intRoll = Random1D6()

                Select Case intRoll
                    
                    Case 1 To 5:
                        
                        Bomber.Airman(intIndex).Status = KIA_STATUS
                        Bomber.Airman(intIndex).Wounded = True
                        UpdateMessage Bomber.Airman(intIndex).Name & " goes down with plane."
                    
                    Case 6:
                        
                        ' Note d: A roll of 6 is always "Bailout OK", even if
                        ' the airman is LW2.
                        
                        UpdateMessage Bomber.Airman(intIndex).Name & " bailed out OK."
                
                        If blnOverWater = True Then
                            Call G8BailoutOverWater(intIndex)
                        End If
    
                End Select
            
            End If
            
        End If
    
    Next intPos

    Bomber.Status = SHOT_DOWN_STATUS

    If blnOverWater = False Then
        If EnemyTerritory() = True Then
            Call CrewCaptured
        End If
    End If
    
End Sub

'******************************************************************************
' G8BailoutOverWater
'
' INPUT:  An airman's position index.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub G8BailoutOverWater(ByVal intIndex As Integer)
    Dim intRoll As Integer
'    Dim intIndex As Integer
' zxc
'    ' Airman currently in position
'    intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
    
    If Damage.Radio = True Then
                    
        Bomber.Airman(intIndex).Status = MIA_STATUS
        UpdateMessage Bomber.Airman(intIndex).Name & " missing in action."
        
    Else
        
        intRoll = Random1D6()

        If Bomber.Airman(intIndex).Status = LW2_STATUS Then
            intRoll = intRoll - 1
        End If
        
        Select Case intRoll
                
            Case 1 To 4:
                    
                Bomber.Airman(intIndex).Status = KIA_STATUS
                Bomber.Airman(intIndex).Wounded = True
                UpdateMessage Bomber.Airman(intIndex).Name & " dies of drowning or exposure."
                
            Case 5 To 6:
                    
                UpdateMessage Bomber.Airman(intIndex).Name & " rescued by friendly ship."
            
        End Select
    
    End If
    
End Sub

'******************************************************************************
' G9LandingOnLand
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This function is called over both friendly and enemy territory.
'******************************************************************************
Public Sub G9LandingOnLand()
    Dim intRoll As Integer
    Dim blnPayloadExploded As Boolean
    Dim intEnginesOut As Integer
    
    UpdateMessage "Landing on land:"
        
    
    If Mission.Zone(Bomber.CurrentZone).Terrain = ALPS_TER Then
        
        ' There's no place to land in the Alps, so the bomber
        ' automatically crashes into the side of a mountain.
        
        Bomber.Status = CRASHED_STATUS
        UpdateMessage "Bomber crashes into the side of a mountain."
        Call CrewFinish(KIA_STATUS)
        Exit Sub
    
    End If
        
    intRoll = Random2D6()
    
    If intRoll = 12 Then
        
        ' Note a: Miracle roll.
        
        If EnemyTerritory() = True Then
            Bomber.Status = CAPTURED_STATUS
            Call CrewCaptured
            UpdateMessage "Bomber captured by enemy."
        Else
            Bomber.Status = DUTY_STATUS
            UpdateMessage "Bomber made a safe landing."
        End If
                    
    Else
        
        ' O-1 Weather: Note a and Note b, plus "The General" (Volume 24, #6)
        ' variant.
        
        If Mission.Zone(BASE_ZONE).Weather = POOR_WEATHER Then
            intRoll = intRoll - 1
        ElseIf Mission.Zone(BASE_ZONE).Weather = BAD_WEATHER Then
            intRoll = intRoll - 2
        ElseIf Mission.Zone(BASE_ZONE).Weather = STORM_WEATHER Then
            intRoll = intRoll - 3
        End If
        
        If Bomber.Airman(PILOT).Status >= SW_STATUS _
        And Bomber.Airman(COPILOT).Status >= SW_STATUS Then
            ' Note f: Pilot and copilot both unable to operate the controls.
            ' Another airman must land the plane.
            intRoll = intRoll - 11
        ElseIf (Bomber.Airman(PILOT).Status <= LW1_STATUS And Bomber.Airman(PILOT).Mission >= 11) _
        Or (Bomber.Airman(COPILOT).Status <= LW1_STATUS And Bomber.Airman(COPILOT).Mission >= 11) Then
            ' Note b: The controls are manned by the pilot and/or copilot, at
            ' least one whom is a veteran who has not been lightly wounded
            ' more than once.
            intRoll = intRoll + 1
        End If
        
        intEnginesOut = CountEnginesOut()
        
        If intEnginesOut = 3 Then
            ' Note h.
            intRoll = intRoll - 3
        ElseIf intEnginesOut = 4 Then
            ' Note i.
            intRoll = intRoll - 7
        End If
        
        If Damage.Window >= 2 Then
            ' P-2 Pilot Compartment.
            intRoll = intRoll - 1
        End If
        
        If Bomber.BomberModel = B24_D _
        Or Bomber.BomberModel = B24_E _
        Or Bomber.BomberModel = B24_GHJ _
        Or Bomber.BomberModel = B24_LM Then
            
            If Damage.Autopilot = True Then
                ' B24s were more difficult and exhausting to fly.
                intRoll = intRoll - 2
            End If
    
        End If
        
        If Damage.ControlCables >= 2 Then
            ' P-2 Pilot Compartment.
            intRoll = intRoll - 1
        End If
        
        If Damage.ElevatorControls = True _
        Or (Damage.Elevator(PORT_SIDE) = True _
        And Damage.Elevator(STBD_SIDE) = True) Then
            ' P-6 Tail Section: Note b.
            intRoll = intRoll - 1
        End If
        
        ' P-6 Tail Section.
        
        If Bomber.BomberModel = B17_C _
        Or Bomber.BomberModel = B17_E _
        Or Bomber.BomberModel = B17_F _
        Or Bomber.BomberModel = B17_G _
        Or Bomber.BomberModel = YB40 Then
                
            ' A B-17 only has one rudder, so by default it is the 'port side'.
            
            If Damage.RudderControls = True _
            Or Damage.Rudder(PORT_SIDE) >= 3 Then
                intRoll = intRoll - 1
            End If
                
        ElseIf Bomber.BomberModel = B24_D _
        Or Bomber.BomberModel = B24_E _
        Or Bomber.BomberModel = B24_GHJ _
        Or Bomber.BomberModel = B24_LM Then
                
            If Damage.RudderControls = True _
            Or (Damage.Rudder(PORT_SIDE) >= 2 _
            And Damage.Rudder(STBD_SIDE) >= 2) Then
                intRoll = intRoll - 2
            End If

        Else ' Lancaster
        
            If Damage.RudderControls = True _
            Or (Damage.Rudder(PORT_SIDE) >= 2 _
            And Damage.Rudder(STBD_SIDE) >= 2) Then
                intRoll = intRoll - 1
            End If

        End If
        
        If Damage.WingFlapControls = True _
        Or (Damage.WingFlap(PORT_SIDE) = True _
        And Damage.WingFlap(STBD_SIDE) = True) Then
            ' BL-1 Wings: Note b.
            intRoll = intRoll - 1
        End If
        
        If Damage.AileronControls = True _
        Or (Damage.Aileron(PORT_SIDE) = True _
        And Damage.Aileron(STBD_SIDE) = True) Then
            ' BL-1 Wings: Note b.
            intRoll = intRoll - 1
        End If
        
        If Damage.Brake = True Then
            ' BL-1 Wings: Note h.
            intRoll = intRoll - 1
        End If

        If Damage.LandingGear = True Then
            
            If Bomber.BomberModel = B24_D _
            Or Bomber.BomberModel = B24_E _
            Or Bomber.BomberModel = B24_GHJ _
            Or Bomber.BomberModel = B24_LM Then
                ' Flimsy roll up doors tended to collapse when doing belly
                ' landings.
                intRoll = intRoll - 4
            Else
                ' BL-1 Wings: Note i.
                intRoll = intRoll - 3
            End If
        
        End If
        
        If Damage.Tailwheel = True Then
            ' P-6 Tail Section. B-17s and Lancasters only.
            intRoll = intRoll - 1
        ElseIf Damage.NoseWheel = True Then
            ' B-24s are only models which can have nosewheel damage, which makes
            ' it more likely the bomber will flip over.
            intRoll = intRoll - 2
        End If
        
        If Damage.BurstInPlane = True Then
            ' Rule 19.2.d.
            intRoll = intRoll - 4
        End If
        
        If Bomber.CurrentZone <> BASE_ZONE Then
            ' Note j: Landing in enemy territory, or emergency landing
            ' at forward tactical airfield.
            intRoll = intRoll - 3
        End If
        
        If intRoll <= 0 _
        And Random1D6() = 6 Then
            
            If Bomber.BombsOnBoard = True Then
                ' Note e.
                UpdateMessage "Bombs still aboard!"
                blnPayloadExploded = True
            ElseIf Bomber.ExtraFuelInBombBay = True Then
                blnPayloadExploded = True
            ElseIf Bomber.ExtraAmmo > 0 Then
                UpdateMessage "Extra ammo still aboard!"
                blnPayloadExploded = True
            End If
            
        End If
        
        If blnPayloadExploded = True Then
            
            If Bomber.RabbitsFoot >= 1 Then
                ' Expend luck to prevent loss of aircraft.
                UpdateMessage "Luckily, payload did not explode"
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                Bomber.Status = DUTY_STATUS
                UpdateMessage "Bomber made a safe landing."
            Else
                    
                Bomber.Status = CRASHED_STATUS
                UpdateMessage "Explosion. Bomber destroyed."
                Call CrewFinish(KIA_STATUS)
            
            End If
        
        Else
            
            If Damage.BurstInPlane = True _
            Or Damage.PeckhamPoints >= PECKHAM_SCRAP_LEVEL Then
                
                ' Rule 19.2.d: If the bomber suffered a BIP or other heavy
                ' damamge, then later managed to make a safe landing, the
                ' bomber is still permanently out of commission.
                
                If intRoll >= 1 Then
                    ' Roll 0 is "irrepairably damaged", but crew suffers no
                    ' additional ill effects.
                    intRoll = 0
                End If
            
            End If
        
            Select Case intRoll
                    
                Case Is <= -3:
                        
                    Bomber.Status = CRASHED_STATUS
                    UpdateMessage "Bomber wrecked."
                    Call CrewFinish(KIA_STATUS)
                
                Case -2:
                        
                    Bomber.Status = CRASHED_STATUS
                    UpdateMessage "Bomber wrecked."
                    Call CrewFinish(BAD_CRASH_STATUS)
                    
                    If EnemyTerritory() = True Then
                        UpdateMessage "Wreckage captured by enemy."
                        Call CrewCaptured
                    End If
                
                Case -1:
                        
                    Bomber.Status = CRASHED_STATUS
                    UpdateMessage "Bomber wrecked."
                    Call CrewFinish(CRASHED_STATUS)
                    
                    If EnemyTerritory() = True Then
                        UpdateMessage "Wreckage captured by enemy."
                        Call CrewCaptured
                    End If
                
                Case 0:
                        
                    If EnemyTerritory() = True Then
                        Bomber.Status = CRASHED_STATUS
                        Call CrewCaptured
                        UpdateMessage "Wreckage captured by enemy."
                    ElseIf Damage.BurstInPlane = True _
                    Or Damage.PeckhamPoints >= PECKHAM_SCRAP_LEVEL Then
                        Bomber.Status = SCRAPPED_STATUS
                        UpdateMessage "Bomber scrapped due to excessive damage."
                    Else
                        Bomber.Status = CRASHED_STATUS
                        UpdateMessage "Bomber crashed; irrepairably damaged."
                    End If
                    
                Case 1:
                        
                    If EnemyTerritory() = True Then
                        Bomber.Status = CAPTURED_STATUS
                        Call CrewCaptured
                        UpdateMessage "Bomber captured by enemy."
                    Else
                        Bomber.Status = DUTY_STATUS
                        UpdateMessage "Bomber crashed; repairable by next mission."
                    End If
                    
                Case Is >= 2:
                        
                    If EnemyTerritory() = True Then
                        Bomber.Status = CAPTURED_STATUS
                        Call CrewCaptured
                        UpdateMessage "Bomber captured by enemy."
                    Else
                        Bomber.Status = DUTY_STATUS
                        UpdateMessage "Bomber made a safe landing."
                    End If
                    
            End Select
        
        End If
    
    End If

End Sub

Public Function G9TakeOff() As Boolean
    Dim blnPayloadExploded As Boolean
    Dim intRoll As Integer
    'From "B-24 variant" (The Boardgamer, volume 8 no. 4?)
    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
    
        ' Basically, this is the G9LandingOnLand procedure, but only with
        ' modifiers for weight, weather and pilot/copilot skill.
    
        intRoll = Random2D6()
        
        ' Adjust the roll for the extra weight carried by the B-24.
        
        intRoll = intRoll - 1
        
        'O-1 Weather Note a. Bad weather causes -2 on G9/G10
        'O-1 Weather: Note b. Poor weather causes -1 on G9/G10
        
        
        If Mission.Zone(BASE_ZONE).Weather = POOR_WEATHER Then
            intRoll = intRoll - 1
        ElseIf Mission.Zone(BASE_ZONE).Weather = BAD_WEATHER Then
            intRoll = intRoll - 2
        ElseIf Mission.Zone(BASE_ZONE).Weather = STORM_WEATHER Then
            intRoll = intRoll - 3
        End If
            
        ' Adjust for pilot and copilot experience.
            
        If Bomber.Airman(PILOT).Mission >= 11 _
        And Bomber.Airman(COPILOT).Mission >= 11 Then
            intRoll = intRoll + 1
        End If
        
        If intRoll <= 1 Then
            UpdateMessage "Bomber fails to clear end of runway."
        End If
        If intRoll <= 0 Then
            If Bomber.BombsOnBoard _
                Or Bomber.ExtraFuelInBombBay _
                Or Bomber.ExtraAmmo > 0 _
                Then
                If Random1D6 = 6 Then
                    blnPayloadExploded = True
                End If
            End If
        End If
        If blnPayloadExploded Then
            Bomber.Status = CRASHED_STATUS
            UpdateMessage "Explosion. Bomber destroyed, killing all crew."
            Call CrewFinish(KIA_STATUS)
            G9TakeOff = False
        Else
            Select Case intRoll
                Case Is <= -3:
                    Bomber.Status = CRASHED_STATUS
                    UpdateMessage "Bomber wrecked."
                    Call CrewFinish(KIA_STATUS)
                
                Case -2:
                    Bomber.Status = CRASHED_STATUS
                    UpdateMessage "Bomber wrecked."
                    Call CrewFinish(BAD_CRASH_STATUS)
                    
                Case -1:
                    Bomber.Status = CRASHED_STATUS
                    UpdateMessage "Bomber wrecked."
                    Call CrewFinish(CRASHED_STATUS)
                    
                Case 0:
                    Bomber.Status = CRASHED_STATUS
                    UpdateMessage "Bomber crashed; irrepairably damaged."
                    
                Case 1:
                    Bomber.Status = DUTY_STATUS
                    UpdateMessage "Bomber crashed; repairable by next mission."
                    
                Case Is >= 2:
                    G9TakeOff = True
            
            End Select
            
        End If
    
    Else
        ' B-17s and Lancasters always successfully takeoff.
        G9TakeOff = True
    End If
End Function

'******************************************************************************
' G10LandingInWater
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  All water is considered "friendly territory", the airmen only have
'         to survive the elements.
'******************************************************************************
Public Sub G10LandingInWater()
    Dim intRoll As Integer
    Dim blnPayloadExploded As Boolean
    Dim intEnginesOut As Integer
    
    UpdateMessage "Landing in water:"
        
    intRoll = Random2D6()
    
    If intRoll = 12 Then
        
        ' Note a: Miracle roll.
        
        UpdateMessage "Bomber is lost, but the crew is rescued."
        ' Airmen's states are unchanged.
        Bomber.Status = DITCHED_STATUS
    
    Else
        
        ' O-1 Weather: Note a and Note b, plus "The General" (Volume 24, #6)
        ' variant.
        
        If Mission.Zone(BASE_ZONE).Weather = POOR_WEATHER Then
            intRoll = intRoll - 1
        ElseIf Mission.Zone(BASE_ZONE).Weather = BAD_WEATHER Then
            intRoll = intRoll - 2
        ElseIf Mission.Zone(BASE_ZONE).Weather = STORM_WEATHER Then
            intRoll = intRoll - 3
        End If
        
        If Bomber.Airman(PILOT).Status >= SW_STATUS _
        And Bomber.Airman(COPILOT).Status >= SW_STATUS Then
            ' Note f: Pilot and copilot both unable to operate the controls.
            ' Another airman must land the plane.
            intRoll = intRoll - 11
        ElseIf (Bomber.Airman(PILOT).Status <= LW1_STATUS And Bomber.Airman(PILOT).Mission >= 11) _
        Or (Bomber.Airman(COPILOT).Status <= LW1_STATUS And Bomber.Airman(COPILOT).Mission >= 11) Then
            ' Note b: The controls are manned by the pilot and/or copilot, at
            ' least one whom is a veteran who has not been lightly wounded
            ' more than once.
            intRoll = intRoll + 1
        End If
        
        intEnginesOut = CountEnginesOut()
        
        If intEnginesOut = 3 Then
            ' Note h.
            intRoll = intRoll - 3
        ElseIf intEnginesOut = 4 Then
            ' Note i.
            intRoll = intRoll - 7
        End If
        
        If Damage.Window >= 2 Then
            ' P-2 Pilot Compartment.
            intRoll = intRoll - 1
        End If
        
        If Bomber.BomberModel = B24_D _
        Or Bomber.BomberModel = B24_E _
        Or Bomber.BomberModel = B24_GHJ _
        Or Bomber.BomberModel = B24_LM Then
            
            If Damage.Autopilot = True Then
                ' B24s were more difficult and exhausting to fly.
                intRoll = intRoll - 2
            End If
    
        End If
        
        If Damage.ControlCables >= 2 Then
            ' P-2 Pilot Compartment.
            intRoll = intRoll - 1
        End If
        
        If Damage.ElevatorControls = True _
        Or (Damage.Elevator(PORT_SIDE) = True _
        And Damage.Elevator(STBD_SIDE) = True) Then
            ' P-6 Tail Section: Note b.
            intRoll = intRoll - 1
        End If
        
        ' P-6 Tail Section.
        
        If Bomber.BomberModel = B17_C _
        Or Bomber.BomberModel = B17_E _
        Or Bomber.BomberModel = B17_F _
        Or Bomber.BomberModel = B17_G _
        Or Bomber.BomberModel = YB40 Then
                
            ' A B-17 only has one rudder, so by default it is the 'port side'.
            
            If Damage.RudderControls = True _
            Or Damage.Rudder(PORT_SIDE) >= 3 Then
                intRoll = intRoll - 1
            End If
                
        ElseIf Bomber.BomberModel = B24_D _
        Or Bomber.BomberModel = B24_E _
        Or Bomber.BomberModel = B24_GHJ _
        Or Bomber.BomberModel = B24_LM Then
                
            If Damage.RudderControls = True _
            Or (Damage.Rudder(PORT_SIDE) >= 2 _
            And Damage.Rudder(STBD_SIDE) >= 2) Then
                intRoll = intRoll - 2
            End If

        Else ' Lancaster
        
            If Damage.RudderControls = True _
            Or (Damage.Rudder(PORT_SIDE) >= 2 _
            And Damage.Rudder(STBD_SIDE) >= 2) Then
                intRoll = intRoll - 1
            End If

        End If
        
        If Damage.WingFlapControls = True _
        Or (Damage.WingFlap(PORT_SIDE) = True _
        And Damage.WingFlap(STBD_SIDE) = True) Then
            ' BL-1 Wings: Note b.
            intRoll = intRoll - 1
        End If
        
        If Damage.AileronControls = True _
        Or (Damage.Aileron(PORT_SIDE) = True _
        And Damage.Aileron(STBD_SIDE) = True) Then
            ' BL-1 Wings: Note b.
            intRoll = intRoll - 1
        End If
        
        ' There are no negative modifiers for brake, landing gear, nosewheel or
        ' tailwheel damage when landing in water, because they are pointless
        ' without a runway.
        
        If Damage.BurstInPlane = True Then
            ' Rule 19.2.d.
            intRoll = intRoll - 4
        End If
        
        If Damage.Radio = True _
        And Bomber.InFormation = False Then
            ' Note g.
            intRoll = intRoll - 6
        End If
                
        If Bomber.BomberModel = B24_D _
        Or Bomber.BomberModel = B24_E _
        Or Bomber.BomberModel = B24_GHJ _
        Or Bomber.BomberModel = B24_LM Then
            ' Less flotation due to roll up doors and high wings.
            intRoll = intRoll - 2
        End If
        
        If intRoll <= 0 _
        And Random1D6() = 6 Then
            
            If Bomber.BombsOnBoard = True Then
                ' Note e.
                UpdateMessage "Bombs still aboard!"
                blnPayloadExploded = True
            ElseIf Bomber.ExtraFuelInBombBay = True Then
                blnPayloadExploded = True
            ElseIf Bomber.ExtraAmmo > 0 Then
                UpdateMessage "Extra ammo still aboard!"
                blnPayloadExploded = True
            End If
            
        End If
        
        If blnPayloadExploded = True Then
            
            Bomber.Status = DITCHED_STATUS
            UpdateMessage "Explosion. Bomber destroyed."
            Call CrewFinish(KIA_STATUS)
        
        Else
            
            Select Case intRoll
                        
                Case Is <= 3:
                            
                    Bomber.Status = DITCHED_STATUS
                    UpdateMessage "Crew lost at sea."
                    ' Even though some of the crew may have survived, no
                    ' one will ever know ...
                    Call CrewFinish(MIA_STATUS)
                    
                Case Is >= 4:
                            
                    Bomber.Status = DITCHED_STATUS
                    ' Airmen's states are unchanged.
                    UpdateMessage "Crew rescued."
                        
            End Select
        
        End If
    
    End If

End Sub

'******************************************************************************
' CountEnginesOut
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Number of inoperable engines.
'
' NOTES:  n/a
'******************************************************************************
Public Function CountEnginesOut() As Integer
    Dim intEngine As Integer
    
    CountEnginesOut = 0
    
    For intEngine = 1 To 4
        If Damage.EngineOut(intEngine) = True Then
            CountEnginesOut = CountEnginesOut + 1
        End If
    Next intEngine

End Function

'******************************************************************************
' OverWater
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if over water, otherwise false.
'
' NOTES:  Water territory (W) is obviously always water, but if the bomber
'         is over half water territory (W-N), it may be over water or land at
'         the current moment.
'******************************************************************************
Public Function OverWater() As Boolean
    Dim intRoll As Integer
    Dim strTer As String
    
    OverWater = False
    
    strTer = Mission.Zone(Bomber.CurrentZone).Terrain
    
    If strTer = WATER_TER Then
        
        ' All water zone.
        OverWater = True
    
    ElseIf Left(strTer, 1) = WATER_TER _
    And Mid(strTer, 1, 1) = "-" Then
        
        ' Half water, half land, zone.
        
        intRoll = Random1D6()
    
        If intRoll <= 3 Then
            OverWater = True
        End If

    End If
    
End Function

'******************************************************************************
' OverWater
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if over water, otherwise false.
'
' NOTES:  Water territory (W) is obviously always water, but if the bomber
'         is over half water territory (W-N), it may be over water or land at
'         the current moment.
'******************************************************************************
Public Function OverWater2(ByVal intZone As Integer) As Boolean
    Dim intRoll As Integer
    Dim strTer As String
    
    OverWater2 = False
    
    strTer = Mission.Zone(intZone).Terrain
    
    If strTer = WATER_TER Then
        
        ' All water zone.
        OverWater2 = True
    
    ElseIf Left(strTer, 1) = WATER_TER _
    And Mid(strTer, 1, 1) = "-" Then
        
        ' Half water, half land, zone.
        
        intRoll = Random1D6()
    
        If intRoll <= 3 Then
            OverWater2 = True
        End If

    End If
    
End Function

'******************************************************************************
' EnemyTerritory
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if over enemy territory, otherwise false.
'
' NOTES:  If the comber is over land (B), the bomber's current location may be
'         be over friendly or enemy held land, depending on when the country
'         was liberated.
'******************************************************************************
Public Function EnemyTerritory() As Boolean
    Dim intLiberationPct As Integer
    Dim strTer As String
    Dim intDate As Integer
    Dim intRoll As Integer
    
    EnemyTerritory = False
    
' ??? is this comment relevant ???
' If the country is England, it is always in Allied hands; if the
' country is Norway, it is always in Axis hands.
    
    strTer = Mission.Zone(Bomber.CurrentZone).Terrain
    intDate = Mission.Date
    
    intLiberationPct = GetLiberationPct(strTer, intDate)

    intRoll = RandomD100()

    If intRoll > intLiberationPct Then
        EnemyTerritory = True
    End If

End Function

'******************************************************************************
' CrewCaptured
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Crew was captured. Check each wounded airman to see if he dies in
'         captivity due to wounds. Any airman which does not die of wounds,
'         becomes a POW. (Unless rescued by partisans.) Disregard frostbite
'         and invalids.
'******************************************************************************
Private Sub CrewCaptured()
    Dim intPos As Integer
    Dim intIndex As Integer

    UpdateMessage "Crew Captured:"
    
    ' Cycle through the crew's originally assigned positions.
    For intPos = PILOT To AMMO_STOCKER
        
        If PosOccupied(intPos) = True Then
            
            ' Airman currently in position
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
            
            If BallGunnerCrushed(intPos, intIndex) = True Then
                GoTo Continue
            End If
            
            If Mission.Zone(Bomber.CurrentZone).Terrain = ALPS_TER Then
    
                ' The Alps do not contain partisans, so if the airman does not
                ' freeze fall thru. The RescuedByPartians() calls will return
                ' false, resulting in the airman's capture.
    
                If FreezeInAlps() = True Then
                    Bomber.Airman(intIndex).Status = KIA_STATUS
                    UpdateMessage Bomber.Airman(intIndex).Name & " freezes to death."
                End If
            
            End If
            
            Select Case Bomber.Airman(intIndex).Status
                Case DUTY_STATUS:

                    If RescuedByPartians() = True Then
                        Bomber.Airman(intIndex).Status = DUTY_STATUS
                        UpdateMessage Bomber.Airman(intIndex).Name & " rescued by partisans."
                    Else
                        Bomber.Airman(intIndex).Status = POW_STATUS
                        UpdateMessage Bomber.Airman(intIndex).Name & ": POW"
                    End If
                
                Case LW1_STATUS To LW2_STATUS:
                    
                    If RescuedByPartians() = True Then
                        Bomber.Airman(intIndex).Status = DUTY_STATUS
                        UpdateMessage Bomber.Airman(intIndex).Name & " rescued by partisans."
                    Else
                        Bomber.Airman(intIndex).Status = POW_STATUS
                        UpdateMessage Bomber.Airman(intIndex).Name & " recovers from his wounds: POW"
                    End If
                        
                Case SW_STATUS:

                    Select Case Random1D6()
                        Case 1 To 5
                            If RescuedByPartians() = True Then
                                Bomber.Airman(intIndex).Status = DUTY_STATUS
                                UpdateMessage Bomber.Airman(intIndex).Name & " rescued by partisans."
                            Else
                                Bomber.Airman(intIndex).Status = POW_STATUS
                                UpdateMessage Bomber.Airman(intIndex).Name & " recovers from his wounds: POW"
                            End If
                        Case 6
                            Bomber.Airman(intIndex).Status = DOW_STATUS
                            UpdateMessage Bomber.Airman(intIndex).Name & " dies from his wounds."
                    End Select
                
            End Select
                    
        End If
        
Continue:
        
    Next intPos

End Sub

'******************************************************************************
' FreezeInAlps
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if airman froze to death, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Private Function FreezeInAlps() As Boolean
    
    FreezeInAlps = False
    
    If Random1D6() <= 5 Then
        FreezeInAlps = True
    End If

End Function

'******************************************************************************
' AlpsDirection
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Indicate the location of the Alps relative to the bomber's direction
'         of travel.
'
' NOTES:  If the bomber does not cross the Alps on this mission, that
'         is also indicated.
'******************************************************************************
Public Function AlpsDirection() As Integer
    
    AlpsDirection = ALPS_NOWHERE
    
    If Mission.AlpsZone = ALPS_NOWHERE Then
        
        AlpsDirection = ALPS_NOWHERE
    
    ElseIf Bomber.CurrentZone = Mission.AlpsZone Then
        
        AlpsDirection = ALPS_BELOW
    
    ElseIf Bomber.Direction = OUTBOUND _
    And Bomber.CurrentZone = (Mission.AlpsZone - 1) _
    Or (Bomber.Direction = RETURN_TRIP _
    And Bomber.CurrentZone = (Mission.AlpsZone + 1)) Then
            
        AlpsDirection = ALPS_NEXT_ZONE
        
    ElseIf (Bomber.Direction = OUTBOUND _
    And Bomber.CurrentZone < Mission.AlpsZone) _
    Or (Bomber.Direction = RETURN_TRIP _
    And Bomber.CurrentZone > Mission.AlpsZone) Then
            
        AlpsDirection = ALPS_AHEAD
        
    Else
            
        AlpsDirection = ALPS_BEHIND
        
    End If
    
End Function

'******************************************************************************
' AlpsZone
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The zone where the mission crosses the Alps, otherwise nowhere.
'
' NOTES:  n/a
'******************************************************************************
Public Function AlpsZone() As Integer
    Dim intZone As Integer
    
    AlpsZone = ALPS_NOWHERE
    
    For intZone = BASE_ZONE To MAX_ZONE
        If Mission.Zone(intZone).Terrain = ALPS_TER Then
            AlpsZone = intZone
            Exit For
        End If
    Next intZone
    
End Function

'******************************************************************************
' RescuedByPartians
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if rescued by partisans, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Private Function RescuedByPartians() As Boolean
    Dim strTer As String
    Dim arrTer() As String
    
    RescuedByPartians = False
    
    strTer = Mission.Zone(Bomber.CurrentZone).Terrain
            
    If InStr(strTer, "-") > 0 Then
        ' The terrain is evenly divided between two countries, or a
        ' country and a body of water. Pick which side the crew is on.
        arrTer = Split(strTer, "-")
                
        If Random1D6() <= 3 Then
            strTer = arrTer(0)
        Else
            strTer = arrTer(1)
        End If
            
    End If
    
    If strTer = FRANCE_TER _
    Or strTer = BELGIUM_TER _
    Or strTer = ITALY_TER _
    Or strTer = GREECE_TER Then
        If Random1D6() = 6 Then
            RescuedByPartians = True
        End If
    ElseIf strTer = YUGOSLAVIA_TER Then
        ' Yugoslav partisans didn't just engage in guerilla tactics, they
        ' actually controlled large chunks of the country, so airmen are
        ' more likely to be rescued there.
        If Random1D6() >= 5 Then
            RescuedByPartians = True
        End If
    End If
            
End Function

'******************************************************************************
' CrewFinish
'
' INPUT:  The final status of all surviving members of the crew.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Public Sub CrewFinish(ByVal intLevel As Integer)
    Dim intPos As Integer
    Dim intIndex As Integer

    UpdateMessage "Crew Finish:"
    
    ' Cycle through the crew's originally assigned positions.
    For intPos = PILOT To AMMO_STOCKER
        
        If PosOccupied(intPos) = True Then
            
            ' Airman currently in position
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
            
            Select Case intLevel
                
                Case KIA_STATUS:
                    
                    Bomber.Airman(intIndex).Status = KIA_STATUS
                    Bomber.Airman(intIndex).Wounded = True
                    UpdateMessage Bomber.Airman(intIndex).Name & ": KIA"
            
                Case MIA_STATUS:
                    
                    Bomber.Airman(intIndex).Status = MIA_STATUS
                    UpdateMessage Bomber.Airman(intIndex).Name & ": MIA"
            
                Case BAD_CRASH_STATUS:
                    
                    If BallGunnerCrushed(intPos, intIndex) = False Then
                        UpdateMessage Bomber.Airman(intIndex).Name & ": " & BL4Wound(intPos, BAD_CRASH_STATUS)
                    End If
            
                Case CRASHED_STATUS:
                    
                    If BallGunnerCrushed(intPos, intIndex) = False Then
                        UpdateMessage Bomber.Airman(intIndex).Name & ": " & BL4Wound(intPos, CRASHED_STATUS)
                    End If
            
            End Select
        
        End If
        
    Next intPos

End Sub

'******************************************************************************
' BallGunnerCrushed
'
' INPUT:  The position being evaluated and the airman occupying the position.
'
' OUTPUT: n/a
'
' RETURN: True if the ball gunner was crushed, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function BallGunnerCrushed(ByVal intPos As Integer, ByVal intIndex As Integer) As Boolean

    BallGunnerCrushed = False
    
    If intPos <> BALL_GUNNER _
    Or Bomber.Airman(intIndex).Status = KIA_STATUS Then
        ' Either this is not the ball turret position, or the airman is already
        ' dead.
        Exit Function
    End If

    If Damage.BallTurretMech = True _
    And Damage.LandingGear = True Then
        
        ' P-5 Waist: Note c. Ball gunner is crushed in turret.

        If Bomber.Airman(intIndex).Status <= SW_STATUS _
        And Bomber.RabbitsFoot >= 1 Then

            ' Expend luck to prevent horrible fate.
            UpdateMessage "Ball turret miraculously retracts"
            Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1

        Else

            BallGunnerCrushed = True
            Bomber.Airman(intIndex).Status = KIA_STATUS
            Bomber.Airman(intIndex).Wounded = True
            UpdateMessage Bomber.Airman(intIndex).Name & " crushed in turret: KIA"

        End If

    End If
            
End Function

'******************************************************************************
' BL4Wound
'
' INPUT:  The airman's position index, and any wounding penalty.
'
' OUTPUT: n/a
'
' RETURN: Wound message.
'
' NOTES:  n/a
'******************************************************************************
Private Function BL4Wound(ByVal intPos As Integer, Optional ByVal intSeverity As Integer) As String
' TODO: This function should not be called if the position is hidden. Note
' that a position can be hidden, but a weapon not, or vice versa. ???
    Dim intPrevState As Integer
    Dim intRoll As Integer
    Dim intWound As Integer
    Dim intNewState As Integer
    Dim intIndex As Integer
    
    Dim strMessage As String

    BL4Wound = ""

    ' Airman currently in position
    intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
            
    ' Get the current status of the airman.
    
    intPrevState = Bomber.Airman(intIndex).Status
    
    If intPrevState = KIA_STATUS Then
        
        strMessage = "already KIA"
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    
    Else
        
        intRoll = Random1D6()

        If IsMissing(intSeverity) = False Then
            If intSeverity = BAD_CRASH_STATUS Then
                intRoll = intRoll + 1
            End If
        End If
        
        Select Case intRoll
            Case 1 To 3:
                intWound = LW1_STATUS
                Damage.PeckhamPoints = Damage.PeckhamPoints + 2
            Case 4 To 5:
                intWound = SW_STATUS
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Case Is >= 6:
                intWound = KIA_STATUS
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        End Select
        
        intNewState = intPrevState + intWound
        
        If intNewState >= SW_STATUS _
        And Bomber.Airman(intIndex).Kills >= 5 _
        And Bomber.RabbitsFoot >= 1 Then
            
            ' Expend luck to prevent loss of ace gunner.
            UpdateMessage "Ace gunner luckily escaped (further) injury"
            Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
            intNewState = intPrevState
            
        End If
        
        If intNewState >= LW1_STATUS Then
            Bomber.Airman(intIndex).Wounded = True
        End If
        
        If intNewState > KIA_STATUS Then
            intNewState = KIA_STATUS
        End If
        
        Select Case intNewState
            
            Case DUTY_STATUS:
                
                strMessage = "No wound"
            
            Case LW1_STATUS:
                
                strMessage = "LW1"
                
            Case LW2_STATUS:
                
                strMessage = "LW2"
            
            Case SW_STATUS:
                
                strMessage = "SW"
            
            Case KIA_STATUS:
                
                strMessage = "KIA"
        
        End Select
        
'        ' Position and guns become unmanned when an airman is dead or seriously
'        ' wounded, even though the position is not empty. When another airman
'        ' takes over the position, he swaps positions with the wounded/dead
'        ' airman: So, even though non-hidden positions can never be empty while
'        ' the bomber is in flight, the guns associated with the position may
'        ' become unmanned.
'
'         If intNewState >= SW_STATUS Then
'
'            ' Regardless of bomber model, or position, if the airman is
'            ' manning a weapon, the weapon becomes unmanned.
'
'            Call UnmanGun(intIndex)
'
'         End If
        
        ' Assign the new state to the airman.
        
        Bomber.Airman(intIndex).Status = intNewState
    
    End If
    
    BL4Wound = strMessage
    
End Function

'******************************************************************************
' BL5Frostbite
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  Cycle through all positions. If the heater is broken assess if the
'         airman currently occupying the position sustains frostbite.
'******************************************************************************
Public Sub BL5Frostbite()
    Dim intPos As Integer
    Dim intIndex As Integer

    ' Cycle through the existing positions on the bomber.
    For intPos = PILOT To AMMO_STOCKER
        
        If PosOccupied(intPos) = True Then
            
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
            
            If Bomber.Airman(intIndex).Status <= SW_STATUS _
            And Bomber.Airman(intIndex).Frostbite = False _
            And Damage.Heater(intPos) = True Then
                
                ' The airman currently manning the position is alive, and not
                ' yet frostbitten, but the position has no heat. Determine if
                ' the airman gets frostbite.
                
                If Random1D6() <= 3 Then
                    Bomber.Airman(intIndex).Frostbite = True
                    UpdateMessage Bomber.Airman(intIndex).Name & ": Frostbite"
                End If
                
            End If
            
        End If
    
    Next intPos

    ' No Peckham Points for frostbite.
    
End Sub

'******************************************************************************
' EndMission
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  TODO
'******************************************************************************
Public Sub EndMission()
    Dim intRoll As Integer
    Dim strMessage As String
    Dim strIgnore As String
    Dim strErrMsg As String
    
    Dim strWound As String
    Dim intPos As Integer
    Dim intIndex As Integer
    
    ' By the time we reach this point, the bomber's and crews' status has
    ' been set. What happens to the crew after the mission -- awards, death
    ' from wounds, etc. -- now has to be determined.
    
    If LookupBomber(Bomber.KeyField, LOOKUP_BY_KEYFIELD, strIgnore) = False Then
        strErrMsg = "Could not find bomber #" & Bomber.KeyField & _
                    " in the database. The mission will be aborted."
        MsgBox strErrMsg, (vbCritical + vbOKOnly)
        Exit Sub
    End If
                    
    If prsBomber![Default] = True Then
        ' If the bomber is default, then so are the airmen, squadron and group
        ' group. Do not perform any updates.
        GoTo Continue
    End If
                
    prsBomber![Status] = Bomber.Status
    
    ' This is a user created bomber. Therefore the airmen, squadron, group and
    ' bomber must all be updated. Point at the bomber's squadron.
                
    If LookupSquadron(prsBomber![Squadron], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
        strErrMsg = "Could not find squadron #" & prsBomber![Squadron] & _
                    " in the database. The mission will be aborted."
        MsgBox strErrMsg, (vbCritical + vbOKOnly)
        Exit Sub
    End If
                
' TODO: Make third param in LookUpXXX() functions optional.

    ' Point at the squadron's group.

    If LookupGroup(prsSquadron![Group], LOOKUP_BY_KEYFIELD, strIgnore) = False Then
        strErrMsg = "Could not find group #" & prsSquadron![Group] & _
                    " in the database. The mission will be aborted."
        MsgBox strErrMsg, (vbCritical + vbOKOnly)
        Exit Sub
    End If
                
    ' Sequentially point at each of the airmen assigned to the bomber by
    ' cycling through the crew's originally assigned positions.
    
    For intPos = PILOT To AMMO_STOCKER
        
        If PosOccupied(intPos) = True Then
            
            ' Airman currently in position
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
            
            ' Point at the airman.
                
            If LookupAirman(Bomber.Airman(intIndex).SerialNumber, LOOKUP_BY_KEYFIELD, strIgnore) = False Then
                strErrMsg = "Airman " & Bomber.Airman(intIndex).Name & _
                            "(#" & Bomber.Airman(intIndex).SerialNumber & ") " & _
                            "could not be found for post-mission update. He " & _
                            "will remain in his pre-mission state. " & vbCrLf & vbCrLf & _
                            "Continue with other post-mission updates."
    
                MsgBox strErrMsg, (vbExclamation + vbOKOnly)
                
                GoTo Continue
            End If
prsAirman![Status] = Bomber.Airman(intIndex).Status ' Nov04
' TODO: Only default entities -- airmen, bombers, squadrons and groups --
' may be associated with each other.
        
'            ' If the airman is not a default airman, update his personnel
'            ' file. Non-Default airmen can only be part of non-default
'            ' bombers, squadrons and groups, so those entities may be
'            ' updated as well.
'
'            If prsAirman![Default] = True Then
'                GoTo Continue
'            End If

            ' First the airman recovers from any wounds.
            
            If Bomber.Airman(intIndex).Wounded = True Then
            
                ' Determine if the airman recovers from his wounds. If the
                ' airman recovers, increment the unit wound counters; if he
                ' dies, increment the unit KIA counters. POWs and airmen
                ' rescued by partisans were already checked for recovery;
                ' their status will be duty, POW or DOW, so they they are not
                ' counted in unit wound/KIA(???) totals.
            
                Select Case Bomber.Airman(intIndex).Status
                    Case LW1_STATUS To LW2_STATUS:
    
'                        Bomber.Airman(intIndex).Status = DUTY_STATUS
                        prsAirman![Status] = DUTY_STATUS
                        prsSquadron![Wounded] = prsSquadron![Wounded] + 1
                        prsGroup![Wounded] = prsGroup![Wounded] + 1
                        strMessage = Bomber.Airman(intIndex).Name & " recovers from his wounds."
                    
                    Case SW_STATUS:
    
                        intRoll = Random1D6()
            
                        If intRoll >= 4 _
                        And Bomber.Airman(intIndex).Kills >= 5 _
                        And Bomber.RabbitsFoot >= 1 Then
        
                            ' Expend luck to prevent loss of ace gunner.
                            UpdateMessage "Ace gunner luckily escapes sure death"
                            Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                            intRoll = 1
                        
                        End If
                        
                        Select Case intRoll
                            Case 1 To 3
                                
                                prsAirman![Status] = DUTY_STATUS
                                prsSquadron![Wounded] = prsSquadron![Wounded] + 1
                                prsGroup![Wounded] = prsGroup![Wounded] + 1
                                strMessage = Bomber.Airman(intIndex).Name & " recovers from his wounds."
                            
                            Case 4 To 5
                                
                                prsAirman![Status] = INVALID_STATUS
                                If prsBomber![Status] = DUTY_STATUS Then prsBomber![Status] = STAND_DOWN_STATUS
                                prsSquadron![Wounded] = prsSquadron![Wounded] + 1
                                prsGroup![Wounded] = prsGroup![Wounded] + 1
                                strMessage = Bomber.Airman(intIndex).Name & " is invalided home due to wounds. His war is over."
                            
                            Case 6
                                
                                prsAirman![Status] = DOW_STATUS
                                If prsBomber![Status] = DUTY_STATUS Then prsBomber![Status] = STAND_DOWN_STATUS
                                prsSquadron![KIA] = prsSquadron![KIA] + 1
                                prsGroup![KIA] = prsGroup![KIA] + 1
                                strMessage = Bomber.Airman(intIndex).Name & " dies from his wounds."
                        
                        End Select
                        
                    Case KIA_STATUS:
                        
'                        If Bomber.Airman(intIndex).Kills >= 5 _
'                        And Bomber.RabbitsFoot >= 1 Then
'
'                            ' Expend luck to prevent loss of ace gunner.
'                            UpdateMessage "Ace gunner luckily escapes sure death"
'                            Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
'
'                            prsSquadron![Wounded] = prsSquadron![Wounded] + 1
'                            prsGroup![Wounded] = prsGroup![Wounded] + 1
'                            strMessage = Bomber.Airman(intIndex).Name & " recovers from his wounds."
'
'                        Else
                        
                            prsAirman![Status] = KIA_STATUS
                            If prsBomber![Status] = DUTY_STATUS Then prsBomber![Status] = STAND_DOWN_STATUS
                            prsSquadron![KIA] = prsSquadron![KIA] + 1
                            prsGroup![KIA] = prsGroup![KIA] + 1
                            strMessage = Bomber.Airman(intIndex).Name & " died during mission."
                    
'                        End If
                        
                End Select
    
                UpdateMessage strMessage
            
            End If
    
            ' If airman was wounded (or dead), or has frostbite, he gets a
            ' Purple Heart.
            
            If Bomber.Airman(intIndex).Wounded = True _
            Or Bomber.Airman(intIndex).Frostbite = True Then
                prsAirman![PurpleHeart] = prsAirman![PurpleHeart] + 1
                prsSquadron![PurpleHeart] = prsSquadron![PurpleHeart] + 1
                prsGroup![PurpleHeart] = prsGroup![PurpleHeart] + 1
            End If
            
            ' Frostbite is tracked separately from wounds. Even if the airman
            ' recovered from a serious wound, he may be invalided due to
            ' frostbite.
                
            If prsAirman![Status] = DUTY_STATUS _
            And Bomber.Airman(intIndex).Frostbite = True Then
                
                intRoll = Random1D6()
    
                If intRoll <= 3 _
                And Bomber.Airman(intIndex).Kills >= 5 _
                And Bomber.RabbitsFoot >= 1 Then

                    ' Expend luck to prevent loss of ace gunner.
                    UpdateMessage "Ace gunner luckily recovers from frostbite"
                    Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                    intRoll = 6
                
                End If
                
                Select Case intRoll
                    Case 1 To 3:
                        prsAirman![Status] = INVALID_STATUS
                        If prsBomber![Status] = DUTY_STATUS Then prsBomber![Status] = STAND_DOWN_STATUS
                        strMessage = Bomber.Airman(intIndex).Name & " is invalided home due to frostbite. His war is over."
                    Case 4 To 6:
                        strMessage = Bomber.Airman(intIndex).Name & " recovers from frostbite."
                End Select
                
                UpdateMessage strMessage
                
            End If

' TODO: The number of required missions varied by nationality and year.
' If Bomber.Airman(intIndex).Status = DUTY_STATUS ' Nov04
            If prsAirman![Status] = DUTY_STATUS _
            And Bomber.Airman(intIndex).Mission >= MAX_MISSIONS Then
                ' If the airman made it through this mission, and he completed
                ' the required number of missions, then his tour is complete.
                prsAirman![Status] = TOUR_COMPLETE_STATUS
                If prsBomber![Status] = DUTY_STATUS Then prsBomber![Status] = STAND_DOWN_STATUS
                
                strMessage = Bomber.Airman(intIndex).Name & " finished " & _
                             Bomber.Airman(intIndex).Mission & " missions. " & _
                             "His tour is complete!"
                
                UpdateMessage strMessage
            
            End If
Dim a
a = 1
' If Bomber.Airman(intIndex).Status <> DUTY_STATUS ' Nov04
            If prsAirman![Status] <> DUTY_STATUS _
            And (prsBomber![Status] = DUTY_STATUS _
            Or prsBomber![Status] = STAND_DOWN_STATUS) Then
                
                ' Airman was captured, went home, or did not recover from
                ' wounds. His last assignment is retained in his record.
                ' Since the bomber returned to base, and is fit for further
                ' duty, the airman should be removed from its roster.

                Select Case intIndex ' intPos
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
' ElseIf Bomber.Airman(intIndex).Status = DUTY_STATUS ' Nov04

ElseIf prsAirman![Status] = DUTY_STATUS _
And prsBomber![Status] <> DUTY_STATUS _
And prsBomber![Status] <> STAND_DOWN_STATUS Then
prsAirman![Assignment] = ADMIN_DUTY
            End If
            
            ' Record the airman's end of mission information.
            
            prsAirman![LeadCrewExp] = Bomber.Airman(intIndex).LeadCrewExp
            prsAirman![Sorties] = Bomber.Airman(intIndex).Mission
            
            ' Increment the bomber, squadron and group kills by the difference
            ' between the airman's post-mission and pre-mission kills. For
            ' example: 102 = 100 + ( 3 - 1 ) ... or ... 100 = 100 + ( 5 - 5 )
            
            prsBomber![Kills] = prsBomber![Kills] + (Bomber.Airman(intIndex).Kills - prsAirman![Kills])
            prsSquadron![Kills] = prsSquadron![Kills] + (Bomber.Airman(intIndex).Kills - prsAirman![Kills])
            prsGroup![Kills] = prsGroup![Kills] + (Bomber.Airman(intIndex).Kills - prsAirman![Kills])

            ' We no longer need this airman's pre-mission kills, so save his
            ' post-mission kills.
            
            prsAirman![Kills] = Bomber.Airman(intIndex).Kills
            
            ' Record remaining airman-oriented unit information.
            
            If prsAirman![Status] = POW_STATUS Then
                prsSquadron![POW] = prsSquadron![POW] + 1
                prsGroup![POW] = prsGroup![POW] + 1
            ElseIf prsAirman![Status] = MIA_STATUS Then
                prsSquadron![MIA] = prsSquadron![MIA] + 1
                prsGroup![MIA] = prsGroup![MIA] + 1
            End If
    
' TODO: Other awards. (Option 1) Popup so the player can make awards.
'                     (Option 2) Determine automatically.
                
        End If ' not a hidden position
    
    Next intPos

'    ' Now that the crew has been updated, update the bomber. As with airmen,
'    ' non-Default bombers can only be part of non-default squadrons and
'    ' groups, so those entities may be updated as well.
'
'    If prsBomber![Default] = False Then
        
        If Bomber.Status = CRASHED_STATUS _
        Or Bomber.Status = CAPTURED_STATUS _
        Or Bomber.Status = DITCHED_STATUS _
        Or Bomber.Status = SHOT_DOWN_STATUS _
        Or Bomber.Status = SCRAPPED_STATUS Then
            
'            prsBomber![Status] = Bomber.Status
            prsSquadron![PlanesLost] = prsSquadron![PlanesLost] + 1
            prsGroup![PlanesLost] = prsGroup![PlanesLost] + 1
        
            ' Even though the plane was lost, retain its last assigned crew.
            ' Surviving individual airmen will be available for assignment
            ' to other bombers.
        
        End If
            
        prsBomber![Sorties] = prsBomber![Sorties] + 1
        prsSquadron![Sorties] = prsSquadron![Sorties] + 1
        prsGroup![Sorties] = prsGroup![Sorties] + 1
    
'    End If

    ' Save all mission end changes at the same time.
    
    pintOpenTrans = pintOpenTrans + pobjConn.BeginTrans()
            
        prsAirman.UpdateBatch
        prsBomber.UpdateBatch
        prsSquadron.UpdateBatch
        prsGroup.UpdateBatch
        
    pobjConn.CommitTrans
        
    pintOpenTrans = pintOpenTrans - 1
    
Continue:
    
    ' Reset bookmarks so the pointed record is in synch with the record
    ' on the tabs.
    
' TODO: Are the tabbed records refreshed when the above update is performed?
    
    prsAirman.Bookmark = varAirmanCurrentlyOnTab
    prsBomber.Bookmark = varBomberCurrentlyOnTab
    prsSquadron.Bookmark = varSquadronCurrentlyOnTab
    prsGroup.Bookmark = varGroupCurrentlyOnTab
    
'MsgBox "frmMission.Hide"
'
'    Unload frmHelpBrowser
'    frmMainMenu.Show

End Sub

'******************************************************************************
' O1Weather
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The weather level.
'
' NOTES:
'
'Mission.Zone(Bomber.CurrentZone).Weather
'Clear: Apply +1 to B-1, B-2, M-4, O-2, O-6
'Good: No changes
'Poor: Apply -1 to B-1, B-2, M-4, O-2x, O-6x, G-9x, G-10x
'Bad: Apply -2 to B-1, B-2, M-4, O-2x, O-6x, G-9x, G-10x
'Storm: Apply -3 to B-1, B-2, G-9, G-10. No flak: Bomber must
'       abort or attack alternate target.
'
'******************************************************************************
Public Function O1Weather() As Integer
        
    O1Weather = GOOD_WEATHER
    
    If Not Mission.Options.AlternateWeather Then
        
        ' Use the traditional weather table.
        Select Case Random2D6()
            Case 2, 12:
                O1Weather = BAD_WEATHER
            Case 3, 11:
                O1Weather = POOR_WEATHER
            Case 4 To 10:
                O1Weather = GOOD_WEATHER
        End Select
    
    Else
        
        ' Use the alternate weather table. Modified from the "Theater
        ' Modifications" article in "The General" (Volume 24, #6).
        
        Select Case Random2D6()
            Case 2, 12:
                O1Weather = STORM_WEATHER
            Case 3, 11:
                O1Weather = BAD_WEATHER
            Case 4, 10:
                O1Weather = POOR_WEATHER
            Case 5, 6, 8, 9:
                O1Weather = GOOD_WEATHER ' Contrails may form, Zone 2 or later
            Case 7:
                O1Weather = CLEAR_WEATHER ' Contrails may form, Zone 2 or later
        End Select
    
    End If
    
End Function

'******************************************************************************
' O2FlakOverTarget
'
' INPUT:  Boolean indicating if the numeric or string flak level is being
'         returned.
'
' OUTPUT: n/a
'
' RETURN: Flak density as a string or number.
'
' NOTES:  n/a
'******************************************************************************
Public Function O2FlakOverTarget(ByVal blnReturnString As Boolean) As Variant
' TODO: If there is stormy weather in the target zone, this function should
' not be called.
    
    ' If there is a storm in the target zone, there is no flak and the bomber
    ' may not bomb the target. The bomber must abort or attack an alternate
    ' target.
    
    Dim intRoll As Integer
    Dim vntFlak As Variant
   
    intRoll = Random1D6()
   
    ' O-1 Weather: Note a and Note b, plus "The General" (Volume 24, #6)
    ' variant.
        
    If Mission.Zone(Bomber.CurrentZone).Weather = CLEAR_WEATHER Then
        intRoll = intRoll + 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = POOR_WEATHER Then
        intRoll = intRoll - 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = BAD_WEATHER Then
        intRoll = intRoll - 2
    End If
    
    If Mission.Zone(Bomber.CurrentZone).Contrail = True Then
        intRoll = intRoll + 1
    End If
    
    If Bomber.Altitude = LOW_ALTITUDE Then
        UpdateMessage "Can't evade flak at low level."
        intRoll = intRoll + 2
    ElseIf Mission.Options.EvadeFlak = True Then
        intRoll = intRoll - 2
        If blnReturnString = False Then UpdateMessage "Evading flak."
    End If
    
    ' Lancaster variant.
    
    If Bomber.SpottedBySearchLight = True Then
        intRoll = intRoll + 1
    End If
    
    Select Case intRoll
        Case Is <= 1:
            
            If blnReturnString = False Then
                vntFlak = NO_FLAK
                UpdateMessage "No flak."
            Else
                vntFlak = "No flak"
            End If
        
        Case 2 To 3:
            
            If blnReturnString = False Then
                vntFlak = LIGHT_FLAK
                UpdateMessage "Light flak."
            Else
                vntFlak = "Light flak"
            End If
        
        Case 4 To 5:
            
            If blnReturnString = False Then
                vntFlak = MEDIUM_FLAK
                UpdateMessage "Medium flak."
            Else
                vntFlak = "Medium flak"
            End If
        
        Case Is >= 6:
            
            If blnReturnString = False Then
                vntFlak = HEAVY_FLAK
                UpdateMessage "Heavy flak."
            Else
                vntFlak = "Heavy flak"
            End If
    
    End Select
   
    O2FlakOverTarget = vntFlak
   
End Function

'******************************************************************************
' O3FlakToHitBomber
'
' INPUT:  Flak density.
'
' OUTPUT: n/a
'
' RETURN: Number of bursts that hit the bomber.
'
' NOTES:  n/a
'******************************************************************************
Public Function O3FlakToHitBomber(ByVal intDensity As Integer, ByVal blnTargetZoneCombat As Boolean) As Integer
    Dim intRoll As Integer
    Dim intIndex As Integer
    Dim intBurstMax As Integer
    Dim intBursts As Integer
            
    O3FlakToHitBomber = 0
    
    If blnTargetZoneCombat = True Then
        intBurstMax = 3
    Else
        ' Wandering around the countryside ...
        intBurstMax = 2
    End If
    
    intBursts = 0
   
    For intIndex = 1 To intBurstMax
        
        intRoll = Random2D6()

        If intDensity = LIGHT_FLAK Then
            
            If intRoll = 2 _
            Or intRoll = 12 Then
                intBursts = intBursts + 1
            End If
        
        ElseIf intDensity = MEDIUM_FLAK Then
            
            If intRoll = 2 _
            Or intRoll = 3 _
            Or intRoll = 7 _
            Or intRoll = 12 Then
                intBursts = intBursts + 1
            End If
        
        ElseIf intDensity = HEAVY_FLAK Then
            
            If intRoll = 2 _
            Or intRoll = 3 _
            Or intRoll = 5 _
            Or intRoll = 7 _
            Or intRoll = 9 _
            Or intRoll = 11 _
            Or intRoll = 12 Then
                intBursts = intBursts + 1
            End If
        
        End If
    
    Next intIndex

    UpdateMessage "Bomber hit by " & intBursts & " flak bursts."

    O3FlakToHitBomber = intBursts

End Function

'******************************************************************************
' O4EffectOfFlakHits
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The number of flak hits caused by a given burst.
'
' NOTES:  This function should be called once per burst.
'******************************************************************************
Public Function O4EffectOfFlakHits() As Integer
    Dim intRoll As Integer
    Dim intHits As Integer
    
    intHits = 0
    
    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2: intHits = BURST_IN_PLANE
        Case 3: intHits = 1
        Case 4: intHits = 4
        Case 5: intHits = 3
        Case 6: intHits = 2
        Case 7: intHits = 1
        Case 8: intHits = 2
        Case 9: intHits = 3
        Case 10: intHits = 4
        Case 11: intHits = 1
        Case 12: intHits = 4
    End Select

    O4EffectOfFlakHits = intHits

End Function

'******************************************************************************
' O5AreaAffectedByFlakHit
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the flak caused catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Public Function O5AreaAffectedByFlakHit() As Integer
    Dim intRoll As Integer
    
    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2: O5AreaAffectedByFlakHit = P3BombBayDamage()
        Case 3: O5AreaAffectedByFlakHit = BL1WingDamage(PORT_SIDE)
        Case 4:
            
            ' B-24 and Lancaster radio operators were located on the flight
            ' deck, not in a separate compartment.
    
            If Bomber.BomberModel = B24_D _
            Or Bomber.BomberModel = B24_E _
            Or Bomber.BomberModel = B24_GHJ _
            Or Bomber.BomberModel = B24_LM _
            Or Bomber.BomberModel = AVRO_LANCASTER Then
                UpdateMessage "Flight Deck: Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                O5AreaAffectedByFlakHit = P4RadioRoomDamage()
            End If
        
        Case 5: O5AreaAffectedByFlakHit = P6TailDamage()
        Case 6: O5AreaAffectedByFlakHit = BL1WingDamage(PORT_SIDE)
        Case 7: O5AreaAffectedByFlakHit = P6TailDamage()
        Case 8: O5AreaAffectedByFlakHit = BL1WingDamage(STBD_SIDE)
        Case 9: O5AreaAffectedByFlakHit = P5WaistDamage()
        Case 10: O5AreaAffectedByFlakHit = P1NoseDamage()
        Case 11: O5AreaAffectedByFlakHit = BL1WingDamage(STBD_SIDE)
        Case 12: O5AreaAffectedByFlakHit = P2FlightDeckDamage()
    End Select
    
End Function

'******************************************************************************
' LeadCrewExp
'
' INPUT:  The position to be evaluated.
'
' OUTPUT: n/a
'
' RETURN: True if the airman at the position has lead crew experience against
'         the target, otherwise false.
'
' NOTES:  "The General" (Volume 24, #6) variant.
'******************************************************************************
Private Function LeadCrewExp(ByVal intPos As Integer) As Boolean
    Dim strLeadCrewExp As String
    Dim arrLeadCrewExp() As String
    Dim intCount As Integer
    Dim intLen As Integer
    Dim intPipes As Integer
    Dim intMissions As Integer
    
    LeadCrewExp = False

    strLeadCrewExp = Bomber.Airman(intPos).LeadCrewExp

    intLen = Len(strLeadCrewExp)

'MsgBox "intLen = '" & intLen & "'"

    If intLen >= 1 Then
        
        For intCount = 1 To intLen
            If Mid(strLeadCrewExp, intCount, 1) = "|" Then
                intPipes = intPipes + 1
            End If
        Next intCount
    
'        ' The last element did not end with a pipe
'        intPipes = intPipes + 1
'
'MsgBox "intPipes = '" & intPipes & "'"

        arrLeadCrewExp = Split(strLeadCrewExp, "|")

        For intCount = 0 To intPipes
'MsgBox intCount & " = '" & arrLeadCrewExp(intCount) & "'"
            
            If CInt(arrLeadCrewExp(intCount)) = prsBomberTarget![KeyField] Then
                
                intMissions = intMissions + 1
'MsgBox "intMissions = '" & intMissions & "'"
                
                If intMissions >= 2 Then
                    LeadCrewExp = True
                    Exit For
                End If
            
            End If
        
        Next intCount

    End If

'MsgBox "Lead Crew Exp = '" & LeadCrewExp & "'"

End Function

'******************************************************************************
' O6BombRun
'
' INPUT:  The number of flak hits the bomber sustained.
'
' OUTPUT: n/a
'
' RETURN: Whether or not the bombs were on target.
'
' NOTES:  n/a
'******************************************************************************
Public Function O6BombRun(ByVal intFlakHits As Integer) As Boolean
    Dim intRoll As Integer
    
    O6BombRun = False

    If Damage.BombSight = True _
    Or Damage.ControlCables >= 2 _
    Or Bomber.Airman(BOMBARDIER).Status >= SW_STATUS _
    Or (Damage.Autopilot = True And Damage.IntercomSystem = True) Then
        
        ' P-1 Nose.
        ' P-2 Pilot Compartment.
        ' O-6 Bomb Run: Note c.
        ' BL-2 Instruments: Note a.
        
        O6BombRun = False
        Exit Function
    
    End If
    
    If (Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM) _
    And Damage.Autopilot = True Then
        
        ' B24s were more difficult to fly.
        O6BombRun = False
        Exit Function
    
    End If
    
    intRoll = Random1D6()
    
    ' O-1 Weather: Note a and Note b, plus "The General" (Volume 24, #6)
    ' variant.
        
    If Mission.Zone(Bomber.CurrentZone).Weather = CLEAR_WEATHER Then
        intRoll = intRoll + 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = POOR_WEATHER Then
        intRoll = intRoll - 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = BAD_WEATHER Then
        intRoll = intRoll - 2
    End If
    
    ' "The General" (Volume 24, #6) variant.
    
    If Mission.Options.CrewExperience = True _
    And LeadCrewExp(BOMBARDIER) = True _
    And LeadCrewExp(NAVIGATOR) = True Then
    
        intRoll = intRoll + 1
    
    End If
    
    If Mission.Options.TimePeriodSpecificFormations = True _
    And Mission.Date = AUG_1942 Then
        intRoll = intRoll - 1
    End If
    
    If Mission.Options.EvadeFlak = True Then
        intRoll = intRoll - 3
    End If

    ' Note a.
    
    If Bomber.Airman(BOMBARDIER).Mission >= 11 _
    And Bomber.Airman(BOMBARDIER).Status <= LW1_STATUS Then
        intRoll = intRoll + 1
    End If
    
    ' Note b.
    
    If intFlakHits > 0 Then
        intRoll = intRoll - 1
    End If
    
    If Bomber.Airman(BOMBARDIER).Status = LW2_STATUS Then
        intRoll = intRoll - 1
    End If
    
    ' P-1 Nose: Note c.
    ' P-3 BombBay.
    
    If Damage.BombControls = True _
    Or Damage.BombRelease = True Then
        ' Bomb must be manually dropped.
        intRoll = intRoll - 3
    End If
    
    ' BL-2 Instruments.
    
    If Damage.Autopilot = True Then
        intRoll = intRoll - 2
    End If
    
    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
        ' Less drag due to roll up doors.
        intRoll = intRoll + 1
    End If
    
    If intRoll <= 2 Then
        O6BombRun = False
    Else
        O6BombRun = True
    End If
        
End Function

'******************************************************************************
' O7BombingAccuracy
'
' INPUT:  Whether or not the bombs were on target.
'
' OUTPUT: n/a
'
' RETURN: The percentage of bombs that hit the target.
'
' NOTES:  n/a
'******************************************************************************
Public Function O7BombingAccuracy(ByVal blnOnTarget As Boolean) As Integer
    Dim intRoll As Integer
    Dim intAccuracy As Integer
    
    intRoll = Random2D6()
    
    If blnOnTarget = True Then
        
        Select Case intRoll
            Case 2: intAccuracy = 75
            Case 3: intAccuracy = 60
            Case 4, 6, 8, 10: intAccuracy = 30
            Case 5, 9: intAccuracy = 20
            Case 7: intAccuracy = 40
            Case 11: intAccuracy = 50
            Case 12: intAccuracy = 88 + Random2D6()
        End Select
    
    Else ' off target
        
        Select Case intRoll
            Case 2, 12: intAccuracy = 10
            Case 3, 11: intAccuracy = 5
            Case 4 To 10: intAccuracy = 0
        End Select
    
    End If
    
    O7BombingAccuracy = intAccuracy
    
End Function

'******************************************************************************
' HeaterHit
'
' INPUT:  Index to the position whose heater was hit.
'
' OUTPUT: A damage string.
'
' RETURN: True if the heater is out, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Private Function HeaterHit(ByVal intPos As Integer, ByRef strEffect As String) As Boolean

    Dim strCrewPosition As String

    HeaterHit = True

    If Damage.Heater(intPos) = True Then
        HeaterHit = False
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        If LookupCrewPosition(intPos, strCrewPosition) = False Then
            HeaterHit = False
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Exit Function
        End If

        Damage.Heater(intPos) = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    
        If strEffect <> "" Then
            ' Some other airman's heater was affected by the same shot.
            strEffect = strEffect & " \ "
        End If
    
        strEffect = strEffect & strCrewPosition & " heat out"
    
    End If
    
End Function

'******************************************************************************
' OxygenHit
'
' INPUT:  Index to the position whose oxygen was hit.
'
' OUTPUT: A damage string.
'
' RETURN: True if the oxygen is out, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Private Function OxygenHit(ByVal intPos As Integer, ByVal intHits As Integer, ByRef strEffect As String) As Boolean
    ' There are only two oxygen bottles available at each position. As long
    ' as an airman is alive, he must be placed at a position where air is
    ' available.
' TODO: Transfer oxygen bottles between positions

    Dim strCrewPosition As String

    OxygenHit = True

    If Damage.Oxygen(intPos) = 2 Then
        OxygenHit = False
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        If LookupCrewPosition(intPos, strCrewPosition) = False Then
            OxygenHit = False
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Exit Function
        End If

        If Damage.Oxygen(intPos) = 0 Then
            If intHits = 1 Then
                Damage.Oxygen(intPos) = Damage.Oxygen(intPos) + 1
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            ElseIf intHits = 2 Then
                Damage.Oxygen(intPos) = Damage.Oxygen(intPos) + 2
                Damage.PeckhamPoints = Damage.PeckhamPoints + 15
            End If
        ElseIf Damage.Oxygen(intPos) = 1 Then
            Damage.Oxygen(intPos) = Damage.Oxygen(intPos) + 1
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        End If
        
        If strEffect <> "" Then
            ' Some previous airman's oxygen was affected by the same shot.
            strEffect = strEffect & " \ "
        End If
    
        If Damage.Oxygen(intPos) = 1 Then
            strEffect = strEffect & strCrewPosition & " oxygen hit"
        ElseIf Damage.Oxygen(intPos) = 2 Then
            strEffect = strEffect & strCrewPosition & " oxygen out"
        End If
    End If
    
End Function

'******************************************************************************
' OxygenFire
'
' INPUT:  1-4 indices to the positions whose oxygen was consumed by the fire.
'
' OUTPUT: A damage string.
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub OxygenFire(ByRef strEffect As String, ByVal intPos1 As Integer, Optional ByVal intPos2 As Integer = UNMANNED_POSITION, Optional ByVal intPos3 As Integer = UNMANNED_POSITION, Optional ByVal intPos4 As Integer = UNMANNED_POSITION)

    Dim strCrewPosition As String

    Call OxygenHit(intPos1, 2, strEffect)
    
    If PosExists(intPos2) = True Then
        Call OxygenHit(intPos2, 2, strEffect)
    End If
    
    If PosExists(intPos3) = True Then
        Call OxygenHit(intPos3, 2, strEffect)
    End If
    
    If PosExists(intPos4) = True Then
        Call OxygenHit(intPos4, 2, strEffect)
    End If
    
    strEffect = strEffect & " FIRE!"
    
End Sub

'******************************************************************************
' BL3ExtinguishFire
'
' INPUT:  The number of (hand-held or engine) extinguishers remaining; is this
'         an engine fire?
'
' OUTPUT: The number of (hand-held or engine) extinguishers remaining.
'
' RETURN: True if the fire was extinguished, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Private Function BL3ExtinguishFire(ByRef intExtinguishers As Integer, ByVal blnEngFire As Boolean) As Boolean

    ' BL-3 modified to handle BL-1 engine fires as well.
    
    Dim intMaxAttempts As Integer
    Dim intCount As Integer
    
    BL3ExtinguishFire = False
    
    If intExtinguishers <= 3 Then
        intMaxAttempts = intExtinguishers
    Else
        intMaxAttempts = 3
    End If

    intCount = 0

    Do While intCount < intMaxAttempts _
    And BL3ExtinguishFire = False
        
        If Bomber.BomberModel = B17_C _
        And blnEngFire = True Then
            ' Variant: B-17Cs lacked self-sealing fuel tanks, so their engine
            ' fires are more difficult to extinguish.
            If Random1D6() <= 2 Then
                BL3ExtinguishFire = True
            End If
        Else
            If Random1D6() <= 4 Then
                BL3ExtinguishFire = True
            End If
        End If
    
        intCount = intCount + 1
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    
    Loop
        
    intExtinguishers = intExtinguishers - intCount
        
    If BL3ExtinguishFire = False _
    And Bomber.RabbitsFoot >= 1 Then
        ' Expend luck to prevent loss of aircraft.
        UpdateMessage "Fire luckily sputters out"
        Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
    End If

End Function

'******************************************************************************
' UnmanGun
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  TODO: Is this function still being used ???
'******************************************************************************
Private Sub UnmanGun(ByRef intIndex As Integer)

    Dim intGunPos As Integer

    ' Find the position that the airman currently occupies. If the position
    ' has a weapon, mark it as unmanned.

    For intGunPos = MID_UPPER_MG To TAIL_MG
        If Bomber.Gun(intGunPos).MannedBy = Bomber.Airman(intIndex).SerialNumber Then
            If GunExists(intGunPos) = True Then
                Bomber.Gun(intGunPos).MannedBy = UNMANNED_MG
                Exit Sub
            End If
        End If
    Next intGunPos

End Sub

'******************************************************************************
' GetAirmanIndexBySerialNumber
'
' INPUT:  An airman's serial number.
'
' OUTPUT: n/a
'
' RETURN: The index in the Bomber.Airman() array where the airman is located
'         (i.e., the airman's position).
'
' NOTES:  n/a
'******************************************************************************
Public Function GetAirmanIndexBySerialNumber(ByRef intSerialNumber As Integer) As Integer
    ' Here is an example of how this function can be used:
    '
    ' 1) Get serial number of airman currently occupying position
    ' 2) Get that airman's original position
    '    intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(NAVIGATOR).CurrentSerialNum)
    ' 3) Get that airman's name
    ' 4) pass that position to the wound function
    '    strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
    
    Dim intIndex As Integer

    GetAirmanIndexBySerialNumber = UNMANNED_POSITION

    ' Find the position that the airman was originally asigned to.

    For intIndex = PILOT To AMMO_STOCKER
        
        If Bomber.Airman(intIndex).SerialNumber = intSerialNumber Then
            GetAirmanIndexBySerialNumber = intIndex
            Exit For
        End If
    
    Next intIndex

End Function

'******************************************************************************
' P1NoseDamage
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P1NoseDamage() As Integer
    
    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.BomberModel = B17_F _
    Or Bomber.BomberModel = B17_G _
    Or Bomber.BomberModel = YB40 Then
    
        P1NoseDamage = P1NoseB17()
    
    ElseIf Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
    
        P1NoseDamage = P1NoseB24()
    
    Else
    
        P1NoseDamage = P1NoseLanc()
    
    End If

End Function
    
'******************************************************************************
' P1NoseB17
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P1NoseB17() As Integer
   
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            strArea = "Bomb sight"
            
            If Bomber.BomberModel = YB40 _
            Or Damage.BombSight = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "Bomb run automatically off target."
                Damage.BombSight = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 3:
            
            strArea = "Armament"
            
            Select Case Random1D6()
                Case 1 To 2:
                    
                    If Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    ElseIf Bomber.BomberModel = B17_G _
                    Or Bomber.BomberModel = YB40 Then
                        
' TODO
                        ' if enemy attacking from low then
                            strEffect = "Nose turret inoperable"
                            Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                        ' else
                        '   strArea = "Superficial damage"
                        '   Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                        ' end if
                    
                    Else
                        
                        strEffect = "Nose gun inoperable"
                        Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    
                    End If
            
                Case 3 To 4:
                    
                    If Bomber.BomberModel = B17_C _
                    Or Bomber.BomberModel = B17_E _
                    Or Bomber.Gun(PORT_CHEEK_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Port cheek gun inoperable"
                        Bomber.Gun(PORT_CHEEK_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If

                Case 5 To 6:
                    
                    If Bomber.BomberModel = B17_C _
                    Or Bomber.BomberModel = B17_E _
                    Or Bomber.Gun(STBD_CHEEK_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Starboard cheek gun inoperable"
                        Bomber.Gun(STBD_CHEEK_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
            
            End Select
        
        Case 4:
            
            strArea = "Crew"
            
            If Bomber.BomberModel = YB40 Then
                
                intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(NAVIGATOR).CurrentSerialNum)
                strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
            
            ElseIf Bomber.BomberModel = B17_C Then
                
                intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(BOMBARDIER).CurrentSerialNum)
                strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex) & " / "
            
            Else
                intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(BOMBARDIER).CurrentSerialNum)
                strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex) & " / "
                
                intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(NAVIGATOR).CurrentSerialNum)
                strEffect = strEffect & Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
            End If
            
        Case 5:
            
            If Bomber.BomberModel = B17_C Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strArea = "Crew"
                intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(NAVIGATOR).CurrentSerialNum)
                strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
            End If
        
        Case 6:
            
            strArea = "Crew"
            
            If Bomber.BomberModel = YB40 Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(BOMBARDIER).CurrentSerialNum)
                strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
            End If
        
        Case 7 To 9:
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 10:
            
            strArea = "Equipment"
            
            Select Case Random1D6()
                
                Case 1 To 3:
                    
' TODO: if Out of Formation, 50% chance of spending a second turn in the zone
                    If Damage.NavigationEquipment = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Navigation equipment inoperable"
                        Damage.NavigationEquipment = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                End If
                
                Case 4 To 6:
                    
                    If Bomber.BomberModel = YB40 _
                    Or Damage.BombControls = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Bomb controls inoperable"
                        Damage.BombControls = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
        
            End Select
        
        Case 11:
            
            strArea = "Heater"
            
            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    If Bomber.BomberModel = YB40 _
                    Or HeaterHit(BOMBARDIER, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 3 To 4:
                    
                    If Bomber.BomberModel = B17_C _
                    Or HeaterHit(NAVIGATOR, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 5 To 6:
                    
                    If Bomber.BomberModel = YB40 Then
                        If HeaterHit(NAVIGATOR, strEffect) = False Then
                            strArea = "Superficial damage"
                        End If
                    ElseIf Bomber.BomberModel = B17_C Then
                        If HeaterHit(BOMBARDIER, strEffect) = False Then
                            strArea = "Superficial damage"
                        End If
                    Else
                        If HeaterHit(BOMBARDIER, strEffect) = False _
                        And HeaterHit(NAVIGATOR, strEffect) = False Then
                            strArea = "Superficial damage"
                        End If
                    End If
                    
            End Select
        
        Case 12:
            
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    If Bomber.BomberModel = YB40 _
                    Or OxygenHit(BOMBARDIER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 3 To 4:
                    
                    If Bomber.BomberModel = B17_C _
                    Or OxygenHit(NAVIGATOR, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 5:
                    
                    If Bomber.BomberModel = YB40 Then
                        If OxygenHit(NAVIGATOR, 1, strEffect) = False Then
                            strArea = "Superficial damage"
                        End If
                    ElseIf Bomber.BomberModel = B17_C Then
                        If OxygenHit(BOMBARDIER, 1, strEffect) = False Then
                            strArea = "Superficial damage"
                        End If
                    Else
                        If OxygenHit(BOMBARDIER, 1, strEffect) = False _
                        And OxygenHit(NAVIGATOR, 1, strEffect) = False Then
                            strArea = "Superficial damage"
                        End If
                    End If
            
                Case 6:
                    
                    If Bomber.BomberModel = YB40 Then
                        Call OxygenFire(strEffect, NAVIGATOR)
                    ElseIf Bomber.BomberModel = B17_C Then
                        Call OxygenFire(strEffect, BOMBARDIER)
                    Else
                        Call OxygenFire(strEffect, NAVIGATOR, BOMBARDIER)
                    End If
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Nose: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P1NoseB17 = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select

    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Nose: " & strArea
    Else
        UpdateMessage "Nose: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' P1NoseB24
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P1NoseB24() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    strArea = "Bomb sight"
                    
                    If Damage.BombSight = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Bomb run automatically off target."
                        Damage.BombSight = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
                Case 4 To 6:
            
                    strArea = "Wheel"
                    
                    If Damage.NoseWheel = True Then
                        strArea = "Superficial damage"
                    Else
                        strEffect = "Nose wheel inoperable"
                        Damage.NoseWheel = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    End If
                
            End Select
        
        Case 3:
            
            strArea = "Armament"

            If Bomber.BomberModel = B24_D Then
                
                If Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE Then
                    strArea = "Superficial damage"
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                Else
                    strEffect = "Nose gun inoperable"
                    Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                End If

            ElseIf Bomber.BomberModel = B24_E Then
                        
                Select Case Random1D6()
                    Case 1 To 2:
                    
                        If Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE Then
                            strArea = "Superficial damage"
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                        Else
                            strEffect = "Nose gun inoperable"
                            Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                        End If
            
                    Case 3 To 4:

                        If Bomber.Gun(PORT_CHEEK_MG).Status = MG_INOPERABLE Then
                            strArea = "Superficial damage"
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                        Else
                            strEffect = "Port cheek gun inoperable"
                            Bomber.Gun(PORT_CHEEK_MG).Status = MG_INOPERABLE
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                        End If

                    Case 5 To 6:
            
                        If Bomber.Gun(STBD_CHEEK_MG).Status = MG_INOPERABLE Then
                            strArea = "Superficial damage"
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                        Else
                            strEffect = "Starboard cheek gun inoperable"
                            Bomber.Gun(STBD_CHEEK_MG).Status = MG_INOPERABLE
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                        End If
                
                End Select
            
            Else
            
                If Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE Then
                    strArea = "Superficial damage"
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                Else
' TODO
                    ' if enemy attacking from high then
                        strEffect = "Nose turret inoperable"
                        Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    ' else
                    '   strArea = "Superficial damage"
                    '   Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    ' end if
                End If
                    
            End If
                
        Case 4:
            
            strArea = "Crew"
            strEffect = BL4WoundB24(BOMBARDIER, strArea)
        
        Case 5:
            
            strArea = "Crew"
            strEffect = BL4WoundB24(NAVIGATOR, strArea)
        
        Case 6:
            
            strArea = "Crew"
            
            If Bomber.BomberModel = B24_D _
            Or Bomber.BomberModel = B24_E Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = BL4WoundB24(NOSE_GUNNER, strArea)
            End If
        
        Case 7 To 9:
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 10:
            
            strArea = "Equipment"
            
            Select Case Random1D6()
                
                Case 1 To 3:
                    
' TODO: if Out of Formation, 50% chance of spending a second turn in the zone
                    If Damage.NavigationEquipment = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Navigation equipment inoperable"
                        Damage.NavigationEquipment = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
                Case 4 To 6:
                    
                    If Damage.BombControls = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Bomb controls inoperable"
                        Damage.BombControls = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
        
            End Select
        
        Case 11:
            
            strArea = "Heater"
            
            Select Case Random1D6()
                Case 1 To 2:
                    
                    If HeaterHit(BOMBARDIER, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 3 To 4:
                    
                    If HeaterHit(NAVIGATOR, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 5 To 6:
                
                    If Bomber.BomberModel = B24_D _
                    Or Bomber.BomberModel = B24_E Then
                        
                        If HeaterHit(BOMBARDIER, strEffect) = False _
                        And HeaterHit(NAVIGATOR, strEffect) = False Then
                            strArea = "Superficial damage"
                        End If
                    
                    Else
                    
                        If HeaterHit(NOSE_GUNNER, strEffect) = False Then
                            strArea = "Superficial damage"
                        End If
                    
                    End If
        
            End Select
                    
        Case 12:
            
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    If OxygenHit(BOMBARDIER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 3 To 4:
                    
                    If OxygenHit(NAVIGATOR, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 5:
                    
                    If Bomber.BomberModel = B24_D _
                    Or Bomber.BomberModel = B24_E Then
                        strArea = "Superficial damage"
                    Else
                        If OxygenHit(NOSE_GUNNER, 1, strEffect) = False Then
                            strArea = "Superficial damage"
                        End If
                    End If
            
                Case 6:
                    
                    If Bomber.BomberModel = B24_D _
                    Or Bomber.BomberModel = B24_E Then
                        Call OxygenFire(strEffect, NAVIGATOR, BOMBARDIER)
                    Else
                        Call OxygenFire(strEffect, NAVIGATOR, BOMBARDIER, NOSE_GUNNER)
                    End If
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Nose: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P1NoseB24 = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select

    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Nose: " & strArea
    Else
        UpdateMessage "Nose: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' P1NoseLanc
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P1NoseLanc() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            strArea = "Bomb sight"
            
            If Damage.BombSight = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "Bomb run automatically off target."
                Damage.BombSight = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
                
        Case 3:
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    
        Case 4:
            
            Select Case Random1D6()
                Case 1 To 3:
                
                    strArea = "Armament"
                    
                    If Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Nose turret inoperable"
                        Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If
                
                Case 4 To 5:
                
                    strArea = "Crew"
                    intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(BOMBARDIER).CurrentSerialNum)
                    strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
            
                Case 6:
                    
                    strArea = "Crew / Armament"
                    
                    If Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE Then
                        strArea = "Crew"
                        intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(BOMBARDIER).CurrentSerialNum)
                        strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(BOMBARDIER).CurrentSerialNum)
                        strEffect = "Nose turret inoperable and " & _
                                    Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                        
                        Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If
            
            End Select
            
        Case 5:
            
            strArea = "Crew"
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(BOMBARDIER).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 6 To 8:
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 9:
            
            strArea = "Crew"
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(BOMBARDIER).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 10:
            
            strArea = "Bomb controls"
            
            If Damage.BombControls = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "inoperable"
                Damage.BombControls = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 11:
            
            strArea = "Heater"
            
            If HeaterHit(BOMBARDIER, strEffect) = False Then
                strArea = "Superficial damage"
            End If
        
        Case 12:
            
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1 To 5:
                    
                    If OxygenHit(BOMBARDIER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 6:
                    
                    Call OxygenFire(strEffect, BOMBARDIER)
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Nose: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P1NoseLanc = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select

    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Nose: " & strArea
    Else
        UpdateMessage "Nose: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' ControlCableHit
'
' INPUT:  n/a
'
' OUTPUT: A damage string.
'
' RETURN: True if both control cables are severed, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Private Function ControlCableHit(ByRef strEffect As String) As Boolean
    
    ControlCableHit = True

    If Damage.ControlCables >= 2 Then
        
        ControlCableHit = False
    
    Else

        Damage.ControlCables = Damage.ControlCables + 1
    
        If Damage.ControlCables = 1 Then
            strEffect = "Hit"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 25
        ElseIf Damage.ControlCables >= 2 Then
            strEffect = "Severed"
        End If

    End If
    
End Function

'******************************************************************************
' WindowHit
'
' INPUT:  n/a
'
' OUTPUT: A damage string.
'
' RETURN: True if both windows are out, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Private Function WindowHit(ByRef strEffect As String) As Boolean
    
    WindowHit = True

    If Damage.Window >= 2 Then
        
        WindowHit = False
    
    Else

        Damage.Window = Damage.Window + 1
    
        If Damage.Window = 1 Then
            strEffect = "Hit"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 25
        ElseIf Damage.Window >= 2 Then
            strEffect = "Shot out"
        End If

    End If
    
End Function

'******************************************************************************
' P2FlightDeckDamage
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P2FlightDeckDamage() As Integer
    
    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.BomberModel = B17_F _
    Or Bomber.BomberModel = B17_G _
    Or Bomber.BomberModel = YB40 Then
    
        P2FlightDeckDamage = P2FlightDeckB17()
    
    ElseIf Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
    
        P2FlightDeckDamage = P2FlightDeckB24()
    
    Else
    
        P2FlightDeckDamage = P2FlightDeckLanc()
    
    End If

End Function
    
'******************************************************************************
' P2FlightDeckB17
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P2FlightDeckB17() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            strArea = "Heater"
            
            If HeaterHit(PILOT, strEffect) = False _
            And HeaterHit(COPILOT, strEffect) = False Then
                strArea = "Superficial damage"
            End If
            
        Case 3
            
            strArea = "Crew"
            
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(PILOT).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex) & " / "
            
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(COPILOT).CurrentSerialNum)
            strEffect = strEffect & Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 4
            
            strArea = "Crew"
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(PILOT).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 5
            
            strArea = "Crew"
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(COPILOT).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 6 To 7
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 8
            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    strArea = "Armament"
                    
                    If Bomber.Gun(TOP_TURRET_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Top turret inoperable"
                        Bomber.Gun(TOP_TURRET_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If
                
                Case 3 To 5
                    
                    strArea = "Crew"
                    intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(ENGINEER).CurrentSerialNum)
                    strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                
                Case 6
                    
                    strArea = "Crew / Armament"
                    
                    If Bomber.Gun(TOP_TURRET_MG).Status = MG_INOPERABLE Then
                        strArea = "Crew"
                        intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(ENGINEER).CurrentSerialNum)
                        strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(ENGINEER).CurrentSerialNum)
                        strEffect = "Top turret inoperable and " & _
                                    Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                        Bomber.Gun(TOP_TURRET_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If

            End Select
        
        Case 9
                    
            P2FlightDeckB17 = BL2Instruments()
            Exit Function
        
        Case 10
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1:
                    
                    If OxygenHit(PILOT, 1, strEffect) = False _
                    And OxygenHit(COPILOT, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 2:
                    
                    If OxygenHit(PILOT, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 3:
                    
                    If OxygenHit(COPILOT, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 4 To 5:
                    
                    If OxygenHit(ENGINEER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 6:
                    
                    Call OxygenFire(strEffect, PILOT, COPILOT, ENGINEER)
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Flight Deck: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P2FlightDeckB17 = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select

        Case 11:
            
            strArea = "Window"

            If WindowHit(strEffect) = False Then
                strArea = "Superficial damage"
            End If
        
        Case 12:
            
            strArea = "Control Cables"
            
            If ControlCableHit(strEffect) = False Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Flight Deck: " & strArea
    Else
        UpdateMessage "Flight Deck: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' P2FlightDeckB24
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  B-24 radio operators were located on the flight deck, not in a
'         separate compartment.
'******************************************************************************
Private Function P2FlightDeckB24() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            strArea = "Heater"
            
            If HeaterHit(PILOT, strEffect) = False _
            And HeaterHit(COPILOT, strEffect) = False Then
                strArea = "Superficial damage"
            End If
            
        Case 3
            
            strArea = "Equipment"
            
            If Damage.IntercomSystem = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                Damage.IntercomSystem = True
                strEffect = "Intercom out"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 4
            
            strArea = "Crew"
            strEffect = BL4WoundB24(PILOT, strArea)
        
        Case 5
            
            strArea = "Equipment"
            
            If Damage.Radio = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                Damage.Radio = True
                strEffect = "Radio out"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 6

            strArea = "Crew"
            
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    strEffect = BL4WoundB24(COPILOT, strArea)
                
                Case 4 To 6:
                    
                    strEffect = BL4WoundB24(RADIO_OPERATOR, strArea)
            
            End Select
        
        Case 7
        
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 8
            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    strArea = "Armament"
                    
                    If Bomber.Gun(TOP_TURRET_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Top turret inoperable"
                        Bomber.Gun(TOP_TURRET_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If
                
                Case 3 To 5
                    
                    strArea = "Crew"
                    strEffect = BL4WoundB24(ENGINEER, strArea)
                
                Case 6
                    
                    strArea = "Crew / Armament"
                    
                    If Bomber.Gun(TOP_TURRET_MG).Status = MG_INOPERABLE Then
                        strArea = "Crew"
                        strEffect = BL4WoundB24(ENGINEER, strArea)
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Top turret inoperable and " & BL4WoundB24(ENGINEER, strArea)
                        Bomber.Gun(TOP_TURRET_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If

            End Select
        
        Case 9
            
            P2FlightDeckB24 = BL2Instruments()
            Exit Function
        
        Case 10
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1:
                    
                    If OxygenHit(PILOT, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 2:
                    
                    If OxygenHit(COPILOT, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 3:
                    
                    If OxygenHit(ENGINEER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 4:
                    
                    If OxygenHit(RADIO_OPERATOR, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 5:
                    
                    If OxygenHit(PILOT, 1, strEffect) = False _
                    And OxygenHit(COPILOT, 1, strEffect) = False _
                    And OxygenHit(ENGINEER, 1, strEffect) = False _
                    And OxygenHit(RADIO_OPERATOR, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 6:
                    
                    Call OxygenFire(strEffect, PILOT, COPILOT, ENGINEER, RADIO_OPERATOR)
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Flight Deck: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P2FlightDeckB24 = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If
            
            End Select

        Case 11:
            
            strArea = "Window"

            If WindowHit(strEffect) = False Then
                strArea = "Superficial damage"
            End If
        
        Case 12:
            
            strArea = "Control Cables"
            
            If ControlCableHit(strEffect) = False Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Flight Deck: " & strArea
    Else
        UpdateMessage "Flight Deck: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' P2FlightDeckLanc
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  Lancaster radio operators were located on the flight deck, not in a
'         separate compartment.
'******************************************************************************
Private Function P2FlightDeckLanc() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            Call P3BombBayDamage
            Exit Function
            
        Case 3
            
            strArea = "Heater"
            
            If HeaterHit(PILOT, strEffect) = False _
            And HeaterHit(NAVIGATOR, strEffect) = False _
            And HeaterHit(ENGINEER, strEffect) = False _
            And HeaterHit(RADIO_OPERATOR, strEffect) = False Then
                strArea = "Superficial damage"
            End If
            
        Case 4
            
            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    strArea = "Equipment"
                    
                    If Damage.NavigationEquipment = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        Damage.NavigationEquipment = True
                        strEffect = "Navigation equipment inoperable"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
                Case 3 To 5
                    
                    strArea = "Crew"
                    intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(NAVIGATOR).CurrentSerialNum)
                    strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                
                Case 6
                    
                    strArea = "Crew / Equipment"
                    
                    If Damage.NavigationEquipment = True Then
                        strArea = "Crew"
                        intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(NAVIGATOR).CurrentSerialNum)
                        strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(NAVIGATOR).CurrentSerialNum)
                        strEffect = "Navigation equipment inoperable and " & _
                                    Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                        Damage.NavigationEquipment = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If

            End Select
            
        Case 5
            
            strArea = "Crew"
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(PILOT).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 6

            strArea = "Crew"
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(ENGINEER).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 7
        
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 8
            
            Select Case Random1D6()
                
                Case 1:
                    
                    strArea = "Equipment"
                    
                    If Damage.IntercomSystem = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        Damage.IntercomSystem = True
                        strEffect = "Intercom out"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
                Case 2 To 3:
                    
                    strArea = "Equipment"
                    
                    If Damage.Radio = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        Damage.Radio = True
                        strEffect = "Radio out"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
                Case 4 To 6:
                    
                    strArea = "Crew"
                    intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(RADIO_OPERATOR).CurrentSerialNum)
                    strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)

            End Select
        
        Case 9
                    
            P2FlightDeckLanc = BL2Instruments()
            Exit Function
        
        Case 10
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1:
                    
                    If OxygenHit(PILOT, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 2:
                    
                    If OxygenHit(ENGINEER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 3:
                    
                    If OxygenHit(NAVIGATOR, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 4:
                    
                    If OxygenHit(RADIO_OPERATOR, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 5:
                    
                    If OxygenHit(PILOT, 1, strEffect) = False _
                    And OxygenHit(ENGINEER, 1, strEffect) = False _
                    And OxygenHit(NAVIGATOR, 1, strEffect) = False _
                    And OxygenHit(RADIO_OPERATOR, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 6:
                    
                    Call OxygenFire(strEffect, PILOT, ENGINEER, NAVIGATOR, RADIO_OPERATOR)
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Flight Deck: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P2FlightDeckLanc = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select

        Case 11:
            
            strArea = "Window"

            If WindowHit(strEffect) = False Then
                strArea = "Superficial damage"
            End If
        
        Case 12:
            
            strArea = "Control Cables"
            
            If ControlCableHit(strEffect) = False Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Flight Deck: " & strArea
    Else
        UpdateMessage "Flight Deck: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' BL2Instruments
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function BL2Instruments() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
    
            If Damage.Autopilot = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strArea = "Autopilot"
                strEffect = "Out"
                Damage.Autopilot = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 3:
            
            If Damage.LandingGear = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strArea = "Landing gear controls"
                strEffect = "Out"
                Damage.LandingGear = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 4:
        
            If Damage.IntercomSystem = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strArea = "Intercom"
                strEffect = "Out"
                Damage.IntercomSystem = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 5:
            
            If Damage.OxygenSystem = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strArea = "Oxygen system"
                strEffect = "Out"
                Damage.OxygenSystem = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 6:

            If Damage.WingFlapControls = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strArea = "Flap controls"
                strEffect = "Out"
                Damage.WingFlapControls = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 7:
            
            If Damage.AileronControls = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strArea = "Aileron controls"
                strEffect = "Out"
                Damage.AileronControls = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 8:
            
            If Damage.ElevatorControls = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strArea = "Elevator controls"
                strEffect = "Out"
                Damage.ElevatorControls = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
            
        Case 9:
            
            If Damage.RudderControls = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strArea = "Rudder controls"
                strEffect = "Out"
                Damage.RudderControls = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
            
        Case 10:
            
            If Damage.FeatheringCtrl = True Then
                strArea = "Superficial damage"
            Else
                strArea = "Propellor feathering controls"
                strEffect = "Inoperable"
                Damage.FeatheringCtrl = True
            End If
            
        Case 11:

            If Damage.EngineExtCtrl = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strArea = "Engine extinguisher controls"
                strEffect = "Inoperable"
                Damage.EngineExtCtrl = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
            
        Case 12:

            If Bomber.RabbitsFoot >= 1 Then
                ' Expend luck to prevent loss of aircraft.
                strArea = "Electrical system"
                strEffect = "Luckily throws sparks & smoke"
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
            Else
            
                Damage.Electrical = True
                UpdateMessage "Instruments: Electrical system - Catastrophic failure."
                G6ControlledBailout (OverWater())
                BL2Instruments = END_MISSION
                Exit Function
            
            End If
    
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Instruments: " & strArea
    Else
        UpdateMessage "Instruments: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' P3BombBayDamage
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  There are few differences between bomber models, so we have a common
'         chart for all of them. Model-specific damage is handled in the child
'         subroutines.
'******************************************************************************
Private Function P3BombBayDamage() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            strArea = "Equipment"
            
            If Damage.BombRelease = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "Bomb release mechanism inoperable"
                Damage.BombRelease = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
            
        Case 3, 9, 11:
            
            strArea = "Bombs"
            
            If Bomber.BombsOnBoard = True Then
                
                Select Case Random1D6()
                    
                    Case 1 To 4:
                    
                        strEffect = "No effect"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    
                    Case 5 To 6:
                    
                        UpdateMessage "Bombs: Ordnance detonates! Bomber disintegrates in mid-air!"
                        Bomber.Status = SHOT_DOWN_STATUS
                        Call CrewFinish(KIA_STATUS)
                        P3BombBayDamage = END_MISSION
                        Exit Function
                        
                End Select
            
            Else
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
        Case 4
            
            strArea = "Equipment"
            
            If Damage.RubberRafts = True _
            Or Bomber.BomberModel = AVRO_LANCASTER Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "Rubber rafts holed"
                Damage.RubberRafts = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 5, 10:
            
            strArea = "Equipment"
            
            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    If Damage.BombBayDoors = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Doors inoperable"
                        Damage.BombBayDoors = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
                Case 3 To 6:
                    
                    strArea = "Superficial damage"
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            
            End Select
        
        Case 6 To 8
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 12:
            
            strArea = "Control Cables"
            
            If ControlCableHit(strEffect) = False Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Bomb Bay: " & strArea
    Else
        UpdateMessage "Bomb Bay: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' P4RadioRoomDamage
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P4RadioRoomDamage() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            strArea = "Heater"
            
            If HeaterHit(RADIO_OPERATOR, strEffect) = False Then
                strArea = "Superficial damage"
            End If
            
        Case 3:
            
            strArea = "Equipment"
            
            If Damage.IntercomSystem = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                Damage.IntercomSystem = True
                strEffect = "Intercom out"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 4 To 5:
            
            strArea = "Equipment"
            
            If Damage.Radio = True Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                Damage.Radio = True
                strEffect = "Radio out"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 6:
            
            strArea = "Crew"
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(RADIO_OPERATOR).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 7 To 10:
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 11:
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1 To 5:
                    
                    If OxygenHit(RADIO_OPERATOR, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 6:
                    
                    Call OxygenFire(strEffect, RADIO_OPERATOR)
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Radio Room: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P4RadioRoomDamage = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select

        Case 12:
            
            strArea = "Control Cables"
            
            If ControlCableHit(strEffect) = False Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Radio Room: " & strArea
    Else
        UpdateMessage "Radio Room: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' P5WaistDamage
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P5WaistDamage() As Integer
    
    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.BomberModel = B17_F _
    Or Bomber.BomberModel = B17_G _
    Or Bomber.BomberModel = YB40 Then
    
        P5WaistDamage = P5WaistB17()
    
    ElseIf Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
    
        P5WaistDamage = P5WaistB24()
    
    Else
    
        P5WaistDamage = P5WaistLanc()
    
    End If

End Function
    
'******************************************************************************
' P5WaistB17
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P5WaistB17() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            strArea = "Oxygen supply"
            
            Select Case Random1D6()

                Case 1 To 2:
                    
                    If OxygenHit(PORT_WAIST_GUNNER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 3 To 4:
                    
                    If OxygenHit(STBD_WAIST_GUNNER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 5:
                    
                    If OxygenHit(BALL_GUNNER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 6:
                    
                    If Bomber.BomberModel = YB40 Then
                        Call OxygenFire(strEffect, PORT_WAIST_GUNNER, STBD_WAIST_GUNNER, BALL_GUNNER, MID_UPPER_GUNNER)
                    Else
                        Call OxygenFire(strEffect, PORT_WAIST_GUNNER, STBD_WAIST_GUNNER, BALL_GUNNER)
                    End If
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Waist: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P5WaistB17 = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select
            
        Case 3:
            
            strArea = "Armament"
                    
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    If Bomber.Gun(PORT_WAIST_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Port waist MG inoperable"
                        Bomber.Gun(PORT_WAIST_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
                Case 4 To 6:
                    
                    If Bomber.Gun(STBD_WAIST_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Starboard waist MG inoperable"
                        Bomber.Gun(STBD_WAIST_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
            End Select
            
        Case 4:
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 5:
            
            If Bomber.BomberModel <> YB40 Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                    
                strArea = "Mid-Upper Turret"
                
                Select Case Random1D6()
                    
                    Case 1 To 2:
                    
                        intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(MID_UPPER_GUNNER).CurrentSerialNum)
                        strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                    
                    Case 3 To 4:
                    
                        If Bomber.Gun(MID_UPPER_MG).Status = MG_INOPERABLE Then
                            strArea = "Superficial damage"
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                        Else
                            strEffect = "Mid-upper turret inoperable"
                            Bomber.Gun(MID_UPPER_MG).Status = MG_INOPERABLE
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                        End If
                    
                    Case 5:
                        
                        If HeaterHit(MID_UPPER_GUNNER, strEffect) = False Then
                            strArea = "Superficial damage"
                        End If
                    
                    Case 6:
                        
                        If OxygenHit(MID_UPPER_GUNNER, 1, strEffect) = False Then
                            strArea = "Superficial damage"
                        End If
                
                End Select
            
            End If
        
        Case 6:
            
            strArea = "Crew"
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 7:
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 8:
            
            strArea = "Crew"
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 9:
            
            strArea = "Ball Turret"
            
            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    strArea = "Crew"
                    intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(BALL_GUNNER).CurrentSerialNum)
                    strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
                Case 3:
                    
                    strArea = "Heater"
                    
                    If HeaterHit(BALL_GUNNER, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 4 To 5:
                    
                    If Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Ball MG inoperable"
                        Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If
        
                Case 6:
                    
                    If Damage.BallTurretMech = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        ' Though it may still be capable of firing, the MG
                        ' is useless without traverse. Note that the other
                        ' mechanical turrets simply have their gun set to
                        ' inoperable if the turret mechanism is jammed; the
                        ' ball turret is different due to the need to check
                        ' the gunner's status if landing wheels up.
                        strEffect = "Ball turret inoperable. Gunner trapped."
                        Damage.BallTurretMech = True
                        Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If

            End Select
            
        Case 10:
            
            strArea = "Crew"
            
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(PORT_WAIST_GUNNER).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex) & " / "
            
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(STBD_WAIST_GUNNER).CurrentSerialNum)
            strEffect = strEffect & Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
            
        Case 11:
        
            strArea = "Heater"
            
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    If HeaterHit(PORT_WAIST_GUNNER, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 4 To 6:
                    
                    If HeaterHit(STBD_WAIST_GUNNER, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
            
            End Select
                
        Case 12:
            
            strArea = "Control Cables"
            
            If ControlCableHit(strEffect) = False Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Waist: " & strArea
    Else
        UpdateMessage "Waist: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' BL4WoundB24
'
' INPUT:  The airman's position index.
'
' OUTPUT: "Superficial damage", if position is unmanned.
'
' RETURN: n/a
'
' NOTES:  B-24s are the only bomber where one position is always unmanned:
'         MGs may be unmanned on all bombers, but airmen -- fit for duty,
'         wounded or KIA -- always occupy all positions. Therefore, this is
'         a B-24 specific wrapper to the BL4Wound() function.
'******************************************************************************
Private Function BL4WoundB24(ByVal intPos As Integer, ByRef strArea As String) As String
    Dim intIndex
    
    BL4WoundB24 = ""
    intIndex = 0
    
    If PosManned(intPos) = False Then
        strArea = "Superficial damage"
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        ' Point at airman currently occupying the position.
        
        intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
    
        BL4WoundB24 = Bomber.Airman(intIndex).Name & ": " & _
                      BL4Wound(intIndex)
    End If

End Function

'******************************************************************************
' P5WaistB24
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P5WaistB24() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()

    Select Case intRoll
        Case 2:
            
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    If OxygenHit(PORT_WAIST_GUNNER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 3 To 4:
                    ' Assume a separate oxygen system from the port gunner, so
                    ' that some other crew member may man the stbd position.
                    
                    If OxygenHit(STBD_WAIST_GUNNER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 5:
                    
                    If OxygenHit(BALL_GUNNER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 6:
                    
                    Call OxygenFire(strEffect, PORT_WAIST_GUNNER, STBD_WAIST_GUNNER, BALL_GUNNER)
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Waist: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P5WaistB24 = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select
            
        Case 3:
            
            strArea = "Armament"
                    
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    If Bomber.Gun(PORT_WAIST_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Port waist MG inoperable"
                        Bomber.Gun(PORT_WAIST_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
                Case 4 To 6:
                    
                    If Bomber.Gun(STBD_WAIST_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Starboard waist MG inoperable"
                        Bomber.Gun(STBD_WAIST_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
            End Select
            
        Case 4 To 5:
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 6:

            strArea = "Crew"
            strEffect = BL4WoundB24(PORT_WAIST_GUNNER, strArea)
        
        Case 7:
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 8:
            
            strArea = "Crew"
            strEffect = BL4WoundB24(STBD_WAIST_GUNNER, strArea)
        
        Case 9:
            
            strArea = "Ball Turret"
            
            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    strArea = "Crew"
                    strEffect = BL4WoundB24(BALL_GUNNER, strArea)
        
                Case 3:
                    
                    strArea = "Heater"
                    
                    If HeaterHit(BALL_GUNNER, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 4 To 5:
                    
                    If Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    ElseIf Bomber.BomberModel = B24_D _
                    Or Bomber.BomberModel = B24_E Then
                        strEffect = "Tunnel MG inoperable"
                        Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    ElseIf Bomber.BomberModel = B24_LM Then
                        strEffect = "Floor ring MG inoperable."
                        Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    Else
                        strEffect = "Ball MG inoperable"
                        Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If
        
                Case 6:
                    
                    If Bomber.BomberModel = B24_D _
                    Or Bomber.BomberModel = B24_E _
                    Or Damage.BallTurretMech = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    ElseIf Bomber.BomberModel = B24_LM Then
                        strEffect = "Floor ring inoperable."
                        Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    Else
                        ' Though it may still be capable of firing, the MG
                        ' is useless without traverse. Note that the other
                        ' mechanical turrets simply have their gun set to
                        ' inoperable if the turret mechanism is jammed; the
                        ' ball turret is different due to the need to check
                        ' the gunner's status if landing wheels up.
                        strEffect = "Ball turret inoperable. Gunner trapped."
                        Damage.BallTurretMech = True
                        Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If
            
            End Select
            
        Case 10:
            
            strArea = "Crew"
            strEffect = BL4WoundB24(PORT_WAIST_GUNNER, strArea)
            strEffect = BL4WoundB24(STBD_WAIST_GUNNER, strArea)
        
        Case 11:
        
            strArea = "Heater"
            
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    If HeaterHit(PORT_WAIST_GUNNER, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 4 To 6:
                    
                    If HeaterHit(STBD_WAIST_GUNNER, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
            
            End Select
                
        Case 12:
            
            strArea = "Control Cables"
            
            If ControlCableHit(strEffect) = False Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Waist: " & strArea
    Else
        UpdateMessage "Waist: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' P5WaistLanc
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P5WaistLanc() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            Call P3BombBayDamage
            Exit Function
            
        Case 3:
            
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1 To 5:
                    
                    If OxygenHit(MID_UPPER_GUNNER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 6:
                    
                    Call OxygenFire(strEffect, MID_UPPER_GUNNER)
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Waist: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P5WaistLanc = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select
            
        Case 4:

            strArea = "Equipment"
            
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    If Damage.PortAmmoBox = 2 Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Port ammo box inoperable"
                        Damage.PortAmmoBox = Damage.PortAmmoBox + 1
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
                Case 4 To 6:
                    
                    If Damage.StbdAmmoBox = 2 Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Starboard ammo box inoperable"
                        Damage.StbdAmmoBox = Damage.StbdAmmoBox + 1
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                    
            End Select
                
        Case 5 To 7:
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 8:
            
            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    strArea = "Armament"
                    
                    If Bomber.Gun(MID_UPPER_MG).Status = MG_INOPERABLE Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Mid-upper turret inoperable"
                        Bomber.Gun(MID_UPPER_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If
                
                Case 3 To 5
                    
                    strArea = "Crew"
                    intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(MID_UPPER_GUNNER).CurrentSerialNum)
                    strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                
                Case 6
                    
                    strArea = "Crew / Armament"
                    
                    If Bomber.Gun(MID_UPPER_MG).Status = MG_INOPERABLE Then
                        strArea = "Crew"
                        intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(MID_UPPER_GUNNER).CurrentSerialNum)
                        strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(MID_UPPER_GUNNER).CurrentSerialNum)
                        strEffect = "Mid-upper turret inoperable and " & _
                                    Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
                        Bomber.Gun(MID_UPPER_MG).Status = MG_INOPERABLE
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
                    End If

            End Select
        
        Case 9:
            
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 10:
            
            strArea = "Equipment"
            
            Select Case Random1D6()
                
                Case 1 To 4:
                    
                    strArea = "Superficial damage"
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
                Case 5:
                    
                    If Damage.PortAmmoTrack = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Port ammo track inoperable"
                        Damage.PortAmmoTrack = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
                Case 6:
                    
                    If Damage.StbdAmmoTrack = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Starboard ammo track inoperable"
                        Damage.StbdAmmoTrack = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                    
            End Select
                
        Case 11:
        
            strArea = "Heater"
            
            If HeaterHit(MID_UPPER_GUNNER, strEffect) = False Then
                strArea = "Superficial damage"
            End If
                
        Case 12:
            
            strArea = "Control Cables"
            
            If ControlCableHit(strEffect) = False Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Waist: " & strArea
    Else
        UpdateMessage "Waist: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' RudderHit
'
' INPUT:  Whether the port or starboard rudder was hit.
'
' OUTPUT: A damage string.
'
' RETURN: n/a
'
' NOTES:  A B-17 only has one rudder, so by default it is the 'port side'.
'******************************************************************************
Private Function RudderHit(ByVal intSide As Integer, ByRef strEffect As String) As Boolean
    
    RudderHit = True

    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.BomberModel = B17_F _
    Or Bomber.BomberModel = B17_G _
    Or Bomber.BomberModel = YB40 Then
        
        If Damage.Rudder(intSide) >= 3 Then
            
            RudderHit = False
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Else
    
            Damage.Rudder(intSide) = Damage.Rudder(intSide) + 1
            Damage.PeckhamPoints = Damage.PeckhamPoints + 25
        
            If Damage.Rudder(intSide) <= 2 Then
                strEffect = "Hit"
            ElseIf Damage.Rudder(intSide) >= 3 Then
                strEffect = "Shot out"
            End If
    
        End If
    
    Else

        If Damage.Rudder(intSide) >= 2 Then
            
            RudderHit = False
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Else
    
            Damage.Rudder(intSide) = Damage.Rudder(intSide) + 1
            Damage.PeckhamPoints = Damage.PeckhamPoints + 25
        
            If Damage.Rudder(intSide) = 1 Then
                strEffect = "Hit"
            ElseIf Damage.Rudder(intSide) >= 2 Then
                strEffect = "Shot out"
            End If
    
        End If
    
    End If

End Function

'******************************************************************************
' TailplaneRootHit
'
' INPUT:  Whether the port or starboard tailplane was hit.
'
' OUTPUT: A damage string.
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  A B-17 only has one tailplane, so by default it is the 'port side'.
'******************************************************************************
Private Function TailplaneRootHit(ByVal intSide As Integer, ByRef strEffect As String) As Integer
    Dim strSide As String
    
    strSide = SideText(intSide)
    
    If Damage.TailplaneRoot(intSide) + 1 >= 3 _
    And Bomber.RabbitsFoot >= 1 Then
        ' Expend luck to prevent loss of aircraft.
        strEffect = "That was a lucky miss"
        Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
        Exit Function
    End If

    Damage.TailplaneRoot(intSide) = Damage.TailplaneRoot(intSide) + 1
                
    If Damage.TailplaneRoot(intSide) = 1 Then
        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
    End If

    If Damage.TailplaneRoot(intSide) <= 2 Then
        
        strEffect = strSide & " root hit"
    
    ElseIf Damage.TailplaneRoot(intSide) >= 3 Then
        
        strEffect = strSide & " tailplane shot off"

' TODO: B-24 loses entire empennage
    
    End If

    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.BomberModel = B17_F _
    Or Bomber.BomberModel = B17_G _
    Or Bomber.BomberModel = YB40 Then
    
        ' B-17s only had one tailplane. Port side is the default.
    
        If Bomber.RabbitsFoot >= 1 Then
            ' Expend luck to prevent loss of aircraft.
            strEffect = "That was a lucky miss"
            Damage.TailplaneRoot(intSide) = Damage.TailplaneRoot(intSide) - 1
            Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
            Exit Function
        Else
            If Damage.TailplaneRoot(intSide) >= 3 Then
                UpdateMessage "Tailplane shot off. Bomber dives out of control."
                G7UncontrolledBailout (OverWater())
                TailplaneRootHit = END_MISSION
            End If
        End If

    Else ' B-24 or Lancaster

        ' Twin-tailed bombers could fly with one missing tailplane, but not
        ' with both missing.

        If Damage.TailplaneRoot(PORT_SIDE) >= 3 _
        And Damage.TailplaneRoot(STBD_SIDE) >= 3 Then
            
            If Bomber.RabbitsFoot >= 1 Then
                ' Expend luck to prevent loss of aircraft.
                strEffect = "That was a lucky miss"
                Damage.TailplaneRoot(intSide) = Damage.TailplaneRoot(intSide) - 1
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                Exit Function
            Else
                
                UpdateMessage "Both tailplanes shot off. Bomber dives out of control."
                G7UncontrolledBailout (OverWater())
                TailplaneRootHit = END_MISSION
            
            End If
        
        End If
    
    End If
    
End Function

'******************************************************************************
' P6TailDamage
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P6TailDamage() As Integer
    
    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.BomberModel = B17_F _
    Or Bomber.BomberModel = B17_G _
    Or Bomber.BomberModel = YB40 Then
    
        P6TailDamage = P6TailB17()
    
    ElseIf Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
    
        P6TailDamage = P6TailB24()
    
    Else
    
        P6TailDamage = P6TailLanc()
    
    End If

End Function
    
'******************************************************************************
' P6TailB17
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P6TailB17() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            strArea = "Heater"
            
            If HeaterHit(TAIL_GUNNER, strEffect) = False Then
                strArea = "Superficial damage"
            End If
                
        Case 3:
            
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    strArea = "Tail Wheel"
                    
                    If Damage.Tailwheel = True Then
                        strArea = "Superficial damage"
                    Else
                        strEffect = "Tail wheel inoperable"
                        Damage.Tailwheel = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    End If
                
                Case 4 To 6:
                    
                    strArea = "Instruments"
                    
                    If Damage.Autopilot = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Autopilot inoperable"
                        Damage.Autopilot = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
            End Select
            
        Case 4:
        
            strArea = "Armament"
            
            If Bomber.Gun(TAIL_MG).Status = MG_INOPERABLE Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "Tail turret inoperable"
                Bomber.Gun(TAIL_MG).Status = MG_INOPERABLE
                Damage.PeckhamPoints = Damage.PeckhamPoints + 20
            End If
        
        Case 5:
        
            strArea = "Crew"
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(TAIL_GUNNER).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 6:
        
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 7:
        
            strArea = "Rudder"
            
            ' A B-17 only has one rudder, so by default it is the 'port side'.
            
            If RudderHit(PORT_SIDE, strEffect) = False Then
                strArea = "Superficial damage"
            End If
        
        Case 8:

            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5

        Case 9 To 10:
        
            strArea = "Tailplane"

            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    strArea = "Superficial damage"
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 5

                Case 3:
                    
                    If Damage.Elevator(PORT_SIDE) = True Then
                        strArea = "Superficial damage"
                    Else
                        strEffect = "Port elevator inoperable"
                        Damage.Elevator(PORT_SIDE) = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    End If
                
                Case 4:
                    
                    If Damage.Elevator(STBD_SIDE) = True Then
                        strArea = "Superficial damage"
                    Else
                        strEffect = "Starboard elevator inoperable"
                        Damage.Elevator(STBD_SIDE) = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    End If
                
                Case 5:
                    
                    P6TailB17 = TailplaneRootHit(PORT_SIDE, strEffect)
        
                Case 6:
                    
                    P6TailB17 = TailplaneRootHit(STBD_SIDE, strEffect)
        
            End Select
        
        Case 11:
        
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1 To 5:
                    
                    If OxygenHit(TAIL_GUNNER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 6:
                    
                    Call OxygenFire(strEffect, TAIL_GUNNER)
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Tail: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P6TailB17 = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select
            
        Case 12:
            
            strArea = "Control Cables"
            
            If ControlCableHit(strEffect) = False Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Tail: " & strArea
    Else
        UpdateMessage "Tail: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' P6TailB24
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P6TailB24() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            strArea = "Heater"
            
            If HeaterHit(TAIL_GUNNER, strEffect) = False Then
                strArea = "Superficial damage"
            End If
                
        Case 3:
            
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    strArea = "Superficial damage"
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                
                Case 4 To 6:
                    
                    strArea = "Instruments"
                    
                    If Damage.Autopilot = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Autopilot inoperable"
                        Damage.Autopilot = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
            End Select
            
        Case 4:
        
            strArea = "Armament"
            
            If Bomber.Gun(TAIL_MG).Status = MG_INOPERABLE Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "Tail turret inoperable"
                Bomber.Gun(TAIL_MG).Status = MG_INOPERABLE
                Damage.PeckhamPoints = Damage.PeckhamPoints + 20
            End If
        
        Case 5:
        
            strArea = "Crew"
            strEffect = BL4WoundB24(TAIL_GUNNER, strArea)
        
        Case 6:
        
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 7:
        
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    strArea = "Port Rudder"
            
                    If RudderHit(PORT_SIDE, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
        
                Case 4 To 6:
                    
                    strArea = "Starboard Rudder"
            
                    If RudderHit(STBD_SIDE, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
        
            End Select
        
        Case 8:

            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5

        Case 9 To 10:
        
            strArea = "Tailplane"

            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    strArea = "Superficial damage"
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 5

                Case 3:
                    
                    If Damage.Elevator(PORT_SIDE) = True Then
                        strArea = "Superficial damage"
                    Else
                        strEffect = "Port elevator inoperable"
                        Damage.Elevator(PORT_SIDE) = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    End If
                
                Case 4:
                    
                    If Damage.Elevator(STBD_SIDE) = True Then
                        strArea = "Superficial damage"
                    Else
                        strEffect = "Starboard elevator inoperable"
                        Damage.Elevator(STBD_SIDE) = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    End If
                
                Case 5:
                    
                    P6TailB24 = TailplaneRootHit(PORT_SIDE, strEffect)
        
                Case 6:
                    
                    P6TailB24 = TailplaneRootHit(STBD_SIDE, strEffect)
        
            End Select
        
        Case 11:
        
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1 To 5:
                    
                    If OxygenHit(TAIL_GUNNER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 6:
                    
                    Call OxygenFire(strEffect, TAIL_GUNNER)
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Tail: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P6TailB24 = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select
            
        Case 12:
            
            strArea = "Control Cables"
            
            If ControlCableHit(strEffect) = False Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Tail: " & strArea
    Else
        UpdateMessage "Tail: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' P6TailLanc
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function P6TailLanc() As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intIndex As Integer

    intRoll = Random2D6()
    
    Select Case intRoll
        Case 2:
            
            strArea = "Heater"
            
            If HeaterHit(TAIL_GUNNER, strEffect) = False Then
                strArea = "Superficial damage"
            End If
                
        Case 3:
            
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    strArea = "Tail Wheel"
                    
                    If Damage.Tailwheel = True Then
                        strArea = "Superficial damage"
                    Else
                        strEffect = "Tail wheel inoperable"
                        Damage.Tailwheel = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    End If
                
                Case 4 To 6:
                    
                    strArea = "Instruments"
                    
                    If Damage.Autopilot = True Then
                        strArea = "Superficial damage"
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                    Else
                        strEffect = "Autopilot inoperable"
                        Damage.Autopilot = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    End If
                
            End Select
            
        Case 4:
        
            strArea = "Armament"
            
            If Bomber.Gun(TAIL_MG).Status = MG_INOPERABLE Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "Tail turret inoperable"
                Bomber.Gun(TAIL_MG).Status = MG_INOPERABLE
                Damage.PeckhamPoints = Damage.PeckhamPoints + 20
            End If
        
        Case 5:
        
            strArea = "Crew"
            intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(TAIL_GUNNER).CurrentSerialNum)
            strEffect = Bomber.Airman(intIndex).Name & ": " & BL4Wound(intIndex)
        
        Case 6:
        
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 7:
        
            Select Case Random1D6()
                
                Case 1 To 3:
                
                    strArea = "Port Rudder"
            
                    If RudderHit(PORT_SIDE, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                    
                Case 4 To 6:
                
                    strArea = "Starboard Rudder"
            
                    If RudderHit(STBD_SIDE, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
        
            End Select
        
        Case 8:

            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5

        Case 9 To 10:
        
            strArea = "Tailplane"

            Select Case Random1D6()
                
                Case 1 To 2:
                    
                    strArea = "Superficial damage"
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 5

                Case 3:
                    
                    If Damage.Elevator(PORT_SIDE) = True Then
                        strArea = "Superficial damage"
                    Else
                        strEffect = "Port elevator inoperable"
                        Damage.Elevator(PORT_SIDE) = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    End If
                
                Case 4:
                    
                    If Damage.Elevator(STBD_SIDE) = True Then
                        strArea = "Superficial damage"
                    Else
                        strEffect = "Starboard elevator inoperable"
                        Damage.Elevator(STBD_SIDE) = True
                        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    End If
                
                Case 5:
                    
                    P6TailLanc = TailplaneRootHit(PORT_SIDE, strEffect)
        
                Case 6:
                    
                    P6TailLanc = TailplaneRootHit(STBD_SIDE, strEffect)
        
            End Select
        
        Case 11:
        
            strArea = "Oxygen supply"
            
            Select Case Random1D6()
                
                Case 1 To 5:
                    
                    If OxygenHit(TAIL_GUNNER, 1, strEffect) = False Then
                        strArea = "Superficial damage"
                    End If
                
                Case 6:
                    
                    Call OxygenFire(strEffect, TAIL_GUNNER)
                    
                    If BL3ExtinguishFire(Bomber.HandHeldExtinguishers, False) = False Then
                        UpdateMessage "Tail: Uncontrolled oxygen fire."
                        G6ControlledBailout (OverWater())
                        P6TailLanc = END_MISSION
                        Exit Function
                    Else
                        strEffect = strEffect & " " & _
                                    Bomber.HandHeldExtinguishers & _
                                    " handheld extinguishers remain."
                    End If

            End Select
            
        Case 12:
            
            strArea = "Control Cables"
            
            If ControlCableHit(strEffect) = False Then
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            End If
        
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage "Tail: " & strArea
    Else
        UpdateMessage "Tail: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' EngineHit
'
' INPUT:  The number of the affected engine.
'
' OUTPUT: A damage string.
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function EngineHit(ByVal intEngine As Integer, ByRef strEffect As String) As Integer
    
    Dim intRoll As Integer
    
    intRoll = Random1D6()
    
    Select Case intRoll
        Case 1 To 2:
            
            strEffect = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5

        Case 3 To 4:
            
            If Damage.EngineOut(intEngine) = True Then
    
                strEffect = "Superficial damage"

            Else
            
                ' BL-1 Wings: Note c.
                
                If Damage.FeatheringCtrl = True Then
                
                    strEffect = "Out; prop not feathered."
                    Damage.EngineOut(intEngine) = True
                    Damage.EngineDrag(intEngine) = True
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    Call DropOutOfFormation

                Else
                
                    Select Case Random1D6()
                        
                        Case 1 To 5:
                            
                            strEffect = "Out; prop feathered."
                            Damage.EngineOut(intEngine) = True
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                            
                        Case 6:
                        
                            strEffect = "Out; prop not feathered."
                            Damage.EngineOut(intEngine) = True
                            Damage.EngineDrag(intEngine) = True
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                            Call DropOutOfFormation
                    
                    End Select
                
                End If
            
            End If
            
        Case 5:
            
            If Damage.EngineOut(intEngine) = True Then
    
                strEffect = "Superficial damage"

            Else
            
                ' BL-1 Wings: Note d.
                
                If Damage.FeatheringCtrl = True Then
                
                    If Bomber.RabbitsFoot >= 1 Then
                        ' Expend luck to prevent loss of aircraft.
                        strEffect = "That was a lucky miss"
                        Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                    Else
                    
                        UpdateMessage "Wing: Runaway #" & intEngine & " engine."
                        G6ControlledBailout (OverWater())
                        EngineHit = END_MISSION
                    
                    End If
                    
                Else
                
                    Select Case Random1D6()
                        
                        Case 1 To 5:
                            
                            strEffect = "Runaway; engine shut off and " & _
                                        "prop feathered."
                            Damage.EngineOut(intEngine) = True
                            Damage.EngineDrag(intEngine) = True
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                            
                        Case 6:
                        
                            If Bomber.RabbitsFoot >= 1 Then
                                ' Expend luck to prevent loss of aircraft.
                                strEffect = "That was a lucky miss"
                                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                            Else
                        
                                UpdateMessage "Wing: Runaway #" & intEngine & " engine."
                                G6ControlledBailout (OverWater())
                                EngineHit = END_MISSION
                    
                            End If
                    
                    End Select
                
                End If
            
            End If
            
        Case 6:
            
            ' BL-1 Wings: Note e.
            
            EngineHit = OilTankHit(intEngine, strEffect)
            
    End Select
    
End Function

'******************************************************************************
' OilTankHit
'
' INPUT:  The number of the affected engine.
'
' OUTPUT: A damage string.
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Public Function OilTankHit(ByVal intEngine As Integer, ByRef strEffect As String) As Integer
    
    ' BL-1 Wings: Note e.
    
    Dim intRoll As Integer
    Dim intLeakRoll As Integer
    
    If Damage.OilTankLeak(intEngine) = NO_OIL Then
    
        strEffect = "Superficial damage"
    
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
                
    Else
    
        intRoll = Random1D6()
        
        Select Case intRoll
            
            Case 1 To 2:
                
                If BL3ExtinguishFire(Bomber.EngineExtinguisher(intEngine), True) = False Then
                    UpdateMessage "Wing: Uncontrolled oil tank fire, #" & _
                                  intEngine & " engine."
                    G6ControlledBailout (OverWater())
                    OilTankHit = END_MISSION
                Else
                    strEffect = "#" & intEngine & " oil tank fire out; " & _
                                Bomber.EngineExtinguisher(intEngine) & _
                                " extinguishers remain."
                
                End If
    
            Case 3 To 4:
                
                If Damage.OilTankLeak(intEngine) = NO_LEAK Then
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                Else
                    Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                End If
                
                intLeakRoll = Random1D6()
        
                Select Case intLeakRoll
                    
                    Case 1 To 2:
                                
                        ' Heavy leak.
                                
                        If Damage.OilTankLeak(intEngine) >= LT_LEAK Then
                            ' Tank is emptied.
                            Damage.OilTankLeak(intEngine) = NO_OIL
                        Else
                            ' Increment leakage.
                            Damage.OilTankLeak(intEngine) = Damage.OilTankLeak(intEngine) + HVY_LEAK
                        End If
                                
                    Case 3 To 4:
                                
                        ' Medium leak.
                                
                        If Damage.OilTankLeak(intEngine) >= MED_LEAK Then
                            ' Tank is emptied.
                            Damage.OilTankLeak(intEngine) = NO_OIL
                        Else
                            ' Increment leakage.
                            Damage.OilTankLeak(intEngine) = Damage.OilTankLeak(intEngine) + MED_LEAK
                        End If
                                
                    Case 5 To 6:
                    
                        ' Light leak.
                        
                        If Damage.OilTankLeak(intEngine) = HVY_LEAK Then
                            ' Tank is emptied.
                            Damage.OilTankLeak(intEngine) = NO_OIL
                        Else
                            ' Increment leakage.
                            Damage.OilTankLeak(intEngine) = Damage.OilTankLeak(intEngine) + LT_LEAK
                        End If
                                
                End Select
        
                strEffect = "#" & intEngine & " oil tank leak; shut off engine after " & _
                            CStr(4 - Damage.OilTankLeak(intEngine)) & _
                            " more zones."
                
            Case 5 To 6:
                    
                strEffect = "#" & intEngine & " oil tank leak; self-sealed."
        
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                    
        End Select
    
    End If
    
End Function

'******************************************************************************
' FuelTankHit
'
' INPUT:  n/a
'
' OUTPUT: A damage string.
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function FuelTankHit(ByRef strEffect As String) As Integer
    
    Dim intRoll As Integer
    
    intRoll = Random1D6()
'intRoll = 2 ' debug Nov04
    If Bomber.BomberModel = B17_C Then
        ' Variant: B-17Cs did not have self-sealing tanks.
        intRoll = intRoll - 1
    End If
    
    Select Case intRoll
        
        Case 0 To 2:
            
            ' BL-1 Wings: Note f.
    
            If Bomber.RabbitsFoot >= 1 Then
                ' Expend luck to prevent loss of aircraft.
                strEffect = "That was a lucky miss"
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
            Else
    
                Select Case Random1D6()
                    
                    Case 1 To 5:
                                
                        UpdateMessage "Wing: Fuel tank fire."
                        G6ControlledBailout (OverWater())
                        FuelTankHit = END_MISSION
                    
                    Case 6:
                                
                        UpdateMessage "Wing: Fuel tank explodes. Wing rips " & _
                                      "off; bomber dives out of control."
                        G7UncontrolledBailout (OverWater())
                        FuelTankHit = END_MISSION
                
                End Select

            End If
        
        Case 3 To 4:
            
            ' BL-1 Wings: Note g, plus variant.

            If Damage.FuelTankHits <= 2 Then
                Damage.FuelTankHits = Damage.FuelTankHits + 1
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
            End If
            
            strEffect = "Fuel tank leak; must land bomber in "
    
            Select Case Damage.FuelTankHits
                Case 1:
                
                    Select Case Random1D6()
                        
                        Case 1:
                            
                            Bomber.FuelPoints = 4
                        
                        Case 2:
                            
                            If Bomber.BomberModel = B24_D _
                            Or Bomber.BomberModel = B24_E _
                            Or Bomber.BomberModel = B24_GHJ _
                            Or Bomber.BomberModel = B24_LM Then
                                
                                Bomber.FuelPoints = 5
                            
                            Else
                                
                                Bomber.FuelPoints = 4
                            
                            End If
                        
                        Case 3:
                            
                            Bomber.FuelPoints = 5
                        
                        Case 4:
                            
                            If Bomber.BomberModel = B24_D _
                            Or Bomber.BomberModel = B24_E _
                            Or Bomber.BomberModel = B24_GHJ _
                            Or Bomber.BomberModel = B24_LM Then
                                
                                Bomber.FuelPoints = 6
                            
                            Else
                                
                                Bomber.FuelPoints = 5
                            
                            End If
                        
                        Case 5 To 6:
                            
                            Bomber.FuelPoints = 6
                    
                    End Select
                
                Case 2:
                    
                    If Bomber.BomberModel = B24_D _
                    Or Bomber.BomberModel = B24_E _
                    Or Bomber.BomberModel = B24_GHJ _
                    Or Bomber.BomberModel = B24_LM Then
    
                        Bomber.FuelPoints = 3
                        
                    Else
                    
                        Bomber.FuelPoints = 2
                        
                    End If
                    
                    
                Case 3:
                    
                    Bomber.FuelPoints = 1
                    
            End Select
            
            strEffect = strEffect & Bomber.FuelPoints & " more zones."
            
        Case 5 To 6:
                
            strEffect = "Fuel tank hit; self-sealed."
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    
    End Select
    
End Function

'******************************************************************************
' WingRootHit
'
' INPUT:  The side of the affected wing.
'
' OUTPUT: A damage string.
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function WingRootHit(ByVal intSide As Integer, ByRef strEffect As String) As Integer
    
    Damage.WingRoot(intSide) = Damage.WingRoot(intSide) + 1
    
    If Damage.WingRoot(intSide) = 1 Then
        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
    End If

    strEffect = "Hit"

    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
    
        If Damage.WingRoot(intSide) >= 3 Then
            
            If Bomber.RabbitsFoot >= 1 Then
                ' Expend luck to prevent loss of aircraft.
                Damage.WingRoot(intSide) = 2
                strEffect = "That was a lucky miss"
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
            Else
    
                UpdateMessage SideText(intSide) & _
                              " wing shot off; bomber dives out of control."
                G7UncontrolledBailout (OverWater())
                WingRootHit = END_MISSION
            
            End If
        
        End If

    Else ' B-17 or Lancaster

        If Damage.WingRoot(intSide) >= 5 Then
            
            If Bomber.RabbitsFoot >= 1 Then
                ' Expend luck to prevent loss of aircraft.
                Damage.WingRoot(intSide) = 4
                strEffect = "That was a lucky miss"
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
            Else
    
                UpdateMessage SideText(intSide) & _
                              " wing shot off; bomber dives out of control."
                G7UncontrolledBailout (OverWater())
                WingRootHit = END_MISSION
        
            End If
        
        End If

    End If
    
End Function

'******************************************************************************
' LandingGearHit
'
' INPUT:  n/a
'
' OUTPUT: A damage string.
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub LandingGearHit(ByRef strEffect As String)

    Dim intRoll As Integer
    
    intRoll = Random1D6()
    
    Select Case intRoll
            
        Case 1 To 3:
            
            If Damage.Brake = True Then
                strEffect = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "Brakes shot out"
                Damage.Brake = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
            End If
        
        Case 4 To 6:
            
            If Damage.LandingGear = True Then
                strEffect = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "Inoperable"
                Damage.LandingGear = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
            End If
        
    End Select
            
End Sub

'******************************************************************************
' AileronHit
'
' INPUT:  The side of the affected aileron.
'
' OUTPUT: A damage string.
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub AileronHit(ByVal intSide As Integer, ByRef strEffect As String)

    Dim intRoll As Integer
    
    intRoll = Random1D6()
    
    Select Case intRoll
            
        Case 1 To 3:
            
            If Damage.Aileron(intSide) = True Then
                strEffect = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "Shot out"
                Damage.Aileron(intSide) = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
            End If
        
        Case 4 To 6:
            
            strEffect = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
    End Select

End Sub
    
'******************************************************************************
' FlapHit
'
' INPUT:  The side of the affected flap.
'
' OUTPUT: A damage string.
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub FlapHit(ByVal intSide As Integer, ByRef strEffect As String)

    Dim intRoll As Integer
    
    intRoll = Random1D6()
    
    Select Case intRoll
            
        Case 1 To 3:
            
            If Damage.WingFlap(intSide) = True Then
                strEffect = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            Else
                strEffect = "Shot out"
                Damage.WingFlap(intSide) = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
            End If
        
        Case 4 To 6:
            
            strEffect = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
    End Select

End Sub
    
'******************************************************************************
' SideText
'
' INPUT:  Side as a number.
'
' OUTPUT: n/a
'
' RETURN: Side as text.
'
' NOTES:  n/a
'******************************************************************************
Private Function SideText(ByVal intSide As Integer) As String
    
    If intSide = PORT_SIDE Then
        SideText = "Port"
    Else
        SideText = "Starboard"
    End If

End Function

'******************************************************************************
' WeatherText
'
' INPUT:  Weather as a number.
'
' OUTPUT: n/a
'
' RETURN: Weather as text.
'
' NOTES:  n/a
'******************************************************************************
Public Function WeatherText(ByVal intVal As Integer) As String
    
    WeatherText = ""
    
    If intVal = CLEAR_WEATHER Then
        WeatherText = "Clear"
    ElseIf intVal = GOOD_WEATHER Then
        WeatherText = "Good"
    ElseIf intVal = POOR_WEATHER Then
        WeatherText = "Poor"
    ElseIf intVal = BAD_WEATHER Then
        WeatherText = "Bad"
    ElseIf intVal = STORM_WEATHER Then
        WeatherText = "Storm"
    End If

End Function

'******************************************************************************
' BL1WingDamage
'
' INPUT:  The side of the affected wing.
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  There are few differences between bomber models, so we have a common
'         chart for all of them. Model-specific damage is handled in the child
'         subroutines.
'******************************************************************************
Private Function BL1WingDamage(ByVal intSide As Integer) As Integer
    
    Dim intRoll As Integer
    Dim strArea As String
    Dim strEffect As String
    Dim intEngine As Integer
    
    intRoll = Random2D6()
'intRoll = 2 ' debug Nov04
'intSide = PORT_SIDE ' debug Nov04
    Select Case intRoll
        Case 2 To 3:
            
            strArea = "Root"
            
            If WingRootHit(intSide, strEffect) = END_MISSION Then
                BL1WingDamage = END_MISSION
                Exit Function
            End If
                
        Case 4:
            
            strArea = "Flap"
            
            Call FlapHit(intSide, strEffect)
            
        Case 5:
        
            strArea = "Aileron"
            
            Call AileronHit(intSide, strEffect)
            
        Case 6:
        
            If Bomber.BomberModel = AVRO_LANCASTER _
            And intSide = STBD_SIDE _
            And Damage.RubberRafts = False Then
            
                strEffect = "Rubber rafts holed"
                Damage.RubberRafts = True
                Damage.PeckhamPoints = Damage.PeckhamPoints + 10
                
            Else
                
                strArea = "Superficial damage"
                Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            
            End If
        
        Case 7 To 8:
        
            strArea = "Superficial damage"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
        Case 9:
        
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    If intSide = PORT_SIDE Then
                        strArea = "#1 Engine"
                        intEngine = 1
                    Else
                        strArea = "#3 Engine"
                        intEngine = 3
                    End If

                Case 4 To 6:
                    
                    If intSide = PORT_SIDE Then
                        strArea = "#2 Engine"
                        intEngine = 2
                    Else
                        strArea = "#4 Engine"
                        intEngine = 4
                    End If

            End Select

            If EngineHit(intEngine, strEffect) = END_MISSION Then
                BL1WingDamage = END_MISSION
                Exit Function
            End If

        Case 10:
        
            Select Case Random1D6()
                
                Case 1 To 3:
                    
                    strArea = "Outboard Fuel Tank"
                    
                Case 4 To 6:
                    
                    strArea = "Inboard Fuel Tank"
                    
            End Select

            If FuelTankHit(strEffect) = END_MISSION Then
                BL1WingDamage = END_MISSION
                Exit Function
            End If
            
        Case 11:
        
            strArea = "Root"
            
            If WingRootHit(intSide, strEffect) = END_MISSION Then
                BL1WingDamage = END_MISSION
                Exit Function
            End If
        
        Case 12:
            
            strArea = "Landing Gear"
            
            Call LandingGearHit(strEffect)
                
    End Select

    If strArea = "Superficial damage" Then
        UpdateMessage SideText(intSide) & " Wing: " & strArea
    Else
        UpdateMessage SideText(intSide) & " Wing: " & strArea & " - " & strEffect
    End If

End Function

'******************************************************************************
' KillAirman
'
' INPUT:  The position currently occupied by the airman who is being killed.
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  n/a
'******************************************************************************
Private Sub KillAirman(ByVal intPos As Integer)
    
    Dim intIndex As Integer

    If PosOccupied(intPos) = True Then
    
        ' The position exists on this bomber, and it is currently occupied.
        ' Determine who is occupying the position.
    
        intIndex = GetAirmanIndexBySerialNumber(Bomber.Position(intPos).CurrentSerialNum)
    
        ' If the airman occupying the position is not dead, kill him.
        
        If Bomber.Airman(intIndex).Status <> KIA_STATUS Then
            Bomber.Airman(intIndex).Status = KIA_STATUS
            Bomber.Airman(intIndex).Wounded = True
            UpdateMessage Bomber.Airman(intIndex).Name & ": KIA"
        End If
    
    End If

    Damage.PeckhamPoints = Damage.PeckhamPoints + 5

End Sub

'******************************************************************************
' BurstInPlane
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Public Function BurstInPlane() As Integer

    ' Rule 19.0
    
    Dim intRoll As Integer
    Dim blnPassThru As Boolean
    
    blnPassThru = False
    
    intRoll = Random2D6()
    
    Select Case intRoll
        
        Case 2:
            
            If Bomber.RabbitsFoot = 0 Then
                BurstInPlane = BurstInBombBay()
            Else
                ' Expend luck to prevent severe damage to aircraft.
                UpdateMessage "Flak shell passed through bomb bay without exploding."
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                blnPassThru = True
            End If
        
        Case 3, 6:
            
            If Bomber.RabbitsFoot = 0 Then
                BurstInPlane = BurstInWing(PORT_SIDE)
            Else
                ' Expend luck to prevent loss of aircraft.
                UpdateMessage "Flak shell passed through port wing without exploding."
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                blnPassThru = True
            End If
        
        Case 4:
            
            ' B-24 and Lancaster radio operators were located on the flight
            ' deck, not in a separate compartment.
    
            If Bomber.BomberModel = B24_D _
            Or Bomber.BomberModel = B24_E _
            Or Bomber.BomberModel = B24_GHJ _
            Or Bomber.BomberModel = B24_LM _
            Or Bomber.BomberModel = AVRO_LANCASTER Then
                UpdateMessage "Flak shell passed through flight deck without exploding."
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                blnPassThru = True
            ElseIf Bomber.RabbitsFoot >= 1 Then
                ' Expend luck to prevent severe damage to aircraft.
                UpdateMessage "Flak shell passed through radio room without exploding."
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                blnPassThru = True
            Else
                BurstInRadioRoom
            End If
        
        Case 5, 7:
            
            If Bomber.RabbitsFoot = 0 Then
                BurstInPlane = BurstInTail()
            Else
                ' Expend luck to prevent loss of aircraft.
                UpdateMessage "Flak shell passed through tail without exploding."
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                blnPassThru = True
            End If
        
        Case 8, 11:
            
            If Bomber.RabbitsFoot = 0 Then
                BurstInPlane = BurstInWing(STBD_SIDE)
            Else
                ' Expend luck to prevent loss of aircraft.
                UpdateMessage "Flak shell passed through starboard wing without exploding."
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                blnPassThru = True
            End If
        
        Case 9:
            
            If Bomber.RabbitsFoot = 0 Then
                BurstInWaist
            Else
                ' Expend luck to prevent severe damage to aircraft.
                UpdateMessage "Flak shell passed through waist without exploding."
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                blnPassThru = True
            End If
        
        Case 10:
            
            If Bomber.RabbitsFoot = 0 Then
                BurstInNose
            Else
                ' Expend luck to prevent severe damage to aircraft.
                UpdateMessage "Flak burst rocks nose."
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                blnPassThru = True
            End If
        
        Case 12:
            
            If Bomber.RabbitsFoot = 0 Then
                BurstInPlane = BurstInFlightDeck()
            Else
                ' Expend luck to prevent loss of aircraft.
                UpdateMessage "Flak shell passed through flight deck without exploding."
                Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                blnPassThru = True
            End If
        
    End Select

    If blnPassThru = False Then
        Damage.BurstInPlane = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 100
        
        Call DropOutOfFormation

        If AlpsDirection() = ALPS_BELOW Then
            
            ' Bomber must maintain high altitude over the Alps. Since
            ' it cannot, the crew must bailout.
            UpdateMessage "Bomber cannot maintain altitude in Alps."
            G6ControlledBailout (False)
            BurstInPlane = END_MISSION
            Exit Function
        
        Else
            
            Call LoseAltitude
        
        End If
    
    End If

    'TODO: Results
    'x    Out of Formation
    '    Two turns in each following zone (13.0) ???
    'x    No evasive action
    'x    automatically irreparably damaged
    
End Function

'******************************************************************************
' BurstInNose
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is bad, but not bad enough to imediately end the mission.
'******************************************************************************
Private Sub BurstInNose()
    
    UpdateMessage "Flak burst in nose!"
    
    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.BomberModel = B17_F _
    Or Bomber.BomberModel = B17_G _
    Or Bomber.BomberModel = YB40 Then
    
        Call BurstInNoseB17
    
    ElseIf Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
    
        Call BurstInNoseB24
    
    Else
    
        Call BurstInNoseLanc
    
    End If

End Sub
    
'******************************************************************************
' BurstInNoseB17
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is bad, but not bad enough to imediately end the mission.
'******************************************************************************
Private Sub BurstInNoseB17()
    
    ' Rule 19.2.d
    
    Dim strEffect As String
    
    ' Case 2:
            
    If Bomber.BomberModel = YB40 _
    Or Damage.BombSight = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Bomb sight destroyed."
        Damage.BombSight = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
    
    ' Case 3:
            
    If Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    ElseIf Bomber.BomberModel = B17_G _
    Or Bomber.BomberModel = YB40 Then
        UpdateMessage "Nose turret destroyed."
        Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
    Else
        UpdateMessage "Nose gun destroyed."
        Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
            
    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.Gun(PORT_CHEEK_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Port cheek gun destroyed."
        Bomber.Gun(PORT_CHEEK_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If

    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.Gun(STBD_CHEEK_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Starboard cheek gun destroyed."
        Bomber.Gun(STBD_CHEEK_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
            
    ' Case 4:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            
    ' Case 5:

    Call KillAirman(NAVIGATOR)
        
    ' Case 6:
            
    Call KillAirman(BOMBARDIER)
        
    ' Case 7 To 9:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 15
        
    ' Case 10:
            
    If Damage.NavigationEquipment = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Navigation equipment destroyed."
        Damage.NavigationEquipment = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
                
    If Bomber.BomberModel = YB40 _
    Or Damage.BombControls = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Bomb controls destroyed."
        Damage.BombControls = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
        
    ' Case 11:
            
    If Bomber.BomberModel = YB40 Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        Call HeaterHit(BOMBARDIER, strEffect)
    End If
                
    If Bomber.BomberModel = B17_C Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        Call HeaterHit(NAVIGATOR, strEffect)
    End If
                
    ' Case 12:
            
    If Bomber.BomberModel = YB40 Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        Call OxygenHit(BOMBARDIER, 2, strEffect)
    End If
                
    If Bomber.BomberModel = B17_C Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        Call OxygenHit(NAVIGATOR, 2, strEffect)
    End If
                
End Sub

'******************************************************************************
' BurstInNoseB24
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is bad, but not bad enough to imediately end the mission.
'******************************************************************************
Private Sub BurstInNoseB24()
    
    ' Rule 19.2.d
    
    Dim strEffect As String

    ' Case 2:
            
    If Damage.BombSight = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Bomb sight destroyed."
        Damage.BombSight = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
                
    If Damage.NoseWheel = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Bomb sight destroyed."
        Damage.NoseWheel = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 25
    End If
                
    ' Case 3:
            
    If Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    ElseIf Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E Then
        UpdateMessage "Nose gun destroyed."
        Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    Else
        UpdateMessage "Nose turret destroyed."
        Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
    End If
                    
    If Bomber.BomberModel = B24_E Then

        If Bomber.Gun(PORT_CHEEK_MG).Status = MG_INOPERABLE Then
            ' Superficial damage
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        Else
            UpdateMessage "Port cheek gun destroyed."
            Bomber.Gun(PORT_CHEEK_MG).Status = MG_INOPERABLE
            Damage.PeckhamPoints = Damage.PeckhamPoints + 10
        End If
    
        If Bomber.Gun(STBD_CHEEK_MG).Status = MG_INOPERABLE Then
            ' Superficial damage
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        Else
            UpdateMessage "Starboard cheek gun destroyed."
            Bomber.Gun(STBD_CHEEK_MG).Status = MG_INOPERABLE
            Damage.PeckhamPoints = Damage.PeckhamPoints + 10
        End If
    
    End If

    ' Case 4:
            
    Call KillAirman(BOMBARDIER)
            
    ' Case 5:
            
    Call KillAirman(NAVIGATOR)
        
    ' Case 6:
            
    Call KillAirman(NOSE_GUNNER)
        
    ' Case 7 To 9:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 15
        
    ' Case 10:
            
    If Damage.NavigationEquipment = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Navigation equipment destroyed."
        Damage.NavigationEquipment = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
                
    If Damage.BombControls = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Bomb controls destroyed."
        Damage.BombControls = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
        
    ' Case 11:
            
    Call HeaterHit(BOMBARDIER, strEffect)
                
    Call HeaterHit(NAVIGATOR, strEffect)
    
    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        Call HeaterHit(NOSE_GUNNER, strEffect)
    End If
                
    ' Case 12:
            
    Call OxygenHit(BOMBARDIER, 2, strEffect)
    
    Call OxygenHit(NAVIGATOR, 2, strEffect)
    
    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        Call OxygenHit(NOSE_GUNNER, 2, strEffect)
    End If
                
End Sub

'******************************************************************************
' BurstInNoseLanc
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is bad, but not bad enough to imediately end the mission.
'******************************************************************************
Private Sub BurstInNoseLanc()
    
    ' Rule 19.2.d
    
    Dim strEffect As String

    ' Case 2:
            
    If Damage.BombSight = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Bomb sight destroyed."
        Damage.BombSight = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
                
    ' Case 3:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    
    ' Case 4:
            
    If Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Nose turret destroyed."
        Bomber.Gun(NOSE_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
    End If
            
    ' Case 5:
            
    Call KillAirman(BOMBARDIER)
        
    ' Case 6 To 9:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 20
        
    ' Case 10:
            
    If Damage.BombControls = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Bomb controls destroyed."
        Damage.BombControls = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
        
    ' Case 11:
            
    Call HeaterHit(BOMBARDIER, strEffect)
        
    ' Case 12:
            
    Call OxygenHit(BOMBARDIER, 2, strEffect)
                    
End Sub

'******************************************************************************
' BurstInFlightDeck
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission.
'
' NOTES:  n/a
'******************************************************************************
Public Function BurstInFlightDeck() As Integer

    ' Rule 19.2.b
    
    UpdateMessage "Flak burst in flight deck!"
    
    Call KillAirman(PILOT)
    
    If Bomber.BomberModel <> AVRO_LANCASTER Then
        Call KillAirman(COPILOT)
    End If
    
    If Bomber.BomberModel = AVRO_LANCASTER Then
        Call KillAirman(NAVIGATOR)
    End If
    
    Call KillAirman(ENGINEER)
    
    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM _
    Or Bomber.BomberModel = AVRO_LANCASTER Then
        Call KillAirman(RADIO_OPERATOR)
    End If
    
    UpdateMessage "Bomber dives out of control."
    
    G7UncontrolledBailout (OverWater())
    BurstInFlightDeck = END_MISSION

End Function

'******************************************************************************
' BurstInBombBay
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Public Function BurstInBombBay() As Integer

    Dim strEffect As String
    
    ' Rule 19.2.c and 19.2.d
    
    UpdateMessage "Flak burst in bomb bay!"
    
    If Bomber.BombsOnBoard = True _
    Or Bomber.ExtraAmmo = True _
    Or Bomber.ExtraFuelInBombBay = True Then
        
        UpdateMessage "Ordnance detonates! Bomber disintegrates in mid-air!"
        Bomber.Status = SHOT_DOWN_STATUS
        Call CrewFinish(KIA_STATUS)
        BurstInBombBay = END_MISSION
        
    Else
    
        ' Case 2:
        
        If Damage.BombRelease = True Then
            ' Superficial damage
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        Else
            UpdateMessage "Bomb release mechanism destroyed."
            Damage.BombRelease = True
            Damage.PeckhamPoints = Damage.PeckhamPoints + 10
        End If
    
        ' Case 3, 9, 11:
        
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 15
            
        ' Case 4:
        
        If Damage.RubberRafts = True _
        Or Bomber.BomberModel = AVRO_LANCASTER Then
            ' Superficial damage
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        Else
            UpdateMessage "Rubber rafts destroyed."
            Damage.RubberRafts = True
            Damage.PeckhamPoints = Damage.PeckhamPoints + 10
        End If
        
        ' Case 5 To 8
        
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
        
        ' Case 10:
            
        If Damage.BombBayDoors = True Then
            ' Superficial damage
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        Else
            UpdateMessage "Bomb bay doors blown off."
            Damage.BombBayDoors = True
            Damage.PeckhamPoints = Damage.PeckhamPoints + 25
        End If
        
        ' Case 12:
            
        Call ControlCableHit(strEffect)
        Call ControlCableHit(strEffect)
        UpdateMessage "Both control cables severed."
        
    End If

End Function

'******************************************************************************
' BurstInRadioRoom
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is bad, but not bad enough to imediately end the mission.
'******************************************************************************
Private Sub BurstInRadioRoom()
    ' B-24 and Lancaster radio operators were located on the flight deck, not
    ' in a separate compartment.
    
    ' Rule 19.2.d
    
    Dim strEffect As String

    ' Case 2:
            
    Call HeaterHit(RADIO_OPERATOR, strEffect)
            
    ' Case 3:
            
    If Damage.IntercomSystem = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        Damage.IntercomSystem = True
        strEffect = "Intercom destroyed."
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
        
    ' Case 4 To 5:
            
    If Damage.Radio = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        Damage.Radio = True
        strEffect = "Radio destroyed."
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
        
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
    ' Case 6:
            
    Call KillAirman(RADIO_OPERATOR)
        
    ' Case 7 To 10:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 20
        
    ' Case 11:
            
    Call OxygenHit(RADIO_OPERATOR, 2, strEffect)
    
    ' Case 12:
            
    Call ControlCableHit(strEffect)
    Call ControlCableHit(strEffect)
    UpdateMessage "Both control cables severed."

End Sub

'******************************************************************************
' BurstInWaist
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is bad, but not bad enough to imediately end the mission.
'******************************************************************************
Private Sub BurstInWaist()

    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.BomberModel = B17_F _
    Or Bomber.BomberModel = B17_G _
    Or Bomber.BomberModel = YB40 Then
    
        Call BurstInWaistB17
    
    ElseIf Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
    
        Call BurstInWaistB24
    
    Else
    
        Call BurstInWaistLanc
    
    End If

End Sub

'******************************************************************************
' BurstInWaistB17
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is bad, but not bad enough to imediately end the mission.
'******************************************************************************
Private Sub BurstInWaistB17()
    
    ' Rule 19.2.d
    
    Dim strEffect As String

    ' Case 2:
            
    Call OxygenHit(PORT_WAIST_GUNNER, 2, strEffect)
    Call OxygenHit(STBD_WAIST_GUNNER, 2, strEffect)
    
    ' Case 3:
            
    If Bomber.Gun(PORT_WAIST_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Port waist MG destroyed."
        Bomber.Gun(PORT_WAIST_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
                
    If Bomber.Gun(STBD_WAIST_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Starboard waist MG destroyed."
        Bomber.Gun(STBD_WAIST_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
                
    ' Case 4:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
    ' Case 5:
            
    If Bomber.BomberModel <> YB40 Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
            
        If Bomber.Gun(MID_UPPER_MG).Status = MG_INOPERABLE Then
            ' Superficial damage
            Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        Else
            UpdateMessage "Mid-upper turret destroyed."
            Bomber.Gun(MID_UPPER_MG).Status = MG_INOPERABLE
            Damage.PeckhamPoints = Damage.PeckhamPoints + 20
        End If
            
        Call KillAirman(MID_UPPER_GUNNER)
        Call HeaterHit(MID_UPPER_GUNNER, strEffect)
        Call OxygenHit(MID_UPPER_GUNNER, 2, strEffect)
            
    End If
        
    ' Case 6:
            
    Call KillAirman(PORT_WAIST_GUNNER)
            
    ' Case 7:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
    ' Case 8:
            
    Call KillAirman(STBD_WAIST_GUNNER)
        
    ' Case 9:
            
    If Damage.BallTurretMech = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        Damage.BallTurretMech = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    End If
        
    If Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Ball turret destroyed."
        Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
    End If
        
    Call KillAirman(BALL_GUNNER)
    Call HeaterHit(BALL_GUNNER, strEffect)
    Call OxygenHit(BALL_GUNNER, 2, strEffect)
                    
    ' Case 10:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
            
    ' Case 11:
        
    Call HeaterHit(PORT_WAIST_GUNNER, strEffect)
    Call HeaterHit(STBD_WAIST_GUNNER, strEffect)
            
    ' Case 12:
            
    Call ControlCableHit(strEffect)
    Call ControlCableHit(strEffect)
    UpdateMessage "Both control cables severed."

End Sub

'******************************************************************************
' BurstInWaistB24
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is bad, but not bad enough to imediately end the mission.
'******************************************************************************
Private Sub BurstInWaistB24()
    
    ' Rule 19.2.d
    
    Dim strEffect As String

    ' Case 2:
            
    Call OxygenHit(PORT_WAIST_GUNNER, 2, strEffect)
    ' Assume a separate oxygen system from the port gunner, so
    ' that some other crew member may man the stbd position.
    Call OxygenHit(STBD_WAIST_GUNNER, 2, strEffect)
    
    ' Case 3:
            
    If Bomber.Gun(PORT_WAIST_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Port waist MG destroyed."
        Bomber.Gun(PORT_WAIST_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
                
    If Bomber.Gun(STBD_WAIST_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Starboard waist MG destroyed."
        Bomber.Gun(STBD_WAIST_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
                
    ' Case 4 To 5:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 10
        
    ' Case 6:
            
    Call KillAirman(PORT_WAIST_GUNNER)
            
    ' Case 7:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
    ' Case 8:
            
    Call KillAirman(STBD_WAIST_GUNNER)
        
    ' Case 9:
            
    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Damage.BallTurretMech = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        Damage.BallTurretMech = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    End If
    
    If Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    ElseIf Bomber.BomberModel = B24_LM Then
        UpdateMessage "Floor ring MG destroyed."
        Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
    Else
        UpdateMessage "Ball turret destroyed."
        Bomber.Gun(BALL_TURRET_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
    End If

    Call KillAirman(BALL_GUNNER)
    Call HeaterHit(BALL_GUNNER, strEffect)
    Call OxygenHit(BALL_GUNNER, 2, strEffect)
            
    ' Case 10:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
    ' Case 11:
        
    Call HeaterHit(PORT_WAIST_GUNNER, strEffect)
    Call HeaterHit(STBD_WAIST_GUNNER, strEffect)
            
    ' Case 12:
            
    Call ControlCableHit(strEffect)
    Call ControlCableHit(strEffect)
    UpdateMessage "Both control cables severed."

End Sub

'******************************************************************************
' BurstInWaistLanc
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  This is bad, but not bad enough to imediately end the mission.
'******************************************************************************
Private Sub BurstInWaistLanc()
    
    ' Rule 19.2.d
    
    Dim strEffect As String

    ' Case 2:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
    ' Case 3:
            
    Call OxygenHit(MID_UPPER_GUNNER, 2, strEffect)
    
    ' Case 4:

    If Damage.PortAmmoBox = 2 Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Port ammo box destroyed."
        Damage.PortAmmoBox = 2
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
                
    If Damage.StbdAmmoBox = 2 Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Starboard ammo box destroyed."
        Damage.StbdAmmoBox = 2
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
                    
    ' Case 5 To 7:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 15
        
    ' Case 8:
            
    If Bomber.Gun(MID_UPPER_MG).Status = MG_INOPERABLE Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Mid-upper turret destroyed."
        Bomber.Gun(MID_UPPER_MG).Status = MG_INOPERABLE
        Damage.PeckhamPoints = Damage.PeckhamPoints + 20
    End If
                
    Call KillAirman(MID_UPPER_GUNNER)
        
    ' Case 9:
            
    ' Superficial damage
    Damage.PeckhamPoints = Damage.PeckhamPoints + 5
        
    ' Case 10:
            
    If Damage.PortAmmoTrack = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Port ammo track destroyed."
        Damage.PortAmmoTrack = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If

    If Damage.StbdAmmoTrack = True Then
        ' Superficial damage
        Damage.PeckhamPoints = Damage.PeckhamPoints + 5
    Else
        UpdateMessage "Starboard ammo track destroyed."
        Damage.StbdAmmoTrack = True
        Damage.PeckhamPoints = Damage.PeckhamPoints + 10
    End If
                    
    ' Case 11:
        
    Call HeaterHit(MID_UPPER_GUNNER, strEffect)
                
    ' Case 12:
            
    Call ControlCableHit(strEffect)
    Call ControlCableHit(strEffect)
    UpdateMessage "Both control cables severed."

End Sub

'******************************************************************************
' BurstInTail
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission.
'
' NOTES:  n/a
'******************************************************************************
Public Function BurstInTail() As Integer

    ' Rule 19.2.b
    
    UpdateMessage "Flak burst blows off tail!"
    
    If Bomber.BomberModel <> B17_C Then
        Call KillAirman(TAIL_GUNNER)
    End If
    
    UpdateMessage "Bomber dives out of control."
    
    G7UncontrolledBailout (OverWater())
    BurstInTail = END_MISSION

End Function

'******************************************************************************
' BurstInWing
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission.
'
' NOTES:  n/a
'******************************************************************************
Public Function BurstInWing(ByVal intSide As Integer) As Integer
            
    ' Rule 19.2.b
    
    UpdateMessage "Flak burst blows off " & SideText(intSide) & " wing!"
    
    UpdateMessage "Bomber dives out of control."
    
    G7UncontrolledBailout (OverWater())
    BurstInWing = END_MISSION

End Function

'******************************************************************************
' DropOutOfFormation
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: n/a
'
' NOTES:  There is no provision for re-joining a formation, as the formation
'         would not throttle down for stragglers.
'******************************************************************************
Public Sub DropOutOfFormation()
    
    ' Lancasters are part of a "stream", rather than a formation, so they
    ' can't fall out of formation.
    
    If Bomber.BomberModel <> AVRO_LANCASTER _
    And Bomber.InFormation = True Then
        
        Bomber.InFormation = False

        Bomber.FormationPos = MIDDLE_PLANE
        Bomber.SquadronPos = MIDDLE_SQUADRON

        UpdateMessage "Bomber drops out of formation."
    
    End If

End Sub

' NOTES:  Bomber must drop out of formation before it can lose altitude. The
'         DropToLowAltitude() asks if the user wants to descend to low altitude;
'         this function actually does the descent.
Public Sub LoseAltitude()
    'We won't worry about the formation dropout, that's the job of DropOutOfFormation()
    If Bomber.Altitude = HIGH_ALTITUDE Then
        'don't want to report that we were descending when we were already low.
        Bomber.Altitude = LOW_ALTITUDE
    
        UpdateMessage "Bomber descends to low altitude (" & _
                      LOW_ALTITUDE & " feet)."
    End If
End Sub

Public Sub GainAltitude()
    
    If Bomber.Altitude = LOW_ALTITUDE Then
        Bomber.Altitude = HIGH_ALTITUDE
        UpdateMessage "Bomber climbs to high altitude (" & _
              HIGH_ALTITUDE & " feet)."

    End If
End Sub

'******************************************************************************
' B1NumberOfGermanFighterWaves
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The number of waves in the current zone.
'
' NOTES:  n/a
'******************************************************************************
Public Function B1NumberOfGermanFighterWaves() As Integer
    Dim intRoll As Integer
    Dim intWaves As Integer

    ' B-1 Number of German Fighter Waves in non-designated target zone
    ' B-2 Number of German Fighter Waves in designated target zone
    
    B1NumberOfGermanFighterWaves = 0
    
    intRoll = Random1D6()
    
    If AlpsDirection() = ALPS_BELOW Then

        If (Mission.Zone(Bomber.CurrentZone).Weather = CLEAR_WEATHER _
        Or Mission.Zone(Bomber.CurrentZone).Weather = GOOD_WEATHER) _
        And intRoll = 6 Then
        
            ' Only one German wave may appear over Alps.
            B1NumberOfGermanFighterWaves = 1
        
        ElseIf Mission.Zone(Bomber.CurrentZone).Weather = POOR_WEATHER Then
    
            ' No German fighters appear in fog over the Alps. If the weather
            ' over the Alps was bad or storm, then the bomber should have
            ' aborted before entering the Alps.
            B1NumberOfGermanFighterWaves = 0
        
        End If
    
        Exit Function
    
    End If
    
    ' O-1 Weather: Note a and Note b, plus "The General" (Volume 24, #6)
    ' variant.
        
    If Mission.Zone(Bomber.CurrentZone).Weather = CLEAR_WEATHER Then
        intRoll = intRoll + 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = POOR_WEATHER Then
        intRoll = intRoll - 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = BAD_WEATHER Then
        intRoll = intRoll - 2
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = STORM_WEATHER Then
        intRoll = intRoll - 3
    End If
    
    If Mission.Zone(Bomber.CurrentZone).Contrail = True Then
        intRoll = intRoll + 1
    End If
    
    ' Rule 18.0
    
    If RandomEvent.LooseFormation Then
        intRoll = intRoll + 1
    ElseIf RandomEvent.TightFormation = True Then
        intRoll = intRoll - 1
    End If
    
    ' G-11 Flight Log Gazeteer.

    intRoll = intRoll + Mission.Zone(Bomber.CurrentZone).Modifier

    ' Mission Chart.
    
    If Bomber.SquadronPos = LOW_SQUADRON Then
        intRoll = intRoll + 1
    ElseIf Bomber.SquadronPos = MIDDLE_SQUADRON Then
        intRoll = intRoll - 1
    End If

    If Bomber.CurrentZone = Mission.TargetZone Then
        
        Select Case intRoll
            Case Is <= 3: intWaves = 1
            Case 4 To 5:  intWaves = 2
            Case Is >= 6: intWaves = 3
        End Select
    
    Else
        
        Select Case intRoll
            Case Is <= 2: intWaves = 0
            Case 3 To 5:  intWaves = 1
            Case Is >= 6: intWaves = 2
        End Select
    
    End If

    UpdateMessage intWaves & " German fighter waves"
    
    B1NumberOfGermanFighterWaves = intWaves

End Function

'******************************************************************************
' B1TameBoarWave
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The number of waves in the current zone (always 0 or 1).
'
' NOTES:  n/a
'******************************************************************************
Public Function B1TameBoarWave() As Integer
    ' Lancaster variant.
    Dim intRoll As Integer
    
    B1TameBoarWave = 0

    If Bomber.CurrentZone <= 3 Then
        ' Tame Boar fighters were not encountered in zones 1 to 3.
        Exit Function
    End If

    intRoll = Random1D6()
    
    ' O-1 Weather: Note a and Note b, plus "The General" (Volume 24, #6)
    ' variant.
        
    If Mission.Zone(Bomber.CurrentZone).Weather = CLEAR_WEATHER Then
        intRoll = intRoll + 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = POOR_WEATHER Then
        intRoll = intRoll - 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = BAD_WEATHER Then
        intRoll = intRoll - 2
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = STORM_WEATHER Then
        intRoll = intRoll - 3
    End If
    
    ' G-11 Flight Log Gazeteer. (Ignored by Lancasters.)
    
    ' Mission Chart. (Ignored by Lancasters.)
    
    If Bomber.Direction = OUTBOUND Then
    
        If (Bomber.CurrentZone = Mission.TargetZone And intRoll >= 3) _
        Or (Bomber.CurrentZone = 4 And intRoll >= 6) _
        Or (Bomber.CurrentZone = 5 And intRoll >= 5) _
        Or (Bomber.CurrentZone >= 6 And intRoll >= 4) Then
            B1TameBoarWave = 1
        End If
    
    Else ' Returning to base
    
        If (Bomber.CurrentZone = Mission.TargetZone And intRoll >= 5) _
        Or intRoll >= 6 Then
            B1TameBoarWave = 1
        End If
    
    End If
    
    UpdateMessage B1TameBoarWave & " Tame Boar waves"
    
End Function

'******************************************************************************
' B1WildBoarWave
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The number of waves in the current zone (always 0 or 1).
'
' NOTES:  n/a
'******************************************************************************
Public Function B1WildBoarWave() As Integer
    ' Lancaster variant.

    B1WildBoarWave = 0

    ' Wild Boar fighters are only encountered over the target, after flak,
    ' but before the bomb run.
    
    If Bomber.SpottedBySearchLight = True _
    Or Random1D6() = 6 Then
        B1WildBoarWave = 1
    End If

    UpdateMessage B1WildBoarWave & " Wild Boar waves"
    
End Function

'******************************************************************************
' SpottedBySearchLight
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if spotted, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function SpottedBySearchLight() As Boolean
    Dim intRoll As Integer

    SpottedBySearchLight = False

    intRoll = Random1D6()
    
    ' O-1 Weather: Note a and Note b, plus "The General" (Volume 24, #6)
    ' and Lancaster variants.
        
    If Mission.Zone(Bomber.CurrentZone).Weather = CLEAR_WEATHER Then
        intRoll = intRoll + 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = POOR_WEATHER Then
        intRoll = intRoll - 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = BAD_WEATHER Then
        intRoll = intRoll - 2
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = STORM_WEATHER Then
        intRoll = intRoll - 3
    End If
    
    If intRoll >= 5 Then
        SpottedBySearchLight = True
    End If

End Function

'******************************************************************************
' FriendlyFire
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function FriendlyFire() As Integer
    Dim intRoll As Integer
    Dim intHits As Integer
    Dim intCount As Integer
    
    UpdateMessage "Bomber hit by friendly fire"
    
    If Random1D6() = 6 Then
    
        intRoll = Random2D6()
        
        If intRoll = 2 _
        Or intRoll = 12 Then
            intHits = 2
        Else
            intHits = 1
        End If
    
        For intCount = 1 To intHits

            Select Case Random2D6()
                Case 2, 12: FriendlyFire = P1NoseDamage()
                Case 3, 11: FriendlyFire = P2FlightDeckDamage()
                Case 4, 10: FriendlyFire = P3BombBayDamage()
                Case 5:
                    If Bomber.BomberModel = B24_D _
                    Or Bomber.BomberModel = B24_E _
                    Or Bomber.BomberModel = B24_GHJ _
                    Or Bomber.BomberModel = B24_LM Then
                        FriendlyFire = P1NoseDamage()
                    Else
                        FriendlyFire = P4RadioRoomDamage()
                    End If
                Case 6: FriendlyFire = BL1WingDamage(PORT_SIDE)
                Case 7: FriendlyFire = P5WaistDamage()
                Case 8: FriendlyFire = BL1WingDamage(STBD_SIDE)
                Case 9: FriendlyFire = P6TailDamage()
            End Select
    
        Next intCount
    
    End If
    
End Function

'******************************************************************************
' B3AttackingFighterWave
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if random events caused catastrophic damage,
'         otherwise the number of fighters in the wave.
'
' NOTES:  n/a
'******************************************************************************
Public Function B3AttackingFighterWave() As Integer
    Dim intWaveRoll As Integer
    Dim intFighterIndex As Integer
    Dim intAcePilots As Integer
    Dim intAceOfAces As Integer
    Dim intFighters As Integer
    Dim rsFighterWave As New ADODB.Recordset

    Wave.JG26 = False
    Wave.Ju88 = False
    Wave.Attack = 0 ' Wave has yet to attack

ReRoll:
    
    intWaveRoll = RandomD66()

' intWaveRoll = 11
'UpdateMessage "intWaveRoll = " & intWaveRoll
If intWaveRoll >= 36 Then
'    intWaveRoll = 25 '63
'    UpdateMessage "*** random events ***"
Else
'    intWaveRoll = 24 '14
End If

    If AlpsDirection() = ALPS_BELOW _
    And (intWaveRoll = 15 _
    Or intWaveRoll = 21 _
    Or intWaveRoll = 31 _
    Or intWaveRoll = 34 _
    Or intWaveRoll = 61 _
    Or intWaveRoll = 65) Then
        ' Vertical dive and vertical climb attacks are not allowed in the Alps.
        GoTo ReRoll
    End If

    Select Case intWaveRoll
        
        ' Handle "non-wave" exceptions
        
        Case 16, 26, 36, 46, 56:
            
            If Bomber.InFormation = False Then
                GoTo ReRoll
            End If
            
            If Mission.Options.TimePeriodSpecificFormations = True Then
                ' "The General" (Volume 24, #6) variant.
                If Mission.Date = AUG_1942 Then
                
                    GoTo ReRoll
                
                ElseIf Mission.Date >= SEP_1942 _
                And Mission.Date <= APR_1943 Then
                
                    If intWaveRoll = 16 _
                    Or intWaveRoll = 36 _
                    Or intWaveRoll = 56 Then
                        GoTo ReRoll
                    End If
                End If
            End If
            
            UpdateMessage "Bandits driven off by other bombers."
            
            If Mission.Options.FormationDefensiveGunnery = True Then
                ' "The General" (Volume 24, #6) variant.
                If FriendlyFire() = END_MISSION Then
                    B3AttackingFighterWave = END_MISSION
                    Exit Function
                End If
            End If
            
        Case 66:
            
            ' Rule 18.0
        
            If Mission.Options.RandomEvents = True Then
                If B7RandomEvents() = END_MISSION Then
                    B3AttackingFighterWave = END_MISSION
                    Exit Function
                End If
            Else
                frmMission.lblMiscWave.Caption = "No Attackers"
                UpdateMessage "No attackers."
            End If
        
        ' Handle normal waves
        
        Case Else:
            
            ' Initialize variables for the current wave.
            
            intFighters = 0
            intAceOfAces = 0
            intAcePilots = 0
            
            ' Determine if this is a special wave.

            Wave.JG26 = WaveIsJG26(intWaveRoll)
            
            If Wave.JG26 = False Then
                ' A wave can't be both JG26 and Ju88. Only check to see if it
                ' Ju88s if it is not JG26.
                Wave.Ju88 = WaveIsJu88(intWaveRoll)
'UpdateMessage "Wave.Ju88 = " & Wave.Ju88
            End If
            
            If Wave.Ju88 = True Then
                ' Ju88s have their own rows in the WaveSelection table.
                intWaveRoll = intWaveRoll + 100
            End If
            
            If GetFighterWaveRecordset(intWaveRoll, rsFighterWave) = False Then
                ' Try to continue despite the error.
                frmMission.lblMiscWave.Caption = "No Attackers" ' ???
                UpdateMessage "Bandits driven off by other bombers."
            End If
            
            If Wave.Ju88 = True _
            Or Bomber.FormationPos = MIDDLE_PLANE _
            Or Bomber.InFormation = False Then ' True Then
                ' No additional opposition: The current wave has the
                ' number of fighters found for the roll.
                intFighters = rsFighterWave.RecordCount
            Else
                ' Lead and tail bombers saw extra opposition, as did bombers
                ' that were out of formation. The extra fighter will be set
                ' to a Me-109 further down.
                intFighters = rsFighterWave.RecordCount + 1
            End If
            
            B3AttackingFighterWave = intFighters
            
            ' Configure the new wave structure, overwriting any previous data
            ' in the structure.
            
            ReDim Wave.Fighter(1 To intFighters)
            
            rsFighterWave.MoveFirst
            
            ' Configure each fighter in the wave.
            For intFighterIndex = 1 To intFighters
                
                ' Determine fighter pilot skill.
            
                If Wave.JG26 = True Then
                    
                    ' JG-26 had more top pilots than other German fighter
                    ' groups.
                    
                    Wave.Fighter(intFighterIndex).PilotSkill = JG26FighterPilotSkill()
                
                    If Wave.Fighter(intFighterIndex).PilotSkill = ACE_OF_ACES Then
                        
                        intAceOfAces = intAceOfAces + 1
                        
                        ' However, even JG-26 only had so many 100+ kill
                        ' pilots.
                        
                        If intAceOfAces >= 2 Then
                            Wave.Fighter(intFighterIndex).PilotSkill = VET_PILOT
                            intAceOfAces = intAceOfAces - 1
                        End If
                    
                    ElseIf Wave.Fighter(intFighterIndex).PilotSkill = ACE_PILOT Then

                        intAcePilots = intAcePilots + 1
                        
                        ' There is also a limit of the number of aces,
                        ' because even Galland got shot down a few times.
                        
                        If intAcePilots >= 3 Then
                            Wave.Fighter(intFighterIndex).PilotSkill = VET_PILOT
                            intAcePilots = intAcePilots - 1
                        End If
                    
                    End If
                
                ElseIf Wave.Ju88 = True Then
                
                    ' Only veteran and ace pilots fly the Ju-88.
                    
                    Wave.Fighter(intFighterIndex).PilotSkill = M6FighterPilotStatus()
                    
                    If Wave.Fighter(intFighterIndex).PilotSkill = GREEN_PILOT Then
                        Wave.Fighter(intFighterIndex).PilotSkill = VET_PILOT
                    End If
                
                Else

                    ' Normal German fighter group.
                    
                    Wave.Fighter(intFighterIndex).PilotSkill = M6FighterPilotStatus()
                
                End If
                
                ' Determine fighter type and position.
            
                If intFighterIndex <> intFighters _
                Or Wave.Ju88 = True Then
                    
                    ' Assign table type and position.
                    
                    If Bomber.BomberModel = B17_C _
                    And rsFighterWave![Type] = "FW190" Then
                        ' FW-190s were not in widespread use when the B-17 was.
                        ' So, change the 190s to 109s. Note that if the last
                        ' fighter in the wave is a 190, it will remain a 190.
                        Wave.Fighter(intFighterIndex).Type = "Me109"
                    Else
                        Wave.Fighter(intFighterIndex).Type = rsFighterWave![Type]
                    End If
                    
                    Wave.Fighter(intFighterIndex).Position = rsFighterWave![Position]
                    
                    If IsNull(rsFighterWave![Special]) = False Then
                        Wave.Fighter(intFighterIndex).Special = rsFighterWave![Special]
                    End If
                
                    rsFighterWave.MoveNext
            
                Else
                    
                    ' This is the last fighter in the list. Determine if
                    ' it is the "bonus" fighter.
                
                    If Bomber.FormationPos = LEAD_PLANE Then
                        
                        ' "Bonus" fighter.
                        
                        Wave.Fighter(intFighterIndex).Type = "Me109"
                        Wave.Fighter(intFighterIndex).Position = F12_HIGH
                        Wave.Fighter(intFighterIndex).Special = ""
                    
                    ElseIf Bomber.FormationPos = TAIL_PLANE Then
                        
                        ' "Bonus" fighter.
                        
                        Wave.Fighter(intFighterIndex).Type = "Me109"
                        Wave.Fighter(intFighterIndex).Position = F6_HIGH
                        Wave.Fighter(intFighterIndex).Special = ""
                    
                    ElseIf Bomber.InFormation = False Then
            
                        ' "Bonus" fighter.
                        
                        Wave.Fighter(intFighterIndex).Type = "Me109"
                        Wave.Fighter(intFighterIndex).Position = F12_LEVEL
                        Wave.Fighter(intFighterIndex).Special = ""
            
                    Else
                        
                        ' Middle of formation. Assign table type and
                        ' position. (This should never occur, but put it
                        ' here as a catch.)
                        
                        Wave.Fighter(intFighterIndex).Type = rsFighterWave![Type]
                        Wave.Fighter(intFighterIndex).Position = rsFighterWave![Position]
                    
                        If IsNull(rsFighterWave![Special]) = False Then
                            Wave.Fighter(intFighterIndex).Special = rsFighterWave![Special]
                        End If
                    
                    End If
                
                End If
                
                Wave.Fighter(intFighterIndex).Damage = NO_DAMAGE
                Wave.Fighter(intFighterIndex).Status = ""
            
            Next intFighterIndex ' Configure next fighter in the wave
    
    End Select ' Done configuring fighters for the current wave
    
End Function

'******************************************************************************
' B3TameBoarFighter
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if random events caused catastrophic damage,
'         otherwise the number of fighters in the wave (always 0 or 1).
'
' NOTES:  n/a
'******************************************************************************
Public Function B3TameBoarFighter() As Integer
    Dim intWaveRoll As Integer
    
    intWaveRoll = RandomD66()

    Select Case intWaveRoll
    
        Case 66:
            
            ' Rule 18.0
        
            If Mission.Options.RandomEvents = True Then
                B3TameBoarFighter = B7RandomEvents()
            Else
                frmMission.lblMiscWave.Caption = "No Attackers"
                UpdateMessage "No attackers."
                B3TameBoarFighter = 0
            End If
        
        ' Handle normal waves
        
        Case 11 To 65:
    
            Wave.JG26 = False
            Wave.Ju88 = False
            Wave.Attack = 0 ' Wave has yet to attack
    
            ' Tame Boar waves always consist of exactly one fighter.
            
            ReDim Wave.Fighter(1)
                    
            ' Configure fighter.
            
            Wave.Fighter(1).Type = "Me110"
            
            If Random1D6() >= 3 Then
                ' Tame Boar has surprise.
                Wave.Fighter(1).Position = F6_LOW
            Else
                Wave.Fighter(1).Position = VERT_CLIMB
            End If
            
            Wave.Fighter(1).PilotSkill = M6FighterPilotStatus()
            Wave.Fighter(1).Damage = NO_DAMAGE
            Wave.Fighter(1).Status = ""
            Wave.Fighter(1).Special = "TameBoar"
            
            B3TameBoarFighter = 1
    
    End Select
    
End Function

'******************************************************************************
' B3WildBoarFighter
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: End of mission, if random events caused catastrophic damage,
'         otherwise the number of fighters in the wave (always 0 or 1).
'
' NOTES:  n/a
'******************************************************************************
Public Function B3WildBoarFighter() As Integer
    Dim intWaveRoll As Integer
    
    intWaveRoll = RandomD66()

    Select Case intWaveRoll
    
        Case 66:
            
            ' Rule 18.0
        
            If Mission.Options.RandomEvents = True Then
                B3WildBoarFighter = B7RandomEvents()
            Else
                frmMission.lblMiscWave.Caption = "No Attackers"
                UpdateMessage "No attackers."
                B3WildBoarFighter = 0
    
            End If
        
        ' Handle normal waves
        
        Case 11 To 65:
    
            Wave.JG26 = False
            Wave.Ju88 = False
            Wave.Attack = 0 ' Wave has yet to attack
        
            ' Wild Boar waves always consist of exactly one fighter.
            
            ReDim Wave.Fighter(1)
                    
            ' Configure fighter.
            
            Wave.Fighter(1).Type = "Me109"
            Wave.Fighter(1).Position = B6SuccessiveAttacks()
            Wave.Fighter(1).PilotSkill = M6FighterPilotStatus()
            Wave.Fighter(1).Damage = NO_DAMAGE
            Wave.Fighter(1).Status = ""
            Wave.Fighter(1).Special = "WildBoar"
    
            B3WildBoarFighter = 1
    
    End Select

End Function

'******************************************************************************
' JG26FighterPilotSkill
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: A JG-26 pilot's skill level.
'
' NOTES:  n/a
'******************************************************************************
Private Function JG26FighterPilotSkill() As Integer
    Dim intRoll As Integer
    
    JG26FighterPilotSkill = VET_PILOT
        
    intRoll = Random2D6()

    Select Case intRoll
        Case 2 To 4:
            JG26FighterPilotSkill = ACE_OF_ACES
        Case 5, 10:
            JG26FighterPilotSkill = ACE_PILOT
        Case 6 To 9, 11 To 12:
            JG26FighterPilotSkill = VET_PILOT
    End Select
    
End Function

'******************************************************************************
' WaveIsJG26
'
' INPUT:  The type of wave.
'
' OUTPUT: n/a
'
' RETURN: True is the wave is from JG-26, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Private Function WaveIsJG26(ByVal intWaveRoll As Integer) As Boolean
    Dim intRoll As Integer
    
    WaveIsJG26 = False

    ' The if statements in this function were broken up to keep the
    ' logic simple.
        
    If Mission.Options.JG26StationedInAbbeville = False Then
        Exit Function
    End If

    ' The JG26 variant was chosen for this mission.
    
    If intWaveRoll = 12 _
    Or intWaveRoll = 14 _
    Or intWaveRoll = 24 _
    Or intWaveRoll = 33 _
    Or intWaveRoll = 42 _
    Or intWaveRoll = 54 _
    Or intWaveRoll = 65 Then

        ' Although JG-26 did fly other fighters, in this variant they only
        ' fly Me-109s. Therefore, only the "all Me-109" waves get to this
        ' point.
    
        intRoll = Random1D6()
        
        If intRoll <= 3 Then

            If Mission.Zone(Bomber.CurrentZone).Terrain = NETHERLANDS_TER _
            Or Mission.Zone(Bomber.CurrentZone).Terrain = NETHERLANDS_GERMANY_TER _
            Or Mission.Zone(Bomber.CurrentZone).Terrain = WATER_NETHERLANDS_TER _
            Or Mission.Zone(Bomber.CurrentZone).Terrain = BELGIUM_TER _
            Or Mission.Zone(Bomber.CurrentZone).Terrain = BELGIUM_GERMANY_TER Then
                ' Bomber is either attacking a target in the Low Countries,
                ' or it is en route to Germany.
                WaveIsJG26 = True
            End If
        
            If Mission.Zone(Bomber.CurrentZone).Terrain = FRANCE_TER _
            Or Mission.Zone(Bomber.CurrentZone).Terrain = WATER_FRANCE_TER Then
                If Mission.TargetName = "Abbeville" _
                Or Mission.TargetName = "Amiens" _
                Or Mission.TargetName = "Lille" _
                Or Mission.TargetName = "Meaulte" _
                Or Mission.TargetName = "Romilly-S.S." _
                Or Mission.TargetName = "St. Omer" _
                Or Mission.TargetName = "Paris" _
                Or Mission.TargetName = "Reims" Then
                    ' Bomber is attacking, or en route to, a target in
                    ' northern France.
                    WaveIsJG26 = True
                End If
            End If
    
        End If ' variant in effect
    
    End If ' intWaveRoll

    If WaveIsJG26 = True Then
        UpdateMessage "It's JG-26, the Abbeville Kids!"
    End If

End Function

'******************************************************************************
' WaveIsJu88
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: True if the wave consists of Ju-88 heavy fighters, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Private Function WaveIsJu88(ByVal intWaveRoll As Integer) As Boolean
    Dim intRoll As Integer
    
    WaveIsJu88 = False

    ' The if statements in this function were broken up to keep the
    ' logic simple.
        
    If Mission.Options.Ju88sUsedAsFighters = False Then
        Exit Function
    End If

    ' The Ju-88 variant was chosen for this mission.
    
    If intWaveRoll = 11 _
    Or intWaveRoll = 22 _
    Or intWaveRoll = 32 _
    Or intWaveRoll = 41 _
    Or intWaveRoll = 51 _
    Or intWaveRoll = 62 Then
        
        ' The Ju-88 only appears in zones with no fighter cover.
        
        If Mission.Zone(Bomber.CurrentZone).CoverOut = NO_COVER _
        And Bomber.Direction = OUTBOUND Then
        
            WaveIsJu88 = True
        
        ElseIf Mission.Zone(Bomber.CurrentZone).CoverBack = NO_COVER _
        And Bomber.Direction = RETURN_TRIP Then
        
            WaveIsJu88 = True
        
        End If
    
    End If
    
    If WaveIsJu88 = True Then
        UpdateMessage "Bomber being attacked by Ju-88s."
    End If

End Function

'******************************************************************************
' B4ShellHitsByArea
'
' INPUT:  The direction from which the fighter is attacking, and the type of
'         fighter it is.
'
' OUTPUT: n/a
'
' RETURN: The number of hits.
'
' NOTES:  n/a
'******************************************************************************
Public Function B4ShellHitsByArea(ByVal intFighterPos As Integer, ByVal strFighterType As String) As Integer
    Dim intRoll As Integer
    Dim intHits As Integer
    
    B4ShellHitsByArea = 0

    intRoll = Random2D6()
    
    Select Case intFighterPos
        
        Case F12_HIGH To F12_LOW, F130_HIGH To F130_LOW, F1030_HIGH To F1030_LOW:
            
            Select Case intRoll
                Case 6, 7, 8: intHits = 1
                Case 3, 4, 5, 9, 10, 11: intHits = 2
                Case 2: intHits = 3
                Case 12: intHits = 4
            End Select
        
        Case F3_HIGH To F3_LOW, F9_HIGH To F9_LOW
            
            Select Case intRoll
                Case 7: intHits = 1
                Case 6, 8: intHits = 2
                Case 3, 4, 5, 9, 10, 11: intHits = 3
                Case 2: intHits = 4
                Case 12: intHits = 5
            End Select
        
        Case F6_HIGH To F6_LOW
            
            Select Case intRoll
                Case 6, 7, 8: intHits = 2
                Case 5, 9: intHits = 3
                Case 4, 10: intHits = 4
                Case 3, 11: intHits = 5
                Case 2: intHits = 6
                Case 12: intHits = 7
            End Select
        
        Case VERT_CLIMB
            
            Select Case intRoll
                Case 7: intHits = 1
                Case 5, 6, 8, 9: intHits = 2
                Case 4, 10: intHits = 3
                Case 2, 3, 11: intHits = 4
                Case 12: intHits = 5
            End Select
        
        Case VERT_DIVE
            
            Select Case intRoll
                Case 5, 6, 7, 8, 9: intHits = 1
                Case 3, 4, 10, 11: intHits = 2
                Case 2: intHits = 3
                Case 12: intHits = 4
            End Select
    
    End Select

    If strFighterType = "Me110" _
    Or strFighterType = "Ju88" Then
        intHits = intHits + 1
    ElseIf strFighterType = "FW190" Then
        intHits = Int(intHits * 1.5) ' round down
    End If

    B4ShellHitsByArea = intHits

End Function

'******************************************************************************
' B5AreaDamage
'
' INPUT:  The direction from which the fighter is attacking and the affected
'         side, if any.
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Public Function B5AreaDamage(ByVal intPosition As Integer, Optional ByVal intSide As Integer) As Integer
    
    B5AreaDamage = NO_EFFECT_HIT
    
    If Bomber.BomberModel = B17_C _
    Or Bomber.BomberModel = B17_E _
    Or Bomber.BomberModel = B17_F _
    Or Bomber.BomberModel = B17_G _
    Or Bomber.BomberModel = YB40 Then
    
        B5AreaDamage = B5AreaDamageB17(intPosition, intSide)
    
    ElseIf Bomber.BomberModel = B24_D _
    Or Bomber.BomberModel = B24_E _
    Or Bomber.BomberModel = B24_GHJ _
    Or Bomber.BomberModel = B24_LM Then
    
        B5AreaDamage = B5AreaDamageB24(intPosition, intSide)
    
    Else
    
        B5AreaDamage = B5AreaDamageLanc(intPosition)
    
    End If

End Function

'******************************************************************************
' B5AreaDamageLanc
'
' INPUT:  The direction from which the fighter is attacking.
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function B5AreaDamageLanc(ByVal intPosition As Integer) As Integer
    Dim intRoll As Integer

    B5AreaDamageLanc = NO_EFFECT_HIT
    
    ' There is no "side" for Lancaster attacks, nor is the attacker's clock
    ' position taken into account.
                        
    intRoll = Random2D6()

    Select Case intRoll
        Case 2: B5AreaDamageLanc = NOSE_HIT
        Case 3: B5AreaDamageLanc = BOMB_BAY_HIT
        Case 4: B5AreaDamageLanc = NO_EFFECT_HIT
        Case 5: B5AreaDamageLanc = PORT_WING_HIT
        Case 6: B5AreaDamageLanc = TAIL_HIT
        Case 7:
            If Random1D6() <= 3 Then
                B5AreaDamageLanc = PORT_WING_HIT
            Else
                B5AreaDamageLanc = STBD_WING_HIT
            End If
        Case 8: B5AreaDamageLanc = WAIST_HIT
        Case 9: B5AreaDamageLanc = STBD_WING_HIT
        Case 10: B5AreaDamageLanc = NO_EFFECT_HIT
        Case 11: B5AreaDamageLanc = FLIGHT_DECK_HIT
        Case 12: B5AreaDamageLanc = BOMB_BAY_HIT
    End Select

End Function

'******************************************************************************
' B5AreaDamageB17
'
' INPUT:  The direction from which the fighter is attacking.
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function B5AreaDamageB17(ByVal intPosition As Integer, ByVal intSide As Integer) As Integer
    Dim intRoll As Integer
    
    B5AreaDamageB17 = NO_EFFECT_HIT
    
    ' Determine the area that was hit, if any. Damage will be determined after
    ' all hit areas are queued, so that if there is a walking hit we will be
    ' able to delete all the non-walking hits.
    
    intRoll = Random2D6()
    
    Select Case intPosition
        
        Case F12_HIGH, F130_HIGH, F1030_HIGH:
            
            Select Case intRoll
                Case 2 To 4: B5AreaDamageB17 = NO_EFFECT_HIT
                Case 5: B5AreaDamageB17 = RADIO_ROOM_HIT
                Case 6: B5AreaDamageB17 = NOSE_HIT
                Case 7: B5AreaDamageB17 = FLIGHT_DECK_HIT
                Case 8:
                    If Random1D6() <= 3 Then
                        B5AreaDamageB17 = PORT_WING_HIT
                    Else
                        B5AreaDamageB17 = STBD_WING_HIT
                    End If
                Case 9: B5AreaDamageB17 = WAIST_HIT
                Case 10: B5AreaDamageB17 = TAIL_HIT
                Case 11: B5AreaDamageB17 = BOMB_BAY_HIT
                Case 12: B5AreaDamageB17 = WALKING_HITS_FUSELAGE
            End Select
        
        Case F12_LEVEL, F130_LEVEL, F1030_LEVEL:
            
            Select Case intRoll
                Case 2, 6: B5AreaDamageB17 = PORT_WING_HIT
                Case 3 To 5: B5AreaDamageB17 = NO_EFFECT_HIT
                Case 7: B5AreaDamageB17 = NOSE_HIT
                Case 8, 12: B5AreaDamageB17 = STBD_WING_HIT
                Case 9: B5AreaDamageB17 = FLIGHT_DECK_HIT
                Case 10 To 11: B5AreaDamageB17 = NO_EFFECT_HIT
            End Select
        
        Case F12_LOW, F130_LOW, F1030_LOW:
            
            Select Case intRoll
                Case 2 To 4: B5AreaDamageB17 = NO_EFFECT_HIT
                Case 5: B5AreaDamageB17 = RADIO_ROOM_HIT
                Case 6: B5AreaDamageB17 = NOSE_HIT
                Case 7: B5AreaDamageB17 = BOMB_BAY_HIT
                Case 8:
                    If Random1D6() <= 3 Then
                        B5AreaDamageB17 = PORT_WING_HIT
                    Else
                        B5AreaDamageB17 = STBD_WING_HIT
                    End If
                Case 9: B5AreaDamageB17 = WAIST_HIT
                Case 10: B5AreaDamageB17 = TAIL_HIT
                Case 11: B5AreaDamageB17 = FLIGHT_DECK_HIT
                Case 12: B5AreaDamageB17 = WALKING_HITS_FUSELAGE
            End Select
        
        Case F3_HIGH, F9_HIGH:
            
            Select Case intRoll
                Case 2: B5AreaDamageB17 = WALKING_HITS_WINGS
                Case 3: B5AreaDamageB17 = NOSE_HIT
                Case 4: B5AreaDamageB17 = FLIGHT_DECK_HIT
                Case 5: B5AreaDamageB17 = BOMB_BAY_HIT
                Case 6: B5AreaDamageB17 = PORT_WING_HIT
                Case 7: B5AreaDamageB17 = TAIL_HIT
                Case 8: B5AreaDamageB17 = STBD_WING_HIT
                Case 9: B5AreaDamageB17 = RADIO_ROOM_HIT
                Case 10: B5AreaDamageB17 = WAIST_HIT
                Case 11: B5AreaDamageB17 = NO_EFFECT_HIT
                Case 12: B5AreaDamageB17 = WALKING_HITS_FUSELAGE
            End Select
        
        Case F3_LEVEL, F9_LEVEL:
            
            Select Case intRoll
                Case 2: B5AreaDamageB17 = BOMB_BAY_HIT
                Case 3, 6:
                    If intSide = PORT_SIDE Then
                        B5AreaDamageB17 = PORT_WING_HIT
                    Else
                        B5AreaDamageB17 = STBD_WING_HIT
                    End If
                Case 4: B5AreaDamageB17 = NOSE_HIT
                Case 5: B5AreaDamageB17 = WAIST_HIT
                Case 7 To 8, 12: B5AreaDamageB17 = NO_EFFECT_HIT
                Case 9: B5AreaDamageB17 = TAIL_HIT
                Case 10: B5AreaDamageB17 = RADIO_ROOM_HIT
                Case 11: B5AreaDamageB17 = FLIGHT_DECK_HIT
                Case 12: B5AreaDamageB17 = WALKING_HITS_BOTH
            End Select
        
        Case F3_LOW, F9_LOW:
            
            Select Case intRoll
                Case 2: B5AreaDamageB17 = WALKING_HITS_WINGS
                Case 3: B5AreaDamageB17 = NOSE_HIT
                Case 4: B5AreaDamageB17 = BOMB_BAY_HIT
                Case 5: B5AreaDamageB17 = RADIO_ROOM_HIT
                Case 6: B5AreaDamageB17 = TAIL_HIT
                Case 7: B5AreaDamageB17 = PORT_WING_HIT
                Case 8: B5AreaDamageB17 = STBD_WING_HIT
                Case 9: B5AreaDamageB17 = NO_EFFECT_HIT
                Case 10: B5AreaDamageB17 = WAIST_HIT
                Case 11: B5AreaDamageB17 = FLIGHT_DECK_HIT
                Case 12: B5AreaDamageB17 = WALKING_HITS_FUSELAGE
            End Select
        
        Case F6_HIGH:
            
            Select Case intRoll
                Case 2 To 3: B5AreaDamageB17 = NO_EFFECT_HIT
                Case 4: B5AreaDamageB17 = RADIO_ROOM_HIT
                Case 5: B5AreaDamageB17 = BOMB_BAY_HIT
                Case 6: B5AreaDamageB17 = PORT_WING_HIT
                Case 7: B5AreaDamageB17 = TAIL_HIT
                Case 8: B5AreaDamageB17 = STBD_WING_HIT
                Case 9: B5AreaDamageB17 = WAIST_HIT
                Case 10: B5AreaDamageB17 = FLIGHT_DECK_HIT
                Case 11: B5AreaDamageB17 = WALKING_HITS_FUSELAGE
                Case 12: B5AreaDamageB17 = NOSE_HIT
            End Select
        
        Case F6_LEVEL:
            
            Select Case intRoll
                Case 2 To 3, 12: B5AreaDamageB17 = NO_EFFECT_HIT
                Case 4 To 5, 7, 9 To 10: B5AreaDamageB17 = TAIL_HIT
                Case 6: B5AreaDamageB17 = PORT_WING_HIT
                Case 8: B5AreaDamageB17 = STBD_WING_HIT
                Case 11: B5AreaDamageB17 = WAIST_HIT
            End Select
        
        Case F6_LOW:
            
            Select Case intRoll
                Case 2: B5AreaDamageB17 = NOSE_HIT
                Case 3 To 4: B5AreaDamageB17 = NO_EFFECT_HIT
                Case 5: B5AreaDamageB17 = BOMB_BAY_HIT
                Case 6: B5AreaDamageB17 = PORT_WING_HIT
                Case 7: B5AreaDamageB17 = TAIL_HIT
                Case 8: B5AreaDamageB17 = STBD_WING_HIT
                Case 9: B5AreaDamageB17 = WAIST_HIT
                Case 10: B5AreaDamageB17 = RADIO_ROOM_HIT
                Case 11: B5AreaDamageB17 = WALKING_HITS_FUSELAGE
                Case 12: B5AreaDamageB17 = NOSE_HIT
            End Select
        
        Case VERT_CLIMB:
            
            Select Case intRoll
                Case 2 To 3: B5AreaDamageB17 = NO_EFFECT_HIT
                Case 4: B5AreaDamageB17 = BOMB_BAY_HIT
                Case 5: B5AreaDamageB17 = TAIL_HIT
                Case 6: B5AreaDamageB17 = PORT_WING_HIT
                Case 7: B5AreaDamageB17 = RADIO_ROOM_HIT
                Case 8: B5AreaDamageB17 = STBD_WING_HIT
                Case 9: B5AreaDamageB17 = WALKING_HITS_FUSELAGE
                Case 10: B5AreaDamageB17 = FLIGHT_DECK_HIT
                Case 11: B5AreaDamageB17 = WAIST_HIT
                Case 12: B5AreaDamageB17 = NOSE_HIT
            End Select
        
        Case VERT_DIVE:
            
            Select Case intRoll
                Case 2 To 3: B5AreaDamageB17 = NO_EFFECT_HIT
                Case 4: B5AreaDamageB17 = BOMB_BAY_HIT
                Case 5: B5AreaDamageB17 = RADIO_ROOM_HIT
                Case 6: B5AreaDamageB17 = PORT_WING_HIT
                Case 7: B5AreaDamageB17 = WALKING_HITS_FUSELAGE
                Case 8: B5AreaDamageB17 = STBD_WING_HIT
                Case 9: B5AreaDamageB17 = FLIGHT_DECK_HIT
                Case 10: B5AreaDamageB17 = TAIL_HIT
                Case 11: B5AreaDamageB17 = WAIST_HIT
                Case 12: B5AreaDamageB17 = NOSE_HIT
            End Select
    
    End Select
' B5AreaDamageB17 = PORT_WING_HIT ' debug Nov04
End Function

'******************************************************************************
' B5AreaDamageB24
'
' INPUT:  The direction from which the fighter is attacking.
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Private Function B5AreaDamageB24(ByVal intPosition As Integer, ByVal intSide As Integer) As Integer
    Dim intRoll As Integer
    
    B5AreaDamageB24 = NO_EFFECT_HIT
    
    ' Determine the area that was hit, if any. Damage will be determined after
    ' all hit areas are queued, so that if there is a walking hit we will be
    ' able to delete all the non-walking hits.
    
    intRoll = Random2D6()
    
    Select Case intPosition
        
        Case F12_HIGH, F130_HIGH, F1030_HIGH:
            
            Select Case intRoll
                Case 2 To 5: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 6: B5AreaDamageB24 = NOSE_HIT
                Case 7: B5AreaDamageB24 = FLIGHT_DECK_HIT
                Case 8:
                    If Random1D6() <= 3 Then
                        B5AreaDamageB24 = PORT_WING_HIT
                    Else
                        B5AreaDamageB24 = STBD_WING_HIT
                    End If
                Case 9: B5AreaDamageB24 = WAIST_HIT
                Case 10: B5AreaDamageB24 = TAIL_HIT
                Case 11: B5AreaDamageB24 = BOMB_BAY_HIT
                Case 12: B5AreaDamageB24 = WALKING_HITS_FUSELAGE
            End Select
        
        Case F12_LEVEL, F130_LEVEL, F1030_LEVEL:
            
            Select Case intRoll
                Case 2, 6: B5AreaDamageB24 = PORT_WING_HIT
                Case 3 To 5: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 7: B5AreaDamageB24 = NOSE_HIT
                Case 8, 12: B5AreaDamageB24 = STBD_WING_HIT
                Case 9: B5AreaDamageB24 = FLIGHT_DECK_HIT
                Case 10 To 11: B5AreaDamageB24 = NO_EFFECT_HIT
            End Select
        
        Case F12_LOW, F130_LOW, F1030_LOW:
            
            Select Case intRoll
                Case 2 To 5: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 6: B5AreaDamageB24 = NOSE_HIT
                Case 7: B5AreaDamageB24 = BOMB_BAY_HIT
                Case 8:
                    If Random1D6() <= 3 Then
                        B5AreaDamageB24 = PORT_WING_HIT
                    Else
                        B5AreaDamageB24 = STBD_WING_HIT
                    End If
                Case 9: B5AreaDamageB24 = WAIST_HIT
                Case 10: B5AreaDamageB24 = TAIL_HIT
                Case 11: B5AreaDamageB24 = FLIGHT_DECK_HIT
                Case 12: B5AreaDamageB24 = WALKING_HITS_FUSELAGE
            End Select
        
        Case F3_HIGH, F9_HIGH:
            
            Select Case intRoll
                Case 2: B5AreaDamageB24 = WALKING_HITS_WINGS
                Case 3: B5AreaDamageB24 = NOSE_HIT
                Case 4: B5AreaDamageB24 = FLIGHT_DECK_HIT
                Case 5: B5AreaDamageB24 = BOMB_BAY_HIT
                Case 6: B5AreaDamageB24 = PORT_WING_HIT
                Case 7: B5AreaDamageB24 = TAIL_HIT
                Case 8: B5AreaDamageB24 = STBD_WING_HIT
                Case 9: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 10: B5AreaDamageB24 = WAIST_HIT
                Case 11: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 12: B5AreaDamageB24 = WALKING_HITS_FUSELAGE
            End Select
        
        Case F3_LEVEL, F9_LEVEL:
            
            Select Case intRoll
                Case 2: B5AreaDamageB24 = BOMB_BAY_HIT
                Case 3, 6:
                    If intSide = PORT_SIDE Then
                        B5AreaDamageB24 = PORT_WING_HIT
                    Else
                        B5AreaDamageB24 = STBD_WING_HIT
                    End If
                Case 4: B5AreaDamageB24 = NOSE_HIT
                Case 5: B5AreaDamageB24 = WAIST_HIT
                Case 7 To 8, 12: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 9: B5AreaDamageB24 = TAIL_HIT
                Case 10: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 11: B5AreaDamageB24 = FLIGHT_DECK_HIT
                Case 12: B5AreaDamageB24 = WALKING_HITS_BOTH
            End Select
        
        Case F3_LOW, F9_LOW:
            
            Select Case intRoll
                Case 2: B5AreaDamageB24 = WALKING_HITS_WINGS
                Case 3: B5AreaDamageB24 = NOSE_HIT
                Case 4: B5AreaDamageB24 = BOMB_BAY_HIT
                Case 5: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 6: B5AreaDamageB24 = TAIL_HIT
                Case 7: B5AreaDamageB24 = PORT_WING_HIT
                Case 8: B5AreaDamageB24 = STBD_WING_HIT
                Case 9: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 10: B5AreaDamageB24 = WAIST_HIT
                Case 11: B5AreaDamageB24 = FLIGHT_DECK_HIT
                Case 12: B5AreaDamageB24 = WALKING_HITS_FUSELAGE
            End Select
        
        Case F6_HIGH:
            
            Select Case intRoll
                Case 2 To 3: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 4: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 5: B5AreaDamageB24 = BOMB_BAY_HIT
                Case 6: B5AreaDamageB24 = PORT_WING_HIT
                Case 7: B5AreaDamageB24 = TAIL_HIT
                Case 8: B5AreaDamageB24 = STBD_WING_HIT
                Case 9: B5AreaDamageB24 = WAIST_HIT
                Case 10: B5AreaDamageB24 = FLIGHT_DECK_HIT
                Case 11: B5AreaDamageB24 = WALKING_HITS_FUSELAGE
                Case 12: B5AreaDamageB24 = NOSE_HIT
            End Select
        
        Case F6_LEVEL:
            
            Select Case intRoll
                Case 2 To 3, 12: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 4 To 5, 7, 9 To 10: B5AreaDamageB24 = TAIL_HIT
                Case 6: B5AreaDamageB24 = PORT_WING_HIT
                Case 8: B5AreaDamageB24 = STBD_WING_HIT
                Case 11: B5AreaDamageB24 = WAIST_HIT
            End Select
        
        Case F6_LOW:
            
            Select Case intRoll
                Case 2: B5AreaDamageB24 = NOSE_HIT
                Case 3 To 4: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 5: B5AreaDamageB24 = BOMB_BAY_HIT
                Case 6: B5AreaDamageB24 = PORT_WING_HIT
                Case 7: B5AreaDamageB24 = TAIL_HIT
                Case 8: B5AreaDamageB24 = STBD_WING_HIT
                Case 9: B5AreaDamageB24 = WAIST_HIT
                Case 10: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 11: B5AreaDamageB24 = WALKING_HITS_FUSELAGE
                Case 12: B5AreaDamageB24 = NOSE_HIT
            End Select
        
        Case VERT_CLIMB:
            
            Select Case intRoll
                Case 2 To 3: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 4: B5AreaDamageB24 = BOMB_BAY_HIT
                Case 5: B5AreaDamageB24 = TAIL_HIT
                Case 6: B5AreaDamageB24 = PORT_WING_HIT
                Case 7: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 8: B5AreaDamageB24 = STBD_WING_HIT
                Case 9: B5AreaDamageB24 = WALKING_HITS_FUSELAGE
                Case 10: B5AreaDamageB24 = FLIGHT_DECK_HIT
                Case 11: B5AreaDamageB24 = WAIST_HIT
                Case 12: B5AreaDamageB24 = NOSE_HIT
            End Select
        
        Case VERT_DIVE:
            
            Select Case intRoll
                Case 2 To 3: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 4: B5AreaDamageB24 = BOMB_BAY_HIT
                Case 5: B5AreaDamageB24 = NO_EFFECT_HIT
                Case 6: B5AreaDamageB24 = PORT_WING_HIT
                Case 7: B5AreaDamageB24 = WALKING_HITS_FUSELAGE
                Case 8: B5AreaDamageB24 = STBD_WING_HIT
                Case 9: B5AreaDamageB24 = FLIGHT_DECK_HIT
                Case 10: B5AreaDamageB24 = TAIL_HIT
                Case 11: B5AreaDamageB24 = WAIST_HIT
                Case 12: B5AreaDamageB24 = NOSE_HIT
            End Select
    
    End Select
        
End Function

'******************************************************************************
' B5AreaDamageRouter
'
' INPUT:  The area affected by the hit, and the side, if any.
'
' OUTPUT: n/a
'
' RETURN: End of mission, if the section suffered catastrophic damage.
'
' NOTES:  n/a
'******************************************************************************
Public Function B5AreaDamageRouter(ByVal intHitArea As Integer, ByVal intSide As Integer) As Integer

    ' All the hit areas have been queued. Inflict the damage appropriate to
    ' the current hit.
    
    Select Case intHitArea
        
        Case NO_EFFECT_HIT:
            UpdateMessage "No effect hit"
            Damage.PeckhamPoints = Damage.PeckhamPoints + 1
        Case PORT_WING_HIT: B5AreaDamageRouter = BL1WingDamage(PORT_SIDE)
        Case STBD_WING_HIT: B5AreaDamageRouter = BL1WingDamage(STBD_SIDE)
        Case RADIO_ROOM_HIT: B5AreaDamageRouter = P4RadioRoomDamage()
        Case NOSE_HIT: B5AreaDamageRouter = P1NoseDamage
        Case FLIGHT_DECK_HIT: B5AreaDamageRouter = P2FlightDeckDamage
        Case WAIST_HIT: B5AreaDamageRouter = P5WaistDamage
        Case TAIL_HIT: B5AreaDamageRouter = P6TailDamage
        Case BOMB_BAY_HIT: B5AreaDamageRouter = P3BombBayDamage
        Case WALKING_HITS_FUSELAGE:
            UpdateMessage "Walking hits!"
            
            B5AreaDamageRouter = P1NoseDamage()
            B5AreaDamageRouter = B5AreaDamageRouter + P2FlightDeckDamage()
            B5AreaDamageRouter = B5AreaDamageRouter + P3BombBayDamage()
            B5AreaDamageRouter = B5AreaDamageRouter + P4RadioRoomDamage()
            B5AreaDamageRouter = B5AreaDamageRouter + P5WaistDamage()
            B5AreaDamageRouter = B5AreaDamageRouter + P6TailDamage()
        
        Case WALKING_HITS_WINGS:
            UpdateMessage "Walking hits!"
            
            B5AreaDamageRouter = BL1WingDamage(PORT_SIDE)
            B5AreaDamageRouter = B5AreaDamageRouter + BL1WingDamage(PORT_SIDE)
            B5AreaDamageRouter = B5AreaDamageRouter + BL1WingDamage(STBD_SIDE)
            B5AreaDamageRouter = B5AreaDamageRouter + BL1WingDamage(STBD_SIDE)
        
        Case WALKING_HITS_BOTH:
            UpdateMessage "Walking hits!"
            
            B5AreaDamageRouter = P1NoseDamage()
            B5AreaDamageRouter = B5AreaDamageRouter + P5WaistDamage()
            B5AreaDamageRouter = B5AreaDamageRouter + P6TailDamage()
            B5AreaDamageRouter = B5AreaDamageRouter + BL1WingDamage(intSide)
    
    End Select

    If B5AreaDamageRouter <= END_MISSION Then
        B5AreaDamageRouter = END_MISSION
    End If
            
End Function

'******************************************************************************
' B6SuccessiveAttacks
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The next direction from which the fighter will attack.
'
' NOTES:  n/a
'******************************************************************************
Public Function B6SuccessiveAttacks() As Integer
        
    ' Pass back a new position. The position should be assigned to a fighter.
        
    Select Case Random2D6()
        
        Case 2: B6SuccessiveAttacks = F6_HIGH
        Case 3: B6SuccessiveAttacks = F6_LEVEL
        Case 4: B6SuccessiveAttacks = F9_LEVEL
        Case 5: B6SuccessiveAttacks = F12_LEVEL
        Case 6: B6SuccessiveAttacks = F1030_LEVEL
        Case 7: B6SuccessiveAttacks = F12_HIGH
        Case 8: B6SuccessiveAttacks = F130_LEVEL
        Case 9: B6SuccessiveAttacks = F12_LEVEL
        Case 10: B6SuccessiveAttacks = F3_LEVEL
        Case 11: B6SuccessiveAttacks = F9_HIGH
        Case 12: B6SuccessiveAttacks = F3_HIGH
    
    End Select
    
End Function

' RETURN: End of mission, if the section suffered catastrophic damage.
'******************************************************************************
Private Function B7RandomEvents() As Integer
    Dim intRoll As Integer
    Dim intGun As Integer
    Dim strMessage As String
    Dim bRetryRoll As Boolean
    Dim bNoEvent As Boolean
    
    If Not Bomber.BomberModel = AVRO_LANCASTER Then
        Do
            bRetryRoll = False
            '"The Battle of Berlin", Rule 7E: The Random Events table is never used
            intRoll = Random2D6()
            
            frmMission.lblMiscWave.Caption = vbNullString
        
            Select Case intRoll
            
                Case 2:
                    B702RandomEngineFailure
                Case 3:
                    'Formation casualties
                    ' Note h: This may only happen once per mission, as it makes no
                    ' sense to move from one end of the formation to the other.
                    
                    If RandomEvent.FormationCasualties Or Not Bomber.InFormation Then
                        'Rule 18.0, note b. If this random event is rolled again, ignore and re-roll.
                        'Rule 18.0, note h. If you are out of formation, reroll also.
                        bRetryRoll = True
                    ElseIf Bomber.FormationPos = MIDDLE_PLANE Then
                            strMessage = "Formation casualties. You are now the "
                        
                            If Random1D6() <= 3 Then
                                strMessage = strMessage & "lead"
                                Bomber.FormationPos = LEAD_PLANE
                            Else
                                strMessage = strMessage & "tail"
                                Bomber.FormationPos = TAIL_PLANE
                            End If
                        
                            strMessage = strMessage & " bomber."
                        
                            UpdateMessage strMessage
                        
                            RandomEvent.FormationCasualties = True
                    Else
                        'Rule 18.0, note h. If you are already the lead or tail plane, ignore and re-roll.
                        bRetryRoll = True
                    End If
                
                Case 4:
                    'Loose formation.
                    'NOTRAW: A formation cannot be both loose and tight.
                    
                    If RandomEvent.LooseFormation Then
                        'Rule 18.0, note b. If this random event is rolled again, ignore and re-roll.
                        bRetryRoll = True
                    Else
                        'Formation was either tight or normal.
                        'Rule 18.0, note i. This event still has the same effect, even if out of formation.
                        RandomEvent.TightFormation = False
                        RandomEvent.LooseFormation = True
                        UpdateMessage "Loose formation."
                    End If
                
                Case 5:
                    'Aggressive "Little Friends".
                    If RandomEvent.AggressiveCover Then
                        'Rule 18.0, note b. If this random event is rolled again, ignore and reroll.
                        bRetryRoll = True
                    ElseIf Mission.Options.Unescorted Then
                        'If we're unescorted we can't have aggressive cover.
                        'NOTRAW: in this case, reroll.
                        bRetryRoll = True
                    Else
                        RandomEvent.AggressiveCover = True
                        UpdateMessage "Aggressive fighter cover."
                    End If
                
                Case 6, 8:
                    ' Tight formation.
                    ' NOTRAW: A formation cannot be both loose and tight.
                
                    If RandomEvent.TightFormation Then
                        'Rule 18.0, note b. If this random event is rolled again, ignore and re-roll.
                        bRetryRoll = True
                    Else
                        RandomEvent.LooseFormation = False
                        RandomEvent.TightFormation = True
                        UpdateMessage "Tight formation."
                    End If
                
                Case 7:
                    
                    ' Note c: This may happen multiple times. Rather than having the
                    ' user decide when to expend some luck, the luck will automatically
                    ' be expended when some horrible tragedy occurs, such as burst in
                    ' plane or the plane is shot down.
                    'FIXME: offer the rabbit's foot on any negative die result.
                    
                    Bomber.RabbitsFoot = Bomber.RabbitsFoot + 1
                    UpdateMessage "Your crew feels luckier ..."
                
                Case 9:
                    
                    ' Note d: This may happen multiple times.
                    RandomEvent.BadLuftwaffeComm = Not RandomEvent.BadLuftwaffeComm
                    
                    'NOTE Should the user be notified of this? Maybe make it optional?
                    If RandomEvent.BadLuftwaffeComm Then
                        UpdateMessage "The Luftwaffe seem poorly co-ordinated"
                    Else
                        UpdateMessage "The Luftwaffe seem better co-ordinated"
                    End If
                    
                Case 10:
                
                    If Not Bomber.Altitude = HIGH_ALTITUDE Then
                        'Rule 18.0 note e. If you are out of formation at 10,000 feet, ignore this result and re-roll.
                        
                        bRetryRoll = True
                    Else
                    
                        UpdateMessage "Extreme cold."
                    
                        For intGun = MID_UPPER_MG To TAIL_MG
                            
                            If GunExists(intGun) Then
                                
                                If Random1D6() = 6 Then
                                    Bomber.Gun(intGun).Status = MG_JAMMED
                                    Bomber.Gun(intGun).Frozen = True
                                    UpdateMessage frmMission.lblGunName(intGun).Caption & " jammed due to cold."
                                    'FIXME: Needs to be auto-repaired at 10K feet.
                                
                                End If
                            
                            End If
                        
                        Next intGun
                    End If
                
                    RandomEvent.ExtremeCold = True
                
                Case 11: 'ace for a day
                    Dim AcePosition As Integer
                    UpdateMessage "TODO: ace for a day."
                    
                    Select Case Random1D6
                        Case 1 To 2:
                            AcePosition = ENGINEER
                        Case 3 To 4:
                            AcePosition = BALL_GUNNER
                        Case 5 To 6:
                            AcePosition = TAIL_GUNNER
                    End Select
                    Bomber.Airman(AcePosition).AceForADay = True
                    UpdateMessage Bomber.Airman(AcePosition).Name & " is having a great day!"
                
                Case 12: 'mid-air accident
                    B712MidAirAccident
                    
            End Select
        Loop Until bRetryRoll = False
        If bNoEvent Then
            frmMission.lblMiscWave.Caption = "No Attackers"
            UpdateMessage "No attackers."
        End If
    End If
    
End Function


Public Sub B702RandomEngineFailure()
                    
    'Rules Page 9, section 18.0
    'Note a: If this random event is rolled again, the previously-failed engine may be able to restart.
    Dim intEngine As Integer
    Dim intFeathering As Integer
    
    If RandomEvent.EngineFailure = 0 Then
        'Engine failure.
        'No current random engine failure. Break a random engine.
        RandomEvent.EngineFailure = RandomDX(4)
        Damage.EngineOut(RandomEvent.EngineFailure) = True
        UpdateMessage "Engine failure in # " & intEngine & "."
        
        'Table BL-1 note c.
        'Check if prop was feathered when it failed.
        intFeathering = Random1D6
        If intFeathering = 6 Then
            UpdateMessage "  Prop not feathered."
            Damage.EngineDrag(intEngine) = True
            DropOutOfFormation
        Else
            UpdateMessage "  Prop feathered."
        End If
    Else
        'An engine was previously failed.
        'NOTRAW It is able to restart if it was not shutdown due to damage.
        '(either engine hit or oil leak runout)
        'FIXME: a battle-damaged engine should stay failed
        If Damage.OilTankLeak(RandomEvent.EngineFailure) > NO_OIL Then
            Damage.EngineDrag(RandomEvent.EngineFailure) = False
            Damage.EngineOut(RandomEvent.EngineFailure) = False
        End If
        
        RandomEvent.EngineFailure = 0
        UpdateMessage "#" & RandomEvent.EngineFailure & " engine restarted."
    End If
End Sub

Private Sub B712MidAirAccident()
    'NOTE: this previously indicated that "Theater Modifications: More Expansions
    'for B-17" said that mid-air collision were impossible in August 1942.
    'I cannot find this in that variant. It does say "While risk of collision was
    'low [prior to September 1942], but it doesn't suggest any rules change to
    'account for that.
    Dim intRoll As Integer
    
    If Not Bomber.InFormation Then
        'Rule 18.0 note g. If you are out of formation, treat this result as #2 (engine failure) instead.
        B702RandomEngineFailure
    Else
    
        RandomEvent.MidAirAccident = True
        
        intRoll = Random2D6()
        
        If (Bomber.Position(PILOT).CurrentSerialNum = Bomber.Position(PILOT).AssignedSerialNum _
        And Bomber.Airman(PILOT).Status <= LW2_STATUS) _
        Or (Bomber.Position(PILOT).CurrentSerialNum = Bomber.Position(COPILOT).AssignedSerialNum _
        And Bomber.Airman(COPILOT).Status <= LW2_STATUS) Then

            ' Crew Experience: "The General" (Volume 24, #6) variant.
            ' The airman occupying the pilot's seat is a pilot or copilot,
            ' who is not incapacitated.
            
            If Mission.Options.CrewExperience Then
                
                ' Crew experience variant from the "Theater Modifications"
                ' article in "The General" (Volume 24, #6).

                If Bomber.Airman(PILOT).Mission <= 5 _
                And Bomber.Airman(COPILOT).Mission <= 5 Then

                    ' Novice pilot is more likely to have severe accident.
                    intRoll = intRoll + 1
            
                ElseIf Bomber.Airman(PILOT).Mission >= 11 _
                Or Bomber.Airman(COPILOT).Mission >= 11 Then

                    ' Veteran pilot is less likely to have severe accident.
                    intRoll = intRoll - 1
                
                End If
            
            End If
            
        Else
            'NOTRAW: Non-pilot even more likely to have more severe collision.
            intRoll = intRoll + 2
        
        End If

        Select Case intRoll
            Case 2 To 8:
                
                UpdateMessage "Close call, but no effect."
                
                'Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    
            Case 9 To 10:
            
                UpdateMessage "Avoided mid-air collision with a shallow dive."
                
                'Damage.PeckhamPoints = Damage.PeckhamPoints + 25
                    
                Call DropOutOfFormation
                'FIXME: we should be able to return to formation in next zone
            
            Case 11:
            
                If Bomber.RabbitsFoot >= 1 Then
                    ' Expend luck to prevent loss of aircraft.
                    UpdateMessage "Luckily avoided mid-air collision"
                    Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                Else
                    
                    UpdateMessage "Avoided mid-air collision with a steep dive."
                    
                    ' Damage.PeckhamPoints = Damage.PeckhamPoints + 50
                    
                    If Random1D6() = 6 Then
                        
                        UpdateMessage "Stress tears off port wing!"
                        G7UncontrolledBailout (OverWater())
                    
                    ElseIf Random1D6() = 6 Then
                        
                        UpdateMessage "Stress tears off starboard wing!"
                        G7UncontrolledBailout (OverWater())
                    Else
                        If Bomber.BombsOnBoard = True _
                        Or Bomber.ExtraAmmo >= 1 _
                        Or Bomber.ExtraFuelInBombBay = True Then
                        
                            UpdateMessage "Payload rips through bomb bay doors."
                            
                            Bomber.BombsOnBoard = False
                            Bomber.ExtraAmmo = 0
                            Bomber.ExtraFuelInBombBay = False
                            
                            Damage.BombBayDoors = True
                            Damage.BombRelease = True
                            Damage.PeckhamPoints = Damage.PeckhamPoints + 35
                        
                        End If
                        
                    End If
                
                    Call DropOutOfFormation

                    If AlpsDirection() = ALPS_BELOW Then
                        G7UncontrolledBailout (False)
                        Exit Sub
                    Else
                        Call LoseAltitude
                    End If
                
                End If

            Case 12:
                
                If Bomber.RabbitsFoot >= 1 Then
                    ' Expend luck to prevent loss of aircraft.
                    UpdateMessage "Luckily avoided mid-air collision"
                    Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
                Else
                        
                    UpdateMessage "Mid-air collision!"
                    G7UncontrolledBailout (OverWater())
                    Exit Sub
            
                End If
        
        End Select
    
    End If

End Sub

'******************************************************************************
' M1DefensiveFire
'
' INPUT:  The bomber's model, the gun that is firing, and the type of fighter
'         being fired at.
'
' OUTPUT: n/a
'
' RETURN: The roll required to hit the fighter.
'
' NOTES:  n/a
'******************************************************************************
Public Function M1DefensiveFire(ByVal intBomberModel As Integer, ByVal intGun As Integer, ByVal intPosition As Integer, ByVal strTargetType As String, ByVal intAirmanStatus As Integer) As Integer
    Dim rsGunnery As New ADODB.Recordset
    Dim intToHit As Integer
    Dim strErrMsg As String
    
    M1DefensiveFire = 0
    intToHit = 0
    
    pobjCmnd.CommandText = " SELECT * FROM Gunnery" & _
                           " WHERE BomberModel = " & intBomberModel & _
                           " AND GunPos = " & intGun & _
                           " AND FighterPos = " & intPosition
    
    rsGunnery.CursorLocation = adUseClient
    rsGunnery.Open pobjCmnd, , adOpenStatic, adLockBatchOptimistic
    
    If RecordsInSet(rsGunnery) = 0 Then
        
        intToHit = 0
    
    Else
        
        If rsGunnery.RecordCount > 0 Then
            
            Select Case strTargetType
                
                Case "Ju88":
                    intToHit = rsGunnery![Ju88]
                Case "Me109":
                    intToHit = rsGunnery![Me109]
                Case "Me110":
                    intToHit = rsGunnery![Me110]
                Case "FW190":
                    intToHit = rsGunnery![FW190]
            
            End Select
    
            ' BL-2 Instruments: Note b.
            
            If Damage.IntercomSystem = True Then
            
                ' Override queried value when intercom is out.
            
                If intGun = TAIL_MG Then
                    intToHit = 5
                Else
                    intToHit = 6
                End If
            End If
            
            ' BL-4 Wounds: Note a.
            
            If intAirmanStatus = LW2_STATUS Then
                ' Override any previous value if the airman is twice wounded.
                intToHit = 6
            End If
                
        End If
    
    End If
    
    M1DefensiveFire = intToHit
    
CleanUp:
   
    If Not rsGunnery Is Nothing Then
        If rsGunnery.State = adStateClosed Then rsGunnery.Close
        Set rsGunnery = Nothing
    End If
   
    Exit Function
   
ErrorTrap:
    
    strErrMsg = "Error " & CStr(Err.Number) & vbCrLf & vbCrLf & _
                "M1DefensiveFire() " & vbCrLf & vbCrLf & _
                Err.Description

    MsgBox strErrMsg, (vbCritical + vbOKOnly)
    
    Err.Clear
    
    M1DefensiveFire = 0
    
    Resume CleanUp

End Function

'******************************************************************************
' M2HitDamageAgainstGermanFighter
'
' INPUT:  The bonus (1) for twin .50s and quad .303s, and whether or not the
'         fighter is a FW-190.
'
' OUTPUT: n/a
'
' RETURN: The damage inflicted on the fighter.
'
' NOTES:  n/a
'******************************************************************************
Public Function M2HitDamageAgainstGermanFighter(ByVal intBonus As Integer, ByVal blnFW190 As Boolean) As Integer
    Dim intRoll As Integer
    
    M2HitDamageAgainstGermanFighter = 0

    intRoll = Random1D6() + intBonus
    
    If blnFW190 = True Then
        intRoll = intRoll - 1
    End If
    
    Select Case intRoll
        Case Is <= 2: M2HitDamageAgainstGermanFighter = FCA_DAMAGE
        Case 3 To 4: M2HitDamageAgainstGermanFighter = FBOA_DAMAGE
        Case Is >= 5: M2HitDamageAgainstGermanFighter = SHOT_DOWN_DAMAGE
    End Select
    
End Function

'******************************************************************************
' M3GermanOffensiveFire
'
' INPUT:  The direction from which the fighter is attacking, the skill of the
'         fighter pilot, the damage already suffered by the fighter, and whether
'         or not the bomber is performing evasive maneuvers.
'
' OUTPUT: n/a
'
' RETURN: True if the fighter hit the bomber, otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function M3GermanOffensiveFire(ByVal intPosition As Integer, ByVal intFighterPilotSkill As Integer, ByVal intDamage As Integer, ByVal blnEvadeFighters As Boolean) As Boolean
    Dim intRoll As Integer
    Dim intToHit As Integer
    Dim intEnginesOut As Integer

    M3GermanOffensiveFire = False

    intRoll = Random1D6()

    If intRoll = 6 Then
        ' 6 is always a hit, regardless of modifiers
        M3GermanOffensiveFire = True
        Exit Function
    End If
    
    ' Rule 15.1.a
    
    If blnEvadeFighters = True Then
        intRoll = intRoll - 1
    End If

    ' Rule 20.0, plus variant.
    
    Select Case intFighterPilotSkill
        Case GREEN_PILOT:
            intRoll = intRoll - 1
        Case VET_PILOT:
            ' No modifier
        Case ACE_PILOT:
            intRoll = intRoll + 1
        Case ACE_OF_ACES:
            intRoll = intRoll + 2
    End Select

    ' Rule 10.2: Engine(s) out = slow plane = easier to hit
        
    intEnginesOut = CountEnginesOut()
        
    If intEnginesOut >= 2 Then
        intRoll = intRoll + 1
    End If

    Select Case intDamage
        Case Is >= SHOT_DOWN_DAMAGE:
            ' A shot down fighter should never get to this point.
            Exit Function
        Case Is >= FBOA_DAMAGE: intRoll = intRoll - 2
        Case FCA_DAMAGE: intRoll = intRoll - 1
    End Select

    ' Now that the roll has been modified by all applicable rules,
    ' find out what the fighter needs to hit.
    
    Select Case intPosition
        Case F12_HIGH To F12_LOW: intToHit = 5
        Case F130_HIGH To F130_LOW: intToHit = 5
        Case F3_HIGH To F3_LOW: intToHit = 4
        Case F6_HIGH To F6_LOW: intToHit = 3
        Case F9_HIGH To F9_LOW: intToHit = 4
        Case F1030_HIGH To F1030_LOW: intToHit = 5
        Case VERT_DIVE: intToHit = 5
        Case VERT_CLIMB: intToHit = 4
    End Select

    ' If the roll exceeds the to hit, there was a hit.
    
    If intRoll >= intToHit Then
        M3GermanOffensiveFire = True
    End If
    
End Function


'******************************************************************************
' M4FighterCoverDefense
'
' INPUT:  The amount of friendly fighter cover in the zone.
'
' OUTPUT: n/a
'
' RETURN: The number of German fighters that may be driven off.
'
' NOTES:  G5FighterCover determines how good coverage is. M4FighterCoverDefense
'         determines how many German fighters are chased away.
'******************************************************************************
Public Function M4FighterCoverDefense(ByVal strCover As String) As Integer

    Dim intRoll As Integer

    M4FighterCoverDefense = 0
    
    If strCover = NO_COVER Then
        Exit Function
    End If

    intRoll = Random1D6()
    
'UpdateMessage vbCrLf & "strCover = '" & strCover & "'" ' DEBUG
'UpdateMessage "Wave.Attack = [" & Wave.Attack & "]" ' DEBUG
'UpdateMessage "intRoll (Bef) = [" & intRoll & "]" ' DEBUG

    ' O-1 Weather: Note a and Note b, plus "The General" (Volume 24, #6)
    ' variant.
        
    If Mission.Zone(Bomber.CurrentZone).Weather = CLEAR_WEATHER Then
        intRoll = intRoll + 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = POOR_WEATHER Then
        intRoll = intRoll - 1
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = BAD_WEATHER Then
        intRoll = intRoll - 2
    ElseIf Mission.Zone(Bomber.CurrentZone).Weather = STORM_WEATHER Then
        ' No fighter protection is possible.
        M4FighterCoverDefense = 0
        Exit Function
    End If

    ' Rule 18.0
    
    If RandomEvent.AggressiveCover = True Then
        intRoll = intRoll + 1
    End If
    
    ' JG26 variant.

    If Wave.JG26 = True Then
        strCover = POOR_COVER
    End If

    ' The three levels of coverage really form one table column, if a
    ' modifier is applied to the roll.
    
    If strCover = FAIR_COVER Then
        intRoll = intRoll - 2
    ElseIf strCover = POOR_COVER Then
        intRoll = intRoll - 4
    End If
    
    If Wave.Attack >= 2 Then
        intRoll = intRoll - 2
    End If
    
    ' Return the number of German fighters that were chased off by friendly
    ' fighters.
    
'UpdateMessage "intRoll (Aft) = [" & intRoll & "]" ' DEBUG
    
    Select Case intRoll
        Case Is <= 0: M4FighterCoverDefense = 0
        Case 1 To 2: M4FighterCoverDefense = 1
        Case 3 To 4: M4FighterCoverDefense = 2
        Case Is >= 5: M4FighterCoverDefense = 3
    End Select

'UpdateMessage "M4FighterCoverDefense = [" & M4FighterCoverDefense & "]" ' DEBUG
    
End Function

'******************************************************************************
' M5SprayAreaFire
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: The spray fire's effect on a German fighter.
'
' NOTES:  n/a
'******************************************************************************
Public Function M5SprayAreaFire() As Integer
    Dim intRoll As Integer

    M5SprayAreaFire = 0

    intRoll = Random1D6()
    
    ' 1 MG Jams / Fighter attacks
    ' 2 Fighter attacks
    ' 3 Fighter attacks
    ' 4 Fighter breaks off
    ' 5 Fighter breaks off
    ' 6 Fighter hit, roll on M2HitDamageAgainstGermanFighter
    
    Select Case intRoll
        Case 1:
            ' MG jams
            M5SprayAreaFire = SPRAY_FIRE_JAM
        Case 2 To 3:
            ' No effect
            M5SprayAreaFire = SPRAY_FIRE_NOEFFECT
        Case 4 To 5:
            ' Fighter chased off prior to attacking
            M5SprayAreaFire = SPRAY_FIRE_BREAKOFF
        Case 6:
            ' Fighter hit, roll on M2HitDamageAgainstGermanFighter
            M5SprayAreaFire = SPRAY_FIRE_HIT
    End Select

End Function

'******************************************************************************
' M6FighterPilotStatus
'
' INPUT:  n/a
'
' OUTPUT: n/a
'
' RETURN: Normal pilot's skill level.
'
' NOTES:  n/a
'******************************************************************************
Private Function M6FighterPilotStatus() As Integer
    Dim intRoll As Integer
    
    M6FighterPilotStatus = VET_PILOT
        
    intRoll = Random2D6()

    If Mission.Options.GermanFighterPilotSkill = True Then
        
        Select Case intRoll
            Case 2 To 3:
                M6FighterPilotStatus = GREEN_PILOT
            Case 4 To 10:
                M6FighterPilotStatus = VET_PILOT
            Case 11 To 12:
                M6FighterPilotStatus = ACE_PILOT
        End Select
    
    End If

End Function

Private Function BadRoll() As Boolean
    'Return value: whether or not the bad result should be re-rolled
    BadRoll = False
    Dim response As VbMsgBoxResult
    If Bomber.RabbitsFoot > 0 Then
        response = MsgBox("Would you like to use a rabbit's foot?", vbYesNo, "Uh-oh")
        If response = vbYes Then
            UpdateMessage "...but you luckily avoid it"
            Bomber.RabbitsFoot = Bomber.RabbitsFoot - 1
            BadRoll = True
        End If
    End If
End Function
