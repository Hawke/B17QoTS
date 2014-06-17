'******************************************************************************
' modUserDefinedTypes.bas
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

Attribute VB_Name = "modUserDefinedTypes"
Option Explicit

Public gblnCrewAssigned As Boolean

'******************************************************************************
' ZoneInfo
'
' NOTES: Describes the zone itself, plus any fighter coverage one may expect.
'******************************************************************************
Private Type ZoneInfo
    Modifier As Integer
    Terrain As String
    Weather As Integer
    Contrail As Boolean
    CoverBack As String
    CoverOut As String
End Type

'******************************************************************************
' MissionOptions
'
' NOTES: The options chosen by the use on the Generate Mission Screen.
'
'        ExtraAmmoInBombBay is an amount, not a boolean, so it part of the
'        BomberInfoNew structure (ExtraAmmo).
'******************************************************************************
Private Type MissionOptions
    RandomEvents As Boolean
    MechanicalFailures As Boolean
    TimePeriodSpecificFormations As Boolean
    FormationDefensiveGunnery As Boolean
    EvadeFlak As Boolean
    CrewExperience As Boolean
    AlternateWeather As Boolean
    GermanFighterPilotSkill As Boolean
    JG26StationedInAbbeville As Boolean
    Ju88sUsedAsFighters As Boolean
    Unescorted As Boolean
    RedTailAngels As Boolean
    Delay As Integer
End Type

'******************************************************************************
' MissionInfo
'
' NOTES: Describes the mission and target, including the flight plan as Zone()
'        and any options. Declare the public structure "Mission".
'******************************************************************************
Private Type MissionInfo
    TargetName As String
    TargetType As String
    HeavyFlak As Boolean
    TargetZone As Integer ' Controls looping Zone()
    Zone(BASE_ZONE To MAX_ZONE) As ZoneInfo
    Date As Integer
    Options As MissionOptions
    AlpsZone As Integer
End Type

Public Mission As MissionInfo

'******************************************************************************
' AirmanInfoNew
'
' NOTES: Describes one airman.
'******************************************************************************
Private Type AirmanInfoNew ' ne: CrewInfo
    Name As String
    SerialNumber As Integer
    Mission As Integer
    Kills As Integer
    Status As Integer
    LeadCrewExp As String
    Wounded As Boolean
    Frostbite As Boolean
    AssignedPosition As Integer ' The position the airman started the mission at
End Type

'******************************************************************************
' CrewPositionInfoNew
'
' NOTES: Track who is, or was, at the position.
'******************************************************************************
Private Type CrewPositionInfoNew
    AssignedSerialNum As Integer ' The airman who started the mission
    CurrentSerialNum As Integer ' The airman currently in the position
End Type

'Nose (1g-15a) / Nose Turret (2g-15a or 2g-30a)
'Stbd Cheek (1g-10a)
'Port Cheek (1g-10a)
'Top Turret (2g-16a or 2g-32a)
'Mid-Upper (2g-16a)
'Radio Room (1g-10a)
'Stbd Waist (1g-20a or 2g-20a)
'Port Waist  (1g-20a or 2g-20a)
'Ball Turret (2g-20a or 2g-24a) / Ventral Blister (2g-20a) / Tunnel (1g-20a) / Floor Ring (2g-20a)
'Tail (2g-23a or 2g-25a or 4g-63a) / Stinger (2g-23a)
'
'Things That Describe Gun Positions
'----------------------------------
'Firepower Bonus
'Ammo
'Status
'Airman Manning It
'Is Airman Qualified?

'Field of Fire(20)
'******************************************************************************
' GunInfoNew
'
' NOTES: Describes a gun or turret.
'
' TODO: Is QualifiedGunner being used anywhere? Should it be used?
'******************************************************************************
Private Type GunInfoNew
   Bonus As Integer ' 0 = normal guns; 1 = twin .50s or quad .303s
   Ammo As Integer  ' current ammo
   MaxAmmo As Integer ' maximum ammo
   Status As Integer
   MannedBy As Integer ' TODO: not necessary if gun is a property of position
   QualifiedGunner As Boolean 'Is airman qualified? If not, 6 to hit.
End Type

'******************************************************************************
' BomberInfoNew
'
' NOTES: Describes a bomber, including its crew, positions and weapons. Declare
'        the public structure "Bomber".
'******************************************************************************
Private Type BomberInfoNew
    Name As String
    KeyField As Integer
    BomberModel As Integer
    TailNumber As String
    Manufacturer As String
    Plant As String
    Mission As Integer ' Missions flown, including the current mission
    Status As Integer
    SquadronPos As String
    FormationPos As String
    CurrentZone As Integer
    Direction As Integer
    InFormation As Boolean
    Altitude As Integer
    TurnsInZone As Integer
    SpottedBySearchLight As Boolean
    Airman(PILOT To AMMO_STOCKER) As AirmanInfoNew
    Position(PILOT To AMMO_STOCKER) As CrewPositionInfoNew
    Gun(MID_UPPER_MG To TAIL_MG) As GunInfoNew
    HandHeldExtinguishers As Integer ' 5 total
    EngineExtinguisher(1 To 4) As Integer ' 2 per engine
    FuelPoints As Integer
    ExtraFuelInBombBay As Boolean
    ExtraAmmo As Integer ' 0 on all planes but the YB40
    BombsOnBoard As Boolean ' True (except YB40), changed to False after Bomb Run
    RabbitsFoot As Integer
End Type

Public Bomber As BomberInfoNew

'******************************************************************************
' FighterInfoNew
'
' NOTES: Describe an enemy fighter.
'******************************************************************************
Private Type FighterInfoNew
   Type As String
   Position As Integer
   PilotSkill As String
   Damage As Integer
   Status As String
   Special As String
End Type

'******************************************************************************
' WaveInfoNew
'
' NOTES: Describe a wave of enemy fighters. There are 1 to 6 fighters in every
'        wave, discounting waves where there are no fighters at all, and thus
'        are not really waves. (e.g., 16, 26, etc.) If there are any waves,
'        then 1 to 3 waves will be encountered. Declare the public structure
'        "Wave".
'******************************************************************************
Public Type WaveInfoNew
    JG26 As Boolean
    Ju88 As Boolean
    Fighter() As FighterInfoNew ' Up to five fighters (six for formation lead)
    Attack As Integer ' Each wave may attack up to three times
End Type

Public Wave As WaveInfoNew

'******************************************************************************
' RandomEventNew
'
' NOTES: Declare the public structure "RandomEvent".
'
'        RabbitsFoot is an amount, not a boolean, so it part of the
'        BomberInfoNew structure.
'******************************************************************************
Private Type RandomEventNew
    EngineFailure As Boolean
    FormationCasualties As Boolean
    LooseFormation As Boolean
    AggressiveCover As Boolean
    TightFormation As Boolean
    BadLuftwaffeComm As Boolean
    ExtremeCold As Boolean
    AceForADay As Boolean ' TODO: How to handle this?
    MidAirAccident As Boolean
End Type

Public RandomEvent As RandomEventNew

'******************************************************************************
' DamageInfoNew
'
' NOTES: Describe every possible type of damage a bomber may sustain. Declare
'        the public structure "Damage".
'******************************************************************************
Private Type DamageInfoNew
    PeckhamPoints As Integer         ' 0
    BombSight As Boolean             ' False
    NoseWheel As Boolean             ' False
    NavigationEquipment As Boolean   ' False
    BombControls As Boolean          ' False (Located in Nose Section)
    Window As Integer                ' 0
    ControlCables As Integer         ' 0
    RubberRafts As Boolean           ' False
    BombRelease As Boolean           ' False (Located in Bomb Bay)
    BombBayDoors As Boolean          ' False
    IntercomSystem As Boolean        ' False
    OxygenSystem As Boolean          ' False
    WingFlapControls As Boolean      ' False
    AileronControls As Boolean       ' False
    ElevatorControls As Boolean      ' False
    RudderControls As Boolean        ' False
    Radio As Boolean                 ' False
    BallTurretMech As Boolean        ' False
    Tailwheel As Boolean             ' False
    Autopilot As Boolean             ' False
   
    PortAmmoTrack As Boolean         ' False
    StbdAmmoTrack As Boolean         ' False
    AmmoTrackFirings As Integer      ' 0
    PortAmmoBox As Integer           ' 0
    StbdAmmoBox As Integer           ' 0
    AmmoBoxFirings As Integer        ' 0
    
    Rudder(1 To 2) As Integer        ' 0
    TailplaneRoot(1 To 2) As Integer ' 0
    WingRoot(1 To 2) As Integer      ' 0
    WingFlap(1 To 2) As Boolean      ' False
    Aileron(1 To 2) As Boolean       ' False
    Elevator(1 To 2) As Boolean      ' False
    EngineOut(1 To 4) As Boolean     ' False
    EngineDrag(1 To 4) As Boolean    ' False
    OilTankLeak(1 To 4) As Integer   ' 0
    FuelTankHits As Integer          ' 0
    FuelTransferSystem As Boolean    ' False
    HeatingSystem As Boolean         ' False
    Turbocharger(1 To 4) As Boolean  ' False
    BurstInPlane As Boolean          ' False
    Heater(PILOT To AMMO_STOCKER) As Boolean ' False
    Oxygen(PILOT To AMMO_STOCKER) As Integer ' 0
    
    Brake As Boolean                 ' False
    LandingGear As Boolean           ' False
    FeatheringCtrl As Boolean        ' False
    EngineExtCtrl As Boolean         ' False
    Electrical As Boolean            ' False
    LastZone As Integer              ' BASE_ZONE
End Type

Public Damage As DamageInfoNew

