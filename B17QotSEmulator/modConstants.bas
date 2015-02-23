Attribute VB_Name = "modConstants"
'******************************************************************************
' modConstants.bas
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

Public pintDoo(1 To 2) As Integer
Public pintDah As Integer
Public pblnDoBailOrCrash As Boolean

' ---( Date Constants )-------------------------------------------------------

Public Const CAMPAIGN_DURATION As Integer = 34
Public Const AUG_1942 As Integer = 0
Public Const SEP_1942 As Integer = 1
Public Const OCT_1942 As Integer = 2
Public Const NOV_1942 As Integer = 3
Public Const DEC_1942 As Integer = 4
Public Const JAN_1943 As Integer = 5
Public Const FEB_1943 As Integer = 6
Public Const MAR_1943 As Integer = 7
Public Const APR_1943 As Integer = 8
Public Const MAY_1943 As Integer = 9
Public Const JUN_1943 As Integer = 10
Public Const JUL_1943 As Integer = 11
Public Const AUG_1943 As Integer = 12
Public Const SEP_1943 As Integer = 13
Public Const OCT_1943 As Integer = 14
Public Const NOV_1943 As Integer = 15
Public Const DEC_1943 As Integer = 16
Public Const JAN_1944 As Integer = 17
Public Const FEB_1944 As Integer = 18
Public Const MAR_1944 As Integer = 19
Public Const APR_1944 As Integer = 20
Public Const MAY_1944 As Integer = 21
Public Const JUN_1944 As Integer = 22
Public Const JUL_1944 As Integer = 23
Public Const AUG_1944 As Integer = 24
Public Const SEP_1944 As Integer = 25
Public Const OCT_1944 As Integer = 26
Public Const NOV_1944 As Integer = 27
Public Const DEC_1944 As Integer = 28
Public Const JAN_1945 As Integer = 29
Public Const FEB_1945 As Integer = 30
Public Const MAR_1945 As Integer = 31
Public Const APR_1945 As Integer = 32
Public Const MAY_1945 As Integer = 33

' ---( Tab Constants )--------------------------------------------------------

Public Const GROUP_TAB As Integer = 0
Public Const SQUADRON_TAB As Integer = 1
Public Const BOMBER_TAB As Integer = 2
Public Const AIRMAN_TAB As Integer = 3
Public Const MISSION_TAB As Integer = 4

' ---( Terrain Constants )----------------------------------------------------

Public Const ALBANIA_TER As String = "Alb"
Public Const ALBANIA_YUGOSLAVIA_TER As String = "Alb-Y"
Public Const ALPS_TER As String = "Alps"
Public Const AUSTRIA_TER As String = "A"
Public Const BELGIUM_TER As String = "B"
Public Const BELGIUM_GERMANY_TER As String = "B-G"
Public Const BULGARIA_TER As String = "Bul"
Public Const BULGARIA_RUMANIA_TER As String = "Bul-R"
Public Const CZECH_TER As String = "Cze"
Public Const DENMARK_TER As String = "Den"
Public Const ENGLAND_TER As String = "E"
Public Const FRANCE_TER As String = "F"
Public Const GERMANY_TER As String = "G"
Public Const GERMANY_CZECH_TER As String = "G-Cze"
Public Const GREECE_TER As String = "Gre"
Public Const HUNGARY_TER As String = "H"
Public Const ITALY_TER As String = "I"
Public Const ITALY_YUGOSLAVIA_TER As String = "I-Y"
Public Const NETHERLANDS_TER As String = "N"
Public Const NETHERLANDS_GERMANY_TER As String = "N-G"
Public Const NORWAY_TER As String = "Nwy"
Public Const RUMANIA_TER As String = "R"
Public Const WATER_TER As String = "W"
Public Const WATER_ALBANIA_TER As String = "W-Alb"
Public Const WATER_FRANCE_TER As String = "W-F"
Public Const WATER_GERMANY_TER As String = "W-G"
Public Const WATER_GREECE_TER As String = "W-Gre"
Public Const WATER_ITALY_TER As String = "W-I"
Public Const WATER_NETHERLANDS_TER As String = "W-N"
Public Const YUGOSLAVIA_TER As String = "Y"
Public Const YUGOSLAVIA_AUSTRIA_TER As String = "Y-A"
Public Const YUGOSLAVIA_BULGARIA_TER As String = "Y-Bul"

Public Const ALPS_NOWHERE As Integer = 0
Public Const ALPS_AHEAD As Integer = 1
Public Const ALPS_BEHIND As Integer = 2
Public Const ALPS_BELOW As Integer = 3
Public Const ALPS_NEXT_ZONE As Integer = 4

' ---( Assignment Constants )-------------------------------------------------

Public Const ADMIN_DUTY As Integer = 0

' ---( Airman States )--------------------------------------------------------

Public Const DUTY_STATUS As Integer = 0
Public Const LW1_STATUS As Integer = 1
Public Const LW2_STATUS As Integer = 2
Public Const SW_STATUS As Integer = 3
Public Const KIA_STATUS As Integer = 4
Public Const INVALID_STATUS As Integer = 5
Public Const POW_STATUS As Integer = 6
Public Const MIA_STATUS As Integer = 7
Public Const DOW_STATUS As Integer = 8
Public Const TOUR_COMPLETE_STATUS As Integer = 9

Public Const MAX_MISSIONS As Integer = 25

' ---( Crash States )---------------------------------------------------

'Public Const CRASHED_STATUS As Integer = 1
Public Const BAD_CRASH_STATUS As Integer = 2
'Public Const KIA_STATUS As Integer = 4
'Public Const MIA_STATUS As Integer = 7

' ---( Bomber States )--------------------------------------------------------

' Public Const DUTY_STATUS As Integer = 0
Public Const CRASHED_STATUS As Integer = 1
Public Const CAPTURED_STATUS As Integer = 2
Public Const DITCHED_STATUS As Integer = 3
Public Const SHOT_DOWN_STATUS As Integer = 4
Public Const RETIRED_STATUS As Integer = 5
Public Const STAND_DOWN_STATUS As Integer = 6 ' Does not have full crew
Public Const SCRAPPED_STATUS As Integer = 7

' ---( Lookup Constants )-----------------------------------------------------

Public Const LOOKUP_BY_LISTINDEX As Integer = 1
Public Const LOOKUP_BY_KEYFIELD As Integer = 2

' ---( Bomber Types )---------------------------------------------------------

Public Const B17_TYPE As Integer = 1
Public Const B24_TYPE As Integer = 2
Public Const AVRO_TYPE As Integer = 3

' ---( Bomber Models )--------------------------------------------------------

Public Const B17_C As Integer = 1
Public Const B17_E As Integer = 2
Public Const B17_F As Integer = 3
Public Const B17_G As Integer = 4
Public Const YB40 As Integer = 5
Public Const B24_D As Integer = 6
Public Const B24_E As Integer = 7
Public Const B24_GHJ As Integer = 8
Public Const B24_LM As Integer = 9
Public Const AVRO_LANCASTER As Integer = 10

' ---( Gunnery Constants )----------------------------------------------------

Public Const HIDDEN_MG As Integer = -1
Public Const UNMANNED_MG As Integer = 0
Public Const MID_UPPER_MG As Integer = 1
Public Const NOSE_MG As Integer = 2
Public Const PORT_CHEEK_MG As Integer = 3
Public Const STBD_CHEEK_MG As Integer = 4
Public Const RADIO_ROOM_MG As Integer = 5
Public Const PORT_WAIST_MG As Integer = 6
Public Const STBD_WAIST_MG As Integer = 7
Public Const TOP_TURRET_MG As Integer = 8
Public Const BALL_TURRET_MG As Integer = 9
Public Const TAIL_MG As Integer = 10

' ---( Field of Fire Constants )----------------------------------------------

' Constants may not begin with a number (unfortunately), so a "F" is prepended
' to the clock constant name. These values existed, and were used this way,
' in the original app, but the constants are new. (Searching for "F12_HIGH"
' is easier than searching for "1".)

Public Const F12_HIGH As Integer = 1
Public Const F12_LEVEL As Integer = 2
Public Const F12_LOW As Integer = 3
Public Const F130_HIGH As Integer = 4
Public Const F130_LEVEL As Integer = 5
Public Const F130_LOW As Integer = 6
Public Const F3_HIGH As Integer = 7
Public Const F3_LEVEL As Integer = 8
Public Const F3_LOW As Integer = 9
Public Const F6_HIGH As Integer = 10
Public Const F6_LEVEL As Integer = 11
Public Const F6_LOW As Integer = 12
Public Const F9_HIGH As Integer = 13
Public Const F9_LEVEL As Integer = 14
Public Const F9_LOW As Integer = 15
Public Const F1030_HIGH As Integer = 16
Public Const F1030_LEVEL As Integer = 17
Public Const F1030_LOW As Integer = 18
Public Const VERT_CLIMB As Integer = 19
Public Const VERT_DIVE As Integer = 20

' *** Original constants ***
'Public Const NOSE As Integer = 2
'Public Const PCHEEK As Integer = 3
'Public Const SCHEEK As Integer = 4
'Public Const RADIO As Integer = 5
'Public Const PWAIST As Integer = 6
'Public Const SWAIST As Integer = 7
'Public Const TOP As Integer = 8
'Public Const BALL As Integer = 9
'Public Const TAIL As Integer = 10

'Things That Describe Gun Positions
'----------------------------------
'Ammo
'Field of Fire(20)
'Airman Manning It
'Is Airman Qualified?
'Status
'FirepowerBonus
'
'Mid-Upper (2g-16a)
'Nose (1g-15a) / Nose Turret (2g-15a or 2g-30a)
'Port Cheek (1g-10a)
'Stbd Cheek (1g-10a)
'Radio Room (1g-10a)
'Port Waist  (1g-20a or 2g-20a)
'Stbd Waist (1g-20a or 2g-20a)
'Top Turret (2g-16a or 2g-32a)
'Ball Turret (2g-20a or 2g-24a) / Ventral Blister (2g-20a) / Tunnel (1g-20a) / Floor Ring (2g-20a)
'Tail (2g-23a or 2g-25a or 4g-63a) / Stinger (2g-23a)

' ---( Visual Basic Constants )-----------------------------------------------

Public Const FILE_DIALOG_CANCEL As Integer = 32755

' ---( Crew Positions )-------------------------------------------------------

' *** Original constants ***
'Public Const PILOT As Integer = 1
'Public Const COPILOT As Integer = 2
'Public Const BOMBARDIER As Integer = 3
'Public Const NAVIGATOR As Integer = 4
'Public Const RADIO_OPERATOR As Integer = 5
'Public Const PWAIST_GUNNER As Integer = 6
'Public Const SWAIST_GUNNER As Integer = 7
'Public Const ENGINEER As Integer = 8
'Public Const BALL_GUNNER As Integer = 9
'Public Const TAIL_GUNNER As Integer = 10

Public Const HIDDEN_POSITION As Integer = -1
Public Const UNMANNED_POSITION As Integer = 0
Public Const PILOT As Integer = 1
Public Const COPILOT As Integer = 2
Public Const BOMBARDIER As Integer = 3
Public Const NAVIGATOR As Integer = 4
Public Const ENGINEER As Integer = 5
Public Const RADIO_OPERATOR As Integer = 6
Public Const NOSE_GUNNER As Integer = 7
Public Const MID_UPPER_GUNNER As Integer = 8
Public Const BALL_GUNNER As Integer = 9
Public Const PORT_WAIST_GUNNER As Integer = 10
Public Const STBD_WAIST_GUNNER As Integer = 11
Public Const TAIL_GUNNER As Integer = 12
Public Const AMMO_STOCKER As Integer = 13

' ---( General Constants )----------------------------------------------------

Public Const PORT_SIDE As Integer = 1
Public Const STBD_SIDE As Integer = 2

' ---( Damage Constants )-----------------------------------------------------

Public Const NO_LEAK As Integer = 0
Public Const LT_LEAK As Integer = 1
Public Const MED_LEAK As Integer = 2
Public Const HVY_LEAK As Integer = 3
Public Const NO_OIL As Integer = 4
Public Const PECKHAM_SCRAP_LEVEL As Integer = 400

' ---( Hit Location Constants )-----------------------------------------------

Public Const END_MISSION As Integer = -1
Public Const NO_EFFECT_HIT As Integer = 0
Public Const PORT_WING_HIT As Integer = 1
Public Const STBD_WING_HIT As Integer = 2
Public Const RADIO_ROOM_HIT As Integer = 3
Public Const NOSE_HIT As Integer = 4
Public Const FLIGHT_DECK_HIT As Integer = 5
Public Const WAIST_HIT As Integer = 6
Public Const TAIL_HIT As Integer = 7
Public Const BOMB_BAY_HIT As Integer = 8
Public Const WALKING_HITS_FUSELAGE As Integer = 9
Public Const WALKING_HITS_WINGS As Integer = 10
Public Const WALKING_HITS_BOTH As Integer = 11
                
' ---( Mission Constants )----------------------------------------------------

Public Const OUTBOUND As Integer = 1
Public Const RETURN_TRIP As Integer = 2

Public Const MIN_FIGHTERS As Integer = 1
Public Const MAX_FIGHTERS As Integer = 6

Public Const NO_WAVE As Integer = 0
Public Const MAX_WAVE As Integer = 3

Public Const BASE_ZONE As Integer = 1
Public Const MAX_ZONE As Integer = 12

Public Const NO_ALTITUDE As Integer = 0
Public Const LOW_ALTITUDE As Integer = 10000
Public Const HIGH_ALTITUDE As Integer = 25000

Public Const BURST_IN_PLANE As Integer = -1
Public Const NO_FLAK As Integer = 0
Public Const LIGHT_FLAK As Integer = 1
Public Const MEDIUM_FLAK As Integer = 2
Public Const HEAVY_FLAK As Integer = 3

Public Const NO_COVER As String = "None"
Public Const POOR_COVER As String = "Poor"
Public Const FAIR_COVER As String = "Fair"
Public Const GOOD_COVER As String = "Good"

Public Const CLEAR_WEATHER As Integer = 0
Public Const GOOD_WEATHER As Integer = 1
Public Const POOR_WEATHER As Integer = 2
Public Const BAD_WEATHER As Integer = 3
Public Const STORM_WEATHER As Integer = 4

'Public Const CLEAR_WEATHER As String = "Clear"
'Public Const GOOD_WEATHER As String = "Good"
'Public Const POOR_WEATHER As String = "Poor"
'Public Const BAD_WEATHER As String = "Bad"
'Public Const STORM_WEATHER As String = "Storm"

Public Const LOW_SQUADRON As String = "Low"
Public Const MIDDLE_SQUADRON As String = "Middle"
Public Const HIGH_SQUADRON As String = "High"

Public Const LEAD_PLANE As String = "Lead"
Public Const MIDDLE_PLANE As String = "Middle"
Public Const TAIL_PLANE As String = "Tail"

' ---( Pilot Skill Levels )---------------------------------------------------

Public Const GREEN_PILOT As Integer = 1
Public Const VET_PILOT As Integer = 2
Public Const ACE_PILOT As Integer = 3
Public Const ACE_OF_ACES As Integer = 4

' ---( Fighter Damage Constants )---------------------------------------------

Public Const NO_DAMAGE As Integer = 0
Public Const FCA_DAMAGE As Integer = 1
Public Const FBOA_DAMAGE As Integer = 2
Public Const SHOT_DOWN_DAMAGE As Integer = 4

' ---( Next Step Button Text )------------------------------------------------

Public Const REMOVE_ENEMY_FIGHTERS As String = "Remove Enemy Fighters"
Public Const FIRE_GUNS As String = "Fire Guns"
Public Const PASSING_FIRE As String = "Passing Fire"
Public Const SWAP_AMMO As String = "Swap Ammo"
Public Const BAIL_OR_CRASH As String = "Emergency"
Public Const BAILOUT_WATER As String = "Bailout (water)"
Public Const BAILOUT_LAND As String = "Bailout (land)"
Public Const BAILOUT_BASE As String = "Bailout (base)"
Public Const DITCH_WATER As String = "Ditch"
Public Const CRASH_LAND As String = "Crash"
Public Const ATTEMPT_LANDING As String = "Land"
Public Const ABORT_MISSION As String = "Abort Mission"
Public Const DESCEND_ALTITUDE As String = "Descend"
Public Const ASCEND_ALTITUDE As String = "Ascend"
Public Const JETTISON_EXCESS As String = "Jettison Excess"
Public Const FINISH_MISSION As String = "Finish Mission"
Public Const EXIT_MISSION As String = "Exit Mission"
Public Const QUIT_BEFORE_WATER As String = "Quit Before Water"

'---( lbl tags )---------------

Public Const REMOVE_FIGHTER As String = "Remove Fighter"
Public Const FIGHTER_MISSED As String = "Fighter Missed"

' ---( Spray Fire )-----------------------------------------------------------

Public Const SPRAY_FIRE_JAM As Integer = 1
Public Const SPRAY_FIRE_NOEFFECT As Integer = 2
Public Const SPRAY_FIRE_BREAKOFF As Integer = 3
Public Const SPRAY_FIRE_HIT As Integer = 4

' ---( Ammo Constants )-------------------------------------------------------

Public Const SINGLE_GUN_AMMO As Integer = 1
Public Const TWIN_GUN_AMMO As Integer = 2

' ---( Gun States )-----------------------------------------------------------

Public Const MG_OKAY As Integer = 0
Public Const MG_JAMMED As Integer = 1
Public Const MG_INOPERABLE As Integer = 2



