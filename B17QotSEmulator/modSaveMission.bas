Attribute VB_Name = "modSaveMission"
'******************************************************************************
' modSaveMission.bas
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
' CreateMissionHTML
'
' INPUT:  n/a
'
' OUTPUT: HTML header, body and footer blocks.
'
' RETURN: n/a
'
' NOTES:  Dynamically assemble the HTML based on the mission parameters.
'******************************************************************************
Public Sub CreateMissionHTML(ByRef strHeader As String, ByRef strBody As String, ByRef strFooter As String)
    Dim intCount As Integer
    Dim strCover As String
    Dim strSign As String

    With frmMainMenu
    
        ' Header
        
        strHeader = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.01 Transitional//EN" & Chr(34) & " " & Chr(34) & "http://www.w3.org/TR/html4/loose.dtd" & Chr(34) & ">" & vbCrLf & _
                    "<html>" & vbCrLf & _
                    "<head>" & vbCrLf & _
                    vbTab & "<meta http-equiv=" & Chr(34) & "content-type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=iso-8859-1" & Chr(34) & ">" & vbCrLf & _
                    vbTab & "<meta name=" & Chr(34) & "keywords" & Chr(34) & " content=" & Chr(34) & "B-17 B-24 Lancaster Queen Skies Flying Fortress Boxcar wargame emulator World War II WWII aerial strategic bombardment Europe 8th Air Force England Germany flight simulator" & Chr(34) & ">" & vbCrLf & _
                    vbTab & "<meta name=" & Chr(34) & "description" & Chr(34) & " content=" & Chr(34) & "B-17: Queen of the Skies Emulator" & Chr(34) & ">" & vbCrLf & _
                    vbTab & "<meta name=" & Chr(34) & "author" & Chr(34) & " content=" & Chr(34) & "B17QotS Emulator by Preston V. McMurry III" & Chr(34) & ">" & vbCrLf & _
                    vbTab & "<meta name=" & Chr(34) & "copyright" & Chr(34) & " content=" & Chr(34) & "B17QotS Emulator &copy; Copyright 2004 Preston V. McMurry III" & Chr(34) & ">" & vbCrLf & _
                    vbTab & "<title>B17QotS Target: " & Mission.TargetName & "</title>" & vbCrLf & _
                    "</head>" & vbCrLf & _
                    "<body>" & vbCrLf & _
                    vbCrLf & _
                    "<center>" & vbCrLf & _
                    vbTab & "<font size=" & Chr(34) & "-1" & Chr(34) & ">" & Chr(34) & "B-17: Queen of the Skies" & Chr(34) & " Emulator</font><br />" & vbCrLf & _
                    vbTab & "<font size=" & Chr(34) & "+4" & Chr(34) & " face=" & Chr(34) & "stencil" & Chr(34) & ">Target: " & Mission.TargetName & "</font><p />" & vbCrLf & _
                    vbTab & .cboMonth.Text & ", " & .cboYear.Text & vbCrLf & _
                    "</center>" & vbCrLf & _
                    vbCrLf & _
                    "<p />" & vbCrLf & _
                    "<hr style=" & Chr(34) & "COLOR: #000000" & Chr(34) & " />" & vbCrLf & _
                    "<p />" & vbCrLf
              
        ' Body
        
        strBody = "<b>Bomber:</b> " & Bomber.Name & " (" & .cboBomberModel(MISSION_TAB).Text & ")" & "<br />" & vbCrLf & _
                  "<b>Target:</b> " & Mission.TargetName & "<br />" & vbCrLf & _
                  "<b>Type:</b> " & Mission.TargetType & "<p />" & vbCrLf & vbCrLf & _
                  "<b>Squadron:</b> " & Bomber.SquadronPos & "<br />" & vbCrLf & _
                  "<b>Formation:</b> " & Bomber.FormationPos
                  
        If Bomber.FormationPos = LEAD_PLANE Then
            strBody = strBody & " (add 1 Me-109/12 High if fighters appear)" & "<p />" & vbCrLf & vbCrLf
        ElseIf Bomber.FormationPos = TAIL_PLANE Then
            strBody = strBody & " (add 1 Me-109/6 High if fighters appear)" & "<p />" & vbCrLf & vbCrLf
        Else ' Middle
            strBody = strBody & "<p />" & vbCrLf & vbCrLf
        End If
        
        strBody = strBody & "<b>Weather Over Target:</b> " & WeatherText(Mission.Zone(Mission.TargetZone).Weather) & "<br />" & vbCrLf & _
                            "<b>Flak Over Target:</b> " & O2FlakOverTarget(True) & "<br />" & vbCrLf & _
                            "<b>Weather Over Base:</b> " & WeatherText(Mission.Zone(BASE_ZONE).Weather) & "<p />" & vbCrLf & vbCrLf & _
                            "<b>Options:</b><br />" & vbCrLf & _
                            "<ul>" & vbCrLf

        ' Options
        
        If Mission.Options.RandomEvents = True Then
            strBody = strBody & vbTab & "<li />Random Events" & vbCrLf
        End If

        If Mission.Options.MechanicalFailures = True Then
            strBody = strBody & vbTab & "<li />Mechanical Failures" & vbCrLf
        End If

        If Mission.Options.TimePeriodSpecificFormations = True Then
            strBody = strBody & vbTab & "<li />Time Period Specific Formations" & vbCrLf
        End If

        If Mission.Options.FormationDefensiveGunnery = True Then
            strBody = strBody & vbTab & "<li />Formation Defensive Gunnery" & vbCrLf
        End If

        If Mission.Options.EvadeFlak = True Then
            strBody = strBody & vbTab & "<li />Evade Flak" & vbCrLf
        End If

        If Mission.Options.CrewExperience = True Then
            strBody = strBody & vbTab & "<li />Crew Experience" & vbCrLf
        End If

        If Mission.Options.AlternateWeather = True Then
            strBody = strBody & vbTab & "<li />Alternate Weather" & vbCrLf
        End If

        If Mission.Options.GermanFighterPilotSkill = True Then
            strBody = strBody & vbTab & "<li />German Fighter Pilot Skill" & vbCrLf
        End If

        If Mission.Options.JG26StationedInAbbeville = True Then
            strBody = strBody & vbTab & "<li />JG26 Stationed In Abbeville" & vbCrLf
        End If

        If Mission.Options.Ju88sUsedAsFighters = True Then
            strBody = strBody & vbTab & "<li />Ju88s Used As Fighters" & vbCrLf
        End If

        If Mission.Options.Unescorted = True Then
            strBody = strBody & vbTab & "<li />Unescorted" & vbCrLf
        End If

        If Mission.Options.RedTailAngels = True Then
            strBody = strBody & vbTab & "<li />Red Tail Angels" & vbCrLf
        End If

        strBody = strBody & "</ul><p />" & vbCrLf & vbCrLf & _
                            "<table border=" & Chr(34) & "1" & Chr(34) & " cellpadding=" & Chr(34) & "5" & Chr(34) & ">" & vbCrLf & _
                            vbTab & "<tr>" & vbCrLf & _
                            vbTab & vbTab & "<td><b>Fighter Cover</b></td>" & vbCrLf

        ' Fighter Cover Table
        
        For intCount = 2 To Mission.TargetZone
                strBody = strBody & vbTab & vbTab & "<td bgcolor=" & Chr(34) & "#cccccc" & Chr(34) & ">Zone " & intCount & "</td>" & vbCrLf
        Next intCount

        strBody = strBody & vbTab & "</tr>" & vbCrLf & _
                            vbTab & "<tr>" & vbCrLf & _
                            vbTab & vbTab & "<td>Outbound</td>" & vbCrLf

        For intCount = 2 To Mission.TargetZone
            If Mission.Zone(intCount).CoverOut = NO_COVER Then
                strCover = ""
            Else
                strCover = Mission.Zone(intCount).CoverOut
            End If
                
            strBody = strBody & vbTab & vbTab & "<td>" & strCover & "</td>" & vbCrLf
        Next intCount

        strBody = strBody & vbTab & "</tr>" & vbCrLf & _
                            vbTab & "<tr>" & vbCrLf & _
                            vbTab & vbTab & "<td>Inbound</td>" & vbCrLf

        For intCount = 2 To Mission.TargetZone
            If Mission.Zone(intCount).CoverBack = NO_COVER Then
                strCover = ""
            Else
                strCover = Mission.Zone(intCount).CoverBack
            End If
                
            strBody = strBody & vbTab & vbTab & "<td>" & strCover & "</td>" & vbCrLf
        Next intCount
    
        strBody = strBody & vbTab & "</tr>" & vbCrLf & _
                            "</table>" & vbCrLf & vbCrLf & _
                            "<p />" & vbCrLf & vbCrLf
              
        ' Enemy Waves Table
        
        strBody = strBody & "<table border=" & Chr(34) & "1" & Chr(34) & " cellpadding=" & Chr(34) & "5" & Chr(34) & ">" & vbCrLf & _
                            vbTab & "<tr>" & vbCrLf & _
                            vbTab & vbTab & "<td><b>Enemy Waves</b></td>" & vbCrLf

        For intCount = 2 To Mission.TargetZone
                strBody = strBody & vbTab & vbTab & "<td bgcolor=" & Chr(34) & "#cccccc" & Chr(34) & ">Zone " & intCount & "</td>" & vbCrLf
        Next intCount

        strBody = strBody & vbTab & "</tr>" & vbCrLf & _
                            vbTab & "<tr>" & vbCrLf & _
                            vbTab & vbTab & "<td></td>" & vbCrLf

        For intCount = 2 To Mission.TargetZone
            If Mission.Zone(intCount).Modifier > 0 Then
                strSign = "+"
            Else
                strSign = ""
            End If
                
            strBody = strBody & vbTab & vbTab & "<td>" & strSign & Mission.Zone(intCount).Modifier & " / " & Mission.Zone(intCount).Terrain & "</td>" & vbCrLf
        Next intCount
    
        strBody = strBody & vbTab & "</tr>" & vbCrLf & _
                            "</table>" & vbCrLf & vbCrLf & _
                            "<p />" & vbCrLf & vbCrLf
              
        ' Footer
        
        strFooter = "<p />" & vbCrLf & _
                    "<hr style=" & Chr(34) & "COLOR: #000000" & Chr(34) & " />" & vbCrLf & _
                    "<p />" & vbCrLf & _
                    vbCrLf & _
                    "</body>" & vbCrLf & _
                    "</html>"
    
    End With

End Sub

