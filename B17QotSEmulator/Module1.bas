'******************************************************************************
' Module1.bas
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

Attribute VB_Name = "Module1"
Option Explicit

Sub BomberBuildData()
    
    Dim intRoll As Integer
    Dim strManufacturer As String
    Dim strPlant As String
    Dim strTailNumber As String

    strManufacturer = "Unknown"
    strPlant = "Unknown"
    strTailNumber = ""

    If prsBomber![BomberModel] = B17_C Then

        ' Built = 38
        ' 40-2042 to 40-2079

        strManufacturer = "Boeing"
        strPlant = "Seattle, WA"

        ' 40-2042 to 40-2079
        strTailNumber = "40-" & CStr(2041 + RandomDX(38)) & "-BO"
        
    ElseIf prsBomber![BomberModel] = B17_E Then

        ' Built = 512
        ' 41-2393 to 41-2669 and 41-9011 to 41-9245
        
        strManufacturer = "Boeing"
        strPlant = "Seattle, WA"

        intRoll = RandomDX(512)
        
        If intRoll <= 277 Then
            
            ' 41-2393 to 41-2669
            strTailNumber = "41-" & CStr(2392 + RandomDX(277)) & "-BO"
        
        ElseIf intRoll >= 278 _
        And intRoll <= 512 Then
            
            ' 41-9011 to 41-9245
            strTailNumber = "41-" & CStr(9010 + RandomDX(235)) & "-BO"
        
        End If

    ElseIf prsBomber![BomberModel] = B17_F Then

        ' Built = 3406
        '
        ' Boeing production: 41-24340 to 24639; 42-5050 to 42-5484; 42-29467 to 42-31031
        ' Douglas production: 42-2964 to 42-3562; 42-37714 to 42-37720
        ' Lockheed-Vega production: 42-5705 to 42-6204
        '
        ' 2300 aircraft built by Boeing
        ' 606 aircraft built by Douglas
        ' 500 aircraft built by Lockheed-Vega

        intRoll = RandomDX(3406)
        
        If intRoll <= 2300 Then

            strManufacturer = "Boeing"
            strPlant = "Seattle, WA"

            ' Built = 2300
            
            intRoll = RandomDX(2300)
            
            If intRoll <= 300 Then

                ' 41-24340 to 24639
                strTailNumber = "41-" & CStr(24339 + RandomDX(300)) & "-BO"

            ElseIf intRoll >= 301 _
            And intRoll <= 735 Then

                ' 42-5050 to 42-5484
                strTailNumber = "42-" & CStr(5049 + RandomDX(435)) & "-BO"

            ElseIf intRoll >= 736 _
            And intRoll <= 2300 Then
                
                ' 42-29467 to 42-31031
                strTailNumber = "42-" & CStr(29466 + RandomDX(1565)) & "-BO"
            
            End If

        ElseIf intRoll >= 2301 _
        And intRoll <= 2906 Then

            strManufacturer = "Douglas"
            strPlant = "Long Beach, CA"

            ' Built = 606
            
            intRoll = RandomDX(606)
            
            If intRoll <= 599 Then
                
                ' 42-2964 to 42-3562
                strTailNumber = "42-" & CStr(2963 + RandomDX(599)) & "-DL"
            
            ElseIf intRoll >= 600 _
            And intRoll <= 606 Then
                
                ' 42-37714 to 42-37720
                strTailNumber = "42-" & CStr(37713 + RandomDX(7)) & "-DL"
            
            End If
    
        ElseIf intRoll >= 2907 _
        And intRoll <= 3406 Then

            strManufacturer = "Lockheed-Vega"
            strPlant = "Burbank, CA"

            ' Built = 500
            
            ' 42-5705 to 42-6204
            strTailNumber = "42-" & CStr(5704 + RandomDX(500)) & "-VE"

        End If

    ElseIf prsBomber![BomberModel] = B17_G Then

        ' Built = 8679
        '
        ' Boeing production: 42-31032 to 42-32116; 42-97058 to 42-97407; 42-102379 to 42-102978; 43-37509 to 43-39508
        ' Douglas production: 42-37716; 42-37721 to 42-38213; 42-106984 to 42-107233; 44-6001 to 44-7000; 44-83236 to 44-83885
        ' Lockheed-Vega production: 42-39758 to 42-40057; 42-97436 to 42-9798035; 44-8001 to 44-9000; 44-85492 to 44-85841
        '
        ' 4035 aircraft built by Boeing
        ' 2394 aircraft built by Douglas
        ' 2250 aircraft built by Lockheed-Vega

        intRoll = RandomDX(8679)
        
        If intRoll <= 4035 Then

            strManufacturer = "Boeing"
            strPlant = "Seattle, WA"

            ' Built = 4035
            
            intRoll = RandomDX(4035)
            
            If intRoll <= 1085 Then

                ' 42-31032 to 42-32116
                strTailNumber = "42-" & CStr(31031 + RandomDX(1085)) & "-BO"

            ElseIf intRoll >= 1086 _
            And intRoll <= 2435 Then

                ' 42-97058 to 42-97407
                strTailNumber = "42-" & CStr(97057 + RandomDX(350)) & "-BO"

            ElseIf intRoll >= 2436 _
            And intRoll <= 3035 Then

                ' 42-102379 to 42-102978
                strTailNumber = "42-" & CStr(102378 + RandomDX(600)) & "-BO"

            ElseIf intRoll >= 3036 _
            And intRoll <= 4035 Then

                ' 43-37509 to 43-39508
                strTailNumber = "43-" & CStr(37508 + RandomDX(2000)) & "-BO"

            End If

        ElseIf intRoll >= 4306 _
        And intRoll <= 6429 Then
        
            strManufacturer = "Douglas"
            strPlant = "Long Beach, CA"

            ' Built = 2394
            
            intRoll = RandomDX(2394)
            
            If intRoll = 1 Then

                ' 42-37716
                strTailNumber = "42-37716-DL"

            ElseIf intRoll >= 2 _
            And intRoll <= 494 Then

                ' 42-37721 to 42-38213
                strTailNumber = "42-" & CStr(37720 + RandomDX(493)) & "-DL"

            ElseIf intRoll >= 495 _
            And intRoll <= 744 Then

                ' 42-106984 to 42-107233
                strTailNumber = "42-" & CStr(106983 + RandomDX(250)) & "-DL"

            ElseIf intRoll >= 745 _
            And intRoll <= 1744 Then

                ' 44-6001 to 44-7000
                strTailNumber = "44-" & CStr(6000 + RandomDX(1000)) & "-DL"

            ElseIf intRoll >= 1745 _
            And intRoll <= 2394 Then

                ' 44-83236 to 44-83885
                strTailNumber = "44-" & CStr(83235 + RandomDX(650)) & "-DL"

            End If
    
        ElseIf intRoll >= 6430 _
        And intRoll <= 8679 Then

            strManufacturer = "Lockheed-Vega"
            strPlant = "Burbank, CA"

            ' Built = 2250
            
            intRoll = RandomDX(2250)
            
            If intRoll <= 300 Then

                ' 42-39758 to 42-40057
                strTailNumber = "42-" & CStr(39757 + RandomDX(300)) & "-VE"

            ElseIf intRoll >= 301 _
            And intRoll <= 900 Then

                ' 42-97436 to 42-98035
                strTailNumber = "42-" & CStr(97435 + RandomDX(600)) & "-VE"

            ElseIf intRoll >= 901 _
            And intRoll <= 1900 Then

                ' 44-8001 to 44-9000
                strTailNumber = "44-" & CStr(8000 + RandomDX(1000)) & "-VE"

            ElseIf intRoll >= 1901 _
            And intRoll <= 2250 Then

                ' 44-85492 to 44-85841
                strTailNumber = "44-" & CStr(85491 + RandomDX(350)) & "-VE"

            End If

        End If

    ElseIf prsBomber![BomberModel] = YB40 Then
    
        ' 42-5732 to 42-5744 VE
        ' 42-5871 VE
        ' 42-5920 to 42-5921 VE
        ' 42-5923 to 42-5925 VE
        ' 42-5927 VE
    
        strManufacturer = "Lockheed-Vega"
        strPlant = "Burbank, CA"

        ' B-17Fs converted = 20
            
        intRoll = RandomDX(20)
        
        If intRoll <= 13 Then
        
            ' 42-5732 to 42-5744
            strTailNumber = "42-" & CStr(5731 + RandomDX(13)) & "-VE"
        
        ElseIf intRoll >= 14 Then

            ' 42-5871
            strTailNumber = "42-5871-VE"
        
        ElseIf intRoll >= 15 _
        And intRoll <= 16 Then
        
            ' 42-5920 to 42-5921
            strTailNumber = "42-" & CStr(5919 + RandomDX(2)) & "-VE"
        
        ElseIf intRoll >= 17 _
        And intRoll <= 19 Then
        
            ' 42-5923 to 42-5925
            strTailNumber = "42-" & CStr(5922 + RandomDX(3)) & "-VE"
        
        ElseIf intRoll >= 20 Then
        
            ' 42-5927
            strTailNumber = "42-5927-VE"
        
        End If
        
    ElseIf prsBomber![BomberModel] = B24_D Then

        ' Built = 2696
        '
        ' 2,381 built by Consolidated at the San Diego, California plant.
        ' 305 built by Consolidated at the Fort Worth, Texas plant.
        ' 10 built by Douglas at the Tulsa Oklahoma plant.
        '
        ' 40-2349 to 40-2368 (CO) -- first 20 B-24Ds ever
        ' 41-1087 to 41-1142 CO
        ' 41-11598 (CO)
        ' 41-23640 to 41-23755 CO
        ' 41-23759 to 41-24339 CO
        ' 42-40058 to 42-41257 CO
        ' 42-72765 to 42-72963 CO
        ' 41-23756 to 41-23758 DT
        ' 41-29074 and 41-29075 (CF) -- first two CF aircraft
        ' 42-63752 to 42-64046 CF

        intRoll = RandomDX(2696)
        
        If intRoll <= 2381 Then

            strManufacturer = "Consolidated"
            strPlant = "San Diego, CA"
            
            ' Built = 2381
            ' Missing serial numbers = 208

            intRoll = RandomDX(2173)
            
            If intRoll <= 20 Then

                ' 40-2349 to 40-2368
                strTailNumber = "40-" & CStr(2348 + RandomDX(20)) & "-CO"

            ElseIf intRoll >= 21 _
            And intRoll <= 76 Then

                ' 41-1087 to 41-1142
                strTailNumber = "41-" & CStr(1086 + RandomDX(56)) & "-CO"

            ElseIf intRoll = 77 Then

                ' 41-11598
                strTailNumber = "41-11598-CO"

            ElseIf intRoll >= 78 _
            And intRoll <= 193 Then

                ' 41-23640 to 41-23755
                strTailNumber = "41-" & CStr(23639 + RandomDX(116)) & "-CO"

            ElseIf intRoll >= 194 _
            And intRoll <= 774 Then

                ' 41-23759 to 41-24339
                strTailNumber = "41-" & CStr(23758 + RandomDX(581)) & "-CO"

            ElseIf intRoll >= 775 _
            And intRoll <= 1974 Then

                ' 42-40058 to 42-41257
                strTailNumber = "42-" & CStr(40057 + RandomDX(1200)) & "-CO"

            ElseIf intRoll >= 1975 _
            And intRoll <= 2173 Then

                ' 42-72765 to 42-72963
                strTailNumber = "42-" & CStr(72764 + RandomDX(199)) & "-CO"

            End If

        ElseIf intRoll >= 2382 _
        And intRoll <= 2686 Then
        
            strManufacturer = "Consolidated"
            strPlant = "Ft. Worth, TX"

            ' Built = 305
            ' Missing serial numbers = 8

            intRoll = RandomDX(297)
            
            If intRoll <= 2 Then

                ' 41-29074 to 41-29075
                strTailNumber = "41-" & CStr(29073 + RandomDX(2)) & "-CF"

            ElseIf intRoll >= 3 _
            And intRoll <= 297 Then

                ' 42-63752 to 42-64046
                strTailNumber = "42-" & CStr(63751 + RandomDX(295)) & "-CF"

            End If

        ElseIf intRoll >= 2687 _
        And intRoll <= 2696 Then
        
            strManufacturer = "Douglas"
            strPlant = "Tulsa, OK"

            ' Built = 10
            ' Missing serial numbers = 7

            ' 41-23756 to 41-23758
            strTailNumber = "41-" & CStr(23755 + RandomDX(3)) & "-DT"

        End If
        
    ElseIf prsBomber![BomberModel] = B24_E Then

        ' Built = 801
        '
        ' 490 B-24E-FO built by Ford at the Willow plant in 6 production blocks.
        ' 165 B-24E-DT built by Douglas at the Tulsa plant using Ford sub-assemblies (5 blocks)
        ' 146 B-24E-CF built by Consolidated at the Fort Worth plant using Ford sub-assemblies (4 blocks).
        '
        ' 41-28409 to 41-28573; 41-29007 to 41-29115; 42-6976 to 42-7464; 42-7770; 42-64395 to 42-64431;

        intRoll = RandomDX(801)
        
        If intRoll <= 490 Then

            strManufacturer = "Ford"
            strPlant = "Willow Run, MI"

            ' Built = 490
            
            intRoll = RandomDX(490)
            
            If intRoll <= 489 Then

                ' 42-6976 to 42-7464
                strTailNumber = "42-" & CStr(6975 + intRoll) & "-FO"

            ElseIf intRoll = 490 Then

                ' 42-7770
                strTailNumber = "42-7770-FO"

            End If

        ElseIf intRoll >= 491 _
        And intRoll <= 655 Then
        
            strManufacturer = "Douglas"
            strPlant = "Tulsa, OK"

            ' Built = 165
            
            ' 41-28409 to 41-28573
            strTailNumber = "41-" & CStr(28408 + RandomDX(165)) & "-DT"

        ElseIf intRoll >= 656 _
        And intRoll <= 801 Then

            strManufacturer = "Consolidated"
            strPlant = "Ft. Worth, TX"

            ' Built = 146
            
            intRoll = RandomDX(146)
            
            If intRoll <= 109 Then

                ' 41-29007 to 41-29115
                strTailNumber = "41-" & CStr(29006 + RandomDX(109)) & "-CF"

            ElseIf intRoll >= 110 _
            And intRoll <= 146 Then

                ' 42-64395 to 42-64431
                strTailNumber = "42-" & CStr(64394 + RandomDX(37)) & "-CF"

            End If

        End If

    ElseIf prsBomber![BomberModel] = B24_GHJ Then
        
        ' The B-24G, B-24H and B-24J were the same basic model. Because the
        ' emulator treats them the same, we need to break out their actual
        ' serial numbers.
        
        intRoll = RandomDX(10208)

        If intRoll <= 430 Then
        
            ' B-24G
            '
            ' Built = 430
            '
            ' 25 built as B-24G (42-78045 to 42-78069)
            ' 405 built as improved B-24G-1 (block numbers 1, 5, 10, 15, and 16) (42-78070 and later)
            '
            ' Serial numbers: 42-78045 to 42-78474

            strManufacturer = "North American"
            strPlant = "Dallas, TX"

            ' 42-78045 to 42-78474 NT
            strTailNumber = "42-" & CStr(78044 + RandomDX(430)) & "-NT"

        ElseIf intRoll >= 431 _
        And intRoll <= 3530 Then

            ' B-24H
            '
            ' Built = 3100
            '
            ' 1780 built by Ford at their Willow Run plant in 7 production blocks
            ' 582 assembled by Douglas at their Tulsa plant.
            ' 738 assembled by Consolidated at their Fort Worth plant.
            '
            ' Serial numbers: 41-28574 to 41-29006; 41-29116 to 41-29608; 42-50277 to 42-50451; 42-51077 to 42-51225; 42-52077 to 42-52776; 42-64432 to 42-64501; 42-7465 to 42-7769; 42-94729 to 42-95503
        
            intRoll = RandomDX(3100)
        
            If intRoll <= 1780 Then
    
                strManufacturer = "Ford"
                strPlant = "Willow Run, MI"
    
                ' Built = 1780
                
                intRoll = RandomDX(1780)
                
                If intRoll <= 305 Then
    
                    ' 42-7465 to 42-7769
                    strTailNumber = "42-" & CStr(7464 + RandomDX(305)) & "-FO"
    
                ElseIf intRoll >= 306 _
                And intRoll <= 1005 Then
    
                    ' 42-52077 to 42-52776
                    strTailNumber = "42-" & CStr(52076 + RandomDX(700)) & "-FO"
    
                ElseIf intRoll >= 1006 _
                And intRoll <= 1780 Then
    
                    ' 42-94729 to 42-95503
                    strTailNumber = "42-" & CStr(94728 + RandomDX(775)) & "-FO"
    
                End If
    
            ElseIf intRoll >= 1781 _
            And intRoll <= 2362 Then
    
                strManufacturer = "Douglas"
                strPlant = "Tulsa, OK"

                ' Built = 582
                
                intRoll = RandomDX(582)
                
                If intRoll <= 433 Then
    
                    ' 41-28574 to 41-29006
                    strTailNumber = "41-" & CStr(28573 + RandomDX(433)) & "-DT"
    
                ElseIf intRoll >= 434 _
                And intRoll <= 582 Then
    
                    ' 42-51077 to 42-51225
                    strTailNumber = "42-" & CStr(51076 + RandomDX(149)) & "-DT"
    
                End If
    
            ElseIf intRoll >= 2363 _
            And intRoll <= 3100 Then
            
                strManufacturer = "Consolidated"
                strPlant = "Ft. Worth, TX"
    
                ' Built = 738
                
                intRoll = RandomDX(738)
                
                If intRoll <= 493 Then
    
                    ' 41-29116 to 41-29608
                    strTailNumber = "41-" & CStr(29115 + RandomDX(493)) & "-CF"
    
                ElseIf intRoll >= 494 _
                And intRoll <= 668 Then
    
                    ' 42-50277 to 42-50451
                    strTailNumber = "42-" & CStr(50276 + RandomDX(175)) & "-CF"
    
                ElseIf intRoll >= 669 _
                And intRoll <= 7380 Then
    
                    ' 42-64432 to 42-64501
                    strTailNumber = "42-" & CStr(64431 + RandomDX(70)) & "-CF"
    
                End If
    
            End If
    
        ElseIf intRoll >= 3531 _
        And intRoll <= 10208 Then

            ' B-24J
            '
            ' Built = 6678
            '
            ' 2792 built by Consolidated at their San Diego plant in 43 production blocks.
            ' 1587 built by Ford at their Willow Run plant in 4 production blocks.
            ' 1558 built by Consolidated at their Fort Worth plant in 25 production blocks.
            ' 536 built by North American at their Dallas plant in 3 production blocks.
            ' 205 built by Douglas at their Tulsa plant in 3 production blocks.
            '
            ' 42-50452 to 42-50508 CF = 57
            ' 42-50509 to 42-51076 FO = 568
            ' 42-51431 to 42-52076 FO = 646
            ' 42-51226 to 42-51430 DT = 205
            ' 42-64047 to 42-64394 CF = 348
            ' 42-72964 to 42-73514 CO = 551
            ' 42-78475 to 42-78794 NT = 320
            ' 42-95504 to 42-95628 FO = 125
            ' 42-99736 to 42-99935 CF = 200
            ' 42-99936 to 42-100435 CO = 500
            ' 42-109789 to 42-110188 CO = 400
            ' 44-10253 to 44-10302 CO = 50
            ' 44-10303 to 44-10752 CF = 450
            ' 44-28061 to 44-28276 NT = 216
            ' 44-28277 to 44-28710 Cancelled contract
            ' 44-40049 to 44-41389 CO = 341
            ' 44-44049 to 44-44501 CF = 453
            ' 44-48754 to 44-49001 FO = 247

        ' 44-28277 to 44-28710 Cancelled contract

            intRoll = RandomDX(6678)
        
            If intRoll <= 1587 Then
    
                strManufacturer = "Ford"
                strPlant = "Willow Run, MI"
    
                ' Built = 1587
                
                intRoll = RandomDX(1587)
        
                If intRoll <= 568 Then
    
                    ' 42-50509 to 42-51076 FO = 568
                    strTailNumber = "42-" & CStr(50508 + RandomDX(568)) & "-FO"
    
                ElseIf intRoll >= 569 _
                And intRoll <= 1214 Then
    
                    ' 42-51431 to 42-52076 FO = 646
                    strTailNumber = "42-" & CStr(51430 + RandomDX(646)) & "-FO"
    
                ElseIf intRoll >= 1215 _
                And intRoll <= 1339 Then
    
                    ' 42-95504 to 42-95628 FO = 125
                    strTailNumber = "42-" & CStr(95503 + RandomDX(125)) & "-FO"
    
                ElseIf intRoll >= 1340 _
                And intRoll <= 1587 Then
    
                    ' 44-48754 to 44-49001 FO = 248 = 1587x
                    strTailNumber = "44-" & CStr(48753 + RandomDX(248)) & "-FO"
    
                End If
    
            ElseIf intRoll >= 1588 _
            And intRoll <= 1792 Then
    
                strManufacturer = "Douglas"
                strPlant = "Tulsa, OK"
    
                ' Built = 205
                
                ' 42-51226 to 42-51430
                strTailNumber = "42-" & CStr(51225 + RandomDX(205)) & "-DT"

            ElseIf intRoll >= 1793 _
            And intRoll <= 4584 Then
    
                strManufacturer = "Consolidated"
                strPlant = "San Diego, CA"
    
                ' Built = 2792
                ' Missing serial numbers = 950

                intRoll = RandomDX(1842)
        
                If intRoll <= 551 Then
    
                    ' 42-72964 to 42-73514 CO = 551
                    strTailNumber = "42-" & CStr(72963 + RandomDX(551)) & "-CO"
    
                ElseIf intRoll >= 552 _
                And intRoll <= 1051 Then
    
                    ' 42-99936 to 42-100435 CO = 500
                    strTailNumber = "42-" & CStr(99935 + RandomDX(500)) & "-CO"
    
                ElseIf intRoll >= 1052 _
                And intRoll <= 1451 Then
    
                    ' 42-109789 to 42-110188 CO = 400
                    strTailNumber = "42-" & CStr(109788 + RandomDX(400)) & "-CO"
    
                ElseIf intRoll >= 1452 _
                And intRoll <= 1501 Then
    
                    ' 44-10253 to 44-10302 CO = 50
                    strTailNumber = "44-" & CStr(10252 + RandomDX(50)) & "-CO"
    
                ElseIf intRoll >= 1502 _
                And intRoll <= 1842 Then
    
                    ' 44-40049 to 44-41389 CO = 341 = 1842 (950 missing)
                    strTailNumber = "44-" & CStr(40048 + RandomDX(341)) & "-CO"
    
                End If
    
            ElseIf intRoll >= 4585 _
            And intRoll <= 6142 Then
    
                strManufacturer = "Consolidated"
                strPlant = "Ft. Worth, TX"
    
                ' Built = 1558
                ' Missing serial numbers = 50

                intRoll = RandomDX(1508)
        
                If intRoll <= 57 Then
    
                    ' 42-50452 to 42-50508 CF = 57
                    strTailNumber = "42-" & CStr(50451 + RandomDX(57)) & "-CF"
    
                ElseIf intRoll >= 58 _
                And intRoll <= 405 Then
    
                    ' 42-64047 to 42-64394 CF = 348
                    strTailNumber = "42-" & CStr(64046 + RandomDX(348)) & "-CF"
    
                ElseIf intRoll >= 406 _
                And intRoll <= 605 Then
    
                    ' 42-99736 to 42-99935 CF = 200
                    strTailNumber = "42-" & CStr(99735 + RandomDX(200)) & "-CF"
    
                ElseIf intRoll >= 606 _
                And intRoll <= 1055 Then
    
                    ' 44-10303 to 44-10752 CF = 450
                    strTailNumber = "44-" & CStr(10302 + RandomDX(450)) & "-CF"
    
                ElseIf intRoll >= 1056 _
                And intRoll <= 1508 Then
    
                    ' 44-44049 to 44-44501 CF = 453 = 1310 (50 missing)
                    strTailNumber = "44-" & CStr(44048 + RandomDX(453)) & "-CF"
    
                End If
    
            ElseIf intRoll >= 6143 _
            And intRoll <= 6678 Then
    
                strManufacturer = "North American"
                strPlant = "Dallas, TX"
    
                ' Built = 536
                
                intRoll = RandomDX(536)
        
                If intRoll <= 320 Then
    
                    ' 42-78475 to 42-78794 NT = 320
                    strTailNumber = "42-" & CStr(78474 + RandomDX(320)) & "-NT"
    
                ElseIf intRoll >= 321 _
                And intRoll <= 536 Then
    
                    ' 44-28061 to 44-28276 NT = 216 = 536x
                    strTailNumber = "44-" & CStr(28060 + RandomDX(216)) & "-NT"
    
                End If
    
            End If
        
        End If

    ElseIf prsBomber![BomberModel] = B24_LM Then
        
        ' The B-24L and B-24M were the same basic model. Because the
        ' emulator treats them the same, we need to break out their actual
        ' serial numbers.
        
        intRoll = RandomDX(4384)

        If intRoll <= 1667 Then
        
            ' B-24J
            '
            ' Built = 1667
            '
            ' Ford built 1,250 aircraft at its Willow Run plant.
            ' Consolidated built 417 aircraft at its San Diego plant.
            '
            ' Serial numbers: 44-41390 to 44-41806; 44-49002 to 44-50251

' 44-41390 to 44-41806 CO = 417
' 44-49002 to 44-50251 FO = 1250

            intRoll = RandomDX(1667)
        
            If intRoll <= 1250 Then
    
                strManufacturer = "Ford"
                strPlant = "Willow Run, MI"
    
                ' Built = 1250
                ' 44-49002 to 44-50251 FO = 1250
                strTailNumber = "44-" & CStr(49001 + RandomDX(1250)) & "-FO"
    
            ElseIf intRoll >= 1251 _
            And intRoll <= 1667 Then
            
                strManufacturer = "Consolidated"
                strPlant = "San Diego, CA"
    
                ' Built = 417
                ' 44-41390 to 44-41806 CO = 417
                strTailNumber = "44-" & CStr(41389 + RandomDX(417)) & "-FO"
        
            End If
        
        ElseIf intRoll >= 1668 _
        And intRoll <= 4384 Then
    
            ' B-24M
            '
            ' Built = 2717
            '
            ' Serial numbers: 44-41807 to 44-42722; 44-50252 to 44-52052
            '
            ' Ford built 1,801 aircraft at its Willow Run plant.
            ' Consolidated built 916 aircraft at its San Diego plant.
' 2717
' 44-41807 to 44-42722 CO = 916
' 44-50252 to 44-52052 FO = 1801

            intRoll = RandomDX(2717)
        
            If intRoll <= 1801 Then
    
                strManufacturer = "Ford"
                strPlant = "Willow Run, MI"
    
                ' Built = 1801
                ' 44-50252 to 44-52052 FO = 1801
                strTailNumber = "44-" & CStr(50251 + RandomDX(1801)) & "-FO"
    
            ElseIf intRoll >= 1802 _
            And intRoll <= 2717 Then
            
                strManufacturer = "Consolidated"
                strPlant = "San Diego, CA"
    
                ' Built = 916
                ' 44-41807 to 44-42722 CO = 916
                strTailNumber = "44-" & CStr(41806 + RandomDX(916)) & "-FO"
        
            End If
        
        End If
        
    ElseIf prsBomber![BomberModel] = AVRO_LANCASTER Then
        
        ' KB726
        ' A(65) to Z(90)
        ' No idea what Lancaster serial numbers were. Based on Mynarski's
        ' Lanc having number KB726, we'll assume two letters followed by
        ' a number from 1-999.
        
        strManufacturer = "Avro"
        strPlant = "Birmingham" ' No idea, just made this up
        
        strTailNumber = Chr(RandomDX(26) + 64) _
                      & Chr(RandomDX(26) + 64) _
                      & CStr(RandomDX(999))
        
    End If

    If strTailNumber = "" Then
    
        ' Something went wrong: Make up a number.
        ' 40-1001 to 44-100000
        strTailNumber = "4" & CStr(RandomDX(4)) & CStr(RandomDX(99000) + 1000) & "-XX"
    
    End If

    ' Populate the structure fields for record insertion.

    prsBomber![TailNumber] = strTailNumber
    prsBomber![Manufacturer] = strManufacturer
    prsBomber![Plant] = strPlant

    ' Populate the BOMBER_TAB fields.

    With frmMainMenu
        .txtTailNumber = strTailNumber
        .txtManufacturer = strManufacturer
        .txtPlant = strPlant
    End With

End Sub
