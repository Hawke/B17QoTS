******************************************************************************
read_me_1st.txt

@author Preston V. McMurry III, http://www.prestonm.com
@copyright (C) Copyright 2002, 2010 by Preston V. McMurry III, http://www.prestonm.com

*****************************************************************************

This file is part of B17QotS, the "B-17: Queen of the Skies" Emulator.

B17QotS is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

B17QotS is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with B17QotS. If not, see <http://www.gnu.org/licenses/>.

---( Introduction )-------------------------------------------------------------

Hi, I'm Preston.

B17QotS is based on Avalon-Hill's "B-17: Queen of the Skies", a solo board game of WWII aerial combat over Europe. It includes all known variants published prior to mid-2005, including the detailed B-24 variant I wrote.

I started working on B17QotS in late 2002, to keep my programming skills semi-sharp after I got laid of during the first tech bust. VB6 was already on its way out, but I could get the compiler cheap as part of a student package at the local community college, so VB6 and Access were what I used. By 2005, when I stopped working on the project, I was already thinking of ways to do version two in some more portable, cross-platform technology, such as PHP and MySQL. This distro includes that design.

Feel free to use my design, or use your own, in the language and database of your choice, as long as it conforms to the GPL (GNU General Public License). I don't have time to do any more development, but if you need a play tester, feel free to contact me at http://www.prestonm.com. (Click on one of the contact links to send me an email.)

---( Credits )------------------------------------------------------------------

B17QotS Emulator:       Preston V. McMurry III.

Original Prototype:     Mike Pranno.

Additional Development: Shaun Foster.

Splash Screen:          Steve Dixon.

Testing:                Steve Dixon, Shaun Edley, Jeff Flint, Eddie Githens,
                        Jay Haygood, Rich Kaumeier, George Martin, Mike Peccolo,
                        Mike Pranno, Kevin Schafer.

Variants:               Bruce Peckham, Mike Haley, John Ockelmann,
                        Jeff Larsen(?), Preston V. McMurry III.

Tab Pseudo-Control:     Stefaan Casier.

---( Contents )-----------------------------------------------------------------

    1) Complete 1.0b2 VB6 and Access source code, database, etc. from February, 2005.
    2) Playable game (1.0b3 executable from June, 2005).
    3) Designs for both v1 and v2.
    5) A copy of the online help docs located at http://www.prestonm.com/military/B17QotS/doc/Index.html.
    4) Many variants. With the exception of the B-24 variant, these were developed by other folks.
    5) This read me file.
    6) A copy of the GPL license.

---( Release History )----------------------------------------------------------

25-Jun-2005

    * 1.0b3 released to test group.

    * FIXED: When the plane went down due to loss of fuel, an additional crew member (navigator) from another aircraft appeared in the crew list, bailing out with the twelve men formally assigned. He also appeared in the final accounting at the end of the mission. (tigone, 10 May 2005)

    * CANNOT RECREATE: My YB-40 appears not to have made the bomb run over the target -- no bombs on board, but exposure to flak there is an important part of the game -- nor did it fly the return leg of the target zone. The plane went from "outbound" in Zone 10 directly to "return" in Zone 9. My presumption is that the YB-40s stayed with the regular groups through the whole mission, and so this seems to skip a couple of important segments of the flight. (tigone, 10 May 2005)

    * CANNOT RECREATE: I was busy fighting off a fighter wave. As I went to assign mg fire, as I clicked on the nose guns I got the following error message, run time error "9". (MikeMarie) Am pretty sure it crashed when I selected the nose gun to assign fire. The RTE was very quick and then the emulator crashed out and closed. (SonBae, v1.0a2, 23-Jul-2004)

--------------------------------------------------------------------------------

21-May-2005

    * FEATURE: Allow menu and mission windows to be minimized. (Brazos, 13 May 2005)

    * FIXED: Review documentation to make sure it matches current functionality. Specifically, mention the Log Speed setting. (Brazos, 13 May 2005)

--------------------------------------------------------------------------------

21-Mar-2005

    * FIXED: On the "Generate Mission" tab: Clicking on the Bomber drop-down, mousing through the list of available bombers, and failing to select one before closing the drop-down causes a "Run-time error '9': Subscript out of range" error. Clicking OK terminates the emulator. Clicking on the Target drop-down, mousing through the list of available targets, and failing to select one before closing the drop-down causes a "LookupBomberTarget() Target 0 not found" error. Clicking OK in this dialog box opens a new dialog box reading "Run-time error '365': Unable to unload within this context". Clicking OK terminates the emulator. The Month, Year, Squadron, and Formation drop-downs do not exhibit this behavior. (rkaumeier)

--------------------------------------------------------------------------------

18-Feb-2005

    * 1.0b2 released to test group.

    * It's amazing to me how much progress has actually been made in the past 10 months. (Since the 8-Apr-2004 update.) The emulator has gone from being a pile of non-functioning, un-compileable, code to an actual game. Seems like it has been much longer than that ... :-)

    * Added splash graphic provided by Steve Dixon.

    * FEATURE: Any chance of having the results of the mission either be a separate pop-up, or have that final list stay up after you select "Exit from Mission". Right now it goes past so quickly, I'm not sure who succumbed to wounds, etc. If that screen would pop up before you can Exit Mission, or that was an additional step before exiting, I think that would be a nice feature. (kfgreer1961, 8 Dec 2004)

    * BUG: Should not be able to assign non-default airman to a default bomber. (The error is possible due to default bombers being listed in the bomber combo on the airman tab.)

    * BUG: Default groups and squadrons should not be updated.

    * BUG: Bombs jettisoned after release damaged.

    * Started discussion on what technology to use for v2.0. (v1.0 is being done in Visual Basic and Access.) I am thinking about doing v2.0 in a cross-platform web-based technology. Probably PHP.

--------------------------------------------------------------------------------

5-Dec-2004

    * 1.0b1 released to test group.

    * BUG: How does POW status work? I have had members of two crews become POWs. The first plane was lost over the Netherlands, and the second deep in Germany on a mission to Chemnitz. Two crwemembers survived the first bail-out, nine the second. On both missions all the crewmembers who survived to become POWs were listed as being on Admin Duty starting with the next mission. I have to admire those nine guys who made it back from the Chemnitz raid. That's a heck of a breakout and journey back to the 8th's home turf! :-) (4 Sep 2004, rkaumeier)

    * Entire crew died in emergency bail out after wing tore off. (Which is okay.) The error occured on exitting the mission: The airmen were unassigned, the bomber positions were still filled, and the airmen were not killed in the database. (6 Jul 2004) Same as above?

    * If airman is captured or killed on his 25th mission, he will actually be set to tour complete status, rather than KIA or POW status. (28 Nov 2004) Same as the two above?

    * When airman was invalided, he was placed on Admin duty, rather than having his last assignment left as his last assigned bomber. (28 Nov 2004)

    * When bomber was retired, airman was not placed on admin duty. He was still assigned to bomber (pvmiii) ... This isn't a bug, per se, because the code was completely forgotten! Also, bomber should be removed from assignment dropdown on airman tab. (28 Nov 2004)

    * If a bomber is shot down, it is showing up in the assignment dropdown on the airman tab. (28 Nov 2004)

--------------------------------------------------------------------------------

16-Nov-2004

    * 1.0a4 released to test group.

--------------------------------------------------------------------------------

7-Nov-2004

    * RE-FIXED: When typing in the name to add a new bomber, when the first letter is typed, the LookupBomberSquadron() function throws a "Squadron 3 not found" error, where it was trying to lookup the squadron number for whatever bomber was current before the user started typing in the new one. (pvmiii)

    * QUESTION: What effect does changing an airman's position have when that airman is assigned to a bomber? If the airman's position is changed to another position on his current bomber, or any other bomber, then the airman is assigned to the new position on that bomber. If a second airman already occupies that position, the second airman is assigned to admin duty. (In reality, a given bomber often had more than one crew, or individual crew were loaned out if a bomber was short for a mission.)

--------------------------------------------------------------------------------

6-Nov-2004

    * FIXED: If I close the program by clicking on the X button in the upper right hand corner of the window, the application appears to terminate but continues to be shown as a running process in the task switcher. Starting another instance of B17 and closing it the same way adds another process. This can continue indefinitely until out of memory. This only happens from the mission screen while a mission is in progress. Using the Windows close button at any other time terminates the program and kills the process. Problem exists in both Win98 and XP. (rkaumeier)

    * FIXED: A minor nitpick concerning one crewmember designation used in the YB-40: Both the bomber maintenance assign crew box and the mission screen refer to a "Mid-Top" gunner. This same position is selected as the "Mid-Upper" gunner on the airman maintenance page. (Rick)

--------------------------------------------------------------------------------

31-Oct-2004

    * Feature: Bailout over base zone, or land (option), if severely damaged or flight pilot and copilot are dead. (AFIntel)

--------------------------------------------------------------------------------

5-Sep-2004

    * FIXED: When a mission is saved to HTML, the document should be HTML 4.01 transitional compliant. I also updated the emulator documentation to be compliant. Here are two samples, one for a B-17, the other for a Lancaster:

      Target: Chemnitz
      Target: Essen

      Note how HTML was dynamically generated based on the targets' specific information.

--------------------------------------------------------------------------------

29-Aug-2004

    * 1.0a3 released to test group.

    * Updating an existing, but unassigned, airman to a specific bomber caused the airman to be assigned to a random default bomber. (Dakota Queen?) Also, new airmen were assigned to admin duty even when a bomber was designated. Now, a new airman may be assigned to admin duty, an empty position on a bomber, or a filled position on a bomber. (In which case the previously assigned airmen is sent to admin duty.) Likewise updating an existing airman. If you try to add or update an airman to a default bomber, the airman will be placed on admin duty. (pvmiii)

    * Bug in TempPosSwap() causes the swapping airman to vanish from the display, which in turn causes a subscript out of range error the next time that position is referenced. At the risk of jinxing myself, after 3+ days, I think I have it fixed.

    * If navigator is seriously wounded, the port cheek should not be fireable. (Stbd cheek is properly set to non-fireable.)

    * On the Bomber Maintenance Screen if you click "Add" and you have not listed a squadron for that bomber you get an "Error 91" message. (SonBae) -- Need to put validation on all add and update tabs.

    * Slow down mission log scroll. The default is 2, which provides a quick, but smooth and readable scroll. Values may be 0 (very quick) to 10 (one second, very slow).

    * Mission Log now automatically scrolls to the bottom of the box -- i.e., the most recent message -- thus saving more clickage.

--------------------------------------------------------------------------------

1-Aug-2004

    * FIXED: When the images and help files are package, the directory is being lost, resulting in the files being places in the program directory rather than /doc or /images. In the SETUP.LST created by Microsoft's package & deployment wizard, make sure that the following settings are correct:

        DefaultDir=$(ProgramFiles)/B17QotSEmulator
        .html files have $(AppPath)/doc
        .jpg files have $(AppPath)/image

      If the settings are not correct, the program may not launch, nor the documentation be found.

    * FIXED: Bomber model should not be modifiable. It causes all sorts of problems. (The docs already state this, but the check wasn't being done.) The only modifiable bomber field is squadron.

    * FIXED: QuitBeforeWater(), which allows the user to bail out before crossing water, was allowing negative last zone, which threw an array error.

    * FIXED: If the bomber is not on duty status at the conclusion of a mission -- either due to being shot down, or due to the loss of crew -- it should not be listed on the bomber combo on the mission tab.

    * FIXED: Create group '9 group'. Try to add squadron '99 squad' to '9 group'. AddSquadron() bails when it tries to insert the '99 squad' into the squadron combo on the bomber tab. The combo index is off because the combo is filtered by type of bomber in squadron. Deleting squadrons also causes a array index error to be thrown on the squadron combo on the bomber page. (pvmiii)

--------------------------------------------------------------------------------

25-Jul-2004

    * 1.0a2 released to test group.

    * FIXED: After running a mission, I clicked on the save button. As the save screen came up, I then clicked on cancel. This caused the program to abort with the following error message: "run time error 32755". (Mike Peccolo)

    * FIXED: If you click on a medium cyan fighter before selecting a medium cyan gun, the fighter is greyed out.

    * Implemented realistic tail numbers. The Lancaster numbers resemble real tail numbers; the American tail numbers are real tail numbers, randomly determined.

    * Users have the option to bail out/crashland in land zones prior to a water zone if the bomber might otherwise have to ditch in the water zone.

    * Red Tail Angels: This option allows for the chance of 15th Air Force bombers being escorted by the 332nd Fighter Group, the famous Tuskegee Airmen, which did not lose a single bomber to German fighters. If this option and unescorted are both chosen, then the mission will be unescorted.

--------------------------------------------------------------------------------

12-Jul-2004

    * Cleaned up code, including fixing form sizes.

    * Added bomb sight and environmental gauges.

    * Hid bomb-oriented information for YB-40s. (Because they didn't carry bombs.)

    * Oops Factor: Missions to Mulhouse and Friedrichshaven may accidentally bomb Switzerland. (Especially in bad weather.)

    * Crew Experience: Inexperienced crews are more likely to miss the target and have accidents, while veteran crews are less likely to experience such mishaps. (Added the Airman.LeadCrewExp field to the database.)

          Gunners:    If gunner has <= 5 missions, DM -1 to hit, plus an unmodified 1 to hit jams the weapon.
                      (After which normal unjamming rules apply.)

          Pilot:      If random event generates aerial collision and pilot and copilot have missions <= 5,
                      DM +1 to severity roll. If pilot or copilot has missions >= 11, DM -1 to severity roll.

          Navigator &
          Bombardier: If both navigator and bombardier have missions >= 11, and both have one "lead crew"
                      mission (flying in lead-low squadron) to a target, then DM +1 to 06BombRun() on each
                      subsequent lead-low mission to the same target. Lead crews may start earning experience
                      in April, 1943.

    * FIXED: Fixed bug that was writing the bonus Me-109 -- for high, low or out of formation, bombers -- over the last fighter in a wave. (Thus removing potential 6th fighter.)

    * FIXED: Mid-Upper gun is appearing on B-17s other than the YB-40 model.

    * FIXED: Even though the ammo box is greyed out, and the ammo is 0, the gun may still be fired at an enemy fighter if the user clicks on the ammo box, then the fighter, then "Fire Guns".

    * FIXED: Selecting a German fighter, without first selecting a weapon, throws a run-time 340 error (array element 0 doesn't exist).

    * FIXED: Inoperative ball turret can still fire:

        Waist: Ball Turret - Ball turret inoperable. Gunner trapped.
        Wave 2 (Attack 1): 3 fighters
        Fighter cover: Poor
        0 enemy fighters chased off by cover
        Nose hit FW190 at 12 High - FCA
        Stbd Cheek missed.
        Stbd Waist missed.
        Top Turret missed.
        Ball Turret hit Me109 at 3 Level - FBOA

    * FIXED: One engine operating: May fly one additional zone, then either ditch, crash or bail out. May fly one zone further than that by throwing all ammo and handheld extinguishers overboard, plus jettisoning bombs. (If bomb bay doors are jammed, this can't be done.)

--------------------------------------------------------------------------------

5-Jul-2004

    * 1.0a released to test group.

    * FIXED: Many issues related to flying over the Alps in the 15th Air Force variant. (This was a bitch to fix.)

    * FIXED: Deleting an airman that is assigned to a bomber does not set the position to unmanned, resulting in a lookup error which shuts down the emulator.

    * FIXED: Unescorted mission appears to occassionally generate a single friendly fighter.

    * FIXED: Switching from one target to another causes problems.

    * FIXED: Roll twice on table O3 (not three times) for light flak when at low altitude.

    * FIXED: Spend two turns in every zone if one engine out and bombs aboard, or two or more engines out. Spend two turns in odd-numbered zones if out of formation and navigation system inoperable.

        I just realized that you only roll for enemy once per zone when out
        formation, not twice in every zone. It's only twice per zone if engines are
        out, or twice every other zone if the navigation system is out. I'd probably
        have a couple of planes still flying if I hadn't being playing incorrectly
        for so long ... :-(

    * Contrails: In zones 2-12, if weather is clear, roll 1d6. On 5 or 6, contrails form. If contrails form, apply +1 to tables B-1, B-2 and O-2. (The variant states O-3, but that makes no sense as O-3 is a 2d6 table, whereas O-2 is a 1d6 table with graduated results.) Contrails apply to both outbound and return trip in the zone in which they are rolled. Contrails do not apply to night missions (Lancasters). From the "Theater Modifications" article in "The General" (Volume 24, #6).

--------------------------------------------------------------------------------

24-Jun-2004

    * FIXED: Bomber spent two turns in every zone, regardless of damage, if out of formation.

    * FIXED: Bug caused bombers flying back-to-back missions to switch base from England to Italy.

    * FIXED: Many issues related to B-17Cs, and some related to B-17Es.

    * Continued working on internal (within code) and external (user visible) documentation.

    * Added option to fly mission completely unescorted. (For the brave or foolhardy ...)

    * Added 30 new targets: 7 in France, 2 in Belgium, 2 in Holland, 1 in Denmark, 2 in Czechoslovakia and 16 in Germany. Two of the targets are in Zone 12! Click and save this image to view it full size. (Only the viewing area has been reduced, not the image itself.) There are now 121 targets in the emulator.

--------------------------------------------------------------------------------

21-Jun-2004

    * Working on internal (within code) and external (user visible) documentation.

--------------------------------------------------------------------------------

16-Jun-2004

    * Continuing to test and debug. I have run hundreds of missions. I'd like to have an alpha release in a couple of weeks.

--------------------------------------------------------------------------------

20-May-2004

    * Added the following new functions and/or functionality:

        UnjamGuns()
        ConsumeOil()
        DropToLowAltitude() <-- ask if the user wants to descend
        LoseAltitude() <-- actually does the descent
        MechanicalFailure()
        RemoveTempOptions()
        DisplayTempOptions()
        cboMonth_Click()
        chkTimePeriodSpecificFormations_Click()
        AdjustMissionOptions()
        DisplayBombBayAmmo()
        txtBombBayAmmo_LostFocus()

    * Did a lot of testing of variants, such as random events, mechanical failures, time period specific formations, formation defensive gunnery, evade flak, alternate weather, German fighter pilot skill, JG26 stationed in Abbeville, and Ju88s used as fighters.

--------------------------------------------------------------------------------

13-May-2004

    * Actively testing and debugging the mission engine. (In other words, playtesting the *game*, rather than the supporting screens.) Also did work on the following functions:

        TempPosSwap() <-- e.g. bombardier fires cheek gun ...
        UndoTempPosSwaps() ... then returns to position
        SprayFireAllowed()

    * Added gunnery tables for the B-17C, B-17E, B-17G, YB-40, B-24D, B-24E, B-24G/H/J, B-24 L/M and Lancaster.

    * Regarding the use of color: Generally speaking, green (or grey) indicates the object is okay. Yellow is a warning, light wound or light damage. Red is system failure, serious damage or serious wound. Black indicates death or very severe damage. Medium cyan indicates guns with potential targets, or the potential targets themselves. Pale cyan indicates a gun that has been assigned to a target, plus the target it was assigned to.

    * If you see a gap in the gun or crew list, that is because the hidden position does not actually exist on the bomber model you are flying.

--------------------------------------------------------------------------------

4-May-2004

    * Continued work on the mission engine. The following functions have either been added, modified or are in progress:

        InitializeMissionInfo()
        RefreshMissionInfo()
        InitializeCrew()
        RefreshCrew()
        InitializeGuns()
        RefreshGuns()
        ExitPrevZone()
        cmdTakeOff_Click()
        cmdInterrupt_Click()
        GetWaveSize()


    * Actually had a successful compile for the first time in over a year. (Since I was heads down just hammering out the basic charts, without any testing.)

--------------------------------------------------------------------------------

30-Apr-2004

    * Added "gauges" to the mission form. The "gauges" are actually a few labels, colored to indicate the status of major systems such as engines, ailerons, fuel, etc. Nothing fancy, but it adds a bit of flavor.

    * Continued work on the mission engine. The following functions have either been added, modified or are in progress:

        B7RandomEvents()
        Swap Ammo Form
        SwapAmmo()
        InitializeGauges()
        RefreshGauges()
        ColorEngineGauge()
        ColorOilPressureGauge()
        ColorFuelGauge()
        OverWater()
        CrashLanding()
        optBailout_Click()
        optCrashDitch_Click()
        NormalZoneCombat()
        TargetZoneCombat()
        AbortMission()
        optAbortMission_Click()
        optContinueMission_Click()

    * Since luck is a bomber parameter, a rabbits foot will always be expended where a die roll results in a bomber being shot down. (Which would cause all accumulated luck to be lost, so it might as well be used.) Luck will also be used to save ace gunners, who are valuable to keeping a bomber in the air, and to save ball gunner from being crushed in an inoperative turret.

--------------------------------------------------------------------------------

23-Apr-2004

    * Added Flak Evasion and Formation Defensive Gunnery variants from "The General" (Volume 24, #6).

    * Continued work on the mission engine. The following functions have either been added, modified or are in progress:

        DefensiveFire()
        M1DefensiveFire() <-- different than above function
        M2HitDamageAgainstGermanFighter()
        M5SprayAreaFire()
        EvasiveActionAllowed()
        OffensiveFire()
        B5AreaDamage()
        B5AreaDamageB17()
        B5AreaDamageB24()
        B5AreaDamageLanc()
        B5AreaDamageRouter()
        B6SuccessiveAttacks()
        PassingFireAllowed()
        PassingFireSetup()
        Miscellaneous coloring functions that provide standard field backgrounds
        lblGun_Click()

--------------------------------------------------------------------------------

17-Apr-2004

    * Have spent the past week working on the mission engine. The following functions have either been added, modified or are in progress:

        FlyMission()
        TakeOff()
        EnterZone()
        FlakCombat()
        AirToAirCombat()
        lblToHit_Click()
        GunIsManned()
        MayFireGun()
        PreAssignGuns()
        HasATarget()
        GetMaxRemovals()
        FightersInWave()
        RemoveFightersFromWave()
        DisplayWave()
        Interrupt()

    * Amongst other things, I am trying to reduce the amount of clickage that occurs while a mission is in progress. After this is done, you will only have to click after you have provided some sort of necessary input. (Such as allocating defensive fire.)

--------------------------------------------------------------------------------

8-Apr-2004

    * I just realized that this was not displaying on the internet, so whatever was wrong has been corrected.

    * I have finished putting the colored cardstock charts into code. No idea if it works, but shoehorning all the various variants -- including the B-24 and Lancaster -- into the code took some doing. The charts are not complete, but the vast majority of the initial effort is done.

    * The next step is to analyze and psuedocode the entire mission process from takeoff through final landing. (Assuming the bomber returns.) That will determine where to place the chart calls within the overall scheme of things. For those who've tested the original alpha of this game, you will realize that much of that is probably already done. But given that the charts had to be substantially rewritten (basically from scratch) to handle the variants, the same effort must be made regarding the mission process. In addition, the mission process may be cleaned up / modularized for ease of future expansion.

    * Despite the fact that this help file has not been updated in 11 months, I have actually been occassionally working on the program ...

--------------------------------------------------------------------------------

9-May-2003

    * Generate Mission completed.

    * Mission sheets completed. If you want to run a paper game, rather than a computer game, the emulator will generate the mission for you.

--------------------------------------------------------------------------------

