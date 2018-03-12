# MedAdvCFG - Mednafen Advanced Configuration Tool

Frontend for <a href="http://mednafen.github.io/releases/">Mednafen</a> v1.x

Tested with the following System Cores:

GB,GBA,GG,MD,NES,PCE/PCE_FAST,PSX,SNES,SS,VB

Homepage: http://www.nigeltodman.com/medadvcfg


<img src="https://i.gyazo.com/3cfc961fddd403f5fe60389f28b3d00d.png">

<img src="https://i.gyazo.com/a748b38a9e2ffb846691f0413da21ea6.png">

<img src="https://i.gyazo.com/0fdd3ab59d2e938e4f22f6062d44b724.png">

<img src="https://i.gyazo.com/3f67974b9b851cc8496a9fc6bd99dae8.png">


v0.4.1-r44 03-12-2018 12:09PM

Desktop Composition is now saved/loaded with settings

Autosave/load State is now saved/loaded with settings

Disc Tray/Cart ejected is now saved/loaded with settings

Saving in Advanced Mode no longer forgets Basic Mode Folder setting

HQ2X is now the hardcoded Video Scaler in Basic Mode.

Frameskip is now set to enabled for Basic Mode.

BIOS & ROM Sanity is now set to disabled for Basic Mode.

Various code optimizations

v0.4.0-r43 03-11-2018 11:13AM

Added option to autosave/load game state

Added option to disable Desktop composition

Renamed some controls to more accurately reflect the parameter they control

'Number of Players' control renamed to netplay.localplayers. Also passes that parameter in addition to setting the 'Controller' to specified type and plugged in for specified number of players

Horizontal Overscan renamed to h_overscan and now also works for Sega Saturn SysCore

Added Twitch icon and link

Updated roadmap

Various code optimizations

v0.4.0-r42 03-09-2018 8:49AM

Region is now autodetected based on BIOS, Defaults to NTSC-U if using unknown BIOS.

Regions added to bios.dat

Added option to boot SysCore with Disc Tray/Cart Ejected

Clicking 'Mode -> Advanced' from the 'Game Info' screen of Basic Mode now actually returns you to Advanced Mode.

Various Code Optimizations

v0.4.0-r41 03-06-2018 5:20PM

Fixed reference to removed control in Basic Mode

Resized game selection form in Basic Mode slightly

Game ID is now displayed in Basic Mode for PSX System Core

v0.4.0-r40 03-06-2018 4:34PM

Added support for Sega Saturn Racing Controller (MK-80107)

Added support for Sega Saturn Mission Stick (MK-80104)

Added support for Sega Saturn Dual Mission Stick (MK-80104)

Added support for Sega Saturn Virtua Gun/Stunner (MK-80113)

Added support for Sega Saturn NetLink US Keyboard (MK-80120)

Added support for Sega Saturn 89-Key JP Keyboard (HSS-0129)


Added '0.9.x.x Mode' - Prevents 1.x parameters from being passed to cmdline. Automatically activates when a known 0.9.x.x version is detected.

Various code optimizations

v0.4.0-r39 02-26-2018 12:14AM

Condensed bios.dat names slightly

Added missing entry to psx-usa-id.dat

Resized main form slightly

Reorganized main form controls

Removed Region checkboxes

Added support for 1.21.0-unstable-win64/win32

Added fps.autoenable (Show FPS)

Added video.fs.display

Removed 'Overlay' video driver option

Fixed bug in passing Sega Saturn region to cmdline

Scanlines and Custom Params are now loaded with settings

Various code optimizations

v0.3.3-r38 02-24-2018 5:14PM

(Mednafen.dat update. Code changes for 1.xx-unstable coming in next build)

Added support for Mednafen 0.9.48.0-win64/win32

Detects Mednafen 1.21.0-unstable-win64/win32

v0.3.3-r38 09-10-2017 12:17PM

(Mednafen.dat update. No code changes required)

Added support for Mednafen v0.9.46.0-win64/win32

Added support for Mednafen v0.9.47.0-win64/win32

v0.3.3-r38 06-10-2017 07:22PM

(Mednafen.dat update. No code changes required)

Added support for Mednafen v0.9.45.1-win64/win32

v0.3.3-r38 06-04-2017 02:30AM

(Mednafen.dat update. No code changes required)

Added support for Mednafen v0.9.45.0-win64/win32

v0.3.3-r38 04-27-2017 5:11PM

(Mednafen.dat update. No code changes required)
Added support for Mednafen v0.9.44.1-win64/win32

v0.3.3-r38 02-26-2017 11:32PM

Added support for Mednafen v0.9.43.0-win64/win32

Goat Scanlines/Progressive controls are now functional

Added Horizontal Overscan (PSX)

Basic Mode now saves settings when games are listed.

Various code optimizations

v0.3.3-r37 02-12-2017 12:26PM

Changed main font to Arial

Added TurboGrafx-16 Controller support

Added PC-FX Bios support (untested)

Various code optimizations

v0.3.3-r36 02-11-2017 11:44AM

Goat Mask control is now functional

Verifying NES with GoodTools

FileNameCleanup for Cover Search cleans gba/gbc/gb/ccd roms

Better handling of first instance of cover/rom search

Updated roadmap.txt

Various Code Optimizations


v0.3.3-r35 02-10-2017 1:09PM

Added snes_faust SysCore support

Added controllers for Sega Genesis/MegaDrive

Added GOAT Shader + Options

Various Code Optimizations

v0.3.2-r34 02-09-2017 4:51AM

Added support for Mednafen v0.9.42.0-win64/win32

Should be verifying MD in both Basic and Advanced Mode now

Removed hardcoded MD5 for BIOS and Mednafen.

Now using bios.dat and Mednafen.dat under /dat/

Fixed 'Path Not Found' error when no bios set and settings are saved/loaded

Removed CoversDB data from GitHub. Added URLS+Mirrors in CoversDB.txt in respective cover folders

Various Code Optimizations


v0.3.2-r33 02-06-2017 8:40PM

More reliable switching of SysCores (Reloads Form3)

Better handling of first instance of SysCores (Creates empty game/cover file)

Verifying GG,MD with NoIntro (www.No-Intro.org)

Resizing NES,GG,MD Covers on Large Cover view

ROM MD5 Copies to Clipboard on click in Large Cover view

Various Code Optimizations


v0.3.2-r32 02-06-2017 7:23AM

New versioning. Allows multiple revisions per minor version.

Changed Advanced Mode Font.

Hiding PSX Game ID on unused cores

Hiding unused BIOS Controls in Large Cover view

Added CoreControls() function to hide/show controls in Adv Mode

Basic Mode Quick Launch setting is now saved

Now verifies SNES Roms with NoIntro

SS and PCE will be verified with REDUMP, Rest with NoIntro

Removed md5.exe credits. No longer used.

Settings now loaded on Game Cover screen. Fixes incorrect Quick Launch cmdline

Various Code Optimizations

v0.3.1 02-04-2017 7:48AM

Added two new menus!

Help & Chat

Under Help is link to Documentation

And Tips for NetPlay


Chat will bring you to new #MedAdvCFG Chat

On Freenode!


Removed md5.exe requirement. Now using Lib cryptdll.dll

Removed much unneeded commented code

About Menu is now functional on all forms.

Now verifies PSX BIOS with REDUMP

REDUMP Verification now displayed in Basic Mode


Added roadmap.txt to GitHub https://github.com/Veritas83/MedAdvCFG

EXE now compressed with UPX v3.91w
1433600 ->    671232   46.82%    win32/pe     MedAdvCFG.exe

Various Code Optimizations

v0.3.0 02-02-2017 4:20AM

Redump verification for PSX

NetPlay controls are now functional*

NetPlay control is now a ComboBox

Several NetPlay servers predefined.

Resolution now selectable from dropdown

Ads, Because who doesn't like ads.

Added link to CoversDB.org

PlayStats are planned for future

Hosted by CoversDB.org

Various Code Optimizations

* Notes on NetPlay Support

Not all SysCores support NetPlay.

Namely, SS does not.

SNES and PSX do :)

To NetPlay, Use Same BIOS and Same ROM

MD5 must match across board. Select same server. Play!

v0.2.9 01-30-2017 9:06PM

Resizing SNES cover in 'Large Cover' view

Now detects PSX Game ID for NTSC-U Region

Game IDs stored in /dat/psx-usa-id.dat

Displays PSX Game ID in Advanced Mode and Basic Mode

Resized Advanced Mode screen and reorganized controls slightly

Better Handling of 'Input Past End of File' Error w\ Large Cover launch

Updated /covers/ on GitHub

Various Code Optimizations


v0.2.8 01-29-2017 2:26AM

Resized and updated some of the Console logos.

Hid unused labels on 1st step of basic mode

Resized console selection window

Increased vertical spacing between covers

Now using www.CoversDB.org Covers!

Now using standardized filenames for cover search
All [space] are _
All - are _
All ,'".*:? symbols are removed
Entire name is lowercased

Various code optimizations


v0.2.7 01-27-2017 3:54PM

Added back buttons to navigate thru Basic Mode

Page Flipping improved further. Can now flip back to Page 1, Flipping beyond last page will rollover to Page 1

Improved Basic Mode ROM Folder loading

Improved Cover Searching, fixes trailing '_The'

Resizing SNES Covers (again)

Resizing NES, Saturn & Genesis Covers

Fixed bug in rom detection if extension was used in folder name.


v0.2.6 01-27-2017 12:01AM

Basic Mode Update.

Page flipping is much faster.

Getting covers for only next page, Not entire list now :)

Basic Mode will now rom/cover search for following SysCores

GB, GBC, LYNX, SS, MD, VB, GG, PCE, NES, SNES, PSX

.dat files storing lists of roms/covers now located in /dat/

Various code optimizations


v0.2.5 01-26-2017 10:32PM                
                                         
Changed .pixshader to .shader            

Added 'nocover.jpg' placeholder to GitHub


v0.2.4 01-26-2017 3:38PM

Basic Mode Update.

Added much needed 'No Cover' place holder.

Adds ability to launch all games from Basic Mode. (Even without covers)

Added 'Quick Launch' (is default)

Not 'Quick Launching' will display an extra large cover with:
Rom Filename, Cover Searched, BIOS, MD5's and a Play/Back button.

Resized SNES Box Art and PSX Box Art slightly

Console Logo is now centered upon selection

Mouse Over Cover will show its 'Cover Searched' name

Basic Mode ROM Folder can now be saved.

BIOS is now auto-verified if it is set.

v0.2.3 01-25-2017 4:50PM


Basic Mode Update. Semi-Functional

Now parsing covers from /covers/psx/, /covers/snes/, etc

Will search for roms/cues for NES, SNES and PSX for now.

Displays detected covers. Click on cover to launch.

Many code optimizations


v0.2.2 01-22-2017 12:13PM


Updated for Mednafen v0.9.41-win32/64

Added 'Basic Mode' - Not finished yet

Normal view up until now is 'Advanced Mode' - Default

Various code optimizations


v0.2.1 09-20-2016 4:34PM


Updated for Mednafen v0.9.39.2-win32/64

Reorganized main form controls

Resized main form slightly

Added non-functional fields for Netplay support

Various code optimizations


v0.2.0 08-26-2016 2:49AM


More accurate population of Controller list

Now verifies BIOS on Load if BIOS is set.

Now actually passes sanity checks to cmdline w\ Saturn Core.

Now saves many more settings. Specifically saves:


Force Mono, Disable Sound, Scanlines, video.blit_timesync

video.glvsync, untrusted_fip_check, cd.image_memcache

Axis Scale, Number of Players & Custom Params


Various code optimizations


v0.1.9 08-25-2016 11:56PM


Updated for Mednafen v0.9.39.1-win32/64

Added cd.image_memcache option

Reorganized main form controls

Changed Window Title/About text to reflect 0.9.x.x support.

Various code optimizations


v0.1.8 08-19-2016 2:54PM


Updated for Mednafen v0.9.39-unstable

Added support for Sega Saturn!

Added Controllers for Saturn

Validates various BIOS for Saturn

Auto Selects US/JP region based on BIOS

Resized main form slightly

Various code optimizations


v0.1.7 08-05-2016 10:01AM


Added menu option to Reset All Settings.

Added Video Driver option.

Added Disable Sound option

Added Force Mono Sound option

Added video.blit_timesync option

Added video.glvsync option

Resized main form slightly

Improved path handling involved spaces

Various code optimizations


v0.1.6 07-06-2016 11:19 AM


Now supports SNES/NES Controllers

Auto-Clears/Auto-Populates based on SysCore

Now ignores BIOS/ROM Sanity checks on all SysCore except PSX

Various code optimizations


v0.1.5 06-20-2016 12:48 AM


Reorganized main controls

Added Scanlines

Added Video Deinterlacer

Added Axis Scale

Added Option to Suppress All Confirmations!

Now hides Detailed Bios Verification information when PSX not emulated

Various code optimizations


v0.1.4 06-19-2016 1:27 AM


Reorganized main controls

Added Social Media Links/Icons

Now supports custom parameters


v0.1.3 06-18-2016 9:34 PM


M3U Generation now supported!

Select File - Generate Multi-Disc M3U. Enter # of Discs, Select them. Done

Outputs multi.m3u into current folder for MedAdvCFG.exe

Support Controller Selection (PSX Only for now)

Added Players Count, Will add this many input(x) to command line for controller


v0.1.2 06-17-2016 5:44 PM


Supports toggle of untrusted_fip_check parameter.

Used for Multi-Disc games specifying CUE/BIN in diff folders

Supports Multi-Disc games via M3U file.

M3U Generation coming in next build.

Now saves RomPath.

If RomPath or BiosPath exists, Resetting defaults to that path


v0.1.1 06-12-2016 12:48 AM


Now supports custom Save/Memory Card/State path

Remembers Bios/Firmware Path

Now validates every PSX Bios that has been dumped to date!

Selection of PSX Region now supported.

Auto selects PSX Region based on BIOS Region if known.

Now actually passes PCE/PSX Bios to cmdline

Resized main form slightly


v0.1.0 06-07-2016 6:08 PM


Suppresses 'Reset ROM/BIOS' Confirmation if a ROM/BIOS is not set.

Now passes video.fs 0 if Fullscreen is unchecked.

Now passes videoip 0 if Bilinear interpolation is unchecked.

Turns off MD5 Checks on System Core change.

Release compression changed to zip format

Changelog now has timestamps


v0.0.9 06-07-2016


ROM/BIOS MD5 Check is now optional

(Note: Works just fine on small bios/rom/cue. But takes time to md5 an entire iso)

Realigned ROM/BIOS 'Set' buttons

Added Mednafen icon.

Global Scaling Factor now saved with 'File -> Save Settings'

Fullscreen Resolution Override now saved with 'File -> Save Settings'

LastPath is now saved with 'File -> Save Settings'

More links in 'About' dialog


v0.0.8 06-06-2016


Reorganized main form controls

Clicking Mednafen EXE/ROM/BIOS MD5 Copies it to clipboard

Added Mednafen logo. Click it to go to Mednafen Homepage.

Added Global Scaling Factor. Applies to all axis, fullscreen and windowed.

Added Fullscreen Resolution Override.

Video Scaler settings are now actually passed to cmdline



v0.0.7 06-06-2016


Tested with PCE/PCE_FAST System Cores

Now supports more Cartridge Based Consoles (GB/GBA/GG/MD/VB)

Adds notice that PCE and PSX Cores expect a BIOS Image w\ filenames

Automatically unsets and disables Sanity Checks on all Cores but PSX

Added button to Clear BIOS/ROM



v0.0.6 06-06-2016


Now supports Cartridge Based Consoles (Tested NES/SNES)

Omits -loadcd parameter on all Cores except PSX, PCE, PCE_FAST

Uses -force_module on all cores except PSX, PCE, PCE_FAST

Added detection for some 0.9.37 Builds


v0.0.5 06-05-2016


No more typing paths by hand!

Added Drive/Folder/File Selection panel

Auto collapses on File Selection

Now performs validation of MedEXE,ROM and BIOS on each change.


v0.0.4 06-05-2016


Added detection for v0.9.38.5 x86/x64

Reviewing Command Line is now optional


v0.0.3 06-05-2016


Added Frameskip

Enabled BIOS Sanity Check

Enabled ROM Sanity Check

Enabled Temporal Blur

Enabled Temporal Blur Accumulate color

Enabled Temporal Blur Amount

Enabled Bilinear interpolation


v0.0.2 06-05-2016


Resized main form slightly

Changed colors

Added HotKeys to main form

Published to GitHub https://github.com/Veritas83/MedAdvCFG


v0.0.1 06-05-2016


First release!

Only PSX System Core is tested.
