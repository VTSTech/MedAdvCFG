# MedAdvCFG - Mednafen Advanced Configuration Tool

Frontend for <a href="http://mednafen.fobby.net/releases/">Mednafen</a> v0.9.x.x

Tested with the following System Cores:

GB,GBA,GG,MD,NES,PCE/PCE_FAST,PSX,SNES,SS,VB

Homepage: http://www.nigeltodman.com/medadvcfg

<img src="https://i.gyazo.com/23d66bd5b868a3492ffc08adca9300d6.png">

<img src="https://i.gyazo.com/92a29a9a1149ff763fc6fb98a8292e3e.png">

<img src="https://i.gyazo.com/c98c6105efcf218ee89e9760b22fd18b.png">

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
