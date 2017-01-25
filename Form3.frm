VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "MedAdvCFG v0.2.3 (Mednafen v0.9.x.x Frontend) by Nigel Todman [BASIC MODE]"
   ClientHeight    =   8670
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12105
   LinkTopic       =   "Form3"
   ScaleHeight     =   8670
   ScaleMode       =   0  'User
   ScaleWidth      =   10.732
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00404040&
      Caption         =   "<"
      Height          =   2650
      Left            =   120
      MaskColor       =   &H00404040&
      TabIndex        =   1
      Top             =   2822
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00404040&
      Caption         =   ">"
      Height          =   2650
      Left            =   11640
      MaskColor       =   &H00404040&
      TabIndex        =   0
      Top             =   2822
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Page: 1/x"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5670
      TabIndex        =   3
      Top             =   8400
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Total Games: x"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   10200
      TabIndex        =   2
      Top             =   8400
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   12
      Left            =   8880
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   11
      Left            =   6120
      Picture         =   "Form3.frx":2B3FA
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   10
      Left            =   3360
      Picture         =   "Form3.frx":567F4
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   9
      Left            =   600
      Picture         =   "Form3.frx":81BEE
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   8
      Left            =   8880
      Picture         =   "Form3.frx":ACFE8
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   7
      Left            =   6120
      Picture         =   "Form3.frx":D83E2
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   6
      Left            =   3360
      Picture         =   "Form3.frx":1037DC
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   5
      Left            =   600
      Picture         =   "Form3.frx":12EBD6
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   4
      Left            =   8880
      Picture         =   "Form3.frx":159FD0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   3
      Left            =   6120
      Picture         =   "Form3.frx":1853CA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   2
      Left            =   3360
      Picture         =   "Form3.frx":1B07C4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   1
      Left            =   600
      Picture         =   "Form3.frx":1DBBBE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   0
      Left            =   11880
      Picture         =   "Form3.frx":206FB8
      Stretch         =   -1  'True
      Top             =   8400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mode 
      Caption         =   "Mode"
      Begin VB.Menu Advanced 
         Caption         =   "Advanced"
      End
      Begin VB.Menu Basic 
         Caption         =   "Basic"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MedEXE, FSO, tmp, tmp2, tmp3(99), BIOSFILE, BIOSPATH, ROMFILE, SystemCore, SYSCORE, BIOSSanity, ROMSanity, Stretch, PixelShader, VideoScaler, x, y, z
Dim cmdstring, Build, Frameskip, Fullscreen, TBlur, TblurAccum, AccumAmount, VideoIP, ActiveFile, XRes, YRes, ScaleFactor, LastPath, SavePath, BiosPathLoad
Dim ResetBios, ResetRom, ResetRomDir, ResetSave, FatalError, SystemRegion, SystemRegionLoad, ROMDIR, M3USize, LastFile, VideoDriver
Dim Bilinear, DisableSound, ForceMono, video_blit_timesync, video_glvsync, untrusted_fip_check, cd_image_memcache, scanlines, numplayers, customparams
Dim MedAdvGAMES, MedAdvCOVERS, MedAdvEXT, TotalGames, PageTotal, PageOn, TotalCovers
Dim GamesList(9999), CoversList(9999), tmparray(9999), CurrName(9999)
Public Function RecursiveDir(colFiles As Collection, _
                             strFolder As String, _
                             strFileSpec As String, _
                             bIncludeSubfolders As Boolean)

    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant

    'Add files in strFolder matching strFileSpec to colFiles
    strFolder = TrailingSlash(strFolder)
    strTemp = Dir(strFolder & strFileSpec)
    Do While strTemp <> vbNullString
        colFiles.Add strFolder & strTemp
        strTemp = Dir
    Loop

    If bIncludeSubfolders Then
        'Fill colFolders with list of subdirectories of strFolder
        strTemp = Dir(strFolder, vbDirectory)
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop

        'Call RecursiveDir for each subfolder in colFolders
        For Each vFolderName In colFolders
            Call RecursiveDir(colFiles, strFolder & vFolderName, strFileSpec, True)
        Next vFolderName
    End If

End Function


Public Function TrailingSlash(strFolder As String) As String
    If Len(strFolder) > 0 Then
        If Right(strFolder, 1) = "\" Then
            TrailingSlash = strFolder
        Else
            TrailingSlash = strFolder & "\"
        End If
    End If
End Function

Private Sub Command1_Click()
If (PageOn + 1) > PageTotal Then
    PageOn = 1
Else
    PageOn = PageOn + 1
    Label2.Caption = "Page: " & PageOn & "/" & PageTotal
    a = Find_Covers()
    Form3.Refresh
    For y = 1 To 12
    Image1(y).Refresh
    Next y
End If
End Sub

Private Sub Form_Load()
If Form1.Combo1.Text = "gb (GameBoy (Color))" Then
    SYSCORE = "gb"
ElseIf Form1.Combo1.Text = "gba (GameBoy Advanced)" Then
    SYSCORE = "gba"
ElseIf Form1.Combo1.Text = "gg (Sega Game Gear)" Then
    SYSCORE = "gg"
ElseIf Form1.Combo1.Text = "lynx (Atari Lynx)" Then
    SYSCORE = "lynx"
ElseIf Form1.Combo1.Text = "md (Sega Genesis/MegaDrive)" Then
    SYSCORE = "md"
ElseIf Form1.Combo1.Text = "nes (Nintendo Entertainment System)" Then
    SYSCORE = "nes"
ElseIf Form1.Combo1.Text = "ngp (Neo Geo Pocket (Color))" Then
    SYSCORE = "ngp"
ElseIf Form1.Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SYSCORE = "pce"
ElseIf Form1.Combo1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SYSCORE = "pce_fast"
ElseIf Form1.Combo1.Text = "pcfx (PC-FX)" Then
    SYSCORE = "pcfx"
ElseIf Form1.Combo1.Text = "psx (Sony PlayStation)" Then
    SYSCORE = "psx"
ElseIf Form1.Combo1.Text = "sms (Sega Master System)" Then
    SYSCORE = "sms"
ElseIf Form1.Combo1.Text = "snes (Super Nintendo Entertainment System)" Then
    SYSCORE = "snes"
'v0.1.8
'Combo1.AddItem "ss (Sega Saturn)", 13
ElseIf Form1.Combo1.Text = "ss (Sega Saturn)" Then
    SYSCORE = "ss"
ElseIf Form1.Combo1.Text = "vb (Virtual Boy)" Then
    SYSCORE = "vb"
ElseIf Form1.Combo1.Text = "wswan (WonderSwan)" Then
    SYSCORE = "wswan"
End If


'12 Games per page..
If SYSCORE = "psx" Then
    MedAdvGAMES = VB.App.Path & "\MedAdvPSX.dat"
    MedAdvCOVERS = VB.App.Path & "\MedAdvPSXCOVERS.dat"
    MedAdvEXT = "cue"
ElseIf SYSCORE = "snes" Then
    MedAdvGAMES = VB.App.Path & "\MedAdvSNES.dat"
    MedAdvCOVERS = VB.App.Path & "\MedAdvSNESCOVERS.dat"
    MedAdvEXT = "smc"
ElseIf SYSCORE = "nes" Then
    MedAdvGAMES = VB.App.Path & "\MedAdvNES.dat"
    MedAdvCOVERS = VB.App.Path & "\MedAdvNESCOVERS.dat"
    MedAdvEXT = "nes"
End If

'Load Settings
Set FSO = CreateObject("Scripting.FileSystemObject")

If FSO.FileExists(VB.App.Path & "\MedAdvCFG.dat") Then

Open VB.App.Path & "\MedAdvCFG.dat" For Input As #1
    For x = 1 To 33
        On Error Resume Next
        Line Input #1, tmp3(x)
    Next x
Close #1

MedEXE = Mid$(tmp3(1), 8, Len(tmp3(1)))
SystemCore = Mid$(tmp3(2), 12, Len(tmp3(2)))
BIOSFILE = Mid$(tmp3(3), 12, Len(tmp3(3)))
BIOSSanity = Mid$(tmp3(4), 12, 1)
ROMFILE = Mid$(tmp3(5), 10, Len(tmp3(5)))
ROMSanity = Mid$(tmp3(6), 11, 1)
Stretch = Mid$(tmp3(7), 9, Len(tmp3(7)))
PixelShader = Mid$(tmp3(8), 13, Len(tmp3(8)))
VideoScaler = Mid$(tmp3(9), 13, Len(tmp3(9)))
Fullscreen = Mid$(tmp3(10), 12, 1)
Frameskip = Mid$(tmp3(11), 11, 1)
TBlur = Mid$(tmp3(12), 7, 1)
TblurAccum = Mid$(tmp3(13), 13, Len(tmp3(13)))
AccumAmount = Mid$(tmp3(14), 11, 1)
VideoIP = Mid$(tmp3(15), 9, 1)
XRes = Mid$(tmp3(16), 6, Len(tmp3(16)))
YRes = Mid$(tmp3(17), 6, Len(tmp3(17)))
ScaleFactor = Mid$(tmp3(18), 13, 1)
LastPath = Mid$(tmp3(19), 10, Len(tmp3(19)))
BiosPathLoad = Mid$(tmp3(20), 14, Len(tmp3(20)))
SavePath = Mid$(tmp3(21), 10, Len(tmp3(21)))
SystemRegion = Mid$(tmp3(22), 14, Len(tmp3(22)))
ROMPathLoad = Mid$(tmp3(23), 9, Len(tmp3(23)))
DisableSound = Mid$(tmp3(24), 14, Len(tmp3(24)))
ForceMono = Mid$(tmp3(25), 11, Len(tmp3(25)))
video_blit_timesync = Mid$(tmp3(26), 21, Len(tmp3(26)))
video_glvsync = Mid$(tmp3(27), 15, Len(tmp3(27)))
untrusted_fip_check = Mid$(tmp3(28), 21, Len(tmp3(28)))
cd_image_memcache = Mid$(tmp3(29), 19, Len(tmp3(29)))
scanlines = Mid$(tmp3(30), 11, Len(tmp3(30)))
axisscale = Mid$(tmp3(31), 11, Len(tmp3(31)))
numplayers = Mid$(tmp3(32), 12, Len(tmp3(32)))
customparams = Mid$(tmp3(33), 14, Len(tmp3(33)))
End If
'End Load Settings
x = 1
y = 1

'Create list downloaded covers
Dim colFiles As New Collection
RecursiveDir colFiles, VB.App.Path & "\covers\" & SYSCORE & "\", "*.jpg", True

Dim vFile As Variant
Close #9
Open MedAdvCOVERS For Output As #9
For Each vFile In colFiles
    'Text1.Text = Text1.Text & vFile & vbCrLf
    Print #9, vFile
Next vFile
Close #9

'Read list of extracted cue sheets
Close #8
Open MedAdvGAMES For Input As #8
While Not EOF(8)
Line Input #8, tmp
GamesList(x) = tmp
x = x + 1
Wend
Close #8

TotalGames = x
PageOn = 1
PageTotal = Int(Round(TotalGames / 12))
Label1.Caption = "Total Games: " & TotalGames
Label2.Caption = "Page: 1/" & PageTotal
x = 1

'Read list of installed covers
Close #10
Open MedAdvCOVERS For Input As #10
While Not EOF(10)
Line Input #10, tmp
CoversList(x) = tmp
x = x + 1
Wend
Close #10

TotalCovers = x
a = Find_Covers()
End Sub
Function Find_Covers()
If PageOn = 1 Then
    y = 1
    For x = 1 To TotalGames - 1
        z = 1
        tmparray(x) = Split(GamesList(x), "\")
        CurrName(x) = tmparray(x)(5)
        tmp = Replace(CurrName(x), ".cue", "")
        tmp = Replace(tmp, ".smc", "")
        tmp = Replace(tmp, ".nes", "")
        tmp = Replace(tmp, ".iso", "")
        tmp = Replace(tmp, ".bin", "")
        tmp = Replace(tmp, " (USA) ", "")
        tmp = Replace(tmp, " (USA)", "")
        tmp = Replace(tmp, " (v1.0)", "")
        tmp = Replace(tmp, " (v1.1)", "")
        tmp = Replace(tmp, " (v1.2)", "")
        tmp = Replace(tmp, " (v1.3)", "")
        tmp = Replace(tmp, " (v1.4)", "")
        tmp = Replace(tmp, " (v1.5)", "")
        tmp = Replace(tmp, " (v1.6)", "")
        tmp = Replace(tmp, "_(Arcade)", "")
        tmp = Replace(tmp, "_(Arcade Mode)", "")
        tmp = Replace(tmp, "(Arcade)", "")
        tmp = Replace(tmp, "(Arcade Mode)", "")
        tmp = Replace(tmp, "_(Simulation)", "")
        tmp = Replace(tmp, "_(Simulation Mode)", "")
        tmp = Replace(tmp, "(Simulation Mode)", "")
        tmp = Replace(tmp, "(Simulation)", "")
        tmp = Replace(tmp, "_(Unl)", "")
        tmp = Replace(tmp, "(Unl)", "")
        tmp = Replace(tmp, "(Disc 1)", "")
        tmp = Replace(tmp, "(Disc 2)", "")
        tmp = Replace(tmp, "(Disc 3)", "")
        tmp = Replace(tmp, "(Disc 4)", "")
        tmp = Replace(tmp, "(Disc 5)", "")
        tmp = Replace(tmp, "(En,Fr,De,Es,It)", "")
        tmp = Replace(tmp, "(En,Fr,De,Sv)", "")
        tmp = Replace(tmp, "(Namco)", "")
        tmp = Replace(tmp, "(Tengen)", "")
        tmp = Replace(tmp, "(Hack)", "")
        tmp = Replace(tmp, "[!]", "")
        tmp = Replace(tmp, "(U)", "")
        tmp = Replace(tmp, "(J)", "")
        tmp = Replace(tmp, "(E)", "")
        tmp = Replace(tmp, "(G)", "")
        tmp = Replace(tmp, "(W)", "")
        tmp = Replace(tmp, "(VC)", "")
        tmp = Replace(tmp, "(JU)", "")
        tmp = Replace(tmp, "(M2)", "")
        tmp = Replace(tmp, "(PRG0)", "")
        tmp = Replace(tmp, "(PRG1)", "")
        tmp = Replace(tmp, "(PC10)", "")
        tmp = Replace(tmp, "[o1]", "")
        tmp = Replace(tmp, "[o2]", "")
        tmp = Replace(tmp, "[o3]", "")
        tmp = Replace(tmp, "[b1]", "")
        tmp = Replace(tmp, "[b2]", "")
        tmp = Replace(tmp, "[b3]", "")
        tmp = Replace(tmp, "[b4]", "")
        tmp = Replace(tmp, "[b5]", "")
        tmp = Replace(tmp, "[b6]", "")
        tmp = Replace(tmp, "[t1]", "")
        tmp = Replace(tmp, "[t2]", "")
        tmp = Replace(tmp, "[t3]", "")
        tmp = Replace(tmp, "[t4]", "")
        tmp = Replace(tmp, "[t5]", "")
        tmp = Replace(tmp, "[t6]", "")
        tmp = Replace(tmp, " - ", "-")
        tmp = Replace(tmp, "  ", " ")
        tmp = Replace(tmp, " ", "_")
            For z = 1 To TotalCovers
            If InStr(1, LCase(CoversList(z)), LCase(tmp), 1) >= 1 Then
                'MsgBox "Cover: " & PSXCovers(x) & " for " & tmp
                If y = 13 Then
                    
                Else
                    Image1(y).Picture = LoadPicture(CoversList(z))
                    Image1(y).Tag = GamesList(x)
                    z = TotalCovers
                    y = y + 1
                End If
            End If
            Next z
        Next x
ElseIf PageOn >= 2 Then
    y = 1
    For x = ((12 * PageOn) - 11) To TotalGames - 1
        z = 1
        tmparray(x) = Split(GamesList(x), "\")
        CurrName(x) = tmparray(x)(5)
        tmp = Replace(CurrName(x), ".cue", "")
        tmp = Replace(tmp, ".smc", "")
        tmp = Replace(tmp, ".nes", "")
        tmp = Replace(tmp, ".iso", "")
        tmp = Replace(tmp, ".bin", "")
        tmp = Replace(tmp, " (USA) ", "")
        tmp = Replace(tmp, " (USA)", "")
        tmp = Replace(tmp, " (v1.0)", "")
        tmp = Replace(tmp, " (v1.1)", "")
        tmp = Replace(tmp, " (v1.2)", "")
        tmp = Replace(tmp, " (v1.3)", "")
        tmp = Replace(tmp, " (v1.4)", "")
        tmp = Replace(tmp, " (v1.5)", "")
        tmp = Replace(tmp, " (v1.6)", "")
        tmp = Replace(tmp, "_(Arcade)", "")
        tmp = Replace(tmp, "_(Arcade Mode)", "")
        tmp = Replace(tmp, "(Arcade)", "")
        tmp = Replace(tmp, "(Arcade Mode)", "")
        tmp = Replace(tmp, "_(Simulation)", "")
        tmp = Replace(tmp, "_(Simulation Mode)", "")
        tmp = Replace(tmp, "(Simulation Mode)", "")
        tmp = Replace(tmp, "(Simulation)", "")
        tmp = Replace(tmp, "_(Unl)", "")
        tmp = Replace(tmp, "(Unl)", "")
        tmp = Replace(tmp, "(Disc 1)", "")
        tmp = Replace(tmp, "(Disc 2)", "")
        tmp = Replace(tmp, "(Disc 3)", "")
        tmp = Replace(tmp, "(Disc 4)", "")
        tmp = Replace(tmp, "(Disc 5)", "")
        tmp = Replace(tmp, "(En,Fr,De,Es,It)", "")
        tmp = Replace(tmp, "(En,Fr,De,Sv)", "")
        tmp = Replace(tmp, "(Namco)", "")
        tmp = Replace(tmp, "(Tengen)", "")
        tmp = Replace(tmp, "(Hack)", "")
        tmp = Replace(tmp, "[!]", "")
        tmp = Replace(tmp, "(U)", "")
        tmp = Replace(tmp, "(J)", "")
        tmp = Replace(tmp, "(E)", "")
        tmp = Replace(tmp, "(G)", "")
        tmp = Replace(tmp, "(W)", "")
        tmp = Replace(tmp, "(VC)", "")
        tmp = Replace(tmp, "(JU)", "")
        tmp = Replace(tmp, "(M2)", "")
        tmp = Replace(tmp, "(PRG0)", "")
        tmp = Replace(tmp, "(PRG1)", "")
        tmp = Replace(tmp, "(PC10)", "")
        tmp = Replace(tmp, "[o1]", "")
        tmp = Replace(tmp, "[o2]", "")
        tmp = Replace(tmp, "[o3]", "")
        tmp = Replace(tmp, "[b1]", "")
        tmp = Replace(tmp, "[b2]", "")
        tmp = Replace(tmp, "[b3]", "")
        tmp = Replace(tmp, "[b4]", "")
        tmp = Replace(tmp, "[b5]", "")
        tmp = Replace(tmp, "[b6]", "")
        tmp = Replace(tmp, "[t1]", "")
        tmp = Replace(tmp, "[t2]", "")
        tmp = Replace(tmp, "[t3]", "")
        tmp = Replace(tmp, "[t4]", "")
        tmp = Replace(tmp, "[t5]", "")
        tmp = Replace(tmp, "[t6]", "")
        tmp = Replace(tmp, " - ", "-")
        tmp = Replace(tmp, "  ", " ")
        tmp = Replace(tmp, " ", "_")
        For z = 1 To TotalCovers
            If InStr(1, LCase(CoversList(z)), LCase(tmp), 1) >= 1 Then
                'MsgBox "Cover: " & PSXCovers(x) & " for " & tmp
                If y = 13 Then
                
                Else
                    Image1(y).Picture = LoadPicture(CoversList(z))
                    Image1(y).Tag = GamesList(x)
                    z = TotalCovers
                    y = y + 1
                End If
            End If
            Next z
        Next x
End If
End Function

Private Sub Image1_Click(Index As Integer)
'Quick Launch
If Form2.Check1.Value = 1 Then
If Form1.Combo1.Text = "gb (GameBoy (Color))" Then
    SYSCORE = "gb"
ElseIf Form1.Combo1.Text = "gba (GameBoy Advanced)" Then
    SYSCORE = "gba"
ElseIf Form1.Combo1.Text = "gg (Sega Game Gear)" Then
    SYSCORE = "gg"
ElseIf Form1.Combo1.Text = "lynx (Atari Lynx)" Then
    SYSCORE = "lynx"
ElseIf Form1.Combo1.Text = "md (Sega Genesis/MegaDrive)" Then
    SYSCORE = "md"
ElseIf Form1.Combo1.Text = "nes (Nintendo Entertainment System)" Then
    SYSCORE = "nes"
ElseIf Form1.Combo1.Text = "ngp (Neo Geo Pocket (Color))" Then
    SYSCORE = "ngp"
ElseIf Form1.Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SYSCORE = "pce"
ElseIf Form1.Combo1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SYSCORE = "pce_fast"
ElseIf Form1.Combo1.Text = "pcfx (PC-FX)" Then
    SYSCORE = "pcfx"
ElseIf Form1.Combo1.Text = "psx (Sony PlayStation)" Then
    SYSCORE = "psx"
ElseIf Form1.Combo1.Text = "sms (Sega Master System)" Then
    SYSCORE = "sms"
ElseIf Form1.Combo1.Text = "snes (Super Nintendo Entertainment System)" Then
    SYSCORE = "snes"
'v0.1.8
'Combo1.AddItem "ss (Sega Saturn)", 13
ElseIf Form1.Combo1.Text = "ss (Sega Saturn)" Then
    SYSCORE = "ss"
ElseIf Form1.Combo1.Text = "vb (Virtual Boy)" Then
    SYSCORE = "vb"
ElseIf Form1.Combo1.Text = "wswan (WonderSwan)" Then
    SYSCORE = "wswan"
End If

If SYSCORE = "psx" Or SYSCORE = "pce" Or SYSCORE = "pce_fast" Or SYSCORE = "ss" Then
    cmdstring = "cmd.exe /c " & Chr(34) & MedEXE & " -loadcd " & SYSCORE
    If SYSCORE = "psx" Then
        If Len(BIOSPATH) > 1 Then
            cmdstring = cmdstring & " -filesys.path_firmware " & Chr(34) & BIOSPATH & Chr(34)
        End If
        If Form1.Check11.Value = 1 Then
            cmdstring = cmdstring & " -psx.bios_na " & Chr(34) & BIOSFILE & Chr(34)
        ElseIf Form1.Check12.Value = 1 Then
            cmdstring = cmdstring & " -psx.bios_jp " & Chr(34) & BIOSFILE & Chr(34)
        ElseIf Form1.Check13.Value = 1 Then
            cmdstring = cmdstring & " -psx.bios_eu " & Chr(34) & BIOSFILE & Chr(34)
        End If
    End If
    If SYSCORE = "pce" Or SYSCORE = "pce_fast" Then
        If Len(BIOSPATH) > 1 Then
            cmdstring = cmdstring & " -filesys.path_firmware " & Chr(34) & BIOSPATH & Chr(34)
        End If
        cmdstring = cmdstring & "-pce.cdbios " & Chr(34) & BIOSFILE & Chr(34)
    End If
    If SYSCORE = "ss" Then
        If Len(BIOSPATH) > 1 Then
            cmdstring = cmdstring & " -filesys.path_firmware " & Chr(34) & BIOSPATH & Chr(34)
        End If
        If Check12.Value = 1 Then
            cmdstring = cmdstring & "-ss.bios_jp " & Chr(34) & BIOSFILE & Chr(34)
        Else
            cmdstring = cmdstring & "-ss.bios_na_eu " & Chr(34) & BIOSFILE & Chr(34)
        End If
    End If
Else
    cmdstring = "cmd.exe /c " & Chr(34) & MedEXE & " -force_module " & SYSCORE
End If

'**
'We'll assume these settings.
If a = a Then
    cmdstring = cmdstring & " -" & SYSCORE & ".videoip 1"
    cmdstring = cmdstring & " -cd.image_memcache 1"
    cmdstring = cmdstring & " -" & SYSCORE & ".stretch full"
    If SYSCORE = "psx" Or SYSCORE = "ss" Then
            cmdstring = cmdstring & " -" & SYSCORE & ".bios_sanity 1"
            cmdstring = cmdstring & " -" & SYSCORE & ".cd_sanity 1"
            cmdstring = cmdstring & " -cd.image_memcache 1"
    End If
End If

'Basic Mode Panel Settings
If Form2.Check23.Value = 1 Then
    cmdstring = cmdstring & " -video.fs 1"
Else
    cmdstring = cmdstring & " -video.fs 0"
End If

If Len(Form2.Text5.Text) > 0 Then
    cmdstring = cmdstring & " -" & SYSCORE & ".xres " & Form2.Text5.Text
End If

If Len(Form2.Text6.Text) > 0 Then
    cmdstring = cmdstring & " -" & SYSCORE & ".yres " & Form2.Text6.Text
End If


'cmdstring = cmdstring & " " & Chr(34) & ROMFILE & Chr(34)
cmdstring = cmdstring & " " & Chr(34) & Image1(Index).Tag & Chr(34)

cmdstring = cmdstring & Chr(34)

If FatalError = False Then
    MsgBox cmdstring
End If

If FatalError = False Then
    Shell (cmdstring)
Else
    FatalError = False
End If

End If
End Sub

Private Sub Quit_Click()
Unload Form1
Unload Form2
Unload Form3
End Sub
