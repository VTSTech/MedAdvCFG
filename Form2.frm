VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "MedAdvCFG v0.2.2 (Mednafen v0.9.38.x Frontend) by Nigel Todman [BASIC MODE]"
   ClientHeight    =   7110
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10875
   LinkTopic       =   "Form2"
   ScaleHeight     =   15979.9
   ScaleMode       =   0  'User
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Play Now!"
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      Text            =   "Not Set"
      Top             =   2520
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "Set"
      Height          =   315
      Left            =   7920
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Text            =   "Not Set"
      Top             =   2880
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "Set"
      Height          =   315
      Left            =   7920
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Text            =   "1366"
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   7320
      TabIndex        =   1
      Text            =   "768"
      Top             =   3360
      Visible         =   0   'False
      Width           =   515
   End
   Begin VB.CheckBox Check23 
      BackColor       =   &H00000000&
      Caption         =   "Fullscreen"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4080
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mednafen EXE:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1920
      TabIndex        =   13
      Top             =   2280
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0.9.38.7-win64 Detected! MD5: 8F0BC836E2B6023371B99E94829B5CF1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3240
      TabIndex        =   12
      Top             =   2280
      Width           =   6225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "BASIC MODE is not Finished yet."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   3420
      TabIndex        =   11
      Top             =   1920
      Width           =   4020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "System BIOS:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "ROM Image:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2075
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Resolution"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   6000
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image15 
      Height          =   1335
      Left            =   4740
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Image Image14 
      Height          =   1335
      Left            =   8880
      Picture         =   "Form2.frx":0BD2
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Image Image13 
      Height          =   1335
      Left            =   6810
      Picture         =   "Form2.frx":457F
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Image Image12 
      Height          =   1335
      Left            =   600
      Picture         =   "Form2.frx":7C10
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image11 
      Height          =   1335
      Left            =   4740
      Picture         =   "Form2.frx":DF37
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image Image10 
      Height          =   1335
      Left            =   2640
      Picture         =   "Form2.frx":1630C
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Image Image9 
      Height          =   1335
      Left            =   600
      Picture         =   "Form2.frx":17A29
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image8 
      Height          =   1335
      Left            =   6810
      Picture         =   "Form2.frx":20B70
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image7 
      Height          =   1335
      Left            =   2670
      Picture         =   "Form2.frx":4309D
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   1335
      Left            =   8880
      Picture         =   "Form2.frx":46FBA
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Image Image5 
      Height          =   1335
      Left            =   6810
      Picture         =   "Form2.frx":4E3A6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   1215
      Left            =   8880
      Picture         =   "Form2.frx":516CB
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   1335
      Left            =   4740
      Picture         =   "Form2.frx":633FA
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   2670
      Picture         =   "Form2.frx":6966B
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   600
      Picture         =   "Form2.frx":6AE27
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu Mode 
      Caption         =   "Mode"
      Begin VB.Menu Advanced 
         Caption         =   "Advanced"
      End
      Begin VB.Menu Basic 
         Caption         =   "Basic"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MedEXE, FSO, tmp, tmp2, tmp3(99), BIOSFILE, BIOSPATH, ROMFILE, SystemCore, SYSCORE, BIOSSanity, ROMSanity, Stretch, PixelShader, VideoScaler, x, y, z
Dim cmdstring, Build, Frameskip, Fullscreen, TBlur, TblurAccum, AccumAmount, VideoIP, ActiveFile, XRes, YRes, ScaleFactor, LastPath, SavePath, BiosPathLoad
Dim ResetBios, ResetRom, ResetSave, FatalError, SystemRegion, SystemRegionLoad, ROMDIR, M3USize, LastFile, VideoDriver
Dim Bilinear, DisableSound, ForceMono, video_blit_timesync, video_glvsync, untrusted_fip_check, cd_image_memcache, scanlines, numplayers, customparams

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Advanced_Click()
Basic.Checked = False
Form1.Visible = True
Form1.Advanced.Checked = True
Form2.Visible = False
End Sub

Private Sub Basic_Click()
Basic.Checked = True
Form1.Visible = False
Form2.Visible = True
End Sub

Private Sub Command1_Click()
If Len(Text1.Text) >= 1 Then
    If Text1.Text = "Not Set" Then
        ResetBios = vbYes
    Else
        If Check15.Value = 1 Then
            a = a
            ResetBios = vbYes
        Else
            ResetBios = MsgBox("Reset Bios?", vbYesNo, "Reset Bios?")
        End If
    End If
    If ResetBios = vbYes Then
        If Len(BiosPathLoad) > 0 Then Dir1.Path = BiosPathLoad
        BIOSFILE = ""
    Else
        a = Form1.Validate_Bios()
    End If
End If
End Sub

Private Sub Command2_Click()
Form2.Width = 12945
ActiveFile = "ROM"
If Len(Text2.Text) >= 1 Then
    If Text2.Text = "Not Set" Then
        ResetRom = vbYes
    Else
        If Form1.Check15.Value = 1 Then
            a = a
            ResetRom = vbYes
        Else
            ResetRom = MsgBox("Reset Rom?", vbYesNo, "Reset Rom?")
        End If
    End If
    If ResetRom = vbYes Then
        If Len(ROMDIR) > 0 Then Dir1.Path = ROMDIR
        ROMFILE = ""
    Else
        a = Form1.Validate_Rom()
    End If
End If
End Sub

Private Sub Command3_Click()
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
ElseIf Combo1.Text = "snes (Super Nintendo Entertainment System)" Then
    SYSCORE = "snes"
'v0.1.8
'Combo1.AddItem "ss (Sega Saturn)", 13
ElseIf Combo1.Text = "ss (Sega Saturn)" Then
    SYSCORE = "ss"
ElseIf Combo1.Text = "vb (Virtual Boy)" Then
    SYSCORE = "vb"
ElseIf Combo1.Text = "wswan (WonderSwan)" Then
    SYSCORE = "wswan"
End If

If SYSCORE = "psx" Or SYSCORE = "pce" Or SYSCORE = "pce_fast" Or SYSCORE = "ss" Then
    cmdstring = "cmd.exe /c " & Chr(34) & MedEXE & " -loadcd " & SYSCORE
    If SYSCORE = "psx" Then
        If Len(BIOSPATH) > 1 Then
            cmdstring = cmdstring & " -filesys.path_firmware " & Chr(34) & BIOSPATH & Chr(34)
        End If
        If Check11.Value = 1 Then
            cmdstring = cmdstring & " -psx.bios_na " & Chr(34) & BIOSFILE & Chr(34)
        ElseIf Check12.Value = 1 Then
            cmdstring = cmdstring & " -psx.bios_jp " & Chr(34) & BIOSFILE & Chr(34)
        ElseIf Check13.Value = 1 Then
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
End Sub

Private Sub Form_Load()
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

a = Form1.Validate_MedEXE()
a = Form1.Validate_Rom()
a = Form1.Validate_Bios()
End If
End Sub

'**
'Combo1.AddItem "gb (GameBoy (Color))", 0
'Combo1.AddItem "gba (GameBoy Advanced)", 1
'Combo1.AddItem "gg (Sega Game Gear)", 2
'Combo1.AddItem "lynx (Atari Lynx)", 3
'Combo1.AddItem "md (Sega Genesis/MegaDrive)", 4
'Combo1.AddItem "nes (Nintendo Entertainment System)", 5
'Combo1.AddItem "ngp (Neo Geo Pocket (Color))", 6
'Combo1.AddItem "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)", 7
'Combo1.AddItem "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)", 8
'Combo1.AddItem "pcfx (PC-FX)", 9
'Combo1.AddItem "psx (Sony PlayStation)", 10
'Combo1.AddItem "sms (Sega Master System)", 11
'Combo1.AddItem "snes (Super Nintendo Entertainment System)", 12
'Combo1.AddItem "ss (Sega Saturn)", 13
'Combo1.AddItem "vb (Virtual Boy)", 14
'Combo1.AddItem "wswan (WonderSwan)", 15
'**
Private Sub Image1_Click()
'PSX
Form1.Combo1.Text = "psx (Sony PlayStation)"
a = Hide_Buttons()
Image1.Visible = True
End Sub

Public Function Hide_Buttons()
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
Image8.Visible = False
Image9.Visible = False
Image10.Visible = False
Image11.Visible = False
Image12.Visible = False
Image13.Visible = False
Image14.Visible = False
Image15.Visible = False
b = Step2()
End Function
Public Function Step2()
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Text1.Visible = True
Text2.Visible = True
Label4.Visible = True
Label7.Visible = True
Label26.Visible = True
Check23.Visible = True
Text5.Visible = True
Text6.Visible = True
End Function
Private Sub Image10_Click()
'NeoGeo
a = Hide_Buttons()
Image10.Visible = True
End Sub

Private Sub Image11_Click()
'Sega Master System
a = Hide_Buttons()
Image11.Visible = True
End Sub

Private Sub Image12_Click()
'PC-Engine/Turbografx
a = Hide_Buttons()
Image12.Visible = True
End Sub

Private Sub Image13_Click()
'PC-FX
a = Hide_Buttons()
Image13.Visible = True
End Sub

Private Sub Image14_Click()
'VB
a = Hide_Buttons()
Image14.Visible = True
End Sub

Private Sub Image15_Click()
'Wonderswan
a = Hide_Buttons()
Image15.Visible = True
End Sub

Private Sub Image2_Click()
'NES
a = Hide_Buttons()
Image2.Visible = True
End Sub

Private Sub Image3_Click()
'SNES
a = Hide_Buttons()
Image3.Visible = True
End Sub

Private Sub Image4_Click()
'Genesis
a = Hide_Buttons()
Image4.Visible = True
End Sub

Private Sub Image5_Click()
'Gameboy
a = Hide_Buttons()
Image5.Visible = True
End Sub

Private Sub Image6_Click()
'GameBoy Adv
a = Hide_Buttons()
Image6.Visible = True
End Sub

Private Sub Image7_Click()
'GameGear
a = Hide_Buttons()
Image7.Visible = True
End Sub

Private Sub Image8_Click()
'Saturn
a = Hide_Buttons()
Image8.Visible = True
End Sub

Private Sub Image9_Click()
'Atari
a = Hide_Buttons()
Image9.Visible = True
End Sub

Private Sub Quit_Click()
Unload Form1
Unload Form2
End Sub
