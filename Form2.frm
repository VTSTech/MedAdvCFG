VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "MedAdvCFG v0.0.0 (Mednafen v0.9.x.x Frontend) by Nigel Todman [BASIC MODE]"
   ClientHeight    =   5250
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   11799.5
   ScaleMode       =   0  'User
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00000000&
      Caption         =   "Back"
      Height          =   315
      Left            =   10320
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Quick Launch (Skip Info)"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3720
      TabIndex        =   20
      Top             =   3360
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   9840
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "Form2.frx":0000
      Top             =   6600
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   11160
      TabIndex        =   16
      Top             =   450
      Width           =   3615
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2820
      Left            =   11160
      TabIndex        =   15
      Top             =   2280
      Width           =   3615
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11160
      TabIndex        =   14
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "List Games"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   3600
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Not Set"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3240
      TabIndex        =   18
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "System BIOS:"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   2040
      TabIndex        =   17
      Top             =   2280
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mednafen EXE:"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   1920
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0.9.41.0-win64 Detected! MD5: 6AADC9A8A196DA610E6DB43367B339B4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3240
      TabIndex        =   12
      Top             =   2040
      Width           =   6135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "BASIC MODE is not Finished yet."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3720
      TabIndex        =   11
      Top             =   0
      Width           =   3480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "System BIOS:"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "ROM Folder:"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   2070
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Resolution"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   6000
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image15 
      Height          =   1335
      Left            =   4440
      Picture         =   "Form2.frx":0007
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Image Image14 
      Height          =   1335
      Left            =   9000
      Picture         =   "Form2.frx":0BD9
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Image Image13 
      Height          =   1335
      Left            =   6360
      Picture         =   "Form2.frx":4586
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Image Image12 
      Height          =   1335
      Left            =   240
      Picture         =   "Form2.frx":5A41
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Image11 
      Height          =   1335
      Left            =   4440
      Picture         =   "Form2.frx":BD68
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image10 
      Height          =   1575
      Left            =   2040
      Picture         =   "Form2.frx":1413D
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Image Image9 
      Height          =   1335
      Left            =   120
      Picture         =   "Form2.frx":1585A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image8 
      Height          =   1335
      Left            =   7080
      Picture         =   "Form2.frx":1E9A1
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Image Image7 
      Height          =   1335
      Left            =   1680
      Picture         =   "Form2.frx":40ECE
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   1575
      Left            =   9000
      Picture         =   "Form2.frx":43473
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   1335
      Left            =   6480
      Picture         =   "Form2.frx":4A85F
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   8280
      Picture         =   "Form2.frx":4EC15
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2535
   End
   Begin VB.Image Image3 
      Height          =   1335
      Left            =   3840
      Picture         =   "Form2.frx":5E225
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   1680
      Picture         =   "Form2.frx":5F682
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1725
      Left            =   120
      Picture         =   "Form2.frx":655C9
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1665
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Save 
         Caption         =   "Save Settings"
      End
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
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu Documentation 
         Caption         =   "Documentation"
      End
   End
   Begin VB.Menu Chat 
      Caption         =   "Chat"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MedEXE, FSO, tmp, tmp2, tmp3(99), BIOSFILE, BIOSPATH, ROMFILE, SystemCore, SysCore, BIOSSanity, ROMSanity, Stretch, PixelShader, VideoScaler, x, y, z
Dim cmdstring, Build, Frameskip, Fullscreen, TBlur, TblurAccum, AccumAmount, VideoIP, ActiveFile, XRes, YRes, ScaleFactor, LastPath, SavePath, BiosPathLoad
Dim ResetBios, ResetRom, ResetRomDir, ResetSave, FatalError, SystemRegion, SystemRegionLoad, ROMDIR, M3USize, LastFile, VideoDriver
Dim Bilinear, DisableSound, ForceMono, video_blit_timesync, video_glvsync, untrusted_fip_check, cd_image_memcache, scanlines, numplayers, customparams
Dim MedAdvGAMES, MedAdvCOVERS, MedAdvEXT, BasicModeFolder, LogoTop, LogoLeft, QuickLaunch, MedMD5, MedDat
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub about_Click()
MsgBox "MedAdvCFG v" & Build & " (Mednafen v0.9.x.x Frontend)" & vbCrLf & "Written by Nigel Todman (nigel@nigeltodman.com)" & vbCrLf & "Primarily written as a PSX Frontend." & vbCrLf & "Tested with the following System Cores:" & vbCrLf & "GB, GBA, GG, MD, NES, PCE, PCE_FAST, PSX, SNES, SS, VB" & vbCrLf & vbCrLf & "Homepage: www.NigelTodman.com" & vbCrLf & "Facebook: facebook.com/nigel.todman.3" & vbCrLf & "Twitter: @Veritas_83" & vbCrLf & "YouTube: Veritas0923" & vbCrLf & "BTC: 18j2Env7QokhGG5MccS3LPBKnjsko6u4NQ"
End Sub

Private Sub Advanced_Click()
Basic.Checked = False
Form1.Basic.Checked = False
Form1.Visible = True
Form1.Advanced.Checked = True
Form2.Visible = False
End Sub

Private Sub Basic_Click()
Basic.Checked = True
Form1.Visible = False
Form2.Visible = True
End Sub

Private Sub Chat_Click()
Shell ("cmd.exe /c start http://bit.ly/2k5E1Xq"), vbHide
End Sub

Private Sub Command1_Click()
Form2.Width = 15075
ActiveFile = "BIOS"
If Len(Text1.Text) >= 1 Then
    If Text1.Text = "Not Set" Then
        ResetBios = vbYes
    Else
        ResetBios = MsgBox("Reset Bios?", vbYesNo, "Reset Bios?")
    End If
    If ResetBios = vbYes Then
        If Len(BiosPathLoad) > 0 Then Dir1.Path = BiosPathLoad
        BIOSFILE = ""
    Else
        a = Validate_Bios()
    End If
End If
End Sub
Function Validate_Bios()
    If FSO.FileExists(BIOSPATH & "\" & BIOSFILE) = True Then
        tmp = Form1.CalcMD5(Form1.ShortPath(BIOSPATH & "\" & BIOSFILE))
        Text1.Text = BIOSPATH & "\" & BIOSFILE
    End If
    If FSO.FileExists(BIOSFILE) = True Then
        tmp = Form1.CalcMD5(Form1.ShortPath(BIOSFILE))
        Text1.Text = BIOSFILE
    End If

    BiosDat = VB.App.Path & "\dat\bios.dat"
    If FSO.FileExists(BiosDat) = True Then
        Close #16
        Open BiosDat For Input As #16
        Do
            Line Input #16, BiosMD5
            If LCase(tmp) = Mid$(BiosMD5, 1, 32) Then
                'MsgBox Mid$(MedMD5, 34, Len(MedMD5))
                Label6.Caption = Mid$(BiosMD5, 34, Len(BiosMD5))
            End If
        Loop Until EOF(16)
        If Label6.Caption = "Not Set" Then
            Label6.Caption = "Set a BIOS"
        End If
    Else
        MsgBox "Fatal Error: bios.dat verification missing. Reinstall MedAdvCFG"
        Unload Form1
    End If
Validate_Bios = tmp
End Function
Function Validate_MedEXE()
If FSO.FileExists(MedEXE) = True Then
    MedDat = VB.App.Path & "\dat\Mednafen.dat"
    If FSO.FileExists(MedDat) = True Then
        tmp = Form1.CalcMD5(Form1.ShortPath(MedEXE))
        Close #16
        Open MedDat For Input As #16
        Do
            Line Input #16, MedMD5
            If tmp = Mid$(MedMD5, 1, 32) Then
                'MsgBox Mid$(MedMD5, 34, Len(MedMD5))
                Label2.Caption = Mid$(MedMD5, 34, Len(MedMD5)) & " Detected! MD5: " & Mid$(MedMD5, 1, 32)
            End If
        Loop Until EOF(16)
    Else
        MsgBox "Fatal Error: Mednafen.dat verification missing. Reinstall MedAdvCFG"
        Unload Form1
    End If
End If
Validate_MedEXE = tmp
End Function
Private Sub Command2_Click()
Form2.Width = 15075
ActiveFile = "ROMDIR"
If Len(Text2.Text) >= 1 Then
    If Text2.Text = "Not Set" Then
        ResetRomDir = vbYes
    Else
        ResetRomDir = MsgBox("Reset Rom Folder?", vbYesNo, "Reset Rom Folder?")
    End If
End If
If ResetRomDir = vbYes Then
   If Len(ROMDIR) > 0 Then Dir1.Path = ROMDIR
        ROMFILE = ""
    Else
        Text2.Text = Dir1.Path
        ROMDIR = Dir1.Path
    End If
End Sub
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
Private Sub Command3_Click()
'Text3.Visible = True
ROMDIR = Text2.Text
CurrFolder = Dir(Text2.Text, vbDirectory)

If SysCore = "gb" Then Form1.Combo1.Text = "gb (GameBoy (Color))"
If SysCore = "gba" Then Form1.Combo1.Text = "gba (GameBoy Advanced)"
If SysCore = "gg" Then Form1.Combo1.Text = "gg (Sega Game Gear)"
If SysCore = "lynx" Then Form1.Combo1.Text = "lynx (Atari Lynx)"
If SysCore = "md" Then Form1.Combo1.Text = "md (Sega Genesis/MegaDrive)"
If SysCore = "nes" Then Form1.Combo1.Text = "nes (Nintendo Entertainment System)"
If SysCore = "ngp" Then Form1.Combo1.Text = "ngp (Neo Geo Pocket (Color))"
If SysCore = "pce" Then Form1.Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)"
If SysCore = "pce_fast" Then Form1.Combo1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)"
If SysCore = "pcfx" Then Form1.Combo1.Text = "pcfx (PC-FX)"
If SysCore = "psx" Then Form1.Combo1.Text = "psx (Sony PlayStation)"
If SysCore = "sms" Then Form1.Combo1.Text = "sms (Sega Master System)"
If SysCore = "snes" Then Form1.Combo1.Text = "snes (Super Nintendo Entertainment System)"
If SysCore = "ss" Then Form1.Combo1.Text = "ss (Sega Saturn)"
If SysCore = "vb" Then Form1.Combo1.Text = "vb (Virtual Boy)"
If SysCore = "wswan" Then Form1.Combo1.Text = "wswan (WonderSwan)"

'Create list of extracted cue sheets
Dim colFiles As New Collection

If SysCore = "psx" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvPSX.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvPSXCOVERS.dat"
    MedAdvEXT = "cue"
ElseIf SysCore = "snes" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvSNES.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvSNESCOVERS.dat"
    MedAdvEXT = "smc"
ElseIf SysCore = "nes" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvNES.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvNESCOVERS.dat"
    MedAdvEXT = "nes"
ElseIf SysCore = "ss" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvSS.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvSSCOVERS.dat"
    MedAdvEXT = "cue"
ElseIf SysCore = "gba" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvGBA.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvGBACOVERS.dat"
    MedAdvEXT = "gba"
ElseIf SysCore = "gb" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvGB.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvGBCOVERS.dat"
    MedAdvEXT = "gbc"
ElseIf SysCore = "gg" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvGG.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvGGCOVERS.dat"
    MedAdvEXT = "gg"
ElseIf SysCore = "pce" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvPCE.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvPCECOVERS.dat"
    MedAdvEXT = "cue"
ElseIf SysCore = "pce_fast" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvPCE.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvPCECOVERS.dat"
    MedAdvEXT = "cue"
ElseIf SysCore = "md" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvMD.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvMDCOVERS.dat"
    MedAdvEXT = "bin"
ElseIf SysCore = "lynx" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvLYNX.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvLYNXCOVERS.dat"
    MedAdvEXT = "lnx"
ElseIf SysCore = "vb" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvVB.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvVBCOVERS.dat"
    MedAdvEXT = "vb"
End If

RecursiveDir colFiles, Text2.Text, "*." & MedAdvEXT, True

If FSO.FileExists(MedAdvGAMES) = False Then
    'MsgBox "cmd.exe /c echo " & Chr(34) & Chr(34) & " >> " & Chr(34) & MedAdvGAMES & Chr(34)
    Shell ("cmd.exe /c echo " & Chr(34) & Chr(34) & " >> " & Chr(34) & MedAdvGAMES & Chr(34))
End If

Dim vFile As Variant
Close #7
Open MedAdvGAMES For Output As #7
For Each vFile In colFiles
    Text3.Text = Text3.Text & vFile & vbCrLf
    Print #7, vFile
Next vFile
Close #7
a = Save_Settings()
Unload Form3
Form3.Refresh
Form2.Visible = False
Form3.Visible = True
End Sub
Private Sub ListFiles(strPath As String, Optional Extention As String)
'Leave Extention blank for all files
    Dim File As String
    
    If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    If Trim$(Extention) = "" Then
        Extention = "*.*"
    ElseIf Left$(Extention, 2) <> "*." Then
        Extention = "*." & Extention
    End If
    
    File = Dir$(strPath & Extention)
    Do While Len(File)
        Text3.Text = Text3.Text & File & vbCrLf
        File = Dir$
    Loop
End Sub

Private Sub Command4_Click()
a = Unhide_Buttons()
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
If ActiveFile = "SAVE" Then
    tmp2 = MsgBox("Set Path: " & Dir1.Path, vbYesNo, "Set this path?")
    If tmp2 = vbYes Then
        Text7.Text = Dir1.Path
        SavePath = Text7.Text
        Form1.Width = 9240
        ActiveFile = "None"
        tmp2 = ""
    End If
ElseIf ActiveFile = "ROMDIR" Then
    tmp2 = MsgBox("Set Path: " & Dir1.Path, vbYesNo, "Set this path?")
    If tmp2 = vbYes Then
        Text2.Text = Dir1.Path
        RomPath = Text2.Text
        ROMDIR = Text2.Text
        Form1.Width = 9240
        ActiveFile = "None"
        tmp2 = ""
    End If
End If
End Sub

Private Sub Documentation_Click()
Shell "cmd.exe /c start https://mednafen.github.io/documentation/", vbHide
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
If ActiveFile = "MEDEXE" Then
        If Check15.Value = 1 Then
            a = a
            tmp2 = vbYes
        Else
            tmp2 = MsgBox("Set File: " & File1.FileName, vbYesNo, "Set this file?")
        End If
    If tmp2 = vbYes Then
        MedEXE = Dir1.Path & "\" & File1.FileName
        Form1.Width = 9240
        ActiveFile = "None"
        tmp2 = ""
        a = Validate_MedEXE()
    End If
End If

If ActiveFile = "BIOS" Then
            tmp2 = MsgBox("Set File: " & File1.FileName, vbYesNo, "Set this file?")
    If tmp2 = vbYes Then
        BIOSPATH = Dir1.Path
        BIOSFILE = File1.FileName
        Text1.Text = BIOSPATH & "\" & BIOSFILE
        Form1.Width = 9240
        ActiveFile = "None"
        tmp2 = ""
        a = Validate_Bios()
    End If
End If

If ActiveFile = "ROM" Then
        If Check15.Value = 1 Then
            a = a
            tmp2 = vbYes
        Else
            tmp2 = MsgBox("Set File: " & File1.FileName, vbYesNo, "Set this file?")
        End If
    If tmp2 = vbYes Then
        Text2.Text = Dir1.Path & "\" & File1.FileName
        ROMDIR = Dir1.Path
        ROMFILE = Text2.Text
        'Form1.Width = 12735
        Form1.Width = 9240
        ActiveFile = "None"
        tmp2 = ""
        a = Validate_Rom()
    End If
End If

If ActiveFile = "M3U" Then
        If z = 0 Then
        If Check15.Value = 1 Then
            a = a
            tmp2 = vbYes
        Else
            tmp2 = MsgBox("Set Disc 1: " & File1.FileName, vbYesNo, "Set this file?")
        End If
            If tmp2 = vbYes Then
                Print #2, Dir1.Path & "\" & File1.FileName
                tmp1 = File1.FileName
            End If
            tmp2 = ""
            z = z + 1
            MsgBox ("Now select the next disc")
        ElseIf z >= 1 And z <= Val(M3USize) Then
            Do
                If File1.FileName <> tmp1 And Len(File1.FileName) > 1 And z <= Val(M3USize) And File1.FileName <> LastFile Then
                    If Check15.Value = 1 Then
                        a = a
                        tmp2 = vbYes
                    Else
                        tmp2 = MsgBox("Set Disc " & z + 1 & ": " & File1.FileName, vbYesNo, "Set this file?")
                    End If
                    LastFile = File1.FileName
                    If tmp2 = vbYes Then
                        Print #2, Dir1.Path & "\" & File1.FileName
                        z = z + 1
                        If z <> Val(M3USize) Then
                            MsgBox ("Now select the next disc")
                        End If
                    ElseIf tmp2 = vbNo Then
                        File1.FileName = ""
                    End If
                End If
                DoEvents
            Loop Until z = Val(M3USize) Or z > Val(M3USize)
            tmp2 = ""
            ActiveFile = "None"
            Form1.Width = 9240
            Close #2
            MsgBox "M3U Generated! Written to: " & VB.App.Path & "\multi.m3u" & vbCrLf & vbCrLf & "You should rename your multi.m3u" & vbCrLf & vbCrLf & "Your Memory Card and Save State filenames will be based on it." & vbCrLf & vbCrLf & "Use F8 to 'Eject/Close' and F6 to cycle thru the Disc Set"
            Text2.Text = VB.App.Path & "\multi.m3u"
        End If
End If
End Sub

Private Sub Form_Load()
Build = Form1.GetBuild()
Form2.Caption = "MedAdvCFG v" & Build & " (Mednafen v0.9.x.x Frontend) by Nigel Todman [BASIC MODE]"
Form2.Width = 11145
Form2.Height = 6120
Basic.Checked = True
Set FSO = CreateObject("Scripting.FileSystemObject")
a = LoadSettings()

If SysCore = "psx" Or SysCore = "ss" Or SysCore = "pce" Then
    If Len(BIOSFILE) = 0 Then
        MsgBox "Please select a BIOS"
    End If
End If
Text1.Text = BIOSFILE
Text2.Text = BasicModeFolder
Text5.Text = XRes
Text6.Text = YRes

Dir1.Path = LastPath
File1.Path = LastPath

ROMDIR = BasicModeFolder
BIOSPATH = BiosPathLoad
a = Validate_MedEXE()
a = Validate_Bios()
Label6.Visible = False
Label5.Visible = False
Label3.Visible = False
Label2.Visible = False
End Sub
Function LoadSettings()
'Load Settings
Set FSO = CreateObject("Scripting.FileSystemObject")

If FSO.FileExists(VB.App.Path & "\MedAdvCFG.dat") Then
Close #66
Open VB.App.Path & "\MedAdvCFG.dat" For Input As #66
    For x = 1 To 35
        On Error Resume Next
        Line Input #66, tmp3(x)
    Next x
Close #66

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
BasicModeFolder = Mid$(tmp3(34), 17, Len(tmp3(34)))
QuickLaunch = Mid$(tmp3(35), 13, 1)
End If
'End Load Settings
End Function

Private Sub Image1_Click()
'PSX
Form1.Combo1.Text = "psx (Sony PlayStation)"
a = Hide_Buttons()
LogoTop = Image1.Top
LogoLeft = Image1.Left
Image1.Left = 4740
Image1.Top = 809.109
Image1.Visible = True
SysCore = "psx"
End Sub
Function Validate_Rom()
If Check9.Value = 1 Then
    If FSO.FileExists(ROMFILE) = True Then
        tmp = Form1.CalcMD5(Form1.ShortPath(ROMFILE))
        Text2.Text = ROMFILE
        Label9.Caption = "MD5: " & tmp
    End If
Else
    Label9.Caption = "MD5: ROM MD5 Disabled"
End If
Validate_Rom = tmp
End Function
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
Public Function Unhide_Buttons()
Image1.Visible = True
Image2.Visible = True
Image3.Visible = True
Image4.Visible = True
Image5.Visible = True
Image6.Visible = True
Image7.Visible = True
Image8.Visible = True
Image9.Visible = True
Image10.Visible = True
Image11.Visible = True
Image12.Visible = True
Image13.Visible = True
Image14.Visible = True
Image15.Visible = True
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Text1.Visible = False
Text2.Visible = False
Label4.Visible = False
Label7.Visible = False
Label26.Visible = False
Check1.Visible = False
Check23.Visible = False
Text5.Visible = False
Text6.Visible = False
Label6.Visible = False
Label5.Visible = False
Label3.Visible = False
Label2.Visible = False

If SysCore = "psx" Then
    Image1.Left = LogoLeft
    Image1.Top = LogoTop
ElseIf SysCore = "nes" Then
    Image2.Left = LogoLeft
    Image2.Top = LogoTop
ElseIf SysCore = "snes" Then
    Image3.Left = LogoLeft
    Image3.Top = LogoTop
ElseIf SysCore = "md" Then
    Image4.Left = LogoLeft
    Image4.Top = LogoTop
ElseIf SysCore = "gb" Then
    Image5.Left = LogoLeft
    Image5.Top = LogoTop
ElseIf SysCore = "gba" Then
    Image6.Left = LogoLeft
    Image6.Top = LogoTop
ElseIf SysCore = "gg" Then
    Image7.Left = LogoLeft
    Image7.Top = LogoTop
ElseIf SysCore = "ss" Then
    Image8.Left = LogoLeft
    Image8.Top = LogoTop
ElseIf SysCore = "lynx" Then
    Image9.Left = LogoLeft
    Image9.Top = LogoTop
ElseIf SysCore = "ngp" Then
    Image10.Left = LogoLeft
    Image10.Top = LogoTop
ElseIf SysCore = "sms" Then
    Image11.Left = LogoLeft
    Image11.Top = LogoTop
ElseIf SysCore = "pce" Then
    Image12.Left = LogoLeft
    Image12.Top = LogoTop
ElseIf SysCore = "pcfx" Then
    Image13.Left = LogoLeft
    Image13.Top = LogoTop
ElseIf SysCore = "vb" Then
    Image14.Left = LogoLeft
    Image14.Top = LogoTop
ElseIf SysCore = "wswan" Then
    Image15.Left = LogoLeft
    Image15.Top = LogoTop
End If
End Function
Public Function Step2()

Command2.Visible = True
Command3.Visible = True
Command4.Visible = True

Text2.Visible = True
Label7.Visible = True
Label26.Visible = True
Check1.Visible = True
Check23.Visible = True
Text5.Visible = True
Text6.Visible = True

SysCore = Form1.SetSysCore
If SysCore = "psx" Or SysCore = "ss" Or SysCore = "pce" Then
    Label6.Visible = True
    Label5.Visible = True
    Label4.Visible = True
    Text1.Visible = True
    Command1.Visible = True
End If

If QuickLaunch = 0 Then Check1.Value = 0
If QuickLaunch = 1 Then Check1.Value = 1

Label3.Visible = True
Label2.Visible = True
End Function
Private Sub Image10_Click()
'NeoGeo
a = Hide_Buttons()
Image10.Visible = True
SysCore = "ngp"
Form1.Combo1.Text = "ngp (Neo Geo Pocket (Color))"
LogoTop = Image10.Top
LogoLeft = Image10.Left
Image10.Left = 4740
Image10.Top = 809.109
End Sub
Private Sub Image11_Click()
'Sega Master System
a = Hide_Buttons()
Image11.Visible = True
SysCore = "sms"
Form1.Combo1.Text = "sms (Sega Master System)"
LogoTop = Image11.Top
LogoLeft = Image11.Left
Image11.Left = 4740
Image11.Top = 809.109
End Sub
Private Sub Image12_Click()
'PC-Engine/Turbografx
a = Hide_Buttons()
Image12.Visible = True
SysCore = "pce"
Form1.Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)"
LogoTop = Image12.Top
LogoLeft = Image12.Left
Image12.Left = 4740
Image12.Top = 809.109
End Sub
Private Sub Image13_Click()
'PC-FX
a = Hide_Buttons()
Image13.Visible = True
SysCore = "pcfx"
Form1.Combo1.Text = "pcfx (PC-FX)"
LogoTop = Image13.Top
LogoLeft = Image13.Left
Image13.Left = 4740
Image13.Top = 809.109
End Sub
Private Sub Image14_Click()
'VB
a = Hide_Buttons()
Image14.Visible = True
SysCore = "vb"
Form1.Combo1.Text = "vb (Virtual Boy)"
LogoTop = Image14.Top
LogoLeft = Image14.Left
Image14.Left = 4740
Image14.Top = 809.109
End Sub
Private Sub Image15_Click()
'Wonderswan
a = Hide_Buttons()
Image15.Visible = True
SysCore = "wswan"
Form1.Combo1.Text = "wswan (WonderSwan)"
LogoTop = Image5.Top
LogoLeft = Image5.Left
Image15.Left = 4740
Image15.Top = 809.109
End Sub
Private Sub Image2_Click()
'NES
a = Hide_Buttons()
Image2.Visible = True
SysCore = "nes"
Form1.Combo1.Text = "nes (Nintendo Entertainment System)"
LogoTop = Image2.Top
LogoLeft = Image2.Left
Image2.Left = 4740
Image2.Top = 809.109
End Sub
Private Sub Image3_Click()
'SNES
a = Hide_Buttons()
Image3.Visible = True
SysCore = "snes"
Form1.Combo1.Text = "snes (Super Nintendo Entertainment System)"
LogoTop = Image3.Top
LogoLeft = Image3.Left
Image3.Left = 4740
Image3.Top = 809.109
End Sub
Private Sub Image4_Click()
'Genesis
a = Hide_Buttons()
Image4.Visible = True
SysCore = "md"
Form1.Combo1.Text = "md (Sega Genesis/MegaDrive)"
SysCore = Form1.SetSysCore
LogoTop = Image4.Top
LogoLeft = Image4.Left
Image4.Left = 4740
Image4.Top = 809.109
End Sub
Private Sub Image5_Click()
'Gameboy
a = Hide_Buttons()
Image5.Visible = True
SysCore = "gb"
Form1.Combo1.Text = "gb (GameBoy (Color))"
LogoTop = Image5.Top
LogoLeft = Image5.Left
Image5.Left = 4740
Image5.Top = 809.109
End Sub
Private Sub Image6_Click()
'GameBoy Adv
a = Hide_Buttons()
Image6.Visible = True
SysCore = "gba"
Form1.Combo1.Text = "gba (GameBoy Advanced)"
LogoTop = Image6.Top
LogoLeft = Image6.Left
Image6.Left = 4740
Image6.Top = 809.109
End Sub
Private Sub Image7_Click()
'GameGear
a = Hide_Buttons()
Image7.Visible = True
SysCore = "gg"
Form1.Combo1.Text = "gg (Sega Game Gear)"
LogoTop = Image7.Top
LogoLeft = Image7.Left
Image7.Left = 4740
Image7.Top = 809.109
End Sub
Private Sub Image8_Click()
'Saturn
a = Hide_Buttons()
Image8.Visible = True
SysCore = "ss"
Form1.Combo1.Text = "ss (Sega Saturn)"
LogoTop = Image8.Top
LogoLeft = Image8.Left
Image8.Left = 4740
Image8.Top = 809.109
End Sub
Private Sub Image9_Click()
'Atari
a = Hide_Buttons()
Image9.Visible = True
SysCore = "lynx"
Form1.Combo1.Text = "lynx (Atari Lynx)"
LogoTop = Image9.Top
LogoLeft = Image9.Left
Image9.Left = 4740
Image9.Top = 809.109
End Sub
Private Sub Quit_Click()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
End Sub
Function Save_Settings()
Set FSO = CreateObject("Scripting.FileSystemObject")
Close #6
Open VB.App.Path & "\MedAdvCFG.dat" For Output As #6
    Print #6, "MedEXE=" & MedEXE
    Print #6, "SystemCore=" & Form1.Combo1.Text
    If FSO.FileExists(BIOSPATH & "\" & BIOSFILE) Then
        Print #6, "SystemBIOS=" & BIOSPATH & "\" & BIOSFILE
    ElseIf FSO.FileExists(BIOSFILE) Then
        Print #6, "SystemBIOS=" & BIOSFILE
    Else
        Print #6, "SystemBIOS="
    End If
    Print #6, "BIOSSanity=0"
    Print #6, "RomImage=" & ROMFILE
    Print #6, "ROMSanity=0"
    Print #6, "Stretch=full - Full"
    Print #6, "PixelShader=None - None/Disabled"
    Print #6, "VideoScaler=None - None/Disabled"
    Print #6, "Fullscreen=1"
    Print #6, "Frameskip=0"
    Print #6, "Tblur=0"
    Print #6, "TblurAccum=0"
    Print #6, "AccumAmount=50"
    Print #6, "VideoIP=1"
    Print #6, "XRes=" & Text5.Text
    Print #6, "YRes=" & Text6.Text
    Print #6, "ScaleFactor=2"
    Print #6, "LastPath=" & ROMDIR
    Print #6, "BiosPathLoad=" & BIOSPATH
    Print #6, "SavePath=" & VB.App.Path & "\sav\"
    Print #6, "SystemRegion=None"
    Print #6, "RomPath=" & ROMDIR
    Print #6, "DisableSound=0"
    Print #6, "ForceMono=0"
    Print #6, "video.blit_timesync=0"
    Print #6, "video.glvsync=0"
    Print #6, "untrusted_fip_check=1"
    Print #6, "cd.image_memcache=1"
    Print #6, "scanlines=0"
    Print #6, "axisscale=1.00"
    Print #6, "numplayers=1"
    Print #6, "customparams="
    Print #6, "BasicModeFolder=" & Text2.Text
    Print #6, "QuickLaunch=" & Check1.Value
Close #6
End Function
Private Sub Save_Click()
a = Save_Settings()
End Sub


