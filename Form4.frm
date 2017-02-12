VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   Caption         =   "Form4"
   ClientHeight    =   7965
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13125
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form4"
   ScaleHeight     =   7965
   ScaleWidth      =   13125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   2
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "REDUMP Database unverified!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   9000
      TabIndex        =   17
      Top             =   2520
      Width           =   2385
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "REDUMP Database unverified!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   9000
      TabIndex        =   16
      Top             =   3000
      Width           =   2385
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Game ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   15
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "GAMEID"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   14
      Top             =   1320
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Cover Searched:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   13
      Top             =   2040
      Width           =   1365
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "6AADC9A8A196DA610E6DB43367B339B4"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   12
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "BIOS MD5:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   11
      Top             =   2520
      Width           =   810
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "6AADC9A8A196DA610E6DB43367B339B4"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   10
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "BIOS"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   9
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "System BIOS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "6AADC9A8A196DA610E6DB43367B339B4"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   7
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "ROM MD5:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   6
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "System Core:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   5
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "ROM Filename:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "FileName.ext"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   1
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "SysCore"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7920
      TabIndex        =   0
      Top             =   360
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   7695
      Index           =   1
      Left            =   120
      Picture         =   "Form4.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7695
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
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MedEXE, FSO, tmp, tmp2, tmp3(99), BIOSFILE, BIOSPATH, ROMFILE, SystemCore, SysCore, BIOSSanity, ROMSanity, Stretch, PixelShader, VideoScaler, x, y, z
Dim cmdstring, Build, Frameskip, Fullscreen, TBlur, TblurAccum, AccumAmount, VideoIP, ActiveFile, XRes, YRes, ScaleFactor, LastPath, SavePath, BiosPathLoad
Dim ResetBios, ResetRom, ResetRomDir, ResetSave, FatalError, SystemRegion, SystemRegionLoad, ROMDIR, M3USize, LastFile, VideoDriver
Dim Bilinear, DisableSound, ForceMono, video_blit_timesync, video_glvsync, untrusted_fip_check, cd_image_memcache, scanlines, numplayers, customparams
Dim MedAdvGAMES, MedAdvCOVERS, MedAdvEXT, TotalGames, PageTotal, PageOn, TotalCovers, t, FNIndex, CurrFN, CoverFound, CoverSearched
Dim GamesList(9999), CoversList(9999), CurrName(9999)
Dim FileName, tmparray, PSXID, RedumpList, REDUMPMD5, ROMMD5, CUEFILE, BINFILE
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub about_Click()
MsgBox "MedAdvCFG v" & Build & " (Mednafen v0.9.x.x Frontend)" & vbCrLf & "Written by Nigel Todman (nigel@nigeltodman.com)" & vbCrLf & "Primarily written as a PSX Frontend." & vbCrLf & "Tested with the following System Cores:" & vbCrLf & "GB, GBA, GG, MD, NES, PCE, PCE_FAST, PSX, SNES, SS, VB" & vbCrLf & vbCrLf & "Homepage: www.NigelTodman.com" & vbCrLf & "Facebook: facebook.com/nigel.todman.3" & vbCrLf & "Twitter: @Veritas_83" & vbCrLf & "YouTube: Veritas0923" & vbCrLf & "BTC: 18j2Env7QokhGG5MccS3LPBKnjsko6u4NQ"
End Sub

Private Sub Basic_Click()
Basic.Checked = False
Form1.Basic.Checked = False
Form1.Visible = True
Form1.Advanced.Checked = True
Form2.Visible = False
End Sub

Private Sub Chat_Click()
Shell ("cmd.exe /c start http://bit.ly/2k5E1Xq"), vbHide
End Sub

Private Sub Command1_Click()
SysCore = Form1.SetSysCore
If SysCore = "psx" Or SysCore = "pce" Or SysCore = "pce_fast" Or SysCore = "ss" Then
    cmdstring = "cmd.exe /c " & Chr(34) & MedEXE & " -loadcd " & SysCore
    If SysCore = "psx" Then
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
            If SysCore = "pce" Or SysCore = "pce_fast" Then
                If Len(BIOSPATH) > 1 Then
                    cmdstring = cmdstring & " -filesys.path_firmware " & Chr(34) & BIOSPATH & Chr(34)
                End If
                cmdstring = cmdstring & "-pce.cdbios " & Chr(34) & BIOSFILE & Chr(34)
            End If
            If SysCore = "ss" Then
                If Len(BIOSPATH) > 1 Then
                    cmdstring = cmdstring & " -filesys.path_firmware " & Chr(34) & BIOSPATH & Chr(34)
                End If
                If Form1.Check12.Value = 1 Then
                    cmdstring = cmdstring & "-ss.bios_jp " & Chr(34) & BIOSFILE & Chr(34)
                Else
                    cmdstring = cmdstring & "-ss.bios_na_eu " & Chr(34) & BIOSFILE & Chr(34)
                End If
            End If
        Else
            cmdstring = "cmd.exe /c " & Chr(34) & MedEXE & " -force_module " & SysCore
        End If

        '**
        'We'll assume these settings.
        If a = a Then
            cmdstring = cmdstring & " -" & SysCore & ".videoip 1"
            cmdstring = cmdstring & " -cd.image_memcache 1"
            cmdstring = cmdstring & " -" & SysCore & ".stretch full"
            If SysCore = "psx" Or SysCore = "ss" Then
                cmdstring = cmdstring & " -" & SysCore & ".bios_sanity 1"
                cmdstring = cmdstring & " -" & SysCore & ".cd_sanity 1"
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
            cmdstring = cmdstring & " -" & SysCore & ".xres " & Form2.Text5.Text
        End If

        If Len(Form2.Text6.Text) > 0 Then
            cmdstring = cmdstring & " -" & SysCore & ".yres " & Form2.Text6.Text
        End If


        cmdstring = cmdstring & " " & Chr(34) & Image1(1).Tag & Chr(34)

        cmdstring = cmdstring & Chr(34)

        If FatalError = False Then
            'MsgBox cmdstring
        End If

        If FatalError = False Then
            Shell (cmdstring)
        Else
            FatalError = False
        End If
    End Sub

    Private Sub Command2_Click()
    Form3.Visible = True
    Form4.Visible = False
End Sub
Public Function GetPSXID()

Set FSO = CreateObject("Scripting.FileSystemObject")

PSXIDList = VB.App.Path & "\dat\psx-usa-id.dat"
CoverName = Label11.Caption
If FSO.FileExists(PSXIDList) Then
    Close #12
    Open PSXIDList For Input As #12
    Do
        Line Input #12, tmp
        If Mid$(tmp, 1, Len(CoverName)) = CoverName And Mid$(tmp, Len(CoverName) + 1, 1) = " " Then
            PSXID = tmp
            'MsgBox "PSXID: " & PSXID
            tmparray = Split(PSXID, " ")
            GetPSXID = tmparray(1)
        End If
    Loop Until EOF(12)
    Close #12
End If

End Function
Function Validate_Rom(ROMFILE As String)
Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(ROMFILE) = True Then
    tmp = Form1.CalcMD5(Form1.ShortPath(ROMFILE))
    ROMMD5 = tmp
    Label6.Caption = ROMMD5
    Shell ("cmd.exe /c del " & Chr(34) & VB.App.Path & "\md5a.txt" & Chr(34)), vbHide
    tmp = Form1.FileNameCleanup()
    'Label35.Visible = True
    CoverName = tmp
    Label13.Caption = GetPSXID()
    a = Redump(ROMMD5)
End If
Validate_Rom = tmp
End Function
Function Validate_Bios()
BIOSFILE = Form2.Text1.Text
Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(BIOSPATH & "\" & BIOSFILE) = True Then
    tmp = Form1.CalcMD5(Form1.ShortPath(BIOSPATH & "\" & BIOSFILE))
End If
If FSO.FileExists(BIOSFILE) = True Then
    tmp = Form1.CalcMD5(Form1.ShortPath(BIOSFILE))
End If

If LCase(tmp) = "239665b1a3dade1b5a52c06338011044" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-J v1.0J BIOS SCPH-1000/DTL-H1000"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
ElseIf LCase(tmp) = "849515939161e62f6b866f6853006780" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-J v1.1J BIOS SCPH-3000/DTL-H1000H"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
ElseIf LCase(tmp) = "dc2b9bf8da62ec93e868cfd29f0d067d" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-U v2.0A BIOS DTL-H1001"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check11.Value = 1
Form1.Check11.Value = 1
ElseIf LCase(tmp) = "54847e693405ffeb0359c6287434cbef" Then
Label9.Caption = tmp
Label8.Caption = "PAL v2.0E BIOS DTL-H1002/SCPH-1002"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check13.Value = 1
Form1.Check13.Value = 1
ElseIf LCase(tmp) = "cba733ceeff5aef5c32254f1d617fa62" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-J v2.1J BIOS SCPH-3500"
ElseIf LCase(tmp) = "da27e8b6dab242d8f91a9b25d80c63b8" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-U v2.1A BIOS DTL-H1101"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check11.Value = 1
Form1.Check11.Value = 1
ElseIf LCase(tmp) = "417b34706319da7cf001e76e40136c23" Then
Label9.Caption = tmp
Label8.Caption = "PAL v2.1E BIOS SCPH-1002/DTL-H1102"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check13.Value = 1
Form1.Check13.Value = 1
ElseIf LCase(tmp) = "57a06303dfa9cf9351222dfcbb4a29d9" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-J v2.2J BIOS SCPH-5000/DTL-H1200/DTL-H3000"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
ElseIf LCase(tmp) = "924e392ed05558ffdb115408c263dccf" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-U v2.2A BIOS SCPH-1001/SCPH-5003/DTL-H1201/DTL-H3001"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check11.Value = 1
Form1.Check11.Value = 1
ElseIf LCase(tmp) = "e2110b8a2b97a8e0b857a45d32f7e187" Then
Label9.Caption = tmp
Label8.Caption = "PAL v2.2E BIOS SCPH-1002/DTL-H1202/DTL-H3002"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check13.Value = 1
Form1.Check13.Value = 1
ElseIf LCase(tmp) = "ca5cfc321f916756e3f0effbfaeba13b" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-J v2.2D BIOS DTL-H1100"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
ElseIf LCase(tmp) = "490f666e1afb15b7362b406ed1cea246" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-U v3.0A BIOS SCPH-5501/SCPH-5503/SCPH-7003"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check11.Value = 1
Form1.Check11.Value = 1
ElseIf LCase(tmp) = "8dd7d5296a650fac7319bce665a6a53c" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-J v3.0J BIOS SCPH-5500"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
ElseIf LCase(tmp) = "32736f17079d0b2b7024407c39bd3050" Then
Label9.Caption = tmp
Label8.Caption = "PAL v3.0E BIOS SCPH-5502/SCPH-5552"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check13.Value = 1
Form1.Check13.Value = 1
ElseIf LCase(tmp) = "8e4c14f567745eff2f0408c8129f72a6" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-J v4.0J BIOS SCPH-7000/SCPH-7500/SCPH-9000"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
ElseIf LCase(tmp) = "b84be139db3ee6cbd075630aa20a6553" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-U v4.1A BIOS SCPH-7000W"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check11.Value = 1
Form1.Check11.Value = 1
ElseIf LCase(tmp) = "1e68c231d0896b7eadcad1d7d8e76129" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-U v4.1A BIOS SCPH-7001/SCPH-7501/SCPH-7503/SCPH-9001/SCPH-9003"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check11.Value = 1
Form1.Check11.Value = 1
ElseIf LCase(tmp) = "b9d9a0286c33dc6b7237bb13cd46fdee" Then
Label9.Caption = tmp
Label8.Caption = "PAL v4.1E BIOS SCPH-7002/SCPH-7502/SCPH-9002"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check13.Value = 1
Form1.Check13.Value = 1
ElseIf LCase(tmp) = "8abc1b549a4a80954addc48ef02c4521" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-J v4.3J BIOS SCPH-100"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
ElseIf LCase(tmp) = "b10f5e0e3d9eb60e5159690680b1e774" Then
Label9.Caption = tmp
Label8.Caption = "PAL v4.4E BIOS SCPH-102"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check13.Value = 1
Form1.Check13.Value = 1
ElseIf LCase(tmp) = "6e3735ff4c7dc899ee98981385f6f3d0" Then
Label9.Caption = tmp
Label8.Caption = "NTSC-U v4.5A BIOS SCPH-101"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check11.Value = 1
Form1.Check11.Value = 1
ElseIf LCase(tmp) = "de93caec13d1a141a40a79f5c86168d6" Then
Label9.Caption = tmp
Label8.Caption = "PAL v4.5E BIOS SCPH-102"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check13.Value = 1
Form1.Check13.Value = 1
ElseIf LCase(tmp) = "3240872c70984b6cbfda1586cab68dbe" Then
Label9.Caption = tmp
Label8.Caption = "SEGA SATURN V1.01A US/EU"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Form1.Check11.Value = 1
Form1.Check11.Value = 1
ElseIf LCase(tmp) = "85ec9ca47d8f6807718151cbcca8b964" Then
Label9.Caption = tmp
Label8.Caption = "SEGA SATURN V1.01 JP"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
ElseIf LCase(tmp) = "af5828fdff51384f99b3c4926be27762" Then
Label9.Caption = tmp
Label8.Caption = "SEGA SATURN V1.00 JP"
Label15.Caption = "REDUMP: verified!"
Label15.ForeColor = RGB(0, 153, 0)
Else
    Label9.Caption = tmp
End If
Validate_Bios = tmp
End Function
Public Function Redump(REDUMPMD5)
REDUMPMD5 = LCase(REDUMPMD5)
Set FSO = CreateObject("Scripting.FileSystemObject")
If SysCore = "psx" Then
    RedumpList = VB.App.Path & "\dat\psx-usa-redump.dat"
ElseIf SysCore = "ss" Then
    RedumpList = VB.App.Path & "\dat\ss-usa-redump.dat"
ElseIf SysCore = "pce" Then
    RedumpList = VB.App.Path & "\dat\pce-usa-redump.dat"
ElseIf SysCore = "nes" Then
    RedumpList = VB.App.Path & "\dat\nes-all-goodtools.dat"
ElseIf SysCore = "snes" Then
    RedumpList = VB.App.Path & "\dat\snes-all-nointro.dat"
ElseIf SysCore = "snes_faust" Then
    RedumpList = VB.App.Path & "\dat\snes-all-nointro.dat"
ElseIf SysCore = "gg" Then
    RedumpList = VB.App.Path & "\dat\gg-all-nointro.dat"
ElseIf SysCore = "gba" Then
    RedumpList = VB.App.Path & "\dat\gba-all-nointro.dat"
ElseIf SysCore = "gbc" Then
    RedumpList = VB.App.Path & "\dat\gbc-all-nointro.dat"
ElseIf SysCore = "vb" Then
    RedumpList = VB.App.Path & "\dat\vb-all-nointro.dat"
ElseIf SysCore = "md" Then
    RedumpList = VB.App.Path & "\dat\md-all-nointro.dat"
End If

If FSO.FileExists(RedumpList) Then
    Close #13
    Open RedumpList For Input As #13
    booltmp = False
    Do
        Line Input #13, tmp
        For x = 1 To Len(tmp)
            If Mid$(tmp, x, 32) = UCase(REDUMPMD5) Or Mid$(tmp, x, 32) = LCase(REDUMPMD5) Then
                If SysCore = "psx" Or SysCore = "ss" Or SysCore = "pce" Then
                    Label23.Caption = "REDUMP: verified!"
                    Label23.ForeColor = RGB(0, 153, 0)
                    booltmp = True
                ElseIf SysCore = "nes" Then
                    Label23.Caption = "GOODTOOLS: verified!"
                    Label23.ForeColor = RGB(0, 153, 0)
                    booltmp = True
                Else
                    Label23.Caption = "NOINTRO: verified!"
                    Label23.ForeColor = RGB(0, 153, 0)
                    booltmp = True
                End If
            End If
        Next x
    Loop Until EOF(13)
    Close #13
    If booltmp = False Then
                If SysCore = "psx" Or SysCore = "ss" Or SysCore = "pce" Then
                    Label23.Caption = "REDUMP: unverified!"
                    Label23.ForeColor = RGB(255, 128, 0)
                    booltmp = True
                ElseIf SysCore = "nes" Then
                    Label23.Caption = "GOODTOOLS: unverified!"
                    Label23.ForeColor = RGB(255, 128, 0)
                    booltmp = True
                Else
                    Label23.Caption = "NOINTRO: unverified!"
                    Label23.ForeColor = RGB(255, 128, 0)
                    booltmp = True
                End If
    End If
End If
End Function
Private Sub Documentation_Click()
    Shell "cmd.exe /c start https://mednafen.github.io/documentation/", vbHide
End Sub
Private Sub Form_Load()
Build = Form1.GetBuild()
Form4.Caption = "MedAdvCFG v" & Build & " (Mednafen v0.9.x.x Frontend) by Nigel Todman [BASIC MODE]"
Form4.Width = 14355
Set FSO = CreateObject("Scripting.FileSystemObject")
SysCore = Form1.SetSysCore
a = LoadSettings()
SysCore = Form1.SetSysCore
Basic.Checked = True
tmparray = Split(Form3.Text1.Text, vbCrLf)
'SysCore
Label1.Caption = tmparray(0)
'FileName
Label2.Caption = tmparray(2)
'CoverSearched
Label11.Caption = tmparray(3)

Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False

a = Form4.Validate_Rom(Image1(1).Tag)
If SysCore = "psx" Or SysCore = "ss" Or SysCore = "pce" Then
    a = Form4.Validate_Bios
    Label7.Visible = True
    Label8.Visible = True
End If

If SysCore = "snes" Then
    Image1(1).Height = 5895
ElseIf SysCore = "nes" Then
    Image1(1).Width = 5895
ElseIf SysCore = "gg" Then
    Image1(1).Width = 5895
ElseIf SysCore = "md" Then
    Image1(1).Width = 5895
End If
End Sub
Function LoadSettings()
'Load Settings
Set FSO = CreateObject("Scripting.FileSystemObject")

If FSO.FileExists(VB.App.Path & "\MedAdvCFG.dat") Then
    Open VB.App.Path & "\MedAdvCFG.dat" For Input As #1
    For x = 1 To 34
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
    BasicModeFolder = Mid$(tmp3(34), 17, Len(tmp3(34)))
End If
'End Load Settings
End Function

Private Sub Label13_Click()
Clipboard.Clear
Clipboard.SetText Label13.Caption
MsgBox "Game ID copied to Clipboard"
End Sub

Private Sub Label6_Click()
Clipboard.Clear
Clipboard.SetText Label6.Caption
MsgBox "MD5 copied to Clipboard"
End Sub

Private Sub Quit_Click()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
End Sub

