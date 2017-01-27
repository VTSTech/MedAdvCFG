VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "MedAdvCFG v0.0.0 (Mednafen v0.9.x.x Frontend) by Nigel Todman [BASIC MODE]"
   ClientHeight    =   7110
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10905
   LinkTopic       =   "Form2"
   ScaleHeight     =   15979.9
   ScaleMode       =   0  'User
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00000000&
      Caption         =   "Back"
      Height          =   315
      Left            =   10200
      TabIndex        =   21
      Top             =   120
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
      Height          =   3015
      Left            =   11160
      TabIndex        =   16
      Top             =   450
      Width           =   3615
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0C0C0&
      Height          =   3210
      Left            =   11160
      TabIndex        =   15
      Top             =   3480
      Width           =   3615
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   18
      Top             =   2280
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "System BIOS:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2040
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mednafen EXE:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1920
      TabIndex        =   13
      Top             =   2040
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0.9.41.0-win64 Detected! MD5: 6AADC9A8A196DA610E6DB43367B339B4"
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
      Top             =   2040
      Width           =   6345
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
      Left            =   3480
      TabIndex        =   11
      Top             =   1680
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
      Caption         =   "ROM Folder:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2070
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
      Picture         =   "Form2.frx":0007
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Image Image14 
      Height          =   1335
      Left            =   8880
      Picture         =   "Form2.frx":0BD9
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Image Image13 
      Height          =   1335
      Left            =   6810
      Picture         =   "Form2.frx":4586
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Image Image12 
      Height          =   1335
      Left            =   600
      Picture         =   "Form2.frx":7C17
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image11 
      Height          =   1335
      Left            =   4740
      Picture         =   "Form2.frx":DF3E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image Image10 
      Height          =   1335
      Left            =   2640
      Picture         =   "Form2.frx":16313
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Image Image9 
      Height          =   1335
      Left            =   600
      Picture         =   "Form2.frx":17A30
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image8 
      Height          =   1335
      Left            =   6810
      Picture         =   "Form2.frx":20B77
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image7 
      Height          =   1335
      Left            =   2670
      Picture         =   "Form2.frx":430A4
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   1335
      Left            =   8880
      Picture         =   "Form2.frx":46FC1
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Image Image5 
      Height          =   1335
      Left            =   6810
      Picture         =   "Form2.frx":4E3AD
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   1215
      Left            =   8880
      Picture         =   "Form2.frx":516D2
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   1335
      Left            =   4740
      Picture         =   "Form2.frx":63401
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   2670
      Picture         =   "Form2.frx":69672
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   600
      Picture         =   "Form2.frx":6AE2E
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1695
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
Dim MedAdvGAMES, MedAdvCOVERS, MedAdvEXT, BasicModeFolder, LogoTop, LogoLeft
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
        Shell ("cmd.exe /c " & Chr(34) & VB.App.Path & "\md5.exe -n " & Chr(34) & BIOSPATH & "\" & BIOSFILE & Chr(34) & " >> " & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
        Sleep (200)
        If FSO.FileExists(VB.App.Path & "\md5.txt") = True Then
            Open VB.App.Path & "\md5.txt" For Input As #5
                Line Input #5, tmp
            Close #5
        End If
        Shell ("cmd.exe /c del " & Chr(34) & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
        Text1.Text = BIOSPATH & "\" & BIOSFILE
    End If
    If FSO.FileExists(BIOSFILE) = True Then
        Shell ("cmd.exe /c " & Chr(34) & VB.App.Path & "\md5.exe -n " & Chr(34) & BIOSFILE & Chr(34) & " >> " & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
        Sleep (200)
        If FSO.FileExists(VB.App.Path & "\md5.txt") = True Then
            Open VB.App.Path & "\md5.txt" For Input As #5
                Line Input #5, tmp
            Close #5
        End If
        Shell ("cmd.exe /c del " & Chr(34) & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
        Text1.Text = BIOSFILE
    End If

        If LCase(tmp) = "239665b1a3dade1b5a52c06338011044" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-J v1.0J BIOS SCPH-1000/DTL-H1000"
            'Check12.Value = 1
            'Check12.Value = 1
        ElseIf LCase(tmp) = "849515939161e62f6b866f6853006780" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-J v1.1J BIOS SCPH-3000/DTL-H1000H"
            'Check12.Value = 1
            'Check12.Value = 1
        ElseIf LCase(tmp) = "dc2b9bf8da62ec93e868cfd29f0d067d" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-U v2.0A BIOS DTL-H1001"
            'Check11.Value = 1
            'Check11.Value = 1
        ElseIf LCase(tmp) = "54847e693405ffeb0359c6287434cbef" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "PAL v2.0E BIOS DTL-H1002/SCPH-1002"
            'Check13.Value = 1
            'Check13.Value = 1
        ElseIf LCase(tmp) = "cba733ceeff5aef5c32254f1d617fa62" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-J v2.1J BIOS SCPH-3500"
            'Check12.Value = 1
            'Check12.Value = 1
        ElseIf LCase(tmp) = "da27e8b6dab242d8f91a9b25d80c63b8" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-U v2.1A BIOS DTL-H1101"
            'Check11.Value = 1
            'Check11.Value = 1
        ElseIf LCase(tmp) = "417b34706319da7cf001e76e40136c23" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "PAL v2.1E BIOS SCPH-1002/DTL-H1102"
            'Check13.Value = 1
            'Check13.Value = 1
        ElseIf LCase(tmp) = "57a06303dfa9cf9351222dfcbb4a29d9" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-J v2.2J BIOS SCPH-5000/DTL-H1200/DTL-H3000"
            'Check12.Value = 1
            'Check12.Value = 1
        ElseIf LCase(tmp) = "924e392ed05558ffdb115408c263dccf" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-U v2.2A BIOS SCPH-1001/SCPH-5003/DTL-H1201/DTL-H3001"
            'Check11.Value = 1
            'Check11.Value = 1
        ElseIf LCase(tmp) = "e2110b8a2b97a8e0b857a45d32f7e187" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label29.Caption = "PAL v2.2E BIOS SCPH-1002/DTL-H1202/DTL-H3002"
            'Check13.Value = 1
            'Check13.Value = 1
        ElseIf LCase(tmp) = "ca5cfc321f916756e3f0effbfaeba13b" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-J v2.2D BIOS DTL-H1100"
            'Check12.Value = 1
            'Check12.Value = 1
        ElseIf LCase(tmp) = "490f666e1afb15b7362b406ed1cea246" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-U v3.0A BIOS SCPH-5501/SCPH-5503/SCPH-7003"
            'Check11.Value = 1
            'Check11.Value = 1
        ElseIf LCase(tmp) = "8dd7d5296a650fac7319bce665a6a53c" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-J v3.0J BIOS SCPH-5500"
            'Check12.Value = 1
            'Check12.Value = 1
        ElseIf LCase(tmp) = "32736f17079d0b2b7024407c39bd3050" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "PAL v3.0E BIOS SCPH-5502/SCPH-5552"
            'Check13.Value = 1
            'Check13.Value = 1
        ElseIf LCase(tmp) = "8e4c14f567745eff2f0408c8129f72a6" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-J v4.0J BIOS SCPH-7000/SCPH-7500/SCPH-9000"
            'Check12.Value = 1
            'Check12.Value = 1
        ElseIf LCase(tmp) = "b84be139db3ee6cbd075630aa20a6553" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-U v4.1A BIOS SCPH-7000W"
            'Check11.Value = 1
            'Check11.Value = 1
        ElseIf LCase(tmp) = "1e68c231d0896b7eadcad1d7d8e76129" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-U v4.1A BIOS SCPH-7001/SCPH-7501/SCPH-7503/SCPH-9001/SCPH-9003"
            'Check11.Value = 1
            'Check11.Value = 1
        ElseIf LCase(tmp) = "b9d9a0286c33dc6b7237bb13cd46fdee" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "PAL v4.1E BIOS SCPH-7002/SCPH-7502/SCPH-9002"
            'Check13.Value = 1
            'Check13.Value = 1
        ElseIf LCase(tmp) = "8abc1b549a4a80954addc48ef02c4521" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-J v4.3J BIOS SCPH-100"
            'Check12.Value = 1
            'Check12.Value = 1
        ElseIf LCase(tmp) = "b10f5e0e3d9eb60e5159690680b1e774" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "PAL v4.4E BIOS SCPH-102"
            'Check13.Value = 1
            'Check13.Value = 1
        ElseIf LCase(tmp) = "6e3735ff4c7dc899ee98981385f6f3d0" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "NTSC-U v4.5A BIOS SCPH-101"
            'Check11.Value = 1
            'Check11.Value = 1
        ElseIf LCase(tmp) = "de93caec13d1a141a40a79f5c86168d6" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "PAL v4.5E BIOS SCPH-102"
            'Check13.Value = 1
            'Check13.Value = 1
        ElseIf LCase(tmp) = "3240872c70984b6cbfda1586cab68dbe" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "SEGA SATURN V1.01A US/EU"
            'Check11.Value = 1
            'Check11.Value = 1
        ElseIf LCase(tmp) = "85ec9ca47d8f6807718151cbcca8b964" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "SEGA SATURN V1.01 JP"
            'Check12.Value = 1
            'Check12.Value = 1
        ElseIf LCase(tmp) = "af5828fdff51384f99b3c4926be27762" Then
            'Label6.Caption = "MD5: " & tmp
            'Label29.Visible = True
            Label6.Caption = "SEGA SATURN V1.00 JP"
            'Check12.Value = 1
            'Check12.Value = 1
        Else
            'Label29.Visible = False
            'Label6.Caption = "MD5: " & tmp
        End If
        Validate_Bios = tmp
End Function
Function Validate_MedEXE()
If FSO.FileExists(MedEXE) = True Then
    Shell ("cmd.exe /c " & Chr(34) & Chr(34) & VB.App.Path & "\md5.exe" & Chr(34) & " -n " & Chr(34) & MedEXE & Chr(34) & " >> " & Chr(34) & VB.App.Path & "\md5.txt" & Chr(34) & Chr(34)), vbHide
    'MsgBox ("cmd.exe /c " & Chr(34) & Chr(34) & VB.App.Path & "\md5.exe" & Chr(34) & " -n " & Chr(34) & MedEXE & Chr(34) & " >> " & Chr(34) & VB.App.Path & "\md5.txt" & Chr(34) & Chr(34)), vbHide
    Sleep (800)
    If FSO.FileExists(VB.App.Path & "\md5.txt") = True Then
        Open VB.App.Path & "\md5.txt" For Input As #4
            Line Input #4, tmp
        Close #4
    End If
    If tmp = "8F0BC836E2B6023371B99E94829B5CF1" Then
        Label2.Caption = "0.9.38.7-win64 Detected! MD5: 8F0BC836E2B6023371B99E94829B5CF1"
    ElseIf tmp = "C2CA5F8A9A4CF93BB05297272F029B9C" Then
        Label2.Caption = "0.9.38.7-win32 Detected! MD5: C2CA5F8A9A4CF93BB05297272F029B9C"
    ElseIf tmp = "D89B755B1616323B7181C9D1931D4E39" Then
        Label2.Caption = "0.9.38.6-win64 Detected! MD5: D89B755B1616323B7181C9D1931D4E39"
    ElseIf tmp = "D6A8592FB42104327EF7E57D4F0C8ED1" Then
        Label2.Caption = "0.9.38.6-win32 Detected! MD5: D6A8592FB42104327EF7E57D4F0C8ED1"
    ElseIf tmp = "E7A5FBC376B2DAA55AB4A3FF9C6AF1E1" Then
        Label2.Caption = "0.9.38.5-win64 Detected! MD5: E7A5FBC376B2DAA55AB4A3FF9C6AF1E1"
    ElseIf tmp = "74B1D63CBAB0CC4F91A9F3FB5020AB78" Then
        Label2.Caption = "0.9.38.5-win32 Detected! MD5: 74B1D63CBAB0CC4F91A9F3FB5020AB78"
    ElseIf tmp = "A30FC82730A62781BBF39DF0652456D2" Then
        Label2.Caption = "0.9.37.1-win64 Detected! MD5: A30FC82730A62781BBF39DF0652456D2"
    ElseIf tmp = "D02DE97F10E0FE544427B8F0BBEED3F1" Then
        Label2.Caption = "0.9.37.1-win32 Detected! MD5: D02DE97F10E0FE544427B8F0BBEED3F1"
    ElseIf tmp = "9357B96CB347AA52E5F2796AB9A062BD" Then
        Label2.Caption = "0.9.39-unstable-win64 Detected! MD5: 9357B96CB347AA52E5F2796AB9A062BD"
    ElseIf tmp = "84438BB30BB5F15488AEAE662D91E56C" Then
        Label2.Caption = "0.9.39-unstable-win32 Detected! MD5: 84438BB30BB5F15488AEAE662D91E56C"
    ElseIf tmp = "942641E6C569B6AC26A66BA9051FCD1C" Then
        Label2.Caption = "0.9.39.1-win64 Detected! MD5: 942641E6C569B6AC26A66BA9051FCD1C"
    ElseIf tmp = "213AE98231F8F96D26B0C409085DAB73" Then
        Label2.Caption = "0.9.39.1-win32 Detected! MD5: 213AE98231F8F96D26B0C409085DAB73"
    ElseIf tmp = "A5AFC70E3B2B267A0E81F64B7366C8C3" Then
        Label2.Caption = "0.9.39.2-win64 Detected! MD5: A5AFC70E3B2B267A0E81F64B7366C8C3"
    ElseIf tmp = "AEF947A6E5A35FF108B954683CD3698A" Then
        Label2.Caption = "0.9.39.2-win32 Detected! MD5: AEF947A6E5A35FF108B954683CD3698A"
    ElseIf tmp = "6AADC9A8A196DA610E6DB43367B339B4" Then
        Label2.Caption = "0.9.41.0-win64 Detected! MD5: 6AADC9A8A196DA610E6DB43367B339B4"
    ElseIf tmp = "74EA6CD12BF60ADF1A93A40854CD4686" Then
        Label2.Caption = "0.9.41.0-win32 Detected! MD5: 74EA6CD12BF60ADF1A93A40854CD4686"
'Unknown Mednafen Version! MD5: A5AFC70E3B2B267A0E81F64B7366C8C3 x64
'Unknown Mednafen Version! MD5: AEF947A6E5A35FF108B954683CD3698A x86
'Unknown Mednafen Version! MD5: 6AADC9A8A196DA610E6DB43367B339B4 x64
'Unknown Mednafen Version! MD5: DE8348EFB4EA79C58711990031FD7505 x86
'Unknown Mednafen Version! MD5: 74EA6CD12BF60ADF1A93A40854CD4686 x86
    Else
        Label2.Caption = "Unknown Mednafen Version! MD5: " & tmp
    End If
    Shell ("cmd.exe /c del " & Chr(34) & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
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
Dim vFile As Variant
Close #7
Open MedAdvGAMES For Output As #7
For Each vFile In colFiles
    Text3.Text = Text3.Text & vFile & vbCrLf
    Print #7, vFile
Next vFile
Close #7

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
            tmp2 = MsgBox("Set File: " & File1.Filename, vbYesNo, "Set this file?")
        End If
    If tmp2 = vbYes Then
        MedEXE = Dir1.Path & "\" & File1.Filename
        Form1.Width = 9240
        ActiveFile = "None"
        tmp2 = ""
        a = Validate_MedEXE()
    End If
End If

If ActiveFile = "BIOS" Then
            tmp2 = MsgBox("Set File: " & File1.Filename, vbYesNo, "Set this file?")
    If tmp2 = vbYes Then
        BIOSPATH = Dir1.Path
        BIOSFILE = File1.Filename
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
            tmp2 = MsgBox("Set File: " & File1.Filename, vbYesNo, "Set this file?")
        End If
    If tmp2 = vbYes Then
        Text2.Text = Dir1.Path & "\" & File1.Filename
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
            tmp2 = MsgBox("Set Disc 1: " & File1.Filename, vbYesNo, "Set this file?")
        End If
            If tmp2 = vbYes Then
                Print #2, Dir1.Path & "\" & File1.Filename
                tmp1 = File1.Filename
            End If
            tmp2 = ""
            z = z + 1
            MsgBox ("Now select the next disc")
        ElseIf z >= 1 And z <= Val(M3USize) Then
            Do
                If File1.Filename <> tmp1 And Len(File1.Filename) > 1 And z <= Val(M3USize) And File1.Filename <> LastFile Then
                    If Check15.Value = 1 Then
                        a = a
                        tmp2 = vbYes
                    Else
                        tmp2 = MsgBox("Set Disc " & z + 1 & ": " & File1.Filename, vbYesNo, "Set this file?")
                    End If
                    LastFile = File1.Filename
                    If tmp2 = vbYes Then
                        Print #2, Dir1.Path & "\" & File1.Filename
                        z = z + 1
                        If z <> Val(M3USize) Then
                            MsgBox ("Now select the next disc")
                        End If
                    ElseIf tmp2 = vbNo Then
                        File1.Filename = ""
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
Build = "0.2.7"
Form2.Caption = "MedAdvCFG v" & Build & " (Mednafen v0.9.x.x Frontend) by Nigel Todman [BASIC MODE]"
Form2.Width = 11145
Set FSO = CreateObject("Scripting.FileSystemObject")
a = LoadSettings()
Text1.Text = BIOSFILE
Text2.Text = BasicModeFolder
'Text4.Text = ScaleFactor
Text5.Text = XRes
Text6.Text = YRes
'Text7.Text = SavePath
'Text8.Text = numplayers

Dir1.Path = LastPath
File1.Path = LastPath

ROMDIR = BasicModeFolder
BIOSPATH = BiosPathLoad
a = Validate_MedEXE()
a = Validate_Bios()

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

'**
'Form1.Combo1.Text "gb (GameBoy (Color))", 0
'Form1.Combo1.Text "gba (GameBoy Advanced)", 1
'Form1.Combo1.Text "gg (Sega Game Gear)", 2
'Form1.Combo1.Text "lynx (Atari Lynx)", 3
'Form1.Combo1.Text "md (Sega Genesis/MegaDrive)", 4
'Form1.Combo1.Text "nes (Nintendo Entertainment System)", 5
'Form1.Combo1.Text "ngp (Neo Geo Pocket (Color))", 6
'Form1.Combo1.Text "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)", 7
'Form1.Combo1.Text "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)", 8
'Form1.Combo1.Text "pcfx (PC-FX)", 9
'Form1.Combo1.Text "psx (Sony PlayStation)", 10
'Form1.Combo1.Text "sms (Sega Master System)", 11
'Form1.Combo1.Text "snes (Super Nintendo Entertainment System)", 12
'Form1.Combo1.Text "ss (Sega Saturn)", 13
'Form1.Combo1.Text "vb (Virtual Boy)", 14
'Form1.Combo1.Text "wswan (WonderSwan)", 15
'**
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
        Shell ("cmd.exe /c " & Chr(34) & VB.App.Path & "\md5.exe -n " & Chr(34) & ROMFILE & Chr(34) & " >> " & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
        Sleep (500)
        If FSO.FileExists(VB.App.Path & "\md5.txt") = True Then
            Open VB.App.Path & "\md5.txt" For Input As #3
                Line Input #3, tmp
            Close #3
        End If
        Shell ("cmd.exe /c del " & Chr(34) & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
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
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Text1.Visible = True
Text2.Visible = True
Label4.Visible = True
Label7.Visible = True
Label26.Visible = True
Check1.Visible = True
Check23.Visible = True
Text5.Visible = True
Text6.Visible = True
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

Private Sub Save_Click()
Set FSO = CreateObject("Scripting.FileSystemObject")
Close #6
Open VB.App.Path & "\MedAdvCFG.dat" For Output As #6
    Print #6, "MedEXE=" & MedEXE
    Print #6, "SystemCore=" & Form1.Combo1.Text
    If FSO.FileExists(BIOSPATH & "\" & BIOSFILE) Then
        Print #6, "SystemBIOS=" & BIOSPATH & "\" & BIOSFILE
    ElseIf FSO.FileExists(BIOSFILE) Then
        Print #6, "SystemBIOS=" & BIOSFILE
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
Close #6
End Sub

