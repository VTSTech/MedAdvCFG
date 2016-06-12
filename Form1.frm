VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "MedAdvCFG v0.0.0 (Mednafen v0.9.38.x Frontend) by Nigel Todman"
   ClientHeight    =   6135
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12555
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   12555
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PAL"
      Height          =   195
      Left            =   4080
      TabIndex        =   58
      Top             =   600
      Width           =   735
   End
   Begin VB.CheckBox Check12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NTSC-J"
      Height          =   195
      Left            =   2640
      TabIndex        =   57
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NTSC-U"
      Height          =   195
      Left            =   1320
      TabIndex        =   56
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1320
      TabIndex        =   54
      Text            =   "Not Set"
      Top             =   3720
      Width           =   4575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set"
      Height          =   315
      Left            =   6000
      TabIndex        =   53
      Top             =   3720
      Width           =   615
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BIOS MD5 Check"
      Height          =   195
      Left            =   7080
      TabIndex        =   52
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM MD5 Check"
      Height          =   195
      Left            =   7080
      TabIndex        =   51
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   7920
      TabIndex        =   49
      Text            =   "1080"
      Top             =   3240
      Width           =   515
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   7440
      TabIndex        =   48
      Text            =   "1920"
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   8160
      TabIndex        =   46
      Text            =   "2"
      Top             =   3000
      Width           =   285
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear BIOS/ROM"
      Height          =   315
      Left            =   6000
      TabIndex        =   45
      Top             =   240
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   9000
      TabIndex        =   44
      Top             =   0
      Width           =   3615
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0C0C0&
      Height          =   2820
      Left            =   9000
      TabIndex        =   41
      Top             =   2400
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0C0C0&
      Height          =   2115
      Left            =   9000
      TabIndex        =   40
      Top             =   360
      Width           =   3615
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Review Command Line before execution?"
      Height          =   435
      Left            =   4440
      TabIndex        =   39
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frameskip"
      Height          =   195
      Left            =   5160
      TabIndex        =   38
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fullscreen (FS)"
      Height          =   195
      Left            =   5160
      TabIndex        =   30
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Launch mednafen.exe"
      Height          =   495
      Left            =   4440
      TabIndex        =   29
      Top             =   4680
      Width           =   1815
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1320
      TabIndex        =   25
      Text            =   "None - None/Disabled"
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bilinear interpolation"
      Height          =   195
      Left            =   5160
      TabIndex        =   24
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   8160
      TabIndex        =   22
      Text            =   "50"
      Top             =   2760
      Width           =   285
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Accumulate color"
      Height          =   195
      Left            =   7080
      TabIndex        =   21
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temporal Blur"
      Height          =   195
      Left            =   7080
      TabIndex        =   20
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1320
      TabIndex        =   18
      Text            =   "None - None/Disabled"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1320
      TabIndex        =   16
      Text            =   "0 - Disabled"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM Sanity Check"
      Height          =   195
      Left            =   7080
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set"
      Height          =   315
      Left            =   6000
      TabIndex        =   12
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Text            =   "Not Set"
      Top             =   1680
      Width           =   4575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BIOS Sanity Check"
      Height          =   195
      Left            =   7080
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set"
      Height          =   315
      Left            =   6000
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "Not Set"
      Top             =   840
      Width           =   4575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Text            =   "gb (GameBoy (Color))"
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Height          =   195
      Left            =   1320
      TabIndex        =   60
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "System Region:"
      Height          =   195
      Left            =   120
      TabIndex        =   59
      Top             =   600
      Width           =   1110
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save Path:"
      Height          =   195
      Left            =   120
      TabIndex        =   55
      Top             =   3720
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   6960
      Picture         =   "Form1.frx":851A
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1785
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Resolution Override (FS)"
      Height          =   195
      Left            =   5160
      TabIndex        =   50
      Top             =   3300
      Width           =   1725
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Global Scaling Factor (FS/W)"
      Height          =   195
      Left            =   5160
      TabIndex        =   47
      Top             =   3050
      Width           =   2085
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click filename to Set and collapse panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9000
      TabIndex        =   43
      Top             =   5640
      Width           =   3435
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Folder to change File List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9360
      TabIndex        =   42
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "F5:  Save state"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "F7:  Load state"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "F10: Reset Console"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   4920
      Width           =   1680
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Alt + Enter: Toggle fullscreen mode."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   5160
      Width           =   3075
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "CTRL + SHIFT + 1: Select Controller Type (Where 1 is which Port/Player)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   33
      Top             =   5400
      Width           =   6270
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALT + SHIFT + 1: Configure buttons for Controller Port (Where 1 is which Port/Player)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   5640
      Width           =   7305
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hotkeys:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   4200
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "www.NigelTodman.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   1890
      TabIndex        =   28
      Top             =   4560
      Width           =   1995
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Written by Nigel Todman"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   1830
      TabIndex        =   27
      Top             =   4320
      Width           =   2115
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Video Scaler"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Blur amount:"
      Height          =   195
      Left            =   7080
      TabIndex        =   23
      Top             =   2760
      Width           =   885
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pixel Shader"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   885
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "FS Stretch"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   750
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Not Set"
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Top             =   2160
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM MD5:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM Image:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Not Set"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "BIOS MD5:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "System BIOS:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "System Core:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   45
      Width           =   6225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mednafen EXE:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   1125
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Save_Settings 
         Caption         =   "Save Settings"
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by Nigel Todman (nigel@nigeltodman.com)
'Website: www.NigelTodman.com
'GitHub: https://github.com/Veritas83/MedAdvCFG

Dim MedEXE, FSO, tmp, tmp2, tmp3(99), BIOSFILE, BIOSPATH, ROMFILE, SystemCore, SYSCORE, BIOSSanity, ROMSanity, Stretch, PixelShader, VideoScaler, x, y, z
Dim cmdstring, Build, Frameskip, Fullscreen, TBlur, TblurAccum, AccumAmount, VideoIP, ActiveFile, XRes, YRes, ScaleFactor, LastPath, SavePath, BiosPathLoad
Dim ResetBios, ResetRom, ResetSave, FatalError, SystemRegion, SystemRegionLoad

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Function Validate_Rom()
If Check9.Value = 1 Then
    If FSO.FileExists(ROMFILE) = True Then
        Shell ("cmd.exe /c " & Chr(34) & VB.App.Path & "\md5.exe -n " & Chr(34) & ROMFILE & Chr(34) & " >> " & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
        Sleep (500)
        If FSO.FileExists(VB.App.Path & "\md5.txt") = True Then
            Open VB.App.Path & "\md5.txt" For Input As #1
                Line Input #1, tmp
            Close #1
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
Private Function Validate_Bios()
If Check10.Value = 1 Then
    If FSO.FileExists(BIOSPATH & "\" & BIOSFILE) = True Then
        Shell ("cmd.exe /c " & Chr(34) & VB.App.Path & "\md5.exe -n " & Chr(34) & BIOSPATH & "\" & BIOSFILE & Chr(34) & " >> " & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
        Sleep (200)
        If FSO.FileExists(VB.App.Path & "\md5.txt") = True Then
            Open VB.App.Path & "\md5.txt" For Input As #1
                Line Input #1, tmp
            Close #1
        End If
        Shell ("cmd.exe /c del " & Chr(34) & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
        Text1.Text = BIOSFILE
        If LCase(tmp) = "239665b1a3dade1b5a52c06338011044" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-J v1.0J BIOS SCPH-1000/DTL-H1000"
            Check12.Value = 1
            Check12.Value = 1
        ElseIf LCase(tmp) = "849515939161e62f6b866f6853006780" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-J v1.1J BIOS SCPH-3000/DTL-H1000H"
            Check12.Value = 1
            Check12.Value = 1
        ElseIf LCase(tmp) = "dc2b9bf8da62ec93e868cfd29f0d067d" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-U v2.0A BIOS DTL-H1001"
            Check11.Value = 1
            Check11.Value = 1
        ElseIf LCase(tmp) = "54847e693405ffeb0359c6287434cbef" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: PAL v2.0E BIOS DTL-H1002/SCPH-1002"
            Check13.Value = 1
            Check13.Value = 1
        ElseIf LCase(tmp) = "cba733ceeff5aef5c32254f1d617fa62" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-J v2.1J BIOS SCPH-3500"
            Check12.Value = 1
            Check12.Value = 1
        ElseIf LCase(tmp) = "da27e8b6dab242d8f91a9b25d80c63b8" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-U v2.1A BIOS DTL-H1101"
            Check11.Value = 1
            Check11.Value = 1
        ElseIf LCase(tmp) = "417b34706319da7cf001e76e40136c23" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: PAL v2.1E BIOS SCPH-1002/DTL-H1102"
            Check13.Value = 1
            Check13.Value = 1
        ElseIf LCase(tmp) = "57a06303dfa9cf9351222dfcbb4a29d9" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-J v2.2J BIOS SCPH-5000/DTL-H1200/DTL-H3000"
            Check12.Value = 1
            Check12.Value = 1
        ElseIf LCase(tmp) = "924e392ed05558ffdb115408c263dccf" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-U v2.2A BIOS SCPH-1001/SCPH-5003/DTL-H1201/DTL-H3001"
            Check11.Value = 1
            Check11.Value = 1
        ElseIf LCase(tmp) = "e2110b8a2b97a8e0b857a45d32f7e187" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: PAL v2.2E BIOS SCPH-1002/DTL-H1202/DTL-H3002"
            Check13.Value = 1
            Check13.Value = 1
        ElseIf LCase(tmp) = "ca5cfc321f916756e3f0effbfaeba13b" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-J v2.2D BIOS DTL-H1100"
            Check12.Value = 1
            Check12.Value = 1
        ElseIf LCase(tmp) = "490f666e1afb15b7362b406ed1cea246" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-U v3.0A BIOS SCPH-5501/SCPH-5503/SCPH-7003"
            Check11.Value = 1
            Check11.Value = 1
        ElseIf LCase(tmp) = "8dd7d5296a650fac7319bce665a6a53c" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-J v3.0J BIOS SCPH-5500"
            Check12.Value = 1
            Check12.Value = 1
        ElseIf LCase(tmp) = "32736f17079d0b2b7024407c39bd3050" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: PAL v3.0E BIOS SCPH-5502/SCPH-5552"
            Check13.Value = 1
            Check13.Value = 1
        ElseIf LCase(tmp) = "8e4c14f567745eff2f0408c8129f72a6" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-J v4.0J BIOS SCPH-7000/SCPH-7500/SCPH-9000"
            Check12.Value = 1
            Check12.Value = 1
        ElseIf LCase(tmp) = "b84be139db3ee6cbd075630aa20a6553" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-U v4.1A BIOS SCPH-7000W"
            Check11.Value = 1
            Check11.Value = 1
        ElseIf LCase(tmp) = "1e68c231d0896b7eadcad1d7d8e76129" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-U v4.1A BIOS SCPH-7001/SCPH-7501/SCPH-7503/SCPH-9001/SCPH-9003"
            Check11.Value = 1
            Check11.Value = 1
        ElseIf LCase(tmp) = "b9d9a0286c33dc6b7237bb13cd46fdee" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: PAL v4.1E BIOS SCPH-7002/SCPH-7502/SCPH-9002"
            Check13.Value = 1
            Check13.Value = 1
        ElseIf LCase(tmp) = "8abc1b549a4a80954addc48ef02c4521" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-J v4.3J BIOS SCPH-100"
            Check12.Value = 1
            Check12.Value = 1
        ElseIf LCase(tmp) = "b10f5e0e3d9eb60e5159690680b1e774" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: PAL v4.4E BIOS SCPH-102"
            Check13.Value = 1
            Check13.Value = 1
        ElseIf LCase(tmp) = "6e3735ff4c7dc899ee98981385f6f3d0" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: NTSC-U v4.5A BIOS SCPH-101"
            Check11.Value = 1
            Check11.Value = 1
        ElseIf LCase(tmp) = "de93caec13d1a141a40a79f5c86168d6" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: PAL v4.5E BIOS SCPH-102"
            Check13.Value = 1
            Check13.Value = 1
        Else
            Label29.Visible = False
            Label6.Caption = "MD5: " & tmp
        End If
    End If
Else
    Label6.Caption = "MD5: BIOS MD5 Disabled"
End If
Validate_Bios = tmp
End Function
Function Validate_MedEXE()
If FSO.FileExists(MedEXE) = True Then
    Shell ("cmd.exe /c " & Chr(34) & VB.App.Path & "\md5.exe -n " & MedEXE & " >> " & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
    Sleep (200)
    If FSO.FileExists(VB.App.Path & "\md5.txt") = True Then
        Open VB.App.Path & "\md5.txt" For Input As #1
            Line Input #1, tmp
        Close #1
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
    Else
        Label2.Caption = "Unknown Mednafen Version! MD5: " & tmp
    End If
    Shell ("cmd.exe /c del " & Chr(34) & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
End If
Validate_MedEXE = tmp
End Function

Private Sub About_Click()
MsgBox "MedAdvCFG v" & Build & " (Mednafen v0.9.38.x Frontend)" & vbCrLf & "Written by Nigel Todman (nigel@nigeltodman.com)" & vbCrLf & "Primarily written as a PSX Frontend." & vbCrLf & "Tested with the following System Cores:" & vbCrLf & "GB, GBA, GG, MD, NES, PCE, PCE_FAST, PSX, SNES, VB" & vbCrLf & vbCrLf & "Homepage: www.NigelTodman.com" & vbCrLf & "Facebook: facebook.com/nigel.todman.3" & vbCrLf & "Twitter: @Veritas_83" & vbCrLf & "YouTube: Veritas0923"
End Sub

Private Sub Check11_Click()
Check12.Value = 0
Check13.Value = 0
End Sub

Private Sub Check12_Click()
Check11.Value = 0
Check13.Value = 0
End Sub

Private Sub Check13_Click()
Check11.Value = 0
Check12.Value = 0
End Sub

Private Sub Combo1_Change()
'MsgBox Combo1.Text
If Combo1.Text = "psx (Sony PlayStation)" Then
    Check1.Enabled = True
    Check2.Enabled = True
    Check11.Enabled = True
    Check12.Enabled = True
    Check13.Enabled = True
Else
    Check1.Value = 0
    Check2.Value = 0
    Check9.Value = 0
    Check10.Value = 0
    Check11.Value = 0
    Check12.Value = 0
    Check13.Value = 0
    Check1.Enabled = False
    Check2.Enabled = False
    Check11.Enabled = False
    Check12.Enabled = False
    Check13.Enabled = False
End If
End Sub

Private Sub Combo1_Click()
'MsgBox Combo1.Text
If Combo1.Text = "psx (Sony PlayStation)" Then
    Check1.Enabled = True
    Check2.Enabled = True
    Check11.Enabled = True
    Check12.Enabled = True
    Check13.Enabled = True
Else
    Check1.Value = 0
    Check2.Value = 0
    Check9.Value = 0
    Check10.Value = 0
    Check11.Value = 0
    Check12.Value = 0
    Check13.Value = 0
    Check1.Enabled = False
    Check2.Enabled = False
    Check11.Enabled = False
    Check12.Enabled = False
    Check13.Enabled = False
End If

'Combo1.AddItem "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)", 7
'Combo1.AddItem "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)", 8
'Combo1.AddItem "pcfx (PC-FX)", 9
'Combo1.AddItem "psx (Sony PlayStation)", 10

If Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Or Combo1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
MsgBox "BIOS Image File: syscard3.pce is expected"
ElseIf Combo1.Text = "psx (Sony PlayStation)" Then
MsgBox "BIOS Image File: scph5500.bin/scph5501.bin/scph5502.bin is expected"
End If
End Sub

Private Sub Command1_Click()
Form1.Width = 12945
ActiveFile = "BIOS"
If Len(Text1.Text) >= 1 Then
    If Text1.Text = "Not Set" Then
        ResetBios = vbYes
    Else
        ResetBios = MsgBox("Reset Bios?", vbYesNo, "Reset Bios?")
    End If
    If ResetBios = vbYes Then
        BIOSFILE = ""
    Else
        a = Validate_Bios()
    End If
End If

End Sub

Private Sub Command2_Click()
Form1.Width = 12945
ActiveFile = "ROM"
If Len(Text2.Text) >= 1 Then
    If Text2.Text = "Not Set" Then
        ResetRom = vbYes
    Else
        ResetRom = MsgBox("Reset Rom?", vbYesNo, "Reset Rom?")
    End If
    If ResetRom = vbYes Then
        ROMFILE = ""
    Else
        a = Validate_Rom()
    End If
End If
End Sub

Private Sub Command3_Click()

If Combo1.Text = "gb (GameBoy (Color))" Then
    SYSCORE = "gb"
ElseIf Combo1.Text = "gba (GameBoy Advanced)" Then
    SYSCORE = "gba"
ElseIf Combo1.Text = "gg (Sega Game Gear)" Then
    SYSCORE = "gg"
ElseIf Combo1.Text = "lynx (Atari Lynx)" Then
    SYSCORE = "lynx"
ElseIf Combo1.Text = "md (Sega Genesis/MegaDrive)" Then
    SYSCORE = "md"
ElseIf Combo1.Text = "nes (Nintendo Entertainment System)" Then
    SYSCORE = "nes"
ElseIf Combo1.Text = "ngp (Neo Geo Pocket (Color))" Then
    SYSCORE = "ngp"
ElseIf Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SYSCORE = "pce"
ElseIf Combo1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SYSCORE = "pce_fast"
ElseIf Combo1.Text = "pcfx (PC-FX)" Then
    SYSCORE = "pcfx"
ElseIf Combo1.Text = "psx (Sony PlayStation)" Then
    SYSCORE = "psx"
    If Check11.Value = 0 And Check12.Value = 0 And Check13.Value = 0 Then
        MsgBox "A System Region must be select for PlayStation!", vbCritical, "Error!"
        FatalError = True
    End If
ElseIf Combo1.Text = "sms (Sega Master System)" Then
    SYSCORE = "sms"
ElseIf Combo1.Text = "snes (Super Nintendo Entertainment System)" Then
    SYSCORE = "snes"
ElseIf Combo1.Text = "VB (Virtual Boy)" Then
    SYSCORE = "vb"
ElseIf Combo1.Text = "wswan (WonderSwan)" Then
    SYSCORE = "wswan"
End If

If SYSCORE = "psx" Or SYSCORE = "pce" Or SYSCORE = "pce_fast" Then
    cmdstring = "cmd.exe /c " & Chr(34) & MedEXE & " -loadcd " & SYSCORE
    If SYSCORE = "psx" Then
        If Check11.Value = 1 Then
            cmdstring = cmdstring & " -psx.bios_na " & Chr(34) & BIOSFILE & Chr(34)
        ElseIf Check12.Value = 1 Then
            cmdstring = cmdstring & " -psx.bios_jp " & Chr(34) & BIOSFILE & Chr(34)
        ElseIf Check13.Value = 1 Then
            cmdstring = cmdstring & " -psx.bios_eu " & Chr(34) & BIOSFILE & Chr(34)
        End If
    End If
    If SYSCORE = "pce" Or SYSCORE = "pce_fast" Then
        cmdstring = cmdstring & "-pce.cdbios " & Chr(34) & BIOSFILE & Chr(34)
    End If
Else
    cmdstring = "cmd.exe /c " & Chr(34) & MedEXE & " -force_module " & SYSCORE
End If

If Combo2.Text = "0 - Disabled" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".stretch 0"
ElseIf Combo2.Text = "full - Full" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".stretch full"
ElseIf Combo2.Text = "aspect - Aspect Preserve" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".stretch aspect"
ElseIf Combo2.Text = "aspect_int - Aspect Preserve + Integer Scale" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".stretch aspect_int"
ElseIf Combo2.Text = "aspect_mult2 - Aspect Preserve + Integer Multiple-of-2 Scale" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".stretch aspect_mult2"
End If

If Combo3.Text = "None - None/Disabled" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".pixshader none"
ElseIf Combo3.Text = "autoip - Auto Interpolation" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".pixshader autoip"
ElseIf Combo3.Text = "autoipsharper - Sharper Auto Interpolation" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".pixshader autoipsharper"
ElseIf Combo3.Text = "scale2x - Scale2x" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".pixshader scale2x"
ElseIf Combo3.Text = "sabr - SABR v3.0" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".pixshader sabr"
ElseIf Combo3.Text = "ipsharper - Sharper bilinear interpolation." Then
    cmdstring = cmdstring & " -" & SYSCORE & ".pixshader ipsharper"
ElseIf Combo3.Text = "ipxnoty - Linear interpolation on X axis only." Then
    cmdstring = cmdstring & " -" & SYSCORE & ".pixshader ipxnoty"
ElseIf Combo3.Text = "ipynotx - Linear interpolation on Y axis only." Then
    cmdstring = cmdstring & " -" & SYSCORE & ".pixshader ipynotx"
ElseIf Combo3.Text = "ipxnotysharper - Sharper version of ipxnoty." Then
    cmdstring = cmdstring & " -" & SYSCORE & ".pixshader ipynotysharper"
ElseIf Combo3.Text = "ipynotxsharper - Sharper version of ipynotx." Then
    cmdstring = cmdstring & " -" & SYSCORE & ".pixshader ipynotxsharper"
End If

If Combo4.Text = "None - None/Disabled" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special none"
ElseIf Combo4.Text = "hq2x - hq2x" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special hq2x"
ElseIf Combo4.Text = "hq3x -hq3x" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special hq3x"
ElseIf Combo4.Text = "hq4x -hq4x" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special hq4x"
ElseIf Combo4.Text = "scale2x -scale2x" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special scale2x"
ElseIf Combo4.Text = "scale3x -scale3x" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special scale3x"
ElseIf Combo4.Text = "scale4x -scale4x" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special scale4x"
ElseIf Combo4.Text = "2xsai - 2xSaI" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special 2xsai"
ElseIf Combo4.Text = "supereagle - Super Eagle" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special supereagle"
ElseIf Combo4.Text = "nn2x - Nearest-neighbor 2x" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special nn2x"
ElseIf Combo4.Text = "nn3x - Nearest-neighbor 3x" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special nn3x"
ElseIf Combo4.Text = "nn4x - Nearest-neighbor 4x" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special nn4x"
ElseIf Combo4.Text = "nny2x - Nearest-neighbor 2x, y axis only" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special nny2x"
ElseIf Combo4.Text = "nny3x - Nearest-neighbor 3x, y axis only" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special nny3x"
ElseIf Combo4.Text = "nny4x - Nearest-neighbor 4x, y axis only" Then
    cmdstring = cmdstring & " -" & SYSCORE & ".special nny4x"
End If

If Check3.Value = 1 Then
    cmdstring = cmdstring & " -" & SYSCORE & ".tblur 1"
End If

If Check4.Value = 1 Then
    cmdstring = cmdstring & " -" & SYSCORE & ".tblur.accum 1"
    Sleep (100)
    cmdstring = cmdstring & " -" & SYSCORE & ".tblur.accum.amount " & Text3.Text
End If

If Check5.Value = 1 Then
    cmdstring = cmdstring & " -" & SYSCORE & ".videoip 1"
Else
    cmdstring = cmdstring & " -" & SYSCORE & ".videoip 0"
End If

If Check6.Value = 1 Then
    cmdstring = cmdstring & " -video.fs 1"
Else
    cmdstring = cmdstring & " -video.fs 0"
End If

If Check7.Value = 1 Then
    cmdstring = cmdstring & " -video.frameskip 1"
Else
    cmdstring = cmdstring & " -video.frameskip 0"
End If

If Check1.Value = 1 Then
    cmdstring = cmdstring & " -" & SYSCORE & ".bios_sanity 1"
End If

If Check2.Value = 1 Then
    cmdstring = cmdstring & " -" & SYSCORE & ".cd_sanity 1"
End If

If Val(Text4.Text) > 1 Then
    cmdstring = cmdstring & " -" & SYSCORE & ".xscale " & Text4.Text & " -" & SYSCORE & ".yscale " & Text4.Text & " -" & SYSCORE & ".xscalefs " & Text4.Text & " -" & SYSCORE & ".yscalefs " & Text4.Text
End If

If Len(Text5.Text) > 0 Then
    cmdstring = cmdstring & " -" & SYSCORE & ".xres " & Text5.Text
End If

If Len(Text6.Text) > 0 Then
    cmdstring = cmdstring & " -" & SYSCORE & ".yres " & Text6.Text
End If

If Len(Text7.Text) > 0 Then
    cmdstring = cmdstring & " -filesys.path_state " & Chr(34) & Text7.Text & Chr(34) & " -filesys.path_sav " & Chr(34) & Text7.Text & Chr(34)
End If
'
cmdstring = cmdstring & " " & Chr(34) & ROMFILE & Chr(34)

cmdstring = cmdstring & Chr(34)

If Check8.Value = 1 And FatalError = False Then
    MsgBox cmdstring
End If

If FatalError = False Then
    Shell (cmdstring)
Else
    FatalError = False
End If
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
Label6.Caption = "Not Set"
Label9.Caption = "Not Set"
End Sub

Private Sub Command5_Click()
Form1.Width = 12945
ActiveFile = "SAVE"
If Len(Text7.Text) >= 1 Then
    If Text7.Text = "Not Set" Then
        ResetSave = vbYes
    Else
        ResetSave = MsgBox("Reset Save Path?", vbYesNo, "Reset Save Path?")
    End If
    If ResetSave = vbYes Then
        SavePath = ""
    Else
        a = a
    End If
End If

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
End If
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Drive1_GotFocus()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
If ActiveFile = "MEDEXE" Then
    tmp2 = MsgBox("Set File: " & File1.FileName, vbYesNo, "Set this file?")
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
    tmp2 = MsgBox("Set File: " & File1.FileName, vbYesNo, "Set this file?")
    If tmp2 = vbYes Then
        Text2.Text = Dir1.Path & "\" & File1.FileName
        ROMFILE = Text2.Text
        Form1.Width = 9240
        ActiveFile = "None"
        tmp2 = ""
        a = Validate_Rom()
    End If
End If

End Sub

Private Sub Form_Load()
'12945
'9240
FatalError = False
Form1.Width = 9240
ActiveFile = "None"
Label29.Visible = False
'Comments
'C:\EMU\mednafen-0.9.38.7-win64\mednafen.exe
'MD5 8F0BC836E2B6023371B99E94829B5CF1

'Credits
'md5.exe Source: https://www.fourmilab.ch/md5/
'MD5.EXE ACKNOWLEDGEMENTS
'The MD5 algorithm was developed by Ron Rivest. The public domain C language implementation used in this program was written by Colin Plumb in 1993.
Build = "0.1.1"
Form1.Caption = "MedAdvCFG v" & Build & " (Mednafen v0.9.38.x Frontend) by Nigel Todman"

Dir1.Path = VB.App.Path
File1.Path = VB.App.Path

Label2.Caption = "Not Set"
'MedEXE = "C:\EMU\mednafen-0.9.38.7-win64\mednafen.exe"

Label16.Caption = "Hotkeys:"
Set FSO = CreateObject("Scripting.FileSystemObject")

If FSO.FileExists(VB.App.Path & "\MedAdvCFG.dat") Then

Open VB.App.Path & "\MedAdvCFG.dat" For Input As #1
    For x = 1 To 22
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

Text1.Text = BIOSFILE
Text2.Text = ROMFILE
Text4.Text = ScaleFactor
Text5.Text = XRes
Text6.Text = YRes
Text7.Text = SavePath

Dir1.Path = LastPath
File1.Path = LastPath

Combo1.Text = SystemCore
Combo2.Text = Stretch
Combo3.Text = PixelShader
Combo4.Text = VideoScaler

If BIOSSanity = 1 Then
    Check1.Value = 1
End If

If ROMSanity = 1 Then
    Check2.Value = 1
End If

If Fullscreen = 1 Then
    Check6.Value = 1
End If

If Frameskip = 1 Then
    Check7.Value = 1
End If

If TBlur = 1 Then
    Check3.Value = 1
End If

If TblurAccum = 1 Then
    Check4.Value = 1
    Text3.Text = AccumAmount
End If

If VideoIP = 1 Then
    Check5.Value = 1
End If

If SystemRegion = None Then
    a = a
ElseIf SystemRegion = "NTSC-U" Then
    Check11.Value = 1
ElseIf SystemRegion = "NTSC-J" Then
    Check12.Value = 1
ElseIf SystemRegion = "PAL" Then
    Check13.Value = 1
End If

a = Validate_MedEXE()
a = Validate_Rom()
a = Validate_Bios()
End If

Combo1.AddItem "gb (GameBoy (Color))", 0
Combo1.AddItem "gba (GameBoy Advanced)", 1
Combo1.AddItem "gg (Sega Game Gear)", 2
Combo1.AddItem "lynx (Atari Lynx)", 3
Combo1.AddItem "md (Sega Genesis/MegaDrive)", 4
Combo1.AddItem "nes (Nintendo Entertainment System)", 5
Combo1.AddItem "ngp (Neo Geo Pocket (Color))", 6
Combo1.AddItem "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)", 7
Combo1.AddItem "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)", 8
Combo1.AddItem "pcfx (PC-FX)", 9
Combo1.AddItem "psx (Sony PlayStation)", 10
Combo1.AddItem "sms (Sega Master System)", 11
Combo1.AddItem "snes (Super Nintendo Entertainment System)", 12
Combo1.AddItem "VB (Virtual Boy)", 13
Combo1.AddItem "wswan (WonderSwan)", 14

Combo2.AddItem "0 - Disabled", 0
Combo2.AddItem "full - Full", 1
Combo2.AddItem "aspect - Aspect Preserve", 2
Combo2.AddItem "aspect_int - Aspect Preserve + Integer Scale", 3
Combo2.AddItem "aspect_mult2 - Aspect Preserve + Integer Multiple-of-2 Scale", 4

Combo3.AddItem "None - None/Disabled", 0
Combo3.AddItem "autoip - Auto Interpolation", 1
Combo3.AddItem "autoipsharper - Sharper Auto Interpolation", 2
Combo3.AddItem "scale2x - Scale2x", 3
Combo3.AddItem "sabr - SABR v3.0", 4
Combo3.AddItem "ipsharper - Sharper bilinear interpolation.", 5
Combo3.AddItem "ipxnoty - Linear interpolation on X axis only.", 6
Combo3.AddItem "ipynotx - Linear interpolation on Y axis only.", 7
Combo3.AddItem "ipxnotysharper - Sharper version of ipxnoty.", 8
Combo3.AddItem "ipynotxsharper - Sharper version of ipynotx.", 9

Combo4.AddItem "None - None/Disabled", 0
Combo4.AddItem "hq2x - hq2x", 1
Combo4.AddItem "hq3x -hq3x", 2
Combo4.AddItem "hq4x -hq4x", 3
Combo4.AddItem "scale2x -scale2x", 4
Combo4.AddItem "scale3x -scale3x", 5
Combo4.AddItem "scale4x -scale4x", 6
Combo4.AddItem "2xsai - 2xSaI", 7
Combo4.AddItem "super2xsai - Super 2xSaI", 8
Combo4.AddItem "supereagle - Super Eagle", 9
Combo4.AddItem "nn2x - Nearest-neighbor 2x", 10
Combo4.AddItem "nn3x - Nearest-neighbor 3x", 11
Combo4.AddItem "nn4x - Nearest-neighbor 4x", 12
Combo4.AddItem "nny2x - Nearest-neighbor 2x, y axis only", 13
Combo4.AddItem "nny3x - Nearest-neighbor 3x, y axis only", 14
Combo4.AddItem "nny4x - Nearest-neighbor 4x, y axis only", 15

If FSO.FileExists(MedEXE) = False Then
    Form1.Width = 12945
    ActiveFile = "MEDEXE"
    MsgBox "Select your Mednafen EXE to get started!"
End If


End Sub

Private Sub Image1_Click()
Shell ("cmd.exe /c start http://mednafen.fobby.net/"), vbHide
End Sub

Private Sub Label15_Click()
Shell ("cmd.exe /c start http://www.NigelTodman.com"), vbHide
End Sub

Private Sub Label2_Click()
Clipboard.Clear
Clipboard.SetText Label2.Caption
MsgBox "MD5 copied to Clipboard"
End Sub

Private Sub Label6_Click()
Clipboard.Clear
Clipboard.SetText Label6.Caption
MsgBox "MD5 copied to Clipboard"
End Sub

Private Sub Label9_Click()
Clipboard.Clear
Clipboard.SetText Label9.Caption
MsgBox "MD5 copied to Clipboard"
End Sub

Private Sub Quit_Click()
Unload Form1
End Sub

Private Sub Save_Settings_Click()
Open VB.App.Path & "\MedAdvCFG.dat" For Output As #1
    Print #1, "MedEXE=" & MedEXE
    Print #1, "SystemCore=" & Combo1.Text
    Print #1, "SystemBIOS=" & BIOSPATH & "\" & BIOSFILE
    Print #1, "BIOSSanity=" & Check1.Value
    Print #1, "RomImage=" & ROMFILE
    Print #1, "ROMSanity=" & Check2.Value
    Print #1, "Stretch=" & Combo2.Text
    Print #1, "PixelShader=" & Combo3.Text
    Print #1, "VideoScaler=" & Combo4.Text
    Print #1, "Fullscreen=" & Check6.Value
    Print #1, "Frameskip=" & Check7.Value
    Print #1, "Tblur=" & Check3.Value
    Print #1, "TblurAccum=" & Check4.Value
    Print #1, "AccumAmount=" & Text3.Text
    Print #1, "VideoIP=" & Check5.Value
    Print #1, "XRes=" & Text5.Text
    Print #1, "YRes=" & Text6.Text
    Print #1, "ScaleFactor=" & Text4.Text
    Print #1, "LastPath=" & File1.Path
    Print #1, "BiosPathLoad=" & BIOSPATH
    Print #1, "SavePath=" & Text7.Text
    If Check11.Value = 0 And Check12.Value = 0 And Check13.Value = 0 Then
        Print #1, "SystemRegion=None"
    ElseIf Check11.Value = 1 Then
        Print #1, "SystemRegion=NTSC-U"
    ElseIf Check12.Value = 1 Then
        Print #1, "SystemRegion=NTSC-J"
    ElseIf Check13.Value = 1 Then
        Print #1, "SystemRegion=PAL"
    End If
Close #1
End Sub


