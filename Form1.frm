VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "MedAdvCFG v0.0.0 (Mednafen v0.9.x.x Frontend) by Nigel Todman [ADV MODE]"
   ClientHeight    =   6810
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12630
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text13 
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
      Height          =   285
      Left            =   6375
      TabIndex        =   86
      Text            =   "Port"
      Top             =   3960
      Width           =   510
   End
   Begin VB.TextBox Text12 
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
      Height          =   285
      Left            =   4800
      TabIndex        =   84
      Text            =   "Hostname/IP"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CheckBox Check23 
      BackColor       =   &H00C0C0C0&
      Caption         =   "cd.image_memcache"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   83
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CheckBox Check22 
      BackColor       =   &H00C0C0C0&
      Caption         =   "video.blit_timesync"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   82
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CheckBox Check21 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Disable Sound"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   81
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CheckBox Check20 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Force Mono"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   80
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CheckBox Check19 
      BackColor       =   &H00C0C0C0&
      Caption         =   "video.glvsync"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   79
      Top             =   3720
      Width           =   1455
   End
   Begin VB.ComboBox Combo6 
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
      Left            =   1320
      TabIndex        =   77
      Text            =   "None - None/Disabled"
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text11 
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
      Height          =   285
      Left            =   8520
      TabIndex        =   73
      Text            =   "1.00"
      Top             =   5880
      Width           =   405
   End
   Begin VB.CheckBox Check18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "weave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   71
      Top             =   2880
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Check17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "bob"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2200
      TabIndex        =   70
      Top             =   2880
      Width           =   615
   End
   Begin VB.CheckBox Check16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "bob_offset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3050
      TabIndex        =   69
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text10 
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
      Height          =   285
      Left            =   8640
      TabIndex        =   67
      Text            =   "0"
      Top             =   5520
      Width           =   285
   End
   Begin VB.CheckBox Check15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Yes to All Prompts(!)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   66
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text9 
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
      Height          =   285
      Left            =   1320
      TabIndex        =   64
      Top             =   5040
      Width           =   4575
   End
   Begin VB.TextBox Text8 
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
      Height          =   285
      Left            =   8640
      TabIndex        =   62
      Text            =   "1"
      Top             =   6240
      Width           =   285
   End
   Begin VB.ComboBox Combo5 
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
      ItemData        =   "Form1.frx":851A
      Left            =   1320
      List            =   "Form1.frx":851C
      TabIndex        =   60
      Text            =   "gamepad - SCPH-1080 PlayStation Digital Gamepad"
      Top             =   4680
      Width           =   4575
   End
   Begin VB.CheckBox Check14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "untrusted_fip_check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   59
      Top             =   3960
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3050
      TabIndex        =   56
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox Check12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NTSC-J"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2200
      TabIndex        =   55
      Top             =   720
      Width           =   975
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NTSC-U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   54
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text7 
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
      Height          =   285
      Left            =   1320
      TabIndex        =   52
      Text            =   "Not Set"
      Top             =   4320
      Width           =   4575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set"
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
      Left            =   6000
      TabIndex        =   51
      Top             =   4320
      Width           =   615
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BIOS MD5 Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   50
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM MD5 Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   49
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text6 
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
      Height          =   285
      Left            =   8400
      TabIndex        =   47
      Text            =   "1080"
      Top             =   4440
      Width           =   515
   End
   Begin VB.TextBox Text5 
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
      Height          =   285
      Left            =   7920
      TabIndex        =   46
      Text            =   "1920"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox Text4 
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
      Height          =   285
      Left            =   8640
      TabIndex        =   44
      Text            =   "2"
      Top             =   4800
      Width           =   285
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear"
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
      Left            =   6000
      TabIndex        =   43
      Top             =   480
      Width           =   615
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
      Left            =   9000
      TabIndex        =   42
      Top             =   120
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
      Height          =   3210
      Left            =   9000
      TabIndex        =   39
      Top             =   3480
      Width           =   3615
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
      Height          =   3015
      Left            =   9000
      TabIndex        =   38
      Top             =   450
      Width           =   3615
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Review Command Line before execution?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   37
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frameskip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   36
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fullscreen (FS)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   28
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Launch mednafen.exe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   27
      Top             =   3000
      Width           =   1815
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   1320
      TabIndex        =   25
      Text            =   "None - None/Disabled"
      Top             =   3960
      Width           =   3375
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bilinear interpolation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   24
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text3 
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
      Height          =   285
      Left            =   8640
      TabIndex        =   22
      Text            =   "50"
      Top             =   5160
      Width           =   285
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Accumulate color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   21
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temporal Blur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   20
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   1320
      TabIndex        =   18
      Text            =   "None - None/Disabled"
      Top             =   3600
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   1320
      TabIndex        =   16
      Text            =   "0 - Disabled"
      Top             =   2520
      Width           =   3375
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM Sanity Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   15
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set"
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
      Left            =   6000
      TabIndex        =   12
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text2 
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
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Text            =   "Not Set"
      Top             =   1920
      Width           =   4575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BIOS Sanity Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   9
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set"
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
      Left            =   6000
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
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
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "Not Set"
      Top             =   960
      Width           =   4575
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1320
      TabIndex        =   2
      Text            =   "gb (GameBoy (Color))"
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Netplay Host:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   85
      Top             =   3720
      Width           =   960
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Video Driver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   78
      Top             =   3240
      Width           =   870
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Homepage: www.NigelTodman.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1320
      TabIndex        =   76
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MedAdvCFG v0.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1320
      TabIndex        =   75
      Top             =   5400
      Width           =   1665
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Axis Scale:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   74
      Top             =   5880
      Width           =   780
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Deinterlacer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   72
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scanlines:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   68
      Top             =   5520
      Width           =   735
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   2760
      Picture         =   "Form1.frx":851E
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   2280
      Picture         =   "Form1.frx":91E8
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   3720
      Picture         =   "Form1.frx":9EB2
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   3240
      Picture         =   "Form1.frx":AB7C
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1800
      Picture         =   "Form1.frx":B846
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1320
      Picture         =   "Form1.frx":C510
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Custom Params:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   65
      Top             =   5040
      Width           =   1140
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Number of Players:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   63
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Controller:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   61
      Top             =   4680
      Width           =   705
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
      TabIndex        =   58
      Top             =   1560
      Width           =   6900
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "System Region:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   57
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   53
      Top             =   4320
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   4920
      Picture         =   "Form1.frx":D1DA
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1665
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Resolution"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   48
      Top             =   4440
      Width           =   750
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scaling Factor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   45
      Top             =   4800
      Width           =   1020
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
      TabIndex        =   41
      Top             =   7380
      Visible         =   0   'False
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
      TabIndex        =   40
      Top             =   7180
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "F5:Save state"
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
      Top             =   5880
      Width           =   1200
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "F7:Load state"
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
      Top             =   6120
      Width           =   1185
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "F10:Reset Console"
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
      Top             =   6360
      Width           =   1620
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Alt+Enter:Toggle fullscreen mode"
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
      Top             =   6600
      Width           =   2835
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "CTRL+SHIFT+1:Select Controller Type"
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
      Left            =   3240
      TabIndex        =   31
      Top             =   6360
      Width           =   3300
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALT+SHIFT+1:Configure buttons for Controller Port "
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
      Left            =   3240
      TabIndex        =   30
      Top             =   6600
      Width           =   4395
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
      TabIndex        =   29
      Top             =   5520
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Video Scaler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temporal Blur amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   23
      Top             =   5160
      Width           =   1590
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pixel Shader"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   885
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "FS Stretch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM MD5:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM Image:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Not Set"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "BIOS MD5:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "System BIOS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "System Core:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   6345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mednafen EXE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Gen_M3U 
         Caption         =   "Generate Multi-Disc M3U"
      End
      Begin VB.Menu Save_Settings 
         Caption         =   "Save Settings"
      End
      Begin VB.Menu Reset_Settings 
         Caption         =   "Reset Settings"
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mode 
      Caption         =   "Mode"
      Begin VB.Menu advanced 
         Caption         =   "Advanced"
         Checked         =   -1  'True
      End
      Begin VB.Menu basic 
         Caption         =   "Basic"
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

Dim MedEXE, FSO, tmp, tmp2, tmp3(99), BIOSFILE, BIOSPATH, ROMFILE, SystemCore, SysCore, BIOSSanity, ROMSanity, Stretch, PixelShader, VideoScaler, x, y, z
Dim cmdstring, Build, Frameskip, Fullscreen, TBlur, TblurAccum, AccumAmount, VideoIP, ActiveFile, XRes, YRes, ScaleFactor, LastPath, SavePath, BiosPathLoad
Dim ResetBios, ResetRom, ResetSave, FatalError, SystemRegion, SystemRegionLoad, ROMDIR, M3USize, LastFile, VideoDriver
Dim Bilinear, DisableSound, ForceMono, video_blit_timesync, video_glvsync, untrusted_fip_check, cd_image_memcache, scanlines, numplayers, customparams

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Function Generate_M3U(M3USize As Integer)
'*** v0.1.3
Form1.Width = 12945
ActiveFile = "M3U"
MsgBox "Select the first disc"
z = 0
Open VB.App.Path & "\multi.m3u" For Output As #2

End Function
Public Function Validate_Rom()
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
Function Validate_Bios()
If Check10.Value = 1 Then
    If FSO.FileExists(BIOSPATH & "\" & BIOSFILE) = True Then
        Shell ("cmd.exe /c " & Chr(34) & VB.App.Path & "\md5.exe -n " & Chr(34) & BIOSPATH & "\" & BIOSFILE & Chr(34) & " >> " & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
        Sleep (200)
        If FSO.FileExists(VB.App.Path & "\md5.txt") = True Then
            Open VB.App.Path & "\md5.txt" For Input As #5
                Line Input #5, tmp
            Close #5
        End If
        Shell ("cmd.exe /c del " & Chr(34) & VB.App.Path & "\md5.txt" & Chr(34)), vbHide
        Text1.Text = BIOSFILE
    End If
    'v0.2.0
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
        ElseIf LCase(tmp) = "3240872c70984b6cbfda1586cab68dbe" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: SEGA SATURN V1.01A US/EU"
            Check11.Value = 1
            Check11.Value = 1
        ElseIf LCase(tmp) = "85ec9ca47d8f6807718151cbcca8b964" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: SEGA SATURN V1.01 JP"
            Check12.Value = 1
            Check12.Value = 1
        ElseIf LCase(tmp) = "af5828fdff51384f99b3c4926be27762" Then
            Label6.Caption = "MD5: " & tmp
            Label29.Visible = True
            Label29.Caption = "Valid: SEGA SATURN V1.00 JP"
            Check12.Value = 1
            Check12.Value = 1
        Else
            Label29.Visible = False
            Label6.Caption = "MD5: " & tmp
        End If
Else
    Label6.Caption = "MD5: BIOS MD5 Disabled"
End If
Validate_Bios = tmp
End Function
Public Function SetSysCore()
If Form1.Combo1.Text = "gb (GameBoy (Color))" Then
    SysCore = "gb"
    Form3.Text1.Text = "gb (GameBoy (Color))" & vbCrLf
ElseIf Form1.Combo1.Text = "gba (GameBoy Advanced)" Then
    SysCore = "gba"
    Form3.Text1.Text = "gba (GameBoy Advanced)" & vbCrLf
ElseIf Form1.Combo1.Text = "gg (Sega Game Gear)" Then
    SysCore = "gg"
    Form3.Text1.Text = "gg (Sega Game Gear)" & vbCrLf
ElseIf Form1.Combo1.Text = "lynx (Atari Lynx)" Then
    SysCore = "lynx"
    Form3.Text1.Text = "lynx (Atari Lynx)" & vbCrLf
ElseIf Form1.Combo1.Text = "md (Sega Genesis/MegaDrive)" Then
    SysCore = "md"
    Form3.Text1.Text = "md (Sega Genesis/MegaDrive)" & vbCrLf
ElseIf Form1.Combo1.Text = "nes (Nintendo Entertainment System)" Then
    SysCore = "nes"
    Form3.Text1.Text = "nes (Nintendo Entertainment System)" & vbCrLf
ElseIf Form1.Combo1.Text = "ngp (Neo Geo Pocket (Color))" Then
    SysCore = "ngp"
    Form3.Text1.Text = "ngp (Neo Geo Pocket (Color))" & vbCrLf
ElseIf Form1.Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SysCore = "pce"
    Form3.Text1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" & vbCrLf
ElseIf Form1.Combo1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SysCore = "pce_fast"
    Form3.Text1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" & vbCrLf
ElseIf Form1.Combo1.Text = "pcfx (PC-FX)" Then
    SysCore = "pcfx"
    Form3.Text1.Text = "pcfx (PC-FX)" & vbCrLf
ElseIf Form1.Combo1.Text = "psx (Sony PlayStation)" Then
    SysCore = "psx"
    Form3.Text1.Text = "psx (Sony PlayStation)" & vbCrLf
ElseIf Form1.Combo1.Text = "sms (Sega Master System)" Then
    SysCore = "sms"
    Form3.Text1.Text = "sms (Sega Master System)" & vbCrLf
ElseIf Form1.Combo1.Text = "snes (Super Nintendo Entertainment System)" Then
    SysCore = "snes"
    Form3.Text1.Text = "snes (Super Nintendo Entertainment System)" & vbCrLf
ElseIf Form1.Combo1.Text = "ss (Sega Saturn)" Then
    SysCore = "ss"
    Form3.Text1.Text = "ss (Sega Saturn)" & vbCrLf
ElseIf Form1.Combo1.Text = "vb (Virtual Boy)" Then
    SysCore = "vb"
    Form3.Text1.Text = "vb (Virtual Boy)" & vbCrLf
ElseIf Form1.Combo1.Text = "wswan (WonderSwan)" Then
    SysCore = "wswan"
    Form3.Text1.Text = "wswan (WonderSwan)" & vbCrLf
End If
SetSysCore = SysCore
End Function
Public Function Validate_MedEXE()
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

Private Sub About_Click()
MsgBox "MedAdvCFG v" & Build & " (Mednafen v0.9.x.x Frontend)" & vbCrLf & "Written by Nigel Todman (nigel@nigeltodman.com)" & vbCrLf & "Primarily written as a PSX Frontend." & vbCrLf & "Tested with the following System Cores:" & vbCrLf & "GB, GBA, GG, MD, NES, PCE, PCE_FAST, PSX, SNES, SS, VB" & vbCrLf & vbCrLf & "Homepage: www.NigelTodman.com" & vbCrLf & "Facebook: facebook.com/nigel.todman.3" & vbCrLf & "Twitter: @Veritas_83" & vbCrLf & "YouTube: Veritas0923"
End Sub

Private Sub Advanced_Click()
advanced.Checked = True
basic.Checked = False
Form1.Visible = True
Form2.Visible = False
End Sub

Private Sub Basic_Click()
advanced.Checked = False
basic.Checked = True
Form1.Visible = False
Form2.Visible = True
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

Private Sub Check15_Click()
If Check15.Value = 1 Then
    a = a
Else
    tmp = MsgBox("Auto Confirm all prompts like this one?", vbYesNo)
End If

If tmp = vbYes Then
    Check15.Value = 1
End If

If tmp = vbNo Then
    Check15.Value = 0
End If
End Sub

Private Sub Check16_Click()
Check17.Value = 0
Check18.Value = 0
End Sub

Private Sub Check17_Click()
Check18.Value = 0
Check16.Value = 0
End Sub

Private Sub Check18_Click()
Check17.Value = 0
Check16.Value = 0
End Sub

Private Sub Combo1_Change()
'MsgBox Combo1.Text
If Combo1.Text = "psx (Sony PlayStation)" Then
    Check1.Enabled = True
    Check2.Enabled = True
    Check11.Enabled = True
    Check12.Enabled = True
    Check13.Enabled = True
    Combo5.Enabled = True
    Check1.Value = 1
    Check2.Value = 1
    Check9.Value = 1
    Check10.Value = 1
    Label29.Visible = True
ElseIf Combo1.Text = "snes (Super Nintendo Entertainment System)" Then
    For z = 0 To Combo5.ListCount
    'Combo5.RemoveItem (z)
    Next z
ElseIf Combo1.Text = "ss (Sega Saturn)" Then
    Check1.Enabled = True
    Check2.Enabled = True
    Check11.Enabled = True
    Check12.Enabled = True
    Check13.Enabled = True
    Combo5.Enabled = True
    Check1.Value = 1
    Check2.Value = 1
    Check9.Value = 1
    Check10.Value = 1
    Label29.Visible = True
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
    Combo5.Enabled = False
    Label29.Visible = False
End If
End Sub

Private Sub Combo1_Click()
'MsgBox Combo1.Text
If Combo1.Text = "psx (Sony PlayStation)" Then
    For z = 1 To Combo5.ListCount
        Combo5.RemoveItem 0
    Next z
    Combo5.AddItem "none", 0
    Combo5.AddItem "gamepad - SCPH-1080 PlayStation Digital Gamepad", 1
    Combo5.AddItem "dualshock - SCPH-1200 PlayStation DualShock Gamepad", 2
    Combo5.AddItem "dualanalog - SCPH-1180 PlayStation DualAnalog Gamepad", 3
    Combo5.AddItem "analogjoy - SCPH-1110 PlayStation Analog Joystick", 4
    Combo5.AddItem "mouse - SCPH-1090 PlayStation Mouse", 5
    Combo5.AddItem "negcon - NPC-101 Namco neGcon", 6
    Combo5.AddItem "guncon - NPC-103 Namco GunCon", 7
    Combo5.AddItem "justifier - SLUH-00017 Konami Justifier", 8
    Combo5.AddItem "dancepad - SLUH-00071 Konami Dancepad", 9
    Combo5.ListIndex = 1
    Check1.Enabled = True
    Check2.Enabled = True
    Check11.Enabled = True
    Check12.Enabled = True
    Check13.Enabled = True
    Combo5.Enabled = True
    Check1.Value = 1
    Check2.Value = 1
    Check9.Value = 1
    Check10.Value = 1
    Label29.Visible = True
ElseIf Combo1.Text = "snes (Super Nintendo Entertainment System)" Then
    For z = 1 To Combo5.ListCount
        Combo5.RemoveItem 0
    Next z
    Combo5.AddItem "none", 0
    Combo5.AddItem "gamepad - SNS-005 Super Nintendo Controller", 1
    Combo5.AddItem "mouse - SNS-016 Super Nintendo Mouse Controller", 2
    Combo5.AddItem "superscope - SNS-013 Super Nintendo Super Scope", 3
    Combo5.ListIndex = 1
    Combo5.Enabled = True
ElseIf Combo1.Text = "nes (Nintendo Entertainment System)" Then
    For z = 1 To Combo5.ListCount
        Combo5.RemoveItem 0
    Next z
    Combo5.AddItem "none", 0
    Combo5.AddItem "gamepad - NES-004 Nintendo Control Pad", 1
    Combo5.AddItem "zapper - NES-005 Nintendo Zapper", 2
    Combo5.AddItem "powerpada - NES-028 Nintendo Power Pad", 3
    Combo5.AddItem "powerpadb - NES-028 Nintendo Power Pad", 4
    Combo5.AddItem "arkanoid - NES-XXX Taito Arkanoid Vaus Controller", 5
    Combo5.ListIndex = 1
    Combo5.Enabled = True
ElseIf Combo1.Text = "ss (Sega Saturn)" Then
    Check1.Enabled = True
    Check2.Enabled = True
    Check9.Value = 1
    Check10.Value = 1
    Label29.Visible = True
    For z = 1 To Combo5.ListCount
        Combo5.RemoveItem 0
    Next z
    Combo5.AddItem "none", 0
    Combo5.AddItem "gamepad - MK-80100 Sega Saturn Controller", 1
    Combo5.AddItem "3dpad - MK-80117 Sega Saturn 3D Control Pad", 2
    Combo5.AddItem "mouse - HSS-0139 Sega Saturn Shuttle Mouse", 3
    Combo5.ListIndex = 1
    Combo5.Enabled = True
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
    Combo5.Enabled = False
    Label29.Visible = False
End If

'Combo1.AddItem "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)", 7
'Combo1.AddItem "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)", 8
'Combo1.AddItem "pcfx (PC-FX)", 9
'Combo1.AddItem "psx (Sony PlayStation)", 10

If Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Or Combo1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    MsgBox "BIOS Image File: syscard3.pce is expected"
ElseIf Combo1.Text = "psx (Sony PlayStation)" Then
    MsgBox "BIOS Image File: scph5500.bin/scph5501.bin/scph5502.bin is expected"
ElseIf Combo1.Text = "ss (Sega Saturn)" Then
    MsgBox "BIOS Image File: sega_101.bin/mpr-17933.bin is expected"
End If
End Sub

Private Sub Command1_Click()
Form1.Width = 12945
ActiveFile = "BIOS"
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
        If Check15.Value = 1 Then
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
        a = Validate_Rom()
    End If
End If
End Sub

Private Sub Command3_Click()
If Combo1.Text = "gb (GameBoy (Color))" Then
    SysCore = "gb"
ElseIf Combo1.Text = "gba (GameBoy Advanced)" Then
    SysCore = "gba"
ElseIf Combo1.Text = "gg (Sega Game Gear)" Then
    SysCore = "gg"
ElseIf Combo1.Text = "lynx (Atari Lynx)" Then
    SysCore = "lynx"
ElseIf Combo1.Text = "md (Sega Genesis/MegaDrive)" Then
    SysCore = "md"
ElseIf Combo1.Text = "nes (Nintendo Entertainment System)" Then
    SysCore = "nes"
ElseIf Combo1.Text = "ngp (Neo Geo Pocket (Color))" Then
    SysCore = "ngp"
ElseIf Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SysCore = "pce"
ElseIf Combo1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SysCore = "pce_fast"
ElseIf Combo1.Text = "pcfx (PC-FX)" Then
    SysCore = "pcfx"
ElseIf Combo1.Text = "psx (Sony PlayStation)" Then
    SysCore = "psx"
    If Check11.Value = 0 And Check12.Value = 0 And Check13.Value = 0 Then
        MsgBox "A System Region must be select for PlayStation!", vbCritical, "Error!"
        FatalError = True
    End If
ElseIf Combo1.Text = "sms (Sega Master System)" Then
    SysCore = "sms"
ElseIf Combo1.Text = "snes (Super Nintendo Entertainment System)" Then
    SysCore = "snes"
'v0.1.8
'Combo1.AddItem "ss (Sega Saturn)", 13
ElseIf Combo1.Text = "ss (Sega Saturn)" Then
    SysCore = "ss"
ElseIf Combo1.Text = "vb (Virtual Boy)" Then
    SysCore = "vb"
ElseIf Combo1.Text = "wswan (WonderSwan)" Then
    SysCore = "wswan"
End If

If SysCore = "psx" Or SysCore = "pce" Or SysCore = "pce_fast" Or SysCore = "ss" Then
    cmdstring = "cmd.exe /c " & Chr(34) & MedEXE & " -loadcd " & SysCore
    If SysCore = "psx" Then
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
        If Check12.Value = 1 Then
            cmdstring = cmdstring & "-ss.bios_jp " & Chr(34) & BIOSFILE & Chr(34)
        Else
            cmdstring = cmdstring & "-ss.bios_na_eu " & Chr(34) & BIOSFILE & Chr(34)
        End If
    End If
Else
    cmdstring = "cmd.exe /c " & Chr(34) & MedEXE & " -force_module " & SysCore
End If

If Combo2.Text = "0 - Disabled" Then
    cmdstring = cmdstring & " -" & SysCore & ".stretch 0"
ElseIf Combo2.Text = "full - Full" Then
    cmdstring = cmdstring & " -" & SysCore & ".stretch full"
ElseIf Combo2.Text = "aspect - Aspect Preserve" Then
    cmdstring = cmdstring & " -" & SysCore & ".stretch aspect"
ElseIf Combo2.Text = "aspect_int - Aspect Preserve + Integer Scale" Then
    cmdstring = cmdstring & " -" & SysCore & ".stretch aspect_int"
ElseIf Combo2.Text = "aspect_mult2 - Aspect Preserve + Integer Multiple-of-2 Scale" Then
    cmdstring = cmdstring & " -" & SysCore & ".stretch aspect_mult2"
End If

If Combo3.Text = "None - None/Disabled" Then
    cmdstring = cmdstring & " -" & SysCore & ".shader none"
ElseIf Combo3.Text = "autoip - Auto Interpolation" Then
    cmdstring = cmdstring & " -" & SysCore & ".shader autoip"
ElseIf Combo3.Text = "autoipsharper - Sharper Auto Interpolation" Then
    cmdstring = cmdstring & " -" & SysCore & ".shader autoipsharper"
ElseIf Combo3.Text = "scale2x - Scale2x" Then
    cmdstring = cmdstring & " -" & SysCore & ".shader scale2x"
ElseIf Combo3.Text = "sabr - SABR v3.0" Then
    cmdstring = cmdstring & " -" & SysCore & ".shader sabr"
ElseIf Combo3.Text = "ipsharper - Sharper bilinear interpolation." Then
    cmdstring = cmdstring & " -" & SysCore & ".shader ipsharper"
ElseIf Combo3.Text = "ipxnoty - Linear interpolation on X axis only." Then
    cmdstring = cmdstring & " -" & SysCore & ".shader ipxnoty"
ElseIf Combo3.Text = "ipynotx - Linear interpolation on Y axis only." Then
    cmdstring = cmdstring & " -" & SysCore & ".shader ipynotx"
ElseIf Combo3.Text = "ipxnotysharper - Sharper version of ipxnoty." Then
    cmdstring = cmdstring & " -" & SysCore & ".shader ipynotysharper"
ElseIf Combo3.Text = "ipynotxsharper - Sharper version of ipynotx." Then
    cmdstring = cmdstring & " -" & SysCore & ".shader ipynotxsharper"
End If

If Combo4.Text = "None - None/Disabled" Then
    cmdstring = cmdstring & " -" & SysCore & ".special none"
ElseIf Combo4.Text = "hq2x - hq2x" Then
    cmdstring = cmdstring & " -" & SysCore & ".special hq2x"
ElseIf Combo4.Text = "hq3x -hq3x" Then
    cmdstring = cmdstring & " -" & SysCore & ".special hq3x"
ElseIf Combo4.Text = "hq4x -hq4x" Then
    cmdstring = cmdstring & " -" & SysCore & ".special hq4x"
ElseIf Combo4.Text = "scale2x -scale2x" Then
    cmdstring = cmdstring & " -" & SysCore & ".special scale2x"
ElseIf Combo4.Text = "scale3x -scale3x" Then
    cmdstring = cmdstring & " -" & SysCore & ".special scale3x"
ElseIf Combo4.Text = "scale4x -scale4x" Then
    cmdstring = cmdstring & " -" & SysCore & ".special scale4x"
ElseIf Combo4.Text = "2xsai - 2xSaI" Then
    cmdstring = cmdstring & " -" & SysCore & ".special 2xsai"
ElseIf Combo4.Text = "supereagle - Super Eagle" Then
    cmdstring = cmdstring & " -" & SysCore & ".special supereagle"
ElseIf Combo4.Text = "nn2x - Nearest-neighbor 2x" Then
    cmdstring = cmdstring & " -" & SysCore & ".special nn2x"
ElseIf Combo4.Text = "nn3x - Nearest-neighbor 3x" Then
    cmdstring = cmdstring & " -" & SysCore & ".special nn3x"
ElseIf Combo4.Text = "nn4x - Nearest-neighbor 4x" Then
    cmdstring = cmdstring & " -" & SysCore & ".special nn4x"
ElseIf Combo4.Text = "nny2x - Nearest-neighbor 2x, y axis only" Then
    cmdstring = cmdstring & " -" & SysCore & ".special nny2x"
ElseIf Combo4.Text = "nny3x - Nearest-neighbor 3x, y axis only" Then
    cmdstring = cmdstring & " -" & SysCore & ".special nny3x"
ElseIf Combo4.Text = "nny4x - Nearest-neighbor 4x, y axis only" Then
    cmdstring = cmdstring & " -" & SysCore & ".special nny4x"
End If

'*** v0.1.3
If Combo5.Enabled = True Then
    If Val(Text8.Text) > 1 Then
        For y = 1 To Val(Text8.Text)
            If SysCore = "psx" Then
                If Combo5.ListIndex = 0 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " none"
                ElseIf Combo5.ListIndex = 1 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " gamepad"
                ElseIf Combo5.ListIndex = 2 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " dualshock"
                ElseIf Combo5.ListIndex = 3 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " dualanalog"
                ElseIf Combo5.ListIndex = 4 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " analogjoy"
                ElseIf Combo5.ListIndex = 5 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " mouse"
                ElseIf Combo5.ListIndex = 6 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " negcon"
                ElseIf Combo5.ListIndex = 7 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " guncon"
                ElseIf Combo5.ListIndex = 8 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " justifier"
                ElseIf Combo5.ListIndex = 9 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " dancepad"
                End If
            ElseIf SysCore = "snes" Then
                If Combo5.ListIndex = 0 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " none"
                ElseIf Combo5.ListIndex = 1 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " gamepad"
                ElseIf Combo5.ListIndex = 2 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " mouse"
                ElseIf Combo5.ListIndex = 3 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 gamepad -" & SysCore & ".input.port2 superscope"
                End If
            ElseIf SysCore = "nes" Then
                If Combo5.ListIndex = 0 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " none"
                ElseIf Combo5.ListIndex = 1 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " gamepad"
                ElseIf Combo5.ListIndex = 2 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " zapper"
                ElseIf Combo5.ListIndex = 3 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " powerpada"
                ElseIf Combo5.ListIndex = 4 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " powerpadb"
                ElseIf Combo5.ListIndex = 5 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " arkanoid"
                End If
            ElseIf SysCore = "ss" Then
                If Combo5.ListIndex = 0 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " none"
                ElseIf Combo5.ListIndex = 1 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " gamepad"
                ElseIf Combo5.ListIndex = 2 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " 3dpad"
                ElseIf Combo5.ListIndex = 3 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & " mouse"
                End If
            End If
        Next y
    ElseIf Val(Text8.Text) = 1 Then
            If SysCore = "psx" Then
                If Combo5.ListIndex = 0 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 none"
                ElseIf Combo5.ListIndex = 1 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 gamepad"
                ElseIf Combo5.ListIndex = 2 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 dualshock"
                ElseIf Combo5.ListIndex = 3 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 dualanalog"
                ElseIf Combo5.ListIndex = 4 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 analogjoy"
                ElseIf Combo5.ListIndex = 5 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 mouse"
                ElseIf Combo5.ListIndex = 6 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 negcon"
                ElseIf Combo5.ListIndex = 7 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 guncon"
                ElseIf Combo5.ListIndex = 8 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 justifier"
                ElseIf Combo5.ListIndex = 9 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 dancepad"
                End If
            ElseIf SysCore = "snes" Then
                If Combo5.ListIndex = 0 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 none"
                ElseIf Combo5.ListIndex = 1 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 gamepad"
                ElseIf Combo5.ListIndex = 2 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 mouse"
                ElseIf Combo5.ListIndex = 3 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 gamepad -" & SysCore & ".input.port2 superscope"
                End If
            ElseIf SysCore = "nes" Then
                If Combo5.ListIndex = 0 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 none"
                ElseIf Combo5.ListIndex = 1 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 gamepad"
                ElseIf Combo5.ListIndex = 2 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 zapper"
                ElseIf Combo5.ListIndex = 3 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 powerpada"
                ElseIf Combo5.ListIndex = 4 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 powerpadb"
                ElseIf Combo5.ListIndex = 5 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 arkanoid"
                End If
            ElseIf SysCore = "ss" Then
                If Combo5.ListIndex = 0 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 none"
                ElseIf Combo5.ListIndex = 1 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 gamepad"
                ElseIf Combo5.ListIndex = 2 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 3dpad"
                ElseIf Combo5.ListIndex = 3 Then
                    cmdstring = cmdstring & " -" & SysCore & ".input.port1 mouse"
                End If
            End If
    End If
'none
'gamepad
'zapper
'powerpada
'powerpadb
'arkanoid
    If Val(Text11.Text) <> "1.00" Then
        If Val(Text8.Text) > 1 Then
            For y = 1 To Val(Text8.Text)
                cmdstring = cmdstring & " -" & SysCore & ".input.port" & y & ".analogjoy.axis_scale " & Val(Text11.Text) & " -" & SysCore & ".input.port" & y & ".dualanalog.axis_scale " & Val(Text11.Text) & " -" & SysCore & ".input.port" & y & ".dualshock.axis_scale " & Val(Text11.Text)
            Next y
        ElseIf Val(Text8.Text) = 1 Then
            cmdstring = cmdstring & " -" & SysCore & ".input.port1.analogjoy.axis_scale " & Val(Text11.Text) & " -" & SysCore & ".input.port1.dualanalog.axis_scale " & Val(Text11.Text) & " -" & SysCore & ".input.port1.dualshock.axis_scale " & Val(Text11.Text)
        End If
    End If
End If


'Combo6.AddItem "OpenGL - OpenGL + SDL", 0
'Combo6.AddItem "SDL - SDL Surface", 1
'Combo6.AddItem "Overlay - SDL Overlay", 2
'**v0.1.7
If Combo6.Text = "OpenGL - OpenGL + SDL" Then
    cmdstring = cmdstring & " -video.driver opengl"
ElseIf Combo6.Text = "SDL - SDL Surface" Then
    cmdstring = cmdstring & " -video.driver sdl"
ElseIf Combo6.Text = "Overlay - SDL Overlay" Then
    cmdstring = cmdstring & " -video.driver overlay"
End If

If Check19.Value = 1 Then
    cmdstring = cmdstring & " -video.glvsync 1"
ElseIf Check19.Value = 0 Then
    cmdstring = cmdstring & " -video.glvsync 0"
End If

If Check20.Value = 1 Then
    cmdstring = cmdstring & " -" & SysCore & ".forcemono 1"
'ElseIf Check20.Value = 0 Then
'    cmdstring = cmdstring & " -" & SYSCORE & ".forcemono 0"
End If

If Check21.Value = 1 Then
    cmdstring = cmdstring & " -sound 0"
ElseIf Check21.Value = 0 Then
    cmdstring = cmdstring & " -sound 1"
End If

If Check22.Value = 1 Then
    cmdstring = cmdstring & " -video.blit_timesync 1"
ElseIf Check22.Value = 0 Then
    cmdstring = cmdstring & " -video.blit_timesync 0"
End If

'v0.1.9
If Check23.Value = 1 Then
    cmdstring = cmdstring & " -cd.image_memcache 1"
'ElseIf Check23.Value = 0 Then
'    cmdstring = cmdstring & " -cd.image_memcache 0"
End If

'**


If Check3.Value = 1 Then
    cmdstring = cmdstring & " -" & SysCore & ".tblur 1"
End If

If Check4.Value = 1 Then
    cmdstring = cmdstring & " -" & SysCore & ".tblur.accum 1"
    Sleep (100)
    cmdstring = cmdstring & " -" & SysCore & ".tblur.accum.amount " & Text3.Text
End If

If Check5.Value = 1 Then
    cmdstring = cmdstring & " -" & SysCore & ".videoip 1"
Else
    cmdstring = cmdstring & " -" & SysCore & ".videoip 0"
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

'video.glvsync
If Check19.Value = 1 Then
    cmdtring = cmdstrig & "-video.glvsync 1"
Else
    cmdtring = cmdstrig & "-video.glvsync 0"
End If

If SysCore = "psx" Or SysCore = "ss" Then
    If Check1.Value = 1 Then
        cmdstring = cmdstring & " -" & SysCore & ".bios_sanity 1"
    End If
    
    If Check2.Value = 1 Then
        cmdstring = cmdstring & " -" & SysCore & ".cd_sanity 1"
    End If
End If

If Check18.Value = 1 Then
    cmdstring = cmdstring & " -video.deinterlacer weave"
ElseIf Check17.Value = 1 Then
    cmdstring = cmdstring & " -video.deinterlacer bob"
ElseIf Check16.Value = 1 Then
    cmdstring = cmdstring & " -video.deinterlacer bob_offset"
End If


If Val(Text4.Text) > 1 Then
    cmdstring = cmdstring & " -" & SysCore & ".xscale " & Text4.Text & " -" & SysCore & ".yscale " & Text4.Text & " -" & SysCore & ".xscalefs " & Text4.Text & " -" & SysCore & ".yscalefs " & Text4.Text
End If

If Len(Text5.Text) > 0 Then
    cmdstring = cmdstring & " -" & SysCore & ".xres " & Text5.Text
End If

If Len(Text6.Text) > 0 Then
    cmdstring = cmdstring & " -" & SysCore & ".yres " & Text6.Text
End If

If Len(Text7.Text) > 0 Then
    cmdstring = cmdstring & " -filesys.path_state " & Chr(34) & Text7.Text & Chr(34) & " -filesys.path_sav " & Chr(34) & Text7.Text & Chr(34)
End If

If Check14.Value = 0 Then
    cmdstring = cmdstring & " -filesys.untrusted_fip_check 0"
End If

If Len(Text9.Text) > 1 Then
cmdstring = cmdstring & " " & Text9.Text
End If

If Val(Text10.Text) <> 0 Then
    cmdstring = cmdstring & " -" & SysCore & ".scanlines " & Val(Text10.Text)
Else
    cmdstring = cmdstring & " -" & SysCore & ".scanlines 0"
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
ROMFILE = ""
BIOSFILE = ""
End Sub

Private Sub Command5_Click()
Form1.Width = 12945
ActiveFile = "SAVE"
If Len(Text7.Text) >= 1 Then
    If Text7.Text = "Not Set" Then
        ResetSave = vbYes
    Else
        If Check15.Value = 1 Then
            a = a
            ResetSave = vbYes
        Else
            ResetSave = MsgBox("Reset Save Path?", vbYesNo, "Reset Save Path?")
        End If
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
        If Check15.Value = 1 Then
            a = a
            tmp2 = vbYes
        Else
            tmp2 = MsgBox("Set Path: " & Dir1.Path, vbYesNo, "Set this path?")
        End If
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
        If Check15.Value = 1 Then
            a = a
            tmp2 = vbYes
        Else
            tmp2 = MsgBox("Set File: " & File1.Filename, vbYesNo, "Set this file?")
        End If
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
BasicModeFolder = Mid$(tmp3(34), 15, Len(tmp3(34)))
End If
'End Load Settings
End Function
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
'Social Media Icons from Rogie King, http://rog.ie/blog/free-social-media-icons
'"This icon set is 100% free under the WTFPL — no link backs or anything needed. All I ask is that you check out my other efforts, Fine Goods and NeonMob."
'You can has link backs.

Build = "0.2.7"
Form1.Caption = "MedAdvCFG v" & Build & " (Mednafen v0.9.x.x Frontend) by Nigel Todman"
Label34.Caption = "MedAdvCFG v" & Build
Dir1.Path = VB.App.Path
File1.Path = VB.App.Path

Label2.Caption = "Not Set"
'MedEXE = "C:\EMU\mednafen-0.9.38.7-win64\mednafen.exe"

Label16.Caption = "Hotkeys:"

a = Form1.LoadSettings()

Text1.Text = BIOSFILE
Text2.Text = ROMFILE
Text4.Text = ScaleFactor
Text5.Text = XRes
Text6.Text = YRes
Text7.Text = SavePath
Text8.Text = numplayers

Dir1.Path = LastPath
File1.Path = LastPath

ROMDIR = ROMPathLoad
BIOSPATH = BiosPathLoad

Combo1.Text = SystemCore
Combo2.Text = Stretch
Combo3.Text = PixelShader
Combo4.Text = VideoScaler
Combo6.Text = VideoDriver

If BIOSSanity = 1 Then
    Check1.Value = 1
End If

If ROMSanity = 1 Then
    Check2.Value = 1
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

If Fullscreen = 1 Then
    Check6.Value = 1
End If

If Frameskip = 1 Then
    Check7.Value = 1
End If

If SystemRegion = none Then
    a = a
ElseIf SystemRegion = "NTSC-U" Then
    Check11.Value = 1
ElseIf SystemRegion = "NTSC-J" Then
    Check12.Value = 1
ElseIf SystemRegion = "PAL" Then
    Check13.Value = 1
End If

If untrusted_fip_check = 1 Then
    Check14.Value = 1
ElseIf untrusted_fip_check = 0 Then
    Check14.Value = 0
End If

If video_glvsync = 1 Then
    Check19.Value = 1
ElseIf video_glvsync = 0 Then
    Check19.Value = 0
End If

If ForceMono = 1 Then
    Check20.Value = 1
ElseIf ForceMono = 0 Then
    Check20.Value = 0
End If

If DisableSound = 1 Then
    Check21.Value = 1
ElseIf DisableSound = 0 Then
    Check21.Value = 0
End If

If video_blit_timesync = 1 Then
    Check22.Value = 1
ElseIf video_blit_timesync = 0 Then
    Check22.Value = 0
End If

If cd_image_memcache = 1 Then
    Check23.Value = 1
ElseIf cd_image_memcache = 0 Then
    Check23.Value = 0
End If

If Combo1.Text = "psx (Sony PlayStation)" Then
    Check1.Enabled = True
    Check2.Enabled = True
    Check11.Enabled = True
    Check12.Enabled = True
    Check13.Enabled = True
    Combo5.Enabled = True
    Check1.Value = 1
    Check2.Value = 1
    Check9.Value = 1
    Check10.Value = 1
    Combo5.AddItem "none", 0
    Combo5.AddItem "gamepad - SCPH-1080 PlayStation Digital Gamepad", 1
    Combo5.AddItem "dualshock - SCPH-1200 PlayStation DualShock Gamepad", 2
    Combo5.AddItem "dualanalog - SCPH-1180 PlayStation DualAnalog Gamepad", 3
    Combo5.AddItem "analogjoy - SCPH-1110 PlayStation Analog Joystick", 4
    Combo5.AddItem "mouse - SCPH-1090 PlayStation Mouse", 5
    Combo5.AddItem "negcon - NPC-101 Namco neGcon", 6
    Combo5.AddItem "guncon - NPC-103 Namco GunCon", 7
    Combo5.AddItem "justifier - SLUH-00017 Konami Justifier", 8
    Combo5.AddItem "dancepad - SLUH-00071 Konami Dancepad", 9
    Combo5.ListIndex = 1
ElseIf Combo1.Text = "snes (Super Nintendo Entertainment System)" Then
    Combo5.AddItem "none", 0
    Combo5.AddItem "gamepad - SNS-005 Super Nintendo Controller", 1
    Combo5.AddItem "mouse - SNS-016 Super Nintendo Mouse Controller", 2
    Combo5.AddItem "superscope - SNS-013 Super Nintendo Super Scope", 3
    Combo5.ListIndex = 1
ElseIf Combo1.Text = "ss (Sega Saturn)" Then
    Check1.Enabled = True
    Check2.Enabled = True
    Check11.Enabled = True
    Check12.Enabled = True
    Check13.Enabled = True
    Combo5.Enabled = True
    Check1.Value = 1
    Check2.Value = 1
    Check9.Value = 1
    Check10.Value = 1
    'Label29.Visible = True
    Combo5.AddItem "none", 0
    Combo5.AddItem "gamepad - MK-80100 Sega Saturn Controller", 1
    Combo5.AddItem "3dpad - MK-80117 Sega Saturn 3D Control Pad", 2
    Combo5.AddItem "mouse - HSS-0139 Sega Saturn Shuttle Mouse", 3
    Combo5.ListIndex = 1
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
    Combo5.Enabled = False
End If

a = Validate_MedEXE()
a = Validate_Rom()
a = Validate_Bios()

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
Combo1.AddItem "ss (Sega Saturn)", 13
Combo1.AddItem "vb (Virtual Boy)", 14
Combo1.AddItem "wswan (WonderSwan)", 15

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

Combo6.AddItem "OpenGL - OpenGL + SDL", 0
Combo6.AddItem "SDL - SDL Surface", 1
Combo6.AddItem "Overlay - SDL Overlay", 2
Combo6.ListIndex = 0

'SySCore = PSX
'none
'gamepad
'dualshock
'dualanalog
'analogjoy
'mouse
'negcon
'guncon
'justifier
'dancepad

'SysCore = SNES
'none
'gamepad
'mouse
'superscope - port2 only

'SysCore = NES
'none
'gamepad
'zapper
'powerpada
'powerpadb
'arkanoid

If FSO.FileExists(MedEXE) = False Then
    Form1.Width = 12900
    ActiveFile = "MEDEXE"
    MsgBox "Select your Mednafen EXE to get started!"
End If


End Sub

Private Sub Gen_M3U_Click()
M3USize = InputBox("How many discs total?", "How many discs total?")
If M3USize <= 1 Then
    MsgBox "Mutli-Disc M3U not needed for a single disc game. If you want to load a single disc game after Gameshark. Enter 2, First is GameShark CD, Second is Game CD"
Else
Generate_M3U (M3USize)
End If
End Sub

Private Sub Image1_Click()
Shell ("cmd.exe /c start http://mednafen.fobby.net/"), vbHide
End Sub

Private Sub Label15_Click()
Shell ("cmd.exe /c start http://www.NigelTodman.com"), vbHide
End Sub

Private Sub Image2_Click()
Shell ("cmd.exe /c start http://www.facebook.com/nigel.todman.3"), vbHide
End Sub

Private Sub Image3_Click()
Shell ("cmd.exe /c start http://www.twitter.com/Veritas_83"), vbHide
End Sub

Private Sub Image4_Click()
Shell ("cmd.exe /c start http://steamcommunity.com/id/veritas_/"), vbHide
End Sub

Private Sub Image5_Click()
Shell ("cmd.exe /c start https://www.youtube.com/user/Veritas0923/videos"), vbHide
End Sub

Private Sub Image6_Click()
Shell ("cmd.exe /c start http://www.nigeltodman.com/"), vbHide
End Sub

Private Sub Image7_Click()
Shell ("cmd.exe /c start https://github.com/Veritas83"), vbHide
End Sub

Private Sub Label2_Click()
Clipboard.Clear
Clipboard.SetText Label2.Caption
MsgBox "MD5 copied to Clipboard"
End Sub

Private Sub Label35_Click()

End Sub

Private Sub Label34_Click()
Shell ("cmd.exe /c start http://www.nigeltodman.com/2016/06/05/medadvcfg-v0-0-1-mednafen-v0-9-38-x-frontend/"), vbHide
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
Unload Form2
Unload Form3
Unload Form4
End Sub

Private Sub Reset_Settings_Click()
tmp = 0
tmp = MsgBox("Are you sure you want to reset settings?", vbYesNo)
If tmp = vbYes Then
Shell ("cmd.exe /c del " & Chr(34) & VB.App.Path & "\MedAdvCFG.dat" & Chr(34))
Sleep (750)
Shell ("cmd.exe /c " & Chr(34) & VB.App.Path & "\MedAdvCFG.exe" & Chr(34))
Unload Form1
End If
End Sub

Private Sub Save_Settings_Click()
Close #6
Open VB.App.Path & "\MedAdvCFG.dat" For Output As #6
    Print #6, "MedEXE=" & MedEXE
    Print #6, "SystemCore=" & Combo1.Text
    If FSO.FileExists(BIOSPATH & "\" & BIOSFILE) Then
        Print #6, "SystemBIOS=" & BIOSPATH & "\" & BIOSFILE
    ElseIf FSO.FileExists(BIOSFILE) Then
        Print #6, "SystemBIOS=" & BIOSFILE
    End If
    Print #6, "BIOSSanity=" & Check1.Value
    Print #6, "RomImage=" & ROMFILE
    Print #6, "ROMSanity=" & Check2.Value
    Print #6, "Stretch=" & Combo2.Text
    Print #6, "PixelShader=" & Combo3.Text
    Print #6, "VideoScaler=" & Combo4.Text
    Print #6, "Fullscreen=" & Check6.Value
    Print #6, "Frameskip=" & Check7.Value
    Print #6, "Tblur=" & Check3.Value
    Print #6, "TblurAccum=" & Check4.Value
    Print #6, "AccumAmount=" & Text3.Text
    Print #6, "VideoIP=" & Check5.Value
    Print #6, "XRes=" & Text5.Text
    Print #6, "YRes=" & Text6.Text
    Print #6, "ScaleFactor=" & Text4.Text
    Print #6, "LastPath=" & File1.Path
    Print #6, "BiosPathLoad=" & BIOSPATH
    Print #6, "SavePath=" & Text7.Text
    If Check11.Value = 0 And Check12.Value = 0 And Check13.Value = 0 Then
        Print #6, "SystemRegion=None"
    ElseIf Check11.Value = 1 Then
        Print #6, "SystemRegion=NTSC-U"
    ElseIf Check12.Value = 1 Then
        Print #6, "SystemRegion=NTSC-J"
    ElseIf Check13.Value = 1 Then
        Print #6, "SystemRegion=PAL"
    End If
    Print #6, "RomPath=" & ROMDIR
    Print #6, "DisableSound=" & Check21.Value
    Print #6, "ForceMono=" & Check20.Value
    Print #6, "video.blit_timesync=" & Check22.Value
    Print #6, "video.glvsync=" & Check19.Value
    Print #6, "untrusted_fip_check=" & Check14.Value
    Print #6, "cd.image_memcache=" & Check23.Value
    Print #6, "scanlines=" & Text10.Text
    Print #6, "axisscale=" & Text11.Text
    Print #6, "numplayers=" & Text8.Text
    Print #6, "customparams=" & Text9.Text
Close #6
End Sub


