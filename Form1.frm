VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "MedAdvCFG v0.0.0 (Mednafen v0.9.x.x Frontend) by Nigel Todman [ADV MODE]"
   ClientHeight    =   8160
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8955
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
   ScaleHeight     =   8160
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo8 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "Form1.frx":851A
      Left            =   1320
      List            =   "Form1.frx":851C
      TabIndex        =   91
      Text            =   "[US] netplay.fobby.net:4046"
      Top             =   5040
      Width           =   4575
   End
   Begin VB.ComboBox Combo7 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "Form1.frx":851E
      Left            =   7720
      List            =   "Form1.frx":8520
      TabIndex        =   90
      Text            =   "800x600"
      Top             =   4800
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   945
      Left            =   0
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   7200
      Width           =   7035
      ExtentX         =   12409
      ExtentY         =   1676
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CheckBox Check23 
      BackColor       =   &H00C0C0C0&
      Caption         =   "cd.image_memcache"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   81
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CheckBox Check22 
      BackColor       =   &H00C0C0C0&
      Caption         =   "video.blit_timesync"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   80
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CheckBox Check21 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Disable Sound"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   79
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CheckBox Check20 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Force Mono"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   78
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox Check19 
      BackColor       =   &H00C0C0C0&
      Caption         =   "video.glvsync"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   77
      Top             =   4080
      Width           =   1455
   End
   Begin VB.ComboBox Combo6 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   75
      Text            =   "None - None/Disabled"
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8520
      TabIndex        =   71
      Text            =   "1.00"
      Top             =   5910
      Width           =   405
   End
   Begin VB.CheckBox Check18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "weave"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   69
      Top             =   3000
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Check17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "bob"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2205
      TabIndex        =   68
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox Check16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "bob_offset"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3045
      TabIndex        =   67
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   65
      Text            =   "0"
      Top             =   5640
      Width           =   285
   End
   Begin VB.CheckBox Check15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Yes to All Prompts(!)"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   64
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   62
      Top             =   5400
      Width           =   4575
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   60
      Text            =   "1"
      Top             =   6200
      Width           =   285
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "Form1.frx":8522
      Left            =   1320
      List            =   "Form1.frx":8524
      TabIndex        =   58
      Text            =   "gamepad - SCPH-1080 PlayStation Digital Gamepad"
      Top             =   4680
      Width           =   4575
   End
   Begin VB.CheckBox Check14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "untrusted_fip_check"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   57
      Top             =   4320
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PAL"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3050
      TabIndex        =   54
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox Check12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NTSC-J"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2200
      TabIndex        =   53
      Top             =   720
      Width           =   975
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NTSC-U"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   52
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   50
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
      Height          =   255
      Left            =   5880
      TabIndex        =   49
      Top             =   4340
      Width           =   495
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BIOS MD5 Check"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   48
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM MD5 Check"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   47
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
      TabIndex        =   45
      Text            =   "1080"
      Top             =   4800
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
      TabIndex        =   44
      Text            =   "1920"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   42
      Text            =   "2"
      Top             =   5080
      Width           =   285
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   41
      Top             =   380
      Width           =   495
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
      TabIndex        =   40
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
      Height          =   3795
      Left            =   9000
      TabIndex        =   39
      Top             =   3960
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
      Height          =   3465
      Left            =   9000
      TabIndex        =   38
      Top             =   450
      Width           =   3615
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Review Command Line before execution?"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   37
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frameskip"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
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
         Name            =   "OpenSymbol"
         Size            =   9
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
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   27
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
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
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   24
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
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
      Top             =   5380
      Width           =   285
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Accumulate color"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   21
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temporal Blur"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
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
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   18
      Text            =   "None - None/Disabled"
      Top             =   3600
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   16
      Text            =   "0 - Disabled"
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM Sanity Check"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
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
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   1820
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
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
      Top             =   1800
      Width           =   4575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BIOS Sanity Check"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
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
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   980
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
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
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Text            =   "gb (GameBoy (Color))"
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "REDUMP: unverified!"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   180
      Left            =   5160
      TabIndex        =   92
      Top             =   1320
      Width           =   1725
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "18j2Env7QokhGG5MccS3LPBKnjsko6u4NQ"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00009900&
      Height          =   180
      Left            =   3240
      TabIndex        =   89
      Top             =   6555
      Width           =   2760
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "www.CoversDB.org"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7200
      TabIndex        =   88
      Top             =   7920
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "REDUMP: unverified!"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   180
      Left            =   5160
      TabIndex        =   87
      Top             =   2160
      Width           =   1725
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "BIOS VER:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   85
      Top             =   1560
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "PSX ID:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   360
      TabIndex        =   84
      Top             =   2400
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Not Set"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1320
      TabIndex        =   83
      Top             =   2400
      Width           =   630
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Netplay Host:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   82
      Top             =   5040
      Width           =   915
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Video Driver"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   76
      Top             =   3240
      Width           =   840
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "www.NigelTodman.com"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7200
      TabIndex        =   74
      Top             =   7680
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MedAdvCFG v0.0.0"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   7200
      TabIndex        =   73
      Top             =   7320
      Width           =   1350
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Axis Scale:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6840
      TabIndex        =   72
      Top             =   6000
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Deinterlacer:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   70
      Top             =   3000
      Width           =   795
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scanlines:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6840
      TabIndex        =   66
      Top             =   5640
      Width           =   645
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   4680
      Picture         =   "Form1.frx":8526
      Top             =   5940
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   4200
      Picture         =   "Form1.frx":91F0
      Top             =   5940
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   5640
      Picture         =   "Form1.frx":9EBA
      Top             =   5940
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   5160
      Picture         =   "Form1.frx":AB84
      Top             =   5940
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3720
      Picture         =   "Form1.frx":B84E
      Top             =   5940
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3240
      Picture         =   "Form1.frx":C518
      Top             =   5940
      Width           =   480
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Custom Params:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   63
      Top             =   5400
      Width           =   1110
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Number of Players:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6840
      TabIndex        =   61
      Top             =   6240
      Width           =   1305
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Controller:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   59
      Top             =   4680
      Width           =   645
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Not Set"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1320
      TabIndex        =   56
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "System Region:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   55
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save Path:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   51
      Top             =   4320
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1200
      Picture         =   "Form1.frx":D1E2
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1665
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Resolution"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6840
      TabIndex        =   46
      Top             =   4800
      Width           =   660
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scaling Factor"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6840
      TabIndex        =   43
      Top             =   5100
      Width           =   945
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "F5:Save state"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   35
      Top             =   6240
      Width           =   1140
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "F7:Load state"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   34
      Top             =   6480
      Width           =   1110
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "F10:Reset Console"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   33
      Top             =   6720
      Width           =   1515
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Alt+Enter:Toggle fullscreen mode"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   32
      Top             =   6960
      Width           =   2715
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "CTRL+SHIFT+1:Select Controller Type"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   3240
      TabIndex        =   31
      Top             =   6765
      Width           =   3180
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALT+SHIFT+1:Configure buttons for Controller Port "
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   3240
      TabIndex        =   30
      Top             =   6960
      Width           =   4365
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hotkeys:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   29
      Top             =   6000
      Width           =   690
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Video Scaler"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   26
      Top             =   3960
      Width           =   840
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temporal Blur amount:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6840
      TabIndex        =   23
      Top             =   5385
      Width           =   1500
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pixel Shader"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   825
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "FS Stretch"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Not Set"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1320
      TabIndex        =   14
      Top             =   2160
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM MD5:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "ROM Image:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Not Set"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1320
      TabIndex        =   8
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "BIOS MD5:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "System BIOS:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "System Core:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "0.9.41.0-win64 Detected! MD5: 6AADC9A8A196DA610E6DB43367B339B4"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   5880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mednafen EXE:"
      BeginProperty Font 
         Name            =   "OpenSymbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1035
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
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu doc 
         Caption         =   "Documentation"
      End
      Begin VB.Menu hlp_netplay 
         Caption         =   "NetPlay Tips"
      End
   End
   Begin VB.Menu Chat 
      Caption         =   "Chat"
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
Dim CoverName, PSXIDList, PSXID, RedumpList, REDUMPMD5, ROMMD5, CUEFILE, BINFILE
Private Type MD5_CTX
  i(1 To 2) As Long
  buf(1 To 4) As Long
  inp(1 To 64) As Byte
  digest(1 To 16) As Byte
End Type

Private Declare Sub MD5Init Lib "cryptdll" (Context As MD5_CTX)
Private Declare Sub MD5Update Lib "cryptdll" (Context As MD5_CTX, ByVal strInput As String, ByVal lLen As Long)
Private Declare Sub MD5Final Lib "cryptdll" (Context As MD5_CTX)
Private Declare Function GetShortPathName Lib "kernel32" _
   Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
   ByVal lpszShortPath As String, ByVal cchBuffer As Long) _
   As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function ShortPath(ByVal strFilename As String) As String
   
   
    Dim strBuffer As String * 255
    Dim lngReturnCode As Long
    
    'FILENAME MUST EXIST FOR API FUNCTION TO WORK
    'SO CREATE THE FILE IF IT DOESN'T EXISTS
    Dim iFileNumber As Integer
    iFileNumber = FreeFile
    If Dir(strFilename) = "" Then
        On Error Resume Next
        Open strFilename For Output As #iFileNumber
        Close #iFileNumber
    End If
    lngReturnCode = GetShortPathName(strFilename, strBuffer, 255)
    ShortPath = Left$(strBuffer, lngReturnCode)
End Function
Public Function CalcMD5(strFilename As String) As String
    Dim strBuffer As String
    Dim myContext As MD5_CTX
    Dim result As String
    Dim lp As Long
    Dim MD5 As String

    strBuffer = Space(FileLen(strFilename))

    Open strFilename For Binary Access Read As #1
        Get #1, , strBuffer
    Close #1

    MD5Init myContext
    MD5Update myContext, strBuffer, Len(strBuffer)
    MD5Final myContext

    result = StrConv(myContext.digest, vbUnicode)
    
    For lp = 1 To Len(result)
            CalcMD5 = CalcMD5 & Right("00" & Hex(Asc(Mid(result, lp, 1))), 2)
    Next
    
End Function
Private Function Generate_M3U(M3USize As Integer)
'*** v0.1.3
Form1.Width = 12945
ActiveFile = "M3U"
MsgBox "Select the first disc"
z = 0
Open VB.App.Path & "\multi.m3u" For Output As #2

End Function
Function FileNameCleanup()
tmparray = Split(ROMFILE, "\")
'MsgBox UBound(tmparray)
tmp = Replace(tmparray(UBound(tmparray)), ".cue", "")
tmp = Replace(tmp, ".smc", "")
tmp = Replace(tmp, ".nes", "")
tmp = Replace(tmp, ".iso", "")
tmp = Replace(tmp, ".bin", "")
tmp = Replace(tmp, ".gg", "")
tmp = Replace(tmp, ".lnx", "")
tmp = Replace(tmp, " (USA) ", "")
tmp = Replace(tmp, " (USA)", "")
tmp = Replace(tmp, "(USA)", "")
For z = 0 To 9
    tmp = Replace(tmp, " (v1." & z & ")", "")
    tmp = Replace(tmp, " (V1." & z & ")", "")
    tmp = Replace(tmp, "(v1." & z & ")", "")
    tmp = Replace(tmp, "(V1." & z & ")", "")
Next z
tmp = Replace(tmp, "(Arcade)", "")
tmp = Replace(tmp, "(Arcade Mode)", "")
tmp = Replace(tmp, "(Arcade Disc)", "")
tmp = Replace(tmp, "(Evolution Disc)", "")
tmp = Replace(tmp, "(Simulation)", "")
tmp = Replace(tmp, "(Simulation Mode)", "")
tmp = Replace(tmp, "(Unl)", "")
tmp = Replace(tmp, "(Disc 1)", "")
tmp = Replace(tmp, "(Disc 2)", "")
tmp = Replace(tmp, "(Disc 3)", "")
tmp = Replace(tmp, "(Disc 4)", "")
tmp = Replace(tmp, "(Disc 5)", "")
tmp = Replace(tmp, "(En,Fr,De,Es,It,Nl,Sv,No,Da,Fi)", "")
tmp = Replace(tmp, "(En,Fr,De,Es,It,Nl,Sv,Da)", "")
tmp = Replace(tmp, "(En,Fr,De,Es,It,Sv)", "")
tmp = Replace(tmp, "(En,Fr,De,Es,It)", "")
tmp = Replace(tmp, "(En,Ja,Fr,De)", "")
tmp = Replace(tmp, "(En,Fr,Es,Pt)", "")
tmp = Replace(tmp, "(En,Fr,De,Es)", "")
tmp = Replace(tmp, "(En,Fr,De,Sv)", "")
tmp = Replace(tmp, "(En,Fr,De)", "")
tmp = Replace(tmp, "(En,Fr,Es)", "")
tmp = Replace(tmp, "(En,Fr)", "")
tmp = Replace(tmp, "(En,Ja)", "")
tmp = Replace(tmp, "(Soviet)", "")
tmp = Replace(tmp, "(Allies)", "")
tmp = Replace(tmp, "(Namco)", "")
tmp = Replace(tmp, "(Tengen)", "")
tmp = Replace(tmp, "(Hack)", "")
tmp = Replace(tmp, "(Beta)", "")
tmp = Replace(tmp, "(Leon)", "")
tmp = Replace(tmp, "(Claire)", "")
tmp = Replace(tmp, "(The Making of)", "")
tmp = Replace(tmp, "[!]", "")
tmp = Replace(tmp, "!!!!", "!")
tmp = Replace(tmp, "!!!", "!")
tmp = Replace(tmp, "!!", "!")
tmp = Replace(tmp, "(U)", "")
tmp = Replace(tmp, "(J)", "")
tmp = Replace(tmp, "(E)", "")
tmp = Replace(tmp, "(G)", "")
tmp = Replace(tmp, "(W)", "")
tmp = Replace(tmp, "(R)", "")
tmp = Replace(tmp, "(VC)", "")
tmp = Replace(tmp, "(JU)", "")
tmp = Replace(tmp, "(JE)", "")
tmp = Replace(tmp, "(M2)", "")
tmp = Replace(tmp, "(GDI)", "")
tmp = Replace(tmp, "(NOD)", "")
tmp = Replace(tmp, "(ECD)", "")
tmp = Replace(tmp, "(PRG0)", "")
tmp = Replace(tmp, "(PRG1)", "")
tmp = Replace(tmp, "(PC10)", "")
tmp = Replace(tmp, "(REV00)", "")
tmp = Replace(tmp, "(REV01)", "")
tmp = Replace(tmp, "(UE)", "")
tmp = Replace(tmp, "[a1]", "")
tmp = Replace(tmp, "[b1]", "")
tmp = Replace(tmp, "[b1+1C]", "")
tmp = Replace(tmp, "[b1+2C]", "")
tmp = Replace(tmp, "[b2]", "")
tmp = Replace(tmp, "[b3]", "")
tmp = Replace(tmp, "[b4]", "")
tmp = Replace(tmp, "[b5]", "")
tmp = Replace(tmp, "[b6]", "")
tmp = Replace(tmp, "[b7]", "")
tmp = Replace(tmp, "[b8]", "")
tmp = Replace(tmp, "[b9]", "")
tmp = Replace(tmp, "[c]", "")
tmp = Replace(tmp, "[f1]", "")
tmp = Replace(tmp, "[f1+1C]", "")
tmp = Replace(tmp, "[f1+2C]", "")
tmp = Replace(tmp, "[f1+3C]", "")
tmp = Replace(tmp, "[f1+4C]", "")
tmp = Replace(tmp, "[f1+5C]", "")
tmp = Replace(tmp, "[f2]", "")
tmp = Replace(tmp, "[f2+1C]", "")
tmp = Replace(tmp, "[f2+2C]", "")
tmp = Replace(tmp, "[f2+3C]", "")
tmp = Replace(tmp, "[o1]", "")
tmp = Replace(tmp, "[o2]", "")
tmp = Replace(tmp, "[o3]", "")
tmp = Replace(tmp, "[p1]", "")
tmp = Replace(tmp, "[p2]", "")
tmp = Replace(tmp, "[p3]", "")
tmp = Replace(tmp, "[p4]", "")
tmp = Replace(tmp, "[p5]", "")
tmp = Replace(tmp, "[t1]", "")
tmp = Replace(tmp, "[t2]", "")
tmp = Replace(tmp, "[t3]", "")
tmp = Replace(tmp, "[t4]", "")
tmp = Replace(tmp, "[t5]", "")
tmp = Replace(tmp, "[t6]", "")
tmp = Replace(tmp, "[hI]", "")
tmp = Replace(tmp, "[hI+C]", "")
tmp = Replace(tmp, "[h1C]", "")
tmp = Replace(tmp, "[h2C]", "")
tmp = Replace(tmp, "[h3C]", "")
tmp = Replace(tmp, "[h4C]", "")
tmp = Replace(tmp, "[h5C]", "")
tmp = Replace(tmp, "'", "")
tmp = Replace(tmp, ",", "")
tmp = Replace(tmp, ".", "")
tmp = Replace(tmp, " - ", "-")
tmp = Replace(tmp, "  ", " ")
tmp = Replace(tmp, "_-_", "-")
tmp = LTrim(RTrim(tmp))
tmp = Replace(tmp, " ", "_")
tmp = Replace(tmp, "-", "_")
tmp = LCase(tmp)
If Mid$(tmp, Len(tmp) - 3, 4) = "_the" Then
    tmp = "the_" & Mid$(tmp, 1, Len(tmp) - 4)
End If
FileNameCleanup = tmp
End Function
Public Function Validate_Rom()
If Check9.Value = 1 Then
    If FSO.FileExists(ROMFILE) = True Then
        ROMMD5 = CalcMD5(ShortPath(ROMFILE))
        Text2.Text = ROMFILE
        Label9.Caption = "MD5: " & ROMMD5
        tmp = FileNameCleanup()
        Label35.Visible = True
        CoverName = tmp
        SysCore = SetSysCore()
        If SysCore = "psx" Then
            a = Redump(ROMMD5)
            Label35.Caption = "Game ID: " & GetPSXID()
        ElseIf SysCore = "ss" Then
            a = Redump(ROMMD5)
            Label35.Visible = False
        ElseIf SysCore = "snes" Then
            a = Redump(ROMMD5)
            Label35.Visible = False
        ElseIf SysCore = "nes" Then
            a = Redump(ROMMD5)
            Label35.Visible = False
        End If
    End If
Else
    Label9.Caption = "MD5: ROM MD5 Disabled"
End If
Validate_Rom = tmp
End Function
Function Validate_Bios()
If Check10.Value = 1 Then
    If FSO.FileExists(BIOSPATH & "\" & BIOSFILE) = True Then
        tmp = CalcMD5(ShortPath(BIOSPATH & "\" & BIOSFILE))
        Text1.Text = BIOSPATH & "\" & BIOSFILE
    End If
    If FSO.FileExists(BIOSFILE) = True Then
        tmp = CalcMD5(ShortPath(BIOSFILE))
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
    If Label29.Visible = True Then
            Label42.Caption = "REDUMP: verified!"
            Label42.ForeColor = RGB(0, 153, 0)
    Else
            Label42.Caption = "REDUMP: unverified!"
            Label42.ForeColor = RGB(255, 128, 0)
    End If
Else
    Label6.Caption = "MD5: BIOS MD5 Disabled"
End If
Validate_Bios = tmp
End Function
Public Function SetSysCore()
If Form1.Combo1.Text = "gb (GameBoy (Color))" Then
    SysCore = "gb"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "gb (GameBoy (Color))" & vbCrLf
ElseIf Form1.Combo1.Text = "gba (GameBoy Advanced)" Then
    SysCore = "gba"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "gba (GameBoy Advanced)" & vbCrLf
ElseIf Form1.Combo1.Text = "gg (Sega Game Gear)" Then
    SysCore = "gg"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "gg (Sega Game Gear)" & vbCrLf
ElseIf Form1.Combo1.Text = "lynx (Atari Lynx)" Then
    SysCore = "lynx"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "lynx (Atari Lynx)" & vbCrLf
ElseIf Form1.Combo1.Text = "md (Sega Genesis/MegaDrive)" Then
    SysCore = "md"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "md (Sega Genesis/MegaDrive)" & vbCrLf
ElseIf Form1.Combo1.Text = "nes (Nintendo Entertainment System)" Then
    SysCore = "nes"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "nes (Nintendo Entertainment System)" & vbCrLf
ElseIf Form1.Combo1.Text = "ngp (Neo Geo Pocket (Color))" Then
    SysCore = "ngp"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "ngp (Neo Geo Pocket (Color))" & vbCrLf
ElseIf Form1.Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SysCore = "pce"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" & vbCrLf
ElseIf Form1.Combo1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SysCore = "pce_fast"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" & vbCrLf
ElseIf Form1.Combo1.Text = "pcfx (PC-FX)" Then
    SysCore = "pcfx"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "pcfx (PC-FX)" & vbCrLf
ElseIf Form1.Combo1.Text = "psx (Sony PlayStation)" Then
    SysCore = "psx"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "psx (Sony PlayStation)" & vbCrLf
ElseIf Form1.Combo1.Text = "sms (Sega Master System)" Then
    SysCore = "sms"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "sms (Sega Master System)" & vbCrLf
ElseIf Form1.Combo1.Text = "snes (Super Nintendo Entertainment System)" Then
    SysCore = "snes"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "snes (Super Nintendo Entertainment System)" & vbCrLf
ElseIf Form1.Combo1.Text = "ss (Sega Saturn)" Then
    SysCore = "ss"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "ss (Sega Saturn)" & vbCrLf
ElseIf Form1.Combo1.Text = "vb (Virtual Boy)" Then
    SysCore = "vb"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "vb (Virtual Boy)" & vbCrLf
ElseIf Form1.Combo1.Text = "wswan (WonderSwan)" Then
    SysCore = "wswan"
    If Len(Form3.Text1.Text) <= 1 Then Form3.Text1.Text = "wswan (WonderSwan)" & vbCrLf
End If
SetSysCore = SysCore
End Function
Public Function Validate_MedEXE()
If FSO.FileExists(MedEXE) = True Then
    tmp = CalcMD5(ShortPath(MedEXE))
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
    Else
        Label2.Caption = "Unknown Mednafen Version! MD5: " & tmp
    End If
End If
Validate_MedEXE = tmp
End Function

Private Sub about_Click()
MsgBox "MedAdvCFG v" & Build & " (Mednafen v0.9.x.x Frontend)" & vbCrLf & "Written by Nigel Todman (nigel@nigeltodman.com)" & vbCrLf & "Primarily written as a PSX Frontend." & vbCrLf & "Tested with the following System Cores:" & vbCrLf & "GB, GBA, GG, MD, NES, PCE, PCE_FAST, PSX, SNES, SS, VB" & vbCrLf & vbCrLf & "Homepage: www.NigelTodman.com" & vbCrLf & "Facebook: facebook.com/nigel.todman.3" & vbCrLf & "Twitter: @Veritas_83" & vbCrLf & "YouTube: Veritas0923" & vbCrLf & "BTC: 18j2Env7QokhGG5MccS3LPBKnjsko6u4NQ"
End Sub

Private Sub Advanced_Click()
Advanced.Checked = True
Basic.Checked = False
Form1.Visible = True
Form2.Visible = False
End Sub

Private Sub Basic_Click()
Advanced.Checked = False
Basic.Checked = True
Form1.Visible = False
Form2.Visible = True
End Sub

Private Sub Chat_Click()
Shell ("cmd.exe /c start http://bit.ly/2k5E1Xq"), vbHide
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
a = CoreControls()
End Sub

Private Sub Combo1_Click()
'MsgBox Combo1.Text

a = CoreControls()

If Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Or Combo1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    MsgBox "BIOS Image File: syscard3.pce is expected"
ElseIf Combo1.Text = "psx (Sony PlayStation)" Then
    MsgBox "BIOS Image File: scph5500.bin/scph5501.bin/scph5502.bin is expected"
ElseIf Combo1.Text = "ss (Sega Saturn)" Then
    MsgBox "BIOS Image File: sega_101.bin/mpr-17933.bin is expected"
End If
End Sub

Private Sub Combo7_Click()
tmparray = Split(Combo7.Text, "x")
Text5.Text = tmparray(0)
Text6.Text = tmparray(1)
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

'v0.3.0 NetPlay
If Combo8.ListIndex >= 1 Then
cmdstring = cmdstring & " -netplay.host " & Chr(34) & Mid$(Combo8.Text, 7, Len(Combo8.Text)) & Chr(34) & " -connect"
End If

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

Private Sub doc_Click()
Shell "cmd.exe /c start https://mednafen.github.io/documentation/", vbHide
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
        If Check15.Value = 1 Then
            a = a
            tmp2 = vbYes
        Else
            tmp2 = MsgBox("Set File: " & File1.FileName, vbYesNo, "Set this file?")
        End If
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
Function LoadSettings()
'Load Settings
Set FSO = CreateObject("Scripting.FileSystemObject")

If FSO.FileExists(VB.App.Path & "\MedAdvCFG.dat") Then

Open VB.App.Path & "\MedAdvCFG.dat" For Input As #1
    For x = 1 To 35
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
QuickLaunch = Mid$(tmp3(35), 13, 1)
End If
'End Load Settings
End Function
Public Function GetBuild()
GetBuild = "0.3.2-r33"
End Function
Public Function ResetSysCore()
SysCore = ""
Combo1.Text = ""
Form3.Text1.Text = ""
ResetSysCore = ""
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
    RedumpList = VB.App.Path & "\dat\nes-all-nointro.dat"
ElseIf SysCore = "snes" Then
    RedumpList = VB.App.Path & "\dat\snes-all-nointro.dat"
ElseIf SysCore = "gg" Then
    RedumpList = VB.App.Path & "\dat\gg-all-nointro.dat"
ElseIf SysCore = "gba" Then
    RedumpList = VB.App.Path & "\dat\gba-all-nointro.dat"
ElseIf SysCore = "gbc" Then
    RedumpList = VB.App.Path & "\dat\gbc-all-nointro.dat"
ElseIf SysCore = "vb" Then
    RedumpList = VB.App.Path & "\dat\vb-all-nointro.dat"
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
                Else
                    Label23.Caption = "NOINTRO: unverified!"
                    Label23.ForeColor = RGB(255, 128, 0)
                    booltmp = True
                End If
    End If
End If
End Function

Public Function GetPSXID()

Set FSO = CreateObject("Scripting.FileSystemObject")

PSXIDList = VB.App.Path & "\dat\psx-usa-id.dat"

If FSO.FileExists(PSXIDList) Then
    Close #12
    Open PSXIDList For Input As #12
    Do
        Line Input #12, tmp
        For x = 1 To Len(tmp)
            If Mid$(tmp, 1, Len(CoverName)) = CoverName And Mid$(tmp, Len(CoverName) + 1, 1) = " " Then
                PSXID = tmp
                'MsgBox "PSXID: " & PSXID
                tmparray = Split(PSXID, " ")
                GetPSXID = tmparray(1)
            End If
        Next x
    Loop Until EOF(12)
    Close #12
End If
End Function

Function CoreControls()
'v0.3.0 Remove trailing \ on Save Path..
If Mid$(Text7.Text, Len(Text7.Text), 1) = " " Then Text7.Text = Mid$(Text7.Text, 1, Len(Text7.Text) - 1)
If Mid$(Text7.Text, Len(Text7.Text), 1) = "\" Then Text7.Text = Mid$(Text7.Text, 1, Len(Text7.Text) - 1)

If Combo1.Text = "psx (Sony PlayStation)" Then
    For z = 1 To Combo5.ListCount
        Combo5.RemoveItem 0
    Next z
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
    Label23.Caption = "REDUMP: unverified!"
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label29.Visible = True
    Label42.Visible = True
    Text1.Visible = True
    Command1.Visible = True
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
    Label23.Caption = "NOINTRO: unverified!"
    Label23.ForeColor = RGB(255, 128, 0)
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label29.Visible = False
    Label42.Visible = False
    Text1.Visible = False
    Command1.Visible = False
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
    Label23.Caption = "NOINTRO: unverified!"
    Label23.ForeColor = RGB(255, 128, 0)
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label29.Visible = False
    Label42.Visible = False
    Text1.Visible = False
    Command1.Visible = False
ElseIf Combo1.Text = "gb (GameBoy (Color))" Then
    Label23.Caption = "NOINTRO: unverified!"
    Label23.ForeColor = RGB(255, 128, 0)
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label29.Visible = False
    Label42.Visible = False
    Text1.Visible = False
    Command1.Visible = False
ElseIf Combo1.Text = "gba (GameBoy Advanced)" Then
    Label23.Caption = "NOINTRO: unverified!"
    Label23.ForeColor = RGB(255, 128, 0)
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label29.Visible = False
    Label42.Visible = False
    Text1.Visible = False
    Command1.Visible = False
ElseIf Combo1.Text = "lynx (Atari Lynx)" Then
    Label23.Caption = "NOINTRO: unverified!"
    Label23.ForeColor = RGB(255, 128, 0)
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label29.Visible = False
    Label42.Visible = False
    Text1.Visible = False
    Command1.Visible = False
ElseIf Combo1.Text = "md (Sega Genesis/MegaDrive)" Then
    Label23.Caption = "NOINTRO: unverified!"
    Label23.ForeColor = RGB(255, 128, 0)
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label29.Visible = False
    Label42.Visible = False
    Text1.Visible = False
    Command1.Visible = False
ElseIf Combo1.Text = "gg (Sega Game Gear)" Then
    Label23.Caption = "NOINTRO: unverified!"
    Label23.ForeColor = RGB(255, 128, 0)
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label29.Visible = False
    Label42.Visible = False
    Text1.Visible = False
    Command1.Visible = False
ElseIf Combo1.Text = "ss (Sega Saturn)" Then
    For z = 1 To Combo5.ListCount
        Combo5.RemoveItem 0
    Next z
    Check1.Enabled = True
    Check2.Enabled = True
    Check11.Enabled = True
    Check12.Enabled = True
    Check13.Enabled = True
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
    Combo5.Enabled = True
    Label23.Caption = "REDUMP: unverified!"
    Label23.ForeColor = RGB(255, 128, 0)
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label29.Visible = True
    Label42.Visible = True
    Text1.Visible = True
    Command1.Visible = True
ElseIf Combo1.Text = "pce (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SysCore = "pce"
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label29.Visible = True
    Label42.Visible = True
    Text1.Visible = True
    Command1.Visible = True
    Label23.Caption = "REDUMP: unverified!"
    Label23.ForeColor = RGB(255, 128, 0)
ElseIf Combo1.Text = "pce_fast (PC Engine (CD)/TurboGrafx 16 (CD)/SuperGrafx)" Then
    SysCore = "pce_fast"
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label29.Visible = True
    Label42.Visible = True
    Text1.Visible = True
    Command1.Visible = True
    Label23.Caption = "REDUMP: unverified!"
    Label23.ForeColor = RGB(255, 128, 0)
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
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label29.Visible = False
    Label42.Visible = False
    Text1.Visible = False
    Command1.Visible = False
End If
End Function

Private Sub Form_Load()
'12945
'9240
FatalError = False
Build = Form1.GetBuild()
Form1.Width = 9240
ActiveFile = "None"

WebBrowser1.StatusBar = False
WebBrowser1.AddressBar = False
WebBrowser1.MenuBar = False
WebBrowser1.Resizable = False
WebBrowser1.ToolBar = False
WebBrowser1.Resizable = False
WebBrowser1.Navigate ("http://ad.a-ads.com/402648?size=468x60&background_color=CCCCCC&title_color=0000ff&link_color=313370&text_color=ffffff&title_hover_color=ff8c00&link_hover_color=ff8c00")

Form1.Caption = "MedAdvCFG v" & Build & " (Mednafen v0.9.x.x Frontend) by Nigel Todman"

Dir1.Path = VB.App.Path
File1.Path = VB.App.Path

Label29.Visible = False
Label35.Visible = False

Label2.Caption = "Not Set"
Label16.Caption = "Hotkeys:"
Label34.Caption = "MedAdvCFG v" & Build

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
Combo7.Text = Text5.Text & "x" & Text6.Text

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

a = CoreControls()

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

Combo7.AddItem "320x200", 0
Combo7.AddItem "320x240", 1
Combo7.AddItem "640x400", 2
Combo7.AddItem "640x480", 3
Combo7.AddItem "800x480", 4
Combo7.AddItem "800x600", 5
Combo7.AddItem "852x480", 6
Combo7.AddItem "1024x768", 7
Combo7.AddItem "1152x864", 8
Combo7.AddItem "1280x720", 9
Combo7.AddItem "1280x768", 10
Combo7.AddItem "1280x800", 11
Combo7.AddItem "1280x960", 12
Combo7.AddItem "1280x1024", 13
Combo7.AddItem "1360x768", 14
Combo7.AddItem "1366x768", 15
Combo7.AddItem "1400x1050", 16
Combo7.AddItem "1440x900", 17
Combo7.AddItem "1600x900", 18
Combo7.AddItem "1600x1200", 19
Combo7.AddItem "1680x1050", 20
Combo7.AddItem "1920x1080", 21
Combo7.AddItem "2048x1536", 22
Combo7.AddItem "2560x2048", 23
Combo7.AddItem "3200x2400", 24
Combo7.AddItem "4000x3000", 25
Combo7.AddItem "5120x4096", 26
Combo7.AddItem "6400x4800", 27

Combo8.AddItem "Disabled/Offline", 0
Combo8.AddItem "[US1] netplay.fobby.net", 1
Combo8.AddItem "[US2] mednafen-us.emuparadise.org", 2
Combo8.AddItem "[NL1] mednafen-nl.emuparadise.org", 3
Combo8.AddItem "[IT1] speedvicio.dtdns.net", 4
Combo8.AddItem "[IT2] scall.org", 5
Combo8.AddItem "[RU1] gs.emu-land.net", 6
Combo8.AddItem "[RU2] emu-russia.net", 7
Combo8.Text = "Disabled/Offline"

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

Private Sub hlp_netplay_Click()
MsgBox "Tips for NetPlay" & vbCrLf & "All players must be using the same:" & vbCrLf & "System BIOS, ROM Image and Mednafen version." & vbCrLf & "Use the provided MD5 to verify all is the same." & vbCrLf & "Select a NetPlay Host from the list. Player 2 must select the same." & vbCrLf & "First player to connect with that game becomes the 'Source/Host' game." & vbCrLf & "When Player 2 connects it will be just like they plugged their controller in."
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

Private Sub Label24_Click()
Shell ("cmd.exe /c start http://www.CoversDB.org"), vbHide
End Sub

Private Sub Label34_Click()
Shell "cmd.exe /c start http://www.NigelTodman.com/medadvcfg", vbHide
End Sub

Private Sub Label35_Click()
Clipboard.Clear
Clipboard.SetText Label35.Caption
MsgBox "Game ID copied to Clipboard"
End Sub

Private Sub Label36_Click()
Shell "cmd.exe /c start http://www.NigelTodman.com/medadvcfg", vbHide
End Sub

Private Sub Label41_Click()
Clipboard.Clear
Clipboard.SetText Label41.Caption
MsgBox "BTC Address copied to Clipboard"
End Sub

Private Sub Label6_Click()
Clipboard.Clear
Clipboard.SetText Mid$(Label6.Caption, 6, 32)
MsgBox "MD5 copied to Clipboard"
End Sub

Private Sub Label9_Click()
Clipboard.Clear
Clipboard.SetText Mid$(Label9.Caption, 6, 32)
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


Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
WebBrowser1.Document.body.Scroll = "no"
End Sub

