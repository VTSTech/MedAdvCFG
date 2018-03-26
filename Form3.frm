VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "MedAdvCFG v0.0.0 (Mednafen v0.9.x.x Frontend) by Nigel Todman [BASIC MODE]"
   ClientHeight    =   8730
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   8730
   ScaleMode       =   0  'User
   ScaleWidth      =   11829.89
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00000000&
      Caption         =   "Back"
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
      Left            =   11520
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11880
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   8520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00404040&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   210
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Total Games: x"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   9480
      TabIndex        =   2
      Top             =   120
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   12
      Left            =   8880
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   11
      Left            =   6120
      Picture         =   "Form3.frx":8341
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   10
      Left            =   3360
      Picture         =   "Form3.frx":10682
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   9
      Left            =   600
      Picture         =   "Form3.frx":189C3
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   8
      Left            =   8880
      Picture         =   "Form3.frx":20D04
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   7
      Left            =   6120
      Picture         =   "Form3.frx":29045
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   6
      Left            =   3360
      Picture         =   "Form3.frx":31386
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   5
      Left            =   600
      Picture         =   "Form3.frx":396C7
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   4
      Left            =   8880
      Picture         =   "Form3.frx":41A08
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   3
      Left            =   6120
      Picture         =   "Form3.frx":49D49
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   2
      Left            =   3360
      Picture         =   "Form3.frx":5208A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   1
      Left            =   600
      Picture         =   "Form3.frx":5A3CB
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   0
      Left            =   11880
      Picture         =   "Form3.frx":6270C
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MedEXE, FSO, tmp, tmp2, tmp3(99), BIOSFILE, BIOSPATH, ROMFILE, SystemCore, SysCore, BIOSSanity, ROMSanity, Stretch, PixelShader, VideoScaler, x, y, z
Dim cmdstring, Build, Frameskip, Fullscreen, TBlur, TblurAccum, AccumAmount, VideoIP, ActiveFile, XRes, YRes, ScaleFactor, LastPath, SavePath, BiosPathLoad
Dim ResetBios, ResetRom, ResetRomDir, ResetSave, FatalError, SystemRegion, SystemRegionLoad, ROMDIR, M3USize, LastFile, VideoDriver
Dim Bilinear, DisableSound, ForceMono, video_blit_timesync, video_glvsync, untrusted_fip_check, cd_image_memcache, scanlines, numplayers, customparams
Dim MedAdvGAMES, MedAdvCOVERS, MedAdvEXT, TotalGames, PageTotal, PageOn, TotalCovers, t, FNIndex, CurrFN, CoverFound, CoverSearched
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

Private Sub about_Click()
MsgBox "MedAdvCFG v" & Build & " (Mednafen v0.9.x.x Frontend)" & vbCrLf & "Written by Nigel Todman (nigel@nigeltodman.com)" & vbCrLf & "Primarily written as a PSX Frontend." & vbCrLf & "Tested with the following System Cores:" & vbCrLf & "GB, GBA, GG, MD, NES, PCE, PCE_FAST, PSX, SNES, SS, VB" & vbCrLf & vbCrLf & "Homepage: www.NigelTodman.com" & vbCrLf & "Facebook: facebook.com/nigel.todman.3" & vbCrLf & "Twitter: @Veritas_83" & vbCrLf & "YouTube: Veritas0923" & vbCrLf & "BTC: 18j2Env7QokhGG5MccS3LPBKnjsko6u4NQ"
End Sub

Private Sub Advanced_Click()
Basic.Checked = False
Form1.Basic.Checked = False
Form1.Visible = True
Form1.Advanced.Checked = True
Form3.Visible = False
Unload Form3
End Sub

Private Sub Chat_Click()
Shell ("cmd.exe /c start http://bit.ly/2k5E1Xq"), vbHide
End Sub

Private Sub Command1_Click()
If (PageOn + 1) > PageTotal Then
    PageOn = 1
    Label2.Caption = "Page: " & PageOn & "/" & PageTotal
    a = Find_Covers()
Else
    PageOn = PageOn + 1
    Label2.Caption = "Page: " & PageOn & "/" & PageTotal
    a = Find_Covers()
End If
End Sub

Private Sub Command2_Click()
If (PageOn - 1) <= 1 Then
    PageOn = 1
    Label2.Caption = "Page: " & PageOn & "/" & PageTotal
    a = Find_Covers()
Else
    PageOn = PageOn - 1
    Label2.Caption = "Page: " & PageOn & "/" & PageTotal
    a = Find_Covers()
End If
End Sub

Private Sub Command4_Click()
Form3.Visible = False
Form2.Visible = True
SysCore = Form1.ResetSysCore()
End Sub

Private Sub Documentation_Click()
Shell "cmd.exe /c start https://mednafen.github.io/documentation/", vbHide
End Sub

Private Sub Form_Load()
Form3.Width = 12345
Build = Form1.GetBuild()
Form3.Caption = "MedAdvCFG v" & Build & " (Mednafen v1.x Frontend) by Nigel Todman"
'Text1.Text = ""
SysCore = Form1.SetSysCore
Basic.Checked = True
Set FSO = CreateObject("Scripting.FileSystemObject")
'12 Games per page..
For y = 1 To 12
    Image1(y).Height = 2592
    Image1(y).Width = 2592
Next y
y = 1
If SysCore = "psx" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvPSX.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvPSXCOVERS.dat"
    MedAdvEXT = "cue"
ElseIf SysCore = "snes" Or SysCore = "snes_faust" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvSNES.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvSNESCOVERS.dat"
    MedAdvEXT = "smc"
    For y = 1 To 12
        Image1(y).Height = 2000
    Next y
    y = 1
ElseIf SysCore = "nes" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvNES.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvNESCOVERS.dat"
    MedAdvEXT = "nes"
    For y = 1 To 12
        Image1(y).Width = 1880
    Next y
    y = 1
ElseIf SysCore = "ss" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvSS.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvSSCOVERS.dat"
    MedAdvEXT = "cue"
    For y = 1 To 12
        Image1(y).Width = 1650
    Next y
    y = 1
ElseIf SysCore = "gba" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvGBA.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvGBACOVERS.dat"
    MedAdvEXT = "gba"
ElseIf SysCore = "gb" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvGBC.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvGBCCOVERS.dat"
    MedAdvEXT = "gbc"
ElseIf SysCore = "gg" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvGG.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvGGCOVERS.dat"
    MedAdvEXT = "gg"
    For y = 1 To 12
        Image1(y).Width = 2000
    Next y
    y = 1
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
    For y = 1 To 12
        Image1(y).Width = 1880
    Next y
    y = 1
ElseIf SysCore = "lynx" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvLYNX.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvLYNXCOVERS.dat"
    MedAdvEXT = "lnx"
ElseIf SysCore = "vb" Then
    MedAdvGAMES = VB.App.Path & "\dat\MedAdvVB.dat"
    MedAdvCOVERS = VB.App.Path & "\dat\MedAdvVBCOVERS.dat"
    MedAdvEXT = "vb"
End If

x = 1
y = 1
Set FSO = CreateObject("Scripting.FileSystemObject")

'Create list downloaded covers
Dim colFiles As New Collection
RecursiveDir colFiles, VB.App.Path & "\covers\" & SysCore & "\", "*.jpg", True
Dim vFile As Variant
If FSO.FileExists(MedAdvCOVERS) Then
    Close #9
    Open MedAdvCOVERS For Output As #9
    For Each vFile In colFiles
        Print #9, vFile
    Next vFile
    Close #9
Else
    Shell ("cmd.exe /c echo " & Chr(34) & Chr(34) & " >> " & MedAdvCOVERS)
End If


If FSO.FileExists(MedAdvGAMES) Then
    'Read list of extracted cue sheets
    Close #8
    Open MedAdvGAMES For Input As #8
    While Not EOF(8)
    Line Input #8, tmp
    GamesList(x) = tmp
    x = x + 1
    Wend
    Close #8
Else
    Shell ("cmd.exe /c echo " & Chr(34) & Chr(34) & " >> " & MedAdvGAMES)
End If

TotalGames = x
PageOn = 1
PageTotal = Int(Round(TotalGames / 12))
Label1.Caption = "Total Games: " & TotalGames
Label2.Caption = "Page: 1/" & PageTotal
x = 1

'Read list of installed covers
If FSO.FileExists(MedAdvCOVERS) Then
    Close #10
    Open MedAdvCOVERS For Input As #10
    While Not EOF(10)
    Line Input #10, tmp
    CoversList(x) = tmp
    x = x + 1
    Wend
    Close #10
Else
    Shell ("cmd.exe /c echo " & Chr(34) & Chr(34) & " >> " & MedAdvCOVERS)
End If

TotalCovers = x
a = Find_Covers()
End Sub
Function FileNameCleanup()
tmp = Replace(CurrName(x), ".cue", "")
tmp = Replace(tmp, ".smc", "")
tmp = Replace(tmp, ".nes", "")
tmp = Replace(tmp, ".iso", "")
tmp = Replace(tmp, ".bin", "")
tmp = Replace(tmp, ".gg", "")
tmp = Replace(tmp, ".lnx", "")
tmp = Replace(tmp, ".gba", "")
tmp = Replace(tmp, ".gbc", "")
tmp = Replace(tmp, ".gb", "")
tmp = Replace(tmp, ".rom", "")
tmp = Replace(tmp, ".ccd", "")
tmp = Replace(tmp, " (USA) ", "")
tmp = Replace(tmp, " (USA)", "")
tmp = Replace(tmp, "(USA)", "")
tmp = Replace(tmp, " (Demo) ", "")
tmp = Replace(tmp, " (Demo)", "")
tmp = Replace(tmp, "(Demo)", "")
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
tmp = Replace(tmp, "(VS)", "")
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
tmp = Replace(tmp, "(NJ037)", "")
tmp = Replace(tmp, "(UE)", "")
tmp = Replace(tmp, "(Ch)", "")
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
Function Find_Covers()
For t = 1 To 32
    On Error Resume Next
    tmparray(t) = Split(GamesList(1), "\")
    If InStr(1, tmparray(1)(t), "." & LCase(MedAdvEXT), 1) <> 0 Then
        FNIndex = t
        t = 32
    End If
Next t
If PageOn = 1 Then
    y = 1
    For x = 1 To TotalGames - 1
        z = 1
        CoverFound = False
        tmparray(x) = Split(GamesList(x), "\")
        CurrName(x) = tmparray(x)(FNIndex)
        ROMFILE = CurrName(x)
        tmp = FileNameCleanup()
        CoverSearched = tmp
            For z = 1 To TotalCovers
            If InStr(1, LCase(CoversList(z)), LCase(CoverSearched), 1) <> 0 Then
                If y >= 12 Then
                    x = TotalGames
                    y = 12
                    Image1(y).Picture = LoadPicture(CoversList(z))
                    Image1(y).Tag = GamesList(x)
                    Image1(y).ToolTipText = CoverSearched
                    z = TotalCovers
                    y = y + 1
                    CoverFound = True
                Else
                    Image1(y).Picture = LoadPicture(CoversList(z))
                    Image1(y).Tag = GamesList(x)
                    Image1(y).ToolTipText = CoverSearched
                    z = TotalCovers
                    y = y + 1
                    CoverFound = True
                End If
            End If
            Next z
        If CoverFound = False Then
            If y >= 12 Then
                x = TotalGames
                y = 12
                Image1(y).Picture = LoadPicture(VB.App.Path & "\covers\" & SysCore & "\nocover.jpg")
                Image1(y).Tag = GamesList(x)
                Image1(y).ToolTipText = CoverSearched
                y = y + 1
            Else
                Image1(y).Picture = LoadPicture(VB.App.Path & "\covers\" & SysCore & "\nocover.jpg")
                Image1(y).Tag = GamesList(x)
                Image1(y).ToolTipText = CoverSearched
                y = y + 1
            End If
        End If
        Next x
ElseIf PageOn >= 2 Then
    y = 1
    For x = ((12 * PageOn) - 12) To TotalGames - 1
        z = 1
        CoverFound = False
        tmparray(x) = Split(GamesList(x), "\")
        CurrName(x) = tmparray(x)(FNIndex)
        tmp = FileNameCleanup()
        CoverSearched = tmp
        For z = 1 To TotalCovers
            If InStr(1, LCase(CoversList(z)), LCase(CoverSearched), 1) <> 0 Then
                If y >= 12 Then
                    x = TotalGames
                    y = 12
                    Image1(y).Picture = LoadPicture(CoversList(z))
                    Image1(y).Tag = GamesList(x)
                    Image1(y).ToolTipText = CoverSearched
                    z = TotalCovers
                    y = y + 1
                    CoverFound = True
                Else
                    Image1(y).Picture = LoadPicture(CoversList(z))
                    Image1(y).Tag = GamesList(x)
                    Image1(y).ToolTipText = CoverSearched
                    z = TotalCovers
                    y = y + 1
                    CoverFound = True
                End If
            End If
            Next z
        If CoverFound = False Then
            If y >= 12 Then
                x = TotalGames
                y = 12
                Image1(y).Picture = LoadPicture(VB.App.Path & "\covers\" & SysCore & "\nocover.jpg")
                Image1(y).Tag = GamesList(x)
                Image1(y).ToolTipText = CoverSearched
                y = y + 1
            Else
                Image1(y).Picture = LoadPicture(VB.App.Path & "\covers\" & SysCore & "\nocover.jpg")
                Image1(y).Tag = GamesList(x)
                Image1(y).ToolTipText = CoverSearched
                y = y + 1
            End If
        End If
        Next x
End If

Form3.Refresh
For x = 1 To 12
    Image1(x).Refresh
Next x

End Function
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
Private Sub Image1_Click(Index As Integer)
'Quick Launch
a = LoadSettings()
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
            If Check12.Value = 1 Then
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
        cmdstring = cmdstring & " -" & SysCore & ".special hq2x"
        cmdstring = cmdstring & " -" & SysCore & ".xscale 2 -" & SysCore & ".xscalefs 2"
        cmdstring = cmdstring & " -" & SysCore & ".yscale 2 -" & SysCore & ".yscalefs 2"
        cmdstring = cmdstring & " -filesys.untrusted_fip_check 0"
        If SysCore = "psx" Or SysCore = "ss" Then
                cmdstring = cmdstring & " -" & SysCore & ".bios_sanity 0"
                cmdstring = cmdstring & " -" & SysCore & ".cd_sanity 0"
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
    
    
    'cmdstring = cmdstring & " " & Chr(34) & ROMFILE & Chr(34)
    cmdstring = cmdstring & " " & Chr(34) & Form3.Image1(Index).Tag & Chr(34)
    cmdstring = cmdstring & Chr(34)
If Form2.Check1.Value = 1 Then
    If FatalError = False Then
        'MsgBox cmdstring
    End If
    
    If FatalError = False Then
        Shell (cmdstring)
    Else
        FatalError = False
    End If
ElseIf Form2.Check1.Value = 0 Then
    On Error Resume Next
    'SysCore & vbcrlf & FullPath to Rom & CoverSearched
    Text1.Text = Text1.Text & vbCrLf & Form3.Image1(Index).Tag & vbCrLf & Image1(Index).ToolTipText & vbCrLf
    tmparray(1) = Split(Form3.Image1(Index).Tag, "\")
    Form3.Visible = False
    Form4.Visible = True
    Form4.Image1(1).Picture = Form3.Image1(Index).Picture
    Form4.Label2.Caption = tmparray(1)(FNIndex)
    Form4.Image1(1).Tag = Form3.Image1(Index).Tag
    Form4.Label11.Caption = Image1(Index).ToolTipText
    a = Form4.Validate_Rom(Form4.Image1(1).Tag)
    Form4.SetFocus
    Form4.Show
    Form4.Refresh
End If
End Sub

Private Sub Quit_Click()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
End Sub



