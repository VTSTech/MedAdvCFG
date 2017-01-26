VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "MedAdvCFG v0.0.0 (Mednafen v0.9.x.x Frontend) by Nigel Todman [BASIC MODE]"
   ClientHeight    =   8670
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11280
   LinkTopic       =   "Form3"
   ScaleHeight     =   8670
   ScaleMode       =   0  'User
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   11880
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form3.frx":0000
      Top             =   6960
      Visible         =   0   'False
      Width           =   2175
   End
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
      Height          =   2595
      Index           =   12
      Left            =   8880
      Picture         =   "Form3.frx":0011
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   11
      Left            =   6120
      Picture         =   "Form3.frx":2B40B
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   10
      Left            =   3360
      Picture         =   "Form3.frx":56805
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   9
      Left            =   600
      Picture         =   "Form3.frx":81BFF
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   8
      Left            =   8880
      Picture         =   "Form3.frx":ACFF9
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   7
      Left            =   6120
      Picture         =   "Form3.frx":D83F3
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   6
      Left            =   3360
      Picture         =   "Form3.frx":1037ED
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2595
      Index           =   5
      Left            =   600
      Picture         =   "Form3.frx":12EBE7
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2592
      Index           =   4
      Left            =   8880
      Picture         =   "Form3.frx":159FE1
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2656
   End
   Begin VB.Image Image1 
      Height          =   2592
      Index           =   3
      Left            =   6120
      Picture         =   "Form3.frx":1853DB
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2656
   End
   Begin VB.Image Image1 
      Height          =   2592
      Index           =   2
      Left            =   3360
      Picture         =   "Form3.frx":1B07D5
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2656
   End
   Begin VB.Image Image1 
      Height          =   2592
      Index           =   1
      Left            =   600
      Picture         =   "Form3.frx":1DBBCF
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2656
   End
   Begin VB.Image Image1 
      Height          =   2655
      Index           =   0
      Left            =   11880
      Picture         =   "Form3.frx":206FC9
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

Private Sub Command2_Click()
If (PageOn - 1) <= 0 Then
    PageOn = 1
Else
    PageOn = PageOn - 1
    Label2.Caption = "Page: " & PageOn & "/" & PageTotal
    a = Find_Covers()
    Form3.Refresh
    For y = 1 To 12
    Image1(y).Refresh
    Next y
End If
End Sub

Private Sub Form_Load()
Form3.Width = 12345
Build = "0.2.4"
Form3.Caption = "MedAdvCFG v" & Build & " (Mednafen v0.9.x.x Frontend) by Nigel Todman [BASIC MODE]"
Text1.Text = ""
SysCore = Form1.SetSysCore
Set FSO = CreateObject("Scripting.FileSystemObject")
'12 Games per page..
If SysCore = "psx" Then
    MedAdvGAMES = VB.App.Path & "\MedAdvPSX.dat"
    MedAdvCOVERS = VB.App.Path & "\MedAdvPSXCOVERS.dat"
    MedAdvEXT = "cue"
    For x = 1 To 12
        Image1(x).Height = 2592
    Next x
ElseIf SysCore = "snes" Then
    MedAdvGAMES = VB.App.Path & "\MedAdvSNES.dat"
    MedAdvCOVERS = VB.App.Path & "\MedAdvSNESCOVERS.dat"
    MedAdvEXT = "smc"
    For x = 1 To 12
        Image1(x).Height = 2000
    Next x
ElseIf SysCore = "nes" Then
    MedAdvGAMES = VB.App.Path & "\MedAdvNES.dat"
    MedAdvCOVERS = VB.App.Path & "\MedAdvNESCOVERS.dat"
    MedAdvEXT = "nes"
End If

x = 1
y = 1

'Create list downloaded covers
Dim colFiles As New Collection
RecursiveDir colFiles, VB.App.Path & "\covers\" & SysCore & "\", "*.jpg", True

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
Function FileNameCleanup()
tmp = Replace(CurrName(x), ".cue", "")
tmp = Replace(tmp, ".smc", "")
tmp = Replace(tmp, ".nes", "")
tmp = Replace(tmp, ".iso", "")
tmp = Replace(tmp, ".bin", "")
tmp = Replace(tmp, " (USA) ", "")
tmp = Replace(tmp, " (USA)", "")
tmp = Replace(tmp, "(USA)", "")
tmp = Replace(tmp, " (v1.0)", "")
tmp = Replace(tmp, " (v1.1)", "")
tmp = Replace(tmp, " (v1.2)", "")
tmp = Replace(tmp, " (v1.3)", "")
tmp = Replace(tmp, " (v1.4)", "")
tmp = Replace(tmp, " (v1.5)", "")
tmp = Replace(tmp, " (v1.6)", "")
tmp = Replace(tmp, "(v1.0)", "")
tmp = Replace(tmp, "(v1.1)", "")
tmp = Replace(tmp, "(v1.2)", "")
tmp = Replace(tmp, "(v1.3)", "")
tmp = Replace(tmp, "(v1.4)", "")
tmp = Replace(tmp, "(v1.5)", "")
tmp = Replace(tmp, "(v1.6)", "")
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
tmp = Replace(tmp, " (En,Fr,De,Es,It)", "")
tmp = Replace(tmp, "(En,Fr,De,Es,It)", "")
tmp = Replace(tmp, " (En,Fr,De,Sv)", "")
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
tmp = Replace(tmp, "(GDI)", "")
tmp = Replace(tmp, "(NOD)", "")
tmp = Replace(tmp, "(ECD)", "")
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
tmp = Replace(tmp, "'", "")
tmp = Replace(tmp, ",", "")
tmp = Replace(tmp, ".", "")
tmp = Replace(tmp, " - ", "-")
tmp = Replace(tmp, "  ", " ")
tmp = Replace(tmp, "_-_", "-")
tmp = LTrim(RTrim(tmp))
tmp = Replace(tmp, " ", "_")
FileNameCleanup = tmp
End Function
Function Find_Covers()
For t = 1 To 32
    On Error Resume Next
    tmparray(t) = Split(GamesList(1), "\")
    If InStr(1, tmparray(1)(t), LCase(MedAdvEXT), 1) <> 0 Then
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
        tmp = FileNameCleanup()
        CoverSearched = tmp
            For z = 1 To TotalCovers
            If InStr(1, LCase(CoversList(z)), LCase(CoverSearched), 1) >= 1 Then
                'MsgBox "Cover: " & PSXCovers(x) & " for " & tmp
                If y = 13 Then
                    y = 1
                Else
                    CoverFound = True
                    Image1(y).Picture = LoadPicture(CoversList(z))
                    Image1(y).Tag = GamesList(x)
                    Image1(y).ToolTipText = CoverSearched
                    z = TotalCovers
                    y = y + 1
                End If
            End If
            Next z
        If CoverFound = False Then
            Image1(y).Picture = LoadPicture(VB.App.Path & "\covers\" & SysCore & "\nocover.jpg")
            Image1(y).Tag = GamesList(x)
            Image1(y).ToolTipText = CoverSearched
            y = y + 1
        End If
        Next x
ElseIf PageOn >= 2 Then
    y = 1
    For x = ((12 * PageOn) - 13) To TotalGames - 1
        z = 1
        CoverFound = False
        tmparray(x) = Split(GamesList(x), "\")
        CurrName(x) = tmparray(x)(FNIndex)
        tmp = FileNameCleanup()
        CoverSearched = tmp
        For z = 1 To TotalCovers
            If InStr(1, LCase(CoversList(z)), LCase(CoverSearched), 1) >= 1 Then
                'MsgBox "Cover: " & PSXCovers(x) & " for " & tmp
                If y = 13 Then
                    y = 1
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
            Image1(y).Picture = LoadPicture(VB.App.Path & "\covers\" & SysCore & "\nocover.jpg")
            Image1(y).Tag = GamesList(x)
            Image1(y).ToolTipText = CoverSearched
            y = y + 1
        End If
        Next x
End If
End Function

Private Sub Image1_Click(Index As Integer)
'Quick Launch
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
    
    
    'cmdstring = cmdstring & " " & Chr(34) & ROMFILE & Chr(34)
    cmdstring = cmdstring & " " & Chr(34) & Form3.Image1(Index).Tag & Chr(34)
    cmdstring = cmdstring & Chr(34)
If Form2.Check1.Value = 1 Then
    If FatalError = False Then
        MsgBox cmdstring
    End If
    
    If FatalError = False Then
        Shell (cmdstring)
    Else
        FatalError = False
    End If
ElseIf Form2.Check1.Value = 0 Then
    'SysCore & vbcrlf & FullPath to Rom & CoverSearched
    Text1.Text = Text1.Text & vbCrLf & Form3.Image1(Index).Tag & vbCrLf & Image1(Index).ToolTipText & vbCrLf
    tmparray(1) = Split(Form3.Image1(Index).Tag, "\")
    Form3.Visible = False
    Form4.Image1(1).Picture = Form3.Image1(Index).Picture
    Form4.Label2.Caption = tmparray(1)(FNIndex)
    Form4.Image1(1).Tag = Form3.Image1(Index).Tag
    Form4.Label11.Caption = Image1(Index).ToolTipText
    a = Form4.Validate_Rom(Form4.Image1(1).Tag)
    Form4.Visible = True
End If
End Sub

Private Sub Quit_Click()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
End Sub
