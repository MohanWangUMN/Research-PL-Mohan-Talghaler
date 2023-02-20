VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCDSoftware 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spectral Products Software Download"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   6240
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   6000
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   855
      Left            =   1320
      Picture         =   "CD2.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   60
      TabIndex        =   4
      Top             =   1200
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   8
      Tab             =   6
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   -2147483637
      TabCaption(0)   =   "SW MANUALS"
      TabPicture(0)   =   "CD2.frx":4864
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame7"
      Tab(0).Control(1)=   "CommonDialog2"
      Tab(0).Control(2)=   "Check34"
      Tab(0).Control(3)=   "Check33"
      Tab(0).Control(4)=   "Check32"
      Tab(0).Control(5)=   "Check30"
      Tab(0).Control(6)=   "Check29"
      Tab(0).Control(7)=   "Check28"
      Tab(0).Control(8)=   "Check27"
      Tab(0).Control(9)=   "Check58"
      Tab(0).Control(10)=   "Check56"
      Tab(0).Control(11)=   "Check55"
      Tab(0).Control(12)=   "Check9"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "MONOCHROMATORS"
      TabPicture(1)   =   "CD2.frx":4880
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "Text16"
      Tab(1).Control(4)=   "Check3"
      Tab(1).Control(5)=   "Check2"
      Tab(1).Control(6)=   "Check1"
      Tab(1).Control(7)=   "Check57"
      Tab(1).Control(8)=   "Text36"
      Tab(1).Control(9)=   "Check78"
      Tab(1).Control(10)=   "Text11"
      Tab(1).Control(11)=   "Text15"
      Tab(1).Control(12)=   "Text35"
      Tab(1).Control(13)=   "Text38"
      Tab(1).Control(14)=   "Check83"
      Tab(1).Control(15)=   "Text40"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "FILTER WHEELS"
      TabPicture(2)   =   "CD2.frx":489C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Text30"
      Tab(2).Control(2)=   "Text29"
      Tab(2).Control(3)=   "Text28"
      Tab(2).Control(4)=   "Check6"
      Tab(2).Control(5)=   "Check5"
      Tab(2).Control(6)=   "Check4"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "LABVIEW RUNTIME"
      TabPicture(3)   =   "CD2.frx":48B8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Check18"
      Tab(3).Control(1)=   "Check19"
      Tab(3).Control(2)=   "Text32"
      Tab(3).Control(3)=   "Text33"
      Tab(3).Control(4)=   "Frame5"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "DETECTORS"
      TabPicture(4)   =   "CD2.frx":48D4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1"
      Tab(4).Control(1)=   "Frame4"
      Tab(4).Control(2)=   "Text25"
      Tab(4).Control(3)=   "Text24"
      Tab(4).Control(4)=   "Check16"
      Tab(4).Control(5)=   "Check15"
      Tab(4).Control(6)=   "Check50"
      Tab(4).Control(7)=   "Text37"
      Tab(4).Control(8)=   "Check7"
      Tab(4).Control(9)=   "Text2"
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "SOURCE CODES"
      TabPicture(5)   =   "CD2.frx":48F0
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame8"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Check48"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Check44"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Check43"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Check41"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Check51"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Check60"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Check62"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Check47"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Check75"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "Check76"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "Check77"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "Check82"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "Check87"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "Check86"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "Check88"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "Check8"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).ControlCount=   17
      TabCaption(6)   =   "INST. MANUALS"
      TabPicture(6)   =   "CD2.frx":490C
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Frame9"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Check74"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Text13"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Check71"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Text10"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "Check73"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "Text12"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "Check61"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "Text3"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "Check64"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).Control(10)=   "Text1"
      Tab(6).Control(10).Enabled=   0   'False
      Tab(6).Control(11)=   "Check65"
      Tab(6).Control(11).Enabled=   0   'False
      Tab(6).Control(12)=   "Text4"
      Tab(6).Control(12).Enabled=   0   'False
      Tab(6).Control(13)=   "Check69"
      Tab(6).Control(13).Enabled=   0   'False
      Tab(6).Control(14)=   "Text8"
      Tab(6).Control(14).Enabled=   0   'False
      Tab(6).Control(15)=   "Check10"
      Tab(6).Control(15).Enabled=   0   'False
      Tab(6).Control(16)=   "Text5"
      Tab(6).Control(16).Enabled=   0   'False
      Tab(6).ControlCount=   17
      TabCaption(7)   =   "MISCELLANEOUS"
      TabPicture(7)   =   "CD2.frx":4928
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame10"
      Tab(7).ControlCount=   1
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6480
         TabIndex        =   95
         Text            =   "318 KB"
         Top             =   2820
         Width           =   735
      End
      Begin VB.CheckBox Check10 
         Caption         =   "AD111 CM/DK User Manual"
         Height          =   255
         Left            =   360
         TabIndex        =   94
         ToolTipText     =   "AD111 CM/DK User Manual"
         Top             =   2820
         Width           =   5295
      End
      Begin VB.CheckBox Check9 
         Caption         =   "AD111 CM/DK LabView Exe"
         Height          =   255
         Left            =   -74640
         TabIndex        =   93
         Top             =   2340
         Width           =   3375
      End
      Begin VB.CheckBox Check8 
         Caption         =   "AD111 CM/DK LabView VI"
         Height          =   255
         Left            =   -74520
         TabIndex        =   92
         Top             =   2100
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68760
         TabIndex        =   91
         Text            =   "1,144 Kb"
         Top             =   1860
         Width           =   1200
      End
      Begin VB.CheckBox Check7 
         Caption         =   "AD111 CM/DK LabView Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   90
         ToolTipText     =   "LabView executable running the SM302 and the CM series monochromators"
         Top             =   1860
         Width           =   3495
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6480
         TabIndex        =   89
         Text            =   "547 KB"
         Top             =   2580
         Width           =   735
      End
      Begin VB.CheckBox Check69 
         Caption         =   "CM110 to AST-XE-175EX, Mounting Instructions"
         Height          =   255
         Left            =   360
         TabIndex        =   88
         ToolTipText     =   "CM110 to AST-XE-175EX"
         Top             =   2580
         Width           =   5295
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6480
         TabIndex        =   87
         Text            =   "638 KB"
         Top             =   2340
         Width           =   735
      End
      Begin VB.CheckBox Check65 
         Caption         =   "SM302 InGaAs Array Spectrometer, User Manual"
         Height          =   255
         Left            =   360
         TabIndex        =   86
         ToolTipText     =   "SM302 InGaAs Array Spectrometer"
         Top             =   2340
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6480
         TabIndex        =   85
         Text            =   "1,216 KB"
         Top             =   2100
         Width           =   735
      End
      Begin VB.CheckBox Check64 
         Caption         =   "DK-Series Monochromator/Spectrograph, User Manual"
         Height          =   255
         Left            =   360
         TabIndex        =   84
         ToolTipText     =   "DK-Series Monochromator/Spectrograph"
         Top             =   2100
         Width           =   5295
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6480
         TabIndex        =   83
         Text            =   "1,184 KB"
         Top             =   1860
         Width           =   735
      End
      Begin VB.CheckBox Check61 
         Caption         =   "CM-Series Monochromator/Spectrograph, User Manual"
         Height          =   315
         Left            =   360
         TabIndex        =   82
         ToolTipText     =   "CM-Series Monochromator/Spectrograph"
         Top             =   1860
         Width           =   5295
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6480
         TabIndex        =   81
         Text            =   "135 KB"
         Top             =   1620
         Width           =   735
      End
      Begin VB.CheckBox Check73 
         Caption         =   "AB300-T Automatic Filter Wheel, User Manual"
         Height          =   315
         Left            =   360
         TabIndex        =   80
         ToolTipText     =   "AB300-T Automatic Filter Wheel"
         Top             =   1620
         Width           =   5295
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6480
         TabIndex        =   79
         Text            =   "436 KB"
         Top             =   1380
         Width           =   735
      End
      Begin VB.CheckBox Check71 
         Caption         =   "AB300-Series Automatic Filter Wheels, User Manual"
         Height          =   315
         Left            =   360
         TabIndex        =   78
         ToolTipText     =   "AB300-Series Automatic Filter Wheels"
         Top             =   1380
         Width           =   5295
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6480
         TabIndex        =   77
         Text            =   "373 KB"
         Top             =   1140
         Width           =   735
      End
      Begin VB.CheckBox Check74 
         Caption         =   "AD131 Photodetector Module, User Manual"
         Height          =   315
         Left            =   360
         TabIndex        =   76
         ToolTipText     =   "AD131 Photodetector Module"
         Top             =   1140
         Width           =   5295
      End
      Begin VB.CheckBox Check55 
         Caption         =   "AD131 CMDK LabView Exe"
         Height          =   255
         Left            =   -74640
         TabIndex        =   75
         Top             =   1380
         Width           =   3375
      End
      Begin VB.CheckBox Check56 
         Caption         =   "SM302CM LabView Exe"
         Height          =   255
         Left            =   -70680
         TabIndex        =   74
         Top             =   2100
         Width           =   3255
      End
      Begin VB.CheckBox Check88 
         Caption         =   "AB300 Series Visual Basic 6"
         Height          =   255
         Left            =   -70440
         TabIndex        =   73
         Top             =   1140
         Width           =   3015
      End
      Begin VB.CheckBox Check86 
         Caption         =   "CM110/112 RS232 Visual Basic 6 "
         Height          =   255
         Left            =   -70440
         TabIndex        =   72
         Top             =   2340
         Width           =   3015
      End
      Begin VB.CheckBox Check87 
         Caption         =   "DK240/242/480 Visual Basic 6"
         Height          =   255
         Left            =   -70440
         TabIndex        =   71
         Top             =   2820
         Width           =   3015
      End
      Begin VB.TextBox Text40 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68640
         TabIndex        =   70
         Text            =   "834,693 bytes"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox Check83 
         Caption         =   "CM110/112 ICS Gpib LabView Demo Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   69
         Top             =   2060
         Width           =   3375
      End
      Begin VB.CheckBox Check82 
         Caption         =   "CM110/112 LabView ICS Gpib VI"
         Height          =   255
         Left            =   -70440
         TabIndex        =   68
         Top             =   1860
         Width           =   3015
      End
      Begin VB.TextBox Text38 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68640
         TabIndex        =   63
         Text            =   "94,208 bytes"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text35 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68640
         TabIndex        =   60
         Text            =   "621,685 bytes"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68640
         TabIndex        =   59
         Text            =   "829,481 bytes"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68640
         TabIndex        =   58
         Text            =   "2,294,464 bytes"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox Check78 
         Caption         =   "CM110/112 RS232 C++ Ver 6 Demo exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   57
         Top             =   2280
         Width           =   3375
      End
      Begin VB.CheckBox Check77 
         Caption         =   "DK240/242/480 LabView Gpib VI"
         Height          =   255
         Left            =   -70440
         TabIndex        =   56
         Top             =   3300
         Width           =   3015
      End
      Begin VB.CheckBox Check76 
         Caption         =   "CM110/112 RS232 Visual C++ Demo "
         Height          =   255
         Left            =   -70440
         TabIndex        =   55
         Top             =   2100
         Width           =   3015
      End
      Begin VB.CheckBox Check75 
         Caption         =   "CM110/112 LabView NI Gpib VI"
         Height          =   255
         Left            =   -70440
         TabIndex        =   54
         Top             =   1620
         Width           =   3015
      End
      Begin VB.CheckBox Check47 
         Caption         =   "CM110/112 LabView RS232 VI"
         Height          =   255
         Left            =   -70440
         TabIndex        =   53
         Top             =   1380
         Width           =   3015
      End
      Begin VB.Frame Frame10 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   52
         Top             =   1080
         Width           =   7215
      End
      Begin VB.TextBox Text37 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68760
         TabIndex        =   51
         Text            =   "931 Kb"
         Top             =   1380
         Width           =   1200
      End
      Begin VB.CheckBox Check62 
         Caption         =   "SM302/CM110/112 LabView Demo"
         Height          =   255
         Left            =   -70440
         TabIndex        =   47
         Top             =   3540
         Width           =   3015
      End
      Begin VB.CheckBox Check60 
         Caption         =   "DK240/242/480 C++ Ver6 Demo"
         Height          =   255
         Left            =   -70440
         TabIndex        =   46
         Top             =   2580
         Width           =   3015
      End
      Begin VB.TextBox Text36 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68760
         TabIndex        =   45
         Text            =   "787,456 bytes"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox Check58 
         Caption         =   "CM110/112 Gpib LabView Demo Exe"
         Height          =   255
         Left            =   -70680
         TabIndex        =   43
         Top             =   1860
         Width           =   3255
      End
      Begin VB.CheckBox Check57 
         Caption         =   "CM110/112 NI Gpib LabView Demo Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   42
         Top             =   1845
         Width           =   3375
      End
      Begin VB.CheckBox Check51 
         Caption         =   "AD131 CM/DK Series LabView VI"
         Height          =   255
         Left            =   -74520
         TabIndex        =   41
         Top             =   1380
         Width           =   3255
      End
      Begin VB.CheckBox Check50 
         Caption         =   "AD131 CM/DK Series LabView Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   39
         ToolTipText     =   "LabView Executable running the AD131 and the DK Series Monochromators"
         Top             =   1380
         Width           =   3495
      End
      Begin VB.CheckBox Check18 
         Caption         =   "Version 5.1"
         Height          =   255
         Left            =   -74520
         TabIndex        =   26
         Top             =   1380
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CM110/112 Visual Basic Demo Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   34
         ToolTipText     =   "Visual Basic Demo executable running the CM110/112"
         Top             =   1380
         Width           =   3375
      End
      Begin VB.CheckBox Check2 
         Caption         =   "CM110/112 LabView Demo Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   33
         ToolTipText     =   "LabView Executable running the CM110/112 Monchromator"
         Top             =   1620
         Width           =   3375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "DK240/242/480 Visual Basic Demo Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   32
         ToolTipText     =   "Visual Basic Executable running the DK240/242/480 Monchromator"
         Top             =   3240
         Width           =   3855
      End
      Begin VB.CheckBox Check4 
         Caption         =   "AB300 Series GPIB LabView Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   31
         ToolTipText     =   "LabView Executable running the AB300 Series Filter Wheels via GPIB"
         Top             =   1260
         Width           =   3495
      End
      Begin VB.CheckBox Check5 
         Caption         =   "AB300 Series RS232 LabView Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   30
         ToolTipText     =   "Labview executable running the AB300 series filter wheels via RS232"
         Top             =   1500
         Width           =   3495
      End
      Begin VB.CheckBox Check6 
         Caption         =   "AB300 Series RS232 Visual Basic Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   29
         ToolTipText     =   "Visual Basic Demo running the AB300 series filter wheels via RS232"
         Top             =   1740
         Width           =   3495
      End
      Begin VB.CheckBox Check15 
         Caption         =   "AD131 Stand Alone LabView Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   28
         ToolTipText     =   "LabView stand alone executable running the AD131 detector"
         Top             =   1140
         Width           =   3495
      End
      Begin VB.CheckBox Check16 
         Caption         =   "SM302 CM110/112 LabView Exe"
         Height          =   255
         Left            =   -74520
         TabIndex        =   27
         ToolTipText     =   "LabView executable running the SM302 and the CM series monochromators"
         Top             =   1620
         Width           =   3495
      End
      Begin VB.CheckBox Check19 
         Caption         =   "Version 6.0"
         Height          =   255
         Left            =   -74520
         TabIndex        =   25
         Top             =   1620
         Width           =   1815
      End
      Begin VB.CheckBox Check27 
         Caption         =   "AD131 Stand Alone LabView Exe"
         Height          =   255
         Left            =   -74640
         TabIndex        =   24
         Top             =   1140
         Width           =   3375
      End
      Begin VB.CheckBox Check28 
         Caption         =   "AB300 Series GPIB LabView Exe"
         Height          =   255
         Left            =   -74640
         TabIndex        =   23
         Top             =   1620
         Width           =   3375
      End
      Begin VB.CheckBox Check29 
         Caption         =   "AB300 Series RS232 LabView Exe"
         Height          =   255
         Left            =   -74640
         TabIndex        =   22
         Top             =   1860
         Width           =   3375
      End
      Begin VB.CheckBox Check30 
         Caption         =   "AB300 Series RS232 Visual Basic Exe"
         Height          =   255
         Left            =   -74640
         TabIndex        =   21
         Top             =   2100
         Width           =   3375
      End
      Begin VB.CheckBox Check32 
         Caption         =   "CM110/112 Visual Basic Demo Exe"
         Height          =   255
         Left            =   -70680
         TabIndex        =   20
         Top             =   1620
         Width           =   3255
      End
      Begin VB.CheckBox Check33 
         Caption         =   "DK240/242/480 Visual Basic Demo Exe"
         Height          =   255
         Left            =   -70680
         TabIndex        =   19
         Top             =   1380
         Width           =   3255
      End
      Begin VB.CheckBox Check34 
         Caption         =   "CM110/112 LabView Demo Exe"
         Height          =   255
         Left            =   -70680
         TabIndex        =   18
         Top             =   1140
         Width           =   3255
      End
      Begin VB.CheckBox Check41 
         Caption         =   "AD131 Stand Alone LabView VI"
         Height          =   255
         Left            =   -74520
         TabIndex        =   17
         Top             =   1140
         Width           =   3255
      End
      Begin VB.CheckBox Check43 
         Caption         =   "AB300 Series GPIB LabView VI"
         Height          =   255
         Left            =   -74520
         TabIndex        =   16
         Top             =   1620
         Width           =   3255
      End
      Begin VB.CheckBox Check44 
         Caption         =   "AB300 Series RS232 LabView VI"
         Height          =   255
         Left            =   -74520
         TabIndex        =   15
         Top             =   1860
         Width           =   3255
      End
      Begin VB.CheckBox Check48 
         Caption         =   "DK240/242/480 LabView RS232 VI"
         Height          =   255
         Left            =   -70440
         TabIndex        =   14
         Top             =   3060
         Width           =   3015
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   7800
         TabIndex        =   13
         Text            =   "2,294,464 bytes"
         Top             =   -240
         Width           =   1215
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68760
         TabIndex        =   12
         Text            =   "2,045,071 bytes"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text24 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68760
         TabIndex        =   11
         Text            =   "600 Kb"
         Top             =   1140
         Width           =   1200
      End
      Begin VB.TextBox Text25 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68760
         TabIndex        =   10
         Text            =   "1,200 Kb"
         Top             =   1620
         Width           =   1200
      End
      Begin VB.TextBox Text28 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68640
         TabIndex        =   9
         Text            =   "619,534 bytes"
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox Text29 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68640
         TabIndex        =   8
         Text            =   "624,810 bytes"
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox Text30 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68640
         TabIndex        =   7
         Text            =   "2,872,401 bytes"
         Top             =   1740
         Width           =   1215
      End
      Begin VB.TextBox Text32 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68760
         TabIndex        =   6
         Text            =   "2,667,128 bytes"
         Top             =   1380
         Width           =   1335
      End
      Begin VB.TextBox Text33 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -68760
         TabIndex        =   5
         Text            =   "12,099,739 bytes"
         Top             =   1620
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   -69360
         Top             =   3900
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame5 
         Caption         =   "LabView RunTime Engine's"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   35
         Top             =   960
         Width           =   7455
      End
      Begin VB.Frame Frame2 
         Caption         =   "DK240/242/480"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   36
         Top             =   2880
         Width           =   7695
         Begin VB.CheckBox Check80 
            Caption         =   "DK240/242/480 Gpib LabView 6 Demo Exe"
            Height          =   255
            Left            =   360
            TabIndex        =   65
            Top             =   1080
            Width           =   3855
         End
         Begin VB.CheckBox Check79 
            Caption         =   "DK240/242/480 RS232 LabView 6 Demo Exe"
            Height          =   255
            Left            =   360
            TabIndex        =   64
            Top             =   840
            Width           =   3855
         End
         Begin VB.CheckBox Check59 
            Caption         =   "DK240/242/480 C++ Ver 6 Demo Exe"
            Height          =   255
            Left            =   360
            TabIndex        =   44
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label Label5 
            Caption         =   "782,806 bytes"
            Height          =   255
            Left            =   6120
            TabIndex        =   67
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "672,949 bytes"
            Height          =   255
            Left            =   6120
            TabIndex        =   66
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   37
         Top             =   1020
         Width           =   7695
      End
      Begin VB.Frame Frame7 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   38
         Top             =   960
         Width           =   7575
      End
      Begin VB.Frame Frame4 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   48
         Top             =   840
         Width           =   7465
      End
      Begin VB.Frame Frame8 
         Height          =   3855
         Left            =   -74760
         TabIndex        =   49
         Top             =   840
         Width           =   7455
      End
      Begin VB.Frame Frame9 
         Height          =   4095
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   7575
      End
      Begin VB.Frame Frame1 
         Caption         =   "CM110/112 "
         Height          =   1815
         Left            =   -74880
         TabIndex        =   62
         Top             =   960
         Width           =   7695
      End
      Begin VB.Label Label2 
         Caption         =   "94,208 bytes"
         Height          =   255
         Left            =   -68760
         TabIndex        =   61
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "931,149 bytes"
         Height          =   255
         Left            =   -68640
         TabIndex        =   40
         Top             =   3060
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCDSoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
retval = Shell("CDSoftware\MONCHROMATORS\Visual Basic\Source\CmVbSource_Setup\Setup.exe", 1)
End Sub

Private Sub Check10_Click()

  If Right$(App.Path, 1) = "\" Then
    On Error GoTo DialogError
    File1 = App.Path & "CDSoftware\DETECTORS\InstManuals\AD111\AD111UserManual.doc"
  Else
    File1 = App.Path & "\CDSoftware\DETECTORS\InstManuals\AD111\AD111UserManual.doc"
  End If
  On Error GoTo DialogError
  With CommonDialog2
    .FileName = "AD111UserManual.doc"
    .CancelError = True
    .Filter = "PDF File (*.pdf)|*.pdf|Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
    .FilterIndex = 1
    .DialogTitle = "Select a Directory to Save"
    .ShowSave
  End With
  File2 = CommonDialog2.FileName
  FileCopy File1, File2
DialogError:


End Sub

Private Sub Check15_Click()
retval = Shell("CDSoftware\DETECTORS\Ad131\Exe\SetUp.exe", 1)
End Sub

Private Sub Check16_Click()
retval = Shell("CDSoftware\DETECTORS\SM302\SetUp\disks\SetUp.exe", 1)
End Sub

Private Sub Check17_Click()
retval = Shell("CDSoftware\DETECTORS\Ad150\SM161Pro\SetUp.exe", 1)
End Sub

Private Sub Check18_Click()
retval = Shell("CDSoftware\LabViewEngine\RunTime\Setup.exe", 1)
End Sub

Private Sub Check19_Click()
retval = Shell("CDSoftware\LabViewEngine\LabView6i\lvrteinstall.exe", 1)
End Sub

Private Sub Check2_Click()
retval = Shell("CDSoftware\Monchromators\LabView\Executable\cm\setup\SETUP.exe", 1)
End Sub

Private Sub Check27_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\DETECTORS\Ad131\AD131.doc"
Else
File1 = App.Path & "\CDSoftware\DETECTORS\Ad131\AD131.doc"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "AD131"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check28_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\FILTER WHEELS\LabView\GPIB\Executable\GpAbe-a.doc"
Else
File1 = App.Path & "\CDSoftware\FILTER WHEELS\LabView\GPIB\Executable\GpAbe-a.doc"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "GpAbe-a"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check29_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\FILTER WHEELS\LabView\RS232\Executable\AB-Read.doc"
Else
File1 = App.Path & "\CDSoftware\FILTER WHEELS\LabView\RS232\Executable\AB-Read.doc"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "AB-Read"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check3_Click()
retval = Shell("CDSoftware\MONCHROMATORS\Visual Basic\Source\DkVbSource_Setup\SetUp.exe", 1)
End Sub

Private Sub Check30_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\FILTER WHEELS\Visual Basic\ABdemo-C.doc"
Else
File1 = App.Path & "\CDSoftware\FILTER WHEELS\Visual Basic\ABdemo-C.doc"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "ABdemo-C.doc"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check32_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\Visual Basic\CM VB RS232\8-2085-a.doc"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\Visual Basic\CM VB RS232\8-2085-a.doc"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "8-2085-a"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check33_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\Visual Basic\DK VB RS232\DkdemovB.doc"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\Visual Basic\DK VB RS232\DkdemovB.doc"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "DkdemovB.doc"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check34_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\LabView\Executable\CM\setup\8-0009-a.doc"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\LabView\Executable\CM\setup\8-0009-a.doc"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "8-0009-a"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check4_Click()
retval = Shell("CDSoftware\FILTER WHEELS\LabView\GPIB\Executable\Setup\SetUp.exe", 1)
End Sub

Private Sub Check41_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\DETECTORS\Ad131\Source\Ad131Is.zip"
Else
File1 = App.Path & "\CDSoftware\DETECTORS\Ad131\Source\Ad131Is.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "Ad131Is.zip"
.CancelError = True
.Filter = "VI (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check43_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\FILTER WHEELS\LabView\GPIB\Source\8-2109-b.zip"
Else
File1 = App.Path & "\CDSoftware\FILTER WHEELS\LabView\GPIB\Source\8-2109-b.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "8-2109-b.zip"
.CancelError = True
.Filter = "VI (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check44_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
  File1 = App.Path & "CDSoftware\FILTER WHEELS\LabView\RS232\Source\8-2002-e.zip"
Else
  File1 = App.Path & "\CDSoftware\FILTER WHEELS\LabView\RS232\Source\8-2002-e.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "8-2002-e.zip"
.CancelError = True
.Filter = "VI (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check47_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\LabView\Source\CM RS232vi\CmRS232.zip"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\LabView\Source\CM RS232vi\CmRS232.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "CmRS232.zip"
.CancelError = True
.Filter = "VI (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check48_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\LabView\Source\DK RS232vi\DkRs232.zip"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\LabView\Source\DK RS232vi\DkRs232.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "DkRs232.zip"
.CancelError = True
.Filter = "ZIP (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check49_Click()
retval = Shell("CDSoftware\DETECTORS\Ad150\AD150CMDK\Exe\Installer\disks\SetUp.exe", 1)
End Sub

Private Sub Check5_Click()
retval = Shell("CDSoftware\FILTER WHEELS\LabView\RS232\Executable\Setup\SetUp.exe", 1)
End Sub

Private Sub Check50_Click()
retval = Shell("CDSoftware\DETECTORS\Ad131\Exe\AD131DK\SetUp.exe", 1)
End Sub

Private Sub Check51_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\DETECTORS\Ad131\Exe\AD131DK\Source\Ad131CmDkSource.zip"
Else
File1 = App.Path & "\CDSoftware\DETECTORS\Ad131\Exe\AD131DK\Source\Ad131CmDkSource.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "Ad131CmDkSource.zip"
.CancelError = True
.Filter = "VI (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check53_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\DETECTORS\Ad150\AD150CMDK\Source\Ad150CmDk.zip"
Else
File1 = App.Path & "\CDSoftware\DETECTORS\Ad150\AD150CMDK\Source\Ad150CmDk.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "Ad150CmDk.zip"
.CancelError = True
.Filter = "VI (*.llb)|*.llb"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:

End Sub

Private Sub Check55_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\DETECTORS\AD131\exe\AD131DK\CMDKAD131.doc"
Else
File1 = App.Path & "\CDSoftware\DETECTORS\AD131\exe\AD131DK\CMDKAD131.doc"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "CMDKAD131.doc"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt|PDF (*.pdf)| *.pdf"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check56_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\DETECTORS\SM302\Setup\disks\SM302CM.doc"
Else
File1 = App.Path & "\CDSoftware\DETECTORS\SM302\Setup\disks\SM302CM.doc"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "SM302CM.doc"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt|PDF (*.pdf)| *.pdf"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check57_Click()
retval = Shell("CDSoftware\MONCHROMATORS\LabView\Executable\CM\CMGpib\SETUP.EXE", 1)
End Sub

Private Sub Check58_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\LabView\Executable\CM\CMGpib\8-2012-a.doc"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\LabView\Executable\CM\CMGpib\8-2012-a.doc"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "8-2012-a.doc"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt|PDF (*.pdf)| *.pdf"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check59_Click()
retval = Shell("CDSoftware\MONCHROMATORS\Visual C++\Exe\DK\SETUP.exe", 1)
End Sub

Private Sub Check6_Click()
retval = Shell("CDSoftware\FILTER WHEELS\Visual Basic\Executable\Setup\SetUp.exe", 1)
End Sub


Private Sub Check60_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\Visual C++\Source\DK\DKVCSource.zip"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\Visual C++\Source\DK\DKVCSource.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "DKVCSource.zip"
.CancelError = True
.Filter = "Documents (*.zip)|*.zip|Text (*.txt)|*.txt|PDF (*.pdf)| *.pdf (*.doc)|*.doc"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub



Private Sub Check61_Click()
  If Right$(App.Path, 1) = "\" Then
    On Error GoTo DialogError
    File1 = App.Path & "CDSoftware\MONCHROMATORS\InstManuals\CmSeries\CmManual.pdf"
  Else
    File1 = App.Path & "\CDSoftware\MONCHROMATORS\InstManuals\CmSeries\CmManual.pdf"
  End If
  On Error GoTo DialogError
  With CommonDialog2
    .FileName = "CmManual.pdf"
    .CancelError = True
    .Filter = "PDF File (*.pdf)|*.pdf|Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
    .FilterIndex = 1
    .DialogTitle = "Select a Directory to Save"
    .ShowSave
  End With
  File2 = CommonDialog2.FileName
  FileCopy File1, File2
DialogError:

End Sub

Private Sub Check62_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\Detectors\SM302\Source\SM302Source.zip"
Else
File1 = App.Path & "\CDSoftware\Detectors\SM302\Source\SM302Source.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "SM302Source.zip"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt|PDF (*.pdf)| *.pdf (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check64_Click()
  If Right$(App.Path, 1) = "\" Then
    On Error GoTo DialogError
    File1 = App.Path & "CDSoftware\MONCHROMATORS\InstManuals\DkSeries\DkManual.pdf"
  Else
    File1 = App.Path & "\CDSoftware\MONCHROMATORS\InstManuals\DkSeries\DkManual.pdf"
  End If
  On Error GoTo DialogError
  With CommonDialog2
    .FileName = "DkManual.pdf"
    .CancelError = True
    .Filter = "PDF File (*.pdf)|*.pdf|Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
    .FilterIndex = 1
    .DialogTitle = "Select a Directory to Save"
    .ShowSave
  End With
  File2 = CommonDialog2.FileName
  FileCopy File1, File2
DialogError:

End Sub

Private Sub Check65_Click()
  If Right$(App.Path, 1) = "\" Then
    On Error GoTo DialogError
    File1 = App.Path & "CDSoftware\Spectrometers\InstManuals\SM302\SM302InGaAs.pdf"
  Else
    File1 = App.Path & "\CDSoftware\Spectrometers\InstManuals\SM302\SM302InGaAs.pdf"
  End If
  On Error GoTo DialogError
  With CommonDialog2
    .FileName = "SM302InGaAs.pdf"
    .CancelError = True
    .Filter = "PDF File (*.pdf)|*.pdf|Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
    .FilterIndex = 1
    .DialogTitle = "Select a Directory to Save"
    .ShowSave
  End With
  File2 = CommonDialog2.FileName
  FileCopy File1, File2
DialogError:

End Sub

Private Sub Check69_Click()
  If Right$(App.Path, 1) = "\" Then
    On Error GoTo DialogError
    File1 = App.Path & "CDSoftware\Light Source\InstManuals\Xenon\CM AST-XE-175EX mounting.pdf"
  Else
    File1 = App.Path & "\CDSoftware\Light Source\InstManuals\Xenon\CM AST-XE-175EX mounting.pdf"
  End If
  On Error GoTo DialogError
  With CommonDialog2
    .FileName = "CM AST-XE-175EX mounting.pdf"
    .CancelError = True
    .Filter = "PDF File (*.pdf)|*.pdf|Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
    .FilterIndex = 1
    .DialogTitle = "Select a Directory to Save"
    .ShowSave
  End With
  File2 = CommonDialog2.FileName
  FileCopy File1, File2
DialogError:

End Sub

Private Sub Check7_Click()
    retval = Shell("CDSoftware\DETECTORS\AD111\SetUp\disks\SetUp.exe", 1)
End Sub

Private Sub Check71_Click()
  If Right$(App.Path, 1) = "\" Then
    On Error GoTo DialogError
    File1 = App.Path & "CDSoftware\FILTER WHEELS\InstManuals\AB300 Series\AB300Series.pdf"
  Else
    File1 = App.Path & "\CDSoftware\FILTER WHEELS\InstManuals\AB300 Series\AB300Series.pdf"
  End If
  On Error GoTo DialogError
  With CommonDialog2
    .FileName = "AB300Series.pdf"
    .CancelError = True
    .Filter = "PDF File (*.pdf)|*.pdf|Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
    .FilterIndex = 1
    .DialogTitle = "Select a Directory to Save"
    .ShowSave
  End With
  File2 = CommonDialog2.FileName
  FileCopy File1, File2
DialogError:

End Sub

Private Sub Check73_Click()
  If Right$(App.Path, 1) = "\" Then
    On Error GoTo DialogError
    File1 = App.Path & "CDSoftware\FILTER WHEELS\InstManuals\AB300 T\AB300-T.pdf"
  Else
    File1 = App.Path & "\CDSoftware\FILTER WHEELS\InstManuals\AB300 T\AB300-T.pdf"
  End If
  On Error GoTo DialogError
  With CommonDialog2
    .FileName = "AB300-T.pdf"
    .CancelError = True
    .Filter = "PDF File (*.pdf)|*.pdf|Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
    .FilterIndex = 1
    .DialogTitle = "Select a Directory to Save"
    .ShowSave
  End With
  File2 = CommonDialog2.FileName
  FileCopy File1, File2
DialogError:

End Sub

Private Sub Check74_Click()
  If Right$(App.Path, 1) = "\" Then
    On Error GoTo DialogError
    File1 = App.Path & "CDSoftware\DETECTORS\InstManuals\AD131\AD131Manual.pdf"
  Else
    File1 = App.Path & "\CDSoftware\DETECTORS\InstManuals\AD131\AD131Manual.pdf"
  End If
  On Error GoTo DialogError
  With CommonDialog2
    .FileName = "AD131Manual.pdf"
    .CancelError = True
    .Filter = "PDF File (*.pdf)|*.pdf|Documents (*.doc)|*.doc|Text (*.txt)|*.txt"
    .FilterIndex = 1
    .DialogTitle = "Select a Directory to Save"
    .ShowSave
  End With
  File2 = CommonDialog2.FileName
  FileCopy File1, File2
DialogError:

End Sub

Private Sub Check75_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\LabView\Source\CM GPIBvi\CMGpibSource.zip"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\LabView\Source\CM Gpibvi\CMGpibSource.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "CMGpibSource.zip"
.CancelError = True
.Filter = "Zipped (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:

End Sub

Private Sub Check76_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\Visual C++\Source\CM\CMC++Source.zip"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\Visual C++\Source\CM\CMC++Source.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "CMC++Source.zip"
.CancelError = True
.Filter = "Zipped (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check77_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\LabView\Source\DK GPIBvi\DKGpibSource.zip"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\LabView\Source\DK GPIBvi\DKGpibSource.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "DKGpibSource.zip"
.CancelError = True
.Filter = "Zipped (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check78_Click()
retval = Shell("CDSoftware\MONCHROMATORS\Visual C++\Exe\CM\SETUP.EXE", 1)
End Sub

Private Sub Check79_Click()
retval = Shell("CDSoftware\MONCHROMATORS\LabView\Executable\DK\DK-Rs232\SETUP.exe", 1)
End Sub

Private Sub Check8_Click()

If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\Detectors\AD111\Source\AD111Source.zip"
Else
File1 = App.Path & "\CDSoftware\Detectors\AD111\Source\AD111Source.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "AD111Source.zip"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt|PDF (*.pdf)| *.pdf (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:

End Sub

Private Sub Check80_Click()
retval = Shell("CDSoftware\MONCHROMATORS\LabView\Executable\DK\DK-Gpib\SETUP.exe", 1)
End Sub

Private Sub Check82_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\LabView\Source\CM ICS Gpibvi\CmIcsGpSource.zip"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\LabView\Source\CM ICS Gpibvi\CmIcsGpSource.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "CmIcsGpSource.zip"
.CancelError = True
.Filter = "Zipped (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:


End Sub

Private Sub Check83_Click()
retval = Shell("CDSoftware\MONCHROMATORS\LabView\Executable\CM\CmIcsGpib\SETUP.EXE", 1)
End Sub
Private Sub Check85_Click()

End Sub

Private Sub Check86_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\Visual Basic\Source\CMVBSource.zip"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\Visual Basic\Source\CMVBSource.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "CMVBSource.zip"
.CancelError = True
.Filter = "Documents (*.zip)|*.zip|Text (*.txt)|*.txt|PDF (*.pdf)| *.pdf (*.doc)|*.doc"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check87_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\MONCHROMATORS\Visual Basic\Source\DKVBSource.zip"
Else
File1 = App.Path & "\CDSoftware\MONCHROMATORS\Visual Basic\Source\DKVBSource.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "DKVBSource.zip"
.CancelError = True
.Filter = "Documents (*.zip)|*.zip|Text (*.txt)|*.txt|PDF (*.pdf)| *.pdf (*.doc)|*.doc"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:

End Sub

Private Sub Check88_Click()
If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\FILTER WHEELS\Visual Basic\Source\VB6AbVE.zip"
Else
File1 = App.Path & "\CDSoftware\FILTER WHEELS\Visual Basic\Source\VB6AbVE.zip"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "VB6AbVE.zip"
.CancelError = True
.Filter = "VI (*.zip)|*.zip"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:
End Sub

Private Sub Check9_Click()

If Right$(App.Path, 1) = "\" Then
On Error GoTo DialogError
File1 = App.Path & "CDSoftware\DETECTORS\InstManuals\AD111\AD111SwManual.doc"
Else
File1 = App.Path & "\CDSoftware\DETECTORS\InstManuals\AD111\AD111SwManual.doc"
End If
On Error GoTo DialogError
With CommonDialog2
.FileName = "AD111SwManual.doc"
.CancelError = True
.Filter = "Documents (*.doc)|*.doc|Text (*.txt)|*.txt|PDF (*.pdf)| *.pdf"
.FilterIndex = 1
.DialogTitle = "Select a Directory to Save"
.ShowSave
End With
File2 = CommonDialog2.FileName
FileCopy File1, File2
DialogError:

End Sub

Private Sub Command2_Click()
End
End Sub

