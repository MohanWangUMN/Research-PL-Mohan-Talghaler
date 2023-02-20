VERSION 5.00
Begin VB.Form About 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4080
   ClientLeft      =   3600
   ClientTop       =   4065
   ClientWidth     =   4320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1007.406
   ScaleMode       =   0  'User
   ScaleWidth      =   996.54
   Begin VB.PictureBox Panel3D1 
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4035
      ScaleWidth      =   4275
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton AboutOK 
         Caption         =   "OK"
         Height          =   315
         Left            =   3120
         MaskColor       =   &H00FF0000&
         TabIndex        =   6
         Top             =   3600
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.PictureBox Panel3D2 
         Height          =   1335
         Left            =   240
         ScaleHeight     =   1275
         ScaleWidth      =   3795
         TabIndex        =   4
         Top             =   2160
         Width           =   3855
         Begin VB.Label SysInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "System Info"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   3615
         End
      End
      Begin VB.Label lblDk 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DK-Series Demo Software"
         BeginProperty Font 
            Name            =   "Script"
            Size            =   24
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Version 2.0a"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1485
         TabIndex        =   1
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (c) Spectral Products, LLC., 2001"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   0
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   240
         X2              =   4080
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   240
         X2              =   4080
         Y1              =   2040
         Y2              =   2040
      End
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AboutOK_Click()
    Unload About
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  AboutOK.Value = True
End Sub

Private Sub Form_Load()
  Dim ms As MEMORYSTATUS
  Dim si As SYSTEM_INFO
  Dim FreeSpace As Long
  Dim FreeResources As Integer
  Dim WinFlags As Long
  Dim TruncRev%, FreeSpaceMsg$, FreeResourceMsg$
  Dim WinFlagMsg$, cpu$

  About.KeyPreview = True

  Left = (Screen.Width - Width) / 2   ' Center form horizontally.
  Top = (Screen.Height - Height) / 2  ' Center form vertically.

  TruncRev% = Int(Val(REVISION) / 10)
  Label2 = "Version " + VERSION + "." + Trim$(Str$(TruncRev%))
    
  GlobalMemoryStatus ms
  FreeSpace = ms.dwAvailPhys
  FreeResources = 100 - ms.dwMemoryLoad

  FreeSpaceMsg$ = "Memory:  " + Trim$(Format(Int(FreeSpace / 1024), "###,###,###,###,###")) + " KB Free"
  FreeResourceMsg$ = "System Resources:  " + Trim$(Str$(FreeResources)) + "%"
    
  GetSystemInfo si
  Select Case si.dwProcessorType
    Case PROCESSOR_INTEL_386
      WinFlagMsg$ = "Intel 386"
    Case PROCESSOR_INTEL_486
      WinFlagMsg$ = "Intel 486"
    Case PROCESSOR_INTEL_PENTIUM
      WinFlagMsg$ = "Intel Pentium"
    Case PROCESSOR_MIPS_R4000
      WinFlagMsg$ = "MIPS_R4000"
    Case PROCESSOR_ALPHA_21064
      WinFlagMsg$ = "ALPHA_21064"
  End Select
  
  SysInfo = WinFlagMsg$ + Chr$(10) + FreeSpaceMsg$ + Chr$(10) + cpu$ & Chr$(10) + FreeResourceMsg$
  About.KeyPreview = True

End Sub

