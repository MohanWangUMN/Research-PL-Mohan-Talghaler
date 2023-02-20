VERSION 5.00
Begin VB.Form baudSearching 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baud Searching"
   ClientHeight    =   2655
   ClientLeft      =   2685
   ClientTop       =   2085
   ClientWidth     =   4350
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
   Icon            =   "BAUDSEAR.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2655
   ScaleWidth      =   4350
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3360
      Picture         =   "BAUDSEAR.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   960
      Width           =   510
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3360
      Picture         =   "BAUDSEAR.frx":0614
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   480
      Width           =   510
   End
   Begin VB.PictureBox pnlSetBaud 
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   4275
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      Begin VB.Frame fraBaud 
         Caption         =   "Baud Rates"
         Height          =   2295
         Left            =   1680
         TabIndex        =   11
         Top             =   120
         Width           =   1215
         Begin VB.OptionButton optBaud 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   12
            Top             =   1950
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optBaud 
            Caption         =   "300"
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
            Index           =   5
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   855
         End
         Begin VB.OptionButton optBaud 
            Caption         =   "600"
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
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Top             =   1410
            Width           =   855
         End
         Begin VB.OptionButton optBaud 
            Caption         =   "1200"
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
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   1140
            Width           =   855
         End
         Begin VB.OptionButton optBaud 
            Caption         =   "2400"
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
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   870
            Width           =   855
         End
         Begin VB.OptionButton optBaud 
            Caption         =   "4800"
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
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton optBaud 
            Caption         =   "9600"
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
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   330
            UseMaskColor    =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame fraCom 
         Caption         =   "COM. Ports"
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1215
         Begin VB.OptionButton optCOM 
            Caption         =   "COM1"
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
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   330
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optCOM 
            Caption         =   "COM2"
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
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton optCOM 
            Caption         =   "COM3"
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
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   870
            Width           =   855
         End
         Begin VB.OptionButton optCOM 
            Caption         =   "COM4"
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
            Index           =   3
            Left            =   120
            TabIndex        =   6
            Top             =   1140
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdCanOK 
         Caption         =   "OK"
         Height          =   315
         Index           =   1
         Left            =   3120
         MaskColor       =   &H00FF0000&
         TabIndex        =   1
         Top             =   2160
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdCanOK 
         Caption         =   "Exit"
         Height          =   315
         Index           =   0
         Left            =   3120
         MaskColor       =   &H000000FF&
         TabIndex        =   0
         Top             =   1800
         UseMaskColor    =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "baudSearching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modBaud%, modPort%

Private Sub cmdCanOK_Click(Index As Integer)
  Dim Status%, Baud%, PortNo%, countBaud%, Selection%
  On Error GoTo cmdCanOKErr
  PortNo% = modPort% + 1
  Select Case Index
    Case 0          ' 0 is Exit (Cancel) button.
      gblOkCancel% = IDCANCEL
      Unload baudSearching
    Case 1          ' 1 is Search (OK) button.
    ' Valid baud rates for CM110/112 are 9600, 4800, 2400, 1200, 600, 300, 150 and 75,
    ' but valid baud rates for Visual Basic are 9600, 4800, 2400, 1200, 600, 300.
    ' I have to use step -1 because if I use step 1, somehow the CM110/112
    ' get into select mode !!!.
      Picture2.Visible = False  'Connected figure
      Picture1.Visible = True   'Disconnected figure
      
      For countBaud% = 5 To 0 Step -1
        baudSearching.optBaud(countBaud%) = True
        Baud% = 9600 / 2 ^ countBaud%
        Status% = DKInitComPort%(PortNo%, Baud%, frmDK.comComm1)
        Status% = DKEcho%(frmDK.comComm1, 1) '2: 2s Timeout%
        If Status% < 128 Then
          gblInitCom% = True
          'MsgBox "The current baud rate is : " & Baud%
          gblPortNo% = modPort%
          gblBaud% = countBaud%
          
          Picture1.Visible = False
          Picture2.Visible = True

          Selection% = MsgBox("Do you want to save the configurarion ?", MB_YESNO)
          If Selection% = IDYES Then
            SaveConfig
          Else
            'If user selects No, then the program does nothing.
          End If

          Exit For
        Else
          Picture1.Visible = True   'Disconnected figure
          Picture2.Visible = False  'Connected figure
          
          gblInitCom% = False
          baudSearching.optBaud(countBaud%) = False
          If countBaud% = 0 Then    '5
            baudSearching.optBaud(6) = True
            'MsgBox "No baud rate found"
          Else
            'MsgBox "Please check connection between PC and CM110/112"
            'Exit For
          End If
        End If
      Next countBaud%
    
    Case Else
      Beep
  End Select

cmdCanOKResume:
  Exit Sub

cmdCanOKErr:
  MsgBox "Error code is : " & Err, MB_ICONEXCLAMATION
  Resume cmdCanOKResume

End Sub

Private Sub Form_Load()
  Centerform baudSearching
  modPort% = gblPortNo%
  optCOM(modPort%) = True
  Picture1.Visible = True
  Picture2.Visible = False

End Sub

Private Sub optCom_Click(Index As Integer)
  Picture1.Visible = True
  Picture2.Visible = False
  modPort% = Index
End Sub

