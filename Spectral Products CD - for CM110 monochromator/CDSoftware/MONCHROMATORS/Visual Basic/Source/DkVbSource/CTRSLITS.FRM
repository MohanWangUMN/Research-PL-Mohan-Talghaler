VERSION 5.00
Begin VB.Form ctrSlits 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slit Adjustment Commands"
   ClientHeight    =   2640
   ClientLeft      =   5460
   ClientTop       =   4905
   ClientWidth     =   4920
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2640
   ScaleWidth      =   4920
   Begin VB.CommandButton cmdCanOK 
      Caption         =   "OK"
      Height          =   315
      Index           =   1
      Left            =   3600
      MaskColor       =   &H00FF0000&
      TabIndex        =   1
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdCanOK 
      Caption         =   "Exit"
      Height          =   315
      Index           =   0
      Left            =   2520
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox pnlSetSlits 
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   0
      Width           =   4935
      Begin VB.Frame fraSlits 
         Caption         =   "SLITS ADJ."
         Height          =   1455
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1335
         Begin VB.OptionButton optSltAdj 
            Caption         =   "S3"
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
            TabIndex        =   23
            Top             =   1140
            Width           =   855
         End
         Begin VB.OptionButton optSltAdj 
            Caption         =   "S2"
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
            TabIndex        =   22
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optSltAdj 
            Caption         =   "S1"
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
            TabIndex        =   21
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton optSltAdj 
            Caption         =   "SLITS"
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
            TabIndex        =   20
            Top             =   330
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.PictureBox pnlSW 
         Height          =   375
         Index           =   0
         Left            =   1800
         ScaleHeight     =   315
         ScaleWidth      =   2715
         TabIndex        =   3
         Top             =   360
         Width           =   2775
         Begin VB.TextBox txtSW 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   4
            Text            =   "0"
            Top             =   35
            Width           =   855
         End
         Begin VB.Label lblSW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Slits Width:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   100
            TabIndex        =   6
            Top             =   75
            Width           =   1095
         End
         Begin VB.Label lblUnitSW 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "um"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   5
            Top             =   45
            Width           =   615
         End
      End
      Begin VB.PictureBox pnlSW 
         Height          =   375
         Index           =   1
         Left            =   1800
         ScaleHeight     =   315
         ScaleWidth      =   2715
         TabIndex        =   7
         Top             =   720
         Width           =   2775
         Begin VB.TextBox txtSW 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   8
            Text            =   "0"
            Top             =   35
            Width           =   855
         End
         Begin VB.Label lblUnitSW 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "um"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   10
            Top             =   45
            Width           =   615
         End
         Begin VB.Label lblSW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "S1 Width:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   100
            TabIndex        =   9
            Top             =   75
            Width           =   1095
         End
      End
      Begin VB.PictureBox pnlSW 
         Height          =   375
         Index           =   2
         Left            =   1800
         ScaleHeight     =   315
         ScaleWidth      =   2715
         TabIndex        =   11
         Top             =   1080
         Width           =   2775
         Begin VB.TextBox txtSW 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   12
            Text            =   "0"
            Top             =   35
            Width           =   855
         End
         Begin VB.Label lblSW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "S2 Width:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   100
            TabIndex        =   14
            Top             =   75
            Width           =   1095
         End
         Begin VB.Label lblUnitSW 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "um"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   13
            Top             =   45
            Width           =   615
         End
      End
      Begin VB.PictureBox pnlSW 
         Height          =   375
         Index           =   3
         Left            =   1800
         ScaleHeight     =   315
         ScaleWidth      =   2715
         TabIndex        =   15
         Top             =   1440
         Width           =   2775
         Begin VB.TextBox txtSW 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   16
            Text            =   "0"
            Top             =   35
            Width           =   855
         End
         Begin VB.Label lblSW 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "S3 Width:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   100
            TabIndex        =   18
            Top             =   75
            Width           =   1095
         End
         Begin VB.Label lblUnitSW 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "um"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   2040
            TabIndex        =   17
            Top             =   45
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "ctrSlits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modSlitsN%

Private Sub cmdCanOK_Click(Index As Integer)
  Dim Status%, SlitByte%, SlitWidth%
  On Error GoTo cmdCanOKErr
  Select Case Index
    Case 0          ' 0 is Cancel button.
      gblOkCancel% = IDCANCEL
      Unload ctrSlits
    Case 1          ' 1 is OK button.
      Select Case modSlitsN%
        Case 0
          SlitByte% = 14  ' Slit adjustment
        Case 1
          SlitByte% = 31  ' S1 adjustment.
        Case 2
          SlitByte% = 32  ' S2 adjustment.
        Case 3
          SlitByte% = 34  ' S3 adjustment.
      End Select
      SlitWidth% = txtSW(modSlitsN%).Text
      Status% = DKSltAdjCal%(SlitByte%, SlitWidth%, frmDK.comComm1, 500)
      If Status% > 127 Then GoTo cmdCanOKErr
      Select Case modSlitsN%
        Case 0
          ' Slit adjustment
          gblS1% = SlitWidth%  ' S1 adjustment.
          gblS2% = SlitWidth%  ' S1 adjustment.
          txtSW(0).Text = SlitWidth%
          txtSW(1).Text = SlitWidth%
          txtSW(2).Text = SlitWidth%
          If (gblNovAdd29& And &H2) Then
            gblS3% = SlitWidth%
            txtSW(3).Text = SlitWidth%
          End If
        Case 1
          gblS1% = SlitWidth%  ' S1 adjustment.
          txtSW(1).Text = SlitWidth%
        Case 2
          gblS2% = SlitWidth%  ' S2 adjustment.
          txtSW(0).Text = SlitWidth%
          txtSW(2).Text = SlitWidth%
        Case 3
          gblS3% = SlitWidth%  ' S3 adjustment.
          txtSW(3).Text = SlitWidth%
      End Select
      gblOkCancel% = IDOK
  End Select
  
cmdCanOKResume:
  Exit Sub

cmdCanOKErr:
  MsgBox "Error code is : " & Err, MB_ICONEXCLAMATION
  Resume cmdCanOKResume

End Sub

Private Sub Form_Load()
  Centerform ctrSlits
  txtSW(0) = gblS2%     ' All slits set at exit slitwidth.
  txtSW(1) = gblS1%     ' Entrance slit width.
  txtSW(2) = gblS2%     ' Entrance slit width.
  txtSW(3) = gblS3%     ' Entrance slit width.
  
  ctrSlits.optSltAdj(0) = True
    txtSW(1).Enabled = False
    txtSW(2).Enabled = False
    txtSW(3).Enabled = False
  
  If (gblNovAdd29& And &H2) Then
    ' For DK242
    
  Else
    optSltAdj(3).Visible = False
    pnlSW(3).Visible = False
  End If
  
  
End Sub





Private Sub optSltAdj_Click(Index As Integer)
  Dim I%
  modSlitsN% = Index
  For I% = 0 To 3
    txtSW(I%).Enabled = False
  Next I
  txtSW(Index).Enabled = True
  
End Sub


Private Sub txtSW_KeyPress(Index As Integer, KeyAscii As Integer)
 Select Case KeyAscii
    Case 13     ' 13 if ascii of CR
      cmdCanOK_Click (1)      ' 1: is OK button
    Case Asc("0") To Asc("9")

    Case 8      ' 8 is ascii of BS
    
    Case Else
      KeyAscii = 0
  End Select
End Sub


