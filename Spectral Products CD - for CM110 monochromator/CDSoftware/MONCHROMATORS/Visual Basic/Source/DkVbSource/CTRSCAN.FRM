VERSION 5.00
Begin VB.Form ctrScan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Command"
   ClientHeight    =   1920
   ClientLeft      =   1020
   ClientTop       =   3090
   ClientWidth     =   6510
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
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1920
   ScaleWidth      =   6510
   Begin VB.CommandButton cmdCanOK 
      Caption         =   "OK"
      Height          =   315
      Index           =   1
      Left            =   4920
      MaskColor       =   &H00FF0000&
      TabIndex        =   2
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdCanOK 
      Caption         =   "Exit"
      Height          =   315
      Index           =   0
      Left            =   3600
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox pnlCtrScan 
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   6435
      TabIndex        =   5
      Top             =   0
      Width           =   6495
      Begin VB.PictureBox pnlWL 
         Height          =   375
         Index           =   1
         Left            =   480
         ScaleHeight     =   315
         ScaleWidth      =   5355
         TabIndex        =   4
         Top             =   720
         Width           =   5415
         Begin VB.TextBox txtWL 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   2280
            MaxLength       =   8
            TabIndex        =   0
            Top             =   35
            Width           =   1335
         End
         Begin VB.Label lblWL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter End Wavelength :"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   9
            Top             =   75
            Width           =   2235
         End
         Begin VB.Label lblUnitWL 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "tenths of Angstroms"
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
            Left            =   3720
            TabIndex        =   10
            Top             =   45
            Width           =   1605
         End
      End
      Begin VB.PictureBox pnlWL 
         Height          =   375
         Index           =   0
         Left            =   480
         ScaleHeight     =   315
         ScaleWidth      =   5355
         TabIndex        =   6
         Top             =   240
         Width           =   5415
         Begin VB.TextBox txtWL 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2280
            MaxLength       =   8
            TabIndex        =   3
            Top             =   35
            Width           =   1335
         End
         Begin VB.Label lblUnitWL 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "tenths of Angstroms"
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
            Left            =   3720
            TabIndex        =   8
            Top             =   45
            Width           =   1605
         End
         Begin VB.Label lblWL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "The Current Wavelength :"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   7
            Top             =   75
            Width           =   2235
         End
      End
   End
End
Attribute VB_Name = "ctrScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCanOK_Click(Index As Integer)
  Dim Status%, WL1&, WL2&, DataL&
  On Error GoTo cmdCanOKErr
  WL1& = Val(txtWL(0))
  WL2& = Val(txtWL(1))
  Select Case Index
    Case 0          ' 0 is Cancel button.
      gblOkCancel% = IDCANCEL
      Unload ctrScan
    Case 1          ' 1 is OK button.
      If (Len(txtWL(0)) <> 0) And (Len(txtWL(1)) <> 0) Then
        Status% = DKScan%(WL2&, frmDK.comComm1, 500)    ' Timeout%
        If Status% < 128 Then
          gblCurWL& = WL2&
        Else
          MsgBox "An error has occurred !", MB_ICONSTOP
        End If
      End If
      gblOkCancel% = IDOK
  End Select
    'Unload ctrScan

cmdCanOKResume:
  Exit Sub

cmdCanOKErr:
  MsgBox "Error code is : " & Err, MB_ICONEXCLAMATION
  Resume cmdCanOKResume

End Sub

Private Sub Form_Load()
  Centerform ctrScan
  txtWL(0).Text = gblCurWL&     '0: start wavelength.
  txtWL(0).SelLength = Len(txtWL(0))
  txtWL(1).Text = gblCurWL&     '1: end wavelength.

End Sub

Private Sub txtWL_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case KeyAscii
    Case 13     ' 13 if ascii of CR
      If Index = 0 Then     '0: Start wavelength.
        txtWL(1).SetFocus
      Else                  '1: End wavelength.
        cmdCanOK_Click (1)      ' 1: is OK button
      End If
    Case Asc("0") To Asc("9")

    Case 8      ' 8 is ascii of BS
    
    Case Else
      KeyAscii = 0
  End Select

End Sub

