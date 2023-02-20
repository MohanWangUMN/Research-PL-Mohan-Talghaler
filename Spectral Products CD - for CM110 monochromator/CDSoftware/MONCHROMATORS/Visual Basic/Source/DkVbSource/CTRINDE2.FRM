VERSION 5.00
Begin VB.Form ctrIncDecM2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inc. and Dec. Turret #2 Commands"
   ClientHeight    =   1815
   ClientLeft      =   1380
   ClientTop       =   4245
   ClientWidth     =   6855
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
   Icon            =   "CTRINDE2.frx":0000
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1815
   ScaleWidth      =   6855
   Begin VB.PictureBox pnlCtrIncDecM2 
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   6795
      TabIndex        =   2
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdCanOK 
         Caption         =   "Exit"
         Height          =   315
         Index           =   0
         Left            =   3840
         MaskColor       =   &H000000FF&
         TabIndex        =   9
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdCanOK 
         Caption         =   "Reset Step"
         Height          =   315
         Index           =   1
         Left            =   5160
         MaskColor       =   &H00FF0000&
         TabIndex        =   8
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.PictureBox pnlBkGnd 
         BackColor       =   &H00FFFF80&
         Height          =   975
         Index           =   0
         Left            =   480
         ScaleHeight     =   915
         ScaleWidth      =   5835
         TabIndex        =   5
         Top             =   120
         Width           =   5895
         Begin VB.PictureBox picStepM2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   510
            Index           =   1
            Left            =   5040
            Picture         =   "CTRINDE2.frx":030A
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   0
            Top             =   240
            Width           =   510
         End
         Begin VB.PictureBox picStep 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   0
            Left            =   720
            ScaleHeight     =   465
            ScaleWidth      =   4305
            TabIndex        =   6
            Top             =   240
            Width           =   4335
            Begin VB.Label lblWL 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
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
               Left            =   1560
               TabIndex        =   11
               Top             =   120
               Width           =   975
            End
            Begin VB.Label lblStepM2 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
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
               Left            =   120
               TabIndex        =   10
               Top             =   120
               Width           =   975
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
               Left            =   2640
               TabIndex        =   3
               Top             =   120
               Width           =   1605
            End
         End
         Begin VB.PictureBox picStepM2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   510
            Index           =   0
            Left            =   240
            Picture         =   "CTRINDE2.frx":0614
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   1
            Top             =   240
            Width           =   510
         End
         Begin VB.Label lblZeroM1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Step +/-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   4
            Top             =   0
            Width           =   1005
         End
         Begin VB.Label lblZeroM1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Wavelength"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   7
            Top             =   0
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "ctrIncDecM2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modStepFrom0M2%

Private Sub cmdCanOK_Click(Index As Integer)
  Dim Status%
  On Error GoTo cmdCanOKErr
  Select Case Index
    Case 0          ' 0 is Exit (Cancel) button.
      gblOkCancel% = IDCANCEL
      If modStepFrom0M2% Then
        Status% = DKGoto%(gblCurWL&, frmDK.comComm1, 5) '5: 5s Timeout%
      End If
      Unload ctrIncDecM2
    Case 1          ' 1 is Reset Step (OK) button.
      If modStepFrom0M2% Then
        Status% = DKGoto%(gblCurWL&, frmDK.comComm1, 5) '5: 5s Timeout%
        modStepFrom0M2% = 0
        lblStepM2.Caption = modStepFrom0M2%
      End If
  End Select

cmdCanOKResume:
  Exit Sub

cmdCanOKErr:
  MsgBox "Error code is : " & Err, MB_ICONEXCLAMATION
  Resume cmdCanOKResume

End Sub

Private Sub Form_Load()
  Centerform ctrIncDecM2
  lblWL.Caption = gblCurWL&
  modStepFrom0M2% = 0
  lblStepM2.Caption = modStepFrom0M2%

End Sub

Private Sub picStepM2_Click(Index As Integer)
  Dim Status%
  picStepM2(Index).Visible = False
  Select Case Index
    Case 0  'Left arrow
      Status% = DKDecM2%(frmDK.comComm1, 2) '2: 2s Timeout%
      If Status% < 128 Then
        modStepFrom0M2% = modStepFrom0M2% - 1
      Else
        MsgBox "An error has occurred !", MB_ICONSTOP
      End If
    Case 1  'Right arrow
      Status% = DKIncM2%(frmDK.comComm1, 2) '2: 2s Timeout%
      If Status% < 128 Then
        modStepFrom0M2% = modStepFrom0M2% + 1
      Else
        MsgBox "An error has occurred !", MB_ICONSTOP
      End If
  End Select
  picStepM2(Index).Visible = True
  lblStepM2.Caption = modStepFrom0M2%

End Sub

