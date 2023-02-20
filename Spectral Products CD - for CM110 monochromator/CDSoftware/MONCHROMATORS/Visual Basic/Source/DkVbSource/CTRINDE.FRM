VERSION 5.00
Begin VB.Form ctrIncDec 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inc. and Dec. Commands"
   ClientHeight    =   1815
   ClientLeft      =   105
   ClientTop       =   4260
   ClientWidth     =   6615
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
   Icon            =   "CTRINDE.frx":0000
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1815
   ScaleWidth      =   6615
   Begin VB.CommandButton cmdCanOK 
      Caption         =   "Reset Step"
      Height          =   315
      Index           =   1
      Left            =   5040
      MaskColor       =   &H00FF0000&
      TabIndex        =   1
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanOK 
      Caption         =   "Exit"
      Height          =   315
      Index           =   0
      Left            =   3720
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox pnlCtrIncDec 
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   6555
      TabIndex        =   4
      Top             =   0
      Width           =   6615
      Begin VB.PictureBox pnlBkGnd 
         BackColor       =   &H00FFFF80&
         Height          =   975
         Index           =   0
         Left            =   480
         ScaleHeight     =   915
         ScaleWidth      =   5715
         TabIndex        =   7
         Top             =   120
         Width           =   5775
         Begin VB.PictureBox picStepM1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   510
            Index           =   1
            Left            =   4920
            Picture         =   "CTRINDE.frx":030A
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   2
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
            ScaleWidth      =   4185
            TabIndex        =   8
            Top             =   240
            Width           =   4215
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
            Begin VB.Label lblStepM1 
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
               TabIndex        =   5
               Top             =   120
               Width           =   1485
            End
         End
         Begin VB.PictureBox picStepM1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   510
            Index           =   0
            Left            =   240
            Picture         =   "CTRINDE.frx":0614
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   3
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
            TabIndex        =   6
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
            TabIndex        =   9
            Top             =   0
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "ctrIncDec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modStepFrom0M1%

Private Sub cmdCanOK_Click(Index As Integer)
  Dim Status%
  On Error GoTo cmdCanOKErr
  Select Case Index
    Case 0          ' 0 is Exit (Cancel) button.
      gblOkCancel% = IDCANCEL
      If modStepFrom0M1% Then
        Status% = DKGoto%(gblCurWL&, frmDK.comComm1, 5) '5: 5s Timeout%
      End If
      Unload ctrIncDec
    Case 1          ' 1 is Reset Step (OK) button.
      If modStepFrom0M1% Then
        Status% = DKGoto%(gblCurWL&, frmDK.comComm1, 5) '5: 5s Timeout%
        modStepFrom0M1% = 0
        lblStepM1.Caption = modStepFrom0M1%
      End If
  End Select

cmdCanOKResume:
  Exit Sub

cmdCanOKErr:
  MsgBox "Error code is : " & Err, MB_ICONEXCLAMATION
  Resume cmdCanOKResume

End Sub

Private Sub Form_Load()
  Centerform ctrIncDec
  lblWL.Caption = gblCurWL&
  modStepFrom0M1% = 0
  lblStepM1.Caption = modStepFrom0M1%

End Sub

Private Sub picStepM1_Click(Index As Integer)
  Dim Status%
  picStepM1(Index).Visible = False
  Select Case Index
    Case 0  'Left arrow
      Status% = DKDecM1%(frmDK.comComm1, 2) '2: 2s Timeout%
      If Status% < 128 Then
        modStepFrom0M1% = modStepFrom0M1% - 1
      Else
        MsgBox "An error has occurred !", MB_ICONSTOP
      End If
    Case 1  'Right arrow
      Status% = DKIncM1%(frmDK.comComm1, 2) '2: 2s Timeout%
      If Status% < 128 Then
        modStepFrom0M1% = modStepFrom0M1% + 1
      Else
        MsgBox "An error has occurred !", MB_ICONSTOP
      End If
  End Select
  
  picStepM1(Index).Visible = True
  lblStepM1.Caption = modStepFrom0M1%

End Sub

Private Sub pnlStepM1_Click()

End Sub
