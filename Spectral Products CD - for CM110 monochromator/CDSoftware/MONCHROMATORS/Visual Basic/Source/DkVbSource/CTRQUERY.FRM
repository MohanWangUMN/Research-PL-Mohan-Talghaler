VERSION 5.00
Begin VB.Form ctrQuery 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query Command"
   ClientHeight    =   3120
   ClientLeft      =   3345
   ClientTop       =   2865
   ClientWidth     =   6975
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
   Icon            =   "CTRQUERY.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3120
   ScaleWidth      =   6975
   Begin VB.PictureBox pnlCtrQuery 
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   6915
      TabIndex        =   3
      Top             =   0
      Width           =   6975
      Begin VB.PictureBox pnlQuery 
         Height          =   1695
         Index           =   2
         Left            =   240
         ScaleHeight     =   1635
         ScaleWidth      =   3075
         TabIndex        =   7
         Top             =   1200
         Width           =   3135
         Begin VB.ListBox lstCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   1395
            Left            =   120
            TabIndex        =   2
            Top             =   150
            Width           =   2895
         End
      End
      Begin VB.PictureBox pnlQuery 
         Height          =   855
         Index           =   1
         Left            =   3720
         ScaleHeight     =   795
         ScaleWidth      =   2475
         TabIndex        =   8
         Top             =   240
         Width           =   2535
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Status Message:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   60
            TabIndex        =   10
            Top             =   75
            Width           =   1755
         End
         Begin VB.Label lblData 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.PictureBox pnlQuery 
         Height          =   855
         Index           =   0
         Left            =   240
         ScaleHeight     =   795
         ScaleWidth      =   3075
         TabIndex        =   4
         Top             =   240
         Width           =   3135
         Begin VB.Label lblCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label lblSelectQuery 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Select Query Command:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   60
            TabIndex        =   5
            Top             =   75
            Width           =   2235
         End
      End
      Begin VB.PictureBox cmdCanOK 
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   0
         Left            =   4080
         ScaleHeight     =   495
         ScaleWidth      =   915
         TabIndex        =   1
         Top             =   1920
         Width           =   975
      End
      Begin VB.PictureBox cmdCanOK 
         ForeColor       =   &H00FF0000&
         Height          =   555
         Index           =   1
         Left            =   5160
         ScaleHeight     =   495
         ScaleWidth      =   915
         TabIndex        =   0
         Top             =   1920
         Width           =   975
      End
   End
End
Attribute VB_Name = "ctrQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QCode%

Private Sub cmdCanOK_Click(Index As Integer)
  Dim Status%, DataL&
  On Error GoTo cmdCanOKErr
  Select Case Index
    Case 0      ' Cancel button
      Unload ctrQuery
    Case 1      ' OK button
      If gblInitCom% Then         ' If Com has been initialized
        Status% = DKQuery%(QCode%, DataL&, frmDK.comComm1, 2) '2: 2s Timeout%
        If Status% < 128 Then
          lblData = DataL&
        Else
          MsgBox "An error has occurred !", MB_ICONSTOP
        End If
      Else
        MsgBox " Did you initialize a serial Port? ", MB_ICONEXCLAMATION, ""
      End If
    Case Else
        MsgBox " An error has occurred "
  End Select

cmdCanOKResume:
  Exit Sub

cmdCanOKErr:
  MsgBox "Error code is : " & Err, MB_ICONEXCLAMATION
  Resume cmdCanOKResume

End Sub

Private Sub Form_Load()
  Centerform ctrQuery
  lstCode.AddItem "Current WaveLength"  '0
  lstCode.AddItem "Monochromator Type"  '1
  lstCode.AddItem "Ruling Of Cur. Sel. Grtg"      '2
  lstCode.AddItem "Blaze WLength Of Cur. Sel. Grtg" '3
  lstCode.AddItem "# Of Cur. Grtg Selected" '4
  lstCode.AddItem "Current Scan Speed"  '5
  lstCode.AddItem "Current Step Size"   '6
  lstCode.AddItem "Zero Offset M1G1"    '7
  lstCode.AddItem "Zero Offset M1G2"    '8
  lstCode.AddItem "Zero Offset M2G1"    '9
  lstCode.AddItem "Zero Offset M2G2"    '10
  lstCode.AddItem "Calibration const. G1"   '11
  lstCode.AddItem "Calibration const. G2"   '12
  lstCode.AddItem "# Of Grtg Installed" '13
  lstCode.AddItem "Current Units Selected"  '14
  lstCode.AddItem "Grating Ruling G1"   '15
  lstCode.AddItem "Blaze Wavelength G1" '16
  lstCode.AddItem "Grating Ruling G2"   '17
  lstCode.AddItem "Blaze Wavelength G2" '18
  lstCode.AddItem "Serial Number"       '19
  lstCode.Selected(0) = True
  lblCode.Caption = lstCode.Text
End Sub

Private Sub lstCode_Click()
 lblCode.Caption = lstCode.Text
 QCode% = lstCode.ListIndex      ' Query code

End Sub

Private Sub lstCode_KeyPress(KeyAscii As Integer)
  cmdCanOK_Click (1)    '1: OK button.
End Sub

