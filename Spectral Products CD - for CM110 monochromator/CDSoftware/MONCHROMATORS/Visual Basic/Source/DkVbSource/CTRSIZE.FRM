VERSION 5.00
Begin VB.Form ctrSize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Size Command"
   ClientHeight    =   1320
   ClientLeft      =   1065
   ClientTop       =   3960
   ClientWidth     =   6630
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
   ScaleHeight     =   1320
   ScaleWidth      =   6630
   Begin VB.PictureBox pnlCtrSize 
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   6555
      TabIndex        =   3
      Top             =   0
      Width           =   6615
      Begin VB.PictureBox pnlSize 
         Height          =   375
         Left            =   480
         ScaleHeight     =   315
         ScaleWidth      =   5595
         TabIndex        =   4
         Top             =   240
         Width           =   5655
         Begin VB.TextBox txtSize 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2520
            MaxLength       =   4
            MultiLine       =   -1  'True
            TabIndex        =   0
            Text            =   "CTRSIZE.frx":0000
            Top             =   30
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
            Left            =   3960
            TabIndex        =   6
            Top             =   45
            Width           =   1605
         End
         Begin VB.Label lblSize 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Size Byte (+/-127) :"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   60
            TabIndex        =   5
            Top             =   75
            Width           =   2355
         End
      End
      Begin VB.PictureBox cmdCanOK 
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   4080
         ScaleHeight     =   255
         ScaleWidth      =   915
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.PictureBox cmdCanOK 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   5160
         ScaleHeight     =   255
         ScaleWidth      =   915
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "ctrSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCanOK_Click(Index As Integer)
  Dim Status%, SizeB%
  On Error GoTo cmdCanOKErr
  SizeB% = Val(txtSize)
  Select Case Index
    Case 0          ' 0 is Cancel button.
      gblOkCancel% = IDCANCEL
      Unload Me
    Case 1          ' 1 is OK button.
      If Len(txtSize) Then
        Status% = DKSize%(SizeB%, frmDK.comComm1, 5) '5: 5s Timeout%
        If Status% < 128 Then
          gblCurSize% = SizeB%
          Unload Me
        Else
          txtSize.SetFocus
        End If
      End If
      gblOkCancel% = IDOK
  End Select

cmdCanOKResume:
  Exit Sub

cmdCanOKErr:
  MsgBox "Error code is : " & Err, MB_ICONEXCLAMATION
  Resume cmdCanOKResume

End Sub

Private Sub Form_Load()
  Dim Size%
  Centerform ctrSize
  If gblCurSize% >= 128 Then
    Size% = 128 - gblCurSize
  Else
    Size% = gblCurSize%
  End If
  txtSize.Text = Size%
  txtSize.SelLength = Len(txtSize)
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
  Dim PosM%, PosP%
  Select Case KeyAscii
    Case 13     ' 13 if ascii of CR
      cmdCanOK_Click (1)      ' 1: is OK button
    Case Asc("-"), Asc("+")
      If InStr(txtSize, "-") Or InStr(txtSize, "+") Then
        KeyAscii = 0
      End If
    Case Asc("0") To Asc("9")

    Case 8      ' 8 is ascii of BS
    
    Case Else
      KeyAscii = 0
  End Select

End Sub

