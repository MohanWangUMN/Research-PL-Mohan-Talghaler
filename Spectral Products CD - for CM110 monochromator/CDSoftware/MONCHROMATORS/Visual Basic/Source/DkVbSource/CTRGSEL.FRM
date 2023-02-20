VERSION 5.00
Begin VB.Form ctrGrtsel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grating Select Command"
   ClientHeight    =   1320
   ClientLeft      =   5955
   ClientTop       =   2580
   ClientWidth     =   4575
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
   ScaleWidth      =   4575
   Begin VB.CommandButton cmdCanOK 
      Caption         =   "OK"
      Height          =   315
      Index           =   1
      Left            =   3120
      MaskColor       =   &H00FF0000&
      TabIndex        =   2
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdCanOK 
      Caption         =   "Exit"
      Height          =   315
      Index           =   0
      Left            =   1920
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox pnlCtrSelect 
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   0
      Width           =   4575
      Begin VB.PictureBox pnlNewGrating 
         Height          =   375
         Left            =   480
         ScaleHeight     =   315
         ScaleWidth      =   3555
         TabIndex        =   4
         Top             =   240
         Width           =   3615
         Begin VB.TextBox txtNGrt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   320
            Left            =   1920
            MaxLength       =   1
            MultiLine       =   -1  'True
            TabIndex        =   0
            Text            =   "CTRGSEL.frx":0000
            Top             =   8
            Width           =   1335
         End
         Begin VB.Label lblNewGrating 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New Grating :"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   60
            TabIndex        =   5
            Top             =   65
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "ctrGrtsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCanOK_Click(Index As Integer)
  Dim Status%, NGrt%
  On Error GoTo cmdCanOKErr
  NGrt% = Val(txtNGrt)
  Select Case Index
    Case 0          ' 0 is Cancel button.
      Unload ctrGrtsel
    Case 1          ' 1 is OK button.
      If Len(txtNGrt) Then
        cmdCanOK(1).Enabled = False
        Status% = DKGrtsel%(NGrt%, frmDK.comComm1, 500)
        cmdCanOK(1).Enabled = True
        
        If Status% < 128 Then
          gblNGrtSelted& = NGrt%
        Else
          MsgBox "An error has occurred !", MB_ICONSTOP
        End If
        'Unload ctrGrtsel
      Else
        MsgBox "The entered number is out of range", MB_ICONEXCLAMATION
        txtNGrt.SetFocus
      End If
  End Select

cmdCanOKResume:
  Exit Sub

cmdCanOKErr:
  MsgBox "Error code is : " & Err, MB_ICONEXCLAMATION
  'Unload ctrGrtsel
  Resume cmdCanOKResume

End Sub

Private Sub Form_Load()
  Centerform ctrGrtsel
  txtNGrt.Text = gblNGrtSelted&
  txtNGrt.SelLength = Len(txtNGrt)
End Sub

Private Sub txtNGrt_KeyPress(KeyAscii As Integer)
  Dim DataL&, Status%
  If gblInitCom% = 0 Then         ' If Com has been initialized
    MsgBox "Did you initialize a serial port ? "
  End If
  ' Need to query for number of gratings
  Select Case KeyAscii
    Case 13     ' 13 if ascii of CR
      If Val(txtNGrt) = 0 Then
        'Do nothing
      Else
        cmdCanOK_Click (1)      ' 1: is OK button
      End If
    Case Asc("1") To gblNGrtInstd% + 48 ' Make gblNGrtInstd%+48 like Asc value

    Case 8      ' 8 is ascii of BS
    
    Case Else
      If KeyAscii > gblNGrtInstd% And gblNGrtInstd% Then
        MsgBox "Sorry, number of gratings installed is : " & gblNGrtInstd%, MB_ICONEXCLAMATION, "Select DkSeries Error"
      End If
      KeyAscii = 0
  End Select

End Sub

