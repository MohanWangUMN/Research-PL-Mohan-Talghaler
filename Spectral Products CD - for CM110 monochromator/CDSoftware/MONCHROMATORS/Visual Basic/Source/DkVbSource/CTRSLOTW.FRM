VERSION 5.00
Begin VB.Form ctrSlotW 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SlotWidth Command"
   ClientHeight    =   1455
   ClientLeft      =   1770
   ClientTop       =   7500
   ClientWidth     =   4455
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
   Icon            =   "CTRSLOTW.frx":0000
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "CTRSLOTW.frx":030A
   ScaleHeight     =   1455
   ScaleWidth      =   4455
   Begin VB.CommandButton cmdCanOK 
      Caption         =   "OK"
      Height          =   315
      Index           =   1
      Left            =   2520
      MaskColor       =   &H00FF0000&
      TabIndex        =   1
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdCanOK 
      Caption         =   "Exit"
      Height          =   315
      Index           =   0
      Left            =   1440
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox pnlCtrEcho 
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   0
      Width           =   4455
      Begin VB.PictureBox pnlEcho 
         Height          =   495
         Left            =   600
         ScaleHeight     =   435
         ScaleWidth      =   3195
         TabIndex        =   4
         Top             =   240
         Width           =   3255
         Begin VB.TextBox txtEchoByte 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   20
            Width           =   1335
         End
         Begin VB.Label lblOldWave 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current SlotWidth :"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   60
            TabIndex        =   5
            Top             =   135
            Width           =   1515
         End
      End
   End
End
Attribute VB_Name = "ctrSlotW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCanOK_Click(Index As Integer)
  Dim Status%
  Select Case Index
    Case 0          ' 0 is Cancel button.
      gblOkCancel% = IDCANCEL
      Unload ctrEcho
    Case 1          ' 1 is OK button.
      txtEchoByte = "......."
      Status% = DKEcho%(frmDK.comComm1, 2) '2: 2s Timeout%
      If Status% < 128 Then
        txtEchoByte = "27"
      Else
        txtEchoByte = "Error"
      End If
  End Select
    

End Sub

Private Sub Form_Load()
  Centerform ctrEcho
End Sub

Private Sub txtEchoByte_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then 'Enter button
    cmdCanOK_Click (1)  '1: OK button
  Else
    KeyAscii = 0
  End If
End Sub

