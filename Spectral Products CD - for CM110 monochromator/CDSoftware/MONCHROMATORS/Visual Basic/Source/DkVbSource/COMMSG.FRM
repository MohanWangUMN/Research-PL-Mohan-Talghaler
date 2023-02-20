VERSION 5.00
Begin VB.Form frmComMsg 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   3045
   ClientLeft      =   4170
   ClientTop       =   4920
   ClientWidth     =   4425
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
   HelpContextID   =   105
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   786.821
   ScaleMode       =   0  'User
   ScaleWidth      =   1244.726
   Begin VB.PictureBox Panel3D3 
      Height          =   3015
      Left            =   0
      ScaleHeight     =   2955
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   0
      Width           =   4395
      Begin VB.CommandButton cmdCanOK 
         Caption         =   "OK"
         Height          =   315
         Index           =   1
         Left            =   3240
         MaskColor       =   &H00FF0000&
         TabIndex        =   3
         Top             =   2520
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DK Not Responding!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   120
         Width           =   2925
      End
      Begin VB.Label TextBody 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1635
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   2730
      End
   End
End
Attribute VB_Name = "frmComMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DoneBtn_Click()
  Unload Me
End Sub

Private Sub cmdCanOK_Click(Index As Integer)
  End
End Sub

Private Sub Form_Load()
  Dim Msg$
  Left = (Screen.Width - Width) / 2   ' Center form horizontally.
  Top = (Screen.Height - Height) / 2  ' Center form vertically.
  Msg$ = "The monochromator is not responding. "
  Msg$ = Msg$ + "Please check the port number, the baud rate and the cable connection. "
  Msg$ = Msg$ + "If they are correct, please check the monchromator manual to make "
  Msg$ = Msg$ + "sure the connection between the computer and the monochromator is right."
  TextBody.Caption = Msg$
End Sub

