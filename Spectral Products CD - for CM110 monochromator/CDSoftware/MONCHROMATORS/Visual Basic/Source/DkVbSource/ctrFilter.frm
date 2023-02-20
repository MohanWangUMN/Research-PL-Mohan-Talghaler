VERSION 5.00
Begin VB.Form ctrFilter 
   Caption         =   "Filter Command"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4635
   LinkTopic       =   "Form2"
   ScaleHeight     =   1380
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pnlCtrSpeed 
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdCanOK 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2280
         MaskColor       =   &H000000FF&
         TabIndex        =   5
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdCanOK 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3360
         MaskColor       =   &H00FF0000&
         TabIndex        =   4
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.PictureBox pnlNewSpeed 
         Height          =   375
         Left            =   480
         ScaleHeight     =   315
         ScaleWidth      =   3675
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         Begin VB.TextBox txtFilter 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   2
            Top             =   20
            Width           =   1335
         End
         Begin VB.Label lblNewFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Select New Filter :"
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
            Height          =   255
            Left            =   60
            TabIndex        =   3
            Top             =   75
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "ctrFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCanOK_Click(Index As Integer)
  Dim Status%, Filter%
  On Error GoTo cmdCanOKErr
  Filter% = Val(txtFilter.Text)
  Select Case Index
    Case 0          ' 0 is Cancel button.
      gblOkCancel% = IDCANCEL
      Unload ctrFilter
    Case 1          ' 1 is OK button.
      Status% = DKFilter%(Filter%, frmDK.comComm1, 5) '5: 5s Timeout%
      If Status% < 128 Then gblFilter% = Filter%
      gblOkCancel% = IDOK
  End Select
  

cmdCanOKResume:
  Exit Sub

cmdCanOKErr:
  MsgBox "Error code is : " & Err, MB_ICONEXCLAMATION
  Resume cmdCanOKResume

End Sub

Private Sub Form_Load()
  On Error GoTo ctrFilterErr
  Centerform ctrFilter
  txtFilter.Text = gblFilter%
  txtFilter.SelLength = Len(txtFilter)

ctrFilterResume:
  Exit Sub

ctrFilterErr:
  MsgBox "Error code is : " & Err, MB_ICONEXCLAMATION
  Resume ctrFilterResume

End Sub

