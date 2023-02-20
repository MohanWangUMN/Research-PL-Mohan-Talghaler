VERSION 5.00
Begin VB.Form ctrReset 
   Caption         =   "Reset Command"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3810
   LinkTopic       =   "Form2"
   ScaleHeight     =   1530
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pnlCtrReset 
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.PictureBox picGauge 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         FillColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         ScaleHeight     =   465
         ScaleWidth      =   3225
         TabIndex        =   3
         Top             =   120
         Width           =   3255
      End
      Begin VB.CommandButton cmdCanOK 
         Caption         =   "OK"
         Height          =   315
         Index           =   1
         Left            =   2520
         MaskColor       =   &H00FF0000&
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "ctrReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCanOK_Click(Index As Integer)
  Dim Status%, I%, J%, PerCenPtr%
  On Error GoTo cmdCanOKErr

  Select Case Index
    Case 0          ' 0 is Cancel button.
      gblOkCancel% = IDCANCEL
      Unload ctrReset
    Case 1          ' 1 is OK button.
        
        picGauge.Cls
        picGauge.CurrentX = (picGauge.ScaleWidth - picGauge.TextWidth("Reading")) \ 2
        picGauge.CurrentY = (picGauge.ScaleHeight - picGauge.TextHeight("Reading")) \ 2
        picGauge.Print "Resetting"
    
      Status% = DKReset%(frmDK.comComm1, 5) '5: 5s Timeout%
          
      I% = 1
      J% = 10         '20
      PerCentPtr% = I% * 100 / J%
      picGauge.Line (0, 0)-((PerCentPtr% * (picGauge.ScaleWidth / 100)), picGauge.ScaleHeight), QBColor(9), BF
      Status% = DKTimeout%(frmDK.comComm1, 15) '5: 5s Timeout%
      
      Do
        Status% = DKEcho%(frmDK.comComm1, 5)
        I% = I% + 1
        If I% > J% Then
            picGauge.Cls
            I% = 0
        End If
        If Status% = 0 Then J% = I%
        
        PerCentPtr% = I% * 100 / J%
        picGauge.Line (0, 0)-((PerCentPtr% * (picGauge.ScaleWidth / 100)), picGauge.ScaleHeight), QBColor(9), BF
      Loop While (Status% <> 0)
            
      gblOkCancel% = IDOK
  End Select
  

cmdCanOKResume:
  Exit Sub

cmdCanOKErr:
  MsgBox "Error code is : " & Err, MB_ICONEXCLAMATION
  Resume cmdCanOKResume

End Sub

Private Sub Form_Load()
  Centerform ctrReset
End Sub
