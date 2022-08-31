VERSION 5.00
Begin VB.Form frmOverride 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Override settings"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3585
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame fmeOverride 
      Caption         =   "Override:"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox chkRandom 
         Caption         =   "Make generation slightly more random"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Note: This could be a good thing or bad thing depending on randomness"
         Top             =   960
         Width           =   3135
      End
      Begin VB.CheckBox chk512 
         Caption         =   "Allow passwords up to 512 character(s)"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox chk256 
         Caption         =   "Allow passwords up to 256 character(s)"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin VB.CheckBox chk128 
         Caption         =   "Allow passwords up to 128 character(s)"
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chk128_Click()
  If chk256.Value = 0 And chk512.Value = 0 Then
    frmPassGen.maxPassLen = 128
  Else
    If chk128.Value = 1 Then
      MsgBox "Cannot set as already overridden to 256 or 512 characters checked.", vbExclamation, "Error"
    End If
    chk128.Value = 0
  End If
End Sub

Private Sub chk256_Click()
  If chk128.Value = 0 And chk512.Value = 0 Then
    frmPassGen.maxPassLen = 256
  Else
    If chk256.Value = 1 Then
      MsgBox "Cannot set as already overridden to 128 or 512 characters checked.", vbExclamation, "Error"
    End If
    chk256.Value = 0
  End If
End Sub

Private Sub chk512_Click()
  If chk128.Value = 0 And chk256.Value = 0 Then
    frmPassGen.maxPassLen = 512
  Else
    If chk512.Value = 1 Then
      MsgBox "Cannot set as already overridden to 128 or 256 characters checked.", vbExclamation, "Error"
    End If
    chk512.Value = 0
  End If
End Sub

Private Sub chkRandom_Click()
  frmPassGen.moreRandomness = chkRandom.Value
End Sub

Private Sub cmdOkay_Click()
  If chk128.Value = 0 And chk256.Value = 0 And chk512.Value = 0 And frmPassGen.maxPassLen <> frmPassGen.defaultMaxPassLen Then
    frmPassGen.maxPassLen = frmPassGen.defaultMaxPassLen
  End If
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
  chk128.Value = IIf(frmPassGen.maxPassLen = 128, 1, 0)
  chk256.Value = IIf(frmPassGen.maxPassLen = 256, 1, 0)
  chk512.Value = IIf(frmPassGen.maxPassLen = 512, 1, 0)
  chkRandom.Value = IIf(frmPassGen.moreRandomness = 1, 1, 0)
End Sub

Private Sub Form_Terminate()
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call frmPassGen.addLenNumbers
End Sub

' EOF
