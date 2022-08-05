VERSION 5.00
Begin VB.Form frmSpecial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Include Special characters"
   ClientHeight    =   3195
   ClientLeft      =   450
   ClientTop       =   540
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdIncludeAll 
      Caption         =   "&Include All"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   34
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame fmeInclude 
      Caption         =   "Include:"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox chkUnderscore 
         Caption         =   "_"
         Height          =   255
         Left            =   3120
         TabIndex        =   32
         Top             =   1200
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkQuote 
         Caption         =   """"
         Height          =   255
         Left            =   3120
         TabIndex        =   31
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkRightBrace 
         Caption         =   "}"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   720
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkLeftBrace 
         Caption         =   "{"
         Height          =   255
         Left            =   3120
         TabIndex        =   29
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkRightBracket 
         Caption         =   "]"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkLeftBracket 
         Caption         =   "["
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   2160
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkRightParent 
         Caption         =   ")"
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   1920
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkLeftParent 
         Caption         =   "("
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   1680
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPipe 
         Caption         =   "|"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   1440
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkComma 
         Caption         =   ","
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   1200
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPercent 
         Caption         =   "%"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkDollar 
         Caption         =   "$"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   720
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkHash 
         Caption         =   "#"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkTilde 
         Caption         =   "~"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkGrave 
         Caption         =   "`"
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   2160
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPower 
         Caption         =   "^"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   1920
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkBackSlash 
         Caption         =   "\"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   1680
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkAtSign 
         Caption         =   "@"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   1440
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkQuestion 
         Caption         =   "?"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   1200
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkGreaterThan 
         Caption         =   ">"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkEquals 
         Caption         =   "="
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkLessThan 
         Caption         =   "<"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkSemiColon 
         Caption         =   ";"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkColon 
         Caption         =   ":"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkForwardSlash 
         Caption         =   "/"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPeriod 
         Caption         =   "."
         Height          =   255
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Help: Period"
         Top             =   1680
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkMinus 
         Caption         =   "-"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Help: Minus"
         Top             =   1440
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkAsterisk 
         Caption         =   "*"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkApostrophe 
         Caption         =   "'"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkAmpersand 
         Caption         =   "&&"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkExclamation 
         Caption         =   "!"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSpecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAmpersand_Click()
  frmPassGen.useAmpersand = chkAmpersand.Value
End Sub

Private Sub chkApostrophe_Click()
   frmPassGen.useApostrophe = chkApostrophe.Value
End Sub

Private Sub chkAsterisk_Click()
  frmPassGen.useAsterisk = chkAsterisk.Value
End Sub

Private Sub chkAtSign_Click()
  frmPassGen.useAtSign = chkAtSign.Value
End Sub

Private Sub chkBackSlash_Click()
  frmPassGen.useBackSlash = chkBackSlash.Value
End Sub

Private Sub chkColon_Click()
  frmPassGen.useColon = chkColon.Value
End Sub

Private Sub chkComma_Click()
  frmPassGen.useComma = chkComma.Value
End Sub

Private Sub chkDollar_Click()
  frmPassGen.useDollar = chkDollar.Value
End Sub

Private Sub chkEquals_Click()
  frmPassGen.useEquals = chkEquals.Value
End Sub

Private Sub chkExclamation_Click()
  frmPassGen.useExclamation = chkExclamation.Value
End Sub

Private Sub chkForwardSlash_Click()
  frmPassGen.useForwardSlash = chkForwardSlash.Value
End Sub

Private Sub chkGrave_Click()
  frmPassGen.useGrave = chkGrave.Value
End Sub

Private Sub chkGreaterThan_Click()
  frmPassGen.useGreaterThan = chkGreaterThan.Value
End Sub

Private Sub chkHash_Click()
  frmPassGen.useHash = chkHash.Value
End Sub

Private Sub chkLeftBrace_Click()
  frmPassGen.useLeftBrace = chkLeftBrace.Value
End Sub

Private Sub chkLeftBracket_Click()
  frmPassGen.useLeftBracket = chkLeftBracket.Value
End Sub

Private Sub chkLeftParent_Click()
  frmPassGen.useLeftParent = chkLeftParent.Value
End Sub

Private Sub chkLessThan_Click()
  frmPassGen.useLessThan = chkLessThan.Value
End Sub

Private Sub chkMinus_Click()
  frmPassGen.useMinus = chkMinus.Value
End Sub

Private Sub chkPercent_Click()
  frmPassGen.usePercent = chkPercent.Value
End Sub

Private Sub chkPeriod_Click()
  frmPassGen.usePeriod = chkPeriod.Value
End Sub

Private Sub chkPipe_Click()
  frmPassGen.usePipe = chkPipe.Value
End Sub

Private Sub chkPlus_Click()
  frmPassGen.usePlus = chkPlus.Value
End Sub

Private Sub chkPower_Click()
  frmPassGen.usePower = chkPower.Value
End Sub

Private Sub chkQuestion_Click()
  frmPassGen.useQuestion = chkQuestion.Value
End Sub

Private Sub chkQuote_Click()
  frmPassGen.useQuote = chkQuote.Value
End Sub

Private Sub chkRightBrace_Click()
  frmPassGen.useRightBrace = chkRightBrace.Value
End Sub

Private Sub chkRightBracket_Click()
  frmPassGen.useRightBracket = chkRightBracket.Value
End Sub

Private Sub chkRightParent_Click()
  frmPassGen.useRightParent = chkRightParent.Value
End Sub

Private Sub chkSemiColon_Click()
  frmPassGen.useSemiColon = chkSemiColon.Value
End Sub

Private Sub chkTilde_Click()
  frmPassGen.useTilde = chkTilde.Value
End Sub

Private Sub chkUnderscore_Click()
  frmPassGen.useUnderscore = chkUnderscore.Value
End Sub

Private Sub cmdIncludeAll_Click()
  chkExclamation.Value = 1
  chkAmpersand.Value = 1
  chkApostrophe.Value = 1
  chkAsterisk.Value = 1
  chkPlus.Value = 1
  chkMinus.Value = 1
  chkPeriod.Value = 1
  chkForwardSlash.Value = 1
  chkColon.Value = 1
  chkSemiColon.Value = 1
  chkLessThan.Value = 1
  chkEquals.Value = 1
  chkGreaterThan.Value = 1
  chkQuestion.Value = 1
  chkAtSign.Value = 1
  chkBackSlash.Value = 1
  chkPower.Value = 1
  chkGrave.Value = 1
  chkTilde.Value = 1
  chkHash.Value = 1
  chkDollar.Value = 1
  chkPercent.Value = 1
  chkComma.Value = 1
  chkPipe.Value = 1
  chkLeftParent.Value = 1
  chkRightParent.Value = 1
  chkLeftBracket.Value = 1
  chkRightBracket.Value = 1
  chkLeftBrace.Value = 1
  chkRightBrace.Value = 1
  chkQuote.Value = 1
  chkUnderscore.Value = 1
End Sub

Private Sub cmdOkay_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  chkExclamation.Value = IIf(frmPassGen.useExclamation = 1, 1, 0)
  chkAmpersand.Value = IIf(frmPassGen.useAmpersand = 1, 1, 0)
  chkApostrophe.Value = IIf(frmPassGen.useApostrophe = 1, 1, 0)
  chkAsterisk.Value = IIf(frmPassGen.useAsterisk = 1, 1, 0)
  chkPlus.Value = IIf(frmPassGen.usePlus = 1, 1, 0)
  chkMinus.Value = IIf(frmPassGen.useMinus = 1, 1, 0)
  chkPeriod.Value = IIf(frmPassGen.usePeriod = 1, 1, 0)
  chkForwardSlash.Value = IIf(frmPassGen.useForwardSlash = 1, 1, 0)
  chkColon.Value = IIf(frmPassGen.useColon = 1, 1, 0)
  chkSemiColon.Value = IIf(frmPassGen.useSemiColon = 1, 1, 0)
  chkLessThan.Value = IIf(frmPassGen.useLessThan = 1, 1, 0)
  chkEquals.Value = IIf(frmPassGen.useEquals = 1, 1, 0)
  chkGreaterThan.Value = IIf(frmPassGen.useGreaterThan = 1, 1, 0)
  chkQuestion.Value = IIf(frmPassGen.useQuestion = 1, 1, 0)
  chkAtSign.Value = IIf(frmPassGen.useAtSign = 1, 1, 0)
  chkBackSlash.Value = IIf(frmPassGen.useBackSlash = 1, 1, 0)
  chkPower.Value = IIf(frmPassGen.usePower = 1, 1, 0)
  chkGrave.Value = IIf(frmPassGen.useGrave = 1, 1, 0)
  chkTilde.Value = IIf(frmPassGen.useTilde = 1, 1, 0)
  chkHash.Value = IIf(frmPassGen.useHash = 1, 1, 0)
  chkDollar.Value = IIf(frmPassGen.useDollar = 1, 1, 0)
  chkPercent.Value = IIf(frmPassGen.usePercent = 1, 1, 0)
  chkComma.Value = IIf(frmPassGen.useComma = 1, 1, 0)
  chkPipe.Value = IIf(frmPassGen.usePipe = 1, 1, 0)
  chkLeftParent.Value = IIf(frmPassGen.useLeftParent = 1, 1, 0)
  chkRightParent.Value = IIf(frmPassGen.useRightParent = 1, 1, 0)
  chkLeftBracket.Value = IIf(frmPassGen.useLeftBracket = 1, 1, 0)
  chkRightBracket.Value = IIf(frmPassGen.useRightBracket = 1, 1, 0)
  chkLeftBrace.Value = IIf(frmPassGen.useLeftBrace = 1, 1, 0)
  chkRightBrace.Value = IIf(frmPassGen.useRightBrace = 1, 1, 0)
  chkQuote.Value = IIf(frmPassGen.useQuote = 1, 1, 0)
  chkUnderscore.Value = IIf(frmPassGen.useUnderscore = 1, 1, 0)
End Sub

Private Sub Form_Terminate()
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmPassGen.makeSpecialString
End Sub
