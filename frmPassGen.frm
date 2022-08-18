VERSION 5.00
Begin VB.Form frmPassGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Generator"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select All"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Timer tmrAutomatic 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   9960
      Top             =   3480
   End
   Begin VB.Timer tmrNoteClear 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   9480
      Top             =   3480
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   10560
      TabIndex        =   15
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame fmeString 
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   3855
      Begin VB.ComboBox cmbPIM 
         Height          =   315
         ItemData        =   "frmPassGen.frx":0000
         Left            =   2760
         List            =   "frmPassGen.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1240
         Width           =   975
      End
      Begin VB.CheckBox chkPIM 
         Caption         =   "Include a random PIM from 1 to "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Note: For when making VeraCrypt file(s)/volume(s)"
         Top             =   1290
         Width           =   3615
      End
      Begin VB.CheckBox chkAutomatic 
         Caption         =   "Automatically generate new password(s) (60s)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   3615
      End
      Begin VB.ComboBox cmbLength 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cmbNumber 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblLength 
         Caption         =   "That are a length of                              characters"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   640
         Width           =   3615
      End
      Begin VB.Label lblNumber 
         Caption         =   "Generate a total of                                password(s)"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   280
         Width           =   3615
      End
   End
   Begin VB.ListBox lstPasswords 
      Height          =   3180
      Left            =   4080
      MultiSelect     =   2  'Extended
      TabIndex        =   14
      Top             =   120
      Width           =   7815
   End
   Begin VB.Frame fmeSettings 
      Caption         =   "Settings:"
      Height          =   1575
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton cmdSpecial 
         Caption         =   "?"
         Height          =   315
         Left            =   3360
         TabIndex        =   4
         Top             =   920
         Width           =   255
      End
      Begin VB.CheckBox chkSpaces 
         Caption         =   "Include Spaces"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   3615
      End
      Begin VB.CheckBox chkSpecialChars 
         Caption         =   "Include Special characters (!, "", #, etc.)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox chkNumChars 
         Caption         =   "Include Numbers"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkUpperChars 
         Caption         =   "Include Uppercase characters"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkLowerChars 
         Caption         =   "Include Lowercase characters"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   3615
      End
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu menuExit 
         Caption         =   "&Exit"
         Index           =   1
      End
   End
   Begin VB.Menu Settings 
      Caption         =   "&Settings"
      Begin VB.Menu Override 
         Caption         =   "&Override"
      End
   End
End
Attribute VB_Name = "frmPassGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ,-> Global variable(s).
Public specialString As String

Public useExclamation As Integer
Public useQuote As Integer
Public useHash As Integer
Public useDollar As Integer
Public usePercent As Integer
Public useAmpersand As Integer
Public useApostrophe As Integer
Public useLeftParenthesis As Integer
Public useRightParenthesis As Integer
Public useAsterisk As Integer
Public usePlus As Integer
Public useComma As Integer
Public useMinus As Integer
Public usePeriod As Integer
Public useForwardSlash As Integer
Public useColon As Integer
Public useSemiColon As Integer
Public useLessThan As Integer
Public useEquals As Integer
Public useGreaterThan As Integer
Public useQuestion As Integer
Public useAtSign As Integer
Public useLeftBracket As Integer
Public useBackSlash As Integer
Public useRightBracket As Integer
Public usePower As Integer
Public useUnderscore As Integer
Public useGrave As Integer
' `-> Note: `
Public useLeftBrace As Integer
Public usePipe As Integer
Public useRightBrace As Integer
Public useTilde As Integer
' `-> ASCII table order.

Public defaultMaxPassLen As Integer
Public defaultPIM As Integer
Public maxPassLen As Integer
Public minPassLen As Integer
Public moreRandomness As Integer

' ,-> Code:

Public Function addLenNumbers()
  If IsNull(cmbLength.Text) = False And cmbLength.Text <> "" Then
    If IsNumeric(cmbLength.Text) = True Then
      If cmbLength.Text > maxPassLen Then
        cmbLength.Tag = "16"
      Else
        cmbLength.Tag = cmbLength.Text
      End If
    Else
      cmbLength.Tag = cmbLength.Text
    End If
  End If
  cmbLength.Clear
  Dim lengthNumber As Integer
  For lengthNumber = minPassLen To maxPassLen
    cmbLength.AddItem lengthNumber
  Next lengthNumber
  cmbLength.AddItem "Rand"
  cmbLength.Text = IIf(IsNull(cmbLength.Tag) = False And cmbLength.Tag <> "", cmbLength.Tag, 16)
  ' `-> The minimum is 8, but nobody should really be using sub 16 character passwords in 2022.
End Function

Private Function makePass(passCount As Integer, inputLength As Variant)
  Randomize
  Dim baseString As String
  If chkUpperChars.Value = 1 Then baseString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  If chkLowerChars.Value = 1 Then baseString = baseString & "abcdefghijklmnopqrstuvwxyz"
  If chkNumChars.Value = 1 Then baseString = baseString & "0123456789"
  If chkSpecialChars.Value = 1 Then
    If IsNull(specialString) = False Or specialString <> "" Then baseString = baseString & specialString
  End If
  If chkSpaces.Value = 1 Then baseString = baseString & Chr(32)
  ' `-> Chr(32) is <SPACE>
  Dim isRand As Integer
  isRand = 0
  If inputLength = "Rand" Then isRand = 1
  lstPasswords.Clear
  Dim makeNumber As Integer, outString As String, randNumber As Integer
  For makeNumber = 0 To passCount - 1
    outString = ""
    Dim lengthNumber As Integer
    If isRand = 1 Then inputLength = Int(Val(minPassLen + Val(Rnd * Val(maxPassLen - Val(minPassLen - 1)))))
    For lengthNumber = 0 To inputLength - 1
      If moreRandomness = 1 Then
        randNumber = Int(Val(1 + Val(Rnd * 6)))
        Select Case randNumber
          Case 1
            outString = outString & Mid(baseString, Int(Val(1 + Val(Rnd * Len(baseString)))), 1)
          Case 2
            outString = outString & Mid(StrReverse(baseString), Int(Val(1 + Val(Rnd * Len(StrReverse(baseString))))), 1)
          Case 3
            ' ,-> Former half
            outString = outString & Mid(Mid(baseString, 1, Val(Len(baseString) / 2)), Int(Val(1 + Val(Rnd * Len(Mid(baseString, 1, Val(Len(baseString) / 2)))))), 1)
          Case 4
            ' ,-> Latter half
            outString = outString & Mid(Mid(baseString, Val(Len(baseString) / 2)), Int(Val(1 + Val(Rnd * Len(Mid(baseString, Val(Len(baseString) / 2)))))), 1)
          Case 5
            ' ,-> Former half (Reverse)
            outString = outString & Mid(Mid(StrReverse(baseString), 1, Val(Len(StrReverse(baseString)) / 2)), Int(Val(1 + Val(Rnd * Len(Mid(StrReverse(baseString), 1, Val(Len(StrReverse(baseString)) / 2)))))), 1)
          Case 6
            ' ,-> Latter half (Reverse)
            outString = outString & Mid(Mid(StrReverse(baseString), Val(Len(StrReverse(baseString)) / 2)), Int(Val(1 + Val(Rnd * Len(Mid(StrReverse(baseString), Val(Len(StrReverse(baseString)) / 2)))))), 1)
        End Select
      Else
        outString = outString & Mid(baseString, Int(Val(1 + Val(Rnd * Len(baseString)))), 1)
      End If
    Next lengthNumber
    If chkPIM.Value = 1 Then
      If Len(outString) <= 64 Then
        ' `-> VeraCrypt length is limited to 64 characters. There is no point to including a PIM if it's longer than that.
        If Len(outString) >= 20 Then
          outString = outString & " ---------- " & Int(Val(1 + Val(Rnd * cmbPIM.Text)))
        Else
          If cmbPIM.Text > defaultPIM Then outString = outString & " ---------- " & Int(Val(defaultPIM + Val(Rnd * Val(cmbPIM.Text - Val(defaultPIM - 1)))))
        End If
      End If
    End If
    lstPasswords.AddItem outString
  Next makeNumber
End Function

Public Function makeSpecialString() As String
  Dim baseString As String
  If useExclamation = 1 Then baseString = "!"
  If useQuote = 1 Then baseString = baseString & Chr(34)
  If useHash = 1 Then baseString = baseString & "#"
  If useDollar = 1 Then baseString = baseString & "$"
  If usePercent = 1 Then baseString = baseString & "%"
  If useAmpersand = 1 Then baseString = baseString & "&"
  If useApostrophe = 1 Then baseString = baseString & "'"
  If useLeftParenthesis = 1 Then baseString = baseString & "("
  If useRightParenthesis = 1 Then baseString = baseString & ")"
  If useAsterisk = 1 Then baseString = baseString & "*"
  If usePlus = 1 Then baseString = baseString & "+"
  If useComma = 1 Then baseString = baseString & ","
  If useMinus = 1 Then baseString = baseString & "-"
  If usePeriod = 1 Then baseString = baseString & "."
  If useForwardSlash = 1 Then baseString = baseString & "/"
  If useColon = 1 Then baseString = baseString & ":"
  If useSemiColon = 1 Then baseString = baseString & ";"
  If useLessThan = 1 Then baseString = baseString & "<"
  If useEquals = 1 Then baseString = baseString & "="
  If useGreaterThan = 1 Then baseString = baseString & ">"
  If useQuestion = 1 Then baseString = baseString & "?"
  If useAtSign = 1 Then baseString = baseString & "@"
  If useLeftBracket = 1 Then baseString = baseString & "["
  If useBackSlash = 1 Then baseString = baseString & "\"
  If useRightBracket = 1 Then baseString = baseString & "]"
  If usePower = 1 Then baseString = baseString & "^"
  If useUnderscore = 1 Then baseString = baseString & "_"
  If useGrave = 1 Then baseString = baseString & "`"
  If useLeftBrace = 1 Then baseString = baseString & "{"
  If usePipe = 1 Then baseString = baseString & "|"
  If useRightBrace = 1 Then baseString = baseString & "}"
  If useTilde = 1 Then baseString = baseString & "~"
  ' `-> ASCII table order.
  specialString = baseString
  If IsNull(baseString) = True Or baseString = "" Then
    chkSpecialChars.Value = 0
  Else
    chkSpecialChars.Value = 1
  End If
End Function

Private Sub chkAutomatic_Click()
  If chkAutomatic.Value = 1 Then
    tmrAutomatic.Enabled = True
  Else
    tmrAutomatic.Enabled = False
  End If
End Sub

Private Sub chkSpecialChars_Click()
  If IsNull(specialString) = True Or specialString = "" And chkSpecialChars.Value = 1 Then
    MsgBox "Cannot enable. You have no special character(s) chosen.", vbExclamation, "Error"
    chkSpecialChars.Value = 0
  End If
End Sub

Private Sub cmdClose_Click()
  End
End Sub

Private Sub cmdCopy_Click()
  On Error GoTo endCopy
    If lstPasswords.SelCount > 0 Then
      Clipboard.Clear
      Dim selectedNumber As Integer
      For selectedNumber = 0 To lstPasswords.ListCount - 1
        If lstPasswords.Selected(selectedNumber) = True Then
          Clipboard.SetText Clipboard.GetText & lstPasswords.List(selectedNumber) & vbNewLine
        End If
      Next
      Me.Caption = Me.Caption & " - Copied to clipboard"
      tmrNoteClear.Enabled = True
    Else
      MsgBox "Please select at least one password to copy.", vbExclamation, "Error"
    End If
    Exit Sub

endCopy:
  MsgBox "Failed to copy to clipboard.", vbExclamation, "Error"
End Sub

Private Sub cmdGenerate_Click()
  If chkUpperChars.Value = 0 And chkLowerChars.Value = 0 And chkNumChars.Value = 0 And chkSpecialChars.Value = 0 Then
    MsgBox "Please select at least one option.", vbExclamation, "Error"
    ' `-> Do not include chkSpace here because it would offer no differences.
  Else
    Call makePass(cmbNumber.Text, cmbLength.Text)
  End If
End Sub

Private Sub cmdSelect_Click()
  If lstPasswords.ListCount > 0 Then
    Dim loopNumber As Integer
    For loopNumber = 0 To lstPasswords.ListCount - 1
      lstPasswords.Selected(loopNumber) = True
    Next
  Else
    MsgBox "There is nothing to select.", vbExclamation, "Error"
  End If
End Sub

Private Sub cmdSpecial_Click()
  frmSpecial.Visible = True
End Sub

Private Sub Form_Load()
  Me.Tag = Me.Caption
  ' ,-> Set the global variable(s).
  useExclamation = 1
  useQuote = 1
  useHash = 1
  useDollar = 1
  usePercent = 1
  useAmpersand = 1
  useApostrophe = 1
  useLeftParenthesis = 1
  useRightParenthesis = 1
  useAsterisk = 1
  usePlus = 1
  useComma = 1
  useMinus = 1
  usePeriod = 1
  useForwardSlash = 1
  useColon = 1
  useSemiColon = 1
  useLessThan = 1
  useEquals = 1
  useGreaterThan = 1
  useQuestion = 1
  useAtSign = 1
  useLeftBracket = 1
  useBackSlash = 1
  useRightBracket = 1
  usePower = 1
  useUnderscore = 1
  useGrave = 1
  useLeftBrace = 1
  usePipe = 1
  useRightBrace = 1
  useTilde = 1
  ' `-> ASCII table order.
  defaultMaxPassLen = 64
  ' `-> Fallback if unsetting from 128, etc.
  defaultPIM = 485
  ' `-> For use when making VeraCrypt file(s)/volume(s).
  '     The default is actually 98 if you use sha512 or Whirlpool, but since I have no way of telling that, I'll use the 2nd default.
  maxPassLen = 64
  minPassLen = 8
  ' `-> WARNING!: DO NOT CHANGE THIS!
  moreRandomness = 0
  Call makeSpecialString
  Dim countNumber As Integer
  For countNumber = 1 To 1024
    cmbNumber.AddItem countNumber
  Next countNumber
  cmbNumber.Text = "64"
  Call addLenNumbers
  cmbPIM.Text = "1024"
End Sub

Private Sub Form_Terminate()
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub menuExit_Click(Index As Integer)
  End
End Sub

Private Sub Override_Click()
  frmOverride.Visible = True
End Sub

Private Sub tmrAutomatic_Timer()
  Call makePass(cmbNumber.Text, cmbLength.Text)
End Sub

Private Sub tmrNoteClear_Timer()
  Me.Caption = Me.Tag
  tmrNoteClear.Enabled = False
End Sub

' EOF
