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
   MinButton       =   0   'False
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
         List            =   "frmPassGen.frx":0013
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
         ToolTipText     =   "For use in VeraCrypt"
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
         Caption         =   "Include Special characters (!, "", $, etc.)"
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

' ,-> Special characters
Public specialString As String

Public useExclamation As Integer
Public useAmpersand As Integer
Public useApostrophe As Integer
Public useAsterisk As Integer
Public usePlus As Integer
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
Public useBackSlash As Integer
Public usePower As Integer
Public useGrave As Integer
' `-> `
Public useTilde As Integer
Public useHash As Integer
Public useDollar As Integer
Public usePercent As Integer
Public useComma As Integer
Public usePipe As Integer
Public useLeftParent As Integer
Public useRightParent As Integer
Public useLeftBracket As Integer
Public useRightBracket As Integer
Public useLeftBrace As Integer
Public useRightBrace As Integer
Public useQuote As Integer
Public useUnderscore As Integer

Public defaultMaxPassLen As Integer
Public maxPassLen As Integer
Public minPassLen As Integer
Public moreRandomness As Integer

Public defaultPIM As Integer


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
      Dim selectNumber As Integer
      For selectNumber = 0 To lstPasswords.ListCount - 1
        If lstPasswords.Selected(selectNumber) = True Then
          Clipboard.SetText Clipboard.GetText & lstPasswords.List(selectNumber) & vbNewLine
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

Public Function makeSpecialString() As String
  Dim baseString As String
  If useExclamation = 1 Then
    baseString = "!"
  End If
  If useAmpersand = 1 Then
    baseString = baseString & "&"
  End If
  If useApostrophe = 1 Then
    baseString = baseString & "'"
  End If
  If useAsterisk = 1 Then
    baseString = baseString & "*"
  End If
  If usePlus = 1 Then
    baseString = baseString & "+"
  End If
  If useMinus = 1 Then
    baseString = baseString & "-"
  End If
  If usePeriod = 1 Then
    baseString = baseString & "."
  End If
  If useForwardSlash = 1 Then
    baseString = baseString & "/"
  End If
  If useColon = 1 Then
    baseString = baseString & ":"
  End If
  If useSemiColon = 1 Then
    baseString = baseString & ";"
  End If
  If useLessThan = 1 Then
    baseString = baseString & "<"
  End If
  If useEquals = 1 Then
    baseString = baseString & "="
  End If
  If useGreaterThan = 1 Then
    baseString = baseString & ">"
  End If
  If useQuestion = 1 Then
    baseString = baseString & "?"
  End If
  If useAtSign = 1 Then
    baseString = baseString & "@"
  End If
  If useBackSlash = 1 Then
    baseString = baseString & "\"
  End If
  If usePower = 1 Then
    baseString = baseString & "^"
  End If
  If useGrave = 1 Then
    baseString = baseString & "`"
  End If
  If useTilde = 1 Then
    baseString = baseString & "~"
  End If
  If useHash = 1 Then
    baseString = baseString & "#"
  End If
  If useDollar = 1 Then
    baseString = baseString & "$"
  End If
  If usePercent = 1 Then
    baseString = baseString & "%"
  End If
  If useComma = 1 Then
    baseString = baseString & ","
  End If
  If usePipe = 1 Then
    baseString = baseString & "|"
  End If
  If useLeftParent = 1 Then
    baseString = baseString & "("
  End If
  If useRightParent = 1 Then
    baseString = baseString & ")"
  End If
  If useLeftBracket = 1 Then
    baseString = baseString & "["
  End If
  If useRightBracket = 1 Then
    baseString = baseString & "]"
  End If
  If useLeftBrace = 1 Then
    baseString = baseString & "{"
  End If
  If useRightBrace = 1 Then
    baseString = baseString & "}"
  End If
  If useQuote = 1 Then
    baseString = baseString & Chr(34)
  End If
  If useUnderscore = 1 Then
    baseString = baseString & "_"
  End If
  specialString = baseString
  If IsNull(baseString) = True Or baseString = "" Then
    chkSpecialChars.Value = 0
  Else
    chkSpecialChars.Value = 1
  End If
End Function

Private Function makePass(passCount As Integer, inputLength As Variant) As String
  Dim baseString As String
  If chkUpperChars.Value = 1 Then
    baseString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  End If
  If chkLowerChars.Value = 1 Then
    baseString = baseString & "abcdefghijklmnopqrstuvwxyz"
  End If
  If chkNumChars.Value = 1 Then
    baseString = baseString & "0123456789"
  End If
  If chkSpecialChars.Value = 1 Then
    If IsNull(specialString) = False Or specialString <> "" Then
      baseString = baseString & specialString
    End If
  End If
  If chkSpaces.Value = 1 Then
    baseString = baseString & Chr(32)
    ' `-> Chr(32) is <SPACE>
  End If
  Dim isRand As Integer
  isRand = 0
  If inputLength = "Rand" Then
    isRand = 1
  End If
  lstPasswords.Clear
  Randomize
  Dim makeNumber As Integer, randNumber As Integer, outString As String
  For makeNumber = 0 To passCount - 1
    outString = ""
    Dim LengthNumber As Integer
    If isRand = 1 Then
      inputLength = Int(Val(minPassLen + Val(Rnd * Val(maxPassLen - Val(minPassLen - 1)))))
    End If
    For LengthNumber = 0 To inputLength - 1
      If moreRandomness = 1 Then
        randNumber = Int(Val(1 + Val(Rnd * 2)))
        If randNumber = 1 Then
          outString = outString & Mid(baseString, Int(Val(1 + Val(Rnd * Len(baseString)))), 1)
        Else
          outString = outString & Mid(StrReverse(baseString), Int(Val(1 + Val(Rnd * Len(StrReverse(baseString))))), 1)
        End If
      Else
        outString = outString & Mid(baseString, Int(Val(1 + Val(Rnd * Len(baseString)))), 1)
      End If
    Next LengthNumber
    If chkPIM.Value = 1 Then
      If Len(outString) <= 64 Then
        ' `-> VeraCrypt length is limited to 64 characters. There is no point to including a PIM if it's longer than that.
        If Len(outString) >= 20 Then
          outString = outString & " ---------- " & Int(Val(1 + Val(Rnd * cmbPIM.Text)))
        Else
          If cmbPIM.Text > defaultPIM Then
            outString = outString & " ---------- " & Int(Val(defaultPIM + Val(Rnd * Val(cmbPIM.Text - Val(defaultPIM - 1)))))
          End If
        End If
      End If
    End If
    lstPasswords.AddItem outString
  Next makeNumber
  ' `-> The random letter aspect was originally part of a separate function. However, it kept making a weird pattern in the
  '     copying tests to notepad. (If you've ever seen the Malbolge code for "99 bottles of beer" you'll understand what I mean.)
End Function

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
  useExclamation = 1
  useAmpersand = 1
  useApostrophe = 1
  useAsterisk = 1
  usePlus = 1
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
  useBackSlash = 1
  usePower = 1
  useGrave = 1
  useTilde = 1
  useHash = 1
  useDollar = 1
  usePercent = 1
  useComma = 1
  usePipe = 1
  useLeftParent = 1
  useRightParent = 1
  useLeftBracket = 1
  useRightBracket = 1
  useLeftBrace = 1
  useRightBrace = 1
  useQuote = 1
  useUnderscore = 1
  maxPassLen = 64
  defaultMaxPassLen = 64
  ' `-> Fallback if unsetting.
  minPassLen = 8
  ' `-> WARNING!: DO NOT CHANGE THIS!
  moreRandomness = 0
  defaultPIM = 485
  ' `-> VeraCrypt use.
  '     It's actually 98 if you use sha512 or Whirlpool, but since I have no way of telling that, I'll use the 2nd default.
  Call makeSpecialString
  Dim loadCountNumber As Integer
  For loadCountNumber = 1 To 1024
    cmbNumber.AddItem loadCountNumber
  Next loadCountNumber
  cmbNumber.Text = "64"
  Call addLenNumbers
  cmbPIM.Text = "1024"
End Sub

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
  Dim loadLengthNumber As Integer
  For loadLengthNumber = minPassLen To maxPassLen
    cmbLength.AddItem loadLengthNumber
  Next loadLengthNumber
  cmbLength.AddItem "Rand"
  cmbLength.Text = IIf(IsNull(cmbLength.Tag) = False And cmbLength.Tag <> "", cmbLength.Tag, 16)
  ' `-> The minimum is 8, but nobody should really be using sub 16 character passwords in 2022.
End Function

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
