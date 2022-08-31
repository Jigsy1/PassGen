VERSION 5.00
Begin VB.Form frmPassGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Generator"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
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
      Caption         =   "Include:"
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
         Caption         =   "Include Numbers (0-9)"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkUpperChars 
         Caption         =   "Include Uppercase characters (A-Z)"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkLowerChars 
         Caption         =   "Include Lowercase characters (a-z)"
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
      Begin VB.Menu menuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu menuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu menuSpecial 
         Caption         =   "&Include Special characters"
         Shortcut        =   ^I
      End
      Begin VB.Menu menuOverride 
         Caption         =   "&Override settings"
         Shortcut        =   ^O
      End
      Begin VB.Menu menuSeparatorA 
         Caption         =   "-"
      End
      Begin VB.Menu menuSave 
         Caption         =   "&Save Settings to Registry on Exit"
      End
   End
   Begin VB.Menu menuAbout 
      Caption         =   "&About"
      Begin VB.Menu menuAboutForm 
         Caption         =   "&About"
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
Public maxPassCount As Integer
Public maxPassLen As Integer
Public minPassLen As Integer
Public moreRandomness As Integer
Public specialsNeeded As Integer

Public ourPath As String
' `-> Registry entry.
' Public useSave As Integer

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
  ' `-> The minimum is eight, but nobody should really be using sub 16 character passwords in 2022.
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
  If IsNull(baseString) = True Or baseString = "" Or Len(baseString) < specialsNeeded Then
    chkSpecialChars.Value = 0
  Else
    chkSpecialChars.Value = 1
  End If
End Function

Public Function isBool(inputBool As Variant)
  If LCase(inputBool) = "true" Or inputBool = "1" Then
    isBool = "1"
  Else
    ' `-> Treat everything else as false.
    isBool = "0"
  End If
End Function

Public Function isRegKey(inputKey As String)
  Dim thisKey As String
  Dim thisObject As Object
  Set thisObject = CreateObject("WScript.Shell")
  On Error Resume Next
  isRegKey = False
  thisKey = thisObject.regRead(inputKey)
  If Err.Number = 0 Then
    isRegKey = True
  ElseIf thisKey = "" Then
    isRegKey = False
  ElseIf IsNull(thisKey) = True Then
    isRegKey = False
  ElseIf CBool(InStr(Err.Description, "Unable")) Then
    isRegKey = False
  End If
  Err.Clear: On Error GoTo 0
End Function

Public Function saveToRegistry()
  If menuSave.Checked = True Then
    Dim qS As Object
    Set qS = CreateObject("WScript.Shell")
    ' `-> q(uick)S(hell).
    ' qS.regWrite ourPath & "useSave", useSave, "REG_DWORD"
    qS.regWrite ourPath & "useUppercase", chkUpperChars.Value, "REG_DWORD"
    qS.regWrite ourPath & "useLowercase", chkLowerChars.Value, "REG_DWORD"
    qS.regWrite ourPath & "useNumbers", chkNumChars.Value, "REG_DWORD"
    qS.regWrite ourPath & "useSpecials", chkSpecialChars.Value, "REG_DWORD"
    qS.regWrite ourPath & "useSpace", chkSpaces.Value, "REG_DWORD"
    qS.regWrite ourPath & "passCount", cmbNumber.Text, "REG_DWORD"
    qS.regWrite ourPath & "passLength", cmbLength.Text, "REG_SZ"
    qS.regWrite ourPath & "useAutomatic", chkAutomatic.Value, "REG_DWORD"
    qS.regWrite ourPath & "usePIM", chkPIM.Value, "REG_DWORD"
    qS.regWrite ourPath & "PIM", cmbPIM.Text, "REG_DWORD"
    ' ,-> Special(s).
    qS.regWrite ourPath & "useExclamation", useExclamation, "REG_DWORD"
    qS.regWrite ourPath & "useQuote", useQuote, "REG_DWORD"
    qS.regWrite ourPath & "useHash", useHash, "REG_DWORD"
    qS.regWrite ourPath & "useDollar", useDollar, "REG_DWORD"
    qS.regWrite ourPath & "usePercent", usePercent, "REG_DWORD"
    qS.regWrite ourPath & "useAmpersand", useAmpersand, "REG_DWORD"
    qS.regWrite ourPath & "useApostrophe", useApostrophe, "REG_DWORD"
    qS.regWrite ourPath & "useLeftParenthesis", useLeftParenthesis, "REG_DWORD"
    qS.regWrite ourPath & "useRightParenthesis", useRightParenthesis, "REG_DWORD"
    qS.regWrite ourPath & "useAsterisk", useAsterisk, "REG_DWORD"
    qS.regWrite ourPath & "usePlus", usePlus, "REG_DWORD"
    qS.regWrite ourPath & "useComma", useComma, "REG_DWORD"
    qS.regWrite ourPath & "useMinus", useMinus, "REG_DWORD"
    qS.regWrite ourPath & "usePeriod", usePeriod, "REG_DWORD"
    qS.regWrite ourPath & "useForwardSlash", useForwardSlash, "REG_DWORD"
    qS.regWrite ourPath & "useColon", useColon, "REG_DWORD"
    qS.regWrite ourPath & "useSemiColon", useSemiColon, "REG_DWORD"
    qS.regWrite ourPath & "useLessThan", useLessThan, "REG_DWORD"
    qS.regWrite ourPath & "useEquals", useEquals, "REG_DWORD"
    qS.regWrite ourPath & "useGreaterThan", useGreaterThan, "REG_DWORD"
    qS.regWrite ourPath & "useQuestion", useQuestion, "REG_DWORD"
    qS.regWrite ourPath & "useAtSign", useAtSign, "REG_DWORD"
    qS.regWrite ourPath & "useLeftBracket", useLeftBracket, "REG_DWORD"
    qS.regWrite ourPath & "useBackSlash", useBackSlash, "REG_DWORD"
    qS.regWrite ourPath & "useRightBracket", useRightBracket, "REG_DWORD"
    qS.regWrite ourPath & "usePower", usePower, "REG_DWORD"
    qS.regWrite ourPath & "useUnderscore", useUnderscore, "REG_DWORD"
    qS.regWrite ourPath & "useGrave", useGrave, "REG_DWORD"
    qS.regWrite ourPath & "useLeftBrace", useLeftBrace, "REG_DWORD"
    qS.regWrite ourPath & "usePipe", usePipe, "REG_DWORD"
    qS.regWrite ourPath & "useRightBrace", useRightBrace, "REG_DWORD"
    qS.regWrite ourPath & "useTilde", useTilde, "REG_DWORD"
    ' ,-> Override
    qS.regWrite ourPath & "override128", IIf(maxPassLen = 128, 1, 0), "REG_DWORD"
    qS.regWrite ourPath & "override256", IIf(maxPassLen = 256, 1, 0), "REG_DWORD"
    qS.regWrite ourPath & "override512", IIf(maxPassLen = 512, 1, 0), "REG_DWORD"
    qS.regWrite ourPath & "moreRandomness", moreRandomness, "REG_DWORD"
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
  If IsNull(specialString) = True Or specialString = "" Or Len(specialString) < specialsNeeded Then
    If chkSpecialChars.Value = 1 Then MsgBox "Cannot enable. You either have no special character(s) chosen, or less than " & specialsNeeded & " enabled.", vbExclamation, "Error"
    chkSpecialChars.Value = 0
  End If
End Sub

Private Sub cmdClose_Click()
  Call saveToRegistry
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then Call makePass(cmbNumber.Text, cmbLength.Text)
  If KeyCode = vbKeyDelete Then lstPasswords.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    Call saveToRegistry
    End
  End If
End Sub

Private Sub Form_Load()
  Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
  Me.Tag = Me.Caption
  ' ,-> Set the default global variable(s).
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
  maxPassCount = 1024
  maxPassLen = 64
  minPassLen = 8
  ' `-> WARNING!: DO NOT CHANGE THIS! NOBODY SHOULD BE USING PASSWORDS LESS THAN EIGHT CHARACTERS ANYMORE!
  moreRandomness = 0
  specialsNeeded = 2
  ' `-> Default number of special character(s) required. (You can drop this to 1 if you like, but things will break if you set it to 0!)
  Call makeSpecialString
  Dim countNumber As Integer
  For countNumber = 1 To maxPassCount
    cmbNumber.AddItem countNumber
  Next countNumber
  cmbNumber.Text = "64"
  Call addLenNumbers
  cmbPIM.Text = "1024"
  ' useSave = 0
  ' ,-> After we've assigned everything first (defaults), check against the registry for saved settings and fiddle with the assignments.
  ourPath = "HKCU\Software\Github\Jigsy1\PassGen\"
  ' `-> This is the registry key we're going to save to. If something goes wrong (because you changed something), just delete the entire key.
  If isRegKey(ourPath) = True Then
    menuSave.Checked = True
    Dim qS As Object
    ' `-> q(uick)S(hell)
    Set qS = CreateObject("WScript.Shell")
    ' If isRegKey(ourPath & "useSave") = True Then useSave = isBool(qS.regRead(ourPath & "useSave"))
    If isRegKey(ourPath & "useUppercase") = True Then chkUpperChars.Value = isBool(qS.regRead(ourPath & "useUppercase"))
    If isRegKey(ourPath & "useLowercase") = True Then chkLowerChars.Value = isBool(qS.regRead(ourPath & "useLowercase"))
    If isRegKey(ourPath & "useNumbers") = True Then chkNumChars.Value = isBool(qS.regRead(ourPath & "useNumbers"))
    ' -> See the note below!
    If isRegKey(ourPath & "useSpace") = True Then chkSpaces.Value = isBool(qS.regRead(ourPath & "useSpace"))
    ' ,-> Specials.
    If isRegKey(ourPath & "useExclamation") = True Then useExclamation = isBool(qS.regRead(ourPath & "useExclamation"))
    If isRegKey(ourPath & "useQuote") = True Then useQuote = isBool(qS.regRead(ourPath & "useQuote"))
    If isRegKey(ourPath & "useHash") = True Then useHash = isBool(qS.regRead(ourPath & "useHash"))
    If isRegKey(ourPath & "useDollar") = True Then useDollar = isBool(qS.regRead(ourPath & "useDollar"))
    If isRegKey(ourPath & "usePercent") = True Then usePercent = isBool(qS.regRead(ourPath & "usePercent"))
    If isRegKey(ourPath & "useAmpersand") = True Then useAmpersand = isBool(qS.regRead(ourPath & "useAmpersand"))
    If isRegKey(ourPath & "useApostrophe") = True Then useApostrophe = isBool(qS.regRead(ourPath & "useApostrophe"))
    If isRegKey(ourPath & "useLeftParenthesis") = True Then useLeftParenthesis = isBool(qS.regRead(ourPath & "useLeftParenthesis"))
    If isRegKey(ourPath & "useRightParenthesis") = True Then useRightParenthesis = isBool(qS.regRead(ourPath & "useRightParenthesis"))
    If isRegKey(ourPath & "useAsterisk") = True Then useAsterisk = isBool(qS.regRead(ourPath & "useAsterisk"))
    If isRegKey(ourPath & "usePlus") = True Then usePlus = isBool(qS.regRead(ourPath & "usePlus"))
    If isRegKey(ourPath & "useComma") = True Then useComma = isBool(qS.regRead(ourPath & "useComma"))
    If isRegKey(ourPath & "useMinus") = True Then useMinus = isBool(qS.regRead(ourPath & "useMinus"))
    If isRegKey(ourPath & "usePeriod") = True Then usePeriod = isBool(qS.regRead(ourPath & "usePeriod"))
    If isRegKey(ourPath & "useForwardSlash") = True Then useForwardSlash = isBool(qS.regRead(ourPath & "useForwardSlash"))
    If isRegKey(ourPath & "useColon") = True Then useColon = isBool(qS.regRead(ourPath & "useColon"))
    If isRegKey(ourPath & "useSemiColon") = True Then useSemiColon = isBool(qS.regRead(ourPath & "useSemiColon"))
    If isRegKey(ourPath & "useLessThan") = True Then useLessThan = isBool(qS.regRead(ourPath & "useLessThan"))
    If isRegKey(ourPath & "useEquals") = True Then useEquals = isBool(qS.regRead(ourPath & "useEquals"))
    If isRegKey(ourPath & "useGreaterThan") = True Then useGreaterThan = isBool(qS.regRead(ourPath & "useGreaterThan"))
    If isRegKey(ourPath & "useQuestion") = True Then useQuestion = isBool(qS.regRead(ourPath & "useQuestion"))
    If isRegKey(ourPath & "useAtSign") = True Then useAtSign = isBool(qS.regRead(ourPath & "useAtSign"))
    If isRegKey(ourPath & "useLeftBracket") = True Then useLeftBracket = isBool(qS.regRead(ourPath & "useLeftBracket"))
    If isRegKey(ourPath & "useBackSlash") = True Then useBackSlash = isBool(qS.regRead(ourPath & "useBackSlash"))
    If isRegKey(ourPath & "useRightBracket") = True Then useRightBracket = isBool(qS.regRead(ourPath & "useRightBracket"))
    If isRegKey(ourPath & "usePower") = True Then usePower = isBool(qS.regRead(ourPath & "usePower"))
    If isRegKey(ourPath & "useUnderscore") = True Then useUnderscore = isBool(qS.regRead(ourPath & "useUnderscore"))
    If isRegKey(ourPath & "useGrave") = True Then useGrave = isBool(qS.regRead(ourPath & "useGrave"))
    If isRegKey(ourPath & "useLeftBrace") = True Then useLeftBrace = isBool(qS.regRead(ourPath & "useLeftBrace"))
    If isRegKey(ourPath & "usePipe") = True Then usePipe = isBool(qS.regRead(ourPath & "usePipe"))
    If isRegKey(ourPath & "useRightBrace") = True Then useRightBrace = isBool(qS.regRead(ourPath & "useRightBrace"))
    If isRegKey(ourPath & "useTilde") = True Then useTilde = isBool(qS.regRead(ourPath & "useTilde"))
    Call makeSpecialString
    If isRegKey(ourPath & "useSpecials") = True Then chkSpecialChars.Value = isBool(qS.regRead(ourPath & "useSpecials"))
    ' `-> We have to do this one here otherwise makeSpecialString will enable it regardless of our choice...
    ' ,-> Override.
    If isRegKey(ourPath & "override128") = True Then
      If isBool(qS.regRead(ourPath & "override128")) = 1 Then maxPassLen = 128
    End If
    If isRegKey(ourPath & "override256") = True Then
      If isBool(qS.regRead(ourPath & "override256")) = 1 And maxPassLen = defaultMaxPassLen Then maxPassLen = 256
    End If
    If isRegKey(ourPath & "override512") = True Then
      If isBool(qS.regRead(ourPath & "override512")) = 1 And maxPassLen = defaultMaxPassLen Then maxPassLen = 512
    End If
    ' `-> If on the off chance someone modifies the registry to include all three, it'll just default to 128. (Or 256?)
    If maxPassLen <> defaultMaxPassLen Then Call addLenNumbers
    If isRegKey(ourPath & "moreRandomness") = True Then moreRandomness = isBool(qS.regRead(ourPath & "moreRandomness"))
    ' ,-> Lastly...
    If isRegKey(ourPath & "passCount") = True Then
      If IsNumeric(qS.regRead(ourPath & "passCount")) = True Then
        Dim tempPassCount As Integer
        tempPassCount = qS.regRead(ourPath & "passCount")
        If tempPassCount > 0 And tempPassCount <= maxPassCount Then cmbNumber.Text = tempPassCount
      End If
    End If
    If isRegKey(ourPath & "passLength") = True Then
      Dim tempPassLength As Variant
      tempPassLength = qS.regRead(ourPath & "passLength")
      If IsNumeric(tempPassLength) = True Then
        If tempPassLength > 0 And tempPassLength <= maxPassLen Then cmbLength.Text = tempPassLength
      Else
        If tempPassLength = "Rand" Then cmbLength.Text = tempPassLength
      End If
    End If
    If isRegKey(ourPath & "useAutomatic") = True Then chkAutomatic.Value = isBool(qS.regRead(ourPath & "useAutomatic"))
    If isRegKey(ourPath & "usePIM") = True Then chkPIM.Value = isBool(qS.regRead(ourPath & "usePIM"))
    If isRegKey(ourPath & "PIM") = True Then
      If IsNumeric(qS.regRead(ourPath & "PIM")) = True Then
        Dim tempPIM As Integer
        tempPIM = qS.regRead(ourPath & "PIM")
        If tempPIM >= cmbPIM.List(0) And tempPIM <= cmbPIM.List(cmbPIM.ListCount - 1) Then cmbPIM.Text = tempPIM
      End If
    End If
  End If
End Sub

Private Sub Form_Terminate()
  Call saveToRegistry
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call saveToRegistry
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub menuAboutForm_Click()
  frmAbout.Visible = True
End Sub

Private Sub menuExit_Click()
  Call saveToRegistry
  End
End Sub

Private Sub menuOverride_Click()
  frmOverride.Visible = True
End Sub

Private Sub menuSave_Click()
  If menuSave.Checked = True Then
    menuSave.Checked = False
    ' ,-> Destroy the registry entry. (It's easier!)
    Dim qS As Object
    Set qS = CreateObject("WScript.Shell")
    If isRegKey(ourPath) = True Then qS.RegDelete ourPath
    Me.Caption = Me.Tag & " - Cleared settings from the registry"
    tmrNoteClear.Enabled = True
  Else
    menuSave.Checked = True
  End If
End Sub

Private Sub menuSpecial_Click()
  frmSpecial.Visible = True
End Sub

Private Sub tmrAutomatic_Timer()
  If chkUpperChars.Value <> 0 Or chkLowerChars.Value <> 0 Or chkNumChars.Value <> 0 Or chkSpecialChars.Value <> 0 Then Call makePass(cmbNumber.Text, cmbLength.Text)
End Sub

Private Sub tmrNoteClear_Timer()
  Me.Caption = Me.Tag
  tmrNoteClear.Enabled = False
End Sub

' EOF
