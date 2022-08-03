VERSION 5.00
Begin VB.Form frmPassGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Generator"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrNoteClear 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3000
      Top             =   3240
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10440
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame fmeString 
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   3735
      Begin VB.ComboBox cmbLength 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   915
      End
      Begin VB.ComboBox cmbNumber 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblLength 
         Caption         =   "That have a length of                        characters"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   645
         Width           =   3495
      End
      Begin VB.Label lblNumber 
         Caption         =   "Generate a total of                          password(s)"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   285
         Width           =   3495
      End
   End
   Begin VB.ListBox lstPasswords 
      Height          =   2790
      Left            =   3960
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   120
      Width           =   7815
   End
   Begin VB.Frame fmeSettings 
      Caption         =   "Settings:"
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3735
      Begin VB.CheckBox chkSpecialChars 
         Caption         =   "Use Special Characters (!, "", $, etc.)"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkNumChars 
         Caption         =   "Use Numbers"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkUpperChars 
         Caption         =   "Use Uppercase characters"
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkLowerChars 
         Caption         =   "Use Lowercase characters"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   555
         Value           =   1  'Checked
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmPassGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  End
End Sub

Private Sub cmdCopy_Click()
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
    MsgBox "Please select at least one password to copy", vbExclamation, "Error"
  End If
End Sub

Private Sub cmdGenerate_Click()
  If chkUpperChars.Value = 0 And chkLowerChars.Value = 0 And chkNumChars.Value = 0 And chkSpecialChars.Value = 0 Then
    MsgBox "Please select at least one option", vbExclamation, "Error"
  Else
    Call makePass(cmbNumber.Text, cmbLength.Text)
  End If
End Sub

Private Function makePass(passCount As Integer, inputLength As Variant) As String
  Randomize
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
    baseString = baseString & "!&'*+-./:;<=>?@\^_`~#$%,|()[]{}" & Chr(34)
    ' `-> Chr(34) is "
  End If
  Dim isRand As Integer
  isRand = 0
  If inputLength = "Rand" Then
    isRand = 1
  End If
  lstPasswords.Clear
  Dim makeNumber As Integer, outString As String
  For makeNumber = 0 To passCount - 1
    outString = ""
    Dim LengthNumber As Integer
    If isRand = 1 Then
      inputLength = Int(Val(8 + Val(Rnd * 56)))
    End If
    For LengthNumber = 0 To inputLength - 1
      outString = outString & Mid(baseString, Int(Val(1 + Val(Rnd * Len(baseString)))), 1)
    Next LengthNumber
    lstPasswords.AddItem outString
  Next makeNumber
  ' `-> The random letter aspect was originally part of a separate function. However, it kept making a weird pattern in the
  '     copying tests to notepad. (If you've ever seen the Malbolge code for "100 bottles of beer" you'll understand what I mean.)
End Function

Private Sub Form_Load()
  Me.Tag = Me.Caption
  Dim loadCountNumber As Integer, loadLengthNumber As Integer
  For loadCountNumber = 1 To 1024
    cmbNumber.AddItem loadCountNumber
  Next loadCountNumber
  For loadLengthNumber = 8 To 64
    cmbLength.AddItem loadLengthNumber
  Next loadLengthNumber
  cmbLength.AddItem "Rand"
  cmbNumber.Text = "64"
  cmbLength.Text = "16"
  ' `-> The minimum is 8, but nobody should really be using sub 16 character passwords in 2022.
End Sub

Private Sub Form_Terminate()
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub tmrNoteClear_Timer()
  Me.Caption = Me.Tag
  tmrNoteClear.Enabled = False
End Sub

' EOF
