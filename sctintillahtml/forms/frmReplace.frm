

VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "&Find Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox chkWrap 
      Caption         =   "Wrap aroun&d"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkWhole 
      Caption         =   "Match &whole word only"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1500
      Width           =   2535
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Match &case"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CheckBox chkRegExp 
      Caption         =   "Regular &expression"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2100
      Width           =   1695
   End
   Begin VB.ComboBox cmbReplace 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.ComboBox cmbFind 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblReplace 
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "Replace with:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   660
      Width           =   975
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "Find what:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cScintilla As clsScintilla

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFindNext_Click()
  cScintilla.FindText cmbFind.Text, False, False, chkWrap.value, chkCase.value, False, chkWhole.value, chkRegExp.value
End Sub

Private Sub cmdReplace_Click()
  Dim iGetPos As Long
  cScintilla.FindText cmbFind.Text, False, False, chkWrap.value, chkCase.value, False, chkWhole.value, chkRegExp.value
  If cmbReplace.Text <> "" Then
    cScintilla.ReplaceSel cmbReplace.Text
    cScintilla.SetSel cScintilla.GetCurPos - 1, cScintilla.GetCurPos - 1 + Len(cmbReplace.Text)
  Else
    iGetPos = cScintilla.GetCurPos
    cScintilla.ReplaceSel ""
    cScintilla.SetSel iGetPos - LenB(cmbFind.Text), iGetPos - LenB(cmbFind.Text)
    'cScintilla.SendEditor SCI_SETTARGETSTART, iGetPos + 1
    'cScintilla.SendEditor SCI_SETTARGETEND, iGetPos + Len(cmbFind.Text)
    cScintilla.FindText cmbFind.Text, False, False, chkWrap.value, chkCase.value, False, chkWhole.value, chkRegExp.value
  End If
  
End Sub

Private Sub cmdReplaceAll_Click()
  Dim iRep As Long
  iRep = cScintilla.ReplaceAll(cmbFind.Text, cmbReplace.Text, chkCase.value, chkRegExp.value, chkWhole.value, False)
  If iRep > 0 Then
    lblReplace.Caption = "Replaced " & iRep & " times"
  Else
    MsgBox "No instances of " & """" & cmbFind.Text & """" & " were found in document"
  End If
End Sub

Private Sub Form_Load()
  Flatten Me
  Me.Left = GetSetting("ScintillaClass", "Settings", "ReplaceLeft", (Screen.Width - Me.Width) \ 2)
  Me.Top = GetSetting("ScintillaClass", "Settings", "ReplaceTop", (Screen.Height - Me.Height) \ 2)
  chkCase.value = GetSetting("ScintillaClass", "Settings", "RchkCase", 0)
  chkRegExp.value = GetSetting("ScintillaClass", "Settings", "RchkRegEx", 0)
  chkWhole.value = GetSetting("ScintillaClass", "Settings", "RchkWhole", 0)
  chkWrap.value = GetSetting("ScintillaClass", "Settings", "RchkWrap", 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting "ScintillaClass", "Settings", "ReplaceLeft", Me.Left
  SaveSetting "ScintillaClass", "Settings", "ReplaceTop", Me.Top
  SaveSetting "ScintillaClass", "Settings", "RchkCase", chkCase.value
  SaveSetting "ScintillaClass", "Settings", "RchkRegEx", chkRegExp.value
  SaveSetting "ScintillaClass", "Settings", "RchkWhole", chkWhole.value
  SaveSetting "ScintillaClass", "Settings", "RchkWrap", chkWrap.value
  cScintilla.SetFocus
End Sub
Private Sub Form_Initialize()
    InitCommonControls

    InitCommonControls
End Sub

