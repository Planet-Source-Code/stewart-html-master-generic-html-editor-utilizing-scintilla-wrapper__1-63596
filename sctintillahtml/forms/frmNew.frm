

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Document"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7620
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2835
      ScaleWidth      =   5595
      TabIndex        =   8
      Top             =   600
      Width           =   5655
      Begin VB.Frame fmAppearance 
         Caption         =   "Appearance"
         Height          =   2295
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   5415
         Begin HTMLMaster.ArielColorBox clrALink 
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SelectedColor   =   255
         End
         Begin HTMLMaster.ArielColorBox clrVLink 
            Height          =   315
            Left            =   1080
            TabIndex        =   5
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SelectedColor   =   16711935
         End
         Begin HTMLMaster.ArielColorBox clrLink 
            Height          =   315
            Left            =   1080
            TabIndex        =   4
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SelectedColor   =   16711680
         End
         Begin HTMLMaster.ArielColorBox clrText 
            Height          =   315
            Left            =   1080
            TabIndex        =   3
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin HTMLMaster.ArielColorBox clrBack 
            Height          =   315
            Left            =   1080
            TabIndex        =   2
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SelectedColor   =   16777215
         End
         Begin VB.PictureBox picSample 
            BackColor       =   &H00FFFFFF&
            Height          =   1575
            Left            =   2400
            ScaleHeight     =   1515
            ScaleWidth      =   2835
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   480
            Width           =   2895
            Begin VB.Label lblALink 
               BackStyle       =   0  'Transparent
               Caption         =   "Active Link"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   1200
               Width           =   2055
            End
            Begin VB.Label lblVLink 
               BackStyle       =   0  'Transparent
               Caption         =   "Visited Link"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C000C0&
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   840
               Width           =   2055
            End
            Begin VB.Label lblLink 
               BackStyle       =   0  'Transparent
               Caption         =   "Link"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label lblText 
               BackStyle       =   0  'Transparent
               Caption         =   "Text"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   120
               Width           =   2055
            End
         End
         Begin VB.Label Label3 
            Caption         =   "Sample:"
            Height          =   375
            Left            =   2400
            TabIndex        =   24
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Active Link"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   1845
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Visited Link:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   1485
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Text:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   765
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Link:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   1125
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Background:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   405
            Width           =   915
         End
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Page &Title:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   150
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   27
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Create &Page"
      Default         =   -1  'True
      Height          =   375
      Left            =   6120
      TabIndex        =   26
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2835
      ScaleWidth      =   5595
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtKeyword 
         Height          =   1245
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox txtDesc 
         Height          =   1245
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   16
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Meta &Keywords:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "&Meta Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   135
         Width           =   2775
      End
   End
   Begin MSComctlLib.TabStrip tbsNew 
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6165
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Basic"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advanced"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clrALink_Click()
  lblALink.ForeColor = clrALink.SelectedColor
End Sub

Private Sub clrBack_Click()
  picSample.BackColor = clrBack.SelectedColor
End Sub

Private Sub clrLink_Click()
  lblLink.ForeColor = clrLink.SelectedColor
End Sub

Private Sub clrText_Click()
  lblText.ForeColor = clrText.SelectedColor
End Sub

Private Sub clrVLink_Click()
  lblVLink.ForeColor = clrVLink.SelectedColor
End Sub

Private Sub Command1_Click()
  
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error Resume Next
  Dim str As String

  str = "<HTML>" & vbCrLf
  str = str & "  <HEAD>" & vbCrLf
  str = str & "    <TITLE>" & txtTitle.Text & "</TITLE>" & vbCrLf
  str = str & "    <META NAME=" & strQ & "KEYWORDS" & strQ & " CONTENT=" & strQ & txtKeyword.Text & strQ & ">" & vbCrLf
  str = str & "    <META NAME=" & strQ & "DESCRIPTION" & strQ & " CONTENT=" & strQ & txtDesc.Text & strQ & ">" & vbCrLf
  str = str & "  </HEAD>" & vbCrLf
  str = str & "  <BODY BGCOLOR=" & strQ & HexRGB(picSample.BackColor) & strQ & " TEXT=" & strQ & HexRGB(clrText.SelectedColor) & strQ & " LINK=" & strQ & HexRGB(clrLink.SelectedColor) & strQ & " ALINK=" & strQ & HexRGB(clrALink.SelectedColor) & strQ & " VLINK=" & strQ & HexRGB(clrVLink.SelectedColor) & strQ & ">" & vbCrLf
  str = str & "    " & vbCrLf
  str = str & "  </BODY>" & vbCrLf
  str = str & "</HTML>"
  Me.Hide
  frmMain.NewDoc "New Document"
  frmMain.ActiveForm.sciMain.Text = str
  frmMain.ActiveForm.sciMain.GotoLineColumn 7, 4
  frmMain.ActiveForm.sciMain.SetFocus
End Sub

Private Sub Form_Load()
  Flatten Me
End Sub


Private Sub TabStrip1_Click()

End Sub

Private Sub tbsNew_Click()
  Picture1.Visible = False
  Picture2.Visible = False
  Select Case tbsNew.SelectedItem.index
    Case 1
      Picture1.Visible = True
    Case 2
      Picture2.Visible = True
  End Select
End Sub
Private Sub Form_Initialize()
    InitCommonControls

    InitCommonControls
End Sub

