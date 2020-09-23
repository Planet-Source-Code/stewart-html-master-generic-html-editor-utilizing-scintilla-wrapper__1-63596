

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Image"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   Icon            =   "frmImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbURL 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Insert"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox picSplit 
      Height          =   53
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   7515
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3720
      Width           =   7575
   End
   Begin VB.ComboBox cmbAlt 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CheckBox chkFile 
      Caption         =   "&Only include filename"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3000
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picHold 
      Height          =   3195
      Left            =   2760
      ScaleHeight     =   3135
      ScaleWidth      =   4875
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   360
      Width           =   4935
      Begin VB.PictureBox picBlock 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   16
         Top             =   1680
         Width           =   255
      End
      Begin VB.HScrollBar vScroll 
         Height          =   255
         LargeChange     =   450
         Left            =   120
         Max             =   100
         SmallChange     =   200
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2655
      End
      Begin VB.VScrollBar hScroll 
         Height          =   1935
         LargeChange     =   450
         Left            =   4200
         SmallChange     =   200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picImage 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   0
         ScaleHeight     =   97
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   153
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdFile 
      Height          =   315
      Left            =   2280
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmImage.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "1"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.TextBox txtHeight 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtWidth 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox cmbFile 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "&Link URL:"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Preview:"
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "&Alternate Text (If no Images)"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "&Height:"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "&Width:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "&Image"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFile_Click()
  On Error GoTo exitclick
  With cd
    .Filter = "All Image Files (*.gif;*.jpg;*.jpeg;*.png)|*.gif;*.jpg;*.jpeg;*.png|Gif Files (*.gif)|*.gif|Jpeg Files (*.jpg;*.jpeg)|*.jpg;*.jpeg|PNG Files (*.png)|*.png"
    .CancelError = True
    .ShowOpen
    cmbFile.Text = .filename
    picImage.Picture = LoadPicture(.filename)
  End With
  txtWidth.Text = picImage.ScaleWidth
  txtHeight.Text = picImage.ScaleHeight
  If picImage.Width > (picHold.Width - hScroll.Width) Then
    vScroll.Max = picImage.Width - (picHold.Width - hScroll.Width - 120)
  Else
    vScroll.Max = 0
  End If
  If picImage.Height > (picHold.Height - vScroll.Height) Then
    hScroll.Max = picImage.Height - (picHold.Height - vScroll.Height - 120)
  Else
    hScroll.Max = 0
  End If
exitclick:
  ' do nothing cancel was selected
End Sub

Private Sub cmdOK_Click()
  On Error Resume Next
  Dim str As String
  Dim imgPath As String
  str = ""
  If cmbURL <> "" Then
    str = "<A HREF=" & strQ & cmbURL.Text & strQ & ">"
  End If
  If chkFile.value = vbChecked Then
    imgPath = ExtractFileName(cmbFile.Text)
  Else
    imgPath = cmbFile.Text
  End If
  str = str & "<IMG SRC=" & strQ & imgPath & strQ
  If cmbAlt.Text <> "" Then
    str = str & " ALT=" & strQ & cmbAlt.Text & strQ
  End If
  str = str & " WIDTH=" & strQ & txtWidth.Text & strQ & " HEIGHT=" & strQ & txtHeight.Text & strQ & " BORDER=" & strQ & "0" & strQ & ">"
  If cmbURL.Text <> "" Then
    str = str & "</A>"
  End If
  Me.Hide
  frmMain.ActiveForm.sciMain.SelText = str
  frmMain.ActiveForm.sciMain.SetFocus
  Unload Me
End Sub

Private Sub Form_Load()
  Flatten Me
  Call vScroll.Move(0, picHold.ScaleHeight - vScroll.Height, picHold.ScaleWidth - hScroll.Width)
  Call hScroll.Move(picHold.ScaleWidth - hScroll.Width, 0, hScroll.Width, picHold.ScaleHeight - hScroll.Width)
  Call picBlock.Move(picHold.ScaleWidth - picBlock.Width, picHold.ScaleHeight - picBlock.Height)
End Sub

Private Sub hScroll_Change()
  picImage.Top = -(hScroll.value)
End Sub

Private Sub hScroll_Scroll()
  picImage.Top = -(hScroll.value)
End Sub

Private Sub vScroll_Change()
  picImage.Left = -(vScroll.value)
End Sub

Private Sub vScroll_Scroll()
  picImage.Left = -(vScroll.value)
End Sub
Private Sub Form_Initialize()
    InitCommonControls

    InitCommonControls
End Sub

