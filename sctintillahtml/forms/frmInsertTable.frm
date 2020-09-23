VERSION 5.00
Begin VB.Form frmInsertTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Table"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   Icon            =   "frmInsertTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5880
      TabIndex        =   19
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Default         =   -1  'True
      Height          =   495
      Left            =   5880
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin HTMLMaster.ArielColorBox clrBorder 
      Height          =   315
      Left            =   1920
      TabIndex        =   17
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
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
      Left            =   1920
      TabIndex        =   16
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
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
   Begin VB.ComboBox cmbWidth 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmInsertTable.frx":000C
      Left            =   1920
      List            =   "frmInsertTable.frx":000E
      TabIndex        =   15
      Text            =   "100%"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtBWidth 
      Height          =   285
      Left            =   4680
      TabIndex        =   13
      Text            =   "1"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtPadd 
      Height          =   285
      Left            =   4680
      TabIndex        =   12
      Text            =   "2"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtSpacing 
      Height          =   285
      Left            =   4680
      TabIndex        =   11
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtCols 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "2"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtRows 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Text            =   "2"
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton optPercent 
      Caption         =   "Percent"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   2880
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optPixels 
      Caption         =   "Pixels"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Table Width:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Border Width:"
      Height          =   195
      Left            =   3480
      TabIndex        =   6
      Top             =   1125
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Cell Padding:"
      Height          =   195
      Left            =   3480
      TabIndex        =   5
      Top             =   645
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Cell Spacing:"
      Height          =   195
      Left            =   3480
      TabIndex        =   4
      Top             =   165
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Border Color:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Background Color:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1140
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Number of Columns:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   645
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of Rows:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   1230
   End
End
Attribute VB_Name = "frmInsertTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ColWidth As String
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdInsert_Click()
  On Error Resume Next
  Dim str As String, strWidth As String
  Dim x As Long, y As Long
  If IsNumeric(txtRows.Text) = False Then
    MsgBox "Unable to continue.  Number of rows must be numeric", vbCritical
    Exit Sub
  End If
  If IsNumeric(txtCols.Text) = False Then
    MsgBox "Unable to continue.  Number of columns must be numeric", vbCritical
    Exit Sub
  End If
  strWidth = cmbWidth.Text
  If optPercent.value = True Then
    If Right(strWidth, 1) <> "%" Then
      strWidth = strWidth & "%"
    End If
  Else
    If Right(strWidth, 1) = "%" Then
      strWidth = Left(strWidth, Len(strWidth) - 1)
    End If
  End If
  str = "<TABLE BORDER=" & strQ & txtBWidth.Text & strQ & " CELLPADDING=" & strQ & txtPadd.Text & strQ & " CELLSPACING=" & strQ & txtSpacing.Text & strQ & " BGCOLOR=" & strQ & HexRGB(clrBack.SelectedColor) & strQ & " BORDERCOLOR=" & strQ & clrBorder.SelectedColor & strQ & " WIDTH=" & strQ & strWidth & strQ & ">" & vbCrLf
  For x = 1 To txtRows.Text
    str = str & ColWidth & "  <TR>" & vbCrLf
    For y = 1 To txtCols.Text
      str = str & ColWidth & "    <TD></TD>" & vbCrLf
    Next y
    str = str & ColWidth & "  </TR>" & vbCrLf
  Next x
  str = str & ColWidth & "</TABLE>"
  Me.Hide
  frmMain.ActiveForm.sciMain.SelText = str
  frmMain.ActiveForm.sciMain.SetFocus
  Unload Me
End Sub

Private Sub Form_Load()
  optPercent_Click
  Flatten Me
End Sub

Private Sub optPercent_Click()
  Dim x As Long
  For x = 1 To 10
    If x <> 10 Then
      Call cmbWidth.AddItem(x * 10 & " %")
    Else
      Call cmbWidth.AddItem(x * 10 & "%")
    End If
  Next x
  cmbWidth.ListIndex = 9
End Sub

Private Sub optPixels_Click()
  cmbWidth.Clear
  cmbWidth.Text = "200"
End Sub
