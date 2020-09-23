VERSION 5.00
Begin VB.Form frmTable 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   765
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FF0000&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   1020
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInfo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1020
      TabIndex        =   0
      Top             =   390
      Width           =   1020
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2325
         TabIndex        =   1
         Top             =   0
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private curSizeX As Long, curSizeY As Long
Private curCols As Long, curRows As Long
Public ColWidth As String
Private Sub resizeForm(x As Long, y As Long)
  Dim g As Long, S As Long
  curSizeX = x
  curSizeY = y
  g = Me.Width - Me.ScaleWidth
  S = Me.Height - Me.ScaleHeight
  Me.Width = (25 * Screen.TwipsPerPixelX) * x + g
  Me.Height = (25 * Screen.TwipsPerPixelY) * y + S + picInfo.Height
End Sub
Private Sub Form_Load()
  Call resizeForm(5, 5)
  StayOnTop Me, True
  PaintLines
End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
  Dim xCell As Integer, yCell As Integer
  If Button = 1 Then
    xCell = (x \ (25 * Screen.TwipsPerPixelX))
    yCell = (y \ (25 * Screen.TwipsPerPixelY)) '
    xCell = xCell + 1
    yCell = yCell + 1
    If xCell + 1 >= curSizeX Then
      curSizeX = xCell + 1
    End If
    If yCell + 1 >= curSizeY Then
      curSizeY = yCell + 1
    End If
    resizeForm curSizeX, curSizeY
    lblInfo.Caption = "Rows: " & yCell & "   Cols: " & xCell
    Me.Cls
    Me.FillStyle = 0
    curCols = xCell
    curRows = yCell
    Me.Line (0, 0)-(((xCell) * (25 * Screen.TwipsPerPixelX)), (yCell) * (25 * Screen.TwipsPerPixelY)), vbBlue, BF
    PaintLines
  End If
  
End Sub

Private Sub PaintLines()
  Dim i As Long, x As Long, f As Long, g As Long
  f = Screen.TwipsPerPixelX * 25
  g = Screen.TwipsPerPixelY * 25
  For i = 1 To Me.ScaleWidth \ f
    Me.Line (i * f, 0)-(i * f, Me.ScaleHeight)
    For x = 1 To (Me.ScaleHeight - picInfo.Height) \ g
      Me.Line (0, x * g)-(Me.ScaleWidth, x * g)
    Next x
  Next i
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
  On Error Resume Next
  Dim str As String
  Dim S As Integer, g As Integer
  str = "<TABLE>" & vbCrLf
  If Button = 1 Then
    For S = 1 To curRows
      str = str & ColWidth & " <TR>" & vbCrLf
      For g = 1 To curCols
        str = str & ColWidth & "    <TD></TD>" & vbCrLf
      Next g
      str = str & ColWidth & "  </TR>" & vbCrLf
    Next S
    
    str = str & ColWidth & "</TABLE>"
  End If
  frmMain.ActiveForm.sciMain.SelText = str
  Unload Me
End Sub

Private Sub Form_Paint()
  PaintLines
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  frmMain.ActiveForm.sciMain.SetFocus
End Sub

Private Sub picInfo_Resize()
  lblInfo.Move (picInfo.ScaleWidth - lblInfo.Width) \ 2, (picInfo.ScaleHeight - lblInfo.Height) \ 2
End Sub
Private Sub Form_Initialize()
    InitCommonControls

    InitCommonControls
End Sub

