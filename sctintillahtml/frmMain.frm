VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "HTML Master"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10350
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgBasic 
      Left            =   6600
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F36
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1048
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":149A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2190
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      Height          =   6540
      Left            =   0
      ScaleHeight     =   6480
      ScaleWidth      =   2865
      TabIndex        =   3
      Top             =   390
      Width           =   2925
      Begin VB.PictureBox picJava 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   480
         ScaleHeight     =   1695
         ScaleWidth      =   1815
         TabIndex        =   18
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
         Begin VB.FileListBox flList 
            Appearance      =   0  'Flat
            Height          =   1200
            Left            =   0
            TabIndex        =   19
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   600
         ScaleHeight     =   2595
         ScaleWidth      =   2265
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   2265
         Begin MSComctlLib.ImageList images 
            Left            =   960
            Top             =   1200
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":372A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":3C7C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView TagsD 
            Height          =   1530
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   2699
            _Version        =   393217
            Indentation     =   5
            LineStyle       =   1
            Style           =   7
            ImageList       =   "images"
            Appearance      =   1
         End
      End
      Begin VB.PictureBox pic16 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         Top             =   45
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox pic32 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   1590
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picFiles 
         BorderStyle     =   0  'None
         Height          =   3555
         Left            =   600
         ScaleHeight     =   3555
         ScaleWidth      =   2505
         TabIndex        =   5
         Top             =   4560
         Width           =   2505
         Begin VB.PictureBox picSizer 
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Height          =   50
            Left            =   240
            ScaleHeight     =   45
            ScaleWidth      =   495
            TabIndex        =   6
            Top             =   1920
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.DirListBox Dir1 
            Height          =   1440
            Left            =   315
            TabIndex        =   8
            Top             =   420
            Width           =   2220
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   300
            TabIndex        =   7
            Top             =   0
            Width           =   2235
         End
         Begin MSComctlLib.ListView File1 
            Height          =   1710
            Left            =   480
            TabIndex        =   9
            Top             =   2040
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   3016
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "File"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Path"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Image imgSize 
            Height          =   45
            Left            =   960
            MouseIcon       =   "frmMain.frx":41CE
            MousePointer    =   99  'Custom
            Top             =   1920
            Width           =   2055
         End
      End
      Begin MSComctlLib.TabStrip tbsSide 
         Height          =   1215
         Left            =   240
         TabIndex        =   4
         Top             =   2880
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2143
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Files"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Tags"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Java"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList iml32 
         Left            =   1200
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   -2147483644
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList iml16 
         Left            =   360
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   -2147483644
         _Version        =   393216
      End
   End
   Begin VB.PictureBox picSize 
      Align           =   3  'Align Left
      Height          =   6540
      Left            =   2925
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6540
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   390
      Width           =   50
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   688
      _CBWidth        =   10350
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinHeight1      =   330
      Width1          =   4800
      NewRow1         =   0   'False
      Child2          =   "tbrFormat"
      MinHeight2      =   330
      Width2          =   4200
      NewRow2         =   0   'False
      Child3          =   "tbrInsert"
      MinHeight3      =   330
      Width3          =   5265
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   330
         Left            =   165
         TabIndex        =   17
         Top             =   30
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "new"
               Object.ToolTipText     =   "New Document"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "open"
               Object.ToolTipText     =   "Open Document"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "close"
               Object.ToolTipText     =   "Close"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "print"
               Object.ToolTipText     =   "Print"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "save"
               Object.ToolTipText     =   "Save Document"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "undo"
               Object.ToolTipText     =   "Undo"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "redo"
               Object.ToolTipText     =   "Redo"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cut"
               Object.ToolTipText     =   "Cut"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "copy"
               Object.ToolTipText     =   "Copy"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "paste"
               Object.ToolTipText     =   "Paste"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "qtable"
               Object.ToolTipText     =   "Quick Table"
               ImageIndex      =   28
               Style           =   5
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrInsert 
         Height          =   330
         Left            =   9225
         TabIndex        =   16
         Top             =   30
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "link"
               Object.ToolTipText     =   "Insert Hyperlink"
               ImageIndex      =   30
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "table"
               Object.ToolTipText     =   "Table"
               ImageIndex      =   31
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "image"
               Object.ToolTipText     =   "Image"
               ImageIndex      =   30
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrFormat 
         Height          =   330
         Left            =   4995
         TabIndex        =   14
         Top             =   30
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgBasic"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bold"
               Object.ToolTipText     =   "Bold"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "italic"
               Object.ToolTipText     =   "Italic"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "underline"
               Object.ToolTipText     =   "Underline"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "strike"
               Object.ToolTipText     =   "Strike Through"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "left"
               Object.ToolTipText     =   "Align Left"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "center"
               Object.ToolTipText     =   "Align Center"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "right"
               Object.ToolTipText     =   "Align Right"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
         EndProperty
         Begin HTMLMaster.ArielColorBox clrSelect 
            Height          =   315
            Left            =   2760
            TabIndex        =   15
            Top             =   0
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
            Palette         =   5
         End
      End
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6930
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7461
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4200
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5640
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4320
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4772
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4BC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5016
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5468
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":615E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":72A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":76F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":83EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8840
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":90E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9536
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9988
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A22C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A67E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AAD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AF22
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B374
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B7C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BB18
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BE6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C1BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Document"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Document"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Document &As"
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExportHTML 
         Caption         =   "&Export to HTML"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Document"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Co&py"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSyntax 
         Caption         =   "&Syntax Settings"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTimeDate 
         Caption         =   "&Time Date"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Format"
      Begin VB.Menu mnuBold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuItalic 
         Caption         =   "&Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuUnderline 
         Caption         =   "&Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuStrike 
         Caption         =   "&Strikethrough"
      End
      Begin VB.Menu mnuSep09 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "Align &Left"
      End
      Begin VB.Menu mnuCenter 
         Caption         =   "Algin &Center"
      End
      Begin VB.Menu mnuRight 
         Caption         =   "Align &Right"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuHyperlink 
         Caption         =   "&Hyperlink"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuImage 
         Caption         =   "&Image"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuTable 
         Caption         =   "&Table"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFindPrevious 
         Caption         =   "Find Previous"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "&Goto"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Window"
      Begin VB.Menu mnuHor 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnuVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowList 
         Caption         =   "&Window List"
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO

  
Private Sub clrSelect_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelText = HexRGB(clrSelect.SelectedColor)
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub Dir1_Change()
Dim path As String

Initialise
path = Dir1.path
FillFile1WithFiles path
GetAllIcons
ShowIcons
End Sub

Private Sub File1_DblClick()
  NewDoc Dir1.path & "\" & File1.SelectedItem.Text
End Sub

Private Sub flList_DblClick()
  On Error Resume Next
  Dim col As String
  Dim str As String, strs As String
  Dim iFile As Integer
  iFile = FreeFile
  col = Space(ActiveForm.sciMain.GetColumn)
  str = ""
  Open flList.path & "\" & flList.filename For Input As #iFile
    Do While Not EOF(iFile)
      Input #iFile, strs
      str = str & col & strs & vbCrLf
    Loop
  Close #iFile
  ActiveForm.sciMain.SelText = str
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub imgSize_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
  picSizer.Visible = True
End Sub

Private Sub imgSize_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
  Dim nxtY As Long
  If Button = 1 Then
    nxtY = (imgSize.Top + y)
    If nxtY < 800 Then nxtY = 800
    If nxtY > (picFiles.ScaleHeight - 800) Then nxtY = picFiles.Height - 800
    picSizer.Top = nxtY
    imgSize.Move picSizer.Left, picSizer.Top, picSizer.Width, picSizer.Height
  End If
End Sub

Private Sub imgSize_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
  picSizer.Visible = False
  Resize
End Sub

Private Sub MDIForm_Load()
  Setup
  flList.path = App.path & "\java"
  pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
  pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
  pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
  pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY
  imgSize.Top = 1920
  Dir1_Change
  addTags
  Resize
  Me.Visible = True
  mnuNew_Click

End Sub

Public Sub Setup()
  On Error Resume Next
  LoadDirectory App.path & "\highlighters"
  Call SetupMenu
  Me.WindowState = GetSetting("ScintillaMDI", "Settings", "MDIWState", 0)
  Me.Left = GetSetting("ScintillaMDI", "Settings", "MDILeft", (Screen.Width - Me.Width) \ 2)
  Me.Top = GetSetting("ScintillaMDI", "Settings", "MDITop", (Screen.Height - Me.Height) \ 2)
  Me.Width = GetSetting("ScintillaMDI", "Settings", "MDIWidth", Me.Width)
  Me.Height = GetSetting("ScintillaMDI", "Settings", "MDIHeight", Me.Height)
  
  Me.Arrange vbCascade
End Sub

Public Function AddMenu(sCaption As String, sTag As String, iIndex As Integer) As Integer
End Function

Public Function SetupMenu()
End Function

Private Sub MDIForm_Unload(Cancel As Integer)
  If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
  SaveSetting "ScintillaMDI", "Settings", "MDIWState", Me.WindowState
  Me.WindowState = vbNormal
  SaveSetting "ScintillaMDI", "Settings", "MDILeft", Me.Left
  SaveSetting "ScintillaMDI", "Settings", "MDITop", Me.Top
  SaveSetting "ScintillaMDI", "Settings", "MDIWidth", Me.Width
  SaveSetting "ScintillaMDI", "Settings", "MDIHeight", Me.Height
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal, Me
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuArrange_Click()
  Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuBold_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelText = "<B></B>"
  Call ActiveForm.sciMain.SetSel(ActiveForm.sciMain.GetCurPos - 4, ActiveForm.sciMain.GetCurPos - 4)
End Sub

Private Sub mnuCascade_Click()
  Me.Arrange vbCascade
End Sub

Private Sub mnuCenter_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelText = "<P ALIGN=" & strQ & "CENTER" & strQ & "></P>"
  Call ActiveForm.sciMain.SetSel(ActiveForm.sciMain.GetCurPos - 4, ActiveForm.sciMain.GetCurPos - 4)
End Sub

Private Sub mnuCopy_Click()
  On Error Resume Next
  ActiveForm.sciMain.Copy
End Sub

Private Sub mnuCut_Click()
  On Error Resume Next
  ActiveForm.sciMain.Cut
End Sub

Private Sub mnuExportHTML_Click()
  On Error Resume Next
  With cd
    .Filter = "HTML Files (*.html, *.htm)|*.html;*.htm)|All Files (*.*)|*.*"
    .ShowSave
    If .filename <> "" Then
      ExportToHTML .filename, ActiveForm.sciMain
    End If
    ActiveForm.sciMain.SetFocus
  End With
End Sub

Private Sub mnuFind_Click()
  On Error Resume Next
  ActiveForm.sciMain.DoFind
End Sub

Private Sub mnuGoto_Click()
  On Error Resume Next
  ActiveForm.sciMain.DoGoto
End Sub

Private Sub mnuHighlighter_Click(index As Integer)
End Sub

Private Sub mnuHor_Click()
  Me.Arrange vbHorizontal
End Sub

Private Sub mnuHyperlink_Click()
  On Error Resume Next
  frmHyperlink.Show vbModal, Me
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuImage_Click()
  frmImage.Show vbModal, Me
End Sub

Private Sub mnuItalic_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelText = "<I></I>"
  Call ActiveForm.sciMain.SetSel(ActiveForm.sciMain.GetCurPos - 4, ActiveForm.sciMain.GetCurPos - 4)
End Sub

Private Sub mnuLeft_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelText = "<P ALIGN=" & strQ & "LEFT" & strQ & "></P>"
  Call ActiveForm.sciMain.SetSel(ActiveForm.sciMain.GetCurPos - 4, ActiveForm.sciMain.GetCurPos - 4)
End Sub

Private Sub mnuNew_Click()
  frmNew.Show vbModal, Me
'  NewDoc "New Document"
'  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuOpen_Click()
  Dim i As Long
  With cd
    .Filter = "All Files (*.*)|*.*|"
    For i = 0 To UBound(Highlighters) - 1
      If Highlighters(i).strFilter <> "" Then .Filter = .Filter & Highlighters(i).strFilter
    Next i
    .ShowOpen
    If .filename <> "" Then NewDoc .filename
  End With
  ActiveForm.sciMain.SetFocus
End Sub

Public Sub NewDoc(strFile As String)
  On Error Resume Next
  Static lDocumentCount As Long
  Dim doc As New frmDoc
  Load doc
  lDocumentCount = lDocumentCount + 1
  If Dir(strFile) <> "" Then
    doc.sciMain.LoadFile (strFile)
    doc.Caption = strFile
    doc.strFile = strFile
    SetHighlighter doc.sciMain, SetHighlighterBasedOnExtension(strFile)
  Else
    doc.Caption = strFile & " " & lDocumentCount
  End If
  doc.Show
  doc.sciMain.SetFocus
End Sub

Private Sub mnuPaste_Click()
  On Error Resume Next
  ActiveForm.sciMain.Paste
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuPrint_Click()
  On Error Resume Next
  ActiveForm.sciMain.PrintDoc
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuRedo_Click()
  On Error Resume Next
  ActiveForm.sciMain.Redo
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuReplace_Click()
  On Error Resume Next
  ActiveForm.sciMain.DoReplace
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuRight_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelText = "<P ALIGN=" & strQ & "RIGHT" & strQ & "></P>"
  Call ActiveForm.sciMain.SetSel(ActiveForm.sciMain.GetCurPos - 4, ActiveForm.sciMain.GetCurPos - 4)
End Sub

Private Sub mnuSave_Click()
  On Error Resume Next
  If ActiveForm.strFile <> "" Then
    ActiveForm.sciMain.SaveToFile (ActiveForm.strFile)
  Else
    DoSaveAs
  End If
  ActiveForm.sciMain.SetFocus
End Sub

Public Sub SaveDoc()
  mnuSave_Click
  ActiveForm.sciMain.SetFocus
End Sub

Public Sub DoSaveAs()
  On Error Resume Next
  Dim i As Long
  With cd
    .Filter = "All Files (*.*)|*.*|"
    For i = 0 To UBound(Highlighters) - 1
      If Highlighters(i).strFilter <> "" Then .Filter = .Filter & Highlighters(i).strFilter
    Next i
    .ShowSave
    If .filename <> "" Then
      ActiveForm.sciMain.SaveToFile .filename
      ActiveForm.strFile = .filename
      ActiveForm.Caption = .filename
      SetHighlighter ActiveForm.sciMain, SetHighlighterBasedOnExtension(.filename)
    End If
  End With
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuSaveAs_Click()
  On Error Resume Next
  DoSaveAs
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuSelAll_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelectAll
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuStrike_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelText = "<S></S>"
  Call ActiveForm.sciMain.SetSel(ActiveForm.sciMain.GetCurPos - 4, ActiveForm.sciMain.GetCurPos - 4)
End Sub

Private Sub mnuSyntax_Click()
  DoSyntaxOptions App.path & "\highlighters\", Me
  ResetSyntaxMDI "frmDoc"
  ActiveForm.sciMain.SetFocus
  SetupMenu
End Sub

Private Sub mnuTable_Click()
  On Error Resume Next
  Load frmInsertTable
  frmInsertTable.ColWidth = Space(ActiveForm.sciMain.GetColumn)
  frmInsertTable.Show vbModal, Me
End Sub

Private Sub mnuTimeDate_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelText = Time & " | " & Date
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuUnderline_Click()
  On Error Resume Next
  ActiveForm.sciMain.SelText = "<U></U>"
  Call ActiveForm.sciMain.SetSel(ActiveForm.sciMain.GetCurPos - 4, ActiveForm.sciMain.GetCurPos - 4)
End Sub

Private Sub mnuUndo_Click()
  On Error Resume Next
  ActiveForm.sciMain.Undo
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub mnuVertical_Click()
  Me.Arrange vbVertical
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub picFiles_Resize()
  Resize
End Sub

Private Sub picJava_Resize()
  flList.Move 0, 0, picJava.ScaleWidth, picJava.ScaleHeight
End Sub

Private Sub picLeft_Resize()
  tbsSide.Move 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
  picFiles.Move tbsSide.ClientLeft, tbsSide.ClientTop, tbsSide.ClientWidth, tbsSide.ClientHeight
  Picture5.Move tbsSide.ClientLeft, tbsSide.ClientTop, tbsSide.ClientWidth, tbsSide.ClientHeight
  picJava.Move tbsSide.ClientLeft, tbsSide.ClientTop, tbsSide.ClientWidth, tbsSide.ClientHeight
End Sub

Private Sub picSize_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    picSize.BackColor = &H8000000C
    picLeft.Width = picLeft.Width + x
  End If
End Sub

Private Sub picSize_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    picSize.BackColor = &H8000000F
  End If
End Sub

Private Sub Resize()
  On Error Resume Next
  imgSize.Left = 0
  imgSize.Width = picFiles.ScaleWidth
  picSizer.Move 0, imgSize.Top, imgSize.Width, imgSize.Height
  Drive1.Move 0, 30, picFiles.ScaleWidth
  Dir1.Move 0, Drive1.Top + Drive1.Height + 30, picFiles.ScaleWidth, imgSize.Top - Dir1.Top - 15
  If Dir1.Height > (picFiles.ScaleHeight - 1500) Then Dir1.Height = picFiles.ScaleHeight - 1500
  imgSize.Move 0, Dir1.Top + Dir1.Height, picFiles.ScaleWidth, 50
  File1.Move 0, imgSize.Top + imgSize.Height + 15, picFiles.ScaleWidth, picFiles.Height - (imgSize.Top + imgSize.Height + 15)
  File1.ColumnHeaders(1).Width = File1.Width - 350
End Sub

Private Sub Picture5_Resize()
  TagsD.Move 0, 30, Picture5.ScaleWidth, Picture5.ScaleHeight - 30
End Sub

Private Sub TagsD_DblClick()
  Dim timedate As String
  On Error Resume Next
  timedate = TagsD.SelectedItem.Text
  ActiveForm.sciMain.SelText = timedate
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub tbrFormat_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case LCase(Button.key)
    Case "bold"
      mnuBold_Click
    Case "italic"
      mnuItalic_Click
    Case "underline"
      mnuUnderline_Click
    Case "strike"
      mnuStrike_Click
    Case "left"
      mnuLeft_Click
    Case "center"
      mnuCenter_Click
    Case "right"
      mnuRight_Click
  End Select
End Sub

Private Sub tbrInsert_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case LCase(Button.key)
    Case "link"
      mnuHyperlink_Click
    Case "image"
      mnuImage_Click
    Case "table"
      mnuTable_Click
  End Select
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next

  
  Select Case LCase(Button.key)
    Case "new"
      mnuNew_Click
    Case "open"
      mnuOpen_Click
    Case "save"
      mnuSave_Click
    Case "print"
      mnuPrint_Click
    Case "cut"
      mnuCut_Click
    Case "copy"
      mnuCopy_Click
    Case "paste"
      mnuPaste_Click
    Case "delete"
      ActiveForm.sciMain.Delete
  End Select
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub tbrMain_ButtonDropDown(ByVal Button As MSComctlLib.Button)
  Dim x As Long, y As Long
  On Error Resume Next
  Select Case Button.key
    Case "qtable"
      Dim p As Rect
      GetWindowRect tbrMain.hWnd, p
      x = p.Left * (Screen.TwipsPerPixelX)
      y = p.Top * (Screen.TwipsPerPixelY)
      frmTable.Top = y + Button.Height
      frmTable.Left = x + Button.Left
      frmTable.ColWidth = Space(ActiveForm.sciMain.GetColumn)
      frmTable.Show
  End Select
End Sub

Private Sub tbrMain_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
  On Error Resume Next
  frmTable.Hide
  Unload frmTable
  ActiveForm.sciMain.SetFocus
End Sub

Private Sub tbsSide_Click()
  picFiles.Visible = False
  Picture5.Visible = False
  picJava.Visible = False
  Select Case tbsSide.SelectedItem.index
    Case 1
      picFiles.Visible = True
    Case 2
      Picture5.Visible = True
    Case 3
      picJava.Visible = True
  End Select
End Sub

Private Sub Initialise()
'-----------------------------------------------
'Initialise the controls
'-----------------------------------------------
On Local Error Resume Next

'Break the link to iml lists
File1.ListItems.Clear
File1.Icons = Nothing
File1.SmallIcons = Nothing

'Clear the image lists
iml32.ListImages.Clear
iml16.ListImages.Clear

End Sub

Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim filename As String

On Local Error Resume Next
For Each Item In File1.ListItems
  filename = Item.SubItems(1) & Item.Text
  GetIcon filename, Item.index
Next

End Sub

Private Function GetIcon(filename As String, index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection



'Get a handle to the small icon
hSIcon = SHGetFileInfo(filename, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(filename, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    ImageList_Draw hLIcon, ShInfo.iIcon, pic32.hDc, 0, 0, ILD_TRANSPARENT
    .Refresh
  End With
  'Small Icon
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    ImageList_Draw hSIcon, ShInfo.iIcon, pic16.hDc, 0, 0, ILD_TRANSPARENT
    .Refresh
  End With
  Set imgObj = iml32.ListImages.Add(index, , pic32.Image)
  Set imgObj = iml16.ListImages.Add(index, , pic16.Image)
End If
End Function
Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the File1
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With File1
  '.ListItems.Clear
  .Icons = iml32        'Large
  .SmallIcons = iml16   'Small
  For Each Item In .ListItems
    Item.Icon = Item.index
    Item.SmallIcon = Item.index
  Next
End With

End Sub

Sub FillFile1WithFiles(ByVal path As String)
'-------------------------------------------
'Scan the selected folder for files
'and add then to the listview
'-------------------------------------------
Dim Item As ListItem
Dim S As String

path = CheckPath(path)    'Add '\' to end if not present
S = Dir(path, vbNormal)
Do While S <> ""
  Set Item = File1.ListItems.Add(, , S)
  Item.key = path & S
  'Item.SmallIcon = "Folder"
  Item.Text = S
  Item.SubItems(1) = path
  S = Dir
Loop

End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

