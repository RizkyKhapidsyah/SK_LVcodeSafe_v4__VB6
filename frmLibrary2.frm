VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.MDIForm frmLibrary 
   BackColor       =   &H8000000C&
   Caption         =   "LaVolpe Source Code Repository"
   ClientHeight    =   10335
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13260
   HelpContextID   =   1
   Icon            =   "frmLibrary2.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList LargeImages 
      Left            =   5955
      Top             =   2355
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":09CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":0E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":16FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":1FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":28B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":318E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":3A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":3D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":40B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":43D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":46F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":4FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":58AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":6186
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":6A62
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":733E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":7C1A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   5985
      Top             =   1185
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picFilters 
      Align           =   3  'Align Left
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   9630
      HelpContextID   =   1
      Left            =   2865
      ScaleHeight     =   9570
      ScaleWidth      =   2790
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   2850
      Begin VB.ComboBox cboFilter 
         Height          =   315
         Index           =   0
         ItemData        =   "frmLibrary2.frx":85F6
         Left            =   315
         List            =   "frmLibrary2.frx":8609
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   7290
         Width           =   2100
      End
      Begin VB.ListBox lstFilter 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6585
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   420
         Width           =   2640
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         Index           =   1
         ItemData        =   "frmLibrary2.frx":8678
         Left            =   330
         List            =   "frmLibrary2.frx":8685
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   8625
         Width           =   2100
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         Index           =   2
         Left            =   315
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   7950
         Width           =   2100
      End
      Begin VB.CommandButton cmdFilterDefault 
         Caption         =   "Remove Filters"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   330
         TabIndex        =   9
         Tag             =   "NoFilter"
         Top             =   9060
         Width           =   1725
      End
      Begin VB.CommandButton cmdCloseFilters 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2055
         TabIndex        =   8
         Tag             =   "Hide Filter Window"
         ToolTipText     =   "Close Filters Panel"
         Top             =   9060
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sort Order"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   315
         TabIndex        =   17
         Top             =   7005
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FILTER(S)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   60
         TabIndex        =   16
         Top             =   135
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Categories"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1395
         TabIndex        =   12
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Attachments"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   345
         TabIndex        =   6
         Top             =   8355
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Language"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   315
         TabIndex        =   5
         Top             =   7665
         Width           =   825
      End
   End
   Begin VB.PictureBox picMain 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   9630
      Left            =   0
      ScaleHeight     =   9570
      ScaleWidth      =   2805
      TabIndex        =   2
      Top             =   360
      Width           =   2865
      Begin RichTextLib.RichTextBox rtfStaging 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   195
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   8415
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   979
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   0   'False
         ReadOnly        =   -1  'True
         RightMargin     =   32650
         TextRTF         =   $"frmLibrary2.frx":86B6
      End
      Begin MSComctlLib.ListView lvCode 
         Height          =   7905
         Left            =   180
         TabIndex        =   0
         Top             =   420
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   13944
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "SmallImages"
         ForeColor       =   -2147483640
         BackColor       =   16776960
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code Title"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Code (right click for action)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   120
         Width           =   2475
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   9990
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Database"
            TextSave        =   "Database"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16404
            MinWidth        =   5080
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1323
            MinWidth        =   1323
            Text            =   "Size"
            TextSave        =   "Size"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   5955
      Top             =   1740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":877F
            Key             =   "IMG1"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":8A99
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":9375
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":9C51
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":A52D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":AE09
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":B6ED
            Key             =   "IMG7"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":BA07
            Key             =   "IMG8"
            Object.Tag             =   "9"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":BD21
            Key             =   "IMG9"
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":C03B
            Key             =   "IMG10"
            Object.Tag             =   "10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":C355
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":CC39
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":D515
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":DDF1
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":E6CD
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":EFB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":F88D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":F9E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":102C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":10BA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":110E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":119C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":1229D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":123B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":1250D
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":12DE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":1323D
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":13399
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":13D75
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":141C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary2.frx":1461D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbTools 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OtherApps"
            Object.ToolTipText     =   "Other Applications"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "URLwindow"
            Object.ToolTipText     =   "Opens Web Links"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Settings"
            Object.ToolTipText     =   "Opens the Settings Window"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Object.ToolTipText     =   "Add new Code"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AddManual"
                  Text            =   "Manual Entry"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AddAuto"
                  Text            =   "From Existing Files"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "List"
            Object.ToolTipText     =   "Display Selected Code"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Search for specific code"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Filter"
            Object.ToolTipText     =   "Filter the Code Listing"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cascade"
            Object.ToolTipText     =   "Cascade Windows"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TileH"
            Object.ToolTipText     =   "Tile Windows Horizontally"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TileV"
            Object.ToolTipText     =   "Tile Windows Vertically"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SelWindow"
            Object.ToolTipText     =   "Activate Window..."
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CloseAll"
            Object.ToolTipText     =   "Close All Windows"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Display Help File"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "Display About Dialog"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "Open &Web Links"
         Index           =   0
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Open &Database"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Compact Database"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Import/Export Wizard"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   7
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Edit"
      Index           =   1
      Begin VB.Menu mnuListing 
         Caption         =   "&View/Edit Code"
         Index           =   0
      End
      Begin VB.Menu mnuListing 
         Caption         =   "&Add New Code"
         Index           =   1
      End
      Begin VB.Menu mnuListing 
         Caption         =   "Add from VB &Files"
         Index           =   2
      End
      Begin VB.Menu mnuListing 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuListing 
         Caption         =   "&Fiind"
         Index           =   4
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuListing 
         Caption         =   "Fi&lter"
         Index           =   5
      End
      Begin VB.Menu mnuListing 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuListing 
         Caption         =   "Refresh Code Listing"
         Index           =   7
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Options"
      Index           =   2
      Begin VB.Menu mnuSettings 
         Caption         =   "Customize &Options"
         Index           =   0
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Category Updates"
         Index           =   2
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Language Updates"
         Index           =   3
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Other &Applications"
         Index           =   4
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Other &Applications"
      Index           =   3
      Begin VB.Menu mnuOtherApps 
         Caption         =   "&Add - Edit - Delete from Listing"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Window"
      Index           =   4
      Begin VB.Menu mnuWindow 
         Caption         =   "Active &Window"
         Index           =   0
         WindowList      =   -1  'True
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "&Cascade"
         Index           =   2
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "Tile &Horizontally"
         Index           =   3
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "Tile &Veritcally"
         Index           =   4
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "Close &All Windows"
         Index           =   6
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Help"
      Index           =   5
      Begin VB.Menu mnuHelp 
         Caption         =   "Help &Topics"
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   2
      End
   End
   Begin VB.Menu mnuViewRecord 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuViewSub 
         Caption         =   "&Copy"
         Index           =   0
         Begin VB.Menu mnuViewCopy 
            Caption         =   "Copy Selected Text     Ctrl+C"
            Index           =   0
         End
         Begin VB.Menu mnuViewCopy 
            Caption         =   "Copy All Text               Ctrl+A, Ctrl+C"
            Index           =   1
         End
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "&Paste               Ctr+V"
         Index           =   1
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "&Font"
         Index           =   3
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Word &Wrap"
         Index           =   4
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "&Recolor Text"
         Index           =   5
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Find New"
         Index           =   7
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Find Next"
         Index           =   8
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Save"
         Index           =   10
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Save As..."
         Index           =   11
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Save to File"
         Index           =   12
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Edit Categories"
         Index           =   14
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Edit Languages"
         Index           =   15
      End
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private iAddType As Integer

Private Sub cboFilter_Click(Index As Integer)

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo cboFilter_Click_General_ErrTrap

If Len(cmdFilterDefault(2).Tag) Then Exit Sub
FilterRecordsetNow
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
cboFilter_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: cboFilter_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub cmdCloseFilters_Click()
'============================================================
' Simply hide & disable the picture box containing the filter options
'============================================================

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo cmdCloseFilters_Click_General_ErrTrap

picFilters.Enabled = False: picFilters.Visible = False
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
cmdCloseFilters_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: cmdCloseFilters_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub cmdFilterDefault_Click(Index As Integer)

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo cmdFilterDefault_Click_General_ErrTrap

FilterRecordsetNow True
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
cmdFilterDefault_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: cmdFilterDefault_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Public Sub FilterRecordsetNow(Optional bReset As Boolean = False)
'============================================================
' Filter the database per user's request
'============================================================

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo FilterRecordsetNow_General_ErrTrap

    If mainDB Is Nothing Then Exit Sub                                  ' Abort if no valid database loaded
    Dim I As Integer
    If bReset = True Then
        cmdFilterDefault(2).Tag = "NoFilter"
        lstFilter.ListIndex = 0
        cboFilter(0).ListIndex = 0
        cboFilter(1).ListIndex = 0
        cboFilter(2).ListIndex = 0
        cmdFilterDefault(2).Tag = ""
    End If
    mainFilterIndex = lstFilter & "|" _
            & lstFilter.ItemData(lstFilter.ListIndex) & "|"
    For I = 1 To 2          ' Construct the filter string which will be parsed in another function
        ' Filter string is:  category | attachments | languages
        mainFilterIndex = mainFilterIndex & cboFilter(I) & "|" _
            & cboFilter(I).ItemData(cboFilter(I).ListIndex) & "|"
    Next
    MousePointer = vbHourglass
    FilterRecordset                         ' Call function filter the database
    MousePointer = vbDefault
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
FilterRecordsetNow_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: FilterRecordsetNow" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub lstFilter_Click()

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo lstFilter_Click_General_ErrTrap

If Len(cmdFilterDefault(2).Tag) Then Exit Sub
FilterRecordsetNow
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
lstFilter_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: lstFilter_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub lvCode_DblClick()
'============================================================
' Allow a double click on the code listview to view selected code
'============================================================

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo lvCode_DblClick_General_ErrTrap

If lvCode.ListItems.Count = 0 Then Exit Sub     ' of course if no code in the listview then abort
mnuListing_Click (0)    ' call sub to display the code
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
lvCode_DblClick_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: lvCode_DblClick" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub lvCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'============================================================
' option gives the user a right click function within the code listing
'============================================================

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo lvCode_MouseUp_General_ErrTrap

If lvCode.ListItems.Count = 0 Then Exit Sub         ' No code in the listing, no right click
If Button = 2 Then PopupMenu mnuMain(1)         ' If right clicked (2), then call the menu
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
lvCode_MouseUp_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: lvCode_MouseUp" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub MDIForm_Load()
'============================================================
' This form is loaded from the Main Sub() and all other functions are performed in that sub, except
'   the applying icons to the menu bar
'============================================================

' Retrieve stored user preference on button size, using Large buttons as a default if this is first run

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo MDIForm_Load_General_ErrTrap

MyDefaults.ButtonSize = CInt(GetSetting("LaVolpeCodeSafe", "Settings", "ButtonSize", "2"))
ApplyButtonFaces MyDefaults.ButtonSize      ' Call function to display the appropriate buttons

' Color the filters picture box to match the rest of the form
Dim I As Integer
For I = Label1.LBound + 1 To Label1.UBound: Label1(I).ForeColor = MyDefaults.LblColorMain: Next
Icon = SmallImages.ListImages(28).ExtractIcon
If MyDefaults.WindowColor = WinBlahColor Then Label1(0).ForeColor = &HFFFFFF Else Label1(0).ForeColor = MyDefaults.LblColorMain
picMain.Align = MyDefaults.Align
picFilters.Align = MyDefaults.Align
Caption = Caption & ", v" & App.Major & "." & App.Minor
iAddType = Val(GetSetting("LaVolpeCodeSafe", "Settings", "AddType", "0"))
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
MDIForm_Load_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: MDIForm_Load" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub MDIForm_Terminate()
'============================================================
' Ensure the forms are unloaded from memory if user clicked the X
'============================================================

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo MDIForm_Terminate_General_ErrTrap

Unload Me
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
MDIForm_Terminate_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: MDIForm_Terminate" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'============================================================
' When unloading, ensure all open forms are also unloaded
'============================================================

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo MDIForm_Unload_General_ErrTrap

Cancel = CloseAllWindows(True)    ' Call function to close all windows, including this window
If Cancel > 0 Then Exit Sub             ' If an open window had unsaved data, stop here & allow data to be saved
On Error Resume Next
mainRS.Close                                  ' Close the main recordset
Set mainRS = Nothing
mainDB.Close                                 ' Close the main database
Set mainDB = Nothing
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
MDIForm_Unload_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: MDIForm_Unload" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub mnuFile_Click(Index As Integer)

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo mnuFile_Click_General_ErrTrap

Select Case Index
Case 0:         ' Web Links
    If IsCheckedOut("frmWeb") Then frmWeb.SetFocus Else frmWeb.Show
Case 2:          ' Open database for editing
    If MsgBox("Use the Options menu button to change the database associated with this program. " _
        & "Press OK to open the current database", vbOKCancel + vbInformation) = vbCancel Then Exit Sub
    OpenThisFile MyDefaults.DBname, 1, "", hWnd
Case 3:
    If Forms.Count > 1 Then
        MsgBox "You must close all windows within this application. " & vbCrLf & _
            "Since each window uses a link to the database, these must be closed before compacting or repairing.", vbInformation + vbOKOnly
        Exit Sub
    End If
    Dim sTempDB As String
    If MyDefaults.DBname = "" Then
        MsgBox "First connect to a valid database.", vbExclamation + vbOKOnly
        Exit Sub
    End If
    mainRS.Close
    mainDB.Close
    Set mainDB = Nothing
    sTempDB = GetTempFolder
    sTempDB = sTempDB & "~LVcsDB.tmp"
    If Len(Dir(sTempDB)) Then Kill sTempDB
    If MsgBox("Will now attempt to compact & connect to compacted database.", vbInformation + vbOKCancel) = vbCancel Then Exit Sub
    DBEngine.CompactDatabase MyDefaults.DBname, sTempDB
    If Len(Dir(sTempDB)) Then
        GP = sTempDB
        On Error Resume Next
        frmDBprompt.Show 1, Me
        Set mainDB = Nothing
        On Error GoTo FailedCompact
        If IsNull(GP) Then  ' Failed connect to compacted db
            MsgBox "Couldn't connect to the compacted database. " & vbCrLf & "Will now reconnect to current database.", vbInformation + vbOKOnly
        Else
            Name MyDefaults.DBname As MyDefaults.DBname & ".bak"
            Name sTempDB As MyDefaults.DBname
            Kill MyDefaults.DBname & ".bak"
            MsgBox "Compacted successfully. Now re-establishing connection.", vbInformation + vbOKOnly
        End If
    Else
        MsgBox "Failed to compact the database. Try manually via Microsoft Access.", vbInformation + vbOKOnly
    End If
    On Error Resume Next
    Set mainDB = Nothing
    Call Main
Case 5              ' Import/Export wizard
    On Error Resume Next
    frmPkgWizard.Show 1, Me
    If InStr(GP, "Requery") Then
        Call mnuListing_Click(7)
        If GP = "Requery with Web" Then frmWeb.Show
    End If
Case 7:             ' Exit
    End
End Select
Exit Sub

FailedCompact:
If Len(Dir(MyDefaults.DBname)) = 0 Then
    MsgBox "The system failed to rename your compacted database back to its original name." & vbCrLf & _
        "Will now attempt to retrieve original database.  Should you not be able to access" & vbCrLf & _
        "any of your records without errors, complete the steps below. Otherwise error corrected." & vbCrLf & vbCrLf & _
        "1. Close this application" & vbCrLf & _
        "2. Use your Windows Explorer and look for the file named " & vbCrLf & "       " & MyDefaults.DBname & ".bak" & vbCrLf & _
        "3. Rename that file by removing the .bak from the end of the file name." & vbCrLf & _
        "4. Reopen this application", vbExclamation + vbOKOnly
    If Len(Dir(MyDefaults.DBname & ".bak")) Then Name MyDefaults.DBname As Left(MyDefaults.DBname, Len(MyDefaults.DBname) - 4)
End If
Set mainDB = Nothing
Call Main
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
mnuFile_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: mnuFile_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub mnuHelp_Click(Index As Integer)

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo mnuHelp_Click_General_ErrTrap

If Index = 2 Then
    frmAbout.Show 1, Me
Else
    If App.HelpFile = "" Then
        MsgBox "The help file is available by request.  Just want to know who's using this application." _
            & vbCrLf & vbCrLf & "Look at the About form to email me.", vbInformation + vbOKOnly
    Else
        SendKeys "{F1}"
    End If
End If
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
mnuHelp_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: mnuHelp_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub mnuListing_Click(Index As Integer)
'============================================================
' Functions used to display exsting code, new code, filter options, or refresh the code listing
'============================================================

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo mnuListing_Click_General_ErrTrap

Select Case Index
    Case 0, 1       ' existing code or new code
        If Index = 0 Then
            If lvCode.ListItems.Count = 0 Then Exit Sub
            ShowCode lvCode.SelectedItem.Key
        Else
            ShowCode "New"
        End If
    Case 2  ' Add from existing VB project files
        frmLoadExisting.Show
    Case 4      ' Find
        If IsCheckedOut("frmLibSearch") Then frmLibSearch.SetFocus Else frmLibSearch.Show
    Case 5      ' Filters
        If picFilters.Visible = True Then                       ' If filter options are open then toggle them closed
            Call cmdCloseFilters_Click
        Else                                                                ' Otherwise, show & enable the filter options
            picFilters.Enabled = True: picFilters.Visible = True
        End If
    Case 7       ' Refresh the listings
        On Error GoTo FailedAction
        If CloseAllWindows(False, "frmClipboard") = 1 Then Exit Sub
        mainRS.Requery
        FilterRecordset             ' Call function to repopulate the code listing & filter the recordset to existing filter(s)
End Select
Exit Sub

FailedAction:
MsgBox Err.Description, vbExclamation + vbOKOnly
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
mnuListing_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: mnuListing_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub mnuOtherApps_Click(Index As Integer)

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo mnuOtherApps_Click_General_ErrTrap

If Index = 1 Then Exit Sub
If Index = 0 Then
    If IsCheckedOut("frmOtherApps") Then frmOtherApps.SetFocus Else frmOtherApps.Show
Else
    OpenThisFile mnuOtherApps(Index).Tag, 1, "", hWnd
End If
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
mnuOtherApps_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: mnuOtherApps_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub mnuSettings_Click(Index As Integer)

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo mnuSettings_Click_General_ErrTrap

Select Case Index
Case 0
    GP = Null
    frmOptions.Show 1, Me                   ' Set global variable & open the Options window
    If GP = "Reload" Then                      ' If the user made changes then the GP variable = Reload
        DoGradient picMain, 1, True         '   and if so, repaint with new colors
        DoGradient picFilters, 1, True
        Dim I As Integer
        For I = Label1.LBound + 1 To Label1.UBound: Label1(I).ForeColor = MyDefaults.LblColorMain: Next
        If MyDefaults.WindowColor = WinBlahColor Then Label1(0).ForeColor = &HFFFFFF Else Label1(0).ForeColor = MyDefaults.LblColorMain
        picMain.Align = MyDefaults.Align
        picFilters.Align = MyDefaults.Align
        iAddType = Val(GetSetting("LaVolpeCodeSafe", "Settings", "AddType", "0"))
        ApplyButtonFaces MyDefaults.ButtonSize  ' and ensure button size to user's request
    End If
Case 2, 3
    GP = Choose(Index + 1, Null, Null, "Cats", "Lang")
    frmCats.Show 1, Me
Case 4
    Call mnuOtherApps_Click(0)
End Select
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
mnuSettings_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: mnuSettings_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub mnuViewCopy_Click(Index As Integer)
mnuViewRecord.Tag = Index + 30
End Sub

Private Sub mnuViewSub_Click(Index As Integer)
If Index Then mnuViewRecord.Tag = Index
End Sub

Private Sub mnuWindow_Click(Index As Integer)
'============================================================
' Cascade, Veritcal, & Horizontal child window options
'============================================================

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo mnuWindow_Click_General_ErrTrap

If Index > 1 Then
    On Error Resume Next
    If Index = 6 Then
        CloseAllWindows False, "frmClipboard"
    Else
        Arrange Choose(Index - 1, vbCascade, vbTileVertical, vbTileHorizontal)
    End If
End If
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
mnuWindow_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: mnuWindow_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub


Private Sub picMain_Resize()
'============================================================
'  Whenever the form resizes, adjust the code listing to fit the window & repaint the window
'============================================================
On Error Resume Next
Dim iOffset As Integer
iOffset = 700
DoGradient picMain, 1, True                     ' Repaint the window
DoGradient picFilters, 1, True
lvCode.Height = picMain.Height - 800      ' Adjust the code listing size
lstFilter.Height = lvCode.Height - 2300
Label1(5).Top = lstFilter.Height + lstFilter.Top + 15
cboFilter(0).Top = Label1(5).Top + 285
Label1(1).Top = cboFilter(0).Top + 405
cboFilter(2).Top = Label1(1).Top + 285
Label1(2).Top = cboFilter(2).Top + 405
cboFilter(1).Top = Label1(2).Top + 270
cmdFilterDefault(2).Top = cboFilter(1).Top + 420
cmdCloseFilters.Top = cmdFilterDefault(2).Top
DoEvents
End Sub

Private Sub tbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    '============================================================
    ' This handles the events sent when the user click on the
    ' toolbar.  Basically, it redirects the input based on the
    ' menus.
    '============================================================
    Select Case Button.Key
        Case "Open"                                         ' Opens the database for direct editing
            Call mnuFile_Click(2)
        Case "URLwindow"                            ' Opens the Web Address window
            Call mnuFile_Click(0)
        Case "Settings"                                     ' Opens the Options window
            Call mnuSettings_Click(0)
        Case "Add"                                          ' Button to add new code to the database
            Call mnuListing_Click(1 + iAddType)
        Case "List"                                           ' Button to show existing code
            Call mnuListing_Click(0)
        Case "Find"                                         ' Button to show a search form
            Call mnuListing_Click(4)
        Case "Filter"                                       ' Button to toggle filter options display
            Call mnuListing_Click(5)
        Case "Cascade"                                  ' Button to cascade child windows
            Arrange vbCascade
        Case "TileV"                                        ' Button to veritcally tile buttons
            Arrange vbTileVertical
        Case "TileH"                                        ' Button to horizontally tile buttons
            Arrange vbTileHorizontal
        Case "CloseAll"                                     ' Button to close all windows
            CloseAllWindows False, "frmClipboard"           ' uses the False switch to prevent main window from closing
        Case "About"                                        ' Button to display the self-gratifying About box
           Call mnuHelp_Click(2)
        Case "Help"                                         ' Button to display the compiled help file
            Call mnuHelp_Click(0)
        Case "OtherApps"                                  ' Button opens a window containing user-defined links to other applications
            Call mnuOtherApps_Click(0)
    End Select
End Sub

Public Sub ApplyButtonFaces(iBigSmall As Integer)
'============================================================
'   Function simply applies either 16x16 icons to 32x32 icons to the menubar
'============================================================

Dim I As Integer

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo ApplyButtonFaces_General_ErrTrap

Set tbTools.ImageList = Nothing                                                             ' First clear existing assignment
If iBigSmall = 0 Then
    tbTools.Enabled = False: tbTools.Visible = False
    Exit Sub
Else
    tbTools.Enabled = True: tbTools.Visible = True
End If
Set tbTools.ImageList = Choose(iBigSmall, SmallImages, LargeImages) ' Now assign new image list
' Loop thru the buttons assigning the appropriate button images
With tbTools.Buttons    ' note some button assignments are 0, meaning no icon will be assigned
    For I = 1 To 17: .Item(I).Image = Choose(I, 16, 3, 2, 0, 4, 5, 6, 15, 0, 7, 9, 8, 0, 10, 0, 11, 12): Next
End With
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
ApplyButtonFaces_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: ApplyButtonFaces" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Function IsCheckedOut(sRecID As String) As Boolean
'============================================================
'   Function sees if an existing code is already displayed in a child window
'============================================================
Dim I As Integer

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo IsCheckedOut_General_ErrTrap

For I = 0 To Forms.Count - 1            ' loop thru each loaded form
    If Left(sRecID, 6) <> "RecID:" Then
        If Forms(I).Name = sRecID Then
            IsCheckedOut = True
            Forms(I).WindowState = 0
            Exit Function
        End If
    End If
    If Forms(I).Tag = sRecID Then       ' see if the Tag property matches the record in question
        IsCheckedOut = True                 ' if so, set the value to true & exit
        If MsgBox("That code is already in one of the open windows. " & vbCrLf & vbCrLf _
            & "Do you want to go to it?", vbQuestion + vbYesNo) = vbYes Then
                Forms(I).WindowState = 0
                Forms(I).SetFocus
        End If
        Exit Function
    End If
Next
Exit Function

' Inserted by LaVolpe OnError Insertion Program.
IsCheckedOut_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: IsCheckedOut" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Function

Public Sub ShowCode(recID As String)
    ' Allow only one instance of displaying this code

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo ShowCode_General_ErrTrap

    If IsCheckedOut(recID) Then Exit Sub
    On Error GoTo FailedAction
    Dim frmView As New frmLibViewRec            ' dimension a copy of the form to display
    If recID = "New" Then                                        ' New code, therefore no db record to reference
        DBrecID = Null
    Else
        DBrecID = Val(Mid(recID, 7))                    ' Reference the db record ID for the selected code
    End If
    MousePointer = vbHourglass
    On Error Resume Next
    frmView.Show                                                ' Open the window
    MousePointer = vbDefault
Exit Sub

FailedAction:
MsgBox "Following error is preventing the requested action." & vbCrLf & Err.Description, vbExclamation + vbOKOnly
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
ShowCode_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: ShowCode" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub tbTools_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo tbTools_ButtonMenuClick_General_ErrTrap

Select Case ButtonMenu.Key
Case "AddManual"
    Call mnuListing_Click(1)
Case "AddAuto"
    Call mnuListing_Click(2)
End Select

Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
tbTools_ButtonMenuClick_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: tbTools_ButtonMenuClick" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub
