VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   HelpContextID   =   3
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Other Options"
      Height          =   360
      Index           =   1
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   795
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Color Options"
      Height          =   360
      Index           =   0
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   795
      Value           =   -1  'True
      Width           =   2685
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   30
      Top             =   6150
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Cancel && Exit"
      Height          =   465
      Index           =   1
      Left            =   90
      TabIndex        =   18
      Top             =   5670
      Width           =   2385
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save && Exit"
      Height          =   465
      Index           =   0
      Left            =   3060
      TabIndex        =   19
      Top             =   5670
      Width           =   2385
   End
   Begin VB.CommandButton cmdNewDB 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4845
      TabIndex        =   0
      Top             =   60
      Width           =   405
   End
   Begin VB.TextBox txtDBname 
      Height          =   285
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "No database has been provided"
      Top             =   345
      Width           =   5205
   End
   Begin VB.Frame Frame1 
      Height          =   4275
      Index           =   0
      Left            =   60
      TabIndex        =   31
      Top             =   1095
      Width           =   5400
      Begin VB.PictureBox picSample 
         Height          =   975
         Index           =   0
         Left            =   90
         ScaleHeight     =   915
         ScaleWidth      =   2415
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   465
         Width           =   2475
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Main Window Main Window Main Window Main Window "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   885
            Index           =   0
            Left            =   555
            TabIndex        =   35
            Top             =   30
            Width           =   1305
         End
         Begin VB.Shape picMainsample 
            BorderWidth     =   2
            Height          =   915
            Index           =   0
            Left            =   0
            Tag             =   "2010"
            Top             =   0
            Width           =   405
         End
         Begin VB.Shape picMainsample 
            BackColor       =   &H00FFFF00&
            BackStyle       =   1  'Opaque
            BorderWidth     =   2
            Height          =   750
            Index           =   1
            Left            =   60
            Tag             =   "2010"
            Top             =   70
            Width           =   285
         End
      End
      Begin VB.PictureBox picSample 
         Height          =   975
         Index           =   1
         Left            =   2835
         ScaleHeight     =   915
         ScaleWidth      =   2415
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   465
         Width           =   2475
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pop-Ups Pop-Ups Pop-Ups Pop-Ups "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   870
            Index           =   1
            Left            =   720
            TabIndex        =   33
            Top             =   30
            Width           =   825
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "None"
         Height          =   375
         Index           =   3
         Left            =   4050
         TabIndex        =   10
         Top             =   1455
         Width           =   1275
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "Change"
         Height          =   375
         Index           =   2
         Left            =   2835
         TabIndex        =   9
         Top             =   1455
         Width           =   1215
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "None"
         Height          =   375
         Index           =   1
         Left            =   1335
         TabIndex        =   5
         Top             =   1455
         Width           =   1245
      End
      Begin VB.CommandButton cmdColors 
         Caption         =   "Change"
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   1455
         Width           =   1245
      End
      Begin VB.CheckBox chkNoGradient 
         Height          =   190
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   1845
         Width           =   210
      End
      Begin VB.CheckBox chkNoGradient 
         Alignment       =   1  'Right Justify
         Height          =   190
         Index           =   1
         Left            =   2865
         TabIndex        =   11
         Top             =   1845
         Width           =   210
      End
      Begin VB.CommandButton cmdLblColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2235
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2115
         Width           =   345
      End
      Begin VB.CommandButton cmdLblColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2115
         Width           =   345
      End
      Begin VB.ComboBox cboAlignment 
         Height          =   315
         ItemData        =   "frmOptions.frx":0000
         Left            =   90
         List            =   "frmOptions.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2430
         Width           =   2475
      End
      Begin VB.CommandButton lblCodeColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2865
         Width           =   345
      End
      Begin VB.CommandButton lblCodeColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3135
         Width           =   345
      End
      Begin VB.CommandButton lblCodeColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3405
         Width           =   345
      End
      Begin VB.CommandButton lblCodeColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3675
         Width           =   345
      End
      Begin VB.CheckBox chkColorImports 
         Height          =   190
         Left            =   135
         TabIndex        =   17
         Top             =   3975
         Width           =   210
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Color Options - Changes reflected in sample windows below"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   105
         TabIndex        =   46
         Top             =   225
         Width           =   5175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click to change label color"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   45
         Top             =   2145
         Width           =   2115
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click to change label color"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   2835
         TabIndex        =   44
         Top             =   2145
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Use solid colors - no gradients"
         Height          =   285
         Index           =   7
         Left            =   345
         TabIndex        =   43
         Top             =   1845
         Width           =   2475
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Use solid colors - no gradients"
         Height          =   285
         Index           =   8
         Left            =   3135
         TabIndex        =   42
         Top             =   1845
         Width           =   2475
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   15
         X2              =   5385
         Y1              =   2805
         Y2              =   2805
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Code lising aligns in main window"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   13
         Left            =   2625
         TabIndex        =   41
         Top             =   2475
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Click to change color for VB Core keywords"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   14
         Left            =   420
         TabIndex        =   40
         Top             =   2895
         Width           =   4905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Click to change color for VB Functions"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   15
         Left            =   420
         TabIndex        =   39
         Top             =   3165
         Width           =   4905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Click to change color for VB Miscellaneous keywords"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   16
         Left            =   420
         TabIndex        =   38
         Top             =   3405
         Width           =   4905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Click to change color for object properties && methods"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   17
         Left            =   420
         TabIndex        =   37
         Top             =   3675
         Width           =   4905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "When importing from other files, color automatically"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   18
         Left            =   420
         TabIndex        =   36
         Top             =   3945
         Width           =   5085
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   4275
      Index           =   1
      Left            =   60
      TabIndex        =   47
      Top             =   1095
      Visible         =   0   'False
      Width           =   5400
      Begin VB.CommandButton btnSample 
         Height          =   525
         Left            =   4605
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1575
         Width           =   645
      End
      Begin VB.ComboBox cboLanguage 
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2475
         Width           =   3075
      End
      Begin VB.OptionButton optBtnSize 
         Caption         =   "Small Buttons"
         Height          =   435
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1665
         Width           =   1455
      End
      Begin VB.OptionButton optBtnSize 
         Caption         =   "Large Buttons"
         Height          =   435
         Index           =   2
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1665
         Width           =   1455
      End
      Begin VB.OptionButton optBtnSize 
         Caption         =   "No Menubar"
         Height          =   435
         Index           =   0
         Left            =   3060
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1665
         Width           =   1455
      End
      Begin VB.ComboBox cboAddType 
         Height          =   315
         ItemData        =   "frmOptions.frx":002E
         Left            =   2790
         List            =   "frmOptions.frx":003B
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3585
         Width           =   2505
      End
      Begin VB.TextBox txtFont 
         Height          =   285
         Index           =   0
         Left            =   495
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Font"
         Top             =   570
         Width           =   2625
      End
      Begin VB.TextBox txtFont 
         Height          =   285
         Index           =   1
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Size"
         Top             =   570
         Width           =   1245
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   20
         Top             =   585
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Graphical Menubar Button Size"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   52
         Top             =   1245
         Width           =   5175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Default Language"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   -15
         TabIndex        =   51
         Top             =   2505
         Width           =   2205
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   -15
         X2              =   5355
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   -45
         X2              =   5325
         Y1              =   3150
         Y2              =   3150
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "When the New Code icon            is clicked which default add option do you want to use?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Index           =   19
         Left            =   450
         TabIndex        =   50
         Top             =   3420
         Width           =   2355
      End
      Begin VB.Image imgAddCode 
         Height          =   330
         Left            =   180
         Top             =   3630
         Width           =   300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Font Size"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   4440
         TabIndex        =   49
         Top             =   585
         Width           =   900
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   0
         X2              =   5370
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Font Options (only affects new code)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   11
         Left            =   -60
         TabIndex        =   48
         Top             =   240
         Width           =   5715
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Note: The color of this window is NOT dependent on these settings."
      Height          =   255
      Index           =   6
      Left            =   90
      TabIndex        =   53
      Top             =   5415
      Width           =   5325
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   60
      X2              =   5430
      Y1              =   5370
      Y2              =   5370
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Database Location.  Click button to change >>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   105
      TabIndex        =   29
      Top             =   105
      Width           =   4515
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private CurWinColor As String, CurPopUpColor As String
Private CurWinSolids As Integer, CurPopUpSolids As Integer
'================================================================
' This form sets the user options stored in the registry, including which master database to use
'================================================================

Private Sub btnSample_GotFocus()
sBar.SimpleText = "Sample button size for task bar."
End Sub

Private Sub cboAddType_GotFocus()
sBar.SimpleText = "Default when you choose to add new files from the toolbar."
End Sub

Private Sub cboAlignment_Click()
' Inserted by LaVolpe
On Error GoTo Sub_cboAlignment_Click_General_ErrTrap_by_LaVolpe
picMainsample(0).Left = cboAlignment.ItemData(cboAlignment.ListIndex)
picMainsample(1).Left = picMainsample(0).Left + 60
Exit Sub

Sub_cboAlignment_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cboAlignment_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cboAlignment_GotFocus()
sBar.SimpleText = "Change code listing to right or left side of Main window."
End Sub

Private Sub cboLanguage_GotFocus()
sBar.SimpleText = "Language automatically selected for new code entries"
End Sub

Private Sub chkColorImports_GotFocus()
sBar.SimpleText = "Check box: When importing text from other files, color as you view?"
End Sub

Private Sub chkNoGradient_Click(Index As Integer)
' Inserted by LaVolpe
On Error GoTo Sub_chkNoGradient_Click_General_ErrTrap_by_LaVolpe
If Index = 0 Then MyDefaults.GradientMain = chkNoGradient(Index) Else MyDefaults.GradientPopup = chkNoGradient(Index)
DoGradient picSample(Index), Index + 1                  ' change the color to a gradient
Exit Sub

Sub_chkNoGradient_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub chkNoGradient_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub chkNoGradient_GotFocus(Index As Integer)
If Index = 0 Then
    sBar.SimpleText = "Check box: Use solid colors only for the Main window"
Else
    sBar.SimpleText = "Check box: Use solid colors only for the Code windows"
End If
End Sub

Private Sub cmdColors_Click(Index As Integer)
'================================================================
'   Sub changes the sample colors and stores them if settings are saved on exiting
'================================================================

Dim PicID As Integer
' Inserted by LaVolpe
On Error GoTo Sub_cmdColors_Click_General_ErrTrap_by_LaVolpe
Select Case Index
Case 1, 3   ' no color option
    If Index = 1 Then PicID = 0 Else PicID = 1              ' identify which picture box sample is being changed
    If PicID = 0 Then                                                       ' depending on which picture box is being changed
        MyDefaults.WindowColor = WinBlahColor           '       set the main color or popup color to gray
    Else
        MyDefaults.PopupColor = WinBlahColor
    End If
    picSample(PicID).BackColor = 0
    picSample(PicID).BackColor = -2147483633
Case 0, 2   ' choose a color
    PicID = Index / 2                                                       ' identify which picture box is being modified
    With frmLibrary.dlgCommon                                       ' display the color dialog box
        On Error GoTo UsrCnx
        .ShowColor
        .Flags = 0
        If Val(PicID) = 0 Then                                              ' depnding on which picture box is being modified
                MyDefaults.WindowColor = CStr(.Color)       ' change the color to user's setting
        Else
                MyDefaults.PopupColor = CStr(.Color)
        End If
    End With
    DoGradient picSample(PicID), PicID + 1                  ' change the color to a gradient
End Select

UsrCnx:
Exit Sub

Sub_cmdColors_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdColors_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdColors_GotFocus(Index As Integer)
Select Case Index
Case 0: sBar.SimpleText = "Change background color of the Main window."
Case 2: sBar.SimpleText = "Change background color of the Code window."
Case 1, 3: sBar.SimpleText = "No special color. Use Window's default color."
End Select
End Sub

Private Sub cmdLblColor_Click(Index As Integer)
On Error GoTo UserCnx
With frmLibrary.dlgCommon
    .Flags = 0
    .ShowColor
    Label1(Index).ForeColor = .Color
    cmdLblColor(Index).BackColor = .Color
UserCnx:
End With
End Sub

Private Sub cmdLblColor_GotFocus(Index As Integer)
sBar.SimpleText = "Change the color of the labels that appear in the window."
End Sub

Private Sub cmdNewDB_Click()
'================================================================
'  Offers user to connect to another database
'================================================================
' Provide this warning first. If the database is being modified at same time this program is accessing the tables,
'   database corruption can occur or skewed record synchronization can cause undesireable affects
' Inserted by LaVolpe
On Error GoTo Sub_cmdNewDB_Click_General_ErrTrap_by_LaVolpe
'If MsgBox("WARNING: If you have code in any open windows, close them before continuing or you may lose " & _
    "the code or corrupt your database. " & vbCrLf & "Continue?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then Exit Sub
GP = Null
frmDBprompt.Show 1, Me                  ' Show the db connection window
If GP = "LoadDatabase" Then             ' if a different db was chosen, then...
    Tag = "Options"
    CloseAllWindows False, Tag
    On Error Resume Next
    mainDB.Close                                ' close the database and call the Main Sub() to reload the entire program
    Set mainDB = Nothing
    Call Main
    txtDBname = mainDB.Name         ' update the current database display text box
Else
    If mainDB Is Nothing Then End
End If
Exit Sub

Sub_cmdNewDB_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdNewDB_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdNewDB_GotFocus()
sBar.SimpleText = "Click to select a different database to store code"
End Sub

Private Sub cmdOK_Click(Index As Integer)
'================================================================
'    Exit function & save settings option
'================================================================

' Inserted by LaVolpe
On Error GoTo Sub_cmdOK_Click_General_ErrTrap_by_LaVolpe
GP = Null
If Index = 0 Then ' Save all changes
    Dim bDirty As Boolean
    ' see if any changes exist between older label color settings and the ones in this form
    If cmdLblColor(0).BackColor <> MyDefaults.LblColorMain Or cmdLblColor(1) <> MyDefaults.LblColorPopup Then bDirty = True
        MyDefaults.LblColorMain = cmdLblColor(0).BackColor
        MyDefaults.LblColorPopup = cmdLblColor(1).BackColor
        SaveSetting "LaVolpeCodeSafe", "Settings", "LblColorMain", CStr(MyDefaults.LblColorMain)
        SaveSetting "LaVolpeCodeSafe", "Settings", "LblColorPopUp", CStr(MyDefaults.LblColorPopup)
    ' see if any changes exist between older color settings and the ones in this form
    If CurWinColor <> MyDefaults.WindowColor Or CurPopUpColor <> MyDefaults.PopupColor Then bDirty = True
        CurWinColor = MyDefaults.WindowColor        ' Save current settings into these variables, used in Unload
        CurPopUpColor = MyDefaults.PopupColor
        SaveSetting "LaVolpeCodeSafe", "Settings", "MainColor", CurWinColor   ' save settings to registry
        SaveSetting "LaVolpeCodeSafe", "Settings", "SecondaryColor", CurPopUpColor
        MyDefaults.WindowColor = CurWinColor                                                ' save settings in global variable
        MyDefaults.PopupColor = CurPopUpColor
    ' see if any changes exist between older gradient/solid settings and the ones in this form
    If CurWinSolids <> MyDefaults.GradientMain Or CurPopUpSolids <> MyDefaults.GradientPopup Then bDirty = True
        CurWinSolids = MyDefaults.GradientMain        ' Save current settings into these variables, used in Unload
        CurPopUpSolids = MyDefaults.GradientPopup
        SaveSetting "LaVolpeCodeSafe", "Settings", "GradientMain", CurWinSolids   ' save settings to registry
        SaveSetting "LaVolpeCodeSafe", "Settings", "GradientPopUp", CurPopUpSolids
        MyDefaults.GradientMain = CurWinSolids                                                ' save settings in global variable
        MyDefaults.GradientPopup = CurPopUpSolids
    ' see if any changes exist between older button size and the setting in this form
    If CInt(Val(btnSample.Tag)) <> MyDefaults.ButtonSize Then bDirty = True
    SaveSetting "LaVolpeCodeSafe", "Settings", "ButtonSize", btnSample.Tag      ' save settings in registry
    MyDefaults.ButtonSize = CInt(Val(btnSample.Tag))                                       ' save settings in global variable
    ' see if any changes exist between older language default and the setting in this form
    If cboLanguage.ItemData(cboLanguage.ListIndex) <> MyDefaults.Language Then bDirty = True
    MyDefaults.Language = cboLanguage.ItemData(cboLanguage.ListIndex)       ' save settings in global variable
    SaveSetting "LaVolpeCodeSafe", "Settings", "Language", CStr(MyDefaults.Language) ' save setting in registry
    ' see if any changes exist between older font name/size and the setting in this form
    If txtFont(0) <> MyDefaults.Font Or txtFont(1) <> MyDefaults.FontSize Then bDirty = True
    MyDefaults.Font = txtFont(0): MyDefaults.FontSize = txtFont(1)    ' save settings in global variable
    SaveSetting "LaVolpeCodeSafe", "Settings", "FontSize", CStr(MyDefaults.FontSize) ' save setting in registry
    SaveSetting "LaVolpeCodeSafe", "Settings", "FontType", MyDefaults.Font ' save setting in registry
    If bDirty = True Then GP = "Reload"     ' identify flag for main window to reload these settings
    ' see if any changes exist between older code list alignment and the setting in this form
    If cboAlignment.ListIndex + 3 <> MyDefaults.Align Then bDirty = True
    MyDefaults.Align = cboAlignment.ListIndex + 3     ' save settings in global variable
    SaveSetting "LaVolpeCodeSafe", "Settings", "Alignment", CStr(MyDefaults.Align) ' save setting in registry
    ' for code coloring, don't force a reload, simply notify that changes will only occur hereafter
    If MyDefaults.KeyWd1 <> lblCodeColor(0).BackColor Or MyDefaults.KeyWd2 <> lblCodeColor(1).BackColor Or _
        MyDefaults.KeyWd3 <> lblCodeColor(2).BackColor Or MyDefaults.KeyWd4 <> lblCodeColor(3).BackColor Then
        MyDefaults.KeyWd1 = lblCodeColor(0).BackColor
        MyDefaults.KeyWd2 = lblCodeColor(1).BackColor
        MyDefaults.KeyWd3 = lblCodeColor(2).BackColor
        MyDefaults.KeyWd4 = lblCodeColor(3).BackColor
        SaveSetting "LaVolpeCodeSafe", "Settings", "CompilerColor", CStr(lblCodeColor(0).BackColor)
        SaveSetting "LaVolpeCodeSafe", "Settings", "FunctionsColor", CStr(lblCodeColor(1).BackColor)
        SaveSetting "LaVolpeCodeSafe", "Settings", "MiscColor", CStr(lblCodeColor(2).BackColor)
        SaveSetting "LaVolpeCodeSafe", "Settings", "PropertiesColor", CStr(lblCodeColor(3).BackColor)
            MsgBox "When code colors are changed, those colors will be updated as new windows open.", vbInformation + vbOKOnly
    End If
    If cboAddType.ListIndex <> Val(cboAddType.Tag) Then bDirty = True
    SaveSetting "LaVolpeCodeSafe", "Settings", "AddType", CStr(cboAddType.ListIndex)
    If chkColorImports <> Val(chkColorImports.Tag) Then
        bDirty = True
        MsgBox "The option to automatically color text imported from other files will take affect next time that window is opened.", vbInformation + vbOKOnly
    End If
    SaveSetting "LaVolpeCodeSafe", "Settings", "ColorImports", CStr(chkColorImports)
End If
Unload Me
Exit Sub

Sub_cmdOK_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdOK_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdOK_GotFocus(Index As Integer)
If Index = 0 Then
    sBar.SimpleText = "Save changes and close this window"
Else
    sBar.SimpleText = "Ignore all changes. Except a database change, if made."
End If
End Sub

Private Sub Command1_Click()

' Inserted by LaVolpe
On Error GoTo Sub_Command1_Click_General_ErrTrap_by_LaVolpe
With frmLibrary.dlgCommon                                       ' display the color dialog box
    On Error GoTo UsrCnx
    .FontName = txtFont(0)
    .FontSize = txtFont(1)
    .Flags = cdlCFBoth
End With
    frmLibrary.dlgCommon.ShowFont
    With frmLibrary.dlgCommon
        txtFont(0) = .FontName
        txtFont(1) = .FontSize
        Label1(11).Font = txtFont(0)
        Label1(11).FontSize = txtFont(1)
    End With
UsrCnx:
Exit Sub

Sub_Command1_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Command1_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Command1_GotFocus()
sBar.SimpleText = "Change default font & size of text in Code windows"
End Sub

Private Sub Form_Load()
'================================================================
'   Syncrhonize form controls with user settings
'================================================================

' Inserted by LaVolpe
On Error GoTo Sub_Form_Load_General_ErrTrap_by_LaVolpe
    DoGradient Me, 1, True                      ' paint this form
    Frame1(0).BackColor = BackColor
    Frame1(1).BackColor = BackColor
    If Not mainDB Is Nothing Then           ' Existing database, so display the name
        txtDBname.Tag = mainDB.Name
        txtDBname.Text = mainDB.Name
    End If
    Dim I As Integer
    With frmLibrary
        Icon = .SmallImages.ListImages(2).ExtractIcon   ' display icon in titlebar
        optBtnSize(MyDefaults.ButtonSize) = True         ' trigger the button size option, per user settings
        cboLanguage.Clear                                           ' load the languages from the main window
        For I = 1 To .cboFilter(2).ListCount - 1
            cboLanguage.AddItem .cboFilter(2).List(I)
            cboLanguage.ItemData(I - 1) = .cboFilter(2).ItemData(I)     ' track the language db record ID
            ' if the default language is found here, then select it
            If cboLanguage.ItemData(I - 1) = MyDefaults.Language Then cboLanguage.ListIndex = I - 1
        Next
    End With
    CurWinColor = MyDefaults.WindowColor            ' set color variables
    CurPopUpColor = MyDefaults.PopupColor           ' the gray color is "-12345"
    CurWinSolids = MyDefaults.GradientMain
    chkNoGradient(0) = MyDefaults.GradientMain
    CurPopUpSolids = MyDefaults.GradientPopup
    chkNoGradient(1) = MyDefaults.GradientPopup
    cmdLblColor(0).BackColor = MyDefaults.LblColorMain
    cmdLblColor(1).BackColor = MyDefaults.LblColorPopup
    For I = Label1.LBound To Label1.UBound
        If MyDefaults.WindowColor = WinBlahColor Then
            Label1(I).ForeColor = &HFFFFFF
        Else
            Label1(I).ForeColor = MyDefaults.LblColorMain
        End If
    Next
    Label1(0).ForeColor = MyDefaults.LblColorMain
    Label1(1).ForeColor = MyDefaults.LblColorPopup
    Call chkNoGradient_Click(0)
    Call chkNoGradient_Click(1)
    txtFont(0) = MyDefaults.Font: txtFont(1) = MyDefaults.FontSize
    Label1(11).Font = txtFont(0): Label1(11).FontSize = txtFont(1)
    cboAlignment.ListIndex = MyDefaults.Align - 3
    lblCodeColor(0).BackColor = MyDefaults.KeyWd1
    lblCodeColor(1).BackColor = MyDefaults.KeyWd2
    lblCodeColor(2).BackColor = MyDefaults.KeyWd3
    lblCodeColor(3).BackColor = MyDefaults.KeyWd4
    cboAddType.ListIndex = Val(GetSetting("LaVolpeCodeSafe", "Settings", "AddType", "0"))
    cboAddType.Tag = cboAddType.ListIndex
    chkColorImports = Val(GetSetting("LaVolpeCodeSafe", "Settings", "ColorImports", "0"))
    chkColorImports.Tag = chkColorImports
    imgAddCode.Picture = frmLibrary.SmallImages.ListImages(4).ExtractIcon
Exit Sub

Sub_Form_Load_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Load]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Terminate()
'================================================================
' Inserted by LaVolpe
On Error GoTo Sub_Form_Terminate_General_ErrTrap_by_LaVolpe
Unload Me
'================================================================
Exit Sub

Sub_Form_Terminate_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Terminate]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
'================================================================
'   Resets colors if user click the X button in titlebar
'================================================================
' Inserted by LaVolpe
On Error GoTo Sub_Form_Unload_General_ErrTrap_by_LaVolpe
MyDefaults.WindowColor = CurWinColor
MyDefaults.PopupColor = CurPopUpColor
MyDefaults.GradientMain = CurWinSolids
MyDefaults.GradientPopup = CurPopUpSolids
Exit Sub

Sub_Form_Unload_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Unload]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
Case 2
    Call cmdLblColor_Click(0)
Case 5
    cmdLblColor_Click (5)
Case 14, 15, 16, 17
    Call lblCodeColor_Click(Index - 14)
Case 18
    chkColorImports = Abs(chkColorImports - 1)
End Select
End Sub

Private Sub lblCodeColor_Click(Index As Integer)
On Error GoTo UserCnx
With frmLibrary.dlgCommon
    .Flags = 0
    .ShowColor
    lblCodeColor(Index).BackColor = .Color
End With
UserCnx:
End Sub

Private Sub lblCodeColor_GotFocus(Index As Integer)
sBar.SimpleText = "Colors used when coloring code."
End Sub

Private Sub optBtnSize_Click(Index As Integer)
'================================================================
'   Changes the sample button display per user's request
'================================================================

' Inserted by LaVolpe
On Error GoTo Sub_optBtnSize_Click_General_ErrTrap_by_LaVolpe
If Index = 1 Then       ' Small buttons
    btnSample.Picture = frmLibrary.SmallImages.ListImages(1).ExtractIcon
Else
    If Index = 2 Then   ' Large buttons
        btnSample.Picture = frmLibrary.LargeImages.ListImages(1).ExtractIcon
    End If
End If
btnSample.Visible = Index       ' visible unless the "no menubar" option (value 0) is selected
btnSample.Tag = Index           ' set the tag to store the selection
Exit Sub

Sub_optBtnSize_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub optBtnSize_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub optBtnSize_GotFocus(Index As Integer)
Select Case Index
Case 1: sBar.SimpleText = "Use small icons on the taskbar"
Case 2: sBar.SimpleText = "Use large icons on the taskbar"
Case 0: sBar.SimpleText = "Do not show a task bar."
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
Frame1(Index).Enabled = True
Frame1(Index).Visible = True
Frame1(Abs(Index - 1)).Visible = False
Frame1(Abs(Index - 1)).Enabled = False
End Sub

Private Sub picSample_GotFocus(Index As Integer)
If Index = 0 Then
    sBar.SimpleText = "Sample main window with colors and code listing position"
Else
    sBar.SimpleText = "Sample code window with colors"
End If
End Sub

Private Sub txtDBname_GotFocus()
sBar.SimpleText = "Current database. Click elipse button (...) to change."
End Sub

Private Sub txtFont_GotFocus(Index As Integer)
If Index = 0 Then
    sBar.SimpleText = "Current font. Label above that is a sample."
Else
    sBar.SimpleText = "Currnt font size. Label above that is a sample."
End If
End Sub
