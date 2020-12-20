VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDBprompt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Name and Location"
   ClientHeight    =   1125
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   HelpContextID   =   10
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgDB 
      Left            =   5430
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.mdb"
      DialogTitle     =   "Location of code library database"
      Filter          =   "MS Access databases|*.mdb"
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      TabIndex        =   0
      ToolTipText     =   "Select database"
      Top             =   615
      Width           =   375
   End
   Begin VB.TextBox txtDBname 
      Height          =   345
      Left            =   570
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   615
      Width           =   3975
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please provide the name and location of the database containing your code"
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
      Height          =   405
      Left            =   570
      TabIndex        =   3
      Top             =   165
      Width           =   4035
   End
End
Attribute VB_Name = "frmDBprompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'============================================================
' This form is used to identify the location of the master database
'   it is called from the Options form and also during intial startup
'============================================================
Private sCurrentDB As String, bReset As Boolean

Private Sub CancelButton_Click()
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_CancelButton_Click_General_ErrTrap_by_LaVolpe
Unload Me
'============================================================
Exit Sub

Sub_CancelButton_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub CancelButton_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdOpen_Click()
'============================================================
'   Allows user to select a master database
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_cmdOpen_Click_General_ErrTrap_by_LaVolpe
dlgDB.Flags = cdlOFNFileMustExist
On Error GoTo UserCnx
dlgDB.ShowOpen                  ' show the Open Dialog box
' if no change in the database name/location, then exit sub
If txtDBname.Tag = dlgDB.FileName Then Exit Sub
' otherwise, set the tag value to the new filename & try to connect
txtDBname.Tag = dlgDB.FileName
Label1.ForeColor = MyDefaults.LblColorPopup
Connect2DB True     ' connect to the database & show results

UserCnx:
Exit Sub

Sub_cmdOpen_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdOpen_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Load()
'============================================================
'   Retrieves current database location & determines if form is to be displayed or not
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_Form_Load_General_ErrTrap_by_LaVolpe
bReset = False
sCurrentDB = ""
If Not IsNull(GP) Then
    sCurrentDB = GP
    txtDBname.Tag = GP
    bReset = True
Else
    If Not mainDB Is Nothing Then           ' User called this from the Options form
        txtDBname = mainDB.Name                 ' set the display text box & tag properties
        txtDBname.Tag = mainDB.Name
        sCurrentDB = txtDBname
    Else                                                    ' probably first use or registry settings deleted
        ' get the location from the registry settings
        txtDBname.Tag = GetSetting("LaVolpeCodeSafe", "Settings", "Path", "No database selected")
        ' if no setting existed, then display the form, otherwise set the txtDBname Tag to the stored setting
        If txtDBname.Tag = "No database selected" Then GoTo DisplayMe Else MyDefaults.DBname = txtDBname.Tag
    End If
End If
Dim bExists As Boolean, bNoDisplay As Boolean
On Error Resume Next
bNoDisplay = (IsNull(GP) = False)
bExists = CBool(Len(Dir(txtDBname.Tag)))    ' ensure the file exists
If bExists = True And Err.Number = 0 And mainDB Is Nothing Then  ' if exists and and the Main Sub() called this routine, then
    If Connect2DB(False, bNoDisplay) = True Then Exit Sub       '   try to connect & don't show results if successful
Else                                ' database selected doesn't exist
    Err.Clear
End If
' If database doesn't exist, or called from the Options form, then display this form
DisplayMe:
If bNoDisplay = True Then
    Unload Me
    GP = Null
    Exit Sub
End If
If DoGradient(Me, 1, True) = False Then Label1.ForeColor = 0
Screen.MousePointer = vbDefault
Show
Exit Sub

Sub_Form_Load_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Load]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Function Connect2DB(bShowResult As Boolean, Optional bNoRegUpdate As Boolean = False) As Boolean
'============================================================
'   Function attempts to connect to a database, unloads this form (if needed), and returns success/failure
'============================================================

On Error Resume Next
mainDB.Close                                            ' close the db, it will be connected to later
Set mainDB = Nothing                                    ' provide success message box
GP = Null
bReset = True
Dim ErrCtr As Integer
On Error GoTo FailedConnect
Set mainDB = OpenDatabase(txtDBname.Tag)    ' try opening the database
txtDBname = txtDBname.Tag                               ' success - update the display textbox
Screen.MousePointer = vbDefault
Connect2DB = True
If bShowResult = True And bNoRegUpdate = False Then     ' if results are to be shown, then how them
    MsgBox "Successfully connected to the database", vbInformation + vbOKOnly
End If
    ' update registry setting and global variable
If MyDefaults.DBname <> txtDBname.Tag Then       ' no change
    If bNoRegUpdate = False Then
        SaveSetting "LaVolpeCodeSafe", "Settings", "Path", txtDBname.Tag
        MyDefaults.DBname = txtDBname.Tag
        
    End If
    GP = "LoadDatabase"                                       ' update GP variable to indicate form is closing
Else
    GP = "LoadDatabase-No"                                      ' update GP variable
End If
Unload Me                                                       ' after successful connection close this form
Exit Function

FailedConnect:
If Err.Number = 3051 And ErrCtr = 0 Then
    SetAttr txtDBname.Tag, vbNormal
    Err.Number = 0
    ErrCtr = 1
    Resume
Else
' on all errors, leave form open & prompt user with error description unless this is a true first run
    If bNoRegUpdate = False Then MsgBox "Try again after addressing the following error:" & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly
    txtDBname.Tag = ""
    GP = Null
End If
End Function

Private Sub Form_Terminate()
'============================================================
Unload Me
'============================================================
End Sub

Private Sub Form_Unload(Cancel As Integer)
'============================================================
'   Form unloads but first warns user that it can't continue without a valid database
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_Form_Unload_General_ErrTrap_by_LaVolpe
If mainDB Is Nothing Then
    If Len(sCurrentDB) Then
        If bReset = True Then
            On Error Resume Next
            mainDB.Close
            Set mainDB = Nothing
            GP = "LoadDatabase"
        End If
    Else
        If MsgBox("Canceling without selecting a database will terminate this application.  Terminate?", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
            Cancel = 1
        Else
            On Error Resume Next
            mainDB.Close
            Set mainDB = Nothing
            GP = "Failure"
        End If
        ' the calling functions, (Options Form & Main Sub()) will terminate the program if needed
    End If
End If
Exit Sub

Sub_Form_Unload_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Unload]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub
