VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLoadExisting 
   Caption         =   "Procedure View"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   HelpContextID   =   11
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Tag             =   "903011100"
   Begin VB.ListBox lstIndex 
      Enabled         =   0   'False
      Height          =   3375
      Left            =   585
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   1425
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   9480
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6350
            Text            =   "Ready"
            TextSave        =   "Ready"
            Object.Tag             =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6350
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   767
            MinWidth        =   220
            Text            =   "Line "
            TextSave        =   "Line "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1349
            MinWidth        =   1358
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   220
            Text            =   "Chars "
            TextSave        =   "Chars "
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1349
            MinWidth        =   1358
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstProcedures 
      BackColor       =   &H00FFFFC0&
      Height          =   4350
      Left            =   105
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   "4350"
      Top             =   4440
      Width           =   2625
   End
   Begin VB.FileListBox lstFiles 
      Height          =   3210
      Left            =   75
      TabIndex        =   2
      Top             =   960
      Width           =   2625
   End
   Begin VB.ComboBox cboActions 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      ItemData        =   "frmViewExisting.frx":0000
      Left            =   870
      List            =   "frmViewExisting.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   15
      Width           =   1785
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse for Project Folder"
      Height          =   345
      Left            =   75
      TabIndex        =   1
      Top             =   360
      Width           =   2625
   End
   Begin MSComDlg.CommonDialog dlgColors 
      Left            =   2205
      Top             =   8310
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin RichTextLib.RichTextBox txtProcedures 
      Height          =   8775
      Left            =   2745
      TabIndex        =   10
      Top             =   15
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   15478
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"frmViewExisting.frx":002D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frameOpts 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   945
      Left            =   90
      TabIndex        =   11
      Top             =   8595
      Width           =   9885
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   330
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Search Text"
         Top             =   510
         Width           =   480
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "Color"
         Height          =   330
         Left            =   1530
         TabIndex        =   5
         ToolTipText     =   "Color VB Keywords"
         Top             =   180
         Width           =   1005
      End
      Begin VB.CheckBox chkSelOnly 
         Caption         =   "Use selected text only"
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         ToolTipText     =   "when copying to Delcarations or Code section of new record"
         Top             =   495
         Width           =   1305
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "To Declarations"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   30
         TabIndex        =   7
         ToolTipText     =   "Add to Declarations section in new record"
         Top             =   510
         Width           =   1485
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "To Code"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1530
         TabIndex        =   8
         ToolTipText     =   "Add to Code Section in New Record"
         Top             =   510
         Width           =   1005
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Default         =   -1  'True
         Height          =   330
         Left            =   30
         TabIndex        =   4
         Top             =   180
         Width           =   1485
      End
      Begin VB.Label lblPath 
         Caption         =   "Path:  "
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   4590
         TabIndex        =   17
         Tag             =   "Path:  "
         Top             =   225
         Width           =   5265
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Procedures found - Click"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   14
      ToolTipText     =   "click listing to display that procedure"
      Top             =   4185
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Items found in folder - Click"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   165
      TabIndex        =   13
      ToolTipText     =   "click listing to display procedures for that file"
      Top             =   735
      Width           =   2505
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   12
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "frmLoadExisting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Private bStop As Boolean
Private bAutoColor As Boolean

Private Sub cboActions_Click()
'===================================================================
' Basic functions for this program - right click functions offer more options
'===================================================================
' Inserted by LaVolpe
On Error GoTo Sub_cboActions_Click_General_ErrTrap_by_LaVolpe
Dim I As Integer, sFileName As String, Index As Integer
Select Case cboActions.ListIndex
Case 0:
    If lstFiles.ListIndex > -1 Then sFileName = lstFiles
    Index = lstProcedures.ListIndex
    lstFiles.Pattern = "*.*"                       ' make file listbox show all files
Case 1:
    If lstFiles.ListIndex > -1 Then sFileName = lstFiles
    Index = lstProcedures.ListIndex
    lstFiles.Pattern = "*.frm;*.bas;*.clt;*.dob;*.cls;*.pag;*.dsr"        ' show only VB files
End Select
Exit Sub

Sub_cboActions_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cboActions_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cboActions_GotFocus()
sBar.Panels(1) = "Change between viewing VB-only files and all files."
End Sub

Private Sub cboActions_LostFocus()
sBar.Panels(1) = sBar.Panels(1).Tag
End Sub

Private Sub chkSelOnly_GotFocus()
sBar.Panels(1) = "Selected text vs entire text"
End Sub

Private Sub cmdAdd_Click(Index As Integer)
Dim sName As String, I As Integer, iResponse As Integer
If Len(txtProcedures.SelText) = 0 And chkSelOnly = 1 Then
    MsgBox "You've opted to only copy the selected text, but haven't selected any text.", vbInformation + vbOKOnly
    Exit Sub
End If
For I = 0 To Forms.Count - 1
    If Forms(I).Caption = "New Record" Then Exit For
Next
If I < Forms.Count Then
    iResponse = MsgBox("You already have a new record open. Add this to the newest record? " & vbCrLf & "Click No to create another new record", vbYesNoCancel + vbQuestion)
    Select Case iResponse
    Case vbCancel:  Exit Sub
    Case vbYes: Tag = I
    Case vbNo
        frmLibrary.ShowCode "New"
        Tag = Forms.Count - 1
    End Select
Else
        frmLibrary.ShowCode "New"
        Tag = Forms.Count - 1
End If
DoEvents
MousePointer = vbHourglass
sBar.Panels(1).Text = "Loading into new record."
With Forms(Val(Tag))
    GP = Null
    If chkSelOnly = 0 Then
        frmLibrary.rtfStaging = txtProcedures
    Else
        frmLibrary.rtfStaging = txtProcedures.SelText
    End If
    If Index = 0 Then
        CheckLine4KeyWords .txtCode, frmLibrary.rtfStaging, , True, pBar
        I = InStr(lstProcedures, " ")
        If I > 0 Then sName = Mid(lstProcedures, I + 1) Else sName = lstProcedures
        .txtCodeName = Left(sName, 150)
        .txtPurpose = Left("Extracted from " & lstFiles, 255)
        .txtPurpose.SetFocus
    Else
        CheckLine4KeyWords .txtDeclarations, frmLibrary.rtfStaging, , True, pBar
        .txtDeclarations.SetFocus
    End If
End With
sBar.Panels(1).Text = "Ready"
MousePointer = vbDefault
End Sub

Private Sub cmdAdd_GotFocus(Index As Integer)
Select Case Index
Case 0:
    sBar.Panels(1) = "Copy to the code section of new repository record"
Case 1:
    sBar.Panels(1) = "Copy to declarations section of new repository record"
End Select
End Sub

Private Sub cmdAdd_LostFocus(Index As Integer)
sBar.Panels(1) = sBar.Panels(1).Tag
End Sub

Private Sub cmdBrowse_GotFocus()
sBar.Panels(1) = "Browse for VB project files"
End Sub

Private Sub cmdBrowse_LostFocus()
sBar.Panels(1) = sBar.Panels(1).Tag
End Sub

'===================================================================
'   Close or Cancel Loading or Cancel Coloring
'===================================================================
Private Sub cmdClose_Click()
' Inserted by LaVolpe
On Error GoTo Sub_cmdClose_Click_General_ErrTrap_by_LaVolpe
If cmdClose.Caption = "&Close" Then
    Unload Me
Else
    bStop = True
End If
Exit Sub

Sub_cmdClose_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdClose_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdBrowse_Click()
'===================================================================
'   Browse
'===================================================================

On Error GoTo UserCnx
With dlgColors
    .Filter = "All Files|*.*|VB Files|*.frm;*.bas;*.dob;*.cls;*.pag;*.dsr"
    .FileName = ""
    .CancelError = True
    .Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
End With
dlgColors.ShowOpen
lstFiles.Path = StripFile(dlgColors.FileName, "P")
lstFiles.Refresh
If cboActions.ListIndex = 0 Then            ' display all files if the was the last option
    lstFiles.Pattern = "*.*"
Else
    cboActions.ListIndex = 1                    ' otherwise only display V B files
End If
lstProcedures.Clear
lstIndex.Clear
txtProcedures.Tag = ""
Exit Sub

UserCnx:
End Sub

Private Sub cmdClose_GotFocus()
If cmdClose.Caption = "&Close" Then
    sBar.Panels(1) = "Close Window"
Else
    sBar.Panels(1) = "Stop reading file"
End If
End Sub

Private Sub cmdClose_LostFocus()
sBar.Panels(1) = sBar.Panels(1).Tag
End Sub

Private Sub cmdColor_Click()
If lstProcedures.ListCount = 0 Or lstProcedures.ListIndex = -1 Then Exit Sub
If cmdColor.Caption = "Color" Then
    ReadProcedure True
Else
    GP = "Stop"
    cmdColor.Caption = "Color"
End If
End Sub

Private Sub cmdColor_GotFocus()
If cmdColor.Caption = "Color" Then
    sBar.Panels(1) = "Color text"
Else
    sBar.Panels(1) = "Abort the coloring process"
End If
End Sub

Private Sub cmdSearch_Click()
Dim sCriteria As String, sMsg As String, lStartSearch As Long, iWholeWord As Integer
sMsg = "Enter the search string below." & vbCrLf & "For whole-word matches start search string with an exclamation mark (!)"
sCriteria = InputBox(sMsg, "Search Criteria", cmdSearch.Tag)
If Trim(sCriteria) = "" Or Trim(sCriteria) = "!" Then Exit Sub
cmdSearch.Tag = sCriteria
With txtProcedures
    If cmdSearch.Tag = "" Then lStartSearch = 0 Else lStartSearch = .SelStart + 1
    If Left(sCriteria, 1) = "!" Then
        iWholeWord = rtfWholeWord
    Else
        iWholeWord = 0
        sCriteria = " " & sCriteria
    End If
    If .Find(Mid(sCriteria, 2), lStartSearch, , iWholeWord) = -1 Then
        If .Find(Mid(sCriteria, 2), 0, , iWholeWord) = -1 Then
            MsgBox "Text not Found", vbInformation + vbOKOnly
        End If
    End If
End With
End Sub

Private Sub cmdSearch_GotFocus()
sBar.Panels(1) = "Search for text within the procedure"
End Sub

Private Sub Form_Load()
'===================================================================
'   Load form & initialize startup variables
'===================================================================

Dim I As Integer
' Inserted by LaVolpe
On Error GoTo Sub_Form_Load_General_ErrTrap_by_LaVolpe
txtProcedures.RightMargin = 32000
cboActions.ListIndex = 1
Call lstFiles_PathChange                              ' reflect current path
Left = 0
Top = 0
Height = Val(Left(Tag, 4))
Width = Val(Mid(Tag, 5))
Tag = ""
cmdSearch.Picture = frmLibrary.SmallImages.ListImages(6).ExtractIcon
Icon = frmLibrary.SmallImages.ListImages(4).ExtractIcon
bAutoColor = CBool(GetSetting("LaVolpeCodeSafe", "Settings", "ColorImports", "0"))
Exit Sub

Sub_Form_Load_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Load]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Resize()
'===================================================================
'   Resize form objects when windows size changes
'===================================================================

' Inserted by LaVolpe
On Error GoTo Sub_Form_Resize_General_ErrTrap_by_LaVolpe
If WindowState = 1 Then Exit Sub
On Error Resume Next
txtProcedures.Width = Width - 2895
txtProcedures.Height = Height - 1065 - 465
lstFiles.Height = txtProcedures.Height * 0.3658
lstProcedures.Top = lstFiles.Height + 1225
lstProcedures.Height = txtProcedures.Height - lstFiles.Height - 1215
Label1(2).Top = lstProcedures.Top - 255
frameOpts.Top = txtProcedures.Height + txtProcedures.Top - 120
Exit Sub

Sub_Form_Resize_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Resize]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Terminate()
' Inserted by LaVolpe
On Error GoTo Sub_Form_Terminate_General_ErrTrap_by_LaVolpe
Unload Me
Exit Sub

Sub_Form_Terminate_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Terminate]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
'===================================================================
'   When closing, remove any temporary files uses
'===================================================================

Dim tmpFile As String
' Inserted by LaVolpe
On Error GoTo Sub_Form_Unload_General_ErrTrap_by_LaVolpe
tmpFile = App.Path & "\" & "VBpro.tmp"
tmpFile = Replace(tmpFile, ":\\", ":\")
On Error Resume Next
Exit Sub

Sub_Form_Unload_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Unload]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub lstFiles_Click()
'===================================================================
' Read procedures contained within the file user clicked on
'===================================================================

' Inserted by LaVolpe
On Error GoTo Sub_lstFiles_Click_General_ErrTrap_by_LaVolpe
If lstFiles.ListIndex < 0 Then
    MsgBox "First select a file from the File Listing provided.", vbInformation + vbOKOnly
    Exit Sub
End If
If lstFiles.ListCount > 0 Then ReadVBfile
lstIndex.Tag = ""
txtProcedures.Tag = ""
Exit Sub

Sub_lstFiles_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lstFiles_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub lstFiles_GotFocus()
' Inserted by LaVolpe
On Error GoTo Sub_lstFiles_GotFocus_General_ErrTrap_by_LaVolpe
sBar.Panels(1).Text = "Click an item to view procedures it contains."
Exit Sub

Sub_lstFiles_GotFocus_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lstFiles_GotFocus]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub lstFiles_LostFocus()
' Inserted by LaVolpe
On Error GoTo Sub_lstFiles_LostFocus_General_ErrTrap_by_LaVolpe
sBar.Panels(1).Text = sBar.Panels(1).Tag
Exit Sub

Sub_lstFiles_LostFocus_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lstFiles_LostFocus]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub lstFiles_PathChange()
' Inserted by LaVolpe
On Error GoTo Sub_lstFiles_PathChange_General_ErrTrap_by_LaVolpe
lblPath.Caption = "Path:  " & lstFiles.Path
Exit Sub

Sub_lstFiles_PathChange_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lstFiles_PathChange]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub lstProcedures_Click()
'===================================================================
' When double clicked the procedure will be read & colored
'===================================================================

' Inserted by LaVolpe
On Error GoTo Sub_lstProcedures_DblClick_General_ErrTrap_by_LaVolpe
If lstFiles.ListIndex < 0 Or lstFiles.ListCount < 0 Then
    MsgBox "First select a file from the File Listing provided.", vbInformation + vbOKOnly
    Exit Sub
End If
If lstProcedures.SelCount = 0 Or lstProcedures.ListIndex < 0 Or lstProcedures.ListCount < 0 Then
    MsgBox "First select a valid procedure from the listing provided.", vbInformation + vbOKOnly
    Exit Sub
End If
On Error Resume Next
BuildReferences
ReadProcedure bAutoColor                                                   ' read the procedures & color it
cmdSearch.Tag = ""
Exit Sub

Sub_lstProcedures_DblClick_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lstProcedures_DblClick]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub lstProcedures_GotFocus()
' Inserted by LaVolpe
On Error GoTo Sub_lstProcedures_GotFocus_General_ErrTrap_by_LaVolpe
sBar.Panels(1).Text = "Click list item to View Procedure"
Exit Sub

Sub_lstProcedures_GotFocus_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lstProcedures_GotFocus]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub lstProcedures_LostFocus()
' Inserted by LaVolpe
On Error GoTo Sub_lstProcedures_LostFocus_General_ErrTrap_by_LaVolpe
sBar.Panels(1).Text = sBar.Panels(1).Tag
Exit Sub

Sub_lstProcedures_LostFocus_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lstProcedures_LostFocus]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub BuildReferences()
On Error Resume Next
myIndex = lstProcedures.ListIndex                       ' reference stored in case user starts clicking on the procedures list while this procedure is loaded
lstIndex.ListIndex = myIndex                                ' sync index listing
lstIndex.Tag = lstIndex.ItemData(myIndex) & "|" & lstIndex                  ' more references stored
txtProcedures.Tag = lstFiles.Path & "\" & lstFiles & "|" & lstProcedures ' more references stored
End Sub

Private Sub ReadVBfile()
'===================================================================
'   Core sub. Reads any text file & looks for specific text to determine where procedures start & stop
'===================================================================

Dim InFile As String, OutFile As String, sIndex As String, sTmp As String
Dim bStartFlag As Boolean, strEndString As String, strText As String
Dim fNr As Integer, FnrOut As Integer, ProCtr As Integer
Dim strBeginLine(1 To 5) As String, strProName As String, strProPrefix As String
Dim LineCtr As Long, ProStart As Long, bytePosition As Long, ItemNdx As Long, lNdx() As Long
Dim bGotDelcarations As Boolean, K As Integer, I As Integer, J As Integer
Dim sTargetWord As String, sDeclareString, sLineConnector As String

' Inserted by LaVolpe
On Error GoTo Sub_ReadVBfile_General_ErrTrap_by_LaVolpe
ReDim lNdx(0)
lstProcedures.Clear        ' clear found procedures listing
InFile = lstFiles.Path & "\" & lstFiles.FileName    ' build path to VB file
InFile = Replace(InFile, ":\\", ":\")
If Len(Dir(InFile)) = 0 Then         ' verify user-selected file exists
    MsgBox "The file [" & lstFiles.FileName & "] doesn't exist at " & lstFiles.Path
    Exit Sub
End If
strBeginLine(1) = "Sub "                    ' Build expected entrances into any VB module
strBeginLine(2) = "Function "             ' if any module can begin with another type of phrase,
strBeginLine(3) = "Property Get "       ' that phrase needs to be added here & the array
strBeginLine(4) = "Property Let "       ' index needs to be increased in the above DIM statements
strBeginLine(5) = "Property Set "

' Another important string of characters....
'   Since files are read top to bottom & some are forms/controls/etc that have a lot of information before
'   the procedures actually start, we need to identify those words which can be found at the beginning
'   of a line within the declarations section of a module. Declaration sections always come before the
'   actual procedures.  If one needs to be added, add it using the format of>  |word|
sDeclareString = "|Option|Private|Public|Global|#|Const|Declare|DefBool|DefByte|" & _
        "DefInt|DefLng|DefCur|DefSng|DefDbl|DefDec|DefDate|DefStr|DefObj|DefVar|"
MousePointer = vbHourglass
fNr = FreeFile()                                ' point to a temp file to hold the results of the scanned object
' Update the status bar & begin scanning
sBar.Panels(1).Text = "Scanning for Procedures..."
Open InFile For Binary As #fNr                   ' Open the file & start a line counter
LineCtr = 0: ProStart = -1
pBar = 0                                                      ' setup the progress bar
pBar.Scrolling = ccScrollingStandard
pBar.Max = LOF(fNr)
sBar.Panels(2).Text = ""
cmdClose.Caption = "Stop Reading"
bStop = False                                               ' flag to let user cancel the read process
Do While Loc(fNr) < pBar.Max                    ' start reading the file
    If bStop = True Then GoTo ResetMousePointer ' abort if user clicked the Stop button
    Line Input #fNr, strText        ' read a line & update the progress bar
    pBar = Loc(fNr)
    If bStartFlag = False Then    ' indicates the 1st procedure was found
        If ProStart < 0 Then    ' procedure entry point -- gets reset after each procedure is found
            sTargetWord = LTrim(strText)    ' parse the 1st word of each line
            I = InStr(sTargetWord, " ")           ' unless it is preceeded by a pound sign, in that case
            If I > 0 Then sTargetWord = Left(sTargetWord, I - 1)    ' use the pound sign as the 1st word
            If Left(sTargetWord, 1) = "#" Then sTargetWord = "#"   ' (i.e., #If, #Else, #End If)
            ' look at each possible procedure entry statement & see if this line of text matches
            For I = 1 To UBound(strBeginLine)
                For J = 1 To 4  ' for each entry point, you can have the following preceeding it, so check
                    strProPrefix = Choose(J, "Public ", "Private ", "", "Friend ")
                    ' do we have an actual starting point?
                    If Left(Trim(strText), Len(strProPrefix & strBeginLine(I))) = strProPrefix & strBeginLine(I) Then
                         bStartFlag = True  ' yep, so set flag to continue reading & set the expected End of the module
                         strEndString = "End " & Choose(I, "Sub", "Function", "Property", "Property", "Property") & " "
                         ProCtr = ProCtr + 1    ' increment a counter of number modules found
                        GoSub GetProcedureName  ' extract the procedure name
                        If bStartFlag = True Then
                            ' Also set where procedure found
                            ProStart = bytePosition - 1
                            Exit For
                        End If
                    End If
                Next
                If J < 5 Then Exit For  ' only < 5 if a procedure/module entry point was found
            Next
        ' not a procedure starting point, but is it a declarations starting point?
            If bGotDelcarations = False Then    ' we didn't find a delcarations section so let's check now
                If InStr(sDeclareString, "|" & sTargetWord & "|") > 0 Then
                    bGotDelcarations = True     ' got a known line within the declarations section of a module
                    GoSub GetProcedureName  ' call routine to store location where this was found
                End If
            End If
        End If
    Else    ' found a procedure now we're checking for the end of the procedure
        strText = strText & " "         ' add a space to end of string to force nullstrings to at least 1 character
        If Left(LTrim(strText), Len(strEndString)) = strEndString Then    ' end of sub/function
            GoSub UpdateListing       ' call routine to update listviews
            bStartFlag = False            ' reset flag to start looking for another procedures
        End If
    End If
ContinueParsing:
bytePosition = Loc(fNr)                 ' keep track of where each procedure line is located
DoEvents
Loop

Close #fNr
' Whew found all procedures we could find,  now let's tidy up and let user continue
' Update the delcarations statement in the listbox to indicate the results
If UBound(lNdx) = 0 Then ReDim Preserve lNdx(0 To 1)
If lNdx(LBound(lNdx)) = lNdx(LBound(lNdx) + 1) Then
    ' When the starting & ending lines of a procedure are equal, procedure wasn't found
    lstProcedures.List(0) = "Declarations - Section not Found"
Else    ' otherwise, it was found
    lstProcedures.List(0) = "Declarations"
End If
lstProcedures.AddItem "[Entire Document]", 0
GoSub SortIndex     ' sync the list index box with the procedure listing entries
Erase lNdx                  ' clear memory variables
ResetMousePointer:      ' finish updating the screen
MousePointer = vbDefault
If bStop = True Then lstProcedures.Clear
bStop = False
cmdClose.Caption = "&Close"
sBar.Panels(1).Text = sBar.Panels(1).Tag
sBar.Panels(2).Text = "Procedures Found: " & lstProcedures.ListCount - 3
pBar = 0
Close                           ' close any open files
Exit Sub

GetProcedureName:
' When a procedure/module or Declarations section is found, store its data
' If no items added to the procedures listing, and the GotDeclarations flag is set to true then this must
'   be a Declarations section we've found
sBar.Panels(2) = "Procedures Found: " & lstProcedures.ListCount + 1
If lstProcedures.ListCount = 0 And bGotDelcarations = True Then
    lstProcedures.AddItem ""                          ' add blank line to listitems
    ItemNdx = bytePosition - 1                       ' this is the holder for the Declarations section
    GoSub AddIndex
Else    ' No GotDeclarations flag, so we've found a procedure/module
    K = InStr(strText, "(")                             ' start stripping the module name.. find first parenthesis
    If K = 0 Then
        bStartFlag = False
        Return
    End If
    strProName = Left(strText, K - 1)          ' truncate line to that point i.e.,  Private Sub Name
    K = InStrRev(strProName, " ")               ' find first space from end of string       ^     ^
    K = InStrRev(strProName, " ", K - 1)      ' find the next space before that space  ||
    strProName = Mid(strProName, K + 1) ' module name including Sub, Function, Get, Let, or Set
    ' but if this is a property, include the word "property" with the module name & store to temp file
    If InStr(strText, "Property ") > 0 Then strProName = "Property " & strProName
    ' depending on whether not a declarations section was found, update the listitems as follows
    If lstProcedures.ListCount = 0 Then
        ' no declarations section found, so add a blank entry as a holder
        lstProcedures.AddItem ""        ' Declarations section will be added later
        ItemNdx = bytePosition - 1
        GoSub AddIndex
    End If
End If
Return

UpdateListing:
' This routine is called when a funciton is found and the subsequent End statement was found
'   Only then will a procedure be displayed in the listview
lstProcedures.AddItem strProName       ' Add the procedure to listview
ItemNdx = ProStart
GoSub AddIndex
ProStart = -1               ' reset flag to indicate no new procedure found yet
sBar.Panels(2) = "Procedures Found: " & lstProcedures.ListCount
Return

AddIndex:
' Since the procedures listing is sorted, we need to track where each procedure starts & ends
'   because the procedures listed don't necessarily follow each other as listed in the file
ReDim Preserve lNdx(0 To lstProcedures.ListCount)   ' increment the variable by 1 for each procedure
lNdx(UBound(lNdx) - 1) = ItemNdx                            ' set the starting location to that variable
lstProcedures.ItemData(lstProcedures.NewIndex) = UBound(lNdx) - 1   ' set a reference in the listbox
Return                                                                                                       ' to its variable location

SortIndex:
' After all procedures are read, upate the index listbox synchronizing the starting/stopping locations
'   of each procedure in the file
lstIndex.Clear
lNdx(UBound(lNdx)) = pBar.Max               ' set the last array index to the file length
For I = 0 To lstProcedures.ListCount - 1     ' loop thru each procedure listed in the procedures listing
    lstIndex.AddItem "Index"                         ' and add a blank entry in the index listing
Next
lstIndex.List(0) = FileLen(InFile)
lstIndex.ItemData(0) = 0
For I = 1 To lstProcedures.ListCount - 1    ' now go thur each procedure again
    ItemNdx = lstProcedures.ItemData(I)     ' extract the reference for that procedure to the variable
    lstIndex.List(I) = lNdx(ItemNdx + 1) - 1    ' containing its start/stop position & update the
    lstIndex.ItemData(I) = lNdx(ItemNdx)        ' index listing with that information
Next
Return
Exit Sub

Sub_ReadVBfile_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub ReadVBfile]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
Resume ResetMousePointer
End Sub

Private Sub ReadProcedure(Optional bColor As Boolean = False)
'===================================================================
'   The other key procedure for this program. Reads specific procedures within a file
'===================================================================

Dim fNr As Integer, InFile As String, strText As String
Dim bStartFlag As Boolean, strEndString As String
Dim strCriteria As String, LineCounter As Long
Dim lStart As Long, lStop As Long
Dim OutFile As String, FnrO As Integer, I As Integer, Looper As Long, nrChunks As Integer
Dim BuffBytes() As Byte, bytesRemaining As Long, lChunks As Long, bytesNeeded As Long

' Inserted by LaVolpe
On Error GoTo Sub_ReadProcedure_General_ErrTrap_by_LaVolpe
InFile = Left(txtProcedures.Tag, InStr(txtProcedures.Tag, "|") - 1) 'get the file name
If Len(Dir(InFile)) = 0 Then                                                        ' and ensure it exists
    MousePointer = vbDefault
    MsgBox "The file [" & StripFile(InFile, "N") & "] doesn't exist at " & StripFile(InFile, "P")
    Exit Sub
End If
pBar = 0                                            ' set up some display objects
Caption = "Procedure View for " & Mid(txtProcedures.Tag, InStr(txtProcedures.Tag, "|") + 1)
sBar.Panels(1) = "Loading Procedure(s)"
OutFile = App.Path & "\" & "VBpro.tmp"  ' identify a temporary file to hold the procedure
OutFile = Replace(OutFile, ":\\", ":\")         ' the procdure is actually extracted & displayed in its own file
If Len(Dir(OutFile)) > 0 Then
    SetAttr OutFile, vbNormal
    Kill OutFile  ' if that temp file already exists, kill it now
End If
' identify where the selected procedure starts & stops
lStart = Val(Left(lstIndex.Tag, InStr(lstIndex.Tag, "|") - 1))
lStop = Val(Mid(lstIndex.Tag, InStr(lstIndex.Tag, "|") + 1))
' if the procedure doesn't exist (no Declaration Section for example), notify
If lStop - lStart < 1 Then
    txtProcedures.Text = ""
    cmdAdd(0).Enabled = False
    cmdAdd(1).Enabled = False
    MsgBox "No code was found for the procedure: " & Mid(txtProcedures.Tag, InStr(txtProcedures.Tag, "|") + 1), vbInformation + vbOKOnly
    GoTo ResetMousePointer
End If
lChunks = 32768                             ' size of bytes to read at a time
fNr = FreeFile()                                ' open the source file
Open InFile For Binary As #fNr
FnrO = FreeFile()
Open OutFile For Binary As #FnrO    ' open the temporary file
If lStart < 0 Then lStart = 0
ReDim BuffBytes(lChunks)
nrChunks = Int(((lStart) / lChunks))            ' see how many times to loop thru reading chunks
bytesRemaining = (lStart) Mod lChunks      ' and how many to finish reading at end of loop
For Looper = 1 To nrChunks                      ' this is done to strip file to starting point of procedure
    Get #fNr, , BuffBytes()                             ' and the procedure is extremely fast
Next
If bytesRemaining - nrChunks > 0 Then
    ReDim BuffBytes(bytesRemaining - nrChunks)
    Get #fNr, , BuffBytes()
End If

ReDim BuffBytes(lChunks)                            ' now at starting point,need to read to end point of
nrChunks = Int(((lStop - lStart) / lChunks))    ' procedure within the file. Same principle as above
bytesRemaining = (lStop - lStart) Mod lChunks   'except this time, we write what is read to the temp file
For Looper = 1 To nrChunks
    Get #fNr, , BuffBytes()
    Put #FnrO, , BuffBytes()
Next
ReDim BuffBytes(bytesRemaining - nrChunks)
Get #fNr, , BuffBytes()
Put #FnrO, , BuffBytes()
'----------------------------------------------------------------
Close #fNr: Close #FnrO                         ' close & load the file
If bColor = False Then
    txtProcedures = ""
    txtProcedures.SelStart = 0
    txtProcedures.SelLength = Len(txtProcedures.Text)
    txtProcedures.SelColor = 0
    txtProcedures.LoadFile OutFile, 1                  ' update the screen info & prep the hidden text box for coloring
Else
    pBar = 0
    sBar.Panels(1) = "Loading Procedure"
    frmLibrary.rtfStaging.LoadFile OutFile, 1
    cmdColor.Caption = "Stop Color"
    GP = Null
    CheckLine4KeyWords txtProcedures, frmLibrary.rtfStaging, , True, pBar
    cmdColor.Caption = "Color"
End If
sBar.Panels(6).Text = Format(Len(txtProcedures.Text), "#,000")
ResetMousePointer:                                  ' enable text box & reset base variables used to inform
MousePointer = vbDefault
sBar.Panels(1).Text = sBar.Panels(1).Tag
cmdAdd(0).Enabled = (Len(txtProcedures.Text) > 0)
cmdAdd(1).Enabled = cmdAdd(0).Enabled
Exit Sub

Sub_ReadProcedure_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub ReadProcedure]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub txtProcedures_GotFocus()
sBar.Panels(1) = "Edit contents before adding to repository."
sBar.Panels(4) = txtProcedures.GetLineFromChar(txtProcedures.SelStart) + 1
End Sub

Private Sub txtProcedures_LostFocus()
sBar.Panels(1) = sBar.Panels(1).Tag
End Sub

Private Sub txtProcedures_SelChange()
sBar.Panels(4) = txtProcedures.GetLineFromChar(txtProcedures.SelStart) + 1
End Sub
