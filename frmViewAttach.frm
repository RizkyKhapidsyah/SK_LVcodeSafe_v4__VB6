VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewAttach 
   Caption         =   "Attachment - "
   ClientHeight    =   8325
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8580
   HelpContextID   =   7
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Height          =   300
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   30
      Width           =   960
   End
   Begin VB.CheckBox chkWrap 
      Caption         =   "Word Wrap Active"
      Height          =   225
      Left            =   1485
      TabIndex        =   3
      Top             =   60
      Value           =   1  'Checked
      Width           =   1800
   End
   Begin VB.OptionButton optView 
      Caption         =   "Save as Rich Text Format (RTF)"
      Height          =   225
      Index           =   0
      Left            =   3570
      TabIndex        =   2
      Top             =   90
      Width           =   2895
   End
   Begin VB.OptionButton optView 
      Caption         =   "Save as ASCII text"
      Height          =   225
      Index           =   1
      Left            =   6555
      TabIndex        =   1
      Top             =   90
      Width           =   1875
   End
   Begin RichTextLib.RichTextBox rtfAttach 
      Height          =   7965
      Left            =   30
      TabIndex        =   0
      Top             =   300
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   14049
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmViewAttach.frx":0000
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Reload - Undo"
         Index           =   2
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   4
         Shortcut        =   +{F4}
      End
   End
End
Attribute VB_Name = "frmViewAttach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bDirty As Boolean                       ' set to true if data changes on this form
Private bInvalidRecord As Boolean           ' set to true if record not found in database (should never happen
                                                                '   unless user is simultaneously modifying database while using this program)
Private sFileIn As String                           ' contains the filename

Private Sub DoExtraction()
'================================================================
'   Function will extract call function to extract data to a temporary file
'================================================================
Dim sCaption As String, ActualSize As Long
On Error Resume Next
bInvalidRecord = True       ' initially set this value to true
rtfAttach = ""                      ' set the RTF text box to nothing
If rsAttachment.RecordCount = 0 Then    ' If no record exists to display then...
    rtfAttach = vbCrLf & vbCrLf & "NO ATTACHMENT FOUND MATCHING REQUESTED RECORD"
    sCaption = "Failed Attachment Load" ' update the caption with this error
Else                                    ' if a record does exist then, try to extract it to a file
    ' now attempt to extract the db record to the filename
    ' because of the way Access saves binary text information, the FieldSize property of the
    '   attachment may exceed its actual size. When the attachment was originally saved, the
    '   actual file size in bytes was saved to the no longer used Field: Viewer.
    '   However, prior to v3 this value wasn't being saved so we need to check both fields
    '   to determine the actual size
    If rsAttachment.Fields("Viewer") > 0 Then
        ActualSize = rsAttachment.Fields("Viewer")
    Else
        ActualSize = rsAttachment.Fields("Attachment").FieldSize
    End If
    If ExtractAttachment(rsAttachment.Fields("Attachment"), sFileIn, ActualSize) = True Then
        rtfAttach.LoadFile sFileIn                                      ' success
        sCaption = rsAttachment.Fields("Description")       ' update caption
    Else                                                                           ' failure - loading file, update caption & display error
        rtfAttach = vbCrLf & vbCrLf & "ERROR EXTRACTING ATTACHMENT"
        sCaption = "Failed Attachment Load"
    End If
End If
Caption = Left(Caption, 13) & sCaption          ' actual caption update
bDirty = False                                                  ' set flag to false initially
End Sub

Private Sub chkWrap_Click()
SendMessageLong rtfAttach.hWnd, EM_SETTARGETDEVICE, 0, Abs(chkWrap.Value - 1)
End Sub

Private Sub Form_Load()
'================================================================
' Sub simply sets some default values and calls function to load attachment to RTF textbox
'================================================================
Dim strSQL As String
' Inserted by LaVolpe
On Error GoTo Sub_Form_Load_General_ErrTrap_by_LaVolpe
optView(1) = True           ' set save as Text as the default
Icon = frmLibrary.SmallImages.ListImages(13).ExtractIcon    ' load the attachment icon in titlebar
sFileIn = GetUniqueFileName(".any", CStr(GP))                   ' set a temporary filename to use
rtfAttach.Font.Name = MyDefaults.Font
rtfAttach.Font.Size = CSng(MyDefaults.FontSize)
cmdSearch.Picture = frmLibrary.SmallImages.ListImages(6).ExtractIcon
DoExtraction                                                                        ' call function to load attachment in RTF textbox
rtfAttach.HideSelection = False
Exit Sub

Sub_Form_Load_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Load]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then Exit Sub
On Error Resume Next
rtfAttach.Width = Width - 195
rtfAttach.Height = Height - 1050
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
'   Unloads the form & offers to save any changes if not already saved
'================================================================
On Error Resume Next
If bDirty = True Then           ' changes made
    Dim I As Integer                ' offer to save changes
    I = MsgBox("The attachment was changed. Do you want to save the update?", vbYesNo + vbQuestion)
    If I = vbYes Then SaveChanges              ' ok, changes should be saved
End If
bDirty = False
End Sub

Private Sub SaveChanges()
'================================================================
'   Saves the current text as the attachment
'================================================================
' Inserted by LaVolpe
On Error GoTo Sub_SaveChanges_General_ErrTrap_by_LaVolpe
Dim ActualSize As Long
rtfAttach.SaveFile sFileIn, Val(optView(0).Tag)  ' save the textbox contents first
' attempt a save and prompt with success/failure
rsAttachment.Edit
ActualSize = LoadAttach(rsAttachment.Fields("Attachment"), sFileIn)
If ActualSize > -1 Then
    rsAttachment.Fields("Viewer") = ActualSize
    rsAttachment.Update
    MsgBox "Attachment Updated"
    rsAttachment.Requery            ' requery
    bDirty = False
Else
    rsAttachment.CancelUpdate
    MsgBox "Couldn't save changes. Try again.", vbExclamation + vbOKOnly + vbExclamation
End If
Exit Sub

Sub_SaveChanges_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub SaveChanges]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub
Private Sub mnuFile_Click(Index As Integer)
'================================================================
'   The File menu bar option
'================================================================
Dim sFileName As String
' Inserted by LaVolpe
On Error GoTo Sub_mnuFile_Click_General_ErrTrap_by_LaVolpe
Select Case Index
Case 0:         ' Save
    SaveChanges
Case 2:         ' Reload/Undo
    DoExtraction
Case 5:         ' Quit
    Unload Me
End Select
Exit Sub

Sub_mnuFile_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub mnuFile_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub optView_Click(Index As Integer)
'================================================================
'   Changes the option to save as text or RTF
'================================================================
' Inserted by LaVolpe
On Error GoTo Sub_optView_Click_General_ErrTrap_by_LaVolpe
optView(0).Tag = Index
Exit Sub

Sub_optView_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub optView_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub rtfAttach_Change()
'================================================================
'   Track if attachment has been modified
'================================================================
' Inserted by LaVolpe
On Error GoTo Sub_rtfAttach_Change_General_ErrTrap_by_LaVolpe
bDirty = True
Exit Sub

Sub_rtfAttach_Change_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub rtfAttach_Change]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub


Private Sub cmdSearch_Click()
Dim sCriteria As String, sMsg As String, lStartSearch As Long, iWholeWord As Integer
sMsg = "Enter the search string below." & vbCrLf & "For whole-word matches start search string with an exclamation mark (!)"
sCriteria = InputBox(sMsg, "Search Criteria", cmdSearch.Tag)
If Trim(sCriteria) = "" Or Trim(sCriteria) = "!" Then Exit Sub
cmdSearch.Tag = sCriteria
With rtfAttach
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

