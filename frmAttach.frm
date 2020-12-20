VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAttachments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Attachments"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   HelpContextID   =   7
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkNoPrompt 
      Height          =   195
      Left            =   4785
      TabIndex        =   7
      ToolTipText     =   "otherwise you will be prompted for a name"
      Top             =   6600
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CheckBox chkNoNamePrompt 
      Height          =   195
      Left            =   7200
      TabIndex        =   13
      ToolTipText     =   "otherwise you will be prompted for a name"
      Top             =   5625
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Save A&LL attachments to Files"
      Height          =   585
      Index           =   3
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5940
      Width           =   1845
   End
   Begin VB.FileListBox fileAll 
      Enabled         =   0   'False
      Height          =   675
      Hidden          =   -1  'True
      Left            =   4425
      Pattern         =   "~Atch*.*"
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5865
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton cmdSaveDescription 
      Caption         =   "Save Change"
      Height          =   285
      Index           =   3
      Left            =   8010
      TabIndex        =   9
      Top             =   1080
      Width           =   1185
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View on Screen Only (Text-Files)"
      Height          =   585
      Index           =   0
      Left            =   6030
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2985
      Width           =   2295
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View in default &Application for attachment's extension"
      Height          =   645
      Index           =   1
      Left            =   6030
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4290
      Width           =   2295
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Save selected attachments to File(s)"
      Height          =   585
      Index           =   2
      Left            =   5385
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5940
      Width           =   1845
   End
   Begin VB.TextBox txtDescription 
      Height          =   315
      Left            =   5160
      TabIndex        =   8
      Text            =   "Description"
      ToolTipText     =   "for Display purposes only"
      Top             =   1350
      Width           =   4035
   End
   Begin VB.CommandButton cmdRemoveAttach 
      Caption         =   "Save && Delete Selected"
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      ToolTipText     =   "Save to a file & remove attachment"
      Top             =   2730
      Width           =   2535
   End
   Begin VB.CommandButton cmdRemoveAttach 
      Caption         =   "Delete Selected"
      Height          =   375
      Index           =   0
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Permanently delete attachment"
      Top             =   2730
      Width           =   2415
   End
   Begin VB.CommandButton cmdRemoveAttach 
      Caption         =   "Remove those Selected"
      Height          =   375
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   6855
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   5325
      Top             =   4350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvAttach 
      Height          =   2400
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   315
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   8291
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "E&xit"
      Height          =   375
      Index           =   1
      Left            =   5370
      TabIndex        =   15
      Top             =   6855
      Width           =   3825
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Add ALL as Attachments"
      Height          =   375
      Index           =   0
      Left            =   2535
      TabIndex        =   6
      Top             =   6855
      Width           =   2505
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "New Attachment(s)"
      Height          =   495
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3285
      Width           =   1890
   End
   Begin MSComctlLib.ListView lvAttach 
      Height          =   2400
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   4170
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   4233
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File - Click Header to Toggle Display"
         Object.Width           =   8291
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "When adding, use file name as attachment name"
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
      Height          =   195
      Index           =   11
      Left            =   165
      TabIndex        =   29
      ToolTipText     =   "otherwise you will be prompted for a name"
      Top             =   6600
      Width           =   4545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Options for the most recent attachment selected in your listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   5
      Left            =   5205
      TabIndex        =   26
      Top             =   195
      Width           =   3915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The file will simply be copied to a location you choose"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   5340
      TabIndex        =   25
      Top             =   5085
      Width           =   3825
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The file will be saved to a location you choose and will then be opened"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Index           =   3
      Left            =   5370
      TabIndex        =   24
      Top             =   3720
      Width           =   3750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "An on-screen copy will be displayed and you have the option to save to a file"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Index           =   1
      Left            =   5310
      TabIndex        =   23
      Top             =   2475
      Width           =   3870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description (Change as needed)"
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
      Height          =   285
      Index           =   6
      Left            =   5205
      TabIndex        =   22
      Top             =   1125
      Width           =   4035
   End
   Begin VB.Label lblData 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Original file name"
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
      Height          =   300
      Left            =   5175
      TabIndex        =   21
      Top             =   2010
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Original file name (FYI only, can't be changed)"
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
      Height          =   285
      Index           =   9
      Left            =   5190
      TabIndex        =   20
      Top             =   1755
      Width           =   4035
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click to Add Multiple Files at once"
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
      Height          =   255
      Index           =   8
      Left            =   2025
      TabIndex        =   19
      Top             =   3525
      Width           =   2985
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click to Add 1 file at a time"
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
      Height          =   255
      Index           =   7
      Left            =   2025
      TabIndex        =   18
      Top             =   3285
      Width           =   2985
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attachments to Add for above code"
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
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   17
      Top             =   3930
      Width           =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Listing of Attachments for above code"
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
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   16
      Top             =   75
      Width           =   3930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Original filename to be      used as new filename"
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
      Height          =   195
      Index           =   10
      Left            =   5175
      TabIndex        =   28
      ToolTipText     =   "otherwise you will be prompted for a name"
      Top             =   5610
      Width           =   4110
   End
End
Attribute VB_Name = "frmAttachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bDirty As Boolean           ' flag to identify if data changed
Private sMyPath As String           ' default path to store temporary file from the database
'============================================================
' This is the main form used for all Attachment actions....
'   - Adding attachments to a record
'   - Deleting attachments from a record
'   - Editing the attachment description
'   - Extracting attachments from the db to local files
'   - Viewing attachments on screen or via default viewer
'============================================================

Private Sub cmdAdd_Click()
'============================================================
' Function loads files from the hard drive into a listbox
'============================================================
On Error GoTo CnxAdd
With dlgCommon                  ' set up the open dialog box
    .DialogTitle = "Which file(s) do you want to add?"
    .Filter = "All Files|*.*"
    .FileName = ""
    .CancelError = True
    .Flags = Val(cmdAdd.Tag)    ' this tag is set when user clicks option to select 1 or multiple files
    .MaxFileSize = 256 + (Abs(CBool(Val(cmdAdd.Tag))) * 4 * 256)    ' set size in relation to 1 or multi files
End With
dlgCommon.ShowOpen              ' open the dialog box
MousePointer = vbHourglass      ' Begin loading selected file(s) into the listbox
On Error GoTo FailedOpen
Dim I As Integer, sDescription As String, J As Integer, K As Integer, sPath As String
Dim sLongFN As String, sFileName As String, itmX As ListItem
'============================================================
'   Since the multiple file selection window truncates all file names down to their DOS 8.3 filenames,
'   this function will locate the file and return its long file name for display purposes
'   why be forced to display the old 8.3 formats?
'============================================================
If Val(cmdAdd.Tag) = 0 Then                 ' user opted for single file Adds
    sPath = StripFile(dlgCommon.FileName, "P")          ' identify the path from the dialog box
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    fileAll.Path = sPath                                                  ' set file list box path to the same
    sFileName = dlgCommon.FileTitle                          ' identify the file name from the dialog box
    GoSub GetLongFileName4Display                         ' call sub to add item to the list
Else                                                        ' user opted for multiple file Adds
    ' since these use the 8.3 method of filenames, each individual file in the list returned
    '           are delimited by a space, including the path which is delimited by a space
    '   one exception: if only one file was returned then the filename includes the full path & file name, no spaces
    sPath = ExtractData(dlgCommon.FileName, " ", 1)         ' extract the first field in the multi string
    If sPath = dlgCommon.FileName Then                          ' if the 1st field is the entire filename, then only 1 file was chosen
        sPath = StripFile(dlgCommon.FileName, "P")          ' extract just the path
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        fileAll.Path = sPath                                                   ' set file list box path to the same
        sFileName = StripFile(dlgCommon.FileName, "N")  ' now identify the filename
        GoSub GetLongFileName4Display                           ' call sub to add item to list
    Else        ' Now we have at least 2 files selected
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"      ' add trailing slash to path & set file list box path to same
        fileAll.Path = sPath
        For J = 2 To 500                                                      ' we're assuming no more than 500 files being added
            sFileName = ExtractData(dlgCommon.FileName, " ", J) ' extract each filename
            If sFileName = "" Then Exit For                                     ' if no more filenames, then exit the loop
            GoSub GetLongFileName4Display                               ' call sub to add filename to listing
        Next
    End If
End If

RestoreMaxFileSize:                                                             ' cleans up MaxFileSize if it was increased earlier
dlgCommon.MaxFileSize = 256
MousePointer = vbDefault
Exit Sub

GetLongFileName4Display:
'============================================================
' Subroutine will try to pull the long filename from a short filename & then add the item to the list box
'   I use a little trick here.  The file list box seems to always return long file names, even if you
'   pass a short filename to it in the pattern property.
'============================================================
fileAll.Pattern = sFileName      ' set the file listbox pattern to display only the filename passed
Select Case fileAll.ListCount
Case 0:     ' no matching files found, shouldn't  happen, but if it does just use the passed filename
    sLongFN = sFileName
Case 1:     ' should always be the case
    sLongFN = fileAll.List(0)
Case Else:  ' more than one file name returned, so let's compare them against each in the list
    For K = 0 To fileAll.ListCount - 1
        ' pass the full LONG file name to a function which will return the DOS 8.3 filename
        sLongFN = StripFile(GetShortPathName(sPath & fileAll.List(K)), "N")
        If sLongFN = sFileName Then Exit For    ' if the two filenames match then exit the loop
    Next
    ' if we have a match, use the LONG filename, otherwise use the passed short filename
    If K < fileAll.ListCount Then sLongFN = fileAll.List(K) Else sLongFN = sFileName
End Select
' Now let's add it to the listbox
If lvAttach(1).ColumnHeaders(1).Tag = "1" Then          ' toggle is displaying full path & filename
    Set itmX = lvAttach(1).ListItems.Add(, , sPath & sLongFN)   ' add as full path & filename
Else                                                                              ' toggle is display just the filename
    Set itmX = lvAttach(1).ListItems.Add(, , sLongFN)               ' add filename only
End If
itmX.Tag = sPath                                                          ' place path in Tag property for use with toggle
Return

CnxAdd:                                         ' Error routine
MousePointer = vbDefault
If Err.Number = 20476 Then          ' When multiple files are being selected, they are returned as a string
    ' if the length of the string is too short, this error appears. Instead of playing the game of increasing memory
    '   incrementally to see if it's long enough, let's just force the user to select less files at a time
    MsgBox "Selected too many files to store in memory. Please try again and select a few less at a time.", vbInformation + vbOKOnly
    Resume
Else    ' other errors
    If Err.Number = 32755 Then Exit Sub     ' this is the error when user presses the Cancel button
    ' for all other errors, display the error message
    MsgBox "Following error preventing including the files you chose." & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End If
Resume RestoreMaxFileSize  ' if not resuming within the program, resume to the point of reseting memory file size
Exit Sub

FailedOpen:
MsgBox "Some or all of the requested file to load failed for the following reason." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Function ListSelected(LstID As Integer) As Boolean
'============================================================
' routine to verify an item in the passed listview is selected or not
'============================================================
If lvAttach(LstID).SelectedItem Is Nothing Then ListSelected = False Else ListSelected = True
' if no items are selected, then show this message box
If ListSelected = False Then MsgBox "First select an item from the list.", vbInformation + vbOKOnly
End Function

Private Sub cmdRemoveAttach_Click(Index As Integer)
'============================================================
' Option to remove attachments from the 2 listviews
'============================================================
If ListSelected(Choose(Index + 1, 0, 1, 0)) = False Then Exit Sub       ' verify items selected first

Dim sMsg As String, I As Integer, bBypass As Boolean, bRefresh As Boolean
Select Case Index
Case 0      ' Removing attachment from the database
    sMsg = "Are you absolutely sure you want the selected attachments deleted?" & vbCrLf & vbCrLf _
        & "WARNING: Ensure these are not your only copies. If so, suggest saving them to a file before deleting them here."
Case 1      ' Removing from the initial file-select listing
    sMsg = "Are you sure you want selected files removed? They haven't been added as attachments yet."
Case 2      ' Removing from the database, but saving to a file first
    sMsg = "Are you sure you want the selected attachments deleted from the database after they are saved to files?"
End Select
' Provide the confirmation warning
I = MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If I = vbNo Then Exit Sub

If Index = 1 Then   ' not in database yet, so just remove them from the listview
    For I = lvAttach(1).ListItems.Count To 1 Step -1    ' Remove all selected items
        If lvAttach(1).ListItems(I).Selected = True Then lvAttach(1).ListItems.Remove I
    Next
Else        ' Ok, these are already in the database
    If Index = 2 Then
        I = MsgBox("Press NO to use the original file name as the new file name" & vbCrLf & vbCrLf _
            & "Press YES to prompt for file names", vbYesNoCancel + vbQuestion, "Prompt for file names?")
        If I = vbCancel Then Exit Sub
        If I = vbNo Then chkNoNamePrompt = 1 Else chkNoNamePrompt = 0
        SaveAttachment 2, True
        bDirty = True
    Else
        With lvAttach(0)
        For I = .ListItems.Count To 1 Step -1
            If .ListItems(I).Selected = True Then
                rsAttachment.FindFirst "[ID]=" & Mid(.ListItems(I).Key, 7)
                If rsAttachment.NoMatch = False Then
                    rsAttachment.Delete
                    .ListItems.Remove I
                    bDirty = True
                End If
            End If
        Next
        End With
    End If
End If
' Cleanup
UserCnx:
Call lvAttach_Click(0)
If bDirty = True Then        ' deletions were made from the db, so let's requery & redisplay
    rsAttachment.Requery
    Call lvAttach_Click(0)
End If
End Sub

Private Sub cmdSave_Click(Index As Integer)
'============================================================
' Function saves attachments to the database or exits the window
'============================================================

If Index = 0 Then           ' Saving vs Exiting
    If lvAttach(1).ListItems.Count = 0 Then Exit Sub        ' can't save if no files are identified
    Dim itmX As ListItem, sDescription As String, I As Integer, sFileName As String
    Dim ActualSize As Long, bFailedSave As Boolean, newID As Long
    With rsAttachment
        For I = lvAttach(1).ListItems.Count To 1 Step -1    ' for every file identified, let's try to save it to the db
            sFileName = StripFile(lvAttach(1).ListItems(I).Text, "N")   ' get the filename of the new attachment
            If chkNoPrompt = 0 Then         ' prompt for filename, otherwise use filename as attachment name
                ' Now provide a popup to alter the filename for description purposes, providing current filename as default
                sDescription = InputBox("Enter a description for the following file or press cancel to skip adding this file", _
                    "Attachment Description", sFileName)
            Else
                sDescription = sFileName
            End If
            If Len(sDescription) Then   ' If user pressed cancel on above Input box, then these steps would be
                .AddNew                       '     skipped for this record only
                .Fields("Description") = Left(sDescription, 150)    ' add the description to db field
                .Fields("RecIDRef") = CLng(Tag)                         ' add this codes db ID to db field
                .Fields("FileName") = sFileName                         ' add attachment filename to db field
                ' Now call function to store the file into the OLE db field and if it fails we abort the add function
                ActualSize = LoadAttach(.Fields("Attachment"), lvAttach(1).ListItems(I).Tag & sFileName)
                If ActualSize > -1 Then
                    .Fields("Viewer") = ActualSize
                    newID = .Fields("ID")   ' success! let's finish the db record.  Track new record id
                    .Update                         ' update the database & add the record to the attachment listing
                    Set itmX = lvAttach(0).ListItems.Add(, "RecID:" & newID, sDescription)
                    itmX.Tag = sFileName
                    lvAttach(1).ListItems.Remove I  ' remove the file from the pre-save listview
                    bDirty = True                             ' changes made, flag it
                Else            ' if the file failed loading into the db, then this would be needed 'cause we started the Add
                    .CancelUpdate: bFailedSave = True
                End If
            End If
        Next
    End With
    Set itmX = Nothing
    If bFailedSave = True Then  ' If an item failed to load into db, prompt user
        MsgBox "The files remaining in the left listing could not be saved. Try saving them again, one at a time.", vbExclamation + vbOKOnly
    Else                                    ' otherwise if we did save any attachments, prompt with success
        If bDirty = True Then MsgBox "Attachments have been saved to the database", vbInformation + vbOKOnly
    End If
    Call lvAttach_Click(0)
Else                        ' if user is exiting, then exit here
    Unload Me
End If
End Sub

Private Sub cmdSaveDescription_Click(Index As Integer)
'============================================================
' Changes the attachment name in the db and on the form
'============================================================

If ListSelected(0) = False Then Exit Sub        ' ensure a list item is selected
If Len(txtDescription) = 0 Then                     ' don't allow blank descriptions
    MsgBox "The description is blank.  No update made", vbInformation + vbOKOnly
    Exit Sub
End If
' Save procedure: simply find the matching record in the db, update it with new description
'   then update the form & flag the update status
With rsAttachment
    .FindFirst "[ID]=" & Mid(lvAttach(0).SelectedItem.Key, 7)
    .Edit
    .Fields("Description") = txtDescription
    .Update
    bDirty = True
End With
lvAttach(0).SelectedItem.Text = txtDescription
End Sub

Private Sub Form_Load()
'============================================================
' Set up the initial display of the window
'============================================================
' following is default where db records are extracted to for the purpose of viewing
sMyPath = App.Path                                  ' Going to set a default path to the program's path
sMyPath = Replace(sMyPath, ":\\", ":\")       ' Check on VB error
If Right(sMyPath, 1) <> "\" Then sMyPath = sMyPath & "\"    ' add trailing backslash

' When this form is loaded, the GP variable comes in 1 of 2 formats:
'   1:  previous form's caption only
'   2: specific attachment ID | previous form's caption
lvAttach(1).ListItems.Clear
If InStr(GP, "|") Then  ' 2nd format
    Caption = "Attachment(s) for - " & Mid(GP, InStr(GP, "|") + 1)  ' set caption
    GP = "RecID:" & Val(GP)                                                         ' set Key property for later use
Else
    Caption = "Attachment(s) for - " & GP                                       ' set caption
    GP = Null                                                                                   ' set Key property to default
End If
Tag = DBrecID                                           ' keep track of the parent's db Record ID
LoadRecordset                                           ' load all attachments, if any
DoGradient Me, 2                                       ' repaint form & change label colors as needed
Dim I As Integer
For I = Label1.LBound To Label1.UBound: Label1(I).ForeColor = MyDefaults.LblColorPopup: Next
lblData.ForeColor = MyDefaults.LblColorPopup
Call lvAttach_Click(0)                                  ' Load selected attachment info to description field
End Sub

Private Sub LoadRecordset()
'============================================================
'   Sub will load the attachment listing from the db
'============================================================

On Error Resume Next
rsAttachment.Close      ' Close recordset if already open

Set rsAttachment = mainDB.OpenRecordset("Select * From tblAttachments " & _
    "Where ( RecIDref =" & Tag & ");", dbOpenDynaset)       ' open recordset for reference/editing
Dim itmX As ListItem
lvAttach(0).ListItems.Clear
With rsAttachment
    If .RecordCount > 0 Then                                        ' for each attachment, list it in the listview
        .MoveFirst
        Do While .EOF = False
            Set itmX = lvAttach(0).ListItems.Add(, "RecID:" & .Fields("ID"), .Fields("Description"))
            itmX.Tag = .Fields("FileName")
            If Not IsNull(GP) Then                                  ' if a specific attachment passed to form, look for it
                If itmX.Key = GP Then                               ' found a match?
                    lvAttach(0).SelectedItem.Selected = False   ' yep, so let's unselect any selected item
                    itmX.Selected = True
                End If
            End If
            .MoveNext
        Loop
    End If
End With
Set itmX = Nothing
' now if a match was found, let's select it here
If Not IsNull(GP) Then lvAttach(0).ListItems(GP).Selected = True
End Sub

Private Sub Form_Terminate()
'============================================================
Unload Me
'============================================================
End Sub

Private Sub Form_Unload(Cancel As Integer)
'============================================================
'    Attempts to delete any temporary files created while window was open and identifies
'       net result of whether or not any attachments were added/deleted/edited so calling form
'       can run appropriate routines
'============================================================
DestroyTmpFiles
If bDirty = True Then GP = "Reload" Else GP = Null
End Sub

Private Sub Label1_Click(Index As Integer)
'============================================================
' The labels toggling between single attachment selection or multi-selection
'============================================================
Select Case Index
Case 7, 8   'option to select 1 file or multi-files
    ' Toggle the back colors of the two labels
    If Index = 7 Then Label1(8).BackStyle = 0 Else Label1(7).BackStyle = 0
    Label1(Index).BackStyle = 1
    Label1(Index).Refresh
    ' Set a Tag property to be used by the Dialog Box "Flags" property when dialog box is displayed
    cmdAdd.Tag = Choose(Index - 6, 0, cdlOFNAllowMultiselect)
Case 10
    chkNoNamePrompt = Abs(chkNoNamePrompt.Value - 1)
Case 11
    chkNoPrompt = Abs(chkNoPrompt.Value - 1)
End Select
End Sub

Private Sub lvAttach_Click(Index As Integer)
'============================================================
' Updates description and original file info each time an attachment is selected
'============================================================

If lvAttach(Index).ListItems.Count = 0 Or Index = 1 Then Exit Sub   ' None listed, no can do
' Update the 2 fields, one of which can be edited
txtDescription = lvAttach(Index).SelectedItem
lblData.Caption = lvAttach(Index).SelectedItem.Tag
End Sub

Private Sub lvAttach_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'============================================================
'   Toggle function on pre-save attachment listing
'============================================================

If Index = 0 Then Exit Sub      ' not applicable on saved attachment listing
Dim I As Integer, bWithPath As Boolean, sUpdate As String

' Value of the Tag property:  0=no path displayed, 1=path displayed also
If Val(ColumnHeader.Tag) = 0 Then bWithPath = True
ColumnHeader.Tag = CStr(Abs(CInt(bWithPath)))   ' toggle Tag property to either 0 or 1
' Loop thru each item in the listing and either add or remove the path from the display
For I = 1 To lvAttach(1).ListItems.Count
    sUpdate = StripFile(lvAttach(1).ListItems(I).Text, "N")     ' extract just the filename for now
    ' if the path is to be displayed, then add the path to the filename just extracted, otherwise filename only is good
    If bWithPath Then sUpdate = lvAttach(1).ListItems(I).Tag & sUpdate
    lvAttach(1).ListItems(I).Text = sUpdate     ' update the display
Next
End Sub

Private Sub cmdView_Click(Index As Integer)
'============================================================
'   Option to view or save to file, saved attachments in db
'============================================================

If ListSelected(0) = False Then Exit Sub        ' if no attachment is selected, no can do
Dim sFileName As String, sPath As String, I As Integer
Dim ActualSize As Long

Select Case Index
Case 0:         ' View on screen, find the correct record in the db
    rsAttachment.FindFirst "[ID]=" & Mid(lvAttach(0).SelectedItem.Key, 7)
    GP = sMyPath                        ' set variable to default path
    frmViewAttach.Show 1, Me    ' show view attachment form, changes made there don't affect this form
    GP = Null
Case 1:         ' View via default viewer, find the correct record in the db
    ' if the file is an executable, warn that this action will in effect, activate the executable
    If InStr("|.exe|.bat|.com", "|." & StripFile(lvAttach(0).SelectedItem.Tag, "E")) Then
        I = MsgBox("Viewing an application isn't possible. If you continue, the application will be executed." _
            & vbCrLf & vbCrLf & "Execute the application?", vbYesNo + vbDefaultButton2 + vbExclamation, "Confirmation")
        If I = vbNo Then Exit Sub
    End If
    rsAttachment.FindFirst "[ID]=" & Mid(lvAttach(0).SelectedItem.Key, 7)
    ' call function to identify a unique filename
    sFileName = GetUniqueFileName(StripFile(lvAttach(0).SelectedItem.Tag, "E"), sMyPath)
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
    ExtractAttachment rsAttachment.Fields("Attachment"), sFileName, ActualSize
    ' call function to open the file using registered default extensions. Errors here are reported via that function
    OpenThisFile sFileName, 1, "", frmLibrary.hwnd
Case 2, 3:        ' Save to file permanently, so we need to prompt for a file name
    SaveAttachment Index, False
End Select
UserCnx:
Err.Clear
End Sub

Private Sub SaveAttachment(SaveType As Integer, bDeleteToo As Boolean)
Dim bNoToAll As Boolean, bYesToAll As Boolean, I As Integer, J As Integer
Dim sFile As String, bSave As Boolean, ActualSize As Long

 If chkNoNamePrompt = 1 Then GoSub Prompt4FileName    ' bypass prompting for filenames so get the folder to save to
With frmLibrary.dlgCommon
    For I = lvAttach(0).ListItems.Count To 1 Step -1                ' loop thru the list
        ' if user is saving selected files only the continue, or if user is saving ALL files then continue
        If lvAttach(0).ListItems(I).Selected = True Or SaveType = 3 Then
            If chkNoNamePrompt = 0 Then                         ' prompting for filenames
                GoSub Prompt4FileName                                   ' so get a filename
                sFile = .FileName                                           ' set to filename chosen
            Else                                        ' otherwise, not prompting for filenames, so
                sFile = StripFile(.FileName, "P") & lvAttach(0).ListItems(I).Tag    ' set to original filename of attachment
            End If
            ' if not checking for filenames & not overwriting all files then check to see if file exists
            If chkNoNamePrompt = 1 And bYesToAll = False Then
                    J = Len(Dir(sFile))                                     ' see if file exists
                    If J Then                                                   ' yep, now do another check
                        If bNoToAll = True Then                     ' is user preventing any overwrites?
                            bSave = False                                   '   if so, don't overwrite
                        Else                                                    ' otherwise, user may want to prevent overwrite
                            GP = StripFile(sFile, "N")               ' call save dialog box
                            frmSaveDialog.Show 1, Me
                                If GP = "Cancel" Then Exit Sub  ' if user cancelled then abort sub here
                                ' otherwise, set variables depending on what user selected in dialog box
                                If GP = "No To ALL" Then bNoToAll = True: bYesToAll = False: bSave = False
                                If GP = "No" Then bSave = False
                                If GP = "Yes To ALL" Then bYesToAll = True: bNoToAll = False: bSave = True
                                If GP = "Yes" Then bSave = True
                        End If
                    Else                ' file doesn't exist so saving is OK
                        bSave = True
                    End If
            Else                        ' checking for filenames (overwrite problem resolved with dialog box) or
                bSave = True    ' user opted to overwrite all, therefore save is OK
            End If
            If bSave = True Then    ' if save is OK then save file
                rsAttachment.FindFirst "[ID]=" & Mid(lvAttach(0).ListItems(I).Key, 7)   ' find the one to save
                ' extract the db record to the requested filename
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
                If ExtractAttachment(rsAttachment.Fields("Attachment"), sFile, ActualSize) = False Then
                    MsgBox "Failed to save. Ensure file not in use (if overwriting) or sufficient disk space" _
                        & vbCrLf & vbCrLf & "File: " & StripFile(sFile, "N"), vbInformation + vbOKOnly
                Else
                    If bDeleteToo = True Then
                        lvAttach(0).ListItems.Remove I
                        rsAttachment.Delete
                        bDirty = True
                    End If
                End If
            End If
        End If
    Next
End With
MsgBox "Saving is complete", vbInformation + vbOKOnly
Exit Sub

Prompt4FileName:
    On Error GoTo UserCnx
    With frmLibrary.dlgCommon                                   ' setup the dialog box
        If chkNoNamePrompt = 0 Then                         ' prompt for filename
            .FileName = lvAttach(0).ListItems(I).Tag          ' use original filename as default
            .DefaultExt = StripFile(lvAttach(0).ListItems(I).Tag, "E")  ' use original file extension as default
            .Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
            .DialogTitle = "Save as... Enter file name and location"
        Else
            .FileName = "Open folder to save to"          ' use original filename as default
            .DefaultExt = ""
            .Flags = cdlOFNPathMustExist
            .DialogTitle = "Which folder to save to?"
        End If
        .Filter = "All Files|*.*"
    End With
    frmLibrary.dlgCommon.ShowSave                           ' show the dialog box
Return
UserCnx:
End Sub

Private Sub DestroyTmpFiles()
'============================================================
'   Function simply looks for specific files created by this form & removes them if posssible
'============================================================

Dim I As Integer
On Error Resume Next
fileAll.Path = sMyPath                  ' set file list box path & pattern & see if any files show up
fileAll.Pattern = "~Atch*.*"
For I = fileAll.ListCount - 1 To 0 Step -1  ' for each of those shown...
    ' see if the file is being used by another app (default viewers) & if not, delete it
    If FileInUse(sMyPath & fileAll.List(I)) = False Then Kill sMyPath & fileAll.List(I)
Next
End Sub

Private Sub lvAttach_DblClick(Index As Integer)
If Index = 0 Then
    If lvAttach(Index).ListItems.Count > 0 Then Call cmdView_Click(0)
End If
End Sub
