VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWeb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Pages"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   HelpContextID   =   8
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   7050
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   3495
      MultiSelect     =   2  'Extended
      Pattern         =   "*.url"
      TabIndex        =   8
      Top             =   5490
      Width           =   3345
   End
   Begin VB.CommandButton cmdMassUpdate 
      Caption         =   "Add all selected internet shortcuts"
      Height          =   495
      Left            =   3540
      TabIndex        =   9
      Top             =   7920
      Width           =   3345
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Choose folder above or Click to go to your Default Favorites Folder"
      Height          =   495
      Index           =   1
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7920
      Width           =   3345
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   105
      TabIndex        =   6
      Top             =   5775
      Width           =   3345
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   5475
      Width           =   3345
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Open Web Page"
      Height          =   465
      Index           =   3
      Left            =   105
      TabIndex        =   1
      Top             =   4740
      Width           =   1360
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Edit Selected Item"
      Height          =   465
      Index           =   2
      Left            =   3300
      TabIndex        =   3
      Top             =   4740
      Width           =   1545
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Add New Item Manually"
      Height          =   465
      Index           =   0
      Left            =   4845
      TabIndex        =   4
      Top             =   4740
      Width           =   2085
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Delete Selected Item"
      Height          =   465
      Index           =   1
      Left            =   1470
      TabIndex        =   2
      Top             =   4740
      Width           =   1815
   End
   Begin MSComctlLib.ListView lvURLs 
      Height          =   4185
      Left            =   90
      TabIndex        =   0
      Top             =   525
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7382
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Web Page - Click to Toggle"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Web Address"
         Object.Width           =   5644
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Navigate to folder or click button below"
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
      Height          =   360
      Index           =   2
      Left            =   60
      TabIndex        =   12
      Top             =   5220
      Width           =   3465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select as many as you want added"
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
      Height          =   360
      Index           =   1
      Left            =   3630
      TabIndex        =   11
      Top             =   5220
      Width           =   3315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWeb.frx":0000
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
      Height          =   450
      Index           =   0
      Left            =   150
      TabIndex        =   10
      Top             =   60
      Width           =   6765
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'================================================================
'   form simply displays a listing of saved web addresses and allows them to be activated
'================================================================
Private rsWeb As DAO.Recordset
'   Retrieves a pointer to the ITEMIDLIST structure of a special folder.
Private Declare Function apiSHGetSpecialFolderLocation Lib "shell32" _
    Alias "SHGetSpecialFolderLocation" _
    (ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    ppidl As Long) _
    As Long

'   Converts an item identifier list to a file system path.
Private Declare Function apiSHGetPathFromIDList Lib "shell32" _
    Alias "SHGetPathFromIDList" _
    (pidl As Long, _
    ByVal pszPath As String) _
    As Long

'   Frees a block of task memory previously allocated through a call to
'   the CoTaskMemAlloc or CoTaskMemRealloc function.
Private Declare Sub sapiCoTaskMemFree Lib "ole32" _
    Alias "CoTaskMemFree" _
    (ByVal pv As Long)
    
Private Function GetFavoritesFolder() As String
'================================================================
'   Returns path to a special folder on the machine
'   without a trailing backslash.
'================================================================
Dim lngRet As Long
Dim strLocation As String
Dim pidl As Long
Const MAX_PATH = 260
Const NOERROR = 0
Const lngCSIDL = &H6
    
    '   retrieve a PIDL for the specified location
' Inserted by LaVolpe
On Error GoTo Function_GetFavoritesFolder_General_ErrTrap_by_LaVolpe
    lngRet = apiSHGetSpecialFolderLocation(hWnd, lngCSIDL, pidl)
    If lngRet = NOERROR Then
        strLocation = Space$(MAX_PATH)
        '  convert the pidl to a physical path
        lngRet = apiSHGetPathFromIDList(ByVal pidl, strLocation)
        If Not lngRet = 0 Then
            '   if successful, return the location
            GetFavoritesFolder = Left$(strLocation, _
                                InStr(strLocation, vbNullChar) - 1)
        End If
        '   calling application is responsible for freeing the allocated memory
        '   for pidl when calling SHGetSpecialFolderLocation. We have to
        '   call IMalloc::Release, but to get to IMalloc, a tlb is required.
        '
        '   According to Kraig Brockschmidt in Inside OLE,   CoTaskMemAlloc,
        '   CoTaskMemFree, and CoTaskMemRealloc take the same parameters
        '   as the interface functions and internally call CoGetMalloc, the
        '   appropriate IMalloc function, and then IMalloc::Release.
        Call sapiCoTaskMemFree(pidl)
    End If
Exit Function

Function_GetFavoritesFolder_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Function GetFavoritesFolder]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Function

Private Sub cmdBrowse_Click(Index As Integer)
'================================================================
'   Sub sets the drive, directory & file controls to the location of the system Favorites folder
'================================================================
Dim sPath As String
' Inserted by LaVolpe
On Error GoTo Sub_cmdBrowse_Click_General_ErrTrap_by_LaVolpe
sPath = GetFavoritesFolder                  ' get path location
Drive1.Drive = StripFile(sPath, "D")    ' synchronize the drive control
Dir1.Path = sPath                               ' synchronize the directory control
Exit Sub

Sub_cmdBrowse_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdBrowse_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdMassUpdate_Click()
'================================================================
'   Sub reads all selected internet shortcuts and loads them into the listing
'================================================================
Dim sURL As String, sURLname As String, sSection As String, sKey As String, sPath As String
Dim I As Integer, bUpdate As Boolean
On Error GoTo FailedUpdate
sSection = "InternetShortcut"           ' Shortcut section name in INI file
sKey = "URL"                                 ' Shortcut key name in INI file
sPath = File1.Path                           ' Set the path to the file list control's path
If Left(sPath, 1) <> "\" Then sPath = sPath & "\"   ' include a trailing backslash
' for each selected item in the file list, add it to the main listing
For I = 0 To File1.ListCount - 1
    If File1.Selected(I) = True Then
        sURLname = sPath & File1.List(I)        ' set the full path & file name
        sURL = ReadWriteINI("Get", sURLname, sSection, sKey, "")    ' call function to read the INI URL
        If Len(sURL) Then   ' if it exists then continue
            rsWeb.AddNew
            rsWeb.Fields("URL") = Left(sURL, 255)   ' use only the first 255 characters (may truncate valid web URLs)
            rsWeb.Fields("Description") = Left(StripFile(File1.List(I), "m"), 150) ' use first 150 characters of filename for description
            rsWeb.Update
            File1.Selected(I) = False       ' unselect the added shortcut
            bUpdate = True                    ' set flag to indicate changes made
        End If
    End If
Next
If bUpdate Then LoadURLs            ' call function to populate listing with db records
Exit Sub

FailedUpdate:
MsgBox Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub cmdUpdate_Click(Index As Integer)
'================================================================
'   Sub to let users update the URL or description
'================================================================
' Inserted by LaVolpe
On Error GoTo Sub_cmdUpdate_Click_General_ErrTrap_by_LaVolpe
If lvURLs.SelectedItem Is Nothing And Index Then            ' ensure one is selected in order to update it
    MsgBox "First select an item from the listing.", vbInformation + vbOKOnly
    Exit Sub
End If
Select Case Index
Case 0, 1, 2:   ' Add, delete, edit
    AddEditURL Index + 1        ' call function to add/delete/edit
Case 3:     ' Activate web link
    Dim sRunURL As String
    sRunURL = lvURLs.SelectedItem.SubItems(1)           ' set the command line string
    If OpenThisFile(sRunURL, 1, "", hWnd) = -1 Then     ' send to function for opening based on default app
        MsgBox "Couldn't activate the web link. Ensure it is a valid web address.", vbInformation + vbOKOnly
    End If
End Select
Exit Sub

Sub_cmdUpdate_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdUpdate_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub AddEditURL(iCode As Integer)
'================================================================
'   Function updates, adds or deletes web links from the database and screen
'================================================================
On Error GoTo FailedAdd
' If deleting or updating, gotta find the right record
If iCode - 1 Then rsWeb.FindFirst "[ID]=" & Mid(lvURLs.SelectedItem.Key, 7)
Dim I As Integer
If iCode = 2 Then       ' were deleting here, provide confirmation message
    I = MsgBox("Are you sure you want the selected web site removed from the database?", vbYesNo + vbDefaultButton2 + vbQuestion, "Confirmation")
    If I = vbYes Then rsWeb.Delete Else Exit Sub
Else
    Dim sURL As String, sURLname As String
    If iCode = 3 Then   ' were updating
        sURL = lvURLs.SelectedItem.SubItems(1)      ' get web address
        sURLname = lvURLs.SelectedItem                  ' get web name
    End If                                                                  ' provide input box for changing above values
    sURL = InputBox("Enter the complete web address below for" & vbCrLf & sURLname, "Web Address", sURL)
    If sURL = "" Then Exit Sub  ' user pressed cancel
    sURLname = InputBox("Enter the name/description for the web address of: " & sURL, "Web Address Description", sURLname)
    If sURLname = "" Then Exit Sub  ' user pressed cancel
    With rsWeb                                  ' otherwise, update the db record
        If iCode = 1 Then .AddNew Else .Edit
        .Fields("Description") = Left(sURLname, 150)
        .Fields("URL") = Left(sURL, 255)
        .Update
    End With
End If
rsWeb.Requery
LoadURLs                ' call function to load listing from db records
Exit Sub

FailedAdd:
MsgBox Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub Dir1_Change()
'================================================================
'   syncrhonize file listing with directory control
'================================================================
' Inserted by LaVolpe
On Error GoTo Sub_Dir1_Change_General_ErrTrap_by_LaVolpe
File1.Path = Dir1
Exit Sub

Sub_Dir1_Change_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Dir1_Change]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Drive1_Change()
'================================================================
'   Synchronize directory control with drive control
'================================================================
On Error Resume Next
Dir1.Path = Drive1      ' Change drive
If Err.Number Then      ' errors occur if moving to floppy/CD-ROM and nothing in drive
    Dim LastErr As Long
    LastErr = Err.Number                    ' save error code
    Err.Clear
    On Error GoTo DisplayError
    Err.Raise LastErr                           ' display the error
End If
Exit Sub

DisplayError:
MsgBox Err.Description, vbOKOnly    ' display error - generally floppy not in drive or drive not available
Drive1.Drive = StripFile(Dir1, "D")       ' reset drive control to directory control's path
End Sub

Private Sub Form_Load()
'================================================================
'   Set default values
'================================================================
' Inserted by LaVolpe
On Error GoTo Sub_Form_Load_General_ErrTrap_by_LaVolpe
Icon = frmLibrary.SmallImages.ListImages(3).ExtractIcon                 ' load icon in titlebar
On Error GoTo FailedLoad
Set rsWeb = mainDB.OpenRecordset("tblURLs", dbOpenDynaset)  ' open recordset to edit
LoadURLs                                                        ' call function to load web links from database
Dim I As Integer
    For I = Label1.LBound To Label1.UBound: Label1(I).ForeColor = MyDefaults.LblColorPopup: Next
Call lvURLs_ColumnClick(lvURLs.ColumnHeaders(1))
Exit Sub

FailedLoad:
MsgBox Err.Description, vbExclamation + vbOKOnly
Exit Sub

Sub_Form_Load_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Load]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Resize()
'================================================================
'   Repaint function when form is resized
'================================================================
' Inserted by LaVolpe
On Error GoTo Sub_Form_Resize_General_ErrTrap_by_LaVolpe
DoGradient Me, 2
Exit Sub

Sub_Form_Resize_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Resize]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
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
On Error Resume Next
rsWeb.Close
Set rsWeb = Nothing
End Sub

Private Sub LoadURLs()
'================================================================
'   Load URLs from the database
'================================================================
On Error GoTo AbortLoad
lvURLs.ListItems.Clear
Dim nrURL As Long, itmX As ListItem
If rsWeb.RecordCount > 0 Then       ' loop thru each record and load it into the listbox
    With rsWeb
        .MoveFirst
        Do While .EOF = False
            Set itmX = lvURLs.ListItems.Add(, "RecID:" & .Fields("ID"), .Fields("Description"))
            itmX.SubItems(1) = .Fields("URL")
            .MoveNext
        Loop
    End With
End If
Exit Sub

AbortLoad:
MsgBox Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub lvURLs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'================================================================
'   Toggle function to display one or two columns
'================================================================
' Inserted by LaVolpe
On Error GoTo Sub_lvURLs_ColumnClick_General_ErrTrap_by_LaVolpe
If ColumnHeader.Index = 2 Then Exit Sub
With lvURLs
    If Val(.ColumnHeaders(1).Tag) = "0" Then
        .ColumnHeaders(1).Tag = "1"
        .ColumnHeaders(2).Width = 0
        .ColumnHeaders(1).Width = .Width - 250
    Else
        .ColumnHeaders(1).Tag = "0"
        .ColumnHeaders(1).Width = (.Width - 250) / 2
        .ColumnHeaders(2).Width = .ColumnHeaders(1).Width
    End If
End With
Exit Sub

Sub_lvURLs_ColumnClick_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lvURLs_ColumnClick]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub lvURLs_DblClick()
'================================================================
' Inserted by LaVolpe
On Error GoTo Sub_lvURLs_DblClick_General_ErrTrap_by_LaVolpe
Call cmdUpdate_Click(3)     ' let double clicking act as activating web link
'================================================================
Exit Sub

Sub_lvURLs_DblClick_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lvURLs_DblClick]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub
