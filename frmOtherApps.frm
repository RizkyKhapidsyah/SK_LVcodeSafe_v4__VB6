VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOtherApps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Other Applications"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   HelpContextID   =   9
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   5670
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Delete"
      Height          =   405
      Index           =   2
      Left            =   3570
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Edit"
      Height          =   405
      Index           =   1
      Left            =   2970
      TabIndex        =   3
      Top             =   3480
      Width           =   585
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Add"
      Height          =   405
      Index           =   0
      Left            =   2355
      TabIndex        =   2
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "E&xecute selected Application"
      Height          =   405
      Left            =   105
      TabIndex        =   1
      Top             =   3480
      Width           =   2235
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   405
      Left            =   4425
      TabIndex        =   5
      Top             =   3480
      Width           =   1125
   End
   Begin MSComctlLib.ListView lvApps 
      Height          =   3435
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Application - Click to Toggle"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Location"
         Object.Width           =   4586
      EndProperty
   End
End
Attribute VB_Name = "frmOtherApps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private rsApps As DAO.Recordset, bDirty As Boolean

Private Sub cmdClose_Click()
' Inserted by LaVolpe
On Error GoTo Sub_cmdClose_Click_General_ErrTrap_by_LaVolpe
Unload Me
Exit Sub

Sub_cmdClose_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdClose_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdExecute_Click()
' Inserted by LaVolpe
On Error GoTo Sub_cmdExecute_Click_General_ErrTrap_by_LaVolpe
If lvApps.ListItems.Count = 0 Or lvApps.SelectedItem Is Nothing Then
    MsgBox "First select an item from the listing.", vbInformation + vbOKOnly
    Exit Sub
End If
OpenThisFile lvApps.SelectedItem.SubItems(1), 1, "", hWnd
Exit Sub

Sub_cmdExecute_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdExecute_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdUpdate_Click(Index As Integer)
On Error GoTo FailedAction
If Index > 0 Then
    If lvApps.ListItems.Count = 0 Or lvApps.SelectedItem Is Nothing Then
        MsgBox "First select an item from the listing.", vbInformation + vbOKOnly
        Exit Sub
    Else
        rsApps.FindFirst "[ID]=" & Mid(lvApps.SelectedItem.Key, 7)
        If rsApps.NoMatch = True Then
            MsgBox "Can't edit/delete that item. It is no longer in the database", vbInformation + vbOKOnly
            Exit Sub
        End If
    End If
End If
Dim sApp As String, sExe As String, i As Integer, bReload As Boolean, itmX As ListItem
If Index = 2 Then   ' we're deleting
    If MsgBox("Are you sure you want the following applicaton removed from the database?" & vbCrLf & vbCrLf _
        & sApp, vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation") = vbNo Then Exit Sub
    rsApps.Delete
    rsApps.Requery
    bReload = True
Else
    i = vbYes
    If Index = 1 Then   ' we're editing
        i = MsgBox("Do you want to change the location of the application?", vbYesNo + vbQuestion)
        sApp = lvApps.SelectedItem
        sExe = lvApps.SelectedItem.SubItems(1)
    End If
    If i = vbYes Then
        On Error GoTo UserCnx
        With frmLibrary.dlgCommon
            .DialogTitle = "Where is application you want " & Choose(Index + 1, "Added", "to Update")
            .Filter = "Applications|*.exe;*.com;*.bat|All Files|*.*"
            .FilterIndex = 0
            .Flags = cdlOFNFileMustExist
        End With
        frmLibrary.dlgCommon.ShowOpen
        sExe = frmLibrary.dlgCommon.FileName
        If Len(sExe) > 255 Then sExe = GetShortPathName(sExe)
        If Len(sExe) > 255 Then
            MsgBox "Can't add that to the list. The filename and path exceed 255 characters. Sorry", vbInformation + vbOKOnly
            Exit Sub
        End If
    End If
    On Error GoTo FailedAction
    sApp = InputBox("Provide the name of this application", "Application Title", Choose(Index + 1, StripFile(sExe, "m"), sApp))
    If sApp = "" Then Exit Sub
    sApp = Left(sApp, 50)
    With rsApps
        If Index = 0 Then .AddNew Else .Edit
        .Fields("AppName") = sApp
        .Fields("AppExe") = sExe
        .Update
        bReload = True
    End With
End If
If bReload = True Then
    bDirty = True
    LoadApps
End If
Exit Sub

FailedAction:
MsgBox Err.Description, vbExclamation + vbOKOnly
UserCnx:
End Sub

Private Sub Form_Load()
On Error GoTo FailedLoad
Icon = frmLibrary.SmallImages.ListImages(16).ExtractIcon
DoGradient Me, 2
Set rsApps = mainDB.OpenRecordset("tblApplications", dbOpenDynaset)
LoadApps
Call lvApps_ColumnClick(lvApps.ColumnHeaders(1))
bDirty = False
Exit Sub

FailedLoad:
MsgBox Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub LoadApps()
On Error GoTo FailedLoad
lvApps.ListItems.Clear
If rsApps.RecordCount = 0 Then Exit Sub
Dim itmX As ListItem
With rsApps
    .MoveFirst
    Do While .EOF = False
        Set itmX = lvApps.ListItems.Add(, "RecID:" & .Fields("ID"), .Fields("AppName"))
        itmX.SubItems(1) = .Fields("AppExe")
        .MoveNext
    Loop
End With
Set itmX = Nothing
Exit Sub
FailedLoad:
MsgBox Err.Description, vbExclamation + vbOKOnly
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
On Error Resume Next
rsApps.Close
Set rsApps = Nothing
If bDirty = True Then RefreshOtherApps
End Sub

Private Sub lvApps_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' Inserted by LaVolpe
On Error GoTo Sub_lvApps_ColumnClick_General_ErrTrap_by_LaVolpe
If ColumnHeader.Index > 1 Then Exit Sub
With lvApps.ColumnHeaders
    If Val(.Item(1).Tag) = 0 Then
        .Item(2).Width = 0
        .Item(1).Tag = "1"
        .Item(1).Width = lvApps.Width - 250
    Else
        .Item(1).Tag = "0"
        .Item(1).Width = (lvApps.Width - 250) / 2
        .Item(2).Width = .Item(1).Width
    End If
End With
Exit Sub

Sub_lvApps_ColumnClick_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lvApps_ColumnClick]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub lvApps_DblClick()
' Inserted by LaVolpe
On Error GoTo Sub_lvApps_DblClick_General_ErrTrap_by_LaVolpe
Call cmdExecute_Click
Exit Sub

Sub_lvApps_DblClick_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lvApps_DblClick]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub
