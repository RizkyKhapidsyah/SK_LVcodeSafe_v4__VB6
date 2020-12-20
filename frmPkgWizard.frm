VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPkgWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import/Export Package Wizard"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   HelpContextID   =   13
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmWiz 
      Caption         =   "Import Wizard - Step 1 - Choosing which code to import..."
      Enabled         =   0   'False
      Height          =   3495
      Index           =   3
      Left            =   105
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   7065
      Begin VB.CheckBox chkNoDups 
         Caption         =   "Check to unselect duplicate code items"
         Height          =   240
         Left            =   3675
         TabIndex        =   40
         Top             =   2745
         Width           =   3240
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Toggle to select/unselect all of the above"
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   37
         Top             =   2760
         Width           =   3360
      End
      Begin VB.Label Label1 
         Caption         =   $"frmPkgWizard.frx":0000
         Height          =   420
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   3000
         Width           =   6765
      End
   End
   Begin VB.Frame frmWiz 
      Caption         =   "Export Wizard - Step 1 - Choose which code to export"
      Enabled         =   0   'False
      Height          =   3495
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   7065
      Begin VB.ListBox lstCodeImport 
         Height          =   2205
         Left            =   180
         TabIndex        =   27
         Top             =   480
         Width           =   2385
      End
      Begin VB.CommandButton cmdAddRemove 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2610
         TabIndex        =   4
         ToolTipText     =   "Remove Selected Code"
         Top             =   1395
         Width           =   315
      End
      Begin VB.CommandButton cmdAddRemove 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2610
         TabIndex        =   3
         ToolTipText     =   "Add selected code"
         Top             =   735
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "TIP: Hold the control or shift key down to select multiple items."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   35
         Top             =   3195
         Width           =   6735
      End
      Begin VB.Label Label1 
         Caption         =   "Select Code from listing >>"
         Height          =   240
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label Label1 
         Caption         =   $"frmPkgWizard.frx":00AE
         Height          =   450
         Index           =   1
         Left            =   195
         TabIndex        =   5
         Top             =   2775
         Width           =   6765
      End
   End
   Begin VB.Frame frmWiz 
      Caption         =   "Import Wizard - Final Step"
      Enabled         =   0   'False
      Height          =   3495
      Index           =   6
      Left            =   105
      TabIndex        =   14
      Tag             =   "Export Wizard - Final Step"
      Top             =   120
      Visible         =   0   'False
      Width           =   7065
      Begin VB.CheckBox chkDupURLs 
         Caption         =   $"frmPkgWizard.frx":0139
         Height          =   390
         Left            =   150
         TabIndex        =   39
         Top             =   3060
         Width           =   6795
      End
      Begin MSComctlLib.ListView lvConfirm 
         Height          =   2385
         Left            =   120
         TabIndex        =   38
         Top             =   675
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   4207
         View            =   3
         LabelEdit       =   1
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
            Text            =   "Code to Import/Export"
            Object.Width           =   5909
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Web Sites to Import/Export"
            Object.Width           =   5909
         EndProperty
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   375
         Left            =   5340
         TabIndex        =   16
         Tag             =   "Export"
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   $"frmPkgWizard.frx":01D6
         Height          =   420
         Index           =   7
         Left            =   180
         TabIndex        =   15
         Tag             =   "Click EXPORT to begin the export process"
         Top             =   270
         Width           =   5145
      End
   End
   Begin VB.Frame frmWiz 
      Caption         =   "Import/Export Wizard"
      Height          =   3495
      Index           =   0
      Left            =   105
      TabIndex        =   22
      Tag             =   "Export Wizard - Final Step"
      Top             =   120
      Width           =   7065
      Begin VB.CommandButton cmdOption 
         Caption         =   "I want to Import"
         Height          =   465
         Index           =   1
         Left            =   2085
         TabIndex        =   26
         Top             =   2760
         Width           =   2475
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "I want to Export"
         Height          =   465
         Index           =   0
         Left            =   2085
         TabIndex        =   25
         Top             =   1485
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "The wizard will also import the export file created by another LaVolpe Code Safe application."
         Height          =   675
         Index           =   10
         Left            =   105
         TabIndex        =   24
         Tag             =   "Click EXPORT to begin the export process"
         Top             =   2340
         Width           =   6765
      End
      Begin VB.Label Label1 
         Caption         =   $"frmPkgWizard.frx":0263
         Height          =   675
         Index           =   9
         Left            =   150
         TabIndex        =   23
         Tag             =   "Click EXPORT to begin the export process"
         Top             =   780
         Width           =   6765
      End
   End
   Begin MSComctlLib.ListView lvMaster 
      Height          =   2370
      Left            =   3045
      TabIndex        =   7
      Tag             =   "00285067950304504035"
      Top             =   480
      Visible         =   0   'False
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   4180
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Purpose"
         Object.Width           =   5080
      EndProperty
   End
   Begin VB.CommandButton cmdStep 
      Cancel          =   -1  'True
      Caption         =   "Cancel && Close"
      Height          =   435
      Index           =   2
      Left            =   4650
      TabIndex        =   21
      Top             =   3630
      Width           =   2505
   End
   Begin VB.CommandButton cmdStep 
      Caption         =   "Next Step"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   435
      Index           =   1
      Left            =   2160
      TabIndex        =   20
      Top             =   3630
      Width           =   2490
   End
   Begin VB.CommandButton cmdStep 
      Caption         =   "Previous Step"
      Enabled         =   0   'False
      Height          =   435
      Index           =   0
      Left            =   105
      TabIndex        =   19
      Top             =   3630
      Width           =   2055
   End
   Begin MSComctlLib.ListView lvURLs 
      Height          =   2400
      Left            =   210
      TabIndex        =   18
      Tag             =   "285|6795"
      Top             =   450
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
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
         Text            =   "URL"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Frame frmWiz 
      Caption         =   "Import Wizard - Step 3 - Choose which Web Sites to Import"
      Enabled         =   0   'False
      Height          =   3495
      Index           =   5
      Left            =   105
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   7065
      Begin VB.CheckBox chkAll 
         Caption         =   "Toggle to select/unselect all of the above"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   36
         Top             =   2730
         Width           =   4020
      End
      Begin VB.Label Label1 
         Caption         =   "Select any websites listed above, if any, to import into this database. Click NEXT to finish."
         Height          =   420
         Index           =   6
         Left            =   135
         TabIndex        =   13
         Top             =   3015
         Width           =   6765
      End
   End
   Begin VB.Frame frmWiz 
      Caption         =   "Import Wizard - Step 2 - Resolve category and language names not in this database"
      Enabled         =   0   'False
      Height          =   3495
      Index           =   4
      Left            =   105
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   7065
      Begin VB.CommandButton cmdResolve 
         Caption         =   "Set"
         Height          =   300
         Left            =   5280
         TabIndex        =   32
         Top             =   2670
         Width           =   1500
      End
      Begin VB.ComboBox cboResolve 
         Height          =   315
         ItemData        =   "frmPkgWizard.frx":032C
         Left            =   1095
         List            =   "frmPkgWizard.frx":0339
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2655
         Width           =   4155
      End
      Begin VB.OptionButton optResolve 
         Caption         =   "Languages that Need Resolving"
         Height          =   330
         Index           =   1
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   270
         Width           =   3390
      End
      Begin VB.OptionButton optResolve 
         Caption         =   "Categories that Need Resolving"
         Height          =   330
         Index           =   0
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   270
         Value           =   -1  'True
         Width           =   3420
      End
      Begin MSComctlLib.ListView lvResolve 
         Height          =   2010
         Left            =   135
         TabIndex        =   28
         Top             =   615
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   3545
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Category"
            Object.Width           =   5539
         EndProperty
      End
      Begin VB.ListBox lstCats 
         Height          =   2010
         ItemData        =   "frmPkgWizard.frx":03A5
         Left            =   3630
         List            =   "frmPkgWizard.frx":03A7
         TabIndex        =   11
         Top             =   615
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "ACTION >>"
         Height          =   270
         Index           =   4
         Left            =   135
         TabIndex        =   33
         Top             =   2685
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   $"frmPkgWizard.frx":03A9
         Height          =   420
         Index           =   3
         Left            =   105
         TabIndex        =   10
         Top             =   3030
         Width           =   6765
      End
   End
   Begin VB.Frame frmWiz 
      Caption         =   "Export Wizard - Step 2 - Choose which websites to import"
      Enabled         =   0   'False
      Height          =   3495
      Index           =   2
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   7065
      Begin VB.CheckBox chkAll 
         Caption         =   "Toggle to select/unselect all of the above"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   34
         Top             =   2745
         Width           =   4020
      End
      Begin VB.Label Label1 
         Caption         =   "Highlight which web sites you want exported and click NEXT to finish."
         Height          =   345
         Index           =   8
         Left            =   105
         TabIndex        =   17
         Top             =   3075
         Width           =   6765
      End
   End
End
Attribute VB_Name = "frmPkgWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iStep As Integer            ' Current step in the Wizard
Private iImEx As Integer            ' Import or Export flag 0=Export, 1=Import
Private sFileName As String, sFileNameBinary As String

Private Sub chkAll_Click(Index As Integer)
' Check boxes used to select all shown URLs and/or Code List items
Dim I As Integer

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto chkAll_Click_General_ErrTrap

If Index < 2 Then       ' URLs
    For I = 1 To lvURLs.ListItems.Count
        lvURLs.ListItems(I).Selected = CBool(chkAll(Index).Value)
    Next
Else                    ' Code
    For I = 1 To lvMaster.ListItems.Count
        lvMaster.ListItems(I).Selected = CBool(chkAll(Index).Value)
    Next
End If
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
chkAll_Click_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: chkAll_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub chkNoDups_Click()
Dim I As Integer

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto chkNoDups_Click_General_ErrTrap

For I = 1 To lvMaster.ListItems.Count
    If lvMaster.ListItems(I).ForeColor = 255 Then lvMaster.ListItems(I).Selected = False
Next
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
chkNoDups_Click_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: chkNoDups_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub cmdAddRemove_Click(Index As Integer)
' Adds/removes code from the listing of Code to be Imported
Dim I As Long, J As Long, itmX As ListItem

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto cmdAddRemove_Click_General_ErrTrap

If Index Then    ' remove
    For I = lstCodeImport.ListCount - 1 To 0 Step -1
       If lstCodeImport.Selected(I) Then lstCodeImport.RemoveItem I
    Next
Else             ' add
    For J = 1 To lvMaster.ListItems.Count
        If lvMaster.ListItems(J).Selected = True Then
            Set itmX = lvMaster.ListItems(J)
            For I = 0 To lstCodeImport.ListCount - 1
                If lstCodeImport.ItemData(I) = Val(Mid(itmX.Key, 5)) Then Exit For
            Next
            If I = lstCodeImport.ListCount Then
                lstCodeImport.AddItem itmX.Text
                lstCodeImport.ItemData(lstCodeImport.ListCount - 1) = Val(Mid(itmX.Key, 5))
            End If
        End If
    Next
End If
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
cmdAddRemove_Click_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: cmdAddRemove_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub cmdImport_Click()
' The actual command to import or export

' ensure something was selected to import/export

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto cmdImport_Click_General_ErrTrap

If lvConfirm.ListItems.Count = 0 Then
    MsgBox "You haven't selected any code or web sites to " & Choose(iImEx + 1, "export.", "import."), vbInformation + vbOKOnly
    Exit Sub
End If

GP = ""                 ' flag for main form when importing data
If iImEx = 0 Then
    If ExportData = True Then        ' routine that exports data
        ' update the data filename in the index & provide success message
        If lstCodeImport.ListCount = 0 Then
            sFileNameBinary = "None"
            MsgBox "One .CSX file was created." & vbCrLf & _
                "Do not modify this file at all until it is imported.", vbInformation + vbOKOnly, "Export Complete"
        Else
            MsgBox "Two .CSX files were created. Both are needed in order to import into another" & vbCrLf & _
                "LaVolpe Code Safe database.  Do not modify those files at all until they are imported.", vbInformation + vbOKOnly, "Export Complete"
        End If
        ReadWriteINI "Write", sFileName, "Data Files", "File1", sFileNameBinary
    End If
Else
    ' Import function
    ' First we ensure each language and/or category that needed to be resolved was resolved
    Dim lResolved(0 To 1) As Long, sVal As String, lTotals(0 To 1) As Long
    sVal = ReadWriteINI("Get", sFileName, "tblCategories", "Resolved", "0")
    lResolved(0) = Val(sVal)
    sVal = ReadWriteINI("Get", sFileName, "tblCategories", "AutoResolved", "0")
    lResolved(0) = lResolved(0) & Val(sVal)
    sVal = ReadWriteINI("Get", sFileName, "tblCategories", "Number", "0")
    lTotals(0) = Val(sVal)
    '   Above checked for category resolution & below checks for Language
    sVal = ReadWriteINI("Get", sFileName, "tblLanguage", "Resolved", "0")
    lResolved(1) = Val(sVal)
    sVal = ReadWriteINI("Get", sFileName, "tblLanguage", "AutoResolved", "0")
    lResolved(1) = lResolved(1) & Val(sVal)
    sVal = ReadWriteINI("Get", sFileName, "tblLanguage", "Number", "0")
    lTotals(1) = Val(sVal)
    ' If additional items still require resolution, then prompt & go there
    If lResolved(0) < lTotals(0) Then
        MsgBox "Returning to Step 2 of the Wizard. There is at least one Category that needs to be resolved.", vbInformation + vbOKOnly, "Unresolved Categories"
        optResolve(0).Value = True
        ShowStep -1
        ShowStep -1
        Exit Sub
    Else
        If lResolved(1) < lTotals(1) Then
            MsgBox "Returning to Step 2 of the Wizard. There is at least one Language that needs to be resolved.", vbInformation + vbOKOnly, "Unresolved Languages"
            optResolve(1).Value = True
            ShowStep -1
            ShowStep -1
            Exit Sub
        End If
    End If
    ImportExports           ' function that performs the imports
    CleanUpIndex            ' removes temporary settings from the index file
    RefreshCategories       ' refresh library Categories
    RefreshLanguages        ' refresh library Languages
    MsgBox "Import Complete", vbInformation + vbOKOnly
    GP = "Requery"           ' flag to indicate code listing should be requeried
    Dim I As Integer         ' if user had the Web form open, then requery that form also
    For I = 0 To Forms.Count - 1
        If Forms(I).Name = "frmWeb" Then
            GP = "Requery with Web"
            Exit For
        End If
    Next
End If
Unload Me
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
cmdImport_Click_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: cmdImport_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub cmdOption_Click(Index As Integer)
' Option button to begin the Import/Export Wizard


' Inserted by LaVolpe OnError Insertion Program.
On Error Goto cmdOption_Click_General_ErrTrap

iImEx = Index       ' set import or export function flag
chkAll(0) = 0: chkAll(1) = 0: chkAll(2) = 0     ' reset all check boxes
chkDupURLs.Visible = iImEx                      ' show checkbox on final pane only when importing
chkDupURLs.Enabled = iImEx
SetFileName                                     ' prompt for import filename or export location
If Len(sFileName) = 0 Then Exit Sub
Select Case Index
Case 0  ' Exporting
    lvMaster.Left = Val(Mid(lvMaster.Tag, 11, 5))   ' adjust Code Listing dimensions
    lvMaster.Width = Val(Right(lvMaster.Tag, 5))    ' and set up screen for Exporting
    Label1(7).Caption = Replace(Label1(7).Caption, "IMPORT", "EXPORT")
    cmdImport.Caption = "EXPORT"
    frmWiz(frmWiz.UBound).Caption = "Export Wizard - Final Step"
    PopulateLists "tblSourceCode"
    PopulateLists "tblURLs"
Case 1  ' Importing
    optResolve(0).Value = False                     ' reset resolution option buttons
    optResolve(1).Value = False: optResolve(1).Tag = ""
    lvMaster.Left = Val(Left(lvMaster.Tag, 5))      ' adjust Code Listing dimensions
    lvMaster.Width = Val(Mid(lvMaster.Tag, 6, 5))   ' and set up screen for Importing
    Label1(7).Caption = Replace(Label1(7).Caption, "EXPORT", "IMPORT")
    cmdImport.Caption = "IMPORT"
    frmWiz(frmWiz.UBound).Caption = "Import Wizard - Final Step"
    chkDupURLs.Value = 1
    ReadExportData       ' read code listing & web links that can be imported
End Select
' initialize the Step Buttons & current step number
cmdStep(0).Enabled = True
cmdStep(1).Enabled = True
iStep = 0
ShowStep 1      ' show step 1
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
cmdOption_Click_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: cmdOption_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub cmdResolve_Click()

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto cmdResolve_Click_General_ErrTrap

If lvResolve.ListItems.Count = 0 Then Exit Sub
' this routine sets what happens to categories/languages being imported but do not
'   exist in the current database.  They need to be resolved by identifying them as
'   either New, not to be loaded, or changed to an existing language/category

Dim sSection As String
sSection = Mid(lvResolve.ColumnHeaders(1).Text, 8)      ' index section to read/write
If lvResolve.SelectedItem Is Nothing Then
    ' Ensure user selected an item to resolve
    MsgBox "Ensure an " & lvResolve.ColumnHeaders(1).Text & " is selected.", vbInformation + vbOKOnly
    Exit Sub
Else
    ' If changing to an existing item, ensure new & existing items are selected
    If cboResolve.ListIndex = 2 And lstCats.ListIndex < 0 Then
        MsgBox "Ensure a database " & sSection & " is selected from the right-hand listing.", vbInformation + vbOKOnly
        Exit Sub
    End If
End If

Dim sItem As String, sValue As String
Dim sActionXtra As String, sAction As String, iIcon As Integer
Dim lRec As Long, lStart As Long, lStop As Long

' Set index section & then reset the current setting for the seleced item to be resolved
sSection = "tbl" & Replace(sSection, "y", "ies")
ReadWriteINI "Write", sFileName, sSection, "XRef" & Mid(lvResolve.SelectedItem.Key, 5), ""
Select Case cboResolve.ListIndex
Case 0: ' New Category/Language
    sAction = "0": sActionXtra = " {NEW}": iIcon = 30
Case 1: ' Delete Category/Language
    sAction = "1": sActionXtra = " {Don't Add}": iIcon = 29
Case 2: ' Change
    sAction = "X" & lstCats.ListIndex: iIcon = 31
    sActionXtra = " {chg to " & lstCats.List(lstCats.ListIndex) & "}"
    ReadWriteINI "Write", sFileName, sSection, "XRef" & Mid(lvResolve.SelectedItem.Key, 5), CStr(lstCats.ItemData(lstCats.ListIndex))
Case 3: ' ALL are new
    lStart = 1
    lStop = lvResolve.ListItems.Count
    sAction = "0": sActionXtra = " {NEW}": iIcon = 30
Case 4: ' Revert back to original value
    sItem = Mid(lvResolve.ColumnHeaders(1).Text, 8) & Mid(lvResolve.SelectedItem.Key, 5)
    GoSub ReadLog                           ' read value from index
    lvResolve.SelectedItem.Text = sValue    ' reset to that value
    lvResolve.SelectedItem.Tag = sValue
    lvResolve.SelectedItem.SmallIcon = 0    ' flag indicating resolution required
    sItem = "Action" & Mid(lvResolve.SelectedItem.Key, 5): sValue = ""
    GoSub WriteToLog                        ' reset current pending action
    lStart = 1                              ' prevent Loop below from occurring
End Select
If lStart = 0 Then                          ' initially loop for current item only
    lStart = lvResolve.SelectedItem.Index
    lStop = lStart
End If
For lRec = lStart To lStop
    ' write the resolve action to the index
    sItem = "Action" & Mid(lvResolve.ListItems(lRec).Key, 5): sValue = sAction: GoSub WriteToLog
    ' Update display with confirmation of the action requested & set icon appropriately
    lvResolve.ListItems(lRec).Text = lvResolve.ListItems(lRec).Tag & sActionXtra
    lvResolve.ListItems(lRec).SmallIcon = iIcon
    If iIcon = 30 Then      ' New objects
        ' here we save the new name to the index vs writing over it, this way we can revert back if needed
        sItem = "NewItem" & Mid(lvResolve.ListItems(lRec).Key, 5)
        sValue = lvResolve.ListItems(lRec).Tag: GoSub WriteToLog
    End If
Next
' Count how many shown objects have pending resolve actions & save that number to the index
lRec = 0
For lStart = 1 To lvResolve.ListItems.Count
    If Not IsEmpty(lvResolve.ListItems(lStart)) Then lRec = lRec + 1
Next
sItem = "Resolved": sValue = lRec: GoSub WriteToLog
Exit Sub

ReadLog:
sValue = ReadWriteINI("Get", sFileName, sSection, sItem, "")
Return
WriteToLog:
sValue = ReadWriteINI("Write", sFileName, sSection, sItem, sValue)
Return
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
cmdResolve_Click_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: cmdResolve_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub cmdStep_Click(Index As Integer)
' buttons to goto the Next or Previous steps

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto cmdStep_Click_General_ErrTrap

If Index = 2 Then
    Unload Me       ' Cancel button
Else
    ShowStep Index * 2 - 1      ' -1 for previous step, 1 for next step
End If
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
cmdStep_Click_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: cmdStep_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub ShowStep(iIncr As Integer)
' Function displays the current step & performs other miscellaneous functions
' depending on the current step

Dim I As Integer, iOffset As Integer


' Inserted by LaVolpe OnError Insertion Program.
On Error Goto ShowStep_General_ErrTrap

iStep = iStep + iIncr           ' set current step number
' The IMPORT wizard has 3 screens & then the final screen; 1st screen is screen #1
' The EXPORT wizard has 2 screens & then the final screen; 1st screen is screen #3

' So we just adjust the acutal screen reference here, and it depends whether or not the
'   user is going forward or backwards.
If iImEx = 0 And ((iStep = 3 And iIncr > 0)) Then iOffset = 3
If iImEx = 1 And iStep > 0 Then iOffset = (iImEx * 2)
' display the current screen
frmWiz(iStep + iOffset).Visible = True
frmWiz(iStep + iOffset).Enabled = True

' this listview is not bound to a tab since it used on 2 tabs. So display it only when
'   user is on step 1
lvMaster.Visible = (iStep = 1)
' this listview is not bound to a tab since it used on 2 tabs. So display it only when
'   user is on step 2 when Exporting, or on step 3 when Importing
lvURLs.Visible = (iStep = 2 And iImEx = 0) Or (iStep = 3 And iImEx = 1)
' Enable the above listviews dependent upon their visible property & bring them to the top
lvMaster.Enabled = lvMaster.Visible
lvURLs.Enabled = lvURLs.Visible
If lvMaster.Visible Then lvMaster.ZOrder
If lvURLs.Visible Then lvURLs.ZOrder

' Now for each screen/tab not displayed, ensure it is hidden & disabled
For I = frmWiz.LBound To frmWiz.UBound
    If I <> iStep + iOffset Then
        frmWiz(I).Visible = False
        frmWiz(I).Enabled = False
    End If
Next
' turn the Next button off if at the beginning or the very end
cmdStep(1).Enabled = (frmWiz(frmWiz.UBound).Visible = False) And iStep > 0
' turn the Previous button off if at the very beginning
cmdStep(0).Enabled = (iStep > 0)

' Now when importing and on step 2, we trigger the Resolve functions if it hasn't been triggered yet
If iImEx = 1 And iStep = 2 Then
    If optResolve(1).Tag = "" Then      ' not triggered yet, so we trigger it now
        ' by setting each button to true, it will determine whether or not the
        '   categories/languages need to be resolved
        optResolve(0).Enabled = True: optResolve(1).Enabled = True
        optResolve(1).Value = True
        optResolve(0).Value = True
        optResolve(1).Tag = "Done"      ' flag to prevent triggering functions again
    End If
    ' if nothing needs to be resolved then skip the step
    If optResolve(0).Enabled = False And optResolve(1).Enabled = False Then
        If iIncr = 1 Then   ' going forward - if going backwards, message was already seen
            If optResolve(1).Tag = "Done" Then  ' has this message been displayed before?
                ' if not, then we display it once & ignore it thereafter
                optResolve(1).Tag = "Done and shown Message"
                If MsgBox("No categories or languages need to be resolved.", vbInformation + vbOKCancel) = vbCancel Then Exit Sub
            End If
            ShowStep iIncr      ' skip to the next step
        Else
            ShowStep iIncr      ' skip to the previous step
        End If
    End If
End If

' Only when on the final step, we are going to display the confirmation listview
If frmWiz(frmWiz.UBound).Visible = True Then
    Dim itmX As ListItem, lRec As Long
    lvConfirm.ListItems.Clear
    If iImEx = 0 Then       ' Exporting
        ' fill in the view with each Code Item to be exported
        For I = 0 To lstCodeImport.ListCount - 1
            Set itmX = lvConfirm.ListItems.Add(, , lstCodeImport.List(I))
            If IsEmpty(lvMaster.ListItems("Rec#" & lstCodeImport.ItemData(I)).SmallIcon) = False Then itmX.SmallIcon = 13
        Next
    Else
        ' fill in the view with each Code Item to be imported
        For I = 1 To lvMaster.ListItems.Count
            If lvMaster.ListItems(I).Selected Then
                Set itmX = lvConfirm.ListItems.Add(, , lvMaster.ListItems(I).Text, , lvMaster.ListItems(I).SmallIcon)
                itmX.Tag = "Dup"
            End If
        Next
    End If
    ' now we add the web links to be imported/exported
    lRec = 0
    For I = 1 To lvURLs.ListItems.Count
        If lvURLs.ListItems(I).Selected = True Then
            lRec = lRec + 1
            If lRec > lvConfirm.ListItems.Count Then
                Set itmX = lvConfirm.ListItems.Add(, , "")
            Else
                Set itmX = lvConfirm.ListItems(lRec)
            End If
            itmX.SubItems(1) = lvURLs.ListItems(I).Text
        End If
    Next
    lvConfirm.Refresh
    Set itmX = Nothing
End If
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
ShowStep_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: ShowStep" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub PopulateLists(sTable As String)
' This function basically & simply populates listviews & listboxes with information
'   from the current database

Dim rstImport As DAO.Recordset, sSQL As String, lvTemp As Object, itmX As ListItem
Dim lFields As Integer, sFieldValue As String
Dim bListBox As Boolean, sTmp As String
Dim sSection As String, sItem As String, sValue As String, lRecs(0 To 1) As Long
Dim iSection As Integer


' Inserted by LaVolpe OnError Insertion Program.
On Error Goto PopulateLists_General_ErrTrap

Select Case sTable
Case "tblSourceCode"
    ' Display current code items from the database
    sSQL = "Select IDnr, CodeName, Purpose From " & sTable & " Order by CodeName;"
    Set rstImport = mainDB.OpenRecordset(sSQL, dbOpenDynaset)
    Set lvTemp = lvMaster
    GoSub PopulateNow
    ' now flag each one that has an attachment with the attachment icon
    sSQL = "Select Distinct RecIDRef From tblAttachments"
    Set rstImport = mainDB.OpenRecordset(sSQL, dbOpenDynaset)
    If rstImport.RecordCount Then
        With rstImport
            .MoveFirst
            Do While .EOF = False
                On Error Resume Next
                Set itmX = Nothing
                Set itmX = lvTemp.ListItems("Rec#" & .Fields(0))
                If Not itmX Is Nothing Then itmX.SmallIcon = 13
            .MoveNext
            Loop
        End With
    End If
    rstImport.Close
Case "tblURLs"
    ' Display current web links from the database
    sSQL = "Select ID, URL, Description From " & sTable & " Order by URL;"
    Set rstImport = mainDB.OpenRecordset(sSQL, dbOpenDynaset)
    Set lvTemp = lvURLs
    GoSub PopulateNow
Case "tblCategories"
    ' Display listing of current Categories from the database
    sSQL = "Select ID, Category From " & sTable & " Order by Category;"
    Set rstImport = mainDB.OpenRecordset(sSQL, dbOpenDynaset)
    Set lvTemp = lstCats
    bListBox = True     ' this is a ListBox control vs a ListView
    GoSub PopulateNow
    With cboResolve     ' change references in the Action's combo to reflect Category
        .Clear
        .AddItem "Make this a brand new Category"
        .AddItem "Don't add this category to the database"
        .AddItem "Change category to highlighted entry on right"
        .AddItem "Make ALL brand new Categories"
        .AddItem "Revert back to original value"
        .ListIndex = 3
    End With
    lvResolve.ColumnHeaders(1).Text = "Import Category"
Case "tblLanguage"
    ' Display listing of current Languages from the database
    sSQL = "Select ID, Language From " & sTable & " Order by Language;"
    Set rstImport = mainDB.OpenRecordset(sSQL, dbOpenDynaset)
    Set lvTemp = lstCats
    bListBox = True         ' this is a listbox control vs a listview
    GoSub PopulateNow
    With cboResolve         ' change references in the Action's combo to reflect Language
        .Clear
        .AddItem "Make this a brand new Language"
        .AddItem "Don't add this language to the database"
        .AddItem "Change language to highlighted entry on right"
        .AddItem "Make ALL brand new Languages"
        .AddItem "Revert back to original value"
        .ListIndex = 3
    End With
    lvResolve.ColumnHeaders(1).Text = "Import Language"
Case "ImportCategories", "ImportLanguages"
    ' Display listing of categories that will be imported with the code, if any
    If sTable = "ImportCategories" Then
        sSection = "tblCategories": sItem = "Number": GoSub ReadLog
    Else
        sSection = "tblLanguage": sItem = "Number": GoSub ReadLog
    End If
    If sTable = "ImportCategories" Then iSection = 1 Else iSection = 2
    lRecs(0) = 1
    lvResolve.ListItems.Clear
    ' loop thru each Index category/language & add it to the lvResolve listview
    sItem = "ID" & lRecs(0): GoSub ReadLog
    Do While Len(sValue)
        lRecs(1) = Val(sValue)
        sItem = Choose(iSection, "Category", "Language") & lRecs(1): GoSub ReadLog
        sTmp = sValue
        GoSub LoadExportedCatLangData
        lRecs(0) = lRecs(0) + 1
        sItem = "ID" & lRecs(0): GoSub ReadLog
    Loop
    GoSub ResolveNow        ' auto-resolve exact matches between export & import databases
    lvResolve.Refresh
End Select
On Error Resume Next
Set rstImport = Nothing
Set lvTemp = Nothing
Set itmX = Nothing
Exit Sub

ReadLog:
sValue = ReadWriteINI("Get", sFileName, sSection, sItem, "")
Return

ResolveNow:
' Routine tries to auto-resolve languages/categories that exist in both import/export dbs
lRecs(1) = 0
For lRecs(0) = 0 To lstCats.ListCount - 1
    'see if import item matches a current db item
    Set itmX = lvResolve.FindItem(lstCats.List(lRecs(0)))
    If Not itmX Is Nothing Then             ' a match
        If IsEmpty(itmX.SmallIcon) Then     ' but not if it has been modified
            ' write the cross-reference to the current db item & remove from view
            ReadWriteINI "Write", sFileName, sSection, "XRef" & Mid(itmX.Key, 5), CStr(lstCats.ItemData(lRecs(0)))
            lvResolve.ListItems.Remove itmX.Index
            lRecs(1) = lRecs(1) + 1
        End If
    End If
Next
ReadWriteINI "Write", sFileName, sSection, "AutoResolved", CStr(lRecs(1))
' depending on whether or not something needs to be resolved, display those items
If lvResolve.ListItems.Count = 0 Then
    If optResolve(0).Value = True Then
        optResolve(0).Enabled = False
        If optResolve(1).Enabled = True Then optResolve(1).Value = True
    Else
        If optResolve(1).Value = True Then
            optResolve(1).Enabled = False
            If optResolve(0).Enabled = True Then optResolve(0).Value = True
        End If
    End If
End If
Return

LoadExportedCatLangData:
' display all langauges/categories including any pending actions
Set itmX = lvResolve.ListItems.Add(, "Rec#" & lRecs(1), "...")
sItem = "Action" & lRecs(1): GoSub ReadLog
Select Case sValue
Case "", "0" ' no action set yet or New category/language
    If sValue = "0" Then        ' New category
        itmX.SmallIcon = 30
        sValue = sTmp & " {NEW}"
    Else                        ' No action pending
        sValue = sTmp
    End If
Case "1"    ' Delete
    sValue = sTmp & " {Don't Add}"
    itmX.SmallIcon = 29
Case Else   ' Edit
    If Left(sValue, 1) = "C" Then   ' auto-change/auto-resolved to existing database item
        sValue = sTmp
    Else                            ' change to existing database item
        sValue = sTmp & " {chg to " & lstCats.List(Val(Mid(sValue, 2))) & "}"
        itmX.SmallIcon = 31
    End If
End Select
itmX.Text = sValue
itmX.Tag = sTmp
Return

PopulateNow:
' popuplates a listview or listbox with the current recordset
If bListBox = False Then lvTemp.ListItems.Clear Else lvTemp.Clear
If rstImport.RecordCount Then
    rstImport.MoveFirst
    With rstImport
        Do While .EOF = False
            If bListBox Then                    ' list box item
                lvTemp.AddItem .Fields(1)       ' set text & ref to RecordID
                lvTemp.ItemData(lvTemp.ListCount - 1) = .Fields(0)
            Else                                ' listview item
                Set itmX = lvTemp.ListItems.Add(, "Rec#" & .Fields(0), .Fields(1))
                For lFields = 2 To .Fields.Count - 1
                    If IsNull(.Fields(lFields)) Then sFieldValue = "" Else sFieldValue = .Fields(lFields)
                    itmX.SubItems(lFields - 1) = sFieldValue
                Next
            End If
        .MoveNext
        Loop
    End With
End If
rstImport.Close
' when initially displaying these items, ensure no listview item is selected
If bListBox Then
    lvTemp.ListIndex = -1
Else
    If Not lvTemp.SelectedItem Is Nothing Then lvTemp.ListItems(lvTemp.SelectedItem.Index).Selected = False
    Set lvTemp.SelectedItem = Nothing
    lvTemp.Refresh
End If
Return

Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
PopulateLists_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: PopulateLists" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub Form_Load()
' initialize listview imagelists

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto Form_Load_General_ErrTrap

Set lvMaster.SmallIcons = frmLibrary.SmallImages
Set lvResolve.SmallIcons = frmLibrary.SmallImages
Set lvConfirm.SmallIcons = frmLibrary.SmallImages
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
Form_Load_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: Form_Load" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Function ExportData() As Boolean
' Core function that exports data to a text file. Two files are built
'   1) Index file & 2) a data file containing memo fields & OLE binary fields

Dim sSQL As String, rstExport(0 To 1) As DAO.Recordset
Dim lRecs(0 To 1) As Long, bChunk() As Byte, lRtn As Long
Dim I As Long, J As Long, fNr As Integer, bExports(0 To 1) As Boolean
Dim sSection As String, sItem As String, sValue As String, sTmp As String, vFields() As String
Dim maxBytes As Long, nrBytes As Long, bytesWrote As Long, bytes2read As Long

' remove any existing export file

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto ExportData_General_ErrTrap

If Len(Dir(sFileName)) Then Kill sFileName
If Len(Dir(sFileNameBinary)) Then Kill sFileNameBinary

If lstCodeImport.ListCount Then
' Step 1 - Write Code information
    maxBytes = 32000
    fNr = FreeFile()
    Open sFileNameBinary For Binary As #fNr
    sSQL = "Select * From tblSourceCode"
    Set rstExport(0) = mainDB.OpenRecordset(sSQL, dbOpenDynaset)
    sSQL = "Select * From tblAttachments"
    Set rstExport(1) = mainDB.OpenRecordset(sSQL, dbOpenDynaset)
    ReDim vFields(1 To 4)
    vFields(1) = "OrigDate": vFields(2) = "UpdateDate"
    vFields(3) = "CodeName": vFields(4) = "Purpose"
    For I = 0 To lstCodeImport.ListCount - 1
        sSection = "tblSourceCode"          ' index section name
        lRecs(0) = 0
        If rstExport(0).RecordCount Then
            With rstExport(0)
                ' locate the record to export
                .FindFirst "[IDnr]=" & lstCodeImport.ItemData(I)
                If .NoMatch = False Then
                    ' write text/date fields to the index file
                    GoSub WriteFieldValues
                    ' write the memo fields
                    GoSub WriteMemoFields
                End If
            End With
        End If
    Next
' Step 2 - Write attachment information for Exported Code
    For I = 0 To lstCodeImport.ListCount - 1
    lRecs(0) = 0
        With rstExport(1)
            .FindFirst "[RecIDRef]=" & lstCodeImport.ItemData(I)
            Do While .NoMatch = False
                sSection = "tblAttachments"
                lRecs(0) = lRecs(0) + 1
                sItem = "Attachment" & .Fields("ID")
                sValue = .Fields("Viewer")
                If Val(sValue) = 0 Then sValue = .Fields("Attachment").FieldSize
                nrBytes = Val(sValue)
                GoSub WriteToLog
                bytesWrote = 0
                Do While bytesWrote < nrBytes
                    If nrBytes - bytesWrote < maxBytes Then bytes2read = nrBytes - bytesWrote Else bytes2read = maxBytes
                    bChunk = .Fields("Attachment").GetChunk(bytesWrote, bytes2read)
                    Put #fNr, , bChunk()
                    bytesWrote = bytesWrote + bytes2read
                Loop
                sItem = "Description" & .Fields("ID")
                sValue = .Fields("Description"): GoSub WriteToLog
                sItem = "FileName" & .Fields("ID")
                sValue = .Fields("FileName"): GoSub WriteToLog
                sItem = "Viewer" & .Fields("ID")
                sValue = .Fields("Viewer"): GoSub WriteToLog
                sSection = "tblSourceCode"
                sItem = "Attachment" & I + 1 & "_" & lRecs(0)
                sValue = .Fields("ID"): GoSub WriteToLog
                .FindNext "[RecIDRef]=" & lstCodeImport.ItemData(I)
            Loop
            sSection = "tblSourceCode"
            sItem = "Attachments" & I + 1: sValue = lRecs(0): GoSub WriteToLog
        End With
    Next
    Close #fNr
    rstExport(0).Close
    rstExport(1).Close
    sSection = "tblCategories": sItem = "Number": sValue = lstCodeImport.ListCount: GoSub WriteToLog
    sSection = "Import Objects": sItem = "tblSourceCode": sValue = "YES": GoSub WriteToLog

' Step 3 - Write each of the Language and Category names
    For J = 1 To 2
        lRecs(0) = 0
        sSection = Choose(J, "tblLanguage", "tblCategories")
        sSQL = "Select * From " & sSection & " Order by " & Choose(J, "Language", "Category")
        Set rstExport(0) = mainDB.OpenRecordset(sSQL, dbOpenDynaset)
        If rstExport(0).RecordCount = 0 Then
            sItem = "Number": sValue = "0"
            GoSub WriteToLog
        Else
            With rstExport(0)
                .MoveFirst
                Do While .EOF = False
                    lRecs(0) = lRecs(0) + 1
                    sItem = .Fields(0).Name & lRecs(0): sValue = CStr(.Fields(0)): GoSub WriteToLog
                    sItem = .Fields(1).Name & .Fields(0): sValue = CStr(.Fields(1)): GoSub WriteToLog
                    .MoveNext
                Loop
            End With
            sItem = "Number": sValue = CStr(rstExport(0).RecordCount)
            GoSub WriteToLog
        End If
        rstExport(0).Close
    Next
' Step 4 - Write each Language & Category cross-reference for Exported Code
    sSection = "tblSourceCode"
    sSQL = "Select CatID as TblID, CodeID, 0 As TblRef From tblCodeCatXref " & _
        "Union Select LangID as TblID, CodeID, 1 As TblRef From tblCodeLangXref"
    Set rstExport(0) = mainDB.OpenRecordset(sSQL, dbOpenDynaset)
    For I = 0 To lstCodeImport.ListCount - 1
        lRecs(0) = 0: lRecs(1) = 0
        If rstExport(0).RecordCount Then
            With rstExport(0)
                .FindFirst "[CodeID]=" & lstCodeImport.ItemData(I)
                Do While .NoMatch = False
                    lRecs(.Fields(2)) = lRecs(.Fields(2)) + 1
                    sItem = Choose(.Fields(2) + 1, "Category", "Language") & I + 1 & "_" & lRecs(.Fields(2))
                    sValue = .Fields(0): GoSub WriteToLog
                    .FindNext "[CodeID]=" & lstCodeImport.ItemData(I)
                Loop
            End With
            sItem = "Categories" & I + 1: sValue = lRecs(0): GoSub WriteToLog
            sItem = "Languages" & I + 1: sValue = lRecs(1): GoSub WriteToLog
        End If
    Next
    rstExport(0).Close
End If

' Step 5 - Write each of the URLs, if any
If Not lvURLs.SelectedItem Is Nothing Then
    sSection = "Import Objects": sItem = "tblURLs": sValue = "YES"
    GoSub WriteToLog
    sSection = "tblURLs"
    lRecs(0) = 0
    For I = 1 To lvURLs.ListItems.Count
        If lvURLs.ListItems(I).Selected = True Then
            With lvURLs.ListItems(I)
                lRecs(0) = lRecs(0) + 1
                sItem = "URL" & lRecs(0): sValue = .Text: GoSub WriteToLog
                sItem = "Description" & lRecs(0): sValue = .SubItems(1): GoSub WriteToLog
            End With
        End If
    Next
    sItem = "Number": sValue = CStr(lRecs(0)): GoSub WriteToLog
End If
ExportData = True
On Error Resume Next
Set rstExport(0) = Nothing
Set rstExport(1) = Nothing
Exit Function

WriteMemoFields:
For J = 1 To 2
    sTmp = Choose(J, "Code", "Declarations")
    sValue = rstExport(0).Fields(sTmp).FieldSize
    sItem = sTmp & I + 1: GoSub WriteToLog
    nrBytes = Val(sValue) - 1
    bytesWrote = 0
    Do While bytesWrote < nrBytes
        bChunk = rstExport(0).Fields(sTmp).GetChunk(bytesWrote, maxBytes)
        Put #fNr, , bChunk()
        bytesWrote = bytesWrote + UBound(bChunk)
    Loop
Next
Return

WriteFieldValues:
For J = LBound(vFields) To UBound(vFields)
    sItem = vFields(J) & I + 1
    If J < 3 Then
        If IsNull(rstExport(0).Fields(vFields(J))) Then sValue = CStr(Date) Else sValue = rstExport(0).Fields(vFields(J))
    Else
        If IsNull(rstExport(0).Fields(vFields(J))) Then sValue = "Not Provided" Else sValue = rstExport(0).Fields(vFields(J))
    End If
    GoSub WriteToLog
Next
Return

WriteToLog:
ReadWriteINI "Write", sFileName, sSection, sItem, sValue
Return

SetupFieldNames:
ReDim vFields(0 To rstExport(0).Fields.Count - 1)
For I = 0 To UBound(vFields)
    vFields(I) = rstExport(0).Fields(I).Name
Next
Return
Exit  Function

' Inserted by LaVolpe OnError Insertion Program.
ExportData_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: ExportData" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Function

Private Sub ReadExportData()

Dim sSQL As String, rstImport As DAO.Recordset
Dim lRecs(0 To 1) As Long, bDups As Boolean
Dim I As Long, fNr As Integer, bExports(0 To 1) As Boolean, itmX As ListItem
Dim sSection As String, sItem As String, sValue As String, vFields() As String
Dim maxBytes As Long, nrBytes As Long, bytesWrote As Long, bChunk() As Byte


' Inserted by LaVolpe OnError Insertion Program.
On Error Goto ReadExportData_General_ErrTrap

sSection = "Import Objects": sItem = "tblURLs": GoSub ReadLog
If sValue = "YES" Then
    lvURLs.ListItems.Clear
    sSection = "tblURLs"
    lRecs(0) = 1
    sItem = "URL" & lRecs(0): GoSub ReadLog
    Do While Len(sValue)
        Set itmX = lvURLs.ListItems.Add(, "Rec#" & lRecs(0), sValue)
        sItem = "Description" & lRecs(0): GoSub ReadLog
        itmX.SubItems(1) = sValue
        sItem = "URL" & lRecs(0): GoSub ReadLog
        lRecs(0) = lRecs(0) + 1
        sItem = "URL" & lRecs(0): GoSub ReadLog
    Loop
    If lRecs(0) > 1 Then lvURLs.ListItems(lvURLs.SelectedItem.Index).Selected = False
    Set lvURLs.SelectedItem = Nothing
End If
sSection = "Import Objects": sItem = "tblSourceCode": GoSub ReadLog
If sValue = "No" Then
    optResolve(0).Enabled = False
    Exit Sub
End If
lvMaster.ListItems.Clear
sSection = "tblSourceCode"
Set rstImport = mainDB.OpenRecordset(sSection, dbOpenDynaset)
lRecs(0) = 1
sItem = "CodeName" & lRecs(0): GoSub ReadLog
Do While Len(sValue)
    Set itmX = lvMaster.ListItems.Add(, "Rec#" & lRecs(0), sValue)
    sItem = "Purpose" & lRecs(0): GoSub ReadLog
    itmX.SubItems(1) = sValue
    sItem = "Attachments" & lRecs(0): GoSub ReadLog
    If Val(sValue) Then itmX.SmallIcon = 13
    rstImport.FindFirst "[CodeName]=" & Chr(34) & itmX.Text & Chr(34)
    If rstImport.NoMatch = False Then
        itmX.ForeColor = 255
        bDups = True
    End If
    lRecs(0) = lRecs(0) + 1
    sItem = "CodeName" & lRecs(0): GoSub ReadLog
Loop
rstImport.Close
Set rstImport = Nothing
If lRecs(0) > 1 Then lvMaster.ListItems(lvMaster.SelectedItem.Index).Selected = False
Set lvMaster.SelectedItem = Nothing
If bDups = True Then
    MsgBox "The items displayed in RED have names already in your database." & _
        "Should you choose to import them, you will be prompted to provide a different " & _
        "name for the code if desired.", vbInformation + vbOKOnly
End If
chkNoDups.Enabled = bDups
Exit Sub

ReadLog:
sValue = ReadWriteINI("Get", sFileName, sSection, sItem, "")
Return

Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
ReadExportData_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: ReadExportData" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub
Private Sub lstCodeImport_Click()

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto lstCodeImport_Click_General_ErrTrap

lvMaster.ListItems("Rec#" & lstCodeImport.ItemData(lstCodeImport.ListIndex)).Selected = True
lvMaster.ListItems("Rec#" & lstCodeImport.ItemData(lstCodeImport.ListIndex)).EnsureVisible
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
lstCodeImport_Click_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: lstCodeImport_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub lvResolve_AfterLabelEdit(Cancel As Integer, NewString As String)

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto lvResolve_AfterLabelEdit_General_ErrTrap

If Len(NewString) = 0 Then
    Cancel = 1
    Exit Sub
End If
Dim I As Integer
For I = 0 To lstCats.ListCount - 1
    If lstCats.List(I) = NewString Then
        lstCats.ListIndex = I
        cboResolve.ListIndex = 2
        NewString = lvResolve.SelectedItem.Tag
        Cancel = 1
        Exit For
    End If
Next
If I = lstCats.ListCount Then
    cboResolve.ListIndex = 0
    lvResolve.SelectedItem.Tag = NewString
End If
Call cmdResolve_Click
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
lvResolve_AfterLabelEdit_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: lvResolve_AfterLabelEdit" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub optResolve_Click(Index As Integer)

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto optResolve_Click_General_ErrTrap

If Index = 0 Then
    PopulateLists "tblCategories"
    PopulateLists "ImportCategories"
Else
    PopulateLists "tblLanguage"
    PopulateLists "ImportLanguages"
End If
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
optResolve_Click_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: optResolve_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub ImportExports()

Dim lOffset As Long, hFile As Long, Looper As Long, iAttach As Integer
Dim curRec As Long, newRecNr As Long
Dim rstImport(0 To 3) As DAO.Recordset, bChunk() As Byte
Dim maxBytes As Long, bytesRead As Long, lRtn As Long
Dim sSection As String, sItem As String, sValue As String
Dim lRecs(0 To 1) As Long, sFldVals() As Variant, I As Integer, J As Integer
Dim lSourceHdl As Long


' Inserted by LaVolpe OnError Insertion Program.
On Error Goto ImportExports_General_ErrTrap

If sFileNameBinary <> "None" Then
    If Len(Dir(sFileNameBinary)) = 0 Then
         MsgBox "The associated data file shown below doesn't exist in the expected drive." & vbCrLf & _
             sFileNameBinary, vbExclamation + vbOKOnly, "Import Wizard is aborting."
        Exit Sub
    End If
    lSourceHdl = CreateFile(sFileNameBinary, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    maxBytes = 32000
End If
' Step 1 - Append any new categories/languages
For I = 1 To 2
    lRecs(0) = 1
    ReDim sFldVals(0 To 1)
    sSection = Choose(I, "tblCategories", "tblLanguage")
    Set rstImport(0) = mainDB.OpenRecordset(sSection, dbOpenDynaset)
    sItem = "ID" & lRecs(0): GoSub ReadLog
    Do While Len(sValue)
        curRec = Val(sValue)
        sItem = "Action" & curRec: GoSub ReadLog
        With rstImport(0)
            If sValue = "0" Then    ' New category/language
                sItem = "NewItem" & curRec: GoSub ReadLog
                .AddNew
                .Fields(1) = sValue
                newRecNr = .Fields(0)
                .Update
                sItem = "XRef" & curRec: sValue = CStr(newRecNr): GoSub WriteToLog
            End If
        End With
        lRecs(0) = lRecs(0) + 1
        sItem = "ID" & lRecs(0): GoSub ReadLog
    Loop
    rstImport(0).Close
Next

' step 2 - Fill the core part of the record
lRecs(0) = 0
Set rstImport(0) = mainDB.OpenRecordset("tblSourceCode", dbOpenDynaset)
Set rstImport(1) = mainDB.OpenRecordset("tblAttachments", dbOpenDynaset)
Set rstImport(2) = mainDB.OpenRecordset("tblCodeCatXref", dbOpenDynaset)
Set rstImport(3) = mainDB.OpenRecordset("tblCodeLangXref", dbOpenDynaset)
For Looper = 1 To lvMaster.ListItems.Count
    ReDim sFldVals(0 To rstImport(0).Fields.Count - 1)
    sSection = "tblSourceCode"
    If lvMaster.ListItems(Looper).Selected Then
        If lvMaster.ListItems(Looper).ForeColor = 255 Then
            sValue = InputBox("The following title for this code already exists in your database." & _
                vbCrLf & "Choose a different title if you desire, otherwise it will be added with " & _
                "the following name.", "Duplicate Record Name", lvMaster.ListItems(Looper).Text)
            If Len(sValue) Then lvMaster.ListItems(Looper).Text = sValue
        Else
            sValue = lvMaster.ListItems(Looper).Text
        End If
        If Len(sValue) Then sFldVals(3) = sValue Else sValue = "Unknown"
        sItem = "OrigDate" & Looper: GoSub ReadLog
        sFldVals(1) = sValue
        sItem = "UpdateDate" & Looper: GoSub ReadLog
        sItem = "Purpose" & Looper: GoSub ReadLog
        If Len(sValue) Then sFldVals(4) = sValue Else sValue = "Not Provided"
        With rstImport(0)
            .AddNew
            For I = 1 To UBound(sFldVals)
                .Fields(I) = sFldVals(I)
            Next
            newRecNr = .Fields(0)
            lvMaster.ListItems(Looper).Tag = newRecNr
        End With
        ' get the Code & Description fields
        bytesRead = 0
        sItem = "Code" & Looper: GoSub ReadLog
            Do While bytesRead < Val(sValue)
                If Val(sValue) - bytesRead < maxBytes Then ReDim bChunk(1 To Val(sValue) - bytesRead) Else ReDim bChunk(1 To maxBytes)
                ReadFile lSourceHdl, bChunk(1), UBound(bChunk), lRtn, ByVal 0& 'Read from the file
                If lRtn < UBound(bChunk) Then MsgBox "Error reading file ...": Exit Do                         'Check for errors
                bytesRead = bytesRead + lRtn
                rstImport(0).Fields("Code").AppendChunk CStr(bChunk())
            Loop
            lOffset = lOffset + bytesRead
        bytesRead = 0
        sItem = "Declarations" & Looper: GoSub ReadLog
            Do While bytesRead < Val(sValue)
                If Val(sValue) - bytesRead < maxBytes Then ReDim bChunk(1 To Val(sValue) - bytesRead) Else ReDim bChunk(1 To maxBytes)
                ReadFile lSourceHdl, bChunk(1), UBound(bChunk), lRtn, ByVal 0& 'Read from the file
                If lRtn < UBound(bChunk) Then MsgBox "Error reading file ...": Exit Do                         'Check for errors
                bytesRead = bytesRead + lRtn
                rstImport(0).Fields("Declarations").AppendChunk CStr(bChunk())
            Loop
            lOffset = lOffset + bytesRead
        rstImport(0).Update
    Else
        ' move file pointer past skipped records
        sItem = "Code" & Looper: GoSub ReadLog
        lOffset = Val(sValue)
        sItem = "Declarations" & Looper: GoSub ReadLog
        lOffset = Val(sValue) + lOffset
        SetFilePointer lSourceHdl, lOffset, 0, FILE_CURRENT
    End If
Next
For Looper = 1 To lvMaster.ListItems.Count
    ' Update any attachments
    sSection = "tblSourceCode"
    sItem = "Attachments" & Looper: GoSub ReadLog
    If Val(sValue) Then ' then code has attachments
        iAttach = Val(sValue): lOffset = 0
        For I = 1 To iAttach
            sItem = "Attachment" & Looper & "_" & I: GoSub ReadLog
            curRec = Val(sValue)
            sValue = ReadWriteINI("Get", sFileName, "tblAttachments", "Attachment" & curRec, "0")
            If lvMaster.ListItems(Looper).Selected = True Then
                bytesRead = 0
                With rstImport(1)
                    .AddNew
                    Do While bytesRead < Val(sValue)
                        If Val(sValue) - bytesRead < maxBytes Then ReDim bChunk(1 To Val(sValue) - bytesRead) Else ReDim bChunk(1 To maxBytes)
                        ReadFile lSourceHdl, bChunk(1), UBound(bChunk), lRtn, ByVal 0& 'Read from the file
                        If lRtn < UBound(bChunk) Then MsgBox "Error reading file ...": Exit Do                         'Check for errors
                        bytesRead = bytesRead + lRtn
                        .Fields("Attachment").AppendChunk bChunk()
                    Loop
                    lOffset = lOffset + bytesRead
                    sValue = ReadWriteINI("Get", sFileName, "tblAttachments", "Description" & curRec, "None")
                    .Fields("Description") = sValue
                    sValue = ReadWriteINI("Get", sFileName, "tblAttachments", "FileName" & curRec, "Unknown")
                    .Fields("FileName") = sValue
                    sValue = ReadWriteINI("Get", sFileName, "tblAttachments", "Viewer" & curRec, "0")
                    .Fields("Viewer") = Val(sValue)
                    .Fields("RecIDRef") = Val(lvMaster.ListItems(Looper).Tag)
                    .Update
                End With
            Else
                lOffset = lOffset + Val(sValue)
            End If
        Next
        If lvMaster.ListItems(Looper).Selected = False Then SetFilePointer lSourceHdl, lOffset, 0, FILE_CURRENT
    End If
    ' Now to build the language & category cross references
    If lvMaster.ListItems(Looper).Selected = True Then
        For J = 1 To 2
            sItem = Choose(J, "Categories", "Languages")
            sValue = ReadWriteINI("Get", sFileName, "tblSourceCode", sItem & Looper, "0")
            iAttach = Val(sValue)
            For I = 1 To iAttach
                sSection = "tblSourceCode"
                sItem = Choose(J, "Category", "Language") & Looper & "_" & I: GoSub ReadLog
                curRec = Val(sValue)
                sSection = Choose(J, "tblCategories", "tblLanguage")
                sItem = "XRef" & curRec: GoSub ReadLog
                If sValue <> "" Then
                    With rstImport(J + 1)
                        .AddNew
                        .Fields(0) = Val(sValue)
                        .Fields(1) = Val(lvMaster.ListItems(Looper).Tag)
                        .Update
                    End With
                End If
            Next
        Next
    End If
Next
On Error Resume Next
CloseHandle lSourceHdl
For I = 1 To 3
    rstImport(I).Close
    Set rstImport(I) = Nothing
Next
If Not lvURLs.SelectedItem Is Nothing Then
    Set rstImport(0) = mainDB.OpenRecordset("tblURLs", dbOpenDynaset)
    For Looper = 1 To lvURLs.ListItems.Count
        If lvURLs.ListItems(Looper).Selected Then
            With rstImport(0)
                .FindFirst "[URL]=" & Chr(34) & lvURLs.ListItems(Looper).Text & Chr(34)
                If .NoMatch = False And chkDupURLs.Value = 1 Then .Edit Else .AddNew
                .Fields(1) = lvURLs.ListItems(Looper).Text
                .Fields(2) = lvURLs.ListItems(Looper).SubItems(1)
                .Update
            End With
        End If
    Next
End If
rstImport(0).Close
Set rstImport(0) = Nothing
Exit Sub

ReadLog:
sValue = ReadWriteINI("Get", sFileName, sSection, sItem, "")
Return
WriteToLog:
ReadWriteINI "Write", sFileName, sSection, sItem, sValue
Return
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
ImportExports_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: ImportExports" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub SetFileName()

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto SetFileName_General_ErrTrap

With frmLibrary.dlgCommon
    If iImEx = 0 Then
        .DialogTitle = "Location for Export Files"
        .Flags = cdlOFNPathMustExist
        .FileName = "Select folder and click OPEN button"
        .Filter = ""
    Else
        .DialogTitle = "LaVolpe Code Safe Export Files"
        .Flags = cdlOFNFileMustExist
        .FileName = ""
        .Filter = "LaVolpe Code Safe Exports|*.csx"
    End If
End With
On Error GoTo UserCnx
frmLibrary.dlgCommon.ShowOpen
On Error GoTo BadFile
If iImEx = 0 Then
    sFileName = StripFile(frmLibrary.dlgCommon.FileName, "P") & "LvCSFexport_"
    sFileNameBinary = sFileName & "dat_" & Format(Date, "yymmdd") & ".csx"
    sFileName = sFileName & Format(Date, "yymmdd") & ".csx"
    If Len(Dir(sFileName)) Then
        If MsgBox("An export file with today's date already exists there. Overwrite it?", vbInformation + vbYesNo) = vbNo Then
            MsgBox "Ok. Select another folder or rename/move that file.", vbInformation + vbOKOnly
            sFileName = ""
            sFileNameBinary = ""
            Exit Sub
        End If
    End If
Else
    sFileName = frmLibrary.dlgCommon.FileName
    sFileNameBinary = ReadWriteINI("Get", sFileName, "Data Files", "File1", "")
    If Len(sFileNameBinary) = 0 Then
        If InStr(sFileName, "_dat_") Then
            sFileName = Replace(sFileName, "_dat", "")
            sFileNameBinary = ReadWriteINI("Get", sFileName, "Data Files", "File1", "")
            If Len(sFileNameBinary) = 0 Then Err.Raise 5
        Else
            Err.Raise 5
        End If
    End If
    
End If
Exit Sub

BadFile:
MsgBox "Select another CSX file. That was not the right one." & vbCrLf & _
    "The file will look like LvCSFexport_date.csx" & vbCrLf & vbCrLf & _
    "Note: There are at least two of those files." & vbCrLf & _
    "Do not select the one with the _dat_ in the middle of the name.", vbInformation + vbOKOnly
sFileName = ""
UserCnx:
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
SetFileName_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: SetFileName" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Public Sub CleanUpIndex()
Dim sSection As String, sItem As String, sItem2 As String, sValue As String, Looper As Long
Dim I As Integer, lRecs As Long

' Inserted by LaVolpe OnError Insertion Program.
On Error Goto CleanUpIndex_General_ErrTrap

For I = 1 To 2
    sSection = Choose(I, "tblCategories", "tblLanguage")
    sItem = "Number": GoSub ReadLog
    lRecs = Val(sValue)
    For Looper = 1 To lRecs
        sItem = "ID" & Looper: GoSub ReadLog
        sItem = "XRef" & sValue
        sItem2 = sValue: sValue = "": GoSub WriteToLog
        sItem = "NewItem" & sItem2: sValue = "": GoSub WriteToLog
        sItem = "Action" & sItem2: sValue = "": GoSub WriteToLog
    Next
    sItem = "Resolved": sValue = "0": GoSub WriteToLog
    sItem = "AutoResolved": sValue = "0": GoSub WriteToLog
Next
Exit Sub

ReadLog:
sValue = ReadWriteINI("Get", sFileName, sSection, sItem, "")
Return
WriteToLog:
ReadWriteINI "Write", sFileName, sSection, sItem, sValue
Return
Exit  Sub

' Inserted by LaVolpe OnError Insertion Program.
CleanUpIndex_General_ErrTrap:
Msgbox "Err: " & Err.Number & " - Procedure: CleanUpIndex" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub
