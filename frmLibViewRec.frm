VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLibViewRec 
   Caption         =   "View Code"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   HasDC           =   0   'False
   HelpContextID   =   2
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Tag             =   "924010170"
   Begin VB.CommandButton cmdWordWrap 
      Caption         =   "Wrap"
      Height          =   375
      Index           =   1
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Word Wrap on/off"
      Top             =   5130
      Width           =   615
   End
   Begin VB.CommandButton cmdWordWrap 
      Caption         =   "Wrap"
      Height          =   375
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Word Wrap on/off"
      Top             =   3225
      Width           =   615
   End
   Begin MSComctlLib.Toolbar tbSub 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Z"
            Object.ToolTipText     =   "Abort / Exit"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BtnX"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "A"
            Object.ToolTipText     =   "Save"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnuSaveDefault"
                  Text            =   "Save Record"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnuSave2File"
                  Text            =   "Save Declaration-Code to File"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "B"
            Object.ToolTipText     =   "Save As..."
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "C"
            Object.ToolTipText     =   "Recover from Last Save"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "D"
            Object.ToolTipText     =   "Delete"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BtnY"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E"
            Object.ToolTipText     =   "Add, Edit Attachments"
            Style           =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BtnZ"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "F"
            Object.ToolTipText     =   "Copy to Clipboard Options"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CopyPurpose"
                  Object.Tag             =   "1"
                  Text            =   "from Purpose"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CopyDeclare"
                  Object.Tag             =   "2"
                  Text            =   "from Delcarations Section"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CopyCode"
                  Object.Tag             =   "3"
                  Text            =   "from Code Section"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CopyCodeName"
                  Object.Tag             =   "4"
                  Text            =   "from Code Name"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DD"
                  Text            =   "Load in Multi-Code Clipboard"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Jump2"
            Object.ToolTipText     =   "Go to field... {Default Code Section}"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Jump2Purpose"
                  Object.Tag             =   "1"
                  Text            =   "Purpose"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Jump2Declarations"
                  Object.Tag             =   "2"
                  Text            =   "Declarations"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Jump2Code"
                  Object.Tag             =   "3"
                  Text            =   "Code"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Jump2CodeName"
                  Object.Tag             =   "4"
                  Text            =   "Code Name"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "G"
            Object.ToolTipText     =   "Restore to Normal Screen"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstCategories 
      Height          =   840
      Left            =   4950
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1410
      Width           =   4965
   End
   Begin RichTextLib.RichTextBox txtDeclarations 
      Height          =   1665
      Left            =   705
      TabIndex        =   5
      Top             =   2565
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   2937
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"frmLibViewRec.frx":0000
   End
   Begin VB.TextBox txtCodeName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4935
      MaxLength       =   150
      TabIndex        =   1
      Text            =   "What's my title?"
      Top             =   765
      Width           =   4965
   End
   Begin VB.TextBox txtPurpose 
      Height          =   945
      Left            =   735
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmLibViewRec.frx":00D8
      Top             =   765
      Width           =   4035
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   315
      Left            =   735
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1950
      Width           =   4065
   End
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   4005
      Left            =   705
      TabIndex        =   8
      Top             =   4485
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   7064
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"frmLibViewRec.frx":00EA
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   60
      TabIndex        =   11
      ToolTipText     =   "Change font name and size"
      Top             =   4815
      Width           =   615
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   60
      TabIndex        =   7
      ToolTipText     =   "Change font name and size"
      Top             =   2910
      Width           =   615
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   300
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Search Text"
      Top             =   2595
      Width           =   615
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   300
      Index           =   1
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Search Text"
      Top             =   4500
      Width           =   615
   End
   Begin VB.TextBox txtOrigin 
      Height          =   315
      Left            =   1875
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Date of Origin"
      Top             =   8520
      Width           =   2445
   End
   Begin VB.TextBox txtUpdate 
      Height          =   315
      Left            =   6075
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Date of Last Update"
      Top             =   8520
      Width           =   2445
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Right Click fields for more"
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
      Height          =   1185
      Index           =   6
      Left            =   15
      TabIndex        =   23
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Title of Code (maximum of 150 characters)"
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
      Index           =   2
      Left            =   4965
      TabIndex        =   15
      Top             =   510
      Width           =   4665
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category for Organization (Select as many as needed)"
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
      Index           =   5
      Left            =   4965
      TabIndex        =   18
      Top             =   1170
      Width           =   4695
   End
   Begin VB.Image imgAttach 
      Height          =   360
      Left            =   135
      ToolTipText     =   "Attachments exist for this record"
      Top             =   765
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CODE SECTION   ( Ctrl-A, Ctrl-C, Ctrl-V, Ctrl-Z && Ctrl-X work in all of these these fields )"
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
      Left            =   750
      TabIndex        =   17
      Top             =   4230
      Width           =   8505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DECLARATIONS SECTION -- Include any Public/Private declarations, API functions, etc."
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
      Index           =   3
      Left            =   750
      TabIndex        =   16
      Top             =   2340
      Width           =   8565
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose of this Code (maximu of 255 characters)"
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
      Index           =   1
      Left            =   780
      TabIndex        =   14
      Top             =   510
      Width           =   4275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Language of Code "
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
      Index           =   0
      Left            =   735
      TabIndex        =   13
      Top             =   1695
      Width           =   4035
   End
   Begin VB.Label lblUpdate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Last Updated: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   4245
      TabIndex        =   22
      Top             =   8550
      Width           =   1785
   End
   Begin VB.Label lblOrigin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Added: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   735
      TabIndex        =   20
      Top             =   8550
      Width           =   1125
   End
End
Attribute VB_Name = "frmLibViewRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Private myRecRef As Variant                     ' reference to the db record ID of the record being displayed
Private CatLoaded As Date                       ' date/time when categories were loaded into toolbar
Private LangLoaded As Date                      ' date/time when languages were loaded into combo box
Private bDataChanged As Boolean             ' set to true if data on this form changes without being saved
Private PreviousValue() As Variant                ' used to help identify if data was changed or not
Private lLastLine(0 To 1) As Long
Private bColorCheck(0 To 1) As Boolean
Private ColorCodes As Variant
Private bUpdatable As Boolean

Private Sub DisplayRecord()
'============================================================
'   Function loads all data from the database, adjusts the toolbar as necessary & sets initial values
'============================================================

' Inserted by LaVolpe
On Error GoTo Sub_DisplayRecord_General_ErrTrap_by_LaVolpe
BeginDisplay:
If IsNull(myRecRef) Then         ' New record
    ReloadCategories                ' Load Categories into the listbox
    ReloadLanguages                 ' Load languages into the combo box
    ResetTextBoxes                  ' Blank out all text/RTF boxes
    Caption = "New Record"      ' Adjust caption to show new record & assign icon to titlebar
    Icon = frmLibrary.SmallImages.ListImages(4).ExtractIcon
Else                                        ' Existing record
    Icon = frmLibrary.SmallImages.ListImages(5).ExtractIcon ' Assign icon to titlebar
    LoadAttachmentListing                                           ' Load attachments (if any) to the toolbar
    ReloadCategories                                                   ' Load categories into the listbox
    ReloadLanguages                                                   ' Load lanaguages into the combo box
    With mainRS     ' using the same recordset as the main program
        .FindFirst "[IDnr] = " & myRecRef       ' Locate the associated record clicked on from the main window
        If .NoMatch = True Then                     ' There should always be a match, but just in case....
            mainRS.Requery                              ' requery the database & if there is still no match then give following warning
            .FindFirst "[IDnr] = " & myRecRef       ' Locate the associated record clicked on from the main window
            If .NoMatch = True Then                     ' This should always be a match, but just in case....
                Dim iResponse As Integer
                iResponse = MsgBox("That record no longer seems to be in the database or the filter is preventing it from being displayed." & vbCrLf & vbCrLf & _
                    "Do you want to remove any filters?", vbQuestion + vbYesNo, "Record Not Found")
                If iResponse = vbYes Then
                    .Close
                    frmLibrary.FilterRecordsetNow True
                    GoTo BeginDisplay
                End If
                On Error Resume Next
                frmLibrary.lvCode.ListItems("RecID:" & myRecRef).Tag = ""   ' check back in since it isn't being used
                myRecRef = Null                             ' Reset the reference to existing record so it doesn't get overwritten -- just in case
                bDataChanged = False
                Tag = ""
                Exit Sub
            End If
        End If
        ' Recaption titlebar to include code name if this is not a new record
        Caption = .Fields("CodeName")
        ' Ok, so far so good -- gotta populate the form with the recordset data
        txtCodeName = .Fields("CodeName")   ' the title of the code & the date code was first added to db
        txtOrigin = Format(.Fields("OrigDate"), "mmm d, yyyy h:nn AM/PM")
        txtPurpose = .Fields("Purpose")              ' the purpose of the code & the date the code was last modified
        txtUpdate = Format(.Fields("UpdateDate"), "mmm d, yyyy h:nn AM/PM")
        frmLibrary.rtfStaging = .Fields("Code")
        CheckLine4KeyWords txtCode, frmLibrary.rtfStaging, , True
        frmLibrary.rtfStaging = .Fields("Declarations")
        CheckLine4KeyWords txtDeclarations, frmLibrary.rtfStaging, , True
        lblUpdate.Visible = True                        ' since this is an existing record, show the Date Last Updated label
        txtUpdate.Visible = True                        ' otherwise it would be hidden for "New" records
        tbSub.Buttons("C").Enabled = True       ' Enable the Undo button (not active on "New" records)
    End With
End If
bDataChanged = False                                    ' Set flag to False
bColorCheck(0) = False
bColorCheck(1) = False
Show                                                                ' Ensure form is visible
DoGradient Me, 1                                            ' Repaint the form to user defined colors
If IsNull(myRecRef) Then Tag = "NewRecord" Else Tag = "RecID:" & myRecRef
bUpdatable = True
Exit Sub

Sub_DisplayRecord_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub DisplayRecord]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub ResetTextBoxes()
'============================================================
' Function simply zeros out text boxes and other data as needed when a new record or UNDO happens
'============================================================

Dim I As Integer
' Inserted by LaVolpe
On Error GoTo Sub_ResetTextBoxes_General_ErrTrap_by_LaVolpe
txtCode = "No Code"                                     ' Default for the code section
txtCodeName = "What's my title?"                  ' Default for the code name/title
txtDeclarations = "'No Declarations"                 ' Default for the declaration section
txtPurpose = "Purpose of Code"                      ' Default for the purpose section
' Unselect all categories
For I = 0 To lstCategories.ListCount - 1: lstCategories.Selected(I) = False: Next
tbSub.Buttons("B").Enabled = False              ' Disable the save as button
tbSub.Buttons("C").Enabled = False              ' Disable the UNDO button
tbSub.Buttons("E").Enabled = False               ' Disable the Attachments button
' Hide the field showing when code was last updated
lblUpdate.Visible = False: txtUpdate.Visible = False
' Set the code origination date to current date/time
txtOrigin = Format(Now(), "mmm d, yyyy h:nn AM/PM")
' Reset window's icon
Icon = frmLibrary.SmallImages.ListImages(4).ExtractIcon
Exit Sub

Sub_ResetTextBoxes_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub ResetTextBoxes]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub LoadAttachmentListing()
'============================================================
' Function will add any attachments to the toolbar
'============================================================

Dim I As Integer, iconID As Integer
' First get rid of any existing attachments
' Inserted by LaVolpe
On Error GoTo Sub_LoadAttachmentListing_General_ErrTrap_by_LaVolpe
For I = tbSub.Buttons("E").ButtonMenus.Count To 1 Step -1
    tbSub.Buttons("E").ButtonMenus.Remove I
Next

If IsNull(myRecRef) = False Then        ' If this is a new record, it won't have any attachments
    ' Now query database to see if any attachments exist for this record
    Dim rsAttach As DAO.Recordset, strSQL As String
    strSQL = "Select * From tblAttachments Where (tblAttachments.RecIDRef)=" & myRecRef _
        & " ORDER BY tblAttachments.Description"
    Set rsAttach = mainDB.OpenRecordset(strSQL, dbOpenDynaset)
    With rsAttach
        If .RecordCount Then                 ' Attachments exist
            .MoveFirst                              ' Goto first one
            I = 1                                       ' start a counter
            Do While .EOF = False           ' Loop thru each attachment & add to the toolbar
                ' the Key property will always be RecID:####, where #### is the value of the individual record's db ID
                tbSub.Buttons("E").ButtonMenus.Add , "RecID:" & .Fields("ID"), I & ". " & .Fields("Description")
                I = I + 1                               ' increment the counter & goto the next attachment, if any
                .MoveNext
            Loop                                        ' show the Attachment icon
            iconID = 13
        End If
    .Close                                              ' close & reset the temporary recordset
    End With
    Set rsAttach = Nothing
End If
tbSub.Buttons("E").Enabled = Not IsNull(myRecRef)   ' disable the Attachment button for new records
On Error Resume Next
If iconID Then
    frmLibrary.lvCode.ListItems("RecID:" & myRecRef).SmallIcon = iconID
    imgAttach.Picture = frmLibrary.LargeImages.ListImages(iconID).ExtractIcon
Else
    Set frmLibrary.lvCode.ListItems("RecID:" & myRecRef).SmallIcon = Empty
    imgAttach.Picture = LoadPicture("")
End If
frmLibrary.lvCode.Refresh
Exit Sub

Sub_LoadAttachmentListing_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub LoadAttachmentListing]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cboLanguage_Click()
'============================================================
'  Function sees if user changed the value
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_cboLanguage_Click_General_ErrTrap_by_LaVolpe
If Len(cboLanguage.Tag) Then Exit Sub       ' Tag property when tracking changes is not wanted (initial load)
' the Previous(0) variable was set in the Got Focus module & if is different now, then tag record as changed
If cboLanguage.ItemData(cboLanguage.ListIndex) <> PreviousValue(0) Then bDataChanged = True
Exit Sub

Sub_cboLanguage_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cboLanguage_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cboLanguage_DblClick()
Call cmdLanguageAdd_Click
End Sub

Private Sub cboLanguage_GotFocus()
'============================================================
' Stores the value of the Language setting before any action taken to change it
'============================================================
On Error Resume Next
ReDim PreviousValue(0)
PreviousValue(0) = cboLanguage.ItemData(cboLanguage.ListIndex)
End Sub

Private Sub cmdAddCat_Click()
'============================================================
'   Sub calls the form that allows adding/editing/deleting categories for this program
'============================================================
Dim I As Integer, sNewValue As String, J As Integer
' Keep track of how many categories are selected and which ones, before they are changed
' Inserted by LaVolpe
On Error GoTo Sub_cmdAddCat_Click_General_ErrTrap_by_LaVolpe
ReDim PreviousValue(0 To lstCategories.ListCount)
PreviousValue(0) = lstCategories.SelCount
For I = 0 To lstCategories.ListCount - 1
    If lstCategories.Selected(I) = True Then
        J = J + 1                                                           ' increment counter
        PreviousValue(J) = lstCategories.ItemData(I)    ' store record ID of this category
    End If
Next
' Now call the form which modifies categories
GP = "Cats"
frmCats.Show 1, frmLibrary
' Check to see if any changes were made by comparing the date when categories were last updated compared
'   to when the ones on this form were last loaded
If LastCatUpdate > CatLoaded Then
    ReloadCategories                                ' they changed, so we have to reload & compare to see if any changed
    If lstCategories.SelCount <> PreviousValue(0) Then
        bDataChanged = True             ' if the number of those selected is different than before, changes made
    Else                                            ' otherwise, lets see if the ones selected are the same as before
        For J = 1 To UBound(PreviousValue)
            For I = 0 To lstCategories.ListCount - 1
                ' See if a selected new category was also a selected category before any changes were made
                If lstCategories.Selected(I) = True And lstCategories.List(I) = PreviousValue(J) Then Exit For
            Next
            If I = lstCategories.ListCount Then         ' went thru the list without finding a match from before, so
                bDataChanged = True
                Exit For
            End If
        Next
    End If
End If
Exit Sub

Sub_cmdAddCat_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdAddCat_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdFont_Click(Index As Integer)

' Inserted by LaVolpe
On Error GoTo Sub_cmdFont_Click_General_ErrTrap_by_LaVolpe
With frmLibrary.dlgCommon                                       ' display the color dialog box
    On Error GoTo UserCnx
    .FontName = Choose(Index + 1, txtDeclarations.Font.Name, txtCode.Font.Name)
    .FontSize = Choose(Index + 1, txtDeclarations.Font.Size, txtCode.Font.Size)
    .Flags = cdlCFBoth
End With
frmLibrary.dlgCommon.ShowFont
If Index = 0 Then
    txtDeclarations.SelStart = 0
    txtDeclarations.SelLength = Len(txtDeclarations.TextRTF)
    txtDeclarations.SelFontName = frmLibrary.dlgCommon.FontName
    txtDeclarations.SelFontSize = frmLibrary.dlgCommon.FontSize
    txtDeclarations.SelStart = 0
Else
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.TextRTF)
    txtCode.SelFontName = frmLibrary.dlgCommon.FontName
    txtCode.SelFontSize = frmLibrary.dlgCommon.FontSize
    txtCode.SelLength = 0
End If

UserCnx:
Exit Sub

Sub_cmdFont_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdFont_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdLanguageAdd_Click()
'============================================================
' Calls form which allows user to add to and delete languages for the database
'============================================================

' First, gotta keep track of what the current language value is
' Inserted by LaVolpe
On Error GoTo Sub_cmdLanguageAdd_Click_General_ErrTrap_by_LaVolpe
ReDim PreviousValue(0)
PreviousValue(0) = cboLanguage.ItemData(cboLanguage.ListIndex)
' Now show the form
GP = "Lang"
frmCats.Show 1, frmLibrary
' Check to see if any changes were made by comparing the date when languages were last updated compared
'   to when the ones on this form were last loaded
If LastLangUpdate > LangLoaded Then     ' they changed, so reload them and compare any changes
    ReloadLanguages                                    ' see if the language changed & tag record as changed if needed
    If PreviousValue(0) <> cboLanguage.ItemData(cboLanguage.ListIndex) Then bDataChanged = True
End If
Exit Sub

Sub_cmdLanguageAdd_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdLanguageAdd_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdSearch_Click(Index As Integer)
Dim sCriteria As String, sMsg As String, lStartSearch As Long, iWholeWord As Integer, rtfCtrl As RichTextBox
Dim iOffset As Integer
If Index > 2 Then iOffset = 20
If cmdSearch(Index - iOffset).Tag = "" Or Index < 2 Then
    sMsg = "Enter the search string below." & vbCrLf & "For whole-word matches start search string with an exclamation mark (!)"
    sCriteria = InputBox(sMsg, "Search Criteria", cmdSearch(Index - iOffset).Tag)
    If Trim(sCriteria) = "" Or Trim(sCriteria) = "!" Then Exit Sub
Else
     sCriteria = cmdSearch(Index - iOffset).Tag
End If
Index = Index - iOffset
If Index = 0 Then Set rtfCtrl = txtDeclarations Else Set rtfCtrl = txtCode
cmdSearch(Index).Tag = sCriteria
With rtfCtrl
    If cmdSearch(Index).Tag = "" Then lStartSearch = 0 Else lStartSearch = .SelStart + 1
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
Set rtfCtrl = Nothing
End Sub

Private Sub cmdWordWrap_Click(Index As Integer)
Dim iOnOff As Integer
If cmdWordWrap(Index).BackColor = 65535 Then iOnOff = 1 Else iOnOff = 0
SendMessageLong Choose(Index + 1, txtDeclarations.hWnd, txtCode.hWnd), EM_SETTARGETDEVICE, 0, iOnOff
cmdWordWrap(Index).BackColor = Choose(iOnOff + 1, 65535, &H8000000F)
End Sub

Private Sub Form_Activate()
On Error Resume Next
If Left(Tag, 5) <> "RecID" Then Exit Sub
frmLibrary.lvCode.ListItems(Tag).Selected = True
frmLibrary.lvCode.ListItems(Tag).EnsureVisible
End Sub

Private Sub Form_Load()
'============================================================
'   Function simply assigns menubar buttons and calls routine to load the data
'============================================================

Dim I As Integer
' Inserted by LaVolpe
On Error GoTo Sub_Form_Load_General_ErrTrap_by_LaVolpe
myRecRef = DBrecID                                          ' save the reference of the db record ID, if any
Set tbSub.ImageList = frmLibrary.SmallImages    ' assign image list to the toolbar
With frmLibrary.SmallImages                                 ' now just assign buttons to the toolbar
    For I = 1 To 9
        tbSub.Buttons.Item(Choose(I, 1, 3, 4, 5, 6, 8, 10, 12, 14)).Image = I + 16
    Next
    For I = 0 To 1
        cmdSearch(I).Picture = .ListImages(6).ExtractIcon
    Next
End With
For I = Label1.LBound To Label1.UBound
    Label1(I).ForeColor = MyDefaults.LblColorMain
    Label1(I).FontBold = (MyDefaults.WindowColor <> WinBlahColor)
Next
txtCode.Font.Name = MyDefaults.Font: txtCode.Font.Size = MyDefaults.FontSize
txtCodeName.Font.Name = MyDefaults.Font: txtCodeName.Font.Size = MyDefaults.FontSize
txtDeclarations.Font.Name = MyDefaults.Font: txtDeclarations.Font.Size = MyDefaults.FontSize
txtPurpose.Font.Name = MyDefaults.Font: txtPurpose.Font.Size = MyDefaults.FontSize
Left = 0
Top = 0
Height = Val(Left(Tag, 4))
Width = Val(Mid(Tag, 5))
SendMessageLong txtCode.hWnd, EM_SETTARGETDEVICE, 0, 1
SendMessageLong txtDeclarations.hWnd, EM_SETTARGETDEVICE, 0, 1
DisplayRecord                                                      ' call function to load form with db data
If Tag = "" Then Unload Me
Exit Sub

Sub_Form_Load_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Load]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then Exit Sub
On Error Resume Next
txtDeclarations.Width = Width - 900
txtCode.Width = Width - 900
txtDeclarations.Height = Height * 0.180194805194805
txtCode.Height = Height - txtDeclarations.Height - 3570
Label1(4).Top = txtDeclarations.Top + txtDeclarations.Height
txtCode.Top = txtDeclarations.Top + txtDeclarations.Height + 255
cmdFont(1).Top = Label1(4).Top
cmdSearch(1).Top = Label1(4).Top
lblOrigin.Top = txtCode.Top + txtCode.Height + 60
lblUpdate.Top = lblOrigin.Top
txtOrigin.Top = txtCode.Top + txtCode.Height + 30
txtUpdate.Top = txtOrigin.Top
cmdSearch(0).Top = txtDeclarations.Top
cmdSearch(1).Top = txtCode.Top
cmdFont(0).Top = cmdSearch(0).Top + cmdSearch(0).Height + 15
cmdFont(1).Top = cmdSearch(1).Top + cmdSearch(1).Height + 15
cmdWordWrap(0).Top = cmdFont(0).Top + cmdFont(0).Height + 15
cmdWordWrap(1).Top = cmdFont(1).Top + cmdFont(1).Height + 15
End Sub

Private Sub Form_Terminate()
'============================================================
'       Ensure form is unloaded from memory when user hits the X button
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_Form_Terminate_General_ErrTrap_by_LaVolpe
Unload Me
Exit Sub

Sub_Form_Terminate_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Terminate]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
'============================================================
'  When a user closes the window, prompt if data changed but wasn't saved
'============================================================
On Error Resume Next
If bDataChanged = True Then
    If MsgBox("Fields were changed in above window. Exiting will cause those changes to be lost. Continue?", _
        vbExclamation + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
            bAppClose = True        ' set global variable in case window is being shut down remotely
            Cancel = 1                      ' abort the Close
            Exit Sub
    End If
End If
End Sub

Private Sub lstCategories_Click()
'============================================================
'   Flag record as updated if neeed
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_lstCategories_Click_General_ErrTrap_by_LaVolpe
If Len(lstCategories.Tag) Then Exit Sub
bDataChanged = True
Exit Sub

Sub_lstCategories_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub lstCategories_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub DeleteRecord()
'============================================================
'   Deletes a record, per request
'============================================================
Dim I As Integer, strSQL As String

' First, provide a confirmation message
' Inserted by LaVolpe
On Error GoTo Sub_DeleteRecord_General_ErrTrap_by_LaVolpe
I = MsgBox("Are you sure you want this code deleted?" & vbCrLf & vbCrLf _
    & "IMPORTANT: This will also permaently delete any attachments you may have to this record", _
    vbYesNo + vbDefaultButton2 + vbExclamation, "Confirmation")
If I = vbNo Then Exit Sub

' If deleting an existing record then run the following queries to rid the record from the db
If IsNull(myRecRef) = False Then
    ' delete the record from the tblSourceCode
    strSQL = "DELETE * FROM tblSourceCode Where (tblSourceCode.IDnr) = " & myRecRef
    mainDB.Execute strSQL
    ' delete related records from the tblCodeLangXref
    strSQL = "DELETE * FROM tblCodeLangXref Where (tblCodeLangXref.CodeID) = " & myRecRef
    mainDB.Execute strSQL
    ' delete related records from the tblCodeCatXRef
    strSQL = "DELETE * FROM tblCodeCatXref Where (tblCodeCatXref.CodeID) = " & myRecRef
    mainDB.Execute strSQL
    ' delete any attachments from the tblAttachments
    strSQL = "DELETE * FROM tblAttachments Where (tblAttachments.RecIDRef) = " & myRecRef
    mainDB.Execute strSQL
    On Error Resume Next
    ' remove the record from the main window's code listing
    frmLibrary.lvCode.ListItems.Remove frmLibrary.lvCode.ListItems("RecID:" & myRecRef).Index
End If
Unload Me
Exit Sub

Sub_DeleteRecord_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub DeleteRecord]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub SaveChanges()
'============================================================
' Save changes made, if any
'============================================================
Dim I As Integer, J As Integer, rsChange As DAO.Recordset, strSQL As String
Dim bNewRec As Boolean, iIcon As Integer

' If the category listing was updated but somehow didn't get updated in this window, prompt to load the categories first
' Inserted by LaVolpe
On Error GoTo Sub_SaveChanges_General_ErrTrap_by_LaVolpe
If LastCatUpdate > CatLoaded Then
    I = MsgBox("The categories have been updated since this window opened. Cannot save changes until " & _
        "the categories are reloaded.  Reload now?", vbYesNo + vbQuestion, "Data Updated Since Code Displayed")
    If I = vbNo Then Exit Sub
Else
    ' if the language listing was updated but somehow didn't get updated in this window, prompt to load the languages
    If LastLangUpdate > LangLoaded Then
        I = MsgBox("The languages have been updated since this window opened. Cannot save changes until " & _
            "the languages are reloaded.  Reload now?", vbYesNo + vbQuestion, "Data Updated Since Code Displayed")
        If I = vbNo Then Exit Sub
    End If
End If
' Ok, categories/languages are updated, now continue with changes
If I Then
    If LastCatUpdate > CatLoaded Then ReloadCategories
    If LastLangUpdate > LangLoaded Then ReloadLanguages
    MsgBox "Categories and Language listings have been updated. Please verify any changes & try saving again.", vbInformation + vbOKOnly
    Exit Sub
End If
' don't allow record to be saved if the user didn't provide a name for the code
If txtCodeName = "" Then
    MsgBox "First provide a name/title for this code", vbInformation + vbOKOnly
    Exit Sub
End If
' Now provide a final confirmation
I = MsgBox("Once changes are saved, there is no undo. Continue?", vbYesNo + vbQuestion, "Confirmation")
If I = vbNo Then Exit Sub

' Update the tblSourceCode first
    If IsNull(myRecRef) Then                ' New record
            txtOrigin = Format(Now(), "mmm d, yyyy h:nn AM/PM")    ' ensure origin date is current
            bNewRec = True                                                                  ' set flag
            strSQL = "tblSourceCode"                                                    ' set db query string
    Else                                                ' Existing record
        ' set the db query string to only reference this record
        strSQL = "SELECT * FROM tblSourceCode Where (tblSourceCode.IDnr)=" & myRecRef
    End If
    Set rsChange = mainDB.OpenRecordset(strSQL, dbOpenDynaset)  ' open the recordset to edit
    With rsChange
        If IsNull(myRecRef) Then            ' new record, so Add
            .AddNew
            .Fields("OrigDate") = Now()
        Else
            If .RecordCount = 0 Then        ' otherwise, verify record about to update is still in db
                I = MsgBox("This code no longer exists in the database. Recreate it?", vbExclamation + vbYesNo, "Missing Code")
                If I = vbNo Then GoTo CleanUp   ' if not, and user wants it added as a new record then do so
                .AddNew
                .Fields("OrigDate") = Now()
            Else
                .MoveFirst
                .Edit                                   ' existing record & it's still in the db
            End If
        End If
        .Fields("UpdateDate") = Now()   ' Update the last modified date & update the display textbox
        txtUpdate = Format(.Fields("UpdateDate"), "mmm d, yyyy h:nn AM/PM")
        ' Update the declarations section, providing a default if needed
        If txtDeclarations.Text = "" Then txtDeclarations = "'No Declarations"
        .Fields("Declarations") = txtDeclarations.Text
        ' Update the code section, providing a default if needed
        If txtCode.Text = "" Then txtCode = "No Code"
        .Fields("Code") = txtCode.Text
        .Fields("CodeName") = txtCodeName        ' Update the code name/title
        ' Update the purpose section, providing a default if needed
        If Len(txtPurpose.Text) = 0 Then txtPurpose = "Not Provided"
        .Fields("Purpose") = txtPurpose.Text
        myRecRef = .Fields("IDnr")                          ' keep track of the new db record ID assigned by Access
        .Update                                                         ' Done with tblSourceCode
    End With
rsChange.Close

' Now update tblCodeCatXref next
strSQL = "Delete * From tblCodeCatXref  Where (tblCodeCatXref.CodeID)=" & myRecRef
mainDB.Execute strSQL           ' First we delete any previous references to this record, since this table can
                                                ' have several entries per record
Set rsChange = mainDB.OpenRecordset("tblCodeCatXref", dbOpenDynaset)    ' open recordset to append
For I = 0 To lstCategories.ListCount - 1                            ' loop thru each category to see if it is selected
    With rsChange
        If lstCategories.Selected(I) = True Then                    ' and if so append the table
            .AddNew
            .Fields("CatID") = lstCategories.ItemData(I)        ' Set cross-reference to the category's db record ID
            .Fields("CodeID") = myRecRef                            ' Set cross-reference to this record's db ID
            .Update                                                               ' Done with this table
        End If
    End With
Next
rsChange.Close

' Now update tblCodeLangXref
strSQL = "Select * From tblCodeLangXref  Where (tblCodeLangXref.CodeID)=" & myRecRef
Set rsChange = mainDB.OpenRecordset(strSQL, dbOpenDynaset)  ' open recordset to edit
With rsChange
    If cboLanguage.ListIndex > -1 Then                                              ' if a language is selected, then
        If .RecordCount Then .Edit Else .AddNew                                 ' Add or Update as necessary
        ' set cross-references to this record and the language chosen within this window
        .Fields("LangID") = cboLanguage.ItemData(cboLanguage.ListIndex)
        .Fields("CodeID") = myRecRef
        .Update                                                                 ' Done with this table
    End If
End With
tbSub.Buttons("B").Enabled = True                            ' Enable the save as button
tbSub.Buttons("E").Enabled = True                               ' Enable the Attachment menu button
tbSub.Buttons("C").Enabled = True                               ' Enable the Undo menu button
txtUpdate.Visible = True: lblUpdate.Visible = True        ' Show the last modified display textbox

' Now to finish by adjusting controls on this form as necessary
Caption = txtCodeName               ' Ensure the titlebar caption references this code
' If attachments exist on the menubar, then set the reference the Attachment icon to be used later
If tbSub.Buttons("E").ButtonMenus.Count Then iIcon = 13
' Last but no least, update the main form where necessary
On Error Resume Next
With frmLibrary.lvCode
    If bNewRec Then     ' For new records, add them to the main window's code listing
        .ListItems.Add , "RecID:" & myRecRef, txtCodeName, , iIcon      ' add record
        .ListItems("RecID:" & myRecRef).Tag = "Loaded"                  ' tag as open in another window
        Icon = frmLibrary.SmallImages.ListImages(5).ExtractIcon         ' change this form's icon from New to Edit
    Else
        .ListItems("RecID:" & myRecRef).SmallIcon = iIcon               ' Update attachment icon as needed
        .ListItems("RecID:" & myRecRef).Text = Caption                  ' Update the listing title as needed
    End If
    FilterRecordset
    On Error Resume Next
    .ListItems("RecID:" & myRecRef).Selected = True
    .ListItems("RecID:" & myRecRef).EnsureVisible
End With
bDataChanged = False                                                    ' set flag to record now not changed
If IsNull(myRecRef) Then Tag = "NewRecord" Else Tag = "RecID:" & myRecRef

CleanUp:
On Error Resume Next
rsChange.Close                                                      ' Close recordset
Set rsChange = Nothing
Exit Sub

Sub_SaveChanges_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub SaveChanges]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
Resume CleanUp
End Sub

Private Sub ReloadCategories()
'============================================================
' Loads categories into a listbox
'============================================================

Dim I As Integer
' Inserted by LaVolpe
On Error GoTo Sub_ReloadCategories_General_ErrTrap_by_LaVolpe
lstCategories.Tag = "Reloading"                     ' Set flag to prevent premature flagging as data changed
lstCategories.Clear                                         ' Clear listbox
With frmLibrary.lstFilter                          ' Since categories are updated on the main window,
    For I = 1 To .ListCount - 1                         '       loop thru them to append to our listbox
        lstCategories.AddItem .List(I)
        lstCategories.ItemData(I - 1) = .ItemData(I)
    Next
End With
' If we have categories and this is not a new record, then let's find those categories referenced by this record
If lstCategories.ListCount And IsNull(myRecRef) = False Then
        Dim rsCats As DAO.Recordset, strSQL As String
        ' build a query string to extract only the categories referenced by this record
        strSQL = "SELECT tblCategories.Category FROM tblCategories INNER JOIN " & _
            "tblCodeCatXref ON tblCategories.ID = tblCodeCatXref.CatID " & _
            "WHERE (((tblCodeCatXref.CodeID)=" & myRecRef & "));"
        Set rsCats = mainDB.OpenRecordset(strSQL, dbOpenDynaset)    ' open the recordset for reference
        If rsCats.RecordCount > 0 Then
            With rsCats
                .MoveFirst
                Do While .EOF = False            ' loop thru each and highlight the listbox entry when they match
                    For I = 0 To lstCategories.ListCount - 1
                        If .Fields("Category") = lstCategories.List(I) Then     ' did we match?
                            lstCategories.Selected(I) = True                            ' yep, so highlight the category
                            Exit For
                        End If
                    Next
                .MoveNext                                 ' continue looping
                Loop
            End With
        End If
    rsCats.Close: Set rsCats = Nothing          ' done, so close the recordset
End If
lstCategories.Tag = ""                                  ' reset Tag so changes can be tracked
CatLoaded = Now()                                   ' reset the date these categories were loaded on this form
Exit Sub

Sub_ReloadCategories_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub ReloadCategories]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub ReloadLanguages()
'============================================================
'  Loads languages into a combo box
'============================================================

Dim I As Integer, iDefault As Integer
' Inserted by LaVolpe
On Error GoTo Sub_ReloadLanguages_General_ErrTrap_by_LaVolpe
iDefault = -1
cboLanguage.Tag = "Reloading"               ' Set Tag to prevent premature flagging as an updated record
cboLanguage.Clear                                   ' Clear the combo box
With frmLibrary.cboFilter(2)                     ' Since languages are updated in the main window, we need to
    For I = 1 To .ListCount - 1                   '         loop thru each one and add them to our combo box
       cboLanguage.AddItem .List(I)
        cboLanguage.ItemData(I - 1) = .ItemData(I)
                ' just in case we don't find a match, let's see if the default language is in the list,
                '         and if it is, track its position in case we want to use it later
        If cboLanguage.ItemData(I - 1) = MyDefaults.Language Then iDefault = I - 1
    Next
End With
' If languages exist and this is not a new record then let's find the one referenced by this record
If cboLanguage.ListCount And IsNull(myRecRef) = False Then
        Dim rsCats As DAO.Recordset, strSQL As String
        ' build the query string
        strSQL = "SELECT tblLanguage.Language FROM tblLanguage INNER JOIN " & _
            "tblCodeLangXref ON tblLanguage.ID= tblCodeLangXref.LangID " & _
            "WHERE (((tblCodeLangXref.CodeID)=" & myRecRef & "));"
        Set rsCats = mainDB.OpenRecordset(strSQL, dbOpenDynaset)    ' open the recordset for reference
        If rsCats.RecordCount > 0 Then
            With rsCats
                ' loop thru each item in the combo box looking for a match
                For I = 0 To cboLanguage.ListCount - 1
                    If .Fields("Language") = cboLanguage.List(I) Then   ' do we have a match?
                        cboLanguage.ListIndex = I                                   ' yep, so select it
                        Exit For
                    End If
                Next
            End With
        End If
    rsCats.Close: Set rsCats = Nothing          ' Done, close the recordset
End If
If cboLanguage.ListIndex < 0 Then    ' didn't find a match, so lets try to apply default or any one
    If iDefault > -1 Then       ' no match, but Default language is in list, so we'll use it
        cboLanguage.ListIndex = iDefault
    Else                                ' no match & no default, let's use the first entry
        If cboLanguage.ListCount Then cboLanguage.ListIndex = 0
    End If
End If
cboLanguage.Tag = ""                                ' Remove flag so tracking of updates can be done
LangLoaded = Now()                                ' Update the time these langugaes were last loaded
Exit Sub

Sub_ReloadLanguages_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub ReloadLanguages]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub lstCategories_DblClick()
Call cmdAddCat_Click
End Sub

Private Sub lstCategories_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then ShowRightClick 4
End Sub

Private Sub tbSub_ButtonClick(ByVal Button As MSComctlLib.Button)
'============================================================
'   Menu button clicks
'============================================================
    On Error Resume Next
    If Button.Key = "Z" Then                ' Abort/Exit button
        Unload Me
    Else
        Select Case Asc(Right(Button.Key, 1)) - 65
        Case 0          ' Save
            SaveChanges
        Case 1      ' Save As...
            Dim sNewName As String, I As Long
            If mainRS.RecordCount Then
Look4Duplicates:
                sNewName = Replace(txtCodeName, Chr(34), Chr(34) & Chr(34))
                sNewName = Replace(txtCodeName, "'", "''")
                mainRS.FindFirst "[CodeName]=" & Chr(34) & sNewName & Chr(34)
                If mainRS.NoMatch = False Then
                    sNewName = InputBox("The code name of [" & txtCodeName & "] already exists. What do you want to call this code?" _
                        & vbCrLf & vbCrLf & "Note: Any attachments will not be copied.", "What name do you want for this duplicate record?")
                    If sNewName = "" Then Exit Sub
                    txtCodeName = sNewName
                    GoTo Look4Duplicates
                End If
            End If
            myRecRef = Null
            For I = tbSub.Buttons("E").ButtonMenus.Count To 1 Step -1
                tbSub.Buttons("E").ButtonMenus.Remove I
            Next
            imgAttach.Picture = LoadPicture("")
            SaveChanges
        Case 2          ' Reset
            DBrecID = myRecRef
            DisplayRecord
        Case 3          ' Delete
            DeleteRecord
        Case 4
            '============================================================
            ' User clicked the Add/Edit/Delete attachment dropdown
            '============================================================
            DBrecID = myRecRef                                                      ' set reference to this records db ID
            GP = Caption                                                                    ' set variable to include form's caption
            frmAttachments.Show 1, frmLibrary                                   ' show the Attachments form
            If GP = "Reload" Then LoadAttachmentListing                   ' reload attachments if needed
        Case 5
            Call tbSub_ButtonMenuClick(tbSub.Buttons(10).ButtonMenus(6))
        Case 6                          ' resize to defaults
            If WindowState = vbMaximized Then WindowState = vbNormal
            Width = 10170
            Height = 9240
        Case -15
            txtCode.SetFocus
        End Select
    End If
End Sub

Private Sub tbSub_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'============================================================
' Dropdown menu button clicks
'============================================================

'============================================================
Dim iKey As Integer, lWidth As Long, lHeight As Long, iResponse As Integer, I As Integer
' Inserted by LaVolpe
On Error GoTo Sub_tbSub_ButtonMenuClick_General_ErrTrap_by_LaVolpe
Select Case ButtonMenu.Key
Case "Jump2CodeName":
    iKey = 1
Case "Jump2Purpose":
    iKey = 2
Case "Jump2Declarations":
    iKey = 3
Case "Jump2Code":
    iKey = 4
Case "NoZoom"
    '============================================================
    ' The following are the dropdown option to copy to or paste from specific fields
    '============================================================
Case "CopyPurpose", "CopyDeclare", "CopyCode", "CopyCodeName"
    Clipboard.Clear
    Clipboard.SetText (Choose(Val(ButtonMenu.Tag), txtPurpose, txtDeclarations.Text, txtCode.Text, txtCodeName))
Case "Paste2Purpose", "Paste2Declarations", "Paste2Code", "Paste2CodeName"
    Select Case Val(ButtonMenu.Tag)
    Case 1: txtPurpose = Clipboard.GetText
    Case 2: txtDeclarations = Clipboard.GetText
    Case 3: txtCode = Clipboard.GetText
    Case 4: txtCodeName = Clipboard.GetText
    End Select
    Exit Sub
Case "DD"
    If Tag = "NewRecord" Then
        MsgBox "Only available on saved records.", vbInformation + vbOKOnly
        Exit Sub
    End If
    For I = 0 To Forms.Count - 1            ' loop thru each loaded form
        If Forms(I).Name = "frmClipboard" Then Exit For
    Next
    With frmClipboard.lstMemory
        For I = 0 To .ListCount - 1
            If .ItemData(I) = Val(Mid(Tag, 7)) Then Exit For
        Next
        If bDataChanged = True Then
            iResponse = MsgBox("The multi-code clipboard only accesses saved data. The data in this record" & vbCrLf & _
                "has changed. Do you want to save it now?", vbYesNo + vbQuestion)
            If iResponse = vbYes Then SaveChanges
        End If
        If I = .ListCount Then
            If Len(txtCodeName.Text) > 50 Then
                .AddItem Left(txtCodeName.Text, 50) & "..."
            Else
                .AddItem txtCodeName.Text
            End If
            .ItemData(.NewIndex) = Val(Mid(Tag, 7))
            I = .NewIndex
        End If
        .ListIndex = I
    End With
    frmClipboard.Show
    Me.SetFocus
Case "mnuSaveDefault"
    SaveChanges
Case "mnuSave2File"
    With frmLibrary.dlgCommon
        .DialogTitle = "Select File to add Code to..."
        .DefaultExt = "txt"
        .FileName = ""
        .Filter = "Text Files|*.txt|All Files|*.*"
        .FilterIndex = 1
        .Flags = cdlOFNPathMustExist
    End With
    On Error GoTo UserCnx
    frmLibrary.dlgCommon.ShowSave
    If Len(Dir(frmLibrary.dlgCommon.FileName)) Then
        iResponse = MsgBox("That file already exists. What do you want to do?" & vbCrLf & vbCrLf & _
            "Press YES to overwrite it" & vbCrLf & _
            "Press NO to append to end of file" & vbCrLf & _
            "Press CANCEL to abort", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Overwrite File?")
        If iResponse = vbCancel Then Exit Sub
    Else
        iResponse = vbYes
    End If
    On Error GoTo Sub_tbSub_ButtonMenuClick_General_ErrTrap_by_LaVolpe
    Dim fNr As Integer
    fNr = FreeFile()
    If iResponse = vbNo Then
        Open frmLibrary.dlgCommon.FileName For Append As #fNr
        Print #fNr, ""
        Print #fNr, "-------------------------------------------------------------------"
        Print #fNr, ""
    Else
        Open frmLibrary.dlgCommon.FileName For Output As #fNr
    End If
    Print #fNr, "' >>>>> Declarations <<<<<<"
    Print #fNr, txtDeclarations.Text
    Print #fNr, ""
    Print #fNr, "' >>>>> Procedure/Code <<<<<<"
    Print #fNr, txtCode.Text
    Close #fNr
Case Else:
    '============================================================
    ' See if user clicked on a specific attachment in dropdown & if not, exit sub at this point
    '============================================================
    If Left(ButtonMenu.Key, 6) <> "RecID:" Then Exit Sub
    DBrecID = myRecRef                                                  ' set reference this records db ID
    GP = Val(Mid(ButtonMenu.Key, 7)) & "|" & Caption    ' set GP variable to include the form's caption &
                                                                                                ' the specific Attachment's db ID
    frmAttachments.Show 1, frmLibrary                              ' show the Attachments form
    If GP = "Reload" Then LoadAttachmentListing               ' reload attachments if needed
    Exit Sub
End Select
GP = Null
Select Case iKey
    Case 1: txtCodeName.SetFocus
    Case 2: txtPurpose.SetFocus
    Case 3: txtDeclarations.SetFocus
    Case 4: txtCode.SetFocus
End Select
UserCnx:
Exit Sub

Sub_tbSub_ButtonMenuClick_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If fNr Then Close #fNr
If MsgBox("Error " & Err.Number & " - Procedure [Sub tbSub_ButtonMenuClick]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub txtCode_Change()
'============================================================
' If  the code section changed, flag as updated
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_txtCode_Change_General_ErrTrap_by_LaVolpe
If bUpdatable Then
    bDataChanged = True
    lLastLine(1) = txtCode.SelStart
    bColorCheck(1) = True
End If

Exit Sub

Sub_txtCode_Change_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub txtCode_Change]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub txtCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then ShowRightClick 1
End Sub

Private Sub txtCode_SelChange()
'===================================================================
' After user moves cursor to another line, re-color any changed text if needed
'===================================================================
If bColorCheck(1) = False Then Exit Sub
RecolorText 1
End Sub

Private Sub txtCodeName_Change()
'============================================================
' If the code title/name changed, flag as updated
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_txtCodeName_Change_General_ErrTrap_by_LaVolpe
If bUpdatable Then bDataChanged = True
Exit Sub

Sub_txtCodeName_Change_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub txtCodeName_Change]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub txtCodeName_GotFocus()
'============================================================
' When the code title/name field has the focus, select all the text in the field
'============================================================
On Error Resume Next
With txtCodeName
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodeName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then ShowRightClick 2
End Sub

Private Sub txtDeclarations_Change()
'============================================================
' If the declarations section changed, flag as updated
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_txtDeclarations_Change_General_ErrTrap_by_LaVolpe
If bUpdatable Then
    bDataChanged = True
    lLastLine(0) = txtDeclarations.SelStart
    bColorCheck(0) = True
End If
Exit Sub

Sub_txtDeclarations_Change_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub txtDeclarations_Change]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub txtDeclarations_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then ShowRightClick 0
End Sub

Private Sub txtDeclarations_SelChange()
'===================================================================
' After user moves cursor to another line, re-color any changed text if needed
'===================================================================
If bColorCheck(0) = False Then Exit Sub
RecolorText 0
End Sub

Private Sub RecolorText(txtBoxID As Integer, Optional bColorAll As Boolean = False)
On Error GoTo Sub_txtProcedures_SelChange_General_ErrTrap_by_LaVolpe
Dim lFirstChar As Long, lThisLine As Long, lLastChar As Long, lPrevLine As Long, lThisChar As Long
Dim iStep As Integer, lSelStart As Long, txtRange(0 To 1) As Variant, txtObj As RichTextBox
If txtBoxID = 0 Then Set txtObj = txtDeclarations Else Set txtObj = txtCode
With txtObj
    If bColorAll = False Then
        lThisLine = .GetLineFromChar(.SelStart)       ' get current line number
        lPrevLine = .GetLineFromChar(lLastLine(txtBoxID)) ' get line number when text change made
        If lPrevLine = lThisLine Then Exit Sub              ' if on the same line, exit for now
    End If
    bUpdatable = False                                  ' set flag to prevent premature flagging as updated
    If bColorAll = False Then
        lSelStart = .SelStart                                       ' save cursor position for later repositioning
        For lFirstChar = lLastLine(txtBoxID) To 1 Step -1         ' changed lines & changed text
            ' find the first character of the changed line
            If Mid(.Text, lFirstChar, 1) = vbLf Or Mid(.Text, lFirstChar, 1) = vbCr Then Exit For
        Next
        '  Determine where the last character of the line is. Function returns the nr of characters on the line
        lLastChar = lFirstChar + apiSendMessage(.hWnd, EM_LINELENGTH, lFirstChar, 0) + 1
        txtRange(0) = lFirstChar + 1                    ' Identify the first visible character on the line
        txtRange(1) = lLastChar - txtRange(0)      ' Identify how many visible characters on the line
        ' recolor line using the range identified above
        CheckLine4KeyWords txtObj, frmLibrary.rtfStaging, txtRange
    Else
        CheckLine4KeyWords txtObj, frmLibrary.rtfStaging
    End If
    bColorCheck(txtBoxID) = False
    bUpdatable = True
    .SelStart = lSelStart
End With
Exit Sub

Sub_txtProcedures_SelChange_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub txtProcedures_SelChange]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub txtPurpose_Change()
'============================================================
' If the purpose section changed, flag as updated
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_txtPurpose_Change_General_ErrTrap_by_LaVolpe
If bUpdatable Then bDataChanged = True
Exit Sub

Sub_txtPurpose_Change_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub txtPurpose_Change]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub txtPurpose_GotFocus()
'============================================================
' When the purpose field has the focus, select all the text in the field
'============================================================
On Error Resume Next
With txtPurpose
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub ShowRightClick(objIndex As Integer)
Dim I As Integer
With frmLibrary.mnuViewSub
    Select Case objIndex
    Case 0, 1:  ' Declarations & Code
        For I = 0 To 15
            If .Item(I).Caption <> "-" Then .Item(I).Enabled = True
        Next
        .Item(4).Checked = (cmdWordWrap(objIndex).BackColor = 65535)
        .Item(4).Enabled = True
    Case 2, 3:  ' Title, Purpose
        For I = 1 To 8
            If .Item(I).Caption <> "-" Then .Item(I).Enabled = False
        Next
        .Item(0).Enabled = True
        .Item(4).Enabled = False
        .Item(4).Checked = True
    Case 4, 5:  ' Categories, Languages
        For I = 0 To 8
            If .Item(I).Caption <> "-" Then .Item(I).Enabled = False
        Next
        .Item(4).Enabled = False
        .Item(4).Checked = False
    End Select
    .Item(11).Enabled = (Tag <> "NewRecord")
    frmLibrary.mnuViewRecord.Tag = ""
End With
PopupMenu frmLibrary.mnuViewRecord
Dim txtControl As Object
I = Val(frmLibrary.mnuViewRecord.Tag)
Select Case I
Case 1: ' Paste
    On Error GoTo FailedPaste
    Set txtControl = Choose(objIndex + 1, txtDeclarations, txtCode, txtCodeName, txtPurpose)
    txtControl.SelText = Clipboard.GetText & " "
Case 3: ' Font
    Call cmdFont_Click(objIndex)
Case 4: ' Wordwrap
    Call cmdWordWrap_Click(objIndex)
Case 5: ' Recolor text
    RecolorText objIndex, True
Case 7, 8:  ' Find New
    If I = 7 Then
        Call cmdSearch_Click(objIndex)
    Else
        Call cmdSearch_Click(Choose(objIndex + 1, 20, 21))
    End If
Case 10: ' Save
    SaveChanges
Case 11: ' Save As...
    Call tbSub_ButtonClick(tbSub.Buttons(4))
Case 12:    ' Save as File
    Call tbSub_ButtonMenuClick(tbSub.Buttons(3).ButtonMenus(2))
Case 14:    ' Edit Categories
    Call cmdAddCat_Click
Case 15:    ' Edit Languages
    cmdLanguageAdd_Click
Case 30, 31
    On Error GoTo FailedPaste
    Set txtControl = Choose(objIndex + 1, txtDeclarations, txtCode)
    If I = 31 Then
        txtControl.HideSelection = True
        txtControl.SelStart = 0
        txtControl.SelLength = Len(txtControl.Text)
    End If
    Clipboard.Clear
    Clipboard.SetText txtControl.SelText, vbCFText
    If I = 31 Then txtControl.SelLength = 0
    txtControl.HideSelection = False
End Select
Set txtControl = Nothing
Exit Sub
FailedPaste:
MsgBox "Failed to access the clipboard. Sorry.", vbInformation + vbOKOnly
Exit Sub
End Sub

Private Sub txtPurpose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then ShowRightClick 3
End Sub
