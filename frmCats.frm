VERSION 5.00
Begin VB.Form frmCats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit "
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   4020
      Width           =   1245
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Edit "
      Height          =   555
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   1740
      Width           =   1185
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete "
      Height          =   555
      Left            =   4200
      TabIndex        =   2
      Top             =   1050
      Width           =   1185
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "New "
      Height          =   555
      Index           =   0
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   1185
   End
   Begin VB.ListBox lstCats 
      Height          =   4155
      Left            =   135
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   345
      Width           =   3885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Make any change as needed (changes are immediate)"
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
      Left            =   195
      TabIndex        =   5
      Top             =   90
      Width           =   5145
   End
End
Attribute VB_Name = "frmCats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private rsCats As DAO.Recordset, bReload As Boolean
Private sToggle(0 To 3) As String, iToggle As Integer
'============================================================
' This form handles both adding, editing & deleting Categories & Languages
'============================================================

Private Sub cmdAdd_Click(Index As Integer)
'============================================================
' Function to add or edit an existing record
'============================================================
On Error GoTo FailedUpdate
Dim sCatName As String, i As Integer, sMsg As String, sHdr As String
If Index = 0 Then       ' Adding new record
    sHdr = "New "       '  these are used to format the inputbox displays
Else
    If lstCats.ListIndex < 0 Then   ' Editing a new record, but ensure an item is selected
        MsgBox "First select a " & sToggle(0) & " to edit", vbInformation + vbOKOnly
        Exit Sub
    End If
    sHdr = "Updated "
End If
' build initial inputbox message
sMsg = "Enter the " & LCase(sHdr) & sToggle(0) & " below (max of 150 characters)"
' display the inputbox
sCatName = InputBox(sMsg, sHdr & sToggle(2), Choose(Index + 1, "", lstCats))
If sCatName = "" Then Exit Sub          ' user pressed cancel
sCatName = Left(sCatName, 150)      ' use the 1st 150 characters

For i = 0 To lstCats.ListCount - 1      ' check for a duplicate entry
    If sCatName = lstCats.List(i) Then  ' oops, duplicate
        If i <> lstCats.ListIndex Then                    ' but if it's the same record we're editing, that's ok, otherwise...
            sMsg = "The " & sToggle(0) & " you provided below is an existing category. Do you want to duplicate it?"
            i = MsgBox(sMsg & vbCrLf & vbCrLf & sCatName, vbQuestion + vbYesNo + vbDefaultButton2, "Duplicate Entry")
            If i = vbNo Then Exit Sub       ' don't allow duplicates to be added
            Exit For
        End If
    End If
Next
With rsCats                                     ' now to make the change
    If Index Then                               ' editing, so find the specific db record
        .FindFirst "[ID] = " & lstCats.ItemData(lstCats.ListIndex)
        .MoveFirst
        .Edit
    Else
        .AddNew
    End If
    .Fields(sToggle(2)) = sCatName  ' update the db field, simple
    .Update
End With
bReload = True                                  ' changes made, so flag it
rsCats.Close
LoadCategories                                  ' reload the listboxes to display the changes
Exit Sub

FailedUpdate:
MsgBox Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub cmdClose_Click()
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_cmdClose_Click_General_ErrTrap_by_LaVolpe
Unload Me
'============================================================
Exit Sub

Sub_cmdClose_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdClose_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdDel_Click()
'============================================================
' Sub will delete the record from the database and screen
'============================================================

' Inserted by LaVolpe
On Error GoTo Sub_cmdDel_Click_General_ErrTrap_by_LaVolpe
If lstCats.ListIndex < 0 Then   ' If one wasn't selected, then exit here
    MsgBox "First select a " & sToggle(0) & " to delete", vbInformation + vbOKOnly
    Exit Sub
End If
Dim sMsg As String, strSQL As String, i As Integer
' provide a confirmation message first
sMsg = "Are you absolutely sure you want the following " & sToggle(0) & " deleted?" & vbCrLf & vbCrLf
i = MsgBox(sMsg & lstCats, vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If i = vbNo Then Exit Sub       ' user opted against deleting
' Otherwise construct the query string to delete the record from the database
strSQL = "DELETE tbl" & sToggle(3) & ".* From tbl" & sToggle(3) & _
    " WHERE (((tbl" & sToggle(3) & ".ID)=" & lstCats.ItemData(lstCats.ListIndex) & "));"
rsCats.Close                        ' close the db and delete the record from the attachments table
mainDB.Execute strSQL
' now build a query string to delete any references in the cross-reference table
strSQL = Choose(iToggle, "Cat", "Lang")
strSQL = "DELETE tblCode" & strSQL & "Xref.* From tblCode" & strSQL & "Xref " & _
    " WHERE (((tblCode" & strSQL & "Xref." & strSQL & "ID)=" & lstCats.ItemData(lstCats.ListIndex) & "));"
mainDB.Execute strSQL       ' delete any cross-references
LoadCategories                    ' reload the records
bReload = True                    ' flag that updates occurred
Exit Sub

Sub_cmdDel_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdDel_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Load()
'============================================================
' Sub displays the form & sets the toggles, depending on whether the form is being used to
'   modify Categories or Languages
'============================================================

' Inserted by LaVolpe
On Error GoTo Sub_Form_Load_General_ErrTrap_by_LaVolpe
If GP = "Cats" Then     ' Categories being modified
    iToggle = 1: Caption = Caption & "Organizational Categories"
    sToggle(0) = "category": sToggle(1) = "categories"
    sToggle(2) = "Category": sToggle(3) = "Categories"
Else                            ' Languages being modified
    iToggle = 2: Caption = Caption & "Code Languages"
    sToggle(0) = "language": sToggle(1) = "languages"
    sToggle(2) = "Language": sToggle(3) = "Language"
End If
HelpContextID = Choose(iToggle, 4, 5)
' relabel button captions appropriately
cmdAdd(0).Caption = "New " & sToggle(2)
cmdAdd(1).Caption = "Edit " & sToggle(2)
cmdDel.Caption = "Delete " & sToggle(2)
GP = Null   ' reset variable and load appropriate titlebar icon
Icon = frmLibrary.SmallImages.ListImages(25 + iToggle).ExtractIcon
LoadCategories  ' call function to open recordset and populate listbox
DoGradient Me, 2 ' repaint form and label colors as necessary
Label1.ForeColor = MyDefaults.LblColorPopup
bReload = False ' set update flag to false
Exit Sub

Sub_Form_Load_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Load]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub LoadCategories()
'============================================================
'   Sub will open the recordset shared by this form and the on-screen viewer.  Also will
'       populate the listbox with records
'============================================================
On Error GoTo FailedLoad
lstCats.Clear   ' Clear the listbox and open the recordset for edting
Set rsCats = mainDB.OpenRecordset("tbl" & sToggle(3), dbOpenDynaset)
If rsCats.RecordCount > 0 Then      ' loop thru each record, if any, and add it to the listbox
    With rsCats
        .MoveFirst
        Do While .EOF = False
            lstCats.AddItem .Fields(sToggle(2))     ' Add to listbox
            lstCats.ItemData(lstCats.NewIndex) = .Fields("ID")  ' track the db record ID
            .MoveNext
        Loop
    End With
End If
' disable the delete & edit buttons if no records exist, otherwise, enable them
cmdDel.Enabled = lstCats.ListCount
cmdAdd(1).Enabled = lstCats.ListCount
Exit Sub

FailedLoad:
MsgBox Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub Form_Terminate()
'============================================================
' Inserted by LaVolpe
On Error GoTo Sub_Form_Terminate_General_ErrTrap_by_LaVolpe
Unload Me
'============================================================
Exit Sub

Sub_Form_Terminate_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Terminate]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
'============================================================
'   Closes the recordset and calls function to repopulate the main window categories & langauges
'   combo boxes as necessary.  Those combo boxes are used by other forms also
'============================================================

On Error Resume Next
rsCats.Close
Set rsCats = Nothing
If bReload = True Then              ' Changes made
    If iToggle = 1 Then                 ' modifying categories
        RefreshCategories               ' function to repopulate main window categories
        LastCatUpdate = Now()       ' update when categories were last changed
    Else
        RefreshLanguages                ' modifying languages, so call function to repopulate main window
        LastLangUpdate = Now()      ' update when languages were last changed
    End If
End If
End Sub
