VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLibSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Locator"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   HelpContextID   =   6
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   5400
   Begin MSComctlLib.ListView lvResults 
      Height          =   1965
      Left            =   165
      TabIndex        =   6
      Top             =   2295
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   3466
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code Name"
         Object.Width           =   8467
      EndProperty
   End
   Begin VB.CheckBox chkAttach 
      Height          =   190
      Left            =   420
      TabIndex        =   16
      Top             =   1710
      Width           =   210
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View Selected Code"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   4305
      Width           =   2145
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   180
      TabIndex        =   8
      Top             =   4305
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4050
      TabIndex        =   10
      Top             =   4305
      Width           =   1245
   End
   Begin VB.CheckBox chkPartial 
      Height          =   190
      Left            =   435
      TabIndex        =   3
      Top             =   1080
      Width           =   210
   End
   Begin VB.CheckBox chkCriteria 
      Height          =   190
      Index           =   1
      Left            =   165
      TabIndex        =   5
      Top             =   1485
      Width           =   210
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CheckBox chkCriteria 
      Height          =   190
      Index           =   0
      Left            =   165
      TabIndex        =   2
      Top             =   870
      Width           =   210
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Index           =   0
      Left            =   2400
      TabIndex        =   4
      Top             =   825
      Width           =   2895
   End
   Begin VB.CheckBox chkCriteria 
      Height          =   190
      Index           =   2
      Left            =   165
      TabIndex        =   0
      Top             =   330
      Width           =   210
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   315
      ItemData        =   "frmLibSearch.frx":0000
      Left            =   2400
      List            =   "frmLibSearch.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   270
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Attachments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   660
      TabIndex        =   17
      Top             =   1710
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search results - icon means criteria found in attachment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   2100
      Width           =   5160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exact match only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   675
      TabIndex        =   14
      Top             =   1080
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "In code/declarations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   435
      TabIndex        =   13
      Top             =   1485
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "In the Title/Purpose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   435
      TabIndex        =   12
      Top             =   870
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Having this language"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   435
      TabIndex        =   11
      Top             =   330
      Width           =   1995
   End
End
Attribute VB_Name = "frmLibSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAttach_Click()
If chkAttach.Value = 1 Then chkCriteria(1).Value = 1
End Sub

Private Sub chkPartial_Click()
If chkPartial.Value Then chkCriteria(0).Value = 1
End Sub

Private Sub cmdFind_Click()
Dim I As Integer, strSQL As String, findRS As DAO.Recordset, lOffset As Long
' Inserted by LaVolpe
On Error GoTo Sub_cmdFind_Click_General_ErrTrap_by_LaVolpe
For I = 0 To chkCriteria.UBound
    If chkCriteria(I) = 1 Then
        If I = 2 Then
            If cboLanguage.ListIndex > -1 Then Exit For Else chkCriteria(I) = 0
        Else
            If Len(txtName(I)) Then Exit For Else chkCriteria(I) = 0
        End If
    End If
Next
If I > chkCriteria.UBound Then
    MsgBox "At least one of the search options must be selected.", vbInformation + vbOKOnly
    Exit Sub
End If

On Error Resume Next
Dim strLanguage As String, strCode(0 To 1) As String
Dim strSearch As String, fBytes() As Byte
Dim bAnd As Boolean, bMatchOK As Boolean, itmX As ListItem
Dim sTmp As String, lFldSize As Long

lvResults.ListItems.Clear
strLanguage = "(tblCodeLangXref.LangID = " & "%%%" & ")"
strCode(0) = "((tblSourceCode.CodeName Like " & Chr(34) & "*%%%*" & Chr(34) & ") OR " & _
    "(tblSourceCode.Purpose Like " & Chr(34) & "*%%%*" & Chr(34) & "))"
strCode(1) = "((tblSourceCode.Code Like " & Chr(34) & "*%%%*" & Chr(34) & ") OR " & _
    "(tblSourceCode.Declarations Like " & Chr(34) & "*%%%*" & Chr(34) & "))"

If chkCriteria(2) = 1 And Len(cboLanguage) Then
    strSearch = strSearch & Replace(strLanguage, "%%%", cboLanguage.ItemData(cboLanguage.ListIndex))
    bAnd = True
End If
For I = 0 To 1
    If chkCriteria(I) = 1 And Len(txtName(I)) Then
        If bAnd = True Then strSearch = strSearch & " AND "
        sTmp = Replace(txtName(I), Chr(34), Chr(34) & Chr(34))
        sTmp = Replace(sTmp, "'", "''")
        sTmp = Replace(strCode(I), "%%%", sTmp)
        If chkPartial = 1 And I = 0 Then         ' exact matches only
            sTmp = Replace(sTmp, "Like ", "=")
            sTmp = Replace(sTmp, "*", "")
        End If
    strSearch = strSearch & sTmp
    bAnd = True
    End If
Next
strSearch = " WHERE (" & strSearch & ");"
sTmp = "SELECT IDnr, CodeName FROM tblSourceCode"
If chkCriteria(2) = 1 Then
    sTmp = sTmp & " INNER JOIN tblCodeLangXref  ON tblCodeLangXref.CodeID = tblSourceCode.IDnr"
End If
strSearch = sTmp & strSearch

Set findRS = mainDB.OpenRecordset(strSearch, dbOpenDynaset)
If findRS.RecordCount Then
    bMatchOK = True
    findRS.MoveFirst
    Do While findRS.EOF = False
        Set itmX = lvResults.ListItems.Add(, "Rec#" & findRS.Fields("IDnr"), findRS.Fields("CodeName"))
        findRS.MoveNext
    Loop
End If
If chkAttach = 1 And Len(txtName(1).Text) > 0 Then
    strSearch = "SELECT tblAttachments.Attachment, tblAttachments.Description, tblSourceCode.IDnr, tblSourceCode.CodeName " & _
        "FROM tblAttachments INNER JOIN tblSourceCode ON tblAttachments.RecIDRef = tblSourceCode.IDnr;"
    Set findRS = mainDB.OpenRecordset(strSearch, dbOpenDynaset)
    With findRS
        If .RecordCount Then
            bMatchOK = True
            .MoveFirst
            Do While .EOF = False
                lOffset = 0
                lFldSize = .Fields("Attachment").FieldSize
                Do While lOffset < lFldSize
                    ' here we chunk out the binary field around 32K per chunk
                    fBytes = .Fields("Attachment").GetChunk(lOffset, 32000)
                    ' we now convert bytes to string for comparison
                    If InStr(LCase(StrConv(fBytes(), vbUnicode)), LCase(txtName(1).Text)) Then
                        ' this would normally cause an error except for the Resume Next above
                        ' by letting the error go, we include all attachments where search found
                        ' matches in one line of the listview.
                        Set itmX = lvResults.ListItems.Add(, "Rec#" & findRS.Fields("IDnr"), findRS.Fields("CodeName"))
                        itmX.Text = itmX.Text & " {" & .Fields("Description") & "}"
                        itmX.SmallIcon = 13
                        Exit Do
                    End If
                    lOffset = lOffset + 32000
                Loop
                .MoveNext
            Loop
        End If
    End With
End If
findRS.Close
Set findRS = Nothing
If bMatchOK Then
    MsgBox "Your search is complete.", vbInformation + vbOKOnly, "Search Complete"
Else
    MsgBox "No matches found for that criteria.", vbInformation + vbOKOnly, "No Matches"
End If
Exit Sub

Sub_cmdFind_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdFind_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub cmdView_Click()
' Inserted by LaVolpe
On Error GoTo Sub_cmdView_Click_General_ErrTrap_by_LaVolpe
If lvResults.SelectedItem Is Nothing Then Exit Sub
frmLibrary.ShowCode "RecID:" & Mid(lvResults.SelectedItem.Key, 5)
Me.SetFocus
Exit Sub

Sub_cmdView_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub cmdView_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Command1_Click()
' Inserted by LaVolpe
On Error GoTo Sub_Command1_Click_General_ErrTrap_by_LaVolpe
WindowState = 1
Exit Sub

Sub_Command1_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Command1_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Load()
Dim I As Integer
' Inserted by LaVolpe
On Error GoTo Sub_Form_Load_General_ErrTrap_by_LaVolpe
cboLanguage.Clear
Set lvResults.SmallIcons = frmLibrary.SmallImages
With frmLibrary.cboFilter(2)
    For I = 1 To .ListCount - 1
        cboLanguage.AddItem .List(I)
        cboLanguage.ItemData(I - 1) = .ItemData(I)
        If .ItemData(I) = MyDefaults.Language Then cboLanguage.ListIndex = I - 1
    Next
End With
If cboLanguage.ListIndex < 0 And cboLanguage.ListCount Then cboLanguage.ListIndex = 0
Icon = frmLibrary.SmallImages.ListImages(6).ExtractIcon
DoGradient Me, 2
For I = Label1.LBound To Label1.UBound: Label1(I).ForeColor = MyDefaults.LblColorPopup: Next
Exit Sub

Sub_Form_Load_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Load]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Terminate()
' Inserted by LaVolpe
On Error GoTo Sub_Form_Terminate_General_ErrTrap_by_LaVolpe
Unload Me
Exit Sub

Sub_Form_Terminate_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Terminate]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Label1_Click(Index As Integer)
' Inserted by LaVolpe
On Error GoTo Sub_Label1_Click_General_ErrTrap_by_LaVolpe
Select Case Index
Case 0, 1, 2
    If chkCriteria(Index) = 1 Then chkCriteria(Index) = 0 Else chkCriteria(Index) = 1
Case 3
    If chkPartial = 1 Then chkPartial = 0 Else chkPartial = 1
Case 5
    If chkAttach = 1 Then chkAttach = 0 Else chkAttach = 1
Case Else
End Select
Exit Sub

Sub_Label1_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Label1_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub lvResults_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Not lvResults.SelectedItem Is Nothing Then Call cmdView_Click
End Sub
