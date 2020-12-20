VERSION 5.00
Begin VB.Form frmClipboard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi-Code Clipboard"
   ClientHeight    =   2820
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5085
   Icon            =   "frmClipboard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMultiDrag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   90
      ScaleHeight     =   105
      ScaleWidth      =   4890
      TabIndex        =   2
      ToolTipText     =   "Drag Tool"
      Top             =   1560
      Width           =   4920
   End
   Begin VB.ListBox lstMemory 
      Height          =   1230
      Left            =   75
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   285
      Width           =   4950
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Double click bar to clear the clipboard."
      Height          =   225
      Index           =   1
      Left            =   135
      TabIndex        =   4
      Top             =   2595
      Width           =   4860
   End
   Begin VB.Label Label2 
      Caption         =   $"frmClipboard.frx":000C
      Height          =   885
      Left            =   75
      TabIndex        =   3
      Top             =   1710
      Width           =   4995
   End
   Begin VB.Label Label1 
      Caption         =   "Drag and drop the selected code item(s)"
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   4860
   End
   Begin VB.Menu mnuCopyPaste 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuCopyPasteSub 
         Caption         =   "Drop Declarations Here"
         Index           =   0
      End
      Begin VB.Menu mnuCopyPasteSub 
         Caption         =   "Drop ALL Declarations"
         Index           =   1
      End
      Begin VB.Menu mnuCopyPasteSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCopyPasteSub 
         Caption         =   "Drop Code/Procedure Here"
         Index           =   3
      End
      Begin VB.Menu mnuCopyPasteSub 
         Caption         =   "Drop ALL Code/Procedures"
         Index           =   4
      End
      Begin VB.Menu mnuCopyPasteSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCopyPasteSub 
         Caption         =   "Cancel"
         Index           =   6
      End
   End
End
Attribute VB_Name = "frmClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private rstMemory As DAO.Recordset

Private Sub FormOntop(bOnTop As Boolean)
'Make a form not always ontop of other windows, parameters...

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo FormOntop_General_ErrTrap

Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Dim lOnTop As Long
If bOnTop = False Then lOnTop = -2 Else lOnTop = -1
On Error GoTo Error
Call SetWindowPos(hWnd, lOnTop, 0&, 0&, 0&, 0&, Flags)
Exit Sub
Error:  MsgBox Err.Description, vbExclamation, "Error"
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
FormOntop_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: FormOntop" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub Form_Activate()

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo Form_Activate_General_ErrTrap

FormOntop True
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
Form_Activate_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: Form_Activate" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub Form_Load()

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo Form_Load_General_ErrTrap
Icon = frmLibrary.SmallImages.ListImages(4).ExtractIcon
Tag = "frmClipboard"
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
Form_Load_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: Form_Load" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub lstMemory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo lstMemory_MouseMove_General_ErrTrap

If Button = vbLeftButton Then lstMemory.OLEDrag
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
lstMemory_MouseMove_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: lstMemory_MouseMove" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub lstMemory_OLESetData(Data As DataObject, DataFormat As Integer)
Dim I As Integer, lChunkSize As Long, sfldName As String, Looper As Long
Dim destWin As Long, MouseXY As POINTAPI

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo lstMemory_OLESetData_General_ErrTrap

mnuCopyPasteSub(1).Enabled = lstMemory.SelCount > 1
mnuCopyPasteSub(4).Enabled = lstMemory.SelCount > 1

mnuCopyPaste.Tag = 6
GetCursorPos MouseXY
destWin = WindowFromPoint(MouseXY.X, MouseXY.Y)
PopupMenu mnuCopyPaste
On Error GoTo PasteError
I = Val(mnuCopyPaste.Tag)
If I = 6 Then Err.Raise -2

Set rstMemory = mainDB.OpenRecordset("Select IDnr, Code, Declarations FROM tblSourceCode;", dbOpenDynaset)
If rstMemory.RecordCount = 0 Then
    MsgBox "No records exists in your database.", vbInformation + vbOKOnly
    rstMemory.Close
    Err.Raise -1
End If
Dim vClipBoard As Variant
If I < 3 Then sfldName = "Declarations" Else sfldName = "Code"
GoSub SetClipboardData
Data.SetData vClipBoard
Clipboard.Clear
If I < 3 Then sfldName = "Code" Else sfldName = "Declarations"
GoSub SetClipboardData
Clipboard.SetText vClipBoard
SetForegroundWindow destWin
CleanUp:
rstMemory.Close
Exit Sub

SetClipboardData:
vClipBoard = Null
For Looper = 0 To lstMemory.ListCount - 1
    If lstMemory.Selected(Looper) = True Then
        With rstMemory
            .FindFirst "[IDnr]=" & lstMemory.ItemData(Looper)
            If .NoMatch = False Then
                vClipBoard = vClipBoard & .Fields(sfldName) & vbCrLf
            End If
        End With
    End If
Next
Return
Exit Sub

PasteError:
Data.SetData "", vbCFText
Select Case Err.Number
Case -1:    Resume CleanUp
Case -2:
Case Else:  MsgBox "Error: " & Err.Description, vbExclamation + vbOKOnly
End Select
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
lstMemory_OLESetData_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: lstMemory_OLESetData" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub lstMemory_OLEStartDrag(Data As DataObject, AllowedEffects As Long)

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo lstMemory_OLEStartDrag_General_ErrTrap

AllowedEffects = vbDropEffectCopy
Data.Clear
Data.SetData , vbCFText
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
lstMemory_OLEStartDrag_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: lstMemory_OLEStartDrag" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub mnuCopyPasteSub_Click(Index As Integer)

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo mnuCopyPasteSub_Click_General_ErrTrap

mnuCopyPaste.Tag = Index
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
mnuCopyPasteSub_Click_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: mnuCopyPasteSub_Click" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub picMultiDrag_DblClick()

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo picMultiDrag_DblClick_General_ErrTrap

Clipboard.Clear
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
picMultiDrag_DblClick_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: picMultiDrag_DblClick" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub

Private Sub picMultiDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo picMultiDrag_MouseMove_General_ErrTrap

If Button = vbLeftButton Then
    If lstMemory.SelCount = 0 Then
        FormOntop False
        MsgBox "First select at least one code item from the listing.", vbInformation + vbOKOnly
        FormOntop True
    Else
        lstMemory.OLEDrag
    End If
End If
Exit Sub

' Inserted by LaVolpe OnError Insertion Program.
picMultiDrag_MouseMove_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: picMultiDrag_MouseMove" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Sub
