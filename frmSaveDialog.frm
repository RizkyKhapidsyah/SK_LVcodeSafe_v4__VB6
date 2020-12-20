VERSION 5.00
Begin VB.Form frmSaveDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Overwrite Existing File?"
   ClientHeight    =   2445
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ControlBox      =   0   'False
   HelpContextID   =   7
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      Caption         =   "Yes To ALL"
      Height          =   375
      Index           =   3
      Left            =   3810
      TabIndex        =   4
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Yes"
      Height          =   375
      Index           =   2
      Left            =   2580
      TabIndex        =   3
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "No To ALL"
      Height          =   375
      Index           =   1
      Left            =   1350
      TabIndex        =   2
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   1980
      Width           =   855
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "No"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label lblFile 
      Caption         =   "File Name"
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
      Height          =   345
      Left            =   300
      TabIndex        =   10
      Top             =   480
      Width           =   5355
   End
   Begin VB.Label Label1 
      Caption         =   "Press YES TO ALL to replace all existing files"
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   5715
   End
   Begin VB.Label Label1 
      Caption         =   "Press YES to overwrite this file && continue recieving alerts to possible overwrites"
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1410
      Width           =   5715
   End
   Begin VB.Label Label1 
      Caption         =   "Press NO TO ALL to prevent any existing files from being overwritten"
      Height          =   345
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1170
      Width           =   5715
   End
   Begin VB.Label Label1 
      Caption         =   "Press NO to skip this file and continue recieving alerts to possible overwrites"
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   930
      Width           =   5715
   End
   Begin VB.Label Label1 
      Caption         =   "The following file already exists in the target folder."
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   195
      Width           =   5715
   End
End
Attribute VB_Name = "frmSaveDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
' Inserted by LaVolpe
On Error GoTo Sub_CancelButton_Click_General_ErrTrap_by_LaVolpe
GP = "Cancel"
Unload Me
Exit Sub

Sub_CancelButton_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub CancelButton_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub Form_Load()
' Inserted by LaVolpe
On Error GoTo Sub_Form_Load_General_ErrTrap_by_LaVolpe
lblFile.Caption = GP
GP = Null
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

Private Sub Form_Unload(Cancel As Integer)
' Inserted by LaVolpe
On Error GoTo Sub_Form_Unload_General_ErrTrap_by_LaVolpe
If IsNull(GP) Then
    MsgBox "Select one of the buttons please", vbInformation + vbOKOnly
    Cancel = 1
End If
Exit Sub

Sub_Form_Unload_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub Form_Unload]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub

Private Sub OKButton_Click(Index As Integer)
' Inserted by LaVolpe
On Error GoTo Sub_OKButton_Click_General_ErrTrap_by_LaVolpe
GP = OKButton(Index).Caption
Unload Me
Exit Sub

Sub_OKButton_Click_General_ErrTrap_by_LaVolpe:    ' Inserted by Lavolpe
If MsgBox("Error " & Err.Number & " - Procedure [Sub OKButton_Click]" & vbCrLf & Err.Description, vbExclamation + vbRetryCancel) = vbRetry Then Resume
End Sub
