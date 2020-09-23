VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Renanmer"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   2535
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   2040
      Width           =   5415
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000010&
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   4920
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rename"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1320
      Width           =   5415
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   5415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   360
      Width           =   5415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   3600
      Left            =   120
      TabIndex        =   0
      Top             =   3800
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Files"
      Height          =   195
      Left            =   7560
      TabIndex        =   12
      Top             =   1800
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Error(s):"
      Height          =   195
      Left            =   2520
      TabIndex        =   11
      Top             =   4680
      Width           =   540
   End
   Begin VB.Label Label3 
      Caption         =   "Status"
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Replace with:"
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Find text in FileName:"
      Height          =   195
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Err:
Text1.Text = "Selected Directory: " & Dir1.Path
Text4.Text = ""
Dim FileLen, FindTxt, Before, After, OldName, NewName

'start loop
For a = 0 To File1.ListCount - 1

' Count total number of characters in the selected file
FileLen = Len(File1.List(a))

'Check whether the selected file contains the Find text
FindTxt = InStr(1, File1.List(a), Text2.Text)

'If Selected file contains the given find text then rename it
If FindTxt <> 0 Then

'part of file name before the given find text
Before = Left(File1.List(a), FindTxt - 1)

'part of file name after the given find text
After = Right(File1.List(a), FileLen - Len(Text2.Text) - Len(Before))

'Old file name
OldName = Dir1.Path & "\" & File1.List(a)

'Nem file name
NewName = Dir1.Path & "\" & Before & Text3.Text & After

'Rename the files in the selected folder
Name OldName As NewName

'Write in the status box when file name is changed
Text1.Text = Text1.Text & vbCrLf & a & ": File Rename As: " & Before & Text3.Text & After

Else

'If Selected file does not contain the given find text then write in error box
Text4.Text = Text4.Text & vbCrLf & a & ": File Can not Rename: " & File1.List(a)

End If

Next

'Rename completed
Text1.Text = Text1.Text & vbCrLf & "Renaming Successfully Completed..."

'When there is no error found.
If Text4.Text = "" Then
Text4.Text = "No error found. "
End If
File1.Refresh
perr:
Screen.MousePointer = vbDefault
Exit Sub

Err:
Text4.Text = Err.Description
Resume perr:

End Sub

Private Sub Dir1_Change()
'display the files contain in the selected folder.
File1.Path = Dir1.Path
' Count total number of files in the selected folder
Label5.Caption = File1.ListCount & " File(s) in the selcted folder"
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Label5.Caption = File1.ListCount & " File(s) in the selcted folder"
End Sub
