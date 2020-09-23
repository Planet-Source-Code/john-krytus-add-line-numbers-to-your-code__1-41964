VERSION 5.00
Begin VB.Form frmFiles 
   Caption         =   "Files"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program Name = LineNumbers
'Programmer = John Krytus  Email:  jkrytus@cris.com
'Program date = December 2002
'


Option Explicit

Private Sub Dir1_Change()
    'connect to the selected file
    File1.Pattern = "*.frm;*.frp"
    File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()
    'connect to the selected drive
    Dir1.Path = Drive1.Drive
    
End Sub

Private Sub File1_Click()
    
    'format the filename
    On Error GoTo FileError
    If Right(Dir1.Path, 1) = "\" Then
        gstrFileName = File1.Path & File1.FileName
    Else
        gstrFileName = File1.Path & "\" & File1.FileName
    End If
    
    'place the source filename into the text combo box
    frmLineNumbers.cboSourceFile.Text = gstrFileName
    
    'add the source filename to the combo box
    frmLineNumbers.cboSourceFile.AddItem gstrFileName
    
    'Clear the destination box
    frmLineNumbers.txtDestination = ""
    frmLineNumbers.lblMessage = ""
    
    frmFiles.Hide
    frmLineNumbers.Show
    
    Exit Sub
    
FileError:
    MsgBox "FileError!"
End Sub
