VERSION 5.00
Begin VB.Form frmLineNumbers 
   Caption         =   "Line Numbers"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   315
      Left            =   7440
      TabIndex        =   10
      ToolTipText     =   "Click this button to view the output file"
      Top             =   840
      Width           =   1020
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   495
      Left            =   4417
      TabIndex        =   9
      ToolTipText     =   "Click this button to get help"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   3075
      TabIndex        =   8
      ToolTipText     =   "Click here to clear source file choices in dropdown"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "Browse"
      Height          =   315
      Left            =   7440
      TabIndex        =   7
      ToolTipText     =   "Click this button to browse for a source .frm file"
      Top             =   360
      Width           =   1020
   End
   Begin VB.ComboBox cboSourceFile 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      ToolTipText     =   "Enter a path and filename to process"
      Top             =   360
      Width           =   6255
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      ToolTipText     =   "This button when clicked will exit the program"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtDestination 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Enter a path and filename for the output file"
      Top             =   840
      Width           =   6255
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "&Analyze"
      Height          =   495
      Left            =   1733
      TabIndex        =   0
      ToolTipText     =   "This button when clicked with process the source file"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblOutputStatus 
      AutoSize        =   -1  'True
      Caption         =   "Output Status"
      Height          =   195
      Left            =   2280
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8640
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblDestinationLabel 
      AutoSize        =   -1  'True
      Caption         =   "Destination"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   795
   End
   Begin VB.Label lblSourceFileLabel 
      AutoSize        =   -1  'True
      Caption         =   "Source"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   510
   End
   Begin VB.Label lblMessage 
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAnalyze 
         Caption         =   "&Analyze"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditBrowse 
         Caption         =   "&Browse for source file"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "&Clear souce file list"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "frmLineNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program Name = LineNumbers
'Programmer = John Krytus  Email:  jkrytus@cris.com
'Program date = December 2002
'

'use the option explicit to help the programmer
Option Explicit
Private Sub cboSourceFile_Click()
    'assign the selected path & filename to the global filename variable
    gstrFileName = cboSourceFile.List(cboSourceFile.ListIndex)
End Sub
Private Sub cmdAnalyze_Click()
    Dim strPastFilenames As String
    Dim intIndex As Integer
    Dim strLineData As String
    Dim intNewLine As Integer
    Dim strNewFilename As String
    Dim intFound As Integer
    
    'make sure you've got a file to process
    If gstrFileName = "" Then
        MsgBox "No file to process", vbInformation, "File Name Error"
        cboSourceFile.SetFocus
        Exit Sub
    End If
    
    'Open .frm file
    Open gstrFileName For Input As #1
    
    strNewFilename = Left(gstrFileName, Len(gstrFileName) - 3) & "frp"
    txtDestination.Text = strNewFilename
    Open strNewFilename For Output As #2
        
        'read from file until code is found
        'code might start at option explicit or attribute VB_Exposed
        intFound = 0
        Do Until EOF(1)
            Line Input #1, strLineData
            If strLineData = "Option Explicit" Then
                intFound = 1
                Exit Do
            End If
        Loop
        'if option explicit is found, don't look any further
        If intFound = 0 Then
            'rewind the souce file to look again
            Close #1
            Open gstrFileName For Input As #1
            Do Until EOF(1)
                Line Input #1, strLineData
                If InStr(strLineData, "Attribute VB_Exposed") > 0 Then
                    Line Input #1, strLineData
                    Exit Do
                End If
            Loop
        End If
        'print out a title for the output
        Print #2, intNewLine; Tab(10); strLineData
        
        intNewLine = 1
        Do Until EOF(1)
            Line Input #1, strLineData
            Print #2, intNewLine; Tab(10); strLineData
            intNewLine = intNewLine + 1
        Loop
    Close #2
    Close #1
    
    'write out the output status
    lblMessage = intNewLine & " Output Lines written"
    
    'Save the current file source name
    Open App.Path & "\PastFilenames.dat" For Output As #1
            For intIndex = 0 To cboSourceFile.ListCount - 1
                strPastFilenames = cboSourceFile.List(intIndex)
                Print #1, strPastFilenames
            Next
    Close #1
    
    'refresh the displayed files for next time
    frmFiles.File1.Refresh
    
       
End Sub
Private Sub cmdbrowse_Click()
    'Get the path & filename of the source file
    frmLineNumbers.Hide
    frmFiles.Show
End Sub
Private Sub cmdClear_Click()
    'delete the combo box entries
    Kill (App.Path & "\PastFilenames.dat")
End Sub
Private Sub cmdExit_Click()
    
    'end the program
    Unload frmLineNumbers
    Unload frmFiles
    Unload frmHelpFile
    Unload frmViewOutput
    Unload frmHelpAnswer
    Unload frmAbout
    End
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'end the program
    Unload frmLineNumbers
    Unload frmFiles
    Unload frmHelpFile
    Unload frmViewOutput
    Unload frmHelpAnswer
    Unload frmAbout
    End
End Sub
Private Sub mnuFileExit_Click()
    'end the program
    Unload frmLineNumbers
    Unload frmFiles
    Unload frmHelpFile
    Unload frmViewOutput
    Unload frmHelpAnswer
    Unload frmAbout
    End
End Sub
Private Sub mnuEditClear_Click()
    'delete the combo box entries
    Kill (App.Path & "\PastFilenames.dat")
End Sub
Private Sub cmdHelp_Click()
    'display help form
    frmLineNumbers.Hide
    frmHelpFile.Show

End Sub

Private Sub cmdView_Click()
    'view the destination file
    'create variables to use below
    Dim strMyFile As String
    Dim strFileData As String
            
    'check if text box has a filename in it
    If txtDestination = "" Then
        MsgBox "No file to view", vbInformation, "File Name Error"
        txtDestination.SetFocus
        Exit Sub
    End If

    'check to see if the file name is valid and file is there
    strMyFile = Dir(txtDestination)
    If strMyFile = "" Then
        MsgBox "File does not exist", vbInformation, "File Access Error"
        txtDestination.SetFocus
        Exit Sub
    End If

    'read in the output file
    Open txtDestination For Input As #1
        gstrTextBoxData = "File Name = " & txtDestination & vbCrLf
        Do While Not EOF(1)
            Line Input #1, strFileData
            gstrTextBoxData = gstrTextBoxData & vbCrLf & strFileData
        Loop
    Close #1
    frmLineNumbers.Hide
    frmViewOutput.Show
    
End Sub

Private Sub Form_Load()
    
    Dim strPastFilenames As String
    
    'set up error handler for no file
    On Error GoTo NoFile
    
    'get past file name jobs
    If cboSourceFile.ListCount <> -1 Then
        Open App.Path & "\PastFilenames.dat" For Input As #1
            Do While Not EOF(1)
                Line Input #1, strPastFilenames
                cboSourceFile.AddItem strPastFilenames
            Loop
        Close #1
    End If

NoFile:
    Close #1
    
End Sub

Private Sub mnuEditBrowse_Click()
    'Get the path & filename of the source file
    frmLineNumbers.Hide
    frmFiles.Show
End Sub

Private Sub mnuHelpAbout_Click()
    'show the about screen
    frmLineNumbers.Hide
    frmAbout.Show
End Sub

Private Sub mnuHelpHelp_Click()
    'display help form
    frmLineNumbers.Hide
    frmHelpFile.Show
    
End Sub
