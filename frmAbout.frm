VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3120
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2835
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6120
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5400
         Top             =   240
      End
      Begin VB.Label lblClose 
         AutoSize        =   -1  'True
         Caption         =   "Click here to close"
         Height          =   195
         Left            =   4560
         TabIndex        =   5
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "jkrytus@cris.com"
         Height          =   195
         Left            =   2400
         TabIndex        =   4
         Top             =   2183
         Width           =   1200
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   1703
         Width           =   1275
      End
      Begin VB.Label lblAuthor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "By: J. Krytus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   2
         Top             =   1223
         Width           =   1935
      End
      Begin VB.Label cmdProgram 
         AutoSize        =   -1  'True
         Caption         =   "Line Numbers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   743
         Width           =   2400
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program Name = LineNumbers
'Programmer = John Krytus  Email:  jkrytus@cris.com
'Program date = December 2002
'

Option Explicit
Dim intSplashShow

Private Sub Form_KeyPress(KeyAscii As Integer)
    'close the about form if the user hits a key
    Unload Me
    Timer1.Enabled = False
    frmLineNumbers.Show
End Sub
Private Sub Form_Load()
    'start the display time timer
    Timer1.Enabled = True
End Sub

Private Sub Frame1_Click()
    'close the about form if the user clicks the frame
    Unload Me
    Timer1.Enabled = False
    frmLineNumbers.Show
End Sub
Private Sub lblClose_Click()
    'Close the About screen if user clicks on close label
    Timer1.Enabled = False
    frmAbout.Hide
    frmLineNumbers.Show
End Sub
Private Sub Timer1_Timer()
    'during the time the about form is displayed, check every
    'second to see if it's time to close.  Otherwise increase
    'the SplashShow display variable by one
    If intSplashShow >= 5 Then
        Timer1.Enabled = False
        frmAbout.Hide
        frmLineNumbers.Show
    End If
    intSplashShow = intSplashShow + 1
End Sub
