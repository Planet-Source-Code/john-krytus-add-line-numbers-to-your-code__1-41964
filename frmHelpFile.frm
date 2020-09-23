VERSION 5.00
Begin VB.Form frmHelpFile 
   BackColor       =   &H80000018&
   Caption         =   "Help File"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Label lblQuestion1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Q1. Is there a quick start procedure provided?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblQuestion4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Q4. What does the Destination Output do?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   3030
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Close This Window"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   4080
      Width           =   1395
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblQuestion3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Q3. What does the source input do?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2580
   End
   Begin VB.Label lblQuestion2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Q2. What does this program do?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblHelpfileLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Line Number Help File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   158
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmHelpFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program Name = LineNumbers
'Programmer = John Krytus  Email:  jkrytus@cris.com
'Program date = December 2002
'

Option Explicit


Private Sub lblClose_Click()
    Unload frmHelpFile
    frmLineNumbers.Show
End Sub

Private Sub lblQuestion1_Click()
    'answer the question 'What does the destination box do?'
    gintHelpQuestion = 1
    frmHelpFile.Hide
    frmHelpAnswer.Show
End Sub

Private Sub lblQuestion2_Click()
    'answer the question 'What this program does
    gintHelpQuestion = 2
    frmHelpFile.Hide
    frmHelpAnswer.Show
End Sub

Private Sub lblQuestion3_Click()
  'answer the question 'What does the source input do?
    gintHelpQuestion = 3
    frmHelpFile.Hide
    frmHelpAnswer.Show
End Sub

Private Sub lblQuestion4_Click()
    'answer the question 'What does the destination box do?'
    gintHelpQuestion = 4
    frmHelpFile.Hide
    frmHelpAnswer.Show
End Sub
