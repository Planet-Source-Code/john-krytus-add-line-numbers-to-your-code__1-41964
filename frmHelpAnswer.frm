VERSION 5.00
Begin VB.Form frmHelpAnswer 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmHelpAnswer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblQuestionNumber 
      BackColor       =   &H80000018&
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Close This Window"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   3960
      Width           =   1395
   End
   Begin VB.Label lblHelpAnswer 
      BackColor       =   &H80000018&
      ForeColor       =   &H00808000&
      Height          =   2055
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label lblHelpQuestion 
      BackColor       =   &H80000018&
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
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   6045
   End
   Begin VB.Line Line1 
      X1              =   743
      X2              =   6623
      Y1              =   840
      Y2              =   840
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
      Left            =   803
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmHelpAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program Name = LineNumbers
'Programmer = John Krytus  Email:  jkrytus@cris.com
'Program date = December 2002
'

Option Explicit

Private Sub Form_GotFocus()
    'display question and answer for the help question selected
    
    If gintHelpQuestion = 1 Then
    'Answer Question 'Show a quick start'
    lblHelpQuestion = "Q1. Is there a Quick Start procedure provided"
    lblQuestionNumber = "A1."
    lblHelpAnswer = "Yes. " & vbCrLf & _
                    "Step #1 - select a source VB .frm file to process. Either " & _
                    "type in path and                         filename or " & _
                    "browse for the file." & vbCrLf & _
                    "Step #2 - Click the analyze button." & vbCrLf & _
                    "Step #3 - If everything has gone well, the output " & _
                    "file name is placed in the                     " & _
                    "destination box. The " & _
                    "file is now processed and available for " & _
                    "                           viewing."
    End If
    If gintHelpQuestion = 2 Then
    'Answer Question 'What does this program do?'
    lblHelpQuestion = "Q2. What this program does."
    lblQuestionNumber = "A2."
    lblHelpAnswer = "This program reads in a Visual Basic .frm file " & _
                    "and looks for the start of the vb code. Upon " & _
                    "finding it, starts adding line numbers to each " & _
                    "line of code as it writes the newly modified " & _
                    "code to the specified output file."
    End If
    If gintHelpQuestion = 3 Then
    'Answer Question 'What does the source input do?
    lblHelpQuestion = "Q3. What does the source input do?"
    lblQuestionNumber = "A3."
    lblHelpAnswer = "This input box will accept the path and filename of the " & _
                    "souce file to be processed.  In VB, the code for " & _
                    "the project is in the .frm file right after the " & _
                    "control properties.  A browse button is provided " & _
                    "to search for the form file needed." & vbCrLf & vbCrLf & _
                    "Also, a drop down is provided to select previously " & _
                    "processed files."
    End If
    If gintHelpQuestion = 4 Then
    'Answer Question 'What does the destination output do?'
    lblHelpQuestion = "Q4. What does the destination output do?"
    lblQuestionNumber = "A4."
    lblHelpAnswer = "This input box will accept the path and filename of the " & _
                    "destination file to write the output to. "
    End If
End Sub

Private Sub lblClose_Click()
    'go back to help screen
    frmHelpAnswer.Hide
    frmHelpFile.Show
End Sub
