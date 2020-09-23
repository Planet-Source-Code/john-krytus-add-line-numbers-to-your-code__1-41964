VERSION 5.00
Begin VB.Form frmViewOutput 
   Caption         =   "View Output File"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   4598
      TabIndex        =   1
      Top             =   8280
      Width           =   1215
   End
   Begin VB.TextBox txtOutput 
      Height          =   8055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "frmViewOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program Name = LineNumbers
'Programmer = John Krytus  Email:  jkrytus@cris.com
'Program date = December 2002
'

Option Explicit
Private Sub cmdClose_Click()
    'get back to main form
    frmViewOutput.Hide
    frmLineNumbers.Show
    
End Sub

Private Sub Form_Load()
    'load the text box with file data
    txtOutput.Text = gstrTextBoxData
End Sub
