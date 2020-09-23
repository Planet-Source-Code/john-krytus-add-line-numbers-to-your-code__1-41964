Attribute VB_Name = "Module1"
'Program Name = LineNumbers
'Programmer = John Krytus  Email:  jkrytus@cris.com
'Program date = December 2002
'

'use the option explicit to help the programmer
Option Explicit

'create a global variable so that the source filename is always available
Public gstrFileName As String
Public gstrTextBoxData As String
Public gintHelpQuestion As Integer


Public Sub Main()
    'Load forms for faster program operation
    Load frmFiles
    Load frmLineNumbers
    
    'after the line number is executed, go on to the line number form
    frmLineNumbers.Show
End Sub
