VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H80000018&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3480
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4335
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


 
 
 

  
 

Private Sub Form_Load()
   
   'place this window on top
   Call SetWindowPos( _
                  -1, 0, 0, 0, 0, 0, _
                    &H1 Or &H2)
End Sub
