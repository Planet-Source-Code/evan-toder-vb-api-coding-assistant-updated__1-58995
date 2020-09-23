VERSION 5.00
Begin VB.Form frmWaitForClipboardRemove 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   Icon            =   "frmWaitForClipboardRemove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&X"
      Height          =   270
      Left            =   2295
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   675
      Width           =   315
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Now Click Clipboard Item You Wish To Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   870
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   2490
   End
End
Attribute VB_Name = "frmWaitForClipboardRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

      frmBar.bRemoveMenuMode = True
      Unload Me
End Sub

Private Sub Form_Load()
     'this on top
      SetWindowPos _
             hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub
