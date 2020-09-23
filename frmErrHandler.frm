VERSION 5.00
Begin VB.Form frmErrHandler 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Your Err handling code"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5205
   Icon            =   "frmErrHandler.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtERR 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1500
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   45
      Width           =   5115
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1575
      Width           =   2235
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FF8080&
      Caption         =   "&Save"
      Height          =   300
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "show mail sender"
      Top             =   1575
      Width           =   2235
   End
End
Attribute VB_Name = "frmErrHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

       Unload Me
End Sub

Private Sub cmdSave_Click()
       
       'save the Errhandler to reg
       SaveSetting _
           "VB_ClipboardBuddy", _
           "ErrHandler", _
           "Value", _
           txtERR.Text
       'update the variable that holds the errhandler
       frmBar.strErrHandler = txtERR.Text
       Unload Me
End Sub

Private Sub Form_Load()
     'this on top
      SetWindowPos _
             hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
            
       'get the text for txtErr from reg
        txtERR.Text = frmBar.strErrHandler
End Sub

