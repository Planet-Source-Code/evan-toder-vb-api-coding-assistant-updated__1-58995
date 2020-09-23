VERSION 5.00
Begin VB.Form FrmComment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Your procedure header comment"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5250
   Icon            =   "FrmComment.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FF8080&
      Caption         =   "&Save"
      Height          =   300
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1755
      Width           =   2235
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1755
      Width           =   2235
   End
   Begin VB.TextBox txtComment 
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
      ForeColor       =   &H00008000&
      Height          =   1635
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   45
      Width           =   5175
   End
End
Attribute VB_Name = "FrmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

       Unload Me
End Sub

Private Sub cmdSave_Click()
       
       'save the comment to reg
       SaveSetting _
           "VB_ClipboardBuddy", _
           "Comment", _
           "Value", _
           txtComment.Text
       'update the variable that holds the comment header
       frmBar.strHeaderComment = txtComment.Text
       Unload Me
End Sub

Private Sub Form_Load()
     'this on top
      SetWindowPos _
             hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
            
       'get the text for txtcomment from reg
       txtComment = frmBar.strHeaderComment
End Sub
