VERSION 5.00
Begin VB.Form frmClipboard 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add item to clipboard"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5160
   Icon            =   "frmClipboard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00EFEFEF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2025
      Width           =   255
   End
   Begin VB.CheckBox ckKeyLiteral 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Use keyboard literal"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1530
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2070
      Width           =   2040
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      Height          =   500
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2340
      Width           =   2550
   End
   Begin VB.CommandButton cmdPlaceOnClipboard 
      BackColor       =   &H00FF8080&
      Caption         =   "&Place on clipboard"
      Enabled         =   0   'False
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2340
      Width           =   2550
   End
   Begin VB.TextBox txtClipboard 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H002F2F2F&
      Height          =   2000
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      ToolTipText     =   "place text to add to clipboard here"
      Top             =   45
      Width           =   5160
   End
End
Attribute VB_Name = "frmClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Enum DT
      ClipboardData
      CodeBlockData
End Enum


Public enumDataType  As DT
 

Private Sub cmdCancel_Click()

       Unload Me
End Sub

Private Sub cmdHelp_Click()
'
     'this not on top
      SetWindowPos _
             hwnd, 1, 0, 0, 0, 0, &H1 Or &H2
             
             'msgbox shows help
               MsgBox _
"with this checked the literal interpretation  " & vbCrLf & _
"of your keypresses is used." & vbCrLf & vbCrLf & _
"For example, with this checked, if you type" & vbCrLf & _
"username {tab} password {Enter}" & vbCrLf & _
"when you paste from the clipboard at " & vbCrLf & _
"a future time your username will be typed," & vbCrLf & _
"the cursor will tab to the next item," & vbCrLf & _
"the password will be typed, then the enter" & vbCrLf & _
"key pressed." & vbCrLf & _
"Without this checked, username will be typed" & vbCrLf & _
"followed be whatever spacing is produced by" & vbCrLf & _
"the tab key, password typed, and the enter " & vbCrLf & _
"key will have no effect." & vbCrLf & _
"works for  ENTER   and   TAB   keys only.", _
              vbInformation
              
              'this back on top
              Call Form_Load

End Sub

Private Sub cmdPlaceOnClipboard_Click()
'
   With frmBar
       If enumDataType = ClipboardData Then
          .cClipboard.AddToClipboard Clip_Board, txtClipboard, frmBar.mnuArrClipboard
       Else
          .cClipboard.AddToClipboard Code_Block, txtClipboard, frmBar.mnuArrCode
       End If
   End With
   
   Unload Me
End Sub
 
 
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Private Sub Form_Load()
     'this on top
      SetWindowPos _
             hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
      
      'set the proper forms caption
      If enumDataType = ClipboardData Then
           Caption = "Add item to clipboard"
           cmdPlaceOnClipboard.Caption = "Place onto clipboard"
      Else
           Caption = "Add item to code block"
           cmdPlaceOnClipboard.Caption = "Place in codeblock library"
      End If
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Private Sub Form_Paint()

       txtClipboard.SetFocus
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |disable cmd if txtbox is empty..enable if not
'----------------------------------------------------------------------
Private Sub txtClipboard_Change()
        '
        If Len(txtClipboard) > 0 Then
             cmdPlaceOnClipboard.Enabled = True
        Else
             cmdPlaceOnClipboard.Enabled = False
        End If
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Private Sub txtClipboard_KeyDown(KeyCode As Integer, Shift As Integer)
 '

         If ckKeyLiteral.Value = vbChecked Then
             Dim str$
             
             Select Case KeyCode
                   Case Is = 13 'enter
                         str = vbCrLf & Chr(129) & "ENTER" & Chr(129)
                         GoTo DONEXT:
                   Case Is = 9 'tab
                         str = vbCrLf & Chr(129) & "TAB" & Chr(129) & vbCrLf
                         GoTo DONEXT:
             End Select
Exit Sub
DONEXT:
             'place cursor at end of string sent
             txtClipboard = (txtClipboard & str & "    ")
             txtClipboard.SelStart = Len(txtClipboard)
             
             'compensate for tab key by sending home
             If KeyCode = 9 Then
                  SendKeys "{HOME}"
             End If
         End If
End Sub
 
 


