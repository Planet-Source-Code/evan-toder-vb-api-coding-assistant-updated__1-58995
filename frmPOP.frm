VERSION 5.00
Begin VB.Form frmPOP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Email configuration"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
   Icon            =   "frmPOP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboPOP 
      Height          =   315
      Left            =   3510
      TabIndex        =   10
      Top             =   1620
      Width           =   1140
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancel"
      Height          =   300
      Index           =   1
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2295
      Width           =   2235
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   2385
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2295
      Width           =   2235
   End
   Begin VB.Frame fmePOP 
      Caption         =   "Account info"
      ForeColor       =   &H00FF0000&
      Height          =   1500
      Left            =   1530
      TabIndex        =   1
      Top             =   90
      Width           =   3165
      Begin VB.TextBox txtPOP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   945
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   990
         Width           =   2130
      End
      Begin VB.TextBox txtPOP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   945
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   675
         Width           =   2130
      End
      Begin VB.TextBox txtPOP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   945
         MultiLine       =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "i.e.  pop.snet.net"
         Top             =   360
         Width           =   2130
      End
      Begin VB.Label lblPOP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "password"
         ForeColor       =   &H00FF8080&
         Height          =   210
         Index           =   4
         Left            =   45
         TabIndex        =   12
         Top             =   1035
         Width           =   840
      End
      Begin VB.Label lblPOP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "username"
         ForeColor       =   &H00FF8080&
         Height          =   210
         Index           =   3
         Left            =   45
         TabIndex        =   11
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblPOP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "mail server address "
         ForeColor       =   &H00FF8080&
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   270
         Width           =   930
      End
   End
   Begin VB.ListBox lstAccount 
      Height          =   1035
      Left            =   45
      TabIndex        =   0
      Top             =   810
      Width           =   1410
   End
   Begin VB.CheckBox ckDisablePopCHecking 
      Caption         =   "&disable POP checking"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   90
      TabIndex        =   13
      Top             =   1890
      Width           =   1590
   End
   Begin VB.Label lblPOP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check for new mail every:"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   2
      Left            =   1575
      TabIndex        =   9
      Top             =   1665
      Width           =   1935
   End
   Begin VB.Label lblPOP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "this service checks if email exists without downloading it"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   840
      Index           =   1
      Left            =   -90
      TabIndex        =   8
      Top             =   0
      Width           =   1620
   End
End
Attribute VB_Name = "frmPOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private iCurrIndex         As Integer
Private bIsDirty           As Boolean

 
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |This means the user is/may be changing mail checking
'            frequency so we need to save/updata
'----------------------------------------------------------------------
Private Sub cboPOP_Click()

       cmd(0).Enabled = True
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |DISABLE/ENABLE ALL CONTROLS ON FORM BASED ON CHECKBOX VAL
'----------------------------------------------------------------------
Private Sub ckDisablePopCHecking_Click()
  Dim CTL        As Control
  Dim b          As Boolean
  
  On Error GoTo CTLNEXT:
      'loop through all controls
      For Each CTL In Controls
               CTL.Enabled = (ckDisablePopCHecking.Value - 1)
CTLNEXT: 'if ctl doesnt support this property then resumes to next ctl
      Next CTL
      
      ckDisablePopCHecking.Enabled = True
      cmd(0).Enabled = False
      'updata public(modPub) variable(saved to file)
      modPub.bEnablePOPchecking = (ckDisablePopCHecking.Value - 1)
      
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Private Sub cmd_Click(Index As Integer)

       If Index = 0 Then
            'save data entered in textboxes to modPub.POPaccountInfo
            'the index being the current index in listbox with focus
            modPub.POPaccountInfo(0, iCurrIndex) = txtPOP(0)
            modPub.POPaccountInfo(1, iCurrIndex) = txtPOP(1)
            modPub.POPaccountInfo(2, iCurrIndex) = txtPOP(2)
            modPub.MailCheckFrequency = cboPOP.ListIndex
            cmd(0).Enabled = False
            'in case cboPOP.listindex(mail check frequ) has change
            'call the this will update the label and check if its
            'timer to check mail based upon new mailCheckFrequ
            Call frmBar.timerMailCheck_Timer
       Else
            Unload Me
       End If
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Private Sub Form_Load()

      'add 4 account names to list box
      With lstAccount
             .AddItem "Account #1"
             .AddItem "Account #2"
             .AddItem "Account #3"
             .AddItem "Account #4"
             .ListIndex = 0
      End With
      
      'add mail checking frequ to cboPOP
      With cboPOP
           .AddItem "5 minutes"
           .AddItem "15 minutes"
           .AddItem "30 minutes"
           .AddItem "60 minutes"
      End With
      
      'check "Disable POP checking if the global
      If bEnablePOPchecking = False Then
          Me.ckDisablePopCHecking.Value = vbChecked
      Else
          Me.ckDisablePopCHecking.Value = vbUnchecked
      End If
      
      cboPOP.ListIndex = MailCheckFrequency
      
      'color any nonfilled textboxes the light blue color
      'as a visual cue to user to fill in the values
      Call txtPOP_KeyDown(0, 0, 0)
      
     'this on top
      SetWindowPos _
             hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Private Sub lstAccount_Click()
       Caption = "Click the Save button for each account created"
       iCurrIndex = lstAccount.ListIndex
      'retrieve data from modPub.POPaccountInfo
      'index being the same as listindex clicked on
      'and place values in the textboxes
       txtPOP(0) = modPub.POPaccountInfo(0, iCurrIndex)
       txtPOP(1) = modPub.POPaccountInfo(1, iCurrIndex)
       txtPOP(2) = modPub.POPaccountInfo(2, iCurrIndex)
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Private Sub txtPOP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim i                As Integer
  Dim bMT              As Boolean
  
       cmd(0).Enabled = False
       'enable the save button of all the
       'textboxes have values
       For i = 0 To 2
          If Len(Trim(txtPOP(i))) = 0 Then
               txtPOP(i).BackColor = RGB(220, 220, 255)
               cmd(0).Enabled = False
               bMT = True
          Else
               txtPOP(i).BackColor = vbWhite
          End If
       Next i
          
       cmd(0).Enabled = Not bMT
End Sub


