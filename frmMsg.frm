VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9645
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerClose 
      Enabled         =   0   'False
      Left            =   5940
      Top             =   450
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


 

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Private Sub Form_Load()
     FrmMsgIsLoaded = True
     '
     'auto closes this in 7 seconds
     TimerClose.Interval = 7000
     TimerClose.Enabled = True
End Sub
 
 
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |kill timer
'----------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
        TimerClose.Interval = 0
        TimerClose.Enabled = False
        FrmMsgIsLoaded = False
End Sub
 
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |timer closes this form
'----------------------------------------------------------------------
Private Sub TimerClose_Timer()
'
        Unload Me
End Sub
