VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCodeView 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "To send only part of code..highlight that part"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6075
   Icon            =   "frmCodeView.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtCodeView 
      Height          =   3165
      Left            =   0
      TabIndex        =   2
      Top             =   45
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   5583
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmCodeView.frx":08CA
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      Height          =   500
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3285
      Width           =   2460
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF8080&
      Caption         =   "&Paste Code"
      Height          =   500
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3285
      Width           =   2460
   End
End
Attribute VB_Name = "frmCodeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

       Unload Me
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | sends code block related to moving controls on the form
'             user can send only part of code by highlighting that part
'----------------------------------------------------------------------
Private Sub cmdOK_Click()
       
        With txtCodeView
           If Len(.SelText) <> 0 Then
               Call frmBar.SendData(.SelText)
           Else
               Call frmBar.SendData(.Text)
           End If
        End With
       
        Unload Me
End Sub

 

Private Sub Form_Load()
     'this on top
      SetWindowPos _
             hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | search through the rich text control looking for lines that
'             are comments and color them the standard green comment color
'----------------------------------------------------------------------
Private Sub ColorCodeText()
 Dim foundPos               As Integer
 Dim foundCR                As Integer
 Dim startPos               As Integer
 
      With txtCodeView
        .Locked = False
        'start point for search is very beginning of textbox
        startPos = 0
        'continuing searching till not found anymore
        While foundPos <> -1
           'looking for "'"
            foundPos = .Find("'", startPos, Len(.Text))
           
            If foundPos <> -1 Then
                 .SelStart = foundPos
                  foundCR = .Find(vbCrLf, (foundPos + 1), Len(.Text))
                
                .SelStart = foundPos
                .SelLength = (foundCR - foundPos)
                .SelColor = RGB(0, 130, 0)
             End If
           
             DoEvents
            'foundpos returns the position in the textbox of the "'"
            startPos = (foundCR + 1)
         Wend
         
         .SelStart = 0
         .Locked = True
      End With
End Sub
 
 

Private Sub Form_Paint()
       
        Call ColorCodeText
End Sub

Private Sub txtCodeView_KeyPress(KeyAscii As Integer)

       KeyAscii = 0
End Sub
