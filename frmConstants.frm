VERSION 5.00
Begin VB.Form frmConstants 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select the desired constant & click  ""OK"""
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4965
   Icon            =   "frmConstants.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optScope 
      Caption         =   "Public"
      Height          =   330
      Index           =   1
      Left            =   2475
      TabIndex        =   4
      Top             =   5130
      Value           =   -1  'True
      Width           =   960
   End
   Begin VB.OptionButton optScope 
      Caption         =   "Private"
      Height          =   330
      Index           =   0
      Left            =   1485
      TabIndex        =   3
      Top             =   5130
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF8080&
      Caption         =   "&OK"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5490
      Width           =   2460
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      Height          =   500
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5490
      Width           =   2460
   End
   Begin VB.ListBox lstConstants 
      Appearance      =   0  'Flat
      Height          =   5100
      ItemData        =   "frmConstants.frx":08CA
      Left            =   90
      List            =   "frmConstants.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4830
   End
End
Attribute VB_Name = "frmConstants"
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
' COMMENTS: | Places selected constants from the listbox to variable
'             strToSend and sends them to our app
'----------------------------------------------------------------------
Private Sub cmdOK_Click()
 Dim i                   As Integer
 Dim strToSend           As String
 Dim strScope            As String

      With lstConstants
         If .SelCount > 0 Then
            'specify the scope by which of the
            '2 opt buttons is selected
            If optScope(0).Value = True Then
               strScope = "Private Const "
            Else
               strScope = "Public Const "
            End If
            'loop through and get the selected items
            For i = 0 To (.ListCount - 1)
                If .Selected(i) = True Then
                    'add those items to the string to post to clipboard
                    strToSend = (strToSend & strScope & .List(i)) & vbCrLf
                End If
             Next i
          End If
       End With
 
       modPub.bSendingApiNotConst = True
       
       Clipboard.Clear
       Clipboard.SetText strToSend$
       
       'show msg telling user he has to hit
       'ctrl+V to send constant or type
       Set cScreenText = New clsScreenText
       cScreenText.ScreenMsg "To paste, press " _
                             & Chr(34) & "Ctrl + V" & Chr(34), _
                             20, vbBlue, True, 2000, 3000
       Set cScreenText = Nothing
       Unload Me
End Sub

Private Sub Form_Load()

     'this on top
      SetWindowPos _
             hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub
