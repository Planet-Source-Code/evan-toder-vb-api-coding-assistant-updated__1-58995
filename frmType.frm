VERSION 5.00
Begin VB.Form frmType 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select the desired type(s) then press OK"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3705
   Icon            =   "frmType.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      Height          =   500
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3285
      Width           =   1740
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF8080&
      Caption         =   "&OK"
      Height          =   500
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3285
      Width           =   1740
   End
   Begin VB.OptionButton optScope 
      Caption         =   "Private"
      Height          =   330
      Index           =   0
      Left            =   855
      TabIndex        =   2
      Top             =   2880
      Width           =   960
   End
   Begin VB.OptionButton optScope 
      Caption         =   "Public"
      Height          =   330
      Index           =   1
      Left            =   1845
      TabIndex        =   1
      Top             =   2880
      Value           =   -1  'True
      Width           =   960
   End
   Begin VB.ListBox lstType 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   45
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   45
      Width           =   3615
   End
End
Attribute VB_Name = "frmType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | gets the selected types selected from listbox
'             to be sent to our app
'----------------------------------------------------------------------
Private Sub cmdOK_Click()
 Dim i                   As Integer
 Dim strToSend           As String
 Dim strScope            As String

      With lstType
         If .SelCount > 0 Then
            'specify the scope by which of the
            '2 opt buttons is selected
            If optScope(0).Value = True Then
               strScope = "Private "
            Else
               strScope = "Public "
            End If
            'loop through and get the selected items
            For i = 0 To (.ListCount - 1)
            
             'add those items to the string to post to clipboard
                If .Selected(i) = True Then
                    'with types only add scope to the "Type"
                    If LCase(Trim(Left(.List(i), 5))) = "type" Then
                        strToSend = (strToSend & strScope & .List(i)) & vbCrLf
                    Else
                        strToSend = (strToSend & .List(i)) & vbCrLf
                    End If
                End If
             Next i
          End If
       End With
     
       modPub.bSendingApiNotConst = True
       
        'types to clipboard
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



Private Sub cmdCancel_Click()

       Unload Me
End Sub


 


Private Sub ListSelMultiple(ListBox As ListBox, strTopMatch As String, strBottomMatch As String)
  Dim lenMatch(1)              As Integer
  Dim i                        As Integer
  Dim bSel                     As Boolean
         
         'store len of matching strings
         lenMatch(0) = Len(strTopMatch)
         lenMatch(1) = Len(strBottomMatch)
         
         'boolean val indicated whether current clicked
         'item has just been selected or deselected
         bSel = ListBox.Selected(ListBox.ListIndex)

         With ListBox
                'if the item we just selected is blank
                ' or empty the deselect it and exit
                If Trim(.List(.ListIndex)) = "" Then
                    .Selected(.ListIndex) = False
                    Exit Sub
                End If
                
                'start from curr selected and work backwards
                For i = .ListIndex To 0 Step -1
                    'make the sel state of this item same
                    'as first item clicked
                    .Selected(i) = bSel
                    'if the top item were looking for is a match
                    'exit to stop the selection process
                    If Trim(LCase(Left(.List(i), lenMatch(0)))) = Trim(LCase(strTopMatch)) Then
                        Exit For
                    End If
                Next i
                
                'here were doing the same as the prev loop
                'except were working from the sel item DOWN
                For i = (.ListIndex) To (.ListCount - 1)
                    .Selected(i) = bSel
                    'weve found the bottom match so exit
                    If Trim(LCase(Left(.List(i), lenMatch(1)))) = Trim(LCase(strBottomMatch)) Then
                         Exit For
                    End If
                Next i
         End With
End Sub
 
 
 

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | we need to select from the opening to the ending type
'----------------------------------------------------------------------
 
Private Sub lstType_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
       'make multiselection "seamless"
       Call LockWindowUpdate(lstType.hwnd)
       Call ListSelMultiple(lstType, "Type", "End Type")
       Call LockWindowUpdate(0)
End Sub
