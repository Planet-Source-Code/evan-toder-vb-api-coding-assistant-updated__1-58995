VERSION 5.00
Begin VB.Form F 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   Icon            =   "F.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   -90
      ScaleHeight     =   270
      ScaleWidth      =   4860
      TabIndex        =   21
      Top             =   2385
      Width           =   4920
   End
   Begin VB.PictureBox picGuage 
      Height          =   1275
      Left            =   3960
      ScaleHeight     =   1215
      ScaleWidth      =   315
      TabIndex        =   19
      Top             =   1080
      Width           =   375
      Begin VB.Label lblGuage 
         BackColor       =   &H00000000&
         Height          =   150
         Left            =   0
         TabIndex        =   20
         Top             =   1080
         Width           =   420
      End
   End
   Begin RamPurge.Tray Tray1 
      Left            =   2475
      Top             =   450
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.Timer tmrRAM 
      Enabled         =   0   'False
      Left            =   3600
      Top             =   1935
   End
   Begin VB.CommandButton btnPurge 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Purge Now"
      Height          =   330
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   630
      Width           =   1140
   End
   Begin VB.CommandButton btnTray 
      BackColor       =   &H00FFFFFF&
      Caption         =   "T"
      Height          =   330
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "send to tray"
      Top             =   180
      Width           =   330
   End
   Begin VB.CommandButton btnDummy 
      Caption         =   "Command1"
      Height          =   420
      Left            =   5265
      TabIndex        =   13
      Top             =   2475
      Width           =   420
   End
   Begin VB.CommandButton btnClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      Height          =   330
      Left            =   4095
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "close & exit"
      Top             =   180
      Width           =   330
   End
   Begin VB.Frame fmeAgressiveness 
      BackColor       =   &H00FF8080&
      Caption         =   "RAM purge agrressiveness"
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   315
      TabIndex        =   1
      Top             =   180
      Width           =   2445
      Begin VB.OptionButton optAgressiveness 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   1395
         TabIndex        =   4
         Top             =   360
         Width           =   240
      End
      Begin VB.OptionButton optAgressiveness 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   1035
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   240
      End
      Begin VB.OptionButton optAgressiveness 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   195
         Index           =   0
         Left            =   675
         TabIndex        =   2
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblSlider 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   60
         Left            =   675
         TabIndex        =   8
         Top             =   405
         Width           =   915
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "minimal"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   7
         ToolTipText     =   "10% of used RAM is attempted to be freed"
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "agressive"
         Height          =   240
         Index           =   1
         Left            =   1620
         TabIndex        =   6
         ToolTipText     =   "20% of used RAM is attempted to be freed"
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "moderate"
         Height          =   240
         Index           =   2
         Left            =   810
         TabIndex        =   5
         ToolTipText     =   "15% of used RAM is attempted to be freed"
         Top             =   540
         Width           =   690
      End
   End
   Begin VB.Frame fmeOptions 
      BackColor       =   &H00FF8080&
      Caption         =   "Pruge RAM..."
      ForeColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   270
      TabIndex        =   0
      Top             =   990
      Width           =   3660
      Begin VB.OptionButton optPurgePercent 
         BackColor       =   &H00FF8080&
         Caption         =   "30%"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   18
         Top             =   1035
         Width           =   735
      End
      Begin VB.OptionButton optPurgePercent 
         BackColor       =   &H00FF8080&
         Caption         =   "20%"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   17
         Top             =   1035
         Width           =   735
      End
      Begin VB.OptionButton optPurgePercent 
         BackColor       =   &H00FF8080&
         Caption         =   "10%"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   16
         Top             =   1035
         Width           =   735
      End
      Begin VB.ComboBox cbo1 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   450
         TabIndex        =   11
         Top             =   450
         Width           =   1590
      End
      Begin VB.CheckBox ckOptions 
         BackColor       =   &H00FF8080&
         Caption         =   "when free RAM  drops below certain level"
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   810
         Width           =   3525
      End
      Begin VB.CheckBox ckOptions 
         BackColor       =   &H00FF8080&
         Caption         =   "at timed intervals"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label lblNextPurge 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   2250
         TabIndex        =   22
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   2760
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4470
   End
End
Attribute VB_Name = "F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'SCOPE  VAR NAME          TYPE           ALTERED BY
Private m_fFIle        As Integer '  get the next free file number for file access
Private m_FilePath     As String  '  used for loading/saving program settings
Private b_InTray       As Boolean '  clicking button to send to tray
Private m_DropPerc     As Single '   clicking on of the optPurgePercent opt buttons
Private m_percAggress  As Integer '  click one of the optAgressiveness opt buttons
Private m_TimerPasses  As Long '     used in timer to compare elapsed time to determine when next purge
Private m_PurgeTime    As Long '     combo box to select timed intervals of ram purge
Private m_FreeRAM      As Long '     amount of free ram in the system
Private m_TotRAM       As Long '     total ram in the system
Private m_UsedRAM      As Long '     used ram in the system
Private m_buff()       As Variant '  used to eat up memory in the loop in sub PurgeRam




' _ __      .    . __ _     .  '    .     .          .
'| '_ \_   _  _ _ / _` | ___   ' _ _  __ _  _ __ __
'| |_)  | | || '_\ (_| |/ _ \  '| '_\/ _` || '_ ` _ \
'| .__/ |_| || |  \__, |  __/  '| |   (_| || | | | | |
'|_|   \__,_||_|  |___/ \___|  '|_|  \__,_||_| |_| |_|


Private Sub PurgeRam(MBamount%)
On Error GoTo ERR:
'---------------------------------
'
'---------------------------------
'VARIABLES:
  Dim i%
'CODE:
  Me.AutoRedraw = False
  
  ReDim m_buff(MBamount)
  picGuage.ScaleHeight = MBamount
  'if the form is hidden, increasing
  'timers interval allows user to see
  'effects of ram purging while its happening
  If b_InTray = False Then
     tmrRAM.Interval = 500
  End If
 
  For i = 0 To (MBamount - 1)
      'this = exactly 1 megabyte
      'so we loop the number of
      'megabytes we want to purge
      'from memory
      m_buff(i) = Space(500000)
       'prevent guage flicker
      LockWindowUpdate Me.hWnd
      lblGuage.Top = (picGuage.ScaleHeight - i)
      lblGuage.Height = i
      DoEvents
      'allow the update
      LockWindowUpdate 0
      'update tray tooltip
      Tray1.TrayToolTip = "Purging RAM now!"
      'without calling this the updated tooltip
      'wont display
      Tray1.Show True
      DoEvents
   Next i
   
  'set timer back to normal
  If b_InTray = False Then
     tmrRAM.Interval = 1000
  Else
     tmrRAM.Interval = 6000
  End If
  '
  lblGuage.Height = 0
  lblGuage.Top = (picGuage.Top - picGuage.Height)
  Erase m_buff
  Beep
  
'END CODE:
'exit sub
ERR:
  Debug.Print ERR.Description
End Sub
 

Private Sub btnClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        'close form
        btnDummy.SetFocus
        Unload Me
        End
End Sub
Private Sub btnPurge_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ERR:
'---------------------------------
' 'purge ram now
'---------------------------------
'VARIABLES:
  
'CODE:
   btnDummy.SetFocus
   Call PurgeRam(DeterminPurgeAmount(m_UsedRAM, m_FreeRAM, m_percAggress))
'END CODE:
'exit sub
ERR:
  Debug.Print ERR.Description

End Sub
 
Private Sub cbo1_click()
        'we use seconds because we use api GetTickCount
        'to determine when to purge base upon time
        Select Case cbo1.ListIndex
           Case Is = 0
               m_PurgeTime = 300000    '5 minutes in seconds
               m_TimerPasses = GetTickCount
           Case Is = 1
               m_PurgeTime = 900000   '15 minutes in seconds
               m_TimerPasses = GetTickCount
           Case Is = 2
               m_PurgeTime = 1800000 ' 30 minutes in seconds
               m_TimerPasses = GetTickCount
        End Select
End Sub

Private Sub ckOptions_Click(Index As Integer)
        '
        Select Case Index
           Case Is = 0 'user checking option to purge ram
                       'at timed interval
              If ckOptions(0).Value = vbUnchecked Then
                   m_PurgeTime = 0
              End If
              
           Case Is = 1 'user checkin option for purging ram
                       'when drops below certain level
              If ckOptions(1).Value = vbChecked Then
                  'check 20%
                  optPurgePercent(1).Value = True
                  Call optPurgePercent_Click(1)
              Else 'uncheck all percent levels
                  Dim i
                  For i = 0 To 2
                     optPurgePercent(i).Value = False
                  Next i
              End If
        End Select
End Sub







'     .    .     .     '
'_____  _ _  __ _ _ _
'_   _|| '_\/ _` | |_| |
' | |  | |   (_| |\__, |
' |_|  |_|  \__,_||___/


Private Sub btnTray_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        'tray the form
        btnDummy.SetFocus
        'with app in tray it is not
        'as important to update ram
        'as frequently as when visible
        tmrRAM.Interval = 6000
        'send to tray
        Tray1.Show True
        Me.Visible = False
        b_InTray = True
End Sub

Private Sub optAgressiveness_Click(Index As Integer)
      'm_percAgrress if factored into the calc
      'of how much ram to purge
      'see:= DeterminPurgeAmount in modAPI
      Select Case Index
         Case Is = 0
           m_percAggress = 10
         Case Is = 1
           m_percAggress = 15
         Case Is = 2
           m_percAggress = 20
      End Select
End Sub

Private Sub optPurgePercent_Click(Index As Integer)
        'variable m_dropPerc is nothing
        'more than a storehouse for which
        'option button was last clicked
        'but this save the timer from
        'having to do that calculations
        Select Case Index
           Case Is = 0
              m_DropPerc = 0.1
           Case Is = 1
              m_DropPerc = 0.2
           Case Is = 2
              m_DropPerc = 0.3
        End Select
End Sub

Private Sub Tray1_LeftClick()
        'tray click..show form
        Tray1.Show False
        Me.Visible = True
        b_InTray = False
        tmrRAM.Interval = 1000
End Sub



''  _     .    .          .  '    . _ '  _     .    .     '    .    .    .
' / _| ___  _ _  _ __ __     ' _   (_) / _| ___  ___ _   _  ___  _    ___
'| |_ / _ \| '_\| '_ ` _ \   '| |  | || |_ / _ \/ __| |_| |/ __|| |  / _ \
'|  _| (_) | |  | | | | | |  '| |_ | ||  _|  __/ (__ \__, | (__ | |_   __/
'|_|  \___/|_|  |_| |_| |_|  '|___||_||_|  \___|\___||___/ \___||___|\___|

Private Sub Form_Load()
On Error GoTo ERR:
'VARIABLES:

'CODE:
        Sx = Screen.TwipsPerPixelX
        Sy = Screen.TwipsPerPixelY
        'shape form and button regions
        Call ModAPI.SetRgn(Me, 50)
        Call ModAPI.SetRgn(btnClose, 50)
        Call ModAPI.SetRgn(btnTray, 50)
        Call ModAPI.SetRgn(btnPurge, 70)
        'add selections to combo box
        With cbo1
           .AddItem "Every 5 minutes"
           .AddItem "Every 15 minutes"
           .AddItem "Every 30 minutes"
        End With
        'display of ram data
        picInfo.ForeColor = vbRed
        'acts as progress bar during ram purge
        lblGuage.Height = 0
        lblGuage.Top = (picGuage.Top - picGuage.Height)
        'load prog data
        Call PopulateVarAndControls
        'updates ram info
        tmrRAM.Interval = 1000
        tmrRAM.Enabled = True
        '
        DoEvents
        DoEvents
        Show
'END CODE:
Exit Sub
ERR:
  Debug.Print "Form_Load: " & ERR.Description
End Sub

Private Sub PopulateVarAndControls()
On Error GoTo ERR:
'---------------------------------
'here we load the programs data from file
'and select the appropriate controls settings
'based upon the values
'---------------------------------
'VARIABLES:
  Dim i%
  Dim sArr() As Variant
'CODE:
        'load data
        m_FilePath = App.Path & "\progSettings.ini"
        m_fFIle = FreeFile
        'loads sArr with array of progdata
        sArr = ModAPI.func_FileLoadData(m_FilePath)
        
        'fill progs values
        For i = LBound(sArr) To UBound(sArr)
             Select Case i
                Case Is = 0
                  'select the appropriate option button
                  m_percAggress = sArr(i)
                  Select Case m_percAggress
                     Case Is = 10
                        optAgressiveness(0).Value = True
                     Case Is = 15
                        optAgressiveness(1).Value = True
                     Case Is = 20
                        optAgressiveness(2).Value = True
                  End Select
                  
               Case Is = 1 'checkbox
                  ckOptions(0).Value = sArr(i)
                  
               Case Is = 2 'if the checkbox is checked then
                          'select the approp. corresp. opt btn
                 If ckOptions(0).Value = vbChecked Then
                     m_PurgeTime = sArr(i)
                     Select Case m_PurgeTime
                        Case Is = 300000
                           cbo1.ListIndex = 0
                        Case Is = 900000
                           cbo1.ListIndex = 1
                        Case Is = 1800000
                           cbo1.ListIndex = 2
                     End Select
                  End If
                  
               Case Is = 3 'checkbox
                  ckOptions(1).Value = sArr(i)
                  
               Case Is = 4 'if the checkbox is checked then
                          'select the approp. corresp. opt btn
                  If ckOptions(1).Value = vbChecked Then
                      m_DropPerc = sArr(i)
                      Select Case m_DropPerc
                        Case Is = 0.1
                           optPurgePercent(0).Value = True
                        Case Is = 0.2
                           optPurgePercent(1).Value = True
                        Case Is = 0.3
                           optPurgePercent(2).Value = True
                      End Select
                  End If
            End Select
        Next i
'END CODE:
'exit sub
ERR:
  Debug.Print ERR.Description

End Sub
Private Sub Form_Resize()
        '
        With Shape1
          .Left = 150: .Top = 150: .Width = (Width - 300): .Height = (Height - 300)
        End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '
        Erase m_buff
        'save prog data
        Call ModAPI.sub_FileSaveData( _
               m_FilePath, _
               m_percAggress, ckOptions(0).Value, m_PurgeTime, _
               ckOptions(0).Value, m_DropPerc _
               )
End Sub




'          .    .     .    .  ''  _     .    .          .
' _ __ __    ___ _   __ ___   ' / _| ___  _ _  _ __ __
'| '_ ` _ \ / _ \ \ / // _ \  '| |_ / _ \| '_\| '_ ` _ \
'| | | | | | (_) \ V /   __/  '|  _| (_) | |  | | | | | |
'|_| |_| |_|\___/ \_/  \___|  '|_|  \___/|_|  |_| |_| |_|


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        '
        If Button = 1 Then Call mod_Move(hWnd)
End Sub
Private Sub fmeAgressiveness_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        '
        If Button = 1 Then Call mod_Move(hWnd)
End Sub

Private Sub fmeOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        '
        If Button = 1 Then Call mod_Move(hWnd)
End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        '
         If Button = 1 Then Call mod_Move(hWnd)
End Sub

 
 

'    .     .          .  ' _       .'  _     .
' _ _  __ _  _ __ __     '(_) _ __   / _| ___
'| '_\/ _` || '_ ` _ \   '| || '_ ` | |_ / _ \
'| |   (_| || | | | | |  '| || | | ||  _| (_)
'|_|  \__,_||_| |_| |_|  '|_||_| |_||_|  \___/
 

Private Sub tmrRAM_Timer()
'On Error GoTo ERR:
'---------------------------------
'obtain ram info
'if form is visible then the interval
' is 1.5 seconds, if its trayed
'its once every 6 seconds
'if were purging its every second
'---------------------------------
'VARIABLES:
  Dim arr() As Long, Elapsed&
  Dim strInfo(1) As String, totStr$
  Dim i%
'CODE:
  'were returning 2 part array from
  'the function in modAPI
  arr = func_GetMemoryStats
  m_TotRAM = arr(0)
  m_FreeRAM = arr(1)
  'first parm is total sys ram
  strInfo(0) = CStr(arr(0))
  'second part is avail ram
  strInfo(1) = CStr(arr(1))
  '
  m_UsedRAM = (arr(0) - arr(1))
  '
  With picInfo
     .Cls
     .CurrentX = 200
     .CurrentY = 50
  End With
  '
  totStr = "total RAM=" & strInfo(0) & " mb" & _
          "    used RAM=" & CStr(m_UsedRAM) & " mb" & _
          "    free RAM=" & strInfo(1) & " mb"
          '
  'if form visible, print info in picbox
  If b_InTray = False Then
      picInfo.Print totStr
  'otherwise, show info in tooltip in tray
  Else
      Tray1.TrayToolTip = totStr
      'this updates the tooltip in the tray
      Tray1.Show True
  End If
  
  
  ' if the free ram drops below specified percent level
  ' (determined by which optPurgePercent index selected)
  ' then purge ram
  If ckOptions(1).Value = vbChecked Then
      If (m_FreeRAM / m_TotRAM) < m_DropPerc Then
         Call PurgeRam(DeterminPurgeAmount(m_UsedRAM, m_FreeRAM, m_percAggress))
      End If
  End If
  
  
  'this means the user has checked to
  'purge ram at timed intervals
  If m_PurgeTime > 0 Then
      Elapsed = (GetTickCount - m_TimerPasses)
      'only do the calc if form is not in tray and can be seen
      If b_InTray = False Then
          'when time remaining is less than a minute
          'instead of show 0 as time remaining, which
          'might be confusing to user, specifiy..less than 1 minute
          If ((m_PurgeTime - Elapsed) \ 60000) <= 0 Then
                lblNextPurge = "Next RAM purge: " & vbCrLf & _
                         "less than 1 minute"
          Else
             'format remaining time til next purge in label
              lblNextPurge = "Next RAM purge: " & vbCrLf & _
                         ((m_PurgeTime - Elapsed) \ 60000) & _
                         " minutes"
          End If
      End If
      'timer select by combo box has expired
      If Elapsed >= m_PurgeTime Then
           m_TimerPasses = GetTickCount
           Call PurgeRam(DeterminPurgeAmount(m_UsedRAM, m_FreeRAM, m_percAggress))
      End If
  End If
'END CODE:
'exit sub
ERR:
  Debug.Print ERR.Description
End Sub

 
