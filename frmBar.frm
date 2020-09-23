VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VB_Codeboard_Buddy"
   ClientHeight    =   2910
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   11475
   Icon            =   "frmBar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerMailFlash 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5175
      Top             =   765
   End
   Begin MSWinsockLib.Winsock sockPOP 
      Left            =   6750
      Top             =   1305
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TimerRam 
      Left            =   7425
      Top             =   450
   End
   Begin VB.Timer TimerPaste 
      Left            =   9000
      Top             =   810
   End
   Begin VB.Timer TimerAnimate 
      Enabled         =   0   'False
      Left            =   8325
      Top             =   765
   End
   Begin VB.Timer timerMailCheck 
      Left            =   7695
      Top             =   855
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   6030
      Top             =   855
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   12075
      Begin VB.CommandButton btnLoadRamPurge 
         BackColor       =   &H00EFEFEF&
         Height          =   225
         Left            =   1215
         TabIndex        =   19
         ToolTipText     =   "load RAM purge"
         Top             =   270
         Width           =   225
      End
      Begin VB.ListBox lstSort 
         Height          =   1035
         Left            =   3690
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   1440
         Width           =   4245
      End
      Begin VB.Frame fmeEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Email check"
         ForeColor       =   &H00FF0000&
         Height          =   1320
         Left            =   45
         TabIndex        =   7
         ToolTipText     =   "Right click to set/alter accounts"
         Top             =   945
         Width           =   1320
         Begin VB.CheckBox ckKillMailFlash 
            BackColor       =   &H00FFFFFF&
            Caption         =   "kill flashing"
            ForeColor       =   &H00808080&
            Height          =   150
            Left            =   90
            TabIndex        =   18
            Top             =   1125
            Width           =   1095
         End
         Begin VB.Label lblPopName 
            BackStyle       =   0  'Transparent
            Caption         =   "4th account"
            ForeColor       =   &H00FF8080&
            Height          =   210
            Index           =   3
            Left            =   315
            TabIndex        =   15
            ToolTipText     =   "right click to edit/alter mail info"
            Top             =   855
            Width           =   885
         End
         Begin VB.Label lblPopName 
            BackStyle       =   0  'Transparent
            Caption         =   "3rd account"
            ForeColor       =   &H00FF8080&
            Height          =   210
            Index           =   2
            Left            =   315
            TabIndex        =   14
            ToolTipText     =   "right click to edit/alter mail info"
            Top             =   630
            Width           =   885
         End
         Begin VB.Label lblPopName 
            BackStyle       =   0  'Transparent
            Caption         =   "2nd account"
            ForeColor       =   &H00FF8080&
            Height          =   225
            Index           =   1
            Left            =   315
            TabIndex        =   13
            ToolTipText     =   "right click to edit/alter mail info"
            Top             =   405
            Width           =   930
         End
         Begin VB.Label lblPopName 
            BackStyle       =   0  'Transparent
            Caption         =   "1st account"
            ForeColor       =   &H00FF8080&
            Height          =   210
            Index           =   0
            Left            =   315
            TabIndex        =   12
            ToolTipText     =   "right click to edit/alter mail info"
            Top             =   180
            Width           =   885
         End
         Begin VB.Label lblPOP 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   11
            ToolTipText     =   "right click to check now"
            Top             =   630
            Width           =   180
         End
         Begin VB.Label lblPOP 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   2
            Left            =   90
            TabIndex        =   10
            ToolTipText     =   "right click to check now"
            Top             =   855
            Width           =   180
         End
         Begin VB.Label lblPOP 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   1
            Left            =   90
            TabIndex        =   9
            ToolTipText     =   "right click to check now"
            Top             =   405
            Width           =   180
         End
         Begin VB.Label lblPOP 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   8
            ToolTipText     =   "right click to check now"
            Top             =   180
            Width           =   180
         End
      End
      Begin VB.PictureBox picCPUload 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   90
         ScaleHeight     =   120
         ScaleWidth      =   750
         TabIndex        =   5
         Top             =   720
         Width           =   780
      End
      Begin VB.PictureBox picRAM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   90
         ScaleHeight     =   120
         ScaleWidth      =   750
         TabIndex        =   1
         Top             =   315
         Width           =   780
      End
      Begin VB.Label lblNextMailCheck 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   45
         TabIndex        =   16
         Top             =   2250
         Width           =   1350
      End
      Begin VB.Label lblCpuLoad 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   150
         Index           =   1
         Left            =   855
         TabIndex        =   6
         Top             =   720
         Width           =   390
      End
      Begin VB.Label lblCpuLoad 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CPU load"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   525
         Width           =   885
      End
      Begin VB.Label lblTotalRam 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   150
         Left            =   855
         TabIndex        =   3
         Top             =   315
         Width           =   300
      End
      Begin VB.Label lblUsedRam 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   -45
         TabIndex        =   2
         Top             =   120
         Width           =   1305
      End
   End
   Begin VB.Menu mnuVBfunctions 
      Caption         =   "vb functions"
      Begin VB.Menu mnuDoLoop 
         Caption         =   "Do...Loop"
      End
      Begin VB.Menu mnuIfEndIf 
         Caption         =   "If...End If"
      End
      Begin VB.Menu mnuIfElseifEndif 
         Caption         =   "If...Elseif...End If"
      End
      Begin VB.Menu mnuSelectCase 
         Caption         =   "Select Case..."
      End
      Begin VB.Menu mnuForI 
         Caption         =   "For I..."
      End
      Begin VB.Menu mnuWhileWend 
         Caption         =   "While...Wend"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuErrResume 
         Caption         =   "On Error Resume Next"
      End
      Begin VB.Menu mnuFullErrHandler 
         Caption         =   "Full Err handler"
         Begin VB.Menu mnuFullErrHandlerViewEdit 
            Caption         =   "View/Edit"
         End
         Begin VB.Menu mnuFullErrHandlerInsert 
            Caption         =   "Insert"
         End
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMsgboxShellvbYesNo 
         Caption         =   "MsgBox Shell (vbYesNo)"
      End
      Begin VB.Menu mnuInpBoxShell 
         Caption         =   "InputBox Shell"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProcHeadComment 
         Caption         =   "Procedure Header Comment"
         Begin VB.Menu mnuPrcHeadCommentViewEdit 
            Caption         =   "View/Edit"
         End
         Begin VB.Menu mnuPrcHeadCommentInsert 
            Caption         =   "Insert"
         End
      End
      Begin VB.Menu mnuSubShell 
         Caption         =   "Sub Routine Shell"
      End
      Begin VB.Menu mnuFunction 
         Caption         =   "Function Shell"
      End
      Begin VB.Menu mnuPropLetGetShell 
         Caption         =   "Property Get/Let Shell"
      End
   End
   Begin VB.Menu mnuOtherFunctions 
      Caption         =   "api functions"
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "window: is form loaded"
         Index           =   25
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "window: move without titlebar"
         Index           =   50
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "window: z-order"
         Index           =   100
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "menu: make columns"
         Index           =   200
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "general: basic subclass code"
         Index           =   300
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "general: twips to pixel"
         Index           =   310
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "general: textbox numbers && backspace only"
         Index           =   320
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "graphics: color to html"
         Index           =   400
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "graphics: color Long to RGB"
         Index           =   401
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "graphics: text to screen (full class)"
         Index           =   402
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "file: ini manipulation class"
         Index           =   500
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "browser: get .hwnd and .hdc"
         Index           =   600
      End
      Begin VB.Menu mnuSub_apiFunctions 
         Caption         =   "array: is it initialized"
         Index           =   700
      End
   End
   Begin VB.Menu mnuSubsCode 
      Caption         =   "code"
      Begin VB.Menu mnuLoadCodeFile 
         Caption         =   "Load Code File"
      End
      Begin VB.Menu mnuSaveCodeBlocksToFile 
         Caption         =   "Save these code blocks to file"
      End
      Begin VB.Menu sep40 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddCodeBlock 
         Caption         =   "Add Sub/Function/Code block"
      End
      Begin VB.Menu mnuRemoveCodeBlock 
         Caption         =   "Remove Sub/Function/Code block"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuClearCodeBlocksFromMenu 
         Caption         =   "Clear these code blocks from menu"
      End
      Begin VB.Menu sep93 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrCode 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuApiWindow 
      Caption         =   "window"
      Begin VB.Menu mnuWindowShell 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuMouse 
      Caption         =   "mouse"
      Begin VB.Menu mnuMouseShell 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuAPIGraphx 
      Caption         =   "graphx"
      Begin VB.Menu mnuGraphxShell 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuAPIRect 
      Caption         =   "rect/rgn"
      Begin VB.Menu mnuRectShell 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuAPImenu 
      Caption         =   "menu"
      Begin VB.Menu mnuMenuShell 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuMsgMenu 
      Caption         =   "message"
      Begin VB.Menu mnuMsgShell 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu sep25 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMsgTypesConstants 
         Caption         =   "BM messages (button)"
         Index           =   5
      End
      Begin VB.Menu mnuMsgTypesConstants 
         Caption         =   "CB messages (combo box)"
         Index           =   10
      End
      Begin VB.Menu mnuMsgTypesConstants 
         Caption         =   "EM messages (textbox)"
         Index           =   15
      End
      Begin VB.Menu mnuMsgTypesConstants 
         Caption         =   "LB messages (listbox)"
         Index           =   20
      End
      Begin VB.Menu mnuMsgTypesConstants 
         Caption         =   "WM messages (general windows)"
         Index           =   25
      End
      Begin VB.Menu mnuMsgTypesConstants 
         Caption         =   "-"
         Index           =   30
      End
      Begin VB.Menu mnuMsgTypesConstants 
         Caption         =   "KeyCode constants"
         Index           =   35
      End
   End
   Begin VB.Menu mnuMiscMenu 
      Caption         =   "misc."
      Begin VB.Menu mnuMiscShell 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuClipboard 
      Caption         =   "clipboard"
      Begin VB.Menu mnuClipboardLoad 
         Caption         =   "&Load clipboard file"
      End
      Begin VB.Menu mnuClipboardSave 
         Caption         =   "&Save this clipboard to file"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClipboardAdd 
         Caption         =   "&Add item..."
      End
      Begin VB.Menu mnuClipboardRemove 
         Caption         =   "&Remove item..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuClipboardClear 
         Caption         =   "&Clear clipboard contents"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrClipboard 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "options"
      Begin VB.Menu mnuOptionsShowCompStats 
         Caption         =   "Show computer statistics"
      End
      Begin VB.Menu mnuOptionsHideCompStats 
         Caption         =   "Hide computer statistics"
      End
      Begin VB.Menu sep99 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoShowContTypesForm 
         Caption         =   "AutoShow const/types form(s) on API paste"
      End
      Begin VB.Menu mnuDontShowContTypesForm 
         Caption         =   "Dont show const/types form(s) on API paste"
      End
      Begin VB.Menu sep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnableScreenHelp 
         Caption         =   "Enable Screen text tips"
      End
      Begin VB.Menu mnuDisableScreenHelp 
         Caption         =   "Disable Screen Text Tips"
      End
      Begin VB.Menu sep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsEnableEmailChecking 
         Caption         =   "Enable Email checking"
      End
      Begin VB.Menu mnuOptionsDisableEmailChecking 
         Caption         =   "Disable Email checking"
      End
      Begin VB.Menu sep50 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsEmail 
         Caption         =   "Email options"
      End
      Begin VB.Menu mnuShowDefMailClient 
         Caption         =   "Show default mail client"
      End
      Begin VB.Menu sep38 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadRamPurge 
         Caption         =   "Load RAM purge"
      End
   End
   Begin VB.Menu mnuApiResources 
      Caption         =   "API resources"
      Begin VB.Menu mnuSubAdditApiRes 
         Caption         =   "Launch APT-Guide"
         Index           =   10
      End
      Begin VB.Menu mnuSubAdditApiRes 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuSubAdditApiRes 
         Caption         =   "Adrea's API call list"
         Index           =   20
      End
      Begin VB.Menu mnuSubAdditApiRes 
         Caption         =   "&Extensive api listings"
         Index           =   30
      End
      Begin VB.Menu mnuSubAdditApiRes 
         Caption         =   "&MSDN api listing "
         Index           =   40
      End
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&X "
   End
End
Attribute VB_Name = "frmBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Enum APIT
         api_window
         api_Mouse
         api_RectRgn
         api_Graphx
         api_Menu
         api_Msg
         api_Misc
         api_TypesConst
End Enum
 
 

 'boolean
Public bRemoveMenuMode As Boolean, m_bAddPrivate As Boolean
Private b_Loading  As Boolean, b_SysInfoVisible   As Boolean

 'string
Private strClipBoardData$, strAPInameLastSent$
Public strHeaderComment$, strErrHandler$

'integer
Private bMailFlash(3) As Boolean

 'long
Private LastMailCheckTickCount&

'class variables
Private WithEvents cPOP               As clsPOP
Attribute cPOP.VB_VarHelpID = -1
Private WithEvents cSysInfo           As clsSysMonitor
Attribute cSysInfo.VB_VarHelpID = -1
Private cScreenMsg                    As clsScreenText
Public cClipboard                     As clsClipboard




'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Private Sub LoadApiMenus(APItype As APIT)
'VARIABLES:
 Dim iUpper%, iLower%, i%
 Dim str$, str2$, str3$
 Dim mnuName As Variant
 On Error Resume Next
'CODE:
        'iLower is lower number of API function
        'string in res file
        'iUpper is the upper number
        If APItype = api_Graphx Then
           iUpper% = 370
           iLower% = 301
           Set mnuName = mnuGraphxShell
           
        ElseIf APItype = api_Misc Then
           iUpper% = 813
           iLower% = 801
           Set mnuName = mnuMiscShell
           
        ElseIf APItype = api_Menu Then
           iUpper% = 520
           iLower% = 501
           Set mnuName = mnuMenuShell
           
        ElseIf APItype = api_RectRgn Then
           iUpper% = 427
           iLower% = 401
           Set mnuName = mnuRectShell
           
        ElseIf APItype = api_Mouse Then
           iUpper% = 210
           iLower% = 201
           Set mnuName = mnuMouseShell
           
        ElseIf APItype = api_Msg Then
           iUpper% = 612
           iLower% = 601
           Set mnuName = mnuMsgShell
           
        ElseIf APItype = api_window Then
           iUpper% = 141
           iLower = 101
           Set mnuName = mnuWindowShell
        
        End If
 
        'clear listbox
        lstSort.Clear
        
        
        'place strings from res file into listbox
        'sort prop = true so the api's will become
        'alphabetically sorted
        For i = iLower To iUpper
               str = LoadResString(i)
               
               'if this is the types/const menu
               'were adding directly to menu from
               'res file without parsing at all
               If APItype = api_TypesConst Then
                     lstSort.AddItem str
               Else
                     lstSort.AddItem str & (i)
               End If
        Next i
       
        
        'place the items in the menu
        For i = 1 To lstSort.ListCount
            Load mnuName(i)
            
             If APItype = api_TypesConst Then
                  mnuName(i).Caption = lstSort.List(i - 1)
             Else
                  mnuName(i).Caption = func_ParseAPIname(lstSort.List(i - 1))
             End If
             'show the menu
             mnuName(i).Visible = True
             mnuName(i).Tag = Right(lstSort.List(i - 1), 3)
              
         Next i
        '
       ' make index 0 invisible
        mnuName(0).Visible = False
        
        'free mem
        Set mnuName = Nothing
'CODE END:

End Sub

 
 


Private Function funcSplitMenu(mnuCaption$) As String()
 Dim i%
 Dim sArr(1) As String, sParts() As String
 
       'seperates the api string from the res number
       'i.e.. "DrawText     334" the following code
       'will split the string by the " " giving us
       'part 1 which is the api name and part2 the res #
       sParts = Split(mnuCaption, Chr(32))
       
       sArr(0) = sParts(0)
       
       'extract the api(str) and res number
       For i = 1 To UBound(sParts)
           If Trim(sParts(i)) <> "" Then
               'the res number
                sArr(1) = Trim(sParts(i))
                Exit For
           End If
       Next i
       
       funcSplitMenu = sArr
End Function
 

 

 

 








 

 '         . _ __   '    .    .     .    .    .  '          .    .      .     .
'_____  ___ | '_ \  ' _    ___ _   __ ___  _     ' _ __ __    ___  _ __  _ _
'_   _|/ _ \| |_)   '| |  / _ \ \ / // _ \| |    '| '_ ` _ \ / _ \| '_ `  | | |
' | |   (_) | .__/  '| |_   __/\ V /   __/| |_   '| | | | | |  __/| | | | |_| |
' |_|  \___/|_|     '|___|\___| \_/  \___||___|  '|_| |_| |_|\___||_| |_|\__,_|
'    .    . _     . _   .    .
' ___  _   (_) ___ | | __ _ _'
'/ __|| |  | |/ __|| |/ // __|
' (__ | |_ | | (__ |   < \__ \
'\___||___||_|\___||_|\_\|___/

 
 

 

Private Sub ckKillMailFlash_Click()
        '-----------------------------
        'turn off timer that causes flashing
        'of labels to notify user of mail waiting
        '----------------------------
        ckKillMailFlash.Value = vbUnchecked
        '
        If TimerMailFlash.Enabled = True Then
            Call TimerFlashOff
        End If
End Sub

Private Sub mnuAPIGraphx_Click()
        'screen message instructing user how to do this
         If bEnableScreenTips = True Then
            Set cScreenMsg = New clsScreenText
            cScreenMsg.ScreenMsg _
                    "Holding  SHIFT  key while clicking menu item will " & vbCrLf & _
                    "show form containing constants/types for that API " & vbCrLf & _
                    "otherwise, the API will be copied to clipboard.   " & vbCrLf & _
                    "Holding down  CTRL key while clicking " & vbCrLf & _
                    "menu item adds " & Chr(34) & "Private" & Chr(34) & _
                    "to the API", 12, vbBlue, _
                    True, 2500, 5000
            Set cScreenMsg = Nothing
         End If
End Sub

Private Sub mnuAPImenu_Click()
        'screen message instructing user how to do this
         If bEnableScreenTips = True Then
            Set cScreenMsg = New clsScreenText
            cScreenMsg.ScreenMsg _
                    "Holding  SHIFT  key while clicking menu item will " & vbCrLf & _
                    "show form containing constants/types for that API " & vbCrLf & _
                    "otherwise, the API will be copied to clipboard.   " & vbCrLf & _
                    "Holding down  CTRL key while clicking " & vbCrLf & _
                    "menu item adds " & Chr(34) & "Private" & Chr(34) & _
                    "to the API", 12, vbBlue, _
                    True, 2500, 5000
            Set cScreenMsg = Nothing
         End If
End Sub

Private Sub mnuAPIRect_Click()
        'screen message instructing user how to do this
         If bEnableScreenTips = True Then
            Set cScreenMsg = New clsScreenText
            cScreenMsg.ScreenMsg _
                    "Holding  SHIFT  key while clicking menu item will " & vbCrLf & _
                    "show form containing constants/types for that API " & vbCrLf & _
                    "otherwise, the API will be copied to clipboard.   " & vbCrLf & _
                    "Holding down  CTRL key while clicking " & vbCrLf & _
                    "menu item adds " & Chr(34) & "Private" & Chr(34) & _
                    "to the API", 12, vbBlue, _
                    True, 2500, 5000
            Set cScreenMsg = Nothing
         End If
End Sub

Private Sub mnuApiWindow_Click()
        'screen message instructing user how to do this
         If bEnableScreenTips = True Then
            Set cScreenMsg = New clsScreenText
            cScreenMsg.ScreenMsg _
                    "Holding  SHIFT  key while clicking menu item will " & vbCrLf & _
                    "show form containing constants/types for that API " & vbCrLf & _
                    "otherwise, the API will be copied to clipboard.   " & vbCrLf & _
                    "Holding down  CTRL key while clicking " & vbCrLf & _
                    "menu item adds " & Chr(34) & "Private" & Chr(34) & _
                    "to the API", 12, vbBlue, _
                    True, 2500, 5000
            Set cScreenMsg = Nothing
         End If
End Sub



Private Sub mnuMouse_Click()
        'screen message instructing user how to do this
         If bEnableScreenTips = True Then
            Set cScreenMsg = New clsScreenText
            cScreenMsg.ScreenMsg _
                    "Holding  SHIFT  key while clicking menu item will " & vbCrLf & _
                    "show form containing constants/types for that API " & vbCrLf & _
                    "otherwise, the API will be copied to clipboard.   " & vbCrLf & _
                    "Holding down  CTRL key while clicking " & vbCrLf & _
                    "menu item adds " & Chr(34) & "Private" & Chr(34) & _
                    "to the API", 12, vbBlue, _
                    True, 2500, 5000
            Set cScreenMsg = Nothing
         End If
End Sub

Private Sub mnuMsgMenu_Click()
        'screen message instructing user how to do this
         If bEnableScreenTips = True Then
            Set cScreenMsg = New clsScreenText
            cScreenMsg.ScreenMsg _
                    "Holding  SHIFT  key while clicking menu item will " & vbCrLf & _
                    "show form containing constants/types for that API " & vbCrLf & _
                    "otherwise, the API will be copied to clipboard.   " & vbCrLf & _
                    "Holding down  CTRL key while clicking " & vbCrLf & _
                    "menu item adds " & Chr(34) & "Private" & Chr(34) & _
                    "to the API", 12, vbBlue, _
                    True, 2500, 5000
            Set cScreenMsg = Nothing
         End If
End Sub

Private Sub mnuMiscMenu_Click()
        'screen message instructing user how to do this
         If bEnableScreenTips = True Then
            Set cScreenMsg = New clsScreenText
            cScreenMsg.ScreenMsg _
                    "Holding  SHIFT  key while clicking menu item will " & vbCrLf & _
                    "show form containing constants/types for that API " & vbCrLf & _
                    "otherwise, the API will be copied to clipboard.   " & vbCrLf & _
                    "Holding down  CTRL key while clicking " & vbCrLf & _
                    "menu item adds " & Chr(34) & "Private" & Chr(34) & _
                    "to the API", 12, vbBlue, _
                    True, 2500, 5000
            Set cScreenMsg = Nothing
         End If
End Sub








 
 

 

 

 

 

'          .    .      .     .  '    .    . _     . _   .    .
' _ __ __    ___  _ __  _   _   ' ___  _   (_) ___ | | __ _ _'
'| '_ ` _ \ / _ \| '_ `  | | |  '/ __|| |  | |/ __|| |/ // __|
'| | | | | |  __/| | | | |_| |  ' (__ | |_ | | (__ |   < \__ \
'|_| |_| |_|\___||_| |_|\__,_|  '\___||___||_|\___||_|\_\|___/

Private Sub mnuShowDefMailClient_Click()
On Error GoTo ERR:
'---------------------------------
'launch mail client
'---------------------------------
'VARIABLES:

'CODE:
   'check to see if mailclient path has been established
   If Trim(modPub.sMailClientPath) = "" Then
       'if not..notify and allow user to select a mail client
       MsgBox "No mail client has been specified.", vbExclamation
       With cmDlg
           .Filter = "exe's|*.exe"
           .CancelError = True
           .ShowOpen
            'if he chose a valid file, set its path to the variable
           If .FileName <> "" Then
               modPub.sMailClientPath = .FileName
               'launch
               ShellExecute hwnd, "open", modPub.sMailClientPath, vbNull, "c:\", 1
           End If
       End With
    Else
       'valid mail client path specified so launch
       ShellExecute hwnd, "open", modPub.sMailClientPath, vbNull, "c:\", 1
    End If
'END CODE:
Exit Sub
ERR:
  Debug.Print ERR.Description
End Sub
'
'various popular, frequently used functions library
Private Sub mnuSub_apiFunctions_Click(Index As Integer)
  
  Clipboard.Clear
  
  With frmCodeHolder
      Select Case Index
      
          Case Is = 25 'is form loaded
              Call SendData(.windowFormIsLoaded.Text)
          Case Is = 50 'move control or window without titlebar
              Call SendData(.windowMoveWithoutTitltebar.Text)
          Case Is = 100 'window z order
              Call SendData(.windowZorder.Text)
          Case Is = 200 'menu into columns
               Call SendData(.menuMakeColumns.Text)
          Case Is = 300 'general subclass
              Call SendData(.subclassBasic.Text)
          Case Is = 310 'twips to pixels
              Call SendData(.generalTwipToPix.Text)
          Case Is = 320 'allow only numbers and BACK in textbox
              Call SendData(.generalTextboxNumbersOnly.Text)
          Case Is = 400 'color to html compatible
              Call SendData(.graphicsColorToHtml.Text)
          Case Is = 401
              Call SendData(.graphicsColorLongToRgb.Text)
          Case Is = 402
              Call SendData(.graphicsTextToScreen.Text)
          Case Is = 500 'get browsers hwnd and hdc
              Call SendData(.generalIniFileManipulationClass)
          Case Is = 600 'get browser hwnd and hdc
              Call SendData(.browserGetHwnd.Text)
          Case Is = 700 'is array initialized
              Call SendData(.arrayIsInitialized.Text)
              
      End Select
  End With
  
End Sub

Private Sub mnuSubAdditApiRes_Click(Index As Integer)
 
 Dim strsite    As String
 
 Select Case Index
    Case Is = 10 'launch api guide
       Dim path_buff As String, str_root As String
       Dim path_to_apiguide As String
       
       path_buff = String(256, " ")
       strsite = ""
       '  check for the existence of api guide on the
       '  users computer. If its not found then go to
       '  download page for it, if user wants
         GetWindowsDirectory path_buff, 256
         str_root = Split(path_buff, "\")(0)
         '  if api-guide exists it will prolly be on this path
         path_to_apiguide = str_root & "\Program Files\API-Guide\API-Guide.exe"
         
         'if its found launch it
         If Len(Trim$(Dir(path_to_apiguide))) > 0 Then
            Shell path_to_apiguide
         Else
            ' if not found go to download site
            strsite = "http://www.mentalis.org/agnet/apiguide.shtml"
         End If
         
    Case Is = 20 'go to andreas api list website
        strsite = "http://www.andreavb.com/API_List.html"
    
    Case Is = 30
        strsite = "http://custom.programming-in.net/articles/art9-2.asp?lib=kernel32.dll"
       
    Case Is = 40 'msdn api reference
        strsite = "http://msdn.microsoft.com/library/en-us/winprog/winprog/functions_by_category.asp?"
 End Select
 
 ' if there is a site specified in on of the select cases ..go there
 If Len(Trim$(strsite)) > 0 Then
    ShellExecute hwnd, "open", strsite, vbNullString, _
    vbNullString, 1
 End If
 
End Sub

Private Sub mnuWindowShell_Click(Index As Integer)
      
  'if shift key is being pressed while API
  'menu item is being clicked on then dont
  'send API..just dispay form showing any
  'Constants or types for the associated API
  If funcCheckForShift = True Then
      'extract the api name
      strAPInameLastSent = Trim(LCase(mnuWindowShell(Index).Caption))
      Call ShowApiConstForm
  Else
      'user wants to add "Private" to the api call
      If funcCheckForControl = True Then
          m_bAddPrivate = True
      End If
      '
      Call MnuBeingClicked(api_window, Index, mnuWindowShell(Index))
  End If
  
End Sub
 
Private Sub mnuMouseShell_Click(Index As Integer)
   '
   'if shift key is being pressed while API
   'menu item is being clicked on then dont
   'send API..just dispay form showing any
   'Constants or types for the associated API
   If funcCheckForShift = True Then
       'extract the api name
       strAPInameLastSent = Trim(LCase(mnuMouseShell(Index).Caption))
       Call ShowApiConstForm
   Else
       'user wants to add "Private" to the api call
       If funcCheckForControl = True Then
           m_bAddPrivate = True
       End If
       '
       Call MnuBeingClicked(api_Mouse, Index, mnuMouseShell(Index))
   End If
   
End Sub
  
Private Sub mnuMenuShell_Click(Index As Integer)
   '
   'if shift key is being pressed while API
   'menu item is being clicked on then dont
   'send API..just dispay form showing any
   'Constants or types for the associated API
   If funcCheckForShift = True Then
       'extract the api name
       strAPInameLastSent = Trim(LCase(mnuMenuShell(Index).Caption))
       Call ShowApiConstForm
   Else
       'user wants to add "Private" to the api call
       If funcCheckForControl = True Then
           m_bAddPrivate = True
       End If
       '
       Call MnuBeingClicked(api_Menu, Index, mnuMenuShell(Index))
   End If
   
End Sub
Private Sub mnuMsgShell_Click(Index As Integer)
   '
   'if shift key is being pressed while API
   'menu item is being clicked on then dont
   'send API..just dispay form showing any
   'Constants or types for the associated API
   If funcCheckForShift = True Then
       'extract the api name
       strAPInameLastSent = Trim(LCase(mnuMsgShell(Index).Caption))
       Call ShowApiConstForm
   Else
       'user wants to add "Private" to the api call
       If funcCheckForControl = True Then
           m_bAddPrivate = True
       End If
       '
       Call MnuBeingClicked(api_Msg, Index, mnuMsgShell(Index))
   End If
   
End Sub
Private Sub mnuGraphxshell_Click(Index As Integer)
   '
   'if shift key is being pressed while API
   'menu item is being clicked on then dont
   'send API..just dispay form showing any
   'Constants or types for the associated API
   If funcCheckForShift = True Then
       'extract the api name
       strAPInameLastSent = Trim(LCase(mnuGraphxShell(Index).Caption))
       Call ShowApiConstForm
   Else
       'user wants to add "Private" to the api call
       If funcCheckForControl = True Then
           m_bAddPrivate = True
       End If
       '
       Call MnuBeingClicked(api_Graphx, Index, mnuGraphxShell(Index))
   End If
   
End Sub
 
Private Sub mnuRectShell_Click(Index As Integer)
 '
 'if shift key is being pressed while API
 'menu item is being clicked on then dont
 'send API..just dispay form showing any
 'Constants or types for the associated API
 If funcCheckForShift = True Then
     'extract the api name
     strAPInameLastSent = Trim(LCase(mnuRectShell(Index).Caption))
     Call ShowApiConstForm
 Else
     'user wants to add "Private" to the api call
     If funcCheckForControl = True Then
         m_bAddPrivate = True
     End If
     '
     Call MnuBeingClicked(api_RectRgn, Index, mnuRectShell(Index))
 End If
 
End Sub
Private Sub mnuMiscShell_Click(Index As Integer)
  '
  'if shift key is being pressed while API
  'menu item is being clicked on then dont
  'send API..just dispay form showing any
  'Constants or types for the associated API
  If funcCheckForShift = True Then
      'extract the api name
      strAPInameLastSent = Trim(LCase(mnuMiscShell(Index).Caption))
      Call ShowApiConstForm
  Else
      'user wants to add "Private" to the api call
      If funcCheckForControl = True Then
          m_bAddPrivate = True
      End If
      '
      Call MnuBeingClicked(api_Misc, Index, mnuMiscShell(Index))
  End If
  
End Sub

 


'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |the index of the res is loaded into str
'            str is passed to the clipboard and sent
'            on mouse down (see senddata)
'----------------------------------------------------------------------
Private Sub MnuBeingClicked( _
                         APItype As APIT, _
                         Index As Integer, _
                         menu As menu)
'VARIABLES
  Dim str$, sParts() As String
  Dim i%, resNum%
  
'CODE:

       'the api name
       '(ShowApiConstForm)
        strAPInameLastSent$ = Trim(LCase(menu.Caption))
       
       'load the res string into var str
        str = LoadResString(CLng(menu.Tag))
        
        'this is toggled to true if the user
        'is holding down the letter p while
        'clicking the menu item
        If m_bAddPrivate = True Then
            str = ("Private " & str)
            m_bAddPrivate = False
        End If
        
        modPub.bSendingApiNotConst = True
       
       'paste
        Call SendData(str, _
             "whereever you click next, the API will be sent...or...press escape, " & _
                        "then   CTL + V when you wish to paste")
'END CODE:

End Sub

Private Sub mnuMsgTypesConstants_Click(Index As Integer)
'---------------------------------
'show constants related to SendMessage API
'---------------------------------
        Select Case Index
            Case Is = 5 'BTN
                Call modConstants.BMmessage(frmConstants.lstConstants)
            Case Is = 10 'CB
                Call modConstants.CBmessage(frmConstants.lstConstants)
            Case Is = 15 'EM
                Call modConstants.EMmessage(frmConstants.lstConstants)
            Case Is = 20 'LB
                Call modConstants.LBmessage(frmConstants.lstConstants)
            Case Is = 25 'WM
                Call modConstants.WMmessage(frmConstants.lstConstants)
            Case Is = 35 'KeyCode constants
                Call modConstants.KeyCodeConst(frmConstants.lstConstants)
        End Select
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | 'here we examine the name of the api
               'and show the form(s) the present
               'required const/types
'----------------------------------------------------------------------
Private Sub ShowApiConstForm()
       
'CODE:
       'var in declarations
       Select Case LCase(Trim(strAPInameLastSent))
       
'API(Misc)========
                Case Is = "systemparametersinfo"
                     Call modConstants.SystemParametersInfo(frmConstants.lstConstants)
                     
                Case Is = "drawtext"
                     Call modTypes.TypeRect(frmType.lstType)
                     
                Case Is = "messagebeep"
                     Call modConstants.MessageBeepConst(frmConstants.lstConstants)
                     
                Case Is = "playsound"
                     Call modConstants.SoundConst(frmConstants.lstConstants)
                     
                Case Is = "registerhotkey"
                     Call modConstants.RegisterHotKeyConst(frmConstants.lstConstants)
                
                Case Is = "shellexecute"
                      Call modConstants.ShowWindowConst(frmConstants.lstConstants)
                
                Case Is = "createfile"
                      Call modConstants.CreateFileConst(frmConstants.lstConstants)
'API(Message)========
        
                Case Is = "broadcastsystemmessage"
                      Call modConstants.BroadcastSysMsgCont(frmConstants.lstConstants)
                      
                Case Is = "dispatchmessage"
                      Call modTypes.TypeMessage(frmType.lstType)
                           
                Case Is = "getmessage"
                      Call modTypes.TypeMessage(frmType.lstType)
                      
                Case Is = "getquestatus"
                      Call modConstants.GetQueStatusConst(frmConstants.lstConstants)
                      
                Case Is = "peekmessage"
                      Call modTypes.TypeMessage(frmType.lstType)
                      Call modConstants.PeekMsgConst(frmConstants.lstConstants)
                      
                Case Is = "postmessage"
                      Call modConstants.PostMsgConst(frmConstants.lstConstants)
                      
                Case Is = "sendmessagetimeout"
                      Call modConstants.SendMsgTimeoutConst(frmConstants.lstConstants)
                      
                Case Is = "translatemessage"
                      Call modTypes.TypeMessage(frmType.lstType)
                
'API(Menu)========
                Case Is = "appendmenu"
                      Call modConstants.MenuConstants(frmConstants.lstConstants)
                
                Case Is = "getmenuiteminfo"
                      Call modTypes.TypeMenuItemInfo(frmType.lstType)
                      
                Case Is = "getmenuitemrect"
                      Call modTypes.TypeRect(frmType.lstType)
                      
                Case Is = "getsystemmenu"
                      Call modConstants.MenuConstants(frmConstants.lstConstants)
                      
                Case "insertmenuitem"
                      Call modTypes.TypeMenuItemInfo(frmType.lstType)
                
                Case Is = "modifymenu", "removemenu", "setmenuitembitmaps"
                      Call modConstants.MenuConstants(frmConstants.lstConstants)
                      
                Case Is = "setmenuiteminfo"
                      Call modTypes.TypeMenuItemInfo(frmType.lstType)
                      
                Case Is = "trackpopupmenu"
                      Call modConstants.TrackPopupMenu(frmConstants.lstConstants)
                      Call modTypes.TypeRect(frmType.lstType)
                      
                Case Is = "trackpopupmenuex"
                      Call modConstants.TrackPopupMenu(frmConstants.lstConstants)
'API(Rect/Rgn)========
                
                Case Is = "adjustwindowrect", "adjustwindowrectex"
                      Call modConstants.SetWindLongWSConst(frmConstants.lstConstants)
                      Call modTypes.TypeRect(frmType.lstType)
                      
                Case Is = "copyrect", "getclientrect", "getwindowrect", _
                          "inflaterect", "intersectrect", "isrectempty", _
                          "offsetrect", "ptinrect", "setrect", _
                          "setrectempty", "subtractrect", "unionrect", _
                          "createellipticrgnindirect", "createrectrgnindirect"
                      Call modTypes.TypeRect(frmType.lstType)
                      
                Case Is = "combinergn"
                      Call CombineRgnConst(frmConstants.lstConstants)
                      
                Case Is = "createpolygonrgn", "createpolypolygonrgn"
                      Call modTypes.PolygonRgn(frmType.lstType)
                      Call modConstants.PolygonRgnConst(frmConstants.lstConstants)
                      
 'API(Graphix)========
                Case Is = "bitblt", "stretchblt"
                      Call modConstants.BitSRCConst(frmConstants.lstConstants)
                      
                Case Is = "copyimage", "loadimage"
                      Call modConstants.CopyImage(frmConstants.lstConstants)
                
                Case Is = "createbrushindirect"
                      Call modTypes.LogBrush(frmType.lstType)
                      
                Case Is = "createdc"
                      Call modTypes.DEVMODE(frmType.lstType)
                
                Case Is = "createdibpatternbrushpt"
                      Call modConstants.DIB(frmConstants.lstConstants)
                      
                Case Is = "createdibsection"
                      Call modTypes.BITMAPINFO(frmType.lstType)
                      Call modConstants.DIB(frmConstants.lstConstants)
                      
                Case Is = "createhatchbrush"
                      Call modConstants.CreateHatchBrushConst(frmConstants.lstConstants)
                      
                Case Is = "createpen"
                     Call modConstants.PenStylesConst(frmConstants.lstConstants)
                     
                Case Is = "createpenindirect"
                     Call modTypes.LOGPEN(frmType.lstType)
                
                Case Is = "drawanimatedrects"
                     Call modTypes.TypeRect(frmType.lstType)
                     Call modConstants.IDANI(frmConstants.lstConstants)
                     
                Case Is = "drawcaption"
                     Call modConstants.DC(frmConstants.lstConstants)
                     
                Case Is = "drawedge"
                     Call modConstants.DrawEdge(frmConstants.lstConstants)
                     
                 Case Is = "drawfocusrect"
                     Call modTypes.TypeRect(frmType.lstType)
                     
                 Case Is = "drawframecontrol"
                     Call modConstants.DrawFrameControl(frmConstants.lstConstants)
                     
                 Case Is = "drawstate"
                      Call modConstants.DSS_DST(frmConstants.lstConstants)
                 
                 Case Is = "drawiconex"
                      Call modConstants.DI(frmConstants.lstConstants)
                 
                 Case Is = "extfloodfill"
                      Call modConstants.FloodFill(frmConstants.lstConstants)
                 
                 Case Is = "fillrect", "framerect"
                      Call modTypes.TypeRect(frmType.lstType)
                 
                 Case Is = "gdialphablend"
                      Call modTypes.BLENDFUNCTION(frmType.lstType)
                 
                 Case Is = "getbimapbits", "setbitmapbits"
                      Call modTypes.BITMAP(frmType.lstType)
                 
                 Case Is = "getcoloradjustment", "setcoloradjustment"
                      Call modTypes.COLORADJUSTMENT(frmType.lstType)
                 
                 Case Is = "getdibits", "setdibitstodevice"
                      Call modTypes.GetDiBits(frmType.lstType)
                      Call modConstants.DIB(frmConstants.lstConstants)
                      
                 Case Is = "getsyscolorbrush"
                      Call modConstants.SysColor(frmConstants.lstConstants)
                 
                 Case Is = "invertrect"
                      Call modTypes.TypeRect(frmType.lstType)
                      
                 Case Is = "loadcursor"
                      Call modConstants.IDC(frmConstants.lstConstants)
                 
                 Case Is = "loadicon"
                      Call modConstants.IDI(frmConstants.lstConstants)

                 Case Is = "patblt"
                      Call modConstants.PatBlt(frmConstants.lstConstants)

                 Case Is = "scrolldc"
                      Call modTypes.TypeRect(frmType.lstType)

                 Case Is = "setstretchbltmode"
                      Call modConstants.SetStretchBltmode(frmConstants.lstConstants)
                 
                 Case Is = "setsystemcursor"
                      Call modConstants.OCR(frmConstants.lstConstants)

'API(Mouse)========
                 Case Is = "clienttoscreen"
                      Call modTypes.PT(frmType.lstType)
                 
                 Case Is = "clipcursor"
                      Call modTypes.TypeRect(frmType.lstType)
                       
                 Case Is = "getcursorpos"
                      Call modTypes.PT(frmType.lstType)
                      
                 Case Is = "mouse_event"
                      Call modConstants.Mouse_Event(frmConstants.lstConstants)
'API(Window)========
 
                 Case Is = "animatewindow"
                      Call modConstants.AW(frmConstants.lstConstants)
                 
                 Case Is = "flashwindowex"
                      Call modTypes.FlashWindowEx(frmType.lstType)
                      
                 Case Is = "getclassinfo"
                      Call modTypes.WNDCLASS(frmType.lstType)
                
                 Case Is = "getititlebarinfo"
                      Call modTypes.TypeTitleBarInfo(frmType.lstType)
                
                 Case Is = "getwindow"
                      Call modConstants.GetWindowConst(frmConstants.lstConstants)
                
                 Case Is = "getwindowlong", "setwindowlong"
                      Call modConstants.GetWindowLongConst(frmConstants.lstConstants)
                 
                 Case Is = "getwindowplacement", "setwindowplacement"
                      Call modTypes.WINDOWPLACEMENT(frmType.lstType)
                 
                 Case Is = "redrawwindow"
                      Call modTypes.TypeRect(frmType.lstType)
                      Call modConstants.RedrawWindowConst(frmConstants.lstConstants)
                 
                 Case Is = "setwindowpos"
                      Call modConstants.SetWindPosConst(frmConstants.lstConstants)
                 
                 Case Is = "showwindow"
                      Call modConstants.ShowWindowConst(frmConstants.lstConstants)
                 
                 Case Is = "windowfrompoint"
                      Call modTypes.PT(frmType.lstType)
       End Select
'END CODE:

End Sub






'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | we split the api by " " scanning throught the
'             arr looking for the word "function" or "sub"
'             the spart() or word that follows is the API name
'----------------------------------------------------------------------
Private Function func_ParseAPIname(strAPI$) As String
  Dim sParts()   As String
  Dim i%
  
           sParts = Split(strAPI, " ")
           
           For i = 0 To UBound(sParts)
               If LCase(sParts(i)) = "function" Or LCase(sParts(i)) = "sub" Then
                   func_ParseAPIname = sParts(i + 1)
                   Exit Function
               End If
           Next i
End Function



















































 
 

'    .     '    . _       .'  _     .  '    .    .     .    .    .
' _ _'_   _  _ _'(_) _ __   / _| ___   ' ___  _    __ _  _ _' _ _'
'/ __| |_| |/ __|| || '_ ` | |_ / _ \  '/ __|| |  / _` |/ __|/ __|
'\__ \\__, |\__ \| || | | ||  _| (_)   ' (__ | |_  (_| |\__ \\__ \
'|___/|___/ |___/|_||_| |_||_|  \___/  '\___||___|\__,_||___/|___/


 


'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: | NONE
' COMMENTS: | shows used ram..event passed from class
'----------------------------------------------------------------------
Private Sub cSysInfo_RamInfo(sAvailRam As String, sUsedram As String)

       lblUsedRam = "used RAM: " & sUsedram & " mb"
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: | NONE
' COMMENTS: | this systems total avail physical ram..event passed from class
'----------------------------------------------------------------------
Private Sub cSysInfo_TotSysRam(sVal As String)

      lblTotalRam = sVal
End Sub
'----------------------------------------------------------------------
'   INPUTS: | data passed from class clsSysInfo
'  RETURNS: | NONE
' COMMENTS: | cpu load info
'----------------------------------------------------------------------
Private Sub cSysInfo_CPUloadInfo(sCpuPercentLoad As String)
 
       lblCpuLoad(1) = sCpuPercentLoad & " %"
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | initialize clsSysInfo and its object associations
'----------------------------------------------------------------------
Private Sub ReviveClassSysInfo()
 
       Set cSysInfo = New clsSysMonitor
       Set cSysInfo.picCPU = picCPUload
       Set cSysInfo.picRAM = picRAM
       Set cSysInfo.Timer = TimerRam
       cSysInfo.StartMonitoring
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | Destroy clsSysInfo and its object associations
'----------------------------------------------------------------------
Private Sub TerminateclassSysInfo()
    On Error Resume Next
    
       Set cSysInfo.Timer = Nothing
       Set cSysInfo.picCPU = Nothing
       Set cSysInfo.picRAM = Nothing
       Set cSysInfo = Nothing
End Sub


 

 

 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
'----------------------------------------------------------------------
' take a menu and format it into columns
'----------------------------------------------------------------------
Private Sub subMakeMenuColumns(mnuIndex&, mnuItemsPerColumn&)
Dim Buffer$, mnuItemText$
Dim i%, mnuItemCnt%
Dim mnuItemID&, hMenu&, hSubMenu&, Result&
 
    'Get the hwnd of the menu
   hMenu = GetMenu(Me.hwnd)
    'Get the hwnd of the submenu
   hSubMenu = GetSubMenu(hMenu, mnuIndex)
    'number of subitems in the menu
   mnuItemCnt = GetMenuItemCount(hSubMenu)
    
   'make the API window menu 2 columns because its so long
   For i = mnuItemsPerColumn& To mnuItemCnt Step mnuItemsPerColumn&
        'create MT string ready for data
        Buffer = Space$(256)
        Result = GetMenuString(hSubMenu, (i - 1), Buffer, Len(Buffer), &H400&)
        mnuItemText = Left$(Buffer, Result)
        mnuItemID = GetMenuItemID(hSubMenu, (i - 1))
        'Modify the menu to a column menu
        Call ModifyMenu(hSubMenu, (i - 1), &H400& Or &H20&, mnuItemID, mnuItemText)
   Next i
   
End Sub
















 
''  _     .    .          .  '    .    .    .     .     .    .    _.
' / _| ___  _ _  _ __ __     ' _ _  ___  _    __ _ _____  ___  __| |
'| |_ / _ \| '_\| '_ ` _ \   '| '_\/ _ \| |  / _` |_   _|/ _ \/ _  |
'|  _| (_) | |  | | | | | |  '| |    __/| |_  (_| | | |    __/ (_| |
'|_|  \___/|_|  |_| |_| |_|  '|_|  \___||___|\__,_| |_|  \___|\__,_|
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | we want to cut out the form so only the menu is left showing
'----------------------------------------------------------------------
Private Sub Form_Load()

On Error GoTo ERR:
 
'VARIABLES:
 Dim strComment$, strErr$
 Dim i%
 Dim Result&
 Dim arrProgData()    As Variant
 
'CODE:
   Visible = False
   
   Set cClipboard = New clsClipboard
   'load all the API menus from res file
   Call LoadApiMenus(api_window)
   Call LoadApiMenus(api_Mouse)
   Call LoadApiMenus(api_Graphx)
   Call LoadApiMenus(api_RectRgn)
   Call LoadApiMenus(api_Menu)
   Call LoadApiMenus(api_Msg)
   Call LoadApiMenus(api_Misc)
   Call LoadApiMenus(api_TypesConst)
  
    'make longest menus 2 column
   Call subMakeMenuColumns(3, 35)
   Call subMakeMenuColumns(5, 35)

   strErr = "  On Error goto ERR:" & vbCrLf & vbCrLf & _
           "exit sub" & vbCrLf & _
           "ERR:" & vbCrLf & _
           "   If Err.number <> 0 then" & vbCrLf & _
           "        msgbox Err.number & vbcrlf & Err.description" & vbCrLf & _
           "   End if"
              

   strComment = "'" & String(50, "-") & vbCrLf & _
              "'   INPUTS: |" & vbCrLf & _
              "'  RETURNS: |" & vbCrLf & _
              "' COMMENTS: |" & vbCrLf & _
              "'" & String(50, "-")
              
   'retrieve errhandler string and procedure
   'header comment string from reg and plad
   'in public var
   'when frmcomment or frmerrhandler are shown
   'contents of these variables will be placed
   'in that forms textbox
   strErrHandler = GetSetting("VB_ClipboardBuddy", _
                  "ErrHandler", _
                  "Value", _
                  strErr)
                  
   strHeaderComment = GetSetting( _
                   "VB_ClipboardBuddy", _
                   "Comment", _
                   "Value", _
                   strComment)
              
  'this carves window regions to initally
  'show computer stats..ie ram,cpu load, etc
   Call mnuOptionsShowCompStats_Click
   
  'start form center/left screen
   Me.Move 20, (Screen.Height * 0.5)

   'this timer will animate the form upwards
   b_Loading = True
  
   'load programs data(function in modPub)
   arrProgData = func_FileLoadData(App.Path & "\progdata.txt")
   
   For i = LBound(arrProgData) To UBound(arrProgData)
         Select Case i
             Case Is = 0
                   'with this checked there are occasional
                   'wav files played to assist the use
                   bEnableScreenTips = CBool(arrProgData(i))
                   'now click the approp menu item
                   If bEnableScreenTips = True Then
                        mnuEnableScreenHelp_Click
                   Else
                        mnuDisableScreenHelp_Click
                   End If
                   '
             Case Is = 1
                   'with this true, a small, rect panel on
                   'the left is shown displaying computer stats
                   bShowComputerStats = CBool(arrProgData(i))
                    'now click the approp menu item
                    If bShowComputerStats = True Then
                         mnuOptionsShowCompStats_Click
                    Else
                         mnuOptionsHideCompStats_Click
                    End If
                    '
             Case Is = 2
                    'if true, when user pastes and API, form(s)
                    'that show const/types required by the api
                    'are/is shown
                    bAutoShowConstForm = CBool(arrProgData(i))
                    'now click the approp menu item
                    If bAutoShowConstForm = True Then
                         mnuAutoShowContTypesForm_Click
                    Else
                         mnudontShowContTypesForm_Click
                    End If
             
             Case Is = 3
                  'how often pop mail (frmPOP) is checked
                   MailCheckFrequency = CInt(arrProgData(i))
             Case Is = 4
                  'if pop mail is checked
                   bEnablePOPchecking = CBool(arrProgData(i))
                   '
                   If bEnablePOPchecking = True Then
                         Call mnuOptionsEnableEmailChecking_Click
                   Else
                         Call mnuOptionsDisableEmailChecking_Click
                   End If
              Case Is = 5
                   sMailClientPath = CStr(arrProgData(i))
                   
   ':::::::: THE FOLLOW ARE THE POP SERVER ADDRESS/USERNAME/:::::
   ':::::::: AND PASSWORD HOLDERS FOR 4 POP MAIL ACCOUNTS :::::::
             Case Is = 6
                   POPaccountInfo(0, 0) = CStr(arrProgData(i))
             Case Is = 7
                   POPaccountInfo(0, 1) = CStr(arrProgData(i))
             Case Is = 8
                   POPaccountInfo(0, 2) = CStr(arrProgData(i))
             Case Is = 9
                   POPaccountInfo(0, 3) = CStr(arrProgData(i))
             Case Is = 10
                   POPaccountInfo(1, 0) = CStr(arrProgData(i))
             Case Is = 11
                   POPaccountInfo(1, 1) = CStr(arrProgData(i))
             Case Is = 12
                   POPaccountInfo(1, 2) = CStr(arrProgData(i))
             Case Is = 13
                   POPaccountInfo(1, 3) = CStr(arrProgData(i))
             Case Is = 14
                   POPaccountInfo(2, 0) = CStr(arrProgData(i))
             Case Is = 15
                   POPaccountInfo(2, 1) = CStr(arrProgData(i))
             Case Is = 16
                   POPaccountInfo(2, 2) = CStr(arrProgData(i))
             Case Is = 17
                   POPaccountInfo(2, 3) = CStr(arrProgData(i))
         End Select
    Next i

  'set labels to autosize so no matter what the values
  'the labels can expand to show entire contents and
  'this we avoid clumsy looking wordwrap
  lblUsedRam.AutoSize = True
  lblTotalRam.AutoSize = True
  
  Visible = True
  
  TimerAnimate.Interval = 10
  TimerAnimate.Enabled = True
  
 'this on top
  SetWindowPos hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
       
  On Error GoTo ERR:

Exit Sub
ERR:
   If ERR.Number <> 0 Then
        MsgBox ERR.Number & vbCrLf & ERR.Description
   End If
End Sub
'----------------------------------------------------------------------
'   INPUTS: |NONE
'  RETURNS: |NONE
' COMMENTS: |turn off timers and delete anything that can rob memory
'----------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'VARIABLES:
  Dim CTL As Control
'CODE:
 On Error GoTo ERR:
      
             Set cClipboard = Nothing
            'save program data to file (modPub variables)
             Call sub_FileSaveData(App.Path & "\progdata.txt", _
                                         bEnableScreenTips, _
                                         bShowComputerStats, _
                                         bAutoShowConstForm, _
                                         MailCheckFrequency, _
                                         bEnablePOPchecking, _
                                         sMailClientPath, _
                                         POPaccountInfo(0, 0), _
                                         POPaccountInfo(0, 1), _
                                         POPaccountInfo(0, 2), _
                                         POPaccountInfo(0, 3), _
                                         POPaccountInfo(1, 0), _
                                         POPaccountInfo(1, 1), _
                                         POPaccountInfo(1, 2), _
                                         POPaccountInfo(1, 3), _
                                         POPaccountInfo(2, 0), _
                                         POPaccountInfo(2, 1), _
                                         POPaccountInfo(2, 2), _
                                         POPaccountInfo(2, 3))
   
        'kill all timers
        TimerRam.Interval = 0
        TimerRam.Enabled = False
        TimerPaste.Interval = 0
        TimerPaste.Enabled = False
        TimerAnimate.Interval = 0
        TimerAnimate.Enabled = False
        timerMailCheck.Interval = 0
        timerMailCheck.Enabled = False
        'kill class objects
        Call TerminateclassSysInfo
        Set cPOP.YourPopSock = Nothing
        Set cPOP = Nothing
 
'END CODE:
Exit Sub
ERR:
   If ERR.Number <> 0 Then
        If ERR.Number = 91 Then
            ERR.Clear
            Exit Sub
        Else
           MsgBox ERR.Number & vbCrLf & ERR.Description
        End If
   End If
End Sub
 
 
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | End App
'----------------------------------------------------------------------
Private Sub mnuClose_Click()
 Dim i                     As Integer
 
        For i = Forms.Count - 1 To 0 Step -1
            DoEvents
            Unload Forms(i)
        Next i
        
        End
End Sub
'----------------------------------------------------------------------------
'   INPUTS: |NONE
'  RETURNS: |NONE
' COMMENTS: |we use setwindowpos  in the paint instead of form load
'            to insure that no matter what this stays on top
'            if another window tries to go on top of this one
'            it will invoke a paint, then this will be back on top
'----------------------------------------------------------------------------
Private Sub Form_Paint()

       SetWindowPos _
             Me.hwnd, _
            -1, _
             0, 0, 0, 0, _
             &H1 Or &H2
End Sub
'----------------------------------------------------------------------
'   INPUTS: |NONE
'  RETURNS: |NONE
' COMMENTS: |keep this form nothing but a titlebar (to send messages
'            to the user)  and a menu bar(this stays out of VB'er's way
'----------------------------------------------------------------------
Private Sub Form_Resize()
        
      If b_Loading = False Then
           Me.Move 100, -250
           Width = 13000
      End If
End Sub




 
 





















 

 

 
 

 

 

 
 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

 

'          .    .      .     .
'1st MENU LIST ON LEFT:::::::::::::::::::::::::::::::::::::::::::
'| '_ ` _ \ / _ \| '_ `  | | |
'| | | | | |  __/| | | | |_| |
'|_| |_| |_|\___||_| |_|\__,_|
'     . _   .  ''  _      .      .    .     . _     .      .    .
'_   __| |__   ' / _|_   _  _ __   ___ _____ (_) ___  _ __   _ _'
' \ / /|  _ \  '| |_  | | || '_ ` / __|_   _|| |/ _ \| '_ ` / __|
'\ V / | |_)   '|  _| |_| || | | | (__  | |  | | (_) | | | |\__ \
' \_/  |_.__/  '|_|  \__,_||_| |_|\___| |_|  |_|\___/|_| |_||___/

 
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | user wants to view the contents of procedure header comment
'----------------------------------------------------------------------
Private Sub mnuPrcHeadCommentViewEdit_Click()

       FrmComment.Show
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | send For I Loop to next app with focus
'----------------------------------------------------------------------
Private Sub mnuForI_Click()

       Call SendData("For I = 0 to Ubound()" & vbCrLf & _
                     "  Doevents" & vbCrLf & vbCrLf & _
                     "Next I" & vbCrLf, "click where you wish to paste  For I  loop")
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | send MsgBox shell (vbYesNo)
'----------------------------------------------------------------------
Private Sub mnuMsgboxShellvbYesNo_Click()
  Dim str               As String
  
       str = "Dim iYesNo      as integer" & vbCrLf & _
             "Dim sText       as string" & vbCrLf & _
             "Dim sTitle      as string" & vbCrLf & _
             "    iYesNo = MsgBox(sText, vbYesNo, sTitle)" & vbCrLf & _
             "    'If user selects yes then   " & vbCrLf & _
             "    If iYesNo = vbYes then" & vbCrLf & vbCrLf & _
             "    End If"
             
       Call SendData(str, "click where you wish to paste  msgbox shell")
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | send InputBox Shell
'----------------------------------------------------------------------
Private Sub mnuInpBoxShell_Click()
  Dim str        As String
  
       str = "Dim sInp           as string" & vbCrLf & _
            "   sInp = InputBox(sPrompt,sDefault)" & vbCrLf & _
            "   '       " & vbCrLf & _
            "   If sInp = sVal Then" & vbCrLf & vbCrLf & _
            "   End If"
            
      Call SendData(str, "click where you wish to paste  inputbox shell")
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | send "On Error Resume Next"
'----------------------------------------------------------------------
Private Sub mnuErrResume_Click()

       Call SendData("  On Error Resume Next", "click where you wish to paste  Error handler")
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | send "If..Elseif..End If" structure to next app with focus
'----------------------------------------------------------------------
Private Sub mnuIfElseifEndif_Click()
  Dim i                      As Integer
  Dim ii                     As Integer
  Dim sTemp                  As String
  
  On Error GoTo ERR:
  
       'default "case ="  is 3
       i = CInt(InpBox( _
           "How many " & Chr(34) & "ElseIf's" & Chr(34) & " do you want", _
           2))
           
      'beginnin part of select case
      sTemp = "If condition Then" & vbCrLf & vbCrLf
      'add additional "Case =" with each loop
      For ii = 1 To i
          sTemp = (sTemp & "Elseif condition Then") & vbCrLf & vbCrLf
      Next ii
      'add end part of select case
      sTemp = (sTemp & "End If" & vbCrLf)
      'send data to clipboard
      Call SendData(sTemp, "click where you wish to paste  If structure")

        
   Exit Sub
ERR:
   If ERR.Number <> 0 Then
      If ERR.Number = 13 Then
          Exit Sub
      End If
    End If
End Sub

 

 

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | send Select case structure to next app with focus
'----------------------------------------------------------------------
Private Sub mnuSelectCase_Click()
  Dim i                      As Integer
  Dim ii                     As Integer
  Dim sTemp                  As String
  
  On Error GoTo ERR:
  
       'default "case ="  is 3
       i = CInt(InpBox( _
           "How many " & Chr(34) & "Case =" & Chr(34) & " do you want", _
           3))
           
      'beginnin part of select case
      sTemp = "Select Case " & vbCrLf
      'add additional "Case =" with each loop
      For ii = 0 To i
          sTemp = (sTemp & "    Case = " & ii) & vbCrLf & vbCrLf
      Next ii
      'add end part of select case
      sTemp = (sTemp & "End Select" & vbCrLf)
      'send data to clipboard
      Call SendData(sTemp, "click where you wish to paste  Select Case structure")

        
   Exit Sub
ERR:
   If ERR.Number <> 0 Then
      If ERR.Number = 13 Then
          Exit Sub
      End If
    End If
End Sub



'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | This will send a Sub routine shell to next app with focus
'----------------------------------------------------------------------
Private Sub mnuSubShell_Click()
 Dim sTemp(2)                     As String

  On Error GoTo ERR:
  
       'get a name for the sub
       sTemp(0) = InpBox( _
           "Enter a name for the subroutine")
           
        'get the desired scope
       sTemp(1) = InpBox( _
           "Public, Private, or Friend", "Private")
          
        'include errhandler?
       If Trim(sTemp(2)) = InpBox( _
             "Include basic Err.handling ?", "Yes") _
             = "yes" Then
            sTemp(2) = strErrHandler
        Else
            sTemp(4) = vbCrLf & vbCrLf
        End If
        
       DoEvents
       
       'place sub onto clipboard
       Call SendData( _
            sTemp(1) & " Sub " & sTemp(0) & sTemp(2) & _
            "End Sub" & vbCrLf, "click where you wish to paste  Sub shell")
            
   Exit Sub
ERR:
   If ERR.Number <> 0 Then
      If ERR.Number = 13 Then
          Exit Sub
      End If
    End If
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | This will send a Function routine shell to next app with focus
'----------------------------------------------------------------------
Private Sub mnuFunction_Click()
  Dim i                     As Integer
  Dim sTemp(4)              As String
  Dim sPrompt(3)            As String
  Dim sDefault(3)           As String

  sPrompt(0) = "Enter a name for the Function"
  sPrompt(1) = "Public, Private, or Friend"
  sPrompt(2) = "Arguments...seperate with comma's"
  sPrompt(3) = "Functions Return Value"
  sDefault(0) = ""
  sDefault(1) = "Private"
  sDefault(2) = ""
  sDefault(3) = "String"
  
         '4 input boxes to get the values we need
       For i = 0 To 3
           sTemp(i) = InpBox( _
           sPrompt(i), sDefault(i))
           'if any of the vals entered are empty
           'then the user must want to cancel
           If Trim(sTemp(i)) = "" Then
             'its possible for there not to be args
             If i <> 2 Then
                 Exit Sub
             End If
           End If
       Next i
            
        'include errhandler?
        sTemp(4) = InpBox( _
             "Include basic Err.handling ?", _
             "Yes")
        
        If Trim(LCase(sTemp(4))) = "yes" Then
             sTemp(4) = "  On Error goto ERR:" & vbCrLf & sTemp(0) & " = " _
                      & Chr(34) & Chr(34) & vbCrLf & vbCrLf & _
                     "exit sub" & vbCrLf & _
                     "ERR:" & vbCrLf & _
                     "   If Err.number <> 0 then" & vbCrLf & _
                     "       msgbox Err.number & vbcrkf & Err.description" & vbCrLf & _
                     "   End if"
        Else
            sTemp(4) = vbCrLf & sTemp(0) & " = " _
                      & Chr(34) & Chr(34) & vbCrLf
        End If
        
       'place sub onto clipboard
       Call SendData( _
            sTemp(1) & " Function " & sTemp(0) & _
            "(" & sTemp(2) & ") as " & sTemp(3) & vbCrLf & _
            sTemp(4) & vbCrLf & _
            "End Function", "click where you wish to paste  Function shell")

   Exit Sub
ERR:
   If ERR.Number <> 0 Then
      If ERR.Number = 13 Then
          Exit Sub
      End If
    End If
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: |  This will send a Property let/get shell to next app with focus
'----------------------------------------------------------------------
Private Sub mnuPropLetGetShell_Click()
  Dim i                     As Integer
  Dim sTemp(3)              As String
  Dim sPrompt(3)            As String
  Dim sDefault(3)           As String
  Dim sComma                As String

  sPrompt(0) = "Enter a name for the Property"
  sPrompt(1) = "Public, Private, or Friend"
  sPrompt(2) = "Arguments...seperate with comma's"
  sPrompt(3) = "Properties Return Value"
  sDefault(0) = ""
  sDefault(1) = "Friend"
  sDefault(2) = ""
  sDefault(3) = "String"
       
       '4 input boxes to get the values we need
       For i = 0 To 3
           sTemp(i) = InpBox( _
           sPrompt(i), sDefault(i))
           'if any of the vals entered are empty
           'then the user must want to cancel
           If Trim(sTemp(i)) = "" Then
             'its possible for there not to be args
              If i <> 2 Then
                 Exit Sub
              End If
           End If
       Next i
       
       'insert comma if there are arguments to seperate
       If sTemp(2) <> "" Then
             sComma = ", "
       Else
             sComma = ""
       End If
       
       'send the property shell
       Call SendData( _
             sTemp(1) & " Property Get " _
             & sTemp(0) & "(" & sTemp(2) & ") As " & sTemp(3) _
             & vbCrLf & vbCrLf & _
             "End Property" & vbCrLf & _
             sTemp(1) & " Property Let " & sTemp(0) _
             & "(" & sTemp(2) & sComma & " ByVal vNewValue as " & sTemp(3) & ")" _
             & vbCrLf & vbCrLf & _
             "End Property", "click where you wish to paste  Property shell")
End Sub

 

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | send While...Wend to next app with focus
'----------------------------------------------------------------------
Private Sub mnuWhileWend_Click()

       Call SendData("While condition" & vbCrLf & _
                     "  Doevents" & vbCrLf & vbCrLf & _
                     "Wend" & vbCrLf, "click where you wish to paste  While Wend")
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | send Do...Loop to next app with focus
'----------------------------------------------------------------------
Private Sub mnuDoLoop_Click()

       Call SendData("Do" & vbCrLf & _
                     "  Doevents" & vbCrLf & vbCrLf & _
                     "Loop until condition" & vbCrLf, "click where you wish to paste Do loop")
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | send If..End If to next app with focus
'----------------------------------------------------------------------
Private Sub mnuIfEndIf_Click()

       Call SendData("If condition Then" & _
                     vbCrLf & vbCrLf & _
                     "End If" & vbCrLf, "click where you wish to paste  If structure")
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | send Procedure header comment
'----------------------------------------------------------------------
Private Sub mnuPrcHeadCommentInsert_Click()

       Call SendData(strHeaderComment, "click where you wish to paste  Header comment")
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | view full errhandler string/show frmErrorHandler
'----------------------------------------------------------------------
Private Sub mnuFullErrHandlerViewEdit_Click()
       
       frmErrHandler.Show
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | send full errhandler string
'----------------------------------------------------------------------
Private Sub mnuFullErrHandlerInsert_Click()

       Call SendData(strErrHandler, "click where you wish to paste  Error handler")
End Sub











































 
 
 
























 '          .    .      .     .  '    . _ __      . _     .      .    .
' _ __ __    ___  _ __  _   _   ' ___ | '_ \_____ (_) ___  _ __   _ _'
'| '_ ` _ \ / _ \| '_ `  | | |  '/ _ \| |_) _   _|| |/ _ \| '_ ` / __|
'| | | | | |  __/| | | | |_| |  ' (_) | .__/ | |  | | (_) | | | |\__ \
'|_| |_| |_|\___||_| |_|\__,_|  '\___/|_|    |_|  |_|\___/|_| |_||___/

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | load RAM purge .exe
'----------------------------------------------------------------------
Private Sub mnuLoadRamPurge_Click()
    'run RamPurge.exe
    Call LoadRamPurge
End Sub
Private Sub btnLoadRamPurge_Click()
    'run RamPurge.exe
    Call LoadRamPurge
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | show left pane that reveals computer stats ie ram ,cpu load
'----------------------------------------------------------------------
Private Sub mnuOptionsShowCompStats_Click()
 Dim hRgnTop                 As Long
 Dim hRgnLeft                As Long
 Dim RgnCombined             As Long
 
  If b_SysInfoVisible = False Then

    'top thin bar part
    'top thin bar part
     hRgnTop = CreateRectRgn(0, 7, _
        (Width / Screen.TwipsPerPixelX) * 0.95, 42)
     
     'the panel part that will display
     ' RAM and other stuff
     hRgnLeft = CreateRoundRectRgn(0, 35, 100, 210, 5, 5)
 
    'the pallette to paint the 2 previous regions on to
     RgnCombined = CreateRectRgn(0, 0, 0, 0)
     
     CombineRgn RgnCombined, hRgnTop, hRgnLeft, RGN_OR

     SetWindowRgn Me.hwnd, RgnCombined, True
     
     'raise class clsSysInfo from the dead
     Call ReviveClassSysInfo
    
     b_SysInfoVisible = True
 End If
 
 'toggle public variable
  bShowComputerStats = True
  
 'toggle check this menu item
 Call ToggleMenuOn(mnuOptionsShowCompStats, mnuOptionsHideCompStats)
 
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | hide left pane that reveals computer stats ie ram ,cpu load
'----------------------------------------------------------------------
Private Sub mnuOptionsHideCompStats_Click()
 Dim hRgnTop                 As Long
 
 On Error Resume Next
 
    'top thin bar part
     hRgnTop = CreateRectRgn(0, 7, _
        (Width / Screen.TwipsPerPixelX) * 0.95, 42)

       SetWindowRgn _
              Me.hwnd, _
              hRgnTop, _
              True
              
       'kill the associated class
       Call TerminateclassSysInfo
       
       b_SysInfoVisible = False
       
       'toggle public variable
       bShowComputerStats = False
       
       'toggle check this menu item
       Call ToggleMenuOn(mnuOptionsHideCompStats, mnuOptionsShowCompStats)
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | auto show form(s) that have the const/types needed
'             by and apicall that was just pasted
'----------------------------------------------------------------------
Private Sub mnuAutoShowContTypesForm_Click()
'
       bAutoShowConstForm = True
       Call ToggleMenuOn(mnuAutoShowContTypesForm, mnuDontShowContTypesForm)
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | dont show form(s) that have the const/types needed
'             by and apicall that was just pasted
'----------------------------------------------------------------------
Private Sub mnudontShowContTypesForm_Click()
'
       bAutoShowConstForm = False
       Call ToggleMenuOn(mnuDontShowContTypesForm, mnuAutoShowContTypesForm)
End Sub
 
  
  
  
  
  
  
  
  
  
  
  
  
  
  







'    .    .    .     .    .  ''  _      .      .    .     . _     .      .    .
' _    ___  ___  __ _  _     ' / _|_   _  _ __   ___ _____ (_) ___  _ __   _ _'
'| |  / _ \/ __|/ _` || |    '| |_  | | || '_ ` / __|_   _|| |/ _ \| '_ ` / __|
'| |_  (_)  (__  (_| || |_   '|  _| |_| || | | | (__  | |  | | (_) | | | |\__ \
'|___|\___/\___|\__,_||___|  '|_|  \__,_||_| |_|\___| |_|  |_|\___/|_| |_||___/



 
'----------------------------------------------------------------------
'   INPUTS: | Data to paste to the clipboard(sDataToSend$)
'  RETURNS: | NONE
' COMMENTS: | this starts the timer that will wait for focus(foreground window)
'             to shift away from this window to another. When that happens
'             is will send clipboard contents to the new app that has focus
'----------------------------------------------------------------------
Public Sub SendData(sDataToSend$, Optional strScreenMsg$)
'VARIABLES:
  
'CODE:
'==='start timer that monitors for when to past
      TimerPaste.Interval = 100
      TimerPaste.Enabled = True
'===place api string on clipboard
      Clipboard.Clear
      Clipboard.SetText sDataToSend$
      strClipBoardData = sDataToSend
      
      'show screen msg
      If Len(Trim(strScreenMsg)) > 0 Then
         If bEnableScreenTips = True Then
            Set cScreenMsg = New clsScreenText
            cScreenMsg.ScreenMsg _
                    "click where you wish to immediately send data, or " & vbCrLf & _
                    "press ESC. and paste (Ctrl+V)", 12, vbBlue, _
                    True, 2500, 3000
            Set cScreenMsg = Nothing
         End If
      End If
'END CODE:

End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | This will show input box and parse to make sure data isnt
'             empty and then pass back to calling routine
'----------------------------------------------------------------------
Private Function InpBox(sPrompt$, Optional sDefault$) As String

             InpBox = InputBox(sPrompt, , sDefault)
             
             If Len(Trim(InpBox)) = 0 Then
                  InpBox = ""
             End If
End Function
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | Show form to set up email checking accounts
'              frame being clicked on
'----------------------------------------------------------------------
Private Sub fmeEmail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

       If Button = 2 Then
             frmPOP.Show vbModeless, Me
       End If
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | Show form to set up email checking accounts
'             label being clicked on
'----------------------------------------------------------------------
Private Sub lblPopName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

      If Button = 2 Then
          frmPOP.Show vbModeless, Me
      End If
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | Show form to set up email checking accounts
'             Options menu (Email Options) being clicked on
'----------------------------------------------------------------------
Private Sub mnuOptionsEmail_Click()

      frmPOP.Show vbModeless, Me
End Sub

'----------------------------------------------------------------------
'   INPUTS: | MENU ITEM TO CHECK, MENU ITEMS TO UNCHECK
'  RETURNS: | NONE
' COMMENTS: | CHECKS A MENU ITEM AND UNCHECKS ALL LISTED IN MenuNamesUncheck
'----------------------------------------------------------------------
Sub ToggleMenuOn(MenuNameCheck As menu, ParamArray MenuNamesUncheck())
 Dim i         As Integer
   
       For i = 0 To UBound(MenuNamesUncheck)
             MenuNamesUncheck(i).Checked = False
       Next i
       
       MenuNameCheck.Checked = True
End Sub























'    . _ __      . _     .      .    .  '          .    .      .     .
' ___ | '_ \_____ (_) ___  _ __   _ _'  ' _ __ __    ___  _ __  _ _
'/ _ \| |_) _   _|| |/ _ \| '_ ` / __|  '| '_ ` _ \ / _ \| '_ `  | | |
' (_) | .__/ | |  | | (_) | | | |\__ \  '| | | | | |  __/| | | | |_| |
'\___/|_|    |_|  |_|\___/|_| |_||___/  '|_| |_| |_|\___||_| |_|\__,_|



'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: | NONE
' COMMENTS: | runs the RamPurge Program
'----------------------------------------------------------------------
Private Sub LoadRamPurge()
        'if the program isnt compiled then notify
        If Dir(App.Path & "\RamPurge\RamPurge.exe") = "" Then
            MsgBox "Compile the RamPurge project included" & vbCrLf & _
                   "in this programs folder saving it as" & vbCrLf & _
                    Chr(34) & "RamPurge.exe" & Chr(34)
        Else
            'run it
            Call ShellExecute( _
                  hwnd, "open", App.Path & "\RamPurge\RamPurge.exe", _
                  vbNull, "c:\", 1 _
                  )
        End If
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |VOICE HELP/TIPS ENABLED
'----------------------------------------------------------------------
Private Sub mnuEnableScreenHelp_Click()

       bEnableScreenTips = True
       'toggle this menu item on
       Call ToggleMenuOn(mnuEnableScreenHelp, mnuDisableScreenHelp)
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |VOICE HELP/TIPS DISABLED
'----------------------------------------------------------------------
Private Sub mnuDisableScreenHelp_Click()

        bEnableScreenTips = False
       'toggle this menu item on
       Call ToggleMenuOn(mnuDisableScreenHelp, mnuEnableScreenHelp)
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | email checking enabled
'----------------------------------------------------------------------
Private Sub mnuOptionsEnableEmailChecking_Click()
      
      bEnablePOPchecking = True
      LastMailCheckTickCount = GetTickCount
      timerMailCheck.Interval = 60000
      timerMailCheck.Enabled = True
      Call ToggleMenuOn(mnuOptionsEnableEmailChecking, mnuOptionsDisableEmailChecking)
      Call timerMailCheck_Timer
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | email checking Disabled
'----------------------------------------------------------------------
Private Sub mnuOptionsDisableEmailChecking_Click()
      
      bEnablePOPchecking = False
      timerMailCheck.Interval = 0
      timerMailCheck.Enabled = False
      Call ToggleMenuOn(mnuOptionsDisableEmailChecking, mnuOptionsEnableEmailChecking)
End Sub

































'    .    . _  _ __  _   .    .     .    .    _.
' ___  _   (_)| '_ \| |__  ___  __ _  _ _  __| |
'/ __|| |  | || |_) |  _ \/ _ \/ _` || '_\/ _  |
' (__ | |_ | || .__/| |_)  (_)  (_| || |   (_| |
'\___||___||_||_|   |_.__/\___/\__,_||_|  \__,_|




 

'add single item to clipboard
Private Sub mnuClipboardAdd_Click()
       '
       frmClipboard.enumDataType = ClipboardData
       frmClipboard.Show vbModeless, Me
End Sub
'add single item to Code Block
Private Sub mnuAddCodeBlock_Click()
        '
       frmClipboard.enumDataType = CodeBlockData
       frmClipboard.Show vbModeless, Me
End Sub






'clipboard menu item clicked on
Private Sub mnuArrClipboard_Click(Index As Integer)

        'this prevents frmConst or frmTypes
        'from show after code paste
        modPub.bSendingApiNotConst = False
        Call ClipboardOrCodeBlockMenuItemClick(Index, ClipboardItem)
End Sub

'Code menu item clicked on
Private Sub mnuArrCode_Click(Index As Integer)

        'this prevents frmConst or frmTypes
        'from show after code paste
        modPub.bSendingApiNotConst = False
        Call ClipboardOrCodeBlockMenuItemClick(Index, CodeBlock)
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |either the clipboard or codeblock menu has been clicked on
'----------------------------------------------------------------------
Private Sub ClipboardOrCodeBlockMenuItemClick(Index As Integer, Which As enumClipboardOrCodeBlock)
 Dim i%
        If Which = ClipboardItem Then
            If bRemoveMenuMode = False Then
                 ' paste/send contents of clipboard
                  Call SendData(cClipboard.CurrClipContents(Clip_Board, Index), _
                          "the next place you click the mouse the clipboard text " & _
                          "will be sent...or...press  Esc  then  Ctrl + V when you " & _
                          "wish to paste")
                  
            'user wants to delete a menu entry
            Else
                bRemoveMenuMode = False
                cClipboard.RemoveFromClipboard Clip_Board, Index, mnuArrClipboard
            End If
        Else
            If bRemoveMenuMode = False Then
                  Call SendData(cClipboard.CurrClipContents(Code_Block, Index), _
                          "the next place you click the mouse the clipboard text " & _
                          "will be sent...or...press  Esc  then  Ctrl + V when you " & _
                          "wish to paste")
                  
            'user wants to delete a menu entry
            Else
                bRemoveMenuMode = False
                cClipboard.RemoveFromClipboard Code_Block, Index, mnuArrCode
            End If
        End If
End Sub
'save clipboard file
Public Sub mnuClipboardSave_Click()
        '
        cClipboard.SaveClipContents Clip_Board, Me.cmDlg
End Sub
'save code block file
Public Sub mnuSaveCodeBlocksToFile_Click()
        '
        cClipboard.SaveClipContents Code_Block, Me.cmDlg
End Sub
'removes single item from code block
Private Sub mnuRemoveCodeBlock_Click()
        '
      bRemoveMenuMode = True
      frmWaitForClipboardRemove.Show
End Sub


'removes single item from clipboard
Private Sub mnuClipboardRemove_Click()
      '
      bRemoveMenuMode = True
      frmWaitForClipboardRemove.Show
End Sub
'load clipboard file
Private Sub mnuClipboardLoad_Click()
        '
      cClipboard.LoadClipContents Clip_Board, mnuArrClipboard, Me.cmDlg
End Sub
'load code block file
Private Sub mnuLoadCodeFile_Click()
        '
       cClipboard.LoadClipContents Code_Block, mnuArrCode, Me.cmDlg
End Sub
'clear menu items from clipboard menu
Private Sub mnuClipboardClear_Click()
        '
       cClipboard.ClearAll Clip_Board
End Sub
'clear menu items from code block menu
Private Sub mnuClearCodeBlocksFromMenu_Click()
        '
       cClipboard.ClearAll Code_Block
End Sub










' _ __     . _ __   '          .     . _     .
'| '_ \ ___ | '_ \  ' _ __ __    __ _ (_) _
'| |_) / _ \| |_)   '| '_ ` _ \ / _` || || |
'| .__/ (_) | .__/  '| | | | | | (_| || || |_
'|_|   \___/|_|     '|_| |_| |_|\__,_||_||___|

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |number of mail messages waiting on mailserver
' COMMENTS: |
'----------------------------------------------------------------------
Sub cPOP_POPmsgsWaiting(NumMsgs As Integer)
 Dim bMsgWaiting As Boolean
  
 
        'means there are messages waiting on POP
       If NumMsgs > 0 Then
          Dim strFile        As String
             
             'select the mailbox number wav based upon
             'which mail account number has messages right now
             '----------
             'var bMailFlash() tells us which labels in the
             'computer statistics panel
             Select Case iPub_CurrPopAccountNum
                   Case Is = 0
                         strFile = "inmailbox1.wav"
                         bMsgWaiting = True
                         bMailFlash(0) = True
                   Case Is = 1
                         strFile = "inmailbox2.wav"
                         bMsgWaiting = True
                         bMailFlash(1) = True
                   Case Is = 2
                         strFile = "inmailbox2.wav"
                         bMsgWaiting = True
                         bMailFlash(2) = True
                   Case Is = 3
                         strFile = "inmailbox3.wav"
                         bMsgWaiting = True
                         bMailFlash(3) = True
             End Select
              
             PlaySound App.Path & "\sounds\youhave.wav", 0&, SND_SYNC Or SND_NODEFAULT
             PlaySound App.Path & "\sounds\" & NumMsgs & ".wav", 0&, SND_SYNC Or SND_NODEFAULT
             PlaySound App.Path & "\sounds\messages.wav", 0&, SND_SYNC Or SND_NODEFAULT
             PlaySound App.Path & "\sounds\" & strFile, 0&, SND_SYNC Or SND_NODEFAULT
             
 
             'if the panel that will flash mail waiting is visible
             If bShowComputerStats = True Then
               If bMsgWaiting = True Then
                   '
                   If TimerMailFlash.Enabled = False Then
                      TimerMailFlash.Interval = 1000
                      TimerMailFlash.Enabled = True
                   End If
               End If
             End If
       End If


      'done checking this accounts mail..if we havent
      'checked the last one then check the next one
       If iPub_CurrPopAccountNum < 3 Then
          iPub_CurrPopAccountNum = (iPub_CurrPopAccountNum + 1)
          Set cPOP.YourPopSock = Nothing
          Set cPOP = Nothing
          Call subGetPopmail(iPub_CurrPopAccountNum)
       Else
          'this means there are no msg waiting so
          'turn of TimerMailFlash
            If bMsgWaiting = False Then
                      If TimerMailFlash.Enabled = True Then
                          Call TimerFlashOff
                      End If
            End If
       End If
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |check mail
'----------------------------------------------------------------------
Sub subGetPopmail(iCurrPopAccountNum As Integer)
  Dim i         As Integer
         
  On Error GoTo ERR:
    'public var keeps track of account to check next
    iPub_CurrPopAccountNum = iCurrPopAccountNum
    
    Set cPOP = New clsPOP
    Set cPOP.YourPopSock = sockPOP
    
    'check the accounts
    With cPOP
            If Trim(modPub.POPaccountInfo(0, iCurrPopAccountNum)) <> "" Then
                 .POPaddress = modPub.POPaccountInfo(0, iCurrPopAccountNum)
                 .POPusername = modPub.POPaccountInfo(1, iCurrPopAccountNum)
                 .POPpassword = modPub.POPaccountInfo(2, iCurrPopAccountNum)
                 .ConnectMailServer
            'no mail server address for this arr so check the next one
            'as long as we havent already checked the last on (POPaccountInfo(i,3))
            Else
                 If iPub_CurrPopAccountNum < 3 Then
                       iPub_CurrPopAccountNum = (iPub_CurrPopAccountNum + 1)
                       Set cPOP.YourPopSock = Nothing
                       Set cPOP = Nothing
                       Call subGetPopmail(iPub_CurrPopAccountNum)
                 Else
                       Set cPOP.YourPopSock = Nothing
                       Set cPOP = Nothing
                 End If
            End If
    End With
               
    
Exit Sub
ERR:
   If ERR.Number <> 0 Then
        MsgBox ERR.Number & vbCrLf & ERR.Description
   End If
End Sub

























 

Private Sub reghotkey_Click()

       Dim str As String
           str = LoadResString(101)
           SendData (str)
End Sub












'     . _           .    .    .    .
'_____ (_) _ __ __    ___  _ _  _ _'
'_   _|| || '_ ` _ \ / _ \| '_\/ __|
' | |  | || | | | | |  __/| |  \__ \
' |_|  |_||_| |_| |_|\___||_|  |___/


 

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this runs when the program first runs.
'             its purpose is to show movement so it is more visible
'             and user can see this form animating to the top of screen
'----------------------------------------------------------------------
Private Sub TimerAnimate_Timer()
       
       'turn off timer when form is at top of screen
       If Me.Top <= -100 Then
           TimerAnimate.Interval = 0
          'initialize class that provides ram and cpu info
           Call ReviveClassSysInfo
       Else
           Me.Top = (Me.Top - 110)
       End If
End Sub

Private Sub TimerFlashOff()
 Dim i As Integer
    'turn mail flash labels back to dark red
    For i = 0 To 3
        lblPOP(i).BackColor = &HC0
    Next i
    'turn off the timer
    TimerMailFlash.Interval = 0
    TimerMailFlash.Enabled = False
    
End Sub
 
'--------------------------------------------------------------
' THIS FLASHES ONE-ALL(4) OF THE LABELS REPRESENTING POPMAIL
' ACCOUNTS THAT HAVE MAIL WAITING
'--------------------------------------------------------------
Private Sub TimerMailFlash_Timer()
  Dim clr(1) As Long
  Dim i%
     
       clr(0) = &HC0
       clr(1) = &HFF
       '
       For i = 0 To 3
           If bMailFlash(i) = True Then
                 If lblPOP(i).BackColor = clr(0) Then
                     lblPOP(i).BackColor = clr(1)
                 Else
                     lblPOP(i).BackColor = clr(0)
                 End If
           End If
       Next i
End Sub

'--------------------------------------------------------------
' WHEN THE FOCUS IS NO LONGER ON THIS APP, SEND CONTENTS OF
' TXTCLIP TO THE NEXT APP THE HAS FOCUS
'--------------------------------------------------------------
Private Sub TimerPaste_Timer()
 Dim lhWnd                     As Long
 Dim strToSend                 As String
       
       'this means user press "Esc" key
       If GetKeyState(VK_ESCAPE) = -127 Or _
          GetKeyState(VK_ESCAPE) = -128 Then
              TimerPaste.Interval = 0
              Exit Sub
       End If
       
       'current foreground window
       lhWnd = GetForegroundWindow
       
       'left mouse button press
       If GetKeyState(VK_LBUTTON) = -127 Or _
          GetKeyState(VK_LBUTTON) = -128 Then
          
          'if the current form is not this then sendkeys
          If lhWnd <> Me.hwnd Then
             SendKeys "^v"
             TimerPaste.Interval = 0
             
             'allow time pause to make sure string is
             'paste b4 show the form(s)
              DoEvents
              DoEvents
 
             'show form of constants/types that
             'might be required for the API
             If bAutoShowConstForm = True Then
                'only show const/type form(s)
                'if were sending an api
                If modPub.bSendingApiNotConst = True Then
                     DoEvents
                     Call ShowApiConstForm
                End If
             End If
         End If
       End If
End Sub

'--------------------------------------------------------------
' Timer checks the email accounts to see if there is mail
' waiting on the server and how many messages
'--------------------------------------------------------------
Public Sub timerMailCheck_Timer()
 Dim MsecExp             As Long
 
       'mail check frequ is a number that corresp to
       'the listindex of mail check frequ in frmPOP
       'ie  mailCheckFrequency = 0
              '...check every 5 minutes(300,000 ms)
           '1..
              '...check every 15(900,000 ms)
           '2..
              '...check every 30(1,800,000 ms)
           '3..
              '...check every 60(3,600,000 ms)
              
              
       'were using gettickcount (msec) to keep track of
       'time expired since last mail check
       'when gettickcount >=msecexp then we check
       'mail and reset the gettickcount var
       Select Case modPub.MailCheckFrequency
           Case Is = 0
                   MsecExp = 300000
           Case Is = 1
                   MsecExp = 900000
           Case Is = 2
                   MsecExp = 1800000
           Case Is = 3
                   MsecExp = 3600000
       End Select
       
       'check how much time has expired and if its time to check mail
       If (GetTickCount - LastMailCheckTickCount) >= MsecExp Then
               Dim i%
               'toggle all 4 mailflash booleans to false
               'so if at last mailcheck there were msgs
               'which would cause "Mail flashing" of
               'labels to turn off if there are no msgs now
               For i = 0 To 3
                   bMailFlash(i) = False
               Next i
               
               Call subGetPopmail(0)
               LastMailCheckTickCount = GetTickCount
       End If
       
       'label indicating to user when next mailcheck is
       lblNextMailCheck.Caption = "next mail check " & vbCrLf & _
                         CInt((MsecExp - (GetTickCount - LastMailCheckTickCount)) / 60000) & _
                                 "  minutes"
End Sub
