Attribute VB_Name = "modConstants"
Option Explicit



'     . _ __  _   '    .    .      .    .     .     .      .     .    .
' __ _ | '_ \(_)  ' ___  ___  _ __   _ _'_____  __ _  _ __  _____  _ _'
'/ _` || |_) | |  '/ __|/ _ \| '_ ` / __|_   _|/ _` || '_ ` _   _|/ __|
' (_| || .__/| |  ' (__  (_) | | | |\__ \ | |   (_| || | | | | |  \__ \
'\__,_||_|   |_|  '\___|\___/|_| |_||___/ |_|  \__,_||_| |_| |_|  |___/


'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the CB (sendmessage) constants in
'             frmConstants
'----------------------------------------------------------------------
Sub CBmessage(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "CB_ADDSTRING As Long = &H143"
.AddItem "CB_GETCURSEL As Long = &H147"
.AddItem "CB_ERRSPACE As Long = (-2)"
.AddItem "CB_DELETESTRING As Long = &H144"
.AddItem "CB_DIR As Long = &H145"
.AddItem "CB_ERR As Long = (-1)"
.AddItem "CB_FINDSTRING As Long = &H14C"
.AddItem "CB_FINDSTRINGEXACT As Long = &H158"
.AddItem "CB_GETCOMBOBOXINFO As Long = &H164"
.AddItem "CB_GETCOUNT As Long = &H146"
.AddItem "CB_GETDROPPEDCONTROLRECT As Long = &H152"
.AddItem "CB_GETDROPPEDSTATE As Long = &H157"
.AddItem "CB_GETDROPPEDWIDTH As Long = &H15F"
.AddItem "CB_GETEDITSEL As Long = &H140"
.AddItem "CB_GETEXTENDEDUI As Long = &H156"
.AddItem "CB_GETHORIZONTALEXTENT As Long = &H15D"
.AddItem "CB_GETITEMDATA As Long = &H150"
.AddItem "CB_GETLBTEXT As Long = &H148"
.AddItem "CB_GETITEMHEIGHT As Long = &H154"
.AddItem "CB_GETLBTEXTLEN As Long = &H149"
.AddItem "CB_GETLOCALE As Long = &H15A"
.AddItem "CB_GETTOPINDEX As Long = &H15B"
.AddItem "CB_INITSTORAGE As Long = &H161"
.AddItem "CB_INSERTSTRING As Long = &H14A"
.AddItem "CB_LIMITTEXT As Long = &H141"
.AddItem "CB_MAX_CAB_PATH As Long = 256"
.AddItem "CB_MAX_CABINET_NAME As Long = 256"
.AddItem "CB_MAX_CHUNK As Long = 32768"
.AddItem "CB_MAX_DISK As Long = &H7FFFFFFF"
.AddItem "CB_MAX_FILENAME As Long = 256"
.AddItem "CB_MSGMAX As Long = &H15B"
.AddItem "CB_MULTIPLEADDSTRING As Long = &H163"
.AddItem "CB_OKAY As Long = 0"
.AddItem "CB_RESETCONTENT As Long = &H14B"
.AddItem "CB_SELECTSTRING As Long = &H14D"
.AddItem "CB_SETCURSEL As Long = &H14E"
.AddItem "CB_SETDROPPEDWIDTH As Long = &H160"
.AddItem "CB_SETEDITSEL As Long = &H142"
.AddItem "CB_SETHORIZONTALEXTENT As Long = &H15E"
.AddItem "CB_SETEXTENDEDUI As Long = &H155"
.AddItem "CB_SETITEMDATA As Long = &H151"
.AddItem "CB_SETITEMHEIGHT As Long = &H153"
.AddItem "CB_SETLOCALE As Long = &H159"
.AddItem "CB_SETTOPINDEX As Long = &H15C"
.AddItem "CB_SHOWDROPDOWN As Long = &H14F"
           End With
End Sub


'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the BM (sendmessage) constants in
'             frmConstants
'----------------------------------------------------------------------
Sub BMmessage(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
           
           End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the EM (sendmessage) constants in
'             frmConstants
'----------------------------------------------------------------------
Sub EMmessage(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "EM_AUTOURLDETECT As Long = (WM_USER + 91)"
.AddItem "EM_CANPASTE As Long = (WM_USER + 50)"
.AddItem "EM_CANREDO As Long = (WM_USER + 85)"
.AddItem "EM_CANUNDO As Long = &HC6"
.AddItem "EM_CHARFROMPOS As Long = &HD7"
.AddItem "EM_CONVPOSITION As Long = (WM_USER + 108)"
.AddItem "EM_DISPLAYBAND As Long = (WM_USER + 51)"
.AddItem "EM_EMPTYUNDOBUFFER As Long = &HCD"
.AddItem "EM_EXGETSEL As Long = (WM_USER + 52)"
.AddItem "EM_EXLIMITTEXT As Long = (WM_USER + 53)"
.AddItem "EM_EXLINEFROMCHAR As Long = (WM_USER + 54)"
.AddItem "EM_EXSETSEL As Long = (WM_USER + 55)"
.AddItem "EM_FINDTEXT As Long = (WM_USER + 56)"
.AddItem "EM_FINDTEXTEX As Long = (WM_USER + 79)"
.AddItem "EM_FINDTEXTEXW As Long = (WM_USER + 124)"
.AddItem "EM_FINDTEXTW As Long = (WM_USER + 123)"
.AddItem "EM_FINDWORDBREAK As Long = (WM_USER + 76)"
.AddItem "EM_FMTLINES As Long = &HC8"
.AddItem "EM_FORMATRANGE As Long = (WM_USER + 57)"
.AddItem "EM_GETAUTOURLDETECT As Long = (WM_USER + 92)"
.AddItem "EM_GETBIDIOPTIONS As Long = (WM_USER + 201)"
.AddItem "EM_GETCHARFORMAT As Long = (WM_USER + 58)"
.AddItem "EM_GETEDITSTYLE As Long = (WM_USER + 205)"
.AddItem "EM_GETEVENTMASK As Long = (WM_USER + 59)"
.AddItem "EM_GETFIRSTVISIBLELINE As Long = &HCE"
.AddItem "EM_GETHANDLE As Long = &HBD"
.AddItem "EM_GETIMECOLOR As Long = (WM_USER + 105)"
.AddItem "EM_GETIMECOMPMODE As Long = (WM_USER + 122)"
.AddItem "EM_GETIMEMODEBIAS As Long = (WM_USER + 127)"
.AddItem "EM_GETIMEOPTIONS As Long = (WM_USER + 107)"
.AddItem "EM_GETIMESTATUS As Long = &HD9"
.AddItem "EM_GETLANGOPTIONS As Long = (WM_USER + 121)"
.AddItem "EM_GETLIMITTEXT As Long = (WM_USER + 37)"
.AddItem "EM_GETLINE As Long = &HC4"
.AddItem "EM_GETLINECOUNT As Long = &HBA"
.AddItem "EM_GETMARGINS As Long = &HD4"
.AddItem "EM_GETOLEINTERFACE As Long = (WM_USER + 60)"
.AddItem "EM_GETMODIFY As Long = &HB8"
.AddItem "EM_GETOPTIONS As Long = (WM_USER + 78)"
.AddItem "EM_GETPARAFORMAT As Long = (WM_USER + 61)"
.AddItem "EM_GETPASSWORDCHAR As Long = &HD2"
.AddItem "EM_GETPUNCTUATION As Long = (WM_USER + 101)"
.AddItem "EM_GETRECT As Long = &HB2"
.AddItem "EM_GETREDONAME As Long = (WM_USER + 87)"
.AddItem "EM_GETSCROLLPOS As Long = (WM_USER + 221)"
.AddItem "EM_GETSEL As Long = &HB0"
.AddItem "EM_GETSELTEXT As Long = (WM_USER + 62)"
.AddItem "EM_GETTEXTEX As Long = (WM_USER + 94)"
.AddItem "EM_GETTEXTLENGTHEX As Long = (WM_USER + 95)"
.AddItem "EM_GETTEXTMODE As Long = (WM_USER + 90)"
.AddItem "EM_GETTEXTRANGE As Long = (WM_USER + 75)"
.AddItem "EM_GETTHUMB As Long = &HBE"
.AddItem "EM_GETTYPOGRAPHYOPTIONS As Long = (WM_USER + 203)"
.AddItem "EM_GETUNDONAME As Long = (WM_USER + 86)"
.AddItem "EM_GETWORDBREAKPROC As Long = &HD1"
.AddItem "EM_GETWORDBREAKPROCEX As Long = (WM_USER + 80)"
.AddItem "EM_GETWORDWRAPMODE As Long = (WM_USER + 103)"
.AddItem "EM_GETZOOM As Long = (WM_USER + 224)"
.AddItem "EM_HIDESELECTION As Long = (WM_USER + 63)"
.AddItem "EM_LIMITTEXT As Long = &HC5"
.AddItem "EM_LINEFROMCHAR As Long = &HC9"
.AddItem "EM_LINEINDEX As Long = &HBB"
.AddItem "EM_LINELENGTH As Long = &HC1"
.AddItem "EM_LINESCROLL As Long = &HB6"
.AddItem "EM_OUTLINE As Long = (WM_USER + 220)"
.AddItem "EM_PASTESPECIAL As Long = (WM_USER + 64)"
.AddItem "EM_POSFROMCHAR As Long = (WM_USER + 38)"
.AddItem "EM_RECONVERSION As Long = (WM_USER + 125)"
.AddItem "EM_REDO As Long = (WM_USER + 84)"
.AddItem "EM_REPLACESEL As Long = &HC2"
.AddItem "EM_REQUESTRESIZE As Long = (WM_USER + 65)"
.AddItem "EM_SCROLL As Long = &HB5"
.AddItem "EM_SCROLLCARET As Long = &HB7"
.AddItem "EM_SELECTIONTYPE As Long = (WM_USER + 66)"
.AddItem "EM_SETBIDIOPTIONS As Long = (WM_USER + 200)"
.AddItem "EM_SETBKGNDCOLOR As Long = (WM_USER + 67)"
.AddItem "EM_SETCHARFORMAT As Long = (WM_USER + 68)"
.AddItem "EM_SETCUEBANNER As Long = (ECM_FIRST + 1)"
.AddItem "EM_SETEDITSTYLE As Long = (WM_USER + 204)"
.AddItem "EM_SETEVENTMASK As Long = (WM_USER + 69)"
.AddItem "EM_SETFONTSIZE As Long = (WM_USER + 223)"
.AddItem "EM_SETHANDLE As Long = &HBC"
.AddItem "EM_SETIMECOLOR As Long = (WM_USER + 104)"
.AddItem "EM_SETIMEMODEBIAS As Long = (WM_USER + 126)"
.AddItem "EM_SETIMEOPTIONS As Long = (WM_USER + 106)"
.AddItem "EM_SETIMESTATUS As Long = &HD8"
.AddItem "EM_SETLANGOPTIONS As Long = (WM_USER + 120)"
.AddItem "EM_SETLIMITTEXT As Long = EM_LIMITTEXT"
.AddItem "EM_SETMARGINS As Long = &HD3"
.AddItem "EM_SETMODIFY As Long = &HB9"
.AddItem "EM_SETOLECALLBACK As Long = (WM_USER + 70)"
.AddItem "EM_SETOPTIONS As Long = (WM_USER + 77)"
.AddItem "EM_SETPALETTE As Long = (WM_USER + 93)"
.AddItem "EM_SETPARAFORMAT As Long = (WM_USER + 71)"
.AddItem "EM_SETPASSWORDCHAR As Long = &HCC"
.AddItem "EM_SETPUNCTUATION As Long = (WM_USER + 100)"
.AddItem "EM_SETREADONLY As Long = &HCF"
.AddItem "EM_SETRECT As Long = &HB3"
.AddItem "EM_SETRECTNP As Long = &HB4"
.AddItem "EM_SETSCROLLPOS As Long = (WM_USER + 222)"
.AddItem "EM_SETSEL As Long = &HB1"
.AddItem "EM_SETTABSTOPS As Long = &HCB"
.AddItem "EM_SETTARGETDEVICE As Long = (WM_USER + 72)"
.AddItem "EM_SETTEXTEX As Long = (WM_USER + 97)"
.AddItem "EM_SETTEXTMODE As Long = (WM_USER + 89)"
.AddItem "EM_SETTYPOGRAPHYOPTIONS As Long = (WM_USER + 202)"
.AddItem "EM_SETUNDOLIMIT As Long = (WM_USER + 82)"
.AddItem "EM_SETWORDBREAKPROCEX As Long = (WM_USER + 81)"
.AddItem "EM_SETWORDWRAPMODE As Long = (WM_USER + 102)"
.AddItem "EM_SHOWSCROLLBAR As Long = (WM_USER + 96)"
.AddItem "EM_STOPGROUPTYPING As Long = (WM_USER + 88)"
.AddItem "EM_STREAMIN As Long = (WM_USER + 73)"
.AddItem "EM_STREAMOUT As Long = (WM_USER + 74)"
.AddItem "EM_UNDO As Long = &HC7"

           End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the LB (sendmessage) constants in
'             frmConstants
'----------------------------------------------------------------------
Sub LBmessage(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "LB_ADDFILE As Long = &H196"
.AddItem "LB_ADDSTRING As Long = &H180"
.AddItem "LB_CTLCODE As Long = 0&"
.AddItem "LB_DELETESTRING As Long = &H182"
.AddItem "LB_DIR As Long = &H18D"
.AddItem "LB_DST_ADDR_USE_DSTADDR_FLAG As Long = &H8"
.AddItem "LB_DST_ADDR_USE_SRCADDR_FLAG As Long = &H4"
.AddItem "LB_DST_MASK_LATE_FLAG As Long = &H20"
.AddItem "LB_ERR As Long = (-1)"
.AddItem "LB_FINDSTRING As Long = &H18F"
.AddItem "LB_FINDSTRINGEXACT As Long = &H1A2"
.AddItem "LB_GETANCHORINDEX As Long = &H19D"
.AddItem "LB_GETCARETINDEX As Long = &H19F"
.AddItem "LB_GETCOUNT As Long = &H18B"
.AddItem "LB_GETCURSEL As Long = &H188"
.AddItem "LB_GETHORIZONTALEXTENT As Long = &H193"
.AddItem "LB_GETITEMDATA As Long = &H199"
.AddItem "LB_GETITEMHEIGHT As Long = &H1A1"
.AddItem "LB_GETLOCALE As Long = &H1A6"
.AddItem "LB_GETITEMRECT As Long = &H198"
.AddItem "LB_GETSEL As Long = &H187"
.AddItem "LB_GETSELCOUNT As Long = &H190"
.AddItem "LB_GETSELITEMS As Long = &H191"
.AddItem "LB_GETTEXT As Long = &H189"
.AddItem "LB_GETTEXTLEN As Long = &H18A"
.AddItem "LB_GETTOPINDEX As Long = &H18E"
.AddItem "LB_INITSTORAGE As Long = &H1A8"
.AddItem "LB_ITEMFROMPOINT As Long = &H1A9"
.AddItem "LB_MSGMAX As Long = &H1A8"
.AddItem "LB_MULTIPLEADDSTRING As Long = &H1B1"
.AddItem "LB_OKAY As Long = 0"
.AddItem "LB_RESETCONTENT As Long = &H184"
.AddItem "LB_SELECTSTRING As Long = &H18C"
.AddItem "LB_SELITEMRANGE As Long = &H19B"
.AddItem "LB_SELITEMRANGEEX As Long = &H183"
.AddItem "LB_SETANCHORINDEX As Long = &H19C"
.AddItem "LB_SETCARETINDEX As Long = &H19E"
.AddItem "LB_SETCOLUMNWIDTH As Long = &H195"
.AddItem "LB_SETCOUNT As Long = &H1A7"
.AddItem "LB_SETCURSEL As Long = &H186"
.AddItem "LB_SETITEMDATA As Long = &H19A"
.AddItem "LB_SETITEMHEIGHT As Long = &H1A0"
.AddItem "LB_SETLOCALE As Long = &H1A5"
.AddItem "LB_SETSEL As Long = &H185"
.AddItem "LB_SETTABSTOPS As Long = &H192"
.AddItem "LB_SETTOPINDEX As Long = &H197"

           End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the WM (sendmessage) constants in
'             frmConstants
'----------------------------------------------------------------------
Sub WMmessage(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "WM_ACTIVATE As Long = &H6"
.AddItem "WM_ACTIVATEAPP As Long = &H1C"
.AddItem "WM_AFXFIRST As Long = &H360"
.AddItem "WM_AFXLAST As Long = &H37F"
.AddItem "WM_APP As Long = &H8000"
.AddItem "WM_APPCOMMAND As Long = &H319"
.AddItem "WM_ASKCBFORMATNAME As Long = &H30C"
.AddItem "WM_CANCELJOURNAL As Long = &H4B"
.AddItem "WM_CANCELMODE As Long = &H1F"
.AddItem "WM_CHAR As Long = &H102"
.AddItem "WM_CHARTOITEM As Long = &H2F"
.AddItem "WM_CHILDACTIVATE As Long = &H22"
.AddItem "WM_CLEAR As Long = &H303"
.AddItem "WM_CLOSE As Long = &H10"
.AddItem "WM_COMMAND As Long = &H111"
.AddItem "WM_COMMNOTIFY As Long = &H44"
.AddItem "WM_COMPACTING As Long = &H41"
.AddItem "WM_COMPAREITEM As Long = &H39"
.AddItem "WM_CONVERTREQUEST As Long = &H10A"
.AddItem "WM_CONVERTREQUESTEX As Long = &H108"
.AddItem "WM_CONVERTRESULT As Long = &H10B"
.AddItem "WM_COPY As Long = &H301"
.AddItem "WM_CREATE As Long = &H1"
.AddItem "WM_CTLCOLOR As Long = &H19"
.AddItem "WM_CTLCOLOREDIT As Long = &H133"
.AddItem "WM_CTLCOLORSCROLLBAR As Long = &H137"
.AddItem "WM_CUT As Long = &H300"
.AddItem "WM_DESTROY As Long = &H2"
.AddItem "WM_DELETEITEM As Long = &H2D"
.AddItem "WM_DRAWCLIPBOARD As Long = &H308"
.AddItem "WM_DRAWITEM As Long = &H2B"
.AddItem "WM_ENABLE As Long = &HA"
.AddItem "WM_DROPFILES As Long = &H233"
.AddItem "WM_GETFONT As Long = &H31"
.AddItem "WM_GETHOTKEY As Long = &H33"
.AddItem "WM_GETICON As Long = &H7F"
.AddItem "WM_GETTEXT As Long = &HD"
.AddItem "WM_HELP As Long = &H53"
.AddItem "WM_HOTKEY As Long = &H312"
.AddItem "WM_HSCROLL As Long = &H114"
.AddItem "WM_LBUTTONDBLCLK As Long = &H203"
.AddItem "WM_KILLFOCUS As Long = &H8"
.AddItem "WM_KEYUP As Long = &H101"
.AddItem "WM_KEYLAST As Long = &H108"
.AddItem "WM_KEYFIRST As Long = &H100"
.AddItem "WM_KEYDOWN As Long = &H100"
.AddItem "WM_LBUTTONDOWN As Long = &H201"
.AddItem "WM_LBUTTONUP As Long = &H202"
.AddItem "WM_MBUTTONDBLCLK As Long = &H209"
.AddItem "WM_MBUTTONDOWN As Long = &H207"
.AddItem "WM_MBUTTONUP As Long = &H208"
.AddItem "WM_MENUCHAR As Long = &H120"
.AddItem "WM_MENUCOMMAND As Long = &H126"
.AddItem "WM_MENUDRAG As Long = &H123"
.AddItem "WM_MENUGETOBJECT As Long = &H124"
.AddItem "WM_MENURBUTTONUP As Long = &H122"
.AddItem "WM_MENUSELECT As Long = &H11F"
.AddItem "WM_MOUSEACTIVATE As Long = &H21"
.AddItem "WM_MOUSEFIRST As Long = &H200"
.AddItem "WM_MOUSEHOVER As Long = &H2A1"
.AddItem "WM_MOUSELAST As Long = &H209"
.AddItem "WM_MOUSELEAVE As Long = &H2A3"
.AddItem "WM_MOUSEMOVE As Long = &H200"
.AddItem "WM_MOVING As Long = &H216"
.AddItem "WM_NCACTIVATE As Long = &H86"
.AddItem "WM_NCCALCSIZE As Long = &H83"
.AddItem "WM_NCCREATE As Long = &H81"
.AddItem "WM_NCDESTROY As Long = &H82"
.AddItem "WM_NCLBUTTONDBLCLK As Long = &HA3"
.AddItem "WM_NCLBUTTONDOWN As Long = &HA1"
.AddItem "WM_NCLBUTTONUP As Long = &HA2"
.AddItem "WM_NCMBUTTONDBLCLK As Long = &HA9"
.AddItem "WM_NCMBUTTONDOWN As Long = &HA7"
.AddItem "WM_NCMBUTTONUP As Long = &HA8"
.AddItem "WM_NCMOUSEHOVER As Long = &H2A0"
.AddItem "WM_NCMOUSELEAVE As Long = &H2A2"
.AddItem "WM_NCMOUSEMOVE As Long = &HA0"
.AddItem "WM_NCPAINT As Long = &H85"
.AddItem "WM_NCRBUTTONDBLCLK As Long = &HA6"
.AddItem "WM_NCRBUTTONDOWN As Long = &HA4"
.AddItem "WM_NCRBUTTONUP As Long = &HA5"
.AddItem "WM_NEXTMENU As Long = &H213"
.AddItem "WM_NOTIFY As Long = &H4E"
.AddItem "WM_PAINT As Long = &HF&"
.AddItem "WM_PAINTCLIPBOARD As Long = &H309"
.AddItem "WM_PAINTICON As Long = &H26"
.AddItem "WM_PASTE As Long = &H302"
.AddItem "WM_PRINT As Long = &H317"
.AddItem "WM_QUIT As Long = &H12"
.AddItem "WM_RBUTTONDBLCLK As Long = &H206"
.AddItem "WM_RBUTTONDOWN As Long = &H204"
.AddItem "WM_RBUTTONUP As Long = &H205"
.AddItem "WM_SETCURSOR As Long = &H20"
.AddItem "WM_SETFOCUS As Long = &H7"
.AddItem "WM_SETFONT As Long = &H30"
.AddItem "WM_SETHOTKEY As Long = &H32"
.AddItem "WM_SETICON As Long = &H80"
.AddItem "WM_SETREDRAW As Long = &HB"
.AddItem "WM_SETTEXT As Long = &HC"
.AddItem "WM_SHOWWINDOW As Long = &H18"
.AddItem "WM_SIZE As Long = &H5"
.AddItem "WM_SIZECLIPBOARD As Long = &H30B"
.AddItem "WM_SIZING As Long = &H214"
.AddItem "WM_SYSCOMMAND As Long = &H112"
.AddItem "WM_TIMER As Long = &H113"
.AddItem "WM_TIMECHANGE As Long = &H1E"
.AddItem "WM_UNDO As Long = &H304"
.AddItem "WM_USER As Long = &H400"
.AddItem "WM_USERCHANGED As Long = &H54"
.AddItem "WM_VSCROLL As Long = &H115"
.AddItem "WM_WINDOWPOSCHANGED As Long = &H47"
.AddItem "WM_WINDOWPOSCHANGING As Long = &H46"
.AddItem "WM_WININICHANGE As Long = &H1A"

           End With
End Sub






'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the AnimateWindow constants (cursor)
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub AW(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "AW_ACTIVATE As Long = &H20000"
.AddItem "AW_BLEND As Long = &H80000"
.AddItem "AW_CENTER As Long = &H10"
.AddItem "AW_HIDE As Long = &H10000"
.AddItem "AW_HOR_NEGATIVE As Long = &H2"
.AddItem "AW_HOR_POSITIVE As Long = &H1"
.AddItem "AW_SLIDE As Long = &H40000"
.AddItem "AW_VER_NEGATIVE As Long = &H8"
.AddItem "AW_VER_POSITIVE As Long = &H4"
         End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the Mouse_Event constants (cursor)
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub Mouse_Event(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "MOUSEEVENTF_ABSOLUTE As Long = &H8000"
.AddItem "MOUSEEVENTF_LEFTDOWN As Long = &H2"
.AddItem "MOUSEEVENTF_LEFTUP As Long = &H4"
.AddItem "MOUSEEVENTF_MIDDLEDOWN As Long = &H20"
.AddItem "MOUSEEVENTF_MIDDLEUP As Long = &H40"
.AddItem "MOUSEEVENTF_MOVE As Long = &H1"
.AddItem "MOUSEEVENTF_RIGHTDOWN As Long = &H8"
.AddItem "MOUSEEVENTF_RIGHTUP As Long = &H10"
.AddItem "MOUSEEVENTF_VIRTUALDESK As Long = &H4000"
.AddItem "MOUSEEVENTF_WHEEL As Long = &H800"
.AddItem "MOUSEEVENTF_XDOWN As Long = &H80"
.AddItem "MOUSEEVENTF_XUP As Long = &H100"
          End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the IDC constants (cursor)
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub OCR(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "OCR_APPSTARTING As Long = 32650"
.AddItem "OCR_CROSS As Long = 32515"
.AddItem "OCR_HAND As Long = 32649"
.AddItem "OCR_IBEAM As Long = 32513"
.AddItem "OCR_ICOCUR As Long = 32647"
.AddItem "OCR_ICON As Long = 32641"
.AddItem "OCR_NO As Long = 32648"
.AddItem "OCR_NORMAL As Long = 32512"
.AddItem "OCR_SIZE As Long = 32640"
.AddItem "OCR_SIZEALL As Long = 32646"
.AddItem "OCR_SIZENESW As Long = 32643"
.AddItem "OCR_SIZENS As Long = 32645"
.AddItem "OCR_SIZENWSE As Long = 32642"
.AddItem "OCR_SIZEWE As Long = 32644"
.AddItem "OCR_UP As Long = 32516"
.AddItem "OCR_WAIT As Long = 32514"

             End With
End Sub


'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the SetStretchBltMode constants (cursor)
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub SetStretchBltmode(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "BLACKONWHITE As Long = 1"
.AddItem "COLORONCOLOR As Long = 3"
.AddItem "HALFTONE As Long = 4"
.AddItem "WHITEONBLACK As Long = 2"

             End With
End Sub


'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the IDC constants (cursor)
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub PatBlt(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "PATCOPY = &HF00021"
.AddItem "PATINVERT = &H5A0049"
.AddItem "PATPAINT = &HFB0A09"
.AddItem "DSTINVERT = &H550009"
.AddItem "BLACKNESS = &H42"
.AddItem "WHITENESS = &HFF0062"
             End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the IDI constants (cursor)
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub IDI(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "IDI_APPLICATION As Long = 32512&"
.AddItem "IDI_ASTERISK As Long = 32516&"
.AddItem "IDI_CLASSICON_OVERLAYFIRST As Long = 500"
.AddItem "IDI_CLASSICON_OVERLAYLAST As Long = 502"
.AddItem "IDI_CONFLICT As Long = 161"
.AddItem "IDI_DISABLED_OVL As Long = 501"
.AddItem "IDI_EXCLAMATION As Long = 32515&"
.AddItem "IDI_FORCED_OVL As Long = 502"
.AddItem "IDI_HAND As Long = 32513&"
.AddItem "IDI_PROBLEM_OVL As Long = 500"
.AddItem "IDI_QUESTION As Long = 32514&"
.AddItem "IDI_RESOURCE As Long = 159"
.AddItem "IDI_RESOURCEFIRST As Long = 159"
.AddItem "IDI_RESOURCELAST As Long = 161"
.AddItem "IDI_RESOURCEOVERLAYFIRST As Long = 161"
.AddItem "IDI_RESOURCEOVERLAYLAST As Long = 161"
.AddItem "IDI_WINLOGO As Long = 32517"
               End With
End Sub
 
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the IDC constants (cursor)
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub IDC(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "IDC_APPSTARTING As Long = 32650&"
.AddItem "IDC_ARROW As Long = 32512&"
.AddItem "IDC_CROSS As Long = 32515&"
.AddItem "IDC_HAND As Long = 32649"
.AddItem "IDC_HELP As Long = 32651"
.AddItem "IDC_IBEAM As Long = 32513&"
.AddItem "IDC_ICON As Long = 32641&"
.AddItem "IDC_NO As Long = 32648&"
.AddItem "IDC_SIZE As Long = 32640&"
.AddItem "IDC_SIZEALL As Long = 32646&"
.AddItem "IDC_SIZENESW As Long = 32643&"
.AddItem "IDC_SIZENS As Long = 32645&"
.AddItem "IDC_SIZENWSE As Long = 32642&"
.AddItem "IDC_SIZEWE As Long = 32644&"
.AddItem "IDC_UPARROW As Long = 32516&"
.AddItem "IDC_WAIT As Long = 32514&"
                End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for SysColor and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub SysColor(ListName As ListBox)

 frmConstants.Visible = True
           With ListName
.AddItem "COLOR_3DDKSHADOW As Long = 21"
.AddItem "COLOR_3DLIGHT As Long = 22"
.AddItem "COLOR_ACTIVEBORDER As Long = 10"
.AddItem "COLOR_ACTIVECAPTION As Long = 2"
.AddItem "COLOR_ADD As Long = 712"
.AddItem "COLOR_ADJ_MAX As Long = 100"
.AddItem "COLOR_ADJ_MIN As Long = -100"
.AddItem "COLOR_APPWORKSPACE As Long = 12"
.AddItem "COLOR_BACKGROUND As Long = 1"
.AddItem "COLOR_BLUE As Long = 708"
.AddItem "COLOR_BLUEACCEL As Long = 728"
.AddItem "COLOR_BOX1 As Long = 720"
.AddItem "COLOR_BTNFACE As Long = 15"
.AddItem "COLOR_BTNHIGHLIGHT As Long = 20"
.AddItem "COLOR_BTNSHADOW As Long = 16"
.AddItem "COLOR_BTNTEXT As Long = 18"
.AddItem "COLOR_CAPTIONTEXT As Long = 9"
.AddItem "COLOR_CURRENT As Long = 709"
.AddItem "COLOR_CUSTOM1 As Long = 721"
.AddItem "COLOR_ELEMENT As Long = 716"
.AddItem "COLOR_GRAYTEXT As Long = 17"
.AddItem "COLOR_GREEN As Long = 707"
.AddItem "COLOR_GREENACCEL As Long = 727"
.AddItem "COLOR_HIGHLIGHT As Long = 13"
.AddItem "COLOR_HIGHLIGHTTEXT As Long = 14"
.AddItem "COLOR_HOTLIGHT As Long = 26"
.AddItem "COLOR_HUE As Long = 703"
.AddItem "COLOR_HUEACCEL As Long = 723"
.AddItem "COLOR_HUESCROLL As Long = 700"
.AddItem "COLOR_INACTIVEBORDER As Long = 11"
.AddItem "COLOR_INACTIVECAPTION As Long = 3"
.AddItem "COLOR_INACTIVECAPTIONTEXT As Long = 19"
.AddItem "COLOR_INFOBK As Long = 24"
.AddItem "COLOR_INFOTEXT As Long = 23"
.AddItem "COLOR_LUM As Long = 705"
.AddItem "COLOR_LUMACCEL As Long = 725"
.AddItem "COLOR_LUMSCROLL As Long = 702"
.AddItem "COLOR_MATCH_VERSION As Long = &H200"
.AddItem "COLOR_MENU As Long = 4"
.AddItem "COLOR_MENUTEXT As Long = 7"
.AddItem "COLOR_MIX As Long = 719"
.AddItem "COLOR_NO_TRANSPARENT As Long = &HFFFFFFFF"
.AddItem "COLOR_PALETTE As Long = 718"
.AddItem "COLOR_RAINBOW As Long = 710"
.AddItem "COLOR_RED As Long = 706"
.AddItem "COLOR_REDACCEL As Long = 726"
.AddItem "COLOR_SAMPLES As Long = 717"
.AddItem "COLOR_SAT As Long = 704"
.AddItem "COLOR_SATACCEL As Long = 724"
.AddItem "COLOR_SATSCROLL As Long = 701"
.AddItem "COLOR_SAVE As Long = 711"
.AddItem "COLOR_SCHEMES As Long = 715"
.AddItem "COLOR_SCROLLBAR As Long = 0"
.AddItem "COLOR_SOLID As Long = 713"
.AddItem "COLOR_SOLID_LEFT As Long = 730"
.AddItem "COLOR_SOLID_RIGHT As Long = 731"
.AddItem "COLOR_TUNE As Long = 714"
.AddItem "COLOR_WINDOW As Long = 5"
.AddItem "COLOR_WINDOWFRAME As Long = 6"
.AddItem "COLOR_WINDOWTEXT As Long = 8"
                  End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for GetDiBits api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub GetDiBits(ListName As ListBox)

 frmConstants.Visible = True
            With ListName
.AddItem "DIB_RGB_COLORS As Long = 0"
.AddItem "DIB_PAL_PHYSINDICES As Long = 2"
.AddItem "DIB_PAL_LOGINDICES As Long = 4"
.AddItem "DIB_PAL_INDICES As Long = 2"
.AddItem "DIB_PAL_COLORS As Long = 1"
                      End With
End Sub


'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for FloodFill api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub FloodFill(ListName As ListBox)

 frmConstants.Visible = True
            With ListName
.AddItem "FLOODFILLBORDER As Long = 0"
.AddItem "FLOODFILLSURFACE As Long = 1"
        End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for drawiconex api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub DI(ListName As ListBox)

 frmConstants.Visible = True
            With ListName
.AddItem "DI_MASK = &H1"
.AddItem "DI_IMAGE = &H2"
.AddItem "DI_NORMAL = &H3"
.AddItem "DI_COMPAT = &H4"
.AddItem "DI_DEFAULTSIZE = &H8"
            End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for drawstate api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub DSS_DST(ListName As ListBox)

 frmConstants.Visible = True
            With ListName
.AddItem "DSS_DISABLED As Long = &H20"
.AddItem "DSS_HIDEPREFIX As Long = &H200"
.AddItem "DSS_MONO As Long = &H80"
.AddItem "DSS_NORMAL As Long = &H0"
.AddItem "DSS_PREFIXONLY As Long = &H400"
.AddItem "DSS_RIGHT As Long = &H8000"
.AddItem "DSS_UNION As Long = &H10"
.AddItem "DST_BITMAP As Long = &H4"
.AddItem "DST_COMPLEX As Long = &H0"
.AddItem "DST_ICON As Long = &H3"
.AddItem "DST_PREFIXTEXT As Long = &H2"
.AddItem "DST_TEXT As Long = &H1"
            End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for drawFrameControl api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub DrawFrameControl(ListName As ListBox)

 frmConstants.Visible = True
            With ListName
            
.AddItem "DFC_BUTTON As Long = 4"
.AddItem "DFC_CAPTION As Long = 1"
.AddItem "DFC_POPUPMENU As Long = 5"
.AddItem "DFC_MENU As Long = 2"
.AddItem "DFC_SCROLL As Long = 3"
.AddItem "DFCS_ADJUSTRECT As Long = &H2000"
.AddItem "DFCS_BUTTON3STATE As Long = &H8"
.AddItem "DFCS_BUTTONCHECK As Long = &H0"
.AddItem "DFCS_BUTTONPUSH As Long = &H10"
.AddItem "DFCS_BUTTONRADIO As Long = &H4"
.AddItem "DFCS_BUTTONRADIOIMAGE As Long = &H1"
.AddItem "DFCS_BUTTONRADIOMASK As Long = &H2"
.AddItem "DFCS_CAPTIONCLOSE As Long = &H0"
.AddItem "DFCS_CAPTIONHELP As Long = &H4"
.AddItem "DFCS_CAPTIONMAX As Long = &H2"
.AddItem "DFCS_CAPTIONMIN As Long = &H1"
.AddItem "DFCS_CAPTIONRESTORE As Long = &H3"
.AddItem "DFCS_CHECKED As Long = &H400"
.AddItem "DFCS_FLAT As Long = &H4000"
.AddItem "DFCS_HOT As Long = &H1000"
.AddItem "DFCS_INACTIVE As Long = &H100"
.AddItem "DFCS_MENUARROW As Long = &H0"
.AddItem "DFCS_MENUARROWRIGHT As Long = &H4"
.AddItem "DFCS_MENUBULLET As Long = &H2"
.AddItem "DFCS_MENUCHECK As Long = &H1"
.AddItem "DFCS_MONO As Long = &H8000"
.AddItem "DFCS_PUSHED As Long = &H200"
.AddItem "DFCS_SCROLLCOMBOBOX As Long = &H5"
.AddItem "DFCS_SCROLLLEFT As Long = &H2"
.AddItem "DFCS_SCROLLRIGHT As Long = &H3"
.AddItem "DFCS_SCROLLSIZEGRIP As Long = &H8"
.AddItem "DFCS_SCROLLSIZEGRIPRIGHT As Long = &H10"
.AddItem "DFCS_SCROLLUP As Long = &H0"
.AddItem "DFCS_TRANSPARENT As Long = &H800"
            
            End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for drawedge api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub DrawEdge(ListName As ListBox)
 
 frmConstants.Visible = True
            With ListName
            
.AddItem "BDR_INNER As Long = &HC"
.AddItem "BDR_OUTER As Long = &H3"
.AddItem "BDR_RAISED As Long = &H5"
.AddItem "BDR_RAISEDINNER As Long = &H4"
.AddItem "BDR_RAISEDOUTER As Long = &H1"
.AddItem "BDR_SUNKEN As Long = &HA"
.AddItem "BDR_SUNKENINNER As Long = &H8"
.AddItem "BDR_SUNKENOUTER As Long = &H2"
.AddItem "EDGE_BUMP As Long = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)"
.AddItem "EDGE_ETCHED As Long = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)"
.AddItem "EDGE_RAISED As Long = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)"
.AddItem "EDGE_SUNKEN As Long = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)"
.AddItem "BF_ADJUST As Long = &H2000"
.AddItem "BF_BOTTOM As Long = &H8"
.AddItem "BF_BOTTOMLEFT As Long = (BF_BOTTOM Or BF_LEFT)"
.AddItem "BF_BOTTOMRIGHT As Long = (BF_BOTTOM Or BF_RIGHT)"
.AddItem "BF_DIAGONAL As Long = &H10"
.AddItem "BF_DIAGONAL_ENDBOTTOMRIGHT As Long = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)"
.AddItem "BF_DIAGONAL_ENDTOPLEFT As Long = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)"
.AddItem "BF_DIAGONAL_ENDTOPRIGHT As Long = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)"
.AddItem "BF_FLAT As Long = &H4000"
.AddItem "BF_LEFT As Long = &H1"
.AddItem "BF_MIDDLE As Long = &H800"
.AddItem "BF_MONO As Long = &H8000"
.AddItem "BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)"
.AddItem "BF_RIGHT As Long = &H4"
.AddItem "BF_SOFT As Long = &H1000"
.AddItem "BF_TOP As Long = &H2"
.AddItem "BF_TOPLEFT As Long = (BF_TOP Or BF_LEFT)"
.AddItem "BF_TOPRIGHT As Long = (BF_TOP Or BF_RIGHT)"
            
            End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for drawcaption api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub DC(ListName As ListBox)
 
 frmConstants.Visible = True
            With ListName
.AddItem "DC_ACTIVE As Long = &H1"
.AddItem "DC_BINADJUST As Long = 19"
.AddItem "DC_BINNAMES As Long = 12"
.AddItem "DC_BINS As Long = 6"
.AddItem "DC_BRUSH As Long = 18"
.AddItem "DC_COLLATE As Long = 22"
.AddItem "DC_COLORDEVICE As Long = 32"
.AddItem "DC_COPIES As Long = 18"
.AddItem "DC_DATATYPE_PRODUCED As Long = 21"
.AddItem "DC_DRIVER As Long = 11"
.AddItem "DC_DUPLEX As Long = 7"
.AddItem "DC_ENUMRESOLUTIONS As Long = 13"
.AddItem "DC_EXTRA As Long = 9"
.AddItem "DC_FIELDS As Long = 1"
.AddItem "DC_FILEDEPENDENCIES As Long = 14"
.AddItem "DC_GRADIENT As Long = &H20"
.AddItem "DC_HASDEFID As Long = &H534"
.AddItem "DC_ICON As Long = &H4"
.AddItem "DC_INBUTTON As Long = &H10"
.AddItem "DC_MANUFACTURER As Long = 23"
.AddItem "DC_MAXEXTENT As Long = 5"
.AddItem "DC_MEDIAREADY As Long = 29"
.AddItem "DC_MODEL As Long = 24"
.AddItem "DC_NUP As Long = 33"
.AddItem "DC_ORIENTATION As Long = 17"
.AddItem "DC_PAPERNAMES As Long = 16"
.AddItem "DC_PAPERS As Long = 2"
.AddItem "DC_PAPERSIZE As Long = 3"
.AddItem "DC_PEN As Long = 19"
.AddItem "DC_PRINTERMEM As Long = 28"
.AddItem "DC_PRINTRATE As Long = 26"
.AddItem "DC_PRINTRATEPPM As Long = 31"
.AddItem "DC_PRINTRATEUNIT As Long = 27"
.AddItem "DC_SIZE As Long = 8"
.AddItem "DC_SMALLCAP As Long = &H2"
.AddItem "DC_STAPLE As Long = 30"
.AddItem "DC_TEXT As Long = &H8"
.AddItem "DC_TRUETYPE As Long = 15"
.AddItem "DC_VERSION As Long = 10"

            End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for drawanimatedrects api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub IDANI(ListName As ListBox)
 
 frmConstants.Visible = True
            With ListName
            
.AddItem "IDANI_CAPTION As Long = &H3"
.AddItem "IDANI_CLOSE As Long = &H2"
.AddItem "IDANI_OPEN As Long = &H1"
      
             End With
End Sub

Sub SystemParametersInfo(ListName As ListBox)
 
 frmConstants.Visible = True
             With ListName
             
 .AddItem "SPI_GETACCESSTIMEOUT = 60"
 .AddItem "SPI_GETACTIVEWINDOWTRACKING = &H1000"
 .AddItem "SPI_GETACTIVEWNDTRKTIMEOUT = 8194"
 .AddItem "SPI_GETACTIVEWNDTRKZORDER = &H100C"
 .AddItem "SPI_GETANIMATION = 72"
 .AddItem "SPI_GETBEEP = 1"
 .AddItem "SPI_GETBLOCKSENDINPUTRESETS = &H1026"
 .AddItem "SPI_GETBORDER = 5"
 .AddItem "SPI_GETCARETWIDTH = &H2006"
 .AddItem "SPI_GETCOMBOBOXANIMATION = &H1004"
 .AddItem "SPI_GETCURSORSHADOW = &H101A"
 .AddItem "SPI_GETDEFAULTINPUTLANG = 89"
 .AddItem "SPI_GETDESKWALLPAPER = 115"
 .AddItem "SPI_GETDRAGFULLWINDOWS = 38"
 .AddItem "SPI_GETDROPSHADOW = &H1024"
 .AddItem "SPI_GETFASTTASKSWITCH = 35"
 .AddItem "SPI_GETFILTERKEYS = 50"
 .AddItem "SPI_GETFLATMENU = &H1022"
 .AddItem "SPI_GETFOCUSBORDERHEIGHT = &H2010"
 .AddItem "SPI_GETFOCUSBORDERWIDTH = &H200E"
 .AddItem "SPI_GETFONTSMOOTHING = 74"
 .AddItem "SPI_GETFONTSMOOTHINGCONTRAST = &H200C"
 .AddItem "SPI_GETFONTSMOOTHINGORIENTATION = &H2012"
 .AddItem "SPI_GETFONTSMOOTHINGTYPE = &H200A"
 .AddItem "SPI_GETFOREGROUNDFLASHCOUNT = &H2004"
 .AddItem "SPI_GETFOREGROUNDLOCKTIMEOUT = &H2000"
 .AddItem "SPI_GETGRADIENTCAPTIONS = &H1008"
 .AddItem "SPI_GETGRIDGRANULARITY=18"
 .AddItem "SPI_GETHIGHCONTRAST = 66"
 .AddItem "SPI_GETHOTTRACKING = &H100E"
 .AddItem "SPI_GETICONMETRICS = 45"
 .AddItem "SPI_GETICONTITLELOGFONT = 31"
 .AddItem "SPI_GETICONTITLEWRAP = 25"
 .AddItem "SPI_GETKEYBOARDCUES = &H100A"
 .AddItem "SPI_GETKEYBOARDDELAY = 22"
 .AddItem "SPI_GETKEYBOARDPREF = 68"
 .AddItem "SPI_GETKEYBOARDSPEED = 10"
 .AddItem "SPI_GETLISTBOXSMOOTHSCROLLING = &H1006"
 .AddItem "SPI_GETLOWPOWERACTIVE = 83"
 .AddItem "SPI_GETLOWPOWERTIMEOUT = 79"
 .AddItem "SPI_GETMENUANIMATION = &H1002"
 .AddItem "SPI_GETMENUDROPALIGNMENT = 27"
 .AddItem "SPI_GETMENUFADE = &H1012"
 .AddItem "SPI_GETMENUSHOWDELAY = 106"
 .AddItem "SPI_GETMINIMIZEDMETRICS = 43"
 .AddItem "SPI_GETMOUSE = 3"
 .AddItem "SPI_GETMOUSECLICKLOCK = &H101E"
 .AddItem "SPI_GETMOUSECLICKLOCKTIME = &H2008"
 .AddItem "SPI_GETMOUSEHOVERHEIGHT = 100"
 .AddItem "SPI_GETMOUSEHOVERTIME = 102"
 .AddItem "SPI_GETMOUSEHOVERWIDTH = 98"
 .AddItem "SPI_GETMOUSEKEYS = 54"
 .AddItem "SPI_GETMOUSESONAR = &H101C"
 .AddItem "SPI_GETMOUSESPEED = 112"
 .AddItem "SPI_GETMOUSETRAILS = 94"
 .AddItem "SPI_GETMOUSEVANISH = &H1020"
 .AddItem "SPI_GETNONCLIENTMETRICS = 41"
 .AddItem "SPI_GETPOWEROFFACTIVE = 84"
 .AddItem "SPI_GETPOWEROFFTIMEOUT = 80"
 .AddItem "SPI_GETSCREENREADER = 70"
 .AddItem "SPI_GETSCREENSAVEACTIVE = 16"
 .AddItem "SPI_GETSCREENSAVERRUNNING = 114"
 .AddItem "SPI_GETSCREENSAVETIMEOUT = 14"
 .AddItem "SPI_GETSELECTIONFADE = &H1014"
 .AddItem "SPI_GETSERIALKEYS = 62"
 .AddItem "SPI_GETSHOWIMEUI = 110"
 .AddItem "SPI_GETSHOWSOUNDS = 56"
 .AddItem "SPI_GETSNAPTODEFBUTTON = 95"
 .AddItem "SPI_GETSOUNDSENTRY = 64"
 .AddItem "SPI_GETSTICKYKEYS = 58"
 .AddItem "SPI_GETTOGGLEKEYS = 52"
 .AddItem "SPI_GETTOOLTIPANIMATION = &H1016"
 .AddItem "SPI_GETTOOLTIPFADE = &H1018"
 .AddItem "SPI_GETUIEFFECTS = &H103E"
 .AddItem "SPI_GETWHEELSCROLLLINES = 104"
 .AddItem "SPI_GETWINDOWSEXTENSION = 92"
 .AddItem "SPI_GETWORKAREA = 48"
 .AddItem "SPI_ICONHORIZONTALSPACING = 13"
 .AddItem "SPI_ICONVERTICALSPACING = 24"
 .AddItem "SPI_LANGDRIVER = 12"
 .AddItem "SPI_SCREENSAVERRUNNING = 97"
 .AddItem "SPI_SETACCESSTIMEOUT = 61"
 .AddItem "SPI_SETACTIVEWINDOWTRACKING = &H1001"
 .AddItem "SPI_SETACTIVEWNDTRKTIMEOUT = &H2003"
 .AddItem "SPI_SETACTIVEWNDTRKZORDER = &H100D"
 .AddItem "SPI_SETANIMATION = 73"
 .AddItem "SPI_SETBEEP = 2"
 .AddItem "SPI_SETBLOCKSENDINPUTRESETS = &H1027"
 .AddItem "SPI_SETBORDER = 6"
 .AddItem "SPI_SETCARETWIDTH = 8199"
 .AddItem "SPI_SETCOMBOBOXANIMATION = &H1005"
 .AddItem "SPI_SETCURSORS = 87"
 .AddItem "SPI_SETCURSORSHADOW = &H101B"
 .AddItem "SPI_SETDEFAULTINPUTLANG = 90"
 .AddItem "SPI_SETDESKPATTERN = 21"
 .AddItem "SPI_SETDESKWALLPAPER = 20"
 .AddItem "SPI_SETDOUBLECLICKTIME = 32"
 .AddItem "SPI_SETDOUBLECLKHEIGHT = 30"
 .AddItem "SPI_SETDOUBLECLKWIDTH = 29"
 .AddItem "SPI_SETDRAGFULLWINDOWS = 37"
 .AddItem "SPI_SETDRAGHEIGHT = 77"
 .AddItem "SPI_SETDRAGWIDTH = 76"
 .AddItem "SPI_SETDROPSHADOW = 4133"
 .AddItem "SPI_SETFASTTASKSWITCH = 36"
 .AddItem "SPI_SETFILTERKEYS = 51"
 .AddItem "SPI_SETFLATMENU = &H1023"
 .AddItem "SPI_SETFOCUSBORDERHEIGHT = &H2011"
 .AddItem "SPI_SETFOCUSBORDERWIDTH = &H200F"
 .AddItem "SPI_SETFONTSMOOTHING = 75"
 .AddItem "SPI_SETFONTSMOOTHINGCONTRAST = &H200D"
 .AddItem "SPI_SETFONTSMOOTHINGORIENTATION = &H2013"
 .AddItem "SPI_SETFONTSMOOTHINGTYPE = &H200B"
 .AddItem "SPI_SETFOREGROUNDFLASHCOUNT = &H2005"
 .AddItem "SPI_SETFOREGROUNDLOCKTIMEOUT = &H2001"
 .AddItem "SPI_SETGRADIENTCAPTIONS = &H1009"
 .AddItem "SPI_SETGRIDGRANULARITY = 19"
 .AddItem "SPI_SETHANDHELD = 78"
 .AddItem "SPI_SETHIGHCONTRAST = 67"
 .AddItem "SPI_SETHOTTRACKING = &H100F"
 .AddItem "SPI_SETICONMETRICS = 46"
 .AddItem "SPI_SETICONS = 88"
 .AddItem "SPI_SETICONTITLELOGFONT = 34"
 .AddItem "SPI_SETICONTITLEWRAP = 26"
 .AddItem "SPI_SETKEYBOARDCUES = &H100B"
 .AddItem "SPI_SETKEYBOARDDELAY = 23"
 .AddItem "SPI_SETKEYBOARDPREF = 69"
 .AddItem "SPI_SETKEYBOARDSPEED = 11"
 .AddItem "SPI_SETLANGTOGGLE = 91"
 .AddItem "SPI_SETLISTBOXSMOOTHSCROLLING = &H1007"
 .AddItem "SPI_SETLOWPOWERACTIVE = 85"
 .AddItem "SPI_SETLOWPOWERTIMEOUT = 81"
 .AddItem "SPI_SETMENUANIMATION = &H1003"
 .AddItem "SPI_SETMENUDROPALIGNMENT = 28"
 .AddItem "SPI_SETMENUFADE = &H1013"
 .AddItem "SPI_SETMENUSHOWDELAY = 107"
 .AddItem "SPI_SETMINIMIZEDMETRICS = 44"
 .AddItem "SPI_SETMOUSE = 4"
 .AddItem "SPI_SETMOUSEBUTTONSWAP = 33"
 .AddItem "SPI_SETMOUSECLICKLOCK = &H101F"
 .AddItem "SPI_SETMOUSECLICKLOCKTIME = &H2009"
 .AddItem "SPI_SETMOUSEHOVERHEIGHT = 101"
 .AddItem "SPI_SETMOUSEHOVERTIME = 103"
 .AddItem "SPI_SETMOUSEHOVERWIDTH = 99"
 .AddItem "SPI_SETMOUSEKEYS = 55"
 .AddItem "SPI_SETMOUSESONAR = &H101D"
 .AddItem "SPI_SETMOUSESPEED = 113"
 .AddItem "SPI_SETMOUSETRAILS=93"
 .AddItem "SPI_SETMOUSEVANISH = &H1021"
 .AddItem "SPI_SETNONCLIENTMETRICS = 42"
 .AddItem "SPI_SETPENWINDOWS = 49"
 .AddItem "SPI_SETPOWEROFFACTIVE = 86"
 .AddItem "SPI_SETPOWEROFFTIMEOUT = 82"
 .AddItem "SPI_SETSCREENREADER = 71"
 .AddItem "SPI_SETSCREENSAVEACTIVE = 17"
 .AddItem "SPI_SETSCREENSAVERRUNNING = 97"
 .AddItem "SPI_SETSCREENSAVETIMEOUT = 15"
 .AddItem "SPI_SETSELECTIONFADE = &H1015"
 .AddItem "SPI_SETSERIALKEYS = 63"
 .AddItem "SPI_SETSHOWIMEUI = 111"
 .AddItem "SPI_SETSHOWSOUNDS=57"
 .AddItem "SPI_SETSNAPTODEFBUTTON = 96"
 .AddItem "SPI_SETSOUNDSENTRY = 65"
 .AddItem "SPI_SETSTICKYKEYS = 59"
 .AddItem "SPI_SETTOGGLEKEYS = 53"
 .AddItem "SPI_SETTOOLTIPANIMATION = &H1017"
 .AddItem "SPI_SETTOOLTIPFADE = &H1019"
 .AddItem "SPI_SETUIEFFECTS = &H103F"
 .AddItem "SPI_SETWHEELSCROLLLINES = 105"
 .AddItem "SPI_SETWORKAREA = 47"
     
             End With
 
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub TrackPopupMenu(ListName As ListBox)
 
 frmConstants.Visible = True
            With ListName
.AddItem "TPM_BOTTOMALIGN = &H20&"
.AddItem "TPM_CENTERALIGN = &H4&"
.AddItem "TPM_HORIZONTAL = &H0&"
.AddItem "TPM_HORNEGANIMATION = &H800&"
.AddItem "TPM_HORPOSANIMATION = &H400&"
.AddItem "TPM_LEFTALIGN = &H0&"
.AddItem "TPM_LEFTBUTTON = &H0&"
.AddItem "TPM_NOANIMATION = &H4000&"
.AddItem "TPM_NONOTIFY = &H80&"
.AddItem "TPM_RECURSE = &H1&"
.AddItem "TPM_RETURNCMD = &H100&"
.AddItem "TPM_RIGHTALIGN = &H8&"
.AddItem "TPM_RIGHTBUTTON = &H2&"
.AddItem "TPM_TOPALIGN = &H0&"
.AddItem "TPM_VCENTERALIGN = &H10&"
.AddItem "TPM_VERNEGANIMATION = &H2000&"
.AddItem "TPM_VERPOSANIMATION = &H1000&"
.AddItem "TPM_VERTICAL = &H40&"
     
            End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub SetWindPosConst(ListName As ListBox)
 
 frmConstants.Visible = True
            With ListName
           
   .AddItem "SWP_HIDEWINDOW = &H80"
   .AddItem "SWP_ASYNCWINDOWPOS = &H4000"
   .AddItem "SWP_DEFERERASE = &H2000"
   .AddItem "SWP_FRAMECHANGED = &H20"
   .AddItem "SWP_NOACTIVATE = &H10"
   .AddItem "SWP_NOCOPYBITS = &H100"
   .AddItem "SWP_NOMOVE = &H2"
   .AddItem "SWP_NOOWNERZORDER = &H200"
   .AddItem "SWP_NOSENDCHANGING = &H400"
   .AddItem "SWP_NOSIZE = &H1"
   .AddItem "SWP_NOZORDER = &H4"
   .AddItem "SWP_SHOWWINDOW = &H40"
   .AddItem "SWP_REFRESH = (&H1 or &H2 or &H4 or &H20)"
   .AddItem "SWP_ONTOP =(&H1 or &H2)"
    
              End With
End Sub


'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
 Sub GetWindowLongConst(ListName As ListBox)
 
 frmConstants.Visible = True
           With ListName
           
.AddItem "GWL_EXSTYLE = -20"
.AddItem "GWL_HINSTANCE = -6"
.AddItem "GWL_HWNDPARENT = -8"
.AddItem "GWL_ID = -12"
.AddItem "GWL_STYLE = -16"
.AddItem "GWL_USERDATA = -21"
.AddItem "GWL_WNDPROC = -4"
.AddItem "DWL_DLGPROC = 4"
.AddItem "DWL_MSGRESULT = 0"
.AddItem "DWL_USER = 8"
.AddItem "DWLP_DLGPROC = DWLP_MSGRESULT + sizeof(LRESULT)"
.AddItem "DWLP_MSGRESULT = 0"

           End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for menu related and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub MenuConstants(ListName As ListBox)
 
 frmConstants.Visible = True
           With ListName
           
.AddItem "MF_BITMAP = &H4&"
.AddItem "MF_BYCOMMAND = &H0&"
.AddItem "MF_BYPOSITION = &H400&"
.AddItem "MF_APPEND = &H100&"
.AddItem "MF_CALLBACKS = &H8000000"
.AddItem "MF_CHANGE = &H80&"
.AddItem "MF_CHECKED = &H8&"
.AddItem "MF_CONV = &H40000000"
.AddItem "MF_DEFAULT = &H1000&"
.AddItem "MF_DELETE = &H200&"
.AddItem "MF_DISABLED = &H2&"
.AddItem "MF_ENABLED = &H0&"
.AddItem "MF_END = &H80"
.AddItem "MF_ERRORS = &H10000000"
.AddItem "MF_GRAYED = &H1&"
.AddItem "MF_REMOVE = &H1000&"
.AddItem "MF_RIGHTJUSTIFY = &H4000&"
.AddItem "MF_POSTMSGS = &H4000000"
.AddItem "MF_POPUP = &H10&"
.AddItem "MF_OWNERDRAW = &H100&"
.AddItem "MF_MOUSESELECT = &H8000&"
.AddItem "MF_MENUBREAK = &H40&"
.AddItem "MF_MENUBARBREAK = &H20&"
.AddItem "MF_MASK = &HFF000000"
.AddItem "MF_LINKS = &H20000000"
.AddItem "MF_INSERT = &H0&"
.AddItem "MF_HSZ_INFO = &H1000000"
.AddItem "MF_HILITE = &H80&"
.AddItem "MF_UNCHECKED = &H0&"
.AddItem "MF_SYSMENU = &H2000&"
.AddItem "MF_STRING = &H0&"
.AddItem "MF_SEPARATOR = &H800&"
.AddItem "MF_SENDMSGS = &H2000000"
.AddItem "MF_UNHILITE = &H0&"
.AddItem "MF_USECHECKBITMAPS = &H200&"
               
               End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub SetWindLongWSConst(ListName As ListBox)
 
 frmConstants.Visible = True
           With ListName
           
.AddItem "GWL_STYLE = -16"
.AddItem "WS_ACTIVECAPTION = &H1"
.AddItem "WS_BORDER = &H800000"
.AddItem "WS_CAPTION = &HC00000"
.AddItem "WS_CHILD = &H40000000"
.AddItem "WS_CLIPCHILDREN = &H2000000"
.AddItem "WS_DISABLED = &H8000000"
.AddItem "WS_DLGFRAME = &H400000"
.AddItem "WS_GROUP = &H20000"
.AddItem "WS_TABSTOP = &H10000"
.AddItem "WS_HSCROLL = &H100000"
.AddItem "WS_ICONIC = WS_MINIMIZE"
.AddItem "WS_MINIMIZE = &H20000000"
.AddItem "WS_MAXIMIZE = &H1000000"
.AddItem "WS_MAXIMIZEBOX = &H10000"
.AddItem "WS_MINIMIZEBOX = &H20000"
.AddItem "WS_OVERLAPPED = &H0&"
.AddItem "WS_SYSMENU = &H80000"
.AddItem "WS_THICKFRAME = &H40000"
.AddItem "WS_POPUP = &H80000000"
.AddItem "WS_SIZEBOX = WS_THICKFRAME"
.AddItem "WS_TILED = WS_OVERLAPPED"
.AddItem "WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW"
      
                         End With
End Sub



'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub SetWindLongWS_EX_Const(ListName As ListBox)
 
 frmConstants.Visible = True
           With ListName
           
.AddItem "GWL_EXSTYLE = -20"
.AddItem "WS_EX_CLIENTEDGE = &H200&"
.AddItem "WS_EX_CONTEXTHELP = &H400&"
.AddItem "WS_EX_CONTROLPARENT = &H10000&"
.AddItem "WS_EX_DLGMODALFRAME = &H1&"
.AddItem "WS_EX_LAYERED = &H80000"
.AddItem "WS_EX_LAYOUTRTL = &H400000&"
.AddItem "WS_EX_LEFT = &H0&"
.AddItem "WS_EX_LEFTSCROLLBAR = &H4000&"
.AddItem "WS_EX_LTRREADING = &H0&"
.AddItem "WS_EX_MDICHILD = &H40&"
.AddItem "WS_EX_NOACTIVATE = &H8000000&"
.AddItem "WS_EX_NOINHERITLAYOUT = &H100000&"
.AddItem "WS_EX_NOPARENTNOTIFY = &H4&"
.AddItem "WS_EX_WINDOWEDGE = &H100&"
.AddItem "WS_EX_TOOLWINDOW = &H80&"
.AddItem "WS_EX_TOPMOST = &H8&"
.AddItem "WS_EX_RIGHT = &H1000&"
.AddItem "WS_EX_RIGHTSCROLLBAR = &H0&"
.AddItem "WS_EX_RTLREADING = &H2000&"
.AddItem "WS_EX_STATICEDGE = &H20000&"
.AddItem "WS_EX_TRANSPARENT = &H20&"
              
                 End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub RedrawWindowConst(ListName As ListBox)
 
 frmConstants.Visible = True
           With ListName
           
.AddItem "RDW_ALLCHILDREN = &H80"
.AddItem "RDW_ERASE = &H4"
.AddItem "RDW_ERASENOW = &H200"
.AddItem "RDW_FRAME = &H400"
.AddItem "RDW_INTERNALPAINT = &H2"
.AddItem "RDW_INVALIDATE = &H1"
.AddItem "RDW_NOCHILDREN = &H40"
.AddItem "RDW_NOERASE = &H20"
.AddItem "RDW_NOFRAME = &H800"
.AddItem "RDW_NOINTERNALPAINT = &H10"
.AddItem "RDW_UPDATENOW = &H100"
.AddItem "RDW_VALIDATE = &H8"

                 End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub GetWindowConst(ListName As ListBox)
 
 frmConstants.Visible = True
               With ListName
           
.AddItem "GW_CHILD = 5"
.AddItem "GW_ENABLEDPOPUP = 6"
.AddItem "GW_HWNDFIRST = 0"
.AddItem "GW_HWNDLAST = 1"
.AddItem "GW_HWNDNEXT = 2"
.AddItem "GW_HWNDPREV = 3"
.AddItem "GW_MAX = 5"
.AddItem "GW_OWNER = 4"

                End With
End Sub



'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will show input box that has the user choose
'            public or private and then pastes the CreateHatchBrush Constants
'----------------------------------------------------------------------
Sub CreateHatchBrushConst(ListName As ListBox)
 
 frmConstants.Visible = True
               With ListName
           
.AddItem "HS_BDIAGONAL = 3"
.AddItem "HS_BDIAGONAL1 = 7"
.AddItem "HS_CROSS = 4"
.AddItem "HS_DENSE1 = 9"
.AddItem "HS_DENSE2 = 10"
.AddItem "HS_DENSE3 = 11"
.AddItem "HS_DENSE4 = 12"
.AddItem "HS_DENSE5 = 13"
.AddItem "HS_DENSE6 = 14"
.AddItem "HS_DENSE7 = 15"
.AddItem "HS_DENSE8 = 16"
.AddItem "HS_DIAGCROSS = 5"
.AddItem "HS_DITHEREDBKCLR = 24"
.AddItem "HS_DITHEREDCLR = 20"
.AddItem "HS_DITHEREDTEXTCLR = 22"
.AddItem "HS_FDIAGONAL = 2"
.AddItem "HS_FDIAGONAL1 = 6"
.AddItem "HS_HALFTONE = 18"
.AddItem "HS_HORIZONTAL = 0"
.AddItem "HS_NOSHADE = 17"
.AddItem "HS_SOLID = 8"
.AddItem "HS_SOLIDBKCLR = 23"
.AddItem "HS_SOLIDCLR = 19"
.AddItem "HS_SOLIDTEXTCLR = 21"
.AddItem "HS_VERTICAL = 1"

                End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api(Device indep bmp and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub DIB(ListName As ListBox)
 
 frmConstants.Visible = True
               With ListName
           
.AddItem "DIB_PAL_COLORS As Long = 1"
.AddItem "DIB_RGB_COLORS As Long = 0"

       End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub DrawFrameControlConst(ListName As ListBox)
 
 frmConstants.Visible = True
               With ListName
           
.AddItem "DFCBUTTON = 4"
.AddItem "DFCCAPTION = 1"
.AddItem "DFCMENU = 2"
.AddItem "DFCPOPUPMENU = 5"
.AddItem "DFCSCROLL = 3"
.AddItem "DFCSADJUSTRECT = &H2000"
.AddItem "DFCSBUTTON3STATE = &H8"
.AddItem "DFCSBUTTONCHECK = &H0"
.AddItem "DFCSBUTTONPUSH = &H10"
.AddItem "DFCSBUTTONRADIO = &H4"
.AddItem "DFCSBUTTONRADIOIMAGE = &H1"
.AddItem "DFCSBUTTONRADIOMASK = &H2"
.AddItem "DFCSCAPTIONCLOSE = &H0"
.AddItem "DFCSCAPTIONHELP = &H4"
.AddItem "DFCSCAPTIONMAX = &H2"
.AddItem "DFCSCAPTIONMIN = &H1"
.AddItem "DFCSCAPTIONRESTORE = &H3"
.AddItem "DFCSFLAT = &H4000"
.AddItem "DFCSHOT = &H1000"
.AddItem "DFCSINACTIVE = &H100"
.AddItem "DFCSMENUARROW = &H0"
.AddItem "DFCSMENUARROWRIGHT = &H4"
.AddItem "DFCSMENUBULLET = &H2"
.AddItem "DFCSMENUCHECK = &H1"
.AddItem "DFCSMONO = &H8000"
.AddItem "DFCSPUSHED = &H200"
.AddItem "DFCSSCROLLCOMBOBOX = &H5"
.AddItem "DFCSSCROLLDOWN = &H1"
.AddItem "DFCSSCROLLLEFT = &H2"
.AddItem "DFCSSCROLLRIGHT = &H3"
.AddItem "DFCSSCROLLSIZEGRIP = &H8"
.AddItem "DFCSSCROLLSIZEGRIPRIGHT = &H10"
.AddItem "DFCSSCROLLUP = &H0"
.AddItem "DFCSTRANSPARENT = &H800"

                 End With
End Sub



'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub GetSysColorBrushConst(ListName As ListBox)
 
 frmConstants.Visible = True
               With ListName
           
.AddItem "COLOR3DDKSHADOW = 21"
.AddItem "COLORBTNFACE = 15"
.AddItem "COLORBTNHIGHLIGHT = 20"
.AddItem "COLOR3DLIGHT = 22"
.AddItem "COLORBTNSHADOW = 16"
.AddItem "COLORACTIVEBORDER = 10"
.AddItem "COLORACTIVECAPTION = 2"
.AddItem "COLORADD = 712"
.AddItem "COLORADJMAX = 100"
.AddItem "COLORADJMIN = -100"
.AddItem "COLORAPPWORKSPACE = 12"
.AddItem "COLORBACKGROUND = 1"
.AddItem "COLORBLUE = 708"
.AddItem "COLORBLUEACCEL = 728"
.AddItem "COLORBOX1 = 720"
.AddItem "COLORBTNTEXT = 18"
.AddItem "COLORCAPTIONTEXT = 9"
.AddItem "COLORCURRENT = 709"
.AddItem "COLORCUSTOM1 = 721"
.AddItem "COLORELEMENT = 716"
.AddItem "COLORGRADIENTACTIVECAPTION = 27"
.AddItem "COLORGRADIENTINACTIVECAPTION = 28"
.AddItem "COLORGRAYTEXT = 17"
.AddItem "COLORGREEN = 707"
.AddItem "COLORGREENACCEL = 727"
.AddItem "COLORHIGHLIGHT = 13"
.AddItem "COLORHIGHLIGHTTEXT = 14"
.AddItem "COLORHOTLIGHT = 26"
.AddItem "COLORHUE = 703"
.AddItem "COLORHUEACCEL = 723"
.AddItem "COLORHUESCROLL = 700"
.AddItem "COLORINACTIVEBORDER = 11"
.AddItem "COLORINACTIVECAPTION = 3"
.AddItem "COLORINACTIVECAPTIONTEXT = 19"
.AddItem "COLORINFOBK = 24"
.AddItem "COLORINFOTEXT = 23"
.AddItem "COLORLUM = 705"
.AddItem "COLORLUMACCEL = 725"
.AddItem "COLORLUMSCROLL = 702"
.AddItem "COLORMATCHVERSION = &H200"
.AddItem "COLORMENU = 4"
.AddItem "COLORMENUTEXT = 7"
.AddItem "COLORMIX = 719"
.AddItem "COLORNOTRANSPARENT = &HFFFFFFFF"
.AddItem "COLORPALETTE = 718"
.AddItem "COLORRAINBOW = 710"
.AddItem "COLORRED = 706"
.AddItem "COLORREDACCEL = 726"
.AddItem "COLORSAMPLES = 717"
.AddItem "COLORSAT = 704"
.AddItem "COLORSATACCEL = 724"
.AddItem "COLORSATSCROLL = 701"
.AddItem "COLORSAVE = 711"
.AddItem "COLORSCHEMES = 715"
.AddItem "COLORSCROLLBAR = 0"
.AddItem "COLORSOLID = 713"
.AddItem "COLORSOLIDLEFT = 730"
.AddItem "COLORSOLIDRIGHT = 731"
.AddItem "COLORTUNE = 714"
.AddItem "COLORWINDOW = 5"
.AddItem "COLORWINDOWFRAME = 6"
.AddItem "COLORWINDOWTEXT = 8"

                   End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub ShowWindowConst(ListName As ListBox)

 frmConstants.Visible = True
               With ListName
               
 .AddItem "SW_AUTOPROF_LOAD_MASK = &H1"
 .AddItem "SW_AUTOPROF_SAVE_MASK = &H2"
 .AddItem "SW_ERASE = &H4"
 .AddItem "SW_FORCEMINIMIZE = 11"
 .AddItem "SW_HIDE = 0"
 .AddItem "SW_INVALIDATE = &H2"
 .AddItem "SW_MAXIMIZE = 3"
 .AddItem "SW_MAX = 10"
 .AddItem "SW_MINIMIZE = 6"
 .AddItem "SW_OTHERUNZOOM = 4"
 .AddItem "SW_OTHERZOOM = 2"
 .AddItem "SW_PARENTCLOSING = 1"
 .AddItem "SW_PARENTOPENING = 3"
 .AddItem "SW_RESTORE = 9"
 .AddItem "SW_SCROLLCHILDREN = &H1"
 .AddItem "SW_SHOW = 5"
 .AddItem "SW_SHOWMAXIMIZED = 3"
 .AddItem "SW_SHOWDEFAULT = 10"
 .AddItem "SW_SHOWMINIMIZED = 2"
 .AddItem "SW_SHOWMINNOACTIVE = 7"
 .AddItem "SW_SHOWNA = 8"
 .AddItem "SW_SHOWNOACTIVATE = 4"
 .AddItem "SW_SHOWNORMAL = 1"
 .AddItem "SW_SMOOTHSCROLL = &H10"

               End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for this api and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub Mouse_EventConst(ListName As ListBox)

 frmConstants.Visible = True
               With ListName
               
 .AddItem "MOUSEEVENTF_ABSOLUTE = &H8000"
 .AddItem "MOUSEEVENTF_LEFTDOWN = &H2"
 .AddItem "MOUSEEVENTF_LEFTUP = &H4"
 .AddItem "MOUSEEVENTF_MIDDLEDOWN = &H20"
 .AddItem "MOUSEEVENTF_MIDDLEUP = &H40"
 .AddItem "MOUSEEVENTF_MOVE = &H1"
 .AddItem "MOUSEEVENTF_RIGHTDOWN = &H8"
 .AddItem "MOUSEEVENTF_RIGHTUP = &H10"
 .AddItem "MOUSEEVENTF_VIRTUALDESK = &H4000"
 .AddItem "MOUSEEVENTF_WHEEL = &H800"
 .AddItem "MOUSEEVENTF_XDOWN = &H80"
 .AddItem "MOUSEEVENTF_XUP = &H100"
               
               End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for graphx API(copyimage) and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub CopyImage(ListName As ListBox)

 frmConstants.Visible = True
               With ListName
.AddItem "IMAGE_BITMAP = 0"
.AddItem "IMAGE_ICON = 1"
.AddItem "IMAGE_CURSOR = 2"
.AddItem "LR_COPYDELETEORG = &H8"
.AddItem "LR_COPYRETURNORG = &H4"
.AddItem "LR_MONOCHROME = &H1"
.AddItem "LR_LOADTRANSPARENT = &H20"
.AddItem "LR_LOADMAP3DCOLORS = &H1000"
.AddItem "LR_LOADFROMFILE = &H10"
.AddItem "LR_DEFAULTSIZE = &H40"
.AddItem "LR_DEFAULTCOLOR = &H0"
.AddItem "LR_CREATEDIBSECTION = &H2000"
.AddItem "LR_COPYFROMRESOURCE = &H4000"
.AddItem "LR_COLOR = &H2"
.AddItem "LR_SHARED = &H8000"
.AddItem "LR_VGACOLOR = &H80"

      End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for graphx API and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub BitSRCConst(ListName As ListBox)

 frmConstants.Visible = True
               With ListName
               
.AddItem "SRCAND = &H8800C6"
.AddItem "SRCCOPY = &HCC0020"
.AddItem "SRCERASE = &H440328"
.AddItem "SRCINVERT = &H660046"
.AddItem "SRCPAINT = &HEE0086"
.AddItem "BLACKNESS = &H42"
.AddItem "DSTINVERT = &H550009"
.AddItem "MERGECOPY = &HC000CA"
.AddItem "MERGEPAINT = &HBB0226"
.AddItem "NOTSRCCOPY = &H330008"
.AddItem "NOTSRCERASE = &H1100A6"
.AddItem "PATCOPY = &HF00021"
.AddItem "PATINVERT = &H5A0049"
.AddItem "PATPAINT = &HFB0A09"
.AddItem "WHITENESS = &HFF0062"
         
             End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for CreatePen API and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub PenStylesConst(ListName As ListBox)

 frmConstants.Visible = True
               With ListName
               
.AddItem "PS_ALTERNATE = 8"
.AddItem "PS_COSMETIC = &H0"
.AddItem "PS_DASH = 1"
.AddItem "PS_DASHDOT = 3"
.AddItem "PS_DASHDOTDOT = 4"
.AddItem "PS_DOT = 2"
.AddItem "PS_ENDCAP_FLAT = &H200"
.AddItem "PS_ENDCAP_MASK = &HF00"
.AddItem "PS_ENDCAP_ROUND = &H0"
.AddItem "PS_ENDCAP_SQUARE = &H100"
.AddItem "PS_GEOMETRIC = &H10000"
.AddItem "PS_INSIDEFRAME = 6"
.AddItem "PS_JOIN_BEVEL = &H1000"
.AddItem "PS_JOIN_MASK = &HF000"
.AddItem "PS_JOIN_MITER = &H2000"
.AddItem "PS_JOIN_ROUND = &H0"
.AddItem "PS_MAXLINKTYPES = 8"
.AddItem "PS_NULL = 5"
.AddItem "PS_OPENTYPE_FONTTYPE = &H10000"
.AddItem "PS_SOLID = 0"
          
            End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for drawedge API and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub DrawEdgeConst(ListName As ListBox)

 frmConstants.Visible = True
               With ListName
               
.AddItem "BDR_INNER = &HC"
.AddItem "BDR_RAISED = &H5"
.AddItem "BDR_OUTER = &H3"
.AddItem "BDR_RAISEDINNER = &H4"
.AddItem "BDR_RAISEDOUTER = &H1"
.AddItem "BDR_SUNKEN = &HA"
.AddItem "BDR_SUNKENINNER = &H8"
.AddItem "BDR_SUNKENOUTER = &H2"
.AddItem "EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)"
.AddItem "EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)"
.AddItem "EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)"
.AddItem "EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)"

              End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for PolyGonRgn API and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub PolygonRgnConst(ListName As ListBox)

 frmConstants.Visible = True
              With ListName
               
.AddItem "ALTERNATE = 1"
.AddItem "WINDING = 2"

              End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this will place the constants for CombineRgn API and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub CombineRgnConst(ListName As ListBox)

 frmConstants.Visible = True
            With ListName
               
.AddItem "RGN_AND = 1"
.AddItem "RGN_COPY = 5"
.AddItem "RGN_DIFF = 4"
.AddItem "RGN_OR = 2"
.AddItem "RGN_XOR = 3"

            End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | this will place the constants for PlaySound API and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub SoundConst(ListName As ListBox)

 frmConstants.Visible = True
            With ListName
            
.AddItem "SND_ALIAS = &H10000"
.AddItem "SND_ALIAS_ID = &H110000"
.AddItem "SND_ALIAS_START = 0"
.AddItem "SND_APPLICATION = &H80"
.AddItem "SND_ASYNC = &H1"
.AddItem "SND_FILENAME = &H20000"
.AddItem "SND_LOOP = &H8"
.AddItem "SND_MEMORY = &H4"
.AddItem "SND_NODEFAULT = &H2"
.AddItem "SND_NOSTOP = &H10"
.AddItem "SND_NOWAIT = &H2000"
.AddItem "SND_PURGE = &H40"
.AddItem "SND_RESERVED = &HFF000000"
.AddItem "SND_RESOURCE = &H40004"
.AddItem "SND_SYNC = &H0"
.AddItem "SND_TYPE_MASK = &H170007"
.AddItem "SND_VALID = &H1F"
.AddItem "SND_VALIDFLAGS = &H17201F"

.AddItem "' the following values get placed directly in the  lpszName   parameter"
.AddItem Chr(34) & "EmptyRecycleBin" & Chr(34) & " ' when recycle bin is emptied"
.AddItem Chr(34) & "SystemExclamation" & Chr(34) & " ' when windows shows a warning"
.AddItem Chr(34) & "SystemExit" & Chr(34) & "       ' when Windows shuts down"
.AddItem Chr(34) & "Maximize" & Chr(34) & "        ' when a program is maximized"
.AddItem Chr(34) & "MenuCommand" & Chr(34) & "      ' when a menu item is clicked on"
.AddItem Chr(34) & "MenuPopup" & Chr(34) & "      ' when a (sub)menu pops up"
.AddItem Chr(34) & "Minimize" & Chr(34) & "      ' when a program is minimized to taskbar"
.AddItem Chr(34) & "MailBeep" & Chr(34) & "        ' when email is received"
.AddItem Chr(34) & "Open" & Chr(34) & "              ' when a program is opened"
.AddItem Chr(34) & "SystemHand" & Chr(34) & "      ' when a critical stop occurs"
.AddItem Chr(34) & "AppGPFault" & Chr(34) & "       ' when a program causes an error"
.AddItem Chr(34) & "SystemQuestion" & Chr(34) & "  ' when a system question occurs"
.AddItem Chr(34) & "RestoreDown" & Chr(34) & "    ' when a program is restored to normal size"
.AddItem Chr(34) & "RestoreUp" & Chr(34) & "      ' when a program is restored to normal size from taskbar"
.AddItem Chr(34) & "SystemStart" & Chr(34) & "    ' when Windows starts up"
.AddItem Chr(34) & "Close" & Chr(34) & "           ' when program is closed"
.AddItem Chr(34) & "Ringout" & Chr(34) & "       ' when (fax) call is made outbound and the line is ringing"
.AddItem Chr(34) & "RingIn" & Chr(34) & "         ' incoming (fax) call"

              End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | this will place the constants for MessageBeep API and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub MessageBeepConst(ListName As ListBox)
 
 frmConstants.Visible = True
            With ListName
            
.AddItem "MB_ICONHAND = &H10&"
.AddItem "MB_ICONEXCLAMATION = &H30&"
.AddItem "MB_ICONASTERISK = &H40&"
.AddItem "MB_ICONQUESTION = &H20&"

           End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | this will place the constants for CreateFile API and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub CreateFileConst(ListName As ListBox)
 
 frmConstants.Visible = True
            With ListName
            
.AddItem "GENERIC_WRITE = &H40000000"
.AddItem "GENERIC_READ = &H80000000"
.AddItem "FILE_SHARE_WRITE = &H2"
.AddItem "FILE_SHARE_READ = &H1"
.AddItem "FILE_SHARE_DELETE = &H4"
.AddItem "CREATE_NEW = 1"
.AddItem "CREATE_ALWAYS = 2"
.AddItem "OPEN_ALWAYS = 4"
.AddItem "OPEN_EXISTING = 3"
.AddItem "TRUNCATE_EXISTING = 5"
.AddItem "FILE_ATTRIBUTE_ARCHIVE = &H20"
.AddItem "FILE_ATTRIBUTE_COMPRESSED = &H800"
.AddItem "FILE_ATTRIBUTE_NORMAL = &H80"
.AddItem "FILE_ATTRIBUTE_HIDDEN = &H2"
.AddItem "FILE_ATTRIBUTE_OFFLINE = &H1000"
.AddItem "FILE_ATTRIBUTE_READONLY = &H1"
.AddItem "FILE_ATTRIBUTE_SYSTEM = &H4"
.AddItem "FILE_ATTRIBUTE_TEMPORARY = &H100"
.AddItem "FILE_FLAG_WRITE_THROUGH = &H80000000"
.AddItem "FILE_FLAG_OVERLAPPED = &H40000000"
.AddItem "FILE_FLAG_NO_BUFFERING = &H20000000"
.AddItem "FILE_FLAG_RANDOM_ACCESS = &H10000000"
.AddItem "FILE_FLAG_SEQUENTIAL_SCAN = &H8000000"
.AddItem "FILE_FLAG_DELETE_ON_CLOSE = &H4000000"
.AddItem "FILE_FLAG_BACKUP_SEMANTICS = &H2000000"
.AddItem "FILE_FLAG_POSIX_SEMANTICS = &H1000000"
.AddItem "SECURITY_ANONYMOUS = (SecurityAnonymous * 2 ^ 16)"
.AddItem "SecurityAnonymous = 1"
.AddItem "SECURITY_SQOS_PRESENT = &H100000"
.AddItem "SECURITY_IDENTIFICATION = (SecurityIdentification * 2 ^ 16)"
.AddItem "SecurityIdentification = 2"
.AddItem "SECURITY_IMPERSONATION = (SecurityImpersonation * 2 ^ 16)"
.AddItem "SECURITY_DELEGATION = (SecurityDelegation * 2 ^ 16)"
.AddItem "SECURITY_CONTEXT_TRACKING = &H40000"
.AddItem "SECURITY_EFFECTIVE_ONLY = &H80000"

            End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | this will place the constants for RegisterHotkey API and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub RegisterHotKeyConst(ListName As ListBox)
         
frmConstants.Visible = True
          With ListName
          
.AddItem "MOD_ALT = &H1"
.AddItem "MOD_CONTROL = &H2"
.AddItem "MOD_SHIFT = &H4"

          End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | this will place the constants for DrawText API and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub DrawTextConst(ListName As ListBox)

 frmConstants.Visible = True
            With ListName
            
.AddItem "DT_BOTTOM = &H8"
.AddItem "DT_CALCRECT = &H400"
.AddItem "DT_CENTER = &H1"
.AddItem "DT_LEFT = &H0"
.AddItem "DT_MULTILINE = (&H1)"
.AddItem "DT_NOCLIP = &H100"
.AddItem "DT_PASSWORD_EDIT = (&H10)"
.AddItem "DT_RIGHT = &H2"
.AddItem "DT_RTLREADING = &H20000"
.AddItem "DT_TOP = &H0"
.AddItem "DT_SINGLELINE = &H20"
           
           End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | this will place most of the keycode constants and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub KeyCodeConst(ListName As ListBox)
             
frmConstants.Visible = True
             With ListName
             
.AddItem "VK_ACCEPT = &H1E"
.AddItem "VK_ADD = &H6B"
.AddItem "VK_ATTN = &HF6"
.AddItem "VK_BACK = &H8"
.AddItem "VK_BROWSER_BACK = &HA6"
.AddItem "VK_BROWSER_FAVORITES = &HAB"
.AddItem "VK_BROWSER_FORWARD = &HA7"
.AddItem "VK_BROWSER_HOME = &HAC"
.AddItem "VK_BROWSER_REFRESH = &HA8"
.AddItem "VK_BROWSER_SEARCH = &HAA"
.AddItem "VK_BROWSER_STOP = &HA9"
.AddItem "VK_CANCEL = &H3"
.AddItem "VK_CAPITAL = &H14"
.AddItem "VK_CLEAR = &HC"
.AddItem "VK_CONTROL = &H11"
.AddItem "VK_CONVERT = &H1C"
.AddItem "VK_CRSEL = &HF7"
.AddItem "VK_DECIMAL = &H6E"
.AddItem "VK_DELETE = &H2E"
.AddItem "VK_DIVIDE = &H6F"
.AddItem "VK_DOWN = &H28"
.AddItem "VK_END = &H23"
.AddItem "VK_ESCAPE = &H1B"
.AddItem "VK_EXECUTE = &H2B"
.AddItem "VK_EXSEL = &HF8"
.AddItem "VK_F1 = &H70"
.AddItem "VK_F10 = &H79"
.AddItem "VK_F11 = &H7A"
.AddItem "VK_F12 = &H7B"
.AddItem "VK_F2 = &H71"
.AddItem "VK_F3 = &H72"
.AddItem "VK_F5 = &H74"
.AddItem "VK_F4 = &H73"
.AddItem "VK_F6 = &H75"
.AddItem "VK_F7 = &H76"
.AddItem "VK_F8 = &H77"
.AddItem "VK_F9 = &H78"
.AddItem "VK_HELP = &H2F"
.AddItem "VK_HOME = &H24"
.AddItem "VK_INSERT = &H2D"
.AddItem "VK_LBUTTON = &H1"
.AddItem "VK_LAUNCH_MAIL = &HB4"
.AddItem "VK_LAUNCH_APP1 = &HB6"
.AddItem "VK_LAUNCH_APP2 = &HB7"
.AddItem "VK_LCONTROL = &HA2"
.AddItem "VK_LEFT = &H25"
.AddItem "VK_LMENU = &HA4"
.AddItem "VK_LSHIFT = &HA0"
.AddItem "VK_MBUTTON = &H4"
.AddItem "VK_MULTIPLY = &H6A"
.AddItem "VK_MENU = &H12"
.AddItem "VK_NEXT = &H22"
.AddItem "VK_NUMLOCK = &H90"
.AddItem "VK_NUMPAD0 = &H60"
.AddItem "VK_NUMPAD1 = &H61"
.AddItem "VK_NUMPAD2 = &H62"
.AddItem "VK_NUMPAD3 = &H63"
.AddItem "VK_NUMPAD4 = &H64"
.AddItem "VK_NUMPAD5 = &H65"
.AddItem "VK_NUMPAD6 = &H66"
.AddItem "VK_NUMPAD8 = &H68"
.AddItem "VK_NUMPAD7 = &H67"
.AddItem "VK_NUMPAD9 = &H69"
.AddItem "VK_PRINT = &H2A"
.AddItem "VK_PRIOR = &H21"
.AddItem "VK_RBUTTON = &H2"
.AddItem "VK_RCONTROL = &HA3"
.AddItem "VK_RETURN = &HD"
.AddItem "VK_RIGHT = &H27"
.AddItem "VK_RMENU = &HA5"
.AddItem "VK_RSHIFT = &HA1"
.AddItem "VK_SCROLL = &H91"
.AddItem "VK_SELECT = &H29"
.AddItem "VK_SEPARATOR = &H6C"
.AddItem "VK_SHIFT = &H10"
.AddItem "VK_SLEEP = &H5F"
.AddItem "VK_SNAPSHOT = &H2C"
.AddItem "VK_SPACE = &H20"
.AddItem "VK_SUBTRACT = &H6D"
.AddItem "VK_TAB = &H9"
.AddItem "VK_UP = &H26"
.AddItem "VK_VOLUME_DOWN = &HAE"
.AddItem "VK_VOLUME_MUTE = &HAD"
.AddItem "VK_VOLUME_UP = &HAF"

             End With
End Sub
 
 
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | this will place most of the BroadcastSystemMessage constants and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub BroadcastSysMsgCont(ListName As ListBox)

frmConstants.Visible = True
          With ListName
          
.AddItem "BSF_FLUSHDISK = &H4"
.AddItem "BSF_FORCEIFHUNG = &H20"
.AddItem "BSF_IGNORECURRENTTASK = &H2"
.AddItem "BSF_NOHANG = &H8"
.AddItem "BSF_POSTMESSAGE = &H10"
.AddItem "BSF_NOTIMEOUTIFNOTHUNG = &H40"
.AddItem "BSF_QUERY = &H1"
.AddItem "BSM_ALLCOMPONENTS = &H0"
.AddItem "BSM_ALLDESKTOPS = &H10"
.AddItem "BSM_APPLICATIONS = &H8"
.AddItem "BSM_INSTALLABLEDRIVERS = &H4"
.AddItem "BSM_NETDRIVER = &H2"
.AddItem "BSM_VXDS = &H1"

           End With
End Sub


'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | this will place most of the GetQueStatus constants and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub GetQueStatusConst(ListName As ListBox)

frmConstants.Visible = True
          With ListName
          
.AddItem "QS_SENDMESSAGE = &H40"
.AddItem "QS_MOUSEBUTTON = &H4"
.AddItem "QS_MOUSEMOVE = &H2"
.AddItem "QS_KEY = &H1"
.AddItem "QS_ALLPOSTMESSAGE = &H100"


           End With
End Sub



'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | this will place most of the PeekMessage constants and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub PeekMsgConst(ListName As ListBox)

frmConstants.Visible = True
          With ListName
          
.AddItem "PM_NOREMOVE = &H0"
.AddItem "PM_REMOVE = &H1"

           End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | this will place most of the PostMessage constants and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub PostMsgConst(ListName As ListBox)

frmConstants.Visible = True
          With ListName
          
.AddItem "HWND_BROADCAST = &HFFFF&"

           End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: | this will place most of the SendMsgTimeout  constants and place
'             in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub SendMsgTimeoutConst(ListName As ListBox)

frmConstants.Visible = True
          With ListName
          
.AddItem "SMTO_ABORTIFHUNG = &H2"
.AddItem "SMTO_BLOCK = &H1"
.AddItem "SMTO_NORMAL = &H0"
.AddItem "SMTO_NOTIMEOUTIFNOTHUNG = &H8"


           End With
End Sub
