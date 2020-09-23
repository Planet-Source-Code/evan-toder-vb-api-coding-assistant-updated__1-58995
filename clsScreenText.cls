VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private Const WM_SYSCOMMAND As Long = &H112

Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private WithEvents Frm      As Form
Attribute Frm.VB_VarHelpID = -1



'----------------------------------------------------------------------
'   INPUTS: |          F: formname
'                   sMsg: text to show on screen                    |string
'              iFontSize: size of the font                          |integer
'              sFontType: style of font to use  i.e "Comic Sans"    |string
'             lFontColor: color of font to use                      |long
'             bFontBold: if the font is bold or not                 |boolean
'            bTextFloat: places form(text) on top                   |boolean
'        bAllowTextMove: user can drag and move the text            |boolean
'             iTextLeft: Left starting point of the text            |integer
'              iTextTop: Top starting point of the text             |integer
'  RETURNS: |
' COMMENTS: | text is printed on a form (F) and where the text is
'             painted is recorded as the path
'             tben the forms visible region is defined by
'             that path
'----------------------------------------------------------------------
Sub ScreenMsg(sMsg$, iFontSize%, _
                            lFontColor&, bFontBold As Boolean, _
                            iTextLeft%, iTextTop%, _
                            Optional sFontType$ = "Arial Black")
'VARIABLES:
    Dim hRgn&
    Dim i%, iTop%
    Dim sParts() As String
'CODE:
    Unload frmMsg
    
    With frmMsg
    'the font to use for the screen message
     .FontName = sFontType
    
    'font size
     .FontSize = iFontSize%
    
    'fontbold?
      .FontBold = bFontBold
    
    'set the backcolor lFontColor
     .BackColor = lFontColor&

    'open a path bracket
     BeginPath .hdc
    
    'split the message by CR
     sParts = Split(sMsg, vbCrLf)
    
    'icrement the Y part of text drawing
    'to = the fontsize + 1
     For i = 0 To UBound(sParts)
       'draw the text
       TextOut .hdc, 2, iTop, sParts(i), Len(sParts(i))
       iTop = (iTop + (iFontSize + 2))
     Next i
    
    'close the path bracket
     EndPath .hdc
    
    'convert the path to a region
     hRgn = PathToRegion(.hdc)
    
    'set the Window-region
     SetWindowRgn .hwnd, hRgn, True
    
    'destroy our region
    DeleteObject hRgn
    
    Set Frm = frmMsg
    
    'show the form
    Load frmMsg
    
    'form(text) position
     .Move iTextLeft, iTextTop
    
   'show, dont activate
     ShowWindow .hwnd, 4
     
  End With
'END CODE:

End Sub

 
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |kill form reference
'----------------------------------------------------------------------
Private Sub Class_Terminate()
   On Error Resume Next
   
       Set Frm = Nothing
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |keep the msg on top
'----------------------------------------------------------------------
Private Sub Frm_Load()
'
       SetWindowPos Frm.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub

 
