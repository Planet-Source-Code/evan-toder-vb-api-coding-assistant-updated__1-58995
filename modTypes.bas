Attribute VB_Name = "modTypes"
Option Explicit


'     .     ' _ __     .    .
'_____ _   _ | '_ \ ___  _ _'
'_   _| |_| || |_) / _ \/ __|
' | |  \__, || .__/  __/\__ \
' |_|  |___/ |_|   \___||___/






'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the WINDOWPLACEMENT structure
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub WINDOWPLACEMENT(ListName As ListBox)

 frmType.Visible = True
               With ListName
.AddItem "Type POINTAPI"
.AddItem "        x As Long"
.AddItem "        y As Long"
.AddItem "End Type"
.AddItem "Type RECT"
.AddItem "        Left As Long"
.AddItem "        Top As Long"
.AddItem "        Right As Long"
.AddItem "        Bottom As Long"
.AddItem "End Type"
.AddItem "Type WINDOWPLACEMENT"
.AddItem "        Length As Long"
.AddItem "        flags As Long"
.AddItem "        showCmd As Long"
.AddItem "        ptMinPosition As POINTAPI"
.AddItem "        ptMaxPosition As POINTAPI"
.AddItem "        rcNormalPosition As RECT"
.AddItem "End Type"
         End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the WNDCLASS structure
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub WNDCLASS(ListName As ListBox)

 frmType.Visible = True
               With ListName
.AddItem "Private Type WNDCLASS"
.AddItem "    style As Long"
.AddItem "    lpfnwndproc As Long"
.AddItem "    cbClsextra As Long"
.AddItem "    cbWndExtra2 As Long"
.AddItem "    hInstance As Long"
.AddItem "    hIcon As Long"
.AddItem "    hCursor As Long"
.AddItem "    hbrBackground As Long"
.AddItem "    lpszMenuName As String"
.AddItem "    lpszClassName As String"
.AddItem "End Type"
          End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the FLASHWINF0 structure
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub FlashWindowEx(ListName As ListBox)

 frmType.Visible = True
               With ListName
.AddItem "Private Type FLASHWINFO"
.AddItem "    cbSize As Long"
.AddItem "    hwnd As Long"
.AddItem "    dwFlags As Long"
.AddItem "    uCount As Long"
.AddItem "    dwTimeout As Long"
.AddItem "End Type"
          End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the PT structure
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub PT(ListName As ListBox)

 frmType.Visible = True
               With ListName
.AddItem "Type POINTAPI"
.AddItem "    X As Long"
.AddItem "    Y As Long"
.AddItem "End Type"
          End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the getDiBits related api and place
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub GetDiBits(ListName As ListBox)

 frmType.Visible = True
               With ListName
.AddItem "Type BITMAPINFOHEADER '40 bytes"
.AddItem "        biSize As Long"
.AddItem "        biWidth As Long"
.AddItem "        biHeight As Long"
.AddItem "        biPlanes As Integer"
.AddItem "        biBitCount As Integer"
.AddItem "        biCompression As Long"
.AddItem "        biSizeImage As Long"
.AddItem "        biXPelsPerMeter As Long"
.AddItem "        biYPelsPerMeter As Long"
.AddItem "        biClrUsed As Long"
.AddItem "        biClrImportant As Long"
.AddItem "End Type"
.AddItem "Type RGBQUAD"
.AddItem "        rgbBlue As Byte"
.AddItem "        rgbGreen As Byte"
.AddItem "        rgbRed As Byte"
.AddItem "        rgbReserved As Byte"
.AddItem "End Type"
.AddItem "Type BITMAPINFO"
.AddItem "        bmiHeader As BITMAPINFOHEADER"
.AddItem "        bmiColors As RGBQUAD"
.AddItem "End Type"
                       End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the getColorAdjustment related api and place
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub COLORADJUSTMENT(ListName As ListBox)

 frmType.Visible = True
               With ListName
.AddItem "Type COLORADJUSTMENT"
.AddItem "        caSize As Integer"
.AddItem "        caFlags As Integer"
.AddItem "        caIlluminantIndex As Integer"
.AddItem "        caRedGamma As Integer"
.AddItem "        caGreenGamma As Integer"
.AddItem "        caBlueGamma As Integer"
.AddItem "        caReferenceBlack As Integer"
.AddItem "        caReferenceWhite As Integer"
.AddItem "        caContrast As Integer"
.AddItem "        caBrightness As Integer"
.AddItem "        caColorfulness As Integer"
.AddItem "        caRedGreenTint As Integer"
.AddItem "End Type"
                  End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the getBitmapBits related api and place
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub BITMAP(ListName As ListBox)

 frmType.Visible = True
               With ListName
               
.AddItem "Type BITMAP"
.AddItem "    bmType As Long"
.AddItem "    bmWidth As Long"
.AddItem "    bmHeight As Long"
.AddItem "    bmWidthBytes As Long"
.AddItem "    bmPlanes As Integer"
.AddItem "    bmBitsPixel As Integer"
.AddItem "    bmBits As Long"
.AddItem "End Type"
             End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the gdialphablend related api and place
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub BLENDFUNCTION(ListName As ListBox)

 frmType.Visible = True
               With ListName
               
.AddItem "Type BLENDFUNCTION"
.AddItem "  BlendOp As Byte"
.AddItem "  BlendFlags As Byte"
.AddItem "  SourceConstantAlpha As Byte"
.AddItem "  AlphaFormat As Byte"
.AddItem "End Type"
              End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the createpenindirect related api and place
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub LOGPEN(ListName As ListBox)

 frmType.Visible = True
               With ListName
               
.AddItem "Type LOGPEN"
.AddItem "        lopnStyle As Long"
.AddItem "        lopnWidth As POINTAPI"
.AddItem "        lopnColor As Long"
.AddItem "End Type"

          End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the BitmapInfo related api and place
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub BITMAPINFO(ListName As ListBox)

 frmType.Visible = True
               With ListName
               
.AddItem "Type BITMAPINFOHEADER '40 bytes"
.AddItem "        biSize As Long"
.AddItem "        biWidth As Long"
.AddItem "        biHeight As Long"
.AddItem "        biPlanes As Integer"
.AddItem "        biBitCount As Integer"
.AddItem "        biCompression As Long"
.AddItem "        biSizeImage As Long"
.AddItem "        biXPelsPerMeter As Long"
.AddItem "        biYPelsPerMeter As Long"
.AddItem "        biClrUsed As Long"
.AddItem "        biClrImportant As Long"
.AddItem "End Type"

.AddItem "Type BITMAPINFO"
.AddItem "        bmiHeader As BITMAPINFOHEADER"
.AddItem "        bmiColors As RGBQUAD"
.AddItem "End Type"

               End With
End Sub



'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the menuiteminfo related api and place
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub TypeMenuItemInfo(ListName As ListBox)

 frmType.Visible = True
               With ListName

.AddItem "Private Type MENUITEMINFO"
.AddItem "   cbSize As Long"
.AddItem "   fMask As Long"
.AddItem "   fType As Long"
.AddItem "   fState As Long"
.AddItem "   wID As Long"
.AddItem "   hSubMenu As Long"
.AddItem "   hbmpChecked As Long"
.AddItem "   hbmpUnchecked As Long"
.AddItem "   dwItemData As Long"
.AddItem "   dwTypeData As String"
.AddItem "   cch As Long"
.AddItem "End Type"

         End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the brush related api and place
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub LogBrush(ListName As ListBox)

 frmType.Visible = True
               With ListName
.AddItem "Type LogBrush"
.AddItem "    lbStyle As Long"
.AddItem "    lbColor As Long"
.AddItem "    lbHatch As Long"
.AddItem "End Type"

          End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the polygon rgn for this api and place
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub PolygonRgn(ListName As ListBox)

 frmType.Visible = True
               With ListName
   
.AddItem "Type COORD"
.AddItem "   X As Long"
.AddItem "   Y As Long"
.AddItem "End Type"
   
   End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the createdc for this api and place
'            in frmTypes.lstTypes
'----------------------------------------------------------------------
Sub DEVMODE(ListName As ListBox)

 frmType.Visible = True
               With ListName
   
.AddItem "Type DEVMODE"
.AddItem "    dmDeviceName As String * CCDEVICENAME"
.AddItem "    dmSpecVersion As Integer"
.AddItem "    dmDriverVersion As Integer"
.AddItem "    dmSize As Integer"
.AddItem "    dmDriverExtra As Integer"
.AddItem "    dmFields As Long"
.AddItem "    dmOrientation As Integer"
.AddItem "    dmPaperSize As Integer"
.AddItem "    dmPaperLength As Integer"
.AddItem "    dmPaperWidth As Integer"
.AddItem "    dmScale As Integer"
.AddItem "    dmCopies As Integer"
.AddItem "    dmDefaultSource As Integer"
.AddItem "    dmPrintQuality As Integer"
.AddItem "    dmColor As Integer"
.AddItem "    dmDuplex As Integer"
.AddItem "    dmYResolution As Integer"
.AddItem "    dmTTOption As Integer"
.AddItem "    dmCollate As Integer"
.AddItem "    dmFormName As String * CCFORMNAME"
.AddItem "    dmUnusedPadding As Integer"
.AddItem "    dmBitsPerPel As Integer"
.AddItem "    dmPelsWidth As Long"
.AddItem "    dmPelsHeight As Long"
.AddItem "    dmDisplayFlags As Long"
.AddItem "    dmDisplayFrequency As Long"
.AddItem "End Type"
                   End With
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the TITLEBAR structure for this api and place
'            in frmTypes.lstType
'----------------------------------------------------------------------
Sub TypeTitleBarInfo(ListName As ListBox)

 frmType.Visible = True
               With ListName
               
.AddItem "Type RECT"
.AddItem "    Left As Long"
.AddItem "    Top As Long"
.AddItem "    Right As Long"
.AddItem "    Bottom As Long"
.AddItem "End Type"
.AddItem ""
.AddItem "Type TITLEBARINFO"
.AddItem "       cbSize As Long"
.AddItem "       rcTitleBar As RECT"
.AddItem "       rgstate(CCHILDREN_TITLEBAR) As Long"
.AddItem "End Type"

                  End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the RECT structure for this api and place
'            in frmType.lstTypes
'----------------------------------------------------------------------
Sub TypeRect(ListName As ListBox)

 frmType.Visible = True
               With ListName
 _
.AddItem "Type RECT"
.AddItem "    Left As Long"
.AddItem "    Top As Long"
.AddItem "    Right As Long"
.AddItem "    Bottom As Long"
.AddItem "End Type"

                 End With
End Sub

 

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the Polygon structure for this api and place
'            in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub TypePolygonRgn(ListName As ListBox)

 frmType.Visible = True
               With ListName
               
.AddItem "Type COORD"
.AddItem "    x As Long"
.AddItem "    y As Long"
.AddItem "End Type"
   
               End With
End Sub

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will place the Message structure for msg related
'            api and place in frmConstants.lstConstants
'----------------------------------------------------------------------
Sub TypeMessage(ListName As ListBox)

 frmType.Visible = True
               With ListName
.AddItem "Type POINTAPI"
.AddItem "    x As Long"
.AddItem "    y As Long"
.AddItem "End Type"
.AddItem "Type Msg"
.AddItem "    hWnd As Long"
.AddItem "    message As Long"
.AddItem "    wParam As Long"
.AddItem "    lParam As Long"
.AddItem "    time As Long"
.AddItem "    pt As POINTAPI"
.AddItem "End Type"

               End With
End Sub
