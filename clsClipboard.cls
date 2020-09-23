VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum ClipType
      Clip_Board
      Code_Block
End Enum


Private fFile              As Integer
Private mod_Clipboard()    As String
Private mod_CodeBlock()    As String


Public Property Get CurrClipContents(Which As ClipType, Index%) As String
  '---------------------------------------
  'since arr cant be public in object module
  'this passes the value ,to frmBar,
  'of the clipboard item the user wants to send
  '--------------------------------------
   If Which = Clip_Board Then
       CurrClipContents = mod_Clipboard(Index)
   Else
       CurrClipContents = mod_CodeBlock(Index)
   End If
End Property


 
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this is called when user clicks "Apply" in frmClip
'----------------------------------------------------------------------
Sub AddToClipboard(Which As ClipType, sData$, mnuName As Variant, _
                        Optional bLoadingCodeBlocks As Boolean = False, _
                        Optional strCodeMenu As String)
'VARIABLES::
    Dim iUpp%
    Dim mnuLeft$
'CODE::
    'redimension both the arr holding clipb data
    'and the clipboard item itself
    If Which = Clip_Board Then
       iUpp = UBound(mod_Clipboard) + 1
       Load mnuName(iUpp)
       ReDim Preserve mod_Clipboard(iUpp)
      'add the new clipboard data
       mod_Clipboard(iUpp) = sData
      'show first 40 chr's in menu as reminder
      'to what clipboard contents are
       mnuLeft = Left(sData, 40)
       ' enable "remove" menu item
       If frmBar.mnuClipboardRemove.Enabled = False Then
           frmBar.mnuClipboardRemove.Enabled = True
       End If
    Else
       iUpp = UBound(mod_CodeBlock) + 1
       Load mnuName(iUpp)
       ReDim Preserve mod_CodeBlock(iUpp)
       'add the new clipboard data
       mod_CodeBlock(iUpp) = sData
       '
       If bLoadingCodeBlocks = False Then
            mnuLeft = InputBox("Enter a description for this code block that " & _
                      "will serve as the menu caption.")
       Else
            mnuLeft = strCodeMenu
       End If
       
       ' enable "remove" menu item
       If frmBar.mnuRemoveCodeBlock.Enabled = False Then
           frmBar.mnuRemoveCodeBlock.Enabled = True
       End If
    End If
    '
    mnuName(iUpp).Caption = mnuLeft
    mnuName(iUpp).Visible = True
'END CODE::
End Sub

Sub RemoveFromClipboard(Which As ClipType, Index%, mnuName As Variant)
'VARIABLES::
  Dim i%, iUpper%
'CODE::
    iUpper = mnuName.UBound
   'move elements from menu 1 step down to
   'the point of the index of the item to be removed
    For i = Index To (iUpper - 1)
        mnuName(i).Caption = mnuName(i + 1).Caption
        '
        If Which = Clip_Board Then
           'do the same with the holder of clipb contents
            mod_Clipboard(i) = mod_Clipboard(i + 1)
        Else
            mod_CodeBlock(i) = (mod_CodeBlock(i + 1))
        End If
    Next i
    
    'remove highest menu item
    Unload mnuName(i)
    
    If Which = Clip_Board Then
       'and the highest arr of dec_Clipboard
       ReDim Preserve mod_Clipboard(i - 1)
       'disable "remove" if there are no more mnuitems to remove
       If iUpper <= 1 Then
           frmBar.mnuClipboardRemove.Enabled = False
       End If
    Else
       'and the highest arr of dec_Clipboard
       ReDim Preserve mod_CodeBlock(i - 1)
       'disable "remove" if there are no more mnuitems to remove
       If iUpper <= 1 Then
           frmBar.mnuRemoveCodeBlock.Enabled = False
       End If
    End If
    '
    Unload frmWaitForClipboardRemove
'END CODE::
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |saves the contents of dec_Clipboard to file
'----------------------------------------------------------------------
Sub SaveClipContents(Which As ClipType, cmDialog As CommonDialog)
 On Error GoTo ERR:
'VARIABLES:
  Dim i%
'CODE:
  fFile = FreeFile
  '
  Open CmnDlgReturn(cmDialog, False) For Output As #fFile
    If Which = Clip_Board Then
        'saving clipboard contents
        For i = 1 To UBound(mod_Clipboard)
           Print #fFile, mod_Clipboard(i) & vbCrLf & "ENDL"
        Next i
    Else
        'saving codeblock contents
        For i = 1 To UBound(mod_CodeBlock)
           Print #fFile, mod_CodeBlock(i) & vbCrLf
           Print #fFile, "menucaption: " & frmBar.mnuArrCode(i).Caption
        Next i
    End If
  Close #fFile
'END CODE:
ERR:
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |load contents of clipboard file to mod_Clipboard
'----------------------------------------------------------------------
Sub LoadClipContents(Which As ClipType, mnuItem As Variant, cmDialog As CommonDialog)
 On Error GoTo ERR:
'VARIABLES:
  Dim sTemp$, sTemp2$
'CODE:
  fFile = FreeFile
  '
  Open CmnDlgReturn(cmDialog, True) For Input As #fFile
    Do Until EOF(fFile)
       Line Input #fFile, sTemp
          '
          If Which = Clip_Board Then
              If Trim(sTemp) = "ENDL" Then
                'function adds item to clipboard and the menu
                 Call AddToClipboard(Clip_Board, sTemp2, mnuItem)
                 sTemp2 = ""
              Else
                  sTemp2 = (sTemp2 & sTemp & vbCrLf)
              End If
          Else
              If Left(Trim(sTemp), 12) = "menucaption:" Then
                 Dim strMnuCaption  As String
                 strMnuCaption = Mid(sTemp, 13, Len(sTemp) - 12)
                 'function adds item to codeblock and the menu
                 Call AddToClipboard(Code_Block, sTemp2, mnuItem, True, strMnuCaption)
                 sTemp2 = ""
              Else
                  sTemp2 = (sTemp2 & sTemp & vbCrLf)
              End If
          End If
          '
          DoEvents
    Loop
  Close #fFile
'END CODE:
ERR:
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |this will remove all items from the menu and mod_clipboard
'----------------------------------------------------------------------
Sub ClearAll(Which As ClipType)
'VARIABLES:
  Dim i%, iYN%, mnuUpper%
  Dim s$
'CODE:
  If Which = Clip_Board Then
    'user wants to save the menu items to file
    iYN = MsgBox("Do you wish to save the " & CStr(UBound(mod_Clipboard)) & _
                      " clipboard items ", vbYesNo)
    If iYN = vbYes Then
        Call SaveClipContents(Clip_Board, frmBar.cmDlg)
    End If
    'erase the arr holding the clipboards strings
    Erase mod_Clipboard
    ReDim mod_Clipboard(0)
   'clear the menu items representing the arr
    For i = 1 To frmBar.mnuArrClipboard.UBound
        Unload frmBar.mnuArrClipboard(i)
    Next i
  Else
    'user wants to save the menu items to file
    iYN = MsgBox("Do you wish to save the " & CStr(UBound(mod_CodeBlock)) & _
                      " code block items ")
    If iYN = vbYes Then
       Call SaveClipContents(Clip_Board, frmBar.cmDlg)
    End If
    'erase the array holding the codeblocks
    Erase mod_CodeBlock
    ReDim mod_CodeBlock(0)
   'clear the menu items representing the arr
    For i = 1 To frmBar.mnuArrCode.UBound
        Unload frmBar.mnuArrCode(i)
    Next i
  End If
'END CODE:
End Sub
 

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |the filepath chosen
' COMMENTS: |preset basic common dialog settings..used in load clipboard file
'            or save clipboard to file
'----------------------------------------------------------------------
Function CmnDlgReturn(cmDlgName As CommonDialog, bShowOpen As Boolean) As String
   On Error GoTo ERR:
'CODE::
  With cmDlgName
     .Filter = "Valid text based files(*.txt, *.wpd, *.ini)|*.txt;*.wpd;*.ini"
     .InitDir = "C:\My Documents"
     .Flags = &H2 'warn b4 overwriting
     .CancelError = True
     
     If bShowOpen = True Then
        .ShowOpen
     Else
        .ShowSave
     End If
     '
     CmnDlgReturn = .FileName
  End With
Exit Function
'END CODE::
ERR:
  If ERR.Number = 32755 Then
     CmnDlgReturn = ""
  End If
End Function

Private Sub Class_Initialize()
     ReDim mod_Clipboard(0)
     ReDim mod_CodeBlock(0)
End Sub
