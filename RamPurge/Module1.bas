Attribute VB_Name = "ModAPI"
Option Explicit






Public Sx&, Sy&


Public Const WM_SYSCOMMAND As Long = &H112
Public Const CB_SHOWDROPDOWN As Long = &H14F

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function ReleaseCapture Lib "user32.dll" () As Long
Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long



'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | STRING
' COMMENTS: | holds the code related to moving controls
'----------------------------------------------------------------------
Sub mod_Move(lHwnd&)
'CODE:
        ReleaseCapture
        SendMessage lHwnd, WM_SYSCOMMAND, &HF012&, 0&
'END CODE:
End Sub

Sub DropDownComboBox(lHwnd)
On Error GoTo ERR:
'---------------------------------
'opens a combo box with code
'---------------------------------
'VARIABLES:

'CODE:
  SendMessage lHwnd, CB_SHOWDROPDOWN, True, 0&
'END CODE:
'exit sub
ERR:
  Debug.Print ERR.Description

End Sub

Sub SetRgn(F As Variant, CornerRoundness%)
On Error GoTo ERR:
'---------------------------------
'create a forms  shape
'---------------------------------
'VARIABLES:
  Dim hRgn&
'CODE:
  hRgn = CreateRoundRectRgn( _
         1, 1, (F.Width / Sx), (F.Height / Sy), _
         CornerRoundness, CornerRoundness _
         )
  Call SetWindowRgn(F.hWnd, hRgn, True)
'END CODE:
'exit sub
ERR:
  Debug.Print ERR.Description
End Sub


Function func_GetMemoryStats() As Long()
On Error GoTo ERR:
'---------------------------------
'returns total system ram
' and available ram
'---------------------------------
'VARIABLES:
  Dim mem As MEMORYSTATUS
  Dim arr(1) As Long
'CODE:
  Call GlobalMemoryStatus(mem)
  'return megabytes (totalbytes/one million)
  arr(0) = (mem.dwTotalPhys / 1000000)
  arr(1) = (mem.dwAvailPhys / 1000000)
  '
  func_GetMemoryStats = arr
'END CODE:
Exit Function
ERR:
  Debug.Print ERR.Description
End Function


Function DeterminPurgeAmount(used&, free&, percent%) As Long
On Error GoTo ERR:
'---------------------------------
'the formula to use:
'5-15% or used ram (5% if minimal
'10% if moderate, 15% if aggressive)
' + the free ram
'so if you have 250 mg ram
'120 mg is used (assume were using
'aggressive setting) 15% of that is
' 18 + 130(free ram)
'so well purge 148 mg ram
'remember..each loop in sub PurgeRam (
'Main Form) purges 1 mb of ram so in this
'instance we need to loop 148 times
'---------------------------------
'VARIABLES:

'CODE:
  DeterminPurgeAmount = (((percent * 0.01) * used) + free)
'END CODE:
Exit Function
ERR:
  Debug.Print ERR.Description
End Function


'----------------------------------------------------------------------
'   INPUTS: |sPath: path saving program data to
'            varVals: the array of program variables your saving to file
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Sub sub_FileSaveData(sPath As String, ParamArray varVals())
'VARIABLES:
  Dim fFile%, i%
'CODE:
      fFile = FreeFile
      
      Open sPath For Output As #fFile
           'each variable value to be saved
           '(varVals) saved on its own line
           For i = 0 To UBound(varVals)
               Write #fFile, varVals(i)
           Next i
      Close fFile
'END CODE:
      'SAMPLE USAGE__________________________
           'say you have two variables from
           'your program you want to filesave
              'bTimerEnabled
              'intTimerCount
              ':
      'Private Sub Form_Unload ____________________
      '
      '   call sub_FileSaveData("strWhateverPathYouWant", _
      '                        bTimerEnabled, intTimerCount)
      'End Sub_____________________________________
End Sub
'----------------------------------------------------------------------
'   INPUTS: |sPath:path to the file that holds variabled for our program
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Function func_FileLoadData(sPath As String) As Variant()
'VARIABLES:
   Dim s$
   Dim fFile%, i%
   Dim arr()   As Variant
'CODE:
   
      fFile = FreeFile
  '!!!#### IMPORTANT !!! : The first time the person runs the prog,
 '                         the file to load has not yet been created.
 '      YOU MUST:
 '                   1)hand create the file, which gives
 '                      you the flexibility of presetting prog
 '                      defaults to whatever you wish
 '         OR:
 '                   2)user error handling in this sub, using
 '                     API call "CreateFile"
 '                     (in the "API miscellaneous" menu)
 '                     followed by  "exit sub"
      Open sPath For Input As #fFile
             Do Until EOF(fFile)
                Input #fFile, s
                ReDim Preserve arr(i)
                arr(i) = s
                i = (i + 1)
            Loop
      Close fFile
      
      func_FileLoadData = arr
'END CODE:
      
      'SAMPLE USEAGE___________________________
      'say your program has two variable values
      'you want filled from file(bTimerEnabled, intTimerCount)
               ':
      'Private Sub Form_Load_____________________
       ' Dim arr()   As Variant
       ' Dim i       as integer
       '
       '  arr = func_FileLoadData("strWhateverPathYouWant")
       '
       '  For i = LBound(arr) To UBound(arr)
       '      Select Case i
       '          Case Is = 0
       '                bTimerEnabled = Cbool(arr(i))
       '          Case =1
       '                intTimerCount = Cint(arr(i))
       '      End Select
       '  Next i
       'End Sub____________________________________
End Function
 

