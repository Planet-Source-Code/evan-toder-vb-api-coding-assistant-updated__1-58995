Attribute VB_Name = "modPub"
Option Explicit

'this is referenced from
 'modSaveClipboardCodeBlockData  and
Public Enum enumClipboardOrCodeBlock
         ClipboardItem = 0
         CodeBlock = 1
End Enum



'############# PUBLIC VARIABLES TO SAVE TO FILE ##############

'options menu on frmBar
Public bEnableScreenTips                     As Boolean
Public bShowComputerStats                    As Boolean
Public bAutoShowConstForm                    As Boolean

'on frmPOP(combobox)
'the number being the listindex of the combobox
Public MailCheckFrequency                    As Integer

'(0,x)=pop mail server
'(1,x)=pop mail username
'(2,x) = pop mail password
Public POPaccountInfo(2, 3)                  As String

'checkbox on frmPOP
Public bEnablePOPchecking                    As Boolean
'path to mail client
Public sMailClientPath                       As String
'##############################################################
'true if frmMsg is loaded..prevents more than 1 instance
Public FrmMsgIsLoaded                        As Boolean
Public Const APP_NAME = "VBcodePaste"

Public bSendingApiNotConst                   As Boolean
'keeps track of the current
'arrnum(POPaccountInfo) were currently checking
Public iPub_CurrPopAccountNum                As Integer

Public cScreenText                           As clsScreenText




 
'----------------------------------------------------------------------
'   INPUTS: |sPath: path saving program data to
'            varVals: the array of program variables your saving to file
'  RETURNS: |
' COMMENTS: |
'----------------------------------------------------------------------
Sub sub_FileSaveData(sPath As String, ParamArray varVals())
  Dim fFile        As Integer
  Dim i            As Integer
      
      fFile = FreeFile
      
      Open sPath For Output As #fFile
           'each variable value to be saved
           '(varVals) saved on its own line
           For i = 0 To UBound(varVals)
               Write #fFile, varVals(i)
           Next i
      Close fFile
      
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
  Dim s            As String
  Dim fFile        As Integer
  Dim arr()        As Variant
  Dim i            As Integer
  
  On Error GoTo ERR:
  
      fFile = FreeFile
      
      Open sPath For Input As #fFile
             Do Until EOF(fFile)
                Input #fFile, s
                ReDim Preserve arr(i)
                arr(i) = s
                i = (i + 1)
            Loop
      Close fFile
      
      func_FileLoadData = arr
      
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
       
ERR:
   Exit Function
End Function
 

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |toggles a menu item on and multiple off
'----------------------------------------------------------------------
Sub ToggleMenuOn(MenuNameCheck As Menu, ParamArray MenuNamesUncheck())
 Dim i         As Integer
   
       ' untoggle items in menuNameUnchecked
       For i = 0 To UBound(MenuNamesUncheck)
             MenuNamesUncheck(i).Checked = False
       Next i
       
       ' toggle menu item on
       MenuNameCheck.Checked = True
End Sub




