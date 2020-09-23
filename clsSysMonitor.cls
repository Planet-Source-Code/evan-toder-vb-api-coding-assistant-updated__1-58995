VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public picRAM                     As PictureBox
Public picCPU                     As PictureBox
Public WithEvents Timer           As Timer
Attribute Timer.VB_VarHelpID = -1

  
Dim bOnce As Boolean

'''''''RAM monitoring variables
Private clrFactor    As Currency
Private TotRam       As Long
Private AvailRam     As Long
Private UsedRam      As Long
''''''''''''''''''''''''''''''''

''''''CPU monitoring variables
Private PercCPUload  As Long
Private sk As Long
Private HQ As Long  ' handle to Query
Private counter As Long
Private Const lType = 4
Private Const lSize = 4
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const HKEY_DYN_DATA As Long = &H80000006
'''''''''''''''''''''''''''''''


Event TotSysRam(sVal As String)
Event RamInfo(sAvailRam As String, sUsedram As String)
Event CPUloadInfo(sCpuPercentLoad As String)
 
 
Private Sub Class_Initialize()

       bOnce = True
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | starts the timer
'----------------------------------------------------------------------
Sub StartMonitoring()

        Timer.Interval = 1500
        Timer.Enabled = True
End Sub

 

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | calls the 2 main functions every 1.5 sec that tell us
'             what we want to know about current Ram and CPU load
'----------------------------------------------------------------------
Private Sub Timer_Timer()
       
       'init function is called once to initilize settings
        If bOnce = True Then
             Call InitializeCPUstuff
             Call InitializeRAMstuff
             bOnce = False
        End If
        
       Call GetRam
       DoEvents
       Call GetCpu
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | updates the status of cpu load every 1.5 seconds
'----------------------------------------------------------------------
Private Sub GetCpu()
Dim lData As Long
Dim hKey As Long
 
       
        If IsOsWinXP Then  '                      if windows xp
             Call PdhCollectQueryData(HQ)
             PercCPUload = CLng(PdhVbGetDoubleCounterValue( _
                  counter, lData))
                  
             RaiseEvent CPUloadInfo(CStr(PercCPUload))
             
        Else  '                                    if os other than xp
             Call RegOpenKey( _
                  HKEY_DYN_DATA, _
                  "PerfStats\StartStat", _
                  hKey)
                  
             Call RegQueryValueEx( _
                  hKey, _
                  "KERNEL\CPUUsage", _
                  0, lType, lData, lSize)
                  
             Call RegCloseKey(hKey)
             
             Call RegOpenKey( _
                  HKEY_DYN_DATA, _
                  "PerfStats\StatData", sk)
     
             Call RegQueryValueEx( _
                  sk, "KERNEL\CPUUsage", _
                  0, lType, lData, lSize)
                  
             PercCPUload = lData
             
             RaiseEvent CPUloadInfo(CLng(PercCPUload))
          End If

          Call PaintCpuGraph
End Sub

'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | updates the status of ram every 1.5 seconds
'----------------------------------------------------------------------
Private Sub GetRam()
    Dim MemStat As MEMORYSTATUSEX
    Dim TotalPhys As Long
    'initialize structure
    MemStat.dwLength = Len(MemStat)
    'retireve memory information
    GlobalMemoryStatusEx MemStat
    'convert large integer to currency
    TotalPhys = LargeIntToCurrency(MemStat.ullTotalPhys)
    'show result
    AvailRam = Format(((MemStat.ullAvailPhys.LowPart / 1024) / 1000))
    TotRam = (TotalPhys / 1024 ^ 2)
    UsedRam = (TotRam - AvailRam)
    RaiseEvent RamInfo(CStr(AvailRam), CStr(UsedRam))
    
    Call PaintRamGraph
End Sub



'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | to properly obtain cpu info we need to know OS
'----------------------------------------------------------------------
Private Function IsOsWinXP() As Boolean

    Dim vi As OSVERSIONINFO
    vi.dwOSVersionInfoSize = Len(vi)
    Call GetVersionEx(vi)
    IsOsWinXP = (vi.dwPlatformId = VER_PLATFORM_WIN32_NT)
    
End Function


'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | stuff that needs to be processed once
'----------------------------------------------------------------------
Private Sub InitializeCPUstuff()
Dim lData           As Long
Dim hKey            As Long
Dim r               As Long

      picCPU.AutoRedraw = True
      picCPU.ScaleWidth = 100 'percent
      
      If IsOsWinXP Then
           Call PdhVbOpenQuery(HQ)
           Call PdhVbAddCounter(HQ, "\Processor(0)\% Processor Time", counter)
        
           Call PdhCollectQueryData(HQ)
           Call PdhVbGetDoubleCounterValue(counter, lData)
      End If
End Sub
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: |  stuff that needs to be processed once
'----------------------------------------------------------------------
Private Sub InitializeRAMstuff()
        Call GetRam
        picRAM.ScaleWidth = TotRam
        picRAM.AutoRedraw = True
        'the purpose of this is to create uniform color
        'gradiency in the picturebox regardless of the
        'total amount of ram the user has
        'this variable acts a a picurebox color buffer
        clrFactor = (TotRam / 255)
        
        RaiseEvent TotSysRam(CStr(TotRam))
        picRAM.ToolTipText = "total system RAM: " & TotRam & " mb"
        Timer.Interval = 1500
        Timer.Enabled = True
End Sub

 
'----------------------------------------------------------------------
'   INPUTS: | NONE
'  RETURNS: | NONE
' COMMENTS: | this actually paints the RAM representation in picRam
'----------------------------------------------------------------------
Private Sub PaintRamGraph()
 Dim i                     As Integer
 
 On Error Resume Next
 
       With picRAM
         .Cls
         'picbox scale width is val of used ram
          For i = 0 To UsedRam
            'the more ram thats used, the farther to the right
            'edge of picbox drawing goes, and the less blue
            'and more red show to visually create warning lever
            'on low ram
            picRAM.Line _
                (i, 0)-(i, .ScaleHeight), _
                RGB((i * clrFactor), 100, (250 - (i * clrFactor)))
          Next i
       End With
End Sub

Private Sub PaintCpuGraph()
 Dim i                     As Integer
 
 On Error Resume Next
 
       With picCPU
         .Cls
         'picbox scale width is val of used ram
          For i = 0 To PercCPUload
            'the more cpu load there is, the farther to the right
            'edge of picbox drawing goes, and the less blue
            'and more red show to visually create warning lever
            'on low ram
            picCPU.Line _
                (i, 0)-(i, .ScaleHeight), _
                RGB((i * 2), 100, (250 - (i * 2)))
          Next i
       End With

End Sub


Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
    'copy 8 bytes from the large integer to an empty currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
    'adjust it
    LargeIntToCurrency = LargeIntToCurrency * 10000
End Function


 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
