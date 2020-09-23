VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




Event POPconnected()
Event POPmsgsWaiting(NumMsgs As Integer)
Event POPstatechange(currState As enumPOPstate, POPserverCommunication As String, _
                                                POPclientCommunication As String)

Public Enum enumPOPstate
        POP_CONNECTED = 0
        POP_USEROK = 1
        POP_PASSOK = 2
        POP_MAILSTAT = 3
        POP_TOP = 4
End Enum


Private iMsgCount                   As Integer
Private m_POPaddress                As String
Private m_POPusername               As String
Private m_POPpassword               As String
Private m_POPserverCommunication    As String
Private m_POPclientCommunication    As String
Private m_popState                  As enumPOPstate
Public WithEvents YourPopSock       As Winsock
Attribute YourPopSock.VB_VarHelpID = -1
 
 
' _ __      . _   .    . _     .  '    .     . _   .    .
'| '_ \_   _ | |__  _   (_) ___   ' _ _'_   _ | |__  _ _'
'| |_)  | | ||  _ \| |  | |/ __|  '/ __| | | ||  _ \/ __|
'| .__/ |_| || |_) | |_ | | (__   '\__ \ |_| || |_) \__ \
'|_|   \__,_||_.__/|___||_|\___|  '|___/\__,_||_.__/|___/



'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |CONNECT WITH THE MAIL SERVER
'----------------------------------------------------------------------
Sub ConnectMailServer()

       YourPopSock.Close
       YourPopSock.Connect POPaddress, 110
End Sub
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |SEND TERMINATE COMMUNICATION NOTICE TO SERVER AND CLOSE SOCKET
'----------------------------------------------------------------------
Sub CloseMailServer()

      If YourPopSock.State = 2 Then
          YourPopSock.SendData "QUIT" & vbCrLf
          DoEvents
          YourPopSock.Close
      Else
          YourPopSock.Close
      End If
End Sub
Private Sub YourPopSock_Connect()
      
      RaiseEvent POPconnected
End Sub

Private Sub YourPopSock_DataArrival(ByVal bytesTotal As Long)
 Dim sData               As String
 
    With YourPopSock
          .GetData sData
          POPserverCommunication = sData
          
          'we must first make sure has sent an OK reply b4 proceeding
          If LCase(Left(Trim(sData), 3)) = "+ok" Then
                'send mail server the username
                If m_popState = POP_CONNECTED Then
                    POPclientCommunication = "USER " & POPusername & vbCrLf
                    .SendData POPclientCommunication
                    m_popState = POP_USEROK
                    '
                    RaiseEvent POPstatechange(POP_USEROK, _
                                               POPserverCommunication, _
                                               POPclientCommunication)
                    
                'send mail server the password
                ElseIf m_popState = POP_USEROK Then
                    POPclientCommunication = "PASS " & POPpassword & vbCrLf
                    .SendData POPclientCommunication
                    m_popState = POP_PASSOK
                    '
                    RaiseEvent POPstatechange(POP_PASSOK, _
                                              POPserverCommunication, _
                                              POPclientCommunication)
                                              
                'ask mail server how many messages and how big they are
                ElseIf m_popState = POP_PASSOK Then
                     POPclientCommunication = "STAT " & vbCrLf
                     .SendData POPclientCommunication
                     m_popState = POP_MAILSTAT
                     '
                     RaiseEvent POPstatechange(POP_MAILSTAT, _
                                              POPserverCommunication, _
                                              POPclientCommunication)
                    
                'at this point we will get a response like so...
                ' +OK 5 23990 which means 5 message are waiting
                'that are 23,990 byes in size total
                'we will split this by the spaces..and
                'this middle value is number of msgs
                'which will will place in integer so we
                'can use in loop to deal with each indiv msg
                ElseIf m_popState = POP_MAILSTAT Then
                      iMsgCount = funcMsgCount(POPserverCommunication)
                      RaiseEvent POPmsgsWaiting(iMsgCount)
                      RaiseEvent POPstatechange(POP_TOP, _
                                              POPserverCommunication, _
                                              POPclientCommunication)
                                         
                End If
               
          'we didnt recieve an ok response from the server
           Else
                 ERR.Raise 43434, "clsPOP", "Mail server returned error code"
           End If
          
     End With
End Sub
 
 
 
 
 
 
 

' _ __     .    . _ __     .    .     . _     .    .
'| '_ \ _ _  ___ | '_ \ ___  _ _ _____ (_) ___  _ _'
'| |_) | '_\/ _ \| |_) / _ \| '_\_   _|| |/ _ \/ __|
'| .__/| |   (_) | .__/  __/| |   | |  | |  __/\__ \
'|_|   |_|  \___/|_|   \___||_|   |_|  |_|\___||___/



'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: | THE USERNAME FOR THE POP EMAIL ACCOUNT
' COMMENTS: |
'----------------------------------------------------------------------
Public Property Get POPusername() As String
        
        POPusername = m_POPusername
End Property
Public Property Let POPusername(ByVal vNewValue As String)

        m_POPusername = vNewValue
End Property
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: | THE PASSWORD FOR THE POP EMAIL ACCOUNT
' COMMENTS: |
'----------------------------------------------------------------------
Public Property Get POPpassword() As String

       POPpassword = m_POPpassword
End Property
Public Property Let POPpassword(ByVal vNewValue As String)

        m_POPpassword = vNewValue
End Property
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: | THE NET ADDRESS FOR THE POP EMAIL ACCOUNT
' COMMENTS: |
'----------------------------------------------------------------------
Public Property Get POPaddress() As String

        POPaddress = m_POPaddress
End Property
Public Property Let POPaddress(ByVal vNewValue As String)

        m_POPaddress = vNewValue
End Property

'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: | THE DATA/COMMUNICATION THE POP SERVER IS SENDING TO US
' COMMENTS: | ALLOWS USER OF CLASS TO SEE THE COMMUNICATION AS WELL
'----------------------------------------------------------------------
Public Property Get POPserverCommunication() As String

        POPserverCommunication = m_POPserverCommunication
End Property
Private Property Let POPserverCommunication(ByVal vNewValue As String)

        m_POPserverCommunication = vNewValue
End Property
'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: | THE DATA/COMMUNICATION THAT WERE SENDING POP SERVER
' COMMENTS: | ALLOWS USER OF CLASS TO SEE THE COMMUNICATION AS WELL
'----------------------------------------------------------------------
Public Property Get POPclientCommunication() As String

        POPclientCommunication = m_POPclientCommunication
End Property
Private Property Let POPclientCommunication(ByVal vNewValue As String)

        m_POPclientCommunication = vNewValue
End Property







' _ __     . _      .     .     .    .  '    .     . _   .    .
'| '_ \ _ _ (_)_   __ __ _ _____  ___   ' _ _'_   _ | |__  _ _'
'| |_) | '_\| | \ / // _` |_   _|/ _ \  '/ __| | | ||  _ \/ __|
'| .__/| |  | |\ V /  (_| | | |    __/  '\__ \ |_| || |_) \__ \
'|_|   |_|  |_| \_/  \__,_| |_|  \___|  '|___/\__,_||_.__/|___/
''  _      .      .    .     . _     .      .    .
' / _|_   _  _ __   ___ _____ (_) ___  _ __   _ _'
'| |_  | | || '_ ` / __|_   _|| |/ _ \| '_ ` / __|
'|  _| |_| || | | | (__  | |  | | (_) | | | |\__ \
'|_|  \__,_||_| |_|\___| |_|  |_|\___/|_| |_||___/









'----------------------------------------------------------------------
'   INPUTS: |THE MAIL SERVER RESPONSE TO "STAT" WHICH GIVES US
'            THE EXAMPLE STRING   +OK 5 23990, WE NEED TO PARSE OUT THE 5
'  RETURNS: |NUMBER OF MESSAGES WAITING ON THE SERVER
' COMMENTS: |
'----------------------------------------------------------------------
Private Function funcMsgCount(strToParse) As Integer
 Dim sParts()           As String
 Dim i                  As Integer
 
        sParts = Split(strToParse, " ")
        funcMsgCount = Trim(sParts(1))
        
End Function


