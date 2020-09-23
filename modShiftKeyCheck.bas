Attribute VB_Name = "modShiftKeyCheck"
Option Explicit



'----------------------------------------------------------------------
'   INPUTS: |
'  RETURNS: |
' COMMENTS: |check to see if shift key is being pressed
'            If shift is being presses while an api menu
'            item is being clicked on, then the types or
'            constants related to that api are shown.
'            If shift is not present, then the API is sent
'----------------------------------------------------------------------
Function funcCheckForShift() As Boolean
        '
        If GetKeyState(VK_SHIFT) = -127 Or _
           GetKeyState(VK_SHIFT) = -128 Then
           
               funcCheckForShift = True
        Else
               funcCheckForShift = False
        End If
End Function
 
'-----------------------------------------------------
'if user is pressing p while clicking a menu
'api call..it means he wants to add Private
'to the api declaration
'-----------------------------------------------------
Function funcCheckForControl() As Boolean
        '
        If GetKeyState(VK_CONTROL) = -127 Or _
           GetKeyState(VK_CONTROL) = -128 Then
           
               funcCheckForControl = True
        Else
               funcCheckForControl = False
        End If
End Function
