Attribute VB_Name = "ShowWindow"
Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
'
' Constants used with APIs
'
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9
Public Const GW_OWNER = 4
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_APPWINDOW = &H40000
'
' Listbox messages
'
Public Const LB_ADDSTRING = &H180
Public Const LB_SETITEMDATA = &H19A
Public Sub pSetForegroundWindow(ByVal hwnd As Long)
Dim lForeThreadID As Long
Dim lThisThreadID As Long
Dim lReturn       As Long
'
' Make a window, specified by its handle (hwnd)
' the foreground window.
'
' If it is already the foreground window, exit.
'
If hwnd <> GetForegroundWindow() Then
    '
    ' Get the threads for this window and the foreground window.
    '
    lForeThreadID = GetWindowThreadProcessId(GetForegroundWindow, ByVal 0&)
    lThisThreadID = GetWindowThreadProcessId(hwnd, ByVal 0&)
    '
    ' By sharing input state, threads share their concept of
    ' the active window.
    '
    If lForeThreadID <> lThisThreadID Then
        ' Attach the foreground thread to this window.
        Call AttachThreadInput(lForeThreadID, lThisThreadID, True)
        ' Make this window the foreground window.
        lReturn = SetForegroundWindow(hwnd)
        ' Detach the foreground window's thread from this window.
        Call AttachThreadInput(lForeThreadID, lThisThreadID, False)
    Else
       lReturn = SetForegroundWindow(hwnd)
    End If
    '
    ' Restore this window to its normal size.
    '
    If IsIconic(hwnd) Then
       Call ShowWindow(hwnd, SW_RESTORE)
    Else
       Call ShowWindow(hwnd, SW_SHOW)
    End If
End If
End Sub

Public Function fEnumWindows(lst As ListBox) As Long
'
' Clear list, then fill it with the running
' tasks. Return the number of tasks.
'
' The EnumWindows function enumerates all top-level windows
' on the screen by passing the handle of each window, in turn,
' to an application-defined callback function. EnumWindows
' continues until the last top-level window is enumerated or
' the callback function returns FALSE.
'
With lst
    .Clear
    Call EnumWindows(AddressOf fEnumWindowsCallBack, .hwnd)
    fEnumWindows = .ListCount
End With
End Function

Private Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim lReturn     As Long
Dim lExStyle    As Long
Dim bNoOwner    As Boolean
Dim sWindowText As String
'
' This callback function is called by Windows (from
' the EnumWindows API call) for EVERY window that exists.
' It populates the listbox with a list of windows that we
' are interested in.
'
' Windows to display are those that:
'   -   are not this app's
'   -   are visible
'   -   do not have a parent
'   -   have no owner and are not Tool windows OR
'       have an owner and are App windows
'
If hwnd <> frmAltTab.hwnd Then
    If IsWindowVisible(hwnd) Then
        If GetParent(hwnd) = 0 Then
            bNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
            
            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
                ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
                '
                ' Get the window's caption.
                '
                sWindowText = Space$(256)
                lReturn = GetWindowText(hwnd, sWindowText, Len(sWindowText))
                If lReturn Then
                   '
                   ' Add it to our list.
                   '
                   sWindowText = Left$(sWindowText, lReturn)
                   lReturn = SendMessage(lParam, LB_ADDSTRING, 0, ByVal sWindowText)
                   Call SendMessage(lParam, LB_SETITEMDATA, lReturn, ByVal hwnd)
                End If
            End If
        End If
    End If
End If
fEnumWindowsCallBack = True
End Function

