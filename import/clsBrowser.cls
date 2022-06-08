VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
' Automating Chromium-Based Browsers with Chrome Dev Protocol API and VBA
'------------------------------------------------------------------------------------------------
' Author(s)   :
'       ChrisK23 (Code Project)
' Contributors:
'       Long Vh (long.hoang.vu@hsbc.com.sg)
' Last Update :
'       07/06/22 Long Vh: - adjusted .wait method + added "restore sessions" popup-handling.
'                         - converted .attach to Private function. Added .attach to .start.
'                         - removed blankStart argument from .start. Made blank-start default.
'                         - added userDir argument to .start for ease of multi-session init.
'                         - added InvokeMethod for advanced automation needs on browser object.
'                         - enhanced .show method with new arguments for windows repositioning.
'                         - added multi-tabs automation methods for working with tabs.
'                         - added .maximized and .minimized to the native methods of clsBrowser.
'       03/06/22 Long Vh: - added extensive commentations + extended codes for HSBC usages
' References  :
'       Microsoft Scripting Runtime
' Sources     :
'       CDP Tutorials: https://github.com/cyrus-and/chrome-remote-interface/issues/282
'       CodeProject: https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA
'       Javascript Tutorials: https://www.w3schools.com/js/
'       Chrome CDP Documentation: https://chromedevtools.github.io/devtools-protocol/
'================================================================================================

'===================================
' Win APIs Declarations
'===================================

Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" ( _
    ByVal dwMilliseconds As Long)
    
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As Long, _
    ByVal hWndChildAfter As Long, _
    ByVal lpszWindowClass As String, _
    ByVal lpszWindowTitle As String) As Long
    
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long
    
Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Boolean
    
'===================================
' Constants
'===================================

Private Enum SW
    SW_HIDE = 0            'Hides the window and activates another window.
    SW_NORMAL = 1          'Activates and displays a window. If the window is minimized or maximized, the system restores it to its original
    SW_SHOWNORMAL = 1      'Resize and position. An application should specify this flag when displaying the window for the first time.
    SW_SHOWMINIMIZED = 2   'Activates the window and displays it as a minimized window.
    SW_SHOWMAXIMIZED = 3   'Activates the window and displays it as a maximized window.
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4  'Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except that the window is not activated.
    SW_SHOW = 5            'Activates the window and displays it in its current size and position.
    SW_MINIMIZE = 6        'Minimizes the specified window and activates the next top-level window in the Z order.
    SW_SHOWMINNOACTIVE = 7 'Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated.
    SW_SHOWNA = 8          'Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated.
    SW_RESTORE = 9         'Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    SW_SHOWDEFAULT = 10    'Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application.
    SW_FORCEMINIMIZE = 11  'Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread.
End Enum
 
Private Const chrome As String = "chrome.exe"
Private Const edge As String = "msedge.exe"
Private Const chromeWindowClass As String = "Chrome_WidgetWin_1" ' same window class for Edge
Private Const chromeWindowTitle As String = " - Google Chrome"
Private Const edgeWindowTitle As String = " - Microsoft? Edge" ' ? is a zero width space "&H200B" if comparing as Unicode

'===================================
' Shared Variables
'===================================

Private objCore As clsCore
Private lngLastID As Long           'Every message sent to edge has an ID
Private strBuffer As String         'Buffer string, required for various win APIs
Private brHWnd As Long              'Browser object window handle
Private brWId As String             'Useful for windows repositioning
Private brType As String            'Browser type: either 'chrome' or 'edge'
Private brUser As String            'Browser user profile
Public SessionID As String          'The ID to the current active tab.
 
 
'=============================================================================
' PUBLIC FUNCTIONS
'=============================================================================

 
Public Function start(Optional name = "chrome", Optional cleanActiveSession As Boolean = True, Optional userProfile As String = "Default", Optional addArguments As String = "")
'-------------------------------------------------------------------------------------------------------------
' This function initiates a new browser instance.
' By default, this function initializes chrome as the default choice.
' To init edge, parse 'edge' to the first parameter (eg. browser.start "edge").
' If cleanActiveSession is False, the method will skip cleaning currently open instances. This will save
' initialization time but if an instance is already running and not closed, pipe error will be triggered.
' To add additional arguments to the call string, use the addArguments parameter for instance:
' .start addArguments:="--disable-info-bar" (https://peter.sh/experiments/chromium-command-line-switches/)
' Note: not all arguments work the same for Chrome and Edge. Some may work for one but not the others
'-------------------------------------------------------------------------------------------------------------
    
    Dim strCall As String
    Dim intRes As Integer
                               
    timeStart = Timer
                               
    Select Case LCase(name)
        Case "edge"
            brType = edge
            strCall = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"" --remote-debugging-pipe --enable-automation --homepage ""about:blank"""
        Case "chrome"
            brType = chrome
            strCall = """C:\Program Files\Google\Chrome\Application\chrome.exe"" --remote-debugging-pipe --enable-automation --homepage ""about:blank"""
        Case Else
            err.Raise -900, Description:="Browser initiation fails: only 'edge' or 'chrome' should be indicated as browser name."
    End Select
    
   'Start with a choice profile
    If userProfile <> "Default" Then
        brUser = userProfile
        Select Case brType
            Case edge: addProfile = " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Microsoft\Edge\User Data\" & userProfile & """"
            Case chrome: addProfile = " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Google\Chrome\User Data\" & userProfile & """"
        End Select
        strCall = strCall & addProfile
    End If
    
   'Add additional arguments if instructed
    If addArguments <> "" Then strCall = strCall & " " & addArguments
    
   'Clean up all running browser instances to prevent pipe error
    If cleanActiveSession Then cleanUpSessions
    
   'Prevent "Restore Session" popup
    setCrashStateNormal
    
    Set objCore = New clsCore
    intRes = objCore.init(strCall, aboolSerializable:=True)  'Serializable is set to True by default for ease of accessibility
    
    If intRes <> 0 Then err.Raise -99, Description:="Unable to start browser."
    lngLastID = 1
    
   'Attach & finish initialization
    attach
    Debug.Print "Browser initialized after " & Format(Timer - timeStart, "0.000") & "ms"
    
End Function


Public Function serialize() As String
'------------------------------------------------------------------
' This method transcribes Chrome's Json message into a string.
' The serialized string can then be used to automate an active
' session with a known serial.
'------------------------------------------------------------------
    
    Dim objSerialize As New Dictionary
    
    objSerialize.Add "objCore", objCore.serialize()
    objSerialize.Add "lngLastID", lngLastID
    objSerialize.Add "SessionID", SessionID
    
    serialize = objCore.jsConverter.ConvertToJson(objSerialize)
    Debug.Print "Serialized Json: " & serialize
    
End Function

 
Public Function deserialize(strSerialized As String)
'------------------------------------------------------------------
' This method deciphers a serialized string to obtain the currently
' active session ID of the already open instance. The serial
' must be obtained first from the serialize method.
'------------------------------------------------------------------
    
    Dim objSerialize As Dictionary
    Set objCore = New clsCore
    
    Set objSerialize = objCore.jsConverter.ParseJson(strSerialized)
    objCore.deserialize objSerialize.Item("objCore")
    
    lngLastID = objSerialize.Item("lngLastID")
    SessionID = objSerialize.Item("SessionID")
    Debug.Print "Reattached to Session ID: " & SessionID
    
End Function
 
 
Public Function jsEval(strString As String, Optional dbgMsg As Boolean = True) As Variant
'------------------------------------------------------------
' This function commands the attached browser instance to
' execute a Javascript expression to perform required tasks.
' Note: the original function has a bool retry variable which
' was removed for simplicity.
' Added dbgMsg to read debug line on the go when needed.
'------------------------------------------------------------
    
    Dim results As Dictionary
    Dim strMessage As String
    Dim objMessage As Dictionary
    
    strMessage = "{""method"":""Runtime.evaluate"",""params"":{""expression"":""1+1;""}}"
    Set objMessage = objCore.jsConverter.ParseJson(strMessage)
    objMessage.Item("params").Item("expression") = strString & ";"
    strMessage = objCore.jsConverter.ConvertToJson(objMessage)
    
    Set results = sendMessage(strMessage)
    If results Is Nothing And boolRetry Then err.Raise -900, Description:="Unable to contact the Browser instance."
    If (results.Exists("error")) Then err.Raise -900, Description:="The Browser was unable to evaluate the Js expression: " & strString
    
   'If the return type has a specific type we can return the result
    With results.Item("result").Item("result")
        If .Item("type") = "string" Or _
           .Item("type") = "boolean" Or _
           .Item("type") = "number" Then _
                  jsEval = .Item("value")
    End With
    
    If dbgMsg Then Debug.Print "Script executed: " & strString & " | Returned: " & jsEval
    
End Function


Public Function invokeMethod(methodName As String, Optional params As Dictionary) As Dictionary
'-----------------------------------------------------------------
' This function enables advanced method invocation under CDP.
' For a list of all methods available, see:
' https://chromedevtools.github.io/devtools-protocol/tot/Browser/
' Error message will be prompted automatically by invokeError in
' the Immediate Window.
' The return object is Json String that has been converted to a
' dictionary object.
' Credit: https://github.com/PerditionC/VBAChromeDevProtocol
'-----------------------------------------------------------------
    
    Dim paramStr As String
    Dim results As Dictionary
    
    If params Is Nothing Then paramStr = "{}" Else paramStr = objCore.jsConverter.ConvertToJson(params)
    Set results = sendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}")
    Debug.Print "Invoked " & methodName
    
   'Print error if there is
    If invokeError(results) Then
        Debug.Print "Failed to invoke method: " & methodName
        err.Raise -100, Description:="Failed to invoke method: " & methodName
    End If
    
   'Return result object
    If results.Exists("result") Then Set invokeMethod = results.Item("result")
    
End Function


Public Function navigate(strUrl As String)
'-----------------------------------------------------------------------
' This function is a quick wraper to navigate to an url
'-----------------------------------------------------------------------
    
    Dim results As Dictionary
    Set results = sendMessage("{""method"":""Page.navigate"",""params"":{""url"":""" & strUrl & """}}")
    Debug.Print "Navigated to " & strUrl
    
End Function

 
Public Function quit()
'-----------------------------------------------------------------------------------
' Closes the currently attached browser instance
' Original Name: closeBrowser
'-----------------------------------------------------------------------------------
    
    Dim results As Dictionary
    
    On Error Resume Next
    Set results = sendMessage("{""method"":""Browser.close"",""params"":{}}")
 
    sleep 1   'It seems without waitng a bit the browser crashes and the next time it wants ro recover from a crash
    Debug.Print "Browser was terminated."
    
End Function

 
Public Function isLive() As Boolean
'-----------------------------------------------------------------------------------
' Check if there is ongoing connection with an attached instance
' Original Name: connectionAlive
'-----------------------------------------------------------------------------------
    
    On Error GoTo ifErr
    Dim strLoc As String
    strLoc = jsEval("window.location.href")
    isLive = True
    Exit Function
    
ifErr:

    isLive = False
    
End Function
 
 
Public Function wait(Optional till As String = "complete", Optional dbgState As Boolean = False)
'--------------------------------------------------------------------------------------------------
' Wait until browser reaches a specific state. Default is until it fully loads ("complete")
' Possible states:  loading / interactive / complete
' Complete state is preferable when certain information is only showed when fully loaded.
' Otherwise, an interactive state is good enough and can potentially save a lot of wait time.
' To read ReadyState on the go, set the dbgState to True to print the state to the Immediate Window
' Original Name: waitCompletion
'--------------------------------------------------------------------------------------------------
    
    Do Until strState = till Or strState = "complete"
        strState = jsEval("document.readyState", dbgMsg:=False)         ' dbgMsg = False to not spam Immediate Window
        If dbgState Then Debug.Print "Current ReadyState: " & strState  ' option to print ReadyState info during iteration
        sleep 0.1                                                       ' reduce sleep will speed up but will cost cpu power
    Loop
    
End Function


Public Function maximized()
'-----------------------------------------------
' Maximize the browser window
'-----------------------------------------------

    If Not brHWnd <> 0 Then brHWnd = getHandle
    ShowWindow brHWnd, SW.SW_RESTORE
    ShowWindow brHWnd, SW.SW_MAXIMIZE
    sleep 0.3

End Function


Public Function minimized()
'-----------------------------------------------
' Minimize the browser window
'-----------------------------------------------

    If Not brHWnd <> 0 Then brHWnd = getHandle
    ShowWindow brHWnd, SW.SW_RESTORE
    ShowWindow brHWnd, SW.SW_MINIMIZE
    sleep 0.3

End Function


Public Function hide()
'-----------------------------------------------------------------------------------
' Make the browser window entirely invisible. Use .show to make it visible again.
'-----------------------------------------------------------------------------------
    
    If Not brHWnd <> 0 Then brHWnd = getHandle
    ShowWindow brHWnd, SW.SW_HIDE
    
End Function

 
Public Function show(Optional left As Variant, Optional top As Variant, Optional width As Variant, Optional height As Variant)
'-----------------------------------------------------------------------------------
' Restore browser visibility after being made invisible with .hide method.
' Can also use after .attach to bring the window topmost as well. By default at
' start, the browser window is minimized.
'-----------------------------------------------------------------------------------
    
    If Not brHWnd <> 0 Then brHWnd = getHandle
    ShowWindow brHWnd, SW.SW_RESTORE                    'Make visible
    ShowWindow brHWnd, SW.SW_SHOWNORMAL
    
   'Resize to specific position and size
   'Will not work properly if window is currently maximized
    If Not (IsMissing(left) And IsMissing(top) And IsMissing(width) And IsMissing(height)) Then 'Note: IsMissing only works with Variant type
        
        sleep 0.3      'Sleep a bit for previous Win Api executions to finish, else the resize effect will be accidentally cancelled out
        
        Dim params As New Dictionary
        If brWId = "" Then brWId = invokeMethod("Browser.getWindowForTarget").Item("windowId")
        params("windowId") = CInt(brWId)
                
        Dim bounds As New Dictionary
        If Not IsMissing(left) Then bounds("left") = CLng(left)
        If Not IsMissing(top) Then bounds("top") = CLng(top)
        If Not IsMissing(width) Then bounds("width") = CLng(width)
        If Not IsMissing(height) Then bounds("height") = CLng(height)
        Set params("bounds") = bounds

        invokeMethod "Browser.setWindowBounds", params
        
    End If
    
End Function


Public Function newTab(Optional url As String, Optional newWindow As Boolean) As String
'------------------------------------------------------------------
' Create a new blank tab within the same instance.
' url: The initial URL the page will be navigated to. An empty
' string indicates about:blank.
' newWindow: If true will open a separate window instead of a tab.
' In the original CDP document, each new tab is called "Target" and
' the equivalence of this method is Target.createTarget
' This returns the Session ID of the newly created tab.
'------------------------------------------------------------------

    Dim params As New Dictionary
    params("url") = CStr(url)
    If Not IsMissing(newWindow) Then params("newWindow") = CBool(newWindow)
    
    Dim results As Dictionary
    Set results = invokeMethod("Target.createTarget", params)
       
   'Return the new tab's Session ID
    Dim tabId As String
    tabId = results("targetId")
    newTab = attachToTab(tabId)
    
End Function


Public Function closeTab()
'------------------------------------------------------------------
' Close the current tab of the session. Requires targetId which can
' be retrieved by using getTargetId. This is the targetId of the
' active tab under focus of the automation pipe.
' WARNING: DON'T close the main tab (the tab window that is open
' with .start method) as this tab is needed to keep the pipe
' running. Closing this will cause pipe message automation to fail.
' This method is only for closing tabs that are subsequently open
' during the automation.
'------------------------------------------------------------------
    
    Dim params As New Dictionary
    params("targetId") = CStr(getTargetId)
    invokeMethod "Target.closeTarget", params

End Function


Public Function getNewTab() As String
'------------------------------------------------------------------
' Search and find the Session ID of the new tab open by interaction
' with the webpage and not open by newTab.
' This is useful when clicking a link in a webpage opens a new tab
' that is not yet attached to the automation pipe and needs to be
' attached for further processing on the new tab.
'------------------------------------------------------------------

    Dim params As New Dictionary
    Dim results As Dictionary
    Set results = invokeMethod("Target.getTargets", params)
    Set targets = results("targetInfos")
    
   'Look for the new tab Id
    Dim tabId As String
    For Each Target In targets
        If Target("type") = "page" And Target("attached") = False Then
            tabId = Target("targetId")
            Exit For
        End If
    Next
    
   'Return the new tab's Session ID
    getNewTab = attachToTab(tabId)

End Function


Public Function switchTo(SessionID_ As String)
' ------------------------------------------------------------------------
' Switch automation object to the specified Session ID. Used for handling
' multiple tabs of the same instance. Use newTab to create and get the
' Session ID of the new tab.
' ------------------------------------------------------------------------
            
    SessionID = SessionID_
        
End Function


Public Function clone() As clsBrowser
' ------------------------------------------------------------------------
' Clone the parameters of the current instance to be parsed to a new
' instance object. This is useful when it is needed to handle multiple
' with different Session IDs but the same remaining pipe parameters.
' ------------------------------------------------------------------------

    Dim clone_ As New clsBrowser
    clone_.deserialize serialize     'Copy the current session's serial
    Set clone = clone_               'Return the new clsBrowser object

End Function


Public Function checkTabsInfo() As String
'------------------------------------------------------------------
' Get all current tabs info. Useful for debugging.
'------------------------------------------------------------------

    Dim params As New Dictionary
    Dim results As Dictionary
    Set results = invokeMethod("Target.getTargets", params)
    Set targets = results("targetInfos")
    
   'Read info
    Debug.Print "---------------------------------------------------"
    Debug.Print " Tab Verification"
    Debug.Print "---------------------------------------------------"
    i = 1
    For Each target_ In targets
        If target_("type") = "page" Then
            Debug.Print "#" & i, _
                        "type: " & target_("type"), _
                        "attached: " & target_("attached"), _
                        "title: " & target_("title"), _
                        "targetId: " & target_("targetId"), _
                        "url: " & target_("url")
        End If
        i = i + 1
    Next

End Function


'=============================================================================
' PRIVATE FUNCTIONS
'=============================================================================


Private Function sendMessage(strMessage As String, Optional objAllMessages As Dictionary) As Dictionary
'------------------------------------------------------------
' Send message to CDP and check if the message is received
' yet by the Chromium browser.
'------------------------------------------------------------
    
    Dim intRes As Long
    Dim strRes As String
    Dim lngCurrentId As Long
    Dim results As Dictionary    'to store a collection of IDs obtained from CDP
    Dim intCounter As Integer: intCounter = 0
    
    lngCurrentId = lngLastID
    lngLastID = lngLastID + 1   'We increase the global ID counter
    
    ' Before sending a message the messagebuffer is emptied
    ' All messages that we have received so far cannot be an answer to the message that we will send
    ' So they can be safely discarded
    
    If objAllMessages Is Nothing Then
    
        intRes = 1
        Do Until intRes < 1
            intRes = objCore.readProcCDP(strRes)
            If intRes > 0 Then strBuffer = strBuffer & strRes
        Loop
        
        Dim lngNullCharPos As Long
        lngNullCharPos = searchNull()
        
        Do Until lngNullCharPos = 0
            'Debug.Print (Left(strBuffer, lngNullCharPos))
            strBuffer = Right(strBuffer, Len(strBuffer) - lngNullCharPos)
            lngNullCharPos = searchNull()
        Loop
        
    End If
    
    intRes = objCore.readProcSTD(strRes)             'Sometimes edge writes to stdout so we clear stdout too.
    strMessage = left(strMessage, Len(strMessage) - 1)  'We add the currentID and sessionID to the message
    
    If SessionID <> "" Then strMessage = strMessage & ", ""sessionId"":""" & SessionID & """"
    strMessage = strMessage & ", ""id"":" & lngCurrentId & "}" & vbNullChar
    
    objCore.writeProc strMessage                     'Send message to the browser instance
    
    Do Until intCounter > 100                           'Add failsafe counter timeout in order not to loop forever for the returning message
        intRes = 1
        Do Until intRes < 1                             'We read from edge and process messages until we receive a message with our ID
            intRes = objCore.readProcCDP(strRes)
            If intRes > 0 Then strBuffer = strBuffer & strRes
        Loop
        
        lngNullCharPos = searchNull()
            
        Do Until lngNullCharPos = 0
           strRes = left(strBuffer, lngNullCharPos - 1)
            strBuffer = Right(strBuffer, Len(strBuffer) - lngNullCharPos)
            
            'Debug.Print (strRes)
            If strRes <> "" Then
                Set results = objCore.jsConverter.ParseJson(strRes)
                If Not objAllMessages Is Nothing Then objAllMessages.Add CStr(objAllMessages.Count), results
                If results.Exists("id") Then
                    If results.Item("id") = lngCurrentId Then
                        Set sendMessage = results
                        Exit Function
                    End If
                End If
            End If
            lngNullCharPos = searchNull()
        Loop
        
        DoEvents
        sleep 0.1
        intCounter = intCounter + 1
    Loop
    
ifTimeOut:

    err.Raise Number:=-900, Description:="Timeout from the sendMessage function. The strMessage is: " & strMessage
    
End Function

 
Private Function attach(Optional strUrl As String = "") As Integer
'-----------------------------------------------------------------------
' This function must be called after start and before all other methods.
' This function attaches to a session of the browser.
' By default, if strUrl is not indicated, this function will try to
' attach to the instance that is opening a blank tab.
'-----------------------------------------------------------------------
    
    Dim results As Dictionary
    Dim objDic2 As Dictionary
    Dim objDic3 As Dictionary
    Dim objAllMessages As New Dictionary
    Dim i As Integer
    Dim boolFound As Boolean
    Dim strKey As Variant
            
    timeStart = Timer
            
    Set results = sendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":true}}", objAllMessages)
    
   'Read message to confirm the requested url is found ----------------------------
    For Each strKey In objAllMessages.Keys
        Set results = objAllMessages.Item(strKey)
        If results.Exists("params") Then
            Set objDic2 = results.Item("params")
            If objDic2.Exists("targetInfo") Then
                Set objDic3 = objDic2.Item("targetInfo")
                If objDic3.Item("type") = "page" Then
                    If objDic3.Item("url") = strUrl Or strUrl = "" Then
                        boolFound = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next strKey
   '-------------------------------------------------------------------------------
    
    If Not boolFound Then
        attach = -1 'ie. not found
        Exit Function
    End If
    
   'If the instance is found
    Set results = sendMessage("{""method"":""Target.attachToTarget"",""params"":{""targetId"":""" & objDic3.Item("targetId") & """,""flatten"":true}}")
    SessionID = results.Item("result").Item("sessionId")
    Set results = sendMessage("{""method"":""Runtime.enable"",""params"":{}}")
    Set results = sendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":false}}")
    attach = 0
    
    timeEnd = Timer
    Debug.Print "Attached successfully after " & Format(timeEnd - timeStart, "0.000") & "ms"
    
End Function


Private Function searchNull() As Long
'------------------------------------------------------------
' CDP messages received from chrome are null-terminated
' It seemed to me you cant search for vbnull in a string
' in vba. Thats why i re-implemented the search function
'------------------------------------------------------------
    
    Dim i As Long
    Dim lngBufferLength As Long
    
    lngBufferLength = Len(strBuffer)
    searchNull = 0
    
    If lngBufferLength = 0 Then Exit Function
    
    For i = 1 To lngBufferLength
        If Mid(strBuffer, i, 1) = vbNullChar Then
            searchNull = i
            Exit Function
        End If
    Next i
    
End Function


Private Function invokeError(result As Dictionary) As Boolean
'-----------------------------------------------------------------------------------
' Check if InvokeMethod fails and what the error message returned is.
' Error message will be shown via Debug.Print
' Credit: https://github.com/PerditionC/VBAChromeDevProtocol
'-----------------------------------------------------------------------------------

    If result Is Nothing Then
        ErrorCode = -1 'arbitrary nonzero value, may change in future
        ErrorMessage = "Timeout or No results"
        ErrorData = vbEmpty
        Debug.Print "InvokeMethod Error: [" & ErrorCode & "] " & ErrorMessage
        invokeError = True
    ElseIf result.Exists("error") Then
        Dim errorDetails As Dictionary
        Set errorDetails = result("error")
        ErrorCode = errorDetails("code")
        ErrorMessage = errorDetails("message")
        Debug.Print "InvokeMethod Error: [" & ErrorCode & "] " & ErrorMessage
        If errorDetails.Exists("data") Then
            ErrorData = errorDetails("data")
            Debug.Print "       " & ErrorData
        Else
            ErrorData = vbEmpty ' vbNullString? is errorDetails("data") a string or object?
        End If
        invokeError = True
    Else
        ErrorCode = 0
        ErrorMessage = vbNullString
        ErrorData = vbEmpty
    End If
    
End Function

 
Private Function getHandle() As Long
'-----------------------------------------------------------------------------------
' Added by Long: obtain the window handle of the current browser session
'-----------------------------------------------------------------------------------
    
    Dim bufferString As String
    Dim titleLen As Long
    
    sessionTitle = jsEval("document.getElementsByTagName(""title"")[0].innerHTML")
    
   'Obtain handles to the correct browser window
   'Both Edge and Chrome use the same class chromeWindowClass
    hwnd = FindWindowEx(0&, 0&, chromeWindowClass, vbNullString)        'To return Chrome / Edge window handles
    Do
                    
       'Get Window Title based on specified handles
       'Requires string buffering declaration as string
        bufferString = String(1024, " ")
        titleLen = GetWindowText(hwnd, bufferString, Len(bufferString))
        windowTitle = left(bufferString, titleLen)
        
       'Verify if matching with the sessionTitle and is really an edge/chrome window
        If windowTitle <> "" Then _
            If InStr(Split(windowTitle, "-")(0), sessionTitle) > 0 Then _
                If InStr(windowTitle, edgeWindowTitle) Or _
                   InStr(windowTitle, chromeWindowTitle) Then Exit Do
        
        hwnd = FindWindowEx(0&, hwnd, chromeWindowClass, vbNullString)  'Continue with the next handle value
        
    Loop Until hwnd = 0 ' no more matching windows
    
    getHandle = hwnd
    
End Function
 
 
Private Function sleep(Optional second As Double = 0.5)
'------------------------------------------------------------------
' Custom sleep function. Sleep by 1s by default
' Change sleep period if processing is not robust
'------------------------------------------------------------------
    
    Const sleepPeriod As Long = 1000    'millisecs
    
    sleep2 sleepPeriod * second
    DoEvents
    
End Function
 
 
Private Function cleanUpSessions()
'------------------------------------------------------------------
' Provides cleaning before firing up a new session to prevent
' pipe error
'------------------------------------------------------------------
    
   'Clean up all running sessions
    On Error Resume Next
    For Each proc In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
        If proc.name = brType Then
            proc.Terminate
        End If
        DoEvents
    Next
    On Error GoTo 0
 
End Function


Private Function setCrashStateNormal()
'------------------------------------------------------------------
' Set Crashed state to Normal to prevent the "Restore Sessions?"
' popup from appearing every time the browser is started after
' being closed by browser.quit method.
' Requires user data dir profile stored in private var brUser and
' the browser type stored in private variable brType.
'------------------------------------------------------------------

   'Mark 'Normal' on the preference file to prevent Restore Page popup on start
   'This file locates at %localappdata%/Google/User Data/Default/Preferences
    If brUser = "Default" Then prefFolder = brUser Else prefFolder = brUser & "\Default"
    Select Case brType
        Case chrome: filePath = Environ("LOCALAPPDATA") & "\Google\Chrome\User Data\" & prefFolder & "\Preferences"
        Case edge: filePath = Environ("LOCALAPPDATA") & "\Microsoft\Edge\User Data\" & prefFolder & "\Preferences"
    End Select
    
    Dim fso As New Scripting.FileSystemObject
    If Not fso.FileExists(filePath) Then Exit Function
    Set prefFile = fso.OpenTextFile(filePath, ForReading)
    strContents = prefFile.ReadAll
    strNewContents = Replace(strContents, "Crashed", "Normal", Compare:=vbBinaryCompare)   'ie. Exact match only
    prefFile.Close
    
    Set prefFile = fso.OpenTextFile(filePath, ForWriting)
    prefFile.Write strNewContents
    prefFile.Close

End Function


Private Function attachToTab(tabId As String) As String
'----------------------------------------------------------------------------------
' Attach to a specific tab with its tabId (a.k.a targetId)
' flatten: Enables "flat" access to the session via specifying sessionId attribute
' in the commands. Set to True by default to make attachToTab work.
' Returns: a session ID string that is needed for CDP automation.
'----------------------------------------------------------------------------------
    
    Dim params As New Dictionary
    params("targetId") = CStr(tabId)
    params("flatten") = True

    Dim results As Dictionary
    Set results = invokeMethod("Target.attachToTarget", params)
    attachToTab = results("sessionId")
    
End Function


Private Function getTargetId() As String
'----------------------------------------------------------------------------------
' Retrieve the targetId of the current tab (a.k.k "target" in CDP Documentation).
' This targetId is useful for tab automation, such as the .closeTab method.
'----------------------------------------------------------------------------------

    Dim results As Dictionary
    Set results = invokeMethod("Target.getTargetInfo")
    getTargetId = results("targetInfo").Item("targetId")
    
End Function


'===================================
' Experimental Functions
'===================================
