'================================================================================================
' Automating Chromium-Based Browsers with Chrome Dev Protocol API and VBA
'------------------------------------------------------------------------------------------------
' Author(s)   :
'       ChrisK23 (Code Project)
' Contributors:
'       Long Vh (long.hoang.vu@hsbc.com.sg)
'       Daniel Polak (SYS, Supporting Your Systems B.V.)
' Last Update :
'       22/01/26 Long Vh:      - some tweaks to the .start method to improve the usage of user-data-dir switch
'                              - enhanced the framework with a new feature called multi-instances operation (see demo example demoMultiProfileOperation)
'                              - removed auto reattach feature from the framework and added .reattach method to suit
'                                the new multi-instances operation feature
'       31/12/25 Daniel Polak: - added BringWindowToTop and bringtoforeground
'       24/10/25 Daniel Polak: - sendMessage: added soft re-attach attempt when pipe error
'                              - snapPage: added CDP native screenshot capability which avoids use of external JavaScript library
'                              - getTab: improved tab selection
'                              - quit: avoid immediate hard kills, reduces race conditions, and makes shutdown behavior predictable and robust in non-responsive cases
'                              - searchNull: new faster version
'                              - use constants (ERR_TIMEOUT, ERR_PIPE, ERR_PROTOCOL, TIMEOUT_DEFAULT, TIMEOUT_NAVIGATE) instead of numbers for clarity
'       23/10/25 Daniel Polak: - replaced unnecessary use of FileSystemObject
'                              - added Option Explicit and added all required variable declarations
'                              - quit: delete the one and only cdpid.txt
'                              - Class_Initialize: prevent compilation errors
'                              - cleanUpSessions: solved bug
'                              - added DeleteCDPID so we use a function instead of repeating same code
'                              - 64-bit safety: window handles as LongPtr
'                              - start: added missing space before --homepage
'                              - jsAddLib: fixed bug in check whether JS library is loaded
'                              - getTab: fixed typo timeStart instead of timerStart and fixed bug 60000s instead of 60s
'                              - cleanUpSessions: only terminate automation-processes
'                              - snapPage: fixed bug where the all files in the Downloads folder were deleted instead of just the temporary screenshot
'                              - getBrowserPath: added fallback path when registry lookup fails
'       18/08/23 Long Vh:      - set core to Private i/o Public as it is meant to be Private.
'       06/06/23 Long Vh:      - applied recommendations by MitsOmitsaras for Access & Word compat.
'       26/05/23 Long Vh:      - made some tweaks for compatibility with Access & Word out-of-the-box
'       26/04/23 Long Vh:      - added .setMain for better tab management.
'       22/03/23 Long Vh:      - added many .getElementsBy methods.
'                              - added advanced methods to extend features of CDP apps.
'                              - changed deserialize/serialize methods to read from a temp file.
'                              - added .snapPage method for capturing an entire web screen.
'                              - added interactions with the new element class object.
'                              - many other changes to improve debug handlings.
' References  :
'       Microsoft Scripting Runtime
' Sources     :
'       Main Git: https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA
'       CDP Tutorials: https://github.com/cyrus-and/chrome-remote-interface/issues/282
'       CodeProject: https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA
'       Javascript Tutorials: https://www.w3schools.com/js/
'       Chrome CDP Documentation: https://chromedevtools.github.io/devtools-protocol/
'================================================================================================

Option Explicit

Private Const CDPVer As String = "3.1.0"

'===================================
' Win APIs Declarations
'===================================

Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" ( _
    ByVal dwMilliseconds As Long)

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As LongPtr, _
    ByVal hWndChildAfter As LongPtr, _
    ByVal lpszWindowClass As String, _
    ByVal lpszWindowTitle As String) As LongPtr

Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hWnd As LongPtr, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long

Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nCmdShow As Long) As Boolean

Private Declare PtrSafe Function BringWindowToTop Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

'DJP 20241105 to solve problems with terminating sessions on terminal server
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" ( _
    ) As Long

Private Declare PtrSafe Function ProcessIdToSessionId Lib "kernel32" ( _
    ByVal dwProcessId As Long, _
    ByRef pSessionId As Long) As Long

'===================================
' Constants
'===================================

Public Enum WinState
    asHidden = 0            'Hides the window and activates another window.
    asNormal = 1            'Activates and displays a window. If the window is minimized or maximized, the system restores it to its original
    asMinimized = 2         'Activates the window and displays it as a minimized window.
    asMaximized = 3         'Activates the window and displays it as a maximized window.
    doShowNoActivate = 4    'Displays a window in its most recent size and position. This value is similar to asNormal, except that the window is not activated.
    doShow = 5              'Activates the window and displays it in its current size and position.
    doMinimize = 6          'Minimizes the specified window and activates the next top-level window in the Z order.
    doShowMinNoActivate = 7 'Displays the window as a minimized window. This value is similar to asMinimized, except the window is not activated.
    doShowNA = 8            'Displays the window in its current size and position. This value is similar to doShow, except that the window is not activated.
    doRestore = 9           'Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    doShowDefault = 10      'Sets the show state based on the WS_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application.
    doForceMin = 11         'Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread.
End Enum

Public Enum ReadyState      'Used for .wait method
    isLoading = 0           'equivalence of the browser's "loading" state
    isInteractive = 1       'equivalence of the browser's "interactive" state
    isComplete = 2          'equivalence of the browser's "complete" state
End Enum

Private Const brChrome As String = "chrome.exe"
Private Const brEdge As String = "msedge.exe"
Private Const chromeWindowClass As String = "Chrome_WidgetWin_1"    'same window class for Edge
Private Const seeRawSendMsgDbg As Boolean = False                   'set to True if needing to see the raw json msg received from the browser for debugging
Private Const ERR_TIMEOUT As Long = 900
Private Const ERR_PIPE As Long = 901
Private Const ERR_PROTOCOL As Long = 902
Private Const TIMEOUT_DEFAULT As Single = 30
Private Const TIMEOUT_NAVIGATE As Single = 45

'===================================
' Global Variables
'===================================

Private core As CDPCore                 'Must be initialized for CDP engine
Private lngLastID As Long               'Every message sent to edge has an ID
Private strBuffer As String             'Buffer string, required for various win APIs
Private brHWnd As LongPtr               'Browser object window handle
Private brWID As String                 'Useful for windows repositioning
Private brType As String                'Browser type: either 'chrome' or 'edge'
Private brUser As String                'Browser user profile
Private brPath As String                'Browser default installation path
Private brID As String                  'The ID of the current CDPBrowser object. Useful when reading debugs of multiple tabs.
Private brName As String                'The focused tab name / window name of the active CDPBrowser
Private brURL As String                 'The focused tab url of the active CDPBrowser
Private brVer As String                 'The current browser version
Private brReAttach As Boolean           'Remember reAttach preference from .reattach method - updated by Long 22/01/2026
Private brCallString As String          'The initial call string used by .start
Private varIDs As Scripting.Dictionary  'Collection of all unique element IDs for CDPElement reference - Added 02/03/2023
Private sessionID As String             'The ID to the current session tab.
Private targetID As String              'The targetID of the current tab for InvokeMethod
Public thisApp As Object                'A dymanic variable used in CDPBrowser initialize procedure to determine if the VBA host is an Excel / Word / Access

'=============================================================================
' PUBLIC FUNCTIONS
'=============================================================================

Public Function start(Optional name As String = "chrome", _
    Optional appUrl As String, _
    Optional cleanActive As Boolean = True, _
    Optional userProfile As String = "CDP1", Optional addArgs As String)
'--------------------------------------------------------------------------------------------------------------
' This function initiates a new browser instance.
' By default, this function initializes chrome as the default choice.
' To init edge, parse 'edge' to the first parameter (eg. browser.start "edge").
' To open in App Mode, input the starting URL for appURL parameter as app mode cannot be open on a blank tab.
' If cleanActive is False, the method will skip cleaning currently open instances. This will save
' initialization time but if an instance is already running and not closed, pipe error will be triggered.
' If reAttach is True, the method will always try to reattach to the previous automation session if it is
' still active. If not, it will then move on to start a brand new session. If False, it will always open a new
' session regardless whether there is an on going one or not.
' To add additional arguments to the call string, use the addArguments parameter for instance:
' .start addArgs:="--disable-info-bar" (https://peter.sh/experiments/chromium-command-line-switches/)
' Note: not all arguments work the same for Chrome and Edge. Some may work for one but not the others.
' Updated: 22/03/23 Long Vh:      - Enhanced to dynamically retrieve installation path of the browser
'          26/03/23 Long Vh:      - Made some quality-of-life changes
'          25/05/25 Long Vh:      - set default profile to "CDP" so that it can work with the latest chrome update
'          23/10/25 Daniel Polak: - Added missing space before --homepage
'          22/01/26 Long Vh:      - Removed automatic reattachment from the start method as this restricts
'                                   the possibility of running multiple CDP browser instances under different
'                                   profiles - which is a very useful feature.
'                                 - Tweaked the addProfile section and the cleanUpSessions to reflect the
'                                   recent Chromium updates and the Edge policy discovery.
'--------------------------------------------------------------------------------------------------------------

    Dim strCall As String, strLoc As String, addProfile As String
    Dim intRes As Integer
    Dim timerStart As Single, timeStart As Single, timerOut As Single

    timeStart = Timer

   'Determine default browser parameters -------------------------------------------------------------------------------------
   '07/01/26: avoid using --enable-automation tag as it triggers many web apps to perform bot check
    Select Case name
        Case "edge": brType = brEdge
        Case "chrome": brType = brChrome
        Case Else: printMsg "Browser initiation fails: only the defined browser types can be accepted", doRaiseError:=True
    End Select
    brPath = getBrowserPath(brType)
    strCall = """" & brPath & """ --remote-debugging-pipe --no-first-run"
   '--------------------------------------------------------------------------------------------------------------------------

   'Add starting URL to the call string -------------------------------------------------------------------------------------------------------------------
    If appUrl <> "" Then
        If appUrl = "about:blank" Then printMsg "Browser cannot be open under Application Mode when homepage is set to ""about:blank""", doRaiseError:=True
        If InStr(appUrl, "http") = 0 Then appUrl = "http://" & appUrl   'ie. devs can just simply input "google.com" i/o "http://google.com"
        strCall = strCall & " --app=""" & appUrl & """"
    Else
        strCall = strCall & " --homepage ""about:blank"""
    End If
   '-------------------------------------------------------------------------------------------------------------------------------------------------------

   'Start with the choice profile -------------------------------------------------------------------------------------------------------------
   '22/01/26: I have recently noted that if an organization sets the group policy of Edge of UserDataDir to a specific structure, the argument --user-data-dir
   'will be effectively disabled and blocked to the specified data dir by the policy. In this scenario, CDP session can only load the browser state from the default
   'data dir which already cannot be allowed to run --remote-debugging-pipe from Google's latest Chromium. If Edge's Chromium updates to the latest, this policy will
   'cause the CDP initialization to fail. If Edge's Chromium is still behind the latest Chromium version, CDP init can still be possible with the --profile-directory
   'switch but the requirement is that all running Edge instances that had been started up manually will need to be closed first otherwise the PeekNamedPipe error
   'will result. For Chrome, v136 and later does not honor remote-debugging request if user profile is loaded from the Chrome\User Data\ folder.
   'As a result, here we set it to Chrome\CDP\ folder instead (auto-created if not exists).
   'Additionally, Multi-Instances Operation requires that each instance is assigned to a different user data folder (not profile - each user data folder can have multiple profiles).
   'as a result, if your organization blocks with edge policy, multi-instances operation will not work there. You will have to enable the line with the switch --profile-directory instead.
    Select Case brType
        Case brEdge:
            addProfile = " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Microsoft\Edge\CDP\" & userProfile & """" & " --profile-directory=""Default"""
'            addProfile = " --profile-directory=""" & userProfile & """" 'Use this switch instead if your organization disables UserDataDir with policy
        Case brChrome:
            addProfile = " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Google\Chrome\CDP\" & userProfile & """" & " --profile-directory=""Default"""
    End Select
    strCall = strCall & addProfile
    brUser = userProfile 'store the user profile name for cpid serialization
   '-----------------------------------------------------------------------------------------------------------------------------------------

   'Clean up previous sessions if they could not be reattached to avoid PeekNamedPipe error --------------------------------------------
   'Note: The error usually occurs when the same User Profile is already being used by another pipe. Hence if this
   'session is unable to connect to the previous pipe (for instance, due to the session application host was abruptly terminated),
   'PeekNamedPipe error will be thrown.
    If cleanActive Then cleanUpSessions onlyCdpSessions:=IIf(brType = brEdge, False, True)  'Note: you can set onlyCdpSessions to True by default if your organization does not force UserDataDir policy
    setCrashStateNormal 'Prevent "Restore Session" popup
   '------------------------------------------------------------------------------------------------------------------------------------

   'Prepare for initialization ------------------------------------------------------------------------
    strCall = strCall & " " & addArgs  'Add additional arguments if instructed
   '---------------------------------------------------------------------------------------------------

   'Initialize CDPCore class ----------------------------------------------------------------------------------------------------------
    Set core = New CDPCore
    brCallString = strCall
    intRes = core.init(brCallString, aboolSerializable:=True)  'Serializable is set to True by default for ease of accessibility
    If intRes <> 0 Then printMsg "Unable to initialize CDPCore", doRaiseError:=True Else printMsg "CDPCore initialized successfully"
    lngLastID = 1 'reset the pipe message ID
   '-----------------------------------------------------------------------------------------------------------------------------------

   'Complete initialization ------------------------------------------------------
    attachToSession     'Find and attach the newly open Session ID
    getSessionInfo      'Store initial session info for printParams debugging
    serialize           'Create a serial string for future attachment
   '------------------------------------------------------------------------------

   'Make sure the page is fully loaded for inAppMode ------------------------------------------------------------------------------------
   '07/01/25: Long Vh: simplified the check as there are web apps that does not end up loading the target url but points to another url.
   'As a result, the below check simply checks if there is "movement" in the name of the page instead
    If appUrl <> "" Then
        timerStart = Timer: timerOut = 45
        Do
            If Me.name <> "" Then Exit Do
            If Timer - timerStart > timerOut Then Err.Raise ERR_TIMEOUT, , "Timeout waiting for the browser to load the homepage."
            DoEvents
        Loop
       wait
    End If
   '-------------------------------------------------------------------------------------------------------------------------------------

   'Finally print the session details for debugging ------------------------------
    printParams
    printMsg "Browser initialization completed after " & Timer - timeStart & "s"
   '------------------------------------------------------------------------------

End Function

Private Function serialize() As String
'------------------------------------------------------------------
' This method transcribes Chrome's Json message into a string.
' The serialized string can then be used to automate an active
' session with a known serial.
' Updated: 03/05/23 - Replace sessionID with targetID since it is
'                     constant and stable on every CDP session
'          22/01/26 - Added specification to user profile name to
'                     the cpid file name to allow targetting of
'                     specific profile for reattachment.
'------------------------------------------------------------------

    Dim objSerialize As New Scripting.Dictionary
    Dim f As String
    Dim intF As Integer

    objSerialize.Add "core", core.serialize()
    objSerialize.Add "lngLastID", lngLastID
    objSerialize.Add "targetID", targetID

    serialize = core.jsConverter.ConvertToJson(objSerialize)
    printMsg "New serial string: " & serialize

   'Write CdpSessionID to Temp folder
    f = Environ("TEMP") & "\" & Replace$(LCase$(brUser), " ", "") & ".cdpid.txt"
    intF = FreeFile
    Open f For Binary As #intF
        Put #intF, 1, serialize
    Close #intF

End Function

Private Function deserialize(strSerialized As String)
'------------------------------------------------------------------
' This method deciphers a serialized string to obtain the currently
' active session ID of the already open instance. The serial
' must be obtained first from the serialize method. The isLive
' public function checks if reattachment is successful.
' Updated: 03/05/23 - replace sessionID with targetID which is
'                     constant and stable on every CDP session
'------------------------------------------------------------------

    On Error GoTo ifErr

    Dim objSerialize As Scripting.Dictionary
    Set core = New CDPCore

1   Set objSerialize = core.jsConverter.ParseJson(strSerialized)
2   core.deserialize objSerialize.Item("core")
3   lngLastID = objSerialize.Item("lngLastID")
4   targetID = objSerialize.Item("targetID")
5   sessionID = attachToTab(targetID)
    Exit Function

ifErr: 'if deserialize fails, simply move on to starting anew

    printMsg "deserialization failed at line " & Erl & " for serial string " & strSerialized

End Function

Public Function jsEval(strString As String, Optional dbgMsg As Boolean = True) As Variant
'------------------------------------------------------------
' This function commands the attached browser instance to
' execute a Javascript expression to perform required tasks.
' Note: the original function has a bool retry variable which
' was removed for simplicity.
' Added dbgMsg to read debug line on the go when needed.
'------------------------------------------------------------

    Dim results As Scripting.Dictionary
    Dim strMessage As String
    Dim objMessage As Scripting.Dictionary
    Dim ErrMsg As Variant, result As Variant
    Dim retried As Boolean, boolRetry As Boolean
    Dim strResult As String

step1:

   'Construct message for sending ----------------------------------------------------------
    strMessage = "{""method"":""Runtime.evaluate"",""params"":{""expression"":""1+1;""}}"
    Set objMessage = core.jsConverter.ParseJson(strMessage)
    objMessage.Item("params").Item("expression") = strString & ";"
    strMessage = core.jsConverter.ConvertToJson(objMessage)
   '----------------------------------------------------------------------------------------

   'Send and check result -------------------------------------------------------------------------------------------
    Set results = sendMessage(strMessage, dbgMsg:=dbgMsg)
    If results Is Nothing And boolRetry Then printMsg "Unable to contact the Browser instance!", doRaiseError:=True
    If Not results.Exists("error") Then GoTo step2
   '-----------------------------------------------------------------------------------------------------------------

   'Handle errors ---------------------------------------------------------------------------------------------------
    ErrMsg = results("error")("message")
    printMsg "jsEval Error: " & ErrMsg
    If ErrMsg = "'Runtime.evaluate' wasn't found" Then printMsg "Possibly the main session has been abruptly closed."
    If ErrMsg = "Session with given id not found." And Not retried Then
        printMsg "Attempting to update Session ID " & sessionID
        sessionID = "": sessionID = attachToTab(targetID)
        printMsg "New Session ID obtained: " & sessionID
        retried = True: GoTo step1
    End If
    printMsg "The Browser was unable to evaluate the JS expression: " & strString, doRaiseError:=True
   '-----------------------------------------------------------------------------------------------------------------

step2:

   'If the return type has a specific type we can return the result -------------------------------------------
    Set result = results.Item("result").Item("result")
    Select Case True
        Case result.Item("subtype") = "error"
            jsEval = Split(result.Item("description"), vbNewLine)(0)
        Case result.Item("type") = "string", result.Item("type") = "boolean", result.Item("type") = "number"
            jsEval = result.Item("value")
        Case result.Item("subtype") = "node"
            jsEval = result.Item("className") & " (class='" & result.Item("description") & "')"
        Case result.Item("subtype") = "null"
            jsEval = "null"
    End Select
   '-----------------------------------------------------------------------------------------------------------

step3:

   'Output to Immediate Window if enabled -------------------------------------------------------------------
    If Not jsEval = "" Then strResult = " | Returned: " & jsEval Else strResult = ""
    strResult = Replace(strResult, vbNewLine, "")
    If Len(jsEval) > 200 Then strResult = Left(strResult, 200) & "(showing 200/" & Len(jsEval) & " chars)"
    If dbgMsg Then printMsg "Script executed: " & strString & strResult
   '---------------------------------------------------------------------------------------------------------

End Function

Public Function navigate(strURL As String, Optional till As ReadyState = isComplete)
'-----------------------------------------------------------------------
' This function is a quick wrapper to navigate to an url.
' Updated 23/03/23
'-----------------------------------------------------------------------

   'Check if already navigated
    Dim curURL As String
    curURL = jsEval("window.location.href", False)
    If curURL = strURL Then
        printMsg "Already navigated to """ & strURL & """"
        Exit Function
    End If

   'Quick formatting for simple URL input
    If strURL <> "about:blank" And InStr(strURL, "http") = 0 Then strURL = "http://" & strURL

   'If not yet then do navigate
    Dim results As Scripting.Dictionary
    Set results = sendMessage("{""method"":""Page.navigate"",""params"":{""url"":""" & strURL & """}}")
    wait till:=till
    getSessionInfo 'update new session info for printParams
    printMsg "Navigated to: " & strURL

End Function

Public Function quit()
'-----------------------------------------------------------------------------------
' Closes the currently attached browser instance
' Original Name: closeBrowser
' Update: 24/10/2025: Daniel Polak - avoid immediate hard kills, reduce race conditions, and make shutdown behavior predictable and robust in non-responsive cases
'-----------------------------------------------------------------------------------

    Dim results As Scripting.Dictionary
    Dim sngT As Single

    On Error Resume Next
    Set results = sendMessage("{""method"":""Browser.close"",""params"":{}}")

    sngT = Timer
    Do While isLive And Timer - sngT < 5
        sleep 0.1
        DoEvents
    Loop
    If isLive Then
        printMsg "Browser didn't close gracefully within 5s; invoking cleanUpSessions."
        cleanUpSessions
    End If
    deleteCDPID False

    sleep 1 'it seems without waiting a bit the browser crashes and the next time it wants to recover from a crash
    printMsg "The browser has been terminated."

End Function

Public Function isLive() As Boolean
'-----------------------------------------------------------------------------------
' Check if there is ongoing connection with an attached instance
' Original Name: connectionAlive
'-----------------------------------------------------------------------------------

    On Error GoTo ifErr
    Dim strLoc As String
    strLoc = jsEval("window.location.href", dbgMsg:=False)
    isLive = True
    Exit Function

ifErr:
    isLive = False

End Function

Public Function wait(Optional till As ReadyState = isComplete, Optional dbgState As Boolean = False)
'--------------------------------------------------------------------------------------------------
' Wait until browser reaches a specific state. Default is until it fully loads ("complete")
' Possible states:  loading / interactive / complete
' Complete state is preferable when certain information is only showed when fully loaded.
' Otherwise, an interactive state is good enough and can potentially save a lot of wait time.
' However, do note that some web apps only fully displays elements at complete readystate.
' To read ReadyState on the go, set the dbgState to True to print the state to the Immediate Window
' Original Name: waitCompletion
' Updated: 26/04/23 - Enhanced 'till' to Enum type for ease of use
'                   - Added exit for targetState = interactive in case CDP misses the state change
'--------------------------------------------------------------------------------------------------

    Dim strState As String, targetState As String
    Select Case till
        Case isLoading: targetState = "loading"
        Case isInteractive: targetState = "interactive"
        Case isComplete: targetState = "complete"
    End Select

    Do Until strState = targetState
        strState = jsEval("document.readyState", dbgMsg:=False)                 'dbgMsg = False to not spam Immediate Window
        If dbgState Then printMsg "Current ReadyState: " & strState             'option to print ReadyState info during iteration
        If targetState = "interactive" And strState = "complete" Then Exit Do   'also exit if better than interactive is satisfied
        sleep 0.1                                                               'reduce sleep will speed up but will cost cpu power
    Loop

    If dbgState Then printMsg "Wait '" & targetState & "' satisfied"

End Function

Public Function show(Optional state As WinState = asNormal, Optional xywh As String = "0 0 0 0")
'-----------------------------------------------------------------------------------
' Display the window with various optional state. The default state is "normal".
' Can also use after .attach to bring the window topmost as well. By default at
' start, the browser window's state will be that of the previous session run.
' xywh is used for inputing left top width height parameters:
' eg.1: xywh = "100 200" left = 100 and top = 200, width and height unchanged
' eg.2: xywh = "100 100 300 400" left = top = 100, width = 300 and height = 400
' eg.3: xywh = "0 0 0 0" the method will ignore and not make any resizing
'-----------------------------------------------------------------------------------
' Note: This method targets the window based on its name so it is best applied when
' the browser window is not refreshing a URL which may potentially change the name.
' Updated: 02/03/23 - Changed CInt to Int to prevent Overflow error.
'-----------------------------------------------------------------------------------

    Dim h As Long, w As Long, x As Long, y As Long

    If brHWnd = 0 Then brHWnd = getHandle
    ShowWindow brHWnd, doShow 'Make visible
    ShowWindow brHWnd, state

   'Resize to specific position and size
   'Will not work properly if window is currently maximized
    If xywh <> "0 0 0 0" Then

        sleep 0.3      'Sleep a bit for previous Win Api executions to finish, else the resize effect will be accidentally cancelled out

        Dim params As New Scripting.Dictionary
        If brWID = "" Then brWID = invokeMethod("Browser.getWindowForTarget").Item("windowId")
        params("windowId") = Int(brWID)     '02/03/23 CInt() will throw error, Int() is better

        Dim bounds As New Scripting.Dictionary
        On Error Resume Next
        x = CLng(Split(xywh, " ")(0))
        y = CLng(Split(xywh, " ")(1))
        w = CLng(Split(xywh, " ")(2))
        h = CLng(Split(xywh, " ")(3))
        On Error GoTo 0

        If x <> 0 Then bounds("left") = x
        If y <> 0 Then bounds("top") = y
        If w <> 0 Then bounds("width") = w
        If h <> 0 Then bounds("height") = h
        Set params("bounds") = bounds
        invokeMethod "Browser.setWindowBounds", params

    End If

End Function

Public Function hide()
'------------------------------------------------------------------
' This is the equivalent of show "hidden" but this function is
' added to make it easy to access and use the same feature.
'------------------------------------------------------------------

    If brHWnd = 0 Then brHWnd = getHandle
    ShowWindow brHWnd, asHidden

End Function

Public Function bringToForeground()
'------------------------------------------------------------------
' Brings the browser window to the foreground
' Updated: Long Vh - 06/01/26 should be a function like others
'------------------------------------------------------------------

    If brHWnd = 0 Then brHWnd = getHandle
    If brHWnd <> 0 Then
        ShowWindow brHWnd, doShow
        BringWindowToTop brHWnd
        SetForegroundWindow brHWnd
    End If

End Function

Public Function newTab(Optional url As String, Optional newWindow As Boolean, Optional setMain As Boolean) As CDPBrowser
'------------------------------------------------------------------
' Create a new blank tab within the same instance.
' url: The initial URL the page will be navigated to. An empty
' string indicates about:blank.
' newWindow: If true will open a separate window instead of a tab.
' In the original CDP document, each new tab is called "Target" and
' the equivalence of this method is Target.createTarget
' This returns the Session ID of the newly created tab.
' Updated: 20/03/23 Return the result as a new Browser object
'          26/04/23 Format url argument to accept simplified input
'          22/01/26 Added .reattach where appropriate
'------------------------------------------------------------------

    Dim params As New Scripting.Dictionary
    Dim SessionID_ As String

    If url <> "" Then url = IIf(url <> "about:blank" And InStr(url, "http") = 0, "http://" & url, url)
    params("url") = CStr(url)
    If newWindow Then params("newWindow") = CBool(newWindow)

    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.createTarget", params)

   'Return the new tab's Session ID
    Dim tabId As String
    tabId = results("targetId")
    SessionID_ = attachToTab(tabId)

   'Switch attachment to the new tab
   '22/01/26: From v3.1, reattachment is no longer automatic so .reattach is needed
    printMsg "newTab will now open a new CDP instance"
    Set newTab = New CDPBrowser
    newTab.reattach brUser
    newTab.id = SessionID_
    If url <> "" Then newTab.wait

   'Set as main session tab if required
    If Not setMain Then Exit Function
    Me.id = SessionID_
    getSessionInfo
    serialize
    printMsg "Tab [" & brName & " | " & targetID & "] is now the main session tab for " & brID

End Function

Public Function getTab(Optional tabName As String, Optional url As String, Optional setMain As Boolean, Optional doRetry As Boolean = False) As CDPBrowser
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Search and find the Session ID of the tab already open based on
' its name or its url. tabName and url arguments can be partial. By
' omitting both these arguments, getTab will try to attach to the
' nearest unattached tab.
' Update: 06/06/23 - Added doRetry argument
'         03/05/23 - Combined the method with the getTabNew
'         24/10/25 Daniel Polak - Predictable, precise, and script-friendly tab selection that scales well when there are many similar tabs open
'                                 Exact title and URL-prefix matching avoids accidentally selecting tabs with similar titles/URLs
'         23/10/25 Daniel Polak - Fixed typo timeStart instead of timerStart and fixed bug 60000s instead of 60s
'         22/01/26 Long Vh      - Added .reattach where appropriate
'-----------------------------------------------------------------------------------------------------------------------------------------------------

    Dim results As Scripting.Dictionary, obj As Scripting.Dictionary
    Dim tabId As String, candidateTabId As String, ErrMsg As String
    Dim SessionID_ As String
    Dim targets As Collection
    Dim timerStart As Single
    Dim tabfound As Boolean
    Dim bestScore As Long, score As Long
    Dim info As String

    If doRetry Then timerStart = Timer

searchit:

    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")

    bestScore = -1
    candidateTabId = ""
    For Each obj In targets
        If obj("type") = "page" Then
            score = -1
            Select Case True
                Case tabName <> ""
                    If obj.Exists("title") Then
                        If obj("title") = tabName Then
                            score = 100
                        ElseIf InStr(obj("title"), tabName) Then
                            score = 50
                        Else
                            ErrMsg = "with name containing """ & tabName & """"
                        End If
                    End If
                Case url <> ""
                    If obj.Exists("url") Then
                        If Left$(obj("url"), Len(url)) = url Then
                            score = 90
                        ElseIf InStr(obj("url"), url) Then
                            score = 40
                        Else
                            ErrMsg = "with url containing """ & url & """"
                        End If
                    End If
                Case Else
                    If obj("attached") = False Then
                        score = 80
                    Else
                        score = 30
                        ErrMsg = "as no new unattached tab could be found"
                    End If
            End Select
            If score > bestScore Then
                bestScore = score
                candidateTabId = obj("targetId")
            End If
        End If
    Next

    If bestScore >= 0 Then
        tabfound = True
        tabId = candidateTabId
    End If

   'Do retry as long as not timeout (default 60s)
    If Not tabfound And doRetry And Timer - timerStart < 60 Then GoTo searchit

   'Returns nothing if not found
   'Nothing is preferred over throwing error
    If Not tabfound Then
        info = IIf(tabName <> "", tabName, IIf(url <> "", url, "(no new tab found)"))
        printMsg "Unable to find the tab: " & info
        Exit Function
    Else
        printMsg "Found the tab with ID: " & tabId
    End If

searchfound:

   'Return the session object for reference
    SessionID_ = attachToTab(tabId)

   'Return the new CDP object
   '22/01/26 .reattach brUser is needed from v3.1 onwards to enable multi-instances automation
    Set getTab = New CDPBrowser
    getTab.reattach brUser
    getTab.id = SessionID_  'will also update the targetID
    sleep 'give some time so that the debugging session can switch to the new tab

   'Set as main session tab if required
    If Not setMain Then Exit Function
    Me.id = SessionID_
    getSessionInfo
    serialize
    printMsg "Tab [" & targetID & "] is now the main session tab for " & brID

End Function

Public Function closeTab() As Boolean
'------------------------------------------------------------------
' Close the current tab of the session. Requires targetId which can
' be retrieved by using getTargetId. This is the targetId of the
' active tab under focus of the automation pipe.
'------------------------------------------------------------------

   'Get the original serialized tab ID
    Dim mainID As Variant, serial As Variant
    serial = getSerialString()
    mainID = core.jsConverter.ParseJson(serial).Item("targetID")
    If targetID = mainID Then
        printMsg "Unable to close this tab " & targetID & " as it is the main session tab"
        Exit Function
    End If

    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(getTargetID)
    invokeMethod "Target.closeTarget", params
    printMsg brID & " was closed"

   'Return the function result
    closeTab = True

End Function

Public Function getElementByID(strID As String) As CDPElement
'----------------------------------------------------------------------------------
' Get an element class from the element ID, similar to the getElementByID of JS.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Get a new unique varID
    Dim varID As String: varID = newVarID

   'Get the object and assign to the varID in the browser
    Dim varJS As String: Dim result As String
    varJS = varID & " = document.getElementById(""" & strID & """)"
    result = jsEval(varJS)

   'Return result
    Set getElementByID = New CDPElement
    getElementByID.init Me, varIDs, varID, varJS, result

End Function

Public Function getElementByXPath(strXPath As String) As CDPElement
'---------------------------------------------------------------------------------------
' Get an element class from XPath. The element class is init in CDPElement module.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation.
' XPath Syntax help: https://www.w3schools.com/xml/xpath_syntax.asp
' Updated: 02/03/2023
' Examples -----------------------------------------------------------------------------
' eg.1 "//b[text()='T-BILL CURVE']": tag "b" with innertext to be 'T-BILL CURVE'
' eg.2 "//input[@name='username']": tag "input" & attribute "name" to be 'username'
' eg.3 "//a[@a1='abc' and @a2='xyz']": tag "a" with attribute a1 = abc and a2 = xyz
' eg.4 "//div[@id='abc']/div": get first div child of the div with id = abc
' eg.5 "//div[@id='abc'].//div": get first div descendant of the div with id = abc
' eg.6 "div": get first div child from the prescribed contextNode (default = 'document')
'---------------------------------------------------------------------------------------

   'Get a new unique varID
    Dim varID As String: varID = newVarID

   'Verify double slash syntax
   'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")

   'Get the object and assign to the varID in the browser
    Dim varJS As String: Dim result As String
    varJS = varID & " = document.evaluate(""" & strXPath & """, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
    result = jsEval(varJS)

   'Return result
    Set getElementByXPath = New CDPElement
    getElementByXPath.init Me, varIDs, varID, varJS, result

End Function

Public Function getElementsByXPath(strXPath As String) As Collection
'----------------------------------------------------------------------------------
' Get a collection of elements by XPath. For further info on the particular XPath
' syntax used for retrieving the elements, refer to Mozilla doc below:
' https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
' Note that index for accessing the collection members start at one (not zero).
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Get a new unique varID
    Dim varID As String: varID = newVarID & "s"     's to signify a collection varID
    Dim length_ As Long, n As Long
    Dim result As Variant
    Dim tElement As Object
    Dim elements As Collection

   'Verify double slash syntax
   'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")

   'Get the object and assign to the varID in the browser
    result = jsEval(varID & " = document.evaluate(""" & strXPath & """, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null)")

   'Return result
    If result = "null" Then Exit Function Else Set elements = New Collection
   'Store element members to the collection
    Dim varJS As String                                 'Store each new member access JS string
    length_ = jsEval(varID & ".snapshotLength")
    For n = 0 To length_ - 1                            'HTML index starts at zero
        varJS = varID & ".snapshotItem(" & n & ")"
        Set tElement = New CDPElement
        tElement.init Me, varIDs, varJS_:=varJS
        elements.Add tElement
    Next

    Set getElementsByXPath = elements

End Function

Public Function getElementByQuery(strQuery As String) As CDPElement
'----------------------------------------------------------------------------------
' Get an element class by QuerySelector - similar to that of the same in JS.
' The element class has useful properties and methods similar to the HTML element
' object in previous IE automation. Refer to the below for some examples.
' https://www.w3schools.com/jsref/met_document_queryselector.asp
' https://stackoverflow.com/a/29937844
' https://www.w3.org/TR/selectors-3/#context
' Updated: 08/11/2023
' Examples ------------------------------------------------------------------------
' eg.1 "#abc": get the first element with id attribute = abc
' eg.2 "tb.xyz": get the first element with tag tb and with class xyz
' eg.3 "a[name='abc']": get the first element with tag a and attribute name = abc
' eg.4 "div > p": get the first element with tag 'p' and has parent with tag 'div'
' eg.5 "[a1='abc'][a2='xyz']": get element with attribute a1 = abc and a2 = 123
' eg.6 "[a1='abc'] [a2='xyz']": get element with attribute a2 = xyz and is a child
' of an element with attribute a1 = abc
' eg.7 "#processType [id*='select_option']": get child element with id text that
' contains "select_option" and is a child of element with id "processType"
' eg.8 "#upload:not([disabled])": get the first element with id=upload and does not
' have attribute 'disabled'.
' eg.9 "#upload:disabled": get the first element with id=upload and has attribute
' 'disabled'.
' eg.10 "#id1 tr:last-child [id*='invoiceno']": first get element with id text
' 'id1', then look for the last child with tag 'tr', then finally look within the
' child's children for the first element with id containing 'invoiceno'.
'----------------------------------------------------------------------------------

   'Get a new unique varID
    Dim varID As String: varID = newVarID

   'Get the object and assign to the varID in the browser
    Dim varJS As String: Dim result As String
    varJS = varID & " = document.querySelector(""" & strQuery & """)"
    result = jsEval(varJS)

   'Return result
    Set getElementByQuery = New CDPElement
    getElementByQuery.init Me, varIDs, varID, varJS, result

End Function

Public Function getElementsByQuery(strQuery As String) As Collection
'----------------------------------------------------------------------------------
' Get a collection of elements by QuerySelectorAll. The syntax is as below:
' https://developer.mozilla.org/en-US/docs/Web/API/XPathResult/iterateNext
' Note that index for accessing the collection members start at one (not zero).
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Get a new unique varID
    Dim varID As String: varID = newVarID & "s"     's to signify a collection varID
    Dim length_ As Long, n As Long
    Dim result As Variant
    Dim tElement As Object
    Dim elements As Collection

   'Get the object and assign to the varID in the browser
    result = jsEval(varID & " = document.querySelectorAll(""" & strQuery & """)")

   'Return result
    If result = "null" Then Exit Function Else Set elements = New Collection

   'Store element members to the collection
    Dim tElem As CDPElement
    Dim varJS As String                                 'Store each new member access JS string
    length_ = jsEval(varID & ".length")
    For n = 0 To length_ - 1                            'HTML index starts at zero
        varJS = varID & ".item(" & n & ")"
        Set tElement = New CDPElement
        tElement.init Me, varIDs, varJS_:=varJS
        elements.Add tElement
    Next

    Set getElementsByQuery = elements

End Function

Public Function jsAddLib(strURL)
'----------------------------------------------------------------------------------
' Load the external JS Library to the document for use of specialized JS methods.
' This is very useful when built-in JS commands are not sufficient to execute a
' task easily, such as when taking a screenshot of an entire webpage while this can
' be done relatively easy using external JS libraries. This function shall enables
' massive possibility of extended features.
' https://stackoverflow.com/a/57572107
' Updated: 23/03/2023
'          23/10/2025 Daniel Polak - fixed bug in check whether JS library is loaded
'----------------------------------------------------------------------------------
' Note: the script library should only be added after the page is fully loaded. If
' it is loaded when the HTML document is refereshing, it risks adding the script to
' the old HTML document which will be replaced by the new one causing the script to
' be inexistent and resulting in "ReferenceError" due to its functions not defined.
'----------------------------------------------------------------------------------

   'Check if already loaded, if so then no need to add again
    If getElementByQuery("script[src='" & strURL & "']").varResult <> "null" Then
        printMsg "The HTML already contains the JS Library: """ & strURL & """"
        Exit Function
    End If

   'Add script loading function
   '(Javascript Original)
    '    function loadScript(name) {
    '        var script = document.createElement('script');
    '        script.type = 'text/javascript'; script.src = name;
    '        document.head.appendChild(script);
    '        script.onload = () => {};                              //added to make sure script has been loaded
    '        return 'Script Library fully loaded';
    '    }
    jsEval "function loadScript(name) { var script = document.createElement('script'); script.type = 'text/javascript'; script.src = name; document.head.appendChild(script); script.onload = () => {}; return 'Script library loaded'; }"

   'Load the script URL
    jsEval "loadScript('" & strURL & "')"
    printMsg "Successfully added the JS Library: """ & strURL & """"

End Function

Public Function jsAddScript(strLocalScriptPath)
'----------------------------------------------------------------------------------
' Load the contents of the local script file directly to the HTML page. This is
' useful when you have a custom local script to add in a certain feature to a web
' application (eg adding a custom popup message).
' Updated: 01/12/2025
'----------------------------------------------------------------------------------
' Note: the library should only be added after the page is fully loaded. If
' it is loaded when the HTML document is refereshing, it risks adding the script to
' the old HTML document which will be replaced by the new one causing the script to
' be inexistent and resulting in "ReferenceError" due to its functions not defined.
'----------------------------------------------------------------------------------

    If Dir(strLocalScriptPath) = "" Then Exit Function

    Dim fso As New FileSystemObject
    jsEval fso.OpenTextFile(strLocalScriptPath).ReadAll 'Load the script contents
    printMsg "Successfully added the local script: """ & strLocalScriptPath & """"

    Set fso = Nothing

End Function

Public Function reattach(Optional userProfile As String = "CDP1") As Boolean
'----------------------------------------------------------------------------------
' Reattach to the target CPID file. This method enables the management of multiple
' CDP instances at once under different Chrome/Edge user profiles. The benefit of
' this is that if one instance fails/crashes, it does not affect the other ones.
' and thus enables the possibility of running multiple CDP automation operations
' on entirely different tasks simultaneously.
'----------------------------------------------------------------------------------

    Dim strSerialized As String

    brUser = userProfile
    strSerialized = getSerialString()
    deserialize strSerialized

    If isLive Then
        reattach = True
        printParams
    Else
        reattach = False
        printMsg "Failed to reattach to previous CPID."
    End If

End Function

Public Function notify(msg)
'----------------------------------------------------------------------------------
' Send a custom notification message on the bottom right corner of the screen. This
' popup message is nicer than alert() and it is not modal: ie. does not the user to
' interact with it while pausing the automation.
' Credit: Claude AI
' Updated: 01/12/2025
'----------------------------------------------------------------------------------

   'Make sure backslash if any is escaped
    msg = Replace(msg, "\", "\\")

    jsEval "function showStatus(msg, duration = 10000) {" & _
                "let c = document.getElementById('statusContainer');" & _
                "if (!c) {" & _
                    "c = document.createElement('div');" & _
                    "c.id = 'statusContainer';" & _
                    "c.style.cssText = 'position:fixed;bottom:20px;right:20px;z-index:9999;';" & _
                    "document.body.appendChild(c);" & _
                "}" & _
                "const m = document.createElement('div');" & _
                "m.style.cssText = 'background:#4CAF50;color:white;padding:16px 24px;margin-top:10px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.3);white-space:nowrap;transition:all 0.3s;opacity:0;transform:translateX(400px);font-size:16px;font-family:system-ui,-apple-system,BlinkMacSystemFont,Roboto,sans-serif;font-weight:300;line-height:1.5;';" & _
                "m.textContent = msg;" & _
                "c.appendChild(m);" & _
                "setTimeout(() => { m.style.opacity = '1'; m.style.transform = 'translateX(0)'; }, 10);" & _
                "setTimeout(() => {" & _
                    "m.style.opacity = '0';" & _
                    "m.style.transform = 'translateX(400px)';" & _
                    "setTimeout(() => c.removeChild(m), 300);" & _
                "}, duration);" & _
            "}"

    jsEval "showStatus('" & msg & "')"

End Function

Public Function snapPage(fullNameWithPngExtension As String, Optional getFullPage As Boolean = False)
'----------------------------------------------------------------------------------
' Capture a screenshot of the entire HTML page from the top to bottom. The method
' employs html2canvas external library is which loaded via addJsLibrary. To snap
' only a certain element within the page, indicate the same with the targetElement
' argument (default is "document.body" for the entire page).
' Original source: https://stackoverflow.com/a/51478809
' Updated: 14/04/23
'          23/10/25 Daniel Polak - fixed bug where the all files in the Downloads folder were deleted instead of just the temporary screenshot
'          24/10/25 Daniel Polak - added CDP native screenshot capability which avoids use of external JavaScript library
'          07/01/26 Long Vh      - removed the use of html2canvas js lib in favor of using the native CDP method to capture the page. Overhauled the function design.
'----------------------------------------------------------------------------------
' Note : if getFullPage = false, snapPage will only get a screenshot of the current
'        view. If true, it will get all the way from the start to the bottom. See
'        demo.getSnapShot sub procedure for example.
'----------------------------------------------------------------------------------

    Dim params As New Scripting.Dictionary
    Dim r As Scripting.Dictionary
    Dim b64 As String, tmpPng As String

   'Verify if the page can be captured with screenshot
    Set r = invokeMethod("Page.getLayoutMetrics")
    If invokeError(r) Then
        printMsg "Unable to perform snapPage on this screen as it has no layout metrics."
        Exit Function
    End If

   'Perform the snip
    If getFullPage Then params.Add "captureBeyondViewport", True
    Set r = invokeMethod("Page.captureScreenshot", params)

   'Save to the same folder if not indicated
    If Not invokeError(r) Then
        b64 = r("data")
        SaveBase64ToFile b64, fullNameWithPngExtension
        printMsg "Screenshot saved to: " & fullNameWithPngExtension
    End If

End Function

Public Function printMsg(strMsg As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
'----------------------------------------------------------------------------------
' A custom printMsg that besides printing to the Immediate Window, it will
' also print to a dated log file. This is useful at user level where
' the Immediate Window is not easily accessible. doRaiseError will also raise an
' error aside from printing the error message to the log file.
' Updated: 07/01/26 Long Vh: the message output is disabled by default to avoid
'                            spamming info. This can be turned on by setting globally
'                            noDbgMsg = False or noLogging = False
'----------------------------------------------------------------------------------

    Dim intFile As Integer
    Dim logName As String, logPath As String, strFormattedMsg As String

    If isHeader Then strFormattedMsg = String(100, "-") & vbNewLine & strMsg & vbNewLine & String(100, "-") _
    Else strFormattedMsg = Format(Now, "hh:mm:ss") & " | " & brID & " | " & strMsg

   'Perform msg output to Immediate Window if needed
    If doPrintDbgMsg Then
        Debug.Print strFormattedMsg
    End If

   'Perform msg logging if needed
    If doLog Then
        logPath = thisApp.Path & "\Logs\"
        logName = "log" & UCase(Format(Now, "ddMMMyy")) & ".cdp.txt"
        If Dir(logPath, vbDirectory) = "" Then MkDir logPath
        intFile = FreeFile
        Open logPath & logName For Append As #intFile
        Print #intFile, strFormattedMsg
        Close #intFile
    End If

    If doRaiseError Then Err.Raise ERR_PROTOCOL, Description:=strMsg

End Function

Public Function printTabs()
'------------------------------------------------------------------
' Print all current tabs info. Useful for debugging.
'------------------------------------------------------------------

    Dim results As Scripting.Dictionary, target_ As Scripting.Dictionary
    Dim i As Long
    Dim targets As Collection
    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")

   'Read info
    printMsg "---------------------------------------------------"
    printMsg "Tabs Information"
    printMsg "---------------------------------------------------"

    i = 1
    For Each target_ In targets
        If target_("type") = "page" Then
            printMsg "#" & i & vbTab & _
                        "type: " & target_("type") & " | " & _
                        "attached: " & target_("attached") & " | " & _
                        "title: " & target_("title") & " | " & _
                        "url: " & target_("url") & " | " & _
                        "target ID: " & target_("targetId")
            i = i + 1
        End If
    Next

End Function

Public Function printParams()
'------------------------------------------------------------------
' Print all parameters of the current session for debuggings.
'------------------------------------------------------------------

    printMsg "---------------------------------------------------"
    printMsg brID & " Information"
    printMsg "---------------------------------------------------"

    printMsg "Session ID    | sessionID | " & sessionID
    printMsg "Last ID       | lngLastID | " & lngLastID
    printMsg "Window ID     | brWId     | " & brWID
    printMsg "Browser Type  | brType    | " & brType
    printMsg "Browser Ver   | brVer     | " & brVer
    printMsg "Browser Path  | brPath    | " & brPath
    printMsg "Tab URL       | brUrl     | " & brURL
    printMsg "Tab Title     | brName    | " & brName
    printMsg "Tab ID        | targetID  | " & targetID
    printMsg "Core StdInWr  | hStdInWr  | " & core.hStdInWr
    printMsg "Core StdOutRd | hStdOutRd | " & core.hStdOutRd
    printMsg "Core CDPInWr  | hCDPInWr  | " & core.hCDPInWr
    printMsg "Core CDPOutRd | hCDPOutRd | " & core.hCDPOutRd
    printMsg "CDP Version   | CDPVer    | " & CDPVer

End Function

Public Function tabCount() As String
'------------------------------------------------------------------------------
' Print all current tabs info. Useful for debugging.
' Updated : 05/02/25 - changed from printTabs to tabCount to get the count too
'------------------------------------------------------------------------------

    Dim results, targets, target_ As Scripting.Dictionary
    Dim i As Integer

    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")

   'Read info
    printMsg "---------------------------------------------------"
    printMsg "Tabs Information"
    printMsg "---------------------------------------------------"

    i = 0
    For Each target_ In targets
        If target_("type") = "page" Then
            printMsg "#" & i & vbTab & _
                        "type: " & target_("type") & " | " & _
                        "attached: " & target_("attached") & " | " & _
                        "title: " & target_("title") & " | " & _
                        "url: " & target_("url") & " | " & _
                        "target ID: " & target_("targetId")
            i = i + 1
        End If
    Next

    tabCount = i

End Function

Public Function sleep(Optional seconds As Double = 0.5)
'------------------------------------------------------------------
' Custom sleep function. Sleep by 0.5s by default.
' Useful for a quick necessary pause when needed.
'------------------------------------------------------------------

    Const baseUnit As Long = 1000    'ie. millisecs

    sleep2 seconds * baseUnit
    DoEvents

End Function

Public Function activate()
'----------------------------------------------------------------------------------
' Switch focus to the current tab being attached to under the current BRID.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(getTargetID)
    invokeMethod "Target.activateTarget", params

End Function

Public Function handleDialog(accept As Boolean, Optional promptText As Variant)
'----------------------------------------------------------------------------------
' Accepts or dismisses a JavaScript initiated dialog (alert, confirm, prompt, or
' onbeforeunload). accept = True to accept the dialog, else dimiss it. promptText
' is the text to enter into the dialog prompt before accepting, used only if this is a
' prompt dialog.
' Note: this method requires the execution of "pageEnable" before the dialog appears
' so that the dialog event is probably exposed and captured as an event.
' Updated: 23/07/2025 - added as requested by Jane Zhang for her project
'----------------------------------------------------------------------------------

    Dim params As New Scripting.Dictionary

    params("accept") = CBool(accept)
    If Not IsMissing(promptText) Then params("promptText") = CStr(promptText)
    invokeMethod "Page.handleJavaScriptDialog", params

End Function

Public Function pageEnable()
'----------------------------------------------------------------------------------
' Enable the page for events listening. Required for handleDialog for instance.
'----------------------------------------------------------------------------------

    invokeMethod "Page.enable"

End Function

'=============================================================================
' PUBLIC PROPERTIES
'=============================================================================

Property Get url() As String
'------------------------------------------------------------
' Get the URL of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    brURL = jsEval("document.URL", False)
    url = brURL

End Property

Property Let url(newURL As String)
'------------------------------------------------------------
' Get the URL of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    navigate newURL
    brURL = newURL

End Property

Property Get name() As String
'------------------------------------------------------------
' Get the title of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    brName = jsEval("document.title", False)
    name = brName

End Property

Property Let name(newName As String)
'------------------------------------------------------------
' Set the title of the current document/tab
' Added 02/03/2023
'------------------------------------------------------------

    jsEval "document.title = '" & newName & "'", False
    printMsg "Tab name changed to " & newName
    brName = newName

End Property

Property Get id() As String
'------------------------------------------------------------
' Get the session ID of the current browser
' Added 02/03/2023
'------------------------------------------------------------

    id = sessionID

End Property

Property Let id(newID As String)
'------------------------------------------------------------
' Set the session ID of the current browser to a new value
' Added 02/03/2023
'------------------------------------------------------------

    sessionID = newID
    getSessionInfo        'update all session infos
    printMsg brID & " updated Session ID to " & sessionID

End Property

Property Get html() As String
'------------------------------------------------------------
' Get the html text of the entire document
' Added 02/03/2023
'------------------------------------------------------------

    html = jsEval("document.documentElement.innerHTML")

End Property

'=============================================================================
' PRIVATE FUNCTIONS
'=============================================================================

Public Function invokeMethod(methodName As String, Optional params As Scripting.Dictionary, Optional dbgMsg As Boolean = False) As Scripting.Dictionary
'-----------------------------------------------------------------
' This function enables advanced method invocation under CDP.
' For a list of all methods available, see:
' https://chromedevtools.github.io/devtools-protocol/tot/Browser/
' Error message will be prompted automatically by invokeError in
' the Immediate Window.
' The return object is Json String that has been converted to a
' dictionary object.
' Credit: https://github.com/PerditionC/VBAChromeDevProtocol
' Updated: 28/04/23 - Set this to Public as it is useful
'          20/03/23 - Added printMsg msgs & set to Private
'-----------------------------------------------------------------

    Dim paramStr As String
    Dim results As Scripting.Dictionary

    If params Is Nothing Then paramStr = "{}" Else paramStr = core.jsConverter.ConvertToJson(params)
    Set results = sendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}")

   'Print error if there is
    If invokeError(results) Then printMsg "Failed to invoke method """ & methodName & """", doRaiseError:=True

   'Return result object
    If results.Exists("result") Then Set invokeMethod = results.Item("result")
    If dbgMsg Then printMsg "Invoked """ & methodName & """ with result: " & core.jsConverter.ConvertToJson(results("result"))

End Function

Private Function sendMessage(strMessage As String, Optional objAllMessages As Scripting.Dictionary, Optional dbgMsg As Boolean = True) As Scripting.Dictionary
'--------------------------------------------------------------
' Send message to CDP and check if the message is received
' yet by the Chromium browser.
' Updated: 03/05/23 - Added soft-reset on peeknamepipe error
'          22/03/23 - Added error handling for better debugging
'          24/10/25 Daniel Polak - Added soft re-attach attempt when pipe error
'--------------------------------------------------------------

    Dim intRes As Long
    Dim strRes As String
    Dim lngCurrentId As Long
    Dim results As Scripting.Dictionary       'to store a collection of IDs obtained from CDP
    Dim errorDesc As String
    Dim errorLine As Long
    Dim timerOut As Single, timerStart As Single
    Dim blnDidSoftRetry As Boolean, blnDoSoftRetry As Boolean

    lngCurrentId = lngLastID
    lngLastID = lngLastID + 1       'We increase the global ID counter

RetrySend:

    On Error GoTo ifError

   'Before sending a message the message-buffer is emptied
   'All messages that we have received so far cannot be an answer to the message that we will send
   'So they can be safely discarded
1   If objAllMessages Is Nothing Then
        intRes = 1
        Do Until intRes < 1
2           intRes = core.readProcCDP(strRes)
3           If intRes > 0 Then strBuffer = strBuffer & strRes
        Loop

        Dim lngNullCharPos As Long
4       lngNullCharPos = searchNull()

        Do Until lngNullCharPos = 0
5           strBuffer = Right(strBuffer, Len(strBuffer) - lngNullCharPos)                        'if necessary, add "printMsg (left(strBuffer, lngNullCharPos))" above this for debugging
6           lngNullCharPos = searchNull()
        Loop
    End If

   '--------------------------------------------------
   'Send message to the browser
   '--------------------------------------------------
7   intRes = core.readProcSTD(strRes)                                                            'Sometimes edge writes to stdout so we clear stdout too.
8   strMessage = Left(strMessage, Len(strMessage) - 1)                                           'We add the currentID and sessionID to the message
9   If sessionID <> "" Then strMessage = strMessage & ", ""sessionId"":""" & sessionID & """"
10  strMessage = strMessage & ", ""id"":" & lngCurrentId & "}" & vbNullChar
11  core.writeProc strMessage                                                                    'Send message to the browser instance

   '------------------------------------------------
   'Read message returned by browser
   '------------------------------------------------
    timerStart = Timer                                                                           'Failsafe timer to ensure loop ends when server fails to respond
    timerOut = 30                                                                                'Max timeout period in seconds, snapPage usually takes 15s so the timeout needs to be sufficient
    Do
        intRes = 1
        Do Until intRes < 1                                                                      'We read from edge and process messages until we receive a message with our ID
12          intRes = core.readProcCDP(strRes)
13          If intRes > 0 Then strBuffer = strBuffer & strRes
        Loop

14      lngNullCharPos = searchNull()

        Do Until lngNullCharPos = 0
15          strRes = Left(strBuffer, lngNullCharPos - 1)
16          strBuffer = Right(strBuffer, Len(strBuffer) - lngNullCharPos)
            If seeRawSendMsgDbg Then printMsg "strRes = " & strRes                              'Only enable if needs to see every raw JS message for debugging/development
            If strRes <> "" Then
17              Set results = core.jsConverter.ParseJson(strRes)
18              If Not objAllMessages Is Nothing Then objAllMessages.Add CStr(objAllMessages.Count), results
19              If results.Exists("id") Then
20                  If results.Item("id") = lngCurrentId Then
21                      Set sendMessage = results
                        Exit Function
                    End If
                End If
            End If
22          lngNullCharPos = searchNull()
        Loop

        DoEvents
23      If Timer - timerStart > timerOut Then Err.Raise 900, , "Timeout waiting for browser to return message."
    Loop
    Exit Function

afterErr:

    printMsg "Attempt to soft re-attach when pipe error because sometimes a simple retry is all it is needed to make it work :)"
    On Error Resume Next
    attachToSession
    GoTo RetrySend

ifError:

    If Not dbgMsg Then Exit Function    'Allow the option to bypass error if intended

   'Record errmsg and print helpful infos
    errorLine = Erl
    errorDesc = Err.Description
    printMsg "sendMessage error at line " & errorLine & ": """ & errorDesc & """"
    printMsg "strMessage used by sendMessage: " & strMessage

    blnDoSoftRetry = False
    If Not blnDidSoftRetry Then
        If InStr(errorDesc, "PeekNamedPipe") > 0 And brReAttach Then
            blnDoSoftRetry = True
            blnDidSoftRetry = True
            Err.Clear
        End If
    End If
    If blnDoSoftRetry Then
        Resume afterErr
    End If

    If InStr(errorDesc, "PeekNamedPipe") > 0 Then
        printMsg "Possible reasons: (1) the browser is not actually running (ie. call string failure); (2) the browser window was accidentally closed; (3) the CDP Pipe connection to the browser was cut-off somehow."
        deleteCDPID False
    End If

    Err.Raise IIf(InStr(errorDesc, "PeekNamedPipe") > 0, ERR_PIPE, ERR_PROTOCOL), , "sendMessage error at line " & errorLine & ": """ & errorDesc & """"

End Function

Private Function attachToSession() As Integer
'-----------------------------------------------------------------------
' This function must be called after start and before all other methods.
' This function attaches to a session of the browser.
' Updated: 03/05/23 - Added printMsg msgs & removed URL checking req.
'-----------------------------------------------------------------------
' Note: if needed to read the returned message result from sendMessage,
' type ?core.jsConverter.ConvertToJson(objAllMessages) in the ImmWnd
'-----------------------------------------------------------------------

    Dim results As Scripting.Dictionary
    Dim results_ As Scripting.Dictionary
    Dim results__ As Scripting.Dictionary
    Dim results___ As Scripting.Dictionary
    Dim boolFound As Boolean
    Dim errorDesc As String
    Dim errorLine As Long

    Dim strKey As Variant

    On Error GoTo ifError

    Dim objAllMessages As New Scripting.Dictionary
1   Set results = sendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":true}}", objAllMessages)
    printMsg "Now looking for the browser window"

   'Read message to confirm the requested url is found
   'Checking results__.Item("url") has been skipped for simplicity since v2.3
   'A new session always has the property attached = false and method = Target.targetCreated
   'This way of verification vastly improves the stability over checking by URL as sometimes
   'there are rogues targets with similar URL but they are not attachable (ie. no Session ID)
    For Each strKey In objAllMessages.Keys
       Set results = objAllMessages.Item(strKey)
       If results.Exists("params") And results.Item("method") = "Target.targetCreated" Then
            Set results_ = results.Item("params")
            If results_.Exists("targetInfo") Then
                Set results__ = results_.Item("targetInfo")
                If results__.Item("attached") = False And results__.Item("type") = "page" Then
                    boolFound = True
                    Exit For
                End If
            End If
        End If
    Next strKey
2   If Not boolFound Then printMsg "Unable to find the browser session", doRaiseError:=True _
    Else printMsg "Now attaching CDP pipes to the browser session"

   'Establish connection with the CDP pipes
3   targetID = results__.Item("targetId")
4   Set results___ = sendMessage("{""method"":""Target.attachToTarget"",""params"":{""targetId"":""" & results__.Item("targetId") & """,""flatten"":true}}")
5   sessionID = results___.Item("result").Item("sessionId")
6   Set results___ = sendMessage("{""method"":""Runtime.enable"",""params"":{}}")
7   Set results___ = sendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":false}}")

    printMsg "Successfully attached to Session ID " & sessionID
    Exit Function

ifError:

   'Record errmsg and print helpful infos
    errorLine = Erl
    errorDesc = Err.Description
    printMsg "attachToSession error at line " & errorLine & ": """ & errorDesc & """"
    printMsg "objAllMessages = " & core.jsConverter.ConvertToJson(objAllMessages)
    printMsg "results = " & core.jsConverter.ConvertToJson(results)
    printMsg "results_ = " & core.jsConverter.ConvertToJson(results_)
    printMsg "results__ = " & core.jsConverter.ConvertToJson(results__)
    printMsg "results___ = " & core.jsConverter.ConvertToJson(results___)
    printMsg ".start call string: " & brCallString

    printParams
    Err.Raise ERR_PROTOCOL, , "attachToSession error at line " & errorLine & ": """ & errorDesc & """"

End Function

Private Function searchNull() As Long
'------------------------------------------------------------
' CDP messages received from chrome are null-terminated
' Updated: 25/10/25: Daniel Polak - new faster version
'------------------------------------------------------------

    Dim lngPos As Long

    If Len(strBuffer) = 0 Then Exit Function
    lngPos = InStr(1, strBuffer, vbNullChar, vbBinaryCompare)
    If lngPos > 0 Then
        searchNull = lngPos
    Else
        searchNull = 0
    End If

End Function

Private Function invokeError(result As Scripting.Dictionary) As Boolean
'-----------------------------------------------------------------------------------
' Check if InvokeMethod fails and what the error message returned is.
' Error message will be shown via printMsg
' Credit: https://github.com/PerditionC/VBAChromeDevProtocol
'-----------------------------------------------------------------------------------

    Dim ErrorCode As Long
    Dim ErrorData As Variant, ErrorMessage As Variant

    If result Is Nothing Then
        ErrorCode = -1 'arbitrary nonzero value, may change in future
        ErrorMessage = "Timeout or No results"
        ErrorData = vbEmpty
        printMsg "InvokeMethod Error: [" & ErrorCode & "] " & ErrorMessage
        invokeError = True
    ElseIf result.Exists("error") Then
        Dim errorDetails As Scripting.Dictionary
        Set errorDetails = result("error")
        ErrorCode = errorDetails("code")
        ErrorMessage = errorDetails("message")
        printMsg "InvokeMethod Error: [" & ErrorCode & "] " & ErrorMessage
        If errorDetails.Exists("data") Then
            ErrorData = errorDetails("data")
            printMsg "       " & ErrorData
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

Private Function getHandle() As LongPtr
'-----------------------------------------------------------------------------------
' Added by Long: obtain the window handle of the current browser session
' Updated: 02/03/23 - Add mechanism to work with browser windows that are blank.
'-----------------------------------------------------------------------------------

    Dim bufferString As String
    Dim titleLen As Long
    Dim hWnd As LongPtr
    Dim sessionTitle As String, windowTitle As String

   'Determine current tab title, if error: the tab is a blank tab
    sessionTitle = jsEval("document.getElementsByTagName(""title"")[0].innerHTML", False)
    If InStr(sessionTitle, "TypeError") Then sessionTitle = "about:blank"

   'Obtain handles to the correct browser window
   'Both Edge and Chrome use the same class chromeWindowClass
    hWnd = FindWindowEx(0&, 0&, chromeWindowClass, vbNullString)        'To return Chrome / Edge window handles
    Do

       'Get Window Title based on specified handles
       'Requires string buffering declaration as string
        bufferString = String(1024, " ")
        titleLen = GetWindowText(hWnd, bufferString, Len(bufferString))
        windowTitle = Left(bufferString, titleLen)

       'Verify if matching with the sessionTitle and is really an edge/chrome window
        If windowTitle <> "" Then If InStr(windowTitle, sessionTitle) Then Exit Do

        hWnd = FindWindowEx(0&, hWnd, chromeWindowClass, vbNullString)  'Continue with the next handle value

    Loop Until hWnd = 0 ' no more matching windows

    getHandle = hWnd

End Function

Private Function cleanUpSessions(Optional onlyCdpSessions As Boolean = True)
'-----------------------------------------------------------------------------------------------------------------------------------------
' Provides cleaning before firing up a new session to prevent
' pipe error
' Update: 05/11/24: Daniel Polak - modified to solve problems with closing sessions (that do not belong to this user) on terminal server
'         23/10/25: Daniel Polak - p.name and p.terminate were not correct, replaced by proc.name and proc.terminate (discovered this bug because we are using Option Explicit now). On Error Resume Next was hiding the bug too.
'                                - only terminate automation-processes (where command line contains --remote-debugging-pipe)
'-----------------------------------------------------------------------------------------------------------------------------------------

    Dim proc As Object
    Dim lProcessId As Long
    Dim lSessionId As Long
    Dim strFName As String

    lProcessId = GetCurrentProcessId()          'Get the current process ID
    ProcessIdToSessionId lProcessId, lSessionId 'Get the Session ID for the current process

   'Clean up all running sessions
    On Error Resume Next
    For Each proc In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process Where SessionId = " & lSessionId)
        If LCase$(proc.name) = LCase$(brType) Then
            If InStr(proc.CommandLine, "--remote-debugging-pipe") Then  'a CDP process
                If InStr(proc.CommandLine, brUser) Then proc.Terminate
            Else                                                        'for browsers open manually
                If Not onlyCdpSessions Then proc.Terminate
            End If
        End If
        DoEvents
    Next
    On Error GoTo 0
    Err.Number = 0

    deleteCDPID False

   'Clean core params for new session
    lngLastID = 0   'To make a clean slate for new session
    sessionID = ""  'Likewise

    printMsg "cleanUpSessions completed and previous CdpSessionID removed"

End Function

Private Function setCrashStateNormal()
'------------------------------------------------------------------
' Set Crashed state to Normal to prevent the "Restore Sessions?"
' popup from appearing every time the browser is started after
' being closed by browser.quit method.
' Requires user data dir profile stored in private var brUser and
' the browser type stored in private variable brType.
' Update: 22/01/26 - Tweaked to work well with the new multi-instance
'                    structure.
'------------------------------------------------------------------

    Dim strPrefFolder As String, strFilePath As String
    Dim strContents As String, strNewContents As String
    Dim intF As Integer
    Dim paths As New Collection, p As Variant
       
    paths.Add Environ("LOCALAPPDATA") & "\Google\Chrome\User Data\CDP\" & brUser & "\Preferences"
    paths.Add Environ("LOCALAPPDATA") & "\Microsoft\Edge\User Data\CDP\" & brUser & "\Preferences"
    paths.Add strFilePath = Environ("LOCALAPPDATA") & "\Microsoft\Edge\User Data\" & brUser & "\Preferences"
    
   'Mark 'Normal' on the preference file to prevent Restore Page popup on start
    For Each p In paths
        strFilePath = p
        If LenB(Dir(strFilePath)) > 0 Then
        
            intF = FreeFile
            Open strFilePath For Binary As #intF
                strContents = Input$(LOF(intF), #intF)
            Close #intF
        
           'Precise replacement: only replace exit_type value
            strNewContents = Replace(strContents, """exit_type"":""Crashed""", """exit_type"":""Normal""", 1, -1, vbBinaryCompare)
        
            If strNewContents <> strContents Then
                intF = FreeFile
                Open strFilePath For Binary As #intF
                    Put #intF, 1, strNewContents
                Close #intF
            End If
        
            printMsg "setCrashStateNormal completed"
            
        End If
    Next
    
    Set paths = Nothing

End Function

Private Function attachToTab(tabId As String) As String
'----------------------------------------------------------------------------------
' Attach to a specific tab with its tabId (a.k.a targetId)
' flatten: Enables "flat" access to the session via specifying sessionId attribute
' in the commands. Set to True by default to make attachToTab work.
' Returns: a session ID string that is needed for CDP automation.
'----------------------------------------------------------------------------------

    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(tabId)
    params("flatten") = True

    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.attachToTarget", params)
    attachToTab = results("sessionId")

End Function

Private Function getTargetID() As String
'----------------------------------------------------------------------------------
' Retrieve the targetId of the current tab (a.k.k "target" in CDP Documentation).
' This targetId is useful for tab automation, such as the .closeTab method.
'----------------------------------------------------------------------------------

    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.getTargetInfo")
    getTargetID = results("targetInfo").Item("targetId")

End Function

Private Function getSerialString() As String
'----------------------------------------------------------------------------------
' Check if a SessionID has already been created to determine if an automated
' browser session is already on going. The return is a Serialized String in which
' SessionID is a part of the serials.
' Updated: 21/01/2026
'----------------------------------------------------------------------------------

    Dim strContents As String
    Dim intF As Integer
    Dim strFName As String

   'Append user profile to the cpid file name for specific targetting
   'This is usually needed when multiple instances of CDP are running and subsequent calls
   'need to point to different CPIDs from time to time to reattach
    strFName = Environ("TEMP") & "\" & Replace$(LCase$(brUser), " ", "") & ".cdpid.txt"

    If Len(Dir(strFName)) > 0 Then
        deleteCDPID True
        intF = FreeFile
        Open strFName For Binary As #intF
            strContents = Input$(LOF(intF), #intF)
        Close #intF
        getSerialString = strContents
    End If

End Function


Private Function newVarID() As String
'----------------------------------------------------------------------------------
' Create a unique varID for element reference with CDPElement. If the ID is not
' unique, it is possible that new element assignment will override existing ones
' under the same ID by chance if working on a large number of elements.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

    Dim newID As String

   'Init the ID collection
    If varIDs Is Nothing Then Set varIDs = New Scripting.Dictionary

    Do: newID = "varID" & Format(Rnd * 100000, "000000")
    Loop Until Not varIDs.Exists(newID)
    varIDs.Add newID, newID

   'Return the new varID
    newVarID = newID

End Function

Private Function getBrowserPath(strBrowserType As String) As String
'----------------------------------------------------------------------------------
' Obtain the installation path of the browser program. This is useful when
' the installation path is changed by the user or their corporation for security.
' Updated: 22/03/2023
' 20251023 Daniel Polak - added fallback path when registry lookup fails (in a locked-down environment)
'----------------------------------------------------------------------------------

    Dim objWSHost As Object, reg As Object
    Dim strKeyPath As String

    Select Case True
        Case InStr(LCase(strBrowserType), "chrome"): brType = brChrome
        Case InStr(LCase(strBrowserType), "edg"): brType = brEdge         'On edge it is just "Edg"
        Case Else: printMsg "Unable to getBrowserPath for browser type: " & brType, doRaiseError:=True
    End Select

   'Return the path
    On Error Resume Next
    Set reg = GetObject("winmgmts://./root/default:StdRegProv")
    Set objWSHost = CreateObject("WScript.Shell")
    strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & brType & "\Path"
    getBrowserPath = objWSHost.RegRead(strKeyPath) & "\" & brType

    If Err.Number <> 0 Or LenB(getBrowserPath) = 0 Then
        Err.Clear
        Select Case brType
            Case brChrome
                If LenB(Dir(Environ("ProgramFiles") & "\Google\Chrome\Application\" & brType)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles") & "\Google\Chrome\Application\" & brType
                ElseIf LenB(Dir(Environ("ProgramFiles(x86)") & "\Google\Chrome\Application\" & brType)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles(x86)") & "\Google\Chrome\Application\" & brType
                End If
            Case brEdge
                If LenB(Dir(Environ("ProgramFiles") & "\Microsoft\Edge\Application\" & brType)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles") & "\Microsoft\Edge\Application\" & brType
                ElseIf LenB(Dir(Environ("ProgramFiles(x86)") & "\Microsoft\Edge\Application\" & brType)) > 0 Then
                    getBrowserPath = Environ("ProgramFiles(x86)") & "\Microsoft\Edge\Application\" & brType
                End If
        End Select
        If LenB(getBrowserPath) > 0 Then
            If LenB(Dir(getBrowserPath)) = 0 Then getBrowserPath = ""
        End If
        If LenB(getBrowserPath) = 0 Then printMsg "Unable to locate " & brType & " installation path.", doRaiseError:=True
    End If

End Function

Private Function getSessionInfo()
'----------------------------------------------------------------------------------
' Obtain session info on the fly for printParams.
' Updated: 23/01/26 - remove the requirement to activate the window and the handle
'                     retrieval
'----------------------------------------------------------------------------------

   'Get infos
    brURL = url
    brName = name
    brWID = invokeMethod("Browser.getWindowForTarget").Item("windowId")
    brType = invokeMethod("Browser.getVersion").Item("product")
    brVer = Split(brType, "/")(1)
    brPath = getBrowserPath(brType)
    targetID = getTargetID

End Function

Private Function deleteCDPID(ByVal blnOnlyDeleteIfCreatedBeforeToday As Boolean)
'----------------------------------------------------------------------------------
' By: Daniel Polak
'----------------------------------------------------------------------------------

    Dim strFName As String

    strFName = Environ("TEMP") & "\" & Replace$(LCase$(brUser), " ", "") & ".cdpid.txt"
    If Len(Dir(strFName)) > 0 Then
        If blnOnlyDeleteIfCreatedBeforeToday = False Or DateDiff("d", FileDateTime(strFName), Now) > 0 Then
            On Error Resume Next
            Kill strFName
            On Error GoTo 0
        End If
    End If

End Function

Private Function SaveBase64ToFile(strB64 As String, strPath As String)
'----------------------------------------------------------------------------------
' Decode a base64 string and save to file without external dependencies (late bound).
'----------------------------------------------------------------------------------

    Dim dom As Object, el As Object, bytes() As Byte, stm As Object

    Set dom = CreateObject("MSXML2.DOMDocument")
    Set el = dom.createElement("tmp")
    el.DataType = "bin.base64"
    el.Text = strB64
    bytes = el.nodeTypedValue

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write bytes
    stm.SaveToFile strPath, 2
    stm.Close

End Function

Private Sub Class_Initialize()
'----------------------------------------------------------------------------------
' A quick way to quickly check if an automated session is already active and open.
' If so, automatically attempt to reattach to it.
' Updated: 07/06/23
'          22/10/25 Daniel Polak - prevent compilation errors because depending on whether we are running in Excel, Word or Access only one of ThisWorkbook, ActiveDocument and CurrentProject is available
'          22/01/26 Long Vh - removed the auto deserialization feature to running multiple instances of browser profiles.
'----------------------------------------------------------------------------------

   'Determine project host type for compatibility
   'CallByName is used to avoid errors on object names not existing during compilation
    Select Case True
        Case InStr(Application.name, "Excel"):  Set thisApp = CallByName(Application, "ThisWorkbook", VbGet)
        Case InStr(Application.name, "Word"):   Set thisApp = CallByName(Application, "ActiveDocument", VbGet)
        Case InStr(Application.name, "Access"): Set thisApp = CallByName(Application, "CurrentProject", VbGet)
    End Select

   'Assign an objID for easy debug identification
    brID = "BRID" & Format(Rnd * 1000, "000")

End Sub

'===============================================================================================
' Experimental Functions
'===============================================================================================

