VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPBrowser"
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
'       18/08/23 Long Vh: - set core to Private i/o Public as it is meant to be Private.
'       06/06/23 Long Vh: - applied recommendations by MitsOmitsaras for Access & Word compat.
'       26/05/23 Long Vh: - made some tweaks for compatibility with Access & Word out-of-the-box
'       26/04/23 Long Vh: - added .setMain for better tab management.
'       22/03/23 Long Vh: - added many .getElementsBy methods.
'                         - added advanced methods to extend features of CDP apps.
'                         - changed deserialize/serialize methods to read from a temp file.
'                         - added .snapPage method for capturing an entire web screen.
'                         - added interactions with the new element class object.
'                         - many other changes to improve debug handlings.
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

Private Const CDPVer As String = "2.7.4"

'===================================
' Win APIs Declarations
'===================================

Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" ( _
    ByVal dwMilliseconds As Long)
    
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExW" ( _
    ByVal hwndParent As LongPtr, _
    ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As LongPtr, _
    ByVal lpszWindow As LongPtr) As LongPtr
    
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextW" ( _
    ByVal hWnd As LongPtr, _
    ByVal lpString As LongPtr, _
    ByVal nMaxCount As Long) As Long
    
Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nCmdShow As Long) As Boolean
   
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

'===================================
' Global Variables
'===================================

Private core As CDPCore                 'Must be initialized for CDP engine
Private lngLastID As Long               'Every message sent to edge has an ID
Private strBuffer As String             'Buffer string, required for various win APIs
Private brHWnd As LongPtr                  'Browser object window handle
Private brWID As String                 'Useful for windows repositioning
Private brType As String                'Browser type: either 'chrome' or 'edge'
Private brUser As String                'Browser user profile
Private brPath As String                'Browser default installation path
Private brID As String                  'The ID of the current CDPBrowser object. Useful when reading debugs of multiple tabs.
Private brName As String                'The focused tab name / window name of the active CDPBrowser
Private brURL As String                 'The focused tab url of the active CDPBrowser
Private brVer As String                 'The current browser version
Private brCallString As String          'The initial call string used by .start
Private varIDs As Scripting.Dictionary  'Collection of all unique element IDs for CDPElement reference - Added 02/03/2023
Private sessionID As String             'The ID to the current session tab.
Private targetID As String              'The targetID of the current tab for InvokeMethod
Private thisApp As Object               'A dymanic variable used in CDPBrowser initialize procedure to determine if the VBA host is an Excel / Word / Access

'=============================================================================
' PUBLIC FUNCTIONS
'=============================================================================

Public Sub start(Optional name As String = "chrome", _
    Optional appUrl As String, _
    Optional cleanActive As Boolean = True, Optional reAttach As Boolean = True, _
    Optional userProfile As String = "Default", Optional addArgs As String)
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
' Updated: 22/03/23 - Enhanced to dynamically retrieve installation path of the browser
'          26/03/23 - Made some quality-of-life changes
'--------------------------------------------------------------------------------------------------------------

    Dim strCall As String
    Dim result As Long
    Dim timeStart As Single
                           
    timeStart = Timer

   'Determine if already open -----------------------------------------------------
    If reAttach And isLive Then
        printParams     'Useful for checking
        Exit Sub
    End If
    printMsg "Reattachment aborted as reAttach=" & reAttach & " & isLive=" & isLive
    printMsg "New Browser Session Started", isHeader:=True
   '-------------------------------------------------------------------------------
                                   
   'Determine default browser parameters -------------------------------------------------------------------------------------
    Select Case name
        Case "edge": brType = brEdge
        Case "chrome": brType = brChrome
        Case Else: printMsg "Browser initiation fails: only the defined browser types can be accepted", doRaiseError:=True
    End Select
    brPath = getBrowserPath(brType)
    strCall = """" & brPath & """ --remote-debugging-pipe --enable-automation"
   '--------------------------------------------------------------------------------------------------------------------------

   'Add starting URL to the call string -------------------------------------------------------------------------------------------------------------------
    If appUrl <> "" Then
        If appUrl = "about:blank" Then printMsg "Browser cannot be open under Application Mode when homepage is set to ""about:blank""", doRaiseError:=True
        If InStr(appUrl, "http") = 0 Then appUrl = "http://" & appUrl   'ie. devs can just simply input "google.com" i/o "http://google.com"
        strCall = strCall & " --app=""" & appUrl & """"
    Else
        strCall = strCall & "--homepage ""about:blank"""
    End If
   '-------------------------------------------------------------------------------------------------------------------------------------------------------

   'Start with a choice profile -------------------------------------------------------------------------------------------------------------
   Dim addProfile As String
   If userProfile <> "Default" Then
        Select Case brType
            Case brEdge: addProfile = " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Microsoft\Edge\User Data\" & userProfile & """"
            Case brChrome: addProfile = " --user-data-dir=""" & Environ("LOCALAPPDATA") & "\Google\Chrome\User Data\" & userProfile & """"
        End Select
        strCall = strCall & addProfile
    End If
    brUser = userProfile
   '-----------------------------------------------------------------------------------------------------------------------------------------
       
   'Prepare for initialization ------------------------------------------------------------------------
    strCall = strCall & " " & addArgs    'Add additional arguments if instructed
    If cleanActive Then cleanUpSessions  'Clean up all running browser instances to prevent pipe error
    setCrashStateNormal                  'Prevent "Restore Session" popup
   '---------------------------------------------------------------------------------------------------

   'Initialize CDPCore class ----------------------------------------------------------------------------------------------------------
    Set core = New CDPCore
    brCallString = strCall
    result = core.init(brCallString)
    If result <> 0 Then printMsg "Unable to initialize CDPCore", doRaiseError:=True Else printMsg "CDPCore initialized successfully"
    lngLastID = 1
   '-----------------------------------------------------------------------------------------------------------------------------------
       
   'Complete initialization ------------------------------------------------------
    attachToSession     'Find and attach the newly open Session ID
    getSessionInfo      'Store initial session info for printParams debugging
    serialize           'Create a serial string for future attachment
    printParams         'Print session info for debugging
    printMsg "Browser initialization completed after " & Timer - timeStart & "s"
   '------------------------------------------------------------------------------
    
   'Make sure the page is fully loaded for inAppMode ------------------------------------------------------------------------------------
    Dim timerStart As Single, timerOut As Long, strLoc As String
    If appUrl <> "" Then
        timerStart = Timer: timerOut = 45
        Do
            strLoc = jsEval("window.location.href", dbgMsg:=False)
            If InStr(strLoc, appUrl) Then Exit Do 'Should be fuzzy check with Instr only as browser often mutates the original string
            If Timer - timerStart > timerOut Then Err.Raise 900, , "Timeout waiting for the browser to load the homepage."
            sleep 0.1
        Loop
       wait
    End If
   '-------------------------------------------------------------------------------------------------------------------------------------
End Sub

Private Function serialize() As String
'------------------------------------------------------------------
' This method transcribes Chrome's Json message into a string.
' The serialized string can then be used to automate an active
' session with a known serial.
' Updated: 03/05/23 - Replace sessionID with targetID since it is
'                     constant and stable on every CDP session
'------------------------------------------------------------------
        
    Dim objSerialize As New Scripting.Dictionary
    Dim fso As Object, f As String

    objSerialize.Add "core", core.serialize()
    objSerialize.Add "lngLastID", lngLastID
    objSerialize.Add "targetID", targetID

    serialize = core.jsConverter.ConvertToJson(objSerialize)
    printMsg "New serial string: " & serialize

   'Write CdpSessionID to Temp folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    f = Environ("Temp") & "\cdpid.txt"
    fso.CreateTextFile(f).Write serialize
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
           
    On Error GoTo iferr
           
    Dim objSerialize As Scripting.Dictionary
    Set core = New CDPCore
    
1   Set objSerialize = core.jsConverter.ParseJson(strSerialized)
2   core.deserialize objSerialize.Item("core")
    
3   lngLastID = objSerialize.Item("lngLastID")
4   targetID = objSerialize.Item("targetID")
5   sessionID = attachToTab(targetID)
    Exit Function
        
iferr: 'if deserialize fails, simply move on to starting anew
        
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
    
step1:
    
   'Construct message for sending ----------------------------------------------------------
    strMessage = "{""method"":""Runtime.evaluate"",""params"":{""expression"":""1+1;""}}"
    Set objMessage = core.jsConverter.ParseJson(strMessage)
    objMessage.Item("params").Item("expression") = strString & ";"
    strMessage = core.jsConverter.ConvertToJson(objMessage)
   '----------------------------------------------------------------------------------------
        
   'Send and check result -------------------------------------------------------------------------------------------
    Set results = sendMessage(strMessage, dbgMsg:=dbgMsg)
    If results Is Nothing Then printMsg "Unable to contact the Browser instance!", doRaiseError:=True
    If Not results.Exists("error") Then GoTo step2
   '-----------------------------------------------------------------------------------------------------------------
    
   'Handle errors ---------------------------------------------------------------------------------------------------
    Dim errMsg As String, retried As Boolean
    errMsg = results("error")("message")
    printMsg "jsEval Error: " & errMsg
    If errMsg = "'Runtime.evaluate' wasn't found" Then printMsg "Possibly the main session has been abruptly closed."
    If errMsg = "Session with given id not found." And Not retried Then
        printMsg "Attempting to update Session ID " & sessionID
        sessionID = "": sessionID = attachToTab(targetID)
        printMsg "New Session ID obtained: " & sessionID
        retried = True: GoTo step1
    End If
    printMsg "The Browser was unable to evaluate the JS expression: " & strString, doRaiseError:=True
   '-----------------------------------------------------------------------------------------------------------------
    
step2:

   'If the return type has a specific type we can return the result -------------------------------------------
    Dim result As Object
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
    Dim strResult As String
    If Not jsEval = "" Then strResult = " | Returned: " & jsEval Else strResult = ""
    strResult = Replace(strResult, vbNewLine, "")
    If Len(jsEval) > 200 Then strResult = Left(strResult, 200) & "(showing 200/" & Len(jsEval) & " chars)"
    If dbgMsg Then printMsg "Script executed: " & strString & strResult
   '---------------------------------------------------------------------------------------------------------
        
End Function

Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
'-----------------------------------------------------------------------
' This sub is a quick wraper to navigate to an url.
' Updated 23/03/23
'-----------------------------------------------------------------------
    Dim curURL As String
   'Check if already navigated
    curURL = jsEval("window.location.href", False)
    If curURL = strURL Then
        printMsg "Already navigated to """ & strURL & """"
        Exit Sub
    End If

   'Quick formatting for simple URL input
    If strURL <> "about:blank" And InStr(strURL, "http") = 0 Then strURL = "http://" & strURL

   'If not yet then do navigate
    Dim results As Scripting.Dictionary
    Set results = sendMessage("{""method"":""Page.navigate"",""params"":{""url"":""" & strURL & """}}")
    'should also add a maxTimeLimit to Wait
    wait till:=till
    getSessionInfo 'update new session info for printParams
    printMsg "Navigated to: " & strURL
End Sub

Public Sub quit()
'-----------------------------------------------------------------------------------
' Closes the currently attached browser instance
' Original Name: closeBrowser
'-----------------------------------------------------------------------------------
    
    Dim results As Scripting.Dictionary
    Dim f As File, tempFolder As Folder, fso As New FileSystemObject
    
    On Error Resume Next
    Set results = sendMessage("{""method"":""Browser.close"",""params"":{}}")
    
    Set tempFolder = fso.GetFolder(Environ("Temp"))
   'Clean up CdpSessionID file
    For Each f In tempFolder.Files
        If InStr(f, "cdpid.txt") Then
            fso.DeleteFile f
            Exit For
        End If
    Next
    sleep 1   'It seems without waitng a bit the browser crashes and the next time it wants ro recover from a crash
    printMsg "Browser was terminated."
    
End Sub

Public Function isLive() As Boolean
'-----------------------------------------------------------------------------------
' Check if there is ongoing connection with an attached instance
' Original Name: connectionAlive
'-----------------------------------------------------------------------------------

    On Error GoTo iferr
    Dim strLoc As String
    strLoc = jsEval("window.location.href", dbgMsg:=False)
    isLive = True
    Exit Function

iferr:
    isLive = False
End Function


Public Sub wait(Optional till As ReadyState = isComplete, Optional dbgState As Boolean = False)
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
    Dim targetState As String, strState As String
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
End Sub


Public Sub show(Optional state As WinState = asNormal, Optional xywh As String = "0 0 0 0")
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

    If brHWnd = 0 Then brHWnd = getHandle
    ShowWindow brHWnd, doShow  'Make visible
    ShowWindow brHWnd, state

   'Resize to specific position and size
   'Will not work properly if window is currently maximized
    If xywh <> "0 0 0 0" Then
    
        sleep 0.3      'Sleep a bit for previous Win Api executions to finish, else the resize effect will be accidentally cancelled out
                    
        Dim params As New Scripting.Dictionary
        If brWID = "" Then brWID = invokeMethod("Browser.getWindowForTarget").Item("windowId")
        params("windowId") = Int(brWID)     '02/03/23 CInt() will throw error, Int() is better
                                
        Dim bounds As New Scripting.Dictionary
        Dim x As Long, y As Long, w As Long, h As Long
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
End Sub


Public Sub hide()
'------------------------------------------------------------------
' This is the equivalent of show "hidden" but this function is
' added to make it easy to access and use the same feature.
'------------------------------------------------------------------
    If brHWnd = 0 Then brHWnd = getHandle
    ShowWindow brHWnd, asHidden
End Sub


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
'------------------------------------------------------------------
    Dim sessionID_ As String
    Dim params As New Scripting.Dictionary
    If url <> "" Then url = IIf(url <> "about:blank" And InStr(url, "http") = 0, "http://" & url, url)
    params("url") = CStr(url)
    If newWindow Then params("newWindow") = CBool(newWindow)

    Dim results As Scripting.Dictionary
    Set results = invokeMethod("Target.createTarget", params)
   
   'Return the new tab's Session ID
    Dim tabId As String
    tabId = results("targetId")
    sessionID_ = attachToTab(tabId)

   'Switch attachment to the new tab
    printMsg "newTab will now open a new CDP instance"
    Set newTab = New CDPBrowser
    newTab.id = sessionID_
    If url <> "" Then newTab.wait

   'Set as main session tab if required
    If Not setMain Then Exit Function
    Me.id = sessionID_
    getSessionInfo
    serialize
    printMsg "Tab [" & brName & " | " & targetID & "] is now the main session tab for " & brID
End Function


Public Function getTab(Optional tabName As String, Optional url As String, Optional setMain As Boolean, Optional doRetry As Boolean = False) As CDPBrowser
'------------------------------------------------------------------
' Search and find the Session ID of the tab already open based on
' its name or its url. tabName and url arguments can be partial. By
' omitting both these arguments, getTab will try to attach to the
' nearest unattached tab.
' Update: 06/06/23 - Added doRetry argument
'         03/05/23 - Combined the method with the getTabNew
'------------------------------------------------------------------

    Dim results As Scripting.Dictionary
    Dim tabId As String
    Dim timerStart As Single
    Dim targets As Object, obj As Object
    Dim tabfound As Boolean
    Dim errMsg As String
    Dim timeStart As Long

    If doRetry Then timerStart = Timer

searchit:

    Set results = invokeMethod("Target.getTargets")
    Set targets = results("targetInfos")
    
    For Each obj In targets
        If obj("type") = "page" Then
            Select Case True
                Case tabName <> ""
                    If InStr(obj("title"), tabName) Then tabfound = True Else errMsg = "with name containing """ & tabName & """"
                Case url <> ""
                    If InStr(obj("url"), url) Then tabfound = True Else errMsg = "with url containing """ & url & """"
                Case Else
                    If obj("attached") = False Then tabfound = True Else errMsg = "as no new unattached tab could be found"
            End Select
            If tabfound Then
                tabId = obj("targetId")
                Exit For
            End If
        End If
    Next

   'Do retry as long as not timeout (default 60s)
    If Not tabfound And doRetry And Timer - timeStart < 60000 Then GoTo searchit

   'Returns nothing if not found
   'Nothing is preferred over throwing error
    If Not tabfound Then Exit Function

    printMsg "Found the tab with ID: " & tabId

searchsound:
   Dim sessionID_ As String
   'Return the session object for reference
    sessionID_ = attachToTab(tabId)

   'Return the new CDP object
    Set getTab = New CDPBrowser
    getTab.id = sessionID_  'will also update the targetID

   'Set as main session tab if required
    If Not setMain Then Exit Function
    Me.id = sessionID_
    getSessionInfo
    serialize
    printMsg "Tab [" & targetID & "] is now the main session tab for " & brID
End Function


Public Sub closeTab()
'------------------------------------------------------------------
' Close the current tab of the session. Requires targetId which can
' be retrieved by using getTargetId. This is the targetId of the
' active tab under focus of the automation pipe.
'------------------------------------------------------------------
    Dim serial As String, mainID As String
   'Get the original serialized tab ID
    serial = getSerialString
    mainID = core.jsConverter.ParseJson(serial).Item("targetID")
    If targetID = mainID Then printMsg "Unable to close this tab " & targetID & " as it is the main session tab", doRaiseError:=True
           
    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(getTargetID)
    invokeMethod "Target.closeTarget", params
    printMsg brID & " was closed"
End Sub


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
       
   'Verify double slash syntax
   'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")
       
   'Get the object and assign to the varID in the browser
    Dim result As Variant
    Dim elements As Collection
    result = jsEval(varID & " = document.evaluate(""" & strXPath & """, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null)")

   'Return result
    If result = "null" Then Exit Function Else Set elements = New Collection
   'Store element members to the collection
    Dim varJS As String                                 'Store each new member access JS string
    Dim length_ As Long
    length_ = jsEval(varID & ".snapshotLength")
    Dim tElement As CDPElement, N As Long
    For N = 0 To length_ - 1                            'HTML index starts at zero
        varJS = varID & ".snapshotItem(" & N & ")"
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
' Updated: 02/03/2023
' Examples ------------------------------------------------------------------------
' eg.1 "#abc": get the first element with id attribute = abc
' eg.2 "tb.xyz": get the first element with tag tb and with class xyz
' eg.3 "a[name='abc']": get the first element with tag a and attribute name = abc
' eg.4 "div > p": get the first element with tag 'p' and has parent with tag 'div'
' eg.5 "[a1='abc'][a2='xyz']": get element with attribute a1 = abc and a2 = 123
' eg.6 "[a1='abc'] [a2='xyz']": get element with attribute a2 = xyz and is a child
' of an element with attribute a1 = abc
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
           
   'Get the object and assign to the varID in the browser
    Dim result As Variant, elements As Collection
    result = jsEval(varID & " = document.querySelectorAll(""" & strQuery & """)")
    
   'Return result
    If result = "null" Then Exit Function Else Set elements = New Collection
            
   'Store element members to the collection
    Dim tElement As CDPElement
    Dim varJS As String                                 'Store each new member access JS string
    Dim length_ As Long, N As Long
    length_ = jsEval(varID & ".length")
    For N = 0 To length_ - 1                            'HTML index starts at zero
        varJS = varID & ".item(" & N & ")"
        Set tElement = New CDPElement
        tElement.init Me, varIDs, varJS_:=varJS
        elements.Add tElement
    Next
    
    Set getElementsByQuery = elements
    
End Function


Public Sub jsAddLib(strURL As String)
'----------------------------------------------------------------------------------
' Load the external JS Library to the document for use of specialized JS methods.
' This is very useful when built-in JS commands are not sufficient to execute a
' task easily, such as when taking a screenshot of an entire webpage while this can
' be done relatively easy using external JS libraries. This function shall enables
' massive possibility of extended features.
' https://stackoverflow.com/a/57572107
' Updated: 23/03/2023
'----------------------------------------------------------------------------------
' Note: the script library should only be added after the page is fully loaded. If
' it is loaded when the HTML document is refereshing, it risks adding the script to
' the old HTML document which will be replaced by the new one causing the script to
' be inexistent and resulting in "ReferenceError" due to its functions not defined.
'----------------------------------------------------------------------------------
    
   'Check if already loaded, if so then no need to add again
    If getElementByQuery("script[src='" & strURL & "']").varResult = "null" Then
        printMsg "The HTML already contains the JS Library: """ & strURL & """"
        Exit Sub
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
End Sub

Public Sub snapPage(Optional fileName As String = "mysnapshot", Optional targetElement As CDPElement)
'----------------------------------------------------------------------------------
' Capture a screenshot of the entire HTML page from the top to bottom. The method
' employs html2canvas external library is which loaded via addJsLibrary. To snap
' only a certain element within the page, indicate the same with the targetElement
' argument (default is "document.body" for the entire page).
' Original source: https://stackoverflow.com/a/51478809
' Updated: 14/04/2023
'----------------------------------------------------------------------------------
' Note : Avoid saving the fileName with . as Chromium browser will likely reject
'        the file and does not allow it to be downloaded. This function assumes the
'        download folder is the default download folder.
'----------------------------------------------------------------------------------

   'Add the html2canvas library location ---------------------------------
    jsAddLib "https://html2canvas.hertzen.com/dist/html2canvas.min.js"
   '----------------------------------------------------------------------

   'Clean up the Downloads folder ---------------------------------------------
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(Environ("USERPROFILE") & "\Downloads\").Files
        f.Delete
    Next f
   '---------------------------------------------------------------------------

   'Run screen capture script
   '(Javascript Original)
    '    html2canvas(document.body, { scrollY: -window.scrollY }).then(function (canvas) {
    '        var link = document.createElement('a');     //create an element tag a
    '        link.href = canvas.toDataURL();             //save img data to href
    '        link.download = 'file-name.png';            //file name
    '        document.body.appendChild(link);            //add link to page
    '        link.click();                               //click it for downloading
    '        document.body.removeChild(link);            //then remove it
    '    });
    Dim callString As String, targetJS As String
    If targetElement Is Nothing Then targetJS = "document.body" Else targetJS = targetElement.varName
    callString = "html2canvas(" & targetJS & ", { scrollY: -window.scrollY }).then(function (canvas) { var link = document.createElement('a'); link.href = canvas.toDataURL(); link.download = '" & fileName & "'; document.body.appendChild(link); link.click(); document.body.removeChild(link); });"
    Dim isRetrying As Boolean: isRetrying = False
    Dim timerStart As Single, timerOut As Long, strReturn As String
    timerStart = Timer: timerOut = 10   'seconds
    Do
        strReturn = jsEval(callString, Not isRetrying)
        If InStr(strReturn, "html2canvas is not defined") = 0 Then Exit Do   'if html2canvas is loaded, move on, else retry connecting with html2canvas
        If Timer - timerStart > timerOut Then printMsg "Timeout waiting for html2canvas to response.", doRaiseError:=True
        If Not isRetrying Then printMsg "Retrying html2canvas execution...": sleep 0.1
        isRetrying = True    'so that jsEval is not going to spam immediate window
    Loop

   'Grab the file from Downloads ------------------------------------------------------------------------------------------------
   'Auto remove the dot char from original name to prevent download error
    fileName = Replace(fileName, ".", "") & ".png"    'default file extension to .png
    Dim filePath As String
    filePath = Environ("USERPROFILE") & "\Downloads\" & fileName
    timerStart = Timer: timerOut = 60   'seconds
    While Dir(filePath) = ""   'Wait till image file is created
        DoEvents
        If Dir(filePath & ".crdownload") <> "" Then filePath = filePath & ".crdownload"   'Sometimes chrome just added this tail
        If Timer - timerStart > timerOut Then printMsg "Timeout waiting for the screenshot download.", doRaiseError:=True
    Wend
   '-----------------------------------------------------------------------------------------------------------------------------

   'Save to destination folder ----------------------------------
    Dim destnPath As String
    destnPath = thisApp.Path & "\Saved Screens\"
    If Dir(destnPath, vbDirectory) = "" Then MkDir (destnPath)
    fso.GetFile(filePath).Copy destnPath & fileName, True
   '-------------------------------------------------------------

    printMsg "snapPage sucessfully saved the image to """ & destnPath & fileName & """"
End Sub

Public Sub printMsg(strMsg As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean = False)
'----------------------------------------------------------------------------------
' A custom printMsg that besides printing to the Immediate Window, it will
' also print to a dated log file. This is useful at user level where
' the Immediate Window is not easily accessible. doRaiseError will also raise an
' error aside from printing the error message to the log file.
' Updated: 20/03/2023
'----------------------------------------------------------------------------------
    Dim strFormattedMsg As String
    If isHeader Then strFormattedMsg = String(100, "-") & vbNewLine & strMsg & vbNewLine & String(100, "-") _
    Else strFormattedMsg = Format(Now, "hh:mm:ss") & " | " & brID & " | " & strMsg

    Debug.Print strFormattedMsg  'standard output to Immediate Window
    
    Dim logPath As String, logName As String
    logPath = thisApp.Path & "\Logs\"

    logName = "log" & UCase(Format(Now, "ddMMMyy")) & ".cdp.txt"
    If Dir(logPath, vbDirectory) = "" Then MkDir logPath
    Open logPath & logName For Append As #1
    Print #1, strFormattedMsg
    Close #1

On Error Resume Next 'GCUser99
    If doRaiseError Then Err.Raise 900, Description:=strMsg
End Sub

Public Sub printTabs()
'------------------------------------------------------------------
' Print all current tabs info. Useful for debugging.
'------------------------------------------------------------------
    Dim targets As Object, target_ As Object, i As Long
    Dim results As Scripting.Dictionary
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
End Sub

Public Sub printParams()
'------------------------------------------------------------------
' Print all parameters of the current session for debuggings.
'------------------------------------------------------------------

    printMsg "---------------------------------------------------"
    printMsg brID & " Information"
    printMsg "---------------------------------------------------"

    printMsg "Session ID    | sessionID | " & sessionID
    printMsg "Last ID       | lngLastID | " & lngLastID
    printMsg "Window Handle | brHWnd    | " & brHWnd
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
End Sub

Public Sub sleep(Optional seconds As Double = 0.5)
'------------------------------------------------------------------
' Custom sleep function. Sleep by 0.5s by default.
' Useful for a quick necessary pause when needed.
'------------------------------------------------------------------

    Const baseUnit As Long = 1000    'ie. millisecs

    sleep2 seconds * baseUnit
    DoEvents
End Sub

Public Sub activate()
'----------------------------------------------------------------------------------
' Switch focus to the current tab being attached to under the current BRID.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------
    Dim params As New Scripting.Dictionary
    params("targetId") = CStr(getTargetID)
    invokeMethod "Target.activateTarget", params
End Sub

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

'=============================================================================
' PRIVATE FUNCTIONS
'=============================================================================

Private Function sendMessage(strMessage As String, Optional objAllMessages As Scripting.Dictionary, Optional dbgMsg As Boolean = True) As Scripting.Dictionary
'--------------------------------------------------------------
' Send message to CDP and check if the message is received
' yet by the Chromium browser.
' Updated: 03/05/23 - Added soft-reset on peeknamepipe error
'          22/03/23 - Added error handling for better debugging
'--------------------------------------------------------------
    
    Dim intRes As Long
    Dim strRes As String
    Dim lngCurrentId As Long
    Dim results As Scripting.Dictionary       'to store a collection of IDs obtained from CDP
            
    lngCurrentId = lngLastID
    lngLastID = lngLastID + 1       'We increase the global ID counter
            
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
    Dim timerStart As Double, timerOut As Long
   
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
    
ifError:
                  
    If Not dbgMsg Then Exit Function    'Allow the option to bypass error if intended
                  
   'Record errmsg and print helpful infos
    Dim errorLine As Long, errorDesc As String, fso As Object
    errorLine = Erl
    errorDesc = Err.Description
    printMsg "sendMessage error at line " & errorLine & ": """ & errorDesc & """"
    printMsg "strMessage used by sendMessage: " & strMessage
    If InStr(Err.Description, "PeekNamedPipe") Then
        printMsg "Possible reasons: (1) the browser is not actually running (ie. call string failure); (2) the browser window was accidentally closed; (3) the CDP Pipe connection to the browser was cut-off somehow."
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFile Environ("Temp") & "\cdpid.txt"   'soft reset
    End If
    
    Err.Raise 900, , "sendMessage error at line " & errorLine & ": """ & errorDesc & """"
        
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
    Dim boolFound As Boolean
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
    Dim errorLine As Long, errorDesc As String
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
    Err.Raise 900, , "attachToSession error at line " & errorLine & ": """ & errorDesc & """"
    
End Function


Private Function searchNull() As Long
'------------------------------------------------------------
' CDP messages received from chrome are null-terminated
' It seemed to me you cant search for vbnull in a string
' in vba. Thats why i re-implemented the search function
'------------------------------------------------------------
'yes you can search for vbNullChar!
'SearchNull = InStr(strBuffer, vbNullChar)

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


Private Function invokeError(result As Scripting.Dictionary) As Boolean
'-----------------------------------------------------------------------------------
' Check if InvokeMethod fails and what the error message returned is.
' Error message will be shown via printMsg
' Credit: https://github.com/PerditionC/VBAChromeDevProtocol
'-----------------------------------------------------------------------------------

    If result Is Nothing Then
        Dim ErrorCode As Long, ErrorMessage As String, ErrorData As Variant
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

    'Determine current tab title, if error: the tab is a blank tab
    Dim sessionTitle As String, windowTitle As String
    sessionTitle = jsEval("document.getElementsByTagName(""title"")[0].innerHTML", False)
    If InStr(sessionTitle, "TypeError") Then sessionTitle = "about:blank"

    'Obtain handles to the correct browser window
    'Both Edge and Chrome use the same class chromeWindowClass
    hWnd = FindWindowEx(0&, 0&, StrPtr(chromeWindowClass), 0&)        'To return Chrome / Edge window handles
    Do
                
        'Get Window Title based on specified handles
        'Requires string buffering declaration as string
        bufferString = String(1024, " ")
        titleLen = GetWindowText(hWnd, StrPtr(bufferString), Len(bufferString))
        windowTitle = Left(bufferString, titleLen)
    
        'Verify if matching with the sessionTitle and is really an edge/chrome window
        If windowTitle <> "" Then If InStr(windowTitle, sessionTitle) Then Exit Do
    
        hWnd = FindWindowEx(0&, hWnd, StrPtr(chromeWindowClass), 0&)      'Continue with the next handle value
    
    Loop Until hWnd = 0 ' no more matching windows

    getHandle = hWnd
End Function

Private Sub cleanUpSessions()
'------------------------------------------------------------------
' Provides cleaning before firing up a new session to prevent
' pipe error
'------------------------------------------------------------------

   'Clean up all running sessions
   Dim p As Object
    On Error Resume Next
    For Each p In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
        If p.name = brType Then
            p.Terminate
        End If
        DoEvents
    Next
    On Error GoTo 0
    Err.Number = 0

   'Also clean up all prev CDP
   Dim fso As Object, tempFolder As Object, fName As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tempFolder = fso.GetFolder(Environ("Temp"))
    fName = tempFolder & "\cdpid.txt"
    If Dir(fName) <> "" Then fso.DeleteFile fName

   'Clean core params for new session
    lngLastID = 0   'To make a clean slate for new session
    sessionID = ""  'Likewise

    printMsg "cleanUpSessions completed and previous CdpSessionID removed"
End Sub


Private Sub setCrashStateNormal()
'------------------------------------------------------------------
' Set Crashed state to Normal to prevent the "Restore Sessions?"
' popup from appearing every time the browser is started after
' being closed by browser.quit method.
' Requires user data dir profile stored in private var brUser and
' the browser type stored in private variable brType.
'------------------------------------------------------------------

   'Mark 'Normal' on the preference file to prevent Restore Page popup on start
   'This file locates at %localappdata%/Google/User Data/Default/Preferences
   Dim filePath As String, prefFolder As String, prefFile As Object, strContents As String, strNewContents As String
    If brUser = "Default" Then prefFolder = brUser Else prefFolder = brUser & "\Default"
    Select Case brType
        Case brChrome: filePath = Environ("LOCALAPPDATA") & "\Google\Chrome\User Data\" & prefFolder & "\Preferences"
        Case brEdge: filePath = Environ("LOCALAPPDATA") & "\Microsoft\Edge\User Data\" & prefFolder & "\Preferences"
    End Select

    Dim fso As New Scripting.FileSystemObject
    If Not fso.FileExists(filePath) Then Exit Sub
    Set prefFile = fso.OpenTextFile(filePath, ForReading)
    strContents = prefFile.ReadAll
    strNewContents = Replace(strContents, "Crashed", "Normal", Compare:=vbBinaryCompare)   'ie. Exact match only
    prefFile.Close

    Set prefFile = fso.OpenTextFile(filePath, ForWriting)
    prefFile.Write strNewContents
    prefFile.Close

    printMsg "setCrashStateNormal completed"
End Sub


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
' Updated: 02/03/2023
'----------------------------------------------------------------------------------
    Dim fso As Object, tempFolder As Object, fName As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tempFolder = fso.GetFolder(Environ("Temp"))
    fName = tempFolder & "\cdpid.txt"
    If Dir(fName) = "" Then Exit Function

   'Check if CdpSessionID is not today
    If DateDiff("d", FileDateTime(fName), Now) > 0 Then fso.DeleteFile fName _
    Else getSerialString = fso.OpenTextFile(fName, ForReading, True).ReadAll
End Function


Private Function newVarID() As String
'----------------------------------------------------------------------------------
' Create a unique varID for element reference with CDPElement. If the ID is not
' unique, it is possible that new element assignment will override existing ones
' under the same ID by chance if working on a large number of elements.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------

   'Init the ID collection
    If varIDs Is Nothing Then Set varIDs = New Scripting.Dictionary
    Dim newID As String
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
'----------------------------------------------------------------------------------
    
    Select Case True
        Case InStr(LCase(strBrowserType), "chrome"): brType = brChrome
        Case InStr(LCase(strBrowserType), "edg"): brType = brEdge         'On edge it is just "Edg"
        Case Else: printMsg "Unable to getBrowserPath for browser type: " & brType, doRaiseError:=True
    End Select

   'Return the path
    On Error Resume Next
    Dim objWSHost As Object, reg As Object, strKeyPath As String
    Set reg = GetObject("winmgmts://./root/default:StdRegProv")
    Set objWSHost = CreateObject("WScript.Shell")
    strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & brType & "\Path"
    getBrowserPath = objWSHost.regread(strKeyPath) & "\" & brType

    If Err.Number <> 0 Then printMsg Err.Description, doRaiseError:=True
End Function


Private Sub getSessionInfo()
    '----------------------------------------------------------------------------------
    ' Obtain session info on the fly for printParams.
    ' Updated: 22/03/2023
    '----------------------------------------------------------------------------------
   'Activate to getHandle correctly
    activate

   'Get infos
    brURL = url
    brName = name
    brHWnd = getHandle
    brWID = invokeMethod("Browser.getWindowForTarget").Item("windowId")
    brType = invokeMethod("Browser.getVersion").Item("product")
    brVer = Split(brType, "/")(1)
    brPath = getBrowserPath(brType)
    targetID = getTargetID
End Sub


Private Sub Class_Initialize()
'----------------------------------------------------------------------------------
' A quick way to quickly check if an automated session is already active and open.
' If so, automatically attempt to reattach to it.
' Updated: 07/06/2023
'----------------------------------------------------------------------------------
    
   'Determine project host type for compatibility
    Select Case True
        Case InStr(Application.name, "Excel"): Set thisApp = ThisWorkbook
        'Case InStr(Application.Name, "Word"): Set thisApp = ThisDocument
        'Case InStr(Application.Name, "Access"): Set thisApp = CurrentProject
    End Select

   'Assign an objID for easy debug identification
    brID = "BRID" & Format(Rnd * 1000, "000")
    
   'Check for SessionID created
    Dim strSerialized As String: strSerialized = getSerialString
    If strSerialized = "" Then Exit Sub
   
   'Attempt to reattach and check if successful
    deserialize strSerialized
    
    If Not isLive Then
        lngLastID = 0   'To make a clean slate for new session
        sessionID = ""  'No need to clean targetID as it will be retrieved anyway
        printMsg "Failed to reattach using the previous session parameters"
        Exit Sub
    End If
            
   'Obtain instance useful info if already live
    printMsg "Browser ID : " & brID & vbNewLine & _
             "Tab ID     : " & targetID, isHeader:=True
    printMsg brID & " successfully attached to Tab ID " & targetID
    getSessionInfo
        
End Sub
 
'===============================================================================================
' Experimental Functions
'===============================================================================================
