Attribute VB_Name = "Demo"
'===================================================================================================
' Automating Chromium-Based Browsers with Chrome Dev Protocol API and VBA
'---------------------------------------------------------------------------------------------------
' Author(s)   :
'       ChrisK23 (Code Project)
' Contributors:
'       Long Vh (long.hoang.vu@hsbc.com.sg)
' Last Update :
'       07/06/22 Long Vh: corrected typos in comments + more examples
'       03/06/22 Long Vh: codes edited + notes added + added extensive comments for HSBC colleagues
' References  :
'       Microsoft Scripting Runtime
' Notes       :
'       The framework does not need a matching webdriver as this is not a webdriver-based API.
'       This module includes a few examples of automating browsers using CDP. For the
'       engine codes, refer to the class modules clsBrowser, clsCore, and clsJsConverter.
'       For original examples, refer to Chris' article on CodeProject:
'       https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA
'===================================================================================================
 
 
Sub runEdge()
'------------------------------------------------------
' This is an example of how to use the browser classes
' This demo tries to access a webpage of a famous movie
' and retrieve its current view count.
'------------------------------------------------------
 
   'Start Browser
   'If no browser name is indicated, chrome is started by default.
   'If edge is to be started, use this instead: objBrowser.start "edge"
   'Homepage has been disabled to speed up by default.
   'To skip cleaning active sessions, set cleanActiveSession to False.
   'This will make browser starts faster but at the risk of pipe error if
   'there are other chrome instances already running.
    Dim objBrowser As New clsBrowser
    objBrowser.start "edge", cleanActiveSession:=True
        
   'By default, the new window is minimized, use .show to bring it out
    objBrowser.show
    
   'Navigate and wait
   'The wait method, if till argument is omitted, will by default wait until ReadyState = complete
    objBrowser.navigate "https://www.livingwaters.com/movie/the-atheist-delusion/"
    objBrowser.wait till:="interactive" 'only need to wait until page is interactable. Refer to definition for other options
    
   'Get view count
    viewCount = objBrowser.jsEval("document.evaluate(""//h3[contains(., 'Total Views')]/*[1]"", document).iterateNext().innerText")
    objBrowser.jsEval "alert(""This free movie has already reached " & viewCount & " views! Wow!"")"
    
End Sub
 
 
Sub runChrome1()
'--------------------------------------------------------------------------------
' This example prints the serialized json string of the running instance.
' This json string can then be parsed onto the next example to attach to the same
' running browser window.
' The string is saved to Cell A1 of the Excel table.
'--------------------------------------------------------------------------------
   
    Dim chrome As New clsBrowser
    
    chrome.start
    chrome.show
    chrome.navigate "https://google.de"
                
    Cells(1, 1) = chrome.serialize()
    MsgBox "The serialized string has been saved to A1."
    
End Sub
 
 
Sub runChrome2()
'--------------------------------------------------------------------------------
' Read serialized string from Cell(1, 1) of the Excel table and try to
' attach to the running instance with the same serial.
'--------------------------------------------------------------------------------
 
    Dim objBrowser As New clsBrowser
            
   'Attempt to decipher the serialized string to attach to the running session
    objBrowser.deserialize Cells(1, 1)
    If Not objBrowser.isLive Then err.Raise -900, Description:="Unable to find the session with the current serial in A1."
    
   'If found, confirm to user
    Cells(1, 1) = "Open VBE screen to see the Demo module"
    objBrowser.jsEval "alert(""Found the target session!"")"
    
End Sub
 
 
Sub runHidden()
'---------------------------------------------------------------------------------
' Demonstrate background running of an automated session.
' This demo will try to open Google in the background, then search for an article
' of CodeProject and retrieve its vote count. Once done, it will prompt a message
' to display the browser window.
' It is recommended to make Immediate Window visible so that you can see the
' activity that is running in the background.
' To confirm the result, you can perform the following steps:
'   1. Go to Google.com
'   2. Type "automate edge vba" and click Search
'   3. Click on the first result to reach the CodeProject's article
'   4. The vote count is seen there.
'---------------------------------------------------------------------------------
       
    Dim chrome As New clsBrowser
   
   'Start and hide
    chrome.start
    chrome.hide
    
   'Perform automation in the background
    chrome.navigate "https://google.com"
    chrome.wait till:="interactive"
    chrome.jsEval "document.getElementsByName(""q"")[0].value=""automate edge vba"""
    chrome.jsEval "document.getElementsByName(""q"")[0].form.submit()"
    chrome.wait till:="interactive"
    chrome.jsEval "document.evaluate("".//h3[text()='Automate Chrome / Edge using VBA - CodeProject']"", document).iterateNext().click()"
    chrome.wait till:="interactive"
    voteCount = chrome.jsEval("ctl00_RateArticle_VoteCountNoHist.innerText")
    
   'Confirm result and display
    userChoice = MsgBox("Automation completed. Current vote counts: " & voteCount & ". Do you want to see the window?", vbYesNo)
    If userChoice = vbYes Then chrome.show Else chrome.quit
    
End Sub


Sub runInstances()
'---------------------------------------------------------------------------------
' Demonstrate the automation of multiple browser instances concurrently using
' different user profiles. This is useful when the automation needs to automate
' multiple sessions at the same time, search as one session for form input and
' another for data collection from various places. The pros of this method is each
' instance can start with different settings as needed. In this case, each
' instance runs on an unique user profile.
'---------------------------------------------------------------------------------

    Dim chrome1 As New clsBrowser
    Dim chrome2 As New clsBrowser
    Dim chrome3 As New clsBrowser
    
   'Start multiple chrome instances at the same time
   'Multiple profiles are needed to avoid PeekPipeNamed error
   'cleanActiveSession=False on subsequent .start to prevent accidentally closing all windows
    chrome1.start cleanActiveSession:=True, userProfile:="User G"
    chrome2.start cleanActiveSession:=False, userProfile:="User H"
    chrome3.start cleanActiveSession:=False, userProfile:="User I"

   'Navigate each instance to a different url
    chrome1.navigate "https://google.com"
    chrome2.navigate "https://yahoo.com"
    chrome3.navigate "https://bing.com"

   'Reposition and display the windows
    chrome1.show 0, 0, 1000, 700
    chrome2.show 0, 100, 1000, 700
    chrome3.show 0, 200, 1000, 700
    
End Sub


Sub runTabs1()
'--------------------------------------------------------------------------
' Demonstrate the automation of multiple tabs in a single browser instance.
' Here, the switchTo method handles automation focus across the tabs.
' switchTo directs the current automation pipe to the new session ID.
' Similar to the runInstances example but this is with multiple tabs in
' the same instance instead.
'--------------------------------------------------------------------------
    
    Dim chrome As New clsBrowser
    chrome.start
    chrome.show
    
   'Get Session IDs for each tab
    Dim sIdTab1 As String: sIdTab1 = chrome.SessionID
    Dim sIdTab2 As String: sIdTab2 = chrome.newTab
    Dim sIdTab3 As String: sIdTab3 = chrome.newTab
       
   'Automate Tab 1
    chrome.switchTo sIdTab1
    chrome.navigate "https://google.com"
    
   'Automate Tab 2
    chrome.switchTo sIdTab2
    chrome.navigate "https://yahoo.com"
    
   'Automate Tab 3
    chrome.switchTo sIdTab3
    chrome.navigate "https://bing.com"
    
   'Resize to complete
    chrome.show 0, 20, 1000, 700

End Sub


Sub runTabs2()
'--------------------------------------------------------------------------
' Demonstrate the automation of multiple tabs in a single browser instance.
' Here, each tab is neatly assigned to a new tab object that is initiated
' by .clone to copy the default browser's pipe parameters. This new tab
' object is then given a Session ID corresponding to the new open tab. This
' is like having 3 automation instances running together like runInstances.
' However, each tabs will have to share the same start settings, unlike
' the case of runInstances where each instance can be setup with a different
' settings to each other.
' Note: because VBA does not allow copying object without referring to it,
' clone has been added to the class function of clsBrowser for this use.
'--------------------------------------------------------------------------
    
    Dim chrome As New clsBrowser
    chrome.start
    chrome.show
    
   'Create and assign tabs
    Dim tab1 As New clsBrowser
    Dim tab2 As New clsBrowser
    Dim tab3 As New clsBrowser
    Set tab1 = chrome                                                           'The first tab is open by default after .start
    Set tab2 = chrome.clone: tab2.SessionID = chrome.newTab(newWindow:=True)    'newWindow: open tab as a new window instead of a tab
    Set tab3 = chrome.clone: tab3.SessionID = chrome.newTab(newWindow:=True)
    
   'Automate each tabs
    tab1.navigate "https://google.com"
    tab2.navigate "https://yahoo.com"
    tab3.navigate "https://bing.com"
    
   'Resize to complete
    tab1.show 0, 10, 1000, 700
    tab2.show 0, 45, 1000, 700
    tab3.show 0, 90, 1000, 700

End Sub


Sub runTabs3()
'--------------------------------------------------------------------------
' This example demonstrates:
' 1. The use of advanced arguments feature added by Long Vh to v1.0 to
'    allow the choice of additional settings for the automation pipe. See
'    https://peter.sh/experiments/chromium-command-line-switches/
' 2. The xPath technique to directly modify the current HTML element
'    so that it will behave in a new way that it was not so before.
' 3. The technique employed to integrate the new tab open spontaneously
'    by interaction with the webpage (instead of using .newTab) into the
'    automation pipe for further processing on the new tab.
'--------------------------------------------------------------------------
    
   'Init browser with custom arguments
    Dim chrome As New clsBrowser
    chrome.start addArguments:="--disable-popup-blocking"    'The disable-popup-blocking argument is needed to allow opening link in a new tab
    chrome.maximized
    
   'Perform standard google search
    chrome.navigate "https://google.com"
    chrome.wait till:="interactive"
    chrome.jsEval "document.getElementsByName(""q"")[0].value=""newstarget.com"""
    chrome.jsEval "document.getElementsByName(""q"")[0].form.submit()"
    chrome.wait till:="interactive"
    
   'Google search result returns links that open in the same tab window
   'For this demonstration, we need to make it open in a new tab window instead
    chrome.jsEval "el = document.evaluate("".//a[contains(@href, 'https://www.newstarget.com/')]"", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
    chrome.jsEval "el.setAttribute(""target"",""_blank"")"      'Modify the element attribute to open in a new tab instead
    chrome.jsEval "el.click()"                                  'Click the link, a new tab will be spontaneously open
    
   'Since the new tab is spontaneously open and not with the .newTab method (see runTabs1 & runTabs2 for examples),
   'we need to find its Session ID and attach the new tab to the automation pipe for further working
    Dim newSessionId As String
    newSessionId = chrome.getNewTab     'retrieve the Session Id of the new tab
    chrome.switchTo newSessionId        'switch focus to the new tab
    chrome.wait                         'without argument, the method will wait until full page loaded (ie. till:="complete")
     
   'Feed the top news title for today
    chrome.jsEval "context = document.querySelector(""#FeaturedA"")"
    chrome.jsEval "topTitle = context.querySelector(""div[class='Headline']"").innerText"
    chrome.jsEval "alert(""Today's Top Headline is... \n\n"" + topTitle.toUpperCase())"

End Sub
