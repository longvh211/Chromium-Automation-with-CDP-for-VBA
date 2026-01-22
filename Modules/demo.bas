'===================================================================================================
' Automating Chromium-Based Browsers with Chrome Dev Protocol API and VBA
'---------------------------------------------------------------------------------------------------
' Author(s)   :
'       ChrisK23 (Code Project)
' Contributors:
'       Long Vh (long.hoang.vu@hsbc.com.sg)
' Last Update :
'       22/01/26 Long Vh: added demoMultiProfileOperation and demoReattachment examples
'       07/01/26 Long Vh: update the sub procedures to show case the new .notify function
'       27/04/23 Long Vh: made many improvements with v2.5 to make methods even more intuitive.
'       07/06/22 Long Vh: corrected typos in comments + more examples
'       03/06/22 Long Vh: codes edited + notes added + added extensive comments for HSBC colleagues
' References  :
'       Microsoft Scripting Runtime
' Notes       :
'       The framework does not need a matching webdriver as this is not a webdriver-based API.
'       This module includes a few examples of automating browsers using CDP. For the
'       engine codes, refer to the class modules CDPBrowser, CDPCore, CDPElement, and CDPJConv
'       For original examples, refer to Chris' article on CodeProject:
'       https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA
'       For the latest update of the CDP Framework by Long Vh:
'       https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA
'===================================================================================================
 
Sub runEdge()
'------------------------------------------------------
' This is an example of how to use the browser classes
' This demo tries to access a webpage of a famous movie
' and retrieve its current view count.
'------------------------------------------------------
 
   'Start Browser
   'If no browser name is indicated, chrome is started by default.
   'Homepage has been disabled to speed up by default.
   'To skip cleaning active sessions, set cleanActive to False.
   'This will make browser starts faster but at the risk of pipe error if
   'there are other chrome instances already running.
   'If reAttach = False, .start will not automatically try to reattach
   'to previous instances open by CDP but will start a brand new instead.
    Dim edge As New CDPBrowser
    edge.start "edge", cleanActive:=True
 
   'Navigate and wait
   'If till argument is omitted, will by default wait until ReadyState = complete
    edge.navigate "https://livingwaters.com/movie/the-atheist-delusion/", isInteractive
 
   'Get view count via the new notify method
    viewCount = edge.getElementByQuery("[data-id='4b9a4b19']").innerText
    edge.notify "This free movie has already reached " & viewCount & " views! Wow!"
 
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
 
    Dim chrome As New CDPBrowser
 
   'Start and hide
    chrome.start
    chrome.hide
 
   'Perform automation in the background
    chrome.navigate "https://google.com", isInteractive
    chrome.getElementByQuery("[name='q']").value = "automate edge vba"
    chrome.getElementByQuery("[name='q']").submit
    
   'Click the target result link
    chrome.getElementByXPath("//h3[text()='Automate Chrome / Edge using VBA']").click
    
   'Get the vote count only once the target element appears on screen
   'The onExists method is needed as this element appears after ReadyState = "complete"
    voteCount = chrome.getElementByID("ctl00_RateArticle_VoteCountNoHist").onExist.innerHTML
    
   'Confirm result and display
    userChoice = MsgBox("Automation completed. Current vote counts: " & voteCount & ". Do you want to see the window?", vbYesNo)
    If userChoice = vbYes Then chrome.show Else chrome.quit
    
End Sub
 
Sub runTabsAsOne()
'--------------------------------------------------------------------------
' Demonstrate the automation of multiple tabs in a single browser instance.
' Similar to the runInstances example but this is with multiple tabs in
' the same instance instead.
'--------------------------------------------------------------------------
 
    Dim chrome As New CDPBrowser
    chrome.start
    chrome.show
    
   'Automate Tabs
    chrome.url = "google.com"   'or [chrome.navigate "google.com"]
    chrome.newTab "sg.yahoo.com"
    chrome.newTab "bing.com"
 
   'Resize to complete
    chrome.show xywh:="0 20 1000 700"
 
End Sub
 
Sub runTabsAsMany()
'-------------------------------------------------------------------------------
' Demonstrate the automation of multiple tabs in a single browser instance.
' This is like having 3 automation instances running together like runInstances.
' However, each tab will have to share the same start settings, unlike
' the case of runInstances where each instance can be setup with a different
' settings to each other.
'-------------------------------------------------------------------------------
 
    Dim chrome As New CDPBrowser
    chrome.start
    chrome.show
 
   'Create and assign tabs
    Dim tab1 As New CDPBrowser                   'The keyword "New" is a must
    Dim tab2 As New CDPBrowser
    Dim tab3 As New CDPBrowser
    Set tab1 = chrome                            'The first tab is open by default after .start
    Set tab2 = chrome.newTab(newWindow:=True)    'newWindow: open tab as a new window instead of a tab
    Set tab3 = chrome.newTab(newWindow:=True)
 
   'Automate each tabs
    tab1.navigate "google.com"
    tab2.navigate "sg.yahoo.com"
    tab3.navigate "bing.com"
 
   'Resize to complete
    tab1.show xywh:="0 10 1000 700"
    tab2.show xywh:="0 45 1000 700"
    tab3.show xywh:="0 90 1000 700"
 
End Sub
 
Sub runNewTab()
'--------------------------------------------------------------------------
' This example demonstrates:
' 1. The use of advanced arguments feature added by Long Vh to
'    allow the choice of additional settings for the automation pipe. See
'    https://peter.sh/experiments/chromium-command-line-switches/
' 2. The xPath technique to directly modify the current HTML element
'    so that it will behave in a new way that it was not so before.
' 3. The technique employed to integrate the new tab open spontaneously
'    by interaction with the webpage (instead of using .newTab) into the
'    automation pipe for further processing on the new tab.
'--------------------------------------------------------------------------
 
   'Init browser with custom arguments
    Dim chrome As New CDPBrowser
    chrome.start addArgs:="--disable-popup-blocking"    'The disable-popup-blocking argument is needed to allow opening link in a new tab
    chrome.show asMaximized
    
   'Perform standard google search
    chrome.navigate "https://google.com"
    chrome.getElementByQuery("[name='q']").value = "newstarget.com"
    chrome.getElementByQuery("[name='q']").submit
 
   'Google search result returns links that open in the same tab window
   'For this demonstration, we need to make it open in a new tab window instead
    Dim targetElement As CDPElement
    Set targetElement = chrome.getElementByXPath(".//a[contains(@href, 'https://www.newstarget.com/')]")
    targetElement.setAttribute "target", "_blank"   'Modify the element attribute to open in a new tab instead
    targetElement.click                             'Click the link, a new tab will be spontaneously open
 
   'Use getTabNew to quickly refer to the next newly open tab
    Dim targetTab As New CDPBrowser
    Set targetTab = chrome.getTab
    targetTab.wait
 
   'Feed the top news title for today
    firstTitle = targetTab.getElementByQuery("div[class='Headline']").innerText
    targetTab.notify "Top popular headline for the day is """ & firstTitle & """."
 
End Sub
 
Sub runIFrame()
'--------------------------------------------------------------------------
' This example demonstrates the CDP Framework v2.5 getIFrame technique for
' accessing iFrame element intuitively, an improvement over 1.0:
' 1. The use of App Mode via appUrl argument of the .start method.
' 2. The use of getIframe to easily access iFrame elements on the web page.
' 3. Working with a complex web design where nested iFrames are employed.
'--------------------------------------------------------------------------
    
    Dim demoUrl As String
    demoUrl = "https://www.w3schools.com/html/tryit.asp?filename=tryhtml_iframe_height_width"
    
    Dim chrome As New CDPBrowser
    chrome.start appUrl:=demoUrl
    
    Dim iFrame1 As CDPElement
    Dim iFrame2 As CDPElement
    Set iFrame1 = chrome.getElementByID("iframeResult").getIFrame
    Set iFrame2 = iFrame1.getElementByQuery("iframe[title='Iframe Example']").getIFrame
    
    txt = iFrame2.getElementByQuery("h1").innerText
    chrome.notify "Retrieved text from the iFrame: """ & txt & """"
    
End Sub
 
Sub getSnapShot()
'--------------------------------------------------------------------------
' This example demonstrates the easy handling of capturing a screenshot of
' the current page under CDP method. The second argument of the snapPage
' method can be set to True to capture the entire page or to False (default)
' to capture only the current view section of the page.
'--------------------------------------------------------------------------
 
    Dim demoUrl As String
    demoUrl = "https://www.google.com/search?q=1sgd+to+vnd"
    
    Dim chrome As New CDPBrowser
    chrome.start                'not App Mode as sometimes Chrome App Mode does not allow file downloading
    chrome.navigate demoUrl
    
   'Snap a portion of the page based on the element indicator
   'If the second argument is omitted, snapPage will snap the entire page
    Dim fileName As String
    fileName = ThisWorkbook.Path & "\todaySGDvsVND.png"
    chrome.snapPage fileName 'chrome.snapPage(fileName, True) to capture the entire page instead
    chrome.notify "Screenshot captured under " & fileName
 
End Sub
 
Sub fillReactForm()
'-------------------------------------------------------------------------
' This example demonstrates the power of 2.6 on working natively
' with React form fields, which are notoriously complex to automate
' due to the fact that React form uses its own internal event handlings.
' The demo aims to:
' 1. Fill in the name field on the page.
' 2. Press submit.
' 3. If the field input is recognized by React, alert will tell its value.
' Updated: 07/01/26: .sendKeys has been replaced with .sendString
'-------------------------------------------------------------------------
 
    Dim demoUrl As String
    demoUrl = "https://cdpn.io/gaearon/fullpage/VmmPgp?anon=true&editors=0010&view="
    
    Dim chrome As New CDPBrowser
    chrome.start
    chrome.navigate demoUrl
        
   'Get the target fields
    Dim ip As CDPElement
    Dim sb As CDPElement
    Set ip = chrome.getElementByID("result").getIFrame.getElementByQuery("input[type='text']")
    Set sb = chrome.getElementByID("result").getIFrame.getElementByQuery("input[type='submit']")
        
   'This traditional input method will fail as this is a React field
    chrome.jsEval ip.varName & ".value = 'TEST1'"
    chrome.jsEval ip.varName & ".dispatchEvent(new Event('input', { bubbles: true, simulated: true }))"
    sb.click 'you will not see "TEST1" in the alert result
 
   'This will succeed by using 2.6-enhanced .value property
    ip.value = "TEST2" '.value property is now overloaded with a smart React field detection & inputing
    sb.click
    
   'This will succeed as it mimicks sending raw keys but to a specific element
    ip.sendString "TEST3"
    sb.click
 
End Sub

Sub switchMain()
'---------------------------------------------------------------
' This example demonstrate the use of argument setMain to switch
' the main session tab to another tab so that future
' reattachment will hook this tab directly. This is useful if
' the main tab is supposed to be a tab open subsequently during
' the automation process by the target web link. The setMain
' method is preferrable to using "Set chrome = chrome.getTab..."
' because the latter method does not update the serial string
' for future reattachment.
'---------------------------------------------------------------

    Dim chrome As New CDPBrowser
    chrome.start "edge"
    chrome.newTab "google.com", setMain:=True   'the chrome object will now directly refer to the Google tab
    chrome.getTab("about:blank").closeTab       'prior 2.7, the next line will throw an error due to no main-switching mechanism
    chrome.printParams

End Sub

Sub demoMultiProfileOperation()
'----------------------------------------------------------------------------------------
' This example demonstrates a powerful feature of v3.1 called multi-instances operation.
' Under multi-instances, our framework can open browsers as separate independent
' instances; thereby enables advanced automation tactics such as robotic process
' automation and asynchronous operation. The procedure below attempts to open 2 CDP
' instances and runs at the same time asynchronously - something that VBA natively does
' not support. You will be able to see from the Immediate Window that (1) execBot2 is
' started first then execBot1 and (2) both bot operations run simultaneously and bot 1
' finishes first (likely as yahoo.com has less thing to load then finance.yahoo.com)
' even though it is started after bot 2. This implies that thanks to the multi-
' instances framework, we can achieve asynchronous operation.
'
' Without this feature, the closest to this application is to open a CDP session with
' multiple tabs but in that scenario, you can not achieve asynchronous operation as
' Chrome Devtools Protocol is tied to a single user profile so automation on a tab has
' to wait for one another. Additionally, if one tab causes the browser to crash, other
' running tabs will crash as well.
'----------------------------------------------------------------------------------------

    Application.OnTime Now + TimeValue("00:00:01"), "execBot1"
    execBot2

End Sub

Function execBot1()
'----------------------------------------------------------------------------------------
' Refer to the demoMultiProfileOperation
'----------------------------------------------------------------------------------------
    
    Debug.Print Format(Now, "hh:mm:ss") & " execBot1 started."
    
    Dim e1 As New CDPBrowser
    e1.start userProfile:="CDP1"
    e1.navigate "yahoo.com"
    
    Debug.Print Format(Now, "hh:mm:ss") & " execBot1 completed."

End Function

Function execBot2()
'----------------------------------------------------------------------------------------
' Refer to the demoMultiProfileOperation
'----------------------------------------------------------------------------------------

    Debug.Print Format(Now, "hh:mm:ss") & " execBot2 started."

    Dim e2 As New CDPBrowser
    e2.start userProfile:="CDP2"
    e2.navigate "finance.yahoo.com"
    
    Debug.Print Format(Now, "hh:mm:ss") & " execBot2 completed."

End Function

Sub demoReattachmentPart1()
'----------------------------------------------------------------------------------------
' From v3.1, .reattach is necessary to perform reattachment to the current CDP instances
' as each instance is now identified with a unique user profile for multi-instances
' operation. The below procedure starts a new CDP session under profile CDP2. After
' running demoReattachmentPart1, you can run demoReattachmentPart2 to see the correct
' way of applying .reattach to the CDP2 session.
'----------------------------------------------------------------------------------------

    Dim c As New CDPBrowser
    c.start userProfile:="CDP2"
    c.navigate "google.com"

End Sub


Sub demoReattachmentPart2()
'----------------------------------------------------------------------------------------
' See notes in demoReattachmentPart1
'----------------------------------------------------------------------------------------

    Dim c As New CDPBrowser
    If c.reattach("CDP2") = True Then c.navigate "wikipedia.com" _
    Else Debug.Print "Failed to reattach. Perhaps the reattach profile CDP2 is incorrect?"

End Sub
