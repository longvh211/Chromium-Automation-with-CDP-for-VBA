VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================================================
' CDP HTML Element Class
'------------------------------------------------------------------------------------------------
' Author(s)   :
'       Long Vh (long.hoang.vu@hsbc.com.sg)
' Last Update :
'       02/03/23 Long Vh: - Added the element class for easy object reference.
' Description :
'       This class helps to simplify HTML element object referencing task using CDP. This shall
'       enables the development of 'getElementBy' methods that are commonly used for automating
'       browser sessions. Without this class, all interaction with the HTML elements will have
'       to be parsed via the jsEval method which can be counter-intuitive.
'================================================================================================
Option Explicit

Private main As CDPBrowser
Private varIDs As Scripting.Dictionary
Private varID As String
Private varJS As String
Private varType As String
Private Const printDbgMsg = True
 
 
Public Sub init(targetObjMain As CDPBrowser, targetObjIdColl As Scripting.Dictionary, Optional varID_ As String, Optional varJS_ As String, Optional varType_ As String)
'----------------------------------------------------------------------------------
' Initialize element object that has properties and methods useful for quickly
' interact with the target object in the browser. The varID name can also be used
' directly to the console to interact with the same object in the console. The
' varJS is an alternative method to access the element with the full Javascript
' command that can be parsed to the browser. This alternative method is useful
' for getElementsBy methods wherein not all elements in a returned collection need
' to be accessed immediately and should save considerable amount of execution time.
' varType is used by onExists/onNonExists method to wait until the element is found.
' varType stores the returned result of getElement methods.
' Updated: 02/03/2023
' Note ----------------------------------------------------------------------------
' At initialization, it is ideal to have only varID or varJS indicated. varID is
' most useful when the element is the target element that needs to be interacted
' with while varJS is most useful where only its path of access is needed for now.
' Too many variables in the browser can cause memory leak and slow.
'----------------------------------------------------------------------------------
       
    Set main = targetObjMain        'Needed to use public methods of the active CDPBrowser
    Set varIDs = targetObjIdColl    'Needed to use the same varIDs book from CDPBrowser
    varID = varID_
    varJS = varJS_
    varType = varType_
    
End Sub
 
 
Property Get varName() As String
'----------------------------------------------------------------------------------
' Retrieve the name of this variable initialized. The same name can also be used
' directly in the browser console window if the element is assigned successfully.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------
    
    varName = varID
    
End Property
 
 
Property Get varPath() As String
'----------------------------------------------------------------------------------
' Retrieve the path of this variable initialized. The same name can also be used
' directly in the browser console window if the element is assigned successfully.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------
    
    varPath = varJS
    
End Property
 
 
Property Get varResult() As String
'----------------------------------------------------------------------------------
' Retrieve the type of variable returned by getElement method. If the return is
' nothing, varResult will be "null". This method is in exchange for "If element Is
' Nothing" check as the getElement has to be returned for onExist and onExistNot.
' Updated: 27/04/2023
'----------------------------------------------------------------------------------
    
    varResult = varType
    
End Property
 
 
Private Function getVarID() As String
'----------------------------------------------------------------------------------
' Convert varJS to varID for ease of access when needed to further interact with
' the element post using varJS to initially access it in the browser
' Updated: 26/04/2023
'----------------------------------------------------------------------------------
    
   'If varID already assigned
    getVarID = varID: If getVarID <> "" Then Exit Function
    
   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID
       
   'Assign the new varID using varJS
    main.jsEval varID_ & " = " & varJS, printDbgMsg
    
    varID = varID_  'Store the new ID to the main ID var
    getVarID = varID_
    
End Function
 
 
Private Function newVarID() As String
'----------------------------------------------------------------------------------
' Create a unique varID for element reference with CDPElement. If the ID is not
' unique, it is possible that new element assignment will override existing ones
' under the same ID by chance if working on a large number of elements.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------
    Dim newID As String
    Do: newID = "varID" & Format(Rnd * 100000, "000000")
    Loop Until Not varIDs.Exists(newID)
    varIDs.Add newID, newID
    
   'Return the new varID
    newVarID = newID
    
End Function
 
'========================================================================================================
' ELEMENT PROPERTY INTERACTIONS
'========================================================================================================
 
Property Get value() As String
'----------------------------------------------------------
' Get the value property of the element in the HTML.
' Updated: 02/03/2023
'----------------------------------------------------------
    
    value = main.jsEval(getVarID & ".value", printDbgMsg)
    
End Property

 
Property Let value(strTextVal As String)
'-----------------------------------------------------------------------------------------------
' Change the value property of the element in the HTML. Enhanced to work even with REACT fields.
' For React fields, CDP need to input value to the element prototype in order to trigger the
' React event handler. For further references refer below:
' https://stackoverflow.com/a/46012210
' https://hustle.bizongo.in/simulate-react-on-change-on-controlled-components-baa336920e04
' Updated: 28/04/2023 - React fields value inputing enhancement
'-----------------------------------------------------------------------------------------------
                
   'Help devs know if the field might is a React field
    Dim e1 As String, isReact As Boolean
    e1 = main.jsEval("Object.entries(" & getVarID & ")[0][0]", False)
    If InStr(e1, "react") Or InStr(e1, "Tracker") Or InStr(e1, "State") Then isReact = True
    If isReact Then main.printMsg "FYI: this is a REACT field as the first method is " & e1
                
   'Perform inputing using all the methods
    If Not isReact Then
        main.jsEval getVarID & ".value = """ & strTextVal & """", printDbgMsg
    Else
        main.jsEval "Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(" & getVarID & ", '" & strTextVal & "')"
    End If
    
   'Most input elements need this
    fireEvent "input", printDbg:=False
    
End Property
 
 
Property Get innerText() As String
'----------------------------------------------------------
' Get the innerText property of the element in the HTML.
' Updated: 02/03/2023
'----------------------------------------------------------
    
    innerText = main.jsEval(getVarID & ".innerText", printDbgMsg)
    
End Property
 
 
Property Let innerText(strTextVal As String)
'----------------------------------------------------------
' Change the innerText property of the element in the HTML.
' Updated: 02/03/2023
'----------------------------------------------------------

    main.jsEval getVarID & ".innerText = """ & strTextVal & """", printDbgMsg
    
End Property

 
Property Get innerHTML() As String
'----------------------------------------------------------
' Get the innerHTML property of the element in the HTML.
' Updated: 02/03/2023
'----------------------------------------------------------
    
    innerHTML = main.jsEval(getVarID & ".innerHTML", printDbgMsg)
    
End Property
 
 
Public Function getAttribute(strAttributeName As String) As String
'----------------------------------------------------------
' Get the value of an element attribute based on its name.
' Updated: 02/03/2023
'----------------------------------------------------------

    getAttribute = main.jsEval(getVarID & ".getAttribute(""" & strAttributeName & """)", printDbgMsg)
    
End Function
 
 
Public Function setAttribute(strAttributeName As String, strValue As String) As String
'----------------------------------------------------------
' Set value to an element attribute based on its name.
' Updated: 02/03/2023
'----------------------------------------------------------

    main.jsEval getVarID & ".setAttribute(""" & strAttributeName & """, """ & strValue & """)", printDbgMsg
    
End Function
 
'========================================================================================================
' ELEMENT OBJECT INTERACTIONS
'========================================================================================================
 
Public Function click(Optional till As ReadyState = isComplete)
'----------------------------------------------------------
' Perform a standard click on the element. Remember to
' perform wait if necessary for the page to be refreshed.
' Updated: 02/03/2023
'----------------------------------------------------------

    main.jsEval getVarID & ".click()", printDbgMsg
    main.wait till:=till
    
End Function
 
 
Public Function submit(Optional till As ReadyState = isComplete)
'----------------------------------------------------------
' Submit a form element. This is equivalence to sending
' "enter" after completing a form. Most of the times, the
' standard way is to simply click the 'Submit' button that
' HTML forms usually have by design.
' Updated: 02/03/2023
'----------------------------------------------------------

    main.jsEval getVarID & ".form.submit()", printDbgMsg
    main.wait till:=till
 
End Function
 
 
Public Function fireEvent(strEventName As String, Optional till As ReadyState = isComplete, Optional printDbg As Boolean = printDbgMsg)
'------------------------------------------------------------
' Dispatch the custom event on the element. Unlike the IE
' DOM event name, the argument strEventName for JS should
' not start with "on" such as "onblur" but only "blur".
' Updated: 27/04/2023
'------------------------------------------------------------

    strEventName = Replace(strEventName, "on", "")  'eg. "onchange" -> "change" - correct JS event name
    
   'Fire the event; simulated=true is in case form is REACT 15.5 or lower
    main.jsEval getVarID & ".dispatchEvent(new Event('" & strEventName & "', { bubbles: true, simulated: true }))", printDbg
    
   'Wait till applied
    main.wait till:=till, dbgState:=printDbg
    
End Function


Public Function isExist() As Boolean
'----------------------------------------------------------
' Easy way to check if the element is existing within the
' current page being referred to by the CDP object.
' Updated: 06/06/2023
'----------------------------------------------------------
            
    If varType = "null" Then isExist = False Else isExist = True
 
End Function

 
 
Public Function onExist(Optional timeOutInSeconds As Double = 30) As CDPElement
'----------------------------------------------------------
' Wait until the element applying this method comes into
' existance on the HTML page.
' Updated: 26/04/2023
'----------------------------------------------------------
    Dim timeStart As Single
    timeStart = Timer
    Do While varType = "null"
        If Timer - timeStart > timeOutInSeconds Then
            main.printMsg "onExists timeout on varJS = " & varJS
            Exit Function
        End If
        varType = main.jsEval(varJS, False)
        main.sleep 0.1
    Loop
        
    Set onExist = Me
    main.printMsg "onExists satisfied: " & getVarID & " is now available for interaction"
 
End Function
 
 
Public Function onExistNot(Optional timeOutInSeconds As Double = 30) As Boolean
'----------------------------------------------------------
' Wait until the element applying this method comes into
' non-existance on the HTML page. Unlike onExists, this
' does not return anything as the element would be out of
' existence hence no further interactions should be needed.
' Updated: 04/05/2023 - make function returns True
'----------------------------------------------------------
    Dim timeStart As Single
    timeStart = Timer
    Do While varType <> "null"
        If Timer - timeStart > timeOutInSeconds Then
            main.printMsg "onExistsNot timeout on varJS = " & varJS
            Exit Function
        End If
        varType = main.jsEval(varJS, False)
        main.sleep 0.1
    Loop
                
    onExistNot = True
                
    main.printMsg "onExistsNot satisfied: " & getVarID & " is no longer available for interaction"
 
End Function
 
 
Public Function sendKeys(textToSend As String)

'----------------------------------------------------------------
' Attempt to simulate physical key input instead of DOM
' manipulation via JavaScript. This sendKeys method will
' work on a specific element and thus is better than the generic
' VBA SendKeys method. Useful when other means of value
' inputing fail or the target field required isTrusted = True.
' This is a very powerful technique but requires the window to be
' visible (ie. it must not be hidden with .hide).
' Updated: 28/04/2023 - Created & debugged the "+" sign issue.
'----------------------------------------------------------------
        
   'First clear the current text of the field
    main.jsEval getVarID & ".value = ''", False
        
   'Get objectId
    Dim params As New Scripting.Dictionary, objectId As String
    params("expression") = getVarID
    objectId = CStr(main.invokeMethod("Runtime.evaluate", params)("result")("objectId"))
    
   'If the objectId returned is not negative, an implicit positive
   'needs to be added to prevent DOM.focus hanging. This is likely because
   'the original positive sign of objectId was incorrectly removed by invokeMethod/Runtime.evaluate
    If Left(objectId, 1) <> "-" Then objectId = "+" & objectId
 
   'Set focus on the node
   'Without focus, the sendkeys will fly to infinity
    params.RemoveAll
    params("objectId") = objectId
    main.invokeMethod "DOM.focus", params
 
   'Sendkeys to the node
    params.RemoveAll
    params("text") = textToSend
    main.invokeMethod "Input.insertText", params
    
    main.printMsg "Executed sendKeys(""" & textToSend & """)"
 
End Function
 
'========================================================================================================
' NODE-TREE TRAVERSAL METHODS
'========================================================================================================
 
Public Function getParent() As CDPElement
'----------------------------------------------------------
' Get the parent node of the current element.
' Updated: 02/03/2023
'----------------------------------------------------------
   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID
    
   'Get the object and assign to the new ID in the browser
    Dim varJS_ As String: Dim result As String
    varJS_ = varID_ & " = " & getVarID & ".parentElement"
    result = main.jsEval(varJS_, printDbgMsg)
    
   'Return result
    Set getParent = New CDPElement
    getParent.init main, varIDs, varID_, varJS_, result
    
End Function
 
 
Public Function getNextSibling() As CDPElement
'----------------------------------------------------------
' Get the next sibling node of the current element.
' Updated: 02/03/2023
'----------------------------------------------------------
     
   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID
    
   'Get the object and assign to the new ID in the browser
    Dim varJS_ As String: Dim result As String
    varJS_ = varID_ & " = " & getVarID & ".nextElementSibling"
    result = main.jsEval(varJS_, printDbgMsg)
    
   'Return result
    Set getNextSibling = New CDPElement
    getNextSibling.init main, varIDs, varID_, varJS_, result
    
End Function
 
 
Public Function getPrevSibling() As CDPElement
'----------------------------------------------------------
' Get the previous sibling node of the current element.
' Updated: 02/03/2023
'----------------------------------------------------------
    
   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID
    
   'Get the object and assign to the new ID in the browser
    Dim varJS_ As String: Dim result As String
    varJS_ = varID_ & " = " & getVarID & ".previousElementSibling"
    result = main.jsEval(varJS_, printDbgMsg)
    
   'Return result
    Set getNextSibling = New CDPElement
    getPrevSibling.init main, varIDs, varID_, varJS_, result
    
End Function
 
 
Public Function getFirstChild() As CDPElement
'----------------------------------------------------------
' Get the next sibling node of the current element.
' Updated: 02/03/2023
'----------------------------------------------------------
    
   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID
    
   'Get the object and assign to the new ID in the browser
    Dim varJS_ As String: Dim result As String
    varJS_ = varID_ & " = " & getVarID & ".firstElementChild"
    result = main.jsEval(varJS_, printDbgMsg)
    
   'Return result
    Set getFirstChild = New CDPElement
    getFirstChild.init main, varIDs, varID_, varJS_, result
    
End Function
 
 
Public Function getElementByID(strID As String) As CDPElement
'----------------------------------------------------------------------------------
' Obtain an element that is the child of the current element with the defined HTML
' ID. This enables multiple getElementBy methods to be stringed together. Thus for
' this to work, the varID of the current element is employed instead of 'document'.
' Updated: 02/03/2023
'----------------------------------------------------------------------------------
    
   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID
    
   'Get the object and assign to the varID in the browser
    Dim varJS_ As String: Dim result As String
    varJS_ = varID_ & " = " & getVarID & ".getElementById(""" & strID & """)"
    result = main.jsEval(varJS_)
    
   'Return result
    Set getElementByID = New CDPElement
    getElementByID.init main, varIDs, varID_, varJS_, result
    
End Function
 
 
Public Function getElementByQuery(strQuery As String) As CDPElement
'----------------------------------------------------------------------------------
' Obtain an element that is the child of the current element using QuerySelector.
' This enables multiple getElementBy methods to be stringed together. Thus for
' this to work, the varID of the current element is employed instead of 'document'.
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
    Dim varID_ As String: varID_ = newVarID
    
   'Get the object and assign to the varID in the browser
    Dim varJS_ As String: Dim result As String
    varJS_ = varID_ & " = " & getVarID & ".querySelector(""" & strQuery & """)"
    result = main.jsEval(varJS_)
    
   'Return result
    Set getElementByQuery = New CDPElement
    getElementByQuery.init main, varIDs, varID_, varJS_, result
    
End Function
 
 
Public Function getElementsByQuery(strQuery As String) As Collection
'----------------------------------------------------------------------------------
' Obtain a collection of elements which are the child of the current element using
' QuerySelectorAll. This enables multiple getElementBy methods to be stringed
' together, ie. the varID of the current element is employed instead of 'document'.
' Updated: 29/03/2023
'----------------------------------------------------------------------------------
           
   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID & "s"     's to signify a collection varID
      
   'Get the object and assign to the varID in the browser
    Dim result As Variant
    result = main.jsEval(varID_ & " = " & getVarID & ".querySelectorAll(""" & strQuery & """)")
    
   'Return result
    Dim elements As Collection
    If result = "null" Then Exit Function Else Set elements = New Collection
    
   'Assign IDs to each of the element in the collection
    Dim varJS__ As String                                'Store each new member's JS access string
    Dim length_ As Long, tElement As CDPElement, N As Long
    length_ = main.jsEval(varID_ & ".length")
    For N = 0 To length_ - 1                             'HTML index starts at zero
        varJS__ = varID_ & ".item(" & N & ")"
        Set tElement = New CDPElement
        tElement.init main, varIDs, varJS_:=varJS__
        elements.Add tElement
    Next
    
    Set getElementsByQuery = elements
    
End Function
 
 
Public Function getElementByXPath(strXPath As String) As CDPElement
'----------------------------------------------------------------------------------
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
    Dim varID_ As String: varID_ = newVarID
           
   'Verify double slash syntax
   'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")
           
   'Check if strXPath has invalid // syntax
   'Per XPath technique, for acquiring child elements, the starting call string must not be "\\"
   'https://bytes.com/topic/javascript/answers/542739-how-use-contextnode-document-evaluate
    If Left(strXPath, 2) = "//" Then
        main.printMsg "Note: '//' has been removed from the next XPath call string """ & strXPath & """ as it is an invalid syntax for the contextNode"
        strXPath = Right(strXPath, Len(strXPath) - 2)
    End If
           
   'Get the object and assign to the varID in the browser
    Dim varJS_ As String: Dim result As String
    varJS_ = varID_ & " = " & "document" & ".evaluate(""" & strXPath & """, " & getVarID & ", null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue"
    result = main.jsEval(varJS_)
    
   'Return result
    Set getElementByXPath = New CDPElement
    getElementByXPath.init main, varIDs, varID_, varJS_, result
    
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
    Dim varID_ As String: varID_ = newVarID & "s"     's to signify a collection varID
        
   'Verify double slash syntax
   'People normally get this wrong when using XPath
    strXPath = Replace(strXPath, "\", "/")
           
   'Check if strXPath has invalid // syntax
   'Per XPath technique, for acquiring child elements, the starting call string must not be "\\"
   'https://bytes.com/topic/javascript/answers/542739-how-use-contextnode-document-evaluate
    If Left(strXPath, 2) = "//" Then
        main.printMsg "Note: '//' has been removed from the next XPath call string """ & strXPath & """ as it is an invalid syntax for the contextNode"
        strXPath = Right(strXPath, Len(strXPath) - 2)
    End If
           
   'Get the object and assign to the varID in the browser
    Dim result As Variant
    result = main.jsEval(varID_ & " = " & "document" & ".evaluate(""" & strXPath & """, " & getVarID & ", null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null)")
    
   'Return result
    Dim elements As Collection
    If result = "null" Then Exit Function Else Set elements = New Collection
    
   'Store element members to the collection
    Dim varJS__ As String                                    'Store each new member access JS string
    Dim length_ As Long, tElement As CDPElement, N As Long
    length_ = main.jsEval(varID_ & ".snapshotLength")
    For N = 0 To length_ - 1                                'HTML index starts at zero
        varJS__ = varID_ & ".snapshotItem(" & N & ")"
        Set tElement = New CDPElement
        tElement.init main, varIDs, varJS_:=varJS__
        elements.Add tElement
    Next
    
    Set getElementsByXPath = elements
    
End Function
 
 
Public Function getIFrame() As CDPElement
'----------------------------------------------------------------------------------
' Obtain the iFrame (inline frame) document of an iFrame HTML element. Some
' websites still use iFrame for form interactions with users so this method will be
' useful to quickly access the document object of the iFrame. iFrame cannot be
' directly interacted with getElement methods but must be accessed via its document
' object using .contentDocument. If an iFrame has a source, it is likely loaded not
' on the same domain. As a result, only indirect automation to the iFrame document
' may work. For further reference, visit w3 info:
' https://www.w3schools.com/jsref/prop_frame_contentdocument.asp
' Updated: 02/03/2023
'----------------------------------------------------------------------------------
           
   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID
    
   'Check if the iFrame is hosted on a different domain
    Dim src As String
    src = getAttribute("src")
    If InStr(src, "http") Then main.printMsg "This iFrame may not be hosted in the same domain. src=""" & src & """"
    
   'Get the object and assign to the varID in the browser
    Dim varJS_ As String: Dim result As String
    varJS_ = varID_ & " = " & getVarID & ".contentDocument"
    result = main.jsEval(varJS_)
 
   'Return result
    Set getIFrame = New CDPElement
    getIFrame.init main, varIDs, varID_, varJS_, result
    
End Function
