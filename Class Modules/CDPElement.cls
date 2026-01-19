'================================================================================================
' CDP HTML Element Class
'------------------------------------------------------------------------------------------------
' Author(s)   :
'       Long Vh (long.hoang.vu@hsbc.com.sg)
' Last Update :
'       24/10/25 Daniel Polak: Dim of variables and added Option Explicit, removed extraneous blank lines
'                              getPrevSibling: fixed bug where wrong assignment led to no return result
'                              setAttribute: changed Function to Sub
'       02/03/23 Long Vh: - Added the element class for easy object reference.
'       13/11/23 Long Vh: - Added getChildren & property let innerHTML. Updated .value to work
'                           with TextArea node type.
'       23/11/23 Long Vh: - Added raiseTimeoutError option for .onExist and .onExistNot
'       12/12/23 Long Vh: - Added checkExist to only execute things if True as well as ifExist
'       12/07/24 Long Vh: - Added the sendKey method.
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
Private checkExist As Boolean       'Needed for the isExist check
Private Const printDbgMsg = True

Public Enum keyboardCode    'Used for .sendKey operation in the browser class and element class

    keyEnter = 13
    keyDelete = 46
    keyTab = 9
    keyEsc = 27
    keyBackspace = 8

End Enum

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
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    value = main.jsEval(getVarID & ".value", printDbgMsg)

End Property

Property Let value(strTextVal As String)
'-----------------------------------------------------------------------------------------------
' Change the value property of the element in the HTML. Enhanced to work even with React fields.
' For React fields, CDP need to input value to the element prototype in order to trigger the
' React event handler. For further references refer below:
' https://stackoverflow.com/a/46012210
' https://hustle.bizongo.in/simulate-react-on-change-on-controlled-components-baa336920e04
' Updated: 28/04/23 - React fields value inputing enhancement
'          12/12/23 - added checkExist
'-----------------------------------------------------------------------------------------------

    Dim e1 As String
    Dim isReact As Boolean

   'Gracefully exit the value input if the element is found not exist under isExist
    If checkExist Then If Not isExist Then Exit Property

   'Help devs know if the field is a React field
    e1 = main.jsEval("Object.entries(" & getVarID & ")[0][0]", False)
    If InStr(e1, "react") Or InStr(e1, "Tracker") Or InStr(e1, "State") Then isReact = True
    If isReact Then main.printMsg "FYI: this is a REACT field as the first method is " & e1

   'Perform inputing using all the methods
    If Not isReact Then
        main.jsEval getVarID & ".value = """ & strTextVal & """", printDbgMsg
    Else
        main.jsEval "Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(" & getVarID & ", '" & strTextVal & "')"
        main.jsEval "Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, ""value"").set.call(" & getVarID & ", """ & strTextVal & """)"
    End If

   'Most input elements need this
    fireEvent "input", printDbg:=False

End Property

Property Get innerText() As String
'----------------------------------------------------------
' Get the innerText property of the element in the HTML.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    innerText = main.jsEval(getVarID & ".innerText", printDbgMsg)

End Property

Property Let innerText(strTextVal As String)
'----------------------------------------------------------
' Change the innerText property of the element in the HTML.
' Updated: 02/03/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    main.jsEval getVarID & ".scrollIntoView(false)", False

   'Escape quotation and apostrophe
    strTextVal = Replace(strTextVal, "'", "\'")
    strTextVal = Replace(strTextVal, """", "\""")

    main.jsEval getVarID & ".innerText = " & "'" & strTextVal & "'", printDbgMsg

End Property

Property Get innerHTML() As String
'----------------------------------------------------------
' Get the innerHTML property of the element in the HTML.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    innerHTML = main.jsEval(getVarID & ".innerHTML", printDbgMsg)

End Property

Property Let innerHTML(strTextVal As String)
'----------------------------------------------------------
' Change the innerHTML property of the element.
' Updated: 13/11/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    main.jsEval getVarID & ".scrollIntoView(false)", False

   'Escape quotation and apostrophe
    strTextVal = Replace(strTextVal, "'", "\'")
    strTextVal = Replace(strTextVal, """", "\""")

    main.jsEval getVarID & ".innerHTML = " & "'" & strTextVal & "'", printDbgMsg

End Property

Property Let checked(boolChecked As Boolean)
'----------------------------------------------------------
' Change the checked value of a check box element
' Updated: 04/12/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    main.jsEval getVarID & ".checked = " & LCase(boolChecked), False

End Property

Property Get checked() As Boolean
'----------------------------------------------------------
' Change the checked value of a check box element
' Updated: 04/12/2023
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    checked = CBool(main.jsEval(getVarID & ".checked", printDbgMsg))

End Property

Property Let selected(selectedOption As String)
'----------------------------------------------------------
' Change the selected value of an html select element type.
' eg. br.getElementbyId("list").selected = 2
' Updated: 06/03/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    main.jsEval getVarID & ".selectedIndex = " & selectedOption, False
    fireEvent "change"

End Property

Property Get selected() As String
'----------------------------------------------------------
' Get the selected option. Only returns the first option.
' Updated: 06/03/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Property
    selected = main.jsEval(getVarID & ".selectedOptions[0]", printDbgMsg)

End Property

'========================================================================================================
' ELEMENT OBJECT INTERACTIONS
'========================================================================================================

Public Function click(Optional till As ReadyState = isComplete) As Boolean
'----------------------------------------------------------
' Perform a standard click on the element. Remember to
' perform wait if necessary for the page to be refreshed.
' Updated: 12/12/2023
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

   'Scroll into the view for better visibility of the automation
    main.jsEval getVarID & ".scrollIntoView(false)", False

   'Some fields require the browser to manually focus before allowing
   'the click action, such as the input type 'file' field. The below
   'attempts to simulate a manual activation on the browser to thus
   'enable the click acceptance of the html.
    Dim params As New Scripting.Dictionary
    params.RemoveAll
    params("text") = ""
    main.invokeMethod "Input.insertText", params

   'Perform click
    main.jsEval getVarID & ".dispatchEvent(new MouseEvent('click', {view: window, bubbles:true, cancelable: true}))", printDbgMsg
    main.wait till:=till
    click = True

End Function

Public Function submit(Optional till As ReadyState = isComplete) As Boolean
'----------------------------------------------------------
' Submit a form element. This is equivalence to sending
' "enter" after completing a form. Most of the times, the
' standard way is to simply click the 'Submit' button that
' HTML forms usually have by design.
' Updated: 12/12/2023
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

    main.jsEval getVarID & ".form.submit()", printDbgMsg
    main.wait till:=till
    submit = True

End Function

Public Function getAttribute(strAttributeName As String) As String
'----------------------------------------------------------
' Get the value of an element attribute based on its name.
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

    getAttribute = main.jsEval(getVarID & ".getAttribute(""" & strAttributeName & """)", printDbgMsg)

End Function

Public Function setAttribute(strAttributeName As String, strValue As String)
'----------------------------------------------------------
' Set value to an element attribute based on its name.
' Daniel Polak recommended this procedure to be a sub i/o
' function but I decided to leave it as function to make
' the presentation consistent with the nature of the module
' Updated: 12/12/23 - added checkExist
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

    main.jsEval getVarID & ".setAttribute(""" & strAttributeName & """, """ & strValue & """)", printDbgMsg

End Function

Public Function setSelection(strOptionName As String)
'----------------------------------------------------------
' Set option value to a select element.
' Updated: 06/03/2025
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

    main.jsEval getVarID & ".scrollIntoView(false)", False
    Me.value = strOptionName
    fireEvent "change"

End Function

Public Function selectText()
'----------------------------------------------------------
' Select all the text held by the element.
' Updated: 18/09/2025
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

    main.jsEval "r = document.createRange()", False
    main.jsEval "s = window.getSelection()", False
    main.jsEval "r.selectNodeContents(" & getVarID & ")"
    main.jsEval "s.removeAllRanges()"
    main.jsEval "s.addRange(r)"

End Function

Public Function focus()
'----------------------------------------------------------
' Set focus to an element.
' Updated: 18/09/2025
'----------------------------------------------------------

    If checkExist Then If Not isExist Then Exit Function
    main.jsEval getVarID & ".focus()", False

End Function

Public Function fireEvent(strEventName As String, Optional till As ReadyState = isComplete, Optional printDbg As Boolean = printDbgMsg) As Boolean
'------------------------------------------------------------
' Dispatch the custom event on the element. Unlike the IE
' DOM event name, the argument strEventName for JS should
' not start with "on" such as "onblur" but only "blur".
' Updated: 12/12/2023
'------------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

   'Auto format some common event name mistakes
    strEventName = Replace(strEventName, "on", "")  'eg. "onchange" -> "change" - correct JS event name

   'Fire the event; simulated=true is in case form is REACT 15.5 or lower
    main.jsEval getVarID & ".dispatchEvent(new Event('" & strEventName & "', { bubbles: true, simulated: true }))", printDbg

   'Wait till applied
    main.wait till:=till, dbgState:=printDbg
    fireEvent = True

End Function

Public Function sendString(textToSend As String) As Boolean
'----------------------------------------------------------------
' Attempt to directly send the text string to the element via
' backdoor interaction with CDP instead of manipulating the DOM.
' This method allows bypassing many JavaScript input requirements
' when dealing with different type of html input elements.
' Updated: 28/04/2023 - Created & debugged the "+" sign issue.
'----------------------------------------------------------------

    Dim objectId As String

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

   'First clear the current text of the field
    main.jsEval getVarID & ".value = ''", False

   'Get objectId
    Dim params As New Scripting.Dictionary
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

    sendString = True
    main.printMsg "Executed sendString(""" & textToSend & """) successfully on " & getVarID

End Function

Public Function sendClick() As Boolean
'----------------------------------------------------------------
' Attempt to simulate physical left mouse click input instead of
' DOM manipulation with JavaScript dispatchEvent method.
' Similarly to sendString, sendClick interacts with the target
' element via backdoor manipulation using CDP ports and thus can
' bypass many complex JavaScript requirements to trigger a click.
'----------------------------------------------------------------

    Dim x, y As String

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

   'Get element coordinates
    x = main.jsEval(getVarID & ".getBoundingClientRect().x")
    y = main.jsEval(getVarID & ".getBoundingClientRect().y")

   'Sendkeys to the node
    Dim params As New Scripting.Dictionary
    params("type") = "mousePressed"
    params("x") = x
    params("y") = y
    params("button") = "left"
    params("clickCount") = 1
    main.invokeMethod "Input.dispatchMouseEvent", params

    sendClick = True
    main.printMsg "Executed sendClick successfully on " & getVarID

End Function

Public Function sendKey(key As keyboardCode, Optional altKey As Boolean = False) As Boolean
'----------------------------------------------------------------
' Attempt to directly send a unique keyboard input to the element
' such as backspace, enter, esc etc. This will help with some
' nasty field that only allows certain action to be executed with
' raw keyboard handling.
'----------------------------------------------------------------

    Dim objectId As String

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

   'Get objectId
    Dim params As New Scripting.Dictionary
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

   'Send key down event to the browser
    params.RemoveAll
    params("type") = "keyDown"
    params("windowsVirtualKeyCode") = key
    params("altKey") = altKey
    main.invokeMethod "Input.dispatchKeyEvent", params

   'Send key up event to the browser to release the key
    params.RemoveAll
    params("type") = "keyUp"
    params("windowsVirtualKeyCode") = key
    params("altKey") = False
    main.invokeMethod "Input.dispatchKeyEvent", params

    sendKey = True
    main.printMsg "Executed sendKey successfully on " & getVarID

End Function

Public Function clearValue() As Boolean
'----------------------------------------------------------------
' Clear the value currently had by the input element
'----------------------------------------------------------------

    Dim e1 As String
    Dim isReact As Boolean

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then If Not isExist Then Exit Function

   'Scroll into view
    main.jsEval getVarID & ".scrollIntoView(false)", False

   'Help devs know if the field might is a React field
    e1 = main.jsEval("Object.entries(" & getVarID & ")[0][0]", False)
    If InStr(e1, "react") Or InStr(e1, "Tracker") Or InStr(e1, "State") Then isReact = True

   'Perform inputing using all the methods
    If Not isReact Then
        main.jsEval getVarID & ".value = """"""", printDbgMsg
    Else
        main.jsEval "Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(" & getVarID & ", '')"
        main.jsEval "Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set.call(" & getVarID & ", '')"
    End If

   'Most input elements need this
    fireEvent "input", printDbg:=False
    clearValue = True

End Function

Public Function isExist() As Boolean
'----------------------------------------------------------
' Easy way to check if the element is existing within the
' current page being referred to by the CDP object.
' Updated: 06/06/2023
'----------------------------------------------------------

    If varType = "null" Then isExist = False Else isExist = True

End Function

Public Function ifExist() As CDPElement
'----------------------------------------------------------
' Easy way to check if the element is existing within the
' current page being referred to by the CDP object. When the
' element does not exist, it returns an empty CDPElement
' object with checkExist tag = False, effectively skipping
' all subsequent functions related to it.
' Updated: 12/12/2023
'----------------------------------------------------------

    checkExist = True
    Set ifExist = Me
    main.printMsg "ifExist returned " & isExist & " on [" & varJS & "]." & IIf(Not isExist, " Subsequent actions on this element are disregarded.", "")

End Function

Public Function onExist(Optional timeOutInSeconds As Double = 30, Optional raiseTimeoutError As Boolean = True) As CDPElement
'----------------------------------------------------------
' Wait until the element applying this method comes into
' existance on the HTML page.
' Updated: 23/11/2023
'----------------------------------------------------------

    Dim timeStart As Single

    timeStart = Timer
    Do While varType = "null"
        If Timer - timeStart > timeOutInSeconds Then
            main.printMsg "onExists timeout on varJS = " & varJS, doRaiseError:=raiseTimeoutError
            Exit Function
        End If
        varType = main.jsEval(varJS, False)
        main.sleep 0.1
    Loop

    Set onExist = Me
    main.printMsg "onExists satisfied: " & getVarID & " is now available for interaction"

End Function

Public Function onExistNot(Optional timeOutInSeconds As Double = 30, Optional raiseTimeoutError As Boolean = True) As Boolean
'----------------------------------------------------------
' Wait until the element applying this method comes into
' non-existance on the HTML page. Unlike onExists, this
' does not return anything as the element would be out of
' existence hence no further interactions should be needed.
' Updated: 23/11/2023 - make function returns True
'----------------------------------------------------------

    Dim timeStart As Single

    timeStart = Timer
    Do While varType <> "null"
        If Timer - timeStart > timeOutInSeconds Then
            main.printMsg "onExistsNot timeout on varJS = " & varJS, doRaiseError:=raiseTimeoutError
            Exit Function
        End If
        varType = main.jsEval(varJS, False)
        main.sleep 0.1
    Loop

    onExistNot = True
    main.printMsg "onExistsNot satisfied: " & getVarID & " is no longer available for interaction"

End Function

'========================================================================================================
' NODE-TREE TRAVERSAL METHODS
'========================================================================================================

Public Function getParent() As CDPElement
'----------------------------------------------------------
' Get the parent node of the current element.
' Updated: 12/12/23
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

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
' Updated: 12/12/23
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

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
' Updated: 24/10/2025: Daniel Polak - fixed bug where wrong assignment led to no return result
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID

   'Get the object and assign to the new ID in the browser
    Dim varJS_ As String: Dim result As String
    varJS_ = varID_ & " = " & getVarID & ".previousElementSibling"
    result = main.jsEval(varJS_, printDbgMsg)

   'Return result
    Set getPrevSibling = New CDPElement
    getPrevSibling.init main, varIDs, varID_, varJS_, result

End Function

Public Function getFirstChild() As CDPElement
'----------------------------------------------------------
' Get the next sibling node of the current element.
' Updated: 12/12/2023
'----------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

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

Public Function getChildren() As Collection
'----------------------------------------------------------------------------------
' Obtain a collection of elements which are the child of the current element.
' This enables multiple getElementBy methods to be stringed together, ie. the varID
' of the current element is employed instead of 'document'.
' Updated: 10/11/2023
'----------------------------------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist And Not isExist Then
        Set getChildren = New Collection 'set an empty collection
        Exit Function
    End If

   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID & "s"     's to signify a collection varID

   'Get the object and assign to the varID in the browser
    Dim result As String
    result = main.jsEval(varID_ & " = " & getVarID & ".children")

   'Return result
    Dim elements As Collection
    If result = "null" Then Exit Function Else Set elements = New Collection

   'Assign IDs to each of the element in the collection
    Dim varJS__ As String                                'Store each new member's JS access string
    Dim length_ As Integer, n As Integer
    Dim tElement As CDPElement
    length_ = main.jsEval(varID_ & ".length")
    For n = 0 To length_ - 1                             'HTML index starts at zero
        varJS__ = varID_ & ".item(" & n & ")"
        Set tElement = New CDPElement
        tElement.init main, varIDs, varJS_:=varJS__
        elements.Add tElement
    Next

    Set getChildren = elements

End Function

Public Function getElementByID(strID As String) As CDPElement
'----------------------------------------------------------------------------------
' Obtain an element that is the child of the current element with the defined HTML
' ID. This enables multiple getElementBy methods to be stringed together. Thus for
' this to work, the varID of the current element is employed instead of 'document'.
' Updated: 12/12/2023
'----------------------------------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

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
' Updated: 12/12/2023
' Examples ------------------------------------------------------------------------
' eg.1 "#abc": get the first element with id attribute = abc
' eg.2 "tb.xyz": get the first element with tag tb and with class xyz
' eg.3 "a[name='abc']": get the first element with tag a and attribute name = abc
' eg.4 "div > p": get the first element with tag 'p' and has parent with tag 'div'
' eg.5 "[a1='abc'][a2='xyz']": get element with attribute a1 = abc and a2 = 123
' eg.6 "[a1='abc'] [a2='xyz']": get element with attribute a2 = xyz and is a child
' of an element with attribute a1 = abc
'----------------------------------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

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
' Updated: 12/12/2023
'----------------------------------------------------------------------------------

    Dim result As String
    Dim elements As Collection
    Dim n As Long, length_ As Long
    Dim tElement As CDPElement

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID & "s"     's to signify a collection varID

   'Get the object and assign to the varID in the browser
    result = main.jsEval(varID_ & " = " & getVarID & ".querySelectorAll(""" & strQuery & """)")

   'Return result
    If result = "null" Then Exit Function Else Set elements = New Collection

   'Assign IDs to each of the element in the collection
    Dim varJS__ As String                                'Store each new member's JS access string
    length_ = main.jsEval(varID_ & ".length")
    For n = 0 To length_ - 1                             'HTML index starts at zero
        varJS__ = varID_ & ".item(" & n & ")"
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
' Updated: 12/12/2023
' Examples -----------------------------------------------------------------------------
' eg.1 "//b[text()='T-BILL CURVE']": tag "b" with innertext to be 'T-BILL CURVE'
' eg.2 "//input[@name='username']": tag "input" & attribute "name" to be 'username'
' eg.3 "//a[@a1='abc' and @a2='xyz']": tag "a" with attribute a1 = abc and a2 = xyz
' eg.4 "//div[@id='abc']/div": get first div child of the div with id = abc
' eg.5 "//div[@id='abc'].//div": get first div descendant of the div with id = abc
' eg.6 "div": get first div child from the prescribed contextNode (default = 'document')
'---------------------------------------------------------------------------------------

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

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
' Updated: 12/12/2023
'----------------------------------------------------------------------------------

    Dim result As String
    Dim elements As Collection
    Dim n As Long, length_ As Long
    Dim tElement As CDPElement

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

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
    result = main.jsEval(varID_ & " = " & "document" & ".evaluate(""" & strXPath & """, " & getVarID & ", null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null)")

   'Return result
    If result = "null" Then Exit Function Else Set elements = New Collection

   'Store element members to the collection
    Dim varJS__ As String                                    'Store each new member access JS string
    length_ = main.jsEval(varID_ & ".snapshotLength")
    For n = 0 To length_ - 1                                'HTML index starts at zero
        varJS__ = varID_ & ".snapshotItem(" & n & ")"
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
' Updated: 12/12/2023
'----------------------------------------------------------------------------------

    Dim src As String

   'When ifExist is employed, the function should only execute if the element exists
    If checkExist Then
        If Not isExist Then
            Set getParent = Me
            Exit Function
        End If
    End If

   'Get a new unique varID
    Dim varID_ As String: varID_ = newVarID

   'Check if the iFrame is hosted on a different domain
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
