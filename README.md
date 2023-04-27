# Chromium Automation with CDP for VBA
This is a method to directly automate Chromium-based web browsers, such as Chrome, Edge, and Firefox, using VBA for Office applications by following the Chrome DevTools Protocol framework. This git is an enhanced framework based on the original pioneering article by ChrisK23 on CodeProject. You can find the original article as well as his example here at https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA\

**What It Can Do**

This method enables direct automation with Chromium-based web browsers for VBA without the need for a third-party software like SeleniumBasic. The framework also includes many examples and useful functions added to the original repository while keeping the whole design as simple as possible to help you understand and get started quickly with deploying the CDP framework for your VBA solutions.

Features greatly improved on v2.5:
1. Added getElementByID, getElementByQuery (querySelector equivalence), getElementByXPath, getElementsByQuery, getElementsByXPath.
2. Added helpful examples to the getElement methods' function definitions for ease of learning and employment.
3. Overhauled the error debugging system of CDP 1.0. The Immediate Window is now filled with highly detailed and useful debugging information.
4. Added AddJsLib, a powerful function to integrate external JS Library to greatly widen the automation scope of the framework.
5. Added snapPage to take snapshot of the web page or an element within the page. This demonstrates the power of the AddJsLib function.
6. No longer required for dev to use .deserialize or .serialize to rehook ongoing CDP Session. This is now done intuitively by the CDPBrowser class.
7. Added CDPElement class with many element specific methods for HTML element interactions, such as:
   - .getIFrame to easily access and work on an iFrame element.
   - .value, .innerHTML, .innerText, .click, .submit, .setAttribute, .fireEvent for diverse interaction needs.
   - .getParent, .getNextSibling, .getPrevSibling, .getFirstChild for diverse node tree traversal requirements.
   - .onExist and .onExistNot to smart wait until the target element appears/disappears.
8. Added .html to easily extract the entire html of the current web page. Useful for devs who need the html for processing.
9. Enhanced .start to automatically detect browser installation path and start the browser there. v1.0 was failing when the user chooses a non-standard path.
10.Many other minor bug fixes and improvements over the remaining functions.

Functions that are added over the original:
1. A method to make the browser visible and invisible.
2. New methods to create and manage multiple tabs at the same time.
3. A method to handle browser window state, such as maximizing, minimizing, and resizing.
4. A method to parse additional arguments to allow setting up a browser automation session with advanced requirements.
5. A method to easily start Edge or Chrome for automation at the user's choice.
  
**For Demo**

Open CDP Framework.xlam and look for the module named "Demo" inside there.

**For Installation**

1. Download CDP Framework.xlam and open it.
2. Copy CDPBrowser, CDPCore, CDPElement, CDPJConv classes over to your VBA project.
3. Make sure your project has Microsoft Scripting Runtime reference.

**Notes**

This framework does not work for Edge IE Mode. For a framework that works on Edge IE Mode, see this git of mine instead:

https://github.com/longvh211/Edge-IE-Mode-Automation-with-IES-for-VBA/tree/main

**Credits**

ChrisK23 for the great original source: https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA\

PerditionC for plenty of helpful CDP examples: https://github.com/PerditionC/VBAChromeDevProtocol
