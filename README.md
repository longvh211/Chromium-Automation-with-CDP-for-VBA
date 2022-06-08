# Chromium Automation with CDP for VBA
This is a method to directly automate Chromium-based web browsers, such as Chrome, Edge, and Firefox, using VBA for Office applications by following the Chrome DevTools Protocol framework. This git is an enhanced framework based on the original pioneering article by ChrisK23 on CodeProject. You can find the original article as well as his example here at https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA\

**What It Can Do**

This method enables direct automation with Chromium-based web browsers for VBA without the need for a third-party tool software as SeleniumBasic. The framework also includes many examples and useful new functions added to the original repository while keeping the whole design as simple as possible to help you understand and get started quickly with deploying the CDP framework for your VBA solutions.

Functions that are added over the original:
1. A method to make the browser visible and invisible.
2. New methods to create and manage multiple tabs at the same time.
3. A method to handle browser window state, such as maximizing, minimizing, and resizing.
4. A method to parse additional arguments to allow setting up a browser automation session with advanced requirements.
5. A method to easily start Edge or Chrome for automation at the user's choice.
  
**For Demo**

The demo file has been prepared to help you get on-board easily with the framework. You can download the CDP Framework Excel macro file (.xlsm).

**For Installation**

You can download the module files in the import folder and add them to your VBA application using the Import Modules option in the VBIDE screen. Or you can also use the modules already setup in the demo file mentioned above. They are the same modules as well. Note that the modules require Microsoft Scripting Runtime reference for the Dictionary object to work.

**Notes**

This framework does not work for Edge IE Mode. For a framework that works on Edge IE Mode, see this git of mine instead:

https://github.com/longvh211/Edge-IE-Mode-Automation-with-IES-for-VBA/tree/main

**Credits**

ChrisK23 for the great original source: https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA\
PerditionC for plenty of helpful CDP examples: https://github.com/PerditionC/VBAChromeDevProtocol
