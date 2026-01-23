'==============================================================================================================
' Procedures for late-binding call on CDP object
'--------------------------------------------------------------------------------------------------------------
' In your external workbook, simply have cdp.xlam open via Workbooks.Open and make referrence to its function
' using Run("cdp.xlam!cdp"). E.g. to execute a newBrowser: Set e = Run("cdp.xlam!cdp").newBrowser
' Late-binding call has advantage of flexbile addin storage location, which means it can be shared to users of
' different computers while Early-binding call (via referrence-setting) will only works if the package cdp.xlam
' is stored in the same folder as your VBA project file.
'==============================================================================================================

Public customLogPath As String      'Enables the host project to set a custom log path if needed
Public doPrintDbgMsg As Boolean     'If true then all CDP Debug.Print msgs will not be printed
Public doLog As Boolean             'If true then no log will be produced at all

Public Function CDP() As CDPInit
'------------------------------------------------------------------
' Instantiate the class from external projects. Needed for late
' binding to xlam or initing class objects. The class has to be
' set to "Public Not Creatable" (F4 on class module). Note, to
' avoid conflict, ensure the name of this function, of the target
' class module, of the vb project, and of the module hosting this
' function are totally different names. Otherwise, external call
' will likely fail with exception "macro not enabled".
' Credit: https://stackoverflow.com/a/10016190
'------------------------------------------------------------------

    Set CDP = New CDPInit

End Function
