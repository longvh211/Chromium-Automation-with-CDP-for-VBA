'==============================================================================================================
' Procedures for late-binding call on CDP object
'--------------------------------------------------------------------------------------------------------------
' In your external workbook, simply have cdp.xlam open via Workbooks.Open and make referrence to its function
' using Run("cdp.xlam!cdp"). E.g. to execute a newBrowser: Set e = Run("cdp.xlam!cdp").newBrowser
' Late-binding call has advantage of flexbile addin storage location, which means it can be shared to users of
' different computers while Early-binding call (via referrence-setting) will only works if the package cdp.xlam
' is stored in the same folder as your VBA project file.
'==============================================================================================================

Public Function newBrowser(Optional customLogPath As String = "", Optional doPrintDbgMsg As Boolean = False, Optional doLog As Boolean = False) As CDPBrowser
'------------------------------------------------------------------
' Instantiate the class from external projects. The class has to be
' set to "Public Not Creatable" (F4 on class module).
' Credit: https://stackoverflow.com/a/10016190
'------------------------------------------------------------------
' Note: debug message and logging are disabled by default to avoid
' spamming information but if needed, simply switch them to true.
'------------------------------------------------------------------

    init.customLogPath = customLogPath
    init.doPrintDbgMsg = doPrintDbgMsg
    init.doLog = doLog
    Set newBrowser = New CDPBrowser

End Function

Public Function newElement() As CDPElement
'------------------------------------------------------------------
' Instantiate the class from external projects. The class has to be
' set to "Public Not Creatable" (F4 on class module).
' Credit: https://stackoverflow.com/a/10016190
'------------------------------------------------------------------

    Set newElement = New CDPElement

End Function
