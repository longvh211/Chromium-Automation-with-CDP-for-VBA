VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
' Core APIs to generate Pipes and read CDP message
' This class is used by Browser class (clsBrowser)
'--------------------------------------------------------------------------------------------------------------
' Author(s)   :
'       ChrisK23 (Code Project)
' Contributors:
'       Long Vh (long.hoang.vu@hsbc.com.sg)
' Last Update :
'       03/06/22 - Long Vh: added extensive commentations + refactored codes for HSBC usages
' Notes       :
'       This class should only be initiated internally by the clsBrowser class and should not need to be
'       initiated directly at any times.
' References  :
'       Nil
' Sources     :
'       Creating Pipe: https://stackoverflow.com/questions/62172551/error-with-createpipe-in-vba-office-64bit
'==============================================================================================================

'===================================
' Win APIs Declarations
'===================================

Private Declare PtrSafe Function CreatePipe Lib "kernel32" ( _
    phReadPipe As LongPtr, _
    phWritePipe As LongPtr, _
    lpPipeAttributes As SECURITY_ATTRIBUTES, _
    ByVal nSize As Long) As Long

Private Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Private Declare PtrSafe Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As Any, _
    lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As Any, _
    lpProcessInformation As Any) As Long

Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr) As Long

Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As LongPtr, _
    lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    lpBytesRead As Long, _
    lpTotalBytesAvail As Long, _
    lpBytesLeftThisMessage As Long) As Long

Private Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    lpOverlapped As Long) As Long

Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As LongPtr)

'===================================
' Types & Structures Declarations
'===================================

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

' This is the structure to pass more than 3 fds to a child process (https://github.com/libuv/libuv/blob/v1.x/src/win/process-stdio.c)
' The fields crt_flags and os_handle must lie contigously in memory i.e. should not be aligned to byte boundaries
' You cannot define a packed struct in VBA, that's why we need to have multiple structs

Private Type STDIO_BUFFER
    number_of_fds As Long
    crt_flags(0 To 4) As Byte
    os_handle(0 To 4) As LongPtr
End Type

#If Win64 Then

    Private Type STDIO_BUFFER2
        number_of_fds As Long
        raw_bytes(0 To 44) As Byte
    End Type
    
#Else

    Private Type STDIO_BUFFER2
        number_of_fds As Long
        raw_bytes(0 To 24) As Byte
    End Type
    
#End If

'===================================
' Constants
'===================================

Private Const STARTF_USESTDHANDLES = &H100&
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESHOWWINDOW As Long = &H1&
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_RESTORE = 9

'===================================
' Shared Variables
'===================================

Private hStdOutWr As LongPtr
Private hStdOutRd As LongPtr
Private hStdInWr As LongPtr
Private hStdInRd As LongPtr
Private hCDPOutWr As LongPtr
Private hCDPOutRd As LongPtr
Private hCDPInWr As LongPtr
Private hCDPInRd As LongPtr
Private boolSerializable As Boolean
Private hProcess As LongPtr
Public jsConverter As New clsJsConverter

'===================================
' Class Functions
'===================================

Public Function serialize() As String
'-------------------------------------------------------------------
' Create serialized Json string from the current browser instance.
'-------------------------------------------------------------------

    If Not boolSerializable Then err.Raise -904, Description:="This instance is not allowed to serialize."

    Dim objSerialize As New Dictionary
    
    objSerialize.Add "hStdOutRd", hStdOutRd
    objSerialize.Add "hStdInWr", hStdInWr
    objSerialize.Add "hCDPOutRd", hCDPOutRd
    objSerialize.Add "hCDPInWr", hCDPInWr
    
    serialize = jsConverter.ConvertToJson(objSerialize)

End Function


Public Function deserialize(strSerialized As String)
'--------------------------------------------------------------------
' Decipher serialized Json string and assign values to CDP variables.
'--------------------------------------------------------------------

    Dim objSerialize As Dictionary
    Set objSerialize = jsConverter.ParseJson(strSerialized)
    
    hStdOutRd = objSerialize.Item("hStdOutRd")
    hStdInWr = objSerialize.Item("hStdInWr")
    hCDPOutRd = objSerialize.Item("hCDPOutRd")
    hCDPInWr = objSerialize.Item("hCDPInWr")

End Function


Public Function init(strExec As String, Optional aboolSerializable As Boolean = False) As Integer
'-------------------------------------------------------------------
' Initialize new browser instance via pipe creation.
'-------------------------------------------------------------------

    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As LongPtr, hWritePipe As LongPtr
    Dim L As Long, result As Long, bSuccess As Long
    Dim Buffer As String
    Dim k As Long
    
    Dim pipes As STDIO_BUFFER
    Dim pipes2 As STDIO_BUFFER2
    
    boolSerializable = aboolSerializable

   'First we create all 4 pipes
   'We start with stdout of the edge process
   'This pipe is used for stderr, too
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    result = CreatePipe(hStdOutRd, hStdOutWr, sa, 0)
    If result = 0 Then
        init = -2
        Exit Function
    End If
    
   'Then stdin
    result = CreatePipe(hStdInRd, hStdInWr, sa, 0)
    If result = 0 Then
        init = -2
        Exit Function
    End If
    
   'Then the out pipe for the CDP Protocol
    result = CreatePipe(hCDPOutRd, hCDPOutWr, sa, 2 ^ 20)
    If result = 0 Then
        init = -2
        Exit Function
    End If
    
   'And finally the in pipe
    result = CreatePipe(hCDPInRd, hCDPInWr, sa, 0)
    If result = 0 Then
        init = -2
        Exit Function
    End If
    
   'Then we fill the special structure for passing arbitrary pipes (i.e. fds) to a process
    pipes.number_of_fds = 5
    pipes.os_handle(0) = hStdInRd
    pipes.os_handle(1) = hStdOutWr
    pipes.os_handle(2) = hStdOutWr
    pipes.os_handle(3) = hCDPInRd
    pipes.os_handle(4) = hCDPOutWr
    pipes.crt_flags(0) = 9
    pipes.crt_flags(1) = 9
    pipes.crt_flags(2) = 9
    pipes.crt_flags(3) = 9
    pipes.crt_flags(4) = 9
    
   'pipes2 is filled by copying memory from pipes
    pipes2.number_of_fds = pipes.number_of_fds
    
    MoveMemory pipes2.raw_bytes(0), pipes.crt_flags(0), 5
    MoveMemory pipes2.raw_bytes(5), pipes.os_handle(0), UBound(pipes2.raw_bytes) - 4
    
   'Define new instance parameters and start it
    With start
        .cb = Len(start)
        .dwFlags = STARTF_USESHOWWINDOW
        .hStdOutput = hStdOutWr
        .hStdInput = hStdInRd
        .hStdError = hStdOutWr
        .wShowWindow = SW_SHOWMINIMIZED 'minimize the window at start
        .cbReserved2 = Len(pipes2)
        .lpReserved2 = VarPtr(pipes2)
    End With
    
    result = CreateProcessA(0&, strExec, sa, sa, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    If result = 0 Then init = -1
    
   'We close the sides of the handles that we dont need anymore
    CloseHandle hStdOutWr
    CloseHandle hStdInRd
    CloseHandle hCDPOutWr
    CloseHandle hCDPInRd
    
    hProcess = proc.hProcess
    init = 0
    
End Function


Public Function readProcCDP(ByRef strData As String) As Long
' --------------------------------------------------------------------------------
' This function tries to read from the CDP out pipe.
' Reading is non-blocking, if there are no bytes ro read the function returns 0,
' otherwise the number of bytes read.
'---------------------------------------------------------------------------------

    Dim lPeekData As Long
    Dim lngRes As Long
    
    lngRes = PeekNamedPipe(hCDPOutRd, ByVal 0&, 0&, ByVal 0&, lPeekData, ByVal 0&)
    If lngRes = 0 Then err.Raise 901, Description:="Error PeekNamedPipe in readProcCDP."
    
    If lPeekData > 0 Then
        Dim Buffer As String
        Dim L As Long
        Dim bSuccess As Long
        Buffer = Space$(lPeekData)
        bSuccess = ReadFile(hCDPOutRd, Buffer, Len(Buffer), L, 0&)
        
        If bSuccess = 1 Then
            strData = Buffer
            readProcCDP = Len(strData)
        Else
            readProcCDP = -2
        End If
        
    Else
        readProcCDP = -1
    End If

End Function


Public Function readProcSTD(ByRef strData As String) As Integer
' --------------------------------------------------------------------------------
' Same as ReadProcCDP.
'---------------------------------------------------------------------------------

    Dim lPeekData As Long
    
    PeekNamedPipe hStdOutRd, ByVal 0&, 0&, ByVal 0&, lPeekData, ByVal 0&
    
    If lPeekData > 0 Then
        Dim Buffer As String
        Dim L As Long
        Dim bSuccess As Long
        Buffer = Space$(lPeekData)
        bSuccess = ReadFile(hStdOutRd, Buffer, Len(Buffer), L, 0&)
        
        If bSuccess = 1 Then
            strData = Buffer
            readProcSTD = Len(strData)
        Else
            readProcSTD = -2
        End If
        
    Else
        readProcSTD = -1
    End If

End Function


Public Function writeProc(ByVal strData As String) As Integer
' --------------------------------------------------------------------------------
' This functions sends a CDP message to the Chromium browser.
'---------------------------------------------------------------------------------

    Dim lngWritten As Long
    
    writeProc = WriteFile(hCDPInWr, ByVal strData, Len(strData), lngWritten, ByVal 0&)

End Function


Private Sub Class_Terminate()
' --------------------------------------------------------------------------------
' Auto-run when the class is closed.
'---------------------------------------------------------------------------------

    If boolSerializable Then Exit Sub

    CloseHandle hStdOutRd
    CloseHandle hStdOutWr
    CloseHandle hStdInRd
    CloseHandle hStdInWr

    CloseHandle hCDPOutRd
    CloseHandle hCDPOutWr
    CloseHandle hCDPInRd
    CloseHandle hCDPInWr

End Sub
