VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DosStdOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
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
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Const NORMAL_PRIORITY_CLASS As Long = &H20
Private Const STARTF_USESTDHANDLES As Long = &H100
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const SW_HIDE As Long = 0

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function CreatePipe Lib "kernel32.dll" (phReadPipe As Long, _
phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long

Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, _
ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, _
ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, _
ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, _
lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, _
lpProcessInformation As PROCESS_INFORMATION) As Long
    
Private m_Outputs As String

Public Property Get Outputs()
    Outputs = m_Outputs
End Property

Public Function DosPipe(lpFileCommand As String, Optional PipeReadSize As Long) As Boolean
Dim si As STARTUPINFO, pi As PROCESS_INFORMATION, sa As SECURITY_ATTRIBUTES
Dim hPipe As Long
Dim lBytesRead As Long
Dim StrBuffer As String
Dim StrA As String
Dim iRet As Long

Dim hReadPipe As Long
Dim hWritePipe As Long

On Error Resume Next

    DosPipe = False
    StrBuffer = Space(PipeReadSize)

    If Len(lpFileCommand) = 0 Then
        Exit Function
        'No lpFileCommand was found
    End If
    
    If (PipeReadSize = 0) Then
        'Exit if no PipeReadSize was found
        Exit Function
    End If
    
    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1&
        .lpSecurityDescriptor = 0&
    End With
    ' Create a new pipe
    hPipe = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    
    If (hPipe = 0) Then
        'if pipe is zero we exit
        Exit Function
    Else
        'Fill in the start up information
        With si
            .cb = Len(si) 'Size of this type
            .dwFlags = (STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW)
            'Output pipe hangle
            .hStdOutput = hWritePipe
            .hStdError = hWritePipe
            'Hide window
            .wShowWindow = SW_HIDE
        End With
    End If

    iRet = CreateProcessA(0&, lpFileCommand, sa, sa, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, si, pi)
              
    If (iRet = 0) Then
        Exit Function
    Else
        'Close the WritePipe
        CloseHandle hWritePipe
        Do
            iRet = ReadFile(hReadPipe, StrBuffer, PipeReadSize, lBytesRead, 0&)
            StrA = StrA & Left(StrBuffer, lBytesRead)
        Loop While (iRet <> 0)
    End If
    
    'Close the hangles
    CloseHandle pi.hProcess
    CloseHandle pi.hThread
    CloseHandle hReadPipe
    
    m_Outputs = StrA
    StrA = vbNullString
    StrBuffer = vbNullString
    lBytesRead = 0
    DosPipe = True
    
End Function

Private Sub Class_Initialize()
    m_Outputs = vbNullString
End Sub

Private Sub Class_Terminate()
    m_Outputs = vbNullString
End Sub
