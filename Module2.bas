Attribute VB_Name = "Module2"
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const ERROR_SUCCESS = 0&
Const REG_OPTION_NON_VOLATILE = &O0
Const KEY_ALL_CLASSES As Long = &HF0063
Const KEY_ALL_ACCESS = &H3F
Const REG_SZ As Long = 1

Public Function RegGetString$(hInKey As Long, ByVal subkey$, ByVal valname$)
    Dim RetVal$, hSubKey As Long, dwType As Long, SZ As Long, v$, r As Long
    
    RetVal$ = ""
    
    r = RegOpenKeyEx(hInKey, subkey$, 0, KEY_ALL_CLASSES, hSubKey)
    If r <> ERROR_SUCCESS Then GoTo Quit_Now
    SZ = 256: v$ = String$(SZ, 0)
    r = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal v$, SZ)
    If r = ERROR_SUCCESS And dwType = REG_SZ Then
        RetVal$ = Left(v$, SZ - 1)
    Else
        RetVal$ = ""
    End If
    If hInKey = 0 Then r = RegCloseKey(hSubKey)
Quit_Now:
    RegGetString$ = RetVal$

End Function

Public Sub ConnectW3(url$)
On Error GoTo fout_connectw3

    Dim strProgram$, strDDETopic$, strDDEItem$
    Dim intLoaded%

'make on Form1 a invisible textbox named DDEText
    strProgram = RegGetString(HKEY_CLASSES_ROOT, "http\shell\open\command", "")
    strDDETopic = UCase(RegGetString(HKEY_CLASSES_ROOT, "http\shell\open\ddeexec\Application", "")) & "|" & RegGetString(HKEY_CLASSES_ROOT, "http\shell\open\ddeexec\Topic", "")
    strDDEItem = url$
    With Form4.DDEText
        .LinkTopic = strDDETopic
        .LinkItem = strDDEItem & ",," & -1
        .LinkMode = 2
        .LinkRequest
    End With
    Exit Sub
    
fout_connectw3:
    If Err.Number = 282 Then
        If intLoaded = 0 Then
            Shell strProgram, vbNormalFocus
            intLoaded = 1
        ElseIf intLoaded <= 5 Then
            intLoaded = intLoaded + 1
        Else
            Err.Number = vbObjectError + 1
            GoTo fout_connectw3
        End If
        Resume
    ElseIf Err.Number <> 0 Then
        MsgBox "Fatal error while communicating to browser"
        Exit Sub
    End If

End Sub

