Attribute VB_Name = "modxrunas"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long


Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(0) As LUID_AND_ATTRIBUTES
End Type
Private Type PRIVILEGES_SET
    PrivilegeCount As Long
    Control As Long
    Privileges(0) As LUID_AND_ATTRIBUTES
End Type
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" ( _
     ByVal ProcessHandle As Long, _
     ByVal DesiredAccess As Long, _
     ByRef TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" ( _
     ByVal lpSystemName As Long, _
     ByVal lpName As String, _
     ByRef lpLuid As LUID) As Long
Private Declare Function PrivilegeCheck Lib "advapi32.dll" ( _
     ByVal TokenHandle As Long, _
     ByRef RequiredPrivileges As PRIVILEGES_SET, _
     ByRef pfResult As Long) As Long
' Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" ( _
'     ByVal TokenHandle As Long, _
'     ByVal DisableAllPrivileges As Long, _
'     ByRef NewState As TOKEN_PRIVILEGES, _
'     ByVal BufferLength As Long, _
'     ByRef PreviousState As Long, _
'     ByRef ReturnLength As Long) As Long
     
     
Public Function Nz(v, Optional nv = vbNullString)
Dim vt%: vt = VarType(v): If VarType(v) And vbArray Or vt < vbInteger Or vt = vbObject Or vt = 10 Then Nz = nv Else Nz = v
End Function

Public Function ShellEx(ByVal sPathName, Optional ByVal nShowCmd = 5, Optional ByVal sOperation, Optional ByVal sParameters, Optional ByVal sDirectory) As Long
On Error Resume Next
Dim sh As SHELLEXECUTEINFO
sh.cbSize = LenB(sh)
sh.fMask = &H40& 'SEE_MASK_NOCLOSEPROCESS (0x00000040)
'sh.hwnd = 0
sh.lpVerb = "" & sOperation
sh.lpFile = "" & sPathName
sh.lpParameters = "" & sParameters
sh.lpDirectory = "" & sDirectory
sh.nShow = nShowCmd
'sh.hInstApp = 0
ShellEx = ShellExecuteEx(sh)

End Function

Sub Main()
Dim b&, s$: s = Trim(Command)
's = "D:\ADPVODA\XVB\xTypeLib.dll"
's = "D:\ADPVODA\XVB\xTemplate.dlg"
's = "D:\ADPVODA\XVB\want.JPG"
's = "D:\ADPVODA\XVB\RICHED_20.DLL"
's = "/s D:\ADPVODA\XVB\XVBEngine.tlb"
's = "/s D:\ADPVODA\XVB\XVBEngine.exe"
    'MsgBox s


Dim q1&, q& 'MSBOX RESULT
If InStr(s, "/s ") Then q1 = 1: q = vbYes: s = Trim(Replace(s, "/s ", "", , 1))
If InStr(s, "/u ") Then q = vbNo: s = Trim(Replace(s, "/u ", "", , 1))
'MsgBox s & vbCrLf & "q=" & IIf(q, q = vbYes, "EMPTY") & " q1=" & q1
If Len(s) Then
    On Error Resume Next

    If SEPrivilege("SeImpersonatePrivilege") = 0 Then ShellEx App.Path & "\" & App.EXEName & ".exe", 5, "runas", s: Exit Sub
    If LCase(Right(s, 13)) = "xvbengine.exe" Then If InstallXVB(q, q1) Then Exit Sub  'setup ActiveX
    
    Select Case LCase(IIf(InStr(s, " "), "", Right(s, 3)))
    Case "tlb"
        If q Then b = q Else b = MsgBox("Register = " & s, vbYesNoCancel + vbQuestion, "TLB Register/UnRegister ")
        If b <> vbCancel Then TLBRegisterUnRegister s, b = vbYes, 0, q1
    Case "dll", "ocx"
        If q Then b = q Else b = MsgBox("Register = " & s, vbYesNoCancel + vbQuestion, "DLL/OCX Register/UnRegister ")
        If b <> vbCancel Then Shell "regsvr32.exe " & IIf(b = vbYes, "", "/u ") & IIf(q1, "/s ", "") & s
    Case Else

        If q1 = 0 Then ShellEx s
    End Select
    
Else 'len(s)
    'MsgBox "elevate:" & vbCrLf & "xrunas path_to_exe[,params]" & vbCrLf & "tlb register/unregister" & vbCrLf & "dll/ocx register/unregister"
    Load xrunasfrm
    'xrunasfrm.Show 1
    xrunasfrm.Visible = 1
End If
End Sub

Function TLBRegisterUnRegister(ByVal tlb_pth$, bRegister As Boolean, bxvb As Boolean, q1 As Long) As Boolean
On Error Resume Next
Dim tlb As ITypeLib: Set tlb = LoadTypeLib(tlb_pth)
If Not tlb Is Nothing Then
    If bRegister Then
        RegisterTypeLib tlb, tlb_pth, vbNullString
    Else
        If bxvb Then
            Dim tia As TYPEATTR
            CopyMemory tia, ByVal tlb.GetTypeInfo(1).GetTypeAttr, Len(tia)
            RegValue(0, "CLSID\" & SysAllocString(StringFromCLSID(tia.IID)), "") = Null
            RegValue(0, "XVBEngine.Public", "") = Null
        End If
        Dim ta As TLIBATTR: CopyMemory ta, ByVal tlb.GetLibAttr, Len(ta)
        UnRegisterTypeLib ta.IID, ta.wMajorVerNum, ta.wMinorVerNum, ta.lcid, SYS_WIN32
    End If
End If
TLBRegisterUnRegister = Not bRegister
If q1 = 0 Then MsgBox tlb_pth & vbCrLf & IIf(bRegister, "", "Un") & "Register = " & IIf(Err, Err.Description, "OK"), vbInformation
Err.Clear
End Function

Function InstallXVB(q&, q1 As Long) As Long
Dim s$: s = "XVBEngine.tlb"
If Len(ResRestore(s)) Then
    Dim b&: If q Then b = q Else b = MsgBox("Register = " & s, vbYesNoCancel)
    If b <> vbCancel Then InstallXVB = TLBRegisterUnRegister(App.Path & "\" & s, b = vbYes, 1, q1)
'Kill App.Path & "\" & s
End If
End Function

Function ResRestore(ByVal s$) As String
On Error Resume Next
Dim st As New Stream: st.Type = adTypeBinary: st.Open: st.Write LoadResData(s, "CUSTOM")
If st.Size = 0 Then Exit Function
ResRestore = s
If st.Size <> LenFile(App.Path & "\" & s) Then st.SaveToFile App.Path & "\" & s, adSaveCreateOverWrite
End Function

Function LenFile&(ByVal pth$)
On Error Resume Next
'LenFile = FileLen(Replace("" & pth, "..\", xAppPath & "\"))
If InStr(pth, "\") Then LenFile = FileLen(pth)
Err.Clear
End Function
Function SEPrivilege(ByVal sPrivilege$) As Long
Dim hToken&
Dim tp As TOKEN_PRIVILEGES  ' token privileges
Dim tps As PRIVILEGES_SET  ' token privileges
Dim LUID As LUID
Dim ret&, v&
'Dim b(15) As Long
'Const TOKEN_QUERY = &H8
'Const TOKEN_ADJUST_PRIVILEGES = &H20
'SE_PRIVILEGE_ENABLED  (it is 0x00000002L in WinNT.h)
'SE_PRIVILEGE_ENABLED_BY_DEFAULT  (it is 0x00000001L in WinNT.h)
'SE_PRIVILEGE_REMOVED  (it is 0x00000004L in WinNT.h)
'SE_PRIVILEGE_USED_FOR_ACCESS  (it is 0x80000000L in WinNT.h)
v = -1: ret = -1

'If Len(sPrivilege) = 0 Then Exit Function
If OpenProcessToken(GetCurrentProcess(), &H8, hToken) Then
    If LookupPrivilegeValue(0, sPrivilege, LUID) Then
        tps.PrivilegeCount = 1: tps.Control = 1: tps.Privileges(0).pLuid = LUID
        If PrivilegeCheck(hToken, tps, ret) Then
'             If v > -1 Then 'LET
'                 If (ret <> v) Then
'                    tp.PrivilegeCount = 1: tp.Privileges(0).pLuid = LUID
'                    tp.Privileges(0).Attributes = IIf(v, 2, 0)
'                    If AdjustTokenPrivileges(hToken, False, tp, LenB(tp), 0, 0) Then
'                        tps.Privileges(0).Attributes = 0: ret = -1
'                        If PrivilegeCheck(hToken, tps, ret) Then 'get updated privilege value
'                            'SEPrivilege = ret
'                            'Debug.Print "LET "; sPrivilege; " ="; v, v = ret
'                        End If
'                    End If
'                Else
'                    'Debug.Print "NO LET "; sPrivilege; " ="; v, v = ret
'                End If
'            Else
'                'Debug.Print "GET "; sPrivilege; " ="; SEPrivilege
'            End If
        End If
    End If
    CloseHandle hToken
End If
SEPrivilege = ret
End Function





Property Let RegValue(ByVal root, ByVal sSection, ByVal sKey, v) ' v=Null = DELETE KEY
'    HKEY_CLASSES_ROOT = &H80000000
'    HKEY_CURRENT_USER = &H80000001
'    HKEY_LOCAL_MACHINE = &H80000002
'    HKEY_USERS = &H80000003
Dim hRoot&, hKey&, strBuf$, lBuf&, m
hRoot = CW_USEDEFAULT Or (CLng(root) And 7&)
If VarType(v) = vbNull Then 'DELETE
    If RegOpenKeyEx(hRoot, sSection, 0&, 6&, hKey) Then Exit Property
    If sKey = "" Then
        For Each m In RegKeys(root, sSection)
            RegValue(root, sSection & "\" & m, "") = Null
        Next
        lBuf = RegDeleteKey(hKey, ""): 'DELETE SECTION
    Else
        lBuf = RegDeleteValue(hKey, Replace(sKey, Chr(0), "")) 'DELETE KEY
    End If

Else 'CREATE UPDATE
    If RegCreateKey(hRoot, sSection, hKey) Then Exit Property
    Select Case VarType(v)
    Case vbEmpty 'Удаление значения по умолчанию
        RegDeleteValue hKey, sKey
    Case vbObject
        RegSetValueEx hKey, sKey, 0, 0, 0, 0
    Case vbLong, vbByte, vbInteger, vbBoolean, vbDate  'DWORD
        lBuf = v: RegSetValueEx hKey, sKey, 0, 4, lBuf, 4
    Case Else 'STRING
        strBuf = v
        RegSetValueEx hKey, sKey, 0, IIf(InStr(strBuf, Chr(0)), 3, 2), ByVal strBuf, Len(strBuf)
    End Select
    RegCloseKey hKey
End If
End Property

Property Get RegValue(ByVal root, ByVal sSection, ByVal sKey) 'as Value or Null if no exist
'    HKEY_CLASSES_ROOT = &H80000000
'    HKEY_CURRENT_USER = &H80000001
'    HKEY_LOCAL_MACHINE = &H80000002
'    HKEY_USERS = &H80000003

'Private Const REG_BINARY As Long = 3
'Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4
'Private Const REG_DWORD_BIG_ENDIAN As Long = 5
'Private Const REG_EXPAND_SZ As Long = 2
'Private Const REG_LINK As Long = 6
'Private Const REG_MULTI_SZ As Long = 7
'Private Const REG_NONE As Long = 0
'Private Const REG_RESOURCE_LIST As Long = 8
'Private Const REG_SZ As Long = 1



Dim hRoot&, hKey&
Dim lValueType&, lResult&, lDataBufSize&
Dim strBuf$
hRoot = CW_USEDEFAULT Or (CLng(root) And 7&)  ' -(2 ^ 31)=&H80000000
RegValue = Null
If RegOpenKeyEx(hRoot, sSection, 0&, 1&, hKey) Then Exit Property
If RegQueryValueEx(hKey, sKey, 0&, lValueType, ByVal 0&, lDataBufSize) = 0 Then
    Select Case lValueType
    Case 1, 2, 3, 6, 7 'REG_SZ,  REG_EXPAND_SZ, REG_BINARY, REG_LINK, REG_MULTI_SZ
        strBuf = String$(lDataBufSize, 0)
        If RegQueryValueEx(hKey, sKey, 0, 0, ByVal strBuf, lDataBufSize) = 0 Then
        
            If lValueType = 3 Then 'REG_BINARY
                RegValue = Left$(strBuf, lDataBufSize)
            Else 'Отрезаем ch0ch0 или ch0 в конце строки
                If Len(strBuf) Then RegValue = Left$(strBuf, InStr(strBuf, IIf(lValueType = 7, Chr$(0), vbNullString) & Chr$(0)) - 1)
            End If
            
        End If
    Case 4, 5 'REG_DWORD, REG_DWORD_BIG_ENDIAN
        If RegQueryValueEx(hKey, sKey, 0&, lValueType, lResult, lDataBufSize) = 0 Then RegValue = lResult
    End Select
    RegCloseKey hKey
End If
End Property


Function RegKeys(ByVal root, ByVal sSection) ' As Collection
'    HKEY_CLASSES_ROOT = &H80000000
'    HKEY_CURRENT_USER = &H80000001
'    HKEY_LOCAL_MACHINE = &H80000002
'    HKEY_USERS = &H80000003

'returns a list of sub keys in the current section
Dim hRoot&, hKey&
hRoot = CW_USEDEFAULT Or (CLng(root) And 7&)
Dim car 'c$ ' As New Collection
Dim Cnt As Long, sName As String, ret As Long
Const BUFFERSIZE As Long = 255
ret = BUFFERSIZE
car = Array()
'Open the registry key
If RegOpenKeyEx(hRoot, sSection, 0&, 8&, hKey) = 0 Then
    sName = Space(BUFFERSIZE)
    While RegEnumKeyEx(hKey, Cnt, sName, ret, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&) <> 259&
        'gAddIndex car, Left(sName, ret)
        ReDim Preserve car(Cnt)
        car(Cnt) = Left(sName, ret)
        Cnt = Cnt + 1
        sName = Space(BUFFERSIZE)
        ret = BUFFERSIZE
    Wend
    RegCloseKey hKey
End If
RegKeys = car
End Function

Property Get SendTo(ByVal pth$) As String
Dim f, s$, wsh, sh
Set wsh = CreateObject("WScript.Shell")
For Each f In CreateObject("Scripting.FileSystemObject").GetFolder(wsh.SpecialFolders("SendTo")).Files
s = f: If LCase(Right(s, 3)) = "lnk" Then s = wsh.CreateShortcut(s).TargetPath
'Set sh = wsh.CreateShortcut(s)
's = sh.TargetPath
'End If
If StrComp(pth, s, vbTextCompare) = 0 Then SendTo = f: Exit For
'Debug.Print s
Next
End Property

Property Let SendTo(ByVal pth$, v$)
Dim ar, wsh, sh, p$
If Len(pth) = 0 Then Exit Property
On Error Resume Next
Set wsh = CreateObject("WScript.Shell")
p = wsh.SpecialFolders("SendTo")
Set sh = wsh.CreateShortcut(p & "\" & pth & ".lnk")
If Len(v) Then 'CREATE UPDATE
    ar = Split(v & vbCrLf & vbCrLf, vbCrLf)
    With sh
        .TargetPath = ar(0)
'        .Arguments = ar(1)
'        .IconLocation = ar(2)
        .Save
    End With
Else 'REMOVE
    p = sh.fullname
    Kill p
End If
End Property

