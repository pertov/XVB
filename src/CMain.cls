VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'Private Type synfmt
'fn(31) As Byte 'font name
'fu As Byte 'font underline
'fi As Byte 'font italic
'fb As Byte 'font bold
'sz As Byte 'font size
'
'bcb As Byte 'back blue
'bcg As Byte 'back green
'bcr As Byte 'back red
'
'fcb As Byte 'color blue
'fcg As Byte 'color green
'fcr As Byte 'color red
'End Type

Private Type syndef2
lang As String 'script language
spc As String 'pattern for comments, strings  directives
rm As String 'remove from code /*
ms As String 'multiline string start $"
sq As String 'strings quot char  "
sj As String 'strings join  +vbCrLf+    +
msj As String 'multiline join    \
ps As String 'pattern for procedures
'fp As String 'replace from code float
sp As String 'precedure separator   :=58 ;=59
pf As String 'pf pattern=  ;function f_%_(){;@;return(%);};    ;function f_%_():@:f_%_=0:end function:
End Type
'Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long

'Private Type MD5_CTX
'  i(1 To 2) As Long
'  buf(1 To 4) As Long
'  inp(1 To 64) As Byte
'  digest(1 To 16) As Byte
'End Type
'Private Declare Sub MD5Init Lib "cryptdll.dll" (Context As MD5_CTX)
'Private Declare Sub MD5Update Lib "cryptdll.dll" (Context As MD5_CTX, ByVal strInput As String, ByVal lLen As Long)
'Private Declare Sub MD5Final Lib "cryptdll.dll" (Context As MD5_CTX)

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

'Public Enum TIME_ZONE
'    TIME_ZONE_ID_INVALID = 0
'    TIME_ZONE_STANDARD = 1
'    TIME_ZONE_DAYLIGHT = 2
'End Enum
' Required Windows API Declares
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByRef lpBuffer As Long, ByVal nSize As Long, ByVal Arguments As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

'Const MAX_PATH As Long = 260&


Private m_DbgWnd As Long

Private m_LogWnd As Long

Private m_bytes() As Byte

'Private mVB6 As Collection
'Private Declare Function WaitMessage Lib "user32.dll" () As Long
Private Declare Sub InternetSetCookie Lib "wininet.dll" Alias "InternetSetCookieA" ( _
     ByVal lpszUrl As String, _
     ByVal lpszCookieName As String, _
     ByVal lpszCookieData As String)
Private Declare Sub InternetGetCookie Lib "wininet.dll" Alias "InternetGetCookieA" ( _
    ByVal lpszUrl As String, _
    ByVal lpszCookieName As String, _
    ByVal lpCookieData As String, _
    ByRef lpdwSize As Long)
     
Private Declare Sub DeleteUrlCacheEntry Lib "wininet.dll" (ByVal lpszUrlName As String)
     
'Private Declare Function SetProcessPreferredUILanguages Lib "kernel32.dll" ( _
'    ByVal dwFlags As Long, _
'    ByVal pwszLanguagesBuffer As String, _
'    pulNumLanguages As Long) As Long

Public Event DumyEvent(Cookie, args)
'Private Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


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


'Private Declare Sub ObjectFromLresult Lib "OLEACC.dll" ( _
     ByVal lResult As Long, _
      riid As UUID, _
     ByVal wParam As Long, _
     ByRef ppvObject As Any)
     
''============CallInterface====================
'Private Const GMEM_FIXED As Long = &H0
'Private Const asmPUSH_imm32 As Byte = &H68
'Private Const asmRET_imm16 As Byte = &HC2
'Private Const asmCALL_rel32 As Byte = &HE8
''============CallInterface====================

'=================  RUN AND WAIT   ===============
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
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
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long

Private Declare Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (lpszSrc As Any, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long

Private Declare Function CreateProcessWithLogon Lib "advapi32" Alias "CreateProcessWithLogonW" (ByVal lpUsername As Long, ByVal lpDomain As Long, ByVal lpPassword As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInfo As PROCESS_INFORMATION) As Long

'=================  RUN AND WAIT   ===============


'Private Declare Function GetProcessVersion Lib "kernel32.dll" (ByVal ProcessId As Long) As Long

'Private Declare Function apiFindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function apiFindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long


'Private Type SntpVersion3Msg
'    BitFields As Long
'    RootDelay As Long
'    RootDispersion As Long
'    ReferenceIdentifier(3) As Byte
'    ReferenceTimestampBE(1) As Long
'    OriginateTimestampBE(1) As Long
'    ReceiveTimestampBE(1) As Long
'    TransmitTimestamp(1) As Long
'End Type
'Private Declare Function ntohl Lib "ws2_32" (ByVal netlong As Long) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Private Declare Function crc32 Lib "zlib.dll" (ByVal crc As Long, ByVal buf As Long, ByVal buflen As Long) As Long  'adler32

'Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As IBindStatusCallback) As Long
'Private Declare Sub DeleteUrlCacheEntry Lib "wininet.dll" (ByVal lpszUrlName As String)


'Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
'Private Declare Function CallByNameAndArgsArray Lib "msvbvm60.dll" Alias "rtcCallByName" (ByVal Object As Object, ByVal ProcName As Long, ByVal CallType As VbCallType, args() As Variant, Optional ByVal lcid As Long) As Variant
'Public Declare Function rtcCallByName Lib "MSVBVM60.dll" (ByVal Object As Object, ByVal ProcName As Long, ByVal CallType As VbCallType, ByRef args() As Any, Optional ByVal lcid As Long) As Variant




'Private Declare Function rtcCallByName Lib "MSVBVM60.dll" (ByVal Object As Object, ByVal ProcName As Long, ByVal CallType As VbCallType, Optional ByVal pArgs As Long, Optional ByVal lcid As Long) As Variant

    
'Private m_Storage As Collection 'Для хранения разных объектов во время исполнения
Private m_Buffer As CParam 'ClipBoard
Private m_Param As New CParam 'Common CParam
Private mRegEx As RegExp 'Object ' 'SYNTAX
Private cpRegEx As CParam 'SYNTAX

Private mSC As ScriptControl 'INNER ScriptControl

'Private m_Connection As Connection  'Server conection
Private m_Cons As New Collection  'Conections

'Public mSchemaColumns As CParam ' SchemaColumns TABLE,PKNAME

Private m_Base As CBase 'Base properties
Private m_iniBase As String

Public JETProvider As String
Const script_engine_ = "script_engine_"

Private WithEvents tm As CTimer
Attribute tm.VB_VarHelpID = -1
Private arCmd As New Collection

'Private WithEvents mCloseTimer As CTimer
'Private ctmar

'Private WithEvents tmWait As CTimer
'Private tmWaitParams() As WAITPARAMS


'Public Port& 'Номер порта для обмена событиями

'Private gCollection As New Collection
'Private gFunctions As New CParam
Public xObject As Object 'Нужен для IUnknown
'Public xObjectHTTP As WinHttpRequest

Private Type SERROR
    Number As Long
    Description As String
End Type
Private mError As SERROR

'Private WithEvents http As WinHttpRequest
'Private http_handle As xControl



'Const PROCESS_QUERY_INFORMATION = 1024
'Const PROCESS_VM_READ = 16

'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal pid As Long) As Long
'Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

'Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)


Private Declare Function VarTokenizeFormatString Lib "oleaut32.dll" (ByVal pstrFormat As Long, ByRef rgbTok As Any, ByVal cbTok As Long, ByVal iFirstDay As VbDayOfWeek, ByVal iFirstWeek As VbFirstWeekOfYear, ByVal lcid As Long, Optional ByVal pcbActual As Long) As Long
Private Declare Function VarFormatFromTokens Lib "oleaut32.dll" (ByRef pvarIn As Variant, ByVal pstrFormat As Long, ByRef pbTokCur As Any, ByVal dwFlags As Long, ByVal pbstrOut As Long, ByVal lcid As Long) As Long
 
Function aVal(ByVal vExpression)
On Error Resume Next
aVal = 0
aVal = Val(Replace$(Replace$(Replace(vExpression, ",", "."), " ", vbNullString), Chr$(160), vbNullString))
Err.Clear
'Dim vt&
'vt = VarType(vExpression)
'If vt < 2 Or vt = 10 Or (vt And vbArray) > 0 Then aVal = 0 Else aVal = Val(Replace$(Replace$(Replace(vExpression, ",", "."), " ", vbNullString), Chr$(160), vbNullString))
End Function
Public Function num(ByVal v) As Double
On Error Resume Next: num = Val("" & v)
End Function
Public Function IFF(Expression, Optional TruePart = Null, Optional FalsePart = Null)
On Error Resume Next
'IFF = IIf(Expression,  TruePart, FalsePart)
If Expression Then IFF = TruePart Else IFF = FalsePart
End Function

Public Function IF0(Expression, Optional Not0Value = Null)
On Error Resume Next
IF0 = Expression
If aVal(Expression) = 0 Then IF0 = Not0Value
End Function


Public Function Nz(Expression, Optional NotNullExpression = vbNullString)
Dim vt&: vt = VarType(Expression)
If vt < 2 Or vt = 10 Or (vt And vbArray) > 0 Then Nz = NotNullExpression Else Nz = Expression
'If vt < 2 Or vt = 10 Or (vt And vbArray) > 0 Then CopyMemory Nz, NotNullExpression, 16 Else CopyMemory Nz, Expression, 16
End Function

'Public Function NzS(Expression, Optional NotZeroString = vbNullString)
Public Function NzS(Expression, Optional NotNullString = vbNullString, Optional sPrefix = vbNullString, Optional sSuffix = vbNullString)
Dim vt&: vt = VarType(Expression)
If vt < 2 Or vt = 10 Or (vt And vbArray) > 0 Then
    NzS = NotNullString
Else
    If Len(Expression) Then NzS = Expression Else NzS = NotNullString
End If
If Len(vbNullString & NzS) Then NzS = sPrefix & NzS & sSuffix
End Function


Function CallByNameEx(ByVal Object As Object, ByVal ProcName As String, ByVal CallType As Long, ParamArray args())
On Error Resume Next
Dim ar(): ar = args
CallByNameEx = rtcCallByName(Object, StrPtr(ProcName), CallType, ar)
If Err Then Err.Clear
End Function

'Function PLet(x, ByVal p, ByVal v)
'On Error Resume Next
'If Not IsEmpty(v) Then CallByName x, p, VbLet, v
'PLet = CallByName(x, p, VbGet)
'End Function
Public Function Cast(vValue, ByVal vt, Optional ByVal bAcceptNull As Boolean)
Cast = modMain.Cast(vValue, vt, bAcceptNull)
End Function
Public Function CastAny(vValue, ByVal vt, Optional ByVal bAcceptNull As Boolean)
CastAny = modMain.CastAny(vValue, vt, bAcceptNull)
End Function




Public Function RegEx(Optional ByVal MIGPattern) As RegExp
Set RegEx = New RegExp
RegEx.Global = True: RegEx.IgnoreCase = True: RegEx.MultiLine = True
Dim pattern$: pattern = S_(MIGPattern)
RegEx.pattern = pattern
If Len(pattern) < 3 Then Exit Function
If Left$(pattern, 1) <> "~" Then Exit Function
'If InStr("01234567", Mid(pattern, 2, 1)) = 0 Then Exit Function
Dim MIG&: MIG = Val(Mid(pattern, 2, 1)): RegEx.pattern = Mid(pattern, 3)
RegEx.Global = MIG And 1: RegEx.IgnoreCase = MIG And 2: RegEx.MultiLine = MIG And 4
End Function

Public Function RegExMatch(ByVal Source, ByVal pattern, Optional ByVal nFromPos, Optional rx As RegExp)
'nFromPos
'   ?>0     Return Matches(0).FirstIndex start from pos
'   ?=0     Return Matches(0).Value
'   ?=-1   Return Matches Object
'   ?<-1   Return Matches(0).SubMatches(-nFromPos-2)
Dim rex As RegExp
RegExMatch = vbNullString
nFromPos = L_(nFromPos)
'If nFromPos Then If nFromPos > 0 Then RegExMatch = 0 Else Set RegExMatch = New Collection
If nFromPos > 0 Then RegExMatch = 0
If nFromPos = -1 Then Set RegExMatch = New Collection
pattern = S_(pattern): Source = S_(Source)
If Len(pattern) > 0 And Len(Source) > 0 Then
    If rx Is Nothing Then
        Set rex = RegEx(pattern)
    Else
        Set rex = rx: rex.pattern = pattern
    End If
    On Error Resume Next
    If nFromPos > 0 Then Source = Mid(Source, nFromPos)
    If Not rex.Test(Source) Then Exit Function
    Dim m As MatchCollection
    Set m = rex.Execute(Source)
    If nFromPos = 0 Then
        RegExMatch = m.Item(0).Value
    ElseIf nFromPos = -1 Then
        Set RegExMatch = m
    ElseIf nFromPos > 0 Then 'InStr
        RegExMatch = nFromPos + m.Item(0).FirstIndex
    Else
        If m.Item(0).SubMatches.count > -nFromPos - 2 Then RegExMatch = m.Item(0).SubMatches(-nFromPos - 2)
    End If
End If
End Function
Public Function RegExReplace(ByVal Source, ByVal pattern, ByVal str_replace, Optional sPreserveLength = "")
Dim rex As RegExp
RegExReplace = Source
If Len(pattern) > 0 And Len(Source) > 0 Then
    Set rex = RegEx(pattern)
    On Error Resume Next
    If Len(S_(sPreserveLength)) Then
        Dim ms, m, s$, i, res$
        res = Source
        Set ms = rex.Execute(res)
        For i = ms.count - 1 To 0 Step -1
            Set m = ms(i)
            s = Mid(res, m.FirstIndex + 1, m.length)
            res = Left(res, m.FirstIndex) & String(Len(s), sPreserveLength) & Mid(res, m.FirstIndex + 1 + m.length)
        Next
        RegExReplace = res
    Else
        RegExReplace = rex.Replace(Source, str_replace)
    End If
End If
End Function

Public Property Get Base() As CBase
Set Base = m_Base
End Property

Public Property Get MainWnd() As xControl
Set MainWnd = xMainWnd
End Property
Public Function xActiveControl() As xControl
'Set xActiveControl = hxControl(GetFocus())
Dim p&: p = GetFocus
Do While (p <> 0) And (xActiveControl Is Nothing)
    Set xActiveControl = hxControl(p)
    If xActiveControl Is Nothing Then p = GetParent(p)
Loop 'Ищем xControl если фркус был на чужом  окне
End Function
Function xControl(ByVal hWnd) As xControl
Set xControl = hxControl(0& + hWnd)
End Function
'Function xFocus()
'xFocus = GetFocus
'End Function

'Public Function TextDocument(ByVal hWnd) As ITextDocument
'Set TextDocument = ITextDocument(hWnd)
'End Function

Function HTMLDocument(ByVal h)
Dim s$, doc As Object, hWnd&
Dim lParam&, mUUID As UUID
DoEvents
hWnd = Cast(h, vbLong)
Do While hWnd <> 0 And s = ""
    s = pWindowClass(hWnd)
    If s = "Internet Explorer_Server" Then Exit Do Else hWnd = GetWindow(hWnd, GW_CHILD): s = ""
Loop

If hWnd Then
    SendMessageTimeout hWnd, RegisterWindowMessage("WM_HTML_GETOBJECT"), 0, 0, 2, 2000, lParam
    CLSIDFromString "{626FC520-A41E-11CF-A731-00A0C9082637}", mUUID
    ObjectFromLresult lParam, mUUID, 0, doc
End If
Set HTMLDocument = doc
End Function

'Public Function GetInterface(obj_or_pointer, ByVal sInterfaceID)
'Dim pobj As IUnknown
'Dim mUUID As UUID
'Dim IID$, lParam&
'IID = Nz(sInterfaceID)
'Dim IUnk As XVBHelpLib.IUnknown
''IID_IHTMLDocument = "{626FC520-A41E-11CF-A731-00A0C9082637}"
'On Error Resume Next
'If Len(IID) Then
'    If IsObject(obj_or_pointer) Then
'        Set IUnk = obj_or_pointer
'    Else
'        lParam = Cast(obj_or_pointer, vbLong)
'        GetMemObj lParam, IUnk
'        'ObjectFromLresult lParam, mUUID, 0, pobj
'    End If
'    CLSIDFromString IID, mUUID
'    IUnk.QueryInterface mUUID, pobj
'End If
'Set GetInterface = pobj
'Set pobj = Nothing
'End Function

Function ObjFromPtr(ByVal oPtr) As Object
On Error Resume Next
GetMemObj 0& + oPtr, ObjFromPtr
End Function
Function ObjToPtr&(obj)
On Error Resume Next
ObjToPtr = ObjPtr(obj)
End Function
Function ObjFromStrPtr(ByVal oPtr) As Object
On Error Resume Next
GetMemObj StringToLong(StrConv(oPtr, vbUnicode)), ObjFromStrPtr
End Function
Function ObjToStrPtr(obj)
On Error Resume Next
ObjToStrPtr = StrConv(LongToString(ObjPtr(obj)), vbFromUnicode)
End Function

'Function VarToPtr&(var)
'On Error Resume Next
'VarToPtr = VarPtr(var)
'End Function
Public Function FindWFP&(ByVal ph, Optional xy)
'FindWFP = WindowUnderCursor(0& + h, xy)
'End Function

'Public Function WindowUnderCursor(ByVal hWndParent&, Optional xy) 'x,y - координаты внутри hWndParent
'Static cals&
'xMain.DebugPrint 0, "ENTER WindowUnderCursor cals=" & cals
'cals = cals + 1
Dim hwp&: hwp = Cast(ph, 3)
Dim h&, h0&, pt As POINTAPI
GetCursorPos pt
ScreenToClient hwp, pt
If Not IsMissing(xy) Then xy = MakeDWord(pt.x, pt.Y)
h = ChildWindowFromPoint(hwp, pt.x, pt.Y)
If h <> 0 And h <> hwp Then
    h0 = FindWFP(h)
    If h0 <> 0 Then h = h0
End If
FindWFP = h
'cals = cals - 1
'xMain.DebugPrint 0, "EXIT WindowUnderCursor cals=" & cals
End Function

Sub xDoEvents(Optional ByVal n&)
Dim t&: t = GetTickCount
'On Error Resume Next

Do
'Debug.Print
'DoEvents
API_DoEvents
'Debug.Print "xDoEvents", n, t, Timer, Timer - t
If GetTickCount - t > n Then Exit Do
Loop
'Loop While GetTickCount - t < n
End Sub

Public Function SendMSG(ByVal hWnd, ByVal iMsg, Optional ByVal wParam = 0, Optional ByVal lParam = 0)
Dim wp&, lp
wp = wParam: lp = lParam
SendMSG = SendMessage(0& + hWnd, iMsg + 0&, wp, lp)
wParam = wp: lParam = lp
End Function
Public Function SendMSGstring(ByVal hWnd, ByVal iMsg, Optional ByVal wParam = 0, Optional ByVal lParam = vbNullString)
Dim wp&, lp$
wp = wParam: lp = S_(lParam)
SendMSGstring = SendMessageString(0& + hWnd, iMsg + 0&, wp, lp)
wParam = wp: lParam = lp
End Function

Public Function PostMSG(ByVal hWnd, ByVal iMsg, Optional ByVal wParam = 0, Optional ByVal lParam = 0)
Dim wp&, lp&
wp = wParam: lp = lParam
PostMSG = PostMessage(0& + hWnd, iMsg + 0&, wp, lp)
wParam = wp: lParam = lp
End Function

'Public Function ApiMsgBox(ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long)
'ApiMsgBox = MessageBoxA(ByVal hWnd, ByVal lpText, ByVal lpCaption, ByVal wType)
'End Function

Function FindWindow(ByVal pCaption, Optional ByVal pClassName = vbNullString, Optional ByVal hWndParent = 0, Optional ByVal hwndChildAfter = 0) As Long
Dim phWnd&, chWnd&, sClassName$, sCaption$
sClassName = NzS(pClassName, vbNullString)
sCaption = NzS(pCaption, vbNullString)
phWnd = L_(hWndParent)
chWnd = L_(hwndChildAfter)
FindWindow = apiFindWindowEx(phWnd, chWnd, sClassName, sCaption)
'Debug.Print "FindWindowByName(" & lpWindowName & ")=" & FindWindowByName
End Function


'Sub TransColor(hWnd&, nColor&)
'    SetWindowLong hWnd, GWL_EXSTYLE, WS_EX_LAYERED  'Or WS_EX_NOPARENTNOTIFY
'    SetLayeredWindowAttributes hWnd, nColor, 255, 3&
'    UpdateWindow hWnd
'End Sub


'Public Function OLEColor(c)
'OLEColor = TranslateColor(0 + c)
'End Function

Public Sub RZBind(ByVal ph, ByVal ch)
ReleaseCapture
If rz Is Nothing Then Set rz = New CResizer
rz.Bind Cast(ph, vbLong), Cast(ch, vbLong)
ReleaseCapture
End Sub

Function ToMatrix(ByVal items, ParamArray parFields()) As CMatrix
'fields(N) =name[`~!@#$%^&*+:[ SIZE[:SCALE]$-: }+] [[=]Expression]
If UBound(parFields) = -1 Then Exit Function
Dim hdr, Item, p_Items
Dim i&, n&, index&
Dim ar, fld, row
Dim sitem$

On Error Resume Next
AssignResult p_Items, pToItems(items)
If VarType(p_Items) Then
    Dim sc As ScriptControl
    hdr = parFields: If UBound(hdr) = 0 Then If IsArray(hdr(0)) Then hdr = parFields(0)
    hdr = pToPrepare(sc, hdr)
    n = UBound(hdr)
    If n > -1 Then
        ReDim ar(n, -1 To -1) 'ar = NewArray("8204,0,0," & n & ",-1,1")
        For i = 0 To n
            fld = StrFieldAppend(hdr(i))  'ARRAY(Name, Type, Size, Attributes, Precision, NumericScale, ExpressionDefaultValue)
            If Len(fld(6)) Then hdr(i) = fld(6) Else hdr(i) = fld(0)
            ReDim Preserve fld(5)
            ar(i, -1) = fld
        Next
        For Each Item In p_Items
            CallByName sc.CodeObject, "Item", IIf(IsObject(Item), VbSet, VbLet), Item
            If IsObject(Item) Then sitem = "item." Else sitem = ""
            ReDim row(n)
            
            For i = 0 To n
                row(i) = pCallItem(sc, pToRIndex(hdr(i), index), sitem)
            Next
            index = index + 1
            pArrayLine ar, index, row, 2 'Добавляем в конец
        Next
    End If
End If
Set ToMatrix = CMatrix(3, ar)
End Function

Private Function pToItems(items)
Dim Item
On Error Resume Next
For Each Item In items: Exit For: Next 'Try to get collection item
If Err Then 'Not a collection
    If IsObject(items) Then pToItems = Array(items) 'object ITEM
Else
    If IsObject(items) Then Set pToItems = items Else pToItems = items
End If
End Function

Private Function pToPrepare(sc As ScriptControl, items, Optional ByVal sDef = "")
Dim Item, v, ar: ar = Array()
Set sc = VBScript("*", "Dim Item,items,pitems")
For Each Item In items
    If TypeName(Item) = "CDataObject" Then ' xVBS.AddObject Item.Name, Item.Value
        'If IsObject(Item.Value) Then sc.addObject Item.Name, Item.Value Else sc.AddCode vbRemoveComents(Item.Value)
        If IsObject(Item.Value) Then sc.addObject Item.Name, Item.Value Else sc.AddCode PrepareCode(Item.Value, sc.Language)
    Else
        v = NzS(Item, sDef)
        If Len(v) Then ReDim Preserve ar(UBound(ar) + 1): ar(UBound(ar)) = v
    End If
Next
pToPrepare = ar
End Function

Private Function pToRIndex(s, i)
pToRIndex = Replace(Replace(s, "#i#", i), "#index#", i)
End Function

Function ToArrayProp(ByVal items, ParamArray propList())
ToArrayProp = ToArray(items, propList)
depricate "ToArrayProp", "ToArray"
End Function

Function ToArray(ByVal items, ParamArray propList())
'sar = For EACH     Array1Dim,Collection
'return array of propName's
Dim ar, p_Items
ar = Array()
On Error Resume Next
AssignResult p_Items, pToItems(items)
If VarType(p_Items) Then
    Dim sc As ScriptControl
    Dim Item, prop, props, index&, sitem
    If UBound(propList) = -1 Then prop = Array(Null) Else prop = propList
    If UBound(prop) = 0 Then If IsArray(prop(0)) Then prop = propList(0)
    props = pToPrepare(sc, prop, "Name")
    CallByName sc.CodeObject, "pItems", IIf(IsObject(p_Items), VbSet, VbLet), p_Items
    For Each Item In p_Items
        CallByName sc.CodeObject, "Item", IIf(IsObject(Item), VbSet, VbLet), Item
        If IsObject(Item) Then sitem = "item." Else sitem = ""
        For Each prop In props
            GoSub additem
'            ReDim Preserve ar(UBound(ar) + 1)
'            sc.CodeObject.items = ar
'            ar(UBound(ar)) = pCallItem(sc, pToRIndex(prop, index), sitem)
        Next
        index = index + 1
    Next
    sc.Reset
End If
ToArray = ar
Exit Function

additem:
    sc.CodeObject.items = ar
    Dim v: v = pCallItem(sc, pToRIndex(prop, index), sitem)
    If VarType(v) <> vbError Then
        ReDim Preserve ar(UBound(ar) + 1)
        ar(UBound(ar)) = v
    Else
        'index = index - 1
    '    Debug.Print index, prop; "=VTError"
    End If
    Return
End Function

Function VTError()
VarVT(VTError) = vbError
End Function

Private Function pCallItem(xVBS As ScriptControl, ByVal Expression, ByVal sitem)  'Eval String Pah with ITEM
If Len(Expression) Then
    On Error Resume Next
    If Left$(Expression, 1) = "=" Then
        pCallItem = xVBS.Eval(Mid$(Expression, 2))
    Else
        pCallItem = xVBS.Eval(sitem & Expression)
    End If
End If
End Function


'Function SplitV(src, ByVal splitter) 'As Variant()
'Dim a, ar, v, vt
'ar = Array()
'For Each a In Split(src, splitter)
'    ReDim Preserve ar(UBound(ar) + 1)
'    If IsNumeric(a) Then
'        For Each vt In Array(vbInteger, vbLong, vbDouble)
'        v = modMain.Cast(a, vt)
'        If a = "" & v Then Exit For
'        Next
'        ar(UBound(ar)) = v
'    Else
'        ar(UBound(ar)) = a
'    End If
'
'Next
'SplitV = ar
'End Function


'Sub ListHords()
'Debug.Print "ListHords = {" & hrds.Names & "}"
'End Sub

Public Property Get Buffer(ByVal index&) As CParam
Select Case index
Case 0: Set Buffer = m_Buffer '.CParam
Case 1: Set Buffer = GlobalGDI '.CParam
Case 2: Set Buffer = GlobalGDICount '.CParam
Case 3: Set Buffer = cpRegEx '.CParam
Case 4: Set Buffer = EnumSubClass '.CParam
Case 5: Set Buffer = hrds '.CParam

End Select
'Debug.Print "GET Buffer=" & Buffer.ParamCount
End Property

'========== CLIPBOARD ============
Public Function ClipboardClear()
'If OpenClipboard(0&) Then
'    EmptyClipboard
'    CloseClipboard
'End If
Clipboard.Clear
End Function
Public Property Let ClipboardText(ByVal v)
Clipboard.Clear
'Dim cl&
'cl = GetKeyboardLayout(0)
'If cl <> 68748313 Then ActivateKeyboardLayout 68748313, 0
Clipboard.SetText Join2(v)
'If cl <> 68748313 Then ActivateKeyboardLayout cl, 0
End Property
Public Property Get ClipboardText()
ClipboardText = Clipboard.GetText
End Property

Function ClipBoardGetData(ByVal sFormat, Optional ByVal bCompressed As Boolean = True)
Dim nFormat As Long
Dim buf() As Byte, nSize&
Dim lCB&, lPtr&
Dim v$
nFormat = ClipBoardFormat(sFormat)

If nFormat < &HC000 Or nFormat = 0 Then Exit Function

If OpenClipboard(0&) Then
    lCB = GetClipboardData(nFormat)
    If lCB Then
        nSize = GlobalSize(lCB)
        If nSize > 0& Then lPtr = GlobalLock(lCB)
        If lPtr Then
            ReDim buf(0 To nSize - 1&)
            CopyMemory buf(0), ByVal lPtr, nSize
            GlobalUnlock lCB
            
            v = StrConv(buf, vbUnicode)
            If bCompressed Then
                If gIsCompressed(v) Then v = gDecompress(v)
                ClipBoardGetData = modMain.ToVariant(v)
            Else
                ClipBoardGetData = v
            End If
            
        End If
    End If
    CloseClipboard
End If
End Function

Function ClipBoardSetData(dVal, ByVal sFormat, Optional ByVal nCompress As Long = -1) As Long
Dim hMem As Long, lPtr As Long
Dim nSize&, buf() As Byte
Dim nFormat As Long
If VarType(dVal) < 2 Or Len(S_(sFormat)) = 0 Then Exit Function
If OpenClipboard(0&) Then
    If VarType(sFormat) = vbString Then nFormat = RegisterFormat("" & sFormat) Else nFormat = sFormat
    If nFormat Then
        If VarType(dVal) = vbLong Then
            ReDim buf(3)
            CopyMemory buf(0), dVal, 4
        ElseIf nCompress = -1 Then
            buf = StrConv(gCompress(modMain.ToString(dVal)), vbFromUnicode)
        ElseIf nCompress = 0 Then
            buf = StrConv(S_(dVal), vbFromUnicode)
        End If
        nSize = 1 + UBound(buf)
        hMem = GlobalAlloc(&H2000, nSize)
        If hMem Then
            lPtr = GlobalLock(hMem)
            If lPtr Then CopyMemory ByVal lPtr, buf(0), nSize
            GlobalUnlock hMem
            ClipBoardSetData = Not (SetClipboardData(nFormat, hMem) = 0&)
        End If
    End If
    CloseClipboard
End If
End Function

Function ClipFormatName(ByVal nFormat)
Dim s$: s = Space$(256&): ClipFormatName = Left$(s, GetClipboardFormatName(Cast(nFormat, 3), s, 255&))
End Function
Function RegisterFormat(ByVal sFormat) As Integer
sFormat = S_(sFormat): If Not IsNum(sFormat) Then RegisterFormat = LOWORD(RegisterClipboardFormat(sFormat))
End Function

Function ClipBoardFormat(ByVal sFormat) As Long 'return if present ID
Dim n&, nFormat As Long
sFormat = S_(sFormat): If IsNum(sFormat) Then nFormat = sFormat
If OpenClipboard(0&) Then
    Do
        n = EnumClipboardFormats(n): If n = 0 Then Exit Do
        If nFormat = n Then 'Поиск по номеру
            ClipBoardFormat = n: Exit Do
        Else 'Поиск по имени
            If sFormat = ClipFormatName(n) Then ClipBoardFormat = n: Exit Do
        End If
    Loop
    CloseClipboard
End If
End Function
'========== CLIPBOARD ============

'Function TextZoom(ByVal TexthWnd)
'Dim cp As New CParam
''cp.Value = EXEParam("TxZm", "LIB.xc")
'cp.Value = CParam("+TxZm")
'Dim rc As RECT, drc As RECT
'GetWindowRect GetDesktopWindow(), drc
'GetWindowRect TexthWnd, rc
'If rc.Bottom + 150 > drc.Bottom - 30 Then rc.Bottom = rc.Top - 150
'If rc.Left + 300 > drc.Right Then rc.Left = drc.Right - 300
'
'cp.AddParams Array("DropX", rc.Left, "DropY", rc.Bottom, "Modal", TexthWnd, "ActiveControl", GetFocus())
'FormValue cp.Source
'End Function

Function InBox(Optional ByVal Prompt = "<имя параметра>", Optional ByVal vDefault = vbNullString, Optional theParams)
Dim cp As New CParam
cp.Load "+IBOX"
'Debug.Print cp.json("*cx")
cp.AddParams Array("cx\Help", Prompt, "cx\Tag", vDefault, "cx\DefaultValue", "cx.Tag")
'Debug.Print CParam(theParams).json("*cx")
cp.AddParams theParams, , True
'Debug.Print cp.json("*cx")
InBox = DoModal(cp)
End Function

Function MsBox(ByVal Prompt, Optional ByVal nStyle = 0, Optional ByVal sTitle = vbNullString, Optional ByVal vDefault = Null, Optional theParams = Null)
Dim cp As New CParam
cp.Load "+MBOX"
cp.AddParams Array("tx\Tag", Prompt, "Style", nStyle, "Caption", sTitle, "Tag", vDefault)
cp.AddParams theParams
MsBox = Nz(DoModal(cp), vDefault)
End Function

Public Function DoModal(ByVal OpenArgs, Optional ByVal Owner)
Dim cp As New CParam
'If TypeName(OpenArgs) = "CParam" Then cp.Value = OpenArgs.Source Else
cp.Value = OpenArgs
Dim af&
af = cp.ParamDef("ActiveControl", GetFocus())
cp("ActiveControl") = af
Dim p As xControl, ap&
ap = af
Do While (ap <> 0) And (p Is Nothing)
    Set p = hxControl(ap)
    If p Is Nothing Then ap = GetParent(ap)
Loop 'Ищем xControl если фркус был на чужом  окне
If Not p Is Nothing Then cp("Modal") = p.PophWnd
cp("ModalMode") = 1
Dim cd As New CDataObject
If TypeName(Owner) = xs.sxControl Then Set p = Owner Else Set p = xMainWnd
Dim hWnd&: hWnd = NewForm(vbNullString, p, cp.Source, cd)
If hWnd Then
    MessageLoopWhileVisible hWnd
    DoModal = cd.Value
End If
apiSetFocus af
End Function


Public Function FormValue(ByVal OpenArgs, Optional ByVal Owner)
Dim hWnd&, f&
Dim p As xControl
Dim cd As New CDataObject
f = GetFocus()
If TypeName(Owner) = xs.sxControl Then Set p = Owner Else Set p = xMainWnd
hWnd = NewForm(vbNullString, p, OpenArgs, cd)
If hWnd Then MessageLoopWhileVisible hWnd
DoEvents
FormValue = cd.Value
'PostMessage hWnd, WM_CLOSE, 0, 0
apiSetFocus f
API_DoEvents
End Function

Public Function DropValueEx(xOwner, xDrop, v, Optional xPos = 1000, Optional yPos = 1000, Optional hPos = 1000)
If TypeName(xDrop) <> xs.sxControl Then Exit Function
If TypeName(xOwner) <> xs.sxControl Then Exit Function
If Not xDrop.isDrop Then Exit Function
Dim pOwner As xControl, prevDrop As xControl
Set pOwner = xDrop.Owner 'Old Owner
If pOwner.DropDroped Then pOwner.DropHide ': Exit Function

Set prevDrop = xOwner.Drop 'Old Owner drop
xOwner.DropSet xDrop 'Init New Drop
Set xDrop.Owner = xOwner 'Set New Owner
xOwner.Activate
xDrop.DropShow v, L_(xPos), L_(yPos), L_(hPos)
MessageLoopWhileVisible xDrop.hWnd
Set xDrop.Owner = pOwner
xOwner.DropSet prevDrop
If xDrop.HasChanges Then DropValueEx = xDrop.Value
End Function

Public Function DropValue(ByVal pOwner, ByVal OpenArgs, v)
Dim Drop As xControl, h&, prevDrop As xControl
If pOwner.DropDroped Then pOwner.DropHide ': Exit Function
If TypeName(pOwner) <> xs.sxControl Then Exit Function
Set prevDrop = pOwner.Drop 'SAVE OLD DROP
'If Not prevDrop Is Nothing Then Debug.Print prevDrop.Name
Dim cd As New CDataObject
Set Drop = xControl(NewForm(vbNullString, , OpenArgs, cd))
If Drop.hWnd = 0 Then Exit Function
Set Drop.Owner = pOwner 'SET NEW DROP OWNER =OWNER
pOwner.DropSet Drop 'SET OWNER NEW DROP
pOwner.Activate
Drop.DropShow v
MessageLoopWhileVisible Drop.hWnd

DropValue = cd.Value

pOwner.DropSet prevDrop 'RESTORE OWNER OLD DROP
Set Drop.Owner = xMainWnd 'RESTORE DROP OWNER
xMainWnd.DropSet Nothing 'RESET NEWDROP
Drop.CloseForm 'CLOSE NEWDROP
Set Drop = Nothing
End Function



Public Function NewForm(ByVal FileName$, Optional ByVal Owner, Optional ByVal OpenArgs, Optional dataobject) As Long
Dim cp As CParam, ar, cur&, oa$

Set cp = New CParam
'If Len(FileName) Then oa = Nz(cp.Load(FileName)) 'Имя файла
If Len(FileName) Then cp.Load FileName

'if cp.ParamCount And Len(oa) > 0 Then cp.Param("SourcePath") = oa
cur = GetCursor()
SetCursor hCursor_WAIT

'Dim opa As New CParam
If IsArray(OpenArgs) Then
    cp.AddParams OpenArgs, , 1
ElseIf TypeName(OpenArgs) = "CParam" Then
    If cp.ParamCount Then cp.AddParams OpenArgs, , 1 Else Set cp = OpenArgs
End If
'cp.AddParams OpenArgs, , 1

Dim Prepare$, ret As Boolean
Prepare = cp.ParamDef("Prepare", vbNullString) ', 1)

Dim p As xControl: If TypeName(Owner) = xs.sxControl Then Set p = Owner Else Set p = xMainWnd
If p.hWnd Then cp("hOwner") = p.hWnd


If Len(Prepare) Then
    On Error Resume Next
    Dim sc As ScriptControl
    'Set sc = VBScript("*+", Empty, "", Prepare, "Result", dataobject)
    Set sc = VBScript("Result", dataobject, "*+", Prepare)
    sc.AllowUI = 1
    ret = L_(sc.CodeObject.Prepare(cp))
    sc.Reset
    Set sc = Nothing
    Err.Clear
End If


If Not ret Then
'GRANT = Разрешение
'DENY = Запрет
'PERMIT = Допуск

Dim x As New xControl
'DBG "BEFORE CREATE============= StartPos=" & cp.Param("StartPos") & " Width=" & cp.Param("Width") & " Height=" & cp.Param("Height")
'If TypeName(dataobject) = "CDataObject" Then

Set x.FormResult = dataobject
NewForm = x.pCreate(p, cp, Null) ', ap.Source)


'DebugPrint 0, "NewForm cp.url=" & x.CPSource.Url

Set x = Nothing


End If

'Set p = Nothing
'Set cp = Nothing
SetCursor cur
ChekWaitSocket

End Function

Sub ChekWaitSocket() 'Проверим наличие слухательного сокета и сделаем его если его нет
If L_(m_Param(App.ProductName & "\NoWaitSocket")) = 0 And xMainWnd.Controls.count > 0 Then
    Dim p&: p = 999
    Dim xc As xControl
    Set xc = xMainWnd.Controls(1)
    If Not xc Is Nothing Then
        If xc.SocketPort = 0 Then
            While Not xc.WaitData(p)
            p = p + 1
            Wend
'            Dim fn&, fns$
'            fn = FreeFile(): fns = App.Path & "\z.xvb." & xMain.INIBase & "." & Format(Now, "ddmmyy.hhnnss") & "." & GetCurrentProcessId & "." & Environ("UserName") & "." & xc.SocketPort
'            Open fns For Output Lock Read Write As fn
'            If fn Then Close fn: Kill fns
'
        End If
        Set xc = Nothing
    End If
End If
End Sub

'===========STRINGS ===================
'Public Function CastVar(ByVal vValue, ByVal vt, Optional ByVal bAcceptNull As Boolean)
'On Error Resume Next
'CastVar = modMain.Cast(vValue, vt, bAcceptNull)
'Err.Clear
'End Function
'Function aVal(ByVal vExpression)
'aVal = modMain.aVal(vExpression)
'End Function


Function aRound(ByVal vNumber, ByVal nNumDigits) ' 5mks
Dim d&
On Error Resume Next
d = Val(Nz(nNumDigits, 0))
'aRound = (Round(aVal(vNumber) * (10 ^ d))) / (10 ^ d)
aRound = (Fix(aVal(vNumber) * (10 ^ d))) / (10 ^ d)
End Function

Function nRound(ByVal vNumer, ByVal nNominator)
nNominator = IF0(L_(nNominator), 1)
nRound = Round(aVal(vNumer) / nNominator) * nNominator
End Function

Function SNumer(ByVal vNumer)
SNumer = Replace(aVal(vNumer), ",", ".")
End Function

Public Function BETag(ByVal v, Optional ByVal bStr, Optional ByVal eStr, Optional nEscapeOptions&, Optional bParseIndex As Boolean)
Dim cDims%, ar, a, i&, s
cDims = ArrayDims(v)
If cDims > 0 Then ar = v Else ar = Array(v)
bStr = S_(bStr): eStr = S_(eStr)
For i = LBound(ar, IIf(cDims = 0, 1, cDims)) To UBound(ar, IIf(cDims = 0, 1, cDims))
    If cDims < 2 Then
        a = ar(i)
        If VarType(a) And vbArray Then
            a = BETag(a, bStr, eStr, nEscapeOptions, bParseIndex)
        ElseIf (VarType(a) = vbString) And nEscapeOptions Then
            a = bStr & modMain.JSONEscape(a) & eStr
        ElseIf VarType(a) > vbNull Then
            a = bStr & a & eStr
        Else
            a = bStr & eStr
        End If
        If bParseIndex Then a = Replace(a, "#index#", i)
        ar(i) = a
    Else
        pArrayLine ar, i, BETag(pArrayLine(ar, i), bStr, eStr, nEscapeOptions, bParseIndex), 1
    End If
Next
If cDims Then BETag = ar Else BETag = ar(0)

End Function

Public Function Quot(sar, Optional ByVal ch = 34, Optional ByVal bQuotAll As Boolean)
'ch>=0 qut string
'ch<0 UN quot string
'bQuotAll =  false = quot only strings
'bQuotAll =  true = quot all values byte/long...
Dim i&, j&, ar, char$
ch = Nz(ch)
If Len(ch) = 0 Then Quot = sar: Exit Function
With mSafeArray(sar)
If .cDims = 0 And VarType(sar) > 1 Then
    ar = S_(sar)
    'If (VarType(ar) = vbString) Or bQuotAll Then
    If VarType(ch) = vbString Then
        char = ch
        ch = 1
    Else
        char = Chr$(modMain.Cast(Abs(ch), vbByte))
    End If
    
    If ch > -1 And ch < 256 Then 'Quot
        If ch > 0 Then ar = char & Replace$(ar, char, char & char) & char
    Else 'UnQuot
        If Len(ar) > 1 Or bQuotAll Then
            'If Left$(ar, Len(char)) = char And Right$(ar, Len(char)) = char Then ar = Replace$(Mid$(ar, Len(char) + 1, Len(ar) - Len(char) - 1), char & char, char)
            If Left$(ar, Len(char)) = char And Right$(ar, Len(char)) = char Then ar = Mid(ar, 2, Len(ar) - 2 * Len(char))
            ar = Replace$(ar, char & char, char)
        End If
    End If
    Quot = ar
    Exit Function
End If

For i = .Bounds(1).lLbound To .Bounds(1).lLbound + .Bounds(1).cElements - 1
    If .cDims = 1 Then
        sar(i) = Quot(sar(i), ch, bQuotAll)
    ElseIf .cDims = 2 Then
        For j = .Bounds(2).lLbound To .Bounds(2).lLbound + .Bounds(2).cElements - 1
            sar(j, i) = Quot(sar(j, i), ch, bQuotAll)
        Next
    Else
        pArrayLine sar, i, Quot(pArrayLine(sar, i), ch, bQuotAll), 2
    End If
Next
Quot = sar
End With

'Quot = ar
End Function

'Public Function NzAfter(v, Optional sSuffix = vbNullString)
'Dim vt&: vt = VarType(v)
'If vt < 2 Or vt = 10 Or (vt And vbArray) > 0 Then NzAfter = vbNullString Else NzAfter = v & vbNullString & sSuffix
'End Function

'Public Function NzSAfter(v, Optional sSuffix = vbNullString)
''Dim vt&: vt = VarType(v)
''NzSAfter = vbNullString
''If vt < 2 Or vt = 10 Or (vt And vbArray) > 0 Then Exit Function
''If Len(vbNullString & v) Then NzSAfter = v & vbNullString & sSuffix
'NzSAfter = NzS(v, , , sSuffix)
'End Function

'Public Function NzPre(v, Optional sPrefix = vbNullString)
'Dim vt&: vt = VarType(v)
'If vt < 2 Or vt = 10 Or (vt And vbArray) > 0 Then NzPre = vbNullString Else NzPre = sPrefix & vbNullString & v
'End Function
'Public Function NzSPre(v, Optional sPrefix = vbNullString)
''Dim vt&: vt = VarType(v)
''NzSPre = vbNullString
''If vt < 2 Or vt = 10 Or (vt And vbArray) > 0 Then Exit Function
''If Len(vbNullString & v) Then NzSPre = sPrefix & vbNullString & v
'NzSPre = NzS(v, , sPrefix)
'End Function

Function LeftStr(str, find_str, Optional ByVal reverse = 0, Optional ByVal start = 0)
'reverse = 0   С начала строки [str]+start  до первого вхождения  [find_str]
'reverse <> 0   С начала строки [str]+start до первого вхождения  [find_str] начиная с конца
'если [find_str]  не надено то вся строка
On Error Resume Next
Dim i&, s&: s = start
If s < 1 Then s = 1
If Val(reverse) Then i = InStrRev(str, find_str, , vbTextCompare) Else i = InStr(s, str, find_str, vbTextCompare)
If i < 1 Then i = 1
LeftStr = Mid(Left(str, i - 1), s)
End Function
Function RightStr(str, find_str, Optional ByVal reverse = 0, Optional ByVal nLeft = 0)
'reverse = 0   С  первого вхождения  [find_str] начиная с конца до конца строки + обрезать nleft
'reverse <> 0   С  первого вхождения  [find_str] начиная с начала  до конца строки + обрезать nleft
'если [find_str]  не надено то вся строка
On Error Resume Next
Dim i&, n&: n = nLeft
If n < 1 Then n = Len(str)
If L_(reverse) Then i = InStr(1, str, find_str, vbTextCompare) Else i = InStrRev(str, find_str, , vbTextCompare)
If i Then i = 1 + Len(str) - Len(find_str) - i
RightStr = Left(Right(str, i), n)
End Function



Public Function EnvironVB(sExpression) As String
On Error Resume Next: EnvironVB = Environ(sExpression)
End Function
'Function ReplaceEnviron(ByVal str, Optional ByVal quotstr = "%") As String
'On Error Resume Next: ReplaceEnviron = pReplaceEnviron(str, quotstr)
'End Function

Public Function ShellVB(ByVal sPathName, Optional ByVal nWindowStyle) As Double
On Error Resume Next: ShellVB = Shell(S_(sPathName), L_(nWindowStyle))
End Function

'Public Function ShellEx(sPathName, Optional nShowCmd, Optional sOperation, Optional sParameters, Optional sDirectory) As Double
'On Error Resume Next
'ShellEx = apiShellExecute(0, Nz(sOperation), Nz(sPathName), Nz(sParameters), Nz(sDirectory), aVal(nShowCmd))
'
'End Function
Public Function ShellEx(ByVal sPathName, Optional ByVal nShowCmd = 5, Optional ByVal sOperation, Optional ByVal sParameters, Optional ByVal sDirectory, Optional ByVal bAsync = 1) As Long
On Error Resume Next
Dim sh As SHELLEXECUTEINFO
sh.cbSize = LenB(sh)
sh.fMask = &H40& 'SEE_MASK_NOCLOSEPROCESS (0x00000040)
'sh.hwnd = 0
sh.lpVerb = sOperation
sh.lpFile = sPathName
sh.lpParameters = sParameters
sh.lpDirectory = sDirectory
sh.nShow = nShowCmd
'sh.hInstApp = 0
ShellEx = ShellExecuteEx(sh)
If L_(bAsync) = 0 Then WaitForSingleObject sh.hProcess, &HFFFF
End Function
'SW_HIDE (0) Hides the window and activates another window.
'SW_MAXIMIZE (3) Maximizes the specified window.
'SW_MINIMIZE (6) Minimizes the specified window and activates the next top-level window in the z-order.
'SW_RESTORE (9) Activates and displays the window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when restoring a minimized window.
'SW_SHOW (5) Activates the window and displays it in its current size and position.
'SW_SHOWDEFAULT (10) Sets the show state based on the SW_ flag specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application. An application should call ShowWindow with this flag to set the initial show state of its main window.
'SW_SHOWMAXIMIZED (3) Activates the window and displays it as a maximized window.
'SW_SHOWMINIMIZED (2) Activates the window and displays it as a minimized window.
'SW_SHOWMINNOACTIVE (7) Displays the window as a minimized window. The active window remains active.
'SW_SHOWNA (8) Displays the window in its current state. The active window remains active.
'SW_SHOWNOACTIVATE (4) Displays a window in its most recent size and position. The active window remains active.
'SW_SHOWNORMAL (1) Activates and displays a window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when displaying the window for the first time.

Public Function ChooseVB(ByVal index, ParamArray arChoice())
On Error Resume Next
Dim ar
If IsArray(arChoice(0)) Then ar = arChoice(0) Else ar = arChoice
ChooseVB = ar(L_(index))
End Function
Public Function SwitchVB(ParamArray arCaseValue())
'arVarsExpr = Array(case1,value1,[case2,value2,...],case else)
On Error Resume Next
Dim i&, ar
ar = arCaseValue
If UBound(ar) = 0 Then If IsArray(ar(0)) Then ar = arCaseValue(0)
For i = 0 To UBound(ar) Step 2
If i = UBound(ar) Then 'Switch ELSE
    SwitchVB = ar(i)
Else 'Switch Case
    If ar(i) Then SwitchVB = ar(i + 1): Exit For
End If
Next
End Function

Public Function SelectCase(ByVal expr, ParamArray arCaseValue())
'arVarsExpr = Array(case1,value1,[case2,value2,...],case else)
On Error Resume Next
Dim i&, ar
ar = arCaseValue
If UBound(ar) = 0 Then If IsArray(ar(0)) Then ar = arCaseValue(0)
For i = 0 To UBound(ar) Step 2
If i = UBound(ar) Then 'Case ELSE
    SelectCase = ar(i)
Else 'Select Case
    If gFindIndex(ar(i), expr) > -1 Then SelectCase = ar(i + 1): Exit For
End If
Next
End Function

Public Function VBDir(ByVal pth) As String
On Error Resume Next
VBDir = Dir(S_(pth))
End Function

'Public Function FormatVB(ByVal Expression As Variant, ByVal sFormat As String, Optional ByVal nFirstDayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek, Optional ByVal nFirstWeekOfYear As VbFirstWeekOfYear = vbUseSystem, Optional ByVal TargetLCID As Long = 1049)
'On Error Resume Next
'Static hRes As Long, b(0 To 4095) As Byte
'If TargetLCID <> 1049 Then
''1049=RU
''1033=EN
'    Dim psFormat$, res$
'    psFormat = sFormat
'     hRes = VarTokenizeFormatString(StrPtr(psFormat), b(0), UBound(b) + 1, nFirstDayOfWeek, nFirstWeekOfYear, 1049)
'    ' If hRes Then Err.Raise hRes
'     hRes = VarFormatFromTokens(Expression, StrPtr(psFormat), b(0), 0, ByVal VarPtr(res), TargetLCID)
'     'If hRes Then Err.Raise hRes
'     FormatVB = res
' Else
'    FormatVB = Format$(Expression, sFormat, nFirstDayOfWeek, nFirstWeekOfYear)
' End If
'End Function

Public Function FormatVB(ByVal Expression As Variant, ByVal sFormat As String, Optional ByVal TargetLCID As Long = 1049)
On Error Resume Next
Static hres As Long, b(0 To 4095) As Byte
If TargetLCID <> 1049 Then
'1049=RU
'1033=EN
    Dim psFormat$, res$
    psFormat = sFormat
     hres = VarTokenizeFormatString(StrPtr(psFormat), b(0), UBound(b) + 1, 0, 0, 1049)
    ' If hRes Then Err.Raise hRes
     hres = VarFormatFromTokens(Expression, StrPtr(psFormat), b(0), 0, ByVal VarPtr(res), TargetLCID)
     'If hRes Then Err.Raise hRes
     FormatVB = res
 Else
    FormatVB = Format$(Expression, sFormat) ', nFirstDayOfWeek, nFirstWeekOfYear)
 End If
End Function


Public Function Replicate$(nCount, sString)
Replicate = modMain.Replicate(0 + nCount, S_(sString))
End Function

'=================================================
'=================================================
'src="select {field_list} from dc where dcid={dcid} order by {orderby}"
'cpar=(CPARAM) Array("field_list","*",dcid,"1","orderby","dcnm")
'ReplaceArgs(src,cpar)="select * from dc where dcid=1 order by dcnm"
'ReplaceArgs(src, cpar as CParam)
'ReplaceArgs(src, cpar as Array (CParam))
Function ReplaceArgs(ByVal src, var, Optional ByVal stg = "{", Optional ByVal etg = "}", Optional ByVal bClear = 1)
'bClear  = полная зачистка
Dim res$, nm$, i&
Dim cp As New CParam
Dim nv$: bClear = L_(bClear)
res = S_(src)
If Val(GTag(res, -1, stg, etg)) > 0 Then
    cp.Value = var
    If cp.ParamCount Then
        Do
            nm = GTag(res, 1, stg, etg) ' Получаем имя аргумента
            If bClear = 0 Then nv = Chr(1) & nm & Chr(2)
            res = Replace(res, stg & nm & etg, Nz(cp(nm), nv)) 'Заменяем аргумент на его значение
            i = i + 1
        Loop While Len(nm) > 0 And i < 1000
    End If
End If
If bClear = 0 Then res = Replace(Replace(res, Chr$(1), stg), Chr$(2), etg)
ReplaceArgs = res
End Function

Function Replace2(ByVal expr, frar, Optional ByVal start As Long = 1, Optional ByVal count As Long = -1, Optional ByVal Compare = 1) As String
Dim i&
expr = "" & expr
For i = 0 To UBound(frar) - (UBound(frar) + 1) Mod 2 Step 2
    If Len(expr) = 0 Then Exit For
    'Debug.Print InStr(expr, frar(i))
    expr = Replace(expr, frar(i), frar(i + 1), start, count, Compare)
Next
Replace2 = expr

End Function
'===========STRINGS ===================

Sub LetControls(propName As String, vLetValue, ParamArray Controls())
Dim c, cs
On Error Resume Next
cs = Controls
If UBound(Controls) = 0 Then
    If IsArray(Controls(0)) Then
        cs = Controls(0)
    ElseIf TypeName(Controls(0)) = "Collection" Then
        Set cs = Controls(0)
    End If
End If
For Each c In cs
If CallByNameEx(c, propName, VbGet) <> vLetValue Then CallByNameEx c, propName, VbLet, vLetValue
Next
Err.Clear
End Sub

Function xLet(obj, propName, vLetValue)
On Error Resume Next
CallByNameEx obj, propName, VbLet, vLetValue
End Function

Public Sub SetVals(src, ParamArray ar())
'LetValues (ar())=src
'ar(n)=src(n) или ar(n)=src(0)(n)
'ar(0)(n)=src(n) или ar(0)(n)=src(0)(n)
Dim i&, tp&, nar&, nsrc&
If UBound(ar) < 0 Then Exit Sub 'NO DESTINATION VARS/CONTROLS
If VarType(ar(0)) And vbArray Then tp = tp Or 2: nar = UBound(ar(0)) Else nar = UBound(ar)
If VarType(src) And vbArray Then tp = tp Or 1: nsrc = UBound(src) Else nsrc = nar
On Error Resume Next
For i = 0 To IIf(nsrc >= nar, nar, nsrc)
Select Case tp
Case 0 'ar(i)=src
    If Not IsObject(ar(i)) Then ar(i) = src Else ar(i).Value = src
Case 1 'ar(i)=src(i)
    If Not IsObject(ar(i)) Then ar(i) = src(i) Else ar(i).Value = src(i)
Case 2 ' ar0(i)=src
    If Not IsObject(ar(0)(i)) Then ar(0)(i) = src Else ar(0)(i).Value = src
Case 3 'ar0(i)=src(i)
    If Not IsObject(ar(0)(i)) Then ar(0)(i) = src(i) Else ar(0)(i).Value = src(i)
End Select
Next
Err.Clear
End Sub


'Function GetValues(ParamArray src()) 'Заменить в массиве объекты на их значения
'Dim res, i&
'On Error Resume Next
'If ArrayDims(src(0)) = 1 Then res = src(0) Else res = src
'For i = 0 To UBound(res)
'    If IsObject(res(i)) Then res(i) = res(i).Value
'    If VarType(res(i)) = 1 Then res(i) = Empty
'Next
''RemoveRefs = res
'GetValues = res
'Err.Clear
'End Function


'Property Let LetValues(ParamArray target(), thevalue) 'Установка значений target[(0)]()[.Value] = theValue или значеними элементов имассива theValue()
'Dim i&, tar, n&, btAr As Boolean, bvar As Boolean
'If UBound(target) = -1 Then Exit Property
'If ArrayDims(target(0)) = 1 Then n = UBound(target(0)): btAr = True Else n = UBound(target)
'If ArrayDims(thevalue) = 1 Then bvar = True
'On Error Resume Next
'For i = 0 To n
'If bvar Then
'    If i <= UBound(thevalue) Then
'        If btAr Then
'            If IsObject(target(0)(i)) Then target(i).Value = thevalue(i) Else target(0)(i) = thevalue(i)
'        Else
'            If IsObject(target(i)) Then target(i).Value = thevalue(i) Else target(i) = thevalue(i)
'        End If
'    End If
'Else
'    If btAr Then
'        If IsObject(target(0)(i)) Then target(0)(i).Value = thevalue Else target(0)(i) = thevalue
'    Else
'        If IsObject(target(i)) Then target(i).Value = thevalue Else target(i) = thevalue
'    End If
'End If
'Next
'Err.Clear
'End Property

'=================VBScript=================
Function GlobalVar(ByVal Name, Optional Value)  ', Optional ByVal bReplace = 0)
On Error Resume Next
If Name <> RegExMatch(Name, "^[A-Z]+[A-Z0-9_]*$") Then Exit Function
If Len(Name) = 0 Then Exit Function
With xMainWnd.GlobalModule
    AssignResult GlobalVar, .Eval(Name)
    If VarType(Value) And Not IsMissing(Value) Then
        If VarType(GlobalVar) = 0 Then .ExecuteStatement IIf(IsObject(Value), "Set " & Name & "=Nothing", Name & "=Empty")
        AssignResult GlobalVar, Value
        CallByName .CodeObject, Name, IIf(IsObject(Value), VbSet, VbLet), GlobalVar
    End If
End With
End Function

'Function GlobalLoad(ByVal mpth)
'On Error Resume Next
'VBAddCode xMainWnd.GlobalModule, mpth
'End Function
Function JScript(ParamArray items()) As ScriptControl
Dim ar: If UBound(items) = 0 Then ar = items(0) Else ar = items
Set JScript = Script("jscript", ar)
End Function
Function VBScript(ParamArray items()) As ScriptControl
Dim ar: If UBound(items) = 0 Then ar = items(0) Else ar = items
Set VBScript = Script("vbscript", ar)
End Function

Function Script(ByVal lang, ParamArray items()) As ScriptControl
'End Function
''items = (name,value,,)
'Function VBScript(ParamArray items()) As ScriptControl
Dim sc As New ScriptControl
'sc.Language = "VBSCRIPT"
On Error Resume Next
'lang = S_(lang)

sc.Language = S_(lang)

sc.timeOut = -1: sc.AllowUI = 0
Dim i&, nm$, ar, bAddMembers As Boolean, v
Dim ob As Object, md_nm
'sc.addObject script_engine_ & sc.Language, sc
If UBound(items) > -1 Then
    'If ArrayDims(items(0)) = 1 Then ar = items(0) Else ar = items
    If UBound(items) = 0 Then ar = items(0) Else ar = items
    If (UBound(ar) Mod 2) = 0 Then sc.timeOut = IF0(1000 * Abs(L_(ar(UBound(ar)))), -1) 'last ubound element =set timeout
    If ArrayDims(ar) = 1 Then
    For i = 0 To UBound(ar) - ((1 + UBound(ar)) Mod 2) Step 2
        nm = S_(ar(i))
        bAddMembers = True: If InStr(nm, "-") Then nm = Replace(nm, "-", ""): bAddMembers = False
        If InStr(nm, "*") Then nm = Replace(nm, "*", ""): sc.addObject "xMain", xMain, True: sc.addObject "App", App, True
        If InStr(nm, "+") Then nm = Replace(nm, "+", ""): sc.addObject "xGlobal", MainWnd.GlobalModule.CodeObject, True
        If InStr(nm, "?") Then nm = Replace(nm, "?", ""): md_nm = ""
        Set ob = Nothing
        If IsObject(ar(i + 1)) Then Set ob = ar(i + 1) Else If VarType(ar(i + 1)) = vbLong Then GetMemObj 0& + ar(i + 1), ob
        If Not ob Is Nothing Then   'AddObject
            If InStr(nm, "%") Then nm = Replace(nm, "%", ""): sc.SitehWnd = ob.hWnd
            sc.addObject nm, ob, bAddMembers
        ElseIf Len(nm) Then 'Add variable and init
            sc.AddCode "Dim " & nm: v = ar(i + 1): CallByName sc.CodeObject, nm, VbLet, v
        ElseIf Len(S_(ar(i + 1))) Then 'AddCode
            SCAddCode sc, "" & ar(i + 1), md_nm
        End If
    Next
    End If
End If
Set Script = sc
End Function

'Friend Property Get mVBScript(ParamArray arAddObject()) As ScriptControl 'Optional vbCode
'Friend Function mVBScript() As ScriptControl 'Optional vbCode
'mSC.Reset
''mSC.addObject "xMain", Me, -1  'Добавляем ссылку на глобальный класс
''mSC.AddObject "prn", Printer, 0  'Добавляем ссылку на глобальный класс
'Set mVBScript = mSC
'End Function

Function EvalAr(hVB As Object, ar)
Dim i&, sc As Object
If hVB Is Nothing Then Set sc = mSC Else Set sc = hVB
On Error Resume Next
EvalAr = ar
Dim ear: If IsArray(ar) Then ear = ar Else ear = Array(ar)
For i = 0 To UBound(ear): ear(i) = sc.Eval(ear(i)): Next
If IsArray(ar) Then EvalAr = ear Else EvalAr = ear(0)
End Function

Function AddComponent(sc As ScriptControl, ByVal lang$, Name$, md) As Boolean
On Error Resume Next
'md=Array(name,openargs,url,code)
Dim i&, v
v = lang: v = sc.CodeObject.Me.Name
DebugPrint 8421376, v & "::#AddComponent " & lang & "." & Name
Dim se As ScriptControl
Set se = sc.Eval(script_engine_ & lang)
If se Is Nothing Then Set se = Script(lang, "*", Empty): sc.addObject script_engine_ & lang, se
v = "": lang = ""
For i = 0 To UBound(md) Step 4
    v = NzS(v, , , vbCrLf) & md(i + 3) & NzS(JSONUnEscape(md(i + 1)), , vbCrLf)
    lang = NzS(lang, , , ",") & md(i + 2)
Next
v = SCAddCode(se, v, Name)
If IsArray(v) Then
    DebugPrint -240, " = " & lang
Else
    sc.addObject Name, se.Modules(Name).CodeObject
    DebugPrint -8421376, " = " & lang
    AddComponent = True
    'sc.CodeObject.Test.testfun
    
End If
End Function

'Public Function VBAddCode(qVBS, ByVal code$, Optional ByVal modName = Empty) ', Optional ByVal bTest As Boolean) ')
'VBAddCode = SCAddCode(qVBS, code, modName)
'End Function
Friend Function SCError(sc As ScriptControl, Optional ByVal dy&)
Dim s$, i&
With sc.Error
If .Number = 0 Then Exit Function
    s = .Text
    i = Val(RegExMatch(s, " f_(\d+)_", -2))
    If i Then
        Dim sx0 As syndef2, sx As syndef2: CopyMemory sx0, ByVal Syn2(sc.Language), Len(sx): Let sx = sx0: FillMemory sx0, Len(sx0), 0
        s = RegExMatch(s, Replace(Replace(sx.pf, "%", i), "@", "([^\0]*?)"), -2)
        i = InStr(.Text, s)
    End If
    s = .Description & vbCrLf & "Line:" & .Line - dy & " Pos:" & .Column - i & NzS(s, , " Text={", "}")
    On Error Resume Next
    If sc.SitehWnd <> 0 Then s = xControl(sc.SitehWnd).Name & " " & s
    s = "ERROR " & .Number & " " & s
    DebugPrint 240, s
    SCError = .Line - dy
End With
End Function

Function SCTestCode(lang$, code$)
'test when edit
SCTestCode = Empty
If Len(lang) = 0 Or Len(code) = 0 Then Exit Function
On Error Resume Next
mSC.Reset: mSC.Language = lang
mSC.addObject "xMain", xMain, 1
'Dim ern&
mSC.AddCode PrepareCode(code, lang, 1, SCTestCode)
If mSC.Error.Number Then SCTestCode = SCError(mSC) ' Else If ern Then SCTestCode = ern
End Function


Public Function SCAddCode(qVBS, ByVal code$, Optional ByVal modName = Empty) ', Optional ByVal bTest As Boolean) ')
'Run code
SCAddCode = Empty
If Len(code) = 0 Then Exit Function
On Error Resume Next

Dim tmp$, v, sc As ScriptControl
'If TypeName(qVBS) = "ScriptControl" Then Set sc = qVBS Else Set sc = mSC: sc.Language = NzS(qVBS, sc.Language)
If TypeName(qVBS) = "ScriptControl" Then Set sc = qVBS Else Exit Function
'mSC.Language = sc.Language

code = PrepareCode(code, sc.Language)
Dim i&, md, m As Match, so As xControl, oa
'#COMPONENT proccessing
For Each m In RegExMatch(code, "^#(\w+) (\w+)=([^\r\n]*)$", -1) 'Add Component
    tmp = m.SubMatches(2)
    GoSub cpcode 'md=(name,adcode,url,code)()
    AddComponent sc, m.SubMatches(0), m.SubMatches(1), md ', modName
    'code = Replace(code, m.Value, String$(Len(m.Value), 32))
    Mid(code, m.FirstIndex + 1, m.length) = String$(m.length, 32)
Next
'#COMPONENT proccessing
    
'#INCLUDE proccessing
'For Each m In RegExMatch(code, "^'#include ([@*~#+]?[^\r\n]*)$", -1) 'Add code
For Each m In RegExMatch(code, "^#include ([@*~#+]?[^\r\n]*)$", -1) 'Add code
    'code = Replace(code, m.Value, String$(Len(m.Value), 32))
    Mid(code, m.FirstIndex + 1, m.length) = String$(m.length, 32)
    tmp = m.SubMatches(0)
    If Len(tmp) > 1 Then
        GoSub cpcode
        For i = 0 To UBound(md) Step 4
            v = Array()
            If VarType(md(i + 3)) = vbString Then
                If Len(md(i + 3)) Then
                    If Len(md(i + 1)) Then md(i + 3) = md(i + 3) & vbCrLf & JSONUnEscape(md(i + 1)) & vbCrLf
                    v = Empty: v = sc.Eval(md(i))
                    'If IsEmpty(v) Then v = SCAddCode(sc, md(i + 3), IIf(VarType(modName) = vbString, md(i), Null))
                    If IsEmpty(v) Then v = SCAddCode(sc, md(i + 3), IIf(VarType(modName) = vbString, sNzS(modName, md(i)), Null))
                End If
            End If
            tmp = Nz(Nz(modName, sc.CodeObject.Me.Name), "NoMe")
            If IsArray(v) Then 'ERROR
                DebugPrint 240, tmp & "::#include " & md(i + 2)
                sc.AddCode tmp & "::#include " & md(i + 2)
            ElseIf IsEmpty(v) Then 'SUCCESS
                sc.AddCode "Dim " & md(i) & ":" & md(i) & "=" & Quot(md(i + 2))
                DebugPrint 32768, tmp & "::#include " & md(i) & " = " & md(i + 2)
            Else
                DebugPrint 100 + 32768, tmp & "::#include " & md(i) & " = " & md(i + 2)
            End If
            
        Next
    End If
Next
'#INCLUDE proccessing
    
Err.Clear
If Len(Replace(Replace(code, " ", ""), vbCrLf, "")) Then 'ADD VB CODE TO XCONTOL MODULE
    If NzS(modName, GlobalModule) = GlobalModule Then
        sc.AddCode code
    Else
    
#If xScript Then
        sc.AddCode code, modName
#Else
        sc.Modules.Add modName
        If Err Then Err.Clear
        sc.Modules(modName).AddCode code ': Debug.Print Join2(ToArray(sc.Modules(modName).Procedures), ",")
#End If
    End If
    If Err Then SCError sc
End If
    
'End If '  TypeName(qVBS) = "ScriptControl" Then 'NO ERROR ADD RUNTIME CODE

'mSC.Reset
Err.Clear
Exit Function

cpcode: 'Find code in.tmp=cpath(;;;), in.sc,  out.md=result (name,url,code)
    md = Array(): i = 0
    'For Each v In Split(tmp, ",")
    oa = RightStr(tmp, vbTab, 1)  'OpenArgs
    For Each v In Split(LeftStr(tmp & vbTab, vbTab), ",")
        ReDim Preserve md(UBound(md) + 4)
        md(i + 0) = LeftStr(v & vbTab, vbTab) 'Name
        'md(i + 1) = RightStr(v, vbTab, 1)  'OpenArgs
        With CParam(md(i + 0))
            md(i + 3) = .Name 'code
            If VarType(md(i + 3)) <> vbString Then
                If sc.SitehWnd = 0 Then Set so = ObjFromPtr(sc.CodeObject.SourceObjectPtr) Else Set so = sc.CodeObject.Me
                md(i + 3) = so.CPSource(md(i + 0))
                md(i + 2) = so.CPSource.Url
                .Url = NzS(md(i + 2), , , IIf(InStr(md(i + 2), "/"), "\", "/")) & md(i + 0)
            End If
            md(i + 2) = .Url
        End With
        md(i + 0) = RegExReplace(SplitIndex(SplitIndex(md(i + 0)), "/"), "[#@~*+ ]*", "") 'vbName
        i = i + 4
    Next
    If Len(oa) Then If UBound(md) = -1 Then md = Array("", oa, "", "") Else md(UBound(md) - 2) = oa
    Return
End Function


Function RemoveSQLComments$(ByVal sql, Optional conn) 'RemoveSQLComents(sql)
'REMOVE SQL multiline comments [/*  sql_comments */]  AND singleline coments [--  sql_comments ]
'sql = RegExReplace(Nz(sql), "(--[^\r\n]*)|(/\*[\w\W]*?(?=\*/)\*/)", vbNullString)
'sql = mSTR(sql)
'sql = RemoveSlashStarBlocks(S_(sql), , "--[^\r]*")
sql = PrepareCode("" & sql, "SQL")
Dim dbms$, prov$, p, s$
Dim cp As CParam ', ar ', ard
Set cp = CParam(Split("MS Jet,JET,Microsoft SQL Server,SQL,MySQL,MY,SQL Anywhere,SA", ","))
'ar = cp.GetArray(1)
'SQL PROVIDER SPEC:
'<JET: ~JET_SQL~ :JET> <SQL: ~SQL_SQL~ :SQL> <Provider: ~Provider_SQL~ :Provider>
On Error Resume Next: dbms = conn.Properties("DBMS Name"): Err.Clear
dbms = NzS(dbms, "SQL")

prov = Nz(cp(dbms), dbms) '"SQL")
Dim m As Match

'For Each m In RegExMatch(sql, "<([A-Z ]+):([\w\W]*?(?=:\1>)):\1>", -1)
For Each m In RegExMatch(sql, "<(\w+):([\w\W]*?(?=:\1>)):\1>", -1)
    s = m.Value: p = m.SubMatches(0)
    If p = prov Then p = m.SubMatches(1) Else p = ""
    sql = Replace(sql, s, p)
Next


's = "'[^']*'||\($\)|"
Select Case prov
Case "JET", "EXCEL"
    'CDate(A) * -> JET Format(A,'dd.mm.yy')
    'CAST(a as varchar) -> CStr(A)
    'ISNULL(A,B) SQL -> JET iif(IsNull(A),B,A)
    'NULLIF(A,B) SQL -> JET iif(A=B,null,A)
    s = "'[^']*'++CDate\(($)\)+Format(<a>,'dd.mm.yy')+CAST\(($) as varchar\)+(''&<a>)+CStr\(($)\)+(''& <a> &'')+\($\)++isnull\(($),($)\)+iif(IsNull(<a>),<b>,<a>)+nullif\(($),($)\)+iif(<a>=<b>,Null,<a>)+iif\($,$,$\)+"
Case "SQL" 'JET->SQL
    'CDate(A) JET -> SQL CONVERT(varchar,CAST(A-2 as datetime),4)
    'Format(A,'dd.mm.yy') JET -> SQL CONVERT(varchar,CAST(A-2 as datetime),4)
    'IIF(A,B,C) JET -> SQL CASE WHEN A THEN B ELSE C END
    s = "'[^']*'++CDate\(($)\)+CONVERT(varchar,CAST(<a>-2 as datetime),4)+CStr\(($)\)+CAST(ISNULL(<a>,'') as varchar)+\($\)++\($,$\)++iif\(($),($),($)\)+CASE WHEN <a> THEN <b> ELSE <c> END"
        
Case Else
    s = ""
End Select
Dim i&, ar: ar = Array(): cp.Clear
If Len(s) Then cp.Value = Split(Replace(s, "$", "[^(,)]*"), "+")
Do
For Each p In cp.Map
    'Debug.Print ">>>"; p.Name
    For Each m In RegExMatch(sql, p.Name, -1)
    's = m.Value
    'If InStr(s, "(") + InStr(s, ")") Then
        i = UBound(ar) + 1: ReDim Preserve ar(i)
        If m.SubMatches.count Then ar(i) = Replace(p.Value, "<a>", m.SubMatches(0)) Else ar(i) = m.Value
        If m.SubMatches.count > 1 Then ar(i) = Replace(ar(i), "<b>", m.SubMatches(1))
        If m.SubMatches.count > 2 Then ar(i) = Replace(ar(i), "<c>", m.SubMatches(2))
        sql = Replace(sql, m.Value, Chr(1) & i & Chr(1))
    'End If
    Next
    'Debug.Print sql
Next
Loop While RegExMatch(sql, Join(cp.GetArray, "|"), -1).count
For i = UBound(ar) To 0 Step -1: sql = Replace(sql, Chr(1) & i & Chr(1), ar(i)): Next


Err.Clear
'sql = Replace(sql, "[[", "{"): sql = Replace(sql, "]]", "}") 'SHAPE {} SYNTAX  [[ ]] PRESERVE <==> xcontrol.ReplaseVars + ReplaceArgs

RemoveSQLComments = sql
End Function

'Function RemoveSlashStarBlocks(src, Optional bSaveLength As Boolean, Optional sReplacePattern = "", Optional bLiteral34 As Boolean)
'Dim s$
's = S_(src)
'If Len(s) Then
'Dim i&, m$, m0$, m1$, n&
'Dim ns, pattern$
'
'
'
'If bLiteral34 Then
'    'pattern = "(?:""(?:""""|[^""\r])*"")"
'    pattern = """(?:""""|[^""\r\n])*""" & "|^'#.*$"
'    s = Replace(s, String$(2, 34), Chr$(20))  ' Шифруем DBLQUOT
'Else 'SQL strings
'    'pattern = "(?:'(?:''|[^'\r])*')"
'    'pattern = "'(?:''|[^'\r\n])*'"
'    pattern = "'(?:''|[^'\r\n])*'|""(?:""""|[^""\r\n])*"""
'    s = Replace(s, String$(2, 39), Chr$(21))
'End If
'
'
's = CompactString(s, ns, 22, pattern) 'Забираем строки
'
''i = 0
'Do
'    m = GTag(s, 1, "/*", "*/")
'    m1 = ""
'    'Debug.Print m
'    If bSaveLength Then
'        m0 = "/*" & m & "*/"
'        If Len(m) > 2 Then
'            For i = 0 To UBound(ns) 'Вставляем строки чтобы посчитать символы
'            m0 = Replace(m0, Chr$(22) & i & Chr$(22), ns(i))
'            Next
'        End If
'        m1 = RegExReplace(m0, "[^\s]", " ") 'Удаляем не пробельные символы на пробелы
'    End If
'    s = Replace(s, "/*" & m & "*/", m1) ',1,1)
'    'i = i + 1
'    If Len(m) = 0 Then m = Chr$(0): If GTag(s, -1, "/*", "*/") = 0 Then Exit Do
'Loop While Len(m) > 0 'And i < 1000 'Чистим вложенные slash star кооменты
'
'If Len(S_(sReplacePattern)) Then s = RegExReplace(s, sReplacePattern, "", IIf(bSaveLength, " ", ""))  ' удаляем SINGLE LINE comments
'
's = CompactString(s, ns, 22) 'Вставляем строки
'
's = Replace(s, Chr$(21), String$(2, 39)) ' Восстанавливаем DBLQUOT
's = Replace(s, Chr$(20), String$(2, 34))
'End If
'RemoveSlashStarBlocks = s
''Clipboard.SetText ""
''Clipboard.SetText s
'End Function

'Function CompactArray(sar, qsa, ByVal qMark, Optional ByVal pattern, Optional ByVal nest = 3)
'CompactArray = Compact(sar, qsa, qMark, pattern, nest)
'End Function
'Function CompactString(ByVal src$, qsa, ByVal qMark, Optional ByVal pattern, Optional ByVal nest = 3)
'CompactString = Compact(src, qsa, qMark, pattern, nest)
'End Function

Function Compact(ByVal sar, qsa, ByVal qMark, Optional ByVal pattern, Optional ByVal nest = 3)
On Error Resume Next
Dim i&, lv&, ar, ari&, src$
If IsArray(sar) Then ar = sar Else ar = Array(sar)
For ari = 0 To UBound(ar)
src = ar(ari)
lv = Val(nest)
If IsArray(qsa) And VarType(pattern) <> vbString Then 'Вставляем строки из массива sa()
    Do
        For i = UBound(qsa) To 0 Step -1
            src = Replace(src, ChrW$(qMark) & i & ChrW$(qMark), qsa(i))
            'Debug.Print src
            If InStr(src, ChrW$(qMark)) = 0 Then Exit For
        Next
        lv = lv - 1
    Loop While InStr(src, ChrW$(qMark)) And lv > 0
ElseIf VarType(pattern) = vbString Then  'Вырезаем строки и ложим их в массив sa()
    'Dim n&: If Not IsArray(qsa) Then qsa = Split("", ",") Else n = UBound(qsa) + 1
    Dim n&: If Not IsArray(qsa) Then qsa = Array() Else n = UBound(qsa) + 1
    If Len(src) Then
        'qMark = Nz(qMark)
        Dim rx As RegExp, b&, m As Match, ms: Set rx = RegEx
        Dim s$, p&, sb&: pattern = Split(pattern, vbCrLf)
        If UBound(pattern) > 0 Then sb = pattern(1)
        pattern = pattern(0)
        Do
            Set ms = RegExMatch(src, pattern, -1, rx) ': i = 0
            For i = ms.count - 1 To 0 Step -1
                p = 0: Set m = ms(i)
                If sb Then s = m.SubMatches(sb - 1): p = InStr(m.Value, m.SubMatches(sb - 1)) Else s = m.Value
                If Len(s) Then
                    n = gAddIndex(qsa, s)
                    If VarType(qMark) <> vbString Then   'Replace match with  [qmark$+n&+qmark$]
                        src = Left(src, m.FirstIndex + p) & ChrW$(qMark) & n & ChrW$(qMark) & Mid(src, m.FirstIndex + 1 + p + Len(s))
                    ElseIf Len(qMark) = 1 Then 'replace match with preserve length
                        'Debug.Print "<<" & String(Len(s), qMark) & ">>"
                        Mid(src, 1 + m.FirstIndex + p, Len(s)) = String(Len(s), qMark)
                    Else
                        src = Left(src, m.FirstIndex + p) & qMark & Mid(src, m.FirstIndex + 1 + p + Len(s))
                    End If
                    'Debug.Print src
                End If
            Next
            lv = lv - 1
        Loop While ms.count And lv > 0
    End If
End If
ar(ari) = src
Next
If IsArray(sar) Then Compact = ar Else Compact = ar(0)
End Function

'Function vbRemoveComents(code) As String
''replace /* multiline c++ comments */
'
''#string string_id = string_expression^$  ====>> replace #string_id# by Quot(string_expression)
''#define define_id = value^$  ====>> replace #define_id# by value
''#compute compute_id = value^$  ====>> replace #compute_id# by Eval(value)
'
''MULTILINE  =
''#(define|string)< (define_id|string_id)=FIRST LINE of MULTILINE STRING^$
''..........^$
''Last LINE of MULTILINE STRING^$
''#(define|string)> (define_id|string_id)^$
'
'Dim Item, dq As New CParam, s$, nm$, i&, v
'Dim fitem$ ', ns&
'Dim s0$
'
''code = mSTR(code)
'
'vbRemoveComents = vbLf & RemoveSlashStarBlocks(code, True, , 1)
''vbRemoveComents = RegExReplace(vbRemoveComents, "^\s?'[^#].*$", "", 32)
'
'For Each Item In Split("compute ,define< ,define ,string< ,string ", ",") '
'dq.Clear
'fitem = vbLf & "'#" & Item
''If InStr(vbRemoveComents, fitem) > 0 Then     'Есть что менять.
'    'For i = 1 To GTag(vbRemoveComents, -1, fitem, vbCr)  'Сколько объявлений
'    Do While InStr(vbRemoveComents, fitem)
'        s0 = GTag(vbRemoveComents, 1, fitem, vbCr) 'Берем 1-у строку
'        'ns = InStr(vbRemoveComents, vbCr & fitem & s & vbCr) + Len(fitem & s0) - 1 'MULTILINE COMENT START
'        nm = LeftStr(s0, "=") 'Получаем имя
'        If Len(nm) Then 'Есть имя
'            If InStr(Item, "<") Then
'                s = GTag(vbRemoveComents, 1, fitem & nm & "=" & vbCrLf, vbCrLf & "'#" & Replace(Item, "< ", "> ") & nm & vbCrLf)
'                s0 = nm & "="
'                If Len(s) Then s0 = s0 & vbCrLf & s & vbCrLf & "'#" & Replace(Item, "<", ">") & nm & vbCrLf
'            Else
'                s = RightStr(s0, "=", 1) 'получаем значение
'                's0 = nm & "=" & s
'            End If
'            If Len(s) Then
'            'Debug.Print s
'                'STRING
'                If InStr(Item, "string") Then
'                    If InStr(s, vbCrLf) > 0 Then
'                        s = Replace(s, vbCrLf, """&vbCrLf&_" & vbCrLf & """")
'                    Else
'                    'Debug.Print Len(s)
'                        s = Replace(s, Chr$(34), String$(2, 34))
'                    End If
'                End If
'                If Item = "compute " Then
'                    On Error Resume Next
'                    'v = S_(VBScript("*", Empty).Eval(s))
'                    'v = "": v = Nz(mVBScript.Eval(s))
'                    v = "": v = Nz(mSC.Eval(s))
'                    s = v
'                    DebugPrint &HFF0000, ":#compute " & nm & " = [" & s & "] " & Err.Description
'                End If
'                dq(nm) = s
'            Else
'                'Exit Do
'
'            End If
'        End If
'        'Debug.Print vbRemoveComents
'        vbRemoveComents = Replace(vbRemoveComents, fitem & s0, vbLf & String(CharCount(s0, vbCrLf), vbCrLf)) 'REMOVE '#....
'        'Debug.Print vbRemoveComents
'    Loop 'Next
'    For i = 0 To dq.ParamCount - 1
'        vbRemoveComents = Replace(vbRemoveComents, "#" & dq.ParamName(i) & "#", dq.Param(i))
'    Next
''End If '#item = exist
'
'Next 'item
'
'
'If InStr(vbRemoveComents, "<$""") Then 'MULTILINE STRING
'    Do
'        s = GTag(vbRemoveComents, 1, "<$""", """$>")
'        If Len(s) Then vbRemoveComents = Replace(vbRemoveComents, "<$""" & s & """$>", """" & Replace(s, vbCrLf, """&vbCrLf&_" & vbCrLf & """") & """")
'    Loop While Len(s)
'End If
'If Left(vbRemoveComents, 1) = vbLf Then vbRemoveComents = Mid(vbRemoveComents, 2)
''Debug.Print
''Debug.Print vbRemoveComents
'End Function
'

'
'Function vbDecorateFloatCode(ByVal code)
''Remove FLOAT STATEMENT для теста прогрммы
''Replace fullLINE comments = ""
''ДЛЯ ПРОВЕРКИ СИНТАКСИСА коментируем все кроме Sub/ function
'Dim ar, sas, sa, i&, c$
'code = Replace(code, vbCrLf, vbLf)
'code = CompactString(code, sas, 16, """(?:""""|[^""\r\n])*""") 'Забираем строки
'code = CompactString(code, sa, 15, "^(?: |\t)*(?:public(?: |\t)*(?:default(?: |\t)*)?|private(?: |\t)*)?(class|sub|function)(?: |\t)*[\S\s]+?\s?end[\t ]+\1.*?$") 'Забираем sub+function+class в массив SA()
'ar = Split(code, vbLf)
'For i = 0 To UBound(ar)
'    If Len(ar(i)) Then
'        c = Left$(ar(i), 1)
'        If Not (c = Chr$(15) Or c = Chr$(39)) Then ar(i) = "'" & ar(i) 'comment vb line
'    End If
'Next
'code = Join(ar, vbCrLf)
'code = CompactString(code, sa, 15) ' Вставляем обратно  sub+function мз массива SA()
'code = CompactString(code, sas, 16) ' Вставляем обратно  строки
'vbDecorateFloatCode = code
''Debug.Print code
'End Function
'=================VBScript=================



'=======GET DATE ===============
Public Function GetDate(ByVal hWndMain, ByVal OpenPos, ByVal mDate, Optional ByVal dmin = 0, Optional ByVal dmax = 99999)
Dim v, ow As xControl, cp As New CParam
'cp.Value = ArrayShift(Split(",AfterUpdate,this.FormResult=this.Value,Visible,1,Type,Date,Name,cal,ToolWindow,1,Border,1,BackGround,15,Format,d mmmm yyyyг.,vmin," & dmin & ",vmax," & dmax, ","), -1)
cp.Value = Split("NoEvent,1,Visible,1,Type,Date,Name,cal,ToolWindow,1,Border,1,BackGround,15,Format,d mmmm yyyyг.,vmin," & dmin & ",vmax," & dmax, ",")
Set ow = xControl(hWndMain)
If OpenPos = 0 Then ' MODAL POPUP
    cp.AddParams Split("Border,3,TitleBar,1,Popup,1,StartPos,2,~Events,&H200,Create,this.Value=" & mDate, ",")
    v = DoModal(cp.Source, ow)
Else 'DROP VALUE
    cp.AddParams Array("IamDrop", 1, "DropX", LOWORD(L_(OpenPos)), "DropY", HIWORD(L_(OpenPos)))
    v = DropValue(ow, cp.Source, mDate)
End If
If VarType(v) > 1 Then GetDate = modMain.Cast(v, vbLong)
End Function

Function FLDat(ByVal dat As Long, dType, ByVal num As Byte) As Long
'DType={0;ПЕРИОД; 1;ДЕНЬ; 2;НЕДЕЛЯ; 3;МЕСЯЦ; 4;КВАРТАЛ; 5;ГОД}
'num={1; Начало периода; Not 1; Конец периода}
Dim d As Date
d = dat
On Error Resume Next
Select Case dType
'0;ПЕРИОД;1;ДЕНЬ
Case 0, 1: FLDat = dat
'2;НЕДЕЛЯ
Case 2: If num = 1 Then FLDat = DateAdd("d", -Val(Format$(d, "w", 2)) + 1, dat) Else FLDat = DateAdd("d", 7 - Val(Format$(d, "w", 2)), dat)
'3;МЕСЯЦ
Case 3: If num = 1 Then FLDat = Fdat(dat) Else FLDat = LDat(dat)
'4;КВАРТАЛ
Case 4: If num = 1 Then FLDat = Fdat(CDate("2." & CStr((Val(Format$(d, "q"))) * 3 - 2) & "." & Format$(d, "yyyy"))) Else FLDat = LDat(CDate("2." & CStr((Val(Format$(d, "q"))) * 3) & "." & Format$(d, "yyyy")))
'5;ГОД
Case 5: If num = 1 Then FLDat = CDate("1.1." & Format$(d, "yyyy")) Else FLDat = DateAdd("yyyy", 1, CDate("1.1." & Format$(d, "yyyy"))) - 1
End Select

Err.Clear
End Function

Function Fdat(ByVal d As Long) As Long 'Первое число месяца
Fdat = d: On Error Resume Next: Fdat = CDate("1." & Format$(d, "mm.yyyy")): Err.Clear
End Function

Function LDat(ByVal d As Long) As Long 'Последнее число месяца
LDat = d: On Error Resume Next: LDat = DateAdd("m", 1, Fdat(d)) - 1: Err.Clear
End Function

Sub GetPerDate(ByVal d, ByVal per, ByRef d1, ByRef d2)    'Период где есть эта дата
On Error Resume Next: d1 = FLDat(d, per, 1): d2 = FLDat(d, per, 2): Err.Clear
End Sub

Sub ADNDatePeriodEx(ByVal dp0, ByRef dp1, ByRef dp2, ByVal a As Integer)
Dim per As String ', d As Long
dp0 = Cast(dp0, vbLong) Mod 6
per = Choose(dp0 + 2, "Y", "Y", "Y", "ww", "m", "Q", "yyyy", "Y", "Y", "Y")
If dp0 > 0 And dp0 < 6 Then 'день,неделя,квартал,год
    dp1 = CLng(FLDat(DateAdd(per, a, dp1), dp0, 1))
    dp2 = CLng(FLDat(DateAdd(per, a, dp2), dp0, 2))
ElseIf dp0 = 0 Then 'период
    dp1 = dp1 + a
    dp2 = dp2 + a
End If
End Sub

Function GetDPerText(ByVal dat1, ByVal per&, ByVal dat2) As String
On Error Resume Next
Dim s$
If dat1 = dat2 Then per = 1
Select Case per
Case -1: GetDPerText = ".. -  ..": s = Format$(dat1, "dd.mm.yy") & "-" & Format$(dat2, "dd.mm.yy")
Case 1: GetDPerText = "..": s = Format$(dat1, "ddd d mmm yyyyг.")
Case 2: GetDPerText = "Неделя": s = Format$(dat1 + 3, "ww неделя yyyyг.", 1) '2)
Case 3: GetDPerText = "Месяц": s = Format$(dat1, "mmmm yyyyг.")
Case 4: GetDPerText = "Квартал": s = Format$(dat1, "q  квартал yyyyг.")
Case 5: GetDPerText = "Год": s = Format$(dat1, "yyyy год")
Case Else: GetDPerText = "Период": s = "< Период " & CLng(dat2 - dat1 + 1) & ">"
End Select
If Len(s) Then GetDPerText = s
Err.Clear
End Function
'=======GET DATE ===============


'=======================================================

'============= CURSORS  ===============================


'=======================================================

'============== POPUP MENU ===========================
Function PopupMenu(hWndOwner, src, Optional Left, Optional Top, Optional ReturnCMD = 1)
Dim pm As New CMenu
pm.Bild 1, S_(src) ', Base.UserRole
PopupMenu = pm.ShowPopup(L_(hWndOwner), L_(Left), L_(Top), L_(ReturnCMD))
End Function


Function MenuItemState(ByVal hWnd&, ByVal bSysMenu As Boolean, ByVal uIDItem&, ByVal NewState&)
Dim hMenu&
If bSysMenu Then hMenu = GetSystemMenu(hWnd, 0) Else hMenu = GetMenu(hWnd)
If NewState = -1 Then
    RemoveMenuItem hMenu, uIDItem, 0
Else
    MenuItemState = EnableMenuItem(hMenu, uIDItem, NewState)
End If
End Function

Function MenuItemText(ByVal hWnd, ByVal bSysMenu As Boolean, ByVal uIDItem&, ByVal sText)
Dim hMenu&, s$
If bSysMenu Then hMenu = GetSystemMenu(hWnd, 0) Else hMenu = GetMenu(hWnd)
If VarType(sText) < 2 Then
    RemoveMenuItem hMenu, uIDItem, 0
Else
    s = sText
    If InsertMenuItem(hMenu, -1, &H1000&, uIDItem, s) Then MenuItemText = s
End If
End Function



'Function WindowMenu(ByVal hWnd, ByVal getSysMenu As Boolean, Optional newhMenu As Long = -1)
'If newhMenu = -1 Then 'GET
'    If getSysMenu Then WindowMenu = GetSystemMenu(0& + hWnd, 1) Else WindowMenu = GetMenu(0& + hWnd)
'Else 'LET
'    If getSysMenu Then WindowMenu = SetSystemMenu(0& + hWnd, 0& + newhMenu) Else WindowMenu = SetMenu(0& + hWnd, 0& + newhMenu)
'End If
'End Function

'Function NewMenuItem(ByVal hMenu, ByVal uIDItem, ByVal wFlags, ByVal sProp) As Long
'Dim s$
's = Nz(sProp)
'If hMenu = 0 And uIDItem = 0 And wFlags = 0 Then
'    NewMenuItem = CreatePopupMenu() 'CreateMenu()
'Else
'    If InsertMenuItem(0& + hMenu, -1, 0& + wFlags, (0& + uIDItem), s) Then NewMenuItem = uIDItem
'End If
''Const MF_BYCOMMAND As Long = &H0&
''Const MF_STRING As Long = &H0&
''Const MF_ENABLED As Long = &H0&
''Const MF_GRAYED As Long = &H1&
''Const MF_DISABLED As Long = &H2&
''Const MF_BITMAP As Long = &H4&
''Const MF_CHECKED As Long = &H8&
''Const MF_POPUP As Long = &H10&
''Const MF_MENUBARBREAK As Long = &H20&
''Const MF_MENUBREAK As Long = &H40&
''Const MF_HILITE As Long = &H80&
''Const MF_OWNERDRAW As Long = &H100&
''Const MF_DELETE As Long = &H200&
''Const MF_BYPOSITION As Long = &H400&
''Const MF_SEPARATOR As Long = &H800&
''Const MF_DEFAULT As Long = &H1000&
'
'End Function
'Property Get MenuItem(ByVal hMenu, ByVal uIDItem, ByVal sProp)
''(uIDItem and &H400)=0  ---> BYCOMMAND command=uIDItem
''(uIDItem and  &H400)=&H400  -->  BYPOSITION position = (uIDItem AND &HFAFF)
'hMenu = aVal(hMenu)
'If hMenu = 0 Then Exit Property
'Dim mi As MENUITEMINFO, s$
'mi.cbSize = LenB(mi)
''Public Const MIIM_ID As Long = &H2
''Public Const MIIM_STATE As Long = &H1
''Public Const MIIM_SUBMENU As Long = &H4
''Public Const MIIM_CHECKMARKS As Long = &H8
''Public Const MIIM_TYPE As Long = &H10
''Public Const MIIM_DATA As Long = &H20
''Public Const MIIM_STRING As Long = &H40
''Public Const MIIM_BITMAP As Long = &H80
''Public Const MIIM_FTYPE As Long = &H100
'Select Case sProp
'Case "State", "Enabled", "Checked", "Default": mi.fMask = 1
'Case "ID": mi.fMask = 2
'Case "SubMenu": mi.fMask = &H4
''Case "CheckMark": mi.fMask = &H8
'Case "Type": mi.fMask = &H10 '0
'Case "Data": mi.fMask = &H20
'Case "Text": mi.fMask = &H40: mi.cch = 255: s = String(mi.cch, 0): mi.dwTypeData = StrPtr(s)
'Case "Bitmap": mi.fMask = &H80
'Case Else: Exit Property
'End Select
'GetMenuItemInfo 0& + hMenu, (0& + uIDItem) And &HFAFF&, (uIDItem And &H400) = &H400, mi
'Select Case sProp
'Case "Text": MenuItem = StrConv(s, vbUnicode)
'Case "ID": MenuItem = mi.wID
'Case "Type": MenuItem = mi.fType
'Case "State": MenuItem = mi.fState
'Case "Enabled": MenuItem = (mi.fState And 2) = 0
'Case "Checked": MenuItem = (mi.fState And 8) > 0
'Case "Default": MenuItem = (mi.fState And &H1000) > 0
'Case "SubMenu": MenuItem = mi.hSubMenu
'Case "Data": MenuItem = mi.dwItemData
''Case "CheckMark": MenuItem = mi.hbmpChecked
'Case "Bitmap": MenuItem = mi.hbmpItem
'End Select
'End Property

'Property Let MenuItem(ByVal hMenu, ByVal uIDItem, ByVal sProp, ByVal newValue)
'hMenu = aVal(hMenu)
'If hMenu = 0 Then Exit Property
'Dim mi As MENUITEMINFO, fState&, s$
'mi.cbSize = LenB(mi)
'fState = &HFF And GetMenuState(0& + hMenu, (0& + uIDItem) And &HFAFF&, (0& + uIDItem) And &H400)
'Select Case sProp
'Case "State": mi.fMask = 1: mi.fState = aVal(newValue)
'Case "Enabled": mi.fMask = 1:: mi.fState = (fState And (Not 3&)) Or IIf(aVal(newValue), 0&, 3&)
'Case "Checked": mi.fMask = 1: mi.fState = (fState And (Not 8&)) Or IIf(aVal(newValue), 8&, 0&)
'Case "Default": mi.fMask = 1: mi.fState = (fState And (Not &H1000&)) Or IIf(aVal(newValue), &H1000&, 0&)
'Case "ID": mi.fMask = 2: mi.wID = aVal(newValue)
''Const SC_SIZE As Long = &HF000&
''Const SC_MOVE As Long = &HF010&
''Const SC_MINIMIZE As Long = &HF020&
''Const SC_MAXIMIZE As Long = &HF030&
''Const SC_NEXTWINDOW As Long = &HF040&
''Const SC_PREVWINDOW As Long = &HF050&
''Const SC_CLOSE As Long = &HF060&
''Const SC_VSCROLL As Long = &HF070&
''Const SC_HSCROLL As Long = &HF080&
''Const SC_MOUSEMENU As Long = &HF090&
''Const SC_KEYMENU As Long = &HF100&
''Const SC_ARRANGE As Long = &HF110&
''Const SC_RESTORE As Long = &HF120&
''Const SC_TASKLIST As Long = &HF130&
''Const SC_SCREENSAVE As Long = &HF140&
''Const SC_HOTKEY As Long = &HF150&
''Const SC_DEFAULT As Long = &HF160&
''Const SC_MONITORPOWER As Long = &HF170&
''Const SC_CONTEXTHELP As Long = &HF180&
'
''Const IDOK As Long = 1
''Const IDCANCEL As Long = 2
''Const IDABORT As Long = 3
''Const IDRETRY As Long = 4
''Const IDIGNORE As Long = 5
''Const IDYES As Long = 6
''Const IDNO As Long = 7
''Const IDCLOSE As Long = 8
''Const IDHELP As Long = 9
''Const IDTRYAGAIN As Long = 10
''Const IDCONTINUE As Long = 11
'
'Case "SubMenu": mi.fMask = &H4: mi.hSubMenu = aVal(newValue)
'Case "Type": mi.fMask = &H10: mi.fType = aVal(newValue)
''MFT_STRING 0x00000000L
''MFT_BITMAP 0x00000004L
''MFT_MENUBARBREAK 0x00000020L
''MFT_MENUBREAK 0x00000040L
''MFT_OWNERDRAW 0x00000100L
''MFT_RADIOCHECK 0x00000200L
''MFT_RIGHTJUSTIFY 0x00004000L
''MFT_RIGHTORDER 0x00002000L
''MFT_SEPARATOR 0x00000800L
'
'Case "Data": mi.fMask = &H20: mi.dwItemData = aVal(newValue)
'Case "Text": mi.fMask = &H40: s = StrConv(Nz(newValue), vbFromUnicode): mi.cch = LenB(s): mi.dwTypeData = StrPtr(s)
'Case "Bitmap": mi.fMask = &H20: mi.hbmpItem = aVal(newValue) '1..11
''HBMMENU_CALLBACK ((HBITMAP) -1)  A bitmap that is drawn by the window that owns the menu. The application must process the WM_MEASUREITEM and WM_DRAWITEM messages.
''HBMMENU_SYSTEM ((HBITMAP) 1) Windows icon or the icon of the window specified in dwItemData.
''HBMMENU_MBAR_RESTORE 0((HBITMAP) 2) Restore button for the menu bar.
''HBMMENU_MBAR_MINIMIZE ((HBITMAP) 3) Minimize button for the menu bar.
''(HBITMAP) 4)
''HBMMENU_MBAR_CLOSE ((HBITMAP) 5) Close button for the menu bar.
''HBMMENU_MBAR_CLOSE_D ((HBITMAP) 6) Disabled close button for the menu bar.
''HBMMENU_MBAR_MINIMIZE_D ((HBITMAP) 7) Disabled minimize button for the menu bar.
''HBMMENU_POPUP_CLOSE ((HBITMAP) 8) Close button for the submenu.
''HBMMENU_POPUP_RESTORE ((HBITMAP) 9) Restore button for the submenu.
''HBMMENU_POPUP_MAXIMIZE ((HBITMAP) 10) Maximize button for the submenu.
''HBMMENU_POPUP_MINIMIZE ((HBITMAP) 11) Minimize button for the submenu.
'
'Case Else: Exit Property
'End Select
'SetMenuItemInfo 0& + hMenu, (0& + uIDItem) And &HFAFF&, (0& + uIDItem) And &H400, mi
''RedrawMenu
'End Property

'Function ReDrawMenu(ByVal hWnd)
'DrawMenuBar 0& + hWnd
'End Function


'============== POPUP MENU ===========================

Function CollectionKeys(oColl As Collection)
Dim s, n&, p&, i&, k&
p = ObjPtr(oColl): If p = 0 Then Exit Function
s = Split("", ",") 's = Array()
i = GMem4(ByVal p + 16) 'items count
If i Then
    ReDim s(i - 1): i = p
    Do While i
        i = GMem4(ByVal i + 24) 'item ptr
        If i Then
            k = GMem4(ByVal i + 16) 'string key ptr
            If k Then s(n) = SysAllocString(k) 'Else s(n) = n + 1
            n = n + 1
        End If
    Loop
End If
CollectionKeys = s
End Function

'============== CONNECTION  ===========================
Property Get xConnectionName(Optional ByVal key As String) As String
On Error Resume Next
If key = "*" Then
    xConnectionName = Join(CollectionKeys(m_Cons), ",")
Else
    Dim x As Connection: Set x = xcon(key)
    If Not x Is Nothing Then
        With x.Properties
             'xConnectionName = NzS(NzS(, .Item("Data Source")), , "(" & .Item("DBMS Name") & ") ")
'             Debug.Print key, .Item("Server Name")
'             Debug.Print key, .Item("Current Catalog")
'             Debug.Print key, .Item("Data Source")
'             Debug.Print key, .Item("DBMS Name")
             xConnectionName = NzS(.Item("Server Name"), , , ":" & .Item("Current Catalog"))
             xConnectionName = NzS(xConnectionName, .Item("Data Source"))
             xConnectionName = NzS(xConnectionName, , "(" & .Item("DBMS Name") & ") ")
             'xConnectionName = NzS(NzS(NzS(.Item("Server Name"), , , ":" & .Item("Current Catalog")), .Item("Data Source")), , "(" & .Item("DBMS Name") & ") ")
        End With
    End If
End If
End Property
Property Get XConnection(Optional ByVal key As String) As Connection
On Error Resume Next: Set XConnection = xcon(key)
End Property
Private Function xcon(ByVal key As String) As Connection
key = Trim(key)
On Error Resume Next: If Len(key) Then Set xcon = m_Cons(IIf(IsNum(key), L_(key), key)) Else Set xcon = m_Cons(1)
End Function
Property Set XConnection(Optional ByVal key As String, ByVal cs)
XConnection(key) = cs
End Property
'Property Let XConnection(Optional ByVal key As String, ByVal cs)
'If Not xcon(key) Is Nothing Then If Len(key) Then m_Cons.Remove key Else m_Cons.Remove 1
'If Len(key) And m_Cons.count = 0 Then m_Cons.Add New Connection 'Add EMPTY DEFAULT CONECTION
'Dim x As Connection
'On Error Resume Next
'If TypeName(cs) = "Connection" Then Set x = cs Else If Len(Nz(cs)) Then Set x = New ADODB.Connection: x.Open cs
'If Not x Is Nothing Then If Len(key) Then m_Cons.Add x, key Else If m_Cons.count Then m_Cons.Add x, , m_Cons(1) Else m_Cons.Add x
'If Len(key) = 0 Then m_Param("schema") = Null
'End Property

Property Let XConnection(Optional ByVal key As String, ByVal cs)
If Not xcon(key) Is Nothing Then If Len(key) Then m_Cons.Remove key Else m_Cons.Remove 1
Dim x As Connection
On Error Resume Next
If TypeName(cs) = "Connection" Then Set x = cs Else Set x = New ADODB.Connection: If Len(Nz(cs)) Then x.Open cs
If Len(key) Then
    If Len(cs) Then m_Cons.Add x, key
Else
    If m_Cons.count Then m_Cons.Add x, , 1 Else m_Cons.Add x
End If
If Len(key) = 0 Then m_Param("schema") = Null
End Property

Sub xConnectionReset()
Set m_Cons = New Collection
XConnection = New Connection
XConnection.CommandTimeout = 10
End Sub
Property Get ConState(Optional ByVal xcon) As Byte
On Error Resume Next
If IsObject(xcon) Then ConState = xcon.State Else ConState = XConnection.State
End Property
Friend Function pxcon(sql) As Connection
Dim ar: ar = Split(sql, Chr(1))
On Error Resume Next
If UBound(ar) = 1 Then
    Set pxcon = New Connection
    pxcon.Open ar(0)
    sql = ar(1)
ElseIf Left(sql, 1) = "@" Then
    sql = Mid(sql, 2)
    Set pxcon = xmdb_conn
ElseIf Left(sql, 1) = "#" Then
    ar = Split(sql, "#")
    Set pxcon = XConnection(ar(1))
    sql = Right(sql, Len(sql) - Len(ar(1)) - 2)
Else
    Set pxcon = XConnection
End If
End Function

Public Function RunSQL(ByVal sql0, ParamArray nFields())
On Error Resume Next

Dim ar, i&, sql$, isql
Dim xconn As Connection, rs As Recordset

Select Case TypeName(sql0)
Case "String"
'    ar = Split(sql0, Chr(0))
'    If UBound(ar) = 1 Then
'        Set xconn = New Connection
'        xconn.Open ar(0)
'        sql0 = ar(1)
'    ElseIf Left(ar(0), 1) = "@" Then
'        sql0 = Mid(ar(0), 2)
'        Set xconn = xmdb_conn
'    ElseIf Left(ar(0), 1) = "#" Then
'        sql0 = Mid(ar(0), 2)
'        Set xconn = xmdb_conn
'    Else
'        Set xconn = XConnection
'    End If
'    If xconn.State = 0 Then Exit Function
'    sql = RemoveSQLComments(sql0, xconn)
    
    Set xconn = pxcon(sql0)
    If xconn.State = 0 Then Exit Function
    sql = RemoveSQLComments(sql0, xconn)

Case "Recordset"
    Set rs = sql0
Case Else
    Exit Function
End Select

Err.Clear
If UBound(nFields) > -1 Then
    If rs Is Nothing Then For Each isql In Split(sql, ";" & vbCrLf): Set rs = xconn.Execute(isql, , IIf(InStr(isql, " "), adCmdText, -1)): Next
    ar = rs.GetRows(1)
    Dim a2 As Boolean: a2 = modMain.ArrayDims(ar) = 2
    If Not a2 Then ReDim ar(UBound(nFields), 0)
    
    'If modMain.ArrayDims(ar) = 2 Then
    If UBound(nFields) = 0 Then
        If nFields(0) < 0 Then  'RunSQL({sql0} ,-1) return FieldName
        '??????????????????????????????????????????????????
            RunSQL = rs.Fields(Abs(nFields(0)) - 1).Name
        '??????????????????????????????????????????????????
        Else 'Return Value
            RunSQL = ar(0 + nFields(0), 0)
        End If
    Else 'Update nFields + Return nFields
        For i = 0 To UBound(nFields)
            If IsObject(nFields(i)) Then nFields(i).Value = ar(i, 0) Else nFields(i) = ar(i, 0)
        Next
        If a2 Then RunSQL = nFields
    End If
    'End If
    Set rs = Nothing
Else 'Execute SQL
    For Each isql In Split(sql, ";" & vbCrLf): xconn.Execute isql, , adExecuteNoRecords: Next
   
End If
'If Err <> 0 And (Err <> 3021) Then LastError: DebugPrint 0, "RunSQL = " & sql & vbCrLf & LastErrorDescription & " " & LastErrorNumber
'If Err Then Err.Clear
Err.Clear
End Function


Public Function getString$(Source, Optional ByVal colStart$, Optional ByVal colEnd$, Optional ByVal rowStart$, Optional ByVal rowEnd$, Optional ByVal sarColumns, Optional nEscapeOptions&)
Dim s$, ar
'Dim xar, sar
On Error Resume Next
'sar = Array() 'Очередность столбцов в результате
If TypeName(Source) = "Recordset" Then
    ar = Source.GetRows(, , sarColumns)
ElseIf VarType(Source) And vbArray Then
    ar = Source
'    If Not IsMissing(sarColumns) Then If VarType(sarColumns) Then If VarType(sarColumns) And vbArray Then sar = sarColumns Else sar = Split(vbNullString & sarColumns, ",")
Else
    'If ConState = 0 Then Exit Function
    s = Source
    Dim xc As Connection: Set xc = pxcon(s)
    'ar = XConnection.Execute(RemoveSQLComments(Source, XConnection)).GetRows(, , sarColumns)
    ar = xc.Execute(RemoveSQLComments(s, xc)).GetRows(, , sarColumns)
End If
Dim cDims%, i&
'X = ArrayDims(ar)
cDims = ArrayDims(ar)
If cDims = 0 Or cDims > 2 Then Exit Function
'If x = 0 Then Err.Clear: Exit Function
    
If cDims = 1 Then
    s = Join( _
        BETag(ar, colStart, colEnd, nEscapeOptions), _
        IIf(rowStart = rowEnd, rowStart, rowEnd & rowStart))
    If rowStart <> rowEnd Then s = rowStart & s & rowEnd
ElseIf cDims = 2 Then
    Dim st As New Stream: st.Type = adTypeText: st.Open
    For i = LBound(ar, 2) To UBound(ar, 2)
    
    st.WriteText _
        IIf(rowStart <> rowEnd, rowStart, "") & _
        Join( _
            BETag(pArrayLine(ar, i), colStart, colEnd, nEscapeOptions), _
            "") & _
        IIf(rowStart <> rowEnd, rowEnd, "")
    
    Next
    st.Position = 0: s = st.ReadText
End If


getString = s
depricate "cmain.getString"

End Function




Public Function GetRows(ByVal sql, Optional ByVal ret = 10, Optional ByVal spCols, Optional ByVal spNames)
'Dim mx As CMatrix: Set mx = CMatrix(IIf(TypeName(sql) = "Recordset", 8, 5), sql)
'GetRows = mx.GetRows(IIf(WithHeader, 1, 10), spCols, spNames)
GetRows = CMatrix(IIf(TypeName(sql) = "Recordset", 8, 5), sql).GetRows(ret, spCols, spNames)
depricate "cmain.getrows", "cmatrix.getrows"
End Function


'Function LastErr(Optional ByVal flag& = -1)
'Static ar: If Not IsArray(ar) Then ar = Array(Array(0, "")): ReDim Preserve ar(9)
'If flag < 0 Then   'SAVE ERR
'    If ar(0)(0) <> Err.Number Then pArrayLine ar, 0, Array(Err.Number, Err.Description), 2: ReDim Preserve ar(9)
'End If
'flag = flag And 30
'LastErr = ar(flag Mod 10)(0) 'RET LAST ERR NUMBER
'If flag > 9 Then LastErr = IIf(flag > 19, LastErr & ":", "") & ar(flag Mod 10)(1) 'RET LAST ERR DESCRIPTION
'End Function

Public Function SetFields(sql, ByVal retFieldName, ParamArray arNewValues())
'
Dim i&, rs As New Recordset, xcon As Connection, ar
Dim arValues, arNames ', arv
On Error Resume Next
retFieldName = S_(retFieldName)
If TypeName(sql) = "Recordset" Then
    Set rs = sql
    Set xcon = rs.ActiveConnection
Else
'    If ConState = 0 Then Exit Function
'    sql = S_(sql)
'    If Len(sql) = 0 Or UBound(arNewValues) = -1 Then Exit Function
'    Set xcon = XConnection

    Set xcon = pxcon(sql)
    If Len(sql) = 0 Or UBound(arNewValues) = -1 Then Exit Function
    If ConState(xcon) = 0 Then Exit Function
    
    If xcon.Provider = "SQLOLEDB.1" Then
        rs.CursorLocation = adUseClient 'Поле счетчика для новой записи
        rs.Open RemoveSQLComments(sql, xcon), xcon, adOpenForwardOnly, adLockBatchOptimistic  'adLockBatchOptimistic  'adLockOptimistic      'adOpenForwardOnly, adLockOptimistic  'adOpenKeyset adOpenStatic
    Else
        rs.Open RemoveSQLComments(sql, xcon), xcon, adOpenKeyset, adLockOptimistic    'adLockBatchOptimistic  'adOpenDynamic      'adOpenForwardOnly, adLockOptimistic  'adOpenKeyset adOpenStatic
    End If
End If

If rs.State = adStateOpen Then
    If ArrayDims(arNewValues(0)) = 1 Then ar = arNewValues(0) Else ar = arNewValues
    i = UBound(ar)
    If i > rs.Fields.count - 1 Then i = rs.Fields.count - 1
    ReDim arNames(i)
    ReDim arValues(i)
'    ReDim arv(4, i)
    
    For i = 0 To UBound(arNames)
        arNames(i) = rs.Fields(i).Name
        arValues(i) = modMain.Cast(ar(i), rs.Fields(i).Type, rs.Fields(i).Attributes And adFldIsNullable)
'        If (rs.Fields(i).Attributes And adFldIsNullable) = 0 And VarType(ar(i)) < 2 Then Debug.Print "SETFIELDS field." & arNames(i) & " NOT NULLABLE CAST =" & arValues(i)
'        arv(0, i) = arNames(i)
'        arv(1, i) = rs.Fields(i).Type
'        arv(2, i) = ar(i)
'        arv(3, i) = TypeName(ar(i))
'        arv(4, i) = TypeName(arValues(i))
    Next
'    Debug.Print "SETFIELDS ", sql, Join2(arv, , , "<Null>")
    Err.Clear

    xcon.BeginTrans
    If (rs.RecordCount <= 0 Or (rs.BOF And rs.EOF)) And Len(retFieldName) > 0 Then 'ADD NEW
        rs.AddNew arNames, arValues
        rs.UpdateBatch ' adAffectCurrent
        If Len(retFieldName) Then SetFields = rs.Fields(retFieldName).Value
    Else 'UPDATE
        xcon.Errors.Clear
        Do While Not rs.EOF And xcon.Errors.count = 0 'Пока есть записи для обновления DOLOOP
            rs.UpDate arNames, arValues
            rs.UpdateBatch adAffectCurrent
            If Len(retFieldName) Then SetFields = rs.Fields(retFieldName).Value
            If Not rs.EOF Then rs.MoveNext
        Loop
    End If
    rs.Close
    If Err Then
        xcon.RollbackTrans
        GoTo exit_function
    End If
    xcon.CommitTrans
End If

exit_function:
If Err Then
    DebugPrint 240, "SetFields:" & Err.Description & vbCrLf & sql
    SetFields = Empty
    Err.Clear
End If
End Function

'
'============== CONNECTION ===========================



'==============DEBUG =====================================

'Function DBG(Optional s$)
'DebugPrint 0, "DBG " & Now() & ": " & s
'End Function

Public Sub DbgWnd(ByVal h)
Dim x As xControl: Set x = xControl(h)
m_DbgWnd = x.hWnd: If m_DbgWnd Then x.RichEditEventMask 0
DebugPrint 0, vbTab
xa.gDebugAll = GetShiftState
End Sub

Public Sub DebugPrint(ByVal tabc, ParamArray s())
On Error Resume Next
'Static tm#
Dim dx$, sar, c&, t&
sar = s


c = tabc ' RGB  RED = 0000 [tab:000][bold:0]
If c < 0 Then t = -1: c = Abs(c) Else t = (c And 14) \ 2
c = c And &HFFFFF1

If t > 0 Then dx = dx & String(t, vbTab)
dx = dx & Replace(modMain.Join2(sar, vbTab), Chr(0), "")

'dx = dx & vbTab & vbTab & Round(1000 * (Timer - tm), 3) & "ms"
'tm = Timer

If InStr(dx, vbCrLf) Then dx = Join2(Split(dx, vbCrLf), vbCrLf & String(t, vbTab))

gDebugPrint dx, t <> -1

If m_DbgWnd = 0 Then Exit Sub

With ITextDocument(m_DbgWnd).Selection.Duplicate
    If .StoryLength > 32768 Or dx = vbTab Then .SetRange 0, IIf(dx = vbTab, .StoryLength, 16384): .Text = ""
    .start = .StoryLength
'    If dx <> vbTab Then .Text = dx & vbCrLf: .Font.ForeColor = c
    If dx <> vbTab Then .Text = IIf(t = -1, "", vbCrLf) & dx: .Font.ForeColor = c: .Font.Bold = IIf(c Mod 2, -1, 0)
    .Collapse 0
    Dim si As SCROLLINFO: si.cbSize = Len(si): si.fMask = SIF_RANGE Or SIF_POS Or SIF_PAGE: GetScrollInfo m_DbgWnd, 1, si
    If si.npos > si.nMax - 2 * si.nPage Or Len(dx) > 1000 Then .Select
End With

End Sub
'==============DEBUG =====================================


'============= SYNTAX ============================
Public Sub Syntax(key, v)
If cpRegEx Is Nothing Then Set cpRegEx = New CParam  ': cpRegEx.Name = "RegEx"
cpRegEx(key) = v
End Sub

Public Function ApplySyntax(ByVal lang$, ByVal hWnd&, Optional ByVal Syn, Optional ByVal funspat) 'Подсветить синтакс после изменений в тексте
'funspat = function pattern => return matches matrix (0$+1$+2$, pos_in_text)
'Dim hWnd&: hWnd = L_(h): If hWnd = 0 Then Exit Sub
If hWnd = 0 Or Len(lang) = 0 Then Exit Function
'If mRegEx Is Nothing Then Set mRegEx = RegEx
'Debug.Print "START VBSyntax"

On Error Resume Next
Dim doc As ITextDocument: Set doc = ITextDocument(hWnd): If doc Is Nothing Then Exit Function

'Dim tm&: tm = GetTickCount

Dim i&, n&, b As Boolean, s$, pat
Dim old_EVENTMASK&

Dim tr As ITextRange, tfd As ITextFont, f$

'name= [italic(BOOL):0][bold(BOOL):0][backcolor(HEX):000000][size(NUM):00][forecolor(HEX):0000000]
doc.Undo tomSuspend '-9999995 'tomSuspend
old_EVENTMASK = CallOldWindowProc(hWnd, EM_GETEVENTMASK, 0, 0)
CallOldWindowProc hWnd, EM_SETEVENTMASK, 0, 0
CallOldWindowProc hWnd, WM_SETREDRAW, 0&, 0& 'WM_SETREDRAW

'sp = SelScrollPos(hWnd)
'doc.Freeze

s = Replace(doc.Range(0, doc.Selection.StoryLength - 1).Text, vbCr, vbLf)

Dim m As Match, ms 'colMatches As MatchCollection
Dim rspc&, nx& ' FLAG: >0 = Replace Match  with spaces from textstring, =1 NESTED
Dim nxf&
Dim ar, nm$, plist$
If TypeName(Syn) = "CParam" Then
    ar = Syn.Param(lang) 'Получаем и заливаем указанный синтакс
Else
    ar = cpRegEx.Param(lang) 'Получаем и заливаем указанный синтакс
End If
If Not IsArray(ar) Then ar = Array()

Dim tfds As New Collection, pats$, fmt$, fm, p$, bb As Boolean

For i = 0 To UBound(ar) Step 2
    pat = ar(i + 1) 'pattern or array(format:pattern,..,)
    fmt = ar(i)
    If InStr(fmt, "@") Then 'pattern for generate procedures list
        plist = ar(i + 1)
    Else 'nm next
        'pat = ar(i + 1)
        Set tr = doc.Range(0, 0).Duplicate
        If aVal(pat) > 0 Then 'Это номер параграфа который надо подкрасить
            n = pat
            tr.MoveStart tomParagraph, n - 1: tr.Expand tomParagraph 'Номер линии
            rspc = "&H" & Right$(fmt, 6): tr.Font.ForeColor = rspc 'ForeColor

        ElseIf VarType(pat) > 1 Then
            
            bb = False
            Do
                If IsArray(pat) Then
                    If Not bb Then
                        pats = "": fm = ""
                        For n = 0 To UBound(pat) Step 2
                            If Len(fm) Then fm = fm & ","
                            If Not b Or InStr(pat(n + 1), Chr(1) & Chr(1)) Then pats = NzS(pats, , , "|") & "(" & pat(n + 1) & ")": fm = fm & n
                            'If Not b Or InStr(pat(n + 1), Chr(1) & Chr(1)) Then pats = NzS(pats, , , "|") & "" & pat(n + 1) & "": fm = fm & n
                        Next
                        fm = Split(fm, ",")
                        If b Then bb = True 'nested 2
                    End If
                    b = False
                Else
                    pats = pat: p = pat ': fmt = ar(i)
                    GoSub getstyle 'fmt
                End If
                
                'b = False
                Set ms = RegExMatch(s, Replace(pats, Chr(1), ""), -1)
                'DebugPrint 4, "time SX " & lang & "+" & (GetTickCount - tm) & "ms =" & ms.count, pats
                For Each m In ms
                    If m.length Then
                    
                        tr.SetRange m.FirstIndex, m.FirstIndex + m.length 'SELECT RANGE
                        If IsArray(pat) Then
                            For n = 0 To m.SubMatches.count - 1 'find format by submatch index
                                If Not IsEmpty(m.SubMatches(n)) Then nx = Val(fm(n)): fmt = pat(nx): p = pat(nx + 1): Exit For
                            Next
'    fmt = "0": p = ""
                            b = b Or InStr(p, Chr(1) & Chr(1))  'NESTED pattern
                            GoSub getstyle 'fmt
                        End If
                        'b = b Or InStr(p, Chr(1) & Chr(1))  'NESTED pattern
                        'GoSub getstyle 'fmt
                        If InStr(p, Chr(1)) Then 'Replace by space
                            Mid$(s, m.FirstIndex + 1, m.length) = String$(m.length, 32)
'CORRECT CAPS <<< ==========================
                        ElseIf nxf And 1& Then 'SET ForeColor =EXIST
                            If tfd.ForeColor And 1& Then 'Text CAPS  ForeColor and 1& =true
                                'f = Mid(pat, InStr(1, pat, m.Value , vbTextCompare) , m.length)
                                f = Mid(pat, InStr(1, pat, "|" & m.Value & "|", vbTextCompare) + 1, m.length)
                                If Len(f) = m.length Then
                                    If StrComp(f, m.Value, vbTextCompare) = 0 And StrComp(f, m.Value, vbBinaryCompare) <> 0 Then tr.Text = f 'DebugPrint 240, m.Value & ">>" & f
                                End If
                            End If
'CORRECT CAPS >>> ==========================
                        End If
'                        GoSub setstyle
'SET STYLE >>> ==========================
    With tr.Font
        If nxf And 1 Then .ForeColor = tfd.ForeColor
        If nxf And 2 Then .BackColor = tfd.BackColor
        If nxf And 4 Then .Size = tfd.Size
        If nxf And 8 Then .Bold = tfd.Bold
        If nxf And 16 Then .Italic = tfd.Italic
        If nxf And 32 Then .Underline = tfd.Underline
        If nxf And 64 Then .Name = tfd.Name
    End With
'SET STYLE >>> ==========================
                    End If
                Next
                
                
            Loop While b


'DebugPrint 6, "END time SX " & lang & "+" & (GetTickCount - tm) & "ms =" & ms.count, pats
        End If 'pat next
    End If 'nm next
Next

'DebugPrint 240, "END1 time SX " & lang & "+" & (GetTickCount - tm) & "ms =" & ms.count, pats
Set tfds = New Collection
'DebugPrint 240, "END2 time SX " & lang & "+" & (GetTickCount - tm) & "ms =" & ms.count, pats
'DebugPrint 3, "color " & lang & "=" & (GetTickCount - tm)

funspat = S_(funspat)
If Len(funspat) Then 'LIST PROCEDURES
    If Len(plist) = 0 Then plist = funspat
    nm = "s$" & vbTab & "i&" & vbTab & "e&"
    ar = Split(plist & vbCrLf, vbCrLf)
    plist = ar(0) 'proc list pattern
    nx = Val(ar(1)) ' jscript body { }
'SaveToFile App.Path & "\zzzzz_syn_test." & IIf(lang = "vbscript", "vbs", "js"), plist & vbCrLf & s
'ar = Array(vbTab, " ", ", ", ",", " ,", ",", "( ", "(", " )", ")")
ar = Split(vbTab & "= =, =,= ,=,=( =(= )=)=  = =" & vbLf & "==,_=,=(_=(=_)=)", "=")
'DebugPrint 240, "START PLIST time SX " & lang & "+" & (GetTickCount - tm) & "ms =" & ms.count, pats

Set ms = RegExMatch(s, plist, -1, mRegEx)
    For Each m In ms
    

        f = "": f = m.SubMatches(0)
        If Len(f) = 0 Then f = m.Value
        nxf = m.FirstIndex + InStr(m.Value, f) - 1 'start pos
        
        'f = Replace(f, "\", "\\")
'        While InStr(f, ", "): f = Replace(f, ", ", ","): Wend
'        While InStr(f, " ,"): f = Replace(f, " ,", ","): Wend
'        While InStr(f, "( "): f = Replace(f, "( ", "("): Wend
'        While InStr(f, " )"): f = Replace(f, " )", ")"): Wend
        Do: plist = f: f = Replace2(f, ar): Loop While f <> plist
        If InStr(f, "\") Then f = JSONEscape(f, , mRegEx)
        'Debug.Print "<<" & f & ">>"
        'Debug.Print "<<" & m.Value & ">>"
        rspc = m.length: If nx Then rspc = rspc + Len(GTag(s, 1, "{", "}", m.FirstIndex + 1)) + 2
        nm = nm & vbCrLf & f & vbTab & nxf & vbTab & m.FirstIndex + rspc
    Next
    ApplySyntax = nm
End If
'SelScrollPos(hWnd) = sp
'If Err Then Debug.Print "ERROR " & Err.Description: Err.Clear

doc.Undo tomResume
'doc.Unfreeze
CallOldWindowProc hWnd, EM_SETEVENTMASK, 0, old_EVENTMASK
CallOldWindowProc hWnd, WM_SETREDRAW, 1&, 0& 'WM_SETREDRAW
RedrawWindow hWnd, 0, 0, 1
doc.Selection.Select
Set doc = Nothing
'speed = (speed + 1) Mod 2
'DebugPrint -1, "=" & (GetTickCount - tm) & "ms"
Exit Function



getstyle:


    f = String$(17, " ") & fmt
'
'f = Right(Replace(f, " ", Chr(0)), 42)
'Dim sf As synfmt
'CopyMemory sf, StrPtr(f), 42

    
nm = Replace(LTrim(f), " ", "_")
Set tfd = Nothing: Set tfd = tfds(nm)
If tfd Is Nothing Then

    nxf = 0
    Set tfd = tr.Font.Duplicate
    With tfd
        .Reset tomUndefined
'        If speed Then
n = Len(f) + 1
nx = sNzS(sNzS(Trim(Mid(f, n - 6, 6)), , "&H"), -1): If nx > -1 Then .ForeColor = nx: nxf = nxf Or 1
nx = sNzS(sNzS(Trim(Mid(f, n - 12, 6)), , "&H"), -1): If nx > -1 Then .BackColor = nx: nxf = nxf Or 2
nx = sNzS(sNzS(Trim(Mid(f, n - 13, 1)), , "&H"), -1): If nx > -1 Then .Size = nx: nxf = nxf Or 4
nx = sNzS(Trim(Mid(f, n - 14, 1)), -1): If nx > -1 Then .Bold = nx > 0: nxf = nxf Or 8
nx = sNzS(Trim(Mid(f, n - 15, 1)), -1): If nx > -1 Then .Italic = nx > 0: nxf = nxf Or 16
nx = sNzS(Trim(Mid(f, n - 16, 1)), -1): If nx > -1 Then .Underline = nx > 0: nxf = nxf Or 32
'Else
'        nx = NzS(NzS(Trim$(Right$(f, 6)), , "&H"), -1): If nx > -1 Then .ForeColor = nx: nxf = nxf Or 1
'        nx = NzS(NzS(Trim$(Left$(Right$(f, 12), 6)), , "&H"), -1): If nx > -1 Then .BackColor = nx: nxf = nxf Or 2
'        nx = NzS(Trim$(Left$(Right$(f, 13), 1)), -1): If nx > -1 Then .SIZE = nx: nxf = nxf Or 4
'        nx = NzS(Trim$(Left$(Right$(f, 14), 1)), -1): If nx > -1 Then .bold = nx > 0: nxf = nxf Or 8
'        nx = NzS(Trim$(Left$(Right$(f, 15), 1)), -1): If nx > -1 Then .italic = nx > 0: nxf = nxf Or 16
'        nx = NzS(Trim$(Left$(Right$(f, 16), 1)), -1): If nx > -1 Then .Underline = nx > 0: nxf = nxf Or 32
'End If
        f = Trim(Left$(f, Len(f) - 16)): If Len(f) Then .Name = f: nxf = nxf Or 64
    End With
    
    tfds.Add tfd, nm
    tfds.Add nxf, nm & "x"
Else
    nxf = tfds(nm & "x")
End If

    Return
'
'setstyle:
'    With tr.Font
'        If nxf And 1 Then .ForeColor = tfd.ForeColor
'        If nxf And 2 Then .BackColor = tfd.BackColor
'        If nxf And 4 Then .SIZE = tfd.SIZE
'        If nxf And 8 Then .Bold = tfd.Bold
'        If nxf And 16 Then .Italic = tfd.Italic
'        If nxf And 32 Then .Underline = tfd.Underline
'        If nxf And 64 Then .Name = tfd.Name
'    End With
'    Return
End Function




'Public Property Get SelScrollPos(ByVal h)
'Dim hWnd&, pos As CHARRANGE
'hWnd = h
'If hWnd = 0 Then Exit Property
'CallOldWindowProc hWnd, EM_EXGETSEL, 0, VarPtr(pos) 'EM_EXGETSEL
'SelScrollPos = Join(Array(pos.cpMin, pos.cpMax, GetScrollPos(hWnd, 0), GetScrollPos(hWnd, 1)), "x")
'End Property
'Public Property Let SelScrollPos(ByVal h, ByVal v)
'Dim hWnd&, pos As CHARRANGE, ar
'hWnd = h
'If VarType(v) < 2 Or hWnd = 0 Then Exit Property
'ar = Split(v & "xxx", "x")
'aValAr ar
'If GetScrollPos(hWnd, 0) <> CLng(ar(2)) Then CallOldWindowProc hWnd, &H114, MakeDWord(4, CLng(ar(2))), 0 'WM_HSCROLL
'If GetScrollPos(hWnd, 1) <> CLng(ar(3)) Then CallOldWindowProc hWnd, &H115, MakeDWord(4, CLng(ar(3))), 0 'WM_VSCROLL
'pos.cpMin = ar(0): pos.cpMax = ar(1)
'CallOldWindowProc hWnd, EM_EXSETSEL, 0, VarPtr(pos) 'EM_EXSETSEL
'End Property

'============= SYNTAX ============================

Function GetFileName$(ByVal hWndOwner, ByVal inPath, Optional ByVal inFilter = "*.*", Optional ByVal ShowSaveAs As Boolean)
Dim hWnd&, Path$, sFilter$, bSave As Boolean
hWnd = L_(hWndOwner): Path = S_(inPath): sFilter = Nz(inFilter, "*.*"): bSave = modMain.Cast(ShowSaveAs, vbBoolean)
GetFileName = mGetFileName(hWnd&, ByVal Path$, ByVal sFilter, bSave)
End Function
Function BrowseFolder$(ByVal hWndOwner, ByVal sTitle, ByVal inPath, Optional nFlags = &H5, Optional sRootDir$)
'BIF_RETURNONLYFSDIRS (0x00000001)
'BIF_STATUSTEXT (0x00000004)
'BIF_NEWDIALOGSTYLE (0x00000040) SIZED BORDER
'BIF_BROWSEINCLUDEURLS (0x00000080)
'BIF_BROWSEINCLUDEFILES (0x00004000)
BrowseFolder = mBrowseForFolder(L_(hWndOwner), S_(sTitle), S_(inPath), L_(nFlags), sRootDir)
End Function
'Public Function HSL(ByVal Hue As Long, ByVal Saturation As Long, ByVal Luminance As Long) As Long
'HSL = HSL2RGB(Hue, Saturation, Luminance, True)
'End Function
Function GetColor(ByVal hWndOwner, ByVal InColor)
Dim color&: color = L_(InColor)
If color > 0 And color < 32 Then color = GetSysColor(color)
GetColor = mGetColor(L_(hWndOwner), color)
End Function
Function GetFont(ByVal hWndOwner, ByVal fontText)
GetFont = mGetFont(L_(hWndOwner), S_(fontText))
End Function

'============= INDEX ARRAY ============================
Public Function AddIndex(ByRef ar, ByVal IndexItem, Optional ByVal retAR = 0)  'RETURN INDEX OF  ADDED ITEM
AddIndex = gAddIndex(ar, IndexItem, L_(retAR))
End Function
Public Function RemoveIndex(ByRef ar, ByVal IndexItem) 'RETURN ARRAY
RemoveIndex = gRemoveIndex(ar, IndexItem)
End Function
Public Function FindIndex(ByRef ar, ByVal IndexItem, Optional fOptions) As Long 'RETURN INDEX OF  ITEM
FindIndex = gFindIndex(ar, IndexItem, L_(fOptions))
End Function
'============= INDEX ARRAY ============================


Function PingAddress(ByVal HostName) As Long
If VarType(HostName) = vbString Then PingAddress = mPing(pIPAddress(HostName, 1))
End Function
Function PingTCP(HostNamePort, Optional nTimeOut = 100) As Long
PingTCP = pPingTCP(S_(HostNamePort), L_(nTimeOut))
End Function

Function IPAddress$(HostName, Optional ByVal retLong As Long)
IPAddress = pIPAddress(HostName, retLong)
End Function
Function MACAddress$(ByVal IPAddress)
MACAddress = pMACAddress(vbNullString & IPAddress)
End Function

Function IPLongFromName&(ByVal HostName)
IPLongFromName = 0
IPLongFromName = pIPAddress(HostName, -1)
End Function
Function IPNameFromLong$(sin_addr)
Dim a1&
a1 = inet_ntoa(0& + sin_addr)
IPNameFromLong = StringFromPointer(a1)
End Function
Function IPCompNameFromLong$(sin_addr)
Dim a1&, a2$
a1 = inet_ntoa(0& + sin_addr)
a2 = StringFromPointer(a1)
IPAddress a2
IPCompNameFromLong = a2
End Function

Function Compress$(s)
Compress = gCompress(s & vbNullString)
'Debug.Print Len(Compress)
End Function
Function isCompressed(s) As Boolean
isCompressed = gIsCompressed(Cast(s, vbString) & vbNullString)
End Function
Function Decompress$(s)
Decompress = gDecompress(s)
End Function

Function sDecompress$(s)
If gIsCompressed(s & vbNullString) Then sDecompress = gDecompress(s) Else sDecompress = s
End Function

Function IsFileCompressed(Url) As Boolean
On Error Resume Next
IsFileCompressed = LenFile(Url) = StringToLong(ReadFile(Url, 4))
End Function

'перекодировка строки win-1251 в кодировку utf-8 и обратно
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToUTF8(ByRef sText, Optional ByVal getPercentEncode)
Dim lngOutLen As Long
On Error Resume Next
ToUTF8 = ""
If VarType(sText) <> vbString Then Exit Function
Dim Utf8() As Byte
lngOutLen = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
ReDim Utf8(lngOutLen - 1)
WideCharToMultiByte CP_UTF8, 0, StrPtr(sText), Len(sText), VarPtr(Utf8(0)), lngOutLen, 0, 0
If L_(getPercentEncode) Then
    Dim i&
    Dim st As New Stream: st.charset = "Windows-1251": st.Type = adTypeText: st.mode = adModeReadWrite: st.Open
    For i = 0 To UBound(Utf8)
        st.WriteText "%" & Right$("0" & Hex$(Utf8(i)), 2)
    Next i
    st.Position = 0
    ToUTF8 = st.ReadText
Else
    ToUTF8 = Utf8
End If
End Function

Public Function FromUTF8(ByRef sText) As String
Dim Utf8() As Byte
On Error Resume Next
If IsArray(sText) Then Utf8 = sText Else Utf8 = S_(sText)
Dim lngOutLen As Long
If UBound(Utf8) = -1 Then Exit Function
lngOutLen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf8(0)), UBound(Utf8) + 1, 0, 0)
If lngOutLen = 0 Then
    'Err.Raise vbObjectError Or &HC312&, "FromUTF8", "Failed to decode string, system error " & CStr(Err.LastDllError)
Else
    FromUTF8 = String$(lngOutLen, 0)
    lngOutLen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf8(0)), UBound(Utf8) + 1, StrPtr(FromUTF8), lngOutLen)
    If lngOutLen = 0 Then
        'Err.Raise vbObjectError Or &HC312&, "FromUTF8", "Failed to decode string, system error " & CStr(Err.LastDllError)
        FromUTF8 = vbNullString
    End If
End If
End Function

Function ToString(vVariant, Optional ByVal bReturnArray)
ToString = modMain.ToString(vVariant, Cast(bReturnArray, 3))
End Function
Function ToVariant(sString)
ToVariant = modMain.ToVariant(S_(sString))
End Function

'Function MD5File$(ByVal sFile)
'On Error Resume Next
'MD5File = CalcMD5(vbNullString & modMain.Cast(sFile, vbString))
'End Function
'Function MD5Value$(src)
'If VarType(src) And vbArray Then MD5Value = CalcMD5_String(ToString(src)) Else If Len(Nz(src)) Then MD5Value = CalcMD5_String(Nz(src))
'End Function

'Function MD5(src) As String
'Dim strBuffer As String
'
'If VarType(src) And vbArray Then
'    'MD5Value = CalcMD5_String(ToString(src))
'    strBuffer = ToString(src)
'
'Else
'    'If Len(Nz(src)) Then MD5Value = CalcMD5_String(Nz(src))
'    strBuffer = Nz(src)
'End If
'
'
'
'    Dim myContext As MD5_CTX
'    Dim result As String
'    Dim lp As Long
'    MD5Init myContext
'    MD5Update myContext, strBuffer, Len(strBuffer)
'    MD5Final myContext
'    result = StrConv(myContext.digest, vbUnicode)
'    For lp = 1 To Len(result)
'        MD5 = MD5 & Right("00" & Hex(Asc(Mid(result, lp, 1))), 2)
'    Next
'
'End Function

Function MD5(src) As String
Dim s As String: If VarType(src) And vbArray Then s = ToString(src) Else s = S_(src)
MD5 = Hash(s, &H8003&)
End Function


''Function ToBase64(Data) As String  'base64encode(data)
'''On Error Resume Next
''With CreateObject("CDO.Message").BodyPart
''    '.Charset = "utf-8"
''    .charset = "windows-1251"
''    .ContentTransferEncoding = "base64"
''    With .GetDecodedContentStream
''        .WriteText Data
''        .flush
''    End With
''    With .GetEncodedContentStream
''        .charset = "windows-1251"
''        ToBase64 = Replace(.ReadText(.SIZE - 2), vbCrLf, "") 'Отрезаем перевод строки на хвосте
''    End With
''End With
''End Function
'Public Function ToBase64(ByVal s) As String
'Dim iXml As New MSXML2.DOMDocument30
'Dim b() As Byte
'With iXml.createElement("x")
'    .dataType = "bin.base64"
'    b = StrConv(s, vbFromUnicode)
'    .nodeTypedValue = b
'    ToBase64 = .Text
'End With
'End Function
Public Function ToBase64(ByVal s) As String '=CAST(s,82)
Dim n As Long, b() As Byte
If VarType(s) = vbArray + vbByte Then b = s Else b = StrConv(s, vbFromUnicode)
If CryptBinaryToString(b(0), UBound(b) + 1, 1, 0, n) Then    'Convert binary to Base64.
    ToBase64 = String$(n - 1, 0)
    If CryptBinaryToString(b(0), UBound(b) + 1, 1, StrPtr(ToBase64), n) Then ToBase64 = Replace(ToBase64, vbCrLf, "")
End If
'ToBase64 = modMain.Cast(s, 82)
End Function

''Function FromBase64(Data) As String  'base64decode(data)
'''On Error Resume Next
''With CreateObject("CDO.Message").BodyPart
''    .charset = "windows-1251"
''    .ContentTransferEncoding = "base64"
''    With .GetEncodedContentStream
''        .WriteText Data
''        .flush
''    End With
''    With .GetDecodedContentStream
''    .charset = "windows-1251"
'''            .Charset = "utf-8"
''        FromBase64 = .ReadText
''    End With
''End With
''End Function
'Public Function FromBase64(ByVal s) As String
'Dim iXml As New MSXML2.DOMDocument30
'Dim b() As Byte
'With iXml.createElement("x")
'    .dataType = "bin.base64"
'    .Text = s
'    b = .nodeTypedValue
'    FromBase64 = StrConv(b, vbUnicode)
'End With
'End Function

Public Function FromBase64(ByVal s) '=CAST(s,81)
Dim n As Long, b() As Byte, dwFlags As Long, s0 As String: s0 = s
If CryptStringToBinary(StrPtr(s0), Len(s0), 1, 0, n, 0&, dwFlags) Then   'Convert Base64 to binary.
    ReDim b(n - 1)
    If CryptStringToBinary(StrPtr(s0), Len(s0), 1, VarPtr(b(0)), n, 0&, dwFlags) Then FromBase64 = StrConv(b, vbUnicode)
End If
'FromBase64 = modMain.Cast(s, 81)
End Function

'Function LunaCheck(ByVal num$, Optional ByVal nLen = 16)
'    Dim i&, sum&, p&
'    LunaCheck = IF0(Abs(aVal(nLen)), 1)
'    num = Replace(num, " ", "")
'    num = RegExReplace(num, "[^0-9]", Fix(Rnd * 10))
'    If Len(num) = LunaCheck Then
'        For i = 1 To Len(num)
'            p = Val(Mid$(num, i, 1)) * (i Mod 2 + 1)
'            sum = sum + IIf(p > 9, p - 9, p)
'        Next i
'        LunaCheck = sum Mod 10
'    End If
'End Function

Function WindowClass(ByVal hWnd) As String
WindowClass = pWindowClass(Cast(hWnd, vbLong))
End Function
Function WindowText(ByVal hWnd) As String
WindowText = pWindowText(Cast(hWnd, vbLong))
End Function


Function WaitProcessWindow(ByVal sClass, ByVal sName, Optional ByVal nTimeOut = 5000, Optional ByVal arExclude)
Dim ar0, ar, tm&, s$
sClass = S_(sClass)
sName = S_(sName)
If IsArray(arExclude) Or Len(S_(arExclude)) > 0 Then
    If Not IsArray(arExclude) Then arExclude = Split(arExclude, ",")
    arExclude = CastAR(arExclude, vbLong)
    tm = GetTickCount
    nTimeOut = Cast(nTimeOut, vbLong)
    Do
        DoEvents
        ar = EnumProcessWindows(sClass, sName)
        gRemoveIndex ar, arExclude
    Loop While ((GetTickCount - tm) < nTimeOut) And (UBound(ar) < 0)
    If UBound(ar) > -1 Then WaitProcessWindow = ar(0)

Else
    WaitProcessWindow = modMain.Join2(EnumProcessWindows(sClass, sName), ",")
End If
End Function


'Function OLECmdExec(pHTML, Optional ByVal template)
Function OLECmdExec(pOle, ByVal vCmd, ByVal vOpt, ByVal pvIn, Optional pvOut)
OLECmdExec = ""
On Error Resume Next

Dim nCmd&, nOpt&, pv
nCmd = L_(vCmd)
If nCmd = 0 Then Exit Function
nOpt = L_(vOpt)
Dim ot As IOleCommandTarget
Set ot = pOle
If ot Is Nothing Then Exit Function

Select Case TypeName(pOle)
Case "HTMLDocument"
    Dim s$, tm&: tm = Timer    'Debug.Print html.readystate
    Do
        DoEvents
        s = pOle.readyState
    Loop While s <> "complete" And (Timer - tm) < 5000
End Select

ot.Exec ByVal 0, nCmd, nOpt, pvIn, pv
If Not IsMissing(pvOut) Then pvOut = pv
DoEvents

End Function

Function WindowMove(ByVal h, Optional ByVal x = -999, Optional ByVal Y = -999, Optional ByVal cx = -999, Optional ByVal cy = -999, Optional ByVal nShow = -999)
Dim hWnd&
hWnd = Cast(h, vbLong)
Dim wp As WINDOWPOS, f&
GetWindowPos hWnd, wp
f = 7
If x <> -999 Then If wp.x <> x Then wp.x = x: f = f And 5&
If Y <> -999 Then If wp.Y <> Y Then wp.Y = Y: f = f And 5&
If cx <> -999 Then If wp.cx <> cx Then wp.cx = cx: f = f And 6&
If cy <> -999 Then If wp.cy <> cy Then wp.cy = cy: f = f And 6&
If Not ((f And 3&) = 3&) Then SetWindowPos hWnd, 0, wp.x, wp.Y, wp.cx, wp.cy, f Or &H224 Or SWP_NOACTIVATE
If nShow <> -999 Then ShowWindow hWnd, Cast(nShow, vbLong)
End Function

Function WindowGet(ByVal h, ByVal uCmd)
WindowGet = GetWindow(Cast(h, vbLong), Cast(uCmd, vbLong))
End Function
'Sub WindowFocus(ByVal h)
'apiSetFocus aVal(h)
'End Sub
Property Get WindowLong(ByVal hWnd, ByVal nIndex, Optional ByVal mask = -1)
WindowLong = GetWindowLong(modMain.Cast(hWnd, vbLong), modMain.Cast(nIndex, vbLong)) And modMain.Cast(mask, vbLong)
End Property

Property Let WindowLong(ByVal hWnd, ByVal nIndex, Optional ByVal mask = -1, ByVal newLong)
'Public Const GWL_STYLE As Long = -16
'Public Const GWL_EXSTYLE As Long = -20
'Public Const GWL_WNDPROC As Long = -4
Dim h&, i&, m&
h = modMain.Cast(hWnd, vbLong)
i = modMain.Cast(nIndex, vbLong)
m = modMain.Cast(mask, vbLong)
Dim wl&, nl&
nl = modMain.Cast(newLong, vbLong)
wl = GetWindowLong(h, i)
SetWindowLong h, i, (wl And Not mask) Or (nl And mask)
End Property

Function DWord(x, Y)
DWord = MakeDWord(modMain.Cast(x, 2), modMain.Cast(Y, 2))
End Function
Function DWordEx(dw, Optional x = 0, Optional Y = 0)
DWordEx = MakeDWord(LOWORD(Cast(dw, 3)) + Cast(x, 2), HIWORD(Cast(dw, 3)) + Cast(Y, 2))
End Function

Function MousePos(Optional hWnd) As Long
Dim pt As POINTAPI
GetCursorPos pt
If L_(hWnd) Then ScreenToClient L_(hWnd), pt
MousePos = MakeDWord(pt.x, pt.Y)
End Function

Sub ShowToolTip(ByVal hWnd, ByVal Text$, Optional ByVal npos)
ToolTipHide
Dim rc As RECT, pt As POINTAPI, pos&

pos = L_(npos)
If pos = -1 Then  'CurrenCursorPos
    GetCursorPos pt: pt.x = pt.x + 16: pt.Y = pt.Y + 16
Else
    GetWindowRect 0& + hWnd, rc
    If pos = 0 Then 'Center CLIENT
        pt.x = (rc.Left + rc.Right) \ 2
        pt.Y = (rc.Top + rc.Bottom) \ 2
    Else 'ScreenPos
        pt.x = LOWORD(pos)
        pt.Y = HIWORD(pos)
        If L_(hWnd) Then 'Ofset hWnd
            pt.x = pt.x + rc.Left
            pt.Y = pt.Y + rc.Top
        End If
    End If
End If
ToolTipTrackTip pt.x, pt.Y, Text
End Sub
Sub HideTip()
ToolTipHide
End Sub
Sub ShowBalloon(ByVal hWnd, ByVal sText, ByVal sTitle, ByVal nIcon, Optional ByVal npos)
ToolTipHide
Dim rc As RECT, pt As POINTAPI
npos = L_(npos)
If npos = -1 Then 'Mouse
    GetCursorPos pt
Else
    GetWindowRect 0& + hWnd, rc
    pt.x = LOWORD(0 + npos)
    pt.Y = HIWORD(0 + npos)
    If pt.x = 0 Then pt.x = (rc.Left + rc.Right) \ 2 Else pt.x = pt.x + rc.Left
    If pt.Y = 0 Then pt.Y = (rc.Top + rc.Bottom) \ 2 Else pt.Y = pt.Y + rc.Top
End If
ToolTipTrackBalloon pt.x, pt.Y, vbNullString & sText, vbNullString & sTitle, 0& + nIcon
'aSetFocus hWnd
End Sub


Sub KillTask(pid)
Dim process As Object, n&
On Error GoTo e_x
For Each process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
n = process.handle
'Debug.Print Process.Caption & " " & Process.handle & " " & n
If n = pid Then
    process.Terminate
    Exit For
End If
Next
e_x:
End Sub

Function CurrentProcessID() As Long
CurrentProcessID = GetCurrentProcessId
End Function

'Function isProcess(pid) As Long
'isProcess = GetProcessVersion(Cast(pid, 3))  'Windows Version
'Debug.Print "GetProcessVersion"; pid; HIWORD(isProcess); LOWORD(isProcess)
'End Function


Function NewIDE(oargs)
Dim cp As New CParam
On Error Resume Next
cp.Load "*"
cp.Url = ""
'cp("SourcePath") = cp.Url
cp.AddParams oargs
Select Case cp.ParamDef("Mode", 0)
Case 1 'DoModal
    NewIDE = DoModal(cp.Value)
Case 2 'FormValue
    NewIDE = FormValue(cp.Value)
Case Else 'NewForm
    NewIDE = NewForm(vbNullString, , cp)
End Select
End Function


Function GetTag(src, n, stg, etg) As String
'Функция = взять содержимое n-ого тага  в границах [<stg] [etg/>]
'Функция = (n<0) посчитать кол-во тагов  в границах [<stg] [etg/>]
GetTag = GTag(vbNullString & src, 0 + n, vbNullString & stg, vbNullString & etg)
End Function


'Function IsEqual(v1, v2) As Boolean
'Dim i&, t1&, t2& ', res As Boolean
't1 = VarType(v1): t2 = VarType(v2)
'If t1 <> t2 Then Exit Function 'And t1 > 0 And t2 > 0 Then Exit Function
'If t1 And vbArray Then
'    If LBound(v1) = LBound(v2) And UBound(v1) = UBound(v2) Then
'        IsEqual = True
'        For i = LBound(v1) To UBound(v1)
'            IsEqual = IsEqual And IsEqual(v1(i), v2(i))
'            If Not IsEqual Then Exit For
'        Next
'    End If
'ElseIf t1 < 2 Then
'    IsEqual = True
'Else
'    IsEqual = (v1 = v2)
'    'Debug.Print "IsEqual " & v1 & " = " & v2
'End If
'End Function


#If LogLock Then '==================LOGLOCK=================

Function GetArraysDiff(hdr, ar0, ar1, Optional ReturnFields As Long)
'hdr - список полей/описаний полей
'ar0 - исходные данные
'ar1 - измененные данные
'ReturnFields
'<-2 - return array of bool changes
'=-2 - return array of new values
'=-1 - return text fields with changes
'=0 -  return text log only fields with changes
'>0 - return array fields with changes

Dim i&, n&, s$, bOK As Boolean, ar, cp As New CParam
ar = Array()
If (VarType(hdr) And VarType(ar0) And VarType(ar1) And vbArray) = 0 Then Exit Function
n = UBound(hdr)
If UBound(ar0) < n Then n = UBound(ar0)
If UBound(ar1) < n Then n = UBound(ar1)
For i = 0 To n
If Len(modMain.Cast(hdr(i), vbString)) Then
    bOK = 0
    If VarType(ar0(i)) <> VarType(ar1(i)) Then bOK = 1 Else If VarType(ar0(i)) > 1 Then If ar0(i) <> ar1(i) Then bOK = 1
    If ReturnFields < -2 Then 'Return array of bool changes
        ReDim Preserve ar(UBound(ar) + 1)
        ar(UBound(ar)) = bOK
    ElseIf bOK Then
        If ReturnFields = -2 Then 'Return array of NewValues
            ReDim Preserve ar(UBound(ar) + 1)
            ar(UBound(ar)) = ar1(i)
        ElseIf ReturnFields = -1 Then 'Return fields
            s = s & IIf(Len(s), ",", vbNullString) & hdr(i)
        ElseIf ReturnFields = 0 Then 'Return TextLOG
            s = s & IIf(Len(s), vbCrLf, vbNullString) & hdr(i) & ": {" & Nz(ar0(i), "Null") & "} ->{" & Nz(ar1(i), "Null") & "}"
        Else 'If ReturnFields = 1 Then 'Return LOGAR (OldValues)
            cp.Param(hdr(i)) = Nz(ar0(i), "Null")
        End If
    End If
End If
Next
If ReturnFields < -1 Then
    GetArraysDiff = ar
ElseIf ReturnFields < 1 Then
    GetArraysDiff = s
Else
    If cp.ParamCount Then GetArraysDiff = cp.Source Else GetArraysDiff = vbNullString
End If
End Function
#End If '==================LOGLOCK=================

'=================================================
Function SendEvent(ByVal evname, ByVal rcvList, Optional ByVal nWaitTimeOut = 0) 'Послать собщение списку получателей(через запятую (IP или HostName)))
Dim a, pt
Dim r, cp As New CParam
cp.Name = evname
For Each a In Split(rcvList, ",")
    pt = Split(a & ":" & Port, ":")
    r = SocketSendData(S_(evname), L_(pt(1)), Trim$(pt(0)), 0 + nWaitTimeOut)
    'If nWaitTimeOut Then
    cp(a) = r
Next
SendEvent = cp.Source
End Function

Function FireEvent(ByVal evname)
FireEvent = FireRSEvent(evname)
'SendEvent evname, IPAddress("")
End Function
Sub WaitEvent(Optional ByVal evname = vbNullString, Optional ByVal hWnd = 0, Optional ByVal proc = vbNullString) 'Ждать событие, при получении выполнить hxControl(hWnd).hEvent proc, evtxt, RemoteHost
xEvent evname, hWnd, proc
End Sub
Function SocketReply(sSocket, txtData)
SocketReply = pSocketReply(S_(sSocket), S_(txtData))
End Function
'=================================================



'=================================================
'===========RTF REPORT FUNCTIONS ===============
Public Property Get RangeRTF(tr) 'As String
RangeRTF = RangeRTFparam(tr, CF_RTF)
End Property
Public Property Get RangeText(tr) As String
RangeText = RangeRTFparam(tr, 1)
End Property

Private Function RangeRTFparam(tr, cpformaT) As String
Dim IDO As IDataObject, fmt As FORMATETC, stm As STGMEDIUM
Dim lLen As Long
On Error Resume Next
If tr.start < tr.End Then
    tr.Copy IDO  ' Copy the text to a IDataObject_Implements
    fmt.cfFormat = cpformaT 'CF_RTF ' Set the FMT structure
    fmt.TYMED = TYMED_HGLOBAL
    IDO.GetData fmt, stm   ' Get the text
    Dim b() As Byte
    lLen = lstrlen(stm.Data) ' Get the text length
    ReDim b(lLen - 1) As Byte ' Initialize the buffer
    CopyMemory b(0), ByVal stm.Data, lLen ' Copy the text to the buffer
    RangeRTFparam = StrConv(b, vbUnicode) ' Convert the string to Unicode
    '    Debug.Print RangeRTFparam & "==="
    ReleaseStgMedium stm ' Release the data on STM
End If

End Function

Public Property Let RangeRTF(tr, rtf)
On Error Resume Next
Dim cDO As New CDataObject
Dim IDO As IDataObject
Set IDO = cDO
cDO.TextRTF = S_(rtf)
tr.Paste IDO, CF_RTF
End Property

Function RangeGetPoint(tr, ta, Optional hWnd = 0) As Long
Dim pt As POINTAPI
On Error Resume Next
tr.GetPoint 0& + ta, pt.x, pt.Y
If L_(hWnd) Then ScreenToClient L_(hWnd), pt
RangeGetPoint = MakeDWord(pt.x, pt.Y)
End Function
'


'Function Num2Word$(cNumber, Optional ByVal sWordParam = vbNullString, Optional ByVal iWordView = 0) 'Optional ByVal NumDigitsAfterDecimal = 0,
''sWordParam = "ISO4217alfa3;rc;cp1;cp2;cp5;cs1;cs2;cs5;rd;dp1;dp2;dp5;ds1;ds2;ds5;ISO4217digits"
''   rc: род целых  0 - женский, 1 - мужской, [] - средний
''   cp1,cp2, cp5 - пропись целых
''   cs1, cs2, cs5 - сокращенная пропись целых
''   rc: род дробных  0 - женский, 1 - мужской, [] - средний
''   dp1, dp2, dp5 - пропись дробных
''   ds1, ds2, ds5 - сокращенная пропись дробных
''   decview: 0..3 количество знаков дробной части
'sWordParam = NzS(sWordParam, "BYN;1;рубль;рубля;рублей;руб.;руб.;руб.;0;копейка;копейки;копеек;коп.;коп.;коп.;2")
'
''iWordView =
''0 - Одна тысяча один рубль 01 копейка +
''1 - 1001 рубль 01 копейка +
''2 - 10001,01 рублей
''3 - 10001,01 руб. (Одна тысяча один рубль 01 копейка) +
''4 - 1001 руб. 01 коп. +
''5 - 1001,01 BYN
''6 - Одна тысяча один
'
'Dim tr&(5), ns0, ns1, ns2, ts0, ts1, ts2, ts5
'Dim ezn As Integer, enm As String
'Dim sm@, sm0@, tx$, tx0$, tmp$, n&, i&, vc, enm0$
'If Not IsNumeric(cNumber) Then Exit Function
'vc = Split(sWordParam & String$(15, ";"), ";")
'On Error Resume Next
'sm0 = CCur(cNumber)
'If Err Then sm = cNumber
'enm = vc(0)
'ezn = vc(15): If ezn < 0 Or ezn > 3 Then ezn = 0
'sm0 = Round(Abs(sm0), ezn)
'tr(0) = Round((10 ^ ezn) * (sm0 - Fix(sm0))) 'Дробная часть = стотые
'sm0 = Fix(sm0): sm = sm0
'ns1 = Split(",,двадцать ,тридцать ,сорок ,пятьдесят ,шестьдесят ,семьдесят ,восемьдесят ,девяносто ", ",")
'ns2 = Split(",сто ,двести ,триста ,четыреста ,пятьсот ,шестьсот ,семьсот ,восемьсот ,девятьсот ,", ",")
'For i = 1 To 5: tr(i) = sm - Fix(sm / 1000) * 1000: sm = Fix(sm / 1000): Next
'If iWordView Then
'    If iWordView = 1 Or iWordView = 4 Then
'        If iWordView = 1 Then i = 0 Else i = 3
'        tx0 = Format$(sm0, "0")  '#,#
'        n = tr(1) - 10 * Fix(tr(1) / 10)
'        If n = 0 Or n > 4 Then tmp = vc(4 + i) Else If n = 1 Then tmp = vc(2 + i) Else tmp = vc(3 + i)
'        tx0 = tx0 & " " & tmp
'        If ezn Then
'            tx0 = tx0 & " " & Format$(tr(0), String$(ezn, "0"))
'            n = tr(0) - 10 * Fix(tr(0) / 10)
'            If n = 0 Or n > 4 Then tmp = vc(11 + i) Else If n = 1 Then tmp = vc(9 + i) Else tmp = vc(10 + i)
'            tx0 = tx0 & " " & tmp
'        End If
'    Else
'        enm0 = " " & IIf(iWordView = 2, vc(4), IIf(iWordView = 5, vc(0), vc(7)))
'        tx0 = Format$(sm0 + tr(0) / (10 ^ ezn), IIf(ezn, "0." & String$(ezn, "0"), vbNullString)) & enm0
'    End If
'End If
'If iWordView = 0 Or iWordView = 3 Or iWordView = 6 Then
'    ts0 = Split(vc(5) & "," & vc(1) & ",0,1,1,1", ",")
'    ts1 = Split(vc(9) & " ," & vc(2) & " ,тысяча ,миллион ,миллиард ,триллион ", ",")
'    ts2 = Split(vc(10) & " ," & vc(3) & " ,тысячи ,миллиона ,миллиарда ,триллиона ", ",")
'    ts5 = Split(vc(11) & " ," & vc(4) & " ,тысяч ,миллионов ,миллиардов ,триллионов ", ",")
'
'    For i = 5 To IIf(ezn = 0, 1, 0) Step -1
'        If tr(i) > 0 Or i < 2 Then
'            If Len(ts0(i)) = 0 Then tmp = ",одно ,два ," Else tmp = IIf(Val(ts0(i)), ",один ,два ,", ",одна ,две ,")
'            ns0 = Split(tmp & "три ,четыре ,пять ,шесть ,семь ,восемь ,девять ,десять ,одиннадцать ,двенадцать ,тринадцать ,четырнадцать ,пятнадцать ,шестнадцать ,семнадцать ,восемнадцать ,девятнадцать ", ",")
'            If sm0 = 0 Then ns0(0) = "ноль "
'            n = Fix(tr(i) / 100) 'кол-во сотен
'            If i > 0 Then tx = tx & ns2(n)
'            n = Fix((tr(i) - n * 100) / 10) 'кол-во десятков
'            If i > 0 Then tx = tx & ns1(n)
'            If n < 2 Then n = tr(i) - (Fix(tr(i) / 10) - n) * 10 Else n = tr(i) - Fix(tr(i) / 10) * 10
'            'кол-во единиц
'            If iWordView = 6 And i = 1 Then
'                tmp = ""
'            Else
'                If n = 0 Or n > 4 Then tmp = ts5(i) Else If n = 1 Then tmp = ts1(i) Else tmp = ts2(i)
'            End If
'            If i = 0 Then ns0(n) = Format$(tr(i), String$(ezn, "0")) & " "
'            If iWordView = 6 And i = 0 Then
'
'            Else
'                tx = tx & ns0(n) & tmp
'            End If
'        End If
'    Next
'End If
'
'tx = Trim$(UCase$(Left$(tx, 1)) & Mid$(tx, 2))
'If iWordView Then
'    Num2Word = tx0
'    If iWordView = 3 Then
'        Num2Word = Num2Word & " (" & tx & ")"
'    ElseIf iWordView = 6 Then
'        Num2Word = tx
'    End If
'Else
'    Num2Word = tx
'End If
'
'End Function


Function LastInputInfo&() 'LastINPUT in seconds
'MouseDown LastTime
'Dim plii As PLASTINPUTINFO
'plii.cbSize = Len(plii)
'GetLastInputInfo plii
'LastInputInfo = (GetTickCount() - plii.dwTime) / 1000
LastInputInfo = (GetTickCount - mLastMDWM) / 1000
End Function

Sub PressKey(Optional vKey = 7)
keybd_event &HFF And Val(vKey), 0, 0, 0
keybd_event &HFF And Val(vKey), 2, 0, 0
End Sub



'============= INI SETTINGS ============================
Public Function CPBasePath(ByRef keypath, Optional ByVal sBase)
'Return cparam path   update:BYREF keypath = file name without extention
CPBasePath = modMain.CPFilePath(keypath, , RTrim$(NzS(sBase, Base.Name)))
End Function

Public Property Get AppPath() As Variant
AppPath = xAppPath
End Property
Public Property Let AppPath(ByVal v As Variant)
xAppPath = PathFromFile(v)
End Property

Public Property Get CMDLine$()
'CMDLine = NzS(Command$, gapp) 'CommandLine
CMDLine = NzS(CommandLine, gapp)
End Property
Public Property Get CMDName$(Optional ByVal cmd = vbNullString)
CMDName = SplitIndex(SplitIndex(NzS(NzS(cmd, CMDLine), App.ProductName), ";", 0))
If Len(CMDName) Then CMDName = SplitIndex(CMDName, "/", 0)
End Property

Public Property Get INIBase()
If Len(m_iniBase) Then INIBase = m_iniBase Else INIBase = CMDName
End Property
Public Property Let INIBase(v)
If Len(S_(v)) Then m_iniBase = SplitIndex(v) Else m_iniBase = CMDName
End Property

Private Function pINIFIle(section) As String
section = NzS(section, vbNullString)
pINIFIle = xAppPath & "\" & INIBase & ".ini"
Dim ar: ar = Split(section, "*")
If UBound(ar) > 0 Then
    pINIFIle = ar(0): section = ar(1)
Else
    If InStr(section, "#") Then pINIFIle = gapp & ".ini"
    section = Replace(section, "#", "")
End If
End Function
Public Property Get INISetting(ByVal section, ByVal key)
Dim INIFile$, buf$, n&, buflen&
'section = NzS(section, vbNullString)
'INIFile = xAppPath & "\" & INIBase & ".ini"
INIFile = pINIFIle(section)
If VarType(key) < 2 Then INISetting = INIFile: Exit Property

buflen = 128

key = NzS(key, vbNullString)
Do
buflen = buflen * 2
buf = Space$(buflen)

n = GetPrivateProfileString(section, key, vbNullString, buf, buflen, INIFile)
Loop While n = buflen - 1 And n > 0

If (Len(section) = 0 Or Len(key) = 0) And n > 0 Then n = n - 1
buf = Left$(buf, n)

If Len(buf) > 4 Then
    If (Len(buf) - 1) Mod 4 = 0 Then
        If Left(buf, 1) = "*" Then
            If RegExMatch(Mid(buf, 2), "^[a-z0-9\+/]*={0,2}") = Mid(buf, 2) Then buf = FromBase64(Mid(buf, 2))
        End If
    End If
End If

INISetting = buf


End Property
Public Property Let INISetting(ByVal section, ByVal key, ByVal newValue)
Dim INIFile$, sKey$, v$
INIFile = pINIFIle(section)
sKey = NzS(key, vbNullString)
v = NzS(newValue, vbNullString)
If InStr(v, vbCr) Or InStr(v, vbLf) Or InStr(v, vbTab) Then v = "*" & ToBase64(v)
MakeDir INIFile
WritePrivateProfileString vbNullString & section, sKey, v, INIFile
End Property
'============= INI SETTINGS ============================

'============= REG SETTINGS ============================
'Public Property Get Reg2Value(key) 'Глобальные настройки
'Reg2Value = pRegKey(2, "SOFTWARE\" & App.EXEName & NzS(Base.Name, , "\"), key)
'End Property
'Public Property Let Reg2Value(key, v)
'pRegKey(2, "SOFTWARE\" & App.EXEName & NzS(Base.Name, , "\"), key) = v
'End Property
'Public Property Get Reg1Value(key) 'Настройки пользователя
'Reg1Value = pRegKey(1, "SOFTWARE\" & App.EXEName & NzS(Base.Name, , "\"), key)
'End Property
'Public Property Let Reg1Value(key, v)
'pRegKey(1, "SOFTWARE\" & App.EXEName & NzS(Base.Name, , "\"), key) = v
'End Property

Public Property Get RegSetting(sect, key) 'Настройки пользователя
RegSetting = pRegKey(1, "SOFTWARE\" & App.EXEName & NzS(sect, , "\"), key)
End Property
Public Property Let RegSetting(sect, key, v)
pRegKey(1, "SOFTWARE\" & App.EXEName & NzS(sect, , "\"), key) = v
End Property

'============= REG SETTINGS ============================
'HKEY_CLASSES_ROOT As Long = &H80000000
'HKEY_CURRENT_USER As Long = &H80000001
'HKEY_LOCAL_MACHINE As Long = &H80000002
'HKEY_USERS As Long = &H80000003
'HKEY_PERFORMANCE_DATA As Long = &H80000004
'HKEY_CURRENT_CONFIG As Long = &H80000005
'HKEY_DYN_DATA As Long = &H80000006

Public Property Get RegValue(root, sSection, sKey)  'as Value or Null if no exist
RegValue = pRegKey(L_(root), S_(sSection), S_(sKey))
End Property
Public Property Let RegValue(root, sSection, sKey, v)
pRegKey(L_(root), S_(sSection), S_(sKey)) = v
End Property
'Public Property Get RegistryValue(root, sSection, sKey)  'as Value or Null if no exist
'RegistryValue = RegKey(aVal(root), Nz(sSection), Nz(sKey))
'End Property
'Public Property Let RegistryValue(root, sSection, sKey, v)
'RegKey(aVal(root), Nz(sSection), Nz(sKey)) = v
'End Property

Public Function RegKeys(ByVal root, ByVal sSection) ' As Array of keys
RegKeys = pRegKeys(L_(root), S_(sSection))
End Function
Public Function RegValues(root, sSection) ' As Array of values
RegValues = pRegValues(L_(root), S_(sSection))
End Function
'Sub RemoveRegTree(root, section)
'If Len(section) > 23 Then
'  Dim key, keys
'  keys = RegistryValues(root, section) 'Child Values
'  For Each key In keys
'  RegistryValue(root, section, key) = Null
'  Next
'  keys = RegistryKeys(root, section) 'Child sections
'  For Each key In keys
'  RemoveRegTree root, section & "\" & key
'  Next
'  RegistryValue(root, section, "") = Null 'Remove section
'End If
'End Sub

#If ACL Then
Function DenyAccess(ByVal root&, ByVal sKey$, ByVal sUser$, Optional ByVal nAccess = 0)
'nAccess = 0 READ DENY
'nAccess = 1 REMOVE DENY
'nAccess = -1 SET DENY
'#If ACL Then
DenyAccess = pRegKeyDenyAccess(root, sKey, sUser, nAccess)
'#End If
'Return 0= ALLOW -1=DENY
End Function
#End If
'============= REG SETTINGS ============================
'======== REGISTRY =========================

'=============== TIME SYSTEM ===============
Function NTPTime(Optional timehost$ = "time.windows.com") '  "time.windows.com") 'time.nist.gov"
Dim Bytes() As Byte
Dim ulngSecs As Long
Dim dblSecs As Double
Dim src$
'src = SocketSendData(Chr(&H1B) & String(47, 0), 123, timehost, 1) 'V3
src = SocketSendData(Chr$(&HA3) & String$(47, 0), 123, timehost, 1)  'V4
Bytes = StrConv(src, vbFromUnicode)
If UBound(Bytes) < 47 Then Exit Function
Dim TransmitTimestamp As Size
CopyMemory TransmitTimestamp, Bytes(40), LenB(TransmitTimestamp)
ulngSecs = ntohl(TransmitTimestamp.cx)
dblSecs = CDbl(ulngSecs And &H7FFFFFFF)
If ulngSecs And &H80000000 Then dblSecs = dblSecs + 2147483648#
If ntohl(TransmitTimestamp.cy) And &H80000000 Then dblSecs = dblSecs + 0.5
NTPTime = CDate(2 + dblSecs / 86400)
End Function

Function UTCTime(ByVal pLocalTime) As Date
''UTCTime = modAPI.pUTCTime(pLocalTime)
'Dim system_time As SYSTEMTIME
'Dim local_file_time As FILETIME
'Dim utc_file_time As FILETIME
'' Convert it into a SYSTEMTIME.
''DateToSystemTime the_date, system_time
'VariantTimeToSystemTime ByVal CDbl(pLocalTime), system_time
'' Convert it to a FILETIME.
'SystemTimeToFileTime system_time, local_file_time
'' Convert it to a UTC time.
'LocalFileTimeToFileTime local_file_time, utc_file_time
'' Convert it to a SYSTEMTIME.
'FileTimeToSystemTime utc_file_time, system_time
'' Convert it to a Date.
''SystemTimeToDate system_time, the_date
'Dim d As Double
'SystemTimeToVariantTime system_time, d
'UTCTime = CDate(d) ' the_date
UTCTime = DateAdd("n", UTCBias, pLocalTime)

End Function

Function LocalTime(ByVal pUTCTime) As Date
''LocalTime = modAPI.pLocalTime(pUTCTime)
'Dim system_time As SYSTEMTIME
'Dim local_file_time As FILETIME
'Dim utc_file_time As FILETIME
'' Convert it into a SYSTEMTIME.
'VariantTimeToSystemTime ByVal CDbl(pUTCTime), system_time
'' Convert it to a FILETIME.
'SystemTimeToFileTime system_time, local_file_time
'' Convert it to a Local time.
'FileTimeToLocalFileTime local_file_time, utc_file_time
'' Convert it to a SYSTEMTIME.
'FileTimeToSystemTime utc_file_time, system_time
''SystemTimeToDate system_time, the_date
'Dim d As Double
'SystemTimeToVariantTime system_time, d
'LocalTime = CDate(d) ' the_date
LocalTime = DateAdd("n", -UTCBias, pUTCTime)
End Function

Function SetDateTime(ByVal dtm) As Date
If VarType(dtm) <> vbDate Then Exit Function
If dtm < 40000 Then Exit Function
Date = Fix(dtm)
If dtm - Fix(dtm) > 0 Then Time = CDate(dtm - Fix(dtm))
SetDateTime = Now
End Function
'=============== TIME SYSTEM ===============

'============= FILE TIMES ============================
Public Property Get FileTimes(ByVal sFile, ByVal CreateAccessModify) 'Const GENERIC_READ = &H80000000
Dim hFile&, res As Date
Const FILE_FLAG_BACKUP_SEMANTICS As Long = &H2000000 'ДЛЯ КАТАЛОГОВ
Const FILE_READ_ATTRIBUTES As Long = &H80
Const OPEN_EXISTING As Long = 3
hFile = CreateFile(sFile & vbNullString, 0&, 0&, ByVal 0&, OPEN_EXISTING, FILE_READ_ATTRIBUTES, 0&)
If hFile = -1 Then hFile = CreateFile(sFile & vbNullString, 0&, 0&, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS Or FILE_READ_ATTRIBUTES, 0&)    'DIRECTORY
Dim ftCreated As FILETIME, ftAccessed As FILETIME, ftModified As FILETIME
If hFile <> -1 Then
Call GetFileTime(hFile, ftCreated, ftAccessed, ftModified)
If CreateAccessModify And 1& Then
    FileTimeToVBDate ftCreated, res
ElseIf CreateAccessModify And 2& Then
    FileTimeToVBDate ftAccessed, res
ElseIf CreateAccessModify And 4& Then
    FileTimeToVBDate ftModified, res
Else
    res = 0
End If
CloseHandle hFile
End If
FileTimes = res
End Property

Public Property Let FileTimes(ByVal sFile, ByVal CreateAccessModify, ByVal NewVBDate)  'Const GENERIC_WRITE = &H40000000
If VarType(NewVBDate) And vbArray Then Exit Property
If VarType(NewVBDate) < 2 Then Exit Property
NewVBDate = CDate(NewVBDate)
Dim hFile&
On Error Resume Next
hFile = CreateFile(vbNullString & sFile, &H40000000, 3&, ByVal 0&, 3&, &H10000080, 0&)
Dim ftCreated As FILETIME, ftAccessed As FILETIME, ftModified As FILETIME
If hFile <> -1 Then
Call GetFileTime(hFile, ftCreated, ftAccessed, ftModified)
If CreateAccessModify And 1& Then VBDateToFileTime NewVBDate, ftCreated
If CreateAccessModify And 2& Then VBDateToFileTime NewVBDate, ftAccessed
If CreateAccessModify And 4& Then VBDateToFileTime NewVBDate, ftModified
Call SetFileTime(hFile, ftCreated, ftAccessed, ftModified)
CloseHandle hFile
'Else
'MsgBox "Let FileTime", 16
End If '
End Property
'============= FILE TIMES ============================


'=============== FILE SYSTEM ===============
Function MakeDir(ByVal pth, Optional bCreate = 1)
MakeDir = modMain.MakeDir(pth, L_(bCreate))
End Function

Function KillFile(ByVal pth) As String
On Error Resume Next:: Kill pth
KillFile = pth
End Function

Function LenFile&(ByVal pth)
On Error Resume Next
'LenFile = FileLen(Replace("" & pth, "..\", xAppPath & "\"))
If InStr(pth, "\") Then LenFile = FileLen(pth)
Err.Clear
End Function

Function PathFromFile(ByVal v)
On Error Resume Next
Dim ar: ar = Split(v, "\")
ReDim Preserve ar(UBound(ar) - 1)
PathFromFile = Join(ar, "\")
Err.Clear
End Function



'Function ListDir$(pth, Optional ByVal bRoot = 1)
'Dim objFolder, objAny, objCollection, res$, tp&
'Set objFolder = CreateObject("Scripting.FileSystemObject").GetFolder(pth)
'res = Replace("Path$500$,tp&,Icon&,Name$,Size&,Type$,Modified~,Created~", ",", vbTab)
''If Not objFolder.parentfolder Is Nothing Then res = res & vbCrLf & Replace(objFolder.parentfolder.Path & ",0," & SysImageListIndex(objFolder.parentfolder.Path, 1) & ",..", ",", vbTab)
'If Not objFolder.parentfolder Is Nothing Then If aVal(bRoot) Then res = res & vbCrLf & Replace(objFolder.parentfolder.Path & ",0," & SysImageListIndex("") & ",..", ",", vbTab)
'tp = 1
'For Each objCollection In Array(objFolder.subfolders, objFolder.files)
'For Each objAny In objCollection
'    'res = res & vbCrLf & Join(Array(Escape(objAny.Path), tp, SysImageListIndex(IIf(tp = 1, "", objAny.Path)), objAny.Name, objAny.SIZE, objAny.Type, objAny.DateLastModified, objAny.DateCreated), vbTab)
'    res = res & vbCrLf & Join(Array(modMain.JSONEscape(objAny.Path), tp, SysImageListIndex(IIf(tp = 1, "", objAny.Path)), objAny.Name, objAny.SIZE, objAny.Type, objAny.DateLastModified, objAny.DateCreated), vbTab)
'Next
'tp = tp + 1
'Next
'ListDir = CMatrix(0, res, "tp,Name").GetRows(0)
'End Function

Function SysIcon(ByVal sFileExtention) As Long
'If Len(sFileExtention) Then sys16 = 0
SysIcon = SysImageListIndex(sFileExtention)
'GetIcon SysIcon, sys16
End Function
Function SysColor(ByVal nColor, Optional ByVal v) As Long
'SysColor = c
'If SysColor > -1 And SysColor < 32 Then SysColor = GetSysColor(aVal(c))
SysColor = pSysColor(L_(nColor), v)
'If IsArray(SysColor) Then SysColor = ToArray(SysColor, "#i#", "item")
End Function

Public Function SaveToFile(ByVal Url, src)
Dim n&
n = FreeFile
Url = Replace(Url, "..\", xAppPath & "\")
On Error Resume Next
Kill Url
Open Url For Binary As n
Dim b() As Byte
If IsArray(src) Then b = src Else b = StrConv(src, vbFromUnicode)
Put n, 1, b
'Width #n, UBound(b) + 3
Close n
End Function

Public Function ReadFile(ByVal Url, Optional ByVal nReadLength& = -1, Optional ByVal charset = "", Optional ByVal nStartPos&) As String
Url = Replace(Url, "..\", xAppPath & "\")
If LenFile(Url) = 0 Then Exit Function
On Error Resume Next
Dim st As ADODB.Stream
Set st = New ADODB.Stream
st.Open
st.charset = NzS(charset, "windows-1251")

'nStartPos = Abs(nStartPos)
Dim n&
If nReadLength < 0 Then 'ALL CONTENT
    st.LoadFromFile Url
    If nStartPos Then
        If nStartPos > -1 Then
            st.Position = nStartPos
        Else
            st.Position = st.Size + nStartPos
        End If
    End If
ElseIf nReadLength > 0 Then  'PART OF FILE
    n = FreeFile
    ReDim buf(nReadLength - 1) As Byte
    Open Url For Binary As n
    
    If nStartPos > -1 Then
        Get n, nStartPos + 1, buf
    Else
        Get n, LOF(n) + nStartPos, buf
    End If
    Close n
    st.Type = adTypeBinary
    st.Write buf
    st.Position = 0
    st.Type = adTypeText
    'If nStartPos Then
'        st.charset = NzS(charset, "windows-1251")
'        Err.Clear
'        If Err Then st.charset = "Unicode"
    'End If
End If


ReadFile = st.ReadText(nReadLength)

st.Close

End Function


'Function UnChar(ByVal str, ByVal ch) As String
'UnChar = S_(str)
'ch = S_(ch)
'If IsNumeric(ch) Then ch = Chr(ch)
'If Left(UnChar, Len(ch)) = ch Then UnChar = Mid(UnChar, Len(ch) + 1)
'If Right(UnChar, Len(ch)) = ch Then UnChar = Left(UnChar, Len(UnChar) - Len(ch))
'End Function

Function TrimS(ByVal str, ByVal cs, Optional ByVal lR = -1) As String
str = S_(str): cs = S_(cs)
If Len(cs) Then
    If lR And 1 Then While Len(str) > 0 And InStr(cs, Left(str, 1)) > 0: str = Mid(str, 2): Wend
    If lR And 2 Then While Len(str) > 0 And InStr(cs, Right(str, 1)) > 0: str = Left(str, Len(str) - 1): Wend
End If
TrimS = str
End Function

Private Function pURL_Location(ByVal host, ByVal location) As String
Dim ar
If Left(location, 1) = "/" Then
    ar = Split(host, "/")
    ReDim Preserve ar(2) '  http://host = Array(http:,,host)
    'pURL_Location = UnChar(Join(ar, "/"), "/") & "/" & UnChar(location, "/")
    pURL_Location = TrimS(Join(ar, "/"), "/") & "/" & TrimS(location, "/")
Else
    pURL_Location = location
End If
End Function

'Public Function DownLoadFile(url, Optional ByVal timeOut = 30, Optional ByVal toFileorCharSet, Optional bAcyncVar, Optional ret) ', Optional sCharSet) 'as string
Public Function DownLoadFile(ByVal sUrl, Optional ByVal timeOut = 30, Optional ByVal toFileorCharSet, Optional hdr)  ', Optional sCharSet) 'as string
''[in] sUrl=             [POST |GET |]url_address [ | SendParam ] [ RequestHeader| RequestHeaderValue||||  ]   [ | Proxy:Port,ProxyUser,ProxyPassword,ServerUser,ServerPassword]
''[in] timeOut:   0=SyncDownload >0=AsyncDownload (timeOut)
''[in] toFileorCharSet:        filename to save response data or charset for return string
''[in] hdr:
''           >0 as Number = DISABLE Redirects
''           <0 as number =Return CDataObject + hdr=sendparam
''[out] hdr = status \r\n +rewsponseheaders
On Error Resume Next

Dim bErrSecurity As Boolean, bResp As Boolean, whr As WinHttpRequest, nAttempts&, sResponseHeaders$
Dim st As Stream
Dim method$, ar, Url$, sendparam, rh, proxy, i&
Dim bAsync As Boolean: bAsync = L_(timeOut) > 0
toFileorCharSet = S_(toFileorCharSet)
Dim eh As CDataObject

Dim bEn As Boolean: bEn = L_(hdr) <= 0 'ENABLE RDIRECTS

Do 'repeat_send: if bErrSecurity=true
    sendparam = ""
    ar = Split(sUrl, " "): method = "GET": Url = sUrl: If UBound(ar) > 0 And Len(ar(0)) < 10 Then method = ar(0): pArrayLine ar, 0, , 3: Url = Join(ar, " ")
    rh = Split(Url, "|"): If UBound(rh) > 0 Then Url = rh(0): pArrayLine rh, 0, , 3: sendparam = rh(0) Else rh = Array()
    Url = Replace(Url, Chr(1), "|")
    Set whr = New WinHttpRequest
    'If Not eh Is Nothing Then eh.Request.Abort
    If UBound(rh) Mod 2 And UBound(rh) > -1 Then 'Всегда в конце нечетный по счету параметр
        proxy = Split(rh(UBound(rh)), ",") ' = proxyserver:port,proxyuser,proxypassword,user,password
        ReDim Preserve proxy(4)
        If Len(proxy(0)) Then whr.SetProxy 2, proxy(0)
    End If
    
    whr.Option(WinHttpRequestOption_EnableRedirects) = bEn 'ENABLES REDIRECTS
    whr.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = SslErrorFlag_Ignore_All 'ENABLES REDIRECTS

If TypeName(hdr) = xs.sxControl Then
    Set eh = CDataObject(NzS(hdr.CPSource("HTTP_Name"), "http"), whr) 'EventProcPrefix
    eh.Value = L_(hdr.CPSource("HTTP_Cookie")) 'COOKIE(Event type): <0:ByName,  >=0:Event()
    eh.xHttpListener = hdr.hWnd 'xControl event handler
End If
    
    whr.Open method, Replace(Url, Chr(1), "|"), bAsync
    For i = 1 To UBound(rh) - (UBound(rh) Mod 2) Step 2: whr.SetRequestHeader rh(i), rh(i + 1): Next
    'whr.Option(WinHttpRequestOption_UserAgentString) = "Murzilla"
    If ArrayDims(proxy) = 1 Then
        If Len(proxy(1)) Then whr.SetCredentials proxy(1), proxy(2), 1 'PROXY  USER /PASSWORD
        If Len(proxy(3)) Then whr.SetCredentials proxy(3), proxy(4), 0 'SERVER USER /PASSWORD
    End If
    
    If Left(sendparam, 1) = "@" Then 'Первый символ = отправка файлов
        Dim f, fs, Boundary$
        Const DefaultBoundary = "--"
        ar = Split(sendparam, "@") 'Получаем список файлов        'sendparam =@file.ext@
        If UBound(ar) < 2 Then
            nAttempts = 10
        Else
            Boundary = ar(1)
            If RegExMatch(Boundary, "[A-Za-z0-9.]*") <> Boundary Then
                nAttempts = 10
            Else
                modMain.pArrayLine ar, 0, , 3: modMain.pArrayLine ar, 0, , 3
                ReDim fs(UBound(ar))
                For i = 0 To UBound(fs)
                    f = Split(ar(i) & "*", "*")
                    fs(i) = Array(Quot(Boundary & "[]"), f(0), Quot(NzS(f(1), SplitIndex(f(0)))))
                Next
                Boundary = String$(22, ".")
                Dim ch$
                For i = 1 To 12
                    Randomize
                    ch = Chr(CLng(Rnd * 25) + 97)
                    If i Mod 2 Then ch = UCase(ch)
                    Boundary = Boundary & ch
                Next
                whr.SetRequestHeader "Content-type", "multipart/form-data; boundary=" & Boundary
            'Формируем заголовки и тело запроса
                Dim buf() As Byte
                Dim fst As New Stream: fst.Type = adTypeBinary: fst.mode = adModeReadWrite: fst.Open
                Set st = New Stream: st.Type = adTypeBinary: st.mode = adModeReadWrite: st.Open
                For Each f In fs
                    If LenFile(f(1)) Then
                        buf = StrConv(vbCrLf & DefaultBoundary & Boundary & vbCrLf & "Content-Disposition: form-data; name=" & f(0) & "; filename=" & f(2) & vbCrLf & vbCrLf, vbFromUnicode)
                        st.Write buf: fst.LoadFromFile (Replace(f(1), "..\", xAppPath & "\")): st.Write fst.Read
                    End If
                Next
                buf = StrConv(vbCrLf & DefaultBoundary & Boundary & DefaultBoundary, vbFromUnicode)
                st.Write buf: st.Position = 0: sendparam = st.Read
            End If
        End If
        
        
    'Формируем заголовки и тело запроса
    Else 'IF Left(sendparam, 1) = "@" Then 'Первый символ = отправка файлов
        sendparam = Replace(sendparam, Chr(1), "|")
        If Left$(sendparam, 1) = ":" Then
            sendparam = Mid$(sendparam, 2) 'COMPUTED PARAM
            If Left$(sendparam, 1) = "~" Then
                sendparam = MainWnd.Eval(Mid$(sendparam, 2)) 'EVAL PARAM
            Else
                sendparam = m_Param(sendparam) 'PREPARED PARAM
            End If
        End If
        
        
    End If
    
    
    
    If L_(hdr) < 0 Then
        Set DownLoadFile = CDataObject("whr", whr)
        hdr = sendparam
        nAttempts = 10
        Exit Do
    End If


    
    
    Do While nAttempts < 10
        nAttempts = nAttempts + 1
        Err.Clear
        
'LogMsg vbMagenta, vbTab & "SEND " & sUrl
        
        whr.Send sendparam
        If bAsync Then
            bResp = whr.WaitForResponse(L_(timeOut))
        Else
            bResp = True: whr.WaitForResponse
        End If
'LogMsg vbCyan, vbTab & "Response " & sUrl
        
        If Err Then
            'Debug.Print Url: Debug.Print Err, Err.Description
            'i = Err: method = Err.Description
            'LogMsg 255, Url: LogMsg 255, i & ":" & method
            'i = LastErr(-31)
            'i = Err
            If Err.Number = -2147012721 And Not bErrSecurity Then 'repeat_send
                bErrSecurity = True: Exit Do 'GoTo repeat_send
            Else
                nAttempts = 10: bErrSecurity = False 'GoTo Exit_Function
            End If
        Else
            bErrSecurity = False
            sResponseHeaders = sResponseHeaders & whr.GetAllResponseHeaders
            Select Case whr.Status
            Case 301, 302
                If Not bEn Then Exit Do 'DISABLE REDIRECTS
                Url = pURL_Location(whr.Option(WinHttpRequestOption_URL), whr.GetResponseHeader("Location"))
                whr.Open method, Url, bAsync   'True
                For i = 1 To UBound(rh) - (UBound(rh) Mod 2) Step 2: whr.SetRequestHeader rh(i), rh(i + 1): Next

            Case 401, 407
                If ArrayDims(proxy) = 1 Then
                    If Len(proxy(1)) Then whr.SetCredentials proxy(1), proxy(2), 1 'PROXY  USER /PASSWORD
                    If Len(proxy(3)) Then whr.SetCredentials proxy(3), proxy(4), 0 'SERVER USER /PASSWORD
                Else
                    Exit Do
                End If
            Case Else: Exit Do
            End Select
        End If
    Loop

Loop While bErrSecurity

If nAttempts < 10 Then
    Set st = ADOStream(whr.ResponseStream)
    
    Dim cs$: cs = whr.GetResponseHeader("Content-Encoding")
    Select Case cs
    Case "gzip", "deflate", "x-gzip", "gzip, deflate"
        'st.Position = 0
        'st.Write CommputeParams(gDecompress(st.Read), CallByName(st, "Position", VbLet, 0))  'WRITE BYTES TO THE STREAM
        st.Write gDecompress(st.Read, CallByName(st, "Position", VbLet, 0))  'WRITE BYTES TO THE STREAM
        st.SetEOS
    End Select
    
    st.Position = 0
    If InStr(toFileorCharSet, "\") > 0 Then
        If st.Size > 0 Then
            MakeDir toFileorCharSet
            st.SaveToFile toFileorCharSet, adSaveCreateOverWrite
            DownLoadFile = toFileorCharSet
        End If
    ElseIf Val(toFileorCharSet) Then 'return stream 'Byte()
        Set DownLoadFile = st
    Else
        st.Type = adTypeText
        If Len(toFileorCharSet) Then
            st.charset = toFileorCharSet
        Else
            st.charset = "_autodetect_all"
            cs = "": cs = whr.GetResponseHeader("Content-Type")
            cs = Trim(Split(cs & ";", ";")(1))
            If InStr(cs, "charset") Then st.charset = Split(cs, "=")(1)
        End If
        DownLoadFile = st.ReadText
    End If
    hdr = whr.Status & vbCrLf & sResponseHeaders
    hdr = Replace(hdr, vbCrLf & vbCrLf, vbCrLf)
    hdr = Left(hdr, Len(hdr) - 2)
Else
    hdr = ""
End If
'LogMsg 255, vbTab & "End download file " & sUrl
End Function
'Function CommputeParams(buf, ParamArray Params())
'CommputeParams = buf
'End Function
'=============== FILE SYSTEM ===============

Function XMLAttributePrefix(Optional ByVal v = Null)
If VarType(v) > 1 Then gXMLAttributePrefix = Left(S_(v), 1)
XMLAttributePrefix = gXMLAttributePrefix
End Function

Function CParam(ar, Optional ByVal tp) As CParam
Set CParam = New CParam
Select Case S_(tp)
Case "json": CParam.json = S_(ar)
Case "xml": CParam.XML = S_(ar)
'Case "names": CParam.FillNames ar, Empty
'Case "file": CParam.Load S_(ar)
Case Else
    If VarType(ar) <> vbString Then CParam.Value = ar Else If Len(ar) Then CParam.Load ar
End Select
End Function

Function CMatrix(sRowSourceType, vRowSource, Optional sSortFilterGroupBy, Optional nRowMode = 0) As CMatrix
Set CMatrix = New CMatrix
CMatrix.OpenRows sRowSourceType, vRowSource, L_(nRowMode), , , S_(sSortFilterGroupBy)
End Function

Function CZip(zip_local_file) As CZip
Set CZip = New CZip:
CZip.OpenZip S_(zip_local_file)
End Function

Function CDataObject(nm, v) As CDataObject
Set CDataObject = New CDataObject
CDataObject.Name = S_(nm)
CDataObject.Value = v
End Function

'==========================================================================
'=========DIFFERENSE======================================================
'Function CParamDIFF(ByVal ar1, ByVal ar2)
''AR1<<DIFF=AR2
'Dim cp As New CParam
'Dim cp1 As New CParam
'If ArrayDims(ar1) <> 1 Or ArrayDims(ar2) <> 1 Then Exit Function
'
'cp1.Value = ar2
'
'Dim i&, nm$
'Dim p1, p2
'
'For i = 0 To UBound(ar1) - ((1 + UBound(ar1)) Mod 2) Step 2
'nm = ar1(i)
'p1 = ar1(i + 1)
'p2 = cp1.Param(nm)
'If VarType(p2) = vbNull Then 'Нету такого в результате
'    If VarType(p1) > vbNull Then cp.Param(nm) = Empty    'Пишем что надо удалить его
'Else 'PARAM EXIST
'    If IsArray(p2) Then
'        If ArrayDims(p2) = 1 Then
'            p2 = CParamDIFF(p1, p2)
'            If UBound(p2) > -1 Then cp.Param(nm) = p2
'        Else
'            cp.Param(nm) = p2
'        End If
'    Else 'Есть отличие
'        If p1 <> p2 Then cp.Param(nm) = p2
'    End If
'End If
'cp1.Param(nm) = Null
'Next
'
'ar2 = cp1.Source
'For i = 0 To UBound(ar2) Step 2
'    nm = ar2(i)
'    p2 = ar2(i + 1)
'    If VarType(p2) = 1 Then cp.Param(nm) = Empty Else cp.Param(nm) = p2
'Next
'CParamDIFF = cp.Source
'End Function

'============ GLOBAL PARAM =========================================
Public Function p_()
Set p_ = Dispatch(m_Param)
End Function
Public Property Get Param(ByVal key)
Param = m_Param(key)
End Property
Public Function Param_(ByVal key, v)
Param(key) = v
Param_ = m_Param(key)
End Function
Public Property Let Param(ByVal key, v)
'If m_Param Is Nothing Then Set m_Param = New CParam
m_Param(key) = v
If Left$(key, 10) = App.ProductName & "\" Then
Select Case Mid$(key, Len(App.ProductName) + 2)
Case "hEventDebug": xa.ghEventDebug = L_(v)
Case "WriteLog": xa.gWriteLog = L_(v)
Case "AllowCPE": xa.gAllowCPE = L_(v)
Case "AllowF6": xa.gAllowF6 = L_(v)
Case "LabelForeColor": xa.gLabelForeColor = L_(v)
Case "Font": xMainWnd.Font = S_(v)
Case "SysColor": pSysColor 0, CParam(v, "json").Source
'Case "MenuIcons": Dim m, p: For Each m In CParam(v, "json").Map: For Each p In CParam(m.Value).Map: ReplaceIcon p.Name, m.Name, p.Value: Next: Next

'XVBEngine\hEventDebug
End Select
End If
End Property
Public Property Get ParamXV(ByVal nm)
ParamXV = m_Param("xv\" & nm)
End Property
Public Property Let ParamXV(ByVal nm, v)
m_Param("xv\" & nm) = v
End Property

'============ GLOBAL PARAM =========================================


'==========================================================================
'============ARRAYS===================================================

Function VarInfo(ar)
VarInfo = modMain.VarInfo(ar)
End Function

'Function NewArray(Optional ByVal sarBounds = "8204,0,0,0")  'Dim array to custom dims and  bounds
'NewArray = modMain.pNewArray(sarBounds)
'End Function

Function ArrayDims(ar) As Long
ArrayDims = modMain.mSafeArray(ar).cDims
End Function
Function ArrayShift(ar, ByVal delta)
ArrayShift = modMain.ArrayShift(ar, 0 + delta)
End Function

'Function ArrayReDim(ar, Optional ByVal sarBounds = "0,-1")
''sarBounds = "LB1,UB1,LB2,UB2,...."?=>1DimArray()
'ArrayReDim = modMain.pArrayReDim(ar, sarBounds)
'End Function
Function ArrayGetLine(ar, ByVal index)
ArrayGetLine = modMain.pArrayLine(ar, 0& + index)
End Function
Function ArrayLetLine(ar, ByVal index, vLine)
modMain.pArrayLine ar, 0& + index, vLine, 1
ArrayLetLine = ar
End Function
Function ArrayInsertLine(ar, ByVal iBeforeIndex, Optional newLineValue, Optional ByVal count = 1)
Do While count > 0
modMain.pArrayLine ar, 0& + iBeforeIndex, newLineValue, 2
count = count - 1
Loop
ArrayInsertLine = ar
End Function

Function ArrayDeleteLine(ar, ByVal index, Optional ByVal count = 1)
Do While count > 0
modMain.pArrayLine ar, 0& + index, , 3
count = count - 1
Loop
ArrayDeleteLine = ar
End Function

'Function SplitInsert$(ByVal sar, ByVal v, Optional ByVal delimeter = "\", Optional ByVal index = -1)
''Вставляет элемент в массив перед указанным индексом
''index <0 начиная с конца
'Dim ar
'If Len(Nz(delimeter)) Then
'    ar = Split(Nz(sar), delimeter)
'    index = aVal(index)
'    If index < 0 Then index = UBound(ar) + index + 1
'    If index > UBound(ar) Then index = UBound(ar) + 1
'     modMain.pArrayLine ar, 0& + index, v, 2
'    sar = Join(ar, delimeter)
'End If
'SplitInsert = Nz(sar)
'End Function

Function SplitRemove$(ByVal sar, Optional ByVal delimeter = "\", Optional ByVal index = -1)
'Удаляет указанный индекс из SPLIT ARRAYя
'index <0 начиная с конца
Dim ar
If Len(S_(delimeter)) Then
    ar = Split(S_(sar), delimeter)
    index = L_(index)
    If index < 0 Then index = UBound(ar) + index + 1
    If index > UBound(ar) Then index = UBound(ar)
    modMain.pArrayLine ar, 0& + index, , 3
    sar = Join(ar, delimeter)
End If
SplitRemove = S_(sar)
End Function

Function SplitIndex$(ByVal sar, Optional ByVal delimeter = "\", Optional ByVal index = -1)
'Возвращает указааный элемент масива
'index <0 начиная с конца
Dim ar
If Len(S_(delimeter)) Then
ar = Split(S_(sar), delimeter)
index = L_(index)
If index < 0 Then index = UBound(ar) + index + 1
'If index > UBound(ar) Then index = UBound(ar)
On Error Resume Next
SplitIndex = ar(index)
Err.Clear
End If
End Function

'Function SplitValue(ByVal sar, ByVal v, Optional ByVal delimeter = "\", Optional ByVal index = -1)
'Dim ar
'SplitValue = sar
'If Len(Nz(delimeter)) Then
'    ar = Split(vbNullString & sar, delimeter)
'    If index < 0 Then index = UBound(ar) + index + 1
'    If index <= UBound(ar) Then ar(index) = v:  SplitValue = Join(ar, delimeter)
'End If
'End Function
Function Join2(sar, Optional ByVal RowDelimeter$ = vbCrLf, Optional ByVal ColumnDelimeter$ = vbTab, Optional NullExpr$)
Join2 = modMain.Join2(sar, RowDelimeter$, ColumnDelimeter$, NullExpr)
End Function

'Function Split2(ByVal src, Optional ByVal RowDelimeter$ = vbCrLf, Optional ByVal ColumnDelimeter$ = vbTab)
''Split2 = modMain.Split2(s, RowDelimeter$, ColumnDelimeter$, NullExpr)
'Dim i&, j&, ar, ar2, u1&, u2&, cs, cols$
'src = Nz(src)
'ar = Split(src, RowDelimeter) 'Строки
'u2 = UBound(ar) 'Количество строк -1
'If InStr(1, src, ColumnDelimeter) Then
'    u1 = UBound(Split(ar(0), ColumnDelimeter))  'Сколько столбцов в первой строке -1
'    cols = Replicate(u1, ColumnDelimeter)
'    ReDim ar2(u1, u2)
'    For i = 0 To u2 'Столбцы
'        cs = Split(ar(i) & cols, ColumnDelimeter)
'        For j = 0 To u1
'            ar2(j, i) = cs(j)
'        Next
'    Next
'    Split2 = ar2
'Else
'    Split2 = ar
'End If
'End Function

Function Rotate(sar)
Dim x&, Y&, ar, r As RECT: ar = sar
With mSafeArray(sar)
If .cDims = 2 Then
    SetRect r, 0, 0, .Bounds(2).cElements - 1, .Bounds(1).cElements - 1
    OffsetRect r, .Bounds(2).lLbound, .Bounds(1).lLbound
    ReDim ar(r.Top To r.Bottom, r.Left To r.Right)
    For Y = r.Top To r.Bottom: For x = r.Left To r.Right: ar(Y, x) = sar(x, Y): Next: Next
End If
End With
Rotate = ar
End Function

'=============  SARRAYS ============================

'JSON FROM ARRAY
Function ArrayToJson(ar, Optional ByVal nEscapeOptions = 1)
ArrayToJson = modMain.ArrayToJson(ar, , Cast(nEscapeOptions, 3))
End Function
 'ARRAY FROM JSON
Function JsonToArray(sJson, Optional b2Dim As Boolean)
JsonToArray = modMain.JsonToArray(sJson, b2Dim)
End Function

'Function JSONEscape(src, Optional ByVal bQuot As Boolean) As String
'JSONEscape = modMain.JSONEscape(src, bQuot)
'End Function
'Function JSONUnEscape(src) As String
'JSONUnEscape = modMain.JSONUnEscape(src)
'End Function
'Function EscapeUnEscape(sar, Optional bUnEscape = 0, Optional ByVal nOptions = 0) 'As Boolean)
'On Error Resume Next
'EscapeUnEscape = modMain.pEscapeUnEscape(sar, bUnEscape, aVal(nOptions))
'End Function
Function xEscape(sar, Optional ByVal nEscapeOptions = 0)
On Error Resume Next
'xEscape = EscapeUnEscape(sar, 0, nEscapeOptions)
xEscape = pEscapeUnEscape(sar, 0, nEscapeOptions)
End Function
Function xUnEscape(sar)
On Error Resume Next
'xUnEscape = EscapeUnEscape(sar, 1)
xUnEscape = pEscapeUnEscape(sar, 1)
End Function


Sub aValAr(ParamArray ar())
On Error Resume Next
Dim i, inAR0 As Boolean
If UBound(ar) = 0 Then If ArrayDims(ar(0)) = 1 Then inAR0 = 1
If inAR0 Then
    For i = 0 To UBound(ar(0))
    If IsObject(ar(0)(i)) Then ar(0)(i).Value = aVal(ar(0)(i)) Else ar(0)(i) = aVal(ar(0)(i))
    Next
Else
    For i = 0 To UBound(ar)
    If IsObject(ar(i)) Then ar(i).Value = aVal(ar(i)) Else ar(i) = aVal(ar(i))
    Next
End If
End Sub

Public Function ArraySort(ByRef ArrayToSort As Variant, Optional ByVal vtSortCast = vbVariant, Optional ByVal order = 0)
If ArrayDims(ArrayToSort) <> 1 Then Exit Function
On Error Resume Next
Dim vt%, i&, j&
Dim imin&, imax&
order = L_(order) > 0
Dim varSwap As Variant
Dim varSwapVt As Variant
Dim v As Boolean, vc
imin = LBound(ArrayToSort) + 1
imax = UBound(ArrayToSort)
vt = vtSortCast
If vt < 2 Then vt = vbVariant
For i = imin To imax
    varSwap = ArrayToSort(i)
    varSwapVt = modMain.Cast(varSwap, vt)
    For j = i To imin Step -1
        vc = modMain.Cast(ArrayToSort(j - 1), vt)
        If varSwapVt = vc Then v = varSwap < ArrayToSort(j - 1) Else v = varSwapVt < vc
        If order Then v = Not v
        If v Then ArrayToSort(j) = ArrayToSort(j - 1) Else Exit For
    Next j
    ArrayToSort(j) = varSwap
Next i
ArraySort = ArrayToSort
End Function

Public Function ArrayGroup(ArrayToGroup)
If ArrayDims(ArrayToGroup) <> 1 Then Exit Function
On Error Resume Next
Dim i&, p
p = ArrayToGroup(0): i = 1
While i <= UBound(ArrayToGroup)
If p = ArrayToGroup(i) Then gRemoveIndex ArrayToGroup, p Else p = ArrayToGroup(i): i = i + 1
Wend
ArrayGroup = ArrayToGroup
End Function


'============ARRAYS===================================================
'==========================================================================
'==========================================================================


''*******SHEDULE******************
'Function SheduleTime(ByVal curDateTime, ByVal prepareArgs, ByVal nStartDate, _
'    Optional ByVal nStartTime, _
'    Optional ByVal nEndDate, _
'    Optional ByVal nEndTime, _
'    Optional ByVal nRunTime, _
'    Optional ByVal nEveryTime, _
'    Optional ByVal nWaitTime, _
'    Optional ByVal nType, _
'    Optional ByVal nDay, _
'    Optional ByVal nWeek, _
'    Optional ByVal nWeekDays, _
'    Optional ByVal nMonthDay, _
'    Optional ByVal nWeekIndex, _
'    Optional ByVal nWeekDay, _
'    Optional ByVal nMonths)
''возвращает >=0 склько времени осталось ждать до старта задания
''возвращает <0 склько времени уже идет задание
'Dim cp As New CParam
'If aVal(prepareArgs) Then
'    If ArrayDims(nStartDate) = 1 Then
'        cp.AddParams nStartDate
'        nStartDate = cp.Param("StartDate")
'        nStartTime = cp.Param("StartTime")
'        nEndDate = cp.Param("EndDate")
'        nEndTime = cp.Param("EndTime")
'        nRunTime = cp.Param("RunTime")
'        nEveryTime = cp.Param("EveryTime")
'        nWaitTime = cp.Param("WaitTime")
'        nType = cp.Param("Type")
'        nDay = cp.Param("Day")
'        nWeek = cp.Param("Week")
'        nWeekDays = cp.Param("WeekDays")
'        nMonthDay = cp.Param("MonthDay")
'        nWeekIndex = cp.Param("WeekIndex")
'        nWeekDay = cp.Param("WeekDay")
'        nMonths = cp.Param("Months")
'    End If
'    nStartDate = Cast(nStartDate, 3)
'    nStartTime = aVal(nStartTime)
'    nEndDate = IF0(Cast(nEndDate, 3), 99999)
'    nEndTime = IF0(aVal(nEndTime), 0.99999)
'    nRunTime = IF0(aVal(nRunTime), 0.5)
'    nEveryTime = aVal(nEveryTime)
'    nWaitTime = aVal(nWaitTime)
'    nType = Cast(nType, 3)
'    nDay = Cast(nDay, 3)
'    nWeek = Cast(nWeek, 3)
'    nWeekDays = vbNullString & NzS(nWeekDays, 1)
'    nMonthDay = Cast(nMonthDay, 3)
'    nWeekIndex = Cast(nWeekIndex, 3)
'    nWeekDay = Cast(nWeekDay, 3)
'    nMonths = vbNullString & NzS(nMonths, 1)
'End If
'
'Dim ret As Double
'Dim i&, retWeeks&, curWeekDay&, iwd&
'Dim cTime As Double, retTime As Double
'Dim elapsedTime As Double
'ret = nStartDate + nStartTime - curDateTime 'Сколько ждать времени до начала
'If ret > 0 Then SheduleTime = ret: Exit Function  'Не надо выполнять = рано еще
''If nEndDate + nEndTime < curDateTime Then JobTime = nEndDate + nEndTime - curDateTime: Exit Function 'Не надо выполнять = уже поздно
'If nEndDate + nEndTime < curDateTime Then SheduleTime = 999: Exit Function 'Не надо выполнять = уже поздно
'
'ret = 0
'Select Case nType
'Case 0 'Однократно
'  'If nStartDate + nStartTime + nRunTime > curDateTime Then JobTime = nStartDate + nStartTime + nRunTime - curDateTime: Exit Function  'Не надо выполнять = уже поздно
'  If nStartDate + nStartTime + nRunTime > curDateTime Then SheduleTime = 999: Exit Function  'Не надо выполнять = уже поздно
'
'Case 1 'Ежедневно
'  ret = (Fix(curDateTime) - nStartDate) Mod nDay 'Через сколько дней выполнить
'
'Case 2 'Еженедельно
'  i = Fix((Fix(curDateTime) - nStartDate + 7 - aVal(Weekday(CDate(nStartDate), 2))) / 7) 'Сколько недель прошло от понедельника недели старта
'  retWeeks = nWeek + (i Mod nWeek) + i 'Через сколько целых недель выполнено
'  If retWeeks Mod nWeek > 0 Then ret = retWeeks * 7  'Через сколько дней (целых недель) выполнить
'  curWeekDay = Weekday(CDate(Fix(curDateTime)), 2) 'Номер дня недели пн=1,вт=2,..,вс=7
'  For i = curWeekDay To 6 + curWeekDay
'    If i > 7 Then iwd = i - 7 Else iwd = i
'    If InStr(1, "," & nWeekDays & ",", "," & iwd & ",") > 0 Then Exit For 'День недели совпал
'    ret = ret + 1 'Через сколько дней выполнить
'  Next
'
'Case 3 'Ежемесячно
'  ret = NextMonthDayDelta(Fix(curDateTime), nMonthDay, nWeekIndex, nWeekDay, nMonths, nEndDate)
'
'End Select
'
''ret=Сколько дней ждать до следующего повтора от текущего [curTime] =1 день
'
'cTime = curDateTime - Fix(curDateTime) 'Текущее время
'retTime = nStartTime - cTime 'Через сколько часов выполнить
'If retTime > 0 Or ret > 0 Then SheduleTime = ret + retTime: Exit Function
'
'If nEveryTime > 0 Then 'Надо повторять в течении дня
'  If cTime > IIf(nWaitTime > 0, nStartTime + nWaitTime, -nWaitTime) And ret = 0 Then 'На сегодня повторы исчерпаны
'    SheduleTime = 1 - cTime + SheduleTime(Fix(curDateTime) + 1, 0, nStartDate, nStartTime, nEndDate, nEndTime, nRunTime, nEveryTime, nWaitTime, nType, nDay, nWeek, nWeekDays, nMonthDay, nWeekIndex, nWeekDay, nMonths)
'  Else 'Можно повторять
'    i = Fix((cTime - nStartTime) / nEveryTime) + 1 'Сколько раз повторилось за день +1
'    elapsedTime = i * nEveryTime 'Продолжительность до следующего повтора
'    If nStartTime + elapsedTime > IIf(nWaitTime > 0, nStartTime + nWaitTime, -nWaitTime) And ret = 0 Then 'На сегодня повторы исчерпаны
'      SheduleTime = 1 - cTime + SheduleTime(Fix(curDateTime) + 1, 0, nStartDate, nStartTime, nEndDate, nEndTime, nRunTime, nEveryTime, nWaitTime, nType, nDay, nWeek, nWeekDays, nMonthDay, nWeekIndex, nWeekDay, nMonths)
'    Else 'Еще есть время для повтора
'      If nStartTime + elapsedTime - nEveryTime + nRunTime > cTime Then 'Попадаем в промежуток выполнения
'        SheduleTime = nStartTime + elapsedTime - nEveryTime - cTime '<0 Сколько уже идет текущее провторение выполнения
'      Else 'Сколько ждать до начала повтора
'        SheduleTime = ret + nStartTime + elapsedTime - cTime
'      End If
'    End If
'  End If
'Else 'Не надо повторять
'  If cTime > nEndTime And ret = 0 Then 'На сегодня повторы исчерпаны
'    SheduleTime = 1 - cTime + SheduleTime(Fix(curDateTime) + 1, 0, nStartDate, nStartTime, nEndDate, nEndTime, nRunTime, nEveryTime, nWaitTime, nType, nDay, nWeek, nWeekDays, nMonthDay, nWeekIndex, nWeekDay, nMonths)
'  Else 'Выполнять
'    SheduleTime = nStartTime - cTime '<0 Сколько уже идет выполнение
'  End If
'End If 'Не надо повторять
'End Function
'
'
'Private Function NextMonthDayDelta(curDate, nMonthDay, nWeekIndex, nWeekDay, nMonths, nEndDate) 'Через сколько дней стартовать задание
'Dim delta&, mDays&, fndDate&
'NextMonthDayDelta = 999
''If curDate > nEndDate Then NextMonthDayDelta = nEndDate - curDate: Exit Function
'If curDate > nEndDate Then Exit Function
'mDays = Day(LDat(CDate(curDate))) - Day(curDate) + 1 'Добавка дней до начала следующего месяца
'If InStr(1, "," & nMonths & ",", "," & Month(curDate) & ",") > 0 Then 'Месяц совпал
'  If nMonthDay > 0 Then 'По номеру дня месяца
'    delta = nMonthDay - Day(curDate)
'    If delta < 0 Then 'Надо найти в следующем месяце
'      NextMonthDayDelta = mDays + NextMonthDayDelta(curDate + mDays, nMonthDay, nWeekIndex, nWeekDay, nMonths, nEndDate)
'    Else
'      NextMonthDayDelta = delta
'    End If
'  Else 'По номерам дней недели nDayIndex,nWeekDay
'    delta = nWeekDay - Weekday(Fdat(curDate), 2)
'    If delta < 0 Then delta = delta + 7
'    fndDate = Fdat(curDate) + delta + (nWeekIndex - 1) * 7
'    While Month(fndDate) <> Month(curDate)
'        fndDate = fndDate - 7
'    Wend
'    delta = fndDate - CLng(curDate)
'    If delta >= 0 Then 'Нашол
'      NextMonthDayDelta = delta
'    Else 'Надо найти в следующем месяце
'      NextMonthDayDelta = mDays + NextMonthDayDelta(curDate + mDays, nMonthDay, nWeekIndex, nWeekDay, nMonths, nEndDate)
'    End If
'  End If 'По номерам дней недели nDayIndex,nWeekDay
'Else 'Надо найти в следующем месяце
'  NextMonthDayDelta = mDays + NextMonthDayDelta(curDate + mDays, nMonthDay, nWeekIndex, nWeekDay, nMonths, nEndDate)
'End If
'End Function
'
'Function SheduleText(jobSRC)
'Dim res$, cp As CParam, n, ar, a, tmp$
'Set cp = CParam(jobSRC)
'res = "начать в " & FormatVB(aVal(cp.Param("StartTime")), "hh:nn") & ","
'Select Case aVal(cp.Param("Type"))
'Case 0: res = res & " однократно, "
'Case 1
'  n = aVal(cp.Param("Day"))
'  If n < 2 Then res = res & " ежедневно, " Else res = res & " каждый " & n & " день, "
'Case 2
'  n = aVal(cp.Param("Week"))
'  res = res & " каждую " & IIf(n > 1, n & " ", "") & "неделю"
'  ar = Split(NzS(cp.Param("WeekDays"), "1"), ",")
'  aValAr ar
'  ArraySort ar
'  For Each a In ar
'    tmp = tmp & WeekdayName(a, 1, 2) & ", "
'  Next
'  res = res & " по " & tmp
'Case 3
'  n = aVal(cp.Param("MonthDay"))
'  If n > 0 Then
'    res = res & " " & n & " числа "
'  Else
'    n = aVal(cp.Param("WeekIndex"))
'    If n > 4 Then n = 5
'    res = res & " по "
'    Select Case n
'    Case 1: tmp = "первым"
'    Case 2: tmp = "вторым"
'    Case 3: tmp = "третьим"
'    Case 4: tmp = "четвертым"
'    Case 5: tmp = "последним"
'    Case Else: tmp = ""
'    End Select
'    res = res & tmp & " "
'    res = res & WeekdayName(aVal(cp.Param("WeekDay")), 1, 2) & " "
'  End If
'  ar = Split(NzS(cp.Param("Months"), "1"), ",")
'  If UBound(ar) < 11 Then
'    tmp = ""
'    aValAr ar
'    ArraySort ar
'    For Each a In ar
'      If Len(tmp) Then tmp = tmp & ","
'      tmp = tmp & MonthName(a, 1)
'    Next
'    res = res & tmp & ", "
'  Else
'    res = res & "каждого месяца, "
'  End If
'Case Else
'  res = res & "не пойму что это, "
'End Select
'n = aVal(cp.Param("EveryTime"))
'If n > 0 Then
'  res = "каждые " & Format$(n, "hч.nnм.") & ", " & res
'  n = aVal(cp.Param("WaitTime"))
'  If n < 0 Then res = res & "до " & Format$(CDate(Abs(n)), "hh:nn") Else res = res & "в течении " & Format$(CDate(n), "hч.nnм.")
'  res = res & ", "
'Else
'  n = 0
'End If
'If n >= 0 Then
'  n = aVal(cp.Param("EndTime"))
'  If n > 0 Then res = res & "до " & FormatVB(n, "hh:nn") & ", "
'End If
'res = res & "начиная с " & Format$(aVal(cp.Param("StartDate")), "dd.mm.yyyy")
'
'n = aVal(cp.Param("EndDate"))
'If n > 0 Then res = res & " и заканчивая " & Format$(n, "dd.mm.yyyy")
'SheduleText = res
'
'End Function
''*******SHEDULE******************

''======= RECORD LOCKING ==============================================
#If LogLock Then '==================LOGLOCK=================
Sub CheckLGLC()
On Error Resume Next
xMainWnd.GlobalModule.CodeObject.CheckLLS
b_LC_TABLE = L_(xMainWnd.Vars("LC_TABLE"))
b_LG_TABLE = L_(xMainWnd.Vars("LG_TABLE"))
End Sub
#End If '==================LOGLOCK=================
'Sub CheckLGLOCKS()
'On Error Resume Next
'b_LG_TABLE = gRemoveIndexArray(Split("dd,uu,tbl,tp,id,log", ","), ToArray(m_Connection.Execute("SELECT * FROM lg WHERE 0=1").Fields)) = 0
'b_LC_TABLE = gRemoveIndexArray(Split("lid,tbl,id,sid,hwnd,dd", ","), ToArray(m_Connection.Execute("SELECT * FROM locks WHERE 0=1").Fields)) = 0
'End Sub

'Public Function RowLocked&(ByVal tbl, ByVal ID, ByVal hWnd) 'Return LockID
''RowLocked >0 заблокировано мной
''RowLocked =0 на заблокированно
''RowLocked <0 заблокировано не мной
'If Not b_LC_TABLE Then Exit Function
'Dim LockID, SessionID, WindowID
'ID = aVal(ID): tbl = Nz(tbl): hWnd = aVal(hWnd)
'If ConState = 1 And ID > 0 And Len(tbl) Then
'    GetFields "SELECT lid, sid, hwnd FROM locks WHERE tbl='" & tbl & "' AND id=" & ID, LockID, SessionID, WindowID 'Получаем реквизиты текущей блокировки
'    aValAr LockID, SessionID, WindowID
'    If LockID = 0 Then 'не блокированно
'        LockID = RecInfo(tbl, ID)
'    ElseIf SessionID = Base.SID Then 'Заблокированно текцщим процессом
'        If WindowID <> hWnd Then LockID = -LockID 'Заблокированно другим окном
'    Else
'        LockID = -LockID 'Заблокированно другим процессом
'    End If
'    RowLocked = LockID
'End If
'End Function
'
'Sub RemoveLock(ByVal LockID)
'If Not b_LOCKS_TABLE Then Exit Sub
'LockID = aVal(LockID)
'If LockID <> 0 Then If ConState = 1 Then RunSQL "DELETE locks" & IIf(XConnection.Provider = "SQLOLEDB.1", "", ".* FROM locks") & " WHERE lid=" & LockID
'End Sub
'
'Public Function RowLock&(ByVal tbl, ByVal ID, ByVal hWnd, Optional ByVal bLock As Boolean)   'Set/Reset Lock RETURN LOCKID
'If Not b_LOCKS_TABLE Then Exit Function
''RowLock(bLock=0) =0 удачное саморазблокирование
''RowLock(bLock=1) >0 удачное блокирование
''RowLock(bLock=0/1) <0 неудачное разблокирование/блокирование
'Dim LockID, SessionID, WindowID, sq0$
'ID = aVal(ID): tbl = Nz(tbl): hWnd = aVal(hWnd)
'If ConState = 1 And ID > 0 And Len(tbl) Then
'    sq0 = IIf(XConnection.Provider = "SQLOLEDB.1", "", ".* FROM locks")
'    RunSQL "DELETE locks" & sq0 & " INNER JOIN ss ON locks.sid=ss.id WHERE ss.tm2>0" 'Удаляем недействительные блокировки
'    GetFields "SELECT lid, sid, hwnd FROM locks WHERE tbl='" & tbl & "' AND id=" & ID, LockID, SessionID, WindowID 'Получаем реквизиты текущей блокировки
'    aValAr LockID, SessionID, WindowID
'    If LockID = 0 And bLock Then  'Не заблокированно = БЛОКИРУЕМ
'        If RecInfo(tbl, ID) Then 'Не могу заблокировать
'            LockID = -1
'        Else
'            LockID = aVal(SetFields("SELECT sid, tbl, id, dd, hwnd, lid FROM locks WHERE lid=0", "lid", Base.SID, tbl, ID, Now(), hWnd))
'        End If
'        RowLock = LockID
'    ElseIf SessionID = Base.SID And WindowID = hWnd Then  'Заблокированно текущей сессией и указанным окном
'        If bLock Then
'            RowLock = LockID
'        Else 'Удаляем свою блокировку
'            RunSQL "DELETE locks" & sq0 & " WHERE lid=" & LockID
'            RowLock = aVal(RunSQL("SELECT lid FROM locks WHERE tbl='" & tbl & "' AND id=" & ID, 0))
'        End If
'    Else 'Заблокированно Другой сессией = Return LockID
'        RowLock = -LockID
'    End If
'End If
'End Function

'Function LockInfo(LockID, retArray)
'If b_LC_TABLE=0 Then Exit Function
'If ConState<>1 Then exit function
'LockID=Abs(aVal(LockID))
'If LockID=0 Then exit function
'Dim UserID,UserName,LockTime,UserIP,UserPort,UserWindow,sql
'RunSQL"SELECT uid,unm,lc.dd,ip,port,hwnd FROM lc INNER JOIN (SELECT ss.id,ss.ip,ss.port,ss.uid,uz.unm FROM ss INNER JOIN uz ON ss.uid=uz.uid) as hp ON lc.sid = hp.id WHERE lc.lid="&LockID,UserID,UserName,LockTime,UserIP,UserPort,UserWindow
'If retArray Then
'   LockInfo=Array(UserID,UserName,LockTime,UserIP,UserPort,UserWindow)
'Else
'   LockInfo="пользователь:"&Nz(UserName)&vbcrlf&"комп:"&IPCompNameFromLong(UserIP)&"("&IPNameFromLong(UserIP)&":"&UserPort & ")"&vbcrlf&"время:"&FormatVB(LockTime,"dd.mm.yy hh:nn:ss")&vbcrlf&"окно:"& Hex(UserWindow)
'End If
'End Function

'Public Function RecInfo&(ByVal tbl, ByVal ID, Optional ByRef s)
''RecInfo = False = ALLOW EDIT
''RecInfo = True = DENY EDIT
'Dim uu, dd, pkey
'ID = aVal(ID): tbl = Nz(tbl)
'If ConState = 1 And ID > 0 And Len(tbl) Then
'    pkey = PrimaryKey("" & tbl)
'    If Len(pkey) Then
'        If GetFields("SELECT uu, dd FROM " & Nz(tbl) & " WHERE " & Nz(pkey) & "=" & aVal(ID), uu, dd) Then
'            aValAr uu, dd
'            RecInfo = Base.UserLevel() < Base.UserLevel(uu)
'            s = Base.UserName(uu) & " Level:" & Base.UserLevel(uu) & " " & CDate(dd)
'        End If
'    End If
'End If
'End Function
'
'Sub LogRecord(ByVal tbl, ByVal ID, ByVal hWnd, txLog, Optional ByVal tp = 0&) 'As Boolean
'If Not b_LG_TABLE Then Exit Sub
'If ConState = 1 Then
'    If RowLocked(tbl, ID, hWnd) > -1 Then
'        If Len(Nz(txLog)) Then SetFields "SELECT dd, uu, tbl, tp, id, log FROM lg WHERE dd=0", "dd", Now(), Base.UID, tbl, tp, ID, txLog
'        SetFields "SELECT dd,uu," & PrimaryKey("" & tbl) & " FROM " & tbl & " WHERE " & PrimaryKey("" & tbl) & "=" & ID, "", Now(), Base.UID
'        'Debug.Print "LogRecord=", tp
'    End If
'End If
'End Sub
'
'
'Property Get PrimaryKey(tbl As String) As String
'Dim res
'If ConState = 1 Then
'res = mSchemaColumns.Param(tbl)
'If IsNull(res) Then
'    res = XConnection.OpenSchema(28, Array(Empty, Empty, "" & tbl)).Fields(3)
'    mSchemaColumns.Param(tbl) = Nz(res)
'End If
'PrimaryKey = res
'End If
'End Property
''============ RECORD LOCKING =========================================




' ============== XYXYXYXYXYXY ==================
'Public Function LineFromPos&(h, p)
'Dim hWnd&, pos&
'hWnd = h: pos = p
'If hWnd = 0 Then Exit Function
'LineFromPos = 1 + CallOldWindowProc(hWnd, EM_EXLINEFROMCHAR, 0, pos) 'EM_EXLINEFROMCHAR=&H400+54
'End Function

'Public Function CurrentPos&(h)
'Dim hWnd&, pos As CHARRANGE
'hWnd = h
'If hWnd = 0 Then Exit Function
'CallOldWindowProc hWnd, EM_EXGETSEL, 0, VarPtr(pos) 'EM_EXGETSEL
'CurrentPos = pos.cpMin
'End Function

'Function ScrollPos(ByVal hwnd, ByVal nBar)
'ScrollPos = GetScrollPos(0& + hwnd, 0& + nBar)
'End Function

Property Get ScrollPos(ByVal hWnd, ByVal nBar, Optional ByVal iMember = 0)
Dim si As SCROLLINFO: si.cbSize = Len(si): si.fMask = SIF_RANGE Or SIF_POS Or SIF_PAGE
GetScrollInfo hWnd, nBar, si
Dim ar: ar = Array(si.npos, si.nMin, si.nMax, si.nPage)
If iMember > -1 Then ScrollPos = ar(iMember And 3) Else ScrollPos = ar
'ScrollPos = GetScrollPos(0& + hWnd, 0& + nBar)
End Property
Property Let ScrollPos(ByVal hWnd, ByVal nBar, Optional ByVal iMember = 0, v)
'Dim si As SCROLLINFO: si.cbSize = Len(si): si.fMask = SIF_RANGE Or SIF_POS Or SIF_PAGE
'GetScrollInfo hWnd, nBar, si
'Dim ar: ar = Array(si.npos, si.nMin, si.nMax, si.nPage)
'If iMember > -1 Then ar(iMember And 3) = v Else ar = v
'si.npos = ar(0): si.nMin = ar(1): si.nMax = ar(2): si.nPage = ar(3)
'SetScrollInfo hWnd, nBar, si, 1
SetScrollPos 0& + hWnd, 0& + nBar, 0& + v, 1
End Property


'Function SBInfo(hwnd&, nBar&, Pos0Min1Max2&) As Long
'Dim si As SCROLLINFO, res&
'si.cbSize = LenB(si)
'si.fMask = SIF_RANGE Or SIF_POS
'GetScrollInfo hwnd, nBar, si
'If Pos0Min1Max2 = 0 Then
'    res = si.npos
'ElseIf Pos0Min1Max2 = 1 Then
'    res = si.nMin
'ElseIf Pos0Min1Max2 = 2 Then
'    res = si.nMax
'End If
'SBInfo = res
'End Function
' ============== XYXYXYXYXYXY ==================


'Function LogAR(hdr, ar0, ar1)
'Dim i&, n&, s$
'If VarType(hdr) And VarType(ar0) And VarType(ar1) And vbArray Then
'    n = UBound(hdr)
'    For i = 0 To n
'    If Nz(hdr(i)) <> "" Then If Nz(ar0(i)) <> Nz(ar1(i)) Then s = s & hdr(i) & ": {" & Nz(ar0(i), "Null") & "}->{" & Nz(ar1(i), "Null") & "}" & vbCrLf
'    Next
'End If
'LogAR = s
'End Function

Public Property Let MidVB(ByRef sString, nStart, nWidth, sValue)
On Error Resume Next
Mid(sString, nStart, nWidth) = sValue
End Property

Public Function xMulDiv(nNumber, nNumerator, nDenominator)
On Error Resume Next
xMulDiv = MulDiv(0& + nNumber, 0& + nNumerator, 0& + nDenominator)
End Function
Public Function CharCount(ByVal sString, ByVal sChar) As Long
  Dim lPos As Long
  On Error Resume Next
  sString = S_(sString): sChar = S_(sChar)
  If Len(sString) = 0 Or Len(sChar) = 0 Then Exit Function
  lPos = 1 - Len(sChar)
  Do
    lPos = InStr(lPos + Len(sChar), sString, sChar, vbBinaryCompare)
    If lPos > 0 Then CharCount = CharCount + 1
  Loop Until lPos = 0
End Function

Function TextBlockSize(ByVal txt, Optional ByVal returnWidth = 0, Optional ByVal sFont$ = "") As Long
Dim rc As RECT
Dim oldFont&, hDC&, lf As LOGFONT, hFont&
CreateMemDC hDC, 1, 1
On Error Resume Next
'lf = ncm.lfMessageFont
lf = mncm.lfMenuFont
hFont = GlobalFontSRC(lf, sFont, 1)
oldFont = SelectObject(hDC, hFont)
returnWidth = L_(returnWidth)
DrawText hDC, txt, Len(txt), rc, DT_CALCRECT Or DT_EXPANDTABS  ' Вычисляем размер текста
If rc.Right > Abs(returnWidth) Then
    rc.Right = Abs(returnWidth)
    Call DrawText(hDC, txt, Len(txt), rc, DT_WORDBREAK Or DT_EDITCONTROL Or DT_CALCRECT Or DT_EXPANDTABS)  ' Вычисляем размер текста
End If
SelectObject hDC, oldFont
CreateMemDC hDC, 0, 0

TextBlockSize = IIf(returnWidth > 0, rc.Right, rc.Bottom)
End Function


Function GDIP_ThumbToJPG(sourcefile, destfile, Optional maxWidth = 0, Optional maxHeight = 0, Optional quality = 80)
GDIP_ThumbToJPG = GDIP_ThumbImage(S_(sourcefile), L_(maxWidth), L_(maxHeight), S_(destfile), quality)
End Function

'*****************************************************************


Function IsInIDE() As Boolean
IsInIDE = modMain.IsInIDE
End Function
Function ShiftState()
ShiftState = GetShiftState
End Function
Function KeyState(ByVal key, Optional ByVal mask = &H8000&)
KeyState = GetKeyState(Cast(key, vbLong)) And Cast(mask, vbLong)
End Function
Function InsertState(ByVal hWnd) As Long
InsertState = GetProp(0& + hWnd, "ins") = 0
End Function
'=================  RUN AND WAIT   ===============
'Function RunAs(ByVal sUser, ByVal sPassword, ByVal sCommandLine, Optional ByVal nShowWindow = 5, Optional ByVal sDesktop As String)
'
''    Const LOGON_WITH_PROFILE = &H1&
''    Const LOGON_NETCREDENTIALS_ONLY = &H2&
'
''    Const CREATE_DEFAULT_ERROR_MODE = &H4000000
''    Const CREATE_NEW_CONSOLE = &H10&
''    Const CREATE_NEW_PROCESS_GROUP = &H200&
''    Const CREATE_SEPARATE_WOW_VDM = &H800&
''    Const CREATE_SUSPENDED = &H4&
''    Const CREATE_UNICODE_ENVIRONMENT = &H400&
''    Const ABOVE_NORMAL_PRIORITY_CLASS = &H8000&
''    Const BELOW_NORMAL_PRIORITY_CLASS = &H4000&
''    Const HIGH_PRIORITY_CLASS = &H80&
''    Const IDLE_PRIORITY_CLASS = &H40&
''    Const NORMAL_PRIORITY_CLASS = &H20&
''    Const REALTIME_PRIORITY_CLASS = &H100&
'
'    Dim lpUsername As String, lpDomain As String, lpPassword As String, lpApplicationName As String
'    Dim lpCommandLine As String, lpCurrentDirectory As String
'    Dim StartInfo As STARTUPINFO, ProcessInfo As PROCESS_INFORMATION
'
'    lpUsername = NzS(sUser, "ADMIN")
'
'    If CreateObject("WScript.Network").UserName = lpUsername Then
'        RunAs = ShellVB(sCommandLine, nShowWindow)
'        Exit Function
'    End If
'    lpDomain = ""
'    lpPassword = NzS(sPassword, "")
'    lpApplicationName = vbNullString '"C:\WINNT\NOTEPAD.EXE"
'    lpCommandLine = Nz(sCommandLine, vbNullString) 'use the same as lpApplicationName
'
'If Len(lpCommandLine) = 0 Then Exit Function
'
'    lpCurrentDirectory = vbNullString 'use standard directory
'
'    StartInfo.wShowWindow = aVal(nShowWindow)
'    StartInfo.cb = LenB(StartInfo) 'initialize structure
'    StartInfo.dwFlags = 2&
'    StartInfo.lpDesktop = sDesktop '"WinSta0\Default"
'    'CreateProcessWithLogon StrPtr(lpUsername), StrPtr(lpDomain), StrPtr(lpPassword), LOGON_WITH_PROFILE, StrPtr(lpApplicationName), StrPtr(lpCommandLine), CREATE_DEFAULT_ERROR_MODE Or CREATE_NEW_CONSOLE Or CREATE_NEW_PROCESS_GROUP, ByVal 0&, StrPtr(lpCurrentDirectory), StartInfo, ProcessInfo
'    CreateProcessWithLogon StrPtr(lpUsername), StrPtr(lpDomain), StrPtr(lpPassword), 0, StrPtr(lpApplicationName), StrPtr(lpCommandLine), &H4000210, ByVal 0&, StrPtr(lpCurrentDirectory), StartInfo, ProcessInfo
'
'    RunAs = ProcessInfo.dwProcessId
'    'CloseHandle ProcessInfo.hThread 'close the handle to the main thread, since we don't use it
'    'CloseHandle ProcessInfo.hProcess 'close the handle to the process, since we don't use it
'    'note that closing the handles of the main thread and the process do not terminate the process
'    'unload this application
'
'End Function


'Public Function RunAndWait(sComLine As String, sStartInFolder As String, nShowFlag As Long) As Boolean
'Dim si As STARTUPINFO
'Dim pi As PROCESS_INFORMATION
'Dim zz As Long
'si.wShowWindow = nShowFlag
'si.dwFlags = &H1
'If CreateProcess(vbNullString, sComLine, ByVal 0, ByVal 0, False, 0, ByVal 0, sStartInFolder, si, pi) Then
'    WaitForSingleObject pi.hProcess, &HFFFF
'    GetExitCodeProcess pi.hProcess, zz
'    TerminateProcess pi.hProcess, zz
'    CloseHandle pi.hProcess
'    RunAndWait = True
'    Exit Function
'End If
'RunAndWait = False
'End Function

' Call this sub to execute and capture a console app.
' Ex: Call ExecAndCapture("ping localhost", Text1)
Public Function ExecToStream(ByVal sCommandLine As String, Optional ByVal sStartInFolder As String = vbNullString, Optional ByVal sDumyCookie$) As Stream
Const BUFSIZE  As Long = 8192
Dim hPipeRead&, hPipeWrite&
Dim sa As SECURITY_ATTRIBUTES, si As STARTUPINFO, pi As PROCESS_INFORMATION
Dim st As New Stream: st.Type = adTypeBinary: st.Open
Dim baOutput(BUFSIZE) As Byte
Dim lBytesRead&
Dim buf() As Byte
With sa
    .nLength = Len(sa)
    .bInheritHandle = 1    ' get inheritable pipe handles
End With 'SA
If CreatePipe(hPipeRead, hPipeWrite, sa, 0) = 0 Then Exit Function
With si
    .cb = Len(si)
    .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
    .wShowWindow = SW_HIDE          ' hide the window
    .hStdOutput = hPipeWrite
    .hStdError = hPipeWrite
End With 'SI
If CreateProcess(vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, sStartInFolder, si, pi) Then
    Call CloseHandle(hPipeWrite)
    Call CloseHandle(pi.hThread)
    If Len(sDumyCookie) Then
        RaiseEvent DumyEvent(sDumyCookie, Array("Start", st))
        If st.Type = adTypeText Then Exit Function
    End If
    hPipeWrite = 0
    Do
        If apiReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then  'Здесь зависает
            If Len(sDumyCookie) Then RaiseEvent DumyEvent(sDumyCookie, Array("End", st))
            st.Position = 0: st.Type = adTypeText
            st.charset = "_autodetect_all"
            Set ExecToStream = st
            Exit Do
        End If
        ReDim Preserve buf(lBytesRead - 1)
        CopyMemory buf(0), baOutput(0), lBytesRead
        st.Write buf
        DoEvents
        If Len(sDumyCookie) Then
            RaiseEvent DumyEvent(sDumyCookie, Array("Progress", st))
            If st.Type = adTypeText Then Exit Do
        End If
    Loop
    Call CloseHandle(pi.hProcess)
End If
' To make sure...
Call CloseHandle(hPipeRead)
Call CloseHandle(hPipeWrite)
End Function


'=================  RUN AND WAIT   ===============

'Function RWBits(ByVal longVal, ByVal nBitIndex, Optional ByVal bWrite = Null) As Long
''Чтение/ Запись указанных битов
''bWrite=Null,Empty    Чтение          return =  longVal.BIT(nBitIndex) as boolean
''bWrite=0                  Запись 0        longVal.BIT(nBitIndex)=0, return = longVal as long
''bWrite<>0                Запись 1        longVal.BIT(nBitIndex)=1, return = longVal as long
'Dim v&, i&, m&
'v = modMain.Cast(longVal, vbLong)
'RWBits = v
'i = modMain.Cast(nBitIndex, vbLong)
'If i < 0 Or i > 31 Then Exit Function
'm = 2 ^ i
'If VarType(bWrite) > vbNull Then 'Write and return result
'    RWBits = (v And (Not m)) Or (m And (aVal(bWrite) <> 0))
'Else 'Read bit
'    RWBits = (v And m) = m
'End If
'End Function

Function UDPList(Optional ByVal nPID = -1, Optional ByVal nRet = 0) As String
'return json  pid\port=ip   filter by PID
'nRet: 1=json 0=List 2=Matrix
If nPID = -1 Then nPID = GetCurrentProcessId
UDPList = modSocket.UDPList(aVal(nPID), aVal(nRet))
End Function

'Function LANList(Optional ByVal sCommentSeparator = "")
'LANList = modSocket.LANList(Nz(sCommentSeparator))
'End Function

Function Dumy(ByVal Cookie, ParamArray args())
Dim ar: ar = args
If IsNull(Cookie) Then
    If UBound(ar) > -1 Then Dumy = ar(0)
Else
    RaiseEvent DumyEvent(Cookie, ar)
    If IsNull(Cookie) Then Dumy = ar
End If
End Function

Function SecurityPrivilege(ByVal nm, Optional ByVal v = -1)
SecurityPrivilege = SEPrivilege(S_(nm), L_(v))
End Function

Public Property Get Port&()
Dim x As xControl
For Each x In xMainWnd.Controls
Port = x.SocketPort
If Port Then Exit For
Next
End Property

Function Dispatch(obj) As Object
On Error Resume Next
Dim uu As UUID
'If TypeName(obj) = "CParam" Then uu.Data1 = 999
'If TypeName(obj) = "CParam" Then
esos.Add obj, "o" & ObjPtr(obj)
Set Dispatch = CreateEventSinkObj(uu, obj, Nothing)
End Function


Function ToRTF(ByVal s)
Dim i&, c$, b() As Byte
Dim st As New ADODB.Stream: st.Open
b = StrConv(Replace(S_(s), vbCrLf, vbCr), vbFromUnicode)
For i = 0 To UBound(b)
Select Case b(i)
Case 9: c = "\tab"
Case 10: c = "\par"
Case 11: c = "\line"
Case 12: c = "\page"
Case 13: c = "\par"
Case Is < 31, Is > 127: c = "\'" & Right("0" & LCase(Hex(b(i))), 2)
Case 92, 123, 125: c = "\" & Chr(b(i))
Case Else: c = Chr(b(i))
End Select
st.WriteText c
Next
st.Position = 0
ToRTF = st.ReadText
End Function



Sub InetSetCookie(ByVal lpszUrl, ByVal lpszCookieName, ByVal lpszCookieData)
InternetSetCookie S_(lpszUrl), S_(lpszCookieName), S_(lpszCookieData)
End Sub

'Function InetGetCookie(ByVal lpszUrl, ByVal lpszCookieName)
'Dim lpCookieData As String, lpdwSize As Long
'lpszUrl = Nz(lpszUrl): lpszCookieName = Nz(lpszCookieName)
'InternetGetCookie lpszUrl, lpszCookieName, lpCookieData, lpdwSize
'If lpdwSize Then
'    lpCookieData = String(lpdwSize, 0)
'    InternetGetCookie lpszUrl, lpszCookieName, lpCookieData, lpdwSize
'    InetGetCookie = Left(lpCookieData, lpdwSize)
'End If
''Debug.Print InetGetCookie
''InetGetCookie = VBScript("", "Function UN(a):UN=UnEscape(a):End Function").Run("UN", InetGetCookie)
''Debug.Print InetGetCookie
'End Function

Sub InetClearCache(ByVal lpszUrlName)
DeleteUrlCacheEntry S_(lpszUrlName)
End Sub

'Function OLEProgID(ByVal clsid_str)
'Dim u As UUID
'CLSIDFromString Nz(clsid_str), u
'Dim p&, s$, s0$
'On Error Resume Next
'p = ProgIDFromCLSID(u)
'OLEProgID = StringWFromPointer(p)
'End Function


Property Get XMDB(Optional ByVal pname = False)
On Error Resume Next
If VarType(pname) = vbBoolean Then
    If pname Then 'GET connection
        Set XMDB = xmdb_conn
    Else 'Get XMDB STATE
        XMDB = False
        XMDB = xmdb_conn.State = 1
    End If
ElseIf Not xmdb_conn Is Nothing Then 'Get Value
    'pname = Replace(Nz(pname), Chr(39), "") 'NameID or ParamPath
    pname = S_(pname)
    If xmdb_conn.State = 1 And Len(pname) > 0 Then
        Dim ar, v, rs As New Recordset
        ar = Split(pname, "\")
        rs.Open "SELECT pv  FROM xmdb WHERE id=" & Quot(ar(0), 39), xmdb_conn, adOpenKeyset, adLockOptimistic
        If Not (rs.BOF And rs.EOF) Then 'READ
            v = rs.Fields(0)
            If VarType(v) = vbString Then
                If gIsCompressed(v) Then v = gDecompress(v)
                v = modMain.ToVariant(v)
                If UBound(ar) > 0 Then v = CParam(v)(RightStr(pname, "\", 1))
            End If
            XMDB = v
        End If
    End If
End If
End Property

Property Let XMDB(Optional ByVal pname = False, ByVal v)
On Error Resume Next
If VarType(pname) = vbBoolean Then 'Let XMDB
    Set xmdb_conn = New Connection
    If Len(S_(v)) Then xmdb_conn.Open JETProvider & v
ElseIf Not xmdb_conn Is Nothing Then
    'pname = Replace(Nz(pname), Chr(39), "")
    pname = S_(pname)
    If xmdb_conn.State = 1 And Len(pname) > 0 Then
        Dim pc, ds, ar, rs As New Recordset
        ar = Split(pname, "\")
        rs.Open "SELECT * FROM xmdb WHERE id=" & Quot(ar(0), 39), xmdb_conn, adOpenKeyset, adLockOptimistic
        Dim pv, cp As New CParam
        If rs.BOF And rs.EOF Then 'AddNew
            If VarType(v) > vbNull Then
                If UBound(ar) > 0 Then cp(pname) = v: pv = cp.Source Else pv = v
                If IsArray(pv) Then cp.Value = pv
                GoSub pc_ds
                pv = modMain.ToString(pv)
                If Len(pv) >= MAX_FIELDINDEX Then pv = gCompress(pv)
                rs.AddNew Array("id", "pv", "tmc", "pc", "ds"), Array(ar(0), pv, Now, pc, ds)
            End If
        Else
            If VarType(v) = vbNull And UBound(ar) = 0 Then
                rs.Delete
            Else
                If UBound(ar) > 0 Then 'PathParam Update
                    cp.Value = modMain.ToVariant(gDecompress(rs.Fields("pv")))
                    cp(RightStr(pname, "\", 1)) = v
                    v = cp.Source
                Else
                    cp.Value = v
                End If
                pv = ToString(v)
                GoSub pc_ds
                If Len(pv) >= MAX_FIELDINDEX Then pv = gCompress(pv)
                rs.UpDate Array("id", "pv", "tmm", "pc", "ds"), Array(ar(0), pv, Now, pc, ds)
            End If
        End If
        If Err Then Debug.Print Err.Description
    End If
End If
Exit Property

pc_ds:
    pc = cp.ParamCount
    If pc Then ds = Left(S_(cp("Description")), 255) Else ds = rs.Fields("ds")
Return
End Property

Sub LoadEngineSettings()
Param(App.ProductName) = Null
Dim Item, a
For Each a In Split("#,", ",") 'EXE.INI + CMD.INI
For Each Item In Split(INISetting(a & App.ProductName, ""), Chr(0)) 'INI  ->> PARAM
Param(App.ProductName & "\" & Item) = INISetting(a & App.ProductName, Item)
Next
Next
'Param(App.ProductName & "\CaptionHeight") = mncm.iCaptionHeight
'Param(App.ProductName & "\MenuHeight") = mncm.iMenuHeight
'Param(App.ProductName & "\BorderWidth") = mncm.iBorderWidth
End Sub

Public Property Get VB6()
'If mVB6 Is Nothing Then
Dim mVB6 As New Collection
'    Set mVB6 = New Collection
    mVB6.Add CDataObject("Printer", Printer)
    mVB6.Add CDataObject("Printers", Printers)
    mVB6.Add CDataObject("Screen", Screen)
'End If
Set VB6 = Dispatch(mVB6)


End Property

Property Get xMain() As CMain
Set xMain = Me
End Property

'Function PrinterPaper(Optional ByVal sPrinterName) As CMatrix
'Dim i&, n&, mx As New CMatrix, tmp$, s$
'mx.OpenRows 0, Replace("id&,nm$64$,w&,h&", ",", vbTab), 3
'Set PrinterPaper = mx
'sPrinterName = NzS(sPrinterName, Printer.DeviceName)
'n = DeviceCapabilities(sPrinterName, "", 16, ByVal 0, 0)
'If n Then
'ReDim pNames(n - 1) As String
'tmp = String$(64 * n, 0)
'DeviceCapabilities sPrinterName, "", 16, ByVal tmp, 0
'ReDim pIDs(n - 1) As Integer
'DeviceCapabilities sPrinterName, "", 2, pIDs(0), 0
'ReDim pSizes(n - 1) As POINTAPI
'DeviceCapabilities sPrinterName, "", 3, pSizes(0), 0
'For i = 0 To n - 1
'pNames(i) = Trim(Replace(Mid$(tmp, 1 + 64 * i, 64), Chr(0), " "))
'mx.AppendRows10 Array(pIDs(i), pNames(i), 10 * pSizes(i).x, 10 * pSizes(i).y)
'Next
'End If
''Debug.Print mx.GetRows(0)
'End Function

Function PageSetupDLG(ByVal hWnd, ByVal psjson)
Dim psd As PAGESETUPDLG_struct
Dim ps, cp As CParam
Dim n&: n = 100
Set cp = CParam(psjson, "json"): Set ps = Dispatch(cp)
With psd

If IsWindow(Abs(hWnd)) Then .hWndOwner = Abs(hWnd)

'Private Const PSD_MINMARGINS As Long = &H1
'Private Const PSD_MARGINS As Long = &H2
'Private Const PSD_INTHOUSANDTHSOFINCHES As Long = &H4
'Private Const PSD_INHUNDREDTHSOFMILLIMETERS As Long = &H8
'Private Const PSD_DISABLEMARGINS As Long = &H10
'Private Const PSD_DISABLEPRINTER As Long = &H20
'Private Const PSD_NOWARNING As Long = &H80
'Private Const PSD_DISABLEORIENTATION As Long = &H100
'Private Const PSD_DISABLEPAPER As Long = &H200
'Private Const PSD_RETURNDEFAULT As Long = &H400
'Private Const PSD_SHOWHELP As Long = &H800
'Private Const PSD_ENABLEPAGESETUPHOOK As Long = &H2000
'Private Const PSD_ENABLEPAGEPAINTHOOK As Long = &H40000
'Private Const PSD_ENABLEPAGESETUPTEMPLATE As Long = &H8000
'Private Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE As Long = &H20000
'Private Const PSD_DISABLEPAGEPAINTING As Long = &H80000
'Private Const PSD_NONETWORKBUTTON As Long = &H200000
If Printers.count = 0 Then Exit Function
.lStructSize = Len(psd)
Dim dmPtr&, dm As DEVMODE
.Flags = &H48A 'PSD_NOWARNING+PSD_RETURNDEFAULT +PSD_MARGINS +PSD_INHUNDREDTHSOFMILLIMETERS
PageSetupDialog psd 'Get DEFAULT DEVMODE
GetMem4 ByVal psd.hDevMode, dmPtr: CopyMemory dm, ByVal dmPtr, LenB(dm)
'Debug.Print Trim(Replace(StrConv(dm.dmDeviceName, vbUnicode), Chr(0), " "))
dm.dmOrientation = IF0(ps.Orientation, dm.dmOrientation)
dm.dmPaperSize = IF0(ps.PaperSize, dm.dmPaperSize)
.rtMargin.Left = IF0(aVal(ps.Left) * n, .rtMargin.Left)
.rtMargin.Top = IF0(aVal(ps.Top) * n, .rtMargin.Top)
.rtMargin.Right = IF0(aVal(ps.Right) * n, .rtMargin.Right)
.rtMargin.Bottom = IF0(aVal(ps.Bottom) * n, .rtMargin.Bottom)
If hWnd < 0 Then .Flags = &H402 Else .Flags = &H22
CopyMemory ByVal dmPtr, dm, LenB(dm) 'Modify DEVMODE
If PageSetupDialog(psd) Or hWnd < 0 Then
    CopyMemory dm, ByVal dmPtr, LenB(dm)
    ps.Width = .ptPaperSize.x \ n
    ps.Height = .ptPaperSize.Y \ n
    ps.PaperSize = dm.dmPaperSize
    ps.Orientation = dm.dmOrientation
    ps.Left = .rtMargin.Left \ n
    ps.Top = .rtMargin.Top \ n
    ps.Right = .rtMargin.Right \ n
    ps.Bottom = .rtMargin.Bottom \ n
    PageSetupDLG = cp.json
    'Debug.Print PageSetup
End If
End With
End Function

Function PrintRTFDLG(ByVal hWnd, ByVal Page_json, Optional ByVal Print_json) ', Optional hEMFs) 'return pages_ends array (end1,end2,..,endN) + PRINT + DIALOG
Dim pages 'RESULT
Dim z As Object, cp As CParam
Set cp = CParam(Page_json, "json"): Set z = Dispatch(cp)
'page=PaperSize,Orientation,Left,Top,Right,Bottom[,Width,Height]
cp.AddParams CParam(Print_json, "json")
'printset: Printer=DeviceName,PrintNow=0,1,2,Name=Title,FromPage,ToPage,Copies

Dim fr As FORMATRANGE, s$
fr.chrg.cpMax = CallOldWindowProc(hWnd, EM_FORMATRANGE, 0, 0)
If fr.chrg.cpMax > 0 Or hWnd = 0 Then

Dim nPrint&: nPrint = aVal(z.PrintNow)
Dim pd As PRINTDLG, hDC& 'PRINT DC

If Printers.count Then
    On Error Resume Next
    s = Printers(1).DeviceName
    If Not Printer Is Nothing Then s = Printer.DeviceName
    s = Nz(z.Printer, s)
End If


If Len(s) Then

Dim dmPtr&, dm As DEVMODE, dmCopies%
pd.lStructSize = Len(pd)

Dim nSize&, lHandle&
nSize = DocumentProperties(0, 0, s, 0, 0, 0)
If nSize > 1 Then
    ReDim buf(nSize - 1) As Byte
    DocumentProperties 0, 0, s, buf(0), 0, 2
    CopyMemory dm, buf(0), LenB(dm)  'copy to dm
    lHandle = GlobalAlloc(2, nSize)
    dmPtr = GlobalLock(lHandle)
    CopyMemory ByVal dmPtr, buf(0), nSize
    GlobalUnlock lHandle
    pd.hDevMode = dmPtr
    pd.hDevNames = 0
Else 'Printer not found = GET PD_RETURNDEFAULT
    pd.Flags = PD_RETURNDEFAULT
    PrintDialog pd
    GetMem4 ByVal pd.hDevMode, dmPtr
End If


dmCopies = Cast(IF0(z.Copies, dm.dmCopies), 2)
dm.dmPaperSize = Cast(IF0(z.PaperSize, dm.dmPaperSize), 2)
dm.dmOrientation = Cast(IF0(z.Orientation, dm.dmOrientation), 2)
dm.dmCopies = dmCopies
pd.nCopies = dmCopies
pd.nFromPage = Cast(IF0(aVal(z.FromPage), 1), 2)
pd.nToPage = Cast(IF0(aVal(z.ToPage), 999), 2)
pd.nMinPage = 1: pd.nMaxPage = pd.nToPage
dm.dmFields = &HFFF3&

If nPrint = 2 Then 'Show PrintDialog
    pd.hWndOwner = hWnd
    pd.Flags = PD_RETURNDC Or PD_NOSELECTION Or PD_HIDEPRINTTOFILE Or PD_USEDEVMODECOPIES
    If lHandle Then CopyMemory ByVal dmPtr, dm, LenB(dm) 'Modify DEVMODE
    
    If PrintDialog(pd) Then
'        CopyMemory dm, ByVal dmPtr, LenB(dm) 'READ DEVMODE  for nCopies
        hDC = pd.hDC
        If (pd.Flags And PD_PAGENUMS) = 0 Then pd.nFromPage = 1: pd.nToPage = 999
    End If
Else
    hDC = CreateDC(vbNullChar, s, vbNullChar, dm)
End If

Else 'Printers.Count=0
    hDC = GetDC(0)
End If

'If hDC = 0 And nPrint = 0 Then If Printers.count = 0 Then hDC = GetDC(0)


If hDC Then
Dim rc As RECT, rcPage As RECT
'Set the rendering rectangle to the pintable area of the page.
'Debug.Print GetDeviceCaps(hDC, PHYSICALWIDTH), GetDeviceCaps(hDC, LOGPIXELSX), dm.dmPrintQuality
rcPage.Right = MulDiv(GetDeviceCaps(hDC, PHYSICALWIDTH), 1440, GetDeviceCaps(hDC, LOGPIXELSX))
If rcPage.Right = 0 Then rcPage.Right = MulDiv(IF0(aVal(z.Width), 210), 14400, 254)
rcPage.Bottom = MulDiv(GetDeviceCaps(hDC, PHYSICALHEIGHT), 1440, GetDeviceCaps(hDC, LOGPIXELSY))
If rcPage.Bottom = 0 Then rcPage.Bottom = MulDiv(IF0(aVal(z.Height), 297), 14400, 254)

'Set the rendering rectangle to the pintable area of the page.
z.Left = aVal(z.Left): z.Right = aVal(z.Right): z.Top = aVal(z.Top): z.Bottom = aVal(z.Bottom)
rc.Left = MulDiv(z.Left * 100, 1440, 2540) 'mm-twip
rc.Right = rcPage.Right - MulDiv(z.Right * 100, 1440, 2540)
rc.Top = MulDiv(z.Top * 100, 1440, 2540)
rc.Bottom = rcPage.Bottom - MulDiv(z.Bottom * 100, 1440, 2540)


Dim di As DOCINFO
If nPrint Then
    di.lpszDocName = Nz(z.Name, "Document " & Hex(hWnd))
    di.cbSize = Len(di)
    'di.lpszOutput = Nz(z.PrintToFile)
    If Right$(di.lpszDocName, 4) = ".xps" Then di.lpszOutput = di.lpszDocName
    
    If StartDoc(hDC, di) < 1 Then nPrint = -nPrint
End If

fr.hDCTarget = hDC
Dim cpMin&, np&, nb& ' As Boolean
fr.rcPage = rcPage
Dim h&, hdc0&

hdc0 = GetDC(0)
'Dim bMeta As Boolean
'bMeta = aVal(hEMFs)
'hEMFs = Array()
Const hx = 150
Do While (fr.chrg.cpMin < fr.chrg.cpMax And np < 100)
    np = np + 1
'    fr.hDC = 0 'GetDC(0)
'    If bMeta Then 'PrintPage to METAFILE
'        Dim ercPage As RECT
'        ercPage.Right = aVal(z.Width) * 100
'        ercPage.Bottom = aVal(z.Height) * 100
'        fr.hDC = CreateEnhMetaFile(hDC, vbNullString, ercPage, "")  ' Nz(z.Name, "Document " & Hex(hWnd)))
'        FillRect fr.hDC, rcPage, GlobalBrush(&HFFFFFF, 0)
'    End If
    nb = IIf(nPrint > 0 And pd.nFromPage <= np And pd.nToPage >= np, 1, 0)
    If nb Then fr.hDC = hDC:  StartPage hDC Else fr.hDC = hdc0
    'If nb Then FillRect fr.hDC, fr.rcPage, GlobalBrush(&HF0F0F0, 0)
    h = 0
    Do
        fr.rc = rc: fr.rc.Bottom = rc.Bottom - h * hx
        cpMin = CallOldWindowProc(hWnd, EM_FORMATRANGE, 0, VarPtr(fr))
        'Debug.Print h, "rc start=", rc.Bottom - h * 180, "max=" & rc.Bottom, "res=" & fr.rc.Bottom, fr.chrg.cpMin, cpMin
        If fr.rc.Bottom - rc.Bottom > hx Then
            h = h + 1
            Call CallOldWindowProc(hWnd, EM_FORMATRANGE, 0, 0)
        End If
    Loop While fr.rc.Bottom - rc.Bottom > hx And 2 * h * hx < rc.Bottom
    fr.rc = rc: fr.rc.Bottom = rc.Bottom - h * hx
    
    cpMin = CallOldWindowProc(hWnd, EM_FORMATRANGE, nb, VarPtr(fr))
    

    If nb Then EndPage hDC
'    If bMeta Then gAddIndex hEMFs, CloseEnhMetaFile(fr.hDC)  'METAFILE
    gAddIndex pages, cpMin

    If (cpMin <= fr.chrg.cpMin) Then Exit Do
    fr.chrg.cpMin = cpMin
Loop
'If nPrint > 0 Then ReleaseDC hWnd, hdc0
Call CallOldWindowProc(hWnd, EM_FORMATRANGE, 0, 0)

If nPrint Then EndDoc hDC
If Abs(nPrint) < 2 Then If Printers.count Then DeleteDC hDC


End If 'HDC<>0
End If 'TEXTLEN>0
If lHandle Then GlobalFree lHandle
If Not IsArray(pages) Then gAddIndex pages, -1

PrintRTFDLG = pages
End Function


Function ADOStream(Optional src = 0, Optional toType = 0) As Stream
On Error Resume Next
Set ADOStream = New Stream: ADOStream.Open
'Select Case TypeName(src)
Select Case VarType(src)
Case VT_UNKNOWN '"Unknown"
    Dim ist As IStream: Set ist = src
    Dim p As Currency
    p = ist.Seek(0, STREAM_SEEK_END)
    ist.Seek 0, STREAM_SEEK_SET
    ADOStream.Type = adTypeBinary
    ist.CopyTo ADOStream, p, 0, 0
Case vbString '"String"
    If Len(src) Then ADOStream.WriteText src
Case vbArray + vbByte, vbByte '"Byte()"
    ADOStream.Type = adTypeBinary
    If LenB(src) Then ADOStream.Write src
Case Else
End Select
ADOStream.Position = 0
If toType Then ADOStream.Type = toType
'Debug.Print ADODBStream.SIZE
'ADOStream.Position = 0
'ADOStream.Type = adTypeText
'ADOStream.charset = "_autodetect_all"
'Dim s$: s = ADOStream.ReadText
'Debug.Print "=" & s & "="
End Function

Sub DeleteGDI(str_gdi$)
While aVal(GlobalGDICount(str_gdi))
ClearGDI str_gdi
Wend
End Sub

Function Bytes(Optional ByVal b = -999, Optional ByVal bset = -1)
Dim s$
On Error Resume Next
Select Case VarType(b)
Case vbNull: Erase m_bytes 'Erase BUFFER
Case vbArray + vbByte, vbString 'Return BUFFER  from b
    m_bytes = b
    's = b
    'm_bytes = StrConv(s, vbFromUnicode)
    Bytes = m_bytes
Case vbObject 'Return string
    s = m_bytes: Bytes = s
Case vbEmpty 'Return stringConv
    's = m_bytes
    Bytes = StrConv(m_bytes, vbFromUnicode)
Case vbInteger, vbLong, vbDouble
    If b = -999 Then 'Get ptr
        Bytes = VarPtr(m_bytes(0))
    ElseIf b = 0 Then 'Get BUFFER as string vbCrLf
        Bytes = Join2(m_bytes)
    ElseIf b < 0 Then 'ReSize BUFFER
        If UBound(m_bytes) > (-b - 1) Then ReDim Preserve m_bytes(-b - 1) Else ReDim m_bytes(-b - 1)
    ElseIf b <= UBound(m_bytes) Then 'Get ITEM
        If bset > -1 Then m_bytes(b - 1) = Cast(bset, vbLong) And 255
        Bytes = m_bytes(b - 1)
    End If
Case Else: Bytes = m_bytes 'Return BUFFER
End Select

'If b = -999 Then 'Get ptr
'    Bytes = VarPtr(m_bytes(0))
'ElseIf b = 0 Then 'Get buffer
'    Bytes = Join2(m_bytes)
'ElseIf b < 0 Then 'ReSize buffer
'    ReDim m_bytes(-b - 1)
'Else
'    Bytes = m_bytes(b - 1)
'End If
End Function


Function ResRestore(ByVal s$) As String
On Error Resume Next
Dim st As Stream: Set st = ADOStream(LoadResData(s, "CUSTOM"))
If st.Size = 0 Then Exit Function
s = modMain.MakeDir(App.Path & "\" & s)
ResRestore = s
If st.Size = LenFile(s) Then Exit Function
st.SaveToFile s, adSaveCreateOverWrite
End Function

Function ParseMenu(ByVal s, Optional ByVal nMode = 0)
'nMode=0    62>=30  60<=28  59;=27                       ESCAPE caption
'nMode=1    9262\>=30  9260\<=28  9259\;= 27    UNESCAPE menu
'nMode=2    30=62>  28=60<   27=59;                      UNESCAPE caption
'nMode=3    30=9262\>  28=9260\<   27=9259\;    ESCAPE to menu
s = Cast(s, vbString)
nMode = Cast(nMode, vbLong)
Dim ch, fnd$, rpl$
For Each ch In Array(59, 60, 62)
    If nMode > 1 Then fnd = Chr$(ch - 32): rpl = Chr$(ch) Else fnd = Chr$(ch): rpl = Chr$(ch - 32)
    If nMode And 1 Then If nMode > 1 Then rpl = Chr$(92) & rpl Else fnd = Chr$(92) & fnd
    s = Replace(s, fnd, rpl)
Next
ParseMenu = s
End Function


Function Layered(ByVal hWnd, Optional ByVal v = -1)
hWnd = Cast(hWnd, vbLong): v = Cast(v, vbLong)
Dim b As Byte
If v > -1 Then
    SetWindowLong hWnd, GWL_EXSTYLE, (Not WS_EX_LAYERED And GetWindowLong(hWnd, GWL_EXSTYLE)) Or IIf(v, WS_EX_LAYERED, 0)
    b = LOBYTE(Cast(v, vbInteger))
    Call SetLayeredWindowAttributes(hWnd, 0, b, 2&)
    UpdateWindow hWnd
End If
b = 0: GetLayeredWindowAttributes hWnd, 0, b, 2&
Layered = b
End Function

Function Flash(ByVal hWnd)
If hWnd = 1 Then KillTask CurrentProcessID Else FlashWindow hWnd, 1
End Function

Function VBStrConv(s, ByVal n)
VBStrConv = StrConv(s, n)
End Function

Function LogMsg(ByVal nColor, ByVal sText) As Long 'call from long time proc
'return false = CONTINUE 'return true = STOP
LogMsg = 1
'DebugPrint nColor, sText
If Not IsWindow(m_LogWnd) = 1 Then Exit Function
On Error Resume Next
LogMsg = aVal(hxControl(m_LogWnd).ASE.Run("Msg", nColor, sText))
End Function

Function LogWnd(Optional h = True) As Long
'setup log listener
If TypeName(h) = xs.sxControl Then
    m_LogWnd = h.hWnd
ElseIf VarType(h) = vbBoolean Then
    If IsWindow(m_LogWnd) = 0 Then m_LogWnd = 0
    LogWnd = m_LogWnd
Else
    m_LogWnd = aVal(h)
End If
End Function

Function Access(Optional ByVal v = Empty)
If VarType(v) Then gAccess = Cast(Nz(v, Base.UserRole), 3)
Access = gAccess
End Function

Function VarComp(v1, v2) As Long
VarComp = modMain.VarComp(v1, v2)
End Function

Function ViewClipFormat(ByVal buf, ByVal nFormat)
If VarType(buf) <> (vbArray + vbByte) Then Exit Function
Dim s$, nm$
nm = S_(nFormat)
If IsNum(nm) Then nm = NzS(ClipFormatName(nFormat), nFormat)
Select Case nm
Case "1": s = RTrimChar(StrConv(buf, vbUnicode))
Case "13": s = Cast(buf, vbString)
Case "FileName", "FileNameW": s = buf
Case "15"
    With ADOStream(buf, adTypeText)
        If .Size > 20 Then .Position = 20: s = Join(Split(.ReadText((.Size - 24) \ 2), Chr(0)), vbCrLf)
    End With
'Case "xControl"
'    Set ViewClipFormat = ObjFromStrPtr(buf)
'    Exit Function
Case Else
        s = StrConv(buf, vbUnicode)
End Select
ViewClipFormat = s
End Function

'Function StrView(ByVal s, ByVal n)
's = Nz(s): If Len(s) > n Then s = Left(s, n) & ".."
'StrView = s
'End Function

Function Min(a, b) ': On Error Resume Next: Min = IIf(a > b, b, a): End Function
On Error Resume Next: If a > b Then Min = b Else Min = a
End Function
Function Max(a, b) ': On Error Resume Next:  Max = IIf(a < b, b, a): End Function
On Error Resume Next: If a > b Then Max = a Else Max = b
End Function

'Function RichMenu(tx As xControl, Optional ByVal smenu)
''dim txh:txh=Array(tx.Send(1074) '/*EM_CANPASTE*/,tx.Send(198,,,1)/*EM_CANUNDO*/,tx.Send(1109,,,1)/*EM_CANREDO*/,Len(tx.SelText)/*SEL_LENGTH*/)
'On Error Resume Next
'Dim cc&: cc = IIf(Len(tx.SelText), 0, 3)
'Dim res, s$
's = "5;<Cut;D;" & cc & ";P;8;I;10><Copy;D;" & cc & ";P;9;I;11><Paste;D;" & IIf(tx.Send(1074), 0, 3) & ";P;10;I;12>"
''s = s & "<\s><Undo;P;148;D;" & IIf(tx.Send(198, , , 1), 0, 3) & ";I;13><Redo;P;149;D;" & IIf(tx.Send(1109, , , 1), 0, 3) & ";I;14>"
'If Left(smenu, 1) = "<" Then s = s & smenu
'res = PopupMenu(tx.hWnd, s)
'Dim r: Set r = tx.RTFSelection
'Select Case aVal(res)
'Case 10: r.Cut 0
'Case 11: If ShiftState Then ClipboardText = RangeRTF(r.Duplicate) Else r.Duplicate.Copy 0
'Case 12: r.Paste 0, 0
'Case 13: tx.RTFDocument.Undo 1
'Case 14: tx.RTFDocument.Redo 1
'End Select
'RichMenu = res
'End Function


Function ReplaceIcon(ByVal ii, ByVal lib, ByVal ID)
ReplaceIcon = pReplacemenu16Icon(aVal(ii), S_(lib), S_(ID))
End Function

Function UTCBias() As Currency ' LocalOffsetFromUTC in minutes
Dim TZI As TIME_ZONE_INFORMATION
GetTimeZoneInformation TZI
UTCBias = TZI.Bias
End Function


'Public Function FMsg(ByVal fs, ParamArray args()) As String
''SYSTEM ERROR DESCRIPTION =  FMsg(err.number)
''HMODULE ERROR DESCRIPTION =  FMsg(err.number,module path)
''FORMAT_MESSAGE = FMsg(format,arg1,arg2,...,argN)
'Dim i&, pbuf&, n&, dwFlag&, lpSource&
'fs = Nz(fs): n = Cast(fs, vbLong)
''Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
''Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
''Const FORMAT_MESSAGE_FROM_STRING As Long = &H400
''Const FORMAT_MESSAGE_FROM_HMODULE As Long = &H800
''Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000&
''Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Long = &H2000
'If fs = n Then
'    If UBound(args) > -1 Then
'        dwFlag = &H800  'FORMAT_MESSAGE_FROM_HMODULE
'        lpSource = LoadLibrary(Nz(args(0))): pbuf = 1
'    Else
'        dwFlag = &H1000&  'FORMAT_MESSAGE_FROM_SYSTEM
'    End If
'Else
'    dwFlag = &H400 'FORMAT_MESSAGE_FROM_STRING
'    lpSource = StrPtr(fs)
'End If
'
'If UBound(args) > pbuf - 1 Then
'    dwFlag = dwFlag Or &H2000 'FORMAT_MESSAGE_ARGUMENT_ARRAY
'    Dim sar() As String, ar() As Long, par&
'    ReDim ar(UBound(args) - pbuf), sar(UBound(args) - pbuf)
'    For i = 0 To UBound(ar): sar(i) = Nz(args(i + pbuf)): ar(i) = StrPtr(sar(i)): Next
'    par = VarPtr(ar(0)): pbuf = 0
'Else
'    dwFlag = dwFlag Or &H200 'FORMAT_MESSAGE_IGNORE_INSERTS
'End If
'
'i = FormatMessage(dwFlag Or &H100, lpSource, n, 0&, pbuf, 1, par)
'If lpSource Then FreeLibrary lpSource
'If i > 0 Then
'    FMsg = SysAllocString(pbuf)
'    If Right(FMsg, 2) = vbCrLf Then FMsg = Left(FMsg, Len(FMsg) - 2) 'Replace$(fmsg, vbCrLf, vbNullString)
'    LocalFree pbuf
'  '  Debug.Print NzS(Hex(n), , "(0x", ") ") & n, FMsg
'    'Debug.Print FMsg
'End If
'End Function


Public Function Hash(ByVal src As String, Optional ByVal calg As Long = &H8003, Optional ByVal bFile = 0) As String
'Const CALG_MD2 = &H8001&   'MD2 hashing algorithm. This algorithm is supported by the Microsoft Base Cryptographic Provider.
'Const CALG_MD4 = &H8002&   'MD4 hashing algorithm.
'Const CALG_MD5 = &H8003&   'MD5 hashing algorithm. This algorithm is supported by the Microsoft Base Cryptographic Provider.
'Const CALG_SHA = &H8004&   'SHA hashing algorithm. This algorithm is supported by the Microsoft Base Cryptographic Provider.
'Const CALG_SHA_256 = &H800C&   '256 bit SHA hashing algorithm. This algorithm is supported by Microsoft Enhanced RSA and AES Cryptographic Provider..Windows XP with SP3: This algorithm is supported by the Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype).
'Const CALG_SHA_384 = &H800D&   '384 bit SHA hashing algorithm. This algorithm is supported by Microsoft Enhanced RSA and AES Cryptographic Provider.Windows XP with SP3: This algorithm is supported by the Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype).
'Const CALG_SHA_512 = &H800E&   '512 bit SHA hashing algorithm. This algorithm is supported by Microsoft Enhanced RSA and AES Cryptographic Provider.Windows XP with SP3: This algorithm is supported by the Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype).
Dim hCtx As Long, hHash As Long, lLen As Long, i As Long, b() As Byte
bFile = Cast(bFile, 3)
If bFile Then If LenFile(src) = 0 Then Exit Function
If CryptAcquireContext(hCtx, "", "", 24&, &HF0000000) Then
    If CryptCreateHash(hCtx, calg, 0&, 0&, hHash) Then
        If bFile Then
            Dim st As New Stream
            st.Type = adTypeBinary: st.Open: st.LoadFromFile src
            Do While Not st.EOS: b = st.Read(&H10000): CryptHashData hHash, b(0), UBound(b) + 1, 0: Loop
            st.Close
        Else
            CryptHashData hHash, ByVal src, Len(src), 0
        End If
        If CryptGetHashParam(hHash, 4, lLen, 4, 0) Then
            ReDim b(0 To lLen - 1)
            If CryptGetHashParam(hHash, 2, b(0), lLen, 0) Then For i = 0 To UBound(b): Hash = Hash & Right$("0" & Hex$(b(i)), 2): Next
        End If
        CryptDestroyHash hHash
    End If
End If
CryptReleaseContext hCtx, 0
End Function

Function FSO() As FileSystemObject
Set FSO = New FileSystemObject
End Function


Private Sub Class_Initialize()
'Debug.Print "XMAIN Class_Initialize"
Set mSC = VBScript("*", Empty)
Set tm = New CTimer
'Set mCloseTimer = New CTimer
JETProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    

Set mRegEx = RegEx
'Port = 8888
'Set m_Connection = New Connection 'TWO THERAD ADD
'Set m_Cons = New Collection
xConnectionReset
'Set XConnection = New Connection

Set m_Buffer = New CParam: m_Buffer.Name = "Buffer"
Set m_Base = New CBase
'gRSType = Split("` ~ ! @ # $ % ^ & * : ' + ?", " ")
gRSType = "`~!@#$%^&*:'+?"
'#If speedc Then
'm_Param("speedc") = 1
'#End If
End Sub

Private Sub Class_Terminate()
On Error Resume Next
'If m_Connection.State Then m_Connection.Cancel: m_Connection.Close
'Set m_Connection = Nothing
xConnectionReset

Set m_Buffer = Nothing
Set m_Base = Nothing
mSC.Reset: Set mSC = Nothing
Set mRegEx = Nothing: Set cpRegEx = Nothing
Set tm = Nothing
#If DragDrop Then
Set iDTH = Nothing
#End If
'Debug.Print "XMAIN Class_Terminate"
End Sub


Private Function Ref(v0, Optional v1): Ref = v0: v1 = v0: End Function


'''===================CallDll=================
#If CallDll Then
Private Function STRUCT(ByVal vtype, cpValue, op As Boolean, Optional ofset& = 0)
'stype =  [ - ] vts
' -      = alloc bufer get ptr
' vts = var definion  =  vartype ( ( F? |  (,lbound,celements)+ |  [fix_celements] |  (dyn_celements) ) |  *BSTR_char_count )
Dim i&, b() As Byte, cb&
Dim rb() As Byte, n&
Dim p&, ptr&, v, vtp
If IsArray(vtype) Then
    If op Then n = -1 Else v = vtype
    For i = 1 To UBound(vtype) Step 2
        If op Then 'write value  to buffer
            rb = STRUCT(vtype(i), cpValue(i), op, cb)
            If cb Then ReDim Preserve b(n + cb): CopyMemory ByVal VarPtr(b(0)) + n + 1, rb(0), cb: n = UBound(b)
        Else 'write buffer to value
            v(i) = STRUCT(vtype(i), cpValue, op, ofset)
        End If
    Next
Else
    vtp = vtype: If Left(vtp, 1) = "-" Then vtp = Mid(vtp, 2)
    n = Val(vtp)
    If op Then 'write value  to buffer
        If vtype = cpValue Then cpValue = Empty
        If InStr(vtype, "~") Then cpValue = StrConv(cpValue, vbFromUnicode)
        Select Case Left(vtype, 1)
            Case "-"  'LPTR ALLOC alloc buf , copy buf  to value, RET =vptr(value)
                cpValue = pNewVariant(vtp, cpValue, n)  'replace  parent value with struct
                b = pNewVariant("17(4)", VPtr(cpValue), 3, cb)  'write long ptr to buffer
            Case Else 'FILL BYTES
                v = pNewVariant(vtp, cpValue, n, cb) 'CAST VALUE to BUFER
                b = pNewVariant("17(" & cb & ")", v, 0, cb)  'write value to buffer
        End Select
        ofset = cb
    Else 'write buffer to value
        p = VPtr(cpValue)
        Select Case Left(vtype, 1)
            Case "-"  'PTR.3 to value
                'If p > &HFFFFF Then p = GMem4(ByVal p + ofset)
                v = pNewVariant(vtp, GMem4(ByVal p + ofset), -1, cb): cb = 4
            Case Else 'Buffer to value
                v = pNewVariant(vtp, p + ofset, -1, cb)
        End Select
        If InStr(vtype, "~") Then v = StrConv(v, vbUnicode)
        If InStr(vtype, "*") Then v = RTrimChar("" & v, vbNullChar & vbNullChar)
        ofset = ofset + cb
    End If
End If
If op Then STRUCT = b Else STRUCT = v
End Function


Function CallDll(dllfun_json, Optional ByVal cpargs) As CParam
Dim cp As New CParam: cp.json = dllfun_json
Dim vta: vta = cp.Source
Dim hLib&:  hLib = LoadLibrary(S_(vta(1))): If hLib = 0 Then Exit Function
Dim fPtr&: fPtr = GetProcAddress(hLib, S_(vta(3))): If fPtr = 0 Then Exit Function
Dim vRtn As Variant, ret_VT%: ret_VT = L_(vta(5))
Dim i%, n%, ptr&
Dim args, vParams, p%
args = Split(Join2(cp.GetArray(, , 1), ","), ",") 'All params paths
Dim exc: exc = Split(Replace(Replace(Join(Filter(Filter(args, "+", False), "-"), ","), "-", ""), "+", ""), ",") 'EXCLUDE -PARAM INCLUDE +PARAM
'Debug.Print Join2(exc, ", ")

Do While UBound(Filter(args, "-")) > -1 Or UBound(Filter(args, "+")) > -1
    For i = 0 To UBound(args) 'To 0 Step -1 'REMOVE  +/- FROM NAMES
    If InStr(args(i), "-") + InStr(args(i), "+") Then
        If VarType(cp(args(i))) > 1 Then
            cp.PathParamName(args(i)) = xMain.SplitIndex(Replace(Replace(args(i), "-", ""), "+", ""))
            vParams = args(i)
            Do While InStr(vParams, "\") > 0
                vParams = xMain.SplitRemove(vParams)
                If VarType(cp(vParams)) > 1 Then cp.PathParamName(vParams) = xMain.SplitIndex(Replace(Replace(vParams, "-", ""), "+", "")) Else Exit Do
            Loop
        End If
    End If
    Next
    args = Split(Join2(cp.GetArray(, , 1), ","), ",") 'All params paths
Loop

vta = cp.Source
n = cp.ParamCount - 4 ' max param index

If Not (IsMissing(cpargs) Or IsEmpty(cpargs)) Then
    If Not IsArray(cpargs) Then cp.json("_inputargs") = cpargs Else cp("_inputargs") = cpargs
    cp.AddParams cp("_inputargs")
    cp("_inputargs") = Null
End If

Dim vParamPtr() As Long, vParamType() As Integer, vBuffer()
ReDim vParamPtr(n), vParamType(n), vBuffer(n)

args = cp.Source 'BYREF
vParams = args 'BYVAL

For i = 7 To UBound(vta) Step 2
    ptr = Val(Nz(vta(i), -1))
    p = (i - 7) \ 2
    If ptr <> -1 Then If args(i) = vta(i) Then args(i) = Empty: vParams(i) = Empty
    Select Case ptr
    Case Is < 0, 8 'BYREF OR STRUCT OR BSTR
        If ptr = -1 Then
            vBuffer(p) = STRUCT(vta(i), args(i), True) 'PREPARE VALUE + buffer as BYTE()
        Else
            vBuffer(p) = STRUCT(Replace(vta(i), "-", ""), args(i), True) 'PREPARE VALUE + buffer as BYTE()
        End If
        vParams(i) = VPtr(vBuffer(p))
    Case Else 'BYVAL PTR/VALUE
        vParams(i) = pNewVariant(vta(i), args(i), ptr)  'PREPARE VALUE
    End Select
    vParamPtr(p) = VarPtr(vParams(i))
    vParamType(p) = VarType(vParams(i))
Next

'cp.Value = args: Debug.Print cp.json

Dim lValue&
lValue = DispCallFunc(0&, fPtr, 4, ret_VT, n + 1, vParamType(0), vParamPtr(0), vRtn)
vParams = Empty: Erase vParamPtr: Erase vParamType

args(5) = lValue
If lValue = 0 Then  ' 0 = S_OK
    For i = 7 To UBound(vta) Step 2
        ptr = Val(Nz(vta(i), -1)): p = (i - 7) \ 2
        If ptr = -1 Then 'STRUCT TO VALUE
            args(i) = STRUCT(vta(i), vBuffer(p), False)
        ElseIf ptr < 0 Or ptr = 8 Then 'Buffer to Value
            args(i) = pNewVariant(Replace(vta(i), "-", ""), vBuffer(p))
            If InStr(vta(i), "*") Then args(i) = RTrimChar("" & args(i), vbNullChar & vbNullChar)
            If InStr(vta(i), "~") Then args(i) = StrConv(args(i), vbUnicode)
        End If
    Next
    If ret_VT = 0 Then args(5) = lValue Else args(5) = vRtn ' return result
End If
Erase vBuffer: vta = Empty
cp.Value = args: args = Empty
cp.FillNames exc, Null: exc = Empty 'EXCLUDE -PARAMS INCLUDE +PARAMS

If lValue Then  ' set error & return Empty
'    SetLast Error lValue
    cp("err") = lValue
End If

Set CallDll = cp
End Function
#End If
'''===================CallDll=================

''''===================URLParams=================
'Function URLParams(ParamArray a())
'depricate "URLParams", "+mod_HTTP.URLParams"
'End Function
'Function URLParams(ByVal v, ByVal pf, ByVal p) 'CParam->URLParams
'Dim i, s
'Dim pp
'If Not IsArray(v) Then
'    s = p & "=" & v
'Else
'    For i = 0 To UBound(v) Step 2
'        If Len(s) Then s = s & "&"
'        pp = pf
'        If p = "" Then
'            pp = pp & v(i)
'        ElseIf Len(RegExMatch(v(i), ".*\.(\d+)$", -2)) Then
'            If IsArray(v(i + 1)) Then
'                pp = p
'            Else
'                pp = p & "[]"
'            End If
'        Else
'            pp = p & "[" & v(i) & "]"
'        End If
'        s = s & URLParams(v(i + 1), "", pp)
'    Next
'End If
'URLParams = s
'End Function
'''''===================URLParams=================



Friend Sub SynInit()
Dim r, c As New CParam
c.AddParams Split(Nz(CParam("+SYN").Name), vbCrLf)
For Each r In c.Map(, 1): With r.Value: Syn2 r.Name, Join(Array(r.Name, .spc, .rm, .ms, .sq, .sj, .msj, .ps, .sp, .pf), vbCrLf): End With: Next
End Sub

Private Function Syn2(ByVal lang, Optional ByVal v) As Long
Static sx() As syndef2
Dim i&, n&, s As syndef2, a
If GMem4(ByVal VarPtrArray(sx)) = 0 Then
    ReDim sx(-1 To -1)
    s.lang = "noscript"

'    s.lang = "vbscript" 'lang
'    s.spc = "^(#\w+(?:|"")=)[^\0]*?\1.*$|^#\w+""?[= ].*$|\$""[^\0]*?""\$|""(?:""""|.)*?""|'.*$|/\*(?:[\S\s](?!/\*))*?\*/|#.+?#"  ' strings+comments+directives+places
'    s.rm = "/*" 'replace with preserve len+lines
'    s.ms = "$""" 'multiline string start chars
'    s.sq = """" 'string quot char
'    s.sj = "+vbCrLf+" 'multi string join
'    s.msj = "_" 'new line string join
'    s.ps = "(?:([\t ]*?(?:(?:public +(?:default +)?|private +))?(?:(sub|function|property)(?: +(?:get|let|set))?) +\w+ *(?:\([^\)]*\))?)[\s\S]*?end +\2\s*?|class +\w+.*|end +class.*|(?:dim|public|private) +.*)" 'list procedures
'    s.sp = ":" 'separator
'    s.pf = ":function f_%_:@:f_%_=%:end function:" 'replace float code
    
'    s.lang = "jscript"
'    s.spc = "{[^{]*?}|^(#\w+""?=)[^\0]*?\1|^#\w+""?[= ].*|'(?:\\'|\\\n|.)*?'|""(?:\\""|\\\n|.)*?""|//.*$|/\*(?:[\S\s](?!/\*))*?\*/"
'    s.rm = " "
'    s.ms = " "
'    s.sq = ""
'    s.sj = "\r\n"
'    s.msj = "\"
'    s.ps = "(?:^|;)(function\s*\w+\s*?\([^\)]*\))\s*?\x05\d+\x05"
'    s.sp = ";" 'separator
'    s.pf = ";function f_%_(){;@;return(%);};"
    
'    s.lang = "sql"
'    s.spc = """(?:""""|.)*?""|'(?:''|.)*?'|--.*$|/\*(?:[\S\s](?!/\*))*?\*/"  ' strings+comments+directives+places
'    s.rm = "/*,--" 'replace with preserve len+lines
'    s.ms = ""
'    s.sq = ""
'    s.sj = "" 'multi string join
'    s.msj = "" 'new line string join
'    s.ps = "" 'list procedures
'    s.sp = "" 'separator
'    s.pf = "" 'replace float code
    
    sx(-1) = s

End If
n = UBound(sx)
lang = S_(lang)
For i = 0 To n
    If sx(i).lang = lang Then Exit For
Next
If i > n Then i = -1
If VarType(v) = vbString Then 'v = "lang+spc+rm+ms+sq+sj+msj+ps+sp+pf".Split("+").join(vbCrLf)
    If i = -1 Then i = n + 1: ReDim Preserve sx(-1 To i)
    a = Split(v, vbCrLf)
    ReDim Preserve a(9)
    CopyMemory ByVal VarPtr(s), ByVal VPtr(a), Len(s)
    sx(i) = s
    FillMemory s, Len(s), 0
End If
Syn2 = VarPtr(sx(i))
End Function

Function PrepareCode(ByVal code, ByVal lang, Optional fortest = 0, Optional errLine)
Dim sx0 As syndef2, sx As syndef2: CopyMemory sx0, ByVal Syn2(lang), Len(sx): Let sx = sx0: FillMemory sx0, Len(sx0), 0
Dim m As Match, mn&, mi&, x&, i&, i1&, s$, src$
Dim ar, nar&: ar = Array()
Dim u&, n&, sn&, ds As New CParam, u5, u5n$, u5v$, dy&
Dim sp$, fp$, rm, rms: rms = Split(sx.rm, ","): sp = vbLf & vbTab & " " & sx.sp
'fp = Replace("@\s*\x06\d+\x06\s*(\x06\d+\x06)*|^(?:\x06\d+\x06\s*){1,}(?=\n|@)", "@", sx.sp)
fp = Replace(Replace("@\s*#\s*(#)*|^(?:#\s*){1,}(?=\n|@)", "@", sx.sp), "#", "\x06\d+\x06")

src = Replace(code, vbCrLf, vbLf)
sn = Len(src)
src = src & String(sn, 32)
On Error Resume Next
Do
    i = 0:  u = 0
    For Each m In RegExMatch(Left$(src, sn), sx.spc, -1)
        s = m.Value: x = 1
        'Debug.Print s
        If Left$(s, 1) = "#" Then '#NAME=
            u5 = Split(s, vbLf): u5v = Mid(u5(0), 2) 'GET FIRST ROW
            u5n = LeftStr(u5v, "=") 'GET NAME
            If Len(u5n) > 0 And InStr(u5n, " ") = 0 Then
                u5v = RightStr(u5v, "=", 1)
                If UBound(u5) Then  'single line def
                    dy = dy - UBound(u5)
                    pArrayLine u5, UBound(u5), , 3: pArrayLine u5, 0, , 3
                    If InStr(u5n, sx.sq) Then u5v = Join(Quot(u5, sx.sq), sx.sj) Else u5v = Join(u5, vbLf)
                Else
                    If InStr(u5n, sx.sq) Then
                        u5v = Quot(u5v, sx.sq)
                    ElseIf Left(u5v, 1) = "=" Then 'compute
                        n = 0: u5v = Nz(mSC.Eval(Mid(u5v, 2))): n = Err.Number
                        If n Then errLine = 1 - (dy - CharCount(Left(src, m.FirstIndex + 1 + i), vbLf)): SCError mSC, -errLine - 1
                        If fortest = 0 Or n Then DebugPrint IIf(n, -1, 1) * &HFF0000, " #" & u5n & "=[" & u5v & "] "
                    End If
                End If
                s = String(UBound(Split(s, vbLf)), vbLf)  'PRESRVE COUNT LINES
                 ds("#" & Replace(u5n, sx.sq, "") & "#") = u5v
            ElseIf Right$(s, 1) <> "#" Then '  #include.... #vb var=path...
                If fortest Then s = ""
            Else '  #var#
                x = 0
            End If
        ElseIf Left$(s, 2) = sx.ms Then ' $"
            u5 = Quot(Split(Mid(s, 3, Len(s) - 4), vbLf), sx.sq)
            s = Join(u5, sx.sj + sx.msj + vbLf)
        Else
            For Each rm In rms
                u5n = Trim(rm)
                If Left$(s, Len(u5n)) = u5n Then ' /* remove from code
                    If InStr(src, u5n) + Len(u5n) < m.FirstIndex + 1 Then u = 1 'NESTED
                    'n = UBound(Split(s, vbLf)): s = String(n, vbLf) & String(m.length - n, 32) 'PRESRVE LEN + COUNT LINES
                    If InStr(rm, " ") Then
                        n = UBound(Split(s, vbLf))
                        s = " " & String$(n, vbLf)  'PRESRVE LEN + COUNT LINES
                        dy = dy - n
                    End If
                End If
            Next
        
        End If
    'compress to ar
    If x Then
        GoSub arradd: ar(nar) = s: s = Chr(6) & nar & Chr(6): nar = nar + 1
        GoSub midsrc
    End If
        'dy = dy - 1
        'If u = 2 Then Exit For
    Next
'    SaveToFile App.Path & "\zzzzz_test." & IIf(lang = "vbscript", "vbs", "js"), sx.spc & vbCrLf & Left(src, sn)
Loop While u

If fortest Then
    'compact procedures
    'i = 0: For Each m In RegExMatch(Left$(src, sn), sx.ps, -1): s = Chr(6) & nar & Chr(6): GoSub midsrc: GoSub arradd: ar(nar) = m.Value: nar = nar + 1: Next
    'SaveToFile App.Path & "\zzzzz_test." & IIf(lang = "vbscript", "vbs", "js"), Left(src, sn)
    i = 0
    'InputBox "", "", sx.ps
    For Each m In RegExMatch(Left$(src, sn), sx.ps, -1)
        s = Chr(6) & nar & Chr(6)
        's = Chr(5) & nar & Chr(5)
        'Debug.Print VarType(m.SubMatches(2)), IsEmpty(m.SubMatches(2))
        'If IsEmpty(m.SubMatches(2)) Then mn = Len(m.SubMatches(0))
        GoSub midsrc
        GoSub arradd
        'If IsEmpty(m.SubMatches(2)) Then ar(nar) = m.SubMatches(0) Else ar(nar) = m.Value
        ar(nar) = m.Value
        nar = nar + 1
    Next
'                                    ReDim Preserve ar(nar - 1):    Compact Left$(src, sn), ar, 5, sx.ps, 0
    'SaveToFile App.Path & "\zzzzz_test." & IIf(lang = "vbscript", "vbs", "js"), Left(src, sn)
    'replace float
    'i = 0: For Each m In RegExMatch(Left$(src, sn), sx.fp, -1): s = Replace(Replace(sx.pf, "%", m.FirstIndex + 1), "@", m.Value): GoSub midsrc: Next
    i = 0
    'SaveToFile App.Path & "\zzzzz_test." & IIf(lang = "vbscript", "vbs", "js"), Left(src, sn)
    x = 0
    For Each m In RegExMatch(Left$(src, sn) & vbLf, fp, -1)
    'For Each m In RegExMatch(Left$(src, sn) & vbLf, "\x05\d+\x05", -1)
        mi = m.FirstIndex: mn = m.length
        GoSub prep6
        If m.length Then x = m.FirstIndex + m.length
    'SaveToFile App.Path & "\zzzzz_test." & IIf(lang = "vbscript", "vbs", "js"), Left(src, sn)
    Next
    mi = sn: mn = 1 ' m.length
    GoSub prep6
    
    'SaveToFile App.Path & "\zzzzz_test." & IIf(lang = "vbscript", "vbs", "js"), Left(src, sn)
End If

'For i = nar - 1 To 0 Step -1:src = Replace(src, Chr(6) & i & Chr(6), ar(i)):Next
'Do While sn And InStrRev(src, Chr(6), sn) > 0
Do While sn
'Debug.Print InStrRev(src, Chr(6), sn), lang
    If InStrRev(src, Chr(6), sn) = 0 Then Exit Do
    i = 0
    'Set rms = RegExMatch(Left(src, sn), "\x06(\d+)\x06", -1)|\x05(\d+)\x05
    For Each m In RegExMatch(Left$(src, sn), "\x06(\d+)\x06", -1)
    n = m.SubMatches(0): s = ar(n)
    'If IsEmpty(m.SubMatches(1)) Then n = m.SubMatches(0) Else n = m.SubMatches(0)
    s = ar(n)
    GoSub midsrc
    Next
Loop
'SaveToFile App.Path & "\zzzzz_test." & IIf(lang = "vbscript", "vbs", "js"), Left(src, sn)
'src = Replace2(src, ds.Value)
i = 0
For Each m In RegExMatch(Left$(src, sn), "#\w+\$?#", -1)
s = Nz(ds(m.Value)): If Len(s) Then GoSub midsrc
Next

'SaveToFile App.Path & "\zzzzz_test." & IIf(lang = "vbscript", "vbs", "js"), Left(src, sn)
PrepareCode = Replace(Left$(src, sn), vbLf, vbCrLf)
'SaveToFile App.Path & "\zzzzz_test." & IIf(lang = "vbscript", "vbs", "js"), PrepareCode
Exit Function

midsrc:
'Debug.Print "<<" & m.Value & ">>"
    If mn = 0 Then mn = m.length
    If mi = 0 Then mi = m.FirstIndex + 1
    mi = mi + i: n = Len(s)
    If sn - mn + n > Len(src) Then src = src & String(10 * (sn - mn + n - Len(src) + 1), 32)
    Mid(src, mi + n) = Mid(src, mi + mn): Mid(src, mi, n) = s
    i = i - mn + n: sn = sn - mn + n
    mn = 0: mi = 0
    'SaveToFile App.Path & "\zzzzz_test." & IIf(lang = "vbscript", "vbs", "js"), Left(src, sn)
    Return
arradd:
    If UBound(ar) < nar Then ReDim Preserve ar(UBound(ar) + 256)
    Return
prep6:
        If mi > x And mn Then
            'Debug.Print "<" & m.Value & ">"
            s = Mid(src, x + 1 + i, mi - x)
            'While Left(s, 1) = vbLf Or Left(s, 1) = vbTab Or Left(s, 1) = " ": x = x + 1: s = Mid(s, 2): Wend
            While Len(s) > 0 And InStr(sp, Left(s, 1)): x = x + 1: s = Mid(s, 2): Wend
            mn = mi - x: mi = x + 1
            'If Len(s) Then While Right(s, 1) = vbLf Or Right(s, 1) = vbTab Or Right(s, 1) = " ": mn = mn - 1: s = Left(s, mn): Wend
            If Len(s) Then While Len(s) > 0 And InStr(sp, Right(s, 1)): mn = mn - 1: s = Left(s, mn): Wend
            If mn Then
            'Debug.Print "<<" & s & ">>"
            s = Replace(Replace(Replace(sx.pf, "\", ""), "%", mi), "@", s)
            GoSub midsrc
            End If
        End If
        mi = 0: mn = 0
Return
End Function




''======== DOCMD =========================
Function DoCmd(ParamArray cmd())
On Error Resume Next
If UBound(cmd) < 0 Then
    Set DoCmd = arCmd
Else 'Viriefy hWNDS
    If IsArray(cmd(0)) And UBound(cmd) = 0 Then
        arCmd.Add cmd(0)
    Else
        arCmd.Add cmd
    End If
    'SendMessageCallBack GetDesktopWindow, WM_USER, 0, 0, AddressOf AsyncCall, 0
    tm.Interval = 10
End If
End Function

Private Sub tm_Timer()
If arCmd.count Then
On Error Resume Next
Dim cmd: cmd = arCmd(1)
arCmd.Remove 1

Dim hWnd&, uMsg&, obj As Object: Set obj = cmd(0)
If obj Is Nothing Then
    Select Case S_(cmd(0))
    Case "EndLoop"
        SendMessageCallBack frm.hWnd, WM_USER, 0, 0, AddressOf EndLoop, 1&
    Case "Quit": xMainWnd.Destroy
    Case "Reset"
     xMainWnd.GlobalModule = Null
    Case "CloseForm"
        hWnd = cmd(1): If IsWindow(hWnd) Then SendMessage hWnd, WM_CLOSE, 0, 0: If IsWindow(hWnd) Then PostMessage hWnd, WM_CLOSE, 0, 0: Exit Sub
    Case "Post"
        hWnd = cmd(1): uMsg = cmd(2): PostMessage hWnd, uMsg, 0&, 0&
    Case "Execute"
        Dim sc As ScriptControl
        Set sc = cmd(2): If sc Is Nothing Then Set sc = xMain.VBScript("*+", Empty)
        'sc.AddCode xMain.vbRemoveComents(cmd(1)) 'Run code
        sc.AddCode xMain.PrepareCode(cmd(1), sc.Language) 'Run code
    End Select
Else
    Dim i&, args(), nm$: nm = cmd(1)
    If UBound(cmd) > 1 Then
        ReDim args(UBound(cmd) - 2)
        For i = 0 To UBound(args): args(i) = cmd(i + 2): Next
    End If
    Call rtcCallByName(obj, StrPtr(nm), VbMethod, args)
End If
End If

'uMsg = 1
'Do While uMsg < arCmd.count
'    cmd = arCmd(uMsg)
'    If cmd(0) = "CloseForm" Or cmd(0) = "Post" Then If IsWindow(cmd(1)) Then arCmd.Remove uMsg: uMsg = uMsg - 1
'    uMsg = uMsg + 1
'Loop
If arCmd.count = 0 Then tm.Interval = 0
'If arCmd.count Then SendMessageCallBack GetDesktopWindow, WM_USER, 0, 0, AddressOf AsyncCall, 0
End Sub
''======== DOCMD =========================



'Function CPublic() As XVBEngine.Public
'Set CPublic = New XVBEngine.Public
'End Function
