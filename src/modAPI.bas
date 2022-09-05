Attribute VB_Name = "modAPI"
Option Explicit


'Public arCmd As New Collection
Declare Function SendMessageCallBack Lib "user32" Alias "SendMessageCallbackA" (ByVal hWnd&, ByVal MSG&, ByVal wParam&, ByVal lParam&, ByVal lpCallBack&, ByVal dwData As Long) As Long

Public Declare Function PostThreadMessage Lib "user32.dll" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Private Declare Function GetClassInfo Lib "user32.dll" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As Long, ByRef lpWndClass As WNDCLASS) As Long
Public Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Public Declare Function StrChr Lib "shell32.dll" Alias "StrChrA" (ByVal lpStart As String, ByVal wMatch As Integer) As Long


#If xScript Then
Public Type TProc
    Name As String
    NumArgs As Long
    'DISPID As Long
End Type
Public Type TERR
    Number As Long
    Description As String
    Line As Long
    pos As Long
    Text As String
    Source As String
End Type
Public Const GlobalModule = "Global"
#End If

Public Declare Function rtcCallByName Lib "MSVBVM60.DLL" (ByVal Object As Object, ByVal ProcName As Long, ByVal CallType As VbCallType, ByRef args() As Any, Optional ByVal lcid As Long) As Variant
'Public Declare Function rtcCallByName2 Lib "MSVBVM60.dll" Alias "rtcCallByName" (ByVal Object As Object, ByVal ProcName As Long, ByVal CallType As VbCallType, ByVal pargs As Long, Optional ByVal lcid As Long) As Variant
'Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (arr() As Any) As Long

'Public Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongW" ( _
'     ByVal hwnd As Long, _
'     ByVal nIndex As Long, _
'     ByVal dwNewLong As Long) As Long
''Public Const GCL_HICON As Long = -14
Public Declare Function CopyIcon Lib "user32.dll" ( _
     ByVal hIcon As Long) As Long


Public Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long

'Private Declare Function CoCreateInstance Lib "ole32" (rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As Any, pVarResult As Any) As Long

'VISTA 'Public Declare Function SHCreateDataObject Lib "shell32" (ByVal pidlFolder As Long, ByVal cidl As Long, ByVal apidl As Long, pdtInner As Any, riid As UUID, ppv As Any) As Long
#If DragDrop Then
Public Declare Function SHCreateFileDataObject Lib "shell32" Alias "#740" (ByVal pidlFolder As Long, ByVal cidl As Long, ByVal apidl As Long, pdtInner As Any, ppv As Any) As Long
''=Public Declare Function SHDoDragDrop Lib "shell32" (ByVal hWnd As Long, ByVal pdtobj As Long, ByVal pdsrc As Long, ByVal dwEffect As Long, pdwEffect As Long) As Long
''=Public Declare Function SHDoDragDrop Lib "shell32" (ByVal hWnd As Long, ByVal pdtobj As IDataObject, ByVal pdsrc As Long, ByVal dwEffect As Long, pdwEffect As Long) As Long

'Public Declare Function SHDoDragDrop Lib "shell32" (ByVal hWnd As Long, ByVal pdtobj As IDataObject, ByVal pdsrc As IDropSource, ByVal dwEffect As Long, pdwEffect As Long) As Long

#End If
'Public Declare Function DragDetect Lib "user32.dll" (ByVal hWnd As Long, pt As POINTAPI) As Long
'Public Declare Function DragDetect2 Lib "user32.dll" Alias "DragDetect" (ByVal hWnd As Long, x As Long, y As Long) As Long

'Public Declare Function ShowCaret Lib "user32.dll" (ByVal hWnd As Long) As Long
'Public Declare Function ScreenToClient2 Lib "user32.dll" Alias "ScreenToClient" (ByVal hWnd As Long, ByRef x As Long, y As Long) As Long


'Public Const xs.sEDIT = "Edit"
'Public Const xs.sRICH = "Rich"
'Public Const xs.sGRID = "Grid"
'Public Const xs.sDATE = "Date"
'Public Const xs.sPERIOD = "Period"
'Public Const xs.sBUTTON = "Button"
'Public Const xs.sCHECK = "Check"
'Public Const xs.sRADIO = "Radio"
'Public Const xs.sGROUP = "Group"
'Public Const xs.sLABEL = "Label"
'Public Const xs.sSTATIC = "Static"
'Public Const xs.sTAB = "Tab"
'Public Const xs.sMDICLIENT = "MDICLIENT"
'Public Const xs.sxControl = "xControl" '!!! TypeName ==xControl  NOT XCONTROL

Public Type xc_const
sEDIT As String ' = "Edit"
sRICH As String ' = "Rich"
sGRID As String ' = "Grid"
sDATE As String ' = "Date"
sPERIOD As String ' = "Period"
sBUTTON As String ' = "Button"
sCHECK As String ' = "Check"
sRADIO As String ' = "Radio"
sGROUP As String ' = "Group"
sLABEL As String ' = "Label"
sSTATIC As String ' = "Static"
sTab As String ' = "Tab"
sMDICLIENT As String ' = "MDICLIENT"
sxControl As String ' = "xControl" '!!! TypeName ==xControl  NOT XCONTROL
End Type
Public xs As xc_const

Public Const MAX_FIELDINDEX = WM_KEYDOWN '=256
Public Const NO_INDEX = -1&

Private Declare Function GetDoubleClickTime Lib "user32.dll" () As Long
Public nDoubleClickTime& '=GetDoubleClickTime

Public Declare Function InflateRect Lib "user32.dll" ( _
     ByRef lpRect As RECT, _
     ByVal x As Long, _
     ByVal Y As Long) As Long

Public Declare Function GetKeyState Lib "user32.dll" ( _
     ByVal nVirtKey As Long) As Integer

'Public Declare Function GetWindowRgn Lib "user32.dll" ( _
'     ByVal hWnd As Long, _
'     ByRef hRgn As Long) As Long

Public Declare Function GetLayeredWindowAttributes Lib "user32.dll" ( _
     ByVal hWnd As Long, _
     ByRef crKey As Long, _
     ByRef bAlpha As Byte, _
     ByRef dwFlags As Long) As Long

'Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long


Public Type PAGESETUPDLG_struct
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    Flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type
Public Declare Function PageSetupDialog Lib "comdlg32.dll" Alias "PageSetupDlgA" (ByRef pPagesetupdlg As PAGESETUPDLG_struct) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" ( _
     ByVal hWnd As Long, _
     ByVal hPrinter As Long, _
     ByVal pDeviceName As String, _
     ByRef pDevModeOutput As Any, _
     ByRef pDevModeInput As Any, _
     ByVal fMode As Long) As Long


'Public wbHack 'Webbrowser hook windows

Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

'Public Declare Function CastPtrToIUnknown Lib "msvbvm60.dll" Alias "VarPtr" (ByVal pUnk As Long) As IUnknown
Public Declare Function RegisterWindowMessage Lib "user32.dll" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function SendMessageTimeout Lib "user32.dll" Alias "SendMessageTimeoutW" ( _
     ByVal hWnd As Long, _
     ByVal MSG As Long, _
     ByVal wParam As Long, _
     ByVal lParam As Long, _
     ByVal fuFlags As Long, _
     ByVal uTimeout As Long, _
     ByRef lpdwResult As Long) As Long
Public Declare Sub ObjectFromLresult Lib "OLEACC.dll" ( _
     ByVal lResult As Long, _
      riid As UUID, _
     ByVal wParam As Long, _
     ByRef ppvObject As Any)
     
Public Declare Sub GetMemObj Lib "MSVBVM60.DLL" (ByRef src As Long, ByRef DST As Object)

'Type EDIT_STREAM
'  dwCookie As Long
'  dwError As Long
'  pfnCallback As Long
'End Type
'
Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

'Private Declare Function LookupPrivilegeName Lib "advapi32.dll" Alias "LookupPrivilegeNameW" ( _
'     ByVal lpSystemName As String, _
'     ByRef lpLuid As LUID, _
'     ByVal lpName As Long, _
'     ByRef cbName As Long) As Long

Private Declare Function GetTokenInformation Lib "advapi32.dll" ( _
    ByVal TokenHandle As Long, _
    ByVal TokenInformationClass As Integer, _
    ByRef TokenInformation As Any, _
    ByVal TokenInformationLength As Long, _
    ByRef ReturnLength As Long) As Long

Private Type TOKEN_ELEVATION
   TokenIsElevated As Long
   TokenIsElevated2 As Long
End Type

Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" ( _
     ByVal lpSystemName As Long, _
     ByVal lpName As String, _
     ByRef lpLuid As LUID) As Long
     
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
Private Declare Function PrivilegeCheck Lib "advapi32.dll" ( _
     ByVal TokenHandle As Long, _
     ByRef RequiredPrivileges As PRIVILEGES_SET, _
     ByRef pfResult As Long) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" ( _
     ByVal TokenHandle As Long, _
     ByVal DisableAllPrivileges As Long, _
     ByRef NewState As TOKEN_PRIVILEGES, _
     ByVal BufferLength As Long, _
     ByRef PreviousState As Long, _
     ByRef ReturnLength As Long) As Long
     
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
'Private Declare Function GetCurrentThread Lib "kernel32.dll" () As Long
'Private Declare Function OpenThreadToken Lib "advapi32.dll" ( _
'    ByVal ThreadHandle As Long, _
'    ByVal DesiredAccess As Long, _
'    ByVal OpenAsSelf As Long, _
'    ByRef TokenHandle As Long) As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" ( _
     ByVal ProcessHandle As Long, _
     ByVal DesiredAccess As Long, _
     ByRef TokenHandle As Long) As Long

Public Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32.dll" () As Long
Public Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function GetClipboardFormatName Lib "user32.dll" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Public Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare Function EmptyClipboard Lib "user32.dll" () As Long


Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Public Const CP_UTF8   As Long = 65001
Public Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Declare Function MultiByteToWideCharStr Lib "kernel32.dll" Alias "MultiByteToWideChar" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

'Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Public Declare Function GetOEMCP Lib "kernel32.dll" () As Long

'Public Declare Function DosDateTimeToFileTime Lib "kernel32.dll" (ByVal wFatDate As Long, ByVal wFatTime As Long, ByRef lpFileTime As FILETIME) As Long
'Public Declare Function FileTimeToDosDateTime Lib "kernel32.dll" (ByRef lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long


'Public Declare Sub HookCOMTable Lib "iehelper.dll" (ByRef pIntObj As Long, ByVal COMTableThunk As Long, ByVal pNewFunc As Long, pOldMethodAddr As Long)

Public Enum TEvents
nEvent_Move = 1
nEvent_Size = 2
nEvent_Moved = 4
nEvent_Sized = 8
nEvent_MouseActivate = &H10
nEvent_ExitSizeMove = &H20
nEvent_AExitSizeMove = &H40
nEvent_CloseForm = &H80
nEvent_NoAutoDrop = &H100
nEvent_Create = &H200
nEvent_Focus = &H400
nEvent_keys = &H800
'nEvent_xvalue = &H1000
End Enum


Private Type BrowseInfo
  hWndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As Long
  lpszTitle      As Long
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHParseDisplayName Lib "shell32.dll" (ByVal pszName As Long, ByVal pbc As Long, ByRef ppidl As Long, ByVal sfgaoIn As Long, ByRef psfgaoOut As Long) As Long
Private m_BrowseInfoCurrentDirectory As String   'The current directory


Public Declare Function InvalidateRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long

'Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
'Public Declare Function SetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal hNewMenu As Long) As Long

'Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long

'Public Declare Function GetMenuItemID Lib "user32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
'Public Declare Function ModifyMenu Lib "user32.dll" Alias "ModifyMenuW" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long

Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function RemoveMenuItem Lib "user32" Alias "RemoveMenu" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long



Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'Public Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Public Declare Function GetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthW" (ByVal hWnd As Long) As Long


Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'Public Declare Function GetTickCount Lib "kernel32" Alias "GetTickCount" () As Long


'Private Declare Function IsDialogMessage Lib "user32.dll" Alias "IsDialogMessageA" (ByVal hDlg As Long, ByRef lpMsg As MSG) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameW" (ByVal hWnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
'Public Declare Function GetDlgCtrlID Lib "user32.dll" (ByVal hwnd As Long) As Long
'Public Declare Function CheckRadioButton Lib "user32.dll" (ByVal hDlg As Long, ByVal nIDFirstButton As Long, ByVal nIDLastButton As Long, ByVal nIDCheckButton As Long) As Long
Public Declare Function GetNextDlgGroupItem Lib "user32" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
Public Declare Function UnionRect Lib "user32" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long


Public Declare Function GetCapture Lib "user32" () As Long
'Public Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Declare Function VariantTimeToSystemTime Lib "oleaut32.dll" (ByRef vtime As Double, ByRef lpSystemTime As SYSTEMTIME) As Long
'Public Declare Function SystemTimeToVariantTime Lib "oleaut32.dll" (ByRef lpSystemTime As SYSTEMTIME, ByRef pvtime As Double) As Long
'Public Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (ByRef lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Long

'Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'Public Declare Function CreateEvent Lib "kernel32.dll" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
'Public Declare Function PulseEvent Lib "kernel32.dll" (ByVal hEvent As Long) As Long
'Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

'Public Const LVM_FIRST As Long = &H1000
'Public Const LVM_SETITEMCOUNT As Long = (LVM_FIRST + 47)
'Public Const LVM_INSERTCOLUMNA As Long = (LVM_FIRST + 27)
'Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
'
'Public Const LVS_OWNERDATA As Long = &H1000
'Public Const LVS_REPORT As Long = &H1
'
'Public Const LVS_EX_FULLROWSELECT As Long = &H20
'Public Const LVS_EX_GRIDLINES As Long = &H1
'
'Public Const LVCF_WIDTH = &H2
'Public Const LVCF_TEXT = &H4
'
'Public Const LVN_FIRST As Long = -100
'Public Const LVN_GETDISPINFOA As Long = (LVN_FIRST - 50)
'Public Const NM_FIRST As Long = 0
'Public Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)
'Public Type LVCOLUMN
'    mask As Long
'    fmt As Long
'    cx As Long
'    pszText As String
'    cchTextMax As Long
'    iSubItem As Long
'    iImage As Long
'    iOrder As Long
'End Type
'
'Public Type LVITEM
'    mask As Long
'    iItem As Long
'    iSubItem As Long
'    State As Long
'    stateMask As Long
'    pszText As String
'    cchTextMax As Long
'    iImage As Long
'    lParam As Long
'    iIndent As Long
'End Type
'Public Type NMLVDISPINFO
'    hdr As NMHDR
'    Item As LVITEM
'End Type


Private Declare Function zlib_compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function zlib_uncompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Type z_stream
    next_in As Long  '/* next input byte */
    avail_in As Long   '/* number of bytes available at next_in */
    total_in As Long   '/* total nb of input bytes read so far */

    next_out As Long '/* next output byte should be put there */
    avail_out As Long  '/* remaining free space at next_out */
    total_out As Long  '/* total nb of bytes output so far */

    MSG As Long      '/* last error message, NULL if no error */
    internal_state As Long '/* not visible by applications */

    zalloc As Long     '/* used to allocate the internal state */
    zfree As Long      '/* used to free the internal state */
    opaque As Long     '/* private data object passed to zalloc and zfree */

    data_type As Long  '/* best guess about the data type: ascii or binary */
    adler As Long      '/* adler32 value of the uncompressed data */
    reserved As Long   '/* reserved for future use */
End Type
 
 
 
Private Declare Function zlib_inflateInit2 Lib "zlib.dll" Alias "inflateInit2_" (ByRef strm As z_stream, ByVal windowBits As Long, ByVal Version As String, ByVal stream_size As Long) As Long
Private Declare Function zlib_inflate Lib "zlib.dll" Alias "inflate" (ByRef strm As z_stream, ByVal flush As Long) As Long
Private Declare Function zlib_inflateEnd Lib "zlib.dll" Alias "inflateEnd" (ByRef strm As z_stream) As Long
Private Declare Function zlib_zlibVersion Lib "zlib.dll" Alias "zlibVersion" () As Long

'Public Type PLASTINPUTINFO
'   cbSize As Long
'   dwTime As Long
'End Type
'Public Declare Function GetLastInputInfo Lib "user32.dll" (ByRef plii As PLASTINPUTINFO) As Long
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'Public Type TBBUTTON
'   iBitmap As Long
'   idCommand As Long
'   fsState As Byte
'   fsStyle As Byte
'   bReserved1 As Byte
'   bReserved2 As Byte
'   dwData As Long
'   iString As Long
'End Type
'Public Type TBADDBITMAP
'    hInst As Long
'    nID As Long
'End Type
'Public Declare Function CreateToolbarEx Lib "COMCTL32.DLL" (ByVal hWndMain As Long, ByVal ws As Long, ByVal wID As Long, ByVal nBitmaps As Long, ByVal hBMInst As Long, ByRef wBMID As Long, ByRef lpButtons As TBBUTTON, ByVal numButtons As Long, ByVal dxButton As Long, ByVal dyButton As Long, ByVal dxBitmap As Long, ByVal dyBitmap As Long, ByVal uStructSize As Long) As Long
'Public Declare Function GetEnhMetaFileDescription Lib "gdi32.dll" Alias "GetEnhMetaFileDescriptionA" (ByVal hemf As Long, ByVal cchBuffer As Long, ByVal lpszDescription As String) As Long

Public Type DOCINFO
    cbSize As Long
    lpszDocName As String
    lpszOutput As String
End Type


'Public Type DevMode1
'    dmDeviceName As String * 32
'    dmSpecVersion As Integer
'    dmDriverVersion As Integer
'    dmSize As Integer
'    dmDriverExtra As Integer
'    dmFields As Long
'    dmOrientation As Integer
'    dmPaperSize As Integer
'    dmPaperLength As Integer
'    dmPaperWidth As Integer
'    dmScale As Integer
'    dmCopies As Integer
'    dmDefaultSource As Integer
'    dmPrintQuality As Integer
'    dmColor As Integer
'    dmDuplex As Integer
'    dmYResolution As Integer
'    dmTTOption As Integer
'    dmCollate As Integer
'    dmFormName As String * 32
'    dmUnusedPadding As Integer
'    dmBitsPerPel As Integer
'    dmPelsWidth As Long
'    dmPelsHeight As Long
'    dmDisplayFlags As Long
'    dmDisplayFrequency As Long
'End Type

'Public Type DEVNAMES
'    wDriverOffset As Integer
'    wDeviceOffset As Integer
'    wOutputOffset As Integer
'    wDefault As Integer
'    'extra As String
'End Type



Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Public Type TCITEM 'HEADER       '// Same for ANSI/Wide
    mask As Long
    lpReserved1 As Long
    lpReserved2 As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
End Type

'Public Enum TCITEM_Mask
'    TCIF_TEXT = &H1
'    TCIF_IMAGE = &H2
'    TCIF_RTLREADING = &H4
'    TCIF_PARAM = &H8
'    TCIF_STATE = &H10
'End Enum

'Const LF_FACESIZE = 32
'Public Type CHARFORMAT
'    cbSize As Integer '2
'    wPad1 As Integer  '4
'    dwMask As Long    '8
'    dwEffects As Long '12
'    yHeight As Long   '16
'    yOffset As Long   '20
'    crTextColor As Long '24
'    bCharSet As Byte    '25
'    bPitchAndFamily As Byte '26
'    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58
'    wPad2 As Integer ' 60
'End Type

'Public Type CHARRANGE
'    cpMin As Long
'    cpMax As Long
'End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long

Private Type tCHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (ByRef pChoosecolor As tCHOOSECOLOR) As Long

Private Type ChooseFontStruct
   lStructSize As Long
   hWndOwner As Long
   hDC As Long
   lpLogFont As Long
   iPointSize As Long
   Flags As Long
   rgbColors As Long
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
   hInstance As Long
   lpszStyle As String
   nFontType As Integer
   MISSING_ALIGNMENT As Integer
   nSizeMin As Long
   nSizeMax As Long
End Type
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (ByRef pChoosefont As ChooseFontStruct) As Long

'Private Type PAGESETUPDLG
'    lStructSize As Long
'    hWndOwner As Long
'    hDevMode As Long
'    hDevNames As Long
'    flags As Long
'    ptPaperSize As POINTAPI
'    rtMinMargin As RECT
'    rtMargin As RECT
'    hInstance As Long
'    lCustData As Long
'    lpfnPageSetupHook As Long
'    lpfnPagePaintHook As Long
'    lpPageSetupTemplateName As String
'    hPageSetupTemplate As Long
'End Type
'Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (ByRef pPagesetupdlg As PAGESETUPDLG) As Long



Public mncm As NONCLIENTMETRICS


'=====================================================================================================
Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum
Private Type TRACKMOUSEEVENT_STRUCT
    cbSize As Long
    dwFlags As TRACKMOUSEEVENT_FLAGS
    hwndTrack As Long
    dwHoverTime As Long
End Type
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
'=====================================================================================================


'=========================================================================================================
'ACCELERATOR
'Type ACCEL
'   fVirt As Byte
'   key As Integer
'   Cmd As Integer
'End Type
''Declare Function LoadAccelerators Lib "user32" Alias "LoadAcceleratorsA" (ByVal hInstance As Long, ByVal lpTableName As String) As Long
'Declare Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableA" (lpaccl As ACCEL, ByVal cEntries As Long) As Long
'Declare Function DestroyAcceleratorTable Lib "user32" (ByVal haccel As Long) As Long
''Declare Function CopyAcceleratorTable Lib "user32" Alias "CopyAcceleratorTableA" (ByVal hAccelSrc As Long, lpAccelDst As ACCEL, ByVal cAccelEntries As Long) As Long
'Declare Function TranslateAccelerator Lib "user32" Alias "TranslateAcceleratorA" (ByVal hWnd As Long, ByVal hAccTable As Long, lpMsg As MSG) As Long
''  Defines for the fVirt field of the Accelerator table structure.
'Public Const FVIRTKEY = True          '  Assumed to be == TRUE
'Public Const FNOINVERT = &H2
'Public Const FSHIFT = &H4
'Public Const FCONTROL = &H8
'Public Const FALT = &H10
'=========================================================================================================

'=========================================================================================================
'SHELLNOTIFYICON

Public Type NOTIFYICONDATA
    cbSize                           As Long
    hWnd                             As Long
    UID                              As Long
    uFlags                           As Long
    uCallbackMessage                 As Long
    hIcon                            As Long
    szTip                            As String * 128
    dwState                          As Long
    dwStateMask                      As Long
    szInfo                           As String * 256
    uTimeout                         As Long
    szInfoTitle                      As String * 64
    dwInfoFlags                      As Long
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

'icon flags
Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_GUID = &H5
Private Const NIIF_ICON_MASK = &HF
Private Const NIIF_NOSOUND = &H10
'=========================================================================================================


' Window messages
'Public Const WM_NULL = &H0

'Private Type tagInitCommonControlsEx
'   lngSize As Long
'   lngICC As Long
'End Type
'Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean


'Private Type WNDCLASS
'    Style As Long
'    lpfnwndproc As Long
'    cbClsextra As Long
'    cbWndExtra As Long
'    hInstance As Long
'    hIcon As Long
'    hCursor As Long
'    hbrBackground As Long
'    lpszMenuName As String
'    lpszClassName As String
'End Type
'Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (pClass As WNDCLASS) As Long
'Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long

'Public App_MessageLoop As Boolean
'Public MessageLoopRun As Boolean

Public nMessageLoops As Long
'Private App_Starting As Boolean
Public mLastMDWM As Long 'Last WM_LBUTONDOWN ime

Private hRich As Long
Public sRichWndClass$
'' Clipboard formats
Public CF_RTF As Long 'Integer
'Public CF_RTFWO As Integer
'Public CF_RTFWOO As Integer
'Public CF_EMBEDDEDOBJECT  As Integer
'Public CF_EMBEDSOURCE As Integer
'Public CF_OBJECTDESCRIPTOR As Integer
'Public CF_HTML As Integer


'=========================
Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    UID As Long
    lpRect As RECT
    hInst As Long
    lpszText As String
    lParam As Long
End Type

Private mActiveToolTip& 'hWnd
Private mToolTip& 'hWnd ToolTip
Private mBalloon& 'hWnd Free Balloon
Private mTip& 'hWnd Free Tip
'=========================
Public ThemeRun As Boolean
'============= ATL ============
Public Declare Function AtlAxWinInit Lib "atl.dll" () As Long
Public Declare Function AtlAxGetControl Lib "atl.dll" (ByVal hWnd As Long, ByRef pp As IUnknown) As Long
'Public Declare Function AtlAxGetHost Lib "atl.dll" (ByVal hWnd As Long, ByRef pp As IUnknown) As Long
'============= ATL ============



'====================================
'======SUBSCRIBING WM_MESSAGE====

Public m_ShinkMsgCount As Long
Public m_ShinkMsg() As Byte 'All Window Messages ARRAY storing ObjectInfo Index
Private m_Shinked As New Collection

Sub WMShink(ByVal xcHandle&, ByVal wTophWnd&, ByVal wMessage&, ByVal wParam&, ByVal wParamMask&, ByVal objectName$, ByVal eventName$, var_vbMacros)
If Abs(wMessage) > 1200 Then Exit Sub

If wMessage = -1 Then 'Clean  ALL Shinks
    ReDim m_ShinkMsg(-1200 To 1200)
    m_ShinkMsgCount = 0
    Set m_Shinked = New Collection
    Exit Sub
ElseIf wMessage = 0 Then 'Clean HANDLE Shinks
    UnWMShink xcHandle
    Exit Sub
End If



Dim p(), pf() As Long
ReDim p(0 To 7)
ReDim pf(0 To 7) As Long
pf(0) = 3: pf(1) = 3: pf(2) = 3: pf(3) = 3 'BYREF
p(4) = xcHandle: p(5) = wParam: p(6) = wParamMask: p(7) = wTophWnd

Dim ev As New CEventInfo
ev.frInitialize eventName, wMessage, p, Nothing, objectName, vbNullString, pf, var_vbMacros
m_Shinked.Add ev, "h" & ObjPtr(ev)

If m_ShinkMsgCount = 0 Then ReDim m_ShinkMsg(-1200 To 1200)
'If UBound(m_ShinkMsg) < wMessage Then ReDim Preserve m_ShinkMsg(wMessage)
m_ShinkMsg(wMessage) = m_ShinkMsg(wMessage) + 1
m_ShinkMsgCount = m_ShinkMsgCount + 1


End Sub
Private Sub pShinkRemove(msgid&, key$)
m_Shinked.Remove key
On Error Resume Next
If m_ShinkMsg(msgid) > 0 Then m_ShinkMsg(msgid) = m_ShinkMsg(msgid) - 1
m_ShinkMsgCount = m_ShinkMsgCount - 1
If m_ShinkMsgCount < 0 Then m_ShinkMsgCount = 0
End Sub

Public Sub UnWMShink(handle&)
Dim i&
Dim ev As CEventInfo
i = 1
While i <= m_Shinked.count
    Set ev = m_Shinked(i)
    If ev.Param(4) = handle Then pShinkRemove ev.DISPID, "h" & ObjPtr(ev) Else i = i + 1
Wend
End Sub


Sub RunShinkPost(wMsg As MSG)
Dim i&, xc As xControl, hWnd&
Dim ev As CEventInfo
Dim bWnd As Boolean
If m_ShinkMsgCount = 0 Then Exit Sub

If Abs(wMsg.Message) > 1200 Then Exit Sub
If m_ShinkMsg(wMsg.Message) = 0 Then Exit Sub

'Locate Listener
On Error Resume Next
i = m_Shinked.count
While i > 0
    Set ev = m_Shinked(i)
    If ev.DISPID = wMsg.Message Then 'Message=OK
        If (wMsg.wParam And ev(6)) = ev(5) Then 'WPARAM=OK
        hWnd = ev(7) 'wTophWnd
        bWnd = (hWnd = 0) 'Фильтр по окну не задан
        
        'Debug.Print Hex(wMsg.hwnd), Hex(Abs(hwnd))
        If Not bWnd Then bWnd = (Abs(hWnd) = wMsg.hWnd) 'Это мое окно
        If Not bWnd And hWnd > 0 Then bWnd = IsLongChild(hWnd, wMsg.hWnd) 'Это окно моего подчиненного
        If bWnd Then 'TopHWND=OK
            ev(0) = wMsg.hWnd
            ev(1) = wMsg.wParam
            ev(2) = wMsg.lParam
            'ev(2) = wMsg.wParam
            GetCursorPos wMsg.pt 'MOUSE POSITION
            'ev(2) = wMsg.pt.x
            'ev(3) = wMsg.pt.y
            ev(3) = MakeDWord(wMsg.pt.x, wMsg.pt.Y)
            'Set xc = ObjectFromPtr(ev(4)) 'INVALID HANDLE
            GetMemObj ev(4), xc
            hWnd = 0: hWnd = xc.hWnd
            
            If hWnd Then 'Post to Window
                Set xc = Nothing
                PostMessage hWnd, UM_SHINKWMSG, ByVal ObjPtr(ev), 0
'            ElseIf ev(4) = 0 Then 'Post To mainWND
'                xc.eax_HandleEvent ev, 0
'                Set xc = Nothing
            Else 'Remove DEAD SHINK
                pShinkRemove wMsg.Message, "h" & ObjPtr(ev)
            End If
            
        End If
        End If
    End If
    Set xc = Nothing
    i = i - 1
Wend
End Sub
'======SUBSCRIBING WM_MESSAGE====
'====================================


'====================================
'Private Declare Function SetCursorAPI Lib "user32.dll" Alias "SetCursor" (ByVal hCursor As Long) As Long

'
'Public Function SetCursor(ByVal hCursor As Long) As Long
'If hCursor = hCursor_WAIT Then
'Debug.Print "SetCursor = hCursor_WAIT "
'
'End If
'SetCursorAPI hCursor
'End Function
'
Public Function apiSetFocus(ByVal hWnd As Long) As Long
If GetFocus() <> hWnd And IsWindow(hWnd) Then apiSetFocus = aSetFocus(hWnd)
End Function

Public Sub StartAPP()
Set xMainWnd = New xControl

'If App_Starting Then Exit Function
xs.sEDIT = "Edit": xs.sRICH = "Rich": xs.sGRID = "Grid": xs.sDATE = "Date"
xs.sPERIOD = "Period": xs.sBUTTON = "Button": xs.sCHECK = "Check": xs.sRADIO = "Radio"
xs.sGROUP = "Group": xs.sLABEL = "Label": xs.sSTATIC = "Static": xs.sTab = "Tab": xs.sMDICLIENT = "MDICLIENT"
xs.sxControl = "xControl"  '!!! TypeName ==xControl  NOT XCONTROL

'Dim iccex As tagINITCOMMONCONTROLSEX, ret&
'iccex.dwSize = LenB(iccex)
'iccex.dwICC = &H4009
'''Const ICC_STANDARD_CLASSES As Long = &H4000
'''Const ICC_TAB_CLASSES As Long = &H8
'''Public Const ICC_LISTVIEW_CLASSES As Long = &H1
'ret = InitCommonControlsEx(iccex) 'TOOLTIP

''If hRich = 0 Then hRich = LoadLibrary(dll_RichOffice12): sRichWndClass = wc_RichEdit20A: App.LogEvent dll_RichOffice12 & "=" & hRich, 4
''If hRich = 0 Then hRich = LoadLibrary(dll_RichOffice11): sRichWndClass = wc_RichEdit20A: App.LogEvent dll_RichOffice11 & "=" & hRich, 4
''If hRich = 0 Then hRich = LoadLibrary(dll_RichMsftedit): sRichWndClass = wc_RichEdit50W: App.LogEvent dll_RichMsftedit & "=" & hRich, 4
''If hRich = 0 Then hRich = LoadLibrary(dll_Rich): sRichWndClass = wc_RichEdit20A: App.LogEvent dll_Rich & "=" & hRich, 4
'
''hRich = LoadLibrary("msftedit.dll")
''hRich = LoadLibrary("msftedit.1515.dll")
''hRich = LoadLibrary("msftedit.2509.dll")
'
''LoadLibrary "comctl32.dll"
'
''hRich = LoadLibrary("msftedit.dll"): sRichWndClass = "RichEdit50W"
''hRich = LoadLibrary("RICHED20.12.DLL"): sRichWndClass = "RichEdit20W"
''If Not IsInIDE Or 1 Then
'hRich = LoadLibrary(App.Path & "\RICHED20.DLL"): sRichWndClass = "RichEdit20W"
''End If
''If hRich = 0 Then hRich = LoadLibrary("RICHED20.DLL"): sRichWndClass = "RichEdit20W"
'If hRich = 0 Then hRich = LoadLibrary("msftedit.dll"): sRichWndClass = "RichEdit50W"
''sRichWndClass = "RichEdit20W"
''hRich = LoadLibrary("riched20.10.dll")
''hRich = LoadLibrary("riched20.12.dll")

hRich = LoadLibrary("msftedit.dll"): sRichWndClass = "RichEdit50W"

'InitClipFormats
CF_RTF = LOWORD(RegisterClipboardFormat("Rich Text Format"))
nDoubleClickTime = GetDoubleClickTime / 2
'xRegisterClass xs.sxControl, 0

mncm.cbSize = Len(mncm)
SystemParametersInfo 41, Len(mncm), mncm, 0 'SPI_GETNONCLIENTMETRICS = 41


'SM_CXVSCROLL = ncm.iScrollWidth 'GetSystemMetrics(2) '  'Ширина полосы прокрутки SM_CXVSCROLL As Long = 2
'Const SM_CXMENUCHECK As Long = 71
'Const SM_CYMENUCHECK As Long = 72
'SM_CXBORDER = GetSystemMetrics(5) ' ширина  границы окна SM_CXBORDER As Long = 5
'SM_CXEDGE = GetSystemMetrics(45) ' размеры трехмерной границы SM_CXEDGE As Long = 45
'SM_CXFIXEDFRAME = GetSystemMetrics(7) 'толщина рамки которая не может размер менять SM_CXFIXEDFRAME As Long = 7

'Debug.Print CurDir
'Debug.Print App.Path
'If CurDir <> App.Path Then ChDir App.Path
'xMain.ResRestore "zlib.dll"


mLastMDWM = GetTickCount
'es_Main 'Edanmo's Event Collection Class v2.0


Set GlobalGDI = New CParam
Set GlobalGDICount = New CParam
EnumSubClass.Clear
EnumSubClass.Name = "SubClass"
GlobalGDI.Clear
GlobalGDI.Name = "GDI"
GlobalGDICount.Clear
GlobalGDICount.Name = "GDICount"

pSysColor 0, Null 'Load Default SysColors

Common_GDI 1, 3
'Set xMainWnd = New xControl

xMain.LoadEngineSettings

'Dim i&, n&: n = 1
'Do
'i = i + 1: If pReplacemenu16Icon(n, "shell32", i) Then n = n + 1: Debug.Print "pReplaceIcon", i, n
'Loop While n < 269 And i < 3000
'pReplacemenu16Icon 5, "", 0
'pReplacemenu16Icon 7, "shell32", 10
'xMainWnd.Properties ("ClientBorder"), "2333"
'xMainWnd.Properties ("TreeIcons"), "2333"
'xMainWnd.Properties ("RowHeight"), "2333"

If xa.gWriteLog Then
    mLogFile = MakeDir(App.Path & "\LOG\" & App.EXEName & "_" & Format(Now(), "yymmddhhnnss") & ".log")
    mLogFileNumber = FreeFile()
    Open mLogFile For Output As mLogFileNumber Len = 10
End If

'MsgBox "StartAPP " & cmd

gDebugPrint "StartAPP(" & GetCurrentProcessId & ":" & App.ThreadID & ") " & Command$
xMainWnd.GlobalModule = Null

'wbHack = Array()
'App_MessageLoop = True
nMessageLoops = 1

End Sub

Sub DesktopRect(rc As RECT)
SystemParametersInfo 48, 0, rc, 0 'SPI_GETWORKAREA As Long = 48
End Sub

' Sub InitClipFormats()
'' RTF
''CF_RTF = &H8000 Or (RegisterClipboardFormat("Rich Text Format") And &H7FFF)
'CF_RTF = LOWORD(RegisterClipboardFormat("Rich Text Format"))
''CF_RTF = RegisterClipboardFormat("Rich Text Format")
'
''' RTF without objects
'''CF_RTFWOO = &H8000 Or (RegisterClipboardFormat("Rich Text Format Without Objects") And &H7FFF)
''CF_RTFWOO = RegisterClipboardFormat("Rich Text Format Without Objects")
''' RTF with objects
'''CF_RTFWO = &H8000 Or (RegisterClipboardFormat("RichEdit Text and Objects") And &H7FFF)
''CF_RTFWO = RegisterClipboardFormat("RichEdit Text and Objects")
''' Object descriptor
'''CF_OBJECTDESCRIPTOR = &H8000 Or (RegisterClipboardFormat("Object Descriptor") And &H7FFF)
''CF_OBJECTDESCRIPTOR = RegisterClipboardFormat("Object Descriptor")
''' Embedded object
'''CF_EMBEDDEDOBJECT = &H8000 Or (RegisterClipboardFormat("Embedded Object") And &H7FFF)
''CF_EMBEDDEDOBJECT = RegisterClipboardFormat("Embedded Object")
''' Embed source
'''CF_EMBEDSOURCE = &H8000 Or (RegisterClipboardFormat("Embed Source") And &H7FFF)
''CF_EMBEDSOURCE = RegisterClipboardFormat("Embed Source")
''' HTML Format
'''CF_HTML = &H8000 Or (RegisterClipboardFormat("HTML Format") And &H7FFF)
''CF_HTML = RegisterClipboardFormat("HTML Format")
''
'
'
'End Sub

'Public Property Get MessageLooping() As Boolean
'   MessageLooping = App_MessageLoop
'End Property
'Public Sub RunMessageLoop()
'If xMainWnd.Controls.count Then MessageLoop
'End Sub


Public Sub StopAPP()
'MsgBox "STOP APP"
xMainWnd.Destroy
Set xMainWnd = Nothing
Set rz = Nothing
'Set xMain = Nothing
ClearGDI vbNullString: comGDI.flag = 0
'App_MessageLoop = 0

If mToolTip Then DestroyWindow mToolTip
If mBalloon Then DestroyWindow mBalloon
If mTip Then DestroyWindow mTip
'FreeLibrary m_hMod
FreeLibrary hRich
'm_hMod = 0
UnregisterClass xs.sxControl, App.hInstance
'gDebugPrint App.ThreadID & " StopAPP " & Command$
gDebugPrint "StopAPP(" & GetCurrentProcessId & ":" & App.ThreadID & ") " & Command$

If mLogFileNumber Then Close mLogFileNumber

End Sub


Public Function xRegisterClass(ByVal cnm$, pop&) As Boolean ', Optional ic&) As Boolean
Dim s$, wc As WNDCLASS
s = StrConv(cnm, vbFromUnicode)
wc.lpszClassName = StrPtr(s)
wc.lpfnWndProc = AddrOf(AddressOf xWndProc)
wc.hInstance = App.hInstance
wc.Style = IIf(pop, 0, CS_PARENTDC)
'ic = GetIcon(ic, menu16)
'xRegisterClass = GetClassInfo(App.hInstance, StrPtr(s), wc)
'If xRegisterClass Then
'    SetClassLong
'Else
'    If ic Then wc.hIcon = GetIcon(ic, menu16)
    xRegisterClass = RegisterClass(wc)   ' Register class
    'xMain.DebugPrint 100, "RegisterClass (" & cnm & ")=" & xRegisterClass
'End If
End Function

'Public Function xUnRegisterClass(cnm$) As Boolean
'UnregisterClass cnm, App.hInstance
'End Function


Function xWndProc(ByVal hWnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Debug.Print Timer, "xWndProc", hWnd, Message
'xMain.DebugPrint 0, "xWndProc", hWnd, Message
Select Case Message
Case WM_DRAWITEM, WM_MEASUREITEM '&H2B ' WM_DRAWITEM=&H2B WM_MEASUREITEM=&H2C
    If wParam = 0 Then    'Для меню
        Dim itemData&
        Dim MenuItem As CMenuItem
        If Message = WM_DRAWITEM Then GetMem4 ByVal lParam + 44, itemData Else GetMem4 ByVal lParam + 20, itemData
        If Not itemData = 0 Then
            GetMemObj itemData, MenuItem
            If Not MenuItem Is Nothing Then
                If Message = WM_DRAWITEM Then MenuItem.DrawItem lParam Else MenuItem.MeasureItem lParam
                xWndProc = True
                Set MenuItem = Nothing
            Else
              '  If Err Then xMain.MsBox Err.Description, 16, "xWndProc"
            End If
            xWndProc = -1
            Exit Function
        End If
'    Else
'        xWndProc = CallOldWindowProc(hWnd, Message, wParam, lParam)
'        Exit Function
    End If
End Select
xWndProc = DefWindowProc(hWnd, Message, wParam, lParam)
End Function



'Private Sub WBTA(wMsg As MSG) 'WebBrowser TranslateAccelerator
'If gFindIndex(wbHack, wMsg.hWnd) = NO_INDEX Then Exit Sub
'If wMsg.Message <> WM_KEYDOWN Then Exit Sub
''If (wMsg.wParam = vbKeyBack) Or (wMsg.wParam = vbKeyLeft) Or (wMsg.wParam = vbKeyRight) Then Exit Sub
'Select Case wMsg.wParam
'Case vbKeyTab, vbKeyDelete, vbKeyC, vbKeyA, vbKeyV, vbKeyX, vbKeyZ
'
'Case Else
'Exit Sub
'End Select
''If Not ((wMsg.wParam = vbKeyTab) Or (wMsg.wParam = vbKeyTab) Or (wMsg.wParam = vbKeyTab)) Then Exit Sub
'
'Dim ipao As IOleInPlaceActiveObject
'AtlAxGetControl GetParent(GetParent(GetParent(wMsg.hWnd))), ipao
'If ipao Is Nothing Then Exit Sub
'ipao.TranslateAccelerator wMsg
''Debug.Print wMsg.wParam, pWindowClass(wMsg.hWnd)
'End Sub

'Sub MessageLoop()
'App_MessageLoop = 1
'MessageLoopRun = 1
'Dim wMsg As MSG
'Do While GetMessage(wMsg, 0, 0, 0) And App_MessageLoop
'    If UBound(wbHack) > NO_INDEX Then WBTA wMsg
'    TranslateMessage wMsg
'    DispatchMessage wMsg
'Loop
'MessageLoopRun = 0
'App_MessageLoop = 0
'End Sub


'Function MessageLoop() As Long '(ByVal a1&, ByVal a2&, ByVal a3&, ByVal a4&)
'xMain.DebugPrint "START MessageLoop"
'App_MessageLoop = 1
'MessageLoopRun = 1
'Dim wMsg As MSG
'Do While GetMessage(wMsg, 0, 0, 0) And App_MessageLoop
'    'If UBound(wbHack) > NO_INDEX Then WBTA wMsg
'    TranslateMessage wMsg
'    DispatchMessage wMsg
'Loop
'MessageLoopRun = 0
'App_MessageLoop = 0
'End Function

Public Sub MessageLoopWhileVisible(ByVal hWnd&)
'Stack 1
nMessageLoops = nMessageLoops + 1
Dim wMsg As MSG, v As Boolean
v = IsWindowVisible(hWnd)
While v
   If GetMessage(wMsg, 0, 0, 0) Then
'        If UBound(wbHack) > NO_INDEX Then WBTA wMsg
        TranslateMessage wMsg
        DispatchMessage wMsg
    Else
        v = 0
    End If
    v = v And IsWindowVisible(hWnd) 'And App_MessageLoop
    If hWnd = wMsg.hWnd Then
        If wMsg.Message = WM_CLOSE Or wMsg.Message = WM_DESTROY Then
            v = 0
        End If
    End If
Wend
nMessageLoops = nMessageLoops - 1
'Stack -1
End Sub

Public Sub API_DoEvents(Optional ok As Boolean = True)
If Not ok Then Exit Sub
Dim wMsg As MSG
While PeekMessage(wMsg, 0, 0, 0, 1)
    Call TranslateMessage(wMsg)
    Call DispatchMessage(wMsg)
Wend
End Sub

'Public Function GetScrollBars(hWnd As Long) As Long
'Dim n As Long
'n = GetWindowLong(hWnd, GWL_STYLE) 'Стиль окна
'If n And WS_HSCROLL Then GetScrollBars = 1
'If n And WS_VSCROLL Then GetScrollBars = GetScrollBars + 2
''SBX_NONE = 0
''SBX_HORZ = 1
''SBX_VERT = 2
''SBX_BOTH = 3
'End Function

Public Function ITextDocument(h) As ITextDocument ' Object
Dim hWnd&
hWnd = h
If h = 0 Then Exit Function
Dim oUnknown As IUnknown
If SendMessage(hWnd, EM_GETOLEINTERFACE, 0, oUnknown) = 0 Then Exit Function 'EM_GETOLEINTERFACE As Long = (WM_USER + 60)
Set ITextDocument = oUnknown
End Function

'Public Function ITextDocument2(h) As ITextDocument2 ' Object
'Dim hWnd&
'hWnd = h
'If h = 0 Then Exit Function
'Dim oUnknown As stdole.IUnknown
'If SendMessage(hWnd, EM_GETOLEINTERFACE, 0, oUnknown) = 0 Then Exit Function 'EM_GETOLEINTERFACE As Long = (WM_USER + 60)
'Set ITextDocument2 = oUnknown
'End Function

Public Function hxControl(ByVal hWnd As Long) As xControl
Dim o As Object, p As Long
If hWnd = 0 Then Set hxControl = xMainWnd: Exit Function
p = GetProp(hWnd, hWnd & "#1")
If p = 0 Then Exit Function 'Set hxControl = xMainWnd: Exit Function
'CopyMemory o, p, 4 'Turn pointer into a reference:
GetMemObj p, o
If TypeName(o) = xs.sxControl Then Set hxControl = o
'CopyMemory o, 0&, 4
End Function

Public Sub GetWindowPos(ByVal h&, wpos As WINDOWPOS)
'======== WINDOW POS===========================================
Dim rc As RECT, pt As POINTAPI, ph& ', wpos As WINDOWPOS
GetWindowRect h, rc
ph = Get_Parent(h)
If ph Then ScreenToClient ph, pt
OffsetRect rc, pt.x, pt.Y
wpos.x = rc.Left: wpos.Y = rc.Top
wpos.cy = rc.Bottom - rc.Top: wpos.cx = rc.Right - rc.Left
'======== WINDOW POS===========================================
End Sub


Public Function GetShiftState() As Long
Dim iR As Long
    iR = iR Or (-vbShiftMask * KeyIsPressed(&H10)) 'VK_SHIFT = &H10&
    iR = iR Or (-vbAltMask * KeyIsPressed(&H12)) 'VK_MENU = &H12& ' Alt key
    iR = iR Or (-vbCtrlMask * KeyIsPressed(&H11)) 'VK_CONTROL = &H11&
    GetShiftState = iR
End Function

Public Function KeyIsPressed(ByVal nVirtKeyCode As KeyCodeConstants, Optional ByVal mask& = &H8000&) As Boolean
Dim lR As Long
lR = GetAsyncKeyState(nVirtKeyCode)
'lR = GetKeyState(nVirtKeyCode)
If (lR And mask) = mask Then KeyIsPressed = True
End Function

'Public Function TranslateColor(ByVal clrColor As Long, Optional HPALETTE As Long = 0) As Long
'    If OleTranslateColor(clrColor, HPALETTE, TranslateColor) Then
'        TranslateColor = &HFFFF
'    End If
'End Function



Function LOBYTE(lWord As Integer) As Byte
LOBYTE = 0 + lWord And &HFF&
End Function
Function HIBYTE(lWord As Integer) As Byte
HIBYTE = 0 + (lWord And &HFF00&) \ &H100
End Function
Function LOWORD(lDWord As Long) As Integer
'If lDWord And &H8000& Then LOWORD = lDWord Or &HFFFF0000 Else LOWORD = lDWord And &HFFFF&
'Debug.Print Hex(lDWord), Hex(LOWORD)
GetMem2 lDWord, LOWORD
'Debug.Print Hex(lDWord), Hex(LOWORD)
End Function
Function HIWORD(lDWord As Long) As Integer
'HIWORD = (lDWord And &HFFFF0000) \ &H10000
'Debug.Print Hex(lDWord), Hex(HIWORD)
GetMem2 ByVal VarPtr(lDWord) + 2, HIWORD
'Debug.Print Hex(lDWord), Hex(HIWORD)
End Function

Function MakeDWord(ByVal LOWORD As Integer, ByVal HIWORD As Integer) As Long
'MakeDWord = (CLng(HIWORD) * &H10000) Or (LOWORD And &HFFFF&)
'Debug.Print Hex(HIWORD), Hex(LOWORD), , Hex(MakeDWord)
PutMem2 MakeDWord, LOWORD
PutMem2 ByVal VarPtr(MakeDWord) + 2, HIWORD
'Debug.Print Hex(HIWORD), Hex(LOWORD), , Hex(MakeDWord)

End Function

Public Sub TrackMouseLeave(ByVal lng_hWnd As Long)
    Dim tme As TRACKMOUSEEVENT_STRUCT
    tme.cbSize = Len(tme)
    tme.dwFlags = TME_LEAVE
    tme.hwndTrack = lng_hWnd
    TrackMouseEvent tme
End Sub

Public Function Get_Owner&(hWnd&)
Get_Owner = GetWindow(hWnd, 4&) 'GW_OWNER
End Function
Public Function Get_Parent&(hWnd&)
If hWnd Then Get_Parent = GetAncestor(hWnd, 1&) 'GA_PARENT
End Function
Public Function Get_OwnerPopup&(hWnd&, Optional MDICHILD_isPopup As Boolean = True)
Dim h&
h = hWnd
While Not IsPopup(h, MDICHILD_isPopup) And h
    h = GetParent(h)
Wend
Get_OwnerPopup = h
End Function

Public Function IsMayBeOwner(ByVal hWnd&) As Boolean
If hWnd = 0 Then Exit Function
IsMayBeOwner = Get_Parent(hWnd) = 0 Or Get_Parent(hWnd) = GetDesktopWindow()
If IsMayBeOwner Then IsMayBeOwner = (GetWindowLong(hWnd, GWL_STYLE) And WS_CHILD) = 0 And (GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_MDICHILD) = 0
'If IsMayBeOwner Then IsMayBeOwner = IsMayBeOwner(GetParent(hWnd))
End Function

Public Function IsPopup(ByVal hWnd&, Optional MDICHILD_isPopup As Boolean = True) As Boolean
If hWnd = 0 Then Exit Function
IsPopup = (GetWindowLong(hWnd, GWL_STYLE) And WS_CHILD) = 0 Or (MDICHILD_isPopup And (GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_MDICHILD))
End Function

Function WindowCanTabTree(ByVal hWnd&) As Boolean
Dim dws&
If IsWindow(hWnd) Then
    dws = GetWindowLong(hWnd, GWL_STYLE)
    'WindowCanTabTree = ((dws And WS_CHILD) = WS_CHILD) And ((dws And WS_VISIBLE) = WS_VISIBLE) And ((dws And WS_DISABLED) = 0)
    WindowCanTabTree = dws And (WS_CHILD Or WS_VISIBLE) And Not WS_DISABLED
    If WindowCanTabTree Then WindowCanTabTree = dws And WS_TABSTOP
End If
End Function

Private Sub WindowsTree(ByRef ar, ByVal hWnd&, Optional ByVal hWndExcludeChilds As Long)
Dim hWndChild&, dws&, s$
If hWnd = 0 Then Exit Sub
Do
    dws = GetWindowLong(hWnd, GWL_STYLE)
    If (((dws And WS_CHILD) = WS_CHILD) And ((dws And WS_VISIBLE) = WS_VISIBLE) And ((dws And WS_DISABLED) = 0)) Then ' And (hWnd <> hWndExclude) Then
        s = String(128, 0): s = Left(s, GetClassName(hWnd, StrPtr(s), Len(s)))
        hWndChild = GetWindow(hWnd, GW_CHILD)
        If hWndChild Then
            ReDim Preserve ar(UBound(ar) + 1): ar(UBound(ar)) = hWnd
            If s <> "AtlAxWin" And (hWnd <> hWndExcludeChilds) Then WindowsTree ar, hWndChild, hWndExcludeChilds
        Else
            If s <> xs.sSTATIC Then ReDim Preserve ar(UBound(ar) + 1): ar(UBound(ar)) = hWnd
        End If
    End If
    hWnd = GetWindow(hWnd, GW_HWNDNEXT)
Loop While hWnd
End Sub

Public Function GetNextWndTabItem&(ByVal hWnd&, ByVal nDirection&, Optional ByVal hWndExcludeChilds&)
Dim ar, i&, n&, h&
If IsWindow(hWnd) = 0 Or hWnd = 0 Then
    If nDirection = 0 Then GetNextWndTabItem = NO_INDEX
    Exit Function
End If
h = hWnd
While Not IsPopup(h) And h
    h = GetParent(h)
Wend
h = GetWindow(h, GW_CHILD) 'Первый чилд главного POPUP окна
ar = Array()
If h Then WindowsTree ar, h, hWndExcludeChilds
n = UBound(ar)
If nDirection = 0 Then
    GetNextWndTabItem = n 'DLGITEMINDEX GetLastDLGItemINDEX()
Else
If n < 0 Then Exit Function
    For i = 0 To n
        If ar(i) = hWnd Then Exit For
    Next
    h = 0
    Do
        i = i + nDirection: h = h + 1
        If i > n Then i = 0
        If i < 0 Then i = n
    Loop While Not WindowCanTabTree(CLng(ar(i))) And h <= n
    If h > n + 1 Then GetNextWndTabItem = hWnd Else GetNextWndTabItem = ar(i)
End If
'*******************************************************************************
'***********HELPER *********************************************************
'If 1 Then
'        Dim arn
'        arn = ar
'        For i = 0 To n
'            Select Case arn(i)
'            Case hWnd: arn(i) = "(" & hxControl(CLng(arn(i))).Name & ")"
'            Case GetNextWndTabItem: arn(i) = "<" & hxControl(arn(i)).Name & ">"
'            Case Else: arn(i) = hxControl(CLng(arn(i))).Name
'            End Select
'        Next
'        Debug.Print "TABORDER: " & Join2(arn, ",")
'End If
'*******************************************************************************
'*******************************************************************************
End Function


Public Function hWndZOrder&(ByVal hWnd&, Optional ByVal zIndex& = 0) ', Optional ByVal bCanTab&)
'zIndex<0 Возвращает максимальный ZOrder в группе
'zIndex=0 Возвращает Z-индекс окна
'zIndex>0 Возвращает hWnd окна по указанному Z-индексу
If hWnd = 0 Then Exit Function
Dim n&, h&, f As Boolean
f = zIndex > 0 'Надо искать окно по индексу
h = GetWindow(Get_Parent(hWnd), 5) ' GW_CHILD=5 Получаем первое окно в родителе
If zIndex = 0 Then n = 1 Else n = IIf(h, 1, 0)
While Not (h = 0) And IIf(f, IIf(zIndex = n, False, True), True) And Not (h = IIf(zIndex = 0, hWnd, -1))
    h = GetWindow(h, 2) ' GW_HWNDNEXT=2 'Получаем следующее окно брата
    If h Then n = n + 1
Wend
If f Then hWndZOrder = h Else hWndZOrder = n
End Function

'Public Function IsVisible(ByVal hWnd&) As Boolean
'If hWnd Then IsVisible = (GetWindowLong(hWnd, GWL_STYLE) And WS_VISIBLE) And WS_VISIBLE
''IsVisible = IsWindowVisible(hWnd)
'End Function

Public Function IsLongChild(ByVal hWndParent&, ByVal hWnd&, Optional ByVal bIncludePopup As Boolean) As Boolean
If hWndParent < 0 Then Exit Function
IsLongChild = (hWndParent = 0) And (hWnd <> 0)
If IsLongChild Then Exit Function
Do
'Debug.Print "wnd=" & Hex(hWnd);
'hWnd = Get_Parent(hWnd)
'Debug.Print " Get_Parent=" & Hex(Get_Parent(hWnd));
hWnd = GetParent(hWnd)
If hWnd = hWndParent Then IsLongChild = 1: Exit Do
If Not bIncludePopup Then If IsPopup(hWnd) Then Exit Do
Loop While hWnd <> hWndParent And hWnd <> 0
End Function


Public Function IsWindowCanFocus(ByVal hWnd&, Optional ByVal wsgroup As Boolean) As Boolean
Dim dws&, b As Boolean
If hWnd = 0 Then Exit Function
dws = GetWindowLong(hWnd, GWL_STYLE)
IsWindowCanFocus = (dws And WS_VISIBLE) <> 0 And (dws And WS_DISABLED) = 0
If IsWindowCanFocus And (dws And WS_CHILD) Then
    If (GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_MDICHILD) = 0 Then
        b = ((dws And WS_TABSTOP) > 0)
        If Not b And wsgroup Then b = ((dws And WS_GROUP) > 0)
        IsWindowCanFocus = b
    End If
End If
End Function



Public Function StringFromPointer(ByVal lPointer As Long) As String
    Dim strTemp As String
    Dim lRetVal As Long
    'prepare the strTemp buffer
    strTemp = String$(lstrlen(ByVal lPointer), 0)
    'copy the string into the strTemp buffer
    lRetVal = lstrcpy(ByVal strTemp, ByVal lPointer)
    'return a string
    If lRetVal Then StringFromPointer = strTemp
End Function
Public Function StringWFromPointer(ByVal lPointer As Long) As String
    Dim strTemp As String
    Dim lRetVal As Long
    'prepare the strTemp buffer
    strTemp = String$(lstrlenW(ByVal lPointer), 0)
    'copy the string into the strTemp buffer
    lRetVal = lstrcpyW(ByVal strTemp, ByVal lPointer)
    'return a string
    If lRetVal Then StringWFromPointer = strTemp
'    Dim s$: s = SysAllocString(lPointer)
'        If strTemp = s Then
'    End If
End Function
Function StringFromBSTR(ByVal pBSTR As Long) As String
    Dim temp As String
    ' copy the pointer into the temporary string's BSTR
    'CopyMemory ByVal VarPtr(temp), pBSTR, 4
    PutMem4 ByVal VarPtr(temp), pBSTR
    ' now Temp points to the original string, so we can copy it
    StringFromBSTR = temp
    ' manually clear then temporary string to avoid GPFs
    'CopyMemory ByVal VarPtr(temp), 0&, 4
    PutMem4 ByVal VarPtr(temp), 0&
End Function

Public Function IntersectWND(hWndParent, hWndChild) As Boolean 'hWndChild.WindowRect видно в hWndParent.ClientRect
Dim rc As RECT, wrc As RECT, pt As POINTAPI
GetWindowRect hWndChild, rc
ScreenToClient hWndParent, pt
OffsetRect rc, pt.x, pt.Y 'Теперь rc в координатах crc
GetClientRect hWndParent, wrc
IntersectWND = IntersectRect(rc, wrc, rc)
End Function

Public Function VScrollPos(hWnd, nDir&) As Boolean
'текущее положение VScrollPos = вначале(nDir<0) или в конце(nDir>0)
'nDir
'<0 MinPos
'=0 MinPos or MaxPos
'>0 MaxPos
Dim si As SCROLLINFO
si.cbSize = Len(si): si.fMask = 5
GetScrollInfo hWnd, 1, si
VScrollPos = ((nDir <= 0 And si.npos = si.nMin) Or (nDir >= 0 And si.npos >= si.nMax - 1))
End Function


'Public Function WindowUnderCursor(ByVal hWndParent&, Optional xy) 'x,y - координаты внутри hWndParent
''Static cals&
''xMain.DebugPrint 0, "ENTER WindowUnderCursor cals=" & cals
''cals = cals + 1
'Dim h&, h0&, pt As POINTAPI
'GetCursorPos pt
'ScreenToClient hWndParent, pt
'If Not IsMissing(xy) Then xy = MakeDWord(pt.x, pt.Y)
'h = ChildWindowFromPoint(hWndParent, pt.x, pt.Y)
'If h <> 0 And h <> hWndParent Then
'    h0 = WindowUnderCursor(h)
'    If h0 <> 0 Then h = h0
'End If
'WindowUnderCursor = h
''cals = cals - 1
''xMain.DebugPrint 0, "EXIT WindowUnderCursor cals=" & cals
'End Function

Public Function mBrowseForFolder(ownerhWnd&, sTitle$, sStartDir$, Optional nFlags&, Optional sRootDir$) As String
  'Opens a Treeview control that displays the directories in a computer
  Dim lpIDList As Long
  Dim szTitle() As Byte
  Dim szRootDir As String
  Dim sBuffer As String
  Dim tBrowseInfo As BrowseInfo
  Dim ppidl As Long
  m_BrowseInfoCurrentDirectory = sStartDir & vbNullChar
  szTitle = StrConv(sTitle & Chr(0), vbFromUnicode)
  With tBrowseInfo
    .hWndOwner = ownerhWnd
    If Len(sRootDir) Then
        szRootDir = sRootDir
        If IsNum(sRootDir) Then
            .pIDLRoot = sRootDir
        Else
            If SHParseDisplayName(StrPtr(szRootDir), 0, ppidl, ByVal 0&, ByVal 0&) = 0 Then .pIDLRoot = ppidl
        End If
    End If
    .lpszTitle = VarPtr(szTitle(0))
    .ulFlags = nFlags '+ BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
    .lpfnCallback = AddrOf(AddressOf BrowseCallbackProc)    'get address of function.
  End With
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer = Space(260)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    mBrowseForFolder = sBuffer
  Else
    mBrowseForFolder = vbNullString
  End If
End Function

Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
  'Dim lpIDList As Long
  Dim ret As Long
  Dim sBuffer As String
  'On Error Resume Next  'Sugested by MS to prevent an error from
                        'propagating back into the calling process.
  Select Case uMsg
    Case 1 'BFFM_INITIALIZED
      'Call SendMessage(hWnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
      Call SendMessageString(hWnd, &H400 + 102, 1, m_BrowseInfoCurrentDirectory)
    Case 2 'BFFM_SELCHANGED
      sBuffer = Space(260)
      ret = SHGetPathFromIDList(lp, sBuffer)
      If ret = 1 Then
        'Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
        Call SendMessageString(hWnd, &H400 + 100, 0, sBuffer)
      End If
  End Select
  BrowseCallbackProc = 0
End Function


Function mGetFileName$(hWnd&, ByVal sPath$, Optional ByVal sFilter$ = "*.*", Optional Save As Boolean)
Dim ofn As OPENFILENAME, s$, res&
ofn.lStructSize = Len(ofn)
ofn.hWndOwner = hWnd
'ofn.lpstrFilter = "Текстовые файлы" & Chr(0) & "*.TXT;*.BAS" & Chr(0) & "Другие файлы" & Chr(0) & "*.CLS" & Chr(0) & "ВСЕ файлы" & Chr(0) & "*.*" & Chr(0) & Chr(0)
'FILTER = "MSWord,*.doc;*.rtf"
If sFilter = "*.*" Then sFilter = "Все файлы,*.*,,"
ofn.lpstrFilter = Replace(sFilter, ",", Chr$(0))
sPath = Replace(sPath, "..\", xAppPath & "\")
ofn.lpstrInitialDir = xMain.PathFromFile(sPath)
ofn.lpstrFile = xMain.SplitIndex(sPath, "\")
ofn.lpstrFile = ofn.lpstrFile & String$(512, 0)
ofn.nMaxFile = 512
'ofn.flags = &H1806
If Save Then
ofn.Flags = &H806
Else
ofn.Flags = &H1000    'OFN_PATHMUSTEXIST=&H800+OFN_HIDEREADONLY=4 +OFN_FILEMUSTEXIST=&H1000 +OFN_CREATEPROMPT=&H2000 +OFN_OVERWRITEPROMPT=&H2
End If
'OFN_NODEREFERENCELINKS=&H100000
If Save Then res = GetSaveFileName(ofn) Else res = GetOpenFileName(ofn)
'ERRR "mGetFileName"
If res = 0 Then Exit Function
ofn.nFileExtension = InStr(1, ofn.lpstrFile, Chr$(0)) - 1
s = Left$(ofn.lpstrFile, ofn.nFileExtension)
mGetFileName = s
End Function

Function mGetColor&(hWndOwner&, ByVal color&)
mGetColor = NO_INDEX
Dim cc As tCHOOSECOLOR, i&
Static c&(15)
If c(0) = 0 Then
For i = 0 To 15
c(i) = GetSysColor(i)
Next
End If
cc.lStructSize = Len(cc)
cc.hWndOwner = hWndOwner
cc.Flags = &H103 'CC_RGBINIT=1 + CC_FULLOPEN=2 +CC_ANYCOLOR = &H100
cc.lpCustColors = VarPtr(c(0))
cc.rgbResult = color
If ChooseColor(cc) Then mGetColor = cc.rgbResult
End Function

Function mGetFont$(hWndOwner&, ByVal fontSRC$)
Dim hFont&, cf As ChooseFontStruct
Dim lf As LOGFONT
cf.lStructSize = LenB(cf)
hFont = GlobalFontSRC(mncm.lfMenuFont, fontSRC)
gdiGetObject hFont, Len(lf), lf
cf.lpLogFont = VarPtr(lf)
cf.hWndOwner = hWndOwner
cf.Flags = &H141 'CF_SCREENFONTS=1 'CF_INITTOLOGFONTSTRUCT As Long = &H40& 'CF_USESTYLE As Long = &H80&
If ChooseFont(cf) Then mGetFont = LogFontSRC(lf) Else mGetFont = vbNullString
End Function

'============= FILE TIMES ============================

'Function DosTimeToVBDate(dosdate&)
'Dim ft As FILETIME
'DosDateTimeToFileTime HIWORD(dosdate), LOWORD(dosdate), ft
'FileTimeToVBDate ft, DosTimeToVBDate
'End Function
Sub FileTimeToVBDate(pFileTime As FILETIME, pVBDate)
Dim pst As SYSTEMTIME, pST0 As SYSTEMTIME
pVBDate = Empty
If FileTimeToSystemTime(pFileTime, pST0) = 0 Or pFileTime.dwHighDateTime = pFileTime.dwLowDateTime Then Exit Sub
SystemTimeToTzSpecificLocalTime 0, pST0, pst
On Error Resume Next
pVBDate = DateSerial(pst.wYear, pst.wMonth, pst.wDay)
pVBDate = pVBDate + TimeSerial(pst.wHour, pst.wMinute, pst.wSecond)
End Sub

Sub VBDateToFileTime(pVBDate, pFileTime As FILETIME)
Dim pst As SYSTEMTIME, pFT As FILETIME
'pst.wYear = Year(pVBDate): pst.wMonth = Month(pVBDate): pst.wDay = Day(pVBDate)
'pst.wHour = Hour(pVBDate): pst.wMinute = Minute(pVBDate): pst.wSecond = Second(pVBDate)
VariantTimeToSystemTime ByVal pVBDate, pst

If SystemTimeToFileTime(pst, pFT) = 0 Then Exit Sub
LocalFileTimeToFileTime pFT, pFileTime
End Sub

'Function FileTimeToDate(pFileTime As FILETIME)
'Dim pVBDate
'FileTimeToVBDate pFileTime, pVBDate
'FileTimeToDate = pVBDate
'End Function
'Function DateToFileTime(pVBDate) As FILETIME
'Dim pFileTime As FILETIME
'VBDateToFileTime pVBDate, pFileTime
'DateToFileTime = pFileTime
'End Function

'Function pUTCTime(ByVal pLocalTime) As Date 'Convert Local Time to UTCTime
'
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
'pUTCTime = CDate(d) ' the_date
'End Function

'Function pLocalTime(ByVal pUTCTime) As Date 'Convert UTCTime to LocalTime
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
'pLocalTime = CDate(d) ' the_date
'End Function
'============= FILE TIMES ============================


'============= COMPRESS/DECOMPRESS ============================
Function LongToString$(ByVal n&)
Dim lw%, hw%
lw = LOWORD(n): hw = HIWORD(n)
LongToString = Chr$(HIBYTE(hw)) & Chr$(LOBYTE(hw)) & Chr$(HIBYTE(lw)) & Chr$(LOBYTE(lw))

'Dim s$
's = "0000"
''CopyMemory 0 + StrPtr(s), 0 + n, 4
'PutMem4 ByVal StrPtr(s), n
'Debug.Print s
End Function

Function StringToLong&(ByVal s$, Optional ByVal bNoReverse As Boolean)
'Dim i&, s0$, b&
's = s & String(4, 0)
'For i = 1 To 4
'b = Asc(Mid(s, i, 1))
'If i = 1 Then b = b And 127
's0 = s0 & Right$("00" & Hex$(b), 2)
'Next
'StringToLong = "&H" & s0

Dim c() As Byte
's = Left$(s & String(4, 0), 4)
If Len(s) > 3 Then
c = StrConv(IIf(bNoReverse, s, StrReverse(s)), vbFromUnicode)
GetMem4 c(0), StringToLong
'Debug.Print Hex(StringToLong)
End If
End Function

Public Function gCompress(ByRef vToCompress)
'If LenB(StringToCompress) = 0 Then Exit Function

Dim src() As Byte, dest() As Byte, srcLen&, destLen&, res&

Dim bString As Boolean
If VarType(vToCompress) And (vbArray + vbByte) Then
    src = vToCompress
ElseIf VarType(vToCompress) = vbString Then
    src = StrConv(vToCompress, vbFromUnicode)  'COMPRESSED SOURCE
    bString = True
Else
    Exit Function
End If

'src = StrConv(StringToCompress, vbFromUnicode)

srcLen = UBound(src) + 1
destLen = srcLen * 1.01 + 12
ReDim dest(7 + destLen) As Byte
res = zlib_compress(dest(8), destLen, src(0), srcLen)
If res Then
    Err.Raise res, "zlib_compress"
Else 'xzip
    ReDim Preserve dest(7 + destLen) As Byte 'Усекаем массив
    Dim b0() As Byte
    b0 = StrConv(LongToString(destLen + 8) & LongToString(srcLen), vbFromUnicode)
    GetMem8 b0(0), dest(0)
    If bString Then
        gCompress = StrConv(dest, vbUnicode)
    Else
        gCompress = dest
    End If
End If

End Function
Public Function gIsCompressed(ByRef CompressedString) As Boolean

If LenB(CompressedString) < 15 Then Exit Function
gIsCompressed = LenB(CompressedString) >= GetLong(CompressedString, 0)

'If Len(CompressedString) = 0 Then Exit Function
'gIsCompressed = Len(CompressedString) >= Abs(StringToLong(Left(CompressedString, 4)))
End Function

Function GetLong(src, pos&, Optional bReverse As Boolean) As Long
Dim i&, b(3) As Byte
If IsArray(src) Then
    For i = 0 To 3
    b(IIf(bReverse, i, 3 - i)) = src(pos + i)
    Next
Else
    For i = 0 To 3
    b(3 - i) = Asc(Mid(src, pos + i + 1, 1))
    Next
End If
GetMem4 b(0), GetLong
End Function

Public Function gDecompress(ByRef CompressedString, Optional o)

If LenB(CompressedString) > 10 Then
Dim dest() As Byte, src() As Byte, srcLen&, destLen&, res&
Dim i&, n&, bString As Boolean

i = VarType(CompressedString)

If i And (vbArray + vbByte) Then
    src = CompressedString
ElseIf i = vbString Then
    src = StrConv(CompressedString, vbFromUnicode)  'COMPRESSED SOURCE
    bString = True
Else
End If


destLen = UBound(src) + 1
n = GetLong(src, 0)

'If destLen = n Then 'xzlib
If destLen >= n And destLen <= n + 3 Then 'xzlib /+LOADRES
    n = 8
    destLen = GetLong(src, 4)
ElseIf n = &H1F8B0800 Then  'gzip/deflate
    n = 10
    destLen = GetLong(src, UBound(src) - 3, 1)
ElseIf (n And &H789C0000) = &H789C0000 Then   'zlib  xњ
    n = 0
    destLen = destLen * 3
Else
    gDecompress = CompressedString
    Exit Function
End If

srcLen = UBound(src) + 1 - n

'MsgBox "gDecompress " & VarInfo(CompressedString) & vbCrLf & " LEN0..3=" & n & vbCrLf & " destLen=" & destLen & vbCrLf & " srcLen=" & srcLen

If n = 10 Then
    Dim zs As z_stream
    zs.next_in = VarPtr(src(0)) + n
    zs.avail_in = srcLen
    res = zlib_inflateInit2(zs, -15, StringFromPointer(zlib_zlibVersion), Len(zs))
    If res = 0 Then
        ReDim dest(destLen - 1) As Byte 'BUFFER FOR UNCOMPRESSED DATA
        zs.next_out = VarPtr(dest(0))
        zs.avail_out = destLen
    
        res = zlib_inflate(zs, 0)
        If res = 1 Then res = 0
        zlib_inflateEnd zs
    End If
Else

    i = destLen
    Do
        destLen = i
        On Error Resume Next
        ReDim dest(destLen - 1) As Byte 'BUFFER FOR UNCOMPRESSED DATA
        res = zlib_uncompress(dest(0), destLen, src(n), srcLen) ', 31)
        If Err Then ReDim dest(0): Debug.Print "zlib_uncompress", Err.Description: Err.Clear
        i = i + destLen
    Loop While res = -5 'Z_BUF_ERROR

End If


If res Then 'ERROR
'    Err.Raise res, "zlib_uncompress"
    gDecompress = CompressedString
    Debug.Assert False
Else

'    If n = 10 Then
'        res = MultiByteToWideChar(CP_UTF8, 0, VarPtr(dest(0)), destLen, 0, 0)
'        gDecompress = String$(res, 0)
'        res = MultiByteToWideChar(CP_UTF8, 0, VarPtr(dest(0)), destLen, StrPtr(gDecompress), res)
    If bString Then
        gDecompress = StrConv(dest, vbUnicode)
    Else
        gDecompress = dest
    End If
End If

End If

End Function


'Public Function GetAdler32(ByVal Data As String) As Long
'    Dim crc As Long
'    ' Get initial value
'    crc = adler32(0, ByVal 0&, 0)
'    crc = adler32(crc, Data, Len(Data))
'    GetAdler32 = crc
'End Function
'
'Public Function GetCRC32(ByVal Data As String) As Long
'    Dim crc As Long
'    ' Get initial value
'    crc = crc32(0, ByVal 0&, 0)
'    crc = crc32(crc, Data, Len(Data))
'    GetCRC32 = crc
'End Function


'============= COMPRESS/DECOMPRESS ============================



'=================== SYSTRAY ==================
Public Sub SystrayOn(frmhWnd As Long, frmIcon As Long, IconTooltipText As String) 'Adds Icon to SysTray
Dim vbTray  As NOTIFYICONDATA
    With vbTray
        .cbSize = Len(vbTray)
        .hWnd = frmhWnd
        .UID = frmhWnd
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .szTip = IconTooltipText & vbNullChar
        .hIcon = frmIcon
    End With
Call Shell_NotifyIcon(NIM_ADD, vbTray)
'App.TaskVisible = False
End Sub

Public Sub SystrayOff(frmhWnd As Long) 'Removes Icon from SysTray
Dim vbTray  As NOTIFYICONDATA
With vbTray
    .cbSize = Len(vbTray)
    .hWnd = frmhWnd
    .UID = frmhWnd
End With
Call Shell_NotifyIcon(NIM_DELETE, vbTray)
End Sub

Public Sub pTrayBalloon(frmhWnd As Long, frmIcon&, mIcon As Long, Message As String, Title As String)
Dim vbTray  As NOTIFYICONDATA
    'Set a Balloon tip on Systray
    'Call RemoveBalloon(frm), This removes any current Balloon Tip that is active.
    'If you want Balloon Tips to "Stack up" and display in sequence
    'after each times out (or you click on the Balloon Tip to clear it),
    'comment out the Call below.
    Call RemoveBalloon(frmhWnd)
    With vbTray
        .cbSize = Len(vbTray)
        .hWnd = frmhWnd
        .UID = frmhWnd
        .uFlags = NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY Or NIF_ICON  'Or   'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frmIcon
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Message & Chr(0)
        .szInfoTitle = Title & Chr(0)
        'Choose the message icon below, NIIF_NONE, NIIF_WARNING, NIIF_ERROR, NIIF_INFO
        .dwInfoFlags = mIcon '* 16 'NIIF_INFO 0..7
    End With
    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)
End Sub

Public Sub RemoveBalloon(frmhWnd As Long)
Dim vbTray  As NOTIFYICONDATA
    'Kill any current Balloon tip on screen for referenced form
    With vbTray
        .cbSize = Len(vbTray)
        .hWnd = frmhWnd
        .UID = frmhWnd
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = 0 'frm.Icon
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Chr(0)
        .szInfoTitle = Chr(0)
        .dwInfoFlags = 0 'NIIF_NONE
    End With
    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)
End Sub
'=================== SYSTRAY ==================

'================ TOOLTIP ============================
Public Sub ToolTipAppendTip(ByVal hWnd&, Text$, rc As RECT)
If mToolTip = 0 Then pCreateToolTipWindow mToolTip, 0
Dim ti As TOOLINFO
ti.cbSize = Len(ti)
ti.hWnd = hWnd
'Debug.Print "ToolTipAppendTip " & Text
If Len(Text) Then 'ADDTOOL
    'Debug.Print Hex(hWnd) & " TTM_ADDTOOLA = " & Text
    ti.lpszText = StrConv(Text, vbFromUnicode)
    ti.uFlags = &H10  'TTF_SUBCLASS = &H10
    If rc.Right = rc.Left And rc.Top = rc.Bottom Then GetClientRect hWnd, ti.lpRect Else ti.lpRect = rc
    SendMessage mToolTip, &H404, 0&, ti 'TTM_ADDTOOLA = WM_USER + 4
ElseIf SendMessage(mToolTip, &H40D, 0&, 0&) > 0 Then 'TTM_GETTOOLCOUNT As Long = (WM_USER + 13)
    While SendMessage(mToolTip, &H408, 0, ti)  ' TTM_GETTOOLINFOA As Long = (WM_USER + 8)
        'Debug.Print Hex(hWnd) & " TTM_DELTOOLA =" & ti.lpszText 'StrConv(ti.lpszText, vbUnicode)
        SendMessage mToolTip, &H405, 0&, ti 'TTM_DELTOOLA As Long = (WM_USER + 5)
    Wend
End If
End Sub

Public Sub ToolTipTrackTip(x&, Y&, sText$)
If mTip = 0 Then pCreateToolTipWindow mTip, 0
Dim ti As TOOLINFO
ti.cbSize = Len(ti)
If SendMessage(mTip, &H40E, 0&, ti) Then 'TTM_ENUMTOOLSA As Long = (WM_USER + 14)
    ti.lpszText = StrConv(sText, vbFromUnicode)
    SendMessage mTip, &H40C, 0&, ti 'TTM_UPDATETIPTEXTA = &H400 + 12
Else 'ADDTOOL
    ti.hWnd = 0
    ti.uFlags = &HA0 'TTF_TRACKABSOLUTE As Long = &H20+&H80
    ti.hInst = App.hInstance
    ti.lpszText = StrConv(sText, vbFromUnicode)
    SendMessage mTip, &H404, 0&, ti 'TTM_ADDTOOLA = WM_USER + 4
End If
SendMessage mTip, &H411, 0&, ti  'Гасим TIP TTM_TRACKACTIVATE = WM_USER + 17
SendMessage mTip, &H412, 0&, ByVal MakeDWord(x, Y) 'Меняем координаты TIPA TTM_TRACKPOSITION = WM_USER + 18
SendMessage mTip, &H411, True, ti 'Показываем TIP TTM_TRACKACTIVATE = WM_USER + 17



mActiveToolTip = mTip
End Sub

Public Sub ToolTipTrackBalloon(ByVal x&, ByVal Y&, sText$, sTitle$, nIcon&)
If mBalloon = 0 Then pCreateToolTipWindow mBalloon, 1
Dim ti As TOOLINFO
ti.cbSize = Len(ti)
ShowWindow mBalloon, 0
If SendMessage(mBalloon, &H40E, 0&, ti) Then 'TTM_ENUMTOOLSA As Long = (WM_USER + 14)
    ti.lpszText = StrConv(sText, vbFromUnicode) 'Text
    SendMessage mBalloon, &H40C, 0&, ti 'TTM_UPDATETIPTEXTA = &H400 + 12
Else 'ADDTOOL
    ti.hWnd = 0
    ti.uFlags = &HA0 'TTF_TRACKABSOLUTE =80  TTF_TRACK=20  TTF_CENTERTIP=2
    ti.hInst = App.hInstance
    ti.lpszText = StrConv(sText, vbFromUnicode)
    SendMessage mBalloon, &H404, 0&, ti 'TTM_ADDTOOLA = WM_USER + 4
End If
'SendMessage mBalloon, &H411, 0, ti 'Гасим TIP TTM_TRACKACTIVATE = WM_USER + 17
SendMessageString mBalloon, &H420, nIcon, ByVal sTitle      'TTM_SETTITLE = WM_USER + 32
SendMessage mBalloon, &H412, 0, ByVal MakeDWord(x, Y)  'Меняем координаты TIPA TTM_TRACKPOSITION = WM_USER + 18
SendMessage mBalloon, &H411, True, ti 'Показываем TIP TTM_TRACKACTIVATE = WM_USER + 17
mActiveToolTip = mBalloon
End Sub

Public Function ToolTipHide()
If mActiveToolTip = 0 Then Exit Function
Dim ti As TOOLINFO
ti.cbSize = Len(ti)
SendMessage mActiveToolTip, &H408, 0, ti '' TTM_GETTOOLINFOA As Long = (WM_USER + 8)
SendMessage mActiveToolTip, &H411, False, ti 'TTM_TRACKACTIVATE = WM_USER + 17
mActiveToolTip = 0
End Function

Private Sub pCreateToolTipWindow(hWnd&, bBalloon As Boolean)
Dim dws&
dws = 0 '&H80000003 'Or WS_POPUP Or WS_BORDER  'Or TTS_ALWAYSTIP Or TTS_NOPREFIX TTS_NOFADE=&H20 TTS_CLOSE=&h80
If bBalloon Then dws = dws Or &H40 'Else dws = dws Or WS_BORDER
hWnd = CreateWindowEx(0, "tooltips_class32", vbNullString, dws, 0, 0, 0, 0, 0, 0&, App.hInstance, 0&)
SendMessage hWnd, &H418, 0, ByVal 400&  'TTM_SETMAXTIPWIDTH=WM_USER + 24
'SetWindowPos hWnd, -1, 0&, 0&, 0&, 0&, &H13
If Not bBalloon Then SendMessage hWnd, &H401, True, ByVal 0 'TTM_ACTIVATE = WM_USER + 1
End Sub
'================ TOOLTIP ============================



Function SEPrivilege(ByVal sPrivilege$, Optional ByVal bEnabled = -1) As Long
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
If L_(bEnabled) > -1 Then v = Sgn(L_(bEnabled))
'If Len(sPrivilege) = 0 Then Exit Function
If OpenProcessToken(GetCurrentProcess(), IIf(v > -1, &H20, 0) Or 8, hToken) Then
    If LookupPrivilegeValue(0, sPrivilege, LUID) Then
        tps.PrivilegeCount = 1: tps.Control = 1: tps.Privileges(0).pLuid = LUID
        If PrivilegeCheck(hToken, tps, ret) Then
             If v > -1 Then 'LET
                 If (ret <> v) Then
                    tp.PrivilegeCount = 1: tp.Privileges(0).pLuid = LUID
                    tp.Privileges(0).Attributes = IIf(v, 2, 0)
                    If AdjustTokenPrivileges(hToken, False, tp, LenB(tp), 0, 0) Then
                        tps.Privileges(0).Attributes = 0: ret = -1
                        If PrivilegeCheck(hToken, tps, ret) Then 'get updated privilege value
                            'SEPrivilege = ret
'                            Debug.Print "LET "; sPrivilege; " ="; v, v = ret
                        End If
                    End If
                Else
'                    Debug.Print "NO LET "; sPrivilege; " ="; v, v = ret
                End If
            Else
'                Debug.Print "GET "; sPrivilege; " ="; ret
            End If
        End If
    End If
    CloseHandle hToken
End If
SEPrivilege = ret
End Function



'Sub PrinterMargins(psd As PAGESETUPDLG_struct) 'ptPaperSize As POINTAPI, rtMargin As RECT)
'Dim m As RECT
'Printer.ScaleMode = vbTwips
'Printer.PaperSize = 9
'With psd
'    Printer.Orientation = IIf(.ptPaperSize.x > .ptPaperSize.y, 2, 1)
'    m.Left = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbHimetric)
'    If .rtMargin.Left < m.Left Then .rtMargin.Left = xMain.nRound(m.Left + 50, 100)
'    m.Right = .ptPaperSize.x - (m.Left + Printer.ScaleX(Printer.ScaleWidth, vbTwips, vbHimetric))
'    If .rtMargin.Right < m.Right Then .rtMargin.Right = xMain.nRound(m.Right + 50, 100)
'    m.Top = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbHimetric)
'    If .rtMargin.Top < m.Top Then .rtMargin.Top = xMain.nRound(m.Top + 50, 100)
'    m.Bottom = .ptPaperSize.y - (m.Top + Printer.ScaleY(Printer.ScaleHeight, vbTwips, vbHimetric))
'    If .rtMargin.Bottom < m.Bottom Then .rtMargin.Bottom = xMain.nRound(m.Bottom + 50, 100)
'End With
'End Sub

#If DragDrop Then


Function IDO_EnumFormatS(pDataObj As IDataObject, bDir As Long)
On Error Resume Next
Dim e As IEnumFORMATETC: Set e = pDataObj.EnumFormatEtc(bDir) 'bDir: 1=DIR_GETData, 2=DIR_SETData
Dim s$, p&, fs$, fm As FORMATETC
Do
e.Next 1, fm, p: If p = 0 Then Exit Do
If Len(fs) Then fs = fs & ","
fs = fs & fm.cfFormat
s = xMain.ClipFormatName(fm.cfFormat)
If Len(s) Then fs = fs & "," & s
'Debug.Print "EnumFormat " & fm.cfFormat & "=" & s '& ", TYMED=" & fm.TYMED & ", dwAspect=" & fm.dwAspect
Loop
IDO_EnumFormatS = Split(fs, ",")
End Function


Sub IDO_SetData(pDataObj As IDataObject, vv)
'vv=VariantArray(fns0,v0,fns1,v1,...fnsN,vN)
Dim i&, v$, nf&
Dim fm As FORMATETC, sm As STGMEDIUM
Dim hMem As Long, ptr As Long, b() As Byte, sz&
On Error Resume Next
Do While i <= UBound(vv)
    v = S_(vv(i))
    If Not IsNum(v) Then If Len(v) > 3 Then v = LOWORD(RegisterClipboardFormat(v))
    fm.cfFormat = v
    If fm.cfFormat Then
        Erase b
        b = Cast(vv(i + 1), vbArray + vbByte)
        sz = UBound(b) + 1
        hMem = GlobalAlloc(GPTR, sz)
        If hMem Then
            ptr = GlobalLock(hMem)
            CopyMemory ByVal ptr, b(0), sz
            GlobalUnlock hMem
            sm.TYMED = TYMED_HGLOBAL
            sm.Data = ptr
            fm.dwAspect = DVASPECT_CONTENT: fm.TYMED = TYMED_HGLOBAL: fm.lIndex = -1
            'Err.Clear
            pDataObj.SetData fm, sm, 1
        End If
        'If Err Then Debug.Print "pDataObj.SetData err=", Err, Hex(Err), Err.Description
    End If
i = i + 2
Loop
End Sub

Function IDO_GetData(pDataObj As IDataObject, ByVal fmts) 'Return CPArray(fns0,v0,fns1,v1,...fnsN,vN)
Dim afs: afs = IDO_EnumFormatS(pDataObj, 1) 'avilable formats list string array
Dim i&, n&, fs
If IsArray(fmts) Then fs = fmts Else fmts = S_(fmts): If fmts = "*" Then fs = afs Else fs = Split(fmts, ",")
Dim fm As FORMATETC, sm As STGMEDIUM
Dim v: v = Array() 'RESULT
ReDim v(-1 To -1)
Dim nSize&, ptr&, buf() As Byte
On Error Resume Next
Do While i <= UBound(fs)
If Not IsNum(fs(i)) Then If Len(fs(i)) Then fs(i) = LOWORD(RegisterClipboardFormat(fs(i)))
If gFindIndex(fs, fs(i)) = i Then
    If gFindIndex(afs, fs(i)) > -1 Then
        n = UBound(v) + 2
        ReDim Preserve v(n)
        fm.cfFormat = fs(i): fm.TYMED = TYMED_HGLOBAL: fm.dwAspect = DVASPECT_CONTENT: fm.lIndex = -1: fm.pDVTARGETDEVICE = 0
        v(n - 1) = fs(i)
        pDataObj.GetData fm, sm
        If Err Then
            Err.Clear
        Else
            nSize = GlobalSize(sm.Data): ptr = 0
            If nSize > 0 Then ptr = GlobalLock(sm.Data)
            If ptr Then
                ReDim buf(0 To nSize - 1&)
                CopyMemory buf(0), ByVal ptr, nSize
                GlobalUnlock sm.Data
                v(n) = buf
            End If
        End If
    End If
End If
i = i + 1
Loop
IDO_GetData = v
End Function


'Function NewDragDropHelper() As DragDropHelper
'Dim dhiid As UUID
'Dim dthiid As UUID
'Const CLSID_DragDropHelper = "{4657278A-411B-11D2-839A-00C04FD918D0}"
'Const IID_IDropTarget = "{4657278B-411B-11D2-839A-00C04FD918D0}"
'Call CLSIDFromString(CLSID_DragDropHelper, dhiid)
'Call CLSIDFromString(IID_IDropTarget, dthiid)
'Call CoCreateInstance(dhiid, 0&, CLSCTX_INPROC_SERVER, dthiid, NewDragDropHelper)
'End Function

Public Function IDS_QueryContinueDrag(ByVal This As IDropSource, ByVal fEscapePressed As Long, ByVal grfKeyState As Long) As Long
'Debug.Print "IDropSource_QueryContinueDrag", fEscapePressed, grfKeyState
If fEscapePressed Then
   IDS_QueryContinueDrag = DRAGDROP_S_CANCEL
ElseIf (grfKeyState And 1) <> 1 Then
   IDS_QueryContinueDrag = DRAGDROP_S_DROP
End If
End Function

'Public Function IDS_GiveFeedback(ByVal This As IDropSource, ByVal dwEffect As DROPEFFECTS) As Long
'Debug.Print "IDropSource_GiveFeedback", dwEffect
'IDS_GiveFeedback = DRAGDROP_S_USEDEFAULTCURSORS 'S_OK
'End Function

Public Function NewVTBLEntry(ByVal pObj As Long, ByVal EntryNumber As Integer, ByVal lpfn As Long) As Long
    Dim lOldAddr As Long
    Dim lpVtableHead As Long
    Dim lpfnAddr As Long
    Dim lOldProtect As Long
    GetMem4 ByVal pObj, lpVtableHead
    lpfnAddr = lpVtableHead + (EntryNumber - 1) * 4
    GetMem4 ByVal lpfnAddr, lOldAddr
    Call VirtualProtect(ByVal lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect)
    CopyMemory ByVal lpfnAddr, lpfn, 4
    Call VirtualProtect(ByVal lpfnAddr, 4, lOldProtect, lOldProtect)
    NewVTBLEntry = lOldAddr
End Function

#End If

'Function SysAllocString(ByVal sPtr&) As String
'PutMem4 SysAllocString, SysAllocStringPtrPtr(sPtr) 'vb6 call SysFreeString
'End Function
'
'Function SysAllocStringLenS(ByVal sPtr&, ByVal cb&) As String
'PutMem4 SysAllocStringLenS, SysAllocStringLen(ByVal sPtr, cb) 'vb6 call SysFreeString
'End Function

'Sub AsyncCall(ByVal hWnd&, ByVal uMsg&, ByVal wParam&, ByVal lParam&)
'If arCmd.count = 0 Then Exit Sub
'On Error Resume Next
'Dim cmd: cmd = arCmd(1)
'arCmd.Remove 1
''DoEvents
'
'Dim obj As Object: Set obj = cmd(0)
'If obj Is Nothing Then
'    Select Case Nz(cmd(0))
'    Case "Quit": xMainWnd.Destroy
''    Case "EndThread"
''        'PostQuitMessage 0
''        PostThreadMessage App.ThreadID, WM_DESTROY, 0, 0
'    Case "CloseForm"
'        hWnd = cmd(1): If IsWindow(hWnd) Then SendMessage hWnd, WM_CLOSE, 0, 0: If IsWindow(hWnd) Then PostMessage hWnd, WM_CLOSE, 0, 0: Exit Sub
'    Case "Post"
'        hWnd = cmd(1): uMsg = cmd(2): PostMessage hWnd, uMsg, 0&, 0&
'    Case "Execute"
'        Dim sc As ScriptControl
'        Set sc = cmd(2): If sc Is Nothing Then Set sc = xMain.VBScript("*+", Empty)
'        sc.AddCode xMain.vbRemoveComents(cmd(1)) 'Run code
'    End Select
'Else
'    Dim i&, args(), nm$: nm = cmd(1)
'    If UBound(cmd) > 1 Then
'        ReDim args(UBound(cmd) - 2)
'        For i = 0 To UBound(args): args(i) = cmd(i + 2): Next
'    End If
'    Call rtcCallByName(obj, StrPtr(nm), VbMethod, args)
'End If
'
''uMsg = 1
''Do While uMsg < arCmd.count
''    cmd = arCmd(uMsg)
''    If cmd(0) = "CloseForm" Or cmd(0) = "Post" Then If IsWindow(cmd(1)) Then arCmd.Remove uMsg: uMsg = uMsg - 1
''    uMsg = uMsg + 1
''Loop
'If arCmd.count = 0 Then tm.Interval = 0
''If arCmd.count Then SendMessageCallBack GetDesktopWindow, WM_USER, 0, 0, AddressOf AsyncCall, 0
'End Sub
