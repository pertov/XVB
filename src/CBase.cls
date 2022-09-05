VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text


'Private Declare Function InternetSetStatusCallback Lib "wininet.dll" (ByVal hInternetSession As Long, ByVal lpfnCallBack As Long) As Long

'====== FTP ========
'Private hFtpOpen As Long
'Private hFtpConnection As Long
'Private mFtpServer As String
'Private Type WIN32_FIND_DATA
'        dwFileAttributes As Long
'        ftCreationTime As FILETIME 'Currency
'        ftLastAccessTime As FILETIME 'Currency
'        ftLastWriteTime As FILETIME 'Currency
'        nFileSizeHigh As Long
'        nFileSizeLow As Long
'        dwReserved0 As Long
'        dwReserved1 As Long
'        cFileName As String * 260
'        cAlternate As String * 14
'End Type

'Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
'Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
'Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
'Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
'Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Boolean
'
'Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszOldName As String, ByVal lpszNewName As String) As Boolean
'Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszName As String) As Boolean
'Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszName As String) As Boolean
'Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
'Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
'Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
'====== FTP ========
'Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean

'Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hhttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
'Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlW" (ByVal hOpen As Long, ByVal sURL As Long, ByVal sHeaders As Long, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'Private Declare Sub InternetTimeToSystemTime Lib "wininet.dll" (ByVal lpszTime As String, ByRef pst As SYSTEMTIME, ByVal dwReserved As Long)



'Private mSQLStorage$
Private mUID&, mBase$
Private mIP&, mIP2&
Private mSID&, mVersion&, mLevel&
'Private mFirm&
Private cp As CParam
Public Vars As CParam

'Implements IBindStatusCallback 'initialize the IBindStatusCallback interface...
'Private mBinding As IBinding  ' set-up so I can access as I wish
'Public xStatus As xControl


Private Sub Class_Initialize()
Set cp = New CParam
Set Vars = New CParam
End Sub
Private Sub Class_Terminate()
Set cp = Nothing
Set Vars = Nothing
'Set xStatus = Nothing
'CloseFtpConnection
'Debug.Print "CBASE Class_Terminate()"
End Sub



Public Property Get Name(): Name = mBase: End Property
Public Property Let Name(v): mBase = S_(v): End Property
Property Get UID(): UID = mUID: End Property
Property Let UID(newUID): mUID = L_(newUID): End Property
Property Get IP(): IP = mIP: End Property
Property Let IP(v): mIP = L_(v): End Property
Property Get IP2(): IP2 = mIP2: End Property
Property Let IP2(v): mIP2 = L_(v): End Property

Property Get SID(): SID = mSID: End Property
Property Let SID(newSID): mSID = L_(newSID): End Property

Property Get Version(): Version = mVersion: End Property
'Property Let Version(v)
'mVersion = aVal(v)
'End Property
'
'Property Get Level()
'Level = mLevel
'End Property
'Property Let Level(v)
'mLevel = aVal(v)
'End Property


Public Function LogIN(ID, ar) As Long
'User ID,User Name,User Role,User Level,User Version
'CloseFtpConnection
Dim i&, n&, s$
If VarType(ar) And vbArray Then
For i = 0 To UBound(ar, 2)
s = "u" & ar(0, i)
cp(s & "\NAM") = ar(1, i)
cp(s & "\ROL") = ar(2, i)
cp(s & "\LVL") = ar(3, i)
cp(s & "\VER") = ar(4, i)
'cp.json("u" & ar(0, i)) = "{nam:" & quot(ar(1, i)) & "",rol:" & ar(2, i) & ",lvl:" & ar(3, i) & ",ver:" & nz(ar(4, i),"null") & "}"
If ID = ar(0, i) Then mVersion = L_(ar(4, i)): mLevel = L_(ar(3, i))
Next
'xMain.MsBox cp.json("!")
LogIN = -1
End If
mUID = ID
xMain.Access UserRole
End Function


Public Function UserID&()
UserID = UID
End Function

Public Function UserLevel&(Optional ByVal ID = -2)
UserLevel = L_(cp("u" & IIf(ID + 0 = -2, mUID, ID) & "\LVL"))
End Function
Public Function UserRole&(Optional ByVal ID = -2)
UserRole = L_(cp("u" & IIf(ID + 0 = -2, mUID, ID) & "\ROL"))
End Function
Public Function UserVer&(Optional ByVal ID = -2)
UserVer = L_(cp("u" & IIf(ID + 0 = -2, mUID, ID) & "\VER"))
End Function
Public Function UserName$(Optional ByVal ID = -2)
UserName = S_(cp("u" & IIf(ID + 0 = -2, mUID, ID) & "\NAM"))
'If Err Then Debug.Print "Base.UserName$ " & Err.Description
End Function

'Property Get SQLStorage()
'SQLStorage = mSQLStorage
'End Property
'Property Let SQLStorage(v)
''mSQLStorage=SQL STATEMENT>>>  SELECT dccode, dcnm, dcbase, dclvl, dcver, dcid FROM ds WHERE dcver=*version* and dsbase= *base* and dskey=*key*
''mSQLStorage=SQL STATEMENT>>>  SELECT src,ver,nm,id,szid FROM sz WHERE nm='*key*' AND ver*ver* ORDER BY ver DESC
''*base* *key* = replacement args
'mSQLStorage = Nz(v)
'End Property

'Function ClearPDOC(Optional wDir$) As Long 'Удаление всех файловых щаблонов из укзанного подкаталога
''+Засосать новые версии файлов из указаного каталога
'If xMain.ConState = 0 Then Exit Function
'Dim s$, p$, sql$
'Dim ar, tmCreate As Double, tmModify As Double, i&, ID&, key$
'On Error Resume Next
'
''  Base.SQLStorage="SELECT src,ver,nm,id,szid FROM sz WHERE nm='*key*' AND ver*ver* ORDER BY ver DESC"
'
''If Len(mSQLStorage) = 0 Then Exit Function
'p = xAppPath & "\" & NzS(wDir, mBase) & "\"
's = Dir(p & "*.xc")
'While Len(s)
'    tmModify = xMain.FileTimes(p & s, 4) 'Время изменения файла
'    tmCreate = xMain.FileTimes(p & s, 1) 'Время создания файла
'    key = Join(ArrayDeleteLine(Split(s, "."), 100), ".")
'    'sql = Replace(Replace(mSQLStorage, "*ver*", "<99999"), "*key*", key)
'    sql = "SELECT ver, id FROM sz WHERE nm='" & key & "' AND ver<99999 ORDER BY ver DESC"
'    'ar = xMain.GetRows(sql, 0, Array(1))  'Список версий сущности
'    ar = xMain.GetRows(sql, 0, Array(0))  'Список версий сущности
'    'ID = IF0(aVal(xMain.RunSQL(sql, 3)), 0)
'    ID = IF0(aVal(xMain.RunSQL(sql, 1)), 0)
'    If ArrayDims(ar) = 2 Then ar = Split(Join2(ar, ":", ""), ":") Else ar = Array()
'    If Abs(tmCreate - tmModify) > 0.0001 Then 'есть изменения файла >8 second
'        'Проверим наличие такой версии в базе
'        If gFindIndex(ar, "" & tmCreate) = -1 Then 'Засосать новую версию файла если сущность зарегистрированна
'            'sql = Replace(Replace(mSQLStorage, "*ver*", "=-55"), "*key*", key)
'            sql = "SELECT src,ver, nm, id, szid FROM sz WHERE szid=-1"
'            xMain.SetFields sql, 1, xMain.ReadFile(p & s), tmCreate, key, ID
'            Kill p & s 'Если в базе есть такое то смело удаляем
'            ClearPDOC = ClearPDOC + 1
'        End If
'    Else
'        If ID > 0 Then If gFindIndex(ar, "" & tmCreate) > -1 Then Kill p & s 'Если в базе есть такое то смело удаляем
'    End If
'    s = "": s = Dir
'Wend
'End Function

'Function GetPDOC(ByVal key) As CParam 'Загрузка актуальной версии сущности
''Ищем сущность в рабочем каталоге
''Засасываем обновленную сущность из каталога в базу
''Выпундриваем актуальную сущность в каталог
''Загружаем сущность из каталога
''  Base.SQLStorage="SELECT src,ver,nm,id,szid FROM sz WHERE nm='*key*' AND ver*ver* ORDER BY ver DESC"
'If xMain.ConState = 0 Then Exit Function
'Dim cp As New CParam
'Dim pth$, sql$, n&
'Dim src, tmBase As Double, tmCreate As Double, tmModify As Double, ID&
'On Error Resume Next
''If Len(mSQLStorage) = 0 Then Exit Function
'pth = xAppPath & "\" & mBase & "\" & key & ".xc"
'
'n = FileLen(pth)
'If n Then
'    tmCreate = xMain.FileTimes(pth, 1) 'Время создания сущности
'    tmModify = xMain.FileTimes(pth, 4) 'Время изменения сущности
'    If Abs(tmModify - tmCreate) > 0.0001 Then 'Есть изменения надо засосать
'
''sql = Replace(Replace("" & mSQLStorage, "*ver*", "<99999"), "*key*", key)
'sql = "SELECT ver, id FROM sz WHERE nm='" & key & "' AND ver<99999 ORDER BY ver DESC"
''ID = xMain.RunSQL(sql, 3)  'Получем ID сущности
'ID = IF0(aVal(xMain.RunSQL(sql, 1)), 0)
'
'        'sql = Replace(Replace("" & mSQLStorage, "*ver*", "=-55"), "*key*", key)
'        sql = "SELECT src,ver, nm, id, szid FROM sz WHERE szid=-1"
'        xMain.SetFields sql, 1, xMain.ReadFile(pth), tmModify, key, ID 'Записываем измененную сущность в базу
'        xMain.FileTimes(pth, 1) = tmModify
'    End If
'End If
''sql = Replace(Replace("" & mSQLStorage, "*ver*", "<=" & IF0(mVersion, 99999) + 1), "*key*", key)
'sql = "SELECT ver, szid  FROM sz WHERE nm='" & key & "' AND ver<=" & (IF0(mVersion, 99999) + 1) & " ORDER BY ver DESC"
''tmBase = aVal(xMain.RunSQL(sql, 0)) 'Читаем версию сущности которая должна работать
'Dim szid&
'tmBase = 0: szid = 1
'xMain.RunSQL sql, tmBase, szid 'Читаем версию и код записи сущности которая должна работать
'
'If n > 0 And tmBase > 0 And Abs(tmBase - tmModify) > 0.0001 Then n = 0: Kill pth   'Эта версия файла слишком новая или старая
'If n = 0 Then ' FROM DB to FILE
'    sql = "SELECT src, ver  FROM sz WHERE szid=" & szid
'    If xMain.GetFields(sql, src, tmBase) Then
'        cp.Value = ToVariant(pDecompress(Nz(src)))
'        If cp.ParamCount Then
'            cp.Save pth 'SAVE TO FILE
'            xMain.FileTimes(pth, 1) = tmBase
'            xMain.FileTimes(pth, 4) = tmBase
'        End If
'    End If
'Else
'    cp.Load pth
'End If
'cp.Param("SourcePath") = pth
'Set GetPDOC = cp
'End Function



'=======================================================

''=======================================================
''==================== FTP ===============================
'Public Property Get FtpServer$()
'If hFtpConnection Then FtpServer = mFtpServer
'End Property
'Public Function OpenFtpConnection(ByVal sServer As String, ByVal sUser As String, ByVal sPassword As String) As Boolean
'    If hFtpConnection <> 0 Then InternetCloseHandle hFtpConnection: hFtpConnection = 0
'    hFtpOpen = InternetOpen("XVB FTPClient", 0&, "", "", 0)
''    hFtpOpen = InternetOpen("TotalCommander", 0&, "", "", 0)
'
'    If hFtpOpen = 0 Then Exit Function
'    hFtpConnection = InternetConnect(hFtpOpen, sServer, 0&, sUser, sPassword, 1&, &H8000000, 0)  'INTERNET_FLAG_PASSIVE = &H8000000
'    OpenFtpConnection = hFtpConnection
'    If hFtpConnection Then mFtpServer = sServer
'End Function
'Public Sub CloseFtpConnection()
'    If hFtpConnection <> 0 Then InternetCloseHandle hFtpConnection
'    hFtpConnection = 0
'    If hFtpOpen Then InternetCloseHandle hFtpOpen
'    hFtpOpen = 0
'End Sub
'Public Function GetFtpDirectory() As String
'If hFtpConnection = 0 Then Exit Function
'Dim szDir As String
'szDir = String$(1024, 0)
'If FtpGetCurrentDirectory(hFtpConnection, szDir, 1024) Then GetFtpDirectory = Left$(szDir, InStr(1, szDir, String$(1, 0), vbBinaryCompare) - 1)
'End Function
'Public Function SetFtpDirectory(ByVal sDir As String) As Boolean
'If hFtpConnection Then SetFtpDirectory = FtpSetCurrentDirectory(hFtpConnection, sDir) ' & Chr(0))
'End Function
'Public Sub GetFtpDirList(ByVal pth$, ByVal sFilter$, cp As CParam, Optional bWithFolders As Boolean = True)
''ROOTDIR =\
''DIR0 =\dir
''DIR1 =\dir\subdir
'If hFtpConnection = 0 Or cp Is Nothing Then Exit Sub
'Dim hFind As Long, i&, cpc As New CParam, s$, sz&
'Dim pData As WIN32_FIND_DATA
'cp.Clear
'If SetFtpDirectory(pth) Then
'    hFind = FtpFindFirstFile(hFtpConnection, sFilter & Chr(0), pData, &H400&, 0)
'    i = hFind
'    While i
'        s = Left$(pData.cFileName, InStr(1, pData.cFileName, String$(1, 0), vbBinaryCompare) - 1)
'        If Not (s = "." Or s = "..") Then
'            If (pData.dwFileAttributes And &H10) = 0 Then cp.Param(s) = FileTimeToDate(pData.ftLastWriteTime) Else If bWithFolders Then cp.Param(s) = -1
'            sz = sz + pData.nFileSizeLow
'        End If
'        pData.cFileName = String$(260, 0)
'        i = InternetFindNextFile(hFind, pData)
'    Wend
'    InternetCloseHandle hFind
'    For i = 0 To cp.ParamCount - 1
'        If cp.Param(i) = -1 Then 'DIR
'            s = cp.ParamName(i)
'            cpc.Value = cp.Param(i)
'            GetFtpDirList pth & IIf(Len(pth) = 1, "", "/") & s, sFilter, cpc
'            sz = sz + Val(cpc.Name)
'            cp.Param(i) = cpc.Source
'        End If
'    Next
'    cp.Name = "" & sz
'End If
'End Sub
'Public Function CreateFtpDirectory(sDirectory As String) As Boolean
'If hFtpConnection Then CreateFtpDirectory = FtpCreateDirectory(hFtpConnection, sDirectory)
'End Function
'Public Function RemoveFtpDirectory(sDirectory As String) As Boolean
'If hFtpConnection Then RemoveFtpDirectory = FtpRemoveDirectory(hFtpConnection, sDirectory)
'End Function
'Public Function RenameFtpFile(sExisting As String, sNewName As String) As Boolean
'If hFtpConnection Then RenameFtpFile = FtpRenameFile(hFtpConnection, sExisting, sNewName)
'End Function
'Public Function DeleteFtpFile(sRemote As String) As Boolean
'If hFtpConnection Then DeleteFtpFile = FtpDeleteFile(hFtpConnection, sRemote)
'End Function
'Public Function SimpleFtpPutFile(sLocal As String, sRemote As String) As Boolean
'If hFtpConnection Then SimpleFtpPutFile = FtpPutFile(hFtpConnection, sLocal, sRemote, 0, 0)
'End Function
'Public Function SimpleFtpGetFile(sLocal As String, sRemote As String) As Boolean
'If hFtpConnection Then SimpleFtpGetFile = FtpGetFile(hFtpConnection, sRemote, sLocal, False, &H80, &H80000000, 0)
'End Function
'Public Function LastResponseInfo() As String
'    Dim nErr As Long, BuffSize As Long
'    InternetGetLastResponseInfo nErr, LastResponseInfo, BuffSize
'    If BuffSize > 0 Then
'        LastResponseInfo = Space(BuffSize)
'        InternetGetLastResponseInfo nErr, LastResponseInfo, BuffSize
'    End If
'End Function
''==================== FTP ===============================
''=======================================================



''=======================================================
'Private Sub IBindStatusCallback_GetBindInfo(grfBINDF As Long, pbindinfo As Long)
''
'grfBINDF = 1
'End Sub
'Private Function IBindStatusCallback_GetPriority() As Long
''
'End Function
'Private Sub IBindStatusCallback_OnDataAvailable(ByVal grfBSCF As Long, ByVal dwSize As Long, pformatetc As XVBHelpLib.FORMATETC, pstgmed As XVBHelpLib.STGMEDIUM)
''
'End Sub
'Private Sub IBindStatusCallback_OnLowResource(ByVal reserved As Long)
''
'End Sub
'Private Sub IBindStatusCallback_OnObjectAvailable(riid As XVBHelpLib.UUID, ByVal punk As stdole.IUnknown)
''
'End Sub
'
'Private Sub IBindStatusCallback_OnProgress(ByVal ulProgress As Long, ByVal ulProgressMax As Long, ByVal ulStatusCode As Long, ByVal szStatusText As Long)
'    'the OnProgress event is what keeps your app from freezing during download ala URLDownloadToFile
'    '--> just an fyi... IBindStatusCallback_OnProgress returns results in terms of bytes...
'    '--> ulProgressMax = the total # bytes to be downloaded
'    '--> ulProgress = the # of bytes downloaded thus far
'    '--> ...you need to base your calculations (if any) in those terms
'
'    'if you downloaded the original version of this project you will remember that ulProgressMax can return null
'    'which really screws up a progress bar... so protect against that possibility here... and
'    'handle the ulProgress error possibilities in the VicDLProg event
''    If ulProgressMax > 0 Then
'        'RaiseEvent VicDLProg(ulProgress, ulProgressMax) 'report progress to the event
''    End If
'    'this is a good place to let a cancel request get in line for processing
'    'Debug.Print LastResponseInfo
''mBinding.Abort
''Debug.Print "START IBindStatusCallback_OnProgress "
'    If Not xStatus Is Nothing Then
'        'xStatus.ListIndex = ulProgressMax
'        'xStatus.Help = SysAllocString(szStatusText)
'        'xStatus.Validate ulProgress
'        xStatus.hEvent "Progress", ulProgress, ulProgressMax, BINDSTATUSText(ulStatusCode), StringWFromPointer(szStatusText)
'        'API_DoEvents
''        If xStatus.Visible Then xStatus.NCRepaint
'        'Debug.Print ulProgress & "/" & ulProgressMax & "  " & BINDSTATUSText(ulStatusCode) & " : "; SysAllocString(szStatusText) 'StringWFromPointer(szStatusText)
'    Else
'        'Debug.Print ulProgress & "/" & ulProgressMax & "  " & BINDSTATUSText(ulStatusCode) & " : "; SysAllocString(szStatusText) 'StringWFromPointer(szStatusText)
'        'Err.Raise E_ABORT
'    End If
''Debug.Print "END IBindStatusCallback_OnProgress "
'    'leave this sub here even if you don't utilize it!
'    'all events need to be exposed... and so on for all IBind subs below
'    'see: http://msdn.microsoft.com/library/default.asp?url=/workshop/networking/moniker/reference/ifaces/urlmon_ref_ifaces_entry.asp
'End Sub
'Private Sub IBindStatusCallback_OnStartBinding(ByVal dwReserved As Long, ByVal pib As XVBHelpLib.IBinding)
'    Set mBinding = pib 'get something to read from callback... and get in the "stream" of things - so to speak.
'    'RaiseEvent VicDLStart
'
'    '(binding has started.  See the frmDnLoad code for some reasons to use this info)
'    'leave this sub here even if you don't utilize it!
'End Sub
'Private Sub IBindStatusCallback_OnStopBinding(ByVal hresult As Long, ByVal szError As Long)
'    'Binding has stopped... figure out why if you want to know...
'    If hresult = 1 Or hresult = S_OK Then
'        'Download Completed normally...
'        'RaiseEvent VicDLDone
'
'    Else
'        'Download was aborted/or cancelled:
'        'You would think that szError would be the logical number to interpret...
'        'But, it's not!  hresult is the error number to interpret...
'        'if it's a failure... it's most likely a server/connection problem/or file not found
'        'RaiseEvent VicDLCrash(hresult, GetVicError(hresult), Cancelled) 'report the failure to the event for handling
'        xStatus.hEvent "Progress", 0, hresult, "ERROR", Hex$(Err.LastDllError), LastResponseInfo()
'    End If
'    Set mBinding = Nothing
'    'If Not xStatus Is Nothing Then Set xStatus = Nothing
'    'leave this sub here even if you don't utilize it!
'    'See the frmDnLoad code for some reasons to use this info
'End Sub
'
'Function BINDSTATUSText(ID)
'Static ar
'On Error Resume Next
'If (VarType(ar) And vbArray) = 0 Then
'ar = Split("CONNECTING,REDIRECTING,BEGINDOWNLOADDATA,DOWNLOADINGDATA,ENDDOWNLOADDATA,BEGINDOWNLOADCOMPONENTS,INSTALLINGCOMPONENTS,ENDDOWNLOADCOMPONENTS,USINGCACHEDCOPY,SENDINGREQUEST,CLASSIDAVAILABLE,MIMETYPEAVAILABLE,CACHEFILENAMEAVAILABLE,BEGINSYNCOPERATION,ENDSYNCOPERATION,BEGINUPLOADDATA,UPLOADINGDATA,ENDUPLOADINGDATA,PROTOCOLCLASSID,ENCODING,VERFIEDMIMETYPEAVAILABLE,CLASSINSTALLLOCATION,DECODING,LOADINGMIMEHANDLER,CONTENTDISPOSITIONATTACH,FILTERREPORTMIMETYPE,CLSIDCANINSTANTIATE,IUNKNOWNAVAILABLE,DIRECTBIND,RAWMIMETYPE,PROXYDETECTING,ACCEPTRANGES,COOKIE_SENT,COMPACT_POLICY_RECEIVED,COOKIE_SUPPRESSED,COOKIE_STATE_UNKNOWN,COOKIE_STATE_ACCEPT,COOKIE_STATE_REJECT,COOKIE_STATE_PROMPT,COOKIE_STATE_LEASH,COOKIE_STATE_DOWNGRADE,POLICY_HREF,P3P_HEADER,SESSION_COOKIE_RECEIVED,PERSISTENT_COOKIE_RECEIVED,SESSION_COOKIES_ALLOWED,CACHECONTROL,CONTENTDISPOSITIONFILENAME,MIMETEXTPLAINMISMATCH,PUBLISHERAVAILABLE,DISPLAYNAMEAVAILABLE,SSLUX_NAVBLOCKED,SERVER_MIMETYPEAVAILABLE,SNIFFED_CLASSIDAVAILABLE", ",")
'End If
'BINDSTATUSText = ar(ID)
'Err.Clear
'End Function
'
'
'
''=======================================================
'
'Function URLFileTime(ByVal URL$)
'Dim hInternet   As Long
'Dim hURL        As Long
'Dim ret         As Long, res As Date
'Dim pst As SYSTEMTIME
'Dim buf() As Byte, buflen&
''Const INTERNET_FLAG_ASYNC = &H10000000
'Const INTERNET_FLAG_ASYNC           As Long = &H10000000
'
'hInternet = InternetOpen("", 0, "", "", 0) 'INTERNET_FLAG_ASYNC)
''Debug.Print "InternetSetStatusCallback=" & InternetSetStatusCallback(hInternet, GetProcAddress(AddressOf InternetCallbackFunc)) 'InternetStatusCallback))
'If hInternet Then
'    hURL = InternetOpenUrl(hInternet, StrPtr(URL), 0, 0, &H84000100, 0) 'INTERNET_FLAG_RELOAD=&H80000000 Or INTERNET_FLAG_PRAGMA_NOCACHE=&H100& Or INTERNET_FLAG_NO_CACHE_WRITE=&H4000000
'    If hURL Then
'        buflen = 99
'        ReDim buf(buflen) As Byte
'        ret = HttpQueryInfo(hURL, 11, buf(0), buflen, 0)  'ByVal StrPtr(buf), Len(buf), 0)
'        InternetCloseHandle hURL
'        If ret Then
'            ReDim Preserve buf(buflen - 1)
'            URLFileTime = StrConv(buf, vbUnicode)
'            'URLFileTime = RTrimNull("" & URLFileTime)
'            InternetTimeToSystemTime URLFileTime, pst, 0
'            Dim d As Double
'            SystemTimeToVariantTime pst, d
'            'URLFileTime =
'            res = CDate(d)
'            'URLFileTime = DateSerial(pst.wYear, pst.wMonth, pst.wDay) + TimeSerial(pst.wHour, pst.wMinute, pst.wSecond)
'            'URLFileTime = URLFileTime & " ->>  " & UTCTime(URLFileTime)
'        End If
'    End If
'    InternetCloseHandle hInternet
'End If
'URLFileTime = res
''Debug.Print "URLFileTime=" & Nz(URLFileTime, "NULL")
'End Function
'
