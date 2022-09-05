VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ZLibVer Lib "zlib.dll" Alias "zlibVersion" () As Long

Private Const API_NULL As Long = 0

Private Enum Z_RESULTS
    Z_VERSION_ERROR = -6
    Z_BUF_ERROR = -5
    Z_MEM_ERROR = -4
    Z_DATA_ERROR = -3
    Z_STREAM_ERROR = -2
    Z_ERRNO = -1
    Z_OK = 0
    Z_STREAM_END = 1
    Z_NEED_DICT = 2
End Enum

Private Const ZIP_CHUNK As Long = 131072  '2 * 65536, could be smaller or larger than this.

Public Enum ZIP_RESULTS
    'In typical "open source" C/C++ programming fashion there is
    'much that goes poorly documented, undocumented, or left
    'open to interpretation due to the crudity of the C and to
    'a lesser extent C++ languages themselves.
    '
    'I am hazarding a guess that these are the result codes
    'for calls to the entrypoints in zlibwapi.dll that we
    'are using here.
    ZIP_INTERNALERROR = -104
    ZIP_BADZIPFILE = -103
    ZIP_PARAMERROR = -102
    ZIP_OK = 0
    ZIP_EOF = 0
    ZIP_ERRNO = Z_ERRNO
    'Extensions for callers checking the Result property:
    ZIP_NOT_OPEN = -666
    ZIP_ALREADY_OPEN = -667
End Enum

Public Enum APPEND_MODES
    APPEND_STATUS_CREATE = 0
    [_APPEND_STATUS_CREATEAFTER] = 1 'Hide this member, seldom used.
    APPEND_STATUS_ADDINZIP = 2
End Enum

Public Enum Z_METHODS
    Z_STORE = 0               'Store without compressing.
    Z_DEFLATED = 8
End Enum

Public Enum Z_LEVELS
    'Your results may vary.   See these comments:
    Z_DEFAULT_COMPRESSION = -1
                              'On many files gives same compression as
                              'Z_BEST_COMPRESSION but runs faster.
    Z_NO_COMPRESSION = 0
    Z_BEST_SPEED = 1          'On many files seems to take about the same time
                              'Z_DEFAULT_COMPRESSION but produces *better*
                              'compression than Z_BEST_COMPRESSION.
    Z_BEST_COMPRESSION = 9    'On many files slower than Z_DEFAULT_COMPRESSION
                              'but only gives same compression.
End Enum

'tm_unz contain date/time info
Private Type tm_zip
    tm_sec  As Long 'seconds after the minute - [0,59]
    tm_min  As Long 'minutes after the hour - [0,59]
    tm_hour  As Long 'hours since midnight - [0,23]
    tm_mday  As Long 'day of the month - [1,31]
    tm_mon  As Long 'months since January - [0,11]
    tm_year  As Long 'years - [1980..2044]
End Type

'unz_global_info structure contain global data about the ZIPfile  These data comes from the end of central dir
Private Type unz_global_info
    number_entry As Long 'total number of entries in the central dir on this disk
    size_comment As Long 'size of the global comment of the zipfile
End Type

'unz_file_info contain information about a file in the zipfile
Private Type unz_file_info
    Version As Long 'version made by                 2 bytes
    version_needed As Long 'version needed to extract       2 bytes
    flag As Long 'general purpose bit flag        2 bytes
    compression_method As Long 'compression method              2 bytes
    dosdate As Long 'last mod file date in Dos fmt   4 bytes
    crc As Long 'crc-32                          4 bytes
    compressed_size As Long 'compressed size                 4 bytes
    uncompressed_size As Long 'uncompressed size               4 bytes
    size_filename As Long 'filename length                 2 bytes
    size_file_extra As Long 'extra field length              2 bytes
    size_file_comment As Long 'file comment length             2 bytes
    disk_num_start As Long 'disk number start               2 bytes
    internal_fa As Long 'internal file attributes        2 bytes
    external_fa As Long 'external file attributes        4 bytes
    tmu_date As tm_zip
End Type





Private Type zip_fileinfo
    'More typical "open source" C programming rears its ugly head
    'here.  The comment below:
    '
    '       if dos_date == 0, tmu_date is used
    '
    '... probably ought to say something like:
    '
    '       if dosDate = 0, tmz_date is used
    '
    'Oh well, that's the assumption I'm relying on and it seems
    'to work.

    tmz_date As tm_zip        'date in understandable format
    dosdate As Long           'if dos_date == 0, tmu_date is used

    'Next line (flag) is commented in the original C header file.
    'flag as long             'general purpose bit flag        2 bytes
    internal_fa As Long       'internal file attributes        2 bytes

    'Next line (external_fa) I have substituted the intrinsic VB
    'Enum as the type.
    'external_fa As Long      'external file attributes        4 bytes
    external_fa As VbFileAttribute
End Type


Private Declare Function unzOpen Lib "zlib.dll" (ByVal sFilePath As String) As Long
Private Declare Function unzClose Lib "zlib.dll" (ByVal hFile As Long) As Long
Private Declare Function unzGetGlobalInfo Lib "zlib.dll" (ByVal hFile As Long, ByRef pglobal_info As unz_global_info) As Long
Private Declare Function unzGetCurrentFileInfo Lib "zlib.dll" (ByVal hFile As Long, _
                         ByRef pfile_info As unz_file_info, _
                         ByVal szFileName As String, _
                         ByVal fileNameBufferSize As Long, _
                         ByRef extraField As Long, _
                         ByVal extraFieldBufferSize As Long, _
                         ByVal szComment As String, _
                         ByVal commentBufferSize As Long) As Long
Private Declare Function unzGoToNextFile Lib "zlib.dll" (ByVal hFile As Long) As Long
Private Declare Function unzLocateFile Lib "zlib.dll" (ByVal hFile As Long, ByVal szFileName As String, ByVal iCaseSensitivity As Long) As Long

Private Declare Function unzOpenCurrentFile Lib "zlib.dll" (ByVal hFile As Long) As Long
Private Declare Function unzReadCurrentFile Lib "zlib.dll" (ByVal hFile As Long, ByRef Buffer As Byte, ByVal BuffLen As Long) As Long
Private Declare Function unzCloseCurrentFile Lib "zlib.dll" (ByVal hFile As Long) As Long



'Private Declare Function zipOpen Lib "zlib.dll" (ByVal ppathname As Long, ByVal append As APPEND_MODES) As Long
Private Declare Function zipOpen Lib "zlib.dll" (ByVal sFilePath As String, ByVal append As APPEND_MODES) As Long
Private Declare Function zipOpenNewFileInZip Lib "zlib.dll" ( _
    ByVal hFile As Long, _
    ByVal pfilename As Long, _
    ByRef zipfi As zip_fileinfo, _
    ByVal pextrafield_local As Long, _
    ByVal size_extrafield_local As Long, _
    ByVal pextrafield_global As Long, _
    ByVal size_extrafield_global As Long, _
    ByVal pcomment As Long, _
    ByVal method As Z_METHODS, _
    ByVal Level As Z_LEVELS) As ZIP_RESULTS

Private Declare Function zipWriteInFileInZip Lib "zlib.dll" (ByVal hFile As Long, ByVal pbuf As Long, ByVal length As Long) As ZIP_RESULTS
Private Declare Function zipCloseFileInZip Lib "zlib.dll" (ByVal hFile As Long) As ZIP_RESULTS
Private Declare Function zipClose Lib "zlib.dll" (ByVal hFile As Long, ByVal pglobal_comment As Long) As ZIP_RESULTS

Private hZip As Long 'Zip file for write
Private hUnZip As Long 'Zip file for read
Private cp As New CParam 'Zip contents
Private m_Name As String 'Zip filename
Private m_path As String 'Zip filepath

Private Const BUFFERSIZE        As Long = 4096 '2048
Private Const MAX_PATH          As Long = 1024 '260
Private Const MAX_COMMENT       As Long = 1024 '255
Private szComment As String * MAX_COMMENT
Private strFileNameBuff As String * MAX_PATH
Private buff()                  As Byte
Private mCancel As Boolean
Private nfiles&, nfile&

'Public Event ZipFiles(sFile$, nTotal&, nCurrent&)
'Public Event ZipBytes(sFile$, nTotal&, nCurrent&)
Public Event UnZipFiles(sFile$, nTotal&, nCurrent&)
Public Event UnZipBytes(sFile$, nTotal&, nCurrent&)
Public Event Cancel()

Public Sub Cancel()
mCancel = True
End Sub

Public Function ListZip()
ListZip = cp.Source
End Function

Public Property Get Name() As String
Name = m_Name
End Property
Public Property Get Path() As String
Path = m_path
End Property

Private Function pListZip()
If hUnZip = 0 Then Exit Function
Dim strFilename As String
Dim i As Long
Dim info As unz_file_info
mCancel = False
Dim ar
ar = Array(Null, "icon", 0, "size", 0, "time", 0) ', "ziptime", 0)
ReDim Preserve ar(-1 To 5)
cp.Clear
Do
    unzGetCurrentFileInfo hUnZip, info, strFileNameBuff, MAX_PATH, 0, 0, szComment, MAX_COMMENT
    strFilename = Left(strFileNameBuff, info.size_filename)
    strFilename = Replace(zip2name(strFilename), "/", "\")
    ar(-1) = zipGetFileName(strFilename)
    'ar(1) = CLng(SysImageListIndex(IIf(info.external_fa = vbDirectory, "", ar(-1))))
    ar(1) = CLng(SysImageListIndex(IIf(info.external_fa = vbDirectory, "", zipGetFileName(strFilename))))
    ar(3) = info.uncompressed_size
    With info.tmu_date
        ar(5) = DateSerial(.tm_year, .tm_mon + 1, .tm_mday) + TimeSerial(.tm_hour, .tm_min, .tm_sec)
    End With
    cp(strFilename) = ar
    If InStr(strFilename, "\") Then cp(zipGetPath(strFilename) & "icon") = CLng(SysImageListIndex(vbNullString))

'    Debug.Print strFilename, ar(1), ar(3)
Loop Until (unzGoToNextFile(hUnZip) <> 0) And Not mCancel
If mCancel Then cp.Clear
'ListZip = cp.Source
If mCancel Then RaiseEvent Cancel
End Function

Private Function zip2name(ByVal sText As String) As String
Dim lngOutLen&
lngOutLen = MultiByteToWideCharStr(GetOEMCP, 0, sText, Len(sText), 0, 0)
zip2name = String$(lngOutLen, 0)
zip2name = Left(zip2name, MultiByteToWideCharStr(GetOEMCP, 0, sText, Len(sText), StrPtr(zip2name), lngOutLen))

'Dim nRet As Long, strRet As String
'strRet = String$(Len(sText), vbNullString)
'nRet = MultiByteToWideChar1(GetOEMCP, 0, sText, Len(sText), StrPtr(strRet), Len(strRet))
'zip2name = Left$(strRet, nRet)
End Function

Private Function name2zip(sText As String) As String
Dim nRet As Long, strRet As String
strRet = String$(Len(sText), vbNullChar)
nRet = WideCharToMultiByte(GetOEMCP, &H0, StrPtr(sText), Len(sText), StrPtr(strRet), Len(sText), 0&, 0&)
name2zip = Left$(StrConv(strRet, vbUnicode), nRet)
End Function

Function OpenZip(ByVal strZipFilePath$) 'open for read/write + create if no exist
cp.Clear
If ZLibVer = 0 Then Exit Function
CloseZip
Dim n&
On Error Resume Next
n = FileLen(strZipFilePath)
If n Then hUnZip = unzOpen(strZipFilePath): pListZip
If n = 0 Then MakeDir strZipFilePath
If hUnZip = 0 Then hZip = zipOpen(strZipFilePath, IIf(n = 0, APPEND_STATUS_CREATE, APPEND_STATUS_ADDINZIP))
If hZip Or hUnZip Then
    m_Name = zipGetFileName(strZipFilePath)
    m_path = zipGetPath(strZipFilePath)
End If
End Function

Function CloseZip() 'Close and Save current zip
If hUnZip Then unzClose hUnZip: hUnZip = 0
If hZip Then zipClose hZip, 0: hZip = 0
m_Name = vbNullString: m_path = vbNullString
End Function

Private Function pFindZipByName(ByVal sfnd$, ByVal pth$)
Dim ar, i&, s$
ar = cp(pth)
For i = 0 To UBound(ar) Step 2
If IsArray(ar(i + 1)) Then
    s = pth & IIf(Len(pth) = 0, vbNullString, "\") & ar(i)
    If ar(i) = sfnd Then pFindZipByName = s: Exit For
    pFindZipByName = pFindZipByName(sfnd, s)
    If Len(pFindZipByName) Then Exit For
End If
Next
End Function

Function ExtractFiles(ByVal zipFilePath$, ByVal remoteFilePath$)
If hUnZip = 0 Then Exit Function
mCancel = False
Dim ret&, cd$

Dim sp As New CParam
sp.Value = cp(zipFilePath) 'Dir/File Content
If sp.ParamCount = 0 Then
    zipFilePath = pFindZipByName(zipFilePath, vbNullString)
    If Len(zipFilePath) Then sp.Value = cp(zipFilePath)
End If
cd = zipGetFileName(zipFilePath)
Dim ar
ReDim ar(-1 To 1)
ar(1) = sp.Source
Select Case True 'sp.ParamCount
Case VarType(sp.Name) < 2 'Это каталог надо распаковать
    remoteFilePath = NzS(remoteFilePath, m_path & xMain.SplitIndex(m_Name, ".", 0) & "\")
    MakeDir remoteFilePath
    nfiles = cpCount(sp.Source)
    If Len(cd) Then
        ar(0) = cd
        'pExtractFiles ar, zipGetPath(zipFilePath), zipGetPath(remoteFilePath)
        pExtractFiles ar, zipGetPath(zipFilePath), remoteFilePath
    Else
        pExtractFiles sp.Source, zipFilePath, zipGetPath(remoteFilePath)
    End If
Case VarType(sp.Name) = vbString 'Одиночный файл надо распаковать по указанным именем
    If Len(remoteFilePath) = 0 Then Exit Function
    cd = MakeDir(remoteFilePath)
    nfiles = 1
    ar(0) = sp.Name
    pExtractFiles ar, zipGetPath(zipFilePath), remoteFilePath
End Select
If mCancel Then RaiseEvent Cancel
End Function

Function ReadBytes(ByVal zipFilePath$) As String
If hUnZip = 0 Then Exit Function

Dim ar, cfile$, n&, ret&
ar = cp(zipFilePath)
If UBound(ar) <> 5 Then Exit Function
n = ar(3)
cfile = Replace(zipFilePath, "\", "/")
cfile = name2zip(cfile)
ret = unzLocateFile(hUnZip, cfile, 2)
If ret = 0 Then 'ReadBytes to buffer
ret = unzOpenCurrentFile(hUnZip)
ReDim buff(n - 1) As Byte
ret = unzReadCurrentFile(hUnZip, buff(0), n)
ret = unzCloseCurrentFile(hUnZip)
ReadBytes = buff 'StrConv(buff, vbUnicode)
ReDim buff(BUFFERSIZE - 1) As Byte
End If
End Function

Private Function zipLocate(ByVal sFile$) As Long
'Dim zfile$
'zfile = name2zip(file)
zipLocate = unzLocateFile(hUnZip, name2zip(sFile), 0)
If zipLocate = 0 Then Exit Function
'zfile = name2zip(Replace(file, "\", "/"))
zipLocate = unzLocateFile(hUnZip, name2zip(Replace(sFile, "\", "/")), 0)

End Function

Private Sub pExtractFiles(ar, ByVal zip_pth$, ByVal dst_pth$)
On Error GoTo err_sub
Dim ret&, i&, cfile$, zfile$
Dim ff#, nb&, nbs&
Dim evs$, evs0$
ReDim buff(BUFFERSIZE - 1) As Byte
For i = 0 To UBound(ar) Step 2
    If IsArray(ar(i + 1)) Then
        If UBound(ar(i + 1)) = 5 And VarType(ar(i + 1)(-1)) > 1 Then   'Extract File
        
            ret = zipLocate(zip_pth & ar(i))
            If ret = 0 Then 'WriteFile to Disk
                nfile = nfile + 1: nbs = ar(i + 1)(3)
                cfile = dst_pth
                If Len(zipGetFileName(dst_pth)) = 0 Then cfile = dst_pth & ar(i)

'Dim info As unz_file_info
'ret = unzGetCurrentFileInfo(hUnZip, info, vbNullString, 0, 0, 0, vbNullString, 0)

                ret = unzOpenCurrentFile(hUnZip)
                If ret Then Exit For                        'ERROR
                ff = FreeFile
                
                Open cfile For Binary As ff
                Do
                    ret = unzReadCurrentFile(hUnZip, buff(0), BUFFERSIZE)
                    If ret = 0 Then Exit Do
                    If ret < BUFFERSIZE Then
                        ReDim tbuff(ret - 1) As Byte
                        CopyMemory tbuff(0), buff(0), ret
                        Put ff, , tbuff
                    Else
                        Put ff, , buff
                    End If
                    DoEvents
                    nb = nb + ret
                    evs0 = Partition(nb, 0, nbs, ZIP_CHUNK)
                    If evs <> evs0 Then
                        RaiseEvent UnZipBytes(cfile, nbs, nb)
                        evs = evs0
                    End If
                Loop Until (ret < BUFFERSIZE) And mCancel = False
                RaiseEvent UnZipBytes(cfile, nbs, nbs)
                ret = unzCloseCurrentFile(hUnZip)
                Close ff
                xMain.FileTimes(cfile, 7) = ar(i + 1)(5)
                If mCancel Then Kill cfile: Exit For
            Else
                xMain.DebugPrint 240, "ERROR Extract " & zip_pth & ar(i)
            End If
        Else 'Extract Catalog
            cfile = dst_pth & ar(i)
            MakeDir cfile & "\"
            nfile = nfile + 1
            'pExtractFiles ar(i + 1), zip_pth & IIf(Len(zip_pth), "\", vbNullString) & ar(i) & "\", cfile & "\"
            pExtractFiles ar(i + 1), zip_pth & ar(i) & "\", cfile & "\"
            RaiseEvent UnZipFiles(cfile, nfiles, nfile)
            If mCancel Then Exit For

        End If
    End If
    If mCancel Then Exit For
Next
Exit Sub
err_sub:
Debug.Print Err.Description
Debug.Assert False
Resume 'Next
End Sub

Function cpCount(ar) As Long 'Return count CP nodes not cp params
Dim i&
If IsArray(ar) Then
For i = 0 To UBound(ar) Step 2
    If IsArray(ar(i + 1)) Then
        cpCount = cpCount + 1
        cpCount = cpCount + cpCount(ar(i + 1))
    End If
Next
End If
End Function

Function zipGetFileName(ByVal strPath As String) As String
Dim i As Integer
        i = InStrRev(strPath, "\")
        zipGetFileName = Right(strPath, Len(strPath) - i)
End Function
'
Function zipGetPath(ByVal strPath As String) As String
Dim i As Integer
        i = InStrRev(strPath, "\")
        zipGetPath = Left(strPath, i)
End Function

Function AddFile(ByVal zipFilePath$, ByVal remoteFilePath$)
If hZip = 0 Then Exit Function
If Len(zipFilePath) = 0 Then Exit Function
If Len(remoteFilePath) = 0 Then Exit Function


'Debug.Print zipFilePath, remoteFilePath
'Exit Function
Dim sFileName$, mResult&, dFileDate As Date
Dim zfi As zip_fileinfo, fnum#, n&
sFileName = StrConv(Replace(zipFilePath, "\", "/"), vbFromUnicode)
With zfi
    With .tmz_date
        dFileDate = xMain.FileTimes(remoteFilePath, 4)
        .tm_sec = Second(dFileDate)
        .tm_min = Minute(dFileDate)
        .tm_hour = Hour(dFileDate)
        .tm_mday = Day(dFileDate)
        .tm_mon = Month(dFileDate) - 1 'Odd but true.
        .tm_year = Year(dFileDate)
    End With
    'Dim ft As FILETIME, lw#, hw#
    'ft = DateToFileTime(dFileDate)
    'FileTimeToDosDateTime ft, lw, hw
    '.dosdate = MakeDWord(lw, hw)
    '.external_fa = vbNormal
    '.internal_fa = 0
End With
mResult = zipOpenNewFileInZip(hZip, StrPtr(sFileName), zfi, 0, 0, 0, 0, 0, Z_DEFLATED, Z_DEFAULT_COMPRESSION)
fnum = FreeFile()
Open remoteFilePath For Binary Access Read As #fnum
n = FileLen(remoteFilePath)
Do
    If n > 0 Then
        If n < ZIP_CHUNK Then ReDim buff(n - 1) Else ReDim buff(ZIP_CHUNK - 1)
        n = n - ZIP_CHUNK
        Get #fnum, , buff
        mResult = zipWriteInFileInZip(hZip, VarPtr(buff(0)), UBound(buff) + 1)
    Else
        Close #fnum
        mResult = zipCloseFileInZip(hZip)
        Exit Do
    End If
Loop
End Function


'Function AddFiles(ByVal zipFilePath$, ByVal remoteFilePath$)
'If hZip = 0 Then Exit Function
'If Len(remoteFilePath) = 0 Then Exit Function
'Dim i&, cfile$, cpath$
'Dim sp As New CParam
'sp.Value = ListDir(Nothing, remoteFilePath) 'ListDir return SingleFile or Directory List with SubDirectories and Files
'
''    If cp.PathParamCount(zipFilePath) Then 'ReCompress all files in zip
''    Else 'AddNew File
''    End If
'
'
'ListZip
'End Function

'Private Function pAddFiles(sp)
'
'
'End Function

'Private Function pAddFile(ByVal zipFilePath$, ByVal remoteStartDir$, ByVal remoteFile$) As Boolean
'If IsArray(cp(zipFilePath)) Then Exit Function
'Dim ar, n&
'On Error Resume Next
'n = FileLen(remoteFile)
'If n = 0 Then Exit Function
'ar = Array(Null, "icon", 0, "size", 0, "time", 0)
'ReDim Preserve ar(-1 To 5)
'ar(-1) = zipGetFileName(zipGetFileName(remoteFile))
''ar(1) = SysImageListIndex(ar(-1), 0)
'ar(1) = SysImageListIndex(ar(-1))
'ar(3) = n
'ar(5) = FileDateTime(remoteStartDir & remoteFile)
'
'
''Write to End of ZIP
'
'cp(zipFilePath) = ar
'pAddFile = True
'End Function

'Function ListDir(ByVal fso As FileSystemObject, ByVal sFilesPath$, Optional ext As String = "*.*")
'If Len(sFilesPath) = 0 Then Exit Function
''If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
'If fso Is Nothing Then Set fso = New FileSystemObject
'On Error Resume Next
'Dim ar, i&, n&
'Dim fold As Folder 'Object
'Dim fold0 As Folder 'Object
'Dim f As File 'Object
'Set fold = fso.GetFolder(sFilesPath)
'If fold Is Nothing Then
'    Set f = fso.GetFile(sFilesPath)
'    If f Is Nothing Then Exit Function
'    ar = Array(f.Name, "icon", 0, "size", 0, "time", 0)
'    'ar(2) = SysImageListIndex(file.Name, 0)
'    ar(2) = SysImageListIndex(f.Name)
'    ar(4) = f.SIZE
'    ar(6) = f.DateLastModified
'    ReDim Preserve ar(-1 To 5)
'Else
'    ar = Array(fold.Name, "icon", 0, "size", 0, "time", 0)
'    'ar(2) = SysImageListIndex(vbNullString, -1)
'    ar(2) = SysImageListIndex(vbNullString)
'    ar(4) = 0
'    ar(6) = fold.DateLastModified
'    ReDim Preserve ar(-1 To 5)
'    n = fold.SubFolders.count + fold.Files.count
'    If n Then
'        ReDim Preserve ar(-1 To 5 + n * 2)
'        i = 6
'        For Each fold0 In fold.SubFolders
'            ar(i) = fold0.Name
'            ar(i + 1) = ListDir(fso, sFilesPath & "\" & ar(i), ext)
'            i = i + 2
'        Next
'    End If
'    If fold.Files.count Then
'        For Each f In fold.Files
'            ar(i) = f.Name
'            ar(i + 1) = ListDir(fso, sFilesPath & "\" & ar(i), ext)
'            i = i + 2
'        Next
'    End If
'End If
'ListDir = ar
'End Function

'Function WriteBytes(ByVal zipFilePath$, buf) 'OverWrite if exist
'
'End Function


Private Sub Class_Terminate()
CloseZip
End Sub
