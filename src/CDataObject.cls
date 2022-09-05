VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Implements IDataObject
'Dim m_IDORTF As String ' RTF text buffer used by the TextRTF property
'Dim m_Image As StdPicture ' Image object used by  the InsertImage function
'Private Type METAFILEPICT
'   mm As Long
'   xExt As Long
'   yExt As Long
'   hMF As Long
'End Type
'Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long


Private Declare Function SendMessageRichEditOleCallback Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As IRichEditOleCallback) As Long
Implements IRichEditOleCallback
Private rich_hwnd&


Implements IDataObject
Private sRTF$
Private hGlobal As Long
Private mValue
Private mName$


Private WithEvents tm As CTimer
Attribute tm.VB_VarHelpID = -1
Private WithEvents whr As WinHttpRequest
Attribute whr.VB_VarHelpID = -1
Public xHttpListener&
Private pos&
'Event OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
'Event OnStart(ByVal Status As Long, ByVal ContentType As String)
'Event OnProgress(ByVal nPosition As Long)
'Event OnFinish(ByVal nTotal As Long)




Public Property Let TextRTF(v$)
sRTF = StrConv(v, vbFromUnicode)
'sRTF = v
'If Len(v) = 0 And hGlobal Then GlobalFree hGlobal: hGlobal = 0
End Property
Public Property Get TextRTF$()
TextRTF = StrConv(sRTF, vbUnicode)
'TextRTF = sRTF
End Property

Public Property Get Name$()
Name = mName
End Property
Public Property Let Name(v$)
mName = v
End Property


'!!!! DEFAULT MEMBER !!!!!
Public Property Let Value(v)
Attribute Value.VB_UserMemId = 0
If TypeName(mValue) = "WinHttpRequest" Then Set whr = Nothing
If TypeName(v) = "WinHttpRequest" Then
    Set whr = v
    pos = 0
Else
    If IsObject(v) Then Set mValue = v Else mValue = v
End If
End Property
'!!!! DEFAULT MEMBER !!!!!
Public Property Get Value()
If IsObject(mValue) Then
    Dim uu As UUID
    Select Case TypeName(mValue)
    Case "Collection": Set Value = CreateEventSinkObj(uu, mValue, Nothing)
    Case Else
        Set Value = mValue
    End Select
Else
    Value = mValue
End If
'Value = mValue
'Debug.Print "CDataObject Get Value =" & Nz(mValue, "NULL")
End Function
'!!!! DEFAULT MEMBER !!!!!

Public Property Get Request() As WinHttpRequest
Set Request = whr
End Property


Private Sub Class_Terminate()
DetachCRichOle
On Error Resume Next
tm.Interval = 0
whr.Abort
Set whr = Nothing
Err.Clear
End Sub


Private Sub WHREvent(ev$, arg) 'ParamArray arg())
If xHttpListener = 0 Then Exit Sub
On Error Resume Next
hxControl(xHttpListener).hEvent "+" & mName & "_" & IIf(L_(mValue) < 0, ev, "Event"), L_(mValue), whr, ev, arg
'Dim e As New CEventInfo, a(), p() As Long
'a = arg
'ReDim p(UBound(a))
'e.frInitialize ev, eid, a, Me, mName, "", p, ""
'hxControl(xHttpListener).hEvent "+" & mName & "_" & IIf(aVal(mValue), ev, "Event"), e
End Sub

Private Sub tm_Timer()
Static pc&
On Error Resume Next
If pc = pos Then Exit Sub
pc = pos
WHREvent "Progress", pos
'RaiseEvent OnProgress(pos)
End Sub

Private Sub whr_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
'On Error Resume Next
If Not tm Is Nothing Then tm.Interval = 0
WHREvent "Error", ErrorNumber & ":" & ErrorDescription
'RaiseEvent OnError(ErrorNumber, ErrorDescription)
'Err.Clear
End Sub
Private Sub whr_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
Set tm = New CTimer
tm.Interval = 50: pos = 0
WHREvent "Start", Status ', ContentType
'RaiseEvent OnStart(Status, ContentType)
End Sub
Private Sub whr_OnResponseDataAvailable(Data() As Byte)
pos = pos + UBound(Data) + 1
End Sub
Private Sub whr_OnResponseFinished()
On Error Resume Next: tm.Interval = 0
'If Not tm Is Nothing Then tm.Interval = 0
WHREvent "Finish", pos
'RaiseEvent OnFinish(pos)
End Sub

'==================IDataObject==================
Private Sub IDataObject_DAdvise(pformatetc As FORMATETC, ByVal advf As Long, ByVal pAdvSink As Long, pdwConnection As Long)
   Err.Raise E_NOTIMPL
End Sub
Private Sub IDataObject_DUnadvise(ByVal dwConnection As Long)
   Err.Raise E_NOTIMPL
End Sub
Private Sub IDataObject_EnumDAdvise(pIEnumAdvise As Long)
   Err.Raise E_NOTIMPL
End Sub
Private Function IDataObject_EnumFormatEtc(ByVal dwDirection As Long) As IEnumFORMATETC
   Err.Raise E_NOTIMPL
End Function
Private Sub IDataObject_GetCanonicalFormatEtc(pformatectIn As FORMATETC, pformatetcOut As FORMATETC)
   Err.Raise E_NOTIMPL
End Sub
Private Sub IDataObject_GetData(pformatetcIn As FORMATETC, pmedium As STGMEDIUM)
' Only paste text in RTF format
If pformatetcIn.cfFormat <> CF_RTF Then Err.Raise E_NOTIMPL
If hGlobal Then GlobalFree ByVal hGlobal: hGlobal = 0
hGlobal = GlobalAlloc(GPTR, LenB(sRTF))  ' Allocate global memory
CopyMemory ByVal hGlobal, ByVal StrPtr(sRTF), LenB(sRTF)  ' Copy the RTF to the global memory
With pmedium ' Set the STGMEDIUM structure
   .Data = hGlobal
   .TYMED = TYMED_HGLOBAL
   .pUnkForRelease = 0 'ObjPtr(Me)
End With
End Sub
Private Sub IDataObject_GetDataHere(pformatetc As FORMATETC, pmedium As STGMEDIUM)
   Err.Raise E_NOTIMPL
End Sub
Private Sub IDataObject_QueryGetData(pformatetc As FORMATETC)
   If pformatetc.cfFormat <> CF_RTF Then Err.Raise E_NOTIMPL
End Sub
Private Sub IDataObject_SetData(pformatetc As FORMATETC, pmedium As STGMEDIUM, ByVal fRelease As Long)
Err.Raise E_NOTIMPL
End Sub
'==================IDataObject==================




'Public Sub InsertImage(ByVal Image As StdPicture) ' Replaces the range with a bitmap or metafile
'Dim IDO As IDataObject
'Dim td As ITextDocument
'Select Case Image.Type
'    Case vbPicTypeBitmap, vbPicTypeMetafile
'        Set td = ITextDocument(m_hWnd)
'        Set m_Image = Image ' Store the image in the global variable
'        Set IDO = Me ' Get a reference to the IDataObject_Implements interface
'        td.Selection.Paste IDO, 0  ' Paste the image
'        'm_Range.Paste IDO, 0 ' Paste the image
'        Set m_Image = Nothing ' Release the image
'    Case Else
'       Err.Raise 481 ' Invalid Picture
'End Select
'End Sub
'Private Sub IDataObject_DAdvise(pformatetc As FORMATETC, ByVal advf As Long, ByVal pAdvSink As Long, pdwConnection As Long)
'   Err.Raise E_NOTIMPL
'End Sub
'Private Sub IDataObject_DUnadvise(ByVal dwConnection As Long)
'   Err.Raise E_NOTIMPL
'End Sub
'Private Sub IDataObject_EnumDAdvise(pIEnumAdvise As Long)
'   Err.Raise E_NOTIMPL
'End Sub
'Private Function IDataObject_EnumFormatEtc(ByVal dwDirection As Long) As IEnumFORMATETC
'   Err.Raise E_NOTIMPL
'End Function
'Private Sub IDataObject_GetCanonicalFormatEtc(pformatectIn As FORMATETC, pformatetcOut As FORMATETC)
'   Err.Raise E_NOTIMPL
'End Sub
'Private Sub IDataObject_GetData(pformatetcIn As FORMATETC, pmedium As STGMEDIUM)
'Dim hGlobal As Long
'If m_Image Is Nothing Then ' Check if it's an image what is being pasted
'   ' Only paste text in RTF format
'   If pformatetcIn.cfFormat <> CF_RTF Then Err.Raise E_NOTIMPL
'   hGlobal = GlobalAlloc(GPTR, LenB(m_IDORTF)) ' Allocate global memory
'   CopyMemory ByVal hGlobal, ByVal StrPtr(m_IDORTF), LenB(m_IDORTF) ' Copy the RTF to the global memory
'   With pmedium ' Set the STGMEDIUM structure
'      .Data = hGlobal
'      .TYMED = TYMED_HGLOBAL
'      .pUnkForRelease = 0
'   End With
'Else
'   With pmedium
'      If m_Image.Type = vbPicTypeBitmap Then
'         .TYMED = TYMED_GDI
'         .Data = m_Image.Handle
'         .pUnkForRelease = 0
'      ElseIf m_Image.Type = vbPicTypeMetafile Then
'         Dim MFP As METAFILEPICT, lPtr As Long
'         With MFP ' Fill the METAFILEPICT structure
'            .mm = 8 'MM_ANISOTROPIC 'Public Const MM_ANISOTROPIC = 8
'            .xExt = m_Image.Width
'            .yExt = m_Image.Height
'            .hMF = m_Image.Handle
'         End With
'         .Data = GlobalAlloc(GHND, LenB(MFP)) ' Allocate global memory for the structure
'         lPtr = GlobalLock(.Data) ' Get a pointer to the memory
'         CopyMemory ByVal lPtr, MFP, LenB(MFP) ' Copy the structure
'         GlobalUnlock .Data ' Release the pointer
'         .TYMED = TYMED_MFPICT
'      End If
'   End With
'End If
'End Sub
'Private Sub IDataObject_GetDataHere(pformatetc As FORMATETC, pmedium As STGMEDIUM)
'   Err.Raise E_NOTIMPL
'End Sub
'Private Sub IDataObject_QueryGetData(pformatetc As FORMATETC)
'If m_Image Is Nothing Then
'   ' Only allow RTF format
'   If pformatetc.cfFormat <> CF_RTF Then Err.Raise E_FAIL
'Else
'   Select Case m_Image.Type
'      Case vbPicTypeMetafile
'         If pformatetc.cfFormat <> vbCFMetafile Then Err.Raise E_FAIL
'      Case vbPicTypeBitmap
'         If pformatetc.cfFormat <> vbCFBitmap Then Err.Raise E_FAIL
'      Case Else
'         Err.Raise E_FAIL
'   End Select
'End If
'End Sub
'Private Sub IDataObject_SetData(pformatetc As FORMATETC, pmedium As STGMEDIUM, ByVal fRelease As Long)
'   Err.Raise E_NOTIMPL
'End Sub

'Sub AsyncGetRows(inMatrix As CMatrix, inAR, inRecordset As Recordset, inSARFields)
'ptrAR = VarPtr(inAR)
''ArrayAssign inAR, pAR
''ArrayAssign inSARFields, mSARFields = inSARFields
'mSARFields = inSARFields
'Set mRecordset = inRecordset
'Set mMatrix = inMatrix
'If tm Is Nothing Then Set tm = New CTimer
'tm.Interval = 1
'tm_Timer
'End Sub

'Private Sub tm_Timer()
'Dim ar
'If Not mRecordset.EOF Then
'    On Error Resume Next
'    Debug.Print mRecordset.AbsolutePosition
'    ar = mRecordset.GetRows(matrixMaxFetchRows, , mSARFields)
'
'    CopyMemory ByVal ptrAR, ar, 16 ' Œœ»–”≈Ã “”ƒ¿
'    FillMemory ar, 16, 0 '”ƒ¿Àﬂ≈Ã «ƒ≈—‹
'End If
'If mRecordset.EOF Then
'    tm.Interval = 0
'    Set tm = Nothing
'    mMatrix.INITARAsync 1
'    Set mRecordset = Nothing
'    Set mMatrix = Nothing
'Else
'    mMatrix.INITARAsync 0
'End If
'End Sub



'=====================IRichEditOleCallback===================================
Friend Sub AttachCRichOle(ByVal hWndRich&, Optional ByVal gos As Boolean)
If hWndRich = 0 Then Exit Sub
rich_hwnd = hWndRich
'pCreateStorage
If gos Then If g_oStorage Is Nothing Then Set g_oStorage = StgCreateDocfile("xvb." & GetCurrentProcessId & ".stg", STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE Or STGM_DELETEONRELEASE, 0)
SendMessageRichEditOleCallback rich_hwnd, EM_SETOLECALLBACK, 0, Me
End Sub
Friend Sub DetachCRichOle()
If rich_hwnd Then SendMessageRichEditOleCallback rich_hwnd, EM_SETOLECALLBACK, 0, Nothing
rich_hwnd = 0
End Sub

Private Sub IRichEditOleCallback_ContextSensitiveHelp(ByVal fEnterMode As Long)
End Sub
Private Sub IRichEditOleCallback_DeleteObject(ByVal lpoleobj As IOleObject)
End Sub
Private Sub IRichEditOleCallback_GetClipboardData(lpchrg As CHARRANGE, ByVal reco As Long, lplpdataobj As IDataObject)
Err.Raise E_NOTIMPL
End Sub
Private Sub IRichEditOleCallback_GetContextMenu(ByVal SelType As Integer, ByVal lpoleobj As IOleObject, lpchrg As CHARRANGE, lphmenu As Long)
PostMessage rich_hwnd, WM_CONTEXTMENU, ByVal lpchrg.cpMin, ByVal lpchrg.cpMax
End Sub
Private Sub IRichEditOleCallback_GetDragDropEffect(ByVal fDrag As Long, ByVal grfKeyState As Long, pdwEffect As Long)
Dim bLocked As Boolean
bLocked = hxControl(rich_hwnd).Locked
'pdwEffect = DROPEFFECT_MOVE 'œÓ ÛÏÓÎ˜‡ÌË˛ ÔÂÂÏÂ˘ÂÌËÂ
If bLocked Then
    If fDrag Then pdwEffect = DROPEFFECT_COPY Else pdwEffect = DROPEFFECT_NONE
Else
    If grfKeyState > 1 Then pdwEffect = DROPEFFECT_COPY '— ÍÌÓÔÍÓÈ ÍÓÔËÓ‚‡ÌËÂ
End If
End Sub
Private Sub IRichEditOleCallback_GetInPlaceContext(lplpFrame As IOleInPlaceFrame, lplpDoc As IOleInPlaceUIWindow, lpFrameInfo As OLEINPLACEFRAMEINFO)
Err.Raise E_NOTIMPL
End Sub
Private Sub IRichEditOleCallback_GetNewStorage(lplpstg As IStorage)
If g_oStorage Is Nothing Then Exit Sub
Set lplpstg = g_oStorage.CreateStorage(GetCurrentProcessId & "." & rich_hwnd, STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE, 0, 0)
End Sub
Private Sub IRichEditOleCallback_QueryAcceptData(ByVal lpdataobj As IDataObject, lpcfFormat As Integer, ByVal reco As Long, ByVal fReally As Long, ByVal hMetaPict As Long)
If hxControl(rich_hwnd).Locked And reco = RECO_DROP Then Err.Raise E_FAIL
End Sub
Private Sub IRichEditOleCallback_QueryInsertObject(lpclsid As UUID, ByVal lpstg As IStorage, ByVal cp As Long)
'Err.Raise E_NOTIMPL
End Sub
Private Sub IRichEditOleCallback_ShowContainerUI(ByVal fShow As Long)
'Err.Raise E_NOTIMPL
'fShow = True
End Sub
'=====================IRichEditOleCallback===================================
