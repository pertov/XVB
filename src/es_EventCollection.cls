VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEventCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 256
Attribute VB_Description = "Edanmo's EventCollection Class v2.0"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************
'
' EventCollection Class
'
'********************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Usage: at your own risk.
'
' Tested on:
'            * Windows XP Pro SP1
'            * VB6 SP5
'
' History:
'           01/02/2003 * This code replaces the old EventCollection
'                        class.
'
'********************************************************************************
Option Explicit

' Collection that will hold the objects
Private m_oCollection As Collection
Attribute m_oCollection.VB_VarHelpID = 272

' HandleEvent
' Raised when an event of a contained object is raised
Event HandleEvent(ByVal EventInfo As CEventInfo, ByVal pVarResult As Long)
'
' pvGetEventName
' Returns the name of an event given its DISPID
' Parameters:
' -----------
' Source - IDispatch interface of source object
' DISPIP - DISPID of the event
' IID    - Interface ID of the source interface
' Return Value: The event name. If the name can't be found
' the function returns the event DISPID.
'
Private Function pvGetEventName( _
      ByVal Source As IDispatch, _
      ByVal DISPID As Long, _
      IID As UUID) As Variant ', ParamFlags() As Long) As Variant
      
Dim oTL As ITypeLib
Dim oTI As ITypeInfo
Dim oTI2 As ITypeInfo2
Dim sEventName As String


On Error Resume Next
Set oTI = Source.GetTypeInfo()
If oTI.GetContainingTypeLib(oTL) Then
    Set oTI2 = oTL.GetTypeInfoOfIID(IID)
    If Not oTI2 Is Nothing Then oTI2.GetNames DISPID, sEventName, 1
End If

If Len(sEventName) = 0 Or IsInIDE Then
Select Case TypeName(Source)
Case "CDataObject": sEventName = Split(",OnError,OnStart,OnProgress,OnFinish", ",")(DISPID)
'Case "CMatrix": sEventName = Split(",DataChanged,StartChange,Canceled", ",")(DISPID)
Case "CMain": sEventName = "DumyEvent"
Case "CParam": sEventName = "Changed"
Case "CZip": sEventName = Split(",ZipFiles,ZipBytes,UnZipFiles,UnZipBytes,Cancel", ",")(DISPID)
End Select
End If


If Len(sEventName) = 0 Then pvGetEventName = "Event" Else pvGetEventName = sEventName
End Function


Public Sub Add( _
   Item As Object, _
   Optional SourceIID As String, _
   Optional key As Variant, _
   Optional ByVal arEvents, Optional ByVal svbMacros As String)
Dim oObjectInfo As CEventObjectInfo
Dim oCPC As IConnectionPointContainer
Dim oEnm As IEnumConnectionPoints
Dim oCP As IConnectionPoint
Dim oUnk As IUnknown
Dim tIID As UUID
Dim lCookie As Long

' Get the IConnectionPointContainer interface
Set oCPC = Item
If LenB(SourceIID) = 0 Then
    Set oEnm = oCPC.EnumConnectionPoints ' Get connection point enumerator
    oEnm.Next 1, oCP ' Get the first connection point
    oCP.GetConnectionInterface tIID ' Get the IID
Else
    CLSIDFromString SourceIID, tIID ' Convert from string to UUID
    On Error Resume Next
    Set oCP = oCPC.FindConnectionPoint(tIID) ' Get the connection point
End If

Set oObjectInfo = New CEventObjectInfo ' Create the ObjectInfo object
Set oUnk = CreateEventSinkObj(tIID, oObjectInfo, Me) ' Create the event sink object
lCookie = oCP.Advise(oUnk) ' Connect the sink object with  the source object

On Error GoTo Disconnect
m_oCollection.Add oObjectInfo, key ' Add the object to the collection
' Initialize the ObjectInfo object
oObjectInfo.frInitialize key, _
                        m_oCollection.count, _
                        lCookie, _
                        Item, _
                        tIID, _
                        "," & NzS(arEvents, "-") & ",", svbMacros
   
Exit Sub
   
Disconnect:
oCP.Unadvise lCookie
Err.Raise Err.Number, , Err.Description
End Sub

Public Sub Clear()
Dim oObjectInfo As CEventObjectInfo
For Each oObjectInfo In m_oCollection ' Disconnect all objects
    pvUnadvise oObjectInfo
Next
Set m_oCollection = New Collection ' Reset the collection
End Sub

Public Property Get count() As Long
   count = m_oCollection.count
End Property

' frRaiseEvent
' Raises the events. This method is called from each event sink object when a
' event is raised from the source object.
' Parameters:
' -----------
' lObjInfo - ObjectInfo object of the source object
' DISPID - DISPID of the event
' lParams - Pointer to a DISPPARAMS structure
'
Friend Sub frRaiseEvent( _
      ByVal ObjInfo As CEventObjectInfo, _
      ByVal DISPID As Long, _
      ByVal lParams As Long, _
      ByVal pVarResult As Long)
Dim oEventInfo As CEventInfo
Dim sEventName As String
Dim ParamFlags() As Long
On Error Resume Next
sEventName = pvGetEventName(ObjInfo.Object, DISPID, ObjInfo.IID) ', ParamFlags)
Set oEventInfo = New CEventInfo ' Create a new EventInfo object
' Initialize the EventInfo object
oEventInfo.frInitialize sEventName, DISPID, pvGetParamArray(lParams, ParamFlags), ObjInfo.Object, ObjInfo.key, ObjInfo.Events, ParamFlags, ObjInfo.VBMacros
RaiseEvent HandleEvent(oEventInfo, pVarResult) ' Raise the event
' Set the event parameters passed ByRef  = ERROR
If UBound(ParamFlags) > 0 Then pvSetParamArray lParams, oEventInfo.Param.frGetParams(), oEventInfo.Param.frGetParamFlags
End Sub

'CUSTOM RAISE EVENT
Public Function FireEvent(ByVal senderName As String, ByVal sEventName As String, ByVal DISPID As Long, oSender As Object, ParamArray args())
Dim ar(), ar3() As Long
ar = args
If UBound(ar) > -1 Then ReDim ar3(UBound(ar))
Dim oEventInfo  As New CEventInfo ' Create a new EventInfo object
sEventName = NzS(sEventName, "Event")
oEventInfo.frInitialize sEventName, DISPID, ar, oSender, senderName, "*,-", ar3, ""
Dim pVarResult As Long
pVarResult = VarPtr(FireEvent)
RaiseEvent HandleEvent(oEventInfo, pVarResult)  ' Raise the event
End Function

Public Property Get Item(key As Variant) As CEventObjectInfo
Attribute Item.VB_UserMemId = 0
On Error Resume Next
Set Item = m_oCollection.Item(key)
End Property

'Public Function NewEnum() As IEnumVARIANT
'   Set NewEnum = m_oCollection.[_NewEnum]
'End Function


' pvGetParamArray
' Returns the event parameters
' Parameters:
' -----------
' Ptr - Pointer to a DISPPARAMS structure
Private Function pvGetParamArray(ByVal ptr As Long, ParamFlags() As Long) As Variant()
Dim tDPAR As xstdole.DISPPARAMS
'Dim mSafeArray As SAFEARRAY_1D
Dim mSA As SAFEARRAY1D
Dim aTmpParams() As Variant
Dim aParams() As Variant
Dim lIdx As Long

CopyMemory tDPAR, ByVal ptr, Len(tDPAR) ' Get the DISPARAMS structure form the poitner
If tDPAR.cArgs = 0 Then Exit Function ' If there're no parameter exit the function
With mSA ' Fill a SAFEARRAY struct  that points to the parameters
    .rgsabound.cElements = tDPAR.cArgs
    .rgsabound.lLbound = 0
    .cDims = 1
    .cbElements = 16
    .pvData = tDPAR.rgPointerToVariantArray
End With
PutMem4 ByVal VarPtrArray(aTmpParams), VarPtr(mSA)   ' Set the array with the SAFEARRAY struct
ReDim aParams(0 To tDPAR.cArgs - 1)
ReDim ParamFlags(0 To tDPAR.cArgs - 1) As Long
For lIdx = 0 To tDPAR.cArgs - 1 ' Copy the array elements
    aParams(lIdx) = aTmpParams(tDPAR.cArgs - lIdx - 1)
    If VarVT(aTmpParams(tDPAR.cArgs - lIdx - 1)) And VT_BYREF Then ParamFlags(lIdx) = 3
Next
pvGetParamArray = aParams
PutMem4 ByVal VarPtrArray(aTmpParams), 0& ' Release the array
End Function


' pvSetParamArray
' Copies the parameters passed ByRef back to the original array
' Parameters:
' -----------
' Ptr       - Pointer to a DISPPARAMS structure
' aParams() - The parameters array returned by the event handler
Private Sub pvSetParamArray(ByVal ptr As Long, aParams() As Variant, aParamsFlag() As Long)
Dim tDPAR As xstdole.DISPPARAMS
'Dim SafeArray As SAFEARRAY_1D
Dim iVarType As Integer
Dim lIdx As Long, lPtr As Long
Dim lVarPtr As Long
CopyMemory tDPAR, ByVal ptr, Len(tDPAR) ' Get the DISPARAMS structure form the poitner
If tDPAR.cArgs = 0 Then Exit Sub ' If there're no parameter exit the sub
lVarPtr = tDPAR.rgPointerToVariantArray ' Get the pointer to the first variant
Dim pindex
For lIdx = 0 To tDPAR.cArgs - 1
    pindex = tDPAR.cArgs - lIdx - 1
    GetMem2 ByVal lVarPtr + (lIdx * 16&), iVarType '' Get variant type. Cannot use VB's VarType because it removes VT_BYREF
    If aParamsFlag(pindex) And &H1000 Then '
        'If (iVarType And VT_BYREF) = VT_BYREF Then
        GetMem4 ByVal lVarPtr + (lIdx * 16&) + 8&, lPtr ' The parameter was passed ByRef   Get the data pointer from the variant
        Select Case iVarType And VT_TYPEMASK
        Case vbString
            Dim lStrPtr As Long
            GetMem4 ByVal lPtr, lStrPtr ' Get the string pointer
            If lStrPtr = 0 Then    ' Allocate a new string
                lStrPtr = SysAllocStringPtr(aParams(pindex))
            Else   ' Reallocate the string
                lStrPtr = SysReAllocString(lStrPtr, StrPtr(aParams(pindex)))
            End If
            PutMem4 ByVal lPtr, lStrPtr ' Copy the string pointer
        Case vbInteger
            PutMem2 ByVal lPtr, CInt(aParams(pindex))
        Case vbBoolean
            PutMem2 ByVal lPtr, CBool(aParams(pindex))
        Case vbInteger
            PutMem2 ByVal lPtr, CInt(aParams(pindex))
        Case vbLong
            PutMem4 ByVal lPtr, CLng(aParams(pindex))
        Case vbSingle
            PutMem4 ByVal lPtr, CSng(aParams(pindex))
        Case vbDouble
            CopyMemory ByVal lPtr, CDbl(aParams(pindex)), 8&
        Case vbDate
            CopyMemory ByVal lPtr, CDate(aParams(pindex)), 8&
        Case vbByte
            PutMem1 ByVal lPtr, CByte(aParams(pindex))
        Case vbCurrency
            CopyMemory ByVal lPtr, CCur(aParams(pindex)), 8&
        Case vbVariant, vbDecimal
            VariantCopyIndPtr ByVal lPtr, ByVal VarPtr(aParams(pindex))
        End Select
    End If
Next
End Sub


' pvUnadvise
' Disconnects the sink from the connection point
Private Sub pvUnadvise(ObjInfo As CEventObjectInfo)
Attribute pvUnadvise.VB_HelpID = 261
Dim oCPC As IConnectionPointContainer
Dim oCP As IConnectionPoint
Set oCPC = ObjInfo.Object ' Get the IConnectionPointContainer interface
On Error Resume Next
Set oCP = oCPC.FindConnectionPoint(ObjInfo.IID) ' Get the IConnectionPoint
oCP.Unadvise ObjInfo.Cookie
End Sub

Public Sub Remove(obj_key As Variant)
Dim oObjectInfo As CEventObjectInfo
'Dim oCPC As IConnectionPointContainer
'Dim oCP As IConnectionPoint
Dim key$
If IsObject(obj_key) Then 'Remove by object
    For Each oObjectInfo In m_oCollection
    If oObjectInfo Is obj_key Then key = oObjectInfo.key: Exit For
    Next
Else 'remove by key
    key = obj_key
End If
On Error Resume Next
Set oObjectInfo = m_oCollection(key) ' Get the object info
If oObjectInfo Is Nothing Then Exit Sub
pvUnadvise oObjectInfo ' Disconnect the object
m_oCollection.Remove key ' Remove the object
'If Err Then
'Debug.Print
'Debug.Print ">>>>>>>>"; Err.Description: Err.Clear
'Debug.Print
'End If
Remove obj_key '”далить еще один если он есть
End Sub

Private Sub Class_Initialize()
Attribute Class_Initialize.VB_HelpID = 259
Set m_oCollection = New Collection ' Initialize the collection
End Sub

Private Sub Class_Terminate()
Attribute Class_Terminate.VB_HelpID = 258
Clear
Set m_oCollection = Nothing
End Sub


