VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "DESCRIPTION OF CPARAM CLASS"
Option Explicit
Option Compare Text
Public Source
Public Url$
Private mParamCount&
Event Changed()

'Dim mCollection As Collection

Private nCallsCount&
Private transcount&
'Private sx 'Сортированный индекс sx(2,ParamCount) параметр/индекс
'Private mSorted As Boolean

'Public Property Get ATestMy() As Variant()
'ATestMy = Array()
'End Property
'Public Property Get ATestMy1(v() As Variant) As Variant()
'ATestMy1 = v
'End Property


Public Property Get Value()
Value = Source
End Property

Public Property Let Value(v)
'FillParams v
nCallsCount = nCallsCount + 1
'Me.Clear
ReDim Source(-1 To -1)

If TypeName(v) = "CParam" Then
    Source = v.Source
    Url = v.Url
ElseIf IsArray(v) Then
    If ArrayDims(v) = 1 Then
        Source = v
        If LBound(v) = 0 Then pArrayLine Source, -1, Empty, 2
    Else
        Source(-1) = v
    End If
Else
    Source(-1) = v
End If
'If VarType(nName) Then Source(-1) = nName
mParamCount = (1 + UBound(Source) - (UBound(Source) + 1) Mod 2) \ 2
nCallsCount = nCallsCount - 1
EventChange
End Property

'Public Property Get ZValue()
'ZValue = gCompress(ToString(Source))
'End Property
'
'Public Property Let ZValue(v0)
'Dim v
'Me.Clear
'v = ToVariant(gDecompress(vbNullString & v0))
'If VarType(v) And vbArray Then
'If ArrayDims(v) <> 1 Then Exit Property
'If LBound(v) <= 0 Then
'    If UBound(v) > -1 Then Source = v Else ReDim Preserve Source(-1 To -1): Source(-1) = v
'    mParamCount = (1 + UBound(Source) - (UBound(Source) + 1) Mod 2) \ 2
'    RaiseEvent Changed
'End If
'End If
'End Property

Public Property Get XML$(Optional ByVal sParamPath$ = vbNullString)
'Replace("<?xml version='1.0' encoding='Windows-1251'?>", "'", Chr(34)) & vbCr &
XML = xml_code(NzS(PathParamName(sParamPath), "root"), PathParam(sParamPath))
End Property

Public Property Let XML(Optional ByVal sParamPath$ = vbNullString, xml_src$)
PathParam(sParamPath) = xml_decode(xml_src)
End Property

Public Property Get json$(Optional ByVal sParamPath$)
'Dim lv&: lv = -1
'Dim nEscapeOptions&, a, i&
'nEscapeOptions = 1: i = 1 '=QUOT STRINGS
'For Each a In Split("*,^,~,"",!,@", ",")
''1 LEVEL TABS           = *
''2 REPLACE ch34  \q = ^
''4 UNICODE                 = ~
''8 QUOT NAMES        = "
''16 NO QUOT ALL STRINGS if NOT [,]{":}       = !
''32 QUOT strings with "  by  «»    = @
''64 LEVEL TABS=1=           = *
'    If InStr(sParamPath, a) Then
'        If i = 1 Then lv = 0 Else nEscapeOptions = nEscapeOptions Or i
'        sParamPath = Replace(sParamPath, a, "")
'    End If
'    i = i * 2
'Next
'json = pCP2JSON2(PathParam(sParamPath), "item", lv, nEscapeOptions)

Dim lv&, nEscapeOptions&, a, i&
nEscapeOptions = 1: i = 2 '=QUOT STRINGS
For Each a In Split("^,~,"",!,@", ",")
'1 LEVEL TABS           = *
'2 REPLACE ch34  \q = ^
'4 UNICODE                 = ~
'8 QUOT NAMES        = "
'16 NO QUOT ALL STRINGS if NOT [,]{":}       = !
'32 QUOT strings with "  by  «»    = @
'64 LEVEL TABS=1=           = *
    If InStr(sParamPath, a) Then
        nEscapeOptions = nEscapeOptions Or i
        sParamPath = Replace(sParamPath, a, "")
    End If
    i = i * 2
Next
Dim lvMax&: lv = -1
Do While InStr(sParamPath, "*"): lv = 0: lvMax = lvMax + 1: sParamPath = Replace(sParamPath, "*", "", , 1): Loop
If lv Then lvMax = -1
json = pCP2JSON2(PathParam(sParamPath), "item", lv, lvMax, nEscapeOptions)

End Property

Public Property Let json(Optional ByVal sParamPath$, ByVal strJson$)
'PathParam(sParamPath) = pJSON2CP(strJson, "item")
'PathParam(sParamPath) = pJson2cpS(strJson, "item")
'PathParam(sParamPath) = pJson2cpN(strJson, "item")
PathParam(sParamPath) = pJson2cpA(StrPtr(strJson), "item")
End Property

Function View(Optional ByVal ParamPath$)
Dim st As New Stream: st.Type = adTypeText: st.Open
pView PathParam(ParamPath), 0, st
st.Position = 0
View = st.ReadText
End Function

Function Map(Optional ByVal ParamPath = "", Optional bDisp As Boolean) As Collection
Set Map = pCollect(PathParam(ParamPath), bDisp)
'For Each ob In cp.Collect("item2")
'    Debug.Print ob.Name, TypeName(ob.Value)
'Next
End Function
Private Function pCollect(ar, bDisp As Boolean) As Collection
Dim i&, pColl As New Collection
Dim ob As CDataObject
If ArrayDims(ar) = 1 Then
On Error Resume Next
    For i = 0 To UBound(ar) Step 2
        Set ob = New CDataObject
        ob.Name = S_(ar(i))
        'ob.Value = ar(i + 1)
        If IsArray(ar(i + 1)) And bDisp Then ob.Value = pCollect(ar(i + 1), bDisp) Else ob.Value = ar(i + 1)
        pColl.Add ob ', ar(i)
    Next
End If
Set pCollect = pColl
End Function

'Function Dispatch(Optional ByVal ParamPath = "") As Object 'READONLY
'Set mCollection = pCollect(PathParam(ParamPath), True)
'Dim uu As UUID
'Set Dispatch = CreateEventSinkObj(uu, mCollection, Nothing)
'
''Sub TestCPDispatch()
''Dim cp As CParam
''Set cp = xMain.CParam("{item1:1,item2:{item20:20,item21:21}}", "json")
''Dim ob As Object
''With cp.Dispatch
''    Debug.Print .item1 =1
''    Debug.Print .item2.item21 =21
''End With
''End Sub
'
'End Function

'Function URLParams(Optional ByVal ParamPath$, Optional ByVal prefix$)
'URLParams = xMain.URLParams(PathParam(ParamPath), prefix)
'End Function

'Property Get Parse(ByVal colSep, ByVal rowSep, Optional ByVal ParamPath$)
''Dim ar, i&, nm$, v, res$
''ar = PathParam(ParamPath)
''For i = 0 To UBound(ar) Step 2
''    v = ar(i + 1)
''    If IsArray(v) Then
''        'res=res & ar(i) &"[]"
''    Else
''        If Len(res) Then res = res & rowSep
''        If VarType(v) = vbBoolean Then v = LCase(xMain.FormatVB(v, "true/false", 1033))
''        res = res & ar(i) & colSep & v
''    End If
''Next
''Parse = res
'xMain.DebugPrint 241, "CParam.Parse DEPRICATED use URLParams"
'Parse = xMain.URLParams(PathParam(ParamPath))
'End Property



'Property Let Parse(ByVal colSep, ByVal rowSep, Optional ByVal ParamPath$, v)
'Dim row, ar, p$
'If Len(ParamPath) Then p = ParamPath & "\"
'For Each row In Split(S_(v), rowSep)
'    ar = Split(row, colSep)
'    If UBound(ar) = 1 Then PathParam(p & ar(0)) = ar(1)
'Next
'End Property


Public Sub BeginUpdate()
transcount = 0
End Sub
Public Sub EndUpdate()
If transcount > 0 Then RaiseEvent Changed
transcount = -1
End Sub

Private Sub EventChange(Optional ByVal VCallsCount&)
If (nCallsCount = VCallsCount) And (transcount = -1) Then
    RaiseEvent Changed
ElseIf transcount > -1 Then
    transcount = transcount + 1
End If
End Sub

Public Function GetArray(Optional ByVal bValues As Boolean, Optional ByVal ParamPath = vbNullString, Optional ByVal requr&)
Dim ar, i&, v
ar = PathParam(ParamPath)
With mSafeArray(ar)
If .cDims = 1 And .Bounds(1).lLbound = -1 And .Bounds(1).cElements > 1 Then
    ReDim res(.Bounds(1).cElements \ 2 - 1)
    For i = 0 To UBound(res)
        If requr And IsArray(ar(i * 2 + 1)) Then
            v = ar(i * 2)
            v = GetArray(bValues, NzS(ParamPath, , , "\") & v, requr)
        Else
            If requr > 0 And Not bValues Then
                v = NzS(ParamPath, , , "\") & ar(i * 2)
            Else
                v = ar(i * 2 + IIf(bValues, 1, 0))
            End If
        End If
        res(i) = v
    Next
    GetArray = res
Else
    If requr > 0 And Not bValues Then GetArray = ParamPath Else GetArray = Null   'Array(ar)
End If
End With
End Function

Public Function Names$(Optional ByVal ParamPath = vbNullString)
Names = xMain.Join2(GetArray(0, ParamPath), ",")
End Function


'Sub ReplaceNamesAR(oldNames, newNames)
'Dim i&, cp As New CParam
'If (VarType(oldNames) And VarType(newNames) And vbArray) = 0 Then Exit Sub
'If UBound(oldNames) <> UBound(newNames) Then Exit Sub
'For i = 0 To UBound(oldNames)
'cp.Param(oldNames(i)) = newNames(i)
'Next
'ReplaceNamesCP cp
'End Sub
'Sub ReplaceNamesCP(cp As CParam)
'Dim i&, nm$
'If cp.ParamCount = 0 Then Exit Sub
'For i = 0 To ParamCount
'nm = cp.ParamDef(ParamName(i), vbNullString)
'If Len(nm) Then ParamName(i) = nm
'Next
'End Sub


Public Sub Save(ByVal key)
Dim n%, buf$, pth

pth = Split(CPFilePath(key) & "/", "/")
On Error Resume Next
If Len(Dir(pth(0) & "\nul", vbDirectory)) = 0 Then MakeDir pth(0)

If Len(pth(1)) Then 'Save to PathParam
    Dim cp As New CParam
    cp.Load pth(0)
    cp(pth(1)) = Source
    cp.Save pth(0)
Else
    Kill pth(0)
    Dim st As New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    st.Write gCompress(ToString(Source, True))
    st.SaveToFile pth(0)
    st.Close
End If
End Sub

Public Function Load(ByVal key$)
mParamCount = 0
Dim tar, pth

If Left(key, 1) = "{" Or Left(key, 1) = "[" Then
    tar = pJson2cpA(StrPtr(key))
Else

pth = Split(CPFilePath(key) & "/", "/")  'pth(0)=file:pth(1)=parampath
Dim p$, a
p = pth(1)
If xMain.LenFile(pth(0)) > 20 Then
    
    Dim sf$: sf = xMain.ReadFile(pth(0), 20)
    If Mid$(sf, 5, 15) = "Standard Jet DB" Then
        Dim xcon As Connection
        Set xcon = xmdb_conn 'xMain.XMDB(True)
        xMain.XMDB(True) = pth(0)
        tar = xMain.XMDB(NzS(pth(1), "Global"))
        Set xmdb_conn = xcon: p = ""
    ElseIf Mid$(sf, 1, 6) = "PK" & Chr(3) & Chr(4) & Chr(20) & Chr(0) Then 'ZIPED
        Dim zip As New CZip: zip.OpenZip pth(0)
        Dim cp As New CParam: cp.Value = zip.ListZip
        p = ""
        For Each a In Split(pth(1), "\")
            p = p & a
            If IsArray(cp(p)) Then
                If L_(cp(p & "\size")) > 0 Then
                    Dim buf() As Byte
                    buf = zip.ReadBytes(p)
                    If gIsCompressed(buf) Then tar = ToVariant(gDecompress(buf)) Else tar = StrConv(buf, vbUnicode)
                    p = Replace(pth(1) & "\", p & "\", "")
                    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
                    'Debug.Print cp.json(pth(1))
                    Exit For
                End If
            Else
                Exit For
            End If
            p = p & "\"
        Next
        zip.CloseZip
    Else
        Dim st As New ADODB.Stream
        st.Open: st.Type = adTypeBinary
        st.LoadFromFile pth(0)
        tar = ToVariant(gDecompress(st.Read))
        If Not IsArray(tar) Then
            tar = Empty
            st.Position = 0: st.Type = adTypeText
            st.charset = "_autodetect_all"
            tar = st.ReadText
            'Debug.Print tar
            'Debug.Print "=" & Left$(tar, 9) & "="
            'Debug.Print "=" & Right$(tar, 1) & "="
            If Len(tar) > 10 Then If (Left$(tar, 9) = xs.sxControl & "{") And (Right$(tar, 1) = "}") Then json = tar: tar = Source
        ElseIf UBound(tar) = -1 Then 'crypted cparam
            If LBound(tar) = -1 Then
            Debug.Print 333333333
'                tar = decrypt(tar(-1), appkeys)
            End If
        End If
        st.Close
        
    End If


ElseIf Len(key) Then
    On Error Resume Next
    tar = modMain.ToVariant(gDecompress(LoadResData(key, "CUSTOM")))
    If Not IsEmpty(tar) Then Url = key
End If

Url = pth(0) & IIf(Len(pth(1)), "/" & pth(1), "")

End If

Err.Clear
If IsEmpty(tar) Then Exit Function
nCallsCount = nCallsCount + 1

Value = tar: tar = Empty
If Len(p) Then Value = PathParam(p)    'указан ParamPath
Load = Url

nCallsCount = nCallsCount - 1
EventChange
End Function

Public Sub Clear()
nCallsCount = nCallsCount + 1
ReDim Source(-1 To -1), sx(1, 0)
mParamCount = 0
nCallsCount = nCallsCount - 1
EventChange
End Sub


Public Function AddParams(cpar, Optional ByVal pth = "", Optional ByVal bAddTree As Boolean) 'Optional ByVal nAddType As Boolean) 'Добавить параметры  Source <+ ArrayCParam
'nAddType = 0  ADD+MODIFY
'nAddType<>0  ONLY MODIFY
Dim i&, nm$, v, ar
On Error Resume Next
Dim cp As New CParam
If TypeName(cpar) = "CParam" Then ar = cpar.Source Else If IsArray(cpar) Then ar = cpar Else If Not VT__(cpar) Then ar = xMain.CParam(cpar, "json").Value
If ArrayDims(ar) = 1 Then
    nCallsCount = nCallsCount + 1
    For i = 0 To UBound(ar) - ((UBound(ar) + 1) Mod 2) Step 2
        nm = ar(i): v = ar(i + 1)
        If Len(pth) Then nm = pth & "\" & nm
        If IsArray(v) And bAddTree Then
            cp.Value = PathParam(nm)
            v = cp.AddParams(v, , bAddTree)
            If UBound(v) = -1 Then ar(i + 1) = v(-1) Else ar(i + 1) = v
        End If
        PathParam(nm) = ar(i + 1)
        
'        PathParam(ar(i)) = ar(i + 1)
        
'        v = ar(i + 1): nm = ar(i)
'        If Len(pth) Then nm = pth & "\" & nm
'        If IsArray(v) Then v = xMain.CParam(PathParam(nm)).AddParams(v, "", nAddType)
'        If nAddType Then If PathParamIndex(nm) = -1 Then nm = ""
'        If Len(nm) Then PathParam(nm) = v

    Next
    nCallsCount = nCallsCount - 1
End If
AddParams = Source
'Set AddParams = Me
If i > 0 Then EventChange
End Function

Public Sub FillNames(ByVal sarPathNames, Optional ByVal iValue)
Dim i&, vt&, ar, v, iValueAr
i = VarType(sarPathNames)
'tar =Array of ParamNames or comma separated string with  ParamNames
'VarType(iValue)=vbEmpty = добавляем отсутствующие
'iVarType(Value)=vbNull = удаляем присутствующие
'iVarType(Value)>vbNull = добавляем отсутствующие/обновляем присутствующие
If i < 2 Or i = 10 Then Exit Sub
If i And vbArray Then ar = sarPathNames Else ar = Split(sarPathNames, ",")
If Not IsArray(ar) Then Exit Sub

If ArrayDims(iValue) = 1 Then iValueAr = iValue: ReDim Preserve iValueAr(UBound(ar))

vt = VarType(iValue): If vt = 10 Then iValue = Empty

nCallsCount = nCallsCount + 1
For i = 0 To UBound(ar)
    'PathParam(ar(i)) = iValue
    If vt And vbArray Then v = Nz(iValueAr(i), Empty) Else v = iValue
    PathParam(ar(i)) = v
Next
nCallsCount = nCallsCount - 1
EventChange
End Sub



'Public Sub FillParams(v, Optional ByVal nName = Empty)
'nCallsCount = nCallsCount + 1
'Me.Clear
'If TypeName(v) = "CParam" Then
'    Source = v.Source
'ElseIf IsArray(v) Then
'    If ArrayDims(v) = 1 Then
'        Source = v
'        If LBound(v) = 0 Then pArrayLine Source, -1, nName, 2
'    Else
'        Source(-1) = v
'    End If
'End If
'If VarType(nName) Then Source(-1) = nName
'mParamCount = (1 + UBound(Source) - (UBound(Source) + 1) Mod 2) \ 2
'nCallsCount = nCallsCount - 1
'EventChange
'End Sub

Public Property Get Name()
On Error Resume Next
Name = Source(-1)
End Property
Public Property Let Name(v)
On Error Resume Next
Source(-1) = v
EventChange
End Property


Public Property Get ParamCount&()
    ParamCount = mParamCount
End Property

Public Property Get ParamName$(ByVal ParamKey)
'на выходе Name параметра
If VarType(ParamKey) < 2 Then Exit Property
Dim index As Long
If VarType(ParamKey) = vbString Then index = GetIndex(ParamKey) Else index = Val(ParamKey)   'Получаем индекс параметра
If index > -1 And index < mParamCount Then 'Есть такой параметр
    ParamName = Source(index * 2 + 0)
'Else
'    ParamName = vbNullString 'Null
End If
End Property
Public Property Let ParamName(ByVal ParamKey, newValue$)
'Dim vt&
'vt = VarType(ParamKey)
'If vt < 2 Then Exit Property
'If vt And vbArray Then Exit Property
If VT__(ParamKey) Then Exit Property
If Len(newValue) = 0 Then Exit Property
If VarType(ParamKey) = vbString Then If ParamKey = vbNullString Then Name = newValue: Exit Property
If GetIndex(newValue) > -1 Then Exit Property 'Нельзя менять имя потомушто такое имя уже есть
Dim index&
If VarType(ParamKey) = vbString Then index = GetIndex(ParamKey) Else index = ParamKey   'Получаем индекс параметра
nCallsCount = nCallsCount + 1
If index > -1 And index < mParamCount Then
    Source(index * 2) = newValue
    EventChange 1
End If
nCallsCount = nCallsCount - 1
End Property

Public Property Get ParamDef(ByVal ParamKey, ByVal Default, Optional ByVal bRemove As Boolean)
'bRemove = 0
Dim res
res = Param(ParamKey)
If bRemove Then If VarType(res) > 1 Then Param(ParamKey) = Null 'Удаляем прочитанный параметр
If IsNull(res) Then ParamDef = Default Else ParamDef = Cast(res, VarType(Default))
End Property

'Public Function Def(ByVal ParamPath, ByVal Default)
'Dim res: res = PathParam(ParamPath)
'If IsNull(res) Then ParamDef = Default Else ParamDef = Cast(res, VarType(Default))
'End Function

Public Property Get Param(ByVal ParamKey)
If VarType(ParamKey) < 2 Then Exit Property
Dim index&, ar
If VarType(ParamKey) = vbString Then index = GetIndex(ParamKey) Else index = ParamKey  'Получаем индекс параметра
If index > -1 And index < mParamCount Then 'Есть такой параметр
    Param = Source(index * 2 + 1) 'значение параметра
Else
    Param = Null
'    If ArrayDims(Source(-1)) = 2 Then '2DIM ARRAY
'        On Error Resume Next
'        ar = Split(ParamKey & ",", ",")
'        Param = Source(-1)(L_(ar(0)), L_(ar(1)))
'        Err.Clear
'    End If
End If
End Property

Public Property Let Param(ByVal ParamKey, newValue)
Dim index&, n&, i&, vt&
'Dim f As Boolean
'vt = VarType(ParamKey)
'If vt < 2 Or vt = 10 Or (vt And vbArray) > 0 Then Exit Property
If VT__(ParamKey) Then Exit Property

nCallsCount = nCallsCount + 1

If VarType(ParamKey) = vbString Then index = GetIndex(ParamKey) Else index = ParamKey     'Получаем индекс параметра
n = mParamCount
vt = VarType(newValue)
If index > -1 And index < n Then 'есть такой индекс
    index = index * 2
    If vt = vbNull Then 'Удаляем параметр NULL
        pArrayLine Source, index, , 3: pArrayLine Source, index, , 3
    Else 'Записываем параметр
        If vt Then Source(index + 1) = newValue Else Source(index + 1) = Null ' isEmpty
    End If
'    f = 1
ElseIf vt <> 1 And VarType(ParamKey) = vbString Then 'Добавляем именованный параметр
'    If ArrayDims(Source(-1)) = 2 Then '2DIM ARRAY
'        On Error Resume Next
'        Dim a2i
'        a2i = Split(ParamKey & ",", ",")
'        Source(-1)(L_(a2i(0)), L_(a2i(1))) = IIf(vt = 0, Null, newValue)
'        Err.Clear
'    Else
        n = (n + 1) * 2 - 1
        ReDim Preserve Source(-1 To n)
        Source(n - 1) = ParamKey
        If vt Then Source(n) = newValue Else Source(n) = Null ' isEmpty
'    End If
'    f = 1
End If
'If f Then mParamCount = (1 + UBound(Source) - (UBound(Source) + 1) Mod 2) \ 2
mParamCount = (1 + UBound(Source) - (UBound(Source) + 1) Mod 2) \ 2
nCallsCount = nCallsCount - 1
'If f Then EventChange
EventChange
End Property


'!!!!!! DEFAULT MEMBER !!!!!!
Public Property Get PathParam(ByVal ParamPath)  'ParamPath$ = "xControls\Control1\xControls\Control1\xColumns\Column1"
Attribute PathParam.VB_UserMemId = 0
'Const speed2 = 0
'If speed Then '+10000%
    Dim ps$, p, v, i&
    If VarType(ParamPath) < 2 Then Exit Property
    ps = ParamPath ': If Len(ps) = 0 Then PathParam = Source: Exit Property
    PathParam = Null
'    If speed2 Then
    CopyMemory v, Source, 16
'    Else
'    v = Source
'    End If
    Do
        For Each p In Split(ps, "\")
            If Not IsArray(v) Then Exit Do
                If Len(p) = 0 Then Exit For
                For i = 0 To UBound(v) Step 2
                    If VarType(v(i)) = vbString Then If StrComp(vbNullString & v(i), p, vbTextCompare) = 0 Then Exit For
                Next
                If i > UBound(v) Then Exit Do
'                If speed2 Then
                CopyMemory v, v(i + 1), 16
'                Else
'                v = v(i + 1)
'                End If
        Next
        PathParam = v
    Exit Do
    Loop
'    If speed2 Then
    FillMemory v, 16, 0
'    End If
'Else
'
'If VarType(ParamPath) < 2 Then Exit Property
'If ParamPath = vbNullString Then PathParam = Source: Exit Property
'Dim pth, n&, key, nm$
''Dim pth, i&, n&, key, nm$, v
'pth = Split(ParamPath, "\"): n = UBound(pth): key = pth(i): v = Param(key)
'If n = 0 Then PathParam = v: Exit Property
'Dim cp As New CParam: cp.Value = v
'For i = 1 To n: nm = nm & IIf(Len(nm), "\", vbNullString) & pth(i): Next
'PathParam = cp(nm)
'End If
End Property
'!!!!!! DEFAULT MEMBER !!!!!!


Public Property Let PathParam(ByVal ParamPath, newValue)
If VarType(ParamPath) < 2 Then Exit Property
Dim pth, i&, n&, key, nm$ ', v
nCallsCount = nCallsCount + 1
If ParamPath = vbNullString Then
    If IsArray(newValue) Then
        Value = newValue
    Else
        If IsArray(Source(-1)) Then
            Param(",") = newValue
        Else
            Clear
        End If
    End If
    mParamCount = (1 + UBound(Source) - (UBound(Source) + 1) Mod 2) \ 2
Else
    pth = Split(ParamPath, "\")
    n = UBound(pth)
    key = pth(0)
    If n = 0 Then
        Param(key) = newValue
    Else
        Dim cp As New CParam
        cp.Value = Param(key)
        For i = 1 To n
        nm = nm & IIf(Len(nm), "\", vbNullString) & pth(i)
        Next
        cp(nm) = newValue
        If cp.ParamCount = 0 Then
            Param(key) = cp.Name
        Else
            Param(key) = cp.Source
        End If
        Set cp = Nothing
    End If
End If
nCallsCount = nCallsCount - 1
EventChange
End Property

'Public Sub MergePathParams(ByVal ParamPath, CPSource, Optional ByVal MergeType&)
''MergeType:
''0=If GetIndex(ParamPath)>-1 Then Modify Else AddNew
''1=If GetIndex(ParamPath)>-1 Then Modify
'If VarType(ParamPath) < 2 Then Exit Sub
'nCallsCount = nCallsCount + 1
'Dim cp As CParam, i&, p$, v, b As Boolean
'Set cp = New CParam
'cp.FillParams vbNullString, CPSource
'For i = 0 To cp.ParamCount - 1
'    v = Empty: v = cp.Param(i): b = IsArray(v)
'    If b Then b = (LBound(v) = -1)
'    p = ParamPath & IIf(ParamPath = vbNullString, vbNullString, "\") & cp.ParamName(i)
'    'Debug.Print b, "=", p
'    If b Then
'        MergePathParams p, v
'    Else
'        PathParam(p) = v
'    End If
'Next
'Set cp = Nothing
'nCallsCount = nCallsCount - 1
'EventChange
'End Sub

Public Property Get PathParamCount&(ByVal ParamPath) 'ParamPath$ = "xControls\Control1\xControls\Control1\xColumns\Column1"
If VarType(ParamPath) < 2 Then Exit Property
Dim pth, i&, n&, key, nm$
pth = Split(ParamPath, "\")
n = UBound(pth)
If n < 0 Then PathParamCount = mParamCount: Exit Property
key = pth(i)
Dim cp As CParam
Set cp = New CParam
cp.Value = Param(key)
For i = 1 To n
nm = nm & IIf(Len(nm), "\", vbNullString) & pth(i)
Next
PathParamCount = cp.PathParamCount(nm)
Set cp = Nothing
End Property


Public Property Get PathParamIndex(ByVal ParamPath)
If VarType(ParamPath) < 2 Then Exit Property
PathParamIndex = -1
If ParamPath = vbNullString Then Exit Property 'Source = newValue: Exit Property
Dim pth, i&, n&, key$, nm$
pth = Split(ParamPath, "\")
n = UBound(pth)
key = pth(0)
If n = 0 Then PathParamIndex = GetIndex(key): Exit Property

Dim cp As CParam
Set cp = New CParam
cp.Value = Param(key)
For i = 1 To n
nm = nm & IIf(Len(nm), "\", vbNullString) & pth(i)
Next
PathParamIndex = cp.PathParamIndex(nm)
Set cp = Nothing
End Property

Public Property Let PathParamIndex(ByVal ParamPath, nIndex)
If VarType(ParamPath) < 2 Then Exit Property
Dim pth, i&, n&, nm$
Dim pi&, pc&, npi&
If ParamPath = vbNullString Then Exit Property
pth = Split(ParamPath, "\")
n = UBound(pth)
For i = 0 To n - 1
nm = nm & IIf(Len(nm), "\", vbNullString) & pth(i)
Next
pc = PathParamCount(nm)
If pc < 2 Then Exit Property 'Параметров токо один или нету
pi = PathParamIndex(ParamPath)
If pi = -1 Then Exit Property 'Параметр не найден
npi = nIndex
If npi < 0 Then npi = 0
If npi > pc - 1 Then npi = pc - 1
If npi = pi Then Exit Property 'Не надо менять индекс он такой -же

nCallsCount = nCallsCount + 1
Dim cp As CParam
Set cp = New CParam
cp.Value = PathParam(nm)
cp.SwapParams pi, npi
If Len(nm) Then
    PathParam(nm) = cp.Source
Else
    Source = cp.Source
    mParamCount = (1 + UBound(Source) - (UBound(Source) + 1) Mod 2) \ 2
    'If mSorted Then sxUpdate "", mParamCount
End If
Set cp = Nothing
nCallsCount = nCallsCount - 1
EventChange

End Property

Friend Sub SwapParams(ByVal pi1&, ByVal pi2&)
Dim buf0, buf1
If pi1 = pi2 Or pi1 < 0 Or pi2 < 0 Or pi1 > mParamCount - 1 Or pi2 > mParamCount - 1 Then Exit Sub
buf0 = Source(pi1 * 2 + 0)
buf1 = Source(pi1 * 2 + 1)
Source(pi1 * 2 + 0) = Source(pi2 * 2 + 0)
Source(pi1 * 2 + 1) = Source(pi2 * 2 + 1)
Source(pi2 * 2 + 0) = buf0
Source(pi2 * 2 + 1) = buf1
EventChange
End Sub

Public Property Get PathParamName(ByVal ParamPath) 'ParamPath$ = "xControls\Control1\xControls\Control1\xColumns\Column1"
If VarType(ParamPath) < 2 Then Exit Property
If ParamPath = vbNullString Then PathParamName = Name: Exit Property
Dim pth, i&, n&, key, nm$
pth = Split(ParamPath, "\")
n = UBound(pth)
key = pth(i)
'If IsNumeric(key) Then key = CLng(key)
If n = 0 Then PathParamName = ParamName(key): Exit Property
Dim cp As CParam
Set cp = New CParam
cp.Value = Param(key)
For i = 1 To n
nm = nm & IIf(Len(nm), "\", vbNullString) & pth(i)
Next
PathParamName = cp.PathParamName(nm)
Set cp = Nothing
End Property

Public Property Let PathParamName(ByVal ParamPath, newValue)
If VarType(ParamPath) < 2 Then Exit Property
nCallsCount = nCallsCount + 1
If ParamPath = vbNullString Then
    Name = newValue
Else
    Dim pth, i&, n&, key, nm$
    pth = Split(ParamPath, "\")
    n = UBound(pth)
    key = pth(0)
    If n = 0 Then
        ParamName(key) = newValue
    Else
        Dim cp As New CParam
        cp.Value = Param(key)
        For i = 1 To n
        nm = nm & IIf(Len(nm), "\", vbNullString) & pth(i)
        Next
        cp.PathParamName(nm) = newValue
        Param(key) = cp.Source
        Set cp = Nothing
    End If
End If
nCallsCount = nCallsCount - 1
EventChange
End Property


'Public Function GetIndex(ByVal key, Optional ByVal sxi, Optional ByVal keyAny = 1) As Long
Public Function GetIndex(ByVal key) As Long
Dim i&, s$
GetIndex = -1: If VT__(key) Then Exit Function
s = key
If s = "\" Then GetIndex = VarPtr(Source): Exit Function
For i = 0 To mParamCount - 1
    If StrComp(vbNullString & Source(i * 2), s, vbTextCompare) = 0 Then GetIndex = i: Exit For
Next
End Function

'Private Sub sxUpdate(key, Optional sxIndex& = -1) 'Инициализация поискового индекса
'If Not mSorted Then Exit Sub
'Dim i&, i0&
'If sxIndex = mParamCount Then 'Добавлен параметр
'    i0 = mParamCount - IIf(key = "", 1, 0)
'    If i0 < 0 Then i0 = 0
'    ReDim Preserve sx(1, i0)
'ElseIf VarType(key) < 2 Then 'Удален параметр
'    If sxIndex > -1 Then 'Позиция удаленного параметра известна
'        i0 = sx(1, sxIndex)
'        For i = 0 To mParamCount - 2
'        If i >= sxIndex Then sx(0, i) = sx(0, i + 1): sx(1, i) = sx(1, i + 1)
'        If sx(1, i) > i0 Then sx(1, i) = sx(1, i) - 1
'        Next
'        i0 = mParamCount - 2
'    Else 'Позиция неизвестна
'        i0 = mParamCount - 1
'    End If
'    If i0 < 0 Then i0 = 0
'    ReDim Preserve sx(1, i0)
'    Exit Sub
'End If
'
'If sxIndex > -1 And Len(key) > 0 Then
'    sx(0, sxIndex) = key
'    If sxIndex = mParamCount Then sx(1, sxIndex) = mParamCount
'Else
'    For i = 0 To mParamCount - 1
'    sx(0, i) = Source(i * 2): sx(1, i) = i
'    Next
'End If
'QuickSort sx
'End Sub
Property Get SVal(ByVal index)
On Error Resume Next
SVal = Source(index)
Err.Clear
End Property
Property Let SVal(ByVal index, v)
On Error Resume Next
Source(index) = v
Err.Clear
End Property
Private Sub Class_Initialize()
ReDim Source(-1 To -1)
'Debug.Print "Initialize  CParam " & ObjPtr(Me)
'ReDim sx(1, 0)
transcount = -1
End Sub

'Private Sub Class_Terminate()
''Debug.Print "Terminate CParam " & ObjPtr(Me)
''Erase sx
''Erase Source ' = Empty
'mParamCount = 0
'End Sub


'Public Property Get NewEnum() As IUnknown
'    Set mCollection = New Collection
'    'this property allows you to enumerate
'    'this collection with the For...Each syntax
'    Dim i&
'    For i = 0 To UBound(Source) Step 2
'    mCollection.Add Source(i + 1), Source(i)
'    Next
'    Set NewEnum = mCollection.[_NewEnum]
'End Property

'Function TreeParam(ar, Optional ByVal vEntry = "")
'Dim i&, n&, v, ret
'ReDim ret(-1 To -1)
'For i = 0 To UBound(ar) Step 2
'    If IsArray(ar(i + 1)) Then
'        n = UBound(ret) + 2
'        ReDim Preserve ret(-1 To n)
'        ret(n - 1) = ar(i)
'        ret(n) = TreeParam(ar(i + 1), vEntry)
'    End If
'Next
'If UBound(ret) > -1 Then TreeParam = ret Else TreeParam = vEntry
'End Function
