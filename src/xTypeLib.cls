VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xTypeLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_isReference As Boolean
Private m_isloaded As Boolean

Private m_FileLoad As String
Private m_FileCreate As String

Private m_Name As String
Private m_DocString As String
Private m_HelpContext As Long
Private m_HelpFile As String

Private la As TLIBATTR

Private m_xTypes As Collection

Private refs As Collection

Private m_tlb As ITypeLib
Private m_ctlb As ICreateTypeLib
Private m_saf
Private m_vbmode As Boolean

Public Property Get TypeFilter() As String
If ArrayDims(m_saf) = 1 Then TypeFilter = PackNums(Join(m_saf, ","))
End Property
Public Property Let TypeFilter(sarFilter As String)
If ArrayDims(sarFilter) = 1 Then m_saf = sarFilter Else m_saf = Split(Nz(sarFilter), ",")
m_saf = Split(UnPackNums(Join(m_saf, ",")), ",")
End Property

Friend Function bSAF(f, Optional nFindOptions As Long) As Boolean
bSAF = True
If IsArray(m_saf) Then
If UBound(m_saf) > -1 Then bSAF = gFindIndex(m_saf, f, nFindOptions) > -1 Or gFindIndex(m_saf, "" & f, nFindOptions) > -1 Else bSAF = f <> 50
End If
End Function

Private Sub Class_Initialize()
pCleanUp
m_vbmode = True
End Sub
Private Sub pCleanUp()
m_isReference = True
m_isloaded = False
m_FileLoad = ""
m_FileCreate = ""
m_Name = ""
m_DocString = ""
m_HelpContext = 0
m_HelpFile = ""
FillMemory la, LenB(la), 0
NewCollection m_xTypes
NewCollection refs
Set m_tlb = Nothing
Set m_ctlb = Nothing
End Sub


Function CreateFile(tlb_path)
On Error GoTo err_fun

Dim file$
file = Nz(tlb_path)
If Len(file) = 0 Then
    If Len(m_FileLoad) = 0 Then Exit Function
    file = m_FileLoad & ".tlb"
End If

'If SplitIndex(Split(m_FileLoad, ".")) = "tlb" Then
If m_FileLoad = "" And TLBFileLen(file) > 0 Then Load file
m_FileCreate = file
Set m_ctlb = CreateTypeLib2(SYS_WIN32, m_FileCreate)
'm_HelpString = m_FileLoad
'If m_xTypes.Count Then m_xTypes.Remove 1

Create
Save

Exit Function
err_fun:
Debug.Print Err.Description
Debug.Assert False
Resume Next
End Function

Function Load(tlb_path) As Boolean
pCleanUp
m_FileLoad = Nz(tlb_path)
If TLBFileLen(m_FileLoad) = 0 Then
    m_FileLoad = ""

    Set m_ctlb = CreateTypeLib2(SYS_WIN32, "")
    Dim uu As UUID
    m_ctlb.SetName "_EmptyTypeLib_"
    CoCreateGuid uu
    m_ctlb.SetGuid uu
    m_ctlb.SetVersion 1, 0
    m_ctlb.SetLcid 0
    m_ctlb.SetDocString "Test Empty TypeLib"
    LoadInfo m_ctlb, False
    m_isloaded = True
    
Else


    LoadInfo LoadTypeLib(m_FileLoad), False
    
    #If LoadTLB Then
    Debug.Print m_Name & "= " & m_HelpString
    #End If
    
    On Error GoTo erfun
    LoadTypes
    Load = True
End If
'GetMembersTree
Exit Function

erfun:
DebugPrint "xTypelib.Load", Err.Description
Err.Clear
Debug.Assert False
'Resume Next
End Function


Friend Function LoadTypes() As Boolean
On Error GoTo erfun
If m_tlb Is Nothing Then Exit Function
NewCollection m_xTypes
LoadTypes = True
Dim index&
For index = 0 To m_tlb.GetTypeInfoCount - 1
    LoadTypes = LoadTypes And PreLoadType(index, m_tlb.GetTypeInfo(index), 0)
Next
Exit Function

erfun:
DebugPrint "xTypelib.LoadTypes", Err.Description
Err.Clear
Debug.Assert False
'Resume Next
End Function

Friend Function PreLoadType(t_index&, cls As ITypeInfo, isob As Boolean) As Boolean
On Error GoTo erfun
Dim x As xType
Dim x0 As xType
Dim nm$
Dim bAdd As Boolean
cls.GetDocumentation DISPID_UNKNOWN, nm, vbNullString, 0, vbNullString
If Len(nm) Then
    PreLoadType = True
    Set x0 = FindTypeByName(nm)
    bAdd = x0 Is Nothing
    If bAdd Then
        Set x = New xType
        m_xTypes.Add x, nm
        If Not x.Load(t_index, cls, Me, isob) Then m_xTypes.Remove nm: PreLoadType = False
    Else
        If t_index > -1 Then x0.index = t_index
    End If
End If
Exit Function

erfun:
DebugPrint "xTypelib.PreLoadType", GetTypeName(cls), Err.Description
Err.Clear
Debug.Assert False
'Resume Next
End Function

'Function RefFindTypeByName(ByVal nm$) As ITypeInfo
'On Error Resume Next
'Dim ref As xTypeLib
'Dim cls As ITypeInfo, xt As xType
'Dim fnd$
'fnd = SplitIndex(nm, ".")
'For Each ref In refs
'Set xt = ref.Types(fnd)
''Set cls = ref.Types(fnd).ITypeInfo
'If Not xt Is Nothing Then Exit For
'Next
'Set RefFindTypeByName = xt.ITypeInfo
''Debug.Print FindTypeByName.Name
'Err.Clear
'
'End Function

Friend Function FindTypeByName(ByVal nm$) As xType
On Error Resume Next
Dim ref As xTypeLib
Dim fnd$
fnd = SplitIndex(nm, ".")
For Each ref In refs
Set FindTypeByName = ref.Types(fnd)
If Not FindTypeByName Is Nothing Then Exit For
Next
Err.Clear
End Function

'Friend Function FindTypeByName(ByVal nm$) As xType
'On Error Resume Next
'Set FindTypeByName = m_xTypes(SplitIndex(nm, "."))
''Debug.Print FindTypeByName.Name
'Err.Clear
'End Function
Public Property Get VBMode() As Boolean
VBMode = m_vbmode
End Property
Public Property Let VBMode(v As Boolean)
m_vbmode = v
End Property

Public Property Get Name() As String
Name = m_Name
End Property
Public Property Let Name(s As String)
m_Name = s
End Property

Public Property Get DocString() As String
DocString = m_DocString
End Property
Public Property Let DocString(s As String)
m_DocString = s
End Property


Function CPSource(Optional cpar)
On Error GoTo err_fun
Dim ar, i&, n&, v, nm$, v0, nm0$
Dim xt As xType
Dim xr As xTypeLib
If VarType(cpar) <> vbError Then
If ArrayDims(cpar) = 1 Then 'Let CPSOURCE
    If m_ctlb Is Nothing And Not m_isReference Then Exit Function
    For i = 0 To UBound(cpar) Step 2
    nm = cpar(i): v = cpar(i + 1)
    n = aVal(GetCParam(v, "tp_"))
    If VarType(v) And vbArray Then
        Select Case True
        Case n > 0 And n < 10 'Type 1..9
            Set xt = New xType
            'Debug.Print "LET xType", nm
            v(-1) = nm
            xt.CPSource v
            m_xTypes.Add xt, nm
        Case n = 10 'TypeLib
            NewCollection refs
            refs.Add Me, nm
            'Debug.Print "LET xTypeLIb", nm
            v0 = v
            nm0 = nm
        Case n = 11  'Reference
            Set xr = New xTypeLib
            xr.CPSource v
            refs.Add xr, nm
            'Debug.Print "LET xTypeLibReference", nm
        End Select
    
    Else 'TypeLib Attributes
        Select Case nm
        Case "tp_": m_isReference = (v = 11)
        Case "uid_": CLSIDFromString "" & v, la.IID
        Case "ver_": ar = Split(v & "", "."): la.wMajorVerNum = CInt(ar(0)): la.wMinorVerNum = CInt(ar(1))
        Case "lcid_": la.lcid = v
        Case "flg_": la.wLibFlags = v
        Case "ds_": m_DocString = v
        Case "hc_": m_HelpContext = v
        Case "hf_": m_HelpFile = v
        End Select
    End If
    Next
    If Len(nm0) Then
        m_Name = nm0
        CPSource v0
    End If
End If

Else 'Get CPSOURCE
    If m_tlb Is Nothing Then Exit Function
    n = StringFromCLSID(la.IID): nm = SysAllocString(n): CoTaskMemFree n
    ar = Array(Null, "tp_", IIf(m_isReference, 11, 10), "uid_", nm, "ver_", la.wMajorVerNum & "." & la.wMinorVerNum, "lcid_", la.lcid, "flg_", la.wLibFlags, "ds_", m_DocString, "hc_", m_HelpContext, "hf_", m_HelpFile, "ic_", CLng(IIf(m_isReference, 11, 10)))
    If m_xTypes.Count Then
        n = UBound(ar) + 1
        ReDim Preserve ar(UBound(ar) + m_xTypes.Count * 2)
        For i = 1 To m_xTypes.Count
            ar(n + (i - 1) * 2) = m_xTypes(i).Name
            ar(n + (i - 1) * 2 + 1) = m_xTypes(i).CPSource
        Next
    End If
    ReDim Preserve ar(-1 To UBound(ar) - 1)
    If Not m_isReference Then
        ar = Array(Null, m_Name, ar)
        ReDim Preserve ar(2 + 2 * (refs.Count - 1))
        For i = 2 To refs.Count
            ar(3 + (i - 2) * 2 + 0) = refs(i).Name
            ar(3 + (i - 2) * 2 + 1) = refs(i).CPSource
        Next
        ReDim Preserve ar(-1 To UBound(ar) - 1)
    End If
    CPSource = ar
End If

Exit Function
err_fun:
Debug.Print Err.Description: Err.Clear
Debug.Assert False
Resume Next
End Function

Function Save() As Boolean
On Error GoTo err_fun

If m_ctlb Is Nothing Then Exit Function
m_ctlb.SaveAllChanges

Exit Function
err_fun:
Debug.Print Err.Description
Debug.Assert False
Resume Next
End Function

Friend Function Create() As Boolean
On Error GoTo err_fun

If m_ctlb Is Nothing Then 'Exit Function
    LoadInfo LoadRegTypeLib(la.IID, la.wMajorVerNum, la.wMinorVerNum, la.lcid), True
    #If createtlb Then
        Debug.Print "+ Reference " & m_Name
    #End If
Else
    #If createtlb Then
        Debug.Print "+ TypeLib " & m_Name
    #End If

    m_ctlb.SetName m_Name
    m_ctlb.SetGuid la.IID
    m_ctlb.SetVersion la.wMajorVerNum, la.wMinorVerNum
    m_ctlb.SetLcid la.lcid
    If Len(m_DocString) Then m_ctlb.SetDocString m_DocString
    m_ctlb.SetHelpContext m_HelpContext
    If Len(m_HelpFile) Then m_ctlb.SetHelpFileName m_HelpFile
    
    
    Set m_tlb = m_ctlb
    Dim ref As xTypeLib
    For Each ref In refs
        If Not ref Is Me Then ref.Create
    Next
    Dim xt As xType
    For Each xt In m_xTypes
        If Not xt.isCreated Then xt.Create Me
    Next
    
End If
Create = True
Exit Function
err_fun:
DebugPrint "xTypeLib.Create", Err.Description: Err.Clear
Debug.Assert False
'Resume

End Function

Property Get guid() As String
Dim n&
n = StringFromCLSID(la.IID): guid = SysAllocString(n): CoTaskMemFree n
End Property
Property Let guid(s As String)
On Error Resume Next
If Len(m_FileLoad) = 0 Then CLSIDFromString "" & s, la.IID
End Property


Friend Function CreateType(tpname$) As Boolean
Dim xt As xType
On Error Resume Next
Set xt = m_xTypes(tpname)
On Error GoTo err_fun
If xt Is Nothing Then
    Debug.Assert False
Else
    If Not xt.isCreated Then
        CreateType = xt.Create(Me)
    Else
        'Debug.Assert False
    End If
End If

Exit Function
err_fun:
Debug.Print Err.Description: Err.Clear
Debug.Assert False
Resume Next
End Function

Function AddTypesFromObject(ob As Object) As Boolean
On Error GoTo erfun

If Not m_isloaded Then Exit Function
Dim oDisp As IDispatch
Set oDisp = ob
Dim cls As ITypeInfo
Set cls = oDisp.GetTypeInfo()
If cls Is Nothing Then Exit Function
AddTypesFromObject = PreLoadType(-1, cls, True)
Exit Function

erfun:
DebugPrint "xTypeLib.AddTypesFromObject", Err.Description: Err.Clear
Debug.Assert False
Resume Next
End Function

Function AddTypeFromFile(lib_pth, tp_name) As Boolean
On Error Resume Next
If Not m_isloaded Then Exit Function
Dim xtlb As New xTypeLib
xtlb.Load "" & lib_pth

PreLoadType -1, xtlb.FindTypeByName("" & tp_name).ITypeInfo, True
AddTypeFromFile = True
End Function



Public Property Get ITypeLib() As ITypeLib
Set ITypeLib = m_tlb
End Property
Public Property Get ICreateTypeLib() As ICreateTypeLib
Set ICreateTypeLib = m_ctlb
End Property
Public Property Get ICreateTypeLib2() As ICreateTypeLib2
Set ICreateTypeLib2 = m_ctlb
End Property

Friend Sub LoadInfo(tlb As ITypeLib, Optional asReference As Boolean = True)
On Error GoTo err_fun
m_isReference = asReference
Set m_tlb = tlb
m_tlb.GetDocumentation DISPID_UNKNOWN, m_Name, m_DocString, m_HelpContext, m_HelpFile
Dim n&
n = InStr(m_Name, Chr(0))
If n > 0 Then m_Name = Left(m_Name, n)
n = InStr(m_DocString, Chr(0))
If n > 0 Then m_DocString = Left(m_DocString, n)

Dim pla As Long
pla = m_tlb.GetLibAttr
CopyMemory la, ByVal pla, LenB(la)
m_tlb.ReleaseTLibAttr pla
NewCollection refs
refs.Add Me
m_isloaded = Not m_tlb Is Nothing
Exit Sub

err_fun:
DebugPrint "xTypelib.LoadInfo " & Err.Description, Err.Source
Err.Clear
Debug.Assert False
Resume Next
End Sub

'Private Sub Class_Terminate()
'Debug.Print "#TYPELIB TERMINATE ", m_Name
''Debug.Assert False
'End Sub

Friend Function GetTypeName(cls As ITypeInfo) As String
'On Error GoTo err_fun
Dim tlb As ITypeLib
Dim ref As xTypeLib
Dim bFound As Boolean
Dim cls_name$, tlb_name$
On Error Resume Next
cls.GetDocumentation DISPID_UNKNOWN, cls_name, vbNullString, 0, vbNullString

cls.GetContainingTypeLib tlb
GetTypeName = cls_name

If tlb Is Nothing Then
    'Debug.Assert False
Else
    If Not m_tlb Is tlb Then
    tlb.GetDocumentation DISPID_UNKNOWN, tlb_name, vbNullString, 0, vbNullString
    GetTypeName = tlb_name & "." & GetTypeName
    End If
End If

Exit Function
'err_fun:
'Debug.Print "xTypelib.GetTypeName", GetTypeName, Err.Description: Err.Clear
'Debug.Assert False
'Resume Next

End Function

Public Property Get Types(Optional ByVal type_name)
On Error Resume Next
type_name = Nz(type_name)
If Len(type_name) Then
    Set Types = m_xTypes(type_name)
Else
    Set Types = m_xTypes
End If
End Property

Friend Function GetTypeInfo(target$) As ITypeInfo
Dim info As ITypeInfo, nm, c%
Dim hash&, memid&, fnd$
nm = Split(target & ".", ".")
If Len(nm(1)) = 0 Then nm(1) = target: nm(0) = m_Name
Dim ref As xTypeLib
For Each ref In refs
    If ref.Name = nm(0) Then
        c = 1
        ref.ITypeLib.FindName ByVal nm(1), hash, info, memid, c
        If Not info Is Nothing Then Exit For
    End If
Next
If Not info Is Nothing Then Set GetTypeInfo = info
End Function

Friend Function GetTypeDesc(ct As ICreateTypeInfo, vts, base$, td() As ARRAYDESC, tdi&) As ARRAYDESC
On Error GoTo erfun
Dim ar
Dim plv&, vt As Integer
Dim vtar, i&, b
Dim ref As ITypeInfo, href&
ReDim Preserve td(tdi + UBound(vts) + 0) 'As TYPEDESC LEVEL POINTERS

For plv = tdi To tdi + UBound(vts)
    vtar = Split(vts(plv - tdi), ";")
    vt = vtar(0)
    If UBound(vtar) > 0 Then 'CARRAY
        td(plv).tdescElem.vt = vtar(0)
        td(plv).cDims = UBound(vtar)
        Debug.Assert UBound(vtar) < 2
        For i = 0 To td(plv).cDims - 1
            b = Split(vtar(i + 1), ":")
            td(plv).rgbounds(i).lLbound = b(0)
            td(plv).rgbounds(i).cElements = b(1)
        Next
    Else
        td(plv).tdescElem.vt = vt
    End If
    Select Case vt
    Case VT_PTR, VT_SAFEARRAY
        td(plv).tdescElem.pTypeDesc = -(plv + 1)
    Case VT_CARRAY
        'Debug.Assert False
        td(plv).tdescElem.pTypeDesc = -(plv + 1)
    Case VT_USERDEFINED
        href = 0
        Set ref = GetTypeInfo(base)
        If ref Is Nothing Then
            CreateType base
            Set ref = GetTypeInfo(base)
            If ref Is Nothing Then
                'Debug.Assert False
                #If createtlb Then
                Debug.Print "ERROR GetTypeInfo " & base
                #End If
                
                td(tdi).tdescElem.vt = VT_VOID
                td(tdi).tdescElem.pTypeDesc = 0
                GetTypeDesc = td(tdi)
                tdi = plv
                Exit Function
            End If
        End If
        ct.AddRefTypeInfo ref, href
        td(plv).tdescElem.pTypeDesc = href
    End Select
Next
GetTypeDesc = td(tdi)
tdi = plv
'On Error GoTo erfun
Exit Function

erfun:
DebugPrint "xTypeLib.GetTypeDesc", Err.Description: Err.Clear
Debug.Assert False
Resume Next

End Function


Friend Function GetTypeNameEx(cls As ITypeInfo, td0 As TYPEDESC, vts$, strbase$) As String
On Error GoTo erfun

Dim ref As ITypeInfo
Dim ref0 As ITypeInfo
Dim td As TYPEDESC
Dim ard  As ARRAYDESC
Dim res$, nm$
Dim ar

td = td0
If Len(vts) Then vts = vts & ","
vts = vts & td.vt
Do
Select Case td.vt
Case VT_PTR: 'Pointer
    CopyMemory td, ByVal td0.pTypeDesc, LenB(td)
    res = GetTypeNameEx(cls, td, vts, strbase)
    If m_vbmode Then
    'strbase = strbase & "*"
    Else
    strbase = strbase & "*"
    End If
Case VT_SAFEARRAY: 'Dynamic array
    'Debug.Assert False
    CopyMemory td, ByVal td0.pTypeDesc, LenB(td)
    res = GetTypeNameEx(cls, td, vts, strbase)
    If m_vbmode Then
        strbase = "() " & strbase
    Else
        strbase = "SAFEARRAY(" & strbase & ")"
    End If
Case VT_CARRAY: 'Fixed array
    ''Debug.Assert False
    CopyMemory ard, ByVal td0.pTypeDesc, 20 'Len(ard)
    td = ard.tdescElem
    res = GetTypeNameEx(cls, td, vts, strbase)
    Debug.Assert ard.cDims = 1
    vts = vts & ";" & ard.rgbounds(0).lLbound & ":" & ard.rgbounds(0).cElements
    If m_vbmode Then
        strbase = " (" & ard.rgbounds(0).cElements & ") " & strbase
    Else
        strbase = strbase & " %" & "[" & ard.rgbounds(0).cElements & "]"
    End If
    
Case VT_USERDEFINED
    On Error Resume Next
    Set ref = cls.GetRefTypeInfo(td0.pTypeDesc)
    On Error GoTo erfun
    If ref Is Nothing Then
        res = "<GetRefTypeInfo failed>"
        ar = Split(vts, ","): ar(UBound(ar)) = VT_VOID: vts = Join(ar, ",")
        Debug.Assert False
        DebugPrint Err.Description, res
    Else
        res = GetTypeName(ref)
        AddReference ref
       ' PreLoadType -1, ref, True
        
    End If
Case Else
    res = InstTypeName(td.vt)

End Select
Loop While res = ""

GetTypeNameEx = res
If Len(strbase) = 0 Then strbase = res
Exit Function

erfun:
Debug.Print Err.Description: Err.Clear
Debug.Assert False
Resume Next
End Function

Friend Function InstTypeName(vt As Integer, Optional vb As Boolean)
Dim res$
Select Case vt
Case VT_HRESULT: res = IIf(vb, " ", "HRESULT")
Case VT_VOID: res = IIf(vb, "Any", "void") 'only used in void* Param, never in void* Param
Case VT_VARIANT: res = IIf(vb, "Variant", "VARIANT")
Case VT_BSTR: res = IIf(vb, "String", "BSTR")
Case VT_UI1: res = IIf(vb, "Byte", "unsigned char")
Case VT_I2: res = IIf(vb, "Integer", "short")
Case VT_I4: res = IIf(vb, "Long", "long")
Case VT_BOOL: res = IIf(vb, "Boolean", "short")
Case VT_R4: res = IIf(vb, "Single", "float")
Case VT_R8: res = IIf(vb, "Double", "double")
Case VT_DISPATCH: res = IIf(vb, "Object", "IDispatch")
Case VT_LPSTR: res = IIf(vb, "String", "LPSTR")
Case VT_LPWSTR: res = IIf(vb, "String", "LPWSTR")
Case VT_UNKNOWN: res = "IUnknown"
Case VT_I1: res = "char"
Case VT_INT: res = "int"
Case VT_UI2: res = "unsigned short"
Case VT_UI4: res = "unsigned long"
Case VT_CY: res = "Currency"
Case VT_DATE: res = "Date"
Case VT_DECIMAL: res = "Decimal"
Case VT_ERROR: res = "SCode"
Case VT_I8: res = "INT64"
Case VT_UI8: res = "UINT64"
Case VT_UINT: res = "UINT"
Case Else
    res = "Unresolved VarType"
    Debug.Assert False
End Select
InstTypeName = res
End Function

Friend Sub AddReference(cls As ITypeInfo)  'As Boolean 'True=MyType False=Reference Added
Dim tlb As ITypeLib, bFound As Boolean
Dim index&

On Error Resume Next
cls.GetContainingTypeLib tlb
On Error GoTo erfun
If tlb Is Nothing Then Set tlb = m_tlb 'refs(1).ITypeLib
Dim ref As xTypeLib
If m_tlb Is tlb Then 'This Type in My TypeLib
    PreLoadType -1, cls, False
Else ' Outer reference
    'For index = 2 To refs.Count 'UBound(refs)
    '    If refs(index).ITypeLib Is tlb Then bFound = True: Exit For
    'Next
    For Each ref In refs
        If ref.ITypeLib Is tlb Then bFound = True: Exit For
    Next
    
    If Not bFound Then 'Add to references
        'Dim ref_tlb As New xTypeLib
        Set ref = New xTypeLib
        'ref_tlb.LoadInfo tlb
        ref.LoadInfo tlb
        'refs.Add ref_tlb
        refs.Add ref
    Else
       ' ref.PreLoadType -1, cls, False
        'Debug.Print "LOAD REF ", ref.Name
    End If
End If



Exit Sub
erfun:
Debug.Print Err.Description: Err.Clear
Debug.Assert False
Resume Next
End Sub

Friend Function GetVarName(tp_name, var_value)
GetVarName = var_value
Dim xt As xType
Dim xv As xVarDesc
On Error Resume Next
Set xt = m_xTypes(tp_name)
If xt Is Nothing Then Exit Function
For Each xv In xt.Vars
    If xv.Icon = 43 Then
        If xv.Value(var_value) = var_value Then GetVarName = xv.Name: Exit For
    End If
Next
End Function

Public Function GetEscapedRows(Optional strRootName As String) As String
On Error GoTo erfun
If Not m_vbmode Then Exit Function

Dim res$, s$
Dim xt As xType
Dim tk
Dim hash&
hash = Timer
'src=join(split("ID$,PID$,NAME$,ICON&,TXT$,VAL$,HELP$",","),vbTab)

If Len(strRootName) = 0 Then res = Join(Array(m_Name, "", m_Name, 10, "", "", Escape(m_DocString)), vbTab)
For Each tk In Array(TKIND_ALIAS, TKIND_COCLASS, TKIND_INTERFACE, TKIND_DISPATCH, TKIND_MODULE, TKIND_ENUM, TKIND_RECORD, TKIND_UNION)
    For Each xt In m_xTypes
        If xt.Kind = tk Then
            s = s & xt.GetEscapedRows(Me, IIf(Len(strRootName), Trim$(strRootName), m_Name), xt.Name, hash, &H1100)
        End If
    Next
Next
'Debug.Print res
If Len(s) Then res = res & s Else res = ""
GetEscapedRows = res
'On Error GoTo erfun
Exit Function
erfun:
Debug.Print Err.Description: Err.Clear
Debug.Assert False
Resume Next

End Function

Function RegTLBFileLen(ByVal guid_ver_lcid) As Long
On Error Resume Next
Dim ar, uu As UUID
ar = Split(guid_ver_lcid, "/")
If InStr(ar(0), "{") = 0 Then ar(0) = "{" & ar(0) & "}"
CLSIDFromString ar(0), uu
Dim v, v0 As Integer, v1 As Integer, v2&
v = Split(ar(1), ".")
v0 = v(0): v1 = v(1)
v2 = ar(2)
Dim file$
file = QueryPathOfRegTypeLib(uu, v0, v1, v2)
RegTLBFileLen = TLBFileLen(file)
End Function

Function ListRegTLB(Optional ByVal sar_references$, Optional ByVal returnFileList As Boolean)
ListRegTLB = pRegTypeLibs(sar_references, returnFileList)
End Function

Function GetTLBInfo(ByVal tlb_pth)
tlb_pth = Nz(tlb_pth)
GetTLBInfo = Null
Dim fn&
fn = TLBFileLen("" & tlb_pth)
If fn = 0 Then Exit Function
On Error GoTo err_fun:
Dim tlb As ITypeLib
Set tlb = LoadTypeLib("" & tlb_pth)
Dim str$, Ptr&, pla0&, la0 As TLIBATTR, nm$, ar
pla0 = tlb.GetLibAttr: CopyMemory la0, ByVal pla0, LenB(la0): tlb.ReleaseTLibAttr pla0
Ptr = StringFromCLSID(la0.IID): str = SysAllocString(Ptr): CoTaskMemFree Ptr
tlb.GetDocumentation DISPID_UNKNOWN, vbNullString, nm, 0, vbNullString
ar = Array(str & "/" & la0.wMajorVerNum & "." & la0.wMinorVerNum & "/" & la0.lcid, nm, tlb_pth, str, la0.wMajorVerNum & "." & la0.wMinorVerNum, la0.lcid, la.wLibFlags, fn, 999)
GetTLBInfo = ar
Register tlb_pth
Exit Function
err_fun:
End Function

Function LoadRegTLB(ByVal guid_ver_lcid) As Boolean
On Error Resume Next
Dim ar, uu As UUID
ar = Split(guid_ver_lcid, "/")
If InStr(ar(0), "{") = 0 Then ar(0) = "{" & ar(0) & "}"
CLSIDFromString ar(0), uu
Dim v, v0 As Integer, v1 As Integer, v2&
v = Split(ar(1), ".")
v0 = v(0): v1 = v(1): v2 = v(2)
LoadInfo LoadRegTypeLib(uu, v0, v1, v2), False
LoadTypes
LoadRegTLB = True
End Function



Function UnRegister(ByVal guid_ver_lcid)
On Error Resume Next
Dim ar, uu As UUID
ar = Split(guid_ver_lcid, "/")
If InStr(ar(0), "{") = 0 Then ar(0) = "{" & ar(0) & "}"
CLSIDFromString ar(0), uu
Dim v, v0 As Integer, v1 As Integer, v2&
v = Split(ar(1), ".")
v0 = v(0): v1 = v(1): v2 = v(2)
UnRegisterTypeLib uu, v0, v1, v2, SYS_WIN32
UnRegister = Err
If Err Then Debug.Print Err.Description, "&H" & Hex(Err)
End Function

Function Register(ByVal tlb_pth) As Boolean
tlb_pth = Nz(tlb_pth)
If TLBFileLen("" & tlb_pth) = 0 Then Exit Function
On Error Resume Next
Dim tlb As ITypeLib
Set tlb = LoadTypeLib("" & tlb_pth)
RegisterTypeLib tlb, tlb_pth, vbNullString
Register = (Err = 0)
End Function



Sub UnionTypes(dest, src, mbofset)
Dim xtd As xType
Dim xts As xType
Dim nm$
Dim b As Boolean

On Error Resume Next
Set xtd = m_xTypes("" & dest)
Set xts = FindTypeByName("" & src)


'If Not b Then b = (xtd.Kind = xts.Kind) And Not (xtd.Kind = TKIND_ALIAS Or xtd.Kind = TKIND_RECORD Or xtd.Kind = TKIND_UNION)
If Not b Then b = (xtd.Kind = TKIND_DISPATCH) And (xts.Kind = TKIND_INTERFACE)
If Not b Then b = (xtd.Kind = TKIND_INTERFACE) And (xts.Kind = TKIND_DISPATCH)
If Not b Then b = (xtd.Kind = TKIND_MODULE) And (xts.Kind = TKIND_MODULE)
If Not b Then b = (xtd.Kind = TKIND_MODULE) And (xts.Kind = TKIND_ENUM)
If b Then

On Error GoTo err_sub
'Dim xi As xImplType
'For Each xi In xts.ImlTypes
'xtd.ImlTypes("").Add xi, xi.Name
'Next

Dim xf As xFuncDesc
For Each xf In xts.Funcs
xf.MemberId = xf.MemberId + mbofset
xtd.Funcs("").Add xf, xf.Name
Next

Dim xv As xVarDesc
For Each xv In xts.Vars
xv.MemberId = xv.MemberId + mbofset

xtd.Vars("").Add xv, xv.Name
Next

On Error Resume Next
m_xTypes.Remove src

End If
Exit Sub
err_sub:
Debug.Print Err.Description
Debug.Assert False
Resume Next
End Sub

Sub Free()
pCleanUp
End Sub

Sub ExtractXVBTLB()
Dim n&, ar() As Byte, a, s$, n1&
On Error Resume Next
a = "XVBEngineHelp.tlb"
s = App.Path & "\" & a
n = FileLen(s)
ar = LoadResData(a, "CUSTOM")
n1 = UBound(ar)
If n1 > 10 Then
If n1 + 1 <> n Then
    Kill s
    n = FreeFile()
    Open s For Binary As n
    Put n, 1, ar
    Close n
    Register s
End If
End If
Err.Clear
End Sub

Public Property Get DebugMsg() As Boolean
DebugMsg = gDebugMsg
End Property
Public Property Let DebugMsg(v As Boolean)
gDebugMsg = v
End Property

