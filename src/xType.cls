VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_index As Long
Private m_Name As String
Private m_DocString As String
Private m_HelpContext As Long

Private asObject As Boolean

Private ta As TYPEATTR

Private m_base As String
Private m_vtsar As String
Private m_idlbase As String

Private m_typekindalias As Long

Private m_ImplTypes As New Collection
Private m_Funcs As New Collection
Private m_Vars As New Collection

Private m_cls As ITypeInfo
Private ct  As ICreateTypeInfo2
Private m_isCreated As Boolean
Private mhash As Long
Private parent As xTypeLib

Public Property Get Kind() As TYPEKIND
Kind = ta.TYPEKIND
End Property

Sub ConvertToModule()
'If ta.TYPEKIND = TKIND_INTERFACE Then ta.TYPEKIND = TKIND_MODULE Else Exit Sub
If ta.TYPEKIND = TKIND_INTERFACE Or ta.TYPEKIND = TKIND_DISPATCH Then ta.TYPEKIND = TKIND_MODULE Else Exit Sub

Dim xf As xFuncDesc
For Each xf In m_Funcs
xf.ConvertToModule
Next
End Sub

Public Sub Layout()
'ta.TYPEKIND = 2
ta.wTypeFlags = 0
Dim xf As xFuncDesc
For Each xf In m_Funcs
xf.Layout
Next
End Sub

Public Property Get ImlTypes(Optional ByVal memb_name)
On Error Resume Next
memb_name = Nz(memb_name)
If Len(memb_name) Then
    Set ImlTypes = m_ImplTypes(memb_name)
Else
    Set ImlTypes = m_ImplTypes
End If
End Property
Public Property Get Funcs(Optional ByVal memb_name)
On Error Resume Next
memb_name = Nz(memb_name)
If Len(memb_name) Then
    Set Funcs = m_Funcs(memb_name)
Else
    Set Funcs = m_Funcs
End If
End Property
Public Property Get Vars(Optional ByVal memb_name)
On Error Resume Next
memb_name = Nz(memb_name)
If Len(memb_name) Then
    Set Vars = m_Vars(memb_name)
Else
    Set Vars = m_Vars
End If
End Property

Friend Function Load(ByVal tp_index&, cls As ITypeInfo, xtlb As xTypeLib, isob As Boolean) As Boolean
On Error GoTo erfun
m_index = tp_index
Dim pta&
Set parent = xtlb

pta = cls.GetTypeAttr: CopyMemory ta, ByVal pta, Len(ta)

If (ta.wTypeFlags And TYPEFLAG_FDUAL) And ta.TYPEKIND = TKIND_DISPATCH Then
#If LoadTLB Then
    Debug.Print "DUAL >>> ";
#End If
    cls.ReleaseTypeAttr pta
    Set cls = cls.GetRefTypeInfo(cls.GetRefTypeOfImplType(-1))
    pta = cls.GetTypeAttr: CopyMemory ta, ByVal pta, Len(ta)
End If
'ta.wTypeFlags = ta.wTypeFlags And (Not TYPEFLAG_FHIDDEN)
cls.GetDocumentation DISPID_UNKNOWN, m_Name, m_DocString, m_HelpContext, vbNullString
Set m_cls = cls
Dim n&
'n = InStr(m_Name, Chr(0))
'If n > 0 Then m_Name = Left(m_Name, n)
n = InStr(m_DocString, Chr(0))
If n > 0 Then m_DocString = Left(m_DocString, n)


If Len(m_Name) = 0 Then
    m_Name = "<TK_" & ta.TYPEKIND & "_" & tp_index & "_Error_GetName_" & xtlb.Name & ">"
End If
#If LoadTLB Then
    Debug.Print m_Name;
#End If

asObject = Not TypeLibExist(cls)
If asObject Or isob Then
    asObject = True
    ta.wTypeFlags = ta.wTypeFlags And Not (TYPEFLAG_FCANCREATE)
    If ta.TYPEKIND = TKIND_INTERFACE Or ta.TYPEKIND = TKIND_DISPATCH Then CoCreateGuid ta.IID
    m_Name = Replace(m_Name, "_", "")
End If
'Debug.Assert m_Name <> "Range"

Dim index&

If ta.TYPEKIND = TKIND_ALIAS Then
        m_typekindalias = -1
        m_base = xtlb.GetTypeNameEx(cls, ta.tdescAlias, m_vtsar, m_idlbase)
        m_typekindalias = -aVal(SplitIndex(m_vtsar, ","))
        If m_typekindalias = -VT_USERDEFINED Then
            m_typekindalias = pGetTypeKindAlias(cls, ta.tdescAlias)
        End If
        m_typekindalias = 1000 + m_typekindalias + 1
        #If LoadTLB Then
            Debug.Print " = " & m_base
        #End If
End If


Dim href&, ref As ITypeInfo
If ta.TYPEKIND = TKIND_COCLASS Then
    #If LoadTLB Then
        Debug.Print
    #End If
   
    Dim xi As xImplType
    Set m_ImplTypes = New Collection
    For index = 0 To ta.cImplTypes - 1
        Set xi = New xImplType
        xi.Load index, cls, xtlb
        m_ImplTypes.Add xi
    Next
    
ElseIf ta.cImplTypes Then
    href = cls.GetRefTypeOfImplType(0)
    Set ref = cls.GetRefTypeInfo(href)
    If ref Is Nothing Then
        m_base = "<GetRefTypeInfo failed>" ' & i
        #If LoadTLB Then
            Debug.Print " as " & m_base
        #End If
    Else
        m_base = xtlb.GetTypeName(ref)
        #If LoadTLB Then
            Debug.Print " as " & m_base
        #End If
        'Debug.Assert m_base <> "ITextRange"
        xtlb.AddReference ref
    End If
Else
    #If LoadTLB Then
        Debug.Print
    #End If
End If


If ta.cVars > 0 Then
    If xtlb.bSAF("41,42,43,44", 2) Then
    
        Dim xv As xVarDesc
        Set m_Vars = New Collection
        For index = 0 To ta.cVars - 1
            Set xv = New xVarDesc
            xv.Load index, cls, xtlb
            m_Vars.Add xv
        Next
    
    End If
End If
'Debug.Assert m_Name <> "ITextRange"
'Debug.Assert m_Name <> "Range"

Dim xf As xFuncDesc
Dim xf0 As xFuncDesc
Dim nm$

Set m_Funcs = New Collection
For index = 0 To ta.cFuncs - 1
    Set xf = New xFuncDesc

    xf.Load index, ta.TYPEKIND, cls, xtlb, asObject
    
    
    Set xf0 = FindFuncDescByName(xf.Name)
    If xf0 Is Nothing Then
        m_Funcs.Add xf, xf.Name
    Else
    'Debug.Assert Not (m_Name = "Range" And xf.Name = "_Default")
        xf0.AppendFunc index, xf
    End If
    
Next

cls.ReleaseTypeAttr pta
Load = True

Exit Function
erfun:
Debug.Print "xType.Load", Err.Description, xtlb.GetTypeName(cls)
'Err.Clear
Debug.Assert False
'Resume Next
End Function

Private Function FindFuncDescByName(ByVal nm$) As xFuncDesc
On Error Resume Next
Set FindFuncDescByName = m_Funcs(nm)
Err.Clear
End Function

Public Property Get Name() As String
Name = m_Name
End Property
Public Property Let Name(v As String)
If parent Is Nothing Then Exit Property
If parent.FindTypeByName(v) Is Nothing Then m_Name = v
End Property

Public Property Get ITypeInfo() As ITypeInfo
Set ITypeInfo = m_cls
End Property

Friend Property Get index() As Long
index = m_index
End Property
Friend Property Let index(v As Long)
m_index = v
End Property
Friend Property Get isCreated() As Boolean
isCreated = m_isCreated
End Property

Private Function pGetTypeKindAlias(cls As ITypeInfo, td0 As TYPEDESC) As TYPEKIND
On Error GoTo erfun
Dim ta As TYPEATTR, pta&
Dim ref As ITypeInfo
Dim td As TYPEDESC
td = td0
Do
Select Case td.vt
Case VT_PTR
    CopyMemory td, ByVal td.pTypeDesc, LenB(td)
Case VT_USERDEFINED
    On Error Resume Next
    Set ref = cls.GetRefTypeInfo(ByVal td.pTypeDesc)
    On Error GoTo erfun
    If ref Is Nothing Then
        pGetTypeKindAlias = 6
    Else
        pta = ref.GetTypeAttr: CopyMemory ta, ByVal pta, LenB(ta)
        If ta.TYPEKIND = TKIND_ALIAS Then
            pGetTypeKindAlias = pGetTypeKindAlias(ref, ta.tdescAlias)
        Else
            pGetTypeKindAlias = ta.TYPEKIND
        End If
        ref.ReleaseTypeAttr pta
    End If
    Exit Do
Case VT_SAFEARRAY
    Debug.Assert False
    Exit Do
Case VT_CARRAY
    Debug.Assert False
    Exit Do
Case Else
    Debug.Assert False
    Exit Do
End Select

Loop


Exit Function
erfun:
Debug.Print Err.Description: Err.Clear
Debug.Assert False
Resume Next
End Function




Function CPSource(Optional cpar)
On Error GoTo err_fun

Dim ar, i&, nm$, v, n&
If VarType(cpar) <> vbError Then
If ArrayDims(cpar) = 1 Then 'Let

    Dim xf As xFuncDesc
    Dim xi As xImplType
    Dim xv As xVarDesc
    NewCollection m_ImplTypes
    NewCollection m_Funcs
    NewCollection m_Vars
    m_Name = cpar(-1)
    For i = 0 To UBound(cpar) Step 2
    nm = cpar(i): v = cpar(i + 1)
    n = aVal(GetCParam(v, "tp_"))
    If VarType(v) And vbArray Then
        v(-1) = nm
        Select Case True
        Case n > 11 And n < 20 'Implied Interface
            Set xi = New xImplType
            m_ImplTypes.Add xi, nm
            Debug.Print , "LET xImplType", nm
            xi.CPSource v
        Case n > 19 And n < 40 'FuncDesc
            Set xf = New xFuncDesc
            m_Funcs.Add xf, nm
            Debug.Print , "LET xFuncDesc", nm
            xf.CPSource v
        Case n > 40 And n < 45 'VarDesc
            Set xv = New xVarDesc
            Debug.Print , "LET xVarDesc", nm
            m_Vars.Add xv, nm
            xv.CPSource v
        End Select
        
    Else 'TypeLib Attributes
        Select Case nm
        Case "tp_": ta.TYPEKIND = (v - 1)
        Case "uid_": CLSIDFromString "" & v, ta.IID
        Case "ver_": ar = Split(v & "", "."): ta.wMajorVerNum = CInt(ar(0)): ta.wMinorVerNum = CInt(ar(1))
        Case "bas_": m_base = v
        Case "vt_": m_vtsar = v
        Case "flg_": ta.wTypeFlags = v
        Case "ob_": If v Then ta.wTypeFlags = ta.wTypeFlags And (Not (TYPEFLAG_FCANCREATE Or TYPEFLAG_FHIDDEN Or TYPEFLAG_FAPPOBJECT))
        Case "ds_": m_DocString = v
        Case "hc_": m_HelpContext = v
        End Select
    End If
    Next
    
End If

Else 'Get
n = StringFromCLSID(ta.IID): nm = SysAllocString(n): CoTaskMemFree n

Dim tk&
tk = 1 + ta.TYPEKIND
ar = Array(Null, "tp_", tk, "uid_", nm, "ver_", ta.wMajorVerNum & "." & ta.wMinorVerNum, "flg_", ta.wTypeFlags, "ds_", m_DocString, "hc_", m_HelpContext, "bas_", "", "vt_", "", "ic_", tk, "baz_", m_idlbase, "ob_", asObject, "ix_", m_index)
If (ta.cImplTypes > 0 And ta.TYPEKIND <> TKIND_COCLASS) Or (ta.TYPEKIND = TKIND_ALIAS) Then
    ar(14) = m_base
    ar(16) = m_vtsar
    If ta.TYPEKIND = TKIND_ALIAS Then ar(18) = m_typekindalias 'ICON
End If

i = m_Vars.Count + m_Funcs.Count + m_ImplTypes.Count
If i > 0 Then
    n = UBound(ar) + 1
    ReDim Preserve ar(UBound(ar) + i * 2)

    If m_ImplTypes.Count Then
        For i = 1 To m_ImplTypes.Count
            ar(n + (i - 1) * 2 + 0) = m_ImplTypes(i).Name
            ar(n + (i - 1) * 2 + 1) = m_ImplTypes(i).CPSource
        Next
        n = n + m_ImplTypes.Count * 2
    End If
    
    If m_Vars.Count Then
        For i = 1 To m_Vars.Count
            ar(n + (i - 1) * 2 + 0) = m_Vars(i).Name
            ar(n + (i - 1) * 2 + 1) = m_Vars(i).CPSource
        Next
        n = n + m_Vars.Count * 2
    End If
    
    If m_Funcs.Count Then
        For i = 1 To m_Funcs.Count
            ar(n + (i - 1) * 2 + 0) = m_Funcs(i).Name
            ar(n + (i - 1) * 2 + 1) = m_Funcs(i).CPSource
        Next
        n = n + m_Funcs.Count * 2
    End If

End If

ReDim Preserve ar(-1 To UBound(ar) - 1)
CPSource = ar
End If

Exit Function
err_fun:
Debug.Print Err.Description: Err.Clear
Debug.Assert False
Resume Next
End Function

Friend Function Create(ctlb As xTypeLib) As Boolean
On Error GoTo err_fun

'Debug.Assert False
#If createtlb Then
    Debug.Print "+ Type " & m_Name;
#End If

Set parent = ctlb

ctlb.ICreateTypeLib.CreateTypeInfo m_Name, ta.TYPEKIND, ct
m_isCreated = True


ct.SetTypeFlags ta.wTypeFlags

ct.SetGuid ta.IID
'ct.SetName m_Name
If Len(m_DocString) Then ct.SetDocString m_DocString
ct.SetHelpContext m_HelpContext

'If Len(m_HelpFile) Then ct.SetHelpFileName m_HelpFile




Dim xf As xFuncDesc
Dim xi As xImplType
Dim xv As xVarDesc
Dim index&
Dim i&


Dim ref As ITypeInfo, href&

'Debug.Assert m_Name <> "Range"

Select Case ta.TYPEKIND
Case TKIND_INTERFACE, TKIND_DISPATCH
    If Len(m_base) Then 'Base Interface
        Set ref = ctlb.GetTypeInfo(m_base)
        #If createtlb Then
            Debug.Print " as " & m_base
        #End If
        If ref Is Nothing Then
            ctlb.CreateType m_base
            Set ref = ctlb.GetTypeInfo(m_base)
        End If
        If ref Is Nothing Then
            #If createtlb Then
            Debug.Print "ERROR GetTypeInfo " & m_base
            #End If
            
        Else
            ct.AddRefTypeInfo ref, href
            ct.AddImplType 0, href
        End If
    Else
        #If createtlb Then
            Debug.Print
        #End If
    End If

Case TKIND_ALIAS
    Dim td() As ARRAYDESC, tdi&
    ReDim td(0)
    ctlb.GetTypeDesc ct, Split(m_vtsar, ","), "" & m_base, td, tdi
    For i = 0 To UBound(td)
        If td(i).tdescElem.pTypeDesc < 0 Then td(i).tdescElem.pTypeDesc = VarPtr(td(Abs(td(i).tdescElem.pTypeDesc)))
    Next
    ct.SetTypeDescAlias td(0).tdescElem
    #If createtlb Then
        Debug.Print " as " & m_base
    #End If
Case Else
    #If createtlb Then
        Debug.Print
    #End If
End Select


index = 0
If m_ImplTypes.Count Then
        #If createtlb Then
            Debug.Print
        #End If
    For Each xi In m_ImplTypes
        xi.Create ct, index, ctlb
    Next
End If


index = 0
For Each xv In m_Vars
    xv.Create ct, index, ctlb
Next

'Debug.Assert m_Name <> "Range"

index = 0
For Each xf In m_Funcs
    xf.Create ct, index, ta.TYPEKIND, ctlb
Next

'ct.LayOut
Create = Err
Exit Function

err_fun:
DebugPrint "xType.Create", m_Name, Err.Description
'Err.Clear
Debug.Assert False
'CoCreateGuid ta.IID
'ctlb.ICreateTypeLib2.DeleteTypeInfo m_Name
Debug.Print asObject, Hex(ta.wTypeFlags)

Resume Next
End Function


'Private Sub Class_Terminate()
'Debug.Print "#TYPE TERMINATE ", m_Name
'End Sub

Function GetEscapedRows(xtlb As xTypeLib, parent$, tp_name$, hash&, options&) As String

On Error GoTo erfun
Dim xt As xType
Dim xf As xFuncDesc
Dim xi As xImplType
Dim xv As xVarDesc
Dim res$, s$, par$
Dim it&, rv

'Debug.Print m_Name

it = options And 7
If hash = mhash And (options And &H10000) = 0 Then Exit Function

If (ta.wTypeFlags And (TYPEFLAG_FRESTRICTED Or TYPEFLAG_FRESTRICTED)) Then Exit Function

If (ta.wTypeFlags And TYPEFLAG_FHIDDEN) And (options And &H100) = 0 Then
    'Debug.Assert False
    Exit Function
End If


If ta.TYPEKIND = TKIND_ALIAS Then
    mhash = hash
    Set xt = xtlb.FindTypeByName(m_base)
    If Not xt Is Nothing Then
        s = xt.GetEscapedRows(xtlb, parent, tp_name, -1, &H11100 Or options)
        res = res & s
    End If
Else
    If (options And &H1000) = &H1000 Then 'Add Header
        par = parent & IIf(Len(parent), ".", "") & tp_name
        s = Join(Array(par, parent, tp_name, Me.Icon, "", "", Escape(m_DocString)), vbTab)
        'Debug.Print s
        res = res & vbCrLf & s
    Else
        par = parent
    End If

    If ta.TYPEKIND = TKIND_COCLASS Then
        mhash = hash
        For Each xi In m_ImplTypes 'CoClass
            Set xt = xtlb.FindTypeByName(xi.Name)
            If Not xt Is Nothing Then
                s = xt.GetEscapedRows(xtlb, par, xt.Name, hash, &H100 + xi.Icon - 12)
                'Debug.Print s
                res = res & s
            End If
        Next
    End If
    If ta.TYPEKIND = TKIND_MODULE Or ta.TYPEKIND = TKIND_INTERFACE Or ta.TYPEKIND = TKIND_DISPATCH Then
        mhash = hash
        Set xt = xtlb.FindTypeByName(m_base)
        If Not xt Is Nothing Then
            s = xt.GetEscapedRows(xtlb, par, xt.Name, hash, &H10100)
            'Debug.Print s
            res = res & s
        End If
        For Each xf In m_Funcs
            If (xf.mFuncFlags And (FUNCFLAG_FRESTRICTED Or 0 * FUNCFLAG_FHIDDEN)) = 0 Then
                s = Join(Array(par & IIf(Len(par), ".", "") & xf.Name, par, xf.Name, xf.Icon(it), xf.VBDecl(xtlb, rv, it), Escape(rv), Escape(xf.DocString)), vbTab)
                'Debug.Print s
                res = res & vbCrLf & s
            End If
        Next
    End If
    'If ta.TYPEKIND = TKIND_ENUM Or ta.TYPEKIND = TKIND_DISPATCH Or ta.TYPEKIND = TKIND_RECORD Or ta.TYPEKIND = TKIND_UNION Or ta.TYPEKIND = TKIND_MODULE Then
        mhash = hash
        For Each xv In m_Vars 'Module,Enum,Record,Union
             s = Join(Array(par & IIf(Len(par), ".", "") & xv.Name, par, xv.Name, xv.Icon, Escape(xv.VBDecl()), Escape(xv.Value(0&)), Escape(xv.DocString)), vbTab)
            'Debug.Print s
            res = res & vbCrLf & s
        Next
    'End If
End If

GetEscapedRows = res

'On Error GoTo erfun
Exit Function
erfun:
Debug.Print Err.Description: Err.Clear
Debug.Assert False
Resume Next

End Function

Public Property Get Icon() As Long
If ta.TYPEKIND = TKIND_ALIAS Then
    Icon = m_typekindalias
Else
    Icon = ta.TYPEKIND + 1
End If
End Property

