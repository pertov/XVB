VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xVarDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_index As Long
Private m_Name As String

Private m_DocString As String
Private m_HelpContext As Long

Dim vd As VARDESC

Private m_base As String
Private m_vtsar As String
Private m_idlbase As String


Private m_Value As Variant

Friend Function Load(ByVal var_index&, cls As ITypeInfo, xtlb As xTypeLib)
On Error GoTo erfun

m_index = var_index
Dim pvd As Long

pvd = cls.GetVarDesc(var_index)
CopyMemory vd, ByVal pvd, LenB(vd)
cls.GetDocumentation vd.memid, m_Name, m_DocString, m_HelpContext, vbNullString
Dim n&
'n = InStr(m_Name, Chr(0))
'If n > 0 Then m_Name = Left(m_Name, n)
n = InStr(m_DocString, Chr(0))
If n > 1 Then m_DocString = Left(m_DocString, n - 1)
    
m_base = "": m_vtsar = "": m_idlbase = ""
m_base = xtlb.GetTypeNameEx(cls, vd.elemdescVar.tdesc, m_vtsar, m_idlbase)
#If LoadTLB Then
Debug.Print , m_Name & " as " & m_base & " [" & m_vtsar & "]";
#End If

Dim dv0 As Variant
m_Value = Empty
If vd.VARKIND = VAR_CONST Then
    CopyMemory dv0, ByVal vd.oInst_varValue, 16
    m_Value = VariantTypeToVBVarType(dv0)
    FillMemory dv0, 16, 0
    m_vtsar = VarVT(m_Value)
    'If IsNumeric(dv) Then cp(nm & "\dvs_") = dv & " (&H" & Hex$(dv) & ")"
    #If LoadTLB Then
    Debug.Print " = " & Nz(m_Value)
    #End If
Else
    #If LoadTLB Then
    Debug.Print
    #End If

End If

cls.ReleaseVarDesc pvd
Exit Function
erfun:
Debug.Print Err.Description: Err.Clear
Debug.Assert False
Resume Next
End Function

Public Property Get Name() As String
Name = m_Name
End Property

Friend Property Get MemberId() As Long
MemberId = vd.memid
End Property
Friend Property Let MemberId(v As Long)
vd.memid = v
End Property


Public Property Get Icon() As Long
Icon = vd.VARKIND + 41
End Property
Public Property Get Value(Optional defvalue)
Value = Empty
If vd.VARKIND = VAR_CONST Then
On Error Resume Next
VariantChangeType Value, m_Value, 0, VarVT(defvalue)
End If
End Property
Public Property Get DocString() As String
DocString = m_DocString
End Property

Public Function VBDecl() As String
If vd.VARKIND = VAR_CONST Then VBDecl = "Const "
VBDecl = VBDecl & m_Name
If vd.VARKIND = VAR_CONST Then
    If VarVT(m_Value) = vbString Then
        VBDecl = VBDecl & " = " & Chr$(34) & m_Value & Chr$(34)
    Else
        VBDecl = VBDecl & " = " & m_Value & " '(&H" & Hex(m_Value) & ")"
    End If
'    If Len(m_DocString) Then VBDecl = VBDecl & " ' " & m_DocString
Else
    Dim pv1&, pb$
    pb = m_base
    pv1 = aVal(Right(m_vtsar, 2))
    If pv1 = 29 Then pb = "<" & pb & ">"
    VBDecl = VBDecl & " as " & pb
End If
'If Len(m_HelpString) Then VBDecl = VBDecl & "   " & Chr$(39) & m_HelpString

End Function


Function CPSource(Optional cpar)
On Error GoTo err_fun
Dim ar, vk&, i&, n&, nm$, v
If VarType(cpar) <> vbError Then
If ArrayDims(cpar) = 1 Then 'Let

    m_Name = cpar(-1)
    For i = 0 To UBound(cpar) Step 2
    nm = cpar(i): v = cpar(i + 1)
    n = aVal(GetCParam(v, "tp_"))
    If VarType(v) And vbArray Then
        Debug.Assert False
    Else 'Function Attributes
        Select Case nm
        Case "tp_": vd.VARKIND = v - 41
        Case "uid_": vd.memid = v
        Case "bas_": m_base = v
        Case "vt_": m_vtsar = v
        Case "flg_": vd.wVarFlags = v
        Case "ds_": m_DocString = v
        Case "hc_": m_HelpContext = v
        Case "dv_": m_Value = v
        End Select
    End If
    Next

End If
Else
    vk = 41 + vd.VARKIND
    ar = Array(Null, "tp_", vk, "uid_", vd.memid, "bas_", m_base, "vt_", m_vtsar, "flg_", vd.wVarFlags, "ds_", m_DocString, "hc_", m_HelpContext, "dv_", Empty, "ic_", vk, "baz_", m_idlbase, "ix_", m_index + 2000)
    
    If vd.VARKIND = VAR_CONST Then ar(14) = m_Value
    
    ReDim Preserve ar(-1 To UBound(ar) - 1)
    CPSource = ar
End If
Exit Function
err_fun:
Debug.Print Err.Description: Err.Clear
Debug.Assert False
Resume Next
End Function

Friend Function Create(ct As ICreateTypeInfo2, index&, ctlb As xTypeLib) As Boolean
On Error GoTo err_fun

Dim dv
Dim td() As ARRAYDESC, tdi&, i&

ReDim td(0)
ctlb.GetTypeDesc ct, Split(m_vtsar, ","), "" & m_base, td, tdi
For i = 0 To UBound(td)
    If td(i).tdescElem.pTypeDesc < 0 Then td(i).tdescElem.pTypeDesc = VarPtr(td(Abs(td(i).tdescElem.pTypeDesc)))
Next
vd.elemdescVar.tdesc = td(0).tdescElem

#If createtlbvars Then
    Debug.Print , "+ VarDesc " & m_Name & " AS " & m_base;
#End If

If vd.VARKIND = VAR_CONST Then
    dv = m_Value
    vd.oInst_varValue = VarPtr(dv)
    VariantChangeType dv, dv, 0, vd.elemdescVar.tdesc.vt
    #If createtlbvars Then
        Debug.Print " = " & dv
    #End If
Else
    #If createtlbvars Then
        Debug.Print
    #End If
End If

ct.AddVarDesc index, vd
If Err = 0 Then
    dv = Empty
    If Len(m_DocString) Then ct.SetVarDocString index, m_DocString
    ct.SetVarHelpContext index, m_HelpContext
    ct.SetVarName index, m_Name
    index = index + 1
End If
Err.Clear


Exit Function

err_fun:
Debug.Print Err.Description
'Err.Clear
Debug.Assert False
Resume Next
End Function

