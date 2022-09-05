VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xImplType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String

Private m_ImplFlags As Long
Private m_index As Long

Friend Function Load(ByVal it_index&, cls As ITypeInfo, xtlb As xTypeLib)
On Error GoTo erfun

Dim href&, ref As ITypeInfo
m_index = it_index
m_ImplFlags = cls.GetImplTypeFlags(it_index)
href = cls.GetRefTypeOfImplType(it_index)

On Error Resume Next
Set ref = cls.GetRefTypeInfo(href)
On Error GoTo erfun

If ref Is Nothing Then
    m_Name = "<GetRefTypeInfo failed>" ' & i
Else
    xtlb.AddReference ref ', True
    m_Name = xtlb.GetTypeName(ref)
    'If xtlb.AddReference(ref) Then
    '    'If bSAF(50) Then pLoadType ref, tlbname, -1, -1
    '    xtlb.PreLoadType -1, ref
    'End If
End If
Exit Function
erfun:
Debug.Print Err.Description: Err.Clear
Debug.Assert False
Resume Next
End Function

Public Property Get Name() As String
Name = m_Name
End Property

Public Property Get Icon(Optional implied_type&) As Long
Icon = m_ImplFlags + 12
End Property

Function CPSource(Optional cpar)
On Error GoTo err_fun
Dim ar, it&, i&, n&, nm$, v
If VarType(cpar) <> vbError Then
If ArrayDims(cpar) = 1 Then 'Let

    For i = 0 To UBound(cpar) Step 2
    nm = cpar(i): v = cpar(i + 1)
    n = aVal(GetCParam(v, "tp_"))
    If VarType(v) And vbArray Then
        Debug.Assert False
    Else 'Function Attributes
        Select Case nm
        Case "tp_": m_ImplFlags = 7 And (v - 12)
        Case "bas_": m_Name = v
        Case "flg_": m_ImplFlags = v
        End Select
    End If
    Next

End If
Else
    it = 12 + (7 And m_ImplFlags)
    ar = Array(Null, "tp_", it, "bas_", m_Name, "flg_", m_ImplFlags, "ic_", it, "baz_", m_Name, "ix_", m_index + 3000)
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
#If createtlbimp Then
Debug.Print , "<< Create ImplType " & index & " " & m_Name
#End If

Dim ref As ITypeInfo, href&

Set ref = ctlb.GetTypeInfo(m_Name)
If ref Is Nothing Then
    ctlb.CreateType m_Name
    Set ref = ctlb.GetTypeInfo(m_Name)
End If

#If createtlbimp Then
Debug.Print , " >> Create ImplType ", m_Name
#End If

If ref Is Nothing Then
    Debug.Assert False
Else
    ct.AddRefTypeInfo ref, href
    ct.AddImplType index, href
    ct.SetImplTypeFlags index, m_ImplFlags
    index = index + 1
End If


Exit Function

err_fun:
Debug.Print Err.Description
'Err.Clear
Debug.Assert False
Resume Next
End Function


