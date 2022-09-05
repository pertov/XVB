VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEventParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 1024
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************
'
' Parameters Class - EventCollection Library
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

Dim m_avParams()
Attribute m_avParams.VB_VarHelpID = 1030
Dim m_paramflags() As Long

Friend Sub frInitialize(Params(), ParamFlags() As Long)
Attribute frInitialize.VB_HelpID = 1029
   m_avParams = Params
   m_paramflags = ParamFlags
End Sub

Friend Function frGetParams() As Variant()
   frGetParams = m_avParams
End Function
Friend Function frGetParamFlags() As Long()
   frGetParamFlags = m_paramflags
End Function

'!!!!!! DEFAULT MEMBER !!!!!!
Public Property Get Item(Optional ByVal indexv)
Attribute Item.VB_UserMemId = 0
Dim index&
index = indexv
If index > count - 1 Or index < 0 Then Exit Property
Item = Cast(m_avParams(index), vbVariant)
End Property
'!!!!!! DEFAULT MEMBER !!!!!!

Public Property Let Item(Optional ByVal indexv, Value)
Dim index&
index = indexv
If index > count - 1 Or index < 0 Then Exit Property
If m_paramflags(index) And 3 Then m_avParams(index) = Value: m_paramflags(index) = m_paramflags(index) Or &H1000
End Property

Public Property Set Item(Optional ByVal indexv, Value)
Dim index&
index = indexv
If index > count - 1 Or index < 0 Then Exit Property
If m_paramflags(index) And 3 Then Set m_avParams(index) = Value: m_paramflags(index) = m_paramflags(index) Or &H1000
End Property

Public Property Get count() As Long
On Error Resume Next
count = UBound(m_avParams) + 1
End Property
'
'Public Function NewEnum() As IUnknown
'   ' Not yet implemented
'End Function
