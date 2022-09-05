VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEventObjectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 512
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************
'
' ObjectInfo Class - EventCollection Library
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

Private m_oObject As IUnknown
Private m_sKey As String
Attribute m_sKey.VB_VarHelpID = 522
Private m_lCookie As Long
Attribute m_lCookie.VB_VarHelpID = 521
Private m_lIndex As Long
Attribute m_lIndex.VB_VarHelpID = 520
Private m_tIID As UUID
Attribute m_tIID.VB_VarHelpID = 519
Private m_Events As String
Private m_vbMacros As String
Private m_AllowAllEvents As Boolean

Friend Property Get IID() As UUID
Attribute IID.VB_HelpID = 518
   IID = m_tIID
End Property

Friend Property Get Cookie() As Long
Attribute Cookie.VB_HelpID = 517
   Cookie = m_lCookie
End Property

Public Property Get key() As String
   key = m_sKey
End Property

Public Property Get Object() As Object
   Set Object = m_oObject
End Property

Public Property Get index() As Long
   index = m_lIndex
End Property
Public Property Get Events() As String
   Events = m_Events
End Property
Public Property Get VBMacros() As String
   VBMacros = m_vbMacros
End Property

Friend Sub frInitialize( _
   ByVal key As String, _
   ByVal index As Long, _
   ByVal Cookie As Long, _
   ByVal Object As Object, _
   IID As UUID, _
   sEvents As String, svbMacros As String)

   m_sKey = key
   m_lIndex = index
   m_lCookie = Cookie
   Set m_oObject = Object
   m_tIID = IID
   m_Events = sEvents
   m_vbMacros = svbMacros
   m_AllowAllEvents = InStr(m_Events, ",-,") > 0
   
End Sub

Function AllowDISPID(ByVal dispidmember As Long) As Boolean
If m_AllowAllEvents Then
    AllowDISPID = True
Else
    AllowDISPID = InStr("," & m_Events & ",", "," & dispidmember & ",") > 0
End If

End Function
