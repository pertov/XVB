VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEventInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 768
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************
'
' EventInfo Class - EventCollection Library
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

Private m_Sender As Object
Private m_Events As String 'Events filter for INVOKE
Private m_sKey As String 'Handler name

Private m_DISPID As Long 'Event ID
Private m_sName As String 'Event name
Private m_oParams As CEventParameters
Attribute m_oParams.VB_VarHelpID = 772
Private m_vbMacros As Variant 'Macro/Arguments

Public Property Get Sender() As Object
   Set Sender = m_Sender
End Property
Public Property Get SenderEvents() As String
   SenderEvents = m_Events
End Property
Public Property Get SenderKey() As String
   SenderKey = m_sKey
End Property
Public Property Get Name() As String
   Name = m_sName
End Property
Public Property Get VBMacros()
   VBMacros = m_vbMacros
End Property
Public Property Get DISPID() As Long
   DISPID = m_DISPID
End Property

'!!!!!! DEFAULT MEMBER !!!!!!
Public Property Get Param() As CEventParameters
Attribute Param.VB_UserMemId = 0
   Set Param = m_oParams
End Property
'!!!!!! DEFAULT MEMBER !!!!!!

Public Property Get Params()
   Params = m_oParams.frGetParams
End Property

Friend Sub frInitialize( _
   ByVal Name As String, _
   ByVal DISPID As Long, _
   Params() As Variant, _
   oSender As Object, _
   ByVal sKey As String, _
   ByVal sEvents As String, ParamFlags() As Long, var_vbMacros)
   
   m_sName = Name 'procedure name
   m_DISPID = DISPID 'procedure DISPID
   Set m_Sender = oSender
   m_Events = sEvents 'events filter for INVOKE
   m_sKey = sKey 'procedure prefix
   m_vbMacros = var_vbMacros 'Macro to run onfire
   Set m_oParams = New CEventParameters
   m_oParams.frInitialize Params, ParamFlags
   
End Sub
