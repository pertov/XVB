VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Public"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Event MSG()
'Public objects As New Collection

Function Run(ByVal code, Optional ByVal lang, Optional ByVal bAsync)
Attribute Run.VB_Description = "code, lang=vbscript/jscript/.."
lang = NzS(lang, "vbscript")
If L_(bAsync) Then
    xMain.DoCmd xMain, "Script", lang, "+*", code
Else
    Dim cd As New CDataObject
    xMain.Script lang, "Result", cd, "*+", S_(code)
    AssignResult Run, cd.Value
End If
End Function

Function Form(src, Optional ByVal OpenArgs, Optional ByVal bWaitResult)
Attribute Form.VB_Description = "src=keypath, openargs=/array/cparam/json/null"
Dim cd As New CDataObject
Dim hWnd&
'xMain.DebugPrint 0, "AppPath=" & xMain.AppPath & vbCrLf & "src=" & S_(src)
'xMain.DebugPrint 0, "public Form START bWaitResult=" & L_(bWaitResult)

hWnd = xMain.NewForm(S_(src), , xMain.CParam(OpenArgs, IIf(InStr(S_(OpenArgs), "{"), "json", "")), cd)
'xMain.DebugPrint 0, "public Form hwnd=" & hWnd

If L_(bWaitResult) Then MessageLoopWhileVisible hWnd
AssignResult Form, cd.Value

'xMain.DebugPrint 0, "public Form END"
End Function

Function CParam(Optional ByVal src, Optional ByVal srcType)
Attribute CParam.VB_Description = "return CParam object"
Set CParam = xMain.CParam(src, srcType)
End Function

Function CMatrix(srcType, src, Optional sSortFilterGroup, Optional RowMode = 0)
Attribute CMatrix.VB_Description = "return CMatrix object"
Set CMatrix = xMain.CMatrix(L_(srcType), src, S_(sSortFilterGroup), L_(RowMode))
End Function

Function Dispatch(obj)
Attribute Dispatch.VB_Description = "obj = cparam,fields,collection(Name,Value)"
Set Dispatch = xMain.Dispatch(obj)
End Function
'
'Property Get ThreadID() As Long
'ThreadID = App.ThreadID
'End Property

Private Sub Class_Initialize()
'MsgBox "PUBLIC  Class_Initialize START"
publicrun = 1
StartLoop
'MsgBox "PUBLIC  Class_Initialize END"
End Sub

Private Sub Class_Terminate()
'MsgBox "PUBLIC  Class_Terminate START"
publicrun = 0
EndLoop 0, 0, 0, 0
'MsgBox "PUBLIC  Class_Terminate END"
End Sub
