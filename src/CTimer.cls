VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private iInterval As Long
Private ID As Long

'Private hWnd_hEvent&
'Public hWnd&
'Public Title$
'Public Progress$
'Public Param&
'Public Message&
'Public TicksElapsed&
'Public TicksEnd&

' User can attach any Variant data they want to the timer
'Public Item As Variant

Public Event Timer()

' SubTimer is independent of VBCore, so it hard codes error handling

'Public Enum EErrorTimer
'    eeBaseTimer = 13650     ' CTimer
'    eeTooManyTimers         ' No more than 10 timers allowed per class
'    eeCantCreateTimer       ' Can't create system timer
'End Enum

'Friend Sub ErrRaise(e As Long)
'    Dim sText As String, sSource As String
'    If e > 1000 Then
'        sSource = App.EXEName & ".WindowProc"
'        Select Case e
'        Case eeTooManyTimers
'            sText = "No more than 10 timers allowed per class"
'        Case eeCantCreateTimer
'            sText = "Can't create system timer"
'        End Select
'        Err.Raise e Or vbObjectError, sSource, sText
'    Else
'        ' Raise standard Visual Basic error
'        Err.Raise e, sSource
'    End If
'End Sub


Property Get Interval() As Long
    Interval = iInterval
End Property

' Can't just change interval--you must kill timer and start a new one
Property Let Interval(iIntervalA As Long)
    Dim f As Boolean
    If iIntervalA > 0 Then
        ' Don't mess with it if interval is the same
        If iInterval = iIntervalA Then Exit Property
        ' Must destroy any existing timer to change interval
        If iInterval Then
            f = TimerDestroy(Me)
            Debug.Assert f       ' Shouldn't fail
        End If
        ' Create new timer with new interval
        iInterval = iIntervalA
        TimerCreate Me 'If TimerCreate(Me) = False Then ErrRaise eeCantCreateTimer
    Else
        If (iInterval > 0) Then
            iInterval = 0
            f = TimerDestroy(Me)
            Debug.Assert f      ' Shouldn't fail
        End If
    End If
End Property

' Must be public so that Timer object can't terminate while client's ThatTime
' event is being processed--Friend wouldn't prevent this disaster
Public Sub PulseTimer()
Attribute PulseTimer.VB_MemberFlags = "40"
    RaiseEvent Timer
End Sub

Friend Property Get TimerID() As Long
    TimerID = ID
End Property

Friend Property Let TimerID(idA As Long)
    ID = idA
End Property

Private Sub Class_Terminate()
    Interval = 0
End Sub

