VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Implements ISubClass
Private m_hWnd& 'Это я
Private Parent_hWnd& 'Окно в котором я лежу
Private Target_hWnd& 'Окно к которому я прицеплен

Private bw&, bw2&
Private bb& 'BRUSH FIX
Private m_rg& 'CURRENT REGION
Private esm As Boolean 'ENTERSIZEMOVE
Private keys As Boolean 'KEY SIZEMOVE

Private w0 As Boolean 'NO FIRST CHILD of p_hwnd' 3 размера менять можно

'Friend
Function ISubclass_WindowProc(ByVal shWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long) As Long
'ISubclass_WindowProc = Me.WindowProc(shWnd, iMsg, wParam, lParam)
'End Function
'Function ISubclass_WindowProc(ByVal shWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long) As Long
Dim rc As RECT, hDC&, pt As POINTAPI, res&
If shWnd = m_hWnd Then 'Мои MSG
    Select Case iMsg
    Case WM_DESTROY: Class_Terminate
    Case WM_KILLFOCUS, WM_MOUSEMOVE
        ToolTipHide
    Case WM_ENTERSIZEMOVE ' WM_ENTERSIZEMOVE = &H231
        'Debug.Print "RESIZER WM_ENTERSIZEMOVE "
        esm = 1
    Case WM_EXITSIZEMOVE ' WM_EXITSIZEMOVE = &H232
        'Debug.Print "RESIZER WM_EXITSIZEMOVE "
        esm = 0
        ToolTipHide
        PostMessage Target_hWnd, WM_EXITSIZEMOVE, 0, 0
    Case WM_NCCALCSIZE
        Dim nc As NCCALCSIZE_PARAMS
        If wParam Then
            ISubclass_WindowProc = DefWindowProc(m_hWnd, iMsg, wParam, lParam)
            CopyMemory nc, ByVal lParam, LenB(nc)
            nc.rc0.Top = nc.rc0.Top + bw
            nc.rc0.Left = nc.rc0.Left + bw
            nc.rc0.Right = nc.rc0.Right - bw
            nc.rc0.Bottom = nc.rc0.Bottom - bw
            CopyMemory ByVal lParam, nc, LenB(nc)
            ISubclass_WindowProc = &H0
            Exit Function
        End If
    Case WM_SIZING
        CopyMemory rc, ByVal lParam, 16
        'Debug.Print "WM_SIZING " & rc.Left & "x" & rc.Top & "x" & rc.Right & "x" & rc.Bottom
        If rc.Right - rc.Left < bw2 Then
            If wParam = 1 Or wParam = 4 Or wParam = 7 Then rc.Left = rc.Right - bw2 Else rc.Right = rc.Left + bw2
            CopyMemory ByVal lParam, rc, 16
        End If
        If rc.Bottom - rc.Top < bw2 Then
            If wParam = 3 Or wParam = 4 Or wParam = 5 Then rc.Top = rc.Bottom - bw2 Else rc.Bottom = rc.Top + bw2
            CopyMemory ByVal lParam, rc, 16
        End If
    Case WM_MOVE
        If esm Then SizeBind
    Case WM_SIZE
        'Debug.Print "RESIZER WM_SIZE"
        GetWindowRect m_hWnd, rc
        OffsetRect rc, rc.Left, rc.Top
        CreateRGN rc.Right - rc.Left - bw2, rc.Bottom - rc.Top - bw2
        SetWindowRgn m_hWnd, m_rg, 1 '0 'True
        If esm Then SizeBind
    Case WM_SETCURSOR
        If LOWORD(lParam) = 2 Then SetCursor icar(0)(4): Exit Function
    Case WM_NCHITTEST
        GetClientRect m_hWnd, rc
        pt.x = LOWORD(lParam): pt.Y = HIWORD(lParam)
        ScreenToClient m_hWnd, pt
        'Debug.Print "WM_NCHITTEST " & pt.X & "x" & pt.Y & "    " & rc.Right & "x" & rc.Bottom
        If w0 And pt.x < 0 Then
            If pt.Y < 0 Then
                res = 2 'HTCAPTION As Long = 2
            ElseIf rc.Bottom > bw2 And pt.Y > (rc.Bottom - bw) \ 2 And pt.Y < (rc.Bottom + bw) \ 2 Then
                res = 10 'HTLEFT As Long = 10
            ElseIf pt.Y > rc.Bottom Then
                res = 16 'HTBOTTOMLEFT As Long = 16
            End If
        ElseIf rc.Right > bw2 And pt.x > (rc.Right - bw) \ 2 And pt.x < (rc.Right + bw) \ 2 Then
            If w0 And pt.Y < 0 Then
                res = 12 'HTTOP As Long = 12
            ElseIf pt.Y > rc.Bottom Then
                res = 15 'HTBOTTOM As Long = 15
            End If
        ElseIf pt.x > rc.Right Then
            If w0 And pt.Y < 0 Then
                res = 14 'HTTOPRIGHT As Long = 14
            ElseIf rc.Bottom > bw2 And pt.Y > (rc.Bottom - bw) \ 2 And pt.Y < (rc.Bottom + bw) \ 2 Then
                res = 11 'HTRIGHT As Long = 11
            ElseIf pt.Y > rc.Bottom Then
                res = 17 'HTBOTTOMRIGHT As Long = 17
            End If
        End If
        ISubclass_WindowProc = res
        Exit Function
    Case WM_ERASEBKGND
        ISubclass_WindowProc = 1
        Exit Function
    Case WM_NCPAINT
        'Debug.Print "RESIZER NCPAINT " & Now
        GetWindowRect m_hWnd, rc
        OffsetRect rc, -rc.Left, -rc.Top
        hDC = GetWindowDC(m_hWnd)
        If Not w0 Then res = SelectObject(hDC, bb)
        Rectangle hDC, 0, 0, bw, bw 'HTTOPLEFT
        Rectangle hDC, 2, 2, bw - 2, bw - 2 'HTCAPTION
        Rectangle hDC, 0, rc.Bottom - bw, bw, rc.Bottom 'HTBOTTOMLEFT
        Rectangle hDC, rc.Right - bw, 0, rc.Right, bw 'HTTOPRIGHT
        If rc.Right > 2 * bw2 Then Rectangle hDC, (rc.Right - bw) \ 2, 0, (rc.Right + bw) \ 2, bw   'HTTOP
        If rc.Bottom > 2 * bw2 Then Rectangle hDC, 0, (rc.Bottom - bw) \ 2, bw, (rc.Bottom + bw) \ 2        'HTLEFT
        If Not w0 Then SelectObject hDC, res
        If rc.Right > 2 * bw2 Then Rectangle hDC, (rc.Right - bw) \ 2, rc.Bottom - bw, (rc.Right + bw) \ 2, rc.Bottom 'HTBOTTOM
        If rc.Bottom > 2 * bw2 Then Rectangle hDC, rc.Right - bw, (rc.Bottom - bw) \ 2, rc.Right, (rc.Bottom + bw) \ 2  'HTRIGHT
        Rectangle hDC, rc.Right - bw, rc.Bottom - bw, rc.Right, rc.Bottom 'HTBOTTOMRIGHT
        ReleaseDC m_hWnd, hDC
        Exit Function
    'Case &H200  'WM_MOUSEMOVE
    '    Debug.Print "RESIZER WM_MOUSEMOVE "
    Case WM_KEYDOWN
        'Debug.Print "RESIZER WM_KEYDOWN " & wParam
        TrackKeyDown wParam
    End Select
    ISubclass_WindowProc = DefWindowProc(shWnd, iMsg, wParam, lParam)
    
ElseIf shWnd = Parent_hWnd Then 'Папа MSG
    Select Case iMsg
    Case WM_DESTROY, WM_CLOSE: StopBind
    Case Else: ToolTipHide
    End Select
    'Debug.Print "RESIZER PARENT SUBCLASS " & Hex$(iMsg) & " " & p_nm
ElseIf shWnd = Target_hWnd Then 'Клиент MSG
    Select Case iMsg
    Case WM_DESTROY, WM_CLOSE: StopBind
    Case WM_MOVE, WM_SIZE: If esm Or keys Then SetMyPos
    End Select
    'Debug.Print "RESIZER CLIENT SUBCLASS " & Hex$(iMsg) & " " & b_nm
End If
End Function



Private Sub Class_Initialize()
bb = GlobalBrush(2, 0, 0)
bw = 6: bw2 = 2 * bw
m_hWnd = CreateWindowEx(0, xs.sxControl, "xxxx RESIZER xxxx", 0, 0, 0, 10, 10, 0, 0, App.hInstance, 0)
SetWindowLong m_hWnd, GWL_STYLE, WS_CHILD 'Or WS_THICKFRAME ''Or WS_BORDER Or WS_THICKFRAME
StartSubClass Me, m_hWnd
End Sub

Private Sub Class_Terminate()
'If m_rg Then DeleteObject m_rg
If m_hWnd Then
    StopBind
    StopSubClass Me, m_hWnd
    DestroyWindow m_hWnd
    m_hWnd = 0
End If
End Sub

Private Sub CreateRGN(cx&, cy&, Optional mx As Boolean)
Dim rg1&, rg2&, crc As RECT
SetRect crc, 0, 0, cx + bw2, cy + bw2
rg2 = CreateRectRgnIndirect(crc) 'Самый большой регион
If mx Then m_rg = rg2: Exit Sub
SetRect crc, bw, bw, cx + bw, cy + bw
rg1 = CreateRectRgnIndirect(crc) 'Внутренний регион
CombineRgn rg2, rg2, rg1, 4
DeleteObject rg1

If cx > bw2 Then
    SetRect crc, bw, 0, (cx + bw) \ 2, cy + bw2
    rg1 = CreateRectRgnIndirect(crc) 'Первый вертикальный
    CombineRgn rg2, rg2, rg1, 4
    DeleteObject rg1
    SetRect crc, (cx + bw) \ 2 + bw, 0, cx + bw, cy + bw2
    rg1 = CreateRectRgnIndirect(crc) 'Второй вертикальный
    CombineRgn rg2, rg2, rg1, 4
    DeleteObject rg1
Else
    SetRect crc, bw, 0, cx + bw, cy + bw2
    rg1 = CreateRectRgnIndirect(crc) 'Первый вертикальный
    CombineRgn rg2, rg2, rg1, 4
    DeleteObject rg1
End If

If cy > bw2 Then
    SetRect crc, 0, bw, cx + bw2, (cy + bw) \ 2
    rg1 = CreateRectRgnIndirect(crc) 'Первый Горизонтальный
    CombineRgn rg2, rg2, rg1, 4
    DeleteObject rg1
    SetRect crc, 0, (cy + bw) \ 2 + bw, cx + bw2, cy + bw
    rg1 = CreateRectRgnIndirect(crc) 'Второй Горизонтальный
    CombineRgn rg2, rg2, rg1, 4
    DeleteObject rg1
Else
    SetRect crc, 0, bw, cx + bw2, cy + bw
    rg1 = CreateRectRgnIndirect(crc) 'Первый Горизонтальный
    CombineRgn rg2, rg2, rg1, 4
    DeleteObject rg1
    
End If
'If m_rg Then DeleteObject m_rg
m_rg = rg2
End Sub


Public Sub Bind(pw&, cw&)
Debug.Print "Bind " & pw & " = " & cw
'If b = 0 Then MsgBox "NO WND"
If Parent_hWnd <> 0 And pw <> Parent_hWnd Then StopSubClass Me, Parent_hWnd: Parent_hWnd = 0
If Target_hWnd <> 0 And cw <> Target_hWnd Then StopSubClass Me, Target_hWnd: Target_hWnd = 0
If cw = 0 Or pw = 0 Then ShowWindow m_hWnd, 0: SetParent m_hWnd, 0   'Гасим и Убираем на стол
'MsgBox "1 RZ_BIND " & p & " " & b
If pw <> 0 And Parent_hWnd <> pw Then Parent_hWnd = pw: StartSubClass Me, Parent_hWnd
'MsgBox "2 RZ_BIND " & p & " " & b
If cw <> 0 And Target_hWnd <> cw Then Target_hWnd = cw: StartSubClass Me, Target_hWnd
If Parent_hWnd <> 0 Then If GetParent(m_hWnd) <> Parent_hWnd Then SetParent m_hWnd, Parent_hWnd
'MsgBox "3 RZ_BIND " & p & " " & b
SetMyPos
'MsgBox "4 RZ_BIND " & p & " " & b
If IsLongChild(GetFocus(), m_hWnd) Then apiSetFocus m_hWnd
'MsgBox "5 RZ_BIND " & p & " " & b
End Sub



Private Sub SetMyPos()
Dim rc As RECT, pt As POINTAPI, tpt As POINTAPI, hc&
If Target_hWnd = 0 Then
    ToolTipHide
    Exit Sub
Else
    GetWindowRect Target_hWnd, rc
    'tpt.X = rc.Left + bw: tpt.Y = rc.Top + bw
    GetCursorPos tpt
    tpt.x = tpt.x + bw: tpt.Y = tpt.Y + bw
    ScreenToClient Parent_hWnd, pt
    OffsetRect rc, pt.x, pt.Y
    w0 = Not (Get_Parent(Target_hWnd) = Parent_hWnd)
    SetWindowPos m_hWnd, 0, rc.Left - bw, rc.Top - bw, rc.Right - rc.Left + bw2, rc.Bottom - rc.Top + bw2, SWP_SHOWWINDOW Or SWP_NOACTIVATE Or SWP_NOSENDCHANGING Or SWP_NOOWNERZORDER Or SWP_NOZORDER
    
    'TOOLTIP
    pt.x = 0: pt.Y = 0
    GetWindowRect Target_hWnd, rc
    ScreenToClient Get_Parent(Target_hWnd), pt
    OffsetRect rc, pt.x, pt.Y
    hc = GetFocus()
    If hc = m_hWnd Or hc = Parent_hWnd Then
        ToolTipTrackTip tpt.x, tpt.Y, IIf(w0, "L=" & rc.Left & " T=" & rc.Top & vbCrLf, "") & "W=" & rc.Right - rc.Left & " H=" & rc.Bottom - rc.Top & vbCrLf & "R=" & rc.Right & " B=" & rc.Bottom
    End If
    'TOOLTIP
End If

End Sub

Private Sub SizeBind()
Dim rc As RECT, pt As POINTAPI
GetWindowRect m_hWnd, rc
ScreenToClient Get_Parent(Target_hWnd), pt
OffsetRect rc, pt.x, pt.Y
SetWindowPos Target_hWnd, 0, rc.Left + bw, rc.Top + bw, rc.Right - rc.Left - bw2, rc.Bottom - rc.Top - bw2, &H214 '=SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER '
End Sub

Public Sub StopBind()
ShowWindow m_hWnd, 0
SetParent m_hWnd, 0
If Parent_hWnd Then StopSubClass Me, Parent_hWnd
If Target_hWnd Then StopSubClass Me, Target_hWnd
Parent_hWnd = 0: Target_hWnd = 0
End Sub

Private Sub TrackKeyDown(wParam&)
Dim Shift&, dx&, dy&, dcx&, dcy&
Dim rc As RECT, pt As POINTAPI
Shift = GetShiftState() And 3&
'Debug.Print "RESIZER WM_KEYDOWN =" & wParam & " Shift=" & Shift
Select Case wParam
Case 37 'LEFT
    If w0 Then dx = Choose(Shift + 1, -5, 0, -1, -1)
    dcx = Choose(Shift + 1, 0, -1, 0, 1)
Case 38 'UP
    If w0 Then dy = Choose(Shift + 1, -5, 0, -1, -1)
    dcy = Choose(Shift + 1, 0, -1, 0, 1)
Case 39 'RIGHT
    If w0 Then dx = Choose(Shift + 1, 5, 0, 1, 1)
    dcx = Choose(Shift + 1, 0, 1, 0, -1)
Case 40 'DOWN
    If w0 Then dy = Choose(Shift + 1, 5, 0, 1, 1)
    dcy = Choose(Shift + 1, 0, 1, 0, -1)

End Select
keys = (Abs(dx) + Abs(dy) + Abs(dcx) + Abs(dcy)) > 0
If keys Then
    GetWindowRect Target_hWnd, rc
    ScreenToClient Get_Parent(Target_hWnd), pt
    OffsetRect rc, pt.x, pt.Y
    SetWindowPos Target_hWnd, 0, rc.Left + dx, rc.Top + dy, rc.Right - rc.Left + dcx, rc.Bottom - rc.Top + dcy, &H214 '=SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
    keys = 0
    PostMessage Target_hWnd, WM_EXITSIZEMOVE, 0, 0
End If
End Sub


