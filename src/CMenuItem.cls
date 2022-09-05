VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text
Private Declare Function LoadBitmap Lib "user32.dll" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long

Public phMenu As Long 'Меню в которм я сижу

Public inTopMenu As Boolean 'я сижу в топ Меню
Public pClassMenu& 'Класс меню где я хранюсь
Public Owner As CMenu

Private wID As Long 'Мой Идентификатор
Public cmd$ 'Моя команда
Public key$ 'Идентификатор Key= "Item_" & wID

Private mCaption As String 'Подпись
'Public Checked As Boolean
Private Popup As Boolean
Private Separator As Boolean

Private maskVisible& 'BITMASK FOR VISIBILITY
Private maskEnabled& 'BITMASK FOR ENABLED

Private mVisible As Boolean
Private mDisabled&

Public wFlag As Long

Public FaceID As Long
Private FaceSTR$

Private mHeight As Long
Private mWidth As Long

Private mTextColor As Long
Public mBackColor As Long
Private hbrBack& 'HBRUSH Back
'Private sFont As String
Private hFont& '
Private FontHeight& '

'Public DrawFormat As String

Private ptrACCESS&
'Private accVE As SIZE
'cx = Access Visible
'cy = Access Enabled


'Private Type MENUITEMINFO
'    cbSize As Long
'    fMask As Long
'    fType As Long
'    fState As Long
'    wID As Long
'    hSubMenu As Long
'    hbmpChecked As Long
'    hbmpUnchecked As Long
'    dwItemData As Long
'    dwTypeData As String
'    cch As Long
'    hbmpItem As Long
'End Type
'Private Declare Function AppendMenu Lib "user32.dll" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
'Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, ByRef lpMenuItemInfo As MENUITEMINFO) As Long
'Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
'Private Declare Function FloodFill Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'Private Declare Function EnableMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

Private Const MF_GRAYED As Long = &H1&
Private Const MF_DISABLED As Long = &H2&

Private Const MF_CHECKED As Long = &H8&
Private Const MF_POPUP As Long = &H10&
Private Const MF_MENUBARBREAK As Long = &H20&
Private Const MF_MENUBREAK As Long = &H40&
Private Const MF_HILITE As Long = &H80&
Private Const MF_OWNERDRAW As Long = &H100&
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_RIGHTJUSTIFY = &H4000&
Private Const DSS_DISABLED As Long = &H20


'Property Get macc() As Long
'    Dim m As CMenu
'    On Error Resume Next
'    Set m = ObjectFromPtr(pClassMenu)
'    If m Is Nothing Then Exit Property
'    macc = m.AccessVisible
'
''CopyMemory mAcc, ByVal mAccess, 4&
'End Property

Friend Function Fill(src$, ByVal nacc&) As Boolean
Dim props$, i&
Dim ar, cn&, m
wFlag = 0

hbrBack = Owner.hBrush

ptrACCESS = nacc

cn = Val(GTag(src, -1, "<", ">")) 'Узнаем количество детей
If cn > 0 Then Popup = 1: wFlag = wFlag Or &H10& 'MF_POPUP
props = GTag(src, 0, "<", ">") 'Получаем свойства пункта меню
If Len(props) = 0 And cn = 0 Then props = src
'props = props & GTag(src, cn + 1, "<", ">") 'Получаем свойства пункта меню
'FaceID = -1

'hFont = GlobalFontSRC&(ncm.lfMenuFont, "", 0)
hFont = Owner.hFont
mTextColor = AnyColor(Owner.nTextColor)
mWidth = Owner.nItemWidth
mHeight = Owner.nItemHeight

props = Replace(props, "\t", vbTab) 'TAB
mVisible = 1
mDisabled = 0

maskVisible = 0
maskEnabled = 0

'If InStr(1, props, "\c") > 0 Then props = Replace(props, "\c", ""): wFlag = wFlag Or MF_CHECKED
'If InStr(1, props, "\s") > 0 Then props = Replace(props, "\s", ""): wFlag = wFlag Or MF_SEPARATOR: Separator = 1: mHeight = 0
'If InStr(1, props, "\n") > 0 Then props = Replace(props, "\n", ""): wFlag = wFlag Or MF_MENUBARBREAK
'If InStr(1, props, "\b") > 0 Then props = Replace(props, "\b", ""): wFlag = wFlag Or MF_MENUBREAK
'If InStr(1, props, "\r") > 0 Then props = Replace(props, "\r", ""): wFlag = wFlag Or MF_RIGHTJUSTIFY


'cn = InStr(1, props, ";") 'Начало свойств
'If cn = 0 Then 'Нету никаких свойств только подпись
'    mCaption = xMain.ParseMenu(props, 2)
'Else
    ar = Split(props, ";")
    mCaption = xMain.ParseMenu(ar(0), 2) 'Сначала идет подпись
    
For Each m In Split("c8,s&H800,n&H20,b&H40,r&H4000", ",")
If InStr(mCaption, "\" & Left$(m, 1)) > 0 Then mCaption = Replace(mCaption, "\" & Left$(m, 1), ""): wFlag = wFlag Or Val(Mid(m, 2))
Next
If wFlag And MF_SEPARATOR Then Separator = 1: mHeight = 0

    Dim v3&, v8$
    For i = 1 To UBound(ar) - UBound(ar) Mod 2 Step 2
    v8 = xMain.ParseMenu(ar(i + 1), 2): v3 = L_(v8)
    Select Case ar(i)
    Case "I": cmd = v8
    Case "D": mDisabled = v3 And 3
    Case "A": maskVisible = v3  'FOR ACCESS VISIBLE FLAG
    Case "E": maskEnabled = v3 'FOR ACCESS ENABLE FLAG
'        'Const MF_DISABLED As Long = &H2&
'        'Const MF_GRAYED As Long = &H1&
    Case "P": If v3 & "" = v8 Then FaceID = v3 Else FaceSTR = v8: FaceID = 1000 'DRAWSOURCE
    Case "F": If Len(v8) Then hFont = GlobalFontSRC(mncm.lfMenuFont, v8)
    Case "C": mTextColor = v3
    Case "B": mBackColor = v3
    Case "V": mVisible = v3
    Case "W": mWidth = v3
    Case "H": mHeight = v3
    End Select
    Next '
'End If
If mCaption = "" Then mCaption = cmd

wFlag = wFlag Or &H100& 'MF_OWNERDRAW
FontHeight = GetTextWidthHeight(hFont, "Mg")
Fill = -1
End Function

Friend Function AppendToMenu(hMenu&, uIDNewItem&) As Boolean
Dim mi As MENUITEMINFO
Const MF_OWNERDRAW As Long = &H100&

'**************************************************
If L_(cmd) And (wFlag And &H10&) = 0 Then uIDNewItem = L_(cmd)
'**************************************************

AppendToMenu = AppendMenu(hMenu, wFlag, uIDNewItem, mCaption)

If Not AppendToMenu Then Exit Function
key = "item_" & uIDNewItem
wID = uIDNewItem
phMenu = hMenu

If (wFlag And MF_OWNERDRAW) = MF_OWNERDRAW Then
    mi.cbSize = LenB(mi)
    mi.fMask = &H20
    mi.dwItemData = ObjPtr(Me)
    AppendToMenu = SetMenuItemInfo(hMenu, wID, 0, mi)
End If

'Debug.Print "AppendToMenu(" & wID & ") =" & AppendToMenu

End Function

Friend Sub MeasureItem(ByVal lParam&) 'WM_MEASUREITEM Размеры пункта меню
Dim ms As MEASUREITEMSTRUCT, w&, h&
CopyMemory ms, ByVal lParam, LenB(ms)

Dim mVisible1 As Boolean

Dim mAccess&: GetMem4 ByVal ptrACCESS, mAccess
mVisible1 = mVisible And ((maskVisible And mAccess) Or maskVisible = 0)
'Debug.Print "WM_MEASUREITEM " & ms.itemID & " Visible=" & mVisible1 & " accVE.cy=" & accVE.cy
Dim s$

If mVisible1 Then
    If Separator Then 'SEPARATOR
        h = 5 + IIf(Len(mCaption) > 0, 12, 0) 'ncm.iMenuHeight
        If inTopMenu Then w = -9 Else w = 10
    Else
        s = Replace(NzS(mCaption, "Mg"), Chr(9), "      ")
        w = GetTextWidthHeight(hFont, s, 1)  'Ширина текста + отступы слева + справа
        'w = w + IIf(FaceID > -1 Or (wFlag And MF_CHECKED), mncm.iMenuHeight, 0) '+Ширина кнопки слева
        w = w + IIf(FaceID Or (wFlag And MF_CHECKED), mncm.iMenuHeight, 0)  '+Ширина кнопки слева
        If Not inTopMenu And (wFlag And MF_POPUP) Then w = w + 12 '+ширина кнопки справа
        h = GetTextWidthHeight(hFont, s, 0) + 4  'Высота текста + отступы сверху + снизу
        'h = ncm.iMenuHeight - 1 + 1
    End If
    ms.itemWidth = IF0(mWidth, w)
    ms.itemHeight = IF0(mHeight, h)
Else
    ms.itemHeight = 0
    ms.itemWidth = -12
End If
CopyMemory ByVal lParam, ms, LenB(ms)
End Sub

Friend Sub DrawItem(ByVal lParam&) 'WM_DRAWITEM Прорисовка пункта меню
Dim ds As DRAWITEMSTRUCT
CopyMemory ds, ByVal lParam, LenB(ds)

Dim rc As RECT, rcb As RECT, hDC&, ox&
Dim mb As Boolean, utop As Boolean, vtop As Boolean
Dim sar, oldFont&, oldFont1&, dy&
Dim ss$
Dim hbr& 'Currentbrush
Dim bFaceID As Boolean
Static mb0 As Boolean
rc = ds.rcItem
hDC = ds.hDC
If rc.Left = rc.Right Then Exit Sub 'VISIBLE=0

Dim mDisabled1 As Long
Dim mAccess&: GetMem4 ByVal ptrACCESS, mAccess
mDisabled1 = mDisabled
If maskEnabled Then If (maskEnabled And mAccess) = 0 Then mDisabled1 = 3

EnableMenuItem phMenu, ds.itemID, IIf(mDisabled1, 3, 0)
'Debug.Print "START DRAWITEM ds.itemID=" & ds.itemID

mb = KeyIsPressed(1)

If inTopMenu Then utop = (ds.itemState And &H1&) > 0 And ((inTopMenu And Popup) Or mb0)
vtop = (ds.itemState And &H41&) > 0

If mBackColor Then
    hbr = GlobalBrush("" & mBackColor, 0)
Else
    hbr = hbrBack 'Стираем рект
End If

FillRect hDC, rc, hbr 'Стираем рект

Dim lb As LOGBRUSH
gdiGetObject hbr, Len(lb), lb
SetBkColor hDC, lb.lbColor
    
oldFont = SelectObject(hDC, hFont)


If Separator Then    'Разделитель групп
'    If Len(mCaption) = 0 Or mCaption = "-" Then
        If inTopMenu Then
            SetRect rcb, (rc.Left + rc.Right) / 2 - 1, rc.Top + 1, (rc.Left + rc.Right) / 2 + 1, rc.Bottom - 1
        Else
            SetRect rcb, rc.Left + 2, (rc.Bottom + rc.Top) / 2 - 1, rc.Right - 2, (rc.Bottom + rc.Top) / 2 + 1
        End If
        DrawEdge hDC, rcb, 2 + 4, IIf(inTopMenu, 1, 8)
'    End If
    If Len(mCaption) And Not inTopMenu Then
        SetBkMode hDC, 2
        SetTextColor hDC, mTextColor
        DrawText hDC, mCaption, Len(mCaption), rc, DT_SINGLELINE Or DT_CENTER
    End If
Else

'bFaceID = FaceID > -1 Or (ds.itemState And 8&)
bFaceID = FaceID Or (ds.itemState And 8&)
'Debug.Print mCaption, ds.itemState And 8, wFlag And 8

rcb = rc
If mDisabled1 = 0 Then
    If inTopMenu Then   'Рамка вокруг меню
        rcb.Left = rcb.Left + 1: rcb.Top = rcb.Top + 1
        If utop Then
            DrawEdge hDC, rcb, 2, 15 'УТОПЛЕНО
        ElseIf vtop Then
            DrawEdge hDC, rcb, 4, 15 'ВЫПУКЛО
        End If
        SetRect rc, rc.Left + 1, rc.Top + 1, rc.Right - 1, rc.Bottom - 1
    'ElseIf FaceID > -1 And FaceID <> 1000 Then
    ElseIf FaceID And FaceID <> 1000 Then
        If vtop Then 'Рамка вокруг кнопки
            ox = IF0(mHeight, mncm.iMenuHeight) - 1
            SetRect rcb, rc.Left, rc.Top, rc.Left + ox, rc.Top + ox
            DrawEdge hDC, rcb, 4, 15 'ВЫПУКЛО
        End If
        SetRect rcb, rc.Left, rc.Top, rc.Right, rc.Bottom
    End If
End If

dy = (rcb.Bottom - rcb.Top - FontHeight) \ 2 ' Отступ сверху для текста
ox = 0
If utop And (mDisabled = 0) Then ox = ox + 1     'Смещаем нажатую картинку

SetBkMode ByVal hDC, 1


Dim diz&, py&
If mDisabled1 And 1 Then diz = DSS_DISABLED
If bFaceID Then
            'Debug.Print mCaption, "FaceID=" & FaceID

    If FaceID = 1000 Then
        Dim dp As xDRAWPARAMS, hit As RECT, rc0 As RECT
        dp.Enabled = mDisabled1 = 0: dp.hFont = hFont
        dp.BackBrush = hbr ': dp.BackColor = mBackColor
        rc0 = rc: rc0.Right = 0
        'ss = Replace(FaceSTR, "{}", mCaption)
        DrawFormat Nothing, 0, rc0, FaceSTR, hit, dp 'GetDrawWidth
        rc0.Right = hit.Right
        DrawFormat Nothing, hDC, rc, FaceSTR, hit, dp
    Else
        
'        If FaceID > 500 Then
'            If FaceID > 999 Then hIcon = GetIcon(FaceID Mod 1000, sys16) Else hIcon = GetIcon(FaceID Mod 500, sm9): szi = 4
'            'hIcon = GetIcon(FaceID - 500, sm9)
'            'DrawStateIcon hDC, 0, 0, hIcon, 0, rcb.Left + ox + 4, rcb.Top + ox + 4, 0, 0, &H3 Or diz
'            DrawStateIcon hDC, 0, 0, hIcon, 0, rcb.Left + ox + szi, rcb.Top + ox + szi, 0, 0, &H3 Or diz
'        Else
'            'Draw Checked
            If FaceID Then
'            If FaceID > -1 Then
'                'szi = 0
'                If FaceID > 999 Then
'                    hIcon = GetIcon(FaceID Mod 1000, sys16)
'                ElseIf FaceID > 499 Then
'                    hIcon = GetIcon(FaceID Mod 500, sm9): szi = 4
'                Else
'                    hIcon = GetIcon(FaceID, menu16)
'                End If
                
                Dim hIcon&, szi& ': szi = 16  'rc.Bottom - rc.Top
'                If FaceID > 499 Then szi = 9 Else szi = 16
'                If FaceID > 999 Then szi = 16
                'hIcon = CopyImage(GetIcon(FaceID, menu16), 1, szi, szi, 0)

                hIcon = GetIcon(FaceID, menu16, szi)
               ' Debug.Print "DrawItem FaceID=" & FaceID, "szi=" & szi
'       Private Const LR_COPYRETURNORG As Long = &H4
'       Private Const LR_COPYFROMRESOURCE As Long = &H4000
'       Private Const LR_COPYDELETEORG As Long = &H8
'       Private Const LR_MONOCHROME As Long = &H1
                
'                py = 0
'                If szi = 9 Then szi = 0: py = 3
                If szi = 9 Then py = 3
                If szi > 16 Then hIcon = CopyImage(hIcon, 1, 16, 16, 0)
'                hIcon = CopyImage(hIcon, 1, szi, szi, IIf(szi, 0, 4))
'
                DrawStateIcon hDC, 0, 0, hIcon, 0, py + rcb.Left + ox, py + rcb.Top + ox, 0, 0, &H3 Or diz
'                DestroyIcon hIcon
                If szi > 16 Then DestroyIcon hIcon
                'DrawStateIcon hDC, 0, 0, hIcon, 0, rcb.Left + ox + szi, rcb.Top + ox + szi, 0, 0, &H3 Or diz
            Else
                Dim hbmp
                hbmp = LoadBitmap(0, 32760)
                DrawStateIcon hDC, 0, 0, hbmp, 0, rcb.Left + 3, rcb.Top + 1, 0, 0, 4 Or diz
                DeleteObject hbmp
            End If
        'End If
        
    End If
End If

If mDisabled1 = 0 Then
    If Not inTopMenu And vtop Then
        SetTextColor hDC, &HFFFFFF
        SetRect rcb, IIf(bFaceID, IIf(FaceID = 1000, rc0.Right, mncm.iMenuHeight + 1), 0) + rc.Left, rc.Top, rc.Right, rc.Bottom
        'rcb = rc
        FillRect hDC, rcb, comGDI.Brush_HightLight
        'If bFaceID Then rcb.Left = rcb.Left - ncm.iMenuHeight - 1
    Else
        SetTextColor hDC, mTextColor
    End If
End If

diz = diz Or 1
If Len(mCaption) Then  'Текст рисуем 'And FaceID <> 1000
    'If (Not inTopMenu Or FaceID > -1) And mWidth = 0 Then rc.Left = rc.Left + ncm.iMenuHeight 'Отступ после картинкоместа
    'If FaceID > -1 Then rc.Left = rc.Left + IIf(FaceID = 1000, rc0.Right, ncm.iMenuHeight) 'Отступ после картинкоместа
    If bFaceID Then rc.Left = rc.Left + IIf(FaceID = 1000, rc0.Right, mncm.iMenuHeight)  'Отступ после картинкоместа
    rcb = rc
    rcb.Left = rcb.Left + 4 + ox
    rcb.Top = rcb.Top + dy + ox
    sar = Split(mCaption, vbTab)
    ss = sar(0)
    DrawStateText hDC, 0, 0, ss, Len(ss), rcb.Left, rcb.Top, 0, 0, diz
    If UBound(sar) > 0 And Not inTopMenu And mWidth = 0 Then 'TAB 'SHORTCUT KEYS
        Dim rcr As RECT
        rcb.Right = rcb.Right - 4 - 12
        DrawText hDC, sar(1), Len(sar(1)), rcr, DT_SINGLELINE Or DT_RIGHT Or DT_CALCRECT
        rcb.Left = rcb.Right - rcr.Right
        ss = sar(1)
        DrawStateText hDC, 0, 0, ss, Len(ss), rcb.Left, rcb.Top, 0, 0, diz
    End If
    'If Len(Font) Then
End If

If Popup And Not inTopMenu Then 'DFCS_MENUARROW
    SetRect rcb, rc.Right - 12, rc.Top, rc.Right, rc.Bottom
    ox = GlobalFont("Marlett", 10, , , , , , 0)
    oldFont1 = SelectObject(hDC, ox)
    DrawStateText hDC, 0, 0, "4", 1, rcb.Left, rcb.Top + 2, 0, 0, diz
    SelectObject hDC, oldFont1
End If


End If
SelectObject hDC, oldFont: oldFont = 0

ExcludeClipRect ds.hDC, ds.rcItem.Left, ds.rcItem.Top, ds.rcItem.Right, ds.rcItem.Bottom 'Не хуй рисовать там где я нарисовал
CopyMemory ByVal lParam, ds, LenB(ds)
mb0 = mb
Exit Sub

'****************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************
'Const ODS_SELECTED As Long = &H1
'Const ODS_GRAYED As Long = &H2
'Const ODS_DISABLED As Long = &H4
'Const ODS_CHECKED As Long = &H8
'Const ODS_FOCUS As Long = &H10
'Const ODS_DEFAULT As Long = &H20
'Const ODS_HOTLIGHT As Long = &H40
'Const ODS_INACTIVE As Long = &H80
'Const ODS_NOACCEL As Long = &H100

End Sub

'Public Property Get Visible() As Boolean
'Visible = mVisible
'End Property
'Public Property Let Visible(v As Boolean)
'If mVisible = v Then Exit Property
'CopyMemory accVE, ByVal ptrACCESS, Len(accVE)
'mVisible = v
'mVisible = mVisible And ((mAccess And accVE.cx) Or mAccess = 0)
'If inTopMenu Then 'Redraw
''    On Error Resume Next
''    Dim m As CMenu
''    Set m = ObjectFromPtr(pClassMenu)
''    If m Is Nothing Then Exit Sub
'    'm.RedrawMenu
'    Owner.Redraw
'End If
'End Property

Public Property Get Caption$()
Caption = mCaption
End Property
Public Property Let Caption(ByVal v$)
mCaption = Replace(v, "\t", vbTab)
'Redraw
End Property

Public Property Get Icon()
If FaceID = 1000 Then Icon = FaceSTR Else Icon = FaceID
End Property
Public Property Let Icon(ByVal v)
If IsNum(v) Then FaceID = v: FaceSTR = "" Else FaceSTR = v: FaceID = 1000
Owner.Redraw
End Property

Property Get Checked()
Checked = (wFlag And 8) = 8
End Property
Property Let Checked(v)
wFlag = (wFlag And Not 8&) Or IIf(v, 8, 0)
Dim mi As MENUITEMINFO
mi.cbSize = Len(mi)
mi.fMask = 1 'MIIM_STATE
GetMenuItemInfo phMenu, wID, 0, mi
mi.fState = (mi.fState And Not (8&)) Or IIf(v, 8, 0)
SetMenuItemInfo phMenu, wID, 0, mi
End Property

Property Get Disabled()
Disabled = mDisabled
End Property
Property Let Disabled(v)
mDisabled = L_(v) And 3
End Property


'    GetMenuItemInfo phMenu, wID, 0, mi

'Public Property Get Enabled() As Boolean
'Enabled = mEnabled
'End Property
'
'Public Property Let Enabled(v As Boolean)
'If v = mEnabled Then Exit Property
'mEnabled = v
'End Property

'Private Sub Redraw()
'Dim mi As MENUITEMINFO, t&
'    mi.cbSize = Len(mi)
'    mi.fMask = &H10
'    GetMenuItemInfo phMenu, wID, 0, mi
'    t = mi.fType
'    mi.fType = 0 '&H4000
'    'Const MF_RIGHTJUSTIFY As Long = &H4000&
'    SetMenuItemInfo phMenu, wID, 0, mi
'    mi.fMask = &H11
'    mi.fType = t
'    'mi.fState = IIf(mEnabled, &H0, &H2)
'    mi.fState = mDisabled
'    SetMenuItemInfo phMenu, wID, 0, mi
''If Not inTopMenu Then Exit Sub
''
''    Dim m As CMenu
''    On Error Resume Next
''    Set m = ObjectFromPtr(pClassMenu)
''    If m Is Nothing Then Exit Sub
''    'm.RedrawMenu
''
'End Sub

