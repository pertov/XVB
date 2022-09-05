VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

'Private Declare Function IsMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
'Private Declare Function CreateMenu Lib "user32.dll" () As Long
'Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
'Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hWnd As Long, ByVal lptpm As Any) As Long
'Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long

'Private Type MENUINFO
'cbSize As Long ' Размер структуры типа (используйте len, чтобы вычислить)
'fMask As Long ' Маска информации / действий, чтобы обработать
'dwStyle As Long ' Стиль меню (не используемый в этом примере)
'cyMax As Long ' Максимальная высота меню в пикселах (не используется в этом примере)
'hbrBack As Long ' Дескриптор фоновой кисти
'dwContextHelpID As Long ' ИДЕНТИФИКАТОР Справочного контекста (не используется в этом примере)
'dwMenuData As Long ' Данные меню (снова не используются в этом примере)
'End Type
'Private Declare Function SetMenuInfo Lib "user32.dll" (ByVal hMenu As Long, ByRef lpcmi As MENUINFO) As Long

Private Type MENUBARINFO
  cbSize As Long
  rcBar As RECT
  hMenu As Long
  hwndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
Private Const OBJID_MENU As Long = &HFFFFFFFD

Private Declare Function GetMenuBarInfo Lib "user32.dll" (ByVal hWnd As Long, ByVal idObject As Long, ByVal IDItem As Long, ByRef pmbi As MENUBARINFO) As Long

Private items As Collection 'список всех созданных пунктов меню в этом классе
Private ARItems As CParam 'список пунктов меню в этом классе по их ID
Private arhMenu 'список всех созданных меню в этом классе
Private m_hMenu& 'Первое созданное меню
Private mSubMenu As Boolean
'Private m_hbrBack& 'Бакграунд для меню
Private mLastwID& 'wID последнего добавленого пункта меню

'Private accessVE As SIZE


Public hBrush& 'MenuBrush
Public hFont& 'MenuFont
Public nItemHeight& 'MenuItemHeight
Public nItemWidth 'ItemWidth
Public nTextColor& 'TextColor

Public hWnd& 'hWndOwner
Public ActiveID& 'Active menuitem id
'accessVE.cx Private nAccess& 'ACCESS VISIBLE FLAG
'accessVE.cy Private nEnabled& 'ACCESS ENABLED FLAG

Private mAccess& 'BITMASK VISIBILITY ITEMS
Private stopRedraw As Boolean

'Event Redraw()

Friend Property Get Height(hWnd&)
Dim pmbi As MENUBARINFO
pmbi.cbSize = Len(pmbi)
GetMenuBarInfo hWnd, OBJID_MENU, 0, pmbi
Height = pmbi.rcBar.Bottom - pmbi.rcBar.Top + 1
End Property

Private Sub Class_Initialize()
m_hMenu = 0
arhMenu = Array()
Set items = New Collection
Set ARItems = New CParam
End Sub

Sub Destroy()
Class_Terminate
End Sub
Private Sub Class_Terminate()
Dim i&, n&
n = UBound(arhMenu)
For i = n To 0 Step -1
    If IsMenu(0& + arhMenu(i)) Then DestroyMenu 0& + arhMenu(i)
Next
If IsMenu(m_hMenu) Then DestroyMenu m_hMenu
Set items = Nothing
End Sub

Public Property Get hMenu&()
hMenu = m_hMenu
End Property

Function ShowPopup(hWnd As Long, dx As Long, dy As Long, Optional ReturnCMD As Long = 1, Optional nPopupIndex As Long)
If IsMenu(m_hMenu) = 0 Then Exit Function
Dim rc As RECT, pt As POINTAPI, r&
If dx <> 0 Or dy <> 0 Then
    GetWindowRect hWnd, rc
    pt.x = rc.Left + dx: pt.Y = rc.Top + dy
Else
    GetCursorPos pt
End If
'Const TPM_RECURSE As Long = &H1&
'Const TPM_RIGHTBUTTON = &H2&
'Const TPM_CENTERALIGN As Long = &H4&
'Const TPM_RIGHTALIGN As Long = &H8&
'Const TPM_BOTTOMALIGN As Long = &H20&
'Const TPM_VERTICAL As Long = &H40&
'Const TPM_NONOTIFY As Long = &H80&
Const TPM_RETURNCMD As Long = &H100&

Dim hm&
If nPopupIndex >= 0 And nPopupIndex <= UBound(arhMenu) Then
'0=fullmenu
'>0 submenu index
    hm = arhMenu(nPopupIndex)
Else
    Exit Function
End If

r = TrackPopupMenuEx(hm, IIf(ReturnCMD = 1 And Not IsWindow(ReturnCMD), TPM_RETURNCMD, 0), pt.x, pt.Y, IIf(IsWindow(ReturnCMD), ReturnCMD, hWnd), ByVal 0&)
'DestroyMenu m_hMenu


If r = 0 Then Exit Function
'START PROCESS ReturnCMD
Dim mi As CMenuItem
On Error Resume Next
Set mi = items("item_" & r)
If mi Is Nothing Then
    ShowPopup = r
Else
    If mi.cmd = "" Then ShowPopup = r Else ShowPopup = mi.cmd
End If
'Debug.Print "ShowPopup=" & ShowPopup
Err.Clear
End Function

Sub Bild(SubMenu As Boolean, ByVal src$) ', ByVal nacc&)
'src = "<;FaceID;28><File;FaceID;2<File ЖОП\<А\> с ушами;TextColor;&Hff0000;FaceID;0><\s><\nSave><Save;TextColor;255><Recent\tA897<Doc1><Doc2><Doc3>><Exit>><\s><Edit;FaceID;1<Cut\tAlt+F4><Copy><Paste><Remove><Select ALL\tQWERTY>><Help<About><Help Contents>><;FaceID;27><;FaceID;28><;FaceID;29>"
'src = "D:\ADPVODA\XVB\want.jpg<;FaceID;0><File<Open><Save\n><Close><Recent<Doc1><Doc2><Doc3>><Exit>><PERIOD><Edit\g<Cut><Copy><\s><Paste><Remove><Select ALL>><Help<About><Help Contents>><Property\g;FaceID;4><\g\r;FaceID;26><;FaceID;27><;FaceID;28><;FaceID;29>"
'src = "24<;FaceID;0><File<Open><Save\n><Close><Recent<Doc1><Doc2><Doc3>><Exit>><PERIOD><Edit\g<Cut><Copy><\s><Paste><Remove><Select ALL>><Help<About><Help Contents>><Property\g;FaceID;4><\g\r;FaceID;26><;FaceID;27><;FaceID;28><;FaceID;29>"
'If src = "<.>" Then src = "<;FaceID;0><File<Open><Save\n><Close><Recent<Doc1><Doc2><Doc3>><Exit>><PERIOD><Edit\g<Cut><Copy><\s><Paste><Remove><Select ALL>><Help<About><Help Contents>><Property\g;FaceID;4><\g\r;FaceID;26><;FaceID;27><;FaceID;28><;FaceID;29>"
If m_hMenu Then Class_Terminate: Class_Initialize
mSubMenu = SubMenu
If mSubMenu Then m_hMenu = CreatePopupMenu() Else m_hMenu = CreateMenu()

src = xMain.ParseMenu(src, 1) 'UNESCAPE \

Dim mi As MENUINFO, b&, s$
s = GTag(src, 0, "<", ">") 'BACKBRUSH
If s = "" Or s = src Then s = 15

Dim ar
ar = Split(s, ";") 'BackBrush;FontSRC;MenuItemHeight;MenuItemWidth;TextColor

ReDim Preserve ar(4)
hBrush = GlobalBrush(NzS(ar(0), 15), b, 0)
'hFont = GlobalFontSRC(ncm.lfMenuFont, Nz(ar(1)))
hFont = GlobalFontSRC(mncm.lfMenuFont, S_(ar(1)))
nItemHeight = IF0(ar(2), mncm.iMenuHeight)
nItemWidth = L_(ar(3))
nTextColor = L_(ar(4))

mi.cbSize = Len(mi)
'mi.hbrBack = m_hbrBack
mi.hbrBack = hBrush
'mi.cyMax = 150
mi.fMask = &H80000002   'MIM_APPLYTOSUBMENUS=&H80000000 + MIM_BACKGROUND=2 +MIM_MAXHEIGHT=1
mLastwID = 999
'accessVE.cx = nacc 'nAccess = nAcc
'accessVE.cy = nacc 'nAccess = nAcc
mAccess = gAccess

InsertMenuItems m_hMenu, src, Not mSubMenu
SetMenuInfo m_hMenu, mi

End Sub

'Property Get AccessVisible() As Long
'AccessVisible = accessVE.cx 'nAccess
'End Property
'Property Let AccessVisible(v As Long)
'accessVE.cx = v
''RedrawMenu
'End Property
'Property Get AccessEnabled() As Long
'AccessEnabled = accessVE.cy 'nEnabled
'End Property
'Property Let AccessEnabled(v As Long)
'accessVE.cy = v
''RedrawMenu
'End Property


Private Function InsertMenuItems&(hMenu&, src$, Optional TopItem As Boolean)
'SRC = 56456456<File<Open><Save><Close><Recent<Doc1><Doc2><Doc3>><Exit>><Edit<Cut><Copy><Paste><Remove><Select ALL>><Help<About><Help Contents>>65456456
'If gAddIndex(arhMenu, hMenu) <> UBound(arhMenu) Then 'Добавляем в список созданых меню
'Debug.Print "Menu.Duplicate"
'End If
gAddIndex arhMenu, hMenu

Dim n&, i&, srcItem$, uIDNewItem& ', ownIDItem&
Dim m As CMenuItem
n = GTag(src, -1, "<", ">") 'Количество пунктов меню которые надо добавить
For i = 1 To n
mLastwID = mLastwID + 1 'ID Нового пункта меню = НЕПОВТОРЯЕТСЯ В ПРЕДЕЛАХ КЛАССА
uIDNewItem = mLastwID
srcItem = GTag(src, i, "<", ">") 'Получили пункт меню

Set m = New CMenuItem 'Создаем класс пункта меню
m.pClassMenu = ObjPtr(Me)
Set m.Owner = Me

m.inTopMenu = TopItem
'On Error Resume Next
'If m.Fill(srcItem, VarPtr(accessVE)) Then   'Заливаем класс пункта меню
If m.Fill(srcItem, VarPtr(mAccess)) Then   'Заливаем класс пункта меню
    If (m.wFlag And &H10&) = &H10& Then 'Есть дети = добавляем субменю MF_POPUP As Long = &H10&
        uIDNewItem = CreatePopupMenu()
        InsertMenuItems uIDNewItem, srcItem  ', wIDStart
    End If
    m.AppendToMenu hMenu, uIDNewItem
    items.Add m, m.key
    If m.cmd <> "" Then ARItems.Param("i" & m.cmd) = m.key

End If
Set m = Nothing
Next

End Function

'Public Property Get IDItem(ByVal ID) As CMenuItem
'On Error Resume Next
'Set IDItem = items("item_" & ID)
'End Property

'!!!!!! DEFAULT MEMBER !!!!!!
Public Property Get Item(ByVal ID) As CMenuItem
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
On Error Resume Next
Set Item = items("item_" & ID)
End Property
'!!!!!! DEFAULT MEMBER !!!!!!


Public Property Get CMDItem(ByVal ID) As CMenuItem 'id=ID_u
On Error Resume Next
Dim nm$
nm = ARItems.ParamDef("i" & ID, "")
If nm = "" Then Exit Property
Set CMDItem = items(nm)
End Property

Public Property Get CapItem(ByVal s) As CMenuItem
Dim Item: s = S_(s)
For Each Item In items
If Item.Caption = s Then Set CapItem = Item: Exit Property
Next
End Property

Public Property Get Access()
Access = mAccess
End Property
Public Property Let Access(v)
'stopRedraw = 1
mAccess = v
'Dim i&, a&, mi As CMenuItem, nv As Boolean
'For i = 1 To Items.count
'    Set mi = Items(i)
'    a = mi.Access
'    If v = 0 Or a = 0 Then
'        nv = 1
'    Else
'        nv = (v And a) > 0
'    End If
'    If Not (mi.Visible) = nv Then mi.Visible = nv
'Next
'If mSubMenu Then Exit Property
'stopRedraw = 0
'RaiseEvent Redraw
End Property

'Public Sub RemoveItem(id) 'Удаление простого пункта меню
'Dim m As CMenuItem
'Set m = Item(id)
'If m Is Nothing Then Exit Sub
'
''Если (ri.wFlag And MF_POPUP) = MF_POPUP значит остался в воздухе ri.wID который был создан CreatePopupMenu
'If (m.wFlag And &H10&) = &H10& Then
'    RemoveMenuItems m.wID
'    RemoveIndex arhMenu, m.wID 'Удаляем меню из списка
'    DestroyMenu m.wID 'Удаляем меню из жизни
'Else
'    Items.Remove m.key 'Удаляем из колекции
'    RemoveMenu m.phMenu, m.wID, 0
'End If
'If Not mSubMenu Then RaiseEvent Redraw
'End Sub
'
'Public Sub RemoveMenuItems(id&) 'Удаление всех пунктов меню
'Dim rwID&, n&
'While GetMenuItemCount(id) > 0 And n < 1000
'    rwID = GetMenuItemID(id, 0)
'    RemoveItem "item_" & rwID 'Удаляем детенка
'    n = n + 1
'Wend
'End Sub
'
Public Sub Redraw()
'If stopRedraw Then Exit Sub
'RaiseEvent Redraw
If hWnd Then DrawMenuBar hWnd
End Sub
