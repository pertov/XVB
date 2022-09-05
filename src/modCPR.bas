Attribute VB_Name = "modCPR"
Option Explicit

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type
Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, ByRef pptDst As Long, ByRef psize As SIZE, ByVal hdcSrc As Long, ByRef pptSrc As Long, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long


Public Const gray_color = &H888888

Public Declare Function CopyImage Lib "user32.dll" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

'Public Declare Function GdiQueryTable Lib "gdi32.dll" () As Long

'Private Declare Function FindResource Lib "kernel32.dll" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As Long, ByVal lpType As Long) As Long
'Private Declare Function LoadResource Lib "kernel32.dll" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
'Private Declare Function LockResource Lib "kernel32.dll" (ByVal hResData As Long) As Long
'Private Declare Function LookupIconIdFromDirectoryEx Lib "user32.dll" (ByVal presbits As Long, ByVal fIcon As Boolean, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
'Private Declare Function CreateIconFromResourceEx Lib "user32.dll" (ByVal presbits As Long, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
'Private Declare Function CreateIconFromResource Lib "user32.dll" (ByVal presbits As Long, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As Long
'Private Declare Function SizeofResource Lib "kernel32.dll" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
'Private Declare Function GetLastError Lib "kernel32.dll" () As Long


Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long


'Private Type ICONINFO
'    fIcon As Long
'    xHotspot As Long
'    yHotspot As Long
'    hbmMask As Long
'    hbmColor As Long
'End Type
'Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
'Private Declare Function CreateIconIndirect Lib "user32.dll" (ByRef piconinfo As ICONINFO) As Long
Private Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long) As Long


Public Declare Function ImageList_BeginDrag Lib "comctl32.dll" (ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
Public Declare Sub ImageList_EndDrag Lib "comctl32.dll" ()
Public Declare Function ImageList_DragEnter Lib "comctl32.dll" (ByVal hwndLock As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ImageList_DragLeave Lib "comctl32.dll" (ByVal hwndLock As Long) As Long
Public Declare Function ImageList_DragMove Lib "comctl32.dll" (ByVal x As Long, ByVal y As Long) As Long
'Public Declare Function ImageList_SetDragCursorImage Lib "comctl32.dll" (ByVal himlDrag As Long, ByVal iDrag As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long

#If DragDrop Then
Public g_DragIML&
Public iDTH As IDropTargetHelper
#End If

'Option Compare Text

'Public Declare Function GetRgnBox Lib "gdi32.dll" ( _
'     ByVal hRgn As Long, _
'     ByRef lpRect As RECT) As Long

'Public Declare Function CreateRectRgn Lib "gdi32.dll" ( _
'     ByVal X1 As Long, _
'     ByVal Y1 As Long, _
'     ByVal X2 As Long, _
'     ByVal Y2 As Long) As Long
     
'Public Declare Function ExtSelectClipRgn Lib "gdi32.dll" ( _
'     ByVal hdc As Long, _
'     ByVal hRgn As Long, _
'     ByVal fnMode As Long) As Long


Public Declare Function ColorAdjustLuma Lib "shlwapi.dll" ( _
     ByVal clrRGB As Long, _
     ByVal n As Long, _
     ByVal fScale As Long) As Long

'Private Declare Sub ColorRGBToHLS Lib "shlwapi.dll" ( _
'     ByVal clrRGB As Long, _
'     ByRef pwHue As Integer, _
'     ByRef pwLuminance As Integer, _
'     ByRef pwSaturation As Integer)
'Private Declare Function ColorHLSToRGB Lib "shlwapi.dll" ( _
'     ByVal wHue As Integer, _
'     ByVal wLuminance As Integer, _
'     ByVal wSaturation As Integer) As Long


Declare Function OleLoadPicture Lib "OLEPRO32.DLL" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
'Private Declare Sub OleLoadFromStream Lib "ole32.dll" ( _
'     ByRef pStm As Any, _
'     ByVal iidInterface As Long, _
'     ByRef ppvObj As Any)

Private Declare Function ImageList_SetBkColor Lib "comctl32.dll" (ByVal himl As Long, ByVal clrBk As Long) As Long
Private Declare Function RoundRect Lib "gdi32.dll" ( _
     ByVal hDC As Long, _
     ByVal X1 As Long, _
     ByVal Y1 As Long, _
     ByVal X2 As Long, _
     ByVal Y2 As Long, _
     ByVal X3 As Long, _
     ByVal Y3 As Long) As Long
Private Declare Function PolyDraw Lib "gdi32.dll" (ByVal hDC As Long, ByRef lppt As POINTAPI, ByRef lpbTypes As Byte, ByVal cCount As Long) As Long
Private Declare Function EndPath Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function BeginPath Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function StrokeAndFillPath Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function StrokePath Lib "gdi32.dll" (ByVal hDC As Long) As Long

Private Declare Function GetPath Lib "gdi32.dll" (ByVal hDC As Long, ByRef lpPoint As POINTAPI, ByRef lpTypes As Byte, ByVal nSize As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32.dll" (ByVal hDC As Long, ByVal iMode As Long) As Long


'============== GDI+ =============
'Private Type GUID
'   Data1 As Long
'   Data2 As Integer
'   Data3 As Integer
'   Data4(0 To 7) As Byte
'End Type
Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type
Private Type EncoderParameter
   GUID As UUID '.GUID
   NumberOfValues As Long
   Type As Long
   Value As Long
End Type
Private Type EncoderParameters
   count As Long
   Parameter As EncoderParameter
End Type
'Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, ID As GUID) As Long
Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long

'Private Declare Function GdipCreateBitmapFromFile Lib "GdiPlus" (ByVal FileName As Long, bitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipGetImageThumbnail Lib "gdiplus" (ByVal Image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, ByVal Callback As Long, ByVal callbackData As Long) As Long
'Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As UUID, encoderParams As Any) As Long
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long

Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, hImage As Long) As Long
'Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As String, ByRef hImage As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, Image As Long) As Long

'Private Declare Function GdipCreateBitmapFromGraphics Lib "GdiPlus" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal hGraphics As Long, ByRef pbitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hbmp As Long, ByVal hpal As Long, ByRef pBmp As Long) As Long

Private Type PicBmp
    SIZE As Long
    Type As Long
    hbmp As Long
    hpal As Long
    reserved As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (PicDesc As PicBmp, RefIID As UUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal Stream As IUnknown, bmp As Long) As Long

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bmp As Long, hbmReturn As Long, ByVal BackGround As Long) As Long

'Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bmp As Long, ByVal x As Long, ByVal Y As Long, color As Long) As Long
'Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, Image As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal hImage As Long, PixelFormat As Long) As Long

'Private Type ColorPalette
'   Flags As Long   ' Palette flags
'   count As Long           ' Number of color entries
'   Entries(255) As Long         ' Palette color entries; this CAN be an array!!!! (Use CopyMemory and a string or byte array as workaround)
'End Type
Private Declare Function GdipBitmapConvertFormat Lib "gdiplus" ( _
   ByVal nBitmap As Long, ByVal PixelFormat As Long, _
   ByVal DitherType As Long, ByVal PaletteType As Long, _
   ByVal nPalette As Long, ByVal alphaThresholdPercent As Single) As Long
'
'Private Declare Function GdipInitializePalette Lib "gdiplus" ( _
'   ByRef nPalette As Long, ByVal PaletteType As Long, _
'   ByVal optimalColors As Long, ByVal useTransparentColor As Long, _
'   ByVal nBitmap As Long) As Long
''Public Declare Function GdipCreateMetafileFromEmf Lib "gdiplus" (ByVal hEmf As Long, ByVal bDeleteEmf As Long, metafile As Long) As Long 'GpStatus
'Public Declare Function GdipPlayMetafileRecord Lib "gdiplus" (ByVal metafile As Long, ByVal recordType As EmfPlusRecordType, ByVal flags As Long, ByVal dataSize As Long, byteData As Any) As GpStatus



Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal mGraphics As Long) As Long

Public Declare Function GdipGetImageBounds Lib "gdiplus" (ByVal nImage As Long, srcRect As RECTF, srcUnit As Long) As Long
Public Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Interpolation As Long) As Long
Public Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, ByVal PixOffsetMode As Long) As Long

Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal img As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Public Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal Callback As Long, ByVal callbackData As Long) As Long
Public Type RECTF
    nLeft As Single
    nTop As Single
    nWidth As Single
    nHeight As Single
End Type
'Private Type PICTDESC
'   cbSizeOfStruct As Long
'   picType As Long
'   hgdiObj As Long
'   hPalOrXYExt As Long
'End Type
'Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PICTDESC, riid As UUID, ByVal fOwn As Boolean, lplpvObj As Object)
'============== GDI+ =============



'Public Type XFORM
'    eM11 As Single
'    eM12 As Single
'    eM21 As Single
'    eM22 As Single
'    eDx As Single
'    eDy As Single
'End Type
'Public Declare Function SetWorldTransform Lib "gdi32.dll" ( _
'     ByVal hDC As Long, _
'     ByRef lpXform As XFORM) As Long
'Public Declare Function GetWorldTransform Lib "gdi32.dll" ( _
'     ByVal hDC As Long, _
'     ByRef lpXform As XFORM) As Long
'Public Declare Function ModifyWorldTransform Lib "gdi32.dll" ( _
'     ByVal hDC As Long, _
'     ByRef lpXform As XFORM, _
'     ByVal iMode As Long) As Long

'Public Declare Function SetGraphicsMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal iMode As Long) As Long
'Public Declare Function GetGraphicsMode Lib "gdi32.dll" (ByVal hDC As Long) As Long


'Private Declare Function DestroyCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Declare Function LoadIcon Lib "user32.dll" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function LoadIconS Lib "user32.dll" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

Declare Function LoadImage Lib "user32.dll" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
'HANDLE LoadImage(
'    HINSTANCE hinst,    // handle of the instance that contains the image
'    LPCTSTR lpszName,   // name or identifier of image
'    UINT uType, // type of image
'    int cxDesired,  // desired width
'    int cyDesired,  // desired height
'    UINT fuLoad // load flags
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long



'Const EM_GETPARAFORMAT = 1085
'Const EM_SETPARAFORMAT = 1095
'





Private Type SHFILEINFO
    hIcon As Long ' : icon
    iIcon As Long ' : icondex
    dwAttributes As Long ' : SFGAO_ flags
    szDisplayName As String * 260 ' : display name (or path)
    szTypeName As String * 80 ' : type name
End Type
'SHGFI_SYSICONINDEX (0x000004000)
'SHGFI_SMALLICON (0x000000001)
'SHGFI_USEFILEATTRIBUTES (0x000000010)
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long



'Private Type BITMAPINFOHEADER
'    biSize As Long
'    biWidth As Long
'    biHeight As Long
'    biPlanes As Integer
'    biBitCount As Integer
'    biCompression As Long
'    biSizeImage As Long
'    biXPelsPerMeter As Long
'    biYPelsPerMeter As Long
'    biClrUsed As Long
'    biClrImportant As Long
'End Type
'
'Private Type RGBQUAD
'    rgbBlue As Byte
'    rgbGreen As Byte
'    rgbRed As Byte
'    rgbReserved As Byte
'End Type
'
'Private Type BITMAPINFO
'    bmiHeader As BITMAPINFOHEADER
'    bmiColors(1) As RGBQUAD
'End Type
'Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long



Public Type xDRAWPOS
VP As POINTAPI 'Отрисованный VP.X/VP.Y
cIndex As Long
colIndex As Long
hIndex As Long
Ready As Boolean
End Type

Public Type xDRAWPARAMS
hFont As Long
FontHeight As Long
TextAlign As Long 'DrawTextFlag DT_CENTER ++++ 1=Center, 2=Rught, 4=VCenter, 8=VBottom
BorderColor As Long '-1 = No Border color
ClientBorder As Long '-1 = No CLIENT Border
ForeColor As Long ' INITAL FORE COLOR
BackBrush As Long 'INITAL BACKBRUSH
BackColor As Long
NCBackBrush As Long 'INITAL NCBRUSH
Enabled As Boolean
NCEnabled As Boolean
Transparent As Boolean
ChildFocus As Boolean
Focus As Boolean
BorderWidth As Byte
GridLines As Byte
SelBrush As Long 'seltype=7 brush

ParentBackBrush As Long 'Dynamic ParentBrush

CurrentBrush As Long 'Dynamic BackBrush
CurrentForeColor As Long 'Dynamic ForeColor
CurrentClientBorder As Long 'Dynamic ClientBorderColor
CurrentFont As Long 'Dynamic ClientFont

CurrentBrushOrignX As Long 'Dynamic BrushOrignX to CLIENTRECT
CurrentBrushOrignY As Long 'Dynamic BrushOrignY to CLIENTRECT

'RowBrush As Long 'BG-3
'RowColor As Long '-4
'RowBorder As Long '-5
End Type

Public Type TEXTRANGE
'    cp1 As Long
'    cp2 As Long
    cp As CHARRANGE
    Text As String
End Type


Private Type iDraw
    bs As Long 'Border Style
    ro As Long 'Ofset RECT
    lh As String 'Line Height
    ta As Long 'Text Align
    tt As Long 'Text Top Margin
    tl As Long 'Text Left Margin
    'tb As Long 'Text Bottom Margin
    tr As Long 'Text Right Margin
    
    lf As Long 'Next Line Offset Y
    
    fr As Long 'FocusRect
    rr As Long 'RoundRect Radius
    
    pd As String 'PolyDraw M,2,-3,L,5,4,B,4,6,5,7,C
    
    nm As String 'FrameName
    nc As Long 'FrameCursor
    
    tc As String 'TextColor
    bg As String 'BackGround
    bx As Long 'Brush ORIGN X
    by As Long 'Brush ORIGN Y
    
    pc As String 'PenColor
    pw As Long 'PenWidth
    df As String 'FrameControl
    ff As String 'Font
    ic As String 'Icon Index
    
    icalign As Long 'Icon align
    icleft As Long 'Icon left
    ictop As Long 'Icon top
    
    sx As String 'Left Start X
    sy As String 'Top Start Y
    dx As Long 'DrawStepX
    dy As Long 'DrawStepY
    ra As String 'Right Start X
'    ay As String 'Bottom Start Y
    
    dw As String 'DrawWidth
    dh As String 'DrawHeight
End Type


Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

Public Type typeCommonGDI 'Хранит глобальные хэндлы GDI
    flag As Long '1=Date
    Font_Small As Long 'Подписи кнопок << < > >>
    Font_Normal As Long 'Числа других месяцев + подписи месяц/год + сегодня + отмена
    Font_Bold As Long 'Числа выбранного месяца
    Brush_Face As Long 'Фон календаря
    Brush_White As Long 'Фон дат
    Brush_ToolTip As Long 'Фон кнопки сегодня и отмена
    Brush_LightGray As Long 'Для подсветки фона когда мышь овер
    Brush_HightLight As Long 'Для подсветки COLOR_HIGHLIGHT
    'Brush_Menu As Long 'Для подсветки COLOR_MENU
    Brush_Gray As Long 'Фон кнопок << < > >>
    'Brush_Gray1 As Long 'Фон кнопок << < > >>
    Brush_Green As Long 'Фон текущей даты когда мышь овер
    Brush_LightGreen As Long 'Фон текущей даты когда мышь аут
    Pen_White As Long 'Белый ПЭН
    Pen_Black As Long 'Черный ПЭН
    Pen_Gray As Long 'Серй ПЭН
End Type


Public comGDI As typeCommonGDI
Public GlobalGDI As CParam 'Куча объектов GDI
Public GlobalGDICount As CParam 'Счетчик объектов GDI
Public cur32&, sm9&, menu16&, sys16&

'Public cur32ar22, ic32ar


Public icar(4) 'cur32ar, ic32ar,menu16ar, sm9ar, sys16ar

#If DragDrop Then

Public DragCursor(3) As Long

#End If

'cur32ar=icar(0), ic32ar=icar(1) menu16ar=icar(2), sm9ar=icar(3), sys16ar=icar(4)

Dim bimlar      'IMAGELIST ПЛЮС МИНУС ЧЕКБОКС
Private m_gdip_Images

Public hCursor_ARROW&, hCursor_WAIT&
Public hrds As New CParam ', HordsOff As Boolean
'Public xDebug As xControl

Private tempHDC&

Public m_GDIP_Token As Long

'Public gdiBitmapCount&

Public Function DrawFrameControl(ByVal hDC As Long, ByRef lpRect As RECT, ByVal un1s, ByVal un2 As Long) As Long ', Optional txt$, Optional DT_FLAG&) As Long
'Exit Function
Dim un1 As Long ', rc As RECT
un1 = L_(un1s) ': un2 = un2s
'Debug.Print "DrawFrameControl =" & un1 & "," & un1s & "," & un2 & " "
'un2 and &HF = "PARTID
'un2 and &H100 DFCS_INACTIVE
'un2 and &H200 = DFCS_PUSHED
'un2 and &H1000 = DFCS_HOT

If un1 < 5 And un1s = un1 & "" Then
    On Error Resume Next
    'If un1 = 0 Then un1 = 4
    'un2 = 1 'un2 And &HEFFF& ''&HE71F&
    DrawFrameControl = DrawFrameControl32(hDC, lpRect, un1, un2)
'    If un2 And &H1000 Then 'HOVER
'        rc = lpRect
'        OffsetRect rc, -1, -1
'        DrawFocusRect hDC, rc
'    '    FrameRect hDC, rc, GetSysColorBrush(0) 'brush()
'    End If
Else
    Dim theme&, sClass$, PartID&, StateId&, un1p&, un2p&, ch$
    PartID = 1: StateId = 1
    un2p = un2
    ch = Left(un1s, 1)
    Select Case ch
    Case "5", "B"
         sClass = xs.sBUTTON
         PartID = 1
         StateId = IIf(un2 And &H100, 4, IIf(un2 And &H200, 3, IIf(un2 And &H1000, 2, 1))) 'Button
    Case "7", "S"
        un1p = 3
        un2p = (un2 - 1) + IIf(un2 = 3, 1, IIf(un2 = 4, -1, 0))
        sClass = "Spin": PartID = (7 And un2p) + 1: StateId = IIf(un2 And &H100, 4, IIf(un2 And &H200, 3, IIf(un2 And &H1000, 2, 1))) 'Spin
    Case "T": sClass = xs.sTab: PartID = 1: StateId = IIf(un2 And &H200, 3, 4) 'Tabs
    Case "C":
        un1p = 4
        un2p = IIf(un2 And 1, &H400, IIf(un2 And 2, &H500, &H0))
        sClass = xs.sBUTTON: PartID = 3: un2 = un2 And 3&: StateId = IIf(un2, IIf(un2 = 2, 9, 5), 1) 'CheckBox 0,1,2
    Case "R"
         un1p = 4
         un2p = 4 + IIf(un2 And 1, &H400, 0)
         sClass = xs.sBUTTON: PartID = 2: un2 = un2 And 3&: StateId = IIf(un2, 5, 1) 'Radio 0,1
    Case "W", "X" 'CAPTION/STARTPANEL
        sClass = IIf(ch = "W", "Window", "STARTPANEL")
        PartID = un2 And &HFF 'WP_CLOSEBUTTON                    18
        StateId = IIf(un2 And &H100, 4, IIf(un2 And &H200, 3, IIf(un2 And &H1000, 2, 1))) 'Button

'/* WINDOW parts */
'#define WP_CAPTION                        1
'#define WP_SMALLCAPTION                   2
'#define WP_MINCAPTION                     3
'#define WP_SMALLMINCAPTION                4
'#define WP_MAXCAPTION                     5
'#define WP_SMALLMAXCAPTION                6
'#define WP_FRAMELEFT                      7
'#define WP_FRAMERIGHT                     8
'#define WP_FRAMEBOTTOM                    9
'#define WP_SMALLFRAMELEFT                 10
'#define WP_SMALLFRAMERIGHT                11
'#define WP_SMALLFRAMEBOTTOM               12
'#define WP_SYSBUTTON                      13
'#define WP_MDISYSBUTTON                   14
'#define WP_MINBUTTON                      15
'#define WP_MDIMINBUTTON                   16
'#define WP_MAXBUTTON                      17
'#define WP_CLOSEBUTTON                    18
'#define WP_SMALLCLOSEBUTTON               19
'#define WP_MDICLOSEBUTTON                 20
'#define WP_RESTOREBUTTON                  21
'#define WP_MDIRESTOREBUTTON               22
'#define WP_HELPBUTTON                     23
'#define WP_MDIHELPBUTTON                  24
'#define WP_HORZSCROLL                     25
'#define WP_HORZTHUMB                      26
'#define WP_VERTSCROLL                     27
'#define WP_VERTTHUMB                      28
'#define WP_DIALOG                         29
'#define WP_CAPTIONSIZINGTEMPLATE          30
'#define WP_SMALLCAPTIONSIZINGTEMPLATE     31
'#define WP_FRAMELEFTSIZINGTEMPLATE        32
'#define WP_SMALLFRAMELEFTSIZINGTEMPLATE   33
'#define WP_FRAMERIGHTSIZINGTEMPLATE       34
'#define WP_SMALLFRAMERIGHTSIZINGTEMPLATE  35
'#define WP_FRAMEBOTTOMSIZINGTEMPLATE      36
'#define WP_SMALLFRAMEBOTTOMSIZINGTEMPLATE 37

    End Select
    
    
    theme = OpenThemeData(0, StrPtr(sClass))
'    ThemeRun = theme

    If theme Then
        DrawThemeBackground theme, hDC, PartID, StateId, lpRect, lpRect
    Else
        DrawFrameControl = DrawFrameControl32(hDC, lpRect, un1p, un2p)
        un2p = IIf(un2p And &H200, 1, 0)
    End If

    CloseThemeData theme
End If

End Function
'Sub DrawTextU(ByVal hDC As Long, ByRef lpRect As RECT, ByVal s$, ta&, PartID&)
'Dim theme&
'theme = OpenThemeData(0, StrPtr(xs.sBUTTON))
'DrawThemeText theme, hDC, PartID, 1&, StrPtr(s), Len(s), ta, 0, lpRect
'CloseThemeData theme
'End Sub
'================ MemDC ==================
Public Function CreateMemDC(hmemDC&, ByVal Width&, ByVal Height&) As Long 'Return HBMP
'Static c&
'Dim c0&: c0 = c
Width = Width And &H1FFF
Height = Height And &H1FFF

Dim lhDCC&, hbmp&, hOldBmp&, bm As BITMAP
lhDCC = tempHDC 'CreateDC("DISPLAY", "", "", ByVal 0&)
If Not (hmemDC = 0) Then 'Надо удалить DC
    hOldBmp = CreateCompatibleBitmap(lhDCC, 1, 1) 'Создаем картинку 1x1
    'If hOldBmp Then gdiBitmapCount = gdiBitmapCount + 1
    hbmp = SelectObject(hmemDC, hOldBmp) 'Берем старую картинку из DC
    gdiGetObject hbmp, Len(bm), bm
    If bm.bmWidth = Width And bm.bmHeight = Height And Not (Width = 0 Or Height = 0) Then
        hOldBmp = SelectObject(hmemDC, hbmp) 'Садим обратно картинку
        'If DeleteObject(hOldBmp) Then gdiBitmapCount = gdiBitmapCount - 1 Else Stop
        DeleteObject hOldBmp: hOldBmp = 0 'Удаляем картинку 1x1
        'Debug.Print "CreateMemDC = DUBLICATED " & Width & "x" & Height
        Height = 0: Width = 0 'Не надо создавать DC потомушто оно такое же как и было
    Else
        'If DeleteObject(hbmp) Then gdiBitmapCount = gdiBitmapCount - 1 Else Stop
        DeleteObject hbmp: hbmp = 0 'Удаляем старую картинку
        'If DeleteObject(hOldBmp) Then gdiBitmapCount = gdiBitmapCount - 1 Else Stop
        DeleteObject hOldBmp: hOldBmp = 0 'Удаляем картинку 1x1
        DeleteDC hmemDC: hmemDC = 0 'Удаляем hmemDC
'        c0 = c0 - 1
    End If
End If
If Width <> 0 And Height <> 0 Then 'Надо создать DC с новыми размерами
'c0 = c0 + 1

    hmemDC = CreateCompatibleDC(lhDCC) 'Создаем совместимое DC
    hbmp = CreateCompatibleBitmap(lhDCC, Width, Height) 'Создаем битмап с новыми размерами
    'If hbmp Then gdiBitmapCount = gdiBitmapCount + 1
    hOldBmp = SelectObject(hmemDC, hbmp) 'Потом удалим её
    'If DeleteObject(hOldBmp) Then gdiBitmapCount = gdiBitmapCount - 0 Else Stop
    DeleteObject hOldBmp: hOldBmp = 0
    
    'Debug.Print "================ CreateMemDC = " & hBmp & " = " & Width & "x" & Height
End If
'If c <> c0 Then Debug.Print "DCCOUNT=" & c
'c = c0
'DeleteDC lhDCC
'CreateMemDC = hbmp
'xMain.DebugPrint 0, "CreateMemDC (" & hmemDC & "," & Width & "," & Height & ") ================ gdiBitmapCount = " & gdiBitmapCount

End Function

'Function GetBitmap(hDC&) As Long
'Dim hOldBmp&
'If hDC = 0 Then Exit Function
'hOldBmp = CreateCompatibleBitmap(hDC, 1, 1) 'Создаем картинку 1x1
'GetBitmap = SelectObject(hDC, hOldBmp) 'Берем старую картинку из DC
'SelectObject hDC, GetBitmap
'DeleteObject hOldBmp
'End Function
Public Function GetTextWidthHeight&(hFont&, txt$, Optional Return_Witdh As Boolean)
'Public Function GetTextWidthHeight&(hWnd&, hFont&, Txt$, Optional Return_Witdh As Boolean)
Dim oldFont&, hDC&, sz As SIZE
'hDC = GetDC(hWnd)
'CreateMemDC hDC, 1, 1

oldFont = SelectObject(tempHDC, hFont)
GetTextExtentPoint32 tempHDC, txt, Len(txt), sz
SelectObject tempHDC, oldFont
'ReleaseDC hWnd, hDC
'CreateMemDC hDC, 0, 0
'DeleteDC hDC

GetTextWidthHeight = IIf(Return_Witdh, sz.cx, sz.cy)
End Function

'================ MemDC ==================


'================ IMAGE LIST ==================
Public Sub ImageList_AddIcons(hImageList&, Optional pic As StdPicture, Optional ByVal OffsetLeft&, Optional ByVal OffsetTop&, Optional ByVal IconSize&, Optional ByVal nColumns&, Optional ByVal nRows&, Optional ByVal StepX&, Optional ByVal StepY&, Optional ByVal crMask& = -1, Optional ByVal icIndex&)
Dim hbmp&, hBmpOld&, hBmpCopy&, hBmp11&
If pic Is Nothing Or icIndex < 0 Then
    ImageList_Destroy hImageList
    hImageList = 0
End If
Dim x&, y&, dc&, dc0&, tbm As BITMAP, a&, cx&, cy&, n&, h&, nc&
If pic Is Nothing Then Exit Sub

If nColumns > 0 And nRows > 0 Then 'CREATE OR ADDTO IMAGELIST
    'tbm.bmBitsPixel = 32 '16 '24
'    tbm.bmWidth = MulDiv(pic.Width, 96, 2540)
'    tbm.bmHeight = MulDiv(pic.Height, 96, 2540)
    gdiGetObject pic.handle, Len(tbm), tbm

    CreateMemDC dc, IconSize, IconSize 'Создаем DC откуда будем брать картинки для IMAGELISTA
    hBmp11 = CreateCompatibleBitmap(dc, 1, 1) 'Заготовка картинки для выбора
    'If hBmp11 Then gdiBitmapCount = gdiBitmapCount + 1
    dc0 = CreateCompatibleDC(dc)
    hBmpOld = SelectObject(dc0, pic.handle) 'Выбираем туда картинку из файла
    If hImageList = 0 Then hImageList = ImageList_Create(IconSize, IconSize, 1 + 32, 0, 0)  'ILC_MASK = &H1&
    cy = 0
    For y = OffsetTop To tbm.bmHeight Step StepY 'Плывем по вертикали и заливаем ICONSы
    cx = 0
    For x = OffsetLeft To tbm.bmWidth Step StepX 'Плывем по горизонтали и заливаем ICONSы
    n = n + 1

 If icIndex < 1 Or icIndex = n Then
        BitBlt dc, 0, 0, IconSize, IconSize, dc0, x, y, vbSrcCopy '&HCC0020   'SRCCOPY = &HCC0020
        crMask = GetPixel(dc, 0, 0)
        hBmpCopy = SelectObject(dc, hBmp11) 'Получаем то что накопировали
End If


If icIndex = n Then 'Restore
    Call ImageList_AddMasked(hImageList, hBmpCopy, crMask)    'Добавляем в конец
    nc = ImageList_GetImageCount(hImageList) - 1
    h = ImageList_GetIcon(menu16, nc, 0) 'Берем икон с конца
    Call ImageList_ReplaceIcon(hImageList, n - 1, h) 'Заменяем
    DestroyIcon h
    Call ImageList_Remove(hImageList, nc)  'Удаляем с конца

ElseIf icIndex < 1 Then 'Fill
    Call ImageList_AddMasked(hImageList, hBmpCopy, crMask)   'Заполняем IMAGELIST
End If



 If icIndex < 1 Or icIndex = n Then
        hBmp11 = SelectObject(dc, hBmpCopy) 'Возвращаем на место
End If

        If cx + 1 = nColumns Then Exit For Else cx = cx + 1
    Next
        If cy + 1 = nRows Then Exit For Else cy = cy + 1
    Next
    'Debug.Print "ImageList_AddIcons=", n
    hbmp = SelectObject(dc0, hBmpOld) 'Отдаем то что было
    DeleteObject hBmp11
    'If DeleteObject(hBmp11) Then gdiBitmapCount = gdiBitmapCount - 1 Else Stop
    CreateMemDC dc0, 0, 0 'Удаляем МЕМДС
    CreateMemDC dc, 0, 0 'Удаляем МЕМДС
End If
'xMain.DebugPrint 0, "ImageList_AddIcons ================ gdiBitmapCount = " & gdiBitmapCount

End Sub



Function GetButtonImageList(index&) As Long
Dim ic&, n&, iml&
If VarType(bimlar) < vbArray Then
    n = ImageList_GetImageCount(menu16)
    ReDim bimlar(-7 To n - 1) As Long
End If
If index < LBound(bimlar) Or index > UBound(bimlar) Then Exit Function
GetButtonImageList = bimlar(index)
If GetButtonImageList = 0 Then
    iml = ImageList_Create(16, 16, &H19, 0, 0)        'ILC_COLOR24 = &H18' ILC_MASK = &H1& 'ILC_COLOR4 = &H4
    'ic = GetIcon(index, menu16)
    ic = ImageList_GetIcon(menu16, index, 0)
    'PBS_NORMAL = 1
    'PBS_HOT = 2
    'PBS_PRESSED = 3
    'PBS_DISABLED = 4
    'PBS_DEFAULTED = 5
    'PBS_STYLUSHOT = 6
    For n = 1 To 6
'    If n = 2 Or n = 3 * 0 Then
'        AddDrawedIcon iml, ic, &H83, index
'    ElseIf n = 4 * 0 Then
'        AddDrawedIcon iml, ic, &H23, index 'DST_ICON Or DSS_DISABLED
'    Else
        ImageList_ReplaceIcon iml, -1, ic
'    End If
    Next
    DestroyIcon ic
    bimlar(index) = iml
    GetButtonImageList = iml
    If GetButtonImageList = 0 Then bimlar(index) = -1 'Больше не загружать
End If
If GetButtonImageList = -1 Then GetButtonImageList = 0

End Function

Public Function GlobalCursor&(index&)
'GlobalCursor = c65553 'ARROW
If cur32 = 0 Or index < -2 Then Exit Function
Dim i&, n&
If mSafeArray(icar(0)).cDims = 0 Then
    n = ImageList_GetImageCount(cur32)
    ReDim ar(-2 To n - 1) As Long
    ar(-2) = hCursor_ARROW 'ARROW
    ar(-1) = LoadCursor(0&, ByVal 32649&)  'HAND
    For i = 0 To n - 1
    ar(i) = ImageList_GetIcon(cur32, i, 0)
    Next
    icar(0) = ar
End If
If index >= LBound(icar(0)) And index <= UBound(icar(0)) Then GlobalCursor = icar(0)(index)
'GlobalCursor = GetIcon(index, cur32)
'Debug.Print "GlobalCursor", index
End Function


Public Function GetIcon&(ByVal index&, Optional ByVal himl&, Optional sz&)
'Return ICON from ImageList
Dim ar, n&, i&
sz = 16
Select Case himl
'Case cur32: i = 0 'icar = cur32ar
'Case ic32: i = 1 'icar = ic32ar
Case menu16: i = 2
    If index > 999 Then himl = sys16: index = index Mod 1000: i = 4
    If index > 499 Then himl = sm9: index = index Mod 500: i = 3: sz = 9

Case sm9: i = 3: sz = 9 ' icar = sm9ar
Case sys16: i = 4 'icar = sys16ar
Case Else
    Exit Function
End Select

If index < 0 And himl <> cur32 Then 'APP ICONS32
    himl = -1: i = 1: index = Abs(index + 1) '-1..-7 => 0..6
    sz = 32
End If
If i = 0 Then Exit Function

On Error Resume Next

If mSafeArray(icar(i)).cDims <> 1 Then icar(i) = Array()

If UBound(icar(i)) = -1 Then
    If himl = -1 Then n = 7 Else If i = 4 Then n = 255 Else n = ImageList_GetImageCount(himl) ': Debug.Print "IML_" & i&; ".COUNT=" & n
    ReDim ar(n - 1) As Long
    icar(i) = ar
End If

If index >= LBound(icar(i)) And index <= UBound(icar(i)) Then
    GetIcon = icar(i)(index)
    If GetIcon = 0 Then
        If himl = -1 Then
            'Const IDI_APPLICATION = 32512&                -1       0
            'Const IDI_ERROR As Long = 32513&           -2       1
            'Const IDI_QUESTION As Long = 32514&    -3       2
            'Const IDI_WARNING As Long = 32515&      -4       3
            'Const IDI_ASTERISK As Long = 32516&      -5       4
            'Const IDI_WINLOGO As Long = 32517&       -6       5
            'Const AAA = AAA&                                               -7       6
            If index = 6 Then GetIcon = LoadIconS(App.hInstance, "AAA") Else GetIcon = LoadIcon(0, 32512& + index)
        Else
            GetIcon = ImageList_GetIcon(himl, index, 0)
            If himl = menu16 And index = 0 Then GetIcon = 1
        End If
        icar(i)(index) = GetIcon
        If GetIcon = 0 Then icar(i)(index) = -1 'Больше не загружать
    End If
    If GetIcon < 2 Then GetIcon = 0
End If
End Function

Function pReplacemenu16Icon(ByVal ii&, ByVal lib$, ByVal ID$)
 'ii= 1-269 'menu16.IconIndex
'lib= module
'id = resindex
If ii > 269 Then Exit Function
Dim h&, i&
If Len(lib) > 2 And ii > 0 Then 'Load and Replace
    i = L_(ID)
    If i < 0 Then
        h = GetIcon(Abs(i), menu16)
'    ElseIf Val(lib) & "" = lib Then 'From xControlDC 16x16
'        h = hWndIcon(L_(lib), sNzS(ID, "0"))
    ElseIf InStr(lib, "+") Then 'From Brush
        h = IconFromBrush(lib, sNzS(ID, "0"))
    Else
        h = LoadImage(LoadLibrary(lib), i, 1, 16, 16, 0)
    End If
    If h Then ImageList_ReplaceIcon menu16, ii, h: If i >= 0 Then DestroyIcon h
Else 'Restore
    If ii Then ImageList_AddIcons menu16, MyPic, 13, 13, 16, 30, 9, 19, 19, , IIf(ii > 0, 1, 0) + ii
End If

If IsArray(icar(2)) Then
    For h = IIf(ii < 1, 0, ii) To IIf(ii < 0, UBound(icar(2)), ii)
        DestroyIcon icar(2)(h)
        icar(2)(h) = 0
    Next
End If

pReplacemenu16Icon = h
End Function

Private Function IconFromBrush(br$, pos$) As Long 'Create Icon16x16 from brush
Const sz& = 16
Dim himl&, bmp&, b0&, dc&, mdc&, crMask&, p
'xMain.Param("Brush\DFIcon") = ds
p = Split(pos & ",", ",")

Dim rc As RECT
SetRect rc, 0, 0, L_(p(0)) + 16, L_(p(1)) + 16
CreateMemDC dc, rc.Right, rc.Bottom 'memdc
FillRect dc, rc, GlobalBrush(br, 0)

CreateMemDC mdc, sz, sz 'memdc
BitBlt mdc, 0, 0, sz, sz, dc, L_(p(0)), L_(p(1)), vbSrcCopy   'copy image to memdc from hwnd.hdc
crMask = GetPixel(mdc, 0, 0)
b0 = CreateCompatibleBitmap(mdc, 1, 1) 'bitmap
bmp = SelectObject(mdc, b0) 'get bitmap with icon
himl = ImageList_Create(sz, sz, 32 + 1, 0, 0) 'create image list
ImageList_AddMasked himl, bmp, crMask
b0 = SelectObject(dc, bmp) 'Возвращаем на место
DeleteObject b0 'delete b0
CreateMemDC mdc, 0, 0 'delete mdc
CreateMemDC dc, 0, 0 'delete dc
IconFromBrush = ImageList_GetIcon(himl, 0, 0) 'copyicon
ImageList_Destroy himl
'xmain.DeleteGDI br
End Function


'Private Function hWndIcon(h&, pos$) As Long 'Create Icon16x16 from HWND.DHC
'Const sz& = 16
'Dim himl&, bmp&, b0&, dc&, mdc&, crMask&, p
'p = Split(pos & ",", ",")
'dc = GetDC(h) 'hwnd.hdc
'CreateMemDC mdc, sz, sz 'memdc
'BitBlt mdc, 0, 0, sz, sz, dc, L_(p(0)), L_(p(1)), vbSrcCopy   'copy image to memdc from hwnd.hdc
'ReleaseDC h, dc
'crMask = GetPixel(mdc, 0, 0)
'b0 = CreateCompatibleBitmap(mdc, 1, 1) 'bitmap
'bmp = SelectObject(mdc, b0) 'get bitmap with icon
'himl = ImageList_Create(sz, sz, 32 + 1, 0, 0) 'create image list
'ImageList_AddMasked himl, bmp, crMask
'b0 = SelectObject(dc, bmp) 'Возвращаем на место
'DeleteObject b0 'delete b0
'CreateMemDC mdc, 0, 0 'delete mdc
'hWndIcon = ImageList_GetIcon(himl, 0, 0) 'copyicon
'ImageList_Destroy himl
'End Function

'================ IMAGE LIST ==================


'===============Common_GDI======================


Public Sub Common_GDI(INIT As Boolean, flag&)

Dim gsi As GdiplusStartupInput: gsi.GdiplusVersion = 1: GdiplusStartup m_GDIP_Token, gsi

'Dim rc As RECT
'DrawFrameControl 0, rc, "B", 0

tempHDC = CreateDC("DISPLAY", "", "", ByVal 0&)
'tempHDC = CreateCompatibleDC(ByVal 0&)
Dim bc&
If (flag And 1) = 1 Then
    If INIT And (comGDI.flag And 1) = 0 Then
        pSysColor 0, Null

        comGDI.flag = comGDI.flag Or 1
        comGDI.Brush_Face = GlobalBrush(15, bc)
        comGDI.Brush_White = GlobalBrush(5, bc)
        comGDI.Brush_ToolTip = GlobalBrush(24, bc)
        comGDI.Brush_Gray = GlobalBrush(gray_color, bc) 'GlobalBrush(&H808080, bc)
        comGDI.Brush_LightGray = GlobalBrush(&HF0F0F0, bc)
        comGDI.Brush_HightLight = GlobalBrush(13, bc) 'COLOR_HIGHLIGHT=13
        'comGDI.Brush_Menu = GlobalBrush(&HF0F6F6, bc) 'COLOR_MENU= 4
        comGDI.Brush_Green = GlobalBrush(&H88FF88, bc)
        comGDI.Brush_LightGreen = GlobalBrush(&HBBFFBB, bc)
        comGDI.Font_Small = GlobalFont("Small Fonts", 6, 1)
        comGDI.Font_Normal = GlobalFont()
        comGDI.Font_Bold = GlobalFont(, , 1)
        comGDI.Pen_White = GetStockObject(6) 'WHITE_PEN
        comGDI.Pen_Black = GetStockObject(7) 'BLACK_PEN
        comGDI.Pen_Gray = CreatePen(0, 0, &HBBBBBB) 'GRAY_PEN
        'Debug.Print Hex$(TranslateColor(&H80000010))
    End If
End If

If (flag And 2) = 2 Then
    hCursor_ARROW = LoadCursor(0, ByVal 32512&) 'IDC_ARROW
    hCursor_WAIT = LoadCursor(0, ByVal 32514&) 'IDC_WAIT
    If cur32 = 0 Then ImageList_AddIcons cur32, MyPic, 2, 212, 32, 7, 1, 33, 33: GlobalCursor 0
    'Const IDC_ARROW As Long = 32512&
    'Const IDC_WAIT As Long = 32514&
    'Const IDC_HAND As Long = 32649
    'Const IDC_UPARROW As Long = 32516&
    'Const IDC_IBEAM As Long = 32513&
'    If menu16 = 0 Then ImageList_AddIcons menu16, Pic, 13, 13, 16, 30, 9, 19, 19       ': MenuIcon 0
    If menu16 = 0 Then pReplacemenu16Icon -1, "", 0
    
    If sm9 = 0 Then ImageList_AddIcons sm9, MyPic, 2, 186, 10, 40, 1, 11, 11      ': MiniIcon 0
    'Set Pic = LoadPicture()
    'Set Pic = Nothing
    
    #If DragDrop Then
    pDragCursor
       
    #End If
End If


End Sub

Private Function MyPic() As StdPicture
Static mPic As StdPicture
'If mPic Is Nothing Then
'    Dim buf() As Byte, mStream As IStream
'    On Error Resume Next
'    buf = LoadResData("xvb.gif", "CUSTOM")
'    Set mStream = CreateStreamOnHGlobal(0, -1)
'    mStream.Write buf(0), UBound(buf) + 1
'    mStream.Seek 0, STREAM_SEEK_SET
'    Dim iPic As IPicture', aGUID(0 To 3) As Long: aGUID(0) = &H7BF80980: aGUID(1) = &H101ABF32: aGUID(2) = &HAA00BB8B: aGUID(3) = &HAB0C3000 ' GUID for IPICTURE
''    OleLoadPicture ByVal ObjPtr(mStream), 0&, 0&, aGUID(0), iPic
'    OleLoadPicture ByVal ObjPtr(mStream), 0&, 0&, IID_IUnknown, iPic
'    Set mPic = iPic
'End If
If mPic Is Nothing Then Set mPic = LoadPictureEx
Set MyPic = mPic
End Function



Function LoadPictureEx(Optional ByVal Url$) As StdPicture ' IPictureDisp
Dim IPic As IPicture ', aGUID(0 To 3) As Long: aGUID(0) = &H7BF80980: aGUID(1) = &H101ABF32: aGUID(2) = &HAA00BB8B: aGUID(3) = &HAB0C3000 ' GUID for IPICTURE
On Error Resume Next
If Len(Url) = 0 Then 'xvb.gif
    Dim buf() As Byte, mStream As IStream
    buf = LoadResData("xvb.gif", "CUSTOM")
    Set mStream = CreateStreamOnHGlobal(0, -1)
    mStream.Write buf(0), UBound(buf) + 1
    mStream.Seek 0, STREAM_SEEK_SET
    OleLoadPicture ByVal ObjPtr(mStream), 0&, 0&, IID_IUnknown, IPic
Else
    Dim w As New XMLHTTP
    w.Open "GET", Url, False
    w.Send
    Set mStream = w.ResponseStream
    Set w = Nothing
    Dim bmp&, hbmp&, pic As PicBmp, cr&
    GdipCreateBitmapFromStream mStream, bmp
    If bmp Then
'Dim pf&
'GdipGetImagePixelFormat bmp, pf
    
        GdipCreateHBITMAPFromBitmap bmp, hbmp, 0
        With pic
           .SIZE = Len(pic)          ' Length of structure.
           .Type = vbPicTypeBitmap  ' Type of Picture (bitmap).
           .hbmp = hbmp              ' Handle to bitmap.
           .hpal = 0
        End With
        OleCreatePictureIndirect pic, IID_IUnknown, 1, IPic
        GdipDisposeImage bmp
    End If
End If

Set LoadPictureEx = IPic
Set mStream = Nothing
End Function


#If DragDrop Then

Private Sub pDragCursor()
Dim hmodole32&, i&
hmodole32 = LoadLibrary("ole32.dll")
For i = 0 To 3
DragCursor(i) = LoadCursor(hmodole32, i + 1)
Next
FreeLibrary hmodole32
End Sub

#End If



'Private Function pvIStreamToPicture(IStream As IUnknown, KeepFormat As Boolean) As IPicture
'    ' function creates a stdPicture from the passed array
'    ' Note: The array was already validated as not empty before this was called
'    Dim aGUID(0 To 3) As Long
'    On Error Resume Next
'    If Not IStream Is Nothing Then
'        aGUID(0) = &H7BF80980    ' GUID for IPICTURE
'        aGUID(1) = &H101ABF32
'        aGUID(2) = &HAA00BB8B
'        aGUID(3) = &HAB0C3000
'        Call OleLoadPicture(ByVal ObjPtr(IStream), 0&, Abs(Not KeepFormat), aGUID(0), pvIStreamToPicture)
'    End If
'End Function


'===============Common_GDI======================

Public Function AnyColor&(nColor) '&)
If nColor > 0 And nColor < 32 Then '32 Then
    'AnyColor = GetSysColor(nColor)
    AnyColor = pSysColor(nColor)
ElseIf nColor = -1 Then
    AnyColor = xa.gLabelForeColor
ElseIf nColor < 0 Then
    AnyColor = -1
Else
    'AnyColor = TranslateColor(nColor)
    AnyColor = &HFFFFFF And nColor
End If
End Function


Function pSysColor(nColor, Optional v) As Long
'nColor = SysColorIndex (0..31)
'?v = Color&: Let SysColor(nColor)=v
'?v = Array(): Let SysColors(1..31)=v(1..31)
'?nColor=0+v = Null: Reset SysColor(1..31)=GetSysColor(1..31)
'?nColor=1..31+v = Null: Reset SysColor(nColor)=GetSysColor(nColor)
Static sc(31) As Long 'SysColorsArray
pSysColor = -1
'Return CustomSysColor n=1..31
If nColor < 0 Or nColor > 31 Then Exit Function
Dim i&
'gDebugPrint "pSysColor(" & nColor & ") Vartype(v)=" & VarType(v)
If VarType(v) = vbError Then 'Get CustomSysColor
    pSysColor = sc(nColor)
ElseIf VarType(v) = vbNull Then 'Reset SysColor/s
    For i = IIf(nColor, nColor, 0) To IIf(nColor, nColor, 31): pSysColor i, GetSysColor(i): Next
ElseIf VarType(v) = vbEmpty Then 'GetSysColor
    pSysColor = GetSysColor(nColor)
ElseIf mSafeArray(v).cDims = 1 Then 'Load SysColors from array
    For i = 0 To UBound(v) Step 2: pSysColor Cast(v(i), vbLong), v(i + 1): Next
Else 'Let SysColor(nColor)=v
    sc(nColor) = AnyColor(Cast(v, vbLong))
    GlobalGDI("brush_" & nColor) = Null

End If

End Function

'Public Function HSL2RGB(ByVal Hue As Long, ByVal Saturation As Long, ByVal Luminance As Long, Optional Validate As Boolean) As Long
'' by Donald (Sterex 1996), donald@xbeat.net, 20011124
'  Dim rr As Single, rG As Single, rB As Single
'  Dim rh As Single, rL As Single, rs As Single
'  Dim rMin As Single, rMax As Single, rDiff As Single
'
'  If Validate Or 1 Then
'    Hue = IIf(Hue < 0, IIf(Hue < -2147483520, 360, 0) + Hue + 2147483520, Hue) Mod 360
'    Saturation = IIf(Saturation < 0, 0, IIf(Saturation > 100, 100, Saturation))
'    Luminance = IIf(Luminance < 0, 0, IIf(Luminance > 100, 100, Luminance))
'  End If
'
'  If Saturation = 0 Then
'    ' CLng(CSng(...)) else 127.5 -> 127
'    HSL2RGB = CLng(CSng(2.55 * Luminance)) * &H10101
'  Else
'    rh = Hue / 60: rs = Saturation / 100: rL = Luminance / 100
'    If rL <= 0.5 Then
'      rMin = rL * (1 - rs)
'    Else
'      rMin = rL - rs * (1 - rL)
'    End If
'    rMax = 2 * rL - rMin
'    rDiff = rMax - rMin
'
'    Select Case Hue \ 60
'    Case 0
'      rr = rMax
'      rB = rMin
'      rG = rh * rDiff + rMin
'    Case 1
'      rG = rMax
'      rB = rMin
'      rr = rMin - (rh - 2) * rDiff
'    Case 2
'      rG = rMax
'      rr = rMin
'      rB = (rh - 2) * rDiff + rMin
'    Case 3
'      rB = rMax
'      rr = rMin
'      rG = rMin - (rh - 4) * rDiff
'    Case 4
'      rB = rMax
'      rG = rMin
'      rr = (rh - 4) * rDiff + rMin
'    Case Else
'      rr = rMax
'      rG = rMin
'      rB = rMin - (rh - 6) * rDiff
'    End Select
'    HSL2RGB = CLng(rB * 255) * &H10000 + CLng(rG * 255) * &H100& + CLng(rr * 255)
'  End If
'
'
'End Function
'
'Public Function RGB2HSL(ByVal RGBValue&) As Long
'' by Donald (Sterex 1996), donald@xbeat.net, 20011116
'
'Dim Hue As Long, Saturation As Long, Luminance As Long
'
'  Dim r As Long, g As Long, b As Long
'  Dim lMax As Long, lMin As Long, lDiff As Long, lSum As Long
'
'  r = RGBValue And &HFF&
'  g = (RGBValue And &HFF00&) \ &H100&
'  b = (RGBValue And &HFF0000) \ &H10000
'
'  If r > g Then lMax = r: lMin = g Else lMax = g: lMin = r
'  If b > lMax Then lMax = b Else If b < lMin Then lMin = b
'
'  lDiff = lMax - lMin
'  lSum = lMax + lMin
'
'  ' Luminance
'  Luminance = lSum / 5.1!
'
'  If lDiff Then
'    ' Saturation
'    If Luminance <= 50& Then
'      Saturation = 100 * lDiff / lSum
'    Else
'      Saturation = 100 * lDiff / (510 - lSum)
'    End If
'    ' Hue
'    Dim q As Single: q = 60 / lDiff
'    Select Case lMax
'    Case r
'      If g < b Then
'        Hue = 360& + q * (g - b)
'      Else
'        Hue = q * (g - b)
'      End If
'    Case g
'      Hue = 120& + q * (b - r)
'    Case b
'      Hue = 240& + q * (r - g)
'    End Select
'  End If
''If Hue = 0 Then Hue = 240
'
'RGB2HSL = Hue * &H10000 + Saturation * &H100& + Luminance
''Debug.Print Hex$(Hue) & " 0" & Hex$(Saturation) & " 0" & Hex$(Luminance); "            " & Hex$(RGB2HSL)
'End Function

'Public Function HSLColor&(ByVal RGBValue&, Optional ByVal NewSaturation& = -1, Optional ByVal NewBrightness& = -1) 'Контрастность/Яркость
'Dim HSL&, Hue&, Saturation&, Brightness&
'HSL = RGB2HSL(RGBValue)
''Debug.Print Hex$(hsl)
''Debug.Print &H63, &HCE
'If NewSaturation = -1 Then Saturation = (HSL And &HFF00&) \ &H100& Else Saturation = NewSaturation And &HFF&
'If NewBrightness = -1 Then Brightness = HSL And &HFF& Else Brightness = NewBrightness And &HFF&
'Hue = (HSL And &H1FF0000) \ &H10000 '0...359(512)
'Hue = Hue Mod 360
'HSLColor = HSL2RGB(Hue, Saturation, Brightness)
'End Function

'Public Function HSLSaturation&(ByVal RGBValue&, ByVal StepSaturation&)   'Контрастность - выше/ниже
'Dim HSL&, Hue&, Saturation&, Brightness&
'HSL = RGB2HSL(RGBValue)
'Brightness = HSL And &HFF
'Saturation = (HSL And &HFF00&) \ &H100&
'Hue = (HSL And &H1FF0000) \ &H10000 '0...359(512)
'Hue = Hue Mod 360
'Saturation = Saturation + Saturation * StepSaturation \ 100
''Saturation = StepSaturation
'Saturation = (Saturation + 1000) Mod 101
'HSLSaturation = HSL2RGB(Hue, Saturation, Brightness)
'End Function



'Public Function HSLBrightness&(ByVal RGBValue&, ByVal nBrightness&)  'Яркость - 0,,100
'HSLBrightness = ColorAdjustLuma(RGBValue, 10 * nBrightness, 1) 'IIf(nBrightness < 0, 1, 1))
'
'End Function


'================ DATE ==================
Public Function HitDateToDate(ByVal hit&, ByVal cPos)
If hit < 13 Or hit > 60 Then Exit Function
Dim d&
cPos = Cast(cPos, vbLong)
d = xMain.Fdat(cPos)  'Начало месяца
If Val(Format$(d, "w", 2)) > 1 Then d = d - Val(Format$(d, "w", 2)) + 1 Else d = d - 7
HitDateToDate = d - 1 + 7 * ((hit - 13) \ 8) + (hit - 13) Mod 8 'DELTA
End Function
Public Function HitDate&(x&, y&, index&, rc As RECT)
Dim r0 As RECT, i&
If index = 0 Then 'Надо вернуть ХИТ по указанным координатам x,y
    If y > 1 And y < 14 Then 'Кнопки
        If x > 1 And x < 16 Then
            HitDate = 1
        ElseIf x > 16 And x < 31 Then
            HitDate = 2
        ElseIf x > 119 And x < 134 Then
            HitDate = 4
        ElseIf x > 134 And x < 149 Then
            HitDate = 5
        End If
    ElseIf y > 25 And y < 103 Then 'Число
        If x > 19 And x < 150 Then
            HitDate = 13 + ((y - 25) \ 13) * 8 + (x - 19) \ 19 + 1
        End If
    ElseIf y > 104 And y < 120 Then 'Сегодня отмена
        If x > 0 And x < 100 Then
            HitDate = 61
        ElseIf x > 103 And x < 149 Then
            HitDate = 62
        End If
    End If
    i = HitDate
Else
    i = index
End If
If i > 0 And i < 63 Then 'Надо вернуть рект по указанному индексу
    Select Case i
    Case 1, 2, 3, 4, 5
        rc.Top = Choose(i, 2, 2, 0, 2, 2)
        rc.Bottom = Choose(i, 13, 13, 14, 13, 13)
        rc.Left = Choose(i, 2, 17, 31, 120, 135)
        rc.Right = rc.Left + Choose(i, 13, 13, 89, 13, 13)
    Case 6, 7, 8, 9, 10, 11, 12
        rc.Top = 13: rc.Bottom = 26
        rc.Left = (i - 5) * 19 - 2
        rc.Right = rc.Left + 18
    Case Is < 61
        rc.Top = 26 + ((i - 13) \ 8) * 13
        rc.Bottom = rc.Top + 12
        rc.Left = ((i - 13) Mod 8) * 19
        If (i - 13) Mod 8 > 0 Then rc.Left = rc.Left - 2
        rc.Right = rc.Left + 18
    Case 61
        rc.Top = 105: rc.Bottom = rc.Top + 15
        rc.Left = 1: rc.Right = 100
    Case 62
        rc.Top = 105: rc.Bottom = rc.Top + 15
        rc.Left = 102: rc.Right = 149
    Case Else
        rc = r0
    End Select
End If
End Function

Public Sub Redraw_Date(hDC&, cPos&, cHit&, index&, vmin&, vmax&)
'Debug.Print "HIT =" & cHit
'If cHit = 62 Then Stop
Dim cm As Boolean, i&, r As RECT, s$, br&, fc&, pn&, fn&, hi& ', r0 As RECT
Dim d&, p0&, b0&, f0& 'OLD_PEN, OLD_BRUSH, OLD_FONT
SetBkMode hDC, 1 'TRANSPARENT
d = xMain.Fdat(cPos)  'Начало месяца
If Val(Format$(d, "w", vbMonday)) > 1 Then d = d - Val(Format$(d, "w", vbMonday)) + 1 Else d = d - 7
For i = IIf(index <= 0, 1, index) To IIf(index <= 0, 62, index)
HitDate 0, 0, i, r
fc = 0: fn = comGDI.Font_Normal
br = 0 'comGDI.Brush_Face
hi = i: pn = 0
Select Case i
Case 1, 2, 3, 4, 5
    s = Choose(i, "<<", "<", Format$(cPos, "mmmm yyyyг."), ">", ">>") 'Кнопки + Месяц год
    If i = 3 Then hi = -1 Else fn = comGDI.Font_Small: br = comGDI.Brush_Gray
Case 6, 7, 8, 9, 10, 11, 12
    s = Format$(3 + i, "ddd") 'Дни недели
    hi = -1: fc = IIf(i < 11, &H808080, IIf(i = 11, &HFF6633, &H8080FF)) 'Черный светлосиний светлокрасный
Case 13, 21, 29, 37, 45, 53
    s = Format$(d + 3, "ww", vbMonday) 'Номер недели
    hi = -1: fc = &HAA6666 '&HDAA5A5
Case 61
    s = "Сегодня " & Format$(Date, "dd.mm.yy") 'Кнопка сегодня
    br = comGDI.Brush_ToolTip
Case 62
    s = "Отмена" 'Кнопка отмена
    br = comGDI.Brush_ToolTip
Case Else 'Число месяца
    If index > 0 Then
        d = d + 7 * ((index - 13) \ 8) + (index - 13) Mod 8
    Else
        d = d + 1 'Прибавляем следующее число
    End If
    s = Day(CDate(d - 1))
    cm = (Month(CDate(cPos)) = Month(CDate(d - 1)))
    br = comGDI.Brush_White
    Select Case (i - 13) Mod 8
    Case 6: fc = IIf(cm, &HFF0000, &HFF6633) 'синий /светлосиний
    Case 7: fc = IIf(cm, &HFF, &H8080FF) 'красный / светлокрасный
    Case Else: fc = IIf(cm, &H0, &H808080) 'черный / серый
    End Select
    If cm Then fn = comGDI.Font_Bold
    
    If d - 1 < vmin Or d - 1 > vmax Then fc = ColorAdjustLuma(fc, 750, 1)

    pn = 1
End Select

f0 = SelectObject(hDC, fn)
SetTextColor hDC, fc

If cHit = hi And i = hi And (hi < 6 Or hi > 60) Then 'Рамка для кнопок
    'Debug.Print "HIT BUTTON =" & hi
    pn = CreatePen(0, 0, 0)
    p0 = SelectObject(hDC, pn)
    If br <> 0 Then b0 = SelectObject(hDC, br)
    Rectangle hDC, r.Left, r.Top, r.Right + 1, r.Bottom  'Фон с рамкой
    SelectObject hDC, p0
    If br <> 0 Then SelectObject hDC, b0: b0 = 0
    DeleteObject pn: pn = 0
Else
    If d - 1 = CLng(Date) And hi = i And i > 5 Then  'Значение календаря
        If i = index Then
            br = comGDI.Brush_LightGreen
        Else
        'Debug.Print index & " =" & Date & " i=" & i
            br = comGDI.Brush_Green
        End If
    ElseIf i = index Then 'HightLight Значение календаря
        'br = comGDI.Brush_Face
        If i = cHit Then br = comGDI.Brush_LightGreen
    End If
    If d - 1 = cPos And pn <> 0 Then 'Рамка вокруг значения
        pn = CreatePen(0, 1, &HEE3333)
        p0 = SelectObject(hDC, pn)
        If br <> 0 Then b0 = SelectObject(hDC, br)
        
        Rectangle hDC, r.Left - 1, r.Top - 1, r.Right + 1, r.Bottom + 1 'Фон с рамкой
        SelectObject hDC, p0
        If br <> 0 Then SelectObject hDC, b0: b0 = 0
        DeleteObject pn: pn = 0
    Else
        If br <> 0 Then FillRect hDC, r, br 'Заливаем фоном без рамки
    End If
End If

DrawText hDC, s, Len(s), r, 1
If b0 <> 0 Then SelectObject hDC, b0
SelectObject hDC, f0
Next
End Sub

'================ DATE ==================

'================ PERIOD ==================
Public Function HitPeriod&(x&, y&, index&, rc As RECT, ByVal style&)
Dim r0 As RECT, i&
If index = 0 Then 'Надо вернуть ХИТ по указанным координатам x,y
    If y > 0 And y < 16 Then
        If x > 1 And x < 13 Then
            HitPeriod = 1
          
        ElseIf x > 17 And x < 107 And style Then
          HitPeriod = 3
            
        ElseIf x > 17 And x < 63 Then
            HitPeriod = 2
        
       
        ElseIf x > 120 And x < 135 And style Then
            HitPeriod = 5
        ElseIf x > 134 And style Then
            HitPeriod = 0
        
        
        
        ElseIf x > 65 And x < 155 Then
            HitPeriod = 3
        ElseIf x > 156 And x < 202 Then
            HitPeriod = 4
            
            
        ElseIf x > 204 And x < 218 Then
            HitPeriod = 5
        End If
    End If
    i = HitPeriod
Else
    i = index
End If
If i > 0 And i < 6 Then 'Надо вернуть рект по указанному индексу
    rc.Top = 0: rc.Bottom = 13
    If i = 1 Then
        rc.Left = 1: rc.Right = 15
        rc.Top = 1: rc.Bottom = 12
    ElseIf i = 2 Then
        rc.Left = 17: rc.Right = 63
    ElseIf i = 3 Then
        If style Then rc.Left = 17: rc.Right = 119 Else rc.Left = 64: rc.Right = 156
    ElseIf i = 4 Then
        rc.Left = 156: rc.Right = 202
    ElseIf i = 5 Then
        If style Then rc.Left = 121: rc.Right = 135 Else rc.Left = 204: rc.Right = 218
        rc.Top = 1: rc.Bottom = 12
    Else
        rc = r0
    End If
End If
End Function
Public Sub Redraw_Period(hDC&, cVal, cHit&, Optional ByVal index&, Optional ByVal style1&) ', Optional OffsetX&, Optional offsetY&)
Dim i&, i2&, ar
Dim rc As RECT
Dim br&, fn&, s$
Dim b0&, f0&, fc&
SetBkMode hDC, 1 'TRANSPARENT
ar = Split(cVal & "xx", "x")
'Debug.Print "Redraw_Period " & cVal
'ЗАЩИТА ОТ ДУРАКА
'ar(0) = Val(ar(0))
'If Not ar(0) > -1 And ar(0) < 6 Then ar(0) = 0
'If Val(ar(1)) = 0 Then ar(1) = 0 'CLng(Date)
'If Val(ar(2)) = 0 Then ar(2) = 0 'xMain.FLDat(ar(1), ar(0), 2) ' = IF0(ar(2), ar(1))
'ЗАЩИТА ОТ ДУРАКА
On Error Resume Next
For i2 = IIf(index = 0, 1, index) To IIf(index = 0, 5, index)
If style1 Then i = Choose(i2, 1, 0, 3, 0, 5) Else i = i2
If i Then
HitPeriod 0, 0, i, rc, style1
'OffsetRect rc, OffsetX, offsetY
fn = comGDI.Font_Normal
SetTextColor hDC, 0
Select Case i
Case 1, 5 'Кнопки
    br = comGDI.Brush_Gray
    fn = comGDI.Font_Small
    s = IIf(i = 1, "<<", ">>") ' IIf(i = 1, "«", "»")
Case 2, 4 'Границы периода
    br = comGDI.Brush_White
    s = ".."
    s = Format$(CDate(ar(IIf(i = 2, 1, 2))), "dd.mm.yy")
    fc = Choose(Weekday(IIf(i = 2, ar(1), ar(2)), vbMonday), 1, 2, 3, 4, 5, &HFF0000, &HFF)
    SetTextColor hDC, fc
Case 3 'Подпись периода
    br = 0
    'Debug.Print
    s = ""
    s = xMain.GetDPerText(ar(1), ar(0), ar(2))
    If ar(0) = 1 Then
        fc = Choose(Weekday(ar(1), vbMonday), 1, 2, 3, 4, 5, &HFF0000, &HFF)
        SetTextColor hDC, fc
    End If
End Select

f0 = SelectObject(hDC, fn)
'b0 = SelectObject(hDC, br)
If index = 1 Or index = 5 Then
    b0 = SelectObject(hDC, br)
    Rectangle hDC, rc.Left, rc.Top, rc.Right, rc.Bottom     'Фон с рамкой
    SelectObject hDC, b0
Else
    If br <> 0 Then FillRect hDC, rc, br 'Заливаем фоном без рамки
End If
If i > 1 And i < 5 Then rc.Top = rc.Top - 1
DrawText hDC, s, Len(s), rc, 1

SelectObject hDC, f0: f0 = 0
End If
Next
End Sub
'================ PERIOD ==================


'================ DRAW FORMAT ==================
Public Function DrawFormat$(xc As xControl, ByVal xHDC As Long, rc0 As RECT, ByVal DrawSRC4Kb$, hit As RECT, DFP As xDRAWPARAMS, Optional ByVal TipText$, Optional NCdx&, Optional NCdy&, Optional nCursor&)
On Error Resume Next
'Hit:
'Hit.Left= mouse.x  <- bound rc[]
'Hit.Top= mouse.y  <- bound rc[]

'Hit.Right= mouse.button ON/OFF

'Hit.Bottom<=-1 DRAW + not FILLBACK
'Hit.Bottom=0 NODRAW ONLY HIT AND BOUNDRECT
'Hit.Bottom=1 DRAW + FILLBACK NC
'Hit.Bottom>=2 DRAW + FILLBACK

If Len(DrawSRC4Kb) = 0 Then Exit Function
If Len(DrawSRC4Kb) > &HFFF Then Exit Function

Dim n&, txt$, i&, gt$, gt0$, gt3&
Dim ar, ic&, c&
Dim ic_i&, ic_sz&, ic_iml&

Dim mOver As Boolean, bCFrame As Boolean
Dim cEnabled As Boolean
Dim prc As RECT 'Область для фрэйма FULLRECT
Dim trc As RECT 'Область для текста
Dim cr As RECT 'Current FRAME RECT
Dim arc As RECT 'Размеры области рисования PAINTRECT
'Dim rsc As Long 'SELTYPE -1=Invert,0=None:>0 =COLOR
Dim hsc As Long 'NAME HIGHLIGHT  -1=Invert,0=None:>0 =COLOR

Dim pt As POINTAPI

Dim toolrc As RECT 'TOOL RECT
Dim txTool$, i0&

Dim bDraw As Boolean 'Надо рисовать
Dim w As iDraw, w0 As iDraw
Dim lf As LOGFONT 'Шрифт текущий
Dim hFont&, oldhFont&
Dim hBrush&, oldhBrush&
Dim hPen&, oldhPen&

Dim hBackBrush& 'Заливать фон
Dim ta&, mTextAlign& 'DT_*** DrawText.Flags
Dim mLineHeight&
Dim vt$, tv$
Dim hIcon&
Dim ic16& ' use_sys16_imagelist
'Dim icon_align As Long

Dim lpt As POINTAPI

Dim s$: s = DrawSRC4Kb
If xc Is Nothing Then
    Do
        vt = GTag(s, 1, "{", "}")
        s = Replace(s, "{" & vt & "}", "")
    Loop While Len(vt)
Else
    s = xc.ReplaceVars(s)
End If

If Len(s) = 0 Then Exit Function

If InStr(s, "‹") Then If InStr(s, "›") Then s = Replace(Replace(s, "‹", "{"), "›", "}")

mLineHeight = DFP.FontHeight + 1


cEnabled = DFP.Enabled

'HIT
bDraw = Not (hit.Bottom = 0) And Not (xHDC = 0) 'DRAW

If hit.Bottom = 1 Then 'NC or MCRC
    hBackBrush = DFP.NCBackBrush
    DFP.CurrentForeColor = DFP.ForeColor
Else
    If hit.Bottom > 1 Then 'CRC
        'hBackBrush = IIf(DFP.Transparent And Not DFP.Focus, DFP.ParentBackBrush, DFP.BackBrush)
        hBackBrush = IIf(DFP.Transparent And Not DFP.Focus, DFP.ParentBackBrush, DFP.CurrentBrush)
    End If
End If

If hBackBrush = 0 Then hBackBrush = DFP.ParentBackBrush
'If hBackBrush = 0 Then hBackBrush = DFP.CurrentBrush
'If hBackBrush = 0 Then hBackBrush = DFP.BackBrush
If hBackBrush = 0 Then hBackBrush = comGDI.Brush_Face




'gdiGetObject DFP.hFont, Len(lf), lf
gdiGetObject DFP.CurrentFont, Len(lf), lf
'f0 = FontSRC(xfnt)

ar = Split(s, "¦") 'Делим на команды ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦ = Chr(&HA6)
Dim x&, y&, rc As RECT: rc = rc0
Dim sr() As RECT, sri&
arc = rc 'Начальные размеры требуемой области для рисования всего что прислано
prc = rc 'Начальный РЕКТ в котором надо рисовать
cr = rc 'Рект начального фрейма
ReDim sr(0)
sr(0) = rc

If bDraw Then
    SetBkMode xHDC, 1 'TRANSPARENT
    If hBackBrush Then FillRect xHDC, cr, hBackBrush   'Заливаем фон
End If 'DRAW

ic16 = menu16
'w.nc = 4: w0.nc = 4
w.nc = 4: w0 = w

n = UBound(ar)
If n > WM_MOUSEMOVE Then n = WM_MOUSEMOVE 'Ограничение на число команд
For i = 0 To n
gt = Mid$(ar(i), 3)

gt3 = Val(gt)
'gt3 = Cast(gt, vbLong)

gt0 = Left$(ar(i), 2)
If i = n Then gt0 = ""

Select Case gt0
Case "BB" 'ClearSCREEN with BRUSH
    If xHDC Then FillRect xHDC, rc, GlobalBrush(gt, DFP.BackColor, False)
    w.bg = gt
Case "BG": w.bg = gt 'BackGround

Case "BX": w.bx = gt3 'BackGround BRUSH ORIGN X
Case "BY": w.by = gt3 'BackGround BRUSH ORIGN Y

Case "TC": w.tc = gt 'TextColor
Case "PC": w.pc = gt: w.pw = 1 'PenColor +/-
Case "PW": w.pw = IF0(gt3, 1) 'PenWidth

Case "BS": w.bs = gt3 'BorderStyle
Case "DF": w.df = gt 'DrawFrameControl
Case "EN": cEnabled = gt3 'ENABLED STATE
Case "FF": w.ff = gt 'Font

Case "IC": w.ic = gt 'IconKey
Case "IA": w.icalign = gt3 'Global IconAlign 0 left+vtop 1 center ,2 right 4-vcenter 8 vbottom, 16=noMBdown
Case "IL": w.icleft = gt3 'IconLeft ofset
Case "IT": w.ictop = gt3 'IconTop ofset
Case "IS": If gt3 Then ic16 = sys16 Else ic16 = menu16

Case "SX": w.sx = gt 'StartX
Case "SY": w.sy = gt 'StartY
Case "DX": w.dx = gt3 'DrawStepX
Case "DY": w.dy = gt3 'DrawStepY
Case "RA": w.ra = gt 'Right Align FRAME
'Case "AY": w.ay = gt 'Align FRAME to Bottom BORDER

Case "DW": w.dw = gt  'DrawWidth
Case "DH": w.dh = gt  'DrawHeight

Case "RO": w.ro = gt3 'RectOfset
Case "RZ": w.ro = gt3 'RectSizeOfset
'Case "RI": w.ri = gt3 'InvertRect
Case "RF": w.fr = gt3 'FocusRect
Case "RR": w.rr = gt3 'RoundRect
'Case "RS": rsc = gt3 'Alpha Sel :-1=InvertRect, >0 ALPHA SEL
Case "HS": hsc = gt3 'NAME Alpha Sel :-1=InvertRect, >0 ALPHA SEL,-2=focusrect,-3=button,-4=border


Case "LH": w.lh = gt3 ': mLineHeight = gt 'LineHeight
Case "LF": w.lf = gt3 ' NextLine  OFFSET Y  if cr.Right+5>rc.Right

Case "FS" 'Frame SAVE  sr(sri)=cr
    sri = Abs(gt3)

Case "FR" 'Frame RESTORE cr=sr(sri)
    sri = -Abs(gt3)

Case "PD" 'POLYDRAW  M[oveTo],L[ineTo],B[ieserTo]
    w.pd = gt
    
Case "TA": w.ta = gt3: mTextAlign = w.ta 'TextAlign
Case "TT": w.tt = gt3 'TextTop Margin
Case "TL": w.tl = gt3 'TextLeft Margin
Case "TR": w.tr = gt3 'TextRight Margin
'Case "TB": w.tb = gt3 'TextBottom Margin

Case "NM": w.nm = gt 'FrameName
Case "NC": w.nc = gt 'FrameCursor

Case "RX"
    If gt3 = 0 Then rc = rc0: cr = rc
    LSet w = w0 'ResetFormat

Case "XX": If gt3 Then Exit For

Case "CF"  'Center Frame DH/DW exist
    If Len(w.dw) Then cr.Left = (rc.Right + rc.Left - Val(w.dw)) \ 2
    If Len(w.dh) Then cr.Top = (rc.Bottom + rc.Top - Val(w.dh)) \ 2
    
Case Else 'THERE TEXT = START DRAW

txt = Replace(ar(i), Chr$(0), "") 'TEXT for DRAW

'==========FONT

If Len(w.ff) Then hFont = GlobalFontSRC(lf, w.ff, 0): mLineHeight = 0 Else hFont = DFP.CurrentFont
If xHDC Then oldhFont = SelectObject(xHDC, hFont)
'==========FONT


If Len(w.lh) Then
    mLineHeight = Val(w.lh)
ElseIf mLineHeight = 0 Then
    mLineHeight = GetTextWidthHeight(hFont, "gM") + 1
End If

If sri < 0 Then
    If -sri <= UBound(sr) Then rc = cr: cr = sr(-sri) 'RESTORE RECT
    'sri = 0
End If

If Len(w.sx) Then 'StartDrawX
    x = Val(w.sx)
    If InStr(w.sx, " ") Then rc.Left = rc.Left + x: w.sx = "": x = 0
    cr.Left = rc.Left:  If x Then If x > 0 Then cr.Left = rc.Left + x Else cr.Left = rc.Right + x
'    If InStr(w.sx, " ") Then rc.Left = rc.Left + w.sx: w.sx = ""
'    cr.Left = rc.Left: If w.sx Then If w.sx > 0 Then cr.Left = rc.Left + w.sx Else cr.Left = rc.Right + w.sx

End If
If Len(w.sy) Then 'StartDrawY
    y = w.sy
    If InStr(w.sy, " ") Then rc.Top = rc.Top + y: w.sy = "": y = 0
    cr.Top = rc.Top: If y Then If y > 0 Then cr.Top = rc.Top + y Else cr.Top = rc.Bottom + y
'    If InStr(w.sy, " ") Then rc.Top = rc.Top + w.sy: w.sy = ""
'    cr.Top = rc.Top: If w.sy Then If w.sy > 0 Then cr.Top = rc.Top + w.sy Else cr.Top = rc.Bottom + w.sy

End If

'If w.dx Then cr.Left = cr.Left + w.dx   'StepDrawX
'If w.dy Then cr.Top = cr.Top + w.dy   'StepDrawY


ic_i = 0: ic_sz = 0
If Len(w.ic) Then 'Есть картинка
    ic = Val(w.ic): prc = cr
    If (Abs(ic) Mod 1000) < 500 Then
        ic_sz = 16: ic_iml = IIf(Abs(ic) < 999, ic16, sys16): ic_i = Abs(ic) Mod 1000
    ElseIf Abs(ic) < 600 Then
        ic_sz = 9: ic_iml = sm9: ic_i = Abs(ic) - 500
    End If
End If

If Len(w.dw) Then 'Ширина фрейма Как ширина текста
    If w.dw > 0 Then cr.Right = cr.Left + w.dw Else cr.Right = rc.Right + w.dw  'DrawWidth
ElseIf sri >= 0 Then
    cr.Right = cr.Left + w.tr + w.tl + ic_sz
    If Len(txt) Then cr.Right = cr.Right + GetTextWidthHeight(hFont, txt, 1) + 3
'    If rc.Right > 0 Then If cr.Right > rc.Right And w.lf = 0 Then cr.Right = rc.Right
End If

If Len(w.dh) Then 'Высота фрейма
    If w.dh > 0 Then cr.Bottom = cr.Top + w.dh Else cr.Bottom = rc.Bottom + w.dh 'DrawHeight
ElseIf sri >= 0 Then
    cr.Bottom = cr.Top + mLineHeight
    If Len(txt) Then cr.Bottom = cr.Bottom + mLineHeight * UBound(Split(txt, vbCrLf))
End If


If Len(w.ra) Then 'Right Align Frame
    x = Val(w.ra)
    If x <= 0 Then x = rc.Right + x
    If InStr(w.ra, " ") Then w.ra = (x - cr.Right + cr.Left + w.dx) & " " 'next frame align
    cr.Left = x - cr.Right + cr.Left: cr.Right = x
End If
'If Len(w.ay) Then 'BStartDrawY
'    y = Val(w.ay): If InStr(w.ay, " ") Then w.ay = (y - cr.Bottom + cr.Top) & " "
'    If y <= 0 Then y = rc.Bottom + y
'    cr.Top = y - cr.Bottom + cr.Top: cr.Bottom = y
'End If
OffsetRect cr, w.dx, w.dy


'If sri >= 0 Then
If w.lf Then  'ALLOW NEXT LINE
    If cr.Right > rc.Right Or w.lf < 0 Then
        If 5 * (cr.Right - rc.Right) > (rc.Right - cr.Left) Or w.lf < 0 Then 'Перенос на сл.строку
            OffsetRect cr, -cr.Left + w.dx + Val(w.sx), Abs(w.lf)
            rc.Top = rc.Top + Abs(w.lf)
        End If
    End If
End If
'End If

If sri < 0 Then sri = 0: rc = rc0

'If sri < 0 Then
'    If -sri <= UBound(sr) Then cr = sr(-sri) 'RESTORE RECT
'    sri = 0
'End If
'End If

If w.ro Then SetRect cr, cr.Left + w.ro, cr.Top + w.ro, cr.Right - w.ro, cr.Bottom - w.ro ': w.ro = 0 'OFSET ZOOM RECT


mOver = 0: bCFrame = 0

If Len(w.nm) Then
    i0 = InStrRev(w.nm, "/")
    txTool = "": If i0 Then txTool = Mid$(w.nm, i0 + 1): w.nm = Left$(w.nm, i0 - 1)
    If Len(TipText) > 0 And Len(txTool) > 0 Then
        toolrc = cr
        If NCdx Or NCdy Then OffsetRect toolrc, NCdx, NCdy
        ToolTipAppendTip xc.hWnd, txTool, toolrc
    End If
    If cEnabled Then
        bCFrame = True
        If PtInRect(cr, hit.Left, hit.Top) Then 'MOUSE OVER FRAME
            mOver = 1
            If Left$(w.ic, 1) = "-" And Left(w.nm, 1) <> "~" Then mOver = 0 Else DrawFormat = w.nm: nCursor = w.nc  'FRAME NAME + FRAME CURSOR
        End If
    End If
End If

If bDraw Then
If IntersectRect(prc, rc, cr) Then
'*************************************************************************
'If Len(w.bg) = 0 Then hBrush = hBackBrush Else hBrush = GlobalBrush(w.bg, DFP.BackColor, False) 'Получаем глобальный бруш ----> если [bc]=-1 то удалять не надо = это картинка
If Len(w.bg) = 0 Then
    hBrush = hBackBrush
Else
'    If w.bg = "-3" Then
'        hBrush = DFP.RowBrush
'    Else
        hBrush = GlobalBrush(w.bg, DFP.BackColor, False) 'Получаем глобальный бруш ----> если [bc]=-1 то удалять не надо = это картинка
'    End If
    If DFP.BackColor = -1 Then SetBrushOrgEx xHDC, 0, 0, pt: SetBrushOrgEx xHDC, w.bx, w.by, pt
    oldhBrush = SelectObject(xHDC, hBrush)

End If


Dim p&, pp, pa, pa0&, pa1&, pv$, ppt() As POINTAPI, pptp() As Byte
p = 0
If Len(w.pd) Then 'POLYDRAW= M[MoveTo]x,y + m[Moveto]dx,dy + L[LineTo]x,y + l[LineTo]dx,dy + C[BiezerTo]x,y + c[BeizerTo]dx,dy + ,[CLOSEFIGURE]
    pp = Split(w.pd, " ") 'M1,1 L3,4 3,4 2,4 4,5 4,6,
    p = UBound(pp) + 1
    ReDim ppt(p), pptp(p)
    ppt(0).x = cr.Left: ppt(0).y = cr.Top
    For p = 0 To UBound(pp)
        pa = Left$(pp(p), 1)
        If IsNumeric(pa) Or pa = "-" Then pa = Split(pp(p), ",") Else pv = pa: pa = Split(Mid(pp(p), 2), ",")
        pa0 = Val(pa(0)): pa1 = Val(pa(1))
        If pv = LCase(pv) Then 'RELATIVE at current point
            ppt(p + 1).x = ppt(0).x + pa0
            ppt(p + 1).y = ppt(0).y + pa1
        Else 'ABSOLUTE at start frame
            ppt(p + 1).x = cr.Left + pa0
            ppt(p + 1).y = cr.Top + pa1
        End If
        Select Case UCase(pv)
        Case "M": pptp(p + 1) = 6 'PT_MOVETO As Long = &H6
        Case "L": pptp(p + 1) = 2 'PT_LINETO As Long = &H2
        Case "C": pptp(p + 1) = 4 'PT_BEZIERTO As Long = &H4
        End Select
        If UBound(pa) > 1 Then pptp(p + 1) = pptp(p + 1) + 1 'PT_CLOSE As Long = &H1
        ppt(0) = ppt(p + 1) 'current point
        
    Next
    BeginPath xHDC
    PolyDraw xHDC, ppt(1), pptp(1), UBound(ppt)
    EndPath xHDC
    w.pd = ""
End If


If Len(w.df) Then   'DrawFrameControl
    'DF00 Пустая кнопка
    'DF10 Крест 'DF11 Минимайз  'DF13 Ресторез 'DF12 Максимайз 'DF14 Хелп
    'DF30 Вверх 'DF31 Вниз 'DF32 Влево 'DF33 Вправо
    'DF5(1,2,3,4) XPкнопка
    DrawFrameControl xHDC, cr, Left$(w.df, 1), CLng(Val(Mid$(w.df, 2))) + IIf(hit.Right > 0 And mOver, &H200, 0) + IIf(mOver, &H1000, 0) + IIf(Len(w.nm) And cEnabled, 0, &H100)
ElseIf Len(w.pc) Then 'Надо рисовать рамку с фоном или POLYDRAW
    ic = Val(w.pc)
    If ic < 0 Then ic = 0: If bCFrame Then ic = mOver    'HIGHLIGHT BORDER
    If w.pc = "0" Then ic = 1
    If ic Then
        oldhPen = SelectObject(xHDC, CreatePen(0, w.pw, AnyColor(Abs(Val(w.pc)))))
        
        If p Then
            If Len(w.bg) Then StrokeAndFillPath xHDC Else StrokePath xHDC
            p = 0
        ElseIf w.rr Then
            RoundRect xHDC, cr.Left, cr.Top, cr.Right, cr.Bottom, w.rr, w.rr
        Else
            Rectangle xHDC, cr.Left, cr.Top, cr.Right, cr.Bottom
        End If
        DeleteObject SelectObject(xHDC, oldhPen)
    End If
Else 'Залифаем фон фрейма
    prc = cr
    If w.bs = 0 Then prc.Bottom = prc.Bottom + 1: prc.Right = prc.Right + 1
    If Len(w.bg) Then FillRect xHDC, prc, 0
End If



If p Then p = 0: If Len(w.bg) Then StrokeAndFillPath xHDC Else StrokePath xHDC


    
If w.bs Then ' [- HIGHLIGHT][DownBS 0..F][-?High/Over BS 0..F]
    ic = Abs(w.bs)
    'If bCFrame Then If mOver Then ic = IIf(hit.Right, ic \ 16, ic) Else ic = IIf(w.bs < 0, 0, ic)
    If bCFrame Then If mOver Then ic = IIf(hit.Right, ic \ 16, ic)
    'ic = IIf(w.bs < 0 And Not mOver, 0, ic)
    If w.bs < 0 And Not mOver Then ic = 0
    If ic Then DrawEdge xHDC, cr, ic And 15&, &HF      'Рисуем рамку
End If
    
If w.fr Then
    ic = w.fr
    If ic < 0 And bCFrame Then ic = mOver 'HIGHLIGHT ФОКУСРЕКТ
    If ic Then DrawFocusRect xHDC, cr    'Рисуем ФОКУСРЕКТ
End If


If ic_i > 0 Then
    ic = Val(w.ic)
    If hit.Right > 0 And mOver And ((w.icalign And 16) <> 16) Then OffsetRect prc, 1, 1
    If Len(w.df) Then prc.Left = prc.Left + 2 'OffsetRect prc, 2, 0
    If w.icalign And 1 Then prc.Left = prc.Left + (prc.Right - prc.Left - ic_sz) / 2 'CENTER
    If w.icalign And 2 Then prc.Left = prc.Right - ic_sz 'RIGHT
    If w.icalign And 4 Then prc.Top = prc.Top + (prc.Bottom - prc.Top - ic_sz) / 2 'VCENTER
    If w.icalign And 8 Then prc.Top = prc.Bottom - ic_sz 'VBOTTOM
    'If Len(w.icleft) Then
    prc.Left = prc.Left + w.icleft
    'If Len(w.ictop) Then
    prc.Top = prc.Top + w.ictop
    If (ic < 0 Or Not cEnabled) And ic_i <> 2000 Then  'DISABLED
        DrawStateIcon xHDC, 0, 0, GetIcon(ic_i, ic_iml), 0, prc.Left, prc.Top, ic_sz, ic_sz, &H23&     'DST_ICON Or DSS_DISABLED
    Else
        ImageList_Draw ic_iml, ic_i, xHDC, prc.Left, prc.Top, 1
    End If
End If

If Len(txt) Then   'TEXT EXIST
    ta = 0: trc = cr 'TEXT FRAME
    If (mTextAlign And (DT_VCENTER Or DT_BOTTOM)) Then 'CALC VCENTER VBOTTOM
        If (trc.Bottom - trc.Top) > mLineHeight And (mTextAlign And DT_SINGLELINE) = 0 Then 'NO POSIBLE CALC VCENTER VBOTTOM
            c = DrawText(xHDC, txt, Len(txt), trc, mTextAlign Or DT_NOPREFIX Or DT_EXPANDTABS Or DT_CALCRECT)   'Return Height of text
            trc = cr
'            If c < cr.Bottom - cr.Top Then 'Текст не вылезет за пределы РЕКТА
                If (mTextAlign And DT_VCENTER) Then 'VCENTER
                    trc.Top = (cr.Bottom + cr.Top - c) / 2
                ElseIf (mTextAlign And DT_BOTTOM) Then 'VBOTTOM
                    trc.Top = cr.Bottom - c
                End If
                ta = mTextAlign 'Or DT_WORDBREAK Or DT_EDITCONTROL 'Добавим перенос слов при отрисовке
'            Else
'                ta = ta And Not (DT_VCENTER Or DT_BOTTOM)
'            End If
        End If
    End If
    trc.Left = trc.Left + 1
    If ta = 0 Then ta = mTextAlign 'And Not (DT_WORDBREAK Or DT_EDITCONTROL))
    If w.tl Then trc.Left = trc.Left + w.tl  'TEXT LEFT MARGIN
    If w.tr Then trc.Right = trc.Right - w.tr  'TEXT RIGHT MARGIN
    'If w.bs > 0 Or Len(w.df) Then trc.Left = trc.Left + 2 'TEXT LEFT FOR BUTTON/FRAME
    If w.tt Then trc.Top = trc.Top + w.tt   'TEXT TOP MARGIN
    'If w.tb Then trc.Top = trc.Bottom - w.tb   'TEXT BOTTOM MARGIN
    
    If ic_i Then 'text with icon
        ic_sz = ic_sz + w.icleft + 1
        If w.icalign And 2 Then trc.Right = trc.Right - ic_sz Else trc.Left = trc.Left + ic_sz
    End If
    'If mOver And hit.Right And w.bs = 0 Then trc.Left = trc.Left + 1: trc.Top = trc.Top + 1 'Смещение нажатой надписи
    If mOver And hit.Right Then trc.Left = trc.Left + 1: trc.Top = trc.Top + 1  'Смещение нажатой надписи
    
    If Not cEnabled Then 'And Len(w.nm) Then
        SetTextColor xHDC, AnyColor(17)  'DISBLAED TEXT
    ElseIf Len(w.tc) Then   'TEXT COLOR
        SetTextColor xHDC, AnyColor(Val(w.tc)) 'CLng(w.tc)
    Else 'DEFAULT COLOR
        'If DFP.CurrentForeColor > 0 And DFP.CurrentForeColor < 32 Then DFP.CurrentForeColor = GetSysColor(DFP.CurrentForeColor)
        'SetTextColor xHDC, DFP.CurrentForeColor
        SetTextColor xHDC, AnyColor(DFP.CurrentForeColor)
    End If

    'If (trc.Bottom - trc.Top) > DFP.FontHeight * 1.5 Then If InStr(txt, vbCr) = 0 Then mTextAlign = DT_SINGLELINE Or (mTextAlign And (Not &H2010))
    
    DrawText xHDC, txt, Len(txt), trc, ta Or DT_NOPREFIX Or DT_EXPANDTABS 'Or DT_NOCLIP
    If oldhFont Then SelectObject xHDC, oldhFont: oldhFont = 0
End If 'TEXT EXIST

hBrush = SelectObject(xHDC, oldhBrush)
'If w.ri Then InvertRect xHDC, cr: w.ri = 0
If hsc And cEnabled Then
    If mOver Then
        If hsc = -1 Then
            InvertRect xHDC, cr
        ElseIf hsc > 0 Then
            hsc = AnyColor(hsc)
            oldhBrush = SelectObject(xHDC, CreateSolidBrush(hsc)) 'SelectObject(xHDC, GlobalBrush(rsc, 0))
            BitBlt xHDC, cr.Left, cr.Top, cr.Right - cr.Left, cr.Bottom - cr.Top, xHDC, cr.Left, cr.Top, &HE20746
            DeleteObject SelectObject(xHDC, oldhBrush)
        End If
        hsc = 0
    End If
End If

'*************************************************************************
End If 'IntersectRect
'If Len(w.nm) And Exclude Then ExcludeClipRect xHDC, cr.Left, cr.Top, cr.Right, cr.Bottom


End If 'DRAW

If xHDC = 0 Or hit.Bottom = 3 Then
    If arc.Left > cr.Left Then arc.Left = cr.Left
    If arc.Top > cr.Top Then arc.Top = cr.Top
    If arc.Right < cr.Right Then arc.Right = cr.Right + 1
    If arc.Bottom < cr.Bottom Then arc.Bottom = cr.Bottom + 1
End If


If sri > 0 Then 'SAVE RECT
    If UBound(sr) < sri Then ReDim Preserve sr(sri)
    sr(sri) = cr: sri = 0
End If

If w.ro Then SetRect cr, cr.Left - w.ro, cr.Top - w.ro, cr.Right + w.ro, cr.Bottom + w.ro: w.ro = 0  'OFSET ZOOM RECT
prc = cr 'Запоминаем координаты нарисованного фрейма
cr.Left = cr.Right  'Начало следующего фрейма

w.nm = "": w.ic = ""
cEnabled = DFP.Enabled
End Select
Next

'If Not DFP.NCEnabled Then
If Not cEnabled Then

DisableDC xHDC, rc

End If
'ExcludeClipRect hDC, rc.Left, rc.Bottom + 1, rc.Right, rc.Bottom + m_RH
'If xHDC Then SelectClipRgn xHDC, 0


hit = arc

End Function


Sub DisableDC(ByVal hDC&, rc As RECT) ', Optional bNC As Boolean)
'EmbosseDC hDC, rc
'Dim hBM&
'hBM = GetHBitmap(hDC, rc.Right - rc.Left, rc.Bottom - rc.Top)
'DitherBlt hDC, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, hBM, 0, 0
'DeleteObject hBM
'Exit Sub
If hDC = 0 Then Exit Sub
Dim h&
Dim br&
'br = CreateSolidBrush(&H818181)
br = CreateSolidBrush(gray_color)
'h = SelectObject(hDC, comGDI.Brush_Gray)
h = SelectObject(hDC, br)
BitBlt hDC, rc.Left, rc.Top, rc.Right, rc.Bottom, hDC, rc.Left, rc.Top, &HFC008A
'SetBkColor hDC, GetPixel(hDC, 0, 0)
'If Not bNC Then DrawParams.DisabledBackColor = GetPixel(hDC, rc.Left, rc.Top): DrawParams.DisabledBrush = GlobalBrush("" & DrawParams.DisabledBackColor, 0&)
h = SelectObject(hDC, h)
DeleteObject br
End Sub



'================ DRAW FORMAT ==================

'================ DRAW DOT LINE ==================
Public Sub TreeLine(ByVal hDC As Long, x&, y&, cx&, cy&, ByVal tp&)
Dim hOldPen As Long
hOldPen = SelectObject(hDC, comGDI.Pen_Gray)
Dim pt(1) As POINTAPI
If tp And 1 Then pt(0).x = x + cx \ 2: pt(0).y = y: pt(1).x = pt(0).x: pt(1).y = y + cy: Polyline hDC, pt(0), 2
If tp = 2 Or tp = 3 Then pt(0).x = x + cx \ 2: pt(0).y = y + cy \ 2: pt(1).x = x + cx: pt(1).y = pt(0).y: Polyline hDC, pt(0), 2
If tp = 2 Then pt(0).x = x + cx \ 2: pt(0).y = y: pt(1).x = pt(0).x: pt(1).y = y + cy \ 2: Polyline hDC, pt(0), 2
SelectObject hDC, hOldPen
End Sub
'================ DRAW DOT LINE ==================





'============= HORD KEEPER ============================
'Public Sub AddHord(ByVal hWndMaster$, ByVal MasterB$, ByVal hWndSlave&, ByVal SlaveB$, ByVal nMin$, ByVal nMax$, ByVal nDelta$, ByVal nValue$, hSlave As xControl)
Public Sub AddHord(ByVal sMaster$, ByVal MasterB$, ByVal SlaveB$, ByVal nMin$, ByVal nMax$, ByVal nDelta$, ByVal nValue$, xSlave As xControl)
Dim h, hMaster&
'Проверим совместимость границ мастера и раба
'MasterB = Left$(MasterB, 1): MasterB = IIf(InStr(1, "LTRBXYWH", MasterB, 1) = 0, "", MasterB)
'SlaveB = Left$(SlaveB, 1): SlaveB = IIf(InStr(1, "LTRBXYWH", SlaveB, 1) = 0, "", SlaveB)
If xSlave.hWnd = 0 Or MasterB = "" Or SlaveB = "" Then Exit Sub
If sMaster = "Parent" Then
    hMaster = xSlave.Parent.hWnd
ElseIf sMaster = "Owner" Then
    hMaster = xSlave.Owner.hWnd
ElseIf sMaster = "Me" Then
    hMaster = xSlave.hWnd
End If
If hMaster Then sMaster = "M" & hMaster
'keySlave = "S" & Join(Array(xSlave.hWnd, MasterB, SlaveB), "|")
h = Replace(Join(Array(nMin, nMax, nDelta, nValue), "|"), ";", ",")
If Len(h) = 3 Then h = Null 'удаление хорды
hrds(sMaster & "\S" & Join(Array(xSlave.hWnd, MasterB, SlaveB), "|")) = h
If hMaster = 0 Then hrds.Name = "NoHWND"
End Sub

Public Sub RemoveHord(hWnd&)
'Debug.Print "RemoveHord " & Hex$(hWnd)
hrds.Param("M" & hWnd) = Null  'Удаляем мастера
'Удаляем рабов

Dim i&, m, mi&, s, si&
For i = 0 To 1 '2 прохода 0=удаление рабов 1=удаление пустых мастеров
m = hrds.Value
For mi = 0 To UBound(m) Step 2
    s = m(mi + 1)
    If IsArray(s) Then
        If i = 0 Then
            For si = 0 To UBound(s) Step 2
                If InStr(s(si), "S" & hWnd) = 1 Then hrds(m(mi) & "\" & s(si)) = Null
            Next
        End If
    Else
        hrds(m(mi)) = Null
    End If
Next
Next

'Dim i&, n&, ar, mn$
'For i = 0 To hrds.ParamCount - 1
'    mn = hrds.ParamName(i) 'Имя мастера
'    ar = hrds.Param(i) 'Массив рабов
'    If IsArray(ar) Then
'        For n = 0 To UBound(ar) Step 2 'Гоним по рабам ищем удаляемого
'            If Split(ar(n), "|")(0) = "S" & hWnd Then hrds(mn & "\" & ar(n)) = Null              'Найден
'        Next
'    Else
'        hrds.Param(i) = Null
'    End If
'Next
'i = 0
'While i < hrds.ParamCount 'Удаляем мастеров без рабов
'    If IsArray(hrds.Param(i)) Then
'        If UBound(hrds.Param(i)) > -1 Then i = i + 1 Else hrds.Param(i) = Null
'    Else
'        hrds.Param(i) = Null
'    End If
'Wend

End Sub

Public Sub LookUpHords(Optional hWnd&)
Dim i&, n, h&
n = EnumSubClass.Value
For i = 0 To UBound(n) Step 2
    h = n(i + 1)
    If hWnd Then 'Двинуть папу и братьев
        If h = Get_Parent(hWnd) Or Get_Parent(h) = Get_Parent(hWnd) Then LookUpHord h
    Else 'Двинуть всех
        LookUpHord h
    End If
Next
End Sub

Public Sub LookUpHord(ByVal hWndMaster&, Optional ByVal hwslave&)
Dim i&, nm$, x As xControl
Dim hWndSlave&, ar, sar
'If HordsOff Or IsIconic(hWndMaster) Then Exit Sub
If IsIconic(hWndMaster) Then Exit Sub
Set x = hxControl(hWndMaster)
'Debug.Print LookUpHord & " " & hWndMaster

If hrds.Name = "NoHWND" Then 'Ищем окно без hWnd по имени
    'On Error Resume Next
    nm = x.Name
    Dim cp As New CParam
    i = hrds.GetIndex(nm)
    If i > -1 Then 'Проверим чилда на предмет родства
        hrds.Name = "Hords"
        ar = hrds.Param(i)
        hWndSlave = Mid$(Split(ar(0), "|")(0), 2) 'Берем hWnd первого детя
        If Get_Parent(hWndSlave) = hWndMaster Or Get_Parent(hWndSlave) = Get_Parent(hWndMaster) Or hWndSlave = Get_Parent(hWndMaster) Then
            If hrds.GetIndex("M" & hWndMaster) = -1 Then 'Окно мастера еще не прописано
                hrds.ParamName(i) = "M" & hWndMaster 'Регистрируем номер окна вместо имени
            Else 'Детей предать в тово который есть
                'СТАРЫЕ ДЕТИ m1 hrds.Param("" & hWndMaster)
                'НОВЫЕ ДЕТИ m2 hrds.Param(i)
                'Надо их объеденить m1=m1+m2
                cp.Value = hrds.Param("M" & hWndMaster) 'Берем старых детей
                cp.AddParams ar 'Добавляем новых детей
                hrds.Param("M" & hWndMaster) = cp.Source 'Получаем обновленных детей
                hrds.Param(i) = Null 'Удаляем нового мастера
            End If
        'Else
            'gDebugPrint "NO POSITION Master=" & nm, ">>> Slave=" & hxControl(hWndSlave).Name
        End If
        For i = 0 To hrds.ParamCount - 1
            If Val(Mid$(hrds.ParamName(i), 2)) = 0 Then hrds.Name = "NoHWND": Exit For
        Next
    End If

'Err.Clear
End If
'If hrds.Name = "NoHWND" Then Debug.Print "LookUpHord ={" & hrds.Names & "}"

'i = hrds.GetIndex("M" & hWndMaster)
'If i = -1 Then Exit Sub
'ar = hrds.Param(i) 'Получаем список рабов и их LAYOUT

ar = hrds("M" & hWndMaster) 'Получаем список рабов и их LAYOUT
If Not IsArray(ar) Then Exit Sub

sar = Array() 'список рабов
For i = 0 To UBound(ar) Step 2 'Групируем рабов
    hWndSlave = Mid$(Split(ar(i), "|")(0), 2)
    If IsWindow(hWndSlave) Then
        If hwslave = 0 Or hWndSlave = hwslave Then
        gAddIndex sar, hWndSlave 'Список рабов
        End If
    End If
Next

'If hwslave Then sar = Array(hwslave)

'Dim Mpapa As Boolean
Dim si&, rcMaster As RECT, rcSlave As RECT, rc As RECT
Dim mcrcMaster As RECT, mcrcSlave As RECT
Dim sp, mb$, sb$, nMin&, nMax&, nDelta&, nValue&
Dim bm&, bs&, wh& 'Границы мастера и раба
Const c9% = -9999

Dim x0 As xControl, bmrc As Boolean

For si = 0 To UBound(sar) 'По всем рабам поехали
hWndSlave = sar(si) 'Окно раба

'Mpapa = (hWndMaster = GetParent(hWndSlave)) 'Отношение к мастеру = мастер : папа или брат
Dim ptm As POINTAPI, pts As POINTAPI, pt0 As POINTAPI
ptm = pt0: pts = pt0
If (hWndMaster = Get_Parent(hWndSlave)) Then 'Мастер = ПАПА берем высоту и ширину МАСТЕРА
    GetClientRect hWndMaster, rcMaster
Else 'Мастер = Брат берем РЕКТ МАСТЕРА на папе
    GetWindowRect hWndMaster, rcMaster
    ScreenToClient Get_Parent(hWndMaster), ptm
    OffsetRect rcMaster, ptm.x, ptm.y
End If

'Берем MARGINS рект папы
x.GetRect mcrcMaster, 2
'Берем MARGINS рект раба
Set x0 = Nothing
Set x0 = hxControl(hWndSlave)
If x0.hWnd = hWndSlave Then 'SLAVE EXIST
    x0.GetRect mcrcSlave, 2


'Берем рект РАБА на папе
GetWindowRect hWndSlave, rcSlave
ScreenToClient Get_Parent(hWndSlave), pts
OffsetRect rcSlave, pts.x, pts.y

'If nn = "TX" Then Debug.Print "===" & StrRect(rcSlave)
rc = rcSlave
bmrc = 0


'Надо отфильтровать то что касается текущего раба
For i = 0 To UBound(ar) Step 2

sp = Split(Mid$(ar(i), 2), "|")
If sp(0) = hWndSlave Then
mb = sp(1): sb = sp(2)
sp = Split(ar(i + 1) & "||||", "|")

'Err.Clear pnMin&, pnMax&, pnDelta&
If IsNumeric(sp(0)) Or sp(0) = "" Then nMin = NzS(sp(0), c9) Else nMin = x.Eval(sp(0), c9)
If IsNumeric(sp(1)) Or sp(1) = "" Then nMax = NzS(sp(1), c9) Else nMax = x.Eval(sp(1), c9)
If IsNumeric(sp(2)) Or sp(2) = "" Then nDelta = NzS(sp(2), c9) Else nDelta = x.Eval(sp(2), c9)
If IsNumeric(sp(3)) Or sp(3) = "" Then nValue = NzS(sp(3), c9) Else nValue = x.Eval(sp(3), c9)
'%50%
'If hxControl(hWndSlave).Name = "BSplit" Then Debug.Print "LookUpHord=" & hxControl(hWndSlave).Name & " " & ar(i + 0) & " " & ar(i + 1) & "  = " & nDelta

'If Err Then rc = rcSlave: Exit For

Select Case mb 'MasterBorder Name
Case "L", "X": bm = rcMaster.Left
Case "T", "Y": bm = rcMaster.Top
Case "R": bm = rcMaster.Right
Case "B": bm = rcMaster.Bottom
Case "W": bm = rcMaster.Right - rcMaster.Left
Case "H": bm = rcMaster.Bottom - rcMaster.Top
Case "ML": bm = mcrcMaster.Left
Case "MT": bm = mcrcMaster.Top
Case "MR": bm = mcrcMaster.Right
Case "MB": bm = mcrcMaster.Bottom
Case "SR": bm = rcMaster.Right - rcMaster.Left - mcrcMaster.Right
Case "SB": bm = rcMaster.Bottom - rcMaster.Top - mcrcMaster.Bottom
Case Else: bm = 0
End Select
'Debug.Print "master", mb, bm

Select Case sb 'SlaveBorder Name
Case "L", "X": bs = rcSlave.Left
Case "T", "Y": bs = rcSlave.Top
Case "R": bs = rcSlave.Right
Case "B": bs = rcSlave.Bottom
Case "W": bs = rcSlave.Right - rcSlave.Left
Case "H": bs = rcSlave.Bottom - rcSlave.Top
Case "ML": bs = mcrcSlave.Left
Case "MT": bs = mcrcSlave.Top
Case "MR": bs = mcrcSlave.Right
Case "MB": bs = mcrcSlave.Bottom
Case "SR": bs = rcSlave.Right - rcSlave.Left - mcrcSlave.Right
Case "SB": bs = rcSlave.Bottom - rcSlave.Top - mcrcSlave.Bottom
Case Else: bs = 0
End Select

'Debug.Print "slave", sb, bs

bs = bm + IIf(nDelta = c9, 0, nDelta)
If Not nValue = c9 Then bs = nValue
If Not nMin = c9 Then If bs < nMin Then bs = nMin
If Not nMax = c9 Then If bs > nMax Then bs = nMax
'Debug.Print "nMin=" & nMin, "nMax=" & nMax, "nDelta=" & nDelta, "nValue=" & nValue

'Debug.Print "slave", sb, bs


Select Case sb
Case "L": rc.Left = bs
Case "T": rc.Top = bs
Case "R": rc.Right = bs
Case "B": rc.Bottom = bs
Case "X": wh = rc.Right - rc.Left: rc.Left = bs: rc.Right = rc.Left + wh
Case "Y": wh = rc.Bottom - rc.Top: rc.Top = bs: rc.Bottom = rc.Top + wh
Case "W": rc.Right = rc.Left + bs
Case "H": rc.Bottom = rc.Top + bs
Case "ML":  bmrc_ mcrcSlave.Left, bs, bmrc 'bmrc = bmrc Or (mcrcSlave.Left <> bs): mcrcSlave.Left = bs
Case "MT": bmrc_ mcrcSlave.Top, bs, bmrc  'bmrc = bmrc Or (mcrcSlave.Top <> bs): mcrcSlave.Top = bs
Case "MR": bmrc_ mcrcSlave.Right, bs, bmrc 'bmrc = bmrc Or (mcrcSlave.Right <> bs): mcrcSlave.Right = bs
Case "MB": bmrc_ mcrcSlave.Bottom, bs, bmrc  'bmrc = bmrc Or (mcrcSlave.Bottom <> bs): mcrcSlave.Bottom = bs
Case "SR": bmrc_ mcrcSlave.Right, rcSlave.Right - rcSlave.Left - bs, bmrc    'bmrc = bmrc Or (mcrcSlave.Right <> rcSlave.Right - rcSlave.Left - bs): mcrcSlave.Right = rcSlave.Right - rcSlave.Left - bs
Case "SB": bmrc_ mcrcSlave.Bottom, rcSlave.Bottom - rcSlave.Top - bs, bmrc   'bmrc = bmrc Or (mcrcSlave.Bottom <> rcSlave.Bottom - rcSlave.Top - bs): mcrcSlave.Bottom = rcSlave.Bottom - rcSlave.Top - bs
End Select

'Debug.Print hxControl(hWndSlave).Name & "." & sB & "=" & ar(i + 1) & " = " & bs

End If
Next

'MCRC corrector
If bmrc Then x0.GetRect mcrcSlave, 3


If rc.Left > rc.Right Then rc.Right = rc.Left
If rc.Top > rc.Bottom Then rc.Bottom = rc.Top

'Все готово для сдвига раба = новая позиция -> [rc]
If EqualRect(rcSlave, rc) = 0 Then  'Если ректы отличаются то двигаем
'Debug.Print "LHORD", hxControl(hWndSlave).Name
    SetWindowPos hWndSlave, 0, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, &H634 ' Or SWP_NOACTIVATE   ' Or &H400
End If
'Debug.Print "LookUpHord " & hxControl(hWndSlave).Name

'Else 'SLAVE NOT !!!!!!!!!!!!!!!!!!!!!! EXIST
'Debug.Print ">>>>>>>>>SLAVE NOT !!!!!!!!!!!!!!!!!!!!!! EXIST"
'Debug.Print ">>>>>>>>>SLAVE NOT !!!!!!!!!!!!!!!!!!!!!! EXIST"
'Debug.Print ">>>>>>>>>SLAVE NOT !!!!!!!!!!!!!!!!!!!!!! EXIST"
End If 'SLAVE EXIST

    'Debug.Print "LookUpHord " & X.Name
Next 'Следующий раб
Set x = Nothing
End Sub

Private Sub bmrc_(v, nv, b As Boolean)
b = b Or v <> nv: v = nv
End Sub
'============= HORD KEEPER ============================

'Function Rich_TAB(hWnd&, nGetShiftState&) As Boolean
'Dim p&, p1&, p2&, s$, v, i&, doc As ITextDocument
'Set doc = ITextDocument(hWnd)
'If doc.Selection.End <= doc.Selection.Start Then Exit Function
'doc.Selection.Expand tomLine
'p1 = doc.Selection.Start
'p2 = doc.Selection.End - 1
'
's = doc.range(p1, p2).Text
'v = Split(s, vbCr)
'For i = 0 To UBound(v)
'    If nGetShiftState = 1 Then 'Remove Line Tab
'        If Len(v(i)) Then If Left(v(i), 1) = vbTab Or Left(v(i), 1) = " " Then v(i) = Mid(v(i), 2): p = p - 1
'    Else 'Add Line Tab
'         v(i) = vbTab & v(i): p = p + 1
'    End If
'Next
'If s <> Join(v, vbCr) Then
'    doc.Freeze
'    doc.range(p1, p2).Text = Join(v, vbCr)
'    doc.range(p1, p2 + p).Select
'    doc.Unfreeze
'End If
'Rich_TAB = True
'End Function



'================ F O N T ================================
Public Function GlobalFontSRC&(ByRef p As LOGFONT, src$, Optional bAdd As Boolean = True) 'NAME,SIZE,BOLD,WIDTH

Dim f(6), ar, a, i&
ar = Split(src, ",")
For Each a In ar
If IsNumeric(a) Then 'SIZE or WIDTH
    If IsEmpty(f(1)) Then 'SIZE
        f(1) = L_(Left("" & Abs(L_(a)), 3))
    ElseIf IsEmpty(f(6)) Then  'WIDTH
        f(6) = L_(Left("" & Abs(L_(a)), 2))
    End If
Else 'NAMEor BOLD or ITALIC
    If Len(a) = 1 Then 'BOLD/ITALIC
        'If IsEmpty(f(1)) Then f(1) = MulDiv(-p.lfHeight, 72, 96) 'DEFAULT SIZE
        If a = "X" Then f(1) = Empty: f(2) = 0: f(3) = 0: f(4) = 0: f(5) = 0: f(6) = 0
        If a = "B" Then f(2) = 1
        If a = "I" Then f(3) = 1
        If a = "U" Then f(4) = 1
        If a = "S" Then f(5) = 1
    ElseIf Len(a) = 0 And IsEmpty(f(0)) Then  'Name=Default
        'f(0) = RTrimNull$(StrConv(p.lfFaceName, vbUnicode))
        f(0) = pTrimS(StrConv(p.lfFaceName, vbUnicode), Chr(0), 2) 'RTRIM$0
        
    ElseIf IsEmpty(f(0)) Then   'Name
        f(0) = Left(a, 31)
    End If
End If
i = i + 1
If i = 8 Then Exit For
Next
'If IsEmpty(f(0)) Then f(0) = RTrimNull$(StrConv(p.lfFaceName, vbUnicode))
If IsEmpty(f(0)) Then f(0) = pTrimS(StrConv(p.lfFaceName, vbUnicode), Chr(0), 2)
If IsEmpty(f(1)) Then f(1) = MulDiv(-p.lfHeight, 72, 96)
If IsEmpty(f(2)) Then f(2) = IIf(p.lfWeight >= 700, 1, 0)
If IsEmpty(f(3)) Then f(3) = IIf(p.lfItalic, 1, 0)
If IsEmpty(f(4)) Then f(4) = IIf(p.lfUnderline, 1, 0)
If IsEmpty(f(5)) Then f(5) = IIf(p.lfStrikeOut, 1, 0)
If IsEmpty(f(6)) Then f(6) = p.lfWidth
GlobalFontSRC = GlobalFont(f(0) & "", 1 * f(1), 1 * f(2), 1 * f(3), 1 * f(4), 1 * f(5), 1 * f(6), bAdd)
End Function
Public Function GlobalFont&(Optional Name$ = "Tahoma", Optional nSize& = 8, Optional bBold As Boolean, Optional bItalic As Boolean, Optional bUnderLine As Boolean, Optional bStrikeOut As Boolean, Optional nWidth&, Optional bAdd As Boolean = True)
Dim f$, n&
f = Name & IIf(nSize, "," & nSize, "") & IIf(bBold, ",B", "") & IIf(bItalic, ",I", "") & IIf(bUnderLine, ",U", "") & IIf(bStrikeOut, ",S", "") & IIf(nWidth, "," & nWidth, "")
GlobalFont = GlobalGDI.ParamDef("font_" & f, 0&)
If GlobalFont <> 0 Then
    If bAdd Then n = GlobalGDICount.ParamDef("font_" & f, 0&)  'Количество использований этого фонта
Else
    Dim lf As LOGFONT, b() As Byte
    b = StrConv(Name & String$(31, 0), vbFromUnicode)
    CopyMemory lf.lfFaceName(0), b(0), 32
    lf.lfHeight = -MulDiv(nSize, 96, 72)
    lf.lfWeight = IIf(bBold, 700, 400)
    lf.lfItalic = IIf(bItalic, 1, 0)
    lf.lfUnderline = IIf(bUnderLine, 1, 0)
    lf.lfStrikeOut = IIf(bStrikeOut, 1, 0)
    lf.lfWidth = nWidth
    lf.lfCharSet = 104
    GlobalFont = CreateFontIndirect(lf) 'Создаем фонт
    If GlobalFont <> 0 Then GlobalGDI.Param("font_" & f) = GlobalFont 'Записываем в его в GlobalGDI
End If
If GlobalFont <> 0 And bAdd Then GlobalGDICount.Param("font_" & f) = n + 1 'Если есть такой фонт то пропишем количество его использований
End Function
Public Function GetFontSRC$(hFont&)
Dim lf As LOGFONT
If hFont = 0 Then Exit Function
gdiGetObject hFont, Len(lf), lf
GetFontSRC = LogFontSRC(lf)
End Function
Public Function LogFontSRC$(lf As LOGFONT)
LogFontSRC = pTrimS(StrConv(lf.lfFaceName, vbUnicode), Chr(0), 2) & "," & MulDiv(-lf.lfHeight, 72, 96) & IIf(lf.lfWeight >= 700, ",B", "") & IIf(lf.lfItalic, ",I", "") & IIf(lf.lfUnderline, ",U", "") & IIf(lf.lfStrikeOut, ",S", "") & IIf(lf.lfWidth, "," & lf.lfWidth, "")
End Function
'================ F O N T ================================

Function GlobalPen&(PenColor&, Optional pw&)
Dim f$, n&, src$
f = Hex$(AnyColor(PenColor)) & IIf(pw > 0, "." & pw, "")
src = "pen_" & f
GlobalPen = GlobalGDI.ParamDef(src, 0&)
If GlobalPen Then
    n = GlobalGDICount.Param(src)
    GlobalGDICount.Param(src) = n + 1
Else
    GlobalPen = CreatePen(0, pw, AnyColor(PenColor))
    GlobalGDI.Param(src) = GlobalPen 'Если создали то пропишем его в GlobalGDI
    GlobalGDICount.Param(src) = 1
End If
End Function

Function GDIP_Image&(sFile$)
'Dim GSI As GdiplusStartupInput
'GSI.GdiplusVersion = 1
'GdiplusStartup m_GDIP_Token, GSI

Dim n&, src$, m_Image&
src = "image_" & sFile
If Len(sFile) = 0 Then Exit Function
If m_GDIP_Token = 0 Then Exit Function
GDIP_Image = GlobalGDI.ParamDef(src, 0&)
If GDIP_Image = -1 Then GDIP_Image = 0: Exit Function
If GDIP_Image = 0 Then
    If IsNum(sFile) Then m_Image = sFile Else GdipLoadImageFromFile StrPtr(sFile), m_Image
    GDIP_Image = m_Image
    If m_Image Then
        GlobalGDI.Param(src) = m_Image
        GlobalGDICount.Param(src) = 1
    End If
Else
    n = GlobalGDICount.Param(src)
    GlobalGDICount.Param(src) = n + 1
End If
End Function

Function GlobalBrush&(str_brush$, back_color&, Optional ByVal bAdd As Boolean = True)
Dim f$, n&, lb As LOGBRUSH, bc&
Dim src$
f = str_brush
'If f = "p" Then f = App.Path & "\want.jpg"

'GlobalBrush = 0 '16
If Len(f) = 0 Then Exit Function
If f = "Null" Then Exit Function
If f = Chr(0) Then Exit Function
'back_color = 0


src = "brush_" & Replace(str_brush, "\", "_")
On Error Resume Next
GlobalBrush = GlobalGDI.ParamDef(src, 0&)
If GlobalBrush Then 'Вернуть цвет фона
    gdiGetObject GlobalBrush, Len(lb), lb
    back_color = lb.lbColor
    If Not lb.lbStyle = 0 Then back_color = -1
    If bAdd Then
        n = GlobalGDICount.ParamDef(src, 0)
        GlobalGDICount.Param(src) = n + 1
    End If
    Exit Function
End If

If IsNumeric(f) Then 'color
    bc = f
    back_color = AnyColor(bc)
    GlobalBrush = CreateSolidBrush(back_color)
    
Else 'string

    On Error Resume Next
    If Left(f, 1) = "*" Then 'ScreenShoot
        GlobalBrush = CreatePatternBrush(back_color)

    
    ElseIf InStr(f, "\") > 0 Or InStr(f, "/") > 0 Then
        'GlobalBrush = CreatePatternBrush(LoadPicture(f))
        'n = LoadPictureEx(f)
        GlobalBrush = CreatePatternBrush(LoadPictureEx(f))
        'DeleteObject n
    
    Else
        f = S_(xMain.Param("Brush\" & f))
        If Len(f) Then
            Dim hmemDC&, dp As xDRAWPARAMS, rc As RECT, RCP As RECT
            'xMainWnd.Controls(1).GetDrawParams VarPtr(dp), ""
            dp.Enabled = 1
            DrawFormat xMainWnd, 0, rc, f, RCP, dp 'Calc SIZE
            rc = RCP
            CreateMemDC hmemDC, RCP.Right, RCP.Bottom
            DrawFormat xMainWnd, hmemDC, rc, f, RCP, dp
            Dim hOldBmp&, hbmp&
            hOldBmp = CreateCompatibleBitmap(hmemDC, 1, 1)  'RCP.Right, RCP.Bottom)
            hbmp = SelectObject(hmemDC, hOldBmp)
            GlobalBrush = CreatePatternBrush(hbmp)
            hOldBmp = SelectObject(hmemDC, hbmp) 'Берем старую картинку из DC
            DeleteObject hOldBmp
            CreateMemDC hmemDC, 0, 0
        End If
    End If
    back_color = -1 'NONE

End If

If GlobalBrush Then  'Если создали то пропишем его в GlobalGDI
    GlobalGDI.Param(src) = GlobalBrush
    GlobalGDICount.Param(src) = 1
End If
End Function



Sub ClearGDI(str_gdi$)
Dim i&, n&, s$
If Len(str_gdi) = 0 Then 'Всех удаляем
    DeleteDC tempHDC
    While GlobalGDI.ParamCount > 0
        n = GlobalGDI.ParamDef(0, 0&)
        s = GlobalGDI.ParamName(0)
        If Left(s, 5) = "image" And n Then
            GdipDisposeImage n  ' m_Image
        Else
            If n And &HFFFE0 Then DeleteObject n
        End If
        GlobalGDI.Param(i) = Null
    Wend
    'Debug.Print "ClearGDI ALL"
'    DeleteObject comGDI.Pen_Gray
'    ImageList_AddIcons cur32
'    ImageList_AddIcons menu16
'    ImageList_AddIcons sm9
'    ImageList_AddIcons sys16

'Debug.Assert Abs(UBound(ic32ar)) < 100
    On Error Resume Next

    Dim Item, ar
    
    For Each Item In Array(cur32, menu16, sm9) ', sys16)
        If Item Then ImageList_AddIcons 0 + Item
    Next
    cur32 = 0: menu16 = 0: sm9 = 0: sys16 = 0
    
    Dim dc: dc = Array()
    #If DragDrop Then
    dc = DragCursor
    ImageList_Destroy g_DragIML
    #End If
    For Each ar In Array(icar(0), icar(2), icar(3), icar(4), dc) 'Array(cur32ar, menu16ar, sys16ar, sm9ar)
        For Each Item In ar
            If Item Then DestroyIcon Item
        Next
    Next



    For Each Item In bimlar
        If Item Then ImageList_AddIcons 0 + Item
    Next
    
    'End If
    'MsgBox "GlobalGDI CLEAR"
    If m_GDIP_Token Then GdiplusShutdown m_GDIP_Token: m_GDIP_Token = 0

Else
    On Error Resume Next
    
    If IsNum(str_gdi) Then str_gdi = GlobalGDI.ParamName((gFindIndex(GlobalGDI.Source, L_(str_gdi)) - 1) \ 2)
    
    n = GlobalGDI.GetIndex(str_gdi)
    If n > -1 Then
    n = GlobalGDICount.Param(str_gdi) 'Количество пользователей этого объекта
    If n = 1 Then
        i = GlobalGDI.Param(str_gdi)
        If Left(str_gdi, 5) = "image" And i Then
            GdipDisposeImage i  ' m_Image
        Else
            DeleteObject i
        End If
        GlobalGDI.Param(str_gdi) = Null
        GlobalGDICount.Param(str_gdi) = Null
        'Debug.Print "ClearGDI " & str_gdi
    Else
        GlobalGDICount.Param(str_gdi) = n - 1
        'Debug.Print (n) & " ClearGDI " & str_gdi
    
    End If
    End If
End If
End Sub

'===============================================================
'===============================================================
'=============SYS IMAGE LIST ====================
Function SysImageListIndex(ByVal sFile) As Long ', Optional isDir As Boolean) As Long
Dim sfi As SHFILEINFO, dwFA&
'Const SFGAO_SHARE = &H20000
'Const SHGFI_SYSICONINDEX = &H4000
'Const SHGFI_SMALLICON = &H1
'Const SHGFI_USEFILEATTRIBUTES = &H10
sFile = S_(sFile)
'If sys16 = 0 Or sFile = 0 Then 'Or Len(sFile) = 0
'    sys16 = SHGetFileInfo("*.*", 0, sfi, LenB(sfi), &H4011)   'SHGFI_SMALLICON Or SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES)
'    ImageList_SetBkColor sys16, &H0 'FFFFFF
'End If
'If isDir Then dwFA = &H10 Else dwFA = &H80
If Len(sFile) Then dwFA = &H80 Else dwFA = &H10
sys16 = SHGetFileInfo(sFile, dwFA, sfi, LenB(sfi), &H4011) '1=SHGFI_SMALLICON Or 4000=SHGFI_SYSICONINDEX Or 10=SHGFI_USEFILEATTRIBUTES
'If GetIcon(sfi.iIcon, sys16) = -1 Then icar(4)(sfi.iIcon) = 0
SysImageListIndex = sfi.iIcon
End Function
'=============SYS IMAGE LIST ====================


'===GDI+

'Function GDIP_ThumbImage(sourcefile$, maxWidth&, maxHeight&, destfile$, Optional pquality = 80) As String
'Dim tSI As GdiplusStartupInput
'Dim quality As Byte
'Dim lres As Long
'Dim lGDIP As Long
'quality = pquality
'tSI.GdiplusVersion = 1
'If GdiplusStartup(lGDIP, tSI) = 0 Then ' Initialize GDI+
'    Dim lBitmap As Long
'    'If GdipCreateBitmapFromFile(StrPtr(sourcefile), lBitmap) = 0 Then ' Open the image file
'    If GdipLoadImageFromFile(StrPtr(sourcefile), lBitmap) = 0 Then ' Open the image file
'        If maxWidth > 15 And maxHeight > 15 Then '? need resize ?
'            Dim cWidth&, cHeight&, ratio As Double
'            Dim nWidth&, nHeight&
'            GdipGetImageWidth lBitmap, cWidth
'            GdipGetImageHeight lBitmap, cHeight
'            If cWidth > maxWidth Or cHeight > maxHeight Then 'Need Resize.
'                ratio = cWidth / cHeight 'исходный image
'                nWidth = maxWidth
'                nHeight = maxHeight
'                If cWidth > cHeight Then nHeight = maxWidth / ratio Else nWidth = maxHeight * ratio
'
'                'RESIZE
''If 0 Then
''                Dim lThumb&
''                GdipGetImageThumbnail lBitmap, nWidth, nHeight, lThumb, 0, 0
''                GdipDisposeImage lBitmap: lBitmap = 0 'Remove lBitmap
''                lBitmap = lThumb 'Resized image
''Else
'                Dim hDC&, hBITMAP&, hGraphics&
'                hDC = CreateCompatibleDC(ByVal 0)
'                hBITMAP = CreateBitmap(nWidth, nHeight, GetDeviceCaps(hDC, 14), GetDeviceCaps(hDC, 12), ByVal 0&)
'                hBITMAP = SelectObject(hDC, hBITMAP)
'                ' Resize the picture
'                GdipCreateFromHDC hDC, hGraphics
'                GdipSetInterpolationMode hGraphics, 7  'InterpolationModeHighQualityBicubic
'                GdipDrawImageRectI hGraphics, lBitmap, 0, 0, nWidth, nHeight
'                GdipDeleteGraphics hGraphics
'                GdipDisposeImage lBitmap: lBitmap = 0
'                ' Get the bitmap back
'                hBITMAP = SelectObject(hDC, hBITMAP)
'                DeleteDC hDC
'                GdipCreateBitmapFromHBITMAP hBITMAP, 0, lBitmap
''End If
'
'            End If
'        End If
'        'SAVE
'        If lBitmap Then
'        Dim tJpgEncoder As UUID 'GUID
'        Dim tParams As EncoderParameters
'        CLSIDFromString "{557CF401-1A04-11D3-9A73-0000F81EF32E}", tJpgEncoder   ' Initialize the encoder GUID
'        tParams.count = 1 ' Initialize the encoder parameters
'        With tParams.Parameter ' Quality
'            CLSIDFromString "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}", .GUID ' Set the Quality GUID
'            .NumberOfValues = 1
'            .Type = 4
'            .Value = VarPtr(quality)
'        End With
'        lres = GdipSaveImageToFile(lBitmap, StrPtr(destfile), tJpgEncoder, tParams) ' Save the image
'        GdipDisposeImage lBitmap
'        If lres = 0 Then GDIP_ThumbImage = destfile
'        End If
'    End If
'    GdiplusShutdown lGDIP
'End If
'End Function


''----------------------------------------------------------
'' Procedure : HandleToPicture
'' Purpose   : Creates a StdPicture object to wrap a GDI
''             image handle
''----------------------------------------------------------
''
'Public Function HandleToPicture(ByVal hGDIHandle As Long, ByVal ObjectType As PictureTypeConstants, Optional ByVal hpal As Long = 0) As StdPicture
'Dim tPictDesc As PICTDESC
'Dim IID_IPicture As UUID 'IID
'Dim oPicture As IPicture
'
'   ' Initialize the PICTDESC structure
'   With tPictDesc
'      .cbSizeOfStruct = Len(tPictDesc)
'      .picType = ObjectType
'      .hgdiObj = hGDIHandle
'      .hPalOrXYExt = hpal
'   End With
'
'   ' Initialize the IPicture interface ID
'   With IID_IPicture
'      .Data1 = &H7BF80981
'      .Data2 = &HBF32
'      .Data3 = &H101A
'      .Data4(0) = &H8B
'      .Data4(1) = &HBB
'      .Data4(3) = &HAA
'      .Data4(5) = &H30
'      .Data4(6) = &HC
'      .Data4(7) = &HAB
'   End With
'
'   ' Create the object
'   OleCreatePictureIndirect tPictDesc, IID_IPicture, True, oPicture
'   ' Return the picture object
'   Set HandleToPicture = oPicture
'
'End Function

'===GDI+

#If DragDrop Then

Function CreateDragImage(xc As xControl, ds$) As Long '(pDataObj As IDataObject)
If g_DragIML Then ImageList_Destroy g_DragIML: g_DragIML = 0
If Len(ds) = 0 Then Exit Function

Dim sz As SIZE, dp As xDRAWPARAMS, x As xControl
If xc Is Nothing Then Set x = xMainWnd Else Set x = xc
x.GetDrawParams VarPtr(dp), ""

Dim rc As RECT, arc As RECT
DrawFormat x, 0, rc, ds, arc, dp
sz.cx = arc.Right: sz.cy = arc.Bottom
If sz.cx > 300 Then sz.cx = 300
If sz.cy > 100 Then sz.cy = 100
SetRect rc, 0, 0, sz.cx, sz.cy
Dim hDC&: CreateMemDC hDC, sz.cx, sz.cy
DrawFormat x, hDC, rc, ds, arc, dp

Dim hOldBmp&: hOldBmp = CreateCompatibleBitmap(hDC, 1, 1) 'Создаем картинку 1x1
Dim hBITMAP&: hBITMAP = SelectObject(hDC, hOldBmp)
g_DragIML = ImageList_Create(arc.Right, arc.Bottom, 1, 0, 0)      'ILC_MASK = &H1&
ImageList_AddMasked g_DragIML, hBITMAP, &HFFFFFF
hOldBmp = SelectObject(hDC, hBITMAP)
DeleteObject hOldBmp
CreateMemDC hDC, 0, 0
CreateDragImage = 1
End Function

#End If

'
'
'Sub polytest()
'Dim dc&
'dc = CreateDC("DISPLAY", "", "", ByVal 0&)
'BeginPath dc
'Dim z&, i&
'z = 1 '000000
'Rectangle dc, 30 * z, 20 * z, 51 * z, 41 * z
''RoundRect dc, 40 * z, 30 * z, 60 * z, 50 * z, 5 * z, 5 * z
''EndPath dc
'SelectClipPath dc, 5
''BeginPath dc
'
'Rectangle dc, 40 * z, 30 * z, 61 * z, 51 * z
''Private Const RGN_AND As Long = 1
''Private Const RGN_OR As Long = 2
''Private Const RGN_XOR As Long = 3
''Private Const RGN_DIFF As Long = 4
''Private Const RGN_COPY As Long = 5
'EndPath dc
'
'
'Dim p0 As POINTAPI, t0 As Byte, n As Long
'Dim p() As POINTAPI, t() As Byte
'n = GetPath(dc, p0, t0, 0)
'If n Then
'    ReDim p(n - 1), t(n - 1)
'    GetPath dc, p(0), t(0), n
'
''Private Const PT_MOVETO As Long = &H6
''Private Const PT_LINETO As Long = &H2
''Private Const PT_BEZIERTO As Long = &H4
''Private Const PT_CLOSEFIGURE As Long = &H1
'Dim s$
'For i = 0 To n - 1
'If Len(s) Then s = s & IIf(t(i) = 6, ",", " ")
's = s & Replace(p(i).x / z & " " & p(i).y / z, ",", ".")
'Next
'Debug.Print s
'End If
'DeleteDC dc
'End Sub
