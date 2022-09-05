VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Parent As xControl 'Родитель столбца
Private m_Control As xControl 'Контрол для редактирования
Private m_DynamicControls As String 'List of dynamic controls   key,name,...,,
Private m_DynamicControl_ColumnIndex&

Private m_Name As String 'Имя столбца


Private m_Width As Long 'Минимальная Ширина столбца
'Private m_MinWidth As Long 'Минимальная Ширина столбца

Private m_Caption As String 'Подпись столбца или DrawFormat
Private m_DefCaption As String 'Подпись столбца из источника строк

'Private m_index As Long 'Индекс столбца
Private m_Enabled$
Public Tag

'Public Locked As Boolean
Private m_Visible As Boolean 'Видимость столбца
Private m_TabStop As Boolean 'TABSTOP
Private m_SelType As Byte 'SELTYPE 0=NONE 1=INVERT 2=FOCUS 3=ALFASELCOLOR
Private m_SelColor As Long ' 0

Private m_ControlSource As String
Private m_FieldIndex& 'Номер столбца в источнике строк родителя
Private m_DrawSource As String 'DrawFormat
Private m_HDrawSource As String 'HEADER DrawFormat
Private m_Font$ 'Font
Private m_vbTreeIcon$ 'Eval Набор иконов
Private mDFP0 As xDRAWPARAMS 'INITAL DRAWPARAMS
Private mDFP As xDRAWPARAMS
Private gdi$(2) 'Font,Back,NCBack
Private m_Format$
Private m_DefaultValue
Private m_ColumnSortInit& '0=NOT ALOWED 1=ALOWED OFF 2=ASC 3=DESC
Private m_ColumnSort& '0=NOT ALOWED 1=ALOWED OFF 2=ASC 3=DESC
'Private m_GridLines&
'Private m_Filter As String
'Private m_datatype As Long 'Тип данных в столбце

Public FastDraw As Boolean

Public LastRow As Boolean
Public LastCell As Boolean
Public OffsetX As Long    'Ширина предыдущих столбцов

Public BandIndex As Long

Public CellWidth As Long 'Текущая ширина ячейки

Public rowNum As Long 'номер строки БЭНДА

Public rowIndex As Long 'номер строки для ячейки в столбце
Public cellIndex As Long 'индекс ячейки в столбце
Public cellPos As Long 'позиция ячейки в столбце
Public Height As Long 'lines count
Public cellX As Long 'cell ofset in column
Public cellDY As Long 'cell Y ofset in column 'LINE BREAK CELL

Public CellsCount As Long 'cells count in column

Public ControlHidden As Boolean 'Не активировать XCONTROL

Private m_CellBack$
Private m_CellColor$
Private m_CellFont$
Private bNoMouse As Boolean

'Property Get ControlHidden() As Boolean
'ControlHidden = m_ControlHidden
'End Property
Property Get CellBack$()
CellBack = m_CellBack
End Property
Property Get CellColor$()
CellColor = m_CellColor
End Property
Property Get CellFont$()
CellFont = m_CellFont
End Property
Private Sub Class_Initialize()
FillMemory mDFP0, LenB(mDFP), 255
mDFP0.Transparent = 0 '-1
m_DrawSource = "" ' "DH0¦DW0¦TL1¦TA&H2010¦{}" 'DT_EDITCONTROL or DT_WORDBREAK
End Sub

Private Sub Class_Terminate()
'Debug.Print "COLUMND TERMINATE " & m_Name
If Not m_Control Is Nothing And Len(m_DynamicControls) = 0 Then
    On Error Resume Next
    If m_Control.ColumnIndex <> -1 Then
        Parent.ControlByName(m_Control.Name).CloseForm
    End If
End If
Set Parent = Nothing
Set m_Control = Nothing
If Len(gdi(0)) Then ClearGDI gdi(0)
If Len(gdi(1)) Then ClearGDI gdi(1)
If Len(gdi(2)) Then ClearGDI gdi(2)
End Sub

Public Property Get DynControl() As Boolean
DynControl = Len(m_DynamicControls)
End Property
Public Property Get Control() As xControl
If Len(m_DynamicControls) Then
On Error Resume Next
    Set Control = Parent.Eval(m_DynamicControls)
    If Not Control Is Nothing Then
        Control.ColumnIndex = m_DynamicControl_ColumnIndex
        m_FieldIndex = Control.FieldIndex
    End If
End If
If Control Is Nothing Then Set Control = m_Control
End Property
Public Property Set Control(xc As xControl)
Set m_Control = xc
If Not xc Is Nothing Then m_DynamicControl_ColumnIndex = xc.ColumnIndex
End Property

'Public Property Get DynControl()
'DynControl = m_DynamicControls
'End Property
'Public Property Let DynControl(ByVal v)
'm_DynamicControls = Nz(v)
'End Property

Sub INIT(ar, tbl As xControl, aColumnIndex&)
Dim i&, v3&, v8$
m_Enabled = 1
Set Parent = tbl
m_SelType = 255
m_SelColor = -1
m_DynamicControl_ColumnIndex = aColumnIndex
For i = 0 To UBound(ar) - (UBound(ar) + 1) Mod 2 Step 2

If Not IsArray(ar(i + 1)) Then
    v8 = S_(ar(i + 1))
    v3 = L_(v8)
    Select Case Cast(ar(i), vbString)
        Case "Name": m_Name = v8
        Case "CellWidth": Width = v8: CellWidth = m_Width   'Abs(v3)
        Case "ColumnVisible": m_Visible = v3
        Case "Enabled": m_Enabled = v3
        Case "ColumnSort": m_ColumnSortInit = v3
        Case "Caption": m_Caption = v8
        Case "SelType": m_SelType = CByte(0 + (v3 And &HFF))
        Case "SelColor": m_SelColor = v3
        Case "ControlSource": m_ControlSource = v8
        Case "TabStop": m_TabStop = v3
        Case "Format": m_Format = v8
        Case "TreeIcon": m_vbTreeIcon = v8
        Case "Tag": Tag = v8
        Case "DefaultValue": m_DefaultValue = v8
        Case "FastDraw": FastDraw = v3
        Case "DrawSource": m_DrawSource = v8
        Case "HDrawSource": m_HDrawSource = v8
        Case "DynControl": m_DynamicControls = v8
        Case "ControlHidden": ControlHidden = v3
        
        Case "BackGround": mDFP0.BackBrush = GlobalBrush(v8, mDFP.BackColor, 0): gdi(1) = "brush_" & v8
            mDFP0.CurrentBrush = mDFP0.BackBrush
            
        Case "NCBackGround":  mDFP0.NCBackBrush = GlobalBrush(v8, 0, 0): gdi(2) = "brush_" & v8
        Case "ForeColor": mDFP0.ForeColor = AnyColor(v3)
        Case "BorderColor":  mDFP0.BorderColor = AnyColor(v3)
        Case "Transparent":  mDFP0.Transparent = v3
        Case "TextAlign":   mDFP0.TextAlign = v3
        
        Case "GridLines":  mDFP0.GridLines = CByte(0 + (v3 And 7)) '&HFF))
        Case "NoFocus":  bNoMouse = 1
        
        Case "Font":  m_Font = v8
        
        Case "RowBack": m_CellBack = v8
        Case "RowColor": m_CellColor = v8
        Case "RowFont": m_CellFont = v8
    
    End Select
End If
Next

m_ColumnSortInit = 1
ControlSource = m_ControlSource
FillDrawParams
End Sub

Friend Sub FillDrawParams() '(ByVal parentDP As Long)
Dim ds$, dp As xDRAWPARAMS
Dim lf As LOGFONT
'Dim b As Boolean
Dim xc As xControl
Set xc = Control
'b = xc Is Nothing 'Простой столбец без контрола
'If Not b Then b = xc.ColumnIndex = -1 'Есть контрол

'Берем параметры форматирования из parentDP
'If parentDP = 0 Then
    Parent.GetDrawParams VarPtr(dp), "" 'Копируем из родителя
'Else
'    CopyMemory dp, ByVal parentDP, LenB(dp)
'End If

'If b Then ' НЕ КОНТРОЛА
If xc Is Nothing Then ' НЕТ КОНТРОЛА
    CopyMemory mDFP, mDFP0, LenB(dp) 'mDFP = mDFP0 'Берем исходные параметры
    mDFP.Focus = 0
    
    If mDFP.BackBrush = -1 Then mDFP.BackBrush = dp.BackBrush 'фон как у родителя
    If mDFP.CurrentBrush = -1 Then mDFP.CurrentBrush = dp.CurrentBrush
    
    If mDFP.ParentBackBrush = -1 Then mDFP.ParentBackBrush = dp.ParentBackBrush
    If mDFP.NCBackBrush = -1 Then mDFP.NCBackBrush = dp.NCBackBrush
    If mDFP.BackColor = -1 Then mDFP.BackColor = dp.BackColor
    'mDFP.Transparent = mDFP.Transparent Or mDFP.BackColor = -1

    If mDFP0.Transparent Then mDFP.BackBrush = 0: mDFP.CurrentBrush = 0
    
    If mDFP.ForeColor = -1 Then mDFP.ForeColor = dp.CurrentForeColor
    mDFP.CurrentForeColor = mDFP.ForeColor
    
    'If mDFP.ForeColor = -1 Then mDFP.CurrentForeColor = dp.ForeColor
    If mDFP.BorderColor = -1 Then mDFP.BorderColor = dp.BorderColor
    If mDFP.GridLines = 255 Then mDFP.GridLines = dp.GridLines
    'If mDFP.hFont = 1 Then mDFP.hFont = dp.hFont
    If Len(m_Font) Then
        gdiGetObject mDFP.hFont, Len(lf), lf
        mDFP.hFont = GlobalFontSRC(lf, m_Font)
        m_Font = GetFontSRC(mDFP.hFont)
        gdi(0) = "font_" & m_Font
        mDFP.CurrentFont = mDFP.hFont
    Else
        mDFP.hFont = dp.hFont
        mDFP.CurrentFont = dp.CurrentFont
        
    End If
    'mDFP.CurrentFont = mDFP.hFont
    'mDFP.CurrentFont = mDFP.CurrentFont
    
    mDFP.FontHeight = GetTextWidthHeight(mDFP.hFont, "gM")     'Высота шрифта
    If mDFP.TextAlign = -1 Then mDFP.TextAlign = dp.TextAlign

Else
    xc.GetDrawParams VarPtr(mDFP), "" 'Копируем из Controla
    'mDFP.Transparent = 1 'mDFP0.Transparent And Not mDFP.Focus ' Or mDFP.BackBrush = 0
    'Debug.Print Control.Name & "  GetDrawParams BackBrush=" & mDFP.BackBrush & " ParentBackBrush=" & mDFP.ParentBackBrush
    'If mDFP0.ForeColor = -1 Then mDFP.ForeColor = dp.CurrentForeColor
    'mDFP.CurrentForeColor = mDFP.ForeColor
'    If mDFP0.Transparent Then mDFP.BackBrush = 0 'Столбцы рисуются
'    mDFP.BackBrush = 5
'    mDFP.CurrentBrush = mDFP.BackBrush
End If
mDFP.Enabled = Enabled
End Sub

Friend Sub GetDrawParams(ByVal pDFP As Long, ds$)
If Len(ds) = 0 Then ds = DrawSource
mDFP.Enabled = Enabled 'С учетом текущей строки в таблице
'If Len(m_DynamicControls) Then 'Берем DrawParams из текущего контрола
'''mDFP << control.DrawParams
'FillDrawParams '0
'End If
If pDFP Then CopyMemory ByVal pDFP, mDFP, LenB(mDFP)
End Sub

Property Get CanDraw() As Boolean
CanDraw = m_Visible And (m_Width > 0)
End Property
Property Get CanTab() As Boolean
'CanTab = m_Visible And (m_Width > 0) And Enabled And Not bNoMouse And m_TabStop
CanTab = CanFocus And m_TabStop
End Property
Property Get CanFocus() As Boolean
CanFocus = m_Visible And (m_Width > 0) And Enabled And Not bNoMouse
End Property
Property Get Enabled() As Boolean
Enabled = 1
If IsNumeric(m_Enabled) Then
    Enabled = m_Enabled
ElseIf Len(m_Enabled) Then 'Вычисляемый Enabled
    'If Control Is Nothing Then
        Enabled = Parent.Eval(m_Enabled)
    'Else
    '    Enabled = Control.Eval(m_Enabled)
    'End If
End If
End Property

Public Property Get FieldIndex&()
FieldIndex = m_FieldIndex
End Property

Public Property Get ControlSource$()
ControlSource = m_ControlSource
End Property
Public Property Let ControlSource(v$)
m_ControlSource = v
'Dim i&
'i = InStr(m_ControlSource, ".")
'If i Then m_ControlSource = Mid(m_ControlSource, i + 1)
If Control Is Nothing Then
    m_FieldIndex = Parent.GetFieldIndex(m_ControlSource)
Else
    If Control.ColumnIndex = -1 Then
        m_FieldIndex = Parent.GetFieldIndex(m_ControlSource)
    Else
'        Control.ControlSource = v
        m_FieldIndex = Control.FieldIndex
    End If
End If
If m_FieldIndex > -1 Then m_DefCaption = Parent.GetFieldName(m_FieldIndex)
m_ColumnSort = Parent.ColumnSort(m_DefCaption) And m_ColumnSortInit
End Property

Public Property Get Name$()
Name = m_Name
End Property

Public Property Get ICaption$()
ICaption = m_Caption
End Property

Public Property Get Caption$()
If Len(m_Caption) Then Caption = m_Caption Else If Len(m_DefCaption) Then Caption = m_DefCaption Else Caption = "[" & m_Name & "]"
End Property
Public Property Let Caption(v$)
m_Caption = v
End Property

Public Property Get DefaultValue0()
DefaultValue0 = NzS(m_DefaultValue, Null)
End Property
'Public Property Get DefaultValue()
'DefaultValue = Parent.Eval(m_DefaultValue)
'End Property
'Public Property Let DefaultValue(v)
'm_DefaultValue = v
''Parent.UpdateDefaults
'If Not Parent.Matrix Is Nothing Then Parent.Matrix.DefaultValues(m_ControlSource) = m_DefaultValue
'
'End Property

Public Property Get DrawSource$()
Dim xc As xControl
Set xc = Control
If xc Is Nothing Or Len(m_DrawSource) Then
    DrawSource = m_DrawSource
Else
    DrawSource = xc.CellDrawSource
End If
End Property

Public Property Let DrawSource(v$)
If Control Is Nothing Then m_DrawSource = v
End Property
Public Property Get HDrawSource$()
Dim ns&, res$, ta&
ta = mDFP.TextAlign 'Or DT_EDITCONTROL Or DT_WORDBREAK
If m_ColumnSort > 0 And m_FieldIndex > -1 Then
    ns = ((m_ColumnSort And 6&) \ 2 + 1) \ 2 ' m_ColumnSort& '0=NOT ALOWED, 1=ALOWED OFF, 2=ASC, 3=DESC
    res = "DH0¦DW0¦TA" & ta & "¦NM~SortChange " & m_Name & "," & IIf(ns = 2, 1, ns + 2) & "¦" '& "RR¦PC2¦RO12¦"
    If Len(m_HDrawSource) Then res = res & m_HDrawSource Else res = res & "{}"
    If ns > 0 Then res = res & "¦SX-11¦IC" & (514 + ns) & " ¦SY2¦"
    HDrawSource = res
Else
    HDrawSource = m_HDrawSource
    If Len(m_HDrawSource) = 0 And mDFP.TextAlign Then HDrawSource = "DW0¦DH0¦TA" & ta & "¦{}"
End If
End Property
Public Property Let HDrawSource(v$)
m_HDrawSource = v
End Property

Public Property Get Sort()
If m_ColumnSort = 0 Then Sort = 0 Else Sort = ((m_ColumnSort And 6&) \ 2 + 1) \ 2 + 1
End Property
Public Property Let Sort(ns)
'0-OFF,1-ALLOW, 2-ASC, 3-DESC
m_ColumnSort = 2 ^ (L_(ns) And 3&) - 1
End Property

Public Property Get Visible() As Boolean
Visible = m_Visible 'And m_ColumnWidth > 0
End Property
Public Property Let Visible(v As Boolean)
m_Visible = v
If Not Control Is Nothing Then Control.Visible = v
Parent.Refresh
End Property

'Public Property Get TextAlign()
'TextAlign = mDFP0.TextAlign And 3
'End Property
'Public Property Let TextAlign(v)
'mDFP0.TextAlign = aVal(v) And 3
'End Property
'
'Public Property Get TabStop() As Boolean
'TabStop = m_TabStop
'End Property
'Public Property Let TabStop(v As Boolean)
'm_TabStop = v
'End Property


Public Property Get Width() 'Ширина родительского столбца
Width = m_Width
End Property
Public Property Let Width(v)
m_Width = Abs(Cast(v, vbLong))
CellWidth = m_Width

If Left$(v, 1) = "+" Then
    If mDFP0.TextAlign = -1 Then mDFP0.TextAlign = 0
    mDFP0.TextAlign = mDFP0.TextAlign Or 1& 'DT_CENTER
End If
If Left$(v, 1) = "-" Then
    If mDFP0.TextAlign = -1 Then mDFP0.TextAlign = 0
    mDFP0.TextAlign = mDFP0.TextAlign Or 2& 'DT_RIGHT
End If
m_Visible = m_Width > 0
End Property

'Public Property Get MinWidth()
'MinWidth = m_MinWidth
'End Property

'Public Property Let MinWidth(v)
'm_MinWidth = v
'End Property

'Public Property Get OffsetX&()
'OffsetX = m_OffsetX
'End Property
'Public Property Let OffsetX(v&)
'm_OffsetX = v
'End Property

Public Property Get TreeIcon&()
On Error Resume Next
If Len(m_vbTreeIcon) Then TreeIcon = Parent.Eval(m_vbTreeIcon)
End Property

Public Function CellValue(Optional row& = -3)
Dim res, vt&
On Error Resume Next
Dim b As Boolean
Dim xc As xControl
Set xc = Control
b = xc Is Nothing
If Not b Then b = xc.ColumnIndex = -1
If b Then 'columnindex=-1
    If m_FieldIndex > -1 Then
        res = Parent.List(m_FieldIndex, row)
        'res = Parent.List(IIf(m_FieldIndex = 1000, m_ControlSource, m_FieldIndex), row)
    ElseIf Len(m_ControlSource) Then
        If Left$(m_ControlSource, 1) = "=" Then
            res = Parent.Eval(Mid$(m_ControlSource, 2))
        Else
            res = Null
        End If
    End If
    
    If Len(m_Format) Then res = Format$(res, m_Format)
Else 'columnindex>-1
    If row = -3 Then
        res = xc.TextValue 'Для прорисовки
        Debug.Assert False
    Else
        res = xc.TextOfValue(xc.DataOwner.List(m_FieldIndex, row))  'Для возврата текста из указаной строки
        'res = xc.TextOfValue(Control.DataOwner.List(IIf(m_FieldIndex = 1000, m_ControlSource, m_FieldIndex), row))  'Для возврата текста из указаной строки
    End If
End If

vt = VarType(res)
Select Case vt
Case 0, 1, 10
    CellValue = ""
Case Else
    If vt And vbArray Then
        CellValue = "{" & VarInfo(res) & "}"
    Else
        If Left(res, 1) = Chr(0) Then If Len(res) = Abs(StringToLong(Left(res, 4))) Then res = "[Compressed]" ' " & StringToLong(Left(res, 4)) & "/" & StringToLong(Mid(res, 5, 4)) & "]"
        CellValue = res
    End If
End Select
Err.Clear
End Function

Public Property Get SelType() As Byte
SelType = m_SelType
End Property
'Public Property Let SelType(v As Byte)
'm_SelType = v
'End Property
Public Property Get SelColor() As Long
SelColor = m_SelColor
End Property
'Public Property Let SelColor(v As Long)
'm_SelColor = v
'End Property

