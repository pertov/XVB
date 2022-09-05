VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

'Private WithEvents rsa As ADODB.Recordset  'Рекордсет
'Private asyncDataReader As CDataObject

Private Type TRNode
    rowIndex As Integer  '0 = RowIndex ar()
    Expanded As Boolean '1 = Expanded
    Parent As Integer '2 = tr_Parent
    ChildCount As Integer '3 = ChildCount
    Level As Integer '4 = Level
    Next As Integer  '5 = Next node
    FirstChild As Integer '6 = FirstChild
    Lines As String '7 = Lines 1..Level 0=[   ], 1=[  |  ], 2=[ |- ] , 3=[ |_ ]
End Type

Private rs As ADODB.Recordset   'Рекордсет
Private ar  As Variant 'Матрица
Private oldvalues 'row oldvalues

Private nRows& 'Количество строк МатрицЫ
Private nCols& 'Количество строк МатрицЫ
'Private cPos& 'Текушая позиция

Private mDefaultValues As CParam 'Значения по умолчанию из СХЕМЫ + mDefValues
Private mAddDefValues As New CParam 'Добавляемые значения по умолчанию

Private curPos& 'Текушая позиция последнего изменения
Private bEvents As Boolean
Private bNotFullRefreshOwners As Boolean

Private m_Tree As Boolean 'Table/Tree f
Private m_TreeExpandedMaxLevel  As Integer 'Expanded node level
'Private m_TreeExpandAll As Byte 'Boolean
Private TreeRootVal ', TreeIDColumn&=0, TreeParentColumn&=1
Private m_RootNode As Boolean

Public TreePathSeparator As String

'Private tr  'Масив дерева
Private trNodes() As TRNode  'Масив дерева

Private tre  'Список строк для обработки
Private trm  'Список видимых строк
Private tr_rows& 'Количество видимых строк
Private tr_id As Long, tr_pid As Long
Private ux, xIndex&  'Номер столбца по которому залит сортированный индекс

Private m_Group As Boolean, gar

Private bOpenRowsCall As Boolean
Private m_invalid As Boolean
Private m_RSType& 'Тип источника строк 0-NoDefined 1-Text with semicolons 2-File ADTG 3-Function 4-CParam 5-SQL
Private m_RowSource 'источник строк
Private m_RowMode& 'Разрешенные действия с записями
'0 - SELECT
'1 - INSERT
'2 - UPDATE
'4 - DELETE
'8 - TABLE
'16 - FREETABLE
'32 - ADDNEW SERVER CURSOR
'Private m_rsFields As String 'Список полей в матрице
Private m_SARFields 'Соответствие столбца матрицы столбцу рекордсета
Private fatr 'FIELDS ATTRIBUTES

Private m_Sort As String
Private m_Filter 'As String
'Private m_Group As String 'GroupBy Column Name

Private m_NewRecordIndex As Long
Private m_New As Long 'rs.EditMode=2
Private m_RemoveNewRecord As Boolean
Private m_Dirty As Boolean

Private m_Connection As ADODB.Connection
Private m_Con As Boolean

Public xOwner As xControl 'Нужен для вычислений DefValues и сообщений из LOGLOCK
'Private hWndOwners 'array of hWnds Owners

Private hWndOwner& 'Owner TopWindowhWnd

'Public LockEdit As Boolean
Private arLockedFields
Public NoINITAR As Boolean


#If LogLock Then '==================LOGLOCK=================
Private b_LogRecord As Boolean 'Писать лог xMain.LogRecord при изменении или удалении данных
Private b_LockRecord As Boolean 'Блокировать запись xMain.LockRecord при изменении данных
Private Type MatrixLOGParams
    TableName As String        'arLG(0)
    PK_FieldIndex As Long     'arLG(1)
    hdr As Variant                 'arLG(2)
    HDRFINDEX As Variant    'arLG(3)
    oldvalues As Variant         'arLG(4)
    NewValues As Variant       'arLG(5)
    LockID As Long                'arLG(6)
End Type

Private mxLOG As MatrixLOGParams
Private xLogCodeObject As Object 'CodeObject with log functions
'Private arLG(6) '0-TableName 1-PK_FieldIndex 2-HDR 3-HDRFINDEX 4-OLD values 5-NEW values 6-LockID
Private locks() As Long 'Массив блокировок записей
Private mLogFields 'Массив имен/индексов полей для ЛОГИРОВАНИЯ   =*/Array() -все поля,  =Array(10,2,5)/="10,2,5"
Private mNoLogFields 'Поля которые не надо протоколировать
#End If '==================LOGLOCK=================

Private WithEvents RCP As CParam
Attribute RCP.VB_VarHelpID = -1

Private m_FullCParam As Boolean
Private b_cpUpdate As Boolean

Private m_CParamPath$ 'Содержание указанного пути

'Private m_SkipNodes 'Список узлов для фильтрации  m_SkipNodes as Array
'Private f_SkipNodes As Integer '-1=Пропускаем  1=не пропуск имен m_SkipNodes as Array

Private m_Nodes 'Список узлов для фильтрации  m_SkipNodes as Array
Private b_Nodes As Boolean 'Есть фильтрация по именам

Private b_SkipArrays As Boolean 'Пропуск узлов со значением ARRAY
'Private m_SkipArrays As Long 'Пропуск узлов со значением ARRAY 0=no skip, 1=skip to json, else=skip
Private m_CPColumns As String 'Список RS столбцов имя/тип/размер через запятую

Private nIncrementField& 'Индекс поля с автоинкрементом

Event OnRefresh(ByVal bFull As Boolean) 'Open Rows
Event OnUpdate(ByVal bAfter As Boolean) 'End edit/add row
Event OnDelete(ByVal index As Long, Cancel As Boolean) 'Before delete
Event OnChange(ByVal nField As Long) 'Edit field
Event OnStartChange() 'Edit field

'Event DataChanged() 'End edit
'Event StartChange() 'Start edit
'Event Change() 'Edit
'Event Canceled() 'Cancel edit

Private Sub Class_Initialize()
bEvents = True
'm_SkipNodes = Array()
m_invalid = -1
m_Nodes = Array()
TreeRootVal = Null
m_TreeExpandedMaxLevel = 1
Set mDefaultValues = New CParam
arLockedFields = Array()
m_RootNode = 0
xIndex = -1
tr_pid = 1
'Set asyncDataReader = New CDataObject
'Set xLogCodeObject = xMain
On Error Resume Next
#If LogLock Then '==================LOGLOCK=================
Set xLogCodeObject = xMainWnd.GlobalModule.CodeObject
#End If '==================LOGLOCK=================
End Sub


Private Sub Class_Terminate()
RemoveNewRecord
#If LogLock Then '==================LOGLOCK=================
If b_LockRecord And m_RSType = 6 Then
    If Not rs Is Nothing Then
        On Error Resume Next
        If mxLOG.LockID > 0 Then xLogCodeObject.RemoveLock mxLOG.LockID
    End If
End If
Set xLogCodeObject = Nothing
#End If '==================LOGLOCK=================
Set rs = Nothing
Set xOwner = Nothing
Set m_Connection = Nothing
End Sub

'Public Property Get xOwners() As Collection
'Set xOwners = New Collection
'If IsArray(hWndOwners) Then
'    On Error Resume Next
'    Dim h, x
'    For Each h In hWndOwners
'        Set x = Nothing: If h Then Set x = hxControl(h)
'        If InStr(TypeName(x), xs.sxControl) Then xOwners.Add x Else gRemoveIndex hWndOwners, h
'    Next
'End If
'End Property
'Public Property Get xOwner()
'On Error Resume Next: Set xOwner = Nothing: With xOwners: Set xOwner = .Item(.count): End With
'End Property
'Public Property Set xOwner(x)
'gAddIndex hWndOwners, x.hWnd
'End Property

Private Function pARFLDSIndex(flds)
On Error Resume Next
If ArrayDims(flds) <> 1 Then
    If VarType(flds) = 8 Then
        If flds = "*" Then flds = xMain.ToArray(rs.Fields, "=#index#") Else flds = Split(flds, ",")
    ElseIf VarType(flds) > 1 Then
        flds = Array("" & flds)
    Else
        flds = Array()
    End If
End If
Dim i&: For i = 0 To UBound(flds): flds(i) = "" & FieldIndex("" & flds(i)): Next
pARFLDSIndex = flds
End Function
Public Property Get LockEdit(Optional ByVal flds = "*") As Boolean
If UBound(arLockedFields) > -1 Then LockEdit = gFindIndex(arLockedFields, pARFLDSIndex(flds), 4) > -1
End Property
Public Property Let LockEdit(Optional ByVal flds = "*", v As Boolean)
If v Then gAddIndex arLockedFields, pARFLDSIndex(flds) Else gRemoveIndex arLockedFields, pARFLDSIndex(flds)
End Property

Private Sub RemoveNewRecord()
If m_RowMode And &H21 And m_New > -1 And m_RemoveNewRecord Then  'AddNewServer =DELETE RECORD
rs.AbsolutePosition = m_New + 1
rs.Delete
m_New = -1
m_NewRecordIndex = -1
On Error Resume Next
m_NewRecordIndex = rs.RecordCount
End If
End Sub

Public Property Get RecordCount&()
On Error Resume Next
If rs Is Nothing Then Exit Property
If m_Tree Or m_Group Then
    RecordCount = tr_rows
    'Debug.Print "TREE.RecordCount=" & RecordCount & " rs.RecordCount=" & rs.RecordCount
Else
    If m_NewRecordIndex > -1 And (m_RowMode And &H10) = 0 Then
        RecordCount = m_NewRecordIndex + 1 'IIf(m_New, 1, 0)
    Else
        RecordCount = nRows       'rs.RecordCount
    End If
End If
End Property

Public Property Get NewRecordIndex&()
NewRecordIndex = m_NewRecordIndex
End Property

Public Property Get RSType&()
RSType = m_RSType
End Property

'Property Get Modifed(ByVal dx, ByVal dy) As Boolean 'Возвращает флаг модификации поля
'On Error Resume Next
'Dim ty&, v, vr  'Индекс строки матрицы
'If m_Tree Or m_Group Then ty = tr(0, trm(dy + 1)) Else ty = dy
'If ty < 0 Or ty >= nRows Then Exit Property
'If aVal(dx) <> "" & dx Then dx = FieldIndex("" & dx)
'
'If dx < 0 Or dx >= nCols Then Exit Property
'
'v = ar(dx, ty)
'If Not m_Dirty Then
'    If cPos <> ty Then cPos = ty: RS.Move ty, 1
'    vr = RS.Fields(dx).Value
'End If
'
'If VarType(v) <> VarType(vr) Then
'    Modifed = True
'ElseIf VarType(v) > 1 Then
'    Modifed = Not (v = vr)
'End If
'
'End Property
'

Public Function BatchEdit(ByVal row, ParamArray var()) As Long
If UBound(var) < 0 Then Exit Function
Dim dy&, i&, v, be As Boolean
dy = L_(Nz(row, curPos)) 'cPos))
If dy >= nRows Then dy = m_NewRecordIndex
If dy > -1 Then
    If m_Tree Or m_Group Then dy = trNodes(dy).rowIndex 'Debug.Print "***************в дереве индекс строки найти надо******************"
    If IsArray(var(0)) Then v = var(0) Else v = var
    If bEvents Then bEvents = 0: be = 1
    row = UBound(v) - (UBound(v) + 1) Mod 2
    For i = 0 To row Step 2
        If i + 1 = row Then If be Then bEvents = 1 'На последнем поле включаем события если они были
        mx(v(i), dy) = v(i + 1)
    Next
End If


'Dim i&, dy&, oe As Boolean 'позиция текущей строки
'If m_New > -1 Then dy = cPos Else dy = rs.AbsolutePosition - 1
'If UBound(var) < 0 Then Exit Function
'If m_Tree Or m_Group Then
'    'dy = tr(0, dy)
'    dy = trNodes(dy).rowIndex
'    Debug.Print "***************в дереве индекс строки найти надо******************"
'    'Debug.Print "*********************************************************"
'    'Debug.Print "*********************************************************"
'
'End If
'Dim v
'BatchEdit = dy
'If BatchEdit > -1 Then
'    If IsArray(var(0)) Then v = var(0) Else v = var
'    If bEvents Then bEvents = 0: oe = 1
'    For i = 0 To UBound(v) - (UBound(v) + 1) Mod 2 Step 2
'        mx(v(i), dy) = v(i + 1)
'    Next
'    If oe Then bEvents = 1
'End If
End Function

'Property Get Position()
'Position = cPos
'End Property
'Property Let Position(ByVal p)
'p = Cast(p, vbLong)
'If p >= nRows Or p < 0 Then Exit Property
'rs.Move p, 1
'cPos = p
'End Property


'!!!!!! DEFAULT MEMBER !!!!!!
Property Get mx(ByVal dx, ByVal dy)   'Возвращает элемент матрицы
Attribute mx.VB_UserMemId = 0
On Error Resume Next
Dim ty&, v 'Индекс строки матрицы

'If m_Tree Or m_Group Then ty = tr(0, trm(dy + 1)) Else ty = dy
'If "" & Val(dx) <> "" & dx Then dx = FieldIndex("" & dx)
If VarType(dx) = vbString Then dx = FieldIndex("" & dx): If dx = -1 Then dx = nCols
dx = Cast(dx, vbLong)
If VarType(dy) = vbString Then dy = Cast(dy, vbLong)
 
'If dy = -1 Then dy = cPos
'If dy < -1 Then Exit Property

If m_Tree Or m_Group Then ty = trNodes(trm(dy + 1)).rowIndex Else ty = dy
 
If dx = -1 Then 'Индекс текущей строки
    mx = ty: rs.Move ty, 1
ElseIf dx < -1 Then 'Специалные данные
    If dx < -nCols - 1 Then 'Код блокировки LockID
        mx = 0
#If LogLock Then '==================LOGLOCK=================
        If b_LockRecord Then mx = locks(ty)
#End If '==================LOGLOCK=================
    Else ' - 2 >= dx >= - (nCols+1) ---->  0 <= Abs(dx)-2 <= nCols-1
        If m_Tree Then mx = GetPathU(ar(tr_id, ty), Abs(dx) - 2)   'Путь в дереве
    End If
End If

If dx < 0 Or dx > nCols - 1 Then
    'If dx > 255 And m_Dirty Then mx = oldvalues(dx Mod 256) 'OLD VALUE BEFORE CHANGE RECORD
    Exit Property
End If

v = Null
If ty = -2 Then
    'v = rs.Fields(dx).Name 'заголовок
    If m_Dirty Then v = oldvalues(dx)   'OLD VALUE BEFORE CHANGE RECORD
ElseIf dy >= m_NewRecordIndex And m_NewRecordIndex > -1 Then 'Значения по умолчанию для новой записи
    If rs.Fields(dx).Properties("ISAUTOINCREMENT").Value Or (m_RSType = 0 And (rs.Fields(dx).Attributes And adFldRowID) > 0) Then
        v = "(Counter)"
    ElseIf mDefaultValues.GetIndex(rs.Fields(dx).Name) > -1 Then
        v = mDefaultValues(rs.Fields(dx).Name)
        v = xMain.EvalAr(xOwner, v)   'DefValue
        v = Cast(v, rs.Fields(dx).Type, rs.Fields(dx).Attributes And adFldMayBeNull)
    End If
Else
    v = ar(dx, ty)
End If

'If DBDataType(rs.Fields(dx).Type) = adBoolean Then
If rs.Fields(dx).Type = adBoolean Then
    If IsNull(v) Then mx = 2& Else mx = CLng(v) And 1&
Else
    mx = v
End If
End Property
'!!!!!! DEFAULT MEMBER !!!!!!

Property Let mx(ByVal dx, ByVal dy, ByVal v)   'Изменение элемента матрицы

If (m_RowMode And 3&) = 0 Or LockEdit(dx) Or m_invalid Then Exit Property 'AddNew or Update
If VarType(dy) = vbString Then dy = Cast(dy, vbLong)


If VarType(dx) = vbString Then dx = FieldIndex("" & dx)
dx = Cast(dx, vbLong)

If dx < 0 Or dx > nCols - 1 Then Exit Property

If m_RSType = 44 Then If dx < 2 Then Exit Property

Dim ty&
'Запись изменений в рекордсет
If m_Tree Or m_Group Then
    If UBound(trm) > dy Then ty = trNodes(trm(dy + 1)).rowIndex Else ty = nRows
Else
    ty = dy
End If
If ty > rs.RecordCount Or ty < 0 Then Exit Property



'Convert dx of ar to dx of recordset
Dim fld As Field

If UBound(m_SARFields) >= dx Then Set fld = rs.Fields(m_SARFields(dx)) Else Exit Property 'Set fld = rs.Fields(dx)

If fld.Type = adBoolean Then v = IIf(v = 2, Null, v)

On Error Resume Next

If m_RSType <> 5 Then
'=================================

'If Not (m_New > -1 And ty = rs.RecordCount) And rs.RecordCount > 0 Then
If m_New <> ty And m_New > -1 And rs.RecordCount > 0 Then

    'If cPos > -1 And cPos <> ty And m_Dirty Then
    If curPos > -1 And curPos <> ty And m_Dirty Then
    'If rs.AbsolutePosition > 0 And (rs.AbsolutePosition <> ty + 1) And m_Dirty Then
        UpDate 'rs.UpdateBatch: m_Dirty = 0
    End If
End If

Dim arNames, arValues, i&, n1&, n2&, nm$

If (fld.Attributes And (adFldUnknownUpdatable Or adFldUpdatable)) = 0 Then Exit Property
If fld.Properties("ISAUTOINCREMENT").Value Then Exit Property
If m_RSType < 5 Then If (fld.Attributes And adFldRowID) > 0 Then Exit Property
'If IsNull(v) And Not (Not m_New > -1 And ty = rs.RecordCount) Then If (fld.Attributes And adFldMayBeNull) = 0 And m_RSType = 6 Then Exit Property
If IsNull(v) And Not (Not m_New > -1 And ty = rs.RecordCount) Then If (fld.Attributes And adFldIsNullable) = 0 And m_RSType = 6 Then Exit Property

If ty = rs.RecordCount And Not m_New > -1 Then 'ADDNEW  'Вставка новой записи
    
    If m_Tree Or m_Group Then Exit Property 'Нельзя добавлять в дерево
    If m_RSType = 4 And dx = 1 Then Exit Property 'Нельзя писать безымянный параметр

    If m_RSType = 44 Then
        If dx <> 2 Then Exit Property
        If Len(RCP.ParamName(S_(v))) Then Exit Property
    End If

    If (m_RowMode And 1&) Then 'ALLOW ADDNEW
        If Not mDefaultValues Is Nothing Then
            arNames = mDefaultValues.GetArray
            arValues = mDefaultValues.GetArray(1) 'arValues
            arValues = xMain.EvalAr(xOwner, arValues)
        End If
        
        Err.Clear
        rs.AddNew
        
        If IsArray(arNames) Then
            For i = 0 To UBound(arNames)   'DEFAULTVALUES FOR NEW RECORD
                With rs.Fields(arNames(i))
                    .Value = Cast(arValues(i), .Type, .Attributes And adFldMayBeNull)
                End With
            Next
        End If
        
        
'        If Err Then
'            gDebugPrint fld.Name & "(" & dx & ") " & Err.Description
'            Err.Clear
'         End If
         
        n1 = UBound(ar, 1): n2 = 0
        If nRows > 0 Then
            n2 = UBound(ar, 2) + 1 'Новая запись
            ReDim Preserve ar(n1, n2)
        End If
    
    '25/05/2017   НЕ СЧИТЫВАЮТСЯ ЗНАЧЕНИЯ СЧЕТЧИКА !!!!
    'Сохраняет запись
        If (m_RowMode And &H21) = &H21 And m_RemoveNewRecord Then
            rs.Resync adAffectCurrent 'Сохраняет запись на сервер = если не было сохранено то удалить эту запись
        End If
    
        For i = 0 To n1 'Значения из таблицы
            If m_RSType < 4 And (rs.Fields(i).Attributes And adFldRowID) > 0 Then rs.Fields(i).Value = NewCounter        'COUNTER
            ar(i, n2) = rs.Fields(i).Value
        Next
    'НЕ СЧИТЫВАЮТСЯ ЗНАЧЕНИЯ СЧЕТЧИКА !!!!
        
        If m_Tree Or m_Group Then ReDim Preserve tre(n2) As Boolean ': tre(n2) = True
        
        m_NewRecordIndex = rs.RecordCount
        If rs.CursorLocation = adUseServer Then m_NewRecordIndex = m_NewRecordIndex + 1
        'cPos = m_NewRecordIndex - 1
        curPos = m_NewRecordIndex - 1 'cPos
        m_New = curPos 'cPos
    
    Else 'DENY ADDNEW
        Exit Property
    End If
        'fld.Value = v 'Записываем поле в новой записи
        fld.Value = Cast(v, fld.Type, (m_RowMode And &H10) Or (fld.Attributes And adFldMayBeNull))

Else ' EDIT RECORD

#If LogLock Then '======================LOGLOCK=================
    If Not m_Dirty And m_RSType = 6 Then 'LOGLOCK
        Dim infotext$, hdr
        If b_LockRecord Then 'LockRecord
            If b_LC_TABLE Then ' Попытка заблокировать запись
                mxLOG.LockID = locks(ty)
                If mxLOG.LockID < 0 Then
                    i = xLogCodeObject.RowLocked(mxLOG.TableName, ar(mxLOG.PK_FieldIndex, ty), hWndOwner)
                    If i = 0 Then
                        If hWndOwner Then xOwner.hEvent "MatrixUnLocked", ty
                        Exit Property 'Запись разблокирована требуется перезагрузка
                    End If
                ElseIf mxLOG.LockID = 0 Then
                    mxLOG.LockID = xLogCodeObject.RowLock(mxLOG.TableName, ar(mxLOG.PK_FieldIndex, ty), hWndOwner, 1)
                    locks(ty) = mxLOG.LockID
                    'If hWndOwner Then xOwner.hEvent "MatrixLocked", mxLOG.LockID
                End If
                
                If mxLOG.LockID < 0 Then
                    If mxLOG.LockID = -1 Then Call xLogCodeObject.RecInfo(mxLOG.TableName, ar(mxLOG.PK_FieldIndex, ty), infotext) Else infotext = xLogCodeObject.LockInfo(mxLOG.LockID, 0)
                    If hWndOwner Then xOwner.hEvent "MatrixLocked", mxLOG.LockID
                    'If hWndOwner Then xMain.ShowBalloon xOwner.ActiveControl.hwnd, infotext, "Запись " & IIf(mxLOG.LockID = -1, "", "временно ") & "заблокированна", 3: RefreshLocks ty
                    Exit Property 'Запись блокирована от изменений не мной
                End If
            End If
        End If
        If b_LogRecord Then  'LogRecord
            'Запоминаем первоначальные значения перед внесением изменений
            hdr = mxLOG.HDRFINDEX
            ReDim arv(UBound(hdr))
            For i = 0 To UBound(hdr)
            If hdr(i) > -1 Then arv(i) = ar(hdr(i), ty)
            Next
            mxLOG.oldvalues = arv
        End If
    End If
    cPos = ty
#End If '======================LOGLOCK=================

    If Not m_New > -1 Then If (rs.AbsolutePosition - 1) <> ty Then curPos = ty: rs.Move ty, 1 ': Debug.Print "EDIT RECORD " & ty
    If Not (m_RSType = 0 And (fld.Attributes And adFldRowID) > 0) Then fld.Value = Cast(v, fld.Type, (m_RowMode And &H10) Or (fld.Attributes And adFldIsNullable))

End If



'Else 'm_RSType = 5
''=================================


End If 'm_RSType != 5
'=================================

'If m_RSType = 5 Then ar(dx, ty) = v Else ar(dx, ty) = fld
'Debug.Print "UPDATE", fld.Name, Nz(fld.Value)
If Not m_Dirty Then oldvalues = pArrayLine(ar, ty) ': Debug.Print curPos, "oldvalues=", NzS(Join2(oldvalues, ","), , "(", ")")


If Not RCP Is Nothing And (m_RSType = 4 Or m_RSType = 40 Or m_RSType = 44) And b_cpUpdate Then     'CPARAM UPDATE
    m_RSType = m_RSType + 100  'DENY to CALL rcp_Changed
    If m_RSType = 144 Then 'CParamColumns
        If dx = 3 Then 'Update value
            RCP(ar(0, ty) & "\" & rs.Fields(dx).Name) = v
        ElseIf dx = 2 Then
            If m_New = ty Then 'NewRecord
                ar(0, ty) = v: rs.Fields(0) = ar(0, ty)
                ar(1, ty) = "": rs.Fields(1) = ar(1, ty)
                Dim arow
                ReDim arow(-1 To 2 * (n1 - 2) - 1)
                For i = 3 To n1
                arow(2 * (i - 3) + 0) = rs.Fields(i).Name
                arow(2 * (i - 3) + 1) = rs.Fields(i).Value
                Next
                RCP(v) = arow
            Else 'Rename
                RCP.ParamName(ar(0, ty)) = v
                ar(0, ty) = IIf(Len(ar(1, ty)), ar(1, ty) & "\", "") & v
                rs.Fields(0) = ar(0, ty)
            End If
        End If
    
    ElseIf m_Tree Then 'CPARAM TREE (ID=PID\NAME,PID,NAME,VALUE)
        If Not (VarType(v) < 2 Or IsArray(v)) Then
        v = S_(v)
        If dx = 2 Then 'Rename
            nm = IIf(Len(ar(1, ty)), ar(1, ty) & "\", "") & v
            If RCP.PathParamIndex(nm) = -1 Then  'Можно переименовать
                'n1 = RCP.PathParamCount(ar(0, ty))
                RCP.PathParamName(ar(0, ty)) = v
'                If n1 = 0 Then
'                    ar(0, ty) = nm: rs.Fields(0) = ar(0, ty)
'                Else
                    m_RSType = m_RSType - 100: rcp_Changed: Exit Property 'm_RSType = m_RSType + 100
'                End If
            Else 'Не удалось переименовать
                rs.Fields(2) = ar(2, ty)
            End If
        ElseIf dx = 3 Then 'Update
            If m_New = -1 Then RCP(ar(0, ty)) = v
        End If
    End If
    
    ElseIf Not m_Tree Then
        If m_RSType = 104 Then 'CParamList  (Name,Value)
            If dx = 0 Then 'Add/Rename ParamName
                nm = fld
                If Len(nm) Then
                    If RCP.GetIndex(nm) = -1 Then  'Можно такое имя
                        If ar(0, ty) <> nm Then RCP.ParamName(ar(0, ty)) = nm Else RCP(nm) = Nz(ar(1, ty), Empty)
'                    Else 'Нельзя такое имя
'                        If ar(0, ty) <> nm Then 'Имя не такое как было
'                            fld = ar(0, ty)
'                        Else 'GEnerate New Name
'                            i = 0
'                            Do
'                            i = i + 1
'                            Loop While RCP.GetIndex(nm & "." & i, , 0) > -1
'                            nm = nm & "." & i
'                            fld = nm
'                            RCP(nm) = Nz(ar(1, ty), Empty)
'                        End If
                    End If
                Else
                    fld = ar(0, ty)
                End If
            Else 'dx=1 Edit Value
                RCP(ar(0, ty)) = v
            End If
        
        ElseIf m_RSType = 140 Then 'CParamColumns m_CParamPath
            'If dy = 0 Then RCP(NzS(RCP.ParamName(dx), sdx)) = v
            If dy = 0 Then RCP(NzS(rs.Fields(dx).Name, Null)) = v
        
        End If
    End If
    m_RSType = m_RSType - 100 'ALLOW to CALL rcp_Changed
    
    ar(dx, ty) = fld
    
Else 'NO CPARAM UPDATE
    If m_RSType = 5 Then ar(dx, ty) = v Else ar(dx, ty) = fld
End If


If Not m_Dirty Then
    m_Dirty = 1
    If bEvents Then RaiseEvent OnStartChange '(dx)
End If
If bEvents Then RaiseEvent OnChange(dx)
End Property

'Private Sub RCPTreeUpdate(ByVal dx, ByVal ty, ByVal v)
'If VarType(v) < 2 Or IsArray(v) Or ty < 0 Then Exit Sub
'v = Nz(v)
'If dx = 2 Then 'Rename + AddNew
'    If m_New = ty Then 'NewRecord
'        RCP(v) = ""
'        ar(0, ty) = v: arf(3, ty) = ""
'        rs.Fields(0) = ar(0, ty): rs.Fields(3) = ar(3, ty)
'    Else 'Rename
'        RCP.PathParamName(ar(0, ty)) = v
'        If RCP.PathParamName(ar(0, ty)) = v Then 'Удалось переименовать
'            ar(0, ty) = IIf(Len(ar(1, ty)), ar(1, ty) & "\", "") & v
'            rs.Fields(0) = ar(0, ty)
'        Else 'Не удалось переименовать
'            ar(2, ty) = RCP.PathParamName(ar(0, ty))
'            rs.Fields(2) = ar(2, ty)
'        End If
'    End If
'ElseIf dx = 3 Then 'Update
'    If m_New = -1 Then RCP(ar(0, ty)) = v
'Else
'    Exit Sub
'End If
'
'End Sub

Public Property Get Dirty() As Boolean
Dirty = m_Dirty
End Property
Public Property Let Dirty(v As Boolean)
m_Dirty = v
End Property
Public Property Get NewRecord() As Boolean
NewRecord = m_New > -1
'Debug.Print "NewRecord=" & NewRecord
End Property

Public Function Cancel(Optional cf) As Boolean
On Error Resume Next
If rs Is Nothing Then Exit Function
If ((m_RowMode And 3&) = 0) Or Not m_Dirty Then Exit Function

'Debug.Print rs.RecordCount
'If m_Dirty Then
Dim i&
i = FieldIndex(S_(cf))
If i > -1 Then
'rs.Fields(i).Value = rs.Fields(i).OriginalValue
    'Debug.Print ar(i, curPos), rs.Fields(i).Value, oldvalues(i)
    ar(i, curPos) = oldvalues(i)
    'Debug.Print ar(i, curPos), rs.Fields(i).Value, oldvalues(i)
    Cancel = 1
    Exit Function
End If
rs.CancelUpdate
'rs.Cancel 'Batch adAffectCurrent

'rs.CancelBatch adAffectCurrent
'Debug.Print rs.RecordCount

RemoveNewRecord
'If m_New Then XConnection.RollbackTrans: RS.Delete: Debug.Print RS.RecordCount, RS.Fields(0)
#If LogLock Then '==================LOGLOCK=================
'Удаляем нашу блокировку если она там есть
If b_LockRecord And m_RSType = 6 Then
    If mxLOG.LockID > 0 Then xLogCodeObject.RemoveLock (mxLOG.LockID)
End If
#End If '==================LOGLOCK=================

'bNotFullRefreshOwners = True
INITAR
'bNotFullRefreshOwners = False

'Err.Clear
'If bEvents Then RaiseEvent evCancel
'Debug.Print "Matrix.Cancel"
End Function

Public Function UpDate(Optional bCancel As Boolean) As Boolean
If ((m_RowMode And 3&) = 0) Or Not m_Dirty Then Exit Function
If bEvents Then RaiseEvent OnUpdate(0)

On Error Resume Next
If (m_RowMode And 16) = 0 Then
    Err.Clear
    rs.UpDate 'Batch adAffectCurrent
    'rs.UpdateBatch adAffectCurrent
    'rs.Fields.UpDate
    If Err Then
        Debug.Print Err.Description
        rs.CancelUpdate
        'rs.Fields.CancelUpdate
        'gDebugPrint "ERROR Matrix.UpdateBatch " & Err.Description & " cpos=" & cPos
    Else
        'Debug.Print "Matrix.UpdateBatch"
    End If
    'If m_New Then XConnection.CommitTrans
    Dim i&, hdr, hdri&
    If m_New > -1 And (Err <> 0 Or bCancel) Then
        ReDim Preserve ar(UBound(ar, 1), rs.RecordCount - 1) 'Удаляем последнюю запись
    Else
    
        
        #If LogLock Then '==================LOGLOCK=================

        If b_LogRecord Then
            hdr = mxLOG.HDRFINDEX
            ReDim arv(UBound(hdr))
            If b_LogRecord And m_New > -1 Then mxLOG.oldvalues = arv
            For i = 0 To nCols - 1
                ar(i, cPos) = rs.Fields(i).Value 'Читаем данные из рекордсета 'Возвращаем значения в исходную
                If b_LogRecord And m_RSType = 6 Then
                    hdri = gFindIndex(hdr, i)
                    If hdri > -1 Then arv(hdri) = ar(i, cPos)
                End If
            Next
    
            If m_RSType = 6 Then
                Dim txLog$
                If m_New > -1 Then
                    mxLOG.NewValues = mxLOG.oldvalues
                    mxLOG.oldvalues = arv  'Запоминаем новые значения после добавления
                    txLog = txLogFun
                    xLogCodeObject.LogRecord mxLOG.TableName, arv(mxLOG.PK_FieldIndex), hWndOwner, txLog, 1 'ADDNEW
                Else
                    mxLOG.NewValues = arv  'LOGLOCK 'Запоминаем измененные значения после обновления
                    txLog = txLogFun
                    If Len(txLog) Then xLogCodeObject.LogRecord mxLOG.TableName, arv(mxLOG.PK_FieldIndex), hWndOwner, txLog, 2 'UPDATE
                End If
            End If
        End If
        #End If '==================LOGLOCK=================
    End If
End If


#If LogLock Then '==================LOGLOCK=================
'Удаляем нашу блокировку если она там есть
If b_LockRecord And mxLOG.LockID > 0 Then
    xLogCodeObject.RemoveLock (mxLOG.LockID)
    mxLOG.LockID = 0: locks(cPos) = 0
End If
#End If '==================LOGLOCK=================

''Err.Clear
'm_Dirty = 0: m_New = -1: cPos = -1: m_NewRecordIndex = -1
''ar = rs.GetRows
'If Not ((rs.Filter = "" Or rs.Filter = 0) And rs.Sort = "") And rs.CursorLocation = adUseClient Then INITAR
'
'If m_RowMode And 1& Then m_NewRecordIndex = rs.RecordCount: nRows = m_NewRecordIndex
bNotFullRefreshOwners = True
INITAR
bNotFullRefreshOwners = False

If bEvents Then RaiseEvent OnUpdate(True)
End Function

#If LogLock Then '==================LOGLOCK=================

Private Function txLogFun$()
Dim oFields
oFields = xMain.GetArraysDiff(mxLOG.hdr, mxLOG.oldvalues, mxLOG.NewValues, 1)
'Debug.Print Join2(oFields, vbTab)
If UBound(oFields) > -1 Then txLogFun = gCompress(ToString(oFields))
End Function
Public Property Get LogRecord() As Boolean
LogRecord = b_LogRecord
End Property
Public Property Let LogRecord(v As Boolean)
b_LogRecord = v And b_LG_TABLE
End Property
Public Property Get LockRecord() As Boolean
LockRecord = b_LockRecord
End Property
Public Property Let LockRecord(v As Boolean)
b_LockRecord = v And b_LC_TABLE
End Property

Public Property Get LogFields()
On Error Resume Next
Dim arf, raf, rf, i&, fi&
If VarType(mLogFields) < 2 Or S_(mLogFields) = "*" Then
    arf = Array()
    arf = xMain.ToArray(rs.Fields)
ElseIf ArrayDims(mLogFields) = 1 Then
    arf = mLogFields
Else
    arf = Split(mLogFields, ",")
End If

i = 0 'LOGFIELDS ColumnName/FieldIndex to BaseName
For Each rf In arf
fi = FieldIndex("" & rf)
If fi > -1 Then arf(i) = rs.Fields(fi).Properties("BASECOLUMNNAME").Value
i = i + 1
Next

i = 0 'REMOVED ColumnName/FieldIndex to BaseName
If ArrayDims(mNoLogFields) = 0 Then raf = Split(S_(mNoLogFields), ",") Else raf = mNoLogFields
For Each rf In raf
fi = FieldIndex("" & rf)
If fi > -1 Then raf(i) = rs.Fields(fi).Properties("BASECOLUMNNAME").Value
i = i + 1
Next

gRemoveIndex arf, raf
LogFields = arf
End Property
Public Property Let LogFields(v)
    mLogFields = v
End Property
Public Property Let NoLogFields(v)
    mNoLogFields = v
End Property

Public Property Let LogLock(v)
LockRecord = aVal(v) And 1
LogRecord = aVal(v) And 2
End Property

Public Sub RefreshLocks(Optional index = -1)
Dim i&
If b_LockRecord Then 'Заливаем список блокировок
    'mxLOG.LockID = 0
    On Error Resume Next
    If index = -1 Then ReDim locks(UBound(ar, 2)) 'Else Erase locks
    For i = IIf(index = -1, 0, index) To IIf(index = -1, UBound(locks), index)
        locks(i) = xLogCodeObject.RowLocked(mxLOG.TableName, ar(mxLOG.PK_FieldIndex, i), hWndOwner)
    Next
    If index > -1 Then mxLOG.LockID = locks(index)
End If
End Sub
#End If '==================LOGLOCK=================

Public Sub DeleteGroup()
If (m_RowMode And 4&) = 0 Then Exit Sub
If m_RSType = 6 Then Exit Sub
On Error Resume Next
If m_Dirty Then rs.UpdateBatch: m_Dirty = 0 'Сохраняем рекордсет перед удалением из него записей
RemoveNewRecord 'Удаляем из него новую запись если была
Do
rs.Delete
rs.MoveFirst
Loop While Not rs.EOF
rs.UpdateBatch
INITAR
End Sub

Public Function Delete(ByVal dy) As Boolean
If (m_RowMode And 4&) = 0 Then Exit Function
If rs.RecordCount = 0 Then Exit Function
On Error Resume Next
dy = L_(dy)
If dy >= nRows Or dy < 0 Then Exit Function

Dim bCancel As Boolean
RaiseEvent OnDelete(dy, bCancel)
If bCancel Then Exit Function

If m_Dirty Then rs.UpdateBatch: m_Dirty = 0 'Сохраняем рекордсет перед удалением из него записей
RemoveNewRecord 'Удаляем из него новую запись если была

'If dy < nRows And dy > -1 Then

'If cPos <> dy + 1 Then cPos = dy: rs.Move dy, 1
If curPos <> dy + 1 Then curPos = dy: rs.Move dy, 1

#If LogLock Then '==================LOGLOCK=================
If m_RSType = 6 Then
If b_LockRecord And mxLOG.LockID > 0 Then
        xLogCodeObject.RemoveLock (mxLOG.LockID)
End If
If b_LogRecord Then
    Dim i&, hdr
    hdr = mxLOG.HDRFINDEX
    ReDim arv(UBound(hdr))
    mxLOG.NewValues = arv
    If m_New > -1 Then ar(mxLOG.PK_FieldIndex, cPos) = rs.Fields(mxLOG.PK_FieldIndex)
    For i = 0 To UBound(hdr)
        If hdr(i) > -1 Then arv(i) = ar(hdr(i), cPos)
    Next
    mxLOG.oldvalues = arv  'LOGLOCK 'Запоминаем старые значения до удаления
    Dim txLog$
    txLog = txLogFun
    If Len(txLog) Then xLogCodeObject.LogRecord mxLOG.TableName, arv(mxLOG.PK_FieldIndex), hWndOwner, txLog, 4 'DELETE
End If
End If
#End If '==================LOGLOCK=================
rs.Delete
rs.UpdateBatch
'End If

bNotFullRefreshOwners = True
INITAR
bNotFullRefreshOwners = False
Delete = 1
If bEvents Then RaiseEvent OnUpdate(True)
End Function

Function NewCounter&(Optional ByVal inital)   'MAX+1
Static NCInital&
If VarType(inital) = 10 Then 'NextCounter
    NCInital = NCInital + 1
'ElseIf NCInital = 0 And inital > -1 Then 'InitCounter
ElseIf VarType(inital) > 1 Then 'InitCounter
    NCInital = inital
End If 'ELSE CurrentCounter
NewCounter = NCInital
End Function

Public Property Get RowMode&()
RowMode = m_RowMode
End Property
Public Property Let RowMode(ByVal v&)
Dim old_RowMode&
old_RowMode = m_RowMode
m_RowMode = v And &H3F&
m_NewRecordIndex = -1
If m_RowMode And 1& Then m_NewRecordIndex = rs.RecordCount
If old_RowMode <> m_RowMode Then RefreshOwners False
End Property

Public Property Get Sort$()
Sort = m_Sort
End Property

Public Property Let Sort(ByVal v$) ' !!!!! CursorLocation=adUseClient  /// lastName DESC, firstName ASC
m_Sort = v
If rs Is Nothing Then Exit Property
bNotFullRefreshOwners = True
INITAR
bNotFullRefreshOwners = False
End Property

Public Property Get Filter()
Filter = m_Filter
End Property

Public Property Let Filter(ByVal v)
'Criteria string — a string made up of one or more individual clauses concatenated with AND or OR operators.
'Array of bookmarks — an array of unique bookmark values that point to records in the Recordset object.
'a FilterGroupEnum value
If m_Filter = v Then Exit Property
m_Filter = v
If rs Is Nothing Then Exit Property
bNotFullRefreshOwners = True
INITAR
bNotFullRefreshOwners = False

End Property

Public Sub OPTCParam(Optional ByVal wCParamOption, Optional ByVal sNodes, Optional ByVal sPathCParam, Optional ByVal bSkipArrays, Optional ByVal sColumns)
On Error Resume Next
Dim n&
n = wCParamOption
m_FullCParam = n And 1 '1= Полное дерево, 0= Только узлы с детями
b_cpUpdate = n And 2 '2=Update CParam
b_Nodes = False

If Len(sNodes) Then  'Есть имена для пропуска или отбора
    b_Nodes = True
    m_Nodes = Split(sNodes, ",") 'Имена которые пропускаем [-] или берем [+]
' +ParamName = берем только эти параметры а остальные не берем
' -ParamName = не берем этот параметр а остальные берем
' +*ParamName = берем только эти узлы а остальные не берем
' -*ParamName = не берем этот узел а остальные берем
'+ = только

End If

m_CPColumns = S_(sColumns)
b_SkipArrays = bSkipArrays 'Массивы не читаем
m_CParamPath = sPathCParam 'Начало отсчета  PathCParam
End Sub

'Function ReadFields(index, far)  'Replace column Names with Values
'Dim i&, n&, tar
'If VarType(far) = vbString Then tar = Split(far, ",") Else tar = far
'If VarType(tar) And vbArray Then
'On Error Resume Next
'For i = 0 To UBound(tar)
'n = -1: n = FieldIndex("" & tar(i))
'tar(i) = mx(n, index)
'If DBDataType(rs.Fields(n).type) = adBoolean Then If Nz(tar(i)) = 2 Then tar(i) = Null
'Next
'ReadFields = tar
''Err.Clear
''Debug.Print "ReadFields=" & Join(ar, ",")
'End If
'End Function

Public Function TreeCParam(key, Optional ByVal spCols, Optional ByVal spNames)
'return CParam.source of matrix.Tree
If Not (m_Tree Or m_Group) Then Exit Function
Dim ty&, res
If m_RSType = 4 Or m_RSType = 40 Then
    ty = GoNode(key)
    res = mx(3, ty) 'CParam Value
    
Else 'Bild Value
    ExpandChilds key 'Expand child nodes
    
    If Not IsMissing(spCols) And VarType(spCols) > 1 Then
        If Not IsArray(spCols) Then spCols = Split(spCols, ",")
    Else
        spCols = xMain.ToArray(rs.Fields, "name")
    End If
    If IsMissing(spNames) Then
        spNames = spCols
    Else
        If Not IsArray(spNames) Then spNames = Split(spNames, ",")
        ReDim Preserve spNames(UBound(spCols))
    End If
    
    Dim fns As New CParam, nm$, fi&
    For ty = 0 To UBound(spCols)
        fi = FieldIndex("" & spCols(ty)):
        nm = S_(spNames(ty))
        If Len(nm) = 0 And fi > -1 Then nm = rs.Fields(fi).Name
        fns("" & fi) = nm
    Next
    
    ty = TRIndex(key, 0, 0)
    res = pCPTree(ty, fns.Source)
End If
If ArrayDims(res) = 1 Then If LBound(res) = -1 Then res(-1) = ar(2, trNodes(ty).rowIndex)
TreeCParam = res
End Function

Private Function pCPTree(ty&, arNames)
Dim cp As New CParam
Dim n&, cy&, i&, nm$
cy = trNodes(ty).FirstChild
While cy > 0
    nm = ar(2, trNodes(cy).rowIndex)
    cp(nm) = pCPTree(cy, arNames)
    cy = trNodes(cy).Next
Wend
'For i = 3 To rs.Fields.count - 1
'cp(rs.Fields(i).Name) = ar(i, trNodes(ty).rowIndex)
'Next
For i = 0 To UBound(arNames) Step 2
    cy = CLng(arNames(i))
    If cy > 2 Then cp(arNames(i + 1)) = ar(cy, trNodes(ty).rowIndex)
Next
pCPTree = cp.Source
End Function

'Public Function RowCParam(dy, Optional ByVal spCols, Optional ByVal spNames)
''Return Array(colname0,colvalue0,...,colnameN,colvalueN) of rowindex=dy
'On Error Resume Next
'Dim ar, i&, fi&, nm$, cp As New CParam
''rs.MoveFirst
'rs.AbsolutePosition = aVal(dy) + 1
'ar = rs.GetRows(1, , spCols)
'If ArrayDims(ar) = 2 Then
'If Not IsMissing(spCols) And VarType(spCols) > 1 Then
'    If Not IsArray(spCols) Then spCols = Split(spCols, ",")
'Else
'    spCols = xMain.ToArray(rs.Fields, "name")
'End If
'If IsMissing(spNames) Then
'    spNames = spCols
'Else
'    If Not IsArray(spNames) Then spNames = Split(spNames, ",")
'    ReDim Preserve spNames(UBound(spCols))
'End If
''spCols=Array of fields
''spNames=Array of fieldnames
'For i = 0 To UBound(ar, 1) 'columns
'    nm = Nz(spNames(i))
'    If Len(nm) = 0 Then fi = FieldIndex("" & spCols(i)): If fi > -1 Then nm = rs.Fields(fi).Name
'    If Len(nm) Then cp.Param(nm) = ar(i, 0)
'Next
'RowCParam = cp.Source
'End If
'End Function


Function Rows() As Collection
Set Rows = New Collection
Dim i&, uu As UUID
uu.Data1 = ObjPtr(rs)
For i = 0 To RecordCount - 1
'    If m_Tree Or m_Group Then
'        uu.Data2 = trNodes(trm(i + 1)).rowIndex
'        uu.Data3 = i
'    Else
        uu.Data2 = i
'    End If
    Rows.Add CreateEventSinkObj(uu, rs.Fields, Nothing)
Next
End Function

Private Function Aggregate(crs As Recordset, ByVal gs) As Recordset
Dim ars As New Recordset
Dim i&, nm$, a, f, h: h = Array()
On Error Resume Next
'Debug.Print "Aggregate", gs
ReDim h(crs.Fields.count - 1)
Dim g, gf: gf = Split(gs, ","): g = Array() 'gGroupBY = g0,g1,g2.(sum[delimeter]|max|min|first|last|count)
gs = gf
For i = 0 To UBound(gf)
    gf(i) = FieldIndex("" & Split(gf(i), ".")(0))
    If gf(i) = "-1" Then gs(i) = "" Else gs(i) = Split(gs(i) & ".", ".")(1)
Next
ReDim g(UBound(h)) '= (name,calc,index)
a = 0
Dim fa
For i = 0 To UBound(h) 'Список выводимых полей
    h(i) = i: f = gFindIndex(gf, "" & i)  'group or calc
    If f = -1 Then nm = " " Else nm = gs(f): a = a + 1
    fa = StrField(crs.Fields(i))  'field definion
    With crs.Fields(i)
        g(i) = Array(.Name, IIf(f > -1, i, -1), nm)
        'If nm = "count" Then ars.Fields.append .Name, adInteger Else ars.Fields.append .Name, .Type, .DefinedSize, .Attributes
        Select Case Left$(nm, 3)
        Case "cou": fa = .Name & "&" 'CAST LONG
        Case "sum": If .Type >= 199 Then fa = .Name & "^" 'CAST MEMO
        End Select
        StrFieldAppend fa, ars
    End With
Next
If a = 0 Then Set Aggregate = crs: Exit Function

ars.Open
Dim v, ff, vv, wh$, n&

'Debug.Print xMain.CMatrix(8, ars).GetRows(0)

Do While Not crs.EOF
f = pArrayLine(crs.GetRows(1), 0, , 0)
wh = ""
For Each a In g
If Len(a(2)) = 0 Then 'GROUP FIELD FILTER
v = f(a(1))
nm = "="
Select Case VarType(v)
Case vbNull: nm = " is null"
Case vbDate: nm = nm & xMain.Quot(v, "#")
Case vbString: nm = nm & xMain.Quot(v, "'") '"='" & Replace(v, "'", "''") & "'"
Case vbDouble, vbSingle, vbCurrency, vbDecimal: nm = nm & Replace(v, ",", ".")
Case Else: nm = nm & v
End Select
If Len(wh) Then wh = wh & " and "
wh = wh & a(0) & nm
End If
Next

ars.Filter = wh
If ars.AbsolutePosition = 1 Then v = pArrayLine(ars.GetRows(1), 0, , 0): ars.AbsolutePosition = 1: n = 1 Else v = f: n = 0

For Each a In g
    i = a(1)
    If i > -1 Then 'AGGREGATE
        ff = f(i): vv = v(i)
        Select Case Left$(a(2), 3)
        Case "sum": If n Then If ars.Fields(i).Type >= 199 Then vv = S_(vv) & Mid$(a(2), 4) & ff Else vv = L_(vv) + ff
        Case "max": If vv < ff Then vv = ff
        Case "min": If vv > ff Then vv = ff
        Case "fir": 'First
        Case "cou": If n Then vv = L_(vv) + 1 Else vv = 1
        'Case "join": If n Then vv = Nz(vv) & Mid$(a(2), 5) & ff
        'Case "": vv = ff
        'Case "last": vv = ff
        Case Else: vv = ff
        End Select
        v(i) = vv
    End If
Next
If ars.AbsolutePosition = 1 Then ars.UpDate h, v Else ars.AddNew h, v
Loop

ars.Filter = ""
Set Aggregate = ars
End Function

Public Function GetRows(ByVal ReturnType, Optional ByVal spCols, Optional ByVal spNames, Optional ByVal gSort = "", Optional ByVal gFilter = "", Optional ByVal gGroupBY = "") ', Optional ByVal gCompute = "First,Last,Sum,Max,Min")
If m_invalid Then Exit Function
Dim i&, tmp$, s$, n&, tmpar, pout&
Dim cp As CParam
Dim st As New Stream

On Error Resume Next

ReturnType = L_(ReturnType)

n = bEvents: bEvents = 0
UpDate
bEvents = n

Dim bNoClone As Boolean
Dim nCount As Long
nCount = -1
Dim crs As Recordset
If m_RSType = 8 Or bNoClone Then Set crs = rs Else Set crs = rs.Clone

If VarType(gSort) <> vbString Then 'Record number
    bNoClone = True
    nCount = L_(gSort): gSort = ""
    If nCount >= nRows Or nCount < 0 Then
        nCount = 0
    Else
        crs.Move nCount, 1
        nCount = 1
    End If
End If
If VarType(gFilter) <> vbString Then 'Record count RETURN
    nCount = L_(gFilter): gFilter = ""
End If

If Not bNoClone And Not crs Is Nothing Then


If Len(gFilter) Then crs.Filter = gFilter '=  String |  Double(bookmark) | Long [0,1,2,3,5]
If Len(gSort) Then crs.Sort = gSort


If Len(gGroupBY) Then
'
'#If oldGR Then
'    Dim ggar, book_ar, groupFields, groupField, gas$
'    groupFields = Split(gGroupBY, ",")
'    ggar = groupFields
'    gRemoveIndex ggar, xMain.ToArray(crs.Fields, "Name")
'    If UBound(ggar) <> UBound(groupFields) Then 'хоть одно поле для групировки = прсутствует
'        ggar = Array(): book_ar = Array(): crs.MoveFirst
'        While Not crs.EOF
'            gas = "" 'Fields Group Value List
'            For Each groupField In groupFields 'BILD CURRENT ROW GROUP
'                 tmp = "": tmp = S_(crs.Fields(groupField))
'                gas = gas & tmp & Chr$(0)
'            Next
'            n = UBound(ggar)
'            i = gAddIndex(ggar, gas) 'Add group row
'            If i > n Then 'New GROUP
'                gAddIndex book_ar, crs.Bookmark  'Add row to filter
'            End If
'            crs.MoveNext
'        Wend
'        If UBound(book_ar) > -1 Then crs.Filter = book_ar Else crs.MoveFirst: gGroupBY = groupFields(0)
'    End If 'group fields = NO exist
'#Else
    'CREATE GRP RS
    tmp = Split(gGroupBY, ";")(0)
    If Len(tmp) Then
        Set crs = Aggregate(crs, tmp)
    End If
'#End If

End If 'Len(gGroupBY)


End If ' Not bNoClone

If nCount = -1 Then rs.MoveFirst

Select Case ReturnType

Case -1 'CMatrix (0)
    pout = 1
    GoSub get_tmpar
    Dim newmx As New CMatrix
    newmx.OpenRows 3, tmpar
    Set GetRows = newmx

Case 20, 21, 22 'ADTG COMPRESSED STRING
    'VARIANT FIELDS IN SOURCE = ERROR
    st.Type = adTypeBinary
    crs.Save st, IIf(ReturnType < 22, adPersistADTG, ReturnType - 21) 'adPersistXML)
    st.Position = 0
    If ReturnType = 20 Then '20 =Compressed ADTG
       GetRows = gCompress("" & StrConv(st.Read, vbUnicode))
    ElseIf ReturnType = 21 Then '21 = ADTG
        GetRows = StrConv(st.Read, vbUnicode)
    Else '22 = XML
        GetRows = st.ReadText
    End If
    
Case 8 'Recordset For Clone
    Set GetRows = crs.Clone(adLockReadOnly)
    
Case 80 'Recordset original
    Set GetRows = rs
    
Case 7, 70 'LIST PACKET
    'pout=0
    GoSub get_tmpar
    If ArrayDims(tmpar) = 2 Then
        If Len(gGroupBY) Then If InStr(gGroupBY, ";") > 0 Then tmp = Mid(gGroupBY, InStr(gGroupBY, ";") + 1)
        tmp = IIf(ReturnType = 7, NzS(tmp, "¶"), "")
        For i = 0 To UBound(tmpar, 2)
            For n = 0 To UBound(tmpar, 1)
                If Not (i = 0 And n = 0) Then s = s & tmp
                s = s & S_(tmpar(n, i))
            Next
        Next
    End If
    GetRows = s

Case 49 ' RCP
    If Not RCP Is Nothing Then Set GetRows = RCP
Case 44 'CParam.Source
    GetRows = TreeCParam(TreeRootVal, spCols, spNames)

Case 4, 40, 400, 41, 401, 42, 402 'CParam LIST spCols, 41
    GoSub get_tmpar
    Set cp = New CParam
    Dim arow
    Dim i1&: If m_Tree And m_RSType = 4 Then i1 = 1 'Исключаем ROOT NODE для дерева

    Select Case ReturnType
    Case 4 'CParam rows col0=name col1=value
        For i = i1 To UBound(tmpar, 2)
            cp("" & tmpar(0, i)) = tmpar(1, i)
        Next
    
    Case 40, 400, 42, 402 'CParam Row(s rt?(42,402)) [{colname:colvalue,...},{},...{}] + WITH JSON DECODE rt?>=400
    '40 = {c0:v0,c1:v1,...cN:vN}
    '400 = {c0:{v0.json},c1:{v1.json},...cN:{vN.json}}
    '42 = [{c0:v0,c1:v1,...cN:vN},{}]
    '402 = [{c0:{v0.json},c1:{v1.json},...cN:{vN.json}},{}]
        For i = i1 To IIf(ReturnType Mod 10, UBound(tmpar, 2), i1)
            tmp = ""
            If ReturnType Mod 10 Then tmp = NzS(gGroupBY, "rows") & "." & i & "\" ' FOR:42,402
            For n = 0 To UBound(spCols)
                GoSub arow_JSON_TO_CP  'JSON to CP
                cp(tmp & spNames(n)) = arow
            Next
        Next
        
    Case 41, 401 'Rows CParam {ar(0,N):{spNames(1):ar(1,N),spNames(2):ar(2,N),.. },..}
        ReDim ggar(-1 To 2 * UBound(spCols) - 1)
        For i = i1 To UBound(tmpar, 2) 'по всем строкам
            For n = 1 To UBound(spCols) 'по выбранным столбцам
                ggar(2 * (n - 1) + 0) = spNames(n)
                GoSub arow_JSON_TO_CP  'JSON to CP
                ggar(2 * (n - 1) + 1) = arow
            Next
            cp.Param("" & tmpar(0, i)) = ggar
        Next
        
'    Case 43 'CParam Rows row(col(0))=JSON columns{ 3,4,5,....} // mod_PrintDoc
'    ElseIf ReturnType = 43 Then 'CParam Rows row(col(0))=JSON columns{ 3,4,5,....} // mod_PrintDoc
'        For i = 0 To UBound(tmpar, 2)
'            tmp = NzS(tmpar(0, i), "row." & i)
'            For n = 3 To UBound(spCols)
'                cp(tmp & "\" & spNames(n)) = NzS(tmpar(n, i), Empty)
'            Next
'            cp(tmp) = cp.json(tmp)
'        Next

    End Select
    'End If
    GetRows = cp.Source
    
Case 0 ', 100 'CSV
    '0 CSV With header
    '100 CSV WithOUT  header
    pout = 2 ' HEADER
    GoSub get_tmpar
    GetRows = modMain.Join2(modMain.pEscapeUnEscape(tmpar, 0, 8), vbCrLf, vbTab)


Case 1, 10, 100, 101, 102, 103
    '1= 2DIMArray (0..nCols,-1 to nRows-1) with header
    '10=2DIMArray (0..nCols,0 to nRows-1) WITHOUT HEADER
    '100=RowsArray(0 to nRows-1) ColsArray(0..nCols)  WITHOUT HEADER
    '101=RowsArray(-1 to nRows-1) ColsArray(0..nCols)  WITH HEADER 1
    '102=RowsArray(-1 to nRows-1) ColsArray(0..nCols)  WITH HEADER 2
    '103=RowsArray(-1 to nRows-1) ColsArray(0..nCols)  WITH HEADER 3
    pout = ReturnType Mod 10
    GoSub get_tmpar
    If ReturnType > 100 And ArrayDims(tmpar) = 2 And (ReturnType Mod 10) > 0 Then 'Convert 2Dim to 1DIM
        ReDim tmpar1(LBound(tmpar, 2) To UBound(tmpar, 2))
        For i = LBound(tmpar, 2) To UBound(tmpar, 2)
            tmpar1(i) = pArrayLine(tmpar, i)
        Next
        GetRows = tmpar1
    Else
        GetRows = tmpar
    End If

'Case 90, 91, 92, 93 'JSON ARRAY with header 0,1,2,3
'    pout = ReturnType - 90
'    GoSub get_tmpar
'    'tmpar = crs.GetRows(nCount, 0, spCols)
'    'pPrepareOut tmpar, spCols, spNames, ReturnType - 90
'    GetRows = ArrayToJson(tmpar)


    
Case 2 'Inital Source
    GetRows = m_RowSource


End Select

tmpar = Empty
'If Len(pFilter) Then rs.Filter = pFilter
'If Len(pSort) Then rs.Sort = pSort
Err.Clear
Exit Function

get_tmpar:
    If Not (crs.BOF And crs.EOF) Then tmpar = crs.GetRows(nCount, 0, spCols)
    pPrepareOut tmpar, spCols, spNames, pout
    Return
arow_JSON_TO_CP:
    arow = tmpar(n, i)
    If ReturnType >= 400 And VarType(arow) = vbString Then
        If InStr("{[", Left(arow, 1)) Then arow = pJson2cpA(StrPtr(arow)): If UBound(arow) = -1 Then arow = tmpar(n, i)
    End If
    If VarType(arow) = vbNull Then arow = Empty
    Return
End Function

Private Sub pPrepareOut(tmpar, spCols, spNames, Optional ByVal nHeader As Long)
'spNames(i): =Array()  =String(,,,)
'   [/Name[newType]=default][vb]                position by name [+new type]+new value
'   [[Name]=default][vb]                                    new value
'   [newName[newType][=default]][vb]       rename [+ retype] [+ new value]
'newType                                                              string= $size$ | decimal = :Precision:NumericScale: | any type =  'adoDataType'DefinedSize'Attributes'
'default                                                                   xExpression (mx)
'vb                                                                           <~vbcode~>

' Name[STYPE]=Expression

Dim i&, n&, defs, nm$, nmar
Dim bDefs As Boolean, bCalcMX As Boolean ', bCalcM As Boolean
Dim fld As Field, b As Boolean

On Error Resume Next
If m_invalid Then Exit Sub
If ArrayDims(tmpar) <> 2 And nHeader > 0 Then 'Надо вставить заголовок
    n = -1
    If IsArray(spCols) Then n = UBound(spCols)
    If n = -1 Then n = rs.Fields.count - 1
    If n < 0 Then n = 0
    ReDim tmpar(n, -1 To -1)
End If
If ArrayDims(tmpar) <> 2 Then Exit Sub
If LBound(tmpar, 2) = 0 And nHeader <> 0 Then pArrayLine tmpar, -1, , 2 'INSERT HEADER

If Not IsArray(spCols) Then ReDim spCols(UBound(tmpar, 1))
n = UBound(spCols)
If IsArray(spNames) Then ReDim Preserve spNames(n) Else spNames = Split(S_(spNames) & String$(1 + n, ","), ",")

ReDim defs(n) 'Значения для заливки
ReDim tps(n) 'Типы данных полей
ReDim tpsz(n) 'Размеры полей
Dim cc$
Dim pns, attr& 'Precision/NumericSacle
Dim fdef
Dim vbcode$

For i = 0 To n
    cc = S_(spCols(i))
    If Len(cc) = 0 Then spCols(i) = i Else spCols(i) = FieldIndex(cc) 'сначала индекс поля в рекордсете
    Set fld = rs.Fields(spCols(i)) 'получаем field
    If Len(S_(spNames(i))) = 0 Then spNames(i) = rs.Fields(spCols(i)).Name  'Имя поля не указано = делаем как у источника
    spCols(i) = fld.Name 'имя поля в рекордсете
    spNames(i) = ExtractCode(spNames(i), vbcode) ' <~vbcode~> забираем VBCODE
    If Left(spNames(i), 1) = "/" And InStr(spNames(i), "=") Then
        attr = UBound(spNames) + 1
        ReDim Preserve spNames(attr)
        spNames(attr) = Mid(spNames(i), 2)
        spNames(i) = spCols(i)
    End If
    If Left(spNames(i), 1) = "=" Or Len(spNames(i)) = 0 Then spNames(i) = spCols(i) & spNames(i)
Next

If attr Then 'REPOS SPNAMES
    For i = n + 1 To UBound(spNames)
        fdef = StrFieldAppend(spNames(i)) 'получаем имя перемещенного поля  = надо найти новую позицию этого поля в spNames
        attr = gFindIndex(spCols, fdef(0))  'Находим индекс этого поля в столбцах
        If attr > -1 Then spNames(attr) = spNames(i)
    Next
    n = UBound(spCols)
    ReDim Preserve spNames(n)
End If

For i = 0 To n
Set fld = rs.Fields(spCols(i)) 'получаем field

fdef = StrFieldAppend(spNames(i), , fld) 'extract default value
If InStr(spNames(i), "=") = 0 Then 'RENAME+copy field
    cc = fdef(0) 'new name name
    fdef = StrFieldAppend(StrField(fld, , , fatr(i)))
    'fdef(3) = fatr(i) 'SCHEMA FIELD ATTRIBUTES
    fdef(0) = cc 'Change Name
End If
    
spNames(i) = fdef(0) 'Имя поля
tps(i) = fdef(1) 'Тип данных
tpsz(i) = IF0(fdef(2), -1) ' Размер поля
attr = fdef(3) 'Атрибуты поля
pns = Array(fdef(4), fdef(5)) ' Параметры NUMERIC
defs(i) = fdef(6)

If Len(S_(fdef(6))) = 0 And fld.Type = tps(i) And fld.DefinedSize = tpsz(i) Then   'Типы сходятся и нету расчетного значения
    defs(i) = Empty
Else
    bDefs = True
    bCalcMX = bCalcMX Or InStr(defs(i), "mx(") Or InStr(defs(i), "mx.")
End If

If nHeader = 0 Then 'NoHeader
ElseIf nHeader = 1 Then 'ArrayHeader
    If tps(i) = 131 Then
        tmpar(i, -1) = Array(spNames(i), tps(i), tpsz(i), attr, pns(0), pns(1))
    Else
        tmpar(i, -1) = Array(spNames(i), tps(i), tpsz(i), attr)
    End If
ElseIf nHeader = 2 Then 'StringHeader
    tmpar(i, -1) = StrField(spNames(i), tps(i), tpsz(i), attr, pns(0), pns(1))
ElseIf nHeader = 3 Then 'Simple ONLY NAME Header
    tmpar(i, -1) = spNames(i)
End If


Next

If bDefs Then 'Залить значения по умолчанию
    Dim xVBS As ScriptControl, v
    Set xVBS = xMain.VBScript("*+", "Dim mx,mxindex,maxrow:maxrow=" & UBound(tmpar, 2) & vbcode)
    Dim cpLine As New CParam, uu As UUID, par
    If bCalcMX Then Set xVBS.CodeObject.mx = CreateEventSinkObj(uu, cpLine, Nothing)  'Dispatch(cpLine)
    For i = 0 To UBound(tmpar, 2) 'FILL DEFAULT VALUE BY ROWS
        If bCalcMX Then cpLine.FillNames spNames, pArrayLine(tmpar, i): xVBS.CodeObject.mxindex = i
        For n = 0 To UBound(defs)
            If VarType(defs(n)) > 1 Then 'Для этого поля есть значение
                    If Len(defs(n)) Then
                        v = Null
                        v = xVBS.Eval(Replace(defs(n), "#i#", i))
                        tmpar(n, i) = v
                    End If
                    tmpar(n, i) = LimitSize(Cast(tmpar(n, i), tps(n), True), tpsz(n)) 'Приводим к типу и размеру
                    If bCalcMX Then cpLine.SVal(2 * n + 1) = tmpar(n, i) 'Пишем значение в модуль
            End If
        Next
    Next
    xVBS.Reset
    Err.Clear
End If

End Sub

Private Function ExtractCode(ByVal s, vbcode) ' <~vbcode~> забираем VBCODE
Dim v$
Do
v = GTag(s, 1, "<~", "~>"): If Len(v) Then vbcode = vbcode & vbCrLf & v
s = Replace(s, "<~" & v & "~>", "")
Loop While Len(v)
ExtractCode = s
End Function


Public Sub Resync(index&)
Dim i&
If m_RSType = 5 Or m_RSType = 6 Then
    If index > -1 And index < nRows Then
        'If cPos <> index Then cPos = index: rs.Move index, 1
        If curPos <> index Then curPos = index: rs.Move index, 1
        rs.Resync adAffectCurrent
        For i = 0 To rs.Fields.count - 1
        'ar(i, cPos) = rs.Fields(i).Value
        ar(i, curPos) = rs.Fields(i).Value
        Next
    End If
End If
End Sub

Public Function Connect(constr) As Long
Dim cs$
'If Not m_Connection Is Nothing Then If m_Connection.State Then m_Connection.Close
On Error Resume Next
m_Con = True
If TypeName(constr) = "Connection" Then
    Set m_Connection = constr
Else
    cs = S_(constr)
    If Len(cs) Then Set m_Connection = New ADODB.Connection: m_Connection.ConnectionTimeout = 5: m_Connection.Open cs Else Set m_Connection = Nothing: m_Con = False
End If

Connect = m_Connection.State
End Function

Property Get XConnection() As ADODB.Connection
'If m_Connection Is Nothing Then
'    Set XConnection = xMain.XConnection
'Else
    Set XConnection = m_Connection
'End If
End Property

Public Property Get RowSource()
RowSource = m_RowSource
End Property
Public Property Let RowSource(v)
OpenRows m_RSType, v, m_RowMode, TreeRootVal
End Property

Public Sub Save(ByVal fn$)
Debug.Assert False
fn = Replace(fn, "..\", xAppPath & "\")
On Error Resume Next
Kill fn
rs.Save fn, adPersistADTG
'If 1 Then 'COMPRESS
    xMain.SaveToFile fn, gCompress(xMain.ReadFile(fn))
'End If
End Sub

Public Function OpenRows(ByVal nType&, varSource, Optional nRowMode&, Optional vTreeRootVal = Null, Optional sExtra$, Optional SortFilterGroupBy$) As Boolean 'Return OK
On Error Resume Next

bOpenRowsCall = True
Dim cur&
cur = GetCursor()
SetCursor hCursor_WAIT

', Optional nTreeID& = -1, Optional nTreePID& = -1
Dim fn$, src$, ar0
Dim fsz&
'Dim n&
Dim u1&, u2&, u2L&, i&, j&, dim2 As Boolean
'Dim cp As CParam
Dim nRMode&, v

'If rs Is Nothing Then Set rs = New ADODB.Recordset Else rs.Close:

'Удаляем нашу блокировку если она там есть
#If LogLock Then '==================LOGLOCK=================
If b_LockRecord Then
    If Not rs Is Nothing Then
        If mxLOG.LockID > 0 Then
            xLogCodeObject.RemoveLock (mxLOG.LockID)
            'mxLOG.LockID = 0
        End If
    End If
End If
#End If '==================LOGLOCK=================
RemoveNewRecord

m_RSType = nType
TreeRootVal = IIf(VarType(vTreeRootVal) < 2, Null, vTreeRootVal)
m_Tree = Not IsNull(TreeRootVal)


Set RCP = Nothing
Set rs = New ADODB.Recordset

hWndOwner = 0
If Not xOwner Is Nothing Then hWndOwner = xOwner.hWnd

m_RemoveNewRecord = 0
nIncrementField = -1

Dim st As New Stream, b() As Byte


If Len(SortFilterGroupBy) Then
    ar0 = Split(SortFilterGroupBy & ";;", ";")
    m_Sort = ar0(0)
    m_Filter = ar0(1)
    m_Group = Len(ar0(2))
    If m_Group Then gar = Split(ar0(2), ",")
End If

If nType < -1 Then Exit Function

m_invalid = 0

m_RowSource = Empty
nRMode = 7

rs.CursorLocation = adUseClient
'rs.CursorLocation = adUseServer

'Dim nNewCounter& 'RECALC COUNTER
'nNewCounter = -1

''Connection string from rowsource
'If VarType(varSource) = vbString And (m_RSType = 5 Or m_RSType = 6) Then
'    ar0 = Split(varSource, Chr(0))
'    If UBound(ar0) = 1 Then
'        Connect ar0(0)
'        varSource = ar0(1)
'    ElseIf Left(varSource, 1) = "@" Then 'XMDB
'        varSource = Mid(varSource, 2)
'        If xMain.XMDB Then Connect xmdb_conn
'    End If
'End If

NewCounter 0

Select Case m_RSType
Case -1 'Copy CMatrix/RecordSet
    'VARIANT FIELDS IN SOURCE = ERROR Err =&H80040E1D
    st.Type = adTypeBinary
    Select Case TypeName(varSource)
    Case "CMatrix"
        varSource.GetRows(80).Save st
    Case "Recordset"
        varSource.Save st
    End Select
    st.Position = 0
    rs.Open st


'Case 7 'String List 1 Column
'    m_RowSource = Cast(varSource, vbString)
'    ar0 = Split(Replace(Replace(m_RowSource, "¶", ";"), ";", ","), ",")
'    rs.Fields.append "col", adVarChar, 255
'    rs.Open
'    For i = 0 To UBound(ar0)
'    rs.AddNew "col", ar0(i)
'    Next
'    'rs.UpdateBatch
    
Case 7  'SAR(,|¶|;)/1DimArray    to IF0(nRowMode\&H10000,1) Columns
    m_RowSource = varSource
    ar0 = varSource
    If VarType(varSource) = vbString Then ar0 = Split(Replace(Replace(m_RowSource, "¶", ";"), ";", ","), ",") ': If UBound(ar0) = -1 Then ar0 = Array("")
    
    'If ArrayDims(ar0) = 1 And (nRowMode And &HFF) > 0 Then
    If ArrayDims(ar0) = 1 Then 'And (nRowMode And &H1F0000) > 0 Then
    u2 = UBound(ar0)
    
    'u1 = nRowMode And &HFF
    u1 = (nRowMode And &HFF0000) \ &H10000: If u1 = 0 Then u1 = 1
    
    u2 = u2 - ((u2 + 1) Mod u1)
    For i = 0 To u2 Step u1
        If i = 0 Then
            ReDim h(u1 - 1)
            If InStr(gRSType, Right(" " & ar0(0), 1)) Then 'HEADER
                For j = 0 To u1 - 1: h(j) = Split(StrFieldAppend(ar0(j), rs), "=")(0): Next
                i = j
            Else 'Standart header
                For j = 0 To u1 - 1: h(j) = "c" & j: rs.Fields.append h(j), adVarChar, 255: Next
            End If
            rs.Open
        End If
        If i <= u2 Then
            ReDim hv(u1 - 1)
            For j = 0 To u1 - 1: hv(j) = Cast(ar0(i + j), rs.Fields(j).Type, True): Next
            rs.AddNew h, hv
        End If
    Next
    'rs.UpdateBatch
    'nRowMode = 0
    End If

Case 0, 1 '0=TSV {tab+crlf}   1=File
    m_RowSource = Cast(varSource, vbString)
    AddCSVRows rs, NzS(xMain.ReadFile(m_RowSource), m_RowSource)
    'rs.UpdateBatch
    
Case 2, 20 '2=File ADTG, 20=ADTG COMPRESSED STRING
    m_RowSource = Cast(varSource, vbString)
    If Len(m_RowSource) Then
    If m_RSType = 2 Then fn = xMain.ReadFile(m_RowSource) Else fn = m_RowSource
    If Len(fn) Then
        b = StrConv(fn, vbFromUnicode)
        If gIsCompressed(fn) Then b = gDecompress(b)
        Set st = xMain.ADOStream(b)
        rs.Open st
        If Err Then Err.Clear: If m_RSType = 2 Then rs.Open fn, , , , 256
    End If
    End If
    m_RSType = 2

'*********************************************************************************************************
'Case 9 'JSON ARRAY
Case 3, 9 'Array 1/2-Dim    Заголовки столбцов в (-1) строке
    m_RowSource = varSource
    Dim sar
    If m_RSType = 9 Then sar = JsonToArray(varSource, True) Else sar = varSource
   
    If VarType(sar) And vbArray Then
        u1 = -1: u1 = UBound(sar, 1) 'Количество столбцов
        u2 = -2: u2 = UBound(sar, 2) 'Количество строк
        If u2 = -2 Then u2 = u1 Else dim2 = True
        If dim2 Then u2L = LBound(sar, 2) Else u2L = LBound(sar, 1)
        If dim2 Then
            ReDim h(u1), c(u1)
            For i = 0 To u1
                If u2L = -1 Then h(i) = sar(i, -1) Else h(i) = "arColumn" & i & "$"
                If VarType(h(i)) And vbArray Then
                    'If (h(i)(3) And adFldRowID) Or (h(i)(3) = 90) Then nNewCounter = i 'NewCounter
                    rs.Fields.append h(i)(0), h(i)(1), h(i)(2), h(i)(3)
                    If h(i)(1) = adNumeric Then
                        If UBound(h(i)) > 4 Then
                            rs.Fields(h(i)(0)).Precision = h(i)(4)
                            rs.Fields(h(i)(0)).NumericScale = h(i)(5)
                        End If
                    End If
                    h(i) = h(i)(0)
                ElseIf VarType(h(i)) = vbString Then
                    h(i) = StrFieldAppend(NzS(h(i), "arColumn" & i & "$"), rs)
                Else
                    rs.Fields.append h(i), adVariant
                End If
            Next
            
        Else 'DIM1
            If u2L = -1 Then fn = sar(-1) Else fn = "Value"
            rs.Fields.append fn, adVariant
        End If
        'rs.CursorLocation = adUseClient
        rs.Open
        
        For i = 0 To u2 'Строки добавим
            If dim2 Then
                'c = pArrayLine(sar, i)
                rs.AddNew h, pArrayLine(sar, i)
            Else
                rs.AddNew fn, sar(i)
            End If
        Next
        
        'rs.UpdateBatch
    End If
'*********************************************************************************************************
    
Case 4, 40, 44 'CParam Let/Set - одномерный именованный  массив с вложениями
    
    If nRowMode And &H1C00 Then OPTCParam (&HC00 And nRowMode) \ &H400, , , nRowMode And &H1000&
    
    If Len(sExtra) Then m_CParamPath = sExtra
    If TypeName(varSource) = "CParam" Then 'SET
        Set RCP = varSource
    Else 'ARRAY LET
        Dim cp As New CParam
        cp.Value = varSource
        Set RCP = cp
    End If
    
    If Len(SortFilterGroupBy) Then
        ar0 = Split(SortFilterGroupBy, ";")
        If UBound(ar0) > 2 Then m_CPColumns = ar0(UBound(ar0))
    End If
    
    rcp_Changed
    nRMode = 3
    If m_RSType = 4 Then 'ListParams
        If Not m_Tree Then nRMode = 7
    ElseIf m_RSType = 40 Then 'CParamColumns
    Else 'm_RSType = 40  'arCParam to Tree matrix (id_$, pid_$, name_$, .. ,fields)
        nRMode = 7
    End If
    'rs.UpdateBatch
    
'*********************************************************************************************************
Case 5, 6
    src = S_(varSource)
    m_RowSource = src
    
    'Dim xc As Connection: Set xc = xMain.pxcon(src)
    'If m_Connection Is Nothing Then Set m_Connection = xMain.pxcon(src)
    If Not m_Con Then Set m_Connection = xMain.pxcon(src)
    
    src = xMain.RemoveSQLComments(src, m_Connection)
    If m_RSType = 6 Then '6 =SQL TABLE SERVER KEYSET
        If m_Connection.Provider = "Microsoft.Jet.OLEDB.4.0" Then rs.CursorLocation = IIf(nRowMode And &H20&, adUseServer, adUseClient) Else m_RemoveNewRecord = (nRowMode And &H21&) = &H21
        rs.Open src, m_Connection, adOpenStatic, adLockOptimistic
        nRMode = 0: fsz = -1: fsz = rs.Fields.count - 1
        For i = 0 To fsz
            If (rs.Fields(i).Attributes And (adFldUnknownUpdatable Or adFldUpdatable)) <> 0 Then nRMode = nRMode Or 7: Exit For
        Next
        nRMode = nRowMode And nRMode
    Else '5=SQL LIST CLIENT SIDE
        rs.Open src, m_Connection, adOpenForwardOnly, adLockReadOnly
        Set rs.ActiveConnection = Nothing
    End If

''*********************************************************************************************************
Case 8 'RECORDSET CLONE
    m_RowSource = ""
    'xMain.DebugPrint 0, "RECORDSET CLONE = " & TypeName(varSource)
    If TypeName(varSource) = "Recordset" Then


'        st.Type = adTypeBinary
'        varSource.Save st
'        st.Position = 0
'        rs.Open st

        Err.Clear
        Set rs = varSource.Clone(adLockReadOnly)
        'If rs.State = 0 Then Set rs = varSource '.Clone(adLockReadOnly)
        If Err Then Set rs = varSource: Err.Clear '.Clone(adLockReadOnly)

    End If

    nRMode = 0
'*********************************************************************************************************
'Case 10 'RECORDSET CLONE FROM STORAGE
'    Set rs = Nothing
'    Set rs = xMain.Storage(m_RowSource).Clone
'    nRMode = 0
Case Else
    m_RowSource = Cast(varSource, vbString)
    nRMode = 0
End Select

'm_rsFields = "" 'ALL FIELDS
If Len(sExtra) And m_RSType <> 4 Then ' m_rsFields = sExtra

'If Len(m_rsFields) Then
    'ar0 = Split(m_rsFields, ",")
    ar0 = Split(sExtra, ",")
    ReDim m_SARFields(UBound(ar0))
    For i = 0 To UBound(ar0)
    m_SARFields(i) = IIf(IsNumeric(ar0(i)), CLng(ar0(i)), ar0(i))
    Next
Else
    ReDim m_SARFields(rs.Fields.count - 1)
    For i = 0 To rs.Fields.count - 1
    m_SARFields(i) = i
    Next
End If
ReDim fatr(-1 To UBound(m_SARFields))

m_RowMode = (nRMode And nRowMode) Or (&HFFFFF8 And nRowMode)
nRowMode = m_RowMode

InitSQLDefaultValues
Set DefaultValues = mAddDefValues

Dim n& 'Column Count
If Not rs Is Nothing Then n = IIf(rs.State, rs.Fields.count, 0) - 1 Else n = -1 'MaxFieldIndex
'If m_Tree And n < 2 Then TreeRootVal = Null: m_Tree = 0  'Проверка рекордсета на подходимость к TREE

If n < 0 Then 'Не получилось списка = Добавим один столбец
    Set rs = New ADODB.Recordset: rs.Fields.append "Value", adVariant: rs.CursorLocation = adUseClient: rs.Open
    rs.AddNew "Value", "NoRecords" 'varSource
    'rs.UpdateBatch
    m_RowMode = 0
    m_invalid = 1
End If

If rs.ActiveConnection Is Nothing Then rs.UpdateBatch

INITAR


SetCursor cur

If m_RSType < 4 And NewCounter(Null) = 0 Then
    For i = 0 To nCols - 1
        If (rs.Fields(i).Attributes And adFldRowID) Or (rs.Fields(i).Attributes = 90 And rs.Fields(i).Type = adInteger) Then
            'nNewCounter = i 'NewCounter
            nIncrementField = i
            NewCounter L_(DMinMax(nIncrementField))
            Exit For
        End If
    Next
End If


bOpenRowsCall = False
End Function


Private Sub InitSQLDefaultValues()
Dim sh As Recordset, i&

'If Not (m_RSType = 5 Or m_RSType = 6) Then Exit Sub
If m_RSType <> 6 Then Exit Sub

#If LogLock Then '==================LOGLOCK=================
If b_LogRecord Or b_LockRecord Then  'Инициализация ARLG LOGLOCK


'Debug.Print xMain.CParam(xMain.ToArray(rs.Fields(0).Properties, "name", "Value")).View

    mxLOG.TableName = rs.Fields(0).Properties("BASETABLENAME").Value 'TABLE NAME
    mxLOG.PK_FieldIndex = FieldIndex(xLogCodeObject.PrimaryKey(mxLOG.TableName), 1)   'PK_FIELD_INDEX
    'ReDim hdr(rs.Fields.count - 1)
    If b_LogRecord Then
        Dim arfs, af
        arfs = LogFields
        ReDim hdr(UBound(arfs))
        ReDim hdrIndex(UBound(arfs))
        For Each af In arfs
            hdr(i) = af 'rs.Fields(af).Properties("BASECOLUMNNAME").Value
            hdrIndex(i) = FieldIndex("" & hdr(i), 1)
            i = i + 1
        Next
        mxLOG.hdr = hdr 'Список полей
        mxLOG.HDRFINDEX = hdrIndex 'Список индексов полей
    End If
End If
#End If '==================LOGLOCK=================


Set mDefaultValues = New CParam

'Debug.Print rs.Fields(0).Properties("BASETABLENAME").Value
Set sh = m_Connection.OpenSchema(adSchemaColumns, Array(Empty, Empty, rs.Fields(0).Properties("BASETABLENAME").Value, Empty))
'xMain.ClipboardText = xMain.CMatrix(8, sh).GetRows(0)
'sh.Filter = "COLUMN_HASDEFAULT=True"
While Not (sh.BOF Or sh.EOF)
    If sh.Fields("COLUMN_HASDEFAULT") Then mDefaultValues(sh.Fields("COLUMN_NAME").Value) = sh.Fields("COLUMN_DEFAULT").Value
    i = sh.Fields("ORDINAL_POSITION") - 1
    fatr(i) = sh.Fields("COLUMN_FLAGS") Or IIf(sh.Fields("IS_NULLABLE"), 32, 0)
    'Debug.Print sh.Fields("COLUMN_NAME"), fatr(i), "notnull=" & ((fatr(i) And 32) = 0)
sh.MoveNext
Wend
Set sh = Nothing
End Sub


Public Property Get DefaultValues() As CParam 'GET ALL
Set DefaultValues = mDefaultValues
'NEED CALL UpdateDefaultValues ""
End Property
Public Property Set DefaultValues(cp As CParam) 'Update DefaultValues SET
Set mAddDefValues = cp
If mAddDefValues Is Nothing Then Set mAddDefValues = New CParam
mDefaultValues.AddParams mAddDefValues.Source
End Property
Public Property Let DefaultValues(cpar) 'Update DefaultValues LET
'If mAddDefValues Is Nothing Then Set mAddDefValues = New CParam
mAddDefValues.AddParams cpar
mDefaultValues.AddParams mAddDefValues.Source
'Debug.Print xOwner.Name & "  MATRIX Let DefaultValues"
End Property


'Sub UpdateDefaultValues(arNamesValues)
''If (m_RowMode And 1&) = 0 Then Exit Sub
'If mDefValues Is Nothing Then Set mDefValues = New CParam
'
'Dim i&, fi&
'Dim arnv
''If mDefValues Is Nothing Then arnv = arNamesValues Else arnv = xMain.CParam(arNamesValues).AddParams(mDefValues.Source)
''arnv = xMain.CParam(arNamesValues).AddParams(mDefValues.Source)
'mDefValues.AddParams arnv
''Debug.Print Join2(arnv, ";")
'
'If VarType(arnv) And vbArray Then
'On Error Resume Next
'For i = 0 To UBound(arnv) Step 2
'    fi = FieldIndex("" & arnv(i))
'    If fi > -1 Then
'        If Not rs.Fields(fi).Properties("ISAUTOINCREMENT") Then mDefValues(rs.Fields(fi).Name) = arnv(i + 1)
'    End If
'Next
'End If
''xOwner.hEvent "UpdateDefaultValues", Join2(arNamesValues, ",")
'End Sub

Private Sub CParamColumns(sar)
Dim i&
Set rs = New ADODB.Recordset
If IsArray(sar) Then
    On Error Resume Next
    For i = 0 To UBound(sar) Step 2
        rs.Fields.append sar(i), adVariant
    Next
    
'If rs.Fields.count = 0 Then rs.Fields.append "", adVariant

    Dim f, fs
    'Debug.Print m_CPColumns
    fs = Split(m_CPColumns, ",")
    gRemoveIndex fs, xMain.ToArray(rs.Fields, "name")
    'MsgBox Join2(xMain.ToArray(rs.Fields, "name"))
    For Each f In fs
        rs.Fields.append f, adVariant
    Next

    rs.CursorLocation = adUseClient
    rs.Open
    rs.AddNew
    For i = 0 To UBound(sar) Step 2
        rs.Fields(sar(i)) = sar(i + 1)
    Next
    
    rs.UpdateBatch
End If
End Sub
Private Sub CParamFill(sar) 'На входе отфильтрованный CParam
Dim v, nm$, i&, b As Boolean
Set rs = New ADODB.Recordset
If m_Tree Then
    rs.Fields.append "_id_", adVariant
    rs.Fields.append "_pid_", adVariant
End If
rs.Fields.append "_nm_", adVarWChar, 255
rs.Fields.append "_val_", adVariant, , adFldMayBeNull
rs.CursorLocation = adUseClient
rs.Open
If IsArray(sar) Then
    On Error Resume Next
    If m_Tree Then 'TREE FILL
        v = sar(-1)
        rs.AddNew Array("_id_", "_pid_", "_nm_", "_val_"), Array("", "ROOT", S_(v), sar)
        FillRSByCParam "", sar
    Else 'PARAM LIST (LEVEL=0)
        
        For i = 0 To UBound(sar) Step 2
            v = sar(i + 1): nm = sar(i)
            b = False
            If VarType(v) And vbArray Then
                b = b_SkipArrays
                If Not b Then If b_Nodes Then b = IIf(gFindIndex(m_Nodes, "+*" & nm) > -1, False, True) And IIf(gFindIndex(m_Nodes, "-*" & nm) > -1, True, False)
            Else
                If b_Nodes Then b = IIf(gFindIndex(m_Nodes, "+" & nm) > -1, False, True) And IIf(gFindIndex(m_Nodes, "-" & nm) > -1, True, False)
            End If
            If Not b Then rs.AddNew Array("_nm_", "_val_"), Array(nm, v)
        Next
    End If
    rs.UpdateBatch
End If
End Sub

Private Sub FillRSByCParam(ByVal pid, tar)
Dim i&, ID$, b As Boolean, par, IDx$
par = Split(S_(pid), "\")

'If UBound(pAR) > -1 And f_SkipNodes <> 0 Then
If (UBound(par) > -1) And b_Nodes Then
    IDx = par(UBound(par))
    If gFindIndex(m_Nodes, "-" & IDx) > -1 Then 'Получить парент путь без m_SkipNode
        If UBound(par) > 0 Then ReDim Preserve par(UBound(par) - 1): pid = Join(par, "\") Else pid = "" 'Путь без m_SkipNode
    Else
        IDx = ""
    End If
End If


Dim pnm$, m1&, m2&, mm$, pidmm$


For i = 0 To UBound(tar) Step 2
    pnm = tar(i)
    If IDx = "" Then
        ID = IIf(pid = "", "", pid & "\") & pnm
    Else
        ID = IIf(pid = "", "", pid & "\") & IDx & "\" & pnm
    End If
    par = Empty: par = tar(i + 1)
    b = 0: b = IsArray(par)
    If b Then b = (LBound(tar(i + 1)) = -1)

    If b And m_Tree Then  'Узел с подчиненными
        b = False
        If b_Nodes Then b = IIf(gFindIndex(m_Nodes, "+*" & pnm) > -1, False, True) And IIf(gFindIndex(m_Nodes, "-*" & pnm) > -1, True, False)
        
        If Not b Then rs.AddNew Array("_id_", "_pid_", "_nm_", "_val_"), Array(ID, pid, pnm, par)
        FillRSByCParam ID, tar(i + 1)
    ElseIf Not m_Tree Or m_FullCParam Then   'Параметр
        b = False
        If b_Nodes Then b = IIf(gFindIndex(m_Nodes, "+" & pnm) > -1, False, True) And IIf(gFindIndex(m_Nodes, "-" & pnm) > -1, True, False)

        If Not b Then rs.AddNew Array("_id_", "_pid_", "_nm_", "_val_"), Array(ID, pid, pnm, par)
        If ArrayDims(par) = 1 Then '1DIM ARRAY
            If LBound(par) = 0 Then
                pidmm = ID
                For m1 = LBound(par, 1) To UBound(par, 1)
                    mm = m1
                    ID = IIf(pid = "", "", pidmm & "\") & mm
                    rs.AddNew Array("_id_", "_pid_", "_nm_", "_val_"), Array(ID, pidmm, mm, par(m1))
                Next
            End If
        ElseIf ArrayDims(par) = 2 Then '2DIM ARRAY
            pidmm = ID
            For m2 = LBound(par, 2) To UBound(par, 2)
            For m1 = LBound(par, 1) To UBound(par, 1)
                mm = m1 & "," & m2
                ID = IIf(pidmm = "", "", pidmm & "\") & mm
                rs.AddNew Array("_id_", "_pid_", "_nm_", "_val_"), Array(ID, pidmm, mm, par(m1, m2))
            Next
            Next
        End If
    End If
Next
End Sub

Function UpdateRows(xmx, ByVal sIndex, Optional ByVal bAddIfNoExist As Boolean, Optional ByVal uc)
'Обновить коресспондирующие поля при соблюдении условия uc
'mx - новые данные
'sIndex - индекс
'bAddIfNoExist - добавлять при отсутствии
'uc - UPDATE CONDITION expression+<~VBCODE~>  = обновлять при соблюдении условия  rs - получатель, mx - источник
'If mx Is Nothing Then Exit Function
If TypeName(xmx) <> "CMatrix" Then Exit Function
Dim mx As CMatrix: Set mx = xmx
If m_invalid Then Exit Function

If m_RSType = 5 Or m_RSType = 6 Then
    If (m_RowMode And 2) = 0 Then Exit Function
    If (m_RowMode And 1) = 0 And bAddIfNoExist Then Exit Function
End If

sIndex = S_(sIndex)
If Len(sIndex) = 0 Then Exit Function

On Error Resume Next
'On Error GoTo err_h

Dim i&, f0&, f1&
f0 = FieldIndex("" & sIndex)
If f0 = -1 Then Exit Function

f1 = mx.FieldIndex("" & sIndex)
If f1 = -1 Then Exit Function


Dim a, h, h0, h1, v, n&, nis As Boolean, b&
h = xMain.ToArray(mx.Fields) 'Список входящих полей
v = xMain.ToArray(Me.Fields) 'Список  полей получателя
a = h
gRemoveIndex a, v 'То чего нет у получателя
gRemoveIndex h, a 'Отнимаем из заголовка то чего нет у получателя
h1 = h 'Список полей с  индексом
h0 = h 'Список полей без индекса

'Debug.Print "ISAUTOINCREMENT="; rs.Fields(f0).Properties("ISAUTOINCREMENT").Value

If rs.Fields(f0).Properties("ISAUTOINCREMENT").Value Then gRemoveIndex h1, sIndex 'Удаляем индекс из списка полей rs.AddNew
gRemoveIndex h0, sIndex 'Удаляем индекс из списка полей rs.Update

If UBound(h0) < 0 Then Exit Function 'Нету кореспондирующих полей

nis = InStr(1, ",8,129,130,200,201,202,203,204,205", "," & rs.Fields(f0).Type & ",") > 0

'rs.Filter = ""
Dim fndv, fnd$

Dim dv As New CParam, m
If mDefaultValues.ParamCount And bAddIfNoExist Then
    For Each m In mDefaultValues.Map
        If gFindIndex(h, m.Name) = -1 Then dv(m.Name) = Cast(xMain.EvalAr(xOwner, m.Value), rs.Fields(m.Name).Type, rs.Fields(m.Name).Attributes And adFldIsNullable)
    Next
    'Debug.Print dv.View
End If

Dim buc As Boolean: buc = Len(S_(uc))
If buc Then
    Dim sc As ScriptControl
    uc = ExtractCode(uc, fnd) ' <~vbcode~> забираем VBCODE
    Set sc = xMain.VBScript("*+dst", xMain.Dispatch(rs.Fields), "src", xMain.Dispatch(mx.Fields), "", fnd, "state", -1)
End If

For i = 0 To mx.RecordCount - IIf(mx.NewRecordIndex > -1, 2, 1)
fndv = mx.mx(f1, i)
If VarType(fndv) > 1 Then
    If buc Then b = mx.mx(-1, i) 'Выбираем текущую запись источника
    'b: 2=UPDATE  1=ADDNEW
    b = IIf(bAddIfNoExist, 1, 0)
    If rs.RecordCount > 0 Then 'Ищем строку для обновления
        rs.AbsolutePosition = 1
        fnd = fndv
        If nis Then fnd = xMain.Quot(fnd, 39)
        fnd = sIndex & "=" & fnd
        'If nis Then fnd = sIndex & "=" & xMain.Quot(mx.mx(f1, i), 39) Else fnd = sIndex & "=" & Nz(mx.mx(f1, i))
        rs.Find fnd, , , 0
'        If rs.EOF Then b = IIf(bAddIfNoExist, 1, 0) Else b = 2  'Строка не/найдена
        If Not rs.EOF Then b = 2   'Строка найдена
    'Else
'        b = IIf(bAddIfNoExist, 1, 0)
    End If
    
    If b > 0 Then
        If b = 1 Then h = h1 Else h = h0
        v = h
        For n = 0 To UBound(v) 'prepare fields
            v(n) = Cast(mx.mx(v(n), i), rs.Fields(h(n)).Type, rs.Fields(h(n)).Attributes And adFldIsNullable)
        Next
        If buc Then sc.CodeObject.State = b
        If b = 1 Then 'ADD 1 RECORD
            If dv.ParamCount Then
                '!!!! DEFAULT VALUES !!!!
                rs.AddNew
                For Each m In dv.Map: rs.Fields(m.Name) = m.Value: Next 'FILL DEFAULT VALUES
                For m = 0 To UBound(h): rs.Fields(h(m)) = v(m): Next 'FILL NEW VALUES
                rs.UpDate
                '!!!! DEFAULT VALUES !!!!
            Else
                 rs.AddNew h, v
            End If
            GoSub evbuc 'NOTIFY ADD NEW
            
        Else 'b=2 UPDATE MATCH RECORDS
            Do While Not rs.EOF
                b = 1 'ALLOW UPDATE
                GoSub evbuc '?? ALLOW UPDATE
                If b Then rs.UpDate h, v: If Err Then Err.Clear: Exit Do 'ERR INDEX UNIQUE
                rs.Find fnd, 1      'FIND NEXT
           Loop
        End If
    End If
 'If Err Then Err.Clear
End If
Next

If buc Then sc.Reset
Refresh
Set UpdateRows = Me
Exit Function
evbuc:
    If buc Then b = sc.Eval(uc)
    Return
End Function

Function AppendRows10(ByVal Rows, Optional ByVal nIndex = Null, Optional ByVal bRetMx As Boolean)  'rows=matrix.GetRows(10)
'Return Last Index Updated /Added
'If m_RSType = 5 Or m_RSType = 6 Then Exit Sub
If m_invalid Then Exit Function
Dim rowsDims%
rowsDims = ArrayDims(Rows)
If rowsDims = 0 Then Exit Function
If rowsDims > 2 Then Exit Function


Dim i&, j&, n&
Dim h, v, ni&, b As Boolean
Dim h1, v1
Dim nid As Long, nis As Boolean 'FIND INDEX AS STRING

n = UBound(Rows, 1)
If rs.Fields.count - 1 <> n Then Exit Function

If rowsDims = 1 Then '1DIM to 2DIM
    ReDim rows1(n, 0)
    For i = 0 To n: rows1(i, 0) = Rows(i): Next
    Rows = rows1
End If


n = 0
On Error Resume Next
h = xMain.ToArray(rs.Fields) 'Список полей текущего рекордсета
ni = FieldIndex(S_(nIndex)) 'Индекс поля для обновления
nid = -1
If ni > -1 Then
    nis = InStr(1, ",8,129,130,200,201,202,203,204,205,", "," & rs.Fields(ni).Type & ",") > 0
    For i = 0 To rs.Fields.count - 1
    If rs.Fields(i).Properties("ISAUTOINCREMENT").Value Then nid = i: Exit For
    Next
    If nid > -1 Then h1 = h: gRemoveIndex h1, rs.Fields(nid).Name
End If

'n = -1
'n = 0
For i = 0 To UBound(Rows, 2) 'строки для добавления/обновления
b = True
If ni > -1 And rs.RecordCount > 0 Then 'Ищем строку для обновления
    rs.AbsolutePosition = 1
'    If nis Then rs.Find h(ni) & "=" & Chr$(39) & S_(Rows(ni, i)) & Chr$(39) Else rs.Find h(ni) & "=" & S_(Rows(ni, i))
'    b = rs.AbsolutePosition = adPosEOF 'Строка не/найдена
    rs.Find h(ni) & "=" & xMain.Quot(Rows(ni, i), IIf(nis, 39, 0))
    If rs.AbsolutePosition > 0 Then b = Cast(rs.Fields(ni), VarType(Rows(ni, i))) <> Rows(ni, i)
End If
v = pArrayLine(Rows, i)
If nIncrementField > -1 And b Then v(nIncrementField) = NewCounter

If nid > -1 Then pArrayLine v, nid, , 3

If b Then 'ADD
    If nid > -1 Then rs.AddNew: rs.Fields(ni) = Rows(ni, i): rs.UpDate h1, v Else rs.AddNew h, v
    'n = rs.RecordCount - 1
    n = n + 1
Else 'UPDATE
    If nid > -1 Then rs.UpDate h1, v Else rs.UpDate h, v
    'n = rs.AbsolutePosition - 1
End If
    
If Err Then Debug.Print "AppendRows10", Err.Description: Err.Clear
Next 'Следующая строка
'rs.Filter = ""
'AppendRows10 = n 'Index of apended row
AppendRows10 = n 'count of apended rows
Refresh

'End If
If bRetMx Then Set AppendRows10 = Me
End Function



Function AddCSVRows(rs_rows As Recordset, ByVal csv_rows, Optional noCreateRS As Long) As Long 'CSV vbTab+vbCrLf
'return count records added
Dim ars As Recordset

'If noCreateRS And m_RSType = 0 Then Exit Function
If rs_rows Is Nothing And noCreateRS Then Set ars = rs Else Set ars = rs_rows
If ars Is Nothing Then Exit Function

Dim i&, j&, n&, h(), c(), Rows, cols, v
Dim vbcode$, sz&
csv_rows = ExtractCode(csv_rows, vbcode)
csv_rows = NzS(csv_rows, "NoRecords"): If InStr(csv_rows, vbCr) = 0 Then Rows = Split(csv_rows, vbLf) Else Rows = Split(csv_rows, vbCrLf)

On Error Resume Next
Dim bDefs As Boolean, defs, sc  As ScriptControl, bCalcMX As Boolean
Dim cpLine As New CParam, uu As UUID
Dim h0, hin, d

For i = 0 To UBound(Rows)
    If i = 0 Then    'Добавляем столбцы в первой строке
        If noCreateRS Then 'Пришли новые строки для добавления
            h = xMain.ToArray(rs.Fields)
            If noCreateRS > 1 Then 'Пришло с заголовком в любом порядке столбцов
                hin = xMain.ToArray(xMain.CMatrix(0, Rows(0)).Fields) 'Приходные столбцы
                h0 = hin
                gRemoveIndex h0, h
                If UBound(h0) > -1 Then Err.Clear: Exit Function 'нет таких столбцов в этой таблице
                h = hin
            End If
            n = UBound(h)
            ReDim c(n)
        Else 'noCreateRS=0
            cols = Split(NzS(Rows(0), "NoRecords"), vbTab) 'Столбцы таблицы
            n = UBound(cols)
            ReDim h(n), c(n), defs(n)
            For j = 0 To n 'Имя и тип данных столбца
                h(j) = StrFieldAppend(Nz(cols(j), "field_" & j), ars)
                defs(j) = xMain.RightStr(h(j), "=", 1)
                h(j) = Split(h(j), "=")(0)
                '==DEFAULT VALUE FOR MATRIX ===   ...fld$50$=<<"defaultvalue"....
                If InStr(defs(j), "<<") Then
                    mDefaultValues(h(j)) = xMain.RightStr(defs(j), "<<")
                    defs(j) = RTrimChar("" & defs(j), "<<")
                End If
                If nIncrementField = -1 Then If InStr(cols(j), "+") Then nIncrementField = j ': NewCounter aVal(defs(j)): defs(j) = ""
                If Len(defs(j)) = 0 Then defs(j) = Empty Else bDefs = True: bCalcMX = bCalcMX Or InStr(defs(j), "mx(") Or InStr(defs(j), "mx.")
            Next
            ars.CursorLocation = adUseClient
            ars.Open
            If bDefs And UBound(Rows) > 0 Then
                Set sc = xMain.VBScript("*+", "Dim mx" & vbcode)
                If bCalcMX Then Set sc.CodeObject.mx = CreateEventSinkObj(uu, cpLine, Nothing) 'Dispatch(cpLine)
            End If
        End If
    End If 'i=0
    
'Start Insert Records

    'If i > 0 Or Not bhdr Or noCreateRS Then 'Следующая строка
    If i > 0 Or noCreateRS = 1 Then 'Следующая строка
    If Len(Rows(i)) Then  'Добавляем только не пустые строки
        cols = Split(Rows(i), vbTab) 'Заливаем данные
        ReDim Preserve cols(n)
        
If bDefs And bCalcMX Then cpLine.FillNames h, cols
        
        For j = 0 To n
            c(j) = cols(j)
            If nIncrementField = j And noCreateRS Then
                c(j) = NewCounter
            Else
                sz = 255
                Select Case rs.Fields(j).Type
                Case adVarWChar, adVarChar, adWChar, adChar, adBinary: sz = ars.Fields(j).DefinedSize
                Case Else: sz = -1
                End Select
                v = JSONUnEscape(c(j))
If bDefs Then
    If VarType(defs(j)) Then 'форматируем входящие данные
        'If bCalcMX Then cpLine.SVal(2 * j + 1) = c(j)
        v = sc.Eval(Replace(defs(j), "#i#", i - 1))
        If bCalcMX Then cpLine.SVal(2 * j + 1) = v
    End If
End If
                c(j) = LimitSize(Cast(v, ars.Fields(j).Type, True), sz)
            End If
        Next
        

        
        'If Len(Rows(i)) Then ars.AddNew h, c 'Добавляем только не пустые строки
        ars.AddNew h, c
    End If
    End If
Next

If Not sc Is Nothing Then sc.Reset
AddCSVRows = UBound(Rows) + (noCreateRS <> 1) 'bhdr

'UpDate NewCounter
'If nIncrementField > -1 And ars.RecordCount > 0 Then NewCounter aVal(DMinMax(nIncrementField))
If noCreateRS Then Refresh
End Function

Public Sub Requery()
Dim tmp: tmp = m_RowSource
OpenRows m_RSType, tmp, m_RowMode, Tree, , Join2(Array(Sort, Filter, Join2(gar, ",")), ";")
End Sub

Public Sub Refresh() '(Optional ByVal bNotFull As Boolean)
'If Not UpDate Then ' INITAR
If UpDate Then Exit Sub
bNotFullRefreshOwners = True
INITAR
bNotFullRefreshOwners = False
'End If
End Sub

Private Sub RefreshOwners(Optional ByVal bFull As Boolean = True)
If bEvents Then RaiseEvent OnRefresh(bFull And Not bNotFullRefreshOwners)
End Sub


Private Sub INITAR()
On Error Resume Next

nRows = 0: nCols = 0: m_NewRecordIndex = -1: m_Dirty = 0: m_New = -1
If rs Is Nothing Then Exit Sub
Dim rsrc As Boolean 'Record count

If Not m_invalid Then
If Len(m_Sort) Or Len(rs.Sort) Then rs.Sort = m_Sort
'If Err Then gDebugPrint Err.Number & " " & Err.Description & " m_Sort=" & m_Sort: Err.Clear

If Len(S_(m_Filter)) Or Len(rs.Filter) Then rs.Filter = "" & m_Filter
'If Err Then gDebugPrint Err.Number & " " & Err.Description & " m_Filter=" & m_Filter: Err.Clear
End If
ar = Empty
rsrc = Not (rs.BOF And rs.EOF)


'If ERRR Then
'Debug.Print "BOF=" & rs.BOF, "EOF=" & rs.EOF, "FC=" & rs.Fields.count, "RC=" & rs.RecordCount
'End If

'cPos = -1
curPos = -1
xIndex = -1
If rsrc Then rs.MoveFirst

If NoINITAR Then Exit Sub

If rsrc Then ar = rs.GetRows(matrixMaxFetchRows, , m_SARFields)


nCols = rs.Fields.count
If ArrayDims(ar) <> 2 Then
    ReDim ar(nCols - 1, 0)
    nRows = 0
Else
    nRows = UBound(ar, 2) + 1
End If

If rsrc Then rs.MoveFirst
If m_RowMode And 1& Then m_NewRecordIndex = nRows

'If m_RSType = 6 Then Debug.Print "rs.AbsolutePosition=" & rs.AbsolutePosition & " xparent=" & xParent.Name & " rs.Filter=" & rs.Filter
If m_Tree Or m_Group Then InitTree
'If rs.Supports(adIndex) And rs.Supports(adSeek) Then
'Debug.Print "rs.Supports(adIndex)=" & rs.Supports(adIndex)
#If LogLock Then '==================LOGLOCK=================
RefreshLocks
#End If '==================LOGLOCK=================
RefreshOwners

End Sub

'Public Sub INITARAsync(Optional endOperation = 0)
'On Error Resume Next
'nCols = rs.Fields.count
'If ArrayDims(ar) <> 2 Then
'    ReDim ar(nCols - 1, 0)
'    nRows = 0
'Else
'    nRows = UBound(ar, 2) + 1
'End If
'
'If endOperation Then
'rs.MoveFirst
'
'
'If m_RowMode And 1& Then m_NewRecordIndex = nRows
'cPos = -1
'
'xIndex = -1
'
''If m_RSType = 6 Then Debug.Print "rs.AbsolutePosition=" & rs.AbsolutePosition & " xparent=" & xParent.Name & " rs.Filter=" & rs.Filter
'If m_Tree Or m_Group Then InitTree
''If rs.Supports(adIndex) And rs.Supports(adSeek) Then
''Debug.Print "rs.Supports(adIndex)=" & rs.Supports(adIndex)
'RefreshLocks
'End If
'
'End Sub




'Private Sub FillIndex(dx&)
'ux = rs.GetRows(, 1, Array(0, dx))
'rs.MoveFirst
'cPos = 0
'Dim i&
'For i = 0 To UBound(ux, 2)
'ux(0, i) = i 'Позиция записи
'Next
'QuickSort ux, , , 1
'xIndex = dx
'End Sub

'Private Sub FillIndex(dx&)
'Dim crs As Recordset
''Set crs = rs.Clone
'If m_RSType = 8 Then Set crs = rs Else Set crs = rs.Clone
'
'crs.Sort = crs.Fields(dx).Name
'crs.MoveFirst
'ux = crs.GetRows(, 1, Array(0, dx))
'If IsArray(ux) Then
''cPos = 0
'Dim i&
'For i = 0 To UBound(ux, 2)
'ux(0, i) = i 'Позиция записи
'Next
'Else
'ReDim ux(1, -1 To -1)
'End If
''QuickSort ux, , , 1
'xIndex = dx
'End Sub


'Public Property Get FieldDefinedSize(ID)
'On Error Resume Next
'FieldDefinedSize = rs.Fields(ID).DefinedSize
''Err.Clear
'End Property
Public Property Get Fields() As Fields
Set Fields = rs.Fields
End Property

Public Property Let Events(v As Boolean)
bEvents = v
End Property

Public Property Get FieldName$(ByVal ID)
On Error Resume Next
Dim i&
If L_(ID) = -999 Then 'MD5 of matrix
    'FieldName = CalcMD5_String(ToString(ar))
    FieldName = xMain.MD5(ToString(ar))
Else
    i = FieldIndex(S_(ID))
    If i > -1 Then FieldName = rs.Fields(i).Name
End If
End Function

Public Function FieldIndex&(v$, Optional atBaseName As Boolean)
FieldIndex = -1
Dim i&, n&
If Len(v) = 0 Then Exit Function
If rs Is Nothing Then Exit Function
If rs.Fields Is Nothing Then Exit Function
'If nCols < 1 Then Exit Function
On Error Resume Next
n = -1
'If m_RSType = 4 Then
'    n = RCP.GetIndex(v) 'номер позиции CPARAM
'    If (n = -1) And (m_RowMode And 1) Then n = RCP.ParamCount
'Else
    nCols = rs.Fields.count
    i = Val(v)
    If "" & i = v Then
        If i > -1 And i < nCols Then n = i
    Else
        For i = 0 To nCols - 1
            If atBaseName Then
                If rs.Fields(i).Properties("BASECOLUMNNAME").Value = v Then n = i: Exit For
            Else
                If rs.Fields(i).Name = v Then n = i: Exit For
            End If
        Next
    End If
'End If
FieldIndex = n
End Function


'=============================================================================
'====================      TREE              ==========================================
'=============================================================================
Public Property Get TreeExpanded()
TreeExpanded = m_TreeExpandedMaxLevel
End Property
Public Property Let TreeExpanded(v)
m_TreeExpandedMaxLevel = IF0(Abs(L_(v)), 1)
End Property

Public Property Get Tree()
Tree = TreeRootVal
If IsNull(TreeRootVal) Then If m_Group Then Tree = 0
End Property
Public Property Let Tree(RootVal)
TreeRootVal = RootVal
m_Tree = 0
If VarType(RootVal) < 2 Then Exit Property
m_Tree = 1
InitTree
End Property

Public Property Get NodeIndex&(row, IndexType)   'ListIndex  родителя (row =Мой ListIndex  , IndexType=0/первого чилда=1/младшего брата=2/старшего брата=3
NodeIndex = -1
If (row > tr_rows - 1) Or (row < 0) Or (IndexType > 3) Then Exit Property
'0 PARENT INDEX
'1 FIRST CHILD INDEX
'2 NEXT SIBLING INDEX
'3 PREV SIBLING INDEX
Dim i&, n&
If IndexType = 3 Then
    n = trNodes(trm(row + 1)).Parent  'Parent tr Index
    n = trNodes(n).FirstChild  'First CHILD = самый старший брат
    For i = 1 To UBound(trm)
    If trm(i) = n Then NodeIndex = i - 1: Exit For
    Next
    If NodeIndex = row Then
        NodeIndex = -1 'Нету старшего
    Else
        n = NodeIndex
        While n <> row
        NodeIndex = n
        n = NodeIndex(NodeIndex, 2) 'Получаем индекс следующего брата
        Wend
    End If

Else
    'n = tr(Choose(IndexType + 1, 2, 6, 5), trm(row + 1))
    If IndexType = 0 Then
        n = trNodes(trm(row + 1)).Parent
    ElseIf IndexType = 1 Then
        n = trNodes(trm(row + 1)).FirstChild
    ElseIf IndexType = 2 Then
        n = trNodes(trm(row + 1)).Next
    End If
    
    If n = 0 Then Exit Property
    For i = 1 To UBound(trm)
    If trm(i) = n Then NodeIndex = i - 1: Exit For
    Next
End If


End Property

Public Property Get NodeInfo(row, nParam)   'Чтение параметров ветки
If (row > tr_rows - 1) Or (row < 0) Then Exit Property
On Error Resume Next
'NodeInfo = tr(Param, trm(row + 1))
Select Case nParam
Case 0: NodeInfo = trNodes(trm(row + 1)).rowIndex 'NODE matrix rowindex
Case 1: NodeInfo = trNodes(trm(row + 1)).Expanded 'EXPANDED
Case 2: NodeInfo = trNodes(trm(row + 1)).Parent 'PARENT LISTINDEX
Case 3: NodeInfo = trNodes(trm(row + 1)).ChildCount 'CHILD COUNT
Case 4: NodeInfo = trNodes(trm(row + 1)).Level 'LEVEL
Case 5: NodeInfo = trNodes(trm(row + 1)).Next 'NEXT SIBLING LISTINDEX
Case 6: NodeInfo = trNodes(trm(row + 1)).FirstChild 'FIRST CHILD LISTINDEX
Case 7: NodeInfo = trNodes(trm(row + 1)).Lines 'LINES
End Select

'0 = RowIndex ar()
'1 = Expanded
'2 = tr_Parent
'3 = ChildCount
'4 = Level
'5 = Next
'6 = FirstChild
'7 = Lines 1..Level 0=[   ], 1=[  |  ], 2=[ |- ] , 3=[ |_ ]
End Property



Private Function InitTree() As Boolean
Dim i&
ReDim trNodes(0) As TRNode  'uIndex, Expanded, tr_Parent, ChildCount, Level, Next, FirstChild, Lines
'Номера строк для обработки
If nRows = 0 Then tre = Array() Else ReDim tre(nRows - 1) As Boolean
'For i = 0 To nRows - 1 'Максимальный индекс строки в матрице
'    tre(i) = True 'Это значит их надо обрабатывать
'Next i
If m_Group Then
    If nRows > 0 Then GetGroups 0, 0    'Заливаем ГРУППЫ
Else
    If nRows > 0 Then
    GetSubTree TreeRootVal, m_TreeExpandedMaxLevel, True         'Заливаем ROOT
    End If
End If
Fill_TRM
End Function

Property Get TreeColumn(RootColumn)
If RootColumn Then TreeColumn = tr_id Else TreeColumn = tr_pid
End Property
Property Let TreeColumn(RootColumn, v)
If RootColumn Then tr_id = v Else tr_pid = v
End Property

Property Get RootNode() As Boolean
RootNode = m_RootNode
End Property

Property Let RootNode(v As Boolean)
m_RootNode = v
INITAR
End Property

Private Function arNoTree() As Boolean
arNoTree = (VarType(ar) And vbArray) = 0
If arNoTree Then Exit Function
arNoTree = (UBound(ar, 1) = 0 And UBound(ar, 2) = 0)
End Function


'===================================================================
'============GROUPING========================================
Private Function TRIndex(grpValue, col&, root&) As Long 'Find existing group by value
TRIndex = -1
Dim i&
For i = 0 To UBound(trNodes)
If trNodes(i).Parent = root Then If ar(col, trNodes(i).rowIndex) = grpValue Then TRIndex = i: Exit For
Next
End Function

Public Property Get GroupBy()
On Error Resume Next
GroupBy = m_Group
End Property
Public Property Let GroupBy(v) 'v=  field_name [ASC|DESC] [, ....]
gar = Split(v, ",")
m_Group = IsArray(gar)
m_Group = UBound(gar) > -1
If rs Is Nothing Then Exit Property
If m_Group Then InitTree
End Property

Public Property Get GroupInfo(row, Param) As Long
GroupInfo = -1
If m_Group Then
    If (row > tr_rows - 1) Or (row < 0) Then Exit Property
    Dim i&, col&, order&
    i = trNodes(trm(row + 1)).Level
    GroupParams i, col, order
    If Param = 0 Then 'Level
        GroupInfo = i
    ElseIf Param = 1 Then 'FieldIndex
        GroupInfo = col
    ElseIf Param = 2 Then 'SortOrder
        GroupInfo = order
    End If
End If
End Property

Private Sub GroupParams(grpIndex&, col&, order&)
Dim gr
gr = Split(gar(grpIndex) & " ", " ")
col = FieldIndex("" & gr(0)) 'FieldIndex
order = 0
If col = -1 Then Exit Sub
If gr(1) = "ASC" Then order = 1 Else If gr(1) = "DESC" Then order = -1
End Sub

Private Sub GetGroups(ByVal grpIndex&, ByVal root&)
Dim i&, n&, ng&
Dim col&, order&
Dim gi&, ok As Boolean
'Dim nx, nxg
n = UBound(trNodes) 'Индекс максимального элемента
If n = 0 Then 'РУТ НОДА токо есть = Инициализируем её
trNodes(0).rowIndex = -1 'Индекс строки матрицы
trNodes(0).Expanded = True  'Expanded
trNodes(0).Parent = -1 'Родитель
'trNodes(0).ChildCount = 0 'ChildCount
trNodes(0).Level = -1  'Level
'trNodes(0).Next = 0 'NextSibling
trNodes(0).FirstChild = -1 'FirstChild
'trNodes(0).Lines = ""  'RootLines

End If

For i = IIf(order > -1, 0, UBound(ar, 2)) To IIf(order > -1, UBound(ar, 2), 0) Step IIf(order > -1, 1, -1)
If Not tre(i) Then
    ok = True
    If grpIndex > 0 Then 'Надо фильтровать предыдущие группы
        ok = False
        For gi = grpIndex - 1 To 0 Step -1
        GroupParams gi, col, order
        'If ar(col, tr(0, root)) = ar(col, i) Then ok = True Else ok = False: Exit For   '- старшая группа должна быть равна группе Parenta
        If ar(col, trNodes(root).rowIndex) = ar(col, i) Then ok = True Else ok = False: Exit For    '- старшая группа должна быть равна группе Parenta
        Next
    End If
    
    If ok Then
    GroupParams grpIndex, col, order
    ng = TRIndex(ar(col, i), col, root) 'Индекс группы

    If ng = -1 Then 'Надо добавить группу
        n = UBound(trNodes): n = n + 1: ReDim Preserve trNodes(n) As TRNode  'Добавляем ноду Группы
        trNodes(n).rowIndex = i  'Индекс строки матрицы
        'trNodes(n).Expanded = 0  'Expanded
        trNodes(n).Parent = root  'Родитель
        'trNodes(n).ChildCount = 0  'ChildCount
        trNodes(n).FirstChild = -1  'ChildCount
        'trNodes(n).Next = 0  'NextSibling
        
        If trNodes(root).ChildCount > 0 Then trNodes(LastChildIndex(root)).Next = n  'PerviosNextSibling=n
        If trNodes(root).ChildCount = 0 Then trNodes(root).FirstChild = n    'ParentFirstChild=n
        trNodes(root).ChildCount = trNodes(root).ChildCount + 1   'Parent ChildCount+1
        trNodes(n).Level = trNodes(root).Level + 1   'Level
        trNodes(n).Expanded = trNodes(n).Level + 1 < m_TreeExpandedMaxLevel
        'trNodes(n).Lines = "" 'RootLines
        If trNodes(root).ChildCount > 1 And order <> 0 Then SortGroupNode root, col, order
        ng = n
        If grpIndex < UBound(gar) Then GetGroups grpIndex + 1, ng 'Добавим подгруппу
    End If
    If Not tre(i) Then
        n = UBound(trNodes): n = n + 1: ReDim Preserve trNodes(n) As TRNode 'Добавляем ноду записи
        trNodes(n).rowIndex = i  'Индекс строки матрицы
        'trNodes(n).Expanded = False  'Expanded
        trNodes(n).Parent = ng 'Родитель
        trNodes(n).FirstChild = -1 'FirstChild
        'trNodes(n).ChildCount = 0  'ChildCount
        If trNodes(ng).ChildCount > 0 Then trNodes(LastChildIndex(ng)).Next = n   'PerviosNextSibling=n
        If trNodes(ng).ChildCount = 0 Then trNodes(ng).FirstChild = n  'ParentFirstChild=n
        trNodes(ng).ChildCount = trNodes(ng).ChildCount + 1  'ParentChildCount+1
        trNodes(n).Level = trNodes(ng).Level + 1   'Level
        trNodes(n).Expanded = trNodes(n).Level + 1 < m_TreeExpandedMaxLevel
        'trNodes(n) = "" 'RootLines
        tre(i) = True 'False
    End If
    End If
End If
Next
End Sub

Private Function LastChildIndex&(ByVal index)
Dim i&
i = trNodes(index).FirstChild 'FirstChild
If i > -1 Then index = i
While trNodes(index).Next  'NextSibling
index = trNodes(index).Next
Wend
LastChildIndex = index
End Function

Private Sub SortGroupNode(root&, col&, order&)
'новый нод самый последний в списке
Dim i&, n&, nd() As Long
ReDim nd(trNodes(root).ChildCount - 1) As Long
Dim ic&: ic = trNodes(root).FirstChild: If ic > -1 Then i = ic

nd(0) = i 'Индекс нода
While trNodes(i).Next
i = trNodes(i).Next  'NextSibling
n = n + 1
nd(n) = i 'Индекс нода
Wend
On Error Resume Next 'POSIBLE COMPARE ERROR
For i = UBound(nd) To 0 Step -1
If i > 0 Then
    If order > 0 Then 'SORT ASC
        If ar(col, trNodes(nd(i)).rowIndex) >= ar(col, trNodes(nd(i - 1)).rowIndex) Then Exit For   'END SORT
    Else 'SORT DESK
        If ar(col, trNodes(nd(i)).rowIndex) <= ar(col, trNodes(nd(i - 1)).rowIndex) Then Exit For   'END SORT
    End If
    n = nd(i - 1): nd(i - 1) = nd(i): nd(i) = n
End If
If i = UBound(nd) Then
    trNodes(nd(i)).Next = 0
ElseIf i < UBound(nd) Then
    trNodes(nd(i)).Next = nd(i + 1)
End If
Next
trNodes(root).FirstChild = nd(0)
End Sub

'============GROUPING========================================
'===================================================================



Private Function GetSubTree(ByVal ParentID, MaxLevel As Integer, Optional ExpandNode As Boolean) As Long 'Заливаем tr() из массива
Dim i&, n&, ic&, b&
Dim Parent&, Level&, rn&
'Debug.Print Hex(VarType(ar)) & " " & Hex(vbArray)
If arNoTree Then Exit Function
n = UBound(trNodes) 'Индекс максимального элемента


If n = 0 Then 'РУТ НОДА токо есть = Инициализируем её
trNodes(0).rowIndex = -1  'Индекс строки матрицы
trNodes(0).Expanded = True  'Expanded
trNodes(0).Parent = -1  'Родитель
trNodes(0).ChildCount = 0 'ChildCount
trNodes(0).Level = -1  'Level
'trNodes(0).Next = -1 'Empty 'NextSibling
trNodes(0).FirstChild = -1 'FirstChild
'trNodes(0).Lines = "" 'RootLines
End If

'If Not IsEmpty(tr(0, 0)) Then n = n + 1 'Номер нового элемента
Parent = 0: Level = 0
For i = 1 To n  'Ищем индекс Родителя и его уровень
    If trNodes(i).rowIndex > -1 Then
        If "" & ar(tr_id, trNodes(i).rowIndex) = "" & ParentID Then
            Parent = i: Level = trNodes(i).Level + 1
            If ExpandNode Then trNodes(i).Expanded = True  'Expanded потомушто счас посчитаем детей
            Exit For
        End If
    End If
Next
If m_RootNode And (n = 0) Then rn = m_RootNode


Do
For i = 0 To UBound(ar, 2)
If Not tre(i) Then 'Смотри токо необработанные строки
    On Error Resume Next
    If ar(tr_id, i) = ar(tr_pid, i) Then ar(tr_pid, i) = Rnd 'TreeRootVal
    
    If rn And m_RootNode And (n = 0) Then b = (ar(tr_id, i) & "" = "" & ParentID) Else b = (ar(tr_pid, i) & "" = "" & ParentID)
    If b Then
        ic = ic + 1
        If MaxLevel > Level Then
            n = UBound(trNodes): n = n + 1 'Номер нового элемента
            ReDim Preserve trNodes(n) As TRNode 'Добавляем ноду
            trNodes(n).rowIndex = i   'Индекс строки матрицы
            trNodes(n).Parent = Parent  'Parent tr Index ((Родитель))
            trNodes(n).Level = Level  'Level
            trNodes(n).FirstChild = -1 'FirstChild
            'trNodes(n).Next = -1 'Next
            'If Parent > 0 And Expan Then tr(1, Parent) = Expan'Експандируем родителя
'            Debug.Print "REGNODE " & n & " = " & ar(2, i)
            If ic = 1 Then trNodes(Parent).FirstChild = n   'FirstChild tr Index  ((Первый деть))
            'Надо проверить наличие старшего брата и прописать ему  код текущего
            If trNodes(n - 1).Parent = Parent Then
                trNodes(n - 1).Next = n  'Next tr Index       ((Я Следующий брат)) !!!!!!!!!!!!!!!
            Else 'Поищем старшего брата путем поиска меня в сыновях родителя
                If trNodes(Parent).FirstChild <> n Then
                'Debug.Print "У меня есть старший брат"
                    b = trNodes(Parent).FirstChild  'Берем первого сына моего отца
                    While b > -1
                    If trNodes(b).Next = 0 Then trNodes(b).Next = n: b = -1 Else b = trNodes(b).Next
                    Wend
                End If
            End If
            tre(i) = True '0 'ЭТА СТРОКА ОБРАБОТАНА
            trNodes(n).ChildCount = GetSubTree(ar(tr_id, i), MaxLevel, (MaxLevel - 1 > Level) And ExpandNode)   'ChildCount
            trNodes(n).Expanded = trNodes(n).ChildCount > 0 And (MaxLevel - 1 > Level)   'm_TreeExpandAll 'Expanded
            'trNodes(n).Lines = ""
        End If
    End If
End If
Next

b = 0
If rn Then rn = 0: b = 1
Loop While rn Or (b > 0) 'ROOT NO ADD



GetSubTree = ic

If Level = 0 And trNodes(0).ChildCount = 0 And MaxLevel > Level Then
    trNodes(0).ChildCount = ic  'ChildCount ROOT
End If

'Err.Clear
End Function

Private Sub Fill_TRM() 'Упорядочить trm
'**************************************************************************
Dim Lines$, cln$
Dim lv&
Dim i As Long, n As Long, c As Long, nf As Boolean
n = UBound(trNodes) 'Количество нодов
ReDim trm(n) 'На всякий случай когда все эксандированы
tr_rows = 0
trm(0) = 0
If n = 0 Then Exit Sub
i = 0: c = 0
While c <= n And i <= n
trm(c) = i 'Сначала самый первый нод = 0
'If tr(0, i) > -1 Then Debug.Print "Node=" & ar(2, tr(0, i)) 'ar(2, tr(0,i))
nf = 0 'Следующий не найден
lv = trNodes(i).Level  'LEVEL
'========  ROOT LINES =========
If lv > -1 Then
'0=[   ], 1=[  |  ], 2=[ |- ] , 3=[ |_ ]
If trNodes(i).Next = 0 Then
    trNodes(i).Lines = Left$(Lines, lv) & "2"
    cln = 0
Else
    trNodes(i).Lines = Left$(Lines, lv) & "3"
    cln = 1
End If
Lines = Left$(Lines, lv) & "" & cln
End If
'========  ROOT LINES =========
'Надо вычислить следующий нод
If trNodes(i).ChildCount > 0 And trNodes(i).Expanded Then   'Проверяем есть дети уи экспандированость текущево нода
    i = trNodes(i).FirstChild: nf = 1  'Получаем первого чилда
Else 'Детей нету - надо искать следующий нод
    If trNodes(i).Next Then 'Есть следующий
        i = trNodes(i).Next: nf = 1
    Else 'Нету следующего
        While trNodes(i).Parent > -1 And nf = 0  'Пока не найден парент со следжующим ищем его
            i = trNodes(i).Parent  'Берем парента
            If i > -1 Then 'Парент есть
                If trNodes(i).Next Then  'Берем у парента его следующего
                    i = trNodes(i).Next: nf = 1
                End If
            End If
        Wend
        If Not nf Then i = n + 1
    End If
End If
c = c + 1

Wend
tr_rows = c - 1
ReDim Preserve trm(c - 1)
End Sub

Public Function ExpandChilds(key)
On Error Resume Next
Dim v, te0&
v = mx(0, GoNode(key))
ExpandChilds = v
GetSubTree v, 100, 1
Fill_TRM
End Function

Public Property Get Expanded(index As Long) As Boolean
If Not (m_Tree Or m_Group) Then Exit Property
If (index > tr_rows - 1) Or (index < 0) Then Exit Property
Expanded = trNodes(trm(index + 1)).Expanded
End Property
Public Property Let Expanded(index As Long, v As Boolean)
If Not (m_Tree Or m_Group) Then Exit Property
If (index > tr_rows - 1) Or (index < 0) Then Exit Property
If UBound(trm) > -1 Then
    If index + 1 > UBound(trm) Or index < 0 Then Exit Property 'Индекс вне диапазона
    If trNodes(trm(index + 1)).ChildCount = 0 Then Exit Property  'Нету чего экспандировать
    If v <> trNodes(trm(index + 1)).Expanded Then trNodes(trm(index + 1)).Expanded = v Else Exit Property
End If
If v Then If trNodes(trm(index + 1)).rowIndex > -1 Then If m_Tree Then GetSubTree ar(tr_id, trNodes(trm(index + 1)).rowIndex), trNodes(trm(index + 1)).Level + 2       'True
Fill_TRM
End Property

Public Function GoNode&(ByVal key) 'Делаем Экспандированной указанную ноду возвращает индекс ноды в списке строк
GoNode = -1
If Not (m_Tree Or m_Group) Then Exit Function
If VarType(key) < 2 Then Exit Function
If tr_rows = 0 Then Exit Function
'Строим маршрут экспандирования до
Dim i As Long, fnd, par, fnd0, n&, i0&
Dim cTreeRootVal
par = Array()
fnd = Cast(key, VarType(ar(tr_id, 0)))
cTreeRootVal = Cast(TreeRootVal, VarType(ar(tr_id, 0)))
fnd0 = fnd

If m_Group Then
    m_Group = 0
    i0 = GetIndex(0, fnd) 'Индекс строки с указанным кодом
    m_Group = 1
    For i = 1 To UBound(trNodes) 'Ищем в TRNodes
        If trNodes(i).ChildCount = 0 Then If trNodes(i).rowIndex = i0 Then n = i: Exit For
    Next
    i0 = n 'Индекс в trNodes
    If n Then 'Expand to ROOT n=trNode.Index
        While trNodes(n).Parent > 0
            n = trNodes(n).Parent
            trNodes(n).Expanded = True
        Wend
    End If
Else 'm_Tree
    m_Tree = 0
    Do
        i = GetIndex(0, fnd) 'Индекс строки с указанным кодом
        If i > -1 Then
            If ar(tr_id, i) = cTreeRootVal Then
                fnd = ar(tr_id, i)
            Else
                fnd = ar(tr_pid, i)
                If gAddIndex(par, fnd) < UBound(par) Then i = -1 'Добавляем код родителя в стек
            End If
        End If
        n = n + 1
    Loop While fnd <> cTreeRootVal And i > -1 And n <= nRows  'tr_rows
    m_Tree = 1

    For i = UBound(par) To 0 Step -1 'Экспандируем до указанного нода
    GetSubTree par(i), UBound(par) - i + 1, -1
    Next
End If
Fill_TRM

If m_Group Then
    If i0 Then 'i0=trNode.Index
        For i = 1 To UBound(trm)
        If trm(i) = i0 Then GoNode = i - 1: Exit For
        Next
    End If
Else
    GoNode = GetIndex(0, fnd0)
End If
End Function

Public Function GetPathU(ID, col As Long, Optional ByVal sSeparator = Null) As String
If Not m_Tree Or NzS(ID) = "" Or nRows = 0 Then Exit Function
sSeparator = Nz(sSeparator, NzS(TreePathSeparator, "\"))
'Debug.Print "GetPathU " & id & " " & col
If NzS(ID) = "" Then Exit Function
'Строим маршрут экспандирования до
Dim i As Long, fnd, c, mt As Boolean, cRoot
c = IIf(col > -1 And col <= UBound(ar, 1), col, 0)
mt = m_Tree
m_Tree = 0
fnd = Cast(ID, VarType(ar(tr_id, 0)))
cRoot = Cast(Nz(TreeRootVal), VarType(ar(tr_pid, 0)))
'If fnd = cRoot Then Exit Function
Do
    i = GetIndex(0, fnd) 'Индекс строки с указанным кодом
    If i > -1 Then
        If ar(c, i) <> "" Then GetPathU = ar(c, i) & IIf(GetPathU = "", "", sSeparator) & GetPathU
        fnd = ar(tr_pid, i)
    End If
Loop While fnd <> cRoot And i > -1
m_Tree = mt
'xIndex = -1
End Function


Public Sub ChildNodesAR(ByVal key, car, Optional m)

'Function TreeNodes(ByVal row)
''Double() Array of child nodes rowindex+1 = AbsolutePosition as double
''sum by group  = DSum("zhsum", TreeNodes(xControl.ListIndex))
'Dim bar: bar = Array()
'If row > -1 And row < tr_rows Then bar = pEnumChildPos(0 + trm(row + 1))
'TreeNodes = bar
'End Function
'Private Function pEnumChildPos(trStartIndex&)
''Return array of child record AbsolutePosition
'Dim bar, i&
'If trNodes(trStartIndex).ChildCount Then
'    i = trNodes(trStartIndex).FirstChild
'    While i > 0
'        gAddIndex bar, pEnumChildPos(i)
'        i = trNodes(i).Next
'    Wend
'    pEnumChildPos = bar
'Else
'    pEnumChildPos = 1# + trNodes(trStartIndex).RowIndex
'End If
'End Function

If arNoTree Then Exit Sub
Dim i&, k$ ', n&
If (VarType(m) And vbArray) = 0 Then ReDim m(nRows - 1) As Boolean
'If m_Group Then 'GroupView

'Else 'TreeView
    For i = 0 To nRows - 1
    If Not m(i) Then
        If "" & ar(tr_pid, i) = "" & key Then
            k = "" & ar(tr_id, i)
            gAddIndex car, k
            m(i) = True
            ChildNodesAR k, car, m
        End If
    End If
    Next
'End If
End Sub

'=============================================================================
'====================      TREE              ==========================================
'=============================================================================


'=============================================================================
'=====================  Статистические функции по подмножеству   =================
'=============================================================================
Private Function DSAR(ByVal fld, ByVal criteria, Optional ByVal countRecords = -1)
Dim crs As Recordset: Set crs = rs.Clone(adLockReadOnly)
On Error Resume Next
If Len(criteria) Then crs.Filter = criteria Else crs.Filter = rs.Filter
Dim flds: If IsArray(fld) Then flds = fld Else flds = Array(fld)
DSAR = crs.GetRows(countRecords, , flds)
'If VarType(fld) And vbArray Then DSAR = crs.GetRows(countRecords, , fld) Else DSAR = crs.GetRows(countRecords, , Array(fld))
'Set crs = Nothing
End Function
Function DCount(Optional ByVal fld = 0, Optional ByVal criteria = "")
Dim tar, i&
DCount = 0
tar = DSAR(fld, criteria)
If ArrayDims(tar) <> 2 Then Exit Function
On Error Resume Next
For i = 0 To UBound(tar, 2)
If VarType(tar(0, i)) > 1 Then DCount = DCount + 1
Next
End Function
Function DSum(Optional ByVal fld = 0, Optional ByVal criteria = "")
Dim tar, i&, v As Double
DSum = 0
tar = DSAR(fld, criteria)
If ArrayDims(tar) <> 2 Then Exit Function
On Error Resume Next
For i = 0 To UBound(tar, 2)
DSum = DSum + tar(0, i)
Next
End Function

'Function DMinMaxLen(Optional ByVal fld = 0, Optional ByVal criteria = "", Optional ByVal bMinMax As Boolean)
'Dim tar, i&
'DMinMaxLen = Null
'tar = DSAR(fld, criteria)
'If ArrayDims(tar) <> 2 Then Exit Function
'On Error Resume Next
'DMinMaxLen = Null
'For i = 0 To UBound(tar, 2)
''If i = 0 Then DMinMaxLen = Cast(DMinMaxLen, VarType(tar(0, i)))
'If bMinMax Then 'Min
'    If Len(Nz(DMinMaxLen)) < Len(Nz(tar(0, i))) Then DMinMaxLen = tar(0, i)
'Else
'    If Len(Nz(DMinMaxLen)) > Len(Nz(tar(0, i))) Then DMinMaxLen = tar(0, i)
'End If
'Next
'End Function

Function DMinMax(Optional ByVal fld = 0, Optional ByVal criteria = "", Optional ByVal bMin As Boolean)
Dim tar, i&
DMinMax = Null
tar = DSAR(fld, criteria)
If ArrayDims(tar) <> 2 Then Exit Function
On Error Resume Next
DMinMax = Null
For i = 0 To UBound(tar, 2)
'If i = 0 Then DMinMax = Cast(DMinMax, VarType(tar(0, i)))
If bMin Then 'MIN
    If DMinMax > tar(0, i) Or IsNull(DMinMax) Then DMinMax = tar(0, i)
Else 'MAX
    If DMinMax < tar(0, i) Or IsNull(DMinMax) Then DMinMax = tar(0, i)
End If
Next
End Function
'Function DMin(Optional ByVal fld = 0, Optional ByVal criteria = "")
'Dim tar, i&
'DMin = Null
'tar = DSAR(fld, criteria)
'If ArrayDims(tar) <> 2 Then Exit Function
'On Error Resume Next
'For i = 0 To UBound(tar, 2)
'If i = 0 Then DMin = Cast(DMin, VarType(tar(0, i)))
'If tar(0, i) < DMin Or IsNull(DMin) Then DMin = tar(0, i)
'Next
'End Function
Function DValue(Optional ByVal fld = 0, Optional ByVal criteria = "")
Dim tar
DValue = Null
tar = DSAR(fld, criteria, 1)
If ArrayDims(tar) <> 2 Then Exit Function
On Error Resume Next
DValue = tar(0, 0)
End Function
Function DJoin(Optional ByVal fld = 0, Optional ByVal sSeparator = ",", Optional ByVal sQuot = "")
Dim tar, i&, v$
DJoin = ""
tar = DSAR(fld, "")
If ArrayDims(tar) <> 2 Then Exit Function
On Error Resume Next
For i = 0 To UBound(tar, 2)
v = "": v = tar(0, i)
If Len(v) Then
    If Len(DJoin) Then DJoin = DJoin & sSeparator
    DJoin = DJoin & sQuot & v & sQuot
End If
Next
End Function
'=============================================================================
'=============================================================================

Function GetValue(ByVal sdxOut, ByVal sdxIn, ByVal fnd, Optional ByVal getValueOptions = 0)
GetValue = Empty
Dim i&, v
Dim tmp$
If VarType(fnd) < 2 Then Exit Function
On Error Resume Next
i = nCols
If i = 0 Then Exit Function

Dim dxOut&, dxIn&
If VarType(sdxIn) = vbString Then dxIn = FieldIndex("" & sdxIn) Else dxIn = sdxIn
If VarType(sdxOut) = vbString Then dxOut = FieldIndex("" & sdxOut) Else dxOut = sdxOut

If dxOut >= i Or (dxOut < 0 And getValueOptions <> 2) Then Exit Function
If dxIn >= i Or dxIn < 0 Then Exit Function
'i = rs.Fields(dxIn).Type
'If i = 0 Then Exit Function

If getValueOptions = 0 Then 'FInd Value
    v = Cast(fnd, rs.Fields(dxIn).Type)
    i = GetIndex(dxIn, v)
    If i > -1 Then GetValue = mx(dxOut, i)
ElseIf getValueOptions = 1 Then 'MULTISELECT VALUE
    fnd = "," & fnd & ","
    For i = 0 To nRows - 1
    v = ar(dxIn, i)
    If VarType(v) > 1 Then If InStr(1, fnd, "," & v & ",") Then If VarType(ar(dxOut, i)) > 1 Then tmp = tmp & IIf(tmp = "", "", ",") & ar(dxOut, i)
    Next
    GetValue = tmp
ElseIf getValueOptions = 2 Then 'Tree PATH
    i = GoNode(Cast(fnd, i))
    If i > -1 Then GetValue = mx(dxOut, i)
ElseIf getValueOptions = 4 Then 'Flag List
    v = Cast(fnd, vbLong)
    For i = 0 To nRows - 1
        If Cast(ar(dxIn, i), vbLong) And v Then tmp = tmp & IIf(tmp = "", "", ",") & ar(dxOut, i)
    Next
    GetValue = tmp
End If

End Function

'Function GetIndex(ByVal dx As Long, ByVal fnd As Variant, Optional pos = "Get") As Long
Function GetIndex(ByVal sdx, ByVal fnd As Variant, Optional pos = "Get") As Long
Dim i As Long, sv, i1&, i1b As Boolean
Dim psmin As Long, psp As Long, ps&, rn&
GetIndex = -1
'If IsArray(ar) Then n = UBound(ar, 1)
On Error Resume Next

Dim dx&, n&
If VarType(sdx) = vbString Then dx = FieldIndex("" & sdx) Else dx = sdx

rn = RecordCount - 1

n = UBound(ar, 1)
If dx < 0 Or dx > n Or VarType(fnd) < 2 Then 'Неправильно указан стлобец поиска или нечево искать
    If dx = -1 Then sv = Cast(fnd, 3): If sv > -1 And sv <= rn Then GetIndex = sv
    Exit Function 'Неправильно указан стлобец поиска или нечево искать
Else
    sv = Cast(fnd, rs.Fields(dx).Type) 'Приводим поисковое значение к типу поля
    'If Len(sv) = 0 Then Exit Function
    psp = -1 'Позиция в матрице
    psmin = 100 'Длина строки
    If pos = "Get" Then 'Pos = "Get"
'        If nRows > 10000000 And Not (m_Tree Or m_Group) Then  'And m_RSType > 4 Then 'Поиск в сортированном массиве
'            If xIndex <> dx Then FillIndex dx 'Заливаем другой индекс для поиска
'
'        'Бинарный или ИНТЕРПОЛИРУЮЩИЙ метод поиска в сотрированном массиве
'        'If dx <> xIndex Then FillIndex dx
'        n = UBound(ux, 2)
'        If IsEmpty(ux(0, 0)) Then n = -1
'        If n > -1 Then
'            psmin = 0: ps = n
''            If IsNumeric(sv) Then 'ИНТЕРПОЛИРУЮЩИЙ ПОИСК по числовому значению
'            If VarType(sv) <> vbString Then 'ИНТЕРПОЛИРУЮЩИЙ ПОИСК по числовому значению
'                While ux(1, psmin) <= sv And ux(1, ps) >= sv And psp = -1 And cn < nRows
'                    i = psmin + ((sv - ux(1, psmin)) * (ps - psmin)) / (ux(1, ps) - ux(1, psmin))
'                    If sv < ux(1, i) Then
'                        ps = i - 1
'                    ElseIf sv > ux(1, i) Then
'                        psmin = i + 1
'                    Else
'                        psp = ux(0, i)
'                    End If
'                    cn = cn + 1
'                Wend
'            Else 'БИНАРНЫЙ ПОИСК по строковым значениям
'                While (psmin <= ps) And psp = -1 And cn < nRows
'                    i = (psmin + ps) \ 2
'                    If sv < ux(1, i) Then
'                        ps = i - 1
'                    ElseIf sv > ux(1, i) Then
'                        psmin = i + 1
'                    Else
'                        psp = ux(0, i)
'                    End If
'                    cn = cn + 1
'                Wend
'            End If 'IsNumeric(sv)
'        End If 'n > -1
'
'
'        Else
            psmin = 0
            For i = 0 To rn 'RecordCount - 1
            If mx(dx, i) = sv Then psp = i: psmin = Len(sv): Exit For
            Next
            
'        End If
        
    Else 'Pos <> "Get" = Find=StartIndex
        psmin = 100: psp = -1
        If Len(sv) Then
            i1 = Abs(L_(pos)) 'StartIndex
            i1b = i1 & "" <> "" & S_(pos)
            For i = i1 To rn 'RecordCount - 1
            ps = InStr(1, mx(dx, i), sv, 1)
            If ps > 0 Then
                If ps < psmin And i1b Then psmin = ps: psp = i   ': Debug.Print "psmin=" & psmin & " GetIndex=" & psp  ' Exit For
                If ps = 1 Then psp = i: Exit For 'Найдено полное соответствие по 2-м и более символам
            End If
            Next
        End If
    End If 'Pos = "Get" Pos = "Find"
    GetIndex = psp
End If
If Not IsMissing(pos) Then pos = psmin
End Function

Private Sub rcp_Changed()
Static bcall As Boolean
If bcall Then Exit Sub
bcall = True
If m_RSType = 4 Then
    CParamFill RCP(m_CParamPath)
ElseIf (m_RSType Mod 100) = 40 Then
    CParamColumns RCP.Source
ElseIf m_RSType = 44 Then
    Set rs = New ADODB.Recordset
    Dim fs ', fs0: fs0 = 1
'    If Len(m_CPColumns) Then
'        fs = Split(m_CPColumns, ",")
'        'If Len(fs(0)) = 0 Then pEnumCPFields fs, RCP.Source
'    Else
'        fs = Array()
'        pEnumCPFields fs, RCP.Source
'    End If
    fs = Array()
    If Len(m_CPColumns) Then fs = Split(Replace(m_CPColumns, "++", ""), ",")  ',,++,, columns = preserve column type
    If UBound(fs) = -1 Or InStr(m_CPColumns, "++") Then pEnumCPFields fs, RCP.Source
    
    AddCSVRows rs, Join(Split("_id_$,_pid_$,_nm_$", ","), vbTab) & vbTab & Join(fs, vbTab)
    nCols = rs.Fields.count
    pCP2RS m_CParamPath, RCP.Source, rs
Else
    bcall = False
    Exit Sub
End If
If Not bOpenRowsCall Then INITAR
bcall = False
End Sub

'Private Sub rs_RecordsetChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'Debug.Print "rs_RecordsetChangeComplete"; adStatus
'INITAR
'End Sub

Public Property Get Invalid() As Boolean
Invalid = m_invalid
End Property


Private Sub pEnumCPFields(fs, cpar) 'список и типы полей по первой строке
    Dim i&, vt&, i0&
    Dim vts, sz&, fnm$
    vts = Split(",,%,&,!,#,@,~,$,,,?,^,,,,,`", ",")
    For i = 0 To UBound(cpar) Step 2
        fnm = cpar(i): vt = VarType(cpar(i + 1))
        If (vt And vbArray) And Not b_SkipArrays Then
            pEnumCPFields fs, cpar(i + 1)
        ElseIf vt < 18 Then
            If b_Nodes Then
                If gFindIndex(m_Nodes, "-", 1) > -1 Then vt = IIf(gFindIndex(m_Nodes, "-" & fnm) > -1, -1, vt)  'Не берем
                If gFindIndex(m_Nodes, "+", 1) > -1 Then vt = IIf(gFindIndex(m_Nodes, "+" & fnm) > -1, vt, -1) 'Берем
            End If
            
            If vt > NO_INDEX Then 'Это имя разрешено для добавления
                If vt > vbNull Then
                    i0 = gFindIndex(fs, fnm)  'Ищем имя поля без указания типа
                    If i0 > NO_INDEX Then gRemoveIndex fs, fs(i0) 'Удаляем это поле которое (без указания типа)
                End If
                If vt = vbString Then If Len(S_(cpar(i + 1))) > 255 Then vt = vbVariant
                gAddIndex fs, fnm & vts(vt)
                'Debug.Print , Join2(fs, ",")
            End If
        End If
    Next
End Sub

Private Sub pCP2RS(p$, cpar, rs As Recordset)
On Error Resume Next
Dim i&, b As Boolean, ID$

'If rs.AbsolutePosition = adPosUnknown Then rs.AddNew
For i = 0 To UBound(cpar) Step 2
'    If VarType(cpar(i + 1)) < 18 Then
        If FieldIndex("" & cpar(i)) > -1 Then
            If rs.EOF And rs.BOF Then
                rs.AddNew Array("_id_", "_pid_", "_nm_"), Array("", "", "")
            End If
            rs.Fields(cpar(i)) = cpar(i + 1)
        End If
'    End If
Next
If rs.AbsolutePosition <> adPosUnknown Then If rs.EditMode = adEditAdd Then rs.UpDate

'If Err Then
'Debug.Print Err.Description
'End If

If Not b_SkipArrays Then
For i = 0 To UBound(cpar) Step 2
    If VarType(cpar(i + 1)) And vbArray Then
                                    If FieldIndex("" & cpar(i)) = -1 Then
'        b = False
'        If b_Nodes And 0 Then
'            If gFindIndex(m_Nodes, "-*", 1) > -1 Then b = IIf(gFindIndex(m_Nodes, "-*" & cpar(i)) > -1, True, False) 'Не берем
'            If gFindIndex(m_Nodes, "+*", 1) > -1 Then b = IIf(gFindIndex(m_Nodes, "+*" & cpar(i)) > -1, False, True) 'Берем
'        End If
'        If Not b Then
            rs.AddNew
            ID$ = IIf(Len(p), p & "\", "") & cpar(i)
            rs.Fields("_id_") = ID
            rs.Fields("_pid_") = p
            rs.Fields("_nm_") = cpar(i)
'            If m_SkipArrays > 0 Then 'TO JSON
'                rs.Fields("" & cpar(i)) = xMain.CParam(cpar(i + 1)).json("!")
'                rs.UpDate
'            Else
                pCP2RS ID, cpar(i + 1), rs
'            End If
'        End If
                                    End If
    End If
Next
End If

End Sub

''****************
'Sub cust11()
'Dim mx As New CMatrix, res
'mx.OpenRows 44, ar, 0, ""
'res = mx.TreeCParam(mx.mx(0, 0))
''res = mx.TreeCParam(mx.mx(0, 0), Split("tid_,lib_,uuid_,ver_,flags_,help_,base_", ","), Split("TypeID,Library,UUID,Version,Flags,Help,Base", ","))
'End Sub
''****************

