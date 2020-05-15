VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Commons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Commons.cls
'''  *<pre>Common Utility
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  */
'
''' /********************************************************
'''  * @class C_Commons
'''  * @classdesc Common Utility Class
'''  * @require C_String, C_Array, JsonConverter
'''  */
'

Dim C_String As New C_String

'//--------------------------------------------------------
'// Public Method (Function)
'//
Public Function GetElement(item As Variant, key As Variant) As Variant
''' /********************************************************
'''  * @function getElement
'''  *</br>������v�f�̎擾�@�i���܂藘�p���l�͂Ȃ��F�z��Ȃ畁�ʂ�idx�ŃQ�b�g�ł���j
'''  * @param {Variant} item �z�񂩁A�A�z�z�񂩁A�����񂩁A���l���A�E�E�E
'''  * @param {Variant} key �v�findex �������́Akey
'''  * @return {Variant} getElement �v�f
'''  *
'''  * XXX����Not Test
'''  */
    If IsArray(item) Then
        If (LBound(item) <= key) And (UBound(item) >= key) Then
            GetElement = item(key)
        Else
            GetElement = False
        End If
    ElseIf IsObject(item) Then
        If TypeName(item) = "Dictionary" Then
                GetElement = item(key)
        ElseIf TypeName(item) = "Collection" Then
                GetElement = item(key)
        Else
                GetElement = False '''FIXME�������̃I�u�W�F�N�g���ǂ��܂ŃT�|�[�g���邩�H
        End If
    Else
        GetElement = item '''�z��A�I�u�W�F�N�g�łȂ����̂�n���ꂽ�ꍇ�́A���̂܂ܕԋp�B
    End If
End Function

Public Sub Bind(target, value)
''' /********************************************************
'''  * @function Bind
'''  *<br/>�@target�ϐ��ɒl(value)��ݒ肷��B
'''  * @param {Variant} target �yByRef�z
'''  * @param {Variant} value
'''  */
    If IsObject(value) Then
        Set target = value
    Else
        target = value
    End If
End Sub

Public Sub BindAt(dicTarget, key, value)
''' /********************************************************
'''  * @function Bind
'''  *<br/>�@target�A�z�z��ϐ��i�������͔z��j�ɁA
'''  *<br/>�@key , value(Value��Object�ɂ��Ή�)
'''  *<br/>�@��ݒ肷��B
'''  * @param {Variant} dicTarget �yByRef�z
'''  * @param {Variant} key
'''  * @param {Variant} value
'''  */
    If IsObject(value) Then
        Set dicTarget(key) = value 'Dictionary
    Else
        dicTarget(key) = value     'Array (key��index)
    End If
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'  Common Class Method for ToString
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Function ToString(ByRef value)
''' /********************************************************
'''  *@function ToString
'''  *<br/> �����œn���ꂽ�l���AJSON�`���ɂ���B
'''  * ���ӁF�����ł̃I�u�W�F�N�g�����p�B�O��Interfase�Ƃ���肷��
'''  *       �ꍇ�́A�O�����C�u�����𗘗p����B
'''  * @param {Variant} value
'''  * @return {Variant} ToString
'''  */
    Dim rtn
    Err.Clear
    '
On Error Resume Next
    If VarType(value) = vbString Then
        If IsObject(value) Then
            rtn = ShowObject(value)
        Else
            'strings
            rtn = ShowString(value)
        End If
    Else
        If IsObject(value) Then
        'Object
            rtn = ShowObject(value)
        ElseIf IsArray(value) Then
        'Array
            rtn = ShowArray(value)
        ElseIf IsEmpty(value) Then
        'Variant Empyt
            rtn = "<empty>"
        ElseIf IsNull(value) Then
        'Variant Null
            rtn = "<null>"
        Else
        'Except strings �iInterger, long, Date, Boolean�j
            rtn = ShowOther(value)
        End If
    End If
    '
    If Err.Number <> 0 Then
        'unkown
        Err.Clear
        rtn = ShowUnknown(value)
    End If
    '
    ToString = rtn
    Exit Function
End Function

Public Function ShowString(value)
''' /********************************************************
'''  * @function ShowString
'''  * @param {Variant} value
'''  * @return {Variant} ShowString
'''  */
  ShowString = """" & value & """"
End Function

Public Function ShowArray(value)
''' /********************************************************
'''  * @function ShowArray
'''  * @param {Variant} value
'''  * @return {Variant} ShowArray
'''  */
''FIXME����Range��Array�Ɣ��f�����ꍇ�ɁA�Ή��ł��Ă��Ȃ��B
''         Array�̎��������`�F�b�N���āA�����ς���B
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowArray"
    Dim i
    Dim clt As Collection
    Set clt = New Collection
    Dim str As String
    Dim C_Array As C_Array
    Set C_Array = New C_Array
    Dim arytype As String
'
'On Error GoTo EXCEPTION
On Error Resume Next
    Let arytype = C_Array.TypeArrayPattern(value)
    '�z��Ɣ��肳�ꂽvalue������
    If arytype = "ArrayNotInit" Then
        ShowArray = "<" & arytype & "> []"
    ElseIf arytype = "ArrayEmpty" Then
        ShowArray = "<" & arytype & "> []"
    ElseIf arytype = "Array" Then
        For Each i In value
            clt.Add ToString(i)
        Next
        For i = 1 To clt.Count
            str = str & clt.item(i) & ","
        Next i
        ShowArray = "<" & arytype & "> [" & Left(str, Len(str) - 1) & "]"
    ElseIf arytype = "JagArray" Then
        str = ShowJagArray(value)
        ShowArray = "<" & arytype & "> " & str
    ElseIf arytype = "ArrayArray" Then
        str = ShowArrayArray(value)
        ShowArray = "<" & arytype & "> " & str
    Else
        ShowArray = "<" & arytype & "> " & ToString(value.name)
    End If
    
    '
    'ToString�ŃG���[
    If Err.Number <> 0 Then
        If IsObject(value) Then
            ShowArray = ToString(value.name)
        End If
        Err.Clear
    End If
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ' Object�J��
    If Err.Number <> 0 Then
        Console.info TypeName(value)
        Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    End If
    Err.Clear
    Set clt = Nothing
    Return
    '
'EXCEPTION:
'    GoSub FINALLY
'    Console.info ("Err.Number = " & Err.Number)
'    Console.info ("Err.Source = " & Err.source)
'    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
'    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
'    Exit Function
    '
End Function

Public Function ShowObject(value)
''' /********************************************************
'''  * @function ShowObject
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim rtn
    Err.Clear
    '
On Error Resume Next
'    Console.info TypeName(value)
'
    'Object�Ɣ��肳�ꂽvalue������
    Select Case TypeName(value)
        Case "Dictionary"
            rtn = ShowDictionary(value)
        Case "Collection"
            rtn = ShowCollection(value)
        Case "ArrayList"
            rtn = ShowArrayList(value)
        Case "Names"
            rtn = ShowNames(value)
        Case "Name"
            rtn = ShowName(value)
        Case "Range"
            rtn = ShowRange(value)
        Case "ListObjects"
            rtn = ShowListObjects(value)
        Case "ListObject"
            rtn = ShowListObject(value)
        Case "ListRows"
            rtn = ShowListRows(value)
        Case "Workbooks"
            rtn = ShowWorkbooks(value)
        Case "Workbook"
            rtn = ShowWorkbook(value)
        Case "Sheets"
            rtn = ShowWorksheets(value)
        Case "Worksheets"
            rtn = ShowWorksheets(value)
        Case "Worksheet"
            rtn = ShowWorksheet(value)
        Case "Shapes"
            rtn = ShowShapes(value)
        Case "Shape"
            rtn = ShowShape(value)
        Case Else
            If Left(TypeName(value), 2) = "C_" Then
                rtn = ShowClass(value)
            ElseIf Left(TypeName(value), 2) = "O_" Then
                rtn = ShowClass(value)
            ElseIf Left(TypeName(value), 3) = "VO_" Then
                rtn = ShowClass(value)
            Else
                If Right(TypeName(value), 10) = "Properties" Then
                    rtn = ShowProperties(value)
                Else
                    rtn = ShowArray(value)
                End If
            End If
    End Select
    '
    If Err.Number <> 0 Then
        Err.Clear
        '''value��ForEach�ŉ񂹂�Object�Ƃ��ď���
        rtn = ShowArray(value)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        '''value.Items��ForEach�ŉ񂹂�Object�Ƃ��ď���
        rtn = ShowArray(value.Items)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        rtn = "<" & TypeName(value) & ">"
    End If
    ShowObject = rtn
End Function

Public Function ShowOther(value)
''' /********************************************************
'''  * @function ShowObject
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim rtn
    Err.Clear
On Error Resume Next
    '���l,���t,�u�[���l�Ɣ��肳�ꂽvalue������
    rtn = CStr(value)
    If Err.Number <> 0 Then
        Err.Clear
        rtn = ShowUnknown(value)
    End If
    ShowOther = rtn
End Function

Public Function ShowUnknown(value)
''' /********************************************************
'''  * @function ShowObject
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    '''�������Y�ꓙ�ŁA�����B
    ShowUnknown = "<unknown:" & VarType(value) & " " & TypeName(value) & ">"
End Function

Public Function ShowDictionary(value)
''' /********************************************************
'''  * @function ShowObject
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    'Dictionary�Ɣ��肳�ꂽvalue������
    Dim i
    Dim str
    Dim clt
    Set clt = New Collection
    '
    For Each i In value.Keys
        clt.Add ToString(i) & ":" & ToString(value(i))
    Next
    For i = 1 To clt.Count
        str = str & clt.item(i) & ","
    Next i
    '
    ShowDictionary = "{" & Left(str, Len(str) - 1) & "}"
    '
    Set clt = Nothing
End Function

Public Function ShowNames(value)
''' /********************************************************
'''  * @function ShowNames
'''  * @param {Variant} value
'''  * @return {Variant} ShowNames
'''  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowNames"
    '
On Error GoTo EXCEPTION
    Dim nm As name
    Dim ary() As Variant
    Dim i As Long
    i = -1
    Dim C_String As C_String
    Set C_String = New C_String
    '
    For Each nm In value
        i = i + 1
        ReDim Preserve ary(i)
        Set ary(i) = CreateObject("Scripting.Dictionary")
        Call ary(i).Add("Name", nm.name)
        Call ary(i).Add("RefersTo", nm.RefersTo)
        If Not C_String.ContainWord(nm.RefersTo, "#REF!") Then
            Call ary(i).Add("Value", nm.RefersToRange)
        End If
        Call ary(i).Add("SheetName", C_String.Chop(C_String.SubStringBefore(nm.RefersTo, "!", 2), 1))
    Next nm
    ShowNames = Me.ToString(ary)
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ' Object�J��
    Return
    '
EXCEPTION:
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
End Function

Public Function ShowName(value)
''' /********************************************************
'''  * @function ShowName
'''  * @param {Variant} value
'''  * @return {Variant} ShowName
'''  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowName"
    '
On Error GoTo EXCEPTION
    Dim C_String As C_String
    Set C_String = New C_String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Call dict.Add("Name", value.name)
    Call dict.Add("RefersTo", value.RefersTo)
    If Not C_String.ContainWord(value.RefersTo, "#REF!") Then
        Call dict.Add("Value", value.RefersToRange)
    End If
    Call dict.Add("SheetName", C_String.Chop(C_String.SubStringBefore(value.RefersTo, "!", 2), 1))
    ShowName = Me.ToString(dict)
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ' Object�J��
    Set dict = Nothing
    Return
    '
EXCEPTION:
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
End Function

Public Function ShowListObjects(value)
''' /********************************************************
'''  * @function ShowListObjects
'''  * @param {Variant} value
'''  * @return {Variant} ShowListObjects
'''  */
    'ListObjects�Ɣ��肳�ꂽvalue������ (Names���ɎQ��#REF�����Ȃ����Ɓj
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowListObjects"
    '
On Error GoTo EXCEPTION
    '
    Dim itm
    Dim str
    Dim i
    If value.Count > 0 Then
        For Each itm In value
            str = str & ToString(itm.name) & ":" & ToString(itm.Range.Address) & ","
        Next
        If Err.Number <> 0 Then
            Err.Clear
            ShowListObjects = ShowUnknown(value)
        End If
    Else
        str = ","
    End If
    '
    ShowListObjects = "[" & Left(str, Len(str) - 1) & "]"
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ' Object�J��
    Return
    '
EXCEPTION:
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
End Function

Public Function ShowListRows(value)
''' /********************************************************
'''  * @function ShowListRows
'''  * @param {Variant} value
'''  * @return {Variant} ShowListRows
'''  */
    'ShowListRows�Ɣ��肳�ꂽvalue������ (Names���ɎQ��#REF�����Ȃ����Ɓj
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowListRows"
    '
On Error GoTo EXCEPTION
    '
    Dim itm
    Dim str
    Dim i
    If value.Count > 0 Then
        For Each itm In value
            str = str & ToString(itm) & ","
        Next
        If Err.Number <> 0 Then
            Err.Clear
            ShowListRows = ShowUnknown(value)
        End If
    Else
        str = ","
    End If
    '
    ShowListRows = "[" & Left(str, Len(str) - 1) & "]"
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ' Object�J��
    Return
    '
EXCEPTION:
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
End Function

Public Function ShowProperties(value)
''' /********************************************************
'''  * @function ShowProperties
'''  * @param {Variant} value
'''  * @return {Variant} ShowProperties
'''  */
    'Properties�Ɣ��肳�ꂽvalue������
    Dim itm
    Dim str
    Dim i
On Error Resume Next ''FIXME����item��Value���擾����ۂ�Error�ƂȂ�ꍇ������
    For Each itm In value
        str = str & ToString(itm.name) & ":" & ToString(itm.value) & ","
    Next
    '
    If Err.Number <> 0 Then
        Err.Clear
'        ShowProperties = ShowUnknown(value)
    End If
    ShowProperties = "[" & Left(str, Len(str) - 1) & "]"
End Function

Public Function ShowCollection(value)
''' /********************************************************
'''  * @function ShowCollection
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
'Collection�Ɣ��肳�ꂽvalue�������BCollection��idx��1����J�n�B
' key��Collection Object����擾�ł��Ȃ��B���AArray�Ɠ����̏o�́B
''
Dim ary() As Variant
ReDim ary(1 To value.Count)
Dim idx As Long
For idx = 1 To value.Count
    ary(idx) = C_Commons.ToString(value(idx))
Next idx
ShowCollection = "<Collection> " & VBA.Join(ary)
''
End Function

Public Function ShowArrayList(value)
''' /********************************************************
'''  * @function ShowArrayList
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
'Collection�Ɣ��肳�ꂽvalue�������BCollection��idx��1����J�n�B
' key��Collection Object����擾�ł��Ȃ��B���AArray�Ɠ����̏o�́B
''
Dim ary() As Variant
ReDim ary(0 To value.Count - 1)
Dim idx As Long
For idx = 0 To value.Count - 1
    ary(idx) = C_Commons.ToString(value.item(idx))
Next idx
ShowArrayList = "<ArrayList> " & VBA.Join(ary, ", ")
''
End Function

Public Function ShowWorkbooks(value)
''' /********************************************************
'''  * @function ShowWorkbooks
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim bk As Workbook
    Dim ary() As Variant
    Dim i As Long
    i = -1
    For Each bk In value
        i = i + 1
        ReDim Preserve ary(i)
        Set ary(i) = CreateObject("Scripting.Dictionary")
        Call ary(i).Add("CodeName", bk.codeName)
        Call ary(i).Add("Name", bk.name)
        Call ary(i).Add("Path", bk.Path)
    Next bk
    ShowWorkbooks = Me.ToString(ary)
End Function

Public Function ShowWorkbook(value)
''' /********************************************************
'''  * @function ShowWorkbook
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Call dict.Add("CodeName", value.codeName)
    Call dict.Add("Name", value.name)
    Call dict.Add("Path", value.Path)
    Call dict.Add("BuiltinDocumentProperties", value.BuiltinDocumentProperties)
    ShowWorkbook = Me.ToString(dict)
    Set dict = Nothing
End Function

Public Function ShowWorksheets(value)
''' /********************************************************
'''  * @function ShowWorksheets
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim sht As Worksheet
    Dim ary() As Variant
    Dim i As Long
    i = -1
    For Each sht In value
        i = i + 1
        ReDim Preserve ary(i)
        Set ary(i) = CreateObject("Scripting.Dictionary")
        Call ary(i).Add("Name", sht.name)
        Call ary(i).Add("Index", sht.index)
    Next sht
    ShowWorksheets = "<Worksheets> " & VBA.Right(Me.ToString(ary), Len(Me.ToString(ary)) - 8)
End Function

Public Function ShowWorksheet(value)
''' /********************************************************
'''  * @function ShowWorksheet
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Call dict.Add("Name", value.name)
    Call dict.Add("Index", value.index)
    Call dict.Add("UsedRange", value.UsedRange.Address)
    Call dict.Add("Parent", value.Parent.name)
    Call dict.Add("CustomProperties", value.CustomProperties)
    ShowWorksheet = Me.ToString(dict)
    Set dict = Nothing
End Function

Public Function ShowShapes(value)
''' /********************************************************
'''  * @function ShowShapes
'''  * @param {Variant} value
'''  * @return {Variant} ShowShapes
'''  */
    Dim shp As Shape
    Dim ary() As Variant
    Dim i As Long
    i = -1
    For Each shp In value
        i = i + 1
        ReDim Preserve ary(i)
        Set ary(i) = CreateObject("Scripting.Dictionary")
        Call ary(i).Add("Name", shp.name)
        Call ary(i).Add("SheetName", shp.Parent.name)
        Call ary(i).Add("Text", shp.TextFrame.Characters.Text)
'        Call ary(i).add("Type", shp.Type)
'        Call ary(i).add("TopLeftCell", shp.TopLeftCell.Address)
'        Call ary(i).add("BottomRightCell", shp.BottomRightCell.Address)
'        Call ary(i).add("FilBackColor", shp.Fill.BackColor.RGB)
    Next shp
    ShowShapes = Me.ToString(ary)
End Function

Public Function ShowShape(value)
''' /********************************************************
'''  * @function ShowShape
'''  * @param {Variant} value
'''  * @return {Variant} ShowShape
'''  */
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Call dict.Add("Name", value.name)
    Call dict.Add("SheetName", value.Parent.name)
    Call dict.Add("Text", value.TextFrame.Characters.Text)
    Call dict.Add("Type", value.Type)
    Call dict.Add("Id", value.ID)
    Call dict.Add("TopLeftCell", value.TopLeftCell.Address)
    Call dict.Add("BottomRightCell", value.BottomRightCell.Address)
    Call dict.Add("FillForeColor", Me.RGBToHexStr(value.Fill.ForeColor.RGB))
    Call dict.Add("FillBackColor", Me.RGBToHexStr(value.Fill.BackColor.RGB))
    Call dict.Add("Top", value.Top)
    Call dict.Add("Left", value.Left)
    Call dict.Add("Height", value.Height)
    Call dict.Add("Width", value.Width)
    ShowShape = Me.ToString(dict)
End Function

Public Function ShowClass(value)
''' /********************************************************
'''  * @function ShowClass
'''  * @param {Variant} value
'''  * @return {Variant} ShowClass
'''  */
    '�Ǝ��N���X�Ɣ��肳�ꂽvalue�������BtoSource��K���������Ă��邱��
    ShowClass = value.ToSource
End Function

Public Function ShowRange(ByRef value)
''' /********************************************************
'''  * @function ShowRange
'''  * @param {Variant} value As Range
'''  * @return {Variant} ShowRange
'''  */
    '' On evaluting as a Range
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowRange"
    ''
On Error GoTo EXCEPTION
    ''
    Dim aryary() As Variant
    Dim ary As Variant
    ''
    Dim ary2 As Variant
    Dim ary3() As Variant
    Dim rng As Range
    Dim k As Long
    If value.Areas.Count > 1 Then
        k = 0
        For Each rng In value.Areas
            ReDim Preserve ary3(k)
            ary3(k) = Me.ToString(rng)
            k = k + 1
        Next rng
        ShowRange = "<Range> [" & Join(ary3, ",") & "]"
        Exit Function
    End If
    ''
    Dim rowCnt As Long
    rowCnt = value.rows.Count
    Dim colCnt As Long
    colCnt = value.columns.Count
    '
    ''a Cell
    If rowCnt = 1 And colCnt = 1 Then
        ShowRange = "<Range> " & C_Commons.ToString(value.value)
        Exit Function
    End If
    '
    ''some Cells
    Let aryary = value.value
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    Dim temp As String
    temp = ShowArrayArray(aryary)
    temp = C_String.Chomp(C_String.Chop(temp, 1), 1)
    ShowRange = "<Range> [" & temp & "]"
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
    '
End Function

Public Function ShowListObject(ByRef value)
''' /********************************************************
'''  * @function ShowListObject
'''  * @param {Variant} value
'''  * @return {Variant} ShowListObject
'''  */
    'C_xxx�i�Ǝ��N���X�j�Ɣ��肳�ꂽvalue�������BtoSource��K���������Ă��邱��
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowListObject"
    '
On Error GoTo EXCEPTION
    '
    Dim rng As Range
    Set rng = value.Range
    '
    Dim jagary As Variant
    Dim ary As Variant
    '
    Dim rowCnt As Long
    rowCnt = rng.rows.Count
    Dim colCnt As Long
    colCnt = rng.columns.Count
    '
    ReDim jagary(rowCnt - 1)
    ReDim ary(colCnt - 1)
    '
    Dim i
    Dim j
    '
    For i = 0 To rowCnt - 1
        jagary(i) = ary
        For j = 0 To colCnt - 1
            Let jagary(i)(j) = rng(i + 1, j + 1)
        Next
    Next
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ShowListObject = ShowJagArray(jagary)
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
    '
End Function

Public Function ShowJagArray(jagary) As Variant
''' /********************************************************
'''  * @function ShowJagArray
'''  * @param {Variant} value
'''  * @return {Variant} ShowJagArray
'''  *�@���ӁFRange/ListObject��JagArray�������̂����������p
'''  *�@���ӁF�����񉻁idump�j�̂��߂̎����BJSON�`���ł͂Ȃ��B
'''  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowJagArray"
    Dim lbjagFirst As Long
    Dim ubjagFirst As Long
    Dim clt As Collection
    Set clt = New Collection
    '
    Dim str As String
    Dim strTemp As String
    '
On Error GoTo EXCEPTION
    'jagArray����������
    If UBound(jagary) <> -1 Then
        lbjagFirst = LBound(jagary(LBound(jagary))) '''jagAry�̍ŏ���Array
        ubjagFirst = UBound(jagary(LBound(jagary)))
        '
        ''�e�l���擾
        Dim r As Long
        Dim c As Long
        For r = LBound(jagary) To UBound(jagary)
            If LBound(jagary(r)) <> lbjagFirst Or UBound(jagary(r)) <> ubjagFirst Then
                Err.Raise 5
            End If
            For c = LBound(jagary(r)) To UBound(jagary(r))
                str = str & ToString(jagary(r)(c)) & ","
            Next c
            clt.Add "[" & Left(str, Len(str) - 1) & "]"
            str = ""
        Next r
        ''
        Dim item As Variant
        For Each item In clt
            strTemp = strTemp & item & ","
        Next item
        ''
        'Console.log strTemp
        ShowJagArray = "[" & Left(strTemp, Len(strTemp) - 1) & "]"
    Else
    '�z�񂪋�
        ShowJagArray = "[]"
    End If
    '
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    '
    Return
    '
EXCEPTION:   'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
End Function

Public Function ShowArrayArray(aryary As Variant) As String
''' /********************************************************
''''  * @function ShowArrayArray
''''  * @param {Variant} value
''''  * @return {Variant}
''''  * @note : for arrayarray of primitive values
''''  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowArrayArray"
    ''
    Dim r As Long
    Dim c As Long
    Dim strRtn As String
On Error GoTo EXCEPTION
    ''
    For r = LBound(aryary, 1) To UBound(aryary, 1)
        For c = LBound(aryary, 2) To UBound(aryary, 2)
            'Console.log (aryary(r, c))
            strRtn = strRtn & ToString(aryary(r, c)) & ","
        Next c
        strRtn = C_String.Chop(strRtn, 1) & ";"
    Next r
    strRtn = "[{" & C_String.Chop(strRtn, 1) & "}]"
    ShowArrayArray = strRtn
    ''
    GoSub FINALLY
    Exit Function
    ''
FINALLY:    ''Label
    ''
    Return
    ''
EXCEPTION:   ''Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'  jsonString��Object�����鋤�ʃN���X���\�b�h
'  �����Ӂ�
'  �O���ƘA�g���Ȃ��ŁA�����Ŕz��ƃI�u�W�F�N�g��JSON�`��
'  �𗘗p����ꍇ�̂݁A�g�p����B
'    JSON�`��������̊��S�Ȃ�parse�́A�O�����C�u�����𗘗p����B
'    ��JsonLib
'      https://github.com/VBA-tools/VBA-JSON
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Function ParseToJson(ByVal str As String) As Variant
'''/********************************************************
''' * @function parseToString
''' * @param {String} str jsonString
''' * @return {Variant} �����I�u�W�F�N�g�i�z��������͘A�z�z��j
''' *
''' */
ParseToJson = JsonConverter.ParseJson(str)
'
End Function

Public Function ParseToString(ByVal str As String) As Variant
'''/********************************************************
''' * @function parseToString
''' * @param {String} str jsonString
''' * @return {Variant} �����I�u�W�F�N�g�i�z��������͘A�z�z��j
''' *
''' */
Dim rtn As Variant
'
If (Left(Trim(str), 1) = "[") Then
        ParseToString = ParseToArray(str)
ElseIf (Left(Trim(str), 1) = "{") Then
        Set ParseToString = ParseToDictionary(str)
Else
        ParseToString = str
End If
'
End Function

Public Function ParseToArray(ByVal str As String, Optional delimiter As Variant = ",") As Variant
'''/********************************************************
''' * @function parseToArray
''' * @param {String} str jsonString
''' * @param {Variant} delimiter �yOptional�z
''' * @return {Variant} parseToArray �z��I�u�W�F�N�g
''' *
''' */
    Dim ary As Variant
    Dim aryTemp As Variant
    Dim strTemp As String
    Dim strTemp2 As String
    Dim i As Long
    If C_String.StartWith(Trim(str)) = "[" And C_String.EndWith(Trim(str)) = "]" Then
        strTemp = C_String.ChopDouble(Trim(str))
    Else
        strTemp = Trim(str)
    End If
    If delimiter = " " Then
        strTemp2 = C_String.SqueezeSpace(strTemp)
    ElseIf delimiter = vbLf Then
        strTemp2 = C_String.DeleteSpace(strTemp)
    Else
        strTemp2 = C_String.DeleteWhiteSpace(strTemp)
    End If
    aryTemp = Split(strTemp2, delimiter)
    '
    For i = LBound(aryTemp) To UBound(aryTemp)
        If C_String.StartWith(aryTemp(i)) = """" And C_String.EndWith(aryTemp(i)) = """" Then
            aryTemp(i) = C_String.TrimSpace(C_String.ChopDouble(aryTemp(i)))
        Else
            aryTemp(i) = C_String.TrimSpace(aryTemp(i))
        End If
    Next i
    '
    ParseToArray = aryTemp
End Function

Public Function ParseToDictionary(ByVal str As String) As Variant
'''/********************************************************
''' * @function parseToDictionary
''' * @param {String} str jsonString
''' * @return {Variant} parseToDictionary �A�z�z��I�u�W�F�N�g
''' */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.parseToDictionary"
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim strTemp As String
    Dim flag As Boolean
    Dim ary  As Variant
    Dim aryTemp As Variant
    Dim i As Long
    Dim key As Variant
    Dim value As Variant
    Dim keyTemp As Variant
    Dim valueTemp As Variant
    '
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// ����
    flag = True
    '
    If C_String.StartWith(Trim(str)) = "{" And C_String.EndWith(Trim(str)) = "}" Then
        strTemp = C_String.ChopDouble(Trim(str))
    Else
        flag = False
    End If
    ary = Split(strTemp, ",")
    For i = 0 To UBound(ary)
        If C_String.Contain(ary(i), ":") Then
            aryTemp = Split(ary(i), ":")
            keyTemp = aryTemp(0)
            valueTemp = aryTemp(1)
            Erase aryTemp
            keyTemp = C_String.Strip(keyTemp, """")
            valueTemp = C_String.Strip(valueTemp, """")
            key = Trim(C_String.TrimSpace(keyTemp))
            value = Trim(valueTemp)
            dict.Add Trim(key), Trim(value)
        Else
            dict.Add Trim(ary(i)), ""
        End If
    Next i
    '
    If flag Then
        Set ParseToDictionary = dict
    Else
        ParseToDictionary = Nothing
    End If
    '
    GoSub FINALLY
    Exit Function
FINALLY: 'Label
    '//--------------------------------------
    '// �I������
    Set dict = Nothing
    Return
    '
EXCEPTION: 'Label
    '//--------------------------------------
    ' ��O����
    GoSub FINALLY
    Console.info ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    '
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'  jsonString��Object�����鋤�ʃN���X���\�b�h�iEval���p�j
'  �����Ӂ�
'  �O���ƘA�g���Ȃ��ŁA�����Ŕz��ƃI�u�W�F�N�g��JSON�`��
'  �𗘗p����ꍇ�̂݁A�g�p����B
'  �@JSON�`��������̊��S�Ȃ�parse�́A�O�����C�u�����𗘗p����B
'  �@��JsonLib
'  �@https://code.google.com/p/vba-json/
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Function ToJson(ByRef vnt As Variant) As String
''' FIXME
ToJson = Me.ToString(vnt)
End Function

Public Function Parse(ByVal jsonString As Variant) As Object
'''/********************************************************
''' * @function parse
''' * @param {String} str jsonString
''' * @return {Object} parse �I�u�W�F�N�g
''' */
Dim objSC As Object
Dim objJson As Object
'
Set objSC = CreateObject("ScriptControl")
objSC.Language = "JScript"
'
objSC.AddCode "function toObject(str){return eval('(' + str + ')' ); } "
'
Set objJson = objSC.CodeObject.ToObject(jsonString)
'
Set Parse = objJson
'
Set objSC = Nothing
Set objJson = Nothing
'
End Function

Private Function EvalScript(ByVal expr As String, ByVal lang As String) As String
    Dim doc As Object: Set doc = CreateObject("HtmlFile")
    doc.parentWindow.execScript "document.write(" & expr & ")", lang
    If Not doc.body Is Nothing Then EvalScript = doc.body.innerHTML
End Function

''' @param vbsExpr As String
''' @return As Variant
Public Function EvalVBScipt(ByVal vbsExpr As String) As String
    EvalVBScipt = EvalScript(vbsExpr, "VBScript")
End Function

''' @param jsExpr As String
''' @return As Variant
Public Function EvalJScript(ByVal jsExpr As String) As Variant
'''/********************************************************
''' * @function EvalJScript
''' * @param {String} str jsonString
''' * @return {Variant} EvalJScript �I�u�W�F�N�g
''' */
    EvalJScript = EvalScript(jsExpr, "JScript")
End Function

Public Function EvalJScriptToObject(ByVal jsExpr As String) As Object
'''/********************************************************
''' * @function EvalJScriptToObject
''' * @param {String} str jsonString
''' * @return {Object} EvalJScript �I�u�W�F�N�g
''' */
    Set EvalJScriptToObject = EvalJScript(jsExpr)
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'  Other Common Class Method
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Public Function RGBToHexStr(lngRGB As Long) As String
'''/********************************************************
''' * @function RGBToHexStr
''' * @param {Long} lngRGB
''' * @return {String} RGBTo16
''' */
''
Dim strR As String
Dim strG As String
Dim strB As String
strR = Right$("00" & Hex$((lngRGB And "&H" & "0000FF")), 2)
strG = Right$("00" & Hex$((lngRGB And "&H" & "00FF00") / 256), 2)
strB = Right$("00" & Hex$((lngRGB And "&H" & "FF0000") / 256 ^ 2), 2)
RGBToHexStr = "#" & strR & strG & strB
''
End Function

Public Function HexStrToRGB(strRGB As String) As Long
'''/********************************************************
''' * @function HexStrToRGB
''' * @param {String} strRGB
''' * @return {Long}
''' */
' FIXME
Dim lngR As String
Dim lngG As String
Dim lngB As String
Dim strR As String
Dim strG As String
Dim strB As String
strR = VBA.Mid(strRGB, 2, 2)
strG = VBA.Mid(strRGB, 4, 2)
strB = VBA.Mid(strRGB, 6, 2)
lngR = VBA.CLng("&H" & strR)
lngG = VBA.CLng("&H" & strG)
lngB = VBA.CLng("&H" & strB)
HexStrToRGB = RGB(lngR, lngG, lngB)
''
End Function

Public Function CreateGUID() As String
''' *************************************************
Do While Len(CreateGUID) < 32
    If Len(CreateGUID) = 16 Then
        '17th character holds version information
        CreateGUID = CreateGUID & Hex$(8 + CInt(Rnd * 3))
    End If
    CreateGUID = CreateGUID & Hex$(CInt(Rnd * 15))
Loop
CreateGUID = Mid(CreateGUID, 1, 8) & "-" & Mid(CreateGUID, 9, 4) & "-" & Mid(CreateGUID, 13, 4) & "-" & Mid(CreateGUID, 17, 4) & "-" & Mid(CreateGUID, 21, 12)
'
End Function

Public Function TypeDataInCell(rng As Range) As String
''' *************************************************
Application.Volatile
Select Case True
    Case VBA.IsEmpty(rng): TypeDataInCell = "Empty"
    Case Application.IsText(rng): TypeDataInCell = "String"
    Case Application.IsLogical(rng): TypeDataInCell = "Boolean"
    Case Application.IsErr(rng): TypeDataInCell = "Error"
    Case VBA.IsDate(rng): TypeDataInCell = "Date"
    Case VBA.InStr(1, rng.Text, ":") <> 0: TypeDataInCell = "Time"
    Case VBA.IsNumeric(rng): TypeDataInCell = "Number"
End Select
'
End Function

'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub unittest()
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim ary1 As Variant
    Dim ary2 As Variant
    Dim ary3 As Variant
    str1 = "[sak , mit , �@saki , miss , ask , ao ]"
    str2 = " sak   mit   saki�@   miss   ask   ao  "
'    str3 = Range(ThisWorkbook.Names("TEST_STRING").RefersTo).text
    ary1 = Me.ParseToArray(str1)
    ary2 = Me.ParseToArray(str2, " ")
'    ary3 = Me.parseToArray(str3, vbLf)
    Console.Dump ary1
    Console.Dump ary2
    Console.Dump ary3
    '
    Dim strr1 As String
    Dim obj1 As Variant
    strr1 = "{saka : mits , �@aki : miss , sak2 : ao }"
    Set obj1 = Me.ParseToDictionary(strr1)
    Console.Dump obj1
    '
    Dim strr2 As String
    Dim obj2 As Variant
    strr2 = "{""sak1"" : mit ,     �@""sass ai"" : miss , ""sak2"" : ao ""}"
    Set obj2 = Me.ParseToDictionary(strr2)
    Console.Dump obj2
    '
    Dim strr3 As String
    Dim obj3 As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "sak", "mit"
    strr3 = C_Commons.ToString(dict)
    '
    Set obj3 = Me.ParseToDictionary(strr3) ''������"��������̂͑Ή��\
    Console.Dump obj3
    '
    Console.Dump Workbooks
    Console.Dump ThisWorkbook.Worksheets
    Console.Dump ThisWorkbook.Names
    Console.Dump ActiveSheet.ListObjects
    '
    Console.Dump Me.ParseToString("saksak")
    '
''''RGBToHexStr,HexStrToRGB
'Dim lng As Long
'Dim str As String
'str = C_Commons.RGBToHexStr(16777215)
'Console.log VBA.IIf(str = "#FFFFFF", "OK", "NG :->" & str)
'str = C_Commons.RGBToHexStr(RGB(255, 255, 255))
'Console.log VBA.IIf(str = "#FFFFFF", "OK", "NG :->" & str)
'lng = C_Commons.HexStrToRGB("#FFFF00")
'Console.log VBA.IIf(lng = 65535, "OK", "NG :->" & lng)
'lng = C_Commons.HexStrToRGB("#FFFFFF")
''Console.log VBA.IIf(lng <> 16777215, "OK", "NG :->" & lng)
'Console.log VBA.IIf(lng = 16777215, "OK", "NG :->" & lng)

''' CreateGUID
'str = C_Commons.CreateGUID()
'Console.log str
End Sub
