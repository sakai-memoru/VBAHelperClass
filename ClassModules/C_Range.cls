VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Range"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
''''  * @file C_Range.cls
''''  *<pre>Range (Cell) Common Utility
''''  *</pre>
''''  * @version 1.00
''''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
''''  */
''
'''' /********************************************************
''''  * @class C_Range
''''  *
''''  *@require C_String,C_Array,O_Template,C_File,C_Template
''''  */
''
Private CONS_RANGE_TYPE As Variant
Private Enum RANGE_TYPE
    cell = 0
    CELLS_AREA = 1
    ARRAY_HORIZONTAL = 2
    ARRAY_VERTICAL = 3
    HASHTABLE = 4
    RECORDROW = 5
    TABLEDATA = 6
End Enum
''//--------------------------------------------------------
''// Class Event
''//
Private Sub Class_Initialize()
'''' /********************************************************
''''  * @event Class_Initialize
''''  */
CONS_RANGE_TYPE = Array("CELL", "AREAS", "ROW", "COLUMN", "HASHTABLE", "RECORD", "TABLE")
''
End Sub

''//--------------------------------------------------------
''// Class Method
''//
Public Function TypeRangeNamePattern(ByVal sheetName As String _
                           , ByVal rangeparam As String _
                           , Optional ByVal bookName As Variant) As String
'''' /********************************************************
''''  *@function typeRangeNamePattern
''''  *<pre>
''''  *対象範囲の形態を返却する。
''''  *</pre>
''''  *@param {String} sheetname
''''  *@param {String} rangename
''''  *@param {Variant<String>} bookname　【Optional】
''''  *@return {Enum<String>} typeRangePattern
''''  *
''''  */
''
Dim rCnt As Long
Dim cCnt As Long
Dim sht As Worksheet
Dim rng As Range
Dim ary() As Variant
Dim itm As Variant
Dim str As Variant
Dim i As Long
Let i = 0
''
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
Set sht = Workbooks(bookName).Worksheets(sheetName)
Set rng = sht.Range(rangeparam)
rCnt = rng.Rows.Count
cCnt = rng.Columns.Count
''
If rCnt = 1 Then
    If cCnt = 1 Then
        str = CONS_RANGE_TYPE(RANGE_TYPE.cell)
    Else
        str = CONS_RANGE_TYPE(RANGE_TYPE.ARRAY_HORIZONTAL)
    End If
ElseIf rCnt = 2 Then
    str = CONS_RANGE_TYPE(RANGE_TYPE.RECORDROW)
Else
    If cCnt = 1 Then
        str = CONS_RANGE_TYPE(RANGE_TYPE.ARRAY_VERTICAL)
    ElseIf cCnt = 2 Then
        str = CONS_RANGE_TYPE(RANGE_TYPE.HASHTABLE)
    Else
        str = CONS_RANGE_TYPE(RANGE_TYPE.TABLEDATA)
    End If
End If
''
TypeRangeNamePattern = str
''
End Function

Public Function TypeRangePattern(ByRef rng As Range) As String
'''' /********************************************************
''''  *@function typeRangePattern
''''  *<pre>
''''  *対象範囲の形態を返却する。
''''  *</pre>
''''  *@param {Range} rng
''''  *@return {Enum<String>} typeRangePattern
'''  *
'''  */
''
Dim rCnt As Long
Dim cCnt As Long
Dim ary() As Variant
Dim itm As Variant
Dim str As String
Dim i As Long
i = 0
''
If rng.Areas.Count > 1 Then
    str = CONS_RANGE_TYPE(RANGE_TYPE.CELLS_AREA)
Else
    rCnt = rng.Rows.Count
    cCnt = rng.Columns.Count
    ''
    If rCnt = 1 Then
        If cCnt = 1 Then
            str = CONS_RANGE_TYPE(RANGE_TYPE.cell)
        Else
            str = CONS_RANGE_TYPE(RANGE_TYPE.ARRAY_HORIZONTAL)
        End If
    ElseIf rCnt = 2 Then
        str = CONS_RANGE_TYPE(RANGE_TYPE.RECORDROW)
    Else
        If cCnt = 1 Then
            str = CONS_RANGE_TYPE(RANGE_TYPE.ARRAY_VERTICAL)
        ElseIf cCnt = 2 Then
            str = CONS_RANGE_TYPE(RANGE_TYPE.HASHTABLE)
        Else
            str = CONS_RANGE_TYPE(RANGE_TYPE.TABLEDATA)
        End If
    End If
End If
TypeRangePattern = str
''
End Function

Public Function ToArray(ByVal sheetName As String _
                    , ByVal rangeparam As String _
                    , Optional ByVal bookName As Variant _
                    , Optional ByVal isCurrent As Variant = False) As Variant
'''' /*******************************************************
''''  *@function ToArray
''''  *<pre>指定の範囲をArrayで返却する
''''  *</pre>
''''  *@param {String} sheetname
''''  *@param {String} rangeparam
''''  *@param {Variant<String>} bookname　【Optional】
''''  *@param {Variant<Boolean>} isCurrent　【Optional】CurrentRegionを取得するか
''''  *@return {Variant<String/Long/Date>} ToArray
''''  *
''''  */
''
Dim rCnt As Long
Dim cCnt As Long
Dim sht As Worksheet
Dim rng As Range
Dim ary() As Variant
Dim itm As Variant
Dim i As Long
i = 0
''
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
''
Set sht = Workbooks(bookName).Worksheets(sheetName)
''
If isCurrent Then
    Set rng = sht.Range(rangeparam).CurrentRegion
Else
    Set rng = sht.Range(rangeparam)
End If
''
rCnt = rng.Rows.Count
cCnt = rng.Columns.Count
''
For Each itm In sht.Range(rng.Cells(1, 1), rng.Cells(rCnt, cCnt))
    ReDim Preserve ary(i)
    ary(i) = itm
    i = i + 1
Next itm
ToArray = ary
''
End Function

Public Function ToArrayArray(ByVal sheetName As String _
                    , ByVal rangeparam As String _
                    , Optional ByVal bookName As Variant _
                    , Optional ByVal isCurrent As Variant = False) As Variant
'''' /********************************************************
''''  *@function ToArrayArray
''''  *<br/>指定の範囲をArrayArrayで返却する
''''  *@param {String} sheetname
''''  *@param {String} rangeparam
''''  *@param {Variant<String>} bookname　【Optional】
''''  *@param {Variant<Boolean>} isCurrent　【Optional】CurrentRegionを取得するか
''''  *@return {Variant<String>} ToArrayArray
''''  */
''
Dim sht As Worksheet
Dim rng As Range
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
''
Set sht = Workbooks(bookName).Worksheets(sheetName)
If isCurrent Then
    Set rng = sht.Range(rangeparam).CurrentRegion
Else
    Set rng = sht.Range(rangeparam)
End If
ToArrayArray = rng.value
''
End Function

Public Function ToJagArray(ByVal sheetName As String _
                    , ByVal rangeparam As String _
                    , Optional ByVal bookName As Variant _
                    , Optional ByVal isCurrent As Variant = False) As Variant
'''' /********************************************************
''''  *@function ToJagArray
''''  *<br/>指定の範囲をArrayArrayで返却する
''''  *@param {String} sheetname
''''  *@param {String} rangeparam
''''  *@param {Variant<String>} bookname　【Optional】
''''  *@param {Variant<Boolean>} isCurrent　【Optional】CurrentRegionを取得するか
''''  *@return {Variant<String>} ToArrayArray
''''  *
''''  *@require C_Array
''''  */
''
Dim rCnt As Long
Dim sht As Worksheet
Dim rng As Range
Dim ary As Variant
ary = Array()
Dim C_Array As C_Array
Set C_Array = New C_Array
''
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
Set sht = Workbooks(bookName).Worksheets(sheetName)
If isCurrent Then
    Set rng = sht.Range(rangeparam).CurrentRegion
Else
    Set rng = sht.Range(rangeparam)
End If
rCnt = rng.Rows.Count
Dim i As Long
For i = 1 To rCnt
    Call C_Array.Add(ary, Me.RangeToArray(rng.Rows(i)))
Next i
ToJagArray = ary
Set C_Array = Nothing
''
End Function

Public Function ArrayToRange(ByVal ary As Variant, rngTarget As Range _
                            , Optional arytype As Variant = "H") As Variant
'''' /********************************************************
''''  *@function ArrayToRange
''''  *<br/>配列（Array(i)を、ターゲットRangeに設定して、Rangeを取得する
''''  *  @param {Variant} ary ByVal 配列
''''  *  @param {Range}   rngTarget
''''  *  @param {Variant} arytype  V:Vertical H;Horizontal
''''  *  @return {Variant} ArrayToRange 範囲オブジェクト
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim aryary As Variant
Dim rngRtn As Range
Dim ub As Long
Dim lb As Long
ub = UBound(ary)
lb = LBound(ary)
If UCase(arytype) = "V" Then
    Set rngRtn = rngTarget.Resize(UBound(ary) - LBound(ary) + 1, 1)
    rngRtn.value = C_Array.ArrayToArrayArrayVertical(ary)
ElseIf UCase(arytype) = "H" Then
    Set rngRtn = rngTarget.Resize(1, UBound(ary) - LBound(ary) + 1)
    rngRtn.value = C_Array.ArrayToArrayArray(ary)
End If
Set ArrayToRange = rngRtn
''
End Function

Public Function ArrayArrayToRange(ByVal aryary As Variant, rngTarget As Range) As Variant
''' /********************************************************
'''  *@function ArrayArrayToRange
'''  *<br/>配列（Array(i)を、ターゲットRangeに設定して、Rangeを取得する
'''  *  @param {Variant} aryary ByVal 二次元配列
'''  *  @param {Range}   rngTarget
'''  *  @param {Variant} type  V:Vertical H;Horizontal
'''  *  @return {Variant} ArrayArrayToRange 範囲オブジェクト
'''  *
'''  */
'
    Dim rngRtn As Range
    Dim rowCnt As Long
    Dim columnCnt As Long
    rowCnt = UBound(aryary, 1) - LBound(aryary, 1) + 1
    columnCnt = UBound(aryary, 2) - LBound(aryary, 2) + 1
    Set rngRtn = rngTarget.Resize(rowCnt, columnCnt)
    rngRtn.value = aryary
    Set ArrayArrayToRange = rngRtn
End Function


Public Function RangeToArray(ByRef rng As Range) As Variant
'''' /********************************************************
''''  *@function RangeToArray
''''  *<br/>指定の範囲を一次元配列にして返却
''''  *@param {Range} rng
''''  *@return {Variant<String>} RangeToArray
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim rCnt As Long
Dim cCnt As Long
Dim ary() As Variant
Dim itm As Variant
Dim i As Long
i = 0 'FIXME
rCnt = rng.Rows.Count
cCnt = rng.Columns.Count
''
For Each itm In Range(rng.Cells(1, 1), rng.Cells(rCnt, cCnt))
    ReDim Preserve ary(i)
    ary(i) = itm
    i = i + 1
Next itm
RangeToArray = C_Array.ChangeIndexOfBegin(ary)
''
End Function

Public Function RangeToArrayArray(ByRef rng As Range) As Variant
'''' /********************************************************
''''  *@function RangeToArrayArray
''''  *<br/>指定の範囲を二次元配列にして返却
''''  *@param {Range} rng
''''  *@return {Variant<String>} RangeToArrayArray
''''  *
''''  */
''
Dim aryary() As Variant
aryary = rng.value
RangeToArrayArray = aryary
''
End Function

Public Function RangeToJagArray(ByRef rng As Range) As Variant
'''' /********************************************************
''''  *@function RangeToJagArray
''''  *<pre>指定の範囲をジャグ配列にして返却
''''  *</pre>
''''  *@param {Range} rng
''''  *@return {Variant<String>} RangeToJagArray
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim aryary() As Variant
aryary = rng
RangeToJagArray = C_Array.ArrayArrayToJagArray(aryary)
Set C_Array = Nothing
''
End Function

Public Function RangeToArrayForAreas(ByRef rng As Range) As Variant
''' /********************************************************
'''  *@function RangeToArrayForAreas
'''  *<br/>複数のエリアを含む指定の範囲を一次元配列にして返却
'''  *@param {Range} rng
'''  *@return {Variant<String>()} RangeToArrayForAreas
'''  *
'''  * FIXME not complete unittest
'''  */
''
    Dim ary As Variant
    Dim rngArea As Range
    ary = Array()
    Dim C_Array As C_Array
    Set C_Array = New C_Array
    '
    For Each rngArea In rng.Areas
        Call C_Array.AddArray(ary, Me.RangeToArray(rngArea)) ''FIXME
    Next rngArea
    RangeToArrayForAreas = ary
    Set C_Array = Nothing
End Function

Public Function UnionEx(ParamArray pAry()) As Range
''' /********************************************************
'''  *@function UnionEx
'''  *<pre>
'''  *範囲名をUnionしたRangeを返却する
'''  *</pre>
'''  *@param {ParamArry} pAry()
'''  *@return {Range}UnionEx As Range
'''  *
'''  * FIXME not complete unittest
'''  */
'
    Dim rngRtn As Range
    Dim i As Long
    Set rngRtn = Range(pAry(0))
    '
    For i = 1 To UBound(pAry)
        Set rngRtn = Union(rngRtn, Range(pAry(i)))
    Next i
    Set UnionEx = rngRtn
    Set rngRtn = Nothing
    '
End Function

Public Function RangeToDictonary(ByRef rng As Range, ByRef aryHeader As Variant) As Dictionary
''' /********************************************************
'''  *@function RangeToDictonary
'''  *<br/>指定の範囲を連想配列にして返却
'''  *@param {Range} rng
'''  *@param {Variant<String>()} aryHeader
'''  *@return {Dictionary} RangeToDictonary
'''  * FIXME
'''  */
'
    Dim rCnt As Long
    Dim cCnt As Long
    Dim ary() As Variant
    Dim itm As Variant
    Dim ub As Long
    Dim i As Long
    i = LBound(aryHeader)
    rCnt = rng.Rows.Count
    cCnt = rng.Columns.Count
    '
    Dim dict As Dictionary
    Set dict = New Dictionary
    For Each itm In Range(rng.Cells(1, 1), rng.Cells(rCnt, cCnt))
        Call dict.Add(aryHeader(i), itm)
        i = i + 1
        If i > UBound(aryHeader) Then
            Exit For
        End If
    Next itm
    Dim j As Long
    If i <= UBound(aryHeader) Then
        For j = i To UBound(aryHeader)
            Call dict.Add(aryHeader(j), "")
        Next j
    End If
    Set RangeToDictonary = dict
    Set dict = Nothing
End Function

Public Function GetHeaderArray(ByRef rng As Range) As Variant
'''' /********************************************************
''''  *@function GetHeaderArray
''''  *<br/> get a header from a range of HASHTABLE, RECORD, TABLE
''''  *@param {Range} rng
''''  *@return {Variant()} GetHeaderArray
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim aryRtn() As Variant
Dim rngPattern As String
Let rngPattern = Me.TypeRangePattern(rng)

If rngPattern = "HASHTABLE" Then
    Dim r As Long
    For r = 1 To rng.Rows.Count
        C_Array.Add aryRtn, rng.Cells(r, 1).value
    Next r
ElseIf rngPattern = "RECORD" Then
    aryRtn = Me.RangeToJagArray(rng)(1)
ElseIf rngPattern = "TABLE" Then
    aryRtn = Me.RangeToJagArray(rng)(1)
Else
    '' no process
End If
''
Let GetHeaderArray = aryRtn
''
End Function

Public Function GetBodyJagArray(ByRef rng As Range) As Variant
'''' /********************************************************
''''  *@function GetBodyJagArray
''''  *<br/> get a header from a range of HASHTABLE, RECORD, TABLE
''''  *@param {Range} rng
''''  *@return {Variant()} GetBodyArray
''''  * If a range is Table Format, return JagArray
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim aryRtn() As Variant
Dim rngPattern As String
Let rngPattern = Me.TypeRangePattern(rng)
Dim ary() As Variant
Dim r As Long
Dim c As Long
''
If rngPattern = "HASHTABLE" Then
    For r = 1 To rng.Rows.Count
        C_Array.Add ary, rng.Cells(r, 2).value
    Next r
    C_Array.Add aryRtn, ary
ElseIf rngPattern = "RECORD" Then
    ary = Me.RangeToJagArray(rng)(2)
    C_Array.Add aryRtn, ary
ElseIf rngPattern = "TABLE" Then
    ary = Me.RangeToJagArray(rng)
    For r = 2 To UBound(ary)
        C_Array.Add aryRtn, ary(r)
    Next r
Else
    '' no process
End If
''
Let GetBodyJagArray = aryRtn
'
End Function

Public Function GetRowRange(rng As Range, idx As Long) As Range
'''' /********************************************************
''''  *@function getRowRange
''''  *<pre>　対象範囲の指定行を取得する
''''  *</pre>
''''  * @param {Range} rng　範囲
''''  * @param {Long} idx　指定行
''''  * @return {Range} getRowRange
''''  */
''
If idx > rng.Rows.Count Then
    Err.Raise 9, "C_Range.getRangeRow", "指定行がインデックスを超えています"
End If
Set GetRowRange = rng.Rows(idx)
''
End Function

Public Function GetRowArray(rng As Variant, idx As Long) As Variant
'''' /********************************************************
''''  *@function getRowRange
''''  *<pre>　対象範囲の指定行を取得する
''''  *</pre>
''''  * @param {Variant<Range>} rng　範囲
''''  * @param {Long} idx　指定行
''''  * @return {Variant<String>()} getRowArray
''''  */
''
If idx > rng.Rows.Count Then
    Err.Raise 9, "C_Range.getRowArray", "指定行がインデックスを超えています"
End If
GetRowArray = Me.RangeToArray(rng.Rows(idx))
''
End Function

'''' Add 200330
Public Function TypeValueOfCell(ByRef rng As Range) As String
''' /**************************************************
'''  *@function TypeValueOfCell
'''  *@param {Range} rng
'''  *@param {String} sheetname Optional
'''  */
Application.Volatile
'If IsMissing(sheetname) Then
'    sheetname = ActiveSheet.name
'End If
Select Case True
    Case Application.IsText(rng.value): TypeValueOfCell = "String"
    Case Application.IsLogical(rng.value): TypeValueOfCell = "Boolean"
    Case Application.IsErr(rng.value): TypeValueOfCell = "Error"
    Case VBA.IsDate(rng.value): TypeValueOfCell = "Date"
    Case VBA.InStr(1, rng.Text, ":") <> 0: TypeValueOfCell = "Time"
    Case VBA.IsNumeric(rng.value): TypeValueOfCell = "Number"
    Case VBA.IsEmpty(rng.value): TypeValueOfCell = "Empty"
End Select
'
End Function

Public Function IsInRange(inRng As Range, exRng As Range) As Boolean
'''' /**************************************************
''''  *@function IsInRange
''''  *</br>Returns True if rng1 is a subset of rng2
''''  *</br>http://codevba.com/excel/range.htm#IsInRange
''''  *</br>?IsInRange(Range("A1"), Range("A1:A2"))=True
'?IsInRange(Range("A1:A2"), Range("A1"))=False
''''  *@param {Range} inRng
''''  *@param {Range} exRng
''''  *@return {Boolean}
''''  */
''
IsInRange = False
If inRng.Parent.Parent.name = exRng.Parent.Parent.name Then
    If inRng.Parent.name = exRng.Parent.name Then
        If Union(inRng, exRng).Address = exRng.Address Then
            IsInRange = True
        End If
    End If
End If
''
End Function

Public Function GetRangeDetail(ByRef rng As Range) As Object
''' /**************************************************
'''  *@function ShowRangeDetail
'''  *@param {Range} rng
'''  *@return {Object<Dictionary>}
'''  */
'
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
Dim sht As Worksheet
Set sht = rng.Worksheet

Dim addr As String
addr = rng.Address
Dim size As Long
size = rng.Count
Dim topLeftAddress As String
Dim bottomRightAddress As String
Dim topRowNo As Long
Dim bottomRowNo As Long
Dim leftColumnNo As Long
Dim rightColumnNo As Long
Dim rowCnt As Long
Dim columnCnt As Long
Dim rowSize As Long
Dim columnSize As Long
Dim ary As Variant
Dim headerRng As Range
Dim bodyRng As Range
Dim bodyRngAddr As String
''
Dim r As Long
Let r = 1
Dim i As Long
Let i = 0
Dim field_type As String
If size = 1 Then
    field_type = "cell"
Else
    field_type = "other"
    ary = Split(addr, ":")
    topLeftAddress = ary(LBound(ary))
    bottomRightAddress = ary(UBound(ary))
    'topRowNo = Split(topLeftAddress, "$")(2)
    topRowNo = rng.Row
    bottomRowNo = Split(bottomRightAddress, "$")(2)
    leftColumnNo = Range(topLeftAddress).column
    rightColumnNo = Range(bottomRightAddress).column
    rowCnt = rng.Rows.Count
    columnCnt = rng.Columns.Count
    If Me.TypeRangeNamePattern(sht.name, rng.Address, sht.Parent.name) = "Table" Then
        Set headerRng = rng.Rows(1)
        Set bodyRng = sht.Range(topLeftAddress).Offset(1, 0).Resize(rowCnt - 1, columnCnt)
        Let bodyRngAddr = bodyRng.Address
    End If
End If
'
dict.Add "address", addr
dict.Add "value", rng.value
''
dict.Add "topLeft", topLeftAddress
dict.Add "bottomRight", bottomRightAddress
''
dict.Add "size", rng.Count
dict.Add "columnCnt", columnCnt
dict.Add "rowCnt", rowCnt
dict.Add "topRowPos", topRowNo
dict.Add "bottomRowPos", bottomRowNo
dict.Add "leftColumnPos", leftColumnNo
dict.Add "rightColumnPos", rightColumnNo
''
dict.Add "sheetName", rng.Worksheet.name
dict.Add "bookName", rng.Worksheet.Parent.name
dict.Add "headerRange", headerRng
dict.Add "bodyRange", bodyRng
dict.Add "bodyAddress", bodyRngAddr

' return and release
Set GetRangeDetail = dict
'Set dict = Nothing
End Function

Public Function GetCellDetail( _
                    ByRef rng As Range, _
                    Optional ByVal sheetName As Variant) As Object
'''' /**************************************************
''''  *@function getCellDetail
''''  *@param {Range} rng
''''  *@param {String} sheetname Optional
''''  *@return {Object<Dictionary>}
''''  */
''
If rng.Count > 2 Then ' Case : range size is 1 only.
    GetCellDetail = ""
    Return
End If
''
If IsMissing(sheetName) Then
    sheetName = rng.Worksheet.name
Else
    Set rng = Worksheets(sheetName).Range(rng.Address)
End If
''
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
''
'dict.add "name", rng.Name '' FIXME:If not set a name
dict.Add "sheetName", rng.Worksheet.name
dict.Add "address", rng.Address
dict.Add "notetext", rng.NoteText
dict.Add "value", rng.value
dict.Add "numberformat", rng.NumberFormatLocal
dict.Add "formula", rng.Formula
'dict.add "presedent", rng.Precedents '' FIXME:If not set a formula
dict.Add "color", C_Commons.RGBToHexStr(rng.Font.Color)
dict.Add "backgroundcolor", C_Commons.RGBToHexStr(rng.Interior.Color)
dict.Add "colNo", rng.Cells.column
dict.Add "rowNo", rng.Cells.Row
'' return
Set GetCellDetail = dict
''
End Function

Public Function Find( _
                ByRef rng As Range, _
                ByVal keyword As String) As Variant
'''' /**************************************************
''''  *@function Find with Perfect matching
''''  *@param {Range} rng
''''  *@param {String} keyword
''''  *@return {String()} :Addresses
''''  *@note return addresses from a range
''''  *      (not including Workbook and worksheet info).
''''  */
Dim aryRtn() As Variant
Dim firstCellAddr As String
''
Dim rngFound As Range
Set rngFound = rng.Find(keyword)
If rngFound Is Nothing Then
    Find = aryRtn
    Exit Function
Else
    '' Continue
    firstCellAddr = rngFound.Address
End If

' Find all cells
Dim i As Long
Let i = 0
Do
    If rngFound.value = keyword Then
        ReDim Preserve aryRtn(i)
        aryRtn(i) = rngFound.Address
        i = i + 1
    End If
    Set rngFound = rng.FindNext(rngFound)
Loop While firstCellAddr <> rngFound.Address
''
Find = aryRtn
''
End Function

Public Function FindInSheet( _
                    ByVal keyword As String, _
                    ByVal sheetName As Variant, _
                    Optional ByVal rangeparam As Variant, _
                    Optional ByVal bookName As Variant) As Variant
'''' /**************************************************
''''  *@function FindInSheet
''''  *@param {String} keyword
''''  *@param {String} sheetname
''''  *@param {String} rangeparam
''''  *@param {String} bookname Optional
''''  *@return {String()} : Addresses
''''  */
''
Dim rng As Range
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
If IsMissing(rangeparam) Then
    Set rng = Workbooks(bookName).Worksheets(sheetName).UsedRange
Else
    Set rng = Workbooks(bookName).Worksheets(sheetName).Range(rangeparam)
End If
''
FindInSheet = Me.Find(rng, keyword)
''
End Function

Public Sub ClearCurrentRegionByName( _
                ByVal paramname As String, _
                ByVal sheetName As Variant, _
                Optional ByVal rowOffset As Long = 0, _
                Optional ByVal colOffset As Long = 0, _
                Optional ByVal bookName As Variant)
'''' /********************************************************
''''  *@function clearCurrentRegionByName
''''  *<br/>対象シートの対象範囲名をターゲットにして、
''''  *<br/>その範囲が含まれるcurrentRegionをクリアする。
''''  *@param {String} paramname
''''  *@param {String} sheetname
''''  *@param {Long} rowOffset Optional
''''  *@param {Long} colOffset Optional
''''  *@param {String} bookname Optional
''''  *
''''  */
''
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
Dim sht As Worksheet
Set sht = Workbooks(bookName).Worksheets(sheetName)
sht.Range(paramname).Offset(rowOffset, colOffset).CurrentRegion.ClearContents
''
End Sub

Public Sub ClearCurrentRegionByKeyword( _
                ByVal keyword As Variant, _
                ByVal sheetName As Variant, _
                Optional ByVal rowOffset As Long = 0, _
                Optional ByVal colOffset As Long = 0, _
                Optional ByVal bookName As Variant, _
                Optional ByVal rangeparam As Variant)
'''' /********************************************************
''''  *@function clearCurrentRegionByKeyword
''''  *<pre>　検索キーワードでヒットしたセルから
''''  *Offsetで指定したセルをもとにしたCurrentRegionを
''''  *クリアする
''''  *</pre>
''''  *@param {String} keyword
''''  *@param {String} sheetname
''''  *@param {Long} rowOffset Optional
''''  *@param {Long} colOffset Optional
''''  *@param {Variant} bookname Optional
''''  *@param {Variant} rangeparam Optional
''''  *
''''  */
''
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim rng As Range
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
If IsMissing(rangeparam) Then
    Set rng = Workbooks(bookName).Worksheets(sheetName).UsedRange
Else
    Set rng = Workbooks(bookName).Worksheets(sheetName).Range(rangeparam)
End If
''
Dim ary As Variant
ary = Me.Find(rng, keyword)
''
If C_Array.TypeArrayPattern(ary) = "ArrayNotInit" Then
    '' continue
ElseIf C_Array.TypeArrayPattern(ary) = "ArrayEmpty" Then
    '' continue
ElseIf C_Array.TypeArrayPattern(ary) = "Array" Then
    If UBound(ary) - LBound(ary) + 1 = 1 Then
        Workbooks(bookName).Worksheets(sheetName).Range(ary(LBound(ary))).CurrentRegion.ClearContents
    End If
Else
    '' continue
End If
''
End Sub

Public Function GetCurrentRegionByKeyword( _
                    ByVal keyword As String, _
                    ByVal sheetName As Variant, _
                    Optional ByVal rowOffset As Long = 0, _
                    Optional ByVal colOffset As Long = 0, _
                    Optional ByVal bookName As Variant, _
                    Optional ByVal rangeparam As Variant) As Range
''' /**************************************************
'''  *@function getCurrentRegionByKeyword
'''  *@param {String} keyword
'''  *@param {String} sheetname
'''  *@param {Long} rowOffset Optional
'''  *@param {Long} colOffset Optional
'''  *@param {Variant} bookname Optional
'''  *@param {Variant} rangeparam Optional
'''  *@return {Range}
'''  */
''
Dim CONS_MODULE_NAME As String
CONS_MODULE_NAME = "C_Range.getCurrentRegionByKeyword"
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim rng As Range
Dim rngTarget As Range
''
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
''
''
On Error GoTo EXCEPTION
    If IsMissing(rangeparam) Then
        Set rngTarget = Workbooks(bookName).Worksheets(sheetName).UsedRange
    Else
        Set rngTarget = Workbooks(bookName).Worksheets(sheetName).Range(rangeparam)
    End If
    ''
    Set rng = Me.GetRangeByKeyword(keyword, sheetName, rowOffset, colOffset, rngTarget.Address, bookName)
    ''
    Set GetCurrentRegionByKeyword = rng.CurrentRegion
    ''
    GoSub FINALLY
    Exit Function
    ''
FINALLY: 'Label
    Set rngTarget = Nothing
    Return
    ''
EXCEPTION: 'Label
    GoSub FINALLY
    Dim desc As String
    desc = CONS_MODULE_NAME & vbCr & Err.description
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Description = " & desc)
    Call Err.Raise(Err.Number, Err.source & "<-" & CONS_MODULE_NAME, desc)
    Exit Function
    ''
End Function

Public Function GetCurrentRegionByName( _
                    ByVal rangeparam As String, _
                    ByVal sheetName As String) As Range
''' /********************************************************
'''  *@function getCurrentRegionByName
'''  *<br/>対象シートの対象範囲名をターゲットにして、
'''  *<br/>その範囲が含まれるcurrentRegionのRangeを取得する。
'''  *@param {String} rangeparam
'''  *@param {String} sheetname
'''  *@return {Range}
'''  */
'
    Set GetCurrentRegionByName = Worksheets(sheetName).Range(rangeparam).CurrentRegion
End Function

Public Function GetRangeAreasByKeyword(ByVal strSheetName As Variant _
            , ByVal strKeyString As Variant _
            , Optional Row As Variant = 0 _
            , Optional column As Variant = 0) As Range
''' /********************************************************
'''  * @function getRangeAreasByKeyword
'''  *<pre>　検索キーワードでヒットした複数セルから
'''  *　Offsetで指定したセルを、Areasにして返却｡
'''  *</pre>
'''  * @param {String} strSheetName　検索対象シート名
'''  * @param {String} strKeyword　検索対象文字
'''  * @param {Variant} row　   【Optional】Offset行位置　【Default】0
'''  * @param {Variant} column　【Optional】Offset列位置　【Default】0
'''  * @return {Variant<Range>} getRangeByKeyword Range
'''  *
'''  * FIXME not do unittest
'''  */
'
Dim rngFindResult As Range
Dim rngFindFirst  As Range
Dim rngUsed As Range
Set rngUsed = ThisWorkbook.Worksheets(strSheetName).UsedRange
Set rngFindResult = rngUsed.Find(What:=strKeyString, Lookat:=xlWhole)
Set rngFindFirst = rngFindResult
Dim strTemp As String
Dim rng As Range
Set rng = rngFindResult.Offset(Row, column)
'
If rngFindResult Is Nothing Then
    Set rng = Nothing
Else
    Do
'            Console.Dump rng
'            Console.info rng.Address
        Set rngFindResult = rngUsed.FindNext(rngFindResult)
        Set rng = Union(rng, rngFindResult.Offset(Row, column)) ''FIXME□□隣接するCellをUnionすると内部で配列表現に変更
    Loop While (rngFindResult.Address <> rngFindFirst.Address)
End If
'
Set GetRangeAreasByKeyword = rng ''FIXME□□Union後の順番は勝手に変わる
'
End Function

Public Function GetValueByKeyword( _
                    ByVal keyword As String, _
                    ByVal sheetName As Variant, _
                    Optional ByVal rowOffset As Long = 0, _
                    Optional ByVal colOffset As Long = 1, _
                    Optional ByVal rangeparam As Variant, _
                    Optional ByVal bookName As Variant) As String
'''' /**************************************************
''''  *@function getValueByKeyword
''''  *  @note : it does not apply to Multiple searches and not_found
''''  *@param {String} keyword
''''  *@param {String} sheetname
''''  *@param {Long} rowOffset Optional
''''  *@param {Long} colOffset Optional
''''  *@param {Variant} rangeparam Optional
''''  *@param {Variant} bookname Optional
''''  *@return {String} : Cell.value
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim rng As Range
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
If IsMissing(rangeparam) Then
    Set rng = Workbooks(bookName).Worksheets(sheetName).UsedRange
Else
    Set rng = Workbooks(bookName).Worksheets(sheetName).Range(rangeparam)
End If
''
Dim ary As Variant
ary = Me.Find(rng, keyword)

If C_Array.TypeArrayPattern(ary) = "ArrayNotInit" Then
    GetValueByKeyword = ""
ElseIf C_Array.TypeArrayPattern(ary) = "ArrayEmpty" Then
    GetValueByKeyword = ""
ElseIf C_Array.TypeArrayPattern(ary) = "Array" Then
    If UBound(ary) - LBound(ary) + 1 = 1 Then
        GetValueByKeyword = Workbooks(bookName).Worksheets(sheetName).Range(ary(LBound(ary))).Offset(rowOffset, colOffset).value
    End If
Else
    GetValueByKeyword = ""
End If
''
End Function

Public Function GetRangeByKeyword( _
                    ByVal keyword As String, _
                    ByVal sheetName As Variant, _
                    Optional ByVal rowOffset As Long = 0, _
                    Optional ByVal colOffset As Long = 0, _
                    Optional ByVal rangeparam As Variant, _
                    Optional ByVal bookName As Variant) As Range
''' /**************************************************
'''  *@function getRangeByKeyword
'''  *@param {String} keyword
'''  *@param {String} sheetname
'''  *@param {Long} rowOffset Optional
'''  *@param {Long} colOffset Optional
'''  *@param {Variant} rangeparam Optional
'''  *@param {Variant} bookname Optional
'''  *@return {Range}
'''  */
'
Dim CONS_MODULE_NAME As String
CONS_MODULE_NAME = "C_Range.getRangeByKeyword"
Dim C_Array As C_Array
Set C_Array = New C_Array
''
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
''
Dim rngTarget As Range
Dim ary As Variant
''
On Error GoTo EXCEPTION
    If IsMissing(sheetName) Then
        sheetName = ActiveSheet.name
    End If
    If IsMissing(rangeparam) Then
        Set rngTarget = Workbooks(bookName).Worksheets(sheetName).UsedRange
    Else
        Set rngTarget = Workbooks(bookName).Worksheets(sheetName).Range(rangeparam)
    End If
    
    ary = Me.Find(rngTarget, keyword)
    Dim arytype As String
    Let arytype = C_Array.TypeArrayPattern(ary)
    
    If arytype = "ArrayNotInit" Then
        Call Err.Raise(Err.Number, Err.source, "keyword not found. -> """ & keyword & """")
    ElseIf arytype = "ArrayEmpty" Then
        Call Err.Raise(Err.Number, Err.source, "keyword not found. -> """ & keyword & """")
    ElseIf arytype = "Array" Then
        If C_Array.Count(ary) = 1 Then
            Set GetRangeByKeyword = Workbooks(bookName).Worksheets(sheetName).Range(ary(LBound(ary))).Offset(rowOffset, colOffset)
        Else
            Set GetRangeByKeyword = Workbooks(bookName).Worksheets(sheetName).Range(ary(UBound(ary))).Offset(rowOffset, colOffset)
        End If
    Else
         Call Err.Raise(Err.Number, Err.source, "not occurs. -> """ & keyword & """")
    End If
    GoSub FINALLY
    Exit Function
    '
FINALLY: 'Label
    Set rngTarget = Nothing
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Dim desc As String
    desc = CONS_MODULE_NAME & vbCr & Err.description
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Description = " & desc)
    Call Err.Raise(Err.Number, Err.source & "<-" & CONS_MODULE_NAME, desc)
    Exit Function
End Function

Public Sub PutArray( _
            ByVal rangeparam As String, _
            ByVal sheetName As String, _
            ByRef ary As Variant, _
            Optional style As Variant = "V", _
            Optional ByVal bookName As Variant)
''' /********************************************************
'''  *@function putArray
'''  *<br/>対象シートの対象範囲名をターゲットにして、配列データを書き込む
'''  *<br/>
'''  *@param {String} sheetname
'''  *@param {String} rangename
'''  *@param {Variant} ary
'''  *@param {Variant} Style    "V":Vertical, "H":Horizontal
'''  *@param {Variant} bookname Optional
'''  *
'''  *@require C_Array
'''  */
'
Dim C_Array As C_Array
Set C_Array = New C_Array
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
Dim rng As Range
Dim rngCurrent As Range
Set rngCurrent = Workbooks(bookName).Worksheets(sheetName).Range(rangeparam).CurrentRegion
Call rngCurrent.ClearContents
If UCase(style) = "V" Then
    Set rng = rngCurrent.Resize(UBound(ary) - LBound(ary) + 1, 1)
    rng.value = C_Array.ArrayToArrayArrayVertical(ary)
ElseIf UCase(style) = "H" Then
    Set rng = rngCurrent.Resize(1, UBound(ary) - LBound(ary) + 1)
    rng.value = C_Array.ArrayToArrayArray(ary)
End If
rng.WrapText = False
'
Set C_Array = Nothing
'
End Sub

Public Sub PutArrayArray( _
            ByVal rangeparam As String, _
            ByVal sheetName As String, _
            ByRef aryary As Variant, _
            Optional ByVal bookName As Variant)
''' /********************************************************
'''  *@function putArrayArray
'''  *<br/>ターゲット範囲に、二次元配列をプットする
'''  *<br/>
'''  *@param {String} rangeparam
'''  *@param {String} sheetname
'''  *@param {Variant} jagary
'''  *@param {Variant} bookname Optional
'''  *
'''  */
'
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
Dim rng As Range
Dim rngCurrent As Range
Set rngCurrent = Workbooks(bookName).Worksheets(sheetName).Range(rangeparam).CurrentRegion
Call rngCurrent.ClearContents
'
Dim clbMax As Long
Dim rlbMax As Long
Dim cubMax As Long
Dim rubMax As Long
rlbMax = LBound(aryary, 1)
clbMax = LBound(aryary, 2)
rubMax = UBound(aryary, 1) 'aryary's first dimension   ''FIXME
cubMax = UBound(aryary, 2) 'aryary's second dimension  ''FIXME
'
Set rng = rngCurrent.Resize(rubMax - rlbMax + 1, cubMax - clbMax + 1)
rng.value = aryary
rng.WrapText = False
'
End Sub

Public Sub PutJagArray( _
            ByVal rangename As String, _
            ByVal sheetName As String, _
            ByRef jagary As Variant, _
            Optional ByVal bookName As Variant)
''' /********************************************************
'''  *@function putJagArray
'''  *<br/>対象シートの対象範囲名をターゲットにして、ジャグ配列データを書き込む
'''  *<br/>
'''  *@param {String} rangeparam
'''  *@param {String} sheetname
'''  *@param {Variant} jagary
'''  *@param {Variant} bookname Optional
'''  *
'''  *@require C_Array
'''  */
'
Dim C_Array As C_Array
Set C_Array = New C_Array
If IsMissing(bookName) Then
    bookName = ThisWorkbook.name
End If
''
Dim rng As Range
Dim rngCurrent As Range
Set rngCurrent = Worksheets(sheetName).Range(rangename).CurrentRegion
Call rngCurrent.ClearContents
''
Dim cubMax As Long
Dim cub As Long
''
Dim r As Long
'
For r = LBound(jagary) To UBound(jagary)
    cub = UBound(jagary(r))
    If cub > cubMax Then
        cubMax = cub
    End If
Next r
'
Set rng = rngCurrent.Resize(UBound(jagary) - LBound(jagary) + 1, cubMax - LBound(jagary(LBound(jagary))) + 1)
rng.value = C_Array.JagArrayToArrayArray(jagary)
rng.WrapText = False
'
Set C_Array = Nothing
'
End Sub

Public Function PutJagArrayTranspose( _
            ByVal rangename As String, _
            ByVal sheetName As String, _
            ByRef jagary As Variant, _
            Optional ByVal bookName As Variant)
''' /********************************************************
'''  *@function putJagArray
'''  *<br/>対象シートの対象範囲名をターゲットにして、ジャグ配列データを書き込む
'''  *<br/>
'''  *@param {String} rangeparam
'''  *@param {String} sheetname
'''  *@param {Variant} jagary
'''  *@param {Variant} bookname Optional
'''  *
'''  *@require C_Array
'''  */
''
'' FIXME
''
End Function

Public Function GetRefertoString(ByRef rng As Range) As String
''' /********************************************************
'''  *@function getRefertoString
'''  *@param {Range} rng
'''  *
'''  */
''
Dim shtName As String
Dim addr As String
Let shtName = rng.Parent.name
Let addr = rng.AddressLocal
Console.log shtName
Console.log addr
GetRefertoString = "=" & shtName & "!" & addr
''
End Function

Public Function GetValueByAddress(ByVal addr As String, _
                Optional ByVal sheetName As Variant, _
                Optional ByRef wb As Variant) As Variant
''' /********************************************************
'''  *@function GetValueByAddress
'''  */
Dim C_String As C_String
Set C_String = New C_String
''
If IsMissing(wb) Then
  Set wb = ThisWorkbook
End If
Dim aryTemp As Variant
If IsMissing(sheetName) Then
    If C_String.Contain(addr, "!") Then
        Let aryTemp = Split(addr, "!")
        Let sheetName = C_String.Chomp(aryTemp(LBound(aryTemp)), 1)
    Else
        Let sheetName = ThisWorkbook.ActiveSheet.name
    End If
End If
Dim sht As Worksheet
Set sht = wb.Worksheets(sheetName)
GetValueByAddress = sht.Range(addr).value
''
End Function

Public Function ForEach(ByRef rng As Range _
                      , ByRef instance As Object _
                      , ByVal methodname As Variant _
                      , ParamArray pAry() As Variant)
''' /********************************************************
'''  *@function ForEach
'''  *<pre>
'''  *rngに対して、各々のCellに、処理を実行する。
'''  *</pre>
'''  *@param {Variant<Range>} rng
'''  *@param {Object} instance
'''  *@param {Variant} methodname
'''  *@param {Variant()} pAry ParamArray
'''  *※呼び出したmethodの戻り値を、対象Cellにセットする。
'''  *　戻り値は必ず設定する必要がある。
'''  *※pAryは、呼び出すmethodにはそのまま渡す。
'''  *　呼び出される側は、item,anAryを引数とする。（Variant型で引き渡す）
'''  *
'''  *　FIXME 使い方が難しすぎる。呼び出すMethodが、Rangeを意識して作る
'''  */
'
    Dim itm As Variant
    Dim sht As Worksheet
    For Each itm In rng
        Set sht = Worksheets(rng.Parent.name)
        sht.Cells(itm.Row, itm.column) = CallByName(instance, methodname, VbMethod, itm, pAry)
    Next itm
    Set sht = Nothing
End Function

Public Function ForEachAreas(ByRef rng As Variant _
                      , ByRef instance As Object _
                      , ByVal methodname As Variant _
                      , ParamArray pAry() As Variant)
''' /********************************************************
'''  *@function ForEachAreas
'''  *<pre>
'''  *複数のAreaを含むrngに対して、各々のCellに、処理を実行する。
'''  *</pre>
'''  *@param {Variant<Range>} rng
'''  *@param {Object} instance
'''  *@param {Variant} methodname
'''  *@param {Variant()} pAry ParamArray
'''  *※呼び出したmethodの戻り値を、対象Cellにセットする。
'''  *　戻り値は必ず設定する必要がある。
'''  *※pAryは、呼び出すmethodにはそのまま渡す。
'''  *　呼び出される側は、item,anAryを引数とする。（Variant型で引き渡す）
'''  */
'
    Dim itm As Variant
    Dim rngArea As Range
    Dim sht As Worksheet
    For Each rngArea In rng.Areas
        For Each itm In rngArea
            Set sht = Worksheets(rng.Parent.name)
            sht.Cells(itm.Row, itm.column) = CallByName(instance, methodname, VbMethod, itm, pAry)
        Next itm
    Next rngArea
    Set sht = Nothing
End Function

Public Sub Flush(ByRef rng As Range _
               , ByVal strFilename As String _
               , Optional ByVal strFoldername As Variant _
               , Optional delimiter As Variant = ",")
''' /********************************************************
'''  *@function flush
'''  *<pre>
'''  *rngをテキスト形式ファイル（CSV形式）に出力する
'''  *</pre>
'''  *@param {Range} rng
'''  *@param {String} strFilename
'''  *@param {String} strFoldername
'''  *@param {Variant} delimiter
'''  *
'''  *@require O_StringBuilder
'''  */
''
Dim objSB As O_StringBuilder
Set objSB = New O_StringBuilder
Dim C_File As C_File
Set C_File = New C_File
If IsMissing(strFoldername) Then
        strFoldername = C_File.GetSpecialFolder()
End If
Dim str As String
str = Me.ToCsv(rng, delimiter)
Call objSB.AppendLine(str)
Call objSB.Flush(strFilename, strFoldername)
Set objSB = Nothing
''
End Sub

Public Function ToCsv(ByRef rng As Range _
               , Optional delimiter As Variant = ",")
''' /********************************************************
'''  *@function ToCsv
'''  *<pre>
'''  *rngをテキスト形式ファイル（CSV形式）に出力する
'''  *</pre>
'''  *@param {Range} rng
'''  *@param {Variant} delimiter
'''  *   @note if using tab, set delimiter vbTab
'''  */
'
Dim str As String
Dim aryTemp As Variant
Dim ary
Dim rngRow As Range
Dim i As Long
'
For Each rngRow In rng.Rows
        aryTemp = Me.RangeToArray(rngRow)
        ReDim ary(UBound(aryTemp))
        For i = 0 To UBound(aryTemp)
            If (aryTemp(i) = "") Or IsEmpty(aryTemp(i)) Then
                ary(i) = ""
            Else
                ary(i) = C_Commons.ToString(aryTemp(i))
            End If
        Next i
        str = str & Join(ary, delimiter) & vbCrLf
Next rngRow
'
ToCsv = Left(str, Len(str) - 2)
'
End Function

Public Function ToXml(rng As Range _
                    , Optional aryHeader As Variant _
                    , Optional recname As Variant = "record" _
                    , Optional isBodyOnly As Variant = "False") As String
''' /********************************************************
'''  *@function ToXml
'''  *<pre>
'''  *rngに対して、XML形式に変換する
'''  *</pre>
'''  *@param {Range} rng
'''  *@param {Variant<String>()} aryHeader
'''  *@param {Variant<String>} recname
'''  *@param {Variant<Boolean>} isBodyOnly
'''  *
'''  *@require C_String,C_Array,C_Template, O_Template
'''  */
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Template As C_Template
Set C_Template = New C_Template
Dim objTemplate As O_Template
Set objTemplate = New O_Template
'
Dim C_String As C_String
Set C_String = New C_String
Dim template As String
Dim strTemp As String
Dim lngRowStart As Long
Dim ary As Variant
Dim dict As Dictionary
'
If IsMissing(aryHeader) Then
        aryHeader = Me.GetRowArray(rng, 1) '一行目をヘッダとする
        lngRowStart = 2
Else
        lngRowStart = 1
End If
'
'bodyの生成
template = C_Template.CreateXmlTemplate(recname, aryHeader)
objTemplate.template = template
'
Dim r As Long
'
For r = lngRowStart To rng.Rows.Count
        Set dict = C_Array.ArrayToDictionary(Me.GetRowArray(rng, r), aryHeader)
        Set objTemplate.dict = dict
        strTemp = strTemp & objTemplate.Merge() & vbCrLf
        Set dict = New Dictionary
Next r
'
strTemp = Left(strTemp, Len(strTemp) - 2)
'
'XMLの生成
Dim strHeader As String
If Not isBodyOnly Then
        strHeader = "<" & recname & "s>" & vbCrLf _
                & "${body}" & vbCrLf & "</" & recname & "s>" & vbCrLf
        strTemp = C_String.Indents(strTemp)
Else
        strHeader = "${body}"
End If
'
Call dict.Add("body", strTemp)
objTemplate.template = strHeader
Set objTemplate.dict = dict
ToXml = objTemplate.Merge()
'
Set C_Array = Nothing
Set C_Template = Nothing
Set objTemplate = Nothing
Set C_String = Nothing
Set dict = Nothing
''
End Function

Public Function ToXml2(rng As Range _
                    , Optional aryHeader As Variant _
                    , Optional recname As Variant = "record" _
                    , Optional isBodyOnly As Variant = "False") As String
''' /********************************************************
'''  *@function ToXml2
'''  *<pre>
'''  *rngに対して、XML形式に変換する
'''  *</pre>
'''  *@param {Range} rng
'''  *@param {Variant<String>()} aryHeader
'''  *@param {Variant<String>} recname
'''  *@param {Variant<Boolean>} isBodyOnly
'''  *
'''  *@require C_Template,C_String,C_Array,O_Template
'''  */
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Template As C_Template
Set C_Template = New C_Template
Dim objTemplate As O_Template
Set objTemplate = New O_Template
'
Dim C_String As C_String
Set C_String = New C_String
Dim template As String
Dim strTemp As String
Dim lngRowStart As Long
Dim ary As Variant
Dim dict As Dictionary
'
If IsMissing(aryHeader) Then
        aryHeader = Me.GetRowArray(rng, 1) '一行目をヘッダとする
        lngRowStart = 2
Else
        lngRowStart = 1
End If
'
'bodyの生成
template = C_Template.CreateXml2Template(recname, aryHeader)
objTemplate.template = template
'
Dim r As Long
'
For r = lngRowStart To rng.Rows.Count
        Set dict = C_Array.ArrayToDictionary(Me.GetRowArray(rng, r), aryHeader)
        Set objTemplate.dict = dict
        strTemp = strTemp & objTemplate.Merge() & vbCrLf
        Set dict = New Dictionary
Next r
'
strTemp = Left(strTemp, Len(strTemp) - 2)
'
'XMLの生成
Dim strHeader As String
If Not isBodyOnly Then
        strHeader = "<" & recname & "s>" & vbCrLf _
                & "${body}" & vbCrLf & "</" & recname & "s>" & vbCrLf
        strTemp = C_String.Indents(strTemp)
Else
        strHeader = "${body}"
End If
'
Call dict.Add("body", strTemp)
objTemplate.template = strHeader
Set objTemplate.dict = dict
ToXml2 = objTemplate.Merge()
'
Set C_Array = Nothing
Set C_Template = Nothing
Set objTemplate = Nothing
Set C_String = Nothing
Set dict = Nothing
''
End Function

Public Function ToJsonObject(ByRef rng As Range) As Variant
''' /********************************************************
'''  *@function ToVariable
'''  *@param {Range} rng
'''  *@return {Variant} ToVariable
'''  */
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_String As C_String
Set C_String = New C_String
'
Dim vntRtn As Variant
Dim ary As Variant
Dim aryHeader As Variant
Dim dict As Dictionary
Set dict = New Dictionary
Dim i As Long
'
Select Case Me.TypeRangePattern(rng)
        Case CONS_RANGE_TYPE(RANGE_TYPE.cell)
            If C_String.IsJsonObject(rng.value) Then
                Set vntRtn = C_Commons.ParseToJson(rng.value)
            Else
                Let vntRtn = C_Commons.ParseToJson(rng.value)
            End If
        Case CONS_RANGE_TYPE(RANGE_TYPE.ARRAY_HORIZONTAL)
            vntRtn = Me.RangeToArray(rng)
        Case CONS_RANGE_TYPE(RANGE_TYPE.ARRAY_VERTICAL)
            vntRtn = Me.RangeToArray(rng)
        Case CONS_RANGE_TYPE(RANGE_TYPE.HASHTABLE)         '''FIXME□□"#"でSkip
            aryHeader = Me.RangeToArray(rng.Columns(1))
            ary = Me.RangeToArray(rng.Columns(2))
            Set dict = C_Array.ArrayToDictionary(ary, aryHeader)
            Set vntRtn = dict
        Case CONS_RANGE_TYPE(RANGE_TYPE.RECORDROW)         '''FIXME□□"#"でSkip
            aryHeader = Me.RangeToArray(rng.Rows(1))
            ary = Me.RangeToArray(rng.Rows(2))
            Set dict = C_Array.ArrayToDictionary(ary, aryHeader)
            Set vntRtn = dict
        Case CONS_RANGE_TYPE(RANGE_TYPE.TABLEDATA)
            vntRtn = Me.RangeToJagArray(rng)
        Case CONS_RANGE_TYPE(RANGE_TYPE.CELLS_AREA)
            ary = Me.RangeToArrayForAreas(rng)
            For i = LBound(ary) To UBound(ary)
                If C_String.IsJsonObject(ary(i)) Then
                    Set ary(i) = C_Commons.ParseToJson(ary(i))
                Else
                    Let ary(i) = C_Commons.ParseToJson(ary(i))
                End If
            Next i
            vntRtn = ary
End Select
'
If IsObject(vntRtn) Then
        Set ToJsonObject = vntRtn
Else
        Let ToJsonObject = vntRtn
End If
'
Set C_Array = Nothing
Set C_String = Nothing
Set dict = Nothing
''
End Function

'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub unittest()
'''' *********************************************************
Dim CONS_MODULE_NAME As String
CONS_MODULE_NAME = "C_Range.unitTest"

On Error GoTo EXCEPTION
    ''''********************************************
    ''
    '' execute unittest
    ''
    'shtTEST.unittest
    ''
    ''''********************************************
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '
    Return

EXCEPTION: 'Label
    GoSub FINALLY
    Dim desc As String
    desc = CONS_MODULE_NAME & vbCr & Err.description
    Console.error ("Err.Number = " & Err.Number)
    Console.error ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.error ("Err.Description = " & desc)
    Call Err.Raise(Err.Number, Err.source & "<-" & CONS_MODULE_NAME, desc)
    Exit Sub

End Sub

