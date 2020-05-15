VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Collection.cls
'''  *<pre>�R���N�V�����p���[�e�B���e�B�E�N���X
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  *
'''  */
'
''' /********************************************************
'''  * @class C_Collection
'''  * @classdesc �R���N�V�����p���[�e�B���e�B�E�N���X
'''  */
'
'//--------------------------------------------------------
'// Class Event
'//
Private Sub Class_Initialize()
''' /********************************************************
'''  * @event Class_Initialize
'''  *<br/>Initialize �֐�
'''  *<br/>Class�������ɁA���s�����B
'''  */
End Sub

'//--------------------------------------------------------
'// Class Method
'//
Public Function TypeNameEx(ByRef clts As Variant) As String
''' /********************************************************
'''  *@function TypeNameEx
'''  *<br/> Collection�I�u�W�F�N�g�̌^���m�F����B�i�o���Ƃ��āj
'''  *  @param {Variant} clts ByRef
'''  *  @return {String} TypeNameEx
'''  */
'
    Dim strType As String
    Dim strRtn As String
    strType = TypeName(clts)
    Select Case strType
        Case "Collection"
            strRtn = "Collection<Item>(idx_or_key)" & " <" & strType & ">"
        Case "Workbooks"
            strRtn = "Collection<Item>(idx_or_name).Name" & " <" & strType & ">"
        Case "Worksheets"
            strRtn = "Collection<Item>(idx_or_name).Name" & " <" & strType & ">"
        Case "Sheets"
            strRtn = "Collection<Item>(idx_or_name).Name" & " <" & strType & ">"
        Case "Shapes"
            strRtn = "Collection<Item>(idx_or_name).Name/TextFrame.Charactors.Text" & " <" & strType & ">"
        Case "ListObjects"
            strRtn = "Collection<Item>(idx_or_name).Name/Range" & " <" & strType & ">"
        Case "Names"
            strRtn = "Collection<Item>(idx_or_name).Name/Address/Value" & " <" & strType & ">"
        Case "Areas"
            strRtn = "Collection<Item>(idx).Range.Address" & " <" & strType & ">"
        Case "Files"
            strRtn = "Collection<Item>(idx).Name" & " <" & strType & ">"
        Case "Folders"
            strRtn = "Collection<Item>(idx).Name" & " <" & strType & ">"
        Case "SubFolders"
            strRtn = "Collection<Item>(idx).Name" & " <" & strType & ">"
        Case Else
            strRtn = "Unknown"
    End Select
    TypeNameEx = strRtn
End Function

Public Function VarTypeEx(ByRef clts As Variant) As String
''' /********************************************************
'''  *@function TypeNameEx
'''  *<br/> Collection�I�u�W�F�N�g�̌^���m�F����B�i�o���Ƃ��āj
'''  *  @param {Variant} clts ByRef
'''  *  @return {String} TypeNameEx
'''  */
'
    Dim strType As String
    Dim strRtn As String
    strType = TypeName(clts)
    Select Case strType
        Case "Collection"
            strRtn = "clts(idx_or_key)"
        Case "Workbooks"
            strRtn = "clts(idx_or_name).CodeName"
        Case "Worksheets"
            strRtn = "clts(idx_or_name).Index"
        Case "Sheets"
            strRtn = "clts(idx_or_name).Index"
        Case "Shapes"
            strRtn = "clts(idx_or_name)"
        Case "ListObjects"
            strRtn = "clts(idx_or_name).Range"
        Case "Names"
            strRtn = "clts(idx_or_name).Index"
        Case "Areas"
            strRtn = "clts(idx).Address"
        Case "Files"
            strRtn = "clts(idx).Name"
        Case "Folders"
            strRtn = "clts(idx).Name"
        Case "SubFolders"
            strRtn = "clts(idx).Name"
        Case Else
            strRtn = "Unknown"
    End Select
    VarTypeEx = strRtn
End Function

Public Function ToString(ByRef clts As Variant) As String
'''' /********************************************************
''''  *@function ToString
''''  *<pre> ���ł��镶����ɕϊ�����
''''  *</pre>
''''  *  @param {Variant} clts ByRef
''''  *  @return {String} ToString
''''  */
''
Dim strRtn As String
Dim strName As String
strName = TypeName(clts)
Select Case strName
        Case "Collection"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Workbooks"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowWorkbooks(clts)
        Case "Worksheets"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Sheets"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Shapes"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "ListObjects"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Names"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Areas"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Files"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Folders"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "SubFolders"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case Else
            strRtn = "<Unknown>" & C_Commons.ShowCollection(clts)
End Select
''
ToString = strRtn
''
End Function

Public Function Exists(ByRef clts As Variant, ByVal key As Variant) As Boolean
'''' /********************************************************
''''  *@function Exists
''''  *<pre> �R���N�V�����ɁA�w�肵��Key�����݂��邩�𔻒肷��B
''''  *</pre>
''''  *  @param {Variant} clts ByRef
''''  *  @param {Variant} key index or key
''''  *  @return {Boolean}
''''  *
''''  */
''
    Dim dammy As Variant
    Dim blnRtn As Boolean
''
On Error Resume Next
    Err.Clear
    ''
    dammy = clts(key)
    ''
    If Err.Number = 0 Then
        blnRtn = True
    Else
        Err.Clear
        dammy = IsObject(clts.item(key)) '''item(key)���擾�ł��Ȃ��Ǝ��s�G���[����
            If Err.Number = 0 Then
                blnRtn = True
            Else
                blnRtn = False
                If Err.Number = 438 Then
                    Err.Description = Err.Description & vbCrLf & "���̃R���N�V������Item�v���p�e�B�����݂��Ȃ����߁A���ʂł��܂���"
                    Console.log Err.Description
                End If
            End If
    End If
    Exists = blnRtn
End Function

Public Sub PutValue(ByRef clts As Collection, ByVal value As Variant, ByVal key As Variant)
''' /********************************************************
'''  *@function
'''  *<pre> �R���N�V�����ɁA�w�肵��Key-Value��put����B
'''  *</pre>
'''  *  @param {Variant} clts ByRef
'''  *  @param {Variant} value
'''  *  @param {Variant} key index or key
'''  *  @note Collection�ϐ���New �ō쐬�����ꍇ�ɁA���Yfunction�̗��p��z��
'''  */
'
    If Me.Exists(clts, key) Then
        Call clts.Remove(key)
        Call clts.Add(value, key)
    Else
        Call clts.Add(value, key)
    End If
End Sub

Public Function ExistsNameOfObject(ByRef clts As Variant, ByVal name As String) As Boolean
'''' /********************************************************
''''  *@function ExistsName
''''  *<br/> Object�̃R���N�V�����ɁA�w��̖��O�����݂��邩�𔻒肷��B
''''  *  @param {Variant} clts ByRef
''''  *  @param {String} Name
''''  *  @return {Boolean}
''''  *
''''  */
''
Dim itm As Object
Dim blnRtn As Boolean
blnRtn = False
For Each itm In clts
        If itm.name = name Then
            blnRtn = True
            Exit For
        End If
Next itm
ExistsNameOfObject = blnRtn
End Function

Public Function FilterNameMatchOfObject(ByRef clts As Variant, ByVal str As String) As Collection
'''' /********************************************************
''''  *@function FilterNameMatch
''''  *<pre> �R���N�V�����̖��O���ALike�ōi�荞��ŁA�i�荞�񂾃R���N�V������ԋp
''''  *</pre>
''''  *@param {Collection} clts ByRef
''''  *@param {String} str
''''  *@return {Variant}  FilterNameMatch
''''  *
''''  */
''
Dim cltRtns As Collection
Set cltRtns = New Collection
Dim cl As Variant
For Each cl In clts
        If LCase(cl.name) Like "*" & LCase(str) & "*" Then
            cltRtns.Add cl
        End If
Next cl
''
Set FilterNameMatchOfObject = cltRtns
Set cltRtns = Nothing
End Function

Public Function FilterNameMatchPrefixOfObject(ByRef clts As Variant, ByVal prefix As String) As Collection
'''' /********************************************************
''''  *@function FilterNameMatchPrefix
''''  *<br/> �R���N�V�����̖��O���A�O����v�ōi�荞��ŁA�i�荞�񂾃R���N�V������ԋp
''''  *  @param {Collection} clts ByRef
''''  *  @param {String} name
''''  *  @return {Variant}  FilterNameMatchPrefix
''''  *
''''  */
''
Dim cltRtns As Collection
Set cltRtns = New Collection
Dim cl As Object
For Each cl In clts
        If cl.name Like prefix & "*" Then
            cltRtns.Add cl
        End If
Next cl
Set FilterNameMatchPrefixOfObject = cltRtns
Set cltRtns = Nothing
End Function

Public Function ContainsNameOfObject(ByRef clts As Variant, ByVal str As String) As Boolean
'''' /********************************************************
''''  *@function ExistsName
''''  *<br/> �R���N�V�����ɁA�w��̕�����Like�ő��݂��邩�𔻒肷��B
''''  *  @param {Variant} clts ByRef
''''  *  @param {String} str
''''  *  @return {Boolean}
''''  *
''''  */
''
Dim itm As Object
Dim blnRtn As Boolean
blnRtn = False
''
Dim cltsTemp As Collection
''
Set cltsTemp = Me.FilterNameMatchOfObject(clts, str)
''
If cltsTemp.Count > 0 Then
        blnRtn = True
Else
        blnRtn = False
End If
''
ContainsNameOfObject = blnRtn
Set cltsTemp = Nothing
End Function

Public Function ContainsNameMatchPrefixOfObject(ByRef clts As Variant, ByVal str As String) As Boolean
'''' /********************************************************
''''  *@function ContainsNameMatchPrefix
''''  *<br/> �R���N�V�����ɁA�w��̕�����Like�ő��݂��邩�𔻒肷��B
''''  *  @param {Variant} clts ByRef
''''  *  @param {String} str
''''  *  @return {Boolean}
''''  *
''''  */
''
Dim itm As Object
Dim blnRtn As Boolean
blnRtn = False
''
Dim cltsTemp As Collection
''
Set cltsTemp = Me.FilterNameMatchPrefixOfObject(clts, str)
''
If cltsTemp.Count > 0 Then
        blnRtn = True
Else
        blnRtn = False
End If
''
ContainsNameMatchPrefixOfObject = blnRtn
Set cltsTemp = Nothing
End Function

Public Function ToArray(ByRef clts As Collection) As Variant
'''' /********************************************************
''''  *@function ToArray
''''  *<br/> �R���N�V�����̃A�C�e����z��ɂ��ĕԋp�ikey�͖����j
''''  *  @param {Collection} clts ByRef
''''  *  @return {Variant} ToArray
''''  */
''
Dim aryRtn() As Variant
Dim cnt As Long
Dim i As Long
''
cnt = clts.Count
If cnt < 1 Then
        ToArray = Array()
        Exit Function
End If
ReDim aryRtn(1 To cnt)
''
For i = 1 To cnt
        If IsObject(clts.item(i)) Then
        Set aryRtn(i) = clts.item(i)
        Else
        aryRtn(i) = clts.item(i)
        End If
Next i
''
ToArray = aryRtn
''
End Function
    
Public Function ToArrayList(ByRef clts As Collection) As Object
'''' /********************************************************
''''  *@function ToArray
''''  *<br/> �R���N�V�����̃A�C�e����z��ɂ��ĕԋp�ikey�͖����j
''''  *  @param {Collection} clts ByRef
''''  *  @return {Variant} ToArray
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim aryList As Object
Dim ary() As Variant
Let ary = Me.ToArray(clts)
''
Set aryList = C_Array.ToArrayList(ary)
''
Set ToArrayList = aryList
''
End Function

Public Function ToArrayNameOfObject(ByRef clts As Variant) As Variant
'''' /********************************************************
''''  *@function ToArrayName
''''  *<pre> �R���N�V�����̖��O��z��ɂ��ĕԋp�B
''''  *���R���N�V��������Object�ɁAName�v���p�e�B��������̂�O��Ƃ���B
''''  *</pre>
''''  *  @param {Collection} clts ByRef
''''  *  @return {Variant}  ToArrayName
''''  *
''''  */
''
Dim aryRtn As Variant
Dim cnt As Long
Dim i As Long
Dim cl As Variant
''
cnt = clts.Count
If cnt > 0 Then
        ReDim aryRtn(cnt - 1)
        i = 0
        For Each cl In clts
            aryRtn(i) = cl.name
            i = i + 1
        Next cl
    ToArrayNameOfObject = aryRtn
Else
    ToArrayNameOfObject = ""
End If
End Function

Public Function ToDictionary(ByVal clts As Variant, ByRef aryHeader As Variant) As Dictionary
'''' /********************************************************
''''  *@function ToDictionary
''''  *<pre> �R���N�V�������A�w��̃w�b�_�z������ƂɘA�z�z��ɂ��ĕԋp�B
''''  *</pre>
''''  *  @param {Variant} ary ByRef
''''  *  @return {Dictionary}  ToDictionary
''''  *  @note Collection�ϐ���New �ō쐬�����ꍇ�ɁA���Yfunction�̗��p��z��
''''  */
''
If Not IsArray(aryHeader) Then Err.Raise 13
If IsObject(clts) Then
        If Not TypeName(clts) = "Collection" Then
            Err.Raise 13
        End If
End If
Dim dict As Dictionary
Set dict = New Dictionary
Dim i As Long
For i = LBound(aryHeader) To UBound(aryHeader)
        If Me.Exists(clts, aryHeader(i)) Then
            Call dict.Add(aryHeader(i), clts(aryHeader(i)))
        Else
            Call dict.Add(aryHeader(i), "")
        End If
Next i
Set ToDictionary = dict
'Set dict = Nothing
End Function

Public Function ArrayToCollection(ByRef ary As Variant) As Collection
'''' /********************************************************
''''  *@function ArrayToCollection
''''  *<pre> �z����R���N�V�����ɂ��ĕԋp�B
''''  *</pre>
''''  *  @param {Variant} ary ByRef
''''  *  @return {Collection}
''''  *
''''  */
''
If Not IsArray(ary) Then Err.Raise 13
Set ArrayToCollection = New Collection
Dim vnt As Variant
For Each vnt In ary
        ArrayToCollection.Add vnt
Next vnt
''
End Function

Public Function ArrayListToCollection(ByRef aryList As Object) As Collection
'''' /********************************************************
''''  *@function ArrayListToCollection
''''  *<pre> ArrayList�z����R���N�V�����ɂ��ĕԋp�B
''''  *</pre>
''''  *  @param {Object} arylist ByRef
''''  *  @return {Collection}
''''  *
''''  */
''
If Not IsObject(aryList) Then Err.Raise 13
Set ArrayListToCollection = New Collection
Dim vnt As Variant
Dim i As Integer
For i = 0 To aryList.Count - 1
    ArrayListToCollection.Add aryList.item(i)
Next i
''
End Function

Public Sub Push(ByRef clts As Collection, ByVal value As Variant)
'''' /********************************************************
''''  *@function Push
''''  *<pre> �l��push
''''  *</pre>
''''  *  @param {Collection} clts ByRef
''''  *  @param {Variant} value ByRef
''''  */
''
clts.Add value
''
End Sub

Public Function Pop(ByVal clts As Collection) As Variant
'''' /********************************************************
''''  *@function Pop
''''  *<pre> �l��Pop
''''  *</pre>
''''  *  @param {Collection} clts ByRef
''''  *  @return {Variant} Pop
''''  */
''
Dim i As Long
i = clts.Count
If IsObject(clts.item(i)) Then
        Set Pop = clts.item(i)
Else
        Let Pop = clts.item(i)
End If
clts.Remove i
''
End Function

Public Sub Unshift(ByVal clts As Collection, ByVal value As Variant, Optional ByVal pos As Long = 1)
'''' /********************************************************
''''  *@function Shift
''''  *<pre> �l��Shift
''''  *</pre>
''''  *  @param {Collection} clts ByRef
''''  *  @param {Variant} value ByRef
''''  */
''
If clts.Count < 1 Then
        clts.Add value
Else
    clts.Add value, , pos
End If
''
End Sub

Public Function Shift(ByVal clts As Collection) As Variant
'''' /********************************************************
''''  *@function Unshift
''''  *<pre> �l��Unshift
''''  *</pre>
''''  *  @param {Collection} clts ByRef
''''  *  @return {Variant} Pop
''''  */
''
Dim i As Long
i = 1
If IsObject(clts.item(i)) Then
    Set Shift = clts.item(i)
Else
    Let Shift = clts.item(i)
End If
clts.Remove i
''
End Function

Public Function Rearrange(ByRef clt As Collection, ByVal aryHeader As Variant, Optional ByRef dict As Variant) As Collection
'''' /********************************************************
''''  *@function Rearrange
''''  *</br> �R���N�V��������ѕς���
''''  *  @param {Collection} clt ByRef
''''  *  @param {String()} aryHeader
''''  *  @param {Dictionary} dict Optional
''''  *  @return {Collection}
''''  */
''
Dim cltOut As Collection
Set cltOut = New Collection
Dim c As Long
If VBA.IsMissing(dict) Then
    Set dict = New Dictionary
    For c = LBound(aryHeader) To UBound(aryHeader)
        dict.Add aryHeader(c), aryHeader(c)
    Next c
End If
''
For c = LBound(aryHeader) To UBound(aryHeader)
    cltOut.Add clt.item(aryHeader(c)), dict.item(aryHeader(c))
Next c
''
Set Rearrange = cltOut
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
'//--------------------------------------
'//variables
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_Collection.unittest"
''
On Error GoTo EXCEPTION
    'Call debugC_Collection.unittest1
    'Call debugC_Collection.unittest2
    'Call debugC_Collection.unittest3
    ''
    GoSub FINALLY
    Exit Sub
    ''
FINALLY: 'Label
    '//--------------------------------------
    '// �I������
    Return

EXCEPTION: 'Label
    '//--------------------------------------
    ' ��O����
    GoSub FINALLY
    Console.info ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
End Sub