VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_FileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_FileIO.cls
'''  *<br/>FileIO共通クラス
'''  * @version 1.00
'''  * @since 2020/03/31
'''  */
'
''' /********************************************************
'''  * @class FileIO
'''  * @classdesc FileIO共通クラス
'''  */
'
'//-----------------------------------------
'// Instance Property　Constant
'    Const CONS_STREAM_CHARSET_UTF As String = "UTF-8"
''' /** @name CONS_STREAM_TYPE_TEXT @constant @type {long} @default @desc {1 : Binary, 2 : Text} */
Const CONS_STREAM_TYPE_TEXT As Long = 2 ' {1 : Binary, 2 : Text}
''' /** @name CONS_STREAM_MODE_TYPE @constant @type {long} @default @desc {1 : New File, 2 : Update} */
Const CONS_STREAM_MODE_TYPE As Long = 1 ' {1 : New File, 2 : Update}

'//-----------------------------------------
'// Instance Property　Variable
''' /** @name objStream @property @type {Object} @desc ADODB.Stream用Object */
Dim objStream As Object

Private Sub Class_Initialize()
''' /********************************************************
'''  * @event Class_Initialize
'''  * @fire Class 生成時
'''  *<br/>Initialize 関数
'''  *<br/>Class生成時に、実行される。
'''  */
    Const CONS_MODULE_NAME As String = "C_FileIO.Class_Initialize"
On Error GoTo EXCEPTION
    ' 出力ストリームを取得
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = CONS_STREAM_TYPE_TEXT
    objStream.Charset = "UTF-8"
    'objStream.Mode = CONS_STREAM_MODE_TYPE
    'objStream.Open
    If Err.Number <> 0 Then
        GoSub EXCEPTION
    End If
    
    Exit Sub
EXCEPTION:
    'FIXME□□ErrのCatchは、Classではなく、呼出し元（Main）にて実行されるのか？
    '　　このClassの実行エラーは、このExceptionルーチンを通らない。
    Call Err.Raise(Err.Number, Err.source & CONS_MODULE_NAME, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
End Sub


Private Sub Class_Terminate()
''' /********************************************************
'''  * @event Class_Terminate
'''  * @fire Class 廃棄時
'''  *<br/>Terminate 関数
'''  *<br/>Class生成時に、実行される。
'''  */
    ' 出力ストリームをClose
    'objStream.Close
    ' Object開放
    Set objStream = Nothing
End Sub


'//-----------------------------------------
'// Instance Method
Public Function WriteTextAllAsUTF8(Optional ByVal filepath As String, Optional ByVal src As String) As Long
''' /********************************************************
'''  *@function writeTextAllAsUTF8
'''  *</br>UTF-8でテキストを出力する
'''  * @param {string} filepath　【Optional】出力するファイルパス名
'''  * @param {string} src 　【Optional】出力内容
'''  * @param {int} 実行結果コード
'''  */
    Const CONS_MODULE_NAME As String = "C_FileIO.writeTextAllAsUTF8"
On Error GoTo EXCEPTION
    '------------------------------
    ' 主処理
    '------------------------------
    objStream.Open
    objStream.WriteText src
    objStream.SaveToFile filepath, 2
    '返却
    WriteTextAllAsUTF8 = 0
    objStream.Close
    Exit Function

EXCEPTION:
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)

End Function

Public Function ReadTextAllAsUTF8(strFilePath As String) As String
''' /****************************************
'''  *@function readTextAllAsUTF8
'''  *</br>UTF-8でテキストを入力する
'''  * @param  {String} strFilePath ファイルパス名
'''  * @return {String} 読み取ったテキスト（全部）
'''  */
    Const CONS_MODULE_NAME As String = "C_FileIO.readTextAllAsUTF8"
On Error GoTo EXCEPTION
    '------------------------------
    ' 主処理
    '------------------------------
    Dim str As String
    objStream.Open
    objStream.LoadFromFile (strFilePath)
    str = objStream.readText()
    ReadTextAllAsUTF8 = str
    objStream.Close
    Exit Function

EXCEPTION:
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)

End Function

Public Function ReadTextAllAsUTF8ToArray(strFilePath As String) As Variant
''' /****************************************
'''  *@function readTextAllAsUTF8ToArray
'''  *</br>UTF-8でテキストを入力する
'''  * @param  {String} strFilePath ファイルパス名
'''  * @return {Variant} 読み取ったテキスト（全部）
'''  */
    Const CONS_MODULE_NAME As String = "C_FileIO.readTextAllAsUTF8ToArray"
On Error GoTo EXCEPTION
    '------------------------------
    ' 主処理
    '------------------------------
    Dim str As String
    Dim aryRtn As Variant
    'objStream.Open
    objStream.LoadFromFile (strFilePath)
    str = objStream.readText()
    
    aryRtn = Split(str, vbCrLf)
    ReadTextAllAsUTF8ToArray = aryRtn
    objStream.Close
    Exit Function

EXCEPTION:
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)

End Function
