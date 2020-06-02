VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /********************************************************
'''  * @file C_Dialog
'''  *<pre>Dialog共通クラス
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  */
'
''' /********************************************************
'''  * @class C_Dialog
'''  * @classdesc Dialog共通クラス
'''  */
'
' ---------------------------------------------------------
Public Function Popup(strText As String _
            , Optional nSecondsToWait As Integer = 0 _
            , Optional strTitle As String = "WSH" _
            , Optional nType As Integer = 1 _
            ) As Integer
''' /********************************************************
'''  * @function Popup
'''  *<br/> WScript.Popupのラッパー
'''  *<br/> 使い方は、WScript.ShellのPopupメソッドを参照
'''  *
'''  *　@param {String} strText 出力したいテキスト
'''  *　@param {Integer} nSecondsToWait 【省略可能】Popup Windowを閉じるまでの秒数。規定値 0（ユーザ操作で閉じるまでずっと表示）
'''  *　@param {String} strTitle 【省略可能】出力したいタイトル。規定値"WSH"
'''  *　@param {Integer} nType 【省略可能】ボタンとアイコンの種類を指定する数値。規定値=1([OK],[Cancel]ボタン）
'''  *  @return {Integer} Popup 1:OK, 2:Cancel, 3:Exit, 4:Retry, 5:Ignore, 6:Yes, 7:No
'''  */
    '//--------------------------------------
    '//変数宣言
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "Popup"
    Dim objWShell As Object
    Set objWShell = CreateObject("WScript.Shell")
    Dim intRet As Integer
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '//処理
    intRet = objWShell.Popup(strText, nSecondsToWait, strTitle, nType)
    
    '//開放と返却
    GoSub FINALLY
    Popup = intRet
    Exit Function
    
FINALLY: 'Label
    '//--------------------------------------
    '// 終了処理
    Set objWShell = Nothing
    Return

EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function

Public Function GetOpenFileName(FileFilter As String, Optional MultiSelect As Boolean = False) As Variant
''' /********************************************************
'''  * @function GetOpenFileName
'''  *<br/> ExcelApplicatioのApplication.GetOpenFileNameのラッパー
'''  *<br/> 使い方は、Application.GetOpenFileNameメソッドを参照
'''  *
'''  *　@param {String} FileFilter 選択ファイルの種類を指定する文字列
'''  *　@param {Boolean} MultiSelect 【省略可能】複数選択を可能とするか否か。規定値 False（単一のみ）
'''  *  @return {Valiant} GetOpenFileName 選択結果に応じた値
'''  *           選択：選択したファイルのフルパス文字列
'''  *           複数選択：選択したファイルのフルパス文字列の配列（１つだけでも配列形式で返却）
'''  *           選択しない：False
'''  */
    '//--------------------------------------
    '//変数宣言
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "GetOpenFileName"
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    Dim vntPath As Variant
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '//処理
    vntPath = objExcel.GetOpenFileName(FileFilter, MultiSelect)
    
    '//開放と返却
    GoSub FINALLY
    GetOpenFileName = vntPath
    Exit Function
    
FINALLY: 'Label
    '//--------------------------------------
    '// 終了処理
    Set objExcel = Nothing
    Return


EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function

Public Function GetSaveAsFilename(FileFilter As String) As Variant
''' /********************************************************
'''  * @function GetSaveAsFilename
'''  *<br/> ExcelApplicatioのApplication.GetOpenFileNameのラッパー
'''  *<br/> 使い方は、Application.GetOpenFileNameメソッドを参照
'''  *
'''  *　@param {String} FileFilter 選択ファイルの種類を指定する文字列
'''  *  @return {Valiant} GetOpenFileName 選択結果に応じた値
'''  *           選択もしくは入力：選択・入力したファイルのフルパス文字列
'''  *           選択しない：False
'''  */
    '//--------------------------------------
    '//変数宣言
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "GetSaveAsFilename"
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    Dim vntPath As Variant
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '//処理
    vntPath = objExcel.GetSaveAsFilename(FileFilter)
    
    '//開放と返却
    GoSub FINALLY
    GetSaveAsFilename = vntPath
    Exit Function
    
FINALLY: 'Label
    '//--------------------------------------
    '// 終了処理
    Set objExcel = Nothing

EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
        
End Function

Public Function PickFolder(Title As String) As Variant
''' /********************************************************
'''  * @function pickFolder
'''  *<br/> ExcelApplicatioのApplication.FileDialog(msoFileDialogFolderPicker)のラッパー
'''  *<br/> 使い方は、Application.FileDialogメソッドを参照
'''  *
'''  *　@param {String} Title Dialogのタイトル文字列
'''  *  @return {Valiant} GetOpenFileName 選択結果に応じた値
'''  *           選択もしくは入力：選択・入力したファイルのフルパス文字列
'''  *           選択しない：False
'''  */
    '//--------------------------------------
    '//変数宣言
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "PickFolder"
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    Dim objDialog As Object
    Set objDialog = objExcel.FileDialog(msoFileDialogFolderPicker)
    Dim vntPath As Variant
    Dim vntRet  As Variant
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '//処理
    objDialog.Title = Title
    vntRet = objDialog.Show()
    
    If vntRet = -1 Then
        vntPath = objDialog.SelectedItems(1)
    Else
        vntPath = False
    End If
    
    '//返却
    GoSub FINALLY
    PickFolder = vntPath
    Exit Function

FINALLY: 'Label
    '//--------------------------------------
    '// 終了処理
    Set objExcel = Nothing
    Set objDialog = Nothing
    Return

EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)

End Function

Public Function BrowseForFolder(strTitle As String, Optional intOption As Integer = &H1 + &H10, Optional vntRootFolder As Variant) As Variant
''' /********************************************************
'''  * @function BrowserForFolder
'''  *<br/> Shell.ApplicationのBrowserForFolderのラッパー
'''  *<br/> 使い方は、objShell.BrowserForFolderメソッドを参照
'''  *
'''  *　@param {String} strTitle
'''  *　@param {Integer} intOption
'''  *　@param {Valiant} vntRootFolder
'''  *  @return {Valiant} BrowserForFolder 選択結果に応じた値
'''  *           選択もしくは入力：選択・入力したファイルのフルパス文字列
'''  *           選択しない：False
'''  */
    '//--------------------------------------
    '//変数宣言
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "BrowseForFolder"
    
    Dim objShell As Object
    Set objShell = CreateObject("Shell.Application")
    Dim objPath As Object
    Dim objWShell As Object
    Set objWShell = CreateObject("WScript.Shell")
    Dim vntPath As Variant
    Dim strDesktopPath As String
    Dim strTmpPath As String
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '//処理
    If IsMissing(vntRootFolder) Then
        strDesktopPath = objWShell.SpecialFolders("Desktop")
        strTmpPath = strDesktopPath & "\"
    Else
        strTmpPath = vntRootFolder & "\"
    End If
    
    Set objPath = objShell.BrowseForFolder(&O0, strTitle, intOption, strTmpPath)
    
    If Not objPath Is Nothing Then
        Select Case objPath
            Case "デスクトップ"
                'objShell.BrowseForFolderでDesktopを選択した場合の処理
                vntPath = strDesktopPath
            Case Else
                vntPath = objPath.Items.item.path
        End Select
    End If
    
    '//開放と返却
    GoSub FINALLY
'    Console.log (vntPath)
    BrowseForFolder = vntPath
    Exit Function
    
FINALLY: 'Label
    '//--------------------------------------
    '// 終了処理
    Set objShell = Nothing
    Set objPath = Nothing
    Set objWShell = Nothing
    Return

EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function


'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub unittest()
' /************************************
'  * 主処理サンプル
'  *　Entry Point
'  *
'  */
    '//--------------------------------------
    '//変数宣言
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "C_Dialog.sample"
    Dim Dialog As New C_Dialog 'Dialog共通クラスの生成
    Dim intRet As Integer
    Dim vnt As Variant
    
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// 初期処理
        
    '//--------------------------------------
    '// 処理
    intRet = Dialog.Popup("メッセージ")
    Console.info ("intRet = " & intRet)
    'vnt = Dialog.GetOpenFileName("Excelブック(*.xlsx;*.xlsm),*.xlsx;*.xlsm")
    'console.info ("vnt = " & vnt)
    'vnt = Dialog.GetSaveAsFilename("Excelブック(*.xlsx;*.xlsm),*.xlsx;*.xlsm")
    'console.info ("vnt = " & vnt)
    'vnt = Dialog.PickFolder("蓄積フォルダを指定してください")
    'Console.info ("vnt = " & vnt)
    'vnt = Dialog.BrowseForFolder("蓄積フォルダを指定してください")
    Console.Dump (vnt)
    
    GoSub FINALLY
    Exit Sub
    
FINALLY: 'Label
    '//--------------------------------------
    '// 終了処理
    Return

EXCEPTION: 'Label
    '//--------------------------------------
    ' 例外処理
    GoSub FINALLY
    Console.info ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    
End Sub

