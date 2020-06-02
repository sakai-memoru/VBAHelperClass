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
'''  *<pre>Dialog���ʃN���X
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  */
'
''' /********************************************************
'''  * @class C_Dialog
'''  * @classdesc Dialog���ʃN���X
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
'''  *<br/> WScript.Popup�̃��b�p�[
'''  *<br/> �g�����́AWScript.Shell��Popup���\�b�h���Q��
'''  *
'''  *�@@param {String} strText �o�͂������e�L�X�g
'''  *�@@param {Integer} nSecondsToWait �y�ȗ��\�zPopup Window�����܂ł̕b���B�K��l 0�i���[�U����ŕ���܂ł����ƕ\���j
'''  *�@@param {String} strTitle �y�ȗ��\�z�o�͂������^�C�g���B�K��l"WSH"
'''  *�@@param {Integer} nType �y�ȗ��\�z�{�^���ƃA�C�R���̎�ނ��w�肷�鐔�l�B�K��l=1([OK],[Cancel]�{�^���j
'''  *  @return {Integer} Popup 1:OK, 2:Cancel, 3:Exit, 4:Retry, 5:Ignore, 6:Yes, 7:No
'''  */
    '//--------------------------------------
    '//�ϐ��錾
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "Popup"
    Dim objWShell As Object
    Set objWShell = CreateObject("WScript.Shell")
    Dim intRet As Integer
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '//����
    intRet = objWShell.Popup(strText, nSecondsToWait, strTitle, nType)
    
    '//�J���ƕԋp
    GoSub FINALLY
    Popup = intRet
    Exit Function
    
FINALLY: 'Label
    '//--------------------------------------
    '// �I������
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
'''  *<br/> ExcelApplicatio��Application.GetOpenFileName�̃��b�p�[
'''  *<br/> �g�����́AApplication.GetOpenFileName���\�b�h���Q��
'''  *
'''  *�@@param {String} FileFilter �I���t�@�C���̎�ނ��w�肷�镶����
'''  *�@@param {Boolean} MultiSelect �y�ȗ��\�z�����I�����\�Ƃ��邩�ۂ��B�K��l False�i�P��̂݁j
'''  *  @return {Valiant} GetOpenFileName �I�����ʂɉ������l
'''  *           �I���F�I�������t�@�C���̃t���p�X������
'''  *           �����I���F�I�������t�@�C���̃t���p�X������̔z��i�P�����ł��z��`���ŕԋp�j
'''  *           �I�����Ȃ��FFalse
'''  */
    '//--------------------------------------
    '//�ϐ��錾
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "GetOpenFileName"
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    Dim vntPath As Variant
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '//����
    vntPath = objExcel.GetOpenFileName(FileFilter, MultiSelect)
    
    '//�J���ƕԋp
    GoSub FINALLY
    GetOpenFileName = vntPath
    Exit Function
    
FINALLY: 'Label
    '//--------------------------------------
    '// �I������
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
'''  *<br/> ExcelApplicatio��Application.GetOpenFileName�̃��b�p�[
'''  *<br/> �g�����́AApplication.GetOpenFileName���\�b�h���Q��
'''  *
'''  *�@@param {String} FileFilter �I���t�@�C���̎�ނ��w�肷�镶����
'''  *  @return {Valiant} GetOpenFileName �I�����ʂɉ������l
'''  *           �I���������͓��́F�I���E���͂����t�@�C���̃t���p�X������
'''  *           �I�����Ȃ��FFalse
'''  */
    '//--------------------------------------
    '//�ϐ��錾
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "GetSaveAsFilename"
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    Dim vntPath As Variant
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '//����
    vntPath = objExcel.GetSaveAsFilename(FileFilter)
    
    '//�J���ƕԋp
    GoSub FINALLY
    GetSaveAsFilename = vntPath
    Exit Function
    
FINALLY: 'Label
    '//--------------------------------------
    '// �I������
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
'''  *<br/> ExcelApplicatio��Application.FileDialog(msoFileDialogFolderPicker)�̃��b�p�[
'''  *<br/> �g�����́AApplication.FileDialog���\�b�h���Q��
'''  *
'''  *�@@param {String} Title Dialog�̃^�C�g��������
'''  *  @return {Valiant} GetOpenFileName �I�����ʂɉ������l
'''  *           �I���������͓��́F�I���E���͂����t�@�C���̃t���p�X������
'''  *           �I�����Ȃ��FFalse
'''  */
    '//--------------------------------------
    '//�ϐ��錾
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
    '//����
    objDialog.Title = Title
    vntRet = objDialog.Show()
    
    If vntRet = -1 Then
        vntPath = objDialog.SelectedItems(1)
    Else
        vntPath = False
    End If
    
    '//�ԋp
    GoSub FINALLY
    PickFolder = vntPath
    Exit Function

FINALLY: 'Label
    '//--------------------------------------
    '// �I������
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
'''  *<br/> Shell.Application��BrowserForFolder�̃��b�p�[
'''  *<br/> �g�����́AobjShell.BrowserForFolder���\�b�h���Q��
'''  *
'''  *�@@param {String} strTitle
'''  *�@@param {Integer} intOption
'''  *�@@param {Valiant} vntRootFolder
'''  *  @return {Valiant} BrowserForFolder �I�����ʂɉ������l
'''  *           �I���������͓��́F�I���E���͂����t�@�C���̃t���p�X������
'''  *           �I�����Ȃ��FFalse
'''  */
    '//--------------------------------------
    '//�ϐ��錾
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
    '//����
    If IsMissing(vntRootFolder) Then
        strDesktopPath = objWShell.SpecialFolders("Desktop")
        strTmpPath = strDesktopPath & "\"
    Else
        strTmpPath = vntRootFolder & "\"
    End If
    
    Set objPath = objShell.BrowseForFolder(&O0, strTitle, intOption, strTmpPath)
    
    If Not objPath Is Nothing Then
        Select Case objPath
            Case "�f�X�N�g�b�v"
                'objShell.BrowseForFolder��Desktop��I�������ꍇ�̏���
                vntPath = strDesktopPath
            Case Else
                vntPath = objPath.Items.item.path
        End Select
    End If
    
    '//�J���ƕԋp
    GoSub FINALLY
'    Console.log (vntPath)
    BrowseForFolder = vntPath
    Exit Function
    
FINALLY: 'Label
    '//--------------------------------------
    '// �I������
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
'  * �又���T���v��
'  *�@Entry Point
'  *
'  */
    '//--------------------------------------
    '//�ϐ��錾
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "C_Dialog.sample"
    Dim Dialog As New C_Dialog 'Dialog���ʃN���X�̐���
    Dim intRet As Integer
    Dim vnt As Variant
    
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// ��������
        
    '//--------------------------------------
    '// ����
    intRet = Dialog.Popup("���b�Z�[�W")
    Console.info ("intRet = " & intRet)
    'vnt = Dialog.GetOpenFileName("Excel�u�b�N(*.xlsx;*.xlsm),*.xlsx;*.xlsm")
    'console.info ("vnt = " & vnt)
    'vnt = Dialog.GetSaveAsFilename("Excel�u�b�N(*.xlsx;*.xlsm),*.xlsx;*.xlsm")
    'console.info ("vnt = " & vnt)
    'vnt = Dialog.PickFolder("�~�σt�H���_���w�肵�Ă�������")
    'Console.info ("vnt = " & vnt)
    'vnt = Dialog.BrowseForFolder("�~�σt�H���_���w�肵�Ă�������")
    Console.Dump (vnt)
    
    GoSub FINALLY
    Exit Sub
    
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
