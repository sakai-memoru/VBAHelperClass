VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @class C_Environment.cls
'''  *<br/>Get WindowOS Environments
'''  *<br/>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  */

'//-----------------------------------------
'// Class Method

Private Declare Function getUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long

Public Function getEnvironment() As Dictionary
'''/****************************************
''' * @function getEnvironment
''' * @return {Dictionary}
''' */
    Dim dict As Dictionary
    Set dict = New Dictionary
    
    dict.Add "OS_VERSION", getOSVersion()
    dict.Add "EXCEL_VERSION", getExcelVersion()
    dict.Add "LOGIN_USER", getLoginUserName()
    dict.Add "COMPUTER_NAME", getConputerName()
    
    Set getEnvironment = dict
    Set dict = Nothing
End Function

Private Function getExcelVersion()
    Dim strVer As String
    Dim str As String
    
    strVer = Application.Version
    
    Select Case strVer
        Case "9.0"
            str = "Excel2000"
        Case "10.0"
            str = "Excel2002"
        Case "11.0"
            str = "Excel2003"
        Case "12.0"
            str = "Excel2007"
        Case "14.0"
            str = "Excel2010"
        Case "15.0"
            str = "Excel2013"
        Case "16.0"
            str = "Excel2016"
        Case Else
            str = "N/A"
    End Select
    getExcelVersion = str & "(" & strVer & ")"
End Function


'''
Private Function getOSVersion() As Variant
    Dim strVer As String
    Dim str As String
    
    strVer = Application.OperatingSystem
    
    If Left(strVer, 7) = "Windows" Then
        Select Case Right(strVer, 5)
            Case " 6.01"
                str = "Windows7"
            Case " 6.00"
                str = "WindowsVista"
            Case " 5.01"
                str = "WindowsXP"
            Case " 5.00"
                str = "Windows2000"
            Case "10.00"
                str = "Windows10"
            Case Else
                str = "N/A"
        End Select
    End If
    
    getOSVersion = str & " (" & strVer & ")"
End Function

Private Function getLoginUserName() As Variant
    Dim buf As String * 255
    getUserName buf, 255
    getLoginUserName = Left(buf, InStr(buf, vbNullChar) - 1)
End Function

Private Function getConputerName() As Variant
    Dim obj As Object
    Set obj = CreateObject("WScript.Shell")
    
    On Error Resume Next
    If Err.Number = 0 Then
        getConputerName = obj.ExpandEnvironmentStrings("%COMPUTERNAME%")
    Else
        Console.info "err: " & Err.Description
    End If

    Set obj = Nothing

End Function

Public Function getEnvironmentStrings(ByVal str As String) As Variant
    Dim obj As Object
    Set obj = CreateObject("WScript.Shell")
    
    On Error Resume Next
    If Err.Number = 0 Then
        getEnvironmentStrings = obj.ExpandEnvironmentStrings(str)
    Else
        Console.info "err: " & Err.Description
    End If

    Set obj = Nothing

End Function

Public Function getUserDomain() As Variant
    Dim obj As Object
    Set obj = CreateObject("WScript.Shell")
    
    On Error Resume Next
    If Err.Number = 0 Then
        getUserDomain = obj.ExpandEnvironmentStrings("%USERDOMAIN%")
    Else
        Console.info "err: " & Err.Description
    End If

    Set obj = Nothing

End Function

'''
Public Function getBookPath() As Variant
    getBookPath = ThisWorkbook.FullName
End Function

Public Function getBookName() As Variant
    getBookName = ThisWorkbook.name
End Function

Public Function getBookFolder() As Variant
    getBookFolder = ThisWorkbook.Path
End Function

'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub unitTest()
''' ********************************************************
Console.info ("// --------------------- start!")
Dim C_Env As C_Environment
Set C_Env = New C_Environment

'Call C_Env.getEnvironment
'''getEnvironment
Dim dict As Dictionary
Set dict = C_Env.getEnvironment()
Console.dump dict

'''getEnvironmentStrings
Dim str As String
str = C_Env.getEnvironmentStrings("%Path%")
Console.info str

'''getUserDomain
Console.info C_Env.getUserDomain

'''getBookPath
Console.info C_Env.getBookPath
'''getBookName
Console.info C_Env.getBookName
'''getBookFolder
Console.info C_Env.getBookFolder

Console.info ("// --------------------- end...")
End Sub


