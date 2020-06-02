VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Clipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Clipboard.cls
'''  *<pre>Clipboard utility
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/04/30
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  */
'
''' @Require refer to Microsoft Form DataObject

Public Function GetClipboard()
'''' ***************************************************
Dim objData As New MSForms.DataObject
Dim strText As String
objData.GetFromClipboard
strText = objData.GetText()
''
Let GetClipboard = strText
''
Set objData = Nothing
End Function

Public Function SetClipboard(ByVal str As String)
'''' ***************************************************
Dim objData As New MSForms.DataObject
''
objData.SetText str
objData.PutInClipboard
''
Set objData = Nothing
''
End Function

