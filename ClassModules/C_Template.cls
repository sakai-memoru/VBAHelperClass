VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Xml.cls
'''  *<pre> XML Format Utility Class
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  *
'''  */
'
''' /********************************************************
'''  * @class C_Xml
'''  * @classdesc XML Format Utility Class
'''  *
'''  */
'
Public Function CreateXmlTemplate(recname As Variant, aryHeader As Variant) As String
''' /********************************************************
'''  *@function createXmlTemplate
'''  *<pre>XML形式Formatを、Header Arrayをもとに生成する。
'''  *  @param {Variant<String>} recname record name
'''  *  @param {Variant<String>()} aryHeader
'''  *  @return {String} createXmlTemplate
'''  */
    Dim str As String
    str = "<" & recname & " "
    Dim i As Long
    For i = LBound(aryHeader) To UBound(aryHeader)
        str = str & aryHeader(i) & "=" & """" & "${" & aryHeader(i) & "}" & """" & " "
    Next i
    str = str & "/>"
    CreateXmlTemplate = str
End Function

Public Function CreateXml2Template(recname As Variant, aryHeader As Variant) As String
''' /********************************************************
'''  *@function createXml2Template
'''  *<pre>XML形式Formatを、Header Arrayをもとに生成する。
'''  *  @param {Variant<String>} recname record name
'''  *  @param {Variant<String>()} aryHeader
'''  *  @return {String} createXml2Template
'''  */
    Dim str As String
    str = "<" & recname & ">" & vbCrLf
    Dim i As Long
    For i = LBound(aryHeader) To UBound(aryHeader)
        str = str & vbTab & "<" & aryHeader(i) & ">${" & aryHeader(i) & "}</" & aryHeader(i) & ">" & vbCrLf
    Next i
    str = str & "</" & recname & ">"
    CreateXml2Template = str
End Function


'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Function unittest()
    Dim aryHeader As Variant

''createXmlTemplate
    aryHeader = Array("no", "name", "address", "postNo", "note", "age")
    Console.info Me.CreateXmlTemplate("user", aryHeader)
    
''createXml2Template
    aryHeader = Array("no", "name", "address", "postNo", "note", "age")
    Console.info Me.CreateXml2Template("user", aryHeader)
    
End Function


