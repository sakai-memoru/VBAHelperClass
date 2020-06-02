VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Generator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const CONS_DESCRIPTION_TAGNAME As String = "descriptionTag"

Public Function GetDescriptionTag(descriptionStr As String, id As String)
'''' ********************************************************
''
Dim C_String As C_String
Set C_String = New C_String
''
Dim strRtn As String
Dim template As String
Let template = "<small id='${id}' class='form-text text-muted'>${description}</small>"
''
''Console.info id
''Console.info descriptionStr
If descriptionStr = "" Then
    Let strRtn = ""
Else
    Let strRtn = template
    Let strRtn = C_String.Substitute4Template(strRtn, "description", descriptionStr)
    Let strRtn = C_String.Substitute4Template(strRtn, "id", id)
End If
''Console.info strRtn
Let GetDescriptionTag = strRtn
''
End Function

Public Function RenderFormTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
''
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.form.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
'' render
Set objTempl.dict = dict
Let RenderFormTag = objTempl.Merge
Call objTempl.Rest
''
End Function

Public Function RenderSelectTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
''
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.select.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
'' populate a dictionary
''' select tag
If dict.Exists("multiple") Then
    Call dict.Add("attributes", dict.item("multiple"))
End If
If Not dict.Exists("size") Then
    Call dict.Add("size", 1)
End If
''' add Description tag
Dim descriptionTag As String
If objTempl.ExistsVariable(CONS_DESCRIPTION_TAGNAME) Then
    Let descriptionTag = Me.GetDescriptionTag(dict.item("description"), dict.item("id"))
    Call dict.Add("descriptionTag", descriptionTag)
End If
'' render
Set objTempl.dict = dict
Let RenderSelectTag = objTempl.MergeWithBlock2
Call objTempl.Rest
''
End Function

Public Function RenderInputTextTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
'' @Description : Can use some types such as text, search, tel, url, email, password
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.input-text.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
'' populate a dictionary
''' input tag
If Not dict.Exists("maxlength") Then
    If dict.item("type") = "url" Then
        Call dict.Add("maxlength", "128")
    ElseIf dict.item("type") = "password" Then
        Call dict.Add("maxlength", "20")
    Else
        Call dict.Add("maxlength", "64")
    End If
End If
''' add Description tag
Dim descriptionTag As String
If objTempl.ExistsVariable(CONS_DESCRIPTION_TAGNAME) Then
    Let descriptionTag = Me.GetDescriptionTag(dict.item("description"), dict.item("id"))
    Call dict.Add("descriptionTag", descriptionTag)
End If
'' render
Set objTempl.dict = dict
Let RenderInputTextTag = objTempl.Merge
Call objTempl.Rest
''
End Function

Public Function RenderInputTextInlineTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
'' @Description : Can use some types such as text, search, tel, url, email, password
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.input-text-inline.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
'' populate a dictionary
''' input tag
If Not dict.Exists("maxlength") Then
    If dict.item("type") = "url" Then
        Call dict.Add("maxlength", "128")
    ElseIf dict.item("type") = "password" Then
        Call dict.Add("maxlength", "20")
    Else
        Call dict.Add("maxlength", "64")
    End If
End If
''' add Description tag
Dim descriptionTag As String
If objTempl.ExistsVariable(CONS_DESCRIPTION_TAGNAME) Then
    Let descriptionTag = Me.GetDescriptionTag(dict.item("description"), dict.item("id"))
    Call dict.Add("descriptionTag", descriptionTag)
End If
'' render
Set objTempl.dict = dict
Let RenderInputTextInlineTag = objTempl.Merge
Call objTempl.Rest
''
End Function

Public Function RenderInputTextInlinesTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
''
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.input-text-inlines.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
'' render
Set objTempl.dict = dict
Let RenderInputTextInlinesTag = objTempl.Merge
Call objTempl.Rest
''
End Function

Public Function RenderInputFileTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
'' @Description : Can use some types such as text, search, tel, url, email, password
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.input-file.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
''' add Description tag
Dim descriptionTag As String
If objTempl.ExistsVariable(CONS_DESCRIPTION_TAGNAME) Then
    Let descriptionTag = Me.GetDescriptionTag(dict.item("description"), dict.item("id"))
    Call dict.Add("descriptionTag", descriptionTag)
End If
'' render
Set objTempl.dict = dict
Let RenderInputFileTag = objTempl.Merge
Call objTempl.Rest
''
End Function

Public Function RenderInputDateTimeTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
'' @Description : Can use some types such as text, search, tel, url, email, password
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.input-datetime.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
''
If Not dict.Exists("step") Then
    If dict.item("type") = "datetim-local" Then
        Call dict.Add("step", "60")
    ElseIf dict.item("type") = "datetime" Then
        Call dict.Add("step", "60")
    ElseIf dict.item("type") = "time" Then
        Call dict.Add("step", "60")
    Else
        Call dict.Add("step", "60")
    End If
End If
''' add Description tag
Dim descriptionTag As String
If objTempl.ExistsVariable(CONS_DESCRIPTION_TAGNAME) Then
    Let descriptionTag = Me.GetDescriptionTag(dict.item("description"), dict.item("id"))
    Call dict.Add("descriptionTag", descriptionTag)
End If
'' render
Set objTempl.dict = dict
Let RenderInputDateTimeTag = objTempl.Merge
Call objTempl.Rest
''
End Function

Public Function RenderTextareaTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.textarea.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
''' add Description tag
Dim descriptionTag As String
If objTempl.ExistsVariable(CONS_DESCRIPTION_TAGNAME) Then
    Let descriptionTag = Me.GetDescriptionTag(dict.item("description"), dict.item("id"))
    Call dict.Add("descriptionTag", descriptionTag)
End If
'' populate a dictionary
''
'' render
Set objTempl.dict = dict
Let RenderTextareaTag = objTempl.Merge
Call objTempl.Rest
''
End Function

Public Function RenderRadioTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
''
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.input-radio.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
'' populate a dictionary
''' add Description tag
Dim descriptionTag As String
If objTempl.ExistsVariable(CONS_DESCRIPTION_TAGNAME) Then
    Let descriptionTag = Me.GetDescriptionTag(dict.item("description"), dict.item("id"))
    Call dict.Add("descriptionTag", descriptionTag)
End If
'' render
Set objTempl.dict = dict
Let RenderRadioTag = objTempl.MergeWithBlock2
Call objTempl.Rest
''
End Function

Public Function RenderCheckboxTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
''
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.input-checkbox.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
'' populate a dictionary
''' add Description tag
Dim descriptionTag As String
If objTempl.ExistsVariable(CONS_DESCRIPTION_TAGNAME) Then
    Let descriptionTag = Me.GetDescriptionTag(dict.item("description"), dict.item("id"))
    Call dict.Add("descriptionTag", descriptionTag)
End If
'' render
Set objTempl.dict = dict
Let RenderCheckboxTag = objTempl.MergeWithBlock2
Call objTempl.Rest
''
End Function

Public Function RenderCheckboxSingleTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
''
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.input-checkbox-single.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
'' populate a dictionary
If Not dict.Exists("value") Then
    Call dict.Add("value", "checked")
End If
''' add Description tag
Dim descriptionTag As String
If objTempl.ExistsVariable(CONS_DESCRIPTION_TAGNAME) Then
    Let descriptionTag = Me.GetDescriptionTag(dict.item("description"), dict.item("id"))
    Call dict.Add("descriptionTag", descriptionTag)
End If
'' render
Set objTempl.dict = dict
Let RenderCheckboxSingleTag = objTempl.Merge
Call objTempl.Rest
''
End Function

Public Function RenderButtonTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
''
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.button.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
'' populate a dictionary
If Not dict.Exists("class") Then
    Call dict.Add("class_exp", "btn-primary")
End If
If Not dict.Exists("button_type") Then
    Call dict.Add("button_type", "button")
End If
''' add Description tag
Dim descriptionTag As String
If objTempl.ExistsVariable(CONS_DESCRIPTION_TAGNAME) Then
    Let descriptionTag = Me.GetDescriptionTag(dict.item("description"), dict.item("id"))
    Call dict.Add("descriptionTag", descriptionTag)
End If
'' render
Set objTempl.dict = dict
Let RenderButtonTag = objTempl.Merge
Call objTempl.Rest
''
End Function

Public Function RenderButtonInlineTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
''
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.button-inline.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
'' populate a dictionary
If Not dict.Exists("class") Then
    Call dict.Add("class_exp", "btn-primary")
End If
If Not dict.Exists("button_type") Then
    Call dict.Add("button_type", "button")
End If
''' add Description tag
Dim descriptionTag As String
If objTempl.ExistsVariable(CONS_DESCRIPTION_TAGNAME) Then
    Let descriptionTag = Me.GetDescriptionTag(dict.item("description"), dict.item("id"))
    Call dict.Add("descriptionTag", descriptionTag)
End If
'' render
Set objTempl.dict = dict
Let RenderButtonInlineTag = objTempl.Merge
Call objTempl.Rest
''
End Function

Public Function RenderButtonInlinesTag(dict As Dictionary, Optional templFile As Variant, Optional templFolder As Variant)
'''' ********************************************************
''
Dim objTempl As O_Template
Set objTempl = New O_Template
Dim CONS_TEMPLATE As String
Let CONS_TEMPLATE = "template\templ.button-inlines.html"
''
Dim strRtn As String
Dim template As String
''
'' set template
If IsMissing(templFile) Then
    Let templFile = CONS_TEMPLATE
    Let template = objTempl.SetTemplateByFile(templFile)
Else
    If IsMissing(templFolder) Then
        Let template = objTempl.SetTemplateByFile(templFile)
    Else
        Let template = objTempl.SetTemplateByFile(templFile, templFolder)
    End If
End If
'' render
Set objTempl.dict = dict
Let RenderButtonInlinesTag = objTempl.Merge
Call objTempl.Rest
''
End Function

