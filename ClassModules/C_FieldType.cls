VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_FieldType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum enumShapeType
'''' ********************************************************
''
Group = msoGroup
TextBox = msoTextBox
AutoShape = msoAutoShape
FormControl = msoFormControl
OLEObject = msoOLEControlObject
Table = msoTable
Chart = msoChart
Comment = msoComment
Picture = msoPicture
PlaceHolder = msoPlaceholder
line = msoLine
''
End Enum

Public Enum enumAutoShapeType
'''' ********************************************************
''
Mixed = msoShapeMixed
Rectangle = msoShapeRectangle
Oval = msoShapeOval
FoldedCorner = msoShapeFoldedCorner
RectangularCallout = msoShapeRectangularCallout
RoundedRectangularCallout = msoShapeRoundedRectangularCallout
OvalCallout = msoShapeOvalCallout
CloudCallout = msoShapeCloudCallout
''
Process = msoShapeFlowchartProcess
AlternateProcess = msoShapeFlowchartAlternateProcess
Decision = msoShapeFlowchartDecision
Data = msoShapeFlowchartData
PredefinedProcess = msoShapeFlowchartPredefinedProcess
InternalStorage = msoShapeFlowchartInternalStorage
Document = msoShapeFlowchartDocument
Preparation = msoShapeFlowchartPreparation
ManualInput = msoShapeFlowchartManualInput
ManualOperation = msoShapeFlowchartManualOperation
Connector = msoShapeFlowchartConnector
OffpageConnector = msoShapeFlowchartOffpageConnector
PunchedTape = msoShapeFlowchartPunchedTape
SummingJunction = msoShapeFlowchartSummingJunction
OOr = msoShapeFlowchartOr
Collate = msoShapeFlowchartCollate
Sort = msoShapeFlowchartSort
Extract = msoShapeFlowchartExtract
Merge = msoShapeFlowchartMerge
StoredData = msoShapeFlowchartStoredData
Delay = msoShapeFlowchartDelay
SequentialAccessStorage = msoShapeFlowchartSequentialAccessStorage
MagneticDisk = msoShapeFlowchartMagneticDisk
DirectAccessStorage = msoShapeFlowchartDirectAccessStorage
Display = msoShapeFlowchartDisplay
''
ActionButtonCustom = msoShapeActionButtonCustom
ActionButtonHome = msoShapeActionButtonHome
ActionButtonHelp = msoShapeActionButtonHelp
ActionButtonInformation = msoShapeActionButtonInformation
ActionButtonBackorPrevious = msoShapeActionButtonBackorPrevious
ActionButtonForwardorNext = msoShapeActionButtonForwardorNext
ActionButtonBeginning = msoShapeActionButtonBeginning
ActionButtonEnd = msoShapeActionButtonEnd
ActionButtonReturn = msoShapeActionButtonReturn
ActionButtonDocument = msoShapeActionButtonDocument
ActionButtonSound = msoShapeActionButtonSound
ActionButtonMovie = msoShapeActionButtonMovie
''
End Enum

Public Enum enumFormControlType
'''' ********************************************************
''
Button = xlButtonControl
CheckBox = xlCheckBox
DropDown = xlDropDown
TextBox = xlEditBox
GroupBox = xlGroupBox
Label = xlLabel
ListBox = xlListBox
Radio = xlOptionButton
ScrollBar = xlScrollBar
Spinner = xlSpinner
''
End Enum

Public Function GetShapeType(value) As String
'''' ********************************************************
''
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
Dim dictShapeType As Dictionary
Set dictShapeType = New Dictionary
''
Call dictShapeType.Add("Grouped", enumShapeType.Group)
Call dictShapeType.Add("TextBox", enumShapeType.TextBox)
Call dictShapeType.Add("AutoShape", enumShapeType.AutoShape)
Call dictShapeType.Add("FormControl", enumShapeType.FormControl)
Call dictShapeType.Add("OLEObject", enumShapeType.OLEObject)
Call dictShapeType.Add("Table", enumShapeType.Table)
Call dictShapeType.Add("Chart", enumShapeType.Chart)
Call dictShapeType.Add("Comment", enumShapeType.Chart)
Call dictShapeType.Add("Picture", enumShapeType.Chart)
Call dictShapeType.Add("PlaceHolder", enumShapeType.Chart)
Call dictShapeType.Add("Line", enumShapeType.line)
''
GetShapeType = C_Dictionary.GetKeyByItem(dictShapeType, value)
''
End Function

Public Function GetAutoShapeType(value) As String
'''' ********************************************************
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
Dim dictAutoShapeType As Dictionary
Set dictAutoShapeType = New Dictionary
''
Call dictAutoShapeType.Add("ShapeMixed", enumAutoShapeType.Mixed)
Call dictAutoShapeType.Add("Rectangle", enumAutoShapeType.Rectangle)
Call dictAutoShapeType.Add("Oval", enumAutoShapeType.Oval)
Call dictAutoShapeType.Add("FoldedCorner", enumAutoShapeType.FoldedCorner)
Call dictAutoShapeType.Add("RectangularCallout", enumAutoShapeType.RectangularCallout)
Call dictAutoShapeType.Add("RoundedRectangularCallout", enumAutoShapeType.RoundedRectangularCallout)
Call dictAutoShapeType.Add("OvalCallout", enumAutoShapeType.OvalCallout)
Call dictAutoShapeType.Add("CloudCallout", enumAutoShapeType.CloudCallout)
Call dictAutoShapeType.Add("Process", enumAutoShapeType.Process)
Call dictAutoShapeType.Add("AlternateProcess", enumAutoShapeType.AlternateProcess)
Call dictAutoShapeType.Add("Decision", enumAutoShapeType.Decision)
Call dictAutoShapeType.Add("Data", enumAutoShapeType.Data)
Call dictAutoShapeType.Add("PredefinedProcess", enumAutoShapeType.PredefinedProcess)
Call dictAutoShapeType.Add("InternalStorage", enumAutoShapeType.InternalStorage)
Call dictAutoShapeType.Add("Document", enumAutoShapeType.Document)
Call dictAutoShapeType.Add("Preparation", enumAutoShapeType.Preparation)
Call dictAutoShapeType.Add("ManualInput", enumAutoShapeType.ManualInput)
Call dictAutoShapeType.Add("ManualOperation", enumAutoShapeType.ManualOperation)
Call dictAutoShapeType.Add("Connector", enumAutoShapeType.Connector)
Call dictAutoShapeType.Add("OffpageConnector", enumAutoShapeType.OffpageConnector)
Call dictAutoShapeType.Add("PunchedTape", enumAutoShapeType.PunchedTape)
Call dictAutoShapeType.Add("SummingJunction", enumAutoShapeType.SummingJunction)
Call dictAutoShapeType.Add("OOr", enumAutoShapeType.OOr)
Call dictAutoShapeType.Add("Collate", enumAutoShapeType.Collate)
Call dictAutoShapeType.Add("Sort", enumAutoShapeType.Sort)
Call dictAutoShapeType.Add("Extract", enumAutoShapeType.Extract)
Call dictAutoShapeType.Add("Merge", enumAutoShapeType.Merge)
Call dictAutoShapeType.Add("StoredData", enumAutoShapeType.StoredData)
Call dictAutoShapeType.Add("Delay", enumAutoShapeType.Delay)
Call dictAutoShapeType.Add("SequentialAccessStorage", enumAutoShapeType.SequentialAccessStorage)
Call dictAutoShapeType.Add("MagneticDisk", enumAutoShapeType.MagneticDisk)
Call dictAutoShapeType.Add("DirectAccessStorage", enumAutoShapeType.DirectAccessStorage)
Call dictAutoShapeType.Add("Display", enumAutoShapeType.Display)
Call dictAutoShapeType.Add("ActionButtonCustom", enumAutoShapeType.ActionButtonCustom)
Call dictAutoShapeType.Add("ActionButtonHome", enumAutoShapeType.ActionButtonHome)
Call dictAutoShapeType.Add("ActionButtonHelp", enumAutoShapeType.ActionButtonHelp)
Call dictAutoShapeType.Add("ActionButtonInformation", enumAutoShapeType.ActionButtonInformation)
Call dictAutoShapeType.Add("ActionButtonBackorPrevious", enumAutoShapeType.ActionButtonBackorPrevious)
Call dictAutoShapeType.Add("ActionButtonForwardorNext", enumAutoShapeType.ActionButtonForwardorNext)
Call dictAutoShapeType.Add("ActionButtonBeginning", enumAutoShapeType.ActionButtonBeginning)
Call dictAutoShapeType.Add("ActionButtonEnd", enumAutoShapeType.ActionButtonEnd)
Call dictAutoShapeType.Add("ActionButtonReturn", enumAutoShapeType.ActionButtonReturn)
Call dictAutoShapeType.Add("ActionButtonDocument", enumAutoShapeType.ActionButtonDocument)
Call dictAutoShapeType.Add("ActionButtonSound", enumAutoShapeType.ActionButtonSound)
Call dictAutoShapeType.Add("ActionButtonMovie", enumAutoShapeType.ActionButtonMovie)
''
GetAutoShapeType = C_Dictionary.GetKeyByItem(dictAutoShapeType, value)
''
End Function

Public Function GetFormControlType(value) As String
'''' ********************************************************
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
Dim dictFormControlType As Dictionary
Set dictFormControlType = New Dictionary
''
Call dictFormControlType.Add("button", enumFormControlType.Button)
Call dictFormControlType.Add("checkbox", enumFormControlType.CheckBox)
Call dictFormControlType.Add("radio", enumFormControlType.Radio)
Call dictFormControlType.Add("dropdown", enumFormControlType.DropDown)
Call dictFormControlType.Add("textbox", enumFormControlType.TextBox)
Call dictFormControlType.Add("groupbox", enumFormControlType.GroupBox)
Call dictFormControlType.Add("label", enumFormControlType.Label)
Call dictFormControlType.Add("listbox", enumFormControlType.ListBox)
Call dictFormControlType.Add("scrollbar", enumFormControlType.ScrollBar)
Call dictFormControlType.Add("spinner", enumFormControlType.Spinner)
''
GetFormControlType = C_Dictionary.GetKeyByItem(dictFormControlType, value)
''
End Function

Public Function GetListOfValidationWithField(nmStr As String, _
                    Optional shtName As Variant, _
                    Optional wb As Variant) As Variant
'''' *************************************************
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Name As C_Name
Set C_Name = New C_Name
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
''
Dim aryRtn() As Variant
Dim dict As Dictionary
Set dict = New Dictionary
Dim rng As Range
Set rng = C_Name.GetRange(nmStr, shtName, wb) ''FIXME
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
If IsMissing(shtName) Then
    Let shtName = rng.Parent.name
End If
''Console.dump rng
''
Dim lstName As String
If rng.Validation.Type = xlValidateList Then
    Let lstName = C_String.Chomp(rng.Validation.Formula1, 1)
    ''Console.log lstName
End If
''
Dim rngList As Range
Set rngList = C_Name.GetRange(lstName, shtName, wb)
''Console.dump rngList
Dim displayString As Variant
Dim val As Variant
Dim r As Long
For r = 1 To rngList.Rows.Count
    Let displayString = rngList.Cells(r, 1)
    Let val = rngList.Cells(r, 1).Offset(0, 1).value
    ''Console.log rngList.Cells(r, 1).Address
    Call dict.Add("displayStringOpt", displayString)
    Call dict.Add("valueOpt", val)
    Call C_Array.Add(aryRtn, C_Dictionary.DeepCopy(dict))
    Call dict.RemoveAll
Next r
''
Let GetListOfValidationWithField = aryRtn
''
End Function

Public Function IsRangeWithListOfValidation(rng As Range) As Boolean
'''' *************************************************
''
On Error Resume Next
Let IsRangeWithListOfValidation = (rng.Validation.Type = xlValidateList)
If Err.Number <> 0 Then
    Let IsRangeWithListOfValidation = False
End If
''
End Function

Public Function GetNameFieldProperty(ByVal nmStr As String, _
                    Optional shtName As Variant, _
                    Optional wb As Variant) As Dictionary
'''' *************************************************
Dim C_Name As C_Name
Set C_Name = New C_Name
Dim C_Range As C_Range
Set C_Range = New C_Range
Dim C_String As C_String
Set C_String = New C_String
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
Dim C_FieldType As C_FieldType
Set C_FieldType = New C_FieldType
''
Dim CONS_OBJECT_NAME As String
Let CONS_OBJECT_NAME = "RangeName"
Dim CONS_TEXT As String
Let CONS_TEXT = "text"
Dim CONS_TEXTBOX As String
Let CONS_TEXTBOX = "input"
Dim CONS_TEXTAREA As String
Let CONS_TEXTAREA = "textarea"
Dim CONS_SELECT As String
Let CONS_SELECT = "select"
''
''Console.info nmstr
Dim dict As Dictionary
Set dict = New Dictionary
Dim rng As Range
Set rng = C_Name.GetRange(nmStr, shtName, wb)
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
Dim isGlobalName As Boolean
If IsMissing(shtName) Then
    Let shtName = rng.Parent.name
End If
If C_String.IsAddress(nmStr) Then
    Let isGlobalName = False
Else
    Let isGlobalName = True
End If
''Console.Dump rng
''
Dim rngMerged As Range
Dim html_tag As String
Dim field_type As String
If rng.Rows.Count = 1 Then
    If Me.IsRangeWithListOfValidation(rng) Then
        Set rngMerged = rng.MergeArea
        Let html_tag = CONS_SELECT
        Let field_type = CONS_SELECT
    Else
    Set rngMerged = rng.MergeArea
    Let html_tag = CONS_TEXTBOX
        Let field_type = CONS_TEXT
    End If
    If rngMerged.Rows.Count > 1 Then
        Let html_tag = CONS_TEXTAREA
        Let field_type = CONS_TEXTAREA
    End If
Else
    Set rngMerged = rng
    Let html_tag = CONS_TEXTAREA
    Let field_type = CONS_TEXTAREA
End If
''
Dim dictRngFieldProp As Dictionary
Set dictRngFieldProp = New Dictionary
Set dictRngFieldProp = Me.GetRangeFieldProperty(rngMerged)
''
Dim rngVType As String
Let rngVType = C_Range.TypeValueOfCell(wb.Worksheets(shtName).Range(dictRngFieldProp.item("topLeft")))
Dim value As Variant
Let value = wb.Worksheets(shtName).Range(dictRngFieldProp.item("topLeft")).value
''
''
Dim dictAdditional As Dictionary
Set dictAdditional = New Dictionary
Dim jsonstr As String
Let jsonstr = C_String.TrimSpace(rng.NoteText)
If C_String.StartsWith(jsonstr, "[") And C_String.EndsWith(jsonstr, "]") Then
    Set dictAdditional = JsonConverter.ParseJson(jsonstr)
ElseIf C_String.StartsWith(jsonstr, "{") And C_String.EndsWith(jsonstr, "}") Then
    Set dictAdditional = JsonConverter.ParseJson(jsonstr)
    ''Set dictAdditional = JsonConverter.ParseJson(C_String.DeleteWhiteSpace(jsonstr))
End If
If Not dictAdditional.Exists("type") Then
    Call dictAdditional.Add("type", field_type)
End If
''
Call dict.Add("sheetName", shtName)
Call dict.Add("name", nmStr)
Call dict.Add("type", field_type)
Call dict.Add("id", nmStr)
Call dict.Add("tag", html_tag)
Call dict.Add("value", value)
Call dict.Add("displayString", dictRngFieldProp.item("title"))
Call dict.Add("inputValue", value)
Call dict.Add("description", dictRngFieldProp.item("description"))
Call dict.Add("rowPos", dictRngFieldProp.item("topRowPos"))
Call dict.Add("colPos", dictRngFieldProp.item("leftColumnPos"))
Call dict.Add("rowCnt", dictRngFieldProp.item("rowCnt"))
Call dict.Add("colCnt", dictRngFieldProp.item("columnCnt"))
Call dict.Add("objectName", CONS_OBJECT_NAME)
Call dict.Add("valuetype", rngVType)
Call dict.Add("address", rng.Address)
''
Call C_Dictionary.MergeDictionary(dict, dictAdditional)
If field_type = CONS_SELECT Then
    Dim optsAry As Variant
    If isGlobalName Then
        Let optsAry = Me.GetListOfValidationWithField(nmStr, shtName, wb)
    Else
        Let optsAry = Me.GetListOfValidationWithField(nmStr, shtName, wb)
    End If
    Call dict.Add("options", optsAry)
End If
Call dict.Add("_id", C_Commons.CreateGUID)
''
Set GetNameFieldProperty = dict
''
End Function

Public Function GetRangeFieldProperty(rng As Range) As Dictionary
'''' ********************************************************
Dim C_Range As C_Range
Set C_Range = New C_Range
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
''
Dim shtName As String
Let shtName = rng.Parent.name
Dim wb As Workbook
Set wb = rng.Parent.Parent
''
Dim dictRtn As Dictionary
Set dictRtn = New Dictionary
Dim dictDetail As Dictionary
Set dictDetail = New Dictionary
Set dictDetail = C_Range.GetRangeDetail(rng)
''Console.Dump dictRngFieldProp
Dim rowPos As Long
Let rowPos = dictDetail.item("topRowPos")
Dim colPos As Long
Let colPos = dictDetail.item("leftColumnPos")
Dim colPosRight As Long
Let colPosRight = dictDetail.item("rightColumnPos")
Dim title As String
Let title = Me.GetDisplayString(rowPos, colPos, colPosRight, shtName, wb)
Dim descriptionStr As String
Let descriptionStr = Me.GetDescription(rowPos, colPos, colPosRight, shtName, wb)
''
Set dictRtn = C_Dictionary.DeepCopy(dictDetail)
Call dictRtn.Add("title", title)
Call dictRtn.Add("description", descriptionStr)
Set GetRangeFieldProperty = dictRtn
''
End Function

Public Function GetAutoShapeProperty(shp As Shape) As Dictionary
'''' ********************************************************
''''  * @function GetShapeProperty
''''  * @param {Shape} shp
''''  * @return {Dictionary}
''''  */
Dim C_String As C_String
Set C_String = New C_String
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
''
Dim CONS_OBJECT_NAME As String
Let CONS_OBJECT_NAME = "AutoShape"
Dim CONS_TAG_NAME As String
Let CONS_TAG_NAME = "div"
Dim CONS_TYPE_NAME As String
Let CONS_TYPE_NAME = "bottun"
''
Dim sht As Worksheet
Set sht = shp.Parent
Dim shtName As String
Let shtName = sht.name
Dim wb As Workbook
Set wb = sht.Parent
''
Dim dict As Dictionary
Set dict = New Dictionary
''
Dim dictAdditional As Dictionary
Set dictAdditional = New Dictionary
Dim jsonstr As String
Let jsonstr = C_String.TrimSpace(shp.AlternativeText)
If Not C_String.IsEmpty(shp.AlternativeText) Then
    If C_String.StartsWith(jsonstr, "[") And C_String.EndsWith(jsonstr, "]") Then
        Set dictAdditional = JsonConverter.ParseJson(jsonstr)
    ElseIf C_String.StartsWith(jsonstr, "{") And C_String.EndsWith(jsonstr, "}") Then
        Set dictAdditional = JsonConverter.ParseJson(jsonstr)
    End If
End If
If Not dictAdditional.Exists("type") Then
    Call dictAdditional.Add("type", CONS_TYPE_NAME)
End If
''
Dim topLeftCell As String
Let topLeftCell = shp.topLeftCell.Address
Dim bottomRightCell As String
Let bottomRightCell = shp.bottomRightCell.Address
Dim rng As Range
Set rng = sht.Range(sht.Range(topLeftCell), sht.Range(bottomRightCell))
Dim dictRngFieldProp As Dictionary
Set dictRngFieldProp = New Dictionary
Set dictRngFieldProp = Me.GetRangeFieldProperty(rng)
''
Call dict.Add("sheetName", shp.Parent.name)
Call dict.Add("name", shp.name)
Call dict.Add("type", "")
Call dict.Add("tag", CONS_TAG_NAME)
Call dict.Add("title", dictRngFieldProp.item("title"))
Call dict.Add("value", shp.TextFrame.Characters.Text)
Call dict.Add("description", dictRngFieldProp.item("description"))
Call dict.Add("rowPos", dictRngFieldProp.item("topRowPos"))
Call dict.Add("colPos", dictRngFieldProp.item("leftColumnPos"))
Call dict.Add("rowCnt", dictRngFieldProp.item("rowCnt"))
Call dict.Add("colCnt", dictRngFieldProp.item("columnCnt"))
Call dict.Add("objectName", CONS_OBJECT_NAME)
Call dict.Add("objectType", Me.GetAutoShapeType(shp.autoShapeType))
Call dict.Add("address", rng.Address)
''
Call C_Dictionary.MergeDictionary(dict, dictAdditional)
Call dict.Add("_id", C_Commons.CreateGUID)
''
Set GetAutoShapeProperty = dict
''
End Function

Public Function IsShapeOleWithLinkedCell(shp As Shape) As Boolean
'''' *************************************************
''
On Error Resume Next
Dim vnt As Variant
Let vnt = shp.OLEFormat.Object.LinkedCell
If Err.Number <> 0 Then
    Let IsShapeOleWithLinkedCell = False
Else
    Let IsShapeOleWithLinkedCell = True
End If
''
End Function

Public Function GetFormControlProperty(shp As Shape) As Dictionary
'''' *****************************************
Dim C_String As C_String
Set C_String = New C_String
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
''
Dim CONS_OBJECT_NAME As String
Let CONS_OBJECT_NAME = "FormControl"
Dim CONS_TAG_INPUT As String
Let CONS_TAG_INPUT = "input"
Dim CONS_TAG_SELECT As String
Let CONS_TAG_SELECT = "select"
''
Dim sht As Worksheet
Set sht = shp.Parent
Dim shtName As String
Let shtName = sht.name
Dim wb As Workbook
Set wb = sht.Parent
''
Dim dict As Dictionary
Set dict = New Dictionary
''
Dim dictAdditional As Dictionary
Set dictAdditional = New Dictionary
Dim jsonstr As String
Let jsonstr = C_String.TrimSpace(shp.AlternativeText)
If Not C_String.IsEmpty(shp.AlternativeText) Then
    If C_String.StartsWith(jsonstr, "[") And C_String.EndsWith(jsonstr, "]") Then
        Set dictAdditional = JsonConverter.ParseJson(jsonstr)
    ElseIf C_String.StartsWith(jsonstr, "{") And C_String.EndsWith(jsonstr, "}") Then
        Set dictAdditional = JsonConverter.ParseJson(jsonstr)
    End If
End If
If Not dictAdditional.Exists("type") Then
    Call dictAdditional.Add("type", CONS_TAG_INPUT)
End If
''
''Console.info "C_FieldType.GetFormControlProperty." & shp.name
''Console.dump shp
''
Dim topLeftCell As String
Let topLeftCell = shp.topLeftCell.Address
Dim bottomRightCell As String
Let bottomRightCell = shp.bottomRightCell.Address
Dim rng As Range
Set rng = sht.Range(sht.Range(topLeftCell).Address, sht.Range(bottomRightCell).Address)
Dim dictRngFieldProp As Dictionary
Set dictRngFieldProp = New Dictionary
Set dictRngFieldProp = Me.GetRangeFieldProperty(rng)
''
Call dict.Add("sheetName", shp.Parent.name)
Call dict.Add("name", shp.name)
Call dict.Add("type", TypeName(shp.OLEFormat.Object))
Call dict.Add("id", shp.name) ''FIXME
Call dict.Add("tag", CONS_TAG_INPUT)
Call dict.Add("text", shp.TextFrame.Characters.Text)
Call dict.Add("displayString", dictRngFieldProp.item("title"))
If Me.IsShapeOleWithLinkedCell(shp) Then
  Call dict.Add("inputValue", sht.Range(shp.OLEFormat.Object.LinkedCell).value)
Else
  Call dict.Add("inputValue", "")
End If
''
Call dict.Add("description", dictRngFieldProp.item("description"))
Call dict.Add("rowPos", dictRngFieldProp.item("topRowPos"))
Call dict.Add("colPos", dictRngFieldProp.item("leftColumnPos"))
Call dict.Add("rowCnt", dictRngFieldProp.item("rowCnt"))
Call dict.Add("colCnt", dictRngFieldProp.item("columnCnt"))
''
If Me.IsShapeOleWithLinkedCell(shp) Then
  Call dict.Add("linkedCell", shp.OLEFormat.Object.LinkedCell)
Else
  Call dict.Add("linkedCell", "")
End If
''
Call dict.Add("objectName", CONS_OBJECT_NAME)
Call dict.Add("objectType", Me.GetFormControlType(shp.formControlType))
Call dict.Add("address", rng.Address)
''
Call C_Dictionary.MergeDictionary(dict, dictAdditional)
Call dict.Add("_id", C_Commons.CreateGUID)
''
Set GetFormControlProperty = dict
''
End Function

''Public Function GetOleProperty(ole As OLEObject) As Dictionary
'''''' *****************************************
''''FIXME : I have made a program for FormControl, so OLEObject can not be available for this tool.
''Dim dict As Dictionary
''Set dict = New Dictionary
''''
''Call dict.Add("name", ole.name)
''Call dict.Add("type", TypeName(ole.Object))
''Call dict.Add("address", ole.topLeftCell.Address)
''Call dict.Add("rowPos", ole.topLeftCell.Row)
''Call dict.Add("colPos", ole.topLeftCell.column)
''Call dict.Add("LinkedCell", ole.LinkedCell)
''Set GetOleProperty = dict
''''
''End Function

Public Function GetDisplayString(rowPos As Long, colPos As Long, colPosRight As Long, _
            ByVal shtName As String, _
            ByVal wb As Workbook, _
            Optional ByRef rngGroup As Variant) As String
'''' *****************************************
Dim C_String As C_String
Set C_String = New C_String
''
If IsMissing(rngGroup) Then
    Set rngGroup = wb.Worksheets(shtName).UsedRange
End If
Dim rngTargetRow1 As Range
Set rngTargetRow1 = wb.Worksheets(shtName).Range(Cells(rowPos - 1, 1).Address, Cells(rowPos, colPos - 1).Address)
Dim rngTargetRow2 As Range
Set rngTargetRow2 = wb.Worksheets(shtName).Range(Cells(rowPos - 1, colPos).Address, Cells(rowPos - 1, colPosRight).Address)
Dim rngIntersected As Range
Set rngIntersected = Application.Intersect(rngGroup, Application.Union(rngTargetRow1, rngTargetRow2))
Dim rng As Range
Dim aryary As Variant
Let GetDisplayString = ""
Dim i As Long
Dim str As String
On Error Resume Next
    Set rng = rngIntersected.SpecialCells(XlSpecialCellsValue.xlTextValues)
    If rng.Areas.Count = 1 Then
        Let GetDisplayString = rng.value
    Else
        For i = 1 To rng.Areas.Count
            str = str + rng.Areas(i).value + ": "
        Next i
        Let GetDisplayString = C_String.Chop(str, 2)
    End If
End Function

Public Function GetDescription(rowPos As Long, colPos As Long, colPosRight As Long, _
            ByVal shtName As String, _
            ByVal wb As Workbook, _
            Optional ByRef rngGroup As Variant) As String
'''' *****************************************
Dim C_String As C_String
Set C_String = New C_String
''
If IsMissing(rngGroup) Then
    Set rngGroup = wb.Worksheets(shtName).UsedRange
End If
Dim rngTargetRow As Range
Set rngTargetRow = wb.Worksheets(shtName).Range(Cells(rowPos + 1, colPos).Address, Cells(rowPos + 1, colPosRight).Address)
Dim rngIntersected As Range
Set rngIntersected = Application.Intersect(rngGroup, rngTargetRow)
Dim rng As Range
Dim aryary As Variant
Let GetDescription = ""
Dim i As Long
Dim str As String
On Error Resume Next
    Set rng = rngIntersected.SpecialCells(XlSpecialCellsValue.xlTextValues)
    If rng.Areas.Count = 1 Then
        Let GetDescription = rng.value
    Else
        For i = 1 To rng.Areas.Count
            str = str + rng.Areas(i).value + "  "
        Next i
        Let GetDescription = C_String.Chop(str, 2)
    End If
End Function


'' 25/05/2020 Add lines

Public Function GetFormControlsWithPrefix(ByVal prefix As String, ByVal sheetName As String, Optional ByRef wb As Variant) As Collection
'''' *****************************************
''
Dim C_Collection As C_Collection
Set C_Collection = New C_Collection
''
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
Dim sht As Worksheet
Set sht = wb.Worksheets(sheetName)
''
Dim shps As Collection
Set shps = C_Collection.FilterNameMatchPrefixOfObject(sht.Shapes, prefix)
''
Set GetFormControlsWithPrefix = shps
''
End Function

Public Function FilterShapeWithType(ByRef shps As Variant, ByVal shapeType As String) As Collection
'''' ********************************************************
''
Dim arySHAPETYPE() As Variant
Let arySHAPETYPE = [{"AutoShape","FormControl","Grouped"}]
''
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim cltRtn As Collection
Set cltRtn = New Collection
Dim cl As Variant
Dim typ As Long
Dim shp As Shape
For Each shp In shps
    Let typ = shp.Type
    If C_Array.Exists(arySHAPETYPE, Me.GetShapeType(typ)) Then
        If Me.GetShapeType(typ) = shapeType Then
            cltRtn.Add shp
        End If
    End If
Next shp
''
Set FilterShapeWithType = cltRtn
''
End Function

Public Function FilterAutoShapeWithType(ByRef shps As Variant, ByVal autoShpType As String) As Collection
'''' ********************************************************
''
Dim CONS_SHAPETYPE As String
Let CONS_SHAPETYPE = "AutoShape"
''
Dim cltRtn As Collection
Set cltRtn = New Collection
Dim clt As Collection
Set clt = New Collection
Set clt = Me.FilterShapeWithType(shps, CONS_SHAPETYPE)
''
Dim typ As Long
Dim shp As Shape
For Each shp In clt
    Let typ = shp.autoShapeType
    If VBA.UCase(Me.GetAutoShapeType(typ)) = VBA.UCase(autoShpType) Then
        cltRtn.Add shp
    End If
Next shp
''
Set FilterAutoShapeWithType = cltRtn
''
End Function

Public Function FilterAutoShapeWithSpecName(ByRef shps As Variant, Optional ByVal shpNameStart As Variant) As Collection
'''' ********************************************************
'' SpecName : Not contain spaces
''
Dim CONS_SHAPETYPE As String
Let CONS_SHAPETYPE = "AutoShape"
''
Dim C_String As C_String
Set C_String = New C_String
Dim isSpecName As Boolean
If IsMissing(shpNameStart) Then
    Let isSpecName = True
Else
    Let isSpecName = False
End If
''
Dim cltRtn As Collection
Set cltRtn = New Collection
Dim clt As Collection
Set clt = New Collection
Set clt = Me.FilterShapeWithType(shps, CONS_SHAPETYPE)
''
Dim nm As String
Dim shp As Shape
For Each shp In clt
    Let nm = shp.name
    If isSpecName Then
        If Not C_String.Contain(nm, " ") Then
            cltRtn.Add shp
        End If
    Else
        If C_String.StartsWith(nm, shpNameStart) Then
            cltRtn.Add shp
        End If
    End If
Next shp
''
Set FilterAutoShapeWithSpecName = cltRtn
''
End Function

Public Function GetAutoShapeFields(ByVal sheetName As String, Optional ByRef wb As Variant) As Collection
'''' ********************************************************
''
Dim cltRtn As Collection
Set cltRtn = New Collection
''
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
Dim shps As Shapes
Set shps = wb.Worksheets(sheetName).Shapes
Set cltRtn = Me.FilterAutoShapeWithSpecName(shps)
''
Set GetAutoShapeFields = cltRtn
''
End Function

Public Function FilterFormControlWithType(ByRef shps As Variant, ByVal formCtrlType As String) As Collection
'''' ********************************************************
''
Dim CONS_SHAPETYPE As String
Let CONS_SHAPETYPE = "FormControl"
''
Dim cltRtn As Collection
Set cltRtn = New Collection
Dim clt As Collection
Set clt = New Collection
Set clt = Me.FilterShapeWithType(shps, CONS_SHAPETYPE)
''
Dim typ As Long
Dim shp As Shape
For Each shp In clt
    Let typ = shp.formControlType
    If VBA.UCase(Me.GetFormControlType(typ)) = VBA.UCase(formCtrlType) Then
        'Console.dump shp
        cltRtn.Add shp
    End If
Next shp
''
Set FilterFormControlWithType = cltRtn
''
End Function

Public Function FilterFormControlWithSpecName(ByRef shps As Variant, _
                   Optional ByVal shpNameStart As Variant, _
                   Optional ByVal formCotrolTypeName As Variant) As Collection
'''' ********************************************************
'' SpecName : Not contain spaces
''
Dim CONS_SHAPETYPE As String
Let CONS_SHAPETYPE = "FormControl"
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim aryControlFieldType() As Variant
Let aryControlFieldType = [{"checkbox","radio","groupbox"}]
''
Dim isSpecName As Boolean
If IsMissing(shpNameStart) Then
    Let isSpecName = True
Else
    Let isSpecName = False
End If

''
Dim cltRtn As Collection
Set cltRtn = New Collection
Dim clt As Collection
Set clt = New Collection
Set clt = Me.FilterShapeWithType(shps, CONS_SHAPETYPE)
''
Dim nm As String
Dim shp As Shape
For Each shp In clt
    Let nm = shp.name
    If C_Array.Exists(aryControlFieldType, Me.GetFormControlType(shp.formControlType)) Then
    If isSpecName Then
        If Not C_String.Contain(nm, " ") Then
            cltRtn.Add shp
        End If
    Else
        If C_String.StartsWith(nm, shpNameStart) Then
            cltRtn.Add shp
        End If
    End If
    End If
Next shp
''
Set FilterFormControlWithSpecName = cltRtn
''
End Function

Public Function GetFormControlFields(ByVal sheetName As String, Optional ByRef wb As Variant) As Collection
'''' ********************************************************
''
Dim cltRtn As Collection
Set cltRtn = New Collection
''
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
Dim shps As Shapes
Set shps = wb.Worksheets(sheetName).Shapes
Set cltRtn = Me.FilterFormControlWithSpecName(shps)
''
Set GetFormControlFields = cltRtn
''
End Function

Public Function GetFormControlFieldsOfGroupbox(ByVal sheetName As String, Optional ByRef wb As Variant) As Collection
'''' ********************************************************
''
Dim cltRtn As Collection
Set cltRtn = New Collection
''
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
Dim shps As Shapes
Set shps = wb.Worksheets(sheetName).Shapes
Set cltRtn = Me.FilterFormControlWithType(shps, "groupbox")
''
Set GetFormControlFieldsOfGroupbox = cltRtn
''
End Function

Public Function GetFormControlFieldsOfCheckbox(ByVal sheetName As String, Optional ByRef wb As Variant) As Collection
'''' ********************************************************
''
Dim cltRtn As Collection
Set cltRtn = New Collection
''
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
Dim shps As Shapes
Set shps = wb.Worksheets(sheetName).Shapes
Set cltRtn = Me.FilterFormControlWithType(shps, "checkbox")
''
Set GetFormControlFieldsOfCheckbox = cltRtn
''
End Function

Public Function GetGroupboxProperty(shpOfGroup As Shape) As Dictionary
'''' *****************************************
''
Dim C_Range As C_Range
Set C_Range = New C_Range
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_String As C_String
Set C_String = New C_String
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
''
Dim CONS_PREFIX_GROUP As String
Let CONS_PREFIX_GROUP = "grp_"
Dim CONS_PREFIX_RADIO As String
Let CONS_PREFIX_RADIO = "rdo_"
Dim CONS_PREFIX_CHECKBOX As String
Let CONS_PREFIX_CHECKBOX = "chk_"
''
Dim shps As Shapes
Set shps = shpOfGroup.Parent.Shapes
''
Dim dictRtn As Dictionary
Set dictRtn = Me.GetFormControlProperty(shpOfGroup)
''
''Console.dump dict
''
Dim rngEx As Range
Set rngEx = shpOfGroup.Parent.Range(dictRtn.item("address"))
''
''' radio's process
Dim dict As Dictionary
Set dict = New Dictionary
Dim rng As Range
Dim ary() As Variant
Dim cl As Shape
Dim clt As Collection
Set clt = Me.FilterFormControlWithType(shps, "radio")
Dim isRadio As Boolean
Let isRadio = False
For Each cl In clt
    ''Console.info "C_FieldType.GetGroupboxProperty." & cl.name
    Set dict = Me.GetFormControlProperty(cl)
    Set rng = cl.Parent.Range(dict.item("address"))
    If C_Range.IsInRange(rng, rngEx) Then
        'Console.dump dict
        Let isRadio = True
        Call C_Array.Add(ary, dict)
    End If
Next cl
''
''' checkbox's process
Dim isCheckbox As Boolean
Let isCheckbox = False
If Not isRadio Then
    Set clt = Me.FilterFormControlWithType(shps, "checkbox")
    For Each cl In clt
        Set dict = Me.GetFormControlProperty(cl)
        Set rng = cl.Parent.Range(dict.item("address"))
        If C_Range.IsInRange(rng, rngEx) Then
            'Console.dump dict
            Let isCheckbox = True
            Call C_Array.Add(ary, dict)
        End If
    Next cl
End If
'' populate
Dim typeNameStr As String
If isRadio Then
    Let typeNameStr = "radio"
ElseIf isCheckbox Then
    Let typeNameStr = "checkbox"
Else
    Let typeNameStr = "text" ''FIXME
End If
Let dictRtn.item("type") = typeNameStr
Let dictRtn.item("name") = C_String.CutPrefix(dictRtn.item("name"), CONS_PREFIX_GROUP)
''
Dim dc As Dictionary
Set dc = New Dictionary
Dim ar() As Variant
Dim arId() As Variant
Dim aryElmValue() As Variant
Dim aryElmInputValue() As Variant
Dim addr As String
Dim dictAdditional As Dictionary
Set dictAdditional = New Dictionary
Dim jsonstr As String
Dim i As Long
For i = LBound(ary) To UBound(ary)
    Set dict = ary(i)
    ''Console.dump dict
    '' name
    Call dc.Add("name", dictRtn.item("name"))
    Call dc.Add("elmType", typeNameStr)
    '' elmId
    If typeNameStr = "radio" Then
        Call dc.Add("elmId", C_String.CutPrefix(dict.item("name"), CONS_PREFIX_RADIO))
    ElseIf typeNameStr = "checkbox" Then
        Call dc.Add("elmId", C_String.CutPrefix(dict.item("name"), CONS_PREFIX_CHECKBOX))
    End If
    '' elmValue
    Let addr = dict.item("linkedCell")
    Call dc.Add("elmLinkedAddress", addr)
    Set rng = shpOfGroup.Parent.Range(addr)
    Let jsonstr = C_String.TrimSpace(rng.NoteText)
    If C_String.StartsWith(jsonstr, "[") And C_String.EndsWith(jsonstr, "]") Then
        Set dictAdditional = JsonConverter.ParseJson(jsonstr)
    ElseIf C_String.StartsWith(jsonstr, "{") And C_String.EndsWith(jsonstr, "}") Then
        Set dictAdditional = JsonConverter.ParseJson(jsonstr)
        ''Set dictAdditional = JsonConverter.ParseJson(C_String.DeleteWhiteSpace(jsonstr))
    End If
    ''
    If dictAdditional.Exists("value") Then
        Call dc.Add("elmValue", dictAdditional.item("value"))
    Else
        Call dc.Add("elmValue", dc.item("elmId"))
    End If
    Call C_Array.Add(aryElmValue, dc.item("elmValue"))
    '' elmDisplayString
    Call dc.Add("elmDisplayString", dict.item("text"))
    Call dc.Add("elmInputValue", dict.item("inputValue"))
    Call C_Array.Add(aryElmInputValue, dc.item("elmInputValue"))
    ''
    Call C_Array.Add(ar, C_Dictionary.DeepCopy(dc))
    Call C_Array.Add(arId, dict.item("name"))
    Call dc.RemoveAll
Next i
''
Let aryElmValue = C_Array.ChangeIndexOfBegin(aryElmValue)
Let aryElmInputValue = C_Array.ChangeIndexOfBegin(aryElmInputValue)
Call dictRtn.Add("values", aryElmValue)
Dim inputValue As Variant
If isRadio Then
    Let inputValue = VBA.Join(C_Array.Unique(aryElmInputValue), ",")
    Let dictRtn.item("inputValue") = dictRtn.item("values")(CInt(inputValue))
ElseIf isCheckbox Then
    Dim aryTemp() As Variant
    For i = LBound(aryElmInputValue) To UBound(aryElmInputValue)
        If aryElmInputValue(i) Then
            Call C_Array.Add(aryTemp, aryElmValue(i))
        End If
        Let inputValue = VBA.Join(aryTemp, ",")
        Let dictRtn.item("inputValue") = inputValue
    Next i
End If
Call dictRtn.Add("inputValue2", aryElmInputValue)
''
Call dictRtn.Add("elms", ar)
Call dictRtn.Add("elmIds", arId)
''
Set GetGroupboxProperty = dictRtn
''
End Function

