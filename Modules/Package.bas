Attribute VB_Name = "Package"
Option Explicit

Public C_String As C_String
''
Public C_Array  As C_Array
Public C_Range  As C_Range
Public C_Name   As C_Name
Public C_ListObject As C_ListObject
Public C_Dictionary As C_Dictionary
Public C_Collection As C_Collection
''
Public C_Template As C_Template
Public O_Template As O_Template
Public template   As MiniTemplator
''
Public C_File   As C_File
Public C_FileIO As C_FileIO
Public C_Book   As C_Book
Public C_Sheet  As C_Sheet
''


Public Function Include()
'''' ********************************************************
Set C_String = New C_String
''
Set C_Array = New C_Array
Set C_Range = New C_Range
Set C_Name = New C_Name
Set C_ListObject = New C_ListObject
Set C_Dictionary = New C_Dictionary
Set C_Collection = New C_Collection
''
Set C_Template = New C_Template
Set O_Template = New O_Template
Set template = New MiniTemplator
''
Set C_File = New C_File
Set C_FileIO = New C_FileIO
Set C_Book = New C_Book
Set C_Sheet = New C_Sheet
''
End Function

Public Function Terminate()
'''' ********************************************************
Set C_String = Nothing
''
Set C_Array = Nothing
Set C_Range = Nothing
Set C_Name = Nothing
Set C_ListObject = Nothing
Set C_Dictionary = Nothing
Set C_Collection = Nothing
''
Set C_Template = Nothing
Set O_Template = Nothing
Set template = Nothing
''
Set C_File = Nothing
Set C_FileIO = Nothing
Set C_Book = Nothing
Set C_Sheet = Nothing
''
End Function
