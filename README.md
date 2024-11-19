## Common VBA _Private Profile_ file services
Simplifies and unifies _Private Profile_ file services by supporting:
- Section separation (default)
- File header and footer (optional)
- Section and value comments (optional)
- Sections and value.names in ascending order (on the fly)
- Any number of _Private Profile_ files in one class instance [^1] 

The following is an example of the simplification:  
- Value read: `<value> = Value(<value_name>, <section_name>[, <file_name>])`
- Value write: `Value(<value_name>, <section_name>[, <file_name>]) = <value>`. 
A unique experience thereby: All sections and values are written in ascending order which eases reading of large cfg-, ini-, and other kinds of _Private Profile_ files.

### The _Value_ service
The service has the following named arguments:

| Argument        | Explanation |
|-----------------|-------------|
| _name\_value_   | String expression, name of the value in the _Private Profile_ file.|
| _name\_section_ | String expression, optional, specifies the _Section_ for the value, defaults to the section specified through the property _Section_.|
| _name\_file_    | String expression, optional, specifies the full name of a _Private Profile_ file, defaults to file specified through the property _FileName_ when omitted.|

### Summary of services

| Service            | Type     | Description |
|--------------------|---------------|-------------|
|_FileName_&nbsp;r/w | Property | Returns/specifies a _Private Profile_ file's full name. When none has been specified the file name defaults to _ThisWorkBook.Path & "\PrivateProfile.dat". |
|_Section_&nbsp;r/w  | Property | Returns/specifies the Section name used throughout subsequent services when the section argument is omitted. |
|_Value_&nbsp;r/w    | Property | Reads from a _Private Profile_ file a value by a provided value-name in a provided section (defaults to the _Section_ property when omitted). When the file name is omitted it defaults to the name specified by the _FileName_ property.|
|          |                    | Writes to a _Private Profile_ file a value by a provided value-name in a provided section (defaults to the _Section_ property when omitted). When the file name is omitted it defaults to the name specified by the _FileName_ property.|
|_ValueRemove_       | Method | Removes one or more values including a possible value comment from one, more or all sections in a _Private Profile_ file, whereby value-names may be provided as a comma separated string and section-names may be provided as a comma separated string. When no file (name_file) is provided it defaults to the file name specified by the property FileName. <br>**Attention!** When no section/s is/are specified, the value/value-name is removed in all sections the name is used. |
|_SectionExists_     | Method   | Returns TRUE when a given section exists in the current valid _Private Profile_ file.|
|_SectionNames_      | Method   | Returns a Dictionary of all section names.|
|_SectionRemove_    | Method   | Removes one (or more specified as a comma delimited string) section. Sections not existing are ignored. When no file (name_file) is provided it defaults to the file name specified by the property FileName.|
|_ValueNameExists_   | Method   | Returns TRUE when a given value-name exists in a provided _Private Profile_ file's section.|
|_ValueNameRename_   | Method   | Replaces an old value name with a new one either in sections provided as a comma delimited string or in all sections when none are provided.|
|_ValueNames_        | Method   | Returns a Dictionary with all value names a _Private Profile_ file with the value name as the key and the value as the item, of all sections if none is provided or those of a provided section's name. When the file name is omitted it defaults to the name specified by the _FileName_ property.<br>***Note:*** The returned value-names are distinct names! I.e. when a value exists in more than one section it is still one distinct value-name.|
| _SectionSeparation_ | Property | Boolean expression, default to True, separates sections by an empty line to improve readability.|

## Installation
1. Download and import [clsPrivProf.cls][1] to your VB project.
2. In the VBE add a Reference to _Microsoft Scripting Runtime_ and _Microsoft VBScript Regular Expression 5.5_

## Usage (example)
The following example uses a class module named `clsConfig` for read/write of values in a "<name>.cfg" file which provides a Get/Let property for each value.

```vb
Option Explicit
Private PPfile As clsPrivProf

Private Sub Class_Initialize()
    Set PPFile = New clsPrivProf
    With PPFile
        .FileName = <the-file's-full-name>
        .FileHeader = "any" ' optional
        .FileFooter = "any" ' optional
        .Section = "Configuration" ' in this example, all values in one section
    End With
End Sub

' -----------------------------------------------------------------------------------
' Sample property for a certain value/value-name.
' -----------------------------------------------------------------------------------
Friend Property Get Example() As String:                Example = Value(<value-name>):          End Property
Friend Property Let Example(ByVal l_value As String):   Value(<value-name>, <value>) = l_value: End Property

' -----------------------------------------------------------------------------------
' Interface to the clsPrivProf class Value service.
' -----------------------------------------------------------------------------------
Private Property Get Value(Optional ByVal v_section_name As String = vbNullString, _
                           Optional ByVal v_value_name As String = vbNullString) As String
    Const PROC = "Value/Let"
    If v_section_name = vbNullString Then Err.Raise AppErr(1), ErrSrc(PROC), "No section-name provided!"
    If v_value_name = vbNullString Then Err.Raise AppErr(2), ErrSrc(PROC), "No value-name provided!"
    Value = PPFile.Value(v_value_name, v_section_name)
End Property

Private Property Let Value(Optional ByVal v_section_name As String = vbNullString, _
                           Optional ByVal v_value_name As String = vbNullString, _
                                    ByVal v_value As String)
    Const PROC = "Value/Let"
    If v_section_name = vbNullString Then Err.Raise AppErr(1), ErrSrc(PROC), "No section-name provided!"
    If v_value_name = vbNullString Then Err.Raise AppErr(2), ErrSrc(PROC), "No value-name provided!"
    PPFile.Value(v_value_name, v_section_name) = v_value
End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error number never conflicts
' with VB runtime error. Thr function returns a given positive number
' (app_err_no) with the vbObjectError added - which turns it to negative. When
' the provided number is negative it returns the original positive "application"
' error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "clsConfig." & e_proc
End Function
```

> This _Common Component_ is prepared to function completely autonomously (download, import, use) but at the same time is prepared to integrate with my personal 'standard' VB-Project design. See [Conflicts with personal and public _Common Components_][2] for more details.

## Contribution
Any kind of contribution is welcome. Respecting the (more or less obvious) coding rules and guidelines will be appreciated. The module is available in a dedicated [Workbook][3] (public GitHub repository). This Workbook also provides a complete regression test covering all public services (methods and properties).

[^1]: It may be more elegant to use a class instance for each individual _Private Profile_ file, provide the file by the _FileName_ property once and omit it in all other properties and methods. This is reflected by the given example.

[1]:https://github.com/warbe-maker/VBA-Private-Profile/blob/main/CompMan/source/clsPrivProf.cls
[2]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
[3]:https://github.com/warbe-maker/VBA-Private-Profile
