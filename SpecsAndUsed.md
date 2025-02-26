
## Services

| Service       | Kind [^1] | Argument    | Explanation |
|---------------|:---------:|-----------------|-------------|
|_FileBaseName_ | | P       |                 |             |
|_FileExtention_| P         |                 |             |
|_FileFullName_ | P r/w     |                 |Returns/specifies a _Private Profile_ file's full name. When none has been specified the file name defaults to _ThisWorkBook.Path & "\PrivateProfile.dat". |
|_Value_        | P r/w     |                 |Reads from / writes to a _Private Profile_ file a value.|
|               |           |_value\_name_   | String expression, name of the value in the _Private Profile_ file.|
|               |           | _section\_name_ | String expression, optional, specifies the _Section_ for the value, defaults to the section specified through the property _Section_.|
|               |           | _file\_name_    | String expression, optional, specifies the full name of a _Private Profile_ file, defaults to file specified through the property _FileName_ when omitted.|
|_ValueRemove_  | M     |                |Removes one or more values including a possible value comment from one, more or all sections in a _Private Profile_ file, whereby value-names may be provided as a comma separated string and section-names may be provided as a comma separated string. When no file (name_file) is provided it defaults to the file name specified by the property FileName. <br>**Attention!** When no section/s is/are specified, the value/value-name is removed in all sections the name is used. |
|_SectionExists_     | M   |  | Returns TRUE when a given section exists in the current valid _Private Profile_ file.|
|_SectionNames_      | M   |  | Returns a Dictionary of all section names.|
|_SectionRemove_    | M   | | Removes one (or more specified as a comma delimited string) section. Sections not existing are ignored. When no file (name_file) is provided it defaults to the file name specified by the property FileName.|
|_ValueNameExists_   | M   | | Returns TRUE when a given value-name exists in a provided _Private Profile_ file's section.|
|_ValueNameRename_   | M   | | Replaces an old value name with a new one either in sections provided as a comma delimited string or in all sections when none are provided.|
|_ValueNames_        | M   | | Returns a Dictionary with all value names a _Private Profile_ file with the value name as the key and the value as the item, of all sections if none is provided or those of a provided section's name. When the file name is omitted it defaults to the name specified by the _FileName_ property.<br>***Note:*** The returned value-names are distinct names! I.e. when a value exists in more than one section it is still one distinct value-name.|
| _SectionSeparation_ | P w | | Boolean expression, default to True, separates sections by an empty line to improve readability.|

[1]: **P**roperty **r**ead/**w**rite or **M**ethod

## Implementation
The module uses a "twin" in the form of a Dictionary/Collections structure together with a copy of the Private Profile file as string. The twin is rebuilt whenever the content changes - either the Private Profile file's content by any other means but the component itself or via one of the modules methods or properties.

## Usage of the component
A kind of best practice may be a class module dedicated to the VB-Project's specific Private Profile values which makes use of the clsPrivProf class module for read/write
```vb
Option Explicit
' ------------------------------------------------------------------------------
' Class Module clsMyPrivProf: Provides the VB-Project's specific values 
' =========================== read/write services.
' ------------------------------------------------------------------------------
Private Const ANY_VALUE_NAME As String = "MyValueName"
Private PP                   As New clsPrivProf

Private Sub Class_Initialize
    PP.FileBaseName = "MyPrivProfFile"
    ' PP.FileLocation = ' optionally modifies the default (ActiveWorkbook's path)
    ' PP.FileExtention = ' optionally modifies the default (.dat)
    ' PP.FileFullName = ' alternatively specifies the file's full name
 End Sub
 
 '~~ clsPrivProf interface
 Private Property Value Get(Optional ByVal v_value_name As String, _
                            Optional ByVal v_section_name As String = vbNullString) As String
         Value = PP.Value(v_value_name, v_section_name)
End Property

 Private Property Value Let(Optional ByVal v_value_name As String, _
                            Optional ByVal v_section_name As String = vbNullString, _
                                     ByVal v_value As String)
         PP.Value(v_value_name, v_section_name) = v_value
End Property

'~~ Example for a VB-Project specific Private Profile value
Public Property Get AnyValue(Optional By Al a_section_name As String = vbNullString) As String
    AnyValue= Value(ANY_VALUE_NAME, a_section_name)
End Property

Public Property Let AnyValue(Optional ByVAl a_section_name As String = vbNullString, 
                                      ByVal  a_value As String)
    Value(ANY_VALUE_NAME, a_section_name) = a_value
End Property

```

