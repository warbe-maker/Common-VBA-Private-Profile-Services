Attribute VB_Name = "mTest"
Option Explicit
Public Const NO_OF_TEST_SECTIONS    As Long = 10
Public Const NO_OF_TEST_VALUE_NAMES As Long = 16

Public PrivProf                     As clsPrivProf
Public PrivProfTest                 As clsPrivProfTest
Public Tests                        As clsTestAid
Public FSo                          As New FileSystemObject

Private Const SECTION_NAME          As String = "Section_"      ' for PrivateProfile services test
Private Const VALUE_NAME_INDIVIDUAL As String = "_Name_"        ' for PrivateProfile services test
Private Const VALUE_NAME            As String = "Value_Name_"   ' for PrivateProfile services test
Private Const VALUE_STRING          As String = "-Value-"       ' for PrivateProfile services test

Private cllTestFiles                As Collection
Private sPrivProfFile               As String
Private sFileName                   As String

Public Property Get FileName() As String:   FileName = sFileName:   End Property

Public Sub Init(Optional ByVal i_default As Boolean = False)
    
    If PrivProf Is Nothing Then Set PrivProf = New clsPrivProf
    If Tests Is Nothing Then Set Tests = New clsTestAid
    Tests.TestedComp = "clsPrivProf"
    sFileName = mTest.PrivateProfile_File(mTest.NO_OF_TEST_SECTIONS, mTest.NO_OF_TEST_VALUE_NAMES)
    If Not i_default Then
        PrivProf.FileName = sFileName
    End If
End Sub

Public Property Let FileString(Optional ByVal f_file_full_name As String, _
                               Optional ByVal f_append As Boolean = False, _
                               Optional ByVal f_exclude_empty As Boolean = False, _
                                        ByVal f_s As String)
' ----------------------------------------------------------------------------
' Writes a string (f_s) with multiple records/lines delimited by a vbCrLf to
' a file (f_file_full_name).
' ----------------------------------------------------------------------------
    
    If f_append _
    Then Open f_file_full_name For Append As #1 _
    Else Open f_file_full_name For Output As #1
    Print #1, f_s
    Close #1
        
End Property

Public Function FileAsString(Optional ByVal f_file_full_name As String, _
                             Optional ByVal f_append As Boolean = False, _
                             Optional ByRef f_split As String = vbCrLf, _
                             Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file_full_name) as a single string plus the
' records/lines delimiter (f_split) which may be vbCrLf, vbCr, or vbLf.
' ----------------------------------------------------------------------------

    Open f_file_full_name For Input As #1
    FileAsString = Input$(LOF(1), 1)
    Close #1
    
    Select Case True
        Case InStr(FileAsString, vbCrLf) <> 0: f_split = vbCrLf
        Case InStr(FileAsString, vbCr) <> 0:   f_split = vbCr
        Case InStr(FileAsString, vbLf) <> 0:   f_split = vbLf
    End Select
    
    '~~ Eliminate a trailing eof if any
    If Right(FileAsString, 1) = VBA.Chr(26) Then
        FileAsString = Left(FileAsString, Len(FileAsString) - 1)
    End If
    
    '~~ Eliminate any trailing split string
    If Right(FileAsString, Len(f_split)) = f_split Then
        FileAsString = Left(FileAsString, Len(FileAsString) - Len(f_split))
    End If
    If f_exclude_empty Then
        FileAsString = FileAsStringEmptyExcluded(FileAsString)
    End If
    
End Function

Public Function FileAsArray(ByVal f_file_full_name As String) As Variant
    Dim sSplit As String
    FileAsArray = Split(FileAsString(f_file_full_name, , sSplit), sSplit)
End Function

Private Property Get FileTemp(Optional ByVal f_path As String = vbNullString, _
                              Optional ByVal f_extension As String = ".tmp") As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(FSo.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = CurDir
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    FileTemp = sTemp
    
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTest." & sProc
End Function

Private Function FileAsStringEmptyExcluded(ByVal s_s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (s_s) with any empty elements excluded. I.e. the string
' returned begins and ends with a non vbNullString character and has no
' ----------------------------------------------------------------------------
    
    s_s = FileStringTrimmed(s_s) ' leading and trailing empty already excluded
    Do While InStr(s_s, vbCrLf & vbCrLf) <> 0
        s_s = Replace(s_s, vbCrLf & vbCrLf, vbCrLf)
    Loop
    FileAsStringEmptyExcluded = s_s
    
End Function

Private Function FileStringTrimmed(ByVal s_s As String, _
                          Optional ByRef s_as_dict As Dictionary = Nothing) As String
' ----------------------------------------------------------------------------
' Returns a file as string (s_s) with any leading and trailing empty items,
' i.e. record, lines, excluded. When a Dictionary is provided
' the string is additionally returned as items with the line number as key.
' ----------------------------------------------------------------------------
    Dim s As String
    Dim i As Long
    Dim v As Variant
    
    s = s_s
    '~~ Eliminate any leading empty items
    Do While Left(s, 2) = vbCrLf
        s = Right(s, Len(s) - 2)
    Loop
    '~~ Eliminate a trailing eof if any
    If Right(s, 1) = VBA.Chr(26) Then
        s = Left(s, Len(s) - 1)
    End If
    '~~ Eliminate any trailing empty items
    Do While Right(s, 2) = vbCrLf
        s = Left(s, Len(s) - 2)
    Loop
    
    FileStringTrimmed = s
    If Not s_as_dict Is Nothing Then
        With s_as_dict
            For Each v In Split(s, vbCrLf)
                i = i + 1
                .Add i, v
            Next v
        End With
    End If
    
End Function

Public Property Get PrivProfFile() As File
    Set PrivProfFile = FSo.GetFile(sPrivProfFile)
End Property

Private Sub ArrayAdd(ByRef a_array As Variant, _
                     ByVal a_str As String)
    On Error Resume Next
    ReDim Preserve a_array(UBound(a_array) + 1)
    If Err.Number <> 0 Then ReDim a_array(0)
    a_array(UBound(a_array)) = a_str
    
End Sub

Public Function PrivateProfile_File(Optional ByVal t_sections As Long = NO_OF_TEST_SECTIONS, _
                                    Optional ByVal t_values As Long = NO_OF_TEST_VALUE_NAMES, _
                                    Optional ByVal t_individual_names As Boolean = True) As String
' ----------------------------------------------------------------------------
' Returns the name of a temporary file with n (t_sections) sections, each
' with m (t_values) values all in descending order. Each test file's name is
' saved to a Collection (cllTestFiles) allowing to delete them all at the end
' of the test.
' When t_individual_names is FALSE all sections have the same set of value
' names.
' ----------------------------------------------------------------------------
    Const PROC = "PrivateProfile_File"

    On Error GoTo eh
    Dim i           As Long
    Dim j           As Long
    Dim sFolder     As String
    Dim arr()       As Variant
    
    sFolder = Tests.TestFolder
    If Not FSo.FolderExists(sFolder) Then FSo.CreateFolder (sFolder)
    
    sPrivProfFile = sFolder & "\" & FSo.GetBaseName(ThisWorkbook.Name) & ".dat"
    If FSo.FileExists(sPrivProfFile) Then FSo.DeleteFile sPrivProfFile
    
    For i = t_sections To 1 Step -2
        ArrayAdd arr, "[" & SectionName(i) & "]"
        For j = t_values To 1 Step -2
            If t_individual_names _
            Then ArrayAdd arr, ValueName(i, j) & "=" & ValueString(i, j) _
            Else ArrayAdd arr, ValueName(, j) & "=" & ValueString(i, j)
        Next j
    Next i
    StringAsFile Join(arr, vbCrLf), sPrivProfFile
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sPrivProfFile
    PrivateProfile_File = sPrivProfFile
    
xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function StringAsFile(ByVal s_strng As String, _
                     Optional ByRef s_file As Variant = vbNullString, _
                     Optional ByVal s_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes a string (s_strng) to a file (s_file) which might be a file object or
' a file's full name. When no file (s_file) is provided, a temporary file is
' returned.
' Note 1: Only when the string has sub-strings delimited by vbCrLf the string
'         is written a records/lines.
' Note 2: When the string has the alternate split indicator "|&|" this one is
'         replaced by vbCrLf.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    Select Case True
        Case s_file = vbNullString: s_file = TempFile
        Case TypeName(s_file) = "File": s_file = s_file.Path
    End Select
    
    If s_file_append _
    Then Open s_file For Append As #1 _
    Else Open s_file For Output As #1
    Print #1, s_strng
    Close #1
    Set StringAsFile = FSo.GetFile(s_file)
    
End Function

Public Sub RemoveTestFiles()

    Dim v As Variant
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    For Each v In cllTestFiles
        If FSo.FileExists(v) Then
            Kill v
        End If
    Next v
    Set cllTestFiles = Nothing
    Set cllTestFiles = New Collection
    
End Sub

Public Function SectionName(ByVal l As Long) As String
    SectionName = SECTION_NAME & Format(l, "00")
End Function

Private Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErH = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.EoP e_proc, e_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP e_proc, e_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP e_proc, e_args
#End If
End Sub

Private Sub BoP(ByVal b_proc As String, _
       Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErH Then          ' serves the mTrc/clsTrc when installed and active
    mErH.BoP b_proc, b_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP b_proc, b_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

Public Function TempFile() As String
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "TempFile"

    On Error GoTo eh
    Dim sFileName   As String

    BoP ErrSrc(PROC)
    sFileName = FileTemp(f_extension:=".dat")
    TempFile = sFileName

    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sPrivProfFile

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function ValueName(Optional ByVal t_section_name As Long = 0, _
                                   Optional ByVal t_value_name As Long = 0) As String
    If t_section_name <> 0 _
    Then ValueName = SECTION_NAME & Format(t_section_name, "00") & VALUE_NAME_INDIVIDUAL & Format(t_value_name, "00") _
    Else ValueName = VALUE_NAME & Format(t_value_name, "00")
    
End Function

Public Function ValueString(ByVal lS As Long, ByVal lV As Long) As String
    ValueString = SECTION_NAME & Format(lS, "00") & VALUE_STRING & Format(lV, "00")
End Function

