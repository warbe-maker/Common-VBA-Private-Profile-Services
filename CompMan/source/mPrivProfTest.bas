Attribute VB_Name = "mPrivProfTest"
Option Explicit
#Const mTrc = 1
' ----------------------------------------------------------------
' Standard Module mPrivProvTest: Test of all services provided by
' ============================== the clsPrivProf class module.
' Usually each test is autonomous and preferrably uses no or only
' tested other Properties/Methods.
'
' Uses:
' - clsTestAid      Common services supporting test including
'                   regression testing.
' - clsPrivProfTests Services supporting tests of methods and
'                   properties of the class module clsPrivProf.
' - mTrc            Execution trace of tests.
'
' W. Rauschenberger, Berlin May 2024
' See also https://github.com/warbe-maker/VBA-Private-Profile.
' ----------------------------------------------------------------
Public TestAid              As clsTestAid

Private PrivProf            As clsPrivProf
Private PrivProfTests       As New clsPrivProfTests
Private cllResultExpectd    As Collection
Private FSo                 As New FileSystemObject
Private vResult             As Variant

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

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
#ElseIf XcTrc_clsTrc Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP b_proc, b_args
#ElseIf XcTrc_mTrc Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mPrivProfTest." & sProc
End Function

Public Sub Prepare(Optional ByVal p_no As Long = 1, _
                   Optional ByVal p_init As Boolean = True)
' ----------------------------------------------------------------------------
' Prepares for a new test or a series of tests:
' 1. A test Private Profile file considering a number (p_no)
' 2. A new clsPrivProf class instance
' 3. By default, initializes the FileName property (p_init)
' Note: By default a file ....1.dat (p_no) is setup from scratch, other
' numbers (p_no) may just copy a backup.
' ----------------------------------------------------------------------------
    Const PROC = "Prepare"
    
    On Error GoTo eh
    Dim sFile As String
    
    If TestAid Is Nothing Then
        Set TestAid = New clsTestAid
        With TestAid
            .ModeRegression = mErH.Regression
            .TestFileExtension = "dat"
            .TestedComp = "clsPrivProf"
        End With
    End If
    
    Set PrivProf = Nothing
    Set PrivProf = New clsPrivProf

    If Not TestAid.ModeRegression Then
        mTrc.FileFullName = TestAid.TestFolder & "\ExecTrace.log"
        mTrc.Title = "Test class module clsPrivProf"
        mTrc.NewFile
    End If
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_000_Regression()
' ----------------------------------------------------------------------------
' Please note: All results are programmatically asserted and thus there is no
' manual intervention during this test. In case an assertion fails the test
' procedure will  n o t  stop but keep a record of the failed assertion.
'
' An execution trace is displayed at the end.
' ----------------------------------------------------------------------------
    Const PROC = "Test_000_Regression"

    On Error GoTo eh
    Dim sTestStatus     As String
    Dim bModeRegression As Boolean
    
    '~~ Initialization (must be done prior the first BoP!)
    Set PrivProfTests = New clsPrivProfTests
    mTrc.FileFullName = TestAid.TestFolder & "\Regression.ExecTrace.log"
    mTrc.Title = "Regression Test class module clsPrivProf"
    mTrc.NewFile
    bModeRegression = True
    mErH.Regression = bModeRegression
    TestAid.ModeRegression = bModeRegression
    TestAid.CleanUp "*Failed.log", "*-Result.*" ' remove any files resulting from individual tests
    
    BoP ErrSrc(PROC)
    sTestStatus = "clsPrivProf Regression Test: "

    mPrivProfTest.Test_100_Property_FileName
    mPrivProfTest.Test_110_Method_Exists
    mPrivProfTest.Test_120_Properties
    mPrivProfTest.Test_300_Method_SectionNames
    mPrivProfTest.Test_400_Method_ValueNames
    mPrivProfTest.Test_410_Method_ValueNameRename
    mPrivProfTest.Test_500_Method_Remove
'    mPrivProfTest.Test_600_Lifecycle
'    mPrivProfTest.Test_700_HskpngNames
    TestAid.ResultSummaryLog
    
xt: EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Set TestAid = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_100_Property_FileName()
    Const PROC = "Test_100_Property_FileName"
    
    On Error GoTo eh:
        
    BoP ErrSrc(PROC)
    Prepare 1, False ' The FileName property is provided in the test
    
    With TestAid
        .TestId = "100-1"
        .TestedProc = "FileName-Let"
        .TestedProcType = "Property"
        .Title = "Property FileName"
        
        .Verification = "Initialize PP-file"
        .ResultExpected = .TestFileFullName("Result")
        .TimerStart
        PrivProf.FileFullName = .TestFileFullName("Result")
        .TimerEnd
        .Result = PrivProf.FileFullName
        ' ======================================================================
        
        .TestId = "100-2"
        .TestedProc = "Let FileName"
        .TestedProcType = "Property"
        
        .Verification = "Specifying the full name of an existing Private Profile file"
        .ResultExpected = .TestFileFullName("Result")
        .TimerStart
        PrivProf.FileFullName = .TestFileFullName("Result") ' continue with current test file
        .TimerEnd
        .Result = PrivProf.FileFullName
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_110_Method_Exists()
    Const PROC = "Test_110_Methods_Exists"

    On Error GoTo eh
    Dim sResultFile As String
    
    BoP ErrSrc(PROC)
    '~~ Test preparation
    Prepare
       
    With TestAid
         TestAid.CleanUp "*110*Failed.log", "*110*-Result.*" ' remove any files resulting from individual tests

        .TestId = "110-1" ' initiates a new test
        .Title = "Exists method"
        .TestedProc = "Exists"
        .TestedProcType = "Method"
        sResultFile = .TestFileFullName("Result")
        PrivProfTests.ProvideTestPrivProf 1, sResultFile ' default private profile file
        PrivProf.FileFullName = sResultFile
        
        ' 1. ----------------------------------------------------------------------------
        .Verification = "Section not exists"
        .TimerStart
        vResult = PrivProf.Exists(sResultFile, PrivProfTests.SectionName(7))
        .TimerEnd
        .Result = vResult
        .ResultExpected = False
        
        ' 2. ----------------------------------------------------------------------------
        .Verification = "Section exists"
        .TimerStart
        vResult = PrivProf.Exists(sResultFile, PrivProfTests.SectionName(8))
        .TimerEnd
        .Result = vResult
        .ResultExpected = True
        
        ' 3. ----------------------------------------------------------------------------
        .Verification = "Value-Name exists"
        .TimerStart
        vResult = PrivProf.Exists(sResultFile, PrivProfTests.SectionName(6), PrivProfTests.ValueName(6, 4))
        .TimerEnd
        .Result = vResult
        .ResultExpected = True
        
        ' 4. ----------------------------------------------------------------------------
        .Verification = "Value-Name not exists"
        .TimerStart
        vResult = PrivProf.Exists(sResultFile, PrivProfTests.SectionName(6), PrivProfTests.ValueName(6, 3))
        .TimerEnd
        .Result = vResult
        .ResultExpected = False
    
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_120_Properties()
' ----------------------------------------------------------------------------
' This test relies on the Value (Let) service.
' The whole test, i.e. all verifications operate on the same Private Profile
' file prepared.
' ----------------------------------------------------------------------------
    Const PROC = "Test_120_Properties"
    
    On Error GoTo eh
    Dim cyValue     As Currency: cyValue = 12345.6789
    Dim fleResult   As File
    Dim sResultFile As String
    
    BoP ErrSrc(PROC)
    Prepare
    
    With TestAid
        .TestId = "120-1"
        .Title = "Properties Value, FileHeader/Footer, SectionComment, ValueComment"
        .TestedProc = "Value-Get"
        .TestedProcType = "Property"
        sResultFile = .TestFileFullName("Result")

        ' 1. -------------------------------------------------------------------------------------------------
        .Verification = "Read non-existing value from a non-existing file returns the specified default value"
        .TimerStart
        vResult = PrivProf.Value(v_value_name:="Any" _
                               , v_section_name:="Any" _
                               , v_value_default:="Not available")
        .TimerEnd
        .Result = vResult
        .ResultExpected = "Not available"
        
        '~~ All subsequent verifications from now on operate on the same Private Profile
        '~~ result file which will have a changed content after any write operation (Property-Let).
        '~~ Made changes a verified versus "result expected Private Profile files".
        PrivProfTests.ProvideTestPrivProf 1, sResultFile ' default private profile file
        PrivProf.FileFullName = sResultFile
        
        ' 2. -------------------------------------------------------------------
        .Verification = "Read existing value"
        .ResultExpected = PrivProfTests.ValueString(2, 4)
        .TimerStart
        vResult = PrivProf.Value(v_value_name:=PrivProfTests.ValueName(2, 4) _
                               , v_section_name:=PrivProfTests.SectionName(2))
        .TimerEnd
        .Result = vResult
        
        ' 3. -------------------------------------------------------------------
        .Verification = "Write value"
        .TestedProc = "Value-Let"
        .TestedProcType = "Property"
        .TimerStart
        PrivProf.Value(v_value_name:=PrivProfTests.ValueName(4, 2) _
                     , v_section_name:=PrivProfTests.SectionName(4)) = "Changed value"
        .TimerEnd
        .Result = PrivProf.Value(v_value_name:=PrivProfTests.ValueName(4, 2) _
                               , v_section_name:=PrivProfTests.SectionName(4))
        .ResultExpected = "Changed value"
        
        ' 4. -------------------------------------------------------------------
        .Verification = "Write new value in existing section"
        .TestedProc = "Value-Let"
        .TestedProcType = "Property"
        .ResultExpected = "New value, existing section"
        .TimerStart
        PrivProf.Value(PrivProfTests.ValueName(2, 17) _
                    , PrivProfTests.SectionName(2)) = "New value, existing section"
        .TimerEnd
        .Result = PrivProf.Value(v_value_name:=PrivProfTests.ValueName(2, 17) _
                              , v_section_name:=PrivProfTests.SectionName(2))
        
        ' 5. -------------------------------------------------------------------
        .Verification = "Write new value in new section"
        .TimerStart
        PrivProf.Value(v_value_name:=PrivProfTests.ValueName(11, 1) _
                     , v_section_name:=PrivProfTests.SectionName(11)) = "New value, new section"
        .TimerEnd
        .Result = FSo.GetFile(sResultFile)
        .ResultExpected = .TestResultExpectedFile
        
        ' 6. -------------------------------------------------------------------
        .Verification = "Change value plus the value and the section comments"
        .TimerStart
        PrivProf.Value(v_value_name:=PrivProfTests.ValueName(11, 1) _
                     , v_section_name:=PrivProfTests.SectionName(11) _
                      ) = "Changed new value, new section"
        .TimerEnd
        .Result = FSo.GetFile(sResultFile)
        .ResultExpected = .TestResultExpectedFile
        
        ' 7. -------------------------------------------------------------------
        .Verification = "Changed again"
        .TimerStart
        PrivProf.Value(v_value_name:=PrivProfTests.ValueName(11, 1) _
                     , v_section_name:=PrivProfTests.SectionName(11) _
                      ) = "Changed again new value, new section"
        .TimerEnd
        .Result = FSo.GetFile(sResultFile)
        .ResultExpected = .TestResultExpectedFile
    
        ' 8. -------------------------------------------------------------------
        .Verification = "Write a file header (not provided delimiter line added by system)"
        .TestedProc = "FileHeader-Let"
        .TestedProcType = "Property"
        .TimerStart
        '~~ Note: For the missing file name the property FileName is used
        '~~ and the missing section- and value-name indicate a file header
        PrivProf.FileHeader() = "File Comment Line 1 (the header delimiter is adjusted to the longest header line)" & vbCrLf & _
                                "File Comment Line 2"
        .TimerEnd
        .Result = FSo.GetFile(sResultFile)
        .ResultExpected = .TestResultExpectedFile

        ' 9. -------------------------------------------------------------------
        .Verification = "File header read (returned as lines delimited by vbCrLf)"
        .TestedProc = "FileHeader-Get"
        .TestedProcType = "Property"
        .TimerStart
        vResult = PrivProf.FileHeader()
        .TimerEnd
        .Result = vResult
        .ResultExpected = "File Comment Line 1 (the header delimiter is adjusted to the longest header line)" & vbCrLf & _
                          "File Comment Line 2"
        
        ' 10. -------------------------------------------------------------------
        .Verification = "Write section comment"
        .TestedProc = "SectionComment-Let"
        .TestedProcType = "Property"
        
        .TimerStart
        PrivProf.SectionComment(PrivProfTests.SectionName(6)) _
                              = "Comment Section 06 Line 1" & vbCrLf & _
                                "Comment Section 06 Line 2"
        .TimerEnd
        .Result = FSo.GetFile(sResultFile)
        .ResultExpected = .TestResultExpectedFile
        
        ' 11. -------------------------------------------------------------------
        .Verification = "Read section comment"
        .TestedProc = "SectionComment-Get"
        .TestedProcType = "Property"
        .TimerStart
        vResult = PrivProf.SectionComment(PrivProfTests.SectionName(6))
        .TimerEnd
        .Result = vResult
        .ResultExpected = "Comment Section 06 Line 1" & vbCrLf & _
                          "Comment Section 06 Line 2"
        
        ' 12. -------------------------------------------------------------------
        .Verification = "Write value comment"
        .TestedProc = "ValueComment-Let"
        .TestedProcType = "Property"
        .TimerStart
        PrivProf.ValueComment(PrivProfTests.ValueName(6, 2), PrivProfTests.SectionName(6)) _
                           = "Comment Section 06 Value 02 Line 1" & vbCrLf & _
                             "Comment Section 06 Value 02 Line 2"
        .TimerEnd
        .Result = FSo.GetFile(sResultFile)
        .ResultExpected = .TestResultExpectedFile
        
        ' 13. -------------------------------------------------------------------
        .Verification = "Read value comment(returned as Collection of lines)"
        .TestedProc = "ValueComment-Get"
        .TestedProcType = "Property"
        .TimerStart
        vResult = PrivProf.ValueComment(PrivProfTests.ValueName(6, 2), PrivProfTests.SectionName(6))
        .TimerEnd
        .Result = vResult
        .ResultExpected = "Comment Section 06 Value 02 Line 1" _
               & vbCrLf & "Comment Section 06 Value 02 Line 2"
        
        ' 14. -------------------------------------------------------------------
        .Verification = "Write file footer (with a delimiter line above)"
        .TestedProc = "FileFooter-Let"
        .TestedProcType = "Property"
        .TimerStart
        PrivProf.FileFooter() = "File footer line 1" & vbCrLf & _
                                "File footer line 2"
        .TimerEnd
        .Result = FSo.GetFile(sResultFile)
        .ResultExpected = .TestResultExpectedFile
        
        ' 15. -------------------------------------------------------------------
        .Verification = "Read file footer (with delimiter line excluded)"
        .TestedProc = "FileFooter-Get"
        .TestedProcType = "Property"
        .TimerStart
        vResult = PrivProf.FileFooter
        .TimerEnd
        .Result = vResult
        .ResultExpected = "File footer line 1" & vbCrLf & _
                          "File footer line 2"
        
        
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_300_Method_SectionNames()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_300_Method_SectionNames"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .Title = "SectionNames service"
        .TestId = "300-1"
        
        .Verification = "Get all section names in a Dictionary"
        .TestedProc = "SectionNames"
        .TestedProcType = "Function"
        .TimerStart
        vResult = PrivProf.SectionNames().Count
        .TimerEnd
        .Result = vResult
        .ResultExpected = 5
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Set dct = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_400_Method_ValueNames()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_400_Method_ValueNames"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .Title = "ValueNames service"
        .TestId = "400-1"
        
        .Verification = "Get all value names of all sections in a Dictionary"
        .TestedProc = "ValueNames"
        .TestedProcType = "Function"
        .TimerStart
        Set dct = PrivProf.ValueNames()
        .TimerEnd
        .Result = dct.Count
        .ResultExpected = 40
        
        .Verification = "Get all value names of a certain section in a Dictionary"
        .TestedProc = "ValueNames"
        .TestedProcType = "Function"
        .TimerStart
        vResult = PrivProf.ValueNames(, PrivProfTests.SectionName(6)).Count
        .TimerEnd
        .Result = vResult
        .ResultExpected = 8
                
        .Verification = "Get all value names of all sections in a Dictionary"
        .TestedProc = "ValueNames"
        .TestedProcType = "Function"
        .TimerStart
        vResult = PrivProf.ValueNames().Count
        .TimerEnd
        .Result = vResult
        .ResultExpected = 6
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_410_Method_ValueNameRename()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_400_Method_ValueNameRename"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
       
    With TestAid
        .Title = "ValueNameRename service"
        .TestId = "410-1"
        .TestedProc = "ValueNameRename"
        .TestedProcType = "Method"
        PrivProf.FileFullName = .TempFileFullName("Result")
        
        .Verification = "Rename a value name in each section."
        .ResultExpected = .TestFile("Result-Expected")
        .TimerStart
        PrivProf.ValueNameRename PrivProfTests.ValueName(2, 2), "Renamed_" & PrivProfTests.ValueName(2, 2)
        .TimerEnd
        .Result = .TestFile("Result")
        ' ======================================================================
    End With
    
xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_500_Method_Remove()
' ----------------------------------------------------------------------------
' The test relies on: - Comment value
' ----------------------------------------------------------------------------
    Const PROC = "Test_500_Method_Remove"
    
    On Error GoTo eh
    Dim sFile As String
    
    BoP ErrSrc(PROC)
    Prepare ' Test preparation
    
    With TestAid
        .Title = "Remove service"
        PrivProf.ValueComment(PrivProfTests.SectionName(6), PrivProfTests.ValueName(6, 4)) = "Comment value 06-04"
        .TestId = "500-1"
        .TestedProc = "ValueRemove"
        .TestedProcType = "Method"
        PrivProf.FileFullName = .TempFileFullName("Result")
        
        .Verification = "Remove a value from a section including its comments."
        .ResultExpected = .TestFile("Result-Expected")
        .TimerStart
        PrivProf.ValueRemove PrivProfTests.ValueName(6, 4), PrivProfTests.SectionName(6)
        .TimerEnd
        .Result = .TestFile("Result")
        
        ' ======================================================================
        .Title = vbNullString
        PrivProf.SectionComment(PrivProfTests.SectionName(6)) = "Comment section 06"
        
        .TestId = "500-2"
        .TestedProc = "SectionRemove"
        .TestedProcType = "Method"
        PrivProf.FileFullName = .TempFileFullName("Result")
        
        .Verification = "Removes a section including its comments."
        .ResultExpected = .TestFile("Result-Expected")
        .TimerStart
        PrivProf.SectionRemove PrivProfTests.SectionName(6)
        .TimerEnd
        .Result = .TestFile("Result")
        
        ' ======================================================================
        .TestId = "500-3"
        Prepare 2
        .TestedProc = "ValueRemove"
        .TestedProcType = "Method"
        
        .Verification = "Remove 2 names in 2 sections."
        .ResultExpected = .TestFile("Result-Expected")
        .TimerStart
        PrivProf.ValueRemove v_value_name:="Last_Modified_AtDateTime,Last_Modified_InWbkFullName", v_section_name:="clsLog,clsQ"
        .TimerEnd
        .Result = .TestFile("Result")
        ' ======================================================================
    
        .TestId = "500-4"
        Prepare 2
        .TestedProc = "ValueRemove"
        .TestedProcType = "Method"
        PrivProf.FileFullName = .TempFileFullName("Result")
        
        .Verification = "Removes all values in one section which removes the section."
        .ResultExpected = .TestFile("Result-Expected")
        .TimerStart
        PrivProf.ValueRemove v_value_name:="ExportFileExtention" & _
                                         ",Last_Modified_AtDateTime" & _
                                         ",Last_Modified_InWbkFullName" & _
                                         ",Last_Modified_InWbkName" & _
                                         ",LastModExpFileFullNameOrigin" _
                           , v_section_name:="clsQ"
        .TimerEnd
        .Result = .TestFile("Result")
        ' ======================================================================
       
        .TestId = "500-5"
        Prepare 2
        .TestedProc = "ValueRemove"
        .TestedProcType = "Method"
        PrivProf.FileFullName = .TempFileFullName("Result")
        
        .Verification = "Remove all values in all sections - file is removed."
        sFile = PrivProfTests.PrivProfFile
        .ResultExpected = False
        .TimerStart
        PrivProf.ValueRemove v_value_name:="ExportFileExtention" & _
                                         ",Last_Modified_AtDateTime" & _
                                         ",Last_Modified_InWbkFullName" & _
                                         ",Last_Modified_InWbkName" & _
                                         ",LastModExpFileFullNameOrigin" & _
                                         ",DoneNamesHskpng"
        .TimerEnd
        .Result = FSo.FileExists(sFile) ' is False
        ' ======================================================================
    End With

xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_600_Lifecycle()
' ----------------------------------------------------------------------------
' Test beginning with a non existing Private Profile file, performing some
' services.
' ----------------------------------------------------------------------------
    Const PROC = "Test_600_Lifecycle"
    
    On Error GoTo eh
    Dim sTestResultFile As String
    
    Prepare 0, False
    BoP ErrSrc(PROC)
    
    With TestAid
        .Title = "A PrivateProfile file life-cycle"
        .TestId = "600-1"
        
        .Verification = "New file, header and footer writing postponed"
        sTestResultFile = .TempFileFullName("Result") ' result test file for the whoöe test procedure
        .TestedProc = "FileHeader-Let, FileFooter-Let"
        .TestedProcType = "Property"
        '~~ Header and/or footer writing to a not yet active file will be
        '~~ postponed until at least one user-value had been written.
        If FSo.FileExists(PrivProfTests.PrivProfFileFullName) Then .FSo.DeleteFile PrivProfTests.PrivProfFileFullName
        Set PrivProf = New clsPrivProf
        PrivProf.FileFullName = sTestResultFile
        .ResultExpected = False
        .TimerStart
        PrivProf.FileFooter() = "File Footer Line 1 (the delimiter below is adjusted to the longest comment)" & vbCrLf & _
                                "File Footer Line 2"
        PrivProf.FileHeader() = "File Comment Line 1 (the delimiter below is adjusted to the longest comment)" & vbCrLf & _
                                "File Comment Line 2"
        .TimerEnd
        .Result = FSo.FileExists(sTestResultFile)
        
        ' -------------------------------------------------------------------
        .Verification = "First value also writes postponed header and footer"
        .TestedProc = "Value-Let"
        .TestedProcType = "Property"
        
        .TimerStart
        PrivProf.Value(v_value_name:="Any-Value-Name" _
                     , v_section_name:="Any-Section-Name" _
                     , v_file_name:=PrivProfTests.PrivProfFileFullName _
                      ) = "Any-Value"
        .TimerEnd
        .ResultExpected = .TestFile("Result-Expected")
        .Result = FSo.GetFile(sTestResultFile)
        
        ' -------------------------------------------------------------------
        .Verification = "Removes the only value in the only section removes file"
        .TestedProc = "ValueRemove"
        .TestedProcType = "Method"
        '~~ Removing the only value in the only section ends with no file
        '~~ Note: This is consequent since there is no Private Profile file without
        '~~       at least on section with one value
        Set PrivProf = New clsPrivProf
        PrivProf.FileFullName = sTestResultFile
        .ResultExpected = True
        .TimerStart
        mErH.Asserted AppErr(1) ' effective only when mErH.Regression = True
        PrivProf.ValueRemove v_value_name:="Any-Value-Name" _
                           , v_section_name:="Any-Section-Name"
        .TimerEnd
        mErH.Asserted ' reset to none
        .Result = Not FSo.FileExists(sTestResultFile)
        ' ======================================================================
    
    End With

xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_700_HskpngNames()
' ----------------------------------------------------------------------------
' Test beginning with a non existing Private Profile file, performing some
' services.
' ----------------------------------------------------------------------------
    Const PROC = " Test_700_HskpngNames"
    
    On Error GoTo eh
    Dim fleTestResult As File
    
    BoP ErrSrc(PROC)
    
    Prepare 2   ' uses a ready for test file copied from a backup
    With TestAid
        .Title = "Names housekeeping"
        .TestId = "700-1"
        .TestedProc = "HouskeepingNames"
        .TestedProcType = "Method"
        
        .Verification = "One Value-name change in two sections"
        .ResultExpected = .TestFile("Result-Expected")
        .TimerStart
        PrivProf.HskpngNames PrivProf.FileFullName, "clsLog:clsQ:Last_Modified_AtDateTime>Last_Modified_UTC_AtDateTime"
        .TimerEnd
        .Result = .TestFile("Result")
        ' ======================================================================
                
        .TestId = "700-2"
        Prepare 2   ' uses a ready for test file copied from a backup
        
        .Verification = "Two value-name changes in all sections"
        .ResultExpected = .TestFile("Result-Expected")
        .TimerStart
        PrivProf.HskpngNames PrivProf.FileFullName, "Last_Modified_AtDateTime>Last_Modified_UTC_AtDateTime" _
                                              , "LastModExpFileFullNameOrigin>Last_Modified_ExpFileFullNameOrigin"
        .TimerEnd
        .Result = .TestFile("Result")
        Set fleTestResult = PrivProfTests.PrivProfFile
        ' ======================================================================
        
        .TestId = "700-3"
        .Verification = "Two value-name changes in all sections (any subsequent once done)"
        .ResultExpected = fleTestResult
        .TimerStart
        PrivProf.HskpngNames fleTestResult.Path, "Last_Modified_AtDateTime>Last_Modified_UTC_AtDateTime" _
                                               , "LastModExpFileFullNameOrigin>Last_Modified_ExpFileFullNameOrigin"
        .TimerEnd
        .Result = .TestFile("Result")
        ' ======================================================================
                
    End With

xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

