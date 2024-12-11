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
Public PrivProf         As clsPrivProf
Public PrivProfTests    As New clsPrivProfTests
Public TestAid          As clsTestAid
Private cllExpctd       As Collection
Private FSo             As New FileSystemObject
Private vResult         As Variant

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

Private Sub BoC(ByVal b_id As String, _
       Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Bnd-of-Code' interface for the Common VBA Execution Trace Service.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mTrc = 1 Then         ' when mTrc is installed and active
    mTrc.BoC b_id, b_args
#ElseIf clsTrc = 1 Then   ' when clsTrc is installed and active
    Trc.BoC b_id, b_args
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
#ElseIf XcTrc_clsTrc Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP b_proc, b_args
#ElseIf XcTrc_mTrc Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

Private Sub EoC(ByVal e_id As String, _
       Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End-of-Code' interface for the Common VBA Execution Trace Service.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mTrc = 1 Then         ' when mTrc is installed and active
    mTrc.EoC e_id, e_args
#ElseIf clsTrc = 1 Then   ' when clsTrc is installed and active
    Trc.EoC e_id, e_args
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
' Prepares for a new test or series of tests:
' 1. A test Private Profile file considering a nmber (p_no)
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
        TestAid.ModeRegression = mErH.Regression
        mTrc.FileFullName = TestAid.TestFolder & "\ExecTrace.log"
        mTrc.Title = "Test class module clsPrivProf"
        mTrc.NewFile
    End If
    
    Set PrivProf = Nothing
    Set PrivProf = New clsPrivProf
    PrivProfTests.ProvideTestPrivProf p_no
    If p_init And p_no <> 0 Then PrivProf.FileName = PrivProfTests.PrivProfFileFullName

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
    TestAid.CleanUp "Result_*" ' remove any files resulting from individual tests
    
    BoP ErrSrc(PROC)
    sTestStatus = "clsPrivProf Regression Test: "

    mPrivProfTest.Test_001_TestAid
    mPrivProfTest.Test_100_Property_FileName
    mPrivProfTest.Test_110_Method_Exists
    mPrivProfTest.Test_120_Property_Value
    mPrivProfTest.Test_200_Property_Comments
    mPrivProfTest.Test_300_Method_SectionNames
    mPrivProfTest.Test_400_Method_ValueNames
    mPrivProfTest.Test_410_Method_ValueNameRename
    mPrivProfTest.Test_500_Method_Remove
    mPrivProfTest.Test_600_Lifecycle
    mPrivProfTest.Test_700_HskpngNames
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

Public Sub Test_001_TestAid()
' ----------------------------------------------------------------------------
' Test of the means (clsTestAid) used by all tests.
' ----------------------------------------------------------------------------
    Const PROC = "Test_001_TestAid"
    
    On Error GoTo eh
    Dim sFileResult     As String
    Dim sFileExpected   As String
    
    BoP ErrSrc(PROC)
    Prepare 1, False
    With TestAid
        .TestId = "001-1"
        .TestedComp = "clsTestAid"  ' remains the default for all subsequent tests
        .TestedProc = "Result-Let, ResultExpected-Let"
        .TestedType = "Property"
        .Verification = "Result is the result expected"
        .ResultExpected = True
        .TimerStart
        .Result = True
        .TimerEnd
        ' ======================================================================
        
        .TestId = "001-3"
        .TestedProc = "Result and .ResultExpected"
        .TestedType = "Property"
        .Verification = "Result is  F a i l e d  because the result/expected boolean differs"
        .ResultExpected = True
        .TimerStart
        .Result = False
        .TimerEnd
        ' ======================================================================
                
        .TestId = "001-5"
        .TestedProc = "Result and .ResultExpected"
        .TestedType = "Property"
        .Verification = "Result is  F a i l e d  because result/expected files differ"
        .ResultExpected = "ResultExpected," & sFileExpected
        .TempTestItem = sFileExpected
        .TimerStart
        .Result = "Result," & sFileResult
        .TimerEnd
        .TempTestItem = sFileResult
        ' ======================================================================
        If Not .ModeRegression Then .ResultSummaryLog
    End With

xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
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
        .TestedComp = "clsPrivProf"
        .TestedProc = "FileName-Let"
        .TestedType = "Property"
        .TestHeadLine = "Property FileName"
        
        .Verification = "Initialize PP-file"
        .ResultExpected = PrivProfTests.PrivProfFileFullName
        .TimerStart
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        .TimerEnd
        .Result = PrivProf.FileName
        ' ======================================================================
        
        .TestId = "100-2"
        .TestedProc = "Let FileName"
        .TestedType = "Property"
        
        .Verification = "Specifying the full name of an existing Private Profile file"
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName ' continue with current test file
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
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
    
    BoP ErrSrc(PROC)
    '~~ Test preparation
    Prepare
       
    With TestAid
        .TestHeadLine = "Exists method"
        .TestedProc = "Exists"
        .TestedType = "Method"
        .TestId = "110-1"
        
        .Verification = "Section not exists"
        .ResultExpected = False
        .TimerStart
        vResult = PrivProf.Exists(PrivProf.FileName, PrivProfTests.SectionName(7))
        .TimerEnd
        .Result = vResult
        ' ======================================================================
        .TestHeadLine = vbNullString
        
        .TestId = "110-3"
        .TestedProc = "Exists"
        .TestedType = "Method"
        
        .Verification = "Section exists"
        .ResultExpected = True
        .TimerStart
        vResult = PrivProf.Exists(PrivProf.FileName, PrivProfTests.SectionName(8))
        .TimerEnd
        .Result = vResult
        ' ======================================================================
        
        .TestId = "110-4"
        .TestedProc = "Exists"
        .TestedType = "Method"
        
        .Verification = "Value-Name exists"
        .ResultExpected = True
        .TimerStart
        vResult = PrivProf.Exists(PrivProf.FileName, PrivProfTests.SectionName(6), PrivProfTests.ValueName(6, 4))
        .TimerEnd
        .Result = vResult
        ' ======================================================================
        
        .TestId = "110-5"
        .TestedProc = "Exists"
        .TestedType = "Method"
        
        .Verification = "Value-Name not exists"
        .ResultExpected = False
        .TimerStart
        vResult = PrivProf.Exists(PrivProf.FileName, PrivProfTests.SectionName(6), PrivProfTests.ValueName(6, 3))
        .TimerEnd
        .Result = vResult
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

Public Sub Test_120_Property_Value()
' ----------------------------------------------------------------------------
' This test relies on the Value (Let) service.
' ----------------------------------------------------------------------------
    Const PROC = "Test_120_Property_Value"
    
    On Error GoTo eh
    Dim cyValue     As Currency: cyValue = 12345.6789
    
    BoP ErrSrc(PROC)
    Prepare
    
    With TestAid
        .TestHeadLine = "Property Value"
        .TestId = "120-1"
        .TestedProc = "Get Value"
        .TestedType = "Property"
        
        .Verification = "Read non-existing value from a non-existing file"
        .ResultExpected = vbNullString
        .TimerStart
        vResult = PrivProf.Value(name_value:="Any" _
                               , name_section:="Any")
        .TimerEnd
        .Result = vResult
        ' ======================================================================
        .TestHeadLine = vbNullString
        
        .TestId = "120-2"
        .TestedProc = "Get Value"
        .TestedType = "Property"
        
        .Verification = "Read existing value"
        .ResultExpected = PrivProfTests.ValueString(2, 4)
        .TimerStart
        vResult = PrivProf.Value(name_value:=PrivProfTests.ValueName(2, 4) _
                               , name_section:=PrivProfTests.SectionName(2))
        .TimerEnd
        .Result = vResult
        ' ======================================================================
        
        .TestId = "120-3"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        
        .Verification = "Write value"
        .ResultExpected = "Changed value"
        .TimerStart
        PrivProf.Value(name_value:=PrivProfTests.ValueName(4, 2) _
                     , name_section:=PrivProfTests.SectionName(4)) = "Changed value"
        .TimerEnd
        .Result = PrivProf.Value(name_value:=PrivProfTests.ValueName(4, 2) _
                               , name_section:=PrivProfTests.SectionName(4))
        ' ======================================================================
        
        .TestId = "120-4"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        
        .Verification = "Write new value in existing section"
        .ResultExpected = "New value, existing section"
        .TimerStart
        PrivProf.Value(PrivProfTests.ValueName(2, 17) _
                    , PrivProfTests.SectionName(2)) = "New value, existing section"
        .TimerEnd
        .Result = PrivProf.Value(name_value:=PrivProfTests.ValueName(2, 17) _
                              , name_section:=PrivProfTests.SectionName(2))
        ' ======================================================================
        
        .TestId = "120-5"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        
        .Verification = "Write new value in new section"
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.Value(name_value:=PrivProfTests.ValueName(11, 1) _
                     , name_section:=PrivProfTests.SectionName(11)) = "New value, new section"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        ' ======================================================================
        
        .TestId = "120-6"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        
        .Verification = "Change value plus the value and the section comments"
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.Value(name_value:=PrivProfTests.ValueName(11, 1) _
                     , name_section:=PrivProfTests.SectionName(11) _
                      ) = "Changed new value, new section"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        ' ======================================================================
        
        .TestId = "120-7"
        .TestedProc = "Let Value"
        .TestedType = "Property"
        
        .Verification = "Changed again"
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.Value(name_value:=PrivProfTests.ValueName(11, 1) _
                     , name_section:=PrivProfTests.SectionName(11) _
                      ) = "Changed again new value, new section"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
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

Public Sub Test_200_Property_Comments()
' ----------------------------------------------------------------------------
' Let/Get comments (FileHeader, FileFooter, SectionComment, ValueComment)
' ----------------------------------------------------------------------------
    Const PROC = "Test_200_Property_Comments"

    On Error GoTo eh
    Dim sHeader         As String
    Dim sResult         As String
    Dim sValue          As String
    Dim sCommentValue   As String
    Dim sCommentSect    As String
    Dim sFileHeader     As String
    Dim sFileFooter     As String
    Dim cllResultExpectd As Collection
    
    Prepare ' Test preparation
       
    With TestAid
        .TestHeadLine = "Comment/header service (file-, section-, value-)"
        If mTrc.FileFullName <> TestAid.TestFolder & "\Regression.ExecTrace.log" Then
            mTrc.FileFullName = TestAid.TestFolder & "\Test200ExecTrace.log"
            mTrc.Title = "Test: " & .TestHeadLine
        End If
        
        BoP ErrSrc(PROC)
        .TestId = "200-1"
        .TestedProc = "FileHeader-Let"
        .TestedType = "Property"

        .Verification = "Write a file header (not provided delimiter line added by system)"
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        '~~ Note: For the missing file name the property FileName is used
        '~~ and the missing section- and value-name indicate a file header
        PrivProf.FileHeader() = "File Comment Line 1 (the header delimiter is adjusted to the longest header line)" & vbCrLf & _
                                "File Comment Line 2"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        .TestHeadLine = vbNullString ' must not been repeated for each subsequent test
        ' ======================================================================

        .TestId = "200-3"
        .TestedProc = "FileHeader-Get"
        .TestedType = "Property"

        .Verification = "File header read (returned without delimiter line)"
        Set cllResultExpectd = New Collection
        cllResultExpectd.Add "File Comment Line 1 (the header delimiter is adjusted to the longest header line)"
        cllResultExpectd.Add "File Comment Line 2"
        .ResultExpected = cllResultExpectd
        .TimerStart
        vResult = PrivProf.FileHeader()
        .TimerEnd
        .Result = vResult
        ' ======================================================================
        
        .TestId = "200-4"
        .TestedProc = "SectionComment-Let"
        .TestedType = "Property"
        
        .Verification = "Write section comment"
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.SectionComment(PrivProfTests.SectionName(6)) = "Comment Section 06 Line 1" _
                                                     & vbCrLf & "Comment Section 06 Line 2"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        ' ======================================================================
        
        .TestId = "200-5"
        .TestedProc = "SectionComment-Get"
        .TestedType = "Property"
        
        .Verification = "Read section comment"
        .ResultExpected = "Comment Section 06 Line 1" _
               & vbCrLf & "Comment Section 06 Line 2"
        .TimerStart
        vResult = PrivProf.SectionComment(PrivProfTests.SectionName(6))
        .TimerEnd
        .Result = vResult
        ' =====================================================================
        
        .TestId = "200-6"
        .TestedProc = "ValueComment-Let"
        .TestedType = "Property"
        
        .Verification = "Write value comment"
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.ValueComment(PrivProfTests.ValueName(6, 2), PrivProfTests.SectionName(6)) = "Comment Section 06 Value 02 Line 1" _
                                                                                  & vbCrLf & "Comment Section 06 Value 02 Line 2"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        ' ======================================================================
        
        .TestId = "200-7"
        .TestedProc = "Value-Comment-Get"
        .TestedType = "Property"
        
        .Verification = "Read value comment"
        .ResultExpected = "Comment Section 06 Value 02 Line 1" _
               & vbCrLf & "Comment Section 06 Value 02 Line 2"
        .TimerStart
        vResult = PrivProf.ValueComment(PrivProfTests.ValueName(6, 2), PrivProfTests.SectionName(6))
        .TimerEnd
        .Result = vResult
        ' ======================================================================
            
    End With

xt: EoP ErrSrc(PROC)
    TestAid.CleanUp
    If Not mErH.Regression Then mTrc.Dsply
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
        .TestHeadLine = "SectionNames service"
        .TestId = "300-1"
        .TestedProc = "SectionNames"
        .TestedType = "Function"
        
        .Verification = "Get all section names in a Dictionary"
        .ResultExpected = 5
        .TimerStart
        vResult = PrivProf.SectionNames().Count
        .TimerEnd
        .Result = vResult
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
        .TestHeadLine = "ValueNames service"
        .TestId = "400-1"
        .TestedProc = "ValueNames"
        .TestedType = "Function"
        .Verification = "Get all value names of all sections in a Dictionary"
        .ResultExpected = 40
        .TimerStart
        Set dct = PrivProf.ValueNames()
        .TimerEnd
        .Result = dct.Count
        ' ======================================================================
        .TestHeadLine = vbNullString
    
        .TestId = "400-2"
        .TestedProc = "ValueNames"
        .TestedType = "Function"
        .Verification = "Get all value names of a certain section in a Dictionary"
        .ResultExpected = 8
        .TimerStart
        vResult = PrivProf.ValueNames(, PrivProfTests.SectionName(6)).Count
        .TimerEnd
        .Result = vResult
        ' ======================================================================
      
        .TestId = "400-3"
        Prepare 2
        .TestedProc = "ValueNames"
        .TestedType = "Function"
        .Verification = "Get all value names of all sections in a Dictionary"
        .ResultExpected = 6
        .TimerStart
        vResult = PrivProf.ValueNames().Count
        .TimerEnd
        .Result = vResult
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
        .TestHeadLine = "ValueNameRename service"
        .TestId = "410-1"
        .TestedProc = "ValueNameRename"
        .TestedType = "Method"
        .Verification = "Rename a value name in each section."
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.ValueNameRename PrivProfTests.ValueName(2, 2), "Renamed_" & PrivProfTests.ValueName(2, 2)
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
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
        .TestHeadLine = "Remove service"
        PrivProf.ValueComment(PrivProfTests.SectionName(6), PrivProfTests.ValueName(6, 4)) = "Comment value 06-04"
        .TestId = "500-1"
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        
        .Verification = "Remove a value from a section including its comments."
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.ValueRemove PrivProfTests.ValueName(6, 4), PrivProfTests.SectionName(6)
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        ' ======================================================================
        .TestHeadLine = vbNullString
        
        PrivProf.SectionComment(PrivProfTests.SectionName(6)) = "Comment section 06"
        .TestId = "500-2"
        .TestedProc = "SectionRemove"
        .TestedType = "Method"
        
        .Verification = "Removes a section including its comments."
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.SectionRemove PrivProfTests.SectionName(6)
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        ' ======================================================================
    
        .TestId = "500-3"
        Prepare 2
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        
        .Verification = "Remove 2 names in 2 sections."
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.ValueRemove name_value:="Last_Modified_AtDateTime,Last_Modified_InWbkFullName", name_section:="clsLog,clsQ"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        ' ======================================================================
    
        .TestId = "500-4"
        Prepare 2
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        
        .Verification = "Removes all values in one section which removes the section."
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.ValueRemove name_value:="ExportFileExtention" & _
                                         ",Last_Modified_AtDateTime" & _
                                         ",Last_Modified_InWbkFullName" & _
                                         ",Last_Modified_InWbkName" & _
                                         ",LastModExpFileFullNameOrigin" _
                           , name_section:="clsQ"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        ' ======================================================================
       
        .TestId = "500-5"
        Prepare 2
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        
        .Verification = "Remove all values in all sections - file is removed."
        sFile = PrivProfTests.PrivProfFile
        .ResultExpected = False
        .TimerStart
        PrivProf.ValueRemove name_value:="ExportFileExtention" & _
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
   
    Prepare 0, False
    BoP ErrSrc(PROC)
    
    With TestAid
        .TestHeadLine = "A PrivateProfile file life-cycle"
        .TestId = "600-1"
        .TestedProc = "FileHeader-Let, FileFooter-Let"
        .TestedType = "Property"
        .Verification = "New file, header and footer writing postponed"
        '~~ Header and/or footer writing to a not yet active file will be
        '~~ postponed until at least one user-value had been written.
        If FSo.FileExists(PrivProfTests.PrivProfFileFullName) Then .FSo.DeleteFile PrivProfTests.PrivProfFileFullName
        Set PrivProf = New clsPrivProf
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        .ResultExpected = True
        .TimerStart
        PrivProf.FileFooter() = "File Footer Line 1 (the delimiter below is adjusted to the longest comment)" & vbCrLf & _
                                "File Footer Line 2"
        PrivProf.FileHeader() = "File Comment Line 1 (the delimiter below is adjusted to the longest comment)" & vbCrLf & _
                                "File Comment Line 2"
        .TimerEnd
        .Result = Not FSo.FileExists(PrivProfTests.PrivProfFileFullName)
        ' ======================================================================
        
        .TestId = "600-2"
        .TestedProc = "Value-Let"
        .TestedType = "Property"
        .Verification = "First value also writes postponed header and footer"
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        .TimerStart
        PrivProf.Value(name_value:="Any-Value-Name" _
                    , name_section:="Any-Section-Name" _
                    , name_file:=PrivProfTests.PrivProfFileFullName _
                      ) = "Any-Value"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        ' ======================================================================
        
        .TestId = "600-3"
        .TestedProc = "ValueRemove"
        .TestedType = "Method"
        .Verification = "Removes the only value in the only section removes file"
        '~~ Removing the only value in the only section ends with no file
        '~~ Note: This is consequent since there is no Private Profile file without
        '~~       at least on section with one value
        Set PrivProf = New clsPrivProf
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        PrivProf.FileName = PrivProfTests.PrivProfFileFullName
        .ResultExpected = True
        .TimerStart
        mErH.Asserted AppErr(1) ' effective only when mErH.Regression = True
        PrivProf.ValueRemove name_value:="Any-Value-Name" _
                           , name_section:="Any-Section-Name"
        .TimerEnd
        mErH.Asserted ' reset to none
        .Result = Not FSo.FileExists(PrivProfTests.PrivProfFileFullName)
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
        .TestHeadLine = "Names housekeeping"
        .TestId = "700-1"
        .TestedProc = "HouskeepingNames"
        .TestedType = "Method"
        
        .Verification = "One Value-name change in two sections"
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.HskpngNames PrivProf.FileName, "clsLog:clsQ:Last_Modified_AtDateTime>Last_Modified_UTC_AtDateTime"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        ' ======================================================================
                
        .TestId = "700-2"
        Prepare 2   ' uses a ready for test file copied from a backup
        
        .Verification = "Two value-name changes in all sections"
        .ResultExpected = PrivProfTests.ExpectedTestResultFile(.TestId)
        .TimerStart
        PrivProf.HskpngNames PrivProf.FileName, "Last_Modified_AtDateTime>Last_Modified_UTC_AtDateTime" _
                                              , "LastModExpFileFullNameOrigin>Last_Modified_ExpFileFullNameOrigin"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
        Set fleTestResult = PrivProfTests.PrivProfFile
        ' ======================================================================
        
        .TestId = "700-3"
        .Verification = "Two value-name changes in all sections (any subsequent once done)"
        .ResultExpected = fleTestResult
        .TimerStart
        PrivProf.HskpngNames fleTestResult.Path, "Last_Modified_AtDateTime>Last_Modified_UTC_AtDateTime" _
                                               , "LastModExpFileFullNameOrigin>Last_Modified_ExpFileFullNameOrigin"
        .TimerEnd
        .Result = PrivProfTests.PrivProfFile
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

