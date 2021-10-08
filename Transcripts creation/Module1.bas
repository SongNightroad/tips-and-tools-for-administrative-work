Attribute VB_Name = "Module1"
Public Sub add_subject(wd As Word.Application, wdDOC As Word.Document, sh As Worksheet, iRow As Long, filePath_transcripts As String, exam_mode As String, current_column As String, subject_name As String)

    Dim number_string_mode_exam As Integer
    Dim number_string_credit As Integer
    
    number_string_mode_exam = 2
    number_string_credit = 3

    'subject_credits
    wd.Selection.GoTo what:=wdGoToBookmark, Name:=subject_name & "_credits"
    If sh.Range(current_column & number_string_credit).Value = "0" Then
        'then credit cell is empty
        Else
        wd.Selection.TypeText Text:=sh.Range(current_column & number_string_credit).Value
    End If
    
    'subject_mode
    wd.Selection.GoTo what:=wdGoToBookmark, Name:=subject_name & "_mode"
    If sh.Range(current_column & number_string_mode_exam).Value = "экз" Then
        wd.Selection.TypeText Text:="Exam"
        
        'grades for exams
    'subject_Academic_results
        wd.Selection.GoTo what:=wdGoToBookmark, Name:=subject_name & "_Academic_results"
        If sh.Range(current_column & iRow).Value = "5" Then
            wd.Selection.TypeText Text:="5"
            ElseIf sh.Range(current_column & iRow).Value = "4" Then
            wd.Selection.TypeText Text:="4"
            ElseIf sh.Range(current_column & iRow).Value = "3" Then
            wd.Selection.TypeText Text:="3"
            Else
            wd.Selection.TypeText Text:="-"
        End If
        
    'subject_Grades
        wd.Selection.GoTo what:=wdGoToBookmark, Name:=subject_name & "_Grades"
        If sh.Range(current_column & iRow).Value = "5" Then
            wd.Selection.TypeText Text:="Excellent"
            ElseIf sh.Range(current_column & iRow).Value = "4" Then
            wd.Selection.TypeText Text:="Good"
            ElseIf sh.Range(current_column & iRow).Value = "3" Then
            wd.Selection.TypeText Text:="Satisfactory"
            Else
            wd.Selection.TypeText Text:="Not passed"
        End If
        
    Else
        'grades for pass test
        wd.Selection.TypeText Text:="Pass/Fail exam"
        
    'subject_Academic_results
        wd.Selection.GoTo what:=wdGoToBookmark, Name:=subject_name & "_Academic_results"
        If sh.Range(current_column & iRow).Value = "5" Then
            wd.Selection.TypeText Text:="Passed"
            ElseIf sh.Range(current_column & iRow).Value = "4" Then
            wd.Selection.TypeText Text:="Passed"
            ElseIf sh.Range(current_column & iRow).Value = "3" Then
            wd.Selection.TypeText Text:="Passed"
            Else
            wd.Selection.TypeText Text:="Not passed"
        End If
        
    'subject_Grades
        wd.Selection.GoTo what:=wdGoToBookmark, Name:=subject_name & "_Grades"
        If sh.Range(current_column & iRow).Value = "5" Then
            wd.Selection.TypeText Text:="Passed"
            ElseIf sh.Range(current_column & iRow).Value = "4" Then
            wd.Selection.TypeText Text:="Passed"
            ElseIf sh.Range(current_column & iRow).Value = "3" Then
            wd.Selection.TypeText Text:="Passed"
            Else
            wd.Selection.TypeText Text:="Not passed"
        End If
    End If


End Sub


Sub SendtoWord()
    'Declare variables
    Dim wd As Word.Application
    Dim wdDOC As Word.Document
    Dim sh As Worksheet 'Variable to refer to the sheet of excel
    Dim iRow As Long 'Variable to hold the starting row and loop through all records in the table
    Dim filePath_transcripts As String
    Dim exam_mode As String
    Dim current_column As String 'Variable to refer to subjects column
    Dim current_column_int As Integer
    Dim subject_name As String 'Variable to make bookmarks in function
    Dim transcripts_template_ch_rus As String 'for making name with _ch and _rus
    
    'block of hard values
    Dim grades_sheet_name As String
    Dim gpa_sheet_name As String
    Dim faculty As String
    Dim year As Integer
    Dim num_chinese_st As Integer
    Dim num_russian_st As Integer
    Dim start_line As Integer 'where student data begins
    Dim transcripts_folder As String
    Dim transcripts_template_name As String
    Dim xRg As Range 'range of grades for counting gpa on gpa list
    
    grades_sheet_name = "оценки"
    gpa_sheet_name = "оценки GPA"
    faculty = "ВМК"
    year = 2018
    num_chinese_st = 27
    num_russian_st = 5 'not important, now we define as all not chinese
    start_line = 4
    transcripts_folder = "Транскрипты без меты"
    transcripts_template_name = "copy_template"
    Set xRg = Range("A1:CL36")
    
    'recalculating GPA
    Set sh = ThisWorkbook.Sheets(gpa_sheet_name)
    Dim rg As Range
    For Each rg In xRg
        With rg
            Select Case .Value
                Case Is = 1
                    .Value = 0
                Case Is = -1
                    .Value = 0
            End Select
        End With
    Next
    
    'Start word and a new Doc
    Set wd = New Word.Application
    Set sh = ThisWorkbook.Sheets(grades_sheet_name)
    
    iRow = start_line
    filePath_transcripts = ThisWorkbook.Path & "\" & transcripts_folder
    
    
    Do While sh.Range("B" & iRow).Value <> ""
               
        'set template define on chinese\russian student
        If iRow < start_line + num_chinese_st Then
            transcripts_template_ch_rus = transcripts_template_name & "_ch"
        Else
            transcripts_template_ch_rus = transcripts_template_name & "_rus"
        End If
                        
        'opening the word template
        Set wdDOC = wd.Documents.Open(ThisWorkbook.Path & "\" & transcripts_template_ch_rus & ".docx")
        'if I want to see how app works try to delete this line
        wd.Visible = False
        
        '##########################################
        '##########################################
        'Code to insert a value from excel to specific bookmark in word doc
        
        'MSU_student_id
        wd.Selection.GoTo what:=wdGoToBookmark, Name:="MSU_student_id"
        wd.Selection.TypeText Text:=sh.Range("E" & iRow).Value
        
        'GPA
        wd.Selection.GoTo what:=wdGoToBookmark, Name:="GPA"
        wd.Selection.TypeText Text:=Round(sh.Range("BL" & iRow).Value, 2)
        
        '##########################################
        '##########################################
        'Russian_language_1
        current_column = "F"
        subject_name = "Russian_language_1"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_spec_1
        current_column = "G"
        subject_name = "Russian_spec_1"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Practical_Russian_1
        current_column = "H"
        subject_name = "Practical_Russian_1"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Modern_history_of_China
        current_column = "I"
        subject_name = "Modern_history_of_China"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Thought_1
        current_column = "J"
        subject_name = "Thought_1"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Fund_life_safety_1
        current_column = "K"
        subject_name = "Fund_life_safety_1"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Physical_training
        current_column = "L"
        subject_name = "Physical_training"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Intro_specialty_ex
        current_column = "N"
        subject_name = "Intro_specialty_ex"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Mao_thought_2
        current_column = "P"
        subject_name = "Mao_thought_2"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Basic_Marxism_2
        current_column = "Q"
        subject_name = "Basic_Marxism_2"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_language_2
        current_column = "R"
        subject_name = "Russian_language_2"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_2_spec
        current_column = "S"
        subject_name = "Russian_2_spec"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Prac_information_2
        current_column = "O"
        subject_name = "Prac_information_2"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Elective_physical_2
        current_column = "T"
        subject_name = "Elective_physical_2"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Mathan_3
        current_column = "V"
        subject_name = "Mathan_3"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Algebra_geometry_3
        current_column = "X"
        subject_name = "Algebra_geometry_3"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Alg_comput_arch_3
        current_column = "Y"
        subject_name = "Alg_comput_arch_3"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Computer_prac_3
        current_column = "Z"
        subject_name = "Computer_prac_3"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_language_3
        current_column = "AA"
        subject_name = "Russian_language_3"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_3_prof
        current_column = "AB"
        subject_name = "Russian_3_prof"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Elective_physical_3
        current_column = "AC"
        subject_name = "Elective_physical_3"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Mathan_4
        current_column = "AE"
        subject_name = "Mathan_4"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Algebra_geometry_4
        current_column = "AG"
        subject_name = "Algebra_geometry_4"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Alg_comput_arch_4
        current_column = "AH"
        subject_name = "Alg_comput_arch_4"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Discret_math_4
        current_column = "AJ"
        subject_name = "Discret_math_4"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Computer_prac_4
        current_column = "AK"
        subject_name = "Computer_prac_4"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_language_4
        current_column = "AL"
        subject_name = "Russian_language_4"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_4_prof
        current_column = "AM"
        subject_name = "Russian_4_prof"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Elective_physical_4
        current_column = "AN"
        subject_name = "Elective_physical_4"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Ord_diff_equat_5
        current_column = "AP"
        subject_name = "Ord_diff_equat_5"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Oper_systems_5
        current_column = "AQ"
        subject_name = "Oper_systems_5"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Mathan_5
        current_column = "AS"
        subject_name = "Mathan_5"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Num_methods_5
        current_column = "AT"
        subject_name = "Num_methods_5"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Computer_prac_5
        current_column = "AU"
        subject_name = "Computer_prac_5"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Databases_5
        current_column = "AV"
        subject_name = "Databases_5"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Fund_obj_or_pr_5
        current_column = "AW"
        subject_name = "Fund_obj_or_pr_5"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Functional_pr_5
        current_column = "AX"
        subject_name = "Functional_pr_5"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_language_5
        current_column = "AY"
        subject_name = "Russian_language_5"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_5_prof
        current_column = "AZ"
        subject_name = "Russian_5_prof"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Mathan_6
        current_column = "BB"
        subject_name = "Mathan_6"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Eq_math_phys_6
        current_column = "BC"
        subject_name = "Eq_math_phys_6"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Num_methods_6
        current_column = "BD"
        subject_name = "Num_methods_6"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Funct_analysis_6
        current_column = "BE"
        subject_name = "Funct_analysis_6"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Prog_languag_6
        current_column = "BF"
        subject_name = "Prog_languag_6"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Comp_meth_linal_6
        current_column = "BG"
        subject_name = "Comp_meth_linal_6"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Comp_prac_econ_6
        current_column = "BH"
        subject_name = "Comp_prac_econ_6"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Spec_seminar_6
        current_column = "BI"
        subject_name = "Spec_seminar_6"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_language_6
        current_column = "BJ"
        subject_name = "Russian_language_6"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_6_prof
        current_column = "BK"
        subject_name = "Russian_6_prof"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        
        
        'save word file with a student name
        wdDOC.SaveAs2 (filePath_transcripts & "\" & year & "_" & faculty & "_" & sh.Range("C" & iRow).Value) & ".docx"
        'close word file
        wdDOC.Close
        Set wdDOC = Nothing
        

        iRow = iRow + 1
    Loop
    


    'close word app
    wd.Quit
    Set wd = Nothing
    
    MsgBox ("Transkripts have been created succesfully!")

End Sub


