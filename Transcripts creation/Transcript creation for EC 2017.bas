Attribute VB_Name = "Module1"
Public Sub add_subject(wd As Word.Application, wdDOC As Word.Document, sh As Worksheet, iRow As Long, filePath_transcripts As String, exam_mode As String, current_column As String, subject_name As String)

    'subject_credits
    wd.Selection.GoTo what:=wdGoToBookmark, Name:=subject_name & "_credits"
    If sh.Range(current_column & 3).Value = "0" Then
        'then credit cell is empty
        Else
        wd.Selection.TypeText Text:=sh.Range(current_column & 3).Value
    End If
    
    'subject_mode
    wd.Selection.GoTo what:=wdGoToBookmark, Name:=subject_name & "_mode"
    If sh.Range(current_column & 2).Value = "экз" Then
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
    Dim subject_name As String 'Variable to make bookmarks in function

    
    'Start word and a new Doc
    Set wd = New Word.Application
    Set sh = ThisWorkbook.Sheets("оценки ЭФ 24.09.21") 'there could be problems because of russian sheet name

    'first, copy template, i don't want to delete bookmarks everywhere
    'i don't know this doesn't work
    'FileCopy ThisWorkbook.Path & "\Transcript template with bookmarks and red.docx", _
    'ThisWorkbook.Path & "temp_copy_Transcript template.docx"
    
    iRow = 4 'where student data begins
    filePath_transcripts = ThisWorkbook.Path & "\student transcripts"
    
    
    'loop while the B cell of irow is not blank
    Do While sh.Range("B" & iRow).Value <> ""
        'opening the word template
        Set wdDOC = wd.Documents.Open(ThisWorkbook.Path & "\copy_template.docx")
        'if I want to see how app works try to delete this line
        wd.Visible = False
        
        '##########################################
        '##########################################
        'Code to insert a value from excel to specific bookmark in word doc
        
        'MSU_student_id
        wd.Selection.GoTo what:=wdGoToBookmark, Name:="MSU_student_id"
        wd.Selection.TypeText Text:=sh.Range("D" & iRow).Value
        
        'GPA
        wd.Selection.GoTo what:=wdGoToBookmark, Name:="GPA"
        wd.Selection.TypeText Text:=Round(sh.Range("BC" & iRow).Value, 2)
        
        '##########################################
        'Russian_language_1
        current_column = "E"
        subject_name = "Russian_language_1"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_spec_1
        current_column = "F"
        subject_name = "Russian_spec_1"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Practical_Russian_1
        current_column = "G"
        subject_name = "Practical_Russian_1"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Modern_history_of_China
        current_column = "H"
        subject_name = "Modern_history_of_China"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Thought_1
        current_column = "I"
        subject_name = "Thought_1"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Fund_life_safety_1
        current_column = "J"
        subject_name = "Fund_life_safety_1"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Physical_training
        current_column = "K"
        subject_name = "Physical_training"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        '##########################################
        '##########################################
        'Intro_specialty_ex
        current_column = "M"
        subject_name = "Intro_specialty_ex"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        
        'Intro_specialty_zach
        'dont add credits, all other same
        current_column = "N"
        subject_name = "Intro_specialty_zach"
        
            'subject_mode
        wd.Selection.GoTo what:=wdGoToBookmark, Name:=subject_name & "_mode"
        If sh.Range(current_column & 2).Value = "экз" Then
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
        
        
        '##########################################
        '##########################################
        '##########################################
        'Mao_thought_2
        current_column = "N"
        subject_name = "Mao_thought_2"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Basic_Marxism_2
        current_column = "O"
        subject_name = "Basic_Marxism_2"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_language_2
        current_column = "P"
        subject_name = "Russian_language_2"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_2_spec
        current_column = "Q"
        subject_name = "Russian_2_spec"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_2_prof
        current_column = "R"
        subject_name = "Russian_2_prof"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Prac_information_2
        current_column = "S"
        subject_name = "Prac_information_2"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Elective_physical_2
        current_column = "T"
        subject_name = "Elective_physical_2"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        '##########################################
        'Mathan
        current_column = "U"
        subject_name = "Mathan"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Linal
        current_column = "V"
        subject_name = "Linal"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Macroec_I
        current_column = "W"
        subject_name = "Macroec_I"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Microec_I
        current_column = "X"
        subject_name = "Microec_I"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Statistics
        current_column = "Y"
        subject_name = "Statistics"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_language_3
        current_column = "Z"
        subject_name = "Russian_language_3"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_3_prof
        current_column = "AA"
        subject_name = "Russian_3_prof"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Elective_physical_3
        current_column = "AB"
        subject_name = "Elective_physical_3"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        '##########################################
        'Microec_II
        current_column = "AC"
        subject_name = "Microec_II"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Opt_solution
        current_column = "AD"
        subject_name = "Opt_solution"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'El_higher_math
        current_column = "AE"
        subject_name = "El_higher_math"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Demography
        current_column = "AF"
        subject_name = "Demography"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'History
        current_column = "AG"
        subject_name = "History"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Philosophy
        current_column = "AH"
        subject_name = "Philosophy"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Ec_informatics
        current_column = "AI"
        subject_name = "Ec_informatics"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_language_4
        current_column = "AJ"
        subject_name = "Russian_language_4"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_4_prof
        current_column = "AK"
        subject_name = "Russian_4_prof"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Elective_physical_4
        current_column = "AL"
        subject_name = "Elective_physical_4"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        '##########################################
        'Probability_th
        current_column = "AM"
        subject_name = "Probability_th"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Macroec_II
        current_column = "AN"
        subject_name = "Macroec_II"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'International_ec
        current_column = "AO"
        subject_name = "International_ec"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Game_theory
        current_column = "AP"
        subject_name = "Game_theory"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Prac_tr_acad
        current_column = "AR"
        subject_name = "Prac_tr_acad"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Fin_theory
        current_column = "AQ"
        subject_name = "Fin_theory"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_language_5
        current_column = "AS"
        subject_name = "Russian_language_5"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_5_prof
        current_column = "AT"
        subject_name = "Russian_5_prof"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        '##########################################
        'Econometrics
        current_column = "AU"
        subject_name = "Econometrics"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Fin_markets
        current_column = "AV"
        subject_name = "Fin_markets"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Intern_fin_acc
        current_column = "AW"
        subject_name = "Intern_fin_acc"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Labor_economics
        current_column = "AX"
        subject_name = "Labor_economics"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Fund_entrepr
        current_column = "AY"
        subject_name = "Fund_entrepr"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Prac_tr_Internship
        current_column = "AZ"
        subject_name = "Prac_tr_Internship"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_language_6
        current_column = "BA"
        subject_name = "Russian_language_6"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        'Russian_6_prof
        current_column = "BB"
        subject_name = "Russian_6_prof"
        add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name
        '##########################################
        
        
        
        
        
        
        
        '##########################################
        
        'save word file with a student name
        wdDOC.SaveAs2 (filePath_transcripts & "\" & sh.Range("C" & iRow).Value) & ".docx"
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


