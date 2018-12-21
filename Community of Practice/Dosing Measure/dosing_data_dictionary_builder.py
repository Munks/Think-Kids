import pandas as pd
import glob as glob
from redcap import Project
import sys
sys.path.append(r'C:\Python Programs\Think-Kids_Private')
import tokens as tk

path = r'C:\Python Programs\Think-Kids\Community of Practice\Dosing Measure'
text = pd.read_excel(path + r'\Dosing Measure Text.xlsx')
path = path + r'/Staff and Student Lists'

tk.dosing_measure_api()

api_data_dictionary = Project(tk.api_url, tk.api_token)
api_data_dictionary = api_data_dictionary.export_metadata(format='df')
blank_data_dictionary = api_data_dictionary[:2]
blank_data_dictionary = blank_data_dictionary.reset_index()


def new_data_dictionary_row(
        variable_name='',
        form_name='dosing_measure',
        section_header='',
        field_type='',
        field_label='',
        choices='',
        field_notes='',
        text_validation='',
        text_validation_min='',
        text_validation_max='',
        identifier='',
        logic='',
        required='',
        custom_alignment='',
        question_number='',
        matrix_name='',
        matrix_ranking='',
        field_annotation=''):
    df = pd.Series(
        [variable_name, form_name, section_header, field_type,
         field_label, choices, field_notes, text_validation,
         text_validation_min, text_validation_max, identifier,
         logic, required, custom_alignment, question_number,
         matrix_name, matrix_ranking, field_annotation],
        index=list(blank_data_dictionary.columns))
    return df


def grades(df, grade, names):
    for i in df.index:
        name = df.loc[i, 'Name']
        if str(df.loc[i, 'Grade']) == grade:
            names.append(name)
    names = names.sort()
    return names


def teacher_grade_level(grade, data, grade_list):
    staff_list = []
    grades(df_staff_list, grade, staff_list)
    staff = []
    for count, staff_name in enumerate(staff_list):
        if staff:
            staff = staff + ' | ' + str(count) + ', ' + staff_name
        else:
            staff = str(count) + ', ' + staff_name
    if staff:
        grade_list.append(grade)
        df = new_data_dictionary_row(
            variable_name=school_abreviation + 'teachers_grade_' + grade,
            field_type='dropdown',
            field_label='What is your name?',
            choices=staff,
            logic=school_abrev_logic + "'" + grade + "'",
            text_validation='autocomplete',
            required='y')
        data = dosing_dictionary.append(df, ignore_index=True)
        return data
    else:
        return data


def student_dosing_row(student_name, grade_word, grade_num, value):
    new_data_dictionary_row(
        variable_name=school_abreviation + student_name,
        section_header=(text.iloc[0][0] + grade_word + text.iloc[0][1]),
        field_type='radio',
        field_label='',
        choices=text.iloc[1][0],
        logic=text.iloc[2][0] + grade_num + text.iloc[2][1],
        matrix_name=value[:5].lower() + '_students'
    )


grades_dictionary = {'0': 'Kindergarten',
                     '1': 'First Grade',
                     '2': 'Second Grade',
                     '3': 'Third Grade',
                     '4': 'Fourth Grade',
                     '5': 'Fifth Grade',
                     '6': 'Sixth Grade',
                     '7': 'Seventh Grade',
                     '8': 'Eighth Grade',
                     '9': 'Ninth Grade',
                     '10': 'Tenth Grade',
                     '11': 'Eleventh Grade',
                     '12': 'Twelth Grade',
                     '20': 'Multiple Grades'}


folder = glob.glob(path + r'/*')

dosing_dictionary = blank_data_dictionary

choices_text = 'select_choices_or_calculations'
variable_text = 'field_name'

schools_dict = {}

school_choices = str(dosing_dictionary.iloc[1][choices_text])

if ' | ' in school_choices:
    content = school_choices.split(' | ')
    for schools in content:
        school = schools.split(', ', 1)
        schools_dict[int(school[0])] = school[1]
else:
    school = school_choices.split(', ', 1)
    schools_dict[school[0]] = school[1]

for folders in folder:
    school_folders = glob.glob(folders + r'\\*')
    for count, school_file in enumerate(school_folders):
        school_name = school_file.split('\\')[-2]
        schools_dict = {value: int(key) for key, value in schools_dict.items()}
        new_value = max(value for value in schools_dict.values()) + 1
        if school_name not in schools_dict.keys():
            dosing_dictionary.iloc[1][choices_text] = dosing_dictionary.iloc[1][choices_text] + ' | ' + str(new_value) + ', ' + school_name
        excel_file = pd.ExcelFile(school_file)
        df_staff_list = pd.read_excel(excel_file, sheet_name='Staff')
        df_student_list = pd.read_excel(excel_file, sheet_name='Students')
        school_abreviation = df_staff_list.iloc[1]['School'] + '_'
        school_abrev_logic = '[' + school_abreviation + 'grades] = '
        dosing_dictionary = dosing_dictionary.append(
            new_data_dictionary_row(
                variable_name=school_abreviation + 'grades',
                field_type='dropdown',
                section_header='<h4>' + school_name + '</h4>',
                field_label='Which grade do you teach?',
                choices=school_abreviation + 'grade',
                text_validation='autocomplete',
                logic="[which_school] = '" + str(
                    schools_dict[school_name]) + "'",
                required='y'), ignore_index=True)
        grade_list = []
        for key in grades_dictionary.keys():
            dosing_dictionary = teacher_grade_level(
                key, dosing_dictionary, grade_list)
        school_grades = []
        for grade in grade_list:
            if school_grades:
                school_grades = school_grades + ' | ' + \
                    grade + ', ' + grades_dictionary[grade]
            else:
                school_grades = grade + ', ' + grades_dictionary[grade]
        for i in dosing_dictionary.index:
            if dosing_dictionary.iloc[i][variable_text] == (
                    school_abreviation + 'grades'):
                dosing_dictionary.iloc[i][choices_text] = school_grades
        dosing_dictionary = dosing_dictionary.append(
            new_data_dictionary_row(
                variable_name='need_to_report',
                section_header='Students had behaviors this week that I need to report.',
                field_type='radio',
                choices='1, Yes | 2, No',
                required='y'), ignore_index=True)
        student_df_by_grade = []
        for key, value in grades_dictionary.items():
            students = []
            grades(df_student_list, key, students)
            first_loop = True
            for student in students:
                if first_loop is True:
                    header = text.iloc[0][0] + value + text.iloc[0][1]
                    first_loop = False
                else:
                    header = ''
                student_name_variable = (
                    school_abreviation + student.replace(' ', '_').lower())
                student_name_variable = student_name_variable.replace("'", '')
                df = new_data_dictionary_row(
                    variable_name=student_name_variable,
                    field_type='radio',
                    section_header=header,
                    field_label=student,
                    choices=text.iloc[1][0],
                    logic=(
                        "[need_to_report] = '1' and (" +
                        school_abrev_logic +
                        "'" +
                        key +
                        "' or " +
                        school_abrev_logic +
                        "'20')"
                    ),
                    matrix_name=value[:5].lower() + '_students')
                student_df_by_grade.append(df)
        student_df_by_grade = pd.DataFrame(student_df_by_grade)
        dosing_dictionary = dosing_dictionary.append(
            student_df_by_grade, ignore_index=True, sort=False)
        text_question_logic = []
        for student in student_df_by_grade['field_label']:
            variable = school_abreviation + student.replace(
                ' ', '_').lower()
            variable = variable.replace("'", '')
            text_question_logic.append('[' + variable + "_q6] > '0'")
        text_question_logic = ' or '.join(text_question_logic)
        for i in range(6):
            print(i)
            first_loop = True
            for student in student_df_by_grade['field_label']:
                variable = school_abreviation + student.replace(
                    ' ', '_').lower()
                variable = variable.replace("'", '')
                logic = "[" + variable + "] > '0'"
                if first_loop is True:
                    header = text.iloc[i + 4][0]
                    first_loop = False
                else:
                    header = ''
                dosing_dictionary = dosing_dictionary.append(
                    new_data_dictionary_row(
                        variable_name=variable + '_q' + str(i + 1),
                        section_header=header,
                        field_label=student,
                        field_type='radio',
                        choices=text.iloc[1][1],
                        logic=logic,
                        matrix_name='question_' + str(i + 1)),
                    ignore_index=True)
        dosing_dictionary = dosing_dictionary.append(
            new_data_dictionary_row(
                variable_name='q6_text_box',
                field_label=text.iloc[9][0],
                field_type='text',
                logic=text_question_logic),
            ignore_index=True)

dosing_dictionary.to_csv(path + '/Dosing Measure Data Dict.csv', index=False)
