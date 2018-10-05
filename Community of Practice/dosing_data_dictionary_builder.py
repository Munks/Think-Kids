import pandas as pd
import glob as glob


# User Interface
#   Think:Kids
print('\n\n')
print("\t\t* * * * * * *   *    *      *")
print("\t\t      *        * *   *    *")
print("\t\t      *         *    *  *")
print("\t\t      *              **")
print("\t\t      *         *    *  *")
print("\t\t      *        * *   *    *")
print("\t\t      *         *    *      *")
print("\n\t\t\tDosing Measures!!\n\n")


drive = r'\\Cifs2\thinkkid$'

folders = r'\Research\Chris\Community of Practice\Staff and Student Lists'

path = drive + folders

dosing_dictionary_template = pd.read_excel(
    path + r'\Excel Templates\Blank Data Dictionary.xlsx')

folder = glob.glob(path + r'\*')

schools = {}

for number, file in enumerate(folder):
    schools[number] = file[78:]

print('Which school are you looking to make a dictionary for?\n')

for key, value in schools.items():
    print(str(key) + '\t' + value[1:])

school_folder = input('\nPlease enter the number related to the school\n\n\t')

school_folder = schools[int(school_folder)]

school = school_folder[1:]

path = path + school_folder

df_staff = pd.read_excel(path + r'/Current Staff List.xlsx')
df_students = pd.read_excel(path + r'/Current Student List.xlsx')

form_name = dosing_dictionary_template.iloc[0]["Form Name"]

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


def create_new_line(variable, form_name, section_header, field_type,
                    field_label, choices_calculations_sliderlabels,
                    field_notes, validation, logic, required, matrix_name):
    df = pd.Series([variable, form_name, section_header, field_type,
                    field_label, choices_calculations_sliderlabels,
                    field_notes, validation, logic, required, matrix_name],
                   index=['Variable / Field Name', 'Form Name',
                          'Section Header', 'Field Type', 'Field Label',
                          'Choices, Calculations, OR Slider Labels',
                          'Field Note', 'Text Validation',
                          'Branching Logic (Show field only if...)',
                          'Required Field?', 'Matrix Group Name'])
    return df


def grades(df, grade, names):
    for i in df.index:
        name = df.loc[i, 'name']
        if str(df.loc[i, 'grade']) == grade:
            names.append(name)
    names = names.sort()
    return names


def teacher_grade_level(grade, data, grade_list):
    staff_list = []
    grades(df_staff, grade, staff_list)
    staff = []
    for count, staff_name in enumerate(staff_list):
        if staff:
            staff = staff + ' | ' + str(count) + ', ' + staff_name
        else:
            staff = str(count) + ', ' + staff_name
    if staff:
        grade_list.append(grade)
        df = create_new_line('teacher_grade_' + grade, form_name, '',
                             'dropdown', 'What is your name?', staff, '', '',
                             "[grades] = '" + grade + "'", 'y', '')
        data = dosing_dictionary.append(df, ignore_index=True)
        return data
    else:
        return data


df = create_new_line('grades', form_name, '', 'dropdown',
                     'Which grade do you teach?', '', '', '',
                     '', 'y', '')


dosing_dictionary = dosing_dictionary_template.append(df, ignore_index=True)
dosing_dictionary.loc[1, "Section Header"] = '<h4>' + school + '</h4>'

grade_list = []
for a in grades_dictionary.keys():
    dosing_dictionary = teacher_grade_level(a, dosing_dictionary, grade_list)

school_grades = []
for grade in grade_list:
    if school_grades:
        school_grades = school_grades + ' | ' + \
            grade + ', ' + grades_dictionary[grade]
    else:
        school_grades = grade + ', ' + grades_dictionary[grade]

dosing_dictionary.loc[
    1, 'Choices, Calculations, OR Slider Labels'] = school_grades

student_df_by_grade = []

for key, value in grades_dictionary.items():
    students = []
    grades(df_students, key, students)
    if students:
        student_name_variable = students[0].replace(' ', '_').lower()
        df = create_new_line(student_name_variable, form_name,
                             'For each of the following ' + value +
                             ' students, how many challenging behaviors did they exhibit over the past week?\n\nPlease skip to the next student if you did not have any interactions with them, or they did not exhibit any challenging behaviors.', 'radio', students[
                                 0],
                             '1, 1-2 | 2, 3-5 | 3, 5-10 | 4, >10', '', '',
                             "[grades] = '" + key + "' or [grades] = '20'", '',
                             value[:5].lower() + '_students')
        student_df_by_grade.append(df)
    for student in students[1:]:
        student_name_variable = student.replace(' ', '_').lower()
        df = create_new_line(student_name_variable, form_name, '',
                             'radio', student,
                             '1, 1-2 | 2, 3-5 | 3, 5-10 | 4, >10', '', '',
                             "[grades] = '" + key + "' or [grades] = '20'", '',
                             value[:5].lower() + '_students')
        student_df_by_grade.append(df)

student_df_by_grade = pd.DataFrame(data=student_df_by_grade)

dosing_dictionary = dosing_dictionary.append(
    student_df_by_grade, ignore_index=True)

for student in student_df_by_grade['Field Label']:
    variable = student.replace(' ', '_').lower()
    logic = "[" + variable + "] > '0'"
    df = pd.DataFrame(
        {'Variable / Field Name': [variable + '_q1',
                                   variable + '_q2',
                                   variable + '_q3',
                                   variable + '_q4',
                                   variable + '_q5'],
         'Form Name': [form_name for x in range(5)],
         'Section Header': ['Answer the following questions for ' + student + '.',
                            'How many times did you address a problem with each of the following methods?',
                            '','',''],
         'Field Type': ['text' for x in range(5)],
         'Field Label': ["How many times was " + student +
                         " behaving inappropriately or wasn't doing something they were asked to do.",
                         'You warned ' + student +
                         ' of a consequence, delivered a consequence, promised a reward for compliance or purposefully ignored the behavior?',
                         "You pulled " + student +
                         " aside some other time to discuss what's getting in the way and to help find a solution.",
                         'You responded in the moment with "what\'s up?" or "tell me more" when the ' +
                         student + ' asked a question or did not follow an expectation/prompt.',
                         'You proactively dropped an expectation (at least temporarily).'],
         'Field Note': ['Number' for x in range(5)],
         'Text Validation Type OR Show Slider Number': ['integer' for x in range(5)],
         'Branching Logic (Show field only if...)': [logic for x in range(5)]})
    dosing_dictionary = dosing_dictionary.append(df, ignore_index=True)

dosing_dictionary = dosing_dictionary.reindex(
    dosing_dictionary_template.columns, axis=1)

dosing_dictionary.to_csv(path + r'\dosing_dictionary_current.csv')
