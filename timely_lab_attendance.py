# -*- coding: utf-8 -*-
"""
Created on Sun Jul 27 11:20:10 2025

@author: Drew Ritcher
"""

import openpyxl


def select_active_tab(workbook, tab_name):
    workbook.active = workbook[tab_name]


def calculate_number_of_labs_attended(active_sheet, tab_name):
    number_of_labs_column = 0
    
    for row in range(2, active_sheet.max_row + 1):
        lab_participation = 0
        student_name = active_sheet["A" + str(row)].value
        
        if student_name is None:
            continue
        
        if active_sheet["B" + str(row)].value not in main_student_hash.keys():
            main_student_hash[active_sheet["B" + str(row)].value] = {}
            main_student_hash[active_sheet["B" + str(row)].value]["student"] = student_name
            main_student_hash[active_sheet["B" + str(row)].value]["sis_user_id"] = active_sheet["C" + str(row)].value
            main_student_hash[active_sheet["B" + str(row)].value]["sis_login_id"] = active_sheet["D" + str(row)].value
            main_student_hash[active_sheet["B" + str(row)].value]["root_account"] = active_sheet["E" + str(row)].value
            main_student_hash[active_sheet["B" + str(row)].value]["section"] = active_sheet["F" + str(row)].value
        
        main_student_hash[active_sheet["B" + str(row)].value][tab_name] = {}
        
        if tab_name not in active_sheet.cell(row=1, column=active_sheet.max_column).value:
            number_of_labs_column = active_sheet.max_column + 1
            active_sheet.cell(row=1, column=number_of_labs_column).value = active_sheet.title + " # of labs attended"
            
            active_sheet_headers.insert(number_of_labs_column, active_sheet.title + " # of labs attended")
        else:
            number_of_labs_column = active_sheet.max_column
            
        
        for column in range(7, number_of_labs_column):
            grade = active_sheet.cell(row=row, column=column).value
            main_student_hash[active_sheet["B" + str(row)].value][tab_name][active_sheet_headers[column]] = grade

            if grade is not None:
                if grade > 0:   
                    lab_participation += 1
        
        
        active_sheet.cell(row=row, column=number_of_labs_column).value = lab_participation
        main_student_hash[active_sheet["B" + str(row)].value][tab_name][active_sheet_headers[column]] = lab_participation
        
    return main_student_hash


def lab_attendance_detection_after_first_lab(num_labs_attended):
    if num_labs_attended < 1:
        return 1
    else:
        return 0
    
    
def lab_attendance_detection_after_second_lab(num_labs_attended):
    if num_labs_attended < 1:
        return 2
    elif num_labs_attended < 2:
        return 1
    else:
        return 0
    
    
def lab_attendance_detection_after_third_lab(num_labs_attended):
    if num_labs_attended < 1:
        return 3
    elif num_labs_attended < 2:
        return 2
    elif num_labs_attended < 3:
        return 1
    else:
        return 0
    
    
def lab_attendance_detection_after_fourth_lab(num_labs_attended):
    if num_labs_attended < 1:
        return 4
    elif num_labs_attended < 2:
        return 3
    elif num_labs_attended < 3:
        return 2
    elif num_labs_attended < 4:
        return 1
    else:
        return 0



# ============================================================================
#  === Change the key lab dates and the excel file (with extension) here ====
# ============================================================================

key_lab_dates = ['2-15', '3-8', '4-5', '4-26']
lab_excel_file = "combined_lab_exercises.xlsx"

# ============================================================================
# ============================================================================
# ============================================================================


from openpyxl import load_workbook

workbook = load_workbook(filename=lab_excel_file)
main_student_hash = {}

if "combined" in workbook.sheetnames:
    workbook.remove(workbook["combined"])

for wb_tab in workbook.sheetnames:
    select_active_tab(workbook, wb_tab)
    active_sheet = workbook.active
    
    header_row = active_sheet[1]
    active_sheet_headers = [cell.value for cell in header_row]

    calculate_number_of_labs_attended(active_sheet, wb_tab)


if "combined" not in active_sheet_headers:
    workbook.create_sheet("combined", 0)
    
select_active_tab(workbook, "combined")
workbook.active.cell(row=1, column=1).value = "Student"
workbook.active.cell(row=1, column=2).value = "ID"
workbook.active.cell(row=1, column=3).value = "SIS User ID"
workbook.active.cell(row=1, column=4).value = "SIS Login ID"
workbook.active.cell(row=1, column=5).value = "Root Account"
workbook.active.cell(row=1, column=6).value = "Section"


all_lab_dates = workbook.sheetnames
all_lab_dates.remove('combined')

for index, lab in enumerate(all_lab_dates):
    workbook.active.cell(row=1, column=7 + index).value = lab


student_index = 2

for key, value in main_student_hash.items():
    total_semester_lab_grade = 20
    num_of_missing_labs = 0
    
    workbook.active.cell(row=student_index, column=1).value = value['student']
    workbook.active.cell(row=student_index, column=2).value = key
    workbook.active.cell(row=student_index, column=3).value = value['sis_user_id']
    workbook.active.cell(row=student_index, column=4).value = value['sis_login_id']
    workbook.active.cell(row=student_index, column=5).value = value['root_account']
    workbook.active.cell(row=student_index, column=6).value = value['section']
    
    current_lab_attendance_detection_formula = 0
    
    for index, lab in enumerate(all_lab_dates):
        workbook.active.cell(row=1, column=7 + index).value = lab + ' labs'
        
        if lab not in value:
            workbook.active.cell(row=student_index, column=7 + index).value = -100
            continue
        
        if lab in key_lab_dates:
            if key_lab_dates.index(lab) == 0:
                total_semester_lab_grade = total_semester_lab_grade - lab_attendance_detection_after_first_lab(value[lab][lab + ' # of labs attended'])
                current_lab_attendance_detection_formula = 0
            elif key_lab_dates.index(lab) == 1:
                total_semester_lab_grade = total_semester_lab_grade - lab_attendance_detection_after_second_lab(value[lab][lab + ' # of labs attended'])
                current_lab_attendance_detection_formula = 1
            elif key_lab_dates.index(lab) == 2:
                total_semester_lab_grade = total_semester_lab_grade - lab_attendance_detection_after_third_lab(value[lab][lab + ' # of labs attended'])
                current_lab_attendance_detection_formula = 2
            elif key_lab_dates.index(lab) == 3:
                total_semester_lab_grade = total_semester_lab_grade - lab_attendance_detection_after_fourth_lab(value[lab][lab + ' # of labs attended'])
                current_lab_attendance_detection_formula = 3
        else:
            if current_lab_attendance_detection_formula == 0:
                 total_semester_lab_grade = total_semester_lab_grade - lab_attendance_detection_after_first_lab(value[lab][lab + ' # of labs attended'])
            elif current_lab_attendance_detection_formula == 1:
                 total_semester_lab_grade = total_semester_lab_grade - lab_attendance_detection_after_second_lab(value[lab][lab + ' # of labs attended'])
            elif current_lab_attendance_detection_formula == 2:
                 total_semester_lab_grade = total_semester_lab_grade - lab_attendance_detection_after_third_lab(value[lab][lab + ' # of labs attended'])
            elif current_lab_attendance_detection_formula == 3:
                 total_semester_lab_grade = total_semester_lab_grade - lab_attendance_detection_after_fourth_lab(value[lab][lab + ' # of labs attended'])
         
        workbook.active.cell(row=student_index, column=7 + index).value = total_semester_lab_grade
        
    student_index = student_index + 1

from datetime import datetime

new_file_created_name = datetime.today().strftime('%Y-%m-%d') + "_" + lab_excel_file

workbook.save(filename=new_file_created_name)
print("Excel file {} updated and saved".format(new_file_created_name))
