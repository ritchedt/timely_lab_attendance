# -*- coding: utf-8 -*-
"""
Created on Sun Jul 27 11:20:10 2025

@author: mnn0073
"""

import openpyxl


def select_active_tab(workbook, tab_name):
    workbook.active = workbook[tab_name]






# Execution

from openpyxl import load_workbook

workbook = load_workbook(filename="combined_labexercises_TEST.xlsx")
main_student_hash = {}


#print("Excel tab names: " + " ".join(str(s) for s in workbook.sheetnames))


# TODO: for loop to iterate through sheets here
select_active_tab(workbook, "2-15")
active_sheet = workbook.active


print ("Excel tab title: " + active_sheet.title)

# ====

active_sheet.cell(row=1, column=active_sheet.max_column + 1).value = active_sheet.title + " # of labs attended"
header_row = active_sheet[1]

active_sheet_headers = [cell.value for cell in header_row]

#print(active_sheet_headers) #class name is .__class__.__name__

#print(active_sheet.max_column) #12

student_index = 1
for row in range(2, active_sheet.max_row + 1):
    student_records = {}
    lab_participation = 0
    student_name = active_sheet["A" + str(row)].value
    
    if student_name is None:
        continue
    
    student_records["student"] = student_name
    student_records["ID"] = active_sheet["B" + str(row)].value
    student_records["sis_user_id"] = active_sheet["C" + str(row)].value
    student_records["sis_login_id"] = active_sheet["D" + str(row)].value
    student_records["root_account"] = active_sheet["E" + str(row)].value
    student_records["section"] = active_sheet["F" + str(row)].value
    #print(student_records)
    
    #print(active_sheet.cell(row=row, column=column).value)
    
    for column in range(6, active_sheet.max_column):
        grade = active_sheet.cell(row=row, column=column+1).value
        student_records[active_sheet_headers[column]] = grade
        if grade is not None:
            if grade > 0:   
                lab_participation += 1
    
    
    active_sheet.cell(row=row, column=active_sheet.max_column).value = lab_participation
    student_records[active_sheet_headers[column]] = lab_participation
    
    main_student_hash[student_index] = student_records
    student_index += 1
    
    

print(main_student_hash)

workbook.save(filename="combined_labexercises_TEST.xlsx")
