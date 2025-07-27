# -*- coding: utf-8 -*-
"""
Created on Sun Jul 27 11:20:10 2025

@author: mnn0073
"""

import openpyxl
        


def select_active_tab(workbook, tab_name):
    workbook.active = workbook[tab_name]  


    

def calculate_number_of_labs_attended(active_sheet, tab_name):
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
        
        for column in range(6, active_sheet.max_column):
            grade = active_sheet.cell(row=row, column=column+1).value
            main_student_hash[active_sheet["B" + str(row)].value][tab_name][active_sheet_headers[column]] = grade

            if grade is not None:
                if grade > 0:   
                    lab_participation += 1
        
        
        active_sheet.cell(row=row, column=active_sheet.max_column).value = lab_participation
        main_student_hash[active_sheet["B" + str(row)].value][tab_name][active_sheet_headers[column]] = lab_participation
        
    return main_student_hash




# Execution

from openpyxl import load_workbook

workbook = load_workbook(filename="combined_labexercises_TEST.xlsx")
main_student_hash = {}

#print("Excel tab names: " + " ".join(str(s) for s in workbook.sheetnames))


# TODO: for loop to iterate through sheets here

for wb_tab in workbook.sheetnames:
    select_active_tab(workbook, wb_tab)
    active_sheet = workbook.active
    
    print ("Excel tab title: " + active_sheet.title)
    
    active_sheet.cell(row=1, column=active_sheet.max_column + 1).value = active_sheet.title + " # of labs attended"
    header_row = active_sheet[1]
    active_sheet_headers = [cell.value for cell in header_row]
    
    
    #print(active_sheet_headers) #class name is .__class__.__name__
    #print(active_sheet.max_column) #12

    calculate_number_of_labs_attended(active_sheet, wb_tab)

# ====


print(main_student_hash)

workbook.save(filename="combined_labexercises_TEST.xlsx")
