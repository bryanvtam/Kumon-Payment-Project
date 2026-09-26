import csv
import tkinter as tk
from tkinter import filedialog, messagebox
from striprtf.striprtf import rtf_to_text

student_master_dict = {}
not_found_in_master_dict = {}
total_charge ={"cash": 0, "check": 0, "credit": 0, "prepayment": 0}
#'r' makes the string a raw string, so that backslashes are treated as literal backslashes and not escape characters. This is useful for Windows file paths.

cash_file_path = r'September 2026\File 3 Cash_2026.txt'
student_master_path = r'September 2026\student-september.csv'
credit_file_path = r'September 2026\CreditCardSeptember.rtf'

#function to create a nested dictionary from the master student csv file
def build_master_dict_from_csv(file_path):
    try:
        with open(file_path, mode='r') as csv_file:
            next(csv_file)  # Skip the header row
            for line in csv_file:
                first_name = line.split(",")[0]
                last_name = line.split(",")[1].replace("\n","")
                if last_name not in student_master_dict: student_master_dict[last_name] = {first_name: {"cash": 0, "check": 0, "credit": 0, "prepayment": "", "notes": ""}}
                else: student_master_dict[last_name][first_name] = {"cash": 0, "check": 0, "credit": 0, "prepayment": "", "notes": ""}

    except FileNotFoundError: print(f"Error: File '{file_path}' not found.")
    except csv.Error as e: print(f"CSV parsing error: {e}")

#function that will be reused in all payment types, adding the identified student to the master dict or not found in mast dict
def add_payment_to_dict(first, last, payment, payment_type):
    if payment_type == "prepayment":
        #prepayment is a string and not a number and will have to be handled differently than the other payment types
        student_master_dict[last][first][payment_type] = payment
    else:
        if last in student_master_dict and first in student_master_dict[last]:
            student_master_dict[last][first][payment_type] = (payment)
        else:
            not_found_in_master_dict[last] = {first: {"cash": 0, "check": 0, "credit": 0, "prepayment": "", "notes": ""}}
            not_found_in_master_dict[last][first][payment_type] = float(payment)

#function to parse the txt cash file  and add the payments to the master dict or not found in master dict
def cash_parse(cash_file_path):
    with open(cash_file_path, "r", encoding="utf-8") as cash_file:
        for i in range(3):
            next(cash_file)  # Skip the header row
        for line in cash_file:
            payment_line = (line.strip().split('\t'))
            if(len(payment_line) == 3):

                add_payment_to_dict(payment_line[0], payment_line[1], payment_line[2], "cash")
                total_charge["cash"] += float(payment_line[2])

def credit_parse(file_path):
    with open(file_path, "r", encoding="utf-8", errors="ignore") as rtf_file:
        for line in rtf_to_text(rtf_file.read()).splitlines():
            temp_dict = {}
            temp_last_name = ""
            if ("Paying for" in line):
                if(line.split("|")[3] == "Approved"):
                    for last_name in student_master_dict:
                        if last_name in line.split("|")[1]:
                            temp_last_name = last_name
                            temp_dict[last_name] = [] 
                            for first_name in student_master_dict[last_name]:
                                if first_name in line.split("|")[1]:
                                    temp_dict[last_name].append(first_name)
                        else:
                              
                #handle case where the payment was not charged    
                else:
                    print(line.split("|"))
            elif("Total approved charges" in line):
                for item in line.replace("," , "").split("|"):
                    try:
                        if (float(item)):
                            total_charge["credit"] = float(item)
                            break
                    except ValueError: continue
                break
            if temp_last_name in temp_dict:
                for first_name in temp_dict[temp_last_name]:
                    add_payment_to_dict(first_name, temp_last_name, (float(line.split("|")[4]))/len(temp_dict[temp_last_name]), "credit")

def main():
    root = tk.Tk()
    root.title("Kumon Project")
    root.geometry("700x500")
    root.resizable(False, False)
    tk.Label(root, text="Kumon Payment Program", font=("Arial", 16, "bold")).pack(pady=(20, 10))

    tk.Label(root, text="Step 1: Select the student database CSV", font=("Arial", 11, "bold")).pack(pady=(10, 4))
    roster_label = tk.Label(root, text="No file selected", fg="gray")
    roster_label.pack()


    root.mainloop() 

build_master_dict_from_csv(student_master_path)
credit_parse(r'CreditCardSeptember.rtf')
cash_parse(cash_file_path)
print(student_master_dict)
print(total_charge)