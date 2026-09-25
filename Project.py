import array
from array import array
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

student_master_dict = {}
#'r' makes the string a raw string, so that backslashes are treated as literal backslashes and not escape characters. This is useful for Windows file paths.
cash_file_path = r"September 2026\File 3 Cash_2026.txt"
student_master_path = r"September 2026\student-september.csv"
#function to create a nested dictionary from the master student csv file
def build_master_dict_from_csv(file_path):
    try:
        with open(file_path, mode='r') as csv_file:
            next(csv_file)  # Skip the header row
            for line in csv_file:
                print(line)
                first_name = line.split(",")[0]
                last_name = line.split(",")[1].replace("\n","")
                if last_name not in student_master_dict: student_master_dict[last_name] = {first_name: {"cash": 0, "check": 0, "credit": 0, "prepayment": "", "notes": ""}}
                else: student_master_dict[last_name][first_name] = {"cash": 0, "check": 0, "credit": 0, "prepayment": "", "notes": ""}

    except FileNotFoundError: print(f"Error: File '{file_path}' not found.")
    except csv.Error as e: print(f"CSV parsing error: {e}")

#function to parse the txt cash file 
def cash_parse(path):
        print("test")

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

with open(cash_file_path, "r", encoding="utf-8") as cash_file:
    for line in cash_file:
        payment_line = (line.strip().split('\t'))
        if(len(payment_line) == 3):
            first_name = (line.strip()).split("\t")[0]
            last_name = (line.strip()).split("\t")[1]
            cash_payment = (line.strip()).split("\t")[2]
            if(last_name in student_master_dict and first_name in student_master_dict[last_name]):
                student_master_dict[last_name][first_name]["cash"] += float(cash_payment)
            else: 
                print(f"Error: Student '{first_name} {last_name}' not found in the master dictionary.")