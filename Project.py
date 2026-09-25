import csv
import tkinter as tk
from tkinter import filedialog, messagebox

student_master_dict = {}
file_path = "Student-september.csv"

def build_master_dict_from_csv(file_path):
    try:
        with open(file_path, mode='r') as csv_file:
            for line in csv_file:
                print(line)
                first_name = line.split(",")[0]
                last_name = line.split(",")[1].replace("\n","")
                if last_name not in student_master_dict: student_master_dict[last_name] = {first_name: {"cash": 0, "check": 0, "credit": 0, "prepayment": "", "notes": ""}}
                else: student_master_dict[last_name][first_name] = {"cash": 0, "check": 0, "credit": 0, "prepayment": "", "notes": ""}

    except FileNotFoundError: print(f"Error: File '{file_path}' not found.")
    except csv.Error as e: print(f"CSV parsing error: {e}")                
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
    
main()