import tkinter as tk
from tkinter import filedialog
import excel_automation

original_opened = False
template_opened = False
save_path_aquierd = False

def open_original_file():
    global original_opened
    global original_file
    original_file = filedialog.askopenfilename()
    original_opened = True
    validation()

def where_is_template():
    global template_opened
    global template_file
    template_file = filedialog.askopenfilename()
    template_opened = True
    validation()

def where_to_save():
    global save_path_aquierd
    global save_path
    save_path = filedialog.askdirectory()
    save_path_aquierd = True
    validation()

def perform_automation():
    user_input = int(output_quantity_box.get())
    file_date = file_name_date_box.get()
    excel_automation.excel_file_automate(
        original_file, 
        template_file, 
        save_path, 
        user_input, 
        file_date
    )

def validation():
    conditions = [
        bool(output_quantity_box.get()),
        bool(file_name_date_box.get()),
        original_opened, 
        template_opened, 
        save_path_aquierd
    ]
    if all(conditions):
        run_code_button.config(state="normal")
    else:
        run_code_button.config(state="disabled")

root = tk.Tk()
root.title("File selector")

open_original_button = tk.Button(
    root, text="Open", command= open_original_file
)
template_button = tk.Button(root, text= "Open", command=where_is_template)
save_to_button = tk.Button(root, text= "Open", command= where_to_save)
output_quantity_box = tk.Entry(root, width=10, borderwidth=2)
file_name_date_box = tk.Entry(root, width=10, borderwidth=2)
run_code_button = tk.Button(
    root, text= "Futtatás", command= perform_automation, state="disabled"
)

original_file_lable = tk.Label(root, text="Eredeti fájl:")
template_lable = tk.Label(root, text="Sablon fájl:")
save_to_lable = tk.Label(root, text="Hova mentse:")
output_quantity_label = tk.Label(root, text= "Hány fájl készüljön:")
file_name_date_lable = tk.Label(root, text="Dátum a fájl névben:")

original_file_lable.grid(row=0, column=0)
open_original_button.grid(row=0, column=1)
template_lable.grid(row=1, column=0)
template_button.grid(row=1, column=1)
save_to_lable.grid(row=2, column=0)
save_to_button.grid(row=2, column=1)
file_name_date_lable.grid(row=3, column=0)
file_name_date_box.grid(row=3, column=1)
output_quantity_label.grid(row=4, column=0)
output_quantity_box.grid(row=4, column=1)
run_code_button.grid(row=5, column=1)

output_quantity_box.bind("<KeyRelease>", lambda event: validation())
file_name_date_box.bind("<KeyRelease>", lambda event: validation())

root.mainloop()