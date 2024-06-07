# -*- coding: utf-8 -*-
"""
@author: S7rasshofer
"""

import tkinter as tk
from tkinter import messagebox
import os
from docx import Document
import shutil
import win32com.client


# Path to the Return Reject Templates folder in Documents
templates_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'Return Reject Templates')

# Path to the program's location
program_location = os.path.dirname(os.path.abspath(__file__))

# List of template files to check for
default_templates = ['Out of Policy.docx', 'Wrong Item.docx', 'Wrong Serial.docx']
template_files = [f for f in os.listdir(templates_folder) if f.endswith('.docx')]

# Create the Return Reject Templates folder if it does not exist
if not os.path.exists(templates_folder):
    os.makedirs(templates_folder)

# Check if each template file exists in the Return Reject Templates folder, and copy it from the program's location if not
for template_file in default_templates:
    template_path = os.path.join(templates_folder, template_file)
    if not os.path.exists(template_path):
        source_template_path = os.path.join(program_location, 'reject_templates', template_file)
        shutil.copyfile(source_template_path, template_file)

def update_reason_menu():
    # Update the reason_choices list based on the current contents of the templates folder
    global reason_choices
    template_files = [f for f in os.listdir(templates_folder) if f.endswith('.docx')]
    reason_choices = [os.path.splitext(template)[0] for template in template_files]

    # Clear and update the drop-down menu
    reason_menu['menu'].delete(0, 'end')
    for choice in reason_choices:
        reason_menu['menu'].add_command(label=choice, command=tk._setit(reason_var, choice))
    reason_var.set(reason_choices[0] if reason_choices else "")

#------------------------------------------------------------------------------


def create_document():
    customer_name = customer_name_entry.get()
    order_no = order_no_entry.get()
    reason = reason_var.get()  # Get the selected reason from the drop-down box
    ordered_item = ordered_item_entry.get()
    returned_item = returned_item_entry.get()
    tracking_number = tracking_number_entry.get()
    num_copies = int(copies_entry.get())  # Get the number of copies

    # Path to the selected template document
    template_path = os.path.join(templates_folder, f"{reason}.docx")

    # Create the Shirejects folder on the desktop if it does not exist
    rejects_folder = os.path.join(os.path.expanduser('~'), 'Desktop', 'Rejects')
    if not os.path.exists(rejects_folder):
        os.makedirs(rejects_folder)

    # Copy the template to a new file
    new_doc_path = os.path.join(rejects_folder, f"{customer_name}_{order_no}.docx")
    shutil.copyfile(template_path, new_doc_path)

    # Open the copied template
    doc = Document(new_doc_path)

    # Replace placeholders with customer information
    for paragraph in doc.paragraphs:
        if "{{customer_name}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{customer_name}}", customer_name)
        if "{{order_no}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{order_no}}", order_no)
        if "{{reason}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{reason}}", reason)
        if "{{ordered_item}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{ordered_item}}", ordered_item)
        if "{{returned_item}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{returned_item}}", returned_item)
        if "{{tracking_number}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{tracking_number}}", tracking_number)

    # Save the modified document
    doc.save(new_doc_path)

    # Print the document the specified number of times
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Make Word application invisible
    for i in range(num_copies):
        doc_to_print = word.Documents.Open(new_doc_path)
        doc_to_print.PrintOut()
        doc_to_print.Close(False)  # Close the document without saving
    word.Quit()

    #print(f"{num_copies} copies printed.")
    response = messagebox.askquestion("Next Order", "Do you want to create another reject? Please wait for printout.")
    if response == 'yes':
        # Clear all text boxes
        customer_name_entry.delete(0, tk.END)
        order_no_entry.delete(0, tk.END)
        reason_var.set(reason_choices[0])  # Reset to default reason
        ordered_item_entry.delete(0, tk.END)
        returned_item_entry.delete(0, tk.END)
        tracking_number_entry.delete(0, tk.END)   

def toggle_stay_on_top():
    current_state = window.attributes('-topmost')
    window.attributes('-topmost', not current_state)
    stay_on_top_var.set(not current_state)
    update_stay_on_top_label()

def update_stay_on_top_label():
    new_label = "Toggle Stay on Top" + (" \u2713" if stay_on_top_var.get() else "")
    file_menu.entryconfig(toggle_stay_on_top_index, label=new_label)

def update_menu():
    stay_on_top_menu_label = "Toggle Stay on Top" + (" \u2713" if stay_on_top_var.get() else "")
    file_menu.entryconfig(toggle_stay_on_top_index, label=stay_on_top_menu_label)

    # Clear and update the drop-down menu
    reason_menu['menu'].delete(0, 'end')
    for choice in reason_choices:
        reason_menu['menu'].add_command(label=choice, command=tk._setit(reason_var, choice))
    reason_var.set(reason_choices[0] if reason_choices else "")

def open_templates_folder():
    templates_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'Return Reject templates')
    os.startfile(templates_folder)

def open_saved_files():
    rejects_folder = os.path.join(os.path.expanduser('~'), 'Desktop', 'Rejects')
    os.startfile(rejects_folder)



#------------------------------------------------------------------------------



def apply_theme(theme):
    themes = {
        "Light Mode": {"bg": "#FFFFFF", "fg": "#000000", "entry_bg": "#F0F0F0", "entry_fg": "#000000"},
        "Dark Mode": {"bg": "#1E1E1E", "fg": "#FFFFFF", "entry_bg": "#333333", "entry_fg": "#FFFFFF"},
        "Cyber Hacker": {"bg": "#0F0F0F", "fg": "#33FF33", "entry_bg": "#0F0F0F", "entry_fg": "#33FF33"},
        "Cottage Core": {"bg": "#F5F5DC", "fg": "#6B4226", "entry_bg": "#FDF5E6", "entry_fg": "#6B4226"},
        "Ocean Blue": {"bg": "#87CEEB", "fg": "#000000", "entry_bg": "#B0E0E6", "entry_fg": "#000000"},
        "Forest Green": {"bg": "#228B22", "fg": "#FFFFFF", "entry_bg": "#32CD32", "entry_fg": "#FFFFFF"},
        "Sunset Orange": {"bg": "#FF7F50", "fg": "#000000", "entry_bg": "#FFA07A", "entry_fg": "#000000"},
        "Space Black": {"bg": "#000000", "fg": "#FFFFFF", "entry_bg": "#1C1C1C", "entry_fg": "#FFFFFF"},
        "Vintage": {"bg": "#FFDAB9", "fg": "#000000", "entry_bg": "#FFE4C4", "entry_fg": "#000000"},
        "Futuristic": {"bg": "#1E90FF", "fg": "#FFFFFF", "entry_bg": "#4682B4", "entry_fg": "#FFFFFF"},
        "Minimalist": {"bg": "#E0E0E0", "fg": "#212121", "entry_bg": "#F5F5F5", "entry_fg": "#212121"},
        "Art Deco": {"bg": "#B22222", "fg": "#F5FFFA", "entry_bg": "#DC143C", "entry_fg": "#F5FFFA"},
        "Tropical Orange": {"bg": "#FF9900", "fg": "#232F3E", "entry_bg": "#F8F8F8", "entry_fg": "#232F3E"},
        "Core Blue": {"bg": "#003B64", "fg": "#FFFFFF", "entry_bg": "#00234E", "entry_fg": "#FFFFFF"}
    }

    # Get the selected theme's colors
    theme_colors = themes.get(theme)

    if theme_colors:
        # Set window background color
        window.config(bg=theme_colors["bg"])

        # Set the styles for all widgets
        for widget in window.winfo_children():
            if isinstance(widget, tk.Label):
                widget.config(bg=theme_colors["bg"], fg=theme_colors["fg"])
            elif isinstance(widget, tk.Entry):
                widget.config(bg=theme_colors["entry_bg"], fg=theme_colors["entry_fg"], insertbackground=theme_colors["fg"])
            elif isinstance(widget, tk.Button):
                widget.config(bg=theme_colors["entry_bg"], fg=theme_colors["fg"], activebackground=theme_colors["bg"], activeforeground=theme_colors["fg"])
            elif isinstance(widget, tk.OptionMenu):
                widget.config(bg=theme_colors["entry_bg"], fg=theme_colors["fg"], activebackground=theme_colors["bg"], activeforeground=theme_colors["fg"])    


#------------------------------------------------------------------------------



window = tk.Tk()
window.title("Order Rejects")
window.iconbitmap("face.ico")
window.geometry("260x290")
window.resizable(False, False)

# Variable to track "stay on top" state
stay_on_top_var = tk.BooleanVar()
stay_on_top_var.set(False)  # Initialize to off

# Create a menu
menu_bar = tk.Menu(window)
window.config(menu=menu_bar)

# Menu options
file_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Update Reason List", command=update_reason_menu)
file_menu.add_command(label="Open Templates Folder", command=open_templates_folder)
file_menu.add_command(label="Open Rejects Folder", command=open_saved_files)


# Add toggle stay on top menu item with checkbutton
stay_on_top_menu_label = "Toggle Stay on Top"
toggle_stay_on_top_index = file_menu.index(tk.END) + 1  # Get index for dynamic update
file_menu.add_checkbutton(label=stay_on_top_menu_label, variable=stay_on_top_var, command=toggle_stay_on_top)

# Create a "Themes" menu
theme_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Themes", menu=theme_menu)
theme_menu.add_command(label="Light Mode", command=lambda: apply_theme("Light Mode"))
theme_menu.add_command(label="Dark Mode", command=lambda: apply_theme("Dark Mode"))
theme_menu.add_command(label="Cyber", command=lambda: apply_theme("Cyber Hacker"))
theme_menu.add_command(label="Tropical Orange", command=lambda: apply_theme("Tropical Orange"))
theme_menu.add_command(label="Core Blue", command=lambda: apply_theme("Core Blue"))
theme_menu.add_command(label="Cottage Core", command=lambda: apply_theme("Cottage Core"))
theme_menu.add_command(label="Ocean Blue", command=lambda: apply_theme("Ocean Blue"))
theme_menu.add_command(label="Forest Green", command=lambda: apply_theme("Forest Green"))
theme_menu.add_command(label="Sunset Orange", command=lambda: apply_theme("Sunset Orange"))
theme_menu.add_command(label="Space Black", command=lambda: apply_theme("Space Black"))
theme_menu.add_command(label="Vintage", command=lambda: apply_theme("Vintage"))
theme_menu.add_command(label="Futuristic", command=lambda: apply_theme("Futuristic"))
theme_menu.add_command(label="Minimalist", command=lambda: apply_theme("Minimalist"))
theme_menu.add_command(label="Art Deco", command=lambda: apply_theme("Art Deco"))


#______________________________________________________________________________

# Create labels and entry fields for each parameter
tk.Label(window, text="Customer Name", anchor="e").grid(row=0, column=0, sticky=tk.E, padx=10, pady=5)
customer_name_entry = tk.Entry(window)
customer_name_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(window, text="Order No.", anchor="e").grid(row=1, column=0, sticky=tk.E, padx=10, pady=5)
order_no_entry = tk.Entry(window)
order_no_entry.grid(row=1, column=1, padx=10, pady=5)

# Drop-down box for Reason
reason_label = tk.Label(window, text="Reason", anchor="e")
reason_label.grid(row=2, column=0, sticky=tk.E, padx=10, pady=5)

# Get the list of available reasons (template filenames without the .docx extension)
reason_choices = [os.path.splitext(template)[0] for template in template_files]
reason_var = tk.StringVar(window)
reason_var.set(reason_choices[0] if reason_choices else "")

reason_menu = tk.OptionMenu(window, reason_var, *reason_choices)
reason_menu.grid(row=2, column=1)

tk.Label(window, text="Ordered Item", anchor="e").grid(row=3, column=0, sticky=tk.E, padx=10, pady=5)
ordered_item_entry = tk.Entry(window)
ordered_item_entry.grid(row=3, column=1, padx=10, pady=5)

tk.Label(window, text="Returned Item", anchor="e").grid(row=4, column=0, sticky=tk.E, padx=10, pady=5)
returned_item_entry = tk.Entry(window)
returned_item_entry.grid(row=4, column=1, padx=10, pady=5)

tk.Label(window, text="Tracking Number", anchor="e").grid(row=5, column=0, sticky=tk.E, padx=10, pady=5)
tracking_number_entry = tk.Entry(window)
tracking_number_entry.grid(row=5, column=1, padx=10, pady=5)

tk.Label(window, text="Copies", anchor="e").grid(row=6, column=0, sticky=tk.E, padx=10, pady=5)
copies_entry = tk.Entry(window)
copies_entry.grid(row=6, column=1, padx=10, pady=5)
copies_entry.insert(0, "2")  # Default value for Copies

# Create a button to generate the document
generate_button = tk.Button(window, text="Create Letter", command=create_document)
generate_button.grid(row=7, columnspan=2, pady=20)


apply_theme("Cyber Hacker")
update_reason_menu()

window.mainloop()
