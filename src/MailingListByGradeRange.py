import tkinter as tk
from tkinter import filedialog, font, messagebox, Button, Label, Entry, OptionMenu, StringVar
from tkinter.scrolledtext import ScrolledText
import pyperclip
import pandas as pd
import webbrowser


try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass


greater_than_percentage = -1
less_than_percentage = 60
file_path = None

range_presets = [
    'Failing CSCI (< 60)',
    'Failing (< 65)',
    'Near failing CSCI (60-66.999)',
    'C- and below (< 73)',
    'A-B range (80-100)',
    'D- (60-62.999)',
    'D  (63-66.999)',
    'C- (70-72.999)',
    'C  (73-76.999)',
    'C  (77-79.999)',
    'B- (80-82.999)',
    'B  (83-86.999)',
    'B+ (87-89.999)',
    'A- (90-92.999)',
    'A  (93-100)',
]

range_presets_dict = {
    'Failing CSCI (< 60)': (-1, 60),
    'Failing (< 65)': (-1, 65),
    'Near failing CSCI (60-66.999)': (59.999, 67),
    'C- and below (< 73)': (-1, 73),
    'A-B range (80-100)': (79.999, 101),
    'D- (60-62.999)': (59.999, 63),
    'D  (63-66.999)': (62.999, 67),
    'C- (70-72.999)': (66.999, 73),
    'C  (73-76.999)': (72.999, 77),
    'C  (77-79.999)': (76.999, 80),
    'B- (80-82.999)': (79.999, 83),
    'B  (83-86.999)': (82.999, 87),
    'B+ (87-89.999)': (86.999, 90),
    'A- (90-92.999)': (89.999, 93),
    'A  (93-100)': (92.999, 101),
}


def copy_mailing_list():
    t = mailing_list_textarea.get(1.0, tk.END); pyperclip.copy(t)


def copy_grades_list():
    t = grades_textarea.get(1.0, tk.END); pyperclip.copy(t)


def apply_preset(val):
    global greater_than_percentage
    global less_than_percentage
    range_values = range_presets_dict[val]
    greater_than_percentage = range_values[0]
    less_than_percentage = range_values[1]
    greater_than_entry.delete(0, tk.END)
    less_than_entry.delete(0, tk.END)
    greater_than_entry.insert(0, greater_than_percentage)
    less_than_entry.insert(0, less_than_percentage)


def check_percentages():
    global greater_than_percentage
    global less_than_percentage
    goodinput = False
    try:
        greater_than_percentage = float(greater_than_entry.get())
        less_than_percentage = float(less_than_entry.get())
        goodinput = True
    except:
        messagebox.showerror(title="Not a number!", message="Enter the percentage as a whole number or decimal with no special characters.")
        greater_than_entry.delete(0, tk.END)
        less_than_entry.delete(0, tk.END)
        greater_than_entry.insert(0, greater_than_percentage)
        less_than_entry.insert(0, less_than_percentage)

    return goodinput


def calc_results():
    global file_path
    check_percentages()
    def _calc():
        df = pd.read_csv(file_path)
        df['Overall Grade'] = df['Calculated Final Grade Scheme Symbol'].str.replace(' %', '')
        df['Overall Grade'] = df['Overall Grade'].astype(float)
        df['Email'] = df['Email'].str.lower()
        result_df = df[
            (df['Overall Grade'] > greater_than_percentage)
            & (df['Overall Grade'] < less_than_percentage)
        ]
        mailing_list = result_df['Email'].str.cat(sep=';')
        grades_list = result_df[['Email', 'Overall Grade']].to_string(index=False, max_cols=4)

        mailing_list_textarea.delete(0.0, tk.END)
        mailing_list_textarea.insert(tk.END, mailing_list)

        grades_textarea.delete(0.0, tk.END)
        grades_textarea.insert(tk.END, grades_list)
    _calc() if file_path else messagebox.showerror('No file!', 'You must import a file first.')


def import_file():
    global file_path
    file_path = filedialog.askopenfilename(title="Select CSV Grades Export file", filetypes=[("Text files", "*.csv"), ("All files", "*.*")])
    if file_path: check_percentages()


def launch_help():
    webbrowser.open("https://github.com/haasr/d2l-mailing-list-by-grade-range-generator/raw/main/Compose%20Email%20List%20of%20Failing%20Students.docx")


def reset():
    global greater_than_percentage
    global less_than_percentage

    greater_than_entry.delete(0, tk.END)
    less_than_entry.delete(0, tk.END)
    mailing_list_textarea.delete(0.0, tk.END)
    grades_textarea.delete(0.0, tk.END)

    range_vals = range_presets_dict[range_presets[0]]
    greater_than_percentage = range_vals[0]
    less_than_percentage = range_vals[1]
    preset.set(range_presets[0])

    greater_than_entry.insert(0, greater_than_percentage)
    less_than_entry.insert(0, less_than_percentage)


# Create the main Tkinter window
root = tk.Tk()
root.configure(background='#F9F9FA')
root.default_font = 'Roboto'
root.default_fontsize = '13'

font_defaults = f"{root.default_font} {root.default_fontsize}"
root.option_add('*Font', font_defaults)
root.option_add('*Label.Font', font_defaults)
root.option_add('*Button.Font', font_defaults)

root.title("Generate Student Mailing List by Grade Range")


# Create an "Import File" button
import_button = Button(root, text="⍈ Import D2L Grades CSV file", command=import_file, background='#86c7fc')
import_button.grid(row=0, pady=10, padx=(24,0))

help_button = Button(root, text="❓ Help", command=launch_help)
help_button.grid(row=0, column=1)

preset = StringVar()
preset.set(range_presets[0])
preset_opt = OptionMenu(root, preset, *range_presets, command=apply_preset)
preset_opt.grid(row=0, column=2, padx=(0, 12))

greater_than_label = Label(root, text='Greater than percentage', background='#F9F9FA').grid(row=1, padx=(19,0))
greater_than_entry = Entry(root)
greater_than_entry.grid(row=1, column=1)
greater_than_entry.insert(0, greater_than_percentage)

calc_button = Button(root, text='⚡ Calculate', command=calc_results, background='#fcba5b', width=12)
calc_button.grid(row=1, column=2, padx=(0, 12))

less_than_label = Label(root, text='Less than percentage', background='#F9F9FA').grid(row=2)
less_than_entry = Entry(root)
less_than_entry.grid(row=2, column=1, pady=5)
less_than_entry.insert(0, less_than_percentage)
reset_button = Button(root, text='⟳ Reset', command=reset, background='#ff8073', width=12)
reset_button.grid(row=2, column=2, padx=(0, 12))

mailing_list_label = Label(root, text='Resulting mailing list:', background='#F9F9FA').grid(row=3, pady=5)
mailing_list_textarea = ScrolledText(root, bg="#E6E6E6", height=10, width=50)
mailing_list_textarea.grid(row=3, column=1, pady=10)

copy_mailing_list = Button(root, text="⎘ Copy", command=copy_mailing_list)
copy_mailing_list.grid(row=4, column=1)

grades_label = Label(root, text='Resulting grades:', background='#F9F9FA').grid(row=6, pady=5)
grades_textarea = ScrolledText(root, bg="#E6E6E6", height=10, width=50)
grades_textarea.grid(row=6, column=1, pady=10)

copy_grades_list = Button(root, text="⎘ Copy", command=copy_grades_list)
copy_grades_list.grid(row=7, column=1, pady=(0, 12))

# Run the Tkinter event loop
root.mainloop()