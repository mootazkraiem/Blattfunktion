from asyncio.windows_events import NULL
import openpyxl
import customtkinter
import tkinter
from tkinter import filedialog
from translate import Translator
from tkinter import messagebox
import ctypes

# Get the screen resolution
user32 = ctypes.windll.user32
screen_width = user32.GetSystemMetrics(0)
screen_height = user32.GetSystemMetrics(1)

# Calculate the window position based on the screen resolution
window_width = 1000
window_height = 800
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2

# Function to browse and select an Excel file
def browse_xl_file():
    global xl_file_path
    xl_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    if not xl_file_path.endswith(".xlsx"):
        print_message("Invalid file format. Please select a .xlsx file.")
        xl_file_path = None
        return


# Function to perform translation to English
def translate_to_english(text):
    # Check if text is not None
    if text is not None:
        # Create a Translator object
        translator = Translator(from_lang='de', to_lang='en')

        try:
            # Perform translation to English
            translated_text = translator.translate(text)

            # Check if the translation result is not empty
            if translated_text:
                return translated_text
            else:
                # Return the original text as fallback
                return text
        except:
            # Handle case when translation fails
            return 'Translation Error'
    else:
        return ''

def generate_text():
    # Check if xl_file_path is defined
    if 'xl_file_path' not in globals():
        print_message("xl_file_path is not defined")
        return

    # Step 1: Read the Excel file
    wb = openpyxl.load_workbook(xl_file_path)
    sheet = wb.active
    signal_mapping = {}

    for row in sheet.iter_rows(values_only=True):
        signal_mapping[row[0]] = row[1]

    # Step 2: Read the text from the entry1 and entry2 textboxes
    text1 = entry1.get("1.0", tkinter.END).strip()
    text2 = entry2.get("1.0", tkinter.END).strip()

    # Check if entry1 is equal to entry2
    if text1 == '':
        print_message("No requirement given")
        entry2.configure(state='normal')
        entry2.delete("1.0", tkinter.END)
        entry2.configure(state='disabled')
        return
    elif text1 == text2:
        print_message("Signal name XL file must've been wrong!")
        entry2.configure(state='normal')
        entry2.delete("1.0", tkinter.END)
        entry2.configure(state='disabled')
        return

    # Step 3: Iterate over signal names
    for signal_name, signal_replacement in signal_mapping.items():
        # Step 4: Check if signal_name is not None
        if signal_name is not None:
            # Step 5: Check if signal_replacement is not None
            if signal_replacement is not None:
                # Step 6: Search and replace in the text
                text1 = text1.replace(signal_name, str(signal_replacement))

    # Step 7: Translate to English if the checkbox is checked
    if translate_var.get():
        text1 = translate_to_english(text1)

    # Step 8: Display the modified text in the entry2 textbox
    entry2.configure(state='normal')
    entry2.delete("1.0", tkinter.END)
    entry2.insert(tkinter.END, text1)
    entry2.configure(state='disabled')



# Create the main application window
app = customtkinter.CTk()
app.geometry(f"{window_width}x{window_height}+{x}+{y}")
app.title('Sheet function')

#print error message

def print_message(text):
    if text == "No changements done,Signals/requirement sheet must've been wrong":
        messagebox.showerror("Error", text)
    else:
        messagebox.showinfo("Success", text)

# Function to handle checkbox change event
def checkbox_changed():
    print(translate_var.get())

# Set appearance mode and default color theme for customtkinter
customtkinter.set_appearance_mode("dark")  # Could be changed with light/system
customtkinter.set_default_color_theme("blue")  # Themes: blue, green

# Create a frame
frame = customtkinter.CTkFrame(master=app, width=960, height=740, corner_radius=15)
frame.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)

# Create a label for "Paste Requirement Here"
l2 = customtkinter.CTkLabel(master=frame, text="Paste Requirement Here", font=('Centery Gothic', 20))
l2.place(x=20, y=30)

# Create the entry1 textbox
entry1 = customtkinter.CTkTextbox(master=frame, width=900, height=200, font=('Centery Gothic', 16))
entry1.place(x=20, y=110)

# Create a label for "Requirement Generated Here"
l3 = customtkinter.CTkLabel(master=frame, text="Requirement Generated Here", font=('Centery Gothic', 20))
l3.place(x=20, y=360)

# Create the entry2 textbox
entry2 = customtkinter.CTkTextbox(master=frame, width=900, height=200, font=('Centery Gothic', 16), state='disabled')
entry2.place(x=20, y=440)

#browse button
browse_button = customtkinter.CTkButton(master=frame, text="Browse", command=browse_xl_file)
browse_button.place(x=20, y=670)

# Create the translate_var variable to store the checkbox state
translate_var = tkinter.BooleanVar()

# Check box
check_box = customtkinter.CTkCheckBox(master=frame, text="Translate to English", variable=translate_var)
check_box.place(x=200, y=670)

# Configure checkbox command
check_box.configure(command=checkbox_changed)

# Create the generate button
generate_button = customtkinter.CTkButton(master=frame, text="Generate", command=generate_text)
generate_button.place(x=780, y=670)

# Run the application
app.mainloop()
