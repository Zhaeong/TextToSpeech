import tkinter as tk
from tkinter import filedialog, messagebox
from gtts import gTTS
from docx import Document

# Function to read the content of a Word document
def read_word_document(file_path):
    doc = Document(file_path)
    full_text = ""
    for para in doc.paragraphs:
        full_text += para.text + "\n"
    return full_text

# Function to convert text to speech and save it as MP3
def convert_text_to_mp3(text, output_mp3, language):
    tts = gTTS(text=text, lang=language)
    tts.save(output_mp3)

# Main function to convert Word document to MP3
def word_to_mp3(word_file, mp3_output, language):
    print(f"Reading the document: {word_file}")
    text = read_word_document(word_file)
    if text:
        print("Converting text to speech...")
        convert_text_to_mp3(text, mp3_output, language)
        print(f"MP3 file saved as: {mp3_output}")
        messagebox.showinfo("Success", f"MP3 file saved as: {mp3_output}")
    else:
        print("The document is empty!")
        messagebox.showerror("Error", "The document is empty!")

# Function to open the file dialog and select a Word document
def open_file_dialog():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

# Function to handle the conversion when the button is clicked
def convert_button_click():
    word_file = file_entry.get()
    if not word_file:
        messagebox.showerror("Error", "Please select a Word document")
        return

    language = language_var.get()
    if language == "Select Language":
        messagebox.showerror("Error", "Please select a language")
        return

    mp3_output = "output.mp3"  # You can allow the user to choose output filename too
    word_to_mp3(word_file, mp3_output, language)

# Create the main GUI window
root = tk.Tk()
root.title("Word to MP3 Converter")

# Create and place the widgets
tk.Label(root, text="Select Word Document").pack(pady=10)

file_entry = tk.Entry(root, width=50)
file_entry.pack(padx=20, pady=5)

browse_button = tk.Button(root, text="Browse", command=open_file_dialog)
browse_button.pack(pady=5)

tk.Label(root, text="Select Language").pack(pady=10)

language_var = tk.StringVar(value="Select Language")
language_options = [
    "en", "es", "fr", "de", "it", "pt", "ja", "zh", "hi", "ru"  # Add more supported languages
]
language_menu = tk.OptionMenu(root, language_var, *language_options)
language_menu.pack(pady=5)

convert_button = tk.Button(root, text="Convert to MP3", command=convert_button_click)
convert_button.pack(pady=20)

# Run the main event loop
root.mainloop()