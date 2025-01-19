import tkinter as tk
from tkinter import filedialog, messagebox

from docx import Document
import os

import sys
sys.path.append('third_party/Matcha-TTS')
from cosyvoice.cli.cosyvoice import CosyVoice, CosyVoice2
from cosyvoice.utils.file_utils import load_wav
import torchaudio


# Function to extract text from the Word document
def extract_text_from_word(file_path):
    doc = Document(file_path)
    text = []
    for para in doc.paragraphs:
        text.append(para.text)
    return '\n'.join(text)

# Function to convert text to speech and save as MP3
def convert_text_to_speech(text, language, slow, save_path):
    try:
        # Adjust the speed of the speech based on the 'slow' parameter


        for i, j in enumerate(cosyvoice.inference_zero_shot(text, '希望你以后能够做的比我还好呦。', prompt_speech_16k, stream=False)):
            torchaudio.save('zero_shot_{}.wav'.format(i), j['tts_speech'], cosyvoice.sample_rate)


        messagebox.showinfo("Success", f"wav file saved as {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert text to MP3: {e}")

# Function to browse and select Word file
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        text = extract_text_from_word(file_path)
        text_box.delete(1.0, tk.END)  # Clear existing text
        text_box.insert(tk.END, text)  # Insert extracted text into Text widget
        file_path_label.config(text=f"File selected: {os.path.basename(file_path)}")

# Function to start conversion process
def convert_to_mp3():
    text = text_box.get(1.0, tk.END).strip()
    if not text:
        messagebox.showerror("Error", "Please load a Word document with text.")
        return
    
    # Get selected language, speed and output file name
    language = language_selection.get()
    slow = speed_selection.get() == "Slow"
    output_file = filedialog.asksaveasfilename(defaultextension=".mp3", filetypes=[("MP3 files", "*.mp3")])

    if output_file:
        convert_text_to_speech(text, language, slow, output_file)

# Set up the main window
root = tk.Tk()
root.title("Mandarin Word to MP3 Converter")

# Set up GUI elements
frame = tk.Frame(root, padx=10, pady=10)
frame.pack(padx=20, pady=20)

file_path_label = tk.Label(frame, text="No file selected", width=50, anchor="w")
file_path_label.grid(row=0, column=0, columnspan=2, pady=10)

browse_button = tk.Button(frame, text="Browse Word Document", command=browse_file)
browse_button.grid(row=1, column=0, columnspan=2)

text_box = tk.Text(frame, height=10, width=50)
text_box.grid(row=2, column=0, columnspan=2, pady=10)

# Language selection dropdown
language_label = tk.Label(frame, text="Select Language:")
language_label.grid(row=3, column=0, padx=10, pady=10)

language_options = ["zh", "en", "zh-TW", "zh-CN"]
language_selection = tk.StringVar(value=language_options[0])  # Default to Mandarin
language_dropdown = tk.OptionMenu(frame, language_selection, *language_options)
language_dropdown.grid(row=3, column=1, pady=10)

# Speed selection dropdown
speed_label = tk.Label(frame, text="Select Speed:")
speed_label.grid(row=4, column=0, padx=10, pady=10)

speed_options = ["Fast", "Normal", "Slow"]
speed_selection = tk.StringVar(value=speed_options[1])  # Default to Normal
speed_dropdown = tk.OptionMenu(frame, speed_selection, *speed_options)
speed_dropdown.grid(row=4, column=1, pady=10)

# Convert button
convert_button = tk.Button(frame, text="Convert to Wav", command=convert_to_mp3)
convert_button.grid(row=5, column=0, columnspan=2, pady=20)

cosyvoice = CosyVoice2('pretrained_models/CosyVoice2-0.5B', load_jit=False, load_trt=False, fp16=False)
prompt_speech_16k = load_wav('./asset/cindy.wav', 16000)

# Run the GUI
root.mainloop()
