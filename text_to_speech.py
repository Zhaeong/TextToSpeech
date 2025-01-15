from docx import Document
from gtts import gTTS

def word_to_mp3(word_file, output_mp3):
    # Step 1: Load the Word document
    doc = Document(word_file)
    full_text = ""
    for paragraph in doc.paragraphs:
        full_text += paragraph.text + "\n"

    # Step 2: Convert the text to speech
    tts = gTTS(full_text, lang='en')  # You can change the language, e.g., 'en', 'es', etc.
    tts.save(output_mp3)
    print(f"MP3 file saved as: {output_mp3}")

# Example usage
word_to_mp3("input.docx", "output.mp3")