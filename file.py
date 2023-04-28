import os
import re
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pdfrw import PdfReader, PdfWriter, IndirectPdfDict
from PIL import Image
import pytesseract
import openai
import pandas as pd
import docx
import PyPDF2
import speech_recognition as sr
from bs4 import BeautifulSoup
import rpy2.robjects as robjects
from excel_utils import read_excel_file, perform_excel_operation
from google.cloud import speech_v1p1beta1 as speech
from google.oauth2 import service_account
import platform
# Set the path to the excel_utils module
EXCEL_UTILS_PATH = "C:/Code/excel_utils.py"

# Add the path to the system path to import the module
import sys
sys.path.append(os.path.dirname(EXCEL_UTILS_PATH))


openai.api_key = "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"

def open_file_with_default_app(file_path):
    try:
        if platform.system() == 'Windows':
            os.startfile(file_path)
        elif platform.system() == 'Darwin':
            subprocess.call(('open', file_path))
        elif platform.system() == 'Linux':
            subprocess.call(('xdg-open', file_path))
        else:
            raise OSError("Unsupported platform")
    except Exception as e:
        messagebox.showerror("Error", f"Unable to open the file: {str(e)}")



def analyze_document():
    file_path = file_entry.get()
    if not file_path:
        messagebox.showerror("Error", "Please select a file to analyze.")
        return

    if not os.path.isfile(file_path):
        messagebox.showerror("Error", "Invalid file path.")
        return

    file_content = ""
    file_ext = os.path.splitext(file_path)[1].lower()
    print(f"File extension: {file_ext}")  # Debugging print statement

try:
    if file_ext == '.pdf':
        with open(file_path, 'rb') as f:
            pdf_reader = PyPDF2.PdfFileReader(f)
            for page in pdf_reader.pages:
                file_content += page.extractText()
    elif file_ext in ('.xlsx', '.xls', '.xlsm', '.xlsb'):
        df = pd.read_excel(file_path, engine='openpyxl')
        file_content = df.to_string()
    elif file_ext in ('.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm'):
        text = textract.process(file_path)
        file_content = text.decode('utf-8')
    elif file_ext in ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.raw'):
        with Image.open(file_path) as img:
            file_content = pytesseract.image_to_string(img)
    elif file_ext in ('.mp3', '.wav', '.aac', '.wma', '.ogg'):
        r = sr.Recognizer()
        with sr.AudioFile(file_path) as audio_file:
            audio_content = r.record(audio_file)
            file_content = r.recognize_google(audio_content)
    elif file_ext in ('.txt', '.rtf', '.odt'):
        with open(file_path, 'r') as f:
            file_content = f.read()
    elif file_ext == '.ppt' or file_ext == '.pptx' or file_ext == '.pptm' or file_ext == '.pps' or file_ext == '.ppsx' or file_ext == '.ppsm':
        text = textract.process(file_path)
        file_content = text.decode('utf-8')
    elif file_ext in ('.zip', '.rar', '.7z', '.iso'):
        with zipfile.ZipFile(file_path, 'r') as archive:
            for file in archive.namelist():
                with archive.open(file, 'r') as f:
                    file_content += f.read().decode('utf-8')
    elif file_ext == '.csv':
        with open(file_path, 'r') as f:
            reader = csv.reader(f)
            for row in reader:
                file_content += ','.join(row) + '\n'
    elif file_ext == '.xml':
        with open(file_path, 'r') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            for child in root:
                file_content += ET.tostring(child, encoding='unicode')
    elif file_ext == '.json':
        with open(file_path, 'r') as f:
            data = json.load(f)
            file_content = json.dumps(data, indent=4)
    elif file_ext in ('.html', '.css', '.js', '.py', '.java', '.cpp', '.c'):
        with open(file_path, 'r') as f:
            file_content = f.read()
    else:
        messagebox.showerror("Error", f"Unsupported file type: {file_ext}")
        return
except Exception as e:
    print(f"Error: {e}")  # Debugging print statement
    messagebox.showerror("Error", "Unable to read file content.")
    return

# Rest of the function

prompt = f"Please help me understand the following content:\n{file_content}\n\nQuestion: {question}\nAnswer:"
response = get_gpt_response(prompt, max_tokens=50)

result_text.delete('1.0', tk.END)  # Clear the chat box
result_text.insert(tk.END, f"Q: {question}\nA: {response}\n")  # Display question and answer

if "open" in question:
    confirm_prompt = f"Do you want to open {file_path}?"
    confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

    if confirm_response == True:
        try:
            subprocess.run(['start', file_path], shell=True, check=True)
        except:
            result_text.insert(tk.END, "Unable to open file.\n")
    else:
        result_text.insert(tk.END, "Operation canceled.\n")

elif "read" in question:
    confirm_prompt = f"Do you want to read {file_path}?"
    confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

    if confirm_response == True:
        try:
            if file_ext == '.pdf':
                text = extract_text_from_pdf(file_path)
            elif file_ext == '.docx':
                text = extract_text_from_resume(file_path)
            elif file_ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.wav', '.mp3', '.c'):
                with open(file_path, 'r') as f:
                    file_content = f.read()
            else:
                messagebox.showerror("Error", f"Unsupported file type: {file_ext}")
                return
        except Exception as e:
            print(f"Error: {e}")  # Debugging print statement
            messagebox.showerror("Error", "Unable to read file content.")
            return
    else:
        confirm_prompt = f"I'm sorry, I didn't understand your question. Do you want to perform the requested operation?"
        confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)


if confirm_response == True:
    if "open" in question:
        try:
            open_file_with_default_app(file_path)
        except:
            result_text.insert(tk.END, "Unable to open file.\n")
    elif "read" in question:
        try:
            if file_ext == '.pdf':
                text = extract_text_from_pdf(file_path)
            elif file_ext == '.docx':
                text = extract_text_from_resume(file_path)
            elif file_ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif'):
                text = extract_text_from_image(file_path)
            elif file_ext in ('.wav', '.mp3'):
                r = sr.Recognizer()
                with sr.AudioFile(file_path) as source:
                    audio = r.record(source)
                text = r.recognize_google(audio)
            else:
                with open(file_path, 'r') as f:
                    text = f.read()
        except Exception as e:
            print(f"Error: {e}")  # Debugging print statement
            messagebox.showerror("Error", "Unable to read file content.")
            return
        if file_ext == '.c':
            with open(file_path, 'r') as f:
                file_content = f.read()
        else:
            messagebox.showerror("Error", f"Unsupported file type: {file_ext}")
            return



    # Analyze the file content using GPT-3
    prompt = f"Please help me understand the following content:\n{file_content}\n\nQuestion: {question}\nAnswer:"
    response = get_gpt_response(prompt, max_tokens=50)

    # Display the result in the chat box
    result_text.delete('1.0', tk.END)  # Clear the chat box
    result_text.insert(tk.END, f"Q: {question}\nA: {response}\n")  # Display question and answer

    # Perform operations based on the user's request
    if "open" in question:
        confirm_prompt = f"Do you want to open {file_path}?"
        confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

        if confirm_response == True:
            try:
                subprocess.run(['start', file_path], shell=True, check=True)
            except:
                result_text.insert(tk.END, "Unable to open file.\n")
        else:
            result_text.insert(tk.END, "Operation canceled.\n")

    elif "read" in question:
        confirm_prompt = f"Do you want to read {file_path}?"
        confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

        if confirm_response == True:
            try:
                if file_ext == '.pdf':
                    text = extract_text_from_pdf(file_path)
                elif file_ext == '.docx':
                    text = extract_text_from_resume(file_path)
                elif file_ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif'):
                    text = extract_text_from_image(file_path)
                elif file_ext in ('.wav', '.mp3'):
                    r = sr.Recognizer()
                    with sr.AudioFile(file_path) as source:
                        audio = r.record(source)
                    text = r.recognize_google(audio)
                else:
                    with open(file_path, 'r') as f:
                        text = f.read()
            except:
                result_text.insert(tk.END, "Unable to read file content.\n")
        else:
            confirm_prompt = f"I'm sorry, I didn't understand your question. Do you want to perform the requested operation?"
            confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

if confirm_response == True:
    if "open" in question:
        try:
            open_file_with_default_app(file_path)
        except:
            result_text.insert(tk.END, "Unable to open file.\n")
    elif "read" in question:
        try:
            if file_ext == '.pdf':
                text = extract_text_from_pdf(file_path)
            elif file_ext == '.docx':
                text = extract_text_from_resume(file_path)
            elif file_ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif'):
                text = extract_text_from_image(file_path)
            elif file_ext in ('.wav', '.mp3'):
                r = sr.Recognizer()
                with sr.AudioFile(file_path) as source:
                    audio = r.record(source)
                text = r.recognize_google(audio)
            elif file_ext == '.c':
                with open(file_path, 'r') as f:
                    file_content = f.read()
            else:
                with open(file_path, 'r') as f:
                    text = f.read()

            prompt = f"Please help me understand the following content:\n{text}\n\nQuestion: {question}\nAnswer:"
            response = get_gpt_response(prompt, max_tokens=50)

            result_text.delete('1.0', tk.END)  # Clear the chat box
            result_text.insert(tk.END, f"Q: {question}\nA: {response}\n")  # Display question and answer

            if "open" in question:
                confirm_prompt = f"Do you want to open {file_path}?"
                confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

                if confirm_response == True:
                    try:
                        subprocess.run(['start', file_path], shell=True, check=True)
                    except:
                        result_text.insert(tk.END, "Unable to open file.\n")
                else:
                    result_text.insert(tk.END, "Operation canceled.\n")

            elif "read" in question:
                confirm_prompt = f"Do you want to read {file_path}?"
                confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

                if confirm_response == True:
                    if file_ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif'):
                        img = Image.open(file_path)
                        img.show()
                    elif file_ext in ('.mp3', '.wav'):
                        play_music(file_path)
                    elif file_ext == '.pdf':
                        open_pdf(file_path)
                    else:
                        result_text.delete('1.0', tk.END)
                        result_text.insert(tk.END, text)
                else:
                    result_text.insert(tk.END, "Operation canceled.\n")
        except Exception as e:
            print(f"Error: {e}")  # Debugging print statement
            messagebox.showerror("Error", "Unable to read file content.")
            return


def get_gpt_response(prompt):
    try:
        response = openai.Completion.create(
            engine="davinci",
            prompt=prompt,
            max_tokens=50,
            n=1,
            stop=None,
            temperature=0.7,
        )
        return response.choices[0].text.strip()
    except Exception as e:
        print(f"An error occurred while getting GPT-3 response: {str(e)}")
        return None


def ask_question(file_content, question):
    prompt = f"Please help me understand the following content:\n{file_content}\n\nQuestion: {question}\nAnswer:"
    response = get_gpt_response(prompt)

    result_text.delete('1.0', tk.END)  # Clear the chat box
    result_text.insert(tk.END, f"Q: {question}\nA: {response}\n")  # Display question and answer

    if "open" in question:
        confirm_prompt = f"Do you want to open {file_content}?"
        confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

        if confirm_response == True:
            open_file_with_default_app(file_content)
        else:
            result_text.insert(tk.END, "Operation canceled.\n")

    elif "read" in question:
        confirm_prompt = f"Do you want to read {file_content}?"
        confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

        if confirm_response == True:
            open_file_with_default_app(file_content)
        else:
            result_text.insert(tk.END, "Operation canceled.\n")

    elif "sum" in question or "average" in question or "count" in question:
        perform_excel_operation(file_content, question)

    else:
        pass



def extract_text_from_resume(resume_path):
    if resume_path.endswith('.pdf'):
        pdf_file = open(resume_path, 'rb')
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)
        page_count = pdf_reader.getNumPages()
        text = ''
        for page_num in range(page_count):
            page = pdf_reader.getPage(page_num)
            text += page.extractText()
        return text
    elif resume_path.endswith('.docx'):
        doc = docx.Document(resume_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)
    else:
        return "Sorry, I can only read PDF and Word documents at this time."

def extract_text_from_image(image_path):
    with Image.open(image_path) as image_file:
        image_text = pytesseract.image_to_string(image_file)
    return image_text

def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as pdf_file:
        read_pdf = PdfReader(pdf_file)
        text = ''
        for page in read_pdf.pages:
            page_text = page.Text
            text += page_text
        return text

def perform_excel(file_path, operation):
    if operation == 'sum':
        return read_excel_file(file_path).sum().sum()
    elif operation == 'average':
        return read_excel_file(file_path).mean().mean()
    elif operation == 'count':
        return read_excel_file(file_path).shape[0]

def perform_excel_operation(file_path, question):
    if 'sum' in question:
        operation = 'sum'
    elif 'average' in question:
        operation = 'average'
    elif 'count' in question:
        operation = 'count'
    else:
        result_text.insert(tk.END, "I'm sorry, I didn't understand your question.\n")
        return

    result = perform_excel(file_path, operation)
    if result is None:
        result_text.insert(tk.END, f"Unable to perform {operation} operation on {file_path}.\n")
    else:
        result_text.insert(tk.END, f"{operation.capitalize()}: {result}\n")
import tkinter as tk
from tkinter import filedialog, messagebox
import os
def browse_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def analyze_document():
    file_path = file_entry.get()
    if not file_path:
        messagebox.showerror("Error", "Please select a file to analyze.")
        return

    if not os.path.isfile(file_path):
        messagebox.showerror("Error", "Invalid file path.")
        return

    try:
        with open(file_path, 'r') as f:
            file_content = f.read()
    except:
        messagebox.showerror("Error", "Unable to read file content.")
        return

    question = question_entry.get()
    if not question:
        messagebox.showerror("Error", "Please enter a question to analyze.")
        return

    prompt = f"Please help me understand the following content:\n{file_content}\n\nQuestion: {question}\nAnswer:"
    response = get_gpt_response(prompt)

    result_text.delete('1.0', tk.END)  # Clear the chat box
    result_text.insert(tk.END, f"Q: {question}\nA: {response}\n")  # Display question and answer

    if "open" in question:
        confirm_prompt = f"Do you want to open {file_path}?"
        confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

        if confirm_response == True:
            open_file_with_default_app(file_path)
        else:
            result_text.insert(tk.END, "Operation canceled.\n")

    elif "read" in question:
        confirm_prompt = f"Do you want to read {file_path}?"
        confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

        if confirm_response == True:
            open_file_with_default_app(file_path)
        else:
            result_text.insert(tk.END, "Operation canceled.\n")

    elif "sum" in question or "average" in question or "count" in question:
        perform_excel_operation(file_path, question)

  

def create_gui():
    global result_text, file_entry, question_entry

    def browse_file():
        file_path = filedialog.askopenfilename()
        if file_path:
            file_entry.delete(0, tk.END)
            file_entry.insert(0, file_path)

    def analyze_document():
        file_path = file_entry.get()
        if not file_path:
            messagebox.showerror("Error", "Please select a file to analyze.")
            return

        if not os.path.isfile(file_path):
            messagebox.showerror("Error", "Invalid file path.")
            return

        try:
            with open(file_path, 'r') as f:
                file_content = f.read()
        except:
            messagebox.showerror("Error", "Unable to read file content.")
            return

        question = question_entry.get()
        if not question:
            messagebox.showerror("Error", "Please enter a question to analyze.")
            return

        prompt = f"Please help me understand the following content:\n{file_content}\n\nQuestion: {question}\nAnswer:"
        response = get_gpt_response(prompt)

        result_text.delete('1.0', tk.END)  # Clear the chat box
        result_text.insert(tk.END, f"Q: {question}\nA: {response}\n")  # Display question and answer

        if "open" in question:
            confirm_prompt = f"Do you want to open {file_path}?"
            confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

            if confirm_response == True:
                open_file_with_default_app(file_path)
            else:
                result_text.insert(tk.END, "Operation canceled.\n")

        elif "read" in question:
            confirm_prompt = f"Do you want to read {file_path}?"
            confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

            if confirm_response == True:
                open_file_with_default_app(file_path)
            else:
                result_text.insert(tk.END, "Operation canceled.\n")

        elif "sum" in question or "average" in question or "count" in question:
            perform_excel_operation(file_path, question)

        else:
            confirm_prompt = f"I'm sorry, I didn't understand your question. Do you want to perform the requested operation?"
            confirm_response = messagebox.askyesno("Confirm Request", confirm_prompt)

            if confirm_response == True:
                if "open" in question:
                    open_file_with_default_app(file_path)
                elif "read" in question:
                    with open(file_path, 'r') as f:
                        content = f.read()
                        result_text.delete('1.0', tk.END)
                        result_text.insert(tk.END, content)
                else:
                    result_text.insert(tk.END, "I'm sorry, I didn't understand your question.\n")
            else:
                result_text.insert(tk.END, "Operation canceled.\n")

root = tk.Tk()
root.title("Document Analyzer")

file_entry_label = tk.Label(root, text="File Path:")
file_entry_label.grid(row=0, column=0, padx=(10, 0), pady=(10, 5), sticky="w")

file_entry = tk.Entry(root)
file_entry.grid(row=0, column=1, padx=(10, 10), pady=(10, 5))

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=(0, 10), pady=(10, 5))

question_entry_label = tk.Label(root, text="Question:")
question_entry_label.grid(row=1, column=0, padx=(10, 0), pady=(5, 5), sticky="w")

question_entry = tk.Entry(root)
question_entry.grid(row=1, column=1, padx=(10, 10), pady=(5, 5), sticky="ew")

analyze_button = tk.Button(root, text="Analyze", command=analyze_document)
analyze_button.grid(row=1, column=2, padx=(0, 10), pady=(5, 5), sticky="ew")

result_label = tk.Label(root, text="Results:")
result_label.grid(row=2, column=0, padx=(10, 0), pady=(5, 5), sticky="w")

result_text = tk.Text(root, wrap="word")
result_text.grid(row=3, column=0, columnspan=3, padx=(10, 10), pady=(5, 5), sticky="nsew")

# Add a scrollbar to the result_text widget
scrollbar = tk.Scrollbar(root, orient=tk.VERTICAL, command=result_text.yview)
scrollbar.grid(row=3, column=2, sticky="ns")

result_text.config(yscrollcommand=scrollbar.set)

root.columnconfigure(1, weight=1)
root.rowconfigure(3, weight=1)

root.mainloop()






create_gui()

