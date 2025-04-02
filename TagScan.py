import cv2
import pytesseract
pytesseract.pytesseract.tesseract_cmd = "C:/Program Files/Tesseract-OCR/tesseract.exe"
import pandas as pd
import tkinter as tk
from tkinter import filedialog, Label, Button, Text
from PIL import Image, ImageTk
import os
import time

# Global variables
image_path = "captured_tag.jpg"
text_result = ""
excel_file = "data.xlsx"

# Function to select image
def select_image():
    global image_path
    file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
    if file_path:
        image_path = file_path
        display_image(file_path)

# Function to capture image from webcam
def capture_image():
    global image_path
    cap = cv2.VideoCapture(0)

    # Allow camera to adjust
    time.sleep(1)

    ret, frame = cap.read()
    cap.release()
    
    if ret:
        image_path = "captured_tag.jpg"
        cv2.imwrite(image_path, frame)
        display_image(image_path)

# Function to display image in GUI
def display_image(img_path):
    img = Image.open(img_path)
    img = img.resize((250, 250), Image.Resampling.LANCZOS)
    img_tk = ImageTk.PhotoImage(img)
    
    label_image.config(image=img_tk)
    label_image.image = img_tk

# Function to extract text from image
def extract_text():
    global text_result
    if not image_path:
        label_status.config(text="Please select or capture an image first!")
        return

    image = cv2.imread(image_path)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Improved OCR processing
    text_result = pytesseract.image_to_string(gray, config="--psm 6").strip()
    
    text_box.delete("1.0", tk.END)
    text_box.insert(tk.END, text_result)
    label_status.config(text="Text extracted successfully!")

# Function to save extracted text to a single Excel file
def save_to_excel():
    if not text_result:
        label_status.config(text="No text to save!")
        return

    try:
        df = pd.read_excel(excel_file, engine="openpyxl")
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Tag Data"])

    # Append new data
    new_data = pd.DataFrame({"Tag Data": [text_result]})
    df = pd.concat([df, new_data], ignore_index=True)

    # Save back to Excel
    df.to_excel(excel_file, index=False, engine="openpyxl")

    label_status.config(text="Data saved to Excel!")

# GUI Components
root = tk.Tk()
root.title("Tag Scanner")

btn_select = Button(root, text="Select Image", command=select_image)
btn_capture = Button(root, text="Capture Image", command=capture_image)
btn_extract = Button(root, text="Extract Text", command=extract_text)
btn_save = Button(root, text="Save to Excel", command=save_to_excel)

label_image = Label(root)
label_status = Label(root, text="")
text_box = Text(root, height=5, width=40)

# Layout
btn_select.pack(pady=5)
btn_capture.pack(pady=5)
label_image.pack(pady=5)
btn_extract.pack(pady=5)
text_box.pack(pady=5)
btn_save.pack(pady=5)
label_status.pack(pady=5)

root.mainloop()
