from pynput import keyboard
import pandas as pd
import pyautogui
import pytesseract
import tkinter as tk
from tkinter import filedialog, messagebox
import time
import os
import re
import io
import numpy as np
import cv2
from PIL import Image

# Set the path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'


keywords_search = [
    '"Heart failure" or "hf" or "HFpEF"  or "HFrEF" or "HFmrEF" or "HFmEF" or "Congestive" or "Cardiomyopathy"',
    '"walker" or "cane" or "roller" or "wheelchair"or "gait" or "mobility" or "exercise" or "walking" or "oxygen"',
    '"interpreter" or "cognitive impairment" OR CARDIAC REHAB OR rehabilitation',
    'Abuse or "alcohol" or "ETOH" or "bipolar" or "schizophrenia" or  "dementia" or cancer or palliative or "lewy"'
]

start_process_flag = False

def start_process():  
    file_path = file_path_entry.get()
    if os.path.exists(file_path) and file_path_entry.get().endswith('.xlsx'):
        dialog = tk.Toplevel(root)
        dialog.title("Start Program")
        
        label = tk.Label(dialog, text="Press 'Start' to begin the program.")
        label.pack(padx=10, pady=10)
        
        start_button = tk.Button(dialog, text="Start", command=lambda: main(file_path_entry.get()))
        start_button.pack(padx=10, pady=10)
    else:
        messagebox.showerror("Error", "Invalid file. Please select a .xlsx file.")

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, filename)

def extract_11_digit_number(s):
    # Regular expression to find 11-digit numbers
    match = re.search(r'\b\d{11}\b', s)
    return match.group(0) if match else None

def read_protected_excel(file_name):
    clients = []
    try:
        df = pd.read_excel(file_name)

        # Check if the 'MRN' column exists
        if 'MRN' in df.columns:
            # Clean the 'MRN' column to ensure it's read as string properly
            df['MRN'] = df['MRN'].str.strip()  # Remove any leading/trailing whitespace
            df['MRN'] = df['MRN'].str.replace('"', '')  # Remove quotes if present

            # Print each value in the MRN column
            for value in df['MRN']:
                eleven_digit_number = extract_11_digit_number(str(value))
                if eleven_digit_number:
                    clients.append(eleven_digit_number)
            return clients
        else:
            print("The 'MRN' column was not found in the Excel file.")
                
    except Exception as e:
        print(f"An error occurred: {e}")
        
def navigate(client):
    # chart Review
    pyautogui.moveTo(137, 36, duration=1)
    pyautogui.click()
    

    # move to name/nrm
    pyautogui.moveTo(111, 111, duration=1)
    pyautogui.click()
    
    # write client number
    pyautogui.write(str(client), interval=0.1)
        
    # move to find
    pyautogui.moveTo(1247, 124, duration=1)
    pyautogui.click()
    time.sleep(2)
    
    # select patient double click
    pyautogui.moveTo(246, 275, duration=1)
    pyautogui.doubleClick()
    time.sleep(4)
    
    output = []
    # pattern search check
    for key in keywords_search:
        
        handled = search_box(key, client)
        
        if handled:
            # implementation to read screen
            # Define the coordinates
            left = 177
            top = 213
            width = 424 - 177
            height = 293 - 213
            
            # Capture a screenshot of the specified area
            screenshot = pyautogui.screenshot(region=(left, top, width, height))
            time.sleep(5)

            # Use OCR to extract text from the screenshot
            extracted_text = pytesseract.image_to_string(screenshot)

            # Check if the specific text is in the screenshot
            if "Hmm..." in extracted_text or "Sorry..." in extracted_text:
                output.append(0)
            else:
                output.append(1)
            
            time.sleep(1)
        else:    
            output.append(-1)
        
    pyautogui.moveTo(338, 59, duration=1)
    pyautogui.click()
    time.sleep(1)
    
    return output

def find_search_box_coordinates(screenshot):
    # Load the smaller template image
    number_of_tries = 0
    while (number_of_tries < 3):
        template = cv2.imread('search_bar.jpg', cv2.IMREAD_GRAYSCALE)

        # Convert the region and template to grayscale
        region_gray = cv2.cvtColor(np.array(screenshot), cv2.COLOR_BGR2GRAY)
        img2 = region_gray.copy()
        w, h = template.shape[::-1]
        
        # Use template matching to find the template in the region
        match_result = cv2.matchTemplate(region_gray, template, cv2.TM_CCOEFF_NORMED)

        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(match_result)
        
        if (max_val > 0.8):
            top_left = max_loc
            bottom_right = (top_left[0] + w, top_left[1] + h)
            
            y_coord = (top_left[1] + bottom_right[1]) // 2
            
            return y_coord
        
        number_of_tries+=1
        print ("Try {}: unable to locate find box".format(number_of_tries))
    
    return None

def search_box(key: str, client: str):
    handled = True
    # Capture a screenshot of the area where the search box is expected to be
    left = 0  # Adjust these coordinates based on your screen
    top = 210
    right = 365
    bottom = 420
    
    width = right-left
    height = bottom-top
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    
    # saveScreenshot(screenshot)
    
    y_coordinate = find_search_box_coordinates(screenshot)

    if y_coordinate:
        handled = True
        y_absolute = top + y_coordinate
        
        print("Search box found at y absolute{}".format(y_absolute))
        
        # Click on the search box
        pyautogui.moveTo(35, y_absolute, duration=1)
        pyautogui.click()
        
        pyautogui.write(key, interval=0.1)
        
        pyautogui.moveTo(16, y_absolute, duration=1)
        pyautogui.click()
        time.sleep(3)
        
    else:
        handled = False
        print("Search box for client {} not found. Manually check the patient".format(client))
    
    return handled

def saveScreenshot(screenshot):
    # Convert the screenshot to a Pillow Image
    screenshot.save('screenshot.png')

def generate_output_filename(base_name, extension):
    count = 1
    while True:
        new_name = f"{base_name}_{count}{extension}"
        if not os.path.exists(new_name):
            return new_name
        count += 1
        
def get_result_string(result):
    if result == -1:
        return "Manual Check Needed"
    elif result == 1:
        return "Yes"
    else:
        return "No"

def main(file_name):
    # Example usage
    clients = read_protected_excel(file_name)
    
    # Create a DataFrame
    df = pd.DataFrame(columns=["Client", "Heart failure", "Mobility", "Interpreter", "Alcohol"])
    
    if clients:
        for client in clients:
            results = navigate(client)

            new_row = pd.DataFrame([{
            "Client": client,
            "Heart failure": get_result_string(results[0]),
            "Mobility": get_result_string(results[1]),
            "Interpreter": get_result_string(results[2]),
            "Alcohol":get_result_string(results[3])
            }])
                
            df = pd.concat([df, new_row], ignore_index=True)
                
        # Save to Excel
        output_file_path = generate_output_filename("client_results",".xlsx")
        df.to_excel(output_file_path, index=False)
        messagebox.showinfo("Process Completed", f"Done!\nOutput file: {output_file_path}")
    else:
        messagebox.showinfo("Client List Empty")
        
if __name__ == "__main__":
    # Set up the tkinter window
    root = tk.Tk()
    root.title("Client Data Processor")

    # File path entry
    file_path_entry = tk.Entry(root, width=50)
    file_path_entry.pack(pady=10)  # Added padding

    # Browse button
    browse_button = tk.Button(root, text="Browse", command=browse_file)
    browse_button.pack(pady=5)  # Added padding

    # Start button
    start_button = tk.Button(root, text="Start", command=start_process)
    start_button.pack(pady=10)  # Added padding

    # Run the tkinter event loop
    root.mainloop()