from pynput import keyboard
import threading
import pandas as pd
import pyautogui
import pytesseract
import tkinter as tk
from tkinter import filedialog, messagebox
import time
import os
import re
import io

# Set the path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'


keywords_search = [
    '"Heart failure" or "hf" or "HFpEF"  or "HFrEF" or "HFmrEF" or "HFmEF" or "Congestive" or "Cardiomyopathy"',
    '"walker" or "cane" or "roller" or "wheelchair"or "gait" or "mobility" or "exercise" or "walking" or "oxygen"',
    '"interpreter" or "cognitive impairment" OR CARDIAC REHAB OR rehabilitation',
    'Abuse or "alcohol" or "ETOH" or "bipolar" or "schizophrenia" or  "dementia" or cancer or palliative or "lewy"'
]

start_process_flag = False
terminate_program = False

# def start():
#     global start_process_flag
#     global terminate_program
    
#     try:
#         if key == keyboard.Key.enter and start_process_flag:
#             threading.Thread(target=main, args=(file_path_entry.get(),)).start()
#         elif key == keyboard.Key.esc and start_process_flag:  # Set the flag to stop the process when 'esc' is pressed
#             terminate_program = True
#     except AttributeError:
#         pass

def start_process():  
    file_path = file_path_entry.get()
    if os.path.exists(file_path) and file_path_entry.get().endswith('.xlsx'):
        dialog = tk.Toplevel(root)
        dialog.title("Start Program")
        
        label = tk.Label(dialog, text="Press 'Start' to begin the program.")
        label.pack(padx=10, pady=10)
        
        start_button = tk.Button(dialog, text="Start", command=main)
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
    if terminate_program:  # Check if the program should terminate
        return None
    pyautogui.moveTo(137, 36, duration=1)
    pyautogui.click()
    
    if terminate_program:  # Check if the program should terminate
        return None

    # move to name/nrm
    pyautogui.moveTo(111, 111, duration=1)
    pyautogui.click()
    if terminate_program:  # Check if the program should terminate
        return None
    
    # write client number
    pyautogui.write(str(client), interval=0.1)
    if terminate_program:  # Check if the program should terminate
        return None
        
    # move to find
    pyautogui.moveTo(1247, 124, duration=1)
    pyautogui.click()
    if terminate_program:  # Check if the program should terminate
        return None
    time.sleep(2)
    
    # select patient double click
    pyautogui.moveTo(246, 275, duration=1)
    pyautogui.doubleClick()
    if terminate_program:  # Check if the program should terminate
        return None
    time.sleep(4)
    
    output = []
    # pattern search check
    for key in keywords_search:
        if terminate_program:  # Check if the program should terminate
            return None
        
        # pyautogui.moveTo(35, 290, duration=1)
        # pyautogui.click()
        pyautogui.keyDown('ctrl')
        pyautogui.press('space')
        pyautogui.keyUp('ctrl')
        if terminate_program:  # Check if the program should terminate
            return None
        time.sleep(0.5)

        pyautogui.write(key, interval=0.1)
        pyautogui.press('enter')
        # pyautogui.moveTo(16, 290, duration=1)
        # pyautogui.click()
        if terminate_program:  # Check if the program should terminate
            return None
        time.sleep(3)
        
        # implementation to read screen
        # Define the coordinates
        left = 177
        top = 213
        width = 424 - 177
        height = 293 - 213
        
        # Capture a screenshot of the specified area
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        if terminate_program:  # Check if the program should terminate
            return None

        # Use OCR to extract text from the screenshot
        extracted_text = pytesseract.image_to_string(screenshot)

        # Check if the specific text is in the screenshot
        if "Hmm..." in extracted_text:
            output.append(False)
        else:
            output.append(True)
        
        time.sleep(1)
        
    if terminate_program:  # Check if the program should terminate
        return None
    pyautogui.moveTo(338, 59, duration=1)
    pyautogui.click()
    time.sleep(1)
    
    return output

def generate_output_filename(base_name, extension):
    count = 1
    while True:
        new_name = f"{base_name}_{count}{extension}"
        if not os.path.exists(new_name):
            return new_name
        count += 1

def main(file_name):
    # Example usage
    clients = read_protected_excel(file_name)
    
    # Create a DataFrame
    df = pd.DataFrame(columns=["Client", "Heart failure", "Mobility", "Interpreter", "Alcohol"])
    
    if clients:
        for client in clients:
            if terminate_program:  # Check if the program should terminate
                break 
            results = navigate(client)
            
            if results == None:
                new_row = pd.DataFrame([{
                    "Client": "Last Client: " + client,
                    "Heart failure": "",
                    "Mobility": "",
                    "Interpreter": "",
                    "Alcohol": ""
                }])
            elif any(results):
                new_row = pd.DataFrame([{
                "Client": client,
                "Heart failure": "Yes" if results[0] else "",
                "Mobility": "Yes" if results[1] else "",
                "Interpreter": "Yes" if results[2] else "",
                "Alcohol": "Yes" if results[3] else ""
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