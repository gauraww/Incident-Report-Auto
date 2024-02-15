import pandas as pd
import pyautogui
from datetime import datetime
import time
import win32com.client as win32
import pygetwindow

def getreport():

    existing_file_path = r"C:\Users\gsingh369\OneDrive - DXC Production\Documents\Python Projects\Incident Report Auto\blanksheet.xlsx"
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Workbooks.Open(existing_file_path)
    excel.Visible = True

    # # After copying the tickets, open Microsoft Excel
    # pyautogui.hotkey('win', 'r')  # Open the Run dialog
    # time.sleep(1)
    # pyautogui.write('excel')  # Type 'excel' and press Enter to open Excel
    # pyautogui.press('enter')

    # Wait for Excel to open
    time.sleep(5)  # Adjust the time as needed

        
    # Get the titles of all visible windows
    windows = pygetwindow.getAllTitles()

    # Check if any window title contains "Excel"
    for window_title in windows:
        if "Excel" in window_title:
            # Switch to the Excel window
            excel_window = pygetwindow.getWindowsWithTitle(window_title)
            if excel_window:
                excel_window[0].activate()
                print(f"Switched to window: {window_title}")
                break  # Stop searching for other Excel windows
    else:
        print("Excel window not found.")
        exit()

    time.sleep(2)

    # Paste the copied data into Excel
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    # center content
    pyautogui.hotkey('ctrl', 'a')  # Select all cells
    time.sleep(1)
    pyautogui.hotkey('alt', 'h')  # Open the Home tab
    time.sleep(1)
    pyautogui.press(['a', 'm'], interval=0.5)
    time.sleep(1)

    pyautogui.hotkey('ctrl', 'a')  # Select all cells
    time.sleep(1)
    pyautogui.hotkey('alt', 'h')  # Open the Home tab
    time.sleep(1)
    pyautogui.press(['a', 'c'], interval=0.5)
    time.sleep(1)

    # Adjust column width to fit content
    pyautogui.hotkey('ctrl', 'a')  # Select all cells
    time.sleep(1)
    pyautogui.hotkey('alt', 'h')  # Open the Home tab
    time.sleep(1)
    pyautogui.press('o')  # Select 'Format' dropdown
    time.sleep(1)
    pyautogui.press('w')  # Select 'AutoFit Column Width'
    time.sleep(1)
    pyautogui.write('10')
    time.sleep(1)
    pyautogui.press('enter') 
    time.sleep(1)

    # Adjust row height to fit content
    time.sleep(1)
    pyautogui.hotkey('alt', 'h')  # Open the Home tab
    time.sleep(1)
    pyautogui.press('o')  # Select 'Format' dropdown
    time.sleep(1)
    pyautogui.press('h')  # Select 'AutoFit Column Width'
    time.sleep(1)
    pyautogui.write('30')
    time.sleep(1)
    pyautogui.press('enter') 
    time.sleep(1)

    # Adjust column D to fit content
    pyautogui.press('right', presses=3, interval=0.5)
    time.sleep(1)
    pyautogui.hotkey('alt', 'h')  # Open the Home tab
    time.sleep(1)
    pyautogui.press('o')  # Select 'Format' dropdown
    time.sleep(1)
    pyautogui.press('w')  # Select 'AutoFit Column Width'
    time.sleep(1)
    pyautogui.write('50')
    time.sleep(1)
    pyautogui.press('enter') 
    time.sleep(1)

    # Adjust column E and F to fit content
    pyautogui.press('right')
    time.sleep(1)
    pyautogui.hotkey('alt', 'h')  # Open the Home tab
    time.sleep(1)
    pyautogui.press('o')  # Select 'Format' dropdown
    time.sleep(1)
    pyautogui.press('w')  # Select 'AutoFit Column Width'
    time.sleep(1)
    pyautogui.write('25')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    # Adjust column E and F to fit content
    pyautogui.press('right')
    time.sleep(1)
    pyautogui.hotkey('alt', 'h')  # Open the Home tab
    time.sleep(1)
    pyautogui.press('o')  # Select 'Format' dropdown
    time.sleep(1)
    pyautogui.press('w')  # Select 'AutoFit Column Width'
    time.sleep(1)
    pyautogui.write('25')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    file_path = r'C:\Users\gsingh369\OneDrive - DXC Production\Desktop\Incident Reports\Sample'
    current_date = datetime.now().strftime("%d %B %Y")
    file_name = f'Last 24 hours TTO-TTIR-TTR breached incidents for {current_date}.xlsx'

    # Save the Excel file as 'Last 24 hours TTO-TTIR-TTR breached incidents for {current_date}.xlsx' at the specified location
    pyautogui.press('f12')  # Open the Save As dialog
    time.sleep(2)  # Wait for the dialog to open
    pyautogui.hotkey('alt', 'd')  # Focus on the file path input field
    time.sleep(1)  # Wait for the field to focus
    pyautogui.typewrite(file_path)
    time.sleep(1)
    pyautogui.press('enter')  # Save the file
    pyautogui.press('tab', interval=0.2,presses=6)
    time.sleep(1)
    pyautogui.typewrite(file_name)  # Type the file name
    time.sleep(2)
    pyautogui.press('enter')  # Save the file
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(1)
    pyautogui.press(['right','enter'], interval=1)  # Save the file
    time.sleep(3)

    # Read the Excel file into a pandas DataFrame
    excel_file = f"{file_path}\{file_name}"
    df = pd.read_excel(excel_file)

    # Check for rows where the "SLT Breached" column has the value "TRUE"
    slt_breached_rows = df[df['SLT Breached'] == True]

    # Save the count of such rows for further use
    slt_breached_count = "no" if len(slt_breached_rows) == 0 else str(len(slt_breached_rows))
    print(slt_breached_count)

    return(slt_breached_count, excel_file)
