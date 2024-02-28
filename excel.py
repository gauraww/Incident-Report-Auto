from pandas import read_excel
from pyautogui import press, hotkey, write
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
    # hotkey('win', 'r')  # Open the Run dialog
    # time.sleep(1)
    # write('excel')  # Type 'excel' and press Enter to open Excel
    # press('enter')

    # Wait for Excel to open
    time.sleep(2)  # Adjust the time as needed

    # Get the titles of all visible windows
    windows = pygetwindow.getAllTitles()
    
    # Check if any window title contains "Excel"
    for window_title in windows:
        if "blanksheet" in window_title:
            # Switch to the Excel window
            excel_window = pygetwindow.getWindowsWithTitle(window_title)
            if excel_window:
                time.sleep(1)
                excel_window[0].activate()
                excel_window[0].maximize()
                break  # Stop searching for other Excel windows

    # Paste the copied data into Excel
    # hotkey('ctrl', 'v')
    # time.sleep(2)

    hotkey('alt', 'h')  # Open the Home tab
    time.sleep(2)
    press(['v', 'm'], interval=2)
    time.sleep(2)

    # # center content
    # hotkey('ctrl', 'a')  # Select all cells
    # time.sleep(0.2)
    # hotkey('alt', 'h')  # Open the Home tab
    # time.sleep(0.2)
    # press(['a', 'm'], interval=0.5)
    # time.sleep(0.2)
    # hotkey('ctrl', 'a')  # Select all cells
    # time.sleep(0.2)
    # hotkey('alt', 'h')  # Open the Home tab
    # time.sleep(0.2)
    # press(['a', 'c'], interval=0.5)
    # time.sleep(0.2)

    # # Adjust column width to fit content
    # hotkey('ctrl', 'a')  # Select all cells
    # time.sleep(0.2)
    # hotkey('alt', 'h')  # Open the Home tab
    # time.sleep(0.2)
    # press('o')  # Select 'Format' dropdown
    # time.sleep(0.2)
    # press('w')  # Select 'AutoFit Column Width'
    # time.sleep(0.2)
    # write('10')
    # time.sleep(0.2)
    # press('enter') 
    # time.sleep(0.2)

    # # Adjust row height to fit content
    # hotkey('alt', 'h')  # Open the Home tab
    # time.sleep(0.2)
    # press('o')  # Select 'Format' dropdown
    # time.sleep(0.2)
    # press('h')  # Select 'AutoFit Column Width'
    # time.sleep(0.2)
    # write('30')
    # time.sleep(0.2)
    # press('enter') 
    # time.sleep(0.2)

    # # Adjust column D to fit content
    # press('right', presses=3, interval=0.5)
    # time.sleep(0.2)
    # hotkey('alt', 'h')  # Open the Home tab
    # time.sleep(0.2)
    # press('o')  # Select 'Format' dropdown
    # time.sleep(0.2)
    # press('w')  # Select 'AutoFit Column Width'
    # time.sleep(0.2)
    # write('50')
    # time.sleep(0.2)
    # press('enter') 
    # time.sleep(0.2)

    # # Adjust column E to fit content
    # press('right')
    # time.sleep(0.2)
    # hotkey('alt', 'h')  # Open the Home tab
    # time.sleep(0.2)
    # press('o')  # Select 'Format' dropdown
    # time.sleep(0.2)
    # press('w')  # Select 'AutoFit Column Width'
    # time.sleep(0.2)
    # write('25')
    # time.sleep(0.2)
    # press('enter')
    # time.sleep(0.2)

    # # Adjust column F to fit content
    # press('right')
    # time.sleep(0.2)
    # hotkey('alt', 'h')  # Open the Home tab
    # time.sleep(0.2)
    # press('o')  # Select 'Format' dropdown
    # time.sleep(0.2)
    # press('w')  # Select 'AutoFit Column Width'
    # time.sleep(0.2)
    # write('25')
    # time.sleep(0.2)
    # press('enter')
    # time.sleep(0.2)


    file_path = r'C:\Users\gsingh369\OneDrive - DXC Production\Desktop\Incident Reports\Sample'
    current_date = datetime.now().strftime("%d %B %Y")
    file_name = f'Last 24 hours TTO-TTIR-TTR breached incidents for {current_date}.xlsx'

    # Save the Excel file as 'Last 24 hours TTO-TTIR-TTR breached incidents for {current_date}.xlsx' at the specified location
    press('f12')  # Open the Save As dialog
    time.sleep(2)  # Wait for the dialog to open
    hotkey('alt', 'd')  # Focus on the file path input field
    time.sleep(1)  # Wait for the field to focus
    write(file_path)
    time.sleep(1)
    press('enter')  # Save the file
    press('tab', interval=0.2,presses=6)
    write(file_name)  # Type the file name
    time.sleep(1)
    press('enter')  # Save the file
    time.sleep(2)

    # Get the titles of all visible windows
    windows = pygetwindow.getAllTitles()
    
    # Check if any window title contains "Excel"
    for window_title in windows:
        if "Confirm Save" in window_title:
            # Switch to the Excel window
            confirm_window = pygetwindow.getWindowsWithTitle(window_title)
            if confirm_window:
                time.sleep(1)
                confirm_window[0].activate()
                time.sleep(1)
                press(['left', 'enter'], interval=0.5)
                break  # Stop searching for other Excel windows

    time.sleep(1)
    hotkey('alt', 'f4')
    time.sleep(1)

    # Read the Excel file into a pandas DataFrame
    excel_file = f"{file_path}\{file_name}"
    df = read_excel(excel_file)

    # Check for rows where the "SLT Breached" column has the value "TRUE"
    slt_breached_rows = df[df['SLT Breached'] == True]

    # Save the count of such rows for further use
    slt_breached_count = "no" if len(slt_breached_rows) == 0 else str(len(slt_breached_rows))

    return(slt_breached_count, excel_file)
