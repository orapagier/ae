
import eel
import os
import subprocess
import time
import win32gui
import win32process
import psutil

# Initialize Eel
eel.init('dist')

# Define a dictionary to store all file paths
excel_files = {
    # Grade 1 Forms
    'q1g1': os.path.join('dist', 'files', 'grade1', 'Grade 1 1st Quarter.xlsx'),
    'q2g1': os.path.join('dist', 'files', 'grade1', 'Grade 1 2nd Quarter.xlsx'),
    'q3g1': os.path.join('dist', 'files', 'grade1', 'Grade 1 3rd Quarter.xlsx'),
    'q4g1': os.path.join('dist', 'files', 'grade1', 'Grade 1 4th Quarter.xlsx'),
    'sumg1': os.path.join('dist', 'files', 'grade1', 'Grade 1 SUMMARY.xlsx'),
    'sf1g1': os.path.join('dist', 'files', 'grade1', 'Grade 1 School Form 1 (SF1).xlsx'),
    'sf2g1': os.path.join('dist', 'files', 'grade1', 'Grade 1 School Form 2 (SF2).xlsx'),
    'sf3g1': os.path.join('dist', 'files', 'grade1', 'Grade 1 School Form 3 (SF3).xlsx'),
    'sf5g1': os.path.join('dist', 'files', 'grade1', 'Grade 1 School Form 5 (SF5).xlsx'),
    'sf8g1': os.path.join('dist', 'files', 'grade1', 'Grade 1 School Form 8 (SF8).xlsx'),
    'sf9g1': os.path.join('dist', 'files', 'grade1', 'Grade 1 Form 138 (SF9).xlsx'),
    'sf10g1': os.path.join('dist', 'files', 'grade1', 'Grade 1 Form 137 (SF10).xlsx'),

    # Grade 2 Forms
    'q1g2': os.path.join('dist', 'files', 'grade2', 'Grade 2 1st Quarter.xlsx'),
    'q2g2': os.path.join('dist', 'files', 'grade2', 'Grade 2 2nd Quarter.xlsx'),
    'q3g2': os.path.join('dist', 'files', 'grade2', 'Grade 2 3rd Quarter.xlsx'),
    'q4g2': os.path.join('dist', 'files', 'grade2', 'Grade 2 4th Quarter.xlsx'),
    'sumg2': os.path.join('dist', 'files', 'grade2', 'Grade 2 SUMMARY.xlsx'),
    'sf1g2': os.path.join('dist', 'files', 'grade2', 'Grade 2 School Form 1 (SF1).xlsx'),
    'sf2g2': os.path.join('dist', 'files', 'grade2', 'Grade 2 School Form 2 (SF2).xlsx'),
    'sf3g2': os.path.join('dist', 'files', 'grade2', 'Grade 2 School Form 3 (SF3).xlsx'),
    'sf5g2': os.path.join('dist', 'files', 'grade2', 'Grade 2 School Form 5 (SF5).xlsx'),
    'sf8g2': os.path.join('dist', 'files', 'grade2', 'Grade 2 School Form 8 (SF8).xlsx'),
    'sf9g2': os.path.join('dist', 'files', 'grade2', 'Grade 2 Form 138 (SF9).xlsx'),
    'sf10g2': os.path.join('dist', 'files', 'grade2', 'Grade 2 Form 137 (SF10).xlsx'),

    # Grade 3 Forms
    'q1g3': os.path.join('dist', 'files', 'grade3', 'Grade 3 1st Quarter.xlsx'),
    'q2g3': os.path.join('dist', 'files', 'grade3', 'Grade 3 2nd Quarter.xlsx'),
    'q3g3': os.path.join('dist', 'files', 'grade3', 'Grade 3 3rd Quarter.xlsx'),
    'q4g3': os.path.join('dist', 'files', 'grade3', 'Grade 3 4th Quarter.xlsx'),
    'sumg3': os.path.join('dist', 'files', 'grade3', 'Grade 3 SUMMARY.xlsx'),
    'sf1g3': os.path.join('dist', 'files', 'grade3', 'Grade 3 School Form 1 (SF1).xlsx'),
    'sf2g3': os.path.join('dist', 'files', 'grade3', 'Grade 3 School Form 2 (SF2).xlsx'),
    'sf3g3': os.path.join('dist', 'files', 'grade3', 'Grade 3 School Form 3 (SF3).xlsx'),
    'sf5g3': os.path.join('dist', 'files', 'grade3', 'Grade 3 School Form 5 (SF5).xlsx'),
    'sf8g3': os.path.join('dist', 'files', 'grade3', 'Grade 3 School Form 8 (SF8).xlsx'),
    'sf9g3': os.path.join('dist', 'files', 'grade3', 'Grade 3 Form 138 (SF9).xlsx'),
    'sf10g3': os.path.join('dist', 'files', 'grade3', 'Grade 3 Form 137 (SF10).xlsx'),

    # Grade 4 Forms
    'q1g4': os.path.join('dist', 'files', 'grade4', 'Grade 4 1st Quarter.xlsx'),
    'q2g4': os.path.join('dist', 'files', 'grade4', 'Grade 4 2nd Quarter.xlsx'),
    'q3g4': os.path.join('dist', 'files', 'grade4', 'Grade 4 3rd Quarter.xlsx'),
    'q4g4': os.path.join('dist', 'files', 'grade4', 'Grade 4 4th Quarter.xlsx'),
    'sumg4': os.path.join('dist', 'files', 'grade4', 'Grade 4 SUMMARY.xlsx'),
    'sf1g4': os.path.join('dist', 'files', 'grade4', 'Grade 4 School Form 1 (SF1).xlsx'),
    'sf2g4': os.path.join('dist', 'files', 'grade4', 'Grade 4 School Form 2 (SF2).xlsx'),
    'sf3g4': os.path.join('dist', 'files', 'grade4', 'Grade 4 School Form 3 (SF3).xlsx'),
    'sf5g4': os.path.join('dist', 'files', 'grade4', 'Grade 4 School Form 5 (SF5).xlsx'),
    'sf8g4': os.path.join('dist', 'files', 'grade4', 'Grade 4 School Form 8 (SF8).xlsx'),
    'sf9g4': os.path.join('dist', 'files', 'grade4', 'Grade 4 Form 138 (SF9).xlsx'),
    'sf10g4': os.path.join('dist', 'files', 'grade4', 'Grade 4 Form 137 (SF10).xlsx'),

    # Grade 5 Forms
    'q1g5': os.path.join('dist', 'files', 'grade5', 'Grade 5 1st Quarter.xlsx'),
    'q2g5': os.path.join('dist', 'files', 'grade5', 'Grade 5 2nd Quarter.xlsx'),
    'q3g5': os.path.join('dist', 'files', 'grade5', 'Grade 5 3rd Quarter.xlsx'),
    'q4g5': os.path.join('dist', 'files', 'grade5', 'Grade 5 4th Quarter.xlsx'),
    'sumg5': os.path.join('dist', 'files', 'grade5', 'Grade 5 SUMMARY.xlsx'),
    'sf1g5': os.path.join('dist', 'files', 'grade5', 'Grade 5 School Form 1 (SF1).xlsx'),
    'sf2g5': os.path.join('dist', 'files', 'grade5', 'Grade 5 School Form 2 (SF2).xlsx'),
    'sf3g5': os.path.join('dist', 'files', 'grade5', 'Grade 5 School Form 3 (SF3).xlsx'),
    'sf5g5': os.path.join('dist', 'files', 'grade5', 'Grade 5 School Form 5 (SF5).xlsx'),
    'sf8g5': os.path.join('dist', 'files', 'grade5', 'Grade 5 School Form 8 (SF8).xlsx'),
    'sf9g5': os.path.join('dist', 'files', 'grade5', 'Grade 5 Form 138 (SF9).xlsx'),
    'sf10g5': os.path.join('dist', 'files', 'grade5', 'Grade 5 Form 137 (SF10).xlsx'),

    # Grade 6 Forms
    'q1g6': os.path.join('dist', 'files', 'grade6', 'Grade 6 1st Quarter.xlsx'),
    'q2g6': os.path.join('dist', 'files', 'grade6', 'Grade 6 2nd Quarter.xlsx'),
    'q3g6': os.path.join('dist', 'files', 'grade6', 'Grade 6 3rd Quarter.xlsx'),
    'q4g6': os.path.join('dist', 'files', 'grade6', 'Grade 6 4th Quarter.xlsx'),
    'sumg6': os.path.join('dist', 'files', 'grade6', 'Grade 6 SUMMARY.xlsx'),
    'sf1g6': os.path.join('dist', 'files', 'grade6', 'Grade 6 School Form 1 (SF1).xlsx'),
    'sf2g6': os.path.join('dist', 'files', 'grade6', 'Grade 6 School Form 2 (SF2).xlsx'),
    'sf3g6': os.path.join('dist', 'files', 'grade6', 'Grade 6 School Form 3 (SF3).xlsx'),
    'sf5g6': os.path.join('dist', 'files', 'grade6', 'Grade 6 School Form 5 (SF5).xlsx'),
    'sf8g6': os.path.join('dist', 'files', 'grade6', 'Grade 6 School Form 8 (SF8).xlsx'),
    'sf9g6': os.path.join('dist', 'files', 'grade6', 'Grade 6 Form 138 (SF9).xlsx'),
    'sf10g6': os.path.join('dist', 'files', 'grade6', 'Grade 6 Form 137 (SF10).xlsx'),

}

def get_hwnds_for_pid(pid):
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if found_pid == pid:
                hwnds.append(hwnd)
        return True
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds

@eel.expose
def open_excel_file(file_key):
    file_path = os.path.abspath(excel_files.get(file_key))
    if not file_path:
        return False

    try:
        # Start Excel process
        excel_process = subprocess.Popen(['start', 'excel', file_path], shell=True)
        
        # Wait for the process to start
        time.sleep(2)
        
        # Get the PID of the Excel process
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == 'EXCEL.EXE' and proc.create_time() > (time.time() - 5):
                excel_pid = proc.info['pid']
                break
        
        # Get the window handle
        hwnds = get_hwnds_for_pid(excel_pid)
        if hwnds:
            # Bring the Excel window to the foreground
            win32gui.SetForegroundWindow(hwnds[0])
        
        return True
    except Exception as e:
        return False

@eel.expose
def ensureFullScreen():
  # Any additional logic before triggering fullscreen
  eel.javascript('goFullscreen()')

eel.start('index.html')


