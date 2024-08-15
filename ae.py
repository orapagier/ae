
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
    'q1g1': os.path.join('dist', 'files', 'grade1', 'G1 1Q.xlsm'),
    'q2g1': os.path.join('dist', 'files', 'grade1', 'G1 2Q.xlsx'),
    'q3g1': os.path.join('dist', 'files', 'grade1', 'G1 3Q.xlsx'),
    'q4g1': os.path.join('dist', 'files', 'grade1', 'G1 4Q.xlsx'),
    'f137g1': os.path.join('dist', 'files', 'grade1', 'FORM 137.xlsm'),
    'f138g1': os.path.join('dist', 'files', 'grade1', 'FORM 138.xlsm'),
    'sumg1': os.path.join('dist', 'files', 'grade1', 'G1 SUMMARY.xlsm'),
    'sf1g1': os.path.join('dist', 'files', 'grade1', 'SF 1.xlsm'),
    'sf2g1': os.path.join('dist', 'files', 'grade1', 'SF 2.xlsm'),
    'sf3g1': os.path.join('dist', 'files', 'grade1', 'SF 3.xlsm'),
    'sf5g1': os.path.join('dist', 'files', 'grade1', 'SF 5.xlsm'),
    'sf8g1': os.path.join('dist', 'files', 'grade1', 'SF 8.xlsm'),
    'sf9g1': os.path.join('dist', 'files', 'grade1', 'SF 9.xlsm'),
    'sf10g1': os.path.join('dist', 'files', 'grade1', 'SF 10.xlsm'),

    # Grade 2 Forms
    'q1g2': os.path.join('dist', 'files', 'grade2', 'G2 1Q.xlsm'),
    'q2g2': os.path.join('dist', 'files', 'grade2', 'G2 2Q.xlsx'),
    'q3g2': os.path.join('dist', 'files', 'grade2', 'G2 3Q.xlsx'),
    'q4g2': os.path.join('dist', 'files', 'grade2', 'G2 4Q.xlsx'),
    'f137g2': os.path.join('dist', 'files', 'grade2', 'FORM 137.xlsm'),
    'f138g2': os.path.join('dist', 'files', 'grade2', 'FORM 138.xlsm'),
    'sumg2': os.path.join('dist', 'files', 'grade2', 'G2 SUMMARY.xlsm'),
    'sf1g2': os.path.join('dist', 'files', 'grade2', 'SF 1.xlsm'),
    'sf2g2': os.path.join('dist', 'files', 'grade2', 'SF 2.xlsm'),
    'sf3g2': os.path.join('dist', 'files', 'grade2', 'SF 3.xlsm'),
    'sf5g2': os.path.join('dist', 'files', 'grade2', 'SF 5.xlsm'),
    'sf8g2': os.path.join('dist', 'files', 'grade2', 'SF 8.xlsm'),
    'sf9g2': os.path.join('dist', 'files', 'grade2', 'SF 9.xlsm'),
    'sf10g2': os.path.join('dist', 'files', 'grade2', 'SF 10.xlsm'),

    # Grade 3 Forms
    'q1g3': os.path.join('dist', 'files', 'grade3', 'G3 1Q.xlsm'),
    'q2g3': os.path.join('dist', 'files', 'grade3', 'G3 2Q.xlsx'),
    'q3g3': os.path.join('dist', 'files', 'grade3', 'G3 3Q.xlsx'),
    'q4g3': os.path.join('dist', 'files', 'grade3', 'G3 4Q.xlsx'),
    'f137g3': os.path.join('dist', 'files', 'grade3', 'FORM 137.xlsm'),
    'f138g3': os.path.join('dist', 'files', 'grade3', 'FORM 138.xlsm'),
    'sumg3': os.path.join('dist', 'files', 'grade3', 'G3 SUMMARY.xlsm'),
    'sf1g3': os.path.join('dist', 'files', 'grade3', 'SF 1.xlsm'),
    'sf2g3': os.path.join('dist', 'files', 'grade3', 'SF 2.xlsm'),
    'sf3g3': os.path.join('dist', 'files', 'grade3', 'SF 3.xlsm'),
    'sf5g3': os.path.join('dist', 'files', 'grade3', 'SF 5.xlsm'),
    'sf8g3': os.path.join('dist', 'files', 'grade3', 'SF 8.xlsm'),
    'sf9g3': os.path.join('dist', 'files', 'grade3', 'SF 9.xlsm'),
    'sf10g3': os.path.join('dist', 'files', 'grade3', 'SF 10.xlsm'),

    # Grade 4 Forms
    'q1g4': os.path.join('dist', 'files', 'grade4', 'G4 1Q.xlsm'),
    'q2g4': os.path.join('dist', 'files', 'grade4', 'G4 2Q.xlsx'),
    'q3g4': os.path.join('dist', 'files', 'grade4', 'G4 3Q.xlsx'),
    'q4g4': os.path.join('dist', 'files', 'grade4', 'G4 4Q.xlsx'),
    'f137g4': os.path.join('dist', 'files', 'grade4', 'FORM 137.xlsm'),
    'f138g4': os.path.join('dist', 'files', 'grade4', 'FORM 138.xlsm'),
    'sumg4': os.path.join('dist', 'files', 'grade4', 'G4 SUMMARY.xlsm'),
    'sf1g4': os.path.join('dist', 'files', 'grade4', 'SF 1.xlsm'),
    'sf2g4': os.path.join('dist', 'files', 'grade4', 'SF 2.xlsm'),
    'sf3g4': os.path.join('dist', 'files', 'grade4', 'SF 3.xlsm'),
    'sf5g4': os.path.join('dist', 'files', 'grade4', 'SF 5.xlsm'),
    'sf8g4': os.path.join('dist', 'files', 'grade4', 'SF 8.xlsm'),
    'sf9g4': os.path.join('dist', 'files', 'grade4', 'SF 9.xlsm'),
    'sf10g4': os.path.join('dist', 'files', 'grade4', 'SF 10.xlsm'),

    # Grade 5 Forms
    'q1g5': os.path.join('dist', 'files', 'grade5', 'G5 1Q.xlsm'),
    'q2g5': os.path.join('dist', 'files', 'grade5', 'G5 2Q.xlsx'),
    'q3g5': os.path.join('dist', 'files', 'grade5', 'G5 3Q.xlsx'),
    'q4g5': os.path.join('dist', 'files', 'grade5', 'G5 4Q.xlsx'),
    'f137g5': os.path.join('dist', 'files', 'grade5', 'FORM 137.xlsm'),
    'f138g5': os.path.join('dist', 'files', 'grade5', 'FORM 138.xlsm'),
    'sumg5': os.path.join('dist', 'files', 'grade5', 'G5 SUMMARY.xlsm'),
    'sf1g5': os.path.join('dist', 'files', 'grade5', 'SF 1.xlsm'),
    'sf2g5': os.path.join('dist', 'files', 'grade5', 'SF 2.xlsm'),
    'sf3g5': os.path.join('dist', 'files', 'grade5', 'SF 3.xlsm'),
    'sf5g5': os.path.join('dist', 'files', 'grade5', 'SF 5.xlsm'),
    'sf8g5': os.path.join('dist', 'files', 'grade5', 'SF 8.xlsm'),
    'sf9g5': os.path.join('dist', 'files', 'grade5', 'SF 9.xlsm'),
    'sf10g5': os.path.join('dist', 'files', 'grade5', 'SF 10.xlsm'),

    # Grade 6 Forms
    'q1g6': os.path.join('dist', 'files', 'grade6', 'G6 1Q.xlsm'),
    'q2g6': os.path.join('dist', 'files', 'grade6', 'G6 2Q.xlsx'),
    'q3g6': os.path.join('dist', 'files', 'grade6', 'G6 3Q.xlsx'),
    'q4g6': os.path.join('dist', 'files', 'grade6', 'G6 4Q.xlsx'),
    'f137g6': os.path.join('dist', 'files', 'grade6', 'FORM 137.xlsm'),
    'f138g6': os.path.join('dist', 'files', 'grade6', 'FORM 138.xlsm'),
    'sumg6': os.path.join('dist', 'files', 'grade6', 'G6 SUMMARY.xlsm'),
    'sf1g6': os.path.join('dist', 'files', 'grade6', 'SF 1.xlsm'),
    'sf2g6': os.path.join('dist', 'files', 'grade6', 'SF 2.xlsm'),
    'sf3g6': os.path.join('dist', 'files', 'grade6', 'SF 3.xlsm'),
    'sf5g6': os.path.join('dist', 'files', 'grade6', 'SF 5.xlsm'),
    'sf8g6': os.path.join('dist', 'files', 'grade6', 'SF 8.xlsm'),
    'sf9g6': os.path.join('dist', 'files', 'grade6', 'SF 9.xlsm'),
    'sf10g6': os.path.join('dist', 'files', 'grade6', 'SF 10.xlsm'),

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


