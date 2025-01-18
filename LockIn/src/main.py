import time
import win32gui

def get_active_window():
    """Retrieve the title of the currently active window."""
    window = win32gui.GetForegroundWindow()
    return win32gui.GetWindowText(window)

def monitor_active_windows():
    """Monitor and print the duration of active windows."""
    active_window = None
    start_time = time.time()

    while True:
        try:
            current_window = get_active_window()
            if current_window != active_window:
                if active_window:
                    duration = time.time() - start_time
                    print(f"{active_window} - {duration:.2f} seconds")
                active_window = current_window
                start_time = time.time()
        except Exception as e:
            print(f"Error: {e}")

        time.sleep(1)

if __name__ == "__main__":
    monitor_active_windows()
