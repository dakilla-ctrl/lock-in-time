import time
import win32gui


def get_active_window():
    window = win32gui.GetForegroundWindow()
    return win32gui.GetWindowText(window)


if __name__ == "__main__":
    active_window = None
    start_time = time.time()

    while True:
        current_window = get_active_window()
        if current_window != active_window:
            if active_window:
                duration = time.time() - start_time
                print(f"{active_window} - {duration:.2f} seconds")
            active_window = current_window
            start_time = time.time()

        time.sleep(1)
