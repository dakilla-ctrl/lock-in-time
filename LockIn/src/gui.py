import tkinter as tk
from tkinter import ttk
import win32gui
import time
import threading

class AppTrackerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("App Tracker")
        self.root.geometry("500x300")

        # Create a Treeview to display application usage
        self.tree = ttk.Treeview(root, columns=("Application", "Time Spent"), show="headings")
        self.tree.heading("Application", text="Application")
        self.tree.heading("Time Spent", text="Time Spent (seconds)")
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Dictionary to store usage times
        self.usage_data = {}

        # Start the tracking in a separate thread
        self.tracking_thread = threading.Thread(target=self.track_usage, daemon=True)
        self.tracking_thread.start()

    def get_active_window(self):
        window = win32gui.GetForegroundWindow()
        return win32gui.GetWindowText(window)

    def update_treeview(self):
        # Clear the treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Insert updated data
        for app, time_spent in self.usage_data.items():
            self.tree.insert("", tk.END, values=(app, time_spent))

    def track_usage(self):
        active_window = None
        start_time = time.time()

        while True:
            current_window = self.get_active_window()
            if current_window != active_window:
                if active_window:
                    duration = time.time() - start_time
                    if active_window in self.usage_data:
                        self.usage_data[active_window] += int(duration)
                    else:
                        self.usage_data[active_window] = int(duration)

                    # Update the Treeview
                    self.update_treeview()

                active_window = current_window
                start_time = time.time()

            time.sleep(1)

if __name__ == "__main__":
    root = tk.Tk()
    app = AppTrackerGUI(root)
    root.mainloop()
