import tkinter as tk
from tkinter import ttk, filedialog
import win32gui
import time
import threading
import csv
import json
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime

class AppTrackerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("App Tracker")
        self.root.geometry("600x500")

        # Create a Treeview to display application usage
        self.tree = ttk.Treeview(
            root, columns=("Application", "Context", "Time Spent", "Time of Day"), show="headings"
        )
        self.tree.heading("Application", text="Application")
        self.tree.heading("Context", text="Context")
        self.tree.heading("Time Spent", text="Time Spent")
        self.tree.heading("Time of Day", text="Time of Day")
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<ButtonRelease-1>", self.save_scroll_position)  # Bind to save scroll position

        # Scrollbar
        self.tree_scroll = ttk.Scrollbar(root, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Frame for buttons and total session time
        button_frame = tk.Frame(root)
        button_frame.pack(fill=tk.X)

        self.start_button = tk.Button(button_frame, text="Start Tracking", command=self.start_tracking)
        self.start_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.stop_button = tk.Button(button_frame, text="Stop Tracking", command=self.stop_tracking, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.export_button = tk.Button(button_frame, text="Export", command=self.export_data)
        self.export_button.pack(side=tk.RIGHT, padx=5, pady=5)

        # Dropdown menu for format selection
        self.file_format = tk.StringVar(value="CSV")
        format_menu = ttk.OptionMenu(button_frame, self.file_format, "CSV", "CSV", "JSON", "XML")
        format_menu.pack(side=tk.RIGHT, padx=5)

        # Label for total session time
        self.total_time_label = tk.Label(root, text="Total Session Time: 00:00", font=("Arial", 12))
        self.total_time_label.pack(pady=10)

        # Variables to track usage
        self.tracking = False
        self.usage_data = defaultdict(float)
        self.start_times = {}  # Start times for each application/context pair
        self.active_window = None
        self.session_start_time = None
        self.total_session_time = 0
        self.tracking_thread = None
        self.logging_thread = None
        self.last_scroll_position = 0  # To remember scroll position
        self.lock = threading.Lock()  # Lock for thread safety

    def get_active_window(self):
        """Retrieve the title of the currently active window."""
        window = win32gui.GetForegroundWindow()
        return win32gui.GetWindowText(window)

    def update_treeview(self):
        """Update the Treeview with current usage data."""
        self.tree.delete(*self.tree.get_children())  # Clear previous data
        for (app_name, context), time_spent in self.usage_data.items():
            formatted_time = self.format_time(int(time_spent))
            time_of_day = self.start_times.get((app_name, context), "N/A")
            self.tree.insert("", tk.END, values=(app_name, context, formatted_time, time_of_day))

        # Restore scroll position
        self.tree.yview_moveto(self.last_scroll_position)
        # Update total session time
        self.update_total_session_time()

    def save_scroll_position(self, event=None):
        """Save the current scroll position."""
        self.last_scroll_position = self.tree.yview()[0]

    def format_time(self, total_seconds):
        """Format time as hh:mm:ss."""
        return time.strftime("%H:%M:%S", time.gmtime(total_seconds))

    def update_total_session_time(self):
        """Update the total session time label."""
        elapsed_time = int(time.time() - self.session_start_time) if self.session_start_time else 0
        self.total_session_time = elapsed_time
        formatted_time = self.format_time(elapsed_time)
        self.total_time_label.config(text=f"Total Session Time: {formatted_time}")

    def track_usage(self):
        """Track active window usage with flexible title parsing."""
        self.active_window = None
        self.session_start_time = time.time()

        while self.tracking:
            try:
                current_window = self.get_active_window()

                # Split the window title into application and context if possible
                window_context = current_window.split(" - ", 1)  # Split at first occurrence of " - "

                if len(window_context) == 2:
                    # Case where the title has the form "App - Context"
                    app_name = window_context[0].strip()
                    context = window_context[1].strip()
                else:
                    # Case where the title doesn't have " - ", use whole title or default context
                    app_name = window_context[0].strip()  # Use the full title as the app name
                    context = "Main"  # Default context (can be refined based on the app)

                formatted_window = (app_name, context)

                if formatted_window != self.active_window:
                    if self.active_window:
                        # Calculate time spent on the previous window
                        duration = time.time() - self.start_time
                        self.usage_data[self.active_window] += duration

                    # Start a new tracking period for this window
                    self.active_window = formatted_window
                    self.start_time = time.time()

                    # Log the start time for the new application/context
                    if formatted_window not in self.start_times:
                        self.start_times[formatted_window] = datetime.now().strftime("%H:%M:%S")

                self.update_treeview()

            except Exception as e:
                print(f"Error: {e}")

            time.sleep(1)

    def incremental_logging(self):
        """Save logs incrementally every minute."""
        while self.tracking:
            with self.lock:
                self.export_to_csv("incremental_log.csv", append=True)
            time.sleep(60)

    def start_tracking(self):
        """Start tracking active window usage."""
        self.tracking = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        # Start tracking threads
        self.tracking_thread = threading.Thread(target=self.track_usage, daemon=True)
        self.logging_thread = threading.Thread(target=self.incremental_logging, daemon=True)
        self.tracking_thread.start()
        self.logging_thread.start()

    def stop_tracking(self):
        """Stop tracking active window usage."""
        self.tracking = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

        # Merge incremental log to main log
        with self.lock:
            self.export_to_csv("log.csv", append=True)
        print("Tracking stopped and logs updated")

    def export_data(self):
        """Export usage data based on selected file format."""
        format_choice = self.file_format.get()
        if format_choice == "CSV":
            self.export_to_csv()
        elif format_choice == "JSON":
            self.export_to_json()
        elif format_choice == "XML":
            self.export_to_xml()

    def export_to_csv(self, file_name=None, append=False):
        """Export usage data to a CSV file."""
        if not self.usage_data:
            print("No data to export!")
            return

        if not file_name:
            file_name = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Save as",
            )
        if file_name:
            mode = "a" if append else "w"
            try:
                with open(file_name, mode, newline="") as file:
                    writer = csv.writer(file)
                    if not append:
                        writer.writerow(["Application", "Context", "Time Spent", "Time of Day"])
                    for (app_name, context), time_spent in self.usage_data.items():
                        time_of_day = self.start_times.get((app_name, context), "N/A")
                        writer.writerow([app_name, context, self.format_time(int(time_spent)), time_of_day])
                print(f"Data successfully exported to {file_name}")
            except Exception as e:
                print(f"Error exporting to CSV: {e}")

    def export_to_json(self):
        """Export usage data to a JSON file."""
        if not self.usage_data:
            print("No data to export!")
            return

        file_name = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Save as",
        )
        if file_name:
            try:
                data = [
                    {
                        "Application": app_name,
                        "Context": context,
                        "Time Spent": self.format_time(int(time_spent)),
                        "Time of Day": self.start_times.get((app_name, context), "N/A"),
                    }
                    for (app_name, context), time_spent in self.usage_data.items()
                ]
                with open(file_name, "w") as file:
                    json.dump(data, file, indent=4)
                print(f"Data successfully exported to {file_name}")
            except Exception as e:
                print(f"Error exporting to JSON: {e}")

    def export_to_xml(self):
        """Export usage data to an XML file."""
        if not self.usage_data:
            print("No data to export!")
            return

        file_name = filedialog.asksaveasfilename(
            defaultextension=".xml",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
            title="Save as",
        )
        if file_name:
            try:
                root = ET.Element("UsageData")
                for (app_name, context), time_spent in self.usage_data.items():
                    entry = ET.SubElement(root, "Entry")
                    ET.SubElement(entry, "Application").text = app_name
                    ET.SubElement(entry, "Context").text = context
                    ET.SubElement(entry, "TimeSpent").text = self.format_time(int(time_spent))
                    ET.SubElement(entry, "TimeOfDay").text = self.start_times.get((app_name, context), "N/A")

                tree = ET.ElementTree(root)
                tree.write(file_name)
                print(f"Data successfully exported to {file_name}")
            except Exception as e:
                print(f"Error exporting to XML: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppTrackerGUI(root)
    root.mainloop()
