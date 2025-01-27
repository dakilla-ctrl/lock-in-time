import time
import win32gui
import csv
from collections import defaultdict


def get_active_window():
    """Retrieve the title of the currently active window."""
    window = win32gui.GetForegroundWindow()
    return win32gui.GetWindowText(window)


def get_parent_window(window_title):
    """Return the parent window title (e.g., for nested contexts)."""
    try:
        parent_window = win32gui.GetParent(win32gui.GetForegroundWindow())
        parent_title = win32gui.GetWindowText(parent_window) if parent_window else "N/A"
        return parent_title
    except Exception as e:
        print(f"Error: {e}")
        return "N/A"


def format_time(seconds):
    """Convert time in seconds to a human-readable format."""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    seconds = int(seconds % 60)
    return f"{hours}h {minutes}m {seconds}s" if hours > 0 else f"{minutes}m {seconds}s"


def monitor_active_windows(app_totals, context_totals, include_apps=None, exclude_apps=None):
    """Monitor and track the active windows and their durations."""
    active_window = None
    start_time = time.time()

    try:
        while True:
            try:
                current_window = get_active_window()
                parent_window = get_parent_window(current_window)

                # Identify website/page (if possible)
                window_context = current_window.split(" - ")

                if len(window_context) >= 2:
                    app_name = window_context[0]
                    website = window_context[1]
                    page = " - ".join(window_context[2:]) if len(window_context) > 2 else "N/A"
                else:
                    app_name = "Unknown"
                    website = "Unknown"
                    page = "Unknown"

                # Format the current window string as: App - Website - Page
                current_window_format = f"{app_name} - {website} - {page}"

                # Apply filters
                if exclude_apps and app_name in exclude_apps:
                    continue
                if include_apps and app_name not in include_apps:
                    continue

                if current_window_format != active_window:
                    if active_window:
                        duration = time.time() - start_time
                        print(f"{active_window} - {format_time(duration)}")

                        # Update totals
                        app_totals[app_name] += duration
                        context_totals[website] += duration

                    active_window = current_window_format
                    start_time = time.time()

            except Exception as e:
                print(f"Error: {e}")

            time.sleep(1)

    except KeyboardInterrupt:
        print("\nTracking stopped by user.")
        return


def save_to_csv(app_totals, context_totals, filename="active_windows_data.csv"):
    """Save application and context totals to a CSV file."""
    with open(filename, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Application', 'Context', 'Page', 'Time Spent'])

        # Write application totals
        for app_name, duration in app_totals.items():
            writer.writerow([app_name, 'Total', '', format_time(duration)])

        # Write context totals
        for website, duration in context_totals.items():
            writer.writerow([website, 'Total', '', format_time(duration)])

    print(f"\nData successfully saved to {filename}")


def print_totals(app_totals, context_totals):
    """Print total time spent on applications and contexts."""
    print("\nTotal time spent per application:")
    for app_name, total_time in app_totals.items():
        print(f"{app_name} - {format_time(total_time)}")

    print("\nTotal time spent per context:")
    for website, total_time in context_totals.items():
        print(f"{website} - {format_time(total_time)}")


def show_session_summary(app_totals, context_totals):
    """Display a summary of the session."""
    print("\n=== SESSION SUMMARY ===")
    print_totals(app_totals, context_totals)

    total_time = sum(app_totals.values())
    print(f"\nTotal session time: {format_time(total_time)}")
    print("========================\n")


if __name__ == "__main__":
    app_totals = defaultdict(float)  # To store total time by application
    context_totals = defaultdict(float)  # To store total time by context/website/page

    print("Press Ctrl+C to stop tracking.\n")
    include_apps = None  # Add specific apps to track (e.g., ["Chrome", "Notepad"])
    exclude_apps = None  # Add specific apps to ignore (e.g., ["Explorer"])

    monitor_active_windows(app_totals, context_totals, include_apps, exclude_apps)

    # Print and save results after exiting the loop
    show_session_summary(app_totals, context_totals)
    save_to_csv(app_totals, context_totals)