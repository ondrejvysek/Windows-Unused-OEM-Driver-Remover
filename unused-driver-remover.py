import subprocess
import re
import tkinter as tk
from tkinter import ttk, messagebox
import ctypes
from threading import Thread

def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def fetch_unused_drivers():
    try:
        # Run the pnputil command to get a list of drivers
        result = subprocess.run(["pnputil", "/enum-drivers"], capture_output=True, text=True, shell=True)
        
        if result.returncode != 0:
            messagebox.showerror("Error", "Failed to run pnputil. Please run the script as an administrator.")
            return []
        
        drivers_output = result.stdout

        # Parse the output to find unused drivers
        driver_blocks = drivers_output.split("\n\n")  # Each driver is separated by a blank line
        unused_drivers = []
        
        for block in driver_blocks:
            if "Published Name" in block and "Driver is in use" not in block:
                # Extract details
                published_name = re.search(r"Published Name\s*:\s*(oem\d+\.inf)", block)
                original_name = re.search(r"Original Name\s*:\s*(.+)", block)
                provider_name = re.search(r"Provider Name\s*:\s*(.+)", block)
                class_name = re.search(r"Class Name\s*:\s*(.+)", block)
                
                if published_name:
                    unused_drivers.append({
                        "Type": "OEM Driver",
                        "File Name": published_name.group(1),
                        "Original INF Name": original_name.group(1) if original_name else "N/A",
                        "Provider Name": provider_name.group(1) if provider_name else "N/A",
                        "Class Name": class_name.group(1) if class_name else "N/A"
                    })

        return unused_drivers

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        return []

def populate_table_async():
    """Populate the table asynchronously with a splash screen."""
    splash_screen = tk.Toplevel(window)
    splash_screen.title("Loading")
    splash_screen.geometry("300x100")
    splash_label = tk.Label(splash_screen, text="Loading drivers, please wait...")
    splash_label.pack(expand=True, padx=20, pady=20)
    splash_screen.update()

    drivers = fetch_unused_drivers()
    if not drivers:
        splash_screen.destroy()
        return

    for driver in drivers:
        tree.insert("", "end", values=(
            driver["Type"],
            driver["File Name"],
            driver["Original INF Name"],
            driver["Provider Name"],
            driver["Class Name"]
        ))

    total_label_var.set(f"Total Entries: {len(drivers)}")
    splash_screen.destroy()

def on_remove_selected():
    if not is_admin():
        messagebox.showerror("Admin Rights Required", "This action requires administrator privileges. Please restart the script as an administrator.")
        return

    selected_items = tree.selection()
    if not selected_items:
        messagebox.showinfo("Info", "No drivers selected for removal.")
        return

    selected_drivers = [tree.item(item)["values"][1] for item in selected_items]
    confirmation = messagebox.askyesno("Confirm Removal", f"Are you sure you want to remove the following drivers?\n\n{', '.join(selected_drivers)}")
    if confirmation:
        for driver in selected_drivers:
            try:
                result = subprocess.run(["pnputil", "/delete-driver", driver, "/uninstall", "/force"], capture_output=True, text=True, shell=True)
                if result.returncode == 0:
                    messagebox.showinfo("Success", f"Driver {driver} removed successfully.")
                else:
                    messagebox.showerror("Error", f"Failed to remove driver {driver}: {result.stderr}")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while removing {driver}: {e}")
        refresh_table()

def select_all():
    for item in tree.get_children():
        tree.selection_add(item)

def deselect_all():
    for item in tree.get_children():
        tree.selection_remove(item)

def refresh_table():
    tree.delete(*tree.get_children())
    Thread(target=populate_table_async).start()

def sort_table(col, reverse):
    """Sort the table by the specified column."""
    data = [(tree.set(k, col), k) for k in tree.get_children('')]
    data.sort(reverse=reverse)
    for index, (val, k) in enumerate(data):
        tree.move(k, '', index)
    tree.heading(col, command=lambda: sort_table(col, not reverse))

# Initialize the app
if not is_admin():
    messagebox.showwarning("Administrator Privileges Required", "This script needs to be run as an administrator for full functionality.")

# Create the main window
window = tk.Tk()
window.title("Unused OEM Drivers")
window.geometry("1000x600")

# Frame for table and buttons
frame = tk.Frame(window)
frame.pack(fill="both", expand=True, padx=10, pady=10)

# Table (Treeview) for displaying drivers
columns = ("Type", "File Name", "Original INF Name", "Provider Name", "Class Name")
tree = ttk.Treeview(frame, columns=columns, show="headings", height=15, selectmode="extended")

# Scrollbar for the table
scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")

tree.pack(side="left", fill="both", expand=True)

# Configure table columns
for col in columns:
    tree.heading(col, text=col, command=lambda _col=col: sort_table(_col, False))
    tree.column(col, anchor="w", width=200)

# Buttons and total label
button_frame = tk.Frame(window)
button_frame.pack(fill="x", pady=10)

note_label = tk.Label(button_frame, text="Hold CTRL and click to select multiple items.", fg="blue")
note_label.pack(side="left", padx=10)

select_all_button = tk.Button(button_frame, text="Select All", command=select_all)
select_all_button.pack(side="left", padx=10)

deselect_all_button = tk.Button(button_frame, text="Deselect All", command=deselect_all)
deselect_all_button.pack(side="left", padx=10)

remove_button = tk.Button(button_frame, text="Remove Selected Drivers", command=on_remove_selected)
remove_button.pack(side="right", padx=10)

total_label_var = tk.StringVar(value="Total Entries: 0")
total_label = tk.Label(button_frame, textvariable=total_label_var)
total_label.pack(side="right", padx=10)

# Populate the table asynchronously
Thread(target=populate_table_async).start()

# Run the main loop
window.mainloop()
