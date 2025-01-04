import os
import json
import tkinter as tk
from tkinter import ttk
from threading import Thread
import time
import pygetwindow as gw

# Path to the projects directory
PROJECTS_PATH = r"C:\Users\ScottMason(Qrometric\OneDrive - Qrometric\Projects"

# File for top-level assignments
TOP_ASSIGNMENTS_FILE = os.path.join(PROJECTS_PATH, "top_assignments.json")


# Function to get a list of projects (folders in the directory)
def get_project_list():
    try:
        return [f.name for f in os.scandir(PROJECTS_PATH) if f.is_dir()]
    except Exception as e:
        print(f"Error reading project directory: {e}")
        return []


# Function to save all assignments to the top-level JSON
def save_top_assignments(assignments):
    try:
        with open(TOP_ASSIGNMENTS_FILE, "w") as f:
            json.dump(assignments, f, indent=4)
    except Exception as e:
        print(f"Error saving top assignments: {e}")


# Function to save programs assigned to a specific project to its folder
def save_project_programs(project_name, programs):
    project_path = os.path.join(PROJECTS_PATH, project_name)
    project_file = os.path.join(project_path, "current_programs.json")
    print("saving to", project_file)

    # Ensure the project folder exists
    if not os.path.exists(project_path):
        os.makedirs(project_path, exist_ok=True)

    # Save the programs to the project-specific JSON
    try:
        with open(project_file, "w") as f:
            json.dump(programs, f, indent=4)
    except Exception as e:
        print(f"Error saving project programs for {project_name}: {e}")


# Function to get a list of open windows
def get_open_windows():
    return gw.getAllTitles()


# Real-time update function for the GUI
def update_assignments(tree, new_windows_tree, assignments, last_assignments, new_assignments):
    while True:
        windows = get_open_windows()
        new_windows = set(windows) - set(last_assignments.keys())

        # Handle new windows: Add to assignments as "Unassigned" and track in new_assignments
        for new_window in new_windows:
            if new_window not in assignments:
                assignments[new_window] = "Unassigned"
                new_assignments[new_window] = "Unassigned"
                if not new_windows_tree.exists(new_window):  # Add to the new windows table
                    new_windows_tree.insert("", tk.END, iid=new_window, values=(new_window,))

        # Remove closed windows from assignments and last_assignments
        closed_windows = set(last_assignments.keys()) - set(windows)
        for closed_window in closed_windows:
            assignments.pop(closed_window, None)
            last_assignments.pop(closed_window, None)

            # Remove closed windows from the new windows tree
            if new_windows_tree.exists(closed_window):
                new_windows_tree.delete(closed_window)

        # Group and update the left-hand Treeview
        assignment_groups = {}  # Group windows by their assigned project
        for window, project in assignments.items():
            if project not in assignment_groups:
                assignment_groups[project] = []
            assignment_groups[project].append(window)

        for project, group_windows in assignment_groups.items():
            if not tree.exists(project):  # Add group header if missing
                tree.insert("", tk.END, iid=project, text=project, open=True)

            existing_items = set(tree.get_children(project))
            for window in group_windows:
                if not tree.exists(window):  # Add only if it doesn't already exist
                    tree.insert(project, tk.END, iid=window, text=window)

        # Remove items no longer present in the current assignments
        all_tree_items = {item for group in tree.get_children() for item in tree.get_children(group)}
        for item in all_tree_items:
            if item not in assignments:
                tree.delete(item)

        # Save updated assignments and project programs
        save_top_assignments(assignments)
        for project, group_windows in assignment_groups.items():
            if project != "Unassigned":
                save_project_programs(project, group_windows)

        # Update last_assignments to match the current state
        last_assignments.clear()
        last_assignments.update(assignments)

        # Reset new_assignments for the next cycle
        new_assignments.clear()

        time.sleep(2)




# GUI setup
def setup_gui():
    projects = get_project_list()

    root = tk.Tk()
    root.title("Window Project Manager")
    root.geometry("1200x800")

    # Left Frame: Assigned Windows by Project
    left_frame = tk.Frame(root)
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    left_label = tk.Label(left_frame, text="Assigned Windows by Project", font=("Arial", 12, "bold"))
    left_label.pack(anchor="w", padx=10, pady=5)

    tree = ttk.Treeview(left_frame, columns=(), show="tree", selectmode="extended")
    tree.heading("#0", text="Assigned Windows")
    tree.pack(fill=tk.BOTH, expand=True)

    # Right Frame: Newly Detected Windows
    right_frame = tk.Frame(root)
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    right_label = tk.Label(right_frame, text="Newly Detected Windows", font=("Arial", 12, "bold"))
    right_label.pack(anchor="w", padx=10, pady=5)

    new_windows_tree = ttk.Treeview(right_frame, columns=("Window",), show="headings")
    new_windows_tree.heading("Window", text="Window Title")
    new_windows_tree.pack(fill=tk.BOTH, expand=True)

    # ComboBox for project selection
    selected_project = tk.StringVar()
    project_dropdown = ttk.Combobox(root, textvariable=selected_project)
    project_dropdown["values"] = projects
    project_dropdown.pack(pady=10)

    # Assignments dictionaries
    assignments = {}
    last_assignments = {}
    new_assignments = {}

    # Function to assign selected windows to a project
    def assign_project():
        selected_items = new_windows_tree.selection()
        if selected_items and selected_project.get():
            project_name = selected_project.get()

            for item in selected_items:
                # Update assignments dictionary
                assignments[item] = project_name
                
                # Remove from new windows tree
                new_windows_tree.delete(item)
                
                # Remove existing item if present
                if tree.exists(item):
                    tree.delete(item)
                    
                # Create project group if needed    
                if not tree.exists(project_name):
                    tree.insert("", tk.END, iid=project_name, text=project_name, open=True)
                    
                # Add to project group
                tree.insert(project_name, tk.END, iid=item, text=item)

                # Save updated assignments
                save_top_assignments(assignments)

            # Save project-specific programs
            project_programs = [w for w, p in assignments.items() if p == project_name]
            save_project_programs(project_name, project_programs)


    # Assign button
    assign_button = ttk.Button(root, text="Assign to Project", command=assign_project)
    assign_button.pack(pady=5)

    # Start the real-time update thread
    update_thread = Thread(target=update_assignments, args=(tree, new_windows_tree, assignments, last_assignments, new_assignments), daemon=True)
    update_thread.start()

    root.mainloop()




if __name__ == "__main__":
    setup_gui()
