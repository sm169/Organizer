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
def update_assignments(tree, assignments):
    while True:
        windows = get_open_windows()
        current_items = tree.get_children()

        # Update the Treeview with new or removed windows
        for window in windows:
            if window and window not in assignments:
                assignments[window] = "Unassigned"
                tree.insert("", tk.END, iid=window, values=(window, assignments[window]))

        for item in current_items:
            if item not in windows:
                tree.delete(item)

        # Update assignment column in real-time
        for window in windows:
            if window in assignments:
                tree.item(window, values=(window, assignments[window]))

        # Save the top assignments file
        save_top_assignments(assignments)

        # Save programs assigned to each project
        project_programs = {}
        for window, project in assignments.items():
            if project != "Unassigned":
                if project not in project_programs:
                    project_programs[project] = []
                project_programs[project].append(window)

        for project, programs in project_programs.items():
            save_project_programs(project, programs)

        time.sleep(2)


# GUI setup
def setup_gui():
    projects = get_project_list()

    root = tk.Tk()
    root.title("Window Project Manager")
    root.geometry("900x600")

    # Treeview for windows and assignments
    tree = ttk.Treeview(root, columns=("Window", "Assignment"), show="headings", selectmode="browse")
    tree.heading("Window", text="Window Title")
    tree.heading("Assignment", text="Project Assignment")
    tree.pack(fill=tk.BOTH, expand=True)

    # ComboBox for project selection
    selected_project = tk.StringVar()
    project_dropdown = ttk.Combobox(root, textvariable=selected_project)
    project_dropdown["values"] = projects
    project_dropdown.pack()

    # Assignments dictionary
    assignments = {}

    # Function to assign a selected project to a selected window
    def assign_project():
        selected_item = tree.focus()  # Get selected item
        if selected_item and selected_project.get():
            assignments[selected_item] = selected_project.get()
            tree.item(selected_item, values=(selected_item, selected_project.get()))

            # Save the top assignments file
            save_top_assignments(assignments)

            # Save programs assigned to the specific project
            project_programs = [win for win, proj in assignments.items() if proj == selected_project.get()]
            save_project_programs(selected_project.get(), project_programs)

    # Assign button
    assign_button = ttk.Button(root, text="Assign to Project", command=assign_project)
    assign_button.pack()

    # Start the real-time update thread
    update_thread = Thread(target=update_assignments, args=(tree, assignments), daemon=True)
    update_thread.start()

    root.mainloop()


if __name__ == "__main__":
    setup_gui()
