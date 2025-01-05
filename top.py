import os
import json
import tkinter as tk
from tkinter import ttk, simpledialog
from threading import Thread
import time
import subprocess
import pygetwindow as gw
import shutil
import psutil
import win32gui
import win32process
from pywinauto import Application
import ctypes
from pychrome import Browser

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
def save_project_programs(project_name, programs_metadata):
    project_path = os.path.join(PROJECTS_PATH, project_name)
    project_file = os.path.join(project_path, "current_programs.json")

    # Ensure the project folder exists
    if not os.path.exists(project_path):
        os.makedirs(project_path, exist_ok=True)

    # Save the programs metadata to the project-specific JSON
    try:
        with open(project_file, "w") as f:
            json.dump(programs_metadata, f, indent=4)
    except Exception as e:
        print(f"Error saving project programs for {project_name}: {e}")


# Function to close windows by title
def close_windows_by_title(window_titles):
    for title in window_titles:
        for win in gw.getWindowsWithTitle(title):
            try:
                win.close()
            except Exception as e:
                print(f"Error closing window {title}: {e}")


# Function to upload project to Git
def upload_project_to_git(project_name):
    project_path = os.path.join(PROJECTS_PATH, project_name)
    try:
        # Add and commit changes
        subprocess.run(["git", "-C", project_path, "add", "."], check=True)
        subprocess.run(["git", "-C", project_path, "commit", "-m", "Project closed"], check=True)
        
        # Try pushing with upstream for first time push
        try:
            subprocess.run(["git", "-C", project_path, "push", "--set-upstream", "origin", "master"], check=True)
        except subprocess.CalledProcessError:
            # If upstream is already set, do regular push
            subprocess.run(["git", "-C", project_path, "push"], check=True)
            
        print(f"Uploaded project {project_name} to Git.")
    except Exception as e:
        print(f"Error uploading project {project_name} to Git: {e}")



# Function to prompt for project notes
def prompt_for_notes(project_name):
    notes_file_path = os.path.join(PROJECTS_PATH, project_name, "closing_notes.txt")
    notes = tk.simpledialog.askstring("Project Notes", f"Enter notes for {project_name}:")
    if notes:
        with open(notes_file_path, "w") as f:
            f.write(notes)
        print(f"Notes saved for project {project_name}.")


# Function to close the active project
def close_active_project(active_project, assignments, tree):
    if not active_project:
        print("No active project set.")
        return

    project_path = os.path.join(PROJECTS_PATH, active_project)
    savestate_path = os.path.join(project_path, "Savedstates")
    last_state_path = os.path.join(project_path, "Savedstates\LastState")
    

    # Get next state number
    existing_states = [
        int(f.replace("state", "").split(".")[0])  # Remove 'state' and extract the numeric part
        for f in os.listdir(savestate_path)
        if f.startswith("state") and f.endswith(".json")  # Ensure the file matches the expected pattern
    ]
    next_state = max(existing_states, default=0) + 1

    # Copy last_state/state.json to savestate/state{n+1}.json if it exists
    last_state_file = os.path.join(last_state_path, "state.json")
    if os.path.exists(last_state_file):
        shutil.copy2(last_state_file, os.path.join(savestate_path, f"state{next_state}.json"))

    # Copy current_programs.json to last_state/state.json
    current_programs = os.path.join(project_path, "current_programs.json")
    if os.path.exists(current_programs):
        shutil.copy2(current_programs, last_state_file)

    # Get all windows associated with the active project
    project_windows = [window for window, project in assignments.items() if project == active_project]

    # Close all associated windows
    close_windows_by_title(project_windows)

    # Upload project folder to Git
    upload_project_to_git(active_project)

    # Prompt for notes
    prompt_for_notes(active_project)

    # Remove project from the GUI
    if tree.exists(active_project):
        tree.delete(active_project)

    # Remove project from assignments
    for window in project_windows:
        assignments.pop(window, None)


# Function to get a list of open windows
def get_open_windows():
    windows_data = []

    # Helper function to get the executable name from a PID
    def get_process_name(pid):
        try:
            process = psutil.Process(pid)
            return process.name(), process.exe()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            return None, None

    # Get all windows using pygetwindow
    windows = gw.getAllWindows()

    # Connect to Chrome debugging protocol for Chrome-specific metadata
    try:
        browser = Browser(url="http://127.0.0.1:9222")
        chrome_tabs = browser.list_tab()
        chrome_data = {tab["id"]: tab for tab in chrome_tabs}
    except Exception:
        chrome_data = {}

    # Iterate through each window to collect metadata
    for window in windows:
        if not window.title.strip():
            continue

        hwnd = window._hWnd
        metadata = {
            "title": window.title,
            "pid": None,
            "process_name": None,
            "exe_path": None,
            "class_name": None,
            "rect": window.box,
            "is_visible": None,
            "chrome_url": None,
            "vscode_workspace": None,
        }

        # Get process and executable info using win32process and psutil
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            metadata["pid"] = pid
            process_name, exe_path = get_process_name(pid)
            metadata["process_name"] = process_name
            metadata["exe_path"] = exe_path
        except Exception:
            pass

        # Get window class name and visibility using win32gui
        try:
            metadata["class_name"] = win32gui.GetClassName(hwnd)
            metadata["is_visible"] = win32gui.IsWindowVisible(hwnd)
        except Exception:
            pass

        # Check if the window belongs to Chrome and get tab URL
        if metadata["process_name"] == "chrome.exe":
            for tab_id, tab_info in chrome_data.items():
                if tab_info["title"] in metadata["title"]:
                    metadata["chrome_url"] = tab_info["url"]

        # Check if the window belongs to VSCode and get workspace info
        if metadata["process_name"] == "Code.exe":
            try:
                app = Application(backend="uia").connect(process=metadata["pid"])
                main_window = app.window(title=metadata["title"])
                workspace = main_window.child_window(auto_id="workbench.parts.editor").texts()
                metadata["vscode_workspace"] = workspace
            except Exception:
                pass

        # Add metadata to the result
        windows_data.append(metadata)

    return windows_data


# Real-time update function for the GUI
def update_assignments(tree, new_windows_tree, assignments, last_assignments, new_assignments):
    while True:
        # Fetch the current open windows (list of dictionaries)
        windows = get_open_windows()  # Replace this with the actual function to retrieve window metadata
        
        # Extract window titles as unique identifiers
        current_window_titles = {window["title"] for window in windows}

        # Handle new windows: Add to assignments as "Unassigned" and track in new_assignments
        new_window_titles = current_window_titles - set(last_assignments.keys())
        for new_window_title in new_window_titles:
            if new_window_title not in assignments:
                new_assignments[new_window_title] = "Unassigned"

                # Add to the "Newly Detected Windows" Treeview in a thread-safe manner
                if not new_windows_tree.exists(new_window_title):
                    new_windows_tree.insert("", tk.END, iid=new_window_title, values=(new_window_title,))

        # Handle closed windows: Remove from assignments and last_assignments
        closed_window_titles = set(last_assignments.keys()) - current_window_titles
        for closed_window_title in closed_window_titles:
            last_assignments.pop(closed_window_title, None)

            # Remove closed windows from the new windows tree
            if new_windows_tree.exists(closed_window_title):
                new_windows_tree.delete(closed_window_title)

        # Group and update the left-hand Treeview (group by project)
        assignment_groups = {}
        metadata_by_project = {}
        for window_title, project in assignments.items():
            if project not in assignment_groups:
                assignment_groups[project] = []
                metadata_by_project[project] = {}
            assignment_groups[project].append(window_title)

            # Collect metadata for each window
            window_metadata = next((w for w in windows if w["title"] == window_title), {})
            metadata_by_project[project][window_title] = window_metadata

        # Update the left-hand Treeview
        for project, group_windows in assignment_groups.items():
            # Add project group if missing
            if not tree.exists(project):
                tree.insert("", tk.END, iid=project, text=project, open=True)

            # Add windows to the project group
            existing_items = set(tree.get_children(project))
            for window_title in group_windows:
                if window_title not in existing_items:
                    tree.insert(project, tk.END, iid=window_title, text=window_title)

        # Remove items no longer present in the current assignments
        for project in tree.get_children():
            for window_title in tree.get_children(project):
                if window_title not in assignments:
                    tree.delete(window_title)

        # Save updated assignments and project programs
        save_top_assignments(assignments)
        for project, metadata in metadata_by_project.items():
            if project != "Unassigned":
                save_project_programs(project, metadata)

        # Update last_assignments to match the current state
        last_assignments.clear()
        last_assignments.update(assignments)

        # Reset new_assignments for the next cycle
        new_assignments.clear()

        time.sleep(2)





# GUI setup
def setup_gui():
    projects = get_project_list()  # Replace with your actual function to get project names

    root = tk.Tk()
    root.title("Window Project Manager")
    root.geometry("1200x800")

    active_project = tk.StringVar()

    # Left Frame: Assigned Windows by Project
    left_frame = tk.Frame(root)
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    left_label = tk.Label(left_frame, text="Assigned Windows by Project", font=("Arial", 12, "bold"))
    left_label.pack(anchor="w", padx=10, pady=5)

    tree = ttk.Treeview(left_frame, show="tree", selectmode="browse")
    tree.heading("#0", text="Assigned Windows by Project")
    tree.pack(fill=tk.BOTH, expand=True)

    # Right Frame: Newly Detected Windows
    right_frame = tk.Frame(root)
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    active_project_label = tk.Label(right_frame, text="Active Project:", font=("Arial", 12, "bold"))
    active_project_label.pack(anchor="w", padx=10, pady=5)

    active_project_dropdown = ttk.Combobox(right_frame, textvariable=active_project)
    active_project_dropdown["values"] = projects
    active_project_dropdown.pack(pady=5)

    right_label = tk.Label(right_frame, text="Newly Detected Windows", font=("Arial", 12, "bold"))
    right_label.pack(anchor="w", padx=10, pady=5)

    # Configure the Treeview for new windows
    new_windows_tree = ttk.Treeview(right_frame, show="tree", selectmode="browse")
    new_windows_tree.heading("#0", text="Newly Detected Windows")
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

        # Function to assign/reassign selected windows to a project
    def assign_project():
        selected_items = new_windows_tree.selection() + tree.selection()  # Combine selections from both tables
        if selected_items and selected_project.get():
            project_name = selected_project.get()

            # Collect metadata for all assigned windows
            project_metadata = {}

            for item in selected_items:
                metadata = {}

                # Check if the item exists in the new windows tree
                if new_windows_tree.exists(item):
                    # Retrieve metadata before removing the item
                    metadata_json = new_windows_tree.item(item, "values")
                    if len(metadata_json) < 2:
                        print(f"Error: Metadata missing for item {item}, values: {metadata_json}")
                        continue
                    metadata = json.loads(metadata_json[1])

                    # Remove from the new windows tree
                    new_windows_tree.delete(item)
                elif tree.exists(item):
                    # Retrieve metadata from the assigned windows tree
                    for child in tree.get_children(item):
                        text = tree.item(child, "text")
                        if ": " in text:
                            key, value = text.split(": ", 1)
                            metadata[key] = value
                        else:
                            print(f"Skipping child node with unexpected text format: {text}")

                    # Remove the window from its current project group
                    for project in tree.get_children():
                        if item in tree.get_children(project):
                            tree.delete(item)
                            break
                else:
                    continue  # Skip if the item doesn't exist in either tree

                # Update assignments dictionary
                assignments[item] = project_name

                # Add the window metadata to the project metadata
                project_metadata[item] = metadata

                # Add the window to the new project group
                if not tree.exists(project_name):
                    tree.insert("", tk.END, iid=project_name, text=project_name, open=True)

                # Add the window title under the project group
                tree.insert(project_name, tk.END, iid=item, text=item)

                # Add metadata as child nodes
                for key, value in metadata.items():
                    tree.insert(item, tk.END, text=f"{key}: {value}")

            # Save updated assignments
            save_top_assignments(assignments)

            # Save project-specific programs with metadata
            save_project_programs(project_name, project_metadata)



    # Assign button
    assign_button = ttk.Button(root, text="Assign to Project", command=assign_project)
    assign_button.pack(pady=5)

    # Function to populate new windows with metadata
    def populate_new_windows():
        windows = get_open_windows()  # Replace with your actual function to get window metadata
        for window in windows:
            metadata_json = json.dumps(window, indent=2)

            # Generate a unique identifier for the Treeview item
            unique_id = f"{window['title']}_{window.get('pid', '')}"

            # Check if the item already exists in the Treeview
            if not new_windows_tree.exists(unique_id):
                # Insert the window title as a parent item
                new_windows_tree.insert("", tk.END, iid=unique_id, text=window["title"], values=(window["title"], metadata_json))

                # Insert metadata as child items
                for key, value in window.items():
                    new_windows_tree.insert(unique_id, tk.END, text=f"{key}: {value}")

    # Function to populate the assigned windows tree
    def populate_assigned_windows():
        for project, group_windows in assignments.items():
            # Create project group if not exists
            if not tree.exists(project):
                tree.insert("", tk.END, iid=project, text=project, open=True)

            # Add windows under the project group
            for window_title, metadata in group_windows.items():
                if not tree.exists(window_title):
                    tree.insert(project, tk.END, iid=window_title, text=window_title)

                # Add metadata as child nodes
                for key, value in metadata.items():
                    tree.insert(window_title, tk.END, text=f"{key}: {value}")

    # Close Project button
    close_button = ttk.Button(right_frame, text="Close Active Project", command=lambda: close_active_project(active_project.get(), assignments, tree))
    close_button.pack(pady=5)

    # Start the real-time update thread
    update_thread = Thread(
        target=update_assignments,
        args=(tree, new_windows_tree, assignments, last_assignments, new_assignments),
        daemon=True  # Ensure the thread stops when the main program exits
    )
    update_thread.start()

    # Populate new windows initially
    populate_new_windows()

    root.mainloop()

if __name__ == "__main__":
    setup_gui()