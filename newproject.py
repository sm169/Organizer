import os
import shutil
import subprocess
import requests


# Define the base and template paths
BASE_PATH = r"C:\Users\ScottMason(Qrometric\OneDrive - Qrometric\Projects"
TEMPLATE_PATH = r"C:\Users\ScottMason(Qrometric\OneDrive - Qrometric\Projects\Templates\ProjectTemplate002"

# Function to check if the project directory was created successfully
def test_project_creation(project_path):
    if os.path.exists(project_path):
        print(f"Success: The project directory '{project_path}' was created.")
        return True
    else:
        print(f"Error: The project directory '{project_path}' could not be created.")
        return False

# Prompt the user for the new project name
new_project_name = input("Enter the project name: ").strip()

# Validate base path
if not os.path.exists(BASE_PATH):
    print(f"Error: Base path '{BASE_PATH}' does not exist. Please check the directory.")
    exit(1)

# Combine the base path and the new project name
new_project_path = os.path.join(BASE_PATH, new_project_name)

# Validate template path
if not os.path.exists(TEMPLATE_PATH):
    print(f"Error: Template path '{TEMPLATE_PATH}' does not exist. Please check the directory.")
    exit(1)

# Check if the project directory already exists
if os.path.exists(new_project_path):
    print(f"Error: A project with the name '{new_project_name}' already exists at '{new_project_path}'.")
    exit(1)

# Attempt to copy the project template to the new project directory
try:
    shutil.copytree(TEMPLATE_PATH, new_project_path)
    if test_project_creation(new_project_path):
        print(f"New project created successfully at {new_project_path}")
    else:
        print("Error: The project directory was not created properly.")
except Exception as e:
    print(f"Error: Failed to copy the project template to '{new_project_path}'. Details: {e}")
    exit(1)

# Function to initialize a Git repository
def initialize_git_repo(project_path):
    try:
        subprocess.run(["git", "-C", project_path, "init"], check=True)
        print(f"Initialized a new Git repository in {project_path}")
    except subprocess.CalledProcessError as e:
        print(f"Error initializing Git repository: {e}")

# Function to link to a remote repository
def link_to_remote_repo(project_path, remote_url):
    if remote_url:
        try:
            subprocess.run(["git", "-C", project_path, "remote", "add", "origin", remote_url], check=True)
            print(f"Linked the project to remote repository: {remote_url}")
        except subprocess.CalledProcessError as e:
            print(f"Error linking to remote repository: {e}")

def create_github_repo(repo_name):
    github_token = "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" #os.getenv("GITHUB_TOKEN")
    if not github_token:
        print("Error: GitHub token not found. Set the GITHUB_TOKEN environment variable.")
        return None

    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json"
    }
    data = {
        "name": repo_name,
        "private": True
    }
    response = requests.post("https://api.github.com/user/repos", headers=headers, json=data)
    if response.status_code == 201:
        repo_url = response.json().get("html_url")
        print(f"Repository created: {repo_url}")
        return repo_url
    else:
        print(f"Error creating repository: {response.json()}")
        return None

# Initialize Git repository and link to remote if needed
initialize_git_repo(new_project_path)
repo_url = create_github_repo(new_project_name)
link_to_remote_repo(new_project_path,repo_url)

print("Project setup complete.")
