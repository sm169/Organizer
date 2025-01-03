# Load required libraries
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Windows API for enumerating windows
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class WinApi {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
}
"@

# Global Variables
$projectsPath = "C:\Users\ScottMason(Qrometric\OneDrive - Qrometric\Projects"
$activeProject = $null
$lastInteractedWindow = $null

# Function to get extra metadata from Chrome tabs
function Get-ChromeMetadata {
    $chromeDebugPort = 9222
    try {
        $endpoints = Invoke-RestMethod -Uri "http://localhost:$chromeDebugPort/json"
        return $endpoints | Where-Object { $_.type -eq "page" } | ForEach-Object {
            @{
                Title = $_.title
                URL   = $_.url
                DebugInfo = $_
            }
        }
    } catch {
        Write-Host "Chrome metadata could not be retrieved."
        return @()
    }
}

# Function to get extra metadata from VSCode windows
function Get-VSCodeMetadata {
    $vscodeProcesses = Get-Process code | Where-Object { $_.MainWindowTitle }
    return $vscodeProcesses | ForEach-Object {
        @{
            Title = $_.MainWindowTitle
            PID   = $_.Id
        }
    }
}

# Function to get all windows with detailed info
function Get-AllWindows {
    [System.Collections.ArrayList]$windows = @()
    $callback = [WinApi+EnumWindowsProc]{
        param($hWnd, $lParam)
        $title = New-Object System.Text.StringBuilder 256
        [WinApi]::GetWindowText($hWnd, $title, $title.Capacity) | Out-Null

        if ([WinApi]::IsWindowVisible($hWnd) -and $title.ToString().Trim()) {
            $windows.Add([PSCustomObject]@{
                Handle = $hWnd
                Title  = $title.ToString().Trim()
            })
        }
        return $true
    }
    [WinApi]::EnumWindows($callback, [IntPtr]::Zero)
    return $windows
}

# Function to display a GUI for assigning windows to projects
function Show-AssignmentDialog {
    param($windows, $projects, $existingAssignments)

    # Ensure $existingAssignments is initialized
    if (-not $existingAssignments) {
        $existingAssignments = @{}
    }

    # Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Assign Windows to Projects"
    $form.Size = New-Object System.Drawing.Size(800, 600)

    # Create the list view
    $listView = New-Object System.Windows.Forms.ListView
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.Columns.Add("Window Title", 400)
    $listView.Columns.Add("Project", 200)
    $listView.Location = New-Object System.Drawing.Point(10, 10)
    $listView.Size = New-Object System.Drawing.Size(560, 500)

    # Populate windows in the list view
    foreach ($window in $windows) {
        $item = New-Object System.Windows.Forms.ListViewItem($window.Title)
        if (-not [string]::IsNullOrWhiteSpace($window.Title) -and $existingAssignments.ContainsKey($window.Title)) {
            $assignedProject = $existingAssignments[$window.Title]
            $item.SubItems.Add($assignedProject)
        } else {
            $item.SubItems.Add("Unassigned")
        }
        $listView.Items.Add($item)
    }

    # Create the combo box for project selection
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = New-Object System.Drawing.Point(580, 10)
    $comboBox.Size = New-Object System.Drawing.Size(200, 30)
    $projects | ForEach-Object { $comboBox.Items.Add($_) }
    $comboBox.Items.Add("Skip All")

    # Assign button
    $btnAssign = New-Object System.Windows.Forms.Button
    $btnAssign.Location = New-Object System.Drawing.Point(580, 50)
    $btnAssign.Size = New-Object System.Drawing.Size(200, 30)
    $btnAssign.Text = "Assign Selected"
    $btnAssign.Add_Click({
        foreach ($item in $listView.SelectedItems) {
            if ($comboBox.SelectedItem -eq "Skip All") {
                $item.SubItems[1].Text = "Skipped"
            } else {
                $item.SubItems[1].Text = $comboBox.SelectedItem
            }
        }
    })

    # Save button (does not close the GUI)
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(580, 500)
    $btnSave.Size = New-Object System.Drawing.Size(200, 30)
    $btnSave.Text = "Save Assignments"
    $btnSave.Add_Click({
        $assignments = @{}
        foreach ($item in $listView.Items) {
            $assignments[$item.Text] = $item.SubItems[1].Text
        }
        Save-FilteredAssignments -assignments $assignments -projectsPath $projectsPath -projects $projects
        Write-Host "Assignments saved successfully."
    })

    # Add controls to the form
    $form.Controls.AddRange(@($listView, $comboBox, $btnAssign, $btnSave))

    # Show the form
    $form.ShowDialog() | Out-Null
}

function Show-ProjectManagementDialog {
    # Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Project Management"
    $form.Size = New-Object System.Drawing.Size(400, 300)

    # Button: Close Active Project
    $btnCloseProject = New-Object System.Windows.Forms.Button
    $btnCloseProject.Location = New-Object System.Drawing.Point(50, 50)
    $btnCloseProject.Size = New-Object System.Drawing.Size(300, 30)
    $btnCloseProject.Text = "Close Active Project"
    $btnCloseProject.Add_Click({
        Close-ActiveProject
    })

    # Button: Set Active Project
    $btnSetActive = New-Object System.Windows.Forms.Button
    $btnSetActive.Location = New-Object System.Drawing.Point(50, 100)
    $btnSetActive.Size = New-Object System.Drawing.Size(300, 30)
    $btnSetActive.Text = "Set Active Project"
    $btnSetActive.Add_Click({
        $newProject = Read-Host "Enter new active project"
        Set-ActiveProject -newProject $newProject
    })

    # Button: Track Last Interacted Window
    $btnTrackWindow = New-Object System.Windows.Forms.Button
    $btnTrackWindow.Location = New-Object System.Drawing.Point(50, 150)
    $btnTrackWindow.Size = New-Object System.Drawing.Size(300, 30)
    $btnTrackWindow.Text = "Track Last Interacted Window"
    $btnTrackWindow.Add_Click({
        Track-LastInteractedWindow
    })

    # Button: Add Last Interacted Window to Active Project
    $btnAddWindow = New-Object System.Windows.Forms.Button
    $btnAddWindow.Location = New-Object System.Drawing.Point(50, 200)
    $btnAddWindow.Size = New-Object System.Drawing.Size(300, 30)
    $btnAddWindow.Text = "Add Last Interacted Window to Active Project"
    $btnAddWindow.Add_Click({
        Add-LastInteractedToActive
    })

    # Add controls to the form
    $form.Controls.AddRange(@($btnCloseProject, $btnSetActive, $btnTrackWindow, $btnAddWindow))

    # Show the form
    $form.ShowDialog() | Out-Null
}



# Function to save assignments to project JSON
function Save-FilteredAssignments {
    param($assignments, $projectsPath, $projects)

    foreach ($project in $projects) {
        $projectPath = Join-Path -Path $projectsPath -ChildPath $project
        $assignmentFilePath = Join-Path -Path $projectPath -ChildPath "assignments.json"

        # Filter assignments for this project
        $filteredAssignments = @{}
        foreach ($key in $assignments.Keys) {
            $assignedProject = $assignments.$key
            if ($assignedProject -eq $project) {
                $filteredAssignments[$key] = $assignedProject
            }
        }

        # Save filtered assignments to the JSON file
        if ($filteredAssignments.Count -gt 0) {
            $filteredAssignments | ConvertTo-Json -Depth 10 | Set-Content -Path $assignmentFilePath
            Write-Host "Saved $($filteredAssignments.Count) assignments to $assignmentFilePath"
        } else {
            @{} | ConvertTo-Json -Depth 10 | Set-Content -Path $assignmentFilePath
            Write-Host "Cleared assignments for project: $project"
        }
    }
}

# Function to handle closing a project
function Close-ActiveProject {
    if ($activeProject) {
        Write-Host "Closing active project: $activeProject"

        # Move laststate to savedstates
        $projectPath = Join-Path -Path $projectsPath -ChildPath $activeProject
        $lastStatePath = Join-Path -Path $projectPath -ChildPath "laststate"
        $savedStatesPath = Join-Path -Path $projectPath -ChildPath "savedstates"
        if (-Not (Test-Path $savedStatesPath)) { New-Item -Path $savedStatesPath -ItemType Directory }
        $timestamp = (Get-Date -Format "yyyyMMdd_HHmmss")
        $destination = Join-Path -Path $savedStatesPath -ChildPath "state_$timestamp"
        Move-Item -Path $lastStatePath -Destination $destination -Force

        # Prompt for closing notes
        $notesFilePath = Join-Path -Path $destination -ChildPath "closing_notes.txt"
        Write-Host "Please add closing notes..."
        notepad.exe $notesFilePath

        # Upload files to Git if configured
        $configPath = Join-Path -Path $projectPath -ChildPath "config.json"
        if (Test-Path $configPath) {
            $config = Get-Content $configPath | ConvertFrom-Json
            if ($config.gitRepo) {
                git -C $projectPath add .
                git -C $projectPath commit -m "Closing project with notes."
                git -C $projectPath push
            }
        }
    }
}

# Function to set a new active project
function Set-ActiveProject {
    param($newProject)
    if ($newProject -ne $activeProject) {
        Close-ActiveProject
        $activeProject = $newProject
        Write-Host "Active project set to: $activeProject"
    }
}

# Function to track the last clicked/interacted window
function Track-LastInteractedWindow {
    $foregroundWindowHandle = [WinApi]::GetForegroundWindow()
    $windows = Get-AllWindows
    $lastInteractedWindow = $windows | Where-Object { $_.Handle -eq $foregroundWindowHandle }
    if ($lastInteractedWindow) {
        Write-Host "Last interacted window: $($lastInteractedWindow.Title)"
    }
}

# Function to add last interacted window to active project
function Add-LastInteractedToActive {
    if ($lastInteractedWindow -and $activeProject) {
        $projectPath = Join-Path -Path $projectsPath -ChildPath $activeProject
        $assignmentsFile = Join-Path -Path $projectPath -ChildPath "assignments.json"
        $currentAssignments = @{}
        if (Test-Path $assignmentsFile) {
            $currentAssignments = Get-Content $assignmentsFile | ConvertFrom-Json
        }

        if (-not $currentAssignments.ContainsKey($lastInteractedWindow.Title)) {
            $currentAssignments[$lastInteractedWindow.Title] = @{
                Handle = $lastInteractedWindow.Handle
            }
            $currentAssignments | ConvertTo-Json -Depth 10 | Set-Content -Path $assignmentsFile
            Write-Host "Added $($lastInteractedWindow.Title) to active project: $activeProject"
        } else {
            Write-Host "Window $($lastInteractedWindow.Title) is already assigned to a project."
        }
    } else {
        Write-Host "No active project or no last interacted window to add."
    }
}

# Main monitoring function
function Start-Monitoring {
    $projects = Get-ChildItem -Path $projectsPath -Directory | Select-Object -ExpandProperty Name
    $assignments = @{}
    $lastCollectedWindows = @()  # Maintain a list of previously collected windows

    while ($true) {
        $windows = Get-AllWindows

        foreach ($window in $windows) {
            # Skip if the window is already in the last collected list
            if ($lastCollectedWindows -contains $window.Title) {
                continue
            }

            $assignedProjects = @()

            # Check if the window is already assigned to any project
            foreach ($project in $projects) {
                $projectPath = Join-Path -Path $projectsPath -ChildPath $project
                $assignmentsFile = Join-Path -Path $projectPath -ChildPath "assignments.json"
                if (Test-Path $assignmentsFile) {
                    $projectAssignments = Get-Content $assignmentsFile | ConvertFrom-Json

                    # Ensure the object is treated as a hashtable
                    if (-not ($projectAssignments -is [hashtable])) {
                        $projectAssignments = @{}
                        foreach ($key in $projectAssignments.PSObject.Properties.Name) {
                            $projectAssignments[$key] = $projectAssignments.PSObject.Properties[$key].Value
                        }
                    }

                    # Check if the title exists in the assignments
                    if ($projectAssignments.PSObject.Properties.Name -contains $window.Title) {
                        $assignedProjects += $project
                    }
                }
            }

            # If not assigned, prompt the user
            if ($assignedProjects.Count -eq 0) {
                $response = Read-Host "Assign window '$($window.Title)' to active project ($activeProject)? (y/n/skip)"
                if ($response -eq 'y' -and $activeProject) {
                    $assignedProjects += $activeProject
                } elseif ($response -eq 'skip') {
                    break
                }
            }
            # Add to assignments
            $assignments[$window.Title] = $assignedProjects
        }

        Write-Host "Launching assignment dialog..."
        $assignments = Show-AssignmentDialog -windows $windows -projects $projects
        

        # Save updated assignments
        Save-FilteredAssignments -assignments $assignments -projectsPath $projectsPath -projects $projects

        # Update the list of collected windows
        $lastCollectedWindows = $windows.Title  # Update with the titles of current windows

        Start-Sleep -Seconds 20
    }
}

# Start the GUI for project management
Show-ProjectManagementDialog

# Start the monitoring process
Start-Monitoring

