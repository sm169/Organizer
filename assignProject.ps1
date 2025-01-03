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

# Function to get all windows with detailed info
function Get-AllWindows {
    [System.Collections.ArrayList]$windows = @()
    $callback = [WinApi+EnumWindowsProc]{
        param($hWnd, $lParam)
        $title = New-Object System.Text.StringBuilder 256
        #$className = New-Object System.Text.StringBuilder 256

        [WinApi]::GetWindowText($hWnd, $title, $title.Capacity) | Out-Null
        #[WinApi]::GetClassName($hWnd, $className, $className.Capacity) | Out-Null

        if ([WinApi]::IsWindowVisible($hWnd) -and $title.ToString().Trim()) {
            $null = $windows.Add([PSCustomObject]@{
                Handle     = $hWnd
                Title      = $title.ToString().Trim()
                #ClassName  = $className.ToString().Trim()
                PID        = (Get-Process | Where-Object { $_.MainWindowHandle -eq $hWnd }).Id
                Visible    = [WinApi]::IsWindowVisible($hWnd)
            })
        }
        return $true
    }
    [WinApi]::EnumWindows($callback, [IntPtr]::Zero)
    return $windows
}

# Function to show assignment dialog
function Show-AssignmentDialog {
    param($windows, $projects)

    # Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Assign Programs to Projects"
    $form.Size = New-Object System.Drawing.Size(800, 600)

    # Create the list view
    $listView = New-Object System.Windows.Forms.ListView
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.Columns.Add("Program", 300)
    $listView.Columns.Add("Project/Tag", 200)
    $listView.Location = New-Object System.Drawing.Point(10, 10)
    $listView.Size = New-Object System.Drawing.Size(560, 500)

    # Populate windows in the list view
    foreach ($window in $windows) {
        $item = New-Object System.Windows.Forms.ListViewItem($window.Title)
        $item.SubItems.Add("Unassigned")
        $listView.Items.Add($item)
    }

    # Create the combo box for project selection
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = New-Object System.Drawing.Point(580, 10)
    $comboBox.Size = New-Object System.Drawing.Size(200, 30)
    $projects | ForEach-Object { $comboBox.Items.Add($_) }
    $comboBox.Items.Add("General")

    # Assign button
    $btnAssign = New-Object System.Windows.Forms.Button
    $btnAssign.Location = New-Object System.Drawing.Point(580, 50)
    $btnAssign.Size = New-Object System.Drawing.Size(200, 30)
    $btnAssign.Text = "Assign Selected"
    $btnAssign.Add_Click({
        foreach ($item in $listView.SelectedItems) {
            $item.SubItems[1].Text = $comboBox.SelectedItem
        }
    })

    # Save button
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(580, 500)
    $btnSave.Size = New-Object System.Drawing.Size(200, 30)
    $btnSave.Text = "Save Assignments"
    $btnSave.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # Add controls to the form
    $form.Controls.AddRange(@($listView, $comboBox, $btnAssign, $btnSave))

    # Show dialog and process results
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $assignments = @{}
        foreach ($item in $listView.Items) {
            $assignments[$item.Text] = $item.SubItems[1].Text
        }
        return $assignments
    }
    return $null
}

# Function to save filtered assignments per project
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

# Main monitoring function
function Start-Monitoring {
    $projectsPath = "C:\Users\ScottMason(Qrometric\OneDrive - Qrometric\Projects"
    $projects = Get-ChildItem -Path $projectsPath -Directory | Select-Object -ExpandProperty Name

    while ($true) {
        Write-Host "Fetching current windows..."
        $windows = Get-AllWindows

        Write-Host "Launching assignment dialog..."
        $assignments = Show-AssignmentDialog -windows $windows -projects $projects

        if ($assignments) {
            Save-FilteredAssignments -assignments $assignments -projectsPath $projectsPath -projects $projects
        }

        Start-Sleep -Seconds 120  # Wait 2 minutes before next loop
    }
}

# Start monitoring
Start-Monitoring
