# Define a function to check if the project was created successfully
function Test-ProjectCreation {
    param (
        [string]$ProjectPath
    )

    if (Test-Path -Path $ProjectPath) {
        Write-Host "Success: The project directory '$ProjectPath' was created." -ForegroundColor Green
        return $true
    } else {
        Write-Host "Error: The project directory '$ProjectPath' could not be created." -ForegroundColor Red
        return $false
    }
}

# Prompt the user for the new project name
$newProjectName = Read-Host "Enter the project name"

# Define the path to the Projects directory
$basePath = "C:\Users\ScottMason(Qrometric\OneDrive - Qrometric\Projects"

# Validate the base path exists
if (-not (Test-Path -Path $basePath)) {
    Write-Host "Error: Base path '$basePath' does not exist. Please check the directory." -ForegroundColor Red
    exit 1
}

# Combine the base path and the new project name
$newProjectPath = Join-Path -Path $basePath -ChildPath $newProjectName

# Define the path to the project template
$templatePath = "C:\Users\ScottMason(Qrometric\OneDrive - Qrometric\Projects\Templates\ProjectTemplate001"

# Validate the template path exists
if (-not (Test-Path -Path $templatePath)) {
    Write-Host "Error: Template path '$templatePath' does not exist. Please check the directory." -ForegroundColor Red
    exit 1
}

# Check if the project directory already exists
if (Test-Path -Path $newProjectPath) {
    Write-Host "Error: A project with the name '$newProjectName' already exists at '$newProjectPath'." -ForegroundColor Yellow
    exit 1
}

# Attempt to copy the project template to the new project directory
try {
    Copy-Item -Recurse -Path $templatePath -Destination $newProjectPath -ErrorAction Stop
    if (Test-ProjectCreation -ProjectPath $newProjectPath) {
        Write-Host "New project created successfully at $newProjectPath" -ForegroundColor Green
    } else {
        Write-Host "Error: The project directory was not created properly." -ForegroundColor Red
    }
} catch {
    Write-Host "Error: Failed to copy the project template to '$newProjectPath'. Details: $_" -ForegroundColor Red
    exit 1
}