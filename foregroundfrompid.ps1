# Add necessary .NET types for interacting with user32.dll
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class User32 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    public const int SW_RESTORE = 9; // Restore the window if minimized
}
"@

# Function to get the handle (hWnd) of a window by its PID
function Get-WindowHandleByPID {
    param (
        [int]$ProcessID
    )

    $TargetWindowHandle = [IntPtr]::Zero

    # Callback function to check each window
    $callback = {
        param ($hWnd, $lParam)
        $processId = 0
        [User32]::GetWindowThreadProcessId($hWnd, [ref]$processId)
        if ($processId -eq $lParam) {
            # Found the window, return its handle
            $script:TargetWindowHandle = $hWnd
            return $false # Stop enumeration
        }
        return $true # Continue enumeration
    }

    # Wrap the callback in a .NET delegate
    $enumProc = [User32+EnumWindowsProc]$callback

    # Enumerate all top-level windows
    [User32]::EnumWindows($enumProc, $ProcessID)

    return $TargetWindowHandle
}

# Prompt the user to enter the PID
$ProcessID = Read-Host "Enter the PID of the process whose window you want to bring to the foreground"

# Get the window handle for the provided PID
$WindowHandle = Get-WindowHandleByPID -ProcessID $ProcessID

if ($WindowHandle -eq [IntPtr]::Zero) {
    Write-Host "No window found for the given PID."
} else {
    # Restore the window if it's minimized
    [User32]::ShowWindow($WindowHandle, [User32]::SW_RESTORE)

    # Bring the window to the foreground
    if ([User32]::SetForegroundWindow($WindowHandle)) {
        Write-Host "Window brought to the foreground successfully."
    } else {
        Write-Host "Failed to bring the window to the foreground."
    }
}
