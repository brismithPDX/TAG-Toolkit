#===========================================================================
# Set up Dependencies and Variables
#===========================================================================
param($ItemPath)
#Core Configuration Variables
[bool]$Dev = $false

#Hide Script windows from user
if($Dev){
    $WindowControl = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
    add-type -name win -member $WindowControl -namespace native
    [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
}
#Create Windows Shell object for handling popup messages
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("FilePath is: $ItemPath",0,"Done",0x0)
