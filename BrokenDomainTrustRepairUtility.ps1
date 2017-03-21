#Insure Admin Applcation runtime Permissions

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

#Hide Script windows from user
$WindowControl = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $WindowControl -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

#===========================================================================
# Create XAML Objects In PowerShell
#===========================================================================
$inputXML =@"
<Window x:Class="WpfApplication1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApplication1"
        mc:Ignorable="d"
        Title="PSU - Broken Domain Trust Repair Utility" Height="140.505" Width="332.091" Background="{DynamicResource {x:Static SystemColors.ControlDarkDarkBrushKey}}">
    <Grid Background="{DynamicResource {x:Static SystemColors.ControlDarkDarkBrushKey}}">
        <Button x:Name="Start" Content="Fix It" HorizontalAlignment="Left" Height="65" Margin="93,23,0,0" VerticalAlignment="Top" Width="142"/>

    </Grid>
</Window>
"@

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
Get-FormVariables
#===========================================================================
# Actually make the objects work
#===========================================================================

function ExecuteMain{
 #Generate Required Data
 $creds = Get-Credential
 $username = $creds.UserName

 #log event to application log for splunk collection
 New-EventLog -LogName Application -Source "PSU - Broken Domain Trust Repair Utility"
 Write-EventLog -LogName Application -source "PSU - Broken Domain Trust Repair Utility" -EntryType Information -EventId 1 -Message "A Domain Trust Repair was Launched by $username"

 #Preform Repair
 Test-ComputerSecureChannel -Repair -Server [CHANGE-ME] -Credential $creds
 #Notify User of condition and exit
 
 #Create Windows Shell object for handling popup messages
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Repair Compleated, please verify new domian trust relationship",0,"Done",0x0)
$form.Close()
}

## Setup button functions
$WPFStart.Add_Click({
    ExecuteMain
    return
    })

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null