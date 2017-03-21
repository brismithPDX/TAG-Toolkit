#===========================================================================
# Set up Dependencies and Variables
#===========================================================================
# Take input from event viewer and task scheduler trigger system
param($eventRecordID)
#Core Configuration Variables
[bool]$Dev = $false

#Hide Script windows from user
if($Dev -eq $false){
    $WindowControl = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
    add-type -name win -member $WindowControl -namespace native
    [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
}

#===========================================================================
# Parse input arguments from event viewer and get event for editing
#===========================================================================

# Search for the event by the ID provided by trigger
$event = get-winevent -LogName "Microsoft-Windows-AppLocker/EXE and DLL" -FilterXPath "*[System[EventRecordID=$eventRecordID]]"

#convert event to XML for data extraction
[xml] $eventXML = $event.ToXml()


[string] $fileComputer = $eventXML.Event.System.Computer
[string] $fileHash = $eventXML.Event.UserData.RuleAndFileData.FileHash
[string] $filePath = $eventXML.Event.UserData.RuleAndFileData.FilePath
[string] $userSID = $eventXML.Event.UserData.RuleAndFileData.TargetUser
if($Dev){
    $fileComputer
    $fileHash
    $filePath
    $userSID
}

#translate use sid into user name for later use
$userObject = New-Object System.Security.Principal.SecurityIdentifier $userSID
$userName = $userObject.Translate([System.Security.Principal.NTAccount])

#skip compression just send hash to gui for delivery to RT

#===========================================================================
# Kick off AUET GUI Components
#===========================================================================
$inputXML =@"
<Window x:Class="WpfApplication1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApplication1"
        mc:Ignorable="d"
        Title="AUET 2016" Height="350" Width="550" Background="{DynamicResource {x:Static SystemColors.ControlDarkDarkBrushKey}}">
    <Grid Background="{DynamicResource {x:Static SystemColors.ControlDarkDarkBrushKey}}">
        <!--Header Text-->
        <TextBlock x:Name="title_Instructions" HorizontalAlignment="Center" Height="37" Margin="10,10,10,0" TextWrapping="Wrap" Text="Please verify the information below before sending the exemption request" VerticalAlignment="Top" Width="497"/>

        <!--Controls-->
        <Button x:Name="sendRequest" Content="Send Request" HorizontalAlignment="Left" Height="33" Margin="416,276,0,0" VerticalAlignment="Top" Width="91"/>
        <Button x:Name="cancel" Content="Cancel" HorizontalAlignment="Left" Height="33" Margin="320,276,0,0" VerticalAlignment="Top" Width="91"/>

        <!--Titles-->
        <TextBlock x:Name="title_UserName" HorizontalAlignment="Left" Height="17" Margin="10,46,0,0" TextWrapping="Wrap" Text="Username:" VerticalAlignment="Top" Width="104"/>
        <TextBlock x:Name="title_AppName" HorizontalAlignment="Left" Height="15" Margin="10,72,0,0" TextWrapping="Wrap" Text="Application Name:" VerticalAlignment="Top" Width="104"/>
        <TextBlock x:Name="title_InstallDirectory" HorizontalAlignment="Left" Height="15" Margin="10,98,0,0" TextWrapping="Wrap" Text="Install Directory:" VerticalAlignment="Top" Width="104"/>
        <TextBlock x:Name="title_Department" HorizontalAlignment="Left" Height="15" Margin="10,124,0,0" TextWrapping="Wrap" Text="Department:" VerticalAlignment="Top" Width="104"/>

        <!--User Inputs & Confermations-->
        <TextBox x:Name="input_UserName" HorizontalAlignment="Left" Height="21" Margin="114,46,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="285"/>
        <TextBox x:Name="input_AppName" HorizontalAlignment="Left" Height="21" Margin="114,72,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="285"/>
        <TextBox x:Name="input_InstallDirectory" HorizontalAlignment="Left" Height="21" Margin="114,98,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="285"/>
        <TextBox x:Name="input_Department" HorizontalAlignment="Left" Height="21" Margin="114,124,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="285"/>
        
        <!--Application Outputs-->
        <TextBlock x:Name="output_MessageBox" HorizontalAlignment="Left" Height="105" Margin="10,168,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="297"/>
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

#Print Available Form Details 
if($Dev){
    Get-FormVariables
}

#===========================================================================
# Core Functions for Application 
#===========================================================================
function RTRequest{
param(
    [string] $restUri, # PSURTRestURI

    [Parameter(Mandatory=$True)]
    [string] $reqType,
    [string] $userName, # PSURTRestUsername
    [string] $userPass, # PSURTRestPassword
    [string] $reqMessage,
    [string] $reqField,
    [string] $reqQueue,
    [string] $reqQuery,
    [string] $reqRequestor,
    [string] $reqSubject,
    [string] $reqTicket, # PSUTicketNumberHTA
    [string] $reqAction = "comment",
    [string] $reqStatus,
    [string] $reqFile   

)

$restUri = "[CHANGE-ME]"
$userName = "[CHANGE-ME]"
$userPass = "[CHANGE-ME]"

if ((($restUri -eq $null) -or ($restUri -eq "")) -or (($userName -eq $null) -or ($userName -eq "")) -or (($userPass -eq $null) -or ($userPass -eq "")))
{
    Write-host "Error: URI or Login Information missing. Cannot proceed"
    Exit(-1)
}

switch($reqType.ToLower())
{
   "create" 
   {
        if ((($reqQueue -ne $null) -and ($reqQueue -ne "")) -and (($reqRequestor -ne $null) -and ($reqRequestor -ne "")) -and (($reqSubject -ne $null) -and `
             ($reqSubject -ne "")) -and (($reqMessage -ne $null) -and ($reqMessage -ne "")))
        {
        
            $restUri += "ticket/new" + "?user=" + $userName + "&pass=" + $userPass
            $reqString = "content=Queue: " + $reqQueue + "`nRequestor: " + $reqRequestor+ "`nStatus: " + $reqStatus + "`nSubject: " + $reqSubject + "`nText: " + $reqMessage + "`n"

            $return = Invoke-RestMethod -Method Post -Uri $restUri -Body $reqString

            # Handle Feedback in $return variable

            Write-host $return
                
        }
        else
        {
            Write-host "Error: Invalid Parameters for Create Request Type."
        }
   }

   "custom" 
   {
        if ((($reqTicket -ne $null) -and ($reqTicket -ne "")) -and (($reqField -ne $null) -and ($reqField -ne "")) `
            -and (($reqMessage -ne $null) -and ($reqMessage -ne "")))
        {
        
            $restUri += "ticket/" + $reqTicket + "/edit" + "?user=" + $userName + "&pass=" + $userPass
            $reqString = "content=CF-" + $reqField + ": " + $reqMessage

            $return = Invoke-RestMethod -Method Post -Uri $restUri -Body $reqString

            # Handle Feedback in $return variable

            if ($return -match "does not apply")
            {

                $restUri = $restUri.Replace("edit","comment")
                
                $reqString = "content=Action: comment" + "`nText: " + $reqField + " = " + $reqMessage + "`nStatus: " + $reqStatus   

                $return = Invoke-RestMethod -Method Post -Uri $restUri -Body $reqString

                Write-Host $return

            }   
            else
            {
                Write-Host $return
            }
        }
        else
        {
            Write-host "Error: Invalid Parameters for Custom Request Type."
        }
   }

   "update"
   {
        if ((($reqTicket -ne $null) -and ($reqTicket -ne "")) -and (($reqMessage -ne $null) -and ($reqMessage -ne "")))
        {
        
            $restUri += "ticket/" + $reqTicket + "/comment" + "?user=" + $userName + "&pass=" + $userPass
            $reqString = "content=Action: " + $reqAction + "`nText: " + $reqMessage + "`nStatus: " + $reqStatus

            $return = Invoke-RestMethod -Method Post -Uri $restUri -Body $reqString

            # Handle Feedback in $return variable

            Write-Host $return

        }
        else
        {
            Write-host "Error: Invalid Parameters for Update Request Type."
        }

   }

   "search"
   {
        if (($reqQuery -ne $null) -and ($reqQuery -ne ""))
        {
            $restUri += "search/ticket" + "?user=" + $userName + "&pass=" + $userPass + "&query=" + $reqQuery.Replace(" ","+")

            $return = Invoke-RestMethod -Method Get -Uri $restUri

            # Handle Feedback in $return variable

            Write-Host $return

        }
        else
        {
            Write-host "Error: Invalid Parameters for Update Request Type."
        }

   }

   default
   {
        Write-Host "Error: Invalid Request Type Encountered = '$reqType'"
   }
}
}
function ValidateInputs{
    [bool] $Error = $false
    if($WPFinput_AppName.Text -eq ""){$Error = $true}
    if($WPFinput_Department.Text -eq ""){$Error = $true}
    if($WPFinput_InstallDirectory.Text -eq ""){$Error = $true}
    if($WPFinput_UserName.Text -eq ""){$Error = $true}
    return $Error
}

#===========================================================================
# Tie in Functions for XAML Objects
#===========================================================================
$WPFcancel.Add_Click({
    $form.Close()
})
$WPFsendRequest.Add_Click({
    #make sure all form content is compleated
    if(ValidateInputs){
        $WPFoutput_MessageBox.Text = "WARNING: A required field is empty, please insure all fields are compleated."
        $WPFoutput_MessageBox.Foreground = "DarkRed"
        return
    }
    #Build RT Messages for sending
    [string] $RTSubject = "Request for Applocker Ruleset Exemption in Department: " + $WPFinput_Department.Text
    [string] $RTMessage = "Request for Applocker Ruleset Exemption in Department: " + $WPFinput_Department.Text + "`n User: " + $WPFinput_UserName.Text + "`n Application Install Directory: " + $WPFinput_InstallDirectory.Text + "`n Application Hash: " + $fileHash + "`n Source Machine: " + $fileComputer + "`n EventRecordID: " + $eventRecordID
    [string] $RequestorID = $WPFinput_UserName.Text
    $RequestorID = $RequestorID.Substring(4)
    
    if($Dev){
        write-host "requestorID" + $RequestorID
    }

    #send Message
    RTRequest -reqType "create" -reqSubject $RTSubject -reqMessage $RTMessage -reqStatus "new" -reqQueue "cis-csa-package" -reqRequestor $RequestorID
    
    #close the form
    $form.Close()

    #pop up closing message
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Request Sent",0,"Done",0x0)
})
#===========================================================================
# Shows the form
#===========================================================================

#set form content
$WPFinput_InstallDirectory.Text = $filePath
$WPFinput_UserName.Text = $userName.Value
if($Dev){
    $WPFinput_InstallDirectory.Text
    $WPFinput_UserName.Text
}

$Form.ShowDialog() | out-null


#===========================================================================
# Clean up and exit
#===========================================================================
