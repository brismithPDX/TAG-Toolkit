<##>

#Hide Script windows from user
$WindowControl = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $WindowControl -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

#===========================================================================
# Create XAML Objects In PowerShell
#===========================================================================
$inputXML =@"
<Window x:Class="WpfApplication2.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApplication2"
        mc:Ignorable="d"
        Title="TAG Troubleshooting Data Collection Utility" Height="350" Width="525">
    <Grid Margin="0,0,0,-2">
        <Label x:Name="label" Content="RT Ticket Number:" HorizontalAlignment="Left" Margin="22,65,0,0" VerticalAlignment="Top"/>
        <Label x:Name="label1" Content="Description of Problem:" HorizontalAlignment="Left" Margin="22,91,0,0" VerticalAlignment="Top" Width="181"/>
        <TextBox x:Name="RTnumber" HorizontalAlignment="Left" Height="21" Margin="134,70,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="148"/>
        <TextBox x:Name="problemDescription" HorizontalAlignment="Left" Height="148" Margin="27,117,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="465"/>
        <Button x:Name="cancel" Content="Cancel" HorizontalAlignment="Left" Margin="329,284,0,0" VerticalAlignment="Top" Width="75"/>
        <Button x:Name="continue" Content="Continue" HorizontalAlignment="Left" Margin="417,284,0,0" VerticalAlignment="Top" Width="75"/>
        <TextBlock x:Name="textBlock" HorizontalAlignment="Left" Height="50" Margin="27,10,0,0" TextWrapping="Wrap" Text="This utility will collect all computer log information and upload the items to a open troubleshooting ticket. Please note the logs may be quite large and collection may take 5-10 minutes." VerticalAlignment="Top" Width="465"/>
        <TextBlock x:Name="errorBlock" HorizontalAlignment="Left" Height="39" Margin="22,265,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="291" Foreground="Red"/>

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
function ExecuteMain{
 ##Validate both fields are filled.
 [bool] $errorCheck = $false
 $WPFerrorBlock.Text = ""
 if($WPFRTnumber.Text -eq "" -or $WPFRTnumber.Text -eq $null){
    $WPFerrorBlock.Text = $WPFErrorBlock.Text+"`nError: No RT Number Provided..."
    $errorCheck = $true
 }
 if($WPFRTnumber.Text.Length -lt 6){
    $WPFerrorBlock.Text = $WPFErrorBlock.Text+"`nError: Invalid RT Number..."
    $errorCheck = $true
 }
 if($WPFproblemDescription.Text -eq "" -or $WPFproblemDescription.Text -eq $null){
    $WPFerrorBlock.Text = $WPFErrorBlock.Text+"`nError: No Description Provided..."
    $errorCheck = $true
 }
 if($errorCheck){
    return
 }
 ##Collect SCCM, System, & McAfee logs
 $StagingArea = "$env:TEMP\LogCollectionUtility"
 New-Item $StagingArea -ItemType directory

 Copy-Item $env:windir\System32\winevt\Logs\Application.evtx $StagingArea
 Copy-Item $env:windir\System32\winevt\Logs\microsoft-windows-mbam%4operational.evtx $StagingArea
 Copy-Item $env:windir\System32\winevt\Logs\microsoft-windows-mbam%4admin.evtx $StagingArea
 Copy-Item $env:windir\System32\winevt\Logs\Microsoft-Windows-PowerShell%4Admin.evtx $StagingArea
 Copy-Item $env:windir\System32\winevt\Logs\Microsoft-Windows-PowerShell%4Operational.evtx $StagingArea
 Copy-Item $env:windir\System32\winevt\Logs\Windows PowerShell.evtx $StagingArea
 Copy-Item $env:windir\CCM\Logs\AppEnforce.log $StagingArea
 Copy-Item $env:windir\CCM\Logs\AppDiscovery.log $StagingArea
 Copy-Item $env:windir\CCM\Logs\CcmEval.log $StagingArea
 Copy-Item $env:windir\CCM\Logs\CcmExec.log $StagingArea

 #Compress logs for shipping
 Compress-Archive -Path $StagingArea -DestinationPath $env:TEMP\TroubleshootingLogArchive.zip

 ##Send collection to RT
[string] $RTSubject = "TDCU - appending logs and user note"
[string] $RTMessage = $WPFproblemDescription.Text
 RTRequest -reqType "update"  -reqTicket $WPFRTnumber -reqSubject $RTSubject -reqMessage $RTMessage -reqFile $env:TEMP\TroubleshootingLogArchive.zip

  #Create Windows Shell object for handling popup messages
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Ticket Update compleated, Logs are appened to the ticket",0,"Done",0x0)
$form.Close()

}

## Setup button functions
$WPFcancel.Add_Click({
    $form.Close()})
$WPFcontinue.Add_Click({
    ExecuteMain
    return
    })

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null
