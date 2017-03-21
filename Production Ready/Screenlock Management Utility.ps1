<##>
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
        Title="TAG ScreenLock Management Utility" Height="350" Width="525">
    <Grid>
        <Label x:Name="label" Content="User Name:" HorizontalAlignment="Left" Margin="22,28,0,0" VerticalAlignment="Top"/>
        <Label x:Name="label1" Content="Password:" HorizontalAlignment="Left" Margin="22,58,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="username" HorizontalAlignment="Left" Height="23" Margin="93,32,0,0" TextWrapping="Wrap" Text="TAG Username" VerticalAlignment="Top" Width="120"/>
        <PasswordBox x:Name="passwordBox" HorizontalAlignment="Left" Margin="93,62,0,0" VerticalAlignment="Top" Width="120" Password="defualt"/>
        <Label x:Name="label2" Content="Computer Name:" HorizontalAlignment="Left" Margin="22,130,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="computername" HorizontalAlignment="Left" Height="23" Margin="127,134,0,0" TextWrapping="Wrap" Text="Target Computer" VerticalAlignment="Top" Width="181"/>
        <Label x:Name="label3" Content="Set Security Level:" HorizontalAlignment="Left" Margin="22,158,0,0" VerticalAlignment="Top"/>
        <ComboBox x:Name="comboBox" HorizontalAlignment="Left" Margin="127,162,0,0" VerticalAlignment="Top" Width="120" SelectedIndex="0">
            <Button Content="PSU Standard" IsDefault="True"/>
            <Button Content="High Security"/>
            <Button Content="Exempt" IsEnabled="False"/>
        </ComboBox>
        <TextBox x:Name="changereason" HorizontalAlignment="Left" Height="106" Margin="27,195,0,0" TextWrapping="Wrap" Text="Reason for change..." VerticalAlignment="Top" Width="281"/>
        <Button x:Name="button" Content="Cancel" HorizontalAlignment="Left" Margin="333,285,0,0" VerticalAlignment="Top" Width="75"/>
        <Button x:Name="button1" Content="Apply" HorizontalAlignment="Left" Margin="422,285,0,0" VerticalAlignment="Top" Width="75"/>
        <TextBlock x:Name="textBlock" HorizontalAlignment="Left" Height="113" Margin="232,16,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="275"><Run Text="This utility will change the target computers Screen Lock Policy. Please remember the principles of least privilege and PSU screen lock policy when making your change."/><LineBreak/><Run/><LineBreak/><Run Text="All changes are monitored and recorded in RT"/></TextBlock>
        <TextBlock x:Name="ErrorBlock" HorizontalAlignment="Left" Height="78" Margin="333,195,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="152" Foreground="#FFFF0000" Text=""/>

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
function CleanGroups{
    param(
        [System.Management.Automation.CredentialAttribute()]
        $Credentals
    )
    try{
        $groups = Get-ADComputer $WPFcomputername.Text -Properties memberof | Select -ExpandProperty memberof
        $ADComputerName = $WPFcomputername.Text+"$"
        if($groups -like "*CSA_Security_ScreenSaverLockOverride_LG*"){
            Remove-ADGroupMember -Identity CSA_Security_ScreenSaverLockOverride_LG -Members $ADComputerName -Credential $Credentals -Confirm:$false
        }
        if($groups -like "*CSA_Security_ScreenSaverLock_Reduced_LG*"){
            Remove-ADGroupMember -Identity CSA_Security_ScreenSaverLock_Reduced_LG -Members $ADComputerName -Credential $Credentals -Confirm:$false
        }
        return
    }catch{
        $WPFErrorBlock.Text = $WPFErrorBlock.Text+"`nError: Could not clean up computers group memberships. Do you have the permissions required for this operation?"
        return
    }
}
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
    ####
    # Field Error Checking
    ####
    [bool] $errorCheck = $false
    $WPFErrorBlock.Text = ""
    if($WPFcomputername.Text -eq "Target Computer" -or $WPFcomputername.Text -eq $null -or $WPFcomputername.Text -eq ""){
        $WPFErrorBlock.Text = $WPFErrorBlock.Text+"`nError: No computer selected..."
        $errorCheck = $true
    }
    if($WPFchangereason.Text -eq "Reason for change..." -or $WPFchangereason.Text -eq $null -or $WPFchangereason.Text -eq ""){
        $WPFErrorBlock.Text = $WPFErrorBlock.Text+"`nError: A reason for the change must be given"
        $errorCheck = $true
    }
    if($WPFpasswordBox.Password -eq "defualt" -or $WPFpasswordBox.Password -eq $null -or $WPFpasswordBox.Password -eq ""){
        $WPFErrorBlock.Text = $WPFErrorBlock.Text+"`nError: No Password Given"
        $errorCheck = $true
    }
    if($WPFusername.Text -eq "Tag Username" -or $WPFusername.Text -eq $null -or $WPFusername.Text -eq ""){
        $WPFErrorBlock.Text = $WPFErrorBlock.Text+"`nError: No Username Given"
        $errorCheck = $true
    }
    if($errorCheck){
        return
    }
    ####
    # Check Credentals for validity
    ####
    $securePassword = ConvertTo-SecureString $WPFpasswordBox.Password -AsPlainText -Force
    $fullusername = "PSU\"+$WPFusername.Text
    $Credentals = New-Object System.Management.Automation.PSCredential ($fullusername, $securePassword)
    try{
        Start-Process -FilePath cmd.exe /c -Credential ($Credentals) -PassThru -Wait
    }
    catch{
        $WPFErrorBlock.Text = $WPFErrorBlock.Text+"`nError: Bad UserName or Password"
        $WPFpasswordBox.Password = ""
        return
    }
    ####
    # Check Valididty of computer Name
    ####
    try{
        Get-ADComputer $WPFcomputername.Text -ErrorAction Stop -Credential $Credentals
    }
    catch{
        $WPFErrorBlock.Text = $WPFErrorBlock.Text+"`nError: Computer can not be found in AD"
        return
    }
    ####
    # Change Computers exclusion
    ####
    CleanGroups $Credentals
    $ADComputerName = $WPFcomputername.Text+"$"
    ##Set computer to standard screen lock policy
    if($WPFcomboBox.Text -eq "PSU Standard" -or $WPFcomboBox.Text -eq ""){
        $WPFErrorBlock.Text = $WPFErrorBlock.Text+"`nINFO: PSU standard is set"
    }
    ##Set computer to High Security lock policy
    if($WPFcomboBox.Text -eq "High Security"){
        $WPFErrorBlock.Text = $WPFErrorBlock.Text+"`nINFO: High Security is set"
        ADD-ADGroupMember -Identity CSA_Security_ScreenSaverLock_Reduced_LG -Members $ADComputerName -Credential $Credentals -Confirm:$false
    }
    ##Set computer to Exempt lock policy
    if($WPFcomboBox.Text -eq "Exempt"){
        $WPFErrorBlock.Text = $WPFErrorBlock.Text+"`nINFO: Exempt is set"
        ADD-ADGroupMember -Identity CSA_Security_ScreenSaverLockOverride_LG -Members $ADComputerName -Credential $Credentals -Confirm:$false
    }
    [string] $RTSubject = "Screenlock Policy change on " + $WPFcomputername.Text +" as Authorized by: " + $WPFusername.Text
    [string] $RTMessage = "Screenlock Policy change on " + $WPFcomputername.Text +" as Authorized by: " + $WPFusername.Text + " Reason: " + $WPFchangereason.Text
    RTRequest -reqType "create" -reqSubject $RTSubject -reqMessage $RTMessage -reqStatus "Resolved" -reqQueue "cis-csa" -reqRequestor $WPFusername.Text

    return
    
}

                        
$WPFbutton.Add_Click({
    $form.Close()})
$WPFbutton1.Add_Click({
    ExecuteMain
    return
    })


#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null