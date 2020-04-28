# AddGuestToTeam
# Eli WInderickx
# Invite Guest user to a Team
# Required:
# - List of adresses in a TXT file
# - Module AzureAD (Install-Module AzureAD)
# - Module MicrsoftTeams (Install-Module MicrosoftTeams)

Import-Module AzureAD
Import-Module MicrosoftTeams

## Connecting to AzureAD and MS Teams
Write-Output "You're going to have to login twice"
Write-Output "If you don't see a prompt, it's probably behind this window"

Connect-MicrosoftTeams
$tenantid= (Connect-AzureAD).TenantId

## Supply another file if users.txt is not found
if(-not(Test-path .\users.txt)){
    Write-Output "users.txt not found"
        Function Get-FileName($initialDirectory)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    }
    Function Get-FileName($initialDirectory)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "TXT (*.txt)| *.txt"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
    }
    $inputfile = Get-FileName Get-Location 

    if(!$inputfile){
        exit
    }
} else{
    $inputfile=".\users.txt"
}

Write-Output "The inputfile is:"
Write-Output $inputfile

# Determine GroupID with a user or with Team name. 
do {
    $input = Read-Host "Do you know the owner or the Team name? (O/T) (Default: O)"
    if (!$input){
        $input="O"
    }

} While(-not($input -eq "T" -or $input -eq "O"))

if ($input -eq "T"){
    do{
        $teamDN = Read-Host "Enter the exact Teamname"
    } while(!$teamDN)
    
    $teamID = (Get-Team -DisplayName $teamDN).GroupId
    Write-out "GroupID is " + $teamID
} else {
    do{
        $teamOw = Read-Host "What is the username of the owner?"
    } while(!$teamOw)
    
    $teamGr = get-team -user $teamOw
    
    $x = 0
    Write-Output "Add users to what team?"
    ForEach ($y in $teamGr.DisplayName) {
        $x = $x + 1
        Write-Output "$x) $y"
    }
    $TeamNu = Read-Host "Select your team (1-$x)"
    $TeamNu = $TeamNu - 1

    $teamID = (Get-Team -DisplayName $teamGr[$TeamNu].DisplayName).GroupId
    Write-Output "Group ID is $teamID"
}

# Get content of inputfile
Get-ChildItem $inputfile | ForEach-Object {
    $users = get-content $_ 
}

## Set up mail 
$messageInfo = New-Object Microsoft.Open.MSGraph.Model.InvitedUserMessageInfo
$teamsurl = "https://teams.microsoft.com/l/team/19%3a2e731f4f66d762c297c31b00053f6973%40thread.tacv2/conversations?groupId=$teamID&tenantId=$tenantid"
$teamName = $teamGr[$TeamNu].DisplayName

## Body of e-mail. Edit between quotes. You can use $teamName to add the teamname to your message
$messageInfo.customizedMessageBody = "Hi, you're being invited to our team, $teamName"

foreach ($email in $users)
{
    ### Invite every user
    $pos=$email.Indexof("@")
    $name=$email.Remove($pos,$email.Length-$pos)
    $DN="Guest " + $name
    New-AzureADMSInvitation -InvitedUserEmailAddress $email -InvitedUserDisplayName $DN -InviteRedirectUrl $teamsurl -InvitedUserMessageInfo $messageInfo -SendInvitationMessage $true
    Write-Output "$DN got invited..." 
    Start-Sleep -Seconds 4

    ### Voeg nieuwe gebruiker toe aan team
    Add-TeamUser -GroupId $teamID -User $email
    Write-Output "... and got added to the team."
}

## Clean up 
Disconnect-MicrosoftTeams
Disconnect-AzureAD
