# AddGuestToTeam
# Eli WInderickx
# Add guest users to a Team
# Prereq's:
# - txt file with e-mailadresses
# - Edit <YOUR MS DOMAIN GOES HERE>
# - PS Module MSOnline
# - PS Module MicrsoftTeams

Import-Module MSOnline
Import-Module MicrosoftTeams

## Connect to MS Teams and MSOnline
Write-Output "You're going to have to login twice into each module"
Write-Output "If you can't see the loginprompt, it's probably behind this window"

Connect-MicrosoftTeams
Connect-MsolService

## if Users.txt doesn't exist in current folder, give another file
if(-not(Test-path .\users.txt)){
    Write-Output "users.txt niet gevonden"
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

# GroupID can be looked up by Team or by Owner
do {
    $input = Read-Host "Do you know the owner or Teamname? (O/T) (Default: O)"
    if (!$input){
        $input="O"
    }

} While(-not($input -eq "T" -or $input -eq "O"))

if ($input -eq "T"){
    do{
        $teamDN = Read-Host "Geef de exacte naam van het team op"
    } while(!$teamDN)
    
    $teamID = (Get-Team -DisplayName $teamDN).GroupId
    Write-out "Groups ID is " + $teamID
} else {
    do{
        $teamOw = Read-Host "What is the owners username (without @domainname)"
    } while(!$teamOw)
    $teamOw = $teamOw + "@ap.be"

    $teamGr = get-team -user $teamOw
    
    $x = 0
    Write-Output "To which team would you like to add the users?"
    ForEach ($y in $teamGr.DisplayName) {
        $x = $x + 1
        Write-Output "$x) $y"
    }
    $TeamNu = Read-Host "Choose Team (1-$x)"
    $TeamNu = $TeamNu - 1

    $teamID = (Get-Team -DisplayName $teamGr[$TeamNu].DisplayName).GroupId
    Write-Output "Groups ID is $teamID"
}

# Get content of txt file
Get-ChildItem $inputfile | ForEach-Object {
    $users = get-content $_ 
}

$password= Read-Host "Supply a password for the users"

foreach ($email in $users)
{
    ### Create user
    $pos=$email.Indexof("@")
    $MNN=$email.Replace("@","_")
    $UPN=$MNN+"#EXT#@<YOUR MS DOMAIN GOES HERE>.onmicrosoft.com"
    $name=$email.Remove($pos,$email.Length-$pos)
    $DN="Gast " + $name
    $OM=$email
    New-MsolUser -UserPrincipalName $UPN -DisplayName $DN -FirstName "Gast" -LastName $name -AlternateEmailAddresses $email -Password $password
    Write-Output "Created user $DN..."
    Start-Sleep -Seconds 4
    $userid=(Get-MsolUser -UserPrincipalName $UPN).objectid

    ### add to team
    Add-TeamUser -GroupId $teamID -User $userid
    Write-Output "and it's added to the team"
}

## Disconnect
Disconnect-MicrosoftTeams
Disconnect-AzureAD
