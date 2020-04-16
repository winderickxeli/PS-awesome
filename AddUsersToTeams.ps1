# AddUsersToTeams
# Eli Winderickx
# Add users to teams via a txt file
# users need to be in the file without @domein
# Requirments : 
# Windows 10 / Server 2019
# Powershell 5 + MicrosoftTeams module
# Install-Module MicrosoftTeams

$domainname=<YOUR DOMAINNAME HERE>

# import Module and connect
Import-Module MicrosoftTeams

Write-Output "Connecting to MS Teams..."
Write-Output "If you don't see a loginprompt, it's probably behind this window"
Connect-MicrosoftTeams

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


# Determing GroupID
do {
    $input = Read-Host "Do you know the owner or the team name (O/T) (Default: O)"
    if (!$input){
        $input="O"
    }

} While(-not($input -eq "T" -or $input -eq "O"))

if ($input -eq "T"){
    do{
        $teamDN = Read-Host "Type in the exact Team name"
    } while(!$teamDN)
    
    $teamID = (Get-Team -DisplayName $teamDN).GroupId
    Write-out "Groups ID is " + $teamID
} else {
    do{
        $teamOw = Read-Host "Type in the username (without domain)"
    } while(!$teamOw)
    $teamOw = $teamOw + "@$domainname"

    $teamGr = get-team -user $teamOw
    
    $x = 0
    Write-Output "To which team would you like to add the users?"
    ForEach ($y in $teamGr.DisplayName) {
        $x = $x + 1
        Write-Output "$x) $y"
    }
    $TeamNu = Read-Host "Choose your team (1-$x)"
    $TeamNu = $TeamNu - 1

    $teamID = (Get-Team -DisplayName $teamGr[$TeamNu].DisplayName).GroupId
    Write-Output "Groups ID is $teamID"
}

# Getting content of txt file
Get-ChildItem $inputfile | ForEach-Object {
    $users = get-content $_ 
}

foreach ($user in $users)
{
    $user = $user + "@$domainname"
    Add-TeamUser -GroupId $teamID -User $user
    write-output "$user toegevoegd"
}

Disconnect-MicrosoftTeams
write-output "Logging off"
