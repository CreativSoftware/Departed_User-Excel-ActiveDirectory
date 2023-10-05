Import-Module ActiveDirectory
Install-Module -Name ImportExcel -Force

#Email Setup
$From = Read-Host -Prompt "Please enter YOUR Email Address"
$EmailTo = "desktoptechs@doi.nyc.gov", "SecurityAlert@doi.nyc.gov"

#Input domain credientials and verifies them.
$authenticate = $true
$attempts = 3
while ($authenticate) {
    $domain_username = Read-Host -Prompt "Enter YOUR ADMIN domain\username"
    $credientials = Get-Credential -UserName $domain_username -Message 'Enter Admin Password'
    try {
        $session = New-PSSession -ComputerName 'doidc02' -Credential $credientials -ErrorAction Stop
        Remove-PSSession $session
        Write-Host "Authentication successful" -ForegroundColor Green
        $authenticate = $false
    } catch {
        $attempts = $attempts - 1
        if ($attempts -eq 0){
            Write-Host "Too many failed attempts. Exiting console." -ForegroundColor Red
            exit
        }
        Write-Host "Failed to authenticate please try again. $attempts attempts remaining." -ForegroundColor Red
    }
}

#Imports users from Excel Document
$externaltempusers = Import-Excel -Path .\ExternalTemps.xlsx | Select-Object Name 

foreach ($user in $externaltempusers){
    
    $names = Get-ADUser -Filter "Name -eq '$($user.Name)'"
    $names = @($names)
    
    foreach ($name in $names) {
        $fullname = $name.Name
        if ($name.distinguishedName -eq "CN=$fullname,OU=ExternalTempUsers,OU=DOI Users,DC=DOI,DC=NYCNET") {
            $username =  $name.SamAccountName
        }
    }

    #Verify the Account Termination
    $account_name = $names.Name
    $username_verify = Read-Host -Prompt "Are you sure you want to Terminate the following user? (Y/N) $account_name"
    if ($username_verify -eq 'Y' -or $username_verify -eq 'y'){
        
    }else{
        exit
    }

    #Assigned memberships
    $assignedgroups = Get-ADPrincipalGroupMembership -Identity $username | Select-Object Name | Out-String

    #Disable user account
    Disable-ADAccount -Identity $username -Credential $credientials

    #clear the Manager and Direct report fields
    Set-ADUser -Identity $username -Clear Manager -Credential $credientials
    $directreports = Get-ADUser -Identity $username -properties DirectReports | select-object -ExpandProperty DirectReports
    foreach($user in $directreports){
        Set-ADUser -Identity $user -Clear Manager -Credential $credientials
    }

    #Remove all memberships from AD account
    $membershipgroups = Get-ADPrincipalGroupMembership -Identity $username

    foreach ($membership in $membershipgroups){
        if ($membership.distinguishedName -eq 'CN=Domain Users,OU=General SG,OU=Security Groups,OU=Groups,DC=DOI,DC=NYCNET')
        {
        continue
        }
        Remove-ADPrincipalGroupMembership -Identity $username -MemberOf $membership.distinguishedName -Credential $credientials -Confirm:$false
    }

    #Move AD account to Departed User's OU
    $username_details = Get-ADUser -Identity $username
    Move-ADObject -Identity $username_details.distinguishedName -TargetPath 'OU=Departed Users,DC=DOI,DC=NYCNET' -Credential $credientials

    #Move the Home and Profile folders to the Archive server. 
    Invoke-Command -ComputerName "doidc02" -Credential $credientials -ScriptBlock {
        $Folder_Name = $using:username
        $Path1 = "\\doiarchive01\home_archive\$Folder_Name"
        New-Item -Path $Path1 -ItemType Directory 
        $Path2 = "\\doiarchive01\profile_archive\$Folder_Name"
        New-Item -Path $Path2 -ItemType Directory 

        $Source_Home_Folder = "\\doi.nycnet\doi_share\home_folder\$Folder_Name"
        $Destination_Home_Folder = "\\DOIARCHIVE01\HOME_ARCHIVE\$Folder_name"

        $Source_Profile_folder = "\\DOIPROFILE01\USER_FOLDER_REDIRECTION\$Folder_name"
        $Destination_Profile_folder = "\\DOIARCHIVE01\PROFILE_ARCHIVE\$Folder_name"

        #Robocopy Execute
        robocopy $Source_Home_Folder $Destination_Home_Folder /COPYALL /Z /E /W:1 /R:2 /tee /Move 
        robocopy $Source_Profile_folder $Destination_Profile_folder /COPYALL /Z /E /W:1 /R:2 /tee /Move 
    }

    #Sends Email with user's memberships
    $fullname = $username_details.Name
    Send-MailMessage -From $From -To $EmailTo -Subject "Departed User $fullname" -body "The Departed account $fullname is now completed. Their home and profile folders have been moved to the Archived Server. Here is a list of Group Memberships he/she was assigned to: `n$assignedgroups" -SmtpServer 'smtp.doi.nycnet' -Port '25'
}
