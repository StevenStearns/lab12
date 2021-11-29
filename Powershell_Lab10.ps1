# Import active directory module for running AD cmdlets
Import-Module activedirectory
  
#Store the data from ADUsers.csv in the $ADUsers variable
$ADUsers = Import-csv C:\Users\Administrator\Downloads\Powershell_1\bulk_user_import.csv

#Loop through each row containing user details in the CSV file 
foreach ($User in $ADUsers)
{
	#Read user data from each field in each row and assign the data to a variable as below
		
	$Username 	= $User.username
	$Password 	= $User.password
	$Firstname 	= $User.firstname
	$Lastname 	= $User.lastname
	$OU 		= $User.ou #This field refers to the OU the user account is to be created in
    $email      = $User.email
    $streetaddress = $User.streetaddress
    $city       = $User.city
    $postalcode  = $User.postalcode
    $state      = $User.state
    $country    = $User.country
    $telephone  = $User.telephone
    $jobtitle   = $User.jobtitle
    $company    = $User.company
    $department = $User.department
    $Password = $User.Password


	#Check to see if the user already exists in AD
	if (Get-ADUser -F {SamAccountName -eq $Username})
	{
		 #If user does exist, give a warning
		 Write-Warning "A user account with username $Username already exist in Active Directory."
	}
	else
	{
		#User does not exist then proceed to create the new user account
		
        #Account will be created in the OU provided by the $OU variable read from the CSV file
		New-ADUser `
            -SamAccountName $Username `
            -UserPrincipalName "$Username@dc.stearn102.local" `
            -Name "$Firstname $Lastname" `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Enabled $True `
            -DisplayName "$Lastname, $Firstname" `
            -Path $OU `
            -City $city `
            -Company $company `
            -Country $country `
            -PostalCode $postalcode `
            -State $state `
            -StreetAddress $streetaddress `
            -OfficePhone $telephone `
            -EmailAddress $email `
            -Title $jobtitle `
            -Department $department `
            -AccountPassword (convertto-securestring $Password -AsPlainText -Force) -ChangePasswordAtLogon $True
            
	}
   
     $label1 = "$Username"
     $label2 = $Username
     $label1 >> C:\user_onboard_info.txt
     $label2 >> C:\user_onboard_info.txt
      #This adds PSSnapin

    if($null -eq (Get-PSSnapin -Name MailEnable.provision.command -ErrorAction SilentlyContinue))
    {
        add-PSSnapin MailEnable.Provision.command
    }

    New-MailEnableMailbox -mailbox "$Username" -domain "stearn102.local" -password "$Password" -right User

    echo $Username
    #Sends email to the users
    #send-mailmessage -from imapmail@stearns102.local -to $Username@stearns102.local -subject "welcome" -body "welcome to the company" -SmtpServer 
    & 'C:\Program Files (x86)\Mail Enable\bin\MESend.Exe' /F:IMAPmail@stearn102.local /T:$Username@stearn102.local /S:Welcome /B:Welcome to the Company /H:127.0.0.1
    
}

#& 'C:\program Files (x86)\Mail Enable\bin\MESend.exe' /F:IMAPmail@stearns102.local /T:IMAPmail@stearn102.local /S:New Users /B:bulkUsers /A:C:\user_onboard_info.txt /H:192.168.1.2