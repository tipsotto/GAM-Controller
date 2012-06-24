# GAM-Controller.ps1
# Created by: 	TIPS
# Version: 		1.0
clear

# VARIBALE TO SET
$execPath = "C:\gam\gam.exe"							# SET THIS TO: "\path\to\gam.exe"
$defaultPassword = ""									# DEFAULT PASSWORD TO USE WHEN CREATING USERS (THIS WILL BE USED WHEN YOU TYPE "\d" FOR PASSWORD IN USER CREATION)

#***********************************************************************************************************************************************************************************************************************
# TEST EXECUTABLE PATH
$execExists = Test-Path $execPath
if ($execExists -eq $False) {Write-Host "Error: Check line 7`n>> gam.exe does not exist at path specified by execPath variable.`n>> Currently execPath = $execPath" -ForegroundColor Red; exit 1}
Set-Alias gam $execPath
# TEST DEFAULT PASSWORD SET?
if ($defaultPassword -eq "") {Write-Host -NoNewLine "Warning: " -ForeGroundColor Yellow; Write-Host "Check line 8. `"`$defaultPassword`" variable not set." -ForeGroundColor Gray}

$start = 0

#############
##FUNCTIONS##		COLLAPSE FUNCTIONS FOR EASIER READING
#############
### MAIN MENU
Function Begin
{
	if ($start -gt 0) {echo "`n`n"}; $start++
	$script:choice = select-item -Caption "*** Google Apps Manager ***" -Message "What do you want to do: " -choice "&User Management", "&Email Settings", "&Domain Settings", "&Reports", "&Quit"  -default 4
	
	if ($choice -eq 4) {echo "`n`nGood Bye`n";exit 0} 	# QUIT
	elseif ($choice -eq 0) {UserManagement}				# USER MANAGEMENT
	elseif ($choice -eq 1) {EmailSettings}				# EMAIL SETTINGS
	elseif ($choice -eq 2) {DomainSettings}				# DOMAIN SETTINGS
	elseif ($choice -eq 3) {Reports}					# REPORTS
}

### USER MANAGEMENT
### SEE WIKI HERE FOR MORE INFO: ( http://code.google.com/p/google-apps-manager/wiki/ExamplesProvisioning )
Function UserManagement
{
	$script:umchoice = select-item -Caption "*** User Management ***" -Message "What do you want to do: " -choice "C&reate User","&Bulk Create Users", "&Delete User(s)", "Get User &Info", "&Cancel"  -default 4; echo ""
	
	Function CreateUser				# CREATE A USER
	{
		Write-Host "Create a New User`n(Leave `'Password`' field blank to cancel)" -ForegroundColor Cyan
		Write-Host -NoNewLine "First Name: " -ForegroundColor Magenta; $firstName = Read-Host
		Write-Host -NoNewLine "Last Name: " -ForegroundColor Magenta; $lastName = Read-Host
		Write-Host -NoNewLine "Username: " -ForegroundColor Magenta; $userName = Read-Host
		Function AskPassword		# ASK FOR PASSWORD SECURELY
		{
			Write-Host -NoNewLine "(Leave blank to cancel, type `"\d`" for default password) Password: " -ForegroundColor Magenta; $script:password = Read-Host -AsSecureString
			$basicString = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
			$script:password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($basicString)
			if ($password -eq "\d") {$script:password = $defaultPassword;$script:confpassword = $defaultPassword; echo "Using default password..."}
			else {
				Write-Host -NoNewLine "Confirm Password: " -ForegroundColor Magenta; $script:confpassword = Read-Host -AsSecureString
				$basicString = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($confpassword)
				$script:confpassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($basicString)
			}
		};AskPassword
		while ($password -ne "")	# IF PASSWORD NOT EMPTY
		{
			while ($password -ne $confpassword) # MAKE SURE PASSWORD AND CONFIRMATION MATCH, IF NOT LOOP AskPassword
				{
					Write-Host "Password does not match! Repeat" -ForegroundColor Red
					AskPassword
				}
			if ($password -eq $confpassword) {gam create user $username firstname $firstName lastname $lastName password $password nohash agreedtoterms on; break} # IF PASS AND CONF MATCH
			else {echo "`n---Password does not match. User was not created, nothing was changed.`n"}
		}
		if ($password -eq "") {echo "`n---Password empty. User was not created, nothing was changed. Cancelled.`n"}
	}
	
	Function BulkCreateUsers		# CREATE USERS IN BULK (MULTIPLE USER CREATION)
	{
		Write-Host "Create New Users in Bulk`n(Follow this syntax: firstname1,lastname1,username1,password1;firstname2,lastname2,username2,password2)" -ForegroundColor Cyan
		Write-Host "(type `"\cancel`" to cancel, type `"\d`" in password field for default password)" -ForegroundColor Cyan
		Write-Host -NoNewLine "`nUsers to Create: " -ForegroundColor Cyan; $script:usersInfo = Read-Host
		$script:usersInfo = $usersInfo.Replace(" ", "")
		if ($usersInfo -ne "\cancel")
		{
			Write-Host "Processing ..."
			if ($usersInfo.substring($usersInfo.length-1) -eq ";") {$script:usersInfo = $usersInfo -replace ".$"}
			$script:usersInfo = $usersInfo.split(";")
			foreach ($user in $usersInfo)
			{
				$user = $user.split(",")
				$fname = $user[0]; $lname = $user[1]; $uname = $user[2]; $script:passd = $user[3]
				if ($script:passd -eq "\d") {$script:passd = "$defaultPassword"}
				gam create user $uname firstname $fname lastname $lname password $passd nohash agreedtoterms on
			}
		}
		else {echo "`n---User creation was Cancelled. No users were created, nothing was changed.`n"}
	}
	
	Function DeleteUser				# DELETE USER(s)
	{
		Write-Host "Delete Existing User(s)" -ForegroundColor Cyan
		Write-Host "(for more than one user, use this syntax: -bulk:username1;username2;username3)" -ForegroundColor Cyan
		Write-Host -NoNewLine "(Leave blank to cancel) Username(s): " -ForegroundColor Magenta; $script:userNames = Read-Host
		if (($userNames -ne $Null) -and ($userNames -ne ""))
		{
			$script:userNames = $script:userNames.Replace(" ", "")
			if ($userNames.substring(0,1) -eq "-")
			{
				if ($userNames.substring(0,6) -eq "-bulk:")
				{
					$script:userNames = $userNames.substring(6,$userNames.length-6)
					$script:userNames = $script:userNames.split(";")
					foreach ($user in $userNames)
					{
						gam delete user $user
					}
				}
			}
			else {gam delete user $userNames}
		}
		else {echo "`n---Input blank. No user was deleted, nothing was changed.`n"}
	}
	
	Function GetUserInfo			# GET A USER'S INFO
	{
		Write-Host "Get User Info" -ForegroundColor Cyan
		Write-Host -NoNewLine "(Leave blank to cancel) Username: " -ForegroundColor Magenta; $userName = Read-Host
		if (($userName -ne $Null) -and ($userName -ne "")) {gam info user $userName}
		else {echo "`n---Request Cancelled.`n"}
	}
	# USER MANAGEMENT OPTION REDIRECTOR
	if ($umchoice -eq 4) {$script:umchoice = -1} 				# GO BACK TO MAIN MENU
	elseif ($umchoice -eq 0) {CreateUser; UserManagement}		# CREATE USER
	elseif ($umchoice -eq 1) {BulkCreateUsers; UserManagement}	# BULK CREATE USERS
	elseif ($umchoice -eq 2) {DeleteUser; UserManagement}		# DELETE USER(s)
	elseif ($umchoice -eq 3) {GetUserInfo; UserManagement}		# GET USER INFO
}

### EMAIL SETINGS
### SEE WIKI HERE FOR MORE INFO: ( http://code.google.com/p/google-apps-manager/wiki/ExamplesEmailSettings )
Function EmailSettings
{
	$script:eschoice = select-item -Caption "*** Email Settings ***" -Message "What do you want to do: " -choice "&Delegation", "&Enable/Disable IMAP", "Disable `"&Web Clips`"", "&Cancel"  -default 3; echo ""
	
	Function Delegation
	{
		$dchoice = select-item -Caption "Email Delegation" -Message "What do you want to do:" -choice "&Create a Delegation", "&Delete a Delegation", "&Cancel"  -default 2
		if ($dchoice -eq 2) {$dchoice = -1}
		elseif ($dchoice -eq 0)
		{
			Write-Host "Note: Users affected MUST be active, and NOT have to change their passwords on next login." -ForeGroundColor Gray
			echo "type \cancel to cancel."
			Write-Host -NoNewLine "(type in username) Give this user access: " -ForegroundColor Cyan; $duser = Read-Host
			if ($duser -ne "\cancel")
			{
				Write-Host -NoNewLine "(type in username) Give $user access to this user's mail: " -ForegroundColor Cyan; $muser = Read-Host
				gam user $muser delegate to $duser
			}
			else {EmailSettings}
		}
		elseif ($dchoice -eq 1)
		{
			Write-Host "Note: Users affected MUST be active, and NOT have to change their passwords on next login." -ForeGroundColor Gray
			echo "type \cancel to cancel."
			Write-Host -NoNewLine "(type in username) Delete this user's access: " -ForegroundColor Cyan; $duser = Read-Host
			if ($duser -ne "\cancel")
			{
				Write-Host -NoNewLine "(type in username) Delete $user access to this user's mail: " -ForegroundColor Cyan; $muser = Read-Host
				gam user $muser delete delegate $duser
			}
			else {EmailSettings}
		}
	}
	
	Function EDImap												# ENABLE/DISABLE IMAP FOR USER(S)
	{
		$script:edichoice = select-item -Caption "Enable/Disable IMAP for user(s)" -Message "Turn IMAP on or off?" -choice "&On", "O&ff", "&Cancel"  -default 2
		if ($edichoice -eq 2) {$edichoice = -1}					# Cancel ENABLE/DISABLE IMAP
		elseif ($edichoice -eq 0)								# ENABLE IMAP
		{
			Write-Host "`nENABLE IMAP. type `"\cancel`" to go to Main Menu, type `"\all`" to apply to all users:" -ForegroundColor Cyan
			Write-Host -NoNewLine "Type in username(s) of User(s) seperated by commas (no spaces): " -ForegroundColor Cyan; $users = Read-Host
			switch ($users)
			{
				"\cancel" {echo "`n---Request Cancelled. Nothing was changed."; EmailSettings}
				"\all" {echo "Enabling IMAP for all users..."; gam all users imap on}
				default {
							$users = $users.split(",")
							foreach ($user in $users) {gam user $users imap on}
						}
			}
		}
		elseif ($edichoice -eq 1)								# DISABLE IMAP
		{
			echo "DISABLE IMAP. Type in username(s) of user(s) seperated by commas (no spaces)"
			Write-Host -NoNewLine "type `"\cancel`" to go to Main Menu, type `"\all`" to apply to all users: "; $users = Read-Host
			switch ($users)
			{
				"\cancel" {echo "`n---Request Cancelled. Nothing was changed."; EmailSettings}
				"\all" {echo "Disabling IMAP for all users..."; gam all users imap off}
				default {
							$users = $users.split(",")
							foreach ($user in $users) {gam user $users imap off}
						}
			}
		}
	}
	
	Function DisableWebClips									# DISABLE WEBCLIPS FOR USER(S)
	{
		echo "DISABLE `"WEB CLIPS`". Type in username(s) of user(s) seperated by commas (no spaces)"
		Write-Host -NoNewLine "type `"\cancel`" to go to Main Menu, type `"\all`" to apply to all users: "; $users = Read-Host
		switch ($users)
		{
			"\cancel" {echo "`n---Request Cancelled. Nothing was changed."; EmailSettings}
			"\all" {echo "Disabling Web ClipsS for all users..."; gam update all users webclips off}
			default {
						$users = $users.split(",")
						foreach ($user in $users) {gam user $users webclips off}
					}
		}
	}
	# EMAIL SETTINGS OPTION REDIRECTOR
	if ($eschoice -eq 3) {$script:eschoice = -1}				# GO BACK TO MAIN MENU
	elseif ($eschoice -eq 0) {Delegation}						# DELEGATE EMAIL ACCOUNTS
	elseif ($eschoice -eq 1) {EDImap}							# ENABLE/DISBALE IMAP
	elseif ($eschoice -eq 2) {DisableWebClips}					# DISABLE "WEB CLIPS"
}

### DOMAIN SETTINGS
### SEE WIKI HERE FOR MORE INFO: ( http://code.google.com/p/google-apps-manager/wiki/DomainSettingsExamples )
Function DomainSettings
{
	$script:dschoice = select-item -Caption "*** Domain Settings ***" -Message "What do you want to do: " -choice "Get Domain &Info", "Enable/Disable User Mail &Migrations", "&Cancel"  -default 2
	echo ""
	Function MailMigration				# ENABLE/DISABLE THE MAIL MIGRATION API
	{
		$script:mmchoice = select-item -Caption "*** Enable/Disable User Mail Migrations ***" -Message "Do you want to Enable or Disable User Mail Migration API: " -choice "&Enable API", "&Disable API", "&Cancel"  -default 2
		if ($mmchoice -eq 2) {DomainSettings}		#GO BACK TO DOMAIN SETTINGS
		elseif ($mmchoice -eq 0) {gam update domain user_migrations true}
		elseif ($mmchoice -eq 1) {gam update domain user_migrations false}
	}
	
	if ($dschoice -eq 2) {$script:dschoice = -1}	# GO BACK TO MAIN MENU
	elseif ($dschoice -eq 0) {gam info domain} 		# GET DOMAIN INFO
	elseif ($dschoice -eq 1) {MailMigration}		# ENABLE/DISABLE USER MAIL MIGRATION
}

### REPORTS
### SEE WIKI HERE FOR MORE INFO: ( http://code.google.com/p/google-apps-manager/wiki/ExamplesCSV )
Function Reports
{
	$script:rchoice = select-item -Caption "*** Reports ***" -Message "What do you want to do: " -choice "Print All &Users", "Print All &Groups", 
	"Print All &Nicknames", "Print All Organizational &Units", "Print All &Resource Calendars", "&Other Reports", "&Cancel"  -default 6; echo ""
	
	Function PrintAllUsers							# REPORT ALL USERS TO CSV
	{
		Write-Host "Printing All Users...`nReport will be saved in a .csv file" -ForegroundColor Cyan
		Write-Host "Where do you want to save report? [ex. C:\Users\user\Desktop\users.csv]"
		Write-Host -NoNewLine "(Leave blank to cancel): "; $csvpath = Read-Host
		if (($csvpath -ne $Null) -and ($csvpath -ne "")) {gam print users firstname lastname username ou suspended changepassword agreed2terms admin aliases groups > $csvpath}
		else {echo "`n---Request Cancelled.`n"}
	}
	
	Function PrintAllGroups							# REPORT ALL GROUPS TO CSV
	{
		Write-Host "Printing All Groups...`nReport will be saved in a .csv file" -ForegroundColor Cyan
		Write-Host "Where do you want to save report? [ex. C:\Users\user\Desktop\groups.csv]"
		Write-Host -NoNewLine "(Leave blank to cancel): "; $csvpath = Read-Host
		if (($csvpath -ne $Null) -and ($csvpath -ne "")) {gam print groups name description permission members owners settings > $csvpath}
		else {echo "`n---Request Cancelled.`n"}
	}
	
	Function PrintAllNicknames						# REPORT ALL NICKNAMES (ALIASES) TO CSV
	{
		Write-Host "Printing All Nicknames...`nReport will be saved in a .csv file" -ForegroundColor Cyan
		Write-Host "Where do you want to save report? [ex. C:\Users\user\Desktop\nicknames.csv]"
		Write-Host -NoNewLine "(Leave blank to cancel): "; $csvpath = Read-Host
		if (($csvpath -ne $Null) -and ($csvpath -ne "")) {gam print nicknames > $csvpath}
		else {echo "`n---Request Cancelled.`n"}
	}
	
	Function PrintAllOrgUnits							# REPORT ALL ORG UNITS TO CSV
	{
		Write-Host "Printing All Organizational Units...`nReport will be saved in a .csv file" -ForegroundColor Cyan
		Write-Host "Where do you want to save report? [ex. C:\Users\user\Desktop\orgunits.csv]"
		Write-Host -NoNewLine "(Leave blank to cancel): "; $csvpath = Read-Host
		if (($csvpath -ne $Null) -and ($csvpath -ne "")) {gam print orgs name description parent inherit > $csvpath}
		else {echo "`n---Request Cancelled.`n"}
	}
	
	Function PrintAllResCals							# REPORT ALL RESOURCE CALENDARS TO CSV
	{
		Write-Host "Printing All Resource Calendars...`nReport will be saved in a .csv file" -ForegroundColor Cyan
		Write-Host "Where do you want to save report? [ex. C:\Users\user\Desktop\resourcecals.csv]"
		Write-Host -NoNewLine "(Leave blank to cancel): "; $csvpath = Read-Host
		if (($csvpath -ne $Null) -and ($csvpath -ne "")) {gam print resources id description email > $csvpath}
		else {echo "`n---Request Cancelled.`n"}
	}
	
	Function OtherReports							# OTHER REPORTS
	{
		$script:orchoice = select-item -Caption "*** Other Reports ***`n`nNote: Allow 24-48 hours for changes to show on these reports." -Message "Which report do you want to run: " -choice "&Accounts", "Acti&vity", "&Disk Space", "&Email Clients", "&Summary", "&Cancel"  -default 5
		
		Function AccountsReport
		{
			Write-Host "Printing Accounts Report..`nReport will be saved in a .csv file" -ForegroundColor Cyan
			Write-Host -NoNewLine "(Leave blank to get most recent) Type in date to get report for that day.[Format: YYYY-MM-DD]: "; $date = Read-Host
			Write-Host "Where do you want to save report? [ex. C:\Users\user\Desktop\accountsreport.csv]"
			Write-Host -NoNewLine "(Leave blank to cancel): "; $csvpath = Read-Host
			if (($csvpath -ne $Null) -and ($csvpath -ne "")) {gam report accounts $date > $csvpath}
			else {echo "`n---Request Cancelled.`n"}
		}
		
		Function ActivityReports
		{
			Write-Host "Printing Activity Report..`nReport will be saved in a .csv file" -ForegroundColor Cyan
			Write-Host -NoNewLine "(Leave blank to get most recent) Type in date to get report for that day.[Format: YYYY-MM-DD]: "; $date = Read-Host
			Write-Host "Where do you want to save report? [ex. C:\Users\user\Desktop\activityreport.csv]"
			Write-Host -NoNewLine "(Leave blank to cancel): "; $csvpath = Read-Host
			if (($csvpath -ne $Null) -and ($csvpath -ne "")) {gam report activity $date > $csvpath}
			else {echo "`n---Request Cancelled.`n"}
		}
		
		Function DiskSpaceReport
		{
			Write-Host "Printing Disk Space Report..`nReport will be saved in a .csv file" -ForegroundColor Cyan
			Write-Host -NoNewLine "(Leave blank to get most recent) Type in date to get report for that day.[Format: YYYY-MM-DD]: "; $date = Read-Host
			Write-Host "Where do you want to save report? [ex. C:\Users\user\Desktop\diskspacereport.csv]"
			Write-Host -NoNewLine "(Leave blank to cancel): "; $csvpath = Read-Host
			if (($csvpath -ne $Null) -and ($csvpath -ne "")) {gam report disk_space $date > $csvpath}
			else {echo "`n---Request Cancelled.`n"}
		}
		
		Function EmailClientsReport
		{
			Write-Host "Printing Email Clients Report..`nReport will be saved in a .csv file" -ForegroundColor Cyan
			Write-Host -NoNewLine "(Leave blank to get most recent) Type in date to get report for that day.[Format: YYYY-MM-DD]: "; $date = Read-Host
			Write-Host "Where do you want to save report? [ex. C:\Users\user\Desktop\emailclientsreport.csv]"
			Write-Host -NoNewLine "(Leave blank to cancel): "; $csvpath = Read-Host
			if (($csvpath -ne $Null) -and ($csvpath -ne "")) {gam report email_clients $date > $csvpath}
			else {echo "`n---Request Cancelled.`n"}
		}
		
		Function SummaryReport
		{
			Write-Host "Printing Summary Report..`nReport will be saved in a .csv file" -ForegroundColor Cyan
			Write-Host -NoNewLine "(Leave blank to get most recent) Type in date to get report for that day.[Format: YYYY-MM-DD]: "; $date = Read-Host
			Write-Host "Where do you want to save report? [ex. C:\Users\user\Desktop\summaryreport.csv]"
			Write-Host -NoNewLine "(Leave blank to cancel): "; $csvpath = Read-Host
			if (($csvpath -ne $Null) -and ($csvpath -ne "")) {gam report summary $date > $csvpath}
			else {echo "`n---Request Cancelled.`n"}
		}
		
		if ($orchoice -eq 5) {$script:orchoice = -1}	# GO BACK TO MAIN MENY
		elseif ($orchoice -eq 0) {AccountsReport}		# PRINT ACCOUNTS REPORT TO .csv FILE
		elseif ($orchoice -eq 1) {ActivityReports}		# PRINT ACTIVITY REPORT TO .csv FILE
		elseif ($orchoice -eq 2) {DiskSpaceReport}		# PRINT DISK SPACE REPORT TO .csv FILE
		elseif ($orchoice -eq 3) {EmailClientsReport}	# PRINT EMAIL CLIENTS REPORT TO .csv FILE
		elseif ($orchoice -eq 4) {SummaryReport}		# PRINT SUMMARY REPORT TO .csv FILE
	}
	
	if ($rchoice -eq 6) {$script:rchoice = -1}			# GO BACK TO MAIN MENU
	elseif ($rchoice -eq 0) {PrintAllUsers}				# PRINT ALL USERS TO .csv FILE
	elseif ($rchoice -eq 1) {PrintAllGroups}			# PRINT ALL GROUPS TO .csv FILE
	elseif ($rchoice -eq 2) {PrintAllNicknames}			# PRINT ALL NICKNAMES TO .csv FILE
	elseif ($rchoice -eq 3) {PrintAllOrgUnits}			# PRINT ALL ORGANIZATIONAL UNITS TO .csv FILE
	elseif ($rchoice -eq 4) {PrintAllResCals}			# PRINT ALL RESOURCE CALENDARS TO .csv FILE
	elseif ($rchoice -eq 5) {OtherReports}				# OTHER REPORTS
}

### SELECT ITEM FUNCTION USED FOR MENUS ###
Function Select-Item 
{
	Param
	(   
		[String[]]$choiceList, 
		[String]$Caption="Please make a selection", 
		[String]$Message="Choices are presented below", 
		[int]$default=0
	) 
	$choicedesc = New-Object System.Collections.ObjectModel.Collection[System.Management.Automation.Host.ChoiceDescription] 
	$choiceList | foreach  { $choicedesc.Add((New-Object "System.Management.Automation.Host.ChoiceDescription" -ArgumentList $_))} 
	$Host.ui.PromptForChoice($caption, $message, $choicedesc, $default) 
} 

while (1) {Begin}									# LOOP SCRIPT TO MAIN MENU