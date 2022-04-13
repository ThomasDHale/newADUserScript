#_  _              _   _               ___         _      _   
#| \| |_____ __ __ | | | |___ ___ _ _  / __| __ _ _(_)_ __| |_ 
#| .` / -_) V  V / | |_| (_-</ -_) '_| \__ \/ _| '_| | '_ \  _|
#|_|\_\___|\_/\_/   \___//__/\___|_|   |___/\__|_| |_| .__/\__|
#                                                    |_|       

# Include Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Specify some variables that might come in handy later
$todayDate = (Get-Date -Format yyyy-MM-dd)

##############################
# Create a regular input box #
##############################

function Get-Input {

    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$TitleBox = "Default Text",
        [string]$Message = "Default Message",
        [string]$DefaultValue = "Default Output"
        [string]$TextBox = ""
    )

    ###
    # Create a box to ask for input
    ###

    # Create a form of 300px and 200px in the middle of the screen
    $form = New-Object System.Windows.Forms.Form
    # The title box value is a parameter of the Get-Input cmdlet
    $form.Text = $TitleBox
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = 'CenterScreen'

    # Draw an OK button in the form
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75,120)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    # Draw a Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(150,120)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    # Draw a label in the form
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    # The message value is a parameter of the Get-Input cmdlet
    $label.Text = $Message
    $form.Controls.Add($label)

    # Draw a text box
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,40)
    $textBox.Size = New-Object System.Drawing.Size(260,20)
    $textBox.Text = $TextBox
    $form.Controls.Add($textBox)

    # Form is on top, otherwise it might end up behind the Powershell window or something
    $form.Topmost = $true

    $form.Add_Shown({$textBox.Select()})
    $result = $form.ShowDialog()

    # Decide what to do with the input of the text box
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        # Set the input as the user's First Name
        $x = $textBox.Text
        $x
    }
    else {
        # If no input was received then set it to a specified default value
        $x = $DefaultValue
        $x
    }
}

############################
# Create a multi-input box #
############################

function Get-MultiInput {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$TitleBox = "Default Text",
        [string]$Message = "Default Message",
        [string]$DefaultValue = "Default Output",
        [array]$MultiValues = "Default"
    )
    
    # Create a form of 300px and 200px in the middle of the screen
    $form = New-Object System.Windows.Forms.Form
    # The title box value is a parameter of the Get-Input cmdlet
    $form.Text = $TitleBox
    $form.Size = New-Object System.Drawing.Size(300,225)
    $form.StartPosition = 'CenterScreen'

    # Draw an OK button in the form
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75,150)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(150,150)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    # Draw a label in the form
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(580,20)
    # The message value is a parameter of the Get-Input cmdlet
    $label.Text = $Message
    $form.Controls.Add($label)

    # Create a list box which will contain a number of items, specified by the $MultiValues parameter
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10,40)
    $listBox.Size = New-Object System.Drawing.Size(260,20)
    $listBox.Height = 100

    # For each $MultiValue that you feed it in the Get-Input cmdlet, it will add them one at a time
    foreach ($multiValue in $MultiValues){
        [void] $listBox.Items.Add($multiValue)
    }

    $form.Controls.Add($listBox)

    # Form is on top, otherwise it might end up behind the Powershell window or something
    $form.Topmost = $true

    $result = $form.ShowDialog()

    # Decide what to do with the input of the text box
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        # Whichever item the user selects becomes the domain name, this will be used in the UserPrincipalName and EmailAddress fields
        $newUserDomainName = $listBox.SelectedItem
        $newUserDomainName
    }
    else {
        # If no input was received then set to a specified default value
        $newUserDomainName = $DefaultValue
        $newUserDomainName
    }
}

############################
# Create a random password #
############################

###
# Use the Dinopass API to fetch a regular password
###

# Download a random weak password from Dinopass using the API: https://www.dinopass.com/api
$newUserPassword = (Invoke-WebRequest -Uri "https://www.dinopass.com/password").Content

# Fiddle with that password so that it has an uppercase character at the start
# You could just use the strong password API instead, but I find those passwords a bit obtuse to be honest
# Passphrases are the strongest passwords because they are loooooooooooooooong https://untroubled.org/pwgen/ppgen.cgi
$newUserPasswordTitleCase = $newUserPassword.Substring(0,1).ToUpper() + $newUserPassword.Substring(1)

# Convert the plaintext string into a 'secure' string
$newUserPasswordString = ConvertTo-SecureString -String $newUserPasswordTitleCase -AsPlainText -Force

############################
# Collect user information #
############################

###
# First Name
###

# Get input. The TitleBox parameter is the title of the window. The Message is the label above the input box. The DefaultValue is what value is set if the input box is empty
$newUserFirstName = Get-Input -TitleBox "First Name" -Message "Please enter the first name of the user:" -DefaultValue "John"
$newUserFirstName

# Make sure that the first letter of the name is capitalised
$newUserFirstNameTitleCase = $newUserFirstName.Substring(0,1).ToUpper() + $newUserFirstName.Substring(1)
$newUserFirstNameTitleCase

# Also create a separate variable to store the lowercase name, this will be used in the UserPrincipalName and EmailAddress attributes
$newUserFirstNameLowerCase = $newUserFirstName.ToLower()

###
# Last Name
###

$newUserLastName = Get-Input -TitleBox "Last Name" -Message "Please enter the last name of the user:" -DefaultValue "Smith"
$newUserLastName

# Make sure that the last name is capitalised
$newUserLastNameTitleCase = $newUserLastName.Substring(0,1).ToUpper() + $newUserLastName.Substring(1)
$newUserLastNameTitleCase

# Also store the lowercase version of the surname as a separate value to be used in the UserPrincipalName and EmailAddress fields
$newUserLastNameLowerCase = $newUserLastName.ToLower()

###
# Domain Name
###

# Get Multi-Input
$newUserDomainName = Get-MultiInput -TitleBox "Select a domain" -Message "Please choose a domain:" -DefaultValue "bob.com" -MultiValues "bob.com","bob.co.uk","fred.co.uk","ted.ie","ben.co.uk"
$newUserDomainName

###
# Display Name
###

# Combine the title case first name and last name to create the display name variable
$newUserDisplayName = "$newUserFirstNameTitleCase $newUserLastNameTitleCase"
$newUserDisplayName

###
# SAM Account Name
###

# Combine the lower case firstname and lastname to create the AccountName
$newUserAccountName = "$newuserFirstNameLowerCase.$newUserLastNameLowerCase"
$newUserAccountName

###
# User Principal Name / Email Address
###

# Combine the lower case firstname, lastname and domainname to create the UPN/mail value. These should be the same for consistency.
$newUserEmailAddress = "$newuserFirstNameLowerCase.$newUserLastNameLowerCase@$newuserDomainName"
$newUserEmailAddress

###
# Job Title
###

# Get input
$newUserJobTitle = Get-Input -TitleBox "Job Title" -Message "Please enter the job title of the user:" -DefaultValue "Assistant Refuse Analyst"
$newUserJobTitle

###
# Manager
###

# Get input
$newUserManager = Get-Input -TitleBox "Manager" -Message "Please enter the manager's username (firstname.lastname):" -DefaultValue "bob.robert"
$newUserManager

###
# Department
###

# Get input
$newUserDepartment = Get-Input -TitleBox "Department" -Message "Please enter the department of the user:" -DefaultValue "Janitorial Services"
$newUserDepartment

###
# Office Location
###

# Get input
$newUserOfficeLocation = Get-MultiInput -TitleBox "Select an Office Location" -Message "Please choose an Office location:" -DefaultValue "Remote" -MultiValues "Remote","Breford","Sorton","Laneham"
$newUserOfficeLocation

# Set the container/OU path, address, postcode, country and additional groups for the user based on their OfficeLocation
if ($newUserOfficeLocation -eq 'Remote')
{
    # The OU of the user
    $newUserPath = "ou=users,ou=remote,ou=sites,ou=uk,dc=bob,dc=com"
    # The address, postcode and country
    $newUserStreetAddress = "Remote,`r`nRemote"
    $newUserCity = 'Remote'
    $newUserState = 'Remote'
    $newUserPostcode = 'Remote'
    $newUserCountry = 'GB'
    # Any additional groups that they might need
    $additionalGroups = "SecUK_IT_RemoteDesktop","DistUK_RemoteWorkers"
    # The company to which they belong
    $newUserCompany = "Bob Robert Industries Co."
}
if ($newUserOfficeLocation -eq 'Breford')
{
    $newUserPath = "ou=users,ou=breford,ou=sites,ou=uk,dc=bob,dc=com"
    $newUserStreetAddress = "10 St. Robert Street"
    $newUserCity = 'Breford'
    $newUserState = 'Brefordshire'
    $newUserPostcode = 'BR1 1AX'
    $newUserCountry = 'GB'
    $additionalGroups = "SecUK_IT_RemoteDesktop","DistUK_BrefordOffice"
    $newUserCompany = "Bob Robert Industries Co."
}
if ($newUserOfficeLocation -eq 'Sorton')
{
    $newUserPath = "ou=users,ou=sorton,ou=sites,ou=uk,dc=bob,dc=com"
    $newUserStreetAddress = "Floor 4,`r`nJoe Bloggs Building"
    $newUserCity = 'Sorton'
    $newUserState = 'Wraghamshire'
    $newUserPostcode = 'SR3 7AB'
    $newUserCountry = 'GB'
    $additionalGroups = "SecUK_IT_RemoteDesktop","DistUK_SortonOffice"
    $newUserCompany = "Bob Robert Industries Co."
}
if ($newUserOfficeLocation -eq 'Laneham')
{
    $newUserPath = "ou=users,ou=laneham,ou=sites,ou=ie,dc=bob,dc=com"
    $newUserStreetAddress = "23 Loxley Road"
    $newUserCity = 'Laneham'
    $newUserState = 'County Connelly'
    $newUserPostcode = 'LH8 9AH'
    $newUserCountry = 'IE'
    $additionalGroups = "SecUK_IT_RemoteDesktop","DistIE_LanehamOffice"
    $newUserCompany = "Bob Robert Industries Co."
}

$newUserPath

###################
# Create the user #
###################

# Create the user with all of the information that has been gathered
New-ADUser -GivenName $newUserFirstNameTitleCase `
-Surname $newUserLastNameTitleCase `
-Name $newUserDisplayName `
-DisplayName $newUserDisplayName `
-SamAccountName $newUserAccountName `
-UserPrincipalName $newUserEmailAddress `
-EmailAddress $newUserEmailAddress `
-Description $newUserJobTitle `
-Title $newUserJobTitle `
-Department $newUserDepartment `
-Manager $newUserManager `
-AccountPassword $newUserPasswordString `
-Office $newUserOfficeLocation `
-Path $newUserPath `
-StreetAddress $newUserStreetAddress `
-City $newUserCity `
-State $newUserState `
-PostalCode $newUserPostcode `
-Country $newUserCountry `
-Enabled $true `

# Set the default groups that all people should be included in
$groups = "SecUK_IT_MFAUsers","DistUK_AllUsers"

# For each of the above groups, add the user to them one at a time
foreach ($group in $groups) {
    Add-ADGroupMember -Identity $group -Members $newUserAccountName
}

# Add the user to any additional groups that they need
foreach ($group in $additionalGroups) {
    Add-ADGroupMember -Identity $group -Members $newUserAccountName
}

# Create a box to show the user's password, so that it can be copied and pasted.
Get-Input -TitleBox "Password" -Message "This is the new user's password:" -DefaultValue "" -TextBox $newUserPasswordTitleCase
