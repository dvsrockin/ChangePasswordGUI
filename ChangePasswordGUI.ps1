# Import the necessary .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'STOP'
Set-Location $PSScriptRoot
#$DomainDNList = get-Content $PSScriptRoot\Domainlist.ini

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Change Password"
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.StartPosition = "CenterScreen"

# Create labels
$lblDomain = New-Object System.Windows.Forms.Label
$lblDomain.Text = "Domain:"
$lblDomain.Location = New-Object System.Drawing.Point(20, 20)
$lblDomain.Size = New-Object System.Drawing.Size(100, 20)

$lblUserId = New-Object System.Windows.Forms.Label
$lblUserId.Text = "User ID:"
$lblUserId.Location = New-Object System.Drawing.Point(20, 60)
$lblUserId.Size = New-Object System.Drawing.Size(100, 20)

$lblOldPassword = New-Object System.Windows.Forms.Label
$lblOldPassword.Text = "Old Password:"
$lblOldPassword.Location = New-Object System.Drawing.Point(20, 140)
$lblOldPassword.Size = New-Object System.Drawing.Size(100, 20)

$lblNewPassword = New-Object System.Windows.Forms.Label
$lblNewPassword.Text = "New Password:"
$lblNewPassword.Location = New-Object System.Drawing.Point(20, 180)
$lblNewPassword.Size = New-Object System.Drawing.Size(100, 20)

$lblUserPath = New-Object System.Windows.Forms.Label
$lblUserPath.Text = "User Path:"
$lblUserPath.Location = New-Object System.Drawing.Point(20, 100)
$lblUserPath.Size = New-Object System.Drawing.Size(100, 20)

# Create input controls
$cbDomain = New-Object System.Windows.Forms.ComboBox
$cbDomain.Location = New-Object System.Drawing.Point(140, 20)
$cbDomain.Size = New-Object System.Drawing.Size(200, 20)
$cbDomain.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

Try{
        $rootDSE = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")
        $configurationNamingContext = $rootDSE.Properties["configurationNamingContext"].Value
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$configurationNamingContext")
        $searcher.Filter = "(objectClass=crossRef)"
        $searcher.SearchScope = "Subtree"
        $searcher.PropertiesToLoad.Add("nCName")

        $results = $searcher.FindAll()
        foreach ($result in $results) {
            $domainDN = $result.Properties["nCName"][0]
            if ($domainDN -notmatch "CN=Configuration|CN=Schema|DC=DomainDnsZones|DC=ForestDnsZones"){
                    $Domainname = $domainDN -replace "DC=","" -replace ",","."
                    $cbDomain.Items.Add($Domainname)
            }
        }
}
Catch{
         [System.Windows.Forms.MessageBox]::Show("Error occured while fetching domains, Please check network connectivity to your domain", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
         Exit
}

#$cbDomain.Items.AddRange(@("if.atcsg.net","za.if.atcsg.net","zb.if.atcsg.net","zd.if.atcsg.net","zf.if.atcsg.net","zh.if.atcsg.net","zi.if.atcsg.net","zj.if.atcsg.net","zk.if.atcsg.net","zl.if.atcsg.net"))
#$cbDomain.Items.AddRange(@($DomainDNList))
Try{
        $cbDomain.SelectedIndex = 0
    }
Catch{
        [System.Windows.Forms.MessageBox]::Show("Domain list not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Break
}

$txtUserId = New-Object System.Windows.Forms.TextBox
$txtUserId.Location = New-Object System.Drawing.Point(140, 60)
$txtUserId.Size = New-Object System.Drawing.Size(200, 20)

$txtUserPath = New-Object System.Windows.Forms.TextBox
$txtUserPath.Location = New-Object System.Drawing.Point(140, 100)
$txtUserPath.Size = New-Object System.Drawing.Size(300, 20)
$txtUserPath.ReadOnly = $true
$txtUserPath.AutoSize = $true


$txtOldPassword = New-Object System.Windows.Forms.TextBox
$txtOldPassword.Location = New-Object System.Drawing.Point(140, 140)
$txtOldPassword.Size = New-Object System.Drawing.Size(200, 20)
$txtOldPassword.UseSystemPasswordChar = $true

$txtNewPassword = New-Object System.Windows.Forms.TextBox
$txtNewPassword.Location = New-Object System.Drawing.Point(140, 180)
$txtNewPassword.Size = New-Object System.Drawing.Size(200, 20)
$txtNewPassword.UseSystemPasswordChar = $true

# Create buttons
$btnFindUser = New-Object System.Windows.Forms.Button
$btnFindUser.Text = "Find User"
$btnFindUser.Location = New-Object System.Drawing.Point(360, 60)
$btnFindUser.Size = New-Object System.Drawing.Size(80, 25)

$btnChangePassword = New-Object System.Windows.Forms.Button
$btnChangePassword.Text = "Change Password"
$btnChangePassword.Location = New-Object System.Drawing.Point(50, 240)
$btnChangePassword.Size = New-Object System.Drawing.Size(120, 30)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(200, 240)
$btnCancel.Size = New-Object System.Drawing.Size(120, 30)

# Global variable to store LDAP path
$global:userLDAPPath = ""

# Add event handler for Find User button
$btnFindUser.Add_Click({
    #$domain = (Get-ADDomain $($cbDomain.SelectedItem)).DistinguishedName
    $domain =  "DC=" + ($($cbDomain.SelectedItem -split '\.') -join ',DC=')
    $userId = $txtUserId.Text

    if ([string]::IsNullOrWhiteSpace($domain) -or [string]::IsNullOrWhiteSpace($userId)) {
        $txtUserPath.Text = "Please enter valid Domain and User ID."
        return
    }

    try {
        # Query Active Directory for the user
        $ldapFilter = "(&(objectClass=user)(sAMAccountName=$userId))"
        $searchRoot = "LDAP://$domain"
        $searcher = New-Object DirectoryServices.DirectorySearcher([adsi]$searchRoot)
        $searcher.Filter = $ldapFilter

        $result = $searcher.FindOne()
        if ($result) {
            $global:userLDAPPath = $result.Path
            $txtUserPath.Text = $global:userLDAPPath
        } else {
            $txtUserPath.Text = "User not found."
        }
    } catch {
        $txtUserPath.Text = "Error querying Active Directory: $_"
    }
})

# Add event handler for Change Password button
$btnChangePassword.Add_Click({
    $oldPassword = $txtOldPassword.Text
    $newPassword = $txtNewPassword.Text

    if ([string]::IsNullOrWhiteSpace($global:userLDAPPath)) {
        [System.Windows.Forms.MessageBox]::Show("Please find the user first.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    if ([string]::IsNullOrWhiteSpace($oldPassword) -or [string]::IsNullOrWhiteSpace($newPassword)) {
        [System.Windows.Forms.MessageBox]::Show("All fields are required.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    try {
        # Attempt to change the password
        [adsi]$user = $global:userLDAPPath
        $user.ChangePassword($oldPassword, $newPassword)

        [System.Windows.Forms.MessageBox]::Show("Password changed successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $form.Close()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to change password. Ensure your inputs are correct.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Add event handler for Cancel button
$btnCancel.Add_Click({
    $form.Close()
})

# Add controls to the form
$form.Controls.Add($lblDomain)
$form.Controls.Add($cbDomain)
$form.Controls.Add($lblUserId)
$form.Controls.Add($txtUserId)
$form.Controls.Add($lblUserPath)
$form.Controls.Add($txtUserPath)
$form.Controls.Add($btnFindUser)
$form.Controls.Add($lblOldPassword)
$form.Controls.Add($txtOldPassword)
$form.Controls.Add($lblNewPassword)
$form.Controls.Add($txtNewPassword)
$form.Controls.Add($btnChangePassword)
$form.Controls.Add($btnCancel)

# Display the form
[void]$form.ShowDialog()
