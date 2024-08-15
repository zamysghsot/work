<#	
	.NOTES
	===========================================================================
	 Created on:   	29.09.2024
	 Created by:   	Patrick Carvalho
	 Organization: 	e-novinfo SA
	 Filename:     	Exchange.ps1
     Modify on :    05.08.2024
	===========================================================================
	.DESCRIPTION
		Script de manipulation O365
    .TODO
        Faire les redirections de mail avec et sans garder le mail dans la boite mail

#>


Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Charger l'assembly Windows Forms
Add-Type -AssemblyName System.Windows.Forms
$firstLoad = $true
#$addresses = null

# Créer la fenêtre (form)
$formWidth = 800
$formHeigth = 600
$form = New-Object System.Windows.Forms.Form
$form.Text = "Manipulation Exchange (O365)"
$form.Size = New-Object System.Drawing.Size($formWidth,$formHeigth)
$form.StartPosition = "CenterScreen"

# Créer une étiquette (lblAction)
$lblAction = New-Object System.Windows.Forms.label
$lblAction.Text = "Choisir l'action"
$lblAction.AutoSize = $true
$lblAction.Location = New-Object System.Drawing.Point(100,30)

# Créer une étiquette (lblAction)
$cmbAction = New-Object System.Windows.Forms.ComboBox
$width = $formWidth - 200
$cmbAction.Size = New-Object System.Drawing.Size($width, 30)
$cmbAction.Location = New-Object System.Drawing.Point(100,50)
$cmbAction.Items.Add("Fixer les droits sur une boite mail")
$cmbAction.Items.Add("Gérer les droits sur les calendriers")
$cmbAction.Items.Add("Supprimer les droits sur une boite mail")
$cmbAction.Items.Add("Voir les droits de l'utilisateur sur les autres boites mail")
$cmbAction.Items.Add("Voir les droits sur une boite mail")
$cmbAction.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

#Créer un groupe pour les éléments
$grpboxContenu = New-Object System.Windows.Forms.GroupBox
$formContenuHeight = $form.Size.Height - 130
$grpboxContenu.Size = New-Object System.Drawing.Size($width, $formContenuHeight)
$grpboxContenu.Location = New-Object System.Drawing.Point(100,70)


$lblLoad = New-Object System.Windows.Forms.Label
$lblLoad.Text = "Chargement ..."
$lblHeight = ($grpboxContenu.ClientSize.Height - 20) / 2
$lblLoad.Location = New-Object System.Drawing.Point(10, $lblHeight)
$lblLoad.Size = New-Object System.Drawing.Size(580, 20)
$lblLoad.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblLoad.Visible = $false

# Ajouter les contrôles à la fenêtre
$form.Controls.Add($lblAction)
$form.Controls.Add($cmbAction)
$form.Controls.Add($grpboxContenu)
$grpboxContenu.controls.Add($lblLoad)

$width = $formWidth - 220

$lblRightIdentity = New-Object System.Windows.Forms.label
$lblRightIdentity.Text = "Identity"
$lblRightIdentity.Location = New-Object System.Drawing.Point(10,20)
$lblRightIdentity.Size = New-Object System.Drawing.Size($width, 20)
    
$lblSearch = New-Object System.Windows.Forms.label
$lblSearch.Text = "Search :"
$lblSearch.Location = New-Object System.Drawing.Point(10,10)
$lblSearch.Size = New-Object System.Drawing.Size(50, 20)

$dgvForm = New-Object System.Windows.Forms.Form
$dgvForm.Text = "Sélectionner une ligne"
$dgvForm.Size = New-Object System.Drawing.Size($width, 400)
$dgvForm.StartPosition = "Manual"
$dgvForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow

$tbxFormIdentity = New-Object System.Windows.Forms.TextBox
$tbxFormIdentity.Location = New-Object System.Drawing.Point(60,7)
$tbxFormIdentity.Size = New-Object System.Drawing.Size(130, 20)

$dgvRight = New-Object System.Windows.Forms.DataGridView
$dgvRight.Location = New-Object System.Drawing.Point(0, 40)
$dgvRight.Size = New-Object System.Drawing.Size($width, 360)
$dgvRight.ReadOnly = $true
$dgvRight.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgvRight.AutoGenerateColumns = $true
$dgvRight.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
$dgvRight.DataSource = $adresses

$dgvRight.ColumnCount = 2
$dgvRight.Columns[0].Name = "PrimarySmtpAddress"
$dgvRight.Columns[1].Name = "DisplayName"

$tbxFormIdentity.Add_TextChanged({
    $searchText = $tbxFormIdentity.Text
    if ($searchText -ne "") {
        $filteredData = $addresses | Where-Object {
            $_.PrimarySmtpAddress -like "*$searchText*" -or
            $_.DisplayName -like "*$searchText*"
        }
    } else {
        $filteredData = $addresses
    }

    $dgvRight.Rows.Clear()

    foreach ($address in $filteredData) {
        $dgvRight.rows.Add($address.DisplayName, $address.PrimarySmtpAddress)
    }
#    $dgvRight.DataSource = $filteredData
})

$dgvForm.Controls.Add($lblSearch)
$dgvForm.Controls.Add($tbxFormIdentity)
#$dgvForm.Controls.Add($btnFormIdentity)
$dgvForm.Controls.Add($dgvRight)

$cbxRightIdentity = New-Object System.Windows.Forms.ComboBox
$cbxRightIdentity.Location = New-Object System.Drawing.Point(10,45)
$cbxRightIdentity.Size = New-Object System.Drawing.Size($width, 20)
$cbxRightIdentity.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown

$lblRightUser = New-Object System.Windows.Forms.label
$lblRightUser.Text = "User"
$lblRightUser.Location = New-Object System.Drawing.Point(10,85)
$lblRightUser.Size = New-Object System.Drawing.Size($width, 20)

$cbxRightUser = New-Object System.Windows.Forms.ComboBox
$cbxRightUser.Location = New-Object System.Drawing.Point(10,105)
$cbxRightUser.Size = New-Object System.Drawing.Size($width, 20)
$cbxRightUser.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown

$lblRightAutomapping = New-Object System.Windows.Forms.label
$lblRightAutomapping.Text = "Automapping"
$lblRightAutomapping.Location = New-Object System.Drawing.Point(10,195)
$lblRightAutomapping.Size = New-Object System.Drawing.Size(80, 20)

$chxRightAutomapping = New-Object System.Windows.Forms.CheckBox
$chxRightAutomapping.Location = New-Object System.Drawing.Point(90,192)
$chxRightAutomapping.Size = New-Object System.Drawing.Size(25, 25)
$chxRightAutomapping.BringToFront()

$lblRightAccess = New-Object System.Windows.Forms.label
$lblRightAccess.Text = "Access Right"
$lblRightAccess.Location = New-Object System.Drawing.Point(10,140)
$lblRightAccess.Size = New-Object System.Drawing.Size(80, 20)

$cbxRightAccess = New-Object System.Windows.Forms.ComboBox
$cbxRightAccess.Location = New-Object System.Drawing.Point(10,165)
$cbxRightAccess.Size = New-Object System.Drawing.Size($width, 20)
$cbxRightAccess.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown

$cbxRightAccess.Items.Add("ChangeOwner")
$cbxRightAccess.Items.Add("ChangePermission")
$cbxRightAccess.Items.Add("DeleteItem")
$cbxRightAccess.Items.Add("ExternalAccount")
$cbxRightAccess.Items.Add("FullAccess")
$cbxRightAccess.Items.Add("ReadPermission")

$cbxRightAccess.SelectedIndex = 4

$btnRightValidate = New-Object System.Windows.Forms.Button
$btnRightValidate.Text = "Soumettre"
$btnRightValidate.Location = New-Object System.Drawing.Point(10,220)
$btnRightValidate.Size = New-Object System.Drawing.Size(580, 20)

$tbxReturn = New-Object System.Windows.Forms.TextBox
$tbxReturn.Location = New-Object System.Drawing.Point(10,245)
$tbxReturnHeigth = $grpboxContenu.Height - 240 - 15
$tbxReturn.Size = New-Object System.Drawing.Size(580,$tbxReturnHeigth)
$tbxReturn.BackColor = [System.Drawing.Color]::FromArgb(1, 36, 86)
$tbxReturn.ForeColor = [System.Drawing.Color]::White
$tbxReturn.Multiline = $true
$tbxReturn.ReadOnly = $true
$tbxReturn.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical

$cmbAction.Add_SelectedIndexChanged({
    if($firstLoad)
    {
        $lblLoad.Visible = $true
        $script:addresses = Get-Mailbox -ResultSize Unlimited | Select PrimarySmtpAddress, DisplayName

        $form.Activate()

        foreach ($address in $addresses) {
            $dgvRight.rows.Add($address.DisplayName, $address.PrimarySmtpAddress)
        }
        $lblLoad.Visible = $false

        $firstLoad = $false
    }
    $grpboxContenu.Controls.Clear()
    switch ($cmbAction.SelectedIndex){
        0 {
            #Fixer les droits sur une boite mail
            $grpboxContenu.Controls.Add($lblRightIdentity)
            $grpboxContenu.Controls.Add($lblRightUser)
            $grpboxContenu.Controls.Add($lblRightAutomapping)
            $grpboxContenu.Controls.Add($lblRightAccess)
            $grpboxContenu.Controls.Add($chxRightAutomapping)
            $grpboxContenu.Controls.Add($cbxRightIdentity)
            $grpboxContenu.Controls.Add($cbxRightUser)
            $grpboxContenu.Controls.Add($cbxRightAccess)
            $grpboxContenu.Controls.Add($btnRightValidate)
            $grpboxContenu.Controls.Add($tbxReturn)
            break
        }

        1{
            #Gérer les droits sur les calendriers
            break
        }

        2{
            #Supprimer les droits sur une boite mail
            $grpboxContenu.Controls.Add($lblRightIdentity)
            $grpboxContenu.Controls.Add($lblRightUser)
            $grpboxContenu.Controls.Add($lblRightAccess)
            $grpboxContenu.Controls.Add($cbxRightIdentity)
            $grpboxContenu.Controls.Add($cbxRightUser)
            $grpboxContenu.Controls.Add($cbxRightAccess)
            $grpboxContenu.Controls.Add($btnRightValidate)
            $grpboxContenu.Controls.Add($tbxReturn)
            break
        }

        3{
            #Voir les droits de l'utilisateur sur les autres boites mail
            break
        }

        4{
            #Voir les droits sur une boite mail
            $grpboxContenu.Controls.Clear()
            $grpboxContenu.Controls.Add($lblRightIdentity)
            $grpboxContenu.Controls.Add($cbxRightIdentity)
            $grpboxContenu.Controls.Add($btnRightValidate)
            $grpboxContenu.Controls.Add($tbxReturn)
            break
        }
    }
    
})

$cbxRightIdentity.Add_Click({
    $dgvForm.Location = $form.PointToScreen($cbxRightIdentity.Location)
    $locationx = $dgvForm.Location.X + $grpboxContenu.Location.X
    $locationy = $dgvForm.Location.Y + $cbxRightIdentity.Height + $grpboxContenu.Location.Y
    $dgvForm.Location = New-Object System.Drawing.Point($locationx, $locationy)
    $tbxFormIdentity.Focus()
    $tbxFormIdentity.SelectAll()
    $global:field = "identity"
    $dgvForm.ShowDialog()
})

$dgvRight.Add_CellDoubleClick({
    $selectedRow = $dgvRight.SelectedRows[0]
    $values = $selectedRow.Cells | ForEach-Object { $_.Value }
    switch($field){
        "identity"{
            $cbxRightIdentity.Text = $values[1]
        }
        "user"{
            $cbxRightUser.Text = $values[1]
        }
    }
    $dgvForm.Hide()
})

$cbxRightUser.Add_Click({
    $dgvForm.Location = $form.PointToScreen($cbxRightUser.Location)
    $locationx = $dgvForm.Location.X + $grpboxContenu.Location.X
    $locationy = $dgvForm.Location.Y + $cbxRightIdentity.Height + $grpboxContenu.Location.Y
    $dgvForm.Location = New-Object System.Drawing.Point($locationx, $locationy)
    $tbxFormIdentity.Focus()
    $tbxFormIdentity.SelectAll()
    $global:field = "User"
    $dgvForm.ShowDialog()
})

$btnRightValidate.Add_Click({
    $tbxReturn.Text = ""
    switch ($cmbAction.SelectedIndex){
        0 {
            #Fixer les droits sur une boite mail
            $command = "Add-MailboxPermission -Identity "
            $command = $command + $cbxRightIdentity.Text
            $command = $command + " -User "
            $command = $command + $cbxRightUser.Text
            $command = $command + " -AccessRights "
            $command = $command + $cbxRightAccess.Text
            if($chxRightAutomapping.Checked){
                $command = $command + " -AutoMapping $$true" 
            } else {
                $command = $command + " -AutoMapping $$false"
            }
            
            try{
                $global:return = Invoke-Expression $command -ErrorAction Stop
                $tbxReturn.Text = "Permission added successfully:`n "
                $tbxReturn.Text = "$tbxReturn.Text $return"
            }
            catch{
                $tbxReturn.Text = "An error occurred:`n` "
                $tbxReturn.Text = "$tbxReturn.Text $_.Exception.Message"
            }
            
            break
        }

        1{
            #Gérer les droits sur les calendriers
            break
        }

        2{
            #Supprimer les droits sur une boite mail
            $command = "Remove-MailboxPermission -Identity "
            $command = $command + $cbxRightIdentity.Text
            $command = $command + " -User "
            $command = $command + $cbxRightUser.Text
            $command = $command + " -AccessRights "
            $command = $command + $cbxRightAccess.Text
            
            $global:return = Invoke-Expression $command
            break
        }

        3{
            #Voir les droits de l'utilisateur sur les autres boites mail
            break
        }

        4{
            #Voir les droits sur une boite mail
            $command = "Get-MailboxPermission -Identity "
            $command += $cbxRightIdentity.Text

            printTbxResult($command)

#            $global:return = Invoke-Expression $command -ErrorAction Stop
            
            
#            $formattedText = "{0,-30} {1,-30} {2,-30}" -f "Identity", "User", "AccessRigths" # En-têtes
#            $formattedText += "-------------------------------------------------------------------------------------------------------------------------------"+ [Environment]::NewLine
#            $tbxReturn.Text = "Identity`tUser`t`tAccessRigths"+ [Environment]::NewLine
#            $tbxReturn.Text += "-------------------------------------------------------------------------------------------------------------------------------"+ [Environment]::NewLine
#            foreach ($permission in $return) {
                # Display each property of the permission object
#                $formattedText += "{0,-20} {1,-30} {2,-30}" -f $permission.Identity, $permission.User, $permission.AccessRights
#                $formattedText += [Environment]::NewLine
#                $tbxReturn.Text += "$($permission.Identity)`t "
#                $tbxReturn.Text += "$($permission.User)`t`t "
#                $tbxReturn.Text += "$($permission.AccessRights)"+ [Environment]::NewLine
#            }
#            $tbxReturn.Text = $formattedText
            break
        }
    }
#    $tbxReturn.Text = $return
})

# Afficher la fenêtre
$form.Add_Shown({ $form.Activate() })
$form.add_FormClosed({ Disconnect-ExchangeOnline -Confirm:$false })

function printTbxResult($command){
    $global:return = Invoke-Expression $command -ErrorAction Stop
                       
    $formattedText = "{0,-30} {1,-30} {2,-30}" -f "Identity", "User", "AccessRigths" # En-têtes
    $formattedText += "-------------------------------------------------------------------------------------------------------------------------------"+ [Environment]::NewLine
#            $tbxReturn.Text = "Identity`tUser`t`tAccessRigths"+ [Environment]::NewLine
#            $tbxReturn.Text += "-------------------------------------------------------------------------------------------------------------------------------"+ [Environment]::NewLine
    foreach ($permission in $return) {
        # Display each property of the permission object
        $formattedText += "{0,-20} {1,-30} {2,-30}" -f $permission.Identity, $permission.User, $permission.AccessRights
        $formattedText += [Environment]::NewLine
#                $tbxReturn.Text += "$($permission.Identity)`t "
#                $tbxReturn.Text += "$($permission.User)`t`t "
#                $tbxReturn.Text += "$($permission.AccessRights)"+ [Environment]::NewLine
    }
    $tbxReturn.Text = $formattedText
}
[System.Windows.Forms.Application]::Run($form)
