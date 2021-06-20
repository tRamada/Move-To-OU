# Move User to OU (MULTI USERS)
#
# Author: (Mutega IT AB Intern) Tiago Ramada
# Date 2021/05/20
#
# Version 1.0


# Assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework


# Set log file location
$logFile = ".\logs\Move_To_OU_$(Get-Date -Format dd-MM-yyyy-hh-mm-ss).log"
Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Program as been started" | Out-File $logFile -Append

$ApplicationName = "Move-To-OU"
$DefaultLog = "Welcome to $ApplicationName"


# Draw window and controls
$Window = New-Object System.Windows.Forms.Form
$Window.Text = "Move-OU"
$Window.Size = New-Object System.Drawing.Size(800, 500)
$Window.FormBorderStyle = "FixedDialog"
$Window.TopMost = $true
$Window.MaximizeBox = $false
$Window.MinimizeBox = $false
$Window.ControlBox = $true
$Window.StartPosition = "CenterScreen"
$Window.Font = "Segoe UI"

# Adding an Icon (Remove '#')
#$Window_Icon = New-Object system.drawing.icon ("Insert icon path here.")
#$Window.Icon = $Window_Icon

# **RICHTEXTBOX-LOG**
$richtextbox_Logs = New-Object 'System.Windows.Forms.RichTextBox'
$richtextbox_Logs.BackColor = 'InactiveBorder'
$richtextbox_Logs.Font = "Consolas, 8.25pt"
$richtextbox_Logs.Location = '10, 320'
$richtextbox_Logs.Name = "richtextbox_Logs"
$richtextbox_Logs.ReadOnly = $True
$richtextbox_Logs.Size = '765, 130'
$richtextbox_Logs.TabIndex = 35
$richtextbox_Logs.Text = ""
$richtextbox_Logs.add_TextChanged($richtextbox_Logs_TextChanged)
$Window.Controls.Add($richtextbox_Logs)

function Add-Logs {
    [CmdletBinding()]
    param ($text)
    $richtextbox_logs.Text += "[$(Get-Date -Format "MM/dd/yyyy HH:mm")] - $text`r"
    Set-Alias alogs Add-Logs -Description "Add content to the RichTextBoxLogs"
    Set-Alias Add-Log Add-Logs -Description "Add content to the RichTextBoxLogs"
}#endregion Add Logs

$richtextbox_Logs_TextChanged = {
    $richtextbox_Logs.SelectionStart = $richtextbox_Logs.Text.Length
    $richtextbox_Logs.ScrollToCaret()
    if ($error[0]) { Add-logs -text $($error[0].Exception.Message) }
}

# Adding Default Log
Add-Logs -text $DefaultLog

# **USER**
# Adding 'User' Label (Single-User)
$lbl_user = New-Object System.Windows.Forms.Label
$lbl_user.Location = New-Object System.Drawing.Point(30, 40)
$lbl_user.AutoSize = $true
$lbl_user.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold) 
$lbl_user.Text = "User:"
$lbl_user.Visible = $false
$Window.Controls.Add($lbl_user)

# Adding 'User' TextBox (Single-User)
$textBox_user = New-Object System.Windows.Forms.TextBox
$textBox_user.Location = New-Object System.Drawing.Point(90, 40)
$textBox_user.Size = New-Object System.Drawing.Size(200, 20)
$textBox_user.Visible = $false
$Window.Controls.Add($textBox_user)

# Adding 'File' Label (Multi-Users)
$lbl_file = New-Object System.Windows.Forms.Label
$lbl_file.Location = New-Object System.Drawing.Point(30, 70)
$lbl_file.AutoSize = $true
$lbl_file.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold) 
$lbl_file.Text = "File:"
$lbl_file.Visible = $false
$Window.Controls.Add($lbl_file)

# Adding 'Total' Label (Multi-Users)
$lbl_users_loaded = New-Object System.Windows.Forms.Label
$lbl_users_loaded.Location = New-Object System.Drawing.Point(8, 90)
$lbl_users_loaded.AutoSize = $true
$lbl_users_loaded.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold) 
$lbl_users_loaded.Text = "Users loaded:"
$lbl_users_loaded.Visible = $false
$Window.Controls.Add($lbl_users_loaded)


# **DETAILS**
# Adding 'file-selected' label (Multi-User)
$lbl_fs = New-Object System.Windows.Forms.Label
$lbl_fs.Location = New-Object System.Drawing.Point(70, 70)
$lbl_fs.AutoSize = $true
$lbl_fs.Font = New-Object System.Drawing.Font("Segoe UI", 9) 
$lbl_fs.Text = "Empty"
$lbl_fs.Visible = $false
$Window.Controls.Add($lbl_fs)

# Adding 'Total of users' label (Multi-User)
$lbl_total = New-Object System.Windows.Forms.Label
$lbl_total.Location = New-Object System.Drawing.Point(90, 90)
$lbl_total.AutoSize = $true
$lbl_total.Font = New-Object System.Drawing.Font("Segoe UI", 9) 
$lbl_total.Text = ""
$lbl_total.Visible = $false
$Window.Controls.Add($lbl_total)


# Adding 'User-Details' Label (Single-User)
$lbl_user_details = New-Object System.Windows.Forms.Label
$lbl_user_details.Location = New-Object System.Drawing.Point(8, 60)
$lbl_user_details.AutoSize = $true
$lbl_user_details.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold) 
$lbl_user_details.Text = "User details:"
$lbl_user_details.Visible = $false
$Window.Controls.Add($lbl_user_details)

# Adding 'user_details_name" Label (Single-User)
$lbl_user_details_name = New-Object System.Windows.Forms.Label
$lbl_user_details_name.Location = New-Object System.Drawing.Point(43, 78)
$lbl_user_details_name.AutoSize = $true
$lbl_user_details_name.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold) 
$lbl_user_details_name.Text = "Name -"
$lbl_user_details_name.Visible = $false
$Window.Controls.Add($lbl_user_details_name)

# Adding 'user_details_name_info" Label (Single-User)
$lbl_user_details_name_info = New-Object System.Windows.Forms.Label
$lbl_user_details_name_info.Location = New-Object System.Drawing.Point(90, 78)
$lbl_user_details_name_info.AutoSize = $true
$lbl_user_details_name_info.Font = New-Object System.Drawing.Font("Segoe UI", 9) 
$lbl_user_details_name_info.Text = "Empty"
$lbl_user_details_name_info.Visible = $false
$Window.Controls.Add($lbl_user_details_name_info)

# Adding 'user_details_ou" Label (Single-User)
$lbl_user_details_ou = New-Object System.Windows.Forms.Label
$lbl_user_details_ou.Location = New-Object System.Drawing.Point(58.8, 100)
$lbl_user_details_ou.AutoSize = $true
$lbl_user_details_ou.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold) 
$lbl_user_details_ou.Text = "OU -"
$lbl_user_details_ou.Visible = $false
$Window.Controls.Add($lbl_user_details_ou)

# Adding 'user_details_ou_info" Label (Single-User)
$lbl_user_details_ou_info = New-Object System.Windows.Forms.Label
$lbl_user_details_ou_info.Location = New-Object System.Drawing.Point(90, 100)
$lbl_user_details_ou_info.AutoSize = $true
$lbl_user_details_ou_info.Font = New-Object System.Drawing.Font("Segoe UI", 9) 
$lbl_user_details_ou_info.Text = "Empty"
$lbl_user_details_ou_info.Visible = $false
$Window.Controls.Add($lbl_user_details_ou_info)


# **DOMAIN**
# Adding 'Domain' Label
$lbl_domain = New-Object System.Windows.Forms.Label
$lbl_domain.Location = New-Object System.Drawing.Point(20, 140)
$lbl_domain.Size = New-Object System.Drawing.Size(240, 32)
$lbl_domain.AutoSize = $true
$lbl_domain.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold) 
$lbl_domain.Text = "Domain:"
$Window.Controls.Add($lbl_domain)
 

# Adding 'Domain' ComboBox
$comboBox_Domain = New-Object System.Windows.Forms.ComboBox
$comboBox_Domain.Location = New-Object System.Drawing.Point(90, 140)
$comboBox_Domain.Size = New-Object System.Drawing.Size(200, 32)
$comboBox_Domain.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
$Window.Controls.Add($comboBox_Domain)

import-csv ".\domain_Multiusers.csv" | ForEach-Object {
    $comboBox_Domain.Items.Add($_.DomainName)
}


# **OU**
# Adding 'OU' Label
$lbl_ou = New-Object System.Windows.Forms.Label
$lbl_ou.Location = New-Object System.Drawing.Point(40, 192)
$lbl_ou.Size = New-Object System.Drawing.Size(240, 32)
$lbl_ou.AutoSize = $true
$lbl_ou.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lbl_ou.Text = "OU:"
$Window.Controls.Add($lbl_ou)

# Adding 'OU' ComboBox
$comboBox_ou = New-Object System.Windows.Forms.ComboBox
$comboBox_ou.Location = New-Object System.Drawing.Point(90, 190)
$comboBox_ou.Size = New-Object System.Drawing.Size(200, 32)
$comboBox_ou.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
$comboBox_ou.add_SelectedIndexChanged($comboBox_ou_SelectedIndexChanged)
$Window.Controls.Add($comboBox_ou)
import-csv ".\ou_Multiusers.csv" | ForEach-Object {
    $comboBox_ou.Items.Add($_.Site)
}

$comboBox_ou_SelectedIndexChanged = {
    #Read file with Site and OU
    $getFile = ".\ou_Multiusers.csv"
    $file = Import-Csv $getFile

    ForEach ($sites in $file) {
        $textBox_ou_path.Clear()
        if ($sites.Site -eq $comboBox_ou.Text) {
            $textBox_ou_path.Text = $sites.OU
            Break
        }
    } 
}
 

# Adding 'OU-Path' Label
$lbl_ou_path = New-Object System.Windows.Forms.Label
$lbl_ou_path.Location = New-Object System.Drawing.Point(25, 220)
$lbl_ou_path.Size = New-Object System.Drawing.Size(240, 32)
$lbl_ou_path.AutoSize = $true
$lbl_ou_path.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold) 
$lbl_ou_path.Text = "OU Path:"
$Window.Controls.Add($lbl_ou_path)

# Adding 'OU-Path' TextBox
$textBox_ou_path = New-Object System.Windows.Forms.TextBox
$textBox_ou_path.Location = New-Object System.Drawing.Point(90, 220)
$textBox_ou_path.Size = New-Object System.Drawing.Size(200, 20)
$Window.Controls.Add($textBox_ou_path)



# **BUTTONS**
# Adding 'Refresh' button (Users-file from Multi-Users)
$bt_rf_users = New-Object System.Windows.Forms.Button
$bt_rf_users.Location = New-Object System.Drawing.Point(260, 70)
$bt_rf_users.Size = New-Object System.Drawing.Size(65, 20)
$bt_rf_users.TextAlign = "MiddleCenter"
$bt_rf_users.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$bt_rf_users.Text = "Refresh"
$bt_rf_users.Visible = $false
$bt_rf_users.Add_Click( {
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Users: Refresh button as been clicked" | Out-File $logFile -Append

        #Read file with Samaccountname, OU and Domain
        $ADUsersFile = ".\users_Multiusers.csv"
        $Global:usernames = Import-Csv -Path $ADUsersFile -Encoding UTF8

        $GLobal:cont = 0
        Foreach ($username in $usernames) {
            $GLobal:cont += 1
        }
        $lbl_total.Text = $GLobal:cont
    })
$Window.Controls.Add($bt_rf_users)



# Adding 'Refresh' button (Domain-file)
$bt_rf = New-Object System.Windows.Forms.Button
$bt_rf.Location = New-Object System.Drawing.Point(370, 140)
$bt_rf.Size = New-Object System.Drawing.Size(65, 20)
$bt_rf.TextAlign = "MiddleCenter"
$bt_rf.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$bt_rf.Text = "Refresh"
$bt_rf.Add_Click( {
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Domain: Refresh button as been clicked" | Out-File $logFile -Append

        # This clear all items from comboBox
        $comboBox_Domain.Items.Clear()

        import-csv ".\domain_Multiusers.csv" | ForEach-Object {
            $comboBox_Domain.Items.Add($_.DomainName)
        }
    })
$Window.Controls.Add($bt_rf)

# Adding 'Edit' button (Domain-file)
$bt_edit_domain = New-Object System.Windows.Forms.Button
$bt_edit_domain.Location = New-Object System.Drawing.Point(300, 140)
$bt_edit_domain.Size = New-Object System.Drawing.Size(65, 20)
$bt_edit_domain.TextAlign = "MiddleCenter"
$bt_edit_domain.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$bt_edit_domain.Text = "Edit"
$bt_edit_domain.Add_Click( {
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Domain: Edit button as been clicked" | Out-File $logFile -Append
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Domain: Openning domain file.." | Out-File $logFile -Append
        Add-Logs -text "Domain: Openning Domain file.."

        notepad.exe ./domain_Multiusers.csv
    })
$Window.Controls.Add($bt_edit_domain)

# Adding 'Edit' button (OU-file)
$bt_edit_ou = New-Object System.Windows.Forms.Button
$bt_edit_ou.Location = New-Object System.Drawing.Point(300, 190)
$bt_edit_ou.Size = New-Object System.Drawing.Size(65, 20)
$bt_edit_ou.TextAlign = "MiddleCenter"
$bt_edit_ou.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$bt_edit_ou.Text = "Edit"
$bt_edit_ou.Add_Click( {
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - OU: Edit button as been clicked" | Out-File $logFile -Append
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - OU: Openning OU file.." | Out-File $logFile -Append
        Add-Logs -text "OU: Openning OU file.."

        notepad.exe ./ou_Multiusers.csv
    })
$Window.Controls.Add($bt_edit_ou)

# Adding 'Refresh' button (Domain-file)
$bt_rf_ou = New-Object System.Windows.Forms.Button
$bt_rf_ou.Location = New-Object System.Drawing.Point(370, 190)
$bt_rf_ou.Size = New-Object System.Drawing.Size(65, 20)
$bt_rf_ou.TextAlign = "MiddleCenter"
$bt_rf_ou.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$bt_rf_ou.Text = "Refresh"
$bt_rf_ou.Add_Click( {
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Domain: Refresh button as been clicked" | Out-File $logFile -Append
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Domain: Refreshing comboBox items.." | Out-File $logFile -Append

        # This clear all items from comboBox
        $comboBox_ou.Items.Clear()

        import-csv ".\ou_Multiusers.csv" | ForEach-Object {
            $comboBox_ou.Items.Add($_.Site)
        }
    })
$Window.Controls.Add($bt_rf_ou)

# Adding '<- Back' button (Single-User and Multi-User)
$bt_back = New-Object System.Windows.Forms.Button
$bt_back.Location = New-Object System.Drawing.Point(10, 10)
$bt_back.Size = New-Object System.Drawing.Size(65, 20)
$bt_back.TextAlign = "MiddleCenter"
$bt_back.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$bt_back.Text = "<- Back"
$bt_back.Visible = $false
$bt_back.Add_Click( {
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Button Back as been clicked" | Out-File $logFile -Append
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Showing Single-user and Multi-Users buttons.." | Out-File $logFile -Append



        $lbl_domain.Select()

        $Global:multi_click = $false
        $Global:single_click = $false

        $bt_single.Visible = $true
        $bt_multi.Visible = $true
        $bt_back.Visible = $false
        $bt_search.Visible = $false
        $lbl_user.Visible = $false
        $textBox_user.Visible = $false
        $lbl_user_details.Visible = $false
        $lbl_user_details_name.Visible = $false
        $lbl_user_details_name_info.Visible = $false
        $lbl_user_details_ou.Visible = $false
        $lbl_user_details_ou_info.Visible = $false
        $bt_back.Visible = $false
        $lbl_file.Visible = $false
        $lbl_fs.Visible = $false
        $bt_edit.Visible = $false
        $lbl_users_loaded.Visible = $false
        $lbl_total.Visible = $false
        $bt_rf_users.Visible = $false

    })
$Window.Controls.Add($bt_back)

# Adding 'Edit' button (Multi-User)
$bt_edit = New-Object System.Windows.Forms.Button
$bt_edit.Location = New-Object System.Drawing.Point(190, 70)
$bt_edit.Size = New-Object System.Drawing.Size(65, 20)
$bt_edit.TextAlign = "MiddleCenter"
$bt_edit.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$bt_edit.Text = "Edit"
$bt_edit.Visible = $false
$bt_edit.Add_Click( {
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - User: Edit button as been clicked" | Out-File $logFile -Append
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - User: Openning Users file.." | Out-File $logFile -Append
        Add-Logs -text "User: Openning Users file.."

        notepad.exe ./users_Multiusers.csv
    })
$Window.Controls.Add($bt_edit)


# Adding 'Search' button (SINGLE-USER)
$bt_search = New-Object System.Windows.Forms.Button
$bt_search.Location = New-Object System.Drawing.Point(300, 40)
$bt_search.Size = New-Object System.Drawing.Size(65, 20)
$bt_search.TextAlign = "MiddleCenter"
$bt_search.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold) 
$bt_search.Text = "Search"
$bt_search.Visible = $false
$bt_search.Add_Click( {
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - User: Button search as been clicked" | Out-File $logFile -Append


        $Global:iD = $textBox_user.Text

        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - User: Searching for $($iD).." | Out-File $logFile -Append
        Add-Logs -text "User: Searching for $($iD).."
        try {
            $o = Get-ADUser "$iD" | Select-Object Name | Format-List
            $rm_ws_o = $o | Out-String
            $rm_ws_o = $rm_ws_o.TrimStart()
            $rm_ws_o = $rm_ws_o.TrimEnd()
            $rm_ws_o = $rm_ws_o -split "Name : "

            $lbl_user_details_name_info.Text = $rm_ws_o

            $user = Get-ADUser -Identity $iD -Properties CanonicalName
            $userOU = "OU=" + ($user.DistinguishedName -split ",OU=", 2)[1]
            $lbl_user_details_ou_info.Text = $userOU
        }
        catch {
            Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - User: $($iD) as not been found" | Out-File $logFile -Append 
            Add-Logs -text "User: $($iD) as not been found"
        }
        

        
    })
$Window.Controls.Add($bt_search)

# Adding 'Cancel' button
$bt_cancel = New-Object System.Windows.Forms.Button
$bt_cancel.Location = New-Object System.Drawing.Point(378, 280)
$bt_cancel.Size = New-Object System.Drawing.Size(100, 25)
$bt_cancel.TextAlign = "MiddleCenter"
$bt_cancel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold) 
$bt_cancel.Text = "Cancel"
$bt_cancel.Add_Click( {
        $Window.Close()
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Button cancel as been clicked" | Out-File $logFile -Append

        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - The Move have been Canceled!" | Out-File $logFile -Append
        Add-Logs -text "The Move have been Canceled!"

    })
$Window.Controls.Add($bt_cancel)

# Adding 'Move' button
$bt_move = New-Object System.Windows.Forms.Button
$bt_move.Location = New-Object System.Drawing.Point(260, 280)
$bt_move.Size = New-Object System.Drawing.Size(100, 25)
$bt_move.TextAlign = "MiddleCenter"
$bt_move.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$bt_move.Text = "Move"
$bt_move.Add_Click( {
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Move button as been clicked" | Out-File $logFile -Append
   

        $domain = $comboBox_Domain.SelectedItem
        $ou = $textBox_ou_path.Text
        
        if ($Global:multi_click -eq $true) {
            $continue = [System.Windows.MessageBox]::Show("Do you want to move $($Global:cont) usernames?", 'Confirmation', 'YesNo');
        }
        elseif ($Global:single_click -eq $true) {
            $continue = [System.Windows.MessageBox]::Show("Do you want to move the user '$($Global:iD)'?", 'Confirmation', 'YesNo');
        }
        

        if ($Global:multi_click -eq $true -and $continue -eq 'Yes') {
            Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Starting moving.." | Out-File $logFile -Append
            Add-Logs -text "Starting moving.."
            Foreach ($username in $usernames) {
                Try {
                    Get-ADUser -Identity $username.samaccountname -Server $domain | Move-ADObject -TargetPath "$ou" -ErrorAction Stop
                    Write-Host "Successfully moved user $($username.samaccountName) to $($ou)" -ForegroundColor Green
                    Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Successfully moved user $($username.samaccountName) to $($ou)" | Out-File $logFile -Append
                    Add-Logs -text "Successfully moved user $($username.samaccountName) to $($ou)"

                }
                catch {
                    Write-Host "Failed moving user $($username.samaccountName) to $($ou)" -ForegroundColor Red
                    Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Failed moving user $($username.samaccountName) to $($ou)" | Out-File $logFile -Append
                    Add-Logs -text "Failed moving user $($username.samaccountName) to $($ou)"

                }
            }
        }
        elseif ($Global:single_click -eq $true -and $continue -eq 'Yes') {
            Get-ADUser -Identity $iD -Server $domain | Move-ADObject -TargetPath "$ou" -ErrorAction Stop
            Write-Host "Successfully moved user $($iD) to $($ou)" -ForegroundColor Green
            Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Successfully moved user $($iD) to $($ou)" | Out-File $logFile -Append
            Add-Logs -text "Successfully moved user $($iD) to $($ou)"

        }
        else {
            Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Move been canceled by the user" | Out-File $logFile -Append
            Add-Logs -text "Move been canceled by the user"
        }
    })
$Window.Controls.Add($bt_move)

# Addding 2 buttons: 'Single-User' and 'Multi-Users'
$bt_single = New-Object System.Windows.Forms.Button
$bt_single.Location = New-Object System.Drawing.Point(70, 55)
$bt_single.AutoSize = $true
$bt_single.TextAlign = "MiddleCenter"
$bt_single.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold) 
$bt_single.Text = "Single-User"
$bt_single.Visible = $true
$bt_single.Add_Click( {
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Button Single-User as been clicked" | Out-File $logFile -Append
        Add-Logs -text "Single-User selected"


        $Global:single_click = $true
        $Global:multi_click = $false

        $lbl_domain.Select()
        $bt_single.Visible = $false
        $bt_multi.Visible = $false
        $bt_back.Visible = $true
        $bt_search.Visible = $true
        $lbl_user.Visible = $true
        $textBox_user.Visible = $true
        $lbl_user_details.Visible = $true
        $lbl_user_details_name.Visible = $true
        $lbl_user_details_name_info.Visible = $true
        $lbl_user_details_ou.Visible = $true
        $lbl_user_details_ou_info.Visible = $true
    })
$Window.Controls.Add($bt_single)

$bt_multi = New-Object System.Windows.Forms.Button
$bt_multi.Location = New-Object System.Drawing.Point(170, 55)
$bt_multi.AutoSize = $true
$bt_multi.TextAlign = "MiddleCenter"
$bt_multi.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold) 
$bt_multi.Text = "Multi-Users"
$bt_multi.Visible = $true
$bt_multi.Add_Click( {
        Write-Output "$(Get-Date -Format "MM/dd/yyyy HH:mm") - Button Multi-User as been clicked" | Out-File $logFile -Append
        Add-Logs -text "Multi-User selected"

        $Global:multi_click = $true
        $Global:single_click = $false

        $lbl_file.Select()
        $bt_single.Visible = $false
        $bt_multi.Visible = $false

        #Read file with Samaccountname, OU and Domain
        $ADUsersFile = ".\users_Multiusers.csv"
        $Global:usernames = Import-Csv -Path $ADUsersFile -Encoding UTF8

        # Get file name and inseted in a label
        $file = Get-ChildItem .\users_Multiusers.csv
        $lbl_fs.Text = $file.Name

        $lbl_users_loaded.Visible = $true
        $lbl_total.Visible = $true
        $bt_back.Visible = $true
        $lbl_file.Visible = $true
        $lbl_fs.Visible = $true
        $bt_edit.Visible = $true
        $bt_rf_users.Visible = $true


        $Global:cont = 0
        Foreach ($username in $usernames) {
            $Global:cont += 1
        }
        $lbl_total.Text = $Global:cont
    })
$Window.Controls.Add($bt_multi)

# Set $lbl_domain select when the script open
$lbl_domain.Select()


# Show Window
$Window.Add_Shown( { $Window.Activate() })
[void] $Window.ShowDialog()