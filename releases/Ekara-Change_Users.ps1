#####################################################################################################
#                           Example of use of the EKARA API                                         #
#####################################################################################################
# Swagger interface : https://api.ekara.ip-label.net                                                #
# To be personalized in the code before use: username / password                                    #
# Purpose of the script : Allows you to modify Ekara user settings (firstname / Lastname / Email)   #
#####################################################################################################
# Author : Guy Sacilotto
# Last Update : 15/02/2025
# Version : 3.0

<#
Authentication :  user / password
Method call : adm-api/users / adm-api/user   
Restitution: Graphical interface
#>

Clear-Host

#region VARIABLES
#========================== SETTING THE VARIABLES ===============================
$error.clear()
add-type -assemblyName "Microsoft.VisualBasic"
$global:API = "https://api.ekara.ip-label.net"                                                # Webservice URL
$global:UserName = ""                                                                         # EKARA Account
$global:PlainPassword = ""                                                                    # EKARA Password
$global:API_KEY = ""                                                                          # EKARA Key account

$global:Result_OK = 0
$global:Result_KO = 0
$global:SelectedUser  = 0
$global:headers = $null
$global:headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"             # Create Header
$headers.Add("Accept","application/json")                                                           # Setting Header
$headers.Add("Content-Type","application/json")                                                     # Setting Header

# Authentication choice
    # 1 = Without asking for an account and password (you must configure the account and password in this script.)
    # 2 = Request the entry of an account and a password (default)
    # 3 = With API-KEY
    $global:Auth = 2
#endregion

#region Functions
function Authentication{
    try{
        Switch($Auth){
            1{
                # Without asking for an account and password
                if(($null -ne $UserName -and $null -ne $PlainPassword) -and ($UserName -ne '' -and $PlainPassword -ne '')){
                    Write-Host "--- Automatic AUTHENTICATION (account) ---------------------------" -BackgroundColor Green
                    $uri = "$API/auth/login"                                                                                                    # Webservice Methode
                    $response = Invoke-RestMethod -Uri $uri -Method POST -Verbose -Body @{ email = "$UserName"; password = "$PlainPassword"}    # Call WebService method
                    $global:Token = $response.token                                                                                             # Register the TOKEN
                    $global:headers.Add("authorization","Bearer $Token")                                                                        # Adding the TOKEN into header
                }Else{
                    Write-Host "--- Account and Password not set ! ---------------------------" -BackgroundColor Red
                    Write-Host "--- To use this connection mode, you must configure the account and password in this script." -ForegroundColor Red
                    Break Script
                }
            }
            2{
                # Requests the entry of an account and a password (default) 
                Write-Host "------------------------------ AUTHENTICATION with account entry ---------------------------" -ForegroundColor Green
                $MyAccount = $Null
                $MyAccount = Get-credential -Message "EKARA login account" -ErrorAction Stop                                            # Request entry of the EKARA Account
                if(($null -ne $MyAccount) -and ($MyAccount.password.Length -gt 0)){
                    $UserName = $MyAccount.GetNetworkCredential().username
                    $PlainPassword = $MyAccount.GetNetworkCredential().Password
                    $uri = "$API/auth/login"
                    $response = Invoke-RestMethod -Uri $uri -Method POST -Body @{ email = "$UserName"; password = "$PlainPassword"} -Verbose     # Call WebService method
                    $Token = $response.token                                                                                            # Register the TOKEN
                    $global:headers.Add("authorization","Bearer $Token")
                }Else{
                    Write-Host "--- Account and password not specified ! ---------------------------" -BackgroundColor Red
                    Write-Host "--- To use this connection mode, you must enter Account and password." -ForegroundColor Red
                    Break Script
                }
            }
            3{
                # With API-KEY
                Write-Host "------------------------------ AUTHENTICATION With API-KEY ---------------------------" -ForegroundColor Green
                if(($null -ne $API_KEY) -and ($API_KEY -ne '')){
                    $global:headers.Add("X-API-KEY", $API_KEY)
                }Else{
                    Write-Host "--- API-KEY not specified ! ---------------------------" -BackgroundColor Red
                    Write-Host "--- To use this connection mode, you must configure API-KEY." -ForegroundColor Red
                    Break Script
                }
            }
        }
    }Catch{

    Write-Host "-------------------------------------------------------------" -ForegroundColor red 
        Write-Host "Erreur ...." -BackgroundColor Red
        Write-Host $Error.exception.Message[0]
        Write-Host $Error[0]
        Write-host $error[0].ScriptStackTrace
        Write-Host "-------------------------------------------------------------" -ForegroundColor red
        Break Script
    }
}

function CountCheckBox { 
    $Global:SelectedUser = 0
    $dataGridView.EndEdit()
    #[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    for($i=0;$i -lt $DataGridView.RowCount;$i++){ 
        if($DataGridView.Rows[$i].Cells['exp'].Value -eq $true){
            $Global:SelectedUser++                                                                          # Count selected CheckBox
            #$DataGridView.DataBindings.DefaultDataSourceUpdateMode = 1
        }
    }
    
    if($Global:SelectedUser -gt 0){
        $label4.ForeColor = "green"
        $label4.Text = ("[" + $Global:SelectedUser + "] users selected out of ["+$global:users.count+"]")   # Update Label
        $button_Update.Enabled = $True
    }else{
        $label4.ForeColor = "black"
        $label4.Text = ("Select the users, then modify the values.")                                        # Update Label
        $button_Update.Enabled = $False
    }
    
    Write-Host("Selected Values " + $Global:SelectedUser);
    $label2.Text = ("Users selected : " + $Global:SelectedUser)                                             # Update value
}

Function List_Users_ID{
    try{
        #========================== adm-api/users ================================
        Write-Host "-------------------------------------------------------------" -ForegroundColor green
        Write-Host "------------------- Liste tous les utilisateurs du compte -------------------" -BackgroundColor "White" -ForegroundColor "DarkCyan"
        $uri ="$API/adm-api/users"                                                                                      # Format API request
        $global:users = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Verbose                             # Call WebService method

        # --------------- Populate the table ------------------------------
        foreach($user in $users){
            $workspaces = "[]"
            If(($user.workspaces).count -gt 0){
                Write-Host ("Number of Workspace ["+($user.workspaces).count+"]")
                $workspaces = ""
                $Counter = 0
                $Total = ($user.workspaces).count

                # Format JSON body response for Workspace
                foreach($workspace in $user.workspaces){
                    Write-Host $workspace.id $workspace.name
                    $WorkspaceId = $workspace.id
                    $WorkspaceName = $workspace.name
                    $Counter++
                    if($Counter -ne $Total){
                        $workspaces = $workspaces + (
                            "{
                                ""id"":""$WorkspaceId"",
                                ""name"":""$WorkspaceName""
                            },")
                    }else{
                        $workspaces = $workspaces + (
                            "{
                                ""id"":""$WorkspaceId"",
                                ""name"":""$WorkspaceName""
                            }")
                    }
                }
                $workspaces = "["+$workspaces+"]"
            }

            [void]$datagridview.Rows.Add($null,$user.ID, $user.firstname, $user.lastname, $user.Email, $user.roleId, $user.roleName, $user.timezone.timeZoneLabel, (($user.emailing.features).tostring()).ToLower(), (($user.emailing.technicals).tostring()).ToLower(), (($user.emailing.communications).ToString()).ToLower(), $workspaces)
            $label4.ForeColor = "black"
            $label4.Text = ("Select the users, then modify the values.")                # Update Label
            $label2.Text = ("Users selected : " + 0)                                    # Update value
        }
        
        $DataGridView.AutoSizeColumnsMode = 'AllCells'                                  # Automatically resize columns

    }catch{
        Write-Host "-------------------------------------------------------------" -ForegroundColor red
        Write-Host "Erreur ...." -BackgroundColor Red
        Write-Host $Error.exception.Message[0]
        Write-Host $Error[0]
        Write-host $error[0].ScriptStackTrace
        Write-Host "-------------------------------------------------------------" -ForegroundColor red
    }    
}

Function Update_users{
    try{
        if ($Global:SelectedUser -gt 0){
            $count = 0
            $label4.ForeColor = "green"
            $label4.Text = ("--> Update started...")                                            # Update Label
            for($i=0;$i -lt $DataGridView.RowCount;$i++){ 
                # Check if CheckBox is checked
                if($DataGridView.Rows[$i].Cells['exp'].Value -eq $true){
                    $count++
                    Write-Host ("------------------- Run Update User " + $count + " of " + $Global:SelectedUser + " - " +$DataGridView.Rows[$i].Cells['Firstname'].Value + " " + $DataGridView.Rows[$i].Cells['Lastname'].Value + " -------------------") -BackgroundColor "White" -ForegroundColor "DarkCyan"
                                            
                    $UserID = $DataGridView.Rows[$i].Cells['ID'].Value
                    $UserEmail = $DataGridView.Rows[$i].Cells['Email'].Value
                    $UserFirstname = $DataGridView.Rows[$i].Cells['firstname'].Value
                    $UserLastName = $DataGridView.Rows[$i].Cells['lastname'].Value
                    $UserRoleId = $DataGridView.Rows[$i].Cells['roleId'].Value
                    $UserRoleName = $DataGridView.Rows[$i].Cells['roleName'].Value
                    $UserTimezone = $DataGridView.Rows[$i].Cells['timezone'].Value
                    $UserEmailing_features = $DataGridView.Rows[$i].Cells['features'].Value
                    $UserEmailing_technicals = $DataGridView.Rows[$i].Cells['technicals'].Value
                    $UserEmailing_communications = $DataGridView.Rows[$i].Cells['communications'].Value
                    $UserWorkspace = $DataGridView.Rows[$i].Cells['Workspaces'].Value
                    
                    # Format JSON body response
                    $body=@"
                    {
                        "email": "$UserEmail",
                        "firstname":"$UserFirstname",
                        "lastname":"$UserLastName",
                        "roleId":"$UserRoleId",
                        "roleName":"$UserRoleName",
                        "timezone":"$UserTimezone",
                        "emailing":{
                            "features":$UserEmailing_features,
                            "technicals":$UserEmailing_technicals,
                            "communications":$UserEmailing_communications
                        },
                        "workspaces":$UserWorkspace
                    }
"@
                    Try{
                        $uri ="$API/adm-api/user/$UserID"           # Format API request
                        $requestWS = Invoke-RestMethod -Uri $uri -Method PATCH -Headers $headers -body $body -ContentType 'application/json' -Verbose
                        Write-Host ("--> User " + $count + " of " + $Global:SelectedUser + " ["+$UserEmail+"] ["+$UserID+"] was updated") -ForegroundColor "Green"
                        $label4.ForeColor = "green"
                        $label4.Text = ("--> User " + $count + " of " + $Global:SelectedUser + " ["+$UserEmail+"] ["+$UserID+"] was updated")       # Update Label
                        $global:Result_OK++ 
                        Start-Sleep -Seconds 1
                    }catch{
                        Write-Host ("Failed update user "+ $count + " of " + $Global:SelectedUser + " ["+$UserEmail+"] ["+$UserID+"]")
                        $label4.ForeColor = "red"
                        $label4.Text = ("Failed update user "+ $count + " of " + $Global:SelectedUser + " ["+$UserEmail+"] ["+$UserID+"]")  
                        $global:Result_KO++
                        Write-Host "-------------------------------------------------------------" -ForegroundColor red 
                        Write-Host "Erreur ...." -BackgroundColor Red
                        Write-Host $Error.exception.Message[0]
                        Write-Host $Error[0]
                        Write-host $error[0].ScriptStackTrace
                        Write-Host "-------------------------------------------------------------" -ForegroundColor red
                    }
                }
            }
            Start-Sleep -Seconds 1
            $label4.ForeColor = "green"
            $label4.Text = ("--> ["+$global:Result_OK +"] users successfully updated / [$global:Result_KO] Failure update.")                         # Update Label
            Start-Sleep -Seconds 1
            $label4.ForeColor = "green"
            $label4.Text = ("Update ended")                                                     # Update Label
            Start-Sleep -Seconds 1
            $datagridview.Rows.Clear()
            List_Users_ID
            $button_Update.Enabled = $False

        }       
    }catch{
        Write-Host -message "-------------------------------------------------------------" -ForegroundColor Red
        Write-Host -message "Erreur " -BackgroundColor "Red"
        Write-Host -message $Error.exception.Message[0]
        Write-Host -message $Error[0]
        Write-Host -message $error[0].ScriptStackTrace
        Write-Host -message "-------------------------------------------------------------" -ForegroundColor red
        #exit
    }  
}
#endregion

Authentication

#region UI
# --------------- Create Form -------------------------------
$Form = New-Object System.Windows.Forms.Form
$Form.Name = "Update users"                                                     # Form name
$Form.Text = 'List of all users of current account'                             # Title
$Form.Size = New-Object System.Drawing.Size(1040,700)                           # Width  / height
$Form.StartPosition = 'CenterScreen'                                            # Set position on the screen
$Form.FormBorderStyle = 'Fixed3D'                                               # Disable Resize and Maximize
$Form.Opacity = 1.0
$Form.TopMost = $false
$Form.ShowIcon = $true                                                          # Enable icon (upper left corner) $ true, disable icon
        
# --------------- Button OK -------------------------------------
$button_Update = New-Object System.Windows.Forms.Button
$button_Update.Location = New-Object System.Drawing.Point(350,600)              # From the left / From the top
$button_Update.Size = New-Object System.Drawing.Size(75,23)                     # Width  / height
$button_Update.Text = 'Update'
$button_Update.AutoSize = $true
$button_Update.Enabled = $false                                                 # Desabled button
$button_Update.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom 
$button_Update.Add_Click($function:Update_users)                                # Execute function after click on button
       
# --------------- Button_quitter ----------------------------------
$button_quitter = New-Object System.Windows.Forms.Button
$button_quitter.Location = New-Object System.Drawing.Point(450,600)             # From the left / From the top
$button_quitter.Name = 'button_quitter'                                         # Button name
$button_quitter.Size = New-Object System.Drawing.Size(75,23)                    # Width  / height
$button_quitter.Text = 'Exit'                                                   # Text displayed
$button_quitter.AutoSize = $true                                                # Auto size button
$button_quitter.UseVisualStyleBackColor = $true 
$button_quitter.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom 
$button_quitter.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$button_quitter.Add_Click({$Form.close()})                                      # Execute function after click on button

# --------------- textarea ---------------------------------------
$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(10,20)                       # From the left / From the top
$label1.Size = New-Object System.Drawing.Size(280,20)                           # Width  / height
$label1.Text = 'Select the users to update :'                                   # Text displayed
$label1.AutoSize = $true                                                        # Auto size label
$label1.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
    -bor [System.Windows.Forms.AnchorStyles]::Bottom `
    -bor [System.Windows.Forms.AnchorStyles]::Left `
    -bor [System.Windows.Forms.AnchorStyles]::Right

# --------------- textarea ---------------------------------------
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(200,630)                      # From the left / From the top
$label2.Size = New-Object System.Drawing.Size(20,20)                            # Width  / height
$label2.Text = ("Users selected : " + $Global:SelectedUser)                     # Text displayed
$label2.AutoSize = $true                                                        # Auto size label
$label2.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom `
    -bor [System.Windows.Forms.AnchorStyles]::Left 

# --------------- textarea ---------------------------------------
$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(10,630)                     # From the left / From the top
$label3.Size = New-Object System.Drawing.Size(20,20)                            # Width  / height
$label3.Text = "Total users : " + $users.Count                                  # Text displayed
$label3.AutoSize = $true                                                        # Auto size label
$label3.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom `
    -bor [System.Windows.Forms.AnchorStyles]::Left 

# --------------- textarea ---------------------------------------
$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(200,560)                     # From the left / From the top
$label4.Size = New-Object System.Drawing.Size(600,20)                           # Width  / height
$label4.BackColor  = "white"
$label4.ForeColor = "black"
$label4.Text = " message affiché en cas d'erreur "
$label4.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom `
    -bor [System.Windows.Forms.AnchorStyles]::Right 

# --------------- DataGrid ---------------------------------------
$DataGridView = New-Object System.Windows.Forms.DataGridView
$DataGridView.Location = New-Object System.Drawing.Point(10,40)                # From the left / From the top
$DataGridView.Size = New-Object System.Drawing.Size(1000,500)                  # Width  / height
$DataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
    -bor [System.Windows.Forms.AnchorStyles]::Bottom `
    -bor [System.Windows.Forms.AnchorStyles]::Left `
    -bor [System.Windows.Forms.AnchorStyles]::Right
$DataGridView.ColumnCount = 11
$DataGridView.ColumnHeadersVisible = $true
$DataGridView.Columns.Insert(0,(New-Object System.Windows.Forms.DataGridViewCheckBoxColumn))
$DataGridView.Columns[0].Name = "Exp"                                           # header name
$DataGridView.Columns[0].Width = '30'                                           # Checkbox Size
$DataGridView.Columns[1].Name = "ID"                                            # header name
$DataGridView.Columns[2].Name = "firstname"                                     # header name
$DataGridView.Columns[3].Name = "lastname"                                      # header name
$DataGridView.Columns[4].Name = "Email"                                         # header name
$DataGridView.Columns[5].Name = "roleId"                                        # header name
$DataGridView.Columns[6].Name = "roleName"                                      # header name
$DataGridView.Columns[7].Name = "timezone"                                      # header name
$DataGridView.Columns[8].Name = "features"                                      # header name
$DataGridView.Columns[9].Name = "technicals"                                    # header name
$DataGridView.Columns[10].Name = "communications"                               # header name
$DataGridView.Columns[11].Name = "Workspaces"                                   # header name

$DataGridView.DefaultCellStyle.WrapMode = 'True'                     

$DataGridView.Columns["ID"].ReadOnly = $true                                    # ReadOnly value in column
$DataGridView.Columns["roleId"].ReadOnly = $true                                # ReadOnly value in column
$DataGridView.Columns["roleName"].ReadOnly = $true                              # ReadOnly value in column
$DataGridView.Columns["timezone"].ReadOnly = $true                              # ReadOnly value in column
$DataGridView.Columns["features"].ReadOnly = $true                              # ReadOnly value in column
$DataGridView.Columns["technicals"].ReadOnly = $true                            # ReadOnly value in column
$DataGridView.Columns["communications"].ReadOnly = $true                        # ReadOnly value in column

$DataGridView.Columns["ID"].DefaultCellStyle.BackColor = "Gray"                 # Background Color in column
$DataGridView.Columns["roleId"].DefaultCellStyle.BackColor = "Gray"             # Background Color in column
$DataGridView.Columns["roleName"].DefaultCellStyle.BackColor = "Gray"           # Background Color in column
$DataGridView.Columns["timezone"].DefaultCellStyle.BackColor = "Gray"           # Background Color in column
$DataGridView.Columns["features"].DefaultCellStyle.BackColor = "Gray"           # Background Color in column
$DataGridView.Columns["technicals"].DefaultCellStyle.BackColor = "Gray"         # Background Color in column
$DataGridView.Columns["communications"].DefaultCellStyle.BackColor = "Gray"     # Background Color in column
$DataGridView.Columns["Workspaces"].DefaultCellStyle.BackColor = "Gray"         # Background Color in column
        
$DataGridView.Columns["Workspaces"].Visible = $False                            # Hide column
$DataGridView.AllowUserToAddRows = $false                                       # Disable row add 
$DataGridView.AllowUserToDeleteRows = $false                                    # Disable row deletion
$dataGridView.MultiSelect   = $true                                             # Enable multiple selection
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect

# --------------- Count selected item -----------------------------
$DataGridView.Add_CellValueChanged($function:CountCheckBox)                     # Update Form when a checkbox is checked
$DataGridView.Sort($DataGridView.Columns["lastname"],'Ascending')               # Ascending / Descending
$DataGridView.RowHeadersVisible = $false                                        # Hide row headers
$DataGridView.AutoSize = $False
$DataGridView.ScrollBars = "Both"


# --------------- Add the components to the form ------------------
$Form.Controls.Add($button_Update)                                              # Add object into form
$Form.Controls.Add($button_quitter)                                             # Add object into form
$Form.Controls.Add($DataGridView)                                               # Add object into form
$Form.Controls.Add($label1)                                                     # Add object into form
$Form.Controls.Add($label2)                                                     # Add object into form
$Form.Controls.Add($label3)                                                     # Add object into form
$Form.Controls.Add($label4)                                                     # Add object into form
$Form.Topmost = $true                                                           # Form Focus
$Form.Add_Load({ List_Users_ID })                                               # Execute function after loading form
$Form = $form.ShowDialog()                                                      # Display form
#endregion
