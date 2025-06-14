Add-Type -AssemblyName System.Core
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data # Required for DataGridView

#region Helper Function: Get-SessionParams
function Get-SessionParams {
    param (
        [string]$ComputerName
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Enter Credentials for $ComputerName"
    $form.Size = New-Object System.Drawing.Size(350, 200)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    
    $lblUsername = New-Object System.Windows.Forms.Label
    $lblUsername.Text = "Username:"
    $lblUsername.Location = New-Object System.Drawing.Point(20, 20)
    $lblUsername.AutoSize = $true
    $form.Controls.Add($lblUsername)
    
    $txtUsername = New-Object System.Windows.Forms.TextBox
    $txtUsername.Location = New-Object System.Drawing.Point(120, 20)
    $txtUsername.Size = New-Object System.Drawing.Size(180, 20)
    $txtUsername.Text = [Environment]::UserName # Pre-fill with current user
    $form.Controls.Add($txtUsername)
    
    $lblPassword = New-Object System.Windows.Forms.Label
    $lblPassword.Text = "Password:"
    $lblPassword.Location = New-Object System.Drawing.Point(20, 50)
    $lblPassword.AutoSize = $true
    $form.Controls.Add($lblPassword)
    
    $txtPassword = New-Object System.Windows.Forms.TextBox
    $txtPassword.Location = New-Object System.Drawing.Point(120, 50)
    $txtPassword.Size = New-Object System.Drawing.Size(180, 20)
    $txtPassword.UseSystemPasswordChar = $true
    $form.Controls.Add($txtPassword)
    
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(120, 90)
    $btnOK.Size = New-Object System.Drawing.Size(75, 23)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOK)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(225, 90)
    $btnCancel.Size = New-Object System.Drawing.Size(75, 23)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)

    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel
    
    $sessionParams = @{
        ComputerName = $ComputerName
    }

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $username = $txtUsername.Text
        $password = $txtPassword.Text | ConvertTo-SecureString -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($username, $password)
        $sessionParams.Add('Credential', $credential)
    } else {
        throw "Credential entry cancelled for $ComputerName."
    }

    return $sessionParams
}
#endregion

#region Helper Function: LoadTasksIntoDataGridView
function LoadTasksIntoDataGridView {
    param (
        [string]$ComputerName,
        [System.Windows.Forms.TextBox]$LogTextBox, # This is the main log textbox
        [System.Windows.Forms.DataGridView]$DataGridView
    )

    $LogTextBox.Clear()
    $DataGridView.DataSource = $null 
    $DataGridView.Columns.Clear() 
    $DataGridView.Rows.Clear()    
    
    if ([string]::IsNullOrWhiteSpace($ComputerName)) {
        $LogTextBox.AppendText("Error: Computer Name cannot be empty.`r`n")
        return
    }

    $LogTextBox.AppendText("Attempting to connect to $ComputerName and load tasks...`r`n")
    $session = $null 
    try {
        $LogTextBox.AppendText("Requesting credentials for $ComputerName...`r`n")
        $sessionParams = Get-SessionParams $ComputerName 

        $LogTextBox.AppendText("Credentials received. Attempting New-PSSession...`r`n")
        $Error.Clear() 
        $session = New-PSSession @sessionParams -ErrorAction Stop

        if ($null -eq $session) {
            $LogTextBox.AppendText("FATAL ERROR: New-PSSession returned null without throwing an exception. This is highly unusual.`r`n")
            if ($Error.Count -gt 0) {
                $LogTextBox.AppendText("Last PowerShell error: $($Error[0].Exception.Message)`r`n")
                $LogTextBox.AppendText("Error category: $($Error[0].CategoryInfo.Reason)`r`n")
            }
            throw "Failed to establish PowerShell session (session object is null)."
        }
        $LogTextBox.AppendText("Successfully established PSSession to $ComputerName. Loading scheduled tasks...`r`n")

        $tasks = Invoke-Command -Session $session -ScriptBlock {
            try {
                Get-ScheduledTask | Select-Object TaskName, State, LastRunTime, LastTaskResult, TaskPath
            } catch {
                Write-Error "Error retrieving scheduled tasks on $env:COMPUTERNAME: $($_.Exception.Message)"
                throw $_ 
            }
        } -ErrorAction Stop

        if ($null -eq $tasks -or -not ($tasks -is [System.Collections.IEnumerable])) {
            $LogTextBox.AppendText("Warning: Get-ScheduledTask returned no tasks or a non-enumerable result. This might be normal if no tasks exist.`r`n")
            $tasks = @() 
        }
        
        $dt = New-Object System.Data.DataTable # This $dt is the target for the DataGridView
        
        $dt.Columns.Add("Name") 
        $dt.Columns.Add("State") 
        $dt.Columns.Add("Last Run Time")
        $dt.Columns.Add("Last Result")
        $dt.Columns.Add("Path")
        
        if ($null -eq $dt -or $dt.Columns.Count -eq 0) {
            $LogTextBox.AppendText("FATAL ERROR: DataTable object is null or has no columns after creation. Cannot bind to DataGridView.`r`n")
            throw "DataTable initialization failed."
        }
    
        $LogTextBox.AppendText("DataTable initialized with $($dt.Columns.Count) columns.`r`n")


        foreach ($task in $tasks) {
            if ($null -eq $task) {
                $LogTextBox.AppendText("WARNING: Found a null task object in the collection. Skipping this entry.`r`n")
                continue 
            }

            try { 
                $row = $dt.NewRow()
                
                $row["Name"] = if ($null -ne $task.TaskName) { $task.TaskName } else { "(Unknown Name)" }
                $row["State"] = if ($null -ne $task.State) { $task.State } else { "(Unknown State)" }
                
                if ($null -ne $task.LastRunTime -and $task.LastRunTime -ne [DateTime]::MinValue) {
                    $row["Last Run Time"] = $task.LastRunTime.ToString("yyyy-MM-dd HH:mm:ss")
                } else {
                    $row["Last Run Time"] = "Never"
                }
                
                $row["Last Result"] = if ($null -ne $task.LastTaskResult) { $task.TaskResult } else { "(Unknown Result)" }
                $row["Path"] = if ($null -ne $task.TaskPath) { $task.TaskPath } else { "(Unknown Path)" }

                $dt.Rows.Add($row)
            } catch {
                $LogTextBox.AppendText("ERROR: Failed to process task '$($task.TaskName)' (or unknown task): $($_.Exception.Message)`r`n")
                $LogTextBox.AppendText("Skipping this task.`r`n")
            }
        }

        $LogTextBox.AppendText("Attempting to set DataGridView.DataSource with $($dt.Rows.Count) rows.`r`n")
        
        if ($null -eq $DataGridView) {
            $LogTextBox.AppendText("FATAL ERROR: \$DataGridView is null just before DataSource assignment! This should not happen.`r`n")
            throw "DataGridView object is null."
        }

        $DataGridView.AutoGenerateColumns = $true 
        
        if (-not $DataGridView.IsHandleCreated) {
            $DataGridView.CreateControl()
            $LogTextBox.AppendText("DataGridView handle forced to create.`r`n")
        }
        
        $DataGridView.SuspendLayout()
        try {
            $LogTextBox.AppendText("Dispatching DataSource assignment to UI thread queue...`r`n")

            # --- CRITICAL FIX IS HERE ---
            # The script block now accepts both $grid and $logBox as explicit parameters.
            $scriptBlockToInvoke = {
                param(
                    [System.Windows.Forms.DataGridView]$grid,
                    [System.Windows.Forms.TextBox]$logBoxParam # The log textbox passed as a parameter
                )

                # Access the DataTable directly from the script's parent scope.
                # This still bypasses the cross-thread parameter marshalling issue for $dt.
                $localDt = $script:dt 
                
                try {
                    if ($null -eq $localDt) {
                        # Now using the passed logBoxParam
                        $logBoxParam.AppendText("FATAL ERROR: DataTable (\$dt) is null when accessed in BeginInvoke script block.`r`n")
                        throw "DataTable is null in UI update."
                    }

                    $grid.DataSource = $localDt 

                    # Now using the passed logBoxParam for all logging inside this block
                    $logBoxParam.AppendText("DataSource set successfully via BeginInvoke.`r`n") 
                    
                    [System.Windows.Forms.Application]::DoEvents() 

                    if ($grid.Rows.Count -gt 0) { 
                        try {
                            $grid.Rows[0].Selected = $true
                            $grid.CurrentCell = $grid.Rows[0].Cells[0]
                            $logBoxParam.AppendText("First row selected and current cell set.`r`n")
                        } catch {
                            $logBoxParam.AppendText("WARNING (BeginInvoke): Could not select first row or set current cell after binding: $($_.Exception.Message)`r`n")
                            $logBoxParam.AppendText("Inner details: $($_.Exception.GetType().FullName)`r`n")
                        }
                    } else {
                        $logBoxParam.AppendText("No rows displayed in DataGridView despite DataTable having $($localDt.Rows.Count) rows (via BeginInvoke).`r`n")
                    }
                } catch {
                    # This is the line that was failing before (line 227 in original error if not changed)
                    $logBoxParam.AppendText("ERROR (BeginInvoke): Failed to set DataSource on UI thread: $($_.Exception.Message)`r`n")
                    $logBoxParam.AppendText("Details: $($_.Exception.GetType().FullName)`r`n")
                }
            }
            
            # Define the Action delegate type, now for TWO parameters (DataGridView, TextBox)
            # Ensure Add-Type -AssemblyName System.Core is at the top of your script.
            $ActionType = [System.Action[System.Windows.Forms.DataGridView, System.Windows.Forms.TextBox]] 

            # Explicitly cast the script block to the specific Action delegate type.
            $DelegateToInvoke = [System.Action[System.Windows.Forms.DataGridView, System.Windows.Forms.TextBox]]$scriptBlockToInvoke 

            # Call BeginInvoke, passing both $DataGridView and $LogTextBox as arguments (in an array).
            $DataGridView.BeginInvoke($DelegateToInvoke, @($DataGridView, $LogTextBox)) 
            # --- END CRITICAL FIX ---

        } finally {
            $DataGridView.ResumeLayout()
            $LogTextBox.AppendText("DataGridView layout resumed.`r`n")
        }

        $LogTextBox.AppendText("Successfully loaded $($dt.Rows.Count) tasks from $ComputerName. (UI update dispatched).`r`n")

    } catch {
        $LogTextBox.AppendText("Error connecting or loading tasks from ${ComputerName}: $($_.Exception.Message)`r`n")
        $LogTextBox.AppendText("Details: $($_.Exception.GetType().FullName)`r`n")
        $LogTextBox.AppendText("Script Line: $($_.InvocationInfo.ScriptLineNumber)`r`n")
        $LogTextBox.AppendText("Stack Trace:`r`n$($_.ScriptStackTrace)`r`n")
    } finally {
        if ($session -is [System.Management.Automation.Runspaces.PSSession]) {
            $LogTextBox.AppendText("Cleaning up PSSession...`r`n")
            Remove-PSSession $session -ErrorAction SilentlyContinue
            $LogTextBox.AppendText("PSSession removed.`r`n")
        } elseif ($null -ne $session) {
            $LogTextBox.AppendText("Warning: Session object existed but was not a PSSession type, skipping removal.`r`n")
        } else {
            $LogTextBox.AppendText("No PSSession to clean up.`r`n")
        }
    }
}
#endregion

#region Helper Function: PerformTaskAction
function PerformTaskAction {
    param (
        [string]$ActionName,
        [string]$ComputerName,
        [string]$TaskName,
        [System.Windows.Forms.TextBox]$LogTextBox,
        [System.Windows.Forms.DataGridView]$DataGridView,
        [scriptblock]$ActionScriptBlock,
        [bool]$Confirm = $false
    )

    if ([string]::IsNullOrWhiteSpace($TaskName)) {
        $LogTextBox.AppendText("Error: No task selected for $ActionName action.`r`n")
        return
    }

    if ($Confirm) {
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to $ActionName task '$TaskName' on '$ComputerName'?",
            "Confirm Action",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
            $LogTextBox.AppendText("$ActionName action cancelled for task '$TaskName'.`r`n")
            return
        }
    }

    $LogTextBox.AppendText("Attempting to $ActionName task '$TaskName' on '$ComputerName'...`r`n")
    $session = $null
    try {
        $sessionParams = Get-SessionParams $ComputerName
        $session = New-PSSession @sessionParams -ErrorAction Stop

        # Pass TaskName to the remote script block for the specific action
        Invoke-Command -Session $session -ScriptBlock $ActionScriptBlock -ArgumentList $TaskName -ErrorAction Stop

        $LogTextBox.AppendText("Successfully $ActionName task '$TaskName' on '$ComputerName'.`r`n")
        # Refresh the task list after successful action
        LoadTasksIntoDataGridView `
            -ComputerName $ComputerName `
            -LogTextBox $LogTextBox `
            -DataGridView $DataGridView

    } catch {
        $LogTextBox.AppendText("Error $ActionName task '$TaskName' on '$ComputerName': $($_.Exception.Message)`r`n")
    } finally {
        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
    }
}
#endregion

#region GUI Elements Setup
$frmMain = New-Object System.Windows.Forms.Form
$frmMain.Text = "Remote Scheduled Task Manager"
$frmMain.Size = New-Object System.Drawing.Size(650, 650) # Increased size for tabs
$frmMain.StartPosition = "CenterScreen"
$frmMain.FormBorderStyle = "FixedSingle"
$frmMain.MaximizeBox = $false
$frmMain.MinimizeBox = $false

# Title Label (still on main form)
$lblMainTitle = New-Object System.Windows.Forms.Label
$lblMainTitle.Text = "Remote Scheduled Task Manager"
$lblMainTitle.Location = New-Object System.Drawing.Point(10, 10)
$lblMainTitle.AutoSize = $true
$lblMainTitle.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
$frmMain.Controls.Add($lblMainTitle)

# TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 40)
$tabControl.Size = New-Object System.Drawing.Size(615, 560)
$frmMain.Controls.Add($tabControl)

#region Tab: Deploy Task
$tabDeploy = New-Object System.Windows.Forms.TabPage("Deploy Task")
$tabControl.Controls.Add($tabDeploy)

# All existing deploy task controls are now added to $tabDeploy
# Computer List
$lblComputers = New-Object System.Windows.Forms.Label
$lblComputers.Text = "Computer List File:"
$lblComputers.Location = New-Object System.Drawing.Point(10, 20)
$lblComputers.AutoSize = $true
$tabDeploy.Controls.Add($lblComputers)

$txtComputers = New-Object System.Windows.Forms.TextBox
$txtComputers.Location = New-Object System.Drawing.Point(140, 20)
$txtComputers.Size = New-Object System.Drawing.Size(300, 20)
$tabDeploy.Controls.Add($txtComputers)

$btnBrowseComputers = New-Object System.Windows.Forms.Button
$btnBrowseComputers.Text = "Browse..."
$btnBrowseComputers.Location = New-Object System.Drawing.Point(450, 18)
$btnBrowseComputers.Size = New-Object System.Drawing.Size(75, 23)
$btnBrowseComputers.Add_Click({
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "Text Files (*.txt)|*.txt|All files (*.*)|*.*"
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtComputers.Text = $OpenFileDialog.FileName
    }
})
$tabDeploy.Controls.Add($btnBrowseComputers)

# Script/Executable Selection
$lblScript = New-Object System.Windows.Forms.Label
$lblScript.Text = "Select Script/Executable:"
$lblScript.Location = New-Object System.Drawing.Point(10, 85)
$lblScript.AutoSize = $true
$tabDeploy.Controls.Add($lblScript)

$txtScript = New-Object System.Windows.Forms.TextBox
$txtScript.Location = New-Object System.Drawing.Point(140, 85)
$txtScript.Size = New-Object System.Drawing.Size(300, 20)
$tabDeploy.Controls.Add($txtScript)

$btnBrowseScript = New-Object System.Windows.Forms.Button
$btnBrowseScript.Text = "Browse..."
$btnBrowseScript.Location = New-Object System.Drawing.Point(450, 83)
$btnBrowseScript.Size = New-Object System.Drawing.Size(75, 23)
$btnBrowseScript.Add_Click({
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "Scripts and Executables (*.ps1;*.bat;*.cmd;*.exe)|*.ps1;*.bat;*.cmd;*.exe|All files (*.*)|*.*"
    $OpenFileDialog.Title = "Select Script or Executable"
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtScript.Text = $OpenFileDialog.FileName
    }
})
$tabDeploy.Controls.Add($btnBrowseScript)

# Task Name
$lblTaskName = New-Object System.Windows.Forms.Label
$lblTaskName.Text = "Scheduled Task Name:"
$lblTaskName.Location = New-Object System.Drawing.Point(10, 120)
$lblTaskName.AutoSize = $true
$tabDeploy.Controls.Add($lblTaskName)

$txtTaskName = New-Object System.Windows.Forms.TextBox
$txtTaskName.Location = New-Object System.Drawing.Point(140, 120)
$txtTaskName.Size = New-Object System.Drawing.Size(300, 20)
$txtTaskName.Text = "MyAutomatedTask" # Default value
$tabDeploy.Controls.Add($txtTaskName)

# Schedule Time
$lblTime = New-Object System.Windows.Forms.Label
$lblTime.Text = "Daily Run Time (HH:mm):"
$lblTime.Location = New-Object System.Drawing.Point(10, 155)
$lblTime.AutoSize = $true
$tabDeploy.Controls.Add($lblTime)

$txtTime = New-Object System.Windows.Forms.TextBox
$txtTime.Location = New-Object System.Drawing.Point(140, 155)
$txtTime.Size = New-Object System.Drawing.Size(80, 20)
$txtTime.Text = "09:00" # Default time
$tabDeploy.Controls.Add($txtTime)

# Option to remove old task
$chkRemoveOld = New-Object System.Windows.Forms.CheckBox
$chkRemoveOld.Text = "Remove old task if exists"
$chkRemoveOld.Location = New-Object System.Drawing.Point(10, 190)
$chkRemoveOld.AutoSize = $true
$chkRemoveOld.Checked = $true # Default to checked
$tabDeploy.Controls.Add($chkRemoveOld)

# Deploy Button
$btnDeploy = New-Object System.Windows.Forms.Button
$btnDeploy.Text = "Deploy Task"
$btnDeploy.Location = New-Object System.Drawing.Point(200, 230)
$btnDeploy.Size = New-Object System.Drawing.Size(120, 35)
$btnDeploy.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$tabDeploy.Controls.Add($btnDeploy)

# Deployment Log Textbox
$txtDeployLog = New-Object System.Windows.Forms.TextBox
$txtDeployLog.Location = New-Object System.Drawing.Point(10, 280)
$txtDeployLog.Size = New-Object System.Drawing.Size(585, 230)
$txtDeployLog.Multiline = $true
$txtDeployLog.ScrollBars = "Vertical"
$txtDeployLog.ReadOnly = $true
$tabDeploy.Controls.Add($txtDeployLog)
#endregion

#region Tab: View & Manage Tasks
$tabManage = New-Object System.Windows.Forms.TabPage("View & Manage Tasks")
$tabControl.Controls.Add($tabManage)

# Computer Name for Management
$lblManageComputer = New-Object System.Windows.Forms.Label
$lblManageComputer.Text = "Computer Name:"
$lblManageComputer.Location = New-Object System.Drawing.Point(10, 20)
$lblManageComputer.AutoSize = $true
$tabManage.Controls.Add($lblManageComputer)

$txtManageComputer = New-Object System.Windows.Forms.TextBox
$txtManageComputer.Location = New-Object System.Drawing.Point(120, 20)
$txtManageComputer.Size = New-Object System.Drawing.Size(200, 20)
$txtManageComputer.Text = $env:COMPUTERNAME # Default to local machine
$tabManage.Controls.Add($txtManageComputer)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect & Load Tasks"
$btnConnect.Location = New-Object System.Drawing.Point(330, 18)
$btnConnect.Size = New-Object System.Drawing.Size(150, 23)
$tabManage.Controls.Add($btnConnect)

# DataGridView to display tasks
$dgvTasks = New-Object System.Windows.Forms.DataGridView
$dgvTasks.Location = New-Object System.Drawing.Point(10, 60)
$dgvTasks.Size = New-Object System.Drawing.Size(585, 350)
$dgvTasks.AllowUserToAddRows = $false
$dgvTasks.ReadOnly = $true
$dgvTasks.AutoSizeColumnsMode = "Fill"
$dgvTasks.SelectionMode = "FullRowSelect"
$dgvTasks.MultiSelect = $false # Start with single select for simpler management actions
$tabManage.Controls.Add($dgvTasks)

# Management Log
$txtManageLog = New-Object System.Windows.Forms.TextBox
$txtManageLog.Location = New-Object System.Drawing.Point(10, 420)
$txtManageLog.Size = New-Object System.Drawing.Size(585, 90)
$txtManageLog.Multiline = $true
$txtManageLog.ScrollBars = "Vertical"
$txtManageLog.ReadOnly = $true
$tabManage.Controls.Add($txtManageLog)

# Management Buttons
$btnRunTask = New-Object System.Windows.Forms.Button
$btnRunTask.Text = "Run Task"
$btnRunTask.Location = New-Object System.Drawing.Point(10, 520)
$btnRunTask.Size = New-Object System.Drawing.Size(90, 25)
$tabManage.Controls.Add($btnRunTask)

$btnEnableTask = New-Object System.Windows.Forms.Button
$btnEnableTask.Text = "Enable Task"
$btnEnableTask.Location = New-Object System.Drawing.Point(110, 520)
$btnEnableTask.Size = New-Object System.Drawing.Size(90, 25)
$tabManage.Controls.Add($btnEnableTask)

$btnDisableTask = New-Object System.Windows.Forms.Button
$btnDisableTask.Text = "Disable Task"
$btnDisableTask.Location = New-Object System.Drawing.Point(210, 520)
$btnDisableTask.Size = New-Object System.Drawing.Size(90, 25)
$tabManage.Controls.Add($btnDisableTask)

$btnDeleteTask = New-Object System.Windows.Forms.Button
$btnDeleteTask.Text = "Delete Task"
$btnDeleteTask.Location = New-Object System.Drawing.Point(310, 520)
$btnDeleteTask.Size = New-Object System.Drawing.Size(90, 25)
$tabManage.Controls.Add($btnDeleteTask)

# Initially disable management buttons immediately after their creation.
# This ensures the objects exist before we try to set their properties for the first time.
$btnRunTask.Enabled = $false
$btnEnableTask.Enabled = $false
$btnDisableTask.Enabled = $false
$btnDeleteTask.Enabled = $false

#endregion Tab: View & Manage Tasks

#region Event Handler: Deploy Button Click (Existing Logic, moved to $tabDeploy)
$btnDeploy.Add_Click({
    $scriptPath = $txtScript.Text
    $listPath = $txtComputers.Text
    $taskName = $txtTaskName.Text
    $time = $txtTime.Text
    $removeOld = $chkRemoveOld.Checked
    $logFile = "$env:TEMP\TaskDeploy_$(Get-Date -f 'yyyyMMdd_HHmmss').csv"

    # Input validation (local machine)
    if (!(Test-Path $listPath)) {
        [System.Windows.Forms.MessageBox]::Show("Computer List file is required and must exist.","Input Missing","OK","Warning")
        return
    }
    
    if (!(Test-Path $scriptPath)) {
        [System.Windows.Forms.MessageBox]::Show("The selected script/executable ('$scriptPath') does not exist on your local machine. Please verify the path.","Script Not Found","OK","Error")
        return
    }

    [datetime]$nullTime = [datetime]::MinValue 
    if (-not [datetime]::TryParseExact($time, 'HH:mm', [cultureinfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::None, [ref]$nullTime)) {
        [System.Windows.Forms.MessageBox]::Show("Invalid time format. Use HH:mm (e.g. 09:30).","Invalid Time","OK","Warning")
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($taskName)) {
        [System.Windows.Forms.MessageBox]::Show("Scheduled Task Name cannot be empty.","Input Missing","OK","Warning")
        return
    }

    $txtDeployLog.AppendText("Deployment started at $(Get-Date)...`r`n")
@"
Computer,Status,Message
"@ | Out-File $logFile -Encoding UTF8

    $localFileName = [System.IO.Path]::GetFileName($scriptPath) 
    $remoteScriptPathOnTarget = Join-Path -Path "C:\temp" -ChildPath $localFileName

    $computers = Get-Content $listPath
    foreach ($computer in $computers) {
        try {
            $txtDeployLog.AppendText("Processing " + $computer + "...`r`n") 
            $sessionParams = Get-SessionParams $computer
            $session = New-PSSession @sessionParams

            Invoke-Command -Session $session -ScriptBlock {
                param($remoteScriptTarget) 
                $remoteTempDir = "C:\temp"
                
                try {
                    New-Item -Path $remoteTempDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                } catch {
                    Write-Error "Failed to create directory $remoteTempDir on $env:COMPUTERNAME: $($_.Exception.Message)"
                    throw 
                }
                
                if (Test-Path $remoteScriptTarget) { 
                    try {
                        Write-Output "Attempting to remove existing file $remoteScriptTarget on $env:COMPUTERNAME..."
                        Remove-Item -Path $remoteScriptTarget -Force -ErrorAction Stop
                        Write-Output "Existing file removed."
                    } catch {
                        Write-Warning "Could not remove existing file $($remoteScriptTarget) on $env:COMPUTERNAME: $($_.Exception.Message). Attempting to overwrite."
                    }
                }
            } -ArgumentList $remoteScriptPathOnTarget 

            Start-Sleep -Milliseconds 500 

            Copy-Item -Path $scriptPath -Destination $remoteScriptPathOnTarget -ToSession $session -Force

            if ($removeOld) {
                Invoke-Command -Session $session -ScriptBlock {
                    param($taskName)
                    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
                        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
                    }
                } -ArgumentList $taskName
            }

            Invoke-Command -Session $session -ScriptBlock {
                param ($taskName, $timeString, $remoteScriptTarget)
                
                $parsedTime = [datetime]::ParseExact($timeString, "HH:mm", [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None)
                $currentDate = (Get-Date).Date 
                $triggerDateTime = $currentDate.AddHours($parsedTime.Hour).AddMinutes($parsedTime.Minute)
                
                $trigger = New-ScheduledTaskTrigger -Daily -At $triggerDateTime

                $fileExtension = [System.IO.Path]::GetExtension($remoteScriptTarget).ToLower()
                $actionExecute = ""
                $actionArguments = ""

                switch ($fileExtension) {
                    ".ps1" {
                        $actionExecute = "powershell.exe"
                        $actionArguments = "-ExecutionPolicy Bypass -File `"$remoteScriptTarget`""
                    }
                    ".bat" { 
                        $actionExecute = "cmd.exe"
                        $actionArguments = "/c `"$remoteScriptTarget`"" 
                    }
                    ".cmd" { 
                        $actionExecute = "cmd.exe"
                        $actionArguments = "/c `"$remoteScriptTarget`"" 
                    }
                    ".exe" {
                        $actionExecute = $remoteScriptTarget 
                        $actionArguments = "" 
                    }
                    default {
                        Write-Warning "Unsupported file type '$fileExtension' detected for '$remoteScriptTarget' on $env:COMPUTERNAME. Attempting direct execution. Verify this is expected."
                        $actionExecute = $remoteScriptTarget 
                        $actionArguments = ""
                    }
                }
                
                $actionParams = @{
                    Execute = $actionExecute
                }
                if (-not [string]::IsNullOrEmpty($actionArguments)) {
                    $actionParams.Add('Argument', $actionArguments)
                }

                $action = New-ScheduledTaskAction @actionParams

                Register-ScheduledTask -TaskName $taskName -Trigger $trigger -Action $action -RunLevel Highest -User "SYSTEM" -Force
                
                $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
                if (-not $task) {
                    throw "Scheduled task registration failed silently on $env:COMPUTERNAME"
                }
            } -ArgumentList $taskName, $time, $remoteScriptPathOnTarget 

            Remove-PSSession $session
            $txtDeployLog.AppendText("Success: " + $computer + ": Task created`r`n") 
            "$computer,Success,Task created" | Out-File -Append $logFile
        } catch {
            $txtDeployLog.AppendText("Error: " + $computer + ": $($_.Exception.Message)`r`n") 
            "$computer,Error,$($_.Exception.Message.Replace(',',''))" | Out-File -Append $logFile
        }
    }

    $txtDeployLog.AppendText("Deployment complete. Log: $logFile`r`n") 
})
#endregion

#region Event Handler: Connect & Load Tasks Button Click
$btnConnect.Add_Click({
    LoadTasksIntoDataGridView `
        -ComputerName $txtManageComputer.Text.Trim() `
        -LogTextBox $txtManageLog `
        -DataGridView $dgvTasks
})
#endregion

#region Event Handler: DataGridView Selection Changed
$dgvTasks.Add_SelectionChanged({
    $hasSelection = ($dgvTasks.SelectedRows.Count -gt 0)
    # FIX 2B: Diagnostic message and null checks for buttons
    $global:txtManageLog.AppendText("DEBUG: Selection Changed. Has selection = $hasSelection, SelectedRows.Count = $($dgvTasks.SelectedRows.Count)`r`n")
    
    if ($global:btnRunTask) { $global:btnRunTask.Enabled = $hasSelection }
    if ($global:btnEnableTask) { $global:btnEnableTask.Enabled = $hasSelection }
    if ($global:btnDisableTask) { $global:btnDisableTask.Enabled = $hasSelection }
    if ($global:btnDeleteTask) { $global:btnDeleteTask.Enabled = $hasSelection }
})
#endregion

#region Event Handlers: Management Buttons
$btnRunTask.Add_Click({
    $selectedRow = $dgvTasks.SelectedRows[0]
    $taskName = $selectedRow.Cells["Name"].Value
    $computerName = $txtManageComputer.Text.Trim()

    $scriptBlock = {
        param($TaskNameToRun)
        Start-ScheduledTask -TaskName $TaskNameToRun -ErrorAction Stop
    }

    PerformTaskAction `
        -ActionName "run" `
        -ComputerName $computerName `
        -TaskName $taskName `
        -LogTextBox $txtManageLog `
        -DataGridView $dgvTasks `
        -ActionScriptBlock $scriptBlock
})

$btnEnableTask.Add_Click({
    $selectedRow = $dgvTasks.SelectedRows[0]
    $taskName = $selectedRow.Cells["Name"].Value
    $computerName = $txtManageComputer.Text.Trim()

    $scriptBlock = {
        param($TaskNameToEnable)
        Enable-ScheduledTask -TaskName $TaskNameToEnable -ErrorAction Stop
    }

    PerformTaskAction `
        -ActionName "enable" `
        -ComputerName $computerName `
        -TaskName $taskName `
        -LogTextBox $txtManageLog `
        -DataGridView $dgvTasks `
        -ActionScriptBlock $scriptBlock
})

$btnDisableTask.Add_Click({
    $selectedRow = $dgvTasks.SelectedRows[0]
    $taskName = $selectedRow.Cells["Name"].Value
    $computerName = $txtManageComputer.Text.Trim()

    $scriptBlock = {
        param($TaskNameToDisable)
        Disable-ScheduledTask -TaskName $TaskNameToDisable -ErrorAction Stop
    }

    PerformTaskAction `
        -ActionName "disable" `
        -ComputerName $computerName `
        -TaskName $taskName `
        -LogTextBox $txtManageLog `
        -DataGridView $dgvTasks `
        -ActionScriptBlock $scriptBlock
})

$btnDeleteTask.Add_Click({
    $selectedRow = $dgvTasks.SelectedRows[0]
    $taskName = $selectedRow.Cells["Name"].Value
    $computerName = $txtManageComputer.Text.Trim()

    $scriptBlock = {
        param($TaskNameToDelete)
        Unregister-ScheduledTask -TaskName $TaskNameToDelete -Confirm:$false -ErrorAction Stop
    }

    PerformTaskAction `
        -ActionName "delete" `
        -ComputerName $computerName `
        -TaskName $taskName `
        -LogTextBox $txtManageLog `
        -DataGridView $dgvTasks `
        -ActionScriptBlock $scriptBlock `
        -Confirm:$true # Add confirmation for delete action
})
#endregion

# Show the form
$frmMain.ShowDialog() | Out-Null
