# Hide PowerShell Console Window
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Mellanox Temperature Monitor"
$form.Size = New-Object System.Drawing.Size(800,810)  # Increased height for chart
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false
$form.MinimizeBox = $true

# Tool path settings
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Location = New-Object System.Drawing.Point(10,10)
$pathLabel.Size = New-Object System.Drawing.Size(100,20)
$pathLabel.Text = "MFT Tool Path:"

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(110,10)
$pathTextBox.Size = New-Object System.Drawing.Size(550,20)
$pathTextBox.Text = "C:\Program Files\Mellanox\WinMFT"

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(670,10)
$browseButton.Size = New-Object System.Drawing.Size(100,25)
$browseButton.Text = "Browse"

# Device list
$deviceLabel = New-Object System.Windows.Forms.Label
$deviceLabel.Location = New-Object System.Drawing.Point(10,40)
$deviceLabel.Size = New-Object System.Drawing.Size(100,20)
$deviceLabel.Text = "Devices:"

$deviceListBox = New-Object System.Windows.Forms.ListBox
$deviceListBox.Location = New-Object System.Drawing.Point(110,40)
$deviceListBox.Size = New-Object System.Drawing.Size(550,100)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(670,40)
$refreshButton.Size = New-Object System.Drawing.Size(100,25)
$refreshButton.Text = "Refresh"

# Temperature display and control
$tempLabel = New-Object System.Windows.Forms.Label
$tempLabel.Location = New-Object System.Drawing.Point(10,155)
$tempLabel.Size = New-Object System.Drawing.Size(100,30)
$tempLabel.Text = "Temperature:"
$tempLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

# Add warning title label
$warningLabel = New-Object System.Windows.Forms.Label
$warningLabel.Location = New-Object System.Drawing.Point(10,190)
$warningLabel.Size = New-Object System.Drawing.Size(100,30)
$warningLabel.Text = "Warnings:"
$warningLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

# Add error message display
$errorDisplay = New-Object System.Windows.Forms.Label
$errorDisplay.Location = New-Object System.Drawing.Point(110,190)
$errorDisplay.Size = New-Object System.Drawing.Size(550,30)
$errorDisplay.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$errorDisplay.ForeColor = [System.Drawing.Color]::LightGray
$errorDisplay.Text = "No active Warning/Errors"
$errorDisplay.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$errorDisplay.BackColor = [System.Drawing.SystemColors]::Window
$errorDisplay.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

$intervalLabel = New-Object System.Windows.Forms.Label
$intervalLabel.Location = New-Object System.Drawing.Point(440,155)
$intervalLabel.Size = New-Object System.Drawing.Size(100,30)
$intervalLabel.Text = "Interval (s):"
$intervalLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

$intervalNumeric = New-Object System.Windows.Forms.NumericUpDown
$intervalNumeric.Location = New-Object System.Drawing.Point(540,155)
$intervalNumeric.Size = New-Object System.Drawing.Size(60,30)
$intervalNumeric.Minimum = 3
$intervalNumeric.Maximum = 10
$intervalNumeric.Value = 5
$intervalNumeric.DecimalPlaces = 0

$tempDisplay = New-Object System.Windows.Forms.TextBox
$tempDisplay.Location = New-Object System.Drawing.Point(110,150)
$tempDisplay.Size = New-Object System.Drawing.Size(100,30)
$tempDisplay.ReadOnly = $true
$tempDisplay.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
$tempDisplay.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center

$getOnceButton = New-Object System.Windows.Forms.Button
$getOnceButton.Location = New-Object System.Drawing.Point(220,150)
$getOnceButton.Size = New-Object System.Drawing.Size(100,30)
$getOnceButton.Text = "Get Once"
$getOnceButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

$monitorButton = New-Object System.Windows.Forms.Button
$monitorButton.Location = New-Object System.Drawing.Point(330,150)
$monitorButton.Size = New-Object System.Drawing.Size(100,30)
$monitorButton.Text = "Start Monitor"
$monitorButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

# Log display
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(10,230)
$logBox.Size = New-Object System.Drawing.Size(650,320)  # Reduced width to make room for export button
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true

# Add export log button
$exportLogButton = New-Object System.Windows.Forms.Button
$exportLogButton.Location = New-Object System.Drawing.Point(670,230)
$exportLogButton.Size = New-Object System.Drawing.Size(100,25)
$exportLogButton.Text = "Export Log"

# Add export chart button
$exportChartButton = New-Object System.Windows.Forms.Button
$exportChartButton.Location = New-Object System.Drawing.Point(670,260)
$exportChartButton.Size = New-Object System.Drawing.Size(100,25)
$exportChartButton.Text = "Export Chart"

# Temperature Chart
$chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chart.Location = New-Object System.Drawing.Point(10,560)
$chart.Size = New-Object System.Drawing.Size(760,200)
$chart.BackColor = [System.Drawing.SystemColors]::Window
$chart.BorderlineDashStyle = [System.Windows.Forms.DataVisualization.Charting.ChartDashStyle]::Solid
$chart.BorderlineColor = [System.Drawing.SystemColors]::ActiveBorder
$chart.BorderlineWidth = 2

# Add prompt text to chart
$chartTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
$chartTitle.Text = "Click 'Start Monitor' to display temperature trend"
$chartTitle.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)  # Update font style
$chartTitle.ForeColor = [System.Drawing.Color]::LightGray
$chartTitle.Docking = [System.Windows.Forms.DataVisualization.Charting.Docking]::Top
$chartTitle.IsDockedInsideChartArea = $true
$chartTitle.Position.X = 25  # Adjust horizontal position to center
$chartTitle.Position.Y = 43  # Adjust vertical position to center
$chartTitle.Alignment = [System.Drawing.StringAlignment]::Center  # Set horizontal center alignment
$chart.Titles.Add($chartTitle)

# Configure chart area
$chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$chart.ChartAreas.Add($chartArea)
$chartArea.AxisX.Title = "Time"
$chartArea.AxisY.Title = "Temperature (C)"
$chartArea.AxisX.LabelStyle.Format = "HH:mm:ss"
$chartArea.BackColor = [System.Drawing.Color]::White

# Configure grid lines
$chartArea.AxisX.MajorGrid.LineColor = [System.Drawing.Color]::LightGray
$chartArea.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::LightGray
$chartArea.AxisX.MinorGrid.Enabled = $true
$chartArea.AxisY.MinorGrid.Enabled = $true
$chartArea.AxisX.MinorGrid.LineColor = [System.Drawing.Color]::FromArgb(240,240,240)
$chartArea.AxisY.MinorGrid.LineColor = [System.Drawing.Color]::FromArgb(240,240,240)

# Set fixed ranges for better stability
$chartArea.AxisY.Minimum = 0
$chartArea.AxisY.Maximum = 100
$chartArea.AxisY.Interval = 10
$chartArea.AxisY.MinorGrid.Interval = 5

# Configure series
$series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
$series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
$series.XValueType = [System.Windows.Forms.DataVisualization.Charting.ChartValueType]::DateTime
$series.Color = [System.Drawing.Color]::Blue
$series.BorderWidth = 2
$chart.Series.Add($series)

# Enable chart zooming and scrolling
$chartArea.CursorX.AutoScroll = $true
$chartArea.CursorX.IsUserSelectionEnabled = $true
$chartArea.CursorY.AutoScroll = $true
$chartArea.CursorY.IsUserSelectionEnabled = $true

# Set initial time range without data points
$currentTime = Get-Date
$chartArea.AxisX.Minimum = $currentTime.AddMinutes(-5).ToOADate()
$chartArea.AxisX.Maximum = $currentTime.AddMinutes(1).ToOADate()

# Global variables
$script:monitoring = $false
$script:timer = $null
$script:maxDataPoints = 50  # Maximum number of points to show on chart

# Function: Write log
function Write-Log {
    param(
        $Message,
        [switch]$IsError
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logBox.AppendText("[$timestamp] $Message`r`n")
    $logBox.ScrollToCaret()
    
    if ($IsError) {
        if ($null -ne $script:syncContext) {
            $script:syncContext.Post({ 
                $errorDisplay.ForeColor = [System.Drawing.Color]::Red
                $errorDisplay.Text = $Message 
            }, $null)
        } else {
            $form.Invoke([Action]{ 
                $errorDisplay.ForeColor = [System.Drawing.Color]::Red
                $errorDisplay.Text = $Message 
            })
        }
    } else {
        if ($null -ne $script:syncContext) {
            $script:syncContext.Post({ 
                $errorDisplay.ForeColor = [System.Drawing.Color]::LightGray
                $errorDisplay.Text = "No active Warning/Errors" 
            }, $null)
        } else {
            $form.Invoke([Action]{ 
                $errorDisplay.ForeColor = [System.Drawing.Color]::LightGray
                $errorDisplay.Text = "No active Warning/Errors" 
            })
        }
    }
}

# Function: Get Device List
function Get-MellanoxDevices {
    try {
        $mdevicesPath = Join-Path $pathTextBox.Text "mdevices.exe"
        
        # Check if file exists
        if (-not (Test-Path $mdevicesPath)) {
            Write-Log "Error: Cannot find mdevices.exe at path: $mdevicesPath" -IsError
            return @()
        }

        Write-Log "Executing command: $mdevicesPath status -vv"
        $output = & $mdevicesPath status -vv 2>&1 | Out-String
        Write-Log "Raw output:`n$output"
        
        # Use more flexible regex pattern to match devices
        $devices = $output | Select-String -Pattern "mt\d+_pciconf\d+|mlx\d+_\d+" -AllMatches | 
                  ForEach-Object { $_.Matches } | 
                  ForEach-Object { $_.Value } |
                  Select-Object -Unique

        if ($devices.Count -eq 0) {
            Write-Log "Warning: No Mellanox devices found" -IsError
        } else {
            Write-Log "Found devices: $($devices -join ', ')"
        }
        
        return $devices
    }
    catch {
        Write-Log "Error: Failed to get device list - $_" -IsError
        return @()
    }
}

# Function: Get Temperature
function Get-Temperature {
    param($DeviceName)
    try {
        $mgetTempPath = Join-Path $pathTextBox.Text "mget_temp_ext.exe"
        Write-Log "Executing: $mgetTempPath -d $DeviceName"
        $output = & $mgetTempPath -d $DeviceName 2>&1
        Write-Log "Output: $output"
        
        $temp = $output | ForEach-Object {
            if ($_ -match '^\s*(\d+)\s*$') {
                return $matches[1]
            }
        } | Select-Object -First 1

        if ($temp) {
            Write-Log "Successfully parsed temperature: $temp"
            return $temp
        }
        Write-Log "Could not parse temperature from output" -IsError
        throw "Invalid temperature output"
    }
    catch {
        Write-Log "Error getting temperature: $_" -IsError
        return $null
    }
}

# Event handler: Browse button
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select Mellanox Tools Directory"
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $pathTextBox.Text = $folderBrowser.SelectedPath
    }
})

# Event handler: Refresh button
$refreshButton.Add_Click({
    $deviceListBox.Items.Clear()
    $devices = Get-MellanoxDevices
    foreach ($device in $devices) {
        if ($null -ne $device -and $device.Trim() -ne '') {
            $deviceListBox.Items.Add($device)
        }
    }
    
    # Add warning if no devices were added
    if ($deviceListBox.Items.Count -eq 0) {
        Write-Log "Warning: No Mellanox devices found" -IsError
    } else {
        Write-Log "Added $($deviceListBox.Items.Count) device(s) to the list"
        # Auto-select the first device
        $deviceListBox.SelectedIndex = 0
        Write-Log "Auto-selected device: $($deviceListBox.SelectedItem)"
    }
})

# Event handler: Get temperature once
$getOnceButton.Add_Click({
    if ($deviceListBox.SelectedItem) {
        $temp = Get-Temperature $deviceListBox.SelectedItem
        if ($temp) {
            $tempDisplay.Text = "$temp C"
        }
        else {
            $tempDisplay.Text = "Error"
        }
    }
    else {
        Write-Log "Please select a device"
    }
})

# Function: Stop monitoring safely with force option
function Stop-Monitoring {
    Write-Log "Attempting to force stop monitoring..."
    
    # Immediately set monitoring flag to false
    $script:monitoring = $false
    
    # Force stop timer with multiple attempts if needed
    try {
        if ($null -ne $script:timer) {
            # Disable timer events first
            $script:timer.Enabled = $false
            
            # Force stop in a separate runspace if needed
            $stopScript = {
                param($timer)
                $timer.Stop()
                $timer.Dispose()
            }
            
            $job = Start-Job -ScriptBlock $stopScript -ArgumentList $script:timer
            
            # Wait for job to complete with timeout
            if (-not (Wait-Job $job -Timeout 5)) {
                Write-Log "Warning: Timer stop operation timed out, forcing cleanup"
                Stop-Job $job
                Remove-Job $job -Force
            } else {
                Remove-Job $job
            }
            
            # Clear timer reference
            $script:timer = $null
            Write-Log "Timer forcefully stopped and disposed"
        }
    }
    catch {
        Write-Log "Error during forced timer cleanup: $_"
        # Last resort cleanup
        $script:timer = $null
    }
    
    # Force UI update
    try {
        $monitorButton.Invoke([Action]{
            $monitorButton.Text = "Start Monitor"
            $monitorButton.Enabled = $true
        })
        Write-Log "Monitoring forcefully stopped"
    }
    catch {
        Write-Log "Error updating UI: $_"
    }
}

# Function: Update Chart
function Update-TemperatureChart {
    param($Temperature)
    
    try {
        if ($chart.Series[0].Points.Count -gt $script:maxDataPoints) {
            $chart.Series[0].Points.RemoveAt(0)
        }
        
        $currentTime = Get-Date
        $chart.Series[0].Points.AddXY($currentTime, $Temperature)
        
        # Update X axis range to show last 5 minutes
        $chartArea.AxisX.Minimum = $currentTime.AddMinutes(-5).ToOADate()
        $chartArea.AxisX.Maximum = $currentTime.AddMinutes(1).ToOADate()
        
        # Only adjust Y axis if temperature is outside current range
        if ($Temperature -gt $chartArea.AxisY.Maximum) {
            $chartArea.AxisY.Maximum = [Math]::Ceiling($Temperature / 10) * 10
        }
        elseif ($Temperature -lt $chartArea.AxisY.Minimum) {
            $chartArea.AxisY.Minimum = [Math]::Floor($Temperature / 10) * 10
        }
        
        $chart.Invalidate()
    }
    catch {
        Write-Log "Error updating chart: $_" -IsError
    }
}

# Event Handler: Start/Stop Monitor
$monitorButton.Add_Click({
    if (-not $script:monitoring) {
        if ($deviceListBox.SelectedItem) {
            try {
                $script:monitoring = $true
                $monitorButton.Text = "Stop Monitor"
                Write-Log "Starting monitoring..."
                
                # Hide the prompt text
                $chart.Titles[0].Text = ""
                
                # Clear existing chart data when starting new monitoring
                $chart.Series[0].Points.Clear()
                $currentTime = Get-Date
                $chartArea.AxisX.Minimum = $currentTime.AddMinutes(-5).ToOADate()
                $chartArea.AxisX.Maximum = $currentTime.AddMinutes(1).ToOADate()
                
                $script:timer = New-Object System.Windows.Forms.Timer
                # Convert seconds to milliseconds
                $script:timer.Interval = $intervalNumeric.Value * 1000
                Write-Log "Monitor interval set to $($intervalNumeric.Value) seconds"
                
                # Store the UI thread's synchronization context at the script level
                $script:syncContext = [System.Threading.SynchronizationContext]::Current
                
                if ($null -eq $script:syncContext) {
                    Write-Log "Warning: No synchronization context available, creating fallback..."
                    [System.Windows.Forms.Application]::EnableVisualStyles()
                    $script:syncContext = [System.Windows.Forms.WindowsFormsSynchronizationContext]::Current
                }
                
                $script:timer.Add_Tick({
                    if (-not $script:monitoring) {
                        Stop-Monitoring
                        return
                    }
                    
                    try {
                        $deviceName = $deviceListBox.SelectedItem
                        $tempValue = Get-Temperature $deviceName
                        
                        if ($null -ne $script:syncContext) {
                            $script:syncContext.Post(
                                {
                                    param($temp)
                                    if ($null -ne $temp -and $temp -ne '') {
                                        $tempDisplay.Text = "$temp C"
                                        Update-TemperatureChart $temp
                                        Write-Log "Temperature updated: $temp C"
                                    } else {
                                        $tempDisplay.Text = "Error"
                                        Write-Log "Failed to get temperature" -IsError
                                    }
                                }, 
                                $tempValue
                            )
                        } else {
                            $form.Invoke(
                                [Action]{
                                    param($temp)
                                    if ($null -ne $temp -and $temp -ne '') {
                                        $tempDisplay.Text = "$temp C"
                                        Update-TemperatureChart $temp
                                        Write-Log "Temperature updated: $temp C"
                                    } else {
                                        $tempDisplay.Text = "Error"
                                        Write-Log "Failed to get temperature" -IsError
                                    }
                                },
                                $tempValue
                            )
                        }
                    }
                    catch {
                        Write-Log "Critical Monitor Error: $_" -IsError
                        Stop-Monitoring
                    }
                })
                
                $script:timer.Start()
                Write-Log "Monitoring started with $($intervalNumeric.Value) second interval"
            }
            catch {
                Write-Log "Error starting monitor: $_"
                Stop-Monitoring
            }
        }
        else {
            Write-Log "Please select a device"
        }
    }
    else {
        $monitorButton.Enabled = $false
        Write-Log "Stop requested, forcing monitoring to stop..."
        Stop-Monitoring
    }
})

# Event Handler: Export Log
$exportLogButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    $saveFileDialog.DefaultExt = "txt"
    $saveFileDialog.AddExtension = $true
    $saveFileDialog.FileName = "$env:COMPUTERNAME`_NIC_Temperature_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $logBox.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
            Write-Log "Log file exported successfully to: $($saveFileDialog.FileName)"
        }
        catch {
            Write-Log "Error exporting log file: $_" -IsError
        }
    }
})

# Event Handler: Export Chart
$exportChartButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "PNG Image (*.png)|*.png|JPEG Image (*.jpg)|*.jpg|All Files (*.*)|*.*"
    $saveFileDialog.DefaultExt = "png"
    $saveFileDialog.AddExtension = $true
    $saveFileDialog.FileName = "$env:COMPUTERNAME`_NIC_Temperature_Chart_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            # Save chart as image
            switch ([System.IO.Path]::GetExtension($saveFileDialog.FileName).ToLower()) {
                ".png" { $chart.SaveImage($saveFileDialog.FileName, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png) }
                ".jpg" { $chart.SaveImage($saveFileDialog.FileName, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Jpeg) }
                default { $chart.SaveImage($saveFileDialog.FileName, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png) }
            }
            Write-Log "Chart exported successfully to: $($saveFileDialog.FileName)"
        }
        catch {
            Write-Log "Error exporting chart: $_" -IsError
        }
    }
})

# Add form load event handler
$form.Add_Load({
    Write-Log "Application started, performing initial device refresh..."
    $refreshButton.PerformClick()
})

# Add form closing event handler
$form.Add_FormClosing({
    param($sender, $e)
    Write-Log "Form closing, cleaning up..."
    Stop-Monitoring
})

# Add controls to form
$form.Controls.AddRange(@(
    $pathLabel, $pathTextBox, $browseButton,
    $deviceLabel, $deviceListBox, $refreshButton,
    $tempLabel, $tempDisplay, $getOnceButton, $monitorButton,
    $intervalLabel, $intervalNumeric,
    $warningLabel, $errorDisplay,
    $logBox, $exportLogButton, $exportChartButton,
    $chart
))

# Show form
$form.ShowDialog()

