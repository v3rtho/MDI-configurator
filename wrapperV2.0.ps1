# Microsoft Defender for Identity - PowerShell GUI Wrapper
# Requires: PowerShell 5.1 or later

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic

# XAML Definition
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Defender for Identity Configurator" 
        Height="700" Width="900"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="200"/>
        </Grid.RowDefinitions>

        <!-- Title -->
        <Grid Grid.Row="0" Margin="0,0,0,15">
            <TextBlock Text="Defender for Identity configurator v1.0" 
                       FontSize="22" 
                       FontWeight="Bold"
                       HorizontalAlignment="Left"
                       VerticalAlignment="Center"/>
            <TextBlock Text="© Thomas Verheyden" 
                       HorizontalAlignment="Right"
                       VerticalAlignment="Center"
                       Foreground="Gray"
                       FontSize="11"/>
        </Grid>

        <!-- Tab Control -->
        <TabControl Grid.Row="1">
            
            <!-- Configuration Tab -->
            <TabItem Header="Configuration">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <Grid Margin="0,0,0,10">
                            <TextBlock Text="Set MDI Configuration" 
                                       FontSize="16" 
                                       FontWeight="Bold"
                                       HorizontalAlignment="Left"
                                       VerticalAlignment="Center"/>
                            <Button x:Name="btnSetConfig" 
                                    Content="Apply Configuration" 
                                    Width="150" 
                                    Height="30"
                                    HorizontalAlignment="Right"/>
                        </Grid>

                        <!-- Mode Selection -->
                        <GroupBox Header="Mode" Margin="0,0,0,15">
                            <StackPanel>
                                <RadioButton x:Name="rbDomain" 
                                           Content="Domain (GPO-based)" 
                                           IsChecked="True" 
                                           Margin="5"/>
                                <RadioButton x:Name="rbLocalMachine" 
                                           Content="LocalMachine (Registry-based)" 
                                           Margin="5"/>
                            </StackPanel>
                        </GroupBox>

                        <!-- Configuration Options -->
                        <GroupBox Header="Configuration Items" Margin="0,0,0,15">
                            <StackPanel>
                                <CheckBox x:Name="chkAll" 
                                        Content="All Configurations" 
                                        FontWeight="Bold"
                                        Margin="5,5,5,10"/>
                                
                                <Separator Margin="5"/>
                                
                                <CheckBox x:Name="chkAdfsAuditing" 
                                        Content="ADFS Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkAdRecycleBin" 
                                        Content="AD Recycle Bin" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkAdvancedAuditPolicyCAs" 
                                        Content="Advanced Audit Policy (CAs)" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkAdvancedAuditPolicyDCs" 
                                        Content="Advanced Audit Policy (DCs)" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkCAAuditing" 
                                        Content="Certificate Authority (CA) Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkConfigurationContainerAuditing" 
                                        Content="Configuration Container Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkEntraConnectAuditing" 
                                        Content="Entra Connect Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkRemoteSAM" 
                                        Content="Remote SAM" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkDomainObjectAuditing" 
                                        Content="Domain Object Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkNTLMAuditing" 
                                        Content="NTLM Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkProcessorPerformance" 
                                        Content="Processor Performance" 
                                        Margin="5"/>
                            </StackPanel>
                        </GroupBox>

                        <!-- GPO Options -->
                        <GroupBox Header="GPO Options" Margin="0,0,0,15">
                            <StackPanel>
                                <CheckBox x:Name="chkSkipGpoLink" 
                                        Content="Skip GPO Link (Don't link GPOs to domain)" 
                                        Margin="5"
                                        ToolTip="When checked, GPOs will be created but not linked to the domain"/>
                                
                                <Label Content="GPO Name Prefix (optional):" Margin="5,10,5,0"/>
                                <TextBox x:Name="txtGpoNamePrefix" 
                                         Margin="5,5,5,5"
                                         ToolTip="Custom prefix for GPO names (e.g., 'MDI-' will create 'MDI-AuditPolicy')"/>
                            </StackPanel>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- Test/Validate Tab -->
            <TabItem Header="Test &amp; Validate">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <Grid Margin="0,0,0,10">
                            <TextBlock Text="Test MDI Configuration" 
                                       FontSize="16" 
                                       FontWeight="Bold"
                                       HorizontalAlignment="Left"
                                       VerticalAlignment="Center"/>
                            <Button x:Name="btnTestConfig" 
                                    Content="Run Test" 
                                    Width="120" 
                                    Height="30"
                                    HorizontalAlignment="Right"/>
                        </Grid>

                        <!-- Test Mode -->
                        <GroupBox Header="Test Mode" Margin="0,0,0,15">
                            <StackPanel>
                                <RadioButton x:Name="rbTestDomain" 
                                           Content="Domain" 
                                           IsChecked="True" 
                                           Margin="5"/>
                                <RadioButton x:Name="rbTestLocalMachine" 
                                           Content="LocalMachine" 
                                           Margin="5"/>
                            </StackPanel>
                        </GroupBox>

                        <!-- Test Configuration -->
                        <GroupBox Header="Configuration to Test" Margin="0,0,0,15">
                            <StackPanel>
                                <CheckBox x:Name="chkTestAll" 
                                        Content="All Configurations" 
                                        FontWeight="Bold"
                                        Margin="5,5,5,10"/>
                                
                                <Separator Margin="5"/>
                                
                                <CheckBox x:Name="chkTestAdfsAuditing" 
                                        Content="ADFS Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkTestAdRecycleBin" 
                                        Content="AD Recycle Bin" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkTestAdvancedAuditPolicyCAs" 
                                        Content="Advanced Audit Policy (CAs)" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkTestAdvancedAuditPolicyDCs" 
                                        Content="Advanced Audit Policy (DCs)" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkTestCAAuditing" 
                                        Content="Certificate Authority (CA) Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkTestConfigurationContainerAuditing" 
                                        Content="Configuration Container Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkTestEntraConnectAuditing" 
                                        Content="Entra Connect Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkTestRemoteSAM" 
                                        Content="Remote SAM" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkTestDomainObjectAuditing" 
                                        Content="Domain Object Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkTestNTLMAuditing" 
                                        Content="NTLM Auditing" 
                                        Margin="5"/>
                                <CheckBox x:Name="chkTestProcessorPerformance" 
                                        Content="Processor Performance" 
                                        Margin="5"/>
                            </StackPanel>
                        </GroupBox>

                        <!-- GPO Options for Testing -->
                        <GroupBox Header="GPO Options" Margin="0,0,0,15">
                            <StackPanel>
                                <Label Content="GPO Name Prefix (optional):" Margin="5,5,5,0"/>
                                <TextBox x:Name="txtTestGpoNamePrefix" 
                                         Margin="5,5,5,5"
                                         ToolTip="Custom prefix for GPO names to test (e.g., 'MDI-' will test 'MDI-AuditPolicy')"/>
                            </StackPanel>
                        </GroupBox>

                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- Reports Tab -->
            <TabItem Header="Reports">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="Generate Configuration Report" 
                                   FontSize="16" 
                                   FontWeight="Bold"
                                   Margin="0,0,0,10"/>

                        <Label Content="Report Output Path:"/>
                        <TextBox x:Name="txtReportPath" 
                                 Text="C:\Temp" 
                                 Margin="0,0,0,10"/>

                        <GroupBox Header="Report Mode" Margin="0,0,0,15">
                            <StackPanel>
                                <RadioButton x:Name="rbReportDomain" 
                                           Content="Domain" 
                                           IsChecked="True" 
                                           Margin="5"/>
                                <RadioButton x:Name="rbReportLocalMachine" 
                                           Content="LocalMachine" 
                                           Margin="5"/>
                            </StackPanel>
                        </GroupBox>

                        <!-- GPO Options for Reports -->
                        <GroupBox Header="GPO Options" Margin="0,0,0,15">
                            <StackPanel>
                                <Label Content="GPO Name Prefix (optional):" Margin="5,5,5,0"/>
                                <TextBox x:Name="txtReportGpoNamePrefix" 
                                         Margin="5,5,5,5"
                                         ToolTip="Custom prefix for GPO names to include in report (e.g., 'MDI-' will report on 'MDI-AuditPolicy')"/>
                            </StackPanel>
                        </GroupBox>

                        <CheckBox x:Name="chkOpenReport" 
                                  Content="Open HTML Report after generation" 
                                  IsChecked="True"
                                  Margin="0,0,0,15"/>

                        <Button x:Name="btnGenerateReport" 
                                Content="Generate Report" 
                                Width="150" 
                                Height="35"
                                HorizontalAlignment="Left"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- Proxy Configuration Tab -->
            <TabItem Header="Proxy">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="Sensor Proxy Configuration" 
                                   FontSize="16" 
                                   FontWeight="Bold"
                                   Margin="0,0,0,10"/>

                        <Label Content="Proxy URL:"/>
                        <TextBox x:Name="txtProxyUrl" 
                                 Text="http://proxy.contoso.com:8080" 
                                 Margin="0,0,0,10"/>

                        <Label Content="Proxy Username (optional):"/>
                        <TextBox x:Name="txtProxyUsername" 
                                 Margin="0,0,0,10"/>

                        <Label Content="Proxy Password (optional):"/>
                        <PasswordBox x:Name="txtProxyPassword" 
                                     Margin="0,0,0,15"/>

                        <StackPanel Orientation="Horizontal">
                            <Button x:Name="btnSetProxy" 
                                    Content="Set Proxy" 
                                    Width="120" 
                                    Height="35"
                                    Margin="0,0,10,0"/>
                            
                            <Button x:Name="btnGetProxy" 
                                    Content="Get Proxy" 
                                    Width="120" 
                                    Height="35"
                                    Margin="0,0,10,0"/>
                            
                            <Button x:Name="btnClearProxy" 
                                    Content="Clear Proxy" 
                                    Width="120" 
                                    Height="35"/>
                        </StackPanel>

                        <Separator Margin="0,20,0,20"/>

                        <Button x:Name="btnTestConnection" 
                                Content="Test Sensor API Connection" 
                                Width="200" 
                                Height="35"
                                HorizontalAlignment="Left"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- gMSA Tab -->
            <TabItem Header="gMSA">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="Group Managed Service Account (gMSA)" 
                                   FontSize="16" 
                                   FontWeight="Bold"
                                   Margin="0,0,0,10"/>

                        <TextBlock TextWrapping="Wrap" 
                                   Margin="0,0,0,15">
                            <Run Text="The gMSA is a Group Managed Service Account used by MDI to query domain controllers for suspicious activities."/>
                        </TextBlock>

                        <Label Content="gMSA Identity (username):"/>
                        <TextBox x:Name="txtGmsaIdentity" 
                                 Margin="0,0,0,10"
                                 ToolTip="e.g., mdiSvc01"/>

                        <Label Content="gMSA Group Name:"/>
                        <TextBox x:Name="txtGmsaGroupName" 
                                 Margin="0,0,0,15"
                                 ToolTip="e.g., mdiSvcGroup01"/>

                        <StackPanel Orientation="Horizontal" Margin="0,0,0,20">
                            <Button x:Name="btnCreateDSA" 
                                    Content="Create gMSA" 
                                    Width="140" 
                                    Height="35"
                                    Margin="0,0,10,0"/>
                            
                            <Button x:Name="btnTestDSA" 
                                    Content="Test gMSA" 
                                    Width="120" 
                                    Height="35"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

        </TabControl>

        <!-- Action Buttons -->
        <StackPanel Grid.Row="2" 
                    Orientation="Horizontal" 
                    HorizontalAlignment="Right"
                    Margin="0,10">
            <Button x:Name="btnClearOutput" 
                    Content="Clear Output" 
                    Width="120" 
                    Height="30"
                    Margin="0,0,10,0"/>
            <Button x:Name="btnCopyOutput" 
                    Content="Copy Output" 
                    Width="120" 
                    Height="30"/>
        </StackPanel>

        <!-- Output Area -->
        <GroupBox Grid.Row="3" Header="Output / Command Result" Margin="0,10,0,0">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <TextBox x:Name="txtOutput" 
                         IsReadOnly="True" 
                         TextWrapping="Wrap"
                         FontFamily="Consolas"
                         Background="#F5F5F5"
                         VerticalScrollBarVisibility="Auto"/>
            </ScrollViewer>
        </GroupBox>
    </Grid>
</Window>
"@

# Load XAML
$Reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($Reader)

# Get all named controls
$Controls = @{}
$nsmgr = New-Object System.Xml.XmlNamespaceManager($XAML.NameTable)
$nsmgr.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml")
$XAML.SelectNodes("//*[@x:Name]", $nsmgr) | ForEach-Object {
    $Controls[$_.Name] = $Window.FindName($_.Name)
}

# Helper function to execute PowerShell commands
function Invoke-MDICommand {
    param([string]$Command)
    
    $output = New-Object System.Text.StringBuilder
    
    try {
        # Add -Verbose to the command if not already present
        if ($Command -notlike "*-Verbose*") {
            $Command += " -Verbose"
        }
        
        $result = Invoke-Expression $Command *>&1
        
        if ($result) {
            foreach ($item in $result) {
                if ($item -is [System.Management.Automation.ErrorRecord]) {
                    [void]$output.AppendLine("Error: $($item.Exception.Message)")
                } elseif ($item -is [System.Management.Automation.VerboseRecord]) {
                    [void]$output.AppendLine("VERBOSE: $($item.Message)")
                } elseif ($item -is [System.Management.Automation.WarningRecord]) {
                    [void]$output.AppendLine("WARNING: $($item.Message)")
                } elseif ($item -is [System.Management.Automation.InformationRecord]) {
                    [void]$output.AppendLine("INFO: $($item.MessageData)")
                }
                # Removed the else block that was showing "Result: ..." for non-verbose output
            }
        }
        # Removed the "Result: False" output for empty results
    }
    catch {
        [void]$output.AppendLine("Exception: $($_.Exception.Message)")
    }
    
    return $output.ToString()
}

# Check if module is installed
function Test-ModuleInstalled {
    try {
        $module = Get-Module -ListAvailable -Name DefenderForIdentity
        if (-not $module) {
            $result = [System.Windows.MessageBox]::Show(
                "DefenderForIdentity PowerShell module is not installed.`n`nDo you want to see installation instructions?",
                "Module Not Found",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning)
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                $Controls.txtOutput.Text = @"
To install the DefenderForIdentity module, run PowerShell as Administrator and execute:

Install-Module -Name DefenderForIdentity -Force

Then restart this application.
"@
            }
        } else {
            $Controls.txtOutput.Text = "DefenderForIdentity module is installed and ready to use."
        }
    }
    catch {
        $Controls.txtOutput.Text = "Error checking module: $($_.Exception.Message)"
    }
}

# Event Handlers

# Check All checkbox handler
$Controls.chkAll.Add_Checked({
    $Controls.chkAdfsAuditing.IsChecked = $true
    $Controls.chkAdRecycleBin.IsChecked = $true
    $Controls.chkAdvancedAuditPolicyCAs.IsChecked = $true
    $Controls.chkAdvancedAuditPolicyDCs.IsChecked = $true
    $Controls.chkCAAuditing.IsChecked = $true
    $Controls.chkConfigurationContainerAuditing.IsChecked = $true
    $Controls.chkEntraConnectAuditing.IsChecked = $true
    $Controls.chkRemoteSAM.IsChecked = $true
    $Controls.chkDomainObjectAuditing.IsChecked = $true
    $Controls.chkNTLMAuditing.IsChecked = $true
    $Controls.chkProcessorPerformance.IsChecked = $true
})

$Controls.chkAll.Add_Unchecked({
    $Controls.chkAdfsAuditing.IsChecked = $false
    $Controls.chkAdRecycleBin.IsChecked = $false
    $Controls.chkAdvancedAuditPolicyCAs.IsChecked = $false
    $Controls.chkAdvancedAuditPolicyDCs.IsChecked = $false
    $Controls.chkCAAuditing.IsChecked = $false
    $Controls.chkConfigurationContainerAuditing.IsChecked = $false
    $Controls.chkEntraConnectAuditing.IsChecked = $false
    $Controls.chkRemoteSAM.IsChecked = $false
    $Controls.chkDomainObjectAuditing.IsChecked = $false
    $Controls.chkNTLMAuditing.IsChecked = $false
    $Controls.chkProcessorPerformance.IsChecked = $false
})

# Configuration tab - Mode selection handlers
$Controls.rbDomain.Add_Checked({
    $Controls.chkSkipGpoLink.IsEnabled = $true
    $Controls.txtGpoNamePrefix.IsEnabled = $true
})

$Controls.rbLocalMachine.Add_Checked({
    $Controls.chkSkipGpoLink.IsEnabled = $false
    $Controls.txtGpoNamePrefix.IsEnabled = $false
})

# Test Check All checkbox handler
$Controls.chkTestAll.Add_Checked({
    $Controls.chkTestAdfsAuditing.IsChecked = $true
    $Controls.chkTestAdRecycleBin.IsChecked = $true
    $Controls.chkTestAdvancedAuditPolicyCAs.IsChecked = $true
    $Controls.chkTestAdvancedAuditPolicyDCs.IsChecked = $true
    $Controls.chkTestCAAuditing.IsChecked = $true
    $Controls.chkTestConfigurationContainerAuditing.IsChecked = $true
    $Controls.chkTestEntraConnectAuditing.IsChecked = $true
    $Controls.chkTestRemoteSAM.IsChecked = $true
    $Controls.chkTestDomainObjectAuditing.IsChecked = $true
    $Controls.chkTestNTLMAuditing.IsChecked = $true
    $Controls.chkTestProcessorPerformance.IsChecked = $true
})

$Controls.chkTestAll.Add_Unchecked({
    $Controls.chkTestAdfsAuditing.IsChecked = $false
    $Controls.chkTestAdRecycleBin.IsChecked = $false
    $Controls.chkTestAdvancedAuditPolicyCAs.IsChecked = $false
    $Controls.chkTestAdvancedAuditPolicyDCs.IsChecked = $false
    $Controls.chkTestCAAuditing.IsChecked = $false
    $Controls.chkTestConfigurationContainerAuditing.IsChecked = $false
    $Controls.chkTestEntraConnectAuditing.IsChecked = $false
    $Controls.chkTestRemoteSAM.IsChecked = $false
    $Controls.chkTestDomainObjectAuditing.IsChecked = $false
    $Controls.chkTestNTLMAuditing.IsChecked = $false
    $Controls.chkTestProcessorPerformance.IsChecked = $false
})

# Test & Validate tab - Mode selection handlers
$Controls.rbTestDomain.Add_Checked({
    $Controls.txtTestGpoNamePrefix.IsEnabled = $true
    $Controls.chkTestAll.IsEnabled = $true
    $Controls.chkTestEntraConnectAuditing.IsEnabled = $true
    $Controls.chkTestRemoteSAM.IsEnabled = $true
})

$Controls.rbTestLocalMachine.Add_Checked({
    $Controls.txtTestGpoNamePrefix.IsEnabled = $false
    $Controls.chkTestAll.IsEnabled = $false
    $Controls.chkTestEntraConnectAuditing.IsEnabled = $false
    $Controls.chkTestRemoteSAM.IsEnabled = $false
})

# Reports tab - Mode selection handlers
$Controls.rbReportDomain.Add_Checked({
    $Controls.txtReportGpoNamePrefix.IsEnabled = $true
})

$Controls.rbReportLocalMachine.Add_Checked({
    $Controls.txtReportGpoNamePrefix.IsEnabled = $false
})

# Set Configuration button
$Controls.btnSetConfig.Add_Click({
    $mode = if ($Controls.rbDomain.IsChecked) { "Domain" } else { "LocalMachine" }
    $configs = @()
    
    if ($Controls.chkAll.IsChecked) {
        $configs = @("All")
    } else {
        if ($Controls.chkAdfsAuditing.IsChecked) { $configs += "AdfsAuditing" }
        if ($Controls.chkAdRecycleBin.IsChecked) { $configs += "AdRecycleBin" }
        if ($Controls.chkAdvancedAuditPolicyCAs.IsChecked) { $configs += "AdvancedAuditPolicyCAs" }
        if ($Controls.chkAdvancedAuditPolicyDCs.IsChecked) { $configs += "AdvancedAuditPolicyDCs" }
        if ($Controls.chkCAAuditing.IsChecked) { $configs += "CAAuditing" }
        if ($Controls.chkConfigurationContainerAuditing.IsChecked) { $configs += "ConfigurationContainerAuditing" }
        if ($Controls.chkEntraConnectAuditing.IsChecked) { $configs += "EntraConnectAuditing" }
        if ($Controls.chkRemoteSAM.IsChecked) { $configs += "RemoteSAM" }
        if ($Controls.chkDomainObjectAuditing.IsChecked) { $configs += "DomainObjectAuditing" }
        if ($Controls.chkNTLMAuditing.IsChecked) { $configs += "NTLMAuditing" }
        if ($Controls.chkProcessorPerformance.IsChecked) { $configs += "ProcessorPerformance" }
    }
    
    if ($configs.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Please select at least one configuration option.",
            "No Selection",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    # Check if EntraConnectAuditing or RemoteSAM is selected (or All) and prompt for service account
    $serviceAccount = $null
    if ($Controls.chkAll.IsChecked -or $Controls.chkEntraConnectAuditing.IsChecked -or $Controls.chkRemoteSAM.IsChecked) {
        $serviceAccount = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter the Service Account name for EntraConnectAuditing or RemoteSAM:",
            "Service Account Required",
            ""
        )
        
        if ([string]::IsNullOrWhiteSpace($serviceAccount)) {
            [System.Windows.MessageBox]::Show(
                "Service Account name is required for EntraConnectAuditing or RemoteSAM configuration.",
                "Missing Service Account",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning)
            return
        }
    }
    
    $configString = $configs -join ","
    
    # Build the command with optional parameters
    $command = "Set-MDIConfiguration -Mode $mode -Configuration $configString"
    
    # Add SkipGpoLink parameter if checkbox is checked
    if ($Controls.chkSkipGpoLink.IsChecked) {
        $command += " -SkipGpoLink"
    }
    
    # Add GpoNamePrefix parameter if text is provided
    $gpoPrefix = $Controls.txtGpoNamePrefix.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($gpoPrefix)) {
        $command += " -GpoNamePrefix '$gpoPrefix'"
    }
    
    # Add Identity parameter if provided
    if ($serviceAccount) {
        $command += " -Identity '$serviceAccount'"
    }
    
    # Set cursor to wait
    $Window.Cursor = [System.Windows.Input.Cursors]::Wait
    
    try {
        $Controls.txtOutput.Text = "Executing: $command`n`n"
        $Controls.txtOutput.Text += "This may take a few moments...`n`n"
        $Controls.txtOutput.Text += Invoke-MDICommand $command
    }
    finally {
        # Reset cursor back to normal
        $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

# Test Configuration button
$Controls.btnTestConfig.Add_Click({
    $mode = if ($Controls.rbTestDomain.IsChecked) { "Domain" } else { "LocalMachine" }
    $configs = @()
    
    if ($Controls.chkTestAll.IsChecked) {
        $configs = @("All")
    } else {
        if ($Controls.chkTestAdfsAuditing.IsChecked) { $configs += "AdfsAuditing" }
        if ($Controls.chkTestAdRecycleBin.IsChecked) { $configs += "AdRecycleBin" }
        if ($Controls.chkTestAdvancedAuditPolicyCAs.IsChecked) { $configs += "AdvancedAuditPolicyCAs" }
        if ($Controls.chkTestAdvancedAuditPolicyDCs.IsChecked) { $configs += "AdvancedAuditPolicyDCs" }
        if ($Controls.chkTestCAAuditing.IsChecked) { $configs += "CAAuditing" }
        if ($Controls.chkTestConfigurationContainerAuditing.IsChecked) { $configs += "ConfigurationContainerAuditing" }
        if ($Controls.chkTestEntraConnectAuditing.IsChecked) { $configs += "EntraConnectAuditing" }
        if ($Controls.chkTestRemoteSAM.IsChecked) { $configs += "RemoteSAM" }
        if ($Controls.chkTestDomainObjectAuditing.IsChecked) { $configs += "DomainObjectAuditing" }
        if ($Controls.chkTestNTLMAuditing.IsChecked) { $configs += "NTLMAuditing" }
        if ($Controls.chkTestProcessorPerformance.IsChecked) { $configs += "ProcessorPerformance" }
    }
    
    if ($configs.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Please select at least one configuration option to test.",
            "No Selection",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    # Check if EntraConnectAuditing or RemoteSAM is selected (or All) and prompt for service account
    $serviceAccount = $null
    if ($Controls.chkTestAll.IsChecked -or $Controls.chkTestEntraConnectAuditing.IsChecked -or $Controls.chkTestRemoteSAM.IsChecked) {
        $serviceAccount = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter the Service Account name for EntraConnectAuditing or RemoteSAM:",
            "Service Account Required",
            ""
        )
        
        if ([string]::IsNullOrWhiteSpace($serviceAccount)) {
            [System.Windows.MessageBox]::Show(
                "Service Account name is required for EntraConnectAuditing or RemoteSAM configuration.",
                "Missing Service Account",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning)
            return
        }
    }
    
    $configString = $configs -join ","
    $command = "Test-MDIConfiguration -Mode $mode -Configuration $configString"
    
    # Add GpoNamePrefix parameter if text is provided
    $testGpoPrefix = $Controls.txtTestGpoNamePrefix.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($testGpoPrefix)) {
        $command += " -GpoNamePrefix '$testGpoPrefix'"
    }
    
    # Add Identity parameter if provided
    if ($serviceAccount) {
        $command += " -Identity '$serviceAccount'"
    }
    
    # Set cursor to wait
    $Window.Cursor = [System.Windows.Input.Cursors]::Wait
    
    try {
        $Controls.txtOutput.Text = "Executing: $command`n`n"
        $Controls.txtOutput.Text += Invoke-MDICommand $command
    }
    finally {
        # Reset cursor back to normal
        $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

# Generate Report button
$Controls.btnGenerateReport.Add_Click({
    $path = $Controls.txtReportPath.Text
    $mode = if ($Controls.rbReportDomain.IsChecked) { "Domain" } else { "LocalMachine" }
    $openReport = if ($Controls.chkOpenReport.IsChecked) { "-OpenHtmlReport" } else { "" }
    
    if ([string]::IsNullOrWhiteSpace($path)) {
        [System.Windows.MessageBox]::Show(
            "Please specify a report output path.",
            "Missing Path",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    # If Domain mode is selected, prompt for gMSA account
    $gmsaAccount = $null
    if ($Controls.rbReportDomain.IsChecked) {
        $gmsaAccount = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter the gMSA account name for the report:",
            "gMSA Account Required",
            ""
        )
        
        if ([string]::IsNullOrWhiteSpace($gmsaAccount)) {
            [System.Windows.MessageBox]::Show(
                "gMSA account name is required for Domain mode reports.",
                "Missing gMSA Account",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning)
            return
        }
    }
    
    $command = "New-MDIConfigurationReport -Path '$path' -Mode $mode $openReport"
    
    # Add GpoNamePrefix parameter if text is provided
    $reportGpoPrefix = $Controls.txtReportGpoNamePrefix.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($reportGpoPrefix)) {
        $command += " -GpoNamePrefix '$reportGpoPrefix'"
    }
    
    # Add Identity parameter if provided
    if ($gmsaAccount) {
        $command += " -Identity '$gmsaAccount'"
    }
    
    # Set cursor to wait
    $Window.Cursor = [System.Windows.Input.Cursors]::Wait
    
    try {
        $Controls.txtOutput.Text = "Executing: $command`n`n"
        $Controls.txtOutput.Text += Invoke-MDICommand $command
    }
    finally {
        # Reset cursor back to normal
        $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

# Set Proxy button
$Controls.btnSetProxy.Add_Click({
    $proxyUrl = $Controls.txtProxyUrl.Text
    $username = $Controls.txtProxyUsername.Text
    $password = $Controls.txtProxyPassword.Password
    
    if ([string]::IsNullOrWhiteSpace($proxyUrl)) {
        [System.Windows.MessageBox]::Show(
            "Please enter a proxy URL.",
            "Missing URL",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    if (-not [string]::IsNullOrWhiteSpace($username) -and -not [string]::IsNullOrWhiteSpace($password)) {
        $secPass = ConvertTo-SecureString $password -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential($username, $secPass)
        $command = "Set-MDISensorProxyConfiguration -ProxyUrl '$proxyUrl' -ProxyCredential `$cred"
        
        $Controls.txtOutput.Text = "Setting proxy configuration with credentials...`n`n"
        $Controls.txtOutput.Text += "Note: This will stop and restart MDI sensor services.`n`n"
        
        try {
            Set-MDISensorProxyConfiguration -ProxyUrl $proxyUrl -ProxyCredential $cred
            $Controls.txtOutput.Text += "Proxy configured successfully."
        }
        catch {
            $Controls.txtOutput.Text += "Error: $($_.Exception.Message)"
        }
    } else {
        $command = "Set-MDISensorProxyConfiguration -ProxyUrl '$proxyUrl'"
        
        $Controls.txtOutput.Text = "Setting proxy configuration...`n`n"
        $Controls.txtOutput.Text += "Note: This will stop and restart MDI sensor services.`n`n"
        $Controls.txtOutput.Text += Invoke-MDICommand $command
    }
})

# Get Proxy button
$Controls.btnGetProxy.Add_Click({
    $command = "Get-MDISensorProxyConfiguration"
    
    $Controls.txtOutput.Text = "Executing: $command`n`n"
    $Controls.txtOutput.Text += Invoke-MDICommand $command
})

# Clear Proxy button
$Controls.btnClearProxy.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Are you sure you want to clear the proxy configuration?",
        "Confirm Clear",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $command = "Clear-MDISensorProxyConfiguration"
        
        $Controls.txtOutput.Text = "Executing: $command`n`n"
        $Controls.txtOutput.Text += Invoke-MDICommand $command
    }
})

# Test Connection button
$Controls.btnTestConnection.Add_Click({
    $command = "Test-MDISensorApiConnection"
    
    # Set cursor to wait
    $Window.Cursor = [System.Windows.Input.Cursors]::Wait
    
    try {
        $Controls.txtOutput.Text = "Executing: $command`n`n"
        $Controls.txtOutput.Text += "Testing connectivity to Defender for Identity cloud services...`n`n"
        $Controls.txtOutput.Text += Invoke-MDICommand $command
    }
    finally {
        # Reset cursor back to normal
        $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

# Create DSA button
$Controls.btnCreateDSA.Add_Click({
    $gmsaIdentity = $Controls.txtGmsaIdentity.Text.Trim()
    $gmsaGroupName = $Controls.txtGmsaGroupName.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($gmsaIdentity) -or [string]::IsNullOrWhiteSpace($gmsaGroupName)) {
        [System.Windows.MessageBox]::Show(
            "Please enter both gMSA Identity and gMSA Group Name.",
            "Missing Information",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $command = "New-MDIDSA -Identity '$gmsaIdentity' -GmsaGroupName '$gmsaGroupName'"
    
    # Set cursor to wait
    $Window.Cursor = [System.Windows.Input.Cursors]::Wait
    
    try {
        $Controls.txtOutput.Text = "Creating gMSA account...`n`n"
        $Controls.txtOutput.Text += "Executing: $command`n`n"
        
        $result = New-MDIDSA -Identity $gmsaIdentity -GmsaGroupName $gmsaGroupName
        $Controls.txtOutput.Text += $result | Out-String
        $Controls.txtOutput.Text += "`n`ngMSA account created successfully!"
    }
    catch {
        $Controls.txtOutput.Text += "Error: $($_.Exception.Message)"
    }
    finally {
        # Reset cursor back to normal
        $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

# Test DSA button
$Controls.btnTestDSA.Add_Click({
    $gmsaIdentity = $Controls.txtGmsaIdentity.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($gmsaIdentity)) {
        [System.Windows.MessageBox]::Show(
            "Please enter the gMSA Identity before testing.",
            "Missing gMSA Identity",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $command = "Test-MDIDSA -Identity '$gmsaIdentity' -Detailed"
    
    # Set cursor to wait
    $Window.Cursor = [System.Windows.Input.Cursors]::Wait
    
    try {
        $Controls.txtOutput.Text = "Testing gMSA account with detailed output...`n`n"
        $Controls.txtOutput.Text += "Executing: $command`n`n"
        $Controls.txtOutput.Text += "Testing Directory Service Account permissions and delegations...`n`n"
        
        $result = Test-MDIDSA -Identity $gmsaIdentity -Detailed
        $Controls.txtOutput.Text += $result | Format-List * | Out-String
    }
    catch {
        $Controls.txtOutput.Text += "Error: $($_.Exception.Message)"
    }
    finally {
        # Reset cursor back to normal
        $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

# Clear Output button
$Controls.btnClearOutput.Add_Click({
    $Controls.txtOutput.Clear()
})

# Copy Output button
$Controls.btnCopyOutput.Add_Click({
    if (-not [string]::IsNullOrEmpty($Controls.txtOutput.Text)) {
        [System.Windows.Clipboard]::SetText($Controls.txtOutput.Text)
        [System.Windows.MessageBox]::Show(
            "Output copied to clipboard!",
            "Success",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information)
    }
})

# Check module on startup
Test-ModuleInstalled

# Show the window
$Window.ShowDialog() | Out-Null
