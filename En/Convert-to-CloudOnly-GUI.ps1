<#
.SYNOPSIS
    GUI tool to convert AD-synced users to cloud-only accounts in Entra ID.
    Supports single-user and bulk (CSV) operations.

.DESCRIPTION
    Steps per user:
      1. Move the AD account to a non-synced OU
      2. Run a delta sync -> the account is soft-deleted in Entra ID
      3. Pause the sync schedule
      4. Restore the account in Entra ID
      5. Clear onPremisesImmutableId via Graph API (handles federated UPN if necessary)
      6. Re-enable the sync schedule and run a delta sync

.REQUIREMENTS
    - Run as administrator
    - RSAT: Active Directory DS Tools installed
    - Microsoft.Graph PowerShell module
    - Network access (WinRM) to the Azure AD Connect server
    - Graph scope: User.ReadWrite.All

.NOTES
    Settings are persisted to Convert-to-CloudOnly-Settings.json next to the script.
#>

#Requires -Version 5.1
Set-StrictMode -Off

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# -------------------------------------------------------------
#  Settings helpers
# -------------------------------------------------------------
$SettingsFile = Join-Path $PSScriptRoot "Convert-to-CloudOnly-Settings.json"

function Load-Settings {
    if (Test-Path $SettingsFile) {
        try { return Get-Content $SettingsFile -Raw | ConvertFrom-Json }
        catch {}
    }
    return [PSCustomObject]@{
        TargetOU            = "OU=Disabled Objects,DC=contoso,DC=com"
        SyncServer          = "azadc.contoso.com"
        ManagedDomainSuffix = "contoso.onmicrosoft.com"
        LogFolder           = "$env:USERPROFILE\Documents\CloudOnly-Logs"
        SyncWaitSeconds     = 180
        RestoreWaitSeconds  = 20
    }
}

function Save-Settings ($s) {
    $s | ConvertTo-Json -Depth 3 | Set-Content $SettingsFile -Encoding UTF8
}

# -------------------------------------------------------------
#  Conversion core
# -------------------------------------------------------------
function Convert-UserToCloudOnly {
    param(
        [string]$SamAccountName,
        [string]$UPN,
        [PSObject]$Settings,
        [System.Management.Automation.PSCredential]$SyncCred,
        [System.Management.Automation.Runspaces.PSSession]$SyncSession,
        [bool]$DryRun,
        [scriptblock]$Log
    )

    function Write-Step ([string]$msg, [string]$lvl = "Info") { & $Log $msg $lvl }

    Write-Step "-- Starting conversion of $UPN --"

    # --- Step 1: Move AD user ---
    Write-Step "[1/6] Fetching and moving AD user '$SamAccountName' to $($Settings.TargetOU)..."
    if (-not $DryRun) {
        try {
            $adUser = Get-ADUser -Identity $SamAccountName -Properties *
        } catch {
            Write-Step "Could not find AD user '$SamAccountName': $_" "Error"
            return $false
        }

        if (-not (Test-Path $Settings.LogFolder)) { New-Item $Settings.LogFolder -ItemType Directory -Force | Out-Null }
        $exportPath = Join-Path $Settings.LogFolder ("$SamAccountName`_AD_pre-move_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml")
        try { $adUser | Export-Clixml -Path $exportPath } catch {}

        try {
            Move-ADObject -Identity $adUser.DistinguishedName -TargetPath $Settings.TargetOU
            Write-Step "AD user moved." "Success"
        } catch {
            Write-Step "Failed to move AD user: $_" "Error"
            return $false
        }
    } else {
        Write-Step "[DRY RUN] Would move '$SamAccountName' to $($Settings.TargetOU)." "Warning"
    }

    # --- Steps 2 & 3: Sync -> soft-delete, disable schedule ---
    Write-Step "[2/6] Running delta sync and disabling the sync schedule..."
    if (-not $DryRun) {
        try {
            Invoke-Command -Session $SyncSession -ScriptBlock {
                Import-Module ADSync -ErrorAction Stop
                Set-ADSyncScheduler -SyncCycleEnabled $false
                # Wait out any in-progress cycle before starting ours
                $deadline = (Get-Date).AddMinutes(5)
                while ((Get-ADSyncConnectorRunStatus) -and (Get-Date) -lt $deadline) { Start-Sleep 5 }
                # Start our delta sync and wait until it completes
                try { Start-ADSyncSyncCycle -PolicyType Delta | Out-Null } catch {
                    Start-Sleep 10
                    Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
                }
                Start-Sleep 5
                $deadline = (Get-Date).AddMinutes(5)
                while ((Get-ADSyncConnectorRunStatus) -and (Get-Date) -lt $deadline) { Start-Sleep 5 }
            }
            Write-Step "Sync complete and schedule paused. Short buffer ($($Settings.SyncWaitSeconds)s) for Entra processing..." "Info"
            Start-Sleep -Seconds $Settings.SyncWaitSeconds
        } catch {
            Write-Step "Problem with the sync server: $_" "Error"
            return $false
        }
    } else {
        Write-Step "[DRY RUN] Would run delta sync and pause the schedule." "Warning"
    }

    # --- Step 4: Connect Graph and restore ---
    Write-Step "[3/6] Connecting to Microsoft Graph..."
    if (-not $DryRun) {
        try {
            Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop -NoWelcome
            Write-Step "Connected to Graph." "Success"
        } catch {
            Write-Step "Failed to connect to Graph: $_" "Error"
            return $false
        }
    }

    Write-Step "[4/6] Looking for deleted user '$UPN' in Entra ID..."
    if (-not $DryRun) {
        $localPart = if ($UPN -match '@') { ($UPN -split '@')[0] } else { $UPN }
        $lpLower   = $localPart.ToLower()
        $deleted = $null
        for ($i = 0; $i -lt 5; $i++) {
            try {
                $all = @(Get-MgDirectoryDeletedItemAsUser -All -ErrorAction Stop)
                Write-Step "Found $($all.Count) deleted users in directory. Searching for '$localPart'..." "Info"

                # 1. Exact UPN match
                $deleted = $all | Where-Object { $_.UserPrincipalName -ieq $UPN } | Select-Object -First 1

                # 2. Local-part exactly before '@'
                if (-not $deleted) {
                    $cands = @($all | Where-Object { ($_.UserPrincipalName -split '@')[0] -ieq $localPart })
                    if ($cands.Count -eq 1) { $deleted = $cands[0] }
                    elseif ($cands.Count -gt 1) {
                        Write-Step "Multiple entries match local-part '$localPart' - specify UPN more precisely." "Error"
                        $cands | ForEach-Object { Write-Step "  - $($_.UserPrincipalName)" "Warning" }
                        return $false
                    }
                }

                # 3. UPN or Mail contains local-part (covers soft-delete prefix like {guid}name@domain)
                if (-not $deleted) {
                    $cands = @($all | Where-Object {
                        ($_.UserPrincipalName -and $_.UserPrincipalName.ToLower().Contains($lpLower)) -or
                        ($_.Mail              -and $_.Mail.ToLower().Contains($lpLower))
                    })
                    if ($cands.Count -eq 1) {
                        $deleted = $cands[0]
                        Write-Step "Found '$($deleted.UserPrincipalName)' via fuzzy match." "Info"
                    } elseif ($cands.Count -gt 1) {
                        Write-Step "Multiple possible deleted users match '$localPart':" "Error"
                        $cands | ForEach-Object { Write-Step "  - UPN=$($_.UserPrincipalName)  Mail=$($_.Mail)  Id=$($_.Id)" "Warning" }
                        return $false
                    }
                }
            } catch { $deleted = $null; Write-Step "Error fetching deleted items: $_" "Warning" }
            if ($deleted) { break }
            if ($i -lt 4) {
                Write-Step "Deleted user not found yet. Running new delta sync and waiting 30s (attempt $($i+1)/5)..." "Warning"
                try {
                    Invoke-Command -Session $SyncSession -ScriptBlock {
                        Import-Module ADSync -ErrorAction Stop
                        $deadline = (Get-Date).AddSeconds(60)
                        while ((Get-ADSyncConnectorRunStatus) -and (Get-Date) -lt $deadline) { Start-Sleep 5 }
                        try { Start-ADSyncSyncCycle -PolicyType Delta | Out-Null } catch {}
                    }
                } catch {}
                Start-Sleep 30
            }
        }
        if (-not $deleted) {
            Write-Step "No deleted user found for UPN '$UPN'. Verify the sync ran." "Error"
            return $false
        }

        $restoredId = $deleted.Id
        Write-Step "Matched deleted: Id=$restoredId  UPN=$($deleted.UserPrincipalName)" "Info"

        try {
            Restore-MgDirectoryDeletedItem -DirectoryObjectId $restoredId -ErrorAction Stop | Out-Null
            Write-Step "User restored. Waiting $($Settings.RestoreWaitSeconds)s..." "Success"
            Start-Sleep -Seconds $Settings.RestoreWaitSeconds
        } catch {
            Write-Step "Failed to restore user: $_" "Error"
            return $false
        }
    } else {
        Write-Step "[DRY RUN] Would restore deleted Entra user for '$UPN'." "Warning"
    }

    # --- Step 5: Clear ImmutableId ---
    Write-Step "[5/6] Clearing onPremisesImmutableId..."
    if (-not $DryRun) {
        $props = "id,userPrincipalName,onPremisesImmutableId,onPremisesSyncEnabled"
        # Use object Id (stable), not UPN - UPN changes on restore
        $mgUser = $null
        for ($r = 0; $r -lt 6; $r++) {
            try {
                $mgUser = Get-MgUser -UserId $restoredId -Property $props -ErrorAction Stop
                break
            } catch {
                if ($r -lt 5) {
                    Write-Step "Entra has not indexed the restore yet, waiting 10s (attempt $($r+1)/6)..." "Warning"
                    Start-Sleep 10
                } else {
                    Write-Step "Could not retrieve Entra user (Id=$restoredId): $_" "Error"
                    return $false
                }
            }
        }
        $UPN = $mgUser.UserPrincipalName
        Write-Step "Restored UPN: $UPN" "Info"

        if ($mgUser.onPremisesSyncEnabled -eq $true) {
            Write-Step "onPremisesSyncEnabled=True - account still appears to be synced. Aborting." "Error"
            return $false
        }

        $managedSuffix = $Settings.ManagedDomainSuffix
        $currentDomain = ($mgUser.UserPrincipalName -split "@")[1]
        $localPart     = ($mgUser.UserPrincipalName -split "@")[0]
        $tempUPN       = "$localPart@$managedSuffix"
        $movedUPN      = $false

        if ($currentDomain -ne $managedSuffix) {
            Write-Step "Temporarily changing UPN to $tempUPN (federated-domain workaround)..."
            try {
                Update-MgUser -UserId $mgUser.Id -UserPrincipalName $tempUPN -ErrorAction Stop
                Start-Sleep 5
                $mgUser = Get-MgUser -UserId $tempUPN -Property $props -ErrorAction Stop
                $movedUPN = $true
            } catch {
                Write-Step "Failed to change UPN: $_" "Error"
                return $false
            }
        }

        $uri  = "https://graph.microsoft.com/v1.0/users/$($mgUser.Id)"
        $body = @{ onPremisesImmutableId = $null }
        try {
            Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body -ErrorAction Stop | Out-Null
        } catch {
            Write-Step "Failed to clear ImmutableId: $_" "Error"
            return $false
        }

        $updated = Get-MgUser -UserId $mgUser.Id -Property "id,userPrincipalName,onPremisesImmutableId"
        if ($null -ne $updated.onPremisesImmutableId) {
            Write-Step "ImmutableId does not appear to have been cleared. Check the account manually." "Error"
            return $false
        }
        Write-Step "ImmutableId cleared." "Success"

        if ($movedUPN) {
            Write-Step "Restoring UPN to $UPN..."
            try {
                Update-MgUser -UserId $mgUser.Id -UserPrincipalName $UPN -ErrorAction Stop
                Write-Step "UPN restored." "Success"
            } catch {
                Write-Step "Warning: Failed to restore UPN. Check manually." "Warning"
            }
        }
    } else {
        Write-Step "[DRY RUN] Would clear ImmutableId for '$UPN'." "Warning"
    }

    # --- Step 6: Re-enable sync ---
    Write-Step "[6/6] Re-enabling the sync schedule and running delta sync..."
    if (-not $DryRun) {
        try {
            Invoke-Command -Session $SyncSession -ScriptBlock {
                Import-Module ADSync -ErrorAction Stop
                Set-ADSyncScheduler -SyncCycleEnabled $true
                Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
            }
            Write-Step "Sync schedule re-enabled." "Success"
        } catch {
            Write-Step "Warning: Failed to restart sync schedule. Start it manually on $($Settings.SyncServer)." "Warning"
        }
    } else {
        Write-Step "[DRY RUN] Would re-enable the sync schedule and run delta sync." "Warning"
    }

    Write-Step "-- Conversion complete for $UPN --" "Success"
    return $true
}

# -------------------------------------------------------------
#  XAML UI
# -------------------------------------------------------------
[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Convert AD User to Cloud-Only (Entra ID)"
    Width="820" Height="720" MinWidth="700" MinHeight="580"
    WindowStartupLocation="CenterScreen"
    Background="#1E1E2E">
  <Window.Resources>
    <Style TargetType="Label">
      <Setter Property="Foreground" Value="#CDD6F4"/>
      <Setter Property="FontSize" Value="13"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#313244"/>
      <Setter Property="Foreground" Value="#CDD6F4"/>
      <Setter Property="BorderBrush" Value="#45475A"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="4,3"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="CaretBrush" Value="#CDD6F4"/>
    </Style>
    <Style TargetType="Button">
      <Setter Property="Background" Value="#89B4FA"/>
      <Setter Property="Foreground" Value="#1E1E2E"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="5" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#B4BEFE"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Background" Value="#45475A"/>
          <Setter Property="Foreground" Value="#6C7086"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#CDD6F4"/>
      <Setter Property="FontSize" Value="13"/>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Background" Value="#313244"/>
      <Setter Property="Foreground" Value="#CDD6F4"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Padding" Value="14,6"/>
    </Style>
    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="#313244"/>
      <Setter Property="Foreground" Value="#CDD6F4"/>
      <Setter Property="BorderBrush" Value="#45475A"/>
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#45475A"/>
      <Setter Property="AlternatingRowBackground" Value="#2A2A3E"/>
      <Setter Property="RowBackground" Value="#313244"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>
  </Window.Resources>

  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="160"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
      <TextBlock Text="Convert AD User &#x2192; Entra ID Cloud-Only" FontSize="18" FontWeight="Bold"
                 Foreground="#CDD6F4" VerticalAlignment="Center"/>
      <TextBlock Text=" (Microsoft Graph)" FontSize="13" Foreground="#6C7086" VerticalAlignment="Bottom" Margin="6,0,0,2"/>
    </StackPanel>

    <!-- Tab control -->
    <TabControl Grid.Row="1" Background="#181825" BorderBrush="#45475A" x:Name="MainTabs">

      <!-- Tab: Single user -->
      <TabItem Header="Single user">
        <Grid Margin="12">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Grid Grid.Row="0" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="160"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Label Grid.Row="0" Grid.Column="0" Content="sAMAccountName (AD):"/>
            <TextBox Grid.Row="0" Grid.Column="1" x:Name="TxtSam" Margin="0,0,0,6"
                     ToolTip="The account's sAMAccountName in Active Directory (logon name without domain)"/>
            <Label Grid.Row="1" Grid.Column="0" Content="UPN (Entra ID):"/>
            <TextBox Grid.Row="1" Grid.Column="1" x:Name="TxtUpn"
                     ToolTip="The user's UserPrincipalName in Entra ID, e.g. john.doe@contoso.com"/>
          </Grid>

          <Separator Grid.Row="1" Background="#45475A" Margin="0,0,0,10"/>

          <CheckBox Grid.Row="2" x:Name="ChkDryRunSingle" Content="Dry run (simulate without making changes)"
                    IsChecked="True" Margin="0,0,0,10" Foreground="#F9E2AF"/>

          <StackPanel Grid.Row="3" Orientation="Horizontal">
            <Button x:Name="BtnRunSingle" Content="Run conversion" Background="#A6E3A1" Foreground="#1E1E2E" Padding="14,8"/>
            <TextBlock x:Name="TxtSingleStatus" VerticalAlignment="Center" Margin="14,0,0,0"
                       Foreground="#6C7086" FontSize="12"/>
          </StackPanel>
        </Grid>
      </TabItem>

      <!-- Tab: Bulk CSV -->
      <TabItem Header="Bulk (CSV file)">
        <Grid Margin="12">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
            <Button x:Name="BtnBrowseCsv" Content="Choose CSV file..." Padding="10,6"/>
            <TextBox x:Name="TxtCsvPath" Width="380" Margin="8,0,0,0" IsReadOnly="True" VerticalContentAlignment="Center"/>
            <Button x:Name="BtnLoadCsv" Content="Load" Margin="8,0,0,0" Padding="10,6"/>
          </StackPanel>

          <TextBlock Grid.Row="1" Text="CSV format: sAMAccountName,UPN  (with header row)"
                     Foreground="#6C7086" FontSize="11" Margin="0,0,0,8"/>

          <DataGrid Grid.Row="2" x:Name="BulkGrid" AutoGenerateColumns="False"
                    CanUserAddRows="False" CanUserDeleteRows="False"
                    SelectionMode="Extended" IsReadOnly="True" Margin="0,0,0,8">
            <DataGrid.Columns>
              <DataGridTextColumn Header="sAMAccountName" Binding="{Binding SamAccountName}" Width="*"/>
              <DataGridTextColumn Header="UPN" Binding="{Binding UPN}" Width="2*"/>
              <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="120">
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock">
                    <Style.Triggers>
                      <Trigger Property="Text" Value="OK">
                        <Setter Property="Foreground" Value="#A6E3A1"/>
                      </Trigger>
                      <Trigger Property="Text" Value="FAIL">
                        <Setter Property="Foreground" Value="#F38BA8"/>
                      </Trigger>
                      <Trigger Property="Text" Value="Running...">
                        <Setter Property="Foreground" Value="#F9E2AF"/>
                      </Trigger>
                    </Style.Triggers>
                  </Style>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
            </DataGrid.Columns>
          </DataGrid>

          <StackPanel Grid.Row="3" Orientation="Horizontal">
            <CheckBox x:Name="ChkDryRunBulk" Content="Dry run" IsChecked="True" Foreground="#F9E2AF" VerticalAlignment="Center"/>
            <Button x:Name="BtnRunBulk" Content="Run all" Background="#A6E3A1" Foreground="#1E1E2E"
                    Margin="16,0,0,0" Padding="14,8" IsEnabled="False"/>
            <TextBlock x:Name="TxtBulkStatus" VerticalAlignment="Center" Margin="14,0,0,0"
                       Foreground="#6C7086" FontSize="12"/>
          </StackPanel>
        </Grid>
      </TabItem>

      <!-- Tab: Settings -->
      <TabItem Header="Settings">
        <ScrollViewer Margin="12" VerticalScrollBarVisibility="Auto">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="220"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Label Grid.Row="0" Grid.Column="0" Content="Target OU (AD):"/>
            <TextBox Grid.Row="0" Grid.Column="1" x:Name="TxtTargetOU" Margin="0,0,0,8"
                     ToolTip="DN of the OU to move users to (a non-synced OU)"/>
            <Label Grid.Row="1" Grid.Column="0" Content="AAD Connect server:"/>
            <TextBox Grid.Row="1" Grid.Column="1" x:Name="TxtSyncServer" Margin="0,0,0,8"
                     ToolTip="Hostname or FQDN of the Azure AD Connect server"/>
            <Label Grid.Row="2" Grid.Column="0" Content="Managed domain suffix:"/>
            <TextBox Grid.Row="2" Grid.Column="1" x:Name="TxtManagedDomain" Margin="0,0,0,8"
                     ToolTip="Your onmicrosoft.com domain (non-federated), e.g. contoso.onmicrosoft.com"/>
            <Label Grid.Row="3" Grid.Column="0" Content="Log folder:"/>
            <TextBox Grid.Row="3" Grid.Column="1" x:Name="TxtLogFolder" Margin="0,0,0,8"
                     ToolTip="Local folder where XML exports and logs are saved"/>
            <Label Grid.Row="4" Grid.Column="0" Content="Sync wait time (seconds):"/>
            <TextBox Grid.Row="4" Grid.Column="1" x:Name="TxtSyncWait" Margin="0,0,0,8" Width="80"
                     HorizontalAlignment="Left" ToolTip="Time to wait after delta sync (default: 180)"/>
            <Label Grid.Row="5" Grid.Column="0" Content="Restore wait time (sec):"/>
            <TextBox Grid.Row="5" Grid.Column="1" x:Name="TxtRestoreWait" Margin="0,0,0,8" Width="80"
                     HorizontalAlignment="Left" ToolTip="Time to wait after restore (default: 20)"/>
          </Grid>
        </ScrollViewer>
      </TabItem>

    </TabControl>

    <!-- Save button -->
    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,8" HorizontalAlignment="Right">
      <Button x:Name="BtnSaveSettings" Content="Save settings" Padding="12,6"/>
    </StackPanel>

    <!-- Log -->
    <Border Grid.Row="3" BorderBrush="#45475A" BorderThickness="1" CornerRadius="4">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" Background="#181825" Margin="6,4">
          <TextBlock Text="Log" Foreground="#6C7086" FontSize="11" FontWeight="Bold" VerticalAlignment="Center"/>
          <Button x:Name="BtnClearLog" Content="Clear" Background="#45475A" Foreground="#CDD6F4"
                  Padding="6,2" Margin="10,0,0,0" FontSize="11"/>
        </StackPanel>
        <RichTextBox Grid.Row="1" x:Name="LogBox" Background="#11111B" BorderThickness="0"
                     IsReadOnly="True" VerticalScrollBarVisibility="Auto"
                     FontFamily="Consolas" FontSize="12" Padding="6,4"/>
      </Grid>
    </Border>

  </Grid>
</Window>
'@

# -------------------------------------------------------------
#  Build UI
# -------------------------------------------------------------
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

function G ($name) { $window.FindName($name) }

$TxtSam           = G "TxtSam"
$TxtUpn           = G "TxtUpn"
$ChkDryRunSingle  = G "ChkDryRunSingle"
$BtnRunSingle     = G "BtnRunSingle"
$TxtSingleStatus  = G "TxtSingleStatus"
$BtnBrowseCsv     = G "BtnBrowseCsv"
$TxtCsvPath       = G "TxtCsvPath"
$BtnLoadCsv       = G "BtnLoadCsv"
$BulkGrid         = G "BulkGrid"
$ChkDryRunBulk    = G "ChkDryRunBulk"
$BtnRunBulk       = G "BtnRunBulk"
$TxtBulkStatus    = G "TxtBulkStatus"
$TxtTargetOU      = G "TxtTargetOU"
$TxtSyncServer    = G "TxtSyncServer"
$TxtManagedDomain = G "TxtManagedDomain"
$TxtLogFolder     = G "TxtLogFolder"
$TxtSyncWait      = G "TxtSyncWait"
$TxtRestoreWait   = G "TxtRestoreWait"
$BtnSaveSettings  = G "BtnSaveSettings"
$BtnClearLog      = G "BtnClearLog"
$LogBox           = G "LogBox"

# -------------------------------------------------------------
#  Log function (UI thread)
# -------------------------------------------------------------
$ColorMap = @{ Info='#CDD6F4'; Success='#A6E3A1'; Warning='#F9E2AF'; Error='#F38BA8' }

function Write-Log ([string]$msg, [string]$level = "Info") {
    $color  = $ColorMap[$level]
    $prefix = switch ($level) {
        "Success" { "[OK] " } "Warning" { "[!!] " } "Error" { "[ERR]" } default { "[   ]" }
    }
    $ts = Get-Date -Format "HH:mm:ss"
    $window.Dispatcher.Invoke([action]{
        $para = [System.Windows.Documents.Paragraph]::new()
        $run  = [System.Windows.Documents.Run]::new("$ts $prefix $msg")
        $run.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString($color)
        $para.Inlines.Add($run)
        $para.Margin = [Windows.Thickness]::new(0)
        $LogBox.Document.Blocks.Add($para)
        $LogBox.ScrollToEnd()
    })
}

# -------------------------------------------------------------
#  Load settings into UI
# -------------------------------------------------------------
$settings = Load-Settings
$TxtTargetOU.Text      = $settings.TargetOU
$TxtSyncServer.Text    = $settings.SyncServer
$TxtManagedDomain.Text = $settings.ManagedDomainSuffix
$TxtLogFolder.Text     = $settings.LogFolder
$TxtSyncWait.Text      = [string]$settings.SyncWaitSeconds
$TxtRestoreWait.Text   = [string]$settings.RestoreWaitSeconds

function Get-UISettings {
    return [PSCustomObject]@{
        TargetOU            = $TxtTargetOU.Text.Trim()
        SyncServer          = $TxtSyncServer.Text.Trim()
        ManagedDomainSuffix = $TxtManagedDomain.Text.Trim()
        LogFolder           = $TxtLogFolder.Text.Trim()
        SyncWaitSeconds     = [int]($TxtSyncWait.Text -replace '\D','')
        RestoreWaitSeconds  = [int]($TxtRestoreWait.Text -replace '\D','')
    }
}

# -------------------------------------------------------------
#  Event: Save settings
# -------------------------------------------------------------
$BtnSaveSettings.Add_Click({
    Save-Settings (Get-UISettings)
    Write-Log "Settings saved to $SettingsFile." "Success"
})

# -------------------------------------------------------------
#  Event: Clear log
# -------------------------------------------------------------
$BtnClearLog.Add_Click({
    $LogBox.Document.Blocks.Clear()
})

# -------------------------------------------------------------
#  Event: Run single user
# -------------------------------------------------------------
$BtnRunSingle.Add_Click({
    $sam = $TxtSam.Text.Trim()
    $upn = $TxtUpn.Text.Trim()
    $dry = [bool]$ChkDryRunSingle.IsChecked

    if (-not $sam -or -not $upn) { Write-Log "Enter sAMAccountName and UPN." "Error"; return }
    $cfg = Get-UISettings

    $syncCred = $null
    if (-not $dry) {
        $syncCred = Get-Credential -Message "Authentication for $($cfg.SyncServer)"
        if (-not $syncCred) { Write-Log "Cancelled." "Warning"; return }
    }

    $BtnRunSingle.IsEnabled = $false
    $TxtSingleStatus.Text   = "Running..."
    $TxtSingleStatus.Foreground = [Windows.Media.Brushes]::Orange

    $adModPath = (Get-Module -ListAvailable -Name ActiveDirectory | Select-Object -First 1).Path
    $_iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $_iss.ImportPSModule(@('ActiveDirectory','Microsoft.Graph.Authentication','Microsoft.Graph.Users','Microsoft.Graph.Identity.DirectoryManagement'))
    $script:_singleRs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($_iss)
    $script:_singleRs.ApartmentState = 'STA'
    $script:_singleRs.ThreadOptions  = 'ReuseThread'
    $script:_singleRs.Open()
    $script:_singleRs.SessionStateProxy.SetVariable('PSModulePath', $env:PSModulePath)
    $script:_singleRs.SessionStateProxy.SetVariable('_adModPath',   $adModPath)
    $script:_singleRs.SessionStateProxy.SetVariable('_dispatcher', $window.Dispatcher)
    $script:_singleRs.SessionStateProxy.SetVariable('_logBox',     $LogBox)
    $script:_singleRs.SessionStateProxy.SetVariable('_convertFn',  [scriptblock]::Create((Get-Item Function:\Convert-UserToCloudOnly).ScriptBlock.ToString()))
    $script:_singleRs.SessionStateProxy.SetVariable('_sam',        $sam)
    $script:_singleRs.SessionStateProxy.SetVariable('_upn',        $upn)
    $script:_singleRs.SessionStateProxy.SetVariable('_cfg',        $cfg)
    $script:_singleRs.SessionStateProxy.SetVariable('_dry',        $dry)
    $script:_singleRs.SessionStateProxy.SetVariable('_syncCred',   $syncCred)

    $script:_singlePs = [PowerShell]::Create()
    $script:_singlePs.Runspace = $script:_singleRs
    $script:_singlePs.AddScript({
        $env:PSModulePath = $PSModulePath

        function Write-Log ([string]$msg, [string]$level = 'Info') {
            $cm  = @{ Info='#CDD6F4'; Success='#A6E3A1'; Warning='#F9E2AF'; Error='#F38BA8' }
            $col = $cm[$level]
            $pfx = switch ($level) { 'Success'{'[OK] '} 'Warning'{'[!!] '} 'Error'{'[ERR]'} default{'[   ]'} }
            $ln  = "$(Get-Date -Format 'HH:mm:ss') $pfx $msg"
            $c   = $col
            $_dispatcher.Invoke([action]{
                $p = [System.Windows.Documents.Paragraph]::new()
                $r = [System.Windows.Documents.Run]::new($ln)
                $r.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString($c)
                $p.Inlines.Add($r)
                $p.Margin = [Windows.Thickness]::new(0)
                $_logBox.Document.Blocks.Add($p)
                $_logBox.ScrollToEnd()
            })
        }

        Set-Item -Path Function:\Convert-UserToCloudOnly -Value $_convertFn

        Write-Log "Preparing..." "Info"

        if (-not $_adModPath) {
            Write-Log "ActiveDirectory module not found. Install RSAT: AD DS Tools and restart the tool." "Error"; return $false
        }
        try { Import-Module $_adModPath -ErrorAction Stop }
        catch { Write-Log "Could not load ActiveDirectory: $_" "Error"; return $false }
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            Write-Log "Microsoft.Graph module missing. Installing (requires internet access)..." "Warning"
            try { Install-Module Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop }
            catch { Write-Log "Installation failed: $_. Run: Install-Module Microsoft.Graph" "Error"; return $false }
        }
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Import-Module Microsoft.Graph.Users -ErrorAction Stop
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
        }
        catch { Write-Log "Could not load Microsoft.Graph modules: $_" "Error"; return $false }

        $session = $null
        if (-not $_dry -and $_syncCred) {
            try {
                $session = New-PSSession -ComputerName $_cfg.SyncServer -Credential $_syncCred -ErrorAction Stop
                Write-Log "Connected to $($_cfg.SyncServer)." "Success"
            } catch {
                Write-Log "Failed to connect to $($_cfg.SyncServer): $_" "Error"
                return $false
            }
        }

        $result = Convert-UserToCloudOnly `
            -SamAccountName $_sam -UPN $_upn `
            -Settings       $_cfg `
            -SyncCred       $_syncCred `
            -SyncSession    $session `
            -DryRun         $_dry `
            -Log            ${function:Write-Log}

        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
        return $result
    }) | Out-Null

    $script:_singleAsync = $script:_singlePs.BeginInvoke()

    $script:_singleTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:_singleTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $script:_singleTimer.Add_Tick({
        if (-not $script:_singleAsync.IsCompleted) { return }
        $script:_singleTimer.Stop()
        $res = $null
        try { $res = $script:_singlePs.EndInvoke($script:_singleAsync) } catch {}
        $script:_singlePs.Dispose()
        $script:_singleRs.Dispose()
        $ok = $res -contains $true
        $BtnRunSingle.IsEnabled = $true
        if ($ok) {
            $TxtSingleStatus.Text       = "Done!"
            $TxtSingleStatus.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString("#A6E3A1")
        } else {
            $TxtSingleStatus.Text       = "Error - see log"
            $TxtSingleStatus.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString("#F38BA8")
        }
    })
    $script:_singleTimer.Start()
})

# -------------------------------------------------------------
#  Event: Browse CSV
# -------------------------------------------------------------
$BtnBrowseCsv.Add_Click({
    $ofd = [System.Windows.Forms.OpenFileDialog]::new()
    $ofd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $ofd.Title  = "Choose CSV file with users"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TxtCsvPath.Text = $ofd.FileName
    }
})

# -------------------------------------------------------------
#  Bulk data model
# -------------------------------------------------------------
class BulkUser {
    [string]$SamAccountName
    [string]$UPN
    [string]$Status
}

$script:BulkUsers = [System.Collections.ObjectModel.ObservableCollection[BulkUser]]::new()
$BulkGrid.ItemsSource = $script:BulkUsers

# -------------------------------------------------------------
#  Event: Load CSV
# -------------------------------------------------------------
$BtnLoadCsv.Add_Click({
    $path = $TxtCsvPath.Text.Trim()
    if (-not (Test-Path $path)) { Write-Log "CSV file not found: $path" "Error"; return }

    $script:BulkUsers.Clear()
    try {
        $rows = Import-Csv -Path $path -ErrorAction Stop
        foreach ($row in $rows) {
            $sam = if ($row.PSObject.Properties['sAMAccountName']) { $row.sAMAccountName }
                   elseif ($row.PSObject.Properties['SamAccountName']) { $row.SamAccountName }
                   else { $null }
            $upn = if ($row.PSObject.Properties['UPN']) { $row.UPN }
                   elseif ($row.PSObject.Properties['UserPrincipalName']) { $row.UserPrincipalName }
                   else { $null }
            if (-not $sam -or -not $upn) { Write-Log "Row missing sAMAccountName/UPN - skipping." "Warning"; continue }
            $u = [BulkUser]::new(); $u.SamAccountName = $sam; $u.UPN = $upn; $u.Status = "-"
            $script:BulkUsers.Add($u)
        }
        Write-Log "Loaded $($script:BulkUsers.Count) users from CSV." "Success"
        $BtnRunBulk.IsEnabled = $script:BulkUsers.Count -gt 0
    } catch {
        Write-Log "Failed to read CSV: $_" "Error"
    }
})

# -------------------------------------------------------------
#  Event: Run bulk
# -------------------------------------------------------------
$BtnRunBulk.Add_Click({
    if ($script:BulkUsers.Count -eq 0) { return }

    $dry = [bool]$ChkDryRunBulk.IsChecked
    $cfg = Get-UISettings

    $syncCred = $null
    if (-not $dry) {
        $syncCred = Get-Credential -Message "Authentication for $($cfg.SyncServer)"
        if (-not $syncCred) { Write-Log "Cancelled." "Warning"; return }
    }

    $BtnRunBulk.IsEnabled = $false
    $TxtBulkStatus.Text   = "Running..."

    $userList  = @($script:BulkUsers | ForEach-Object { [PSCustomObject]@{ Sam=$_.SamAccountName; UPN=$_.UPN } })
    $bulkUsers = $script:BulkUsers

    $adModPath = (Get-Module -ListAvailable -Name ActiveDirectory | Select-Object -First 1).Path
    $_iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $_iss.ImportPSModule(@('ActiveDirectory','Microsoft.Graph.Authentication','Microsoft.Graph.Users','Microsoft.Graph.Identity.DirectoryManagement'))
    $script:_bulkRs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($_iss)
    $script:_bulkRs.ApartmentState = 'STA'
    $script:_bulkRs.ThreadOptions  = 'ReuseThread'
    $script:_bulkRs.Open()
    $script:_bulkRs.SessionStateProxy.SetVariable('PSModulePath', $env:PSModulePath)
    $script:_bulkRs.SessionStateProxy.SetVariable('_adModPath',   $adModPath)
    $script:_bulkRs.SessionStateProxy.SetVariable('_dispatcher', $window.Dispatcher)
    $script:_bulkRs.SessionStateProxy.SetVariable('_logBox',     $LogBox)
    $script:_bulkRs.SessionStateProxy.SetVariable('_bulkGrid',   $BulkGrid)
    $script:_bulkRs.SessionStateProxy.SetVariable('_convertFn',  [scriptblock]::Create((Get-Item Function:\Convert-UserToCloudOnly).ScriptBlock.ToString()))
    $script:_bulkRs.SessionStateProxy.SetVariable('_userList',   $userList)
    $script:_bulkRs.SessionStateProxy.SetVariable('_bulkUsers',  $bulkUsers)
    $script:_bulkRs.SessionStateProxy.SetVariable('_cfg',        $cfg)
    $script:_bulkRs.SessionStateProxy.SetVariable('_dry',        $dry)
    $script:_bulkRs.SessionStateProxy.SetVariable('_syncCred',   $syncCred)

    $script:_bulkPs = [PowerShell]::Create()
    $script:_bulkPs.Runspace = $script:_bulkRs
    $script:_bulkPs.AddScript({
        $env:PSModulePath = $PSModulePath

        function Write-Log ([string]$msg, [string]$level = 'Info') {
            $cm  = @{ Info='#CDD6F4'; Success='#A6E3A1'; Warning='#F9E2AF'; Error='#F38BA8' }
            $col = $cm[$level]
            $pfx = switch ($level) { 'Success'{'[OK] '} 'Warning'{'[!!] '} 'Error'{'[ERR]'} default{'[   ]'} }
            $ln  = "$(Get-Date -Format 'HH:mm:ss') $pfx $msg"
            $c   = $col
            $_dispatcher.Invoke([action]{
                $p = [System.Windows.Documents.Paragraph]::new()
                $r = [System.Windows.Documents.Run]::new($ln)
                $r.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString($c)
                $p.Inlines.Add($r)
                $p.Margin = [Windows.Thickness]::new(0)
                $_logBox.Document.Blocks.Add($p)
                $_logBox.ScrollToEnd()
            })
        }

        Set-Item -Path Function:\Convert-UserToCloudOnly -Value $_convertFn

        if (-not $_adModPath) {
            Write-Log "ActiveDirectory module not found. Install RSAT: AD DS Tools and restart the tool." "Error"; return
        }
        try { Import-Module $_adModPath -ErrorAction Stop }
        catch { Write-Log "Could not load ActiveDirectory: $_" "Error"; return }
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            Write-Log "Microsoft.Graph module missing. Installing (requires internet access)..." "Warning"
            try { Install-Module Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop }
            catch { Write-Log "Installation failed: $_. Run: Install-Module Microsoft.Graph" "Error"; return }
        }
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Import-Module Microsoft.Graph.Users -ErrorAction Stop
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
        }
        catch { Write-Log "Could not load Microsoft.Graph modules: $_" "Error"; return }

        $session = $null
        if (-not $_dry -and $_syncCred) {
            try {
                $session = New-PSSession -ComputerName $_cfg.SyncServer -Credential $_syncCred -ErrorAction Stop
                Write-Log "Connected to $($_cfg.SyncServer)." "Success"
            } catch {
                Write-Log "Failed to connect to $($_cfg.SyncServer): $_" "Error"
                return
            }
        }

        for ($i = 0; $i -lt $_userList.Count; $i++) {
            $u   = $_userList[$i]
            $idx = $i
            $_dispatcher.Invoke([action]{
                $_bulkUsers[$idx].Status = 'Running...'
                $_bulkGrid.Items.Refresh()
            })

            $ok = Convert-UserToCloudOnly `
                -SamAccountName $u.Sam -UPN $u.UPN `
                -Settings       $_cfg `
                -SyncCred       $_syncCred `
                -SyncSession    $session `
                -DryRun         $_dry `
                -Log            ${function:Write-Log}

            $st = if ($ok) { 'OK' } else { 'FAIL' }
            $_dispatcher.Invoke([action]{
                $_bulkUsers[$idx].Status = $st
                $_bulkGrid.Items.Refresh()
            })
        }

        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
    }) | Out-Null

    $script:_bulkAsync = $script:_bulkPs.BeginInvoke()

    $script:_bulkTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:_bulkTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $script:_bulkTimer.Add_Tick({
        if (-not $script:_bulkAsync.IsCompleted) { return }
        $script:_bulkTimer.Stop()
        try { $script:_bulkPs.EndInvoke($script:_bulkAsync) } catch {}
        $script:_bulkPs.Dispose()
        $script:_bulkRs.Dispose()
        $failed = ($script:BulkUsers | Where-Object { $_.Status -eq 'FAIL' }).Count
        $total  = $script:BulkUsers.Count
        $BtnRunBulk.IsEnabled = $true
        $TxtBulkStatus.Text   = "Done: $($total - $failed)/$total succeeded"
        Write-Log "Bulk complete. $($total - $failed) of $total succeeded." $(if ($failed -eq 0) {'Success'} else {'Warning'})
    })
    $script:_bulkTimer.Start()
})

# -------------------------------------------------------------
#  Clean up on close
# -------------------------------------------------------------
$window.Add_Closed({
    if ($script:SyncSession) { Remove-PSSession $script:SyncSession -ErrorAction SilentlyContinue }
})

# -------------------------------------------------------------
#  Show window
# -------------------------------------------------------------
Write-Log "Welcome! Fill in settings and choose Single user or Bulk." "Info"
Write-Log "Dry run is enabled by default - no changes are made until you uncheck it." "Warning"
[void]$window.ShowDialog()
