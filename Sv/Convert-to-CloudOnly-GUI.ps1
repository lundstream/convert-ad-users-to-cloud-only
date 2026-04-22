<#
.SYNOPSIS
    GUI-verktyg för att konvertera AD-synkade användare till cloud-only-konton i Entra ID.
    Stödjer enskild användare och bulk via CSV-fil.

.DESCRIPTION
    Steg per användare:
      1. Flytta AD-kontot till ett icke-synkat OU
      2. Kör en delta-synk → kontot tas bort i Entra ID
      3. Pausa synkschema
      4. Återställ kontot i Entra ID
      5. Rensa onPremisesImmutableId via Graph API (hanterar federerat UPN om nödvändigt)
      6. Återaktivera synkschema och kör delta-synk

.REQUIREMENTS
    - Kör som administratör
    - RSAT: Active Directory DS Tools installerat
    - Microsoft.Graph PowerShell-modul
    - Nätverksåtkomst (WinRM) till Azure AD Connect-servern
    - Graph-behörighet: User.ReadWrite.All

.NOTES
    Inställningar sparas i Convert-to-CloudOnly-Settings.json bredvid skriptet.
#>

#Requires -Version 5.1
Set-StrictMode -Off

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ─────────────────────────────────────────────────────────────
#  Inställningshjälpare
# ─────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────
#  Konverteringslogik (kärna)
# ─────────────────────────────────────────────────────────────
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

    Write-Step "── Startar konvertering av $UPN ──"

    # --- Steg 1: Flytta AD-användare ---
    Write-Step "[1/6] Hämtar och flyttar AD-användare '$SamAccountName' till $($Settings.TargetOU)..."
    if (-not $DryRun) {
        try {
            $adUser = Get-ADUser -Identity $SamAccountName -Properties *
        } catch {
            Write-Step "Kunde inte hitta AD-användare '$SamAccountName': $_" "Error"
            return $false
        }

        if (-not (Test-Path $Settings.LogFolder)) { New-Item $Settings.LogFolder -ItemType Directory -Force | Out-Null }
        $exportPath = Join-Path $Settings.LogFolder ("$SamAccountName`_AD_pre-move_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml")
        try { $adUser | Export-Clixml -Path $exportPath } catch {}

        try {
            Move-ADObject -Identity $adUser.DistinguishedName -TargetPath $Settings.TargetOU
            Write-Step "AD-användaren flyttad." "Success"
        } catch {
            Write-Step "Misslyckades att flytta AD-användare: $_" "Error"
            return $false
        }
    } else {
        Write-Step "[DRY RUN] Skulle flytta '$SamAccountName' till $($Settings.TargetOU)." "Warning"
    }

    # --- Steg 2 & 3: Synk → soft-delete, inaktivera schema ---
    Write-Step "[2/6] Kör delta-synk och inaktiverar synkschema..."
    if (-not $DryRun) {
        try {
            Invoke-Command -Session $SyncSession -ScriptBlock {
                Import-Module ADSync -ErrorAction Stop
                Set-ADSyncScheduler -SyncCycleEnabled $false
                # Vänta ut ev. pågående cykel innan vi startar vår
                $deadline = (Get-Date).AddMinutes(5)
                while ((Get-ADSyncConnectorRunStatus) -and (Get-Date) -lt $deadline) { Start-Sleep 5 }
                # Starta vår delta-synk och vänta tills den är klar
                try { Start-ADSyncSyncCycle -PolicyType Delta | Out-Null } catch {
                    Start-Sleep 10
                    Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
                }
                Start-Sleep 5
                $deadline = (Get-Date).AddMinutes(5)
                while ((Get-ADSyncConnectorRunStatus) -and (Get-Date) -lt $deadline) { Start-Sleep 5 }
            }
            Write-Step "Synk körd och schema pausat. Kort buffert ($($Settings.SyncWaitSeconds)s) för Entra-bearbetning..." "Info"
            Start-Sleep -Seconds $Settings.SyncWaitSeconds
        } catch {
            Write-Step "Problem med synkservern: $_" "Error"
            return $false
        }
    } else {
        Write-Step "[DRY RUN] Skulle köra delta-synk och pausa schema." "Warning"
    }

    # --- Steg 4: Anslut Graph och återställ ---
    Write-Step "[3/6] Ansluter till Microsoft Graph..."
    if (-not $DryRun) {
        try {
            Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop -NoWelcome
            Write-Step "Ansluten till Graph." "Success"
        } catch {
            Write-Step "Misslyckades att ansluta till Graph: $_" "Error"
            return $false
        }
    }

    Write-Step "[4/6] Letar efter borttagen användare '$UPN' i Entra ID..."
    if (-not $DryRun) {
        $localPart = if ($UPN -match '@') { ($UPN -split '@')[0] } else { $UPN }
        $lpLower   = $localPart.ToLower()
        $deleted = $null
        for ($i = 0; $i -lt 5; $i++) {
            try {
                $all = @(Get-MgDirectoryDeletedItemAsUser -All -ErrorAction Stop)
                Write-Step "Hittade $($all.Count) borttagna användare i directory. Söker efter '$localPart'..." "Info"

                # 1. Exakt UPN-match
                $deleted = $all | Where-Object { $_.UserPrincipalName -ieq $UPN } | Select-Object -First 1

                # 2. Local-part exakt före '@'
                if (-not $deleted) {
                    $cands = @($all | Where-Object { ($_.UserPrincipalName -split '@')[0] -ieq $localPart })
                    if ($cands.Count -eq 1) { $deleted = $cands[0] }
                    elseif ($cands.Count -gt 1) {
                        Write-Step "Flera matchar local-part '$localPart' — specificera UPN tydligare." "Error"
                        $cands | ForEach-Object { Write-Step "  - $($_.UserPrincipalName)" "Warning" }
                        return $false
                    }
                }

                # 3. UPN eller Mail innehåller local-part (täcker soft-delete-prefix som {guid}name@domain)
                if (-not $deleted) {
                    $cands = @($all | Where-Object {
                        ($_.UserPrincipalName -and $_.UserPrincipalName.ToLower().Contains($lpLower)) -or
                        ($_.Mail              -and $_.Mail.ToLower().Contains($lpLower))
                    })
                    if ($cands.Count -eq 1) {
                        $deleted = $cands[0]
                        Write-Step "Hittade '$($deleted.UserPrincipalName)' via fuzzy match." "Info"
                    } elseif ($cands.Count -gt 1) {
                        Write-Step "Flera möjliga borttagna användare matchar '$localPart':" "Error"
                        $cands | ForEach-Object { Write-Step "  - UPN=$($_.UserPrincipalName)  Mail=$($_.Mail)  Id=$($_.Id)" "Warning" }
                        return $false
                    }
                }
            } catch { $deleted = $null; Write-Step "Fel vid hämtning av borttagna: $_" "Warning" }
            if ($deleted) { break }
            if ($i -lt 4) {
                Write-Step "Hittade inte borttagen användare ännu. Kör ny delta-synk och väntar 30s (försök $($i+1)/5)..." "Warning"
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
            Write-Step "Ingen borttagen användare hittades med UPN '$UPN'. Kontrollera att synken körde." "Error"
            return $false
        }

        $restoredId = $deleted.Id
        Write-Step "Matchar borttagen: Id=$restoredId  UPN=$($deleted.UserPrincipalName)" "Info"

        try {
            Restore-MgDirectoryDeletedItem -DirectoryObjectId $restoredId -ErrorAction Stop | Out-Null
            Write-Step "Användare återställd. Väntar $($Settings.RestoreWaitSeconds)s..." "Success"
            Start-Sleep -Seconds $Settings.RestoreWaitSeconds
        } catch {
            Write-Step "Misslyckades att återställa användaren: $_" "Error"
            return $false
        }
    } else {
        Write-Step "[DRY RUN] Skulle återställa borttagen Entra-användare för '$UPN'." "Warning"
    }

    # --- Steg 5: Rensa ImmutableId ---
    Write-Step "[5/6] Rensar onPremisesImmutableId..."
    if (-not $DryRun) {
        $props = "id,userPrincipalName,onPremisesImmutableId,onPremisesSyncEnabled"
        # Använd objekt-Id (stabilt), inte UPN — UPN ändras av restore
        $mgUser = $null
        for ($r = 0; $r -lt 6; $r++) {
            try {
                $mgUser = Get-MgUser -UserId $restoredId -Property $props -ErrorAction Stop
                break
            } catch {
                if ($r -lt 5) {
                    Write-Step "Entra har inte indexerat återställningen ännu, väntar 10s (försök $($r+1)/6)..." "Warning"
                    Start-Sleep 10
                } else {
                    Write-Step "Kunde inte hämta Entra-användare (Id=$restoredId): $_" "Error"
                    return $false
                }
            }
        }
        $UPN = $mgUser.UserPrincipalName
        Write-Step "Återställd UPN: $UPN" "Info"

        if ($mgUser.onPremisesSyncEnabled -eq $true) {
            Write-Step "onPremisesSyncEnabled=True — kontot verkar fortfarande vara synkat. Avbryter." "Error"
            return $false
        }

        $managedSuffix = $Settings.ManagedDomainSuffix
        $currentDomain = ($mgUser.UserPrincipalName -split "@")[1]
        $localPart     = ($mgUser.UserPrincipalName -split "@")[0]
        $tempUPN       = "$localPart@$managedSuffix"
        $movedUPN      = $false

        if ($currentDomain -ne $managedSuffix) {
            Write-Step "Byter temporärt UPN till $tempUPN (federerad domän-workaround)..."
            try {
                Update-MgUser -UserId $mgUser.Id -UserPrincipalName $tempUPN -ErrorAction Stop
                Start-Sleep 5
                $mgUser = Get-MgUser -UserId $tempUPN -Property $props -ErrorAction Stop
                $movedUPN = $true
            } catch {
                Write-Step "Misslyckades att byta UPN: $_" "Error"
                return $false
            }
        }

        $uri  = "https://graph.microsoft.com/v1.0/users/$($mgUser.Id)"
        $body = @{ onPremisesImmutableId = $null }
        try {
            Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body -ErrorAction Stop | Out-Null
        } catch {
            Write-Step "Misslyckades att rensa ImmutableId: $_" "Error"
            return $false
        }

        $updated = Get-MgUser -UserId $mgUser.Id -Property "id,userPrincipalName,onPremisesImmutableId"
        if ($null -ne $updated.onPremisesImmutableId) {
            Write-Step "ImmutableId verkar inte ha rensats. Kontrollera kontot manuellt." "Error"
            return $false
        }
        Write-Step "ImmutableId rensad." "Success"

        if ($movedUPN) {
            Write-Step "Återställer UPN till $UPN..."
            try {
                Update-MgUser -UserId $mgUser.Id -UserPrincipalName $UPN -ErrorAction Stop
                Write-Step "UPN återställt." "Success"
            } catch {
                Write-Step "Varning: Misslyckades att återställa UPN. Kontrollera manuellt." "Warning"
            }
        }
    } else {
        Write-Step "[DRY RUN] Skulle rensa ImmutableId för '$UPN'." "Warning"
    }

    # --- Steg 6: Återaktivera synk ---
    Write-Step "[6/6] Startar synkschema igen och kör delta-synk..."
    if (-not $DryRun) {
        try {
            Invoke-Command -Session $SyncSession -ScriptBlock {
                Import-Module ADSync -ErrorAction Stop
                Set-ADSyncScheduler -SyncCycleEnabled $true
                Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
            }
            Write-Step "Synkschema återaktiverat." "Success"
        } catch {
            Write-Step "Varning: Misslyckades att starta om synkschema. Starta manuellt på $($Settings.SyncServer)." "Warning"
        }
    } else {
        Write-Step "[DRY RUN] Skulle återaktivera synkschema och köra delta-synk." "Warning"
    }

    Write-Step "── Konvertering klar för $UPN ──" "Success"
    return $true
}

# ─────────────────────────────────────────────────────────────
#  XAML UI
# ─────────────────────────────────────────────────────────────
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

      <!-- Tab: Enskild anvandare -->
      <TabItem Header="Enskild användare">
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
                     ToolTip="Kontots sAMAccountName i Active Directory (inloggningsnamn utan domän)"/>
            <Label Grid.Row="1" Grid.Column="0" Content="UPN (Entra ID):"/>
            <TextBox Grid.Row="1" Grid.Column="1" x:Name="TxtUpn"
                     ToolTip="Användarens UserPrincipalName i Entra ID, t.ex. john.doe@contoso.com"/>
          </Grid>

          <Separator Grid.Row="1" Background="#45475A" Margin="0,0,0,10"/>

          <CheckBox Grid.Row="2" x:Name="ChkDryRunSingle" Content="Dry run (simulera utan att göra ändringar)"
                    IsChecked="True" Margin="0,0,0,10" Foreground="#F9E2AF"/>

          <StackPanel Grid.Row="3" Orientation="Horizontal">
            <Button x:Name="BtnRunSingle" Content="Kör konvertering" Background="#A6E3A1" Foreground="#1E1E2E" Padding="14,8"/>
            <TextBlock x:Name="TxtSingleStatus" VerticalAlignment="Center" Margin="14,0,0,0"
                       Foreground="#6C7086" FontSize="12"/>
          </StackPanel>
        </Grid>
      </TabItem>

      <!-- Tab: Bulk CSV -->
      <TabItem Header="Bulk (CSV-fil)">
        <Grid Margin="12">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
            <Button x:Name="BtnBrowseCsv" Content="Välj CSV-fil..." Padding="10,6"/>
            <TextBox x:Name="TxtCsvPath" Width="380" Margin="8,0,0,0" IsReadOnly="True" VerticalContentAlignment="Center"/>
            <Button x:Name="BtnLoadCsv" Content="Ladda" Margin="8,0,0,0" Padding="10,6"/>
          </StackPanel>

          <TextBlock Grid.Row="1" Text="CSV-format: sAMAccountName,UPN  (med rubrikrad)"
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
                      <Trigger Property="Text" Value="FEL">
                        <Setter Property="Foreground" Value="#F38BA8"/>
                      </Trigger>
                      <Trigger Property="Text" Value="Kör...">
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
            <Button x:Name="BtnRunBulk" Content="Kör alla" Background="#A6E3A1" Foreground="#1E1E2E"
                    Margin="16,0,0,0" Padding="14,8" IsEnabled="False"/>
            <TextBlock x:Name="TxtBulkStatus" VerticalAlignment="Center" Margin="14,0,0,0"
                       Foreground="#6C7086" FontSize="12"/>
          </StackPanel>
        </Grid>
      </TabItem>

      <!-- Tab: Installningar -->
      <TabItem Header="Inställningar">
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
            <Label Grid.Row="0" Grid.Column="0" Content="Mål-OU (AD):"/>
            <TextBox Grid.Row="0" Grid.Column="1" x:Name="TxtTargetOU" Margin="0,0,0,8"
                     ToolTip="DN för det OU dit användarna ska flyttas (icke-synkat OU)"/>
            <Label Grid.Row="1" Grid.Column="0" Content="AAD Connect-server:"/>
            <TextBox Grid.Row="1" Grid.Column="1" x:Name="TxtSyncServer" Margin="0,0,0,8"
                     ToolTip="Hostname eller FQDN för Azure AD Connect-servern"/>
            <Label Grid.Row="2" Grid.Column="0" Content="Managed domain-suffix:"/>
            <TextBox Grid.Row="2" Grid.Column="1" x:Name="TxtManagedDomain" Margin="0,0,0,8"
                     ToolTip="Din onmicrosoft.com-domän (icke-federerad), t.ex. contoso.onmicrosoft.com"/>
            <Label Grid.Row="3" Grid.Column="0" Content="Loggmapp:"/>
            <TextBox Grid.Row="3" Grid.Column="1" x:Name="TxtLogFolder" Margin="0,0,0,8"
                     ToolTip="Lokal mapp där XML-exportfiler och logg sparas"/>
            <Label Grid.Row="4" Grid.Column="0" Content="Synkväntetid (sekunder):"/>
            <TextBox Grid.Row="4" Grid.Column="1" x:Name="TxtSyncWait" Margin="0,0,0,8" Width="80"
                     HorizontalAlignment="Left" ToolTip="Tid att vänta efter delta-synk (standard: 180)"/>
            <Label Grid.Row="5" Grid.Column="0" Content="Återställningsväntetid (sek):"/>
            <TextBox Grid.Row="5" Grid.Column="1" x:Name="TxtRestoreWait" Margin="0,0,0,8" Width="80"
                     HorizontalAlignment="Left" ToolTip="Tid att vänta efter restore (standard: 20)"/>
          </Grid>
        </ScrollViewer>
      </TabItem>

    </TabControl>

    <!-- Spara-knapp -->
    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,8" HorizontalAlignment="Right">
      <Button x:Name="BtnSaveSettings" Content="Spara inställningar" Padding="12,6"/>
    </StackPanel>

    <!-- Logg -->
    <Border Grid.Row="3" BorderBrush="#45475A" BorderThickness="1" CornerRadius="4">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" Background="#181825" Margin="6,4">
          <TextBlock Text="Logg" Foreground="#6C7086" FontSize="11" FontWeight="Bold" VerticalAlignment="Center"/>
          <Button x:Name="BtnClearLog" Content="Rensa" Background="#45475A" Foreground="#CDD6F4"
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

# ─────────────────────────────────────────────────────────────
#  Bygg UI
# ─────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────
#  Loggfunktion (UI-tråd)
# ─────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────
#  Ladda inställningar i UI
# ─────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────
#  Händelse: Spara inställningar
# ─────────────────────────────────────────────────────────────
$BtnSaveSettings.Add_Click({
    Save-Settings (Get-UISettings)
    Write-Log "Inställningar sparade till $SettingsFile." "Success"
})

# ─────────────────────────────────────────────────────────────
#  Händelse: Rensa logg
# ─────────────────────────────────────────────────────────────
$BtnClearLog.Add_Click({
    $LogBox.Document.Blocks.Clear()
})

# ─────────────────────────────────────────────────────────────
#  Händelse: Kör enskild användare
# ─────────────────────────────────────────────────────────────
$BtnRunSingle.Add_Click({
    $sam = $TxtSam.Text.Trim()
    $upn = $TxtUpn.Text.Trim()
    $dry = [bool]$ChkDryRunSingle.IsChecked

    if (-not $sam -or -not $upn) { Write-Log "Ange sAMAccountName och UPN." "Error"; return }
    $cfg = Get-UISettings

    $syncCred = $null
    if (-not $dry) {
        $syncCred = Get-Credential -Message "Autentisering för $($cfg.SyncServer)"
        if (-not $syncCred) { Write-Log "Avbrutet." "Warning"; return }
    }

    $BtnRunSingle.IsEnabled = $false
    $TxtSingleStatus.Text   = "Kör..."
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

        Write-Log "Förbereder..." "Info"

        if (-not $_adModPath) {
            Write-Log "ActiveDirectory-modulen hittades inte. Installera RSAT: AD DS Tools och starta om verktyget." "Error"; return $false
        }
        try { Import-Module $_adModPath -ErrorAction Stop }
        catch { Write-Log "Kunde inte ladda ActiveDirectory: $_" "Error"; return $false }
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            Write-Log "Microsoft.Graph-modulen saknas. Installerar (kräver internetåtkomst)..." "Warning"
            try { Install-Module Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop }
            catch { Write-Log "Installationen misslyckades: $_. Kör: Install-Module Microsoft.Graph" "Error"; return $false }
        }
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Import-Module Microsoft.Graph.Users -ErrorAction Stop
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
        }
        catch { Write-Log "Kunde inte ladda Microsoft.Graph-moduler: $_" "Error"; return $false }

        $session = $null
        if (-not $_dry -and $_syncCred) {
            try {
                $session = New-PSSession -ComputerName $_cfg.SyncServer -Credential $_syncCred -ErrorAction Stop
                Write-Log "Ansluten till $($_cfg.SyncServer)." "Success"
            } catch {
                Write-Log "Misslyckades att ansluta till $($_cfg.SyncServer): $_" "Error"
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
            $TxtSingleStatus.Text       = "Klart!"
            $TxtSingleStatus.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString("#A6E3A1")
        } else {
            $TxtSingleStatus.Text       = "Fel — se loggen"
            $TxtSingleStatus.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString("#F38BA8")
        }
    })
    $script:_singleTimer.Start()
})

# ─────────────────────────────────────────────────────────────
#  Händelse: Bläddra CSV
# ─────────────────────────────────────────────────────────────
$BtnBrowseCsv.Add_Click({
    $ofd = [System.Windows.Forms.OpenFileDialog]::new()
    $ofd.Filter = "CSV-filer (*.csv)|*.csv|Alla filer (*.*)|*.*"
    $ofd.Title  = "Välj CSV-fil med användare"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TxtCsvPath.Text = $ofd.FileName
    }
})

# ─────────────────────────────────────────────────────────────
#  Bulk-datamodell
# ─────────────────────────────────────────────────────────────
class BulkUser {
    [string]$SamAccountName
    [string]$UPN
    [string]$Status
}

$script:BulkUsers = [System.Collections.ObjectModel.ObservableCollection[BulkUser]]::new()
$BulkGrid.ItemsSource = $script:BulkUsers

# ─────────────────────────────────────────────────────────────
#  Händelse: Ladda CSV
# ─────────────────────────────────────────────────────────────
$BtnLoadCsv.Add_Click({
    $path = $TxtCsvPath.Text.Trim()
    if (-not (Test-Path $path)) { Write-Log "CSV-filen hittades inte: $path" "Error"; return }

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
            if (-not $sam -or -not $upn) { Write-Log "Rad saknar sAMAccountName/UPN — hoppar över." "Warning"; continue }
            $u = [BulkUser]::new(); $u.SamAccountName = $sam; $u.UPN = $upn; $u.Status = "—"
            $script:BulkUsers.Add($u)
        }
        Write-Log "Laddade $($script:BulkUsers.Count) användare från CSV." "Success"
        $BtnRunBulk.IsEnabled = $script:BulkUsers.Count -gt 0
    } catch {
        Write-Log "Misslyckades att läsa CSV: $_" "Error"
    }
})

# ─────────────────────────────────────────────────────────────
#  Händelse: Kör bulk
# ─────────────────────────────────────────────────────────────
$BtnRunBulk.Add_Click({
    if ($script:BulkUsers.Count -eq 0) { return }

    $dry = [bool]$ChkDryRunBulk.IsChecked
    $cfg = Get-UISettings

    $syncCred = $null
    if (-not $dry) {
        $syncCred = Get-Credential -Message "Autentisering för $($cfg.SyncServer)"
        if (-not $syncCred) { Write-Log "Avbrutet." "Warning"; return }
    }

    $BtnRunBulk.IsEnabled = $false
    $TxtBulkStatus.Text   = "Kör..."

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
            Write-Log "ActiveDirectory-modulen hittades inte. Installera RSAT: AD DS Tools och starta om verktyget." "Error"; return
        }
        try { Import-Module $_adModPath -ErrorAction Stop }
        catch { Write-Log "Kunde inte ladda ActiveDirectory: $_" "Error"; return }
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            Write-Log "Microsoft.Graph-modulen saknas. Installerar (kräver internetåtkomst)..." "Warning"
            try { Install-Module Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop }
            catch { Write-Log "Installationen misslyckades: $_. Kör: Install-Module Microsoft.Graph" "Error"; return }
        }
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Import-Module Microsoft.Graph.Users -ErrorAction Stop
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
        }
        catch { Write-Log "Kunde inte ladda Microsoft.Graph-moduler: $_" "Error"; return }

        $session = $null
        if (-not $_dry -and $_syncCred) {
            try {
                $session = New-PSSession -ComputerName $_cfg.SyncServer -Credential $_syncCred -ErrorAction Stop
                Write-Log "Ansluten till $($_cfg.SyncServer)." "Success"
            } catch {
                Write-Log "Misslyckades att ansluta till $($_cfg.SyncServer): $_" "Error"
                return
            }
        }

        for ($i = 0; $i -lt $_userList.Count; $i++) {
            $u   = $_userList[$i]
            $idx = $i
            $_dispatcher.Invoke([action]{
                $_bulkUsers[$idx].Status = 'Kör...'
                $_bulkGrid.Items.Refresh()
            })

            $ok = Convert-UserToCloudOnly `
                -SamAccountName $u.Sam -UPN $u.UPN `
                -Settings       $_cfg `
                -SyncCred       $_syncCred `
                -SyncSession    $session `
                -DryRun         $_dry `
                -Log            ${function:Write-Log}

            $st = if ($ok) { 'OK' } else { 'FEL' }
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
        $failed = ($script:BulkUsers | Where-Object { $_.Status -eq 'FEL' }).Count
        $total  = $script:BulkUsers.Count
        $BtnRunBulk.IsEnabled = $true
        $TxtBulkStatus.Text   = "Klart: $($total - $failed)/$total lyckades"
        Write-Log "Bulk klar. $($total - $failed) av $total lyckades." $(if ($failed -eq 0) {'Success'} else {'Warning'})
    })
    $script:_bulkTimer.Start()
})

# ─────────────────────────────────────────────────────────────
#  Städa vid stängning
# ─────────────────────────────────────────────────────────────
$window.Add_Closed({
    if ($script:SyncSession) { Remove-PSSession $script:SyncSession -ErrorAction SilentlyContinue }
})

# ─────────────────────────────────────────────────────────────
#  Visa fönstret
# ─────────────────────────────────────────────────────────────
Write-Log "Välkommen! Fyll i inställningar och välj Enskild användare eller Bulk." "Info"
Write-Log "Dry run är aktiverat som standard — inga ändringar görs förrän du avmarkerar det." "Warning"
[void]$window.ShowDialog()
