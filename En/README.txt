========================================================
  Convert AD user to Cloud-Only (Entra ID)
========================================================

WHAT THE TOOL DOES
------------------
Converts an AD-synced user into a standalone cloud
account in Entra ID (Azure AD) by:

  1. Moving the AD account to a non-synced OU
  2. Running a delta sync -> the account is deleted in Entra ID
  3. Pausing the sync schedule
  4. Restoring the account in Entra ID
  5. Clearing onPremisesImmutableId via Graph API
  6. Re-enabling the sync schedule and running a new delta sync

REQUIREMENTS
------------
  - Windows PowerShell 5.1 (run as administrator)
  - RSAT: Active Directory DS Tools installed
    (Settings -> Apps -> Optional features)
  - Microsoft.Graph PowerShell module
    (installed via Install-Prerequisites.ps1)
  - Network access (WinRM) to the Azure AD Connect server
  - Permissions: User Administrator or Global Administrator in Entra ID
  - AD rights to move users between OUs

PREREQUISITES (run once per server)
-----------------------------------
  Right-click Install-Prerequisites.ps1 -> Run as administrator

  The script automatically installs missing modules:
    - Microsoft.Graph (Authentication, Users, Identity.DirectoryManagement)
    - NuGet provider (required by PowerShellGet)
    - RSAT: Active Directory DS Tools (if missing)

  On servers WITH internet, everything is pulled directly from PSGallery.
  On servers WITHOUT internet, run Download-Prerequisites.ps1 on an
  internet-connected computer/server first and copy the Offline-Packages\
  folder here - Install-Prerequisites.ps1 detects it automatically.

GETTING STARTED
---------------
  1. Right-click Launch-GUI.bat -> Run as administrator
  2. Go to the "Settings" tab and fill in:
       - Target OU: the OU to move the AD account into (non-synced)
       - AAD Connect server: hostname of your sync server
       - Managed domain suffix: your onmicrosoft.com domain
       - Log folder: where XML backups are saved
  3. Click "Save settings"

SINGLE USER
-----------
  1. Select the "Single user" tab
  2. Fill in sAMAccountName (AD logon name) and UPN (Entra ID)
  3. Uncheck "Dry run" when you are ready to run for real
  4. Click "Run conversion"
  5. Provide credentials for the sync server when prompted
  6. Follow progress in the log pane at the bottom

BULK (CSV FILE)
---------------
  1. Create a CSV file with a header row:
       sAMAccountName,UPN
       john.doe,john.doe@contoso.com
       jane.smith,jane.smith@contoso.com

     (See example-users.csv as a template)

  2. Select the "Bulk (CSV file)" tab
  3. Click "Choose CSV file..." and pick your file
  4. Click "Load" - users appear in the list
  5. Uncheck "Dry run" when ready
  6. Click "Run all"
  7. Per-user status appears in the right-hand column (OK / FAIL)

DRY RUN
-------
Dry run is enabled by default. Nothing is changed in AD or
Entra ID - all steps are logged as if they had run for real.
Use this to verify settings before running in production.

LOGS AND BACKUPS
----------------
For every converted user, an XML file with all attributes of
the AD account is saved in the log folder (configurable).
Files are named: <sAMAccountName>_AD_pre-move_<date>.xml

AIR-GAPPED / OFFLINE NETWORKS
-----------------------------
If the tool is used on a server without internet access, run
the two helper scripts in order:

  Step A - on a computer/server WITH internet:
    Right-click Download-Prerequisites.ps1 -> Run as administrator

    The script downloads and saves to the Offline-Packages\ folder:
      1. NuGet provider DLL
      2. PowerShellGet (updated)
      3. Microsoft.Graph.Authentication  (v2+)
         Microsoft.Graph.Users
         Microsoft.Graph.Identity.DirectoryManagement
      4. Instructions for RSAT (cannot be packaged as a PS module)

    Copy the entire Offline-Packages\ folder to the target server.

  Step B - on the TARGET SERVER (offline):
    Right-click Install-Prerequisites.ps1 -> Run as administrator

    The script:
      1. Installs the NuGet provider from Offline-Packages\NuGetProvider\
      2. Registers Offline-Packages\Modules\ as a temporary PSRepository
      3. Installs the three Graph modules from the local repo
         (fallback: copies module folders directly to WindowsPowerShell\Modules)
      4. Installs RSAT Active Directory tools:
           - Windows Server:  Install-WindowsFeature RSAT-AD-PowerShell
           - Windows 10/11:   Add-WindowsCapability (Rsat.ActiveDirectory...)
      5. Unregisters the temporary PSRepository
      6. Verifies that all modules are available

    Restart the tool once the script completes.

  Note: Modules can also be installed over the internet directly from
  the GUI if Microsoft.Graph is missing - a prompt appears automatically.

TROUBLESHOOTING
---------------
  - "ActiveDirectory module not found"
    -> Install RSAT: Active Directory DS Tools via
       Settings -> Apps -> Optional features
       or run Install-Prerequisites.ps1 (see above)

  - "Microsoft.Graph module missing"
    -> The tool tries to install it automatically, but if that
       fails, run manually:
       Install-Module Microsoft.Graph -Scope CurrentUser
       or use Install-Prerequisites.ps1 for offline installation

  - The window does not open at all
    -> Make sure Launch-GUI.bat is run as administrator and
       that ExecutionPolicy allows scripts:
       Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
