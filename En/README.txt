================================================================
  Convert AD User to Cloud-Only (Entra ID)
  English version
================================================================

OVERVIEW
--------
This tool converts an Active Directory-synced user into a
standalone cloud-only account in Entra ID (Azure AD). It is a
WPF GUI built on Windows PowerShell 5.1 and drives the entire
flow from one place: AD, Azure AD Connect, and Microsoft Graph.

It supports two modes:

  - Single user  - convert one user at a time
  - Bulk         - import a CSV and process many users in sequence

A Dry Run checkbox is enabled by default so you can verify
settings and see the full log output without making any
changes in AD or Entra ID.


HOW IT WORKS (per user)
-----------------------
For each user the tool runs six sequential steps:

  [1/6]  MOVE IN AD
         - Fetches the AD object and exports all attributes to
           an XML file (pre-move backup) in the log folder.
         - Moves the account to the configured non-synced OU so
           AAD Connect will treat it as out-of-scope.

  [2/6]  DELTA SYNC + PAUSE
         - Connects to the Azure AD Connect server via WinRM.
         - Waits for any in-progress sync cycle to finish.
         - Disables the sync schedule, runs a Delta sync so the
           now out-of-scope user is soft-deleted in Entra ID,
           then waits for the cycle to complete.

  [3/6]  CONNECT TO GRAPH
         - Interactive sign-in (Connect-MgGraph) using the
           User.ReadWrite.All scope.

  [4/6]  FIND + RESTORE DELETED USER
         - Searches the recycle bin using three strategies:
             1. Exact UPN match
             2. Exact local-part match (before '@')
             3. Fuzzy "contains" match on UPN and Mail (handles
                the Entra soft-delete format
                {guid}originalname@domain)
         - If no match is found, triggers another Delta sync and
           retries (up to 5 attempts, ~30s apart).
         - Restores the matched user by object Id and waits for
           Entra to re-index.

  [5/6]  CLEAR ImmutableId
         - Reads the restored user by Id (the UPN can change on
           restore).
         - If the current UPN is in a federated/verified domain,
           temporarily renames it to <localpart>@<managed suffix>
           so the ImmutableId write succeeds, then renames back.
         - PATCHes onPremisesImmutableId = null via Graph.
         - Verifies the attribute is cleared.

  [6/6]  RESUME SYNC
         - Re-enables the AAD Connect sync schedule and kicks
           off a final Delta sync.


REQUIREMENTS
------------
Host running the tool:
  - Windows 10/11 or Windows Server 2016+
  - Windows PowerShell 5.1 (NOT PowerShell 7)
  - Running as administrator
  - RSAT: Active Directory DS Tools
  - Microsoft.Graph modules v2+:
      Microsoft.Graph.Authentication
      Microsoft.Graph.Users
      Microsoft.Graph.Identity.DirectoryManagement
  - WinRM access to the AAD Connect server

Identity:
  - Entra role: User Administrator or Global Administrator
  - AD rights to read and move the user object
  - Credentials with admin access on the AAD Connect server

Network:
  - Outbound HTTPS to graph.microsoft.com and login.microsoftonline.com
  - TCP 5985/5986 (WinRM) to the AAD Connect server


INSTALLATION
------------
One-time per host:

  1. Right-click Install-Prerequisites.ps1 -> Run as administrator
     The script installs:
       - NuGet package provider
       - Microsoft.Graph modules (Authentication, Users,
         Identity.DirectoryManagement)
       - RSAT: Active Directory DS Tools (prompts Y/N)
     If Offline-Packages\ is present next to the script, it
     installs from there instead of from PSGallery.

  2. Verify the summary at the end reports every item as [OK].


AIR-GAPPED / OFFLINE INSTALLATION
---------------------------------
Use this when the target host has no internet access.

  Step A - on an internet-connected machine
     Right-click Download-Prerequisites.ps1 -> Run as administrator
     Creates an Offline-Packages\ folder containing:
       - NuGet provider DLL
       - Updated PowerShellGet
       - All three Microsoft.Graph modules (with dependencies)
       - manifest.json (metadata about the bundle)

  Step B - on the target host
     Copy the Offline-Packages\ folder next to
     Install-Prerequisites.ps1.
     Right-click Install-Prerequisites.ps1 -> Run as administrator.
     The script:
       - Installs the NuGet provider from the bundle
       - Registers a temporary local PSRepository
       - Installs the Graph modules from it
       - Unregisters the temporary repository
       - Prompts for RSAT installation (cannot be packaged as a module)

  Note: the GUI can also self-install Microsoft.Graph at runtime
  if the modules are missing and internet is available.


GETTING STARTED
---------------
  1. Right-click Launch-GUI.bat -> Run as administrator.

  2. On the Settings tab, fill in:

       Target OU                 DN of a non-synced OU, e.g.
                                 OU=Disabled Objects,DC=contoso,DC=com

       AAD Connect server        Hostname or FQDN of the sync server,
                                 e.g. azadc.contoso.com

       Managed domain suffix     Your onmicrosoft.com domain (non-federated),
                                 e.g. contoso.onmicrosoft.com
                                 Used as a temporary UPN during the
                                 ImmutableId clearing if the user is on
                                 a federated domain.

       Log folder                Folder for pre-move XML backups and logs.
                                 Default:
                                 %USERPROFILE%\Documents\CloudOnly-Logs

       Sync wait time            Seconds to wait after the Delta sync
                                 before querying the recycle bin
                                 (default 180).

       Restore wait time         Seconds to wait after restoring the user
                                 before further Graph calls (default 20).

  3. Click "Save settings".
     Settings are written to Convert-to-CloudOnly-Settings.json
     next to the script.


SINGLE-USER MODE
----------------
  1. Select the "Single user" tab.
  2. Enter:
       sAMAccountName   The AD logon name (without domain).
       UPN              The user's UserPrincipalName in Entra ID.
  3. Leave "Dry run" enabled for a rehearsal, or uncheck it to
     run for real.
  4. Click "Run conversion".
  5. Provide credentials for the AAD Connect server in the dialog.
  6. Sign in to Microsoft Graph in the browser window that opens.
  7. Follow progress in the log pane. On success the status turns
     green ("Done!"); on failure it turns red ("Error - see log").


BULK (CSV) MODE
---------------
  1. Prepare a CSV with the header:

         sAMAccountName,UPN
         john.doe,john.doe@contoso.com
         jane.smith,jane.smith@contoso.com

     (See example-users.csv next to the script.)

  2. Select the "Bulk (CSV file)" tab.
  3. Click "Choose CSV file..." and pick the file.
  4. Click "Load". Users appear in the grid with Status = "-".
  5. Uncheck "Dry run" when ready.
  6. Click "Run all".
  7. Each row updates to Running... -> OK / FAIL as the tool
     progresses. The final summary shows how many succeeded.

  Header aliases accepted: SamAccountName, UserPrincipalName.


DRY RUN
-------
Dry Run is on by default. In Dry Run:

  - No AD move, no sync run, no Graph restore, no ImmutableId change.
  - All six steps are logged as if they had executed.
  - The sync server credentials are NOT requested.

Use Dry Run to verify your Target OU, sync server name, and
managed domain suffix before running live.


LOGS AND BACKUPS
----------------
  - A live color-coded log is shown in the bottom pane of the GUI.
  - Each AD account is exported to:
        <LogFolder>\<sAMAccountName>_AD_pre-move_<yyyyMMdd_HHmmss>.xml
    This XML is a full Export-Clixml of the user object and can
    be reviewed or replayed with Import-Clixml.
  - Settings are persisted to Convert-to-CloudOnly-Settings.json.


TROUBLESHOOTING
---------------
"ActiveDirectory module not found"
   Install RSAT via Settings -> Apps -> Optional features, or
   run Install-Prerequisites.ps1.

"Microsoft.Graph module missing"
   The tool attempts auto-install. If that fails, run manually:
     Install-Module Microsoft.Graph -Scope CurrentUser
   or use Install-Prerequisites.ps1 (offline-capable).

"Get-ADUser is not recognized" inside the tool
   Usually means the ActiveDirectory module is not installed
   for the account running the GUI. Install RSAT and re-launch.

"No deleted user found for UPN ..."
   The Delta sync has not soft-deleted the user yet. The tool
   will retry with additional sync cycles, but verify:
     - The user really is in the Target OU.
     - The OU is excluded from AAD Connect's sync scope.
     - No synchronization error is reported on the AAD Connect
       server (Synchronization Service Manager -> Operations).

"Multiple entries match local-part ..."
   You have more than one recycle-bin entry with the same
   local-part. Supply a more specific UPN (full mail address)
   on the Single-user tab.

"Failed to change UPN" / federated domain errors
   Verify that Managed domain suffix is set to a verified
   non-federated onmicrosoft.com domain on your tenant.

Window does not open
   Make sure Launch-GUI.bat runs as administrator and that
   ExecutionPolicy allows scripts:
     Set-ExecutionPolicy RemoteSigned -Scope CurrentUser


SECURITY NOTES
--------------
  - The GUI never stores the AAD Connect credentials; they are
    held in memory only for the duration of the PSSession.
  - Graph authentication uses the interactive broker and
    respects Conditional Access / MFA.
  - XML backups in the log folder contain sensitive attributes
    (incl. mail, phone, manager references). Protect the folder
    with NTFS ACLs appropriate to your environment.


FILES
-----
  Convert-to-CloudOnly-GUI.ps1   Main GUI and conversion logic
  Install-Prerequisites.ps1      Installs modules and RSAT
  Download-Prerequisites.ps1     Builds the offline bundle
  Launch-GUI.bat                 Starts the GUI with -Sta -NoProfile
  example-users.csv              Bulk mode template
  README.txt                     This file
