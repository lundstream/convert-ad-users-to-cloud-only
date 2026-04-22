========================================================
  Convert AD-användare till Cloud-Only (Entra ID)
========================================================

VAD VERKTYGET GÖR
-----------------
Konverterar en AD-synkad användare till ett fristående
molnkonto i Entra ID (Azure AD) genom att:

  1. Flytta AD-kontot till ett icke-synkat OU
  2. Köra en delta-synk → kontot tas bort i Entra ID
  3. Pausa synkschema
  4. Återställa kontot i Entra ID
  5. Rensa onPremisesImmutableId via Graph API
  6. Återaktivera synkschema och köra en ny delta-synk

KRAV
----
  - Windows PowerShell 5.1 (körs som administratör)
  - RSAT: Active Directory DS Tools installerat
    (Inställningar → Appar → Valfria funktioner)
  - Microsoft.Graph PowerShell-modul
    (installeras via Install-Prerequisites.ps1)
  - Nätverksåtkomst (WinRM) till Azure AD Connect-servern
  - Behörighet: User Administrator eller Global Administrator i Entra ID
  - AD-behörighet att flytta användare mellan OU:n

FÖRBEREDELSER (kör en gång per server)
--------------------------------------
  Högerklicka på Install-Prerequisites.ps1 → Kör som administratör

  Skriptet installerar automatiskt de moduler som saknas:
    - Microsoft.Graph (Authentication, Users, Identity.DirectoryManagement)
    - NuGet-provider (krävs av PowerShellGet)
    - RSAT: Active Directory DS Tools (om det saknas)

  På servrar MED internet hämtas allt direkt från PSGallery.
  På servar UTAN internet, kör Download-Prerequisites.ps1 på en
  internetansluten dator/server först och kopiera mappen Offline-Packages\
  hit — Install-Prerequisites.ps1 känner av detta automatiskt.

KOM IGÅNG
---------
  1. Högerklicka på Launch-GUI.bat → Kör som administratör
  2. Gå till fliken "Inställningar" och fyll i:
       - Mål-OU: det OU dit AD-kontot ska flyttas (icke-synkat)
       - AAD Connect-server: hostname till din sync-server
       - Managed domain-suffix: din onmicrosoft.com-domän
       - Loggmapp: var XML-säkerhetskopior sparas
  3. Klicka "Spara inställningar"

ENSKILD ANVÄNDARE
-----------------
  1. Välj fliken "Enskild användare"
  2. Fyll i sAMAccountName (AD-inloggning) och UPN (Entra ID)
  3. Avmarkera "Dry run" när du är redo att köra på riktigt
  4. Klicka "Kör konvertering"
  5. Ange autentisering för sync-servern när dialogrutan visas
  6. Följ förloppet i loggrutan längst ned

BULK (CSV-FIL)
--------------
  1. Skapa en CSV-fil med rubrikrad:
       sAMAccountName,UPN
       john.doe,john.doe@contoso.com
       jane.smith,jane.smith@contoso.com

     (Se example-users.csv som mall)

  2. Välj fliken "Bulk (CSV-fil)"
  3. Klicka "Välj CSV-fil..." och välj din fil
  4. Klicka "Ladda" – användarna visas i listan
  5. Avmarkera "Dry run" när du är redo
  6. Klicka "Kör alla"
  7. Status per användare visas i kolumnen till höger (OK / FEL)

DRY RUN
-------
Dry run är aktiverat som standard. Ingenting ändras i AD
eller Entra ID – alla steg loggas som om de körts på riktigt.
Använd detta för att verifiera inställningar innan du kör skarpt.

LOGG OCH SÄKERHETSKOPIA
------------------------
För varje konverterad användare sparas en XML-fil med
AD-kontots alla attribut i loggmappen (inställningsbar).
Filerna döps till: <sAMAccountName>_AD_pre-move_<datum>.xml

SERVRAR UTAN INTERNETUPPKOPPLING
-----------------------------
Används verktyget på en server utan internetåtkomst behöver du
köra de två hjälpskripten i ordning:

  Steg A – på en dator/server MED internet:
    Högerklicka på Download-Prerequisites.ps1 → Kör som administratör

    Skriptet laddar ned och sparar till mappen Offline-Packages\:
      1. NuGet-provider DLL
      2. PowerShellGet (uppdaterad)
      3. Microsoft.Graph.Authentication  (v2+)
         Microsoft.Graph.Users
         Microsoft.Graph.Identity.DirectoryManagement
      4. Instruktioner för RSAT (kan inte paketers som PS-modul)

    Kopiera hela mappen Offline-Packages\ till målservern.

  Steg B – på MÅLSERVERN (offline):
    Högerklicka på Install-Prerequisites.ps1 → Kör som administratör

    Skriptet:
      1. Installerar NuGet-provider från Offline-Packages\NuGetProvider\
      2. Registrerar Offline-Packages\Modules\ som tillfällig PSRepository
      3. Installerar de tre Graph-modulerna från den lokala repot
         (fallback: kopierar modulkataloger direkt till WindowsPowerShell\Modules)
      4. Installerar RSAT Active Directory-verktyg:
           - Windows Server:  Install-WindowsFeature RSAT-AD-PowerShell
           - Windows 10/11:   Add-WindowsCapability (Rsat.ActiveDirectory…)
      5. Avregistrerar den tillfälliga PSRepository
      6. Verifierar att alla moduler är tillgängliga

    Starta om verktyget efter att skriptet slutförts.

  Obs! Moduler kan även installeras med internet direkt från GUI:t om
  Microsoft.Graph saknas – en prompt visas automatiskt.

FELSÖKNING
----------
  - "ActiveDirectory-modulen hittades inte"
    → Installera RSAT: Active Directory DS Tools via
      Inställningar → Appar → Valfria funktioner
      eller kör Install-Prerequisites.ps1 (se ovan)

  - "Microsoft.Graph-modulen saknas"
    → Verktyget försöker installera automatiskt, men om det
      misslyckas, kör manuellt:
      Install-Module Microsoft.Graph -Scope CurrentUser
      eller använd Install-Prerequisites.ps1 för offline-installation

  - Fönstret öppnas inte alls
    → Kontrollera att Launch-GUI.bat körs som administratör
      och att ExecutionPolicy tillåter skript:
      Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
