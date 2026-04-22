================================================================
  Konvertera AD-användare till Cloud-Only (Entra ID)
  Svensk version
================================================================

ÖVERSIKT
--------
Verktyget konverterar en AD-synkad användare till ett fristående
molnkonto i Entra ID (Azure AD). Det är ett WPF-gränssnitt byggt
på Windows PowerShell 5.1 som driver hela flödet från en plats:
AD, Azure AD Connect och Microsoft Graph.

Två lägen stöds:

  - Enskild användare  - konvertera en användare i taget
  - Bulk               - läsa in en CSV och köra många i följd

Kryssrutan "Dry run" är aktiverad som standard så att du kan
verifiera inställningar och se hela loggutdatat utan att något
ändras i AD eller Entra ID.


SÅ FUNGERAR DET (per användare)
-------------------------------
För varje användare kör verktyget sex sekventiella steg:

  [1/6]  FLYTTA I AD
         - Hämtar AD-objektet och exporterar alla attribut till
           en XML-fil (säkerhetskopia före flytt) i loggmappen.
         - Flyttar kontot till det konfigurerade icke-synkade OU:t
           så att AAD Connect behandlar det som utanför sync-omfång.

  [2/6]  DELTA-SYNK + PAUS
         - Ansluter till Azure AD Connect-servern via WinRM.
         - Väntar in eventuell pågående synkcykel.
         - Inaktiverar synkschemat, kör en Delta-synk så att den
           nu utanförstående användaren soft-delete:as i Entra ID,
           och väntar tills cykeln är klar.

  [3/6]  ANSLUT TILL GRAPH
         - Interaktiv inloggning (Connect-MgGraph) med scopet
           User.ReadWrite.All.

  [4/6]  HITTA + ÅTERSTÄLL RADERAD ANVÄNDARE
         - Söker i papperskorgen med tre strategier:
             1. Exakt UPN-matchning
             2. Exakt local-part (före '@')
             3. Fuzzy "contains" på UPN och Mail (hanterar
                Entras soft-delete-format
                {guid}originalnamn@domain)
         - Om ingen matchning hittas körs en ny Delta-synk och
           försöket görs om (upp till 5 försök, ~30 s mellan).
         - Återställer matchad användare via objekt-Id och väntar
           på att Entra ska indexera om.

  [5/6]  RENSA ImmutableId
         - Läser den återställda användaren via Id (UPN kan ändras
           vid återställning).
         - Om den aktuella UPN:en ligger i en federerad/verifierad
           domän byts den tillfälligt till
           <localpart>@<managed-domain> så att ImmutableId-skrivningen
           lyckas, och byts sedan tillbaka.
         - PATCH:ar onPremisesImmutableId = null via Graph.
         - Verifierar att attributet är tomt.

  [6/6]  ÅTERUPPTA SYNK
         - Aktiverar AAD Connects synkschema igen och startar en
           avslutande Delta-synk.


KRAV
----
Dator som kör verktyget:
  - Windows 10/11 eller Windows Server 2016+
  - Windows PowerShell 5.1 (INTE PowerShell 7)
  - Körs som administratör
  - RSAT: Active Directory DS Tools
  - Microsoft.Graph-moduler v2+:
      Microsoft.Graph.Authentication
      Microsoft.Graph.Users
      Microsoft.Graph.Identity.DirectoryManagement
  - WinRM-åtkomst till AAD Connect-servern

Identitet:
  - Entra-roll: User Administrator eller Global Administrator
  - AD-rättighet att läsa och flytta användarobjektet
  - Konto med adminåtkomst på AAD Connect-servern

Nätverk:
  - Utgående HTTPS till graph.microsoft.com och
    login.microsoftonline.com
  - TCP 5985/5986 (WinRM) till AAD Connect-servern


INSTALLATION
------------
En gång per dator:

  1. Högerklicka Install-Prerequisites.ps1 -> Kör som administratör
     Skriptet installerar:
       - NuGet package provider
       - Microsoft.Graph-moduler (Authentication, Users,
         Identity.DirectoryManagement)
       - RSAT: Active Directory DS Tools (frågar J/N)
     Om mappen Offline-Packages\ finns bredvid skriptet
     installeras från den i stället för från PSGallery.

  2. Verifiera att sammanfattningen i slutet rapporterar [OK]
     för samtliga poster.


OFFLINE-INSTALLATION
--------------------------------
Används när måldatorn saknar internetåtkomst.

  Steg A - på en internetansluten dator
     Högerklicka Download-Prerequisites.ps1 ->
     Kör som administratör
     Skapar mappen Offline-Packages\ med:
       - NuGet provider-DLL
       - Uppdaterad PowerShellGet
       - Samtliga tre Microsoft.Graph-moduler (med beroenden)
       - manifest.json (metadata om paketet)

  Steg B - på måldatorn
     Kopiera mappen Offline-Packages\ bredvid
     Install-Prerequisites.ps1.
     Högerklicka Install-Prerequisites.ps1 ->
     Kör som administratör.
     Skriptet:
       - Installerar NuGet-providern från paketet
       - Registrerar en tillfällig lokal PSRepository
       - Installerar Graph-modulerna därifrån
       - Avregistrerar repositoriet
       - Erbjuder RSAT-installation (kan inte paketeras som modul)

  Obs! GUI:t kan också själv-installera Microsoft.Graph vid
  körning om modulerna saknas och internet finns.


KOM IGÅNG
---------
  1. Högerklicka Launch-GUI.bat -> Kör som administratör.

  2. Under fliken "Inställningar", fyll i:

       Mål-OU                    DN till ett icke-synkat OU, t.ex.
                                 OU=Disabled Objects,DC=contoso,DC=com

       AAD Connect-server        Hostname eller FQDN till sync-servern,
                                 t.ex. azadc.contoso.com

       Managed domain-suffix     Din onmicrosoft.com-domän (icke-federerad),
                                 t.ex. contoso.onmicrosoft.com
                                 Används som tillfällig UPN vid rensning
                                 av ImmutableId om användaren ligger på
                                 en federerad domän.

       Loggmapp                  Mapp för XML-säkerhetskopior och logg.
                                 Standard:
                                 %USERPROFILE%\Documents\CloudOnly-Logs

       Synk-väntetid             Sekunder att vänta efter Delta-synk
                                 innan papperskorgen frågas
                                 (standard 180).

       Återställningstid         Sekunder att vänta efter återställning
                                 innan fler Graph-anrop (standard 20).

  3. Klicka "Spara inställningar".
     Inställningarna skrivs till Convert-to-CloudOnly-Settings.json
     bredvid skriptet.


LÄGE: ENSKILD ANVÄNDARE
-----------------------
  1. Välj fliken "Enskild användare".
  2. Fyll i:
       sAMAccountName   AD-inloggningsnamn (utan domän).
       UPN              Användarens UserPrincipalName i Entra ID.
  3. Låt "Dry run" vara ikryssad för generalrepetition, eller
     avmarkera för att köra på riktigt.
  4. Klicka "Kör konvertering".
  5. Ange autentisering till AAD Connect-servern i dialogen.
  6. Logga in mot Microsoft Graph i webbläsaren som öppnas.
  7. Följ förloppet i loggrutan. Vid lyckat resultat blir statusen
     grön ("Klar!"); vid misslyckande röd ("Fel - se logg").


LÄGE: BULK (CSV)
----------------
  1. Förbered en CSV med rubrikraden:

         sAMAccountName,UPN
         john.doe,john.doe@contoso.com
         jane.smith,jane.smith@contoso.com

     (Se example-users.csv bredvid skriptet.)

  2. Välj fliken "Bulk (CSV-fil)".
  3. Klicka "Välj CSV-fil..." och välj filen.
  4. Klicka "Ladda". Användare visas i listan med Status = "-".
  5. Avmarkera "Dry run" när du är klar.
  6. Klicka "Kör alla".
  7. Varje rad ändras från Kör... -> OK / FEL medan verktyget
     arbetar. Slutsammanfattningen visar antal lyckade.

  Kolumnaliasen SamAccountName och UserPrincipalName accepteras.


DRY RUN
-------
Dry Run är aktiverad som standard. I Dry Run:

  - Ingen AD-flytt, ingen synkkörning, ingen Graph-återställning,
    ingen ImmutableId-ändring.
  - Alla sex steg loggas som om de hade körts.
  - Autentisering till sync-servern efterfrågas INTE.

Använd Dry Run för att verifiera Mål-OU, sync-servernamn och
managed domain-suffix innan skarp körning.


LOGG OCH SÄKERHETSKOPIA
-----------------------
  - En levande färgkodad logg visas i panelen längst ned i GUI:t.
  - För varje AD-konto exporteras:
        <Loggmapp>\<sAMAccountName>_AD_pre-move_<yyyyMMdd_HHmmss>.xml
    Filen är en fullständig Export-Clixml av användarobjektet och
    kan granskas eller spelas upp igen med Import-Clixml.
  - Inställningar lagras i Convert-to-CloudOnly-Settings.json.


FELSÖKNING
----------
"ActiveDirectory-modulen hittades inte"
   Installera RSAT via Inställningar -> Appar -> Valfria funktioner,
   eller kör Install-Prerequisites.ps1.

"Microsoft.Graph-modulen saknas"
   Verktyget försöker installera automatiskt. Om det misslyckas,
   kör manuellt:
     Install-Module Microsoft.Graph -Scope CurrentUser
   eller använd Install-Prerequisites.ps1 (stöder offline).

"Get-ADUser är inte känt" i verktyget
   Innebär oftast att ActiveDirectory-modulen inte är installerad
   för kontot som kör GUI:t. Installera RSAT och starta om.

"Ingen raderad användare hittades för UPN ..."
   Delta-synken har ännu inte soft-delete:at användaren. Verktyget
   gör nya försök med ytterligare synkkörningar, men kontrollera:
     - Att användaren verkligen ligger i Mål-OU:t.
     - Att OU:t är uteslutet från AAD Connects sync-omfång.
     - Att inget synkroniseringsfel rapporteras på AAD Connect-servern
       (Synchronization Service Manager -> Operations).

"Flera poster matchar local-part ..."
   Det finns mer än en post i papperskorgen med samma local-part.
   Ange en mer specifik UPN (fullständig e-postadress) på fliken
   Enskild användare.

"Failed to change UPN" / fel kring federerad domän
   Kontrollera att managed domain-suffix är satt till en verifierad
   icke-federerad onmicrosoft.com-domän på din tenant.

Fönstret öppnas inte
   Kontrollera att Launch-GUI.bat körs som administratör och att
   ExecutionPolicy tillåter skript:
     Set-ExecutionPolicy RemoteSigned -Scope CurrentUser


SÄKERHETSANMÄRKNINGAR
---------------------
  - GUI:t lagrar aldrig AAD Connect-autentiseringen; den finns
    bara i minnet under PSSessionens livstid.
  - Graph-autentiseringen använder interaktiv broker och respekterar
    Conditional Access / MFA.
  - XML-säkerhetskopiorna i loggmappen innehåller känsliga attribut
    (bl.a. mail, telefon, manager-referenser). Skydda mappen med
    NTFS-ACL:er anpassade till er miljö.


FILER
-----
  Convert-to-CloudOnly-GUI.ps1   Huvudsakligt GUI och konverteringslogik
  Install-Prerequisites.ps1      Installerar moduler och RSAT
  Download-Prerequisites.ps1     Bygger offline-paketet
  Launch-GUI.bat                 Startar GUI:t med -Sta -NoProfile
  example-users.csv              Mall för bulkläge
  README.txt                     Denna fil
