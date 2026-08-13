# Aanmeldingsanalyse voor Server Health

## Doel

`Get-ServerHealth.ps1` krijgt een afzonderlijke beveiligingsanalyse van recente Windows-aanmeldingen. De uitbreiding signaleert mogelijke brute-forcepogingen, geslaagde RDP-aanmeldingen vanaf publieke IP-adressen en onverwachte privileged logons. Het script rapporteert uitsluitend en voert geen blokkeringen, accountwijzigingen of firewallacties uit.

## Configuratie

Bovenaan het script komen afzonderlijke, eenvoudig aanpasbare instellingen:

- analyseperiode: 24 uur;
- waarschuwing: 5 tot en met 19 mislukte pogingen per bron-IP of doelaccount;
- kritiek: 20 of meer mislukte pogingen per bron-IP of doelaccount;
- maximum aantal ingelezen Security-events om de looptijd te begrenzen.

Een geslaagde RDP-aanmelding vanaf een publiek IP-adres krijgt status `WARN`.

## Gegevensbron

De analyse leest het lokale Windows-logboek `Security` met `Get-WinEvent` voor:

- event 4624: geslaagde aanmeldingen;
- event 4625: mislukte aanmeldingen;
- event 4672: gevoelige privileges die aan een nieuwe logonsessie zijn toegekend.

Alle velden worden uit de event-XML gelezen via benoemde `EventData/Data`-velden. Daarmee blijft de parser onafhankelijk van de Windows-weergavetaal en van de volgorde waarin eventproperties worden aangeboden.

## Normalisatie

Een formatteringsfunctie zet elk relevant event om naar een object met minimaal:

- event-ID en tijdstip;
- account- en domeinnaam;
- logontype;
- bron-IP en workstation;
- logon-ID;
- elevated-token- of privilegelijst waar beschikbaar;
- status en substatus voor mislukte aanmeldingen.

Ontbrekende XML-velden krijgen een veilige lege of herkenbare standaardwaarde en mogen de healthcheck niet afbreken.

## Detectieregels

### Mislukte aanmeldingen

Event 4625 wordt zowel per geldig bron-IP als per niet-leeg doelaccount gegroepeerd:

- 0–4 pogingen: alleen informatief;
- 5–19 pogingen: `WARN`;
- 20 of meer pogingen: `KRIT`.

De hoogste gevonden status telt mee in de algemene healthcheck.

### Geslaagde interactieve aanmeldingen

Van event 4624 worden logontype 2 (interactief) en 10 (Remote Desktop/RemoteInteractive) getoond. Een type-10-aanmelding met een publiek bron-IP krijgt `WARN`. Private, loopback, link-local en ontbrekende adressen krijgen deze waarschuwing niet.

De IP-classificatie ondersteunt IPv4 en IPv6, inclusief RFC1918, loopback, link-local, carrier-grade NAT en IPv6 unique-local ranges.

### Privileged logons

Event 4672 wordt waar mogelijk via het logon-ID aan een 4624-event gekoppeld. Bekende systeemidentiteiten (`SYSTEM`, `LOCAL SERVICE`, `NETWORK SERVICE`) en computeraccounts die eindigen op `$` worden uit waarschuwingen gefilterd. Overige privileged logons worden zichtbaar gemaakt als aandachtspunt, zonder ze uitsluitend op basis van event 4672 kritiek te noemen.

## Console-uitvoer

Een nieuwe sectie `AANMELDINGSBEVEILIGING` toont:

- aantallen geslaagde interactieve en RDP-aanmeldingen;
- totaal aantal mislukte pogingen;
- verdachte bron-IP’s en doelaccounts met aantallen en status;
- geslaagde RDP-aanmeldingen vanaf publieke IP-adressen;
- niet-systeem privileged logons;
- een beperkte lijst van de recentste relevante events.

De bestaande `Write-Rij`-functie registreert waarschuwingen en kritieke resultaten in `$Rapport`, zodat ze automatisch meetellen in de eindstatus.

## HTML-uitvoer

Het HTML-rapport krijgt een afzonderlijke sectie met:

- een beveiligingssamenvatting;
- een tabel met verdachte groepen per IP en account;
- een tabel met recente interactieve/RDP-aanmeldingen;
- een tabel met relevante privileged logons.

Alle dynamische waarden worden HTML-gecodeerd voordat ze in het rapport terechtkomen.

## Foutafhandeling

- Geen toegang tot het Security-logboek geeft `WARN` en stopt de overige healthcheck niet.
- Geen relevante events geeft `WARN`, met de melding dat het auditbeleid gecontroleerd moet worden.
- Een individueel beschadigd of afwijkend event wordt overgeslagen of met veilige standaardwaarden weergegeven.
- Een onbekend IP-formaat wordt niet als publiek beschouwd en veroorzaakt geen fout.

## Teststrategie

Lokale Pester-tests gebruiken representatieve XML-fixtures en geïsoleerde formatteringsfuncties. Ze controleren:

- parsing van 4624, 4625 en 4672;
- veilige verwerking van ontbrekende velden;
- groepering per IP en account;
- statusgrenzen 4, 5, 19 en 20;
- herkenning van publieke en niet-publieke IPv4- en IPv6-adressen;
- filtering van systeem- en computeraccounts;
- HTML-codering van dynamische waarden;
- foutafhandeling wanneer het Security-logboek niet leesbaar is.

Het volledige productiescript en de tests worden daarnaast door de PowerShell-parser gehaald. De lokale testmap blijft conform de bestaande `.gitignore` buiten GitHub.

## Referenties

- Microsoft Learn: [event 4624](https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4624)
- Microsoft Learn: [event 4648 en correlatie via Logon GUID](https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4648)
- Microsoft Learn: [Windows-logontypes](https://learn.microsoft.com/nl-nl/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/basic-audit-logon-events)
