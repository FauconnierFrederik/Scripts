# Null-safe eventweergave voor Server Health

## Doel

`Get-ServerHealth.ps1` moet Windows-events zonder `TimeCreated` of `Message` kunnen weergeven en exporteren zonder dat de launcher de uitvoering met een null-method-fout afbreekt.

## Ontwerp

- Voeg twee kleine formatteringsfuncties toe: één voor het tijdstip en één voor het bericht.
- Een ontbrekend tijdstip wordt `Onbekend`.
- Een leeg of ontbrekend bericht wordt `(Geen bericht beschikbaar)`.
- Consoleberichten gebruiken alleen de eerste regel en worden tot 80 tekens beperkt.
- HTML-berichten behouden meerdere regels en worden zoals voorheen HTML-gecodeerd.
- De bestaande drempelwaarden, eventselectie, statustelling, launcher en installer veranderen niet.

## Foutafhandeling

De formatteringsfuncties accepteren nullwaarden expliciet en roepen pas methoden aan nadat een veilige tekstwaarde is bepaald. Daarmee worden zowel de consoleweergave als de optionele HTML-export beschermd.

## Teststrategie

Een Pester-regressietest laadt alleen de formatteringsfuncties uit het script en controleert:

- een null-tijdstip;
- een null-bericht;
- afkapping van een lang consolebericht;
- HTML-codering en regeleinden voor export.

Daarnaast wordt het volledige script door de PowerShell-parser gehaald. De interactieve serverchecks worden niet lokaal uitgevoerd omdat ze administratorrechten, Windows Server-componenten en een GUI vereisen.
