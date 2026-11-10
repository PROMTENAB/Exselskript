# Optimerad import – manuell installation

Leveransen är en fullständig Inventering_Rapportautomatik.bas för VBA version 2.
Ingen ny XLSM-fil distribueras i denna ändring. Regelmotorn är oförändrad,
förutom rättningen av det reserverade variabelnamnet Local i två funktioner.

1. Spara en kopia av din arbetsbok.
2. Alt+F11: exportera gärna den befintliga modulen för återställning, ta bort den
   gamla Inventering_Rapportautomatik och importera den nya BAS-filen via Arkiv.
   Behåll övriga moduler. Importera inte två kopior av samma makron.
3. Felsök → Kompilera VBAProject. Kompileringen är inte utförd här.
4. Kör ImporteraFranVectorworks. Stäng källfilen först; välj inte målfilen.

## Ändring

Calculation, ScreenUpdating, EnableEvents och StatusBar sparas före arbetet och
återställs på gemensam avslutningsväg, även vid fel och avbrutet filval. Importen
arbetar i manuellt beräkningsläge. Automatisk slutberäkning sker vid återställning
om det var ursprungsläget; inget dubbelt Calculate-anrop görs. Ett ursprungligt
manuellt läge behålls utan framtvingad beräkning.

Källa och befintliga värden läses till arrayer. Sista källrad beräknas en gång.
ID:n, dubbletter, felceller och exporttotaler kontrolleras före skrivning.
Uppdaterade, nya och oförändrade poster hålls isär. En redan verifierad
exporttotal på rad 2 får sina fyra kontrollsummor uppdaterade i minnet, så
att den fortfarande kan identifieras som total efter importen. Endast ändrade celler skrivs,
i sammanhängande kolumnblock. Oförändrade formler och egna kolumner lämnas kvar.
Textblock skrivs med tillfälligt textformat för att bevara text-ID, inledande
nollor och text som börjar med likhetstecken. Tidigare nummerformat återställs.
Blandade format återställs cellvis endast i sådana textblock.

Tabellen utökas vid behov en gång före blockskrivningen. Befintliga tomma
rader i tabellområdet behålls. Statusfältet visar filöppning, säkerhetskopiering,
ID-kontroll, tabelluppdatering, skrivning och slutberäkning. Tabellutökningen görs
före skrivning för att undvika upprepad automatisk utökning.

Behåll bedömningar följer Rapportval: SKICK, KOMMENTAR, ANDELSPARAS och
Återbruk i projekt. Även uttrycklig nolla behålls. Motstridiga icke-tomma
bedömningar loggas i IMPORTKONTROLL med objekt-ID. Källan skriver inga nya
rapporter automatiskt. Strukturella ändringar kräver UppdateraRapporter.

Vid skrivfel kan importen vara delvis genomförd. Makrot anger säkerhetskopians
sökväg och återställer Excel-inställningarna; automatisk återställning av
arbetsbokens data är inte implementerad.

## Kontroller

Godkänt i Python-referensmodell: nya och matchade objekt, bevarade egna fält,
bevarade bedömningar inklusive 0, dubblettavvisning, omkörning och 1 000 objekt.
Statiskt kontrollerat: en beräkning av sista källrad, högst en explicit Resize,
manuell beräkning före filöppning/skrivning, gemensam inställningsåterställning,
blockskrivning och att Local inte används som variabelnamn.
Regelmotorn jämförs med tidigare källkod; inga funktionsändringar har gjorts där.

**Inte testat i Excel för Windows:** kompilering, faktisk importtid, verklig
slutberäkning, formatåterställning, tabellutökning, felvägar och full omkörning.
Python-testerna kör inte VBA och bevisar inte Excel-beteendet.

Byggskriptet innehåller redan XML-rättningen: dataValidations sorteras före
pageMargins. Ingen ny XLSM distribueras innan den ändringen finns i byggaren.
Denna leverans ersätter enbart VBA-källan i din befintliga V2-arbetsbok.
