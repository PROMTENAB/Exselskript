# Testrapport – första testversionen

## Verifierat automatiskt

Fyra referenskontroller kördes i Python mot den bifogade arbetsboken och den
byggda testkopian. **Detta är inte körning av VBA i Excel.**

| Kontroll | Resultat |
|---|---|
| CS1 | 55 objekt i källfilen. Med separat syntetisk JA/NEJ-indelning blir grupperna 53 och 2, totalt 55. Källfilen saknar beslut. |
| KS1 | 6 objekt, BETONGMATERIAL / KANTSTÖD, 161,09071791699728 lpm före avrundning och 161,0 med steg 0,5 m. |
| Smågatsten / SMG2 | Alla tre efterfrågade materialbenämningar matchar NATURSTENSMATERIAL / GATSTEN / SMÅGATSTEN. SMG2: 22 objekt, 106,11854324583435 lpm före avrundning. |
| Bevarande och täckning | Alla ursprungliga rådataceller har samma värden. 519 objekt passerar filtren, 205 filtreras uttryckligen. 1 verifierad exporttotal behandlas separat. Inga saknade eller dubbla ID:n bland de inkluderade källobjekten. |

Originalets ritningsdelar, mediefiler, knappar och rapportmall är byte-identiska
i testkopian. Befintliga rapportblad är också bevarade i paketet och ska byggas
om när UppdateraRapporter körs. De visar alltså inte ännu version 2:s resultat.

VBA:s rutingränser, strängliteraler, radlängder och moduldeklarationer har
kontrollerats statiskt. Den inbäddade VBA-källan extraherades efter paketering
och jämfördes med levererad .bas-fil. Klassmodulernas källkod är oförändrad.
Den kompilerade versionsbundna VBA-cachen har ogiltigförklarats för omkompilering.

## Ej körda kontroller

- Kompilering i Excel för Windows.
- Fullständig VBA-rapportuppdatering och upprepad uppdatering.
- Verkliga formelresultat, knappar, utfällbara rader, fotolänkar och A3-utskrift.
- Import med flyttade kolumner och konflikter i tidigare bedömningar.
- Verklig prestanda med 1 000 objekt och transaktionell återställning vid körfel.

Dessa är **ej testade**, inte godkända. Inga kända fel återstår i de fyra körda
referenskontrollerna; det innebär inte att Excel-versionen är slutverifierad.

## Bevis för oförändrat original

Originalets SHA-256 före och efter:
`8dc3b96be88e046f0d901720557ce90e70e3bb48242e7e9ea2372980e95731fe`

Samma fil finns under original/ i leveranspaketet. Fullständig lista över
ändrade paketdelar och testfilens checksumma finns i paketkontroll.json.

## Manuell kontroll i Excel

1. Öppna testfilen och kompilera VBAProject. Avbryt och rapportera felraden om kompileringen misslyckas.
2. Kör InstalleraKnappar och därefter UppdateraRapporter. Kontrollera KONTROLL.
3. Kontrollera KS1 och SMG2:s mängder samt smågatstensrubrikerna.
4. Kör den tydligt märkta CS1-testfilen för ett syntetiskt 53/2-test. Besluten i
   den filen är testdata och får inte användas som inventeringsbeslut.
5. Kör rapportuppdateringen en andra gång och kontrollera att poster/mängder
   inte dubbleras. Jämför utfällning, bilder, mall, logotyp och utskrift.

## Rättning efter Excel-reparation

DataValidation låg efter pageMargins i arbetsbladets XML, i strid med
OOXML-elementordningen. Ordningen rättades i båda testarbetsböckerna och
byggaren. Alla arbetsblads elementordning har därefter kontrollerats.
Ny öppning i Excel är fortfarande ej verifierad här.
