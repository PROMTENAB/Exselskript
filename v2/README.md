# Inventering – testversion 2

Detta är en separat testversion baserad på VBA-koden i den senast uppladdade
arbetsboken. Den äldre modulen i repots rot är kvar för återställning.
**Excel för Windows har inte körts här. Filen ska först kompileras och provas
på en kopia, inte ersätta en produktionsarbetsbok direkt.**

## Börja här

1. Packa upp leveransens ZIP-fil lokalt. Originalet finns under `original/`.
2. Öppna `Inventering_v2_test.xlsm` i Excel för Windows och tillåt makron om du litar på filen.
3. Alt+F11 → Felsök → Kompilera VBAProject. Detta är fortfarande en ej utförd kontroll.
4. Kör `InstalleraKnappar` en gång. De gamla makronamnen är kvar; regelknappar läggs på Regler.
5. Fyll `Återbruk i projekt` i Rådata. Den tillhandahållna filen saknade fältet;
   inga JA/NEJ-beslut har gissats från SKICK eller ANDELSPARAS.
6. Kör `UppdateraRapporter`. Rapportbladen som följer med testfilen är originalets
   tidigare rapporter tills makrot körts. Granska därefter KONTROLL.

Den nya rapporten visar Återbruk i projekt i kolumn F. Procent och beräknade
återbruksmängder visas inte. Rapportval!B26 ska vara 1; nivå 2–3 är inte implementerade.
En procentnolla betyder aldrig automatiskt NEJ eller kassering.

## Regler på samma flik

De ursprungliga reglerna är bevarade överst som referens. Makrot använder dessa nya tabeller:

| Tabell | Start vid leverans | Syfte |
|---|---|---|
| tblKategorierV2 | A40 | Materialkategori, tre rubriknivåer och grundläggande enhet |
| tblGrupperV2 | A100 | Grundregel, Ersätt eller Komplettera med Ja/Nej per fält |
| tblAterbrukV2 | A120 | Översättning av beslut till JA, NEJ eller EJ FASTSTÄLLT |
| tblEnheterV2 | A135 | Särskilda enhetsval: exakt SMG2 och CLASS Hårdgjort |

Tabellnamnen används, inte tabellernas radnummer. Lägg till rader via tabellen
eller knapparna; Excel ska då flytta tabeller under vid behov. Ändra inte regel-ID
när du sorterar eller ändrar prioritet. Regelknappar skapar nya, inaktiva regler
med tidsstämplade ID:n. Aktivera när villkoren är klara.

Välj Fält 1/Metod 1/Värde 1 och vid behov villkor 2–3. Välj kombination i Logik,
exempelvis `1 OCH (2 ELLER 3)`. Ingen fri kod eller Eval körs från cellerna.
Fält X/Metod X/Värde X är ett faktiskt undantag som utesluter matchningen.
Lokalt delområde behåller den gamla innebörden: en lokal regel har företräde
framför generella regler, därefter gäller högst numerisk prioritet.
Motstridiga regler med samma högsta prioritet stoppar körningen och skrivs i
KONTROLL. Regler som ger identiska resultat loggas som varning.
Beskrivning är en genererad förklaringskolumn och uppdateras vid regelknappar
eller rapportuppdatering. Sökvärden jämförs utan skiftlägeskänslighet och med
normaliserade mellanslag. Exakt KS1 matchar inte KS10.

Kategorireglerna migreras från den bifogade filen; materialpolicyn ändras inte.
Enhetsregler kan överstyra kategoriens enhet. SMG2 är exakt matchning och får lpm.
CLASS Hårdgjort får m². Övriga Auto använder längd, därefter area, därefter antal.

## Gruppering och beräkning

Kategorier, enhet, återbruksbeslut och respektive områdesavgränsning är alltid
separerade. Gruppregler bestämmer ytterligare fält. Standardutrustning grupperas
på littera, övriga standardregler på littera och materialtyp. Den inaktiva
CS1-exempelregeln visar hur SKICK kan komplettera grupperingen.
Tomt och EJ FASTSTÄLLT får samma rapportbeslut. Okända beslut flaggas och redovisas
som EJ FASTSTÄLLT. NEJ betyder inte KASSERAS. Tidigare automatisk
SPARAS/KASSERAS-indelning av utrustning används inte i version 2.

Den dolda fliken `_InventeringV2` kopplar objekt-ID, regel-ID och gruppnycklar
till mängder. ID-uppslag hittar källobjekt även när rader flyttas inom underlaget.
Gruppsummeringen använder gruppidentitet, inte enbart littera/material. Mängder
för befintliga objekt räknas om genom formler. Ändrade beslut, regler, kategorier,
nya objekt eller ändrad kolumnordning kräver `UppdateraRapporter`.
Källkolumner identifieras efter rubriknamn när rapporten/importen körs. Ingen
2000-radersgräns används i version 2. Uppdatering behövs efter omflyttning av kolumner.

## KONTROLL och bevarande

Verifierad exporttotal blir inte rapportobjekt. I denna fil stämmer rad 2 med
summan av övriga ID:n, antal, area och längd. Raden är bevarad i Rådata.
Helt tomma rader ignoreras. Saknade/dubbla ID:n och ogiltiga mängder placeras i
KONTROLL med källrad och länk och tas inte med i mängdsummeringen. Filtrering via
Rapportval loggas. Regeltvister stoppar körningen i stället för godtycklig vinnare.

Varje rapportkörning gör en separat fullständig XLSM-säkerhetskopia innan den
bygger om rapporterna. En dialog varnar om att egna rapportanteckningar och
format ännu inte flyttas automatiskt. Rådata, originalregler, bildregister,
rapportmall och originalbilder lämnas kvar. Genererade rapportblad byggs om som
förut. Fotolänkar fortsätter gå direkt till första filen med antal bildnamn.
Det gamla bildregistret finns kvar men används inte av de direkta fotolänkarna.

## Import

Importen matchar rubriknamn och objekt-ID, accepterar nya kolumner och bygger
inte på fasta kolumnnummer. Den läser första kalkylbladets rad 1 som rubriker.
Oidentifierade importposter och dubbla import-ID stoppar importen före skrivning.
Exporttotal får bara ignoreras efter summeringskontroll. Befintliga rådata raderas
inte. Med Behåll bedömningar=Ja bevaras SKICK, KOMMENTAR, ANDELSPARAS och Återbruk
i projekt. Skillnader i icke-tomma bedömningar loggas i IMPORTKONTROLL.
Rapporten uppdateras med en separat knapp efter import. Importen är ej körd i Excel.

## Återställning

Stäng testfilen och öppna `original/test2.xlsm` eller en automatisk
`Inventering_backup_*.xlsm`. Originalets SHA-256 är
`8dc3b96be88e046f0d901720557ce90e70e3bb48242e7e9ea2372980e95731fe`.
Den gamla versionen i repots rot har inte ersatts. Importera inte enbart V2-modulen
i en gammal bok utan de nya tabellerna: använd den förberedda testarbetsboken.

## Återstår

- Excel-kompilering, full rapportkörning, upprepad körning, import, knappar,
  utskrift, logotyp och faktisk prestanda är inte verifierade i Windows.
- Förhandsgranskning, transaktionell återställning vid körfel, automatisk
  anteckningsöverföring, egen formatering över ombyggnader och ändringshistorik
  är uppskjutna enligt den förenklade beställningen. Säkerhetskopian är återställningsvägen.
- Nivå 2–3, ANDEL BEDÖMD och procentbaserad återbruksmängd är inte implementerade.
- Regelkonflikter stoppar körningen; KONTROLL är en enkel kontrollista, inte en fullständig granskningsdialog.

## Bygga och kontrollera paketet igen

Installera `openpyxl`, `oletools`, `olefile` och `lxml` i en Python-miljö.

```sh
python v2/build_test_workbook.py original.xlsm test.xlsm v2/Inventering_Rapportautomatik.bas
python tests/test_v2_reference.py original.xlsm test.xlsm
```

Byggaren redigerar enbart utvalda XML-delar och VBA-källan. Alla andra ZIP-delar
bevaras. VBA:s versionsbundna kompileringscache ogiltigförklaras; Excel måste
kompilera källan. Samtliga ursprungliga klassmodulers källkod bevaras och den
inbäddade källan extraheras igen för kontroll. Detta ersätter inte Excel-test.
