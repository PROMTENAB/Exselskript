# Inventering – rapportautomatik

Importera `Inventering_Rapportautomatik.bas` i Excel VBA (Alt+F11).
Ersätt den befintliga modulen med samma namn; spara först en säkerhetskopia av arbetsboken.
Spara som `.xlsm` och kör `UppdateraRapporter` för att ersätta befintliga rapportformler.
Modulen är kodad som Windows-1252 för import i svensk Windows-Excel.

## Ändring av skick

Delområden och Sammanställning använder gruppens första rådatacell i kolumn D
som jämförelsevärde, i stället för ett ID-uppslag som kan fallera för numeriska ID:n.
Samma skick på alla matchande objekt visar det gemensamma värdet; olika skick
visar `Varierande`. Alla tomma skick ger tomt; blandat tomt och ifyllt ger
`Varierande`. Villkorstecken i skicktext behandlas bokstavligt.
Formlerna fortsätter följa ändringar i rådata. Kör UppdateraRapporter efter
omgruppering eller borttagning av källrader. Befintlig gräns på 2000 rader kvarstår.

## Kontroll i Excel

Kontrollera både delområdesblad och Sammanställning:
- Två objekt med numeriska ID:n och samma skick visar skicket.
- Ändra ett skick: resultatet blir Varierande. Återställ: skicket visas igen.
- Alla skick tomma ger tom cell; ett tomt och ett ifyllt ger Varierande.
- Olika delområden med samma skick visar gemensamt skick i Sammanställning.
- Skick med *, ? eller inledande = jämförs som bokstavlig text.

VBA och Excel-formler behöver köras i Excel för slutlig verifiering;
Excel finns inte i den här molnmiljön.

## Urklipp

Mallrader skrivs nu direkt utan Range.Copy. Cellinnehåll/formler, grundformat,
kanter, radhöjd och sammanslagningar inom A:L på en rad överförs.
Villkorsstyrd formatering, dataverifiering, kommentarer, blandad teckenformatering
och sammanslagningar över flera rader överförs inte av denna rutin.
Kontrollera rapportens utseende mot Rapportmall i Excel efter import.
Ändringen eliminerar modulens enda Copy-anrop; hastigheten är inte uppmätt.

## Tidsmätning av knapp 2

Importera den uppdaterade modulen i stället för den gamla. Öppna direktfönstret
i VBA-editorn med Alt+F11 och Ctrl+G. Kör UppdateraRapporter (knapp 2).
Tider skrivs för förberedelser, rensning, bildregister, sammanställning,
varje delområde och slutberäkning, samt en totaltid. Excel-statusfältet
visar pågående steg och återställs vid avslut eller fel.
Slutberäkningen inkluderar återaktivering av automatisk beräkning.
Tiderna mäter verklig förfluten tid; totaltiden inkluderar inte slutdialogen.
Vid fel finns endast loggen för redan avslutade steg, ingen slutförd totaltid.
Mätningen är inte en prestandaoptimering och har inte provkörts i Excel här.

## Direktlänkar till foton

Rapporten använder inte längre fliken Bilder eller dess Ja/Nej-val. Den gamla
fliken lämnas orörd. Kör InstalleraKnappar en gång för att ersätta knapparna
och ta bort den gamla knappen för bildregistret.

Ange fotomappen i Rapportval!B11 och bildnamnen i Rådata kolumn G.
Format som stöds: FOTO(bild1.jpg) FOTO(bild2.jpg), ett enkelt filnamn eller
kommaseparerade filnamn. Fullständiga sökvägar används direkt.
Rapportens befintliga val att visa bilder gäller fortfarande.
Länken visar första bildnamnet och antal unika bildnamn för raden/gruppen.
Den öppnar första bilden direkt; övriga bilder öppnas via fotomappen.
Saknas fotomapp för ett relativt filnamn visas en uppmaning utan klickbar länk.
Inga bildfiler öppnas eller kontrolleras under rapportgenereringen.
Kör UppdateraRapporter efter ändring av bildnamn eller fotomapp.
Tidsloggen visar nu Bildreferenser i stället för Bildregister.

Kontrollera i Excel: ett foto, flera foton, upprepade bildnamn, tomt bildfält,
saknad fotomapp och fullständig sökväg. Kör även utan fliken Bilder.
Excel-körning och åtkomst till användarens fotomapp är inte verifierade här.

## Blockskrivning och gemensamma format

Knapp 2 läser Rådata och Regler till minnet och klassificerar varje objekt en
gång per körning. Varje rapport byggs i en buffert med värden och levande
formler och skrivs med Value2 för värden samt Formula för sammanhängande formelblock. Text och numeriska
objekt-ID:n behåller sina typer så att ID-uppslag fungerar.

Mallens åtta radtyper läses till 96 återanvändbara cellformat med namn
Inventering_Auto_*. Dessa format uppdateras från Rapportmall vid varje körning
och appliceras i omgångar om högst 50 rader. Ingen kopiering via Urklipp används.
Sammanslagningar inom mallraden, grupperingar och fotolänkar hanteras separat.
Befintliga begränsningar för villkorsstyrd formatering och sammanslagningar
över flera rader gäller fortfarande. Genererade delområdesblad byggs fortfarande om.
Excel-händelser pausas under genereringen och återställs även vid fel.

Tidsloggen visar även Gruppering och innehåll, Blockskrivning, Formatering
och Bildlänkar inom varje rapport. Dessa är deltider som redan ingår i
rapportens tid; summera dem inte en gång till med huvudtiderna.
Förberedelser inkluderar nu inläsning, regelklassificering och mallformat.

### Verifiera i Excel före ordinarie användning

Kör på en kopia av samma arbetsbok som tidigare tog 351,41 sekunder.
Jämför kategorier, ordning, radantal, mängder, sparprocent, skick, formatering,
utfällbara rader och fotolänkar med föregående version. Kontrollera både
numeriska ID:n och text-ID:n samt littera med inledande nollor. Ändra därefter
mängd och skick i Rådata och kontrollera att befintliga rapportformler reagerar.
Kör knapp 2 två gånger för att kontrollera att format och innehåll inte växer
eller dubbleras. Kontrollera även tomma rapporter och olika rapportnivåer.

Statisk kontroll är gjord, men VBA-kompilering, Excel-beteende och faktisk
tidsvinst behöver verifieras i Excel; Excel finns inte i molnmiljön.

## Rättning av text och formler vid blockskrivning

Blockskrivningen lägger inte längre till apostrofer i textvärden. Text-ID:n och
littera skyddas med textformat före Value2-skrivning; formler skrivs separat
med General-format och kontrolleras med HasFormula. Mallformat sätts därefter.
Den återanvända sammanställningen återställer också tidigare dolda rader.
Tidsloggen listar huvudkategorier per blad för att felsöka saknade kategorier.

Saknade naturstens-/granitkategorier har inte kunnat reproduceras utan den
aktuella arbetsboken. Kontrollera dem efter uppdatering; kategoriloggen visar
vilka huvudkategorier koden faktiskt genererar. Regelprioriteringen är oförändrad.

## Exempeldata i Rapportmall får inte följa med

Rapportmall används nu enbart för utseende, kolumnbredder, radhöjder och
sammanslagningar. Dess exempelvärden och exempelformler överförs aldrig.
Tidigare kunde till exempel 120 från C12 ligga kvar i m²-kolumnen när en
rapportgrupp använde lpm, och exempelbildtext kunde ligga kvar utan foton.
Alla rapportvärden och formler kommer nu från rapportgenereringen.

Uttrycklig enhet i Regler gäller först. För Auto väljs lpm för exakt SMG2,
därefter m² för CLASS som innehåller Hårdgjort. Övriga Auto-objekt använder
längd K före area J och annars antal. SMG20 omfattas inte av SMG2-undantaget. Klassificeringsreglernas högsta prioritet gäller även materialnamn
som innehåller flera matchande ord, exempelvis smågatsten i betong.

## Infällda detaljer och rubrikordning

Efter all formatering fälls rapportens objektrader ihop med Outline.ShowLevels
RowLevels:=1. Rubriker och grupprader visas; plusknapparna öppnar objekten.
Detta görs sist eftersom tilldelning av radhöjd kan visa tidigare dolda rader.

Ordning i kolumn J styr rubrikernas placering; Prioritet i A väljer matchande
regel. Sorteringen håller nu ihop varje huvudkategori, därefter underkategori
och tredje rubriknivå. Varje rubrik får det lägsta ordningstalet bland sina
matchade regler i hela rådata, så samma ordning används mellan delområden.
Naturstensregler med Ordning 10–15 kommer före Betong 20, Murar/Kantstöd 25–27
och Räcken 30–32. Inom samma rubrik sorteras littera och material alfabetiskt.
Sparprocent ändrar inte rubrikordningen; materialslag innebär inte automatiskt
att objektet ska sparas. Ändringar i Regler kräver UppdateraRapporter.

## Skick och utrustning

Reglernas Sökfält stöder nu SKICK (Rådata kolumn D), utöver Materialtyp,
CLASS och LITTRERING. Sökvärdet matchas som deltext utan skiftlägeskänslighet.
Högsta prioritet och lokala undantag gäller som tidigare. En generell SKICK-regel
kan matcha alla material; den är inte automatiskt begränsad till utrustning.

Objekt vars vinnande regel har Rubriknivå 1 UTRUSTNING får automatiskt:
- Nivå 2 SPARAS för sparandel >0, KASSERAS för uttryckligen 0, EJ BEDÖMT för tomt.
- Nivå 3 objektets skick, eller SKICK EJ ANGIVET om det är tomt.
Ordningen är SPARAS, KASSERAS, EJ BEDÖMT. Dessa nivåer ersätter regelns egna
nivå 2 och 3 för UTRUSTNING. Välj tre rubriknivåer i Rapportval för att se skickrubriker.
Andel utanför 0–100 % eller icke-numerisk andel stoppar genereringen med radnummer.

Gruppformlerna avgränsas även på skick och sparstatus så att kassering inte
summeras med sparade objekt. En delvis sparad post hamnar i SPARAS.
Kör UppdateraRapporter efter ändring av skick eller sparstatus för att
flytta objekt, uppdatera rubriker och detaljrader. Gamla rubriker kan inte
flyttas av enbart formler. Skickrubriker skapas inte automatiskt för övriga
huvudkategorier; där kan egna SKICK-regler användas.

När en aktiv SKICK-regel finns hålls olika skick isär även i övriga grupper.
Det förhindrar att en generell materialgrupp summerar samma objekt som en
särskild skickgrupp. Ta bort/inaktivera SKICK-reglerna om denna uppdelning
inte önskas; automatisk skickuppdelning under UTRUSTNING finns alltid kvar.
