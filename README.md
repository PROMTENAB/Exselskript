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
formler och skrivs med ett sammanhängande Formula-anrop. Text och numeriska
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
