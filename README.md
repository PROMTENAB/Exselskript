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
