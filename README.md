# AttachmentSaver
Ms Outlook Add-in to automatically save attachments
Ms Outlook invoegtoepassing om automatisch bijlagen op te slaan


Instructie (NL):

1 Installatie

Zorg dat Outlook uitstaat en run AttachmentSaver (.VSTO) om het programma te installeren.
Als het goed is is de applicatie geinstalleerd, en zie je als je Outlook weer opstart een nieuw tabblad bovenin: “Bijlage-import”

2 Opslagprofielen

Als je op de knop “Opslagprofielen” drukt verschijnt er een venster waarin je zogenaamde profielen kunt configureren. In een profiel kun je instellen dat je bepaalde typen bijlagen van bepaalde afzenders automatisch ergens kunt wegschrijven.

Je geeft een profiel een naam.

In het veld “Verzender” vul je het e-mail adres van de verzender waarvan je bijlagen wilt opslaan.

In het veld “Bestandsnaam-filter (RegEx)” kun je een Regular Expression-filter instellen voor welke bestandsnamen je wilt opslaan. Dit is niet verplicht, en als je dit leeglaat slaat hij alle bestandsnamen op van de verzender. 
Voor meer informatie over Regular Expressions, zie: https://regexr.com/.

In het veld “Bestandsextensies” kun je specificeren welke typen bestanden je op wilt slaan. Als je er geen één specificeert, slaat hij alle soorten bestanden op, dus ook deze instelling is niet verplicht.

In het veld “Locatie” zet je de opslag-directory waar je de bijlagen op wil slaan.

NB. Als je verschillende typen bijlagen van dezelfde afzender op andere locaties wil opslaan, kun je uiteraard meerdere profielen creëren met diezelfde Verzender. Of uiteraard meerdere profielen met dezelfde Locatie, enz.

Verder wijst de rest zich vanzelf. Je drukt op “Opslaan” om het profiel op te slaan, en je kunt later weer op het profiel in de linker-lijst klikken om deze te bewerken.

3 Bijlagen importeren

Als je op de knop “Bijlagen importeren” drukt, importeert hij de bijlagen van mails sinds de laatst uitgevoerde datum zoals gespecificeerd in de opslagprofielen.

Je kunt er ook voor kiezen om hem automatisch uit te laten voeren met de linker knop. Dan voert hij het automatisch uit bij het opstarten van Outlook, en check hij een nieuwe mail als je er een ontvangt.

In het rechtergedeelte kun je specificaties zien over de import.
Als je hierop klikt, kun je een datum instellen. Je kunt de ingestelde datum waarop hij voor het laatst heeft geïmporteerd hier veranderen, zodat hij bij de volgende keer uitvoeren alles vanaf jouw datum op zal halen.

Als je op “View Log” klikt in het rechtergedeelte, kun je alle resultaten bekijken. Zo kun je controleren of alles zoals bedoeld is opgeslagen (en hoe, waar), of dat er iets mis is gegaan.


