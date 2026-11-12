// ─────────────────────────────────────────────────────────────
//  Practicum "Vertraagd!" — Informatiekunde, Universiteit Utrecht
//
//  Genereren (vanuit de map practicum-vertraagd):  ./maak_pdfs.sh
//  Of los:
//    typst compile --font-path fonts --input excel=nl typst/practicum.typ practicum_leerling_nl-excel.pdf
//    typst compile --font-path fonts --input excel=en typst/practicum.typ practicum_leerling_en-excel.pdf
//    typst compile --font-path fonts --input docent=true typst/practicum.typ practicum_begeleiders.pdf
//
//  De tekst is altijd Nederlands; alleen Excel-termen (menu's, opties, functies) volgen de
//  taal van Excel. De begeleidersversie toont standaard beide (NL / EN).
//
//  Alle getallen komen uit antwoorden.json (gemaakt door scripts/bouw_dataset.py).
// ─────────────────────────────────────────────────────────────

#let docent = sys.inputs.at("docent", default: "false") == "true"
#let excel = sys.inputs.at("excel", default: if docent { "beide" } else { "nl" })
#assert(excel in ("nl", "en", "beide"), message: "excel moet nl, en of beide zijn")
// Excel-term in de taal van Excel: xl("Invoegen", "Insert")
#let xl(nl, en) = if excel == "nl" { nl } else if excel == "en" { en } else [#nl / #en]
#let fn(nl, en) = xl(raw(nl), raw(en))
#let a = json("antwoorden.json")

// ── Huisstijl UU (uu.nl/organisatie/huisstijl) ───────────────
// Primair: geel, zwart, wit; rood alleen als accent. Secundaire kleuren niet voor tekstkaders.
#let uu-geel = rgb("#FFCD00")
#let uu-rood = rgb("#C00A35")
#let zwart = rgb("#000000")
#let geel-licht = uu-geel.lighten(80%)
#let grijs-licht = rgb("#F2F2F2")
#let grijs = rgb("#6E6E6E")
// Merriweather voor tekst, Open Sans voor koppen; Arial als ze ontbreken.
#let kop-font = ("Open Sans", "Arial")
#let tekst-font = ("Merriweather", "Arial")

#let sans(..args) = text(font: kop-font, ..args)

// Getallen in Nederlandse notatie: 1.500 en 41,1
#let nl(x) = {
  let delen = str(x).split(".")
  let heel = delen.at(0)
  let neg = heel.starts-with("-")
  if neg { heel = heel.slice(1) }
  let groepen = ()
  while heel.len() > 3 {
    groepen.insert(0, heel.slice(heel.len() - 3))
    heel = heel.slice(0, heel.len() - 3)
  }
  groepen.insert(0, heel)
  let uit = (if neg { "-" } else { "" }) + groepen.join(".")
  if delen.len() > 1 { uit + "," + delen.at(1) } else { uit }
}

#let vv-naam = (ATL: "Atlanta", ORD: "Chicago O'Hare", DEN: "Denver", LAX: "Los Angeles", PHX: "Phoenix", JFK: "New York JFK")
#let an(vv) = a.analyse.find(r => r.vv == vv)
#let kpi(vv) = a.kpi.find(k => k.vv == vv)
#let uur(u) = a.per_uur.find(r => r.uur == u)
#let max-by(arr, f) = arr.fold(arr.first(), (m, x) => if f(x) > f(m) { x } else { m })

// ── Pagina-opmaak ────────────────────────────────────────────
#set document(
  title: if docent { "Vertraagd! — handleiding en nakijkmodel" } else { "Vertraagd! — practicum" },
  author: "Informatiekunde, Universiteit Utrecht",
)
#set page(
  paper: "a4",
  margin: (x: 2cm, top: 3cm, bottom: 2cm),
  header: context {
    if counter(page).get().first() > 1 {
      grid(
        columns: (auto, 1fr),
        align: (left + horizon, right + horizon),
        image("beeld/logo-uu-2021.png", height: 1.5cm),
        sans(8pt, fill: grijs)[
          Practicum Informatiekunde · Vertraagd!
          #if docent [ · #text(fill: uu-rood, weight: "bold")[BEGELEIDERS]]
        ],
      )
      v(-4pt)
      line(length: 100%, stroke: 1.5pt + uu-geel)
    }
  },
  footer: context {
    if counter(page).get().first() > 1 {
      align(center, sans(8pt, fill: grijs, counter(page).display("1")))
    }
  },
)
#set text(font: tekst-font, size: 9.5pt, lang: "nl")
#set par(leading: 0.75em, spacing: 1.1em)
#set list(marker: sans(fill: zwart)[•])
#set enum(numbering: n => sans(weight: "bold")[#n.])
#show heading: set text(font: kop-font)
#show heading.where(level: 1): it => {
  pagebreak(weak: true)
  block(below: 1em)[
    #text(19pt, weight: "bold", it.body)
    #v(-0.6em)
    #box(width: 2.2cm, height: 4pt, fill: uu-geel)
  ]
}
#show heading.where(level: 2): it => block(above: 1.4em, below: 0.7em, text(12pt, weight: "bold", it.body))
#show heading.where(level: 3): it => block(above: 1.1em, below: 0.6em, text(10pt, weight: "bold", it.body))
#show raw: set text(font: ("Menlo", "DejaVu Sans Mono"), size: 8.5pt)
#show table: set text(font: kop-font, size: 8.5pt)
#show table.cell.where(y: 0): set text(weight: "bold")
#set table(stroke: 0.5pt + rgb("#BBBBBB"), fill: (_, y) => if y == 0 { geel-licht })

// ── Bouwstenen ───────────────────────────────────────────────

#let deel-kop(nr, titel, pad, tijd) = {
  pagebreak(weak: true)
  block(width: 100%, fill: zwart, inset: (x: 14pt, y: 12pt), below: 1.2em)[
    #sans(9pt, weight: "bold", fill: uu-geel)[DEEL #nr · #upper(pad)] #h(1fr) #sans(9pt, fill: white)[± #tijd minuten] \
    #sans(20pt, weight: "bold", fill: white, titel)
  ]
}

#let vraag(id, punten: none, bonus: false, body) = block(above: 1.3em, below: 0.6em, breakable: true)[
  #grid(
    columns: (0.95cm, 1fr),
    if bonus {
      box(stroke: 1pt + zwart, inset: (x: 4pt, y: 2pt), sans(8.5pt, weight: "bold", id))
    } else {
      box(fill: zwart, inset: (x: 5pt, y: 3pt), sans(8.5pt, weight: "bold", fill: white, id))
    },
    [
      #if bonus [#sans(8pt, weight: "bold", fill: uu-rood)[BONUS — alleen als je tijd over hebt] \ ]
      #body
      #if punten != none and docent [#h(1fr) #sans(7.5pt, fill: grijs)[(#if bonus [bonus, ]#punten pt)]]
    ],
  )
]

// Antwoordruimte voor leerlingen, nakijkmodel voor begeleiders
#let antwoord(regels: 2, model) = {
  if docent {
    block(width: 100%, fill: geel-licht, stroke: (left: 3pt + uu-geel), inset: 10pt, below: 1em)[
      #sans(7.5pt, weight: "bold")[NAKIJKMODEL] \
      #model
    ]
  } else {
    for _ in range(regels) {
      block(above: 0pt, below: 0pt, height: 0.9cm, width: 100%,
        line(start: (0pt, 0.85cm), length: 100%, stroke: 0.4pt + rgb("#AAAAAA")))
    }
    v(0.4em)
  }
}

#let kader(titel: none, body) = block(width: 100%, fill: grijs-licht, stroke: (left: 3pt + uu-geel), inset: 10pt)[
  #if titel != none [#sans(weight: "bold", titel) ]#body
]
#let tip(body) = kader(titel: "Tip:", body)

// Toont dezelfde formule voor Nederlandse en Engelse Excel (andere functienamen).
// Het scheidingsteken blijft ; : dat hangt af van de regio-instelling van de laptop, niet van
// de taal van Excel, en is op Nederlandse laptops ook in Engelse Excel ; (getest op Mac, nov 2026).
#let formule-en(f) = {
  f.replace("AANTALLEN.ALS", "COUNTIFS").replace("AANTAL.ALS", "COUNTIF").replace("GEMIDDELDE.ALS", "AVERAGEIF")
}
#let formule(f) = align(center, block(fill: grijs-licht, inset: (x: 10pt, y: 6pt))[
  #if excel == "nl" { raw(f) } else if excel == "en" { raw(formule-en(f)) } else {
    grid(columns: 2, column-gutter: 8pt, row-gutter: 5pt, align: (right + horizon, left + horizon),
      sans(7.5pt, fill: grijs)[NL], raw(f),
      sans(7.5pt, fill: grijs)[EN], raw(formule-en(f)),
    )
  }
])

// Menupad; een stap is een tekst (gelijk in beide talen) of ("Nederlands", "English")
#let menu(..pad) = pad.pos().map(p => if type(p) == array { xl(..p) } else { p }).map(p => box(fill: grijs-licht, stroke: 0.5pt + rgb("#AAAAAA"), inset: (x: 4pt, y: 1.5pt), outset: (y: 1pt), sans(8pt, p))).join(sans[ › ])

#let terugblik(body) = block(width: 100%, stroke: 1pt + zwart, inset: 12pt, above: 1.5em)[
  #sans(weight: "bold")[Wat heb je net gedaan?] #body
]

// Na elk deel: iedereen stopt en kijkt naar het scherm
#let stop-moment() = block(width: 100%, fill: uu-geel, inset: (x: 14pt, y: 10pt), above: 1.2em)[
  #sans(11pt, weight: "bold")[STOP — klassikaal moment] #h(1fr) #sans(9pt)[± 3 minuten] \
  #sans(9pt)[Kijk naar het scherm. Klaar voordat de begeleider begint? Werk door aan de bonusvragen.]
]

#let vakje = box(width: 0.75em, height: 0.75em, stroke: 0.8pt + zwart, baseline: 0.05em)

// Spreektekst voor begeleiders
#let praatje(body) = block(width: 100%, stroke: (left: 3pt + uu-geel), inset: (left: 12pt, y: 4pt))[
  #set text(style: "italic")
  #body
]

// ═════════════════════════════════════════════════════════════
//  Voorblad
// ═════════════════════════════════════════════════════════════

#page(header: none, footer: none, margin: 0pt)[
  #place(top + left, dx: 0pt, dy: 0pt, image("beeld/logo-uu-2021.png", width: 7.2cm))
  #place(top, dy: 2.9cm, image("beeld/KBG.jpg", width: 100%, height: 6.2cm, fit: "cover"))
  #place(top, dy: 9.1cm, rect(width: 100%, height: 0.35cm, fill: uu-geel))
  #place(top + left, dx: 2cm, dy: 10.6cm, block(width: 17cm)[
    #sans(10pt, weight: "bold")[PRACTICUM INFORMATIEKUNDE]
    #v(0.1em)
    #sans(48pt, weight: "bold")[Vertraagd!]
    #v(-0.2em)
    #sans(14pt)[Wat #nl(a.aantal) vertraagde vluchten ons vertellen \ over vliegvelden, maatschappijen en mensen]
    #if docent {
      v(0.6em)
      box(fill: zwart, inset: 8pt, sans(11pt, weight: "bold", fill: uu-geel)[Handleiding en nakijkmodel voor begeleiders])
    }
    #v(1.4em)
    #grid(
      columns: (1fr, 1fr, 1fr),
      column-gutter: 10pt,
      ..(("1", "Kennis en data analytics", "Filteren, sorteren en zelf formules schrijven"),
         ("2", "Mens-technologie interactie", "Kleurcodes en grafieken die je in één oogopslag snapt"),
         ("3", "Informatiesystemen", "Van cijfers naar kernwaarden en een advies")).map(((n, t, d)) =>
        block(width: 100%, height: 2.7cm, fill: grijs-licht, stroke: (top: 4pt + uu-geel), inset: 10pt)[
          #sans(8pt, fill: grijs)[DEEL #n] \
          #sans(weight: "bold", t) \ #v(0.1em) #text(8.5pt, d)
        ])
    )
    #v(1.2em)
    #set text(font: kop-font, size: 9.5pt)
    #grid(
      columns: (auto, 1fr), row-gutter: 7pt, column-gutter: 12pt,
      [*Tijd*], [± 65 minuten: drie delen van een kwartier, na elk deel een korte toelichting],
      [*Nodig*], [Laptop met Excel en het bestand `vluchten_practicum.xlsx`],
      [*Excel*], [#(nl: "Deze versie is voor Nederlandstalige Excel", en: "Deze versie is voor Engelstalige Excel (English)", beide: "Excel-termen staan er in het Nederlands én Engels (NL / EN)").at(excel)],
      [*Samenwerken*], [In tweetallen. Eén bestuurt de muis, de ander leest voor. Wissel per deel!],
      ..if not docent { ([*Namen*], box(width: 1fr, line(length: 100%, stroke: 0.4pt + grijs))) },
    )
  ])
]

// ═════════════════════════════════════════════════════════════
//  Handleiding voor begeleiders (alleen in de docentversie)
// ═════════════════════════════════════════════════════════════

#if docent [
  = Handleiding voor begeleiders

  *Doel.* Scholieren proeven in ruim een uur aan alle drie de studiepaden van Informatiekunde, aan de hand van één echte dataset. Ze hoeven vooraf niets te kunnen: alles staat stap voor stap in de opdrachten. Het gaat niet om goed of fout, maar om de "aha": data heeft regels (Kennis en data analytics), vormgeving bepaalt wat mensen zien (Mens-technologie interactie), en welk getal je kiest hangt af van wie het leest (Informatiesystemen). Na elk deel volgt een kort klassikaal moment waarin je laat zien dat dit maar een klein stukje is van wat je in dat studiepad leert.

  #table(
    columns: (auto, auto, 1fr),
    table.header([Blok], [Duur], [Wat er gebeurt]),
    [Start], [± 5 min], [Instructiepraatje (dia 1–4), bestand openen],
    [Deel 1 — Kennis en data analytics], [± 15 min], [Filteren, sorteren, een verborgen regel ontdekken, eigen formules op tabblad _Analyse_ (dia 5 op het scherm)],
    [_Klassikaal 1_], [± 3 min], [Terugblik + studiepad Kennis en data analytics (dia 6–7)],
    [Deel 2 — Mens-technologie interactie], [± 15 min], [Slechte grafiek beoordelen, kleurcodes (voorwaardelijke opmaak), eigen grafiek, peer review (dia 8)],
    [_Klassikaal 2_], [± 3 min], [Terugblik + studiepad Mens-technologie interactie (dia 9–10)],
    [Deel 3 — Informatiesystemen], [± 15 min], [Oorzaken via draaitabel, denkvraag over de data, mini-dashboard voor een stakeholder (dia 11)],
    [_Klassikaal 3_], [± 3 min], [Terugblik + studiepad Informatiesystemen (dia 12–13)],
    [Afsluiting], [± 5 min], [Twee tweetallen pitchen hun mini-dashboard, opbouw van de studie, welk pad past bij jou (dia 14–17)],
  )

  Ga op tijd door naar het volgende deel, ook als niet iedereen klaar is. Deel 2 en 3 gebruiken wel het tabblad _Analyse_ uit deel 1 (vraag 1.5). Is een tweetal daar niet aan toegekomen, help ze dan die tabel even snel te vullen, of laat ze in deel 2 en 3 de getallen van een buurtweetal gebruiken. Vragen met het label *BONUS* zijn voor snelle tweetallen en tellen niet mee.

  == Klassikale momenten
  Na elk deel staat in de leerlingversie een gele balk *STOP — klassikaal moment*. Zet dan de presentatie `presentatie_vertraagd.pptx` op het scherm. Elk moment heeft twee dia's:
  - *Terugblik*: wat hebben ze net gedaan, en wat is het antwoord op de lastigste vraag?
  - *Studiepad*: dit was een klein proefje; zo ziet het echte studiepad eruit (vakken in jaar 2, vervolg na de bachelor).
  De spreektekst staat in de *notities* bij elke dia. Houd het kort (± 3 minuten): de scholieren moeten vooral zelf doen. Wie nog bezig is, mag na het moment het deel afmaken of doorgaan naar het volgende deel.

  == Voorbereiding
  - Zet `vluchten_practicum.xlsx` op elke laptop en open het één keer: krijg je een melding over "beveiligde weergave", klik dan op _#xl("Bewerken inschakelen", "Enable Editing")_.
  - Controleer of de *statusbalk* onderin #xl("Gemiddelde, Aantal en Som", "Average, Count en Sum") laat zien (rechtermuisklik op de statusbalk om ze aan te zetten).
  - Print de leerlingversie: één per tweetal. Er zijn *twee leerlingversies*: `practicum_leerling_nl-excel.pdf` en `practicum_leerling_en-excel.pdf`. De opdrachten zijn hetzelfde; alleen de menunamen en formules verschillen. Kijk per laptop in welke taal Excel staat (staat er bovenin _Invoegen_ of _Insert_?) en geef de passende versie.
  - Zorg voor een beamer met de presentatie `presentatie_vertraagd.pptx`, Excel en een zichtbare timer (bijvoorbeeld op je telefoon of via een website).
  - Doe het practicum zelf één keer door. Dan weet je waar het lastig wordt (meestal: de eerste formule en de draaitabel).
  - In dit document staan Excel-termen in beide talen (NL / EN). Een overzicht staat in de vertaaltabel achterin.

  == Instructiepraatje (± 4 minuten)
  Een voorbeeld, maak het vooral je eigen. Tussen haken staat wat je doet.

  #praatje[
    Welkom! Ik ben [naam], en ik [studeer / werk bij] Informatiekunde hier in Utrecht.

    Even een vraag: wie heeft weleens op Schiphol gezeten terwijl er "vertraagd" achter je vlucht stond? [handen] Vervelend, hè? Voor jou is het een uurtje wachten. Maar voor een vliegveld of een luchtvaartmaatschappij kost elke minuut vertraging heel veel geld.

    Vandaag zijn jullie geen scholieren maar data-adviseurs bij DataVlucht, een adviesbureau. Jullie hebben een bestand gekregen met #nl(a.aantal) echte vertraagde vluchten uit Amerika. De klant wil weten: waar zit de vertraging, waar komt het door, en wat moeten we eraan doen?

    Om die vraag te beantwoorden heb je drie soorten mensen nodig, en dat zijn precies de drie studiepaden van onze opleiding. [dia 3]
    Eén: iemand die snapt wat er écht in de data staat. Dat is *Kennis en data analytics*.
    Twee: iemand die het zo kan laten zien dat mensen het in één oogopslag snappen. Dat is *Mens-technologie interactie*.
    En drie: iemand die bepaalt welke cijfers belangrijk zijn en welk advies daaruit volgt. Dat is *Informatiesystemen*.
    Jullie proberen vandaag alle drie, steeds een kwartier. Na elk deel stoppen we even samen, en vertel ik wat je in dat studiepad écht gaat doen.

    Hoe werkt het? Jullie werken in tweetallen. Eén van jullie heeft de muis, de ander leest de opdracht voor en denkt mee. Na elk deel wisselen jullie. [laat de timer zien] Na een kwartier zeg ik "door", ook als je nog niet klaar bent: dat is helemaal niet erg. Ben je juist snel? Er staan bonusvragen in.

    Je hoeft geen Excel te kunnen: alles staat stap voor stap op het blad. Kom je er niet uit? Steek je hand op, daar zijn wij voor. Aan het eind laten een paar tweetallen hun advies zien. Open nu het bestand en begin bij "De casus". Succes!
  ]

  == Korte aftrap per deel (± 30 seconden)
  / Deel 1: "Data-experts zijn een beetje detectives. Voordat je iets concludeert, moet je weten wat er in je data zit, en vooral wat er níet in zit. En aan het eind schrijven jullie je eerste formules." [Laat op de beamer zien waar de statusbalk zit.]
  / Deel 2: "Jullie hebben nu getallen. Maar een directeur gaat geen tabellen lezen. Hoe zorg je dat iemand in drie seconden ziet wat er aan de hand is? Denk aan het dashboard van je smartwatch of je bank-app. Daar gaat Mens-technologie interactie over: ontwerpen voor hoe mensen kijken."
  / Deel 3: "Nu wordt het zakelijk. Een directeur, een maatschappij en een app-bouwer willen allemaal iets anders weten. Jullie kiezen de juiste getallen en geven een advies." [Doe stap 1–4 van de draaitabel in 3.1 voor op de beamer: dat scheelt veel vragen.]

  == Afsluiting (± 5 minuten)
  - Laat twee tweetallen met een verschillende stakeholder hun mini-dashboard in één minuut pitchen.
  - Vraag: _"Welk deel vonden jullie het leukst? Steek je hand op bij deel 1... 2... 3..."_ Koppel dat aan de studiepaden (dia 15): "Wie deel 1 het leukst vond, zou zich thuis voelen bij Kennis en data analytics", enzovoort.
  - Sluit af met de kernboodschap (dia 14): _"In het eerste jaar maak je kennis met alle drie. In het tweede jaar kies je een studiepad, en je kunt ze ook combineren."_

  == Over de data
  *Bron.* US Bureau of Transportation Statistics (via de Kaggle-dataset "Airline Delay Causes"), jaar 2008. Het origineel bevat 1,9 miljoen vluchten.

  *Bewerking.* Gefilterd op vertrek vanaf ATL, ORD, DEN, LAX, PHX en JFK. Geannuleerde en omgeleide vluchten zijn verwijderd. Daarna zijn er per vliegveld *precies #nl(a.per_vliegveld) vluchten* willekeurig getrokken (seed #a.seed). Mijlen zijn omgerekend naar kilometers en kolomnamen zijn vertaald.

  *Waarom evenveel vluchten per vliegveld?* De bron bevat alleen *binnenlandse* vluchten van Amerikaanse maatschappijen. JFK is een enorm vliegveld, maar vooral voor internationale vluchten. Die zitten niet in deze data. In het origineel heeft JFK daardoor maar ± 34.000 vertraagde vluchten, tegen ± 131.000 voor ATL. Bij een gewone steekproef zou JFK maar een paar honderd rijen krijgen. Dan lijkt JFK "klein", en zijn gemiddelden over JFK minder betrouwbaar. Met gelijke aantallen vergelijk je de vliegvelden eerlijk. Het nadeel: het *aantal* vluchten per vliegveld zegt nu niets meer. Dat is precies de denkvraag in 3.2.

  *De grote valkuil.* De brondata bevat alleen vluchten die minstens 6 minuten te laat vertrokken. Een percentage "op tijd" kun je dus niet uitrekenen. Dit komt terug in vraag 1.1 en 3.2.

  == Veelgemaakte fouten en hoe je helpt
  - *Formule geeft `#NAAM?`*: tikfout in de functienaam of in `Vluchten[...]`. Laat ze de formule opnieuw beginnen en de kolomnaam kiezen uit het lijstje dat Excel toont.
  - *Formule wordt tekst*: de cel begint niet met `=`, of er staat een spatie vóór het `=`-teken.
  - *Verkeerd scheidingsteken*: welk teken Excel tussen de onderdelen van een formule wil (`;` of `,`), hangt af van de *regio-instelling van de laptop*, niet van de taal van Excel. Op Nederlandse laptops is het vrijwel altijd `;`, ook in Engelse Excel (getest). Daarom staat overal `;`. Foutmelding? Laat ze `;` vervangen door `,`.
  - *`">=100"` zonder aanhalingstekens*: dan geeft Excel een foutmelding.
  - *Doortrekken lukt niet*: wijs het kleine vierkantje rechtsonder in de geselecteerde cel aan (de muis wordt een dun kruisje) en sleep naar beneden.
  - *Filters vergeten te wissen*: getallen kloppen niet meer met het nakijkmodel. Check: staat er een trechtertje in een kolomkop? Formules en draaitabellen rekenen trouwens altijd met álle rijen, ook als er een filter aanstaat.
  - *De kolomkop wordt ook geel of rood* bij vraag 2.2: Excel vindt tekst "groter dan 99". Dat is niet erg; het verdwijnt niet vanzelf, maar het klopt verder allemaal.
  - *Voorwaardelijke opmaak: alles rood of alles geel*: de verkeerde kolom geselecteerd, of de regels staan in de verkeerde volgorde. Kijk bij #menu(("Voorwaardelijke opmaak", "Conditional Formatting"), ("Regels beheren", "Manage Rules")).
  - *Draaitabel toont "#xl("Som", "Sum")" in plaats van "#xl("Gemiddelde", "Average")"*: Excel kiest standaard #xl("Som", "Sum"). Laat ze _#xl("Waardeveldinstellingen", "Value Field Settings")_ zoeken.
  - *Kleine aantallen*: een gemiddelde over een paar vluchten is geen inzicht. Dit is de belangrijkste les die in alle drie de studiepaden terugkomt (zie bonusvraag 2.6).

  == Puntentelling (optioneel, bijvoorbeeld voor een wedstrijdje)
  #table(
    columns: (1fr, auto, auto),
    table.header([Deel], [Punten], [Bonus]),
    [Deel 1 — Kennis en data analytics], [13], [+2 (1.6)],
    [Deel 2 — Mens-technologie interactie], [14], [+3 (2.6)],
    [Deel 3 — Informatiesystemen], [10], [+3 (3.4)],
    [*Totaal*], [*37*], [+8],
  )

  == Excel in het Engels
  #table(
    columns: (1fr, 1fr),
    table.header([Nederlands], [Engels]),
    [Gegevens › Filter / Wissen], [Data › Filter / Clear],
    [Invoegen › Draaitabel], [Insert › PivotTable],
    [Start › Voorwaardelijke opmaak], [Home › Conditional Formatting],
    [Markeringsregels voor cellen], [Highlight Cells Rules],
    [Gegevensbalken / Kleurenschalen], [Data Bars / Color Scales],
    [`AANTAL.ALS`], [`COUNTIF`],
    [`AANTALLEN.ALS`], [`COUNTIFS`],
    [`GEMIDDELDE.ALS`], [`AVERAGEIF`],
    [Scheidingsteken `;`], [Scheidingsteken `,`],
  )
]

// ═════════════════════════════════════════════════════════════
//  De casus
// ═════════════════════════════════════════════════════════════

= De casus
Elke dag vertrekken er in de Verenigde Staten zo'n 20.000 vluchten. Een flink deel daarvan vertrekt te laat. Dat kost reizigers tijd en zenuwen, en luchtvaartmaatschappijen en vliegvelden heel veel geld.

Jullie werken vandaag bij *DataVlucht*, een (verzonnen) adviesbureau. Een klant heeft ons een bestand gestuurd met #nl(a.aantal) vertraagde binnenlandse vluchten uit 2008: #nl(a.per_vliegveld) vanaf elk van deze zes grote Amerikaanse vliegvelden.

#align(center, table(
  columns: 2, inset: (x: 12pt, y: 5pt),
  table.header([Code], [Vliegveld]),
  ..("ATL", "ORD", "DEN", "LAX", "PHX", "JFK").map(c => (raw(c), vv-naam.at(c))).flatten()
))

De vraag van de klant: *"Waar zit de vertraging, waar komt het door, en wat moeten we eraan doen?"*

Om die vraag te beantwoorden heb je drie soorten informatiekundigen nodig, precies de drie studiepaden van de bachelor Informatiekunde in Utrecht:

+ *Kennis en data analytics*: Wat staat er eigenlijk in de data? Klopt het? Wat ontbreekt er?
+ *Mens-technologie interactie*: Hoe laat je mensen in één oogopslag zien wat er aan de hand is?
+ *Informatiesystemen*: Welke cijfers zijn belangrijk voor de klant, en welk besluit volgt daaruit?

== Even wennen aan het bestand
Open `vluchten_practicum.xlsx`. Er zijn vier tabbladen:

- *Lees mij*: wat betekent elke kolom? Lees dit eerst even door.
- *Vluchten*: de data. Elke rij is één vlucht.
- *Analyse*: hier schrijf je straks je eigen formules.
- *Codes*: welke afkorting hoort bij welk vliegveld of welke maatschappij.

#tip[Excel heeft onderaan het scherm een *statusbalk*. Selecteer je een kolom met getallen, dan zie je daar meteen het *#xl("Gemiddelde", "Average")*, het *#xl("Aantal", "Count")* en de *#xl("Som", "Sum")*. Die heb je vandaag vaak nodig!]

// ═════════════════════════════════════════════════════════════
//  DEEL 1 — Kennis en data analytics
// ═════════════════════════════════════════════════════════════

#deel-kop(1, "Data-detective", "Kennis en data analytics", 15)

Voordat je conclusies trekt, moet je weten wat er in je data zit, en vooral wat er *niet* in zit. Data-specialisten zijn een beetje detectives: ze filteren, sorteren en vragen zich steeds af: _klopt dit wel?_ Daarna laten ze de computer het rekenwerk doen met *formules*.

#tip[Het tabblad *Vluchten* is een Excel-*tabel*: in elke kolomkop zit een pijltje. Daarmee kun je *filteren* (alleen bepaalde rijen tonen) en *sorteren*. Filters wissen: #menu(("Gegevens", "Data"), ("Wissen", "Clear")).]

#vraag("1.1", punten: 2)[
  Lees het tabblad *Lees mij*. Welke vluchten zitten er *niet* in deze dataset? Waarom is dat belangrijk als je wilt weten hoe goed een vliegveld presteert?
]
#antwoord(regels: 2)[
  Vluchten die op tijd (of minder dan 6 minuten te laat) vertrokken. Ook internationale vluchten ontbreken. Gevolg: je kunt *geen percentage* "op tijd" of "te laat" uitrekenen. Je ziet alleen hoe erg de vertraging is als het misgaat, niet hoe vaak het misgaat. (1 pt welke vluchten, 1 pt gevolg.)
]

#vraag("1.2", punten: 2)[
  Zet *twee filters tegelijk*: *Van* = JFK en *Maand* = 12 (december). Hoeveel vluchten zie je, en wat is hun *gemiddelde* vertrekvertraging? (Selecteer de kolom en kijk in de statusbalk.)
]
#antwoord(regels: 1)[
  *#a.jfk_dec.aantal vluchten*, gemiddeld *#nl(a.jfk_dec.gem) minuten*. Linksonder staat ook "#xl([#a.jfk_dec.aantal van #nl(a.aantal) records gevonden], [#a.jfk_dec.aantal of #a.aantal records found])". Let op: selecteer je de kolomkop mee, dan telt #xl("Aantal", "Count") één te veel. (1 pt aantal, 1 pt gemiddelde.)
]

#vraag("1.3", punten: 2)[
  Wis de filters. *Sorteer* de kolom *Vertrekvertraging_min* van groot naar klein (pijltje › _#xl("Sorteren van groot naar klein", "Sort Largest to Smallest")_; op een Mac heet die knop _#xl("Aflopend", "Descending")_). Welke vlucht had de grootste vertraging? Noteer het vluchtnummer, de datum, van waar naar waar, en hoeveel *uur* de vertraging was.
]
#antwoord(regels: 2)[
  *#a.max_vertraging.vlucht* (#a.max_vertraging.maatschappij) op #a.max_vertraging.datum, van #a.max_vertraging.van naar #a.max_vertraging.naar: #nl(a.max_vertraging.minuten) minuten = *± #nl(a.max_vertraging.uren) uur*. (1 pt vlucht + datum + route, 1 pt omrekening naar uren.) \
  _Doorvraag:_ fout in de data of echt gebeurd? Het kan echt zijn (bijvoorbeeld een defect, waardoor de vlucht naar de volgende dag is verschoven). Maar een data-expert controleert zulke uitschieters altijd.
]

#vraag("1.4", punten: 3)[
  *Detectivevraag.* Wis eerst de filters. Bij een deel van de vluchten zijn de vijf *Oorzaak*-kolommen leeg. Filter *Oorzaak_weer_min* op _#xl("(Lege cellen)", "(Blanks)")_ en sorteer de kolom *Aankomstvertraging_min*. Wat valt op? Kun je een regel bedenken die verklaart wanneer de oorzaak wél wordt ingevuld?
]
#antwoord(regels: 3)[
  Bij de #nl(a.leeg.aantal) vluchten zonder oorzaak is de aankomstvertraging *hooguit #a.leeg.max_aankomst minuten*. Bij vluchten mét oorzaak is die *minstens #a.leeg.min_aankomst_gevuld minuten*. \
  *Regel:* de oorzaak wordt alleen geregistreerd als een vlucht *15 minuten of meer te laat aankomt*. Dat is de officiële Amerikaanse definitie van een "vertraagde vlucht". \
  Bonus-ontdekking: sommige vluchten hebben een *negatieve* aankomstvertraging (laagste: #a.leeg.min_aankomst min). Ze vertrokken te laat, maar haalden de tijd onderweg in. Dat geldt voor #nl(a.leeg.te_vroeg) vluchten. \
  _Punten:_ 1 pt grens 14/15 gezien, 1 pt regel geformuleerd, 1 pt negatieve waarden of een andere slimme observatie.
]

#vraag("1.5", punten: 4)[
  *Je eigen formules.* Wis alle filters en ga naar het tabblad *Analyse*. Daar vul je de gele cellen, niet met de hand maar met formules. Klik in de cel, typ de formule en druk op Enter. _Zegt Excel "There's a problem with this formula" of "Er is een probleem met deze formule"? Vervang dan elke `;` door een `,` (dat hangt af van de instellingen van je laptop)._

  #enum(
    [*B2: aantal vluchten.* #fn("AANTAL.ALS", "COUNTIF") telt hoe vaak iets voorkomt. Tel hoe vaak "ATL" (cel A2) in de kolom Van staat:
      #formule("=AANTAL.ALS(Vluchten[Van];A2)")],
    [*C2: gemiddelde vertrekvertraging.* #fn("GEMIDDELDE.ALS", "AVERAGEIF") neemt het gemiddelde, maar alleen van de rijen die bij A2 horen:
      #formule("=GEMIDDELDE.ALS(Vluchten[Van];A2;Vluchten[Vertrekvertraging_min])")],
    [*D2: gemiddelde taxitijd.* Nu jij! Schrijf zelf een formule voor het gemiddelde van *Taxi_uit_min*. (Tip: hij lijkt heel erg op C2.)],
    [*E2: aantal vluchten met 100 minuten of meer vertraging.* #fn("AANTALLEN.ALS", "COUNTIFS") (#xl("met -LEN", "met een S erachter")) telt met *twee* voorwaarden:
      #formule("=AANTALLEN.ALS(Vluchten[Van];A2;Vluchten[Vertrekvertraging_min];\">=100\")")],
    [*F2: percentage.* Schrijf zelf een formule die E2 deelt door B2. Excel zet het als percentage neer.],
    [*Doortrekken.* Selecteer B2 t/m F2. Pak het kleine vierkantje rechtsonder en sleep het naar rij 7. Nu staan alle zes de vliegvelden er!],
  )
  a) Welk vliegveld heeft de langste taxitijd, en welk het hoogste percentage vluchten met ≥ 100 min vertraging? \
  b) Je typte de formule maar één keer, maar elke rij geeft een ander antwoord. Hoe kan dat?
]
#antwoord(regels: 3)[
  #table(
    columns: (auto, 1fr, 1fr, 1fr, 1fr, 1fr), align: (left, right, right, right, right, right),
    table.header([Vliegveld], [Aantal], [Gem. vertraging], [Gem. taxitijd], [≥ 100 min], [%]),
    ..a.analyse.map(r => (raw(r.vv), nl(r.n), nl(r.vertr), nl(r.taxi), str(r.ernstig), str(r.pct) + "%")).flatten()
  )
  D2: #formule("=GEMIDDELDE.ALS(Vluchten[Van];A2;Vluchten[Taxi_uit_min])") F2: `=E2/B2` \
  a) Langste taxitijd: *#max-by(a.analyse, r => r.taxi).vv* (#nl(max-by(a.analyse, r => r.taxi).taxi) min). Hoogste percentage ≥ 100 min: *#max-by(a.analyse, r => r.ernstig).vv* (#max-by(a.analyse, r => r.ernstig).pct%), op de voet gevolgd door ORD. \
  b) Excel past de verwijzing *A2* bij het doortrekken automatisch aan naar A3, A4, enzovoort: een *relatieve celverwijzing*. Daarom werkt één formule voor de hele tabel. \
  _Punten:_ 1 pt tabel correct, 1 pt D2 zelf geschreven, 1 pt F2 zelf geschreven, 1 pt uitleg b.
]

#vraag("1.6", punten: 2, bonus: true)[
  Wat als de klant "ernstige vertraging" niet bij 100 maar bij 60 minuten wil laten beginnen? Dan wil je niet alle formules opnieuw typen. In cel *I1* staat de drempel (100). Verander de formule in E2 in:
  #formule("=AANTALLEN.ALS(Vluchten[Van];A2;Vluchten[Vertrekvertraging_min];\">=\"&$I$1)")
  Trek hem door naar beneden en verander I1 in 60. Wat gebeurt er? Waarom staan er *dollartekens* in `$I$1`?
]
#antwoord(regels: 2)[
  Alle aantallen en percentages passen zich direct aan: bij 60 minuten bijvoorbeeld #an("JFK").pct60% voor JFK en #an("PHX").pct60% voor PHX. De `$` maakt de verwijzing *absoluut*: bij doortrekken blijft het I1 (en verschuift het niet naar I2, I3). Zo kun je "wat-als"-vragen beantwoorden, iets wat in het bedrijfsleven dagelijks gebeurt.
]

#terugblik[Je hebt data _verkend_, een _verborgen regel_ ontdekt (vraag 1.4) en je eerste _formules_ geschreven. Zo werkt Kennis en data analytics: je moet weten hoe data tot stand komt voordat je er iets over durft te zeggen. In de opleiding Informatiekunde leer je SQL en Python gebruiken om data te analyseren, op miljoenen rijen tegelijk.]

#stop-moment()

// ═════════════════════════════════════════════════════════════
//  DEEL 2 — Mens-technologie interactie
// ═════════════════════════════════════════════════════════════

#deel-kop(2, "Kleur en grafieken die spreken", "Mens-technologie interactie", 15)

Apps gebruiken vaak *dashboards* om informatie aan hun gebruikers te laten zien. Denk aan het activiteitenoverzicht van je smartwatch, of het overzicht van je uitgaven in je bank-app. Zo'n dashboard werkt met grafieken en kleurcodes. Een goed dashboard laat in drie seconden zien wat er aan de hand is. Een slecht dashboard laat mensen zoeken, of erger: het leidt ze naar de verkeerde conclusie.

In het studiepad Mens-technologie interactie leer je ontwerpen _voor mensen_: welke informatie is belangrijk, wat ziet iemand als eerste, wat begrijpt iemand wel en wat niet? In dit deel ontwerp je de bouwstenen van zo'n dashboard.

#vraag("2.1", punten: 4)[
  Een stagiair heeft de *gemiddelde vertrekvertraging per maand* in een grafiek gezet. Noem minstens *vier* dingen die er mis zijn met deze grafiek.
]
#align(center, image("beeld/slechte_grafiek.png", width: 66%))
#antwoord(regels: 3)[
  1 pt per goed punt, max 4. Mogelijke antwoorden:
  - De *titel* zegt niets ("GRAFIEK"). Een goede titel vertelt de conclusie.
  - Een *taartdiagram* is voor delen van een geheel. Gemiddelden per maand vormen geen geheel, dus de percentages zijn *betekenisloos*.
  - Maanden hebben een *volgorde in de tijd*: daarvoor gebruik je een lijn- of staafgrafiek.
  - *3D, schaduw en losse stukken* vervormen de grootte: stukken vooraan lijken groter.
  - *Regenboogkleuren* betekenen niets, je moet steeds heen en weer kijken naar de legenda. Lastig voor kleurenblinden.
  - Maanden als *nummers* in plaats van namen, en *geen eenheid* (minuten?).
  - Druk lettertype, schreeuwende kleuren, gele achtergrond: slecht leesbaar.
  - De stukken zijn bijna even groot (7–10%): je ziet het verschil tussen maanden niet.
]

#vraag("2.2", punten: 3)[
  *Kleurcodes.* Met 9.000 rijen zie je door de bomen het bos niet meer. Laat Excel de ernstige vertragingen automatisch kleuren:
  + Ga naar *Vluchten*. Selecteer de hele kolom *Vertrekvertraging_min* (klik op de letter boven de kolom).
  + #menu(("Start", "Home"), ("Voorwaardelijke opmaak", "Conditional Formatting"), ("Markeringsregels voor cellen", "Highlight Cells Rules"), ("Groter dan...", "Greater Than...")) › vul *99* in › kies _#xl("Gele opvulling met donkergele tekst", "Yellow Fill with Dark Yellow Text")_.
  + Herhaal, maar nu met *179* en _#xl("Lichtrode opvulling met donkerrode tekst", "Light Red Fill with Dark Red Text")_.
  Scroll eens door de data. a) Gebruik het filterpijltje › _#xl("Filteren op kleur", "Filter by Color")_: hoeveel vluchten zijn *rood*? b) Waarom kies je juist rood voor de ergste vertragingen? En wat is het probleem als je *alleen* op kleur vertrouwt?
]
#antwoord(regels: 3)[
  a) Rood (≥ 180 min): *#a.drie_uur.aantal vluchten*. Geel (100–179 min): *#(a.totaal_ernstig - a.drie_uur.aantal) vluchten*. \
  b) Rood betekent voor de meeste mensen "gevaar / stop": het sluit aan bij wat mensen al kennen (een _conventie_). Maar ± 8% van de mannen is kleurenblind, vaak voor rood/groen. Gebruik daarom ook licht/donker-verschil, vetgedrukte tekst of een icoontje. \
  _Punten:_ 1 pt aantal, 1 pt conventie, 1 pt kleurenblindheid of toegankelijkheid. \
  _Als alles rood wordt:_ de regel voor 179 is als laatste toegevoegd en staat bovenaan, dus die wint. Klopt het nog niet, kijk dan bij #menu(("Regels beheren", "Manage Rules")).
]

#vraag("2.3", punten: 2)[
  Ga naar je tabel op *Analyse*. Selecteer F2 t/m F7 en kies #menu(("Start", "Home"), ("Voorwaardelijke opmaak", "Conditional Formatting"), ("Gegevensbalken", "Data Bars")). Selecteer daarna D2 t/m D7 en kies #menu(("Kleurenschalen", "Color Scales")) (groen-geel-rood). \
  Je hebt nu twee soorten kleurcodes gebruikt: *#xl("gegevensbalken", "Data Bars")* (kolom F) en een *#xl("kleurenschaal", "Color Scale")* (kolom D). Welke vliegvelden springen eruit? En met welke van die twee soorten kun je *preciezer* zien hoe groot het verschil tussen twee vliegvelden is? Waarom?
]
#antwoord(regels: 2)[
  JFK springt er bij allebei uit (en ORD bij het percentage). *Gegevensbalken* zijn preciezer: mensen kunnen *lengtes* veel beter vergelijken dan kleurtinten. Of iets "twee keer zo groot" is, zie je aan een balk wel, maar aan een kleur niet. Kleur is goed om iets op te laten _vallen_, lengte om te _vergelijken_. Dat is een basisregel uit de visualisatieleer. (1 pt vliegveld, 1 pt uitleg.)
]

#kader(titel: "Checklist: vijf regels voor een goede grafiek")[
  #set list(marker: vakje)
  - *Titel = conclusie.* Niet "Taxitijd per vliegveld", maar "Op JFK sta je het langst in de rij".
  - *Juiste vorm.* Tijd (uren, maanden) → *lijn*. Categorieën vergelijken → *staaf*, gesorteerd van groot naar klein.
  - *Assen kloppen.* Aslabels met eenheid (min, km). Bij staven begint de as bij 0.
  - *Kleur met een doel.* Alles grijs, en één accentkleur voor het ene ding dat opvalt.
  - *Weg met rommel.* Geen 3D en geen schaduw. Bij één reeks heb je geen legenda nodig.
]

#vraag("2.4", punten: 3)[
  Maak van je *Analyse*-tabel een grafiek van de *gemiddelde taxitijd* per vliegveld:
  + Selecteer de cellen *A1 t/m A7* (de vliegvelden).
  + Houd *Ctrl* (Windows) of *Cmd* (Mac) ingedrukt en selecteer ook *D1 t/m D7* (de taxitijd).
  + Kies #menu(("Invoegen", "Insert")) › het knopje met de staafjes (#xl("Kolom- of staafdiagram invoegen", "Insert Column or Bar Chart; op een Mac heet het Column")) › onder _#xl("2D-staaf", "2-D Bar")_ de eerste (liggende staven).
  Pas daarna de checklist toe:
  - *Sorteren:* klik in de tabel op een getal in kolom D › #menu(("Start", "Home"), ("Sorteren en filteren", "Sort & Filter"), ("Sorteren van klein naar groot", "Sort Smallest to Largest")). De langste staaf staat nu bovenaan.
  - *Accentkleur:* klik één keer op een staaf (alle staven worden geselecteerd) en dan nog één keer op de staaf van het opvallendste vliegveld (nu is alleen die geselecteerd) › rechtermuisklik › _#xl("Gegevenspunt opmaken...", "Format Data Point...")_ › het verfemmertje › _#xl("Opvulling", "Fill")_ › kies een kleur.
  - *Titel:* klik op de titel van de grafiek en typ je conclusie. \
  _Liever het percentage ≥ 100 min? Selecteer dan in stap 2 *F1 t/m F7*._
]
#antwoord(regels: 1)[
  Goed: een gesorteerde (liggende) staafgrafiek, alle staven grijs behalve *JFK* in een accentkleur, en een titel met de conclusie. Volgorde taxitijd: #a.taxi_per_vliegveld.map(t => t.vv + " " + nl(t.gem)).join(" · "). \
  _Punten:_ 1 pt vorm + gesorteerd, 1 pt accentkleur, 1 pt titel als conclusie.
  #align(center, image("beeld/goed_taxi.png", width: 80%))
]

#vraag("2.5", punten: 2)[
  *Peer review.* Wissel van laptop met het tweetal naast je. Loop hun grafiek langs met de checklist: hoeveel van de 5 vinkjes krijgen ze? Geef ze *één concrete tip*.
]
#antwoord(regels: 1)[
  Beoordeel de kwaliteit van de feedback. Is de tip concreet ("zet de staven op volgorde", "zeg in de titel wat er opvalt") of vaag ("ziet er mooi uit")? Zo worden in het studiepad Mens-technologie interactie ontwerpen getest: laat het iemand anders zien.
]

#vraag("2.6", punten: 3, bonus: true)[
  Maakt het *tijdstip van vertrek* uit? Maak een *draaitabel* (#menu(("Invoegen", "Insert"), ("Draaitabel", "PivotTable")), op een nieuw werkblad): sleep *Vertrekuur* naar _#xl("Rijen", "Rows")_ en *Vertrekvertraging_min* naar _#xl("Waarden", "Values")_ (zet die via _#xl("Waardeveldinstellingen", "Value Field Settings")_ op *#xl("Gemiddelde", "Average")*). Voeg ook het _#xl("Aantal", "Count")_ van *Vlucht_ID* toe. (Veel cijfers achter de komma? Klik op een getal › #menu(("Start", "Home")) › het knopje _#xl("Minder decimalen", "Decrease Decimal")_.) Maak er een *lijngrafiek* van. Wat zie je? En waarom moet je voorzichtig zijn met de uren 0 tot 5?
]
#antwoord(regels: 2)[
  *Hoe later op de dag, hoe meer vertraging*: rond 10:00 gemiddeld ± #nl(uur(10).gem) min, rond 19:00 ± #nl(uur(19).gem) min. Vertraging *stapelt zich op*: een vliegtuig vliegt meerdere vluchten per dag, en is de eerste te laat, dan zijn de volgende dat ook. \
  Uren 0–5: bijna geen vluchten (om 1:00 maar #uur(1).n vlucht!). *Bij zo weinig vluchten zegt een gemiddelde niets.*
  #align(center, image("beeld/goed_uur.png", width: 80%))
]

#terugblik[Je hebt een ontwerp _beoordeeld_, met kleur _informatie zichtbaar gemaakt_, zelf iets _ontworpen_ volgens duidelijke regels en het laten _testen_ door een ander. In het studiepad Mens-technologie interactie doe je dit voor echte apps, websites en dashboards: je onderzoekt hoe mensen informatie waarnemen en met technologie omgaan, en ontwerpt daar gericht voor.]

#stop-moment()

// ═════════════════════════════════════════════════════════════
//  DEEL 3 — Informatiesystemen
// ═════════════════════════════════════════════════════════════

#deel-kop(3, "Van data naar besluit", "Informatiesystemen", 15)

Een directeur heeft geen tijd voor 9.000 rijen of tien grafieken. Die wil weten: _hoe staan we ervoor, en wat moet ik doen?_ In het studiepad Informatiesystemen leer je hoe organisaties informatie en technologie inzetten om beslissingen te nemen. Je vertaalt data naar een paar *kernwaarden* (in het Engels: KPI's, _Key Performance Indicators_) en naar een *advies*. Jullie tabel op *Analyse* is al een prima begin!

#vraag("3.1", punten: 3)[
  *Waar komt de vertraging vandaan?* Maak een *draaitabel*:
  + Klik ergens in de data op *Vluchten* › #menu(("Invoegen", "Insert"), ("Draaitabel", "PivotTable")) › _#xl("Nieuw werkblad", "New Worksheet")_ › OK.
  + Sleep *Van* naar het vak _#xl("Rijen", "Rows")_.
  + Sleep alle vijf de *Oorzaak*-kolommen naar het vak _#xl("Waarden", "Values")_. Excel telt de minuten op: "#xl("Som van ...", "Sum of ...")".
  + Selecteer de getallen in de kolom *#xl("Som van", "Sum of") Oorzaak_luchtverkeer_min* en kies #menu(("Start", "Home"), ("Voorwaardelijke opmaak", "Conditional Formatting"), ("Kleurenschalen", "Color Scales")), net als in deel 2.
  Alle vliegvelden hebben evenveel vluchten in het bestand, dus je mag de minuten eerlijk vergelijken. \
  a) Wat is voor *ORD* (Chicago) de grootste oorzaak (de meeste minuten)? \
  b) Welk vliegveld heeft de meeste minuten vertraging door *luchtverkeer*? \
  c) Welke oorzaak kan een *vliegveld* zelf het best aanpakken, en welke niet? Leg uit.
]
#antwoord(regels: 3)[
  #let o = kpi("ORD").oorzaak_min
  #let op = kpi("ORD").oorzaak_pct
  #let lv = max-by(a.kpi, k => k.oorzaak_min.luchtverkeer)
  a) ORD: *laat binnengekomen vliegtuig* (#nl(o.laat_vliegtuig) min, #nl(op.laat_vliegtuig)% van alle oorzaak-minuten), daarna maatschappij (#nl(o.maatschappij) min). Dat is het sneeuwbaleffect: één te late vlucht maakt de volgende ook te laat. \
  b) *#lv.vv*: #nl(lv.oorzaak_min.luchtverkeer) min door luchtverkeer (#nl(lv.oorzaak_pct.luchtverkeer)% van de JFK-oorzaken, het hoogste aandeel van alle zes). Samen met de langste taxitijd betekent dit: op JFK is het vliegveld zélf echt een probleem. \
  c) Een vliegveld kan vooral *luchtverkeer* aanpakken (capaciteit: start- en landingsbanen, gates, spreiding van piekuren), en een beetje *beveiliging*. *Maatschappij* (onderhoud, bemanning) en *laat vliegtuig* liggen bij de luchtvaartmaatschappij. Aan *weer* kan niemand iets doen. \
  _Punten:_ 1 pt a, 1 pt b, 1 pt c. \
  _Let op:_ "#xl("% van rijtotaal", "% of Row Total")" werkt hier *niet*: Excel rekent dat per waardeveld uit, dus elke kolom wordt 100%. Percentages tussen oorzaken kun je in een draaitabel niet maken.
]

#vraag("3.2", punten: 2)[
  *Denkvraag.* In ons bestand heeft elk vliegveld precies #nl(a.per_vliegveld) vluchten. Kun je met dit bestand zeggen welk vliegveld de *meeste* vertraagde vluchten heeft? Waarom zou de maker van het bestand voor gelijke aantallen hebben gekozen?
]
#antwoord(regels: 3)[
  *Nee.* De #nl(a.per_vliegveld) is een keuze van de maker, geen eigenschap van de vliegvelden. Om te weten hoe vaak het misgaat, heb je het *totaal aantal vluchten* nodig, ook de vluchten die op tijd waren (en die ontbreken, zie 1.1). \
  Gelijke aantallen maken *eerlijk vergelijken* mogelijk: elk gemiddelde en percentage is even betrouwbaar. In de bron heeft JFK veel minder rijen (alleen binnenlandse vluchten), dus bij een gewone steekproef had JFK maar een paar honderd vluchten gekregen. \
  _Kernles:_ data wordt gemaakt door mensen die keuzes maken. Vraag altijd: wat zit erin, en waarom? \
  _Punten:_ 1 pt "nee + totaal nodig", 1 pt "eerlijk vergelijken".
]

#vraag("3.3", punten: 5)[
  *Eindproduct: jouw mini-dashboard (KPI-kaart).* Kies één stakeholder hieronder. Maak voor die persoon een klein dashboard (op een nieuw tabblad of op papier) met:
  - *3 kernwaarden* (KPI's) die voor díe persoon belangrijk zijn, met het getal erbij (gebruik je tabel op *Analyse*);
  - *1 grafiek* (die uit deel 2);
  - *1 advies* van maximaal twee zinnen.
  Presenteer je dashboard in *één minuut* aan een ander tweetal, alsof zij de stakeholder zijn.
]
#grid(
  columns: (1fr, 1fr, 1fr), column-gutter: 8pt,
  ..(
    ("A", "Operationeel directeur van JFK", "\"Ons vliegveld staat bekend als traag. Wat kunnen wíj daaraan doen?\""),
    ("B", "Hoofd operatie van een luchtvaartmaatschappij", "\"Op welke vliegvelden moeten we extra buffer inplannen?\""),
    ("C", "Productmanager van een reisapp", "\"Wat moeten we reizigers laten zien als ze een vlucht boeken?\""),
  ).map(((l, w, q)) => block(width: 100%, height: 3.1cm, fill: grijs-licht, stroke: (top: 3pt + uu-geel), inset: 9pt)[
    #sans(weight: "bold")[#l. #w] \ #v(0.2em) #text(8.5pt, style: "italic", q)
  ])
)
#antwoord(regels: 0)[
  *Beoordeling (5 pt):* 2 pt KPI's passen bij de stakeholder · 1 pt grafiek voldoet aan de checklist · 1 pt advies volgt logisch uit de data · 1 pt pitch binnen één minuut en begrijpelijk. \
  *Voorbeelden van goede KPI's:*
  - _A (JFK):_ gem. taxitijd #nl(an("JFK").taxi) min (PHX: #nl(an("PHX").taxi)); #an("JFK").pct% van de vertraagde vluchten ≥ 100 min; aandeel luchtverkeer #nl(kpi("JFK").oorzaak_pct.luchtverkeer)% (hoogste van zes). Advies: vertrekken in de piekuren beter spreiden en de rij voor de startbaan korter maken, bijvoorbeeld door vliegtuigen pas van de gate te laten vertrekken als er een startmoment is.
  - _B (maatschappij):_ % ≥ 100 min per vliegveld (JFK en ORD #an("ORD").pct–#an("JFK").pct%, PHX #an("PHX").pct%), aandeel "laat vliegtuig" (ORD #nl(kpi("ORD").oorzaak_pct.laat_vliegtuig)%), gem. vertraging. Advies: meer buffer tussen vluchten via ORD en JFK.
  - _C (reisapp):_ gem. vertraging per vliegveld, kans op ≥ 100 min, (bonus) tijdstip. Advies: toon bij het boeken "via Phoenix loop je gemiddeld minder vertraging op dan via New York JFK".
  Let op: wie met "% vluchten op tijd" komt, maakt de fout uit 1.1/3.2. Dat percentage kun je met deze data niet uitrekenen!
]

#vraag("3.4", punten: 3, bonus: true)[
  *Vertraging kost geld.* Stel dat elke minuut vertraging een maatschappij *\$#a.kosten_per_minuut* kost (brandstof, personeel, compensatie voor reizigers). Maak een draaitabel met *Maatschappij* in de #xl("rijen", "Rows") en de *#xl("Som", "Sum")* van *Vertrekvertraging_min* in de #xl("waarden", "Values"), en sorteer van groot naar klein. \
  a) Welke maatschappij verliest in onze dataset het meeste geld? Hoeveel ongeveer? \
  b) Een journalist schrijft: _"Dit is de slechtste maatschappij."_ Klopt dat? Voeg _#xl("Aantal", "Count")_ en _#xl("Gemiddelde", "Average")_ toe om het te controleren.
]
#antwoord(regels: 3)[
  #let k = a.kosten.first()
  a) *#k.mij*: #nl(k.som) min × \$#a.kosten_per_minuut = *± \$#nl(calc.round(k.kosten / 1000000, digits: 1)) miljoen*, en dat in een steekproef van maar #nl(a.aantal) vluchten. \
  b) *Nee.* #k.mij heeft gewoon veel vluchten in het bestand (#nl(k.n)), met gemiddeld #nl(k.gem) min. De hoogste gemiddelde vertraging (bij minstens 100 vluchten) heeft *#a.hoogste_gem_mij.mij* (#nl(a.hoogste_gem_mij.gem) min). *Totaal en gemiddelde vertellen een ander verhaal*: voor de financieel directeur telt het totaal, voor de reiziger het gemiddelde.
]

#terugblik[Je hebt duizenden rijen teruggebracht tot een paar *kernwaarden*, gekozen wát je laat zien aan *wie*, en daar een *advies* aan gekoppeld. In het studiepad Informatiesystemen leer je hoe organisaties met informatie en IT betere beslissingen nemen, en hoe je daarover praat met mensen die geen data-expert zijn.]

#stop-moment()

// ═════════════════════════════════════════════════════════════
//  Afsluiting
// ═════════════════════════════════════════════════════════════

= Eén dataset, drie studiepaden
#table(
  columns: (auto, 1fr, 1fr), inset: 8pt,
  table.header([Studiepad], [Wat je vandaag deed], [Wat je in het studiepad leert]),
  [*Kennis en data analytics*], [Filteren, sorteren, een verborgen regel ontdekken, formules schrijven], [Data analyseren, machine learning toepassen, kennissystemen bouwen, verantwoord omgaan met algoritmes en AI],
  [*Mens-technologie interactie*], [Een slecht ontwerp herkennen, kleurcodes, een grafiek ontwerpen en laten testen], [Gebruiksvriendelijkheid testen en verbeteren, hoe denken en emoties ontwerp beïnvloeden, technologie die gedrag verandert, leren met spelelementen],
  [*Informatiesystemen*], [KPI's kiezen per stakeholder, een mini-dashboard en advies maken], [Digitale ecosystemen in kaart brengen, softwareproducten ontwikkelen, informatiebeveiliging, een gesimuleerd bedrijf runnen],
)

#v(1em)
*Om over na te denken:* welk deel vond jij het leukst? Het puzzelen met de data, het ontwerpen van iets wat mensen snappen, of het advies aan de directeur? Dat zegt misschien iets over welk studiepad bij jou past. In het eerste jaar van Informatiekunde maak je kennis met alle drie. In het tweede jaar kies je een studiepad, en je kunt ook paden combineren.

#v(1fr)
#align(center, sans(8pt, fill: grijs)[Data: US Bureau of Transportation Statistics, 2008 · #nl(a.per_vliegveld) vertraagde vluchten per vliegveld · Meer weten? #link("https://www.uu.nl/bachelors/informatiekunde")[uu.nl/bachelors/informatiekunde]])
