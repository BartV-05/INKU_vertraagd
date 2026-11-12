"""Bouwt de presentatie voor de klassikale momenten van het practicum "Vertraagd!".

Gebruik:  ~/.venvs/practicum-vertraagd/bin/python scripts/bouw_presentatie.py <UU-sjabloon.pptx>
Uitvoer:  presentatie_vertraagd.pptx

Het sjabloon is een bestaande UU-presentatie (bijv. de open-dagpresentatie presOD.pptx).
Alle dia's daaruit worden verwijderd; de masters en layouts (huisstijl) blijven.
Getallen komen uit typst/antwoorden.json, grafieken uit typst/beeld/ — net als in het practicum.
"""
import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Cm, Pt

ROOT = Path(__file__).resolve().parent.parent
BEELD = ROOT / "typst" / "beeld"
A = json.loads((ROOT / "typst" / "antwoorden.json").read_text())

KDA, MTI, IS = "Kennis en data analytics", "Mens-technologie interactie", "Informatiesystemen"
KOPTEKST = "Informatiekunde · Vertraagd!"
BRON = "Bron: uu.nl/bachelors/informatiekunde (vakken en masters kunnen per collegejaar wijzigen)"

ZWART, WIT, GEEL = RGBColor(0, 0, 0), RGBColor(0xFF, 0xFF, 0xFF), RGBColor(0xFF, 0xCD, 0x00)
GRIJS = RGBColor(0x55, 0x55, 0x55)

# Layouts in het UU-sjabloon
L_TITEL, L_TEKST = "Title slide basisbeeld", "Title and object"

# Inhoudsgebied onder de witte kopband en boven de onderste band (in cm, dia is 33,87 × 19,05)
X0, BREED = 1.3, 31.2


def nl(x):
    """1500 -> 1.500, 41.1 -> 41,1"""
    heel, _, dec = str(x).partition(".")
    heel = f"{int(heel):,}".replace(",", ".")
    return f"{heel},{dec}" if dec else heel


# ── sjabloon voorbereiden ────────────────────────────────────

def zet_tekst(shape, tekst):
    """Vervangt de tekst van een vorm maar houdt de opmaak van de eerste run."""
    para = shape.text_frame.paragraphs[0]
    runs = para.runs
    runs[0].text = tekst
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    for p in shape.text_frame.paragraphs[1:]:
        p._p.getparent().remove(p._p)


def open_sjabloon(pad):
    prs = Presentation(pad)
    lijst = prs.slides._sldIdLst
    for sld in list(lijst):
        prs.part.drop_rel(sld.rId)
        lijst.remove(sld)
    vormen = list(prs.slide_masters[0].shapes) + list(layout(prs, L_TITEL).shapes)
    for sh in vormen:
        if sh.has_text_frame and sh.text_frame.text.strip() == "Informatica & Wiskunde":
            zet_tekst(sh, KOPTEKST)
    return prs


def layout(prs, naam):
    return next(l for l in prs.slide_layouts if l.name == naam)


def verwijder(shape):
    shape._element.getparent().remove(shape._element)


# ── bouwstenen ───────────────────────────────────────────────

def tekstvak(dia, x, y, w, h, regels, vul=None, rand=None, anker=MSO_ANCHOR.TOP, marge=0.3):
    """regels: lijst van (tekst, grootte_pt, vet, kleur) of strings (18 pt, normaal, zwart)."""
    vorm = dia.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    if vul is not None:
        vorm.fill.solid()
        vorm.fill.fore_color.rgb = vul
    if rand is not None:
        vorm.line.color.rgb = rand
        vorm.line.width = Pt(1.5)
    tf = vorm.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE  # vaste hoogte, zodat vakken naast elkaar even hoog zijn
    tf.vertical_anchor = anker
    for kant in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, kant, Cm(marge))
    for i, regel in enumerate(regels):
        if isinstance(regel, str):
            regel = (regel, 18, False, ZWART)
        tekst, grootte, vet, kleur = regel
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(grootte * 0.35)
        r = p.add_run()
        r.text = tekst
        r.font.size, r.font.bold, r.font.color.rgb = Pt(grootte), vet, kleur
    return vorm


def notities(dia, tekst):
    dia.notes_slide.notes_text_frame.text = tekst.strip()


def tekst_dia(prs, titel, punten=None, notitie=""):
    """Titel + opsomming op de layout 'Title and object'. punten: (tekst, niveau) of strings."""
    dia = prs.slides.add_slide(layout(prs, L_TEKST))
    dia.shapes.title.text = titel
    body = next(ph for ph in dia.placeholders if ph.placeholder_format.idx == 1)
    if punten:
        # De plaatshouder in het sjabloon is maar ~1 cm hoog; zonder dit verkleint Keynote/PowerPoint de tekst
        body.left, body.top, body.width, body.height = Cm(X0), Cm(5.9), Cm(BREED), Cm(10.3)
        tf = body.text_frame
        tf.auto_size = MSO_AUTO_SIZE.NONE
        for i, punt in enumerate(punten):
            tekst, niveau = (punt, 0) if isinstance(punt, str) else punt
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text, p.level = tekst, niveau
            for r in p.runs:
                r.font.size = Pt(24 if niveau == 0 else 20)
            p.space_after = Pt(10)
    else:
        verwijder(body)
    notities(dia, notitie)
    return dia


def afbeelding(dia, pad, x, y, max_w, max_h, kader=True):
    """Plaatst een afbeelding passend in een vak (zonder bijsnijden), gecentreerd, met wit kader."""
    pic = dia.shapes.add_picture(str(pad), Cm(x), Cm(y))
    schaal = min(Cm(max_w) / pic.width, Cm(max_h) / pic.height)
    pic.width, pic.height = int(pic.width * schaal), int(pic.height * schaal)
    pic.left = Cm(x) + (Cm(max_w) - pic.width) // 2
    pic.top = Cm(y) + (Cm(max_h) - pic.height) // 2
    if kader:
        pic.line.color.rgb = WIT
        pic.line.width = Pt(6)
    return pic


def tabel(dia, x, y, w, rijen, kolombreedtes, accent_rij=None, grootte=14):
    t = dia.shapes.add_table(len(rijen), len(rijen[0]), Cm(x), Cm(y), Cm(w), Cm(0.9 * len(rijen))).table
    for k, b in enumerate(kolombreedtes):
        t.columns[k].width = Cm(b)
    for r, rij in enumerate(rijen):
        for k, waarde in enumerate(rij):
            cel = t.cell(r, k)
            cel.text = str(waarde)
            cel.fill.solid()
            kop, accent = r == 0, r == accent_rij
            cel.fill.fore_color.rgb = ZWART if kop or accent else WIT
            cel.margin_top = cel.margin_bottom = Cm(0.1)
            p = cel.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if k == 0 else PP_ALIGN.RIGHT
            for run in p.runs:
                run.font.size = Pt(grootte)
                run.font.bold = kop or accent or k == 0
                run.font.color.rgb = GEEL if accent and not kop else (WIT if kop else ZWART)
    return t


def deel_start(prs, nr, pad, titel, vragen, bonus, extra, notitie):
    """Dia die tijdens het werken op het scherm blijft staan."""
    dia = tekst_dia(prs, f"Deel {nr}", notitie=notitie)
    tekstvak(dia, X0, 5.6, BREED, 4.2, [
        (f"DEEL {nr} · {pad.upper()}", 18, True, GEEL),
        (titel, 44, True, WIT),
    ], vul=ZWART, anker=MSO_ANCHOR.MIDDLE, marge=0.8)
    tekstvak(dia, X0, 10.4, 15, 5.8, [
        ("± 15 minuten", 32, True, ZWART),
        (f"Vragen {vragen}", 22, False, ZWART),
        (f"Klaar? Bonus {bonus}", 22, False, ZWART),
    ])
    tekstvak(dia, 17.5, 10.4, 15, 5.8, [(e, 20, False, ZWART) for e in extra],
             vul=WIT, marge=0.6)
    return dia


def studiepad_dia(prs, pad, deel, kernvraag, vandaag, jaar2, daarna, notitie):
    dia = tekst_dia(prs, f"Studiepad {pad}", notitie=notitie)
    tekstvak(dia, X0, 5.4, BREED, 2.2, [(kernvraag, 20, True, ZWART)])
    kol_w, gat = (BREED - 1.0) / 3, 0.5
    kolommen = [
        (f"Vandaag (deel {deel})", vandaag[0], [vandaag[1]]),
        ("In jaar 2: het studiepad", "4 vakken, onder andere:", jaar2),
        ("Na je bachelor", "Bijvoorbeeld:", daarna),
    ]
    for i, (kop, sub, items) in enumerate(kolommen):
        x = X0 + i * (kol_w + gat)
        vul = ZWART if i == 0 else WIT
        kleur = WIT if i == 0 else ZWART
        tekstvak(dia, x, 7.8, kol_w, 8.3,
                 [(kop, 20, True, GEEL if i == 0 else ZWART), (sub, 16, False, kleur)]
                 + [(("• " if i else "") + it, 16, i == 0, kleur) for it in items],
                 vul=vul, marge=0.5)
    tekstvak(dia, X0, 16.2, BREED, 0.8, [(BRON, 9, False, GRIJS)], marge=0)
    return dia


# ── de dia's ─────────────────────────────────────────────────

def bouw(prs):
    an = {r["vv"]: r for r in A["analyse"]}
    kpi = {k["vv"]: k for k in A["kpi"]}
    hoogste_lv = max(A["kpi"], key=lambda k: k["oorzaak_min"]["luchtverkeer"])

    # 1 Titel
    dia = prs.slides.add_slide(layout(prs, L_TITEL))
    dia.shapes.title.text = "Vertraagd!"
    p = dia.shapes.title.text_frame.add_paragraph()
    p.text = "Practicum Informatiekunde"
    for r in p.runs:
        r.font.size, r.font.bold = Pt(16), False
    foto = next(ph for ph in dia.placeholders if ph.placeholder_format.idx == 14)
    foto.insert_picture(str(BEELD / "KBG.jpg"))
    notities(dia, """
Welkom! Stel jezelf voor: naam, en of je studeert of werkt bij Informatiekunde.
Vraag: wie heeft weleens op Schiphol gezeten terwijl er 'vertraagd' achter je vlucht stond? (handen)
Voor jou is het een uurtje wachten, maar voor een vliegveld of luchtvaartmaatschappij kost elke minuut vertraging heel veel geld.
""")

    # 2 De casus
    tekst_dia(prs, "De casus: jullie zijn data-adviseurs", [
        "Vandaag werken jullie bij DataVlucht, een (verzonnen) adviesbureau",
        f"Jullie klant stuurde {nl(A['aantal'])} echte vertraagde vluchten uit de VS (2008)",
        (f"{nl(A['per_vliegveld'])} vanaf elk van zes grote vliegvelden: ATL · ORD · DEN · LAX · PHX · JFK", 1),
        "De vraag van de klant:",
        ("“Waar zit de vertraging, waar komt het door, en wat moeten we eraan doen?”", 1),
    ], notitie="""
Vandaag zijn jullie geen scholieren maar data-adviseurs bij DataVlucht.
De klant heeft een bestand met echte vertraagde vluchten uit Amerika gestuurd en wil drie dingen weten: waar zit de vertraging, waar komt het door, en wat moeten we eraan doen?
""")

    # 3 Drie studiepaden
    dia = tekst_dia(prs, "Informatiekunde: mens, organisatie en technologie", notitie=f"""
Om die vraag te beantwoorden heb je drie soorten informatiekundigen nodig. Dat zijn precies de drie studiepaden van onze opleiding.
Eén: iemand die snapt wat er écht in de data staat: {KDA}.
Twee: iemand die het zo laat zien dat mensen het in één oogopslag snappen: {MTI}.
Drie: iemand die bepaalt welke cijfers belangrijk zijn en welk advies daaruit volgt: {IS}.
Jullie proberen vandaag alle drie, steeds een kwartier. Na elk deel stoppen we even en vertel ik wat je in dat studiepad écht doet.
""")
    kol_w = (BREED - 1.0) / 3
    for i, (pad, vraag) in enumerate([
        (KDA, "Hoe maak je kennis uit data, en hoe pas je die verantwoord toe?"),
        (MTI, "Hoe werken mensen en systemen zo goed mogelijk samen?"),
        (IS, "Hoe ondersteunt technologie organisaties en de maatschappij slim en zinvol?"),
    ]):
        tekstvak(dia, X0 + i * (kol_w + 0.5), 6.0, kol_w, 9.5, [
            (f"DEEL {i + 1}", 14, True, GEEL),
            (pad, 21, True, WIT),
            (vraag, 18, False, WIT),
        ], vul=ZWART, marge=0.6)

    # 4 Werkwijze
    tekst_dia(prs, "Zo werken we vandaag", [
        "In tweetallen: één bestuurt de muis, de ander leest voor. Wissel per deel!",
        "Drie delen van ± 15 minuten",
        "Na elk deel: STOP. We kijken samen even terug (± 3 minuten)",
        "Snel klaar? Er staan bonusvragen in",
        "Geen Excel-ervaring nodig: alles staat stap voor stap op het blad",
        "Vastgelopen? Steek je hand op!",
        "Open nu vluchten_practicum.xlsx en begin bij “De casus”",
    ], notitie="""
Laat de timer zien. Na een kwartier zeg ik 'stop', ook als je nog niet klaar bent: dat is helemaal niet erg.
Wie snel is, kan door met de bonusvragen.
Je hoeft geen Excel te kunnen. Kom je er niet uit? Steek je hand op, daar zijn wij voor.
""")

    # 5 Deel 1 start
    deel_start(prs, 1, KDA, "Data-detective", "1.1 t/m 1.5", "1.6", [
        "Tip: de statusbalk onderin laat Gemiddelde, Aantal en Som zien (Average, Count, Sum)",
        "Wis je filters voordat je verdergaat!",
    ], notitie="""
Data-experts zijn een beetje detectives. Voordat je iets concludeert, moet je weten wat er in je data zit, en vooral wat er níet in zit.
En aan het eind schrijven jullie je eerste formules.
Laat op de beamer zien waar de statusbalk zit. Laat deze dia staan terwijl ze werken.
""")

    # 6 Terugblik deel 1
    leeg = A["leeg"]
    dia = tekst_dia(prs, "Terugblik deel 1: de verborgen regel", notitie=f"""
Wie heeft de detectivevraag (1.4) opgelost? Laat iemand de regel vertellen.
Antwoord: de oorzaak van vertraging wordt pas geregistreerd als een vlucht 15 minuten of meer te laat aankomt. Bij lege cellen is de aankomstvertraging hooguit {leeg['max_aankomst']} minuten.
Dat is de officiële Amerikaanse definitie van een vertraagde vlucht. Data heeft regels die je pas ziet als je gaat zoeken.
Wijs dan naar de tabel: dit hebben jullie zelf met formules uitgerekend. JFK springt eruit: de langste taxitijd én de meeste ernstige vertragingen.
""")
    tekstvak(dia, X0, 5.6, 14.5, 10.5, [
        ("Lege oorzaak-cellen?", 22, True, ZWART),
        (f"Aankomst hooguit {leeg['max_aankomst']} min te laat", 18, False, ZWART),
        ("De regel:", 22, True, ZWART),
        ("Een oorzaak wordt pas geregistreerd vanaf 15 minuten te laat aankomen", 18, False, ZWART),
        ("Les: data heeft regels die je niet ziet, tot je gaat zoeken.", 18, True, ZWART),
    ], vul=WIT, marge=0.6)
    rijen = [["Vliegveld", "Gem. vertraging", "Gem. taxitijd", "≥ 100 min"]]
    for vv in ("ATL", "ORD", "DEN", "LAX", "PHX", "JFK"):
        r = an[vv]
        rijen.append([vv, f"{nl(r['vertr'])} min", f"{nl(r['taxi'])} min", f"{r['pct']}%"])
    tekstvak(dia, 16.6, 5.4, 15.9, 1.0, [("Jullie eigen formules (tabblad Analyse):", 16, True, ZWART)], marge=0)
    tabel(dia, 16.6, 6.5, 15.9, rijen, [3.3, 4.4, 4.2, 4.0],
          accent_rij=rijen.index(next(r for r in rijen if r[0] == "JFK")), grootte=15)

    # 7 Studiepad KDA
    studiepad_dia(prs, KDA, 1,
        "Hoe maak je kennis uit data, en hoe pas je die verantwoord toe?",
        ("9.000 rijen in Excel:", "filteren, een verborgen regel vinden, je eerste formules"),
        ["Data Analytics", "Applications of Machine Learning",
         "Knowledge-intensive process analysis", "Responsible Data Science"],
        ["Master Applied Data Science", "Master Artificial Intelligence", "Data-analist"],
        notitie=f"""
Wat jullie net deden is maar een klein proefje. In het studiepad {KDA} werk je niet met 9.000 rijen in Excel, maar met miljoenen rijen. Daarvoor leer je in de opleiding SQL en Python gebruiken.
In jaar 2 kies je dit studiepad: vier vakken, zoals Data Analytics en Applications of Machine Learning. Je leert ook nadenken over de ethiek van algoritmes en AI (Responsible Data Science).
Na je bachelor kun je bijvoorbeeld door met de master Applied Data Science of Artificial Intelligence, of als data-analist aan de slag.
Controleer vooraf of de vaknamen nog kloppen voor dit collegejaar (uu.nl/bachelors/informatiekunde/studieprogramma).
""")

    # 8 Deel 2 start
    deel_start(prs, 2, MTI, "Kleur en grafieken die spreken", "2.1 t/m 2.5", "2.6", [
        "Denk aan het dashboard van je smartwatch of je bank-app",
        "Wat zie je in drie seconden?",
    ], notitie=f"""
Jullie hebben nu getallen. Maar een directeur, of iemand die naar zijn smartwatch kijkt, gaat geen tabellen lezen.
Hoe zorg je dat iemand in drie seconden ziet wat er aan de hand is? Daar gaat {MTI} over: ontwerpen voor hoe mensen kijken en met technologie omgaan.
""")

    # 9 Terugblik deel 2
    dia = tekst_dia(prs, "Terugblik deel 2: zoeken of in drie seconden snappen?", notitie="""
Links de grafiek van de stagiair uit vraag 2.1. Wat was er allemaal mis? Laat een paar dingen noemen: geen titel, een taart voor iets wat geen geheel is, 3D, regenboogkleuren...
Rechts dezelfde soort informatie, goed ontworpen: gesorteerd, alles grijs behalve één accentkleur, en een titel die de conclusie vertelt.
De les uit 2.3: kleur is goed om iets op te laten vallen, lengte om te vergelijken.
Als er tijd is: laat één tweetal hun eigen grafiek uit 2.4 op het scherm zien.
""")
    afbeelding(dia, BEELD / "slechte_grafiek.png", X0, 5.7, 14.8, 8.4)
    afbeelding(dia, BEELD / "goed_taxi.png", 17.7, 5.7, 14.8, 8.4)
    tekstvak(dia, X0, 14.4, 14.8, 1.8, [("Zoeken…", 22, True, ZWART)], marge=0)
    tekstvak(dia, 17.7, 14.4, 14.8, 1.8, [("…of in drie seconden snappen", 22, True, ZWART)], marge=0)

    # 10 Studiepad MTI
    studiepad_dia(prs, MTI, 2,
        "Hoe werken mensen en systemen zo goed mogelijk samen?",
        ("Eén grafiek en een kleurcode:", "ontwerpen met duidelijke regels, en laten testen door een ander"),
        ["Usability Engineering and User Experience", "Cognitie en Emotie",
         "Persuasive technologies", "Playful learning"],
        ["Master Human Computer Interaction", "UX-onderzoeker", "Multimedia designer"],
        notitie=f"""
Vandaag hebben jullie één grafiek en een kleurcode ontworpen. In het studiepad {MTI} ontwerp je hele apps, websites en dashboards, en test je die met echte gebruikers.
Je leert hoe mensen denken en wat ze voelen als ze technologie gebruiken (Cognitie en Emotie), hoe technologie gedrag kan veranderen (Persuasive technologies, denk aan je stappenteller) en hoe spelelementen helpen bij leren (Playful learning).
Na je bachelor kun je bijvoorbeeld door met de master Human Computer Interaction, of aan de slag als UX-onderzoeker.
Controleer vooraf of de vaknamen nog kloppen voor dit collegejaar.
""")

    # 11 Deel 3 start
    deel_start(prs, 3, IS, "Van data naar besluit", "3.1 t/m 3.3", "3.4", [
        "Eindproduct: jouw mini-dashboard voor één stakeholder",
        "3 kernwaarden · 1 grafiek · 1 advies",
    ], notitie="""
Nu wordt het zakelijk. Een vliegvelddirecteur, een luchtvaartmaatschappij en een app-bouwer willen allemaal iets anders weten.
Jullie kiezen de juiste getallen en geven een advies.
Doe stap 1 t/m 4 van de draaitabel in vraag 3.1 voor op de beamer: dat scheelt veel vragen.
""")

    # 12 Terugblik deel 3
    o_ord, m_ord = kpi["ORD"]["oorzaak_pct"], kpi["ORD"]["oorzaak_min"]
    dia = tekst_dia(prs, "Terugblik deel 3: wie kan er iets aan doen?", notitie=f"""
Vraag 3.1: bij ORD (Chicago) komt de meeste vertraging ({nl(m_ord['laat_vliegtuig'])} minuten, bijna de helft) door een vliegtuig dat al te laat binnenkwam. Dat is een sneeuwbaleffect: dat moet de maatschappij oplossen.
{hoogste_lv['vv']} heeft de meeste minuten vertraging door luchtverkeer van alle zes ({nl(hoogste_lv['oorzaak_min']['luchtverkeer'])} min), en ook de langste taxitijd. Daar kan het vliegveld zelf dus echt iets doen.
Omdat elk vliegveld evenveel vluchten in het bestand heeft, mag je die minuten eerlijk vergelijken.
Vraag 3.2: elk vliegveld heeft precies {nl(A['per_vliegveld'])} vluchten. Dat is een keuze van de maker van het bestand, geen eigenschap van de vliegvelden. Data wordt gemaakt door mensen die keuzes maken.
""")
    rijen = [["Grootste oorzaak", "Wie lost het op?"],
             [f"ORD: laat binnengekomen vliegtuig ({nl(m_ord['laat_vliegtuig'])} min)", "De maatschappij"],
             [f"{hoogste_lv['vv']}: luchtverkeer ({nl(hoogste_lv['oorzaak_min']['luchtverkeer'])} min, meeste van zes)", "Het vliegveld"],
             ["Weer", "Niemand"]]
    tabel(dia, X0, 5.8, 19.0, rijen, [12.5, 6.5], grootte=18)
    tekstvak(dia, 21.3, 5.8, 11.2, 10.3, [
        ("Elk vliegveld precies 1.500 vluchten?", 20, True, ZWART),
        ("Dat is een keuze van de maker van het bestand.", 17, False, ZWART),
        ("Les: data wordt gemaakt door mensen die keuzes maken. Vraag altijd: wat zit erin, en waarom?", 17, True, ZWART),
    ], vul=WIT, marge=0.6)

    # 13 Studiepad IS
    studiepad_dia(prs, IS, 3,
        "Hoe kan technologie organisaties en de maatschappij efficiënt, effectief en zinvol ondersteunen?",
        ("Kernwaarden en een advies:", "het juiste getal voor de juiste persoon kiezen"),
        ["Architectuur van digitale ecosystemen", "Productsoftware", "Information Security",
         "Strategisch management van organisaties en ICT"],
        ["Master Business Informatics", "IT-adviseur", "Softwareondernemer"],
        notitie=f"""
Vandaag kozen jullie de juiste kernwaarden voor één stakeholder. In het studiepad {IS} kijk je naar hele organisaties: hoe werken bedrijven samen in een digitaal ecosysteem, hoe bouw je een softwareproduct dat veel klanten gebruiken, en hoe bescherm je informatie?
In het vak Strategisch management van organisaties en ICT run je zelfs een gesimuleerd bedrijf.
Na je bachelor kun je bijvoorbeeld door met de master Business Informatics, of aan de slag als IT-adviseur of als softwareondernemer.
Controleer vooraf of de vaknamen nog kloppen voor dit collegejaar.
""")

    # 14 Pitch
    tekst_dia(prs, "Jullie mini-dashboards", [
        "Twee tweetallen, elk met een andere stakeholder",
        "Eén minuut per tweetal",
        "Luister mee: zie je 3 kernwaarden, 1 grafiek en 1 advies?",
        ("Valkuil: een percentage ‘vluchten op tijd’ kun je met deze data niet uitrekenen!", 1),
    ], notitie="""
Kies twee tweetallen met een verschillende stakeholder (A: directeur JFK, B: maatschappij, C: reisapp).
Geef na elke pitch één compliment en één tip. Let op de valkuil: wie met '% op tijd' komt, vergeet dat de data alleen vertraagde vluchten bevat.
""")

    # 15 Opbouw van de studie
    dia = tekst_dia(prs, "Zo ziet de bachelor Informatiekunde eruit", notitie="""
Vandaag hebben jullie aan alle drie de studiepaden geproefd. In de opleiding gaat dat zo:
In jaar 1 volgt iedereen dezelfde acht vakken. Daarin maak je kennis met alle drie de kanten van Informatiekunde.
In jaar 2 kies je een studiepad van vier vakken. Je kunt ook studiepaden combineren.
In jaar 3 kies je keuzevakken, ook bij andere faculteiten of universiteiten, en schrijf je je bachelorscriptie.
""")
    jaren = [
        ("Jaar 1", "Iedereen hetzelfde", ["8 verplichte vakken", "Kennismaken met alle drie", "o.a. Datamodelleren, Computationeel denken, Procesmodelleren"]),
        ("Jaar 2", "Kies je studiepad", ["4 vakken van jouw studiepad", "Paden combineren kan", "Plus verplichte vakken en keuzevakken"]),
        ("Jaar 3", "Jouw eigen richting", ["Keuzevakken, ook buiten de opleiding", "Bachelorscriptie"]),
    ]
    kol_w = (BREED - 1.0) / 3
    for i, (jaar, kop, items) in enumerate(jaren):
        tekstvak(dia, X0 + i * (kol_w + 0.5), 5.8, kol_w, 10.3,
                 [(jaar, 30, True, GEEL if i == 1 else ZWART), (kop, 20, True, WIT if i == 1 else ZWART)]
                 + [("• " + it, 16, False, WIT if i == 1 else ZWART) for it in items],
                 vul=ZWART if i == 1 else WIT, marge=0.6)
    tekstvak(dia, X0, 16.2, BREED, 0.8, [(BRON, 9, False, GRIJS)], marge=0)

    # 16 Welk pad past bij jou?
    dia = tekst_dia(prs, "Welk deel vond jij het leukst?", notitie=f"""
Vraag: steek je hand op bij het deel dat je het leukst vond. Deel 1? Deel 2? Deel 3?
Wie deel 1 het leukst vond, zou zich thuis voelen bij {KDA}. Deel 2: {MTI}. Deel 3: {IS}.
Twijfel je? Dat is helemaal goed: in het eerste jaar maak je kennis met alle drie, en pas in jaar 2 kies je. Je kunt ze ook combineren.
""")
    for i, (deel, wat, pad) in enumerate([
        ("Deel 1", "Puzzelen met data en verborgen regels vinden", KDA),
        ("Deel 2", "Iets ontwerpen wat mensen meteen snappen", MTI),
        ("Deel 3", "Het advies aan de directeur", IS),
    ]):
        tekstvak(dia, X0 + i * (kol_w + 0.5), 6.0, kol_w, 9.5, [
            (deel, 28, True, ZWART), (wat, 20, False, ZWART), ("→ " + pad, 20, True, ZWART),
        ], vul=WIT, marge=0.6)

    # 17 Meer weten
    tekst_dia(prs, "Bedankt! Meer weten?", [
        "uu.nl/bachelors/informatiekunde",
        "Kom een dag meelopen: ervaar een echte studiedag en spreek studenten",
        "Nog vragen? Stel ze nu aan ons!",
    ], notitie="""
Bedank iedereen. Wijs op de meeloopdagen en de website. Vul hier eventueel een contactpersoon of QR-code voor de meeloopdag aan.
""")


def main():
    prs = open_sjabloon(Path(sys.argv[1]).expanduser())
    bouw(prs)
    uit = ROOT / "presentatie_vertraagd.pptx"
    prs.save(uit)
    print(f"{len(prs.slides)} dia's geschreven naar {uit}")


if __name__ == "__main__":
    main()
