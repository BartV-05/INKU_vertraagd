"""Bouwt de gereduceerde dataset (Excel) en de antwoorden (JSON) voor het practicum "Vertraagd!".

Gebruik:  python3 -I scripts/bouw_dataset.py <pad/naar/DelayedFlights.csv>
Uitvoer:  data/vluchten_practicum.xlsx (+ .csv), typst/antwoorden.json, typst/beeld/*.png

Alle getallen in het nakijkmodel komen uit dit script, zodat dataset en antwoorden
altijd bij elkaar passen. Verander je SEED, PER_VLIEGVELD of VLIEGVELDEN, draai dan opnieuw.
"""
import json
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

SEED = 2008
PER_VLIEGVELD = 1_500  # gelijke aantallen: zo zijn vliegvelden eerlijk te vergelijken
DREMPEL = 100          # "ernstige vertraging" in de formule- en kleuropdrachten
KOSTEN_PER_MINUUT = 75  # aanname in dollar, orde van grootte uit sectorrapporten
VLIEGVELDEN = {
    "ATL": "Atlanta (Hartsfield-Jackson)",
    "ORD": "Chicago (O'Hare)",
    "DEN": "Denver",
    "LAX": "Los Angeles",
    "PHX": "Phoenix (Sky Harbor)",
    "JFK": "New York (John F. Kennedy)",
}
MAATSCHAPPIJEN = {
    "WN": "Southwest Airlines", "AA": "American Airlines", "MQ": "American Eagle",
    "UA": "United Airlines", "OO": "SkyWest", "DL": "Delta Air Lines",
    "XE": "ExpressJet", "CO": "Continental Airlines", "US": "US Airways",
    "EV": "Atlantic Southeast", "NW": "Northwest Airlines", "FL": "AirTran Airways",
    "YV": "Mesa Airlines", "B6": "JetBlue Airways", "OH": "Comair",
    "9E": "Pinnacle Airlines", "AS": "Alaska Airlines", "F9": "Frontier Airlines",
    "HA": "Hawaiian Airlines", "AQ": "Aloha Airlines",
}
WEEKDAGEN = {1: "1-ma", 2: "2-di", 3: "3-wo", 4: "4-do", 5: "5-vr", 6: "6-za", 7: "7-zo"}
OORZAKEN = {
    "CarrierDelay": "Oorzaak_maatschappij_min",
    "WeatherDelay": "Oorzaak_weer_min",
    "NASDelay": "Oorzaak_luchtverkeer_min",
    "SecurityDelay": "Oorzaak_beveiliging_min",
    "LateAircraftDelay": "Oorzaak_laat_vliegtuig_min",
}

ROOT = Path(__file__).resolve().parent.parent

KOLOMUITLEG = [
    ("Vlucht_ID", "Uniek nummer van de rij."),
    ("Datum", "Datum van de vlucht (alle vluchten zijn uit 2008)."),
    ("Maand", "Maandnummer: 1 = januari ... 12 = december."),
    ("Weekdag", "Dag van de week: 1-ma t/m 7-zo."),
    ("Maatschappij_code", "Tweeletterige code van de luchtvaartmaatschappij."),
    ("Maatschappij", "Naam van de luchtvaartmaatschappij."),
    ("Vluchtnummer", "Code + nummer, bv. WN335. Hetzelfde vluchtnummer vliegt vaak elke dag."),
    ("Van", "Vertrekvliegveld (zie tabblad Codes)."),
    ("Naar", "Aankomstvliegveld (driecijferige IATA-code)."),
    ("Geplande_vertrektijd", "Vertrektijd volgens de dienstregeling (lokale tijd)."),
    ("Vertrekuur", "Het uur van de geplande vertrektijd (0-23). Handig voor draaitabellen."),
    ("Afstand_km", "Afstand tussen de twee vliegvelden in kilometers."),
    ("Vertrekvertraging_min", "Hoeveel minuten later dan gepland het vliegtuig vertrok."),
    ("Aankomstvertraging_min", "Hoeveel minuten later dan gepland het vliegtuig aankwam. Negatief = te vroeg!"),
    ("Taxi_uit_min", "Minuten tussen loskoppelen van de gate en opstijgen."),
    ("Taxi_in_min", "Minuten tussen landen en aankomst bij de gate."),
    ("Oorzaak_maatschappij_min", "Vertraging door de maatschappij zelf (onderhoud, bemanning, schoonmaak, bagage...)."),
    ("Oorzaak_weer_min", "Vertraging door extreem weer."),
    ("Oorzaak_luchtverkeer_min", "Vertraging door het luchtverkeerssysteem: drukte, verkeersleiding, gesloten banen."),
    ("Oorzaak_beveiliging_min", "Vertraging door beveiliging (ontruiming, lange rijen bij de controle...)."),
    ("Oorzaak_laat_vliegtuig_min", "Vertraging omdat het vliegtuig te laat binnenkwam van een eerdere vlucht."),
]


def laad(csv_pad):
    kol = ["Month", "DayofMonth", "DayOfWeek", "CRSDepTime", "UniqueCarrier", "FlightNum",
           "Origin", "Dest", "Distance", "ArrDelay", "DepDelay", "TaxiIn", "TaxiOut",
           "Cancelled", "Diverted", *OORZAKEN]
    df = pd.read_csv(csv_pad, usecols=kol)
    df = df[(df.Cancelled == 0) & (df.Diverted == 0) & df.Origin.isin(VLIEGVELDEN)]
    # Gestratificeerd: per vliegveld evenveel vluchten. Een proportionele steekproef zou
    # suggereren dat JFK "klein" is, terwijl de bron alleen binnenlandse vluchten bevat.
    return df.groupby("Origin", group_keys=False).sample(n=PER_VLIEGVELD, random_state=SEED)


def vertaal(df):
    datum = pd.to_datetime(dict(year=2008, month=df.Month, day=df.DayofMonth))
    uit = pd.DataFrame({
        "Datum": datum,
        "Maand": df.Month,
        "Weekdag": df.DayOfWeek.map(WEEKDAGEN),
        "Maatschappij_code": df.UniqueCarrier,
        "Maatschappij": df.UniqueCarrier.map(MAATSCHAPPIJEN),
        "Vluchtnummer": df.UniqueCarrier + df.FlightNum.astype(str),
        "Van": df.Origin,
        "Naar": df.Dest,
        "Geplande_vertrektijd": df.CRSDepTime,
        "Vertrekuur": (df.CRSDepTime // 100) % 24,
        "Afstand_km": (df.Distance * 1.609344).round().astype(int),
        "Vertrekvertraging_min": df.DepDelay.astype(int),
        "Aankomstvertraging_min": df.ArrDelay.astype(int),
        "Taxi_uit_min": df.TaxiOut.astype(int),
        "Taxi_in_min": df.TaxiIn.astype(int),
    })
    for bron, doel in OORZAKEN.items():
        uit[doel] = df[bron].astype("Int64")
    uit = uit.sort_values(["Datum", "Geplande_vertrektijd", "Vluchtnummer"]).reset_index(drop=True)
    uit.insert(0, "Vlucht_ID", range(1, len(uit) + 1))
    return uit


def schrijf_excel(df, pad):
    wb = Workbook()
    lees = wb.active
    lees.title = "Lees mij"
    lees["A1"] = "Vertraagd! - vertraagde vluchten in de VS, 2008"
    lees["A1"].font = Font(size=16, bold=True)
    lees["A3"] = (f"Dit bestand bevat {len(df):,} binnenlandse vluchten binnen de VS uit 2008: precies "
                  f"{PER_VLIEGVELD:,} per vertrekvliegveld, willekeurig gekozen. LET OP: in de brondata "
                  "zitten alleen vluchten die minstens 6 minuten te laat vertrokken. Op tijd vertrokken "
                  "vluchten ontbreken dus!").replace(",", ".")
    lees["A3"].alignment = Alignment(wrap_text=True, vertical="top")
    lees.merge_cells("A3:B3")
    lees.row_dimensions[3].height = 62
    lees["A5"], lees["B5"] = "Kolom", "Betekenis"
    for c in ("A5", "B5"):
        lees[c].font = Font(bold=True, color="FFFFFF")
        lees[c].fill = PatternFill("solid", fgColor="1F3864")
    for i, (k, uitleg) in enumerate(KOLOMUITLEG, start=6):
        lees.cell(i, 1, k).font = Font(bold=True)
        lees.cell(i, 2, uitleg).alignment = Alignment(wrap_text=True)
    n = 6 + len(KOLOMUITLEG) + 1
    lees.cell(n, 1, "Lege cellen").font = Font(bold=True)
    lees.cell(n, 2, "Bij sommige vluchten zijn de Oorzaak-kolommen leeg. Waarom? Dat zoek je zelf uit in Deel 1!")
    lees.cell(n + 1, 1, "Bron").font = Font(bold=True)
    lees.cell(n + 1, 2, "US Bureau of Transportation Statistics, via de Kaggle-dataset 'Airline Delay Causes' "
                        f"(gereduceerd: {PER_VLIEGVELD} willekeurige vluchten per vliegveld, seed {SEED}).")
    lees.column_dimensions["A"].width = 28
    lees.column_dimensions["B"].width = 95

    ws = wb.create_sheet("Vluchten")
    ws.append(list(df.columns))
    for rij in df.itertuples(index=False):
        rij = list(rij)
        rij[1] = rij[1].to_pydatetime().date()
        rij[9] = f"{rij[9] // 100:02d}:{rij[9] % 100:02d}"
        ws.append([None if pd.isna(v) else v for v in rij])
    for r in range(2, len(df) + 2):
        ws.cell(r, 2).number_format = "dd-mm-yyyy"
    breedtes = {"Datum": 12, "Maatschappij": 20, "Geplande_vertrektijd": 12}
    for i, k in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = breedtes.get(k, max(10, min(len(k) + 2, 16)))
    ws.row_dimensions[1].height = 45
    for c in ws[1]:
        c.alignment = Alignment(wrap_text=True, vertical="center")
    tabel = Table(displayName="Vluchten", ref=f"A1:{get_column_letter(df.shape[1])}{len(df) + 1}")
    tabel.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tabel)
    ws.freeze_panes = "C2"

    codes = wb.create_sheet("Codes")
    codes.append(["Vliegveld", "Stad / naam"])
    for k, v in VLIEGVELDEN.items():
        codes.append([k, v])
    codes.append([])
    codes.append(["Maatschappij_code", "Maatschappij"])
    for k, v in sorted(MAATSCHAPPIJEN.items(), key=lambda kv: kv[1]):
        if k in set(df.Maatschappij_code):
            codes.append([k, v])
    for r in (1, len(VLIEGVELDEN) + 3):
        for c in codes[r]:
            c.font = Font(bold=True)
    codes.column_dimensions["A"].width = 20
    codes.column_dimensions["B"].width = 32

    analyse = wb.create_sheet("Analyse", 2)
    koppen = ["Vliegveld", "Aantal vluchten", "Gem. vertrekvertraging (min)", "Gem. taxitijd (min)",
              f"Aantal ≥ {DREMPEL} min", f"% ≥ {DREMPEL} min"]
    analyse.append(koppen)
    for vv in VLIEGVELDEN:
        analyse.append([vv])
    for c in analyse[1]:
        c.font = Font(bold=True, color="000000")
        c.fill = PatternFill("solid", fgColor="FFCD00")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    analyse.row_dimensions[1].height = 32
    invul = PatternFill("solid", fgColor="FFF5CC")
    for r in range(2, len(VLIEGVELDEN) + 2):
        analyse.cell(r, 1).font = Font(bold=True)
        for k in range(2, len(koppen) + 1):
            analyse.cell(r, k).fill = invul
            analyse.cell(r, k).number_format = "0%" if k == len(koppen) else "0.0"
        analyse.cell(r, 2).number_format = "0"
        analyse.cell(r, 5).number_format = "0"
    analyse["A9"] = "Hier schrijf je in Deel 1 je eigen formules (gele cellen). Deel 2 en 3 bouwen hierop verder."
    analyse["A9"].font = Font(italic=True, color="666666")
    analyse["H1"], analyse["I1"] = "Drempel (min):", DREMPEL
    analyse["H1"].font = Font(bold=True)
    analyse.column_dimensions["A"].width = 12
    for k in "BCDEF":
        analyse.column_dimensions[k].width = 16
    analyse.column_dimensions["H"].width = 15

    wb.active = 1
    wb.save(pad)


def r1(x):
    return round(float(x), 1)


def antwoorden(df):
    a = {"aantal": len(df), "per_vliegveld": PER_VLIEGVELD, "drempel": DREMPEL, "seed": SEED,
         "kosten_per_minuut": KOSTEN_PER_MINUUT}
    a["per_vliegveld_aantal"] = df.Van.value_counts().to_dict()

    # Deel 1
    top = df.loc[df.Vertrekvertraging_min.idxmax()]
    a["max_vertraging"] = {
        "vlucht": top.Vluchtnummer, "maatschappij": top.Maatschappij,
        "datum": top.Datum.strftime("%d-%m-%Y"), "van": top.Van, "naar": top.Naar,
        "minuten": int(top.Vertrekvertraging_min),
        "uren": r1(top.Vertrekvertraging_min / 60),
    }
    jfk_dec = df[(df.Van == "JFK") & (df.Maand == 12)]
    a["jfk_dec"] = {"aantal": len(jfk_dec), "gem": r1(jfk_dec.Vertrekvertraging_min.mean())}
    lang = df[df.Vertrekvertraging_min >= 180]
    a["drie_uur"] = {"aantal": len(lang), "per_vliegveld": lang.Van.value_counts().to_dict()}
    leeg = df[df.Oorzaak_weer_min.isna()]
    gevuld = df[df.Oorzaak_weer_min.notna()]
    a["leeg"] = {
        "aantal": len(leeg),
        "max_aankomst": int(leeg.Aankomstvertraging_min.max()),
        "min_aankomst": int(leeg.Aankomstvertraging_min.min()),
        "min_aankomst_gevuld": int(gevuld.Aankomstvertraging_min.min()),
        "te_vroeg": int((df.Aankomstvertraging_min < 0).sum()),
    }
    a["analyse"] = []
    for vv in VLIEGVELDEN:
        d = df[df.Van == vv]
        ernstig = int((d.Vertrekvertraging_min >= DREMPEL).sum())
        a["analyse"].append({
            "vv": vv, "n": len(d), "vertr": r1(d.Vertrekvertraging_min.mean()), "taxi": r1(d.Taxi_uit_min.mean()),
            "ernstig": ernstig, "pct": round(ernstig / len(d) * 100),
            "ernstig60": int((d.Vertrekvertraging_min >= 60).sum()),
            "pct60": round((d.Vertrekvertraging_min >= 60).mean() * 100),
        })
    a["totaal_ernstig"] = int((df.Vertrekvertraging_min >= DREMPEL).sum())

    # Deel 2
    per_uur = df.groupby("Vertrekuur").Vertrekvertraging_min.agg(["size", "mean"])
    a["per_uur"] = [{"uur": int(u), "n": int(r["size"]), "gem": r1(r["mean"])} for u, r in per_uur.iterrows()]
    taxi = df.groupby("Van").Taxi_uit_min.mean().sort_values(ascending=False)
    a["taxi_per_vliegveld"] = [{"vv": k, "gem": r1(v)} for k, v in taxi.items()]

    # Deel 3
    kpi = df.groupby("Van").agg(
        n=("Vlucht_ID", "size"),
        vertr=("Vertrekvertraging_min", "mean"),
        taxi=("Taxi_uit_min", "mean"),
        uur_plus=("Vertrekvertraging_min", lambda s: (s >= 60).mean() * 100),
    )
    oorz = df.groupby("Van")[list(OORZAKEN.values())].sum()
    oorz_pct = oorz.div(oorz.sum(axis=1), axis=0) * 100
    a["kpi"] = []
    for vv, r in kpi.sort_values("vertr").iterrows():
        a["kpi"].append({
            "vv": vv, "vertr": r1(r.vertr), "taxi": r1(r.taxi), "uur_plus": r1(r.uur_plus),
            "oorzaak_pct": {k.removeprefix("Oorzaak_").removesuffix("_min"): r1(oorz_pct.loc[vv, k])
                            for k in OORZAKEN.values()},
            # in minuten: dit is wat leerlingen in de draaitabel van 3.1 zien (Som van ...)
            "oorzaak_min": {k.removeprefix("Oorzaak_").removesuffix("_min"): int(oorz.loc[vv, k])
                            for k in OORZAKEN.values()},
        })
    totaal = df.groupby("Maatschappij").Vertrekvertraging_min.agg(["sum", "mean", "size"])
    totaal = totaal.sort_values("sum", ascending=False)
    a["kosten"] = [{"mij": k, "n": int(r["size"]), "som": int(r["sum"]), "gem": r1(r["mean"]),
                    "kosten": int(r["sum"] * KOSTEN_PER_MINUUT)} for k, r in totaal.head(5).iterrows()]
    hoogste_gem = totaal[totaal["size"] >= 100].sort_values("mean", ascending=False).iloc[0]
    a["hoogste_gem_mij"] = {"mij": hoogste_gem.name, "gem": r1(hoogste_gem["mean"]), "n": int(hoogste_gem["size"])}
    return a


def slechte_grafiek(df, pad):
    """Bewust slechte grafiek voor de MMI-opdracht: 3D-taart, regenboog, geen conclusie."""
    per = df.groupby("Maand").Vertrekvertraging_min.mean()
    fig = plt.figure(figsize=(7, 5), dpi=150)
    ax = fig.add_subplot(111)
    kleuren = plt.cm.gist_rainbow([i / 12 for i in range(12)])
    wedges, *_ = ax.pie(per.values, colors=kleuren, shadow=True, startangle=37,
                        explode=[0.06] * 12, autopct="%1.0f%%",
                        textprops={"fontsize": 7, "fontfamily": "fantasy"})
    ax.set_aspect(0.55)  # nep-3D: platgedrukte taart
    ax.legend(wedges, [str(m) for m in per.index], loc="center left", bbox_to_anchor=(1, 0.5),
              fontsize=7, title="Legenda")
    ax.set_title("GRAFIEK", fontsize=18, fontweight="bold", color="magenta", fontfamily="fantasy")
    fig.patch.set_facecolor("#e8e0a0")
    fig.tight_layout()
    fig.savefig(pad)


def voorbeeldgrafieken(df, map_):
    """Voorbeelden van 'goed' voor in het nakijkmodel."""
    grijs, accent, tekst = "#B8B8B8", "#C00A35", "#000000"  # UU-huisstijl: rood alleen als accent
    from matplotlib import font_manager
    for f in (ROOT / "fonts").glob("OpenSans-*.ttf"):
        font_manager.fontManager.addfont(str(f))
    plt.rcParams.update({"font.family": "Open Sans", "axes.spines.top": False,
                         "axes.spines.right": False, "axes.edgecolor": "#999999",
                         "axes.labelcolor": tekst, "xtick.color": tekst, "ytick.color": tekst})

    per = df.groupby("Maand").Vertrekvertraging_min.mean()
    fig, ax = plt.subplots(figsize=(7, 3.6), dpi=200)
    ax.bar(per.index, per.values, color=[accent if m == per.idxmax() else grijs for m in per.index])
    ax.set_xticks(range(1, 13), ["jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec"])
    ax.set_ylabel("gem. vertrekvertraging (min)")
    ax.set_title(f"Vertraagde vluchten lopen in {['januari','februari','maart','april','mei','juni','juli','augustus','september','oktober','november','december'][per.idxmax()-1]} het meest uit",
                 loc="left", fontsize=12, fontweight="bold", color=tekst)
    fig.tight_layout()
    fig.savefig(map_ / "goed_maand.png")
    plt.close(fig)

    per = df[df.Vertrekuur.between(6, 22)].groupby("Vertrekuur").Vertrekvertraging_min.mean()
    fig, ax = plt.subplots(figsize=(7, 3.6), dpi=200)
    ax.plot(per.index, per.values, color=accent, lw=2.5, marker="o", ms=4)
    ax.set_ylim(0, per.max() * 1.15)
    ax.set_xticks(range(6, 23, 2), [f"{u}:00" for u in range(6, 23, 2)])
    ax.set_xlabel("geplande vertrektijd")
    ax.set_ylabel("gem. vertrekvertraging (min)")
    ax.set_title("Hoe later op de dag, hoe groter de vertraging", loc="left", fontsize=12,
                 fontweight="bold", color=tekst)
    fig.tight_layout()
    fig.savefig(map_ / "goed_uur.png")
    plt.close(fig)

    per = df.groupby("Van").Taxi_uit_min.mean().sort_values()
    fig, ax = plt.subplots(figsize=(7, 3.2), dpi=200)
    ax.barh(per.index, per.values, color=[accent if v == per.idxmax() else grijs for v in per.index])
    for i, v in enumerate(per.values):
        ax.text(v + 0.5, i, f"{v:.0f} min", va="center", fontsize=9, color=tekst)
    ax.set_xlabel("gem. taxitijd tot opstijgen (min)")
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="y", length=0)
    ax.set_title("Op JFK sta je bijna twee keer zo lang in de rij voor de startbaan", loc="left",
                 fontsize=12, fontweight="bold", color=tekst)
    fig.tight_layout()
    fig.savefig(map_ / "goed_taxi.png")
    plt.close(fig)


def main():
    csv_pad = Path(sys.argv[1]).expanduser()
    df = vertaal(laad(csv_pad))
    (ROOT / "data").mkdir(exist_ok=True)
    schrijf_excel(df, ROOT / "data" / "vluchten_practicum.xlsx")
    df.to_csv(ROOT / "data" / "vluchten_practicum.csv", index=False, sep=";", decimal=",")
    (ROOT / "typst" / "antwoorden.json").write_text(
        json.dumps(antwoorden(df), indent=2, ensure_ascii=False, default=str))
    slechte_grafiek(df, ROOT / "typst" / "beeld" / "slechte_grafiek.png")
    voorbeeldgrafieken(df, ROOT / "typst" / "beeld")
    print(f"{len(df)} rijen geschreven naar {ROOT / 'data'}")


if __name__ == "__main__":
    main()
