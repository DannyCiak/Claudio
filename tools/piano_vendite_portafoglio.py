"""Genera il Piano Vendite Full Immersion dal report portafoglio Allianz."""
import re
import sys
import datetime as dt
from collections import defaultdict, OrderedDict

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment

SRC, OUT = sys.argv[1], sys.argv[2]
START = dt.date(2026, 10, 12)
TODAY = dt.date(2026, 10, 7)
AGENTE = "Dan Moscaliuc"
CAP_FERIALE = 15   # nuovi primi contatti lun-ven (full immersion)
CAP_SABATO = 5     # sabato: soprattutto appuntamenti + recall
GIORNI = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato", "Domenica"]

df = pd.read_excel(SRC)
df = df.fillna("")


# ----------------------------------------------------------------- utility
def d(s):
    try:
        return dt.datetime.strptime(str(s), "%d/%m/%Y").date()
    except ValueError:
        return None


def fmt(x):
    return x.strftime("%d/%m") if x else ""


def giorno(x):
    return f"{GIORNI[x.weekday()].lower()} {x.strftime('%d/%m')}"


VOW = set("AEIOU")


def _cons(s):
    return [c for c in s if c.isalpha() and c not in VOW]


def _vow(s):
    return [c for c in s if c in VOW]


def cf_surname(s):
    s = re.sub(r"[^A-Z]", "", s.upper())
    return ("".join(_cons(s) + _vow(s)) + "XXX")[:3]


def cf_name(s):
    s = re.sub(r"[^A-Z]", "", s.upper())
    c = _cons(s)
    if len(c) >= 4:
        return c[0] + c[2] + c[3]
    return ("".join(c + _vow(s)) + "XXX")[:3]


def split_name(full, cf):
    toks = full.split()
    cf = str(cf).upper()
    if len(cf) == 16 and len(toks) >= 2:
        for k in range(1, len(toks)):
            sur, nam = " ".join(toks[:k]), " ".join(toks[k:])
            if cf_surname(sur) == cf[:3] and cf_name(nam) == cf[3:6]:
                return sur, nam
    if len(toks) >= 2:
        return toks[0], " ".join(toks[1:])
    return full, ""


def tc(s):
    out = []
    for w in s.lower().split():
        out.append("'".join(p[:1].upper() + p[1:] for p in w.split("'")))
    return " ".join(out)


PROV_LONTANE = {"PR", "VR", "TN", "TS", "BL", "AG", "BZ"}


def cell_phone(v):
    if v == "" or v is None:
        return ""
    s = str(int(float(v)))
    if s.startswith("39") and len(s) > 10:
        return f"+39 {s[2:5]} {s[5:]}"
    return s


# ------------------------------------------------------- arricchimento dati
recs = []
for i, r in df.iterrows():
    cop = r["Cop. Cl"]
    mod = r["Mod. del Cl"]
    ini = r["Iniziative Cl"]
    impresa = r["Sesso"] == "G" or r["Cl Forma Giur."] not in ("",)
    nome_full = r["Contraente"].strip()
    cf = str(r["Cod.Fiscale / P.IVA"])
    titolare = ""
    if impresa:
        m = re.search(r"\bDI (.+)$", nome_full)
        if m:
            t = m.group(1).replace("& C.", "").strip()
            sur, nam = split_name(t, "")
            titolare = tc(nam) if nam else tc(t)
        elif len(cf) == 16:
            sur, nam = split_name(nome_full, cf)
            titolare = tc(nam)
        saluto = titolare or ""
        cognome = ""
    else:
        sur, nam = split_name(nome_full, cf)
        saluto, cognome = tc(nam), tc(sur)
    eta = int(r["Età"]) if r["Età"] != "" else None
    if impresa and r["Sesso"] == "G":
        eta = None
    prof = r["Professione"]
    sesso = r["Sesso"]
    # Lei per over 40, imprese e professioni "alte"; tu per i più giovani
    formale = impresa or eta is None or eta >= 40 or prof in ("Medico", "Dirigente")

    has_tcm = r["Scop. TCM"] == "No"
    has_sal = "ULTRA SALUTE" in cop
    has_casa = r["Scop. Casa"] == "No"
    has_ltc = "LONGEVITY" in cop
    has_prev = any(k in cop for k in ("FONDO PENSIONE", "F.P.A.", "INSIEME"))
    has_inv = "NUOVI ORIZZONTI" in cop
    has_auto = int(r["N. Pol. Auto Cl"] or 0) > 0
    has_infcirc = "INFORTUNI DA CIRCOLAZIONE" in cop
    has_uimp = "ULTRA IMPRESA" in cop
    has_catnat = "CATASTROFI NATURALI IMPRESA" in cop or "Clienti CON copertura" in ini
    sal_mods = []
    for k, lab in (("SPESE MEDICHE", "spese mediche"), ("DIARIA DA RICOVERO", "diaria da ricovero"),
                   ("INVALIDITA' PERMANENTE DA INFORTUNIO", "invalidità permanente da infortunio"),
                   ("INVALIDITA' PERMANENTE DA MALATTIA", "invalidità permanente da malattia")):
        if "AZUSA " + k in mod:
            sal_mods.append(lab)
    sal_miss = [x for x in ("invalidità permanente da malattia", "spese mediche", "diaria da ricovero",
                            "invalidità permanente da infortunio") if x not in sal_mods]
    _eta = int(r["Età"]) if r["Età"] != "" else 45
    if _eta >= 60 or r["Professione"] in ("Pensionato", "Casalinga", "Studente", "Benestante"):
        sal_miss = [x for x in sal_miss if x != "invalidità permanente da malattia"]
    if r["St. Cl"] == "C":
        has_tcm = has_sal = has_casa = has_ltc = has_prev = has_inv = has_auto = has_infcirc = has_uimp = False
        has_catnat = False
    scad = d(r["Dt. Prox Scad Quiet"])
    cliente_dal = d(r["Iniz. Rapp."])
    cessato = r["St. Cl"] == "C"
    disdetta = r["Disd 12 mm Cl"] == "Si"
    prod_cess = tc(r["Cop. Ces. Cl"].strip(" -").split(" - ")[0]) if r["Cop. Ces. Cl"] else ""
    doc = r["Des Doc Scad Cl"].strip(" -")
    privacy = r["Prvy Com"] == "Si"
    otp = r["OTP"] == "Si"
    ibips = r["Quest. Ibips"] == "Si"
    lavoratore = prof not in ("Pensionato", "Casalinga", "Studente", "Benestante", "Condomini",
                              "Associazioni Di Pubblica Assistenza") and not impresa
    autonomo = impresa or prof in ("Imprenditore", "Artigiano Edile", "Artigiano Manifatturiero", "Artigiano",
                                   "Agente", "Medico", "Industria Edile: Costruzione, Install. Impianti",
                                   "Collaboratore")
    ente = prof in ("Condomini", "Associazioni Di Pubblica Assistenza")
    premi = float(r["Premi Annui Cl"] or 0)
    npol = int(r["N. Pol. Tot. Cl"] or 0)

    # ---------------- punteggi di bisogno per prodotto (0-10)
    need = {}
    e = eta or 45
    if not impresa:
        if not has_tcm and 22 <= e <= 67 and prof not in ("Studente",):
            v = 7
            if 30 <= e <= 55:
                v += 2
            if has_casa or prof in ("Imprenditore", "Impiegato", "Funzionario/quadro", "Medico", "Dirigente"):
                v += 1
            if prof in ("Pensionato", "Casalinga"):
                v -= 3
            need["TCM"] = v
        if not has_sal and 18 <= e <= 72:
            v = 7 + (1 if lavoratore or autonomo else 0)
            need["SALUTE"] = v
        elif has_sal and sal_miss and e <= 70:
            need["SALUTE+"] = 5 + (2 if "invalidità permanente da malattia" in sal_miss and (lavoratore or autonomo) else 0)
        if not has_ltc and e >= 48:
            need["LTC"] = 6 + (2 if e >= 58 else 0)
        if not has_prev and 18 <= e <= 60 and (lavoratore or autonomo):
            need["PREV"] = 7 + (2 if 25 <= e <= 50 else 0)
        elif not has_prev and prof == "Casalinga" and e <= 58:
            need["PREV"] = 4
        elif not has_prev and prof == "Studente":
            need["PREV"] = 3
        if not has_inv and 25 <= e <= 78:
            v = 4
            if prof in ("Pensionato", "Benestante", "Imprenditore", "Impiegato", "Funzionario/quadro", "Medico",
                        "Dirigente", "Insegnante"):
                v += 2
            if premi >= 1000:
                v += 1
            need["INV"] = v
        if not has_casa:
            need["CASA"] = 4 if e >= 28 else 2
        if not has_auto and e >= 20 and prof != "Studente":
            need["AUTO"] = 5 if disdetta else 3
        elif has_auto and not has_infcirc and "Infortuni da circolazione" in ini:
            need["INFCIRC"] = 5
    else:
        if ente:
            if not has_casa and prof == "Condomini":
                need["CASA"] = 4
            need["IMPRESA"] = 4
        else:
            if not has_catnat:
                need["CATNAT"] = 10
            if not has_uimp:
                need["IMPRESA"] = 7
            need["KEYMAN"] = 6 if not has_tcm else 0
            if not has_sal:
                need["SALUTE"] = 6
            if "Previdenza" in ini or "PREVIDENZA" in ini:
                need["WELFARE"] = 6
            if not has_prev:
                need.setdefault("PREV", 5)
    # bonus iniziative Allianz già assegnate
    bonus_map = {"LOVIA": "TCM", "SALUTE": "SALUTE", "MALATTIA": "SALUTE", "INFORTUNI": "SALUTE",
                 "PREVIDENZA": "PREV", "Casa e Patrimonio": "CASA", "Longevity": "LTC", "CAT NAT": "CATNAT",
                 "Lithium": "AUTO", "Infortuni da circolazione": "INFCIRC", "Previdenza": "WELFARE",
                 "AZULTRA Impresa": "IMPRESA"}
    for k, p in bonus_map.items():
        if k in ini:
            if p == "SALUTE" and "SALUTE" not in need and "SALUTE+" in need:
                p = "SALUTE+"
            if p in need and need[p] > 0:
                need[p] += 2
    need = {k: v for k, v in need.items() if v > 0}
    # i prodotti "persona" sono il focus richiesto: piccolo peso aggiuntivo
    persona = {"TCM", "SALUTE", "SALUTE+", "LTC", "PREV", "INV", "KEYMAN", "WELFARE"}
    rank = sorted(need.items(), key=lambda kv: -(kv[1] + (1.5 if kv[0] in persona else 0) + (3 if kv[0] == "CATNAT" else 0)))
    prodotti = [k for k, _ in rank]

    # ---------------- score cliente 0-100
    opp = sum(v for _, v in rank[:3])
    s_opp = min(opp, 30) / 30 * 40
    s_val = min(premi, 2000) / 2000 * 15
    s_rel = min(npol, 3) / 3 * 10 + (5 if cliente_dal and (TODAY - cliente_dal).days > 3 * 365 else 0)
    days_to = (scad - START).days if scad else 999
    s_urg = 15 if -15 <= days_to <= 35 else (8 if days_to <= 70 else 0)
    s_age = 10 if 30 <= e <= 60 else (6 if 25 <= e <= 70 else 2)
    s_ini = 5 if ini else 0
    s_dis = 5 if disdetta else 0
    score = round(min(100, s_opp + s_val + s_rel + s_urg + s_age + s_ini + s_dis))
    if cessato:
        score = min(score, 40)
    classe = "A" if score >= 62 else ("B" if score >= 45 else "C")

    recs.append(dict(
        idx=i, nome_full=nome_full, saluto=saluto, cognome=cognome, titolare=titolare, impresa=impresa,
        ente=ente, eta=eta, sesso=sesso, prof=prof, settore=r["U.D.M."], comune=tc(r["Comune"]),
        prov=r["Prov."], cell=cell_phone(r["Cellulare"]), cell_raw=str(r["Cellulare"]), email=r["Email"],
        formale=formale, has_tcm=has_tcm, has_sal=has_sal, has_casa=has_casa, has_ltc=has_ltc,
        has_prev=has_prev, has_inv=has_inv, has_auto=has_auto, has_infcirc=has_infcirc, has_uimp=has_uimp,
        has_catnat=has_catnat, sal_mods=sal_mods, sal_miss=sal_miss, scad=scad, cliente_dal=cliente_dal,
        cessato=cessato, disdetta=disdetta, prod_cess=prod_cess, doc=doc, privacy=privacy, otp=otp,
        ibips=ibips, lavoratore=lavoratore, autonomo=autonomo, premi=premi, npol=npol, cop=cop, mod=mod,
        ini=ini, need=need, prodotti=prodotti, score=score, classe=classe, fine_rapp=r["Dt. Fine Rapporto"],
        via=r["Via"], cap=str(r["C.A.P."]), ind=r["Indirizzo"]))

# ---------------------------------------------------- famiglie / duplicati
by_phone = defaultdict(list)
for c in recs:
    key = c["cell_raw"] if c["cell_raw"] not in ("", "nan") else "N" + c["nome_full"]
    by_phone[key].append(c)
by_addr = defaultdict(list)
for c in recs:
    k = re.sub(r"[^a-z0-9]", "", c["via"].lower()) + c["cap"]
    by_addr[k].append(c)
for k, lst in by_addr.items():
    names = {x["nome_full"] for x in lst}
    for c in lst:
        others = sorted({x["nome_full"] for x in lst if x["nome_full"] != c["nome_full"]})
        c["famiglia"] = others if len(names) > 1 else []
for k, lst in by_phone.items():
    for c in lst:
        c["stesso_tel"] = sorted({x["nome_full"] for x in lst if x is not c})

recs.sort(key=lambda c: c["nome_full"])
NOMI_PERSONE = {c["nome_full"] for c in recs if not c["impresa"]}
for n, c in enumerate(recs, 1):
    c["id"] = f"C{n:03d}"

# ------------------------------------------------------------ testi prodotti
PNAME = {
    "TCM": "Allianz Lovia (TCM – protezione famiglia)",
    "SALUTE": "Allianz Ultra Salute",
    "SALUTE+": "Completamento moduli Ultra Salute",
    "LTC": "Allianz Longevity Care (non autosufficienza)",
    "PREV": "Fondo Pensione Aperto Allianz Previdenza",
    "INV": "Piano di risparmio/investimento vita (es. Nuovi Orizzonti)",
    "CASA": "Allianz Ultra Casa e Patrimonio",
    "AUTO": "Auto (Bonus Malus) – recupero/confronto",
    "INFCIRC": "Infortuni da circolazione",
    "CATNAT": "Catastrofi Naturali Impresa (obbligatoria)",
    "IMPRESA": "Allianz Ultra Impresa (danni + RC)",
    "KEYMAN": "Lovia sul titolare (TCM persona chiave)",
    "WELFARE": "Previdenza/welfare per titolare e dipendenti",
}
PSHORT = {"TCM": "TCM Lovia", "SALUTE": "Ultra Salute", "SALUTE+": "Upgrade Salute", "LTC": "Longevity Care",
          "PREV": "Fondo pensione", "INV": "Risparmio/Invest.", "CASA": "Ultra Casa", "AUTO": "Auto",
          "INFCIRC": "Infortuni circolaz.", "CATNAT": "CAT NAT impresa", "IMPRESA": "Ultra Impresa",
          "KEYMAN": "TCM persona chiave", "WELFARE": "Welfare/previdenza"}


def T(c, tu, lei):
    return lei if c["formale"] else tu


def possesso(c):
    p = []
    if c["has_auto"]:
        p.append("l'auto")
    if c["has_casa"]:
        p.append("la casa")
    if c["has_sal"]:
        p.append("la salute")
    if c["has_tcm"]:
        p.append("la protezione famiglia (Lovia)")
    if c["has_uimp"]:
        p.append("l'attività")
    if c["has_prev"]:
        p.append("il fondo pensione")
    if c["has_inv"]:
        p.append("un piano di risparmio")
    if not p:
        return ""
    return ", ".join(p[:-1]) + (" e " if len(p) > 1 else "") + p[-1]


def perche(c, p):
    e = c["eta"]
    if p == "TCM":
        return (f"Nessuna copertura vita; {e} anni, {c['prof'].lower()}: il reddito va protetto"
                + ("; ha la casa assicurata → probabile mutuo/famiglia" if c["has_casa"] else "") + ".")
    if p == "SALUTE":
        return "Nessuna copertura salute/infortuni: spese mediche, ricovero e invalidità tutte scoperte."
    if p == "SALUTE+":
        return f"Ha Ultra Salute ({', '.join(c['sal_mods']) or 'moduli base'}); mancano: {', '.join(c['sal_miss'])}."
    if p == "LTC":
        return f"{e} anni, nessuna tutela per la non autosufficienza: è l'età giusta per pianificarla."
    if p == "PREV":
        if c["prof"] == "Studente":
            return "Studente: fondo pensione aperto anche a carico dei genitori (deducibile per loro), parte presto."
        if c["prof"] == "Casalinga":
            return "Casalinga senza previdenza propria: fondo pensione anche come familiare a carico."
        return f"{e} anni, lavora e non ha previdenza complementare: gap pensionistico + deduzione fiscale persa."
    if p == "INV":
        return "Nessun prodotto di risparmio/investimento con noi: verificare liquidità ferma sul conto."
    if p == "CASA":
        return "Casa non assicurata con noi (o solo polizza della banca)."
    if p == "AUTO":
        return (f"Disdetta recente ({c['prod_cess']}): recuperare il rapporto." if c["disdetta"]
                else "Nessuna polizza auto con noi: chiedere scadenza attuale e fare confronto.")
    if p == "INFCIRC":
        return "Ha l'auto con noi ma non gli Infortuni da circolazione; c'è un'iniziativa attiva ott–dic 2026."
    if p == "CATNAT":
        return "Impresa senza CAT NAT: obbligo di legge, senza copertura rischia l'esclusione da contributi pubblici."
    if p == "IMPRESA":
        return "Nessuna polizza danni/RC dell'attività con noi."
    if p == "KEYMAN":
        return "Il titolare non ha una TCM: se si ferma lui, si ferma l'impresa."
    if p == "WELFARE":
        return "Iniziativa Allianz Previdenza per imprese fino a 59 dipendenti: welfare e fondo pensione."
    return ""


def bridge(c, p):
    """Frase-ponte sul bisogno, da dire al telefono."""
    t = _bridge(c, p)
    if not possesso(c):
        for pre in ("ma niente sulla ", "ma la casa no. ", "e c'è un tema", "e mi piacerebbe", "ma non c'è niente", "ma niente "):
            if t.startswith(pre):
                if pre == "ma niente sulla ":
                    t = T(c, "Vorrei parlarti della ", "Vorrei parlarle della ") + t[len(pre):]
                elif pre == "ma la casa no. ":
                    t = T(c, "Vorrei parlarti della casa. ", "Vorrei parlarle della casa. ") + t[len(pre):]
                elif pre == "ma non c'è niente":
                    t = "Oggi non c'è niente" + t[len(pre):]
                else:
                    t = t[2].upper() + t[3:]
                break
    return t


def _bridge(c, p):
    pos = possesso(c)
    base = T(c, f"Oggi con noi hai già {pos}, ", f"Oggi con noi ha già {pos}, ") if pos else ""
    if p == "TCM":
        return base + T(c,
            "ma non c'è niente che protegga la tua famiglia e il tuo reddito se dovesse succedere qualcosa a te. "
            "Ci hai mai pensato?",
            "ma non c'è niente che protegga la sua famiglia e il suo reddito se dovesse succedere qualcosa a lei. "
            "È un tema a cui ha mai pensato?")
    if p in ("SALUTE",):
        return base + T(c,
            "ma niente sulla salute. Con le liste d'attesa di oggi molti clienti mi chiedono come fare visite ed esami "
            "in tempi rapidi senza pagare tutto di tasca propria. Tu come ti organizzi?",
            "ma niente sulla salute. Con le liste d'attesa di oggi molti clienti mi chiedono come fare visite ed esami "
            "in tempi rapidi senza pagare tutto di tasca propria. Lei come si organizza?")
    if p == "SALUTE+":
        cosa = {"invalidità permanente da malattia": "è la parte che protegge il reddito quando una malattia impedisce di lavorare",
                "spese mediche": "è la parte che rimborsa visite, esami e interventi",
                "diaria da ricovero": "è un'indennità per ogni giorno di ricovero, utile per le spese extra",
                "invalidità permanente da infortunio": "è la parte che paga un capitale se un infortunio lascia conseguenze permanenti"}[c["sal_miss"][0]]
        return T(c, f"La tua Ultra Salute copre già {', '.join(c['sal_mods']) or 'la parte base'}, ma manca {c['sal_miss'][0]}: {cosa}. Vale la pena guardarla insieme.",
                 f"La sua Ultra Salute copre già {', '.join(c['sal_mods']) or 'la parte base'}, ma manca {c['sal_miss'][0]}: {cosa}. Vale la pena guardarla insieme.")
        return T(c,
            f"La tua Ultra Salute copre già {', '.join(c['sal_mods']) or 'la parte base'}, ma manca "
            f"{c['sal_miss'][0]}: è la parte che protegge il reddito quando non si può lavorare. Vale la pena "
            "guardarla insieme.",
            f"La sua Ultra Salute copre già {', '.join(c['sal_mods']) or 'la parte base'}, ma manca "
            f"{c['sal_miss'][0]}: è la parte che protegge il reddito quando non si può lavorare. Vale la pena "
            "guardarla insieme.")
    if p == "LTC":
        return base + T(c,
            "e c'è un tema di cui si parla poco: la non autosufficienza. Chi ha assistito un familiare sa quanto costa. "
            "Ti è mai capitato da vicino?",
            "e c'è un tema di cui si parla poco: la non autosufficienza. Chi ha assistito un familiare sa quanto costa. "
            "Le è mai capitato da vicino?")
    if p == "PREV":
        if c["prof"] == "Studente":
            return ("Per chi è giovane il fondo pensione è lo strumento più efficiente: partire presto, anche con poco, "
                    "fa la differenza. I versamenti possono farli anche i genitori e dedurli loro.")
        return base + T(c,
            "ma niente sulla pensione. Per chi ha la tua età la pensione pubblica sarà più bassa dell'ultimo stipendio. "
            "Con il fondo pensione costruisci un'integrazione e risparmi tasse: i versamenti si deducono fino a "
            "5.164,57 € l'anno. Sai già dove va il tuo TFR?",
            "ma niente sulla pensione. Per chi ha la sua età la pensione pubblica sarà più bassa dell'ultimo reddito. "
            "Con il fondo pensione si costruisce un'integrazione e si risparmiano tasse: i versamenti si deducono fino a "
            "5.164,57 € l'anno. Sa già come è messo con la previdenza?")
    if p == "INV":
        return base + T(c,
            "e mi piacerebbe capire come sono messi i tuoi risparmi: i soldi fermi sul conto ogni anno perdono valore "
            "con l'inflazione. Si può partire anche con piccole cifre mensili.",
            "e mi piacerebbe capire come sono gestiti i suoi risparmi: la liquidità ferma sul conto ogni anno perde "
            "valore con l'inflazione. Possiamo vedere una soluzione adatta al suo profilo.")
    if p == "CASA":
        return base + T(c,
            "ma la casa no. Tra catastrofi naturali, RC della famiglia e furto, sono rischi che costano poco da coprire. "
            "Hai già qualcosa, magari con la banca del mutuo?",
            "ma la casa no. Tra catastrofi naturali, RC della famiglia e furto, sono rischi che costano poco da coprire. "
            "Ha già qualcosa, magari con la banca del mutuo?")
    if p == "AUTO":
        return T(c,
            "Mi farebbe piacere sapere quando scade l'assicurazione della tua auto: ti preparo un confronto senza impegno.",
            "Mi farebbe piacere sapere quando scade l'assicurazione della sua auto: le preparo un confronto senza impegno.")
    if p == "INFCIRC":
        return T(c,
            "Hai l'auto con noi ma non la copertura infortuni del conducente: se ti fai male in un incidente con colpa, "
            "l'RC auto non paga te. Costa poco e fino a dicembre c'è un'iniziativa dedicata.",
            "Ha l'auto con noi ma non la copertura infortuni del conducente: se si fa male in un incidente con colpa, "
            "l'RC auto non paga lei. Costa poco e fino a dicembre c'è un'iniziativa dedicata.")
    if p == "CATNAT":
        return ("La chiamo per una verifica importante: la polizza catastrofi naturali è obbligatoria per le imprese, "
                "e senza si rischia di perdere l'accesso a contributi e agevolazioni pubbliche. Dalla nostra anagrafica "
                "la vostra attività non risulta coperta: avete già provveduto altrove?")
    if p == "IMPRESA":
        return ("Vorrei fare con lei il punto sui rischi dell'attività: danni ai beni, responsabilità civile verso "
                "clienti e terzi, fermo attività. Oggi con noi non risulta nulla di questo.")
    if p == "KEYMAN":
        return ("C'è una cosa che vedo spesso trascurata nelle piccole imprese: se si ferma il titolare, si ferma tutto. "
                "Una protezione sulla persona chiave tutela l'attività e la famiglia.")
    if p == "WELFARE":
        return ("Allianz ha un'iniziativa per le imprese fino a 59 dipendenti su previdenza e welfare: è un modo per "
                "fidelizzare i dipendenti con vantaggi fiscali per l'azienda. Vorrei mostrargliela.")
    return ""


DISCOVERY = {
    "TCM": ["Chi dipende economicamente da te (partner, figli, genitori)?",
            "Hai un mutuo o finanziamenti? Quanto manca e fino a quando?",
            "Se mancasse il tuo reddito, per quanti anni la famiglia manterrebbe lo stesso tenore di vita?",
            "Hai già coperture vita (abbinate al mutuo, al lavoro, alla banca)? Con quale capitale?"],
    "SALUTE": ["Nell'ultimo anno quante visite/esami hai fatto privatamente e quanto hai speso?",
               "Il tuo contratto di lavoro prevede un fondo sanitario di categoria? Cosa copre?",
               "Se non potessi lavorare per 3 mesi, chi coprirebbe le spese di casa?",
               "Ci sono patologie pregresse da dichiarare nel questionario sanitario?"],
    "SALUTE+": ["Hai mai usato la tua Ultra Salute? Com'è andata?",
                "Cosa succederebbe al tuo reddito se una malattia ti impedisse di lavorare per mesi?",
                "Sei soddisfatto dei massimali attuali o vorresti alzarli?",
                "Ci sono altri familiari da inserire nella copertura?"],
    "LTC": ["Hai genitori anziani o hai vissuto da vicino una situazione di non autosufficienza?",
            "Se avessi bisogno di assistenza continuativa, su chi conteresti?",
            "Preferiresti restare a casa con assistenza o una struttura?",
            "Che entrate avrai in pensione? Basterebbero per pagare un'assistenza?"],
    "PREV": ["A che età pensi di andare in pensione e con quale reddito ti piacerebbe vivere?",
             "Dove va oggi il tuo TFR (azienda, Fondo Tesoreria INPS, fondo di categoria)?",
             "Hai consultato l'estratto conto contributivo su MyINPS?",
             "Quanto riesci a mettere da parte ogni mese senza rinunce?"],
    "INV": ["Hai risparmi fermi sul conto? Con che orizzonte potresti investirli?",
            "Qual è l'obiettivo: casa, studi dei figli, pensione, tranquillità, eredità?",
            "Come reagiresti se l'investimento perdesse il 10% in un anno?",
            "Obbligatorio prima della proposta: questionario di adeguatezza IBIPs."],
    "CASA": ["Casa di proprietà o in affitto? Con mutuo?",
             "Il mutuo ha una polizza della banca? Cosa copre davvero (spesso solo incendio del fabbricato)?",
             "Figli, animali domestici, attività nel tempo libero (RC famiglia)?",
             "Zona soggetta ad allagamenti, grandine, eventi atmosferici?"],
    "AUTO": ["Quando scade la polizza auto attuale e con chi è?",
             "Perché hai cambiato compagnia: prezzo, servizio, sinistro?",
             "Quali garanzie hai oggi (furto/incendio, kasko, cristalli, assistenza)?",
             "Altri veicoli in famiglia?"],
    "INFCIRC": ["Quanti km fai all'anno e per cosa (lavoro/casa)?",
                "Sai che l'RC auto non risarcisce il conducente che causa l'incidente?",
                "Chi altro guida l'auto?",
                "Hai altre coperture infortuni?"],
    "CATNAT": ["Avete già sottoscritto la polizza catastrofi naturali obbligatoria (anche con altre compagnie)?",
               "Valore di fabbricato, impianti, macchinari e attrezzature?",
               "La sede è in proprietà o in affitto?",
               "Avete in programma richieste di contributi o finanziamenti agevolati?"],
    "IMPRESA": ["Quanti dipendenti/collaboratori avete?",
                "Lavorate presso clienti o in cantiere? Che danni potreste causare a terzi?",
                "Se l'attività si fermasse 2 mesi per un danno, come andreste avanti?",
                "Che polizze avete oggi e con chi?"],
    "KEYMAN": ["Se lei si fermasse per malattia o infortunio, l'attività andrebbe avanti?",
               "Ci sono soci, debiti aziendali o fidi garantiti personalmente?",
               "Chi in famiglia dipende dal reddito dell'impresa?",
               "Ha coperture personali vita/infortuni?"],
    "WELFARE": ["Quanti dipendenti avete e con che contratto?",
                "Il TFR dei dipendenti dove va oggi?",
                "Avete già un piano welfare aziendale?",
                "Le interessa ridurre il cuneo fiscale con benefit deducibili?"],
}

DISCOVERY_LEI = {
    "TCM": ["Chi dipende economicamente da lei (coniuge, figli, genitori)?",
            "Ha un mutuo o finanziamenti? Quanto manca e fino a quando?",
            "Se venisse a mancare il suo reddito, per quanti anni la famiglia manterrebbe lo stesso tenore di vita?",
            "Ha già coperture vita (abbinate al mutuo, al lavoro, alla banca)? Con quale capitale?"],
    "SALUTE": ["Nell'ultimo anno quante visite/esami ha fatto privatamente e quanto ha speso?",
               "Il suo contratto di lavoro prevede un fondo sanitario di categoria? Cosa copre?",
               "Se non potesse lavorare per 3 mesi, chi coprirebbe le spese di casa?",
               "Ci sono patologie pregresse da dichiarare nel questionario sanitario?"],
    "SALUTE+": ["Ha mai usato la sua Ultra Salute? Com'è andata?",
                "Cosa succederebbe al suo reddito se una malattia le impedisse di lavorare per mesi?",
                "È soddisfatto dei massimali attuali o vorrebbe alzarli?",
                "Ci sono altri familiari da inserire nella copertura?"],
    "LTC": ["Ha genitori anziani o ha vissuto da vicino una situazione di non autosufficienza?",
            "Se avesse bisogno di assistenza continuativa, su chi potrebbe contare?",
            "Preferirebbe restare a casa con assistenza o una struttura?",
            "Che entrate avrà in pensione? Basterebbero per pagare un'assistenza?"],
    "PREV": ["A che età pensa di andare in pensione e con quale reddito le piacerebbe vivere?",
             "Dove va oggi il suo TFR (azienda, Fondo Tesoreria INPS, fondo di categoria)? Se è autonomo: quanto versa di contributi?",
             "Ha consultato l'estratto conto contributivo su MyINPS?",
             "Quanto riesce a mettere da parte ogni mese senza rinunce?"],
    "INV": ["Ha risparmi fermi sul conto? Con che orizzonte potrebbe investirli?",
            "Qual è l'obiettivo: casa, studi dei figli, pensione, tranquillità, eredità?",
            "Come reagirebbe se l'investimento perdesse il 10% in un anno?",
            "Obbligatorio prima della proposta: questionario di adeguatezza IBIPs."],
    "CASA": ["Casa di proprietà o in affitto? Con mutuo?",
             "Il mutuo ha una polizza della banca? Cosa copre davvero (spesso solo incendio del fabbricato)?",
             "Figli, animali domestici, attività nel tempo libero (RC famiglia)?",
             "Zona soggetta ad allagamenti, grandine, eventi atmosferici?"],
    "AUTO": ["Quando scade la polizza auto attuale e con chi è?",
             "Perché ha cambiato compagnia: prezzo, servizio, sinistro?",
             "Quali garanzie ha oggi (furto/incendio, kasko, cristalli, assistenza)?",
             "Altri veicoli in famiglia?"],
    "INFCIRC": ["Quanti km fa all'anno e per cosa (lavoro/casa)?",
                "Sa che l'RC auto non risarcisce il conducente che causa l'incidente?",
                "Chi altro guida l'auto?",
                "Ha altre coperture infortuni?"],
}
for _k in ("CATNAT", "IMPRESA", "KEYMAN", "WELFARE"):
    DISCOVERY_LEI[_k] = DISCOVERY[_k]

OBIEZIONI = {
    "TCM": ("\"Non ci voglio pensare / porta sfortuna\" oppure \"costa troppo\"",
            "Lo capisco, nessuno ama parlarne: proprio per questo lo si fa una volta, bene, e poi non ci si pensa più. "
            "Non parliamo di morte ma di chi paga mutuo e spese al posto tuo. Sul costo: quanto pensi che costi? "
            "Spesso è meno della bolletta del telefono e si può partire da un capitale più basso."),
    "SALUTE": ("\"C'è già il servizio sanitario / ho il fondo di categoria\"",
               "Ottimo, partiamo da lì: verifichiamo cosa copre davvero il fondo e integriamo solo quello che manca, "
               "senza doppioni. Se sei già coperto bene, te lo dico io."),
    "SALUTE+": ("\"Ho già la polizza salute, basta così\"",
                "Hai fatto la scelta giusta a partire. Il modulo che manca protegge il reddito, non le spese: è "
                "proprio la parte che serve quando non puoi lavorare. Ti mostro la differenza in 5 minuti."),
    "LTC": ("\"È troppo presto / ci penseranno i figli\"",
            "Molti la scelgono proprio per non pesare sui figli. E prima si parte, più il costo è sostenibile."),
    "PREV": ("\"Preferisco tenere i soldi disponibili / non mi fido\"",
             "Giusto tenere una riserva liquida: il fondo si alimenta solo con la parte che non ti serve nel breve. "
             "Ogni euro versato riduce le tasse, e in caso di necessità (prima casa, spese sanitarie) si può chiedere "
             "un'anticipazione. Il fondo è un patrimonio separato e vigilato da COVIP."),
    "INV": ("\"Ho già la banca\"",
            "Diversificare è sempre bene. Guardiamo insieme cosa hai, costi e obiettivi: se quello che hai è adatto "
            "a te, te lo dirò io."),
    "CASA": ("\"Ho già la polizza della banca\"",
             "Spesso copre solo l'incendio del fabbricato, a favore della banca: contenuto, RC della famiglia, "
             "catastrofi naturali e furto restano scoperti. Portami la polizza e la confrontiamo."),
    "AUTO": ("\"Ho trovato più conveniente altrove\"",
             "Ci sta: dimmi solo la scadenza e ti preparo un confronto a parità di garanzie. Se resta più "
             "conveniente, almeno sai di avere la scelta giusta."),
    "INFCIRC": ("\"Ho già l'RC auto\"",
                "L'RC paga i danni che causi agli altri, non quelli che subisci tu se sei alla guida e hai torto. "
                "Questa copertura protegge te."),
    "CATNAT": ("\"Lo faccio più avanti\"",
               "L'obbligo è già in vigore: senza polizza si rischia di perdere contributi e agevolazioni pubbliche. "
               "Facciamo subito il preventivo: con visura e valori in mano bastano 15 minuti."),
    "IMPRESA": ("\"Ho già il mio broker / non ho tempo\"",
                "Nessun problema: le chiedo solo di farmi vedere le polizze. Se sono adeguate glielo confermo per "
                "iscritto, altrimenti le segnalo i buchi. 20 minuti in azienda."),
    "KEYMAN": ("\"Ci penso io, non serve\"",
               "Proprio perché lei è l'impresa, se si ferma lei si ferma tutto. Vediamo solo i numeri: debiti, "
               "fidi, famiglia. Decide lei."),
    "WELFARE": ("\"Costi in più per l'azienda\"",
                "Il welfare ben strutturato è deducibile e spesso costa meno di un aumento in busta paga, a parità "
                "di beneficio per il dipendente."),
}
OBIEZIONI_COMUNI = ("\"Mandami qualcosa via mail\" → Volentieri, ma senza conoscere la situazione ti manderei un "
                    "preventivo generico. 30 minuti e ti porto una proposta su misura. | \"Ne parlo con mia "
                    "moglie/mio marito\" → Giustissimo: facciamo l'appuntamento insieme, così fate le domande entrambi.")

PROPOSTA = {
    "TCM": "Lovia in 3 opzioni: A) capitale = debito residuo/mutuo; B) mutuo + 3 anni di reddito; C) metodo DIME "
           "completo (debiti + 5 anni di reddito + mutuo + studi figli). Durata fino a fine mutuo o "
           "all'indipendenza dei figli.",
    "SALUTE": "Ultra Salute in 3 configurazioni: A) invalidità permanente da infortunio + spese mediche grandi "
              "interventi; B) + diaria da ricovero; C) + invalidità permanente da malattia e spese mediche estese.",
    "SALUTE+": "Upgrade Ultra Salute: aggiungere i moduli mancanti e alzare il livello (Essential → Plus → Premium/Top) "
               "dove serve.",
    "LTC": "Longevity Care: rendita mensile in caso di non autosufficienza; simulare 2 importi di rendita.",
    "PREV": "Fondo Pensione Aperto Allianz Previdenza: simulazione con 2 contributi (es. 50 e 100 €/mese) + "
            "risparmio fiscale + eventuale conferimento TFR.",
    "INV": "Piano di accumulo/risparmio vita (es. Nuovi Orizzonti) da 50–100 €/mese o versamento unico, SOLO dopo "
           "questionario IBIPs e coerente con il profilo di rischio.",
    "CASA": "Ultra Casa e Patrimonio: fabbricato + contenuto + RC famiglia + catastrofi naturali (+ furto, tutela legale).",
    "AUTO": "Preventivo auto alla scadenza a parità di garanzie + infortuni conducente.",
    "INFCIRC": "Infortuni da circolazione abbinata alla polizza auto (iniziativa ott–dic 2026).",
    "CATNAT": "Catastrofi Naturali Impresa con i valori di fabbricato/impianti/merci; consegna certificato di copertura.",
    "IMPRESA": "Ultra Impresa: RC impresa + danni ai beni + fermo attività, configurata sul codice ATECO.",
    "KEYMAN": "Lovia sul titolare con capitale = debiti aziendali garantiti + 1–2 anni di fatturato netto.",
    "WELFARE": "Fondo pensione/welfare per titolare e dipendenti (iniziativa Mondo Impresa Previdenza).",
}
DOCS = {
    "TCM": "piano di ammortamento del mutuo, ultima busta paga o dichiarazione dei redditi",
    "SALUTE": "eventuali referti/terapie in corso (per il questionario sanitario), info fondo sanitario di categoria",
    "SALUTE+": "polizza Ultra Salute attuale",
    "LTC": "nessun documento particolare",
    "PREV": "estratto conto contributivo MyINPS, ultima busta paga/CU, info sul TFR",
    "INV": "estratti conto di investimenti/risparmi attuali",
    "CASA": "metri quadri, anno di costruzione, polizza della banca se presente",
    "AUTO": "libretto e polizza auto attuale (attestato di rischio)",
    "INFCIRC": "nessuno",
    "CATNAT": "visura camerale, valori di fabbricato, impianti, macchinari e merci",
    "IMPRESA": "visura camerale, fatturato, n. dipendenti, polizze in essere",
    "KEYMAN": "situazione debiti aziendali/fidi",
    "WELFARE": "n. dipendenti, CCNL applicato",
}

INIT_HOOK = [
    ("CAT NAT Impresa 2026   Clienti SENZA", "verifica sull'obbligo della polizza catastrofi naturali per l'impresa", ["CATNAT"]),
    ("CROSS SELL LOVIA", "Allianz ha riservato ai clienti come te una valutazione dedicata sulla protezione della famiglia (Lovia)", ["TCM"]),
    ("RAFFORZA LA PROTEZIONE SALUTE", "c'è un'iniziativa riservata ai clienti per rafforzare la protezione salute", ["SALUTE", "SALUTE+"]),
    ("Allianz Ultra Salute 2° sem", "c'è un'iniziativa Ultra Salute del secondo semestre riservata ai clienti", ["SALUTE", "SALUTE+"]),
    ("CROSS SELL MALATTIA", "c'è un'iniziativa dedicata alla copertura malattia per chi è già cliente", ["SALUTE", "SALUTE+"]),
    ("CROSS SELL INFORTUNI", "c'è un'iniziativa dedicata alla copertura infortuni per chi è già cliente", ["SALUTE", "SALUTE+"]),
    ("Clienti Vita   proposta SALUTE", "per chi ha già una polizza vita c'è una proposta salute dedicata", ["SALUTE", "SALUTE+"]),
    ("Mondo Persona Clienti+ | PREVIDENZA", "sto facendo con i clienti il punto sulla previdenza prima di fine anno", ["PREV"]),
    ("Proponi AZ Longevity", "c'è una novità Allianz sulla tutela della non autosufficienza (Longevity Care)", ["LTC"]),
    ("Mondo Impresa Clienti+ Previdenza", "c'è un'iniziativa Allianz su previdenza e welfare per le imprese", ["WELFARE"]),
    ("AZULTRA Impresa", "c'è il restyling di Ultra Impresa con una promozione riservata", ["IMPRESA"]),
    ("Mondo Impresa Clienti+ | SMEG", "c'è un'iniziativa Allianz dedicata alle PMI clienti", ["IMPRESA", "KEYMAN"]),
]



def hook(c):
    """Motivo della chiamata: (codice, testo lungo, testo breve per WhatsApp)."""
    tu = not c["formale"]
    if c["cessato"]:
        st = "stata" if c["sesso"] == "F" else "stato"
        return ("WIN-BACK",
                T(c, f"Ti chiamo perché sei {st} un nostro cliente e mi farebbe piacere sapere come ti trovi oggi: "
                     "sto facendo un check-up gratuito della protezione e vorrei proportelo.",
                  f"La chiamo perché è {st} nostro cliente e mi farebbe piacere sapere come si trova oggi: sto "
                  "facendo un check-up gratuito della protezione e vorrei proporglielo."),
                "vorrei sapere come ti trovi e proporti un check-up gratuito" if tu else
                "vorrei sapere come si trova e proporle un check-up gratuito")
    if c["disdetta"] and c["prod_cess"]:
        return ("RETENTION",
                T(c, f"Ho visto che quest'anno la polizza {c['prod_cess']} non è più con noi: volevo capire com'è "
                     "andata, se c'è qualcosa che potevo fare meglio, e rifare con te il punto su tutte le coperture.",
                  f"Ho visto che quest'anno la polizza {c['prod_cess']} non è più con noi: volevo capire com'è "
                  "andata, se c'è qualcosa che potevo fare meglio, e rifare con lei il punto su tutte le coperture."),
                f"ho visto che la polizza {c['prod_cess']} non è più con noi e vorrei capire com'è andata")
    if c["impresa"] and "CATNAT" in c["need"]:
        return ("OBBLIGO CAT NAT", bridge(c, "CATNAT"), "devo fare una verifica sull'obbligo della polizza catastrofi naturali per l'impresa")
    if c["scad"] and (c["scad"] - START).days <= 35:
        when = "è scaduta il" if c["scad"] < START else "scade il"
        if c["scad"] < START:
            dd = c["scad"].strftime("%d/%m")
            return ("RINNOVO",
                    T(c, f"Ho in evidenza la quietanza del {dd}: la sistemiamo insieme e, già che ci vediamo, ti propongo 30 minuti di check-up gratuito della protezione.",
                      f"Ho in evidenza la quietanza del {dd}: la sistemiamo insieme e, già che ci vediamo, le propongo 30 minuti di check-up gratuito della protezione."),
                    f"ho in evidenza la quietanza del {dd} e vorrei sistemarla insieme")
        return ("RINNOVO",
                T(c, f"La prossima quietanza {when} {c['scad'].strftime('%d/%m')}: invece della solita chiamata per "
                     "il pagamento, ti propongo 30 minuti per fare un check-up completo della tua protezione, gratuito.",
                  f"La prossima quietanza {when} {c['scad'].strftime('%d/%m')}: invece della solita chiamata per il "
                  "pagamento, le propongo 30 minuti per fare un check-up completo della sua protezione, gratuito."),
                f"la tua quietanza {when} {c['scad'].strftime('%d/%m')} e vorrei approfittarne per un check-up" if tu
                else f"la sua quietanza {when} {c['scad'].strftime('%d/%m')} e vorrei approfittarne per un check-up")
    if c["doc"]:
        docbreve = c["doc"].split("(")[0].strip().lower()
        scad_doc = c["doc"].split("SCADUTA IL")[-1].strip(" -")
        return ("DOCUMENTO",
                T(c, f"Nella tua anagrafica risulta la {docbreve} scaduta il {scad_doc}: devo aggiornarla. "
                     "Già che ci vediamo, facciamo 30 minuti di check-up gratuito sulla tua protezione.",
                  f"Nella sua anagrafica risulta la {docbreve} scaduta il {scad_doc}: devo aggiornarla. "
                  "Già che ci vediamo, facciamo 30 minuti di check-up gratuito sulla sua protezione."),
                f"devo aggiornare il documento scaduto in anagrafica e approfittarne per un check-up")
    if c.get("ini_txt"):
        t = c["ini_txt"] if tu else c["ini_txt"].replace("come te", "come lei")
        if True:
            return ("INIZIATIVA ALLIANZ",
                    T(c, f"Ti chiamo perché {t}, e ho pensato subito a te.",
                      f"La chiamo perché {t}, e ho pensato subito a lei."), t)
    if c["scad"] and (c["scad"] - START).days <= 70:
        return ("RINNOVO",
                T(c, f"La prossima quietanza è il {c['scad'].strftime('%d/%m')}: prima di arrivarci vorrei fare con te "
                     "30 minuti di check-up gratuito della protezione.",
                  f"La prossima quietanza è il {c['scad'].strftime('%d/%m')}: prima di arrivarci vorrei fare con lei "
                  "30 minuti di check-up gratuito della protezione."),
                f"prima della quietanza del {c['scad'].strftime('%d/%m')} vorrei fare un check-up")
    return ("CHECK-UP ANNUALE",
            T(c, "Sto facendo con tutti i miei clienti il check-up annuale della protezione: 30 minuti, gratuito, "
                 "e ti lascio una fotografia chiara di cosa sei coperto e cosa no.",
              "Sto facendo con tutti i miei clienti il check-up annuale della protezione: 30 minuti, gratuito, "
              "e le lascio una fotografia chiara di cosa è coperto e cosa no."),
            "sto facendo il check-up annuale gratuito della protezione con tutti i miei clienti")


def luogo(c):
    if c["prov"] in PROV_LONTANE:
        return "Videochiamata (cliente fuori zona)"
    if c["impresa"] or c["prof"] in ("Imprenditore",) or "Edile" in c["prof"] or "Artigiano" in c["prof"]:
        return "In azienda/cantiere del cliente"
    if c["eta"] and c["eta"] >= 68:
        return "A domicilio"
    if c["famiglia"]:
        return "In agenzia o a domicilio, con il partner/familiare"
    return "In agenzia (in alternativa videochiamata)"


def finestra(c):
    p = c["prof"]
    if c["impresa"] or c["autonomo"]:
        return "W1"  # 8:30-9:30 prima del lavoro
    if p in ("Pensionato", "Casalinga", "Benestante"):
        return "W2"  # 9:30-11:30
    if p in ("Studente",):
        return "W4"
    if p in ("Operaio", "Commesso", "Tecnico Con Lavoro Manuale", "Paramedico", "Lavoratore Dipendente Privato"):
        return "W3"  # pausa pranzo
    return "W4"      # impiegati: dopo il lavoro


FIN = OrderedDict([
    ("W1", ("08:30", "09:30", "Imprenditori/autonomi – prima del lavoro")),
    ("W2", ("09:30", "11:30", "Pensionati/casalinghe – mattina")),
    ("W3", ("12:30", "14:00", "Operai/dipendenti – pausa pranzo")),
    ("W4", ("17:30", "19:30", "Impiegati/studenti – dopo il lavoro")),
])
FIN_ALT = {"W1": ["W3", "W4"], "W2": ["W3", "W4"], "W3": ["W4", "W2"], "W4": ["W3", "W2"]}
FIN_CAP = {"W1": 4, "W2": 5, "W3": 5, "W4": 6}

APP_SLOT = {"W1": ("13:30", "in pausa pranzo"), "W2": ("10:30", "in mattinata"),
            "W3": ("18:30", "dopo il lavoro"), "W4": ("18:00", "dopo il lavoro")}

# --------------------------------------------------------------- pianificazione
workdays = []
x = START
while len(workdays) < 36:
    if x.weekday() < 6:
        workdays.append(x)
    x += dt.timedelta(days=1)


def add_wd(day, n, sat=False):
    x, k = day, 0
    while k < n:
        x += dt.timedelta(days=1)
        if x.weekday() < 5 or (sat and x.weekday() == 5):
            k += 1
    return x


# unità di contatto: stesso numero di telefono = una sola chiamata
units = []
seen = set()
for c in sorted(recs, key=lambda c: -c["score"]):
    if c["id"] in seen:
        continue
    grp = [c] + [o for o in recs if o is not c and o["id"] not in seen and o["cell_raw"] == c["cell_raw"]
                 and c["cell_raw"] not in ("", "nan")]
    for g in grp:
        seen.add(g["id"])
    units.append(grp)

for u in units:
    lead = max(u, key=lambda c: c["score"])
    u_score = max(c["score"] for c in u)
    urgent = any(c["disdetta"] or (c["impresa"] and "CATNAT" in c["need"]) or
                 (c["scad"] and (c["scad"] - START).days <= 30) for c in u)
    cess = all(c["cessato"] for c in u)
    rel = START
    sc = min([c["scad"] for c in u if c["scad"]] or [None]) if True else None
    if sc and (sc - START).days <= 30 and not any(c["disdetta"] for c in u):
        rel = max(START, sc - dt.timedelta(days=12))
    for c in u:
        c["unit_lead"] = lead["id"]
    u_meta = dict(u=u, score=u_score, urgent=urgent, cess=cess, release=rel, scad=sc)
    units[units.index(u)] = u_meta

# famiglie (stesso indirizzo): pianificate lo stesso giorno, una dopo l'altra
assigned = {}
day_load = defaultdict(int)
win_load = defaultdict(int)


def cap(day):
    return CAP_SABATO if day.weekday() == 5 else CAP_FERIALE


pool = [u for u in units if not u["cess"]]
cess_pool = [u for u in units if u["cess"]]
urg = sorted([u for u in pool if u["urgent"]], key=lambda u: (u["release"], u["scad"] or dt.date(2099, 1, 1)))
rest = sorted([u for u in pool if not u["urgent"]], key=lambda u: -u["score"])

order_by_day = defaultdict(list)
di = 0
while urg or rest:
    day = workdays[di]
    while day_load[day] < cap(day) and (urg or rest):
        pick = None
        for u in urg:
            if u["release"] <= day:
                pick = u
                break
        if pick:
            urg.remove(pick)
        elif rest:
            pick = rest.pop(0)
        else:
            break
        order_by_day[day].append(pick)
        day_load[day] += 1
        # aggiunge i familiari allo stesso giorno
        for c in pick["u"]:
            for fam in c["famiglia"]:
                for other in list(rest) + list(urg):
                    if any(o["nome_full"] == fam for o in other["u"]):
                        (rest if other in rest else urg).remove(other)
                        order_by_day[day].append(other)
                        day_load[day] += 1
    di += 1
# cessati (win-back) nel giorno successivo all'ultima ondata
wb_day = workdays[di] if di < len(workdays) else workdays[-1]
for u in cess_pool:
    order_by_day[wb_day].append(u)

# orario dentro la giornata
for day, lst in order_by_day.items():
    wl = defaultdict(list)
    sat = day.weekday() == 5
    for u in lst:
        lead = [c for c in u["u"] if c["id"] == u["u"][0]["unit_lead"]][0] if u["u"] else u["u"][0]
        w = finestra(u["u"][0])
        if sat:
            w = "W2"
        cands = [w] + FIN_ALT[w]
        for ww in cands:
            capw = 20 if sat else FIN_CAP[ww]
            if len(wl[ww]) < capw:
                wl[ww].append(u)
                break
        else:
            wl["W4"].append(u)
    for w, ul in wl.items():
        st = dt.datetime.combine(day, dt.time(*map(int, FIN[w][0].split(":"))))
        if sat:
            st = dt.datetime.combine(day, dt.time(9, 0))
        span = (dt.datetime.combine(day, dt.time(*map(int, FIN[w][1].split(":")))) - st).seconds // 60
        step = max(10, min(20, span // max(1, len(ul))))
        for k, u in enumerate(ul):
            t = (st + dt.timedelta(minutes=step * k)).time()
            for c in u["u"]:
                c["day"], c["time"], c["win"] = day, t.strftime("%H:%M"), w

# ----------------------------------------------------- testi personalizzati
for c in recs:
    prods = list(c["prodotti"])
    c["ini_txt"] = ""
    for key, txt, codes in INIT_HOOK:
        hit = [p for p in prods[:3] if p in codes]
        if key in c["ini"] and hit:
            c["ini_txt"] = txt
            if not (prods and prods[0] == "CATNAT"):
                prods.remove(hit[0])
                prods.insert(0, hit[0])
            break
    c["prodotti"] = prods
    prods = prods[:3] or ["CHECK"]
    p1 = prods[0]
    c["p1"], c["p2"] = p1, prods[1] if len(prods) > 1 else ""
    c["p3"] = prods[2] if len(prods) > 2 else ""
    code, hook_long, hook_short = hook(c)
    c["hook_code"] = code
    nome = c["saluto"] or ""
    if c["impresa"] and not c["titolare"]:
        intro_nome = "Buongiorno"
    else:
        sig = ("Sig. " if c["sesso"] == "M" else "Sig.ra ") + c["cognome"]
        intro_nome = T(c, f"Ciao {nome}", f"Buongiorno {sig if (c['eta'] or 0) >= 60 and c['cognome'] else nome}")
    c["intro_nome"] = intro_nome
    w = c.get("win", "W4")
    d1 = add_wd(c["day"], 2)
    sab = c["day"] + dt.timedelta(days=(5 - c["day"].weekday()) % 7 or 7)
    if (sab - c["day"]).days < 2:
        sab += dt.timedelta(days=7)
    t1, quando1 = APP_SLOT[w]
    if c["impresa"] or c["autonomo"]:
        opz1, opz2 = f"{giorno(d1)} alle {t1}", f"{giorno(add_wd(c['day'], 3))} alle 8:30"
    elif c["famiglia"] or (c["eta"] and c["eta"] < 50 and w in ("W3", "W4")):
        opz1, opz2 = f"{giorno(d1)} alle {t1}", f"sabato {sab.strftime('%d/%m')} alle 10:00"
    else:
        opz1, opz2 = f"{giorno(d1)} alle {t1}", f"{giorno(add_wd(c['day'], 3))} alle {'16:00' if w=='W2' else t1}"
    c["opz1"], c["opz2"] = opz1, opz2
    c["luogo"] = luogo(c)
    br = bridge(c, p1) if p1 != "CHECK" else ""
    if code == "OBBLIGO CAT NAT":
        br = bridge(c, c["p2"]) if c["p2"] else ""
    if not c["privacy"]:
        if code == "DOCUMENTO":
            hook_long = hook_long.split(" Già che ci vediamo")[0] + T(c, " Ci vediamo 10 minuti per sistemarla?", " Ci vediamo 10 minuti per sistemarla?")
            hook_short = "devo aggiornare il documento scaduto in anagrafica"
        elif code == "RINNOVO" and c["scad"]:
            dd = c["scad"].strftime("%d/%m")
            hook_long = T(c, f"Ti chiamo per la quietanza del {dd}: organizziamo insieme il rinnovo e aggiorniamo l'anagrafica, bastano 10 minuti.",
                          f"La chiamo per la quietanza del {dd}: organizziamo insieme il rinnovo e aggiorniamo l'anagrafica, bastano 10 minuti.")
            hook_short = f"devo organizzare con {T(c,'te','lei')} la quietanza del {dd}"
        else:
            code = "SERVIZIO"
            hook_long = T(c, "Sto aggiornando le anagrafiche dei miei clienti e mi servono 10 minuti con te per sistemare documenti e consensi.",
                          "Sto aggiornando le anagrafiche dei miei clienti e mi servono 10 minuti con lei per sistemare documenti e consensi.")
            hook_short = "devo aggiornare la tua anagrafica" if not c["formale"] else "devo aggiornare la sua anagrafica"
        c["hook_code"] = code
        br = ""
    s = []
    s.append(f"1) APERTURA: \"{intro_nome}, sono {AGENTE}, {T(c,'il tuo consulente','il suo consulente')} Allianz. "
             f"{T(c,'Hai','Ha')} due minuti?\" (se no: \"Quando {T(c,'ti','la')} richiamo, oggi alle 18 o domani mattina?\")")
    s.append(f"2) MOTIVO: \"{hook_long}\"")
    if not c["privacy"]:
        s.append("3) CONSENSO (prima di tutto): \"In quell'occasione aggiorniamo anche i consensi privacy, così posso tenerla informata "
                 "sulle iniziative utili per lei.\" → SOLO dopo la firma del consenso si passa all'analisi dei bisogni.")
    if br:
        s.append(f"3) PONTE SUL BISOGNO: \"{br}\" → ascolta, non vendere al telefono.")
    if c["privacy"]:
        s.append(f"4) APPUNTAMENTO: \"Facciamo così: {T(c,'ti','le')} preparo un'analisi su misura e ci vediamo 30 minuti "
                 f"({c['luogo'].split('(')[0].strip().lower()}). {T(c,'Ti','Le')} va meglio {opz1} oppure {opz2}?\"")
    else:
        s.append(f"4) APPUNTAMENTO: \"{T(c,'Ti','Le')} va meglio {opz1} oppure {opz2} "
                 f"({c['luogo'].split('(')[0].strip().lower()})?\"")
    fam_p = [tc(x) for x in c["famiglia"] if x in NOMI_PERSONE and x not in c["stesso_tel"]]
    fam_i = [x for x in c["famiglia"] + c["stesso_tel"] if x not in NOMI_PERSONE]
    extra = []
    if fam_p:
        extra.append(f"Se {T(c,'vuoi','vuole')}, facciamo il punto anche per {', '.join(fam_p[:2])}: venite insieme così "
                     "ottimizziamo tutto.")
    if fam_i and not c["impresa"]:
        extra.append(f"E già che ci siamo guardiamo anche le coperture della {T(c,'tua','sua')} attività.")
    if extra:
        s.append("5) FAMIGLIA/ATTIVITÀ: \"" + " ".join(extra) + "\"")
    s.append(f"{'6' if extra else '5'}) CHIUSURA: \"Perfetto, {T(c,'ti','le')} mando subito conferma su WhatsApp "
             f"con cosa portare. A {opz1.split(' alle')[0]}!\"")
    c["script"] = "\n".join(s)

    greet = c["intro_nome"]
    c["wa1"] = (f"{greet}, sono {AGENTE.split()[0]} di Allianz{' 👋' if not c['formale'] else ''}. "
                f"{T(c,'Ti scrivo perché','Le scrivo perché')} {hook_short}. "
                f"Oggi verso le {c['time']} {T(c,'ti','la')} chiamo per 2 minuti: se {T(c,'preferisci','preferisce')} "
                f"un altro orario {T(c,'scrivimelo','me lo scriva')} pure qui.")
    offerta = {
        "TCM": "un check-up gratuito sulla protezione della famiglia",
        "SALUTE": "un check-up gratuito su salute e infortuni",
        "SALUTE+": "una revisione gratuita della tua Ultra Salute" if not c["formale"] else "una revisione gratuita della sua Ultra Salute",
        "LTC": "un check-up gratuito su salute e assistenza futura",
        "PREV": "una simulazione gratuita di pensione e risparmio fiscale",
        "INV": "un check-up gratuito su risparmi e protezione",
        "CASA": "un check-up gratuito su casa e famiglia",
        "AUTO": "un confronto gratuito sull'auto + check-up protezione",
        "INFCIRC": "un check-up gratuito sulle coperture auto e conducente",
        "CATNAT": "la verifica dell'obbligo CAT NAT e dei rischi dell'impresa",
        "IMPRESA": "un check-up gratuito dei rischi dell'impresa",
        "KEYMAN": "un check-up gratuito dei rischi dell'impresa e del titolare",
        "WELFARE": "una presentazione su previdenza e welfare aziendale",
    }.get(p1, "un check-up gratuito della protezione")
    if not c["privacy"]:
        offerta = "un breve incontro per aggiornare documenti e anagrafica"
    c["offerta"] = offerta
    c["wa2"] = (f"{greet}, {T(c,'ho provato a chiamarti','ho provato a chiamarla')} senza fortuna. "
                f"{T(c,'Volevo proporti','Volevo proporle')} {offerta} ({'10' if not c['privacy'] else '30'} minuti). "
                f"{T(c,'Ti va meglio','Le va meglio')} {opz1} o {opz2}? "
                f"{T(c,'Rispondimi anche solo 1 o 2 👍','Mi risponda anche solo 1 o 2, grazie.')}")
    ogg = {
        "PREV": "la tua pensione e le tasse 2026" if not c["formale"] else "pensione e risparmio fiscale 2026",
        "CATNAT": "obbligo polizza catastrofi naturali",
    }.get(p1, "il tuo check-up protezione 2026" if not c["formale"] else "il suo check-up protezione 2026")
    c["email_txt"] = (f"OGGETTO: {greet.replace('Buongiorno ', '').replace('Ciao ', '') + ', ' if greet not in ('Buongiorno',) else ''}{ogg}\n\n"
                  f"{greet},\n{T(c,'ho provato a contattarti','ho provato a contattarla')} nei giorni scorsi. "
                  f"{hook_long}\n\n{T(c,'Ti propongo','Le propongo')} 30 minuti, {opz1} oppure {opz2} "
                  f"({c['luogo'].split('(')[0].strip().lower()}). "
                  f"{T(c,'Se preferisci un altro momento, rispondi pure a questa mail o scrivimi su WhatsApp.','Se preferisce un altro momento, può rispondere a questa mail o scrivermi su WhatsApp.')}\n\n"
                  f"Un saluto,\n{AGENTE} – Allianz")
    dq = []
    for p in [p1, c["p2"]]:
        if p in DISCOVERY:
            dq += DISCOVERY[p][:3 if p == p1 else 2]
    if True:
        dq = []
        src = DISCOVERY_LEI if c["formale"] else DISCOVERY
        for p in [p1, c["p2"]]:
            if p in src:
                dq += src[p][:3 if p == p1 else 2]
        c["discovery"] = "\n".join(f"• {q}" for q in dq)
    else:
        def lei(q):
            rep = [("Hai ", "Ha "), ("hai ", "ha "), ("ti ", "le "), ("Ti ", "Le "), ("tuo ", "suo "), ("tua ", "sua "),
                   ("tuoi ", "suoi "), ("Sei ", "È "), ("sei ", "è "), ("riesci", "riesce"), ("pensi", "pensa"),
                   ("conteresti", "conterebbe"), ("Preferiresti", "Preferirebbe"), ("reagiresti", "reagirebbe"),
                   ("potresti", "potrebbe"), ("avrai", "avrà"), ("potessi", "potesse"), ("Dove va oggi il tuo", "Dove va oggi il suo"),
                   ("da te", "da lei"), ("piacerebbe", "piacerebbe"), ("Fai ", "Fa "), ("fai ", "fa "), ("usato", "usato")]
            for a, b in rep:
                q = q.replace(a, b)
            return q
        c["discovery"] = "\n".join(f"• {lei(q)}" for q in dq)
    ob = OBIEZIONI.get(p1)
    c["obiezione"] = (f"{ob[0]} → {ob[1]}\n" if ob else "") + OBIEZIONI_COMUNI
    prop = [f"1) {PNAME.get(p, p)}: {PROPOSTA.get(p,'')}" for p in [p1] if p in PROPOSTA]
    prop += [f"{k}) {PNAME[p]}" for k, p in ((2, c["p2"]), (3, c["p3"])) if p]
    c["proposta"] = "\n".join(prop) if prop else "Check-up generale: verificare bisogni e consenso privacy."
    docs = ["documento d'identità valido" + (" (QUELLO IN ANAGRAFICA È SCADUTO)" if c["doc"] else "")]
    for p in [p1, c["p2"]]:
        if p in DOCS and DOCS[p] != "nessuno":
            docs.append(DOCS[p])
    c["docs"] = "; ".join(docs)
    if c["otp"]:
        c["firma"] = "OTP attivo → firma digitale anche a distanza (utile se il cliente rimanda)."
    else:
        c["firma"] = "OTP non attivo → firma grafometrica/cartacea in appuntamento; attivare OTP per le prossime."
    comp = []
    if not c["privacy"]:
        comp.append("⚠ CONSENSO PRIVACY COMMERCIALE ASSENTE: contatto solo di servizio (rinnovo/documenti), "
                    "far firmare il consenso PRIMA di proporre prodotti")
    comp.append("IDD: raccolta esigenze e bisogni (demands & needs) + informativa precontrattuale "
                "(Allegati 3, 4/4-bis, 4-ter Reg. IVASS 40/2018) + set informativo/DIP prima della firma")
    if "INV" in (p1, c["p2"], c["p3"]) or "PREV" in (p1, c["p2"]):
        comp.append("Prodotti d'investimento (IBIP): questionario di adeguatezza " +
                    ("già presente, verificare che sia aggiornato" if c["ibips"] else "DA COMPILARE") +
                    " + KID; per il fondo pensione: Nota informativa COVIP e modulo di adesione")
    c["compliance"] = "\n".join(comp)
    alerts = []
    if c["cessato"]:
        alerts.append(f"EX CLIENTE (fine rapporto {c['fine_rapp']}) → win-back")
    if c["disdetta"]:
        alerts.append(f"Disdetta ultimi 12 mesi: {c['prod_cess']}")
    if c["doc"]:
        alerts.append("Documento scaduto")
    if c["scad"] and c["scad"] < START:
        alerts.append(f"Quietanza del {c['scad'].strftime('%d/%m/%Y')} con data passata: verificare se è stata pagata")
    if not c["privacy"]:
        alerts.append("No consenso privacy commerciale")
    if c["stesso_tel"]:
        alerts.append("Stesso numero di: " + ", ".join(tc(x) for x in c["stesso_tel"]))
    if c["famiglia"]:
        alerts.append("Stesso indirizzo di: " + ", ".join(tc(x) for x in c["famiglia"]))
    c["alert"] = " | ".join(alerts)
    if c["formale"]:
        seq = "Chiamata nello slot → se non risponde WhatsApp \"non risponde\" subito → richiamo dopo 2 gg lavorativi"
    else:
        seq = "WhatsApp di preavviso alle 8:00 → chiamata nello slot → se non risponde WhatsApp \"non risponde\" → richiamo dopo 2 gg"
    if c["impresa"]:
        seq = "Chiamata 8:30–9:30 (o 13:00) → se non risponde WhatsApp + email → richiamo dopo 2 gg"
    if not c["privacy"]:
        seq = "Solo chiamata di servizio (rinnovo/documento). Niente messaggi promozionali finché non c'è il consenso."
    c["sequenza"] = seq
    c["canale"] = ("Chiamata" if c["formale"] or c["impresa"] else "WhatsApp + chiamata")

# ===================================================================== EXCEL
wb = Workbook()
F = "Arial"
NAVY = "1F3864"
H_FILL = PatternFill("solid", fgColor=NAVY)
H_FONT = Font(name=F, bold=True, color="FFFFFF", size=10)
B_FONT = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
TITLE = Font(name=F, size=16, bold=True, color=NAVY)
SUB = Font(name=F, size=11, bold=True, color=NAVY)
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
BLUE = Font(name=F, size=10, color="0000FF")
GREEN = Font(name=F, size=10, color="008000")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
CLASS_FILL = {"A": "C6EFCE", "B": "FFEB9C", "C": "F2F2F2"}
STATI = ["1-Da contattare", "2-Non risponde", "3-Da richiamare", "4-Appuntamento fissato", "5-Analisi fatta",
         "6-Proposta presentata", "7-Firmato", "8-Non interessato", "9-Rimandato a prossimo rinnovo"]
ESITI = ["Risposto - appuntamento", "Risposto - richiamare", "Risposto - no", "Non risponde", "WhatsApp inviato",
         "Numero errato"]


def header(ws, row, cols, widths=None, height=30):
    for j, h in enumerate(cols, 1):
        cl = ws.cell(row=row, column=j, value=h)
        cl.font, cl.fill, cl.alignment, cl.border = H_FONT, H_FILL, CENTER, BORDER
    ws.row_dimensions[row].height = height
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w


def put(ws, r, cidx, v, font=B_FONT, align=WRAP, fill=None, fmt_=None, border=True):
    cl = ws.cell(row=r, column=cidx, value=v)
    cl.font, cl.alignment = font, align
    if border:
        cl.border = BORDER
    if fill:
        cl.fill = PatternFill("solid", fgColor=fill) if isinstance(fill, str) else fill
    if fmt_:
        cl.number_format = fmt_
    return cl


# ---------------------------------------------------------------- Schede Clienti
ws_s = wb.active
ws_s.title = "Schede Clienti"
sc_cols = ["ID", "Cliente", "Nome per saluto", "Tipo", "Sesso", "Età", "Professione", "Comune (Prov)",
           "Cellulare", "Email", "Lei / tu", "Privacy marketing", "Firma OTP", "Quest. IBIPs", "Cliente dal",
           "N. polizze", "Premi annui €", "Prossima quietanza", "Portafoglio attuale", "Moduli salute attuali",
           "Gap TCM", "Gap Salute", "Gap LTC", "Gap Previdenza", "Gap Risparmio/Inv.", "Gap Casa", "Gap Auto",
           "Gap Impresa/CAT NAT", "Iniziative Allianz assegnate", "ALERT", "Score", "Classe",
           "Prodotto 1 (focus)", "Perché (dai dati)", "Prodotto 2", "Prodotto 3", "Data 1° contatto", "Ora",
           "Sequenza di contatto", "Tipo di hook", "SCRIPT TELEFONATA", "WhatsApp 1 – preavviso",
           "WhatsApp 2 – se non risponde", "Email (3° tentativo)", "Domande di analisi (appuntamento 1)",
           "Obiezione probabile → risposta", "Proposta da presentare (Good/Better/Best)", "Documenti da portare",
           "Luogo appuntamento", "Appuntamento: 2 opzioni da proporre", "Firma", "Compliance"]
sc_w = [7, 26, 14, 9, 6, 5, 16, 18, 16, 24, 6, 9, 7, 8, 11, 7, 10, 11, 32, 30, 7, 7, 7, 8, 8, 7, 7, 9, 36, 30,
        6, 6, 22, 40, 18, 18, 11, 6, 34, 14, 70, 50, 50, 55, 55, 60, 60, 40, 24, 30, 34, 55]
header(ws_s, 1, sc_cols, sc_w, 36)
ws_s.freeze_panes = "C2"
recs_sched = sorted(recs, key=lambda c: (c["day"], c["time"], c["id"]))
row_of = {}
for r_i, c in enumerate(sorted(recs, key=lambda c: c["id"]), 2):
    row_of[c["id"]] = r_i
    gap = lambda flag: "Sì" if flag else "No"
    imp_gap = c["impresa"] and not c["ente"] and (not c["has_catnat"] or not c["has_uimp"])
    vals = [c["id"], tc(c["nome_full"]) if not c["impresa"] else c["nome_full"],
            c["saluto"], "Impresa/P.IVA" if c["impresa"] else "Persona", c["sesso"], c["eta"], c["prof"],
            f"{c['comune']} ({c['prov']})", c["cell"], c["email"], "Lei" if c["formale"] else "tu",
            "Sì" if c["privacy"] else "NO", "Sì" if c["otp"] else "No", "Sì" if c["ibips"] else "No",
            c["cliente_dal"], c["npol"], c["premi"], c["scad"], c["cop"].strip(" -").replace(" - ", "; ") or "—",
            "; ".join(c["sal_mods"]) or "—",
            (gap(not c["has_tcm"]) if (c["eta"] or 0) <= 70 else "n.a."), gap(not c["has_sal"]), gap(not c["has_ltc"] and (c["eta"] or 0) >= 48),
            (gap(not c["has_prev"]) if (c["eta"] or 0) <= 62 and c["prof"] not in ("Pensionato", "Condomini") else "n.a."),
            gap(not c["has_inv"]), gap(not c["has_casa"]), gap(not c["has_auto"]),
            gap(imp_gap) if c["impresa"] else "—",
            re.sub(r"\s{2,}", " ", c["ini"].strip(" -")).replace(" - ", "\n") or "—", c["alert"] or "—", c["score"], c["classe"],
            PNAME.get(c["p1"], "Check-up generale"), perche(c, c["p1"]), PSHORT.get(c["p2"], ""),
            PSHORT.get(c["p3"], ""), c["day"], c["time"], c["sequenza"], c["hook_code"], c["script"], c["wa1"],
            c["wa2"], c["email_txt"] if c["email"] and c["privacy"] else ("(nessuna email)" if not c["email"] else "(no consenso)"),
            c["discovery"], c["obiezione"], c["proposta"], c["docs"], c["luogo"],
            f"A) {c['opz1']}\nB) {c['opz2']}", c["firma"], c["compliance"]]
    for j, v in enumerate(vals, 1):
        f_ = None
        if j in (15, 18, 37):
            f_ = "dd/mm/yyyy"
        if j == 17:
            f_ = '#,##0.00 "€"'
        put(ws_s, r_i, j, v, fmt_=f_, fill=CLASS_FILL[c["classe"]] if j == 32 else None)
    ws_s.row_dimensions[r_i].height = 210
ws_s.auto_filter.ref = f"A1:{get_column_letter(len(sc_cols))}{len(recs)+1}"
for col in (21, 22, 23, 24, 25, 26, 27, 28):
    L = get_column_letter(col)
    ws_s.conditional_formatting.add(f"{L}2:{L}{len(recs)+1}",
                                    CellIsRule(operator="equal", formula=['"Sì"'], fill=PatternFill("solid", fgColor="F8CBAD")))
L = get_column_letter(12)
ws_s.conditional_formatting.add(f"{L}2:{L}{len(recs)+1}",
                                CellIsRule(operator="equal", formula=['"NO"'], fill=PatternFill("solid", fgColor="FF9999")))

# ---------------------------------------------------------------- Piano & Pipeline
ws_p = wb.create_sheet("Piano & Pipeline", 0)
pp_cols = ["Data", "Giorno", "Ora", "Fascia", "ID", "Cliente", "Cellulare", "Classe", "Score", "Canale",
           "Hook (perché chiamo)", "Prodotto focus", "2° prodotto", "Luogo appuntamento", "Appuntamento da proporre",
           "Scheda →", "Esito tentativo 1", "Richiamo (2° tent.)", "Esito tentativo 2", "Esito tentativo 3",
           "STATO", "Data appuntamento", "Prodotti proposti", "Premio annuo proposto €", "Premio annuo firmato €",
           "Data firma", "Referral raccolti", "Note"]
pp_w = [11, 10, 6, 13, 6, 26, 16, 6, 6, 13, 30, 18, 16, 22, 30, 9, 18, 11, 18, 18, 22, 12, 20, 12, 12, 11, 20, 30]
put(ws_p, 1, 1, "PIANO CHIAMATE & PIPELINE – Full immersion dal 12/10/2026", TITLE, Alignment(vertical="center"), border=False)
put(ws_p, 2, 1, "Ogni mattina filtra la colonna Data sul giorno di oggi. Le celle GIALLE si compilano dopo ogni contatto. "
                "Clicca 'Apri scheda' per lo script completo del cliente.", Font(name=F, size=10, italic=True),
    Alignment(vertical="center"), border=False)
HR = 4
header(ws_p, HR, pp_cols, pp_w, 34)
ws_p.freeze_panes = ws_p.cell(row=HR + 1, column=7)
dv_stato = DataValidation(type="list", formula1=f"=Liste!$A$2:$A${len(STATI)+1}", allow_blank=True)
dv_esito = DataValidation(type="list", formula1=f"=Liste!$B$2:$B${len(ESITI)+1}", allow_blank=True)
ws_p.add_data_validation(dv_stato)
ws_p.add_data_validation(dv_esito)
r = HR + 1
first = r
prev_day = None
for c in recs_sched:
    vals = [c["day"], GIORNI[c["day"].weekday()], c["time"], FIN[c["win"]][2].split(" – ")[0] if c["day"].weekday() < 5 else "Sabato",
            c["id"], tc(c["nome_full"]) if not c["impresa"] else c["nome_full"], c["cell"], c["classe"], c["score"],
            c["canale"], f"{c['hook_code']}", PSHORT.get(c["p1"], "Check-up"), PSHORT.get(c["p2"], ""),
            c["luogo"].split("(")[0].strip(), f"{c['opz1']} / {c['opz2']}", "Apri scheda"]
    for j, v in enumerate(vals, 1):
        f_ = "dd/mm/yyyy" if j == 1 else None
        cl = put(ws_p, r, j, v, fmt_=f_, fill=CLASS_FILL[c["classe"]] if j == 8 else None)
    link = ws_p.cell(row=r, column=16)
    link.hyperlink = f"#'Schede Clienti'!A{row_of[c['id']]}"
    link.font = Font(name=F, size=10, color="0563C1", underline="single")
    put(ws_p, r, 17, None, fill=INPUT_FILL)
    put(ws_p, r, 18, f"=WORKDAY(A{r},2)", fmt_="dd/mm/yyyy")
    for j in range(19, 29):
        put(ws_p, r, j, "1-Da contattare" if j == 21 else None, fill=INPUT_FILL,
            fmt_=("dd/mm/yyyy" if j in (22, 26) else ('#,##0.00 "€"' if j in (24, 25) else None)))
    dv_esito.add(f"Q{r}")
    dv_esito.add(f"S{r}:T{r}")
    dv_stato.add(f"U{r}")
    ws_p.row_dimensions[r].height = 30
    if prev_day and prev_day != c["day"]:
        for j in range(1, 29):
            ws_p.cell(row=r, column=j).border = Border(left=thin, right=thin, bottom=thin,
                                                         top=Side(style="medium", color=NAVY))
    prev_day = c["day"]
    r += 1
last = r - 1
ws_p.auto_filter.ref = f"A{HR}:AB{last}"
ws_p.conditional_formatting.add(f"A{first}:AB{last}", FormulaRule(formula=[f'$U{first}="7-Firmato"'],
                                                                   fill=PatternFill("solid", fgColor="C6EFCE")))
ws_p.conditional_formatting.add(f"A{first}:AB{last}", FormulaRule(formula=[f'$U{first}="4-Appuntamento fissato"'],
                                                                   fill=PatternFill("solid", fgColor="DDEBF7")))
ws_p.conditional_formatting.add(f"A{first}:AB{last}", FormulaRule(formula=[f'$U{first}="8-Non interessato"'],
                                                                   font=Font(color="808080", strike=True)))
PP = "'Piano & Pipeline'"
RNG = lambda col: f"{PP}!${col}${first}:${col}${last}"

# ---------------------------------------------------------------- Liste
ws_l = wb.create_sheet("Liste")
put(ws_l, 1, 1, "Stati pipeline", BOLD)
put(ws_l, 1, 2, "Esiti contatto", BOLD)
for k, s_ in enumerate(STATI, 2):
    put(ws_l, k, 1, s_)
for k, s_ in enumerate(ESITI, 2):
    put(ws_l, k, 2, s_)
ws_l.column_dimensions["A"].width = 32
ws_l.column_dimensions["B"].width = 28
ws_l.sheet_state = "hidden"

# ---------------------------------------------------------------- Dashboard
ws_d = wb.create_sheet("Dashboard", 0)
ws_d.column_dimensions["A"].width = 44
for L_ in "BCDEFGH":
    ws_d.column_dimensions[L_].width = 16
put(ws_d, 1, 1, "DASHBOARD VENDITE – Piano Full Immersion", TITLE, Alignment(vertical="center"), border=False)
put(ws_d, 2, 1, f"Agente: {AGENTE} · Portafoglio analizzato: REPORT.xlsx ({len(recs)} posizioni) · Start: lunedì 12/10/2026",
    Font(name=F, size=10, italic=True), Alignment(vertical="center"), border=False)

put(ws_d, 4, 1, "IPOTESI (modificabili – celle gialle)", SUB, border=False)
assum = [("Tasso di risposta (contatti raggiunti / pianificati)", 0.65,
          "Ipotesi Claude per clienti già in portafoglio con numero salvato. Aggiorna dopo la 1ª settimana con il dato reale."),
         ("Tasso di appuntamento (appuntamenti / raggiunti)", 0.45,
          "Ipotesi Claude: clienti caldi + check-up gratuito. Aggiorna con i dati reali."),
         ("Tasso di chiusura (firme / appuntamenti)", 0.40,
          "Ipotesi Claude per vendite consulenziali su clienti esistenti."),
         ("Premio annuo medio nuova polizza persona (€)", 450,
          "Ipotesi Claude: media prudente tra TCM, salute, previdenza (contributo annuo). Sostituisci con il tuo dato.")]
for k, (lab, v, note) in enumerate(assum, 5):
    put(ws_d, k, 1, lab)
    cl = put(ws_d, k, 2, v, font=BLUE, fill=INPUT_FILL, fmt_=("0%" if v < 1 else '#,##0 "€"'))
    cl.comment = Comment(note, "Claude")
put(ws_d, 10, 1, "PREVISIONE (dalle ipotesi)", SUB, border=False)
fc = [("Contatti pianificati", f"=COUNTA({RNG('E')})", "0"),
      ("Contatti raggiunti attesi", "=B11*B5", "0"),
      ("Appuntamenti attesi", "=B12*B6", "0"),
      ("Contratti attesi", "=B13*B7", "0"),
      ("Nuovo premio annuo atteso €", "=B14*B8", '#,##0 "€"')]
for k, (lab, f_, nf) in enumerate(fc, 11):
    put(ws_d, k, 1, lab)
    put(ws_d, k, 2, f_, fmt_=nf)

put(ws_d, 17, 1, "PIPELINE REALE (si aggiorna da 'Piano & Pipeline')", SUB, border=False)
header_cells = ["Stato", "N. clienti", "% sul totale"]
for j, h in enumerate(header_cells, 1):
    cl = ws_d.cell(row=18, column=j, value=h)
    cl.font, cl.fill, cl.alignment, cl.border = H_FONT, H_FILL, CENTER, BORDER
for k, s_ in enumerate(STATI, 19):
    put(ws_d, k, 1, s_)
    put(ws_d, k, 2, f'=COUNTIF({RNG("U")},A{k})', fmt_="0")
    put(ws_d, k, 3, f"=IFERROR(B{k}/$B$11,0)", fmt_="0.0%")
kp = 19 + len(STATI) + 1
put(ws_d, kp, 1, "KPI CHIAVE", SUB, border=False)
kpis = [("Appuntamenti fissati (stati 4-7)", f"=SUM(B22:B25)", "0"),
        ("Contratti firmati", "=B25", "0"),
        ("Tasso di chiusura reale (firme / appuntamenti)", f"=IFERROR(B{kp+2}/B{kp+1},0)", "0.0%"),
        ("Premio annuo proposto €", f"=SUM({RNG('X')})", '#,##0 "€"'),
        ("Premio annuo firmato €", f"=SUM({RNG('Y')})", '#,##0 "€"'),
        ("Referral raccolti (righe compilate)", f"=COUNTA({RNG('AA')})", "0"),
        ("Avanzamento vs previsione firme", f"=IFERROR(B{kp+2}/B14,0)", "0.0%")]
for k, (lab, f_, nf) in enumerate(kpis, kp + 1):
    put(ws_d, k, 1, lab, BOLD)
    put(ws_d, k, 2, f_, fmt_=nf, font=BOLD)

ga = kp + len(kpis) + 2
put(ws_d, ga, 1, "ANALISI PORTAFOGLIO – dove sono le opportunità", SUB, border=False)
for j, h in enumerate(["Area di bisogno", "Clienti SCOPERTI", "% portafoglio", "di cui classe A"], 1):
    cl = ws_d.cell(row=ga + 1, column=j, value=h)
    cl.font, cl.fill, cl.alignment, cl.border = H_FONT, H_FILL, CENTER, BORDER
SC = "'Schede Clienti'"
n_s = len(recs) + 1
gaps = [("TCM / protezione famiglia", "U"), ("Salute / infortuni", "V"), ("Non autosufficienza (48+ anni)", "W"),
        ("Previdenza complementare", "X"), ("Risparmio / investimento", "Y"), ("Casa", "Z"), ("Auto", "AA"),
        ("Impresa / CAT NAT (solo P.IVA)", "AB")]
for k, (lab, col) in enumerate(gaps, ga + 2):
    put(ws_d, k, 1, lab)
    put(ws_d, k, 2, f'=COUNTIF({SC}!${col}$2:${col}${n_s},"Sì")', fmt_="0")
    put(ws_d, k, 3, f"=IFERROR(B{k}/COUNTA({SC}!$A$2:$A${n_s}),0)", fmt_="0.0%")
    put(ws_d, k, 4, f'=COUNTIFS({SC}!${col}$2:${col}${n_s},"Sì",{SC}!$AF$2:$AF${n_s},"A")', fmt_="0")
pf = ga + 2 + len(gaps) + 1
put(ws_d, pf, 1, "PRODOTTO FOCUS ASSEGNATO (n. clienti)", SUB, border=False)
for j, h in enumerate(["Prodotto 1", "N. clienti", "Classe A", "Classe B"], 1):
    cl = ws_d.cell(row=pf + 1, column=j, value=h)
    cl.font, cl.fill, cl.alignment, cl.border = H_FONT, H_FILL, CENTER, BORDER
plist = [p for p in PNAME.values()] + ["Check-up generale"]
for k, p in enumerate(plist, pf + 2):
    put(ws_d, k, 1, p)
    put(ws_d, k, 2, f'=COUNTIF({SC}!$AG$2:$AG${n_s},A{k})', fmt_="0")
    put(ws_d, k, 3, f'=COUNTIFS({SC}!$AG$2:$AG${n_s},A{k},{SC}!$AF$2:$AF${n_s},"A")', fmt_="0")
    put(ws_d, k, 4, f'=COUNTIFS({SC}!$AG$2:$AG${n_s},A{k},{SC}!$AF$2:$AF${n_s},"B")', fmt_="0")

# ---------------------------------------------------------------- Calendario
ws_c = wb.create_sheet("Calendario 6 settimane", 1)
cal_cols = ["Data", "Giorno", "Sett.", "Fase", "Nuovi contatti pianificati", "di cui classe A",
            "Recall previsti (2° tent.)", "Target appuntamenti svolti", "Focus del giorno", "Fatto ✓"]
header(ws_c, 3, cal_cols, [11, 10, 6, 26, 12, 10, 12, 12, 90, 8], 40)
put(ws_c, 1, 1, "CALENDARIO GIORNO PER GIORNO (lun–sab mattina)", TITLE, Alignment(vertical="center"), border=False)
FASI = {1: "Ondata 1 – Urgenze + Classe A", 2: "Ondata 2 – Classe A/B", 3: "Ondata 3 – Classe B/C + recall",
        4: "Ondata 4 – Chiusura giro + win-back", 5: "Chiusure & secondi appuntamenti", 6: "Firme, referral, fine anno"}
FOCUS = {
    1: ["AVVIO: stampa/filtra il piano di oggi, prepara le 5 schede classe A. Prime chiamate su rinnovi d'ottobre e disdette. "
        "Obiettivo: 5 appuntamenti fissati entro sera.",
        "Chiamate + primi appuntamenti dalla mattina di mercoledì. Rinnovi di fine ottobre: proponi check-up al momento del pagamento.",
        "Primi appuntamenti di analisi (Fase 3). Usa la 'fotografia protezione' (Processo Vendita).",
        "Chiamate + 3 appuntamenti. Prepara le proposte (Good/Better/Best) per gli appuntamenti di lunedì–martedì.",
        "Chiamate + 3 appuntamenti. 17:00 revisione settimana: aggiorna ipotesi nella Dashboard.",
        "SABATO: 9:00–10:00 recall WhatsApp ai 'non risponde' della settimana; 10:00–12:30 appuntamenti coppie/famiglie."],
    2: ["Nuova ondata. Prima ora: recall dei 'non risponde' di giovedì–venerdì. Presentazione delle prime proposte.",
        "Chiamate + 3–4 appuntamenti. Proposte presentate: fissa SEMPRE la data della decisione prima di uscire.",
        "Chiamate + appuntamenti. Imprese: visite in azienda 13:30 (CAT NAT + persona chiave).",
        "Chiamate + appuntamenti. Follow-up a 48h delle proposte di lunedì (WhatsApp 'Riepilogo').",
        "Chiamate + appuntamenti. Revisione settimana: tasso risposta/appuntamento reale vs ipotesi.",
        "SABATO: appuntamenti famiglie + firme a distanza (OTP) delle proposte in sospeso."],
    3: ["Ondata 3. Inizia dal 3° tentativo per i 'non risponde' (email + WhatsApp).",
        "Chiamate + appuntamenti. Previdenza: leva fiscale – versare entro il 31/12 per dedurre nel 2026.",
        "Chiamate + appuntamenti + firme.",
        "Chiamate + appuntamenti + firme. Dopo ogni firma: richiesta referral (2 nomi).",
        "Chiamate + appuntamenti. Revisione settimana.",
        "SABATO (31/10): appuntamenti famiglie; recall WhatsApp."],
    4: ["Ultime chiamate del giro + WIN-BACK ex clienti. Rinnovi di metà novembre.",
        "Appuntamenti + firme. Clienti 'rimandati': fissa promemoria al prossimo rinnovo.",
        "Appuntamenti + firme. Recall finale: 3° tentativo a chi non ha mai risposto.",
        "Appuntamenti + firme + referral.",
        "Revisione: chi è in stato 6 (proposta) da più di 7 giorni → chiamata di decisione.",
        "SABATO: appuntamenti famiglie + referral."],
    5: ["Secondi appuntamenti (proposte) e chiamate ai referral raccolti.",
        "Firme + post-vendita (WhatsApp di benvenuto, consegna documenti).",
        "Appuntamenti con i referral (stesso processo dalla Fase 1).",
        "Firme + upsell ai clienti chiusi (2° prodotto suggerito nella scheda).",
        "Revisione: aggiorna la Dashboard, analizza obiezioni più frequenti.",
        "SABATO: appuntamenti residui."],
    6: ["Campagna previdenza fine anno (deduzione 2026) a tutti i lavoratori in stato 6 o 9.",
        "Firme + post-vendita.",
        "Preparazione ondata rinnovi di DICEMBRE (47 quietanze): pianifica le chiamate 10–14 giorni prima.",
        "Firme + referral.",
        "Chiusura progetto: report risultati, lezioni apprese, piano di gennaio.",
        "SABATO: ultimi appuntamenti e firme."],
}
day_counts_row = {}
for k, day in enumerate(workdays, 4):
    week = (day - START).days // 7 + 1
    put(ws_c, k, 1, day, fmt_="dd/mm/yyyy")
    put(ws_c, k, 2, GIORNI[day.weekday()])
    put(ws_c, k, 3, week, align=CENTER)
    put(ws_c, k, 4, FASI[week])
    put(ws_c, k, 5, f"=COUNTIF({RNG('A')},A{k})", align=CENTER, fmt_="0")
    put(ws_c, k, 6, f'=COUNTIFS({RNG("A")},A{k},{RNG("H")},"A")', align=CENTER, fmt_="0")
    put(ws_c, k, 7, f"=COUNTIF({RNG('R')},A{k})", align=CENTER, fmt_="0")
    tgt = 0 if k == 4 else (3 if day.weekday() < 5 else 3)
    if week >= 2:
        tgt = 4 if day.weekday() < 5 else 3
    put(ws_c, k, 8, tgt, font=BLUE, fill=INPUT_FILL, align=CENTER)
    put(ws_c, k, 9, FOCUS[week][day.weekday()])
    put(ws_c, k, 10, None, fill=INPUT_FILL)
    ws_c.row_dimensions[k].height = 42
    if day.weekday() == 5:
        for j in range(1, 11):
            if j not in (8, 10):
                ws_c.cell(row=k, column=j).fill = PatternFill("solid", fgColor="EDEDED")
endc = 3 + len(workdays)
put(ws_c, endc + 1, 4, "TOTALE", BOLD)
put(ws_c, endc + 1, 5, f"=SUM(E4:E{endc})", BOLD, CENTER, fmt_="0")
put(ws_c, endc + 1, 6, f"=SUM(F4:F{endc})", BOLD, CENTER, fmt_="0")
put(ws_c, endc + 1, 7, f"=SUM(G4:G{endc})", BOLD, CENTER, fmt_="0")
put(ws_c, endc + 1, 8, f"=SUM(H4:H{endc})", BOLD, CENTER, fmt_="0")
put(ws_c, 2, 1, "Colonne E–G: formule dal foglio Piano & Pipeline. Colonna H (gialla): target modificabile.",
    Font(name=F, size=10, italic=True), Alignment(vertical="center"), border=False)
ws_c.freeze_panes = "A4"

# ---------------------------------------------------------------- Giornata tipo
ws_g = wb.create_sheet("Giornata Tipo", 2)
put(ws_g, 1, 1, "GIORNATA TIPO – FULL IMMERSION", TITLE, Alignment(vertical="center"), border=False)
header(ws_g, 3, ["Orario", "Lun–Ven", "Perché in questa fascia", "Sabato"], [14, 60, 55, 50])
gt = [
    ("07:45–08:15", "Preparazione: filtra 'Piano & Pipeline' sul giorno, rileggi le schede classe A, prepara le proposte del pomeriggio.",
     "Si arriva alla chiamata sapendo già cosa ha e cosa manca il cliente.", "08:45–09:00 preparazione"),
    ("08:00", "Invia i WhatsApp di PREAVVISO (colonna 'WhatsApp 1') ai clienti 'tu' del giorno.",
     "Il preavviso alza il tasso di risposta: il cliente sa chi chiama e perché.", ""),
    ("08:30–09:30", "CHIAMATE fascia W1: imprenditori, artigiani, P.IVA.",
     "Prima di entrare in cantiere/negozio rispondono.", "09:00–10:00 recall WhatsApp a tutti i 'non risponde' della settimana"),
    ("09:30–11:30", "CHIAMATE fascia W2: pensionati, casalinghe + richiami (2° tentativo) del giorno.",
     "Fascia più tranquilla per chi è a casa.", "10:00–12:30 APPUNTAMENTI (coppie e famiglie, 3 slot: 10:00, 11:00, 12:00)"),
    ("10:30 / 11:30", "APPUNTAMENTO A (pensionati, domicilio/agenzia) – dalla 2ª settimana.", "", ""),
    ("12:30–14:00", "CHIAMATE fascia W3: operai e dipendenti in pausa pranzo.", "Unico momento libero per chi lavora su turni.", ""),
    ("13:30", "APPUNTAMENTO B (imprese, in azienda).", "Pausa pranzo dell'imprenditore.", ""),
    ("14:30–16:00", "PREVENTIVI e PROPOSTE: preparazione Good/Better/Best, firme a distanza (OTP), invio riepiloghi.", "", ""),
    ("16:00", "APPUNTAMENTO C (agenzia).", "", ""),
    ("17:30–19:30", "CHIAMATE fascia W4: impiegati, studenti. In parallelo APPUNTAMENTO D alle 18:00/18:30.",
     "Dopo il lavoro: tasso di risposta più alto per i dipendenti.", ""),
    ("19:30–19:45", "AGGIORNA la pipeline (esiti, stato, data appuntamento) e invia i WhatsApp di conferma per domani.",
     "Senza aggiornamento la Dashboard non serve.", "12:30 aggiornamento pipeline settimanale"),
]
for k, row_ in enumerate(gt, 4):
    for j, v in enumerate(row_, 1):
        put(ws_g, k, j, v, font=BOLD if j == 1 else B_FONT)
    ws_g.row_dimensions[k].height = 36
put(ws_g, 16, 1, "Regole d'oro", SUB, border=False)
regole = ["Massimo 3 tentativi per cliente: chiamata → chiamata in fascia diversa dopo 2 gg lavorativi → WhatsApp + email. Poi stato 9 (prossimo rinnovo).",
          "Al telefono si vende l'APPUNTAMENTO, non la polizza. Mai preventivi al telefono.",
          "Proponi sempre 2 opzioni di data (tecnica dell'alternativa): sono già scritte nella scheda.",
          "Ogni appuntamento finisce con una data: decisione, firma o secondo incontro.",
          "Dopo ogni firma: chiedi 2 nomi (referral) e segna in colonna AA.",
          "Gli appuntamenti si confermano su WhatsApp subito e si ricordano il giorno prima (template nel foglio Messaggi)."]
for k, t in enumerate(regole, 17):
    put(ws_g, k, 1, f"{k-16}.", BOLD, border=False)
    ws_g.merge_cells(start_row=k, start_column=2, end_row=k, end_column=4)
    put(ws_g, k, 2, t, border=False)
    ws_g.row_dimensions[k].height = 30

# ---------------------------------------------------------------- Processo vendita
ws_v = wb.create_sheet("Processo Vendita", 3)
put(ws_v, 1, 1, "PROCESSO DI VENDITA – dalla hook alla firma", TITLE, Alignment(vertical="center"), border=False)
header(ws_v, 3, ["Fase", "Quando", "Obiettivo", "Cosa fare (passi)", "Strumenti nel file", "KPI / uscita"],
       [22, 18, 30, 80, 30, 26])
fasi = [
    ("0. Preparazione", "Sera prima / 7:45",
     "Sapere prima di chiamare cosa ha e cosa manca il cliente.",
     "• Filtra il giorno in 'Piano & Pipeline'.\n• Apri la scheda: portafoglio, gap, prodotto 1, alert (privacy, documento, disdetta).\n"
     "• Clienti 'tu' (under 40): prepara il WhatsApp di preavviso.",
     "Schede Clienti (colonne S–AH)", "100% delle schede del giorno lette"),
    ("1. Hook", "Mattina (preavviso) + fascia oraria",
     "Ottenere 2 minuti di attenzione con un motivo reale e personale.",
     "• Motivi in ordine di forza: disdetta → obbligo CAT NAT → quietanza in scadenza → documento scaduto → iniziativa Allianz → check-up annuale.\n"
     "• Reciprocità: offri qualcosa di utile (check-up gratuito, fotografia della protezione).\n"
     "• Chiedi il permesso: 'Ha due minuti?'",
     "Colonne 'Tipo di hook', 'WhatsApp 1'", "Tasso di risposta ≥ 60%"),
    ("2. Fissare l'appuntamento", "Stessa telefonata",
     "Vendere solo l'appuntamento (30 min), mai la polizza al telefono.",
     "• Ponte sul bisogno: una domanda sul gap principale, poi ascolta.\n• Alternativa: 2 date già pronte nella scheda.\n"
     "• Coppie/famiglie: invita entrambi (decidono insieme, niente 'ne parlo con mia moglie').\n"
     "• Conferma subito su WhatsApp con cosa portare.",
     "SCRIPT TELEFONATA, 'Appuntamento: 2 opzioni'", "≥ 45% dei raggiunti → appuntamento"),
    ("3. Appuntamento 1 – Analisi (check-up)", "Entro 2–5 gg dalla chiamata",
     "Far emergere il bisogno in modo che sia il cliente a dirlo.",
     "• 5 min: rompere il ghiaccio, ricordare lo scopo e i tempi.\n• 15 min: domande di analisi (nella scheda) + questionario esigenze e bisogni IDD.\n"
     "• 5 min: 'fotografia della protezione' – foglio con 7 aree (vita, invalidità, salute, LTC, pensione, casa, responsabilità civile): verde = coperto, rosso = scoperto.\n"
     "• Quantifica il gap (es. TCM con metodo DIME: Debiti + Income/reddito × anni + Mutuo + Education/studi figli).\n"
     "• Se è semplice (salute, infortuni): proposta nello stesso incontro. Se è complesso (TCM alta, previdenza, investimento): fissa il 2° incontro entro 5 gg.",
     "Domande di analisi, Documenti da portare", "Gap quantificato + data della proposta"),
    ("4. Proposta (Good/Better/Best)", "Stesso incontro o entro 5 gg",
     "Rendere facile dire sì.",
     "• Presenta 3 opzioni partendo dalla più completa (ancoraggio), consiglia quella di mezzo.\n• Parla in €/mese, collegando ogni garanzia a una frase detta dal cliente.\n"
     "• Mostra cosa resta scoperto con ogni opzione (trasparenza = fiducia).\n• Consegna l'informativa precontrattuale (set informativo, DIP, allegati IVASS).",
     "Proposta da presentare", "Proposta presentata → data della decisione"),
    ("5. Obiezioni", "Durante la proposta",
     "Trasformare il dubbio in informazione.",
     "Metodo: 1) Accogli ('capisco') 2) Chiarisci ('cosa la preoccupa esattamente?') 3) Rispondi con i fatti 4) Ricollega all'obiettivo dichiarato.\n"
     "Mai insistere: se il prodotto non è adatto, dillo (vale per l'IDD e per la reputazione).",
     "Foglio 'Script & Obiezioni'", "≥ 40% di chiusura sugli appuntamenti"),
    ("6. Chiusura e firma", "In appuntamento o a distanza (OTP)",
     "Firmare in modo conforme e senza attriti.",
     "• Domanda di chiusura: 'Quale delle tre configurazioni sente più sua?'\n• Verifica adeguatezza, questionario sanitario (salute/TCM), questionario IBIPs (investimento).\n"
     "• Firma grafometrica in agenzia oppure OTP a distanza.\n• Imposta il pagamento (SDD/carta) e la decorrenza.",
     "Colonne 'Firma', 'Compliance'", "Contratto firmato"),
    ("7. Post-vendita e referral", "Entro 24h + 30 gg",
     "Consolidare il cliente e generare nuovi contatti.",
     "• Entro 24h: WhatsApp di ringraziamento con riepilogo.\n• Richiesta referral: 'Chi tra le persone a cui tiene potrebbe avere bisogno dello stesso check-up?' (2 nomi). "
     "Niente premi o regali in cambio dei nominativi: verifica le regole IVASS e dell'agenzia.\n"
     "• A 30 gg: chiamata di cortesia + 2° prodotto suggerito nella scheda.\n• Annota nel CRM il prossimo check-up (12 mesi).",
     "Messaggi WA 'Post-firma', colonna Referral", "≥ 1 referral ogni 2 firme"),
]
for k, row_ in enumerate(fasi, 4):
    for j, v in enumerate(row_, 1):
        put(ws_v, k, j, v, font=BOLD if j == 1 else B_FONT)
    ws_v.row_dimensions[k].height = 150 if k in (6, 7, 8) else 100
put(ws_v, 13, 1, "Leve psicologiche (etiche) usate negli script", SUB, border=False)
leve = [("Reciprocità", "Check-up e 'fotografia' gratuiti prima di chiedere qualcosa."),
        ("Impegno e coerenza", "Piccoli sì progressivi: 'ha 2 minuti?' → appuntamento → analisi → proposta."),
        ("Riprova sociale", "'Molti clienti della sua età mi chiedono…' (solo se è vero)."),
        ("Ancoraggio + opzione centrale", "Good/Better/Best: si mostra prima la più completa, si consiglia la centrale."),
        ("Scadenze reali", "Quietanza in scadenza, iniziative con data, deduzione previdenza entro il 31/12, obbligo CAT NAT. MAI scadenze inventate."),
        ("Avversione alla perdita (con misura)", "Mostrare cosa si perde senza copertura (es. tasse non dedotte), senza fare leva sulla paura.")]
for k, (a, b) in enumerate(leve, 14):
    put(ws_v, k, 1, a, BOLD)
    ws_v.merge_cells(start_row=k, start_column=2, end_row=k, end_column=4)
    put(ws_v, k, 2, b)

# ---------------------------------------------------------------- Script & Obiezioni
ws_o = wb.create_sheet("Script & Obiezioni", 4)
put(ws_o, 1, 1, "LIBRERIA SCRIPT & OBIEZIONI PER PRODOTTO", TITLE, Alignment(vertical="center"), border=False)
header(ws_o, 3, ["Prodotto", "Frase-ponte (telefono)", "Domande di analisi", "Obiezione tipica", "Risposta",
                 "Proposta Good/Better/Best", "Documenti"], [24, 55, 55, 30, 60, 55, 35])
demo = dict(formale=False, has_auto=True, has_casa=False, has_sal=False, has_tcm=False, has_uimp=False,
            has_prev=False, has_inv=False, sal_mods=["invalidità permanente da infortunio"],
            sal_miss=["invalidità permanente da malattia"], prof="Impiegato", eta=40)
for k, p in enumerate(PNAME, 4):
    ob = OBIEZIONI.get(p, ("", ""))
    vals = [PNAME[p], bridge(demo, p), "\n".join("• " + q for q in DISCOVERY.get(p, [])), ob[0], ob[1],
            PROPOSTA.get(p, ""), DOCS.get(p, "")]
    for j, v in enumerate(vals, 1):
        put(ws_o, k, j, v, font=BOLD if j == 1 else B_FONT)
    ws_o.row_dimensions[k].height = 120
k = 4 + len(PNAME) + 1
put(ws_o, k, 1, "Obiezioni trasversali", SUB, border=False)
trasv = [("\"Non ho tempo\"", "Proprio per questo le propongo 30 minuti, dove preferisce lei, anche in videochiamata. Meglio inizio o fine settimana?"),
         ("\"Mandami una mail / un preventivo\"", "Volentieri, ma senza conoscere la sua situazione le manderei un prezzo generico. In 30 minuti le porto una proposta su misura: martedì o giovedì?"),
         ("\"Ne parlo con mio marito/mia moglie\"", "Giustissimo: facciamo l'appuntamento insieme, così fate le domande entrambi. Il sabato mattina va bene?"),
         ("\"Pago già troppe assicurazioni\"", "Allora il check-up serve ancora di più: controlliamo che non ci siano doppioni e che i soldi vadano dove servono."),
         ("\"Non mi serve, sto bene\"", "È proprio il momento migliore: oggi non ci sono esclusioni e i premi sono più bassi. Guardiamo solo i numeri, poi decide lei."),
         ("\"Ci devo pensare\"", "Certo. Cosa le manca per decidere? (ascolta) Fissiamo una telefonata di 5 minuti venerdì così le rispondo su quel punto."),
         ("\"Non mi fido delle assicurazioni / non pagano\"", "Capisco, succede quando le condizioni non sono chiare. Per questo le mostro esattamente cosa è coperto e cosa no, per iscritto.")]
header(ws_o, k + 1, ["Obiezione", "Risposta"], None, 24)
for i_, (a, b) in enumerate(trasv, k + 2):
    put(ws_o, i_, 1, a, BOLD)
    put(ws_o, i_, 2, b)
    ws_o.row_dimensions[i_].height = 40

# ---------------------------------------------------------------- Messaggi
ws_m = wb.create_sheet("Messaggi WA-Email", 5)
put(ws_m, 1, 1, "TEMPLATE MESSAGGI (WhatsApp ed email) – sostituisci i campi tra {parentesi}", TITLE,
    Alignment(vertical="center"), border=False)
header(ws_m, 3, ["Momento", "Versione TU", "Versione LEI", "Note"], [26, 70, 70, 40])
msgs = [
    ("Preavviso (mattina)", "Personalizzato in 'Schede Clienti' → colonna WhatsApp 1", "Personalizzato in 'Schede Clienti' → colonna WhatsApp 1", "Invia alle 8:00 ai clienti del giorno."),
    ("Non risponde", "Personalizzato in 'Schede Clienti' → colonna WhatsApp 2", "Personalizzato in 'Schede Clienti' → colonna WhatsApp 2", "Subito dopo la chiamata senza risposta."),
    ("Conferma appuntamento",
     "Perfetto {Nome}! Ci vediamo {giorno} alle {ora} {luogo}. Porta se puoi: documento d'identità e {documenti}. A presto! – Dan",
     "Grazie {Nome}, confermo l'appuntamento di {giorno} alle {ora} {luogo}. Se possibile porti: documento d'identità e {documenti}. Cordiali saluti, Dan Moscaliuc – Allianz",
     "Subito dopo la telefonata."),
    ("Promemoria (giorno prima)",
     "Ciao {Nome}, ti ricordo il nostro appuntamento di domani alle {ora}. Se hai un imprevisto scrivimi pure 👍",
     "Buongiorno {Nome}, le ricordo l'appuntamento di domani alle {ora}. Per qualsiasi imprevisto mi scriva pure qui. Grazie!",
     "Riduce i 'buchi' in agenda."),
    ("Riepilogo dopo l'analisi",
     "Grazie per oggi {Nome}! Come promesso ti riassumo: oggi sei coperto su {coperto}, restano scoperti {scoperto}. Ci rivediamo {data} con le 3 proposte.",
     "Grazie per il tempo dedicato oggi. Riepilogo: oggi è coperto su {coperto}, restano scoperti {scoperto}. Ci rivediamo {data} con le 3 proposte.",
     "Entro sera. Fissa per iscritto il prossimo passo."),
    ("Follow-up proposta (48h)",
     "Ciao {Nome}, hai avuto modo di guardare le proposte? Se hai dubbi ne parliamo 5 minuti al telefono, quando ti va?",
     "Buongiorno {Nome}, ha avuto modo di valutare le proposte? Se ha domande possiamo sentirci 5 minuti: quando le è comodo?",
     "Mai più di 2 follow-up senza una risposta."),
    ("Firma a distanza (OTP)",
     "Ciao {Nome}, ti ho appena inviato i documenti da firmare: riceverai un SMS con il codice OTP. Se vuoi ti guido al telefono, ci vogliono 3 minuti.",
     "Buongiorno {Nome}, le ho inviato i documenti per la firma: riceverà un SMS con il codice OTP. Se preferisce la guido al telefono, servono 3 minuti.",
     "Solo per clienti con OTP = Sì."),
    ("Post-firma + referral",
     "Benvenuto {Nome}! Da oggi sei protetto su {prodotto}. Grazie per la fiducia 🙏 Se pensi a qualcuno che potrebbe avere bisogno dello stesso check-up, presentamelo: me ne occuperò con la stessa cura.",
     "Grazie {Nome}, da oggi è protetto su {prodotto}. Se conosce qualcuno a cui potrebbe servire lo stesso check-up, mi farà piacere conoscerlo: lo seguirò con la stessa cura.",
     "Niente premi in cambio dei nominativi (verificare regole IVASS/agenzia)."),
    ("Email 3° tentativo", "Personalizzata in 'Schede Clienti' → colonna Email", "Personalizzata in 'Schede Clienti' → colonna Email", "Solo a chi ha dato il consenso email/privacy."),
    ("Previdenza fine anno (novembre)",
     "Ciao {Nome}, promemoria utile: i versamenti nel fondo pensione fatti entro il 31/12 si deducono dal reddito 2026 (fino a 5.164,57 €). Vuoi che ti faccia la simulazione di quanto risparmi?",
     "Buongiorno {Nome}, un promemoria: i versamenti nel fondo pensione fatti entro il 31/12 si deducono dal reddito 2026 (fino a 5.164,57 €). Desidera una simulazione del risparmio fiscale?",
     "Campagna di novembre per i lavoratori in stato 6/9."),
]
for k, row_ in enumerate(msgs, 4):
    for j, v in enumerate(row_, 1):
        put(ws_m, k, j, v, font=BOLD if j == 1 else B_FONT)
    ws_m.row_dimensions[k].height = 70

# ---------------------------------------------------------------- Agenda appuntamenti
ws_a = wb.create_sheet("Agenda Appuntamenti", 6)
put(ws_a, 1, 1, "AGENDA APPUNTAMENTI – compila quando fissi (celle gialle)", TITLE, Alignment(vertical="center"), border=False)
header(ws_a, 3, ["Data", "Giorno", "10:30 (A)", "13:30 (B)", "16:00 (C)", "18:00/18:30 (D)", "Sabato 10:00", "Sabato 11:00", "Sabato 12:00"],
       [11, 10, 24, 24, 24, 24, 24, 24, 24])
for k, day in enumerate(workdays, 4):
    put(ws_a, k, 1, day, fmt_="dd/mm/yyyy")
    put(ws_a, k, 2, GIORNI[day.weekday()])
    for j in range(3, 10):
        sat = day.weekday() == 5
        active = (j >= 7) if sat else (j <= 6)
        put(ws_a, k, j, None, fill=INPUT_FILL if active else "D9D9D9")
    ws_a.row_dimensions[k].height = 22
ws_a.freeze_panes = "C4"

# ---------------------------------------------------------------- Famiglie & Imprese
ws_f = wb.create_sheet("Famiglie & Imprese", 7)
put(ws_f, 1, 1, "NUCLEI FAMILIARI (stesso indirizzo) e IMPRESE – appuntamenti congiunti", TITLE,
    Alignment(vertical="center"), border=False)
header(ws_f, 3, ["Indirizzo", "Persone/posizioni", "Prodotti focus", "Giorno contatto", "Suggerimento"], [38, 50, 45, 14, 60])
k = 4
for key, lst in by_addr.items():
    names = sorted({x["nome_full"] for x in lst})
    if len(names) < 2 and not any(x["impresa"] for x in lst):
        continue
    if len(names) < 2:
        continue
    imp = any(x["impresa"] for x in lst)
    sugg = ("Titolare e impresa: unico appuntamento in azienda (rischi impresa + persona chiave + famiglia)."
            if imp else "Appuntamento di coppia/famiglia (sabato mattina o sera): check-up per entrambi, sconto multi-polizza se previsto.")
    put(ws_f, k, 1, lst[0]["ind"])
    put(ws_f, k, 2, "\n".join(f"{x['id']} – {tc(x['nome_full']) if not x['impresa'] else x['nome_full']} ({x['eta'] or 'impresa'})" for x in lst))
    put(ws_f, k, 3, "\n".join(f"{x['id']}: {PSHORT.get(x['p1'],'Check-up')}" for x in lst))
    put(ws_f, k, 4, min(x["day"] for x in lst), fmt_="dd/mm/yyyy")
    put(ws_f, k, 5, sugg)
    ws_f.row_dimensions[k].height = 15 * max(2, len(lst)) + 6
    k += 1
k += 1
put(ws_f, k, 1, "IMPRESE / PARTITE IVA", SUB, border=False)
k += 1
header(ws_f, k, ["Impresa", "Settore", "Polizze attuali", "Giorno contatto", "Priorità di proposta"], None)
for c in sorted([c for c in recs if c["impresa"]], key=lambda c: c["day"]):
    k += 1
    put(ws_f, k, 1, f"{c['id']} – {c['nome_full']}")
    put(ws_f, k, 2, c["prof"])
    put(ws_f, k, 3, c["cop"].strip(" -").replace(" - ", "; "))
    put(ws_f, k, 4, c["day"], fmt_="dd/mm/yyyy")
    put(ws_f, k, 5, " → ".join(PSHORT.get(p, p) for p in c["prodotti"][:4]))
    ws_f.row_dimensions[k].height = 30

# ---------------------------------------------------------------- Leggimi
ws_r = wb.create_sheet("Leggimi", 0)
ws_r.column_dimensions["A"].width = 4
ws_r.column_dimensions["B"].width = 140
put(ws_r, 1, 2, "PIANO VENDITE FULL IMMERSION – Portafoglio clienti", TITLE, Alignment(vertical="center"), border=False)
n_cls = {k_: sum(1 for c in recs if c["classe"] == k_) for k_ in "ABC"}
lines = [
    ("COME USARLO OGNI GIORNO", None),
    ("1. Apri 'Piano & Pipeline' e filtra la colonna Data sul giorno: trovi chi chiamare, a che ora, perché e cosa proporre.", 0),
    ("2. Clicca 'Apri scheda' per lo script completo: apertura, motivo, frase-ponte, 2 date da proporre, WhatsApp, email, domande, obiezioni, proposta, documenti.", 0),
    ("3. Dopo ogni contatto compila le celle GIALLE (esito, STATO, data appuntamento, premi). La Dashboard e il Calendario si aggiornano da soli.", 0),
    ("4. Quando fissi un appuntamento scrivilo anche in 'Agenda Appuntamenti' (4 slot al giorno, 3 il sabato) per evitare sovrapposizioni.", 0),
    ("5. Segui 'Giornata Tipo' per gli orari e 'Processo Vendita' per le fasi dall'hook alla firma.", 0),
    ("COME È STATO COSTRUITO IL PIANO", None),
    (f"• Analizzate {len(recs)} posizioni del REPORT.xlsx: {sum(1 for c in recs if not c['impresa'])} persone e {sum(1 for c in recs if c['impresa'])} imprese/P.IVA. "
     f"Clienti con lo stesso numero di telefono = una sola chiamata; stesso indirizzo = stesso giorno (appuntamento di famiglia).", 0),
    ("• Per ogni cliente il file calcola i gap (TCM, salute, non autosufficienza, previdenza, risparmio, casa, auto, impresa/CAT NAT) "
     "dalle colonne di scopertura e dai prodotti in essere, e assegna un punteggio di bisogno per prodotto in base a età, professione e iniziative Allianz già assegnate.", 0),
    (f"• Score 0–100 = opportunità (40) + premi annui (15) + relazione/anzianità (15) + urgenza quietanza (15) + età (10) + iniziativa Allianz (5) + disdetta (5). "
     f"Classe A ≥ 62 ({n_cls['A']} clienti), B 45–61 ({n_cls['B']}), C < 45 ({n_cls['C']}).", 0),
    ("• Ordine delle chiamate: prima le URGENZE (disdette negli ultimi 12 mesi, imprese senza CAT NAT, quietanze entro 30 gg, chiamate circa 12 gg prima della scadenza), "
     "poi tutti gli altri per score decrescente. Ex clienti (win-back) nell'ultima giornata del giro.", 0),
    (f"• Ritmo full immersion: {CAP_FERIALE} nuovi contatti lun–ven + {CAP_SABATO} il sabato, più i richiami; 3–4 appuntamenti al giorno. "
     "Il giro completo del portafoglio si chiude in circa 3 settimane e mezzo; le settimane 5–6 sono per chiusure, referral e campagna previdenza di fine anno.", 0),
    ("• Fascia oraria per profilo: imprenditori 8:30–9:30, pensionati/casalinghe 9:30–11:30, operai 12:30–14:00, impiegati/studenti 17:30–19:30.", 0),
    ("• Lei/tu: 'Lei' dai 40 anni in su e per le imprese, 'tu' sotto i 40. Cambialo pure se con il cliente usi già il tu.", 0),
    ("• Luogo dell'appuntamento: in azienda per imprese e artigiani, a domicilio dai 68 anni, videochiamata fuori Veneto, in agenzia per tutti gli altri.", 0),
    ("SKILL APPLICATE", None),
    ("insurance-agent / insurance-adviser (analisi dei 7 rischi, metodo DIME per il capitale TCM) · annuity-life-insurance-sales (fact-finding, previdenza e rendita, gestione obiezioni consulenziale) · "
     "prospecting e sales-enablement (segmentazione, cadenza a 3 tentativi, script, libreria obiezioni) · marketing-psychology (leve etiche: reciprocità, impegno, ancoraggio, scadenze reali) · "
     "sms/WhatsApp ed emails (sequenze messaggi, finestre orarie) · referrals (richiesta nominativi post-firma) · churn-prevention (recupero disdette e win-back) · compliance-officer (IDD, privacy, IBIPs) · xlsx.", 0),
    ("ATTENZIONE – DA VERIFICARE", None),
    ("• Nomi e regole dei prodotti (Lovia, Ultra Salute, Longevity Care, Fondo Pensione Aperto, Nuovi Orizzonti, Ultra Casa, Ultra Impresa, CAT NAT) sono presi dal tuo portafoglio: "
     "verifica in agenzia le condizioni e le promozioni in corso prima di citarle al cliente. Le iniziative (es. 'Infortuni da circolazione 30% OTT DIC') sono riportate come da report: "
     "verifica se il 30% è uno sconto per il cliente o un obiettivo commerciale, prima di nominarlo.", 0),
    ("• Fiscalità: deducibilità dei versamenti al fondo pensione fino a 5.164,57 €/anno (regola generale). Per P.IVA forfettarie e casi particolari rimanda al commercialista.", 0),
    ("• CAT NAT imprese: obbligo introdotto dalla Legge di Bilancio 2024 (L. 213/2023) con scadenze scaglionate nel 2025. Verifica con l'agenzia la normativa aggiornata prima della chiamata.", 0),
    (f"• Privacy: {sum(1 for c in recs if not c['privacy'])} clienti NON hanno il consenso commerciale (segnalati in rosso): solo contatti di servizio finché non lo firmano.", 0),
    ("• Il file contiene dati personali dei clienti: conservalo solo su dispositivi aziendali protetti e non inoltrarlo.", 0),
    ("• Ipotesi di conversione nella Dashboard (celle gialle) sono stime di partenza: aggiornale con i tuoi dati dopo la prima settimana.", 0),
    ("LEGENDA COLORI", None),
    ("Celle GIALLE = da compilare · Testo BLU = ipotesi modificabili · Classe A verde, B gialla, C grigia · Gap 'Sì' in arancio = cliente scoperto · Privacy 'NO' in rosso.", 0),
]
rr = 3
for t, kind in lines:
    if kind is None:
        rr += 1
        put(ws_r, rr, 2, t, SUB, Alignment(vertical="center"), border=False)
    else:
        put(ws_r, rr, 2, t, B_FONT, WRAP, border=False)
        ws_r.row_dimensions[rr].height = 15 * (1 + len(t) // 150)
    rr += 1

# ordine fogli
order = ["Leggimi", "Dashboard", "Calendario 6 settimane", "Giornata Tipo", "Piano & Pipeline", "Schede Clienti",
         "Processo Vendita", "Script & Obiezioni", "Messaggi WA-Email", "Agenda Appuntamenti", "Famiglie & Imprese", "Liste"]
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
wb.active = 0
wb.save(OUT)

# riepilogo per controllo
import collections
print("clienti", len(recs), "unità", len(units))
print("classi", n_cls)
print("p1", collections.Counter(c["p1"] for c in recs))
print("hook", collections.Counter(c["hook_code"] for c in recs))
print("per giorno", sorted(collections.Counter(c["day"] for c in recs).items()))
