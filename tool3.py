# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import datetime
import xlsxwriter
from itertools import combinations

# ======================================
# Costanti disciplinare
# ======================================

RI_DICT = {
    1: 0.0,  2: 0.0,  3: 0.489, 4: 0.805, 5: 1.059,
    6: 1.18, 7: 1.252, 8: 1.317, 9: 1.373, 10: 1.406,
    11: 1.419, 12: 1.445, 13: 1.46, 14: 1.471, 15: 1.485
}
CR_THRESHOLD = 0.10

LIVELLI_GIUDIZIO = [
    ("Ottimo", 1.0),
    ("Più che adeguato", 0.8),
    ("Adeguato", 0.6),
    ("Parzialmente adeguato", 0.4),
    ("Scarsamente adeguato", 0.2),
    ("Inadeguato", 0.0),
]

SCALE_DEF = {
    "Versione 1 (1–9, Saaty)": {
        "valori": [1.0, 2, 3, 4, 5, 6, 7, 8, 9],
        "etichette": {
            1.0: "Parità",
            2: "Preferenza minima / molto piccola",
            3: "Preferenza media",
            4: "Preferenza tra media ed elevata",
            5: "Preferenza elevata",
            6: "Tra elevata e molto elevata",
            7: "Preferenza molto elevata",
            8: "Tra molto elevata e massima",
            9: "Preferenza massima"
        }
    },
    "Versione 2 (1–5 mod.)": {
        "valori": [1.0, 1.25, 1.5, 2, 3, 4, 5],
        "etichette": {
            1.0: "Parità",
            1.25: "Preferenza minima",
            1.5: "Preferenza piccola",
            2: "Preferenza media",
            3: "Preferenza elevata",
            4: "Preferenza molto elevata",
            5: "Preferenza massima"
        }
    },
    "Versione 3 (1–5 mod. fine)": {
        "valori": [1.0, 1.1, 1.25, 1.5, 2, 3, 5],
        "etichette": {
            1.0: "Parità",
            1.1: "Preferenza minima",
            1.25: "Preferenza piccola",
            1.5: "Preferenza media",
            2: "Preferenza elevata",
            3: "Preferenza molto elevata",
            5: "Preferenza massima"
        }
    },
}

# ======================================
# Utility AHP
# ======================================

def compute_ahp_weights(matrix: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    m = np.array(matrix, dtype=float)
    # Evita problemi se presenti parità/valori 1: log(1)=0 ok; evitare zeri
    if np.any(m <= 0):
        raise ValueError("La matrice AHP deve contenere solo valori positivi.")
    geom_means = np.exp(np.mean(np.log(m), axis=1))
    weights = geom_means / np.sum(geom_means) if geom_means.sum() > 0 else np.zeros_like(geom_means)
    return geom_means, weights

def lambda_ci_cr(matrix: np.ndarray, weights: np.ndarray) -> tuple[float, float, float]:
    n = matrix.shape[0]
    if n <= 1:
        return (float(n), 0.0, 0.0)
    A = np.array(matrix, dtype=float)
    w = np.array(weights, dtype=float)
    Aw = A.dot(w)
    lambda_i = Aw / np.where(w != 0, w, 1)  # evita divisioni per zero
    lam = float(np.mean(lambda_i))
    ci = (lam - n) / (n - 1) if n > 1 else 0.0
    ri = RI_DICT.get(n, 1.537)
    cr = ci / ri if ri != 0 else 0.0
    return lam, ci, cr

def make_pairwise_matrix_from_rows(active_competitors: list[str], rows: list[dict]) -> np.ndarray:
    """
    rows: lista di dict con chiavi:
      - 'A', 'B' (nomi concorrenti)
      - 'Preferito' in {"A","B","Parità"}
      - 'Fattore' float (>1 se Preferito != "Parità"; altrimenti 1.0)
    """
    k = len(active_competitors)
    idx = {name: i for i, name in enumerate(active_competitors)}
    M = np.ones((k, k), dtype=float)
    for r in rows:
        a = r["A"]; b = r["B"]
        pref = (r.get("Preferito") or "").strip()
        try:
            f = float(r.get("Fattore", 1.0)) if pref != "Parità" else 1.0
        except Exception:
            f = 1.0
        if a not in idx or b not in idx:
            continue
        i, j = idx[a], idx[b]
        if i == j:
            continue
        if pref == "Parità":
            val = 1.0
        elif pref == "A":
            val = max(float(f), 1.0)
        elif pref == "B":
            val = 1.0 / max(float(f), 1.0)
        else:
            # Se non specificato, trattiamo come parità
            val = 1.0
        M[i, j] = val
        M[j, i] = 1.0 / val if val != 0 else 0.0
    # assicurati diagonale 1
    np.fill_diagonal(M, 1.0)
    return M

def normalize_definitivi(weights: np.ndarray) -> np.ndarray:
    w = np.array(weights, dtype=float)
    m = w.max() if w.size else 1.0
    return w / m if m > 0 else w

def now_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ======================================
# Generazione modelli Excel per commissari
# ======================================

def build_commissioner_workbook_bytes(meta: dict, commissario: str) -> bytes:
    """
    Crea un workbook per un singolo commissario con tutte le schede 'per Lotto'.
    meta contiene:
      - scala_version, schema_punteggio, riparametrizza_aggregati
      - lots, criteria, competitors, ptmax
      - liv_labels, liv_values (per validazione)
      - scale_vals (>1 inclusi e 1.0 in lista per comodità)
      - esclusioni: dict[lot][crit] -> set(nomi concorrenti esclusi)
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book

        fmt_h1 = wb.add_format({"bold": True, "font_size": 14, "font_color": "#FFFFFF", "bg_color": "#2F75B5", "align": "left", "valign": "vcenter"})
        fmt_header = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#4BACC6", "align": "center", "valign": "vcenter", "border": 1})
        fmt_label = wb.add_format({"bold": True, "font_color": "#444444", "bg_color": "#E7E6E6", "align": "left", "valign": "vcenter", "border": 1})
        fmt_bordo = wb.add_format({"border": 1})
        fmt_num = wb.add_format({"num_format": "0.000000", "border": 1})
        fmt_locked = wb.add_format({"border": 1, "align": "center", "valign": "vcenter", "locked": True})
        fmt_input = wb.add_format({"bg_color": "#FFF2CC", "border": 1, "align": "center", "valign": "vcenter", "locked": False})

        # VALIDAZIONI
        ws_val = wb.add_worksheet("VALIDAZIONI")
        ws_val.hide()
        ws_val.write(0, 0, "LivelliLabel"); ws_val.write(0, 1, "LivelliVal")
        for i, (lab, val) in enumerate(meta["livelli"], start=1):
            ws_val.write(i, 0, lab)
            ws_val.write_number(i, 1, float(val))
        ws_val.write(0, 3, "ScalaValori")
        for i, v in enumerate(meta["scale_vals"], start=1):
            ws_val.write_number(i, 3, float(v))

        # Range per data validation
        from xlsxwriter.utility import xl_rowcol_to_cell
        liv_first = xl_rowcol_to_cell(1, 0)
        liv_last  = xl_rowcol_to_cell(len(meta["livelli"]), 0)
        liv_range = f"=VALIDAZIONI!${liv_first}:${liv_last}"

        sc_first = xl_rowcol_to_cell(1, 3)
        sc_last  = xl_rowcol_to_cell(len(meta["scale_vals"]), 3)
        scala_range = f"=VALIDAZIONI!${sc_first}:${sc_last}"

        # IMPOSTAZIONI
        ws_set = wb.add_worksheet("Impostazioni")
        ws_set.set_column("A:A", 28); ws_set.set_column("B:B", 70)
        ws_set.write("A1", "Report generato", fmt_label); ws_set.write("B1", meta["timestamp"], fmt_locked)
        ws_set.write("A2", "Commissario", fmt_label); ws_set.write("B2", commissario, fmt_locked)
        ws_set.write("A3", "Versione scala", fmt_label); ws_set.write("B3", meta["scala_version"], fmt_locked)
        ws_set.write("A4", "Schema punteggio", fmt_label); ws_set.write("B4", meta["schema_punteggio"], fmt_locked)
        ws_set.write("A5", "Rip. finale su aggregati", fmt_label); ws_set.write("B5", "Sì" if meta["riparametrizza_aggregati"] else "No", fmt_locked)

        ws_set.write("A7", "Criteri e PTmax", fmt_h1)
        ws_set.write("A8", "Criterio", fmt_header); ws_set.write("B8", "PTmax", fmt_header)
        for i, (c, p) in enumerate(zip(meta["criteria"], meta["ptmax"]), start=9):
            ws_set.write(i, 0, c, fmt_bordo); ws_set.write_number(i, 1, float(p), fmt_num)

        ws_set.write("D1", "Lotti", fmt_h1)
        ws_set.write("D2", "Elenco lotti", fmt_header)
        for i, l in enumerate(meta["lots"], start=3):
            ws_set.write(i, 3, l, fmt_bordo)

        ws_set.write("F1", "Concorrenti", fmt_h1)
        ws_set.write("F2", "Elenco concorrenti", fmt_header)
        for i, c in enumerate(meta["competitors"], start=3):
            ws_set.write(i, 5, c, fmt_bordo)

        # ISTRUZIONI
        ws_help = wb.add_worksheet("Istruzioni")
        ws_help.set_column("A:A", 120)
        ws_help.write("A1", "Come compilare", fmt_h1)
        ws_help.write("A3", "Per ciascun Lotto e Criterio:", fmt_label)
        ws_help.write("A4", "- Se i concorrenti attivi per il criterio sono meno di 3, compilare la colonna 'Giudizio' scegliendo un'etichetta; il valore verrà calcolato automaticamente.", fmt_bordo)
        ws_help.write("A5", "- Se i concorrenti attivi sono 3 o più, per ogni coppia indicare: Preferito (A/B/Parità) e Fattore (solo se non Parità).", fmt_bordo)
        ws_help.write("A7", f"Scala utilizzata: {meta['scala_version']}", fmt_bordo)
        ws_help.write("A8", "Scegli un fattore coerente con la scala (menu a discesa).", fmt_bordo)

        # Schede per lotto
        for lot in meta["lots"]:
            ws = wb.add_worksheet(lot[:31])
            ws.set_column("A:A", 26); ws.set_column("B:B", 26); ws.set_column("C:E", 20); ws.set_column("F:H", 16)
            ws.merge_range("A1:D1", f"Raccolta giudizi — Lotto: {lot}", fmt_h1)
            ws.write("A2", "Commissario", fmt_label); ws.write("B2", commissario, fmt_locked)
            ws.write("A3", "Scala", fmt_label); ws.write("B3", meta["scala_version"], fmt_locked)
            ws.write("A4", "Schema", fmt_label); ws.write("B4", meta["schema_punteggio"], fmt_locked)
            ws.write("A5", "Generato il", fmt_label); ws.write("B5", meta["timestamp"], fmt_locked)

            r = 7
            for crit_idx, crit in enumerate(meta["criteria"]):
                # concorrenti attivi (eventuali esclusioni predefinite)
                excluded = set(meta.get("esclusioni", {}).get(lot, {}).get(crit, set()))
                active = [c for c in meta["competitors"] if c not in excluded]
                k = len(active)

                ws.write(r, 0, "CRITERIO", fmt_header); ws.write(r, 1, crit, fmt_bordo)
                ws.write(r, 2, "PTmax", fmt_header); ws.write_number(r, 3, float(meta["ptmax"][crit_idx]), fmt_num)
                r += 1

                if k < 3:
                    # Giudizi discreti
                    ws.write(r, 0, "Modalità", fmt_label); ws.write(r, 1, "Discreti (k<3)", fmt_locked); r += 1
                    ws.write(r, 0, "Concorrente", fmt_header)
                    ws.write(r, 1, "Giudizio (etichetta)", fmt_header)
                    ws.write(r, 2, "Valore (auto)", fmt_header); r += 1
                    # Tabella
                    for i, name in enumerate(active):
                        ws.write(r+i, 0, name, fmt_bordo)
                        # Etichetta con data validation
                        ws.write(r+i, 1, "", fmt_input)
                        ws.data_validation(r+i, 1, r+i, 1, {"validate": "list", "source": liv_range})
                        # Valore via VLOOKUP
                        # cerca l'etichetta in VALIDAZIONI colonna A, restituisce colonna B
                        from xlsxwriter.utility import xl_rowcol_to_cell
                        lab_cell = xl_rowcol_to_cell(r+i, 1)
                        ws.write_formula(r+i, 2, f'=IFERROR(VLOOKUP({lab_cell},VALIDAZIONI!$A:$B,2,FALSE),0)', fmt_locked)
                    r = r + k + 2

                else:
                    # Pairwise
                    ws.write(r, 0, "Modalità", fmt_label); ws.write(r, 1, "Confronti a coppie (k≥3)", fmt_locked); r += 1
                    ws.write(r, 0, "A", fmt_header)
                    ws.write(r, 1, "B", fmt_header)
                    ws.write(r, 2, "Preferito (A/B/Parità)", fmt_header)
                    ws.write(r, 3, "Fattore (se A/B)", fmt_header)
                    r += 1
                    # Righe coppie
                    for (a, b) in combinations(active, 2):
                        ws.write(r, 0, a, fmt_bordo)
                        ws.write(r, 1, b, fmt_bordo)
                        ws.write(r, 2, "", fmt_input)
                        ws.data_validation(r, 2, r, 2, {"validate": "list", "source": ["A", "B", "Parità"]})
                        ws.write(r, 3, "", fmt_input)
                        ws.data_validation(r, 3, r, 3, {"validate": "list", "source": scala_range})
                        r += 1
                    r += 1  # spazio tra i criteri

    return buf.getvalue()

# ======================================
# Parsing modelli compilati (upload) e aggregazione
# ======================================

def parse_commissioner_workbook(file_bytes: bytes) -> dict:
    """
    Ritorna:
    {
      'commissario': str,
      'scala_version': str,
      'schema_punteggio': str,
      'lotti': {
         lot: {
           crit: {
             'mode': 'discreti'|'pairwise',
             'active_competitors': [...],
             'discreti': {competitor: value_float}   # se discreti
             'pairwise_rows': [ {'A':..., 'B':..., 'Preferito':..., 'Fattore':...}, ... ]  # se pairwise
           }, ...
         }, ...
      }
    }
    """
    out = {"commissario": "", "scala_version": "", "schema_punteggio": "", "lotti": {}}
    # Leggi foglio impostazioni
    imp = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Impostazioni", header=None, engine="openpyxl")
    try:
        out["commissario"] = str(imp.iloc[1, 1]).strip()
    except Exception:
        out["commissario"] = ""
    try:
        out["scala_version"] = str(imp.iloc[2, 1]).strip()
    except Exception:
        out["scala_version"] = ""
    try:
        out["schema_punteggio"] = str(imp.iloc[3, 1]).strip()
    except Exception:
        out["schema_punteggio"] = ""

    # Elenco fogli (lotti) esclusi i tecnici
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    sheets = [s for s in xls.sheet_names if s not in ("Impostazioni", "VALIDAZIONI", "Istruzioni")]

    for lot in sheets:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=lot, header=None, engine="openpyxl")
        # Scansione: cerca blocchi "CRITERIO"
        lot_dict = {}
        r = 0
        n_rows, n_cols = df.shape
        while r < n_rows:
            cell = str(df.iat[r, 0]).strip() if not pd.isna(df.iat[r, 0]) else ""
            if cell == "CRITERIO":
                crit_name = str(df.iat[r, 1]).strip()
                # PTmax in (r,3) non necessario per parsing (lo abbiamo da setup)
                r += 1
                mode_label = str(df.iat[r, 1]).strip().lower()
                r += 1
                if "discreti" in mode_label:
                    # Header riga: Concorrente | Giudizio | Valore
                    r0 = r
                    discreti = {}
                    # Leggi fino a riga vuota
                    r += 1
                    while r < n_rows:
                        c0 = df.iat[r, 0]
                        if pd.isna(c0) or str(c0).strip() == "":
                            break
                        name = str(c0).strip()
                        val = df.iat[r, 2]
                        try:
                            valf = float(val)
                        except Exception:
                            valf = 0.0
                        discreti[name] = valf
                        r += 1
                    active = list(discreti.keys())
                    lot_dict[crit_name] = {
                        "mode": "discreti",
                        "active_competitors": active,
                        "discreti": discreti
                    }
                    r += 1  # salta riga vuota/spazio
                else:
                    # Pairwise: Header A | B | Preferito | Fattore
                    rows = []
                    # leggi fino a riga vuota
                    while r < n_rows:
                        a = df.iat[r, 0]
                        if pd.isna(a) or str(a).strip() == "":
                            break
                        b = df.iat[r, 1]
                        pref = df.iat[r, 2]
                        fatt = df.iat[r, 3]
                        rows.append({
                            "A": str(a).strip(),
                            "B": str(b).strip(),
                            "Preferito": "" if pd.isna(pref) else str(pref).strip(),
                            "Fattore": 0 if pd.isna(fatt) else fatt
                        })
                        r += 1
                    # attivi = unione nomi A,B
                    names = set()
                    for rr in rows:
                        names.add(rr["A"]); names.add(rr["B"])
                    active = sorted(list(names))
                    lot_dict[crit_name] = {
                        "mode": "pairwise",
                        "active_competitors": active,
                        "pairwise_rows": rows
                    }
                    r += 1  # riga spazio
            else:
                r += 1

        out["lotti"][lot] = lot_dict
    return out

def aggregate_results(meta: dict, uploads: list[dict]) -> dict:
    """
    Aggrega i risultati tra commissari.
    meta: setup dell'app (lots, criteria, competitors, ptmax, schema, riparametrizzazione, esclusioni opzionali)
    uploads: lista di strutture parse dei file dei commissari (parse_commissioner_workbook)
    Ritorna un oggetto con strutture per report finale.
    """
    results = {"per_lotto": {}, "log": []}

    for lot in meta["lots"]:
        lot_res = {
            "criterio_results": {},          # crit -> dict (aggregati, per commissario, punti_criterio)
            "punti_per_criterio": None,
            "graduatoria": None
        }
        # per ogni criterio aggrega
        for crit_idx, crit in enumerate(meta["criteria"]):
            ptmax = float(meta["ptmax"][crit_idx])
            # concorrenti attivi da setup (esclusioni opzionali)
            excluded = set(meta.get("esclusioni", {}).get(lot, {}).get(crit, set()))
            active = [c for c in meta["competitors"] if c not in excluded]
            k = len(active)

            per_comm = []  # lista dict per commissario con definitivi, CR, ecc.
            for up in uploads:
                comm_name = up.get("commissario", "N/D")
                lot_block = up.get("lotti", {}).get(lot, {})
                cblock = lot_block.get(crit, None)
                if not cblock:
                    continue  # commissario non ha valutato questo lot/crit
                # Verifica coerenza attivi
                # Se gli attivi dal file non coincidono, sovrascriviamo con attivi di gara (i dati estranei vengono ignorati)
                act = [c for c in active if c in cblock["active_competitors"]]
                if len(act) < len(active):
                    results["log"].append(f"[{comm_name}] {lot}/{crit}: alcuni concorrenti attivi mancanti nel file; uso insieme atteso dall'app.")
                    act = active[:]  # Mantieni l'ordine definito dall'app

                if k < 3 and cblock["mode"] == "discreti":
                    # Valori discreti -> definitivi = normalizzazione al max 1
                    vals = [float(cblock["discreti"].get(c, 0.0)) for c in act]
                    provv = np.array(vals, dtype=float)
                    definitivi = normalize_definitivi(provv)
                    per_comm.append({
                        "commissario": comm_name,
                        "mode": "discreti",
                        "active_competitors": act,
                        "provvisori": provv,
                        "definitivi": definitivi,
                        "CR": None,
                        "lambda_max": None,
                        "CI": None
                    })
                elif k >= 3 and cblock["mode"] == "pairwise":
                    # Ricostruisci matrice e calcola pesi/CR
                    M = make_pairwise_matrix_from_rows(act, cblock["pairwise_rows"])
                    try:
                        gm, w = compute_ahp_weights(M)
                        lam, ci, cr = lambda_ci_cr(M, w)
                        definitivi = normalize_definitivi(w)
                    except Exception as e:
                        results["log"].append(f"[{comm_name}] {lot}/{crit}: errore nel calcolo AHP ({e}); uso pesi nulli.")
                        w = np.zeros((len(act),)); definitivi = w; lam=ci=cr=None
                    per_comm.append({
                        "commissario": comm_name,
                        "mode": "pairwise",
                        "active_competitors": act,
                        "matrix": M,
                        "provvisori": w,
                        "definitivi": definitivi,
                        "CR": cr,
                        "lambda_max": lam,
                        "CI": ci
                    })
                else:
                    # Modalità non coerente con k: cerchiamo di interpretare
                    if k < 3 and cblock["mode"] == "pairwise":
                        results["log"].append(f"[{comm_name}] {lot}/{crit}: ricevuti confronti a coppie ma attivi <3; interpreto come discreti (media vittorie).")
                        # fallback: conteggio vittorie come proxy (AHP non applicabile)
                        counts = {c: 0 for c in act}
                        for r in cblock["pairwise_rows"]:
                            if r.get("Preferito") == "A":
                                counts[r["A"]] = counts.get(r["A"], 0) + 1
                            elif r.get("Preferito") == "B":
                                counts[r["B"]] = counts.get(r["B"], 0) + 1
                        provv = np.array([counts.get(c, 0) for c in act], dtype=float)
                        definitivi = normalize_definitivi(provv)
                        per_comm.append({
                            "commissario": comm_name,
                            "mode": "discreti",
                            "active_competitors": act,
                            "provvisori": provv,
                            "definitivi": definitivi,
                            "CR": None, "lambda_max": None, "CI": None
                        })
                    else:
                        results["log"].append(f"[{comm_name}] {lot}/{crit}: modalita' discreti ma attivi >=3; impossibile interpretare. Ignoro.")
                        continue

            # Aggregazione
            if per_comm:
                arr_def = np.vstack([c["definitivi"] for c in per_comm])
                def_aggregati = arr_def.mean(axis=0) if arr_def.size else np.zeros((k,))
                if meta["schema_punteggio"].startswith("Interdipendente") and meta["riparametrizza_aggregati"] and def_aggregati.size:
                    def_aggregati = normalize_definitivi(def_aggregati)
            else:
                def_aggregati = np.zeros((k,))

            # Punteggi per criterio
            punti_criterio = {c: 0.0 for c in meta["competitors"]}
            for i, c in enumerate(active):
                punti_criterio[c] = def_aggregati[i] * ptmax
            # esclusi a 0 già previsto

            lot_res["criterio_results"][crit] = {
                "mode": "discreti" if k < 3 else "pairwise",
                "active_competitors": active,
                "commissari": per_comm,
                "def_aggregati": def_aggregati,
                "punti_criterio": punti_criterio
            }

        # Riepilogo lotto: punti per criterio e graduatoria
        punti_per_criterio_df = pd.DataFrame(index=meta["competitors"])
        for crit in meta["criteria"]:
            punti_c = lot_res["criterio_results"][crit]["punti_criterio"]
            punti_per_criterio_df[crit] = pd.Series(punti_c)
        punti_per_criterio_df = punti_per_criterio_df.fillna(0.0)

        serie_totale = punti_per_criterio_df.sum(axis=1).sort_values(ascending=False)
        lot_res["punti_per_criterio"] = punti_per_criterio_df
        lot_res["graduatoria"] = serie_totale

        results["per_lotto"][lot] = lot_res

    return results

# ======================================
# Export report finale Excel
# ======================================

def export_final_report(meta: dict, agg: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_h1 = wb.add_format({"bold": True, "font_size": 14, "font_color": "#FFFFFF", "bg_color": "#2F75B5", "align": "left", "valign": "vcenter"})
        fmt_header = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#4BACC6", "align": "center", "valign": "vcenter", "border": 1})
        fmt_label = wb.add_format({"bold": True, "font_color": "#444444", "bg_color": "#E7E6E6", "align": "left", "valign": "vcenter", "border": 1})
        fmt_bordo = wb.add_format({"border": 1})
        fmt_num = wb.add_format({"num_format": "0.000000", "border": 1})
        fmt_green = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        fmt_red   = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        # Impostazioni
        ws_set = wb.add_worksheet("Impostazioni")
        ws_set.set_column("A:A", 28); ws_set.set_column("B:B", 70)
        ws_set.write("A1", "Report generato", fmt_label); ws_set.write("B1", now_str(), fmt_bordo)
        ws_set.write("A2", "Versione scala", fmt_label); ws_set.write("B2", meta["scala_version"], fmt_bordo)
        ws_set.write("A3", "Schema punteggio", fmt_label); ws_set.write("B3", meta["schema_punteggio"], fmt_bordo)
        ws_set.write("A4", "Rip. finale su aggregati", fmt_label); ws_set.write("B4", "Sì" if meta["riparametrizza_aggregati"] else "No", fmt_bordo)

        ws_set.write("A6", "Criteri e PTmax", fmt_h1)
        ws_set.write("A7", "Criterio", fmt_header); ws_set.write("B7", "PTmax", fmt_header)
        for i, (c, p) in enumerate(zip(meta["criteria"], meta["ptmax"]), start=8):
            ws_set.write(i, 0, c, fmt_bordo); ws_set.write_number(i, 1, float(p), fmt_num)

        ws_set.write("D1", "Lotti", fmt_h1)
        ws_set.write("D2", "Elenco lotti", fmt_header)
        for i, l in enumerate(meta["lots"], start=3):
            ws_set.write(i, 3, l, fmt_bordo)

        ws_set.write("F1", "Concorrenti", fmt_h1)
        ws_set.write("F2", "Elenco concorrenti", fmt_header)
        for i, c in enumerate(meta["competitors"], start=3):
            ws_set.write(i, 5, c, fmt_bordo)

        # Fogli per lotto
        for lot, lot_data in agg["per_lotto"].items():
            ws = wb.add_worksheet(lot[:31])
            ws.set_column("A:A", 22); ws.set_column("B:B", 16); ws.set_column("C:Z", 13)
            ws.merge_range("A1:D1", f"Report Tecnico — Lotto: {lot}", fmt_h1)
            ws.write("A2", "Scala:", fmt_label); ws.write("B2", meta["scala_version"], fmt_bordo)
            ws.write("A3", "Schema:", fmt_label); ws.write("B3", meta["schema_punteggio"], fmt_bordo)
            ws.write("A4", "Rip. finale:", fmt_label); ws.write("B4", "Sì" if meta["riparametrizza_aggregati"] else "No", fmt_bordo)

            r = 6
            ws.write(r, 0, "Criterio", fmt_header)
            ws.write(r, 1, "PTmax", fmt_header)
            ws.write(r, 2, "Modalità", fmt_header)
            ws.write(r, 3, "Definitivi aggregati (max=1 se applicato)", fmt_header)
            ws.write(r, 4, "Punti (× PTmax)", fmt_header)
            r += 1

            # Blocco per criterio
            for crit_idx, crit in enumerate(meta["criteria"]):
                cres = lot_data["criterio_results"][crit]
                ptmax = float(meta["ptmax"][crit_idx])
                active = cres["active_competitors"]
                def_agg = cres["def_aggregati"]
                punti_map = cres["punti_criterio"]

                ws.write(r, 0, crit, fmt_bordo)
                ws.write_number(r, 1, ptmax, fmt_num)
                ws.write(r, 2, "Discreti" if cres["mode"] == "discreti" else "Pairwise", fmt_bordo)
                # definitivi
                for i, c in enumerate(active):
                    ws.write(r, 3 + i, float(def_agg[i]), fmt_num)
                r += 1

                # Dettaglio per commissario (CR evidenziato)
                for cdata in cres["commissari"]:
                    ws.write(r, 0, f"{crit} — {cdata['commissario']}", fmt_label)
                    if cdata["CR"] is not None:
                        ws.write(r, 1, "CR", fmt_header)
                        ws.write_number(r, 2, float(cdata["CR"]), fmt_num)
                        cell_cr = f"{xlsxwriter.utility.xl_rowcol_to_cell(r,2)}"
                        ws.conditional_format(cell_cr, {"type": "cell", "criteria": "<=", "value": CR_THRESHOLD, "format": fmt_green})
                        ws.conditional_format(cell_cr, {"type": "cell", "criteria": ">", "value": CR_THRESHOLD, "format": fmt_red})
                    # pesi definitivi di dettaglio
                    for i, name in enumerate(cdata["active_competitors"]):
                        ws.write(r, 3 + i, float(cdata["definitivi"][i]), fmt_num)
                    r += 1
                r += 1  # spazio

                # Punti criterio per tutti i concorrenti
                ws.write(r, 0, f"Punti {crit} (× PTmax)", fmt_label); r += 1
                ws.write(r, 0, "Concorrente", fmt_header); col = 1
                for name in meta["competitors"]:
                    ws.write(r, col, name, fmt_header); col += 1
                r += 1
                col = 1
                for name in meta["competitors"]:
                    ws.write_number(r, col, float(punti_map.get(name, 0.0)), fmt_num); col += 1
                r += 2

            # Riepilogo e graduatoria
            ws.write(r, 0, "Riepilogo Punti per criterio", fmt_h1); r += 1
            ws.write(r, 0, "Concorrente", fmt_header)
            for j, crit in enumerate(meta["criteria"]):
                ws.write(r, 1 + j, crit, fmt_header)
            ws.write(r, 1 + len(meta["criteria"]), "Totale", fmt_header)
            r += 1
            for i, c in enumerate(meta["competitors"]):
                ws.write(r + i, 0, c, fmt_bordo)
                tot = 0.0
                for j, crit in enumerate(meta["criteria"]):
                    val = float(lot_data["criterio_results"][crit]["punti_criterio"].get(c, 0.0))
                    ws.write_number(r + i, 1 + j, val, fmt_num)
                    tot += val
                ws.write_number(r + i, 1 + len(meta["criteria"]), tot, fmt_num)
            r = r + len(meta["competitors"]) + 2

            ws.write(r, 0, "Graduatoria Tecnica", fmt_h1); r += 1
            grad = lot_data["graduatoria"]
            ws.write(r, 0, "Concorrente", fmt_header); ws.write(r, 1, "Punteggio tecnico", fmt_header)
            r += 1
            for i, (cname, v) in enumerate(grad.items()):
                ws.write(r + i, 0, cname, fmt_bordo)
                ws.write_number(r + i, 1, float(v), fmt_num)

        # Log di consistenza
        ws_log = wb.add_worksheet("Log")
        ws_log.set_column("A:A", 120)
        ws_log.write("A1", "Avvisi e note di consistenza", fmt_h1)
        for i, msg in enumerate(agg.get("log", []), start=3):
            ws_log.write(i, 0, msg, fmt_bordo)

    return buf.getvalue()

# ======================================
# App Streamlit
# ======================================

def main():
    st.set_page_config(page_title="AHP - Modelli Commissari & Aggregazione", layout="wide")
    st.title("AHP — Confronto a Coppie con raccolta esterna dei giudizi")

    st.sidebar.header("Parametri generali")
    num_lots = st.sidebar.number_input("Numero di lotti", min_value=1, value=1, step=1)
    num_criteria = st.sidebar.number_input("Numero di criteri", min_value=1, value=3, step=1)
    num_competitors = st.sidebar.number_input("Numero di concorrenti", min_value=1, value=3, step=1)
    num_commissari = st.sidebar.number_input("Numero di commissari", min_value=1, value=3, step=1)

    st.sidebar.markdown("---")
    scala_version = st.sidebar.selectbox("Versione scala disciplinare", list(SCALE_DEF.keys()), index=0)
    schema_punteggio = st.sidebar.selectbox(
        "Schema di punteggio",
        ["Interdipendente (riparametrizzazioni previste)", "Assoluto (senza riparametrizzazione finale)"],
        index=0
    )
    applica_rip_aggregati = st.sidebar.checkbox(
        "Interdipendente: applica riparametrizzazione finale su aggregati (max=1)",
        value=True
    )

    st.sidebar.markdown("---")
    lot_names = [st.sidebar.text_input(f"Nome Lotto {i+1}", value=f"Lotto{i+1}") for i in range(num_lots)]
    criterion_names = [st.sidebar.text_input(f"Criterio {j+1}", value=f"Criterio{j+1}") for j in range(num_criteria)]
    competitor_names = [st.sidebar.text_input(f"Concorrente {k+1}", value=f"Conc{k+1}") for k in range(num_competitors)]
    commissari_names = [st.sidebar.text_input(f"Commissario {c+1}", value=f"Comm{c+1}") for c in range(num_commissari)]

    st.sidebar.markdown("---")
    st.sidebar.subheader("PTmax per criterio")
    ptmax_values = []
    for j, crit in enumerate(criterion_names):
        v = st.sidebar.number_input(f"PTmax '{crit}'", min_value=0.0, value=1.0, step=0.5, key=f"ptmax_{j}")
        ptmax_values.append(v)

    st.sidebar.markdown("---")
    with st.sidebar.expander("Opzioni avanzate: esclusioni concorrenti per criterio", expanded=False):
        gestisci_esclusioni = st.checkbox("Abilita esclusioni per criterio", value=False, key="excl_enable")
        esclusioni = {}
        if gestisci_esclusioni:
            for lot in lot_names:
                st.markdown(f"**{lot}**")
                esclusioni[lot] = {}
                for crit in criterion_names:
                    sel = st.multiselect(f"Escludi su '{crit}'", competitor_names, key=f"excl_{lot}_{crit}")
                    esclusioni[lot][crit] = set(sel)

    # Mostra contesto
    st.markdown("___")
    st.subheader("Informazioni di contesto")
    st.write(
        f"**Scala:** {scala_version}  |  **Schema:** {schema_punteggio}"
        + (", **Rip. finale su aggregati: Sì**" if (schema_punteggio.startswith("Interdipendente") and applica_rip_aggregati) else ", **Rip. finale su aggregati: No**")
    )
    st.dataframe(pd.DataFrame({"Criterio": criterion_names, "PTmax": ptmax_values}).set_index("Criterio").style.format({"PTmax": "{:.4f}"}))
    st.markdown("___")

    tab1, tab2, tab3 = st.tabs(["① Genera modelli per i commissari", "② Carica modelli compilati", "③ Report & Export"])

    # Stato condiviso
    if "uploaded_parsed" not in st.session_state:
        st.session_state.uploaded_parsed = []  # lista strutture parse
    if "agg_result" not in st.session_state:
        st.session_state.agg_result = None

    # ① Generazione modelli
    with tab1:
        st.markdown("### Genera pacchetto di modelli Excel (uno per commissario)")
        if st.button("📦 Crea ZIP modelli per i commissari"):
            meta = {
                "timestamp": now_str(),
                "scala_version": scala_version,
                "schema_punteggio": schema_punteggio,
                "riparametrizza_aggregati": bool(applica_rip_aggregati),
                "lots": lot_names,
                "criteria": criterion_names,
                "competitors": competitor_names,
                "ptmax": ptmax_values,
                "livelli": LIVELLI_GIUDIZIO,
                "scale_vals": SCALE_DEF[scala_version]["valori"],
                "esclusioni": esclusioni if 'esclusioni' in locals() else {}
            }
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for comm in commissari_names:
                    wb_bytes = build_commissioner_workbook_bytes(meta, commissario=comm)
                    safe_comm = comm.replace(" ", "_").replace("/", "-")
                    z.writestr(f"Modello_Giudizi_{safe_comm}.xlsx", wb_bytes)
            zip_buf.seek(0)
            st.download_button(
                label="Scarica ZIP modelli",
                data=zip_buf.getvalue(),
                file_name="Modelli_Commissari_AHP.zip",
                mime="application/zip"
            )
            st.success("ZIP creato. Invia ogni file al relativo commissario per la compilazione.")

        st.markdown("#### Cosa vedranno i commissari")
        st.write("- Scheda **per Lotto** con blocchi **per Criterio**")
        st.write("- **Discreti (k<3)**: etichetta con menu; il valore è calcolato")
        st.write("- **Pairwise (k≥3)**: per ogni coppia **A-B** indicano Preferito e Fattore (scala)")

    # ② Upload modelli compilati
    with tab2:
        st.markdown("### Carica i file compilati dai commissari")
        files = st.file_uploader(
            "Carica uno o più file .xlsx (i modelli compilati).",
            type=["xlsx"],
            accept_multiple_files=True
        )
        if st.button("📥 Importa e valida file caricati"):
            parsed_list = []
            errs = []
            if not files:
                st.warning("Nessun file caricato.")
            else:
                for f in files:
                    try:
                        file_bytes = f.read()
                        parsed = parse_commissioner_workbook(file_bytes)
                        parsed_list.append(parsed)
                    except Exception as e:
                        errs.append(f"{f.name}: errore in lettura ({e})")
                st.session_state.uploaded_parsed = parsed_list
                if errs:
                    st.error("Alcuni file hanno dato errori:"); st.write("\n".join(errs))
                st.success(f"Importati {len(parsed_list)} file.")
                # mostra elenco commissari importati
                st.write("Commissari importati:")
                st.write([p.get("commissario","N/D") for p in parsed_list])

        # Aggregazione immediata (anteprima)
        if st.session_state.uploaded_parsed:
            st.markdown("#### Anteprima aggregazione")
            meta = {
                "scala_version": scala_version,
                "schema_punteggio": schema_punteggio,
                "riparametrizza_aggregati": bool(applica_rip_aggregati),
                "lots": lot_names,
                "criteria": criterion_names,
                "competitors": competitor_names,
                "ptmax": ptmax_values,
                "esclusioni": esclusioni if 'esclusioni' in locals() else {}
            }
            agg = aggregate_results(meta, st.session_state.uploaded_parsed)
            st.session_state.agg_result = agg

            for lot in lot_names:
                st.subheader(f"Lotto: {lot}")
                lot_data = agg["per_lotto"].get(lot, {})
                if not lot_data:
                    st.info("Nessun dato aggregato per questo lotto.")
                    continue

                # Punti per criterio
                st.write("**Punti per criterio**")
                st.dataframe(lot_data["punti_per_criterio"].style.format("{:.6f}"))
                # Graduatoria
                st.write("**Graduatoria tecnica**")
                st.dataframe(lot_data["graduatoria"].to_frame("Punteggio tecnico").style.format("{:.6f}"))

            if agg.get("log"):
                with st.expander("Log di consistenza e avvisi"):
                    for m in agg["log"]:
                        st.write("- " + m)

    # ③ Export report finale
    with tab3:
        st.markdown("### Esporta report finale")
        if st.session_state.agg_result:
            meta = {
                "scala_version": scala_version,
                "schema_punteggio": schema_punteggio,
                "riparametrizza_aggregati": bool(applica_rip_aggregati),
                "lots": lot_names,
                "criteria": criterion_names,
                "competitors": competitor_names,
                "ptmax": ptmax_values
            }
            if st.button("🔽 Esporta Excel finalizzato"):
                xlsb = export_final_report(meta, st.session_state.agg_result)
                st.download_button(
                    label="Scarica Report Finale",
                    data=xlsb,
                    file_name="Report_Tecnico_AHP_Commissari.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.info("Carica e importa i modelli compilati (Tab ②) per attivare l'export.")

if __name__ == "__main__":
    main()
