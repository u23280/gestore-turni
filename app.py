import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from datetime import datetime, timedelta
import io

# ==============================================================================
# 1. CONFIGURAZIONE COSTANTI & REGOLE
# ==============================================================================

STAFF_SOLO_MATTINA = ["Giuseppe Sergi", "Piero Cappi", "Marco Salierno"]
ROTAZIONE_ORDINE = [
    "Matteo Costanzi", "Augusto Novelli", "Paolo Nucci", "Fabrizio Loria",
    "Walter Araujo", "Alberto Rink", "Marco Celentano", "Marco Lorentino",
    "Simone Esposito"
]
STAFF_NOTTE = ROTAZIONE_ORDINE.copy()
STAFF_ODDS = ["Walter Savino", "Gennaro Auriemma", "Claudio Condemi", "Michele di Chiaro", "Aytac Yener", "Klajd Goxho"]
LISTA_ACCETTATORI = ["Giuseppe Sergi", "Antonio Mandica", "Matteo Costanzi", "Piero Cappi", "Fabrizio Loria", "Augusto Novelli", "Paolo Nucci", "Marco Salierno", "Marco Mirabella", "Alberto Rink", "Marco Lorentino", "Walter Araujo", "Marco Celentano", "Simone Esposito"]
STAFF_SOLO_POMERIGGIO = ["Simone Esposito"]
STAFF_EXTRA = ["Antonio Mandica", "Marco Mirabella"]

ALL_STAFF = list(set(STAFF_SOLO_MATTINA + STAFF_NOTTE + STAFF_ODDS + STAFF_EXTRA + STAFF_SOLO_POMERIGGIO))
ALL_STAFF.sort()
NUM_STAFF = len(ALL_STAFF)

GIORNI_SETT_IT = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]

TEXT_TO_ID = {
    "-": 0, "FERIE": 0, "REQ": 0, "ASSENTE": 0,
    "07:00": 1, "08:00": 2, "09:00": 3, "10:00": 4, "11:00": 5,
    "12:00": 6, "15:00": 7, "16:00": 8, "17:00": 10,
    "CHIUSURA": 9, "18-02": 9, "20-04": 9
}

TURNI_CONFIG = {
    1:  {"txt": "07:00",    "cat": "M", "critico": 2, "ideale": 3, "bg": "#B4C6E7", "font": "#000000"},
    2:  {"txt": "08:00",    "cat": "M", "critico": 0, "ideale": 1, "bg": "#8EA9DB", "font": "#000000"},
    3:  {"txt": "09:00",    "cat": "M", "critico": 1, "ideale": 1, "bg": "#305496", "font": "#FFFFFF"},
    4:  {"txt": "10:00",    "cat": "M", "critico": 0, "ideale": 1, "bg": "#D9E1F2", "font": "#000000"},
    5:  {"txt": "11:00",    "cat": "M", "critico": 0, "ideale": 1, "bg": "#E2EFDA", "font": "#000000"},
    6:  {"txt": "12:00",    "cat": "M", "critico": 1, "ideale": 2, "bg": "#A9D08E", "font": "#000000"},
    7:  {"txt": "15:00",    "cat": "L", "critico": 0, "ideale": 1, "bg": "#FFF2CC", "font": "#000000"},
    8:  {"txt": "16:00",    "cat": "L", "critico": 3, "ideale": 3, "bg": "#FFD966", "font": "#000000"},
    10: {"txt": "17:00",    "cat": "L", "critico": 0, "ideale": 1, "bg": "#F4B084", "font": "#000000"}, 
    9:  {"txt": "CHIUSURA", "cat": "L", "critico": 1, "ideale": 1, "bg": "#C00000", "font": "#FFFFFF"},
}
RIPOSO_CFG = {"txt": "-", "bg": "#FFFFFF", "font": "#D3D3D3"}

# ==============================================================================
# 2. MOTORE DI CALCOLO
# ==============================================================================

def parse_uploaded_schedule(file, start_date_obj, num_days):
    try:
        df = pd.read_excel(file, header=0)
        if 'Dipendente' not in df.columns: return None, "Colonna 'Dipendente' non trovata."
        previous_matrix = {}
        for idx, row in df.iterrows():
            nome = row['Dipendente']
            if nome not in ALL_STAFF: continue
            col_idx = 0
            for col_name in df.columns:
                if any(x in str(col_name) for x in ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]):
                    val_str = str(row[col_name]).strip()
                    t_id = TEXT_TO_ID.get(val_str, 0)
                    if val_str in ["FERIE", "REQ"]: t_id = 0
                    if col_idx < num_days:
                        previous_matrix[(nome, col_idx)] = t_id
                        col_idx += 1
        return previous_matrix, None
    except Exception as e: return None, str(e)

def solve_turni(start_date, weeks_to_generate, list_assenze, list_req_turni, previous_solution=None):
    NUM_SETTIMANE = weeks_to_generate
    NUM_GIORNI = 7 * NUM_SETTIMANE 
    if isinstance(start_date, datetime): start_date = start_date
    else: start_date = datetime.combine(start_date, datetime.min.time())
    
    FERIE_INDICI = {}
    REQ_OFF_INDICI = {}
    REQ_TURNI_INDICI = {}

    for item in list_assenze:
        nome = item['nome']
        tipo = item['tipo']
        d_input = item['data']
        if not isinstance(d_input, datetime): d_input = datetime.combine(d_input, datetime.min.time())
        delta = (d_input - start_date).days
        if 0 <= delta < NUM_GIORNI:
            if tipo == "Ferie": FERIE_INDICI.setdefault(nome, []).append(delta)
            elif tipo == "Richiesta OFF": REQ_OFF_INDICI.setdefault(nome, []).append(delta)

    for item in list_req_turni:
        nome = item['nome']
        d_input = item['data']
        if not isinstance(d_input, datetime): d_input = datetime.combine(d_input, datetime.min.time())
        t_id = int(str(item['turno']).split(" - ")[0])
        delta = (d_input - start_date).days
        if 0 <= delta < NUM_GIORNI: REQ_TURNI_INDICI.setdefault(nome, {})[delta] = t_id

    model = cp_model.CpModel()
    shifts = {}
    for n in range(NUM_STAFF):
        for d in range(NUM_GIORNI): shifts[(n, d)] = model.NewIntVar(0, 10, f's_{n}_{d}')

    # APPLICAZIONI INPUT
    for nome, indices in FERIE_INDICI.items():
        if nome in ALL_STAFF:
            n_idx = ALL_STAFF.index(nome)
            for d in indices: model.Add(shifts[(n_idx, d)] == 0)
    for nome, indices in REQ_OFF_INDICI.items():
        if nome in ALL_STAFF:
            n_idx = ALL_STAFF.index(nome)
            for d in indices: model.Add(shifts[(n_idx, d)] == 0)
    for nome, req_map in REQ_TURNI_INDICI.items():
        if nome in ALL_STAFF:
            n_idx = ALL_STAFF.index(nome)
            for d, t_id in req_map.items(): model.Add(shifts[(n_idx, d)] == t_id)

    # MODALITA' RIPARAZIONE
    score_stability = 0
    if previous_solution:
        for n in range(NUM_STAFF):
            nome = ALL_STAFF[n]
            for d in range(NUM_GIORNI):
                if (nome, d) in previous_solution:
                    old_val = previous_solution[(nome, d)]
                    is_now_ferie = (d in FERIE_INDICI.get(nome, []) or d in REQ_OFF_INDICI.get(nome, []))
                    if not is_now_ferie:
                        is_same = model.NewBoolVar(f'same_{n}_{d}')
                        model.Add(shifts[(n, d)] == old_val).OnlyEnforceIf(is_same)
                        model.Add(shifts[(n, d)] != old_val).OnlyEnforceIf(is_same.Not())
                        score_stability += is_same * 500

    # VINCOLI RUOLI
    for n in range(NUM_STAFF):
        nome = ALL_STAFF[n]
        is_acceptor = nome in LISTA_ACCETTATORI
        for d in range(NUM_GIORNI):
            giorno_sett = d % 7
            is_weekend = (giorno_sett == 5 or giorno_sett == 6)
            is_weekend_night = (giorno_sett == 4 or giorno_sett == 5)

            if nome == "Piero Cappi":
                for t in [4, 5, 6, 7, 8, 10, 9]: model.Add(shifts[(n, d)] != t)
                if not is_weekend: model.Add(shifts[(n, d)] != 2)
                else: 
                    model.Add(shifts[(n, d)] != 1)
                    model.Add(shifts[(n, d)] != 2)
            elif nome in STAFF_SOLO_MATTINA:
                for t in [4, 5, 6, 7, 8, 10, 9]: model.Add(shifts[(n, d)] != t)
            if nome in STAFF_SOLO_POMERIGGIO:
                for t in [1, 2, 3, 4, 5]: model.Add(shifts[(n, d)] != t)
            if nome in STAFF_ODDS:
                model.Add(shifts[(n, d)] != 9)
                model.Add(shifts[(n, d)] != 10)
            if nome == "Marco Mirabella":
                if is_weekend_night: model.Add(shifts[(n, d)] != 9)
            elif nome not in STAFF_NOTTE:
                model.Add(shifts[(n, d)] != 9)
            if not is_acceptor:
                model.Add(shifts[(n, d)] != 10)
                model.Add(shifts[(n, d)] != 9)

    for n in range(NUM_STAFF):
        if ALL_STAFF[n] == "Simone Esposito":
            for w in range(NUM_SETTIMANE):
                s = w * 7; e = s + 7
                count_12 = []
                for d in range(s, e):
                    if d < NUM_GIORNI:
                        b = model.NewBoolVar(f'sim_12_{d}')
                        model.Add(shifts[(n, d)] == 6).OnlyEnforceIf(b)
                        model.Add(shifts[(n, d)] != 6).OnlyEnforceIf(b.Not())
                        count_12.append(b)
                model.Add(sum(count_12) <= 2)

    for w in range(NUM_SETTIMANE):
        idx_rot = w % len(ROTAZIONE_ORDINE)
        nome_rot = ROTAZIONE_ORDINE[idx_rot]
        ven = 4 + (w*7); sab = 5 + (w*7); dom = 6 + (w*7)
        if ven < NUM_GIORNI:
            skip_rot = False
            if nome_rot in FERIE_INDICI and (ven in FERIE_INDICI[nome_rot] or sab in FERIE_INDICI[nome_rot]): skip_rot = True
            if nome_rot in REQ_OFF_INDICI and (ven in REQ_OFF_INDICI[nome_rot] or sab in REQ_OFF_INDICI[nome_rot]): skip_rot = True
            if nome_rot in ALL_STAFF and not skip_rot:
                idx = ALL_STAFF.index(nome_rot)
                model.Add(shifts[(idx, ven)] == 9)
                if sab < NUM_GIORNI: model.Add(shifts[(idx, sab)] == 9)
                if dom < NUM_GIORNI: model.Add(shifts[(idx, dom)] == 0)

    score_coverage = 0 
    score_balance = 0
    score_weekend_balance = 0
    staff_m_counts = [[] for _ in range(NUM_STAFF)]
    staff_l_counts = [[] for _ in range(NUM_STAFF)]

    for d in range(NUM_GIORNI):
        counts = {t: [] for t in TURNI_CONFIG}
        acc_eve = []
        seniors = []
        giorno_sett = d % 7
        is_weekend_night = (giorno_sett == 4 or giorno_sett == 5)
        
        weight_day = 20
        if giorno_sett == 5 or giorno_sett == 6: weight_day = 100 
        elif giorno_sett == 0 or giorno_sett == 4: weight_day = 1   

        for n in range(NUM_STAFF):
            nome = ALL_STAFF[n]
            for t in TURNI_CONFIG:
                b = model.NewBoolVar(f'is_{t}_{n}_{d}')
                model.Add(shifts[(n, d)] == t).OnlyEnforceIf(b)
                model.Add(shifts[(n, d)] != t).OnlyEnforceIf(b.Not())
                counts[t].append(b)
                if TURNI_CONFIG[t]['cat'] == 'M': staff_m_counts[n].append(b)
                elif TURNI_CONFIG[t]['cat'] == 'L': staff_l_counts[n].append(b)

            if nome in LISTA_ACCETTATORI:
                v_eve = model.NewBoolVar(f'acc_eve_{n}_{d}')
                model.Add(shifts[(n, d)] >= 8).OnlyEnforceIf(v_eve)
                model.Add(shifts[(n, d)] < 8).OnlyEnforceIf(v_eve.Not())
                acc_eve.append(v_eve)
                v_pres = model.NewBoolVar(f'pres_{n}_{d}')
                model.Add(shifts[(n, d)] > 0).OnlyEnforceIf(v_pres)
                model.Add(shifts[(n, d)] == 0).OnlyEnforceIf(v_pres.Not())
                seniors.append(v_pres)

        if is_weekend_night: model.Add(sum(counts[10]) == 1)
        else: model.Add(sum(counts[10]) == 0)
        model.Add(sum(counts[9]) == 1)

        for t, cfg in TURNI_CONFIG.items():
            if t == 10 or t == 9: continue
            if cfg['critico'] > 0: model.Add(sum(counts[t]) >= cfg['critico'])
            if cfg['ideale'] > cfg['critico'] or cfg['critico'] == 0:
                covered = model.NewBoolVar(f'cov_{t}_{d}')
                model.Add(sum(counts[t]) >= cfg['ideale']).OnlyEnforceIf(covered)
                weight_shift = 10
                if t == 1: weight_shift = 50 
                elif t == 7: weight_shift = 5 
                score_coverage += covered * weight_day * weight_shift

        model.Add(sum(acc_eve) >= 2)
        model.Add(sum(seniors) >= 3)

    for n in range(NUM_STAFF):
        nome = ALL_STAFF[n]
        if nome not in STAFF_SOLO_MATTINA and nome not in STAFF_SOLO_POMERIGGIO:
            tot_m = sum(staff_m_counts[n])
            tot_l = sum(staff_l_counts[n])
            diff = model.NewIntVar(0, NUM_GIORNI, f'diff_ml_{n}')
            model.Add(diff >= tot_m - tot_l)
            model.Add(diff >= tot_l - tot_m)
            score_balance -= diff * 5

    for n in range(NUM_STAFF):
        we_days_worked = []
        for w in range(NUM_SETTIMANE):
            idx_sab = 5 + (w * 7); idx_dom = 6 + (w * 7)
            if idx_dom < NUM_GIORNI:
                is_we_active = model.NewBoolVar(f'we_act_{n}_{w}')
                model.Add(shifts[(n, idx_sab)] + shifts[(n, idx_dom)] > 0).OnlyEnforceIf(is_we_active)
                model.Add(shifts[(n, idx_sab)] + shifts[(n, idx_dom)] == 0).OnlyEnforceIf(is_we_active.Not())
                we_days_worked.append(is_we_active)
        tot_we = sum(we_days_worked)
        excess = model.NewIntVar(0, NUM_SETTIMANE, f'exc_we_{n}')
        model.Add(tot_we <= int(NUM_SETTIMANE * 0.7) + excess)
        score_weekend_balance -= excess * 50

    for n in range(NUM_STAFF):
        for d in range(NUM_GIORNI - 1):
            is_weekend_night = ((d % 7) == 4 or (d % 7) == 5)
            if is_weekend_night: 
                for v in [1, 2, 3, 4, 5, 6]: model.AddForbiddenAssignments([shifts[(n, d)], shifts[(n, d+1)]], [(9, v)])
            else:
                 for v in [1, 2, 3, 4, 5, 6]: model.AddForbiddenAssignments([shifts[(n, d)], shifts[(n, d+1)]], [(9, v)])
            for v in [1, 2, 3, 4, 5]: model.AddForbiddenAssignments([shifts[(n, d)], shifts[(n, d+1)]], [(10, v)])
            for v in [1, 2, 3, 4]: model.AddForbiddenAssignments([shifts[(n, d)], shifts[(n, d+1)]], [(8, v)])
            for v in [1, 2, 3]: model.AddForbiddenAssignments([shifts[(n, d)], shifts[(n, d+1)]], [(7, v)])

        for w in range(NUM_SETTIMANE):
            s = w*7; e = s+7
            giorni_ferie_sett = 0
            if ALL_STAFF[n] in FERIE_INDICI:
                for fd in FERIE_INDICI[ALL_STAFF[n]]:
                    if s <= fd < e: giorni_ferie_sett += 1
            if ALL_STAFF[n] in REQ_OFF_INDICI:
                for fd in REQ_OFF_INDICI[ALL_STAFF[n]]:
                    if s <= fd < e: giorni_ferie_sett += 1
            target = max(0, 5 - giorni_ferie_sett)
            worked = []
            for d in range(s, min(e, NUM_GIORNI)):
                bw = model.NewBoolVar(f'wk_{n}_{d}')
                model.Add(shifts[(n, d)] > 0).OnlyEnforceIf(bw)
                model.Add(shifts[(n, d)] == 0).OnlyEnforceIf(bw.Not())
                worked.append(bw)
            model.Add(sum(worked) == target)
        
        for d in range(NUM_GIORNI - 6):
            window = [model.NewBoolVar(f'w_{n}_{d+k}') for k in range(6)]
            for k in range(6):
                model.Add(shifts[(n, d+k)] > 0).OnlyEnforceIf(window[k])
                model.Add(shifts[(n, d+k)] == 0).OnlyEnforceIf(window[k].Not())
            model.Add(sum(window) <= 5)

    for d in range(NUM_GIORNI):
        for nome in ["Giuseppe Sergi", "Marco Salierno"]:
            if nome in ALL_STAFF:
                i = ALL_STAFF.index(nome)
                b = model.NewBoolVar(f'p7_{d}_{i}')
                model.Add(shifts[(i, d)] == 1).OnlyEnforceIf(b)
                score_coverage += b * 5

    model.Maximize(score_coverage + score_balance + score_weekend_balance + score_stability)
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 300
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        workbook = writer.book
        sheet = workbook.add_worksheet("Turni")
        
        fmt_base = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_head = workbook.add_format({'border': 1, 'bold': True, 'bg_color': '#E7E6E6', 'align': 'center'})
        fmt_name = workbook.add_format({'border': 1, 'bold': True, 'align': 'left', 'valign': 'vcenter'})
        fmt_pct = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0%'})
        fmt_total = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bold': True, 'bg_color': '#FFFF00'})
        fmt_ferie = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#595959', 'font_color': '#FFFFFF', 'bold': True})
        fmt_req = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E4DFEC', 'font_color': '#000000', 'bold': True})
        
        formats = {}
        formats[0] = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': RIPOSO_CFG['bg'], 'font_color': RIPOSO_CFG['font']})
        for t, cfg in TURNI_CONFIG.items():
            formats[t] = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': cfg['bg'], 'font_color': cfg['font']})
        
        formats_req_shift = {}
        for t, cfg in TURNI_CONFIG.items():
            formats_req_shift[t] = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': cfg['bg'], 'font_color': '#0000CC', 'bold': True})

        headers_cal = ['Ruolo', 'Dipendente']
        for d in range(NUM_GIORNI):
            curr_date = start_date + timedelta(days=d)
            day_str = GIORNI_SETT_IT[d % 7]
            date_str = curr_date.strftime("%d/%m")
            headers_cal.append(f"{day_str} {date_str}")
        
        shift_ids = sorted(TURNI_CONFIG.keys())
        headers_stats = ["Tot", "Weekend", "Mat %", "Sera %", "Bilancio"] + [TURNI_CONFIG[t]['txt'] for t in shift_ids]
        sheet.write_row(0, 0, headers_cal + headers_stats, fmt_head)

        sorted_staff = sorted(ALL_STAFF, key=lambda x: (
            0 if x in STAFF_SOLO_MATTINA else 
            1 if x == "Simone Esposito" else 
            2 if x in STAFF_NOTTE else 
            3 if x in STAFF_ODDS else 4
        ))

        r = 1
        daily_staff_count = {d: 0 for d in range(NUM_GIORNI)}
        daily_shift_breakdown = {t: {d: 0 for d in range(NUM_GIORNI)} for t in TURNI_CONFIG}

        for nome in sorted_staff:
            n = ALL_STAFF.index(nome)
            role = "Staff"
            if nome in STAFF_SOLO_MATTINA: role = "Solo Mattina"
            elif nome == "Simone Esposito": role = "Solo Pom"
            elif nome in STAFF_NOTTE: role = "Team Notte"
            elif nome in STAFF_ODDS: role = "Odds"
            elif nome in LISTA_ACCETTATORI: role = "Accettatore"

            sheet.write(r, 0, role, fmt_base)
            sheet.write(r, 1, nome, fmt_name)

            cnt_tot, cnt_m, cnt_l = 0, 0, 0
            p_counts = {t: 0 for t in TURNI_CONFIG}
            weekend_worked = 0

            for d in range(NUM_GIORNI):
                val = solver.Value(shifts[(n, d)])
                if val > 0:
                    cnt_tot += 1
                    daily_staff_count[d] += 1
                    daily_shift_breakdown[val][d] += 1
                    p_counts[val] += 1
                    if TURNI_CONFIG[val]['cat'] == 'M': cnt_m += 1
                    elif TURNI_CONFIG[val]['cat'] == 'L': cnt_l += 1
                
                is_ferie = nome in FERIE_INDICI and d in FERIE_INDICI[nome]
                is_req = nome in REQ_OFF_INDICI and d in REQ_OFF_INDICI[nome]
                is_req_turn = False
                if nome in REQ_TURNI_INDICI and d in REQ_TURNI_INDICI[nome]:
                    if REQ_TURNI_INDICI[nome][d] == val: is_req_turn = True

                if is_ferie: sheet.write(r, d+2, "FERIE", fmt_ferie)
                elif is_req: sheet.write(r, d+2, "REQ", fmt_req)
                else:
                    txt = TURNI_CONFIG[val]['txt'] if val > 0 else "-"
                    if val == 9: txt = "20-04" if (d%7) in [4, 5] else "18-02"
                    
                    used_fmt = formats_req_shift[val] if is_req_turn and val > 0 else formats[val]
                    sheet.write(r, d+2, txt, used_fmt)
            
            for w in range(NUM_SETTIMANE):
                sab_idx = 5+w*7
                dom_idx = 6+w*7
                if sab_idx < NUM_GIORNI:
                    v_s = solver.Value(shifts[(n, sab_idx)])
                    v_d = 0
                    if dom_idx < NUM_GIORNI: v_d = solver.Value(shifts[(n, dom_idx)])
                    if v_s > 0 or v_d > 0: weekend_worked += 1

            col = NUM_GIORNI + 2
            sheet.write(r, col, cnt_tot, fmt_base)
            sheet.write(r, col+1, weekend_worked, fmt_base)
            sheet.write(r, col+2, cnt_m/cnt_tot if cnt_tot else 0, fmt_pct)
            sheet.write(r, col+3, cnt_l/cnt_tot if cnt_tot else 0, fmt_pct)
            
            bal = "OK"
            if abs((cnt_m/cnt_tot if cnt_tot else 0) - (cnt_l/cnt_tot if cnt_tot else 0)) > 0.4: bal = "Sbil."
            if role in ["Solo Mattina", "Solo Pom"]: bal = "Fix"
            sheet.write(r, col+4, bal, fmt_base)
            
            off = 5
            for t_id in shift_ids:
                sheet.write(r, col+off, p_counts[t_id], fmt_base)
                off += 1
            r += 1

        r += 2
        sheet.write(r, 1, "STAFF AL LAVORO:", fmt_total)
        for d in range(NUM_GIORNI): sheet.write(r, d+2, daily_staff_count[d], fmt_total)
        r += 2
        sheet.write(r, 1, "DETTAGLIO ORARI:", fmt_head)
        r += 1
        for t_id in shift_ids:
            sheet.write(r, 1, TURNI_CONFIG[t_id]['txt'], formats[t_id])
            for d in range(NUM_GIORNI):
                sheet.write(r, d+2, daily_shift_breakdown[t_id][d], fmt_base)
            r += 1

        sheet.set_column(0, 1, 20)
        sheet.set_column(2, NUM_GIORNI+2, 5)
        writer.close()
        return output
    else:
        return None

# ==============================================================================
# UI STREAMLIT
# ==============================================================================

st.set_page_config(page_title="Gestore Turni Pro", layout="wide")
st.title("🧩 Gestore Turni Avanzato")

st.sidebar.header("Impostazioni Generali")
start_d = st.sidebar.date_input("Data Inizio Calendario", datetime(2026, 1, 5))
weeks_num = st.sidebar.slider("Numero di Settimane", min_value=1, max_value=12, value=9)

ALL_STAFF_UI = ALL_STAFF.copy()
ALL_STAFF_UI.sort()

SHIFT_OPTIONS = {
    1: "07:00", 2: "08:00", 3: "09:00", 4: "10:00", 5: "11:00",
    6: "12:00", 7: "15:00", 8: "16:00", 10: "17:00", 9: "CHIUSURA"
}

if 'list_assenze' not in st.session_state: st.session_state.list_assenze = []
if 'list_turni' not in st.session_state: st.session_state.list_turni = []

tab1, tab2 = st.tabs(["🆕 Nuova Pianificazione", "🛠️ Modalità Riparazione"])

with tab1:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. Inserisci Assenze")
        with st.form("form_add_assenza", clear_on_submit=True):
            nome = st.selectbox("Dipendente", ALL_STAFF_UI, key="nome_ass")
            tipo = st.selectbox("Tipo", ["Ferie", "Richiesta OFF"], key="tipo_ass")
            data = st.date_input("Data", key="data_ass")
            submitted = st.form_submit_button("Aggiungi")
            if submitted:
                st.session_state.list_assenze.append({"nome": nome, "tipo": tipo, "data": data})
                st.success(f"Aggiunto: {nome} - {tipo}")
        
        if st.session_state.list_assenze:
            st.write("📋 Lista Assenze:")
            st.table(st.session_state.list_assenze)
            if st.button("Cancella Assenze"):
                st.session_state.list_assenze = []
                st.experimental_rerun()

    with col2:
        st.subheader("2. Richiedi Turno Specifico")
        with st.form("form_add_turno", clear_on_submit=True):
            tnome = st.selectbox("Dipendente", ALL_STAFF_UI, key="tnome")
            tdata = st.date_input("Data", key="tdata")
            tturno = st.selectbox("Orario", [f"{k} - {v}" for k,v in SHIFT_OPTIONS.items()], key="tturno")
            submitted_t = st.form_submit_button("Aggiungi Richiesta")
            if submitted_t:
                st.session_state.list_turni.append({"nome": tnome, "data": tdata, "turno": tturno})
                st.success("Richiesta Aggiunta")

        if st.session_state.list_turni:
            st.write("📋 Lista Richieste Turni:")
            st.table(st.session_state.list_turni)
            if st.button("Cancella Turni"):
                st.session_state.list_turni = []
                st.experimental_rerun()

    st.divider()
    if st.button("🚀 GENERA TURNI", type="primary"):
        with st.spinner("Calcolo in corso..."):
            res = solve_turni(start_d, weeks_num, st.session_state.list_assenze, st.session_state.list_turni)
            if res:
                st.success("Fatto!")
                st.download_button("📥 Scarica Excel", res.getvalue(), "Turni_Generati.xlsx")
            else:
                st.error("Nessuna soluzione trovata.")

with tab2:
    st.info("Carica un Excel esistente per aggiungere assenze impreviste.")
    uploaded_file = st.file_uploader("Carica Excel", type=["xlsx"])
    
    st.write("### Nuova Assenza")
    with st.form("form_repair"):
        r_nome = st.selectbox("Dipendente", ALL_STAFF_UI)
        r_tipo = st.selectbox("Tipo", ["Ferie", "Richiesta OFF"])
        r_data = st.date_input("Data")
        sub_repair = st.form_submit_button("Ricalcola")

    if sub_repair and uploaded_file:
        num_days = 7 * weeks_num
        prev_sol, err = parse_uploaded_schedule(uploaded_file, start_d, num_days)
        if prev_sol:
            new_abs = [{"nome": r_nome, "tipo": r_tipo, "data": r_data}]
            with st.spinner("Riparazione..."):
                res = solve_turni(start_d, weeks_num, new_abs, [], prev_sol)
                if res:
                    st.success("Fatto!")
                    st.download_button("Scarica Excel Aggiornato", res.getvalue(), "Turni_Riparati.xlsx")
                else: st.error("Impossibile riparare.")
        else: st.error(f"Errore file: {err}")
